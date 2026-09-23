/* Desktop (Tauri) integration. Loaded after app.js; a no-op in a plain
   browser. Gives real file semantics: Open/Save As via native dialogs, Save
   writes back to the current path, and the open file is watched so external
   edits (another machine via OneDrive, Excel, etc.) auto-reload. */
(function () {
  'use strict';
  if (!window.__TAURI__) return;

  var dialog = window.__TAURI__.dialog;
  var fs = window.__TAURI__.fs;

  var XLSX_FILTER = [{ name: 'Excel workbook', extensions: ['xlsx'] }];
  var currentPath = null;   // absolute path of the open xlsx (null = unsaved / bundle mode)
  var unwatch = null;       // stops the active directory watcher
  var watchGen = 0;         // bumps on every rewatch so a late event from an old watcher is dropped
  var reloading = false;

  // shared-bundle (folder) document — see the "shared bundle" section below
  var bundleDir = null;      // absolute path of the open <Title>.headway folder
  var activePlanId = null;   // sub-bundle whose entity events are applied live
  var lastShardJson = {};    // rel path → canonical envelope JSON we last read or wrote (echo test)
  var bundleWarnings = [];   // {path, err} — shards skipped on read, never fatal
  var ownUserId = null;      // presence/history identity, from HeadwayDesktop.setUserId

  function app() { return window.HeadwayApp; }
  function RB() { return window.RMBundle; }
  function basename(p) { return String(p).replace(/^.*[\\/]/, ''); }
  function dirname(p) {
    var m = String(p).match(/^(.*)[\\/][^\\/]+$/);
    return m ? m[1] : p;
  }
  function samePath(a, b) {
    // watcher events may use the other separator style on Windows
    return String(a).replace(/\\/g, '/') === String(b).replace(/\\/g, '/');
  }

  function markTitle() {
    var name = bundleDir ? basename(bundleDir).replace(/\.headway$/i, '')
      : currentPath ? basename(currentPath) : null;
    document.title = name ? name + ' — Headway' : 'Headway — Roadmap Planner';
    var t = document.getElementById('docTitle');
    if (t) t.title = bundleDir || currentPath || '';
  }

  function setPath(p) {
    // opening an xlsx ends the bundle session: the watcher below prefers
    // bundleDir, so it must be gone before rewatch()
    if (p && bundleDir) leaveBundle();
    currentPath = p;
    markTitle();
    rewatch();
    if (p) app().noteRecent(p); // the start page's Recent list
  }

  // migrate the pre-start-page single last-path memory into the recents list
  (function () {
    try {
      var p = localStorage.getItem('headway-last-path');
      if (p) { app().noteRecent(p); localStorage.removeItem('headway-last-path'); }
    } catch (e) { /* storage optional */ }
  })();

  // Watch the parent directory, not the file: editors and sync clients
  // (OneDrive included) replace files by rename, which kills a file watch.
  function rewatch() {
    if (unwatch) { try { unwatch(); } catch (e) { /* already gone */ } unwatch = null; }
    var gen = ++watchGen;
    if (bundleDir) {
      // whole folder, recursively: shards live three levels down
      fs.watch(bundleDir, function (event) {
        if (gen !== watchGen) return;
        return onBundleEvent(event);
      }, { recursive: true, delayMs: 800 }).then(function (un) {
        if (gen !== watchGen) { try { un(); } catch (e) { /* stale */ } return; }
        unwatch = un;
      }).catch(function (err) {
        app().toast('Could not watch the shared folder: ' + friendlyFsError(err), 'err');
      });
      return;
    }
    if (!currentPath) return;
    fs.watch(dirname(currentPath), function (event) {
      var paths = (event && event.paths) || [];
      var hit = paths.some(function (p) { return samePath(p, currentPath); });
      // own writes are filtered by content (byte sig + embedded document
      // JSON, see below), never by timing
      // with auto-save off the document on screen is not what is on disk;
      // a reload would throw away the unsaved edits
      if (!hit || reloading || !app().autoSaveOn()) return;
      reloadFromDisk();
    }, { delayMs: 800 }).then(function (un) {
      unwatch = un;
    }).catch(function (err) {
      // plugin errors are plain strings, not Error objects
      app().toast('Could not watch for external changes: ' + (err && err.message || err), 'err');
    });
  }

  // fingerprint of the workbook bytes we last loaded or wrote — reloads are
  // applied (and announced) only when the content actually changed
  var lastSig = null;
  function sigOf(bytes) {
    var h = 2166136261;
    for (var i = 0; i < bytes.length; i++) {
      h ^= bytes[i];
      h = (h * 16777619) >>> 0;
    }
    return bytes.length + ':' + h.toString(16);
  }

  // The byte fingerprint alone cannot tell our own save from a remote edit:
  // sync clients (OneDrive/SharePoint especially) rewrite the xlsx container
  // after upload — injected sync metadata re-zips the file — so the bytes
  // change while the document does not. Track the embedded document JSON
  // (excel.js hides it in the _RoadmapTool sheet) beside the byte sig; a
  // changed file whose JSON still matches is an echo of our own write and is
  // adopted silently instead of announcing a reload.
  var lastStateJson = null;
  function noteLoadedBytes(bytes) {
    lastSig = sigOf(bytes);
    return window.RMExcel.readStateJson(bytes.buffer).then(function (json) {
      lastStateJson = json;
    }, function () {
      lastStateJson = null;
    });
  }

  // Sync clients (OneDrive especially) fire the change event before the new
  // bytes are fully on disk — an immediate read can return the old content
  // or a partial file. Retry with backoff until genuinely new bytes appear.
  var RELOAD_RETRY_MS = [1200, 3000, 8000];
  function reloadFromDisk(attempt) {
    attempt = attempt || 0;
    if (!app().autoSaveOn()) return;
    var p = currentPath;
    reloading = true;
    fs.readFile(p).then(function (bytes) {
      var validZip = bytes.length > 4 && bytes[0] === 0x50 && bytes[1] === 0x4B; // xlsx = 'PK…'
      var sig = validZip ? sigOf(bytes) : null;
      if (!validZip || sig === lastSig) {
        reloading = false;
        if (attempt < RELOAD_RETRY_MS.length) {
          setTimeout(function () {
            if (currentPath === p && !reloading) reloadFromDisk(attempt + 1);
          }, RELOAD_RETRY_MS[attempt]);
        }
        return;
      }
      // new bytes — but is it a new DOCUMENT? A sync client's container
      // rewrite of our own save carries the same embedded JSON: adopt the
      // new bytes quietly and leave the editor alone.
      return window.RMExcel.readStateJson(bytes.buffer).catch(function () {
        return null;
      }).then(function (json) {
        if (json != null && lastStateJson != null && json === lastStateJson) {
          lastSig = sig;
          reloading = false;
          return;
        }
        return app().loadBuffer(bytes.buffer, basename(p), true /* reload in place */).then(function () {
          lastSig = sig;
          lastStateJson = json;
          app().toast('Reloaded “' + basename(p) + '” — changed on disk');
        }).finally(function () {
          reloading = false;
        });
      });
    }).catch(function (err) {
      reloading = false;
      app().toast('Auto-reload failed: ' + (err && err.message || err), 'err');
    });
  }

  // ------------------------------------------------------- shared bundle
  // A shared roadmap is a folder, not a file:
  //   <Title>.headway/headway.json            plan list (merged by plan id)
  //   plans/<planId>/meta.json                meta envelope (id 'meta')
  //   plans/<planId>/{items,phases,team,costs}/<uid>.json   one envelope per entity
  //   history/<userId>.jsonl                  append-only, one writer per file
  //   presence/<userId>.json                  heartbeat
  // Envelope merge rules live in RMBundle (pure); this section is the fs side:
  // read-merge-write per shard, atomic temp+rename, echo detection by
  // canonical content per path, and OneDrive conflict-sibling absorption.
  //
  // App-side contract — methods on window.HeadwayApp. All are OPTIONAL and
  // every call is guarded with typeof, so the app runs before they exist:
  //   applyExternalEntities([{kind, id, env}])  kind ∈ RMBundle.KINDS ('items'|
  //       'phases'|'team'|'costs') or 'meta'; env = the envelope now on disk
  //       (already merged with any conflict sibling). ONE call per watch event.
  //   presenceChanged(userId, obj|null)          parsed presence/<userId>.json;
  //       null when the file was removed or is unreadable. Own id is skipped.
  //   plansChanged(headwayJson)                  parsed headway.json after an
  //       external change.
  //   noteRecent(path, kind)                     kind 'bundle'|'xlsx' (the
  //       legacy one-argument call means xlsx).
  //   beforeClose() → Promise|void               from the window close request;
  //       the window is destroyed once it settles (3 s cap).
  //   planShardChanged(planId)                   an entity shard of a plan OTHER
  //       than the active one changed (its cached Compare copy is stale).
  //   resumeBundle()                             called once desktop.js has loaded,
  //       so the app can re-link the folder its ui snapshot names.
  var ENTITY_DIRS = { items: 'item', phases: 'phase', team: 'team', costs: 'cost' };
  var RENAME_RETRY_MS = [120, 400, 1200];
  var CLOSE_HOOK_MS = 3000;
  var deniedToasted = false; // one toast per session for a read-capability gap

  function norm(p) { return String(p).replace(/\\/g, '/'); }
  function wait(ms) { return new Promise(function (res) { setTimeout(res, ms); }); }
  function errText(err) {
    return err && err.message != null ? String(err.message) : String(err == null ? 'unknown error' : err);
  }
  function isObj(v) { return v !== null && typeof v === 'object' && !Array.isArray(v); }
  function isEnvelope(v) { return isObj(v) && v.id != null; }
  function warn(path, err) { bundleWarnings.push({ path: path, err: errText(err) }); }
  // shards are written pretty (readable, diffable); equality is by canonical
  function pretty(canon) { return JSON.stringify(JSON.parse(canon), null, 2) + '\n'; }

  // path relative to the bundle root, '/'-separated, no leading slash. The
  // drive letter may differ in case between the dialog and the watcher.
  function rel(dir, p) {
    var d = norm(dir).replace(/\/+$/, ''), q = norm(p);
    if (q.toLowerCase().indexOf(d.toLowerCase() + '/') === 0) q = q.slice(d.length + 1);
    else if (q.toLowerCase() === d.toLowerCase()) q = '';
    return q.replace(/^\/+/, '');
  }

  // what a bundle-relative path is. kind: plans|meta|item|phase|team|cost|
  // history|presence|tmp|other; entityKind is the RMBundle kind (dir name) or
  // 'meta' for shards; stem is the file name without extension.
  function classify(relPath) {
    var seg = norm(relPath).split('/').filter(Boolean);
    var file = seg.length ? seg[seg.length - 1] : '';
    var out = { kind: 'other', planId: null, file: file, stem: file.replace(/\.[^.]*$/, ''), entityKind: null, userId: null };
    if (/\.tmp$/i.test(file)) { out.kind = 'tmp'; return out; }
    var json = /\.json$/i.test(file);
    if (seg.length === 1 && file === 'headway.json') { out.kind = 'plans'; return out; }
    if (seg[0] === 'plans' && seg.length >= 3) {
      out.planId = seg[1];
      if (seg.length === 3 && json) { out.kind = 'meta'; out.entityKind = 'meta'; return out; }
      if (seg.length === 4 && json && ENTITY_DIRS[seg[2]]) {
        out.kind = ENTITY_DIRS[seg[2]]; out.entityKind = seg[2]; return out;
      }
      return out;
    }
    if (seg[0] === 'history' && seg.length === 2 && /\.jsonl$/i.test(file)) { out.kind = 'history'; out.userId = out.stem; return out; }
    if (seg[0] === 'presence' && seg.length === 2 && json) { out.kind = 'presence'; out.userId = out.stem; return out; }
    return out;
  }

  // the plugin's WatchEvent.type is an object union ({create:{kind}},
  // {modify:{kind}}, {remove:{kind}}) or the strings 'any' / 'other'
  function evKind(ev) {
    var t = ev && ev.type;
    if (!t) return 'other';
    if (typeof t === 'string') return /^(create|modify|remove)$/.test(t) ? t : 'other';
    if ('create' in t) return 'create';
    if ('modify' in t) return 'modify';
    if ('remove' in t) return 'remove';
    return 'other';
  }

  // plugin permission errors are plain strings such as
  // "fs.rename not allowed. Permissions associated with this command: fs:allow-rename"
  function isDenied(err) { return /not allowed|permission/i.test(errText(err)); }
  function friendlyFsError(err) {
    var msg = errText(err);
    if (!isDenied(err)) return msg;
    var cap = msg.match(/fs:allow-[a-z-]+/);
    if (!cap) {
      var cmd = msg.match(/fs\.([a-zA-Z_]+)/);
      cap = cmd ? ['fs:allow-' + cmd[1].replace(/_/g, '-').replace(/([A-Z])/g, function (m) { return '-' + m.toLowerCase(); })] : ['fs:allow-<command>'];
    }
    return 'Headway is missing capability ' + cap[0];
  }
  function rejectFriendly(err) { throw new Error(friendlyFsError(err)); }

  function shardRelPath(planId, entityKind, id) {
    return entityKind === 'meta'
      ? 'plans/' + planId + '/meta.json'
      : 'plans/' + planId + '/' + entityKind + '/' + id + '.json';
  }
  function readTextOr(path, dflt) {
    return fs.exists(path).then(function (there) { return there ? fs.readTextFile(path) : dflt; });
  }
  // *.json names in a directory (missing directory = empty), sorted
  function listJson(dirPath) {
    return fs.exists(dirPath).then(function (there) {
      return there ? fs.readDir(dirPath) : [];
    }).then(function (entries) {
      return (entries || []).filter(function (e) {
        return e && e.isFile && /\.json$/i.test(e.name) && !/\.tmp$/i.test(e.name);
      }).map(function (e) { return e.name; }).sort();
    });
  }
  // one read, no retry, no warning — presence and sibling checks
  function readJsonOnce(path) {
    return readTextOr(path, null).then(function (text) {
      if (text == null || !String(text).trim()) return null;
      try { return JSON.parse(text); } catch (e) { return null; }
    });
  }
  // A sync client fires the event before the bytes are all there — an empty
  // or half-written shard is retried on the reload ladder; still bad after
  // the last attempt → null + warning. A missing file is null, no warning.
  function readJsonRetry(path, attempt) {
    attempt = attempt || 0;
    return fs.readTextFile(path).then(function (text) {
      if (!String(text).trim()) throw new Error('empty file');
      return JSON.parse(text);
    }).then(function (obj) { return obj; }, function (err) {
      // a capability gap is not a torn file: no ladder, no "corrupt shard" warning
      if (isDenied(err)) { var e = new Error(friendlyFsError(err)); e.denied = true; throw e; }
      return fs.exists(path).then(function (there) {
        if (!there) return null;
        if (attempt < RELOAD_RETRY_MS.length) {
          return wait(RELOAD_RETRY_MS[attempt]).then(function () { return readJsonRetry(path, attempt + 1); });
        }
        warn(path, err);
        return null;
      }, function () { warn(path, err); return null; });
    });
  }

  // write x.json.tmp then rename over x.json, so a reader never sees a torn
  // shard. A sync client may hold the target for a moment: retry the rename,
  // then write in place as a last resort. The thing holding the target is
  // often a peer's write landing, so the in-place text comes from `refresh`
  // (re-read + re-merge) when the caller provides one — never the stale
  // original. A capability gap is never masked.
  function isMissingDir(err) { return /no such file|os error 2\b|os error 3\b|ENOENT|cannot find the (path|file)/i.test(errText(err)); }
  function atomicWriteText(path, text, refresh) {
    var tmp = path + '.tmp';
    function renameAttempt(n) {
      return fs.rename(tmp, path).catch(function (err) {
        if (isDenied(err)) throw err;
        if (n < RENAME_RETRY_MS.length) {
          return wait(RENAME_RETRY_MS[n]).then(function () { return renameAttempt(n + 1); });
        }
        var again = refresh ? refresh() : Promise.resolve(text);
        return again.then(function (t) { return fs.writeTextFile(path, t); }).then(function () {
          return fs.remove(tmp).catch(function () { /* best effort */ });
        });
      });
    }
    return fs.writeTextFile(tmp, text).catch(function (err) {
      if (!isMissingDir(err)) throw err;
      return fs.mkdir(dirname(path), { recursive: true }).then(function () { return fs.writeTextFile(tmp, text); });
    }).then(function () {
      return renameAttempt(0);
    }).catch(function (err) {
      fs.remove(tmp).catch(function () { /* may not exist */ });
      rejectFriendly(err);
    });
  }
  // write a shard: remember its canonical BEFORE the rename lands so the
  // watch event it raises is recognised as our own; roll back on failure
  // Resolves {canon, env} — env is what actually landed on disk, which can be
  // a merge with a peer's shard when the rename never went through.
  function writeShard(dir, relPath, env) {
    var canon = RB().canonicalize(env);
    var prev = lastShardJson[relPath];
    var landed = env;
    lastShardJson[relPath] = canon;
    var abs = dir + '/' + relPath;
    function refresh() {
      return readJsonOnce(abs).then(function (disk) {
        if (isEnvelope(disk) && RB().canonicalize(disk) !== canon) {
          landed = RB().mergeEntity(disk, env);
          canon = RB().canonicalize(landed);
          lastShardJson[relPath] = canon;
        }
        return pretty(canon);
      });
    }
    return atomicWriteText(abs, pretty(canon), refresh).then(function () {
      return { canon: canon, env: landed };
    }, function (err) {
      if (prev === undefined) delete lastShardJson[relPath]; else lastShardJson[relPath] = prev;
      throw err;
    });
  }

  // OneDrive conflict sibling (<name>-<COMPUTERNAME>.json): fold it into the
  // canonical shard and drop it. Resolves {env, changed} — changed is false
  // when the sibling carried nothing the canonical did not already have.
  function absorbSibling(dir, planId, entityKind, sibEnv, sibRel) {
    var canonRel = shardRelPath(planId, entityKind, sibEnv.id);
    return readJsonRetry(dir + '/' + canonRel).then(function (disk) {
      var merged = isEnvelope(disk) ? RB().mergeEntity(disk, sibEnv) : sibEnv;
      var changed = RB().canonicalize(merged) !== lastShardJson[canonRel];
      return writeShard(dir, canonRel, merged).then(function (res) {
        merged = res.env;
        return removeShard(dir, sibRel).catch(function (err) { warn(sibRel, err); });
      }).then(function () { return { env: merged, changed: changed }; });
    });
  }
  // ONLY for sibling cleanup: refuses anything that is not an entity shard
  // whose envelope id differs from its file stem
  function removeShard(dir, relPath) {
    var c = classify(relPath);
    if (!c.entityKind) return Promise.reject(new Error('refusing to remove ' + relPath + ': not an entity shard'));
    var abs = dir + '/' + relPath;
    return readJsonOnce(abs).then(function (env) {
      if (!isEnvelope(env)) throw new Error('refusing to remove ' + relPath + ': unreadable envelope');
      if (String(env.id) === c.stem) throw new Error('refusing to remove ' + relPath + ': it is the canonical shard');
      return fs.remove(abs).catch(rejectFriendly);
    }).then(function () { delete lastShardJson[relPath]; });
  }

  // every shard of one plan: {meta: env|null, envs: {items:[env], …}}.
  // Never throws for one bad shard — it is skipped and warned about.
  function readPlan(dir, planId) {
    var kinds = RB().KINDS.concat('meta');
    var byId = {}, siblings = [];
    var chain = Promise.resolve();
    kinds.forEach(function (kind) {
      byId[kind] = {};
      var kdir = kind === 'meta' ? 'plans/' + planId : 'plans/' + planId + '/' + kind;
      chain = chain.then(function () { return listJson(dir + '/' + kdir); }).then(function (names) {
        return names.reduce(function (c, name) {
          return c.then(function () {
            var r = kdir + '/' + name;
            return readJsonRetry(dir + '/' + r).then(function (env) {
              if (env === null) return;
              if (!isEnvelope(env)) { warn(r, 'not an envelope'); return; }
              lastShardJson[r] = RB().canonicalize(env);
              var stem = name.replace(/\.json$/i, '');
              if (String(env.id) !== stem) { siblings.push({ kind: kind, rel: r, env: env }); return; }
              if (kind === 'meta' && stem !== 'meta') return; // a stray json at the plan root
              byId[kind][env.id] = env;
            });
          });
        }, Promise.resolve());
      });
    });
    // siblings after the walk: their canonical shard may have been read later
    chain = chain.then(function () {
      return siblings.reduce(function (c, s) {
        return c.then(function () {
          return absorbSibling(dir, planId, s.kind, s.env, s.rel).then(function (res) {
            byId[s.kind][res.env.id] = res.env;
          }, function (err) { warn(s.rel, err); });
        });
      }, Promise.resolve());
    });
    return chain.then(function () {
      var envs = {};
      RB().KINDS.forEach(function (kind) {
        envs[kind] = Object.keys(byId[kind]).sort().map(function (id) { return byId[kind][id]; });
      });
      return { meta: byId.meta.meta || null, envs: envs };
    });
  }

  function leaveBundle() {
    var dir = bundleDir, uid = ownUserId;
    watchGen++;
    if (unwatch) { try { unwatch(); } catch (e) { /* already gone */ } unwatch = null; }
    bundleDir = null; activePlanId = null; lastShardJson = {}; bundleWarnings = [];
    if (!dir || !uid) return Promise.resolve();
    return removePresence(dir, uid).catch(function () { /* best effort */ });
  }
  function removePresence(dir, userId) {
    var p = dir + '/presence/' + userId + '.json';
    return fs.remove(p).catch(function (err) {
      return fs.exists(p).then(function (there) { if (there) rejectFriendly(err); });
    });
  }

  // one watch event = a batch of paths → at most ONE applyExternalEntities
  function onBundleEvent(ev) {
    var dir = bundleDir;
    if (!dir) return Promise.resolve();
    var kind = evKind(ev), a = app(), pending = [], seen = {};
    var chain = Promise.resolve();
    ((ev && ev.paths) || []).forEach(function (p) {
      var r = rel(dir, p);
      if (seen[r]) return;
      seen[r] = true;
      var c = classify(r);
      if (c.kind === 'tmp' || c.kind === 'other' || c.kind === 'history') return;
      chain = chain.then(function () {
        if (bundleDir !== dir) return;
        if (c.kind === 'presence') {
          if (c.userId === ownUserId) return; // our own heartbeat
          if (!a || typeof a.presenceChanged !== 'function') return;
          if (kind === 'remove') { a.presenceChanged(c.userId, null); return; }
          return readJsonOnce(dir + '/' + r).then(function (obj) { a.presenceChanged(c.userId, isObj(obj) ? obj : null); });
        }
        if (c.kind === 'plans') {
          return readJsonRetry(dir + '/' + r).then(function (hw) {
            if (isObj(hw) && a && typeof a.plansChanged === 'function') a.plansChanged(hw);
          });
        }
        if (c.planId !== activePlanId) {
          // not applied live — but a Compare overlay may be showing that plan
          if (c.entityKind && a && typeof a.planShardChanged === 'function') a.planShardChanged(c.planId);
          return;
        }
        if (kind === 'remove') return; // deletes are tombstones, never file removals
        return readJsonRetry(dir + '/' + r).then(function (env) {
          if (!isEnvelope(env)) return;
          var canon = RB().canonicalize(env);
          if (canon === lastShardJson[r]) return; // echo of our own write
          lastShardJson[r] = canon;
          if (String(env.id) !== c.stem) {
            return absorbSibling(dir, c.planId, c.entityKind, env, r).then(function (res) {
              if (res.changed) pending.push({ kind: c.entityKind, id: res.env.id, env: res.env });
            });
          }
          pending.push({ kind: c.entityKind, id: env.id, env: env });
        });
      }).catch(function (err) {
        warn(r, err);
        // a denied read would otherwise be an invisible warning per shard forever
        if (err && err.denied && !deniedToasted) {
          deniedToasted = true;
          if (a && typeof a.toast === 'function') a.toast(errText(err), 'err');
        }
      });
    });
    return chain.then(function () {
      if (bundleDir !== dir || !pending.length) return;
      if (a && typeof a.applyExternalEntities === 'function') a.applyExternalEntities(pending);
    });
  }

  window.HeadwayDesktop = {
    // ---- shared bundle (folder) backend ----
    // open a <Title>.headway folder; planId defaults to the first live plan.
    // Resolves {doc, planId, plans, headway, envs, metaEnv, warnings}; the
    // only hard failure is a missing/invalid headway.json.
    openBundle: function (dir, planId) {
      dir = norm(dir).replace(/\/+$/, '');
      return fs.readTextFile(dir + '/headway.json').catch(rejectFriendly).then(function (text) {
        var hw;
        try { hw = JSON.parse(text); } catch (e) { throw new Error('Not a Headway shared folder: headway.json is not valid JSON'); }
        if (!isObj(hw) || !Array.isArray(hw.plans)) throw new Error('Not a Headway shared folder: headway.json has no plan list');
        var live = hw.plans.filter(function (p) { return p && p.id && !p.deleted; });
        var pid = planId || (live[0] && live[0].id);
        if (!pid) throw new Error('This shared folder has no plans');
        if (bundleDir && bundleDir !== dir) leaveBundle();
        bundleWarnings = [];
        lastShardJson = {};
        return readPlan(dir, pid).then(function (plan) {
          bundleDir = dir;
          activePlanId = pid;
          // keeps every xlsx path inert: autosave, renameTo, reload all gate on currentPath
          currentPath = null; lastSig = null; lastStateJson = null;
          markTitle();
          rewatch();
          var a = app();
          if (a && typeof a.noteRecent === 'function') a.noteRecent(dir, 'bundle');
          return {
            doc: RB().assembleState(plan.meta, plan.envs),
            planId: pid, plans: hw.plans, headway: hw,
            envs: plan.envs, metaEnv: plan.meta,
            warnings: bundleWarnings.slice()
          };
        });
      });
    },
    readPlan: readPlan,

    // changes: [{kind, id, env, baseRev}], kind ∈ RMBundle.KINDS or 'meta'.
    // Read-merge-write per shard: a shard whose rev moved past baseRev was
    // written by a peer since our last read, so merge before overwriting.
    // Resolves {written:[{kind,id,canon,env}], merged:[id]} — canon is the
    // envelope canonical (the echo key), env what is now on disk.
    flushShards: function (dir, planId, changes) {
      var written = [], merged = [];
      return (changes || []).reduce(function (chain, ch) {
        return chain.then(function () {
          var r = shardRelPath(planId, ch.kind, ch.id);
          return readJsonRetry(dir + '/' + r).then(function (disk) {
            var env = ch.env;
            var base = isFinite(+ch.baseRev) ? +ch.baseRev : 0;
            if (isEnvelope(disk) && (+disk.rev || 0) > base) {
              env = RB().mergeEntity(disk, env);
              merged.push(ch.id);
            }
            return writeShard(dir, r, env).then(function (res) {
              written.push({ kind: ch.kind, id: ch.id, canon: res.canon, env: res.env });
            });
          });
        });
      }, Promise.resolve()).then(function () { return { written: written, merged: merged }; });
    },

    // history/<userId>.jsonl — we are its only writer, so no temp file
    appendHistory: function (dir, userId, line) {
      var p = dir + '/history/' + userId + '.jsonl';
      return fs.mkdir(dir + '/history', { recursive: true }).then(function () {
        return readTextOr(p, '');
      }).then(function (text) {
        text = String(text || '');
        if (text && !/\n$/.test(text)) text += '\n';
        return fs.writeTextFile(p, text + window.RM.asciiJson(line) + '\n');
      }).catch(rejectFriendly);
    },
    rewriteHistory: function (dir, userId, lines) {
      var p = dir + '/history/' + userId + '.jsonl';
      return fs.mkdir(dir + '/history', { recursive: true }).then(function () {
        return fs.writeTextFile(p, RB().encodeHistory(lines));
      }).catch(rejectFriendly);
    },
    readHistory: function (dir) {
      var out = {};
      return fs.exists(dir + '/history').then(function (there) {
        return there ? fs.readDir(dir + '/history') : [];
      }).then(function (entries) {
        return (entries || []).filter(function (e) { return e && e.isFile && /\.jsonl$/i.test(e.name); })
          .reduce(function (chain, e) {
            return chain.then(function () {
              return fs.readTextFile(dir + '/history/' + e.name).then(function (text) {
                out[e.name.replace(/\.jsonl$/i, '')] = RB().parseHistory(text);
              }, function (err) { warn('history/' + e.name, err); });
            });
          }, Promise.resolve());
      }).then(function () { return out; });
    },

    // presence is advisory: plain writes, unreadable files are skipped
    writePresence: function (dir, userId, obj) {
      return fs.mkdir(dir + '/presence', { recursive: true }).then(function () {
        return fs.writeTextFile(dir + '/presence/' + userId + '.json', JSON.stringify(obj));
      }).catch(rejectFriendly);
    },
    readPresence: function (dir) {
      var out = {};
      return fs.exists(dir + '/presence').then(function (there) {
        return there ? fs.readDir(dir + '/presence') : [];
      }).then(function (entries) {
        return (entries || []).filter(function (e) { return e && e.isFile && /\.json$/i.test(e.name); })
          .reduce(function (chain, e) {
            return chain.then(function () {
              return readJsonOnce(dir + '/presence/' + e.name).then(function (obj) {
                if (isObj(obj)) out[e.name.replace(/\.json$/i, '')] = obj;
              }, function () { /* advisory: a file mid-removal or unreadable is skipped */ });
            });
          }, Promise.resolve());
      }).then(function () { return out; });
    },
    removePresence: removePresence,
    removeShard: removeShard,

    // contents = RMBundle.migrateFromState(...): {headway, plans:{<pid>:{meta,
    // items:[env], …}}, history:{<uid>:[lines]}}. Resolves the folder path.
    createBundle: function (dir, contents) {
      dir = norm(dir).replace(/\/+$/, '');
      var files = []; // [rel, text]
      var shards = []; // [rel, env]
      files.push(['headway.json', JSON.stringify(contents.headway, null, 2) + '\n']);
      Object.keys(contents.plans || {}).forEach(function (pid) {
        var plan = contents.plans[pid];
        if (plan.meta) shards.push([shardRelPath(pid, 'meta', 'meta'), plan.meta]);
        RB().KINDS.forEach(function (kind) {
          (plan[kind] || []).forEach(function (env) { shards.push([shardRelPath(pid, kind, env.id), env]); });
        });
      });
      Object.keys(contents.history || {}).forEach(function (uid) {
        files.push(['history/' + uid + '.jsonl', RB().encodeHistory(contents.history[uid])]);
      });
      var dirs = ['history', 'presence'];
      Object.keys(contents.plans || {}).forEach(function (pid) {
        RB().KINDS.forEach(function (kind) { dirs.push('plans/' + pid + '/' + kind); });
      });
      return fs.mkdir(dir, { recursive: true }).then(function () {
        return dirs.reduce(function (chain, d) {
          return chain.then(function () { return fs.mkdir(dir + '/' + d, { recursive: true }); });
        }, Promise.resolve());
      }).then(function () {
        return files.reduce(function (chain, f) {
          return chain.then(function () { return atomicWriteText(dir + '/' + f[0], f[1]); });
        }, Promise.resolve());
      }).then(function () {
        return shards.reduce(function (chain, s) {
          return chain.then(function () { return writeShard(dir, s[0], s[1]); });
        }, Promise.resolve());
      }).catch(rejectFriendly).then(function () { return dir; });
    },
    openBundleDialog: function () {
      return dialog.open({ directory: true, multiple: false }).then(function (dir) {
        return dir ? window.HeadwayDesktop.openBundle(dir) : null;
      });
    },
    // a parent folder for a new / converted bundle; null on cancel
    pickFolder: function () {
      return dialog.open({ directory: true, multiple: false }).then(function (d) { return d || null; });
    },
    // so Create/Convert can refuse a <parent>/<title>.headway that already exists
    // (createBundle would otherwise write into — and merge with — someone's bundle)
    pathExists: function (p) {
      return fs.exists(norm(p).replace(/\/+$/, '')).catch(function () { return false; });
    },
    readHeadway: function (dir) {
      return readJsonRetry(norm(dir).replace(/\/+$/, '') + '/headway.json');
    },
    // headway.json read-merge-write: hw.plans merge by id into what is on
    // disk (a peer's concurrent create / rename / tombstone survives), other
    // keys overwrite. Resolves the document that landed.
    writeHeadway: function (dir, hw) {
      dir = norm(dir).replace(/\/+$/, '');
      var p = dir + '/headway.json';
      return readJsonRetry(p).then(function (disk) {
        var base = isObj(disk) ? disk : {};
        var out = {};
        Object.keys(base).forEach(function (k) { out[k] = base[k]; });
        Object.keys(hw || {}).forEach(function (k) { if (k !== 'plans') out[k] = hw[k]; });
        out.format = out.format || RB().FORMAT;
        out.plans = RB().mergePlanList(base.plans || [], (hw && hw.plans) || []);
        return atomicWriteText(p, JSON.stringify(out, null, 2) + '\n').then(function () { return out; });
      });
    },
    // write a blob where the user says and adopt NOTHING: currentPath, the
    // watcher and the bundle session are untouched (a bundle's .xlsx export).
    // Resolves the path, or null on cancel.
    exportBlob: function (blob, suggestedName, ext, filterName) {
      ext = ext || 'xlsx';
      return dialog.save({
        defaultPath: suggestedName,
        filters: [{ name: filterName || 'File', extensions: [ext] }]
      }).then(function (p) {
        if (!p) return null;
        if (!new RegExp('\\.' + ext + '$', 'i').test(p)) p += '.' + ext;
        return blob.arrayBuffer().then(function (buf) {
          return fs.writeFile(p, new Uint8Array(buf));
        }).then(function () { return p; });
      });
    },
    // unwatch, drop our presence file, forget the folder
    closeBundle: function () {
      return leaveBundle().then(function () { markTitle(); });
    },
    setUserId: function (id) { ownUserId = id ? String(id) : null; },
    userId: function () { return ownUserId; },
    bundleDir: function () { return bundleDir; },
    activePlanId: function () { return activePlanId; },
    bundleWarnings: function () { return bundleWarnings.slice(); },
    classify: classify,
    friendlyFsError: friendlyFsError,

    // ---- xlsx (single-file) backend ----
    // pick an .xlsx and hand back its bytes WITHOUT adopting it — no
    // currentPath, no watcher, the bundle stays the document (Import from
    // Excel…). Resolves {path, name, buffer}, or null on cancel.
    pickWorkbook: function () {
      return dialog.open({ multiple: false, filters: XLSX_FILTER }).then(function (p) {
        if (!p) return null;
        return fs.readFile(p).then(function (bytes) {
          // an exact ArrayBuffer: a view may sit inside a larger pool
          var buf = bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength);
          return { path: p, name: basename(p), buffer: buf };
        });
      });
    },
    // native open dialog → load → remember + watch the path
    openDialog: function () {
      dialog.open({ multiple: false, filters: XLSX_FILTER }).then(function (p) {
        if (!p) return;
        fs.readFile(p).then(function (bytes) {
          return app().loadBuffer(bytes.buffer, basename(p)).then(function () {
            return noteLoadedBytes(bytes);
          });
        }).then(function () {
          setPath(p);
        }).catch(function (err) {
          app().toast('Could not open: ' + err.message, 'err');
        });
      });
    },

    // write the workbook; dialog only when there's no path yet (or Save As).
    // stateJson: the document JSON embedded in this blob (RMExcel.stateJsonOf)
    // — remembered so a sync client's rewrite of this save is not mistaken
    // for a remote change. Resolves to the saved path, or null on cancel.
    saveBlob: function (blob, suggestedName, forceDialog, stateJson) {
      var target = (currentPath && !forceDialog)
        ? Promise.resolve(currentPath)
        : dialog.save({ defaultPath: suggestedName, filters: XLSX_FILTER });
      return target.then(function (p) {
        if (!p) return null;
        if (!/\.xlsx$/i.test(p)) p += '.xlsx';
        return blob.arrayBuffer().then(function (buf) {
          var u8 = new Uint8Array(buf);
          lastSig = sigOf(u8);
          lastStateJson = stateJson != null ? stateJson : null;
          return fs.writeFile(p, u8);
        }).then(function () {
          if (!samePath(p, currentPath || '')) setPath(p);
          return p;
        });
      });
    },

    // file export (PNG, PPTX, …): ask where via the native dialog, write,
    // then OPEN the file
    saveFileAndOpen: function (blob, suggestedName, filterName, ext) {
      return dialog.save({
        defaultPath: suggestedName,
        filters: [{ name: filterName || 'File', extensions: [ext || 'png'] }]
      }).then(function (p) {
        if (!p) return null;
        if (!new RegExp('\\.' + (ext || 'png') + '$', 'i').test(p)) p += '.' + (ext || 'png');
        return blob.arrayBuffer().then(function (buf) {
          return fs.writeFile(p, new Uint8Array(buf));
        }).then(function () {
          var op = window.__TAURI__.opener;
          if (op && op.openPath) {
            // best-effort: the export succeeded even if opening doesn't
            return op.openPath(p).then(function () { return p; }, function () { return p; });
          }
          return p;
        });
      });
    },

    // a split export writes several files at once: pick a folder, write all
    saveManyToFolder: function (files) {
      return dialog.open({ directory: true, multiple: false }).then(function (dir) {
        if (!dir) return null;
        var chain = Promise.resolve();
        files.forEach(function (f) {
          chain = chain.then(function () {
            return f.blob.arrayBuffer().then(function (buf) {
              return fs.writeFile(dir + '/' + f.name, new Uint8Array(buf));
            });
          });
        });
        return chain.then(function () { return dir; });
      });
    },

    // open an external http(s) link in the OS browser
    openUrl: function (url) {
      var op = window.__TAURI__ && window.__TAURI__.opener;
      if (!op || !op.openUrl) { window.open(url, '_blank', 'noopener'); return Promise.resolve(); }
      return op.openUrl(url).catch(function (err) { app().toast('Could not open link: ' + (err && err.message || err), 'err'); });
    },

    // open a known path (start page recents) — rejects if unreadable
    openPath: function (p) {
      return fs.readFile(p).then(function (bytes) {
        return app().loadBuffer(bytes.buffer, basename(p)).then(function () {
          return noteLoadedBytes(bytes);
        });
      }).then(function () {
        setPath(p);
        return p;
      });
    },

    // rename the open file in place (the title IS the filename). Resolves to
    // the new path; null with no open file; rejects if the target exists or
    // the filesystem refuses.
    renameTo: function (newBase) {
      if (!currentPath) return Promise.resolve(null);
      if (!/\.xlsx$/i.test(newBase)) newBase += '.xlsx';
      var np = dirname(currentPath) + '/' + newBase;
      if (samePath(np, currentPath)) return Promise.resolve(currentPath);
      return fs.exists(np).then(function (there) {
        if (there) throw new Error('“' + newBase + '” already exists in this folder');
        return fs.rename(currentPath, np);
      }).then(function () {
        setPath(np); // re-watches the directory and refreshes the recents list
        return np;
      });
    },

    currentPath: function () { return currentPath; },
    basename: basename,
    // Claude Code process bridge for the AI assistant's "Claude subscription"
    // provider (js/ai.js owns the stream-json protocol). One global event
    // listener fans stdout/stderr/exit lines out to the live handles by pid.
    claude: (function () {
      var invoke = window.__TAURI__.core.invoke;
      var handles = {};
      var listening = false;
      function ensureListener() {
        if (listening) return;
        listening = true;
        window.__TAURI__.event.listen('ai-proc', function (ev) {
          var p = ev.payload || {};
          var h = handles[p.id];
          if (!h) return;
          if (p.kind === 'exit') { delete handles[p.id]; h.alive = false; }
          try { h.onLine(p.kind, p.line); } catch (e) { /* handler error must not kill the pipe */ }
        });
      }
      return {
        // resolve the CLI binary (custom path wins); null when not installed
        path: function (custom) { return invoke('ai_claude_path', { custom: custom || '' }); },
        // spawn → handle { write(line), kill(), alive }
        spawn: function (bin, args, onLine) {
          ensureListener();
          return invoke('ai_spawn', { bin: bin, args: args }).then(function (id) {
            var h = { id: id, alive: true, onLine: onLine,
              write: function (line) { return invoke('ai_write', { id: id, line: line }); },
              kill: function () { h.alive = false; delete handles[id]; return invoke('ai_kill', { id: id }); } };
            handles[id] = h;
            return h;
          });
        }
      };
    })(),
    appVersion: '' // filled asynchronously below
  };

  // the app booted before this file loaded: let it re-link the shared folder
  // its ui snapshot names (a reload mid-session). Optional, like the rest of
  // the app-side contract.
  (function () {
    var a = app();
    if (!a || typeof a.resumeBundle !== 'function') return;
    try { a.resumeBundle(); } catch (e) { /* the app toasts its own failures */ }
  })();

  // app version (start page footer)
  if (window.__TAURI__.app && window.__TAURI__.app.getVersion) {
    window.__TAURI__.app.getVersion().then(function (v) {
      window.HeadwayDesktop.appVersion = v || '';
      // refresh the footer if the start page is already showing
      if (document.body.classList.contains('start') && app() && app().renderStartPage) {
        app().renderStartPage();
      }
      // first launch on a new version → "What's new" (once per version)
      if (app() && app().maybeShowReleaseNotes) app().maybeShowReleaseNotes();
    }).catch(function () { /* fine without it */ });
  }

  // ------------------------------------------------------- window chrome
  // The header doubles as the titlebar. macOS: native titlebar hidden with
  // overlay traffic lights (see tauri.macos.conf.json) — the header gets
  // left padding to clear them. Windows: frameless window with Windows 11
  // style caption buttons on the right. Empty header space drags the window
  // (and double-click zooms/maximizes, per each platform's convention).
  (function chrome() {
    var isMac = navigator.platform.indexOf('Mac') === 0;
    var topbar = document.getElementById('topbar');
    if (!topbar) return;

    // dragging works from header background and passive elements (brand
    // mark, gaps around the view tabs) — never from actual controls
    Array.prototype.forEach.call(
      document.querySelectorAll('#topbar, .tb-brand, .tb-mark, .tb-mark span'),
      function (el) { el.setAttribute('data-tauri-drag-region', ''); }
    );
    document.body.classList.add(isMac ? 'chrome-mac' : 'chrome-win');

    // explicit drag handler — the built-in data-tauri-drag-region listener
    // has proven unreliable here, so start the drag ourselves. Document-level
    // so the start page's bar drags too.
    var win = window.__TAURI__.window.getCurrentWindow();
    document.addEventListener('mousedown', function (e) {
      if (e.button !== 0) return;
      var t = e.target;
      if (!t.hasAttribute || !t.hasAttribute('data-tauri-drag-region')) return;
      e.preventDefault();
      if (e.detail >= 2) win.toggleMaximize();
      else win.startDragging();
    });

    // macOS native fullscreen hides the traffic lights — body.fullscreen
    // lets the CSS reclaim their header inset
    function syncFullscreen() {
      win.isFullscreen().then(function (fs) {
        document.body.classList.toggle('fullscreen', fs);
      }).catch(function () { /* window gone */ });
    }
    win.onResized(syncFullscreen);
    syncFullscreen();

    // closing the window (caption ✕, Alt+F4, the red traffic light) — ONE
    // handler, two duties in order:
    //   1. unsaved .xlsx work asks first — app.js owns the Save / Don't save /
    //      Cancel dialog (guardUnsaved) and flushes silently when autosave
    //      already owns the file. A shared bundle is never "unsaved" that
    //      way (unsavedNow is false for it), so it skips straight to 2.
    //   2. beforeClose(): the app lands its pending bundle flush and drops
    //      its presence file; the window is destroyed once that settles
    //      (CLOSE_HOOK_MS cap). destroy() raises no second close-requested
    //      event, so an approved close cannot re-prompt (needs
    //      core:window:allow-destroy). Absent in jsdom/tests.
    if (typeof win.onCloseRequested === 'function') {
      win.onCloseRequested(function (ev) {
        ev.preventDefault();
        var a = app();
        function finish() {
          var p = null;
          try { p = (a && typeof a.beforeClose === 'function') ? a.beforeClose() : null; } catch (e) { p = null; }
          Promise.race([
            Promise.resolve(p).catch(function () { /* still close */ }),
            new Promise(function (res) { setTimeout(res, CLOSE_HOOK_MS); })
          ]).then(function () { return win.destroy(); }).catch(function () { /* window gone */ });
        }
        if (a && typeof a.guardUnsaved === 'function' && typeof a.unsavedNow === 'function' && a.unsavedNow()) {
          a.guardUnsaved(finish); // Cancel keeps the window: finish never runs
        } else finish();
      });
    }
    if (isMac) return;

    // Windows caption buttons — one set in the editor header, one on the
    // start page bar (whichever surface is visible carries them)
    var GLYPH = {
      min: '<svg viewBox="0 0 10 10" width="10" height="10"><path d="M0 5h10" stroke="currentColor" fill="none"/></svg>',
      max: '<svg viewBox="0 0 10 10" width="10" height="10"><rect x=".5" y=".5" width="9" height="9" stroke="currentColor" fill="none"/></svg>',
      restore: '<svg viewBox="0 0 10 10" width="10" height="10"><rect x=".5" y="2.5" width="7" height="7" stroke="currentColor" fill="none"/><path d="M2.5 2.5v-2h7v7h-2" stroke="currentColor" fill="none"/></svg>',
      close: '<svg viewBox="0 0 10 10" width="10" height="10"><path d="M0 0l10 10M10 0L0 10" stroke="currentColor" fill="none"/></svg>'
    };
    var bars = [];
    function makeCaptionBar(host) {
      if (!host) return;
      var bar = document.createElement('div');
      bar.className = 'win-caption';
      bar.innerHTML =
        '<button class="cap-min" title="Minimize" aria-label="Minimize">' + GLYPH.min + '</button>' +
        '<button class="cap-max" title="Maximize" aria-label="Maximize">' + GLYPH.max + '</button>' +
        '<button class="cap-close" title="Close" aria-label="Close">' + GLYPH.close + '</button>';
      bar.querySelector('.cap-min').addEventListener('click', function () { win.minimize(); });
      bar.querySelector('.cap-max').addEventListener('click', function () {
        win.toggleMaximize().then(syncMaxGlyph);
      });
      bar.querySelector('.cap-close').addEventListener('click', function () { win.close(); });
      host.appendChild(bar);
      bars.push(bar);
    }
    makeCaptionBar(topbar);
    makeCaptionBar(document.getElementById('startBar'));

    function syncMaxGlyph() {
      win.isMaximized().then(function (max) {
        bars.forEach(function (bar) {
          var b = bar.querySelector('.cap-max');
          b.innerHTML = max ? GLYPH.restore : GLYPH.max;
          b.title = max ? 'Restore' : 'Maximize';
          b.setAttribute('aria-label', b.title);
        });
      });
    }
    win.onResized(syncMaxGlyph);
    syncMaxGlyph();
  })();

  // ------------------------------------------------------- macOS menu bar
  // Mirror the in-app File / Edit / View menus into the system menu bar and
  // hide the in-window menu buttons. Rebuilt (debounced, only when the spec
  // actually changes) so checkmarks and enabled states follow the app state.
  (function nativeMenu() {
    var menu = window.__TAURI__.menu;
    var isMac = navigator.platform.indexOf('Mac') === 0;
    if (!isMac || !menu) return;

    var menusNav = document.getElementById('menus');
    if (menusNav) menusNav.style.display = 'none';

    function accel(kbd) {
      if (!kbd || /scroll/.test(kbd)) return undefined;
      var mods = [];
      if (kbd.indexOf('⇧') >= 0) mods.push('Shift');
      if (kbd.indexOf('⌥') >= 0) mods.push('Alt');
      if (kbd.indexOf('⌘') >= 0) mods.push('CmdOrCtrl');
      var key = kbd.replace(/[⇧⌥⌘]/g, '');
      return key ? mods.concat(key.toUpperCase()).join('+') : undefined;
    }

    function fireItem(m) {
      // ⌘Z in a text field should stay the field's own undo, not the app's
      if (m.label === 'Undo' || m.label === 'Redo') {
        var ae = document.activeElement;
        if (ae && (/INPUT|TEXTAREA|SELECT/.test(ae.tagName) || ae.isContentEditable)) {
          document.execCommand(m.label.toLowerCase());
          return;
        }
      }
      m.fn();
    }

    // NativeIcon names that AppKit exposes as *Template images (see muda's
    // macOS NativeIcon → NSImageName table); everything else is a coloured
    // pictogram (Folder, Info, MultipleDocuments, User, Trash…).
    var TEMPLATE_ICONS = {};
    ['Add', 'Remove', 'Share', 'Refresh', 'RefreshFreestanding', 'StopProgress',
      'StopProgressFreestanding', 'FollowLinkFreestanding', 'RevealFreestanding',
      'InvalidDataFreestanding', 'GoLeft', 'GoRight', 'LeftFacingTriangle',
      'RightFacingTriangle', 'Home', 'Bookmarks', 'Bluetooth', 'ColumnView',
      'FlowView', 'IconView', 'ListView', 'EnterFullScreen', 'ExitFullScreen',
      'IChatTheater', 'LockLocked', 'LockUnlocked', 'MenuMixedState',
      'MenuOnState', 'Path', 'QuickLook', 'Slideshow', 'SmartBadge'
    ].forEach(function (n) { TEMPLATE_ICONS[n] = true; });

    function toNative(items) {
      return Promise.all(items.map(function (m) {
        if (m.sep) return menu.PredefinedMenuItem.new({ item: 'Separator' });
        var opts = {
          text: m.label,
          enabled: !m.disabled,
          accelerator: accel(m.kbd),
          action: function () { fireItem(m); }
        };
        if ('checked' in m) {
          return menu.CheckMenuItem.new(Object.assign({ checked: !!m.checked }, opts));
        }
        // macOS template icons where the app names one; plain item otherwise.
        // Only template images (monochrome, tinted by the menu) are allowed —
        // the other NativeIcon names are full-colour pictograms that clash.
        if (m.nativeIcon && TEMPLATE_ICONS[m.nativeIcon] && menu.IconMenuItem && menu.NativeIcon && menu.NativeIcon[m.nativeIcon]) {
          return menu.IconMenuItem.new(Object.assign({ icon: menu.NativeIcon[m.nativeIcon] }, opts))
            .catch(function () { return menu.MenuItem.new(opts); });
        }
        return menu.MenuItem.new(opts);
      }));
    }

    function predefined(names) {
      return Promise.all(names.map(function (n) {
        return menu.PredefinedMenuItem.new({ item: n });
      }));
    }

    // app Edit menu + the system clipboard block (without it, ⌘C/⌘V would
    // stop working in the webview once the default menu is replaced)
    function buildEdit(items) {
      var cut = items.findIndex(function (m) { return m.sep; });
      return Promise.all([
        toNative(items.slice(0, cut)),
        predefined(['Separator', 'Cut', 'Copy', 'Paste', 'SelectAll']),
        toNative(items.slice(cut))
      ]).then(function (parts) {
        return menu.Submenu.new({ text: 'Edit', items: parts[0].concat(parts[1], parts[2]) });
      });
    }

    function buildMenu() {
      var M = window.HeadwayApp.menuItems;
      return Promise.all([
        // the app menu carries the file actions (New/Open/Save/Export/Help),
        // then the standard Hide/Quit block — there is no File submenu on mac
        Promise.all([
          toNative(M('macApp')),
          predefined(['Separator', 'Hide', 'HideOthers', 'ShowAll', 'Separator', 'Quit'])
        ]).then(function (parts) {
          return menu.Submenu.new({ text: 'Headway', items: parts[0].concat(parts[1]) });
        }),
        buildEdit(M('edit')),
        toNative(M('view')).then(function (items) {
          return menu.Submenu.new({ text: 'View', items: items });
        }),
        predefined(['Minimize', 'Maximize', 'Fullscreen', 'Separator', 'CloseWindow']).then(function (items) {
          return menu.Submenu.new({ text: 'Window', items: items });
        })
      ]).then(function (subs) {
        return menu.Menu.new({ items: subs });
      }).then(function (m) {
        return m.setAsAppMenu();
      });
    }

    var sig = '', building = false, dirty = false, timer = null;
    function specSig() {
      var M = window.HeadwayApp.menuItems;
      return JSON.stringify(['macApp', 'edit', 'view'].map(function (n) {
        return M(n).map(function (m) {
          return m.sep ? '-' : [m.label, !!m.disabled, !!m.checked, m.kbd || ''];
        });
      }));
    }
    function syncNow() {
      var s;
      try { s = specSig(); } catch (e) { return; }
      if (s === sig) return;
      if (building) { dirty = true; return; }
      sig = s; building = true;
      buildMenu().catch(function (err) {
        console.warn('menu build failed', err);
        sig = ''; // retry on the next sync
      }).finally(function () {
        building = false;
        if (dirty) { dirty = false; syncNow(); }
      });
    }

    window.HeadwayDesktop.syncMenu = function () {
      clearTimeout(timer);
      timer = setTimeout(syncNow, 200);
    };
    window.HeadwayDesktop.syncMenu();
  })();

  // ------------------------------------------------------- auto-update
  // One updater shared by the header's flashing Update button and the start
  // page's Check-for-updates button. Checks 4s after launch and then every
  // hour; a found update downloads in the background and the buttons flip to
  // "Update to x.y.z". Clicking installs and relaunches.
  (function autoUpdate() {
    var up = window.__TAURI__.updater;
    var proc = window.__TAURI__.process;
    if (!up || !proc) return;

    var CHECK_EVERY = 60 * 60 * 1000;
    var U = {
      state: 'idle',   // idle | checking | downloading | ready | installing
      version: '',     // the downloaded update's version (state ready/installing)
      error: '',       // last manual-check failure, for the start page
      checkedAt: 0,
      update: null,
      listeners: []
    };
    function set(patch) {
      Object.keys(patch).forEach(function (k) { U[k] = patch[k]; });
      U.listeners.forEach(function (fn) { try { fn(U); } catch (e) { /* listener */ } });
      // the start page renders the updater's state into its button
      if (document.body.classList.contains('start') && app() && app().renderStartPage) app().renderStartPage();
    }
    U.onChange = function (fn) { U.listeners.push(fn); };

    // manual: surface "up to date" / failures; automatic: silent
    U.check = function (manual) {
      if (U.state === 'ready' || U.state === 'installing') return Promise.resolve(U.update);
      if (U.state === 'checking' || U.state === 'downloading') return Promise.resolve(null);
      set({ state: 'checking', error: '' });
      return up.check().then(function (update) {
        if (!update) {
          set({ state: 'idle', checkedAt: Date.now() });
          if (manual && app()) app().toast('Headway ' + (window.HeadwayDesktop.appVersion || '') + ' is up to date');
          return null;
        }
        set({ state: 'downloading', version: update.version, checkedAt: Date.now() });
        return update.download().then(function () {
          set({ state: 'ready', update: update });
          showHeaderButton();
          return update;
        });
      }).catch(function (err) {
        // offline, dev build, or no release yet — silent unless asked for
        var msg = (err && err.message) || String(err || 'unknown error');
        set({ state: 'idle', error: manual ? msg : '' });
        if (manual && app()) app().toast('Could not check for updates — ' + msg, 'err');
        return null;
      });
    };

    U.install = function () {
      if (U.state !== 'ready' || !U.update) return Promise.resolve();
      set({ state: 'installing' });
      return U.update.install().then(function () {
        return proc.relaunch(); // NSIS on Windows exits/relaunches itself
      }).catch(function (err) {
        set({ state: 'ready' });
        if (app()) app().toast('Update failed: ' + (err && err.message || err), 'err');
      });
    };

    function showHeaderButton() {
      if (document.getElementById('btnUpdate')) return;
      var right = document.querySelector('.tb-right');
      if (!right) return;
      var b = document.createElement('button');
      b.id = 'btnUpdate';
      b.title = 'Version ' + U.version + ' downloaded';
      b.innerHTML = '<i data-lucide="refresh-cw"></i>Update';
      b.addEventListener('click', function () { U.install(); });
      U.onChange(function (u) {
        b.disabled = u.state === 'installing';
        b.innerHTML = u.state === 'installing' ? 'Updating…' : '<i data-lucide="refresh-cw"></i>Update';
        if (window.lucide) lucide.createIcons();
      });
      right.insertBefore(b, right.firstChild);
      if (window.lucide) lucide.createIcons();
    }

    window.HeadwayDesktop.updater = U;
    setTimeout(function () { U.check(false); }, 4000);
    setInterval(function () { U.check(false); }, CHECK_EVERY);
  })();
})();
