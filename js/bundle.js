/*
 * Headway shared-bundle format — pure data, no filesystem. A bundle is a
 * folder of per-entity JSON "shards"; each shard is an envelope carrying the
 * entity's fields plus a per-field change stamp, so two machines that edited
 * the same item merge field by field (newest stamp wins, ties broken
 * deterministically). Loaded in the browser as window.RMBundle and in node
 * (tests) via require.
 */
(function (root) {
  'use strict';

  var RM = root.RM || (typeof require === 'function' && require('./core.js'));
  var RMBundle = {};

  RMBundle.FORMAT = 'headway-bundle-v1';
  // entity kinds that get one shard per row, in plans/<planId>/<kind>/
  RMBundle.KINDS = ['items', 'phases', 'team', 'costs'];
  // Top-level state keys that travel in the plan's meta.json: everything
  // normalizeState owns at the top level that is neither an entity list nor
  // left outside the plan (history has its own files; optId/optName/options
  // become plans).
  RMBundle.META_KEYS = ['meta', 'wsOrder', 'wsColors', 'epicIcons', 'epicColors', 'epicJira', 'teamTypes'];
  // in-memory-only keys — must not count as a change
  RMBundle.VOLATILE_KEYS = ['holdPos', '_idx', 'leadDays'];
  // keys a bundle-backed document carries that a standalone xlsx must not
  RMBundle.BUNDLE_MARKERS = ['docId', 'bundle', 'planId'];

  function isObj(v) { return v !== null && typeof v === 'object' && !Array.isArray(v); }
  function later(a, b) { return a > b ? a : b; }
  function str(v) { return v == null ? '' : String(v); }
  function has(list, v) { return Array.isArray(list) && list.indexOf(v) !== -1; }
  function sortedCopy(obj) {
    var out = {};
    Object.keys(obj).sort().forEach(function (k) { out[k] = obj[k]; });
    return out;
  }

  // ------------------------------------------------------------ canonical form
  function clean(v) {
    if (Array.isArray(v)) return v.map(clean);
    if (!isObj(v)) return v;
    var out = {};
    Object.keys(v).sort().forEach(function (k) {
      if (RMBundle.VOLATILE_KEYS.indexOf(k) !== -1 || v[k] === undefined) return;
      out[k] = clean(v[k]);
    });
    return out;
  }
  // sorted-key ASCII JSON: equal strings = equal documents (the echo test)
  RMBundle.canonicalize = function (obj) {
    var c = clean(obj);
    return c === undefined ? '' : RM.asciiJson(c);
  };

  // ------------------------------------------------------------ envelopes
  // { id, rev, updatedAt, updatedBy, deleted, fields, fieldsAt } — fieldsAt[k]
  // is the stamp of the last change to field k, carried forward when the
  // field is untouched. stories stamp per story ('stories.<sid>'); deps are
  // an OR-set stamped per add ('deps+<id>') and remove ('deps-<id>').
  function fieldsOf(entity) {
    var f = RM.clone(entity);
    delete f.id;
    RMBundle.VOLATILE_KEYS.forEach(function (k) { delete f[k]; });
    return f;
  }
  function storyMap(list) {
    var m = {};
    (list || []).forEach(function (s) { if (s && s.id != null) m[String(s.id)] = s; });
    return m;
  }
  RMBundle.wrap = function (entity, prev, userId, nowIso) {
    var fields = fieldsOf(entity);
    var pf = prev && !prev.deleted && isObj(prev.fields) ? prev.fields : {};
    var pAt = prev && isObj(prev.fieldsAt) ? prev.fieldsAt : {};
    var at = {};
    // an unchanged field keeps its old stamp (legacy envelopes without one
    // fall back to their updatedAt)
    function carried(k) { return pAt[k] || (prev && prev.updatedAt) || nowIso; }
    var keys = {};
    Object.keys(fields).concat(Object.keys(pf)).forEach(function (k) { keys[k] = true; });
    Object.keys(keys).sort().forEach(function (k) {
      if (k === 'stories') {
        var now = storyMap(fields.stories), was = storyMap(pf.stories), sids = {};
        Object.keys(now).concat(Object.keys(was)).forEach(function (s) { sids[s] = true; });
        Object.keys(sids).forEach(function (sid) {
          var same = now[sid] && was[sid] &&
            RMBundle.canonicalize(now[sid]) === RMBundle.canonicalize(was[sid]);
          at['stories.' + sid] = same ? carried('stories.' + sid) : nowIso;
        });
        return;
      }
      if (k === 'deps') {
        var have = fields.deps || [], had = pf.deps || [];
        Object.keys(pAt).forEach(function (pk) { if (/^deps[+-]/.test(pk)) at[pk] = pAt[pk]; });
        have.forEach(function (d) {
          if (!has(had, d)) at['deps+' + d] = nowIso;
          else if (!at['deps+' + d]) at['deps+' + d] = carried('deps+' + d);
        });
        had.forEach(function (d) { if (!has(have, d)) at['deps-' + d] = nowIso; });
        return;
      }
      var same = (k in pf) && RMBundle.canonicalize(pf[k]) === RMBundle.canonicalize(fields[k]);
      at[k] = same ? carried(k) : nowIso;
    });
    return {
      id: entity.id,
      rev: prev && isFinite(+prev.rev) ? +prev.rev + 1 : 1,
      updatedAt: nowIso,
      updatedBy: userId,
      deleted: false,
      fields: fields,
      fieldsAt: sortedCopy(at)
    };
  };

  // prev may be the previous envelope or just the id
  RMBundle.tombstone = function (prev, userId, nowIso) {
    return {
      id: isObj(prev) ? prev.id : prev,
      rev: isObj(prev) && isFinite(+prev.rev) ? +prev.rev + 1 : 1,
      updatedAt: nowIso,
      updatedBy: userId,
      deleted: true,
      deletedAt: nowIso,
      fields: {},
      fieldsAt: {}
    };
  };

  RMBundle.unwrap = function (env) {
    if (!env || env.deleted) return null;
    var out = { id: env.id };
    var f = RM.clone(isObj(env.fields) ? env.fields : {});
    Object.keys(f).forEach(function (k) { out[k] = f[k]; });
    return out;
  };

  // The newer of two stamped values; ties fall to the greater updatedBy, then
  // the greater canonical value — never to argument order, so merge commutes.
  function pick(a, b, sa, sb, va, vb) {
    if (sa !== sb) return sa > sb ? a : b;
    var ba = str(a.updatedBy), bb = str(b.updatedBy);
    if (ba !== bb) return ba > bb ? a : b;
    return va >= vb ? a : b;
  }

  RMBundle.mergeEntity = function (a, b) {
    if (!a) return b;
    if (!b) return a;
    var ua = str(a.updatedAt), ub = str(b.updatedAt);
    var rev = Math.max(+a.rev || 0, +b.rev || 0);
    var lead = pick(a, b, ua, ub, RMBundle.canonicalize(a), RMBundle.canonicalize(b));

    if (a.deleted || b.deleted) {
      if (a.deleted && b.deleted) {
        return {
          id: a.id, rev: rev, updatedAt: later(ua, ub), updatedBy: lead.updatedBy, deleted: true,
          deletedAt: later(str(a.deletedAt || a.updatedAt), str(b.deletedAt || b.updatedAt)),
          fields: {}, fieldsAt: {}
        };
      }
      // a tombstone loses only to an edit made after the delete
      var tomb = a.deleted ? a : b, edit = a.deleted ? b : a;
      var winner = str(edit.updatedAt) > str(tomb.deletedAt || tomb.updatedAt) ? edit : tomb;
      var out = RM.clone(winner);
      out.rev = rev;
      return out;
    }

    var fa = isObj(a.fields) ? a.fields : {}, fb = isObj(b.fields) ? b.fields : {};
    var aa = isObj(a.fieldsAt) ? a.fieldsAt : {}, ab = isObj(b.fieldsAt) ? b.fieldsAt : {};
    // a field with no stamp of its own counts as changed at the envelope's updatedAt
    function stamp(env, at, present, k) { return str(at[k] || (present ? env.updatedAt : '')); }
    var fields = {}, at = {}, keys = {};
    Object.keys(fa).concat(Object.keys(fb)).forEach(function (k) { keys[k] = true; });

    Object.keys(keys).sort().forEach(function (k) {
      if (k === 'stories' || k === 'deps') return;
      var sa = stamp(a, aa, k in fa, k), sb = stamp(b, ab, k in fb, k);
      var va = RMBundle.canonicalize(fa[k]), vb = RMBundle.canonicalize(fb[k]);
      var src = pick(a, b, sa, sb, va, vb) === a ? fa : fb;
      if (k in src && src[k] !== undefined) fields[k] = RM.clone(src[k]);
      at[k] = later(sa, sb);
    });

    if (keys.stories) {
      var sa2 = storyMap(fa.stories), sb2 = storyMap(fb.stories), sids = {};
      Object.keys(sa2).concat(Object.keys(sb2)).forEach(function (s) { sids[s] = true; });
      // a story deleted on one side survives only as a stamp
      Object.keys(aa).concat(Object.keys(ab)).forEach(function (k) {
        if (k.indexOf('stories.') === 0) sids[k.slice(8)] = true;
      });
      var merged = [];
      Object.keys(sids).sort().forEach(function (sid) {
        var k = 'stories.' + sid;
        var stA = stamp(a, aa, !!sa2[sid], k), stB = stamp(b, ab, !!sb2[sid], k);
        var src = pick(a, b, stA, stB, RMBundle.canonicalize(sa2[sid]), RMBundle.canonicalize(sb2[sid])) === a ? sa2 : sb2;
        if (src[sid]) merged.push(RM.clone(src[sid]));
        if (stA || stB) at[k] = later(stA, stB);
      });
      fields.stories = RM.sortByOrder(merged);
    }

    var depIds = {};
    (fa.deps || []).concat(fb.deps || []).forEach(function (d) { depIds[d] = true; });
    Object.keys(aa).concat(Object.keys(ab)).forEach(function (k) {
      if (/^deps[+-]/.test(k)) depIds[k.slice(5)] = true;
    });
    if (keys.deps || Object.keys(depIds).length) {
      // OR-set: present iff the latest add is newer than the latest remove
      var deps = [];
      Object.keys(depIds).sort().forEach(function (d) {
        var plus = later(stamp(a, aa, has(fa.deps, d), 'deps+' + d), stamp(b, ab, has(fb.deps, d), 'deps+' + d));
        var minus = later(str(aa['deps-' + d]), str(ab['deps-' + d]));
        if (plus && plus > minus) deps.push(d);
        if (plus) at['deps+' + d] = plus;
        if (minus) at['deps-' + d] = minus;
      });
      fields.deps = deps;
    }

    // any remaining stamp: per-key max
    Object.keys(aa).concat(Object.keys(ab)).forEach(function (k) {
      if (!(k in at)) at[k] = later(str(aa[k]), str(ab[k]));
    });

    return {
      id: a.id, rev: rev, updatedAt: later(ua, ub), updatedBy: lead.updatedBy, deleted: false,
      fields: sortedCopy(fields), fieldsAt: sortedCopy(at)
    };
  };

  // ------------------------------------------------------------ whole plans
  // prevEnvs: envelopes by id (or by kind/id) from the last flush
  RMBundle.wrapState = function (state, prevEnvs, userId, nowIso) {
    prevEnvs = prevEnvs || {};
    var out = {};
    RMBundle.KINDS.forEach(function (kind) {
      out[kind] = (state[kind] || []).map(function (ent) {
        var prev = prevEnvs[kind + '/' + ent.id] || prevEnvs[ent.id] || null;
        return RMBundle.wrap(ent, prev, userId, nowIso);
      });
    });
    return out;
  };

  // The plan's meta.json is an envelope too (id 'meta'), so a key such as
  // wsColors merges whole-key LWW through the same mergeEntity.
  RMBundle.metaEntity = function (state) {
    var ent = { id: 'meta' };
    RMBundle.META_KEYS.forEach(function (k) { if (state[k] !== undefined) ent[k] = RM.clone(state[k]); });
    return ent;
  };
  RMBundle.wrapMeta = function (state, prevEnv, userId, nowIso) {
    return RMBundle.wrap(RMBundle.metaEntity(state), prevEnv || null, userId, nowIso);
  };

  // metaShard: the meta envelope (a plain {meta, wsOrder, …} object is
  // accepted too); envsByKind: { items: [env], phases: [env], … }
  RMBundle.assembleState = function (metaShard, envsByKind) {
    var base = metaShard && isObj(metaShard.fields) ? RMBundle.unwrap(metaShard) : (metaShard || {});
    var state = {};
    RMBundle.META_KEYS.forEach(function (k) {
      if (base && base[k] !== undefined) state[k] = RM.clone(base[k]);
    });
    RMBundle.KINDS.forEach(function (kind) {
      state[kind] = ((envsByKind && envsByKind[kind]) || [])
        .map(function (env) { return RMBundle.unwrap(env); })
        .filter(Boolean);
    });
    return RM.normalizeState(state);
  };

  // What changed since the last flush. prevCanon: { 'kind/id': canonical }
  // as returned in `canon` on each changed entry. The meta entity reports as
  // kind 'meta' (never deleted).
  RMBundle.diffEntities = function (prevCanon, state) {
    prevCanon = prevCanon || {};
    var changed = [], deleted = [], seen = {};
    function check(kind, ent) {
      var key = kind + '/' + ent.id;
      seen[key] = true;
      var canon = RMBundle.canonicalize(ent);
      if (prevCanon[key] !== canon) changed.push({ kind: kind, id: ent.id, entity: ent, canon: canon });
    }
    RMBundle.KINDS.forEach(function (kind) {
      (state[kind] || []).forEach(function (ent) { check(kind, ent); });
    });
    check('meta', RMBundle.metaEntity(state));
    Object.keys(prevCanon).forEach(function (key) {
      if (seen[key]) return;
      var cut = key.indexOf('/');
      var kind = key.slice(0, cut);
      if (RMBundle.KINDS.indexOf(kind) !== -1) deleted.push({ kind: kind, id: key.slice(cut + 1) });
    });
    return { changed: changed, deleted: deleted };
  };

  // ------------------------------------------------------------ plan list
  RMBundle.newPlanEntry = function (id, name, nowIso) {
    return { id: id, name: name, createdAt: nowIso, updatedAt: nowIso };
  };
  // by id: a deleted entry always wins, else the later updatedAt (ties by
  // canonical value); sorted by createdAt then id
  RMBundle.mergePlanList = function (a, b) {
    var by = {};
    function take(p) {
      if (!p || !p.id) return;
      var cur = by[p.id];
      if (!cur) { by[p.id] = p; return; }
      if (!!cur.deleted !== !!p.deleted) { by[p.id] = cur.deleted ? cur : p; return; }
      var ta = str(cur.updatedAt), tb = str(p.updatedAt);
      if (ta !== tb) by[p.id] = ta > tb ? cur : p;
      else by[p.id] = RMBundle.canonicalize(cur) >= RMBundle.canonicalize(p) ? cur : p;
    }
    (a || []).forEach(take);
    (b || []).forEach(take);
    return Object.keys(by).map(function (id) { return RM.clone(by[id]); }).sort(function (p, q) {
      var ca = str(p.createdAt), cb = str(q.createdAt);
      if (ca !== cb) return ca < cb ? -1 : 1;
      return str(p.id) < str(q.id) ? -1 : str(p.id) > str(q.id) ? 1 : 0;
    });
  };

  // ------------------------------------------------------------ history
  // history/<userId>.jsonl — one writer per file; every line carries the
  // plan it belongs to. Line shape = a normalized history entry + userId + planId.
  RMBundle.historyLine = function (entry, userId, planId) {
    var e = RM.normalizeHistoryEntry(entry);
    if (!e) return null;
    var line = { t: e.t, u: e.u, userId: str(userId), planId: str(planId), label: e.label, n: e.n };
    if (e.d) line.d = e.d;
    if (e.x) line.x = e.x;
    if (e.tl) line.tl = e.tl;
    return line;
  };
  RMBundle.encodeHistory = function (lines) {
    lines = lines || [];
    return lines.map(function (l) { return RM.asciiJson(l); }).join('\n') + (lines.length ? '\n' : '');
  };
  // malformed lines (a torn write, a hand edit) are skipped, not fatal
  RMBundle.parseHistory = function (text) {
    var out = [];
    str(text).split(/\r?\n/).forEach(function (raw) {
      if (!raw.trim()) return;
      var obj;
      try { obj = JSON.parse(raw); } catch (e) { return; }
      if (!isObj(obj)) return;
      var line = RMBundle.historyLine(obj, obj.userId, obj.planId);
      if (line) out.push(line);
    });
    return out;
  };
  RMBundle.mergeHistory = function (lists, maxN) {
    var all = [];
    (lists || []).forEach(function (l) { (l || []).forEach(function (x) { if (x) all.push(x); }); });
    all.sort(function (p, q) {
      if (p.t !== q.t) return p.t - q.t;
      var ua = str(p.userId), ub = str(q.userId);
      if (ua !== ub) return ua < ub ? -1 : 1;
      return str(p.label) < str(q.label) ? -1 : str(p.label) > str(q.label) ? 1 : 0;
    });
    var cap = isFinite(+maxN) && +maxN > 0 ? +maxN : RM.HISTORY_MAX;
    return all.slice(-cap);
  };

  // ------------------------------------------------------------ export / convert
  // a copy fit for a standalone xlsx: bundle markers gone
  RMBundle.exportableState = function (state) {
    var out = RM.clone(state);
    RMBundle.BUNDLE_MARKERS.forEach(function (k) { delete out[k]; });
    return out;
  };

  // An open xlsx document as bundle contents: every option becomes a plan
  // (the active one first), each with its own envelopes; legacy history
  // lines all land on the active plan, in the converting user's file.
  RMBundle.migrateFromState = function (state, userId, nowIso) {
    var docs = RM.splitOptions(state);
    var activeId = state.optId || RM.uid('plan');
    var plans = {}, entries = [];
    docs.forEach(function (d, i) {
      var pid = i === 0 ? activeId : d.id;
      var plan = RMBundle.wrapState(d.doc, {}, userId, nowIso);
      plan.meta = RMBundle.wrapMeta(d.doc, null, userId, nowIso);
      plans[pid] = plan;
      entries.push(RMBundle.newPlanEntry(pid, d.name, nowIso));
    });
    var history = {};
    history[userId] = (state.history || [])
      .map(function (h) { return RMBundle.historyLine(h, userId, activeId); })
      .filter(Boolean);
    return {
      headway: {
        format: RMBundle.FORMAT,
        docId: RM.uid('doc'),
        title: state.meta && state.meta.title,
        createdAt: nowIso,
        plans: entries
      },
      plans: plans,
      history: history
    };
  };

  if (typeof module !== 'undefined' && module.exports) module.exports = RMBundle;
  root.RMBundle = RMBundle;
})(typeof window !== 'undefined' ? window : globalThis);
