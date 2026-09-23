/*
 * Headway ↔ Jira Cloud sync (REST API v3 + Agile API, Basic auth with an
 * API token).
 *
 * Pure, node-testable pieces:
 *  - discover(client, state, cfg, onProgress) -> what Jira looks like: the
 *    project's issue types, its Start date field, the account ids behind
 *    Headway's assignees, the Scrum board and its sprints, and the linked
 *    issues' current state.
 *  - plan(state, cfg, info) -> what a sync would do: epics / features /
 *    stories to create, issues to update, dependency links, sprints to
 *    create and issues to place in them, and done-flags to pull back.
 *  - apply(plan, client, onProgress) -> Promise<result>: runs the plan;
 *    never touches Headway state — the caller commits the returned keys.
 *  - adf(text): plain text -> Atlassian Document Format.
 *
 * Milestones never sync: they are dates on Headway's timeline, not work.
 * Every synced issue carries Start date + Due date so Jira's timeline
 * shows it; the Start date field is found automatically.
 *
 * Browser glue (window.HeadwayJira): a Setup tab (login is stored on this
 * machine only; site, project and types live in the document as
 * meta.jira), the Sync dialog (previews the plan), a background job with a
 * topbar status button.
 *
 * Transport: Tauri's http plugin in the desktop app (Jira Cloud sends no
 * CORS headers, so a plain browser page cannot call it); fetch otherwise.
 */
(function (root) {
  'use strict';

  var RM = root.RM || (typeof require !== 'undefined' ? require('./core.js') : null);
  var JR = {};

  JR.DEFAULTS = { pushEpics: true, pushStories: true, sprints: true, auto: true, startField: '' };
  JR.lastTypes = null; // the last discovered project's resolved types, for the settings card
  JR.AUTO_MS = 5 * 60 * 1000; // auto-sync cadence once a sync has succeeded
  // Your Jira login (email + API token) lives ONLY in this machine's local
  // storage under this key. The document carries the shared part (site,
  // project, issue types) in meta.jira and never the token.
  JR.LOCAL_KEY = 'headway-jira-v1';

  // ------------------------------------------------------------ helpers
  function slug(s) {
    return String(s || '').toLowerCase().replace(/[^a-z0-9]+/g, '-').replace(/^-+|-+$/g, '');
  }
  // free-form tags push as plain slugged labels
  function tagLabels(o) {
    return ((o && o.tags) || []).map(slug);
  }
  function section(title, html) {
    var t = RM.htmlToText(html);
    return t ? title + ':\n' + t : '';
  }
  function joinSections(parts) {
    return parts.filter(Boolean).join('\n\n');
  }
  function sched(x) { return x && x.startDay != null && x.durDays != null; }
  function iso(meta, day) { return RM.fmtISO(RM.dayToDate(meta, day)); }
  function lc(s) { return String(s || '').trim().toLowerCase(); }
  // readable text for whatever was thrown: Error, string, or a Jira body
  function errText(e) {
    if (e == null) return 'unknown error';
    if (typeof e === 'string') return e;
    if (typeof e.message === 'string' && e.message) return e.message;
    if (e.message != null) return errText(e.message);
    try { return JSON.stringify(e); } catch (x) { return String(e); }
  }
  JR.errText = errText;
  // the messages inside a Jira error body: errorMessages[] plus errors{} —
  // and, for a bulk response, each element's own errorMessages / errors
  function bodyMessages(data) {
    var msg = [];
    if (!data || typeof data !== 'object') return msg;
    (Array.isArray(data.errorMessages) ? data.errorMessages : []).forEach(function (m) { msg.push(errText(m)); });
    if (Array.isArray(data.errors)) {
      data.errors.forEach(function (er) {
        var ee = (er && er.elementErrors) || er || {};
        var inner = bodyMessages(ee);
        msg.push((er && er.failedElementNumber != null ? 'item ' + (er.failedElementNumber + 1) + ': ' : '') + (inner.join('; ') || errText(er)));
      });
    } else if (data.errors && typeof data.errors === 'object') {
      Object.keys(data.errors).forEach(function (k) { msg.push(k + ': ' + errText(data.errors[k])); });
    }
    return msg;
  }
  JR.bodyMessages = bodyMessages;
  function cfgOf(state, creds) {
    var out = {};
    Object.keys(JR.DEFAULTS).forEach(function (k) { out[k] = JR.DEFAULTS[k]; });
    var doc = (state && state.meta && state.meta.jira) || {};
    Object.keys(doc).forEach(function (k) { if (doc[k] != null && doc[k] !== '') out[k] = doc[k]; });
    Object.keys(creds || {}).forEach(function (k) { if (k !== 'site' || !out.site) out[k] = creds[k]; });
    return out;
  }
  JR.cfgOf = cfgOf;
  function memberById(state, id) {
    var team = (state && state.team) || [];
    for (var i = 0; i < team.length; i++) if (team[i].id === id) return team[i];
    return null;
  }
  // the person an issue goes to: Headway's first assignee (Jira holds one)
  function assigneeLabel(state, x) {
    var m = memberById(state, (x && x.assignees || [])[0]);
    return m ? RM.memberLabel(m) : '';
  }

  // Plain text -> ADF: blank lines split paragraphs, single newlines break lines
  JR.adf = function (text) {
    var content = [];
    String(text || '').replace(/\r\n?/g, '\n').split(/\n{2,}/).forEach(function (para) {
      if (!para.trim()) return;
      var nodes = [];
      para.split('\n').forEach(function (line, i) {
        if (i) nodes.push({ type: 'hardBreak' });
        if (line) nodes.push({ type: 'text', text: line });
      });
      content.push({ type: 'paragraph', content: nodes });
    });
    return { type: 'doc', version: 1, content: content };
  };

  // Headway priority letters -> Jira's default priority names
  JR.PRIORITY_NAMES = {
    levels: { C: 'Highest', H: 'High', M: 'Medium', L: 'Low' },
    moscow: { M: 'Highest', S: 'High', C: 'Medium', W: 'Low' }
  };
  function priorityField(scheme, v) {
    var map = JR.PRIORITY_NAMES[scheme];
    return map && v && map[v] ? { name: map[v] } : null;
  }

  // ------------------------------------------------------------ fields
  // info (optional): discover() output — resolved type names, the Start
  // date field id and the assignee account ids
  // the Jira issue type for an epic name / item / story
  function typeName(info, state, obj, kind) {
    var key = RM.typeOf(state, obj, kind).key;
    var r = info && info.types && info.types.byKey && info.types.byKey[key];
    return r ? r.name : RM.jiraTypeName(state, key);
  }
  function isSubtask(info, state, st) {
    var key = RM.typeOf(state, st, 'story').key;
    var r = info && info.types && info.types.byKey && info.types.byKey[key];
    // no project info yet: assume the default Sub-task nesting
    return r ? !!r.subtask : true;
  }
  function withDates(f, meta, startDay, durDays, deadline, info) {
    var on = startDay != null && durDays != null;
    var start = on ? iso(meta, startDay) : null;
    var end = deadline || (on ? RM.fmtISO(RM.spanEndDate(meta, startDay, Math.max(1, durDays))) : null);
    if (end) f.duedate = end;
    var sf = info && info.startField;
    if (sf && start) f[sf] = start;
  }
  function withAssignee(f, info, label) {
    var id = info && info.accounts && label ? info.accounts[label] : null;
    if (id) f.assignee = { accountId: id };
  }
  JR.featureFields = function (state, it, cfg, parentKey, info) {
    var meta = state.meta;
    var phase = null;
    state.phases.forEach(function (p) { if (p.id === it.phaseId) phase = p; });
    var checklist = cfg.pushStories ? '' : it.stories.map(function (s) {
      return (s.done ? '[x] ' : '[ ] ') + s.title;
    }).join('\n');
    var f = {
      project: { key: cfg.project },
      issuetype: { name: typeName(info, state, it, 'feature') },
      summary: it.feature || '(untitled)',
      description: JR.adf(joinSections([
        RM.htmlToText(it.description),
        checklist ? 'Stories:\n' + checklist : '',
        section('Enables', it.enables),
        section('Out of scope', it.outOfScope),
        section('External dependencies', it.extDeps),
        section('Notes', it.notes)
      ])),
      labels: [
        it.workstream ? 'ws-' + slug(it.workstream) : '',
        phase ? 'phase-' + slug(phase.name) : '',
        it.size ? 'size-' + slug(it.size) : ''
      ].concat(tagLabels(it)).filter(Boolean)
    };
    var pr = priorityField(RM.prioritySchemeOf(state), it.priority);
    if (pr) f.priority = pr;
    withDates(f, meta, it.startDay, it.durDays, it.deadline, info);
    withAssignee(f, info, assigneeLabel(state, it));
    if (parentKey) f.parent = { key: parentKey };
    return f;
  };
  JR.storyFields = function (state, it, st, cfg, parentKey, info) {
    var meta = state.meta;
    var f = {
      project: { key: cfg.project },
      issuetype: { name: typeName(info, state, st, 'story') },
      summary: st.title || '(untitled)',
      description: JR.adf(joinSections([
        RM.htmlToText(st.description),
        section('Acceptance criteria', st.ac)
      ])),
      labels: ['feature-' + slug(it.feature), it.workstream ? 'ws-' + slug(it.workstream) : '']
        .concat(tagLabels(st)).filter(Boolean)
    };
    var pr = priorityField(RM.prioritySchemeOf(state, 'story'), st.priority);
    if (pr) f.priority = pr;
    // a story without its own timeline rides along with the feature's dates
    var own = st.startDay != null && st.durDays > 0;
    if (own) withDates(f, meta, st.startDay, st.durDays, st.deadline, info);
    else withDates(f, meta, it.startDay, it.durDays, st.deadline || it.deadline, info);
    withAssignee(f, info, assigneeLabel(state, st) || assigneeLabel(state, it));
    if (parentKey) f.parent = { key: parentKey };
    return f;
  };
  // an epic's dates: the first start and the last end/deadline of its
  // scheduled, non-milestone features — Jira's timeline draws epic bars from
  // the epic's own Start date / Due date, not from the children
  JR.epicSpan = function (state, name) {
    var meta = state.meta, lo = null, hi = null;
    state.items.forEach(function (it) {
      if (it.epic !== name || it.milestone || !sched(it)) return;
      var start = iso(meta, it.startDay);
      var end = it.deadline || RM.fmtISO(RM.spanEndDate(meta, it.startDay, Math.max(1, it.durDays)));
      if (lo == null || start < lo) lo = start;
      if (hi == null || end > hi) hi = end;
    });
    return lo ? { start: lo, end: hi } : null;
  };
  JR.epicFields = function (name, cfg, info, state) {
    var f = { project: { key: cfg.project }, issuetype: { name: typeName(info, state, name, 'epic') }, summary: name };
    var span = state ? JR.epicSpan(state, name) : null;
    if (span) {
      f.duedate = span.end;
      if (info && info.startField) f[info.startField] = span.start;
    }
    return f;
  };

  // ------------------------------------------------------------ sprints
  // Headway sprint number for a day, and the calendar span of a sprint
  JR.sprintOfDay = function (meta, day) {
    return RM.sprintNumForWeek(meta, Math.floor(day / RM.slotsOf(meta)));
  };
  JR.sprintSpan = function (meta, num) {
    var r = RM.sprintRange(meta, num);
    var start = RM.weekStartDate(meta, r.w0);
    var end = new Date(RM.weekStartDate(meta, r.w1).getTime() - 86400000);
    return { num: num, name: 'Sprint ' + num, start: RM.fmtISO(start), end: RM.fmtISO(end) };
  };
  // the Jira sprint whose dates hold the Headway sprint's first day, else
  // one named the same way
  function matchSprint(sprints, span) {
    var t = Date.parse(span.start + 'T12:00:00Z');
    for (var i = 0; i < sprints.length; i++) {
      var s = sprints[i];
      var a = Date.parse(s.startDate || ''), b = Date.parse(s.endDate || '');
      if (a && b && a <= t && t < b) return s;
    }
    for (i = 0; i < sprints.length; i++) if (lc(sprints[i].name) === lc(span.name)) return sprints[i];
    return null;
  }

  // ------------------------------------------------------------ plan
  // info: discover() output ({ remote, types, startField, accounts, board,
  // notes }); an object with only `remote` is fine for tests.
  JR.plan = function (state, cfg, info) {
    info = info || {};
    // a bare { KEY: issue } map (the older signature) still works as `remote`
    var bare = !('remote' in info) && !('types' in info) && !('board' in info) && !('accounts' in info);
    var remote = bare ? info : (info.remote || {});
    if (bare) info = { remote: remote };
    var meta = state.meta;
    var plan = { epics: [], features: [], stories: [], updates: [], links: [], storyLinks: [], pulls: [], transitions: [], missing: [], milestones: [],
      sprints: null, notes: (info.notes || []).slice() };
    // stories nest under their feature only as sub-tasks; any other type
    // sits beside the feature (same level) and is tied to it with a link
    var anyFlat = false;
    var epicKey = {};
    Object.keys(state.epicJira || {}).forEach(function (e) { if (state.epicJira[e]) epicKey[e] = state.epicJira[e]; });
    var work = state.items.filter(function (it) {
      if (it.milestone) { plan.milestones.push({ id: it.id, num: it.num, title: it.feature, key: it.jiraKey || '' }); return false; }
      return true;
    });
    if (cfg.pushEpics) {
      var seen = {};
      work.forEach(function (it) {
        if (!it.epic || epicKey[it.epic] || seen[it.epic]) return;
        seen[it.epic] = true;
        plan.epics.push({ name: it.epic, fields: JR.epicFields(it.epic, cfg, info, state) });
      });
    }
    // epics already in Jira get their rolled-up dates refreshed
    Object.keys(epicKey).forEach(function (name) {
      if (!work.some(function (it) { return it.epic === name; })) return;
      var span = JR.epicSpan(state, name);
      if (!span) return;
      var ef = { duedate: span.end };
      if (info.startField) ef[info.startField] = span.start;
      plan.updates.push({ kind: 'epic', id: 'epic:' + name, title: name, key: epicKey[name], fields: ef });
    });
    var sprintsOn = !!(cfg.sprints && info.board && RM.sprintsEnabled(meta));
    if (sprintsOn) plan.sprints = { boardId: info.board.id, boardName: info.board.name || '', create: [], assign: [] };
    var newSprint = {};
    function placeInSprint(ref, day) {
      if (!sprintsOn || day == null) return;
      var num = JR.sprintOfDay(meta, day);
      var span = JR.sprintSpan(meta, num);
      var found = matchSprint(info.board.sprints || [], span);
      if (!found && !newSprint[num]) { newSprint[num] = true; plan.sprints.create.push(span); }
      plan.sprints.assign.push({ ref: ref, num: num, sprintId: found ? found.id : null });
    }
    work.forEach(function (it) {
      var parent = it.epic ? (epicKey[it.epic] || null) : null;
      var fields = JR.featureFields(state, it, cfg, parent, info);
      var entry = { id: it.id, num: it.num, title: it.feature, epic: it.epic || '', fields: fields, kind: 'feature' };
      if (it.jiraKey) {
        if (remote[it.jiraKey] === null) { plan.missing.push({ kind: 'feature', id: it.id, key: it.jiraKey, title: it.feature }); return; }
        // creates carry project/type; updates must not (Jira rejects type moves here)
        var uf = {}; Object.keys(fields).forEach(function (k) { if (k !== 'project' && k !== 'issuetype') uf[k] = fields[k]; });
        entry.key = it.jiraKey; entry.fields = uf;
        plan.updates.push(entry);
        // done is sticky in both directions: Jira's Done comes back to
        // Headway; Headway's Done moves the Jira issue to a done status
        var r = remote[it.jiraKey];
        if (r && r.done && !it.done) plan.pulls.push({ kind: 'feature', id: it.id, key: it.jiraKey, title: it.feature, done: true, status: r.status });
        if (r && !r.done && it.done) plan.transitions.push({ kind: 'feature', id: it.id, key: it.jiraKey, title: it.feature });
      } else {
        plan.features.push(entry);
        if (it.done) plan.transitions.push({ kind: 'feature', id: it.id, key: null, title: it.feature });
      }
      if (sched(it)) placeInSprint({ kind: 'feature', id: it.id }, it.startDay);
      if (cfg.pushStories) {
        it.stories.forEach(function (st) {
          var nest = isSubtask(info, state, st);
          if (!nest) anyFlat = true;
          var sf = JR.storyFields(state, it, st, cfg, nest ? (it.jiraKey || null) : null, info);
          var se = { id: st.id, itemId: it.id, title: st.title, fields: sf, kind: 'story', nest: nest };
          if (!nest) plan.storyLinks.push({ storyId: st.id, itemId: it.id });
          if (st.jiraKey) {
            if (remote[st.jiraKey] === null) { plan.missing.push({ kind: 'story', id: st.id, itemId: it.id, key: st.jiraKey, title: st.title }); return; }
            var suf = {}; Object.keys(sf).forEach(function (k) { if (k !== 'project' && k !== 'issuetype') suf[k] = sf[k]; });
            se.key = st.jiraKey; se.fields = suf;
            plan.updates.push(se);
            var sr = remote[st.jiraKey];
            if (sr && sr.done && !st.done) plan.pulls.push({ kind: 'story', id: st.id, itemId: it.id, key: st.jiraKey, title: st.title, done: true, status: sr.status });
            if (sr && !sr.done && st.done) plan.transitions.push({ kind: 'story', id: st.id, key: st.jiraKey, title: st.title });
          } else {
            plan.stories.push(se);
            if (st.done) plan.transitions.push({ kind: 'story', id: st.id, key: null, title: st.title });
          }
          // sub-tasks follow their parent's sprint in Jira; other story
          // types are placed by their own start, else the feature's
          if (!nest) {
            var day = st.startDay != null && st.durDays > 0 ? st.startDay : (sched(it) ? it.startDay : null);
            placeInSprint({ kind: 'story', id: st.id }, day);
          }
        });
      }
    });
    if (cfg.pushStories && anyFlat && info.types && info.types.known) {
      plan.notes.push('Some ' + RM.levelLabel(state, 'story', true).toLowerCase() + ' use a type that is not a sub-task in Jira; they are created beside their ' +
        RM.levelLabel(state, 'feature').toLowerCase() + ', linked to it and labelled feature-….');
    }
    // dependency links: "X blocks Y" for every dep whose two ends will both
    // have keys once creates land (resolved in apply); milestones excluded
    var workId = {};
    work.forEach(function (it) { workId[it.id] = true; });
    work.forEach(function (it) {
      it.deps.forEach(function (d) {
        // deps hold item ids (a legacy number resolves by num)
        var dep = RM.itemById(state, d) || RM.itemByNum(state, +d);
        if (!dep || !workId[dep.id]) return;
        plan.links.push({ blockerNum: dep.num, blockerId: dep.id, blockedNum: it.num, blockedId: it.id });
      });
    });
    // story -> story dependencies become the same Blocks link, as long as
    // both stories are being pushed (their features are in this run)
    if (cfg.pushStories) {
      work.forEach(function (it) {
        it.stories.forEach(function (st) {
          var res = RM.resolveStoryDeps(state, st);
          res.deps.forEach(function (ref) {
            if (!workId[ref.it.id]) return;
            plan.links.push({
              story: true,
              blockerNum: ref.st.num, blockerId: ref.st.id, blockerTitle: ref.st.title,
              blockedNum: st.num, blockedId: st.id, blockedTitle: st.title
            });
          });
        });
      });
    }
    // people: who resolved, who did not
    var people = {}, unresolved = [];
    work.forEach(function (it) {
      var l = assigneeLabel(state, it);
      if (l) people[l] = true;
      if (cfg.pushStories) it.stories.forEach(function (st) { var sl = assigneeLabel(state, st); if (sl) people[sl] = true; });
    });
    Object.keys(people).forEach(function (l) { if (!(info.accounts && info.accounts[l])) unresolved.push(l); });
    plan.people = { total: Object.keys(people).length, unresolved: unresolved };
    plan.counts = {
      create: plan.epics.length + plan.features.length + plan.stories.length,
      update: plan.updates.length, link: plan.links.length, pull: plan.pulls.length, missing: plan.missing.length,
      done: plan.transitions.length,
      sprintsCreate: plan.sprints ? plan.sprints.create.length : 0,
      sprintAssign: plan.sprints ? plan.sprints.assign.length : 0
    };
    return plan;
  };
  // the keys a plan wants to look at in Jira (milestones excluded)
  JR.keysOf = function (state, cfg) {
    var keys = [];
    state.items.forEach(function (it) {
      if (it.milestone) return;
      if (it.jiraKey) keys.push(it.jiraKey);
      if (cfg.pushStories) it.stories.forEach(function (st) { if (st.jiraKey) keys.push(st.jiraKey); });
    });
    return keys;
  };
  // the assignee names a sync needs account ids for
  JR.peopleOf = function (state, cfg) {
    var seen = {}, out = [];
    state.items.forEach(function (it) {
      if (it.milestone) return;
      var l = assigneeLabel(state, it);
      if (l && !seen[l]) { seen[l] = true; out.push(l); }
      if (cfg.pushStories) it.stories.forEach(function (st) {
        var sl = assigneeLabel(state, st);
        if (sl && !seen[sl]) { seen[sl] = true; out.push(sl); }
      });
    });
    return out;
  };

  // ------------------------------------------------------------ client
  function b64(s) {
    if (typeof btoa === 'function') return btoa(unescape(encodeURIComponent(s)));
    return Buffer.from(s, 'utf8').toString('base64');
  }
  JR.siteUrl = function (site) {
    var s = String(site || '').trim().replace(/\/+$/, '');
    if (!s) return '';
    if (!/^https?:\/\//i.test(s)) s = 'https://' + s;
    return s;
  };
  // pick the transport: Tauri's http plugin (no CORS) when present
  JR.fetchImpl = null;
  function fetchFn() {
    if (JR.fetchImpl) return JR.fetchImpl;
    var t = root.__TAURI__;
    if (t && t.http && t.http.fetch) return t.http.fetch;
    return root.fetch ? root.fetch.bind(root) : null;
  }
  JR.client = function (creds) {
    var base = JR.siteUrl(creds.site);
    var auth = 'Basic ' + b64((creds.email || '') + ':' + (creds.token || ''));
    function request(method, path, body) {
      var f = fetchFn();
      if (!f) return Promise.reject(new Error('No HTTP transport available'));
      if (!base) return Promise.reject(new Error('Jira site URL is not set'));
      var opts = { method: method, headers: { Authorization: auth, Accept: 'application/json' } };
      if (body !== undefined) { opts.headers['Content-Type'] = 'application/json'; opts.body = JSON.stringify(body); }
      return f(base + path, opts).then(function (res) {
        return res.text().then(function (txt) {
          var data = null;
          try { data = txt ? JSON.parse(txt) : null; } catch (e) { data = null; }
          if (!res.ok) {
            var msg = bodyMessages(data);
            var err = new Error(msg.length ? msg.join('; ') : 'HTTP ' + res.status + (txt ? ' ' + txt.slice(0, 200) : ''));
            err.status = res.status;
            err.data = data;
            err.fields = data && data.errors && !Array.isArray(data.errors) && typeof data.errors === 'object' ? Object.keys(data.errors) : [];
            throw err;
          }
          return data;
        });
      }, function (e) {
        // the transport itself failed (no network, blocked URL, CORS…)
        if (e instanceof Error) throw e;
        throw new Error(errText(e));
      });
    }
    return {
      get: function (p) { return request('GET', p); },
      post: function (p, b) { return request('POST', p, b); },
      put: function (p, b) { return request('PUT', p, b); }
    };
  };

  // ------------------------------------------------------------ discover
  // { KEY: { summary, done, status } | null (not found) } — one GET per key,
  // a few in flight; a bad key must not sink the whole sync
  JR.fetchRemote = function (client, keys, onProgress) {
    var out = {}, i = 0, done = 0;
    var uniq = keys.filter(function (k, idx) { return keys.indexOf(k) === idx; });
    function next() {
      if (i >= uniq.length) return Promise.resolve();
      var key = uniq[i++];
      return client.get('/rest/api/3/issue/' + encodeURIComponent(key) + '?fields=summary,status').then(function (iss) {
        var st = iss && iss.fields && iss.fields.status;
        out[key] = { summary: iss && iss.fields ? iss.fields.summary : '',
          done: !!(st && st.statusCategory && st.statusCategory.key === 'done'), status: st ? st.name : '' };
      }, function (err) {
        if (err && err.status === 404) out[key] = null; else throw err;
      }).then(function () {
        done += 1;
        if (onProgress) onProgress('Reading ' + done + ' of ' + uniq.length + ' linked issues…');
        return next();
      });
    }
    var lanes = [];
    for (var l = 0; l < 4; l++) lanes.push(next());
    return Promise.all(lanes).then(function () { return out; });
  };
  // the project's issue types resolved per Headway type key:
  // { byKey: { key: { name, subtask } }, known, notes }
  JR.resolveTypes = function (issueTypes, state) {
    var list = Array.isArray(issueTypes) ? issueTypes : [];
    function byName(n) { return n ? list.filter(function (t) { return lc(t.name) === lc(n); })[0] : null; }
    function first(pred) { return list.filter(pred)[0]; }
    var epicKeys = RM.levelOf(state, 'epic').types, storyKeys = RM.levelOf(state, 'story').types;
    var out = { byKey: {}, known: list.length > 0, notes: [] };
    RM.itemTypes(state).forEach(function (t) {
      var want = t.jira || t.label;
      var hit = byName(want);
      if (!hit && list.length) {
        if (epicKeys.indexOf(t.key) !== -1) hit = first(function (x) { return x.hierarchyLevel === 1; }) || byName('Epic');
        else if (storyKeys.indexOf(t.key) !== -1) hit = first(function (x) { return !!x.subtask; }) || byName('Story') || byName('Task');
        if (!hit) hit = byName('Story') || byName('Task') || first(function (x) { return !x.subtask && x.hierarchyLevel !== 1 && lc(x.name) !== 'epic'; });
        if (hit) out.notes.push('Type “' + want + '” (' + t.label + ') is not in the project; using “' + hit.name + '”');
      }
      out.byKey[t.key] = { name: hit ? hit.name : want, subtask: !!(hit && hit.subtask) };
    });
    return out;
  };
  // Jira's Start date field: the configured id, else the field named so
  JR.findStartField = function (fields, configured) {
    var list = Array.isArray(fields) ? fields : [];
    if (configured) {
      var hit = list.filter(function (f) { return f.id === configured; })[0];
      return { id: configured, name: hit ? hit.name : configured };
    }
    var cands = list.filter(function (f) {
      var t = f.schema && f.schema.type;
      return lc(f.name) === 'start date' && (t === 'date' || t === 'datetime' || !t);
    });
    cands.sort(function (a, b) { return (a.custom ? 1 : 0) - (b.custom ? 1 : 0); });
    return cands.length ? { id: cands[0].id, name: cands[0].name } : null;
  };
  // the Jira user behind a Headway name: an exact display name or email,
  // else the one candidate carrying every word of the name, else the sole
  // candidate; null when it stays ambiguous. Returns the user object.
  JR.pickUser = function (users, label) {
    var list = Array.isArray(users) ? users : [];
    var want = lc(label);
    var exact = list.filter(function (u) { return lc(u.displayName) === want || lc(u.emailAddress) === want; });
    if (exact.length === 1) return exact[0];
    var words = want.split(/[\s,.]+/).filter(function (w) { return w.length > 1; });
    if (words.length) {
      var all = list.filter(function (u) {
        var hay = lc(u.displayName) + ' ' + lc(u.emailAddress);
        return words.every(function (w) { return hay.indexOf(w) !== -1; });
      });
      if (all.length === 1) return all[0];
    }
    if (list.length === 1) return list[0];
    return null;
  };
  JR.discover = function (client, state, cfg, onProgress) {
    var progress = onProgress || function () {};
    var info = { remote: {}, types: null, startField: cfg.startField || null, startFieldName: '', accounts: {}, accountNames: {}, board: null, notes: [] };
    var cached = (state.meta.jira && state.meta.jira.accounts) || {};
    var cachedNames = (state.meta.jira && state.meta.jira.accountNames) || {};
    return Promise.resolve()
      .then(function () {
        progress('Reading project ' + cfg.project + '…');
        return client.get('/rest/api/3/project/' + encodeURIComponent(cfg.project)).then(function (p) {
          info.types = JR.resolveTypes((p && p.issueTypes) || [], state);
          info.types.notes.forEach(function (n) { info.notes.push(n); });
          JR.lastTypes = info.types;
        }, function (err) { info.notes.push('Could not read the project: ' + errText(err)); });
      })
      .then(function () {
        progress('Finding the Start date field…');
        return client.get('/rest/api/3/field').then(function (fields) {
          var sf = JR.findStartField(fields, cfg.startField);
          if (sf) { info.startField = sf.id; info.startFieldName = sf.name; }
          else info.notes.push('No “Start date” field in this Jira site — only due dates will be set. Enter the field id in Setup → Jira if it goes by another name.');
        }, function (err) { info.notes.push('Could not list fields: ' + errText(err)); });
      })
      .then(function () {
        var people = JR.peopleOf(state, cfg);
        var i = 0;
        function next() {
          if (i >= people.length) return Promise.resolve();
          var label = people[i++];
          if (cached[label]) { info.accounts[label] = cached[label]; info.accountNames[label] = cachedNames[label] || ''; return next(); }
          progress('Matching people (' + i + ' of ' + people.length + ')…');
          return client.get('/rest/api/3/user/assignable/search?project=' + encodeURIComponent(cfg.project) +
            '&query=' + encodeURIComponent(label) + '&maxResults=20').then(function (users) {
            var u = JR.pickUser(users, label);
            info.accounts[label] = u ? u.accountId : null;
            info.accountNames[label] = u ? (u.displayName || '') : '';
          }, function (err) {
            info.accounts[label] = null;
            if (i === 1) info.notes.push('Could not look up people: ' + errText(err));
          }).then(next);
        }
        return next();
      })
      .then(function () {
        if (!cfg.sprints || !RM.sprintsEnabled(state.meta)) return;
        progress('Finding the Scrum board…');
        return client.get('/rest/agile/1.0/board?projectKeyOrId=' + encodeURIComponent(cfg.project) + '&type=scrum').then(function (res) {
          var boards = res && Array.isArray(res.values) ? res.values : [];
          if (!boards.length) { info.notes.push('No Scrum board for ' + cfg.project + ' — issues are not placed in sprints'); return; }
          info.board = { id: boards[0].id, name: boards[0].name, sprints: [] };
          function page(startAt) {
            return client.get('/rest/agile/1.0/board/' + info.board.id + '/sprint?startAt=' + startAt + '&maxResults=50').then(function (r) {
              var vals = r && Array.isArray(r.values) ? r.values : [];
              info.board.sprints = info.board.sprints.concat(vals);
              if (r && r.isLast === false && vals.length) return page(startAt + vals.length);
            });
          }
          return page(0);
        }, function (err) { info.notes.push('Could not read boards: ' + errText(err)); });
      })
      .then(function () {
        return JR.fetchRemote(client, JR.keysOf(state, cfg), progress).then(function (remote) { info.remote = remote; });
      })
      .then(function () { return info; });
  };
  JR.testConnection = function (creds, project) {
    var c = JR.client(creds);
    return c.get('/rest/api/3/myself').then(function (me) {
      var who = (me && me.displayName) || (me && me.emailAddress) || 'signed in';
      if (!project) return { user: who };
      return c.get('/rest/api/3/project/' + encodeURIComponent(project)).then(function (p) {
        return { user: who, project: (p && p.name) || project, issueTypes: (p && p.issueTypes) || [] };
      });
    });
  };

  // ------------------------------------------------------------ apply
  function chunks(arr, n) {
    var out = [];
    for (var i = 0; i < arr.length; i += n) out.push(arr.slice(i, i + n));
    return out;
  }
  // a field Jira refuses (typically "cannot be set — not on the screen") is
  // dropped from the batch and the failed entries go once more without it
  function fieldErrors(errObj) {
    if (!errObj || typeof errObj !== 'object' || Array.isArray(errObj)) return [];
    return Object.keys(errObj).filter(function (k) { return /cannot be set|not on the appropriate screen|unknown/i.test(errText(errObj[k])); });
  }
  // Jira refusing the parent ("Please select valid parent issue") means the
  // type cannot nest there — the entry goes again on its own
  function parentRefused(errObj) {
    if (!errObj || typeof errObj !== 'object' || Array.isArray(errObj)) return false;
    return Object.keys(errObj).some(function (k) { return /^parent/i.test(k) && /parent/i.test(errText(errObj[k])); });
  }
  function dropField(entries, field) {
    entries.forEach(function (e) { delete e.fields[field]; });
  }
  // POST /issue/bulk keeps input order: issues[] carry the keys, errors[]
  // point back at failedElementNumber
  function bulkCreate(client, entries, label, result, onProgress, dropped) {
    var made = 0;
    function run(batch, retry) {
      return client.post('/rest/api/3/issue/bulk', { issueUpdates: batch.map(function (e) { return { fields: e.fields }; }) })
        .then(null, function (err) {
          // every element failed: Jira answers 400 with the same issues[] /
          // errors[] body a partial failure would carry
          if (err && err.data && Array.isArray(err.data.errors)) return { issues: err.data.issues || [], errors: err.data.errors };
          throw err;
        })
        .then(function (res) {
          var issues = (res && res.issues) || [];
          var errs = (res && res.errors) || [];
          var failed = {}, again = [], badFields = {};
          errs.forEach(function (er) {
            failed[er.failedElementNumber] = true;
            var el = batch[er.failedElementNumber];
            var ee = er.elementErrors || {};
            var fe = fieldErrors(ee.errors);
            if (fe.length && !retry) { fe.forEach(function (f) { badFields[f] = true; }); if (el) again.push(el); return; }
            if (el && el.fields.parent && parentRefused(ee.errors) && !retry) {
              delete el.fields.parent;
              el.unnested = true; // it will be linked to its feature instead
              result.parentRefused = (result.parentRefused || 0) + 1;
              again.push(el);
              return;
            }
            var m = bodyMessages(ee);
            result.errors.push((el ? (el.title || el.name) : label) + ': ' + (m.join('; ') || 'not created'));
          });
          var ii = 0;
          batch.forEach(function (e, idx) {
            if (failed[idx]) return;
            var iss = issues[ii++];
            if (iss && iss.key) { e.key = iss.key; made += 1; }
            else result.errors.push((e.title || e.name) + ': no key returned');
          });
          var bf = Object.keys(badFields);
          if (bf.length) bf.forEach(function (f) { dropped[f] = true; dropField(entries, f); });
          if (again.length) return run(again, true);
        });
    }
    return chunks(entries, 50).reduce(function (p, batch) {
      return p.then(function () {
        if (onProgress) onProgress('Creating ' + label + ' (' + made + ' of ' + entries.length + ')…');
        return run(batch, false);
      });
    }, Promise.resolve()).then(function () { return made; });
  }
  JR.apply = function (plan, client, onProgress) {
    var result = { epicKeys: {}, itemKeys: {}, storyKeys: {}, created: 0, updated: 0, linked: 0, storyLinked: 0, transitioned: 0,
      sprintsCreated: 0, sprintAssigned: 0, sprintIds: {}, errors: [], droppedFields: [] };
    var progress = onProgress || function () {};
    var dropped = {};
    return Promise.resolve()
      // 1. epics
      .then(function () { return bulkCreate(client, plan.epics, 'epics', result, progress, dropped); })
      .then(function (n) {
        result.created += n;
        plan.epics.forEach(function (e) { if (e.key) result.epicKeys[e.name] = e.key; });
        // features under a freshly created epic get their parent now
        plan.features.forEach(function (f) {
          if (f.epic && result.epicKeys[f.epic]) f.fields.parent = { key: result.epicKeys[f.epic] };
        });
        plan.updates.forEach(function (u) {
          if (u.kind === 'feature' && u.epic && result.epicKeys[u.epic]) u.fields.parent = { key: result.epicKeys[u.epic] };
        });
        Object.keys(dropped).forEach(function (f) { dropField(plan.features, f); dropField(plan.stories, f); dropField(plan.updates, f); });
        return bulkCreate(client, plan.features, 'features', result, progress, dropped);
      })
      // 2. features
      .then(function (n) {
        result.created += n;
        plan.features.forEach(function (f) { if (f.key) result.itemKeys[f.id] = f.key; });
        plan.stories.forEach(function (s) {
          if (s.nest !== false && !s.fields.parent && result.itemKeys[s.itemId]) s.fields.parent = { key: result.itemKeys[s.itemId] };
        });
        // a nesting story whose feature never got a key cannot be created
        var ready = plan.stories.filter(function (s) { return s.nest === false || !!s.fields.parent; });
        plan.stories.forEach(function (s) { if (s.nest !== false && !s.fields.parent) result.errors.push(s.title + ': feature has no Jira key yet'); });
        Object.keys(dropped).forEach(function (f) { dropField(plan.stories, f); dropField(plan.updates, f); });
        return bulkCreate(client, ready, 'stories', result, progress, dropped);
      })
      // 3. stories, then updates one by one
      .then(function (n) {
        result.created += n;
        plan.stories.forEach(function (s) { if (s.key) result.storyKeys[s.id] = s.key; });
        Object.keys(dropped).forEach(function (f) { dropField(plan.updates, f); });
        var i = 0;
        function put(u, retry) {
          return client.put('/rest/api/3/issue/' + encodeURIComponent(u.key), { fields: u.fields })
            .then(function () { result.updated += 1; }, function (err) {
              var eo = err.data && err.data.errors && !Array.isArray(err.data.errors) ? err.data.errors : {};
              var fe = fieldErrors(eo);
              if (fe.length && !retry) {
                fe.forEach(function (f) { dropped[f] = true; dropField(plan.updates, f); });
                return put(u, true);
              }
              if (u.fields.parent && parentRefused(eo) && !retry) {
                delete u.fields.parent;
                return put(u, true);
              }
              result.errors.push(u.key + ' ' + u.title + ': ' + errText(err));
            });
        }
        return plan.updates.reduce(function (p, u) {
          return p.then(function () {
            i += 1;
            progress('Updating ' + i + ' of ' + plan.updates.length + '…');
            return put(u, false);
          });
        }, Promise.resolve());
      })
      // 4. dependency links
      .then(function () {
        var keyOf = {};
        plan.features.forEach(function (f) { if (f.key) keyOf[f.id] = f.key; });
        plan.updates.forEach(function (u) { if (u.kind === 'feature') keyOf[u.id] = u.key; });
        // story links resolve against the stories created in step 3 and the
        // ones that already had a key (ids never collide across the two)
        plan.stories.forEach(function (s2) { if (s2.key) keyOf[s2.id] = s2.key; });
        plan.updates.forEach(function (u) { if (u.kind === 'story') keyOf[u.id] = u.key; });
        var links = plan.links.filter(function (l) { return keyOf[l.blockerId] && keyOf[l.blockedId]; });
        var i = 0;
        return links.reduce(function (p, l) {
          return p.then(function () {
            i += 1;
            progress('Linking ' + i + ' of ' + links.length + '…');
            return client.post('/rest/api/3/issueLink', {
              type: { name: 'Blocks' },
              outwardIssue: { key: keyOf[l.blockerId] },
              inwardIssue: { key: keyOf[l.blockedId] }
            }).then(function () { result.linked += 1; }, function (err) {
              // an already-present link is fine; anything else is reported
              if (!/already|exist/i.test(errText(err))) result.errors.push('Link ' + keyOf[l.blockerId] + ' → ' + keyOf[l.blockedId] + ': ' + errText(err));
            });
          });
        }, Promise.resolve());
      })
      // 4b. stories that sit beside their feature (no sub-task nesting) are
      //     tied to it with a Relates link
      .then(function () {
        var keyOf = {};
        plan.features.forEach(function (f) { if (f.key) keyOf['feature:' + f.id] = f.key; });
        plan.updates.forEach(function (u) { keyOf[u.kind + ':' + u.id] = u.key; });
        var todo = [];
        plan.stories.forEach(function (st) {
          if (!st.key) return;
          if (st.nest === false || st.unnested) todo.push({ story: st.key, feature: keyOf['feature:' + st.itemId], title: st.title });
        });
        todo = todo.filter(function (t) { return !!t.feature; });
        var i = 0;
        return todo.reduce(function (p, t) {
          return p.then(function () {
            i += 1;
            progress('Linking stories to features (' + i + ' of ' + todo.length + ')…');
            return client.post('/rest/api/3/issueLink', {
              type: { name: 'Relates' }, outwardIssue: { key: t.feature }, inwardIssue: { key: t.story }
            }).then(function () { result.storyLinked += 1; }, function (err) {
              if (!/already|exist/i.test(errText(err))) result.errors.push(t.story + ' ' + t.title + ': ' + errText(err));
            });
          });
        }, Promise.resolve());
      })
      // 5. done in Headway -> a done status in Jira (the first transition
      //    that leads to the Done category)
      .then(function () {
        var keyOf = {};
        plan.features.forEach(function (f) { if (f.key) keyOf['feature:' + f.id] = f.key; });
        plan.stories.forEach(function (s) { if (s.key) keyOf['story:' + s.id] = s.key; });
        var todo = plan.transitions.map(function (t) { return { key: t.key || keyOf[t.kind + ':' + t.id], title: t.title }; })
          .filter(function (t) { return !!t.key; });
        var i = 0;
        return todo.reduce(function (p, t) {
          return p.then(function () {
            i += 1;
            progress('Marking done in Jira (' + i + ' of ' + todo.length + ')…');
            return client.get('/rest/api/3/issue/' + encodeURIComponent(t.key) + '/transitions').then(function (res) {
              var list = res && Array.isArray(res.transitions) ? res.transitions : [];
              var tr = list.filter(function (x) { return x.to && x.to.statusCategory && x.to.statusCategory.key === 'done'; })[0];
              if (!tr) { result.errors.push(t.key + ' ' + t.title + ': no transition to a done status is available'); return; }
              return client.post('/rest/api/3/issue/' + encodeURIComponent(t.key) + '/transitions', { transition: { id: tr.id } })
                .then(function () { result.transitioned += 1; });
            }).then(null, function (err) { result.errors.push(t.key + ' ' + t.title + ': ' + errText(err)); });
          });
        }, Promise.resolve());
      })
      // 6. sprints: create the missing ones, then place issues
      .then(function () {
        if (!plan.sprints) return;
        var sp = plan.sprints;
        var idByNum = {};
        sp.assign.forEach(function (a) { if (a.sprintId) idByNum[a.num] = a.sprintId; });
        var i = 0;
        return sp.create.reduce(function (p, span) {
          return p.then(function () {
            i += 1;
            progress('Creating sprint ' + i + ' of ' + sp.create.length + '…');
            return client.post('/rest/agile/1.0/sprint', {
              name: span.name, originBoardId: sp.boardId,
              startDate: span.start + 'T00:00:00.000Z', endDate: span.end + 'T23:59:00.000Z'
            }).then(function (s) {
              if (s && s.id) { idByNum[span.num] = s.id; result.sprintsCreated += 1; result.sprintIds[span.num] = s.id; }
              else result.errors.push(span.name + ': sprint not created');
            }, function (err) { result.errors.push(span.name + ': ' + errText(err)); });
          });
        }, Promise.resolve()).then(function () {
          var keyOf = {};
          plan.features.forEach(function (f) { if (f.key) keyOf['feature:' + f.id] = f.key; });
          plan.stories.forEach(function (s) { if (s.key) keyOf['story:' + s.id] = s.key; });
          plan.updates.forEach(function (u) { keyOf[u.kind + ':' + u.id] = u.key; });
          var bySprint = {};
          sp.assign.forEach(function (a) {
            var id = idByNum[a.num], key = keyOf[a.ref.kind + ':' + a.ref.id];
            if (!id || !key) return;
            (bySprint[id] = bySprint[id] || []).push(key);
          });
          var ids = Object.keys(bySprint), j = 0;
          return ids.reduce(function (p, id) {
            return p.then(function () {
              j += 1;
              progress('Placing issues in sprints (' + j + ' of ' + ids.length + ')…');
              return chunks(bySprint[id], 50).reduce(function (q, keys) {
                return q.then(function () {
                  return client.post('/rest/agile/1.0/sprint/' + id + '/issue', { issues: keys })
                    .then(function () { result.sprintAssigned += keys.length; },
                      function (err) { result.errors.push('Sprint ' + id + ': ' + errText(err)); });
                });
              }, Promise.resolve());
            });
          }, Promise.resolve());
        });
      })
      .then(function () {
        result.droppedFields = Object.keys(dropped);
        if (result.parentRefused) result.errors.push(result.parentRefused + ' stor' + (result.parentRefused === 1 ? 'y' : 'ies') +
          ' could not nest under the feature (Jira refused the parent) and were created beside it, linked. Add the Subtask issue type to the project to nest stories.');
        return result;
      });
  };

  // write the sync result + pulls into a Headway state (used inside commit)
  JR.applyToState = function (s, plan, result, info) {
    Object.keys(result.epicKeys).forEach(function (e) { s.epicJira[e] = result.epicKeys[e]; });
    s.items.forEach(function (it) {
      if (result.itemKeys[it.id]) it.jiraKey = result.itemKeys[it.id];
      it.stories.forEach(function (st) { if (result.storyKeys[st.id]) st.jiraKey = result.storyKeys[st.id]; });
    });
    plan.pulls.forEach(function (pl) {
      var it = RM.itemById(s, pl.itemId || pl.id);
      if (!it) return;
      if (pl.kind === 'feature') it.done = pl.done;
      else it.stories.forEach(function (st) { if (st.id === pl.id) st.done = pl.done; });
    });
    if (info) {
      s.meta.jira = s.meta.jira || {};
      // remember who matched whom, so the next sync skips the lookups
      var acc = s.meta.jira.accounts || {}, names = s.meta.jira.accountNames || {};
      Object.keys(info.accounts || {}).forEach(function (l) {
        if (!info.accounts[l]) return;
        acc[l] = info.accounts[l];
        if (info.accountNames && info.accountNames[l]) names[l] = info.accountNames[l];
      });
      s.meta.jira.accounts = acc;
      s.meta.jira.accountNames = names;
    }
  };

  // ------------------------------------------------------------ status
  // what a sync would look at, compacted — equal fingerprints mean nothing
  // sync-relevant changed since the last run
  JR.fingerprint = function (state) {
    var parts = [];
    (state.items || []).forEach(function (it) {
      if (it.milestone) return;
      parts.push([it.num, it.jiraKey || '', it.feature, RM.htmlToText(it.description || ''), it.epic || '', it.workstream || '',
        it.size || '', it.priority || '', it.deadline || '', it.startDay, it.durDays, it.done ? 1 : 0, it.phaseId,
        (it.deps || []).join(','), (it.assignees || []).join(','), RM.htmlToText(it.enables || ''), RM.htmlToText(it.outOfScope || ''),
        RM.htmlToText(it.extDeps || ''), RM.htmlToText(it.notes || ''),
        (it.stories || []).map(function (st) {
          return [st.jiraKey || '', st.title, st.done ? 1 : 0, st.priority || '', st.deadline || '', st.startDay, st.durDays,
            (st.assignees || []).join(','), RM.htmlToText(st.description || ''), RM.htmlToText(st.ac || '')].join('');
        }).join('')].join(''));
    });
    var ej = state.epicJira || {};
    parts.push(Object.keys(ej).sort().map(function (k) { return k + '=' + ej[k]; }).join(','));
    var str = parts.join('\n');
    // FNV-1a keeps the stored value short
    var h = 0x811c9dc5;
    for (var i = 0; i < str.length; i++) { h ^= str.charCodeAt(i); h = Math.imul(h, 0x01000193) >>> 0; }
    return ('0000000' + h.toString(16)).slice(-8) + ':' + str.length;
  };
  JR.job = null;        // the running background sync: { progress, title }
  JR.lastResult = null; // { plan, result, at } of the last finished sync (this session)
  // { kind: 'off' | 'busy' | 'never' | 'dirty' | 'clean' | 'error', label, linked, total, lastSync, last }
  JR.status = function (state) {
    var j = (state && state.meta && state.meta.jira) || {};
    var items = ((state && state.items) || []).filter(function (it) { return !it.milestone; });
    var linked = items.filter(function (it) { return !!it.jiraKey; }).length;
    var out = { linked: linked, total: items.length, lastSync: j.lastSync || null, project: j.project || '', last: '' };
    var lc2 = j.lastCounts;
    if (lc2) {
      var bits = [];
      if (lc2.created) bits.push(lc2.created + ' created');
      if (lc2.updated) bits.push(lc2.updated + ' updated');
      if (lc2.linked) bits.push(lc2.linked + ' linked');
      if (lc2.sprints) bits.push(lc2.sprints + ' placed in sprints');
      if (lc2.done) bits.push(lc2.done + ' marked done in Jira');
      if (lc2.pulled) bits.push(lc2.pulled + ' done flag' + (lc2.pulled === 1 ? '' : 's') + ' read back');
      out.last = bits.length ? bits.join(', ') : 'nothing changed';
    }
    if (JR.job) { out.kind = 'busy'; out.label = 'Syncing… ' + (JR.job.progress || ''); return out; }
    if (!JR.hasCreds() || !j.project) { out.kind = 'off'; out.label = 'Jira sync not set up'; return out; }
    if (!j.lastSync) { out.kind = 'never'; out.label = 'Not synced yet'; return out; }
    out.auto = j.auto !== false;
    if (j.lastErrors) { out.kind = 'error'; out.label = 'Last sync had ' + j.lastErrors + ' problem' + (j.lastErrors === 1 ? '' : 's'); return out; }
    if (j.syncedHash && JR.fingerprint(state) === j.syncedHash) { out.kind = 'clean'; out.label = 'In sync with ' + j.project; return out; }
    out.kind = 'dirty'; out.label = 'Changes since the last sync';
    return out;
  };
  function ago(iso) {
    var t = Date.parse(iso);
    if (!t) return '';
    var m = Math.round((Date.now() - t) / 60000);
    if (m < 1) return 'just now';
    if (m < 60) return m + ' min ago';
    var h = Math.round(m / 60);
    if (h < 48) return h + ' h ago';
    return Math.round(h / 24) + ' days ago';
  }

  // ------------------------------------------------------------ browser UI
  function esc(s) {
    return String(s == null ? '' : s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
  }
  JR.loadCreds = function () {
    try { return JSON.parse(root.localStorage.getItem(JR.LOCAL_KEY) || '{}') || {}; } catch (e) { return {}; }
  };
  JR.saveCreds = function (c) {
    try { root.localStorage.setItem(JR.LOCAL_KEY, JSON.stringify(c)); } catch (e) { /* storage blocked */ }
  };
  JR.hasCreds = function () {
    var c = JR.loadCreds();
    return !!(c.email && c.token);
  };
  // login from this machine + site from the document (older setups kept the
  // site with the login; that still works as a fallback)
  JR.credsFor = function (state) {
    var c = JR.loadCreds();
    var site = (state && state.meta && state.meta.jira && state.meta.jira.site) || c.site || '';
    return { site: site, email: c.email || '', token: c.token || '' };
  };
  function app() { return root.HeadwayApp; }
  function docCfg() { return cfgOf(app().ai.state(), null); }

  JR.renderStatus = function (btn, state) {
    if (!btn) return;
    var st = JR.status(state);
    btn.className = st.kind;
    btn.title = st.label + (st.lastSync && st.kind !== 'busy' ? ' · synced ' + ago(st.lastSync) + (st.last ? ': ' + st.last : '') : '') +
      (st.auto && st.kind !== 'busy' && st.kind !== 'off' ? ' · auto-sync every 5 min' : '');
  };
  // re-read the button from the live document (progress ticks, job end)
  JR.refreshStatus = function () {
    var btn = root.document && root.document.getElementById('btnJira');
    var A = app();
    if (btn && A && A.ai.hasDoc()) JR.renderStatus(btn, A.ai.state());
  };
  // Runs a plan in the background: the topbar button spins with the
  // progress, a toast reports the outcome, and the keys / done flags land
  // in the document through one commit. Nothing here needs a dialog open.
  // opts: { quiet } — an automatic run only speaks up about problems;
  // { resume } — the caller already opened JR.job (auto-sync discovery)
  JR.runSync = function (plan, client, info, opts) {
    var A = app();
    opts = opts || {};
    if (JR.job && !opts.resume) { A.toast('A Jira sync is already running'); return Promise.resolve(null); }
    var title = A.ai.state().meta.title;
    JR.job = { progress: 'Starting…', title: title, startedAt: Date.now(), auto: !!opts.quiet };
    JR.refreshStatus();
    return JR.apply(plan, client, function (msg) { JR.job.progress = msg; JR.refreshStatus(); }).then(function (result) {
      JR.job = null;
      if (result.droppedFields.length) {
        result.errors.push('Jira refused ' + result.droppedFields.map(function (f) {
          return f === (info && info.startField) ? 'Start date (' + f + ')' : f;
        }).join(', ') + ' — add the field to the project’s screens so it can be set');
      }
      JR.lastResult = { plan: plan, result: result, at: new Date().toISOString() };
      // the sync belongs to the document it started on; a project switched
      // mid-sync keeps its own keys untouched
      if (!A.ai.hasDoc() || A.ai.state().meta.title !== title) {
        A.toast('Jira sync finished, but a different project is open — keys were not recorded', 'warn');
        JR.refreshStatus();
        return result;
      }
      A.ai.commit('jira sync', function (s) {
        JR.applyToState(s, plan, result, info);
        s.meta.jira = s.meta.jira || {};
        s.meta.jira.lastSync = new Date().toISOString();
        s.meta.jira.lastErrors = result.errors.length;
        s.meta.jira.lastCounts = { created: result.created, updated: result.updated, linked: result.linked,
          sprints: result.sprintAssigned, done: result.transitioned, pulled: plan.pulls.length };
        s.meta.jira.syncedHash = JR.fingerprint(s);
      });
      if (!opts.quiet || result.errors.length) {
        A.toast('Jira sync: ' + result.created + ' created, ' + result.updated + ' updated' +
          (result.linked ? ', ' + result.linked + ' linked' : '') +
          (result.sprintAssigned ? ', ' + result.sprintAssigned + ' placed in sprints' : '') +
          (result.transitioned ? ', ' + result.transitioned + ' marked done in Jira' : '') +
          (plan.pulls.length ? ', ' + plan.pulls.length + ' done flag' + (plan.pulls.length === 1 ? '' : 's') + ' read back' : '') +
          (result.errors.length ? ' · ' + result.errors.length + ' problem' + (result.errors.length === 1 ? '' : 's') : ''),
          result.errors.length ? 'warn' : '');
      }
      JR.refreshStatus();
      return result;
    }, function (err) {
      JR.job = null;
      if (opts.quiet) JR.autoFailedHash = opts.fingerprint || null;
      if (root.console) console.error('Jira sync failed', err);
      JR.lastResult = { plan: plan, result: null, error: errText(err), at: new Date().toISOString() };
      A.toast('Jira sync failed: ' + errText(err), 'err');
      JR.refreshStatus();
      return null;
    });
  };
  // ---- auto-sync: every AUTO_MS, a document with at least one item already
  // linked to Jira (a sync got that far once) and changed since the last run
  // is synced again without a dialog. A failed automatic run waits for the
  // next change rather than retrying every tick.
  JR.autoTimer = null;
  JR.autoFailedHash = null;
  JR.autoEnabled = function (state) {
    var j = (state && state.meta && state.meta.jira) || {};
    return j.auto !== false;
  };
  // the fingerprint an automatic run should sync to, or null when nothing is due
  // anything in the document already linked to Jira: a feature, story or epic
  JR.linkedOnce = function (state) {
    if (!state) return false;
    var ej = state.epicJira || {};
    if (Object.keys(ej).some(function (k) { return !!ej[k]; })) return true;
    return (state.items || []).some(function (it) {
      return !!it.jiraKey || (it.stories || []).some(function (st) { return !!st.jiraKey; });
    });
  };
  JR.autoDue = function (state) {
    if (!state || JR.job) return null;
    var j = (state.meta && state.meta.jira) || {};
    if (!JR.autoEnabled(state) || !JR.hasCreds() || !j.project) return null;
    if (!j.lastSync && !JR.linkedOnce(state)) return null;
    var fp = JR.fingerprint(state);
    if (fp === j.syncedHash || fp === JR.autoFailedHash) return null;
    return fp;
  };
  JR.autoTick = function () {
    var A = app();
    if (!A || !A.ai || !A.ai.hasDoc()) return Promise.resolve(false);
    var state = A.ai.state();
    var fp = JR.autoDue(state);
    if (!fp) return Promise.resolve(false);
    var creds = JR.credsFor(state);
    var cfg = cfgOf(state, creds);
    if (!creds.site) return Promise.resolve(false);
    var client = JR.client(creds);
    JR.job = { progress: 'Looking at Jira…', title: state.meta.title, startedAt: Date.now(), auto: true };
    JR.refreshStatus();
    return JR.discover(client, state, cfg, function (msg) { if (JR.job) { JR.job.progress = msg; JR.refreshStatus(); } }).then(function (info) {
      var plan = JR.plan(state, cfg, info);
      var c = plan.counts;
      if (!c.create && !c.update && !c.pull && !c.link && !c.sprintAssign && !c.done) {
        // nothing Jira needs — remember this shape so the tick stays quiet
        JR.job = null;
        A.ai.commit('jira sync', function (s) { s.meta.jira = s.meta.jira || {}; s.meta.jira.syncedHash = JR.fingerprint(s); });
        JR.refreshStatus();
        return true;
      }
      return JR.runSync(plan, client, info, { quiet: true, resume: true, fingerprint: fp }).then(function () { return true; });
    }, function (err) {
      JR.job = null;
      JR.autoFailedHash = fp;
      if (root.console) console.error('Jira auto-sync failed', err);
      A.toast('Jira auto-sync failed: ' + errText(err), 'err');
      JR.refreshStatus();
      return false;
    });
  };
  // arm the timer once per page; the tick itself decides whether to run
  JR.armAuto = function () {
    if (JR.autoTimer || !root.document || typeof root.setInterval !== 'function') return;
    JR.autoTimer = root.setInterval(function () { JR.autoTick(); }, JR.AUTO_MS);
    root.setTimeout(function () { JR.autoTick(); }, 20000); // shortly after the app opens
  };

  function line(n, what) { return n ? '<li><b>' + n + '</b> ' + what + '</li>' : ''; }
  function list(title, rows) {
    if (!rows.length) return '';
    return '<div class="m-sec"><label>' + title + '</label><div class="jr-list">' + rows.map(function (r) { return '<div title="' + esc(r) + '">' + esc(r) + '</div>'; }).join('') + '</div></div>';
  }
  // the last run's outcome, for the status menu
  JR.resultModal = function () {
    var A = app();
    var lr = JR.lastResult;
    if (!lr) return;
    var body;
    if (lr.error) body = '<div class="m-hint">Sync failed: ' + esc(lr.error) + '</div>';
    else {
      var r = lr.result;
      body = '<ul class="jr-plan">' +
        line(r.created, 'issue' + (r.created === 1 ? '' : 's') + ' created') +
        line(r.updated, 'issue' + (r.updated === 1 ? '' : 's') + ' updated') +
        line(r.linked, 'link' + (r.linked === 1 ? '' : 's') + ' added') +
        line(r.sprintsCreated, 'sprint' + (r.sprintsCreated === 1 ? '' : 's') + ' created') +
        line(r.sprintAssigned, 'issue' + (r.sprintAssigned === 1 ? '' : 's') + ' placed in sprints') +
        line(r.transitioned, 'issue' + (r.transitioned === 1 ? '' : 's') + ' marked done in Jira') +
        line(lr.plan.pulls.length, 'Done flag' + (lr.plan.pulls.length === 1 ? '' : 's') + ' read back') +
        '</ul>' + list('Problems', r.errors);
      if (!r.created && !r.updated && !r.linked && !r.sprintAssigned && !r.transitioned && !lr.plan.pulls.length && !r.errors.length) body = '<div class="m-hint">Nothing changed.</div>';
    }
    A.openModal('<div class="modal" style="width:520px">' +
      '<div class="m-head"><h2>Last Jira sync</h2><button class="p-close" data-m="x"><i data-lucide="x"></i></button></div>' +
      '<div class="m-body">' + body + '<div class="m-hint">' + esc(ago(lr.at)) + '</div></div>' +
      '<div class="m-foot"><button data-m="x2" class="primary">Close</button></div></div>', function (host) {
      host.querySelector('[data-m=x]').onclick = A.closeModal;
      host.querySelector('[data-m=x2]').onclick = A.closeModal;
    });
  };
  JR.statusMenu = function (anchor) {
    var A = app();
    if (!A.ai.hasDoc()) { A.toast('Open a project first'); return; }
    var st = JR.status(A.ai.state());
    var detail = st.kind === 'off' || st.kind === 'busy' ? st.label
      : st.label + ' · ' + st.linked + ' of ' + st.total + ' features linked';
    var lastLine = st.lastSync && st.kind !== 'busy' ? 'Last sync ' + ago(st.lastSync) + (st.last ? ': ' + st.last : '') : '';
    A.openDropdown(anchor, [
      { label: esc(detail), disabled: true },
      lastLine ? { label: esc(lastLine), disabled: true } : null,
      { sep: true },
      { icon: 'refresh-cw', label: 'Sync now…', disabled: !!JR.job, fn: function () { JR.syncModal(); } },
      { icon: 'timer-reset', label: 'Auto-sync every 5 min', checked: JR.autoEnabled(A.ai.state()), fn: function () {
        var on = !JR.autoEnabled(A.ai.state());
        A.ai.commit('jira settings', function (s) { s.meta.jira = s.meta.jira || {}; s.meta.jira.auto = on; });
        A.toast('Jira auto-sync ' + (on ? 'on' : 'off'));
      } },
      JR.lastResult ? { icon: 'list-checks', label: 'Last sync result…', fn: JR.resultModal } : null,
      { icon: 'settings-2', label: 'Jira settings…', fn: function () { A.openSetup('jira'); } }
    ].filter(Boolean));
  };

  JR.settingsHtml = function () {
    var st = app().ai.state();
    var c = JR.loadCreds();
    var d = docCfg();
    var desktop = !!root.__TAURI__;
    function inp(id, val, ph, type) {
      return '<input id="' + id + '" type="' + (type || 'text') + '" style="width:100%" value="' + esc(val || '') + '" placeholder="' + esc(ph || '') + '" autocomplete="off">';
    }
    function ck(id, on, label) {
      return '<label class="p-check" style="margin-top:7px"><input type="checkbox" id="' + id + '"' + (on ? ' checked' : '') + '> ' + label + '</label>';
    }
    return '<h2>Your login (this machine only)</h2>' +
      '<div class="p-grid2">' +
      '<div class="m-sec"><label>Email</label>' + inp('jrEmail', c.email, 'you@example.com') + '</div>' +
      '<div class="m-sec"><label>API token</label>' + inp('jrToken', c.token, '', 'password') + '</div></div>' +
      '<div class="m-hint">Kept in this app’s local storage on this computer and never written to the roadmap file, so nobody who opens the file sees it. Create a token at id.atlassian.com → Security → API tokens.' +
      (desktop ? '' : ' Jira Cloud blocks browser calls; syncing works from the desktop app.') + '</div>' +
      '<div class="p-row" style="margin-top:8px"><button id="jrTest">Test connection</button><span id="jrTestOut" class="m-hint" style="margin:0 0 0 10px"></span></div>' +
      '<h2 style="margin-top:22px">Project (shared in the file)</h2>' +
      '<div class="m-sec"><label>Site</label>' + inp('jrSite', d.site || c.site, 'your-team.atlassian.net') + '</div>' +
      '<div class="p-grid2">' +
      '<div class="m-sec"><label>Project key</label>' + inp('jrProject', d.project, 'e.g. HW') + '</div>' +
      '<div class="m-sec"><label>Start date field id</label>' + inp('jrStartField', d.startField, 'found automatically') + '</div>' +
      '</div>' +
      ck('jrPushEpics', d.pushEpics, 'Create an epic per Headway epic and parent features to it') +
      ck('jrPushStories', d.pushStories, 'Sync stories as their own issues under the feature') +
      ck('jrSprints', d.sprints, 'Place issues in the board’s sprints by start date, creating missing sprints') +
      ck('jrAuto', d.auto, 'Sync automatically every 5 minutes once anything is linked to Jira') +
      '<h2 style="margin-top:18px">Issue types</h2>' +
      '<div class="m-hint">Each Headway type becomes this Jira issue type. Types are defined in Setup → Hierarchy; the Jira name can be edited here or there.</div>' +
      '<table class="hol-table jr-types"><thead><tr><th>Headway type</th><th>Jira issue type</th><th>In project</th></tr></thead><tbody>' +
      RM.itemTypes(st).map(function (t) {
        var r = JR.lastTypes && JR.lastTypes.byKey && JR.lastTypes.byKey[t.key];
        var res = !r ? '' : (lc(r.name) === lc(t.jira || t.label) ? '✓' : 'falls back to ' + esc(r.name));
        return '<tr><td><i data-lucide="' + esc(t.icon) + '"></i> ' + esc(t.label) + '</td>' +
          '<td><input data-jrtype="' + esc(t.key) + '" value="' + esc(t.jira || '') + '" placeholder="' + esc(t.label) + '"></td>' +
          '<td class="m-hint">' + res + '</td></tr>';
      }).join('') + '</tbody></table>' +
      '<div class="m-hint">Site, project and issue types travel with the file, so teammates only add their own login. ' +
      'Every synced issue gets a Start date and a Due date from the timeline; milestones are never sent. ' +
      'Types that do not exist in the project fall back to what it has. Jira keys land back on features, stories and epics after a sync; ' +
      'the Done state of linked issues is read back from Jira.</div>' +
      '<div class="p-row" style="margin-top:10px"><button id="jrSync" class="primary">Sync with Jira…</button></div>' +
      JR.peopleHtml();
  };
  // Headway people -> Jira accounts (the Assignee field). Automatic matches
  // show as such; a search box per person fixes the rest by hand.
  JR.peopleHtml = function () {
    var st = app().ai.state();
    var j = st.meta.jira || {};
    var acc = j.accounts || {}, names = j.accountNames || {};
    var team = (st.team || []).map(function (m) { return RM.memberLabel(m); }).filter(function (l, i, a) { return a.indexOf(l) === i; });
    if (!team.length) return '';
    return '<h2 style="margin-top:22px">People → Jira accounts</h2>' +
      '<div class="m-hint">Each person on the Team becomes the Jira assignee through this mapping. Names are matched automatically during a sync; search here for anyone that did not match.</div>' +
      '<div class="jr-people">' + team.map(function (l) {
        var id = acc[l];
        return '<div class="jr-person" data-jr-person="' + esc(l) + '">' +
          '<span class="jr-pname">' + esc(l) + '</span>' +
          '<span class="jr-pstate' + (id ? ' on' : '') + '">' + (id ? esc(names[l] || id) : 'not matched') + '</span>' +
          '<input class="jr-psearch" placeholder="Search Jira users…" value="" autocomplete="off">' +
          '<button data-jr-find>Find</button>' +
          (id ? '<button data-jr-clear title="Clear">×</button>' : '') +
          '<span class="jr-presults"></span></div>';
      }).join('') + '</div>';
  };
  JR.wireSettings = function (host) {
    function $(sel) { return host.querySelector(sel); }
    function creds() {
      return { email: $('#jrEmail').value.trim(), token: $('#jrToken').value.trim() };
    }
    ['#jrEmail', '#jrToken'].forEach(function (sel) {
      $(sel).addEventListener('change', function () { JR.saveCreds(creds()); });
    });
    function mapping() {
      return {
        site: JR.siteUrl($('#jrSite').value),
        project: $('#jrProject').value.trim().toUpperCase(),
        startField: $('#jrStartField').value.trim(),
        pushEpics: $('#jrPushEpics').checked,
        pushStories: $('#jrPushStories').checked,
        sprints: $('#jrSprints').checked,
        auto: $('#jrAuto').checked
      };
    }
    ['#jrSite', '#jrProject', '#jrStartField', '#jrPushEpics', '#jrPushStories', '#jrSprints', '#jrAuto'].forEach(function (sel) {
      $(sel).addEventListener('change', function () {
        var m = mapping();
        app().ai.commit('jira settings', function (s) {
          var cur = s.meta.jira || {};
          Object.keys(m).forEach(function (k) { cur[k] = m[k]; });
          s.meta.jira = cur;
        });
      });
    });
    host.addEventListener('change', function (ev) {
      var inp = ev.target.closest && ev.target.closest('[data-jrtype]');
      if (!inp) return;
      var key = inp.dataset.jrtype, val = inp.value;
      app().ai.commit('jira issue type', function (s) { RM.setItemTypeJira(s, key, val); });
    });
    $('#jrTest').addEventListener('click', function () {
      var out = $('#jrTestOut');
      var c = creds();
      JR.saveCreds(c);
      var full = { site: JR.siteUrl($('#jrSite').value), email: c.email, token: c.token };
      if (!full.site || !full.email || !full.token) { out.textContent = 'Fill in site, email and token first'; return; }
      out.textContent = 'Connecting…';
      JR.testConnection(full, $('#jrProject').value.trim().toUpperCase()).then(function (r) {
        out.textContent = 'Connected as ' + r.user + (r.project ? ' · project “' + r.project + '”' : '');
        if (r.issueTypes) {
          JR.lastTypes = JR.resolveTypes(r.issueTypes, app().ai.state());
          app().openSetup('jira');
        }
      }, function (err) { out.textContent = 'Failed: ' + errText(err); });
    });
    $('#jrSync').addEventListener('click', function () { JR.syncModal(); });
    // people picker
    host.addEventListener('click', function (e) {
      var row = e.target.closest('[data-jr-person]');
      if (!row) return;
      var label = row.dataset.jrPerson;
      if (e.target.closest('[data-jr-clear]')) {
        app().ai.commit('jira settings', function (s) {
          s.meta.jira = s.meta.jira || {};
          if (s.meta.jira.accounts) delete s.meta.jira.accounts[label];
          if (s.meta.jira.accountNames) delete s.meta.jira.accountNames[label];
        });
        return;
      }
      if (e.target.closest('[data-jr-find]')) {
        var out = row.querySelector('.jr-presults');
        var st = app().ai.state();
        var creds = JR.credsFor(st);
        var project = (st.meta.jira && st.meta.jira.project) || '';
        if (!creds.site || !creds.email || !creds.token || !project) { out.textContent = 'Set the login, site and project key first'; return; }
        var q = row.querySelector('.jr-psearch').value.trim() || label;
        out.textContent = 'Searching…';
        JR.client(creds).get('/rest/api/3/user/assignable/search?project=' + encodeURIComponent(project) + '&query=' + encodeURIComponent(q) + '&maxResults=20')
          .then(function (users) {
            var list = Array.isArray(users) ? users : [];
            if (!list.length) { out.textContent = 'No assignable user matches “' + q + '”'; return; }
            out.innerHTML = '<select data-jr-pick><option value="">Pick…</option>' + list.map(function (u) {
              return '<option value="' + esc(u.accountId) + '" data-name="' + esc(u.displayName || '') + '">' + esc(u.displayName || u.accountId) + (u.emailAddress ? ' · ' + esc(u.emailAddress) : '') + '</option>';
            }).join('') + '</select>';
          }, function (err) { out.textContent = 'Failed: ' + errText(err); });
      }
    });
    host.addEventListener('change', function (e) {
      var sel = e.target.closest('[data-jr-pick]');
      if (!sel || !sel.value) return;
      var row = sel.closest('[data-jr-person]');
      var label = row.dataset.jrPerson, id = sel.value, name = sel.options[sel.selectedIndex].dataset.name || '';
      app().ai.commit('jira settings', function (s) {
        s.meta.jira = s.meta.jira || {};
        s.meta.jira.accounts = s.meta.jira.accounts || {};
        s.meta.jira.accountNames = s.meta.jira.accountNames || {};
        s.meta.jira.accounts[label] = id;
        s.meta.jira.accountNames[label] = name;
      });
    });
  };

  // The Sync dialog: look at Jira, show the plan, hand it to the background
  // job on Apply
  JR.syncModal = function () {
    var A = app();
    if (!A.ai.hasDoc()) { A.toast('Open a project first'); return; }
    var creds = JR.credsFor(A.ai.state());
    var cfg = cfgOf(A.ai.state(), creds);
    function shell(body, foot) {
      return '<div class="modal" style="width:540px">' +
        '<div class="m-head"><h2>Sync with Jira</h2><button class="p-close" data-m="x"><i data-lucide="x"></i></button></div>' +
        '<div class="m-body">' + body + '</div>' +
        '<div class="m-foot">' + foot + '</div></div>';
    }
    if (!creds.site || !creds.email || !creds.token || !cfg.project) {
      A.openModal(shell('<div class="m-hint">Add your Jira login, the site and a project key first.</div>',
        '<button data-m="x2">Close</button><button data-m="setup" class="primary">Open Jira settings</button>'), function (host) {
        host.querySelector('[data-m=x]').onclick = A.closeModal;
        host.querySelector('[data-m=x2]').onclick = A.closeModal;
        host.querySelector('[data-m=setup]').onclick = function () { A.closeModal(); A.openSetup('jira'); };
      });
      return;
    }
    if (JR.job) { A.toast('A Jira sync is already running'); return; }
    var client = JR.client(creds);
    var state = A.ai.state();
    var plan = null, info = null;
    A.openModal(shell('<div id="jrBody" class="m-hint">Looking at Jira…</div>', '<button data-m="x2">Cancel</button><button data-m="go" class="primary" disabled>Apply</button>'), function (host) {
      var body = host.querySelector('#jrBody');
      var go = host.querySelector('[data-m=go]');
      var cancel = host.querySelector('[data-m=x2]');
      host.querySelector('[data-m=x]').onclick = A.closeModal;
      cancel.onclick = A.closeModal;
      JR.discover(client, state, cfg, function (msg) { body.textContent = msg; }).then(function (inf) {
        info = inf;
        plan = JR.plan(state, cfg, info);
        var c = plan.counts;
        if (!c.create && !c.update && !c.pull && !c.link && !c.sprintAssign && !c.done) { body.innerHTML = '<div class="m-hint">Nothing to sync.</div>' + list('Notes', plan.notes); return; }
        body.innerHTML = '<ul class="jr-plan">' +
          line(plan.epics.length, esc(RM.levelLabel(state, 'epic', plan.epics.length !== 1).toLowerCase()) + ' to create in ' + esc(cfg.project)) +
          line(plan.features.length, esc(RM.levelLabel(state, 'feature', plan.features.length !== 1).toLowerCase()) + ' to create') +
          line(plan.stories.length, esc(RM.levelLabel(state, 'story', plan.stories.length !== 1).toLowerCase()) + ' to create') +
          line(c.update, 'linked issue' + (c.update === 1 ? '' : 's') + ' to update') +
          line(c.link, 'dependency link' + (c.link === 1 ? '' : 's') + ' to add') +
          line(c.sprintsCreate, 'sprint' + (c.sprintsCreate === 1 ? '' : 's') + ' to create on ' + esc(plan.sprints ? plan.sprints.boardName : '')) +
          line(c.sprintAssign, 'issue' + (c.sprintAssign === 1 ? '' : 's') + ' to place in sprints') +
          line(c.done, 'issue' + (c.done === 1 ? '' : 's') + ' to mark done in Jira') +
          line(c.pull, 'Done flag' + (c.pull === 1 ? '' : 's') + ' to read back from Jira') +
          '</ul>' +
          '<div class="m-hint">' +
          (info.startField ? 'Start date → ' + esc(info.startFieldName || info.startField) + ', end → Due date; epics span their features. ' : 'Only Due dates (no Start date field). ') +
          (plan.people.total ? (plan.people.total - plan.people.unresolved.length) + ' of ' + plan.people.total + ' people matched to Jira accounts. ' : '') +
          (plan.milestones.length ? plan.milestones.length + ' milestone' + (plan.milestones.length === 1 ? '' : 's') + ' skipped. ' : '') +
          'Apply runs in the background; the sync button in the toolbar shows progress and the result.</div>' +
          list('People not found in Jira (left unassigned)', plan.people.unresolved) +
          list('Notes', plan.notes) +
          list('Done state from Jira', plan.pulls.map(function (p) { return p.key + ' ' + p.title + ' → ' + (p.done ? 'done' : 'not done') + ' (' + p.status + ')'; })) +
          list('Keys not found in Jira (left alone)', plan.missing.map(function (m) { return m.key + ' ' + m.title; })) +
          list('Milestones already in Jira (never updated — delete them there if unwanted)', plan.milestones.filter(function (m) { return m.key; }).map(function (m) { return m.key + ' ' + m.title; }));
        go.disabled = false;
      }, function (err) {
        if (root.console) console.error('Jira discovery failed', err);
        body.innerHTML = '<div class="m-hint">Could not reach Jira: ' + esc(errText(err)) + '</div>';
      });
      go.onclick = function () {
        if (!plan) return;
        A.closeModal();
        JR.runSync(plan, client, info);
      };
    });
  };

  root.HeadwayJira = JR;
  JR.armAuto();
  if (typeof module !== 'undefined' && module.exports) module.exports = JR;
})(typeof window !== 'undefined' ? window : globalThis);
