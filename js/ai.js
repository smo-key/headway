/*
 * Headway AI assistant.
 *
 * Pure, node-testable pieces:
 *  - applyOps(state, ops) -> { state, changes }: path-addressed edits to any
 *    part of a document ('items/#12/feature', 'phases/@p1/name', …), then
 *    RM.normalizeState.
 *  - runTool(name, args, A): the tool set both providers share, executed
 *    against the app bridge (window.HeadwayApp.ai). Reads are clones; every
 *    write goes through A.ai.commit, so it is undoable and lands in Version
 *    history as "<name> · AI".
 *  - openai.*: OpenAI-compatible (LiteLLM) request building, SSE parsing and
 *    delta accumulation (reasoning_content / thinking_blocks / tool_calls).
 *  - claude.*: the `claude -p` stream-json protocol — user lines, event
 *    reduction, and the ```headway-tool fenced text protocol that stands in
 *    for function calling.
 *  - md(text): light, escaped markdown for assistant replies.
 *
 * Browser glue (window.HeadwayAI): the Setup → Personal → AI assistant tab
 * (settings live on this machine only) and the chat drawer (#aiDrawer).
 */
(function (root) {
  'use strict';

  var RM = root.RM || (typeof require !== 'undefined' ? require('./core.js') : null);
  var AI = {};

  AI.LOCAL_KEY = 'headway-ai-v1';      // provider settings — per machine
  AI.CHAT_KEY = 'headway-ai-chat-v1';  // last conversation (file bytes dropped)
  AI.UI_KEY = 'headway-ai-ui-v1';      // drawer open / width
  AI.EFFORTS = [['low', 'Low'], ['medium', 'Medium'], ['high', 'High'], ['max', 'Max']];
  // gateways speak OpenAI's reasoning_effort, which knows three levels
  AI.GATEWAY_EFFORTS = AI.EFFORTS.slice(0, 3);
  // per-model facts learnt from the gateway (`/model_group/info`):
  // { '<model id>': { reasoning: bool } }; null until fetched
  AI.modelInfo = null;
  AI.modelInfoBase = '';
  // the gateway's model list, fetched once per gateway URL (drawer open or
  // Setup → Load); null until fetched
  AI.modelCache = null;
  AI.modelCacheBase = '';
  // 'bedrock/global.us.claude-opus-5' -> 'claude-opus-5' for labels only:
  // the last path segment, minus leading provider / region / vendor tokens
  AI.shortModel = function (id) {
    var s = String(id || '');
    s = s.slice(s.lastIndexOf('/') + 1);
    var pre = /^(global|us|eu|ap|apac|anthropic|amazon|meta|mistral|cohere|ai21|google|openai|azure|bedrock|vertex_ai|vertex)\./i;
    while (pre.test(s)) s = s.replace(pre, '');
    return s;
  };
  // effort levels a provider/model accepts: the CLI takes all four; a gateway
  // model offers OpenAI's three when it reasons, none when it doesn't, and
  // every level while the gateway has not said
  AI.effortsFor = function (s) {
    if (!s || s.provider === 'claude') return AI.EFFORTS.slice();
    var info = AI.modelInfo && AI.modelInfo[s.model];
    if (!info) return AI.EFFORTS.slice();
    return info.reasoning ? AI.GATEWAY_EFFORTS.slice() : [];
  };
  AI.effortAllowed = function (s) {
    return !!(s && s.effort) && AI.effortsFor(s).some(function (e) { return e[0] === s.effort; });
  };
  // the effort to run with after a model change: the current pick when the
  // model offers it, else Medium, else the first level it does offer; a
  // model with no effort support keeps the stored value (the selector hides)
  AI.pickEffort = function (s) {
    var eff = AI.effortsFor(s);
    if (!eff.length || eff.some(function (e) { return e[0] === s.effort; })) return s.effort;
    return eff.some(function (e) { return e[0] === 'medium'; }) ? 'medium' : eff[0][0];
  };
  AI.CLAUDE_MODELS = [['sonnet', 'Sonnet'], ['opus', 'Opus'], ['fable', 'Fable'], ['haiku', 'Haiku']];
  AI.DEFAULTS = {
    provider: 'litellm', baseUrl: '', apiKey: '', model: '', headers: '',
    claudeBin: '', claudeModel: 'sonnet', effort: 'medium'
  };
  AI.MAX_ROUNDS = 12;         // tool-call rounds per user message
  AI.FENCE = 'headway-tool';  // claude -p tool-call fence language
  AI.TEXT_FILE_MAX = 300 * 1024;
  AI.BIN_FILE_MAX = 12 * 1024 * 1024;

  // ------------------------------------------------------------ helpers
  function esc(s) {
    return String(s == null ? '' : s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
  }
  AI.esc = esc;
  function clone(o) { return JSON.parse(JSON.stringify(o)); }
  function isObj(v) { return v && typeof v === 'object' && !Array.isArray(v); }
  function isoOfDay(meta, day) {
    if (day == null) return null;
    var d = RM.dayToDate(meta, day);
    return d ? RM.fmtISO(d) : null;
  }
  function dayOfIso(meta, iso) {
    if (!iso || !/^\d{4}-\d{2}-\d{2}$/.test(String(iso))) return null;
    var d = RM.dateToDay(meta, RM.parseISO(iso));
    if (d == null || d < 0 || d >= RM.numDays(meta)) return null;
    return d;
  }
  AI.isoOfDay = isoOfDay;
  AI.dayOfIso = dayOfIso;
  function spanEndIso(meta, start, dur) {
    if (start == null || dur == null) return null;
    if (dur <= 0) return isoOfDay(meta, start);
    return RM.fmtISO(RM.spanEndDate(meta, start, dur));
  }

  // ------------------------------------------------------------ settings
  AI.loadSettings = function () {
    var out = clone(AI.DEFAULTS);
    try {
      var raw = JSON.parse(root.localStorage.getItem(AI.LOCAL_KEY) || '{}') || {};
      Object.keys(out).forEach(function (k) { if (raw[k] != null) out[k] = raw[k]; });
    } catch (e) { /* storage blocked */ }
    if (out.provider !== 'claude') out.provider = 'litellm';
    if (!AI.EFFORTS.some(function (e) { return e[0] === out.effort; })) out.effort = 'medium';
    return out;
  };
  AI.saveSettings = function (s) {
    try { root.localStorage.setItem(AI.LOCAL_KEY, JSON.stringify(s)); } catch (e) { /* storage blocked */ }
  };
  // { ok, why } — can a turn be sent with these settings?
  AI.ready = function (s, desktop) {
    if (s.provider === 'claude') {
      if (!desktop) return { ok: false, why: 'The Claude subscription provider runs the Claude Code CLI, which only the desktop app can start. Pick the LiteLLM gateway here, or use the desktop app.' };
      return { ok: true };
    }
    if (!String(s.baseUrl || '').trim()) return { ok: false, why: 'Set the gateway URL in AI settings.' };
    if (!String(s.apiKey || '').trim()) return { ok: false, why: 'Set the API key in AI settings.' };
    if (!String(s.model || '').trim()) return { ok: false, why: 'Pick a model in AI settings.' };
    return { ok: true };
  };
  // 'https://host/litellm/v1/' -> 'https://host/litellm'
  AI.baseOf = function (url) {
    var s = String(url || '').trim().replace(/\/+$/, '');
    if (!s) return '';
    if (!/^https?:\/\//i.test(s)) s = 'https://' + s;
    s = s.replace(/\/v1$/i, '').replace(/\/chat\/completions$/i, '').replace(/\/v1$/i, '');
    return s;
  };
  // "Name: value" lines -> {}
  AI.parseHeaders = function (text) {
    var out = {};
    String(text || '').split(/\r?\n/).forEach(function (line) {
      var m = line.match(/^\s*([^:\s]+)\s*:\s*(.*?)\s*$/);
      if (m) out[m[1]] = m[2];
    });
    return out;
  };

  // ------------------------------------------------------------ paths
  // 'items/#12/stories/@s1/done' -> segments; '#12' = item by num (in an
  // array of items), '@id' = element by id, digits = index, '-' = append
  function splitPath(path) {
    return String(path || '').split('/').map(function (s) { return s.trim(); }).filter(function (s) { return s !== ''; });
  }
  function findIndex(arr, seg) {
    if (!Array.isArray(arr)) return -1;
    if (seg.charAt(0) === '#') {
      var n = +seg.slice(1);
      for (var i = 0; i < arr.length; i++) if (arr[i] && arr[i].num === n) return i;
      return -1;
    }
    if (seg.charAt(0) === '@') {
      var id = seg.slice(1);
      for (var j = 0; j < arr.length; j++) if (arr[j] && arr[j].id === id) return j;
      return -1;
    }
    if (/^\d+$/.test(seg)) { var k = +seg; return k < arr.length ? k : -1; }
    return -1;
  }
  // { parent, key } for the last segment; creates intermediate objects when
  // `create` is set. Throws with a readable message when a segment misses.
  AI.resolvePath = function (state, path, create) {
    var segs = splitPath(path);
    if (!segs.length) throw new Error('empty path');
    var cur = state;
    for (var i = 0; i < segs.length - 1; i++) {
      var seg = segs[i];
      var next;
      if (Array.isArray(cur)) {
        var idx = findIndex(cur, seg);
        if (idx === -1) throw new Error('no element "' + seg + '" at ' + segs.slice(0, i + 1).join('/'));
        next = cur[idx];
      } else if (isObj(cur)) {
        next = cur[seg];
        if (next == null) {
          if (!create) throw new Error('nothing at ' + segs.slice(0, i + 1).join('/'));
          next = cur[seg] = /^\d+$/.test(segs[i + 1]) || segs[i + 1] === '-' ? [] : {};
        }
      } else {
        throw new Error('cannot descend into ' + segs.slice(0, i).join('/'));
      }
      cur = next;
    }
    var last = segs[segs.length - 1];
    if (Array.isArray(cur)) {
      if (last === '-') return { parent: cur, key: cur.length, append: true };
      var li = findIndex(cur, last);
      if (li === -1) throw new Error('no element "' + last + '" at ' + path);
      return { parent: cur, key: li };
    }
    if (!isObj(cur)) throw new Error('cannot set ' + path);
    return { parent: cur, key: last };
  };

  // ops: [{ op: 'set' | 'delete' | 'push', path, value }]
  AI.applyOps = function (state, ops) {
    var next = clone(state);
    var changes = [];
    if (!Array.isArray(ops) || !ops.length) throw new Error('ops must be a non-empty array');
    ops.forEach(function (o, i) {
      if (!o || typeof o !== 'object') throw new Error('op ' + i + ' is not an object');
      var op = o.op || 'set';
      if (!o.path) throw new Error('op ' + i + ' has no path');
      if (op === 'set') {
        if (o.value === undefined) throw new Error('op ' + i + ' (set ' + o.path + ') has no value');
        var r = AI.resolvePath(next, o.path, true);
        var before = r.append ? undefined : r.parent[r.key];
        if (r.append) r.parent.push(clone(o.value)); else r.parent[r.key] = clone(o.value);
        changes.push({ op: 'set', path: o.path, before: before, after: o.value });
      } else if (op === 'delete') {
        var d = AI.resolvePath(next, o.path, false);
        var was = d.parent[d.key];
        if (Array.isArray(d.parent)) d.parent.splice(d.key, 1); else delete d.parent[d.key];
        changes.push({ op: 'delete', path: o.path, before: was });
      } else if (op === 'push') {
        if (o.value === undefined) throw new Error('op ' + i + ' (push ' + o.path + ') has no value');
        var p = AI.resolvePath(next, o.path, true);
        var arr = p.append ? p.parent : p.parent[p.key];
        if (arr == null && !p.append) arr = p.parent[p.key] = [];
        if (!Array.isArray(arr)) throw new Error(o.path + ' is not an array');
        arr.push(clone(o.value));
        changes.push({ op: 'push', path: o.path, after: o.value });
      } else {
        throw new Error('unknown op "' + op + '" (use set, delete or push)');
      }
    });
    return { state: RM.normalizeState(next), changes: changes };
  };

  // ------------------------------------------------------------ summaries
  function itemLine(state, it) {
    var meta = state.meta;
    var ph = null;
    state.phases.forEach(function (p) { if (p.id === it.phaseId) ph = p; });
    var o = { num: it.num, feature: it.feature, phase: ph ? ph.name : '', workstream: it.workstream || '', epic: it.epic || '' };
    if (it.size) o.size = it.size;
    if (it.priority) o.priority = it.priority;
    if (it.risk) o.risk = it.risk;
    if (it.milestone) o.milestone = true;
    if (it.type && it.type !== RM.defaultTypeFor(state, 'feature')) o.type = it.type;
    if (it.startDay != null && it.durDays != null) {
      o.start = isoOfDay(meta, it.startDay);
      o.end = spanEndIso(meta, it.startDay, it.durDays);
      o.durDays = it.durDays;
    } else o.start = null;
    if (it.deadline) o.deadline = it.deadline;
    if (it.deps && it.deps.length) o.deps = it.deps.slice();
    if (it.done) o.done = true;
    if (it.locked) o.locked = true;
    if (it.noAuto) o.noAuto = true;
    if (it.flag) o.flag = it.flag.reason || true;
    if (it.jiraKey) o.jiraKey = it.jiraKey;
    if (it.tags && it.tags.length) o.tags = it.tags.slice();
    if (it.stories && it.stories.length) {
      o.stories = it.stories.length;
      var dn = it.stories.filter(function (s) { return s.done; }).length;
      if (dn) o.storiesDone = dn;
      var sn = it.stories.map(function (s) { return s.num; }).filter(function (n) { return n != null; });
      if (sn.length) o.storyNums = sn;
    }
    return o;
  }
  function itemFull(state, it) {
    var meta = state.meta;
    var o = clone(it);
    o.start = isoOfDay(meta, it.startDay);
    o.end = it.startDay != null && it.durDays != null ? spanEndIso(meta, it.startDay, it.durDays) : null;
    ['description', 'ac', 'enables', 'outOfScope', 'notes', 'extDeps'].forEach(function (k) { o[k] = RM.htmlToText(it[k]); });
    o.stories = (it.stories || []).map(function (s) {
      var so = clone(s);
      so.start = isoOfDay(meta, s.startDay);
      so.end = s.startDay != null && s.durDays ? spanEndIso(meta, s.startDay, s.durDays) : null;
      so.description = RM.htmlToText(s.description);
      so.ac = RM.htmlToText(s.ac);
      return so;
    });
    return o;
  }
  AI.summary = function (state) {
    var m = state.meta;
    var epics = {};
    state.items.forEach(function (it) { if (it.epic) epics[it.epic] = (epics[it.epic] || 0) + 1; });
    var counts = {};
    state.items.forEach(function (it) { counts[it.phaseId] = (counts[it.phaseId] || 0) + 1; });
    var stats = RM.scheduleStats(state);
    var si = RM.sprintInfo(m);
    return {
      title: m.title,
      vision: m.vision || '',
      timeline: {
        start: m.timelineStart, end: m.endDate || null, weeks: m.numWeeks, weeksPerSprint: m.weeksPerSprint,
        workDays: m.workDays, sprintAnchor: m.sprintAnchor, sprintAnchorNum: si.firstNum,
        holidays: (m.holidayRanges || []).map(function (r) { return r.name ? r.name + ' ' + r.start + (r.end !== r.start ? '..' + r.end : '') : r.start + (r.end !== r.start ? '..' + r.end : ''); }),
        lastScheduledDay: isoOfDay(m, stats.lastDay != null ? stats.lastDay - 1 : null)
      },
      schemes: {
        featureSize: m.sizeScheme, featureSizes: m.sizeOrder, featureSizeDays: m.sizeDays,
        storySize: m.storySizeScheme, storySizes: m.storySizeOrder,
        featurePriority: m.priorityScheme, storyPriority: m.storyPriorityScheme, risk: m.riskScheme, storyRisk: m.storyRiskScheme,
        capacityEnabled: !!m.capacityEnabled, workstreamsEnabled: m.workstreamsEnabled !== false,
        capMode: m.capMode, defaultPoints: m.defaultPoints, planLevel: m.planLevel
      },
      scopeColumns: (m.scopeCols || []).map(function (c) { return { key: c.key, label: RM.scopeColLabel(c), scope: c.scope || 'both' }; }),
      phases: state.phases.map(function (p) {
        return { id: p.id, name: p.name, bucket: !!p.bucket, items: counts[p.id] || 0,
          start: isoOfDay(m, p.startDay), end: isoOfDay(m, p.endDay) };
      }),
      workstreams: (state.wsOrder || []).map(function (w) { return { name: w, color: state.wsColors[w] || null }; }),
      epics: Object.keys(epics).map(function (e) { return { name: e, items: epics[e], icon: state.epicIcons[e] || null, jiraKey: state.epicJira[e] || null }; }),
      team: (state.team || []).map(function (t) {
        return { id: t.id, name: t.name, role: t.role || '', type: t.type || '', capType: t.capType || '', points: t.points, workstreams: t.workstreams || [], capacity: t.capacity, rate: t.rate || 0, cost: t.cost || 0 };
      }),
      teamTypes: state.teamTypes,
      items: state.items.map(function (it) { return itemLine(state, it); }),
      scheduled: stats.scheduled, unscheduled: stats.unscheduled,
      jira: jiraSummary(state)
    };
  };
  // Jira: mapping from the document, connection from this machine
  function jiraSummary(state) {
    var m = state.meta;
    var JR = root.HeadwayJira;
    var out = { project: (m.jira && m.jira.project) || null, lastSync: (m.jira && m.jira.lastSync) || null, connected: !!(JR && JR.hasCreds && JR.hasCreds()) };
    if (JR && JR.status) {
      var st = JR.status(state);
      out.status = st.label; out.linked = st.linked; out.total = st.total;
    }
    out.canSync = !!(out.connected && out.project);
    return out;
  }
  function validationReport(state, v) {
    var out = { counts: v.counts, global: (v.global || []).map(function (g) { return g.level + ': ' + g.msg; }), items: [] };
    state.items.forEach(function (it) {
      var list = v.byItem[it.id];
      if (list && list.length) out.items.push({ num: it.num, feature: it.feature, issues: list.map(function (x) { return x.level + ': ' + x.msg; }) });
    });
    return out;
  }

  // ------------------------------------------------------------ tools
  AI.PREF_KEYS = {
    theme: "'system' | 'light' | 'dark'",
    userName: 'string — the author name recorded in Version history',
    snapFeat: "'day' | 'week' | 'sprint' — drag/resize grid for feature bars",
    snapStory: "'day' | 'week' | 'sprint' — grid for story bars",
    deps: 'boolean — dependency arrows on the timeline',
    crit: 'boolean — critical path highlight',
    cap: 'boolean — capacity row',
    autoOrder: 'boolean — re-sort rows by start date after moves',
    groupWs: 'boolean — group rows by workstream',
    groupEpic: 'boolean — group rows by epic',
    colorBy: "'workstream' | 'epic' | 'assignee' | 'priority' | 'type' — what bar colours follow",
    autoSave: 'boolean — desktop: write to the open file automatically',
    detailMode: "'feature' | 'story' — Planning row detail level"
  };
  AI.VIEWS = ['scoping', 'prio', 'planning', 'sprints', 'budget', 'reports', 'setup', 'history'];

  AI.TOOLS = [
    {
      name: 'get_project',
      description: 'Read the open project. part="summary" (default) gives the timeline, schemes, phases, workstreams, epics, team and one compact line per feature. part="items" with nums gives full features including stories and text fields (ask for the nums you need, not everything). part="meta" | "phases" | "team" returns those raw sections. part="validation" lists preflight errors and warnings. part="history" lists recent version-history entries.',
      parameters: {
        type: 'object',
        properties: {
          part: { type: 'string', enum: ['summary', 'items', 'meta', 'phases', 'team', 'validation', 'history'] },
          nums: { type: 'array', items: { type: 'integer' }, description: 'feature numbers for part="items"' }
        }
      }
    },
    {
      name: 'get_preferences',
      description: 'Read the current UI state (view, selected feature, desktop or browser) and the personal preferences, with the option lists and the keys set_preference accepts.',
      parameters: { type: 'object', properties: {} }
    },
    {
      name: 'add_items',
      description: 'Create features in a phase. Each item: feature (title, required), type (Feature, Bug, Task, … — a type label or key from Setup → Organization), workstream, epic, size, priority, risk, description, enables, outOfScope, notes, extDeps, deps (feature numbers), start (ISO date), durDays or end (ISO date), deadline (ISO), milestone (boolean, zero duration; milestones ignore size and priority), headcount, teamType, tags (array of strings), stories ([{title, type, description, ac, size, priority, done, tags, deps}]). Stories get their own number from the same pool as features. Returns the new feature numbers.',
      parameters: {
        type: 'object',
        properties: {
          phase: { type: 'string', description: 'phase name or id; the first phase when omitted' },
          items: { type: 'array', items: { type: 'object' } }
        },
        required: ['items']
      }
    },
    {
      name: 'update_items',
      description: 'Change features or stories. Each update: num (a feature number, or a story number to change that story instead — required), story (story id; the older way to reach a story, still accepted), fields (object merged into the target). Feature fields: feature, type (Feature, Bug, Task, … — a type label or key from Setup → Organization), workstream, epic, size, priority, risk, description, enables, outOfScope, notes, extDeps, deps, start (ISO date or null to unschedule), durDays, end (ISO date), deadline, milestone, headcount, teamType, locked, noAuto (true = excluded from the Auto timeline / ⚡: the scheduler leaves it where it is; mutually exclusive with locked — setting one clears the other), done, phase (name or id), assignees (team ids), tags (array of strings — replaces the list), custom ({columnKey: text}), jiraKey, addStories ([{title, type, description, ac, size, priority}] appends stories). Story fields: title, type, description, ac, size, priority, risk, done, noAuto (excluded from the Auto timeline), start, durDays, end, deadline, assignees, tags (array of strings), deps (story numbers this story depends on — stories link to stories, never to features). Use delete: true to remove the target.',
      parameters: {
        type: 'object',
        properties: {
          updates: { type: 'array', items: { type: 'object' } },
          label: { type: 'string', description: 'short history label, e.g. "resize #12"' }
        },
        required: ['updates']
      }
    },
    {
      name: 'update_project',
      description: 'Edit any other part of the document with path operations, e.g. project settings (meta/title, meta/vision, meta/timelineStart, meta/endDate, meta/weeksPerSprint, meta/sprintAnchor, meta/sprintAnchorNum, meta/workDays, meta/sizeScheme, meta/sizeDays/M, meta/priorityScheme, meta/storyPriorityScheme, meta/riskScheme, meta/storyRiskScheme (\"none\"|\"risk\"|\"confidence\"), meta/capacityEnabled, meta/capMode ("person"|"points"), meta/defaultPoints, meta/planLevel, meta/holidayRanges (push {name,start,end}), meta/scopeCols (push {key:"c<slug>", label}), meta/jira, meta/itemTypes (array of {key,label,icon,jira}), meta/hierarchy/levels/<i>/types (each type key at exactly one level), epicTypes/<name>), phases (phases/@id/name, phases/- to append {name, bucket}), team (team/@id/rate, team/- to append {name, role, type, workstreams, capacity, rate, cost}), teamTypes, wsColors/<name>, epicIcons/<name> (lucide icon), epicJira/<name>, wsOrder. Path segments: #num = feature by number, @id = element by id, digits = index, "-" = append. Ops: set (path, value), delete (path), push (path, value). Prefer add_items / update_items for features and stories.',
      parameters: {
        type: 'object',
        properties: {
          ops: { type: 'array', items: { type: 'object', properties: { op: { type: 'string', enum: ['set', 'delete', 'push'] }, path: { type: 'string' }, value: {} }, required: ['path'] } },
          label: { type: 'string', description: 'short history label' }
        },
        required: ['ops']
      }
    },
    {
      name: 'set_preference',
      description: 'Change a personal preference (this machine only). Keys: ' + Object.keys(AI.PREF_KEYS).map(function (k) { return k + ' (' + AI.PREF_KEYS[k] + ')'; }).join('; ') + '.',
      parameters: { type: 'object', properties: { key: { type: 'string' }, value: {} }, required: ['key', 'value'] }
    },
    {
      name: 'sync_jira',
      description: 'Sync the project with Jira Cloud (needs the Jira connection on this machine and a project key in Setup → Jira Integration; get_project summary shows jira.canSync). Creates issues for features (and stories, when that option is on) that have no Jira key, updates linked issues from Headway, adds "blocks" links for dependencies, and reads Done state back from Jira. Call with dryRun: true first to see what would change, and confirm with the user before applying. Adjust the mapping (epics, stories) through update_project on meta/jira, and issue types through meta/itemTypes.',
      parameters: { type: 'object', properties: { dryRun: { type: 'boolean', description: 'true = preview only' } } }
    },
    {
      name: 'navigate',
      description: 'Show something to the user: switch the view (scoping, prio, planning, sprints, budget, reports, setup, history) and/or select a feature by number.',
      parameters: { type: 'object', properties: { view: { type: 'string' }, num: { type: 'integer' } } }
    }
  ];
  AI.toolByName = function (name) {
    for (var i = 0; i < AI.TOOLS.length; i++) if (AI.TOOLS[i].name === name) return AI.TOOLS[i];
    return null;
  };

  function phaseIdOf(state, ref) {
    if (ref == null || ref === '') return state.phases[0] ? state.phases[0].id : null;
    var r = String(ref).trim().toLowerCase();
    for (var i = 0; i < state.phases.length; i++) {
      var p = state.phases[i];
      if (p.id === ref || String(p.name || '').trim().toLowerCase() === r) return p.id;
    }
    throw new Error('no phase "' + ref + '" (phases: ' + state.phases.map(function (p) { return p.name; }).join(', ') + ')');
  }
  function textish(v) { return v == null ? '' : String(v); }
  // shared field merge for features and stories; ISO start/end become days
  function mergeFields(meta, target, fields, isStory) {
    var changed = [];
    Object.keys(fields).forEach(function (k) {
      var v = fields[k];
      if (k === 'start') {
        if (v == null || v === '') { target.startDay = null; if (isStory) target.durDays = target.durDays || null; else target.durDays = null; }
        else {
          var d = dayOfIso(meta, v);
          if (d == null) throw new Error('start "' + v + '" is not an ISO date inside the timeline (' + meta.timelineStart + ' to ' + meta.endDate + '; extend meta/endDate first)');
          target.startDay = d;
          if (target.durDays == null) target.durDays = target.milestone ? 0 : 5;
        }
        changed.push('start');
      } else if (k === 'end') {
        var e = dayOfIso(meta, v);
        if (e == null) throw new Error('end "' + v + '" is not an ISO date inside the timeline (' + meta.timelineStart + ' to ' + meta.endDate + '; extend meta/endDate first)');
        if (target.startDay == null) throw new Error('set start before end');
        target.durDays = Math.max(target.milestone ? 0 : 1, e - target.startDay + 1);
        changed.push('end');
      } else if (k === 'durDays') {
        target.durDays = v == null ? null : Math.max(0, Math.round(+v));
        changed.push('durDays');
      } else if (k === 'phase') {
        target.phaseId = phaseIdOf({ phases: fields.__phases || [] }, v);
        changed.push('phase');
      } else if (k === 'tags') {
        target.tags = RM.normalizeTags(v);
        changed.push('tags');
      } else if (k === 'deps') {
        target.deps = (Array.isArray(v) ? v : [v]).map(Number).filter(function (n) { return !isNaN(n); });
        changed.push('deps');
      } else if (k === 'flag') {
        target.flag = RM.normalizeFlag(v); // true / "reason" / null
        changed.push('flag');
      } else if (k === 'locked' && !isStory) {
        RM.setLocked(target, !!v); // clears noAuto
        changed.push('locked');
      } else if (k === 'noAuto') {
        // Lock and Exclude-from-Auto are exclusive; asked for both, Lock wins
        if (!(v && fields.locked && !isStory)) RM.setNoAuto(target, !!v);
        changed.push('noAuto');
      } else if (k === 'num' || k === 'id') {
        throw new Error(k + ' is assigned by Headway and cannot be set — the user renumbers in the panel');
      } else if (k === 'type') {
        var want = String(v).trim().toLowerCase();
        var hit = RM.itemTypes({ meta: meta }).filter(function (t) { return t.key.toLowerCase() === want || t.label.toLowerCase() === want; })[0];
        if (!hit) throw new Error('unknown type "' + v + '"');
        target.type = hit.key;
        changed.push('type');
      } else if (k === 'stories' || k === 'addStories') {
        if (isStory) throw new Error('a story has no stories of its own');
        var made = (Array.isArray(v) ? v : []).map(function (s) {
          var st = { id: RM.uid('s'), title: '', done: false };
          mergeFields(meta, st, isObj(s) ? s : { title: textish(s) }, true);
          return st;
        });
        target.stories = k === 'stories' ? made : (target.stories || []).concat(made);
        changed.push(k);
      } else if (k === '__phases') {
        /* internal */
      } else if (['description', 'ac', 'enables', 'outOfScope', 'notes', 'extDeps'].indexOf(k) !== -1) {
        target[k] = textToHtml(v);
        changed.push(k);
      } else {
        target[k] = v;
        changed.push(k);
      }
    });
    return changed;
  }
  // Stories the tool just created carry no number yet; normalize would only give
  // them one on the next load, and the targeted apply below compares base and
  // next by JSON — so number them here, in the state the tool is about to commit.
  function numberNewStories(state) {
    state.items.forEach(function (it) {
      (it.stories || []).forEach(function (st) {
        if (st.num == null) st.num = RM.nextNum(state);
      });
    });
  }

  // plain text (blank-line paragraphs) -> the sanitized-HTML shape rich fields hold
  function textToHtml(v) {
    var t = textish(v);
    if (!t) return '';
    if (/<[a-z][\s\S]*>/i.test(t)) return t;
    return t.replace(/\r\n?/g, '\n').split(/\n{2,}/).map(function (para) {
      return '<p>' + esc(para).replace(/\n/g, '<br>') + '</p>';
    }).join('');
  }
  AI.textToHtml = textToHtml;

  // Apply only what the tool changed. The model thinks for seconds between
  // reading the project (base) and writing (next) while the user keeps
  // editing the live document; items and sections the tool left alone stay
  // exactly as the user has them — object identity included — so the screen
  // updates in place instead of reloading.
  function sameJson(a, b) { return JSON.stringify(a === undefined ? null : a) === JSON.stringify(b === undefined ? null : b); }
  function applyItemChanges(s, baseItems, nextItems) {
    var live = Array.isArray(s.items) ? s.items : (s.items = []);
    var baseById = {}, nextById = {}, liveIdx = {}, liveNums = {};
    baseItems.forEach(function (it) { baseById[it.id] = it; });
    nextItems.forEach(function (it) { nextById[it.id] = it; });
    // deleted by the tool
    for (var i = live.length - 1; i >= 0; i--) if (baseById[live[i].id] && !nextById[live[i].id]) live.splice(i, 1);
    live.forEach(function (it, idx) { liveIdx[it.id] = idx; liveNums[it.num] = true; });
    nextItems.forEach(function (it) {
      var b = baseById[it.id];
      if (!b) { // added
        if (liveIdx[it.id] != null) return;
        if (liveNums[it.num]) it.num = RM.nextNum(s); // the user added a feature meanwhile
        live.push(it); liveIdx[it.id] = live.length - 1; liveNums[it.num] = true;
        return;
      }
      if (sameJson(b, it)) return; // untouched: the user's copy stays
      if (liveIdx[it.id] != null) live[liveIdx[it.id]] = it; // (deleted by the user meanwhile: stays deleted)
    });
    // numbers are one pool across features and stories: a number the tool
    // handed a new story (or feature) may meanwhile have gone to something
    // the user added — later occurrences move to the next free number
    // (existing stories keep theirs: a new feature that collides moves)
    var storyNums = {}, seenNum = {};
    live.forEach(function (it) {
      if (baseById[it.id]) (it.stories || []).forEach(function (st) { if (st.num != null) storyNums[st.num] = true; });
    });
    live.forEach(function (it) {
      if (seenNum[it.num] || (!baseById[it.id] && storyNums[it.num])) it.num = RM.nextNum(s);
      seenNum[it.num] = true;
    });
    live.forEach(function (it) {
      (it.stories || []).forEach(function (st) {
        if (st.num == null || seenNum[st.num]) st.num = RM.nextNum(s);
        seenNum[st.num] = true;
      });
    });
    // order: follow the tool only when it reordered something
    var baseOrder = baseItems.map(function (it) { return it.id; }).filter(function (id) { return nextById[id]; });
    var nextOrder = nextItems.map(function (it) { return it.id; }).filter(function (id) { return baseById[id]; });
    if (baseOrder.join('\n') !== nextOrder.join('\n')) {
      var rank = {};
      nextItems.forEach(function (it, i) { rank[it.id] = i; });
      live.sort(function (a, b) {
        var ra = rank[a.id], rb = rank[b.id];
        if (ra == null && rb == null) return 0;
        if (ra == null) return 1;
        if (rb == null) return -1;
        return ra - rb;
      });
    }
  }
  function applyChanges(s, base, next) {
    Object.keys(next).forEach(function (k) {
      if (k === 'history') return;
      if (k === 'items') { applyItemChanges(s, base.items || [], next.items || []); return; }
      if (!sameJson(base[k], next[k])) s[k] = next[k];
    });
    Object.keys(base).forEach(function (k) {
      if (k !== 'history' && k !== 'items' && !(k in next)) delete s[k];
    });
  }
  AI.applyChanges = applyChanges;
  function commitState(A, label, base, next, extra) {
    A.ai.commit(label, function (s) { applyChanges(s, base, next); });
    var v = A.ai.validation();
    var out = { ok: true, validation: v.counts };
    var errs = [];
    Object.keys(v.byItem || {}).forEach(function (id) {
      v.byItem[id].forEach(function (x) { if (x.level === 'error') errs.push(x.msg); });
    });
    if (errs.length) out.errors = errs.slice(0, 12);
    Object.keys(extra || {}).forEach(function (k) { out[k] = extra[k]; });
    return out;
  }

  // -> result object (JSON-able); throws on bad input so the model sees why
  AI.runTool = function (name, args, A) {
    args = isObj(args) ? args : {};
    var tool = AI.toolByName(name);
    if (!tool) throw new Error('unknown tool "' + name + '"');
    if (name === 'get_preferences') {
      var ui = A.ai.ui();
      return { ui: ui, preferenceKeys: AI.PREF_KEYS, views: AI.VIEWS };
    }
    if (name === 'set_preference') {
      if (!AI.PREF_KEYS[args.key]) throw new Error('unknown preference "' + args.key + '"');
      var okp = A.ai.setPref(args.key, args.value);
      if (!okp) throw new Error('value not accepted for ' + args.key);
      return { ok: true, key: args.key, value: args.value };
    }
    if (name === 'navigate') {
      var did = {};
      if (args.view) {
        if (AI.VIEWS.indexOf(args.view) === -1) throw new Error('unknown view "' + args.view + '"');
        did.view = A.ai.setView(args.view);
      }
      if (args.num != null) {
        if (!A.ai.hasDoc()) throw new Error('no project is open');
        did.selected = A.ai.selectNum(+args.num);
        if (!did.selected) throw new Error('no feature #' + args.num);
      }
      return did;
    }
    if (!A.ai.hasDoc()) throw new Error('no project is open — ask the user to open or create one');
    var state = A.ai.state();
    if (name === 'get_project') {
      var part = args.part || 'summary';
      if (part === 'summary') return AI.summary(state);
      if (part === 'items') {
        var nums = Array.isArray(args.nums) ? args.nums.map(Number) : [];
        if (!nums.length) throw new Error('give nums for part="items" (use the summary to find them)');
        var missing = [];
        var found = nums.map(function (n) {
          var it = RM.itemByNum(state, n);
          if (!it) { missing.push(n); return null; }
          return itemFull(state, it);
        }).filter(Boolean);
        var r = { items: found };
        if (missing.length) r.missing = missing;
        return r;
      }
      if (part === 'meta') return state.meta;
      if (part === 'phases') return state.phases.map(function (p) { var o = clone(p); o.start = isoOfDay(state.meta, p.startDay); o.end = isoOfDay(state.meta, p.endDay); return o; });
      if (part === 'team') return state.team;
      if (part === 'validation') return validationReport(state, A.ai.validation());
      if (part === 'history') {
        return (state.history || []).slice(-25).reverse().map(function (h) {
          return { when: new Date(h.t).toISOString(), who: h.u, label: h.label, changes: (h.d || []).slice(0, 8).map(function (d) { return d.join(' · '); }) };
        });
      }
      throw new Error('unknown part "' + part + '"');
    }
    if (name === 'add_items') {
      var list = Array.isArray(args.items) ? args.items : [];
      if (!list.length) throw new Error('items is empty');
      var pid = phaseIdOf(state, args.phase);
      var next = clone(state);
      var nextNum = RM.nextNum(next);
      var made = [];
      list.forEach(function (spec) {
        if (!isObj(spec) || !textish(spec.feature).trim()) throw new Error('every item needs a feature title');
        var it = { id: RM.uid('i'), num: nextNum++, phaseId: pid, feature: '', stories: [], deps: [] };
        var f = clone(spec);
        if (f.phase) { it.phaseId = phaseIdOf(state, f.phase); delete f.phase; }
        if (f.milestone) { it.milestone = true; if (f.durDays == null) f.durDays = 0; }
        mergeFields(next.meta, it, f, false);
        if (it.startDay != null && it.durDays == null) it.durDays = it.milestone ? 0 : 5;
        next.items.push(it);
        made.push({ num: it.num, feature: it.feature });
      });
      numberNewStories(next);
      var norm = RM.normalizeState(next);
      return commitState(A, args.label || ('add ' + (made.length === 1 ? '#' + made[0].num + ' ' + made[0].feature : made.length + ' features')), state, norm, { created: made });
    }
    if (name === 'update_items') {
      var ups = Array.isArray(args.updates) ? args.updates : [];
      if (!ups.length) throw new Error('updates is empty');
      var st2 = clone(state);
      var report = [];
      ups.forEach(function (u) {
        if (!isObj(u) || u.num == null) throw new Error('every update needs a num');
        // one number pool: a num is a feature or a story (the story: <id> form still works)
        var hit = RM.byNum(st2, +u.num);
        if (!hit) throw new Error('no feature or story #' + u.num);
        var it = hit.it, isStory = hit.kind === 'story';
        var target = isStory ? hit.st : it;
        var label = isStory ? '#' + hit.st.num + ' ' + (hit.st.title || 'story') : '#' + it.num;
        if (u.story) {
          if (isStory) throw new Error('#' + u.num + ' is already a story; drop the story field');
          target = null;
          it.stories.forEach(function (s) { if (s.id === u.story) target = s; });
          if (!target) throw new Error('feature #' + it.num + ' has no story "' + u.story + '"');
          isStory = true;
          label += ' story ' + target.title;
        }
        if (u['delete'] === true) {
          if (isStory) it.stories = it.stories.filter(function (s) { return s !== target; });
          else st2.items = st2.items.filter(function (x) { return x !== it; });
          report.push({ target: label, deleted: true });
          return;
        }
        var fields = isObj(u.fields) ? clone(u.fields) : {};
        if (fields.phase != null) fields.__phases = st2.phases;
        var ch = mergeFields(st2.meta, target, fields, isStory);
        report.push({ target: label, changed: ch });
      });
      numberNewStories(st2);
      var lbl = args.label || (report.length === 1 ? 'edit ' + report[0].target : 'edit ' + report.length + ' items');
      return commitState(A, lbl, state, RM.normalizeState(st2), { updated: report });
    }
    if (name === 'update_project') {
      var res = AI.applyOps(state, args.ops);
      var lbl2 = args.label || (res.changes.length === 1 ? res.changes[0].op + ' ' + res.changes[0].path : res.changes.length + ' changes');
      return commitState(A, lbl2, state, res.state, { changes: res.changes.map(function (c) { return c.op + ' ' + c.path; }) });
    }
    if (name === 'sync_jira') return syncJira(A, state, !!args.dryRun);
    throw new Error('tool "' + name + '" is not implemented');
  };
  // -> Promise: the Jira sync (preview or apply) through js/jira.js
  function syncJira(A, state, dryRun) {
    var JR = root.HeadwayJira;
    if (!JR) throw new Error('the Jira module is not loaded');
    var creds = JR.loadCreds();
    var cfg = JR.cfgOf(state, creds);
    if (!creds.site || !creds.email || !creds.token) throw new Error('Jira is not connected on this machine — the user sets site, email and API token in Setup → Jira Integration');
    if (!cfg.project) throw new Error('no Jira project key is mapped — set it in Setup → Jira Integration (or update_project meta/jira/project)');
    var client = JR.client(creds);
    var look = JR.discover ? JR.discover(client, state, cfg) : JR.fetchRemote(client, JR.keysOf(state, cfg)).then(function (remote) { return { remote: remote }; });
    return look.then(function (info) {
      var plan = JR.plan(state, cfg, info);
      var preview = {
        project: cfg.project, dryRun: dryRun, counts: plan.counts,
        create: { epics: plan.epics.map(function (e) { return e.name; }), features: plan.features.map(function (f) { return '#' + f.num + ' ' + f.title; }), stories: plan.stories.map(function (st) { return st.title; }) },
        update: plan.updates.map(function (u) { return u.key + ' ' + u.title; }),
        readBack: plan.pulls.map(function (pl) { return pl.key + ' ' + pl.title + ' → ' + (pl.done ? 'done' : 'not done') + ' (' + pl.status + ')'; }),
        missingInJira: plan.missing.map(function (x) { return x.key + ' ' + x.title; })
      };
      var c = plan.counts;
      if (dryRun) return preview;
      if (plan.notes && plan.notes.length) preview.notes = plan.notes;
      if (plan.people && plan.people.unresolved.length) preview.peopleNotInJira = plan.people.unresolved;
      if (plan.sprints) preview.sprints = { create: plan.sprints.create.map(function (sp) { return sp.name; }), place: plan.sprints.assign.length };
      if (!c.create && !c.update && !c.pull && !c.link && !c.sprintAssign) { preview.nothingToSync = true; return preview; }
      return JR.apply(plan, client).then(function (result) {
        A.ai.commit('jira sync', function (s) {
          JR.applyToState(s, plan, result, info);
          s.meta.jira = s.meta.jira || {};
          s.meta.jira.lastSync = new Date().toISOString();
          s.meta.jira.lastErrors = result.errors.length;
          if (JR.fingerprint) s.meta.jira.syncedHash = JR.fingerprint(s);
        });
        var keys = {};
        plan.features.forEach(function (f) { if (f.key) keys['#' + f.num] = f.key; });
        return { ok: true, project: cfg.project, created: result.created, updated: result.updated, linked: result.linked,
          sprintsCreated: result.sprintsCreated || 0, placedInSprints: result.sprintAssigned || 0,
          readBack: plan.pulls.length, newKeys: keys, errors: result.errors };
      });
    });
  }

  // ------------------------------------------------------------ system prompt
  AI.GUIDE = [
    '# Headway',
    'Headway is a standalone roadmap planning tool (browser page or desktop app) that reads and writes one .xlsx document per project. Everything is local; the file is the source of truth. Version history (who changed what) travels inside the file.',
    '',
    '## Views (top tab group)',
    '- Scoping: a spreadsheet of features and stories — fixed chip columns (Size, Risk, Workstream, Epic) then text columns (Enables, Out of scope, External dependencies, Notes, Description; custom columns can be added in Setup → Columns).',
    '- Prioritizing: a kanban of the phases; cards sort by priority; filters by epic/workstream; RICE scoring when that scheme is on.',
    '- Planning: the timeline / gantt. Bars are features (or stories at story detail); drag to move, edges resize, ⌘-drag pushes dependents. Milestones are zero-duration diamonds. Dependency arrows, critical path, capacity row and the resources panel live here.',
    '- Sprinting: sprint-by-sprint rows; drag rows between sprints to reschedule.',
    '- Budgeting: one row per team role with hourly cost and rate, margin, and week-hours; totals price actual hours.',
    '- Reporting: dashboard (done / scheduled / by workstream or phase).',
    '- Setup (gear): project settings — Timeline (start, end, work week, sprint length and numbering, holidays), Phases, Workstreams, Team, Columns, Sizing (size, risk and priority schemes for features and stories), Jira. Personal: Appearance, Preferences, AI assistant.',
    '- History (clock): version history with timeline diffs.',
    'Undo is ⌘Z / Ctrl+Z. Save writes the .xlsx; the desktop app auto-saves the open file.',
    '',
    '## Model',
    '- Time is counted in working days from meta.timelineStart (weekends and non-work days do not exist in the index). Holidays stretch bars. A sprint = meta.weeksPerSprint weeks; sprint numbers count from meta.sprintAnchor / sprintAnchorNum. Tools accept and report ISO dates; day indexes appear in raw sections.',
    '- Phases hold features (state.phases; each item has phaseId). bucket=true phases are backlog shelves (Next / Future).',
    '- Features (state.items) have num (the user-facing #id), feature (title), workstream, epic, size, risk, priority, deps (numbers of features that must finish first), startDay/durDays (null = unscheduled), deadline, milestone, locked, noAuto (excluded from the Auto timeline: Auto / ⚡ leave it where it sits, Place at earliest slot still moves it; never together with locked), done, headcount, teamType, assignees (team ids), rich-text fields (description, enables, outOfScope, notes, extDeps — plain text is fine when writing), custom column values, jiraKey, tags (free-form labels shared with stories, exported as Jira labels), and stories, type (Feature / Bug / Task …; types live in meta.itemTypes, and meta.hierarchy puts each type at exactly one level (a level\'s items hold children of the level below), and each type\'s jira field is the Jira issue type used by sync).',
    '- Stories belong to a feature: id, num, title, done, noAuto (excluded from the Auto timeline; a feature’s noAuto covers its stories), size, priority, risk, description, ac (acceptance criteria — a built-in column shown on stories by default), optional own startDay/durDays, deadline, assignees, jiraKey, deps. A story number comes from the same pool as feature numbers, so every # in the document is either a feature or a story; refer to a story by its number (update_items takes it as num). Stories can depend on other stories: story deps hold story numbers, never feature numbers.',
    '- Sizing schemes: feature sizes (t-shirt XS–XL with working days per size in meta.sizeDays, or story points), story sizes, risk (none / L-M-H …), priority (none, MoSCoW M/S/C/W, levels C/H/M/L, RICE). Values are validated against the active scheme; read the summary before setting them.',
    '- Team (state.team): people or seats with role, rate-card type, workstreams, capacity (per-person mode: heads at 40 h, 0.5 = half-time; ignored in points mode except that 0 means supplies nothing), hourly rate and cost, weekHours overrides, capType = what they supply (no capType = supplies nothing); points = story points per sprint (points mode, checked per sprint and scaled by hours). Capacity checks only run when meta.capacityEnabled; Auto timeline is a one-shot button on each phase band (it lays that phase out by dependencies and capacity when clicked), not a stored setting.',
    '- Workstreams carry colour (wsColors, order in wsOrder); epics carry a lucide icon (epicIcons) and optionally a Jira epic key (epicJira).',
    '',
    '## How to work',
    '- Read before you write: call get_project (summary) first in a conversation, then get_project items for the features you will touch. Never guess numbers, names or scheme values.',
    '- Edits show up in place in whatever view the user is on; only the items you changed are touched. Do not call navigate after a write unless the user asked to see something — never switch views on your own.',
    '- Make the smallest edit that does the job, in one tool call when possible. Every write is undoable and shows in Version history as "<user> · AI" — mention that briefly after edits.',
    '- Dates are ISO (YYYY-MM-DD) and must fall inside the timeline. Durations are working days. A feature that gets a start but no duration gets 5 days.',
    '- When a request is ambiguous (which phase, which of two similar features), ask instead of picking.',
    '- Check the validation counts a write returns; if it introduced errors, fix or explain them.',
    '- Refer to features as #num. Keep answers short and concrete; use lists sparingly. You may also answer general project-management and planning questions (estimation, sequencing, risk, scope negotiation, sprint planning) from your own knowledge.',
    '- Jira: when the summary says jira.canSync, sync_jira pushes features (and stories if enabled) to Jira Cloud and reads Done state back; preview with dryRun first and confirm before applying. Jira keys live on features/stories (jiraKey) and epics (epicJira); the mapping (project key, issue types, options) is meta/jira. The connection itself (site, email, token) is set by the user in Setup → Jira Integration.',
    '- You cannot open, save or export files, or change AI settings; tell the user how to do it in the UI instead.'
  ].join('\n');

  // per-turn context prepended to the guide
  AI.systemPrompt = function (ctx) {
    var lines = [AI.GUIDE, '', '## Now'];
    lines.push('- Today: ' + ctx.today + (ctx.userName ? '. User: ' + ctx.userName : ''));
    lines.push('- Running in the ' + (ctx.desktop ? 'desktop app' : 'browser') + '; current view: ' + ctx.view + (ctx.selectedNum != null ? '; selected feature #' + ctx.selectedNum : ''));
    if (ctx.doc) lines.push('- Open project: "' + ctx.doc.title + '" — ' + ctx.doc.items + ' features in ' + ctx.doc.phases + ' phases, timeline ' + ctx.doc.start + ' to ' + ctx.doc.end);
    else lines.push('- No project is open: only questions can be answered until the user opens or creates one.');
    if (ctx.textTools) {
      lines.push('', '## Tools');
      lines.push('You have no built-in tools; use these Headway tools instead. To call one, write a fenced block whose language is ' + AI.FENCE + ' containing one JSON object {"name": …, "args": {…}}. You may write several blocks in one message. Write nothing after the last block: the results come back in the next user message as JSON, and you then continue. Tool definitions (JSON Schema):');
      lines.push('```json', JSON.stringify(AI.TOOLS.map(function (t) { return { name: t.name, description: t.description, parameters: t.parameters }; })), '```');
    }
    return lines.join('\n');
  };

  // ------------------------------------------------------------ markdown
  function inline(s) {
    // s is already escaped
    return s
      .replace(/`([^`\n]+)`/g, function (m, c) { return '<code>' + c + '</code>'; })
      .replace(/\*\*([^*\n]+)\*\*/g, '<b>$1</b>')
      .replace(/(^|[\s(])\*([^*\n]+)\*(?=[\s).,;:!?]|$)/g, '$1<i>$2</i>')
      .replace(/\[([^\]\n]+)\]\((https?:\/\/[^)\s]+)\)/g, '<a href="$2" target="_blank" rel="noopener">$1</a>')
      .replace(/(^|[\s(])#(\d+)\b/g, '$1<a class="ai-ref" data-num="$2" href="#">#$2</a>');
  }
  // GitHub-style pipe tables: split a row into cells (\| is a literal pipe)
  function tableCells(line) {
    var t = line.trim();
    if (t.charAt(0) === '|') t = t.slice(1);
    if (t.charAt(t.length - 1) === '|' && t.charAt(t.length - 2) !== '\\') t = t.slice(0, -1);
    var cells = [], cur = '';
    for (var i = 0; i < t.length; i++) {
      var ch = t.charAt(i);
      if (ch === '\\' && t.charAt(i + 1) === '|') { cur += '|'; i += 1; }
      else if (ch === '|') { cells.push(cur.trim()); cur = ''; }
      else cur += ch;
    }
    cells.push(cur.trim());
    return cells;
  }
  // the delimiter row under a header: cells like ---, :---, ---:, :---:
  function tableAligns(line) {
    if (line.indexOf('-') === -1) return null;
    var cells = tableCells(line);
    var aligns = [];
    for (var i = 0; i < cells.length; i++) {
      var m = cells[i].match(/^(:?)-+(:?)$/);
      if (!m) return null;
      aligns.push(m[1] && m[2] ? 'center' : m[2] ? 'right' : m[1] ? 'left' : '');
    }
    return aligns;
  }
  function tableHtml(header, aligns, rows) {
    function cell(tag, txt, i) {
      var a = aligns[i];
      return '<' + tag + (a ? ' style="text-align:' + a + '"' : '') + '>' + inline(esc(txt || '')) + '</' + tag + '>';
    }
    function row(cells, tag) {
      var out = [];
      for (var i = 0; i < header.length; i++) out.push(cell(tag, cells[i], i));
      return '<tr>' + out.join('') + '</tr>';
    }
    return '<div class="ai-tbl"><table><thead>' + row(header, 'th') + '</thead>' +
      '<tbody>' + rows.map(function (r) { return row(r, 'td'); }).join('') + '</tbody></table></div>';
  }
  AI.md = function (text) {
    var src = String(text || '').replace(/\r\n?/g, '\n');
    var out = [];
    var lines = src.split('\n');
    var i = 0;
    var para = [];
    var listType = null, listItems = [];
    function flushPara() {
      if (para.length) { out.push('<p>' + inline(esc(para.join('\n'))).replace(/\n/g, '<br>') + '</p>'); para = []; }
    }
    function flushList() {
      if (listType) { out.push('<' + listType + '>' + listItems.map(function (li) { return '<li>' + li + '</li>'; }).join('') + '</' + listType + '>'); listType = null; listItems = []; }
    }
    while (i < lines.length) {
      var line = lines[i];
      var fence = line.match(/^\s*```([\w-]*)\s*$/);
      if (fence) {
        flushPara(); flushList();
        var code = [];
        i += 1;
        while (i < lines.length && !/^\s*```\s*$/.test(lines[i])) { code.push(lines[i]); i += 1; }
        i += 1;
        if (fence[1] === AI.FENCE) continue; // tool calls render as cards, not code
        out.push('<pre><code>' + esc(code.join('\n')) + '</code></pre>');
        continue;
      }
      // a table: a header line with a pipe, a delimiter line, then rows until a blank line
      if (line.indexOf('|') !== -1 && i + 1 < lines.length) {
        var aligns = tableAligns(lines[i + 1]);
        var header = aligns ? tableCells(line) : null;
        if (header && header.length === aligns.length) {
          flushPara(); flushList();
          var rows = [];
          i += 2;
          while (i < lines.length && lines[i].trim() && lines[i].indexOf('|') !== -1) { rows.push(tableCells(lines[i])); i += 1; }
          out.push(tableHtml(header, aligns, rows));
          continue;
        }
      }
      var h = line.match(/^\s*(#{1,4})\s+(.*)$/);
      if (h) { flushPara(); flushList(); out.push('<h' + (h[1].length + 2) + '>' + inline(esc(h[2])) + '</h' + (h[1].length + 2) + '>'); i += 1; continue; }
      var li = line.match(/^\s*([-*•]|\d+[.)])\s+(.*)$/);
      if (li) {
        flushPara();
        var type = /^\d/.test(li[1]) ? 'ol' : 'ul';
        if (listType && listType !== type) flushList();
        listType = type;
        listItems.push(inline(esc(li[2])));
        i += 1; continue;
      }
      var q = line.match(/^\s*>\s?(.*)$/);
      if (q) { flushPara(); flushList(); out.push('<blockquote>' + inline(esc(q[1])) + '</blockquote>'); i += 1; continue; }
      if (!line.trim()) { flushPara(); flushList(); i += 1; continue; }
      if (listType && /^\s{2,}/.test(line)) { listItems[listItems.length - 1] += '<br>' + inline(esc(line.trim())); i += 1; continue; }
      flushList();
      para.push(line);
      i += 1;
    }
    flushPara(); flushList();
    return out.join('');
  };

  // ------------------------------------------------------------ SSE
  // feed text chunks; returns the complete `data:` payloads found so far
  AI.sseParser = function () {
    var buf = '';
    return {
      push: function (chunk) {
        buf += chunk;
        var out = [];
        var idx;
        while ((idx = buf.indexOf('\n')) !== -1) {
          var line = buf.slice(0, idx).replace(/\r$/, '');
          buf = buf.slice(idx + 1);
          if (line.indexOf('data:') === 0) out.push(line.slice(5).replace(/^ /, ''));
        }
        return out;
      },
      flush: function () {
        var rest = buf; buf = '';
        var m = rest.replace(/\r$/, '');
        return m.indexOf('data:') === 0 ? [m.slice(5).replace(/^ /, '')] : [];
      }
    };
  };

  // ------------------------------------------------------------ OpenAI-compatible (LiteLLM)
  AI.fetchImpl = null;
  function fetchFn() {
    if (AI.fetchImpl) return AI.fetchImpl;
    var t = root.__TAURI__;
    if (t && t.http && t.http.fetch) return t.http.fetch;
    return root.fetch ? root.fetch.bind(root) : null;
  }
  var openai = AI.openai = {};
  openai.effort = function (e) { return e === 'max' ? 'high' : e; };
  openai.tools = function () {
    return AI.TOOLS.map(function (t) { return { type: 'function', function: { name: t.name, description: t.description, parameters: t.parameters } }; });
  };
  function fileParts(files) {
    return (files || []).map(function (f) {
      if (!f.data) return { type: 'text', text: '[attachment ' + f.name + ' — content not available]' };
      if (f.kind === 'image') return { type: 'image_url', image_url: { url: 'data:' + f.type + ';base64,' + f.data } };
      if (f.kind === 'pdf') return { type: 'file', file: { filename: f.name, file_data: 'data:application/pdf;base64,' + f.data } };
      return { type: 'text', text: '--- file: ' + f.name + ' ---\n' + f.data + '\n--- end of ' + f.name + ' ---' };
    });
  }
  openai.messages = function (system, conv) {
    var out = [{ role: 'system', content: system }];
    conv.forEach(function (m) {
      if (m.role === 'user') {
        var parts = fileParts(m.files);
        if (m.text) parts.unshift({ type: 'text', text: m.text });
        if (!parts.length) parts.push({ type: 'text', text: '(empty)' });
        out.push({ role: 'user', content: parts.length === 1 && parts[0].type === 'text' ? parts[0].text : parts });
      } else if (m.role === 'assistant') {
        if (m.error && !m.text && !(m.toolCalls || []).length) return;
        var a = { role: 'assistant', content: m.text || null };
        if (m.toolCalls && m.toolCalls.length) {
          a.tool_calls = m.toolCalls.map(function (c) {
            return { id: c.id, type: 'function', function: { name: c.name, arguments: JSON.stringify(c.args || {}) } };
          });
        }
        if (m.thinkingBlocks && m.thinkingBlocks.length) a.thinking_blocks = m.thinkingBlocks;
        else if (m.thinking) a.reasoning_content = m.thinking;
        if (a.content == null && !a.tool_calls) a.content = '';
        out.push(a);
      } else if (m.role === 'tool') {
        (m.results || []).forEach(function (r) {
          out.push({ role: 'tool', tool_call_id: r.id, content: typeof r.content === 'string' ? r.content : JSON.stringify(r.content) });
        });
      }
    });
    return out;
  };
  // accumulate streamed deltas into { text, thinking, thinkingBlocks, toolCalls, finish }
  openai.accumulator = function (onEvent) {
    var acc = { text: '', thinking: '', thinkingBlocks: [], toolCalls: [], finish: null, usage: null };
    var calls = {};
    return {
      acc: acc,
      push: function (obj) {
        if (!obj) return;
        if (obj.error) throw new Error(obj.error.message || JSON.stringify(obj.error));
        if (obj.usage) acc.usage = obj.usage;
        var ch = obj.choices && obj.choices[0];
        if (!ch) return;
        var d = ch.delta || ch.message || {};
        var think = d.reasoning_content != null ? d.reasoning_content : d.reasoning;
        if (typeof think === 'string' && think) { acc.thinking += think; if (onEvent) onEvent({ type: 'thinking', delta: think }); }
        if (Array.isArray(d.thinking_blocks)) {
          d.thinking_blocks.forEach(function (b) {
            if (!b || typeof b !== 'object') return;
            var last = acc.thinkingBlocks[acc.thinkingBlocks.length - 1];
            // LiteLLM streams one block in pieces: thinking text first, then the signature
            if (last && last.type === 'thinking' && b.type === 'thinking' && !last.signature) {
              if (b.thinking) last.thinking += b.thinking;
              if (b.signature) last.signature = b.signature;
            } else acc.thinkingBlocks.push({ type: b.type || 'thinking', thinking: b.thinking || '', signature: b.signature || '' });
          });
        }
        if (typeof d.content === 'string' && d.content) { acc.text += d.content; if (onEvent) onEvent({ type: 'text', delta: d.content }); }
        if (Array.isArray(d.tool_calls)) {
          d.tool_calls.forEach(function (tc, k) {
            var idx = tc.index != null ? tc.index : k;
            var c = calls[idx];
            if (!c) { c = calls[idx] = { id: tc.id || ('call_' + idx), name: '', argsText: '' }; acc.toolCalls.push(c); }
            if (tc.id) c.id = tc.id;
            if (tc.function) {
              if (tc.function.name) c.name += tc.function.name;
              if (tc.function.arguments) c.argsText += tc.function.arguments;
            }
          });
        }
        if (ch.finish_reason) acc.finish = ch.finish_reason;
      },
      done: function () {
        acc.thinkingBlocks = acc.thinkingBlocks.filter(function (b) { return b.thinking || b.signature; });
        acc.toolCalls = acc.toolCalls.map(function (c) {
          var args = {};
          if (c.argsText) { try { args = JSON.parse(c.argsText); } catch (e) { args = { __parseError: e.message, raw: c.argsText }; } }
          return { id: c.id, name: c.name, args: args };
        });
        return acc;
      }
    };
  };
  function readStream(res, onChunk) {
    var dec = typeof TextDecoder !== 'undefined' ? new TextDecoder() : null;
    if (res.body && res.body.getReader && dec) {
      var reader = res.body.getReader();
      return (function pump() {
        return reader.read().then(function (r) {
          if (r.done) { onChunk(dec.decode()); return; }
          onChunk(dec.decode(r.value, { stream: true }));
          return pump();
        });
      })();
    }
    return res.text().then(function (t) { onChunk(t); });
  }
  function readError(res) {
    return res.text().then(function (txt) {
      var msg = 'HTTP ' + res.status;
      try {
        var j = JSON.parse(txt);
        var e = j.error || j;
        msg += ': ' + (typeof e === 'string' ? e : (e.message || txt.slice(0, 300)));
      } catch (x) { if (txt) msg += ': ' + txt.slice(0, 300); }
      var err = new Error(msg);
      err.status = res.status;
      err.body = txt;
      return err;
    }, function () { return new Error('HTTP ' + res.status); });
  }
  openai.headers = function (s) {
    var h = { 'Content-Type': 'application/json', Accept: 'text/event-stream' };
    if (s.apiKey) h.Authorization = 'Bearer ' + s.apiKey;
    var extra = AI.parseHeaders(s.headers);
    Object.keys(extra).forEach(function (k) { h[k] = extra[k]; });
    return h;
  };
  openai.body = function (s, system, conv, withEffort) {
    var b = { model: s.model, messages: openai.messages(system, conv), stream: true, tools: openai.tools(), tool_choice: 'auto', stream_options: { include_usage: true } };
    if (withEffort && AI.effortAllowed(s)) b.reasoning_effort = openai.effort(s.effort);
    return b;
  };
  // -> Promise<{ text, thinking, thinkingBlocks, toolCalls }>
  openai.run = function (opts) {
    var s = opts.settings;
    var f = fetchFn();
    if (!f) return Promise.reject(new Error('No HTTP transport available'));
    var url = AI.baseOf(s.baseUrl) + '/v1/chat/completions';
    function attempt(withEffort) {
      var acc = openai.accumulator(opts.onEvent);
      var parser = AI.sseParser();
      function feed(datas) {
        datas.forEach(function (d) {
          if (d === '[DONE]') return;
          var obj = null;
          try { obj = JSON.parse(d); } catch (e) { return; }
          acc.push(obj);
        });
      }
      return f(url, { method: 'POST', headers: openai.headers(s), body: JSON.stringify(openai.body(s, opts.system, opts.conv, withEffort)), signal: opts.signal })
        .then(function (res) {
          if (!res.ok) return readError(res).then(function (err) { throw err; });
          var ct = String(res.headers && res.headers.get ? res.headers.get('content-type') || '' : '');
          if (ct.indexOf('json') !== -1 && ct.indexOf('event-stream') === -1) {
            // non-streaming gateway: one full completion
            return res.json().then(function (j) { acc.push(j); });
          }
          return readStream(res, function (chunk) { feed(parser.push(chunk)); }).then(function () { feed(parser.flush()); });
        })
        .then(function () { return acc.done(); });
    }
    return attempt(true).catch(function (err) {
      // gateways that reject reasoning_effort for the model: retry without it
      if (err && err.status === 400 && /reasoning_effort|reasoning/i.test(err.body || err.message || '')) return attempt(false);
      throw err;
    });
  };
  // GET /v1/models -> [ids]
  openai.models = function (s) {
    var f = fetchFn();
    if (!f) return Promise.reject(new Error('No HTTP transport available'));
    var h = openai.headers(s);
    delete h['Content-Type'];
    h.Accept = 'application/json';
    return f(AI.baseOf(s.baseUrl) + '/v1/models', { method: 'GET', headers: h }).then(function (res) {
      if (!res.ok) return readError(res).then(function (err) { throw err; });
      return res.json();
    }).then(function (j) {
      var arr = Array.isArray(j) ? j : (j.data || j.models || []);
      return arr.map(function (m) { return typeof m === 'string' ? m : (m.id || m.name || ''); }).filter(Boolean).sort();
    });
  };

  // GET /model_group/info (LiteLLM) -> { id: { reasoning } }; resolves to
  // null when the gateway has no such endpoint, so callers keep the defaults
  openai.modelInfo = function (s) {
    var f = fetchFn();
    if (!f) return Promise.resolve(null);
    var h = openai.headers(s);
    delete h['Content-Type'];
    h.Accept = 'application/json';
    return f(AI.baseOf(s.baseUrl).replace(/\/v1$/, '') + '/model_group/info', { method: 'GET', headers: h }).then(function (res) {
      if (!res.ok) return null;
      return res.json().then(function (j) {
        var arr = Array.isArray(j) ? j : (j.data || []);
        var map = {};
        arr.forEach(function (g) {
          var id = g.model_group || g.model_name || g.id;
          if (!id) return;
          var params = g.supported_openai_params || [];
          map[id] = { reasoning: g.supports_reasoning === true || params.indexOf('reasoning_effort') !== -1 };
        });
        return map;
      });
    }).catch(function () { return null; });
  };
  // refresh the per-model facts once per gateway; re-renders the drawer header
  AI.refreshModelInfo = function (force) {
    var s = AI.loadSettings();
    if (s.provider !== 'litellm' || !s.baseUrl) return Promise.resolve();
    if (!force && AI.modelInfoBase === s.baseUrl) return Promise.resolve();
    AI.modelInfoBase = s.baseUrl;
    return openai.modelInfo(s).then(function (map) {
      AI.modelInfo = map;
      if (drawer) rebuildHeader();
    });
  };
  // the model list for the drawer's selector: fetched when nothing is cached
  // for this gateway; resolves to the ids, or null when not applicable or
  // the gateway failed (a toast says so and the typed model stays)
  AI.ensureModels = function (s) {
    if (!s || s.provider !== 'litellm' || !s.baseUrl) return Promise.resolve(null);
    if (AI.modelCache && AI.modelCacheBase === s.baseUrl) return Promise.resolve(AI.modelCache);
    return openai.models(s).then(function (ids) {
      AI.modelCache = ids;
      AI.modelCacheBase = s.baseUrl;
      if (drawer) rebuildHeader();
      return ids;
    }, function (err) {
      var a = app();
      if (a && a.toast) a.toast('Could not list models: ' + err.message, 'err');
      return null;
    });
  };
  // per-model facts for the effort selector: refetch when the map has not
  // heard of this model (a model added to the gateway since the last load)
  AI.ensureModelInfo = function (s) {
    var force = !!(AI.modelInfo && s.model && !AI.modelInfo[s.model]);
    return AI.refreshModelInfo(force);
  };

  // ------------------------------------------------------------ claude -p
  var claude = AI.claude = {};
  claude.args = function (s, session) {
    var a = ['-p', '--input-format', 'stream-json', '--output-format', 'stream-json', '--verbose', '--include-partial-messages',
      '--tools', '', '--strict-mcp-config', '--disable-slash-commands'];
    if (s.claudeModel) a.push('--model', s.claudeModel);
    if (s.effort) a.push('--effort', s.effort);
    if (session) a.push('--resume', session);
    return a;
  };
  // one user turn as a stream-json line (files become content blocks)
  claude.userLine = function (m) {
    var content = [];
    if (m.role === 'tool') {
      content.push({ type: 'text', text: 'Tool results:\n```json\n' + JSON.stringify((m.results || []).map(function (r) { return { name: r.name, result: r.content }; })) + '\n```\nContinue.' });
    } else {
      if (m.text) content.push({ type: 'text', text: m.text });
      (m.files || []).forEach(function (f) {
        if (!f.data) { content.push({ type: 'text', text: '[attachment ' + f.name + ' — content not available]' }); return; }
        if (f.kind === 'image') content.push({ type: 'image', source: { type: 'base64', media_type: f.type, data: f.data } });
        else if (f.kind === 'pdf') content.push({ type: 'document', source: { type: 'base64', media_type: 'application/pdf', data: f.data } });
        else content.push({ type: 'text', text: '--- file: ' + f.name + ' ---\n' + f.data + '\n--- end of ' + f.name + ' ---' });
      });
      if (!content.length) content.push({ type: 'text', text: '(empty)' });
    }
    return JSON.stringify({ type: 'user', message: { role: 'user', content: content } });
  };
  // ```headway-tool fences -> { text (fences removed), calls }
  claude.parseFences = function (text) {
    var calls = [];
    var re = new RegExp('```' + AI.FENCE + '[^\\n]*\\n([\\s\\S]*?)```', 'g');
    var out = String(text || '').replace(re, function (m, body) {
      var obj = null;
      try { obj = JSON.parse(body.trim()); } catch (e) { obj = { name: '', args: { __parseError: e.message, raw: body.trim() } }; }
      if (Array.isArray(obj)) obj.forEach(function (o) { calls.push(o); }); else calls.push(obj);
      return '';
    });
    calls = calls.map(function (c, i) {
      c = isObj(c) ? c : {};
      return { id: 'fence_' + Date.now().toString(36) + '_' + i, name: String(c.name || c.tool || ''), args: isObj(c.args) ? c.args : (isObj(c.arguments) ? c.arguments : (isObj(c.input) ? c.input : {})) };
    });
    return { text: out.replace(/\n{3,}/g, '\n\n').trim(), calls: calls };
  };
  // reduce stdout lines into a turn: { session, text, thinking, done, error }
  claude.reducer = function (onEvent) {
    var st = { session: null, text: '', thinking: '', done: false, error: null, finalText: null };
    return {
      state: st,
      push: function (line) {
        var ev = null;
        try { ev = JSON.parse(line); } catch (e) { return; }
        if (!ev || typeof ev !== 'object') return;
        if (ev.session_id) st.session = ev.session_id;
        if (ev.type === 'stream_event' && ev.event) {
          var e = ev.event;
          if (e.type === 'content_block_delta' && e.delta) {
            if (e.delta.type === 'text_delta' && e.delta.text) { st.text += e.delta.text; if (onEvent) onEvent({ type: 'text', delta: e.delta.text }); }
            else if (e.delta.type === 'thinking_delta' && e.delta.thinking) { st.thinking += e.delta.thinking; if (onEvent) onEvent({ type: 'thinking', delta: e.delta.thinking }); }
          } else if (e.type === 'content_block_start' && e.content_block && e.content_block.type === 'thinking' && onEvent) {
            onEvent({ type: 'thinking', delta: '' });
          }
        } else if (ev.type === 'assistant' && ev.message && Array.isArray(ev.message.content)) {
          // the full message is authoritative for text (partials can be lossy)
          var txt = ev.message.content.filter(function (b) { return b.type === 'text'; }).map(function (b) { return b.text; }).join('');
          if (txt) st.finalText = (st.finalText || '') + txt;
        } else if (ev.type === 'result') {
          st.done = true;
          if (ev.is_error) st.error = (Array.isArray(ev.errors) && ev.errors.join('; ')) || ev.result || ev.subtype || 'Claude returned an error';
          else if (ev.subtype && ev.subtype !== 'success' && !st.text && !st.finalText) st.error = ev.subtype.replace(/_/g, ' ');
        } else if (ev.type === 'system' && ev.subtype === 'init' && ev.model && onEvent) {
          onEvent({ type: 'meta', model: ev.model });
        }
      }
    };
  };

  // one process per conversation; `session` lets a fresh process resume
  var proc = null; // { handle, session, model, effort, reducer, resolve, reject, stderr }
  claude.proc = function () { return proc; };
  claude.stop = function () {
    if (proc && proc.handle && proc.handle.alive) { try { proc.handle.kill(); } catch (e) { /* gone */ } }
    var session = proc ? proc.session : null;
    proc = null;
    return session;
  };
  function claudeBridge() { return root.HeadwayDesktop && root.HeadwayDesktop.claude; }
  // -> Promise<{ text, thinking, toolCalls, session }>
  claude.run = function (opts) {
    var s = opts.settings;
    var bridge = claudeBridge();
    if (!bridge) return Promise.reject(new Error('The Claude subscription provider needs the desktop app.'));
    var lastMsg = opts.conv[opts.conv.length - 1];
    if (!lastMsg) return Promise.reject(new Error('nothing to send'));
    var reuse = proc && proc.handle && proc.handle.alive && proc.model === s.claudeModel && proc.effort === s.effort && !proc.busy;
    var session = opts.session || (proc ? proc.session : null);
    var start = reuse ? Promise.resolve(proc) : bridge.path(s.claudeBin).then(function (bin) {
      if (!bin) throw new Error('Claude Code is not installed (or the path in AI settings is wrong). Install it from claude.com/code, then try again.');
      if (proc) claude.stop();
      var args = claude.args(s, session);
      // --system-prompt goes on the command line; the resume keeps the old one
      args.push('--system-prompt', opts.system);
      var p = { handle: null, session: session, model: s.claudeModel, effort: s.effort, reducer: null, busy: false, stderr: [] };
      return bridge.spawn(bin, args, function (kind, line) {
        if (kind === 'out') { if (p.reducer) p.reducer.push(line); if (p.reducer && p.reducer.state.session) p.session = p.reducer.state.session; if (p.reducer && p.reducer.state.done) settle(p); }
        else if (kind === 'err') { p.stderr.push(line); if (p.stderr.length > 40) p.stderr.shift(); }
        else if (kind === 'exit') { p.exited = true; settle(p, true); }
      }).then(function (h) { p.handle = h; proc = p; return p; });
    });
    function settle(p, exited) {
      if (!p.turn) return;
      var t = p.turn;
      var st = p.reducer.state;
      if (!st.done && !exited) return;
      p.turn = null; p.busy = false;
      if (st.error) { t.reject(new Error(st.error)); return; }
      if (!st.done && exited) {
        var why = p.stderr.filter(Boolean).slice(-3).join(' ');
        t.reject(new Error(why ? 'Claude exited: ' + why : 'Claude exited before finishing'));
        return;
      }
      var text = st.text || st.finalText || '';
      var parsed = claude.parseFences(text);
      t.resolve({ text: parsed.text, thinking: st.thinking, thinkingBlocks: [], toolCalls: parsed.calls, session: p.session });
    }
    return start.then(function (p) {
      return new Promise(function (resolve, reject) {
        p.busy = true;
        p.reducer = claude.reducer(opts.onEvent);
        p.turn = { resolve: resolve, reject: reject };
        if (opts.signal) opts.signal.addEventListener('abort', function () {
          if (p.turn) { p.turn = null; p.busy = false; claude.stop(); reject(new Error('stopped')); }
        });
        p.handle.write(claude.userLine(lastMsg)).catch(function (err) {
          p.turn = null; p.busy = false;
          reject(new Error('could not talk to Claude: ' + (err && err.message || err)));
        });
      });
    }).catch(function (err) {
      // a dead resume target: retry once with a fresh session
      if (session && /resume|session|No conversation found/i.test(err.message || '') && !opts._retried) {
        claude.stop();
        return claude.run(Object.assign({}, opts, { session: null, _retried: true }));
      }
      throw err;
    });
  };

  // ------------------------------------------------------------ files
  function kindOf(file) {
    var t = String(file.type || '');
    if (/^image\/(png|jpeg|gif|webp)$/.test(t)) return 'image';
    if (t === 'application/pdf' || /\.pdf$/i.test(file.name)) return 'pdf';
    return 'text';
  }
  // File -> Promise<{ name, type, kind, size, data }>
  AI.readFile = function (file) {
    var kind = kindOf(file);
    var max = kind === 'text' ? AI.TEXT_FILE_MAX : AI.BIN_FILE_MAX;
    if (file.size > max) return Promise.reject(new Error(file.name + ' is too large (' + Math.round(file.size / 1024) + ' KB; limit ' + Math.round(max / 1024) + ' KB)'));
    return new Promise(function (resolve, reject) {
      var r = new FileReader();
      r.onerror = function () { reject(new Error('could not read ' + file.name)); };
      r.onload = function () {
        var out = { name: file.name, type: file.type || (kind === 'pdf' ? 'application/pdf' : 'text/plain'), kind: kind, size: file.size };
        if (kind === 'text') {
          var txt = String(r.result || '');
          if (/[\x00-\x08\x0E-\x1F]/.test(txt.slice(0, 2000))) { reject(new Error(file.name + ' is not a text, image or PDF file')); return; }
          out.data = txt;
        } else out.data = String(r.result || '').replace(/^data:[^,]*,/, '');
        resolve(out);
      };
      if (kind === 'text') r.readAsText(file); else r.readAsDataURL(file);
    });
  };

  // ------------------------------------------------------------ conversation + agent loop
  var conv = [];          // [{ role: 'user', text, files, t } | { role: 'assistant', text, thinking, toolCalls, toolResults, error, streaming } | { role: 'tool', results }]
  var running = null;     // { controller }
  var claudeSession = null;
  var unread = false;
  var listeners = [];
  AI.conversation = function () { return conv; };
  AI.busy = function () { return !!running; };
  AI.onChange = function (fn) { listeners.push(fn); };
  function emit(kind) { listeners.forEach(function (fn) { try { fn(kind); } catch (e) { /* ui */ } }); }
  function app() { return root.HeadwayApp; }
  function isDesktop() { return !!root.HeadwayDesktop; }

  function persistConv() {
    try {
      var slim = conv.map(function (m) {
        var c = clone(m);
        if (c.files) c.files = c.files.map(function (f) { return { name: f.name, type: f.type, kind: f.kind, size: f.size }; });
        delete c.streaming;
        return c;
      });
      root.localStorage.setItem(AI.CHAT_KEY, JSON.stringify({ conv: slim.slice(-80), session: claudeSession, provider: AI.loadSettings().provider }));
    } catch (e) { /* storage optional */ }
  }
  function restoreConv() {
    try {
      var raw = JSON.parse(root.localStorage.getItem(AI.CHAT_KEY) || 'null');
      if (raw && Array.isArray(raw.conv)) { conv = raw.conv; claudeSession = raw.session || null; }
    } catch (e) { conv = []; }
  }
  AI.newChat = function () {
    if (running) AI.stop();
    claude.stop();
    conv = [];
    claudeSession = null;
    unread = false;
    persistConv();
    emit('conv');
  };
  AI.stop = function () {
    if (!running) return;
    try { running.controller.abort(); } catch (e) { /* no signal */ }
  };

  function ctxNow() {
    var A = app();
    var ui = A ? A.ai.ui() : { view: 'planning', selectedNum: null, desktop: isDesktop(), userName: '' };
    var doc = null;
    if (A && A.ai.hasDoc()) {
      var st = A.ai.state();
      doc = { title: st.meta.title, items: st.items.length, phases: st.phases.length, start: st.meta.timelineStart, end: st.meta.endDate };
    }
    return { today: RM.fmtISO(new Date()), userName: ui.userName, desktop: ui.desktop, view: ui.view, selectedNum: ui.selectedNum, doc: doc };
  }

  // send a user message; resolves when the assistant is done (tools included)
  AI.send = function (text, files) {
    if (running) return Promise.reject(new Error('busy'));
    var s = AI.loadSettings();
    var ready = AI.ready(s, isDesktop());
    if (!ready.ok) return Promise.reject(new Error(ready.why));
    conv.push({ role: 'user', text: String(text || ''), files: files || [], t: Date.now() });
    persistConv();
    emit('conv');
    return runTurns(s, 0);
  };
  function runTurns(s, round) {
    var controller = typeof AbortController !== 'undefined' ? new AbortController() : { signal: null, abort: function () {} };
    running = { controller: controller };
    var asst = { role: 'assistant', text: '', thinking: '', toolCalls: [], streaming: true, t: Date.now(), provider: s.provider, model: s.provider === 'claude' ? s.claudeModel : s.model };
    conv.push(asst);
    emit('conv');
    var textTools = s.provider === 'claude';
    var ctx = ctxNow();
    ctx.textTools = textTools;
    var system = AI.systemPrompt(ctx);
    var onEvent = function (ev) {
      if (ev.type === 'text') asst.text += ev.delta;
      else if (ev.type === 'thinking') asst.thinking += ev.delta;
      else if (ev.type === 'meta' && ev.model) asst.model = ev.model;
      emit('stream');
    };
    var provider = s.provider === 'claude' ? claude : openai;
    var history = conv.slice(0, -1);
    return provider.run({ settings: s, system: system, conv: history, signal: controller.signal, session: claudeSession, onEvent: onEvent })
      .then(function (res) {
        asst.text = res.text != null ? res.text : asst.text;
        asst.thinking = res.thinking || asst.thinking;
        asst.thinkingBlocks = res.thinkingBlocks || [];
        asst.toolCalls = res.toolCalls || [];
        if (res.session) claudeSession = res.session;
        asst.streaming = false;
        if (!asst.toolCalls.length) return finish();
        if (round >= AI.MAX_ROUNDS) { asst.error = 'Stopped after ' + AI.MAX_ROUNDS + ' tool rounds.'; return finish(); }
        var A = app();
        asst.toolResults = [];
        // tools run one after another (writes must land in order); some are async
        return asst.toolCalls.reduce(function (p, c) {
          return p.then(function () {
            return Promise.resolve().then(function () {
              if (c.args && c.args.__parseError) throw new Error('arguments were not valid JSON: ' + c.args.__parseError);
              return AI.runTool(c.name, c.args, A);
            }).then(function (content) { return { ok: true, content: content }; },
              function (e) { return { ok: false, content: { error: (e && e.message) || String(e) } }; })
              .then(function (out) {
                asst.toolResults.push({ id: c.id, name: c.name, args: c.args, ok: out.ok, content: out.content });
                emit('conv');
              });
          });
        }, Promise.resolve()).then(function () {
          var results = asst.toolResults;
          conv.push({ role: 'tool', results: results.map(function (r) { return { id: r.id, name: r.name, content: r.content }; }) });
          persistConv();
          emit('conv');
          running = null;
          return runTurns(s, round + 1);
        });
      }, function (err) {
        asst.streaming = false;
        asst.error = err && err.message === 'stopped' ? 'Stopped.' : ((err && err.message) || String(err));
        return finish();
      });
    function finish() {
      running = null;
      unread = true;
      persistConv();
      emit('done');
      return asst;
    }
  }

  // ------------------------------------------------------------ settings UI
  AI.settingsHtml = function () {
    var s = AI.loadSettings();
    var desktop = isDesktop();
    function inp(id, val, ph, type, extra) {
      return '<input id="' + id + '" type="' + (type || 'text') + '" style="width:100%" value="' + esc(val || '') + '" placeholder="' + esc(ph || '') + '" autocomplete="off"' + (extra || '') + '>';
    }
    var provSeg = '<div class="seg" id="aiProvSeg">' +
      '<button data-aiprov="litellm"' + (s.provider === 'litellm' ? ' class="on"' : '') + '>LiteLLM gateway</button>' +
      '<button data-aiprov="claude"' + (s.provider === 'claude' ? ' class="on"' : '') + '>Claude subscription</button></div>';
    var litellm =
      '<div id="aiLitellm"' + (s.provider === 'litellm' ? '' : ' hidden') + '>' +
      '<div class="m-sec"><label>Gateway URL</label>' + inp('aiBase', s.baseUrl, 'https://litellm.example.com') +
      '<div class="m-hint">The LiteLLM proxy base URL (with or without /v1). Any OpenAI-compatible endpoint works.</div></div>' +
      '<div class="p-grid2">' +
      '<div class="m-sec"><label>API key</label>' + inp('aiKey', s.apiKey, 'sk-…', 'password') + '</div>' +
      '<div class="m-sec"><label>Model</label><div class="p-row">' + inp('aiModel', s.model, 'e.g. claude-sonnet-5', 'text', ' list="aiModelList"') +
      '<button id="aiModels" class="fixed" title="List models from the gateway">Load</button></div><datalist id="aiModelList"></datalist></div>' +
      '</div>' +
      '<div class="m-sec"><label>Extra headers</label><textarea id="aiHeaders" rows="2" style="width:100%;font-family:var(--mono);font-size:12px" placeholder="X-Team: platform">' + esc(s.headers || '') + '</textarea>' +
      '<div class="m-hint">One <code>Name: value</code> per line, sent with every request. Optional.</div></div>' +
      '</div>';
    var claudeUi =
      '<div id="aiClaude"' + (s.provider === 'claude' ? '' : ' hidden') + '>' +
      (desktop ? '' : '<div class="m-hint" style="margin-bottom:12px">This provider runs the Claude Code CLI on your machine, so it only works in the desktop app. In the browser, use a LiteLLM gateway.</div>') +
      '<div class="p-grid2">' +
      '<div class="m-sec"><label>Model</label><select id="aiClaudeModel" style="width:100%">' + AI.CLAUDE_MODELS.map(function (m) {
        return '<option value="' + m[0] + '"' + (s.claudeModel === m[0] ? ' selected' : '') + '>' + m[1] + '</option>';
      }).join('') + '</select></div>' +
      '<div class="m-sec"><label>Claude Code path</label>' + inp('aiClaudeBin', s.claudeBin, 'auto-detect') + '</div>' +
      '</div>' +
      '<div class="m-hint">Uses <code>claude -p</code>, billed to your Claude plan. Leave the path blank to find the CLI automatically.</div>' +
      '<div class="p-row" style="margin-top:8px"><button id="aiClaudeCheck" class="fixed"' + (desktop ? '' : ' disabled') + '>Check</button><span id="aiClaudeOut" class="m-hint" style="margin:0 0 0 10px"></span></div>' +
      '</div>';
    return '<h2>Provider</h2>' +
      '<div class="m-sec">' + provSeg + '</div>' +
      litellm + claudeUi +
      '<div class="m-hint" style="margin-top:22px">Stored on this machine only. The assistant can read and edit the open project and your preferences; every edit is undoable and shows in Version history as “you · AI”.</div>' +
      '<div class="p-row" style="margin-top:12px"><button id="aiOpen" class="primary fixed">Open assistant</button></div>';
  };
  AI.wireSettings = function (host) {
    function $(sel) { return host.querySelector(sel); }
    function save(patch) {
      var s = AI.loadSettings();
      Object.keys(patch).forEach(function (k) { s[k] = patch[k]; });
      AI.saveSettings(s);
      emit('settings');
    }
    host.addEventListener('click', function (e) {
      var pb = e.target.closest('[data-aiprov]');
      if (pb) {
        save({ provider: pb.dataset.aiprov });
        host.querySelectorAll('[data-aiprov]').forEach(function (b) { b.classList.toggle('on', b === pb); });
        $('#aiLitellm').hidden = pb.dataset.aiprov !== 'litellm';
        $('#aiClaude').hidden = pb.dataset.aiprov !== 'claude';
      }
    });
    $('#aiBase').addEventListener('change', function () { save({ baseUrl: $('#aiBase').value.trim() }); });
    $('#aiKey').addEventListener('change', function () { save({ apiKey: $('#aiKey').value.trim() }); });
    $('#aiModel').addEventListener('change', function () { save({ model: $('#aiModel').value.trim() }); });
    $('#aiHeaders').addEventListener('change', function () { save({ headers: $('#aiHeaders').value }); });
    $('#aiClaudeModel').addEventListener('change', function () { save({ claudeModel: $('#aiClaudeModel').value }); });
    $('#aiClaudeBin').addEventListener('change', function () { save({ claudeBin: $('#aiClaudeBin').value.trim() }); });
    $('#aiModels').addEventListener('click', function () {
      save({ baseUrl: $('#aiBase').value.trim(), apiKey: $('#aiKey').value.trim(), headers: $('#aiHeaders').value });
      var s = AI.loadSettings();
      var btn = $('#aiModels');
      if (!s.baseUrl) { app().toast('Set the gateway URL first', 'err'); return; }
      btn.disabled = true; btn.textContent = '…';
      openai.models(s).then(function (ids) {
        $('#aiModelList').innerHTML = ids.map(function (id) { return '<option value="' + esc(id) + '">'; }).join('');
        AI.modelCache = ids;
        AI.modelCacheBase = s.baseUrl;
        AI.refreshModelInfo(true);
        btn.disabled = false; btn.textContent = 'Load';
        app().toast(ids.length + ' model' + (ids.length === 1 ? '' : 's') + ' available — pick one in the Model field');
        $('#aiModel').focus();
      }, function (err) {
        btn.disabled = false; btn.textContent = 'Load';
        app().toast('Could not list models: ' + err.message, 'err');
      });
    });
    $('#aiClaudeCheck').addEventListener('click', function () {
      var out = $('#aiClaudeOut');
      var bridge = claudeBridge();
      if (!bridge) { out.textContent = 'Desktop app only'; return; }
      save({ claudeBin: $('#aiClaudeBin').value.trim() });
      out.textContent = 'Looking…';
      bridge.path(AI.loadSettings().claudeBin).then(function (p) {
        out.textContent = p ? 'Found ' + p : 'Not found — install Claude Code or set the path';
      }, function (err) { out.textContent = 'Failed: ' + err.message; });
    });
    $('#aiOpen').addEventListener('click', function () { AI.open(); });
  };

  // ------------------------------------------------------------ drawer UI
  var drawer = null, ui = { open: false, width: 420 };
  var pending = [];   // attachments queued for the next message
  var renderQueued = false;
  function loadUi() {
    try {
      var raw = JSON.parse(root.localStorage.getItem(AI.UI_KEY) || 'null');
      if (raw) { ui.open = !!raw.open; if (raw.width > 300) ui.width = raw.width; }
    } catch (e) { /* storage optional */ }
  }
  function saveUi() { try { root.localStorage.setItem(AI.UI_KEY, JSON.stringify(ui)); } catch (e) { /* storage optional */ } }

  function toolTitle(r) {
    var a = r.args || {};
    if (!r.ok) return r.name + ' failed';
    if (r.name === 'get_project') return 'Read project' + (a.part && a.part !== 'summary' ? ' · ' + a.part + (a.nums ? ' ' + a.nums.map(function (n) { return '#' + n; }).join(', ') : '') : '');
    if (r.name === 'get_preferences') return 'Read preferences';
    if (r.name === 'add_items') { var c = r.content && r.content.created || []; return 'Added ' + c.map(function (x) { return '#' + x.num + ' ' + x.feature; }).join(', '); }
    if (r.name === 'update_items') { var u = r.content && r.content.updated || []; return 'Updated ' + u.map(function (x) { return x.target + (x.deleted ? ' (deleted)' : ''); }).join(', '); }
    if (r.name === 'update_project') { var ch = r.content && r.content.changes || []; return 'Changed ' + (ch.length === 1 ? ch[0] : ch.length + ' settings'); }
    if (r.name === 'set_preference') return 'Preference · ' + a.key + ' = ' + JSON.stringify(a.value);
    if (r.name === 'sync_jira') {
      var jc = r.content || {};
      if (jc.dryRun) return 'Previewed Jira sync · ' + (jc.counts ? jc.counts.create + ' to create, ' + jc.counts.update + ' to update' : '');
      if (jc.nothingToSync) return 'Jira already in sync';
      return 'Synced Jira · ' + jc.created + ' created, ' + jc.updated + ' updated' + (jc.errors && jc.errors.length ? ', ' + jc.errors.length + ' problem(s)' : '');
    }
    if (r.name === 'navigate') return 'Showed ' + (a.view ? a.view : '') + (a.num != null ? ' #' + a.num : '');
    return r.name;
  }
  function isWrite(name) { return /^(add_items|update_items|update_project|set_preference|sync_jira)$/.test(name); }
  function attachChip(f, i, removable) {
    var ico = f.kind === 'image' ? 'image' : f.kind === 'pdf' ? 'file-text' : 'file';
    return '<span class="ai-file" title="' + esc(f.name) + (f.size ? ' · ' + Math.round(f.size / 1024) + ' KB' : '') + '"><i data-lucide="' + ico + '"></i>' + esc(f.name) +
      (removable ? '<button data-airm="' + i + '" title="Remove"><i data-lucide="x"></i></button>' : '') + '</span>';
  }
  function messageHtml(m, idx) {
    if (m.role === 'user') {
      return '<div class="ai-msg user"><div class="ai-bubble">' + esc(m.text).replace(/\n/g, '<br>') +
        (m.files && m.files.length ? '<div class="ai-files">' + m.files.map(function (f) { return attachChip(f); }).join('') + '</div>' : '') +
        '</div></div>';
    }
    if (m.role !== 'assistant') return '';
    var parts = [];
    if (m.thinking) {
      var open = m.streaming && !m.text;
      parts.push('<details class="ai-think"' + (open ? ' open' : '') + '><summary><i data-lucide="brain"></i>' + (open ? 'Thinking…' : 'Thought') + '</summary><div class="ai-think-body">' + esc(m.thinking || '') + '</div></details>');
    }
    if (m.text) parts.push('<div class="ai-body">' + AI.md(m.text) + (m.streaming ? '<span class="ai-cursor"></span>' : '') + '</div>');
    (m.toolResults || []).forEach(function (r, k) {
      var cls = 'ai-tool' + (r.ok ? (isWrite(r.name) ? ' write' : '') : ' fail');
      parts.push('<details class="' + cls + '"><summary><i data-lucide="' + (r.ok ? (isWrite(r.name) ? 'pencil' : 'eye') : 'circle-alert') + '"></i>' + esc(toolTitle(r)) + '</summary>' +
        '<div class="ai-tool-body"><div class="ai-tool-lbl">Call</div><pre>' + esc(JSON.stringify(r.args || {}, null, 1)) + '</pre>' +
        '<div class="ai-tool-lbl">Result</div><pre>' + esc(truncate(JSON.stringify(r.content, null, 1), 6000)) + '</pre></div></details>');
    });
    if (m.error) parts.push('<div class="ai-err"><i data-lucide="circle-alert"></i>' + esc(m.error) + '</div>');
    if (m.streaming && !m.thinking && !m.text) parts.push('<div class="ai-wait"><span></span><span></span><span></span></div>');
    return '<div class="ai-msg asst" data-idx="' + idx + '">' + parts.join('') + '</div>';
  }
  function truncate(s, n) { s = String(s == null ? '' : s); return s.length > n ? s.slice(0, n) + '\n… (' + (s.length - n) + ' more characters)' : s; }

  function drawerHtml() {
    var s = AI.loadSettings();
    var modelOpts = '';
    if (s.provider === 'claude') {
      modelOpts = AI.CLAUDE_MODELS.map(function (m) { return '<option value="' + m[0] + '"' + (s.claudeModel === m[0] ? ' selected' : '') + '>' + m[1] + '</option>'; }).join('');
    } else {
      // the cache belongs to one base URL: ignore it when the endpoint changed
      var ids = (AI.modelCache && AI.modelCacheBase === s.baseUrl ? AI.modelCache : []).slice();
      if (s.model && ids.indexOf(s.model) === -1) ids.unshift(s.model);
      modelOpts = ids.length ? ids.map(function (id) { return '<option value="' + esc(id) + '"' + (s.model === id ? ' selected' : '') + '>' + esc(AI.shortModel(id)) + '</option>'; }).join('') : '<option value="">No model</option>';
    }
    var efforts = AI.effortsFor(s);
    var curModel = s.provider === 'claude' ? s.claudeModel : s.model;
    var effortSel = efforts.length
      ? '<select id="aiEffortSel" title="Effort">' + efforts.map(function (e) { return '<option value="' + e[0] + '"' + (s.effort === e[0] ? ' selected' : '') + '>' + e[1] + '</option>'; }).join('') + '</select>'
      : '';
    return '<div id="aiRz"></div>' +
      '<div class="ai-head">' +
      '<span class="ai-title"><i data-lucide="sparkles"></i>Assistant</span>' +
      '<button id="aiNew" class="ai-ib" title="New chat"><i data-lucide="square-pen"></i></button>' +
      '<button id="aiSettings" class="ai-ib" title="AI settings"><i data-lucide="settings-2"></i></button>' +
      '<button id="aiClose" class="ai-ib" title="Close  ⌘J"><i data-lucide="x"></i></button>' +
      '</div>' +
      '<div class="ai-msgs" id="aiMsgs"></div>' +
      '<div class="ai-compose" id="aiCompose">' +
      '<div class="ai-pending" id="aiPending"></div>' +
      '<textarea id="aiInput" rows="2" placeholder="Ask about the plan, or tell me what to change… (Shift+Enter for a new line)"></textarea>' +
      '<div class="ai-actions">' +
      '<button id="aiAttach" class="ai-ib" title="Attach files"><i data-lucide="paperclip"></i></button>' +
      '<select id="aiModelSel" title="' + esc(curModel || 'Model') + '">' + modelOpts + '</select>' +
      effortSel +
      '<input type="file" id="aiFile" multiple hidden accept="image/png,image/jpeg,image/gif,image/webp,application/pdf,text/*,.md,.csv,.json,.txt">' +
      '<span class="ai-hint" id="aiHint"></span>' +
      '<button id="aiSend" class="primary" title="Send  ↵"><i data-lucide="arrow-up"></i></button>' +
      '<button id="aiStop" class="danger" title="Stop" hidden><i data-lucide="square"></i></button>' +
      '</div></div>';
  }
  function icons() { if (root.lucide) root.lucide.createIcons(); }

  function renderMessages() {
    if (!drawer) return;
    var host = drawer.querySelector('#aiMsgs');
    var nearBottom = host.scrollHeight - host.scrollTop - host.clientHeight < 80;
    var s = AI.loadSettings();
    var ready = AI.ready(s, isDesktop());
    var html = '';
    if (!conv.length) {
      html = '<div class="ai-empty"><i data-lucide="sparkles"></i><div><b>Ask anything about this plan.</b><br>Try “what’s at risk this quarter?”, “add a story to #12 for error handling”, or “move Phase 2 out two weeks”.</div>' +
        (ready.ok ? '' : '<div class="ai-err"><i data-lucide="circle-alert"></i>' + esc(ready.why) + ' <a href="#" id="aiGoSettings">Open AI settings</a></div>') + '</div>';
    } else {
      html = conv.map(messageHtml).join('');
    }
    host.innerHTML = html;
    icons();
    if (nearBottom) host.scrollTop = host.scrollHeight;
    var busy = !!running;
    drawer.querySelector('#aiSend').hidden = busy;
    drawer.querySelector('#aiStop').hidden = !busy;
    drawer.querySelector('#aiInput').disabled = false;
    var hint = drawer.querySelector('#aiHint');
    hint.textContent = busy ? 'Working…' : '';
    var btn = root.document.getElementById('btnAI');
    if (btn) {
      btn.classList.toggle('busy', busy);
      var dot = btn.querySelector('.ai-dot');
      if (dot) dot.hidden = !(unread && !ui.open);
    }
  }
  function renderPending() {
    if (!drawer) return;
    var host = drawer.querySelector('#aiPending');
    host.innerHTML = pending.map(function (f, i) { return attachChip(f, i, true); }).join('');
    host.hidden = !pending.length;
    icons();
  }
  function scheduleRender() {
    if (renderQueued) return;
    renderQueued = true;
    var raf = root.requestAnimationFrame || function (fn) { setTimeout(fn, 16); };
    raf(function () { renderQueued = false; renderMessages(); });
  }
  // the composer grows with its text (two lines at rest, up to ~40% of the drawer)
  function autosize(ta) {
    var cap = Math.max(120, Math.round((drawer ? drawer.clientHeight : 800) * 0.4));
    ta.style.height = 'auto';
    ta.style.height = Math.min(cap, Math.max(0, ta.scrollHeight)) + 'px';
    ta.style.overflowY = ta.scrollHeight > cap ? 'auto' : 'hidden';
  }
  function addFiles(files) {
    var list = Array.prototype.slice.call(files || []);
    if (!list.length) return;
    Promise.all(list.map(function (f) { return AI.readFile(f).then(function (r) { return r; }, function (err) { app().toast(err.message, 'err'); return null; }); }))
      .then(function (res) {
        res.filter(Boolean).forEach(function (r) { pending.push(r); });
        renderPending();
      });
  }
  function submit() {
    var ta = drawer.querySelector('#aiInput');
    var text = ta.value.trim();
    if (!text && !pending.length) return;
    if (running) return;
    var s = AI.loadSettings();
    var ready = AI.ready(s, isDesktop());
    if (!ready.ok) { app().toast(ready.why, 'err'); return; }
    var files = pending.slice();
    pending = [];
    ta.value = '';
    autosize(ta);
    renderPending();
    unread = false;
    AI.send(text, files).catch(function (err) { app().toast(err.message, 'err'); });
  }
  function wireDrawer() {
    var d = drawer;
    var ta = d.querySelector('#aiInput');
    d.querySelector('#aiClose').addEventListener('click', function () { AI.close(); });
    d.querySelector('#aiNew').addEventListener('click', function () { AI.newChat(); ta.focus(); });
    d.querySelector('#aiSettings').addEventListener('click', function () { app().ai.openSettings(); });
    d.querySelector('#aiSend').addEventListener('click', submit);
    d.querySelector('#aiStop').addEventListener('click', function () { AI.stop(); });
    d.querySelector('#aiAttach').addEventListener('click', function () { d.querySelector('#aiFile').click(); });
    d.querySelector('#aiFile').addEventListener('change', function (e) { addFiles(e.target.files); e.target.value = ''; });
    var effSel = d.querySelector('#aiEffortSel');
    if (effSel) effSel.addEventListener('change', function (e) {
      var s = AI.loadSettings(); s.effort = e.target.value; AI.saveSettings(s);
    });
    d.querySelector('#aiModelSel').addEventListener('change', function (e) {
      var s = AI.loadSettings();
      if (s.provider === 'claude') s.claudeModel = e.target.value; else s.model = e.target.value;
      // the effort list follows the model: keep a still-valid pick, else fall back
      s.effort = AI.pickEffort(s);
      AI.saveSettings(s);
      rebuildHeader();
      if (s.provider !== 'litellm') return;
      AI.ensureModelInfo(s).then(function () {
        var s2 = AI.loadSettings();
        var e2 = AI.pickEffort(s2);
        if (e2 !== s2.effort) { s2.effort = e2; AI.saveSettings(s2); }
        rebuildHeader();
      });
    });
    ta.addEventListener('input', function () { autosize(ta); });
    ta.addEventListener('keydown', function (e) {
      if (e.key === 'Enter' && !e.shiftKey && !e.isComposing) { e.preventDefault(); submit(); }
    });
    ta.addEventListener('paste', function (e) {
      var items = e.clipboardData && e.clipboardData.files;
      if (items && items.length) { e.preventDefault(); addFiles(items); }
    });
    d.addEventListener('dragover', function (e) { e.preventDefault(); d.classList.add('drop'); });
    d.addEventListener('dragleave', function () { d.classList.remove('drop'); });
    d.addEventListener('drop', function (e) { e.preventDefault(); d.classList.remove('drop'); if (e.dataTransfer) addFiles(e.dataTransfer.files); });
    d.addEventListener('click', function (e) {
      var rm = e.target.closest('[data-airm]');
      if (rm) { pending.splice(+rm.dataset.airm, 1); renderPending(); return; }
      var ref = e.target.closest('.ai-ref');
      if (ref) { e.preventDefault(); if (!app().ai.selectNum(+ref.dataset.num)) app().toast('No feature #' + ref.dataset.num, 'err'); return; }
      if (e.target.id === 'aiGoSettings') { e.preventDefault(); app().ai.openSettings(); }
    });
    // resize from the left edge
    var rz = d.querySelector('#aiRz');
    rz.addEventListener('pointerdown', function (e) {
      e.preventDefault();
      var x0 = e.clientX, w0 = ui.width;
      function mv(ev) { ui.width = Math.max(320, Math.min(900, w0 + (x0 - ev.clientX))); d.style.setProperty('--ai-w', ui.width + 'px'); }
      function up() { root.removeEventListener('pointermove', mv); root.removeEventListener('pointerup', up); saveUi(); }
      root.addEventListener('pointermove', mv);
      root.addEventListener('pointerup', up);
    });
  }
  function ensureDrawer() {
    if (drawer) return drawer;
    drawer = root.document.getElementById('aiDrawer');
    if (!drawer) return null;
    drawer.innerHTML = drawerHtml();
    drawer.style.setProperty('--ai-w', ui.width + 'px');
    wireDrawer();
    autosize(drawer.querySelector('#aiInput'));
    renderMessages();
    renderPending();
    icons();
    return drawer;
  }
  function rebuildHeader() {
    if (!drawer) return;
    var ta = drawer.querySelector('#aiInput');
    var keep = ta ? ta.value : '';
    drawer.innerHTML = drawerHtml();
    wireDrawer();
    if (keep) { drawer.querySelector('#aiInput').value = keep; }
    autosize(drawer.querySelector('#aiInput'));
    renderMessages();
    renderPending();
  }
  AI.open = function () {
    if (!ensureDrawer()) return;
    ui.open = true;
    unread = false;
    saveUi();
    drawer.hidden = false;
    root.document.body.classList.add('ai-open');
    renderMessages();
    AI.refreshModelInfo();
    AI.ensureModels(AI.loadSettings());
    var ta = drawer.querySelector('#aiInput');
    if (ta) ta.focus();
  };
  AI.close = function () {
    if (!drawer) return;
    ui.open = false;
    saveUi();
    drawer.hidden = true;
    root.document.body.classList.remove('ai-open');
    renderMessages();
  };
  AI.toggle = function () { if (ui.open && drawer && !drawer.hidden) AI.close(); else AI.open(); };
  AI.isOpen = function () { return ui.open && !!drawer && !drawer.hidden; };

  function boot() {
    if (!root.document) return;
    loadUi();
    restoreConv();
    AI.onChange(function (kind) {
      if (kind === 'settings') rebuildHeader();
      else if (kind === 'done') {
        scheduleRender();
        if (!AI.isOpen()) {
          var last = conv[conv.length - 1];
          if (last && last.role === 'assistant' && app()) app().toast(last.error ? 'AI: ' + last.error : 'AI reply ready', last.error ? 'err' : '');
        }
      } else scheduleRender();
    });
    var btn = root.document.getElementById('btnAI');
    if (btn) btn.addEventListener('click', function () { AI.toggle(); });
    root.addEventListener('keydown', function (e) {
      if ((e.metaKey || e.ctrlKey) && !e.shiftKey && !e.altKey && e.key.toLowerCase() === 'j') { e.preventDefault(); AI.toggle(); }
    });
    if (ui.open) {
      // defer: app.js builds the page after this script runs
      setTimeout(function () { if (app() && app().ai.hasDoc()) AI.open(); }, 0);
    }
  }
  if (root.document && root.document.getElementById) {
    // the script sits at the end of <body>, so the toolbar is already parsed
    if (root.document.getElementById('btnAI') || root.document.readyState !== 'loading') boot();
    else root.document.addEventListener('DOMContentLoaded', boot);
  }

  root.HeadwayAI = AI;
  if (typeof module !== 'undefined' && module.exports) module.exports = AI;
})(typeof window !== 'undefined' ? window : globalThis);
