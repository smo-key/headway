/*
 * Headway core — pure logic, no DOM. Loaded in the browser as window.RM and
 * in node (tests) via require. Time is measured in WORKING DAYS from
 * meta.timelineStart (a Monday): 5 per week, weekends don't exist in the index
 * space. week = floor(day / slotsPerWeek) — the slot count follows the
 * selected working days (1-7).
 */
(function (root) {
  'use strict';

  var RM = {};

  // Working days per size — measured in weeks: XS 2d · S 1w · M 2w · L 4w · XL 8w.
  RM.DEFAULT_SIZE_DAYS = { XS: 2, S: 5, M: 10, L: 20, XL: 40 };
  RM.LEGACY_SIZE_DAYS = { XS: 2, S: 3, M: 5, L: 10, XL: 20 };
  RM.SIZE_ORDER = ['XS', 'S', 'M', 'L', 'XL'];

  // Sizing approaches (meta.sizeScheme). Every option maps to working days so
  // scheduling works the same under any approach; 'none' turns sizing off
  // (Kanban / #NoEstimates style — duration is set directly, if at all).
  // Editing options in Setup flips the scheme to 'custom'.
  RM.SIZE_SCHEMES = {
    tshirt: {
      name: 'T-shirt sizes',
      hint: 'XS–XL relative buckets — quick gut-feel estimates',
      sizes: ['XS', 'S', 'M', 'L', 'XL'],
      days: { XS: 2, S: 5, M: 10, L: 20, XL: 40 }
    },
    fibonacci: {
      name: 'Story points',
      hint: 'Fibonacci scale (Scrum) — uncertainty grows with size',
      sizes: ['1', '2', '3', '5', '8', '13'],
      days: { 1: 1, 2: 2, 3: 3, 5: 5, 8: 10, 13: 20 }
    },
    points5: {
      name: 'Points 1–5',
      hint: 'Simple five-step scale',
      sizes: ['1', '2', '3', '4', '5'],
      days: { 1: 2, 2: 5, 3: 10, 4: 20, 5: 40 }
    },
    none: {
      name: 'No sizing',
      hint: 'Kanban / no-estimates — set durations directly when needed',
      sizes: [],
      days: {}
    },
    custom: {
      name: 'Custom',
      hint: 'Your own options and day values',
      sizes: null, // whatever meta.sizeOrder holds
      days: {}
    }
  };
  RM.SIZE_SCHEME_ORDER = ['tshirt', 'fibonacci', 'points5', 'none'];
  // Features and stories size on SEPARATE scales: features under
  // meta.sizeScheme / sizeOrder / sizeDays, stories under the story* twins.
  // Every helper takes an optional kind ('feature' default | 'story').
  RM.DEFAULT_STORY_SIZE_SCHEME = 'fibonacci';
  RM.DEFAULT_STORY_PRIORITY_SCHEME = 'levels';
  function sizeKeys(kind) {
    return kind === 'story'
      ? { scheme: 'storySizeScheme', order: 'storySizeOrder', days: 'storySizeDays' }
      : { scheme: 'sizeScheme', order: 'sizeOrder', days: 'sizeDays' };
  }
  RM.sizeKeys = sizeKeys;
  // walk every sized thing of a kind (features, or every story)
  function eachOfKind(state, kind, fn) {
    (state.items || []).forEach(function (it) {
      if (kind === 'story') (it.stories || []).forEach(fn);
      else fn(it);
    });
  }
  RM.sizeOrderOf = function (state, kind) {
    var m = state.meta || state;
    var k = sizeKeys(kind);
    return m[k.order] || (kind === 'story' ? RM.SIZE_SCHEMES[RM.DEFAULT_STORY_SIZE_SCHEME].sizes : RM.SIZE_ORDER);
  };
  RM.sizingEnabled = function (state, kind) {
    var m = state.meta || state;
    return m[sizeKeys(kind).scheme] !== 'none' && RM.sizeOrderOf(state, kind).length > 0;
  };
  RM.setSizeScheme = function (state, scheme, kind) {
    var def = RM.SIZE_SCHEMES[scheme];
    if (!def || scheme === 'custom') return;
    var m = state.meta, k = sizeKeys(kind);
    m[k.scheme] = scheme;
    m[k.order] = def.sizes.slice();
    m[k.days] = RM.clone(def.days);
  };
  RM.renameSizeOption = function (state, oldLabel, newLabel, kind) {
    var m = state.meta, k = sizeKeys(kind);
    if (!newLabel || oldLabel === newLabel || m[k.order].indexOf(newLabel) !== -1) return;
    m[k.order] = m[k.order].map(function (l) { return l === oldLabel ? newLabel : l; });
    m[k.days][newLabel] = m[k.days][oldLabel];
    delete m[k.days][oldLabel];
    eachOfKind(state, kind, function (o) { if (o.size === oldLabel) o.size = newLabel; });
    m[k.scheme] = 'custom';
  };
  RM.addSizeOption = function (state, label, days, kind) {
    var m = state.meta, k = sizeKeys(kind);
    if (!label || m[k.order].indexOf(label) !== -1) return;
    m[k.order].push(label);
    m[k.days][label] = isFinite(+days) && +days > 0 ? +days : 5;
    m[k.scheme] = 'custom';
  };
  RM.removeSizeOption = function (state, label, kind) {
    var m = state.meta, k = sizeKeys(kind);
    m[k.order] = m[k.order].filter(function (l) { return l !== label; });
    delete m[k.days][label];
    eachOfKind(state, kind, function (o) { if (o.size === label) o.size = null; });
    m[k.scheme] = 'custom';
  };
  RM.RISK_ORDER = ['L', 'M', 'H']; // low / medium / high (severity, not a size)

  // Assessment ("Risk") column schemes. Most projects track nothing here —
  // 'none' is the default for new documents. The alternatives mirror common
  // PM practice: RAID-log style risk severity (probability × impact rolled
  // into L/M/H), estimation confidence (planning-poker style H/M/L),
  // MoSCoW prioritization (Must/Should/Could/Won't), and an auto-computed
  // dependency risk read straight off the graph (cycles, fan-in, slack).
  RM.RISK_SCHEMES = {
    none: { name: 'None', label: 'Risk', desc: 'No assessment column.' },
    risk: { name: 'Risk (manual)', label: 'Risk', order: ['L', 'M', 'H'],
      desc: 'RAID-style severity — probability × impact rolled into Low / Medium / High.' },
    auto: { name: 'Risk (auto)', label: 'Risk', auto: true,
      desc: 'Computed from the dependency graph: cycles, unscheduled or many dependencies, zero slack, deep chains.' },
    confidence: { name: 'Confidence', label: 'Confidence', order: ['H', 'M', 'L'],
      desc: 'Estimation confidence, planning-poker style — High / Medium / Low.' }
  };
  RM.RISK_SCHEME_ORDER = ['none', 'risk', 'auto', 'confidence'];

  // the header apps, in tab order: [key, name, icon, what it does]
  RM.APPS = [
    ['scoping', 'Scoping', 'table-properties', 'Spreadsheet of descriptions and scope'],
    ['prio', 'Prioritizing', 'square-kanban', 'Kanban of features and stories'],
    ['planning', 'Planning', 'chart-gantt', 'Timeline, dependencies and capacity (always on)'],
    ['sprints', 'Sprinting', 'calendar-range', 'Sprint-by-sprint list'],
    ['budget', 'Budgeting', 'wallet', 'Rates, costs and role hours'],
    ['reports', 'Reporting', 'chart-pie', 'Project reporting dashboard']
  ];
  RM.appEnabled = function (state, key) {
    var a = state && state.meta && state.meta.apps;
    if (key === 'planning') return true;
    if (!RM.APPS.some(function (x) { return x[0] === key; })) return true; // not an app (setup, history)
    return !a || a[key] !== false;
  };

  // Priority is its own column (risk measures uncertainty; priority ranks
  // importance): MoSCoW or Critical/High/Medium/Low ladders.
  RM.PRIORITY_SCHEMES = {
    none: { name: 'None', label: 'Priority', desc: 'No priority column.' },
    moscow: { name: 'MoSCoW', label: 'Priority', order: ['M', 'S', 'C', 'W'],
      desc: 'Must / Should / Could / Won’t — classic scope-negotiation priority.' },
    levels: { name: 'Critical / High / Medium / Low', label: 'Priority', order: ['C', 'H', 'M', 'L'],
      desc: 'A severity ladder — Critical, High, Medium, Low.' },
    // computed from the item's four RICE inputs; there is no picked value
    rice: { name: 'RICE score', label: 'RICE', computed: true,
      desc: 'Reach × Impact × Confidence ÷ Effort — evidence-based scoring.' }
  };
  RM.PRIORITY_SCHEME_ORDER = ['none', 'moscow', 'levels', 'rice'];
  // stories carry no RICE inputs, so their scheme list stops at the ladders
  RM.STORY_PRIORITY_SCHEME_ORDER = ['none', 'moscow', 'levels'];
  function prioKey(kind) { return kind === 'story' ? 'storyPriorityScheme' : 'priorityScheme'; }
  RM.prioritySchemeOf = function (state, kind) {
    var s = state && state.meta && state.meta[prioKey(kind)];
    if (kind === 'story' && s === 'rice') return 'none';
    return RM.PRIORITY_SCHEMES[s] ? s : 'none';
  };
  RM.priorityEnabled = function (state, kind) { return RM.prioritySchemeOf(state, kind) !== 'none'; };
  RM.priorityOrderOf = function (state, kind) {
    return (RM.PRIORITY_SCHEMES[RM.prioritySchemeOf(state, kind)].order || []).slice();
  };
  RM.setPriorityScheme = function (state, key, kind) {
    if (!RM.PRIORITY_SCHEMES[key]) return;
    if (kind === 'story' && key === 'rice') return;
    state.meta[prioKey(kind)] = key;
    var order = RM.PRIORITY_SCHEMES[key].order || [];
    eachOfKind(state, kind, function (o) {
      if (o.priority && order.indexOf(o.priority) === -1) o.priority = null;
    });
  };
  RM.riskSchemeOf = function (state) {
    var s = state && state.meta && state.meta.riskScheme;
    return RM.RISK_SCHEMES[s] ? s : 'none';
  };
  RM.riskEnabled = function (state) { return RM.riskSchemeOf(state) !== 'none'; };
  RM.riskOrderOf = function (state) {
    return (RM.RISK_SCHEMES[RM.riskSchemeOf(state)].order || []).slice();
  };
  RM.riskColLabel = function (state) {
    return RM.RISK_SCHEMES[RM.riskSchemeOf(state)].label;
  };
  RM.setRiskScheme = function (state, key) {
    if (!RM.RISK_SCHEMES[key]) return;
    state.meta.riskScheme = key;
    var order = RM.RISK_SCHEMES[key].order || [];
    state.items.forEach(function (it) {
      if (it.risk && order.indexOf(it.risk) === -1) it.risk = null;
      (it.stories || []).forEach(function (st) {
        if (st.risk && order.indexOf(st.risk) === -1) st.risk = null;
      });
    });
  };

  // RICE: reach × impact × confidence% ÷ effort. Null until every input is
  // in — a partial score would sort above honestly-unscored work.
  RM.riceScore = function (it) {
    var r = it && it.rice;
    if (!r || r.reach == null || r.impact == null || r.confidence == null || !r.effort) return null;
    return r.reach * r.impact * (r.confidence / 100) / r.effort;
  };

  // weeks [w0, w1) of a numbered sprint
  RM.sprintRange = function (meta, num) {
    var si = RM.sprintInfo(meta);
    var w0 = si.anchorWeek + (num - si.firstNum) * si.wps;
    return { w0: w0, w1: w0 + si.wps };
  };
  RM.itemInWeeks = function (meta, it, w0, w1) {
    if (it.startDay == null) return false;
    var S = RM.slotsOf(meta);
    var span = Math.max(1, (it.durDays || 0) + (it.riskDays || 0));
    return it.startDay < w1 * S && it.startDay + span > w0 * S;
  };

  RM.DEFAULT_TEAM_TYPES = ['Software Engineer', 'Product Designer', 'Product Manager', 'Data Scientist', 'QA Engineer'];
  RM.WEEK_HOURS = 40; // one person's full week
  RM.HISTORY_MAX = 300; // version-history entries kept per document
  RM.OPTIONS_MAX = 12;  // parked alternate-plan options kept per document
  RM.HISTORY_OPS_MAX = 120; // change-detail rows kept per history entry
  RM.HISTORY_COALESCE_MS = 5 * 60 * 1000; // same-label edits by one person inside this window merge into one entry
  RM.ANY_TYPE = '';

  // Scoping-view columns. Built-ins map to fixed item fields; custom columns
  // ('c…' keys) store their text in item.custom.
  RM.SCOPE_BUILTIN_LABELS = {
    description: 'Description',
    ac: 'Acceptance criteria',
    enables: 'Enables',
    outOfScope: 'Out of scope',
    extDeps: 'External dependencies',
    notes: 'Notes'
  };
  // New documents start with Description and Acceptance criteria; the rest
  // stay available in the add-column menu. Legacy docs infer their list from
  // actual content.
  RM.DEFAULT_SCOPE_COLS = ['description', 'ac'];
  RM.SCOPE_BUILTIN_ORDER = ['description', 'ac', 'enables', 'outOfScope', 'extDeps', 'notes'];
  // built-ins that restrict to one row kind unless the user says otherwise
  RM.SCOPE_BUILTIN_SCOPE = { ac: 'story' };
  // fixed (chip) scoping columns and the canonical full-order template
  // milestone marker shapes; the first is the default
  RM.MS_STYLES = ['diamond', 'star', 'circle'];
  RM.msStyleOf = function (it) {
    return RM.MS_STYLES.indexOf(it && it.msStyle) > 0 ? it.msStyle : 'diamond';
  };
  RM.SCOPE_FIXED_KEYS = ['assignees', 'size', 'risk', 'priority', 'duration', 'start', 'deadline', 'workstream', 'epic'];
  RM.SCOPE_DEFAULT_ORDER = ['description', 'ac', 'epic', 'assignees', 'size', 'risk', 'priority', 'duration', 'start', 'deadline', 'workstream'];

  // 2026 US holiday calendar (company observance table). Merged once into a
  // document's holidays (meta.holidaysV2026 flags the merge so user deletions
  // stick). Weekend-dated entries would be ignored by holidayDaySet anyway.
  RM.US_HOLIDAYS_2026 = [
    '2026-01-01', '2026-01-02', '2026-01-19', '2026-02-16',
    '2026-05-22', '2026-05-25', '2026-06-19', '2026-07-03',
    '2026-09-04', '2026-09-07',
    '2026-11-25', '2026-11-26', '2026-11-27',
    '2026-12-24', '2026-12-25', '2027-01-01'
  ];

  // Categorical bar palette (CVD-validated): product blue / data orange /
  // process green / mixed plum. `neutral` is the gray every workstream
  // falls back to until a color is chosen.
  RM.PALETTE = {
    product: '3273BD',
    data: 'C25E0E',
    process: '08875B',
    mixed: 'A14FBF',
    neutral: '6E7883'
  };
  RM.PALETTE_KEYS = ['product', 'data', 'process', 'mixed'];

  // Color follows the WORKSTREAM (state.wsColors[ws], a palette key or 6-hex);
  // epics carry an ICON instead (state.epicIcons[epic], a lucide icon name).
  // Well-known workstreams get a default color on load (OS = blue).
  RM.DEFAULT_WS_COLORS = {
    'OS': '3273BD',
    'Product': 'A14FBF',
    'Data': 'C25E0E',
    'Process': '08875B',
    'Product / Process': '2A7F8E',
    'All': '6E7883'
  };
  function resolveColor(custom) {
    if (custom) {
      if (RM.PALETTE[custom]) return RM.PALETTE[custom];
      var hex = String(custom).replace(/^#/, '').toUpperCase();
      if (/^[0-9A-F]{6}$/.test(hex)) return hex;
    }
    return null;
  }
  // The DEFAULT workstream: what a null/empty workstream means. It is a
  // real, renamable workstream with its own color (gray unless changed).
  RM.defaultWsName = function (state) {
    var m = state && state.meta;
    return (m && typeof m.defaultWsName === 'string' && m.defaultWsName.trim()) || 'General';
  };
  RM.defaultWsColor = function (state) {
    var m = state && state.meta;
    return (m && resolveColor(m.defaultWsColor)) || RM.PALETTE.neutral;
  };
  RM.colorForWs = function (state, ws) {
    if (!ws) return RM.defaultWsColor(state);
    var c = state && state.wsColors ? resolveColor(state.wsColors[ws]) : null;
    return c || RM.PALETTE.neutral;
  };
  // What bar colors follow. The mode is a UI preference the app sets on
  // load; every renderer and export reads colorForItem, so they all agree.
  RM.COLOR_MODES = ['workstream', 'epic', 'assignee', 'priority'];
  var colorMode = 'workstream';
  RM.setColorMode = function (mode) { colorMode = RM.COLOR_MODES.indexOf(mode) !== -1 ? mode : 'workstream'; };
  RM.colorMode = function () { return colorMode; };
  // a spread of distinct hues for names without a chosen color (epics)
  RM.HASH_PALETTE = ['3273BD', 'C25E0E', '08875B', 'A14FBF', '2A7F8E', 'B8336A', '5B6ABF', '8A7B1E', 'C2402E', '3E8E41'];
  function hashOf(s) {
    var h = 0;
    s = String(s || '');
    for (var i = 0; i < s.length; i++) h = (h * 31 + s.charCodeAt(i)) >>> 0;
    return h;
  }
  function hslToHex(h, s, l) {
    s /= 100; l /= 100;
    var k = function (n) { return (n + h / 30) % 12; };
    var a = s * Math.min(l, 1 - l);
    var f = function (n) { return l - a * Math.max(-1, Math.min(k(n) - 3, Math.min(9 - k(n), 1))); };
    return [f(0), f(8), f(4)].map(function (v) { return ('0' + Math.round(v * 255).toString(16)).slice(-2); }).join('').toUpperCase();
  }
  // Epics without a chosen color share the hash palette; the document's
  // epics are assigned together (alphabetically) so two epics never land on
  // the same swatch while a free one exists, and adding an epic does not
  // recolor the others unless it collides.
  function epicColorTable(state) {
    var names = {};
    ((state && state.items) || []).forEach(function (it) { if (it.epic) names[it.epic] = true; });
    var chosen = state && state.epicColors ? state.epicColors : {};
    var taken = {}, table = {};
    Object.keys(names).sort().forEach(function (e) {
      var c = resolveColor(chosen[e]);
      if (c) { table[e] = c; taken[c] = true; }
    });
    Object.keys(names).sort().forEach(function (e) {
      if (table[e]) return;
      var n = RM.HASH_PALETTE.length;
      var i = ((hashOf(e) * 2654435761) >>> 0) % n;
      for (var k = 0; k < n; k++) {
        var col = RM.HASH_PALETTE[(i + k) % n];
        if (!taken[col]) { table[e] = col; taken[col] = true; return; }
      }
      table[e] = RM.HASH_PALETTE[i]; // more epics than swatches: colors repeat
    });
    return table;
  }
  RM.colorForEpic = function (state, epic) {
    if (!epic) return RM.PALETTE.neutral;
    var c = state && state.epicColors ? resolveColor(state.epicColors[epic]) : null;
    if (c) return c;
    var table = epicColorTable(state);
    return table[epic] || RM.HASH_PALETTE[((hashOf(epic) * 2654435761) >>> 0) % RM.HASH_PALETTE.length];
  };
  RM.colorForMember = function (m) {
    if (!m) return RM.PALETTE.neutral;
    return hslToHex(hashOf(RM.memberLabel(m)) % 360, 52, 42);
  };
  // priority: a fixed ladder from hottest to coolest across the scheme's order
  // (critical = bright red, high = orange, medium = green, low = gray)
  RM.PRIORITY_RAMP = ['E0261B', 'EF8A1F', '2E9E52', '8A929B'];
  RM.PRIORITY_TIERS = ['crit', 'high', 'med', 'low'];
  // which of the four tiers a picked value sits in — Must/Critical are 'crit',
  // Won't/Low are 'low'; null for no value or a scheme without a ladder
  RM.priorityTier = function (state, value, kind) {
    if (!value) return null;
    var order = RM.priorityOrderOf(state, kind);
    var idx = order.indexOf(value);
    if (idx === -1) return null;
    return RM.PRIORITY_TIERS[Math.min(3, Math.floor(idx * 4 / order.length))];
  };
  RM.colorForPriority = function (state, it) {
    var sch = RM.prioritySchemeOf(state);
    if (sch === 'none') return RM.PALETTE.neutral;
    if (sch === 'rice') {
      // relative: quartiles of the scored items
      var mine = RM.riceScore(it);
      if (!(mine > 0)) return RM.PALETTE.neutral;
      var scores = state.items.map(RM.riceScore).filter(function (x) { return x > 0; }).sort(function (a, b) { return b - a; });
      var rank = scores.indexOf(mine);
      return RM.PRIORITY_RAMP[Math.min(3, Math.floor(rank * 4 / scores.length))];
    }
    var order = RM.priorityOrderOf(state);
    var idx = order.indexOf(it.priority);
    if (idx === -1) return RM.PALETTE.neutral;
    return RM.PRIORITY_RAMP[Math.min(3, Math.floor(idx * 4 / order.length))];
  };
  RM.colorForItem = function (state, it) {
    if (colorMode === 'epic') return RM.colorForEpic(state, it.epic);
    if (colorMode === 'assignee') {
      var id = (it.assignees || [])[0];
      var m = id ? (state.team || []).filter(function (x) { return x.id === id; })[0] : null;
      return RM.colorForMember(m);
    }
    if (colorMode === 'priority') return RM.colorForPriority(state, it);
    return RM.colorForWs(state, it.workstream);
  };
  // legend entries for a set of items under the active color mode:
  // [{ name, color }], in first-seen order (workstreams: default last)
  RM.colorLegend = function (state, items) {
    var seen = {}, out = [];
    function add(name, color) { if (!seen[name]) { seen[name] = true; out.push({ name: name, color: color }); } }
    items.forEach(function (it) {
      if (colorMode === 'epic') add(it.epic || 'No epic', RM.colorForEpic(state, it.epic));
      else if (colorMode === 'assignee') {
        var id = (it.assignees || [])[0];
        var m = id ? (state.team || []).filter(function (x) { return x.id === id; })[0] : null;
        add(m ? RM.memberLabel(m) : 'Unassigned', RM.colorForMember(m));
      } else if (colorMode === 'priority') {
        var sch = RM.prioritySchemeOf(state);
        var lbl = sch === 'none' ? 'No priority' : sch === 'rice' ? 'RICE' : (it.priority || 'No priority');
        add(lbl, RM.colorForPriority(state, it));
      } else add(it.workstream || RM.defaultWsName(state), RM.colorForWs(state, it.workstream));
    });
    if (colorMode === 'workstream') {
      var dn = RM.defaultWsName(state);
      out = out.filter(function (e) { return e.name !== dn; }).concat(out.filter(function (e) { return e.name === dn; }));
    }
    return out;
  };
  // profile avatars: initials + a deterministic color from the name
  RM.initialsOf = function (name) {
    var parts = String(name || '').trim().split(/\s+/).filter(Boolean);
    if (!parts.length) return '?';
    var a = parts[0].charAt(0);
    var b = parts.length > 1 ? parts[parts.length - 1].charAt(0) : (parts[0].charAt(1) || '');
    return (a + b).toUpperCase();
  };
  RM.avatarColor = function (name) {
    var h = 0;
    var s = String(name || '');
    for (var i = 0; i < s.length; i++) h = (h * 31 + s.charCodeAt(i)) >>> 0;
    return 'hsl(' + (h % 360) + ', 52%, 42%)';
  };
  RM.iconForEpic = function (state, epic) {
    return (state && state.epicIcons && epic && state.epicIcons[epic]) || null;
  };
  // seeded icons for the dataset's epics (lucide names)
  RM.DEFAULT_EPIC_ICONS = {
    'OS': 'cpu',
    'Integrations': 'plug',
    'Agents': 'bot',
    'Data': 'database',
    'Testing': 'flask-conical',
    'Agent Platform': 'map',
    'Workflows': 'workflow',
    'Process': 'network'
  };

  // Light tint of a hex color (for lead-in segments / soft fills), t in [0,1].
  RM.tint = function (hex, t) {
    var r = parseInt(hex.slice(0, 2), 16), g = parseInt(hex.slice(2, 4), 16), b = parseInt(hex.slice(4, 6), 16);
    function mix(c) { return Math.round(c + (255 - c) * t); }
    function h2(c) { var s = c.toString(16).toUpperCase(); return s.length === 1 ? '0' + s : s; }
    return h2(mix(r)) + h2(mix(g)) + h2(mix(b));
  };

  // ---------------------------------------------------------------- utils
  var uidCounter = 0;
  RM.uid = function (prefix) {
    uidCounter += 1;
    return (prefix || 'x') + Date.now().toString(36) + '-' + uidCounter + '-' + Math.random().toString(36).slice(2, 7);
  };
  RM.clone = function (o) { return JSON.parse(JSON.stringify(o)); };

  // \u-escape every non-ASCII char: the result is still valid JSON, and
  // pure-ASCII text is immune to the surrogate-pair corruption ExcelJS
  // exhibits at certain in-cell offsets (splitting mid-escape is fine —
  // concatenation restores it before JSON.parse). Also the bundle's
  // canonical form, so shard contents compare as plain ASCII strings.
  RM.asciiJson = function (obj) {
    return JSON.stringify(obj).replace(/[\u007F-\uFFFF]/g, function (ch) {
      return '\\u' + ('0000' + ch.charCodeAt(0).toString(16)).slice(-4);
    });
  };
  // Creation time (epoch ms) decoded from a uid's time36 segment; 0 for ids
  // that aren't uids (fixtures like 'p1'). time36 is 8 chars for any date
  // between 1972 and 2059 and the prefix may end in letters too, so the
  // segment is anchored from the right instead of by stripping letters.
  RM.uidTime = function (id) {
    var m = /^[a-z]+([0-9a-z]{8})-\d+-[0-9a-z]*$/.exec(String(id || ''));
    return m ? parseInt(m[1], 36) || 0 : 0;
  };

  // ---- fractional order keys. Entities carry `order`, a base-62 string
  // compared as a plain string; inserting between two rows mints one new key
  // and touches nothing else. Keys never end in '0' (a trailing zero would
  // make two different strings the same fraction).
  var ORDER_DIGITS = '0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz';
  function orderDigit(s, i) {
    var k = i < s.length ? ORDER_DIGITS.indexOf(s.charAt(i)) : 0;
    return k < 0 ? 0 : k;
  }
  // A usable key: base-62 digits only, trailing '0's dropped (they add
  // nothing to the fraction). Anything else — a foreign or hand-edited key —
  // reads as absent and gets regenerated.
  function orderKey(s) {
    if (typeof s !== 'string') return '';
    for (var i = 0; i < s.length; i++) if (ORDER_DIGITS.indexOf(s.charAt(i)) === -1) return '';
    return s.replace(/0+$/, '');
  }
  // A key strictly between a and b (null = open end). No room, or a >= b,
  // extends a with a suffix, so the result is always > a.
  RM.orderBetween = function (a, b) {
    a = orderKey(a);
    b = orderKey(b);
    if (!(b > a)) b = null;
    var n = ORDER_DIGITS.length, p, d;
    if (b == null) {
      // append: bump the last non-'z' digit and drop what follows — keys
      // stay short; only an all-'z' key needs a new digit
      for (p = a.length - 1; p >= 0 && orderDigit(a, p) === n - 1; p--) { /* skip */ }
      if (p < 0) return a + 'V';
      return a.slice(0, p) + ORDER_DIGITS.charAt(orderDigit(a, p) + 1);
    }
    if (!a) {
      // prepend: decrement b's last digit; a would-be trailing '0' becomes '0V'
      p = b.length - 1;
      d = orderDigit(b, p);
      if (d > 1) return b.slice(0, p) + ORDER_DIGITS.charAt(d - 1);
      if (d === 1) return b.slice(0, p) + '0V';
    }
    // between: walk the shared prefix, then bisect the first digit with room
    var out = '';
    for (var i = 0; ; i++) {
      if (b != null && i >= b.length) b = null; // out equals b — nothing fits below it
      var da = orderDigit(a, i);
      var db = b == null ? n : orderDigit(b, i);
      if (db - da > 1) return out + ORDER_DIGITS.charAt(Math.floor((da + db) / 2));
      out += ORDER_DIGITS.charAt(da);
      if (db > da) b = null; // already below b from here on
    }
  };
  RM.orderAfterAll = function (list) {
    var top = null;
    (list || []).forEach(function (x) {
      var k = x ? orderKey(x.order) : '';
      if (k && (top == null || k > top)) top = k;
    });
    return RM.orderBetween(top, null);
  };
  // in place, stable: by key, then id; entries without a key sink to the end
  RM.sortByOrder = function (list) {
    var idx = list.map(function (x, i) { return { x: x, i: i }; });
    idx.sort(function (p, q) {
      var a = p.x, b = q.x;
      var ka = typeof a.order === 'string' && a.order, kb = typeof b.order === 'string' && b.order;
      if (!ka !== !kb) return ka ? -1 : 1;
      if (ka && ka !== kb) return ka < kb ? -1 : 1;
      var ia = String(a.id || ''), ib = String(b.id || '');
      if (ia !== ib) return ia < ib ? -1 : 1;
      return p.i - q.i;
    });
    for (var k = 0; k < idx.length; k++) list[k] = idx[k].x;
    return list;
  };
  // Give every entry a key, interpolated from its array neighbours (so a
  // row spliced in by older code keeps its slot); a trailing keyless row
  // goes after the list's highest key. Then sort by key.
  RM.ensureOrder = function (list) {
    var prev = null, top = null;
    list.forEach(function (x) {
      // sanitize first: a key that fails orderKey is treated as missing
      x.order = orderKey(x.order) || null;
      if (x.order && (top == null || x.order > top)) top = x.order;
    });
    for (var i = 0; i < list.length; i++) {
      var x = list[i];
      if (typeof x.order !== 'string' || !x.order) {
        var next = null;
        for (var j = i + 1; j < list.length && next == null; j++) {
          if (typeof list[j].order === 'string' && list[j].order) next = list[j].order;
        }
        var base = prev;
        if (next == null || (prev != null && next <= prev)) { base = top; next = null; }
        x.order = RM.orderBetween(base, next);
        if (top == null || x.order > top) top = x.order;
      }
      prev = x.order;
    }
    return RM.sortByOrder(list);
  };
  // Move `entry` inside `list` to sit before `beforeId` (null = after the
  // last entry that inGroup accepts): ONE order key changes, relative to the
  // new neighbours, and the array mirrors the move for index-based code.
  function placeInList(list, entry, beforeId, inGroup) {
    var peers = RM.sortByOrder(list.filter(function (x) { return x !== entry && inGroup(x); }));
    var at = -1;
    peers.forEach(function (x, i) { if (x.id === beforeId) at = i; });
    var before = at === -1 ? null : peers[at];
    var prev = at === -1 ? peers[peers.length - 1] : peers[at - 1];
    entry.order = RM.orderBetween(prev ? prev.order : null, before ? before.order : null);
    var cur = list.indexOf(entry);
    if (cur !== -1) list.splice(cur, 1); // an entry not yet in the list is simply inserted
    var pos = list.length;
    if (before) pos = list.indexOf(before);
    else for (var i = list.length - 1; i >= 0; i--) if (inGroup(list[i])) { pos = i + 1; break; }
    list.splice(pos, 0, entry);
    return entry;
  }

  // ---------------------------------------------------------------- calendar
  RM.parseISO = function (iso) {
    var p = String(iso).slice(0, 10).split('-').map(Number);
    return new Date(Date.UTC(p[0], p[1] - 1, p[2]));
  };
  RM.fmtISO = function (dt) {
    if (!dt || !isFinite(dt.getTime())) return '';
    return dt.toISOString().slice(0, 10);
  };

  RM.numDays = function (meta) { return meta.numWeeks * RM.slotsOf(meta); };

  // keep meta.endDate (last working day, a Friday) in sync with numWeeks;
  // call after any change to numWeeks or timelineStart
  RM.syncEndDate = function (meta) {
    meta.endDate = RM.fmtISO(RM.dayToDate(meta, meta.numWeeks * RM.slotsOf(meta) - 1));
  };

  RM.weekStartDate = function (meta, week) {
    var d = RM.parseISO(meta.timelineStart);
    d.setUTCDate(d.getUTCDate() + week * 7);
    return d;
  };

  // ---- which weekdays work: meta.weekStart (0=Sun…6=Sat, default Monday)
  // is the first day of each week column; meta.workDays lists the working
  // weekdays (1–5 of them). The index space stays 5 slots per week — slot i
  // is the i-th working weekday, trailing slots read as non-working.
  RM.weekStartOf = function (metaOrState) {
    var m = metaOrState && metaOrState.meta ? metaOrState.meta : metaOrState;
    var w = m && m.weekStart;
    return isFinite(+w) && +w >= 0 && +w <= 6 ? Math.round(+w) : 1;
  };
  RM.workDaysOf = function (metaOrState) {
    var m = metaOrState && metaOrState.meta ? metaOrState.meta : metaOrState;
    var wd = m && m.workDays;
    if (Array.isArray(wd)) {
      var seen = {}, out = [];
      wd.forEach(function (d) {
        d = Math.round(+d);
        if (isFinite(d) && d >= 0 && d <= 6 && !seen[d]) { seen[d] = true; out.push(d); }
      });
      if (out.length >= 1 && out.length <= 7) return out;
    }
    return [1, 2, 3, 4, 5]; // Mon–Fri
  };
  // slots per index week = number of selected working days (1-7); every
  // slot is a working day — only holidays make a slot non-working
  RM.slotsOf = function (metaOrState) { return RM.workDaysOf(metaOrState).length; };
  // day-offsets (0–6 from the week's first day) of each working slot, ascending
  RM.workOffsets = function (meta) {
    var ws = RM.weekStartOf(meta);
    return RM.workDaysOf(meta)
      .map(function (d) { return ((d - ws) + 7) % 7; })
      .sort(function (a, b) { return a - b; });
  };

  // Re-establish the work-week invariants after ANY weekStart/workDays edit:
  // workDays sorted in week order, daysPerWeek in sync, and the timeline
  // start snapped back to the week's first day. Normalize and the Setup
  // handlers both go through this.
  RM.applyWorkWeek = function (m) {
    m.weekStart = RM.weekStartOf(m);
    var ws = m.weekStart;
    m.workDays = RM.workDaysOf(m).sort(function (a, b) {
      return (((a - ws) + 7) % 7) - (((b - ws) + 7) % 7);
    });
    m.daysPerWeek = m.workDays.length; // kept in sync for older readers
    var d = RM.parseISO(m.timelineStart);
    if (d && isFinite(d.getTime())) {
      var back = ((d.getUTCDay() - ws) + 7) % 7;
      if (back) {
        d.setUTCDate(d.getUTCDate() - back);
        m.timelineStart = RM.fmtISO(d);
      }
    }
    // endDate is left alone: normalize derives numWeeks from it, and a
    // snapped-back start only widens the window by part of a week
  };

  // Re-encode every stored day index from one week shape into another by
  // round-tripping through calendar dates — the one mapping both shapes
  // share. Used when the work week changes (slots per week, week start).
  RM.remapDaySpace = function (state, oldMeta) {
    var m = state.meta;
    function mapDay(d) {
      if (d == null || !isFinite(d)) return d;
      return Math.max(0, RM.dateToDay(m, RM.dayToDate(oldMeta, d)));
    }
    function mapSpan(obj) {
      if (obj.startDay == null) return;
      var s0 = obj.startDay;
      var e0 = s0 + Math.max(1, obj.durDays || 1) - 1;
      var s1 = mapDay(s0);
      var e1 = mapDay(e0);
      obj.startDay = s1;
      if (obj.durDays != null) {
        obj.durDays = obj.milestone ? 0 : Math.max(1, e1 - s1 + 1);
      }
    }
    (state.items || []).forEach(function (it) {
      mapSpan(it);
      (it.stories || []).forEach(mapSpan);
    });
    (state.costs || []).forEach(function (c) {
      c.startDay = mapDay(c.startDay) || 0;
      if (c.endDay != null) c.endDay = mapDay(c.endDay);
    });
    (state.phases || []).forEach(function (p) {
      if (p.startDay != null) p.startDay = mapDay(p.startDay);
      if (p.endDay != null) p.endDay = mapDay(p.endDay);
    });
  };

  // The one entry point for editing the work week on a LIVE document:
  // applies the setting, then re-encodes all day indices so every bar keeps
  // its calendar dates.
  RM.changeWorkWeek = function (state, patch) {
    var m = state.meta;
    var oldMeta = RM.clone(m);
    if (patch && Array.isArray(patch.workDays)) m.workDays = patch.workDays;
    if (patch && patch.weekStart != null) m.weekStart = patch.weekStart;
    RM.applyWorkWeek(m);
    RM.remapDaySpace(state, oldMeta);
  };

  RM.dayToDate = function (meta, day) {
    var S = RM.slotsOf(meta);
    var week = Math.floor(day / S);
    var slot = ((day % S) + S) % S;
    var offs = RM.workOffsets(meta);
    var d = RM.parseISO(meta.timelineStart);
    d.setUTCDate(d.getUTCDate() + week * 7 + offs[Math.min(slot, offs.length - 1)]);
    return d;
  };

  // Inclusive end date of a bar: last working day of the span.
  RM.spanEndDate = function (meta, startDay, durDays) {
    return RM.dayToDate(meta, startDay + Math.max(1, Math.ceil(durDays)) - 1);
  };

  // Returns null for an invalid/unparseable date — callers must handle it.
  RM.dateToDay = function (meta, date) {
    if (!date || !isFinite(date.getTime())) return null;
    var start = RM.parseISO(meta.timelineStart);
    var diff = Math.floor((date.getTime() - start.getTime()) / 86400000);
    var week = Math.floor(diff / 7);
    var rem = diff - week * 7;
    var offs = RM.workOffsets(meta);
    var slot = 0;
    for (var i = 0; i < offs.length; i++) if (offs[i] <= rem) slot = i;
    return week * RM.slotsOf(meta) + slot; // non-working weekday -> previous working slot
  };

  var MONTHS = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];
  RM.fmtShort = function (dt) { return MONTHS[dt.getUTCMonth()] + ' ' + dt.getUTCDate(); };
  RM.fmtShortYear = function (dt) { return MONTHS[dt.getUTCMonth()] + ' ' + dt.getUTCDate() + ' ’' + String(dt.getUTCFullYear()).slice(2); };

  // ---- work week: how many of the 5 index slots are working days, and what
  // a full-time week means in hours. Short weeks (e.g. 4-day) keep the 5-slot
  // index space — the trailing slot(s) simply count as non-working days.
  RM.daysPerWeekOf = function (metaOrState) {
    var m = metaOrState && metaOrState.meta ? metaOrState.meta : metaOrState;
    if (m && Array.isArray(m.workDays)) return RM.workDaysOf(m).length;
    var d = m && m.daysPerWeek; // pre-workDays documents
    return isFinite(+d) && +d >= 1 && +d <= 5 ? Math.round(+d) : 5;
  };
  RM.weekHoursOf = function (metaOrState) {
    var m = metaOrState && metaOrState.meta ? metaOrState.meta : metaOrState;
    var h = m && m.weekHours;
    return isFinite(+h) && +h > 0 ? +h : RM.WEEK_HOURS;
  };
  RM.hoursPerDay = function (metaOrState) {
    return RM.weekHoursOf(metaOrState) / RM.daysPerWeekOf(metaOrState);
  };

  // Holidays are INDIVIDUAL dates (meta.holidays, ISO strings). This builds a
  // { workingDayIndex: true } lookup; weekend-dated holidays are ignored since
  // weekends don't exist in the index space. Slots beyond meta.daysPerWeek
  // (a 4-day week's Fridays, say) read as non-working via offDay().
  RM.holidayDaySet = function (meta) {
    var set = {};
    var ws = RM.weekStartOf(meta);
    var offs = RM.workOffsets(meta);
    (meta.holidays || []).forEach(function (iso) {
      var d = RM.parseISO(iso);
      if (!d || !isFinite(d.getTime())) return;
      // only dates that fall ON a working weekday count — others don't
      // exist in the index space
      var off = ((d.getUTCDay() - ws) + 7) % 7;
      if (offs.indexOf(off) === -1) return;
      var day = RM.dateToDay(meta, d);
      if (day != null) set[day] = true;
    });
    return set;
  };

  // A non-working day slot: every slot is a selected working day, so only
  // an explicit holiday switches one off.
  RM.offDay = function (meta, day, set) {
    return (set || RM.holidayDaySet(meta))[day] === true;
  };

  RM.isHolidayDay = function (meta, day, set) {
    return RM.offDay(meta, day, set);
  };

  // A week is "blacked out" only when ALL its working days are holidays —
  // those weeks are excluded from capacity math entirely.
  RM.isBlackoutWeek = function (meta, week, set) {
    set = set || RM.holidayDaySet(meta);
    var S = RM.slotsOf(meta);
    for (var i = 0; i < S; i++) if (!set[week * S + i]) return false;
    return true;
  };

  // Non-working day slots in a week (explicit holidays + short-week slots).
  RM.holidaysInWeek = function (meta, week, set) {
    set = set || RM.holidayDaySet(meta);
    var n = 0;
    var S = RM.slotsOf(meta);
    for (var i = 0; i < S; i++) if (RM.offDay(meta, week * S + i, set)) n += 1;
    return n;
  };

  // ---- named holiday ranges (meta.holidayRanges: [{ name, start, end }],
  // end inclusive; a single day is a one-day range). meta.holidays stays the
  // derived flat date list every calendar function reads — call
  // syncHolidayDates after any range edit.
  RM.US_HOLIDAY_NAMES = {
    '2026-01-01': 'New Year’s', '2026-01-19': 'MLK Day', '2026-02-16': 'Presidents’ Day',
    '2026-05-22': 'Memorial Day', '2026-05-25': 'Memorial Day', '2026-06-19': 'Juneteenth',
    '2026-07-03': 'Independence Day', '2026-09-04': 'Labor Day', '2026-09-07': 'Labor Day',
    '2026-11-25': 'Thanksgiving', '2026-12-24': 'Christmas', '2027-01-01': 'New Year’s'
  };
  function addDaysIso(iso, n) {
    var d = RM.parseISO(iso);
    d.setUTCDate(d.getUTCDate() + n);
    return RM.fmtISO(d);
  }
  // do two dates belong to one observance? adjacent, or separated only by a
  // weekend (e.g. Fri + Mon around Memorial Day weekend)
  function holidayBridged(endIso, nextIso) {
    var gap = Math.round((RM.parseISO(nextIso) - RM.parseISO(endIso)) / 86400000);
    if (gap === 1) return true;
    if (gap > 3) return false;
    for (var i = 1; i < gap; i++) {
      var dow = RM.parseISO(addDaysIso(endIso, i)).getUTCDay();
      if (dow !== 0 && dow !== 6) return false;
    }
    return true;
  }
  RM.rangesFromDates = function (dates) {
    var out = [];
    (dates || []).slice().sort().forEach(function (iso) {
      var last = out[out.length - 1];
      if (last && holidayBridged(last.end, iso)) {
        last.end = iso;
        if (!last.name) last.name = RM.US_HOLIDAY_NAMES[iso] || '';
        return;
      }
      out.push({ name: RM.US_HOLIDAY_NAMES[iso] || '', start: iso, end: iso });
    });
    return out;
  };
  RM.syncHolidayDates = function (m) {
    var out = [], seen = {};
    (m.holidayRanges || []).forEach(function (r) {
      var iso = r.start, guard = 0;
      while (iso <= r.end && guard++ < 400) {
        if (!seen[iso]) { seen[iso] = true; out.push(iso); }
        iso = addDaysIso(iso, 1);
      }
    });
    m.holidays = out.sort();
  };
  RM.addHolidayRange = function (m, name, start, end) {
    m.holidayRanges.push({
      name: name || '',
      start: start,
      end: end && end >= start ? end : start
    });
    m.holidayRanges.sort(function (a, b) { return a.start < b.start ? -1 : a.start > b.start ? 1 : 0; });
    RM.syncHolidayDates(m);
  };
  // edit one range in place: patch = { name?, start?, end? }; a start past
  // the end (or vice versa) swaps them so the range stays valid
  RM.updateHolidayRange = function (m, idx, patch) {
    var r = (m.holidayRanges || [])[idx];
    if (!r) return false;
    if (patch.name != null) r.name = String(patch.name).trim();
    if (patch.start) r.start = patch.start;
    if (patch.end) r.end = patch.end;
    if (r.end < r.start) { var t = r.start; r.start = r.end; r.end = t; }
    m.holidayRanges.sort(function (a, b) { return a.start < b.start ? -1 : a.start > b.start ? 1 : 0; });
    RM.syncHolidayDates(m);
    return true;
  };
  RM.removeHolidayRange = function (m, idx) {
    m.holidayRanges.splice(idx, 1);
    RM.syncHolidayDates(m);
  };
  // carve [start, end] out of every range (used by the header week toggle)
  RM.clipHolidayRanges = function (m, start, end) {
    var out = [];
    (m.holidayRanges || []).forEach(function (r) {
      if (r.end < start || r.start > end) { out.push(r); return; }
      if (r.start < start) out.push({ name: r.name, start: r.start, end: addDaysIso(start, -1) });
      if (r.end > end) out.push({ name: r.name, start: addDaysIso(end, 1), end: r.end });
    });
    m.holidayRanges = out;
    RM.syncHolidayDates(m);
  };

  // Smallest calendar span (in working-day slots) whose non-holiday days >= workDays.
  RM.stretchSpan = function (meta, startDay, workDays) {
    if (workDays <= 0) return Math.max(0, workDays);
    var set = RM.holidayDaySet(meta);
    var remaining = workDays;
    var d = startDay;
    var guard = 0;
    while (remaining > 0 && guard < 20000) {
      if (!RM.offDay(meta, d, set)) remaining -= 1;
      d += 1;
      guard += 1;
    }
    return d - startDay;
  };

  // Non-holiday working days inside [startDay, startDay + span).
  RM.workInSpan = function (meta, startDay, span) {
    var set = RM.holidayDaySet(meta);
    var n = 0;
    for (var d = startDay; d < startDay + span; d++) {
      if (!RM.offDay(meta, d, set)) n += 1;
    }
    return n;
  };

  // Sprint numbering anchor: sprint boundaries fall every weeksPerSprint weeks
  // aligned to meta.sprintAnchor (default: timelineStart), and the sprint that
  // starts there is numbered meta.sprintAnchorNum (default 1). Weeks before
  // the anchor number down through S0, S-1, …
  // sprints can be disabled (weeksPerSprint = 0): the timeline is plain
  // weeks — internal sprint math then treats every week as its own bucket
  RM.sprintsEnabled = function (metaOrState) {
    var m = metaOrState && metaOrState.meta ? metaOrState.meta : metaOrState;
    return !m || m.weeksPerSprint !== 0;
  };
  RM.sprintInfo = function (meta) {
    var wps = meta.weeksPerSprint > 0 ? meta.weeksPerSprint : (meta.weeksPerSprint === 0 ? 1 : 2);
    var anchorWeek = 0;
    if (meta.sprintAnchor) {
      var d = RM.dateToDay(meta, RM.parseISO(meta.sprintAnchor));
      if (d != null) anchorWeek = Math.round(d / RM.slotsOf(meta));
    }
    return { wps: wps, anchorWeek: anchorWeek, firstNum: meta.sprintAnchorNum != null ? meta.sprintAnchorNum : 1 };
  };

  // working-day slots in one sprint (one week when sprints are off) — the
  // default length of a newly scheduled story or unsized feature
  RM.sprintDays = function (metaOrState) {
    var m = metaOrState && metaOrState.meta ? metaOrState.meta : metaOrState;
    return RM.slotsOf(m) * RM.sprintInfo(m).wps;
  };

  RM.sprintNumForWeek = function (meta, week) {
    var si = RM.sprintInfo(meta);
    return si.firstNum + Math.floor((week - si.anchorWeek) / si.wps);
  };

  // plain-text projection of stored rich text (tooltips, Excel cells, search)
  // Jira issue key ("HW-12") typed in by hand after a CSV import so later
  // exports can parent rows and update existing issues; null when unset.
  RM.jiraKeyOf = function (v) {
    if (typeof v !== 'string') return null;
    var k = v.trim().toUpperCase();
    return k ? k : null;
  };

  RM.htmlToText = function (html) {
    if (!html) return '';
    var t = String(html)
      .replace(/<(script|style)[\s\S]*?<\/\1\s*>/gi, '')
      .replace(/<br\s*\/?>/gi, '\n')
      .replace(/<\/(p|div|li|ul|ol|h[1-6])\s*>/gi, '\n')
      .replace(/<[^>]+>/g, '')
      .replace(/&nbsp;/g, ' ').replace(/&lt;/g, '<').replace(/&gt;/g, '>')
      .replace(/&quot;/g, '"').replace(/&#39;/g, "'").replace(/&amp;/g, '&');
    return t.replace(/[ \t]+\n/g, '\n').replace(/\n{3,}/g, '\n\n').trim();
  };

  RM.sizeDays = function (state, size, kind) {
    var k = sizeKeys(kind);
    var map = (state.meta && state.meta[k.days]) || (kind === 'story' ? RM.SIZE_SCHEMES[RM.DEFAULT_STORY_SIZE_SCHEME].days : RM.DEFAULT_SIZE_DAYS);
    return size && map[size] != null ? map[size] : null;
  };
  // Effective working days for a story: its size, else its own span, else
  // one sprint
  RM.storyEffortDays = function (state, st) {
    var sd = RM.sizeDays(state, st.size, 'story');
    if (sd != null) return sd;
    if (st.durDays != null) return st.durDays;
    return RM.sprintDays(state.meta);
  };

  // Nearest size option for a working-day count (ties resolve to the smaller size).
  RM.sizeForDays = function (state, days) {
    var map = (state.meta && state.meta.sizeDays) || RM.DEFAULT_SIZE_DAYS;
    var best = null, bestDiff = Infinity;
    RM.sizeOrderOf(state).forEach(function (s) {
      if (map[s] == null) return;
      var diff = Math.abs(map[s] - days);
      if (diff < bestDiff) { bestDiff = diff; best = s; }
    });
    return best;
  };

  // ---------------------------------------------------------------- history
  // One version-history entry: {t: epoch ms, u: user name, label, n: coalesce
  // count, d: change details [[category, field label, old, new], …],
  // x: overflow, tl: schedule moves}; null when unusable.
  RM.normalizeHistoryEntry = function (h) {
    if (!h || typeof h !== 'object') return null;
    var ht = +h.t;
    if (!isFinite(ht) || ht <= 0) return null;
    var d = Array.isArray(h.d) ? h.d.slice(0, RM.HISTORY_OPS_MAX).map(function (op) {
      if (!Array.isArray(op)) return null;
      return [String(op[0] || '').slice(0, 20), String(op[1] || '').slice(0, 120),
        String(op[2] == null ? '' : op[2]).slice(0, 400), String(op[3] == null ? '' : op[3]).slice(0, 400)];
    }).filter(Boolean) : [];
    var out = {
      t: Math.round(ht),
      u: typeof h.u === 'string' ? h.u.slice(0, 80) : '',
      label: typeof h.label === 'string' ? h.label.slice(0, 140) : '',
      n: isFinite(+h.n) && +h.n > 1 ? Math.round(+h.n) : 1
    };
    if (d.length) out.d = d;
    if (isFinite(+h.x) && +h.x > 0) out.x = Math.round(+h.x);
    // tl: machine-readable schedule moves for the visual timeline diff
    if (Array.isArray(h.tl)) {
      var tl = h.tl.slice(0, 60).map(function (m2) {
        if (!m2 || typeof m2 !== 'object') return null;
        function day(v) { return v == null || !isFinite(+v) ? null : Math.round(+v); }
        return { id: String(m2.id || ''), n: isFinite(+m2.n) ? +m2.n : 0,
          f: String(m2.f || '').slice(0, 60), ms: m2.ms ? 1 : 0,
          s0: day(m2.s0), d0: day(m2.d0), s1: day(m2.s1), d1: day(m2.d1) };
      }).filter(Boolean);
      if (tl.length) out.tl = tl;
    }
    return out;
  };
  RM.normalizeHistory = function (list) {
    return (Array.isArray(list) ? list : []).map(function (h) { return RM.normalizeHistoryEntry(h); })
      .filter(Boolean).slice(-RM.HISTORY_MAX);
  };
  // merge coalesced ops: same field keeps its FIRST old and LAST new value
  RM.mergeOps = function (base, add) {
    var out = base.slice();
    var at = {};
    // key by field AND label with a separator, so ('a','bc') never collides with ('ab','c')
    out.forEach(function (op, i) { at[op[0] + '\u0001' + op[1]] = i; });
    add.forEach(function (op) {
      var k = op[0] + '\u0001' + op[1];
      if (at[k] != null) out[at[k]] = [op[0], op[1], out[at[k]][2], op[3]];
      else { at[k] = out.length; out.push(op); }
    });
    return out.filter(function (op) { return op[2] !== op[3]; });
  };
  // coalesced schedule moves keep each item's FIRST before and LAST after
  RM.mergeTl = function (base, add) {
    var out = base.slice();
    var by = {};
    out.forEach(function (t, i) { by[t.id] = i; });
    add.forEach(function (t) {
      if (by[t.id] != null) {
        var b = out[by[t.id]];
        out[by[t.id]] = { id: t.id, n: t.n, f: t.f, ms: t.ms, s0: b.s0, d0: b.d0, s1: t.s1, d1: t.d1 };
      } else { by[t.id] = out.length; out.push(t); }
    });
    return out.filter(function (t) { return !(t.s0 === t.s1 && t.d0 === t.d1); });
  };

  // ---------------------------------------------------------------- state
  RM.normalizeState = function (raw) {
    var state = RM.clone(raw || {});
    state.meta = state.meta || {};
    var m = state.meta;
    m.title = m.title || 'Roadmap';
    // Prioritizing view: the product vision line every card ladders up to
    m.vision = typeof m.vision === 'string' ? m.vision : '';
    m.timelineStart = m.timelineStart || '2026-07-27';
    // work week shape: first day of week + which weekdays work (≤5).
    // Legacy daysPerWeek (3/4/5 from Monday) migrates to an explicit list.
    m.weekStart = RM.weekStartOf(m);
    if (!Array.isArray(m.workDays)) {
      m.workDays = [1, 2, 3, 4, 5].slice(0, RM.daysPerWeekOf(m));
    }
    RM.applyWorkWeek(m);
    m.numWeeks = m.numWeeks || (m.numSprints ? m.numSprints * (m.weeksPerSprint || 2) : 48);
    m.weeksPerSprint = [0, 1, 2, 4].indexOf(+m.weeksPerSprint) !== -1 ? +m.weeksPerSprint : 2;
    // capacity feature switch — roster-based scheduling constraints and the
    // capacity header row. OFF by default; enabled per-document in Setup.
    m.capacityEnabled = !!m.capacityEnabled;
    // apps switch (Setup → Apps): which header tabs this project shows. All
    // on by default; Planning is the home view and can never go off.
    var apps = (m.apps && typeof m.apps === 'object') ? m.apps : {};
    m.apps = {};
    RM.APPS.forEach(function (a) { m.apps[a[0]] = a[0] === 'planning' ? true : apps[a[0]] !== false; });
    // the saved project end date (last working day) wins over numWeeks
    if (m.endDate && /^\d{4}-\d{2}-\d{2}$/.test(m.endDate)) {
      var endWeeks = Math.floor((RM.parseISO(m.endDate) - RM.parseISO(m.timelineStart)) / (7 * 86400000)) + 1;
      if (endWeeks >= 1) m.numWeeks = Math.max(4, endWeeks);
    }
    RM.syncEndDate(m);
    m.holidays = (m.holidays || []).filter(function (iso) {
      return typeof iso === 'string' && /^\d{4}-\d{2}-\d{2}$/.test(iso);
    });
    // migrate legacy whole-week blackouts into five individual holiday dates
    if (m.blackoutWeeks && m.blackoutWeeks.length) {
      var seenHol = {};
      m.holidays.forEach(function (iso) { seenHol[iso] = true; });
      m.blackoutWeeks.forEach(function (iso) {
        var day = RM.dateToDay(m, RM.parseISO(iso));
        if (day == null) return;
        var mS = RM.slotsOf(m);
        var week = Math.floor(day / mS);
        for (var i = 0; i < mS; i++) {
          var dIso = RM.fmtISO(RM.dayToDate(m, week * mS + i));
          if (dIso && !seenHol[dIso]) { seenHol[dIso] = true; m.holidays.push(dIso); }
        }
      });
      m.holidays.sort();
    }
    delete m.blackoutWeeks;
    if (!m.holidaysV2026) {
      RM.US_HOLIDAYS_2026.forEach(function (iso) {
        if (m.holidays.indexOf(iso) === -1) m.holidays.push(iso);
      });
      m.holidays.sort();
      m.holidaysV2026 = true;
    }
    // named holiday ranges; docs saved before ranges existed migrate their
    // flat date list (consecutive/weekend-bridged dates merge, known US
    // observances get their names)
    if (!Array.isArray(m.holidayRanges)) {
      m.holidayRanges = RM.rangesFromDates(m.holidays);
    }
    m.holidayRanges = m.holidayRanges
      .map(function (r) {
        if (!r || typeof r.start !== 'string' || !/^\d{4}-\d{2}-\d{2}$/.test(r.start)) return null;
        var rEnd = typeof r.end === 'string' && /^\d{4}-\d{2}-\d{2}$/.test(r.end) && r.end >= r.start ? r.end : r.start;
        return { name: typeof r.name === 'string' ? r.name : '', start: r.start, end: rEnd };
      })
      .filter(Boolean)
      .sort(function (a, b) { return a.start < b.start ? -1 : a.start > b.start ? 1 : 0; });
    RM.syncHolidayDates(m);
    m.sprintAnchor = m.sprintAnchor || m.timelineStart;
    m.sprintAnchorNum = m.sprintAnchorNum != null && isFinite(m.sprintAnchorNum) ? m.sprintAnchorNum : 1;
    // scoping columns: ordered list of { key, label? }; built-in keys fall
    // back to their canonical label, custom keys ('c…') keep the user's
    // label. Any column may carry a user rename. Docs saved before scopeCols
    // existed get Description plus whichever built-ins actually hold content.
    var seenCol = {};
    var inheritedCols = m.scopeCols;
    if (!inheritedCols) {
      inheritedCols = RM.SCOPE_BUILTIN_ORDER.filter(function (k) {
        return RM.DEFAULT_SCOPE_COLS.indexOf(k) !== -1 || (state.items || []).some(function (it) { return it && it[k]; });
      }).map(function (k) { return { key: k }; });
    }
    m.scopeCols = inheritedCols
      .map(function (c) {
        if (typeof c === 'string') c = { key: c };
        if (!c || typeof c.key !== 'string' || !c.key) return null;
        var col;
        if (RM.SCOPE_BUILTIN_LABELS[c.key]) {
          col = c.label ? { key: c.key, label: String(c.label) } : { key: c.key };
        } else {
          col = { key: c.key, label: String(c.label || 'Column') };
        }
        // which rows show the column: 'feature' / 'story'; absent = both.
        // Built-ins with a default scope store an explicit 'both' so the
        // user's choice sticks across loads.
        var defScope = RM.SCOPE_BUILTIN_SCOPE[c.key];
        if (c.scope === 'feature' || c.scope === 'story') col.scope = c.scope;
        else if (defScope) col.scope = c.scope === 'both' ? 'both' : defScope;
        return col;
      })
      .filter(function (c) {
        if (!c || seenCol[c.key]) return false;
        seenCol[c.key] = true;
        return true;
      });
    // one-time migration: Description used to be hidden by default — surface
    // it left of Enables in docs saved before it joined the defaults
    if (!m.scopeDescV1) {
      m.scopeDescV1 = true;
      if (!seenCol.description) {
        var descAt = m.scopeCols.findIndex(function (c) { return c.key === 'enables'; });
        m.scopeCols.splice(descAt === -1 ? 0 : descAt, 0, { key: 'description' });
      }
    }
    // one-time migration: Acceptance criteria became a built-in (stories
    // only by default) right after Description. A custom column by that
    // name takes over the built-in key — its values move to item/story .ac
    // and its position and scope are kept — so nothing shows up twice.
    if (!m.scopeAcV1) {
      m.scopeAcV1 = true;
      if (!seenCol.ac) {
        var acAt = m.scopeCols.findIndex(function (c) {
          return !RM.SCOPE_BUILTIN_LABELS[c.key] && /^acceptance\s+criteria$/i.test(String(c.label || '').trim());
        });
        if (acAt !== -1) {
          var oldKey = m.scopeCols[acAt].key, oldScope = m.scopeCols[acAt].scope;
          m.scopeCols[acAt] = { key: 'ac', scope: oldScope || 'both' };
          if (Array.isArray(m.scopeColOrder)) {
            m.scopeColOrder = m.scopeColOrder.map(function (k) { return k === oldKey ? 'ac' : k; });
          }
          (state.items || []).forEach(function (it) {
            if (!it) return;
            if (it.custom && it.custom[oldKey] != null) { it.ac = it.custom[oldKey]; delete it.custom[oldKey]; }
            (it.stories || []).forEach(function (s) {
              if (s && s.custom && s.custom[oldKey] != null) { s.ac = s.custom[oldKey]; delete s.custom[oldKey]; }
            });
          });
        } else {
          var afterDesc = m.scopeCols.findIndex(function (c) { return c.key === 'description'; });
          m.scopeCols.splice(afterDesc === -1 ? 0 : afterDesc + 1, 0, { key: 'ac', scope: 'story' });
        }
        seenCol.ac = true;
      }
    }
    // full column order across FIXED and text columns (user-reorderable).
    // Default: Description and Epic lead, then the assessment/duration
    // cluster, then Workstream and any remaining text columns.
    (function () {
      var valid = {};
      RM.SCOPE_FIXED_KEYS.forEach(function (k) { valid[k] = true; });
      m.scopeCols.forEach(function (c) { valid[c.key] = true; });
      var out = [], seenK = {};
      function take(k) { if (valid[k] && !seenK[k]) { seenK[k] = true; out.push(k); } }
      // documents saved before the Deadline column slot it after Start,
      // matching the default order, instead of tacking it on at the end
      var savedOrder = Array.isArray(m.scopeColOrder) ? m.scopeColOrder.slice() : [];
      if (savedOrder.length && savedOrder.indexOf('deadline') === -1) {
        var atStart = savedOrder.indexOf('start');
        savedOrder.splice(atStart === -1 ? savedOrder.length : atStart + 1, 0, 'deadline');
      }
      // likewise Acceptance criteria slots right after Description (when
      // Description itself is unsaved both come from the default order)
      if (savedOrder.indexOf('ac') === -1 && savedOrder.indexOf('description') !== -1) {
        savedOrder.splice(savedOrder.indexOf('description') + 1, 0, 'ac');
      }
      savedOrder.forEach(take);
      RM.SCOPE_DEFAULT_ORDER.forEach(take);
      m.scopeCols.forEach(function (c) { take(c.key); });
      m.scopeColOrder = out;
    })();
    m.sizeDays = m.sizeDays || RM.clone(RM.DEFAULT_SIZE_DAYS);
    // migrate documents saved under the pre-2026-08 size metric
    var isLegacyMap = RM.SIZE_ORDER.every(function (s) { return m.sizeDays[s] === RM.LEGACY_SIZE_DAYS[s]; });
    if (isLegacyMap) m.sizeDays = RM.clone(RM.DEFAULT_SIZE_DAYS);
    // sizing approach: preset scheme, or 'custom' once edited; 'none' = off
    m.sizeScheme = RM.SIZE_SCHEMES[m.sizeScheme] ? m.sizeScheme : 'tshirt';
    if (Array.isArray(m.sizeOrder)) {
      var seenSz = {};
      m.sizeOrder = m.sizeOrder.map(String).filter(function (l) {
        if (!l || seenSz[l]) return false;
        seenSz[l] = true;
        return true;
      });
    } else {
      m.sizeOrder = (RM.SIZE_SCHEMES[m.sizeScheme].sizes || RM.SIZE_ORDER).slice();
    }
    var schemeDays = RM.SIZE_SCHEMES[m.sizeScheme].days || {};
    m.sizeOrder.forEach(function (l) {
      if (!isFinite(+m.sizeDays[l]) || +m.sizeDays[l] <= 0) {
        m.sizeDays[l] = schemeDays[l] || 5;
      }
    });
    // the story scale: a doc from before it existed hands stories the
    // feature scale when any story is already sized (so those sizes keep
    // meaning); otherwise stories start on story points
    if (m.storySizeScheme == null) {
      var anyStorySize = (state.items || []).some(function (it) {
        return (it && it.stories || []).some(function (st) { return st && st.size; });
      });
      if (anyStorySize) {
        m.storySizeScheme = m.sizeScheme;
        m.storySizeOrder = m.sizeOrder.slice();
        m.storySizeDays = RM.clone(m.sizeDays);
      } else {
        m.storySizeScheme = RM.DEFAULT_STORY_SIZE_SCHEME;
      }
    }
    m.storySizeScheme = RM.SIZE_SCHEMES[m.storySizeScheme] ? m.storySizeScheme : RM.DEFAULT_STORY_SIZE_SCHEME;
    if (Array.isArray(m.storySizeOrder)) {
      var seenSs = {};
      m.storySizeOrder = m.storySizeOrder.map(String).filter(function (l) {
        if (!l || seenSs[l]) return false;
        seenSs[l] = true;
        return true;
      });
    } else {
      m.storySizeOrder = (RM.SIZE_SCHEMES[m.storySizeScheme].sizes || []).slice();
    }
    var storySchemeDays = RM.SIZE_SCHEMES[m.storySizeScheme].days || {};
    m.storySizeDays = m.storySizeDays && typeof m.storySizeDays === 'object' ? m.storySizeDays : {};
    m.storySizeOrder.forEach(function (l) {
      if (!isFinite(+m.storySizeDays[l]) || +m.storySizeDays[l] <= 0) {
        m.storySizeDays[l] = storySchemeDays[l] || 5;
      }
    });
    // workstream feature switch — ON unless the project turned it off
    m.workstreamsEnabled = m.workstreamsEnabled !== false;
    // the default (null) workstream: user-visible name + color
    m.defaultWsName = typeof m.defaultWsName === 'string' && m.defaultWsName.trim()
      ? m.defaultWsName.trim() : 'General';
    m.defaultWsColor = (function () {
      var hex = String(m.defaultWsColor || '').replace(/^#/, '').toUpperCase();
      return /^[0-9A-F]{6}$/.test(hex) ? hex : RM.PALETTE.neutral;
    })();
    // priority column scheme (own column, separate from risk)
    m.priorityScheme = RM.PRIORITY_SCHEMES[m.priorityScheme] ? m.priorityScheme : 'none';
    // stories rank on their own scheme. A doc from before story schemes
    // existed keeps the feature scheme for stories that already carry a
    // priority; otherwise stories start on the severity ladder.
    var anyStoryPri = (state.items || []).some(function (it) {
      return (it && it.stories || []).some(function (st) { return st && st.priority; });
    });
    if (m.storyPriorityScheme == null) {
      m.storyPriorityScheme = anyStoryPri && m.priorityScheme !== 'rice' ? m.priorityScheme : RM.DEFAULT_STORY_PRIORITY_SCHEME;
    }
    if (!RM.PRIORITY_SCHEMES[m.storyPriorityScheme] || m.storyPriorityScheme === 'rice') {
      m.storyPriorityScheme = RM.DEFAULT_STORY_PRIORITY_SCHEME;
    }
    // a doc saved while MoSCoW lived under Risk migrates to the Priority column
    if (m.riskScheme === 'moscow') {
      m.riskScheme = 'none';
      m.priorityScheme = 'moscow';
      (state.items || []).forEach(function (it) {
        if (it && it.risk && !it.priority) { it.priority = it.risk; it.risk = null; }
      });
    }
    // assessment column scheme — docs that predate schemes keep their Risk
    // column only if they actually used it; new docs start without one
    if (!RM.RISK_SCHEMES[m.riskScheme]) {
      m.riskScheme = (state.items || []).some(function (it) { return it && it.risk; })
        ? 'risk' : 'none';
    }
    // work week: full-time hours + working days per week
    m.weekHours = isFinite(+m.weekHours) && +m.weekHours > 0 ? Math.min(80, +m.weekHours) : RM.WEEK_HOURS;
    m.daysPerWeek = isFinite(+m.daysPerWeek) && +m.daysPerWeek >= 1 && +m.daysPerWeek <= 5
      ? Math.round(+m.daysPerWeek) : 5;
    // rate card: role -> default hourly { rate, cost }
    var rcIn = m.rateCard && typeof m.rateCard === 'object' ? m.rateCard : {};
    m.rateCard = {};
    Object.keys(rcIn).forEach(function (k) {
      var v = rcIn[k] || {};
      var rr = isFinite(+v.rate) && +v.rate > 0 ? +v.rate : 0;
      var cc = isFinite(+v.cost) && +v.cost > 0 ? +v.cost : 0;
      if (rr || cc) m.rateCard[k] = { rate: rr, cost: cc };
    });
    delete m.sprintDates;

    state.phases = (state.phases || []).map(function (p) {
      // startDay/endDay: optional user-pinned phase window (working-day
      // indices); null = auto-derived from the phase's items
      var ps = p.startDay != null && isFinite(p.startDay) ? Math.max(0, Math.round(p.startDay)) : null;
      var pe = p.endDay != null && isFinite(p.endDay) ? Math.max(0, Math.round(p.endDay)) : null;
      if (ps != null && pe != null && pe <= ps) pe = ps + 1;
      return {
        id: p.id || RM.uid('p'),
        order: typeof p.order === 'string' && p.order ? p.order : null,
        name: p.name || 'Phase',
        description: p.description || '',
        bucket: !!p.bucket,
        collapsed: !!p.collapsed,
        startDay: ps,
        endDay: pe
      };
    });
    if (!state.phases.length) {
      state.phases = [{ id: RM.uid('p'), name: 'Phase 1', description: '', bucket: false, collapsed: false }];
    }
    RM.ensureOrder(state.phases);

    var phaseIds = {};
    state.phases.forEach(function (p) { phaseIds[p.id] = true; });
    var fallbackPhase = state.phases[0].id;

    var riskOrder = RM.RISK_SCHEMES[m.riskScheme].order || RM.RISK_ORDER;
    var prioOrder = RM.PRIORITY_SCHEMES[m.priorityScheme].order || [];
    var storyPrioOrder = RM.PRIORITY_SCHEMES[m.storyPriorityScheme].order || [];
    state.items = (state.items || []).map(function (it) {
      return {
        id: it.id || RM.uid('i'),
        order: typeof it.order === 'string' && it.order ? it.order : null,
        num: it.num != null ? it.num : null,
        phaseId: phaseIds[it.phaseId] ? it.phaseId : fallbackPhase,
        feature: it.feature || '',
        description: it.description || '',
        ac: it.ac || '',
        workstream: it.workstream || '',
        epic: it.epic || '',
        enables: it.enables || '',
        outOfScope: it.outOfScope || '',
        notes: it.notes || '',
        // item ids (legacy nums migrate below in migrateDepsToIds)
        deps: (function () {
          var out = [];
          (Array.isArray(it.deps) ? it.deps : []).forEach(function (d) {
            if (d == null || d === '') return;
            var s = String(d);
            if (out.indexOf(s) === -1) out.push(s);
          });
          return out;
        })(),
        depsText: it.depsText || [],
        extDeps: it.extDeps || '',
        size: it.size && RM.SIZE_ORDER.indexOf(it.size) !== -1 ? it.size : (it.size || null),
        // assessment value — validated against the active scheme's options;
        // legacy t-shirt risk values migrate: XS/S → L, XL → H
        risk: (function () {
          var rv = it.risk ? String(it.risk).toUpperCase() : null;
          if (!rv) return null;
          if (riskOrder.indexOf(rv) !== -1) return rv;
          if (riskOrder === RM.RISK_ORDER || riskOrder.indexOf('L') !== -1) {
            if (rv === 'XS' || rv === 'S') return 'L';
            if (rv === 'XL') return 'H';
          }
          return null;
        })(),
        priority: it.priority && prioOrder.indexOf(String(it.priority).toUpperCase()) !== -1
          ? String(it.priority).toUpperCase() : null,
        // hard deadline: a calendar date (ISO), so it survives work-week edits
        deadline: /^\d{4}-\d{2}-\d{2}$/.test(String(it.deadline || '')) ? String(it.deadline) : null,
        headcount: it.headcount != null && it.headcount > 0 ? it.headcount : 1,
        // role is descriptive metadata (capacity is role-agnostic); empty = any role
        teamType: it.teamType != null && it.teamType !== '' ? String(it.teamType) : '',
        // milestones are fixed dates: zero-duration diamonds on the timeline
        milestone: !!it.milestone,
        // milestone marker shape; absent = diamond (kept on bars so a
        // feature converted back and forth remembers its choice)
        msStyle: RM.MS_STYLES.indexOf(it.msStyle) > 0 ? it.msStyle : undefined,
        startDay: it.startDay != null && isFinite(it.startDay) ? it.startDay : null,
        durDays: it.durDays != null && isFinite(it.durDays)
          ? Math.max(it.milestone ? 0 : 1, it.durDays) : null,
        // risk t-shirt is planning metadata only — it never pads the schedule
        riskDays: 0,
        locked: !!it.locked,
        // custom scoping-column values, keyed by column key
        custom: (function () {
          var out = {};
          if (it.custom && typeof it.custom === 'object') {
            Object.keys(it.custom).forEach(function (k) {
              if (it.custom[k] != null && it.custom[k] !== '') out[k] = String(it.custom[k]);
            });
          }
          return out;
        })(),
        // RICE inputs (reach / impact / confidence % / effort) — feed the
        // computed score when the RICE priority scheme is active
        rice: (function () {
          var r = it.rice || {};
          function num(v) { return v != null && isFinite(v) && +v >= 0 ? +v : null; }
          return { reach: num(r.reach), impact: num(r.impact), confidence: num(r.confidence), effort: num(r.effort) };
        })(),
        colorOverride: it.colorOverride || null,
        jiraKey: RM.jiraKeyOf(it.jiraKey),
        // team-member ids working on this feature (validated against the
        // roster once the team is normalized below)
        assignees: Array.isArray(it.assignees) ? it.assignees.map(String) : [],
        done: !!it.done,
        stories: (it.stories || []).map(function (s) {
          // stories may carry their own little timeline (startDay/durDays);
          // both null = no timeline (the default)
          var sched = s.startDay != null && isFinite(s.startDay) && s.durDays > 0;
          return {
            id: s.id || RM.uid('s'), title: s.title || '', done: !!s.done,
            order: typeof s.order === 'string' && s.order ? s.order : null,
            jiraKey: RM.jiraKeyOf(s.jiraKey),
            size: s.size || null,
            priority: s.priority && storyPrioOrder.indexOf(String(s.priority).toUpperCase()) !== -1
              ? String(s.priority).toUpperCase() : null,
            // stories rate risk on the document's risk scheme, like features
            risk: s.risk && riskOrder.indexOf(String(s.risk).toUpperCase()) !== -1
              ? String(s.risk).toUpperCase() : null,
            assignees: Array.isArray(s.assignees) ? s.assignees.map(String) : [],
            // stories carry their own hard deadline, same shape as items
            deadline: /^\d{4}-\d{2}-\d{2}$/.test(String(s.deadline || '')) ? String(s.deadline) : null,
            // rich-text (sanitized HTML) story body + acceptance criteria
            description: typeof s.description === 'string' ? s.description : '',
            ac: typeof s.ac === 'string' ? s.ac : '',
            // scope-column values (same keys as item.custom) — stories share
            // the document's columns; workstream/epic roll up from the item
            custom: (function () {
              var out = {};
              if (s.custom && typeof s.custom === 'object') {
                Object.keys(s.custom).forEach(function (k) {
                  if (s.custom[k] != null && s.custom[k] !== '') out[k] = String(s.custom[k]);
                });
              }
              return out;
            })(),
            startDay: sched ? Math.max(0, Math.round(s.startDay)) : null,
            // an unscheduled story may still carry a duration (used when it
            // lands on the timeline, shown in the scoping grid)
            durDays: sched ? Math.round(s.durDays)
              : (s.durDays != null && isFinite(s.durDays) && s.durDays > 0 ? Math.round(s.durDays) : null)
          };
        })
      };
    });

    state.items.forEach(function (it) {
      delete it.leadDays;
      RM.ensureOrder(it.stories);
    });
    // canonical array order first, so a num collision resolves the same way
    // on every machine; then deps (which reference ids) can be resolved
    RM.ensureOrder(state.items);
    RM.dedupeNums(state);
    RM.migrateDepsToIds(state);

    state.epicColors = state.epicColors || {}; // legacy — display now keys off workstream
    state.wsColors = state.wsColors && typeof state.wsColors === 'object' ? state.wsColors : {};
    state.epicIcons = state.epicIcons && typeof state.epicIcons === 'object' ? state.epicIcons : {};
    // epic name -> Jira epic key (epics are strings on items, like epicIcons)
    var epicJira = {};
    if (state.epicJira && typeof state.epicJira === 'object') {
      Object.keys(state.epicJira).forEach(function (ep) {
        var k = RM.jiraKeyOf(state.epicJira[ep]);
        if (k) epicJira[ep] = k;
      });
    }
    state.epicJira = epicJira;
    // seed default colors/icons for well-known workstreams and epics
    state.items.forEach(function (it) {
      var w = it.workstream;
      if (w && !state.wsColors[w] && RM.DEFAULT_WS_COLORS[w]) state.wsColors[w] = RM.DEFAULT_WS_COLORS[w];
      var ep = it.epic;
      if (ep && !state.epicIcons[ep] && RM.DEFAULT_EPIC_ICONS[ep]) state.epicIcons[ep] = RM.DEFAULT_EPIC_ICONS[ep];
    });
    // workstream display order: saved order first (minus stale entries), then
    // any referenced-but-unlisted workstreams in first-appearance order
    var wsRef = {}, wsRefList = [];
    state.items.map(function (x) { return x.workstream; })
      .concat((state.team || []).reduce(function (a, x) { return a.concat(RM.memberWorkstreams(x)); }, []))
      .concat(Object.keys(state.wsColors))
      .forEach(function (w) { if (w && !wsRef[w]) { wsRef[w] = true; wsRefList.push(w); } });
    var wsOrder = [];
    (Array.isArray(state.wsOrder) ? state.wsOrder : []).forEach(function (w) {
      if (wsRef[w] && wsOrder.indexOf(w) === -1) wsOrder.push(w);
    });
    wsRefList.forEach(function (w) { if (wsOrder.indexOf(w) === -1) wsOrder.push(w); });
    state.wsOrder = wsOrder;
    state.teamTypes = state.teamTypes && state.teamTypes.length ? state.teamTypes : RM.clone(RM.DEFAULT_TEAM_TYPES);
    // fixed & recurring costs (budgeting)
    state.costs = (state.costs || []).map(function (c) {
      if (!c || typeof c !== 'object') return null;
      var kind = ['fixed', 'weekly', 'monthly'].indexOf(c.kind) !== -1 ? c.kind : 'fixed';
      var sd = isFinite(+c.startDay) && +c.startDay >= 0 ? Math.round(+c.startDay) : 0;
      var ed = kind !== 'fixed' && c.endDay != null && isFinite(+c.endDay) && +c.endDay >= sd
        ? Math.round(+c.endDay) : null;
      return {
        id: c.id || RM.uid('cost'),
        order: typeof c.order === 'string' && c.order ? c.order : null,
        name: typeof c.name === 'string' && c.name ? c.name : 'Cost',
        amount: isFinite(+c.amount) && +c.amount >= 0 ? +c.amount : 0,
        kind: kind,
        startDay: sd,
        endDay: ed
      };
    }).filter(Boolean);
    RM.ensureOrder(state.costs);
    // version history (see normalizeHistoryEntry for the entry shape)
    state.history = RM.normalizeHistory(state.history);
    // options — alternate plan versions. The active document carries its own
    // option id/name; the others are parked in state.options as full document
    // snapshots. A parked doc never nests options of its own.
    state.optId = typeof state.optId === 'string' && state.optId ? state.optId.slice(0, 40) : 'opt-default';
    state.optName = typeof state.optName === 'string' && state.optName.trim()
      ? state.optName.trim().slice(0, 60) : 'Default';
    state.options = (Array.isArray(state.options) ? state.options : []).filter(function (o) {
      return o && typeof o === 'object' && o.doc && typeof o.doc === 'object' && Array.isArray(o.doc.items);
    }).slice(0, RM.OPTIONS_MAX).map(function (o) {
      var doc = RM.clone(o.doc);
      delete doc.options;
      return {
        id: typeof o.id === 'string' && o.id ? o.id.slice(0, 40) : RM.uid('opt'),
        name: typeof o.name === 'string' && o.name.trim() ? o.name.trim().slice(0, 60) : 'Option',
        doc: doc
      };
    });
    state.team = (state.team || []).map(function (mbr) {
      // weekHours: { isoMonday: hours } — default 40 for any week not listed.
      // Legacy offWeeks (whole weeks off) migrate to 0-hour entries.
      var wh = {};
      if (mbr.weekHours && typeof mbr.weekHours === 'object') {
        Object.keys(mbr.weekHours).forEach(function (iso) {
          var h = Number(mbr.weekHours[iso]);
          if (/^\d{4}-\d{2}-\d{2}$/.test(iso) && isFinite(h) && h >= 0) wh[iso] = h;
        });
      }
      (mbr.offWeeks || []).forEach(function (iso) {
        if (typeof iso === 'string' && iso && wh[iso] == null) wh[iso] = 0;
      });
      // people can belong to several workstreams; `workstream` (the first)
      // stays for older readers and single-value call sites
      var wss = [];
      (Array.isArray(mbr.workstreams) ? mbr.workstreams
        : mbr.workstream ? [mbr.workstream] : []).forEach(function (w) {
        w = String(w || '').trim();
        if (w && wss.indexOf(w) === -1) wss.push(w);
      });
      return {
        id: mbr.id || RM.uid('t'),
        order: typeof mbr.order === 'string' && mbr.order ? mbr.order : null,
        // name is optional — a row can be a yet-unnamed seat ("Senior Dev TBD")
        name: mbr.name != null ? String(mbr.name) : '',
        // free-text role/title (what they do), independent of the rate card
        role: mbr.role != null ? String(mbr.role) : '',
        // rate-card role (drives default rate/cost); empty = not assigned
        type: mbr.type != null ? String(mbr.type) : '',
        workstream: wss[0] || '',
        workstreams: wss,
        // capacity at 40 h — a 0.5 role contributes half a head even full-time;
        // 0 (or blank) is allowed and contributes nothing
        capacity: mbr.capacity != null && mbr.capacity !== '' && isFinite(+mbr.capacity) && +mbr.capacity >= 0
          ? +mbr.capacity : 1,
        // hourly bill rate & hourly cost (budgeting view); 0 = not set
        rate: isFinite(+mbr.rate) && +mbr.rate >= 0 ? +mbr.rate : 0,
        cost: isFinite(+mbr.cost) && +mbr.cost >= 0 ? +mbr.cost : 0,
        weekHours: wh
      };
    });
    RM.ensureOrder(state.team);
    // every referenced work type must exist in the list
    state.items.forEach(function (it) {
      if (it.teamType && state.teamTypes.indexOf(it.teamType) === -1) state.teamTypes.push(it.teamType);
    });
    state.team.forEach(function (mbr) {
      if (mbr.type && state.teamTypes.indexOf(mbr.type) === -1) state.teamTypes.push(mbr.type);
    });
    var teamIds = {};
    state.team.forEach(function (mbr) { teamIds[mbr.id] = true; });
    function cleanAssignees(obj) {
      var seenA = {};
      obj.assignees = obj.assignees.filter(function (id) {
        if (!teamIds[id] || seenA[id]) return false;
        seenA[id] = true;
        return true;
      });
    }
    state.items.forEach(function (it) {
      cleanAssignees(it);
      it.stories.forEach(cleanAssignees);
    });

    // one-time day-space migration: documents saved before variable
    // slots-per-week kept a fixed 5-slot index week (short weeks read the
    // trailing slots as off). Re-encode their indices into the true
    // slots-per-week space via calendar dates.
    if (!m.dayspaceV2) {
      m.dayspaceV2 = true;
      if (RM.slotsOf(m) !== 5) {
        var legacyMeta = RM.clone(m);
        legacyMeta.workDays = [0, 1, 2, 3, 4].map(function (i) { return (m.weekStart + i) % 7; });
        RM.remapDaySpace(state, legacyMeta);
      }
    }
    return state;
  };

  RM.phaseIndex = function (state, phaseId) {
    for (var i = 0; i < state.phases.length; i++) if (state.phases[i].id === phaseId) return i;
    return -1;
  };

  RM.itemsInPhase = function (state, phaseId) {
    return state.items.filter(function (it) { return it.phaseId === phaseId; });
  };

  RM.itemByNum = function (state, num) {
    for (var i = 0; i < state.items.length; i++) if (state.items[i].num === num) return state.items[i];
    return null;
  };

  RM.itemById = function (state, id) {
    for (var i = 0; i < state.items.length; i++) if (state.items[i].id === id) return state.items[i];
    return null;
  };

  RM.nextNum = function (state) {
    var mx = 0;
    state.items.forEach(function (it) { if (it.num > mx) mx = it.num; });
    return mx + 1;
  };

  // Unique nums: assign missing, renumber collisions. Among duplicates the
  // OLDEST item (uid creation time, then array position) keeps its number —
  // the same answer on every machine that assembles the same shards. Safe
  // because deps reference ids, not nums.
  RM.dedupeNums = function (state) {
    var byNum = {}, maxNum = 0, loose = [];
    state.items.forEach(function (it, i) {
      if (it.num == null || !isFinite(it.num)) { loose.push(it); return; }
      (byNum[it.num] = byNum[it.num] || []).push({ it: it, i: i });
      if (it.num > maxNum) maxNum = it.num;
    });
    Object.keys(byNum).forEach(function (n) {
      var group = byNum[n];
      if (group.length < 2) return;
      group.sort(function (p, q) {
        var tp = RM.uidTime(p.it.id), tq = RM.uidTime(q.it.id);
        return tp !== tq ? tp - tq : p.i - q.i;
      });
      group.slice(1).forEach(function (x) { loose.push(x.it); });
    });
    // fresh numbers go out in array order
    var pos = {};
    state.items.forEach(function (it, i) { pos[it.id] = i; });
    loose.sort(function (p, q) { return pos[p.id] - pos[q.id]; });
    loose.forEach(function (it) { maxNum += 1; it.num = maxNum; });
    return state;
  };

  // deps reference item IDS. Older documents and Excel imports carry nums:
  // one that resolves becomes its id, one that doesn't moves to depsText as
  // '#n' so the reference stays visible. Self-references drop. Idempotent.
  RM.migrateDepsToIds = function (state) {
    state.items.forEach(function (it) {
      var out = [];
      if (!Array.isArray(it.depsText)) it.depsText = [];
      (it.deps || []).forEach(function (d) {
        var key = String(d);
        var dep = RM.itemById(state, key);
        if (!dep && /^\d+$/.test(key)) {
          // a legacy number: resolve it once; one that points nowhere becomes
          // free text so the reference is not lost
          dep = RM.itemByNum(state, Number(key));
          if (!dep) {
            if (it.depsText.indexOf('#' + key) === -1) it.depsText.push('#' + key);
            return;
          }
        }
        // an id whose target is missing stays put — validate flags it, and in
        // a shared bundle the target's shard may simply not have arrived yet
        var id = dep ? dep.id : key;
        if (id !== it.id && out.indexOf(id) === -1) out.push(id);
      });
      it.deps = out;
    });
    return state;
  };

  // Move an item into a phase, before another item (null = end of phase).
  RM.placeItem = function (state, itemId, phaseId, beforeItemId) {
    var it = RM.itemById(state, itemId);
    if (!it) return null;
    if (RM.phaseIndex(state, phaseId) !== -1) it.phaseId = phaseId;
    return placeInList(state.items, it, beforeItemId, function (x) { return x.phaseId === it.phaseId; });
  };
  RM.movePhaseTo = function (state, phaseId, beforeId) {
    var i = RM.phaseIndex(state, phaseId);
    if (i === -1) return null;
    return placeInList(state.phases, state.phases[i], beforeId, function () { return true; });
  };

  // The document as N+1 standalone plans — the active one first, then each
  // parked option — every doc normalized with the option bookkeeping
  // (optId/optName/options) removed, so nothing nests.
  RM.splitOptions = function (state) {
    var s = RM.normalizeState(state);
    function bare(doc) {
      var d = RM.normalizeState(doc);
      delete d.optId; delete d.optName; delete d.options;
      return d;
    }
    var out = [{ id: s.optId, name: s.optName, doc: bare(s) }];
    s.options.forEach(function (o) { out.push({ id: o.id, name: o.name, doc: bare(o.doc) }); });
    return out;
  };

  // ------------------------------------------------------------ import (add-only)
  // Merge a workbook into an open document without overwriting it: unknown
  // features, stories, team members and phases are added; a matched row only
  // gains values for fields it left empty. Where both sides hold a value and
  // disagree the difference is counted, never applied — the shared roadmap
  // wins. planImport is pure; applyImport mutates (inside commit).
  // teamType is not here: an empty role means "any role" (a real value, not
  // a gap), so a workbook must neither fill it nor count it as a conflict
  RM.IMPORT_FILL_FIELDS = ['enables', 'outOfScope', 'notes', 'extDeps', 'description', 'ac', 'size', 'risk'];
  RM.IMPORT_STORY_FILL_FIELDS = ['description', 'ac'];
  function normTitle(s) { return String(s == null ? '' : s).replace(/\s+/g, ' ').trim().toLowerCase(); }
  // Only an RM.uid-shaped id is identity across documents. Template imports
  // mint low-entropy ids ('ph1', 'tm1', …) on BOTH sides, so pairing those by
  // id would marry unrelated rows; they fall through to the name/title key.
  function strongId(x) {
    var id = x && x.id != null ? String(x.id) : '';
    return /^[a-z]+[0-9a-z]{6,}-\d+-[0-9a-z]{4,}$/.test(id) ? id : '';
  }
  function emptyVal(v) { return v == null || v === '' || (Array.isArray(v) && !v.length); }
  function sameVal(a, b) { return JSON.stringify(a) === JSON.stringify(b); }
  function sameSet(a, b) { return sameVal((a || []).slice().sort(), (b || []).slice().sort()); }
  // Pair `want` rows with `have` rows, one key function at a time; a key
  // pairs only when exactly one unmatched row on EACH side carries it, so a
  // shared title never guesses. Rows left over are adds.
  function pairRows(have, want, keyFns) {
    var pairs = [], hFree = have.slice(), wFree = want.slice();
    keyFns.forEach(function (keyOf) {
      var hBy = {}, wBy = {};
      hFree.forEach(function (h, i) { var k = keyOf(h); if (k) (hBy[k] = hBy[k] || []).push(i); });
      wFree.forEach(function (w, i) { var k = keyOf(w); if (k) (wBy[k] = wBy[k] || []).push(i); });
      // rows are tracked by position, not id: a workbook may repeat an id
      var hTaken = {}, wTaken = {};
      wFree.forEach(function (w, i) {
        var k = keyOf(w);
        if (!k || !hBy[k] || hBy[k].length !== 1 || wBy[k].length !== 1) return;
        pairs.push({ have: hFree[hBy[k][0]], want: w });
        hTaken[hBy[k][0]] = true;
        wTaken[i] = true;
      });
      hFree = hFree.filter(function (h, i) { return !hTaken[i]; });
      wFree = wFree.filter(function (w, i) { return !wTaken[i]; });
    });
    return { pairs: pairs, add: wFree };
  }
  function byId(x) { return strongId(x); } // import pairing: uid-shaped ids only (see strongId)
  // fill/conflict pass over one matched pair; returns the fields to fill
  function fillFields(have, want, fields, counter) {
    var out = {}, any = false;
    fields.forEach(function (f) {
      var hv = have[f], wv = want[f];
      if (emptyVal(wv)) return;
      if (emptyVal(hv)) { out[f] = RM.clone(wv); any = true; }
      else if (!sameVal(hv, wv)) counter.conflicts++;
    });
    return any ? out : null;
  }

  // state, incoming: normalized states. Returns a plan for applyImport plus
  // the counts the preview shows; nothing in `state` is touched.
  RM.planImport = function (state, incoming) {
    var plan = {
      items: { add: [], fill: [], matched: 0, conflicts: 0 },
      stories: { add: [], fill: [], matched: 0, conflicts: 0 },
      team: { add: [] },
      phases: { add: [] },
      summary: {}
    };
    function used(list) { var m = {}; (list || []).forEach(function (x) { m[x.id] = true; }); return m; }
    function fresh(id, taken, prefix) { var out = taken[id] ? RM.uid(prefix) : id; taken[out] = true; return out; }

    // phases: id, then name
    var phaseMap = {}, phaseTaken = used(state.phases);
    var ph = pairRows(state.phases, incoming.phases, [byId, function (p) { return normTitle(p.name); }]);
    ph.pairs.forEach(function (pr) { phaseMap[pr.want.id] = pr.have.id; });
    ph.add.forEach(function (p) {
      var np = RM.clone(p);
      np.id = fresh(p.id, phaseTaken, 'p');
      np.order = null;
      phaseMap[p.id] = np.id;
      plan.phases.add.push(np);
    });

    // team: id, then name (an unnamed seat never matches by name)
    var teamMap = {}, teamTaken = used(state.team);
    var tm = pairRows(state.team, incoming.team, [byId, function (m) { return normTitle(m.name); }]);
    tm.pairs.forEach(function (pr) { teamMap[pr.want.id] = pr.have.id; });
    tm.add.forEach(function (m) {
      var nm = RM.clone(m);
      nm.id = fresh(m.id, teamTaken, 't');
      nm.order = null;
      teamMap[m.id] = nm.id;
      plan.team.add.push(nm);
    });
    function mapAssignees(list) {
      var out = [];
      (list || []).forEach(function (a) { var to = teamMap[a]; if (to && out.indexOf(to) === -1) out.push(to); });
      return out;
    }

    // items: id, then num + title, then a title unique on both sides
    var idMap = {}, itemTaken = used(state.items);
    var im = pairRows(state.items, incoming.items, [
      byId,
      function (it) { return it.num != null && normTitle(it.feature) ? it.num + '|' + normTitle(it.feature) : ''; },
      function (it) { return normTitle(it.feature); }
    ]);
    im.pairs.forEach(function (pr) { idMap[pr.want.id] = pr.have.id; });
    var num = RM.nextNum(state), srcOf = {};
    im.add.forEach(function (it) {
      var n = RM.clone(it);
      delete n.holdPos;
      n.id = fresh(it.id, itemTaken, 'i');
      n.num = num++;
      n.order = null;
      n.phaseId = phaseMap[it.phaseId] || state.phases[0].id;
      n.assignees = mapAssignees(it.assignees);
      n.stories.forEach(function (s) { s.assignees = mapAssignees(s.assignees); });
      if (!idMap[it.id]) idMap[it.id] = n.id; // a matched row with this id keeps the mapping
      srcOf[n.id] = it;
      plan.items.add.push(n);
    });
    // deps reference incoming ids: a target that is matched or added maps to
    // its merged id; one already in the roadmap stays; anything else falls to
    // depsText as '#num' so the reference is not lost
    function mapDeps(it, selfId) {
      var deps = [], text = (it.depsText || []).slice();
      (it.deps || []).forEach(function (d) {
        var to = idMap[d] || (RM.itemById(state, d) ? d : null);
        if (to) { if (to !== selfId && deps.indexOf(to) === -1) deps.push(to); return; }
        var tgt = RM.itemById(incoming, d);
        var label = tgt && tgt.num != null ? '#' + tgt.num : (/^\d+$/.test(d) ? '#' + d : null);
        if (label && text.indexOf(label) === -1) text.push(label);
      });
      return { deps: deps, depsText: text };
    }
    plan.items.add.forEach(function (n) {
      var md = mapDeps(srcOf[n.id], n.id);
      n.deps = md.deps;
      n.depsText = md.depsText;
    });
    im.pairs.forEach(function (pr) {
      plan.items.matched++;
      var fields = fillFields(pr.have, pr.want, RM.IMPORT_FILL_FIELDS, plan.items) || {};
      if (normTitle(pr.want.feature) && normTitle(pr.have.feature) !== normTitle(pr.want.feature)) plan.items.conflicts++;
      // a dependency list is filled only when the roadmap has none at all
      var md = mapDeps(pr.want, pr.have.id);
      if (md.deps.length || md.depsText.length) {
        if (emptyVal(pr.have.deps) && emptyVal(pr.have.depsText)) { fields.deps = md.deps; fields.depsText = md.depsText; }
        else if (!sameSet(pr.have.deps, md.deps)) plan.items.conflicts++;
      }
      if (Object.keys(fields).length) plan.items.fill.push({ id: pr.have.id, fields: fields });
      // stories inside a matched item: id, then title
      var sm = pairRows(pr.have.stories, pr.want.stories, [byId, function (s) { return normTitle(s.title); }]);
      var sTaken = used(pr.have.stories);
      sm.add.forEach(function (s) {
        var ns = RM.clone(s);
        ns.id = fresh(s.id, sTaken, 's');
        ns.order = null;
        ns.assignees = mapAssignees(s.assignees);
        plan.stories.add.push({ itemId: pr.have.id, story: ns });
      });
      sm.pairs.forEach(function (sp) {
        plan.stories.matched++;
        var sf = fillFields(sp.have, sp.want, RM.IMPORT_STORY_FILL_FIELDS, plan.stories);
        if (normTitle(sp.want.title) && normTitle(sp.have.title) !== normTitle(sp.want.title)) plan.stories.conflicts++;
        if (sf) plan.stories.fill.push({ itemId: pr.have.id, id: sp.have.id, fields: sf });
      });
    });

    function fieldCount(list) { return list.reduce(function (n, f) { return n + Object.keys(f.fields).length; }, 0); }
    var s = plan.summary;
    s.items = plan.items.add.length;
    s.stories = plan.stories.add.length;
    s.fills = fieldCount(plan.items.fill) + fieldCount(plan.stories.fill);
    s.team = plan.team.add.length;
    s.phases = plan.phases.add.length;
    s.conflicts = plan.items.conflicts + plan.stories.conflicts;
    s.matched = plan.items.matched;
    s.empty = !(s.items || s.stories || s.fills || s.team || s.phases);
    return plan;
  };

  // Apply a plan in place. A fill re-checks that the field is STILL empty, so
  // a value typed since the preview is never overwritten. New rows go after
  // the last row of their list.
  RM.applyImport = function (state, plan) {
    var added = { items: 0, stories: 0, team: 0, phases: 0 }, filled = 0;
    function addType(t) { if (t && state.teamTypes.indexOf(t) === -1) state.teamTypes.push(t); }
    function fillInto(obj, fields) {
      Object.keys(fields).forEach(function (f) {
        if (!emptyVal(obj[f])) return;
        obj[f] = RM.clone(fields[f]);
        filled++;
      });
    }
    (plan.phases.add || []).forEach(function (p) {
      if (RM.phaseIndex(state, p.id) !== -1) return;
      var np = RM.clone(p);
      np.order = RM.orderAfterAll(state.phases);
      state.phases.push(np);
      added.phases++;
    });
    (plan.team.add || []).forEach(function (m) {
      if (state.team.some(function (x) { return x.id === m.id; })) return;
      var nm = RM.clone(m);
      nm.order = RM.orderAfterAll(state.team);
      state.team.push(nm);
      addType(nm.type);
      added.team++;
    });
    (plan.items.add || []).forEach(function (it) {
      if (RM.itemById(state, it.id)) return;
      var n = RM.clone(it);
      n.order = RM.orderAfterAll(state.items);
      if (RM.phaseIndex(state, n.phaseId) === -1) n.phaseId = state.phases[0].id;
      RM.ensureOrder(n.stories);
      state.items.push(n);
      addType(n.teamType);
      added.items++;
    });
    (plan.stories.add || []).forEach(function (a) {
      var it = RM.itemById(state, a.itemId);
      if (!it || it.stories.some(function (s) { return s.id === a.story.id; })) return;
      var ns = RM.clone(a.story);
      ns.order = RM.orderAfterAll(it.stories);
      it.stories.push(ns);
      added.stories++;
    });
    (plan.items.fill || []).forEach(function (f) {
      var it = RM.itemById(state, f.id);
      if (!it) return;
      fillInto(it, f.fields);
      addType(it.teamType);
    });
    (plan.stories.fill || []).forEach(function (f) {
      var it = RM.itemById(state, f.itemId);
      var st = it && it.stories.filter(function (s) { return s.id === f.id; })[0];
      if (st) fillInto(st, f.fields);
    });
    return { added: added, filled: filled };
  };

  // ------------------------------------------------------------ scope columns
  RM.scopeColLabel = function (col) {
    return col.label || RM.SCOPE_BUILTIN_LABELS[col.key] || 'Column';
  };
  RM.renameScopeCol = function (state, key, label) {
    state.meta.scopeCols.forEach(function (c) {
      if (c.key !== key) return;
      if (label && label !== RM.SCOPE_BUILTIN_LABELS[key]) c.label = label;
      else delete c.label; // empty (or canonical) restores the built-in name
    });
  };
  // does a column show on rows of this kind ('feature' | 'story')?
  RM.scopeColShows = function (col, kind) {
    return !col.scope || col.scope === 'both' || col.scope === kind;
  };
  // 'feature' / 'story' restrict the column; 'both' (or anything else) clears
  RM.setScopeColScope = function (state, key, scope) {
    if (scope !== 'feature' && scope !== 'story' && scope !== 'both') return;
    state.meta.scopeCols.forEach(function (c) {
      if (c.key !== key) return;
      if (scope !== 'both') c.scope = scope;
      else if (RM.SCOPE_BUILTIN_SCOPE[key]) c.scope = 'both'; // explicit: overrides the default
      else delete c.scope;
    });
  };
  // stories keep Description and Acceptance criteria as own fields; every
  // other column (built-in or custom) lives in story.custom
  RM.STORY_FIELD_KEYS = { description: true, ac: true };
  RM.storyScopeValue = function (st, key) {
    if (RM.STORY_FIELD_KEYS[key]) return st[key] || '';
    return (st.custom && st.custom[key]) || '';
  };
  RM.setStoryScopeValue = function (st, key, val) {
    if (RM.STORY_FIELD_KEYS[key]) { st[key] = val; return; }
    if (!st.custom) st.custom = {};
    if (val) st.custom[key] = val; else delete st.custom[key];
  };
  RM.scopeValue = function (it, key) {
    if (RM.SCOPE_BUILTIN_LABELS[key]) return it[key] || '';
    return (it.custom && it.custom[key]) || '';
  };
  RM.setScopeValue = function (it, key, val) {
    if (RM.SCOPE_BUILTIN_LABELS[key]) { it[key] = val; return; }
    if (!it.custom) it.custom = {};
    if (val) it.custom[key] = val; else delete it.custom[key];
  };
  // add a custom column (or re-show a hidden built-in when key is given)
  function orderAppend(m, key) {
    if (Array.isArray(m.scopeColOrder) && m.scopeColOrder.indexOf(key) === -1) {
      m.scopeColOrder.push(key);
    }
  }
  RM.addScopeCol = function (state, label, key) {
    var cols = state.meta.scopeCols;
    if (key && RM.SCOPE_BUILTIN_LABELS[key]) {
      if (!cols.some(function (c) { return c.key === key; })) {
        var col = { key: key };
        if (RM.SCOPE_BUILTIN_SCOPE[key]) col.scope = RM.SCOPE_BUILTIN_SCOPE[key];
        cols.push(col);
      }
      orderAppend(state.meta, key);
      return key;
    }
    var k = RM.uid('c');
    cols.push({ key: k, label: label || 'Column' });
    orderAppend(state.meta, k);
    return k;
  };
  RM.removeScopeCol = function (state, key) {
    state.meta.scopeCols = state.meta.scopeCols.filter(function (c) { return c.key !== key; });
    if (Array.isArray(state.meta.scopeColOrder)) {
      state.meta.scopeColOrder = state.meta.scopeColOrder.filter(function (k) { return k !== key; });
    }
    if (!RM.SCOPE_BUILTIN_LABELS[key]) {
      state.items.forEach(function (it) { if (it.custom) delete it.custom[key]; });
    }
  };
  RM.moveScopeCol = function (state, key, delta) {
    var cols = state.meta.scopeCols;
    var i = -1;
    cols.forEach(function (c, ix) { if (c.key === key) i = ix; });
    if (i === -1) return;
    var j = Math.max(0, Math.min(cols.length - 1, i + delta));
    if (j === i) return;
    var c = cols.splice(i, 1)[0];
    cols.splice(j, 0, c);
  };

  // Effective working days of effort for an item.
  RM.effortDays = function (state, it) {
    var sd = RM.sizeDays(state, it.size);
    if (sd != null) return sd;
    if (it.durDays != null) return it.durDays;
    return RM.sprintDays(state.meta);
  };

  // End of an item INCLUDING its risk buffer — dependents plan around the
  // buffer, that's what it's for.
  RM.itemEnd = function (it) {
    return it.startDay != null && it.durDays != null ? it.startDay + it.durDays + (it.riskDays || 0) : null;
  };

  // Total calendar span on the grid (work + risk buffer). Milestones are a
  // point in time — dependents may start the same day.
  RM.itemSpan = function (it) {
    if (it.milestone) return 0;
    return (it.durDays || 0) + (it.riskDays || 0);
  };

  // Risk is metadata only now — it contributes no working days to the plan.
  RM.riskEffortDays = function () { return 0; };

  // ------------------------------------------------------------ dependencies
  // Concrete dependency item list from the explicit deps (item ids); unknown
  // = ids with no item, i.e. deleted. ("All above" support was removed —
  // only specifically-defined deps count.)
  RM.resolveDeps = function (state, it) {
    var out = { deps: [], unknown: [] };
    (it.deps || []).forEach(function (id) {
      var dep = RM.itemById(state, id);
      if (!dep) out.unknown.push(id);
      else if (dep.id !== it.id) out.deps.push(dep);
    });
    return out;
  };

  // All concrete edges as [depItem, item] pairs.
  RM.depEdges = function (state) {
    var edges = [];
    state.items.forEach(function (it) {
      RM.resolveDeps(state, it).deps.forEach(function (dep) {
        edges.push([dep, it]);
      });
    });
    return edges;
  };

  // Set of item ids participating in at least one dependency cycle.
  RM.cycleMembers = function (state) {
    var adj = {};
    state.items.forEach(function (it) { adj[it.id] = []; });
    RM.depEdges(state).forEach(function (e) { adj[e[0].id].push(e[1].id); });

    // Tarjan SCC, iterative.
    var index = 0, stack = [], onStack = {}, idx = {}, low = {}, cyclic = {};
    state.items.forEach(function (root0) {
      if (idx[root0.id] != null) return;
      var work = [[root0.id, 0]];
      while (work.length) {
        var top = work[work.length - 1];
        var v = top[0];
        if (top[1] === 0) {
          idx[v] = low[v] = index++;
          stack.push(v); onStack[v] = true;
        }
        var advanced = false;
        var neighbors = adj[v];
        while (top[1] < neighbors.length) {
          var w = neighbors[top[1]];
          top[1] += 1;
          if (idx[w] == null) { work.push([w, 0]); advanced = true; break; }
          if (onStack[w]) low[v] = Math.min(low[v], idx[w]);
        }
        if (advanced) continue;
        if (low[v] === idx[v]) {
          var comp = [];
          var u;
          do { u = stack.pop(); onStack[u] = false; comp.push(u); } while (u !== v);
          if (comp.length > 1) comp.forEach(function (id) { cyclic[id] = true; });
          else {
            // self-loop
            var self = comp[0];
            if (adj[self].indexOf(self) !== -1) cyclic[self] = true;
          }
        }
        work.pop();
        if (work.length) {
          var parent = work[work.length - 1][0];
          low[parent] = Math.min(low[parent], low[v]);
        }
      }
    });
    return cyclic;
  };

  // ------------------------------------------------------------ capacity
  // A member's hours for a given week index (default 40; keyed by ISO Monday).
  RM.memberHoursForWeek = function (meta, member, week) {
    if (!member.weekHours) return RM.weekHoursOf(meta);
    var iso = RM.fmtISO(RM.weekStartDate(meta, week));
    var h = member.weekHours[iso];
    return h != null && isFinite(h) ? h : RM.weekHoursOf(meta);
  };

  // Back-compat: "off" = a zero-hour week.
  RM.memberOffWeek = function (meta, member, week) {
    return RM.memberHoursForWeek(meta, member, week) === 0;
  };

  // Roster availability for one week, in PEOPLE-EQUIVALENTS (hours ÷ the
  // project's full-time week): the fractional headcount actually available.
  RM.availForWeek = function (state, week) {
    var total = 0;
    var full = RM.weekHoursOf(state.meta);
    state.team.forEach(function (m) {
      var pe = (RM.memberHoursForWeek(state.meta, m, week) / full) * (m.capacity != null ? m.capacity : 1);
      if (pe > 0) total += pe;
    });
    return { total: total };
  };

  // How much team focus one active item consumes, in "focus units".
  // Grounded in Kanban WIP-limit practice (and Little's Law): a team's
  // throughput collapses when concurrent work outgrows the people available,
  // and BIG items eat disproportionate focus — most teams can only run a
  // couple of large initiatives at once. An M (2 working weeks) counts as 1
  // unit; weight scales with the item's working days ÷ 10, clamped to
  // [0.3, 2] so a swarm of XS chores still registers and an XL can't demand
  // more than two people-equivalents of attention.
  RM.wipWeight = function (state, it) {
    var days = it.startDay != null && it.durDays != null
      ? RM.workInSpan(state.meta, it.startDay, it.durDays)
      : (RM.sizeDays(state, it.size) || it.durDays || 5);
    return Math.max(0.3, Math.min(2, days / 10));
  };

  // Weekly size-weighted WIP vs the people available. Blackout weeks carry
  // no demand and no check; member off-weeks lower that week's availability.
  RM.capacity = function (state) {
    var meta = state.meta;
    var weeks = [];
    var teamTotal = state.team.length;
    var w;
    for (w = 0; w < meta.numWeeks; w++) {
      var avail = RM.availForWeek(state, w);
      weeks.push({
        demand: 0, items: [],
        cap: teamTotal > 0 ? avail.total : Infinity,
        blackout: RM.isBlackoutWeek(meta, w),
        over: false
      });
    }
    state.items.forEach(function (it) {
      if (it.startDay == null || it.durDays == null || it.done || it.milestone) return;
      var wt = RM.wipWeight(state, it);
      var S = RM.slotsOf(meta);
      var w0 = Math.floor(it.startDay / S);
      var w1 = Math.floor((it.startDay + it.durDays - 1) / S);
      for (var wk = Math.max(0, w0); wk <= Math.min(meta.numWeeks - 1, w1); wk++) {
        var cell = weeks[wk];
        if (cell.blackout) continue;
        cell.demand += wt;
        cell.items.push(it.id);
      }
    });
    weeks.forEach(function (cell) {
      if (cell.demand > cell.cap + 1e-9) cell.over = true;
    });
    return { weeks: weeks, teamTotal: teamTotal };
  };

  // ------------------------------------------------------------ dependency risk
  // Heuristic risk estimate from the dependency graph, per item.
  // Returns { level: 'none'|'low'|'med'|'high', score, reasons: [] }.
  RM.depRisk = function (state, it, cyclic) {
    cyclic = cyclic || RM.cycleMembers(state);
    var res = RM.resolveDeps(state, it);
    var active = res.deps.filter(function (d) { return !d.done; });
    var score = 0, reasons = [];

    if (cyclic[it.id]) { score += 4; reasons.push('In a dependency cycle'); }
    if (res.unknown.length) { score += 2; reasons.push(res.unknown.length + ' unknown dependenc' + (res.unknown.length > 1 ? 'ies' : 'y')); }
    if (active.length >= 3) { score += 2; reasons.push('Depends on ' + active.length + ' open items'); }
    else if (active.length) { score += 1; reasons.push('Depends on ' + active.length + ' open item' + (active.length > 1 ? 's' : '')); }

    var unsched = active.filter(function (d) { return RM.itemEnd(d) == null; });
    if (unsched.length) { score += 2; reasons.push(unsched.length + ' dependenc' + (unsched.length > 1 ? 'ies' : 'y') + ' not scheduled yet'); }

    if (it.startDay != null && it.durDays != null) {
      var minSlack = null;
      active.forEach(function (d) {
        var e = RM.itemEnd(d);
        if (e == null) return;
        var gap = it.startDay - e;
        if (minSlack == null || gap < minSlack) minSlack = gap;
      });
      if (minSlack != null && minSlack < 0) { score += 3; reasons.push('Starts before a dependency finishes'); }
      else if (minSlack != null && minSlack === 0) { score += 1; reasons.push('Zero slack — any dependency slip pushes this'); }
    }

    // upstream chain depth (bounded walk; cycles already scored above)
    var depth = 0, frontier = active, seen = {}, guard = 0;
    while (frontier.length && depth < 6 && guard < 400) {
      depth += 1;
      var next = [];
      frontier.forEach(function (d) {
        if (seen[d.id]) return;
        seen[d.id] = true;
        RM.resolveDeps(state, d).deps.forEach(function (dd) {
          if (!dd.done && !seen[dd.id]) next.push(dd);
        });
        guard += 1;
      });
      frontier = next;
    }
    if (depth >= 3) { score += 1; reasons.push('Dependency chain ' + depth + ' deep'); }

    var level = score === 0 ? 'none' : score <= 2 ? 'low' : score <= 4 ? 'med' : 'high';
    return { level: level, score: score, reasons: reasons };
  };

  // ------------------------------------------------------------ critical path
  // Edges on the longest scheduled dependency chain (by duration). Cycle
  // members are excluded so the walk terminates. Returns
  // { edges: {"depId>itemId":true}, items: {id:true}, total }.
  RM.criticalPath = function (state) {
    var cyclic = RM.cycleMembers(state);
    var nodes = state.items.filter(function (it) {
      return it.startDay != null && it.durDays != null && !it.done && !cyclic[it.id];
    });
    var byId = {};
    nodes.forEach(function (it) { byId[it.id] = it; });
    var edges = [];
    RM.depEdges(state).forEach(function (e) {
      if (byId[e[0].id] && byId[e[1].id]) edges.push(e);
    });

    var preds = {}, succs = {}, indeg = {}, outdeg = {};
    nodes.forEach(function (it) { preds[it.id] = []; succs[it.id] = []; indeg[it.id] = 0; outdeg[it.id] = 0; });
    edges.forEach(function (e) {
      succs[e[0].id].push(e[1].id);
      preds[e[1].id].push(e[0].id);
      indeg[e[1].id] += 1;
      outdeg[e[0].id] += 1;
    });

    // Kahn topological order.
    var order = [], q = [];
    nodes.forEach(function (it) { if (indeg[it.id] === 0) q.push(it.id); });
    var indegLeft = {};
    nodes.forEach(function (it) { indegLeft[it.id] = indeg[it.id]; });
    while (q.length) {
      var id = q.shift();
      order.push(id);
      succs[id].forEach(function (nid) {
        indegLeft[nid] -= 1;
        if (indegLeft[nid] === 0) q.push(nid);
      });
    }

    var up = {}, down = {};
    order.forEach(function (id2) {
      var best = 0;
      preds[id2].forEach(function (p) { if (up[p] > best) best = up[p]; });
      up[id2] = best + RM.itemSpan(byId[id2]);
    });
    for (var i = order.length - 1; i >= 0; i--) {
      var id3 = order[i];
      var best2 = 0;
      succs[id3].forEach(function (sId) { if (down[sId] > best2) best2 = down[sId]; });
      down[id3] = best2 + RM.itemSpan(byId[id3]);
    }

    var total = 0;
    order.forEach(function (id4) {
      var t = up[id4] + down[id4] - RM.itemSpan(byId[id4]);
      if (t > total) total = t;
    });

    var critItems = {}, critEdges = {};
    if (total > 0) {
      order.forEach(function (id5) {
        if (up[id5] + down[id5] - RM.itemSpan(byId[id5]) === total) critItems[id5] = true;
      });
      edges.forEach(function (e) {
        if (critItems[e[0].id] && critItems[e[1].id] && up[e[0].id] + down[e[1].id] === total) {
          critEdges[e[0].id + '>' + e[1].id] = true;
        }
      });
    }
    return { edges: critEdges, items: critItems, total: total };
  };

  // ------------------------------------------------------------ validation
  RM.validate = function (state) {
    var byItem = {};
    var global = [];
    function add(it, level, code, msg) {
      (byItem[it.id] = byItem[it.id] || []).push({ level: level, code: code, msg: msg });
    }

    // duplicate nums
    var byNum = {};
    state.items.forEach(function (it) { (byNum[it.num] = byNum[it.num] || []).push(it); });
    Object.keys(byNum).forEach(function (n) {
      if (byNum[n].length > 1) {
        byNum[n].forEach(function (it) { add(it, 'error', 'DUP_NUM', 'Duplicate ID #' + n); });
      }
    });

    var cyclic = RM.cycleMembers(state);
    var phaseById = {};
    state.phases.forEach(function (p) { phaseById[p.id] = p; });
    var horizon = RM.numDays(state.meta);

    state.items.forEach(function (it) {
      var res = RM.resolveDeps(state, it);
      res.unknown.forEach(function () {
        add(it, 'warn', 'UNKNOWN_DEP', 'Depends on a deleted item');
      });
      if (cyclic[it.id]) add(it, 'error', 'CYCLE', 'Part of a dependency cycle');
      if (it.deps.indexOf(it.id) !== -1) add(it, 'warn', 'SELF_DEP', 'Depends on itself (ignored)');
      if (!it.feature.trim()) add(it, 'warn', 'NO_TITLE', 'Feature has no title');

      var scheduled = it.startDay != null && it.durDays != null;
      var phase = phaseById[it.phaseId];
      if (scheduled) {
        if (it.startDay < 0 || it.startDay + RM.itemSpan(it) > horizon) {
          add(it, 'warn', 'OFF_TIMELINE', 'Bar extends outside the timeline');
        }
        res.deps.forEach(function (dep) {
          var depEnd = RM.itemEnd(dep);
          if (depEnd == null) {
            if (!dep.done) add(it, 'info', 'DEP_UNSCHEDULED', 'Dependency #' + dep.num + ' (' + dep.feature + ') is not scheduled');
          } else if (it.startDay < depEnd && !dep.done) {
            add(it, 'warn', 'DEP_ORDER', 'Starts before dependency #' + dep.num + ' (' + dep.feature + ') finishes');
          }
        });
      } else if (phase && !phase.bucket && !it.done) {
        add(it, 'info', 'UNSCHEDULED', 'In an active phase but not on the timeline');
      }

    });

    var cap = RM.capacity(state);
    cap.weeks.forEach(function (cell, w) {
      if (!state.meta.capacityEnabled || !cell.over) return;
      var d = RM.weekStartDate(state.meta, w);
      function r1(x) { return Math.round(x * 10) / 10; }
      var what = r1(cell.demand) + ' focus units vs ' + r1(cell.cap) + ' people available';
      global.push({
        level: 'warn', code: 'OVER_CAP', week: w,
        msg: 'Week of ' + RM.fmtShort(d) + ' looks like too much concurrent work (' + what + ')',
        items: cell.items
      });
    });

    var counts = { error: 0, warn: 0, info: 0 };
    Object.keys(byItem).forEach(function (id) {
      byItem[id].forEach(function (v) { counts[v.level] += 1; });
    });
    global.forEach(function (v) { counts[v.level] += 1; });

    return { byItem: byItem, global: global, capacity: cap, counts: counts };
  };

  // ------------------------------------------------------------ scheduling
  // Auto-schedule all unlocked items in non-bucket phases: topological order by
  // dependencies, earliest-start greedy placement under weekly capacity, bars
  // stretched across blackout weeks. Locked/bucket/done items keep their dates
  // and pre-consume capacity. Mutates a clone; returns { state, changed, notes }.
  RM.autoSchedule = function (inputState) {
    var state = RM.clone(inputState);
    var meta = state.meta;
    var notes = [];

    var phaseIdxById = {};
    state.phases.forEach(function (p, i) { phaseIdxById[p.id] = i; });

    var considered = [];
    var fixed = [];
    state.items.forEach(function (it, idx) {
      it._idx = idx;
      var phase = state.phases[phaseIdxById[it.phaseId]];
      // milestones are fixed dates: never moved, dependents plan around them
      if (!phase.bucket && !it.locked && !it.done && !it.milestone) considered.push(it);
      else if (it.startDay != null && it.durDays != null && !it.done) fixed.push(it);
    });

    // capacity ledger (capacity feature off → schedule by dependencies only)
    var teamTotal = state.meta.capacityEnabled ? state.team.length : 0;
    var HORIZON_WEEKS = meta.numWeeks + 104; // allow spill; UI can extend the grid
    var ledgerTotal = new Array(HORIZON_WEEKS);
    for (var w = 0; w < HORIZON_WEEKS; w++) ledgerTotal[w] = 0;

    var SLOTS = RM.slotsOf(meta);
    function occupy(it, startDay, durDays) {
      var wWt = RM.wipWeight(state, it);
      var w0 = Math.floor(startDay / SLOTS);
      var w1 = Math.floor((startDay + durDays - 1) / SLOTS);
      for (var wk = w0; wk <= w1 && wk < HORIZON_WEEKS; wk++) {
        if (RM.isBlackoutWeek(meta, wk)) continue;
        ledgerTotal[wk] += wWt;
      }
    }
    fixed.forEach(function (it) { occupy(it, it.startDay, it.durDays); });

    function fits(it, startDay, durDays) {
      if (teamTotal === 0) return true; // no roster -> no capacity constraint
      var wWt = RM.wipWeight(state, it);
      var w0 = Math.floor(startDay / SLOTS);
      var w1 = Math.floor((startDay + durDays - 1) / SLOTS);
      for (var wk = w0; wk <= w1; wk++) {
        if (wk >= HORIZON_WEEKS) return true;
        if (RM.isBlackoutWeek(meta, wk)) continue;
        var avail = RM.availForWeek(state, wk);
        if (ledgerTotal[wk] + wWt > avail.total + 1e-9) return false;
      }
      return true;
    }

    // topo order over considered items (deps to fixed items are satisfied by date)
    var consideredById = {};
    considered.forEach(function (it) { consideredById[it.id] = it; });
    var pendingDeps = {}; // id -> count of unscheduled considered deps
    var dependents = {};  // id -> [considered items depending on it]
    considered.forEach(function (it) {
      var deps = RM.resolveDeps(state, it).deps;
      var n = 0;
      deps.forEach(function (dep) {
        if (consideredById[dep.id]) {
          n += 1;
          (dependents[dep.id] = dependents[dep.id] || []).push(it);
        }
      });
      pendingDeps[it.id] = n;
    });

    function priority(a, b) {
      var pa = phaseIdxById[a.phaseId], pb = phaseIdxById[b.phaseId];
      if (pa !== pb) return pa - pb;
      return a._idx - b._idx;
    }

    var ready = considered.filter(function (it) { return pendingDeps[it.id] === 0; }).sort(priority);
    var remaining = considered.filter(function (it) { return pendingDeps[it.id] > 0; });
    var endById = {};
    fixed.concat(state.items.filter(function (it) { return it.done; })).forEach(function (it) {
      var e = RM.itemEnd(it);
      if (e != null) endById[it.id] = e;
    });

    var changed = 0;
    var maxDay = 0;

    // An item that can NEVER fit the roster — WIP weight above the PEAK
    // weekly availability (hours-based people-equivalents) — must not
    // trigger an endless capacity walk. Peaks are hours-aware, so a roster
    // of part-time people counts fractionally.
    var peakTotal = 0;
    if (teamTotal > 0) {
      for (var pw = 0; pw < HORIZON_WEEKS; pw++) {
        var pa = RM.availForWeek(state, pw);
        if (pa.total > peakTotal) peakTotal = pa.total;
        if (pw > meta.numWeeks && pa.total === peakTotal) break; // hours settle after overrides end
      }
    }
    function infeasible(it) {
      if (teamTotal === 0) return false;
      return RM.wipWeight(state, it) > peakTotal + 1e-9;
    }

    function place(it) {
      var deps = RM.resolveDeps(state, it).deps;
      var est = 0;
      deps.forEach(function (dep) {
        if (dep.done) return;
        var e = endById[dep.id] != null ? endById[dep.id] : RM.itemEnd(dep);
        if (e != null && e > est) est = e;
      });
      var work = RM.effortDays(state, it);
      var s = est;
      var guard = 0;
      var dur = RM.stretchSpan(meta, s, work);
      // with a roster set, the scheduler NEVER overallocates: an item the
      // roster can't absorb is left unscheduled instead of forced in
      function leaveUnscheduled(why) {
        notes.push('#' + it.num + ' (' + it.feature + ') ' + why + ' — left unscheduled.');
        if (it.startDay != null) changed += 1;
        it.startDay = null;
        it.durDays = null;
        it.riskDays = 0;
        (dependents[it.id] || []).forEach(function (child) {
          pendingDeps[child.id] -= 1;
          if (pendingDeps[child.id] === 0) {
            ready.push(child);
            ready.sort(priority);
            remaining = remaining.filter(function (r) { return r.id !== child.id; });
          }
        });
      }
      if (infeasible(it)) {
        leaveUnscheduled('needs more capacity than the roster ever has in a week');
        return;
      }
      while (!fits(it, s, dur) && guard < HORIZON_WEEKS * SLOTS) {
        s += 1;
        dur = RM.stretchSpan(meta, s, work);
        guard += 1;
      }
      if (guard >= HORIZON_WEEKS * SLOTS) {
        leaveUnscheduled('could not find a capacity-valid slot');
        return;
      }
      var riskSpan = RM.stretchSpan(meta, s + dur, RM.riskEffortDays(state, it));
      if (it.startDay !== s || it.durDays !== dur || (it.riskDays || 0) !== riskSpan) changed += 1;
      if (it.startDay != null) RM.shiftStories(it, s - it.startDay);
      it.startDay = s;
      it.durDays = dur;
      it.riskDays = riskSpan;
      occupy(it, s, dur); // the risk buffer is contingency — it books no capacity
      endById[it.id] = s + dur + riskSpan;
      if (s + dur + riskSpan > maxDay) maxDay = s + dur + riskSpan;
      (dependents[it.id] || []).forEach(function (child) {
        pendingDeps[child.id] -= 1;
        if (pendingDeps[child.id] === 0) {
          ready.push(child);
          ready.sort(priority);
          remaining = remaining.filter(function (r) { return r.id !== child.id; });
        }
      });
    }

    var guard2 = 0;
    while ((ready.length || remaining.length) && guard2 < 5000) {
      guard2 += 1;
      if (!ready.length) {
        // dependency cycle — break it deterministically at the lowest-priority entry
        remaining.sort(priority);
        var forced = remaining.shift();
        notes.push('#' + forced.num + ' is in a dependency cycle; scheduled by row order.');
        pendingDeps[forced.id] = 0;
        ready.push(forced);
      }
      var it = ready.shift();
      place(it);
    }

    var neededWeeks = Math.ceil(maxDay / SLOTS);
    if (neededWeeks > meta.numWeeks) {
      meta.numWeeks = neededWeeks;
      RM.syncEndDate(meta);
      notes.push('Timeline extended to ' + neededWeeks + ' weeks to fit the schedule.');
    }

    state.items.forEach(function (it) { delete it._idx; });
    return { state: state, changed: changed, notes: notes };
  };

  // Earliest dependency- and capacity-valid slot for one item, others fixed.
  // Returns { state, changed, note }; an item the roster can never absorb is
  // left untouched (note explains why) instead of being pushed off the grid.
  RM.snapEarliest = function (inputState, itemId) {
    var state = RM.clone(inputState);
    var it = RM.itemById(state, itemId);
    if (!it) return { state: state, changed: 0, note: null };
    var meta = state.meta;
    var teamTotal = state.meta.capacityEnabled ? state.team.length : 0;

    var snapWt = RM.wipWeight(state, it);
    if (teamTotal > 0 && snapWt > teamTotal) {
      return { state: state, changed: 0, note: 'Needs more focus than the roster of ' + teamTotal + ' can give — no slot can ever fit. Left unchanged.' };
    }

    var stash = { startDay: it.startDay, durDays: it.durDays, riskDays: it.riskDays || 0 };
    it.startDay = null; it.durDays = null; // free own capacity
    var deps = RM.resolveDeps(state, it).deps;
    var est = 0;
    deps.forEach(function (dep) {
      var e = RM.itemEnd(dep);
      if (!dep.done && e != null && e > est) est = e;
    });
    // a milestone occupies no working days — it snaps to the dependency floor
    var work = it.milestone ? 0 : RM.effortDays(state, it);

    var capData = RM.capacity(state);
    var snapS = RM.slotsOf(meta);
    var LIMIT = (meta.numWeeks + 104) * snapS;
    function fits(s, dur) {
      if (teamTotal === 0) return true;
      var w0 = Math.floor(s / snapS), w1 = Math.floor((s + dur - 1) / snapS);
      for (var wk = w0; wk <= w1; wk++) {
        if (RM.isBlackoutWeek(meta, wk)) continue;
        var cell = wk < capData.weeks.length ? capData.weeks[wk] : null;
        var availTotal = cell ? cell.cap : RM.availForWeek(state, wk).total;
        var demand = cell ? cell.demand : 0;
        if (demand + snapWt > availTotal + 1e-9) return false;
      }
      return true;
    }

    var s = est, guard = 0;
    var dur = RM.stretchSpan(meta, s, work);
    while (!fits(s, dur) && s + dur < LIMIT && guard < LIMIT) {
      s += 1; dur = RM.stretchSpan(meta, s, work); guard += 1;
    }
    if (!fits(s, dur)) {
      it.startDay = stash.startDay; it.durDays = stash.durDays; it.riskDays = stash.riskDays;
      return { state: state, changed: 0, note: 'No free slot found — left unchanged.' };
    }
    var riskSpan = RM.stretchSpan(meta, s + dur, RM.riskEffortDays(state, it));
    var note = null;
    var neededWeeks = Math.ceil((s + dur + riskSpan) / snapS);
    if (neededWeeks > meta.numWeeks) {
      meta.numWeeks = neededWeeks;
      RM.syncEndDate(meta);
      note = 'Timeline extended to ' + neededWeeks + ' weeks to fit it.';
    }
    var changed = (stash.startDay !== s || stash.durDays !== dur || stash.riskDays !== riskSpan) ? 1 : 0;
    if (stash.startDay != null) RM.shiftStories(it, s - stash.startDay);
    it.startDay = s;
    it.durDays = dur;
    it.riskDays = riskSpan;
    return { state: state, changed: changed, note: note };
  };

  // Renumber an item. An invalid or already-taken number falls back to the
  // next available one. Deps reference ids, so nothing else changes.
  RM.renumberItem = function (state, itemId, wanted) {
    var it = RM.itemById(state, itemId);
    if (!it) return null;
    var old = it.num;
    var n = parseInt(wanted, 10);
    var taken = {};
    state.items.forEach(function (x) { if (x.id !== itemId) taken[x.num] = true; });
    if (!isFinite(n) || n < 1 || taken[n]) n = RM.nextNum(state);
    if (n === old) return n;
    it.num = n;
    return n;
  };

  // Ripple move: cascade the dragged item's end-change through its dependents
  // ITERATIVELY. The dragged item has already been moved by the caller.
  // Forward (delta > 0): each dependent is pushed only as far as its deps'
  // buffered ends require, and each push chains to the next level in turn.
  // Backward (delta < 0): dependents follow by the same pull, clamped so they
  // never start before another dependency's buffered end (or day 0).
  // Locked/done items never move; since their ends don't change, the chain
  // naturally stops behind them. Returns the number of items moved.
  // a feature's scheduled stories ride along when the feature moves in time
  RM.shiftStories = function (it, delta) {
    if (!delta) return;
    (it.stories || []).forEach(function (st) {
      if (st.startDay != null) st.startDay = Math.max(0, st.startDay + delta);
    });
  };

  // ---- sprint view moves (pure; app.js wraps them in commits)
  // first working-day index of a numbered sprint (never before the timeline)
  RM.sprintStartDay = function (meta, num) {
    return Math.max(0, RM.sprintRange(meta, num).w0 * RM.slotsOf(meta));
  };
  // reorder one item in document order: before another item (adopting its
  // phase, like a row drop in Planning) or to the end of its own phase
  RM.reorderItem = function (state, itemId, beforeId) {
    var t = RM.itemById(state, itemId);
    if (!t || beforeId === itemId) return false;
    var before = beforeId ? RM.itemById(state, beforeId) : null;
    // one order key changes (placeItem mirrors the move in the array)
    RM.placeItem(state, itemId, before ? before.phaseId : t.phaseId, before ? before.id : null);
    delete t.holdPos;
    return true;
  };
  // Sprinting view drop: land the item in sprint `num` (null = unscheduled)
  // keeping its duration, ride its stories along, then reorder it before
  // `beforeId` (or to the end of its phase). Returns true when anything moved.
  RM.moveItemToSprint = function (state, itemId, num, beforeId) {
    var t = RM.itemById(state, itemId);
    if (!t) return false;
    var meta = state.meta;
    if (num == null) {
      if (t.startDay != null) { t.startDay = null; t.durDays = null; t.riskDays = 0; }
    } else {
      var day = RM.sprintStartDay(meta, num);
      var was = t.startDay;
      if (t.durDays == null) {
        t.durDays = t.milestone ? 1 : RM.stretchSpan(meta, day, RM.effortDays(state, t) || RM.sprintDays(meta));
      }
      t.startDay = day;
      if (was != null) RM.shiftStories(t, day - was);
    }
    RM.reorderItem(state, itemId, beforeId);
    return true;
  };
  // story variant: a story gets its own timeline in sprint `num` (one sprint
  // unless it already has a span), or loses it (num = null); reorders only
  // inside its feature's story list (before `beforeStId`, else last)
  RM.moveStoryToSprint = function (state, itemId, stId, num, beforeStId) {
    var it = RM.itemById(state, itemId);
    if (!it) return false;
    var st = (it.stories || []).filter(function (x) { return x.id === stId; })[0];
    if (!st) return false;
    if (num == null) { st.startDay = null; st.durDays = null; }
    else {
      st.startDay = RM.sprintStartDay(state.meta, num);
      if (st.durDays == null) st.durDays = RM.stretchSpan(state.meta, st.startDay, RM.storyEffortDays(state, st));
    }
    // before another story (else last): its order key follows the row
    if (beforeStId !== stId) placeInList(it.stories, st, beforeStId || null, function () { return true; });
    return true;
  };

  // Rename a role everywhere it appears: the role list, people, items and
  // the rate card. Returns false when the new name is empty or taken.
  RM.renameRole = function (state, oldName, newName) {
    newName = String(newName || '').trim();
    if (!newName || newName === oldName) return false;
    if (state.teamTypes.indexOf(newName) !== -1) return false;
    var i = state.teamTypes.indexOf(oldName);
    if (i === -1) return false;
    state.teamTypes[i] = newName;
    state.team.forEach(function (m) { if (m.type === oldName) m.type = newName; });
    state.items.forEach(function (it) { if (it.teamType === oldName) it.teamType = newName; });
    if (state.meta.rateCard && state.meta.rateCard[oldName]) {
      state.meta.rateCard[newName] = state.meta.rateCard[oldName];
      delete state.meta.rateCard[oldName];
    }
    return true;
  };

  // Remove a role everywhere: people and features that had it are left with
  // no role (empty), and its rate card entry goes with it.
  RM.removeRole = function (state, name) {
    var i = state.teamTypes.indexOf(name);
    if (i === -1) return false;
    state.teamTypes.splice(i, 1);
    state.team.forEach(function (m) { if (m.type === name) m.type = ''; });
    state.items.forEach(function (it) { if (it.teamType === name) it.teamType = ''; });
    if (state.meta.rateCard) delete state.meta.rateCard[name];
    return true;
  };

  // opts.rigid (⌘-drag): every transitive dependent moves by the SAME delta
  // as the dragged item — no slack absorption, pull and push alike.
  // Milestones ride along too (the drag is an explicit date change); locked
  // and done items stay put either way.
  RM.shiftDependents = function (state, itemId, delta, opts) {
    if (!delta) return 0;
    var rigid = !!(opts && opts.rigid);
    var childrenBy = {};
    RM.depEdges(state).forEach(function (e) {
      (childrenBy[e[0].id] = childrenBy[e[0].id] || []).push(e[1]);
    });
    var moved = {};
    var pulled = {};
    var q = [{ id: itemId, d: delta }];
    var guard = 0;
    while (q.length && guard++ < 10000) {
      var cur = q.shift();
      /* eslint-disable no-loop-func */
      (childrenBy[cur.id] || []).forEach(function (ch) {
        if (ch.startDay == null || ch.locked || ch.done) return;
        if (rigid) {
          if (pulled[ch.id]) return; // visited once — rigid moves are absolute
          pulled[ch.id] = true;
          var rt = Math.max(0, ch.startDay + delta);
          var rApplied = rt - ch.startDay;
          if (rApplied) {
            ch.startDay = rt;
            RM.shiftStories(ch, rApplied);
            moved[ch.id] = true;
          }
          q.push({ id: ch.id, d: delta });
          return;
        }
        if (ch.milestone) return;
        // never start before any scheduled dependency's buffered end
        var floor = 0;
        ch.deps.forEach(function (id) {
          var dp = RM.itemById(state, id);
          if (dp && dp.startDay != null && !dp.done) floor = Math.max(floor, RM.itemEnd(dp));
        });
        var target;
        if (cur.d < 0 && !pulled[ch.id]) {
          pulled[ch.id] = true;
          target = Math.max(0, ch.startDay + cur.d, floor);
        } else {
          target = Math.max(ch.startDay, floor);
        }
        var applied = target - ch.startDay;
        if (!applied) return;
        ch.startDay = target;
        RM.shiftStories(ch, applied);
        moved[ch.id] = true;
        q.push({ id: ch.id, d: applied });
      });
    }
    return Object.keys(moved).length;
  };

  // Stable-reorder items inside each phase by start day (unscheduled items
  // sink to the end of their phase, keeping their relative order).
  RM.sortItemsByStart = function (state) {
    var out = [];
    state.phases.forEach(function (p) {
      var mine = state.items
        .map(function (it, i) { return { it: it, i: i }; })
        .filter(function (x) { return x.it.phaseId === p.id; });
      // holdPos: a freshly-inserted row stays put until it has a start date
      // (the flag clears itself once one is set)
      var held = [];
      mine = mine.filter(function (x, pi) {
        if (x.it.holdPos && x.it.startDay == null) { held.push({ it: x.it, pi: pi }); return false; }
        if (x.it.holdPos) delete x.it.holdPos;
        return true;
      });
      mine.sort(function (a, b) {
        var sa = a.it.startDay != null ? a.it.startDay : Infinity;
        var sb = b.it.startDay != null ? b.it.startDay : Infinity;
        if (sa !== sb) return sa - sb;
        return a.i - b.i; // stable
      });
      held.forEach(function (h) {
        mine.splice(Math.min(h.pi, mine.length), 0, { it: h.it });
      });
      mine.forEach(function (x) { out.push(x.it); });
    });
    // anything with an unknown phase (shouldn't exist post-normalize) tags along
    state.items.forEach(function (it) { if (out.indexOf(it) === -1) out.push(it); });
    state.items = out;
    return state;
  };

  // Items in display order, as a NEW array: by start day inside each phase
  // when auto-order is on (sortItemsByStart's ordering, run on shallow
  // copies so holdPos flags on the real items are left alone), else by
  // order key. Never mutates state.items.
  RM.viewItems = function (state, opts) {
    if (!(opts && opts.autoOrder)) return RM.sortByOrder(state.items.slice());
    var byId = {};
    state.items.forEach(function (it) { byId[it.id] = it; });
    var tmp = { phases: state.phases, items: state.items.map(function (it) {
      return { id: it.id, phaseId: it.phaseId, startDay: it.startDay, holdPos: it.holdPos };
    }) };
    RM.sortItemsByStart(tmp);
    return tmp.items.map(function (x) { return byId[x.id]; });
  };

  // ------------------------------------------------------------ budgeting
  // Planned vs ACTUAL hours for one role-week: actual clips the planned hours
  // (override or the 40 h default) to the week's workable days — 8 h per
  // non-holiday day. Actual is the basis for totals and cost.
  RM.roleWeekHours = function (state, m, w) {
    var iso = RM.fmtISO(RM.weekStartDate(state.meta, w));
    var planned = m.weekHours[iso] != null ? m.weekHours[iso] : RM.weekHoursOf(state.meta);
    var workable = (RM.slotsOf(state.meta) - RM.holidaysInWeek(state.meta, w)) * RM.hoursPerDay(state.meta);
    return { iso: iso, planned: planned, actual: Math.min(planned, workable) };
  };
  // Per-role ACTUAL hours over the whole project.
  RM.roleTotalHours = function (state, m) {
    var total = 0;
    for (var w = 0; w < state.meta.numWeeks; w++) total += RM.roleWeekHours(state, m, w).actual;
    return total;
  };
  // ---- rate card: per-role default hourly rate/cost (meta.rateCard[role] =
  // { rate, cost }). A person inherits their role's numbers unless they carry
  // an explicit override (0/empty = inherit).
  RM.rateCardFor = function (state, role) {
    var rc = state.meta.rateCard;
    return (role && rc && rc[role]) || null;
  };
  // display label for a member — name first, then role/rate-card fallbacks
  // (both name and rate-card assignment are optional)
  RM.memberLabel = function (m) {
    return (m && (m.name || m.role || m.type)) || 'Unnamed';
  };
  RM.memberRate = function (state, m) {
    if (m.rate > 0) return m.rate;
    var rc = RM.rateCardFor(state, m.type);
    return rc && isFinite(+rc.rate) && +rc.rate > 0 ? +rc.rate : 0;
  };
  RM.memberCost = function (state, m) {
    if (m.cost > 0) return m.cost;
    var rc = RM.rateCardFor(state, m.type);
    return rc && isFinite(+rc.cost) && +rc.cost > 0 ? +rc.cost : 0;
  };

  // People can sit on several workstreams; the primary (first) drives colors.
  RM.memberWorkstreams = function (m) {
    return Array.isArray(m.workstreams) ? m.workstreams : (m.workstream ? [m.workstream] : []);
  };
  RM.setMemberWorkstreams = function (m, list) {
    var out = [];
    (list || []).forEach(function (w) {
      w = String(w || '').trim();
      if (w && out.indexOf(w) === -1) out.push(w);
    });
    m.workstreams = out;
    m.workstream = out[0] || '';
  };

  // Hard deadline as a day index (deadlines live as calendar dates); null if unset.
  RM.deadlineDay = function (meta, it) {
    if (!it || !it.deadline) return null;
    return RM.dateToDay(meta, RM.parseISO(it.deadline));
  };
  // Does the item's scheduled span run past its deadline?
  RM.pastDeadline = function (meta, it) {
    var dl = RM.deadlineDay(meta, it);
    if (dl == null || it.startDay == null) return false;
    var lastDay = it.milestone ? it.startDay
      : it.startDay + Math.max(1, (it.durDays || 1) + (it.riskDays || 0)) - 1;
    return lastDay > dl;
  };

  // Margin: share of the bill rate kept after hourly cost, in %; null if no rate.
  RM.roleMargin = function (state, m) {
    var rate = RM.memberRate(state, m);
    if (!rate) return null;
    return (rate - RM.memberCost(state, m)) / rate * 100;
  };

  // ---- fixed & recurring costs (state.costs). kind: 'fixed' hits once on
  // startDay; 'weekly'/'monthly' repeat from startDay through endDay
  // (endDay null = the end of the timeline).
  RM.costOccurrences = function (state, c) {
    var meta = state.meta;
    var last = RM.numDays(meta) - 1;
    var start = Math.max(0, c.startDay || 0);
    if (c.kind === 'fixed') {
      return start <= last ? [{ day: start, amount: c.amount }] : [];
    }
    var end = c.endDay != null ? Math.min(c.endDay, last) : last;
    var out = [];
    if (c.kind === 'weekly') {
      for (var d = start; d <= end; d += RM.slotsOf(state.meta)) out.push({ day: d, amount: c.amount });
      return out;
    }
    // monthly: same day-of-month as the start date, until end
    var dt = RM.dayToDate(meta, start);
    var guard = 0;
    while (guard++ < 240) {
      var day = RM.dateToDay(meta, dt);
      if (day == null || day > end) break;
      if (day >= start) out.push({ day: day, amount: c.amount });
      dt = new Date(dt.getTime());
      dt.setUTCMonth(dt.getUTCMonth() + 1);
    }
    return out;
  };
  RM.costTotal = function (state, c) {
    return RM.costOccurrences(state, c).reduce(function (a, o) { return a + o.amount; }, 0);
  };
  RM.costsTotal = function (state) {
    return (state.costs || []).reduce(function (a, c) { return a + RM.costTotal(state, c); }, 0);
  };
  // Average hourly COST of roster roles with the given team type; roles with
  // no cost set are ignored; falls back to the blended roster average.
  RM.avgCostRate = function (state, teamType) {
    function eff(m) { return RM.memberCost(state, m); }
    var pool = state.team.filter(function (m) { return eff(m) > 0 && (!teamType || m.type === teamType); });
    if (!pool.length && teamType) pool = state.team.filter(function (m) { return eff(m) > 0; });
    if (!pool.length) {
      // no roster costs: the rate card itself can price the work
      var rc = teamType ? RM.rateCardFor(state, teamType) : null;
      return rc && +rc.cost > 0 ? +rc.cost : 0;
    }
    var s = 0;
    pool.forEach(function (m) { s += eff(m); });
    return s / pool.length;
  };
  // Item effort: scheduled items use their painted working days, unscheduled
  // ones their size estimate. Hours = days × the project's hours-per-day;
  // cost prices those hours at the avg cost rate for the item's role.
  RM.itemEffortInfo = function (state, it) {
    var days = it.startDay != null && it.durDays != null
      ? RM.workInSpan(state.meta, it.startDay, it.durDays)
      : (it.size ? RM.sizeDays(state, it.size) : (it.durDays || 0));
    var hours = days * RM.hoursPerDay(state.meta);
    return { days: days, hours: hours, cost: hours * RM.avgCostRate(state, it.teamType) };
  };
  // Cost/effort rollup for the Reports panel. mode: 'workstream' | 'phase' |
  // 'phase-ws'. Workstream mode also rolls up the roster's own spend
  // (role hours × hourly cost) per workstream.
  RM.costReport = function (state, mode) {
    var rows = {}, order = [];
    function bucket(key) {
      if (!rows[key]) { rows[key] = { key: key, items: 0, days: 0, hours: 0, cost: 0 }; order.push(key); }
      return rows[key];
    }
    var phaseName = {};
    state.phases.forEach(function (p) { phaseName[p.id] = p.name; });
    state.items.forEach(function (it) {
      var info = RM.itemEffortInfo(state, it);
      var key = mode === 'phase' ? phaseName[it.phaseId]
        : mode === 'phase-ws' ? phaseName[it.phaseId] + ' · ' + (it.workstream || RM.defaultWsName(state))
        : (it.workstream || RM.defaultWsName(state));
      var b = bucket(key);
      b.items += 1; b.days += info.days; b.hours += info.hours; b.cost += info.cost;
    });
    if (mode === 'workstream') {
      state.team.forEach(function (m) {
        // a person on several workstreams splits their hours/cost evenly
        var wss = RM.memberWorkstreams(m);
        if (!wss.length) wss = [RM.defaultWsName(state)];
        var h = RM.roleTotalHours(state, m) / wss.length;
        var c = h * RM.memberCost(state, m);
        wss.forEach(function (w) {
          var b = bucket(w);
          b.roleHours = (b.roleHours || 0) + h;
          b.roleCost = (b.roleCost || 0) + c;
        });
      });
    }
    var out = order.map(function (k) { return rows[k]; });
    var total = { key: 'Total', items: 0, days: 0, hours: 0, cost: 0, roleHours: 0, roleCost: 0 };
    out.forEach(function (r) {
      total.items += r.items; total.days += r.days; total.hours += r.hours; total.cost += r.cost;
      total.roleHours += r.roleHours || 0; total.roleCost += r.roleCost || 0;
    });
    return { rows: out, total: total };
  };

  // Phase window in working-day indices: user-pinned startDay/endDay win;
  // whichever side is unset auto-derives from the phase's scheduled items.
  RM.phaseSpan = function (state, phase) {
    var lo = null, hi = null;
    RM.itemsInPhase(state, phase.id).forEach(function (it) {
      if (it.startDay == null || it.durDays == null) return;
      if (lo == null || it.startDay < lo) lo = it.startDay;
      var e = RM.itemEnd(it);
      if (hi == null || e > hi) hi = e;
    });
    if (phase.startDay != null) lo = phase.startDay;
    if (phase.endDay != null) hi = phase.endDay;
    if (lo == null || hi == null || hi <= lo) return null;
    return { lo: lo, hi: hi };
  };

  // ------------------------------------------------------------ misc helpers
  // Suggested finish stats for the header: last scheduled day + item count.
  RM.scheduleStats = function (state) {
    var last = null, scheduled = 0, unscheduled = 0;
    state.items.forEach(function (it) {
      if (it.startDay != null && it.durDays != null) {
        scheduled += 1;
        var e = RM.itemEnd(it);
        if (last == null || e > last) last = e;
      } else unscheduled += 1;
    });
    return { lastDay: last, scheduled: scheduled, unscheduled: unscheduled };
  };

  if (typeof module !== 'undefined' && module.exports) module.exports = RM;
  root.RM = RM;
})(typeof window !== 'undefined' ? window : globalThis);
