/* Headway UI. Vanilla JS, no build step; state logic lives in core.js. */
(function () {
  'use strict';

  var LS_KEY = 'headway-v1';
  var UI_KEY = 'headway-ui-v1';

  // ------------------------------------------------------------ theme
  // Personal, per-machine — deliberately NOT part of uiSnapshot(), so a
  // shared .xlsx never overrides another machine's theme. index.html stamps
  // data-theme before first paint; this manager owns it from then on.
  var THEME_KEY = 'headway-theme-v1';
  var themePref = 'system'; // system | light | dark
  try { themePref = localStorage.getItem(THEME_KEY) || 'system'; } catch (e) { /* storage optional */ }
  if (['system', 'light', 'dark'].indexOf(themePref) === -1) themePref = 'system';
  var themeMedia = window.matchMedia ? window.matchMedia('(prefers-color-scheme: dark)') : null;
  function darkActive() {
    return themePref === 'dark' || (themePref === 'system' && !!(themeMedia && themeMedia.matches));
  }
  function applyTheme() {
    document.documentElement.dataset.theme = darkActive() ? 'dark' : 'light';
  }
  function setTheme(pref) {
    themePref = pref;
    try { localStorage.setItem(THEME_KEY, pref); } catch (e) { /* storage optional */ }
    applyTheme();
    if (state) render(); // theme-aware paints (resource heat cells) refresh
  }
  if (themeMedia && themeMedia.addEventListener) {
    themeMedia.addEventListener('change', function () {
      if (themePref !== 'system') return;
      applyTheme();
      if (state) render();
    });
  }
  applyTheme();
  var THEME_CHOICES = [['system', 'System theme', 'monitor'], ['light', 'Light theme', 'sun'], ['dark', 'Dark theme', 'moon']];
  function themeMenuItems() {
    return THEME_CHOICES.map(function (t) {
      return { icon: t[2], label: t[1], checked: themePref === t[0], fn: function () { setTheme(t[0]); } };
    });
  }

  // ------------------------------------------------------------ app state
  var state;                 // document state (undo-tracked)
  // a standalone HTML export boots with its document inlined: every tab is
  // navigable, nothing edits, and Setup keeps only the theme
  var readOnly = !!(window.HEADWAY_VIEW && window.HEADWAY_VIEW.doc);
  function viewOnlyToast() { toast('This is a view-only copy \u2014 edits are off'); }
  var validation;            // RM.validate cache
  var undoStack = [], redoStack = [];
  var selectedId = null;
  var reloadingDoc = false;  // true while a disk reload re-renders (focus is put back afterwards)
  var expanded = {};         // itemId -> true (stories open)
  var weekPx = 28;
  var view = 'planning';     // planning (timeline) | scoping (spreadsheet)
  var depsMode = 'on';       // on: selected item's explicit deps + violations + critical path | none
  var showCrit = true;       // orange critical-path highlight (bars + arrows)
  var showCap = true;        // weekly capacity row in the planning header
  var repCollapsed = true;   // bottom Reports drawer starts tucked away
  var leftWBudget = 806;     // frozen left-pane width, budgeting view
  var repMode = 'workstream'; // reports grouping: workstream | phase | phase-ws
  var panelW = 372;          // right edit-panel width (resizable)
  var panelOpen = true;      // right panel is persistent on Planning; collapsible
  var leftCollapsed = false; // frozen left pane folded to a slim strip ([ toggles)
  var leftWPlan = 538;       // frozen left-pane width, planning view
  var leftWScope = 538;      // …and scoping view (independently resizable)
  var groupWs = false;       // sub-group rows by workstream inside each phase
  var groupEpic = false;     // …and/or by epic (nested under workstream)
  var exportPrefs = null;    // last-used Export dialog settings (fmt, split, …)
  var jiraPrefs = null;      // last-used Export Jira CSV dialog settings
  var SNAP_MODES = ['day', 'week', 'sprint'];
  var snapFeat = 'week';     // feature drag/place snap: 'day' | 'week' | 'sprint'
  var snapStory = 'sprint';  // …and the story snap, chosen separately
  var detailMode = 'feature'; // Scoping/Planning row detail: phase | feature | story
  var buColW = {};           // budgeting column width overrides (key -> px)
  var buColOrder = null;     // budgeting column order (array of keys; null = default)
  var buColHide = {};        // budgeting columns hidden (key -> true)
  var plColW = {};           // planning left-pane column width overrides (key -> px)
  var plColOrder = null;     // planning left-pane column order
  var plColHide = {};        // planning left-pane columns hidden
  var autoOrder = true;      // after move/resize, reorder rows by start day (stable)
  var setupTab = 'project';    // active section in the Setup view
  var filterText = '';       // planning/scoping row filter (⌘F); transient
  var multiSel = null;       // multi-selection: item ids (null = single-select mode; selectedId stays the anchor)
  var docSaved = false;      // doc matches its last save/open (Save button shows ✓)
  var sessionEdited = false; // an actual edit happened this session (guards never nag a doc that was only opened)
  var autoSave = true;       // desktop: write to the open file after each change
  var resPanelH = 150;       // resources panel height (px)
  var resCollapsed = false;  // resources section collapsed
  var drag = null;           // active drag descriptor
  var lastExport = null;
  var critCache = null;      // RM.criticalPath result, refreshed in render()

  var $ = function (sel, el) { return (el || document).querySelector(sel); };
  var $$ = function (sel, el) { return Array.prototype.slice.call((el || document).querySelectorAll(sel)); };
  var board = $('#board'), grid = $('#grid'), rowsEl = $('#rows');

  function esc(s) {
    return String(s == null ? '' : s)
      .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;')
      .replace(/"/g, '&quot;');
  }
  function SPW() { return RM.slotsOf(state.meta); } // slots (working days) per week
  function dayPx() { return weekPx / SPW(); }
  function leftW() {
    return parseFloat(getComputedStyle(document.documentElement).getPropertyValue('--left-w')) || 478;
  }
  // scrollLeft that puts lane-space x a third of the way into the VISIBLE lane
  // strip (the frozen left column eats leftW() px of the viewport).
  function scrollLaneTo(x) {
    var laneVis = Math.max(80, board.clientWidth - leftW());
    board.scrollLeft = Math.max(0, x - Math.max(24, laneVis / 3));
  }
  function phaseOf(it) {
    for (var i = 0; i < state.phases.length; i++) if (state.phases[i].id === it.phaseId) return state.phases[i];
    return state.phases[0];
  }
  function isScheduled(it) { return it.startDay != null && it.durDays != null; }

  // ------------------------------------------------------------ persistence
  // one place defines "the UI prefs" — written to localStorage on every
  // commit AND carried in the .xlsx (_RoadmapTool sheet) so a saved file
  // restores the exact browser state on any machine
  function uiSnapshot() {
    return { weekPx: weekPx, view: view, depsMode: depsMode, groupWs: groupWs, groupEpic: groupEpic, resCollapsed: resCollapsed, snapFeat: snapFeat, snapStory: snapStory, autoOrder: autoOrder, showCrit: showCrit, showCap: showCap, scopeColW: scopeColW, resPanelH: resPanelH, panelSec: panelSec, leftWPlan: leftWPlan, leftWScope: leftWScope, leftWBudget: leftWBudget, panelW: panelW, expanded: expanded, repCollapsed: repCollapsed, repMode: repMode, autoSave: autoSave, setupTab: setupTab, panelOpen: panelOpen, prioGroup: prioGroup, prioFields: prioFields, prioChipHide: prioChipHide, prioHideUnset: prioHideUnset, prioSort: prioSort, detailMode: detailMode, buColW: buColW, buColOrder: buColOrder, buColHide: buColHide, plColW: plColW, plColOrder: plColOrder, plColHide: plColHide, exportPrefs: exportPrefs, jiraPrefs: jiraPrefs, sprLevel: sprLevel, prioLevel: prioLevel, prioStoryCol: prioStoryCol, prioFeatCol: prioFeatCol, leftCollapsed: leftCollapsed, colorBy: colorBy };
  }
  var COLOR_MODES = [['workstream', 'Workstream'], ['epic', 'Epic'], ['assignee', 'Assignee'], ['priority', 'Priority'], ['type', 'Item type']];
  var colorBy = 'workstream'; // what bar colors follow: 'workstream' | 'epic' | 'assignee' | 'priority' | 'type'
  function setColorBy(mode) {
    colorBy = RM.COLOR_MODES.indexOf(mode) !== -1 ? mode : 'workstream';
    RM.setColorMode(colorBy);
  }
  var prioGroup = 'none';   // Prioritizing view swimlanes: 'none' | 'ws' | 'epic'
  var sprLevel = 'feature'; // Sprinting view rows: 'feature' | 'story'
  var prioLevel = 'feature'; // Prioritizing cards: 'feature' | 'story' (a board of story cards)
  var prioStoryCol = 'priority'; // Story-level columns: 'priority' | 'size' | 'risk'
  var prioFeatCol = 'phase';     // Feature-level columns: 'phase' | 'priority' | 'size' | 'risk'
  var prioFields = [];      // scope-column keys shown on prioritizing cards (compact by default)
  var prioChipHide = [];    // card chips (size / priority / risk / dur / epic / ws) the user hid
  var prioHideUnset = false; // Prioritizing: hide the Unset column (features/stories without a value for the column field)
  var prioSort = 'priority'; // Prioritizing order: 'priority' | 'doc' | 'title' | 'size'
  var prioFEpic = null;     // Prioritizing epic filter: null = all, '' = no epic, else the epic
  var prioFWs = null;       // Prioritizing workstream filter: null = all, '' = default, else the stream
  var prioFPhase = null;    // Prioritizing phase filter: null = all, else a phase id
  var sprFPhase = null;     // Sprinting phase filter: null = all, else a phase id
  var sprFEpic = null;      // Sprinting epic filter: null = all, '' = no epic, else the epic
  var sprFWs = null;        // Sprinting workstream filter: null = all, '' = default, else the stream
  var uiExpandedLoaded = false; // boot skips the auto-expand default when true

  function applyUi(ui) {
    if (!ui) return;
    weekPx = ui.weekPx || 28;
    view = ['scoping', 'setup', 'budget', 'reports', 'history', 'prio', 'sprints'].indexOf(ui.view) !== -1 ? ui.view : 'planning';
    detailMode = ui.detailMode === 'story' ? 'story' : 'feature';
    buColW = ui.buColW && typeof ui.buColW === 'object' ? ui.buColW : {};
    buColOrder = Array.isArray(ui.buColOrder) ? ui.buColOrder : null;
    buColHide = ui.buColHide && typeof ui.buColHide === 'object' ? ui.buColHide : {};
    plColW = ui.plColW && typeof ui.plColW === 'object' ? ui.plColW : {};
    plColOrder = Array.isArray(ui.plColOrder) ? ui.plColOrder : null;
    plColHide = ui.plColHide && typeof ui.plColHide === 'object' ? ui.plColHide : {};
    depsMode = ui.depsMode === 'none' ? 'none' : 'on';
    groupWs = ui.groupWs != null ? !!ui.groupWs : ui.groupBy === 'ws';
    groupEpic = ui.groupEpic != null ? !!ui.groupEpic : (ui.groupBy === 'epic' || !!ui.groupByEpic);
    exportPrefs = ui.exportPrefs && typeof ui.exportPrefs === 'object' ? ui.exportPrefs : null;
    jiraPrefs = ui.jiraPrefs && typeof ui.jiraPrefs === 'object' ? ui.jiraPrefs : null;
    resCollapsed = !!ui.resCollapsed;
    // pre-1.1 snapshots stored one numeric snap (1 / 5 / 10 days) for everything
    var legacySnap = ui.snapDays === 1 ? 'day' : ui.snapDays === 10 ? 'sprint' : null;
    snapFeat = SNAP_MODES.indexOf(ui.snapFeat) !== -1 ? ui.snapFeat : (legacySnap || 'week');
    snapStory = SNAP_MODES.indexOf(ui.snapStory) !== -1 ? ui.snapStory : 'sprint';
    scopeColW = ui.scopeColW && typeof ui.scopeColW === 'object' ? ui.scopeColW : {};
    autoOrder = ui.autoOrder !== false; // default true
    showCrit = ui.showCrit !== false;   // default true
    showCap = ui.showCap !== false;     // default true
    resPanelH = ui.resPanelH > 40 ? ui.resPanelH : 150;
    panelSec = ui.panelSec && typeof ui.panelSec === 'object' ? ui.panelSec : {};
    leftWPlan = ui.leftWPlan > 200 ? ui.leftWPlan : 538;
    leftWScope = ui.leftWScope > 200 ? ui.leftWScope : 538;
    // 660 was the pre-widened-columns default; bump those snapshots to 696
    // 660/696 were earlier defaults — stored values matching them jump to the
    // new default (which fits the added Role column)
    leftWBudget = ui.leftWBudget > 300 && ui.leftWBudget !== 660 && ui.leftWBudget !== 696 ? ui.leftWBudget : 806;
    panelW = ui.panelW > 280 ? ui.panelW : 372;
    repCollapsed = ui.repCollapsed !== false; // default collapsed
    repMode = ['workstream', 'phase', 'phase-ws'].indexOf(ui.repMode) !== -1 ? ui.repMode : 'workstream';
    autoSave = ui.autoSave !== false;   // default true (desktop writes to the open file)
    setupTab = normSetupTab(typeof ui.setupTab === 'string' ? ui.setupTab : 'project');
    panelOpen = ui.panelOpen !== false; // panel is persistent by default
    leftCollapsed = ui.leftCollapsed === true;
    prioGroup = ['ws', 'epic'].indexOf(ui.prioGroup) !== -1 ? ui.prioGroup : 'none';
    setColorBy(ui.colorBy);
    sprLevel = ui.sprLevel === 'story' ? 'story' : 'feature';
    prioLevel = ui.prioLevel === 'story' ? 'story' : 'feature';
    prioStoryCol = ['priority', 'size', 'risk'].indexOf(ui.prioStoryCol) !== -1 ? ui.prioStoryCol : 'priority';
    prioFeatCol = ['phase', 'priority', 'size', 'risk'].indexOf(ui.prioFeatCol) !== -1 ? ui.prioFeatCol : 'phase';
    if (Array.isArray(ui.prioFields)) prioFields = ui.prioFields.map(String);
    if (Array.isArray(ui.prioChipHide)) prioChipHide = ui.prioChipHide.map(String).filter(function (k) { return PR_CHIPS[k]; });
    prioHideUnset = ui.prioHideUnset === true;
    prioSort = ['doc', 'title', 'size'].indexOf(ui.prioSort) !== -1 ? ui.prioSort : 'priority';
    if (ui.expanded && typeof ui.expanded === 'object') {
      expanded = ui.expanded;
      uiExpandedLoaded = true;
    }
  }

  var localSaveBroken = false;
  function saveLocal() {
    if (readOnly) return; // the exported copy never writes the viewer's storage
    if (wz) return; // the new-project wizard's draft is never stored
    try {
      localStorage.setItem(LS_KEY, JSON.stringify(state));
      localStorage.setItem(UI_KEY, JSON.stringify(uiSnapshot()));
      localSaveBroken = false;
    } catch (e) {
      // storage full or blocked (some browsers restrict file:// storage) —
      // say so instead of pretending the edit is safe
      localSaveBroken = true;
    }
    var st = $('#saveStatus');
    if (st) {
      if (localSaveBroken) {
        st.textContent = 'local save unavailable — use Save to keep your work';
        st.classList.add('dirty');
      } else {
        st.textContent = 'saved locally' + (lastExport ? ' · exported ' + lastExport : '');
        st.classList.remove('dirty');
      }
    }
  }

  // people now take their role's capacity type: say so when that moved anyone
  function capTypeToast(n) {
    if (n > 0) toast(n + (n === 1 ? ' person now takes' : ' people now take') + ' the capacity type of their role \u2014 check Setup \u203a Scheduling');
  }
  var localCapTypeChanges = 0; // set by loadLocal, announced once the app is up
  function loadLocal() {
    try {
      applyUi(JSON.parse(localStorage.getItem(UI_KEY) || 'null'));
      var raw = JSON.parse(localStorage.getItem(LS_KEY) || 'null');
      if (raw && raw.items) {
        var restored = RM.normalizeState(raw);
        localCapTypeChanges = RM.lastCapTypeChanges || 0;
        return restored;
      }
    } catch (e) { /* corrupted local copy — fall back to seed */ }
    return null;
  }

  // ------------------------------------------------------------ commits
  // every commit also lands in the document's version history (state.history,
  // persisted with the doc): who made the change, what it was, and when.
  // Rapid same-kind edits by the same person coalesce into one entry.
  var USER_KEY = 'headway-user-v1';
  var HISTORY_COALESCE_MS = 5 * 60 * 1000;
  var sessionStart = Date.now(); // for stamping this session's anonymous edits
  var namePromptShown = false;
  function userName() {
    try { return (localStorage.getItem(USER_KEY) || '').trim(); } catch (e) { return ''; }
  }
  function setUserName(v) {
    try { localStorage.setItem(USER_KEY, String(v || '').trim()); } catch (e) { /* storage optional */ }
  }
  // ---- semantic diff between two document states, for Version History.
  // Ops are [category, field label, old value, new value]; categories map to
  // the tabs on the Version History page. Values are short display strings —
  // this is a human-readable audit trail, not a machine patch.
  var VH_CATS = { timeline: 'Timeline', scope: 'Scope', status: 'Status', budget: 'Team & Costs', setup: 'Setup' };
  function diffStates(a, b) {
    var ops = [];
    var tl = []; // schedule moves, machine-readable, for the visual timeline diff
    var meta = (b && b.meta) || (a && a.meta);
    function push(cat, label, oldV, newV) {
      oldV = oldV == null ? '' : String(oldV);
      newV = newV == null ? '' : String(newV);
      if (oldV !== newV) ops.push([cat, label, oldV, newV]);
    }
    function txt(v) {
      var t = RM.htmlToText(String(v == null ? '' : v)).trim();
      return t.length > 160 ? t.slice(0, 159) + '…' : t;
    }
    function dDate(day) { return day == null ? '' : RM.fmtShortYear(RM.dayToDate(meta, day)); }
    function dWeeks(days) { return days == null ? '' : (Math.round(days / RM.slotsOf(meta) * 100) / 100) + 'w'; }
    function byId(list) {
      var m = {};
      (list || []).forEach(function (x) { if (x && x.id) m[x.id] = x; });
      return m;
    }
    function names(st, ids) {
      return (ids || []).map(function (id) {
        var m2 = null;
        (st.team || []).forEach(function (x) { if (x.id === id) m2 = x; });
        return m2 ? RM.memberLabel(m2) : '?';
      }).join(', ');
    }
    // ---- items (features / milestones) + their stories
    var ia = byId(a.items), ib = byId(b.items);
    function span(x) { return x.durDays == null ? null : x.durDays + (x.riskDays || 0); }
    function tlPush(it, p) {
      // p null = added, it null = removed; only rows whose schedule differs
      var s0 = p ? p.startDay : null, d0 = p ? span(p) : null;
      var s1 = it ? it.startDay : null, d1 = it ? span(it) : null;
      if (s0 === s1 && d0 === d1) return;
      if (s0 == null && s1 == null) return;
      var ref = it || p;
      tl.push({ id: ref.id, n: ref.num, f: shorten(ref.feature || '(untitled)', 40),
        ms: ref.milestone ? 1 : 0, s0: s0, d0: d0, s1: s1, d1: d1 });
    }
    (b.items || []).forEach(function (it) {
      var lbl = '#' + it.num + ' ' + shorten(it.feature || '(untitled)', 28);
      var p = ia[it.id];
      tlPush(it, p || null);
      if (!p) { push('scope', lbl, '', it.milestone ? 'Milestone added' : 'Feature added'); return; }
      push('scope', lbl + ' — Title', p.feature, it.feature);
      function phName(st, pid) {
        var ph = null;
        (st.phases || []).forEach(function (x) { if (x.id === pid) ph = x; });
        return ph ? ph.name : '';
      }
      push('scope', lbl + ' — Phase', phName(a, p.phaseId), phName(b, it.phaseId));
      push('scope', lbl + ' — Workstream', p.workstream, it.workstream);
      push('scope', lbl + ' — Epic', p.epic, it.epic);
      push('scope', lbl + ' — Type', RM.typeOf(a, p, 'feature').label, RM.typeOf(b, it, 'feature').label);
      push('scope', lbl + ' — Size', p.size, it.size);
      push('scope', lbl + ' — ' + RM.riskColLabel(b), p.risk, it.risk);
      push('scope', lbl + ' — Priority', p.priority, it.priority);
      push('scope', lbl + ' — Role', p.teamType, it.teamType);
      push('scope', lbl + ' — Depends on', (p.deps || []).map(function (n) { return '#' + n; }).join(', '),
        (it.deps || []).map(function (n) { return '#' + n; }).join(', '));
      (b.meta.scopeCols || []).forEach(function (c) {
        push('scope', lbl + ' — ' + RM.scopeColLabel(c), txt(RM.scopeValue(p, c.key)), txt(RM.scopeValue(it, c.key)));
      });
      push('timeline', lbl + ' — Start', dDate(p.startDay), dDate(it.startDay));
      push('timeline', lbl + ' — Duration', dWeeks(p.durDays), dWeeks(it.durDays));
      push('timeline', lbl + ' — Risk buffer', dWeeks(p.riskDays || null), dWeeks(it.riskDays || null));
      push('timeline', lbl + ' — Deadline', p.deadline, it.deadline);
      push('timeline', lbl + ' — Milestone', p.milestone ? 'yes' : 'no', it.milestone ? 'yes' : 'no');
      push('timeline', lbl + ' — Locked', p.locked ? 'yes' : 'no', it.locked ? 'yes' : 'no');
      push('timeline', lbl + ' — Excluded from Auto', p.noAuto ? 'yes' : 'no', it.noAuto ? 'yes' : 'no');
      push('status', lbl + ' — Done', p.done ? 'done' : 'not done', it.done ? 'done' : 'not done');
      push('status', lbl + ' — Assignees', names(a, p.assignees), names(b, it.assignees));
      var sa = byId(p.stories), sb = byId(it.stories);
      (it.stories || []).forEach(function (st) {
        var sl = lbl + ' › #' + st.num + ' ' + shorten(st.title || '(story)', 24);
        var sp = sa[st.id];
        if (!sp) { push('scope', sl, '', 'Story added'); return; }
        push('scope', sl + ' — #', sp.num != null ? '#' + sp.num : '', '#' + st.num);
        push('scope', sl + ' — Title', sp.title, st.title);
        push('scope', sl + ' — Depends on', (sp.deps || []).map(function (n) { return '#' + n; }).join(', '),
          (st.deps || []).map(function (n) { return '#' + n; }).join(', '));
        push('scope', sl + ' — Description', txt(sp.description), txt(st.description));
        push('scope', sl + ' — Size', sp.size, st.size);
        push('scope', sl + ' — Type', RM.typeOf(a, sp, 'story').label, RM.typeOf(b, st, 'story').label);
        push('timeline', sl + ' — Start', dDate(sp.startDay), dDate(st.startDay));
        push('timeline', sl + ' — Duration', dWeeks(sp.durDays), dWeeks(st.durDays));
        push('timeline', sl + ' — Deadline', sp.deadline, st.deadline);
        push('timeline', sl + ' — Excluded from Auto', sp.noAuto ? 'yes' : 'no', st.noAuto ? 'yes' : 'no');
        push('status', sl + ' — Done', sp.done ? 'done' : 'not done', st.done ? 'done' : 'not done');
        push('status', sl + ' — Assignees', names(a, sp.assignees), names(b, st.assignees));
      });
      (p.stories || []).forEach(function (st) {
        if (!sb[st.id]) push('scope', lbl + ' › #' + st.num + ' ' + shorten(st.title || '(story)', 24), 'Story removed', '');
      });
    });
    (a.items || []).forEach(function (it) {
      if (!ib[it.id]) {
        tlPush(null, it);
        push('scope', '#' + it.num + ' ' + shorten(it.feature || '(untitled)', 28),
          it.milestone ? 'Milestone removed' : 'Feature removed', '');
      }
    });
    // ---- phases
    var pa = byId(a.phases), pb = byId(b.phases);
    (b.phases || []).forEach(function (p) {
      var prev = pa[p.id];
      var lbl = 'Phase ' + shorten(p.name || '(unnamed)', 28);
      if (!prev) { push('setup', lbl, '', 'Added'); return; }
      push('setup', lbl + ' — Name', prev.name, p.name);
      push('setup', lbl + ' — Description', txt(prev.description), txt(p.description));
      push('setup', lbl + ' — Backlog bucket', prev.bucket ? 'yes' : 'no', p.bucket ? 'yes' : 'no');
      push('timeline', lbl + ' — Pinned start', dDate(prev.startDay), dDate(p.startDay));
      push('timeline', lbl + ' — Pinned end', dDate(prev.endDay), dDate(p.endDay));
    });
    (a.phases || []).forEach(function (p) {
      if (!pb[p.id]) push('setup', 'Phase ' + shorten(p.name || '(unnamed)', 28), 'Removed', '');
    });
    // ---- team + costs (budgeting)
    var ta = byId(a.team), tb = byId(b.team);
    (b.team || []).forEach(function (m) {
      var prev = ta[m.id];
      var lbl = shorten(RM.memberLabel(m), 28);
      if (!prev) { push('budget', lbl, '', 'Person added'); return; }
      push('budget', lbl + ' — Name', prev.name, m.name);
      push('budget', lbl + ' — Role', prev.role, m.role);
      push('budget', lbl + ' — Rate card', prev.type, m.type);
      push('budget', lbl + ' — Workstreams', RM.memberWorkstreams(prev).join(', '), RM.memberWorkstreams(m).join(', '));
      push('budget', lbl + ' — Rate', prev.rate || '', m.rate || '');
      push('budget', lbl + ' — Cost', prev.cost || '', m.cost || '');
      push('budget', lbl + ' — Capacity', prev.capacity != null ? prev.capacity : 1, m.capacity != null ? m.capacity : 1);
      var wk = 0, wa = prev.weekHours || {}, wb2 = m.weekHours || {};
      Object.keys(wb2).forEach(function (k) { if (wa[k] !== wb2[k]) wk++; });
      Object.keys(wa).forEach(function (k) { if (!(k in wb2)) wk++; });
      if (wk) push('budget', lbl + ' — Week hours', '', wk + ' week(s) adjusted');
    });
    (a.team || []).forEach(function (m) {
      if (!tb[m.id]) push('budget', shorten(RM.memberLabel(m), 28), 'Person removed', '');
    });
    var ca = byId(a.costs), cb = byId(b.costs);
    (b.costs || []).forEach(function (c) {
      var prev = ca[c.id];
      var lbl = 'Cost ' + shorten(c.name || '(unnamed)', 26);
      if (!prev) { push('budget', lbl, '', 'Added'); return; }
      push('budget', lbl + ' — Name', prev.name, c.name);
      push('budget', lbl + ' — Amount', prev.amount, c.amount);
      push('budget', lbl + ' — Recurrence', prev.kind, c.kind);
      push('budget', lbl + ' — From', dDate(prev.startDay), dDate(c.startDay));
      push('budget', lbl + ' — Until', dDate(prev.endDay), dDate(c.endDay));
    });
    (a.costs || []).forEach(function (c) {
      if (!cb[c.id]) push('budget', 'Cost ' + shorten(c.name || '(unnamed)', 26), 'Removed', '');
    });
    // ---- setup: meta + rate card + workstream config, compared generically
    function genericDiff(cat, prefix, oa, ob) {
      oa = oa || {}; ob = ob || {};
      var keys = {};
      Object.keys(oa).concat(Object.keys(ob)).forEach(function (k) { keys[k] = 1; });
      Object.keys(keys).forEach(function (k) {
        var va = JSON.stringify(oa[k]), vb = JSON.stringify(ob[k]);
        if (va !== vb) push(cat, prefix + k, txt(va == null ? '' : va), txt(vb == null ? '' : vb));
      });
    }
    genericDiff('setup', 'Setup — ', a.meta, b.meta);
    genericDiff('budget', 'Rate card — ', (a.meta || {}).rateCard, (b.meta || {}).rateCard);
    genericDiff('setup', 'Workstream color — ', a.wsColors, b.wsColors);
    genericDiff('setup', 'Epic type — ', a.epicTypes, b.epicTypes);
    push('setup', 'Roles (rate card)', (a.teamTypes || []).join(', '), (b.teamTypes || []).join(', '));
    // options — renames of the active option and creates/renames/closes of
    // parked ones land in the audit trail (switching bypasses history)
    push('setup', 'Option name', a.optName || 'Default', b.optName || 'Default');
    function optNames(s) { return ((s && s.options) || []).map(function (o) { return o.name; }).join(', '); }
    push('setup', 'Other options', optNames(a), optNames(b));
    // rateCard sits inside meta too — drop the raw duplicate from the generic pass
    ops = ops.filter(function (op) { return op[1] !== 'Setup — rateCard'; });
    return { ops: ops, tl: tl };
  }
  // merge coalesced ops: same field keeps its FIRST old and LAST new value
  function mergeOps(base, add) {
    var out = base.slice();
    var at = {};
    out.forEach(function (op, i) { at[op[0] + '' + op[1]] = i; });
    add.forEach(function (op) {
      var k = op[0] + '' + op[1];
      if (at[k] != null) out[at[k]] = [op[0], op[1], out[at[k]][2], op[3]];
      else { at[k] = out.length; out.push(op); }
    });
    return out.filter(function (op) { return op[2] !== op[3]; });
  }
  // coalesced schedule moves keep each item's FIRST before and LAST after
  function mergeTl(base, add) {
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
  }
  var aiActor = false; // true while the AI assistant applies an edit (history attribution)
  function recordHistory(label, prevJson) {
    if (!label) return;
    var u = userName();
    if (aiActor) u = (u || 'Unknown user') + ' · AI';
    var h = state.history = Array.isArray(state.history) ? state.history : [];
    var ops = [], tl = [];
    if (prevJson) {
      try {
        var df = diffStates(JSON.parse(prevJson), state);
        ops = df.ops;
        tl = df.tl;
      } catch (e) { ops = []; tl = []; }
    }
    // the history array itself always differs between snapshots — never diff it
    ops = ops.filter(function (op) { return op[1].indexOf('history') === -1; });
    var last = h[h.length - 1];
    var now = Date.now();
    if (last && last.label === label && last.u === u && now - last.t < HISTORY_COALESCE_MS) {
      last.t = now;
      last.n = (last.n || 1) + 1;
      var merged = mergeOps(last.d || [], ops);
      last.x = (last.x || 0) + Math.max(0, merged.length - RM.HISTORY_OPS_MAX);
      last.d = merged.slice(0, RM.HISTORY_OPS_MAX);
      var mtl = mergeTl(last.tl || [], tl).slice(0, 60);
      if (mtl.length) last.tl = mtl; else delete last.tl;
      return;
    }
    var en = { t: now, u: u, label: label, n: 1 };
    if (ops.length > RM.HISTORY_OPS_MAX) en.x = ops.length - RM.HISTORY_OPS_MAX;
    en.d = ops.slice(0, RM.HISTORY_OPS_MAX);
    if (tl.length) en.tl = tl.slice(0, 60);
    h.push(en);
    if (h.length > RM.HISTORY_MAX) h.splice(0, h.length - RM.HISTORY_MAX);
  }
  // bumped whenever `state` changes (commit, replaceState, undo/redo, open):
  // the cheap key the Auto timeline dry runs are memoized on
  var stateRev = 0;
  // Auto timeline (one-shot, per phase): what a click would do right now.
  // Memoized per phase on the state revision + snap + today, so re-renders
  // never re-run the scheduler.
  var AUTO_TL_TIP = 'Auto timeline: move this phase\u2019s items to follow dependencies and capacity';
  var autoDry = { rev: -1, key: '', byPhase: {} };
  function autoPhaseDryRun(phaseId) {
    var key = snapFeat + '|' + snapStory + '|' + autoOrder + '|' + RM.todayDay(state.meta);
    if (autoDry.rev !== stateRev || autoDry.key !== key) autoDry = { rev: stateRev, key: key, byPhase: {} };
    if (!autoDry.byPhase[phaseId]) {
      // auto-order rides along inside the action (it re-sorts between passes),
      // so the dry run and the click agree
      var o = snapOpts();
      o.autoOrder = autoOrder;
      autoDry.byPhase[phaseId] = RM.autoPhase(state, phaseId, o);
    }
    return autoDry.byPhase[phaseId];
  }
  // { disabled, tip } for the band button, the band menu entry and the dialog
  function autoPhaseStatus(phaseId) {
    if (!state.meta.capacityEnabled) return { disabled: true, tip: 'Turn on scheduling in Setup \u2192 Scheduling' };
    var ph = state.phases.filter(function (p) { return p.id === phaseId; })[0];
    if (!ph || ph.bucket) return { disabled: true, tip: 'A backlog bucket is never auto-scheduled' };
    if (!autoPhaseDryRun(phaseId).changed) return { disabled: true, tip: 'Everything in this phase is already in place' };
    return { disabled: false, tip: AUTO_TL_TIP };
  }
  // row mark for a feature or story excluded from the Auto timeline — it
  // takes the lock icon's slot (Lock and Exclude never co-occur)
  function noAutoIcon() {
    return '<span class="r-lock r-noauto" title="Excluded from Auto timeline"><i data-lucide="zap-off"></i></span>';
  }
  // the Auto timeline action: lay the phase out (locked, excluded and done work stays put)
  // and, at the Stories level, size its features from their stories — one
  // commit, so one undo takes it all back
  function autoTimelinePhase(phaseId) {
    if (readOnly) { viewOnlyToast(); return 0; }
    var st = autoPhaseStatus(phaseId);
    if (!state.meta.capacityEnabled) { toast(st.tip, 'err'); return 0; }
    // the memoized dry run IS the result (the commit bumps the revision, so
    // the cache never outlives the state it was computed from)
    var r = autoPhaseDryRun(phaseId);
    if (!r.changed) { toast('Nothing to move'); return 0; }
    commit('auto timeline', function (s) {
      s.items = r.state.items;
      s.meta = r.state.meta;
    });
    var sz = r.sized.length;
    var szTxt = sz ? 'resized ' + sz + ' ' + (sz === 1 ? lvl('feature') : lvl('feature', true)).toLowerCase() : '';
    var stuck = (r.notes || []).filter(function (n) { return /never fits/.test(n); }).length;
    toast((r.moved
      ? 'Auto timeline moved ' + r.moved + ' item' + (r.moved === 1 ? '' : 's') + (sz ? ' and ' + szTxt : '')
      : sz ? 'Auto timeline ' + szTxt : 'Auto timeline rebuilt the ' + lvl('feature').toLowerCase() + ' bars') +
      (stuck ? ' \u00b7 ' + stuck + ' could not be placed' : ''));
    return r.changed;
  }
  // every real phase, one after another (the assistant's hook); returns the
  // number of changes
  function autoTimelineAll() {
    var n = 0;
    state.phases.filter(function (p) { return !p.bucket; }).forEach(function (p) {
      if (autoPhaseDryRun(p.id).changed) n += autoTimelinePhase(p.id);
    });
    return n;
  }
  function commit(label, mutate) {
    if (readOnly) { viewOnlyToast(); return; }
    // the new-project wizard edits a draft: no undo, history or saving
    if (wz) { if (mutate) mutate(state); wz.dirty = true; RM.applySizeRollup(state); renderWizard(); return; }
    var prev = JSON.stringify(state);
    undoStack.push(prev);
    if (undoStack.length > 120) undoStack.shift();
    redoStack.length = 0;
    if (mutate) mutate(state);
    recordHistory(label, prev);
    afterChange();
    maybeAskName();
  }
  function replaceState(label, next) {
    if (readOnly) { viewOnlyToast(); return; }
    if (wz) { state = next; wz.dirty = true; renderWizard(); return; }
    var prev = JSON.stringify(state);
    undoStack.push(prev);
    if (undoStack.length > 120) undoStack.shift();
    redoStack.length = 0;
    state = next;
    multiSel = null; // ids from the previous document mean nothing here
    recordHistory(label, prev);
    afterChange();
    maybeAskName();
  }
  // opening a file is not an edit: the document arrives as-is — no version-
  // history entry, nothing unsaved, and no undo back into the previous
  // document (that would autosave the old doc over the new file)
  function adoptState(next) {
    state = next;
    multiSel = null;
    undoStack.length = 0;
    redoStack.length = 0;
    docSaved = true;
    sessionEdited = false;
    stateRev += 1;
    // auto-order applies on open: the rows follow the timeline
    var prevJson = JSON.stringify(state);
    if (autoOrder) RM.sortItemsByStart(state);
    if (JSON.stringify(state) !== prevJson) {
      undoStack.length = 0; redoStack.length = 0;
      recordHistory('auto', prevJson);
      docSaved = false;
      sessionEdited = true;
      // the open left the document unsaved — never silently
      toast('Rows auto-ordered');
    }
    validation = RM.validate(state);
    saveLocal();
    render();
  }
  // the first change someone saves (autosave included) must carry an author:
  // ask for their name once, and stamp this session's anonymous entries
  function maybeAskName() {
    if (wz || userName() || namePromptShown) return;
    namePromptShown = true;
    openModal(
      '<div class="modal" style="width:420px">' +
      '<div class="m-head"><h2>Who’s editing?</h2></div>' +
      '<div class="m-body"><div class="m-sec"><label>Your name</label>' +
      '<input id="vhNameIn" style="width:100%" placeholder="e.g. Alex Rivera" autocomplete="name">' +
      '<div class="m-hint">Your changes just started saving. Every change is recorded in Version history with its author — set the name to record yours under. Stored on this machine only (Personal settings changes it later).</div></div></div>' +
      '<div class="m-foot"><button data-m="later">Not now</button><button id="vhNameSave" class="primary" disabled>Save name</button></div></div>',
      function (host) {
        var inp = $('#vhNameIn', host), save = $('#vhNameSave', host);
        if (inp) inp.focus();
        inp.addEventListener('input', function () { save.disabled = !inp.value.trim(); });
        function doSave() {
          var v = inp.value.trim();
          if (!v) return;
          setUserName(v);
          (state.history || []).forEach(function (en) {
            if (!en.u && en.t >= sessionStart) en.u = v;
          });
          closeModal();
          saveLocal();
          render();
        }
        save.onclick = doSave;
        inp.addEventListener('keydown', function (ev) { if (ev.key === 'Enter') doSave(); });
        $('[data-m=later]', host).onclick = closeModal;
      });
  }
  function undo() {
    if (wz || !undoStack.length) return;
    redoStack.push(JSON.stringify(state));
    state = JSON.parse(undoStack.pop());
    afterChange();
  }
  function redo() {
    if (wz || !redoStack.length) return;
    undoStack.push(JSON.stringify(state));
    state = JSON.parse(redoStack.pop());
    afterChange();
  }
  function afterChange() {
    docSaved = false;
    sessionEdited = true;
    stateRev += 1;
    RM.applySizeRollup(state); // the rollup scheme's feature sizes follow their stories
    validation = RM.validate(state);
    saveLocal();
    render();
    scheduleAutoSave();
  }

  var autoSaveTimer = null;
  function scheduleAutoSave() {
    if (wz) return; // never autosave while the wizard holds a draft
    // desktop only, and only once the doc lives in a real file
    if (!autoSave || !window.HeadwayDesktop || !HeadwayDesktop.currentPath()) return;
    clearTimeout(autoSaveTimer);
    autoSaveTimer = setTimeout(function () {
      if (!docSaved && !savingNow) doSave(false, true);
    }, 1500);
  }

  // ownOnly: judge the feature by its own fields, not its stories' titles
  function tagsHit(tags, q) {
    return (tags || []).some(function (t) { return String(t).toLowerCase().indexOf(q) !== -1; });
  }
  function matchesFilter(it, ownOnly) {
    if (!filterText) return true;
    var q = filterText.toLowerCase();
    // "#urgent" searches tags only — the rest of the fields stay out of it
    if (q.charAt(0) === '#') {
      var tq = q.slice(1);
      if (!tq) return true;
      return tagsHit(it.tags, tq) ||
        (!ownOnly && (it.stories || []).some(function (st) { return tagsHit(st.tags, tq); }));
    }
    return (it.feature || '').toLowerCase().indexOf(q) !== -1 ||
      (it.epic || '').toLowerCase().indexOf(q) !== -1 ||
      (it.workstream || '').toLowerCase().indexOf(q) !== -1 ||
      tagsHit(it.tags, q) ||
      state.meta.scopeCols.some(function (c) {
        return RM.htmlToText(RM.scopeValue(it, c.key)).toLowerCase().indexOf(q) !== -1;
      }) ||
      (!ownOnly && (it.stories || []).some(function (st) {
        return (st.title || '').toLowerCase().indexOf(q) !== -1 || tagsHit(st.tags, q);
      }));
  }

  // ------------------------------------------------------------ toasts, popover, modal
  // shadcn-style toasts: surface card + icon, bottom-center of the screen
  function toast(msg, kind) {
    var el = document.createElement('div');
    el.className = 'toast' + (kind === 'err' ? ' err' : '');
    el.innerHTML =
      '<i data-lucide="' + (kind === 'err' ? 'circle-alert' : 'circle-check') + '"></i>' +
      '<span class="toast-msg"></span>';
    el.querySelector('.toast-msg').textContent = msg;
    $('#toasts').appendChild(el);
    if (window.lucide) lucide.createIcons();
    setTimeout(function () { el.classList.add('gone'); }, 3400);
    setTimeout(function () { el.remove(); }, 3800);
  }

  // ---- tooltips: one styled component for EVERY hint in the app. Titles
  // are lifted to data-tip on first hover (so the native browser tooltip
  // never renders) and shown in a fixed, positioned element after a delay.
  var uiTip = null, uiTipTimer = null, uiTipFor = null;
  function hideUiTip() {
    clearTimeout(uiTipTimer);
    uiTipTimer = null;
    uiTipFor = null;
    if (uiTip) uiTip.hidden = true;
  }
  document.addEventListener('pointerover', function (e) {
    if (!e.target.closest) return;
    var t = e.target.closest('[title], [data-tip]');
    if (!t) { if (uiTipFor) hideUiTip(); return; }
    if (t.hasAttribute && t.hasAttribute('title')) {
      var tv = t.getAttribute('title');
      t.removeAttribute('title');
      if (tv) t.setAttribute('data-tip', tv);
    }
    var text = t.getAttribute && t.getAttribute('data-tip');
    if (!text || drag) { hideUiTip(); return; }
    if (uiTipFor === t) return;
    uiTipFor = t;
    clearTimeout(uiTipTimer);
    uiTipTimer = setTimeout(function () {
      if (uiTipFor !== t || !document.contains(t) || drag) return;
      if (!uiTip) {
        uiTip = document.createElement('div');
        uiTip.id = 'uiTip';
        document.body.appendChild(uiTip);
      }
      uiTip.textContent = t.getAttribute('data-tip');
      uiTip.hidden = false;
      uiTip.style.left = '0px';
      uiTip.style.top = '0px';
      var r = t.getBoundingClientRect();
      var tr = uiTip.getBoundingClientRect();
      var x = Math.min(Math.max(6, r.left + r.width / 2 - tr.width / 2), innerWidth - tr.width - 6);
      var y = r.top - tr.height - 7;
      if (y < 6) y = r.bottom + 7;
      uiTip.style.left = x + 'px';
      uiTip.style.top = y + 'px';
    }, 420);
  });
  document.addEventListener('pointerout', function (e) {
    if (!uiTipFor) return;
    if (e.target === uiTipFor || (uiTipFor.contains && uiTipFor.contains(e.target))) {
      if (!(e.relatedTarget && uiTipFor.contains && uiTipFor.contains(e.relatedTarget))) hideUiTip();
    }
  });
  document.addEventListener('pointerdown', hideUiTip, true);
  document.addEventListener('scroll', hideUiTip, true);
  window.addEventListener('blur', hideUiTip);

  // wire() callbacks attach delegated listeners to the host — swap in a clean
  // clone on every open so listeners never accumulate across opens
  function resetNode(el) {
    var clone = el.cloneNode(false);
    el.parentNode.replaceChild(clone, el);
    return clone;
  }

  var popEl = $('#popover');
  function openPopover(x, y, html, wire) {
    popEl = resetNode(popEl);
    popEl.innerHTML = html;
    popEl.hidden = false;
    var w = popEl.offsetWidth, h = popEl.offsetHeight;
    popEl.style.left = Math.min(x, window.innerWidth - w - 12) + 'px';
    popEl.style.top = Math.min(y, window.innerHeight - h - 12) + 'px';
    if (wire) wire(popEl);
  }
  function closePopover() { popEl.hidden = true; popEl.innerHTML = ''; }
  document.addEventListener('pointerdown', function (e) {
    // the calendar layer floats over popover forms — clicking it must not
    // tear the form underneath down
    if (!popEl.hidden && !popEl.contains(e.target) && !(e.target.closest && e.target.closest('#calPop'))) closePopover();
  }, true);

  var modalHost = $('#modalHost');
  function openModal(html, wire) {
    modalHost = resetNode(modalHost);
    modalHost.addEventListener('pointerdown', function (e) {
      if (e.target === modalHost) closeModal();
    });
    modalHost.innerHTML = html;
    modalHost.hidden = false;
    if (window.lucide) lucide.createIcons();
    if (wire) wire(modalHost);
    var first = $('.modal input, .modal select, .modal button', modalHost);
    if (first) first.focus();
  }
  function closeModal() { modalHost.hidden = true; modalHost.innerHTML = ''; }

  function confirmBox(title, body, okLabel, onOk, danger) {
    openModal(
      '<div class="modal" style="width:420px">' +
      '<div class="m-head"><h2>' + esc(title) + '</h2></div>' +
      '<div class="m-body"><div style="font-size:13px;color:var(--ink-2);line-height:1.5">' + body + '</div></div>' +
      '<div class="m-foot"><button data-m="cancel">Cancel</button>' +
      '<button data-m="ok" class="' + (danger ? 'danger' : 'primary') + '">' + esc(okLabel) + '</button></div></div>',
      function (host) {
        $('[data-m=cancel]', host).onclick = closeModal;
        $('[data-m=ok]', host).onclick = function () { closeModal(); onOk(); };
      });
  }

  // ------------------------------------------------------------ calendar
  // one popover calendar for every date the app edits. openCalendar anchors
  // to a control; readonly inputs with class .cal-in (ISO value) open it on
  // click/Enter and get their value set plus a bubbling change event.
  var CAL_MONTHS = ['January', 'February', 'March', 'April', 'May', 'June',
    'July', 'August', 'September', 'October', 'November', 'December'];
  var CAL_DOW = ['Su', 'Mo', 'Tu', 'We', 'Th', 'Fr', 'Sa'];
  // the calendar gets its own layer so it can float over popover forms
  // (cost dates, engagement windows) without tearing them down
  var calPop = document.createElement('div');
  calPop.id = 'calPop';
  calPop.hidden = true;
  document.body.appendChild(calPop);
  function closeCal() { calPop.hidden = true; calPop.innerHTML = ''; }
  document.addEventListener('pointerdown', function (e) {
    if (!calPop.hidden && !calPop.contains(e.target)) closeCal();
  }, true);
  function openCalendar(anchor, iso, onPick, opts) {
    opts = opts || {};
    var r = anchor.getBoundingClientRect();
    var sel = iso && /^\d{4}-\d{2}-\d{2}$/.test(String(iso)) ? String(iso).slice(0, 10) : '';
    var now = new Date();
    var todayIso = RM.fmtISO(new Date(Date.UTC(now.getFullYear(), now.getMonth(), now.getDate())));
    var base = RM.parseISO(sel || todayIso);
    var vy = base.getUTCFullYear(), vm = base.getUTCMonth();
    var ws = RM.weekStartOf(state.meta);
    var workSet = {};
    RM.workDaysOf(state.meta).forEach(function (d) { workSet[d] = true; });
    function calHtml() {
      var first = new Date(Date.UTC(vy, vm, 1));
      var back = ((first.getUTCDay() - ws) + 7) % 7;
      var cur = new Date(Date.UTC(vy, vm, 1 - back));
      var cells = '';
      for (var i = 0; i < 42; i++) {
        var ciso = RM.fmtISO(cur);
        cells += '<button class="cal-day' + (cur.getUTCMonth() === vm ? '' : ' out') +
          (ciso === todayIso ? ' today' : '') + (sel && ciso === sel ? ' sel' : '') +
          (workSet[cur.getUTCDay()] ? '' : ' off') +
          '" data-iso="' + ciso + '">' + cur.getUTCDate() + '</button>';
        cur.setUTCDate(cur.getUTCDate() + 1);
      }
      var dows = '';
      for (var d2 = 0; d2 < 7; d2++) dows += '<span class="cal-dow">' + CAL_DOW[(ws + d2) % 7] + '</span>';
      return '<div class="cal">' +
        '<div class="cal-head">' +
        '<button class="cal-nav" data-cal="prev" title="Previous month"><i data-lucide="chevron-left"></i></button>' +
        '<span class="cal-title">' + CAL_MONTHS[vm] + ' ' + vy + '</span>' +
        '<button class="cal-nav" data-cal="next" title="Next month"><i data-lucide="chevron-right"></i></button>' +
        '</div>' +
        '<div class="cal-grid">' + dows + cells + '</div>' +
        '<div class="cal-foot">' +
        '<button class="cal-lnk" data-cal="today">Today</button>' +
        (opts.allowClear ? '<button class="cal-lnk cal-clear" data-cal="clear">' + esc(opts.clearLabel || 'Clear') + '</button>' : '') +
        '</div></div>';
    }
    calPop = resetNode(calPop);
    calPop.innerHTML = calHtml();
    calPop.hidden = false;
    var cw = calPop.offsetWidth, ch = calPop.offsetHeight;
    calPop.style.left = Math.min(r.left, window.innerWidth - cw - 12) + 'px';
    calPop.style.top = Math.min(r.bottom + 4, window.innerHeight - ch - 12) + 'px';
    if (window.lucide) lucide.createIcons();
    calPop.addEventListener('click', function (ev) {
      var day = ev.target.closest('.cal-day');
      if (day) { closeCal(); onPick(day.dataset.iso); return; }
      var nav = ev.target.closest('[data-cal]');
      if (!nav) return;
      if (nav.dataset.cal === 'prev' || nav.dataset.cal === 'next') {
        vm += nav.dataset.cal === 'next' ? 1 : -1;
        if (vm < 0) { vm = 11; vy -= 1; }
        if (vm > 11) { vm = 0; vy += 1; }
        calPop.innerHTML = calHtml();
        if (window.lucide) lucide.createIcons();
        return;
      }
      if (nav.dataset.cal === 'today') { closeCal(); onPick(todayIso); }
      else if (nav.dataset.cal === 'clear') { closeCal(); onPick(''); }
    });
  }
  function calInputOpen(inp) {
    openCalendar(inp, inp.value, function (iso) {
      inp.value = iso;
      inp.dispatchEvent(new Event('change', { bubbles: true }));
    }, { allowClear: inp.dataset.calClear != null, clearLabel: inp.dataset.calClear || 'Clear' });
  }
  document.addEventListener('click', function (e) {
    var inp = e.target.closest && e.target.closest('input.cal-in');
    if (inp) calInputOpen(inp);
  });
  document.addEventListener('keydown', function (e) {
    if (e.key !== 'Enter' && e.key !== ' ' && e.key !== 'ArrowDown') return;
    var inp = e.target.closest && e.target.closest('input.cal-in');
    if (inp && document.activeElement === inp) { e.preventDefault(); calInputOpen(inp); }
  });

  // ------------------------------------------------------------ rich text
  // minimal DOM-based whitelist sanitizer for stored WYSIWYG HTML
  var WZ_TAGS = { B: 1, I: 1, U: 1, S: 1, EM: 1, STRONG: 1, UL: 1, OL: 1, LI: 1, P: 1, DIV: 1, BR: 1 };
  function sanitizeHtml(html) {
    if (!html) return '';
    var box = document.createElement('div');
    box.innerHTML = String(html);
    (function walk(node) {
      Array.prototype.slice.call(node.childNodes).forEach(function (ch) {
        if (ch.nodeType === 3) return;
        if (ch.nodeType !== 1 || ch.tagName === 'SCRIPT' || ch.tagName === 'STYLE') { node.removeChild(ch); return; }
        walk(ch);
        if (WZ_TAGS[ch.tagName]) {
          Array.prototype.slice.call(ch.attributes).forEach(function (a) { ch.removeAttribute(a.name); });
        } else {
          // unknown tags unwrap: keep their (already-walked) children
          while (ch.firstChild) node.insertBefore(ch.firstChild, ch);
          node.removeChild(ch);
        }
      });
    })(box);
    var out = box.innerHTML;
    return out === '<br>' ? '' : out;
  }

  // URLs in rich text become links at display time only: the stored value
  // stays plain (sanitizeHtml drops anchors on the way back in), so a link
  // is never more than the URL text the user typed
  var URL_RE = /https?:\/\/[^\s<>"']+/g;
  function linkifyHtml(html) {
    if (!html || html.indexOf('http') === -1) return html;
    var box = document.createElement('div');
    box.innerHTML = html;
    (function walk(node) {
      Array.prototype.slice.call(node.childNodes).forEach(function (ch) {
        if (ch.nodeType === 1) { if (ch.tagName !== 'A') walk(ch); return; }
        if (ch.nodeType !== 3 || ch.nodeValue.indexOf('http') === -1) return;
        var text = ch.nodeValue, frag = document.createDocumentFragment(), last = 0, m;
        URL_RE.lastIndex = 0;
        while ((m = URL_RE.exec(text))) {
          var u = m[0], tail = '';
          var t = u.match(/[.,;:!?)\]}'"]+$/);
          if (t) { tail = t[0]; u = u.slice(0, -tail.length); }
          if (!u) continue;
          frag.appendChild(document.createTextNode(text.slice(last, m.index)));
          var a = document.createElement('a');
          a.className = 'auto-link';
          a.setAttribute('href', u);
          a.setAttribute('title', (/Mac|iPhone|iPad/.test(navigator.platform || '') ? '\u2318' : 'Ctrl') + '-click to open');
          a.textContent = u;
          frag.appendChild(a);
          last = m.index + u.length;
        }
        if (!last) return;
        frag.appendChild(document.createTextNode(text.slice(last)));
        node.replaceChild(frag, ch);
      });
    })(box);
    return box.innerHTML;
  }
  function openExternal(url) {
    if (!/^https?:\/\//i.test(url)) return;
    if (window.HeadwayDesktop && window.HeadwayDesktop.openUrl) { window.HeadwayDesktop.openUrl(url); return; }
    window.open(url, '_blank', 'noopener');
  }
  // \u2318 / Ctrl-click follows an auto-link; a plain click only places the caret
  document.addEventListener('click', function (e) {
    var a = e.target.closest && e.target.closest('a.auto-link');
    if (!a) return;
    e.preventDefault();
    if (e.metaKey || e.ctrlKey) openExternal(a.getAttribute('href'));
  });

  // display a stored scope value as rich HTML; legacy plain-text values keep
  // their line breaks
  // displayHtml is what the editor round-trips to, so no-op guards compare
  // against it; richDisplay adds the display-only links on top.
  function displayHtml(v) {
    if (!v) return '';
    if (/<[a-z][\s\S]*>/i.test(v)) return sanitizeHtml(v);
    return esc(v).replace(/\n/g, '<br>');
  }
  function richDisplay(v) {
    return linkifyHtml(displayHtml(v));
  }

  // WYSIWYG editor block; the host wires the commit (blur / Save)
  function wysHtml(field, html, placeholder) {
    function btn(cmd, title, label) {
      return '<button type="button" tabindex="-1" data-wzc="' + cmd + '" title="' + title + '">' + label + '</button>';
    }
    return '<div class="wz">' +
      '<div class="wz-bar">' +
      btn('bold', 'Bold', '<b>B</b>') +
      btn('italic', 'Italic', '<i>I</i>') +
      btn('insertUnorderedList', 'Bullet list', '<i data-lucide="list"></i>') +
      btn('insertOrderedList', 'Numbered list', '<i data-lucide="list-ordered"></i>') +
      '</div>' +
      '<div class="wz-ed" contenteditable="true" data-f="' + field + '" data-ph="' + esc(placeholder || '') + '">' +
      linkifyHtml(sanitizeHtml(html)) + '</div></div>';
  }
  // toolbar buttons act on their editor without stealing its selection
  document.addEventListener('pointerdown', function (e) {
    var b = e.target.closest && e.target.closest('[data-wzc]');
    if (b) e.preventDefault();
  });
  document.addEventListener('click', function (e) {
    var b = e.target.closest && e.target.closest('[data-wzc]');
    if (!b) return;
    var ed = $('.wz-ed', b.closest('.wz'));
    if (ed) { ed.focus(); document.execCommand(b.dataset.wzc); }
  });

  // ⌘B / ⌘I (Ctrl on Windows/Linux) bold / italic inside any rich editor —
  // the desktop webview has no Format menu, so the shortcut has to be ours
  document.addEventListener('keydown', function (e) {
    if (!(e.metaKey || e.ctrlKey) || e.altKey || e.shiftKey) return;
    var k = e.key.toLowerCase();
    if (k !== 'b' && k !== 'i') return;
    var ed = e.target.closest && e.target.closest('[contenteditable="true"]');
    if (!ed || ed.classList.contains('sc-name') || ed.classList.contains('st-name')) return; // plain-text titles
    e.preventDefault();
    document.execCommand(k === 'b' ? 'bold' : 'italic');
  });

  // typing "- " or "1. " at the start of a line in any rich editor starts a
  // bullet / numbered list (the marker itself is swallowed)
  document.addEventListener('keydown', function (e) {
    if (e.key !== ' ') return;
    var ed = e.target.closest && e.target.closest('[contenteditable="true"]');
    if (!ed || ed.classList.contains('sc-name') || ed.classList.contains('st-name')) return;
    var sel = window.getSelection();
    if (!sel || !sel.rangeCount || !sel.isCollapsed) return;
    var r = sel.getRangeAt(0);
    var node = r.startContainer;
    if (node.nodeType !== 3) return;
    if (node.parentNode && node.parentNode.closest && node.parentNode.closest('li')) return;
    var before = node.textContent.slice(0, r.startOffset);
    var marker = before.replace(/\u00A0/g, ' ').trim();
    if (before.trim() !== before) return; // must be flush at the line start
    if (node.previousSibling) return;
    var isUl = marker === '-' || marker === '*';
    var isOl = /^\d+\.$/.test(marker);
    if (!isUl && !isOl) return;
    e.preventDefault();
    node.textContent = node.textContent.slice(r.startOffset);
    var nr = document.createRange();
    nr.setStart(node, 0);
    nr.collapse(true);
    sel.removeAllRanges();
    sel.addRange(nr);
    document.execCommand(isUl ? 'insertUnorderedList' : 'insertOrderedList');
  });

  // ------------------------------------------------------------ geometry
  function visibleSequence() {
    // ordered visible rows: bands and items (stories excluded)
    var seq = [];
    state.phases.forEach(function (p) {
      seq.push({ kind: 'band', phaseId: p.id });
      if (!p.collapsed) {
        RM.itemsInPhase(state, p.id).filter(matchesFilter).forEach(function (it) {
          seq.push({ kind: 'item', id: it.id, phaseId: p.id });
        });
      }
    });
    return seq;
  }

  // ---- detail level (Scoping + Planning): how deep the row grid goes.
  // feature: features, stories tucked away · story: every story row open
  function lvl(kind, plural) { return RM.levelLabel(state, kind, plural); }
  function dmModes() {
    return [
      ['feature', lvl('feature'), 'rows-3'],
      ['story', lvl('story'), 'list-tree']
    ];
  }
  function setDetailMode(mode) {
    detailMode = mode;
    // choosing the Feature level tucks every per-item story expansion away
    if (mode === 'feature') expanded = {};
    saveLocal();
    render();
  }
  // ONE level dropdown for every view: same markup (icon · label · caret),
  // same title, same place — the far left of the view's toolbar
  function levelBtnInner(modes, current) {
    var m = modes.filter(function (x) { return x[0] === current; })[0] || modes[0];
    return '<i data-lucide="' + m[2] + '"></i><span>' + esc(m[1]) + '</span><i data-lucide="chevron-down" class="dm-caret"></i>';
  }
  function levelBtnTitle(modes, current) {
    var m = modes.filter(function (x) { return x[0] === current; })[0] || modes[0];
    return 'Detail level: ' + m[1];
  }
  function levelBtnHtml(attr, modes, current) {
    return '<button class="dm-btn" ' + attr + ' title="' + esc(levelBtnTitle(modes, current)) + '">' +
      levelBtnInner(modes, current) + '</button>';
  }
  function openLevelMenu(anchor, modes, current, pick) {
    openDropdown(anchor, modes.map(function (m) {
      return { icon: m[2], label: esc(m[1]), checked: current === m[0], fn: function () { pick(m[0]); } };
    }));
  }
  function syncDetailBtn() {
    var b = $('#detailBtn');
    if (!b) return;
    b.innerHTML = levelBtnInner(dmModes(), detailMode);
    b.title = levelBtnTitle(dmModes(), detailMode);
    if (window.lucide) lucide.createIcons();
  }
  $('#detailBtn').addEventListener('click', function () {
    openLevelMenu($('#detailBtn'), dmModes(), detailMode, setDetailMode);
  });

  // ------------------------------------------------------------ options
  // alternate plan versions ("what if we cut phase 3?"). The active document
  // IS the current option; the rest are parked in state.options as full
  // snapshots. Switching swaps documents wholesale; comparing overlays the
  // other option's schedule as dashed ghost bars on the planning timeline.
  var compareOptId = null; // ghost-overlay option id; transient
  var cmpCache = null;     // per-render {id, name, doc, byId} for the overlay
  function cmpEntry() {
    if (!compareOptId) return null;
    return (state.options || []).filter(function (o) { return o.id === compareOptId; })[0] || null;
  }
  function getCmp() {
    var en = cmpEntry();
    if (!en) return null;
    if (!cmpCache || cmpCache.id !== en.id) {
      var doc = RM.normalizeState(en.doc);
      var byId = {};
      doc.items.forEach(function (x) { byId[x.id] = x; });
      cmpCache = { id: en.id, name: en.name, doc: doc, byId: byId };
    }
    return cmpCache;
  }
  function parkCurrent() {
    var doc = JSON.parse(JSON.stringify(state));
    delete doc.options;
    return { id: state.optId, name: state.optName, doc: doc };
  }
  function activateOption(id) {
    var entry = (state.options || []).filter(function (o) { return o.id === id; })[0];
    if (!entry) return;
    var prevId = state.optId;
    var next = RM.normalizeState(entry.doc);
    next.optId = entry.id;
    next.optName = entry.name;
    next.options = (state.options || []).map(function (o) { return o.id === id ? parkCurrent() : o; });
    // undo can't cross option documents — a revert here would fold one
    // option's edits into another option's parked snapshot
    undoStack.length = 0;
    redoStack.length = 0;
    selectedId = null;
    // comparing with the option being activated: keep the overlay meaningful
    // by flipping it to the one being parked
    if (compareOptId === id) compareOptId = prevId;
    state = next;
    stateRev += 1;
    docSaved = false;
    sessionEdited = true;
    validation = RM.validate(state);
    saveLocal();
    render();
    toast('Switched to “' + entry.name + '”');
  }
  function promptOptName(title, hint, initial, onOk) {
    openModal(
      '<div class="modal" style="width:420px"><div class="m-head"><h2>' + esc(title) + '</h2></div>' +
      '<div class="m-body"><div class="m-sec"><label>Option name</label>' +
      '<input id="optNameIn" style="width:100%" placeholder="e.g. Aggressive scope" value="' + esc(initial || '') + '">' +
      (hint ? '<div class="m-hint">' + hint + '</div>' : '') + '</div></div>' +
      '<div class="m-foot"><button data-m="cancel">Cancel</button><button id="optNameOk" class="primary">Save</button></div></div>',
      function (host) {
        var inp = $('#optNameIn', host);
        inp.focus();
        inp.select();
        function doOk() {
          var v = inp.value.trim().slice(0, 60);
          if (!v) return;
          closeModal();
          onOk(v);
        }
        $('#optNameOk', host).onclick = doOk;
        inp.addEventListener('keydown', function (ev) { if (ev.key === 'Enter') doOk(); });
        $('[data-m=cancel]', host).onclick = closeModal;
      });
  }
  function createOption() {
    if ((state.options || []).length >= RM.OPTIONS_MAX) {
      toast('Close an option first — up to ' + (RM.OPTIONS_MAX + 1) + ' open options', 'warn');
      return;
    }
    promptOptName('New option', 'Starts as a copy of “' + esc(state.optName) +
      '” — edits stay in the new option until you switch back.', '', function (nm) {
      commit('new option', function (s) {
        var doc = JSON.parse(JSON.stringify(s));
        delete doc.options;
        s.options = (s.options || []).concat([{ id: s.optId, name: s.optName, doc: doc }]);
        s.optId = RM.uid('opt');
        s.optName = nm;
      });
      toast('Now editing “' + nm + '”');
    });
  }
  function renameOption(id, curName) {
    promptOptName('Rename option', null, curName, function (nm) {
      commit('rename option', function (s) {
        if (id === s.optId) s.optName = nm;
        else (s.options || []).forEach(function (o) { if (o.id === id) o.name = nm; });
      });
    });
  }
  function closeOption(id, name) {
    confirmBox('Close option',
      'This removes “' + esc(name) + '” and its whole plan version. ' +
      (id === state.optId ? 'You will land on the next open option.' : 'The current option is kept.') +
      ' Undo can bring it back.',
      'Close option', function () {
        if (compareOptId === id) compareOptId = null;
        if (id === state.optId) {
          var first = (state.options || [])[0];
          if (!first) return;
          activateOption(first.id); // parks the closing option…
        }
        commit('close option', function (s) { // …then drops it
          s.options = (s.options || []).filter(function (o) { return o.id !== id; });
        });
        toast('Closed “' + name + '”');
      }, true);
  }
  function syncOptBtn() {
    var b = $('#optBtn');
    if (!b) return;
    var n = (state.options || []).length;
    b.innerHTML = '<i data-lucide="git-branch"></i><span>' + esc(state.optName) + '</span>' +
      (compareOptId ? '<span class="opt-cmp-dot" title="Comparing"></span>' : '') +
      '<i data-lucide="chevron-down" class="dm-caret"></i>';
    b.title = 'Option: ' + state.optName;
    if (window.lucide) lucide.createIcons();
  }
  function syncCmpPill() {
    if (compareOptId && !cmpEntry()) compareOptId = null; // option was closed
    var pill = $('#cmpPill');
    var en = cmpEntry();
    if (!en || view !== 'planning') {
      if (pill) pill.remove();
      return;
    }
    if (!pill) {
      pill = document.createElement('div');
      pill.id = 'cmpPill';
      $('#centerCol').appendChild(pill);
      pill.addEventListener('click', function (e) {
        if (e.target.closest('[data-cmpx]')) { compareOptId = null; render(); }
      });
    }
    pill.innerHTML = '<i data-lucide="eye"></i><span>Comparing with “' + esc(en.name) +
      '” — dashed bars</span><button data-cmpx title="Stop comparing"><i data-lucide="x"></i></button>';
    if (window.lucide) lucide.createIcons();
  }
  $('#optBtn').addEventListener('click', function () {
    var hasOthers = (state.options || []).length > 0;
    function toggleCompare(id) {
      compareOptId = compareOptId === id ? null : id;
      // comparing is a timeline overlay — jump there if elsewhere
      if (compareOptId && view !== 'planning') {
        view = 'planning';
        saveLocal();
      }
      render();
    }
    // one row per option: the active one leads with a check instead of the
    // branch icon; every row carries its own Compare / Rename / Delete
    function optionRow(id, name, isCur) {
      return {
        icon: isCur ? 'check' : 'git-branch', checked: isCur, label: esc(name),
        fn: isCur ? function () {} : function () { activateOption(id); },
        actions: [
          isCur ? null : { icon: 'eye', title: 'Compare', on: compareOptId === id,
            fn: function () { toggleCompare(id); } },
          { icon: 'pencil', title: 'Rename', fn: function () { renameOption(id, name); } },
          hasOthers ? { icon: 'trash-2', title: 'Delete',
            fn: function () { closeOption(id, name); } } : null
        ].filter(Boolean)
      };
    }
    var items = [optionRow(state.optId, state.optName, true)];
    (state.options || []).forEach(function (o) { items.push(optionRow(o.id, o.name, false)); });
    items.push({ sep: true });
    items.push({ icon: 'copy-plus', label: 'New option', fn: createOption });
    openDropdown($('#optBtn'), items, { minW: 300 });
  });

  // ------------------------------------------------------------ render
  // scoping view columns come from the document (meta.scopeCols — orderable,
  // removable, plus custom ones); widths are user-resizable and remembered in
  // the browser (UI_KEY)
  var SCOPE_DEFAULT_W = { description: 320, enables: 260, outOfScope: 260, extDeps: 240, notes: 300 };
  // fixed lead columns (field chips, not text): [key, label, width, min width]
  var SCOPE_FIXED = [
    ['assignees', 'Assignees', 100, 64],
    ['size', 'Size', 56, 44],
    ['risk', 'Risk', 56, 44],
    ['priority', 'Priority', 62, 44],
    ['duration', 'Duration', 74, 48],
    ['start', 'Start', 88, 64],
    ['deadline', 'Deadline', 88, 64],
    ['workstream', 'Workstream', 130, 80],
    ['epic', 'Epic', 150, 80]
  ];
  // fixed columns follow the project's feature switches; the assessment
  // column's label follows its scheme (Risk / Confidence / Priority)
  function scopeFixedCols() {
    return SCOPE_FIXED.filter(function (c) {
      if (c[0] === 'size') return RM.sizingEnabled(state);
      if (c[0] === 'risk') return RM.riskEnabled(state) || RM.riskEnabled(state, 'story');
      if (c[0] === 'priority') return RM.priorityEnabled(state);
      if (c[0] === 'workstream') return state.meta.workstreamsEnabled;
      return true;
    }).map(function (c) {
      if (c[0] !== 'risk') return c;
      return [c[0], RM.riskColLabel(state, RM.riskEnabled(state) ? 'feature' : 'story'), c[2], c[3]];
    });
  }
  // human labels for the assessment column's one-letter values, per scheme
  var RISK_VALUE_LABELS = {
    risk: { L: 'Low', M: 'Medium', H: 'High' },
    confidence: { H: 'High', M: 'Medium', L: 'Low' }
  };
  // ladder values draw as glyphs: v Low, = Medium, ^ High, ↑ Critical
  var LEVEL_GLYPHS = { L: 'chevron-down', M: 'equal', H: 'chevron-up', C: 'arrow-up' };
  function levelGlyph(v) {
    return LEVEL_GLYPHS[v] ? '<i data-lucide="' + LEVEL_GLYPHS[v] + '"></i>' : esc(v || '');
  }
  var PRIORITY_VALUE_LABELS = {
    moscow: { M: 'Must', S: 'Should', C: 'Could', W: 'Won’t' },
    levels: { C: 'Critical', H: 'High', M: 'Medium', L: 'Low' }
  };
  function priorityValueLabel(v, kind) {
    var mp = PRIORITY_VALUE_LABELS[RM.prioritySchemeOf(state, kind)] || {};
    return mp[v] || v || '';
  }
  // ' pt-crit' / ' pt-high' / ' pt-med' / ' pt-low' — the color tier class
  // every priority chip and picker button carries
  function priTierClass(v, kind) {
    var t = RM.priorityTier(state, v, kind);
    return t ? ' pt-' + t : '';
  }
  function priTierColor(v, kind) {
    var t = RM.PRIORITY_TIERS.indexOf(RM.priorityTier(state, v, kind));
    return t === -1 ? RM.PALETTE.neutral : RM.PRIORITY_RAMP[t];
  }

  function riskValueLabel(v, kind) {
    var mp = RISK_VALUE_LABELS[RM.riskSchemeOf(state, kind)] || {};
    return mp[v] || v || '';
  }
  // profile avatar: deterministic color + initials. Name, free-text role, and
  // rate-card assignment are all optional — fall back through them for display
  function mLabel(m) { return RM.memberLabel(m); }
  function mSub(m) { return (m && (m.role || m.type)) || ''; }
  function avatarHtml(m, extraCls) {
    var lbl = mLabel(m);
    return '<span class="avatar' + (extraCls ? ' ' + extraCls : '') + '" title="' + esc(lbl) +
      '" style="background:' + RM.avatarColor(lbl) + '">' + esc(RM.initialsOf(lbl)) + '</span>';
  }
  function memberById(id) {
    for (var i = 0; i < state.team.length; i++) if (state.team[i].id === id) return state.team[i];
    return null;
  }
  // small overlapping stack for rows/cards: at most `max` bubbles — when the
  // roster doesn't fit, the last bubble becomes a +N counter
  function avatarStack(ids, max) {
    var ms = (ids || []).map(memberById).filter(Boolean);
    if (!ms.length) return '';
    max = max || 3;
    var shown = ms.length > max ? max - 1 : ms.length;
    var out = ms.slice(0, shown).map(function (m) { return avatarHtml(m, 'sm'); }).join('');
    if (ms.length > shown) out += '<span class="avatar sm more">+' + (ms.length - shown) + '</span>';
    return '<span class="avstack">' + out + '</span>';
  }
  // all columns in the user's order (meta.scopeColOrder spans fixed + text)
  function allScopeCols() {
    var by = {};
    scopeFixedCols().forEach(function (c) { by[c[0]] = c; });
    scopeCols().forEach(function (c) { by[c[0]] = c; });
    var out = [];
    (state.meta.scopeColOrder || []).forEach(function (k) { if (by[k]) { out.push(by[k]); delete by[k]; } });
    Object.keys(by).forEach(function (k) { out.push(by[k]); });
    return out;
  }
  function isFixedColKey(k) { return RM.SCOPE_FIXED_KEYS.indexOf(k) !== -1; }
  function scopeCols() {
    return state.meta.scopeCols.map(function (c) {
      return [c.key, RM.scopeColLabel(c), SCOPE_DEFAULT_W[c.key] || 240];
    });
  }
  var scopeColW = {}; // field -> px override
  function scopeColDef(key) {
    return state.meta.scopeCols.filter(function (c) { return c.key === key; })[0] || null;
  }
  // does text column `key` show on rows of this kind ('feature' | 'story')?
  function colShowsOn(key, kind) {
    var c = scopeColDef(key);
    return !c || RM.scopeColShows(c, kind);
  }
  // the panel's Fields: text columns in the Scoping tab's order, kept to
  // the ones that show on this kind of row
  function panelScopeCols(kind) {
    return allScopeCols().map(function (c) { return scopeColDef(c[0]); })
      .filter(function (c) { return c && RM.scopeColShows(c, kind); });
  }
  function scopeScopeLabel(key) {
    return key === 'both' ? lvl('feature', true) + ' and ' + lvl('story', true).toLowerCase() :
      key === 'feature' ? lvl('feature', true) + ' only' : lvl('story', true) + ' only';
  }
  // milestone marker shapes: lucide icon + panel glyph per style
  var MS_STYLE_ICONS = { diamond: 'gem', star: 'star', circle: 'circle' };
  var MS_STYLE_GLYPHS = { diamond: '◆', star: '★', circle: '●' };
  function msStyleLabel(sname) { return sname.charAt(0).toUpperCase() + sname.slice(1); }
  function setMsStyle(itemId, sname) {
    commit('milestone style', function (s) {
      var t = RM.itemById(s, itemId);
      if (!t) return;
      if (sname === 'diamond') delete t.msStyle; else t.msStyle = sname;
    });
  }
  function msStyleItems(itemId, it) {
    return RM.MS_STYLES.map(function (sname) {
      return { icon: MS_STYLE_ICONS[sname], label: msStyleLabel(sname) + ' marker',
        checked: RM.msStyleOf(it) === sname, fn: function () { setMsStyle(itemId, sname); } };
    });
  }
  function setItemType(itemId, key) {
    commit('type', function (s) { var t = RM.itemById(s, itemId); if (t && RM.itemType(s, key)) t.type = key; });
  }
  function setStoryType(itemId, storyId, key) {
    commit('story type', function (s) {
      var t = RM.itemById(s, itemId); if (!t) return;
      t.stories.forEach(function (st) { if (st.id === storyId && RM.itemType(s, key)) st.type = key; });
    });
  }
  // delete-if-default / set-otherwise: shared by setEpicType and the epic
  // edit modal's save handler
  function applyEpicType(s, name, key) {
    if (!RM.itemType(s, key) || key === RM.defaultTypeFor(s, 'epic')) delete s.epicTypes[name];
    else s.epicTypes[name] = key;
  }
  function setEpicType(name, key) {
    commit('epic type', function (s) { applyEpicType(s, name, key); });
  }
  // dropdown / submenu entries for a level's types; the current type is
  // listed even when it is no longer allowed there
  function typeMenuItems(kind, currentKey, onPick) {
    var list = RM.typesFor(state, kind).slice();
    var cur = RM.itemType(state, currentKey);
    if (cur && !list.some(function (t) { return t.key === cur.key; })) list.push(cur);
    var allowed = RM.typesFor(state, kind).map(function (t) { return t.key; });
    return list.map(function (t) {
      var off = allowed.indexOf(t.key) === -1;
      return { icon: t.icon, label: esc(t.label) + (off ? ' (not allowed here)' : ''), checked: t.key === currentKey, fn: function () { onPick(t.key); } };
    });
  }
  function typeChipHtml(act, kind, obj) {
    var t = RM.typeOf(state, obj, kind);
    return '<button class="p-typechip" data-act="' + act + '" title="Type"><i data-lucide="' + esc(t.icon) + '"></i> ' + esc(t.label) + '</button>';
  }
  // Lucide icons drawn filled rather than stroked when used as a type glyph
  var TYPE_FILL_ICONS = { bookmark: 1 };
  // the glyph left of a title: the item's type icon in the active color-mode
  // color. The Feature default (`square`) is the filled rounded square the
  // rows always had; milestones keep their diamond/star/circle mask.
  // `owner` is the feature a story belongs to (stories borrow its color
  // except in the type color mode).
  function typeGlyphHtml(obj, kind, owner) {
    var it = owner || obj;
    var t = RM.typeOf(state, obj, kind);
    var color = '#' + (kind === 'story' && RM.colorMode() === 'type' ? RM.colorForType(state, t.key) : RM.colorForItem(state, it));
    var title = ' title="' + esc(t.label) + '"';
    if (kind !== 'story' && it.milestone) return '<span class="r-dot msdot ' + RM.msStyleOf(it) + '" style="background:' + color + '"' + title + '></span>';
    if (t.icon === 'square') return '<span class="r-dot" style="background:' + color + '"' + title + '></span>';
    return '<span class="r-type' + (TYPE_FILL_ICONS[t.icon] ? ' fill' : '') + '" style="color:' + color + '"' + title + '><i data-lucide="' + esc(t.icon) + '"></i></span>';
  }
  function typeIconMenu(anchor, key) {
    var t = RM.itemType(state, key);
    openDropdown(anchor, EPIC_ICONS.map(function (ic) {
      return { icon: ic, label: ic, checked: !!t && t.icon === ic, fn: function () {
        commit('type icon', function (s2) { RM.setItemTypeIcon(s2, key, ic); });
      } };
    }));
  }

  // panel sections: key -> collapsed override (persisted in UI_KEY);
  // unlisted keys fall back to the defaults below
  var panelSec = {};
  var PANEL_SEC_CLOSED = { stories: 1 };
  function secOpen(key) { return panelSec[key] != null ? !panelSec[key] : !PANEL_SEC_CLOSED[key]; }
  function scopeColWidth(c) { return Math.max(c[3] || 120, scopeColW[c[0]] || c[2]); }
  function scopeW() {
    return allScopeCols().reduce(function (a, c) { return a + scopeColWidth(c); }, 40);
  }

  // sticky band offsets need the real header and band heights (they vary by
  // view, capacity row, and band content); the CSS defaults are only fallbacks
  function syncHdrH() {
    var root = document.documentElement.style;
    root.setProperty('--hdr-h', $('#hdr').offsetHeight + 'px');
    var b = $('.row.band', rowsEl);
    if (b && b.offsetHeight) root.setProperty('--band-real-h', b.offsetHeight + 'px');
    var eb = $('.row.eband', rowsEl);
    if (eb && eb.offsetHeight) root.setProperty('--eband-real-h', eb.offsetHeight + 'px');
  }

  function render() {
    if (wz) { renderWizard(); return; } // the app behind the wizard waits
    var sx = board.scrollLeft, sy = board.scrollTop;
    if (!RM.appEnabled(state, view)) view = 'planning'; // an app switched off under us (or in a loaded file)
    $$('#viewTabs button[data-view]').forEach(function (b) { b.hidden = !RM.appEnabled(state, b.dataset.view); });
    critCache = RM.criticalPath(state);
    document.documentElement.style.setProperty('--week-px', weekPx + 'px');
    document.body.classList.toggle('no-cap', !showCap || !state.meta.capacityEnabled);
    document.body.classList.toggle('cap-off', !state.meta.capacityEnabled);
    document.body.classList.toggle('no-size', !RM.sizingEnabled(state));
    var boardView = view === 'planning' || view === 'scoping' || view === 'budget';
    document.body.classList.toggle('left-collapsed', leftCollapsed && boardView);
    // Planning's pane grows to fit its columns (see planLeftW)
    var planW = planLeftW();
    document.documentElement.style.setProperty('--left-w',
      (leftCollapsed && boardView ? 0
        : (view === 'scoping' ? leftWScope : view === 'budget' ? leftWBudget : planW)) + 'px');
    // folded: the pane is gone; only the floating reopen button remains
    var lp = $('#leftPeek');
    if (lp) lp.hidden = !(leftCollapsed && boardView);
    document.documentElement.style.setProperty('--panel-w', panelW + 'px');
    applyBuColWidths();
    applyPlColWidths();
    renderHlCols();
    syncDetailBtn();
    cmpCache = null; // parked docs may have changed (switch/rename/close)
    syncOptBtn();
    syncCmpPill();
    hidePlaceGhost();
    document.body.dataset.view = view;
    // a view switch or re-render can replace the rich cell the floating
    // B/I toolbar is anchored to — never leave the bar orphaned on screen
    if (scFmtTarget && ((view !== 'scoping' && view !== 'prio' && view !== 'sprints') || !scFmtTarget.isConnected)) hideScFmtBar();

    renderTopbar();
    syncZoomCtl();
    if (view === 'setup') {
      renderSetup();
      renderPanel();
      return;
    }
    if (view === 'reports') {
      renderReportsPage();
      renderPanel();
      return;
    }
    if (view === 'history') {
      renderHistoryPage();
      renderPanel();
      return;
    }
    if (view === 'prio') {
      renderPrioPage();
      renderPanel();
      return;
    }
    if (view === 'sprints') {
      renderSprintPage();
      renderPanel();
      return;
    }
    if (view === 'budget') {
      // budgeting lives on the same board as planning: shared header
      // (phase lane + dates/sprints), frozen left pane, one scroll surface
      var laneW2 = state.meta.numWeeks * weekPx;
      grid.style.width = 'calc(var(--left-w) + ' + laneW2 + 'px)';
      renderHeader(laneW2);
      renderBgCols(laneW2);
      renderBudgetRows();
      $('#arrowPaths').innerHTML = '';
      $('#arrows').setAttribute('width', 0);
      $('#arrows').setAttribute('height', 0);
      renderPanel();
      syncHdrH();
      board.scrollLeft = sx; board.scrollTop = sy;
      requestAnimationFrame(positionToday);
      return;
    }
    if (view === 'scoping') {
      grid.style.width = 'calc(var(--left-w) + ' + scopeW() + 'px)';
      renderScopeHeader();
      $('#hdrPhases').innerHTML = ''; // line stays visible for the row filter
      $('#bgcols').innerHTML = '';
      $('#hdrCapRows').innerHTML = '';
      $('#arrowPaths').innerHTML = '';
      // the svg keeps its planning-view width attribute otherwise, which
      // stretches the grid past the columns
      $('#arrows').setAttribute('width', 0);
      $('#arrows').setAttribute('height', 0);
      $('#todayLine').hidden = true;
      renderRows();
      autoGrowScope();
      renderPanel();
      syncHdrH();
      board.scrollLeft = sx; board.scrollTop = sy;
      return;
    }
    var laneW = state.meta.numWeeks * weekPx;
    grid.style.width = 'calc(var(--left-w) + ' + laneW + 'px)';
    renderHeader(laneW);
    renderBgCols(laneW);
    renderRows();
    renderResources();
    renderPanel();
    if (readOnly) lockDownRo();
    syncHdrH();
    board.scrollLeft = sx; board.scrollTop = sy;
    requestAnimationFrame(function () { renderArrows(); positionToday(); });
  }
  // view-only: every editor in the rendered surfaces turns inert (the
  // filters stay live — they only narrow the view)
  function lockDownRo() {
    $$('#rows [contenteditable="true"], #panel [contenteditable="true"], #prioView [contenteditable="true"], #sprintView [contenteditable="true"]')
      .forEach(function (el) { el.setAttribute('contenteditable', 'false'); });
    $$('#rows input, #rows textarea, #rows select, #panel input, #panel textarea, #panel select, #resGrid input, #prioView input, #sprintView input')
      .forEach(function (el) {
        if (el.type === 'search' || el.id === 'prFilter' || el.id === 'spFilter' || el.id === 'rowFilter') return;
        if (el.type === 'checkbox' || el.type === 'radio' || el.tagName === 'SELECT') el.disabled = true;
        else el.readOnly = true;
      });
  }

  // ------------------------------------------------------------ reports page
  // Dashboard-style project reporting (replaces the old drawer): KPI cards,
  // progress by phase, effort/cost by workstream, a cumulative planned-cost
  // curve, upcoming milestones, and open flags — the classic status areas
  // (schedule / scope / cost / risk) at a glance.
  function fmtPct(x) { return isFinite(x) ? Math.round(x * 100) + '%' : '—'; }
  function renderReportsPage() {
    var host = $('#reportsView');
    if (!host) return;
    var meta = state.meta;
    var items = state.items;

    // ---- effort-weighted completion (EVM-lite: done working days / total)
    var totalDays = 0, doneDays = 0, doneCount = 0, schedCount = 0, msTotal = 0, msDone = 0;
    items.forEach(function (it) {
      if (it.milestone) {
        msTotal += 1;
        if (it.done || (it.startDay != null && (function () {
          var today = RM.dateToDay(meta, new Date());
          return today != null && it.startDay < today && it.done;
        })())) msDone += it.done ? 1 : 0;
        return;
      }
      var info = RM.itemEffortInfo(state, it);
      totalDays += info.days;
      if (it.done) { doneDays += info.days; doneCount += 1; }
      if (it.startDay != null) schedCount += 1;
    });
    var workItems = items.filter(function (i) { return !i.milestone; });

    // ---- schedule window
    var stats = RM.scheduleStats(state);
    var today = RM.dateToDay(meta, new Date(Date.UTC(new Date().getFullYear(), new Date().getMonth(), new Date().getDate())));
    var lastDay = stats.lastDay != null ? stats.lastDay : RM.numDays(meta);
    var weeksLeft = today != null ? Math.max(0, Math.ceil((lastDay - today) / SPW())) : null;

    // ---- cost: items (estimate), roster (hours × cost), fixed/recurring
    var rep = RM.costReport(state, 'workstream');
    var rosterCost = 0, billing = 0;
    state.team.forEach(function (m) {
      var h = RM.roleTotalHours(state, m);
      rosterCost += h * RM.memberCost(state, m);
      billing += h * RM.memberRate(state, m);
    });
    var fixedCosts = RM.costsTotal(state);
    var planned = rosterCost + fixedCosts;
    var margin = billing > 0 ? (billing - planned) / billing : null;

    // ---- flags
    var flags = [];
    items.forEach(function (it) {
      (validation.byItem[it.id] || []).forEach(function (v) {
        if (v.level !== 'info') flags.push('#' + it.num + ' ' + shorten(it.feature || '(untitled)', 22) + ' — ' + v.msg);
      });
    });
    validation.global.forEach(function (v) { flags.push(v.msg); });

    function kpi(label, value, sub) {
      return '<div class="rp-kpi"><span class="rp-kpi-label">' + label + '</span>' +
        '<span class="rp-kpi-value">' + value + '</span>' +
        (sub ? '<span class="rp-kpi-sub">' + sub + '</span>' : '') + '</div>';
    }
    function barRow(label, frac, right, color) {
      var pct = Math.max(0, Math.min(1, frac));
      return '<div class="rp-bar-row"><span class="rp-bar-label" title="' + esc(label) + '">' + esc(label) + '</span>' +
        '<span class="rp-bar-track"><span class="rp-bar-fill" style="width:' + (pct * 100).toFixed(1) +
        '%;background:' + (color || 'var(--blue)') + '"></span></span>' +
        '<span class="rp-bar-val">' + right + '</span></div>';
    }

    // progress by phase (done effort / total effort)
    var phaseBars = state.phases.map(function (p) {
      var tot = 0, done = 0, n = 0;
      RM.itemsInPhase(state, p.id).forEach(function (it) {
        if (it.milestone) return;
        var d = RM.itemEffortInfo(state, it).days;
        tot += d; n += 1;
        if (it.done) done += d;
      });
      if (!n) return '';
      return barRow(p.name, tot ? done / tot : 0, fmtPct(tot ? done / tot : 0) + ' · ' + n + ' items');
    }).join('') || '<div class="p-none">No items yet.</div>';

    // effort & cost by workstream
    var maxCost = 1;
    rep.rows.forEach(function (r) { if (r.cost > maxCost) maxCost = r.cost; });
    var wsBars = rep.rows.map(function (r) {
      var col = '#' + RM.colorForWs(state, r.key === RM.defaultWsName(state) ? '' : r.key);
      return barRow(r.key, r.cost / maxCost,
        fmtMoney(r.cost) + ' · ' + (Math.round(r.days / SPW() * 10) / 10) + 'w', col);
    }).join('') || '<div class="p-none">Nothing scheduled or sized yet.</div>';

    // cumulative planned cost curve (roster + fixed/recurring), by week
    var weekly = new Array(meta.numWeeks);
    for (var w = 0; w < meta.numWeeks; w++) weekly[w] = 0;
    state.team.forEach(function (m) {
      var c = RM.memberCost(state, m);
      if (!c) return;
      for (var w2 = 0; w2 < meta.numWeeks; w2++) weekly[w2] += RM.roleWeekHours(state, m, w2).actual * c;
    });
    (state.costs || []).forEach(function (cst) {
      RM.costOccurrences(state, cst).forEach(function (o) {
        var wk = Math.min(meta.numWeeks - 1, Math.floor(o.day / SPW()));
        weekly[wk] += o.amount;
      });
    });
    var cum = [], run = 0;
    weekly.forEach(function (v) { run += v; cum.push(run); });
    var curve = '';
    if (run > 0) {
      var W = 560, H = 120;
      var pts = cum.map(function (v, i) {
        return (i / Math.max(1, cum.length - 1) * W).toFixed(1) + ',' + (H - (v / run) * (H - 6)).toFixed(1);
      });
      var todayX = today != null && today >= 0 && today <= RM.numDays(meta)
        ? (today / RM.numDays(meta)) * W : null;
      curve =
        '<svg class="rp-curve" viewBox="0 0 ' + W + ' ' + (H + 4) + '" preserveAspectRatio="none">' +
        '<polyline points="' + pts.join(' ') + '" fill="none" stroke="var(--blue)" stroke-width="2"/>' +
        '<polygon points="0,' + H + ' ' + pts.join(' ') + ' ' + W + ',' + H + '" fill="var(--blue)" opacity=".08"/>' +
        (todayX != null ? '<line x1="' + todayX.toFixed(1) + '" y1="0" x2="' + todayX.toFixed(1) + '" y2="' + H + '" stroke="var(--err)" stroke-dasharray="3 3" stroke-width="1"/>' : '') +
        '</svg>' +
        '<div class="rp-curve-cap"><span>' + esc(RM.fmtShortYear(RM.parseISO(meta.timelineStart))) + '</span>' +
        '<span>' + fmtMoney(run) + ' total</span>' +
        '<span>' + esc(meta.endDate ? RM.fmtShortYear(RM.parseISO(meta.endDate)) : '') + '</span></div>';
    } else {
      curve = '<div class="p-none">Add roster costs or fixed costs to see spend over time.</div>';
    }

    // upcoming milestones
    var msRows = items.filter(function (i) { return i.milestone && i.startDay != null; })
      .sort(function (a, b) { return a.startDay - b.startDay; })
      .slice(0, 6)
      .map(function (m2) {
        var past = today != null && m2.startDay < today;
        return '<div class="rp-ms' + (m2.done ? ' done' : past ? ' late' : '') + '">' +
          '<span class="rp-ms-d"></span>' +
          '<span class="rp-ms-name">' + esc(m2.feature || '(untitled)') + '</span>' +
          '<span class="rp-ms-date">' + RM.fmtShortYear(RM.dayToDate(meta, m2.startDay)) + '</span>' +
          '<span class="rp-ms-tag">' + (m2.done ? 'done' : past ? 'overdue' : 'upcoming') + '</span>' +
          '</div>';
      }).join('') || '<div class="p-none">No milestones on the timeline — convert an item via its context menu.</div>';

    // sprint-level delivery: what lands in each sprint (features + stories)
    // and how much of that scope is already done
    var curSp = currentSprintNum();
    var spBars = sprintNums().map(function (n) {
      var sIts = itemsOverlappingSprint(n);
      var sSts = storiesOverlappingSprint(n, sIts);
      var tot = sIts.length + sSts.length;
      if (!tot) return '';
      var done = sIts.filter(function (i) { return i.done; }).length +
        sSts.filter(function (p2) { return p2.st.done; }).length;
      var cur = n === curSp;
      return barRow(sprintLabel(n) + (cur ? ' · current' : ''), tot ? done / tot : 0,
        fmtPct(done / tot) + ' · ' + sIts.length + ' feature' + (sIts.length === 1 ? '' : 's') +
        (sSts.length ? ' · ' + sSts.length + (sSts.length === 1 ? ' story' : ' stories') : ''));
    }).join('') || '<div class="p-none">Nothing scheduled into sprints yet.</div>';

    var flagRows = flags.slice(0, 8).map(function (f) {
      return '<div class="rp-flag"><i data-lucide="triangle-alert"></i>' + esc(f) + '</div>';
    }).join('') || '<div class="rp-flag ok"><i data-lucide="circle-check"></i>No open warnings.</div>';

    host.innerHTML =
      '<div class="rp-page">' +
      '<h1 class="su-page">Reporting</h1>' +
      '<div class="rp-kpis">' +
      kpi('Complete', fmtPct(totalDays ? doneDays / totalDays : 0), doneCount + ' of ' + workItems.length + ' items done') +
      kpi('Scheduled', fmtPct(workItems.length ? schedCount / workItems.length : 0), schedCount + ' items on the timeline') +
      kpi('Time left', weeksLeft != null ? weeksLeft + 'w' : '—',
        stats.lastDay != null ? 'finishes ' + RM.fmtShortYear(RM.dayToDate(meta, Math.max(0, stats.lastDay - 1))) : 'nothing scheduled') +
      kpi('Effort', (Math.round(totalDays / SPW() * 10) / 10) + 'w', 'estimated person-weeks') +
      kpi('Planned cost', fmtMoney(planned), fmtMoney(rosterCost) + ' team + ' + fmtMoney(fixedCosts) + ' costs') +
      kpi('Billing', billing > 0 ? fmtMoney(billing) : '—',
        margin != null ? fmtPct(margin) + ' margin' : 'no bill rates set') +
      '</div>' +
      '<div class="rp-grid">' +
      '<section class="su-card rp-card"><h2>Progress by phase</h2>' + phaseBars + '</section>' +
      '<section class="su-card rp-card"><h2>Cost &amp; effort by workstream</h2>' + wsBars + '</section>' +
      '<section class="su-card rp-card rp-wide"><h2>Delivery by ' + (RM.sprintsEnabled(meta) ? 'sprint' : 'week') + '</h2>' + spBars + '</section>' +
      '<section class="su-card rp-card rp-wide"><h2>Planned spend over time</h2>' + curve + '</section>' +
      '<section class="su-card rp-card"><h2>Milestones</h2>' + msRows + '</section>' +
      '<section class="su-card rp-card"><h2>Flags</h2>' + flagRows + '</section>' +
      '</div></div>';
    if (window.lucide) lucide.createIcons();
  }

  // --------------------------------------------- sprint helpers (reports page)
  function currentSprintNum() {
    var meta = state.meta;
    var now = new Date();
    var d = RM.dateToDay(meta, new Date(Date.UTC(now.getFullYear(), now.getMonth(), now.getDate())));
    if (d == null) return null;   // today sits outside the timeline: no current sprint
    var wk = Math.max(0, Math.min(meta.numWeeks - 1, Math.floor(d / SPW())));
    return RM.sprintNumForWeek(meta, wk);
  }
  function sprintNums() {
    var meta = state.meta;
    var out = [], seen = {};
    for (var w = 0; w < meta.numWeeks; w++) {
      var n = RM.sprintNumForWeek(meta, w);
      if (!seen[n]) { seen[n] = true; out.push(n); }
    }
    return out;
  }
  function sprintLabel(n) {
    var meta = state.meta;
    var r = RM.sprintRange(meta, n);
    var d = RM.weekStartDate(meta, Math.max(0, r.w0));
    return (RM.sprintsEnabled(meta) ? 'Sprint ' + n + ' \u00B7 ' : 'Week of ') + RM.fmtShort(d);
  }
  function itemsInSprint(num) {
    return state.items.filter(function (it) {
      return it.startDay != null && sprFirstNum(it) === num && matchesFilter(it);
    });
  }
  // Reporting counts differently from Sprinting: a bar shows everything that
  // is *in flight* during the sprint, so a three-sprint feature lands in all
  // three bars. Sprinting lists a row once, in the sprint it starts in.
  function itemsOverlappingSprint(num) {
    var meta = state.meta;
    var r = RM.sprintRange(meta, num);
    return state.items.filter(function (it) {
      return RM.itemInWeeks(meta, it, r.w0, r.w1) && matchesFilter(it);
    });
  }
  function storiesOverlappingSprint(num, items) {
    var meta = state.meta;
    var r = RM.sprintRange(meta, num);
    var inIds = {};
    items.forEach(function (it) { inIds[it.id] = true; });
    var out = [];
    state.items.forEach(function (it) {
      if (!matchesFilter(it)) return;
      (it.stories || []).forEach(function (st) {
        var own = st.startDay != null && st.durDays != null;
        if (own ? RM.itemInWeeks(meta, st, r.w0, r.w1) : inIds[it.id]) out.push({ it: it, st: st });
      });
    });
    return out;
  }
  // ---------------------------------------------------------- prioritizing view
  // Kanban of the phases (the horizons). Cards carry the same size, priority,
  // epic and workstream chips as everywhere else, plus whichever scope
  // columns the Fields menu selects (Description by default). Columns
  // auto-sort by the active priority framework — the RICE scheme scores
  // Reach × Impact × Confidence ÷ Effort from dropdowns — and the Group
  // control lays the board out in workstream or epic swimlanes.
  var RICE_OPTIONS = {
    reach: { label: 'Reach / quarter', opts: [[1, '1'], [10, '10'], [100, '100'], [1000, '1,000'], [10000, '10,000']] },
    impact: { label: 'Impact', opts: [[0.25, '0.25 · Minimal'], [0.5, '0.5 · Low'], [1, '1 · Medium'], [2, '2 · High'], [3, '3 · Massive']] },
    confidence: { label: 'Confidence', opts: [[50, '50% · Low'], [80, '80% · Medium'], [100, '100% · High']] },
    effort: { label: 'Effort', opts: [[0.5, '½ week'], [1, '1 week'], [2, '2 weeks'], [4, '4 weeks'], [8, '8 weeks'], [12, '12 weeks']] }
  };
  function riceScoreLabel(it) {
    var sc = RM.riceScore(it);
    return sc == null ? '' : String(Math.round(sc * 10) / 10);
  }
  function riceSelectsHtml(it) {
    return Object.keys(RICE_OPTIONS).map(function (k) {
      var def = RICE_OPTIONS[k];
      var cur = (it.rice || {})[k];
      return '<label class="rice-lab">' + esc(def.label) +
        '<select data-rk="' + k + '"><option value="">—</option>' +
        def.opts.map(function (o) {
          return '<option value="' + o[0] + '"' + (cur === o[0] ? ' selected' : '') + '>' + esc(o[1]) + '</option>';
        }).join('') + '</select></label>';
    }).join('');
  }
  // the RICE editor every priority chip opens while the RICE scheme is active
  function riceMenu(anchor, itemId) {
    var it = RM.itemById(state, itemId);
    if (!it) return;
    var r = anchor.getBoundingClientRect();
    openPopover(r.left, r.bottom + 4, '<div class="rice-pop">' + riceSelectsHtml(it) + '</div>', function (host) {
      host.addEventListener('change', function (ev) {
        var k = ev.target.dataset.rk;
        if (!k) return;
        var v = ev.target.value === '' ? null : parseFloat(ev.target.value);
        commit('rice', function (s) {
          var x = RM.itemById(s, itemId);
          if (x) x.rice[k] = v;
        });
      });
    });
  }
  // ONE priority chip everywhere: the picked value for MoSCoW / levels, the
  // computed score under RICE
  function priChipContent(it) {
    if (RM.prioritySchemeOf(state) === 'rice') return esc(riceScoreLabel(it) || '·');
    if (!it.priority) return '';
    return RM.prioritySchemeOf(state) === 'levels' ? levelGlyph(it.priority) : esc(it.priority);
  }
  function priChipTitle(it) {
    if (RM.prioritySchemeOf(state) === 'rice') return 'RICE score';
    return 'Priority' + (it.priority ? '\nNow: ' + priorityValueLabel(it.priority) : '');
  }
  function priChipHasValue(it) {
    return !!(RM.prioritySchemeOf(state) === 'rice' ? riceScoreLabel(it) : it.priority);
  }
  function openPriorityEditor(anchor, itemId) {
    var it = RM.itemById(state, itemId);
    if (!it) return;
    if (RM.prioritySchemeOf(state) === 'rice') { riceMenu(anchor, itemId); return; }
    openDropdown(anchor, [{ label: '<i>None</i>', checked: !it.priority, fn: function () {
      commit('priority', function (s) { RM.itemById(s, itemId).priority = null; });
    } }].concat(RM.priorityOrderOf(state).map(function (pv) {
      return { icon: RM.prioritySchemeOf(state) === 'levels' ? LEVEL_GLYPHS[pv] : undefined,
        dot: '#' + priTierColor(pv),
        label: (RM.prioritySchemeOf(state) === 'levels' ? '' : pv + ' · ') + esc(priorityValueLabel(pv)),
        checked: it.priority === pv, fn: function () {
        commit('priority', function (s) { RM.itemById(s, itemId).priority = pv; });
      } };
    })));
  }
  // a story's priority picker — shared by the panel and the Prioritizing board
  function storyPriorityMenu(anchor, itemId, stId) {
    var stPri = (storyById(RM.itemById(state, itemId) || {}, stId) || {}).priority;
    function setPri(pv) {
      commit('story priority', function (s) {
        var st2 = storyById(RM.itemById(s, itemId) || {}, stId);
        if (st2) st2.priority = pv;
      });
    }
    var stLevels = RM.prioritySchemeOf(state, 'story') === 'levels';
    openDropdown(anchor, [{ label: '<i>None</i>', checked: !stPri, fn: function () { setPri(null); } }]
      .concat(RM.priorityOrderOf(state, 'story').map(function (pv) {
        return { icon: stLevels ? LEVEL_GLYPHS[pv] : undefined,
          dot: '#' + priTierColor(pv, 'story'),
          label: (stLevels ? '' : pv + ' · ') + esc(priorityValueLabel(pv, 'story')),
          checked: stPri === pv, fn: function () { setPri(pv); } };
      })));
  }
  // ---- story estimate fields: setters, pickers and ONE chip set shared by
  // Planning, Scoping, Sprinting and Prioritizing (size · priority · risk ·
  // duration, each only when its scheme is on)
  function withStory(label, itemId, stId, fn) {
    commit(label, function (s) {
      var st2 = storyById(RM.itemById(s, itemId) || {}, stId);
      if (st2) fn(st2, s);
    });
  }
  function setStorySize(itemId, stId, sz) {
    withStory('story size', itemId, stId, function (st2, s) {
      st2.size = sz;
      // a 0-point size is zero effort but still one day on the grid
      if (sz && st2.startDay != null) st2.durDays = RM.stretchSpan(s.meta, st2.startDay, snapUpDays(Math.max(1, RM.sizeDays(s, sz, 'story')), 'story'));
    });
  }
  function setStoryRisk(itemId, stId, rv) {
    withStory('story risk', itemId, stId, function (st2) { st2.risk = rv; });
  }
  // days = null clears the duration (and takes the story off the timeline)
  function setStoryDur(itemId, stId, days) {
    withStory('story duration', itemId, stId, function (st2, s) {
      if (days == null) { st2.durDays = null; st2.startDay = null; }
      else {
        var want = snapUpDays(days, 'story'); // a story buys whole snap units
        st2.durDays = st2.startDay != null ? RM.stretchSpan(s.meta, st2.startDay, want) : want;
      }
    });
  }
  function storySizeMenu(anchor, itemId, stId) {
    var cur = (storyById(RM.itemById(state, itemId) || {}, stId) || {}).size;
    openDropdown(anchor, [{ label: '<i>no size</i>', checked: !cur, fn: function () { setStorySize(itemId, stId, null); } }]
      .concat(RM.sizeOrderOf(state, 'story').map(function (sz) {
        return { label: esc(sz) + ' <small>' + sizeHuman(sz, 'story') + '</small>', checked: cur === sz,
          fn: function () { setStorySize(itemId, stId, sz); } };
      })));
  }
  function storyRiskMenu(anchor, itemId, stId) {
    if (!RM.riskEnabled(state, 'story')) return;
    var cur = (storyById(RM.itemById(state, itemId) || {}, stId) || {}).risk;
    openDropdown(anchor, [{ label: '<i>None</i>', checked: !cur, fn: function () { setStoryRisk(itemId, stId, null); } }]
      .concat(RM.riskOrderOf(state, 'story').map(function (rv) {
        return { icon: LEVEL_GLYPHS[rv], label: esc(riskValueLabel(rv, 'story')), checked: cur === rv,
          fn: function () { setStoryRisk(itemId, stId, rv); } };
      })));
  }
  // inline weeks editor for a duration chip: empty clears, a number sets
  // working days; the caller commits through onSave(days | null)
  function inlineWeeksEditor(chip, curDays, onSave) {
    if (chip.querySelector('input')) return;
    var inp = document.createElement('input');
    inp.type = 'number'; inp.min = '0.2'; inp.step = '0.2';
    inp.value = curDays != null ? Math.round(curDays / SPW() * 100) / 100 : '';
    inp.className = 'hc-edit'; inp.style.width = '30px';
    chip.textContent = '';
    chip.appendChild(inp);
    inp.focus(); inp.select();
    var done = false;
    var fin = function (saveIt) {
      if (done) return; done = true;
      if (!saveIt) { render(); return; }
      var v = String(inp.value).trim();
      onSave(v === '' ? null : Math.max(1, Math.round(Math.max(0.2, parseFloat(v) || 1) * SPW())));
    };
    inp.addEventListener('blur', function () { fin(true); });
    inp.addEventListener('keydown', function (ev) {
      ev.stopPropagation();
      if (ev.key === 'Enter') fin(true);
      if (ev.key === 'Escape') fin(false);
    });
  }
  // a story's duration readout: its own span, else its size estimate
  function storyWeeks(st) {
    var days = st.startDay != null && st.durDays != null ? RM.workInSpan(state.meta, st.startDay, st.durDays)
      : st.durDays != null ? st.durDays : RM.sizeDays(state, st.size, 'story');
    return days ? fmtDays(days) : '';
  }
  // story chip HTML; attr names the click hook ('data-act' on rows,
  // 'data-spact' on Sprinting, 'data-prstact' on Prioritizing cards). Values
  // are st-size / st-pri / st-risk / st-wk everywhere.
  function storyChipHtml(key, st, attr, blank) {
    blank = blank == null ? '·' : blank;
    if (key === 'size') {
      if (!RM.sizingEnabled(state, 'story')) return '';
      return '<span class="r-size" tabindex="0" role="button" ' + attr + '="st-size" title="Story size">' +
        (st.size ? esc(st.size) : blank) + '</span>';
    }
    if (key === 'pri') {
      if (!RM.priorityEnabled(state, 'story')) return '';
      return '<span class="r-risk pri' + (st.priority ? ' has-risk' : '') + priTierClass(st.priority, 'story') + '" tabindex="0" role="button" ' + attr + '="st-pri" title="' +
        esc(lvl('story') + ' priority' + (st.priority ? '\nNow: ' + priorityValueLabel(st.priority, 'story') : '')) + '">' +
        (st.priority ? (RM.prioritySchemeOf(state, 'story') === 'levels' ? levelGlyph(st.priority) : esc(st.priority)) : blank) + '</span>';
    }
    if (key === 'risk') {
      if (!RM.riskEnabled(state, 'story')) return '';
      return '<span class="r-risk rk-none' + (st.risk ? ' has-risk' : '') + '" tabindex="0" role="button" ' + attr + '="st-risk" title="' +
        esc(RM.riskColLabel(state, 'story') + (st.risk ? '\nNow: ' + riskValueLabel(st.risk, 'story') : '')) + '">' +
        (st.risk ? levelGlyph(st.risk) : blank) + '</span>';
    }
    if (key === 'dur') {
      return '<span class="r-wk editable" tabindex="0" role="button" ' + attr + '="st-wk" title="Story duration">' +
        (storyWeeks(st) || blank) + '</span>';
    }
    if (key === 'cap') {
      if (!state.meta.capacityEnabled) return '';
      return '<span class="r-cap' + (st.capType ? '' : ' empty') + '" tabindex="0" role="button" ' + attr + '="st-cap" title="' +
        esc('Capacity type' + (st.capType ? '\nNow: ' + st.capType : '')) + '">' + (st.capType ? esc(st.capType) : blank) + '</span>';
    }
    if (key === 'mult') {
      if (!state.meta.capacityEnabled || state.meta.capMode === 'points') return '';
      var mv = st.capMult || 1;
      return '<span class="r-mult editable' + (mv === 1 ? ' one' : '') + '" tabindex="0" role="button" ' + attr + '="st-mult" title="Capacity multiplier — people this story needs at once">×' + fmtPe(mv) + '</span>';
    }
    return '';
  }
  // routes a story chip click; returns true when it handled one
  function storyChipAction(act, anchor, itemId, stId) {
    if (act === 'st-size') storySizeMenu(anchor, itemId, stId);
    else if (act === 'st-pri') storyPriorityMenu(anchor, itemId, stId);
    else if (act === 'st-risk') storyRiskMenu(anchor, itemId, stId);
    else if (act === 'st-wk') {
      var stW = storyById(RM.itemById(state, itemId) || {}, stId);
      inlineWeeksEditor(anchor, stW ? stW.durDays : null, function (days) { setStoryDur(itemId, stId, days); });
    } else if (act === 'st-asg') {
      if (!state.team.length) { toast('Add people in the Resources panel first'); return true; }
      openDropdown(anchor, storyAssignMenuItems(itemId, stId), ASSIGN_DD);
    } else if (act === 'st-cap') openDropdown(anchor, storyCapMenuItems(itemId, stId));
    else if (act === 'st-mult') {
      var stM = storyById(RM.itemById(state, itemId) || {}, stId);
      inlineMultEditor(anchor, stM ? stM.capMult : 1, function (v) {
        commit('story multiplier', function (s) {
          var st2 = storyById(RM.itemById(s, itemId) || {}, stId);
          if (st2) st2.capMult = v;
        });
      });
    } else return false;
    return true;
  }
  // the capacity type a story / feature drains: the document's types, or general
  function storyCapMenuItems(itemId, stId) {
    var cur = (storyById(RM.itemById(state, itemId) || {}, stId) || {}).capType || '';
    return [{ label: '<i>— general —</i>', checked: !cur, fn: function () {
      commit('story capacity type', function (s) { var st2 = storyById(RM.itemById(s, itemId) || {}, stId); if (st2) st2.capType = ''; });
    } }].concat(RM.capTypesOf(state).map(function (t) {
      return { label: esc(t), checked: cur === t, fn: function () {
        commit('story capacity type', function (s) { var st2 = storyById(RM.itemById(s, itemId) || {}, stId); if (st2) st2.capType = t; });
      } };
    }));
  }
  // a feature's planning type comes from its stories while they all agree; then
  // its own capType is only what it falls back to, so the chip reads inherited
  function itemCapInherited(it) {
    return !!it && !it.milestone && RM.itemCapType(state, it) !== (it.capType || '');
  }
  var CAP_INHERIT_TITLE = 'Inherited from its stories — set the stories’ types to change it';
  function itemCapMenuItems(itemId) {
    var itC = RM.itemById(state, itemId);
    var cur = (itC || {}).capType || '';
    var head = itemCapInherited(itC)
      ? [{ label: '<i>Type follows the stories while they all agree</i>', disabled: true, fn: function () {} }, { sep: true }]
      : [];
    // a feature always plans as one of the types — no "general" entry here
    // (stories keep theirs)
    return head.concat(RM.capTypesOf(state).map(function (t) {
      return { label: esc(t), checked: cur === t, fn: function () {
        commit('capacity type', function (s) { var t2 = RM.itemById(s, itemId); if (t2) t2.capType = t; });
      } };
    }));
  }
  // × multiplier chip edits in place like the duration chip
  function inlineMultEditor(chip, cur, onSave) {
    if (chip.querySelector('input')) return;
    var inp = document.createElement('input');
    inp.type = 'number'; inp.min = '0.1'; inp.step = '0.5';
    inp.value = cur != null ? cur : 1;
    inp.className = 'hc-edit'; inp.style.width = '30px';
    chip.textContent = '';
    chip.appendChild(inp);
    inp.focus(); inp.select();
    var done = false;
    var fin = function (saveIt) {
      if (done) return; done = true;
      if (!saveIt) { render(); return; }
      var v = parseFloat(inp.value);
      onSave(isFinite(v) && v > 0 ? v : 1);
    };
    inp.addEventListener('blur', function () { fin(true); });
    inp.addEventListener('keydown', function (ev) {
      ev.stopPropagation();
      if (ev.key === 'Enter') fin(true);
      if (ev.key === 'Escape') fin(false);
    });
  }
  // feature twins for the boards that lack the row chips (Sprinting, Prioritizing)
  function itemRiskMenu(anchor, itemId) {
    var it = RM.itemById(state, itemId);
    if (!it || RM.riskSchemeOf(state) === 'auto') return; // computed, not set
    openDropdown(anchor, [{ label: '<i>None</i>', checked: !it.risk, fn: function () { setItemRisk(itemId, null); } }]
      .concat(RM.riskOrderOf(state).map(function (rv) {
        return { icon: LEVEL_GLYPHS[rv], label: esc(riskValueLabel(rv)), checked: it.risk === rv,
          fn: function () { setItemRisk(itemId, rv); } };
      })));
  }
  function setItemDur(itemId, days) {
    commit('duration', function (s) {
      var t = RM.itemById(s, itemId);
      if (!t || t.milestone) return;
      if (days == null) { t.startDay = null; t.durDays = null; }
      else t.durDays = t.startDay != null ? RM.stretchSpan(s.meta, t.startDay, days) : days;
    });
  }
  function itemChipHtml(key, it, attr) {
    if (it.milestone) return '';
    if (key === 'size') {
      if (!RM.sizingEnabled(state)) return '';
      return '<span class="r-size' + sizeRoCls(it) + '" tabindex="0" role="button" ' + attr + '="size" title="' + sizeChipTitle(it) + '">' + (it.size ? esc(it.size) : '·') + '</span>';
    }
    if (key === 'pri') {
      if (!RM.priorityEnabled(state)) return '';
      return '<span class="r-risk pri' + (priChipHasValue(it) ? ' has-risk' : '') + priTierClass(it.priority) + '" tabindex="0" role="button" ' + attr + '="priority" title="' +
        esc(priChipTitle(it)) + '">' + (priChipContent(it) || '·') + '</span>';
    }
    if (key === 'risk') {
      if (!RM.riskEnabled(state)) return '';
      if (RM.riskSchemeOf(state) === 'auto') {
        var rkA = RM.depRisk(state, it);
        return '<span class="r-risk rk-' + rkA.level + '" title="' + esc(rkA.level === 'none' ? 'No dependency risk detected' : 'Dependency risk ' + rkA.level) + '">' +
          (rkA.level === 'none' ? '·' : levelGlyph(rkA.level.charAt(0).toUpperCase())) + '</span>';
      }
      return '<span class="r-risk rk-none' + (it.risk ? ' has-risk' : '') + '" tabindex="0" role="button" ' + attr + '="risk" title="' +
        esc(RM.riskColLabel(state) + (it.risk ? '\nNow: ' + riskValueLabel(it.risk) : '')) + '">' +
        (it.risk ? levelGlyph(it.risk) : '·') + '</span>';
    }
    if (key === 'dur') {
      return '<span class="r-wk editable" tabindex="0" role="button" ' + attr + '="dur" title="Duration">' + totalWeeks(it) + '</span>';
    }
    return '';
  }
  function itemChipAction(act, anchor, itemId) {
    var itA = RM.itemById(state, itemId);
    if (!itA) return false;
    if (act === 'size') {
      if (sizeLocked(itA)) return true;
      openDropdown(anchor, [{ label: '<i>no size</i>', checked: !itA.size, fn: function () { setItemSize(itemId, null); } }]
        .concat(RM.sizeOrderOf(state).map(function (sz) {
          return { label: esc(sz) + ' <small>' + sizeHuman(sz) + '</small>', checked: itA.size === sz, fn: function () { setItemSize(itemId, sz); } };
        })));
    } else if (act === 'priority') openPriorityEditor(anchor, itemId);
    else if (act === 'risk') itemRiskMenu(anchor, itemId);
    else if (act === 'dur') inlineWeeksEditor(anchor, itA.durDays, function (days) { setItemDur(itemId, days); });
    else if (act === 'cap') openDropdown(anchor, itemCapMenuItems(itemId));
    else if (act === 'mult') inlineMultEditor(anchor, itA.capMult, function (v) {
      commit('multiplier', function (s) { RM.itemById(s, itemId).capMult = v; });
    });
    else if (act === 'asg') {
      if (!state.team.length) { toast('Add people in the Resources panel first'); return true; }
      openDropdown(anchor, assignMenuItems(itemId), ASSIGN_DD);
    } else return false;
    return true;
  }
  // sort under the active framework: RICE score descending, else the
  // scheme's ladder; unset last, ties keep document order
  function prioRank(it) {
    if (RM.prioritySchemeOf(state) === 'rice') {
      var sc = RM.riceScore(it);
      return sc == null ? Infinity : -sc;
    }
    var order = RM.priorityOrderOf(state);
    var i = it.priority ? order.indexOf(it.priority) : -1;
    return i === -1 ? Infinity : i;
  }
  function prSortItems(list) {
    var rank;
    if (prioSort === 'title') {
      rank = null; // compare directly
    } else if (prioSort === 'size') {
      var so = RM.sizeOrderOf(state);
      rank = function (it) { var i2 = it.size ? so.indexOf(it.size) : -1; return i2 === -1 ? Infinity : -i2; }; // big first, unset last
    } else if (prioSort === 'priority' && RM.priorityEnabled(state)) {
      rank = prioRank;
    } else {
      return list.slice(); // document order
    }
    return list.map(function (it, i) { return { it: it, i: i }; })
      .sort(function (a, b) {
        if (!rank) return (a.it.feature || '').localeCompare(b.it.feature || '') || (a.i - b.i);
        return (rank(a.it) - rank(b.it)) || (a.i - b.i);
      })
      .map(function (x) { return x.it; });
  }
  // board-level match: the text filter plus the epic / workstream dropdowns
  function prMatches(it) {
    if (it.milestone) return false; // milestones are dates, not work to rank
    if (!matchesFilter(it)) return false;
    if (prioFPhase != null && it.phaseId !== prioFPhase) return false;
    if (prioFEpic != null && (it.epic || '') !== prioFEpic) return false;
    if (prioFWs != null && (it.workstream || '') !== prioFWs) return false;
    return true;
  }
  // chips a feature card can show; the Fields menu hides any of them.
  // Only chips the document can draw are offered (see prChipAvail).
  var PR_CHIPS = { size: 'Size', priority: 'Priority', risk: 'Risk', dur: 'Duration', epic: 'Epic', ws: 'Workstream' };
  function prChipAvail(k) {
    if (k === 'size') return RM.sizingEnabled(state) && prioFeatCol !== 'size';
    if (k === 'priority') return RM.priorityEnabled(state);
    if (k === 'risk') return RM.riskEnabled(state);
    if (k === 'epic') return prioGroup !== 'epic'; // the swimlane already names the epic
    if (k === 'ws') return !!state.meta.workstreamsEnabled && prioGroup !== 'ws';
    return true; // dur
  }
  function prChipOn(k) { return prChipAvail(k) && prioChipHide.indexOf(k) === -1; }
  function prCardFields() {
    var by = {};
    scopeCols().forEach(function (c) { if (colShowsOn(c[0], 'feature')) by[c[0]] = c; });
    return prioFields.map(function (k) { return by[k]; }).filter(Boolean);
  }
  // a card title is plain text (double-click or Rename… edits it in place);
  // an empty one shows a quiet placeholder so there is still something to hit
  function prTitleHtml(txt, f, cls) {
    var s = txt || '';
    return '<span class="' + cls + (s ? '' : ' pr-ph') + '" data-' + (f === 'story' ? 'prstf="title"' : 'prf="feature"') +
      ' title="Double-click to rename">' + esc(s || (f === 'story' ? 'Story' : 'Name')) + '</span>';
  }
  function prCardHtml(it) {
    var fields = prCardFields().map(function (c) {
      // no label — a quiet "No description"-style hint fills the empty field
      // read-only on the card: the panel (or the Scoping grid) edits these
      return '<div class="pr-rich" data-prsc="' + c[0] + '" data-ph="No ' + esc(c[1].toLowerCase()) + '">' +
        richDisplay(RM.scopeValue(it, c[0])) + '</div>';
    }).join('');
    var wsColor = it.workstream ? RM.colorForWs(state, it.workstream) : RM.defaultWsColor(state);
    // Story level: the card lists its stories — title inline-editable, own
    // priority chip — under the feature's own fields
    var stories = prioLevel === 'story' && !it.milestone
      ? '<div class="pr-stories">' + (it.stories || []).map(function (st) {
          return '<div class="pr-story" data-prst="' + st.id + '">' +
            '<i data-lucide="corner-down-right" class="pr-st-ico"></i>' +
            prTitleHtml(st.title, 'story', 'pr-st-title') +
            ['size', 'pri', 'risk', 'dur'].map(function (k) { return storyChipHtml(k, st, 'data-prstact'); }).join('') +
            '</div>';
        }).join('') +
        ((it.stories || []).length ? '' : // add via the story context menu once one exists
          '<button class="pr-st-add" data-prstadd="1"><i data-lucide="plus"></i>' + esc('Add ' + lvl('story')) + '</button>') +
        '</div>'
      : '';
    return '<div class="sp-card pr-card' + (isSel(it.id) && !selStory ? ' selected' : '') + '" data-prcard="' + it.id + '" style="--ws-c:#' + wsColor + '">' +
      '<div class="pr-head">' + typeGlyphHtml(it, 'feature') +
      prTitleHtml(it.feature, 'feature', 'pr-title') + prFlagHtml(it) + '</div>' +
      fields + stories +
      '<div class="pr-chips">' +
      (prChipOn('size') // the column field's chip is redundant; hidden chips come from the Fields menu
        ? '<span class="r-size' + sizeRoCls(it) + '" tabindex="0" role="button" data-pract="size" title="' + sizeChipTitle(it) + '">' +
          (it.size ? esc(it.size) : '·') + '</span>' : '') +
      (prChipOn('priority')
        ? '<span class="r-risk pri' + (priChipHasValue(it) ? ' has-risk' : '') + priTierClass(it.priority) +
          '" tabindex="0" role="button" data-pract="priority" title="' + esc(priChipTitle(it)) + '">' +
          (priChipContent(it) || '·') + '</span>' : '') +
      (prChipOn('risk') ? itemChipHtml('risk', it, 'data-pract') : '') +
      (prChipOn('dur') ? itemChipHtml('dur', it, 'data-pract') : '') +
      (prChipOn('epic')
        ? '<span class="pr-chip" tabindex="0" role="button" data-pract="epic" title="Epic">' +
          '<i data-lucide="' + (RM.iconForEpic(state, it.epic) || 'tag') + '"></i>' + esc(it.epic || '—') + '</span>' : '') +
      (prChipOn('ws')
        ? '<span class="pr-chip" tabindex="0" role="button" data-pract="ws" title="Workstream">' +
          '<span class="dd-dot" style="background:#' + wsColor + '"></span>' + esc(it.workstream || RM.defaultWsName(state)) + '</span>' : '') +
      '</div></div>';
  }
  // ---- Story level: a board of story cards in columns of one story field
  // (priority / size / risk); dragging a card sets that field on the story
  var PR_STORY_COLS = { priority: ['arrow-down-wide-narrow', 'Priority'], size: ['ruler', 'Size'], risk: ['triangle-alert', 'Risk'] };
  function prStoryColKinds() {
    var out = [];
    if (RM.priorityEnabled(state, 'story')) out.push('priority');
    if (RM.sizingEnabled(state, 'story')) out.push('size');
    if (RM.riskEnabled(state, 'story')) out.push('risk');
    return out;
  }
  function prStoryColLabel(v) {
    if (v === '') return 'Unset';
    if (prioStoryCol === 'priority') return priorityValueLabel(v, 'story');
    if (prioStoryCol === 'risk') return riskValueLabel(v, 'story');
    return v;
  }
  // [{ key, name }] — the field's ladder in scheme order, then Unset
  function prStoryColumns() {
    var vals = prioStoryCol === 'priority' ? RM.priorityOrderOf(state, 'story')
      : prioStoryCol === 'risk' ? RM.riskOrderOf(state, 'story') : RM.sizeOrderOf(state, 'story');
    return vals.map(function (v) { return { key: v, name: prStoryColLabel(v) }; }).concat(prioHideUnset ? [] : [{ key: '', name: 'Unset' }]);
  }
  // the story's value for the column field; anything the ladder does not
  // know (an old scheme's value) files under Unset instead of vanishing
  function prStoryVal(st) {
    var v = st[prioStoryCol] || '';
    if (!v) return '';
    return prStoryColumns().some(function (c) { return c.key === v; }) ? v : '';
  }
  // a story shows when its feature passes the dropdowns and either the story
  // title or the feature itself matches the text filter
  function prStoryMatches(it, st) {
    if (prioFPhase != null && it.phaseId !== prioFPhase) return false;
    if (prioFEpic != null && (it.epic || '') !== prioFEpic) return false;
    if (prioFWs != null && (it.workstream || '') !== prioFWs) return false;
    if (!filterText) return true;
    var pq = filterText.toLowerCase();
    if (pq.charAt(0) === '#') return tagsHit(st.tags, pq.slice(1)) || matchesFilter(it, true);
    return (st.title || '').toLowerCase().indexOf(pq) !== -1 || tagsHit(st.tags, pq) || matchesFilter(it, true);
  }
  // every visible story as { it, st } pairs, in document order
  function prStoryPairs(items) {
    var out = [];
    items.forEach(function (it) {
      if (it.milestone) return;
      (it.stories || []).forEach(function (st) { if (prStoryMatches(it, st)) out.push({ it: it, st: st }); });
    });
    return out;
  }
  function prSortStories(pairs) {
    var rank;
    if (prioSort === 'title') {
      return pairs.map(function (x, i) { return { x: x, i: i }; })
        .sort(function (a, b) { return (a.x.st.title || '').localeCompare(b.x.st.title || '') || (a.i - b.i); })
        .map(function (y) { return y.x; });
    } else if (prioSort === 'size' && RM.sizingEnabled(state, 'story')) {
      var so = RM.sizeOrderOf(state, 'story');
      rank = function (st) { var i2 = st.size ? so.indexOf(st.size) : -1; return i2 === -1 ? Infinity : -i2; };
    } else if (prioSort === 'priority' && RM.priorityEnabled(state, 'story')) {
      var po = RM.priorityOrderOf(state, 'story');
      rank = function (st) { var i3 = st.priority ? po.indexOf(st.priority) : -1; return i3 === -1 ? Infinity : i3; };
    } else return pairs.slice();
    return pairs.map(function (x, i) { return { x: x, i: i }; })
      .sort(function (a, b) { return (rank(a.x.st) - rank(b.x.st)) || (a.i - b.i); })
      .map(function (y) { return y.x; });
  }
  function prFlagHtml(x) {
    return x.flag ? '<span class="pr-flag" title="' + esc(flagTitle(x)) + '"><i data-lucide="flag"></i></span>' : '';
  }
  function prStoryCardHtml(it, st) {
    var wsColor = it.workstream ? RM.colorForWs(state, it.workstream) : RM.defaultWsColor(state);
    var chips = ['size', 'pri', 'risk', 'dur'].filter(function (k) {
      return !(k === 'size' && prioStoryCol === 'size');
    }).map(function (k) { return storyChipHtml(k, st, 'data-prstact'); }).join('');
    return '<div class="sp-card pr-card pr-stcard' + (st.done ? ' done' : '') + (selectedId === it.id && selStory === st.id ? ' selected' : '') + '" data-prcard="' + it.id + '" data-prst="' + st.id + '" style="--ws-c:#' + wsColor + '">' +
      '<div class="pr-stfeat" title="Feature">' +
      '<i data-lucide="' + (RM.iconForEpic(state, it.epic) || 'tag') + '"></i>' +
      '<span class="pr-stfeatname">' + esc(it.feature || '(untitled)') + '</span></div>' +
      '<div class="pr-head">' + typeGlyphHtml(st, 'story', it) +
      prTitleHtml(st.title, 'story', 'pr-title pr-st-title') + prFlagHtml(st) + '</div>' +
      '<div class="pr-chips">' + chips + '</div></div>';
  }
  function prStoryColHtml(col, pairs, laneAttr) {
    var mine = prSortStories(pairs.filter(function (x) { return prStoryVal(x.st) === col.key; }));
    return '<div class="sp-col" data-prcol="' + esc(col.key) + '"' + (laneAttr || '') + '>' +
      '<div class="sp-colbody">' + (mine.length ? mine.map(function (x) { return prStoryCardHtml(x.it, x.st); }).join('') : '<div class="pr-empty">No items</div>') + '</div></div>';
  }
  // ---- Feature level columns: the phases (default) or one feature field's
  // ladder — priority / size / risk — plus Unset; dragging a card sets it
  var PR_FEAT_COLS = { phase: ['flag', 'Phase'], priority: ['arrow-down-wide-narrow', 'Priority'], size: ['ruler', 'Size'], risk: ['triangle-alert', 'Risk'] };
  function prFeatColKinds() {
    var out = ['phase'];
    if (RM.priorityEnabled(state) && RM.prioritySchemeOf(state) !== 'rice') out.push('priority'); // RICE is computed, no ladder
    if (RM.sizingEnabled(state) && !RM.sizeRollup(state)) out.push('size');
    if (RM.riskEnabled(state) && RM.riskSchemeOf(state) !== 'auto') out.push('risk');
    return out;
  }
  // Columns menu tail: show / hide the Unset column (a ladder's catch-all)
  function prUnsetToggleItems() {
    return [{ sep: true }, { icon: 'eye-off', label: 'Unset column', checked: !prioHideUnset, fn: function () {
      prioHideUnset = !prioHideUnset;
      saveLocal();
      render();
    } }];
  }
  function prFeatColumns() {
    if (prioFeatCol === 'phase') return state.phases.map(function (p) { return { key: p.id, name: p.name }; });
    var vals = prioFeatCol === 'priority' ? RM.priorityOrderOf(state)
      : prioFeatCol === 'risk' ? RM.riskOrderOf(state) : RM.sizeOrderOf(state);
    return vals.map(function (v) {
      return { key: v, name: prioFeatCol === 'priority' ? priorityValueLabel(v) : prioFeatCol === 'risk' ? riskValueLabel(v) : v };
    }).concat(prioHideUnset ? [] : [{ key: '', name: 'Unset' }]);
  }
  // the feature's value for the column field; off-ladder values file under Unset
  function prFeatVal(it) {
    if (prioFeatCol === 'phase') return it.phaseId;
    var v = it[prioFeatCol] || '';
    if (!v) return '';
    return prFeatColumns().some(function (c) { return c.key === v; }) ? v : '';
  }
  function prColHtml(col, items, laneAttr) {
    var mine = prSortItems(items.filter(function (it) { return prFeatVal(it) === col.key && prMatches(it); }));
    return '<div class="sp-col" data-prcol="' + esc(col.key) + '"' + (laneAttr || '') + '>' +
      '<div class="sp-colbody">' + (mine.length ? mine.map(prCardHtml).join('') : '<div class="pr-empty">No items</div>') + '</div>' +
      (prioFeatCol === 'phase' && !mine.length // a new feature needs a phase to land in; a filled column adds via the context menu
        ? '<button class="pr-add" data-pradd="' + col.key + '"' + (laneAttr || '') + '><i data-lucide="plus"></i> Add</button>' : '') +
      '</div>';
  }
  function prLaneKey(it) { return prioGroup === 'ws' ? (it.workstream || '') : (it.epic || ''); }
  var PR_GROUPS = { none: ['square-kanban', 'None'], ws: ['layers', 'Workstreams'], epic: ['tag', 'Epics'] };
  var PR_SORTS = { priority: ['arrow-down-wide-narrow', 'Priority'], doc: ['list', 'Document order'],
    title: ['arrow-down-a-z', 'Title'], size: ['ruler', 'Size'] };
  function prDdLabel(pair) { return '<i data-lucide="' + pair[0] + '"></i>' + esc(pair[1]); }
  function renderPrioPage() {
    var host = $('#prioView');
    if (!host) return;
    if (prioGroup === 'ws' && !state.meta.workstreamsEnabled) prioGroup = 'none';
    var storyMode = prioLevel === 'story';
    var stKinds = storyMode ? prStoryColKinds() : [];
    if (storyMode && stKinds.indexOf(prioStoryCol) === -1 && stKinds.length) prioStoryCol = stKinds[0];
    var ftKinds = storyMode ? [] : prFeatColKinds();
    if (!storyMode && ftKinds.indexOf(prioFeatCol) === -1) prioFeatCol = 'phase';
    if (prioFPhase != null && !state.phases.some(function (p) { return p.id === prioFPhase; })) prioFPhase = null;
    var phases = state.phases;
    var grouped = prioGroup !== 'none';
    // columns: phases (or a feature field's ladder) for features, a story field's ladder for stories
    var cols = storyMode ? prStoryColumns() : prFeatColumns();
    var allPairs = storyMode ? prStoryPairs(state.items) : null;
    var head = (grouped ? '<div class="pr-corner"></div>' : '') + cols.map(function (c) {
      var n = storyMode
        ? allPairs.filter(function (x) { return prStoryVal(x.st) === c.key; }).length
        : state.items.filter(function (it) { return prFeatVal(it) === c.key && prMatches(it); }).length;
      return '<div class="pr-phhd">' + esc(c.name) + '<span class="pr-lanect">' + n + '</span></div>';
    }).join('');
    var rows;
    if (storyMode && !stKinds.length) {
      rows = '<div class="p-none" style="grid-column:1/-1;padding:30px">Turn on a story priority, size or risk scheme in Setup to lay stories out in columns.</div>';
    } else if (storyMode && !grouped) {
      rows = cols.map(function (c) { return prStoryColHtml(c, allPairs); }).join('');
    } else if (storyMode) {
      var stLanes = prioGroup === 'ws'
        ? allWorkstreams().map(function (w) { return { key: w, name: w, dot: RM.colorForWs(state, w) }; })
            .concat([{ key: '', name: RM.defaultWsName(state), dot: RM.defaultWsColor(state) }]) // the catch-all lane sits last
        : allEpics().map(function (ep) { return { key: ep, name: ep, icon: RM.iconForEpic(state, ep) || 'tag' }; })
            .concat([{ key: '', name: 'No epic' }]);
      stLanes = stLanes.filter(function (ln) { return allPairs.some(function (x) { return prLaneKey(x.it) === ln.key; }); });
      rows = stLanes.map(function (ln) {
        var lanePairs = allPairs.filter(function (x) { return prLaneKey(x.it) === ln.key; });
        var laneAttr = ' data-prlane="' + esc(ln.key) + '"';
        return '<div class="pr-lanecell">' +
          (ln.dot ? '<span class="dd-dot" style="background:#' + ln.dot + '"></span>' : '') +
          (ln.icon ? '<i data-lucide="' + ln.icon + '"></i>' : '') +
          '<span class="pr-lanename">' + esc(ln.name) + '</span>' +
          '<span class="pr-lanect">' + lanePairs.length + '</span></div>' +
          cols.map(function (c) { return prStoryColHtml(c, lanePairs, laneAttr); }).join('');
      }).join('') || '<div class="p-none" style="grid-column:1/-1;padding:30px">Nothing matches.</div>';
    } else if (!grouped) {
      rows = cols.map(function (c) { return prColHtml(c, state.items); }).join('');
    } else {
      var lanes = prioGroup === 'ws'
        ? allWorkstreams().map(function (w) { return { key: w, name: w, dot: RM.colorForWs(state, w) }; })
            .concat([{ key: '', name: RM.defaultWsName(state), dot: RM.defaultWsColor(state) }]) // the catch-all lane sits last
        : allEpics().map(function (ep) { return { key: ep, name: ep, icon: RM.iconForEpic(state, ep) || 'tag' }; })
            .concat([{ key: '', name: 'No epic' }]);
      lanes = lanes.filter(function (ln) {
        return state.items.some(function (it) { return prLaneKey(it) === ln.key && prMatches(it); });
      });
      rows = lanes.map(function (ln) {
        var laneItems = state.items.filter(function (it) { return prLaneKey(it) === ln.key && prMatches(it); });
        var laneAttr = ' data-prlane="' + esc(ln.key) + '"';
        return '<div class="pr-lanecell">' +
          (ln.dot ? '<span class="dd-dot" style="background:#' + ln.dot + '"></span>' : '') +
          (ln.icon ? '<i data-lucide="' + ln.icon + '"></i>' : '') +
          '<span class="pr-lanename">' + esc(ln.name) + '</span>' +
          '<span class="pr-lanect">' + laneItems.length + '</span></div>' +
          cols.map(function (c) { return prColHtml(c, laneItems, laneAttr); }).join('');
      }).join('') || '<div class="p-none" style="grid-column:1/-1;padding:30px">Nothing matches.</div>';
    }
    // an active filter wears what it picked: the epic's own icon, the
    // workstream's color dot, and a blue outline
    var fPhP = prioFPhase == null ? null : state.phases.filter(function (p) { return p.id === prioFPhase; })[0];
    var phaseFilterDd = state.phases.length > 1
      ? '<button class="dd-btn' + (fPhP ? ' pr-on' : '') + '" data-prdd="fphase" title="Filter by phase">' +
        '<span class="dd-label"><i data-lucide="milestone"></i>' + (fPhP ? esc(fPhP.name) : 'All phases') + '</span>' +
        '<i data-lucide="chevron-down"></i></button>' : '';
    var epicFilterDd = allEpics().length
      ? '<button class="dd-btn' + (prioFEpic == null ? '' : ' pr-on') + '" data-prdd="fepic" title="Filter by epic">' +
        '<span class="dd-label">' + (prioFEpic == null ? '<i data-lucide="tag"></i>All epics'
          : '<i data-lucide="' + (prioFEpic ? (RM.iconForEpic(state, prioFEpic) || 'tag') : 'tag') + '"></i>' + esc(prioFEpic || 'No epic')) + '</span>' +
        '<i data-lucide="chevron-down"></i></button>' : '';
    var wsFilterDd = state.meta.workstreamsEnabled
      ? '<button class="dd-btn' + (prioFWs == null ? '' : ' pr-on') + '" data-prdd="fws" title="Filter by workstream">' +
        '<span class="dd-label">' + (prioFWs == null ? '<i data-lucide="layers"></i>All workstreams'
          : '<span class="dd-dot" style="background:#' + (prioFWs ? RM.colorForWs(state, prioFWs) : RM.defaultWsColor(state)) + '"></span>' +
            esc(prioFWs || RM.defaultWsName(state))) + '</span>' +
        '<i data-lucide="chevron-down"></i></button>' : '';
    // which field the columns represent (phases or a ladder for features; a ladder for stories)
    var colsDd = (storyMode ? stKinds.length : ftKinds.length > 1)
      ? '<span class="pr-lab">Columns</span>' +
        '<button class="dd-btn" data-prdd="cols" title="' + (storyMode ? 'Which story field the columns represent' : 'Which feature field the columns represent') + '">' +
        '<span class="dd-label">' + prDdLabel(storyMode ? PR_STORY_COLS[prioStoryCol] : PR_FEAT_COLS[prioFeatCol]) + '</span><i data-lucide="chevron-down"></i></button>'
      : '';
    host.innerHTML = '<div class="sp-page">' +
      '<div class="pr-bar">' +
      levelBtnHtml('data-prdd="level"', dmModes(), prioLevel) +
      '<span class="filter-wrap pr-filter"><i data-lucide="search" class="filter-ico" aria-hidden="true"></i>' +
      '<input id="prFilter" type="search" placeholder="Filter cards" aria-label="Filter cards" value="' + esc(filterText) + '">' +
      '<kbd class="kbd filter-kbd" aria-hidden="true">⌘F</kbd></span>' +
      phaseFilterDd + epicFilterDd + wsFilterDd +
      '<div class="pr-settings">' +
      colsDd +
      '<span class="pr-lab">Group</span>' +
      '<button class="dd-btn" data-prdd="group" title="Swimlanes">' +
      '<span class="dd-label">' + prDdLabel(PR_GROUPS[prioGroup]) + '</span><i data-lucide="chevron-down"></i></button>' +
      '<span class="pr-lab">Sort</span>' +
      '<button class="dd-btn" data-prdd="sort" title="Order cards inside each column">' +
      '<span class="dd-label">' + prDdLabel(PR_SORTS[prioSort]) + '</span><i data-lucide="chevron-down"></i></button>' +
      (storyMode ? '' : '<button id="prFieldsBtn" title="Choose which fields cards show"><i data-lucide="list-checks"></i>Fields</button>') +
      '</div></div>' +
      '<div class="pr-table' + (storyMode ? ' pr-stboard' : '') + '" style="grid-template-columns:' + (grouped ? '170px ' : '') +
      'repeat(' + cols.length + ', minmax(' + (storyMode ? '200px' : '230px') + ', 1fr))">' + head + rows + '</div></div>';
    if (window.lucide) lucide.createIcons();
  }

  var prFilterTimer = null;
  $('#prioView').addEventListener('input', function (e) {
    if (e.target.id !== 'prFilter') return;
    var v = e.target.value;
    clearTimeout(prFilterTimer);
    prFilterTimer = setTimeout(function () {
      filterText = v.trim();
      render();
      var nf = $('#prFilter'); // the render rebuilt the input — hand focus back
      if (nf) {
        nf.focus();
        nf.setSelectionRange(nf.value.length, nf.value.length);
      }
    }, 120);
  });
  // double-click a card title (or Rename… in the card menu) to edit it in
  // place. The host is re-queried by id: selecting a card re-renders the
  // board, so the element the click started on may already be gone.
  function prStartRename(cardId, stId, inRow) {
    var host = $(inRow
      ? '#prioView [data-prcard="' + cardId + '"] .pr-story[data-prst="' + stId + '"]'
      : stId ? '#prioView .pr-stcard[data-prst="' + stId + '"]'
        : '#prioView .pr-card[data-prcard="' + cardId + '"]:not(.pr-stcard)');
    if (!host) return;
    var span = host.querySelector(inRow ? '.pr-st-title' : '.pr-head .pr-title');
    if (!span || span.tagName === 'INPUT') return;
    var blank = span.classList.contains('pr-ph');
    startInlineEdit(span, function (val) {
      if (stId) {
        commit('rename story', function (s) {
          var st = storyById(RM.itemById(s, cardId) || {}, stId);
          if (st) st.title = val;
        });
      } else {
        commit('rename', function (s) {
          var x = RM.itemById(s, cardId);
          if (x) x.feature = val;
        });
      }
    });
    var inp = host.querySelector('input.st-add-input');
    if (inp) {
      if (stId) inp.dataset.prstf = 'title'; else inp.dataset.prf = 'feature';
      inp.className = 'st-add-input ' + (inRow ? 'pr-st-title' : stId ? 'pr-title pr-st-title' : 'pr-title');
      if (blank) inp.value = '';
      inp.select();
    }
  }
  $('#prioView').addEventListener('dblclick', function (e) {
    var t = e.target.closest('.pr-title,.pr-st-title');
    if (!t || t.tagName === 'INPUT') return;
    var card = t.closest('[data-prcard]');
    if (!card) return;
    e.preventDefault();
    var stRow = t.closest('.pr-story');
    if (stRow) prStartRename(card.dataset.prcard, stRow.dataset.prst, true);
    else prStartRename(card.dataset.prcard, card.dataset.prst || null, false);
  });
  $('#prioView').addEventListener('keydown', function (e) {
    var t = e.target;
    if (t.id === 'prFilter' && e.key === 'Escape') {
      e.stopPropagation();
      t.value = '';
      filterText = '';
      render();
      return;
    }
  });
  $('#prioView').addEventListener('click', function (e) {
    if (dragConsumedClick) { dragConsumedClick = false; return; } // the click that trails a card drag
    var dd = e.target.closest('[data-prdd]');
    if (dd) {
      var kind = dd.dataset.prdd;
      if (kind === 'level') {
        openLevelMenu(dd, dmModes(), prioLevel, function (mode) {
          if (prioLevel !== mode) { prioLevel = mode; saveLocal(); render(); }
        });
      } else if (kind === 'group') {
        openDropdown(dd, Object.keys(PR_GROUPS).map(function (k) {
          if (k === 'ws' && !state.meta.workstreamsEnabled) return null;
          return { icon: PR_GROUPS[k][0], label: esc(PR_GROUPS[k][1]), checked: prioGroup === k, fn: function () {
            prioGroup = k; saveLocal(); render();
          } };
        }).filter(Boolean));
      } else if (kind === 'sort') {
        openDropdown(dd, Object.keys(PR_SORTS).map(function (k) {
          return { icon: PR_SORTS[k][0], label: esc(PR_SORTS[k][1]), checked: prioSort === k, fn: function () {
            prioSort = k; saveLocal(); render();
          } };
        }));
      } else if (kind === 'cols' && prioLevel === 'story') {
        openDropdown(dd, prStoryColKinds().map(function (k) {
          return { icon: PR_STORY_COLS[k][0], label: PR_STORY_COLS[k][1], checked: prioStoryCol === k, fn: function () {
            if (prioStoryCol !== k) { prioStoryCol = k; saveLocal(); render(); }
          } };
        }).concat(prUnsetToggleItems()));
      } else if (kind === 'cols') {
        openDropdown(dd, prFeatColKinds().map(function (k) {
          return { icon: PR_FEAT_COLS[k][0], label: PR_FEAT_COLS[k][1], checked: prioFeatCol === k, fn: function () {
            if (prioFeatCol !== k) { prioFeatCol = k; saveLocal(); render(); }
          } };
        }).concat(prioFeatCol === 'phase' ? [] : prUnsetToggleItems())); // phases have no Unset
      } else if (kind === 'fphase') {
        openDropdown(dd, [{ label: '<i>All phases</i>', checked: prioFPhase == null, fn: function () { prioFPhase = null; render(); } }]
          .concat(state.phases.map(function (p) {
            return { label: esc(p.name), checked: prioFPhase === p.id, fn: function () { prioFPhase = p.id; render(); } };
          })));
      } else if (kind === 'fepic') {
        openDropdown(dd, [{ label: '<i>All epics</i>', checked: prioFEpic == null, fn: function () { prioFEpic = null; render(); } },
          { label: '<i>— no epic —</i>', checked: prioFEpic === '', fn: function () { prioFEpic = ''; render(); } }]
          .concat(allEpics().map(function (ep) {
            return { icon: RM.iconForEpic(state, ep) || 'tag', label: esc(ep), checked: prioFEpic === ep, fn: function () {
              prioFEpic = ep; render();
            } };
          })));
      } else if (kind === 'fws') {
        openDropdown(dd, [{ label: '<i>All workstreams</i>', checked: prioFWs == null, fn: function () { prioFWs = null; render(); } },
          { label: esc(RM.defaultWsName(state)) + ' <i>(default)</i>', dot: '#' + RM.defaultWsColor(state),
            checked: prioFWs === '', fn: function () { prioFWs = ''; render(); } }]
          .concat(allWorkstreams().map(function (w) {
            return { label: esc(w), dot: '#' + RM.colorForWs(state, w), checked: prioFWs === w, fn: function () {
              prioFWs = w; render();
            } };
          })));
      }
      return;
    }
    if (e.target.closest('#prFieldsBtn')) {
      // scope (text) columns opt in; the chips (size / priority / risk /
      // duration / epic / workstream) are on unless the user hides them
      var fItems = scopeCols().filter(function (c) { return colShowsOn(c[0], 'feature'); }).map(function (c) {
        return { label: esc(c[1]), checked: prioFields.indexOf(c[0]) !== -1, fn: function () {
          var at = prioFields.indexOf(c[0]);
          if (at === -1) prioFields.push(c[0]); else prioFields.splice(at, 1);
          saveLocal();
          render();
        } };
      });
      var chipItems = Object.keys(PR_CHIPS).filter(prChipAvail).map(function (k) {
        return { label: PR_CHIPS[k], checked: prioChipHide.indexOf(k) === -1, fn: function () {
          var at2 = prioChipHide.indexOf(k);
          if (at2 === -1) prioChipHide.push(k); else prioChipHide.splice(at2, 1);
          saveLocal();
          render();
        } };
      });
      if (fItems.length && chipItems.length) fItems.push({ sep: true });
      openDropdown(e.target.closest('#prFieldsBtn'), fItems.concat(chipItems));
      return;
    }
    // story rows (Story level): priority chip, add-story button
    var stChipP = e.target.closest('[data-prstact]');
    if (stChipP) {
      var stRowP = stChipP.closest('[data-prst]'), stCardP = stChipP.closest('[data-prcard]');
      if (stRowP && stCardP) storyChipAction(stChipP.dataset.prstact, stChipP, stCardP.dataset.prcard, stRowP.dataset.prst);
      return;
    }
    var stAdd = e.target.closest('[data-prstadd]');
    if (stAdd) {
      var stCardA = stAdd.closest('[data-prcard]');
      if (!stCardA) return;
      var addId = stCardA.dataset.prcard, newSt = RM.uid('s');
      commit('add story', function (s) {
        var x = RM.itemById(s, addId);
        if (x) x.stories.push({ id: newSt, title: '', done: false, num: RM.nextNum(s) });
      });
      prStartRename(addId, newSt, true); // the new line's title opens for typing
      return;
    }
    var chip = e.target.closest('[data-pract]');
    if (chip) {
      var cardC = chip.closest('[data-prcard]');
      if (!cardC) return;
      var cid = cardC.dataset.prcard;
      var itC = RM.itemById(state, cid);
      if (!itC) return;
      var act = chip.dataset.pract;
      if (act === 'size') {
        if (sizeLocked(itC)) return;
        openDropdown(chip, [{ label: '<i>no size</i>', checked: !itC.size, fn: function () {
          setItemSize(cid, null);
        } }].concat(RM.sizeOrderOf(state).map(function (sz) {
          return { label: esc(sz) + ' <small>' + sizeHuman(sz) + '</small>', checked: itC.size === sz, fn: function () {
            setItemSize(cid, sz);
          } };
        })));
      } else if (act === 'priority') openPriorityEditor(chip, cid);
      else if (act === 'risk' || act === 'dur') itemChipAction(act, chip, cid);
      else if (act === 'epic') openDropdown(chip, setEpicMenu(cid, true));
      else if (act === 'ws') {
        openDropdown(chip, wsMenuItems(cid, function () {
          return $('#prioView [data-prcard="' + cid + '"] [data-pract="ws"]');
        }));
      }
      return;
    }
    // a plain click on a card opens the panel on it (clicking the selected
    // card again keeps the selection)
    var cardK = e.target.closest('[data-prcard]');
    if (cardK && !e.target.closest('.pr-story,input,textarea,select,button,[contenteditable="true"],[data-pract],[data-prstact]')) {
      var kid = cardK.dataset.prcard, kst = cardK.dataset.prst || null;
      panelOpen = true;
      if (kst) selectStory(kid, kst); else select(kid);
      return;
    }
    var add = e.target.closest('[data-pradd]');
    if (add) {
      var lane = add.getAttribute('data-prlane');
      var group = prioGroup;
      addFeature(add.dataset.pradd);
      // in a swimlane the new card belongs to the lane it was added in
      if (lane && selectedId) {
        var nid = selectedId;
        commit(group === 'ws' ? 'workstream' : 'epic', function (s) {
          var x = RM.itemById(s, nid);
          if (x) { if (group === 'ws') x.workstream = lane; else x.epic = lane; }
        });
      }
    }
  });
  $('#prioView').addEventListener('contextmenu', function (e) {
    var card = e.target.closest('[data-prcard]');
    if (!card || e.target.closest('input,textarea,select,[contenteditable="true"]')) return;
    e.preventDefault();
    e.stopPropagation(); // the app-chrome theme menu must not replace this
    var cid = card.dataset.prcard;
    var itX = RM.itemById(state, cid);
    if (!itX) return;
    var cx = e.clientX, cy = e.clientY;
    var stRowX = e.target.closest('.pr-story'); // a story line on a feature card
    if (stRowX) {
      var stRowId = stRowX.dataset.prst;
      openContextMenu(cx, cy, [
        { icon: 'pencil', label: 'Rename…', fn: function () { prStartRename(cid, stRowId, true); } },
        moveStoryFeatureEntry(cx, cy, cid, stRowId),
        { sep: true },
        storyInsertEntries(cid, stRowId)[0], storyInsertEntries(cid, stRowId)[1],
        flagMenuEntries(cid, stRowId)[0], flagMenuEntries(cid, stRowId)[1] || null, noAutoEntry(cid, stRowId),
        { icon: 'copy', label: 'Duplicate story', fn: function () { duplicateStory(cid, stRowId); } },
        { icon: 'trash-2', label: 'Delete story', danger: true, fn: function () {
          commit('delete story', function (s) {
            var t = RM.itemById(s, cid);
            if (t) t.stories = (t.stories || []).filter(function (st) { return st.id !== stRowId; });
          });
        } }
      ]);
      return;
    }
    if (card.dataset.prst) { // a story card: jump to its feature, or delete the story
      var stIdX = card.dataset.prst;
      openContextMenu(cx, cy, [
        { icon: 'pencil', label: 'Rename…', fn: function () { prStartRename(cid, stIdX, false); } },
        { icon: 'rows-3', label: 'Go to feature', fn: function () { select(cid, true); } },
        moveStoryFeatureEntry(cx, cy, cid, stIdX),
        { sep: true },
        storyInsertEntries(cid, stIdX)[0], storyInsertEntries(cid, stIdX)[1],
        flagMenuEntries(cid, stIdX)[0], flagMenuEntries(cid, stIdX)[1] || null, noAutoEntry(cid, stIdX),
        { icon: 'copy', label: 'Duplicate story', fn: function () { duplicateStory(cid, stIdX); } },
        { icon: 'trash-2', label: 'Delete story', danger: true, fn: function () {
          commit('delete story', function (s) {
            var t = RM.itemById(s, cid);
            if (t) t.stories = (t.stories || []).filter(function (st) { return st.id !== stIdX; });
          });
        } }
      ]);
      return;
    }
    openContextMenu(cx, cy, [
      { icon: 'pencil', label: 'Rename…', fn: function () { prStartRename(cid, null, false); } },
      { sep: true },
      { icon: 'folder-input', label: 'Move to phase…', fn: function () { openContextMenu(cx, cy, movePhaseMenu(cid)); } },
      { icon: 'tag', label: 'Set epic…', fn: function () { openContextMenu(cx, cy, setEpicMenu(cid, false)); } },
      state.meta.workstreamsEnabled
        ? { icon: 'layers', label: 'Set workstream…', fn: function () {
            openContextMenu(cx, cy, wsMenuItems(cid, function () {
              return $('#prioView [data-prcard="' + cid + '"] [data-pract="ws"]');
            }));
          } }
        : null,
      { sep: true },
      flagMenuEntries(cid, null)[0], flagMenuEntries(cid, null)[1] || null, noAutoEntry(cid, null),
      { sep: true },
      { icon: 'trash-2', label: 'Delete…', danger: true, fn: function () { deleteItemConfirm(cid); } }
    ].filter(Boolean));
  });
  // cards drag between horizons (phase columns) — and across swimlanes
  $('#prioView').addEventListener('pointerdown', function (e) {
    if (e.button !== 0 || drag) return;
    if (e.target.closest('input,textarea,button,select,[contenteditable="true"],[data-pract]')) return;
    var card = e.target.closest('[data-prcard]');
    if (!card) return;
    drag = { kind: 'prcard', id: card.dataset.prcard, st: card.dataset.prst || null, el: card,
      x0: e.clientX, y0: e.clientY, moved: false, ghost: null, colEl: null };
    e.preventDefault();
  });
  function prCardDragMove(e) {
    if (!drag.ghost) {
      drag.ghost = document.createElement('div');
      drag.ghost.className = 'sp-card sp-card-ghost';
      var t = RM.itemById(state, drag.id);
      var stG = drag.st && t ? storyById(t, drag.st) : null;
      drag.ghost.textContent = stG ? (stG.title || 'Story') : ((t && t.feature) || 'Card');
      document.body.appendChild(drag.ghost);
      drag.el.classList.add('sp-dragging');
    }
    drag.ghost.style.left = (e.clientX + 10) + 'px';
    drag.ghost.style.top = (e.clientY + 8) + 'px';
    var under = (document.elementFromPoint && document.elementFromPoint(e.clientX, e.clientY)) || e.target;
    drag.colEl = under && under.closest ? under.closest('[data-prcol]') : null;
    $$('#prioView .sp-col').forEach(function (c) { c.classList.toggle('drop', c === drag.colEl); });
  }
  function prCardDragEnd(d) {
    if (d.ghost) d.ghost.remove();
    d.el.classList.remove('sp-dragging');
    $$('#prioView .sp-col').forEach(function (c) { c.classList.remove('drop'); });
    var t = RM.itemById(state, d.id);
    var toId = d.colEl && d.colEl.dataset.prcol;
    var lane = d.colEl ? d.colEl.getAttribute('data-prlane') : null; // null = no swimlanes
    var group = prioGroup;
    if (d.st) { // story card: the column IS the field value ('' = unset)
      var stD = t ? storyById(t, d.st) : null;
      var field = prioStoryCol;
      if (toId == null || !stD || (stD[field] || '') === toId) { render(); return; }
      commit('story ' + field, function (s) {
        var st2 = storyById(RM.itemById(s, d.id) || {}, d.st);
        if (st2) st2[field] = toId || null;
      });
      return;
    }
    if (prioFeatCol !== 'phase') { // the column IS the field value ('' = unset)
      var ff = prioFeatCol;
      var laneSame = lane == null || prLaneKey(t) === lane;
      if (toId == null || !t || ((t[ff] || '') === toId && laneSame)) { render(); return; }
      commit(ff, function (s) {
        var x = RM.itemById(s, d.id);
        if (!x) return;
        if ((x[ff] || '') !== toId) {
          x[ff] = toId || null;
          if (ff === 'size' && toId && isScheduled(x) && !x.locked) x.durDays = RM.stretchSpan(s.meta, x.startDay, RM.sizeDays(s, toId)); // as setItemSize
          if (ff === 'risk') x.riskDays = 0; // as setItemRisk
        }
        if (lane != null) { if (group === 'ws') x.workstream = lane; else x.epic = lane; }
      });
      return;
    }
    if (!toId || !t || (t.phaseId === toId && (lane == null || prLaneKey(t) === lane))) { render(); return; }
    commit('move card', function (s) {
      var x = RM.itemById(s, d.id);
      s.items = s.items.filter(function (y) { return y.id !== d.id; });
      x.phaseId = toId;
      if (lane != null) {
        if (group === 'ws') x.workstream = lane;
        else x.epic = lane;
      }
      var lastIdx = -1;
      s.items.forEach(function (y, i2) { if (y.phaseId === toId) lastIdx = i2; });
      s.items.splice(lastIdx + 1, 0, x);
    });
  }

  // ------------------------------------------------------------ sprinting page
  // Sprint by sprint: a sidebar of sprints (Unscheduled last) and ONE
  // scrolling page of sections, rows per sprint — features or stories.
  // Dragging a row to another sprint moves it on the timeline (start = that
  // sprint's first day, span kept, stories ride along) and reorders the
  // shared items array, so the new order shows in every view.
  var sprScroll = 0;   // #spvMain scroll offset, kept across re-renders
  function sprSecKey(num) { return num == null ? 'u' : String(num); }
  function sprSecNum(key) { return key === 'u' ? null : Number(key); }
  function sprHasOwn(st) { return st.startDay != null && st.durDays != null; }
  // the sprint an item or story starts in / ends in (numbers, timeline-clamped)
  function sprFirstNum(x) {
    var meta = state.meta;
    return RM.sprintNumForWeek(meta, Math.max(0, Math.min(meta.numWeeks - 1, Math.floor(x.startDay / SPW()))));
  }
  function sprLastNum(x) {
    var meta = state.meta, end = RM.itemEnd(x);
    if (end == null) return sprFirstNum(x);
    return RM.sprintNumForWeek(meta, Math.max(0, Math.min(meta.numWeeks - 1, Math.ceil(end / SPW()) - 1)));
  }
  // stories of one feature that belong to sprint `num` (null = unscheduled):
  // a story's own timeline decides when it has one, else it follows the feature
  function sprStoriesOf(it, num, featIn) {
    return (it.stories || []).filter(function (st) {
      if (sprHasOwn(st)) return num != null && sprFirstNum(st) === num;
      return num == null ? it.startDay == null : featIn;
    });
  }
  // row-level match: the text filter plus the phase / epic / workstream dropdowns
  function sprMatches(it) {
    if (!matchesFilter(it)) return false;
    if (sprFPhase != null && it.phaseId !== sprFPhase) return false;
    if (sprFEpic != null && (it.epic || '') !== sprFEpic) return false;
    if (sprFWs != null && (it.workstream || '') !== sprFWs) return false;
    return true;
  }
  function sprFilterOn() { return !!filterText || sprFPhase != null || sprFEpic != null || sprFWs != null; }
  // sprint totals read in story points when story sizing is on and its scale is
  // numeric (Fibonacci / points 1-5); T-shirt sizes have no points to add up
  function sprPointsOn() {
    if (!RM.sizingEnabled(state, 'story')) return false;
    var order = RM.sizeOrderOf(state, 'story');
    return order.length > 0 && order.every(function (v) {
      return v !== '' && v != null && !isNaN(Number(v));
    });
  }
  // half points make totals fractional; keep them off floating-point noise
  function fmtPts(p) { return Math.round(p * 10) / 10; }
  function sprStoryPts(st) {
    var n = st && st.size ? Number(st.size) : NaN;
    return isNaN(n) ? 0 : n;
  }
  function sprStorySized(st) {
    return !!(st && st.size) && !isNaN(Number(st.size));
  }
  // a points total only means something once a story is sized: until then the
  // header keeps the plain row count rather than reading a hollow "0 pt"
  function sprPointsTotal(ptFeats) {
    var sum = 0, any = false;
    ptFeats.forEach(function (f) {
      f.stories.forEach(function (st) {
        if (sprStorySized(st)) { any = true; sum += sprStoryPts(st); }
      });
    });
    return any ? sum : null;
  }
  // story-points capacity: each sprint's total reads planned / available
  // (needs numeric story sizes to add up and sprints to budget per; else the
  // plain count / points total as before)
  function sprCapXY() {
    return !!state.meta.capacityEnabled && state.meta.capMode === 'points' &&
      RM.sprintsEnabled(state.meta) && sprPointsOn();
  }
  // the points a sprint supplies across every capacity type (the capacity
  // summary's sprint period); null when the sprint is off the timeline
  function sprSupply(num) {
    var cap = validation && validation.capacity;
    if (!cap || !cap.periods) return null;
    var y = null;
    cap.periods.forEach(function (p, i) { if (p.num === num) y = cap.weeks[i].supply; });
    return y;
  }
  function sprSections() {
    var meta = state.meta;
    var pointsOn = sprPointsOn();
    var capXY = sprCapXY();
    var nums = sprintNums().concat([null]); // Unscheduled trails the timeline
    var secs = nums.map(function (n) {
      var items = n == null
        ? state.items.filter(function (it) { return it.startDay == null && sprMatches(it); })
        : itemsInSprint(n).filter(sprMatches);
      var inIds = {};
      items.forEach(function (it) { inIds[it.id] = true; });
      var feats = [];
      // the sprint's stories: what story level lists, and what points totals add
      var ptFeats = sprLevel === 'story' || pointsOn || capXY ? [] : null;
      if (ptFeats) {
        state.items.forEach(function (it) {
          if (!sprMatches(it)) return;
          var sts = sprStoriesOf(it, n, !!inIds[it.id]);
          if (sts.length) ptFeats.push({ it: it, stories: sts });
        });
        if (sprLevel === 'story') feats = ptFeats;
      }
      var w0 = n == null ? 0 : Math.max(0, RM.sprintRange(meta, n).w0);
      var d0 = n == null ? 0 : RM.sprintStartDay(meta, n);
      var wps = RM.sprintInfo(meta).wps;
      var points = pointsOn ? sprPointsTotal(ptFeats) : null;
      // points capacity: X is 0 until a story is sized, never the row count
      if (capXY) points = sprPointsTotal(ptFeats) || 0;
      return {
        num: n, key: sprSecKey(n), points: points,
        supply: capXY && n != null ? sprSupply(n) : null,
        title: n == null ? 'Unscheduled' : (RM.sprintsEnabled(meta) ? 'Sprint ' + n : 'Week of ' + RM.fmtShort(RM.weekStartDate(meta, w0))),
        dates: n == null ? '' : RM.fmtShort(RM.weekStartDate(meta, w0)) + ' – ' +
          RM.fmtShort(RM.spanEndDate(meta, d0, Math.min(wps * SPW(), meta.numWeeks * SPW() - d0))),
        items: items, feats: feats,
        count: sprLevel === 'story'
          ? feats.reduce(function (a, f) { return a + f.stories.length; }, 0) : items.length
      };
    });
    var cur = RM.sprintsEnabled(meta) ? currentSprintNum() : null;
    // count is level-aware (stories in story mode), so a sprint whose features
    // have no visible stories trims away instead of rendering empty
    var keep = function (s) { return s.num == null || s.num === cur || s.count > 0; };
    var first = -1, last = -1;
    secs.forEach(function (s, i) { if (s.num != null && keep(s)) { if (first === -1) first = i; last = i; } });
    var un = secs[secs.length - 1];
    return (first === -1 ? [] : secs.slice(first, last + 1)).concat([un]);
  }
  // rows whose span runs past their sprint get an info glyph saying how far
  function sprFlagHtml(x) {
    return x.flag ? '<span class="spv-info spv-flag" title="' + esc(flagTitle(x)) + '"><i data-lucide="flag"></i></span>' : '';
  }
  function sprCarryHtml(x) {
    if (!isScheduled(x)) return '';
    var s0 = sprFirstNum(x), s1 = sprLastNum(x);
    if (s1 <= s0) return '';
    var n = s1 - s0, unit = RM.sprintsEnabled(state.meta) ? 'sprint' : 'week';
    return '<span class="spv-info" title="Expecting to carryover for ' + n + ' ' + unit + (n === 1 ? '' : 's') + ' (through ' + unit + ' ' + s1 + ')"><i data-lucide="info"></i></span>';
  }
  // the est group's assignee chip (feature and story rows alike): a stack of
  // avatars, or an add-person glyph when nobody's on it yet
  function sprAsgChip(obj, act) {
    return '<span class="r-asg" tabindex="0" role="button" data-spact="' + act + '" title="Assignees">' +
      (avatarStack(obj.assignees, 2) || '<i data-lucide="user-plus"></i>') + '</span>';
  }
  function sprWsChip(it) {
    if (!state.meta.workstreamsEnabled) return '';
    return '<span class="spv-chip" tabindex="0" role="button" data-spact="ws" title="Workstream">' +
      '<span class="dd-dot" style="background:#' + RM.colorForWs(state, it.workstream) + '"></span>' + esc(it.workstream || RM.defaultWsName(state)) + '</span>';
  }
  // a row's title is plain text (double-click or Rename… edits it in place);
  // it takes only the width it needs so the chips sit close by
  function sprTitleHtml(txt, f) {
    var s = txt || '';
    return '<span class="spv-title' + (s ? '' : ' spv-ph') + '" data-spf="' + f + '" title="Double-click to rename">' +
      esc(s || (f === 'story' ? 'Story' : 'Name')) + '</span>';
  }
  function sprRowHtml(it, num) {
    return '<div class="spv-row' + (isSel(it.id) ? ' sel' : '') + (it.done ? ' done' : '') +
      '" data-spid="' + it.id + '" data-spsec="' + sprSecKey(num) + '">' +
      '<span class="spv-grip" title="Drag to another sprint or position"><i data-lucide="grip-vertical"></i></span>' +
      typeGlyphHtml(it, 'feature') +
      sprTitleHtml(it.feature, 'feature') +
      sprCarryHtml(it) + sprFlagHtml(it) + '<span class="spv-fill"></span>' +
      '<span class="spv-chip" tabindex="0" role="button" data-spact="epic" title="Epic">' +
      '<i data-lucide="' + (RM.iconForEpic(state, it.epic) || 'tag') + '"></i>' + esc(it.epic || '—') + '</span>' +
      sprWsChip(it) +
      '<span class="spv-est">' + ['size', 'pri', 'risk'].map(function (k) { return itemChipHtml(k, it, 'data-spact'); }).join('') + sprAsgChip(it, 'asg') + '</span>' +
      '</div>';
  }
  function sprStoryRowHtml(it, st, num) {
    var own = sprHasOwn(st);
    return '<div class="spv-row spv-st' + (isSel(it.id) && selStory === st.id ? ' sel' : '') + (st.done ? ' done' : '') +
      '" data-spid="' + it.id + '" data-spst="' + st.id + '" data-spsec="' + sprSecKey(num) + '">' +
      '<span class="spv-grip" title="Drag to another sprint or position"><i data-lucide="grip-vertical"></i></span>' +
      typeGlyphHtml(st, 'story', it) +
      sprTitleHtml(st.title, 'story') +
      (own ? sprCarryHtml(st) : '') + sprFlagHtml(st) + '<span class="spv-fill"></span>' +
      '<span class="spv-est">' + ['size', 'pri', 'risk'].map(function (k) { return storyChipHtml(k, st, 'data-spact'); }).join('') + sprAsgChip(st, 'st-asg') + '</span>' +
      '</div>';
  }
  // a sprint's total at the right edge: planned / available points under
  // points capacity, else the points total or the plain row count
  function sprCtHtml(sec) {
    var cls = 'pr-lanect spv-ct', txt, title;
    if (sec.supply != null) {
      // a filter leaves out some of the sprint's stories: X is then partial,
      // so it never reads red
      var filtered = sprFilterOn();
      var over = !filtered && sec.points > sec.supply + 1e-9;
      if (over) cls += ' over';
      txt = fmtPts(sec.points) + ' / ' + fmtPts(sec.supply);
      title = fmtPts(sec.points) + ' story points planned' + (filtered ? ' (filtered)' : '') + ' · ' +
        fmtPts(sec.supply) + ' available this sprint' + (over ? ' — over' : '');
    } else if (sec.points != null) {
      txt = fmtPts(sec.points) + ' pt'; title = 'Story points';
    } else {
      txt = String(sec.count); title = null;
    }
    return '<span class="' + cls + '"' + (title ? ' title="' + esc(title) + '"' : '') + '>' + esc(txt) + '</span>';
  }
  function sprSectionHtml(sec) {
    var body;
    if (sprLevel === 'story') {
      body = sec.feats.map(function (f) {
        return '<div class="spv-feat" data-spfeat="' + f.it.id + '" data-spid="' + f.it.id + '">' +
          typeGlyphHtml(f.it, 'feature') +
          '<span class="spv-featname">' + esc(f.it.feature || '(untitled)') + '</span>' +
          '<span class="spv-est">' + ['pri', 'size', 'dur'].map(function (k) { return itemChipHtml(k, f.it, 'data-spact'); }).join('') + '</span></div>' +
          f.stories.map(function (st) { return sprStoryRowHtml(f.it, st, sec.num); }).join('');
      }).join('');
    } else {
      body = sec.items.map(function (it) { return sprRowHtml(it, sec.num); }).join('');
    }
    return '<section class="spv-sec" data-spsec="' + sec.key + '">' +
      '<div class="spv-sechd"><h3>' + esc(sec.title) + '</h3>' +
      (sec.dates ? '<span class="spv-secdates">' + esc(sec.dates) + '</span>' : '') +
      sprCtHtml(sec) + '</div>' +
      '<div class="spv-rows">' + (body || '<div class="spv-empty">Nothing here' + (sprFilterOn() ? ' matches' : '') + '. Drop a row to move it into this sprint.</div>') + '</div>' +
      (sprLevel === 'feature' && !sec.items.length ? '<button class="spv-add" data-spadd="' + sec.key + '"><i data-lucide="plus"></i>' + esc('Add ' + lvl('feature')) + '</button>' : '') +
      '</section>';
  }
  function renderSprintPage() {
    var host = $('#sprintView');
    if (!host) return;
    if (sprFPhase != null && !state.phases.some(function (p) { return p.id === sprFPhase; })) sprFPhase = null;
    if (sprFWs != null && !state.meta.workstreamsEnabled) sprFWs = null;
    var secs = sprSections();
    // filter buttons wear what they picked (phase name, epic icon, workstream dot) and a blue outline
    var fPh = sprFPhase == null ? null : state.phases.filter(function (p) { return p.id === sprFPhase; })[0];
    var phaseFilterDd = state.phases.length > 1
      ? '<button class="dd-btn' + (fPh ? ' pr-on' : '') + '" data-spdd="fphase" title="Filter by phase">' +
        '<span class="dd-label"><i data-lucide="milestone"></i>' + (fPh ? esc(fPh.name) : 'All phases') + '</span>' +
        '<i data-lucide="chevron-down"></i></button>' : '';
    var epicFilterDd = allEpics().length
      ? '<button class="dd-btn' + (sprFEpic == null ? '' : ' pr-on') + '" data-spdd="fepic" title="Filter by epic">' +
        '<span class="dd-label">' + (sprFEpic == null ? '<i data-lucide="tag"></i>All epics'
          : '<i data-lucide="' + (sprFEpic ? (RM.iconForEpic(state, sprFEpic) || 'tag') : 'tag') + '"></i>' + esc(sprFEpic || 'No epic')) + '</span>' +
        '<i data-lucide="chevron-down"></i></button>' : '';
    var wsFilterDd = state.meta.workstreamsEnabled
      ? '<button class="dd-btn' + (sprFWs == null ? '' : ' pr-on') + '" data-spdd="fws" title="Filter by workstream">' +
        '<span class="dd-label">' + (sprFWs == null ? '<i data-lucide="layers"></i>All workstreams'
          : '<span class="dd-dot" style="background:#' + (sprFWs ? RM.colorForWs(state, sprFWs) : RM.defaultWsColor(state)) + '"></span>' +
            esc(sprFWs || RM.defaultWsName(state))) + '</span>' +
        '<i data-lucide="chevron-down"></i></button>' : '';
    var cur = RM.sprintsEnabled(state.meta) ? currentSprintNum() : null;
    var side = secs.map(function (sec) {
      return '<button class="spv-sbtn' + (sec.num != null && sec.num === cur ? ' today' : '') +
        '" data-spside="' + sec.key + '" title="' + esc(sec.title + (sec.dates ? ' · ' + sec.dates : '')) + '">' +
        '<i data-lucide="' + (sec.num == null ? 'inbox' : 'calendar-range') + '"></i>' +
        '<span class="spv-sbtxt"><span class="spv-sbname">' + esc(sec.title) + '</span>' +
        (sec.dates ? '<small>' + esc(sec.dates) + '</small>' : '') + '</span>' +
        sprCtHtml(sec) + '</button>';
    }).join('');
    host.innerHTML = '<div class="spv">' +
      '<aside class="spv-side"><div class="spv-sidehd">Sprints</div>' + side + '</aside>' +
      '<div class="spv-main" id="spvMain"><div class="pr-bar">' +
      levelBtnHtml('data-spdd="level"', dmModes(), sprLevel) +
      '<span class="filter-wrap pr-filter"><i data-lucide="search" class="filter-ico" aria-hidden="true"></i>' +
      '<input id="spFilter" type="search" placeholder="Filter rows" aria-label="Filter rows" value="' + esc(filterText) + '">' +
      '<kbd class="kbd filter-kbd" aria-hidden="true">⌘F</kbd></span>' +
      phaseFilterDd + epicFilterDd + wsFilterDd +
      '</div>' +
      secs.map(sprSectionHtml).join('') + '</div></div>';
    if (window.lucide) lucide.createIcons();
    var main = $('#spvMain');
    main.scrollTop = sprScroll;
    sprSyncSide();
  }
  // the sidebar tracks the section under the top of the scroll area
  function sprSyncSide() {
    var main = $('#spvMain');
    if (!main) return;
    var secs = $$('#spvMain .spv-sec');
    var top = main.getBoundingClientRect().top + 60, on = secs[0];
    secs.forEach(function (s) { if (s.getBoundingClientRect().top <= top) on = s; });
    $$('#sprintView .spv-sbtn').forEach(function (b) {
      b.classList.toggle('on', !!on && b.dataset.spside === on.dataset.spsec);
    });
  }
  $('#sprintView').addEventListener('scroll', function (e) {
    if (e.target.id !== 'spvMain') return;
    sprScroll = e.target.scrollTop;
    sprSyncSide();
  }, true);
  // adds a feature scheduled at the section's sprint (or unscheduled)
  function sprAddFeature(key) {
    var num = sprSecNum(key);
    doAddFeature();
    var nid = selectedId;
    if (nid && num != null) {
      commit('schedule', function (s) { RM.moveItemToSprint(s, nid, num, null); });
    }
  }
  var spFilterTimer = null;
  $('#sprintView').addEventListener('input', function (e) {
    if (e.target.id !== 'spFilter') return;
    var v = e.target.value;
    clearTimeout(spFilterTimer);
    spFilterTimer = setTimeout(function () {
      filterText = v.trim();
      render();
      var nf = $('#spFilter');
      if (nf) { nf.focus(); nf.setSelectionRange(nf.value.length, nf.value.length); }
    }, 120);
  });
  // double-click a title (or Rename… in the row menu) to edit it in place
  function sprStartRename(row) {
    if (!row) return;
    var span = row.querySelector('.spv-title');
    if (!span || span.tagName === 'INPUT') return;
    var id = row.dataset.spid, stId = row.dataset.spst || null;
    var blank = span.classList.contains('spv-ph');
    startInlineEdit(span, function (val) {
      if (stId) {
        commit('rename story', function (s) {
          var st = storyById(RM.itemById(s, id) || {}, stId);
          if (st) st.title = val;
        });
      } else {
        commit('rename', function (s) {
          var x = RM.itemById(s, id);
          if (x) x.feature = val;
        });
      }
    });
    var inp = row.querySelector('input.st-add-input');
    if (inp) {
      inp.dataset.spf = stId ? 'story' : 'feature';
      inp.className = 'st-add-input spv-title';
      if (blank) { inp.value = ''; }
      inp.select();
    }
  }
  $('#sprintView').addEventListener('dblclick', function (e) {
    var t = e.target.closest('.spv-title');
    if (!t || t.tagName === 'INPUT') return;
    e.preventDefault();
    sprStartRename(t.closest('[data-spid]'));
  });
  $('#sprintView').addEventListener('keydown', function (e) {
    var t = e.target;
    if (t.id === 'spFilter' && e.key === 'Escape') {
      e.stopPropagation();
      t.value = '';
      filterText = '';
      render();
      return;
    }
  });
  $('#sprintView').addEventListener('click', function (e) {
    if (dragConsumedClick) { dragConsumedClick = false; return; }
    var sb = e.target.closest('[data-spside]');
    if (sb) {
      var sec = $('#spvMain .spv-sec[data-spsec="' + sb.dataset.spside + '"]');
      var main = $('#spvMain');
      if (sec && main) {
        // land with the sprint heading just under the sticky filter bar,
        // not hidden behind it
        var bar = main.querySelector('.pr-bar');
        main.scrollTop = sec.offsetTop - main.offsetTop - (bar ? bar.offsetHeight : 0);
        sprScroll = main.scrollTop;
        sprSyncSide();
      }
      return;
    }
    var lv = e.target.closest('[data-spdd="level"]');
    if (lv) {
      openLevelMenu(lv, dmModes(), sprLevel, function (mode) {
        if (sprLevel !== mode) { sprLevel = mode; saveLocal(); render(); }
      });
      return;
    }
    var fdd = e.target.closest('[data-spdd]');
    if (fdd) {
      var fk = fdd.dataset.spdd;
      if (fk === 'fphase') {
        openDropdown(fdd, [{ label: '<i>All phases</i>', checked: sprFPhase == null, fn: function () { sprFPhase = null; render(); } }]
          .concat(state.phases.map(function (p) {
            return { label: esc(p.name), checked: sprFPhase === p.id, fn: function () { sprFPhase = p.id; render(); } };
          })));
      } else if (fk === 'fepic') {
        openDropdown(fdd, [{ label: '<i>All epics</i>', checked: sprFEpic == null, fn: function () { sprFEpic = null; render(); } },
          { label: '<i>— no epic —</i>', checked: sprFEpic === '', fn: function () { sprFEpic = ''; render(); } }]
          .concat(allEpics().map(function (ep) {
            return { icon: RM.iconForEpic(state, ep) || 'tag', label: esc(ep), checked: sprFEpic === ep, fn: function () {
              sprFEpic = ep; render();
            } };
          })));
      } else if (fk === 'fws') {
        openDropdown(fdd, [{ label: '<i>All workstreams</i>', checked: sprFWs == null, fn: function () { sprFWs = null; render(); } },
          { label: esc(RM.defaultWsName(state)) + ' <i>(default)</i>', dot: '#' + RM.defaultWsColor(state),
            checked: sprFWs === '', fn: function () { sprFWs = ''; render(); } }]
          .concat(allWorkstreams().map(function (w) {
            return { label: esc(w), dot: '#' + RM.colorForWs(state, w), checked: sprFWs === w, fn: function () {
              sprFWs = w; render();
            } };
          })));
      }
      return;
    }
    var add = e.target.closest('[data-spadd]');
    if (add) { sprAddFeature(add.dataset.spadd); return; }
    var chip = e.target.closest('[data-spact]');
    var row = e.target.closest('[data-spid]');
    if (chip && row) {
      var cid = row.dataset.spid;
      var itC = RM.itemById(state, cid);
      if (!itC) return;
      if (row.dataset.spst) storyChipAction(chip.dataset.spact, chip, cid, row.dataset.spst);
      else if (chip.dataset.spact === 'epic') openDropdown(chip, setEpicMenu(cid, true));
      else if (chip.dataset.spact === 'ws') openDropdown(chip, wsMenuItems(cid, function () { return chip; }));
      else itemChipAction(chip.dataset.spact, chip, cid);
      return;
    }
    if (row && !e.target.closest('input,button')) {
      if (row.dataset.spst) {
        var wasSt = selStory;
        select(row.dataset.spid);
        if (wasSt !== row.dataset.spst) { selStory = row.dataset.spst; render(); }
      } else if (selectedId !== row.dataset.spid || selStory) select(row.dataset.spid);
    }
  });
  // the Rename… entry every titled row's menu opens with (re-queried on click:
  // the menu's own render may have replaced the row element)
  function sprRenameEntry(row) {
    if (!row.querySelector('.spv-title')) return null;
    var sel = row.dataset.spst
      ? '#sprintView .spv-row[data-spst="' + row.dataset.spst + '"]'
      : '#sprintView .spv-row[data-spid="' + row.dataset.spid + '"]:not(.spv-st)';
    return { icon: 'pencil', label: 'Rename…', fn: function () { sprStartRename($(sel)); } };
  }
  $('#sprintView').addEventListener('contextmenu', function (e) {
    var row = e.target.closest('[data-spid]');
    if (!row || e.target.closest('input,textarea,select')) return;
    e.preventDefault();
    e.stopPropagation();
    var cid = row.dataset.spid;
    var itX = RM.itemById(state, cid);
    if (!itX) return;
    var cx = e.clientX, cy = e.clientY;
    if (row.dataset.spst) {
      var sid = row.dataset.spst;
      var stX = storyById(itX, sid);
      openContextMenu(cx, cy, [
        sprRenameEntry(row),
        { icon: 'calendar-range', label: 'Move to sprint…', fn: function () { openContextMenu(cx, cy, moveStorySprintMenu(cid, sid)); } },
        stX && sprHasOwn(stX) ? { icon: 'corner-down-right', label: 'With feature', fn: function () {
          commit('story with feature', function (s) { RM.moveStoryToSprint(s, cid, sid, null, sid); });
        } } : null,
        state.team.length ? { icon: 'users', label: 'Assign…', fn: function () { openContextMenu(cx, cy, storyAssignMenuItems(cid, sid), ASSIGN_DD); } } : null,
        moveStoryFeatureEntry(cx, cy, cid, sid),
        flagMenuEntries(cid, sid)[0], flagMenuEntries(cid, sid)[1] || null, noAutoEntry(cid, sid),
        { sep: true },
        storyInsertEntries(cid, sid)[0], storyInsertEntries(cid, sid)[1],
        { icon: 'copy', label: 'Duplicate story', fn: function () { duplicateStory(cid, sid); } }
      ].filter(Boolean));
      return;
    }
    openContextMenu(cx, cy, [
      sprRenameEntry(row),
      { icon: 'calendar-range', label: 'Move to sprint…', fn: function () { openContextMenu(cx, cy, moveSprintMenu(cid)); } },
      { icon: 'folder-input', label: 'Move to phase…', fn: function () { openContextMenu(cx, cy, movePhaseMenu(cid)); } },
      { icon: 'tag', label: 'Set epic…', fn: function () { openContextMenu(cx, cy, setEpicMenu(cid, false)); } },
      state.meta.workstreamsEnabled
        ? { icon: 'layers', label: 'Set workstream…', fn: function () { openContextMenu(cx, cy, wsMenuItems(cid, function () { return null; })); } }
        : null,
      state.team.length ? { icon: 'users', label: 'Assign…', fn: function () { openContextMenu(cx, cy, assignMenuItems(cid)); } } : null,
      flagMenuEntries(cid, null)[0], flagMenuEntries(cid, null)[1] || null, noAutoEntry(cid, null),
      isScheduled(itX) ? { icon: 'calendar-off', label: 'Unschedule', fn: function () {
        commit('unschedule', function (s) { RM.moveItemToSprint(s, cid, null, null); });
      } } : null,
      { sep: true },
      { icon: 'trash-2', label: 'Delete…', danger: true, fn: function () { deleteItemConfirm(cid); } }
    ].filter(Boolean));
  });
  // rows drag between sprints (sections or sidebar entries) and reorder
  // inside one; stories only reorder among their own feature's stories
  $('#sprintView').addEventListener('pointerdown', function (e) {
    if (e.button !== 0 || drag) return;
    if (e.target.closest('input,textarea,button,select,[data-spact],[contenteditable="true"]')) return;
    var row = e.target.closest('.spv-row');
    if (!row) return;
    drag = { kind: 'sprow', id: row.dataset.spid, stId: row.dataset.spst || null, fromSec: row.dataset.spsec,
      el: row, x0: e.clientX, y0: e.clientY, moved: false, ghost: null, target: null };
    e.preventDefault();
  });
  function sprClearDrop() {
    $$('#sprintView .drop-before,#sprintView .drop-after,#sprintView .drop').forEach(function (n) {
      n.classList.remove('drop-before', 'drop-after', 'drop');
    });
  }
  function sprDragMove(e) {
    if (!drag.ghost) {
      drag.ghost = document.createElement('div');
      drag.ghost.className = 'sp-card sp-card-ghost';
      var t = RM.itemById(state, drag.id);
      var st = t && drag.stId ? storyById(t, drag.stId) : null;
      drag.ghost.textContent = st ? (st.title || 'Story') : ((t && t.feature) || 'Feature');
      document.body.appendChild(drag.ghost);
      drag.el.classList.add('sp-dragging');
    }
    drag.ghost.style.left = (e.clientX + 10) + 'px';
    drag.ghost.style.top = (e.clientY + 8) + 'px';
    var under = (document.elementFromPoint && document.elementFromPoint(e.clientX, e.clientY)) || e.target;
    sprClearDrop();
    drag.target = null;
    if (!under || !under.closest) return;
    var sb = under.closest('[data-spside]');
    if (sb) { sb.classList.add('drop'); drag.target = { sec: sb.dataset.spside, beforeId: null, beforeStId: null }; return; }
    var row = under.closest('.spv-row');
    var sec = under.closest('.spv-sec');
    if (!sec) return;
    var tgt = { sec: sec.dataset.spsec, beforeId: null, beforeStId: null };
    if (row && row !== drag.el) {
      var r = row.getBoundingClientRect();
      var before = e.clientY < r.top + r.height / 2;
      var ref = before ? row : row.nextElementSibling;
      while (ref && !ref.classList.contains('spv-row')) ref = ref.nextElementSibling; // skip feature headings
      if (ref && ref !== drag.el) {
        if (drag.stId) {
          // a story only has a "before" among its own feature's stories
          if (ref.dataset.spid === drag.id) tgt.beforeStId = ref.dataset.spst;
        } else tgt.beforeId = ref.dataset.spid;
      }
      row.classList.add(before ? 'drop-before' : 'drop-after');
    } else sec.classList.add('drop');
    drag.target = tgt;
  }
  function sprDragEnd(d) {
    if (d.ghost) d.ghost.remove();
    d.el.classList.remove('sp-dragging');
    sprClearDrop();
    var t = d.target;
    if (!t) { render(); return; }
    var num = sprSecNum(t.sec), moved = t.sec !== d.fromSec;
    if (d.stId) {
      if (!moved && !t.beforeStId) { render(); return; }
      commit(moved ? 'move story to sprint' : 'reorder story', function (s) {
        RM.moveStoryToSprint(s, d.id, d.stId, moved ? num : sprHasOwn(storyById(RM.itemById(s, d.id) || {}, d.stId)) ? num : null, t.beforeStId);
      });
      return;
    }
    if (!moved && !t.beforeId) { render(); return; }
    commit(moved ? 'move to sprint' : 'reorder', function (s) {
      if (moved) RM.moveItemToSprint(s, d.id, num, t.beforeId);
      else RM.reorderItem(s, d.id, t.beforeId);
      if (moved && autoOrder) RM.sortItemsByStart(s);
    });
  }

  function renderScopeHeader() {
    var out = ['<div class="sc-hrow">'];
    allScopeCols().forEach(function (c) {
      var fixed = isFixedColKey(c[0]);
      out.push('<div class="sc-hcell' + (fixed ? ' sc-fixh' : '') + '" data-col="' + c[0] + '" style="width:' + scopeColWidth(c) + 'px">' +
        '<span class="sc-hlab">' + esc(c[1]) + '</span>' +
        '<span class="sc-rz" data-rz="' + c[0] + '"></span></div>');
    });
    out.push('<div class="sc-hcell sc-hadd"><button class="sc-hbtn" data-coladd title="Add a column">' +
      '<i data-lucide="plus"></i></button></div>');
    out.push('</div>');
    $('#hdrSprints').innerHTML = out.join('');
    $('#hdrSprints').style.width = scopeW() + 'px';
  }

  function renderTopbar() {
    // the desktop shell mirrors File/Edit/View into the macOS menu bar;
    // nudge it so checkmarks and enabled states track the app state
    if (window.HeadwayDesktop && HeadwayDesktop.syncMenu) HeadwayDesktop.syncMenu();
    var t = $('#docTitle');
    if (t.value !== state.meta.title && document.activeElement !== t) t.value = state.meta.title;
    sizeTitle();
    updateSaveBtn();
    var counts = validation.counts;
    var n = counts.error + counts.warn;
    $('#valCount').textContent = n || '✓';
    var dot = $('#valDot');
    dot.className = counts.error ? 'err' : (counts.warn ? 'warn' : '');
    dot.id = 'valDot';
    $$('#viewTabs button, #btnSetup, #btnHistory').forEach(function (b) {
      b.classList.toggle('on', b.dataset.view === view);
    });

    var stats = RM.scheduleStats(state);
    var fs = '';
    if (stats.lastDay != null) {
      fs = 'ships ' + RM.fmtShortYear(RM.dayToDate(state.meta, stats.lastDay - 1)) +
        ' · ' + stats.scheduled + ' scheduled · ' + stats.unscheduled + ' backlog';
    }
    var fsEl = $('#finishStat');
    if (fsEl) fsEl.textContent = fs;
    if (window.HeadwayJira) HeadwayJira.renderStatus($('#btnJira'), state);
  }

  // the widest capacity-cell form that fits the week column: ~5.5px per
  // character in the 9px mono those cells use, with 2px of air each side
  function capCellText(forms, cellPx) {
    var room = (cellPx != null ? cellPx : weekPx) - 4;
    for (var i = 0; i < forms.length; i++) {
      // fmtPe hands back a number, so the bare-ask form needs coercing
      var t = String(forms[i]);
      if (t.length * 5.5 <= room) return t;
    }
    return '';
  }

  function renderHeader(laneW) {
    var meta = state.meta;
    var si = RM.sprintInfo(meta);
    var wps = si.wps;
    var sprintW = wps * weekPx;
    var hset = RM.holidayDaySet(meta);
    var hs = [], hw = [];
    // sprint boundaries align to the numbering anchor; the first boundary is
    // the one at or before week 0
    var firstB = si.anchorWeek - Math.ceil(si.anchorWeek / wps) * wps;
    for (var bw = firstB; bw < meta.numWeeks; bw += wps) {
      var w0v = Math.max(0, bw);
      var w1v = Math.min(meta.numWeeks, bw + wps);
      if (w1v <= w0v) continue;
      var num = RM.sprintNumForWeek(meta, bw);
      var d = RM.weekStartDate(meta, bw);
      var cellW = (w1v - w0v) * weekPx;
      // date is the primary label and is always shown; the sprint number is
      // secondary and only appears when there's room
      var dateTxt = cellW >= 64 ? RM.fmtShort(d) : (d.getUTCMonth() + 1) + '/' + d.getUTCDate();
      var numTag = cellW >= 56 && RM.sprintsEnabled(meta) ? '<span class="sp-num">S' + num + '</span>' : '';
      hs.push('<div class="sprint-cell" style="left:' + (w0v * weekPx) + 'px;width:' + cellW +
        'px" title="Sprint ' + num + '">' +
        '<span class="sp-date">' + dateTxt + '</span>' + numTag + '</div>');
    }

    // the capacity row: ONE row summing every capacity type — each period's
    // total demand vs total supply (a week each in people, a sprint each in
    // story points; sprint cells span the sprint cells above). A sum can hide
    // one type overflowing, so a cell reads over when ANY type is over
    // (overAny, incl. work of a type nobody supplies); the tooltip lists
    // each type's numbers.
    var cap = validation.capacity;
    var unitWord = meta.capMode === 'points' ? 'points' : 'people';
    var bySprint = cap.period === 'sprint';
    var capRowsHtml = [];
    var hcc = [];
    cap.periods.forEach(function (per, pi) {
      var w = per.w0, nW = per.w1 - per.w0, cellPx = nW * weekPx;
      var cell = cap.weeks[pi];
      var hn = 0;
      for (var hw2 = per.w0; hw2 < per.w1; hw2++) hn += RM.holidaysInWeek(meta, hw2, hset);
      var cls, txt2 = '', title;
      if (cell.blackout) { cls = 'blackout'; txt2 = weekPx >= 24 ? '✕' : ''; title = bySprint ? 'Holiday weeks' : 'Holiday week'; }
      else {
        var ratio = cell.supply > 0 ? cell.demand / cell.supply : (cell.demand > 0 ? Infinity : 0);
        cls = cell.overAny ? 'over' : (cell.demand === 0 ? 'idle' : (ratio > 0.85 ? 'mid' : 'ok'));
        // the widest form that actually fits THIS cell's numbers: spaced
        // demand / supply, then tight, then the ask alone, then nothing.
        // Measured off the string, not the zoom — a clipped “4 / 0.” reads
        // as a different number. The tooltip always carries both.
        var dTxt = fmtPe(cell.demand), sTxt = fmtPe(cell.supply);
        txt2 = capCellText([dTxt + ' / ' + sTxt, dTxt + '/' + sTxt, dTxt], cellPx);
        // per type: every type with something asked or supplied this period
        var parts = [];
        cap.types.forEach(function (ct) {
          var tc = cap.rows[ct][pi];
          if (tc.demand <= 1e-9 && tc.supply <= 1e-9) return;
          parts.push(ct + ' ' + fmtPe(tc.demand) + ' / ' + fmtPe(tc.supply) +
            (tc.over ? (tc.supply <= 1e-9 ? ' — no supply, nobody on the roster supplies ' + ct : ' — over') : ''));
        });
        title = fmtPe(cell.demand) + ' ' + unitWord + ' asked · ' + fmtPe(cell.supply) + ' available' +
          (parts.length ? ' · ' + parts.join(' · ') : '');
      }
      var when = !bySprint ? 'Week of ' + RM.fmtShort(RM.weekStartDate(meta, w))
        : (per.num != null ? 'Sprint ' + per.num + ', ' : '') + RM.fmtShort(RM.weekStartDate(meta, w)) + ' – ' +
          RM.fmtShort(RM.spanEndDate(meta, w * SPW(), nW * SPW()));
      // a sprint cell spans its weeks; a click toggles the week under the
      // pointer (data-w1 marks the span)
      hcc.push('<div class="cap-cell ' + cls + (hn && !cell.blackout ? ' part' : '') + '" tabindex="0" data-w="' + w +
        (bySprint ? '" data-w1="' + per.w1 : '') +
        '" style="left:' + (w * weekPx + 1) + 'px;width:' + (cellPx - 2) + 'px" title="' +
        esc(when + ': ' + title +
          (hn ? ' · ' + hn + ' holiday day(s)' : '') + ' · click to toggle holiday week') + '">' + txt2 + '</div>');
    });
    capRowsHtml.push('<div class="hdr-line hdr-cap">' +
      '<div class="hdr-left"><span class="cap-row-lab" title="' +
      esc((bySprint ? 'Each sprint: ' : 'Each week: ') + unitWord + ' asked / available across every capacity type — Setup → Scheduling') + '">' +
      'Capacity (' + unitWord + ')</span></div>' +
      '<div class="hdr-lane">' + hcc.join('') + '</div></div>');
    // phase lane above the dates: user-pinned dates win, otherwise the span
    // auto-derives from the phase's scheduled items; overlapping phases stack
    var PH_H = 20;
    var phCells = [], phLanes = [], phSpans = [];
    state.phases.forEach(function (p) {
      var span = RM.phaseSpan(state, p);
      if (span) phSpans.push({ p: p, lo: span.lo, hi: span.hi });
    });
    phSpans.sort(function (a, b) { return a.lo - b.lo || a.hi - b.hi; });
    phSpans.forEach(function (sp) {
      var px = sp.lo * dayPx(), pw = (sp.hi - sp.lo) * dayPx();
      var lvl = 0;
      while (lvl < phLanes.length && phLanes[lvl] > px + 0.5) lvl++;
      phLanes[lvl] = px + pw;
      var pinned = sp.p.startDay != null || sp.p.endDay != null;
      phCells.push('<div class="ph-cell' + (pinned ? ' pinned' : '') + '" data-phase="' + sp.p.id + '" style="left:' + px + 'px;width:' + pw +
        'px;top:' + (lvl * PH_H + 2) + 'px" data-range="' +
        esc(RM.fmtShort(RM.dayToDate(meta, sp.lo)) + ' → ' +
          RM.fmtShort(RM.dayToDate(meta, Math.max(sp.lo, sp.hi - 1))) + (pinned ? ' · pinned' : ' · auto')) + '">' +
        '<span class="ph-h l" data-phh="l"></span>' +
        '<span class="ph-name">' + esc(sp.p.name) + '</span>' +
        '<span class="ph-h r" data-phh="r"></span></div>');
    });
    var phLine = $('#hdrPhases').parentNode;
    // the line stays visible even with no phase spans, keeping the header
    // heights stable
    phLine.style.height = Math.max(34, phLanes.length ? phLanes.length * PH_H + 4 : 0) + 'px';
    phLine.style.display = '';
    $('#hdrPhases').innerHTML = phCells.join('');
    $('#hdrPhases').style.width = laneW + 'px';

    $('#hdrSprints').innerHTML = hs.join('');
    $('#hdrSprints').style.width = laneW + 'px';
    $('#hdrCapRows').innerHTML = capRowsHtml.join('');
    if (window.lucide) lucide.createIcons();
  }

  // click a header phase span to edit that phase
  $('#hdrPhases').addEventListener('click', function (e) {
    if (dragConsumedClick) { dragConsumedClick = false; return; }
    var c = e.target.closest('[data-phase]');
    if (c) phaseModal(c.dataset.phase);
  });

  // drag a phase span to PIN its dates: body moves the window, the edge
  // handles resize it (each pins just that side; unset sides stay auto)
  $('#hdrPhases').addEventListener('pointerdown', function (e) {
    if (e.button !== 0 || drag) return;
    var c = e.target.closest('.ph-cell');
    if (!c) return;
    var p = null;
    state.phases.forEach(function (x) { if (x.id === c.dataset.phase) p = x; });
    var span = p && RM.phaseSpan(state, p);
    if (!span) return;
    var hh = e.target.closest('[data-phh]');
    drag = {
      kind: 'phspan', phaseId: p.id, el: c,
      mode: hh ? (hh.dataset.phh === 'l' ? 'resize-l' : 'resize-r') : 'move',
      x0: e.clientX, y0: e.clientY, lo0: span.lo, hi0: span.hi, moved: false
    };
    e.preventDefault();
  });

  function phSpanDragMove(e, dx) {
    var dd = daysFromDx(dx);
    var sn = e.altKey ? function (d) { return d; } : function (d) { return snapTo(d, 'feature'); };
    var lo = drag.lo0, hi = drag.hi0;
    if (drag.mode === 'move') {
      lo = Math.max(0, sn(drag.lo0 + dd));
      hi = lo + (drag.hi0 - drag.lo0);
    } else if (drag.mode === 'resize-l') {
      lo = Math.max(0, Math.min(sn(drag.lo0 + dd), hi - 1));
    } else {
      hi = Math.max(lo + 1, sn(drag.hi0 + dd));
    }
    document.body.classList.add('dragging-x');
    drag.el.style.left = (lo * dayPx()) + 'px';
    drag.el.style.width = ((hi - lo) * dayPx()) + 'px';
    drag.lo = lo; drag.hi = hi;
    dragTip.hidden = false;
    dragTip.style.left = (e.clientX + 14) + 'px';
    dragTip.style.top = (e.clientY - 34) + 'px';
    dragTip.innerHTML = '<b>' + RM.fmtShort(RM.dayToDate(state.meta, lo)) + '</b> → ' +
      RM.fmtShort(RM.dayToDate(state.meta, Math.max(lo, hi - 1)));
  }

  function phSpanDragEnd(d) {
    if (d.lo == null) return;
    commit('phase dates', function (s) {
      s.phases.forEach(function (p) {
        if (p.id !== d.phaseId) return;
        if (d.mode !== 'resize-r') p.startDay = d.lo;
        if (d.mode !== 'resize-l') p.endDay = d.hi;
      });
    });
  }

  // hovering a phase span shows a rich tooltip (description + date range)
  // instead of a native title
  var phTip = document.createElement('div');
  phTip.id = 'phTip';
  phTip.hidden = true;
  document.body.appendChild(phTip);
  $('#hdrPhases').addEventListener('mouseover', function (e) {
    var c = e.target.closest('.ph-cell');
    if (!c) return;
    var p = null;
    state.phases.forEach(function (x) { if (x.id === c.dataset.phase) p = x; });
    if (!p) return;
    phTip.innerHTML = '<b>' + esc(p.name) + '</b>' +
      (p.description ? '<div class="pht-desc">' + esc(p.description) + '</div>' : '') +
      '<div class="pht-range">' + esc(c.dataset.range || '') + '</div>';
    phTip.hidden = false;
    var r = c.getBoundingClientRect();
    var tw = phTip.offsetWidth;
    phTip.style.left = Math.max(8, Math.min(e.clientX - tw / 2, window.innerWidth - tw - 12)) + 'px';
    phTip.style.top = (r.bottom + 6) + 'px';
  });
  $('#hdrPhases').addEventListener('mouseout', function (e) {
    if (!e.relatedTarget || !e.relatedTarget.closest('.ph-cell')) phTip.hidden = true;
  });

  // click a capacity cell to toggle that week as a holiday week
  $('#hdrCapRows').addEventListener('click', function (e) {
    if (dragConsumedClick) { dragConsumedClick = false; return; }
    var cell = e.target.closest('[data-w]');
    if (!cell) return;
    var w = parseInt(cell.dataset.w, 10);
    // a sprint-wide cell (points mode): the week under the pointer; a
    // keyboard click (no pointer position) takes the sprint's first week
    if (cell.dataset.w1 && e.clientX) {
      var r = cell.getBoundingClientRect();
      var off = Math.floor((e.clientX - r.left) / weekPx);
      if (off > 0) w = Math.min(parseInt(cell.dataset.w1, 10) - 1, w + off);
    }
    var iso = RM.fmtISO(RM.weekStartDate(state.meta, w));
    var isFull = RM.holidaysInWeek(state.meta, w) === SPW();
    commit('toggle holiday', function (s) {
      var mon = RM.fmtISO(RM.dayToDate(s.meta, w * RM.slotsOf(s.meta)));
      var fri = RM.fmtISO(RM.dayToDate(s.meta, (w + 1) * RM.slotsOf(s.meta) - 1));
      if (isFull) {
        // carve this week out of every holiday range
        RM.clipHolidayRanges(s.meta, mon, fri);
      } else {
        RM.addHolidayRange(s.meta, '', mon, fri);
      }
    });
    toast(isFull ? 'Week of ' + iso + ' is a working week again' : 'Week of ' + iso + ' marked as a holiday week');
  });

  function renderBgCols(laneW) {
    var meta = state.meta;
    var si = RM.sprintInfo(meta);
    var html = [];
    // simplified grid: only sprint boundaries (anchor-aligned) get a line
    var firstB = si.anchorWeek - Math.ceil(si.anchorWeek / si.wps) * si.wps;
    // sticky group bands are opaque, so they redraw this grid themselves:
    // publish the pitch and phase the lines use (lane-local pixels)
    var rs = document.documentElement.style;
    rs.setProperty('--sprint-px', (si.wps * weekPx) + 'px');
    rs.setProperty('--sprint-off', (firstB * weekPx) + 'px');
    for (var w = firstB; w <= meta.numWeeks; w += si.wps) {
      if (w < 0) continue;
      html.push('<div class="bg-week sprint" style="left:calc(var(--left-w) + ' + (w * weekPx) + 'px)"></div>');
    }
    // holiday segments are drawn per DAY, merged into contiguous runs
    var hset = RM.holidayDaySet(meta);
    var numDays = RM.numDays(meta);
    for (var d0 = 0; d0 < numDays; d0++) {
      if (!hset[d0]) continue;
      var d1 = d0;
      while (d1 + 1 < numDays && hset[d1 + 1]) d1 += 1;
      html.push('<div class="bg-blackout" title="holiday" style="left:calc(var(--left-w) + ' + (d0 * dayPx()) +
        'px);width:' + ((d1 - d0 + 1) * dayPx()) + 'px"></div>');
      d0 = d1;
    }
    $('#bgcols').innerHTML = html.join('');
  }

  function positionToday() {
    var now = new Date();
    var today = new Date(Date.UTC(now.getFullYear(), now.getMonth(), now.getDate()));
    var d = RM.dateToDay(state.meta, today);
    var line = $('#todayLine');
    if (d < 0 || d > RM.numDays(state.meta)) { line.hidden = true; return; }
    line.hidden = false;
    line.dataset.x = d * dayPx();
    line.style.left = 'calc(var(--left-w) + ' + (d * dayPx()) + 'px)';
    syncTodayClip();
  }
  // the line paints above the sticky header, so it must duck when scrolled
  // behind the frozen left pane
  function syncTodayClip() {
    var line = $('#todayLine');
    if (!line || line.hidden) return;
    line.classList.toggle('under-left', (+line.dataset.x || 0) - board.scrollLeft < 0);
  }

  // the orange flag: shown in the alert slot, its reason on hover
  function flagTitle(obj) { return 'Flagged' + (obj.flag && obj.flag.reason ? ': ' + obj.flag.reason : ''); }
  function flagBadgeHtml(obj) {
    return '<span class="r-warn flag" data-act="warn" title="' + esc(flagTitle(obj)) + '"><i data-lucide="flag"></i></span>';
  }
  function warnBadge(it) {
    if (it.flag) return flagBadgeHtml(it); // the flag wins the slot; validation stays in the hover
    var list = validation.byItem[it.id] || [];
    if (!list.length) return '<span class="r-warn"></span>';
    var top = 'info';
    list.forEach(function (v) {
      if (v.level === 'error') top = 'err';
      else if (v.level === 'warn' && top !== 'err') top = 'warn';
    });
    var icon = top === 'err' ? 'octagon-alert' : top === 'warn' ? 'triangle-alert' : 'info';
    return '<span class="r-warn ' + top + '" data-act="warn"><i data-lucide="' + icon + '"></i></span>';
  }

  // hovering an alert icon shows its checks immediately (no title delay)
  var hoverTip = document.createElement('div');
  hoverTip.id = 'hoverTip';
  hoverTip.hidden = true;
  document.body.appendChild(hoverTip);
  rowsEl.addEventListener('mouseover', function (e) {
    var badge = e.target.closest('.r-warn[data-act="warn"]');
    if (!badge) return;
    var rowEl = badge.closest('.row');
    var it = rowEl && RM.itemById(state, rowEl.dataset.id);
    if (!it) return;
    var stH = rowEl.dataset.story ? storyById(it, rowEl.dataset.story) : null;
    var obj = stH || it;
    var list = stH ? [] : (validation.byItem[it.id] || []);
    if (!list.length && !obj.flag) return;
    hoverTip.innerHTML = (obj.flag ? '<div class="p-warnitem flag">' + esc(flagTitle(obj)) + '</div>' : '') + list.map(function (v) {
      var cls = v.level === 'error' ? 'err' : v.level;
      return '<div class="p-warnitem ' + cls + '">' + esc(v.msg) + '</div>';
    }).join('');
    hoverTip.hidden = false;
    var r = badge.getBoundingClientRect();
    var w = hoverTip.offsetWidth, h = hoverTip.offsetHeight;
    hoverTip.style.left = Math.min(r.right + 8, window.innerWidth - w - 10) + 'px';
    hoverTip.style.top = Math.max(6, Math.min(r.top - 4, window.innerHeight - h - 10)) + 'px';
  });
  rowsEl.addEventListener('mouseout', function (e) {
    if (e.target.closest && e.target.closest('.r-warn[data-act="warn"]')) hoverTip.hidden = true;
  });

  // outside Scoping a left-pane title is plain text: double-click (or
  // Rename… in the row menu) turns it into an input. The tooltip still
  // carries the description, behind the hint.
  function plNameHtml(it) {
    var s = it.feature || '';
    var tip = 'Double-click to rename' + (s ? '\n' + s : '') +
      (it.description ? '\n' + RM.htmlToText(it.description) : '');
    return '<span class="r-name r-name-txt' + (s ? '' : ' r-ph') + '" data-rowname title="' + esc(tip) + '">' +
      esc(s || '(untitled)') + '</span>';
  }
  function plStoryNameHtml(st) {
    var s = st.title || '';
    return '<span class="st-title st-title-txt' + (st.done ? ' done' : '') + (s ? '' : ' r-ph') +
      '" data-act="st-open" title="Open story · double-click to rename">' + esc(s || 'Story') + '</span>';
  }
  function itemRowsHtml(html, it, cyclic) {
    var meta = state.meta;
    var color = '#' + RM.colorForItem(state, it);
    var vlist = validation.byItem[it.id] || [];
    var sizeCls = !it.size ? (isScheduled(it) ? ' missing' : '') : (sizeMatches(it) ? '' : ' custom');
    var rk = RM.depRisk(state, it, cyclic);
    // the feature's risk chip — same markup in the Scoping grid and the
    // Planning row: computed under the auto scheme, picked otherwise
    function riskChipHtml() {
      if (!RM.riskEnabled(state)) return ''; // the column is up for story risk only
      var sch = RM.riskSchemeOf(state);
      if (sch === 'auto') {
        var autoTitle = rk.level === 'none' ? 'No dependency risk detected'
          : 'Dependency risk ' + rk.level + ':\n\u00B7 ' + rk.reasons.join('\n\u00B7 ');
        return '<span class="r-risk rk-' + rk.level + '" tabindex="0" title="' + esc(autoTitle) + '">' +
          (rk.level === 'none' ? '' : levelGlyph(rk.level.charAt(0).toUpperCase())) + '</span>';
      }
      var hintLevel = sch === 'risk' ? rk.level : 'none'; // graph hint only for risk
      var chipTitle = sch === 'risk' ? riskTitle : RM.riskColLabel(state) +
        (it.risk ? '\nNow: ' + riskValueLabel(it.risk) : '');
      return '<span class="r-risk rk-' + hintLevel + (it.risk ? ' has-risk' : '') +
        '" tabindex="0" role="button" data-act="risk" title="' + esc(chipTitle) + '">' +
        (it.risk ? levelGlyph(it.risk) : (hintLevel === 'none' ? '' : levelGlyph(hintLevel.charAt(0).toUpperCase()))) + '</span>';
    }
    var riskTitle = 'Risk' +
      (rk.level !== 'none' ? '\nDependency risk ' + rk.level + ':\n· ' + rk.reasons.join('\n· ') : '');
    var laneInner = '';
    var ports = '<span class="port p-in" data-port="in" title="Drag to another bar: this depends on it"></span>' +
      '<span class="port p-out" data-port="out" title="Drag to another bar: it depends on this"></span>';
    var doneCk = it.done
      ? '<span class="b-done" title="Done"><i data-lucide="circle-check"></i></span>' : '';
    if (isScheduled(it) && it.milestone) {
      // milestone: a diamond pinned to its fixed date, label at its right
      var msTip = it.feature + '  ·  ' + RM.fmtShort(RM.dayToDate(meta, it.startDay)) + '  ·  milestone' +
        (it.description ? '\n' + RM.htmlToText(it.description) : '');
      laneInner =
        '<div class="bar ms ms-' + RM.msStyleOf(it) + (isSel(it.id) && !selStory ? ' selected' : '') +
        (it.locked ? ' locked' : '') + (it.done ? ' done-bar' : '') +
        (showCrit && critCache && critCache.items[it.id] ? ' crit' : '') +
        '" data-bar="' + it.id + '"' +
        ' style="left:' + (it.startDay * dayPx()) + 'px;--bar-c:' + color + '"' +
        ' title="' + esc(msTip) + '">' +
        '<span class="ms-diamond"></span>' +
        '<span class="b-label out">' + doneCk + esc(it.feature) + '</span>' +
        ports +
        '</div>';
    } else if (isScheduled(it)) {
      var left = it.startDay * dayPx();
      var workW = Math.max(6, it.durDays * dayPx());
      var riskW = (it.riskDays || 0) * dayPx();
      var width = workW + riskW;
      var hcTag = '';
      // label rides inside the bar when it fits (~6.3px/char at 10.5px bold);
      // otherwise it sits just right of the bar in ink
      var labelW = it.feature.length * 6.3 + 16 + (hcTag ? 30 : 0) + (it.done ? 16 : 0);
      var label = '<span class="b-label' + (labelW <= width ? '' : ' out') + '">' + doneCk + esc(it.feature) + '</span>';
      var dates = RM.fmtShort(RM.dayToDate(meta, it.startDay)) + ' → ' +
        RM.fmtShort(RM.spanEndDate(meta, it.startDay, RM.itemSpan(it)));
      var tip = it.feature + '  ·  ' + dates + (it.size ? '  ·  ' + it.size : '') +
        (it.risk ? '  ·  risk ' + it.risk : '') +
        (it.description ? '\n' + RM.htmlToText(it.description) : '');
      // one uniform duration on the timeline — the work/risk split lives in
      // the panel, not the paint
      laneInner =
        '<div class="bar' + (isSel(it.id) && !selStory ? ' selected' : '') +
        (it.locked ? ' locked' : '') + (it.done ? ' done-bar' : '') + (width < 34 ? ' tiny' : '') +
        (showCrit && critCache && critCache.items[it.id] ? ' crit' : '') +
        '" data-bar="' + it.id + '"' +
        ' style="left:' + left + 'px;width:' + width + 'px;--bar-c:' + color + '"' +
        ' title="' + esc(tip) + '">' +
        '<div class="b-h l" data-act="bh-l"></div>' + label + hcTag +
        (it.locked ? '<span class="b-hc bi-lock" style="pointer-events:none"><i data-lucide="lock"></i></span>' : '') +
        // excluded from Auto sits in the lock's slot (the two never co-occur)
        (it.noAuto ? '<span class="b-hc bi-noauto" style="pointer-events:none" title="Excluded from Auto timeline"><i data-lucide="zap-off"></i></span>' : '') +
        '<div class="b-h r" data-act="bh-r"></div>' +
        ports +
        '</div>';
    }
    // option compare: the other option's schedule for this item rides behind
    // the live bar as a dashed ghost (skipped when the two agree). Days map
    // through calendar dates so options with different timelines still align.
    if (view !== 'scoping' && compareOptId) {
      var cmp = getCmp();
      var cIt = cmp && cmp.byId[it.id];
      if (cIt && cIt.startDay != null && (cIt.milestone || cIt.durDays != null)) {
        var cMeta = cmp.doc.meta;
        var gs = RM.dateToDay(meta, RM.dayToDate(cMeta, cIt.startDay));
        var ge = cIt.milestone ? gs
          : RM.dateToDay(meta, RM.spanEndDate(cMeta, cIt.startDay, RM.itemSpan(cIt)));
        if (ge == null || ge < gs) ge = gs + Math.max(1, RM.itemSpan(cIt)) - 1;
        var same = isScheduled(it) && !!cIt.milestone === !!it.milestone && gs === it.startDay &&
          (cIt.milestone || ge === it.startDay + RM.itemSpan(it) - 1);
        if (gs != null && gs >= 0 && !same) {
          var gCol = '#' + RM.colorForItem(cmp.doc, cIt);
          var gTip = cmp.name + ': ' + (cIt.milestone
            ? RM.fmtShort(RM.dayToDate(cMeta, cIt.startDay)) + '  ·  milestone'
            : RM.fmtShort(RM.dayToDate(cMeta, cIt.startDay)) + ' → ' +
              RM.fmtShort(RM.spanEndDate(cMeta, cIt.startDay, RM.itemSpan(cIt))));
          laneInner = (cIt.milestone
            ? '<div class="bar ms cmp ms-' + RM.msStyleOf(cIt) + '" style="left:' + (gs * dayPx()) + 'px;--bar-c:' + gCol +
              '" title="' + esc(gTip) + '"><span class="ms-diamond"></span></div>'
            : '<div class="bar cmp" style="left:' + (gs * dayPx()) + 'px;width:' +
              Math.max(6, (ge - gs + 1) * dayPx()) + 'px;--bar-c:' + gCol +
              '" title="' + esc(gTip) + '"></div>') + laneInner;
        }
      }
    }
    // hard deadline paint: a vertical tick on the deadline day plus a
    // connector from the bar's end — dashed red once the bar runs past it
    var dlHtml = '';
    if (view !== 'scoping' && it.deadline) {
      var dlDay = RM.deadlineDay(meta, it);
      if (dlDay != null) {
        var dlLate = RM.pastDeadline(meta, it);
        var dlX = (dlDay + 1) * dayPx();
        var dlTip = 'Deadline ' + RM.fmtShortYear(RM.parseISO(it.deadline));
        dlHtml = '<span class="r-dl-mark' + (dlLate ? ' late' : '') + '" style="left:' + dlX +
          'px" title="' + esc(dlTip) + '"></span>';
        if (isScheduled(it)) {
          var dlEndX = it.milestone ? (it.startDay + 0.5) * dayPx()
            : (it.startDay + Math.max(1, it.durDays + (it.riskDays || 0))) * dayPx();
          var dlx0 = Math.min(dlEndX, dlX), dlx1 = Math.max(dlEndX, dlX);
          if (dlx1 - dlx0 > 1) {
            dlHtml += '<span class="r-dl-line' + (dlLate ? ' late' : '') + '" style="left:' + dlx0 +
              'px;width:' + (dlx1 - dlx0) + 'px" title="' + esc(dlTip) + '"></span>';
          }
        }
      }
    }
    // unscheduled: the lane stays empty — hovering it previews the landing
    // slot (1 week unless sized), clicking places the item there
    if (view === 'scoping') {
      var cells = ['<div class="sc-row">'];
      var epIco2 = RM.iconForEpic(state, it.epic);
      var fixedContent = {
        size: it.milestone ? '<span class="r-size r-blank"></span>'
          : '<span class="r-size' + sizeCls + sizeRoCls(it) + '" tabindex="0" role="button" data-act="size" title="' + sizeChipTitle(it) + '">' + (it.size ? esc(it.size) : '') + '</span>',
        risk: riskChipHtml(),
        duration: it.milestone
          ? '<span class="r-wk editable" tabindex="0" role="button" data-act="wk" title="Milestone">0w</span>'
          : '<span class="r-wk editable" tabindex="0" role="button" data-act="wk" title="Duration">' +
            (isScheduled(it) || it.durDays != null ? totalWeeks(it) : '') + '</span>',
        workstream: '<span class="r-ws sc-chip" tabindex="0" role="button" data-act="ws" title="Workstream">' +
          '<span class="dd-dot" style="background:#' + RM.colorForWs(state, it.workstream) + '"></span>' +
          (it.workstream ? esc(shorten(it.workstream, 18))
            : '<i class="dws">' + esc(shorten(RM.defaultWsName(state), 18)) + '</i>') + '</span>',
        epic: '<span class="r-ws sc-chip" tabindex="0" role="button" data-act="epic" title="Epic">' +
          (epIco2 ? '<i data-lucide="' + epIco2 + '"></i>' : '') +
          (it.epic ? esc(shorten(it.epic, 20)) : '') + '</span>',
        start: '<span class="r-ws sc-chip" tabindex="0" role="button" data-act="startd" title="Start date">' +
          (isScheduled(it) ? esc(RM.fmtShort(RM.dayToDate(meta, it.startDay))) : '') + '</span>',
        deadline: (function () {
          var lateC = RM.pastDeadline(meta, it);
          return '<span class="r-ws sc-chip dl-chip' + (lateC ? ' late' : '') +
            '" tabindex="0" role="button" data-act="deadline" title="' +
            esc('Hard deadline' + (lateC ? '\nThe item runs past its deadline' : '')) + '">' +
            (it.deadline ? esc(RM.fmtShort(RM.parseISO(it.deadline))) : '') + '</span>';
        })(),
        priority: it.milestone ? '<span class="r-risk r-blank"></span>'
          : '<span class="r-risk pri' + (priChipHasValue(it) ? ' has-risk' : '') + priTierClass(it.priority) +
          '" tabindex="0" role="button" data-act="priority" title="' +
          esc(priChipTitle(it)) + '">' + priChipContent(it) + '</span>',
        assignees: '<span class="r-ws sc-chip" tabindex="0" role="button" data-act="asg" title="Assignees">' +
          (avatarStack(it.assignees) || '<i class="dws">+</i>') + '</span>'
      };
      allScopeCols().forEach(function (c) {
        if (isFixedColKey(c[0])) {
          cells.push('<div class="sc-cell sc-fix" data-col="' + c[0] + '" style="width:' + scopeColWidth(c) + 'px">' +
            fixedContent[c[0]] + '</div>');
          return;
        }
        if (!colShowsOn(c[0], 'feature')) {
          cells.push('<div class="sc-cell sc-na" data-col="' + c[0] + '" style="width:' + scopeColWidth(c) +
            'px" title="' + esc(scopeScopeLabel('story')) + '"></div>');
          return;
        }
        // every scope column holds rich text — edit it in place as such
        var sv = RM.scopeValue(it, c[0]);
        cells.push('<div class="sc-cell" data-col="' + c[0] + '" style="width:' + scopeColWidth(c) + 'px">' +
          '<div class="sc-edit sc-rich" contenteditable="true" data-scope="' + c[0] + '">' + richDisplay(sv) + '</div></div>');
      });
      cells.push('</div>');
      laneInner = cells.join('');
    }
    html.push(
      '<div class="row item' + (view === 'scoping' ? ' scope' : '') + (it.milestone ? ' ms' : '') +
      // when one of its stories is the selection, the feature is only the
      // selection's parent — marked, but never in the selected background
      (isSel(it.id) ? (selStory && selectedId === it.id ? ' sel-parent' : ' selected') : '') + (it.done ? ' done' : '') + (it.flag ? ' flagged' : '') +
      '" data-id="' + it.id + '">' +
      '<div class="row-left">' +
      '<span class="r-grip" data-act="grip"><i data-lucide="grip-vertical"></i></span>' +
      // milestones carry no stories; in story detail every feature is open
      // (no toggle to show); in feature detail the chevron opens/closes
      // just this item's stories
      (it.milestone || detailMode === 'story'
        ? '<span class="r-chev" aria-hidden="true"></span>'
        : '<span class="r-chev' + (expanded[it.id] ? ' open' : '') + '" data-act="stories" title="Stories (' + it.stories.length + ')">' +
          (it.stories.length ? '<i data-lucide="chevron-right"></i>' : '<span style="opacity:.35"><i data-lucide="chevron-right"></i></span>') + '</span>') +
      '<span class="r-num">' + it.num + '</span>' +
      '<div class="r-main">' +
      (it.locked ? '<span class="r-lock"><i data-lucide="lock"></i></span>' : '') +
      (it.noAuto ? noAutoIcon() : '') +
      (it.done ? '<span class="r-doneck" title="Done"><i data-lucide="circle-check"></i></span>' : '') +
      typeGlyphHtml(it, 'feature') +
      (view === 'scoping'
        // scoping: the title is a full-height editable cell in the tab ring,
        // top-aligned and wrapping like every other cell
        ? '<div class="r-name sc-name" contenteditable="true" spellcheck="false" aria-label="Feature title">' + esc(it.feature) + '</div>'
        : plNameHtml(it)) +
      // no epic tag beside the title: the Epic column (Planning) / Epic
      // cell (Scoping) carries it when wanted
      '</div>' +
      (view === 'scoping' ? '' : (function () {
        // the planning chips follow the user's column order/visibility
        var chips = {
          size: RM.sizingEnabled(state) && !it.milestone
            ? '<span class="r-size' + sizeCls + sizeRoCls(it) + '" tabindex="0" role="button" data-act="size" title="' + sizeChipTitle(it) + '">' + (it.size ? esc(it.size) : '·') + '</span>'
            : '<span class="r-size r-blank"></span>',
          pri: RM.priorityEnabled(state) && !it.milestone
            ? '<span class="r-risk pri' + (priChipHasValue(it) ? ' has-risk' : '') + priTierClass(it.priority) + '" tabindex="0" role="button" data-act="priority" title="' + esc(priChipTitle(it)) + '">' + (priChipContent(it) || '·') + '</span>'
            : '<span class="r-risk pri r-blank"></span>',
          risk: riskChipHtml(),
          dur: it.milestone
            ? '<span class="r-wk editable" tabindex="0" role="button" data-act="wk" title="Milestone">◆</span>'
            : '<span class="r-wk editable" tabindex="0" role="button" data-act="wk" title="Duration">' + totalWeeks(it) + '</span>',
          asg: '<span class="r-asg" tabindex="0" role="button" data-act="asg" title="Assignees">' +
            (avatarStack(it.assignees, 2) || '<i data-lucide="user-plus"></i>') + '</span>',
          cap: state.meta.capacityEnabled && RM.planLevel(state) === 'feature' && !it.milestone
            ? '<span class="r-cap' + (RM.itemCapType(state, it) ? '' : ' empty') + (itemCapInherited(it) ? ' inherited' : '') + '" tabindex="0" role="button" data-act="cap" title="' +
              esc(itemCapInherited(it) ? CAP_INHERIT_TITLE : 'Capacity type' + (RM.itemCapType(state, it) ? '\nNow: ' + RM.itemCapType(state, it) : '')) + '">' +
              (RM.itemCapType(state, it) ? esc(RM.itemCapType(state, it)) : '·') + '</span>'
            : '<span class="r-cap r-blank"></span>',
          mult: state.meta.capacityEnabled && RM.planLevel(state) === 'feature' && !it.milestone && state.meta.capMode !== 'points'
            ? '<span class="r-mult editable' + ((it.capMult || 1) === 1 ? ' one' : '') + '" tabindex="0" role="button" data-act="mult" title="Capacity multiplier — people this needs at once">×' + fmtPe(it.capMult || 1) + '</span>'
            : '<span class="r-mult r-blank"></span>',
          // the fixed fields Scoping also shows — same chips, same editors.
          // milestones carry a start date and nothing else here.
          ws: it.milestone ? '<span class="r-ws-col r-blank"></span>'
            : '<span class="r-ws-col" tabindex="0" role="button" data-act="ws" title="Workstream">' +
              '<span class="dd-dot" style="background:#' + RM.colorForWs(state, it.workstream) + '"></span>' +
              (it.workstream ? esc(shorten(it.workstream, 18))
                : '<i class="dws">' + esc(shorten(RM.defaultWsName(state), 18)) + '</i>') + '</span>',
          epic: it.milestone ? '<span class="r-epic-col r-blank"></span>'
            : '<span class="r-epic-col" tabindex="0" role="button" data-act="epic" title="Epic">' +
              (RM.iconForEpic(state, it.epic) ? '<i data-lucide="' + RM.iconForEpic(state, it.epic) + '"></i>' : '') +
              (it.epic ? esc(shorten(it.epic, 18)) : '') + '</span>',
          start: '<span class="r-date-col" tabindex="0" role="button" data-act="startd" title="Start date">' +
            (isScheduled(it) ? esc(RM.fmtShort(RM.dayToDate(meta, it.startDay))) : '') + '</span>',
          deadline: it.milestone ? '<span class="r-date-col dl-chip r-blank"></span>' : (function () {
            var lateP = RM.pastDeadline(meta, it);
            return '<span class="r-date-col dl-chip' + (lateP ? ' late' : '') +
              '" tabindex="0" role="button" data-act="deadline" title="' +
              esc('Hard deadline' + (lateP ? '\nThe item runs past its deadline' : '')) + '">' +
              (it.deadline ? esc(RM.fmtShort(RM.parseISO(it.deadline))) : '') + '</span>';
          })()
        };
        return plColsVisible().map(function (k) { return chips[k]; }).join('');
      })()) +
      warnBadge(it) +
      '</div>' +
      '<div class="row-lane">' + laneInner + dlHtml + '</div>' +
      '</div>');

    if ((detailMode === 'story' || expanded[it.id]) && !it.milestone) {
      it.stories.forEach(function (st) {
        var stSched = st.startDay != null && st.durDays != null;
        html.push(
          '<div class="row story' + (selStory === st.id ? ' selected' : '') + (st.flag ? ' flagged' : '') +
          '" data-story="' + st.id + '" data-id="' + it.id + '">' +
          '<div class="row-left"><span class="st-pad"><span class="st-grip" title="Drag to reorder or move to another feature"><i data-lucide="grip-vertical"></i></span></span>' +
          // the story number sits in the same column as the feature numbers above
          '<span class="r-num st-num">#' + st.num + '</span>' +
          (st.noAuto ? noAutoIcon() : '') +
          (st.done ? '<span class="r-doneck st-doneck" title="Done"><i data-lucide="circle-check"></i></span>' : '') +
          typeGlyphHtml(st, 'story', it) +
          (view === 'scoping'
            // scoping: the story title edits in place like the feature titles
            ? '<div class="st-title st-name' + (st.done ? ' done' : '') + '" contenteditable="true" spellcheck="false" aria-label="Story title">' + esc(st.title) + '</div>'
            : plStoryNameHtml(st)) +
          (view === 'scoping' ? '' : plColsVisible().map(function (k) {
            if (k === 'asg') {
              return '<span class="r-asg" tabindex="0" role="button" data-act="st-asg" title="Story assignees">' +
                (avatarStack(st.assignees, 2) || '<i data-lucide="user-plus"></i>') + '</span>';
            }
            // workstream and epic belong to the feature: shown rolled up,
            // dimmed and italic, and never clickable on a story row
            if (k === 'ws') {
              return '<span class="r-ws-col roll" title="Rolls up from the feature">' +
                '<span class="dd-dot" style="background:#' + RM.colorForWs(state, it.workstream) + '"></span>' +
                esc(shorten(it.workstream || RM.defaultWsName(state), 18)) + '</span>';
            }
            if (k === 'epic') {
              return '<span class="r-epic-col roll" title="Rolls up from the feature">' +
                (RM.iconForEpic(state, it.epic) ? '<i data-lucide="' + RM.iconForEpic(state, it.epic) + '"></i>' : '') +
                (it.epic ? esc(shorten(it.epic, 18)) : '') + '</span>';
            }
            // the story's own dates, on the same calendars Scoping opens
            if (k === 'start') {
              return '<span class="r-date-col" tabindex="0" role="button" data-act="st-startd" title="Story start">' +
                (st.startDay != null ? esc(RM.fmtShort(RM.dayToDate(meta, st.startDay))) : '') + '</span>';
            }
            if (k === 'deadline') {
              var stLateP = RM.pastDeadline(meta, st);
              return '<span class="r-date-col dl-chip' + (stLateP ? ' late' : '') +
                '" tabindex="0" role="button" data-act="st-dl" title="' +
                esc('Story deadline' + (stLateP ? '\nThe story runs past its deadline' : '')) + '">' +
                (st.deadline ? esc(RM.fmtShort(RM.parseISO(st.deadline))) : '') + '</span>';
            }
            // keep the column aligned even when the story scale is off
            return storyChipHtml(k, st, 'data-act') ||
              '<span class="' + (k === 'size' ? 'r-size' : k === 'dur' ? 'r-wk' : k === 'cap' ? 'r-cap'
                : k === 'mult' ? 'r-mult' : k === 'pri' ? 'r-risk pri' : 'r-risk') + ' r-blank"></span>';
          }).join('') + (st.flag ? flagBadgeHtml(st) : '<span class="r-warn"></span>')) +
          '</div>' +
          (view === 'scoping'
            // scoping: stories share the grid — text columns, Size, Assignees,
            // Priority and Duration edit in place; the rest rolls up from the
            // feature and reads dimmed + italic
            ? '<div class="row-lane"><div class="sc-row">' + allScopeCols().map(function (c) {
                var key = c[0];
                var w = scopeColWidth(c);
                if (!isFixedColKey(key)) {
                  if (!colShowsOn(key, 'story')) {
                    return '<div class="sc-cell sc-na" data-col="' + key + '" style="width:' + w +
                      'px" title="' + esc(scopeScopeLabel('feature')) + '"></div>';
                  }
                  var sval = RM.storyScopeValue(st, key);
                  return '<div class="sc-cell" data-col="' + key + '" style="width:' + w + 'px">' +
                    '<div class="sc-edit sc-rich" contenteditable="true" data-stscope="' + key + '">' + richDisplay(sval) + '</div></div>';
                }
                function stFix(inner) {
                  return '<div class="sc-cell sc-fix" data-col="' + key + '" style="width:' + w + 'px">' + inner + '</div>';
                }
                if (key === 'size' && RM.sizingEnabled(state, 'story')) return stFix(storyChipHtml('size', st, 'data-act', ''));
                if (key === 'risk' && RM.riskEnabled(state, 'story')) return stFix(storyChipHtml('risk', st, 'data-act', ''));
                if (key === 'assignees') {
                  return stFix('<span class="r-ws sc-chip" tabindex="0" role="button" data-act="st-asg" title="Story assignees">' +
                    (avatarStack(st.assignees) || '<i class="dws">+</i>') + '</span>');
                }
                // stories rank on their own priority scheme (never RICE)
                if (key === 'priority' && RM.priorityEnabled(state, 'story')) return stFix(storyChipHtml('pri', st, 'data-act', ''));
                if (key === 'duration') return stFix(storyChipHtml('dur', st, 'data-act', ''));
                if (key === 'start') {
                  return stFix('<span class="r-ws sc-chip" tabindex="0" role="button" data-act="st-startd" title="Story start">' +
                    (st.startDay != null ? esc(RM.fmtShort(RM.dayToDate(state.meta, st.startDay))) : '') + '</span>');
                }
                if (key === 'deadline') {
                  var stLate = RM.pastDeadline(state.meta, st);
                  return stFix('<span class="r-ws sc-chip dl-chip' + (stLate ? ' late' : '') +
                    '" tabindex="0" role="button" data-act="st-dl" title="' +
                    esc('Story deadline' + (stLate ? '\nThe story runs past its deadline' : '')) + '">' +
                    (st.deadline ? esc(RM.fmtShort(RM.parseISO(st.deadline))) : '') + '</span>');
                }
                // rolled up from the feature — shown dimmed and italic
                var roll = key === 'workstream' ? esc(it.workstream || RM.defaultWsName(state))
                  : key === 'epic' ? esc(it.epic || '')
                  : '';
                return '<div class="sc-cell sc-fix sc-na" data-col="' + key + '" style="width:' + w +
                  'px" title="Rolls up from the feature">' +
                  (roll ? '<span class="sc-roll">' + roll + '</span>' : '') + '</div>';
              }).join('') + '</div></div></div>'
            : '<div class="row-lane"' + (!stSched && view === 'planning' ? ' title="Double-click to add a timeline"' : '') + '>' +
              (stSched
                ? (function () {
                    var stW = Math.max(6, st.durDays * dayPx());
                    // quiet label: inside when it fits, else right of the bar
                    var stLabW = st.title.length * 5.5 + 12;
                    return '<div class="st-bar' + (selStory === st.id ? ' selected' : '') + '" data-stbar="' + st.id + '" data-id="' + it.id + '" style="left:' + (st.startDay * dayPx()) +
                      'px;width:' + stW + 'px;--bar-c:' + color + '">' +
                      '<span class="stb-label' + (stLabW <= stW ? '' : ' out') + '">' + esc(st.title) + '</span>' +
                      '<span class="bh l" data-act="sh-l"></span><span class="bh r" data-act="sh-r"></span>' +
                      // same in/out circles as feature bars — stories link to stories
                      '<span class="port p-in" data-port="in" title="Drag to another story: this depends on it"></span>' +
                      '<span class="port p-out" data-port="out" title="Drag to another story: it depends on this"></span>' +
                      '</div>';
                  })()
                // no timeline of its own: the title sits, faint and boxless,
                // where the feature starts — it rides along with the feature
                : (isScheduled(it)
                    ? '<span class="st-ghost" style="left:' + (it.startDay * dayPx()) + 'px">' + esc(st.title) + '</span>'
                    : '')) +
              '</div></div>'));
      });
      if (!(it.stories || []).length) html.push( // a feature with stories adds via the story context menu
        '<div class="row story story-add" data-id="' + it.id + '">' +
        '<div class="row-left"><span class="st-pad"></span>' +
        '<i data-lucide="plus" class="st-add-ico"></i>' +
        '<input class="st-add-input" data-act="st-add" placeholder="Add ' + esc(lvl('story').toLowerCase()) + '…">' +
        '</div><div class="row-lane"></div></div>');
    }
  }

  function renderRows() {
    var html = [];
    var cyclic = RM.cycleMembers(state);
    state.phases.forEach(function (p) {
      var items = RM.itemsInPhase(state, p.id).filter(matchesFilter);
      var bandLane = p.description ? '<span class="band-desc" title="' + esc(RM.htmlToText(p.description)) + '">' + esc(RM.htmlToText(p.description)) + '</span>' : '';
      html.push(
        '<div class="row band" data-kind="band" data-phase="' + p.id + '">' +
        '<div class="row-left">' +
        '<span class="band-chev' + (p.collapsed ? '' : ' open') + '" data-act="phase-toggle" title="Collapse / expand"><i data-lucide="chevron-right"></i></span>' +
        '<span class="band-name">' + esc(p.name) + '</span>' +
        '<span class="band-count">' + items.length + '</span>' +
        (p.bucket ? '<span class="band-bucket-tag">backlog</span>' : '') +
        // the phase's actions sit at the right edge of the band, always visible
        '<span class="band-acts">' +
        '<button class="band-add" data-act="phase-additem" title="' + esc('Add a ' + lvl('feature').toLowerCase() + ' to this phase') + '">+ ' + esc(lvl('feature').toLowerCase()) + '</button>' +
        '<button class="band-edit" data-act="phase-edit" title="Edit phase">edit</button>' +
        (p.bucket ? '' : (function () {
          var as = autoPhaseStatus(p.id);
          return '<button class="band-zap" data-act="phase-auto" title="' + esc(as.tip) + '"' + (as.disabled ? ' disabled' : '') +
            '><i data-lucide="zap"></i></button>';
        })()) +
        '</span>' +
        '</div>' +
        '<div class="row-lane">' + bandLane + '</div>' +
        '</div>');
      if (p.collapsed) return;

      // grouping hierarchy: phase > workstream > epic ("no workstream" last)
      function partition(list, field, emptyLast) {
        var keys = [], by = {};
        list.forEach(function (it) {
          var key = it[field] || '';
          if (!by[key]) { by[key] = []; keys.push(key); }
          by[key].push(it);
        });
        if (emptyLast && keys.indexOf('') !== -1) {
          keys = keys.filter(function (k) { return k !== ''; }).concat(['']);
        }
        return { keys: keys, by: by };
      }
      function epicBands(list, sub, wsOf) {
        wsKey = wsOf == null ? null : wsOf;
        var g = partition(list, 'epic', false);
        g.keys.forEach(function (key) {
          html.push(
            '<div class="row eband' + (sub ? ' sub' : '') + '" data-kind="eband" data-phase="' + p.id +
            '" data-epic="' + esc(key) + '">' +
            '<div class="row-left">' +
            '<span class="r-ico"><i data-lucide="' + (RM.iconForEpic(state, key) || 'tag') + '"></i></span>' +
            '<span class="eb-name">' + (key ? esc(key) : '<i>no epic</i>') +
            '</span><span class="band-count">' + g.by[key].length + '</span></div>' +
            '<div class="row-lane"></div></div>');
          g.by[key].forEach(function (it) { itemRowsHtml(html, it, cyclic); });
        });
      }
      // click-to-add row, shown only while the phase has no features: a
      // filled phase adds through Insert feature above / below
      function addRowHtml(ph) {
        return '<div class="row addrow" data-kind="addrow" data-phase="' + ph.id + '">' +
          '<div class="row-left" title="Add a ' + esc(lvl('feature').toLowerCase()) + ' to ' + esc(ph.name) + '"><span class="addrow-lab"><i data-lucide="plus"></i> ' + esc('Add ' + lvl('feature')) + '</span></div>' +
          '<div class="row-lane"></div></div>';
      }
      var wsKey = null;
      if (groupWs && state.meta.workstreamsEnabled) {
        var wg = partition(items, 'workstream', true);
        wg.keys.forEach(function (key) {
          html.push(
            '<div class="row eband wsband" data-kind="eband" data-phase="' + p.id +
            '" data-ws="' + esc(key) + '">' +
            '<div class="row-left">' +
            '<span class="r-dot" style="background:#' + RM.colorForWs(state, key) + '"></span>' +
            '<span class="eb-name">' + (key ? esc(key) : esc(RM.defaultWsName(state)) + ' <i style="font-weight:400">default</i>') +
            '</span><span class="band-count">' + wg.by[key].length + '</span></div>' +
            '<div class="row-lane"></div></div>');
          if (groupEpic) epicBands(wg.by[key], true, key);
          else {
            wg.by[key].forEach(function (it) { itemRowsHtml(html, it, cyclic); });

          }
        });
      } else if (groupEpic) {
        epicBands(items, false);
      } else {
        items.forEach(function (it) { itemRowsHtml(html, it, cyclic); });
      }

      // click-to-add row at the bottom of the phase — only when the phase
      // isn't already split into groups that carry their own add rows
      var grouped = (groupWs && state.meta.workstreamsEnabled) || groupEpic;
      if (!items.length) html.push(addRowHtml(p));
    });

    rowsEl.innerHTML = html.join('');
    if (window.lucide) lucide.createIcons();
  }

  // total effort (work + risk) in weeks for the row readout. Scheduled items
  // count WORKING days only — a bar stretched across a holiday still reads as
  // its clean size (e.g. 8w, not 8.2w)
  function totalWeeks(it) {
    // an explicitly-set duration wins even off the timeline; else the size
    var days = isScheduled(it)
      ? RM.workInSpan(state.meta, it.startDay, RM.itemSpan(it))
      : (it.durDays != null ? it.durDays
        : RM.effortDays(state, it) + RM.riskEffortDays(state, it));
    if (!days) return '·';
    return fmtDays(days);
  }

  function sizeHuman(size, kind) {
    var d = RM.sizeDays(state, size, kind);
    if (d == null) return '';
    if (d < 5) return (Math.round(d * 10) / 10) + 'd';
    return (Math.round(d / SPW() * 100) / 100) + 'w';
  }
  function sizeMatches(it) {
    if (!it.size || !isScheduled(it)) return true;
    // a derived size is whatever the bar says it is — never a mismatch
    if (RM.sizeRollup(state)) return true;
    // at the Stories level a feature whose bar is the hull of its scheduled
    // stories draws from the stories, not from its size
    if (RM.planLevel(state) === 'story' && (it.stories || []).some(function (st) { return st.startDay != null && st.durDays != null; })) return true;
    var work = RM.workInSpan(state.meta, it.startDay, it.durDays);
    return work === RM.itemSizeDays(state, it, snapOpts());
  }

  // ------------------------------------------------------------ arrows
  function barRect(id) {
    var el = rowsEl.querySelector('[data-bar="' + id + '"]');
    if (!el) return null;
    var g = grid.getBoundingClientRect();
    var r = el.getBoundingClientRect();
    return { left: r.left - g.left, right: r.right - g.left, cy: r.top - g.top + r.height / 2, top: r.top - g.top, bottom: r.bottom - g.top };
  }
  // the same measurement for a story's own little bar
  function storyBarRect(stId) {
    var el = rowsEl.querySelector('[data-stbar="' + stId + '"]');
    if (!el) return null;
    var g = grid.getBoundingClientRect();
    var r = el.getBoundingClientRect();
    return { left: r.left - g.left, right: r.right - g.left, cy: r.top - g.top + r.height / 2, top: r.top - g.top, bottom: r.bottom - g.top };
  }

  function renderArrows() {
    var gEl = $('#arrowPaths');
    if (depsMode === 'none' || drag) { gEl.innerHTML = ''; return; }
    var svg = $('#arrows');
    svg.setAttribute('width', grid.scrollWidth);
    svg.setAttribute('height', grid.scrollHeight);
    var edges = RM.depEdges(state);
    var out = [];
    edges.forEach(function (e) {
      var dep = e[0], it = e[1];
      var viol = it.startDay != null && it.startDay < RM.itemEnd(dep) && !dep.done;
      var related = selectedId && (selectedId === dep.id || selectedId === it.id);
      var crit = showCrit && critCache && critCache.edges[dep.id + '>' + it.id];
      var explicit = it.deps.indexOf(dep.num) !== -1;
      var edgeSel = selectedEdge && selectedEdge.fromId === dep.id && selectedEdge.toId === it.id;
      // arrows ON: every specifically-defined dependency renders (selected
      // item's in blue, critical path orange, violations dashed amber)
      if (!explicit && !viol && !crit && !edgeSel) return;
      var a = barRect(dep.id), b = barRect(it.id);
      if (!a || !b) return;
      var d = curvePath(a.right + 1, a.cy, b.left - 3, b.cy);
      out.push('<g class="edge' + (crit ? ' crit' : '') + (viol ? ' viol' : '') + (related ? ' sel-related' : '') +
        (edgeSel ? ' hot' : '') +
        '" data-from="' + dep.id + '" data-to="' + it.id + '" data-explicit="' + explicit + '">' +
        '<path class="hit" d="' + d + '"></path>' +
        '<path class="vis" d="' + d + '"></path></g>');
    });
    // story → story arrows, drawn only when both little bars are on screen
    RM.storyDepEdges(state).forEach(function (e) {
      var dep = e[0], ref = e[1];
      var a = storyBarRect(dep.st.id), b = storyBarRect(ref.st.id);
      if (!a || !b) return;
      var dw = RM.storyWindow(state, dep.it, dep.st);
      var w = RM.storyWindow(state, ref.it, ref.st);
      var viol = !!(w && dw && w.startDay < dw.endDay && !dep.st.done);
      var related = selStory && (selStory === dep.st.id || selStory === ref.st.id);
      var edgeSel = selectedEdge && selectedEdge.sfrom === dep.st.id && selectedEdge.sto === ref.st.id;
      var d = curvePath(a.right + 1, a.cy, b.left - 3, b.cy);
      out.push('<g class="edge story' + (viol ? ' viol' : '') + (related ? ' sel-related' : '') +
        (edgeSel ? ' hot' : '') +
        '" data-sfrom="' + dep.st.id + '" data-sto="' + ref.st.id + '" data-explicit="true">' +
        '<path class="hit" d="' + d + '"></path>' +
        '<path class="vis" d="' + d + '"></path></g>');
    });
    gEl.innerHTML = out.join('');
  }

  // one smooth cubic curve; loops back naturally when the target is behind
  function curvePath(sx, sy, tx, ty) {
    var dx = Math.max(28, Math.min(90, Math.abs(tx - sx) * 0.5));
    return 'M' + sx + ',' + sy + ' C' + (sx + dx) + ',' + sy + ' ' + (tx - dx) + ',' + ty + ' ' + tx + ',' + ty;
  }

  // arrow click selects the edge; Delete removes it
  var selectedEdge = null; // { fromId, toId } for features | { sfrom, sto } for stories
  $('#arrows').addEventListener('click', function (e) {
    var g = e.target.closest('g.edge');
    if (!g) return;
    selectedEdge = g.dataset.sfrom
      ? { sfrom: g.dataset.sfrom, sto: g.dataset.sto }
      : { fromId: g.dataset.from, toId: g.dataset.to };
    requestAnimationFrame(renderArrows);
  });
  function deleteSelectedEdge() {
    if (!selectedEdge) return false;
    if (selectedEdge.sfrom) {
      var dep = RM.storyRef(state, selectedEdge.sfrom);
      var ref = RM.storyRef(state, selectedEdge.sto);
      selectedEdge = null;
      if (!dep || !ref) { requestAnimationFrame(renderArrows); return true; }
      var depNum = dep.st.num;
      commit('remove story dep', function (s) {
        var t = RM.storyRef(s, ref.st.id);
        if (t) t.st.deps = (t.st.deps || []).filter(function (n) { return n !== depNum; });
      });
      toast('Removed: ' + RM.storyLabel(state, ref) + ' no longer depends on ' + RM.storyLabel(state, dep));
      return true;
    }
    var from = RM.itemById(state, selectedEdge.fromId);
    var to = RM.itemById(state, selectedEdge.toId);
    selectedEdge = null;
    if (!from || !to) { requestAnimationFrame(renderArrows); return true; }
    commit('remove dep', function (s) {
      var t = RM.itemById(s, to.id);
      t.deps = t.deps.filter(function (n) { return n !== from.num; });
    });
    toast('Removed: #' + to.num + ' no longer depends on #' + from.num);
    return true;
  }

  // ------------------------------------------------------------ selection & panel
  var selStory = null; // story id when the panel shows a story (selectedId = its item)
  function selIds() { return multiSel || (selectedId ? [selectedId] : []); }
  function isSel(id) { return multiSel ? multiSel.indexOf(id) !== -1 : selectedId === id; }

  function select(id, scrollTo) {
    flushPanelEdit();
    multiSel = null;
    selectedId = id;
    selStory = null;
    selectedEdge = null;
    render();
    if (scrollTo && id) {
      var el = rowsEl.querySelector('.row[data-id="' + id + '"]');
      if (el && el.scrollIntoView) el.scrollIntoView({ block: 'center' });
      var it = RM.itemById(state, id);
      if (it && isScheduled(it) && view === 'planning') {
        var x = it.startDay * dayPx();
        var laneVis = Math.max(80, board.clientWidth - leftW());
        if (x < board.scrollLeft + 10 || x > board.scrollLeft + laneVis - 60) {
          scrollLaneTo(x);
        }
      }
    }
  }
  // shift-click: contiguous range from the anchor, in visible row order
  function extendSelectionTo(itemId) {
    if (!selectedId || selectedId === itemId) { select(itemId); return; }
    flushPanelEdit();
    var order = $$('#rows .row.item').map(function (r) { return r.dataset.id; });
    var a = order.indexOf(selectedId), b = order.indexOf(itemId);
    if (a === -1 || b === -1) { select(itemId); return; }
    multiSel = order.slice(Math.min(a, b), Math.max(a, b) + 1);
    selStory = null;
    selectedEdge = null;
    render();
  }

  function selectStory(itemId, stId) {
    flushPanelEdit();
    multiSel = null;
    selectedId = itemId;
    selStory = stId;
    selectedEdge = null;
    render();
  }

  // the panel is rebuilt from scratch on every render: keep its scroll offset,
  // and during a disk reload hand focus back to the field that had it
  function panelFlagHtml(x) {
    return x.flag ? '<span class="p-flag" title="' + esc(flagTitle(x)) + '"><i data-lucide="flag"></i>' + (x.flag.reason ? esc(shorten(x.flag.reason, 28)) : 'Flagged') + '</span>' : '';
  }
  function renderPanel() {
    var panel = $('#panel');
    var keepTop = panel && !panel.hidden ? panel.scrollTop : 0;
    var ae = document.activeElement;
    var refocus = reloadingDoc && ae && panel && panel.contains(ae) ? panelFieldSelector(ae) : null;
    renderPanelInner();
    // #panelPeek floats over the top-right corner of the board, which is where
    // the assistant drawer's close button lands too. Flag the state on <body>
    // so CSS can pad .ai-head clear of the toggle (Prioritizing hides the
    // toggle in CSS, so it never counts).
    var peekEl = $('#panelPeek');
    document.body.classList.toggle('peek-on', !!(peekEl && !peekEl.hidden) && view !== 'prio');
    if (panel && !panel.hidden && keepTop) panel.scrollTop = keepTop;
    if (refocus) {
      var el = panel.querySelector(refocus);
      if (el && el.focus) el.focus({ preventScroll: true });
    }
  }
  function panelFieldSelector(el) {
    if (el.id) return '#' + el.id;
    var f = el.dataset && (el.dataset.f || el.dataset.stf);
    if (!f) return null;
    var key = el.dataset.f ? 'data-f' : 'data-stf';
    return '[' + key + '="' + f + '"]' + (el.dataset.v != null ? '[data-v="' + el.dataset.v + '"]' : '') +
      (el.dataset.k != null ? '[data-k="' + el.dataset.k + '"]' : '');
  }
  function renderPanelInner() {
    var panel = $('#panel');
    var peek = $('#panelPeek');
    // the panel lives on Planning, Scoping AND Sprinting: persistent, collapsible.
    // On Prioritizing it appears only while a card is selected (click to
    // show, click the same card to hide)
    if (view !== 'planning' && view !== 'scoping' && view !== 'sprints' && view !== 'prio') {
      panel.hidden = true; panel.innerHTML = '';
      if (peek) peek.hidden = true;
      return;
    }
    if (view === 'prio' && !selectedId) {
      panel.hidden = true; panel.innerHTML = '';
      if (peek) peek.hidden = true;
      return;
    }
    if (!panelOpen) {
      panel.hidden = true; panel.innerHTML = '';
      if (peek) peek.hidden = false;
      return;
    }
    if (peek) peek.hidden = true;
    var it = selectedId ? RM.itemById(state, selectedId) : null;
    if (!it) {
      panel.hidden = false;
      panel.innerHTML =
        '<div id="panelRz"></div>' +
        '<div class="p-top"><button class="p-close" data-f="collapse" title="Hide panel  ]"><i data-lucide="panel-right-close"></i></button></div>' +
        '<div class="p-empty">No item selected<span>Click a row ' +
        (view === 'sprints' ? '' : 'on the timeline ') + 'to edit it here.</span></div>';
      if (window.lucide) lucide.createIcons();
      return;
    }
    if (selStory) {
      var stSel = storyById(it, selStory);
      if (!stSel) { selStory = null; } else {
        panel.hidden = false;
        renderStoryPanel(panel, it, stSel);
        return;
      }
    }
    panel.hidden = false;
    var meta = state.meta;
    var res = RM.resolveDeps(state, it);
    var vlist = validation.byItem[it.id] || [];

    var phaseCur = phaseOf(it);
    var phaseDd = ddButton('phase', esc(phaseCur.name) + (phaseCur.bucket ? ' <small>(backlog)</small>' : ''), null, 'Phase');
    var epicDd = ddButton('epic',
      (RM.iconForEpic(state, it.epic) ? '<i data-lucide="' + RM.iconForEpic(state, it.epic) + '"></i> ' : '') +
      (it.epic ? esc(it.epic) : '<i>— none —</i>'), null, 'Epic');
    var typeDd = ddButton('teamType', it.teamType ? esc(it.teamType) : 'Any role', null, 'Which role works this item');

    var sizeBtns = RM.sizeRollup(state)
      ? '<span class="p-rollup" title="Sum of the story points">' +
        (it.size ? esc(it.size) + ' pt <small>' + esc(fmtDays(RM.rollupDays(state, it) || 0)) + '</small>' : '<i>no sized ' + esc(lvl('story', true).toLowerCase()) + '</i>') +
        '</span>'
      : RM.sizeOrderOf(state).map(function (s) {
        return '<button data-f="size" data-v="' + esc(s) + '"' + (it.size === s ? ' class="on"' : '') +
          ' title="' + fmtDays(RM.sizeDays(state, s)) + '">' + esc(s) + '</button>';
      }).join('');
    var riskBtns = ['<button data-f="riskSize" data-v=""' + (!it.risk ? ' class="on"' : '') + ' title="Not set">None</button>']
      .concat(RM.riskOrderOf(state).map(function (s) {
        return '<button data-f="riskSize" data-v="' + s + '"' + (it.risk === s ? ' class="on"' : '') +
          ' title="' + esc(riskValueLabel(s)) + '">' + levelGlyph(s) + '</button>';
      })).join('');
    var priInfo = '';
    if (it.milestone) {
      // milestones carry no priority
    } else if (RM.prioritySchemeOf(state) === 'rice') {
      // the RICE scheme edits through dropdowns; the score is the priority
      priInfo = '<label class="p-lab" style="margin-top:10px">RICE score' +
        (riceScoreLabel(it) ? ' · ' + riceScoreLabel(it) : '') + '</label>' +
        '<div class="rice-pop rice-panel">' + riceSelectsHtml(it) + '</div>';
    } else if (RM.priorityEnabled(state)) {
      var priLevels = RM.prioritySchemeOf(state) === 'levels';
      priInfo = '<label class="p-lab" style="margin-top:10px">Priority</label>' +
        '<div class="seg">' +
        ['<button data-f="priSet" data-v=""' + (!it.priority ? ' class="on"' : '') + ' title="Not set">None</button>']
          .concat(RM.priorityOrderOf(state).map(function (pv) {
            return '<button data-f="priSet" data-v="' + pv + '" class="' + (it.priority === pv ? 'on' : '') + priTierClass(pv) + '"' +
              ' title="' + esc(priorityValueLabel(pv)) + '">' + (priLevels ? levelGlyph(pv) : pv) + '</button>';
          })).join('') +
        '</div>';
    }

    var scheduleInfo = '';
    var startD = isScheduled(it) ? RM.dayToDate(meta, it.startDay) : null;
    if (it.milestone) {
      scheduleInfo = isScheduled(it)
        ? '<div><label class="p-lab">Fixed date</label>' +
          '<input type="text" readonly class="cal-in" data-f="startDate" value="' + RM.fmtISO(startD) + '" style="width:100%"></div>' +
          '<div class="p-row" style="margin-top:8px">' +
          '<button data-f="snap" title="Earliest date after dependencies">Snap earliest</button>' +
          '<button data-f="unschedule">Unschedule</button>' +
          '</div>'
        : '<div class="p-row" style="margin-top:8px">' +
          '<button data-f="schedule-now" class="primary">Place on timeline</button>' +
          '<button data-f="snap">Snap earliest</button>' +
          '</div>';
    } else if (isScheduled(it)) {
      scheduleInfo =
        '<div class="p-grid2">' +
        '<div><label class="p-lab">Start</label>' +
        '<input type="text" readonly class="cal-in" data-f="startDate" value="' + RM.fmtISO(startD) + '" style="width:100%"></div>' +
        '<div><label class="p-lab">Duration (weeks)</label>' +
        '<input type="number" data-f="durWeeks" min="0.2" step="0.2" value="' + (Math.round(it.durDays / SPW() * 100) / 100) + '" style="width:100%"></div>' +
        '</div>' +
        '<div class="p-row" style="margin-top:8px">' +
        '<button data-f="snap" title="Earliest slot after dependencies with free capacity">Snap earliest</button>' +
        '<button data-f="unschedule">Unschedule</button>' +
        '</div>';
    } else {
      scheduleInfo =
        '<div><label class="p-lab">Duration (weeks)</label>' +
        '<input type="number" data-f="durWeeks" min="0.2" step="0.2" value="' +
        (it.durDays != null ? Math.round(it.durDays / SPW() * 100) / 100 : '') + '" placeholder="1" style="width:100%"' +
        ' title="Used when the item lands on the timeline (empty = 1 week, or the size)"></div>' +
        '<div class="p-row" style="margin-top:8px">' +
        '<button data-f="schedule-now" class="primary">Place on timeline</button>' +
        '<button data-f="snap">Snap earliest</button>' +
        '</div>';
    }
    // hard deadline rides with every schedule variant
    scheduleInfo += '<div style="margin-top:8px"><label class="p-lab">Deadline' +
      (RM.pastDeadline(meta, it) ? ' <span class="p-dl-late">— runs past it</span>' : '') + '</label>' +
      '<input type="text" readonly class="cal-in" data-f="deadline" data-cal-clear="Remove deadline" value="' +
      (it.deadline || '') + '" placeholder="None" style="width:100%"></div>';

    // assessment block follows the project's scheme: hidden when off,
    // read-only when auto-computed
    var riskInfo = '';
    if (RM.riskSchemeOf(state) === 'auto') {
      var rkP = RM.depRisk(state, it);
      riskInfo = '<label class="p-lab" style="margin-top:10px">Risk (auto)</label>' +
        '<div class="m-hint" style="margin-top:2px">' +
        (rkP.level === 'none' ? 'No dependency risk detected.'
          : esc(rkP.level.toUpperCase()) + ' \u2014 ' + esc(rkP.reasons.join('; '))) + '</div>';
    } else if (RM.riskEnabled(state)) {
      riskInfo = '<label class="p-lab" style="margin-top:10px">' + esc(RM.riskColLabel(state)) + '</label>' +
        '<div class="seg">' + riskBtns + '</div>';
    }

    var depChips = it.deps.map(function (n) {
      var d = RM.itemByNum(state, n);
      if (!d) return '<span class="dep-chip unknown" title="No item #' + n + '"><i>#' + n + '</i> missing<button class="x" data-deprm="' + n + '"><i data-lucide="x"></i></button></span>';
      return '<span class="dep-chip" data-depgo="' + d.id + '" title="' + esc(d.feature) + '"><i>#' + n + '</i> ' +
        esc(shorten(d.feature, 26)) + '<button class="x" data-deprm="' + n + '"><i data-lucide="x"></i></button></span>';
    }).join('');
    var depTextChips = (it.depsText || []).map(function (t, i) {
      return '<span class="dep-chip text" title="Free-text dependency">' + esc(shorten(t, 30)) +
        '<button class="x" data-deptxtrm="' + i + '"><i data-lucide="x"></i></button></span>';
    }).join('');

    // items that list THIS one as a dependency
    var dependentChips = state.items
      .filter(function (o) { return o.deps.indexOf(it.num) !== -1; })
      .map(function (o) {
        return '<span class="dep-chip" data-depgo="' + o.id + '" title="' + esc(o.feature) + '"><i>#' + o.num + '</i> ' +
          esc(shorten(o.feature || '(untitled)', 26)) +
          '<button class="x" data-rdep="' + o.id + '" title="Remove this link"><i data-lucide="x"></i></button></span>';
      }).join('');

    var storyRows = it.stories.map(function (st) {
      var hasBody = !!(st.description || st.ac);
      return '<div class="p-story" data-pst="' + st.id + '">' +
        '<span class="r-num st-num">#' + st.num + '</span>' +
        '<input type="text" data-pst-title="' + st.id + '" value="' + esc(st.title) + '">' +
        '<button class="st-del st-edit' + (hasBody ? ' has-body' : '') + '" style="opacity:1" data-pst-edit="' + st.id +
        '" title="Description &amp; acceptance criteria"><i data-lucide="pencil"></i></button>' +
        '<button class="st-del" style="opacity:1" data-pst-del="' + st.id + '"><i data-lucide="x"></i></button></div>';
    }).join('');

    // collapsible section card: the body always renders (collapsed hides via CSS)
    var SEC_ICONS = {
      fields: 'text', details: 'info', schedule: 'calendar-range',
      people: 'users', deps: 'git-merge', stories: 'list-todo',
      timeline: 'chart-gantt', checks: 'shield-check',
      meta: 'tags', danger: 'trash-2', integrations: 'plug', tags: 'tag'
    };
    function sec(key, label, summary, body) {
      return '<div class="p-sec c' + (secOpen(key) ? ' open' : '') + '" data-sec="' + key + '">' +
        '<button class="p-sechead" data-sectoggle="' + key + '">' +
        '<i data-lucide="chevron-right"></i>' +
        (SEC_ICONS[key] ? '<i data-lucide="' + SEC_ICONS[key] + '" class="p-secico"></i>' : '') +
        '<span class="p-seclab">' + label + '</span>' +
        '</button><div class="p-secbody">' + body + '</div></div>';
    }

    // every scope column (Description included) edits as rich text in one
    // always-available Fields section, in the document's column order
    var fieldEds = panelScopeCols('feature').map(function (c) {
      return '<div class="p-fld"><label class="p-lab">' + esc(RM.scopeColLabel(c)) + '</label>' +
        wysHtml('col:' + c.key, RM.scopeValue(it, c.key), '') + '</div>';
    }).join('') || '<div class="p-none">No columns — add one in the Scoping view.</div>';

    panel.innerHTML =
      '<div id="panelRz"></div>' +
      '<div class="p-top"><span class="p-lead"><span class="p-num">#<input class="p-num-edit" data-f="num" value="' + it.num +
      '" style="width:' + (String(it.num).length + 1.6) + 'ch" title="Item #"></span>' +
      typeChipHtml('itype', 'feature', it) + panelFlagHtml(it) +
      (it.milestone ? '<button class="p-mschip" data-act="msstyle" title="Milestone">' +
        MS_STYLE_GLYPHS[RM.msStyleOf(it)] + ' Milestone · ' + msStyleLabel(RM.msStyleOf(it)) + '</button>' : '') +
      '</span><button class="p-close" data-f="collapse" title="Hide panel  ]"><i data-lucide="panel-right-close"></i></button></div>' +
      '<textarea class="p-name" data-f="feature" rows="1" placeholder="' + esc(lvl('feature') + ' name') + '">' + esc(it.feature) + '</textarea>' +

      sec('fields', 'Fields', '', fieldEds) +


      sec('details', 'Details', '',
        '<div class="p-grid2">' +
        '<div><label class="p-lab">Phase</label>' + phaseDd + '</div>' +
        (state.meta.workstreamsEnabled
          ? '<div><label class="p-lab">Workstream</label>' +
            ddButton('ws',
              it.workstream ? esc(it.workstream) : esc(RM.defaultWsName(state)) + ' <small>(default)</small>',
              '#' + RM.colorForWs(state, it.workstream), 'Workstream') + '</div>'
          : '<div></div>') +
        '</div>' +
        '<div style="margin-top:8px"><label class="p-lab">Epic</label>' + epicDd + '</div>') +

      sec('schedule', it.milestone || !RM.sizingEnabled(state) ? 'Schedule' : 'Size &amp; schedule', '',
        (it.milestone || !RM.sizingEnabled(state) ? '' :
          '<label class="p-lab">Size</label>' +
          (RM.sizeRollup(state)
            ? '<div style="margin-bottom:8px">' + sizeBtns + '</div>'
            : '<div class="seg" style="margin-bottom:8px">' + sizeBtns +
              '<button data-f="size" data-v=""' + (!it.size ? ' class="on"' : '') + ' title="No size">—</button></div>')) +
        scheduleInfo +
        riskInfo + priInfo +
        '<div class="p-row" style="margin-top:10px">' +
        '<label class="p-check fixed"><input type="checkbox" data-f="locked"' + (it.locked ? ' checked' : '') + '> Locked</label>' +
        '<label class="p-check fixed" title="Auto timeline and ⚡ leave it where it sits; Place at earliest slot still moves it">' +
        '<input type="checkbox" data-f="noAuto"' + (it.noAuto ? ' checked' : '') + '> Excluded from Auto timeline</label>' +
        '<label class="p-check fixed"><input type="checkbox" data-f="done"' + (it.done ? ' checked' : '') + '> Done</label>' +
        '</div>') +

      sec('people', 'People', '',
        (state.meta.capacityEnabled && RM.planLevel(state) === 'feature' && !it.milestone
          ? '<label class="p-lab">Capacity type</label>' +
            ddButton('icap',
              RM.itemCapType(state, it)
                ? (itemCapInherited(it) ? '<span class="inherited">' + esc(RM.itemCapType(state, it)) + '</span>' : esc(RM.itemCapType(state, it)))
                : '<i>— general —</i>',
              null,
              itemCapInherited(it) ? CAP_INHERIT_TITLE : 'What this feature drains at the feature planning level') +
            (state.meta.capMode !== 'points'
              ? '<label class="p-lab" style="margin-top:10px">Multiplier</label>' +
                '<input type="number" min="0.1" step="0.5" data-f="capMult" value="' + (it.capMult || 1) + '" style="width:80px" title="People this feature needs at once">'
              : '') +
            '<label class="p-lab" style="margin-top:10px">Role</label>'
          : '<label class="p-lab">Role</label>') + typeDd +
        '<label class="p-lab" style="margin-top:10px">Assignees</label>' +
        '<div class="chips">' +
        (it.assignees || []).map(function (aid) {
          var mm = memberById(aid);
          if (!mm) return '';
          return '<span class="dep-chip asg-chip">' + avatarHtml(mm, 'sm') + ' ' + esc(mLabel(mm)) +
            '<button class="x" data-asgrm="' + aid + '"><i data-lucide="x"></i></button></span>';
        }).join('') +
        '</div>' +
        (state.team.length
          ? ddButton('assign', '+ Assign\u2026', null, 'Assign people from the roster')
          : '<div class="m-hint">Add people in the Resources panel to assign them.</div>')) +

      sec('tags', 'Tags', (it.tags || []).length ? String(it.tags.length) : '', tagsBody(it)) +
      sec('deps', 'Dependencies', '',
        '<label class="p-lab">Depends on</label>' +
        '<div class="chips">' + depChips + depTextChips + (depChips || depTextChips ? '' : '<span class="p-none">none</span>') + '</div>' +
        '<div class="dep-search"><input data-f="depsearch" placeholder="+ add dependency — search by name…" autocomplete="off" style="width:100%;margin-top:7px">' +
        '<div class="dep-sug" hidden></div></div>' +
        '<label class="p-lab" style="margin-top:10px">Dependents (rely on this)</label>' +
        '<div class="chips">' + dependentChips + (dependentChips ? '' : '<span class="p-none">none</span>') + '</div>' +
        '<div class="m-hint">Tip: hover a bar and drag its edge circles to another bar to link.</div>') +

      (it.milestone ? '' : // milestones carry no stories
        sec('stories', esc(lvl('story', true)), '',
          '<div class="p-stories">' + storyRows + '</div>' +
          ((it.stories || []).length ? '' : // with stories present, the story context menu inserts
            '<input data-f="storyadd" placeholder="' + esc('+ add ' + lvl('story').toLowerCase() + '…') + '" style="width:100%;margin-top:6px">'))) +

      (vlist.length
        ? sec('checks', 'Checks', '',
          '<div class="p-warnlist">' + vlist.map(function (v) {
            var cls = v.level === 'error' ? 'err' : v.level;
            return '<div class="p-warnitem ' + cls + '">' + esc(v.msg) + '</div>';
          }).join('') + '</div>')
        : '') +

      sec('integrations', 'Integrations', '',
        '<label class="p-lab">Jira key</label>' +
        '<input data-f="jiraKey" placeholder="e.g. HW-12" value="' + esc(it.jiraKey || '') +
        '" style="width:100%" title="Jira issue key">') +

      '<datalist id="wsList"><option>Product</option><option>Data</option><option>Process</option><option>Product / Process</option><option>All</option></datalist>';
    if (window.lucide) lucide.createIcons();
    var pn = $('.p-name', panel);
    if (pn) { pn.style.height = 'auto'; pn.style.height = pn.scrollHeight + 'px'; }
  }

  // story detail panel: same scope columns as items; workstream/epic roll up
  function renderStoryPanel(panel, it, st) {
    var epIco = RM.iconForEpic(state, it.epic);
    var fieldEds = panelScopeCols('story').map(function (c) {
      var v = RM.storyScopeValue(st, c.key);
      return '<div class="p-fld"><label class="p-lab">' + esc(RM.scopeColLabel(c)) + '</label>' +
        wysHtml('stcol:' + c.key, v, '') + '</div>';
    }).join('');

    var timeline;
    if (st.startDay != null) {
      timeline =
        '<div class="p-grid2">' +
        '<div><label class="p-lab">Start</label>' +
        '<input type="text" readonly class="cal-in" data-stf="startDate" value="' + RM.fmtISO(RM.dayToDate(state.meta, st.startDay)) + '" style="width:100%"></div>' +
        '<div><label class="p-lab">Weeks</label>' +
        '<input type="number" data-stf="durWeeks" min="0.2" step="0.2" value="' + (Math.round(st.durDays / SPW() * 100) / 100) + '" style="width:100%"></div>' +
        '</div>' +
        '<div class="p-row" style="margin-top:8px"><button data-stf="untimeline">Remove timeline</button></div>';
    } else {
      timeline = '<div class="m-hint">No timeline — double-click the story’s lane on the Planning tab to add one.</div>';
    }
    // estimate: the story's own size / priority / risk scales (Setup → Sizing, priority & risk)
    var stSizeBtns = RM.sizingEnabled(state, 'story')
      ? '<label class="p-lab">Size</label><div class="seg">' +
        ['<button data-stf="size" data-v=""' + (!st.size ? ' class="on"' : '') + ' title="Not set">—</button>']
          .concat(RM.sizeOrderOf(state, 'story').map(function (sz) {
            return '<button data-stf="size" data-v="' + esc(sz) + '"' + (st.size === sz ? ' class="on"' : '') +
              ' title="' + fmtDays(RM.sizeDays(state, sz, 'story')) + '">' + esc(sz) + '</button>';
          })).join('') + '</div>'
      : '';
    var stPriLevels = RM.prioritySchemeOf(state, 'story') === 'levels';
    var stPriBtns = RM.priorityEnabled(state, 'story')
      ? '<label class="p-lab" style="margin-top:10px">Priority</label><div class="seg">' +
        ['<button data-stf="priority" data-v=""' + (!st.priority ? ' class="on"' : '') + ' title="Not set">None</button>']
          .concat(RM.priorityOrderOf(state, 'story').map(function (pv) {
            return '<button data-stf="priority" data-v="' + pv + '" class="' + (st.priority === pv ? 'on' : '') + priTierClass(pv, 'story') + '"' +
              ' title="' + esc(priorityValueLabel(pv, 'story')) + '">' + (stPriLevels ? levelGlyph(pv) : pv) + '</button>';
          })).join('') + '</div>'
      : '';
    var stRiskBtns = RM.riskEnabled(state, 'story')
      ? '<label class="p-lab" style="margin-top:10px">' + esc(RM.riskColLabel(state, 'story')) + '</label><div class="seg">' +
        ['<button data-stf="risk" data-v=""' + (!st.risk ? ' class="on"' : '') + ' title="Not set">None</button>']
          .concat(RM.riskOrderOf(state, 'story').map(function (rv) {
            return '<button data-stf="risk" data-v="' + rv + '"' + (st.risk === rv ? ' class="on"' : '') +
              ' title="' + esc(riskValueLabel(rv, 'story')) + '">' + levelGlyph(rv) + '</button>';
          })).join('') + '</div>'
      : '';
    var estimate = stSizeBtns + stPriBtns + stRiskBtns;

    // dependencies: story → story only (features link separately)
    var rsP = RM.resolveStoryDeps(state, st);
    var stDepChips = rsP.deps.map(function (ref) {
      return '<span class="dep-chip" data-stdepgo="' + ref.st.id + '" title="' + esc(ref.st.title || '') + '"><i>#' + ref.st.num + '</i> ' +
        esc(shorten(ref.st.title || '(untitled)', 26)) + '<button class="x" data-stdeprm="' + ref.st.num + '"><i data-lucide="x"></i></button></span>';
    }).concat(rsP.unknown.map(function (n) {
      return '<span class="dep-chip unknown" title="No ' + esc(lvl('story').toLowerCase()) + ' #' + n + '"><i>#' + n + '</i> missing' +
        '<button class="x" data-stdeprm="' + n + '"><i data-lucide="x"></i></button></span>';
    })).join('');
    var stDependentChips = RM.storyDependents(state, st).map(function (ref) {
      return '<span class="dep-chip" data-stdepgo="' + ref.st.id + '" title="' + esc(ref.st.title || '') + '"><i>#' + ref.st.num + '</i> ' +
        esc(shorten(ref.st.title || '(untitled)', 26)) +
        '<button class="x" data-strdep="' + ref.st.id + '" title="Remove this link"><i data-lucide="x"></i></button></span>';
    }).join('');
    var stDeps =
      '<label class="p-lab">Depends on</label>' +
      '<div class="chips">' + stDepChips + (stDepChips ? '' : '<span class="p-none">none</span>') + '</div>' +
      '<div class="dep-search"><input data-stf="stdepsearch" placeholder="' +
      esc('Add a ' + lvl('story').toLowerCase() + ' by # or name\u2026') + '" autocomplete="off" style="width:100%;margin-top:7px">' +
      '<div class="dep-sug" hidden></div></div>' +
      '<label class="p-lab" style="margin-top:10px">Depended on by</label>' +
      '<div class="chips">' + stDependentChips + (stDependentChips ? '' : '<span class="p-none">none</span>') + '</div>';

    panel.innerHTML =
      '<div id="panelRz"></div>' +
      '<div class="p-top">' +
      '<button class="p-crumb" data-stf="up" title="Back to #' + it.num + '">' +
      '<i data-lucide="corner-left-up"></i>#' + it.num + ' ' + esc(shorten(it.feature || '(untitled)', 26)) + '</button>' +
      '<span class="p-lead"><span class="p-num">#<input class="p-num-edit" data-stf="num" value="' + st.num +
      '" style="width:' + (String(st.num).length + 1.6) + 'ch" title="' + esc(lvl('story')) + ' #"></span></span>' +
      typeChipHtml('stype', 'story', st) + panelFlagHtml(st) +
      '<button class="p-close" data-f="collapse" title="Hide panel  ]"><i data-lucide="panel-right-close"></i></button></div>' +
      '<textarea class="p-name" data-stf="title" rows="1" placeholder="' + esc(lvl('story') + ' title') + '">' + esc(st.title) + '</textarea>' +
      '<div class="p-row" style="margin:6px 0 8px">' +
      '<label class="p-check fixed"><input type="checkbox" data-stf="done"' + (st.done ? ' checked' : '') + '> Done</label>' +
      '<label class="p-check fixed" title="Auto timeline and ⚡ leave it where it sits; Place at earliest slot still moves it">' +
      '<input type="checkbox" data-stf="noAuto"' + (st.noAuto ? ' checked' : '') + '> Excluded from Auto timeline</label>' +
      '</div>' +

      '<div class="p-sec c open"><button class="p-sechead" tabindex="-1">' +
      '<i data-lucide="chevron-right"></i><span class="p-seclab">Rolls up to</span></button>' +
      '<div class="p-secbody"><div class="p-rollup">' +
      '<span class="sc-chip"><span class="dd-dot" style="background:#' + RM.colorForWs(state, it.workstream) + '"></span>' +
      (it.workstream ? esc(it.workstream) : '<i>no workstream</i>') + '</span>' +
      '<span class="sc-chip">' + (epIco ? '<i data-lucide="' + epIco + '"></i>' : '') +
      (it.epic ? esc(it.epic) : '<i>no epic</i>') + '</span>' +
      '</div><div class="m-hint">Stories inherit workstream and epic from their feature.</div></div></div>' +

      sec2('fields', 'Fields', fieldEds) +
      (estimate ? sec2('estimate', 'Estimate', estimate) : '') +
      sec2('people', 'People',
        '<label class="p-lab">Capacity type</label>' +
        ddButton('stcap', st.capType ? esc(st.capType) : '<i>\u2014 general \u2014</i>', null, 'What this story drains and who can take it') +
        (state.meta.capacityEnabled && state.meta.capMode !== 'points'
          ? '<label class="p-lab" style="margin-top:10px">Multiplier</label>' +
            '<input type="number" min="0.1" step="0.5" data-stf="capMult" value="' + (st.capMult || 1) + '" style="width:80px" title="People this story needs at once">'
          : '') +
        '<label class="p-lab" style="margin-top:10px">Assignees</label>' +
        '<div class="chips">' +
        (st.assignees || []).map(function (aid) {
          var mm = memberById(aid);
          if (!mm) return '';
          return '<span class="dep-chip asg-chip">' + avatarHtml(mm, 'sm') + ' ' + esc(mLabel(mm)) +
            '<button class="x" data-stasgrm="' + aid + '"><i data-lucide="x"></i></button></span>';
        }).join('') +
        '</div>' +
        (state.team.length
          ? ddButton('stassign', '+ Assign\u2026', null, 'Assign people from the roster')
          : '<div class="m-hint">Add people in the Resources panel to assign them.</div>')) +
      sec2('tags', 'Tags', tagsBody(st)) +
      sec2('schedule', 'Timeline', timeline) +
      sec2('deps', 'Dependencies', stDeps) +
      sec2('integrations', 'Integrations',
        '<label class="p-lab">Jira key</label>' +
        '<input data-stf="jiraKey" placeholder="e.g. HW-13" value="' + esc(st.jiraKey || '') +
        '" style="width:100%" title="Jira issue key">');
    if (window.lucide) lucide.createIcons();
    var pn = $('.p-name', panel);
    if (pn) { pn.style.height = 'auto'; pn.style.height = pn.scrollHeight + 'px'; }
  }
  // always-open section (story panel keeps things simple)
  function sec2(key, label, body) {
    return '<div class="p-sec c open" data-sec="st-' + key + '">' +
      '<button class="p-sechead" tabindex="-1">' +
      '<i data-lucide="chevron-right"></i><span class="p-seclab">' + label + '</span>' +
      '</button><div class="p-secbody">' + body + '</div></div>';
  }

  // ---- tags: free-form labels, edited only here (rows and cards stay lean).
  // The datalist offers every other tag in the document.
  function tagsBody(o) {
    var tags = (o && o.tags) || [];
    var own = {};
    tags.forEach(function (t) { own[String(t).toLowerCase()] = true; });
    var opts = RM.allTags(state).filter(function (t) { return !own[String(t).toLowerCase()]; });
    return '<div class="chips p-tags">' +
      tags.map(function (t) {
        return '<span class="tag-chip">' + esc(t) +
          '<button data-tagrm="' + esc(t) + '" title="Remove tag"><i data-lucide="x"></i></button></span>';
      }).join('') +
      '<input class="p-tag-in" data-tagadd="1" list="tagOptions" autocomplete="off" placeholder="Add tag\u2026">' +
      '</div>' +
      '<datalist id="tagOptions">' +
      opts.map(function (t) { return '<option value="' + esc(t) + '">'; }).join('') +
      '</datalist>';
  }
  // the object the panel is editing: the open story, else the feature
  function tagTarget() {
    var it = selectedId && RM.itemById(state, selectedId);
    if (!it) return null;
    return selStory ? storyById(it, selStory) : it;
  }
  function tagsOfTarget() {
    var t = tagTarget();
    return (t && t.tags) || [];
  }
  function commitTags(list) {
    var itId = selectedId, stId = selStory;
    commit('tags', function (s) {
      var x = RM.itemById(s, itId);
      if (!x) return;
      var target = stId ? storyById(x, stId) : x;
      if (target) RM.setTags(s, target, list);
    });
  }
  // the panel re-renders on every commit — put the caret back in the input
  function focusTagInput() {
    requestAnimationFrame(function () {
      var el = $('#panel .p-tag-in');
      if (el) el.focus();
    });
  }
  // returns true when something was actually added (the panel then re-rendered)
  function addTagFromInput(inp) {
    var v = (inp.value || '').trim();
    inp.value = '';
    if (!v || !tagTarget()) return false;
    var cur = tagsOfTarget();
    var next = cur.concat([v]);
    if (RM.normalizeTags(next).length === cur.length) return false; // already there
    commitTags(next);
    return true;
  }

  // ---- shared dropdown: a button that opens the same list UI the menu bar
  // uses (icons/dots, blue selected row, check on the right, optional edit)
  function ddButton(dd, label, dot, title) {
    return '<button class="dd-btn" data-dd="' + dd + '" title="' + esc(title || '') + '">' +
      (dot ? '<span class="dd-dot" style="background:' + dot + '"></span>' : '') +
      '<span class="dd-label">' + label + '</span><i data-lucide="chevron-down"></i></button>';
  }
  // one menu renderer for dropdowns and context menus. `m.pre` is raw HTML
  // drawn before the label (an avatar, a type glyph); `opts.search` adds a
  // filter box at the top (a combobox: typing narrows the rows, Enter picks
  // the first match, arrows walk the visible rows)
  function menuListHtml(items, minW, opts) {
    var search = opts && opts.search
      ? '<input class="menu-search" type="text" placeholder="' + esc(opts.search) + '" spellcheck="false" autocomplete="off">' : '';
    return '<div class="menu-list' + (search ? ' has-search' : '') + '"' + (minW ? ' style="min-width:' + minW + 'px"' : '') + '>' +
      search + items.map(function (m, i) {
        if (m.sep) return '<div class="menu-sep"></div>';
        return '<button data-mi="' + i + '"' + (m.disabled ? ' disabled' : '') + (m.title ? ' title="' + esc(m.title) + '"' : '') + (m.checked ? ' class="on"' : '') + '>' +
          (m.dot ? '<span class="dd-dot" style="background:' + m.dot + '"></span>' : '') +
          (m.pre || '') +
          (m.icon ? '<i data-lucide="' + esc(m.icon) + '"></i>' : '') +
          '<span>' + m.label + '</span>' +
          (m.actions ? '<span class="mi-acts">' + m.actions.map(function (a, j) {
            return '<span class="mi-act' + (a.on ? ' on' : '') + '" data-ma="' + i + ':' + j +
              '" title="' + a.title + '"><i data-lucide="' + a.icon + '"></i></span>';
          }).join('') + '</span>' : '')
          +
          (m.edit ? '<span class="mi-edit" data-me="' + i + '" title="Edit"><i data-lucide="pencil"></i></span>' : '') +
          (m.checked && !m.actions ? '<i data-lucide="check" class="mi-check"></i>' : '') +
          '</button>';
      }).join('') + '</div>';
  }
  // keyboard + search wiring shared by both menu kinds
  function wireMenu(host, items) {
    var btns = $$('.menu-list [data-mi]', host);
    var visible = function () { return btns.filter(function (b) { return !b.hidden; }); };
    var inp = host.querySelector('.menu-search');
    if (inp) {
      inp.addEventListener('input', function () {
        var q = inp.value.trim().toLowerCase();
        btns.forEach(function (b) { b.hidden = !!q && b.textContent.toLowerCase().indexOf(q) === -1; });
        $$('.menu-sep', host).forEach(function (sp) { sp.hidden = !!q; });
      });
      inp.addEventListener('keydown', function (ev) {
        if (ev.key === 'Enter') {
          ev.preventDefault();
          var first = visible()[0];
          if (first) first.click();
        } else if (ev.key === 'ArrowDown') {
          ev.preventDefault();
          var f = visible()[0];
          if (f) f.focus({ preventScroll: true });
        } else if (ev.key === 'Escape') {
          ev.stopPropagation();
          closePopover();
        }
      });
      inp.focus({ preventScroll: true });
    } else {
      // focus lands on the current choice
      var start = host.querySelector('.menu-list button.on') || btns[0];
      if (start) start.focus({ preventScroll: true });
    }
    host.addEventListener('keydown', function (ev) {
      if (ev.key !== 'ArrowDown' && ev.key !== 'ArrowUp') return;
      if (ev.target === inp) return;
      ev.preventDefault();
      var vis = visible();
      var i2 = vis.indexOf(document.activeElement);
      var n = ev.key === 'ArrowDown' ? i2 + 1 : i2 - 1;
      if (n < 0) { if (inp) { inp.focus(); return; } n = vis.length - 1; }
      if (n >= vis.length) n = 0;
      if (vis[n]) vis[n].focus({ preventScroll: true });
    });
  }
  function openDropdown(anchor, items, ddOpts) {
    var r = anchor.getBoundingClientRect();
    // remember which chip opened this so focus can return after the re-render
    var restoreSel = null;
    var aRow = anchor.closest && anchor.closest('.row');
    if (aRow && aRow.dataset.id && anchor.dataset.act) {
      restoreSel = '.row[data-id="' + aRow.dataset.id + '"] [data-act="' + anchor.dataset.act + '"]';
    }
    function restoreFocus() {
      if (!restoreSel) return;
      requestAnimationFrame(function () {
        var el = rowsEl.querySelector(restoreSel);
        if (el && el.focus) el.focus({ preventScroll: true });
      });
    }
    var html = menuListHtml(items, Math.max((ddOpts && ddOpts.minW) || 210, Math.round(r.width)), ddOpts);
    openPopover(r.left, r.bottom + 4, html, function (host) {
      if (window.lucide) lucide.createIcons();
      wireMenu(host, items);
      host.addEventListener('click', function (ev) {
        var ma = ev.target.closest('[data-ma]');
        if (ma) {
          ev.stopPropagation();
          closePopover();
          var pq = ma.dataset.ma.split(':');
          items[+pq[0]].actions[+pq[1]].fn();
          restoreFocus();
          return;
        }
        var me = ev.target.closest('[data-me]');
        if (me) {
          ev.stopPropagation();
          closePopover();
          items[parseInt(me.dataset.me, 10)].edit();
          return;
        }
        var mi = ev.target.closest('[data-mi]');
        if (!mi) return;
        closePopover();
        items[parseInt(mi.dataset.mi, 10)].fn();
        restoreFocus();
      });
    });
  }

  function allWorkstreams(sArg) {
    var st = sArg || state;
    var seen = {}, ref = [];
    st.items.map(function (x) { return x.workstream; })
      .concat(st.team.reduce(function (a, x) { return a.concat(RM.memberWorkstreams(x)); }, []))
      .concat(Object.keys(st.wsColors))
      .forEach(function (w) { if (w && !seen[w]) { seen[w] = true; ref.push(w); } });
    var out = [];
    (st.wsOrder || []).forEach(function (w) { if (seen[w] && out.indexOf(w) === -1) out.push(w); });
    ref.forEach(function (w) { if (out.indexOf(w) === -1) out.push(w); });
    return out;
  }

  function allEpics() {
    var seen = {}, out = [];
    state.items.forEach(function (it) {
      if (it.epic && !seen[it.epic]) { seen[it.epic] = true; out.push(it.epic); }
    });
    (state.epicList || []).forEach(function (ep) { if (!seen[ep]) { seen[ep] = true; out.push(ep); } });
    out.sort(function (a, b) { return a.localeCompare(b); });
    return out;
  }

  // edit an epic's label + icon (applies to every item carrying it)
  var EPIC_ICONS = ['tag', 'star', 'flag', 'rocket', 'target', 'layers', 'database', 'shield',
    'zap', 'globe', 'users', 'wrench', 'chart-line', 'box', 'lightbulb', 'compass',
    'cpu', 'plug', 'bot', 'flask-conical', 'map', 'workflow', 'network', 'building',
    'bug', 'check-square', 'list-tree', 'rows-3', 'corner-down-right', 'square', 'bookmark'];
  function epicEditModal(epicName) {
    var setting = state.epicIcons[epicName] || null;
    var count = state.items.filter(function (x) { return x.epic === epicName; }).length;
    var icons = '<button class="iswatch' + (!setting ? ' on' : '') + '" data-eic="" title="No icon">—</button>' +
      EPIC_ICONS.map(function (k) {
        return '<button class="iswatch' + (setting === k ? ' on' : '') + '" data-eic="' + k + '" title="' + k + '">' +
          '<i data-lucide="' + k + '"></i></button>';
      }).join('');
    openModal(
      '<div class="modal" style="width:420px">' +
      '<div class="m-head"><h2>Edit epic</h2><button class="p-close" data-m="x"><i data-lucide="x"></i></button></div>' +
      '<div class="m-body">' +
      '<div class="m-sec"><label>Label</label><input id="epName" style="width:100%" value="' + esc(epicName) + '">' +
      '<div class="m-hint">Renames the epic on all ' + count + ' item(s) that carry it.</div></div>' +
      '<div class="m-sec"><label>Icon</label><div class="iswatches">' + icons + '</div></div>' +
      '<div class="m-sec"><label>Type</label><select id="epType" style="width:100%">' +
      (function () {
        var curKey = RM.typeOf(state, epicName, 'epic').key;
        var list = RM.typesFor(state, 'epic').slice();
        if (!list.some(function (t) { return t.key === curKey; })) {
          var cur = RM.itemType(state, curKey);
          if (cur) list.push(cur);
        }
        return list.map(function (t) {
          return '<option value="' + esc(t.key) + '"' + (t.key === curKey ? ' selected' : '') + '>' + esc(t.label) + '</option>';
        }).join('');
      })() + '</select></div>' +
      '<div class="m-sec"><label>Jira key</label><input id="epJira" style="width:100%" placeholder="e.g. HW-1" value="' + esc(state.epicJira[epicName] || '') + '">' +
      '<div class="m-hint">The Jira epic these features parent to in the Jira CSV export.</div></div>' +
      '</div>' +
      '<div class="m-foot"><button data-m="x2">Cancel</button><button id="epSave" class="primary">Save</button></div></div>',
      function (host) {
        var picked = setting; // icon name or null
        $('[data-m=x]', host).onclick = closeModal;
        $('[data-m=x2]', host).onclick = closeModal;
        host.addEventListener('click', function (ev) {
          var b = ev.target.closest('[data-eic]');
          if (!b) return;
          picked = b.dataset.eic || null;
          $$('.iswatch', host).forEach(function (s) { s.classList.toggle('on', s === b); });
        });
        $('#epSave', host).onclick = function () {
          var newName = $('#epName', host).value.trim();
          var jira = RM.jiraKeyOf($('#epJira', host).value);
          var typeKey = $('#epType', host).value;
          closeModal();
          commit('edit epic', function (s) {
            var name2 = newName || epicName;
            if (name2 !== epicName) {
              s.items.forEach(function (x) { if (x.epic === epicName) x.epic = name2; });
              s.epicList = (s.epicList || []).map(function (e) { return e === epicName ? name2 : e; })
                .filter(function (e, i, a) { return a.indexOf(e) === i; });
              if (s.epicIcons[epicName] != null) { s.epicIcons[name2] = s.epicIcons[epicName]; delete s.epicIcons[epicName]; }
              if (s.epicTypes[epicName] != null) { s.epicTypes[name2] = s.epicTypes[epicName]; delete s.epicTypes[epicName]; }
              delete s.epicJira[epicName];
            }
            if (picked) s.epicIcons[name2] = picked;
            else delete s.epicIcons[name2];
            applyEpicType(s, name2, typeKey);
            if (jira) s.epicJira[name2] = jira;
            else delete s.epicJira[name2];
          });
        };
      });
  }

  // edit the DEFAULT workstream (what a null workstream means): name + color
  function defaultWsModal() {
    var cur = '#' + RM.defaultWsColor(state);
    var setting = state.meta.defaultWsColor || null;
    var count = state.items.filter(function (x) { return !x.workstream; }).length;
    var sw = ['neutral'].concat(RM.PALETTE_KEYS).map(function (k) {
      var on = ('' + setting).toUpperCase() === RM.PALETTE[k] ||
        (k === 'neutral' && !setting);
      return '<button class="swatch' + (on ? ' on' : '') +
        '" data-esw="' + RM.PALETTE[k] + '" style="background:#' + RM.PALETTE[k] +
        '" title="' + (k === 'neutral' ? 'gray' : k) + '"></button>';
    }).join('');
    openModal(
      '<div class="modal" style="width:420px">' +
      '<div class="m-head"><h2>Default workstream</h2><button class="p-close" data-m="x"><i data-lucide="x"></i></button></div>' +
      '<div class="m-body">' +
      '<div class="m-sec"><label>Name</label><input id="dwsName" style="width:100%" value="' + esc(RM.defaultWsName(state)) + '">' +
      '<div class="m-hint">Items without a workstream belong here (' + count + ' item(s) right now).</div></div>' +
      '<div class="m-sec"><label>Color</label><div class="swatches">' + sw +
      '<input type="color" id="dwsColor" value="' + cur + '" title="Custom color"></div></div>' +
      '</div>' +
      '<div class="m-foot"><button data-m="x2">Cancel</button><button id="dwsSave" class="primary">Save</button></div></div>',
      function (host) {
        var picked = null;
        $('[data-m=x]', host).onclick = closeModal;
        $('[data-m=x2]', host).onclick = closeModal;
        host.addEventListener('click', function (ev) {
          var b = ev.target.closest('[data-esw]');
          if (!b) return;
          picked = b.dataset.esw;
          $$('.swatch', host).forEach(function (s) { s.classList.toggle('on', s === b); });
        });
        $('#dwsColor', host).addEventListener('change', function (ev) {
          picked = String(ev.target.value).replace(/^#/, '').toUpperCase();
          $$('.swatch', host).forEach(function (s) { s.classList.remove('on'); });
        });
        $('#dwsSave', host).onclick = function () {
          var nv = $('#dwsName', host).value.trim();
          closeModal();
          commit('default workstream', function (s) {
            if (nv) s.meta.defaultWsName = nv;
            if (picked) s.meta.defaultWsColor = picked;
          });
        };
      });
  }

  // edit a workstream's label + color (applies to items and roles carrying it)
  function wsEditModal(wsName) {
    var setting = state.wsColors[wsName] || null;
    var cur = '#' + RM.colorForWs(state, wsName);
    var count = state.items.filter(function (x) { return x.workstream === wsName; }).length;
    var sw = '<button class="swatch' + (!setting || setting === 'neutral' ? ' on' : '') + '" data-esw="neutral" style="background:#' + RM.PALETTE.neutral + '" title="Default gray"></button>' +
      RM.PALETTE_KEYS.map(function (k) {
        return '<button class="swatch' + (setting === k ? ' on' : '') + '" data-esw="' + k + '" style="background:#' + RM.PALETTE[k] + '" title="' + k + '"></button>';
      }).join('');
    openModal(
      '<div class="modal" style="width:420px">' +
      '<div class="m-head"><h2>Edit workstream</h2><button class="p-close" data-m="x"><i data-lucide="x"></i></button></div>' +
      '<div class="m-body">' +
      '<div class="m-sec"><label>Label</label><input id="wsName" style="width:100%" value="' + esc(wsName) + '">' +
      '<div class="m-hint">Renames the workstream on all ' + count + ' item(s) and any roles that carry it.</div></div>' +
      '<div class="m-sec"><label>Color</label><div class="swatches">' + sw +
      '<input type="color" id="wsColor" value="' + cur + '" title="Custom color"></div></div>' +
      '</div>' +
      '<div class="m-foot">' +
      '<button id="wsDelete" class="danger" style="margin-right:auto">Delete workstream</button>' +
      '<button data-m="x2">Cancel</button><button id="wsSave" class="primary">Save</button></div></div>',
      function (host) {
        var picked = setting; // palette key, hex, or null (default gray)
        $('[data-m=x]', host).onclick = closeModal;
        $('[data-m=x2]', host).onclick = closeModal;
        host.addEventListener('click', function (ev) {
          var b = ev.target.closest('[data-esw]');
          if (!b) return;
          picked = b.dataset.esw || null;
          $$('.swatch', host).forEach(function (s) { s.classList.toggle('on', s === b); });
        });
        $('#wsColor', host).addEventListener('change', function (ev) {
          picked = String(ev.target.value).replace(/^#/, '').toUpperCase();
          $$('.swatch', host).forEach(function (s) { s.classList.remove('on'); });
        });
        $('#wsDelete', host).onclick = function () {
          closeModal();
          confirmBox('Delete workstream “' + esc(wsName) + '”?',
            'Its ' + count + ' item(s) and any roles keep their place with no workstream.',
            'Delete', function () {
              commit('delete workstream', function (s) {
                s.items.forEach(function (x) { if (x.workstream === wsName) x.workstream = ''; });
                s.team.forEach(function (m) {
                  RM.setMemberWorkstreams(m, RM.memberWorkstreams(m).filter(function (w) { return w !== wsName; }));
                });
                delete s.wsColors[wsName];
              });
            }, true);
        };
        $('#wsSave', host).onclick = function () {
          var newName = $('#wsName', host).value.trim();
          closeModal();
          commit('edit workstream', function (s) {
            var name2 = newName || wsName;
            if (name2 !== wsName) {
              s.items.forEach(function (x) { if (x.workstream === wsName) x.workstream = name2; });
              s.team.forEach(function (m) {
                RM.setMemberWorkstreams(m, RM.memberWorkstreams(m).map(function (w) { return w === wsName ? name2 : w; }));
              });
              if (s.wsColors[wsName] != null) { s.wsColors[name2] = s.wsColors[wsName]; delete s.wsColors[wsName]; }
            }
            // store the choice EXPLICITLY — deleting the entry would let the
            // known-workstream default (e.g. Product = plum) re-seed on the
            // next load and silently revert a "Default gray" pick
            s.wsColors[name2] = picked || 'neutral';
          });
        };
      });
  }

  // inline "new epic" input swapped in place of the epic dropdown button
  function newEpicInline(itemId) {
    var btn = $('#panel [data-dd="epic"]');
    if (!btn) return;
    var inp = document.createElement('input');
    inp.placeholder = 'New epic name…';
    inp.style.width = '100%';
    btn.replaceWith(inp);
    inp.focus();
    var done = false;
    function finish(saveIt) {
      if (done) return; done = true;
      var nv = inp.value.trim();
      if (saveIt && nv) commit('epic', function (s) { RM.itemById(s, itemId).epic = nv; });
      else render();
    }
    inp.addEventListener('blur', function () { finish(true); });
    inp.addEventListener('keydown', function (ev) {
      if (ev.key === 'Enter') finish(true);
      if (ev.key === 'Escape') finish(false);
    });
  }

  function fmtPe(x) { return Math.round(x * 10) / 10; }

  // measured in weeks (5 working days); short spans read better in days
  function fmtDays(d) {
    if (d == null) return '';
    if (d < 5) return (Math.round(d * 10) / 10) + 'd';
    var w = d / SPW();
    return (Math.round(w * 10) / 10) + 'w';
  }
  function shorten(s, n) { return s.length > n ? s.slice(0, n - 1) + '…' : s; }

  // panel events (delegated)
  $('#panelPeek').addEventListener('click', function () {
    panelOpen = true;
    saveLocal();
    renderPanel();
  });
  function togglePanel() {
    panelOpen = !panelOpen;
    saveLocal();
    renderPanel();
  }
  function toggleLeftPane() {
    leftCollapsed = !leftCollapsed;
    saveLocal();
    render();
  }
  $('#leftCollapse').addEventListener('click', toggleLeftPane);
  $('#leftPeek').addEventListener('click', toggleLeftPane);

  $('#panel').addEventListener('click', function (e) {
    // collapse works from every panel state (empty state included)
    var clps = e.target.closest('[data-f="collapse"]');
    if (clps) {
      panelOpen = false;
      saveLocal();
      renderPanel();
      return;
    }
    var it = selectedId && RM.itemById(state, selectedId);
    if (!it) return;
    var tagRm = e.target.closest('[data-tagrm]');
    if (tagRm) {
      var gone = tagRm.dataset.tagrm;
      commitTags(tagsOfTarget().filter(function (t) { return t !== gone; }));
      return;
    }
    var msChip = e.target.closest('[data-act="msstyle"]');
    if (msChip) { openDropdown(msChip, msStyleItems(it.id, it)); return; }
    var tyChip = e.target.closest('[data-act="itype"]');
    if (tyChip) { openDropdown(tyChip, typeMenuItems('feature', it.type, function (k) { setItemType(it.id, k); })); return; }
    var secBtn = e.target.closest('[data-sectoggle]');
    if (secBtn) {
      var sk = secBtn.dataset.sectoggle;
      panelSec[sk] = secOpen(sk); // open → store collapsed=true, and vice versa
      saveLocal();
      renderPanel();
      return;
    }
    var stChip = e.target.closest('[data-act="stype"]');
    if (stChip && selStory) {
      var stForType = storyById(it, selStory);
      if (stForType) openDropdown(stChip, typeMenuItems('story', stForType.type, function (k) { setStoryType(it.id, selStory, k); }));
      return;
    }
    // story estimate buttons (size / priority / risk segs in the story panel)
    var stBtn = e.target.closest('button[data-stf][data-v]');
    if (stBtn && selStory) {
      var stv = stBtn.dataset.v || null, stk = stBtn.dataset.stf;
      if (stk === 'size') setStorySize(it.id, selStory, stv);
      else if (stk === 'priority') withStory('story priority', it.id, selStory, function (st2) { st2.priority = stv; });
      else if (stk === 'risk') setStoryRisk(it.id, selStory, stv);
      return;
    }
    // story dependency chips (before the generic [data-stf] return)
    if (selStory) {
      var stDepGo = e.target.closest('[data-stdepgo]');
      var stDepRm = e.target.closest('[data-stdeprm]');
      var stRDep = e.target.closest('[data-strdep]');
      if (stDepRm) {
        var rmNum = parseInt(stDepRm.dataset.stdeprm, 10);
        var rmStId = selStory;
        commit('remove story dep', function (s) {
          var st2 = storyById(RM.itemById(s, it.id) || {}, rmStId);
          if (st2) st2.deps = (st2.deps || []).filter(function (x) { return x !== rmNum; });
        });
        return;
      }
      if (stRDep) {
        var otherId = stRDep.dataset.strdep;
        var meStId = selStory;
        commit('remove story dependent', function (s) {
          var meSt = storyById(RM.itemById(s, it.id) || {}, meStId);
          var oRef = RM.storyRef(s, otherId);
          if (meSt && oRef) oRef.st.deps = (oRef.st.deps || []).filter(function (x) { return x !== meSt.num; });
        });
        return;
      }
      if (stDepGo) {
        var goRef = RM.storyRef(state, stDepGo.dataset.stdepgo);
        if (goRef) selectStory(goRef.it.id, goRef.st.id);
        return;
      }
    }
    // story-panel controls
    var stf = e.target.closest('[data-stf]');
    if (stf && selStory) {
      var stfk = stf.dataset.stf;
      if (stfk === 'up') { select(it.id); return; }
      if (stfk === 'untimeline') {
        commit('story timeline', function (s) {
          var st2 = storyById(RM.itemById(s, it.id) || {}, selStory);
          if (st2) { st2.startDay = null; st2.durDays = null; }
        });
        return;
      }
      return; // title/done/dates commit on change
    }

    var btn = e.target.closest('[data-f]');
    var dep = e.target.closest('[data-deprm]');
    var depTxt = e.target.closest('[data-deptxtrm]');
    var rdep = e.target.closest('[data-rdep]');
    var depGo = e.target.closest('[data-depgo]');
    var stDel = e.target.closest('[data-pst-del]');
    var stEd = e.target.closest('[data-pst-edit]');
    if (stEd) { selectStory(it.id, stEd.dataset.pstEdit); return; }
    if (rdep) {
      // a dependent stops relying on this item
      var depId = rdep.dataset.rdep;
      commit('remove dependent', function (s) {
        var d2 = RM.itemById(s, depId);
        var me = RM.itemById(s, it.id);
        if (d2 && me) d2.deps = d2.deps.filter(function (n) { return n !== me.num; });
      });
      return;
    }
    if (dep) {
      var n = parseInt(dep.dataset.deprm, 10);
      commit('remove dep', function (s) {
        var t = RM.itemById(s, it.id);
        t.deps = t.deps.filter(function (x) { return x !== n; });
      });
      return;
    }
    if (depTxt) {
      var i = parseInt(depTxt.dataset.deptxtrm, 10);
      commit('remove text dep', function (s) { RM.itemById(s, it.id).depsText.splice(i, 1); });
      return;
    }
    if (depGo) { select(depGo.dataset.depgo, true); return; }
    if (stDel) {
      commit('delete story', function (s) {
        var t = RM.itemById(s, it.id);
        t.stories = t.stories.filter(function (st) { return st.id !== stDel.dataset.pstDel; });
      });
      return;
    }
    var dd = e.target.closest('[data-dd]');
    if (dd) {
      var which = dd.dataset.dd;
      if (which === 'phase') {
        openDropdown(dd, movePhaseMenu(it.id));
        return;
      }
      if (which === 'epic') {
        openDropdown(dd, setEpicMenu(it.id, true));
        return;
      }
      if (which === 'ws') {
        openDropdown(dd, wsMenuItems(it.id, function () { return $('#panel [data-dd="ws"] .dd-label'); }));
        return;
      }
      if (which === 'stassign' && selStory) {
        openDropdown(dd, storyAssignMenuItems(it.id, selStory), ASSIGN_DD);
        return;
      }
      if (which === 'stcap' && selStory) {
        openDropdown(dd, storyCapMenuItems(it.id, selStory));
        return;
      }
      if (which === 'icap') {
        openDropdown(dd, itemCapMenuItems(it.id));
        return;
      }
      if (which === 'assign') {
        openDropdown(dd, assignMenuItems(it.id), ASSIGN_DD);
        return;
      }
      if (which === 'teamType') {
        var tItems = [{ label: 'Any role', checked: !it.teamType, fn: function () {
          commit('team type', function (s) { RM.itemById(s, it.id).teamType = ''; });
        } }];
        state.teamTypes.forEach(function (t) {
          tItems.push({ label: esc(t), checked: it.teamType === t, fn: function () {
            commit('team type', function (s) { RM.itemById(s, it.id).teamType = t; });
          } });
        });
        openDropdown(dd, tItems);
        return;
      }
    }
    var stAsgRm = e.target.closest('[data-stasgrm]');
    if (stAsgRm && selStory) {
      var rmSid = selStory, rmPid = stAsgRm.dataset.stasgrm;
      commit('story assignees', function (s) {
        var st2 = storyById(RM.itemById(s, it.id) || {}, rmSid);
        if (st2) st2.assignees = (st2.assignees || []).filter(function (x) { return x !== rmPid; });
      });
      return;
    }
    var asgrm = e.target.closest('[data-asgrm]');
    if (asgrm) {
      var rmId = asgrm.dataset.asgrm;
      commit('assignees', function (s) {
        var t = RM.itemById(s, it.id);
        t.assignees = (t.assignees || []).filter(function (x) { return x !== rmId; });
      });
      return;
    }
    if (!btn) return;
    var f = btn.dataset.f;
    if (f === 'close') { select(null); return; }
    if (f === 'size') {
      if (sizeLocked(it)) return;
      var v = btn.dataset.v || null;
      commit('size', function (s) {
        var t = RM.itemById(s, it.id);
        t.size = v;
        if (v && isScheduled(t) && !t.locked) {
          t.durDays = RM.stretchSpan(s.meta, t.startDay, RM.sizeDays(s, v));
          t.riskDays = RM.stretchSpan(s.meta, t.startDay + t.durDays, RM.riskEffortDays(s, t));
        }
      });
      return;
    }
    if (f === 'riskSize') {
      var rv = btn.dataset.v || null;
      commit('risk size', function (s) {
        var t = RM.itemById(s, it.id);
        t.risk = rv;
        if (isScheduled(t)) {
          t.riskDays = 0;
        } else if (!rv) {
          t.riskDays = 0;
        }
      });
      return;
    }
    if (f === 'priSet') {
      var pv2 = btn.dataset.v || null;
      commit('priority', function (s) {
        RM.itemById(s, it.id).priority = pv2;
      });
      return;
    }
    if (f === 'snap') {
      // locked and finished rows never move (same rule as the context menu)
      if (it.locked || it.done) {
        toast(it.locked ? 'Locked — unlock it to place it' : 'Already done — nothing to place', 'err');
        return;
      }
      var r = RM.placeUnit(state, it.id, null, snapOpts());
      if (r.changed) {
        if (autoOrder) RM.sortItemsByStart(r.state);
        replaceState('snap', r.state);
        toast('Snapped #' + it.num + ' to its earliest open slot' + (r.note ? ' — ' + r.note : ''));
      } else {
        toast(r.note || ('#' + it.num + ' is already at its earliest slot'));
      }
      return;
    }
    if (f === 'schedule-now') {
      commit('schedule', function (s) {
        var t = RM.itemById(s, it.id);
        var today = RM.dateToDay(s.meta, new Date(Date.UTC(new Date().getFullYear(), new Date().getMonth(), new Date().getDate())));
        t.startDay = Math.max(0, Math.min(today, RM.numDays(s.meta) - 5));
        t.durDays = t.milestone ? 0
          : (t.durDays != null ? t.durDays : RM.stretchSpan(s.meta, t.startDay, RM.effortDays(s, t)));
        t.riskDays = t.milestone ? 0 : RM.stretchSpan(s.meta, t.startDay + t.durDays, RM.riskEffortDays(s, t));
      });
      return;
    }
    if (f === 'unschedule') {
      commit('unschedule', function (s) {
        var t = RM.itemById(s, it.id);
        t.startDay = null; t.durDays = null; t.riskDays = 0;
      });
      return;
    }
  });

  // flip an item between a duration bar and a fixed-date milestone diamond
  function toggleMilestone(itemId) {
    commit('milestone', function (s) {
      var t = RM.itemById(s, itemId);
      if (!t) return;
      t.milestone = !t.milestone;
      if (t.milestone) {
        if (t.durDays != null) t.durDays = 0;
        t.riskDays = 0;
        t.size = null; t.priority = null; // milestones carry neither
      } else if (t.durDays != null) {
        // back to a bar: restore a duration from the size (else one week)
        var days = RM.itemSizeDays(s, t, snapOpts()) || 5;
        t.durDays = RM.stretchSpan(s.meta, t.startDay, days);
      }
    });
  }

  // rich editors commit on blur (contenteditable has no change event):
  //   col:<key>   — an item's scope column
  //   stcol:<key> — a story's scope column (Description / Acceptance criteria
  //                 are story fields; the rest live in st.custom)
  $('#panel').addEventListener('focusout', function (e) {
    var tinOut = e.target.closest && e.target.closest('[data-tagadd]');
    if (tinOut) { addTagFromInput(tinOut); return; }
    var ed = e.target.classList && e.target.classList.contains('wz-ed') ? e.target : null;
    if (!ed || !ed.dataset.f) return;
    commitPanelEd(ed);
  });
  function commitPanelEd(ed) {
    var it = selectedId && RM.itemById(state, selectedId);
    if (!it) return;
    var f = ed.dataset.f;
    var v = sanitizeHtml(ed.innerHTML);
    if (selStory && f.indexOf('stcol:') === 0) {
      var st = storyById(it, selStory);
      if (!st) return;
      var cur = RM.storyScopeValue(st, f.slice(6));
      if (v === cur || v === displayHtml(cur)) return;
      var stId = selStory;
      commit('story field', function (s) {
        var st2 = storyById(RM.itemById(s, it.id) || {}, stId);
        if (!st2) return;
        RM.setStoryScopeValue(st2, f.slice(6), v);
      });
      return;
    }
    if (f.indexOf('col:') === 0) {
      var key = f.slice(4);
      if (v === RM.scopeValue(it, key)) return;
      commit('scope ' + key, function (s) { RM.setScopeValue(RM.itemById(s, it.id), key, v); });
    }
  }

  // Blur is not a reliable commit trigger: the drag surfaces preventDefault
  // on pointerdown (so focus never leaves the field) and macOS WebKit
  // buttons don't take focus on click at all. Anything that re-renders the
  // panel goes through here first so the focused field's pending edit lands.
  var flushingEdit = false;
  function flushPanelEdit() {
    if (flushingEdit) return;
    var ae = document.activeElement;
    var panel = $('#panel');
    if (!ae || !panel || panel.hidden || !panel.contains(ae)) return;
    flushingEdit = true;
    try {
      if (ae.classList.contains('wz-ed') && ae.dataset.f) commitPanelEd(ae);
      else if ((ae.tagName === 'INPUT' || ae.tagName === 'TEXTAREA') && !ae.readOnly &&
        (ae.type === 'checkbox' ? ae.checked !== ae.defaultChecked : ae.value !== ae.defaultValue)) {
        ae.dispatchEvent(new Event('change', { bubbles: true }));
      }
    } finally { flushingEdit = false; }
  }

  $('#panel').addEventListener('change', function (e) {
    var it = selectedId && RM.itemById(state, selectedId);
    if (!it) return;
    if (e.target.dataset.rk) { // panel RICE dropdowns (rice priority scheme)
      var rk = e.target.dataset.rk;
      var rv = e.target.value === '' ? null : parseFloat(e.target.value);
      commit('rice', function (s) {
        var x = RM.itemById(s, it.id);
        if (x) x.rice[rk] = rv;
      });
      return;
    }
    var stf = e.target.dataset.stf;
    if (stf && selStory) {
      var stId = selStory;
      if (stf === 'stdepsearch') return; // combobox, not a field
      var sval = e.target.type === 'checkbox' ? e.target.checked : e.target.value;
      if (stf === 'startDate') {
        if (!sval) { render(); return; }
        var sd = RM.dateToDay(state.meta, RM.parseISO(sval));
        if (sd == null) { render(); return; }
        commit('story start', function (s) {
          var st2 = storyById(RM.itemById(s, it.id) || {}, stId);
          if (st2 && st2.startDay != null) st2.startDay = Math.max(0, sd);
        });
        return;
      }
      if (stf === 'durWeeks') {
        var swv = Math.max(0.2, parseFloat(sval) || 1);
        commit('story duration', function (s) {
          var st2 = storyById(RM.itemById(s, it.id) || {}, stId);
          if (st2 && st2.startDay != null) st2.durDays = RM.stretchSpan(s.meta, st2.startDay, snapUpDays(Math.max(1, Math.round(swv * SPW())), 'story'));
        });
        return;
      }
      if (stf === 'capMult') {
        var smv = parseFloat(sval);
        commit('story multiplier', function (s) {
          var st2 = storyById(RM.itemById(s, it.id) || {}, stId);
          if (st2) st2.capMult = isFinite(smv) && smv > 0 ? smv : 1;
        });
        return;
      }
      if (stf === 'num') {
        var stNewNum;
        commit('renumber story', function (s) { stNewNum = RM.renumberStory(s, it.id, stId, sval); });
        if (stNewNum != null && String(stNewNum) !== String(sval).trim()) {
          toast('#' + sval + ' isn\u2019t available \u2014 used #' + stNewNum + ' instead');
        }
        return;
      }
      // the Auto flag reads the same in history as the feature's
      var stLbl = stf === 'noAuto' ? (sval ? 'exclude from auto' : 'include in auto') : 'story ' + stf;
      commit(stLbl, function (s) {
        var st2 = storyById(RM.itemById(s, it.id) || {}, stId);
        if (!st2) return;
        if (stf === 'title') st2.title = String(sval);
        else if (stf === 'done') st2.done = !!sval;
        else if (stf === 'noAuto') RM.setNoAuto(st2, !!sval);
        else if (stf === 'jiraKey') st2.jiraKey = RM.jiraKeyOf(String(sval));
      });
      return;
    }
    var f = e.target.dataset.f;
    var pstTitle = e.target.dataset.pstTitle;
    if (pstTitle) {
      var tv = e.target.value;
      commit('story title', function (s) {
        RM.itemById(s, it.id).stories.forEach(function (st) { if (st.id === pstTitle) st.title = tv; });
      });
      return;
    }
    var cf = e.target.dataset.cf;
    if (cf) {
      var cfv = e.target.value;
      commit('scope ' + cf, function (s) { RM.setScopeValue(RM.itemById(s, it.id), cf, cfv); });
      return;
    }
    if (!f) return;
    var val = e.target.type === 'checkbox' ? e.target.checked : e.target.value;
    var simple = { feature: 1, description: 1, workstream: 1, enables: 1, outOfScope: 1, notes: 1, extDeps: 1 };
    if (simple[f]) {
      commit(f, function (s) { RM.itemById(s, it.id)[f] = val; });
      return;
    }
    if (f === 'jiraKey') { commit('jira key', function (s) { RM.itemById(s, it.id).jiraKey = RM.jiraKeyOf(val); }); return; }
    if (f === 'num') {
      var newNum;
      commit('renumber', function (s) { newNum = RM.renumberItem(s, it.id, val); });
      if (newNum != null && String(newNum) !== String(val).trim()) {
        toast('#' + val + ' isn’t available — used #' + newNum + ' instead');
      }
      return;
    }
    // Locked and Excluded from Auto are exclusive: each clears the other
    if (f === 'locked') { commit('lock', function (s) { RM.setLocked(RM.itemById(s, it.id), !!val); }); return; }
    if (f === 'noAuto') { commit(val ? 'exclude from auto' : 'include in auto', function (s) { RM.setNoAuto(RM.itemById(s, it.id), !!val); }); return; }
    if (f === 'done') { commit('done', function (s) { RM.itemById(s, it.id).done = !!val; }); return; }
    if (f === 'deadline') {
      var dlv = /^\d{4}-\d{2}-\d{2}$/.test(String(val)) ? String(val) : null;
      commit('deadline', function (s) { RM.itemById(s, it.id).deadline = dlv; });
      return;
    }
    if (f === 'startDate') {
      if (!val) { render(); return; } // cleared/partial date: leave the item untouched
      var d2 = RM.parseISO(val);
      var day2 = RM.dateToDay(state.meta, d2);
      if (day2 == null) { render(); return; }
      commit('start date', function (s) {
        var t = RM.itemById(s, it.id);
        var nd2 = Math.max(0, day2);
        if (t.startDay != null) RM.shiftStories(t, nd2 - t.startDay);
        t.startDay = nd2;
        if (autoOrder) RM.sortItemsByStart(s);
      });
      return;
    }
    if (f === 'durWeeks') {
      if (!isScheduled(it) && String(val).trim() === '') {
        // clearing an unscheduled preset returns it to the default
        commit('duration', function (s) {
          var t = RM.itemById(s, it.id);
          if (t.startDay == null) t.durDays = null;
        });
        return;
      }
      var wv = Math.max(0.2, parseFloat(val) || 1);
      commit('duration', function (s) {
        RM.itemById(s, it.id).durDays = Math.max(1, Math.round(wv * SPW()));
      });
      return;
    }
    if (f === 'capMult') {
      var mvv = parseFloat(val);
      commit('multiplier', function (s) {
        RM.itemById(s, it.id).capMult = isFinite(mvv) && mvv > 0 ? mvv : 1;
      });
      return;
    }
    if (f === 'storyadd') {
      var sv = val.trim();
      if (sv) {
        expanded[it.id] = true;
        commit('add story', function (s) {
          RM.itemById(s, it.id).stories.push({ id: RM.uid('s'), title: sv, done: false, num: RM.nextNum(s) });
        });
      }
      return;
    }
  });

  // Enter in the story add input triggers change
  $('#panel').addEventListener('keydown', function (e) {
    var tin = e.target.closest && e.target.closest('[data-tagadd]');
    if (tin) {
      if (e.key === 'Enter' || e.key === ',') {
        e.preventDefault();
        if (addTagFromInput(tin)) focusTagInput();
        return;
      }
      if (e.key === 'Backspace' && !tin.value) {
        var curTags = tagsOfTarget();
        if (curTags.length) {
          e.preventDefault();
          commitTags(curTags.slice(0, curTags.length - 1));
          focusTagInput();
        }
        return;
      }
      if (e.key === 'Escape') { tin.value = ''; return; }
    }
    if (e.key === 'Enter' && e.target.dataset.f === 'storyadd') {
      e.target.blur();
    }
    if (e.target.dataset.f === 'depsearch') {
      var sug = $('#panel .dep-sug');
      if (e.key === 'Enter') {
        e.preventDefault();
        var first = sug && $('[data-addep]', sug);
        if (first) first.click();
      } else if (e.key === 'Escape') {
        if (sug) { sug.hidden = true; }
        e.target.value = '';
      }
    }
    if (e.target.dataset.stf === 'stdepsearch') {
      var stSug = $('#panel .dep-sug');
      if (e.key === 'Enter') {
        e.preventDefault();
        var stFirst = stSug && $('[data-addstdep]', stSug);
        if (stFirst) stFirst.click();
      } else if (e.key === 'Escape') {
        if (stSug) { stSug.hidden = true; }
        e.target.value = '';
      }
    }
  });

  // Enter on any single-line field in the panel or the left panes (the
  // panel title included) leaves the field and saves it. Blur commits in
  // the browser; the change is re-fired only when blur left it uncommitted
  // (the element is still attached with a dirty value).
  document.addEventListener('keydown', function (e) {
    if (e.key !== 'Enter' || e.defaultPrevented || e.altKey || e.ctrlKey || e.metaKey) return;
    var t = e.target;
    if (!t || !t.closest || !t.closest('#panel, #rows, #resPanel')) return;
    var isArea = t.tagName === 'TEXTAREA';
    if (isArea ? (!t.classList.contains('p-name') || e.shiftKey) : t.tagName !== 'INPUT') return;
    if (!isArea && ['checkbox', 'radio', 'range', 'file', 'color'].indexOf(t.type) !== -1) return;
    if (t.dataset.act === 'st-add' || t.dataset.f === 'depsearch' || t.dataset.stf === 'stdepsearch' || t.dataset.tagadd) return;
    e.preventDefault();
    var wasSel = t.value !== t.defaultValue;
    t.blur();
    if (wasSel && t.isConnected && t.value !== t.defaultValue) {
      t.dispatchEvent(new Event('change', { bubbles: true }));
    }
  });

  // dependency search-by-name combobox
  function addDep(itemId, num) {
    var target = RM.itemByNum(state, num);
    var it = RM.itemById(state, itemId);
    if (!target || !it || target.id === it.id) return;
    if (it.deps.indexOf(num) !== -1) { toast('#' + it.num + ' already depends on #' + num); return; }
    commit('add dep', function (s) { RM.itemById(s, itemId).deps.push(num); });
    toast('#' + it.num + ' now depends on #' + num);
  }
  // story → story dependency search: '#n' / 'n' resolves a number, anything
  // else matches the story title or its feature's title
  function addStoryDep(itemId, stId, num) {
    var meRef = RM.storyRef(state, stId);
    var target = RM.storyByNum(state, num);
    if (!meRef || !target || target.st.id === stId) return;
    if ((meRef.st.deps || []).indexOf(num) !== -1) {
      toast('#' + meRef.st.num + ' already depends on #' + num);
      return;
    }
    commit('add story dep', function (s) {
      var st2 = storyById(RM.itemById(s, itemId) || {}, stId);
      if (st2) { st2.deps = st2.deps || []; st2.deps.push(num); }
    });
    toast('\u201C' + (meRef.st.title || '(untitled)') + '\u201D now depends on #' + num);
  }
  $('#panel').addEventListener('input', function (e) {
    if (e.target.dataset.stf === 'stdepsearch' && selStory) {
      var itS = selectedId && RM.itemById(state, selectedId);
      var meSt = itS && storyById(itS, selStory);
      var sugS = e.target.parentElement.querySelector('.dep-sug');
      if (!meSt || !sugS) return;
      var qS = e.target.value.trim();
      if (!qS) { sugS.hidden = true; sugS.innerHTML = ''; return; }
      var mine = meSt.deps || [];
      var stHits = [];
      if (/^#?\d+$/.test(qS)) {
        var oneRef = RM.storyByNum(state, parseInt(qS.replace('#', ''), 10));
        if (oneRef && oneRef.st.id !== meSt.id && mine.indexOf(oneRef.st.num) === -1) stHits.push(oneRef);
      } else {
        var ql = qS.toLowerCase();
        state.items.forEach(function (o) {
          (o.stories || []).forEach(function (os) {
            if (stHits.length >= 8) return;
            if (os.id === meSt.id || mine.indexOf(os.num) !== -1) return;
            if ((os.title || '').toLowerCase().indexOf(ql) !== -1 ||
              (o.feature || '').toLowerCase().indexOf(ql) !== -1) stHits.push({ it: o, st: os });
          });
        });
      }
      sugS.innerHTML = stHits.map(function (r) {
        return '<button data-addstdep="' + r.st.num + '"><i>#' + r.st.num + '</i> ' +
          esc(shorten(r.st.title || '(untitled)', 44)) +
          (r.it.feature ? ' <em>' + esc(shorten(r.it.feature, 28)) + '</em>' : '') + '</button>';
      }).join('') || '<div class="dep-sug-none">No match</div>';
      sugS.hidden = false;
      return;
    }
    if (e.target.dataset.f !== 'depsearch') return;
    var it = selectedId && RM.itemById(state, selectedId);
    if (!it) return;
    var q = e.target.value.trim().toLowerCase();
    var sug = e.target.parentElement.querySelector('.dep-sug');
    if (!q) { sug.hidden = true; sug.innerHTML = ''; return; }
    var hits = state.items.filter(function (o) {
      if (o.id === it.id || it.deps.indexOf(o.num) !== -1) return false;
      return (o.feature || '').toLowerCase().indexOf(q) !== -1 ||
        (o.epic || '').toLowerCase().indexOf(q) !== -1;
    }).slice(0, 8);
    sug.innerHTML = hits.map(function (o) {
      return '<button data-addep="' + o.num + '"><i>#' + o.num + '</i> ' + esc(shorten(o.feature || '(untitled)', 44)) +
        (o.epic ? ' <em>' + esc(o.epic) + '</em>' : '') + '</button>';
    }).join('') || '<div class="dep-sug-none">No match — Enter keeps it as a text note</div>';
    if (!hits.length) {
      sug.innerHTML += '<button data-addeptext="1">Add “' + esc(shorten(e.target.value.trim(), 36)) + '” as a text dependency</button>';
    }
    sug.hidden = false;
  });
  $('#panel').addEventListener('click', function (e) {
    var addSt = e.target.closest('[data-addstdep]');
    if (addSt && selStory) {
      var itSt = selectedId && RM.itemById(state, selectedId);
      if (itSt) addStoryDep(itSt.id, selStory, parseInt(addSt.dataset.addstdep, 10));
      return;
    }
    var add = e.target.closest('[data-addep]');
    var addTxt = e.target.closest('[data-addeptext]');
    if (!add && !addTxt) return;
    var it = selectedId && RM.itemById(state, selectedId);
    if (!it) return;
    if (add) { addDep(it.id, parseInt(add.dataset.addep, 10)); return; }
    var inp = $('#panel [data-f=depsearch]');
    var tv = inp ? inp.value.trim() : '';
    if (tv) commit('add text dep', function (s) { RM.itemById(s, it.id).depsText.push(tv); });
  });

  // ------------------------------------------------------------ rows: click actions
  // manual double-click tracking for placing unscheduled items: the first
  // click deselects, which re-renders and replaces the lane's DOM node — the
  // browser then never fires a native dblclick on it, so we count clicks
  // ourselves (rowsEl survives renders).
  var lanePress = null; // { id, t } of the last empty-lane click

  // a title's double-click is counted by hand: the first click selects the
  // row, which re-renders and replaces the node, so the browser never fires
  // a native dblclick on it (same story as placing on an empty lane)
  var titlePress = null; // { key, t } of the last click on a title span
  rowsEl.addEventListener('click', function (e) {
    if (dragConsumedClick) { dragConsumedClick = false; return; }
    // scoping text cells and inline editors edit in place — selecting would
    // re-render and steal their focus
    if (e.target.closest('.sc-edit,.st-add-input,.hc-edit,input.r-name,.sc-name,.st-name')) return;
    var tEl = e.target.closest('.r-name-txt,.st-title-txt');
    if (tEl) {
      var tRow0 = tEl.closest('.row');
      var tKey = tRow0 ? (tRow0.dataset.story ? 'st:' + tRow0.dataset.story : tRow0.dataset.id) : null;
      var tNow = Date.now();
      if (tKey && titlePress && titlePress.key === tKey && tNow - titlePress.t < 450) {
        titlePress = null;
        plStartRename(tRow0.dataset.id, tRow0.dataset.story || null);
        return;
      }
      titlePress = tKey ? { key: tKey, t: tNow } : null;
    } else titlePress = null;
    var act = e.target.closest('[data-act]');
    var rowEl = e.target.closest('.row');
    if (!rowEl) return;

    if (rowEl.dataset.kind === 'addrow') {
      if (!e.target.closest('.row-left')) return; // the lane is empty space, not a button
      var grp = null;
      if (rowEl.dataset.epic != null || rowEl.dataset.ws != null) {
        grp = {};
        if (rowEl.dataset.epic != null) grp.epic = rowEl.dataset.epic;
        if (rowEl.dataset.ws != null) grp.workstream = rowEl.dataset.ws;
      }
      addFeature(rowEl.dataset.phase, grp);
      return;
    }

    if (rowEl.dataset.kind === 'band') {
      var phase = null;
      state.phases.forEach(function (p) { if (p.id === rowEl.dataset.phase) phase = p; });
      if (!phase) return;
      if (act && act.dataset.act === 'phase-toggle') {
        commit('toggle phase', function (s) {
          s.phases.forEach(function (p) { if (p.id === phase.id) p.collapsed = !p.collapsed; });
        });
      } else if (act && act.dataset.act === 'phase-edit') {
        phaseModal(phase.id);
      } else if (act && act.dataset.act === 'phase-auto') {
        if (!act.disabled) autoTimelinePhase(phase.id);
      } else if (act && act.dataset.act === 'phase-additem') {
        addFeature(phase.id);
      }
      return;
    }

    var itemId = rowEl.dataset.id;
    var it = itemId && RM.itemById(state, itemId);
    if (!it) return;

    var storyEl = e.target.closest('[data-story]');
    if (storyEl && act) {
      var stId = storyEl.dataset.story;
      if (act.dataset.act === 'st-open') {
        selectStory(itemId, stId);
      } else if (storyChipAction(act.dataset.act, act, itemId, stId)) {
        return;
      } else if (act.dataset.act === 'st-asg') {
        if (!state.team.length) { toast('Add people in the Resources panel first'); return; }
        openDropdown(act, storyAssignMenuItems(itemId, stId), ASSIGN_DD);
      } else if (act.dataset.act === 'st-startd') {
        var stObjD = storyById(it, stId);
        openCalendar(act, stObjD && stObjD.startDay != null ? RM.fmtISO(RM.dayToDate(state.meta, stObjD.startDay)) : '',
          function (iso) {
            commit('story start', function (s) {
              var st2 = storyById(RM.itemById(s, itemId) || {}, stId);
              if (!st2) return;
              if (!iso) { st2.startDay = null; return; }
              var sd2 = RM.dateToDay(s.meta, RM.parseISO(iso));
              if (sd2 == null) return;
              st2.startDay = Math.max(0, sd2);
              if (st2.durDays == null) st2.durDays = RM.stretchSpan(s.meta, st2.startDay, snapUpDays(Math.max(1, RM.storyEffortDays(s, st2)), 'story'));
            });
          }, { allowClear: true, clearLabel: 'Unschedule' });
      } else if (act.dataset.act === 'st-dl') {
        var stObjL = storyById(it, stId);
        openCalendar(act, (stObjL && stObjL.deadline) || '', function (iso) {
          commit('story deadline', function (s) {
            var st2 = storyById(RM.itemById(s, itemId) || {}, stId);
            if (st2) st2.deadline = iso || null;
          });
        }, { allowClear: true, clearLabel: 'Remove deadline' });
      }
      return;
    }
    // clicking a story's left pane or bar opens the story panel — but not
    // when the click lands in the editable title (scoping renames in place)
    if (storyEl && (e.target.closest('.row-left') || e.target.closest('[data-stbar]'))) {
      if (e.target.closest('[contenteditable="true"]')) return;
      selectStory(itemId, storyEl.dataset.story);
      return;
    }

    if (act) {
      switch (act.dataset.act) {
        case 'wk': {
          if (act.querySelector('input')) return;
          var wkSched = isScheduled(it);
          var wkInp = document.createElement('input');
          wkInp.type = 'number';
          wkInp.min = '0';
          wkInp.step = '0.2';
          wkInp.value = it.milestone ? 0
            : wkSched
            ? Math.round((RM.workInSpan(state.meta, it.startDay, it.durDays) / SPW()) * 10) / 10
            : (it.durDays != null ? Math.round(it.durDays / SPW() * 100) / 100 : '');
          wkInp.placeholder = '1';
          wkInp.className = 'hc-edit';
          wkInp.style.width = '30px';
          act.textContent = '';
          act.appendChild(wkInp);
          wkInp.focus();
          wkInp.select();
          var wkDone = false;
          var wkFin = function (saveIt) {
            if (wkDone) return; wkDone = true;
            var wv = parseFloat(wkInp.value);
            if (saveIt && String(wkInp.value).trim() === '') {
              // empty = off the timeline; the item keeps its kind
              commit('duration', function (s) {
                var t = RM.itemById(s, itemId);
                t.startDay = null;
                t.durDays = null;
              });
            } else if (saveIt && wv === 0) {
              // 0 duration = a milestone pinned at its start date
              commit('duration', function (s) {
                var t = RM.itemById(s, itemId);
                t.milestone = true;
                t.durDays = t.startDay != null ? 0 : null;
                t.riskDays = 0;
                t.size = null; t.priority = null;
              });
            } else if (saveIt && isFinite(wv) && wv > 0) {
              commit('duration', function (s) {
                var t = RM.itemById(s, itemId);
                t.milestone = false;
                t.durDays = t.startDay != null
                  ? RM.stretchSpan(s.meta, t.startDay, Math.max(1, Math.round(wv * SPW())))
                  : Math.max(1, Math.round(wv * SPW()));
              });
            } else render();
          };
          wkInp.addEventListener('blur', function () { wkFin(true); });
          wkInp.addEventListener('keydown', function (ev) {
            ev.stopPropagation();
            if (ev.key === 'Enter') wkFin(true);
            if (ev.key === 'Escape') wkFin(false);
          });
          return;
        }
        case 'startd': {
          openCalendar(act, isScheduled(it) ? RM.fmtISO(RM.dayToDate(state.meta, it.startDay)) : '', function (iso) {
            if (!iso) {
              if (!isScheduled(it)) return;
              commit('start date', function (s) {
                var t = RM.itemById(s, itemId);
                t.startDay = null;
                t.durDays = null;
              });
              return;
            }
            var day = RM.dateToDay(state.meta, RM.parseISO(iso));
            if (day == null) return;
            commit('start date', function (s) {
              var t = RM.itemById(s, itemId);
              t.startDay = Math.max(0, day);
              if (t.milestone) t.durDays = 0;
              else if (t.durDays == null) t.durDays = RM.stretchSpan(s.meta, t.startDay, RM.effortDays(s, t));
              var need = Math.ceil((t.startDay + (t.durDays || 0) + (t.riskDays || 0)) / RM.slotsOf(s.meta));
              if (need > s.meta.numWeeks) { s.meta.numWeeks = need; RM.syncEndDate(s.meta); }
            });
          }, { allowClear: true, clearLabel: 'Unschedule' });
          return;
        }
        case 'deadline': {
          openCalendar(act, it.deadline || '', function (iso) {
            commit('deadline', function (s) {
              RM.itemById(s, itemId).deadline = iso || null;
            });
          }, { allowClear: true, clearLabel: 'Remove deadline' });
          return;
        }
        case 'priority': {
          openPriorityEditor(act, itemId);
          return;
        }
        case 'asg': {
          if (!state.team.length) { toast('Add people in the Resources panel first'); return; }
          openDropdown(act, assignMenuItems(itemId), ASSIGN_DD);
          return;
        }
        case 'size': {
          if (sizeLocked(it)) return;
          openDropdown(act, [{ label: '<i>no size</i>', checked: !it.size, fn: function () {
            setItemSize(itemId, null);
          } }].concat(RM.sizeOrderOf(state).map(function (sz) {
            return { label: esc(sz) + ' <small>' + sizeHuman(sz) + '</small>', checked: it.size === sz, fn: function () {
              setItemSize(itemId, sz);
            } };
          })));
          return;
        }
        case 'risk': {
          if (RM.riskSchemeOf(state) === 'auto') return; // computed, not set
          openDropdown(act, [{ label: '<i>None</i>', checked: !it.risk, fn: function () {
            setItemRisk(itemId, null);
          } }].concat(RM.riskOrderOf(state).map(function (sz) {
            return { icon: LEVEL_GLYPHS[sz], label: esc(riskValueLabel(sz)), checked: it.risk === sz, fn: function () {
              setItemRisk(itemId, sz);
            } };
          })));
          return;
        }
        case 'cap': {
          openDropdown(act, itemCapMenuItems(itemId));
          return;
        }
        case 'mult': {
          inlineMultEditor(act, it.capMult, function (v) {
            commit('multiplier', function (s) { RM.itemById(s, itemId).capMult = v; });
          });
          return;
        }
        case 'epic': {
          openDropdown(act, setEpicMenu(itemId, true));
          return;
        }
        case 'ws': {
          openDropdown(act, wsMenuItems(itemId, function () {
            return rowsEl.querySelector('.row[data-id="' + itemId + '"] .r-ws, .row[data-id="' + itemId + '"] .r-ws-col');
          }));
          return;
        }
        case 'stories': {
          // toggle just this item's stories — the view level stays put
          expanded[itemId] = !expanded[itemId];
          saveLocal();
          render();
          return;
        }
        case 'warn': {
          select(selectedId === itemId ? null : itemId);
          return;
        }
      }
    }
    // story lane: two quick clicks give the story its own timeline
    if (storyEl && e.target.closest('.row-lane') && !e.target.closest('[data-stbar]')) {
      var stLaneId = storyEl.dataset.story;
      var stObj = stLaneId && storyById(it, stLaneId);
      if (view === 'planning' && stObj && stObj.startDay == null) {
        var stNow = Date.now();
        if (lanePress && lanePress.id === 'st:' + stLaneId && stNow - lanePress.t < 450) {
          lanePress = null;
          placeStoryAt(itemId, stLaneId, e.clientX);
          return;
        }
        lanePress = { id: 'st:' + stLaneId, t: stNow };
      } else lanePress = null;
      selectStory(itemId, stLaneId);
      return;
    }

    // any other click in the row selects the item and opens the panel; two
    // quick clicks on an unscheduled row's lane also place the item (manual
    // count — see lanePress above)
    if (e.target.closest('.row-lane') &&
      !e.target.closest('[data-bar],[data-ghost],.ghost-pill,.port,.sc-row')) {
      if (view === 'planning' && !isScheduled(it) && !rowEl.classList.contains('story')) {
        var now = Date.now();
        if (lanePress && lanePress.id === itemId && now - lanePress.t < 450) {
          lanePress = null;
          placeItemAt(itemId, e.clientX);
          return;
        }
        lanePress = { id: itemId, t: now };
      } else lanePress = null;
      select(itemId);
      return;
    }
    if (e.shiftKey && rowEl.classList.contains('item')) { extendSelectionTo(itemId); return; }
    select(itemId);
  });

  // one workstream picker everywhere an item's workstream can change (scoping
  // chip, right panel, context menu): color dot per stream, default first,
  // pencil to edit, "Other…" types a brand-new one in place
  function wsMenuItems(itemId, chipFinder) {
    var it = RM.itemById(state, itemId);
    var wsItems = [{
      label: esc(RM.defaultWsName(state)) + ' <i>(default)</i>',
      dot: '#' + RM.defaultWsColor(state),
      checked: !it.workstream,
      fn: function () { commit('workstream', function (s) { RM.itemById(s, itemId).workstream = ''; }); },
      edit: function () { defaultWsModal(); }
    }];
    allWorkstreams().forEach(function (w) {
      wsItems.push({
        label: esc(w),
        dot: '#' + RM.colorForWs(state, w),
        checked: it.workstream === w,
        fn: function () { commit('workstream', function (s) { RM.itemById(s, itemId).workstream = w; }); },
        edit: function () { wsEditModal(w); }
      });
    });
    wsItems.push({ sep: true });
    wsItems.push({ icon: 'pencil', label: 'Other…', fn: function () {
      var chip = chipFinder && chipFinder();
      if (!chip) return;
      chip.textContent = it.workstream || '';
      startInlineEdit(chip, function (v) {
        commit('workstream', function (s) { RM.itemById(s, itemId).workstream = v; });
      });
    } });
    return wsItems;
  }

  function assignMenuItems(itemId) {
    var it = RM.itemById(state, itemId);
    return state.team.map(function (mm) {
      var onA = (it.assignees || []).indexOf(mm.id) !== -1;
      return { pre: avatarHtml(mm, 'sm'), label: esc(mLabel(mm)) + (mSub(mm) ? ' <small>' + esc(mSub(mm)) + '</small>' : ''), checked: onA, fn: function () {
        commit('assignees', function (s) {
          var t = RM.itemById(s, itemId);
          t.assignees = t.assignees || [];
          var at = t.assignees.indexOf(mm.id);
          if (at === -1) t.assignees.push(mm.id);
          else t.assignees.splice(at, 1);
        });
      } };
    });
  }

  // every assignee picker is the same searchable list of avatars
  var ASSIGN_DD = { search: 'Search people\u2026' };
  // shared by the panel's story-assignee dropdown and the Sprinting story chip
  function storyAssignMenuItems(itemId, stId) {
    var it = RM.itemById(state, itemId);
    // only people supplying the story's capacity type (everyone when it has
    // none, or nobody supplies it)
    return RM.assignableFor(state, storyById(it, stId)).map(function (mm) {
      var onSA = ((storyById(it, stId) || {}).assignees || []).indexOf(mm.id) !== -1;
      return { pre: avatarHtml(mm, 'sm'), label: esc(mLabel(mm)) + (mSub(mm) ? ' <small>' + esc(mSub(mm)) + '</small>' : ''), checked: onSA, fn: function () {
        commit('story assignees', function (s) {
          var st2 = storyById(RM.itemById(s, itemId) || {}, stId);
          if (!st2) return;
          st2.assignees = st2.assignees || [];
          var atA = st2.assignees.indexOf(mm.id);
          if (atA === -1) st2.assignees.push(mm.id);
          else st2.assignees.splice(atA, 1);
        });
      } };
    });
  }

  // set size / risk from the chip dropdowns
  // under the rollup scheme feature sizes derive from the stories: every
  // size editor says so instead of opening
  function sizeLocked(it) {
    if (!RM.sizeRollup(state)) return false;
    toast(lvl('feature') + ' sizes roll up from ' + lvl('story', true).toLowerCase() + ' \u2014 size the ' + lvl('story', true).toLowerCase() + ' instead');
    return true;
  }
  function sizeRoCls() { return RM.sizeRollup(state) ? ' ro' : ''; }
  function sizeChipTitle() {
    return RM.sizeRollup(state) ? 'Size (sum of story points)' : 'Size';
  }
  function setItemSize(itemId, sz) {
    if (sizeLocked(RM.itemById(state, itemId))) return;
    commit('size', function (s) {
      var t = RM.itemById(s, itemId);
      t.size = sz;
      if (sz && isScheduled(t) && !t.locked) {
        t.durDays = RM.stretchSpan(s.meta, t.startDay, RM.sizeDays(s, sz));
      }
    });
  }
  function setItemRisk(itemId, sz) {
    commit('risk size', function (s) {
      var t = RM.itemById(s, itemId);
      t.risk = sz;
      t.riskDays = 0;
    });
  }


  // right-click context menus on rows
  // right-clicking an empty timeline slot edits the holiday list in place
  function holidayMenuItems(day) {
    var meta = state.meta;
    var iso = RM.fmtISO(RM.dayToDate(meta, day));
    var wk = Math.floor(day / SPW());
    var w0 = RM.fmtISO(RM.dayToDate(meta, wk * SPW()));
    var w1 = RM.fmtISO(RM.dayToDate(meta, (wk + 1) * SPW() - 1));
    var hitIdx = -1, hitName = '';
    (meta.holidayRanges || []).forEach(function (r, i) {
      if (hitIdx === -1 && iso >= r.start && iso <= r.end) { hitIdx = i; hitName = r.name; }
    });
    var items = [];
    if (hitIdx !== -1) {
      var rmIdx = hitIdx;
      items.push({ icon: 'calendar-x', label: 'Remove holiday' + (hitName ? ' \u201C' + esc(shorten(hitName, 22)) + '\u201D' : '') +
        ' \u00B7 ' + esc(RM.fmtShort(RM.parseISO(iso))), fn: function () {
        commit('holiday rm', function (s) { RM.removeHolidayRange(s.meta, rmIdx); });
      } });
    } else {
      items.push({ icon: 'calendar-plus', label: 'Add holiday \u00B7 ' + esc(RM.fmtShort(RM.parseISO(iso))), fn: function () {
        commit('holiday add', function (s) { RM.addHolidayRange(s.meta, '', iso, iso); });
      } });
    }
    var fullWeek = RM.holidaysInWeek(meta, wk) === SPW();
    items.push(fullWeek
      ? { icon: 'calendar-check', label: 'Make week of ' + esc(RM.fmtShort(RM.parseISO(w0))) + ' a working week', fn: function () {
          commit('toggle holiday', function (s) { RM.clipHolidayRanges(s.meta, w0, w1); });
        } }
      : { icon: 'calendar-off', label: 'Mark week of ' + esc(RM.fmtShort(RM.parseISO(w0))) + ' as a holiday week', fn: function () {
          commit('toggle holiday', function (s) { RM.addHolidayRange(s.meta, '', w0, w1); });
        } });
    items.push({ sep: true });
    items.push({ icon: 'settings-2', label: 'Holiday settings\u2026', fn: function () {
      view = 'setup';
      setupTab = 'project';
      saveLocal();
      render();
    } });
    return items;
  }

  rowsEl.addEventListener('contextmenu', function (e) {
    var rowEl = e.target.closest('.row');
    if (!rowEl || e.target.closest('input,textarea,select')) return;
    if ((view === 'planning' || view === 'budget') &&
        e.target.closest('.row-lane') && !e.target.closest('.bar,.st-bar')) {
      e.preventDefault();
      e.stopPropagation();
      openContextMenu(e.clientX, e.clientY, holidayMenuItems(Math.max(0, laneDayAt(e.clientX))));
      return;
    }
    if (view === 'budget') return; // budget rows carry their own menu below
    e.preventDefault();
    e.stopPropagation(); // the document fallback must not replace this menu
    var cx = e.clientX, cy = e.clientY;
    var items = [];
    if (rowEl.dataset.kind === 'band') {
      var phaseId = rowEl.dataset.phase;
      items = [
        { icon: 'plus', label: esc('Add ' + lvl('feature').toLowerCase() + ' here'), fn: function () { addFeature(phaseId); } },
        { icon: 'plus', label: 'New phase…', fn: function () { phaseModal(null); } },
        { sep: true },
        { icon: 'pencil', label: 'Edit phase…', fn: function () { phaseModal(phaseId); } },
        (function () {
          var ph = state.phases.filter(function (p) { return p.id === phaseId; })[0];
          if (!ph || ph.bucket) return null;
          var as = autoPhaseStatus(phaseId);
          return { icon: 'zap', label: 'Auto timeline', disabled: as.disabled, title: as.tip,
            fn: function () { autoTimelinePhase(phaseId); } };
        })(),
        { icon: 'trash-2', label: 'Delete phase…', fn: function () { deletePhaseConfirm(phaseId); } }
      ];
    } else if (rowEl.dataset.kind === 'eband') {
      var epName = rowEl.dataset.epic;
      var wsName = rowEl.dataset.ws;
      if (wsName) {
        items = [{ icon: 'pencil', label: 'Edit workstream…', fn: function () { wsEditModal(wsName); } }];
      } else if (epName) {
        items = [
          { icon: 'pencil', label: 'Edit epic…', fn: function () { epicEditModal(epName); } },
          { sep: true },
          { icon: 'trash-2', label: 'Delete epic…', fn: function () { deleteEpicConfirm(epName); } }
        ];
      } else return; // "no epic" / "no workstream" groups
    } else if (rowEl.classList.contains('story') && rowEl.dataset.story) {
      var stmId = rowEl.dataset.story, stmItemId = rowEl.dataset.id;
      var stmIt = RM.itemById(state, stmItemId);
      var stm = stmIt && storyById(stmIt, stmId);
      if (!stm) return;
      items = [
        view === 'scoping' ? null : plRenameEntry(stmItemId, stmId),
        stm.startDay != null ? { icon: 'calendar-off', label: 'Remove timeline', fn: function () {
          commit('story timeline', function (s) {
            var st = storyById(RM.itemById(s, stmItemId), stmId);
            if (st) { st.startDay = null; st.durDays = null; }
          });
        } } : null,
        placeEntry(stmItemId, stmId),
        { icon: RM.typeOf(state, stm, 'story').icon, label: 'Type: ' + esc(RM.typeOf(state, stm, 'story').label) + '…', fn: function () {
          openContextMenu(cx, cy, typeMenuItems('story', stm.type, function (k) { setStoryType(stmItemId, stmId, k); }));
        } },
        flagMenuEntries(stmItemId, stmId)[0], flagMenuEntries(stmItemId, stmId)[1] || null, noAutoEntry(stmItemId, stmId),
        moveStoryFeatureEntry(cx, cy, stmItemId, stmId),
        { sep: true },
        storyInsertEntries(stmItemId, stmId)[0], storyInsertEntries(stmItemId, stmId)[1],
        { icon: 'copy', label: 'Duplicate story', fn: function () { duplicateStory(stmItemId, stmId); } },
        { icon: 'trash-2', label: 'Delete story', fn: function () {
          commit('delete story', function (s) {
            var t = RM.itemById(s, stmItemId);
            t.stories = t.stories.filter(function (x) { return x.id !== stmId; });
          });
        } }
      ].filter(Boolean);
    } else if (rowEl.classList.contains('item')) {
      var itemId = rowEl.dataset.id;
      var it = RM.itemById(state, itemId);
      if (!it) return;
      if (multiSel && multiSel.length > 1 && multiSel.indexOf(itemId) !== -1) {
        openContextMenu(cx, cy, bulkMenuItems(multiSel.slice(), cx, cy));
        return;
      }
      items = [
        view === 'scoping' ? null : plRenameEntry(itemId, null),
        view === 'scoping' ? null : { sep: true },
        it.milestone ? null : { icon: 'plus', label: esc('Add ' + lvl('story').toLowerCase()), fn: function () { addStoryNear(itemId, null, 0); } },
        { icon: 'plus', label: esc('Insert ' + lvl('feature').toLowerCase() + ' above'), fn: function () { addFeatureNear(itemId, 0); } },
        { icon: 'plus', label: esc('Insert ' + lvl('feature').toLowerCase() + ' below'), fn: function () { addFeatureNear(itemId, 1); } },
        { icon: 'plus', label: 'New phase…', fn: function () { phaseModal(null); } },
        { sep: true },
        { icon: 'folder-input', label: 'Move to phase…', fn: function () { openContextMenu(cx, cy, movePhaseMenu(itemId)); } },
        placeEntry(itemId, null),
        { icon: 'tag', label: 'Set epic…', fn: function () { openContextMenu(cx, cy, setEpicMenu(itemId, false)); } },
        { icon: RM.typeOf(state, it, 'feature').icon, label: 'Type: ' + esc(RM.typeOf(state, it, 'feature').label) + '…', fn: function () {
          openContextMenu(cx, cy, typeMenuItems('feature', it.type, function (k) { setItemType(itemId, k); }));
        } },
        state.meta.workstreamsEnabled
          ? { icon: 'layers', label: 'Set workstream…', fn: function () {
              openContextMenu(cx, cy, wsMenuItems(itemId, function () {
                return rowsEl.querySelector('.row[data-id="' + itemId + '"] .r-ws, .row[data-id="' + itemId + '"] .r-ws-col');
              }));
            } }
          : null,
        { sep: true },
        isScheduled(it) ? { icon: 'calendar-off', label: 'Unschedule', fn: function () {
          commit('unschedule', function (s) {
            var t = RM.itemById(s, itemId);
            t.startDay = null; t.durDays = null; t.riskDays = 0;
          });
        } } : null,
        flagMenuEntries(itemId, null)[0], flagMenuEntries(itemId, null)[1] || null,
        { icon: it.locked ? 'lock-open' : 'lock', label: it.locked ? 'Unlock' : 'Lock', fn: function () {
          commit('lock', function (s) { var t = RM.itemById(s, itemId); RM.setLocked(t, !t.locked); });
        } },
        noAutoEntry(itemId, null),
        { icon: it.done ? 'circle' : 'circle-check', label: it.done ? 'Unmark as done' : 'Mark as done', fn: function () {
          commit('done', function (s) { var t = RM.itemById(s, itemId); t.done = !t.done; });
        } },
        { icon: it.milestone ? 'rectangle-horizontal' : 'gem',
          label: it.milestone ? esc('Convert to ' + lvl('feature').toLowerCase()) : 'Convert to milestone',
          fn: function () { toggleMilestone(itemId); } }
      ].concat(it.milestone ? msStyleItems(itemId, it) : []).concat([
        { sep: true },
        { icon: 'copy', label: 'Duplicate', fn: function () { duplicateItem(itemId); } },
        { icon: 'trash-2', label: 'Delete #' + it.num + '…', fn: function () { deleteItemConfirm(itemId); } }
      ]).filter(Boolean);
    } else return;
    openContextMenu(cx, cy, items);
  });

  function openContextMenu(x, y, items, opts) {
    items = (items || []).filter(Boolean); // builders leave null for entries that do not apply
    var html = menuListHtml(items, 0, opts);
    openPopover(x, y, html, function (host) {
      if (window.lucide) lucide.createIcons();
      if (opts && opts.search) wireMenu(host, items);
      host.addEventListener('click', function (ev) {
        var mi = ev.target.closest('[data-mi]');
        if (!mi) return;
        closePopover();
        items[parseInt(mi.dataset.mi, 10)].fn();
      });
    });
  }

  // right-click on app chrome (header, setup, start page background) →
  // right panel: right-click = the item / story action menu (the actions the
  // old ⋯ button carried — duplicate, convert, delete)
  $('#panel').addEventListener('contextmenu', function (e) {
    if (e.target.closest('input, textarea, select, [contenteditable="true"], button')) return;
    var it = selectedId != null ? RM.itemById(state, selectedId) : null;
    if (!it) return;
    e.preventDefault();
    e.stopPropagation(); // the theme fallback must not replace this menu
    if (selStory) {
      var stMoreId = selStory;
      openContextMenu(e.clientX, e.clientY, [
        moveStoryFeatureEntry(e.clientX, e.clientY, it.id, stMoreId),
        { sep: true },
        storyInsertEntries(it.id, stMoreId)[0], storyInsertEntries(it.id, stMoreId)[1],
        flagMenuEntries(it.id, stMoreId)[0], flagMenuEntries(it.id, stMoreId)[1] || null, noAutoEntry(it.id, stMoreId),
        { icon: 'copy', label: 'Duplicate story', fn: function () { duplicateStory(it.id, stMoreId); } },
        { icon: 'trash-2', label: 'Delete story', danger: true, fn: function () {
          commit('delete story', function (s) {
            var t = RM.itemById(s, it.id);
            t.stories = t.stories.filter(function (x) { return x.id !== stMoreId; });
            selStory = null;
          });
        } }
      ]);
      return;
    }
    openContextMenu(e.clientX, e.clientY, [
      flagMenuEntries(it.id, null)[0], flagMenuEntries(it.id, null)[1] || null, noAutoEntry(it.id, null),
      { icon: 'copy', label: 'Duplicate', fn: function () { duplicateItem(it.id); } },
      { icon: it.milestone ? 'rectangle-horizontal' : 'gem',
        label: it.milestone ? esc('Convert to ' + lvl('feature').toLowerCase()) : 'Convert to milestone',
        fn: function () { toggleMilestone(it.id); } }
    ].concat(it.milestone ? msStyleItems(it.id, it) : []).concat([
      { sep: true },
      { icon: 'trash-2', label: 'Delete…', danger: true, fn: function () { deleteItemConfirm(it.id); } }
    ]));
  });

  // No surface shows the browser's default context menu. Editable text keeps
  // the native menu (copy/paste is essential there); specific surfaces attach
  // their own menus and preventDefault first; everywhere else a right-click
  // does nothing (the theme lives in Settings and the app menu).
  document.addEventListener('contextmenu', function (e) {
    if (e.defaultPrevented) return; // rows / resources / budget handled it already
    if (e.target.closest('input, textarea, select, [contenteditable="true"]')) return;
    // start-page recents: open / remove
    var spRow = e.target.closest('.sp-row[data-sp-open]');
    if (spRow) {
      e.preventDefault();
      var spPath = spRow.dataset.spOpen;
      openContextMenu(e.clientX, e.clientY, [
        { icon: 'folder-open', label: 'Open', fn: function () { openRecent(spPath); } },
        { icon: 'x', label: 'Remove from this list', fn: function () { dropRecent(spPath); renderStartPage(); } }
      ]);
      return;
    }
    // scoping column headers: right-click = the column menu
    var colHd = e.target.closest('#hdrSprints .sc-hcell[data-col]');
    if (colHd && view === 'scoping' && !isFixedColKey(colHd.dataset.col)) {
      e.preventDefault();
      openContextMenu(e.clientX, e.clientY, scopeColMenuItems(colHd.dataset.col));
      return;
    }
    e.preventDefault();
  });

  // right-click on a multi-selection: every action applies to the whole group
  function bulkMenuItems(ids, cx, cy) {
    var sel = ids.map(function (id) { return RM.itemById(state, id); }).filter(Boolean);
    function each(label, fn) {
      commit(label, function (s) {
        ids.forEach(function (id) {
          var t = RM.itemById(s, id);
          if (t) fn(s, t);
        });
      });
    }
    var allDone = sel.every(function (t) { return t.done; });
    var allLocked = sel.every(function (t) { return t.locked; });
    var allNoAuto = sel.every(function (t) { return t.noAuto; });
    return [
      { label: '<i>' + ids.length + ' selected</i>', fn: function () {} },
      { sep: true },
      { icon: 'folder-input', label: 'Move to phase…', fn: function () {
        openContextMenu(cx, cy, state.phases.map(function (p) {
          return { label: esc(p.name) + (p.bucket ? ' <small>(backlog)</small>' : ''),
            checked: sel.every(function (t) { return t.phaseId === p.id; }),
            fn: function () {
              commit('move phase', function (s) {
                // the group lands at the end of the target phase, in its own order
                var moving = s.items.filter(function (x) { return ids.indexOf(x.id) !== -1; });
                s.items = s.items.filter(function (x) { return ids.indexOf(x.id) === -1; });
                moving.forEach(function (t) { t.phaseId = p.id; });
                var lastIdx = -1;
                s.items.forEach(function (x, i2) { if (x.phaseId === p.id) lastIdx = i2; });
                s.items.splice.apply(s.items, [lastIdx + 1, 0].concat(moving));
              });
            } };
        }));
      } },
      { icon: 'tag', label: 'Set epic…', fn: function () {
        var eItems = [{ label: '<i>— none —</i>', checked: sel.every(function (t) { return !t.epic; }),
          fn: function () { each('epic', function (s, t) { t.epic = ''; }); } }];
        allEpics().forEach(function (ep) {
          eItems.push({ label: esc(ep), icon: RM.iconForEpic(state, ep) || 'tag',
            checked: sel.every(function (t) { return t.epic === ep; }),
            fn: function () { each('epic', function (s, t) { t.epic = ep; }); } });
        });
        openContextMenu(cx, cy, eItems);
      } },
      { icon: 'tag', label: 'Type…', fn: function () {
        openContextMenu(cx, cy, typeMenuItems('feature', sel.length && sel.every(function (t) { return t.type === sel[0].type; }) ? sel[0].type : null, function (k) {
          commit('type', function (s) {
            ids.forEach(function (id) {
              var t = RM.itemById(s, id);
              if (t && RM.itemType(s, k)) t.type = k;
            });
          });
        }));
      } },
      state.meta.workstreamsEnabled
        ? { icon: 'layers', label: 'Set workstream…', fn: function () {
            var wItems = [{ label: esc(RM.defaultWsName(state)) + ' <i>(default)</i>',
              dot: '#' + RM.defaultWsColor(state),
              checked: sel.every(function (t) { return !t.workstream; }),
              fn: function () { each('workstream', function (s, t) { t.workstream = ''; }); } }];
            allWorkstreams().forEach(function (w) {
              wItems.push({ label: esc(w), dot: '#' + RM.colorForWs(state, w),
                checked: sel.every(function (t) { return t.workstream === w; }),
                fn: function () { each('workstream', function (s, t) { t.workstream = w; }); } });
            });
            openContextMenu(cx, cy, wItems);
          } }
        : null,
      { sep: true },
      { icon: allDone ? 'circle' : 'circle-check', label: allDone ? 'Unmark as done' : 'Mark as done',
        fn: function () { each('done', function (s, t) { t.done = !allDone; }); } },
      { icon: allLocked ? 'lock-open' : 'lock', label: allLocked ? 'Unlock' : 'Lock',
        fn: function () { each('lock', function (s, t) { RM.setLocked(t, !allLocked); }); } },
      { icon: allNoAuto ? 'zap' : 'zap-off', label: allNoAuto ? 'Include in Auto timeline' : 'Exclude from Auto timeline',
        fn: function () { each(allNoAuto ? 'include in auto' : 'exclude from auto', function (s, t) { RM.setNoAuto(t, !allNoAuto); }); } },
      sel.some(function (t) { return isScheduled(t); })
        ? { icon: 'calendar-off', label: 'Unschedule', fn: function () {
            each('unschedule', function (s, t) { t.startDay = null; t.durDays = null; t.riskDays = 0; });
          } }
        : null,
      { sep: true },
      { icon: 'trash-2', label: 'Delete ' + ids.length + '…', danger: true, fn: function () {
        confirmBox('Delete ' + ids.length + ' items?', 'Every selected item is removed.', 'Delete', function () {
          commit('delete', function (s) {
            s.items = s.items.filter(function (x) { return ids.indexOf(x.id) === -1; });
            selectedId = null;
            multiSel = null;
          });
        }, true);
      } }
    ].filter(Boolean);
  }

  // menu-item builders shared by the panel dropdowns and row context menus
  function movePhaseMenu(itemId) {
    var it = RM.itemById(state, itemId);
    return state.phases.map(function (p) {
      return { label: esc(p.name) + (p.bucket ? ' <small>(backlog)</small>' : ''), checked: p.id === it.phaseId, fn: function () {
        commit('move phase', function (s) {
          var t = RM.itemById(s, itemId);
          t.phaseId = p.id;
          s.items = s.items.filter(function (x) { return x.id !== t.id; });
          var lastIdx = -1;
          s.items.forEach(function (x, i2) { if (x.phaseId === p.id) lastIdx = i2; });
          s.items.splice(lastIdx + 1, 0, t);
        });
      } };
    });
  }
  // Sprinting context menu: every sprint of the timeline, then Unscheduled
  function moveSprintMenu(itemId) {
    var it = RM.itemById(state, itemId);
    var cur = it && isScheduled(it) ? sprFirstNum(it) : null;
    return sprintNums().map(function (n) {
      return { label: esc(sprintLabel(n)), checked: cur === n, fn: function () {
        commit('move to sprint', function (s) { RM.moveItemToSprint(s, itemId, n, null); });
      } };
    }).concat([{ sep: true }, { icon: 'inbox', label: 'Unscheduled', checked: cur == null, fn: function () {
      commit('unschedule', function (s) { RM.moveItemToSprint(s, itemId, null, null); });
    } }]);
  }
  // story variant: a sprint gives the story its own timeline; "With feature" drops it
  // "Move to feature…": every other feature, searchable by title; the story
  // lands at the end of the target's story list and the selection follows
  function moveStoryToFeature(itemId, stId, toItemId) {
    if (toItemId === itemId) return;
    commit('move story', function (s) {
      var from = RM.itemById(s, itemId);
      var st = from && (from.stories || []).filter(function (x) { return x.id === stId; })[0];
      var to = RM.itemById(s, toItemId);
      if (!st || !to || to.milestone) return;
      from.stories = from.stories.filter(function (x) { return x.id !== stId; });
      to.stories = to.stories || [];
      to.stories.push(st);
      if (selStory === stId) selectedId = toItemId;
    });
  }
  function moveStoryFeatureMenu(itemId, stId) {
    return state.items.filter(function (o) { return !o.milestone && o.id !== itemId; }).map(function (o) {
      var ph = phaseOf(o);
      return { pre: typeGlyphHtml(o, 'feature'), label: esc(o.feature || '(untitled)') + (ph ? ' <small>' + esc(ph.name || '') + '</small>' : ''),
        fn: function () { moveStoryToFeature(itemId, stId, o.id); } };
    });
  }
  function moveStoryFeatureEntry(cx, cy, itemId, stId) {
    return { icon: 'folder-input', label: esc('Move to ' + lvl('feature').toLowerCase() + '\u2026'), fn: function () {
      openContextMenu(cx, cy, moveStoryFeatureMenu(itemId, stId), { search: 'Search ' + lvl('feature').toLowerCase() + 's\u2026' });
    } };
  }
  function moveStorySprintMenu(itemId, stId) {
    var st = storyById(RM.itemById(state, itemId) || {}, stId);
    var cur = st && sprHasOwn(st) ? sprFirstNum(st) : null;
    return sprintNums().map(function (n) {
      return { label: esc(sprintLabel(n)), checked: cur === n, fn: function () {
        commit('move story to sprint', function (s) { RM.moveStoryToSprint(s, itemId, stId, n, stId); });
      } };
    }).concat([{ sep: true }, { icon: 'corner-down-right', label: 'With feature', checked: cur == null, fn: function () {
      commit('story with feature', function (s) { RM.moveStoryToSprint(s, itemId, stId, null, stId); });
    } }]);
  }
  function setEpicMenu(itemId, withNew) {
    var it = RM.itemById(state, itemId);
    var eItems = [{ label: '<i>— none —</i>', checked: !it.epic, fn: function () {
      commit('epic', function (s) { RM.itemById(s, itemId).epic = ''; });
    } }];
    allEpics().forEach(function (ep) {
      eItems.push({
        label: esc(ep),
        icon: RM.iconForEpic(state, ep) || 'tag',
        checked: it.epic === ep,
        fn: function () { commit('epic', function (s) { RM.itemById(s, itemId).epic = ep; }); },
        edit: function () { epicEditModal(ep); }
      });
    });
    if (withNew) {
      eItems.push({ sep: true });
      eItems.push({ icon: 'plus', label: 'New epic…', fn: function () { newEpicInline(itemId); } });
    }
    return eItems;
  }

  function deleteEpicConfirm(epicName) {
    var count = state.items.filter(function (x) { return x.epic === epicName; }).length;
    confirmBox('Delete epic “' + esc(epicName) + '”?',
      'Its ' + count + ' item(s) stay on the roadmap with no epic.',
      'Delete', function () {
        commit('delete epic', function (s) {
          s.items.forEach(function (x) { if (x.epic === epicName) x.epic = ''; });
          s.epicList = (s.epicList || []).filter(function (e) { return e !== epicName; });
          delete s.epicColors[epicName];
          delete s.epicTypes[epicName];
        });
      }, true);
  }

  function deletePhaseConfirm(phaseId) {
    var phase = null;
    state.phases.forEach(function (p) { if (p.id === phaseId) phase = p; });
    if (!phase) return;
    if (state.phases.length <= 1) { toast('The last phase can’t be deleted'); return; }
    var items = RM.itemsInPhase(state, phaseId);
    var dest = state.phases.filter(function (p) { return p.id !== phaseId; })[0];
    confirmBox('Delete phase “' + esc(phase.name) + '”?',
      items.length
        ? 'Its ' + items.length + ' item(s) move to “' + esc(dest.name) + '”.'
        : 'The phase is empty.',
      'Delete', function () {
        commit('delete phase', function (s) {
          s.items.forEach(function (x) { if (x.phaseId === phaseId) x.phaseId = dest.id; });
          s.phases = s.phases.filter(function (p) { return p.id !== phaseId; });
        });
      }, true);
  }

  // insert a blank feature right above (offset 0) or below (offset 1) the
  // anchor row. The edit panel stays closed — the new row's inline title input
  // gets focus instead — and holdPos keeps auto-order from moving the row
  // away until it gets a start date.
  function addFeatureNear(itemId, offset) {
    var anchor = RM.itemById(state, itemId);
    if (!anchor) return;
    var newId = RM.uid('i');
    commit('add feature', function (s) {
      var it = RM.normalizeState({
        meta: s.meta, phases: s.phases,
        items: [{ id: newId, num: RM.nextNum(s), phaseId: anchor.phaseId, feature: '', size: 'M', headcount: 1,
          capType: s.capTypes[0] || '',
          epic: anchor.epic || '', workstream: anchor.workstream || '', teamType: anchor.teamType }]
      }).items[0];
      it.id = newId;
      it.num = RM.nextNum(s);
      it.holdPos = true;
      var idx = s.items.indexOf(RM.itemById(s, itemId));
      s.items.splice(idx + offset, 0, it);
    });
    var row = rowsEl.querySelector('.row[data-id="' + newId + '"]');
    if (row && row.scrollIntoView) row.scrollIntoView({ block: 'nearest' });
    // Scoping types straight into the title cell; elsewhere the title is text
    // until the inline editor opens on it
    var inp = row && row.querySelector('.sc-name');
    if (inp) inp.focus();
    else plStartRename(newId, null);
  }

  function duplicateItem(itemId) {
    commit('duplicate', function (s) {
      var t = RM.itemById(s, itemId);
      if (!t) return;
      var copy = RM.clone(t);
      copy.id = RM.uid('i');
      copy.num = RM.nextNum(s);
      copy.feature = t.feature + ' (copy)';
      // fresh ids and numbers for the copied stories; dependencies among them
      // follow the copy, links to stories outside it keep pointing there
      var next = copy.num + 1, numMap = {};
      copy.stories.forEach(function (st) {
        st.id = RM.uid('s');
        if (st.num != null) numMap[st.num] = next;
        st.num = next++;
      });
      RM.remapStoryDeps(copy.stories, numMap);
      s.items.splice(s.items.indexOf(t) + 1, 0, copy);
      selectedId = copy.id;
    });
  }

  // a blank story next to stId (offset 0 = above, 1 = below; no anchor = at
  // the end), opened for editing: the panel title when the panel is up,
  // else the inline title on the Prioritizing card or the Planning row
  function addStoryNear(itemId, stId, offset) {
    var newId = RM.uid('s');
    commit('add story', function (s) {
      var t = RM.itemById(s, itemId);
      if (!t || t.milestone) return;
      t.stories = t.stories || [];
      var anchor = stId ? storyById(t, stId) : null;
      var idx = anchor ? t.stories.indexOf(anchor) + offset : t.stories.length;
      t.stories.splice(idx, 0, { id: newId, title: '', done: false, num: RM.nextNum(s) });
      expanded[itemId] = true;
      selectedId = itemId;
      selStory = newId;
    });
    var ed = $('#panel textarea[data-stf="title"]') ||
      rowsEl.querySelector('.row.story[data-story="' + newId + '"] .st-name');
    if (ed && ed.focus) { ed.focus(); return; }
    // no panel to type in: open the inline editor on the new story's title
    if ($('#prioView .pr-story[data-prst="' + newId + '"]')) prStartRename(itemId, newId, true);
    else if ($('#prioView .pr-stcard[data-prst="' + newId + '"]')) prStartRename(itemId, newId, false);
    else plStartRename(itemId, newId);
  }
  function storyInsertEntries(itemId, stId) {
    return [
      { icon: 'plus', label: esc('Insert ' + lvl('story').toLowerCase() + ' above'), fn: function () { addStoryNear(itemId, stId, 0); } },
      { icon: 'plus', label: esc('Insert ' + lvl('story').toLowerCase() + ' below'), fn: function () { addStoryNear(itemId, stId, 1); } }
    ];
  }
  // ---- attention flags: any feature or story, optional reason, via the context menus
  function flagTarget(s, itemId, stId) {
    var t = RM.itemById(s, itemId);
    if (!t) return null;
    return stId ? storyById(t, stId) : t;
  }
  function setFlag(itemId, stId, flag) {
    commit(flag ? 'flag' : 'unflag', function (s) {
      var x = flagTarget(s, itemId, stId);
      if (x) x.flag = flag ? { reason: String(flag.reason || '').trim() } : null;
    });
  }
  function flagDialog(itemId, stId) {
    var x = flagTarget(state, itemId, stId);
    if (!x) return;
    var cur = x.flag ? x.flag.reason : '';
    openModal(
      '<div class="modal" style="width:440px">' +
      '<div class="m-head"><h2>' + (x.flag ? 'Edit flag' : 'Flag') + ' ' + (stId ? '#' + x.num + ' ' + esc(shorten(x.title || '(untitled)', 40)) : '#' + x.num + ' ' + esc(shorten(x.feature || '(untitled)', 40))) + '</h2></div>' +
      '<div class="m-body"><label class="p-lab">Reason (optional)</label>' +
      '<textarea id="flagReason" rows="3" style="width:100%" placeholder="Why does this need attention?">' + esc(cur) + '</textarea></div>' +
      '<div class="m-foot"><button data-m="cancel">Cancel</button>' +
      '<button data-m="ok" class="primary">' + (x.flag ? 'Save' : 'Flag') + '</button></div></div>',
      function (host) {
        var ta = $('#flagReason', host);
        ta.focus();
        $('[data-m=cancel]', host).onclick = closeModal;
        $('[data-m=ok]', host).onclick = function () { var r = ta.value; closeModal(); setFlag(itemId, stId, { reason: r }); };
        ta.addEventListener('keydown', function (e) { if (e.key === 'Enter' && (e.metaKey || e.ctrlKey)) $('[data-m=ok]', host).click(); });
      });
  }
  // one item or story to its earliest dependency- and capacity-valid slot
  function placeEntry(itemId, storyId) {
    if (!state.meta.capacityEnabled) return null;
    var it = RM.itemById(state, itemId);
    if (!it) return null;
    var st = storyId ? storyById(it, storyId) : null;
    if (storyId && !st) return null;
    return { icon: 'zap', label: 'Place at earliest slot',
      disabled: !!it.locked || !!it.done || !!(st && st.done), fn: function () {
      var r = RM.placeUnit(state, itemId, storyId || null, snapOpts());
      // a partial placement still lands what it could — keep it, and say why
      if (r.changed) {
        if (autoOrder) RM.sortItemsByStart(r.state); // commit never re-sorts on its own
        replaceState('place', r.state);
        toast('Placed at the earliest slot' + (r.note ? ' — ' + r.note : ''));
        return;
      }
      toast(r.note || 'Already at its earliest slot', r.note ? 'err' : undefined);
    } };
  }
  // Exclude from / Include in the Auto timeline — features and stories alike
  // (on a feature it clears Lock: the two are mutually exclusive)
  function noAutoEntry(itemId, stId) {
    var x = flagTarget(state, itemId, stId);
    if (!x) return null;
    var on = !!x.noAuto;
    return { icon: on ? 'zap' : 'zap-off', label: on ? 'Include in Auto timeline' : 'Exclude from Auto timeline',
      title: on ? null : 'Auto timeline and ⚡ leave it where it sits',
      fn: function () {
        commit(on ? 'include in auto' : 'exclude from auto', function (s) {
          var t = flagTarget(s, itemId, stId);
          if (t) RM.setNoAuto(t, !on);
        });
      } };
  }
  function flagMenuEntries(itemId, stId) {
    var x = flagTarget(state, itemId, stId);
    if (!x) return [];
    return x.flag
      ? [{ icon: 'flag', label: 'Edit flag…', fn: function () { flagDialog(itemId, stId); } },
         { icon: 'flag-off', label: 'Unflag', fn: function () { setFlag(itemId, stId, null); } }]
      : [{ icon: 'flag', label: 'Flag…', fn: function () { flagDialog(itemId, stId); } }];
  }
  function duplicateStory(itemId, stId) {
    commit('duplicate story', function (s) {
      var t = RM.itemById(s, itemId);
      var st = t && storyById(t, stId);
      if (!st) return;
      var copy = RM.clone(st);
      copy.id = RM.uid('s');
      copy.num = RM.nextNum(s);
      copy.title = (st.title || 'Story') + ' (copy)';
      copy.jiraKey = '';
      t.stories.splice(t.stories.indexOf(st) + 1, 0, copy);
      selectedId = itemId;
      selStory = copy.id;
    });
  }

  function deleteItemConfirm(itemId) {
    var it = RM.itemById(state, itemId);
    if (!it) return;
    confirmBox('Delete #' + it.num + '?', esc(it.feature) +
      '<br><br>Items depending on it will keep a dangling reference (flagged by validation).',
      'Delete', function () {
        commit('delete', function (s) {
          s.items = s.items.filter(function (x) { return x.id !== itemId; });
          if (selectedId === itemId) selectedId = null;
        });
      }, true);
  }

  function storyById(it, stId) {
    return (it.stories || []).filter(function (s) { return s.id === stId; })[0] || null;
  }

  // the pointer marks the MIDDLE of the landing slot — the placed bar
  // centers on the cursor, then snaps
  function placeDurFor(it) {
    if (!it) return RM.sprintDays(state.meta);
    if (it.milestone) return 1;
    return it.durDays != null ? it.durDays : (RM.effortDays(state, it) || RM.sprintDays(state.meta));
  }
  function centeredDayAt(clientX, dur, kind) {
    return Math.max(0, snapTo(Math.round(laneDayAt(clientX) - dur / 2), kind));
  }

  // give a story its own little timeline where the pointer sits on its lane
  function placeStoryAt(itemId, stId, clientX) {
    var cur = storyById(RM.itemById(state, itemId) || {}, stId);
    if (!cur || cur.startDay != null) return;
    placedAt = Date.now(); placedKey = 'st:' + stId;
    // a story defaults to one sprint
    var day = centeredDayAt(clientX, RM.sprintDays(state.meta), 'story');
    commit('place story', function (s) {
      var st = storyById(RM.itemById(s, itemId), stId);
      if (!st) return;
      st.startDay = day;
      st.durDays = RM.stretchSpan(s.meta, day, RM.sprintDays(s.meta));
    });
  }

  // schedule an unscheduled item where the pointer sits on its lane
  var placedAt = 0, placedKey = null; // swallow the trailing native dblclick after a manual place
  function placeItemAt(itemId, clientX) {
    placedAt = Date.now(); placedKey = itemId;
    var it0 = RM.itemById(state, itemId);
    var pDay = centeredDayAt(clientX, placeDurFor(it0), 'feature');
    commit('place item', function (s2) {
      var t2 = RM.itemById(s2, itemId);
      t2.startDay = pDay;
      // a preset duration wins; else the size estimate; else one sprint
      t2.durDays = t2.milestone ? 0
        : (t2.durDays != null ? t2.durDays
          : RM.stretchSpan(s2.meta, pDay, RM.effortDays(s2, t2) || RM.sprintDays(s2.meta)));
      if (autoOrder) RM.sortItemsByStart(s2);
    });
    select(itemId);
  }

  // hovering an unscheduled row's empty lane previews the landing slot and
  // floats a hint right above it (double-click places the item there)
  var placeGhost = null;
  function hidePlaceGhost() {
    if (placeGhost) {
      if (placeGhost.parentNode) placeGhost.parentNode.removeChild(placeGhost);
      placeGhost = null;
      dragTip.hidden = true;
    }
  }
  // hovering a scheduled row whose bar is scrolled out of view floats a
  // pointer at the near edge of the lane: "← title" or "title →"
  var offTag = null, offRow = null;
  function hideOffTag() {
    if (offTag && offTag.parentNode) offTag.parentNode.removeChild(offTag);
    offTag = null;
    offRow = null;
  }
  // the row's bar in lane px ({x0, x1, title}), or null when it has none
  function laneSpanOf(row) {
    var it = RM.itemById(state, row.dataset.id);
    if (!it) return null;
    var st = row.dataset.story ? storyById(it, row.dataset.story) : null;
    if (row.dataset.story && !st) return null;
    if (st && st.startDay != null && st.durDays != null) {
      return { x0: st.startDay * dayPx(), x1: st.startDay * dayPx() + Math.max(6, st.durDays * dayPx()), title: st.title };
    }
    // a story without its own timeline rides along with the feature
    if (!isScheduled(it)) return null;
    var w = it.milestone ? 1 : Math.max(1, it.durDays + (it.riskDays || 0));
    return { x0: it.startDay * dayPx(), x1: (it.startDay + w) * dayPx(), title: st ? st.title : it.feature };
  }
  function showOffTag(row) {
    var sp = laneSpanOf(row);
    var vis0 = board.scrollLeft, vis1 = board.scrollLeft + Math.max(80, board.clientWidth - leftW());
    var dir = !sp ? null : sp.x1 < vis0 ? 'left' : sp.x0 > vis1 ? 'right' : null;
    if (!dir) { hideOffTag(); return; }
    var lane = row.querySelector('.row-lane');
    if (!lane) { hideOffTag(); return; }
    if (!offTag) offTag = document.createElement('div');
    offRow = row;
    offTag.className = 'lane-off ' + dir;
    offTag.innerHTML = '<span class="lo-arrow">' + (dir === 'left' ? '\u2190' : '\u2192') + '</span>' +
      '<span class="lo-text">' + esc(sp.title || '(untitled)') + '</span>';
    if (offTag.parentNode !== lane) lane.appendChild(offTag);
    offTag.style.left = (dir === 'left' ? vis0 + 8 : vis1 - 8 - offTag.offsetWidth) + 'px';
  }
  board.addEventListener('scroll', function () {
    if (!offRow) return;
    if (document.body.contains(offRow)) showOffTag(offRow); else hideOffTag();
  }, { passive: true });
  rowsEl.addEventListener('pointermove', function (e) {
    if (view !== 'planning' || drag) { hidePlaceGhost(); hideOffTag(); return; }
    var offHit = e.target.closest('.row-lane') && !e.target.closest('.row.story-add')
      ? (e.target.closest('.row.story[data-story]') || e.target.closest('.row.item')) : null;
    if (offHit) showOffTag(offHit); else hideOffTag();
    if (!e.target.closest('.row-lane') ||
      e.target.closest('[data-bar],[data-stbar],[data-ghost],.ghost-pill,.port,.ph-row-bar,.bar.cmp')) { hidePlaceGhost(); return; }
    var stRow = e.target.closest('.row.story[data-story]');
    var itRow = e.target.closest('.row.item');
    var target = null, dur = 5, ms = false;
    if (stRow) {
      var stH = storyById(RM.itemById(state, stRow.dataset.id) || {}, stRow.dataset.story);
      if (stH && stH.startDay == null) { target = stRow; dur = RM.sprintDays(state.meta); }
    } else if (itRow) {
      var itH = RM.itemById(state, itRow.dataset.id);
      if (itH && !isScheduled(itH)) { target = itRow; dur = placeDurFor(itH); ms = !!itH.milestone; }
    }
    if (!target) { hidePlaceGhost(); return; }
    var day = centeredDayAt(e.clientX, dur, stRow ? 'story' : 'feature');
    var lane = target.querySelector('.row-lane');
    if (!placeGhost) {
      placeGhost = document.createElement('div');
      placeGhost.className = 'place-ghost';
    }
    placeGhost.classList.toggle('ms', ms);
    RM.MS_STYLES.forEach(function (sn) { placeGhost.classList.remove(sn); });
    if (ms) placeGhost.classList.add(RM.msStyleOf(itH));
    if (placeGhost.parentNode !== lane) lane.appendChild(placeGhost);
    placeGhost.style.left = (day * dayPx()) + 'px';
    placeGhost.style.width = Math.max(8, dur * dayPx()) + 'px';
    var lr = target.getBoundingClientRect();
    dragTip.hidden = false;
    dragTip.innerHTML = '<b>Double-click to place</b> · ' + esc(RM.fmtShort(RM.dayToDate(state.meta, day)));
    dragTip.style.left = (e.clientX + 14) + 'px';
    dragTip.style.top = (lr.top - 30) + 'px';
  });
  rowsEl.addEventListener('pointerleave', function () { hidePlaceGhost(); hideOffTag(); });
  rowsEl.addEventListener('pointerdown', function () { hidePlaceGhost(); hideOffTag(); });

  function justPlaced(key) { return placedKey === key && Date.now() - placedAt < 500; }
  // double-click a left-pane title (or Rename… in the row menu) to edit it in
  // place. The row is re-queried by id: the click that came first selected it
  // and re-rendered the pane, so the clicked node may already be detached.
  function plStartRename(itemId, stId) {
    var rowEl = rowsEl.querySelector(stId
      ? '.row.story[data-story="' + stId + '"]'
      : '.row.item[data-id="' + itemId + '"]');
    if (!rowEl) return;
    var span = rowEl.querySelector(stId ? '.st-title-txt' : '.r-name-txt');
    if (!span || span.tagName === 'INPUT') return;
    var blank = span.classList.contains('r-ph');
    startInlineEdit(span, function (val) {
      if (stId) {
        commit('rename story', function (s) {
          var st = storyById(RM.itemById(s, itemId) || {}, stId);
          if (st) st.title = val;
        });
      } else {
        commit('rename', function (s) {
          var x = RM.itemById(s, itemId);
          if (x) x.feature = val;
        });
      }
    });
    var inp = rowEl.querySelector('input.st-add-input');
    if (inp) {
      if (!stId) inp.dataset.rowname = '';
      inp.className = 'st-add-input ' + (stId ? 'st-title' : 'r-name');
      if (blank) inp.value = '';
      inp.select();
    }
  }
  // the Rename… entry a titled row's menu opens with
  function plRenameEntry(itemId, stId) {
    return { icon: 'pencil', label: 'Rename…', fn: function () { plStartRename(itemId, stId); } };
  }
  rowsEl.addEventListener('dblclick', function (e) {
    // double-click in a text field selects a word — never steal its focus
    if (e.target.closest('input,textarea,[contenteditable="true"]')) return;
    var titleEl = e.target.closest('.r-name-txt,.st-title-txt');
    if (titleEl) {
      var tRow = titleEl.closest('.row');
      if (tRow) {
        e.preventDefault();
        plStartRename(tRow.dataset.id, tRow.dataset.story || null);
        return;
      }
    }
    var stRowEl = e.target.closest('.row.story[data-story]');
    if (stRowEl) {
      if (view === 'planning' && e.target.closest('.row-lane') && !e.target.closest('[data-stbar]') &&
        !justPlaced('st:' + stRowEl.dataset.story)) {
        placeStoryAt(stRowEl.dataset.id, stRowEl.dataset.story, e.clientX);
      }
      return;
    }
    var rowEl = e.target.closest('.row.item');
    if (!rowEl) return;
    var itemId = rowEl.dataset.id;
    if (justPlaced(itemId)) return;
    var it = RM.itemById(state, itemId);
    // double-click empty lane space on an unscheduled row places the item
    // (usually already handled by the manual click counter — this is the
    // fallback when the first click didn't trigger a re-render)
    if (view === 'planning' && it && !isScheduled(it) &&
      e.target.closest('.row-lane') &&
      !e.target.closest('[data-bar],[data-ghost],.ghost-pill,.port,.sc-row')) {
      placeItemAt(itemId, e.clientX);
      return;
    }
    select(itemId);
    var nameInput = $('#panel .p-name');
    if (nameInput) { nameInput.focus(); nameInput.select(); }
  });

  function startInlineEdit(span, onDone) {
    var input = document.createElement('input');
    input.value = span.textContent;
    input.className = 'st-add-input';
    input.style.flex = '1';
    span.replaceWith(input);
    input.focus(); input.select();
    var done = false;
    function finish(saveIt) {
      if (done) return; done = true;
      if (saveIt && input.value.trim()) onDone(input.value.trim()); else render();
    }
    input.addEventListener('blur', function () { finish(true); });
    input.addEventListener('keydown', function (ev) {
      if (ev.key === 'Enter') finish(true);
      if (ev.key === 'Escape') finish(false);
    });
  }

  // scoping cells auto-grow with their content; every textarea in a row is
  // stretched to the row's tallest cell so there is no dead whitespace
  function autoGrowRow(scRow) {
    var tas = $$('.sc-edit', scRow);
    // story rows are half-height features (2× the plain story row)
    var max = scRow.closest('.row.story') ? 46 : 66;
    tas.forEach(function (t) {
      t.style.height = 'auto';
      max = Math.max(max, t.scrollHeight);
    });
    tas.forEach(function (t) { t.style.height = max + 'px'; });
  }
  function autoGrowScope() {
    $$('.sc-row', rowsEl).forEach(autoGrowRow);
  }
  rowsEl.addEventListener('input', function (e) {
    if (e.target.classList && e.target.classList.contains('sc-edit')) {
      autoGrowRow(e.target.closest('.sc-row'));
    }
  });
  // clicking cell padding focuses its textarea
  rowsEl.addEventListener('mousedown', function (e) {
    var cell = e.target.closest('.sc-cell:not(.sc-fix)');
    if (!cell || e.target.classList.contains('sc-edit')) return;
    var ta = cell.querySelector('.sc-edit');
    if (ta) {
      e.preventDefault();
      ta.focus();
      if (ta.setSelectionRange) {
        ta.setSelectionRange(ta.value.length, ta.value.length);
      } else {
        // rich (contenteditable) cell: caret to the end
        var rg = document.createRange();
        rg.selectNodeContents(ta);
        rg.collapse(false);
        var sel = window.getSelection();
        sel.removeAllRanges();
        sel.addRange(rg);
      }
    }
  });

  // Tab / ⇧Tab hop cell-to-cell across the scoping grid (the rich
  // Description cell would otherwise swallow Tab entirely)
  rowsEl.addEventListener('keydown', function (e) {
    if (e.key !== 'Tab') return;
    var ed = e.target.closest && e.target.closest('.sc-edit');
    if (!ed) return;
    var edits = $$('.sc-edit', rowsEl);
    var i = edits.indexOf(ed);
    if (i === -1) return;
    var next = edits[i + (e.shiftKey ? -1 : 1)];
    if (!next) return;
    e.preventDefault();
    next.focus();
    if (next.setSelectionRange) next.setSelectionRange(0, next.value.length);
  });

  // floating B / I / list toolbar for the rich Description cells (the panel
  // editor has its own inline bar; scoping cells get this shared one)
  var scFmtBar = null;
  function ensureScFmtBar() {
    if (scFmtBar) return scFmtBar;
    scFmtBar = document.createElement('div');
    scFmtBar.id = 'scFmtBar';
    scFmtBar.hidden = true;
    scFmtBar.innerHTML =
      '<button type="button" tabindex="-1" data-scfmt="bold" title="Bold"><b>B</b></button>' +
      '<button type="button" tabindex="-1" data-scfmt="italic" title="Italic"><i>I</i></button>' +
      '<button type="button" tabindex="-1" data-scfmt="insertUnorderedList" title="Bullet list"><i data-lucide="list"></i></button>' +
      '<button type="button" tabindex="-1" data-scfmt="insertOrderedList" title="Numbered list"><i data-lucide="list-ordered"></i></button>';
    // pointerdown is swallowed so the editor keeps focus and selection
    scFmtBar.addEventListener('pointerdown', function (e) { e.preventDefault(); });
    scFmtBar.addEventListener('click', function (e) {
      var b = e.target.closest('[data-scfmt]');
      if (b) document.execCommand(b.dataset.scfmt);
    });
    document.body.appendChild(scFmtBar);
    if (window.lucide) lucide.createIcons();
    return scFmtBar;
  }
  var scFmtTarget = null; // the editor the bar is anchored to
  function placeScFmtBar() {
    if (!scFmtTarget || !scFmtBar || scFmtBar.hidden) return;
    var r = scFmtTarget.getBoundingClientRect();
    scFmtBar.style.left = Math.max(4, r.left) + 'px';
    scFmtBar.style.top = Math.max(4, r.top - 32) + 'px';
  }
  function hideScFmtBar() {
    if (scFmtBar) scFmtBar.hidden = true;
    scFmtTarget = null;
  }
  rowsEl.addEventListener('focusin', function (e) {
    if (!e.target.classList || !e.target.classList.contains('sc-rich')) return;
    ensureScFmtBar().hidden = false;
    scFmtTarget = e.target;
    placeScFmtBar();
  });
  rowsEl.addEventListener('focusout', function (e) {
    if (e.target.classList && e.target.classList.contains('sc-rich') && scFmtBar) hideScFmtBar();
  });
  // focusout doesn't fire when a re-render swaps the editor out from under the
  // bar, and view switches leave it orphaned — close whenever focus lands
  // anywhere that isn't the editor it's anchored to, or the window blurs
  document.addEventListener('focusin', function (e) {
    if (!scFmtBar || scFmtBar.hidden || !scFmtTarget) return;
    if (e.target !== scFmtTarget && !scFmtBar.contains(e.target)) hideScFmtBar();
  });
  window.addEventListener('blur', hideScFmtBar);
  // the bar rides along when the board scrolls under it
  board.addEventListener('scroll', placeScFmtBar, { passive: true });
  window.addEventListener('resize', placeScFmtBar);

  // scoping title cell: plain text, commits on blur; Enter = done
  rowsEl.addEventListener('focusout', function (e) {
    if (!e.target.classList || !e.target.classList.contains('sc-name')) return;
    var rowEl0 = e.target.closest('.row');
    var rid0 = rowEl0 && rowEl0.dataset.id;
    if (!rid0) return;
    var nv0 = e.target.textContent.replace(/\s+/g, ' ').trim();
    var cur0 = (RM.itemById(state, rid0) || {}).feature || '';
    if (nv0 === cur0) return;
    commit('rename', function (s2) { RM.itemById(s2, rid0).feature = nv0; });
  });
  rowsEl.addEventListener('keydown', function (e) {
    if (e.key === 'Enter' && e.target.classList && e.target.classList.contains('sc-name')) {
      e.preventDefault();
      e.target.blur();
    }
  });

  // scoping story title: same in-place rename, committed to the story
  rowsEl.addEventListener('focusout', function (e) {
    if (!e.target.classList || !e.target.classList.contains('st-name')) return;
    var stRow0 = e.target.closest('.row.story[data-story]');
    if (!stRow0) return;
    var pid0 = stRow0.dataset.id, sid0 = stRow0.dataset.story;
    var nv1 = e.target.textContent.replace(/\s+/g, ' ').trim();
    var cur1 = (storyById(RM.itemById(state, pid0) || {}, sid0) || {}).title || '';
    if (nv1 === cur1) return;
    commit('rename story', function (s2) {
      var st2 = storyById(RM.itemById(s2, pid0) || {}, sid0);
      if (st2) st2.title = nv1;
    });
  });
  rowsEl.addEventListener('keydown', function (e) {
    if (e.key === 'Enter' && e.target.classList && e.target.classList.contains('st-name')) {
      e.preventDefault();
      e.target.blur();
    }
  });

  // story scoping cells commit on blur too
  rowsEl.addEventListener('focusout', function (e) {
    var sf = e.target.dataset && e.target.dataset.stscope;
    if (!sf) return;
    if (e.target.getAttribute('contenteditable') !== 'true') return;
    var srowEl = e.target.closest('.row');
    var sItemId = srowEl && srowEl.dataset.id;
    var sStId = srowEl && srowEl.dataset.story;
    if (!sItemId || !sStId) return;
    var sIt = RM.itemById(state, sItemId);
    var sSt = sIt && storyById(sIt, sStId);
    if (!sSt) return;
    var sVal = sanitizeHtml(e.target.innerHTML);
    var sCur = RM.storyScopeValue(sSt, sf);
    if (sVal === sCur || sVal === displayHtml(sCur)) return;
    commit('story field', function (s) {
      var st2 = storyById(RM.itemById(s, sItemId) || {}, sStId);
      if (!st2) return;
      RM.setStoryScopeValue(st2, sf, sVal);
    });
  });

  // scoping cells are all rich editors — they commit on blur
  rowsEl.addEventListener('focusout', function (e) {
    var f = e.target.dataset && e.target.dataset.scope;
    if (!f) return;
    if (e.target.getAttribute('contenteditable') !== 'true') return;
    var rowEl = e.target.closest('.row');
    var itemId = rowEl && rowEl.dataset.id;
    if (!itemId) return;
    var val = sanitizeHtml(e.target.innerHTML);
    var prev = RM.scopeValue(RM.itemById(state, itemId), f);
    if (val === prev || val === displayHtml(prev)) return;
    commit('scope ' + f, function (s) {
      var t = RM.itemById(s, itemId);
      if (t) RM.setScopeValue(t, f, val);
    });
  });

  // keyboard parity: Enter/Space on a focusable chip acts like a click
  rowsEl.addEventListener('keydown', function (e) {
    if ((e.key !== 'Enter' && e.key !== ' ') || e.target.closest('input,textarea')) return;
    var actEl = e.target.closest('[data-act],[data-resadd]');
    if (!actEl && e.target.classList.contains('addrow-lab')) actEl = e.target;
    if (!actEl) return;
    e.preventDefault();
    actEl.dispatchEvent(new MouseEvent('click', { bubbles: true }));
  });

  // story quick-add inputs (change-committed)
  rowsEl.addEventListener('keydown', function (e) {
    if (e.target.dataset && e.target.dataset.act === 'st-add' && e.key === 'Enter') {
      var rowEl = e.target.closest('.row');
      var itemId = rowEl.dataset.id;
      var v = e.target.value.trim();
      if (!v) return;
      commit('add story', function (s) {
        RM.itemById(s, itemId).stories.push({ id: RM.uid('s'), title: v, done: false, num: RM.nextNum(s) });
      });
      requestAnimationFrame(function () {
        var again = rowsEl.querySelector('.row.story-add[data-id="' + itemId + '"] .st-add-input');
        if (again) again.focus();
      });
    }
  });

  // ------------------------------------------------------------ dragging
  var dragConsumedClick = false;
  var dragEndAt = 0; // secondary guard for listeners that run after the flag is consumed
  var dragTip = $('#dragTip');

  rowsEl.addEventListener('pointerdown', function (e) {
    if (e.button !== 0) return;
    var portEl = e.target.closest('[data-port]');
    var barEl = e.target.closest('[data-bar]');
    var ghostEl = e.target.closest('[data-ghost]');
    var leftEl = e.target.closest('.row.item .row-left');
    var handle = e.target.closest('[data-act="bh-l"],[data-act="bh-r"]');

    if (portEl && barEl) {
      startPortDrag(e, barEl.dataset.bar, portEl.dataset.port);
      return;
    }
    var stBarEl = e.target.closest('[data-stbar]');
    if (portEl && stBarEl) {
      startPortDrag(e, stBarEl.dataset.stbar, portEl.dataset.port, true);
      return;
    }
    if (stBarEl) {
      var sh = e.target.closest('[data-act="sh-l"],[data-act="sh-r"]');
      startStoryBarDrag(e, stBarEl.dataset.id, stBarEl.dataset.stbar,
        sh ? (sh.dataset.act === 'sh-l' ? 'resize-l' : 'resize-r') : 'move', stBarEl);
      return;
    }
    var stLeftEl = e.target.closest('.row.story[data-story] .row-left');
    if (stLeftEl && !e.target.closest('input,button,textarea,select,[contenteditable="true"]')) {
      // stories drag too: reorder within their feature or move to another
      var stRowD = stLeftEl.closest('.row.story');
      drag = { kind: 'strow', itemId: stRowD.dataset.id, stId: stRowD.dataset.story,
        x0: e.clientX, y0: e.clientY, moved: false, indicator: null };
      e.preventDefault();
      return;
    }
    if (leftEl && !e.target.closest('input,button,textarea,select,[contenteditable="true"]')) {
      // any part of the left pane drags the row (chips still click if not moved)
      startRowDrag(e, leftEl.closest('.row.item').dataset.id);
      return;
    }
    if (ghostEl) {
      startGhostDrag(e, ghostEl.dataset.ghost);
      return;
    }
    if (barEl) {
      var it = RM.itemById(state, barEl.dataset.bar);
      if (!it || it.locked) return;
      var mode = handle ? (handle.dataset.act === 'bh-l' ? 'resize-l' : 'resize-r') : 'move';
      startBarDrag(e, it.id, mode, barEl);
      return;
    }
  });

  // drag from a bar's edge circle to another bar/row to create a dependency.
  // left circle (in) = this item depends ON the target; right circle (out) =
  // this item is a dependency FOR the target. Esc/Delete cancels mid-draw.
  function startPortDrag(e, itemId, port, story) {
    drag = {
      kind: 'port', itemId: itemId, port: port, story: !!story,
      x0: e.clientX, y0: e.clientY,
      moved: true, lastE: e // drawing starts on mousedown, not after a threshold
    };
    startAutoScroll();
    var svg = $('#arrows');
    svg.setAttribute('width', grid.scrollWidth);
    svg.setAttribute('height', grid.scrollHeight);
    portDragMove(e);
    e.preventDefault();
    e.stopPropagation();
  }
  function cancelPortDrag() {
    if (!drag || drag.kind !== 'port') return;
    drag = null;
    $('#tempLink').setAttribute('hidden', '');
    $$('.bar.link-target,.st-bar.link-target').forEach(function (el) { el.classList.remove('link-target'); });
    dragConsumedClick = true;
  }
  function portDragMove(e) {
    if (drag.story) { storyPortDragMove(e); return; }
    var a = barRect(drag.itemId);
    if (!a) return;
    var g = grid.getBoundingClientRect();
    var temp = $('#tempLink');
    temp.removeAttribute('hidden');
    var sx = drag.port === 'out' ? a.right + 1 : a.left - 3;

    // resolve the hovered ROW (not just the bar) as the target
    $$('.bar.link-target').forEach(function (el) { el.classList.remove('link-target'); });
    var under = document.elementFromPoint(e.clientX, e.clientY);
    var tb = under && under.closest && under.closest('[data-bar],[data-ghost],.row.item');
    var tid = tb && (tb.dataset.bar || tb.dataset.ghost || tb.dataset.id);
    if (tid && tid !== drag.itemId) drag.targetId = tid;
    else drag.targetId = null;

    // snap the line's end to the target's bar edge while hovering a row
    var ex = e.clientX - g.left, ey = e.clientY - g.top;
    if (drag.targetId) {
      var tEl = rowsEl.querySelector('[data-bar="' + drag.targetId + '"]');
      if (tEl) {
        tEl.classList.add('link-target');
        var b = barRect(drag.targetId);
        if (b) {
          ex = drag.port === 'out' ? b.left - 3 : b.right + 1;
          ey = b.cy;
        }
      }
    }
    temp.setAttribute('d', curvePath(sx, a.cy, ex, ey));
  }
  // the story flavour: only story bars and story rows are targets, and a
  // feature under the pointer is politely refused on drop
  function storyPortDragMove(e) {
    var a = storyBarRect(drag.itemId);
    drag.targetStory = null; // never keep a target from a previous move
    drag.overFeature = false;
    if (!a) return;
    var g = grid.getBoundingClientRect();
    var temp = $('#tempLink');
    temp.removeAttribute('hidden');
    var sx = drag.port === 'out' ? a.right + 1 : a.left - 3;

    $$('.st-bar.link-target').forEach(function (el) { el.classList.remove('link-target'); });
    var under = document.elementFromPoint(e.clientX, e.clientY);
    var tb = under && under.closest && under.closest('[data-stbar],.row.story[data-story],[data-bar],.row.item');
    var tid = tb && (tb.dataset.stbar || (tb.classList.contains('story') ? tb.dataset.story : null));
    drag.targetStory = tid && tid !== drag.itemId ? tid : null;
    // a feature bar or row under the pointer is not a target, but remember it
    // so the drop can explain why
    drag.overFeature = !tid && !!tb;

    var ex = e.clientX - g.left, ey = e.clientY - g.top;
    if (drag.targetStory) {
      var tEl = rowsEl.querySelector('[data-stbar="' + drag.targetStory + '"]');
      if (tEl) {
        tEl.classList.add('link-target');
        var b = storyBarRect(drag.targetStory);
        if (b) {
          ex = drag.port === 'out' ? b.left - 3 : b.right + 1;
          ey = b.cy;
        }
      }
    }
    temp.setAttribute('d', curvePath(sx, a.cy, ex, ey));
  }
  function portDragEnd(d) {
    $('#tempLink').setAttribute('hidden', '');
    $$('.bar.link-target,.st-bar.link-target').forEach(function (el) { el.classList.remove('link-target'); });
    if (d.story) { storyPortDragEnd(d); return; }
    if (!d.targetId) return;
    var src = RM.itemById(state, d.itemId);
    var tgt = RM.itemById(state, d.targetId);
    if (!src || !tgt || src.id === tgt.id) return;
    // out-port: target depends on source; in-port: source depends on target
    var depOn = d.port === 'out' ? src : tgt;
    var dependent = d.port === 'out' ? tgt : src;
    if (dependent.deps.indexOf(depOn.num) !== -1) {
      toast('#' + dependent.num + ' already depends on #' + depOn.num);
      return;
    }
    commit('link', function (s) { RM.itemById(s, dependent.id).deps.push(depOn.num); });
    toast('#' + dependent.num + ' now depends on #' + depOn.num);
  }
  function storyPortDragEnd(d) {
    if (!d.targetStory) {
      if (d.overFeature) toast('Stories link to stories — drop on a story bar or row');
      return;
    }
    var src = RM.storyRef(state, d.itemId);
    var tgt = RM.storyRef(state, d.targetStory);
    if (!src || !tgt || src.st.id === tgt.st.id) return;
    // out-port: target depends on source; in-port: source depends on target
    var depOn = d.port === 'out' ? src : tgt;
    var dependent = d.port === 'out' ? tgt : src;
    if ((dependent.st.deps || []).indexOf(depOn.st.num) !== -1) {
      toast('#' + dependent.st.num + ' already depends on #' + depOn.st.num);
      return;
    }
    commit('link stories', function (s) {
      var t = RM.storyRef(s, dependent.st.id);
      if (!t) return;
      t.st.deps = t.st.deps || [];
      t.st.deps.push(depOn.st.num);
    });
    toast('#' + dependent.st.num + ' now depends on #' + depOn.st.num);
  }

  // scoping column resize (widths remembered in the browser); dragging the
  // cell body reorders the column instead
  $('#hdrSprints').addEventListener('pointerdown', function (e) {
    if (view !== 'scoping') return;
    var rz = e.target.closest('[data-rz]');
    if (rz) {
      var field = rz.dataset.rz;
      var col = allScopeCols().filter(function (c) { return c[0] === field; })[0];
      if (!col) return;
      drag = { kind: 'scol', field: field, x0: e.clientX, y0: e.clientY, w0: scopeColWidth(col), moved: false };
      e.preventDefault();
      e.stopPropagation();
      return;
    }
    if (e.target.closest('.sc-hadd,button,input')) return;
    var hc = e.target.closest('.sc-hcell[data-col]');
    if (!hc) return;
    drag = { kind: 'scolmove', key: hc.dataset.col, x0: e.clientX, y0: e.clientY, moved: false, indicator: null };
    e.preventDefault();
    e.stopPropagation();
  });

  function scolMoveMove(e) {
    var cells = $$('#hdrSprints .sc-hcell[data-col]');
    if (!cells.length) return;
    var idx = cells.length;
    for (var i = 0; i < cells.length; i++) {
      var r = cells[i].getBoundingClientRect();
      if (e.clientX < r.left + r.width / 2) { idx = i; break; }
    }
    drag.toIdx = idx;
    if (!drag.indicator) {
      drag.indicator = document.createElement('div');
      drag.indicator.className = 'col-drop-indicator';
      grid.appendChild(drag.indicator);
    }
    var g = grid.getBoundingClientRect();
    var edge = idx < cells.length ? cells[idx].getBoundingClientRect().left
      : cells[cells.length - 1].getBoundingClientRect().right;
    drag.indicator.style.left = (edge - g.left) + 'px';
    document.body.classList.add('dragging-x');
  }

  function scolMoveEnd(d) {
    if (d.toIdx == null) return;
    var keys = allScopeCols().map(function (c) { return c[0]; });
    var from = keys.indexOf(d.key);
    if (from === -1) return;
    var to = d.toIdx > from ? d.toIdx - 1 : d.toIdx;
    if (to === from) return;
    commit('reorder columns', function (s) {
      // reorder within the VISIBLE keys, then write back over the full order
      var vis = keys.slice();
      vis.splice(from, 1);
      vis.splice(to, 0, d.key);
      var visSet = {};
      keys.forEach(function (k) { visSet[k] = true; });
      var vi = 0;
      s.meta.scopeColOrder = s.meta.scopeColOrder.map(function (k) {
        return visSet[k] ? vis[vi++] : k;
      });
    });
  }

  // scoping header: the trailing "+" adds a column (column menus open on
  // right-click — see the document contextmenu handler)
  $('#hdrSprints').addEventListener('click', function (e) {
    if (view !== 'scoping') return;
    var add = e.target.closest('[data-coladd]');
    if (add) {
      var visible = {};
      state.meta.scopeCols.forEach(function (c) { visible[c.key] = true; });
      var items = Object.keys(RM.SCOPE_BUILTIN_LABELS)
        .filter(function (k) { return !visible[k]; })
        .map(function (k) {
          return { icon: 'columns-3', label: RM.SCOPE_BUILTIN_LABELS[k], fn: function () {
            commit('add column', function (s) { RM.addScopeCol(s, null, k); });
          } };
        });
      if (items.length) items.push({ sep: true });
      items.push({ icon: 'plus', label: 'New custom column…', fn: function () { scopeColModal(null); } });
      openDropdown(add, items);
      return;
    }
  });

  // column menu for a (non-fixed) scoping column: reorder / rename / remove;
  // opened by right-clicking the column's header cell
  function scopeColMenuItems(key) {
    var ordKeys = allScopeCols().map(function (c) { return c[0]; });
    var idx = ordKeys.indexOf(key);
    function shiftCol(dir) {
      commit('move column', function (s) {
        var visSet = {};
        ordKeys.forEach(function (k) { visSet[k] = true; });
        var vis = ordKeys.slice();
        vis.splice(idx, 1);
        vis.splice(idx + dir, 0, key);
        var vi = 0;
        s.meta.scopeColOrder = s.meta.scopeColOrder.map(function (k) {
          return visSet[k] ? vis[vi++] : k;
        });
      });
    }
    var items2 = [];
    if (idx > 0) items2.push({ icon: 'arrow-left', label: 'Move left', fn: function () { shiftCol(-1); } });
    if (idx < ordKeys.length - 1) items2.push({ icon: 'arrow-right', label: 'Move right', fn: function () { shiftCol(1); } });
    items2.push({ icon: 'pencil', label: 'Rename…', fn: function () { scopeColModal(key); } });
    items2.push({ sep: true });
    var curScope = (scopeColDef(key) || {}).scope || 'both';
    ['both', 'feature', 'story'].forEach(function (sc) {
      items2.push({ icon: sc === 'both' ? 'layers' : sc === 'feature' ? 'square' : 'list-todo',
        label: esc(scopeScopeLabel(sc)), checked: curScope === sc, fn: function () {
          commit('column scope', function (s) { RM.setScopeColScope(s, key, sc); });
        } });
    });
    items2.push({ sep: true });
    items2.push({ icon: 'trash-2', label: 'Remove column', fn: function () {
      commit('remove column', function (s) { RM.removeScopeCol(s, key); });
    } });
    return items2;
  }

  // create (key == null) or rename a scoping column (built-ins included —
  // clearing the label restores a built-in's canonical name)
  function scopeColModal(key) {
    var cur = '';
    if (key) {
      state.meta.scopeCols.forEach(function (c) {
        if (c.key === key) cur = RM.scopeColLabel(c);
      });
    }
    openModal(
      '<div class="modal" style="width:360px">' +
      '<div class="m-head"><h2>' + (key ? 'Rename column' : 'New column') + '</h2>' +
      '<button class="p-close" data-m="x"><i data-lucide="x"></i></button></div>' +
      '<div class="m-body"><div class="m-sec"><label>Label</label>' +
      '<input id="colName" style="width:100%" value="' + esc(cur) + '" placeholder="Column name"></div></div>' +
      '<div class="m-foot"><button data-m="x2">Cancel</button>' +
      '<button id="colSave" class="primary">' + (key ? 'Save' : 'Add') + '</button></div></div>',
      function (host) {
        $('[data-m=x]', host).onclick = closeModal;
        $('[data-m=x2]', host).onclick = closeModal;
        function save() {
          var v = $('#colName', host).value.trim();
          if (!v && !key) { closeModal(); return; }
          closeModal();
          commit(key ? 'rename column' : 'add column', function (s) {
            if (key) RM.renameScopeCol(s, key, v);
            else RM.addScopeCol(s, v);
          });
        }
        $('#colSave', host).onclick = save;
        $('#colName', host).addEventListener('keydown', function (ev) {
          if (ev.key === 'Enter') save();
        });
      });
  }
  function scolDragMove(e) {
    var w = Math.max(120, drag.w0 + (e.clientX - drag.x0));
    scopeColW[drag.field] = Math.round(w);
    $$('[data-col="' + drag.field + '"]').forEach(function (el) { el.style.width = scopeColW[drag.field] + 'px'; });
    grid.style.width = 'calc(var(--left-w) + ' + scopeW() + 'px)';
    $('#hdrSprints').style.width = scopeW() + 'px';
  }

  // pan the timeline by dragging empty lane space
  board.addEventListener('pointerdown', function (e) {
    if (e.button !== 0 || drag || view === 'scoping') return;
    if (e.target.closest('[data-bar],[data-ghost],[data-port],.row-left,.hdr-left,button,input,textarea,select,.ghost-pill')) return;
    if (!e.target.closest('#grid')) return;
    if (e.shiftKey && view === 'planning') {
      // shift-drag on empty lane: rubber-band select the bars it touches
      drag = { kind: 'marquee', x0: e.clientX, y0: e.clientY, moved: false, box: null, ids: [] };
      e.preventDefault();
      return;
    }
    drag = { kind: 'pan', x0: e.clientX, y0: e.clientY, sl: board.scrollLeft, st: board.scrollTop, moved: false };
  });

  function marqueeMove(e) {
    if (!drag.box) {
      drag.box = document.createElement('div');
      drag.box.id = 'marquee';
      document.body.appendChild(drag.box);
    }
    var x1 = Math.min(drag.x0, e.clientX), x2 = Math.max(drag.x0, e.clientX);
    var y1 = Math.min(drag.y0, e.clientY), y2 = Math.max(drag.y0, e.clientY);
    drag.box.style.left = x1 + 'px';
    drag.box.style.top = y1 + 'px';
    drag.box.style.width = (x2 - x1) + 'px';
    drag.box.style.height = (y2 - y1) + 'px';
    drag.ids = [];
    $$('#rows .bar[data-bar]').forEach(function (b) {
      var r = b.getBoundingClientRect();
      var hit = r.right >= x1 && r.left <= x2 && r.bottom >= y1 && r.top <= y2;
      if (hit) drag.ids.push(b.dataset.bar);
      b.classList.toggle('selected', hit);
    });
  }

  function marqueeEnd(d) {
    flushPanelEdit();
    if (!d.ids.length) { multiSel = null; selectedId = null; selStory = null; render(); return; }
    if (d.ids.length === 1) { select(d.ids[0]); return; }
    multiSel = d.ids;
    if (multiSel.indexOf(selectedId) === -1) selectedId = d.ids[0];
    selStory = null;
    selectedEdge = null;
    render();
  }

  function startBarDrag(e, itemId, mode, barEl) {
    var it = RM.itemById(state, itemId);
    // dragging one bar of a multi-selection moves the whole group with it
    var group = null;
    if (mode === 'move' && multiSel && multiSel.indexOf(itemId) !== -1) {
      group = [];
      multiSel.forEach(function (id) {
        if (id === itemId) return;
        var g = RM.itemById(state, id);
        if (!g || !isScheduled(g) || g.locked) return;
        group.push({ id: id, start0: g.startDay, el: rowsEl.querySelector('[data-bar="' + id + '"]') });
      });
      if (!group.length) group = null;
    }
    drag = {
      kind: 'bar', mode: mode, itemId: itemId, group: group,
      x0: e.clientX, y0: e.clientY,
      start0: it.startDay, dur0: it.durDays, risk0: it.riskDays || 0,
      ripple: e.metaKey || e.ctrlKey, // ⌘-drag pushes dependents along
      el: barEl, moved: false
    };
    barEl.setPointerCapture && barEl.setPointerCapture(e.pointerId);
    e.preventDefault();
  }

  function startGhostDrag(e, itemId) {
    drag = { kind: 'ghost', itemId: itemId, x0: e.clientX, y0: e.clientY, moved: false, preview: null };
    e.preventDefault();
  }

  function startRowDrag(e, itemId) {
    drag = { kind: 'row', itemId: itemId, x0: e.clientX, y0: e.clientY, moved: false, indicator: null };
    e.preventDefault();
  }

  window.addEventListener('pointermove', function (e) {
    if (!drag) return;
    var dx = e.clientX - drag.x0, dy = e.clientY - drag.y0;
    drag.lastE = e;
    if (!drag.moved && Math.abs(dx) < 3 && Math.abs(dy) < 3) return;
    if (!drag.moved) { drag.moved = true; startAutoScroll(); }

    dragMoveDispatch(e, dx);
  });

  function dragMoveDispatch(e, dx) {
    if (dx == null) dx = e.clientX - drag.x0;
    if (drag.kind === 'bar') barDragMove(e, dx);
    else if (drag.kind === 'stbar') stBarDragMove(e, dx);
    else if (drag.kind === 'ghost') ghostDragMove(e);
    else if (drag.kind === 'row') rowDragMove(e);
    else if (drag.kind === 'marquee') marqueeMove(e);
    else if (drag.kind === 'strow') strowMove(e);
    else if (drag.kind === 'phspan') phSpanDragMove(e, dx);
    else if (drag.kind === 'port') portDragMove(e);
    else if (drag.kind === 'scol') scolDragMove(e);
    else if (drag.kind === 'scolmove') scolMoveMove(e);
    else if (drag.kind === 'prcard') prCardDragMove(e);
    else if (drag.kind === 'sprow') sprDragMove(e);
    else if (drag.kind === 'rfill') rfillMove(e);
    else if (drag.kind === 'bfill') bfillMove(e);
    else if (drag.kind === 'rrow') rrowMove(e);
    else if (drag.kind === 'brow') browMove(e);
    else if (drag.kind === 'pan') {
      board.scrollLeft = drag.sl - (e.clientX - drag.x0);
      board.scrollTop = drag.st - (e.clientY - drag.y0);
      document.body.classList.add('panning');
      requestAnimationFrame(renderArrows);
    }
  }

  // dragging near the board's edges scrolls it, so long moves are easy
  var autoScrollOn = false;
  function startAutoScroll() {
    if (autoScrollOn) return;
    autoScrollOn = true;
    (function tick() {
      if (!drag || !drag.moved) { autoScrollOn = false; return; }
      if (drag.kind !== 'pan' && drag.lastE) {
        var r = board.getBoundingClientRect();
        var e = drag.lastE;
        var Z = 56, moved = false;
        if (drag.kind === 'row' || drag.kind === 'port') {
          if (e.clientY < r.top + Z + 40) { board.scrollTop -= Math.ceil((r.top + Z + 40 - e.clientY) / 4); moved = true; }
          else if (e.clientY > r.bottom - Z) { board.scrollTop += Math.ceil((e.clientY - (r.bottom - Z)) / 4); moved = true; }
        }
        if (drag.kind !== 'row') {
          var laneLeft = r.left + leftW();
          if (e.clientX < laneLeft + Z && e.clientX > r.left) { board.scrollLeft -= Math.ceil((laneLeft + Z - e.clientX) / 4); moved = true; }
          else if (e.clientX > r.right - Z) {
            // never scroll past the timeline's own extent — the project end
            // date is a hard wall (stray overflow must not extend the drag)
            var contentW = leftW() + (view === 'scoping' ? scopeW() : state.meta.numWeeks * weekPx);
            var maxSL = Math.max(0, Math.ceil(contentW - board.clientWidth));
            var next = Math.min(maxSL, board.scrollLeft + Math.ceil((e.clientX - (r.right - Z)) / 4));
            if (next > board.scrollLeft) { board.scrollLeft = next; moved = true; }
          }
        }
        if (moved) dragMoveDispatch(drag.lastE);
      }
      requestAnimationFrame(tick);
    })();
  }

  window.addEventListener('pointerup', function (e) {
    if (!drag) return;
    var d = drag; drag = null;
    document.body.classList.remove('dragging-x');
    document.body.classList.remove('panning');
    dragTip.hidden = true;
    if (d.indicator) d.indicator.remove();
    if (d.vIndicator) d.vIndicator.remove();
    if (d.preview) d.preview.remove();
    if (d.box) d.box.remove();
    if (d.kind === 'port') { $('#tempLink').setAttribute('hidden', ''); }
    if (!d.moved) { return; }
    dragConsumedClick = true;
    dragEndAt = Date.now();

    if (d.kind === 'bar') barDragEnd(d, e);
    else if (d.kind === 'stbar') stBarDragEnd(d);
    else if (d.kind === 'ghost') ghostDragEnd(d, e);
    else if (d.kind === 'row') rowDragEnd(d, e);
    else if (d.kind === 'strow') strowEnd(d);
    else if (d.kind === 'marquee') marqueeEnd(d);
    else if (d.kind === 'phspan') phSpanDragEnd(d);
    else if (d.kind === 'port') portDragEnd(d);
    else if (d.kind === 'scol') saveLocal();
    else if (d.kind === 'scolmove') scolMoveEnd(d);
    else if (d.kind === 'prcard') prCardDragEnd(d);
    else if (d.kind === 'sprow') sprDragEnd(d);
    else if (d.kind === 'rfill') rfillEnd(d);
    else if (d.kind === 'bfill') bfillEnd(d);
    else if (d.kind === 'rrow') rrowEnd(d);
    else if (d.kind === 'brow') browEnd(d);
    else if (d.kind === 'pan') requestAnimationFrame(renderArrows);
  });

  function daysFromDx(dx) { return Math.round(dx / dayPx()); }
  // touched items align to the snap grid. Features and stories each have a
  // MODE (day / week / sprint) — the actual grid follows the selected working
  // days, so a week is SPW() slots and week snaps land on week starts
  // regardless of holidays; sprint snaps land on sprint boundaries, which
  // follow the sprint anchor in Setup
  var SNAP_LABELS = { day: 'day', week: 'week', sprint: 'sprint' };
  function snapModeFor(kind) { return kind === 'story' ? snapStory : snapFeat; }
  // core never reads UI prefs: hand the snap modes down explicitly
  function snapOpts() { return { snap: { feature: snapFeat, story: snapStory } }; }
  // a working-day count, rounded up to a whole number of the kind's snap unit
  function snapUpDays(days, kind) { return days == null ? days : RM.snapUpDays(state.meta, days, snapModeFor(kind)); }
  function setSnapMode(kind, mode) { if (kind === 'story') snapStory = mode; else snapFeat = mode; }
  function snapUnit(kind) {
    var mode = snapModeFor(kind);
    return mode === 'day' ? 1 : mode === 'sprint' ? RM.sprintDays(state.meta) : SPW();
  }
  function snapTo(d, kind) {
    var u = snapUnit(kind);
    var off = snapModeFor(kind) === 'sprint' ? ((RM.sprintInfo(state.meta).anchorWeek * SPW()) % u + u) % u : 0;
    return Math.round((d - off) / u) * u + off;
  }

  function barDragMove(e, dx) {
    var it = RM.itemById(state, drag.itemId);
    var dd = daysFromDx(dx);
    var su = e.altKey ? 1 : snapUnit('feature'); // ⌥ ignores the snap grid: any day
    var sn = e.altKey ? function (d) { return d; } : function (d) { return snapTo(d, 'feature'); };
    var ns = drag.start0, nd = drag.dur0, nr = drag.risk0;
    if (drag.mode === 'move') ns = Math.max(0, sn(drag.start0 + dd));
    else if (drag.mode === 'resize-r') {
      // the END edge lands on a grid line; never shorter than one grid step
      nd = sn(drag.start0 + drag.dur0 + dd) - drag.start0;
      if (nd < 1) nd = Math.max(1, su);
    } else if (drag.mode === 'resize-l') {
      ns = Math.max(0, Math.min(sn(drag.start0 + dd), drag.start0 + drag.dur0 - su));
      nd = drag.dur0 + (drag.start0 - ns);
      if (nd < 1) nd = Math.max(1, su);
    }
    drag.el.classList.add('dragging');
    document.body.classList.add('dragging-x');
    drag.el.style.left = (ns * dayPx()) + 'px';
    if (!it.milestone) drag.el.style.width = (Math.max(6, nd * dayPx()) + nr * dayPx()) + 'px';
    drag.ns = ns; drag.nd = nd; drag.nr = nr;
    if (drag.group) {
      var gd = ns - drag.start0;
      drag.group.forEach(function (g) {
        if (g.el) { g.el.classList.add('dragging'); g.el.style.left = (Math.max(0, g.start0 + gd) * dayPx()) + 'px'; }
      });
    }

    // vertical: dragging the bar across other rows reorders / re-phases the
    // item (same drop logic as dragging the row's left pane) — a group drag
    // moves in time only
    if (drag.mode === 'move' && !drag.group) {
      var srcRow = rowsEl.querySelector('.row[data-id="' + drag.itemId + '"]');
      var sr = srcRow && srcRow.getBoundingClientRect();
      drag.vTarget = (sr && (e.clientY < sr.top || e.clientY > sr.bottom)) ? dropTargetAt(e.clientY) : null;
      if (drag.vTarget) {
        if (!drag.vIndicator) {
          drag.vIndicator = document.createElement('div');
          drag.vIndicator.className = 'drop-indicator';
          grid.appendChild(drag.vIndicator);
        }
        var g2 = grid.getBoundingClientRect();
        drag.vIndicator.style.top = (drag.vTarget.y - g2.top) + 'px';
        drag.vIndicator.style.left = board.scrollLeft + 'px';
        drag.vIndicator.style.display = '';
      } else if (drag.vIndicator) drag.vIndicator.style.display = 'none';
    }

    var meta = state.meta;
    dragTip.hidden = false;
    dragTip.style.left = (e.clientX + 14) + 'px';
    dragTip.style.top = (e.clientY - 34) + 'px';
    if (it.milestone) {
      dragTip.innerHTML = '<b>' + RM.fmtShort(RM.dayToDate(meta, ns)) + '</b> · milestone';
    } else {
      var endD = RM.spanEndDate(meta, ns, nd + nr);
      var work = RM.workInSpan(meta, ns, nd);
      var riskWork = RM.workInSpan(meta, ns + nd, nr);
      dragTip.innerHTML = '<b>' + RM.fmtShort(RM.dayToDate(meta, ns)) + '</b> → ' + RM.fmtShort(endD) +
        ' · ' + fmtDays(work) + (it.size ? ' <b>(' + it.size + ')</b>' : '') +
        (nr > 0 ? ' + ' + fmtDays(riskWork) + ' risk' : '');
      if (drag.ripple) dragTip.innerHTML += ' · moves chain';
    }
    if (drag.group) dragTip.innerHTML += ' · ' + (drag.group.length + 1) + ' items';
  }

  function barDragEnd(d) {
    if (d.ns == null) return;
    var vr = d.vTarget ? resolveDrop(d.vTarget, d.itemId) : null;
    var label = d.mode === 'move' ? 'move bar' : 'resize bar';
    var rippleMoved = 0;
    commit(label, function (s) {
      var t = RM.itemById(s, d.itemId);
      var endBefore = t.startDay + t.durDays + (t.riskDays || 0);
      if (d.mode === 'move') RM.shiftStories(t, d.ns - t.startDay);
      t.startDay = d.ns;
      t.durDays = d.nd;
      t.riskDays = d.nr;
      if (d.ripple) {
        // \u2318-drag moves the whole dependent chain rigidly with the bar
        var endDelta = (d.ns + d.nd + d.nr) - endBefore;
        rippleMoved = RM.shiftDependents(s, t.id, endDelta, { rigid: true });
      }
      if (d.group) {
        var gd2 = d.ns - d.start0;
        d.group.forEach(function (g) {
          var gt = RM.itemById(s, g.id);
          if (!gt || gt.startDay == null) return;
          var gns = Math.max(0, g.start0 + gd2);
          RM.shiftStories(gt, gns - gt.startDay);
          gt.startDay = gns;
        });
      }
      if (vr) applyDrop(s, d.itemId, vr);
      if (autoOrder) RM.sortItemsByStart(s);
    });
    if (rippleMoved) toast('Moved ' + rippleMoved + ' chained item(s) along');
  }

  // story bars: move / edge-resize, no ports, no ripple, no auto-order
  function startStoryBarDrag(e, itemId, stId, mode, el) {
    var st = storyById(RM.itemById(state, itemId) || {}, stId);
    if (!st || st.startDay == null) return;
    drag = {
      kind: 'stbar', mode: mode, itemId: itemId, stId: stId,
      x0: e.clientX, y0: e.clientY, start0: st.startDay, dur0: st.durDays,
      el: el, moved: false
    };
    el.setPointerCapture && el.setPointerCapture(e.pointerId);
    e.preventDefault();
  }

  function stBarDragMove(e, dx) {
    var dd = daysFromDx(dx);
    var su = e.altKey ? 1 : snapUnit('story');
    var sn = e.altKey ? function (d) { return d; } : function (d) { return snapTo(d, 'story'); };
    var ns = drag.start0, nd = drag.dur0;
    if (drag.mode === 'move') ns = Math.max(0, sn(drag.start0 + dd));
    else if (drag.mode === 'resize-r') {
      nd = sn(drag.start0 + drag.dur0 + dd) - drag.start0;
      if (nd < 1) nd = Math.max(1, su);
    } else if (drag.mode === 'resize-l') {
      ns = Math.max(0, Math.min(sn(drag.start0 + dd), drag.start0 + drag.dur0 - su));
      nd = drag.dur0 + (drag.start0 - ns);
      if (nd < 1) nd = Math.max(1, su);
    }
    drag.el.classList.add('dragging');
    document.body.classList.add('dragging-x');
    drag.el.style.left = (ns * dayPx()) + 'px';
    drag.el.style.width = Math.max(6, nd * dayPx()) + 'px';
    drag.ns = ns; drag.nd = nd;
    var meta = state.meta;
    dragTip.hidden = false;
    dragTip.style.left = (e.clientX + 14) + 'px';
    dragTip.style.top = (e.clientY - 34) + 'px';
    dragTip.innerHTML = '<b>' + RM.fmtShort(RM.dayToDate(meta, ns)) + '</b> → ' +
      RM.fmtShort(RM.spanEndDate(meta, ns, nd)) + ' · ' + fmtDays(RM.workInSpan(meta, ns, nd));
  }

  function stBarDragEnd(d) {
    if (d.ns == null) return;
    commit(d.mode === 'move' ? 'move story' : 'resize story', function (s) {
      var st = storyById(RM.itemById(s, d.itemId) || {}, d.stId);
      if (st) { st.startDay = d.ns; st.durDays = d.nd; }
    });
  }

  function laneDayAt(clientX) {
    var g = grid.getBoundingClientRect();
    var x = clientX - g.left - parseFloat(getComputedStyle(document.documentElement).getPropertyValue('--left-w'));
    return Math.max(0, Math.round(x / dayPx()));
  }

  function ghostDragMove(e) {
    var it = RM.itemById(state, drag.itemId);
    var day = Math.max(0, e.altKey ? laneDayAt(e.clientX) : snapTo(laneDayAt(e.clientX), 'feature'));
    var dur = it.milestone ? 1
      : (it.durDays != null ? it.durDays : RM.stretchSpan(state.meta, day, RM.effortDays(state, it)));
    if (!drag.preview) {
      var rowEl = rowsEl.querySelector('.row[data-id="' + it.id + '"] .row-lane');
      drag.preview = document.createElement('div');
      drag.preview.className = 'bar preview';
      drag.preview.style.setProperty('--bar-c', '#' + RM.colorForItem(state, it));
      rowEl.appendChild(drag.preview);
    }
    drag.preview.style.left = (day * dayPx()) + 'px';
    drag.preview.style.width = Math.max(6, dur * dayPx()) + 'px';
    drag.day = day; drag.dur = dur;
    dragTip.hidden = false;
    dragTip.style.left = (e.clientX + 14) + 'px';
    dragTip.style.top = (e.clientY - 34) + 'px';
    dragTip.innerHTML = '<b>' + RM.fmtShort(RM.dayToDate(state.meta, day)) + '</b> · ' + fmtDays(RM.effortDays(state, it));
  }

  function ghostDragEnd(d) {
    if (d.day == null) return;
    commit('schedule', function (s) {
      var t = RM.itemById(s, d.itemId);
      t.startDay = d.day;
      t.durDays = t.milestone ? 0 : d.dur;
      t.riskDays = RM.stretchSpan(s.meta, t.startDay + t.durDays, RM.riskEffortDays(s, t));
      if (autoOrder) RM.sortItemsByStart(s);
    });
  }

  function rowDragMove(e) {
    if (!drag.indicator) {
      drag.indicator = document.createElement('div');
      drag.indicator.className = 'drop-indicator';
      grid.appendChild(drag.indicator);
      var srcEl = rowsEl.querySelector('.row[data-id="' + drag.itemId + '"]');
      if (srcEl) srcEl.classList.add('drag-row-ghost');
    }
    var target = dropTargetAt(e.clientY);
    drag.target = target;
    if (target) {
      var g = grid.getBoundingClientRect();
      drag.indicator.style.top = (target.y - g.top) + 'px';
      drag.indicator.style.left = board.scrollLeft + 'px'; // stay under the frozen column
      drag.indicator.style.display = '';
    } else {
      drag.indicator.style.display = 'none';
    }
  }

  function dropTargetAt(clientY) {
    // find insertion point among visible band/epic-band/item rows
    var rowEls = $$('#rows .row').filter(function (el) {
      return el.dataset.kind === 'band' || el.dataset.kind === 'eband' || el.classList.contains('item');
    });
    for (var i = 0; i < rowEls.length; i++) {
      var r = rowEls[i].getBoundingClientRect();
      if (clientY < r.top + r.height / 2) {
        return { beforeEl: rowEls[i], y: r.top, index: i };
      }
    }
    var last = rowEls[rowEls.length - 1];
    if (!last) return null;
    var lr = last.getBoundingClientRect();
    return { beforeEl: null, y: lr.bottom, index: rowEls.length };
  }

  function rowDragEnd(d, e) {
    var srcEl = rowsEl.querySelector('.row[data-id="' + d.itemId + '"]');
    if (srcEl) srcEl.classList.remove('drag-row-ghost');
    var r = d.target && resolveDrop(d.target, d.itemId);
    if (!r) { render(); return; }
    commit('reorder', function (s) { applyDrop(s, d.itemId, r); });
  }

  // story rows drag among story rows: reorder inside a feature, or drop into
  // another feature (before one of its stories, or on its "add story" row)
  function strowMove(e) {
    if (!drag.indicator) {
      drag.indicator = document.createElement('div');
      drag.indicator.className = 'drop-indicator';
      grid.appendChild(drag.indicator);
      var srcEl = rowsEl.querySelector('.row.story[data-story="' + drag.stId + '"]');
      if (srcEl) srcEl.classList.add('drag-row-ghost');
    }
    var els = $$('#rows .row.story');
    var target = null;
    for (var i = 0; i < els.length; i++) {
      var r = els[i].getBoundingClientRect();
      if (e.clientY < r.top + r.height / 2) { target = { beforeEl: els[i], y: r.top }; break; }
    }
    if (!target && els.length) {
      var lr = els[els.length - 1].getBoundingClientRect();
      target = { beforeEl: els[els.length - 1], y: lr.top };
    }
    drag.target = target;
    if (target) {
      var g = grid.getBoundingClientRect();
      drag.indicator.style.top = (target.y - g.top) + 'px';
      drag.indicator.style.left = board.scrollLeft + 'px';
      drag.indicator.style.display = '';
    } else {
      drag.indicator.style.display = 'none';
    }
  }
  function strowEnd(d) {
    var srcEl = rowsEl.querySelector('.row.story[data-story="' + d.stId + '"]');
    if (srcEl) srcEl.classList.remove('drag-row-ghost');
    var t = d.target;
    if (!t || !t.beforeEl) { render(); return; }
    var toItemId = t.beforeEl.dataset.id;
    var beforeStId = t.beforeEl.classList.contains('story-add') ? null : t.beforeEl.dataset.story;
    if (beforeStId === d.stId) { render(); return; }
    commit('move story', function (s) {
      var from = RM.itemById(s, d.itemId);
      var st = from && (from.stories || []).filter(function (x) { return x.id === d.stId; })[0];
      if (!st) return;
      from.stories = from.stories.filter(function (x) { return x.id !== d.stId; });
      var to = RM.itemById(s, toItemId);
      if (!to || to.milestone) { from.stories.push(st); return; }
      var at = to.stories.length;
      if (beforeStId) {
        var bi = to.stories.map(function (x) { return x.id; }).indexOf(beforeStId);
        if (bi !== -1) at = bi;
      }
      to.stories.splice(at, 0, st);
      // the selection follows the moved story to its new feature
      if (selStory === d.stId) selectedId = toItemId;
    });
  }

  // Where does the insertion point sit? Determine governing phase + before-item.
  // Shared by the left-pane row drag and the bar's vertical drag.
  function resolveDrop(target, itemId) {
    var beforeItemId = null, phaseId = null, epicTo, wsTo;
    if (target.beforeEl) {
      if (target.beforeEl.dataset.kind === 'band') {
        // dropping right above a band = end of previous phase; find previous band
        var idx = RM.phaseIndex(state, target.beforeEl.dataset.phase);
        if (idx <= 0) { phaseId = target.beforeEl.dataset.phase; }
        else phaseId = state.phases[idx - 1].id;
      } else if (target.beforeEl.dataset.kind === 'eband') {
        // dropping right above an epic band = end of the previous epic group
        phaseId = target.beforeEl.dataset.phase;
        var seq = $$('#rows .row');
        var pos = seq.indexOf(target.beforeEl);
        for (var k = pos - 1; k >= 0; k--) {
          if (seq[k].classList.contains('item')) {
            var prevIt = RM.itemById(state, seq[k].dataset.id);
            if (prevIt && prevIt.phaseId === phaseId) {
              if (groupEpic) epicTo = prevIt.epic;
              if (groupWs) wsTo = prevIt.workstream;
            }
            break;
          }
          if (seq[k].dataset.kind === 'band') break;
        }
      } else {
        beforeItemId = target.beforeEl.dataset.id;
        var beforeIt = RM.itemById(state, beforeItemId);
        phaseId = beforeIt.phaseId;
        if (groupEpic) epicTo = beforeIt.epic;
        if (groupWs) wsTo = beforeIt.workstream;
      }
    } else {
      phaseId = state.phases[state.phases.length - 1].id;
    }
    if (beforeItemId === itemId) return null;
    return { phaseId: phaseId, beforeItemId: beforeItemId, epicTo: epicTo, wsTo: wsTo };
  }

  function applyDrop(s, itemId, r) {
    var t = RM.itemById(s, itemId);
    if (!t) return;
    s.items = s.items.filter(function (x) { return x.id !== t.id; });
    t.phaseId = r.phaseId;
    if (groupEpic && r.epicTo !== undefined) t.epic = r.epicTo || '';
    if (groupWs && r.wsTo !== undefined) t.workstream = r.wsTo || '';
    var insertAt = s.items.length;
    if (r.beforeItemId) {
      insertAt = s.items.indexOf(RM.itemById(s, r.beforeItemId));
    } else {
      // end of phase: after last item with phaseId (or global end)
      var lastIdx = -1;
      s.items.forEach(function (x, i2) { if (x.phaseId === r.phaseId) lastIdx = i2; });
      insertAt = lastIdx + 1;
      if (lastIdx === -1) insertAt = s.items.length;
    }
    s.items.splice(insertAt, 0, t);
  }

  // ------------------------------------------------------------ topbar: title, tabs, menus
  // the title IS the file name (shown without .xlsx) — renaming the roadmap
  // renames the open file on disk in the desktop shell
  function titleFromFileName(name) {
    return String(name || '').replace(/\.xlsx$/i, '').trim() || 'Roadmap';
  }
  $('#docTitle').addEventListener('change', function (e) {
    var v = titleFromFileName(e.target.value);
    if (v === state.meta.title) { e.target.value = v; return; }
    commit('title', function (s) { s.meta.title = v; });
    e.target.value = state.meta.title;
    if (window.HeadwayDesktop && HeadwayDesktop.renameTo && HeadwayDesktop.currentPath()) {
      HeadwayDesktop.renameTo(saveFileName()).then(function (p) {
        if (p) toast('Renamed file to “' + HeadwayDesktop.basename(p) + '”');
      }, function (err) {
        toast('Could not rename the file: ' + (err && err.message || err), 'err');
      });
    }
  });
  // the title edits in place — click it to rename. It rests readonly so
  // header clicks can still drag the desktop window; the click unlocks it,
  // blur/Enter commits, Esc reverts.
  var titleMeasure = null;
  function sizeTitle() {
    var t = $('#docTitle');
    if (!titleMeasure) titleMeasure = document.createElement('canvas').getContext('2d');
    var w;
    if (titleMeasure) {
      var cs = getComputedStyle(t);
      titleMeasure.font = cs.fontWeight + ' ' + cs.fontSize + ' ' + cs.fontFamily;
      w = Math.ceil(titleMeasure.measureText(t.value || '').width) + 18;
    } else {
      w = (t.value || '').length * 8 + 18; // headless fallback (no canvas 2d)
    }
    var editing = !t.hasAttribute('readonly');
    t.style.width = Math.min(420, Math.max(editing ? 220 : 30, w)) + 'px';
  }
  function setTitleEditing(on) {
    var t = $('#docTitle');
    if (on) t.removeAttribute('readonly');
    else t.setAttribute('readonly', '');
    sizeTitle();
    if (on) { t.focus(); t.select(); }
  }
  $('#docTitle').addEventListener('click', function (e) {
    if (e.target.hasAttribute('readonly')) setTitleEditing(true);
  });
  $('#docTitle').addEventListener('blur', function () {
    setTitleEditing(false);
  });
  $('#docTitle').addEventListener('keydown', function (e) {
    if (e.key === 'Enter') e.target.blur();
    if (e.key === 'Escape') { e.target.value = state.meta.title; e.target.blur(); }
  });
  $('#docTitle').addEventListener('input', sizeTitle);

  var filterTimer = null;
  $('#rowFilter').addEventListener('input', function (e) {
    clearTimeout(filterTimer);
    filterTimer = setTimeout(function () {
      filterText = e.target.value.trim();
      render();
    }, 120);
  });
  $('#rowFilter').addEventListener('keydown', function (e) {
    if (e.key === 'Escape') {
      e.stopPropagation();
      e.target.value = '';
      filterText = '';
      e.target.blur();
      render();
    }
  });

  function switchView(e) {
    var b = e.target.closest('[data-view]');
    if (!b || b.dataset.view === view) return;
    flushPanelEdit();
    view = b.dataset.view;
    // entering Version History always lands on the newest change
    if (view === 'history') { vhSel = null; vhPick = []; vhTab = null; }
    saveLocal();
    render();
    if (view === 'planning') requestAnimationFrame(goToday);
  }
  $('#viewTabs').addEventListener('click', switchView);
  $('#btnSetup').addEventListener('click', switchView);
  $('#btnHistory').addEventListener('click', switchView);

  // ---------------------------------------------------------- version history
  // Timeline of document changes (newest first): who, what, when. Entries are
  // written by commit() and travel WITH the document (saved into the .xlsx).
  function vhTime(t) {
    var d = new Date(t);
    return d.toLocaleDateString(undefined, { month: 'short', day: 'numeric' }) + ' ' +
      d.toLocaleTimeString(undefined, { hour: 'numeric', minute: '2-digit' });
  }
  function vhLabel(label) {
    var s = String(label || 'edit');
    return s.charAt(0).toUpperCase() + s.slice(1);
  }
  // Version History is a full page: the change feed on the left (newest
  // first), the selected change's visual diff on the right — grouped into
  // tabs (Timeline / Scope / Status / Team & Costs / Setup) with old values
  // in red and new values in green. Ticking two entries diffs the document
  // ACROSS them (every change in between, merged).
  var vhSel = null;   // selected entry index into state.history
  var vhPick = [];    // ticked entries for a two-version compare (max 2)
  var vhTab = null;   // active category tab (null = first non-empty)
  var VH_TAB_ORDER = ['timeline', 'scope', 'status', 'budget', 'setup'];
  var VH_TAB_ICONS = { timeline: 'chart-gantt', scope: 'table-properties', status: 'square-kanban', budget: 'wallet', setup: 'settings-2' };
  function vhEntryTitle(en) {
    return vhLabel(en.label) + (en.n > 1 ? ' ×' + en.n : '');
  }
  // mini planning-style timeline for the Timeline tab: one row per item whose
  // schedule changed — the old position in red, the new one in green. Hovering
  // a row lists every field that changed on that item.
  function vtHtml(tlList, ops) {
    if (!tlList || !tlList.length) return '';
    var meta = state.meta;
    var wpx = Math.max(6, Math.min(22, Math.round(820 / meta.numWeeks)));
    var dpx = wpx / RM.slotsOf(meta);
    var laneW = Math.round(meta.numWeeks * wpx);
    var si = RM.sprintInfo(meta);
    var firstB = si.anchorWeek - Math.ceil(si.anchorWeek / si.wps) * si.wps;
    var ticks = '', glines = '';
    for (var bw = Math.max(0, firstB); bw < meta.numWeeks; bw += si.wps) {
      var x = Math.round(bw * wpx);
      glines += '<span class="vt-gridline" style="left:' + x + 'px"></span>';
      if (x + 46 < laneW) {
        ticks += '<span class="vt-tick" style="left:' + x + 'px">' + esc(RM.fmtShort(RM.weekStartDate(meta, bw))) + '</span>';
      }
    }
    var rows = tlList.map(function (t) {
      var opsFor = (ops || []).filter(function (op) { return op[1].indexOf('#' + t.n + ' ') === 0; });
      var tip = opsFor.map(function (op) {
        return op[1].split('— ').pop() + ': ' + (op[2] || '∅') + ' → ' + (op[3] || '∅');
      }).join('\n') || 'Schedule changed';
      function bar(cls, s, d) {
        if (s == null) return '';
        var w = Math.max(4, (d || 1) * dpx);
        return '<span class="vt-bar ' + cls + (t.ms && !d ? ' msb' : '') + '" style="left:' + (s * dpx).toFixed(1) + 'px;width:' + w.toFixed(1) + 'px"></span>';
      }
      return '<div class="vt-row" title="' + esc('#' + t.n + ' ' + t.f + '\n' + tip) + '">' +
        '<div class="vt-left"><span class="r-num">#' + t.n + '</span>' + esc(t.f) + '</div>' +
        '<div class="vt-lane" style="width:' + laneW + 'px">' + glines +
        bar('old', t.s0, t.d0) + bar('new', t.s1, t.d1) + '</div></div>';
    }).join('');
    return '<div class="vt-wrap"><div class="vt-grid">' +
      '<div class="vt-hdr"><div class="vt-left"></div><div class="vt-lane" style="width:' + laneW + 'px">' + ticks + '</div></div>' +
      rows + '</div></div>' +
      '<div class="vt-key"><span class="k"><span class="sw old"></span>before</span>' +
      '<span class="k"><span class="sw new"></span>after</span>' +
      '<span class="k">hover a row for every changed field</span></div>';
  }
  function renderHistoryPage() {
    var host = $('#historyView');
    if (!host) return;
    var h = state.history || [];
    if (vhSel == null || !h[vhSel]) vhSel = h.length ? h.length - 1 : null;
    vhPick = vhPick.filter(function (i) { return h[i]; });
    var listRows = [];
    for (var i = h.length - 1; i >= 0; i--) {
      var en = h[i];
      var who = en.u || 'Unknown user';
      var isSel = vhPick.length === 2 ? vhPick.indexOf(i) !== -1 : i === vhSel;
      listRows.push('<div class="vh-item hv-item' + (isSel ? ' sel' : '') + '" data-vh="' + i + '"><span class="vh-dot"></span>' +
        '<input type="checkbox" class="hv-ck" data-vhck="' + i + '"' + (vhPick.indexOf(i) !== -1 ? ' checked' : '') +
        ' title="Tick two changes to compare across them">' +
        '<span class="avatar sm" title="' + esc(who) + '" style="background:' +
        (en.u ? RM.avatarColor(en.u) : 'var(--ink-3)') + '">' + esc(en.u ? RM.initialsOf(en.u) : '?') + '</span>' +
        '<div class="vh-main"><b>' + esc(who) + '</b> <span class="vh-what">— ' + esc(vhLabel(en.label)) + '</span>' +
        (en.n > 1 ? '<span class="vh-count">×' + en.n + '</span>' : '') + '</div>' +
        '<span class="vh-time" title="' + esc(new Date(en.t).toLocaleString()) + '">' +
        esc(relTime(en.t)) + ' · ' + esc(vhTime(en.t)) + '</span></div>');
    }
    var meNote = userName()
      ? 'Editing as <b>' + esc(userName()) + '</b> — <button data-vh-me>change</button>'
      : 'Your edits record as “Unknown user” — <button data-vh-me>set your name</button>';

    // ---- assemble the diff for the current selection
    var ops = [], tlSel = [], head = '', sub = '', noDetail = '', overflow = 0;
    if (vhPick.length === 2) {
      var lo = Math.min(vhPick[0], vhPick[1]), hi = Math.max(vhPick[0], vhPick[1]);
      for (var k = lo + 1; k <= hi; k++) {
        ops = mergeOps(ops, h[k].d || []);
        tlSel = mergeTl(tlSel, h[k].tl || []);
        overflow += h[k].x || 0;
        if (!h[k].d) noDetail = 'Some of the compared changes predate change tracking — their details aren’t included.';
      }
      head = 'Comparing ' + (hi - lo) + ' change' + (hi - lo > 1 ? 's' : '');
      sub = esc(vhTime(h[lo].t)) + ' (' + esc(vhEntryTitle(h[lo])) + ') → ' + esc(vhTime(h[hi].t)) + ' (' + esc(vhEntryTitle(h[hi])) + ')';
    } else if (vhSel != null) {
      var se = h[vhSel];
      ops = se.d || [];
      tlSel = se.tl || [];
      overflow = se.x || 0;
      head = vhEntryTitle(se);
      sub = '<b>' + esc(se.u || 'Unknown user') + '</b> · ' + esc(relTime(se.t)) + ' · ' + esc(new Date(se.t).toLocaleString());
      if (!se.d) noDetail = 'No detail was recorded for this change — it predates change tracking. New edits record full details.';
    }
    var byCat = {};
    ops.forEach(function (op) { (byCat[op[0]] = byCat[op[0]] || []).push(op); });
    var cats = VH_TAB_ORDER.filter(function (c) { return byCat[c] && byCat[c].length; });
    var activeTab = cats.indexOf(vhTab) !== -1 ? vhTab : cats[0];
    var tabs = cats.map(function (c) {
      return '<button class="hv-tab' + (c === activeTab ? ' on' : '') + '" data-vhtab="' + c + '">' +
        '<i data-lucide="' + VH_TAB_ICONS[c] + '"></i>' + esc(VH_CATS[c] || c) +
        '<span class="hv-tabn">' + byCat[c].length + '</span></button>';
    }).join('');
    var diffRows = (byCat[activeTab] || []).map(function (op) {
      return '<div class="vd-row"><span class="vd-lab">' + esc(op[1]) + '</span><span class="vd-vals">' +
        (op[2] ? '<span class="vd-old">' + esc(op[2]) + '</span>' : '') +
        (op[2] && op[3] ? '<span class="vd-arr">→</span>' : '') +
        (op[3] ? '<span class="vd-new">' + esc(op[3]) + '</span>' : '') +
        '</span></div>';
    }).join('');
    var detail;
    if (!h.length) {
      detail = '<div class="vh-empty">No changes recorded yet — edits made from now on appear here, newest first.</div>';
    } else {
      // the Timeline tab leads with a mini planning-style timeline: every
      // moved item as a row — before in red, after in green
      var tlVisual = activeTab === 'timeline' ? vtHtml(tlSel, ops) : '';
      detail =
        '<div class="hv-dhead"><h2>' + esc(head) + '</h2><div class="hv-dsub">' + sub + '</div></div>' +
        (noDetail ? '<div class="hv-nodetail">' + esc(noDetail) + '</div>' : '') +
        (cats.length
          ? '<div class="hv-tabs">' + tabs + '</div>' + tlVisual + '<div class="hv-diff">' + diffRows +
            (overflow ? '<div class="hv-more">…and ' + overflow + ' more change(s) not kept in detail</div>' : '') + '</div>'
          : (noDetail ? '' : '<div class="vh-empty">Nothing to show — this change left no recorded field differences.</div>'));
    }
    host.innerHTML =
      '<div class="hv-wrap">' +
      '<aside class="hv-side">' +
      '<div class="hv-sidehead"><h2>Version history</h2><span class="band-count">' + h.length + '</span></div>' +
      '<div class="hv-hint">Click a change to see its diff · tick two to compare across them</div>' +
      (listRows.length ? '<div class="vh-list hv-list">' + listRows.join('') + '</div>'
        : '<div class="vh-empty">No changes yet.</div>') +
      '<div class="hv-foot"><span class="vh-me">' + meNote + '</span></div>' +
      '</aside>' +
      '<section class="hv-detail">' + detail + '</section>' +
      '</div>';
    if (window.lucide) lucide.createIcons();
  }
  $('#historyView').addEventListener('click', function (e) {
    var me = e.target.closest('[data-vh-me]');
    if (me) { personalSettingsModal(); return; }
    var tab = e.target.closest('[data-vhtab]');
    if (tab) { vhTab = tab.dataset.vhtab; renderHistoryPage(); return; }
    var ck = e.target.closest('[data-vhck]');
    if (ck) {
      var ci = parseInt(ck.dataset.vhck, 10);
      var at = vhPick.indexOf(ci);
      if (at !== -1) vhPick.splice(at, 1);
      else {
        vhPick.push(ci);
        if (vhPick.length > 2) vhPick.shift();
      }
      renderHistoryPage();
      return;
    }
    var row = e.target.closest('[data-vh]');
    if (row) {
      vhSel = parseInt(row.dataset.vh, 10);
      vhPick = [];
      renderHistoryPage();
    }
  });

  function goToday() {
    var now = new Date();
    var d = RM.dateToDay(state.meta, new Date(Date.UTC(now.getFullYear(), now.getMonth(), now.getDate())));
    if (d != null) scrollLaneTo(d * dayPx());
  }

  function doAddFeature() {
    if (!state.phases.length) {
      commit('add phase', function (s) {
        s.phases.push({ id: RM.uid('p'), name: 'Phase 1', description: '', bucket: false, collapsed: false });
      });
    }
    var phaseId = selectedId && RM.itemById(state, selectedId)
      ? RM.itemById(state, selectedId).phaseId
      : (state.phases.filter(function (p) { return !p.bucket; })[0] || state.phases[0]).id;
    addFeature(phaseId);
  }

  function zoomBy(f) {
    weekPx = Math.max(14, Math.min(80, Math.round(weekPx * f)));
    saveLocal();
    render();
  }

  // menu bar (File / Edit / View)
  // while the new-project wizard is up every menu item is inert (the wizard
  // has its own Back / Continue / Close)
  function menuItems(name) {
    var items = menuItemsFor(name);
    if (!wz) return items;
    return items.map(function (m) {
      if (m.sep) return m;
      var c = {};
      Object.keys(m).forEach(function (k) { c[k] = m[k]; });
      c.disabled = true;
      c.fn = function () {};
      return c;
    });
  }
  function menuItemsFor(name) {
    var isMacDesktop = !!window.HeadwayDesktop && navigator.platform.indexOf('Mac') === 0;
    function openProject() {
      if (window.HeadwayDesktop) HeadwayDesktop.openDialog();
      else $('#filePick').click();
    }
    function toggleAutoSave() {
      autoSave = !autoSave;
      saveLocal(); renderTopbar();
      if (autoSave) scheduleAutoSave();
      toast('Auto save ' + (autoSave ? 'on — writes to the open file' : 'off'));
    }
    if (name === 'macApp') {
      // macOS: file actions live in the app-name menu (desktop.js appends the
      // standard Hide/Quit block after these). nativeIcon must name a macOS
      // *template* image (monochrome, tints with the menu); Folder / Info /
      // MultipleDocuments etc. are full-colour pictograms and look out of
      // place next to the template ones, so those items carry no icon.
      return [
        { icon: 'file-plus-2', nativeIcon: 'Add', label: 'New project', fn: openWizard },
        { icon: 'folder-open', label: 'Open project', fn: openProject },
        { icon: 'file-spreadsheet', label: 'Download template', fn: downloadTemplate },
        { sep: true },
        { icon: 'download', label: 'Save', kbd: '⌘S', fn: function () { $('#btnSave').click(); } },
        { icon: 'save', label: 'Save as…', kbd: '⇧⌘S', fn: function () { window.HeadwayApp.save(true); } },
        { icon: 'timer-reset', label: 'Auto save', checked: autoSave, fn: toggleAutoSave },
        { sep: true },
        { icon: 'share', nativeIcon: 'Share', label: 'Export…', fn: function () { $('#btnExport').click(); } },
        { icon: 'refresh-cw', label: 'Sync with Jira…', fn: function () { HeadwayJira.syncModal(); } },
        { sep: true },
        { icon: 'circle-help', label: 'Help', fn: helpModal }
      ];
    }
    if (name === 'file') {
      return [
        { icon: 'house', label: 'Start page', fn: showStart },
        { sep: true },
        { icon: 'file-plus-2', label: 'New project…', fn: openWizard },
        { icon: 'folder-open', label: 'Open project…', fn: openProject },
        { sep: true },
        { icon: 'download', label: 'Save', kbd: '⌘S', fn: function () { $('#btnSave').click(); } },
        window.HeadwayDesktop
          ? { icon: 'save', label: 'Save as…', fn: function () { window.HeadwayApp.save(true); } }
          : null,
        window.HeadwayDesktop
          ? { icon: 'timer-reset', label: 'Auto save', checked: autoSave, fn: toggleAutoSave }
          : null,
        { sep: true },
        { icon: 'share', label: 'Export…', fn: function () { $('#btnExport').click(); } },
        window.HeadwayJira ? { icon: 'refresh-cw', label: 'Sync with Jira…', fn: function () { HeadwayJira.syncModal(); } } : null,
        { icon: 'file-spreadsheet', label: 'Download template', fn: downloadTemplate },
        { sep: true },
        { icon: 'circle-help', label: 'Shortcuts & help', fn: helpModal }
      ].filter(Boolean);
    }
    if (name === 'edit') {
      return [
        { icon: 'undo-2', label: 'Undo', kbd: '⌘Z', fn: undo, disabled: !undoStack.length },
        { icon: 'redo-2', label: 'Redo', kbd: '⇧⌘Z', fn: redo, disabled: !redoStack.length },
        { sep: true },
        { icon: 'plus', label: esc('Add ' + lvl('feature').toLowerCase()), fn: doAddFeature },
        { icon: 'plus', label: 'Add phase…', fn: function () { phaseModal(null); } },
        { sep: true },
        { icon: 'unlink', label: 'Clear all dependencies…', fn: function () {
          var linked = state.items.filter(function (x) { return x.deps.length || (x.depsText || []).length; }).length;
          confirmBox('Clear all dependencies?', 'Removes every dependency link from ' + linked + ' item(s). Undo works.',
            'Clear', function () {
              commit('clear deps', function (s) {
                s.items.forEach(function (x) { x.deps = []; x.depsText = []; });
              });
              toast('All dependencies cleared');
            }, true);
        } },
        { sep: true },
        { icon: 'settings-2', label: 'Setup', fn: function () {
          view = 'setup';
          saveLocal();
          render();
        } }
      ];
    }
    // view
    var snapItems = [];
    [['feature', lvl('feature', true)], ['story', lvl('story', true)]].forEach(function (k, ki) {
      if (ki) snapItems.push({ sep: true });
      var raw = k[1], labelHtml = esc(raw);
      SNAP_MODES.forEach(function (mode) {
        snapItems.push({ icon: 'magnet', label: labelHtml + ' snap to ' + SNAP_LABELS[mode], checked: snapModeFor(k[0]) === mode, fn: function () {
          setSnapMode(k[0], mode);
          saveLocal(); render(); // the rows too: the Auto timeline buttons depend on the snap
          toast(raw + ' snap: ' + SNAP_LABELS[mode]);
        } });
      });
    });
    return snapItems.concat([
      { sep: true },
      { icon: 'spline', label: 'Dependency arrows', checked: depsMode === 'on', fn: function () {
        depsMode = depsMode === 'on' ? 'none' : 'on';
        saveLocal(); render();
      } },
      { icon: 'flame', label: 'Critical path highlight', checked: showCrit, fn: function () {
        showCrit = !showCrit;
        saveLocal(); render();
        toast('Critical path highlight ' + (showCrit ? 'on' : 'off'));
      } },
      state.meta.capacityEnabled ? { icon: 'gauge', label: 'Capacity rows', checked: showCap, fn: function () {
        showCap = !showCap;
        saveLocal(); render();
        toast('Capacity rows ' + (showCap ? 'shown' : 'hidden'));
      } } : null,
      { sep: true },
      state.meta.workstreamsEnabled ? { icon: 'layers', label: 'Group by workstream', checked: groupWs, fn: function () {
        groupWs = !groupWs;
        saveLocal(); render();
      } } : null,
      { icon: 'layers', label: 'Group by epic', checked: groupEpic, fn: function () {
        groupEpic = !groupEpic;
        saveLocal(); render();
      } },
      { sep: true }
    ].concat(COLOR_MODES.map(function (cm) {
      return { icon: 'palette', label: 'Color by ' + cm[1].toLowerCase(), checked: colorBy === cm[0], fn: function () {
        setColorBy(cm[0]);
        saveLocal(); render();
      } };
    })).concat([
      { sep: true },
      { icon: 'chevrons-up-down', label: 'Expand all features', fn: function () {
        setDetailMode('story');
      } },
      { icon: 'chevrons-down-up', label: 'Collapse all features', fn: function () {
        setDetailMode('feature');
      } },
      { icon: 'arrow-down-narrow-wide', label: 'Auto-order rows by start', checked: autoOrder, fn: function () {
        autoOrder = !autoOrder;
        saveLocal();
        if (autoOrder) commit('auto-order', function (s) { RM.sortItemsByStart(s); });
        else render(); // the Auto timeline buttons' dry runs follow auto-order
        toast('Auto-order ' + (autoOrder ? 'on — rows follow the timeline' : 'off'));
      } },
      state.meta.capacityEnabled ? { icon: 'zap', label: 'Auto timeline: the \u26a1 button on a phase band', disabled: true, fn: function () {} } : null,
      { sep: true },
      // the keys work on Planning; ⌘scroll over the timeline zooms too
      { icon: 'zoom-in', label: 'Zoom in', kbd: '⌘+', title: 'Also ⌘scroll over the timeline', fn: function () { zoomBy(1.2); } },
      { icon: 'zoom-out', label: 'Zoom out', kbd: '⌘−', title: 'Also ⌘scroll over the timeline', fn: function () { zoomBy(1 / 1.2); } },
      { icon: 'crosshair', label: 'Scroll to today', fn: goToday },
      { sep: true },
    ].concat(themeMenuItems())
      // macOS lost the native File menu (its actions moved to the app menu) —
      // keep the start page reachable from View there
      .concat(isMacDesktop ? [{ sep: true }, { icon: 'house', label: 'Start page', fn: showStart }] : []))).filter(Boolean);
  }

  var openMenuName = null;
  function openMenu(b) {
    var name = b.dataset.menu;
    var items = menuItems(name);
    var r = b.getBoundingClientRect();
    var html = '<div class="menu-list">' + items.map(function (m, i) {
      if (m.sep) return '<div class="menu-sep"></div>';
      return '<button data-mi="' + i + '"' + (m.disabled ? ' disabled' : '') + (m.checked ? ' class="on"' : '') +
        (m.title ? ' title="' + esc(m.title) + '"' : '') + '>' +
        '<i data-lucide="' + m.icon + '"></i><span>' + m.label + '</span>' +
        (m.kbd ? '<span class="kbd">' + m.kbd + '</span>' : '') +
        (m.checked ? '<i data-lucide="check" class="mi-check"></i>' : '') +
        '</button>';
    }).join('') + '</div>';
    openPopover(r.left, r.bottom + 6, html, function (host) {
      if (window.lucide) lucide.createIcons();
      host.addEventListener('click', function (ev) {
        var mi = ev.target.closest('[data-mi]');
        if (!mi) return;
        openMenuName = null;
        closePopover();
        items[parseInt(mi.dataset.mi, 10)].fn();
      });
    });
    openMenuName = name;
  }
  $('#menus').addEventListener('click', function (e) {
    var b = e.target.closest('[data-menu]');
    if (!b) return;
    e.stopPropagation();
    if (openMenuName === b.dataset.menu && !popEl.hidden) {
      openMenuName = null;
      closePopover();
      return;
    }
    openMenu(b);
  });
  // once a menu is open, hovering a sibling menu button switches to it;
  // hovering any menu button for 0.4s opens it
  var menuHoverTimer = null;
  $('#menus').addEventListener('mouseover', function (e) {
    var b = e.target.closest('[data-menu]');
    if (!b) return;
    if (openMenuName && !popEl.hidden && popEl.querySelector('.menu-list')) {
      if (openMenuName !== b.dataset.menu) openMenu(b);
      return;
    }
    clearTimeout(menuHoverTimer);
    menuHoverTimer = setTimeout(function () {
      if (b.matches(':hover') && (!openMenuName || popEl.hidden)) openMenu(b);
    }, 400);
  });
  $('#menus').addEventListener('mouseout', function (e) {
    if (e.target.closest('[data-menu]')) clearTimeout(menuHoverTimer);
  });
  document.addEventListener('pointerdown', function (e) {
    if (openMenuName && !popEl.contains(e.target) && !e.target.closest('[data-menu]')) openMenuName = null;
  }, true);

  // ⌘/ctrl-scroll zooms the timeline around the cursor
  // ⌘/ctrl-scroll zooms on Planning AND Budgeting — smooth (scaled by the
  // wheel delta, so trackpad pinches glide instead of jumping in fixed steps)
  board.addEventListener('wheel', function (e) {
    if ((view !== 'planning' && view !== 'budget') || (!e.ctrlKey && !e.metaKey)) return;
    e.preventDefault();
    var r = board.getBoundingClientRect();
    var laneX = board.scrollLeft + (e.clientX - r.left) - leftW();
    var day = laneX / dayPx();
    var f = Math.exp(-e.deltaY * 0.006);
    f = Math.max(0.7, Math.min(1.4, f));
    weekPx = Math.max(14, Math.min(80, weekPx * f));
    saveLocal();
    render();
    board.scrollLeft = Math.max(0, day * dayPx() - (e.clientX - r.left - leftW()));
    requestAnimationFrame(renderArrows);
  }, { passive: false });

  // floating zoom cluster (bottom right of the timeline)
  function syncZoomCtl() {
    var z = $('#zoomCtl');
    if (z) z.hidden = !(view === 'planning' || view === 'budget');
    // the cluster sits above the Resources panel when that is visible
    var rp = $('#resPanel');
    document.documentElement.style.setProperty('--res-total',
      (rp && view === 'planning' ? rp.offsetHeight : 0) + 'px');
  }
  $('#zoomInBtn').addEventListener('click', function () { zoomBy(1.2); });
  $('#zoomOutBtn').addEventListener('click', function () { zoomBy(1 / 1.2); });

  // group = { epic?, workstream? }: the feature joins that group (its values
  // set) and lands right after the group's last row, not the phase's end
  function addFeature(phaseId, group) {
    var newId = RM.uid('i');
    commit('add feature', function (s) {
      var it = RM.normalizeState({
        meta: s.meta, phases: s.phases,
        items: [{ id: newId, num: RM.nextNum(s), phaseId: phaseId, feature: '', size: 'M', headcount: 1,
          capType: s.capTypes[0] || '',
          epic: group && group.epic != null ? group.epic : '',
          workstream: group && group.workstream != null ? group.workstream : '' }]
      }).items[0];
      it.id = newId;
      it.num = RM.nextNum(s);
      var lastIdx = -1;
      var inGroup = function (x) {
        return x.phaseId === phaseId &&
          (!group || group.epic == null || (x.epic || '') === group.epic) &&
          (!group || group.workstream == null || (x.workstream || '') === group.workstream);
      };
      s.items.forEach(function (x, i2) { if (inGroup(x)) lastIdx = i2; });
      if (lastIdx === -1) s.items.forEach(function (x, i2) { if (x.phaseId === phaseId) lastIdx = i2; });
      s.items.splice(lastIdx === -1 ? s.items.length : lastIdx + 1, 0, it);
      s.phases.forEach(function (p) { if (p.id === phaseId) p.collapsed = false; });
      selectedId = newId;
    });
    select(newId, true);
    focusRowTitle(newId);
  }
  // put the caret in a row's title in the left pane (Scoping cell, or the
  // inline editor on a Planning row); the panel's name field is the fallback
  function focusRowTitle(itemId) {
    var nameEl = rowsEl.querySelector('.row[data-id="' + itemId + '"] .sc-name');
    if (!nameEl && view === 'planning' &&
      rowsEl.querySelector('.row.item[data-id="' + itemId + '"] .r-name-txt')) {
      plStartRename(itemId, null);
      return;
    }
    if (nameEl) {
      nameEl.focus();
      if (nameEl.isContentEditable && window.getSelection) {
        var rg = document.createRange();
        rg.selectNodeContents(nameEl);
        rg.collapse(false);
        var sl = window.getSelection();
        sl.removeAllRanges();
        sl.addRange(rg);
      }
      return;
    }
    var nameInput = $('#panel .p-name');
    if (nameInput) nameInput.focus();
  }

  // ------------------------------------------------------------ phase modal
  function phaseModal(phaseId) {
    var phase = null;
    state.phases.forEach(function (p) { if (p.id === phaseId) phase = p; });
    var isNew = !phase;
    var itemCount = phase ? RM.itemsInPhase(state, phase.id).length : 0;
    openModal(
      '<div class="modal" style="width:520px">' +
      '<div class="m-head"><h2>' + (isNew ? 'New phase' : 'Edit phase') + '</h2>' +
      '<button class="p-close" data-m="x"><i data-lucide="x"></i></button></div>' +
      '<div class="m-body">' +
      '<div class="m-sec"><label>Name</label><input id="phName" style="width:100%" value="' + esc(phase ? phase.name : '') + '" placeholder="MVP: Measurement"></div>' +
      '<div class="m-sec"><label>Description</label><div id="phDescEd">' +
      wysHtml('phdesc', phase ? phase.description : '', 'What this phase delivers\u2026') + '</div></div>' +
      '<div class="m-sec"><label class="p-check"><input type="checkbox" id="phBucket"' + (phase && phase.bucket ? ' checked' : '') + '> Backlog bucket (items parked here aren’t auto-scheduled)</label></div>' +
      '<div class="m-sec"><label>Dates</label><div class="p-grid2">' +
      '<div><label class="p-lab">Start</label><input type="text" readonly class="cal-in" data-cal-clear="Auto" id="phStart" style="width:100%" value="' +
        (phase && phase.startDay != null ? esc(RM.fmtISO(RM.dayToDate(state.meta, phase.startDay))) : '') + '"></div>' +
      '<div><label class="p-lab">End</label><input type="text" readonly class="cal-in" data-cal-clear="Auto" id="phEnd" style="width:100%" value="' +
        (phase && phase.endDay != null ? esc(RM.fmtISO(RM.dayToDate(state.meta, Math.max(0, phase.endDay - 1)))) : '') + '"></div>' +
      '</div><div class="m-hint">Blank = automatic (derived from the phase’s items). A set date pins that side; the header span drags/resizes too.</div></div>' +
      (isNew ? '' :
        '<div class="m-sec"><label>Move</label><div class="p-row">' +
        '<button id="phUp">↑ Move up</button><button id="phDown">↓ Move down</button></div></div>') +
      '</div>' +
      '<div class="m-foot">' +
      (isNew || itemCount > 0 || state.phases.length <= 1 ? '' : '<button id="phDelete" class="danger" style="margin-right:auto">Delete phase</button>') +
      (!isNew && itemCount > 0 ? '<span style="margin-right:auto;font-size:11.5px;color:var(--ink-3);align-self:center">' + itemCount + ' item(s) — move them out to delete</span>' : '') +
      (!isNew && itemCount === 0 && state.phases.length <= 1 ? '<span style="margin-right:auto;font-size:11.5px;color:var(--ink-3);align-self:center">the last phase can’t be deleted</span>' : '') +
      (isNew || phase.bucket ? '' : (function () {
        var as = autoPhaseStatus(phase.id);
        return '<button id="phAutoRun" title="' + esc(as.tip) + '"' + (as.disabled ? ' disabled' : '') + '><i data-lucide="zap"></i> Auto timeline</button>';
      })()) +
      '<button data-m="x2">Cancel</button><button id="phSave" class="primary">' + (isNew ? 'Add phase' : 'Save') + '</button>' +
      '</div></div>',
      function (host) {
        $('[data-m=x]', host).onclick = closeModal;
        $('[data-m=x2]', host).onclick = closeModal;
        // Save, and the Auto timeline button, which saves the pending edits first
        function savePhase() {
          var name = $('#phName', host).value.trim() || 'Phase';
          var desc = sanitizeHtml($('#phDescEd .wz-ed', host).innerHTML);
          var bucket = $('#phBucket', host).checked;
          function pinDay(val, isEnd) {
            if (!val) return null;
            var d = RM.dateToDay(state.meta, RM.parseISO(val));
            if (d == null || !isFinite(d)) return null;
            return Math.max(0, d) + (isEnd ? 1 : 0); // endDay is exclusive
          }
          var pStart = pinDay($('#phStart', host).value, false);
          var pEnd = pinDay($('#phEnd', host).value, true);
          closeModal();
          if (isNew) {
            commit('add phase', function (s) {
              s.phases.push({ id: RM.uid('p'), name: name, description: desc, bucket: bucket, collapsed: false, startDay: pStart, endDay: pEnd });
            });
          } else {
            commit('edit phase', function (s) {
              s.phases.forEach(function (p) {
                if (p.id === phaseId) {
                  p.name = name; p.description = desc; p.bucket = bucket;
                  p.startDay = pStart; p.endDay = pEnd;
                }
              });
            });
          }
        }
        $('#phSave', host).onclick = savePhase;
        var autoRun = $('#phAutoRun', host);
        if (autoRun) autoRun.onclick = function () { savePhase(); autoTimelinePhase(phaseId); };
        var del = $('#phDelete', host);
        if (del) del.onclick = function () {
          closeModal();
          commit('delete phase', function (s) {
            s.phases = s.phases.filter(function (p) { return p.id !== phaseId; });
          });
        };
        var up = $('#phUp', host), down = $('#phDown', host);
        if (up) up.onclick = function () { movePhase(phaseId, -1); closeModal(); };
        if (down) down.onclick = function () { movePhase(phaseId, 1); closeModal(); };
      });
  }
  function movePhase(phaseId, dir) {
    commit('move phase', function (s) {
      var i = -1;
      s.phases.forEach(function (p, k) { if (p.id === phaseId) i = k; });
      var j = i + dir;
      if (i < 0 || j < 0 || j >= s.phases.length) return;
      var tmp = s.phases[i]; s.phases[i] = s.phases[j]; s.phases[j] = tmp;
    });
  }

  // ------------------------------------------------------------ team modal
  // ------------------------------------------------------------ validation modal
  $('#btnValidation').addEventListener('click', validationModal);
  $('#btnJira').addEventListener('click', function (e) { if (window.HeadwayJira) HeadwayJira.statusMenu(e.currentTarget); });
  function validationModal() {
    var groups = { error: [], warn: [], info: [] };
    state.items.forEach(function (it) {
      (validation.byItem[it.id] || []).forEach(function (v) {
        groups[v.level === 'error' ? 'error' : v.level].push({ it: it, v: v });
      });
    });
    validation.global.forEach(function (v) {
      groups[v.level].push({ it: null, v: v });
    });
    function section(title, arr, cls) {
      if (!arr.length) return '';
      return '<div class="val-group"><div class="m-label">' + title + ' (' + arr.length + ')</div>' +
        arr.map(function (x) {
          // story findings live in global but name their feature and story:
          // they jump to the feature like any feature finding
          var stRef = !x.it && x.v.storyId ? RM.storyRef(state, x.v.storyId) : null;
          return '<div class="val-item ' + cls + '" data-goto="' + (x.it ? x.it.id : (x.v.itemId || '')) + '" data-week="' + (x.v.week != null ? x.v.week : '') + '">' +
            '<span class="vi-id">' + (x.it ? '#' + x.it.num : stRef ? '#' + stRef.st.num : '⧗') + '</span><span>' + esc(x.v.msg) +
            (x.it ? ' — ' + esc(shorten(x.it.feature, 42)) : '') + '</span></div>';
        }).join('') + '</div>';
    }
    var total = groups.error.length + groups.warn.length + groups.info.length;
    openModal(
      '<div class="modal wide">' +
      '<div class="m-head"><h2>Validation — ' + (total ? total + ' finding(s)' : 'all clear') + '</h2>' +
      '<button class="p-close" data-m="x"><i data-lucide="x"></i></button></div>' +
      '<div class="m-body">' +
      (total === 0 ? '<div class="val-empty"><b>✓</b>No issues. Dependencies, sizes, capacity and timeline all check out.</div>' : '') +
      section('Errors', groups.error, 'err') +
      section('Warnings', groups.warn, 'warn') +
      section('Notes', groups.info, 'info') +
      '</div></div>',
      function (host) {
        $('[data-m=x]', host).onclick = closeModal;
        host.addEventListener('click', function (e) {
          var g = e.target.closest('[data-goto]');
          if (!g) return;
          closeModal();
          if (g.dataset.goto) select(g.dataset.goto, true);
          else if (g.dataset.week) {
            scrollLaneTo(parseInt(g.dataset.week, 10) * weekPx);
          }
        });
      });
  }

  // ------------------------------------------------------------ settings modal
  // ------------------------------------------------------------ setup view
  // ------------------------------------------------------------ budgeting view
  function fmtMoney(n) {
    if (!isFinite(n)) return '—';
    return '$' + Math.round(n).toString().replace(/\B(?=(\d{3})+(?!\d))/g, ',');
  }
  // budgeting hours read to one decimal place
  function fmtH(n) { return Math.round(n * 10) / 10; }

  // budgeting rows render into the shared board: frozen left pane columns
  // (name · role · rate card · workstream · cost · rate · margin · total) +
  // week cells. Name is optional; Role is free text; Rate card is the
  // rate-card role selector (empty = not assigned).
  var BU_COLS = { role: 110, type: 96, cap: 100, ws: 108, cost: 76, rate: 76, margin: 54, total: 80 };
  // header labels + the tooltip that explains each column (tooltips live on
  // the HEADERS — the cells themselves stay quiet)
  var BU_COL_DEFS = {
    role: ['Role', 'Role'],
    type: ['Rate card', 'Rate card'],
    cap: ['Capacity', 'Capacity type this person supplies — stories of that type can be assigned to them and drain it'],
    ws: ['Workstream', 'Workstreams this person works on'],
    cost: ['Cost', 'Cost (hourly)'],
    rate: ['Rate', 'Rate (hourly)'],
    margin: ['Margin', 'Margin'],
    total: ['Total', 'Total']
  };
  var BU_KEYS = ['role', 'type', 'cap', 'ws', 'cost', 'rate', 'margin', 'total'];
  // [short header label, tooltip, default width px, minimum width px]
  var PL_COL_DEFS = {
    size: ['Size', 'Size', 34, 26],
    pri: ['Pri', 'Priority', 26, 22],
    risk: ['Risk', 'Risk', 26, 22],
    dur: ['Wks', 'Duration in weeks', 34, 26],
    asg: ['Ppl', 'Assignees', 40, 30],
    cap: ['Cap', 'Capacity type', 58, 34],
    mult: ['×', 'Capacity multiplier (per-person demand)', 26, 22],
    ws: ['WS', 'Workstream', 96, 44],
    epic: ['Epic', 'Epic', 96, 44],
    start: ['Start', 'Start date', 62, 40],
    deadline: ['Due', 'Deadline', 62, 40]
  };
  var PL_KEYS = ['size', 'pri', 'risk', 'dur', 'asg', 'cap', 'mult', 'ws', 'epic', 'start', 'deadline'];
  // the fixed fields added later stay out of the way until asked for
  var PL_DEFAULT_HIDDEN = ['ws', 'epic', 'start', 'deadline'];
  function plHidden(k) {
    return plColHide[k] != null ? !!plColHide[k] : PL_DEFAULT_HIDDEN.indexOf(k) !== -1;
  }
  function orderedCols(order, allKeys) {
    var out = (order || []).filter(function (k) { return allKeys.indexOf(k) !== -1; });
    allKeys.forEach(function (k) { if (out.indexOf(k) === -1) out.push(k); });
    return out;
  }
  function buColsOrdered() { return orderedCols(buColOrder, BU_KEYS); }
  function plColsOrdered() { return orderedCols(plColOrder, PL_KEYS); }
  function buColsVisible() {
    return buColsOrdered().filter(function (k) { return !buColHide[k]; });
  }
  function plColsVisible() {
    return plColsOrdered().filter(function (k) {
      if (plHidden(k)) return false;
      if (k === 'size') return RM.sizingEnabled(state) || RM.sizingEnabled(state, 'story');
      if (k === 'pri') return RM.priorityEnabled(state) || RM.priorityEnabled(state, 'story');
      if (k === 'risk') return RM.riskEnabled(state) || RM.riskEnabled(state, 'story');
      if (k === 'cap') return !!state.meta.capacityEnabled;
      if (k === 'mult') return !!state.meta.capacityEnabled && state.meta.capMode !== 'points';
      if (k === 'ws') return !!state.meta.workstreamsEnabled;
      return true;
    });
  }
  // columns are user-resizable (drag the header edges); widths flow to every
  // row (and the Resources panel) through --bu-w-* CSS variables
  function buW(k) { return Math.max(44, Math.min(420, parseInt(buColW[k], 10) || BU_COLS[k])); }
  function applyBuColWidths() {
    var rs = document.documentElement.style;
    Object.keys(BU_COLS).forEach(function (k) { rs.setProperty('--bu-w-' + k, buW(k) + 'px'); });
  }
  // the planning chip columns resize the same way, through --pl-w-*
  function plW(k) {
    var d = PL_COL_DEFS[k] || [];
    return Math.max(d[3] || 22, Math.min(240, parseInt(plColW[k], 10) || d[2] || 34));
  }
  function applyPlColWidths() {
    var rs = document.documentElement.style;
    PL_KEYS.forEach(function (k) { rs.setProperty('--pl-w-' + k, plW(k) + 'px'); });
  }
  // the left pane the visible planning columns actually need: fixed row
  // chrome (grip + chevron + dot + warning badge) + every chip column + a
  // floor under the title, so turning columns on never squeezes it away
  var PL_ROW_CHROME = 100, PL_TITLE_MIN = 120;
  function plColsWidth() {
    var sum = 0;
    plColsVisible().forEach(function (k) { sum += plW(k); });
    return PL_ROW_CHROME + sum + PL_TITLE_MIN;
  }
  // the wider of the user's dragged width and what the visible columns need,
  // capped at 70vw. Derived, never persisted: leftWPlan stays the user's floor
  function planLeftW() {
    return Math.max(leftWPlan, Math.min(Math.round(window.innerWidth * 0.7), plColsWidth()));
  }
  // the column-label strip is rendered per view: labels, tooltips, resize
  // handles (budget), and drag-to-reorder all live here
  function renderHlCols() {
    var el = $('#hlCols');
    if (!el) return;
    if (view === 'budget') {
      el.innerHTML = buColsVisible().map(function (k) {
        return '<i class="bu-only" data-bucol="' + k + '" title="' +
          esc(BU_COL_DEFS[k][1]) +
          '" style="width:var(--bu-w-' + k + ')">' + esc(BU_COL_DEFS[k][0]) +
          '<span class="bu-rz" data-burz="' + k + '"></span></i>';
      }).join('');
    } else {
      el.innerHTML = plColsVisible().map(function (k) {
        return '<i class="pl-only' + (k === 'size' ? ' sz-lab' : '') + '" data-plcol="' + k + '" title="' +
          esc(PL_COL_DEFS[k][1]) +
          '" style="width:var(--pl-w-' + k + ')">' + esc(PL_COL_DEFS[k][0]) +
          '<span class="pl-rz" data-plrz="' + k + '"></span></i>';
      }).join('') +
        '<button id="plColsAdd" class="pl-add" title="Columns…" aria-label="Columns"><i data-lucide="plus"></i></button>';
    }
  }
  function columnsMenuItems(kind) {
    var keys = kind === 'bu' ? buColsOrdered() : plColsOrdered();
    var defs = kind === 'bu' ? BU_COL_DEFS : PL_COL_DEFS;
    var hide = kind === 'bu' ? buColHide : plColHide;
    // the menu spells every column out in full, whatever its cramped header says
    var PL_MENU_LBL = {
      size: 'Size', pri: 'Priority', risk: 'Risk', dur: 'Duration', asg: 'Assignees',
      cap: 'Capacity type', mult: 'Multiplier', ws: 'Workstream', epic: 'Epic',
      start: 'Start', deadline: 'Deadline'
    };
    // hidden-ness is a per-kind question: planning columns have defaults
    function isHid(k) { return kind === 'bu' ? !!hide[k] : plHidden(k); }
    var items = keys.map(function (k) {
      var lbl = kind === 'bu' ? (defs[k][0] || k) : (PL_MENU_LBL[k] || defs[k][0] || k);
      return { icon: isHid(k) ? 'eye-off' : 'eye', label: esc(lbl), checked: !isHid(k), fn: function () {
        hide[k] = !isHid(k);
        if (keys.every(function (x) { return isHid(x); })) hide[k] = false; // keep one column
        saveLocal();
        render();
      } };
    });
    items.push({ sep: true });
    items.push({ icon: 'rotate-ccw', label: 'Reset columns', fn: function () {
      if (kind === 'bu') { buColOrder = null; buColHide = {}; buColW = {}; }
      else { plColOrder = null; plColHide = {}; plColW = {}; }
      saveLocal();
      render();
    } });
    return items;
  }
  function startBuColResize(e, k) {
    e.preventDefault();
    e.stopPropagation();
    var w0 = buW(k), x0 = e.clientX;
    function mv(ev) {
      buColW[k] = Math.max(44, Math.min(420, Math.round(w0 + ev.clientX - x0)));
      document.documentElement.style.setProperty('--bu-w-' + k, buW(k) + 'px');
    }
    function up() {
      window.removeEventListener('pointermove', mv);
      window.removeEventListener('pointerup', up);
      saveLocal();
    }
    window.addEventListener('pointermove', mv);
    window.addEventListener('pointerup', up);
  }
  function startPlColResize(e, k) {
    e.preventDefault();
    e.stopPropagation();
    var w0 = plW(k), x0 = e.clientX;
    function mv(ev) {
      plColW[k] = Math.max(PL_COL_DEFS[k][3] || 22, Math.min(240, Math.round(w0 + ev.clientX - x0)));
      document.documentElement.style.setProperty('--pl-w-' + k, plW(k) + 'px');
    }
    function up() {
      window.removeEventListener('pointermove', mv);
      window.removeEventListener('pointerup', up);
      saveLocal();
    }
    window.addEventListener('pointermove', mv);
    window.addEventListener('pointerup', up);
  }
  // drag a header label to reorder its column; the columns menu
  // (show / hide / reset) opens on right-click instead
  function startHdrColDrag(e, cell) {
    e.preventDefault();
    var isBu = cell.dataset.bucol != null;
    var attr = isBu ? 'bucol' : 'plcol';
    var key = cell.dataset[attr];
    var host = $('#hlCols');
    var x0 = e.clientX, moved = false, ind = null, at = null;
    function cells() { return $$('i[data-' + attr + ']', host); }
    function insertIndexAt(cx) {
      var cs = cells(), i2;
      for (i2 = 0; i2 < cs.length; i2++) {
        var r = cs[i2].getBoundingClientRect();
        if (cx < r.left + r.width / 2) return i2;
      }
      return cs.length;
    }
    function mv(ev) {
      if (!moved && Math.abs(ev.clientX - x0) < 4) return;
      moved = true;
      cell.classList.add('colgrab');
      if (!ind) {
        ind = document.createElement('span');
        ind.className = 'col-indicator';
        host.style.position = 'relative';
        host.appendChild(ind);
      }
      at = insertIndexAt(ev.clientX);
      var cs = cells();
      var hr = host.getBoundingClientRect();
      var x = at < cs.length ? cs[at].getBoundingClientRect().left : cs[cs.length - 1].getBoundingClientRect().right;
      ind.style.left = (x - hr.left) + 'px';
    }
    function up(ev) {
      window.removeEventListener('pointermove', mv);
      window.removeEventListener('pointerup', up);
      cell.classList.remove('colgrab');
      if (ind) ind.remove();
      if (!moved) return;
      var vis = (isBu ? buColsVisible() : plColsVisible()).slice();
      var from = vis.indexOf(key);
      if (from === -1 || at == null) return;
      var to = at;
      vis.splice(from, 1);
      if (from < to) to -= 1;
      vis.splice(to, 0, key);
      var full = vis.concat((isBu ? buColsOrdered() : plColsOrdered()).filter(function (k2) {
        return vis.indexOf(k2) === -1;
      }));
      if (isBu) buColOrder = full; else plColOrder = full;
      saveLocal();
      render();
    }
    window.addEventListener('pointermove', mv);
    window.addEventListener('pointerup', up);
  }
  $('#hlCols').addEventListener('pointerdown', function (e) {
    if (e.button !== 0) return;
    var rz = e.target.closest('.bu-rz');
    if (rz) { startBuColResize(e, rz.dataset.burz); return; }
    var prz = e.target.closest('.pl-rz');
    if (prz) { startPlColResize(e, prz.dataset.plrz); return; }
    if (e.target.closest('#plColsAdd')) return; // the + is a click, not a drag
    var cell = e.target.closest('i[data-bucol],i[data-plcol]');
    if (cell) startHdrColDrag(e, cell);
  });
  // the + at the end of the planning strip opens the same show/hide menu
  $('#hlCols').addEventListener('click', function (e) {
    var add = e.target.closest('#plColsAdd');
    if (!add) return;
    e.preventDefault();
    e.stopPropagation();
    openDropdown(add, columnsMenuItems('pl'), { minW: 190 });
  });
  // right-click a column label (or the header corner, below) = columns menu
  $('#hlCols').addEventListener('contextmenu', function (e) {
    var cell = e.target.closest('i[data-bucol],i[data-plcol]');
    if (!cell) return;
    e.preventDefault();
    e.stopPropagation(); // the corner + theme fallbacks must not double up
    openContextMenu(e.clientX, e.clientY, columnsMenuItems(cell.dataset.bucol != null ? 'bu' : 'pl'));
  });
  document.addEventListener('contextmenu', function (e) {
    var corner = e.target.closest('.hdr-left.corner');
    if (!corner || (view !== 'budget' && view !== 'planning')) return;
    e.preventDefault();
    e.stopPropagation();
    openContextMenu(e.clientX, e.clientY, columnsMenuItems(view === 'budget' ? 'bu' : 'pl'));
  });
  // shared left-hand member columns — the Resources panel under the timeline
  // renders these SAME columns in the user's order (budgeting adds the money
  // columns, also orderable/hideable). moneyMap: key -> html for cost/rate/
  // margin/total; absent (Resources panel) skips those keys entirely.
  function memberColsHtml(m, moneyMap) {
    var mws = RM.memberWorkstreams(m);
    var wsHex = RM.colorForWs(state, mws[0] || '');
    var cols = {
      role: '<input class="bu-in bu-col" style="width:var(--bu-w-role)" data-bud="role" value="' + esc(m.role || '') +
        '" placeholder="Role">',
      type: '<span class="r-ws sc-chip bu-col bu-chip' + (m.type ? '' : ' empty') +
        '" style="width:var(--bu-w-type)" tabindex="0" role="button" data-bact="type">' +
        (m.type ? esc(shorten(m.type, 13)) : '—') + '</span>',
      cap: '<span class="r-ws sc-chip bu-col bu-chip' + (m.capType ? '' : ' empty') +
        '" style="width:var(--bu-w-cap)" tabindex="0" role="button" data-bact="cap" title="Capacity type (from the role)">' +
        (m.capType ? esc(shorten(m.capType, 13)) : '—') + '</span>',
      ws: '<span class="r-ws sc-chip bu-col bu-chip' + (mws.length ? '' : ' empty') +
        '" style="width:var(--bu-w-ws)" tabindex="0" role="button" data-bact="ws">' +
        (mws.length ? '<span class="dd-dot" style="background:#' + wsHex + '"></span>' + esc(shorten(mws[0], 12)) +
          (mws.length > 1 ? '<small class="bu-wsmore">+' + (mws.length - 1) + '</small>' : '') : '—') + '</span>'
    };
    var out = '<span class="bu-nm">' + avatarHtml(m, 'sm') +
      '<input class="bu-in bu-nm-in" data-bud="name" value="' + esc(m.name || '') +
      '" placeholder="Name"></span>';
    buColsVisible().forEach(function (k) {
      if (cols[k]) out += cols[k];
      else if (moneyMap && moneyMap[k] != null) out += moneyMap[k];
    });
    return out;
  }
  function renderBudgetRows() {
    var meta = state.meta;
    var html = [];
    var sumTotal = 0;
    var fullWeek = RM.weekHoursOf(meta);
    state.team.forEach(function (m) {
      var mws = RM.memberWorkstreams(m);
      var wsHex = RM.colorForWs(state, mws[0] || '');
      var cr = parseInt(wsHex.slice(0, 2), 16), cg = parseInt(wsHex.slice(2, 4), 16), cb = parseInt(wsHex.slice(4, 6), 16);
      var hours = RM.roleTotalHours(state, m);
      var effRate = RM.memberRate(state, m);
      var effCost = RM.memberCost(state, m);
      var total = hours * effRate; // billing total: actual hours × effective rate
      sumTotal += total;
      var margin = RM.roleMargin(state, m);
      var rc = RM.rateCardFor(state, m.type);
      var cells = [];
      for (var w = 0; w < meta.numWeeks; w++) {
        var wh = RM.roleWeekHours(state, m, w);
        var a = Math.max(0, Math.min(1, wh.actual / fullWeek));
        // clipped weeks (holidays) show the ACTUAL hours below in small text —
        // that's what totals and cost are built from
        var clipped = wh.actual < wh.planned;
        cells.push('<div class="bu-cell' + (clipped ? ' clipped' : '') + '" tabindex="0" data-w="' + w + '" data-iso="' + wh.iso +
          '" style="left:' + (w * weekPx) + 'px;width:' + weekPx +
          'px;background:rgba(' + cr + ',' + cg + ',' + cb + ',' + (a * 0.30).toFixed(3) + ')">' +
          (weekPx >= 20
            ? '<span>' + (fmtH(wh.planned) || '') + '</span>' + (clipped ? '<span class="bu-sub">(' + fmtH(wh.actual) + ')</span>' : '')
            : '') + '</div>');
      }
      var moneyMap = {
        cost: '<input class="bu-in bu-col' + (!m.cost && effCost ? ' inherited' : '') +
          '" style="width:var(--bu-w-cost)" type="number" min="0" data-bud="cost" value="' + (m.cost || '') +
          '" placeholder="' + (rc && rc.cost ? rc.cost : 0) + '">',
        rate: '<input class="bu-in bu-col' + (!m.rate && effRate ? ' inherited' : '') +
          '" style="width:var(--bu-w-rate)" type="number" min="0" data-bud="rate" value="' + (m.rate || '') +
          '" placeholder="' + (rc && rc.rate ? rc.rate : 0) + '">',
        margin: '<span class="bu-col bu-ro" style="width:var(--bu-w-margin)">' + (margin == null ? '—' : Math.round(margin) + '%') + '</span>',
        total: '<span class="bu-col bu-ro" style="width:var(--bu-w-total)">' + fmtMoney(total) + '</span>'
      };
      html.push('<div class="row brole" data-mid="' + m.id + '">' +
        '<div class="row-left">' +
        '<span class="r-grip bu-grip"><i data-lucide="grip-vertical"></i></span>' +
        '<span class="r-dot" style="background:#' + wsHex + '"></span>' +
        memberColsHtml(m, moneyMap) +
        '</div><div class="row-lane">' + cells.join('') + '</div></div>');
    });
    // click-to-add role row (same flow as the Resources panel)
    html.push('<div class="row addrow" data-kind="baddrole" title="Add a person to the roster">' +
      '<div class="row-left"><span class="addrow-lab"><i data-lucide="plus"></i> Add role</span></div>' +
      '<div class="row-lane"></div></div>');

    // ---- fixed & recurring costs: their own band + rows with timeline markers
    var costs = state.costs || [];
    var costsTotal = RM.costsTotal(state);
    html.push('<div class="row band bcosts-band"><div class="row-left">' +
      '<span class="band-name">COSTS</span><span class="band-count">' + costs.length + '</span></div>' +
      '<div class="row-lane"></div></div>');
    costs.forEach(function (c) {
      var occ = RM.costOccurrences(state, c);
      var marks = occ.map(function (o) {
        return '<span class="bu-costmark" title="' + esc(c.name) + '" style="left:' + (o.day * dayPx()) + 'px"></span>';
      }).join('');
      var kindLabel = c.kind === 'fixed' ? 'One-time' : c.kind === 'weekly' ? 'Weekly' : 'Monthly';
      var cTotal = RM.costTotal(state, c);
      sumTotal = sumTotal; // billing total stays roster-only; costs are spend
      // cost rows fill the same visible column slots as the people above:
      // kind chip in the Role slot, the date range across Rate card +
      // Workstream (merged only when those two sit side by side)
      var vis = buColsVisible();
      var datesTxt = esc(RM.fmtShort(RM.dayToDate(meta, c.startDay)) +
        (c.kind !== 'fixed' ? ' →' + (c.endDay != null ? ' ' + RM.fmtShort(RM.dayToDate(meta, c.endDay)) : ' end') : ''));
      var datesAt = vis.indexOf('type') !== -1 ? 'type' : (vis.indexOf('ws') !== -1 ? 'ws' : null);
      var mergeWs = datesAt === 'type' && vis[vis.indexOf('type') + 1] === 'ws';
      var costCols = '';
      vis.forEach(function (k) {
        if (k === 'role') {
          costCols += '<span class="r-ws sc-chip bu-col bu-chip" style="width:var(--bu-w-role)" tabindex="0" role="button" data-cact="kind">' + kindLabel + '</span>';
        } else if (k === datesAt) {
          costCols += '<span class="r-ws sc-chip bu-col bu-chip" style="width:' +
            (mergeWs ? 'calc(var(--bu-w-type) + var(--bu-w-ws))' : 'var(--bu-w-' + k + ')') +
            '" tabindex="0" role="button" data-cact="dates">' + datesTxt + '</span>';
        } else if (k === 'ws' && mergeWs) {
          // swallowed by the merged dates chip
        } else if (k === 'cost') {
          costCols += '<input class="bu-in bu-col" style="width:var(--bu-w-cost)" type="number" min="0" data-cf="amount" value="' + c.amount + '">';
        } else if (k === 'total') {
          costCols += '<span class="bu-col bu-ro" style="width:var(--bu-w-total)">' + fmtMoney(cTotal) + '</span>';
        } else {
          costCols += '<span class="bu-col" style="width:var(--bu-w-' + k + ')"></span>';
        }
      });
      html.push('<div class="row brole bcost" data-cost="' + c.id + '">' +
        '<div class="row-left">' +
        '<span class="r-grip bu-grip"><i data-lucide="grip-vertical"></i></span>' +
        '<span class="r-dot nodot"></span>' +
        '<span class="bu-nm"><input class="bu-in bu-nm-in" data-cf="name" value="' + esc(c.name) + '"></span>' +
        costCols +
        '</div><div class="row-lane">' + marks + '</div></div>');
    });
    html.push('<div class="row addrow" data-kind="baddcost" title="Add a fixed or recurring cost">' +
      '<div class="row-left"><span class="addrow-lab"><i data-lucide="plus"></i> Add cost</span></div>' +
      '<div class="row-lane"></div></div>');

    if (state.team.length || costs.length) {
      html.push('<div class="row brole btotal"><div class="row-left">' +
        '<span class="r-dot" style="background:transparent"></span>' +
        '<span class="bu-nm">Billing ' + fmtMoney(sumTotal) + (costsTotal ? ' · Costs ' + fmtMoney(costsTotal) : '') + '</span>' +
        '</div><div class="row-lane"></div></div>');
    }
    rowsEl.innerHTML = html.join('');
    if (window.lucide) lucide.createIcons();
  }

  // typing into a focused money cell replaces its contents; focused week
  // cells grow a spreadsheet fill handle
  rowsEl.addEventListener('focusin', function (e) {
    if (view !== 'budget') return;
    if (e.target.classList && e.target.classList.contains('bu-in')) e.target.select();
    var td = e.target.closest && e.target.closest('.bu-cell');
    if (td && !td.querySelector('.bu-fill')) addBuFillHandle(td);
  });
  rowsEl.addEventListener('focusout', function (e) {
    var h = e.target.querySelector && e.target.querySelector('.bu-fill');
    if (h && !(drag && drag.kind === 'bfill')) h.remove();
  });

  // clicking a cell while another is mid-edit: the blur-commit re-render
  // destroys the pressed element before the click lands, so remember what
  // was pressed and put focus there on release
  var buPendingFocus = null;
  rowsEl.addEventListener('pointerdown', function (e) {
    if (view !== 'budget') { buPendingFocus = null; return; }
    var t = e.target.closest && e.target.closest('input[data-bud], .bu-cell');
    buPendingFocus = t ? buFocusSelector(t) : null;
  });
  window.addEventListener('pointerup', function () {
    if (!buPendingFocus) return;
    var sel = buPendingFocus;
    buPendingFocus = null;
    if (view !== 'budget') return;
    var el = rowsEl.querySelector(sel);
    if (!el || el === document.activeElement) return;
    if (el.classList.contains('bu-cell')) openBuCellEditor(el);
    else { el.focus(); if (el.select) el.select(); }
  });

  // spreadsheet fill: drag the focused cell's corner handle across weeks to
  // spread its hour value (same commit semantics as the resources panel)
  function addBuFillHandle(td) {
    var h = document.createElement('span');
    h.className = 'bu-fill';
    h.title = 'Drag to fill weeks with this value';
    h.addEventListener('pointerdown', function (e) {
      if (e.button !== 0 || drag) return;
      e.preventDefault();
      e.stopPropagation();
      var rowEl2 = td.closest('[data-mid]');
      var m = null;
      state.team.forEach(function (x) { if (x.id === rowEl2.dataset.mid) m = x; });
      if (!m) return;
      var w = parseInt(td.dataset.w, 10);
      var r0 = $$('#rows .row.brole[data-mid]').indexOf(rowEl2);
      drag = {
        kind: 'bfill', mid: m.id, w0: w, w1: w, r0: r0, r1: r0,
        val: RM.memberHoursForWeek(state.meta, m, w),
        x0: e.clientX, y0: e.clientY, moved: false
      };
    });
    td.appendChild(h);
  }
  function bfillMove(e) {
    var rows = $$('#rows .row.brole[data-mid]');
    var srcRow = rows[drag.r0];
    if (!srcRow) return;
    var r = srcRow.querySelector('.row-lane').getBoundingClientRect();
    drag.w1 = Math.max(0, Math.min(state.meta.numWeeks - 1, Math.floor((e.clientX - r.left) / weekPx)));
    // vertical: the handle also fills downward/upward across roles
    drag.r1 = drag.r0;
    rows.forEach(function (rw, i) {
      var rr = rw.getBoundingClientRect();
      if (e.clientY >= rr.top && e.clientY < rr.bottom) drag.r1 = i;
    });
    var lo = Math.min(drag.w0, drag.w1), hi = Math.max(drag.w0, drag.w1);
    var rlo = Math.min(drag.r0, drag.r1), rhi = Math.max(drag.r0, drag.r1);
    rows.forEach(function (rw, i) {
      var inRows = i >= rlo && i <= rhi;
      $$('.bu-cell', rw).forEach(function (c) {
        var w = parseInt(c.dataset.w, 10);
        c.classList.toggle('bu-fillsel', inRows && w >= lo && w <= hi);
      });
    });
    dragTip.hidden = false;
    dragTip.style.left = (e.clientX + 14) + 'px';
    dragTip.style.top = (e.clientY - 34) + 'px';
    dragTip.textContent = fmtH(drag.val) + 'h × ' + (hi - lo + 1) + 'w' +
      (rhi > rlo ? ' × ' + (rhi - rlo + 1) + ' roles' : '');
  }
  function bfillEnd(d) {
    var rows = $$('#rows .row.brole[data-mid]');
    var rlo = Math.min(d.r0, d.r1), rhi = Math.max(d.r0, d.r1);
    var mids = rows.slice(rlo, rhi + 1).map(function (rw) { return rw.dataset.mid; });
    fillWeekHours(mids.length ? mids : [d.mid], d.w0, d.w1, d.val);
  }

  // drag a row's grip to reorder people or costs (budget view); people
  // reorder within the roster, costs within the costs band
  rowsEl.addEventListener('pointerdown', function (e) {
    if (view !== 'budget' || e.button !== 0 || drag) return;
    var grip = e.target.closest('.bu-grip');
    if (!grip) return;
    var pr = grip.closest('.row[data-mid]');
    var cr = grip.closest('.row[data-cost]');
    if (!pr && !cr) return;
    drag = {
      kind: 'brow',
      id: cr ? cr.dataset.cost : pr.dataset.mid,
      isCost: !!cr,
      x0: e.clientX, y0: e.clientY, moved: false
    };
    e.preventDefault();
  });
  function browRows(d) {
    return $$(d.isCost ? '#rows .row.bcost[data-cost]' : '#rows .row.brole[data-mid]');
  }
  function browRowId(rw, d) { return d.isCost ? rw.dataset.cost : rw.dataset.mid; }
  function browMove(e) {
    var rows = browRows(drag);
    if (!drag.indicator) {
      drag.indicator = document.createElement('div');
      drag.indicator.className = 'rrow-indicator';
      rowsEl.appendChild(drag.indicator);
      rows.forEach(function (rw) {
        if (browRowId(rw, drag) === drag.id) rw.classList.add('drag-row-ghost');
      });
    }
    drag.before = null;
    var g = rowsEl.getBoundingClientRect();
    var y = null;
    for (var i = 0; i < rows.length; i++) {
      var r = rows[i].getBoundingClientRect();
      if (e.clientY < r.top + r.height / 2) { drag.before = browRowId(rows[i], drag); y = r.top - g.top; break; }
    }
    if (y == null && rows.length) y = rows[rows.length - 1].getBoundingClientRect().bottom - g.top;
    drag.indicator.style.top = (y || 0) + 'px';
  }
  function browEnd(d) {
    browRows(d).forEach(function (rw) { rw.classList.remove('drag-row-ghost'); });
    if (d.before === d.id) { render(); return; }
    commit(d.isCost ? 'reorder cost' : 'reorder person', function (s) {
      var list = d.isCost ? (s.costs || []) : s.team;
      var moved = null;
      list.forEach(function (x) { if (x.id === d.id) moved = x; });
      if (!moved) return;
      list = list.filter(function (x) { return x.id !== d.id; });
      var at = d.before ? list.map(function (x) { return x.id; }).indexOf(d.before) : list.length;
      if (at < 0) at = list.length;
      list.splice(at, 0, moved);
      if (d.isCost) s.costs = list; else s.team = list;
    });
  }

  // rate/cost commits (budget rows). The commit re-renders, which would eat
  // the focus mid-Tab — so remember where focus LANDED and put it back.
  function buFocusSelector(el) {
    var row = el && el.closest && el.closest('[data-mid]');
    if (!row) return null;
    var part = el.dataset.bud ? '[data-bud="' + el.dataset.bud + '"]'
      : el.dataset.bact ? '[data-bact="' + el.dataset.bact + '"]'
      : el.dataset.iso ? '.bu-cell[data-iso="' + el.dataset.iso + '"]' : null;
    return part ? '#rows [data-mid="' + row.dataset.mid + '"] ' + part : null;
  }
  rowsEl.addEventListener('change', function (e) {
    if (view !== 'budget') return;
    var f = e.target.dataset.bud;
    var rowEl2 = e.target.closest('[data-mid]');
    if (!f || !rowEl2) return;
    var roleId = rowEl2.dataset.mid;
    if (f === 'name' || f === 'role') {
      var tv = e.target.value.trim();
      commit(f === 'name' ? 'rename person' : 'person role', function (s) {
        s.team.forEach(function (m) { if (m.id === roleId) m[f] = tv; });
      });
      return;
    }
    var v = Math.max(0, parseFloat(e.target.value) || 0);
    commit('role ' + f, function (s) {
      s.team.forEach(function (m) { if (m.id === roleId) m[f] = v; });
    });
  });

  // keyboard: Enter/Space activates chips and week cells; Tab through the
  // budget controls survives the commit re-render (we steer it ourselves)
  rowsEl.addEventListener('keydown', function (e) {
    if (view !== 'budget') return;
    if (e.key === 'Tab') {
      var inp = e.target.closest && e.target.closest('input[data-bud]');
      if (!inp) return; // chips/cells tab natively; cell editors hop themselves
      e.preventDefault();
      var focusables = $$('#rows input[data-bud], #rows [data-bact], #rows .bu-cell');
      var i = focusables.indexOf(inp);
      var next = focusables[i + (e.shiftKey ? -1 : 1)];
      var nextSel = next && buFocusSelector(next);
      inp.blur(); // fires change → commit → re-render when the value changed
      setTimeout(function () {
        var nf = (nextSel && rowsEl.querySelector(nextSel)) || (next && next.isConnected ? next : null);
        if (nf) { nf.focus(); if (nf.select) nf.select(); }
      }, 0);
      return;
    }
    if (e.key !== 'Enter' && e.key !== ' ') return;
    var chip = e.target.closest && e.target.closest('[data-bact]');
    if (chip) { e.preventDefault(); chip.click(); return; }
    var td = e.target.closest && e.target.closest('.bu-cell');
    if (td && !td.querySelector('input')) { e.preventDefault(); openBuCellEditor(td); }
  });

  // member chips (Budgeting rows AND the Resources panel): rate-card role and
  // workstream pick from the shared dropdown
  function openMemberTypeDropdown(chip, roleId) {
    var m = memberById(roleId);
    if (!m) return;
    var items = [{ label: '<i>— none —</i>', checked: !m.type, fn: function () {
      commit('rate card role', function (s) {
        s.team.forEach(function (x) { if (x.id === roleId) x.type = ''; });
      });
    } }];
    var rct = state.roleCapTypes || {};
    state.teamTypes.forEach(function (t) {
      items.push({ label: esc(t) + (rct[t] ? ' <small>' + esc(rct[t]) + '</small>' : ''), checked: m.type === t, fn: function () {
        commit('rate card role', function (s) {
          s.team.forEach(function (x) { if (x.id === roleId) x.type = t; });
          RM.syncMemberCapTypes(s); // the role brings its capacity type
        });
      } });
    });
    openDropdown(chip, items);
  }
  // a person's capacity type comes from their role: with no role, pick one;
  // with a role, the menu points to where the role's type is set
  function openMemberCapDropdown(chip, roleId) {
    var m = memberById(roleId);
    if (!m) return;
    if (!m.type) { openMemberTypeDropdown(chip, roleId); return; }
    openDropdown(chip, [{ icon: 'gauge',
      label: esc(m.capType ? m.type + ' supplies ' + m.capType : m.type + ' supplies no capacity type') + ' \u2014 change in Setup \u203a Scheduling\u2026',
      fn: function () { setupTab = 'scheduling'; view = 'setup'; saveLocal(); render(); } }]);
  }
  function openMemberWsDropdown(chip, roleId) {
    var m = memberById(roleId);
    if (!m) return;
    // people can sit on several workstreams — each click toggles membership
    var curWs = RM.memberWorkstreams(m);
    var wsItems = [{ label: '<i>— none —</i>', checked: !curWs.length, fn: function () {
      commit('role workstream', function (s) {
        s.team.forEach(function (x) { if (x.id === roleId) RM.setMemberWorkstreams(x, []); });
      });
    } }];
    allWorkstreams().forEach(function (wv) {
      var onWs = curWs.indexOf(wv) !== -1;
      wsItems.push({
        label: esc(wv), dot: '#' + RM.colorForWs(state, wv), checked: onWs,
        fn: function () {
          commit('role workstream', function (s) {
            s.team.forEach(function (x) {
              if (x.id !== roleId) return;
              var list = RM.memberWorkstreams(x).slice();
              var at = list.indexOf(wv);
              if (at === -1) list.push(wv);
              else list.splice(at, 1);
              RM.setMemberWorkstreams(x, list);
            });
          });
        }
      });
    });
    openDropdown(chip, wsItems);
  }
  rowsEl.addEventListener('click', function (e) {
    if (view !== 'budget') return;
    var chip = e.target.closest('[data-bact]');
    var rowEl2 = e.target.closest('[data-mid]');
    if (!chip || !rowEl2 || rowEl2.classList.contains('btotal')) return;
    if (chip.dataset.bact === 'type') openMemberTypeDropdown(chip, rowEl2.dataset.mid);
    else if (chip.dataset.bact === 'cap') openMemberCapDropdown(chip, rowEl2.dataset.mid);
    else openMemberWsDropdown(chip, rowEl2.dataset.mid);
  });

  // add-role / add-cost rows + cost field edits (budget view)
  function addBudgetRole(rowEl) {
    var lab = rowEl.querySelector('.addrow-lab');
    if (!lab || lab.querySelector('input')) return;
    var inp = document.createElement('input');
    inp.placeholder = 'Name…';
    inp.className = 'st-add-input';
    lab.innerHTML = '';
    lab.appendChild(inp);
    inp.focus();
    var done = false;
    function finish(saveIt) {
      if (done) return; done = true;
      var v = inp.value.trim();
      if (saveIt && v) {
        commit('add person', function (s) {
          s.team.push({ id: RM.uid('t'), name: v, role: '', type: '', weekHours: {} });
        });
      } else render();
    }
    inp.addEventListener('blur', function () { finish(true); });
    inp.addEventListener('keydown', function (ev) {
      ev.stopPropagation();
      if (ev.key === 'Enter') finish(true);
      if (ev.key === 'Escape') finish(false);
    });
  }
  function costById(id) {
    return (state.costs || []).filter(function (c) { return c.id === id; })[0] || null;
  }
  function costDatesPopover(anchor, costId) {
    var c = costById(costId);
    if (!c) return;
    var r = anchor.getBoundingClientRect();
    var meta = state.meta;
    openPopover(r.left, r.bottom + 4,
      '<div class="rolep">' +
      '<label class="p-lab">First occurrence</label>' +
      '<input type="text" readonly class="cal-in" id="csStart" value="' + RM.fmtISO(RM.dayToDate(meta, c.startDay)) + '">' +
      (c.kind !== 'fixed'
        ? '<label class="p-lab" style="margin-top:8px">Last occurrence (empty = timeline end)</label>' +
          '<input type="text" readonly class="cal-in" data-cal-clear="Timeline end" id="csEnd" value="' + (c.endDay != null ? RM.fmtISO(RM.dayToDate(meta, c.endDay)) : '') + '">'
        : '') +
      '<div class="p-row" style="margin-top:9px"><button id="csApply" class="primary fixed">Apply</button></div>' +
      '</div>',
      function (host) {
        function apply() {
          var sd = RM.dateToDay(state.meta, RM.parseISO($('#csStart', host).value));
          var edEl = $('#csEnd', host);
          var ed = edEl && edEl.value ? RM.dateToDay(state.meta, RM.parseISO(edEl.value)) : null;
          closePopover();
          commit('cost dates', function (s) {
            var c2 = (s.costs || []).filter(function (x) { return x.id === costId; })[0];
            if (!c2) return;
            if (sd != null) c2.startDay = Math.max(0, sd);
            c2.endDay = ed != null && ed >= c2.startDay ? ed : null;
          });
        }
        $('#csApply', host).onclick = apply;
        host.addEventListener('keydown', function (ke) { if (ke.key === 'Enter') apply(); });
      });
  }
  rowsEl.addEventListener('click', function (e) {
    if (view !== 'budget') return;
    var add = e.target.closest('.row.addrow');
    if (add) {
      if (add.dataset.kind === 'baddrole') { addBudgetRole(add); return; }
      if (add.dataset.kind === 'baddcost') {
        commit('add cost', function (s) {
          s.costs.push({ id: RM.uid('cost'), name: 'New cost', amount: 0, kind: 'fixed', startDay: 0, endDay: null });
        });
        requestAnimationFrame(function () {
          var last = $$('#rows .row.bcost').pop();
          var inp = last && last.querySelector('[data-cf="name"]');
          if (inp) { inp.focus(); inp.select(); }
        });
        return;
      }
    }
    var cact = e.target.closest('[data-cact]');
    var costRow = e.target.closest('[data-cost]');
    if (cact && costRow) {
      var costId = costRow.dataset.cost;
      if (cact.dataset.cact === 'kind') {
        openDropdown(cact, [['fixed', 'One-time'], ['weekly', 'Weekly'], ['monthly', 'Monthly']].map(function (k) {
          return { label: k[1], checked: (costById(costId) || {}).kind === k[0], fn: function () {
            commit('cost kind', function (s) {
              var c2 = (s.costs || []).filter(function (x) { return x.id === costId; })[0];
              if (c2) c2.kind = k[0];
            });
          } };
        }));
      } else if (cact.dataset.cact === 'dates') {
        costDatesPopover(cact, costId);
      }
    }
  });
  rowsEl.addEventListener('change', function (e) {
    if (view !== 'budget') return;
    var cf = e.target.dataset.cf;
    var costRow = e.target.closest('[data-cost]');
    if (!cf || !costRow) return;
    var costId = costRow.dataset.cost;
    var val = cf === 'amount' ? Math.max(0, parseFloat(e.target.value) || 0) : e.target.value.trim();
    commit('cost ' + cf, function (s) {
      var c2 = (s.costs || []).filter(function (x) { return x.id === costId; })[0];
      if (!c2) return;
      if (cf === 'amount') c2.amount = val;
      else if (val) c2.name = val;
    });
  });

  // right-click: budget people and cost rows get their own menus
  rowsEl.addEventListener('contextmenu', function (e) {
    if (view !== 'budget') return;
    if (e.target.closest('.row-lane')) return; // lane = holiday quick-edit menu
    var costRow = e.target.closest('[data-cost]');
    var roleRow = e.target.closest('[data-mid]');
    if (!costRow && !roleRow) return;
    e.preventDefault();
    e.stopPropagation();
    if (costRow) {
      var cid = costRow.dataset.cost;
      openContextMenu(e.clientX, e.clientY, [
        { icon: 'trash-2', label: 'Remove cost', fn: function () {
          commit('remove cost', function (s) {
            s.costs = (s.costs || []).filter(function (x) { return x.id !== cid; });
          });
        } }
      ]);
      return;
    }
    if (roleRow.classList.contains('btotal')) return;
    var mid2 = roleRow.dataset.mid;
    var mm2 = null;
    state.team.forEach(function (x) { if (x.id === mid2) mm2 = x; });
    if (!mm2) return;
    openContextMenu(e.clientX, e.clientY, [
      { icon: 'pencil', label: 'Rename', fn: function () {
        var nm = roleRow.querySelector('.bu-nm');
        if (nm) startInlineEdit(nm, function (v) {
          if (v.trim()) commit('rename person', function (s) {
            s.team.forEach(function (x) { if (x.id === mid2) x.name = v.trim(); });
          });
        });
      } },
      { icon: 'tags', label: 'Rate card…', fn: function () {
        var chipT = roleRow.querySelector('[data-bact="type"]');
        if (chipT) chipT.click();
      } },
      { icon: 'layers', label: 'Workstream…', fn: function () {
        var chipW = roleRow.querySelector('[data-bact="ws"]');
        if (chipW) chipW.click();
      } },
      { sep: true },
      { icon: 'trash-2', label: 'Remove person…', fn: function () { deleteRoleConfirm(mid2); } }
    ]);
  });

  // week-hour cells edit in place (budget rows); Tab/Shift+Tab commits and
  // hops to the neighbouring week like the resources spreadsheet
  function openBuCellEditor(td) {
    if (!td || td.querySelector('input')) return;
    var rowEl2 = td.closest('[data-mid]');
    if (!rowEl2) return;
    var roleId = rowEl2.dataset.mid, iso = td.dataset.iso;
    var mm = null;
    state.team.forEach(function (x) { if (x.id === roleId) mm = x; });
    if (!mm) return;
    var cur = mm.weekHours[iso] != null ? mm.weekHours[iso] : RM.weekHoursOf(state.meta);
    var inp = document.createElement('input');
    inp.type = 'number';
    inp.min = '0';
    inp.className = 'bu-in';
    inp.value = cur;
    td.textContent = '';
    td.appendChild(inp);
    inp.focus(); inp.select();
    var done = false;
    function fin(saveIt, hopDir) {
      if (done) return; done = true;
      var h = Math.max(0, parseFloat(inp.value));
      if (saveIt && isFinite(h)) {
        commit('role hours', function (s) {
          s.team.forEach(function (m2) {
            if (m2.id !== roleId) return;
            if (h === RM.weekHoursOf(s.meta)) delete m2.weekHours[iso]; else m2.weekHours[iso] = h;
          });
        });
      } else render();
      if (hopDir) {
        var cells = $$('#rows [data-mid="' + roleId + '"] .bu-cell');
        var idx = -1;
        cells.forEach(function (c, i) { if (c.dataset.iso === iso) idx = i; });
        var next = cells[idx + hopDir];
        if (next) { next.focus(); openBuCellEditor(next); }
      } else if (saveIt) {
        var back = rowsEl.querySelector('[data-mid="' + roleId + '"] .bu-cell[data-iso="' + iso + '"]');
        if (back) back.focus();
      }
    }
    inp.addEventListener('blur', function () { fin(true); });
    inp.addEventListener('keydown', function (ev) {
      ev.stopPropagation();
      if (ev.key === 'Tab') { ev.preventDefault(); fin(true, ev.shiftKey ? -1 : 1); }
      if (ev.key === 'Enter') fin(true);
      if (ev.key === 'Escape') fin(false);
    });
  }
  rowsEl.addEventListener('click', function (e) {
    if (view !== 'budget' || Date.now() - dragEndAt < 120) return;
    openBuCellEditor(e.target.closest('.bu-cell'));
  });

  // every Setup section's cards, keyed by section; the wizard shows the same
  // bodies one step at a time
  function setupBodies(mode) {
    var m = state.meta;

    // the size option tables stay editable (label → days); editing options
    // flips the scheme to Custom. Story controls carry data-kind="story".
    function sizeOptRowsFor(kind) {
      var kAttr = kind === 'story' ? ' data-kind="story"' : '';
      var days = m[RM.sizeKeys(kind).days] || {};
      return RM.sizeOrderOf(state, kind).map(function (s2) {
        return '<tr>' +
          '<td><input data-suszlabel="' + esc(s2) + '"' + kAttr + ' value="' + esc(s2) + '" aria-label="Option label"></td>' +
          '<td><input type="number" min="0" step="0.5" data-susz="' + esc(s2) + '"' + kAttr + ' value="' + days[s2] + '" aria-label="Working days"></td>' +
          '<td class="hol-x"><button data-suszrm="' + esc(s2) + '"' + kAttr + ' title="Remove option"><i data-lucide="x"></i></button></td>' +
          '</tr>';
      }).join('');
    }
    function sizingCardsFor(kind) {
      var label = kind === 'story' ? esc(lvl('story') + ' size options') : 'Size options';
      if (kind !== 'story' && RM.sizeRollup(state)) {
        return '<section class="su-card"><h2>' + label + '</h2>' +
          '<div class="m-hint">' + esc('Each ' + lvl('feature').toLowerCase() + '\u2019s size is the sum of its ' + lvl('story', true).toLowerCase() +
          '\u2019 points, and its working days are their days added up. Sizes are not set by hand in this mode \u2014 size the ' + lvl('story', true).toLowerCase() + ' instead.') + '</div>' +
          '</section>';
      }
      return RM.sizingEnabled(state, kind)
        ? '<section class="su-card"><h2>' + label + '</h2>' +
          '<table class="hol-table"><thead><tr><th>Label</th><th>Working days</th><th></th></tr></thead>' +
          '<tbody>' + sizeOptRowsFor(kind) + '</tbody></table>' +
          '<div class="p-row" style="margin-top:8px"><button id="suSzAdd' + (kind === 'story' ? 'Story' : '') + '"><i data-lucide="plus"></i> Add option</button>' +
          (m[RM.sizeKeys(kind).scheme] === 'custom' && m[RM.sizeKeys(kind).base]
            ? '<button data-suszreset="' + esc(m[RM.sizeKeys(kind).base]) + '"' + (kind === 'story' ? ' data-kind="story"' : '') + '><i data-lucide="rotate-ccw"></i> Reset to ' +
              esc(RM.SIZE_SCHEMES[m[RM.sizeKeys(kind).base]].name) + '</button>'
            : '') + '</div>' +
          (kind === 'story'
            ? '<div class="m-hint">A sized story that lands on the timeline takes its size in working days; unsized stories take one sprint.</div>'
            : '<div class="m-hint">Working days drive scheduling and the duration estimate; the risk rating stays a separate L/M/H flag.</div>') +
          '</section>'
        : '<section class="su-card"><h2>' + label + '</h2>' +
          '<div class="m-hint">Sizing is off — set durations directly ' + (kind === 'story' ? 'on stories' : 'on items') + ' when you need them.</div>' +
          '</section>';
    }
    // priority and risk ladders are fixed: shown as chips, not edited
    function ladderCardFor(kind) {
      function line(label, sch, order, labelFn) {
        return '<div class="su-ladder"><label class="p-lab">' + esc(label) + '</label>' +
          (order.length ? ladderChips(order, labelFn) : '<div class="m-hint">' + esc(sch.desc) + '</div>') + '</div>';
      }
      return '<section class="su-card"><h2>' + esc(lvl(kind) + ' priority & risk') + '</h2>' +
        line('Priority', RM.PRIORITY_SCHEMES[RM.prioritySchemeOf(state, kind)], RM.priorityOrderOf(state, kind),
          function (v) { return priorityValueLabel(v, kind); }) +
        line(RM.riskColLabel(state, kind), RM.RISK_SCHEMES[RM.riskSchemeOf(state, kind)], RM.riskOrderOf(state, kind),
          function (v) { return riskValueLabel(v, kind); }) +
        '</section>';
    }

    // every range edits in place: name is free text, the dates open the calendar
    var holRows = (m.holidayRanges || []).map(function (r, i) {
      return '<tr>' +
        '<td><input data-suholname="' + i + '" value="' + esc(r.name || '') + '" placeholder="—" aria-label="Holiday name" title="Rename"></td>' +
        '<td class="hol-dates"><span class="hol-range">' +
        '<input type="text" readonly class="cal-in" data-suholstart="' + i + '" value="' + esc(r.start) + '" aria-label="First day" title="First day">' +
        '<span class="hol-arrow">→</span>' +
        '<input type="text" readonly class="cal-in" data-suholend="' + i + '" value="' + esc(r.end) + '" aria-label="Last day" title="Last day">' +
        '</span></td>' +
        '<td class="hol-x"><button data-suholrm="' + i + '" title="Remove"><i data-lucide="x"></i></button></td>' +
        '</tr>';
    }).join('');

    function grip() {
      return '<span class="su-grip"><i data-lucide="grip-vertical"></i></span>';
    }
    // the default workstream (null in the data) leads the list — renamable,
    // recolorable, not deletable and not draggable
    var defaultWsRow = (function () {
      var count = state.items.filter(function (x) { return !x.workstream; }).length;
      return '<div class="su-row su-defws">' +
        '<span class="dd-dot" style="background:#' + RM.defaultWsColor(state) + '"></span>' +
        '<span class="su-name">' + esc(RM.defaultWsName(state)) + '</span>' +
        '<span class="band-bucket-tag">default</span>' +
        '<span class="band-count">' + count + '</span>' +
        '<button data-sudefws title="Edit"><i data-lucide="pencil"></i></button>' +
        '</div>';
    })();
    var wsRows = allWorkstreams().map(function (w) {
      var count = state.items.filter(function (x) { return x.workstream === w; }).length;
      return '<div class="su-row" data-key="' + esc(w) + '">' + grip() +
        '<span class="dd-dot" style="background:#' + RM.colorForWs(state, w) + '"></span>' +
        '<span class="su-name">' + esc(w) + '</span>' +
        '<span class="band-count">' + count + '</span>' +
        '<button data-suwsedit="' + esc(w) + '" title="Edit"><i data-lucide="pencil"></i></button>' +
        '</div>';
    }).join('');

    var phaseRows = state.phases.map(function (ph) {
      var count = RM.itemsInPhase(state, ph.id).length;
      return '<div class="su-row" data-key="' + ph.id + '">' + grip() +
        '<span class="su-name">' + esc(ph.name) + '</span>' +
        (ph.bucket ? '<span class="band-bucket-tag">backlog</span>' : '') +
        '<span class="band-count">' + count + '</span>' +
        '<button data-suphedit="' + ph.id + '" title="Edit"><i data-lucide="pencil"></i></button>' +
        '<button data-suphdel="' + ph.id + '" class="danger" title="Delete"><i data-lucide="trash-2"></i></button>' +
        '</div>';
    }).join('');

    var typeCounts = {};
    state.team.forEach(function (mm) { typeCounts[mm.type] = (typeCounts[mm.type] || 0) + 1; });
    var DAY_NAMES = ['Sunday', 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday'];
    var firstDayName = DAY_NAMES[RM.weekStartOf(m)];
    // roles carry the RATE CARD: default hourly rate/cost that people of the
    // role inherit unless overridden on the person (Budgeting)
    var typeRows = state.teamTypes.map(function (t) {
      var rc = (m.rateCard && m.rateCard[t]) || {};
      return '<div class="su-row" data-key="' + esc(t) + '">' + grip() +
        '<input class="su-name su-name-in" data-rcname="' + esc(t) + '" value="' + esc(t) + '" title="Rename role">' +
        '<span class="band-count">' + (typeCounts[t] || 0) + '</span>' +
        '<input class="su-rc" type="number" min="0" data-rccost="' + esc(t) + '" value="' + (rc.cost || '') + '" placeholder="cost/h" title="Default hourly cost for this role">' +
        '<input class="su-rc" type="number" min="0" data-rcrate="' + esc(t) + '" value="' + (rc.rate || '') + '" placeholder="rate/h" title="Default hourly bill rate for this role">' +
        '<button data-suttrm="' + esc(t) + '" class="danger" title="Remove role"><i data-lucide="x"></i></button>' +
        '</div>';
    }).join('');

    var capTypeCounts = {};
    state.team.forEach(function (mm) { if (mm.capType) capTypeCounts[mm.capType] = (capTypeCounts[mm.capType] || 0) + 1; });
    var capTypeRows = RM.capTypesOf(state).map(function (t) {
      return '<div class="su-row" data-key="' + esc(t) + '">' + grip() +
        '<input class="su-name su-name-in" data-capname="' + esc(t) + '" value="' + esc(t) + '" title="Rename capacity type">' +
        '<span class="su-supplied">' + RM.capTypeRoles(state, t).map(function (r) {
          return '<span class="su-chip su-rolechip">' + esc(r) + '<button data-surolerm="' + esc(r) + '" title="Stop ' + esc(r) + ' supplying ' + esc(t) + '" aria-label="Remove ' + esc(r) + '">\u00d7</button></span>';
        }).join('') + '<button class="su-roleadd" data-suroleadd="' + esc(t) + '" title="Pick a role that supplies ' + esc(t) + '">+ Role</button></span>' +
        '<span class="band-count" title="People supplying it">' + (capTypeCounts[t] || 0) + '</span>' +
        '<button data-sucaprm="' + esc(t) + '" class="danger" title="Remove capacity type"><i data-lucide="x"></i></button>' +
        '</div>';
    }).join('');

    var epicRows = allEpics().map(function (ep) {
      var count = state.items.filter(function (x) { return x.epic === ep; }).length;
      var ico = RM.iconForEpic(state, ep);
      return '<div class="su-row" data-key="' + esc(ep) + '">' +
        '<span class="r-ico">' + (ico ? '<i data-lucide="' + ico + '"></i>' : '<i data-lucide="tag"></i>') + '</span>' +
        '<span class="su-name">' + esc(ep) + '</span>' +
        '<span class="band-count">' + count + '</span>' +
        '<button data-suepedit="' + esc(ep) + '" title="Edit"><i data-lucide="pencil"></i></button>' +
        '<button data-suepdel="' + esc(ep) + '" class="danger" title="Delete"><i data-lucide="trash-2"></i></button>' +
        '</div>';
    }).join('');

    // Hierarchy: one block per level, top to bottom; each type is listed
    // once, under the level it lives at
    var hier = m.hierarchy;
    var levelBlocks = hier.levels.map(function (lv, li) {
      var only = lv.types.length === 1;
      var rows = lv.types.map(function (k, ti) {
        var t = RM.itemType(state, k);
        if (!t) return '';
        return '<tr><td><button class="su-hicon" data-suhticon="' + esc(t.key) + '" title="Icon"><i data-lucide="' + esc(t.icon) + '"></i></button>' +
          '<input type="color" class="su-hcolor" data-suhtcolor="' + esc(t.key) + '" value="#' + esc(t.color) + '" title="Color (Color by item type)"></td>' +
          '<td><div class="su-hname"><input data-suhtlabel="' + esc(t.key) + '" value="' + esc(t.label) + '" aria-label="Type label">' +
          (ti === 0 ? '<span class="su-hdef" title="New ' + esc(lv.label.toLowerCase()) + ' items start as this type">default</span>' : '') + '</div></td>' +
          '<td><input data-suhtjira="' + esc(t.key) + '" value="' + esc(t.jira) + '" placeholder="Jira issue type" aria-label="Jira issue type"></td>' +
          '<td><select data-suhtlevel="' + esc(t.key) + '" aria-label="Level"' + (only ? ' disabled title="The only type at this level"' : '') + '>' +
          hier.levels.map(function (o) { return '<option value="' + o.key + '"' + (o.key === lv.key ? ' selected' : '') + '>' + esc(o.label) + '</option>'; }).join('') +
          '</select></td>' +
          '<td class="hol-x"><button data-suhtrm="' + esc(t.key) + '"' + (only ? ' disabled title="The only type at this level"' : ' title="Remove type"') + '><i data-lucide="x"></i></button></td></tr>';
      }).join('');
      return '<div class="su-hlv" data-suhlevel="' + lv.key + '">' +
        '<div class="su-hlv-head"><span class="su-hlv-n">Level ' + (li + 1) + '</span>' +
        '<input data-suhlabel="' + lv.key + '" value="' + esc(lv.label) + '" aria-label="Level ' + (li + 1) + ' name">' +
        (li ? '<span class="m-hint">under ' + esc(hier.levels[li - 1].label) + '</span>' : '') + '</div>' +
        '<table class="hol-table su-htypes">' + (li ? '' : '<thead><tr><th></th><th>Type</th><th>Jira issue type</th><th>Level</th><th></th></tr></thead>') +
        '<tbody>' + rows + '</tbody></table>' +
        '<button class="su-hadd" data-suhadd="' + lv.key + '"><i data-lucide="plus"></i> Add ' + esc(lv.label.toLowerCase()) + ' type</button></div>';
    }).join('');
    var hierCard =
      '<section class="su-card"><h2>Hierarchy</h2>' +
      '<div class="m-hint">Each item type lives at one level, and items hold children from the level below. The first type at a level is the default for new items.</div>' +
      '<div class="su-hlvs">' + levelBlocks + '</div>' +
      '<div class="m-hint">Behavior follows the level, not the type. The Jira issue type is what sync and the CSV export use.</div>' +
      '</section>';

    // the cards, then grouped into sections below
    var parts = {
      timeline:
        '<section class="su-card"><h2>Project</h2>' +
        '<label class="p-lab" for="suTitle">Name' + (wz ? ' <span class="req" aria-hidden="true">*</span>' : '') + '</label>' +
        '<input id="suTitle" style="width:100%" maxlength="120" value="' + esc(m.title || '') + '" placeholder="' + (wz ? 'e.g. Mobile app launch' : 'Roadmap name') + '"' +
        (wz ? ' required aria-required="true"' : '') + '>' +
        '</section>' +
        '<section class="su-card"><h2>Timeline</h2>' +
        '<div class="p-grid2">' +
        '<div><label class="p-lab">Start (' + firstDayName + ')</label><input type="text" readonly class="cal-in" id="suStart" value="' + esc(m.timelineStart) + '" style="width:100%"></div>' +
        '<div><label class="p-lab">End (last working day)</label><input type="text" readonly class="cal-in" id="suEnd" value="' + esc(m.endDate || '') + '" style="width:100%"></div>' +
        '</div>' +
        '</section>',
      sprints:
        '<section class="su-card"><h2>Sprints</h2>' +
        '<div><label class="p-lab">Sprint length</label>' +
        '<div class="seg">' + [[0, 'Off'], [1, '1 week'], [2, '2 weeks'], [3, '3 weeks'], [4, '4 weeks']].map(function (o) {
          return '<button data-suwps="' + o[0] + '"' + (m.weeksPerSprint === o[0] ? ' class="on"' : '') + '>' + o[1] + '</button>';
        }).join('') + '</div></div>' +
        (RM.sprintsEnabled(m)
          ? '<div class="p-grid2" style="margin-top:10px">' +
            '<div><label class="p-lab">Sprint starts on (' + firstDayName + ')</label><input type="text" readonly class="cal-in" id="suAnchor" value="' + esc(m.sprintAnchor || m.timelineStart) + '" style="width:100%"></div>' +
            '<div><label class="p-lab">…and is sprint #</label><input type="number" id="suAnchorNum" step="1" value="' + (m.sprintAnchorNum != null ? m.sprintAnchorNum : 1) + '" style="width:100%"></div>' +
            '</div>'
          : '<div class="m-hint">No sprint numbers on the timeline, and the Sprinting tab is hidden.</div>') +
        (RM.sprintsEnabled(m) ? sprintStripHtml() : '') +
        '</section>',
      holidays:
        '<section class="su-card"><h2>Holidays</h2>' +
        (holRows
          ? '<table class="hol-table"><thead><tr><th>Name</th><th>Dates</th><th></th></tr></thead><tbody>' + holRows + '</tbody></table>'
          : '<div class="m-hint">none</div>') +
        '<div class="hol-add">' +
        '<input id="suHolName" placeholder="Name (optional)">' +
        '<input type="text" readonly class="cal-in" id="suHolStart" placeholder="First day" aria-label="First day">' +
        '<input type="text" readonly class="cal-in" data-cal-clear="Single day" id="suHolEnd" placeholder="Last day" aria-label="Last day (optional)" title="Last day">' +
        '<button id="suHolAddBtn" class="fixed">Add</button>' +
        '</div>' +
        '<div class="m-hint">Single days or ranges (e.g. Christmas Eve through New Year’s). Non-working days stretch bars that span them.</div>' +
        '</section>',
      // Organization: three lists side by side, each with an Add box that
      // takes a pasted list (one per line); levels & item types fold away
      org:
        '<div class="su-org">' +
        '<section class="su-card su-org-col"><h2>Phases</h2>' +
        '<div class="su-rows" data-sulist="phase">' + phaseRows + '</div>' +
        addListHtml('suPhAddIn', 'suPhAddBtn', 'New phases, one per line') +
        '</section>' +
        '<section class="su-card su-org-col"><h2>Workstreams</h2>' +
        '<label class="p-check" title="Color-codes items and enables workstream grouping"><input type="checkbox" id="suWsEnable"' + (m.workstreamsEnabled ? ' checked' : '') + '> Use workstreams</label>' +
        (m.workstreamsEnabled
          ? '<div class="su-rows" style="margin-top:10px">' + defaultWsRow + '</div>' +
            '<div class="su-rows" data-sulist="ws">' + (wsRows || '') + '</div>' +
            addListHtml('suWsAdd', 'suWsAddBtn', 'New ones, one per line')
          : '<div class="m-hint">Workstreams are off — items keep a neutral color and the Scoping column is hidden.</div>') +
        '</section>' +
        '<section class="su-card su-org-col"><h2>Epics</h2>' +
        '<div class="su-rows">' + (epicRows || '<div class="m-hint">None yet.</div>') + '</div>' +
        addListHtml('suEpAdd', 'suEpAddBtn', 'New epics, one per line') +
        '</section>' +
        '</div>' +
        (orgHierOpen ? hierCard
          : '<section class="su-card su-fold"><div><h2>Levels &amp; item types</h2><div class="m-hint">' +
            esc(hier.levels.map(function (lv) { return lv.label; }).join(' › ')) + ' · ' + RM.itemTypes(state).length + ' item types</div></div>' +
            '<button data-suhieropen>Customize</button></section>'),
      team:
        '<section class="su-card"><h2>Roles &amp; rate card</h2>' +
        '<div class="su-rc-head"><span></span><span>Cost/h</span><span>Rate/h</span><span></span></div>' +
        '<div class="su-rows" data-sulist="type">' + typeRows + '</div>' +
        '<div class="p-row" style="margin-top:8px"><input id="suTypeAdd" placeholder="New role, e.g. Data Scientist"><button id="suTypeAddBtn" class="fixed">Add</button></div>' +
        '<div class="m-hint">People inherit their role’s hourly cost and bill rate; either can be overridden per person in Budgeting.</div>' +
        '</section>',
      workweek:
        '<section class="su-card"><h2>Work week</h2>' +
        '<div class="p-grid2">' +
        '<div><label class="p-lab">Full-time hours per week</label>' +
        '<input type="number" id="suWeekHours" min="1" max="80" value="' + m.weekHours + '" style="width:100%"></div>' +
        '<div><label class="p-lab">First day of week</label>' +
        '<select id="suWeekStart" style="width:100%">' + DAY_NAMES.map(function (dn, di) {
          return '<option value="' + di + '"' + (RM.weekStartOf(m) === di ? ' selected' : '') + '>' + dn + '</option>';
        }).join('') + '</select></div>' +
        '</div>' +
        '<div style="margin-top:12px"><label class="p-lab">Working days</label>' +
        '<div class="su-wdays">' + DAY_NAMES.map(function (dn, di) {
          var onDay = RM.workDaysOf(m).indexOf(di) !== -1;
          return '<label class="p-check wd"><input type="checkbox" data-suwday="' + di + '"' +
            (onDay ? ' checked' : '') + '> ' + dn.slice(0, 3) + '</label>';
        }).join('') + '</div></div>' +
        '<div class="m-hint">Defines what one full-time person means — the schedule plans across exactly the days checked (1–7); bars keep their calendar dates when this changes.</div>' +
        '</section>',
      capacity:
        '<section class="su-card"><h2>Schedule against capacity</h2>' +
        '<label class="p-check" title="The roster limits scheduling and validation; shows the capacity row"><input type="checkbox" id="suCapEnable"' + (m.capacityEnabled ? ' checked' : '') + '> Schedule against capacity</label>' +
        '<div class="m-hint">People and their weekly hours live in Setup \u203a Team and the Resources panel; each person supplies their role\u2019s capacity type (Supplied by, under Capacity types). Auto timeline (per phase) and Place at earliest slot need this on.</div>' +
        '</section>' +
        '<section class="su-card"><h2>Planning level</h2><div class="su-schemes">' +
        [['feature', esc(lvl('feature', true)), 'Capacity follows the ' + esc(lvl('feature', true).toLowerCase()) + ' and their capacity type; ' + esc(lvl('story', true).toLowerCase()) + ' need no details'],
         ['story', esc(lvl('story', true)), esc(lvl('feature')) + ' bars become the hull of their ' + esc(lvl('story', true).toLowerCase()) + ' \u2014 work is planned on the ' + esc(lvl('story', true).toLowerCase()) + ' and their capacity types']].map(function (o) {
          var on = RM.planLevel(state) === o[0];
          return '<button class="su-scheme' + (on ? ' on' : '') + '" data-suplan="' + o[0] + '">' +
            '<span class="su-scheme-check"><i data-lucide="' + (on ? 'circle-check' : 'circle') + '"></i></span>' +
            '<span class="su-scheme-main"><b>' + o[1] + '</b><span>' + o[2] + '</span></span></button>';
        }).join('') + '</div></section>' +
        '<section class="su-card"><h2>Demand</h2><div class="su-schemes">' +
        [['person', 'Per person', 'A unit in flight uses one person of its capacity type, times its multiplier'],
         ['points', 'Story points', 'A unit\u2019s points spread over its weeks; each person supplies points per sprint']].map(function (o) {
          var on = (m.capMode || 'person') === o[0];
          return '<button class="su-scheme' + (on ? ' on' : '') + '" data-sucapmode="' + o[0] + '">' +
            '<span class="su-scheme-check"><i data-lucide="' + (on ? 'circle-check' : 'circle') + '"></i></span>' +
            '<span class="su-scheme-main"><b>' + o[1] + '</b><span>' + o[2] + '</span></span></button>';
        }).join('') + '</div>' +
        (m.capMode === 'points'
          ? '<div style="margin-top:10px"><label class="p-lab">Default points per person per sprint</label>' +
            '<input type="number" id="suDefPoints" min="0" step="1" value="' + m.defaultPoints + '" style="width:140px">' +
            '<div class="m-hint">Each person can override this in the Resources panel. With sprints off, points are per two weeks.</div></div>'
          : '') +
        '</section>' +
        '<section class="su-card"><h2>Capacity types</h2>' +
        '<div class="su-rows" data-sulist="captype">' + capTypeRows + '</div>' +
        '<div class="p-row" style="margin-top:8px"><input id="suCapTypeAdd" placeholder="New capacity type, e.g. Data"><button id="suCapTypeAddBtn" class="fixed">Add</button></div>' +
        '<div class="m-hint">A ' + esc(lvl('story').toLowerCase()) + '\u2019s capacity type says what it drains and who can take it. Each role supplies one type (Supplied by) and its people supply it. Drag the grips to reorder.</div>' +
        '</section>' + schedExplainerHtml(),
      columns: (function () {
        var offNotes = [];
        if (!RM.sizingEnabled(state) && !RM.sizingEnabled(state, 'story')) offNotes.push('Size (enable in Sizing, priority & risk)');
        if (!RM.riskEnabled(state) && !RM.riskEnabled(state, 'story')) offNotes.push('Risk (pick a scheme in Sizing, priority & risk)');
        if (!RM.priorityEnabled(state) && !RM.priorityEnabled(state, 'story')) offNotes.push('Priority (pick a scheme in Sizing, priority & risk)');
        if (!state.meta.workstreamsEnabled) offNotes.push('Workstream (enable in Organization)');
        var colRows = allScopeCols().map(function (c) {
          var fixed = isFixedColKey(c[0]);
          return '<div class="su-row" data-key="' + esc(c[0]) + '">' + grip() +
            (fixed
              ? '<span class="su-name">' + esc(c[1]) + '</span><span class="band-bucket-tag">built-in</span>'
              : '<input class="su-name su-name-in" data-sucolname="' + esc(c[0]) + '" value="' + esc(c[1]) + '" title="Rename column">') +
            (fixed ? '' : '<select data-sucolscope="' + esc(c[0]) + '" title="Which rows show this column">' +
              ['both', 'feature', 'story'].map(function (sc) {
                return '<option value="' + sc + '"' + (((scopeColDef(c[0]) || {}).scope || 'both') === sc ? ' selected' : '') + '>' +
                  esc(scopeScopeLabel(sc)) + '</option>';
              }).join('') + '</select>') +
            (fixed
              ? '<button data-sucoldel="' + esc(c[0]) + '" class="danger" disabled title="Built-in columns can\u2019t be deleted"><i data-lucide="x"></i></button>'
              : '<button data-sucoldel="' + esc(c[0]) + '" class="danger" title="Delete column"><i data-lucide="x"></i></button>') +
            '</div>';
        }).join('');
        return '<section class="su-card"><h2>Scoping columns</h2>' +
          '<div class="su-rows" data-sulist="scol">' + colRows + '</div>' +
          '<div class="p-row" style="margin-top:8px"><input id="suColAdd" placeholder="New column, e.g. Owner"><button id="suColAddBtn" class="fixed">Add</button></div>' +
          '<div class="m-hint">Drag to reorder — the Scoping grid and the panel follow this order. Text columns rename, pick which rows show them (features, stories, or both) and remove here; built-ins carry fixed data.' +
          (offNotes.length ? '<br>Hidden right now: ' + esc(offNotes.join(', ')) + '.' : '') + '</div>' +
          '</section>';
      })(),
      sizing:
        estGridHtml() +
        sizingCardsFor('feature') + ladderCardFor('feature') +
        sizingCardsFor('story') + ladderCardFor('story'),
      views:
        '<section class="su-card"><h2>Views</h2>' +
        '<div class="m-hint" style="margin:0 0 10px">Which tabs this project shows. Turning a view off hides its tab; its data stays in the document.</div>' +
        RM.APPS.map(function (a) {
          // Planning is home; Sprinting and Budgeting follow their sections
          var fixedPill = a[0] === 'planning' ? 'Always on'
            : a[0] === 'sprints' ? 'Follows Sprints' : a[0] === 'budget' ? 'Follows Budgeting' : '';
          var ctl = fixedPill
            ? '<span class="su-pill">' + fixedPill + '</span>'
            : '<input type="checkbox" data-suapp="' + a[0] + '"' + (RM.appEnabled(state, a[0]) ? ' checked' : '') + '> ';
          return '<label class="p-check su-app' + (fixedPill ? ' fixed' : '') + '" data-suview="' + a[0] + '" title="' + esc(a[3]) + '">' +
            (fixedPill ? '<span class="su-app-sp"></span>' : ctl) +
            '<i data-lucide="' + a[2] + '"></i>' + esc(a[1]) + '<span class="su-app-desc">' + esc(a[3]) + '</span>' +
            (fixedPill ? ctl : '') + '</label>';
        }).join('') +
        '</section>',
      appearance:
        '<section class="su-card">' + personalFieldsHtml('appearance') +
        '<div class="m-hint">System follows your OS.</div>' +
        '</section>',
      prefs:
        '<section class="su-card">' + personalFieldsHtml('behavior') + '</section>',
      ai:
        '<section class="su-card" id="aiSettingsCard">' +
        (window.HeadwayAI ? HeadwayAI.settingsHtml() : '<div class="m-hint">AI module not loaded.</div>') +
        '</section>',
      jira:
        '<section class="su-card" id="jiraSettingsCard">' +
        (window.HeadwayJira ? HeadwayJira.settingsHtml() : '<div class="m-hint">Jira module not loaded.</div>') +
        '</section>'
    };
    return {
      project: mode === 'wizard'
        ? parts.timeline + presetCardsHtml() + (wz && wz.expandWorkWeek ? parts.workweek + parts.holidays : workWeekSummaryHtml())
        : parts.timeline + parts.workweek + parts.holidays,
      sprints: parts.sprints,
      org: parts.org,
      est: parts.sizing,
      budget: '<section class="su-card">' +
        '<label class="p-check"><input type="checkbox" data-subudget' + (RM.appEnabled(state, 'budget') ? ' checked' : '') + '> Track budget</label>' +
        '<div class="m-hint">Adds the Budgeting tab: hours and cost by role, week and phase.</div></section>' +
        (RM.appEnabled(state, 'budget') ? parts.team : ''),
      team: '<section class="su-card">' + teamTableHtml() + '</section>',
      scheduling: parts.capacity,
      columns: parts.columns,
      views: parts.views,
      appearance: parts.appearance,
      prefs: parts.prefs,
      ai: parts.ai,
      jira: parts.jira
    };
  }
  // Scheduling: pick a role to supply this capacity type (a role supplies one)
  function roleSupplyMenu(anchor, capType) {
    var map = state.roleCapTypes || {};
    if (!state.teamTypes.length) { openDropdown(anchor, [{ label: '<i>No roles yet \u2014 add them on the Team section</i>', fn: function () {} }]); return; }
    openDropdown(anchor, state.teamTypes.map(function (r) {
      var cur = map[r];
      return { label: esc(r) + (cur && cur !== capType ? ' <small>now ' + esc(cur) + '</small>' : ''), checked: cur === capType,
        fn: function () { commit('role capacity type', function (s2) { RM.setRoleCapType(s2, r, capType); }); } };
    }));
  }
  // an Add box that takes one name or a pasted list (one per line)
  function addListHtml(inId, btnId, ph) {
    return '<div class="p-row su-addlist"><textarea id="' + inId + '" rows="1" placeholder="' + esc(ph) + '"></textarea>' +
      '<button id="' + btnId + '" class="fixed">Add</button></div>';
  }
  function addListNames(v, existing) {
    var out = [];
    String(v || '').split(/\r?\n/).forEach(function (l) {
      l = l.trim();
      if (l && out.indexOf(l) === -1 && (existing || []).indexOf(l) === -1) out.push(l);
    });
    return out;
  }
  var orgHierOpen = false; // Setup → Organization: Levels & item types unfolded
  function sectionBody(key, mode) { return setupBodies(mode)[key] || ''; }
  // Setup → Team: everyone on the project, every field editable but the
  // capacity type, which comes from the role (Setup → Scheduling)
  var teamNewRoleFor = null; // member id whose Role cell is naming a new role
  function teamTableHtml() {
    var cap = !!state.meta.capacityEnabled;
    var pts = cap && state.meta.capMode === 'points';
    var bud = RM.appEnabled(state, 'budget');
    var wsOn = !!state.meta.workstreamsEnabled;
    var head = ['Name', 'Title', 'Role'].concat(cap ? ['Capacity type'] : [], wsOn ? ['Workstreams'] : [], ['Allocation'],
      pts ? ['Points / sprint'] : [], bud ? ['Rate/h', 'Cost/h'] : [], ['']);
    function roleOpts(cur) {
      return '<option value="">\u2014</option>' + state.teamTypes.map(function (r) {
        return '<option value="' + esc(r) + '"' + (r === cur ? ' selected' : '') + '>' + esc(r) + '</option>';
      }).join('') + '<option value="__new">New role\u2026</option>';
    }
    var rc = state.meta.rateCard || {};
    if (!state.team.length) {
      return '<div class="m-hint">No one yet \u2014 add the people who work on this project.</div>' +
        '<button id="suTmAdd" style="margin-top:8px"><i data-lucide="plus"></i> Add person</button>';
    }
    return '<div class="su-team-wrap"><table class="hol-table su-team"><thead><tr>' + head.map(function (h) { return '<th>' + h + '</th>'; }).join('') + '</tr></thead><tbody>' +
      state.team.map(function (m) {
        var id = ' data-mid="' + esc(m.id) + '"';
        var card = rc[m.type] || {};
        var roleCell = teamNewRoleFor === m.id
          ? '<input data-sutmnewrole' + id + ' placeholder="New role name" aria-label="New role name">'
          : '<select data-sutm="type"' + id + ' aria-label="Role">' + roleOpts(m.type) + '</select>';
        return '<tr' + id + '>' +
          '<td><input data-sutm="name"' + id + ' value="' + esc(m.name) + '" placeholder="Name" aria-label="Name"></td>' +
          '<td><input data-sutm="role"' + id + ' value="' + esc(m.role || '') + '" placeholder="Title" aria-label="Title"></td>' +
          '<td>' + roleCell + '</td>' +
          (cap ? '<td class="su-tm-cap" title="' + (m.type ? 'From the role \u2014 change it in Setup \u203a Scheduling' : 'Pick a role to set it') + '">' +
            esc(m.capType || '\u2014') + '</td>' : '') +
          (wsOn ? '<td class="su-tm-ws"><button class="su-tm-wsbtn' + (RM.memberWorkstreams(m).length ? '' : ' empty') + '" data-sutmws' + id +
            ' title="Workstreams">' + esc(RM.memberWorkstreams(m).join(', ') || '\u2014') + '</button></td>' : '') +
          '<td class="su-tm-num"><input type="number" min="0" step="5" data-sutm="capacity"' + id + ' value="' + Math.round((m.capacity != null ? m.capacity : 1) * 100) + '" aria-label="Allocation percent"> %</td>' +
          (pts ? '<td class="su-tm-num"><input type="number" min="0" data-sutm="points"' + id + ' value="' + (m.points == null ? '' : m.points) + '" placeholder="' + state.meta.defaultPoints + '" aria-label="Points per sprint"></td>' : '') +
          (bud ? '<td class="su-tm-num"><input type="number" min="0" data-sutm="rate"' + id + ' value="' + (m.rate || '') + '" placeholder="' + (card.rate || '') + '" aria-label="Hourly rate"></td>' +
                 '<td class="su-tm-num"><input type="number" min="0" data-sutm="cost"' + id + ' value="' + (m.cost || '') + '" placeholder="' + (card.cost || '') + '" aria-label="Hourly cost"></td>' : '') +
          '<td class="hol-x"><button data-sutmdel="' + esc(m.id) + '" title="Remove person"><i data-lucide="x"></i></button></td></tr>';
      }).join('') + '</tbody></table></div>' +
      '<button id="suTmAdd" style="margin-top:8px"><i data-lucide="plus"></i> Add person</button>' +
      '<div class="m-hint">Allocation here is each person\u2019s default. You can change allocation and hours week by week later in the Resources panel under the timeline.</div>';
  }
  function renderSetupHost() { if (wz) renderWizard(); else renderSetup(); }
  // Setup → Sprints: the next six sprints from the anchor, with their dates
  function sprintStripHtml() {
    var m = state.meta, si = RM.sprintInfo(m), wd = RM.workDaysOf(m);
    var out = [];
    for (var i = 0; i < 6; i++) {
      var wk = si.anchorWeek + i * si.wps;
      var start = RM.weekStartDate(m, wk), end = RM.weekStartDate(m, wk + si.wps);
      end.setUTCDate(end.getUTCDate() - 1);
      for (var g = 0; g < 7 && wd.indexOf(end.getUTCDay()) === -1; g++) end.setUTCDate(end.getUTCDate() - 1);
      out.push('<span class="su-spr"><b>S' + RM.sprintNumForWeek(m, wk) + '</b>' + RM.fmtShort(start) + ' – ' + RM.fmtShort(end) + '</span>');
    }
    return '<div class="su-sprs" aria-label="Upcoming sprints">' + out.join('') + '</div>';
  }
  // wizard Project step: work week and holidays fold into one line
  function workWeekSummaryHtml() {
    var m = state.meta;
    var DN = ['Sun', 'Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat'];
    var days = RM.workDaysOf(m).map(function (d) { return DN[d]; }).join(', ');
    var nHol = (m.holidayRanges || []).length;
    return '<section class="su-card wz-fold"><div><h2>Work week &amp; holidays</h2><div class="m-hint">' + esc(days) + ' · ' +
      m.weekHours + ' h per week · ' + nHol + (nHol === 1 ? ' holiday' : ' holidays') + '</div></div>' +
      '<button data-wzexpand>Edit</button></section>';
  }
  // wizard Project step: the required preset, each card with a tiny timeline
  function pcBar(cls, x, w, y, tx) { return '<span class="pc-' + cls + '" style="left:' + x + '%;width:' + w + '%;top:' + y + 'px">' + esc(tx || '') + '</span>'; }
  function pcFeat(x, w, y, tx) { return pcBar('feat', x, w, y, tx); }
  function pcHull(x, w, y, tx) { return pcBar('hull', x, w, y, tx); }
  function pcStory(x, w, y, tx) { return pcBar('story', x, w, y, tx); }
  function pcRisk(x, y) { return pcBar('risk', x, 2.2, y, ''); }
  function pcSprints() { return [0, 1, 2, 3, 4].map(function (i) { return pcBar('hdr', 1 + i * 19.6, 19, 4, 'S' + (i + 1)); }).join(''); }
  function pcMonths() { return ['Oct', 'Nov', 'Dec', 'Jan', 'Feb'].map(function (mo, i) { return pcBar('mon', 1 + i * 19.6, 19, 4, mo); }).join(''); }
  var PC_SKETCH = {
    scrum: pcSprints() + pcHull(1, 44, 26, 'Checkout') + pcStory(1, 13, 46, '3') + pcStory(15, 11, 46, '5') + pcStory(27, 18, 46, '2') +
      pcHull(40, 56, 66, 'Search') + pcStory(40, 22, 86, '8') + pcStory(63, 16, 86, '3'),
    ascrum: pcSprints() + pcFeat(1, 44, 26, 'Checkout · L · Must') + pcRisk(42, 30) + pcStory(1, 13, 46, '3 H') + pcStory(15, 11, 46, '5 M') +
      pcStory(27, 18, 46, '2 L') + pcFeat(40, 56, 66, 'Search · XL · Should') + pcRisk(93, 70) + pcStory(40, 22, 86, '8 M') + pcStory(63, 16, 86, '3 L'),
    rapid: pcMonths() + pcFeat(1, 30, 26, 'Login · M') + pcFeat(32, 44, 26, 'Billing · L') + pcFeat(1, 18, 48, 'Export · S') +
      pcFeat(20, 30, 48, 'Search · M') + pcFeat(51, 26, 48, 'Alerts · M') + pcFeat(78, 20, 70, 'Admin · S'),
    contract: pcMonths() + pcFeat(1, 30, 26, 'Login · M · Must') + pcFeat(32, 44, 26, 'Billing · L · Must') + pcRisk(73, 30) +
      pcFeat(1, 30, 48, 'Export · S · Could') + pcFeat(32, 34, 48, 'Alerts · M · Should') + pcRisk(63, 52) + pcFeat(67, 31, 70, 'Admin · S · Won’t'),
    innovation: pcMonths() + pcFeat(1, 30, 26, 'Pilot · RICE 84') + pcRisk(28, 30) + pcFeat(32, 40, 26, 'Chat assist · RICE 61') +
      pcFeat(1, 36, 48, 'Voice · RICE 40') + pcRisk(34, 52) + pcFeat(38, 30, 48, 'Offline · RICE 22') + pcFeat(69, 29, 70, 'AR view · RICE 9'),
    minimal: pcMonths() + pcFeat(1, 30, 26, 'Login · M') + pcFeat(24, 44, 46, 'Billing · L') + pcFeat(8, 18, 66, 'Export · S') +
      pcFeat(60, 36, 86, 'Search · M')
  };
  function presetCardsHtml() {
    var cur = state.meta.preset;
    return '<section class="su-card"><h2>Start from a preset <span class="req" aria-hidden="true">*</span></h2>' +
      '<div class="m-hint">Pick how you work. It fills in the next steps, and you can change any of it on the way.</div>' +
      '<div class="pc-grid" role="radiogroup" aria-label="Preset">' + RM.PRESETS.map(function (p, i) {
        var on = cur === p.key;
        var stop = on || (!cur && i === 0); // one tab stop; arrows move within
        return '<button class="pcard' + (on ? ' on' : '') + '" role="radio" aria-checked="' + on + '" tabindex="' + (stop ? 0 : -1) + '" data-wzpreset="' + p.key + '">' +
          (on ? '<span class="pc-check" aria-hidden="true">✓</span>' : '') +
          '<b>' + esc(p.name) + '</b><span class="pc-desc">' + esc(p.desc) + '</span>' +
          '<span class="pc-sketch" aria-hidden="true">' + PC_SKETCH[p.key] + '</span></button>';
      }).join('') + '</div>' + (cur ? '' : '<div class="m-hint req-hint">Choose a preset to continue.</div>') + '</section>';
  }
  // Setup → Scheduling: supply and demand, drawn on a small fixed example
  function schedExplainerHtml() {
    var SUP = [2, 2, 2, 1.2, 2, 2, 2, 2, 2, 2];
    function chart(title, note, dem) {
      return '<div class="su-sched-chart"><div><b>' + title + '</b> <span class="m-hint">' + note + '</span></div><div class="su-sched-bars">' +
        dem.map(function (d, i) {
          var fit = Math.min(d, SUP[i]), over = Math.max(0, d - SUP[i]);
          return '<div class="su-sched-wk"><i class="over" style="height:' + Math.round(over * 36) + 'px"></i><i class="ok" style="height:' + Math.round(fit * 36) + 'px"></i>' +
            '<i class="sup" style="bottom:' + Math.round(SUP[i] * 36) + 'px"></i></div>';
        }).join('') + '</div></div>';
    }
    return '<section class="su-card"><h2>How scheduling works</h2>' +
      '<div class="m-hint">Every week, for each capacity type, the work in flight has to fit inside what the people supply.</div>' +
      '<div class="su-sched-keys">' +
      '<div><b>Supply</b><div class="m-hint">Each person gives their role\u2019s capacity type their weekly hours, or points per sprint. Holidays and time off lower it.</div></div>' +
      '<div><b>Demand</b><div class="m-hint">Each ' + esc(lvl(RM.planLevel(state)).toLowerCase()) + ' in flight asks its capacity type for one person \u00d7 its multiplier, or its points spread over its weeks.</div></div>' +
      '<div><b>Over capacity</b><div class="m-hint">Demand above the line shows red on the timeline. Auto timeline moves work to the first week with room, after its dependencies.</div></div></div>' +
      '<div class="su-sched-charts">' + chart('As drawn', 'weeks 3\u20134 over capacity', [1, 2, 3, 3, 1, 1, 0, 0, 0, 0]) +
      chart('After Auto timeline', 'overflow moved to weeks 5\u20137', [1, 2, 2, 1.2, 2, 2, 0.8, 0, 0, 0]) + '</div></section>';
  }
  // Sizing, priority & risk: one select per field × level
  function schemeSelect(what, level, cur, options) {
    return '<select data-suscheme-kind="' + what + '" data-kind="' + level + '" aria-label="' + esc(lvl(level) + ' ' + what) + '">' +
      options.map(function (o) {
        return '<option value="' + o[0] + '"' + (o[0] === cur ? ' selected' : '') + '>' + esc(o[1]) + '</option>';
      }).join('') + '</select>';
  }
  function estGridHtml() {
    var m = state.meta;
    function sizeOpts(kind) {
      var o = RM.sizeSchemesFor(kind).map(function (k) { return [k, RM.SIZE_SCHEMES[k].name]; });
      if (m[RM.sizeKeys(kind).scheme] === 'custom') o.push(['custom', 'Custom']);
      return o;
    }
    function prioOpts(kind) {
      return (kind === 'story' ? RM.STORY_PRIORITY_SCHEME_ORDER : RM.PRIORITY_SCHEME_ORDER)
        .map(function (k) { return [k, RM.PRIORITY_SCHEMES[k].name]; });
    }
    function riskOpts(kind) {
      return RM.RISK_SCHEME_ORDER.filter(function (k) { return kind !== 'story' || RM.STORY_RISK_SCHEMES.indexOf(k) !== -1; })
        .map(function (k) { return [k, RM.RISK_SCHEMES[k].name]; });
    }
    var rows = [
      ['Size', 'Sets bar length', 'size', function (k) { return m[RM.sizeKeys(k).scheme]; }, sizeOpts],
      ['Priority', 'Ranks the backlog', 'prio', function (k) { return RM.prioritySchemeOf(state, k); }, prioOpts],
      ['Risk', 'Flags what may slip', 'risk', function (k) { return RM.riskSchemeOf(state, k); }, riskOpts]
    ];
    return '<section class="su-card"><div class="su-grid"><span></span><b>' + esc(lvl('feature')) + ' level</b><b>' + esc(lvl('story')) + ' level</b>' +
      rows.map(function (r) {
        return '<div><b>' + r[0] + '</b><div class="m-hint">' + r[1] + '</div></div>' +
          schemeSelect(r[2], 'feature', r[3]('feature'), r[4]('feature')) +
          schemeSelect(r[2], 'story', r[3]('story'), r[4]('story'));
      }).join('') + '</div></section>';
  }
  function ladderChips(order, labelFn) {
    return '<div class="su-chips">' + order.map(function (v) { return '<span class="su-chip">' + esc(labelFn(v)) + '</span>'; }).join('') + '</div>';
  }
  // the rail's on/off note for sections that switch a feature
  function setupPill(key) {
    if (key === 'sprints') return RM.sprintsEnabled(state.meta) ? state.meta.weeksPerSprint + ' wk' : 'Off';
    if (key === 'budget') return RM.appEnabled(state, 'budget') ? 'On' : 'Off';
    if (key === 'scheduling') return state.meta.capacityEnabled ? 'On' : 'Off';
    return '';
  }

  function renderSetup() {
    if (wz) { renderWizard(); return; }
    var host = $('#setupView');
    setupTab = normSetupTab(setupTab);
    var tabBodies = setupBodies();

    // view-only copies keep the theme and nothing else
    var sections = readOnly
      ? SETUP_SECTIONS.map(function (sec) {
          return [sec[0], sec[1].filter(function (t) { return t[0] === 'appearance'; })];
        }).filter(function (sec) { return sec[1].length; })
      : SETUP_SECTIONS;
    if (readOnly) setupTab = 'appearance';
    var rail = sections.map(function (sec) {
      return (sec[0] ? '<div class="su-rail-hd">' + esc(sec[0]) + '</div>' : '') +
        sec[1].map(function (t) {
          var pill = setupPill(t[0]);
          return '<button class="su-tab' + (setupTab === t[0] ? ' on' : '') + '" data-sutab="' + t[0] + '">' +
            '<i data-lucide="' + t[2] + '"></i>' + esc(t[1]) + (pill ? '<span class="su-pill">' + pill + '</span>' : '') + '</button>';
        }).join('');
    }).join('');

    // page title = the active tab's rail label (typography hierarchy:
    // page h1 > card h2 > field sub-headings)
    var pageTitle = '';
    sections.forEach(function (sec) {
      sec[1].forEach(function (t) { if (t[0] === setupTab) pageTitle = t[1]; });
    });

    host.innerHTML =
      '<div class="su-layout">' +
      '<nav class="su-rail" aria-label="Settings sections">' + rail + '</nav>' +
      '<div class="su-content"><h1 class="su-page">' + esc(pageTitle) + '</h1>' + tabBodies[setupTab] + '</div>' +
      '</div>';
    var nrIn = teamNewRoleFor && host.querySelector('[data-sutmnewrole]');
    if (nrIn) nrIn.focus();
    if (setupTab === 'ai' && window.HeadwayAI) HeadwayAI.wireSettings($('#aiSettingsCard', host));
    if (setupTab === 'jira' && window.HeadwayJira) HeadwayJira.wireSettings($('#jiraSettingsCard', host));
    if (window.lucide) lucide.createIcons();
  }

  var SETUP_SECTIONS = [
    ['', [
      ['project', 'Project', 'calendar-range'],
      ['sprints', 'Sprints', 'repeat'],
      ['org', 'Organization', 'layers'],
      ['est', 'Sizing, priority & risk', 'ruler'],
      ['budget', 'Budgeting', 'wallet'],
      ['team', 'Team', 'users'],
      ['scheduling', 'Scheduling', 'gauge'],
      ['columns', 'Custom columns', 'columns-3'],
      ['views', 'Views', 'layout-grid']
    ]],
    ['Personal \u00b7 this computer only', [
      ['appearance', 'Appearance', 'palette'],
      ['prefs', 'Preferences', 'sliders-horizontal'],
      ['ai', 'AI assistant', 'sparkles'],
      ['jira', 'Jira Integration', 'link']
    ]]
  ];
  // section keys from before the regroup (UI snapshots, links)
  var SETUP_KEY_MAP = { timeline: 'project', phases: 'org', workstreams: 'org', sizing: 'est',
    capacity: 'scheduling', apps: 'views' };
  function normSetupTab(k) {
    k = SETUP_KEY_MAP[k] || k;
    var all = [];
    SETUP_SECTIONS.forEach(function (sec) { sec[1].forEach(function (t) { all.push(t[0]); }); });
    return all.indexOf(k) !== -1 ? k : 'project';
  }

  // ------------------------------------------------------------ new-project wizard
  // Full screen under the app's top bar. While it is open the global `state`
  // IS a draft project: the Setup section bodies and their handlers edit it
  // unchanged, and commit() only mutates it and re-renders the wizard. The
  // open document (with its undo, redo and saved flag) waits in wz.stash.
  var wz = null;
  var WZ_STEPS = SETUP_SECTIONS[0][1].slice(0, 8).map(function (t) { return t[0]; }).concat(['review']);
  var ONBOARDED_KEY = 'headway-onboarded-v1';
  function localStorageGet(k) { try { return localStorage.getItem(k); } catch (e) { return null; } }
  function openWizard() {
    if (wz || readOnly) return;
    flushPanelEdit(); // a half-typed panel field belongs to the open document
    closePopover();
    wz = { step: 'project', dirty: false, welcomed: false, stash: {
      state: state, undo: undoStack.slice(), redo: redoStack.slice(),
      docSaved: docSaved, sessionEdited: sessionEdited, view: view, reloaded: false,
      start: document.body.classList.contains('start') } };
    // the start page hides the top bar; the wizard needs it (drag, close)
    document.body.classList.remove('start');
    state = blankState();
    state.meta.title = ''; // required: the user names the project
    undoStack.length = 0; redoStack.length = 0;
    // the first project on this machine starts with a Welcome (name + theme)
    if (!userName() && !localStorageGet(ONBOARDED_KEY)) { wz.step = 'welcome'; wz.hasWelcome = true; }
    $('#setupView').innerHTML = ''; // one copy of the section controls on screen
    document.body.classList.add('wizard-open');
    var tb = $('#topbar');
    var x = document.createElement('button');
    x.id = 'wzClose';
    x.title = 'Close';
    x.setAttribute('aria-label', 'Close');
    x.innerHTML = '<i data-lucide="x"></i>';
    tb.insertBefore(x, $('.win-caption', tb) || null);
    // measured with the close button in place (it sets the bar's height),
    // so the wizard — and its progress bar — start right below it
    document.documentElement.style.setProperty('--topbar-h', (tb.offsetHeight || 46) + 'px');
    $('#docTitle').value = 'New project';
    // the progress bar stays put under the top bar; only its fill moves
    $('#wizard').innerHTML = '<div class="wz-progress" id="wzProgress" role="progressbar" aria-label="New project progress" aria-valuemin="0" aria-valuemax="100">' +
      '<i></i></div><div class="wz-body"></div>';
    $('#wizard').hidden = false;
    renderWizard();
    if (window.HeadwayDesktop && HeadwayDesktop.syncMenu) HeadwayDesktop.syncMenu();
  }
  // the desktop watcher reloaded the open file while the wizard is up: the
  // stash takes it, and closing adopts it the way any reload does
  function wizardStashReload(next) {
    wz.stash.state = next;
    wz.stash.docSaved = true;
    wz.stash.reloaded = true;
  }
  // AI / Jira hooks act on the open document even while the wizard is up:
  // run fn with the stash swapped in, then put the draft back
  function withStash(fn) {
    var w = wz, draft = state, du = undoStack.slice(), dr = redoStack.slice();
    state = w.stash.state;
    undoStack.length = 0; Array.prototype.push.apply(undoStack, w.stash.undo);
    redoStack.length = 0; Array.prototype.push.apply(redoStack, w.stash.redo);
    wz = null;
    try { return fn(); } finally {
      w.stash.state = state;
      w.stash.undo = undoStack.slice(); w.stash.redo = redoStack.slice();
      w.stash.docSaved = docSaved; w.stash.sessionEdited = sessionEdited;
      state = draft;
      undoStack.length = 0; Array.prototype.push.apply(undoStack, du);
      redoStack.length = 0; Array.prototype.push.apply(redoStack, dr);
      wz = w;
      $('#setupView').innerHTML = ''; // one copy of the section controls on screen
      document.body.classList.remove('start');
      renderWizard();
    }
  }
  // discard: true drops the draft; false hands it back (Create)
  function closeWizard(discard) {
    if (!wz) return null;
    var st = wz.stash, draft = state;
    state = st.state;
    undoStack.length = 0; Array.prototype.push.apply(undoStack, st.undo);
    redoStack.length = 0; Array.prototype.push.apply(redoStack, st.redo);
    docSaved = st.docSaved; sessionEdited = st.sessionEdited; view = st.view;
    wz = null;
    if (st.start) document.body.classList.add('start');
    teamNewRoleFor = null;
    closePopover();
    document.body.classList.remove('wizard-open');
    var x = $('#wzClose');
    if (x) x.remove();
    $('#wizard').hidden = true;
    $('#wizard').innerHTML = '';
    stateRev += 1;
    if (st.reloaded) {
      // a reload while the wizard was up: adopt it now (fresh undo, auto-order)
      if (selectedId && !RM.itemById(state, selectedId)) { selectedId = null; selStory = null; }
      reloadingDoc = true;
      try { adoptState(state); } finally { reloadingDoc = false; }
    } else {
      validation = RM.validate(state);
      render();
    }
    updateSaveBtn();
    if (!docSaved) scheduleAutoSave(); // an autosave the wizard held back
    if (window.HeadwayDesktop && HeadwayDesktop.syncMenu) HeadwayDesktop.syncMenu();
    return discard ? null : draft;
  }
  function wizardTryClose() {
    if (!wz) return;
    if (!wz.dirty) { closeWizard(true); return; }
    confirmBox('Discard this new project?', 'The settings you picked are not saved anywhere yet.', 'Discard',
      function () { closeWizard(true); }, true);
  }
  function wizardCreate() {
    var keep = wz && { welcomed: wz.welcomed, expandWorkWeek: wz.expandWorkWeek };
    var draft = closeWizard(false);
    if (!draft) return;
    var saved = RM.clone(draft);
    // a canceled Save dialog or a failed export must not lose the setup:
    // reopen the wizard at Review with the draft
    function reopen() {
      openWizard();
      if (!wz) return;
      state = saved;
      wz.step = 'review'; wz.dirty = true;
      wz.welcomed = keep.welcomed; wz.expandWorkWeek = keep.expandWorkWeek;
      renderWizard();
    }
    function done() { try { localStorage.setItem(ONBOARDED_KEY, '1'); } catch (e) { /* storage optional */ } }
    // flush pending edits to the currently-open file before the paths
    // switch, so the last few seconds of work can't land in the wrong file
    var flush = (window.HeadwayDesktop && !docSaved && HeadwayDesktop.currentPath())
      ? (doSave(false, true) || Promise.resolve()) : Promise.resolve();
    var go = function () { createProjectOnDisk(draft, { onDone: done, onFail: reopen }); };
    flush.then(go, go);
  }
  function wizardGo(k) {
    if (!wz) return;
    wz.step = k;
    teamNewRoleFor = null;
    renderWizard();
    var c = $('#wizard .wz-content');
    if (c) c.scrollTop = 0;
  }
  function wizardStepLabel(k) {
    if (k === 'review') return 'Review & create';
    var t = SETUP_SECTIONS[0][1].filter(function (x) { return x[0] === k; })[0];
    return t ? t[1] : k;
  }
  function welcomeHtml() {
    var cur = themePref || 'system';
    var sw = { light: ['#F5F2EC', '#FFFFFF', '#0057B8'], dark: ['#161B21', '#242C35', '#5B9BE0'],
      system: ['linear-gradient(90deg,#F5F2EC 50%,#161B21 50%)', 'rgba(255,255,255,.55)', '#0057B8'] };
    return '<section class="su-card wz-hello"><div class="wz-kicker">FIRST TIME ONLY</div><h1 class="wz-h1">Welcome to Headway</h1>' +
      '<div class="m-hint">Two things about you before the project. They’re saved on this computer, not in project files.</div>' +
      '<label class="p-lab" for="wzName">Your name</label><input id="wzName" maxlength="60" placeholder="e.g. Sam Rivera" value="' + esc(userName()) + '">' +
      '<div class="p-lab">Theme</div><div class="wz-themes">' + ['light', 'dark', 'system'].map(function (k) {
        var c = sw[k];
        return '<button class="pcard' + (cur === k ? ' on' : '') + '" data-wztheme="' + k + '" aria-pressed="' + (cur === k) + '">' +
          (cur === k ? '<span class="pc-check" aria-hidden="true">✓</span>' : '') +
          '<span class="wz-sw" style="background:' + c[0] + '"><i style="background:' + c[1] + '"></i><i style="background:' + c[2] + '"></i></span>' +
          (k === 'system' ? 'Match system' : k.charAt(0).toUpperCase() + k.slice(1)) + '</button>';
      }).join('') + '</div>' +
      '<div class="wz-hello-foot"><button class="primary" data-wz="next">Continue</button></div></section>';
  }
  // the Review step: one line per section, each with a way back
  var WZ_SUMMARY = {
    project: function () {
      var m = state.meta, p = RM.PRESETS.filter(function (x) { return x.key === m.preset; })[0];
      return (m.title || 'Untitled') + ' · ' + RM.fmtShort(RM.parseISO(m.timelineStart)) +
        (m.endDate ? ' – ' + RM.fmtShort(RM.parseISO(m.endDate)) : '') + (p ? ' · ' + p.name + ' preset' : '');
    },
    sprints: function () {
      return RM.sprintsEnabled(state.meta) ? state.meta.weeksPerSprint + '-week sprints from S' + RM.sprintInfo(state.meta).firstNum : 'No sprints';
    },
    org: function () {
      var n = state.phases.filter(function (p) { return !p.bucket; }).length;
      return n + (n === 1 ? ' phase' : ' phases') + (state.meta.workstreamsEnabled ? ' · ' + allWorkstreams().length + ' workstreams' : '');
    },
    est: function () {
      var m = state.meta;
      function nm(k) { return (RM.SIZE_SCHEMES[k] || { name: 'Custom' }).name; }
      return RM.levelLabel(state, 'feature') + ': ' + nm(m.sizeScheme) + ' · ' + RM.levelLabel(state, 'story') + ': ' + nm(m.storySizeScheme);
    },
    budget: function () {
      return RM.appEnabled(state, 'budget') ? 'On · ' + state.teamTypes.length + (state.teamTypes.length === 1 ? ' role' : ' roles') + ' on the rate card' : 'Off';
    },
    team: function () { return state.team.length + (state.team.length === 1 ? ' person' : ' people'); },
    scheduling: function () {
      var m = state.meta;
      return m.capacityEnabled ? (m.planLevel === 'story' ? 'Stories' : 'Features') + ', ' + (m.capMode === 'points' ? 'story points' : 'per person') +
        ', ' + state.capTypes.length + ' capacity types' : 'Off';
    },
    columns: function () {
      var n = allScopeCols().filter(function (c) { return !isFixedColKey(c[0]); }).length;
      return n + (n === 1 ? ' custom column' : ' custom columns');
    }
  };
  function reviewHtml() {
    return '<h1 class="wz-h1">Review &amp; create</h1><div class="m-hint">Check the essentials. Anything can be changed later in Setup.</div>' +
      '<section class="su-card wz-sum">' + WZ_STEPS.slice(0, 8).map(function (k, i) {
        return '<div class="wz-sum-row"><span class="wz-num">' + (i + 1) + '</span><b>' + esc(wizardStepLabel(k)) + '</b><span class="wz-sum-v">' +
          esc(WZ_SUMMARY[k]()) + '</span><button data-wzgo="' + k + '">Edit</button></div>';
      }).join('') + '</section>';
  }
  // required fields still empty on a step (Continue stays disabled)
  function wizardMissing(step) {
    if (step === 'project') return !String(state.meta.title || '').trim() || !state.meta.preset;
    return false;
  }
  function wizardSyncNext() {
    var nb = $('#wizard [data-wz="next"]');
    if (nb && wz && wz.step !== 'welcome') nb.disabled = wizardMissing(wz.step);
    var open = !wizardMissing('project');
    $$('#wizard .wz-step').forEach(function (b) { if (b.dataset.wzgo !== 'project') b.disabled = !open; });
  }
  // a preset switch overwrites its fields: ask first when the draft changed
  // any of them since the current preset was applied
  function presetEdited() {
    var cur = state.meta.preset;
    if (!cur) return false;
    var c = RM.clone(state);
    RM.applyPreset(c, cur);
    RM.applySizeRollup(c);
    return JSON.stringify(c) !== JSON.stringify(state);
  }
  function pickPreset(key, then) {
    if (key === state.meta.preset) { if (then) then(); return; }
    function go() { commit('preset', function (s2) { RM.applyPreset(s2, key); }); if (then) then(); }
    if (!presetEdited()) { go(); return; }
    var p = RM.PRESETS.filter(function (x) { return x.key === key; })[0];
    confirmBox('Switch to ' + p.name + '?',
      'The preset resets sizing, priority, risk, sprints, budget, scheduling and item types. The changes you made to those will be lost.',
      'Switch preset', go, true);
  }
  function renderWizard() {
    if (!wz) return;
    var host = $('#wizard .wz-body');
    if (!host) return;
    var keep = host.querySelector('.wz-content');
    var scroll = keep ? keep.scrollTop : 0;
    var idx = WZ_STEPS.indexOf(wz.step);
    var hasPreset = !wizardMissing('project');
    var need = wizardMissing(wz.step);
    var seq = (wz.hasWelcome ? ['welcome'] : []).concat(WZ_STEPS);
    var pct = Math.round((seq.indexOf(wz.step) + 1) / seq.length * 100);
    var bar = $('#wzProgress');
    if (bar) { bar.setAttribute('aria-valuenow', pct); bar.firstChild.style.width = pct + '%'; }
    if (wz.step === 'welcome') {
      host.innerHTML = '<div class="wz-welcome">' + welcomeHtml() + '</div>';
    } else {
      var rail = WZ_STEPS.map(function (k, i) {
        var done = i < idx;
        return '<button class="wz-step' + (k === wz.step ? ' on' : '') + (done ? ' done' : '') + '" data-wzgo="' + k + '"' +
          (!hasPreset && k !== 'project' ? ' disabled' : '') + '><span class="wz-num">' + (done ? '✓' : i + 1) + '</span>' +
          esc(wizardStepLabel(k)) + '</button>';
      }).join('');
      var body = wz.step === 'review' ? reviewHtml() : sectionBody(wz.step, 'wizard');
      host.innerHTML =
        '<div class="wz-layout"><nav class="wz-rail" aria-label="New project steps">' + rail + '</nav>' +
        '<div class="wz-main"><div class="wz-content">' +
        (wz.step === 'review' ? '' : '<h1 class="wz-h1">' + esc(wizardStepLabel(wz.step)) + '</h1>') + body + '</div>' +
        '<div class="wz-foot"><span></span><button data-wz="back"' + (idx === 0 && !wz.welcomed ? ' disabled' : '') + '>Back</button>' +
        (idx > 0 && wz.step !== 'review' ? '<button data-wz="skip">Skip</button>' : '') +
        '<button class="primary" data-wz="next"' + (need ? ' disabled' : '') + '>' + (wz.step === 'review' ? 'Create project' : 'Continue') + '</button></div>' +
        '</div></div>';
      var c = host.querySelector('.wz-content');
      if (c && keep) c.scrollTop = scroll;
    }
    var nrIn = teamNewRoleFor && host.querySelector('[data-sutmnewrole]');
    if (nrIn) nrIn.focus();
    if (window.lucide) lucide.createIcons();
    // the top bar settles once its icons render: keep the wizard (and its
    // progress bar) right below it
    var tbH = $('#topbar').offsetHeight;
    if (tbH) document.documentElement.style.setProperty('--topbar-h', tbH + 'px');
  }
  $('#wizard').addEventListener('click', function (e) {
    if (!wz) return;
    var go = e.target.closest('[data-wzgo]');
    if (go) { if (!go.disabled) wizardGo(go.dataset.wzgo); return; }
    var pc = e.target.closest('[data-wzpreset]');
    if (pc) {
      pickPreset(pc.dataset.wzpreset);
      return;
    }
    var th = e.target.closest('[data-wztheme]');
    if (th) { setTheme(th.dataset.wztheme); renderWizard(); return; }
    if (e.target.closest('[data-wzexpand]')) { wz.expandWorkWeek = true; renderWizard(); return; }
    var b = e.target.closest('[data-wz]');
    if (!b || b.disabled) return;
    var idx = WZ_STEPS.indexOf(wz.step);
    if (b.dataset.wz === 'next' && wz.step === 'welcome') { wz.welcomed = true; wizardGo('project'); return; }
    if (b.dataset.wz === 'next' && wz.step === 'review') { wizardCreate(); return; }
    if (b.dataset.wz === 'next' || b.dataset.wz === 'skip') { wizardGo(WZ_STEPS[Math.min(idx + 1, WZ_STEPS.length - 1)]); return; }
    if (b.dataset.wz === 'back') {
      if (idx > 0) wizardGo(WZ_STEPS[idx - 1]);
      else if (wz.welcomed) wizardGo('welcome');
    }
  });
  $('#wizard').addEventListener('change', function (e) {
    if (e.target.id === 'wzName') setUserName(e.target.value);
  });
  // required fields: Continue follows the typing, not just the commit
  $('#wizard').addEventListener('input', function (e) {
    if (!wz || e.target.id !== 'suTitle') return;
    state.meta.title = e.target.value.trim();
    wz.dirty = true;
    wizardSyncNext();
  });
  // the preset radio group: arrow keys pick the next / previous preset
  $('#wizard').addEventListener('keydown', function (e) {
    var card = e.target.closest && e.target.closest('[data-wzpreset]');
    if (!wz || !card) return;
    var d = e.key === 'ArrowRight' || e.key === 'ArrowDown' ? 1 : e.key === 'ArrowLeft' || e.key === 'ArrowUp' ? -1 : 0;
    if (!d) return;
    e.preventDefault();
    var keys = RM.PRESETS.map(function (p) { return p.key; });
    var next = keys[(keys.indexOf(card.dataset.wzpreset) + d + keys.length) % keys.length];
    pickPreset(next, function () {
      var nb = $('#wizard [data-wzpreset="' + next + '"]');
      if (nb) nb.focus();
    });
  });
  $('#topbar').addEventListener('click', function (e) {
    if (e.target.closest('#wzClose')) wizardTryClose();
  });

  // the Setup section handlers, shared by #setupView and the wizard
  function bindSectionHandlers(suHost) {
  // vertical settings rail; personal controls share the modal's wiring
  suHost.addEventListener('click', function (e) {
    var tab = e.target.closest('[data-sutab]');
    if (!tab) return;
    setupTab = tab.dataset.sutab;
    saveLocal();
    renderSetup();
  });
  wirePersonalFields(suHost, function () {
    if (wz) renderWizard();
    else if (view === 'setup') renderSetup();
  });

  suHost.addEventListener('change', function (e) {
    if (e.target.matches('[data-subudget]')) {
      var budOn = e.target.checked;
      commit('budgeting', function (s2) { s2.meta.apps.budget = budOn; });
      return;
    }
    var tf = e.target.closest('[data-sutm]');
    if (tf) {
      var tmId = tf.dataset.mid, tmF = tf.dataset.sutm, tmV = tf.value;
      if (tmF === 'type' && tmV === '__new') { teamNewRoleFor = tmId; renderSetupHost(); return; }
      commit('team', function (s2) {
        var mm = s2.team.filter(function (x) { return x.id === tmId; })[0];
        if (!mm) return;
        if (tmF === 'capacity') mm.capacity = Math.max(0, (+tmV || 0) / 100);
        else if (tmF === 'points') mm.points = tmV === '' ? null : Math.max(0, +tmV || 0);
        else if (tmF === 'rate' || tmF === 'cost') mm[tmF] = Math.max(0, +tmV || 0);
        else mm[tmF] = String(tmV);
        if (tmF === 'type') RM.syncMemberCapTypes(s2);
      });
      return;
    }
    if (e.target.matches('[data-sutmnewrole]')) {
      var nrId = e.target.dataset.mid, nrName = e.target.value.trim();
      teamNewRoleFor = null;
      if (!nrName) { renderSetupHost(); return; }
      commit('add role', function (s2) {
        if (s2.teamTypes.indexOf(nrName) === -1) s2.teamTypes.push(nrName);
        s2.team.forEach(function (x) { if (x.id === nrId) x.type = nrName; });
        RM.syncMemberCapTypes(s2);
      });
      return;
    }
    var sk = e.target.closest('[data-suscheme-kind]');
    if (sk) {
      var skKind = sk.dataset.kind, skV = sk.value;
      if (sk.dataset.suschemeKind === 'size') {
        if (skV === state.meta[RM.sizeKeys(skKind).scheme]) return;
        commit('sizing approach', function (s2) { RM.setSizeScheme(s2, skV, skKind); });
      } else if (sk.dataset.suschemeKind === 'prio') {
        if (skV === RM.prioritySchemeOf(state, skKind)) return;
        commit('priority column', function (s2) { RM.setPriorityScheme(s2, skV, skKind); });
      } else {
        if (skV === RM.riskSchemeOf(state, skKind)) return;
        commit('risk column', function (s2) { RM.setRiskScheme(s2, skV, skKind); });
      }
      return;
    }
    if (e.target.id === 'suTitle') {
      var tv2 = e.target.value.trim() || (wz ? '' : 'Roadmap');
      commit('title', function (s2) { s2.meta.title = tv2; });
      return;
    }
    var t = e.target;
    if (t.id === 'suCapEnable') {
      var on = t.checked;
      commit('capacity feature', function (s2) { s2.meta.capacityEnabled = on; });
      toast('Scheduling ' + (on ? 'on' : 'off'));
      return;
    }
    if (t.id === 'suWsEnable') {
      var wsOn = t.checked;
      commit('workstream feature', function (s2) { s2.meta.workstreamsEnabled = wsOn; });
      toast('Workstreams ' + (wsOn ? 'enabled' : 'disabled'));
      return;
    }
    if (t.dataset.suhlabel) { var hk2 = t.dataset.suhlabel, hv = t.value; commit('level label', function (s2) { RM.setLevelLabel(s2, hk2, hv); }); return; }
    if (t.dataset.suhtlabel) { var tk2 = t.dataset.suhtlabel, tlv2 = t.value; commit('rename type', function (s2) { RM.renameItemType(s2, tk2, tlv2); }); return; }
    if (t.dataset.suhtcolor) { var tk4 = t.dataset.suhtcolor, cv = t.value; commit('type color', function (s2) { RM.setItemTypeColor(s2, tk4, cv); }); return; }
    if (t.dataset.suhtjira) { var tk3 = t.dataset.suhtjira, jv = t.value; commit('type jira name', function (s2) { RM.setItemTypeJira(s2, tk3, jv); }); return; }
    if (t.dataset.suhtlevel) {
      var tk5 = t.dataset.suhtlevel, lvTo = t.value, okL = true;
      commit('type level', function (s2) { okL = RM.setTypeLevel(s2, tk5, lvTo); });
      if (!okL) toast('A level needs at least one type');
      return;
    }
    if (t.dataset.suapp) {
      var appKey = t.dataset.suapp, appOn = t.checked;
      var appName = (RM.APPS.filter(function (a) { return a[0] === appKey; })[0] || [])[1] || appKey;
      commit('app ' + (appOn ? 'on' : 'off'), function (s2) { s2.meta.apps[appKey] = appOn; });
      toast(appName + (appOn ? ' enabled' : ' hidden'));
      return;
    }
    if (t.dataset.suholname != null || t.dataset.suholstart != null || t.dataset.suholend != null) {
      var hIdx = parseInt(t.dataset.suholname != null ? t.dataset.suholname
        : t.dataset.suholstart != null ? t.dataset.suholstart : t.dataset.suholend, 10);
      var hPatch = t.dataset.suholname != null ? { name: t.value }
        : t.dataset.suholstart != null ? { start: t.value } : { end: t.value };
      if ((hPatch.start != null && !hPatch.start) || (hPatch.end != null && !hPatch.end)) { render(); return; }
      commit('holiday edit', function (s2) { RM.updateHolidayRange(s2.meta, hIdx, hPatch); });
      return;
    }
    if (t.id === 'suWeekHours') {
      var wh2 = Math.max(1, Math.min(80, parseFloat(t.value) || RM.WEEK_HOURS));
      commit('week hours', function (s2) { s2.meta.weekHours = wh2; });
      return;
    }
    if (t.id === 'suWeekStart') {
      var wsd = parseInt(t.value, 10);
      commit('week start', function (s2) {
        RM.changeWorkWeek(s2, { weekStart: wsd });
      });
      return;
    }
    if (t.dataset.suwday != null) {
      var dayN = parseInt(t.dataset.suwday, 10);
      var list = RM.workDaysOf(state.meta).slice();
      if (t.checked) { if (list.indexOf(dayN) === -1) list.push(dayN); }
      else list = list.filter(function (x) { return x !== dayN; });
      if (!list.length) { render(); return; }
      commit('work days', function (s2) {
        RM.changeWorkWeek(s2, { workDays: list });
      });
      return;
    }
    if (t.dataset.sucolscope != null) {
      var scKey = t.dataset.sucolscope, scVal = t.value;
      commit('column scope', function (s2) { RM.setScopeColScope(s2, scKey, scVal); });
      return;
    }
    if (t.dataset.sucolname != null) {
      var cKey = t.dataset.sucolname;
      var cLbl = t.value.trim();
      commit('rename column', function (s2) { RM.renameScopeCol(s2, cKey, cLbl); });
      return;
    }
    if (t.dataset.capname != null) {
      var oldCap = t.dataset.capname;
      var newCap = t.value.trim();
      if (!newCap || newCap === oldCap) { render(); return; }
      if (state.capTypes.indexOf(newCap) !== -1) {
        toast('\u201C' + newCap + '\u201D already exists', 'err');
        render();
        return;
      }
      commit('rename capacity type', function (s2) { RM.renameCapType(s2, oldCap, newCap); });
      return;
    }
    if (t.id === 'suDefPoints') {
      var dp = parseFloat(t.value);
      if (!isFinite(dp) || dp < 0) { render(); return; }
      commit('default points', function (s2) { s2.meta.defaultPoints = dp; });
      return;
    }
    if (t.dataset.rcname != null) {
      var oldRole = t.dataset.rcname;
      var newRole = t.value.trim();
      if (!newRole || newRole === oldRole) { render(); return; }
      if (state.teamTypes.indexOf(newRole) !== -1) {
        toast('\u201C' + newRole + '\u201D already exists', 'err');
        render();
        return;
      }
      commit('rename role', function (s2) { RM.renameRole(s2, oldRole, newRole); });
      return;
    }
    if (t.dataset.rcrate != null || t.dataset.rccost != null) {
      var rcRole = t.dataset.rcrate != null ? t.dataset.rcrate : t.dataset.rccost;
      var rcField = t.dataset.rcrate != null ? 'rate' : 'cost';
      var rcVal = Math.max(0, parseFloat(t.value) || 0);
      commit('rate card', function (s2) {
        var card = s2.meta.rateCard[rcRole] || { rate: 0, cost: 0 };
        card[rcField] = rcVal;
        if (card.rate || card.cost) s2.meta.rateCard[rcRole] = card;
        else delete s2.meta.rateCard[rcRole];
      });
      return;
    }
    if (t.dataset.suszlabel != null) {
      var oldLbl = t.dataset.suszlabel, lblKind = t.dataset.kind || 'feature';
      var newLbl = t.value.trim();
      if (!newLbl || newLbl === oldLbl) { render(); return; }
      if (RM.sizeOrderOf(state, lblKind).indexOf(newLbl) !== -1) {
        toast('“' + newLbl + '” already exists', 'err');
        render();
        return;
      }
      commit('rename size', function (s2) { RM.renameSizeOption(s2, oldLbl, newLbl, lblKind); });
      return;
    }
    if (t.id === 'suStart') {
      var d = RM.parseISO(t.value || state.meta.timelineStart);
      if (!isFinite(d.getTime())) { render(); return; }
      d.setUTCDate(d.getUTCDate() - ((d.getUTCDay() + 6) % 7)); // snap Monday
      commit('timeline start', function (s2) {
        s2.meta.timelineStart = RM.fmtISO(d);
        RM.syncEndDate(s2.meta);
      });
      return;
    }
    if (t.id === 'suEnd') {
      var ed = RM.parseISO(t.value);
      if (!isFinite(ed.getTime())) { render(); return; }
      commit('timeline end', function (s2) {
        var w = Math.floor((ed - RM.parseISO(s2.meta.timelineStart)) / (7 * 86400000)) + 1;
        if (isFinite(w)) s2.meta.numWeeks = Math.max(4, w);
        RM.syncEndDate(s2.meta);
      });
      return;
    }
    if (t.id === 'suAnchor') {
      var ad = RM.parseISO(t.value || state.meta.sprintAnchor);
      if (!isFinite(ad.getTime())) { render(); return; }
      ad.setUTCDate(ad.getUTCDate() - ((ad.getUTCDay() + 6) % 7));
      commit('sprint anchor', function (s2) { s2.meta.sprintAnchor = RM.fmtISO(ad); });
      return;
    }
    if (t.id === 'suAnchorNum') {
      var n = parseInt(t.value, 10);
      if (!isFinite(n)) { render(); return; }
      commit('sprint number', function (s2) { s2.meta.sprintAnchorNum = n; });
      return;
    }
    if (t.dataset.susz) {
      var sz = t.dataset.susz, szKeys = RM.sizeKeys(t.dataset.kind || 'feature');
      var cur = (state.meta[szKeys.days] || {})[sz];
      var pv = parseFloat(t.value);
      // 0 is a real day value (a 0-point story); anything below it is not
      var v = isFinite(pv) && pv >= 0 ? pv : (cur != null ? cur : 5);
      commit('size days', function (s2) {
        s2.meta[szKeys.days][sz] = v;
        RM.markSizeCustom(s2, t.dataset.kind || 'feature');
      });
      return;
    }
  });

  suHost.addEventListener('click', function (e) {
    var t = e.target.closest('button');
    if (!t) return;
    if (t.id === 'suHolAddBtn') {
      var hName = $('#suHolName').value.trim();
      var hStart = $('#suHolStart').value;
      var hEnd = $('#suHolEnd').value || hStart;
      if (!hStart) return;
      if (hEnd < hStart) { var swp = hStart; hStart = hEnd; hEnd = swp; }
      commit('holiday add', function (s2) {
        RM.addHolidayRange(s2.meta, hName, hStart, hEnd);
      });
      return;
    }
    // workstreams: the same multi-pick menu as the Budgeting / Resources rows
    if (t.dataset.sutmws != null) { openMemberWsDropdown(t, t.dataset.mid); return; }
    if (t.dataset.sutmdel != null) {
      var tdId = t.dataset.sutmdel;
      commit('remove person', function (s2) { s2.team = s2.team.filter(function (x) { return x.id !== tdId; }); });
      return;
    }
    if (t.id === 'suTmAdd') {
      commit('add person', function (s2) {
        s2.team.push({ id: RM.uid('t'), name: '', role: '', type: '', capType: '', workstream: '', workstreams: [],
          capacity: 1, points: null, rate: 0, cost: 0, weekHours: {} });
      });
      return;
    }
    var holrm = t.dataset.suholrm;
    if (holrm != null) {
      var holIdx = parseInt(holrm, 10);
      commit('holiday rm', function (s2) { RM.removeHolidayRange(s2.meta, holIdx); });
      return;
    }
    if (t.id === 'suSzAdd' || t.id === 'suSzAddStory') {
      var addKind = t.id === 'suSzAddStory' ? 'story' : 'feature';
      commit('add size', function (s2) {
        var lbl = 'New', n2 = 2;
        while (s2.meta[RM.sizeKeys(addKind).order].indexOf(lbl) !== -1) lbl = 'New ' + n2++;
        RM.addSizeOption(s2, lbl, 5, addKind);
      });
      return;
    }
    if (t.dataset.suszrm) {
      var rmLbl = t.dataset.suszrm, rmKind = t.dataset.kind || 'feature';
      commit('remove size', function (s2) { RM.removeSizeOption(s2, rmLbl, rmKind); });
      return;
    }
    if (t.id === 'suWsAddBtn') {
      var wsNames = addListNames($('#suWsAdd').value);
      if (!wsNames.length) return;
      commit('add workstream', function (s2) {
        wsNames.forEach(function (wv) {
          if (!s2.wsColors[wv]) s2.wsColors[wv] = RM.DEFAULT_WS_COLORS[wv] || 'neutral';
        });
      });
      return;
    }
    if (t.id === 'suPhAddBtn') {
      var phNames = addListNames($('#suPhAddIn').value, state.phases.map(function (p2) { return p2.name; }));
      if (!phNames.length) return;
      commit('add phase', function (s2) {
        phNames.forEach(function (nm) {
          s2.phases.push({ id: RM.uid('p'), name: nm, description: '', bucket: false, collapsed: false });
        });
      });
      return;
    }
    if (t.id === 'suEpAddBtn') {
      var epNames = addListNames($('#suEpAdd').value, allEpics());
      if (!epNames.length) return;
      commit('add epic', function (s2) {
        s2.epicList = (s2.epicList || []).concat(epNames);
      });
      return;
    }
    if (t.dataset.suhieropen != null) { orgHierOpen = true; renderSetupHost(); return; }
    if (t.dataset.suszreset) {
      var rsKind = t.dataset.kind || 'feature', rsScheme = t.dataset.suszreset;
      commit('sizing approach', function (s2) { RM.setSizeScheme(s2, rsScheme, rsKind); });
      return;
    }
    if (t.dataset.suroleadd) { roleSupplyMenu(t, t.dataset.suroleadd); return; }
    if (t.dataset.surolerm) {
      var rmRole = t.dataset.surolerm;
      commit('role capacity type', function (s2) { RM.setRoleCapType(s2, rmRole, ''); });
      return;
    }
    if (t.id === 'suColAddBtn') {
      var colV = $('#suColAdd').value.trim();
      if (!colV) return;
      commit('add column', function (s2) { RM.addScopeCol(s2, colV); });
      return;
    }
    if (t.dataset.sucoldel) {
      if (t.disabled || isFixedColKey(t.dataset.sucoldel)) return; // built-ins stay
      var rmKey = t.dataset.sucoldel;
      commit('remove column', function (s2) { RM.removeScopeCol(s2, rmKey); });
      return;
    }
    if (t.dataset.suwps != null) {
      var wpsV = parseInt(t.dataset.suwps, 10);
      commit('sprint length', function (s2) { s2.meta.weeksPerSprint = wpsV; });
      return;
    }
    if (t.dataset.suwsedit) { wsEditModal(t.dataset.suwsedit); return; }
    if (t.dataset.sudefws != null) { defaultWsModal(); return; }
    if (t.dataset.suhtrm) {
      var rmTypeKey = t.dataset.suhtrm, okR = true;
      commit('remove type', function (s2) { okR = RM.removeItemType(s2, rmTypeKey); });
      if (!okR) toast('A level needs at least one type');
      return;
    }
    if (t.dataset.suhticon) { typeIconMenu(t, t.dataset.suhticon); return; }
    if (t.dataset.suhadd) {
      var addLv = t.dataset.suhadd, newKey = '';
      commit('add type', function (s2) { newKey = RM.addItemType(s2, 'New type', 'tag', '', addLv); });
      requestAnimationFrame(function () { var inp = $('input[data-suhtlabel="' + newKey + '"]', suHost); if (inp) { inp.focus(); inp.select(); } });
      return;
    }
    if (t.dataset.suepedit) { epicEditModal(t.dataset.suepedit); return; }
    if (t.dataset.suepdel) { deleteEpicConfirm(t.dataset.suepdel); return; }
    if (t.dataset.suphedit) { phaseModal(t.dataset.suphedit); return; }
    if (t.dataset.suphdel) { deletePhaseConfirm(t.dataset.suphdel); return; }
    if (t.dataset.suplan) {
      var planV = t.dataset.suplan;
      if (planV === RM.planLevel(state)) return;
      commit('planning level', function (s2) { s2.meta.planLevel = planV; });
      return;
    }
    if (t.dataset.sucapmode) {
      var cmV = t.dataset.sucapmode;
      commit('capacity mode', function (s2) { s2.meta.capMode = cmV; });
      toast('Demand: ' + (cmV === 'points' ? 'story points' : 'per person'));
      return;
    }
    if (t.id === 'suCapTypeAddBtn') {
      var cv = $('#suCapTypeAdd').value.trim();
      if (!cv) return;
      if (state.capTypes.indexOf(cv) !== -1) { toast('Capacity type already exists'); return; }
      commit('add capacity type', function (s2) { s2.capTypes.push(cv); });
      return;
    }
    if (t.dataset.sucaprm) {
      var caprm = t.dataset.sucaprm;
      commit('remove capacity type', function (s2) { RM.removeCapType(s2, caprm); });
      return;
    }
    if (t.id === 'suTypeAddBtn') {
      var tv = $('#suTypeAdd').value.trim();
      if (!tv) return;
      if (state.teamTypes.indexOf(tv) !== -1) { toast('Role already exists'); return; }
      commit('add type', function (s2) { s2.teamTypes.push(tv); });
      return;
    }
    var ttrm = t.dataset.suttrm;
    if (ttrm) {
      // people and features that had the role simply lose it
      var usedP = state.team.filter(function (mm) { return mm.type === ttrm; }).length;
      var usedF = state.items.filter(function (x) { return x.teamType === ttrm; }).length;
      commit('remove type', function (s2) { RM.removeRole(s2, ttrm); });
      if (usedP || usedF) {
        var parts = [];
        if (usedP) parts.push(usedP + (usedP === 1 ? ' person' : ' people'));
        if (usedF) parts.push(usedF + (usedF === 1 ? ' feature' : ' features'));
        toast('Removed “' + ttrm + '” — ' + parts.join(' and ') + ' now have no role');
      }
      return;
    }
  });

  // drag-to-reorder for the Setup lists (workstreams / phases / team types)
  function moveKeyBefore(keys, key, beforeKey) {
    var out = keys.filter(function (k) { return k !== key; });
    var i = beforeKey == null ? out.length : out.indexOf(beforeKey);
    if (i < 0) i = out.length;
    out.splice(i, 0, key);
    return out;
  }
  (function () {
    var d = null;
    suHost.addEventListener('pointerdown', function (e) {
      var g = e.target.closest('.su-grip');
      if (!g) return;
      var row = g.closest('.su-row');
      var list = row && row.parentNode;
      if (!row || !list || !list.dataset.sulist) return;
      d = { kind: list.dataset.sulist, key: row.dataset.key, list: list, row: row, before: undefined, indicator: null };
      row.classList.add('drag-row-ghost');
      e.preventDefault();
    });
    window.addEventListener('pointermove', function (e) {
      if (!d) return;
      if (!d.indicator) {
        d.indicator = document.createElement('div');
        d.indicator.className = 'su-indicator';
        d.list.appendChild(d.indicator);
      }
      var rows = $$('.su-row', d.list);
      d.before = null;
      var y = null;
      var lr = d.list.getBoundingClientRect();
      for (var i = 0; i < rows.length; i++) {
        var r = rows[i].getBoundingClientRect();
        if (e.clientY < r.top + r.height / 2) { d.before = rows[i].dataset.key; y = r.top - lr.top; break; }
      }
      if (y == null && rows.length) y = rows[rows.length - 1].getBoundingClientRect().bottom - lr.top;
      d.indicator.style.top = (y || 0) + 'px';
    });
    window.addEventListener('pointerup', function () {
      if (!d) return;
      var dd = d;
      d = null;
      dd.row.classList.remove('drag-row-ghost');
      if (dd.indicator) dd.indicator.remove();
      if (dd.before === undefined || dd.before === dd.key) return; // never moved / no-op
      commit('reorder ' + dd.kind, function (s2) {
        if (dd.kind === 'type') {
          s2.teamTypes = moveKeyBefore(s2.teamTypes, dd.key, dd.before);
        } else if (dd.kind === 'captype') {
          s2.capTypes = moveKeyBefore(s2.capTypes, dd.key, dd.before);
        } else if (dd.kind === 'ws') {
          s2.wsOrder = moveKeyBefore(allWorkstreams(s2), dd.key, dd.before);
        } else if (dd.kind === 'scol') {
          s2.meta.scopeColOrder = moveKeyBefore(s2.meta.scopeColOrder, dd.key, dd.before);
        } else if (dd.kind === 'phase') {
          var ids = s2.phases.map(function (p2) { return p2.id; });
          var order = moveKeyBefore(ids, dd.key, dd.before);
          s2.phases.sort(function (a, b) { return order.indexOf(a.id) - order.indexOf(b.id); });
        }
      });
    });
  })();
  }
  bindSectionHandlers($('#setupView'));
  bindSectionHandlers($('#wizard'));

  // template workbook: a blank roadmap carrying ONE worked example so every
  // sheet's shape is visible (feature with size/deps columns + a story)
  function templateState() {
    var s = blankState();
    s.meta.title = 'Roadmap Template';
    s.items = [{
      id: RM.uid('i'), num: 1, phaseId: 'now',
      feature: 'Example feature — replace me',
      description: 'One example row. Add features below it; sizes are XS 2d · S 1w · M 2w · L 4w · XL 8w.',
      workstream: 'Example workstream', epic: 'Example epic',
      size: 'M', headcount: 1,
      startDay: 0, durDays: 10,
      stories: [{ id: RM.uid('s'), title: 'Example story', startDay: 0, durDays: 5 }]
    }];
    return RM.normalizeState(s);
  }

  function downloadTemplate() {
    RMExcel.exportWorkbook(templateState()).then(function (blob) {
      var a = document.createElement('a');
      a.href = URL.createObjectURL(blob);
      a.download = 'roadmap-template.xlsx';
      a.click();
      setTimeout(function () { URL.revokeObjectURL(a.href); }, 4000);
      toast('Saved roadmap-template.xlsx');
    }).catch(function (err) {
      toast('Template export failed: ' + err.message, 'err');
    });
  }

  function blankState() {
    var now = new Date();
    var mon = new Date(Date.UTC(now.getFullYear(), now.getMonth(), now.getDate()));
    mon.setUTCDate(mon.getUTCDate() - ((mon.getUTCDay() + 6) % 7));
    return RM.normalizeState({
      meta: { title: 'New Roadmap', timelineStart: RM.fmtISO(mon), numWeeks: 32, holidays: [] },
      phases: [
        { id: 'now', name: 'Now', bucket: false },
        { id: 'next', name: 'Next', bucket: true },
        { id: 'future', name: 'Future', bucket: true }
      ],
      items: [], team: []
    });
  }

  function helpModal() {
    openModal(
      '<div class="modal">' +
      '<div class="m-head"><h2>Headway</h2><button class="p-close" data-m="x"><i data-lucide="x"></i></button></div>' +
      '<div class="m-body" style="font-size:12.5px;line-height:1.75;color:var(--ink-2)">' +
      '<b>Views</b> — Scoping (spreadsheet, resizable columns) and Planning (timeline), tabs up top<br>' +
      '<b>Bars</b> — drag to move · edges resize · <span class="kbd">⌘drag</span> pushes dependents along · <span class="kbd">←</span><span class="kbd">→</span> nudge (<span class="kbd">⇧</span> = 1 week) · snap grids (features / stories) in View<br>' +
      '<b>Sizes in weeks</b> — XS 2d · S 1w · M 2w · L 4w · XL 8w; the risk buffer uses the same scale (panel), shown as one duration<br>' +
      '<b>Rows</b> — drag the left pane to reorder / move phase (auto-order can re-sort by start) · right-click for insert/delete · chips cycle on click<br>' +
      '<b>Links</b> — drag a bar’s edge circles onto another row (left = depends ON it, right = dependency FOR it; Esc cancels) · click an arrow, Delete removes · orange = critical path<br>' +
      '<b>Timeline</b> — drag empty space to pan · <span class="kbd">⌘+</span> / <span class="kbd">⌘−</span> or <span class="kbd">⌘scroll</span> zoom · click a capacity cell to toggle a holiday week; single dates + sprint numbering (e.g. S1 = Sep 7) in Settings<br>' +
      '<b>Availability row</b> — people available per week (fractional when hours dip); flags weeks with too much concurrent work<br>' +
      '<b>Resources panel</b> — bottom, resizable/collapsible; hours per person per week (default = the project full-time week) — click a cell to type, drag to fill; drag the grip to reorder people<br>' +
      '<b>Auto</b> — dependency-ordered, capacity-aware schedule; locked items and those excluded from Auto stay put<br>' +
      '<span class="kbd">⌘Z</span> undo · <span class="kbd">⇧⌘Z</span> redo · <span class="kbd">⌘S</span> save · <span class="kbd">Del</span> delete · <span class="kbd">Esc</span> close' +
      '</div>' +
      '<div class="m-foot"><button data-m="x2" class="primary">Got it</button></div></div>',
      function (host) {
        $('[data-m=x]', host).onclick = closeModal;
        $('[data-m=x2]', host).onclick = closeModal;
      });
  }

  // ------------------------------------------------------------ resources panel
  // Bottom panel: one row per person, hours per week (default 40), edited
  // like a spreadsheet — click a cell to type, drag across cells to fill.
  var resBody = $('#resBody'), resGrid = $('#resGrid');

  function fmtH(h) { return h % 1 === 0 ? String(h) : String(Math.round(h * 10) / 10); }

  // blend hex a → hex b (no '#') by t in [0,1]
  function mixHex(a, b, t) {
    function c(h, i) { return parseInt(h.slice(i, i + 2), 16); }
    function h2(v) { var s = Math.round(v).toString(16).toUpperCase(); return s.length === 1 ? '0' + s : s; }
    return h2(c(a, 0) + (c(b, 0) - c(a, 0)) * t) +
      h2(c(a, 2) + (c(b, 2) - c(a, 2)) * t) +
      h2(c(a, 4) + (c(b, 4) - c(a, 4)) * t);
  }

  function renderResources() {
    $('#resCount').textContent = state.team.length;
    $('#resPanel').classList.toggle('collapsed', resCollapsed);
    $('#resChev').classList.toggle('open', !resCollapsed);
    resBody.style.height = resCollapsed ? '0px' : resPanelH + 'px';
    if (resCollapsed) { resGrid.innerHTML = ''; return; }
    var meta = state.meta;
    var laneW = meta.numWeeks * weekPx;
    var html = [];
    state.team.forEach(function (m) {
      var cells = [];
      for (var w = 0; w < meta.numWeeks; w++) {
        var h = RM.memberHoursForWeek(meta, m, w);
        var cls = h > RM.weekHoursOf(meta) ? ' rh-over' : '';
        // 0h → full-week ramps toward blue from the theme's ground: white in
        // light (dark label), surface in dark (light label) — see --cell-ink
        var t2 = Math.max(0, Math.min(1, h / RM.weekHoursOf(meta)));
        var bg = '#' + (darkActive() ? mixHex('242C35', '31597F', t2) : RM.tint('7FAEDD', 1 - t2));
        cells.push('<div class="rh' + cls + '" data-w="' + w + '" style="left:' + (w * weekPx) +
          'px;width:' + weekPx + 'px;background:' + bg + '" title="' + esc(mLabel(m)) + '">' +
          (weekPx >= 20 ? fmtH(h) : '') + '</div>');
      }
      // SAME left columns as the Budgeting view (name · role · rate card ·
      // workstream) — budgeting alone adds cost/rate/margin/total after them
      html.push(
        '<div class="rrow" data-mid="' + m.id + '">' +
        '<div class="rleft">' +
        '<span class="r-grip rr-grip"><i data-lucide="grip-vertical"></i></span>' +
        memberColsHtml(m) +
        // one or the other: the × seat multiplier in per-person mode, points
        // per sprint in story-points mode — never both
        // untyped people supply nothing: a quiet prompt to give them a type
        // instead (their seat and points are kept for when they get one)
        (state.meta.capacityEnabled && !m.capType
          ? '<span class="res-cap res-untyped" tabindex="0" role="button" data-runtyped="' + m.id +
            '" title="No capacity type — supplies nothing. Click to pick a role">set type</span>'
          : '') +
        (state.meta.capacityEnabled && m.capType && state.meta.capMode !== 'points'
          ? '<span class="res-cap" tabindex="0" role="button" data-rcap="' + m.id +
            '" title="Capacity at full-time hours">' + fmtPe(m.capacity != null ? m.capacity : 1) + '×</span>'
          : '') +
        (state.meta.capacityEnabled && m.capType && state.meta.capMode === 'points'
          ? '<span class="res-cap res-pts' + (m.points == null ? ' dflt' : '') + '" tabindex="0" role="button" data-rpts="' + m.id +
            '" title="Story points per sprint (blank = document default)">' + fmtPe(RM.memberPoints(state, m)) + ' pt</span>'
          : '') +
        '</div>' +
        '<div class="rlane" style="width:' + laneW + 'px">' + cells.join('') + '</div>' +
        '</div>');
    });
    // blank click-to-add row (types a role in place; defaults to 40 h/week)
    html.push('<div class="rrow raddrow" data-resadd title="Add a role">' +
      '<div class="rleft"><span class="addrow-lab"><i data-lucide="plus"></i> New role</span></div>' +
      '<div class="rlane" style="width:' + laneW + 'px"></div></div>');
    resGrid.innerHTML = html.join('');
    resGrid.style.width = 'calc(var(--left-w) + ' + laneW + 'px)';
    if (window.lucide) lucide.createIcons();
    resBody.scrollLeft = board.scrollLeft;
  }

  $('#hdrCapRows').addEventListener('keydown', function (e) {
    if ((e.key === 'Enter' || e.key === ' ') && e.target.closest('[data-w]')) {
      e.preventDefault();
      e.target.dispatchEvent(new MouseEvent('click', { bubbles: true }));
    }
  });

  // the whole header bar toggles the panel (not just the chevron)
  $('#resHead').addEventListener('click', function (e) {
    if (e.target.closest('#resManage')) return;
    resCollapsed = !resCollapsed;
    saveLocal();
    renderResources();
  });
  $('#resManage').addEventListener('click', function () {
    view = 'setup';
    setupTab = 'team';
    saveLocal();
    render();
  });

  // right panel width — drag its left edge
  (function () {
    var prs = null;
    $('#panel').addEventListener('pointerdown', function (e) {
      if (e.target.id !== 'panelRz') return;
      prs = { x0: e.clientX, w0: panelW };
      e.preventDefault();
    });
    window.addEventListener('pointermove', function (e) {
      if (!prs) return;
      panelW = Math.round(Math.max(280, Math.min(window.innerWidth * 0.6, prs.w0 - (e.clientX - prs.x0))));
      document.documentElement.style.setProperty('--panel-w', panelW + 'px');
    });
    window.addEventListener('pointerup', function () {
      if (prs) { prs = null; saveLocal(); }
    });
  })();

  // frozen left-pane width — resizable per view (independent of the drag machine)
  (function () {
    var lrs = null;
    function startLeftRz(e) {
      lrs = { x0: e.clientX, w0: view === 'scoping' ? leftWScope : view === 'budget' ? leftWBudget : planLeftW() };
      e.preventDefault();
      e.stopPropagation();
    }
    $('#leftRz').addEventListener('pointerdown', startLeftRz);
    $('#leftRzLine').addEventListener('pointerdown', startLeftRz);
    window.addEventListener('pointermove', function (e) {
      if (!lrs) return;
      var w = Math.round(Math.max(240, Math.min(window.innerWidth * 0.7, lrs.w0 + (e.clientX - lrs.x0))));
      var shown = w;
      if (view === 'scoping') leftWScope = w;
      else if (view === 'budget') leftWBudget = w;
      // the raw drag is the user's floor; the column floor still applies to
      // what is shown, so a narrow drag previews exactly what pointerup keeps
      else { leftWPlan = w; shown = planLeftW(); }
      document.documentElement.style.setProperty('--left-w', shown + 'px');
    });
    window.addEventListener('pointerup', function () {
      if (lrs) { lrs = null; saveLocal(); render(); }
    });
  })();

  // 70vw moves with the window, so Planning's derived width has to follow it
  (function () {
    var t = null;
    window.addEventListener('resize', function () {
      if (t) clearTimeout(t);
      t = setTimeout(function () {
        t = null;
        if (view !== 'planning' || leftCollapsed) return;
        document.documentElement.style.setProperty('--left-w', planLeftW() + 'px');
      }, 120);
    });
  })();

  // vertical resize (independent of the drag state machine)
  (function () {
    var rs = null;
    $('#resResize').addEventListener('pointerdown', function (e) {
      if (resCollapsed) return;
      rs = { y0: e.clientY, h0: resPanelH };
      e.preventDefault();
    });
    window.addEventListener('pointermove', function (e) {
      if (!rs) return;
      resPanelH = Math.max(48, Math.min(window.innerHeight * 0.6, rs.h0 - (e.clientY - rs.y0)));
      resBody.style.height = resPanelH + 'px';
    });
    window.addEventListener('pointerup', function () {
      if (rs) { rs = null; saveLocal(); }
    });
  })();

  // horizontal scroll sync with the timeline
  var syncing = false;
  board.addEventListener('scroll', function () {
    syncTodayClip();
    if (syncing) return;
    syncing = true; resBody.scrollLeft = board.scrollLeft; syncing = false;
  });
  resBody.addEventListener('scroll', function () {
    if (syncing) return;
    syncing = true; board.scrollLeft = resBody.scrollLeft; syncing = false;
  });

  // cell interactions: drag = fill with the source value, click = edit
  resGrid.addEventListener('pointerdown', function (e) {
    if (e.button !== 0) return;
    var grip = e.target.closest('.rr-grip');
    if (grip) {
      drag = { kind: 'rrow', mid: grip.closest('.rrow').dataset.mid, x0: e.clientX, y0: e.clientY, moved: false };
      e.preventDefault();
      return;
    }
    var cell = e.target.closest('.rh');
    if (!cell || e.target.tagName === 'INPUT') return;
    var rrow = cell.closest('.rrow');
    var m = null;
    state.team.forEach(function (x) { if (x.id === rrow.dataset.mid) m = x; });
    if (!m) return;
    var w = parseInt(cell.dataset.w, 10);
    drag = {
      kind: 'rfill', mid: m.id, w0: w, w1: w,
      val: RM.memberHoursForWeek(state.meta, m, w),
      x0: e.clientX, y0: e.clientY, moved: false
    };
    e.preventDefault();
  });

  function rfillMove(e) {
    var row = resGrid.querySelector('.rrow[data-mid="' + drag.mid + '"]');
    if (!row) return;
    var lane = row.querySelector('.rlane');
    var r = lane.getBoundingClientRect();
    drag.w1 = Math.max(0, Math.min(state.meta.numWeeks - 1, Math.floor((e.clientX - r.left) / weekPx)));
    var lo = Math.min(drag.w0, drag.w1), hi = Math.max(drag.w0, drag.w1);
    $$('.rh', row).forEach(function (c) {
      var w = parseInt(c.dataset.w, 10);
      c.classList.toggle('rh-fill', w >= lo && w <= hi);
    });
    dragTip.hidden = false;
    dragTip.style.left = (e.clientX + 14) + 'px';
    dragTip.style.top = (e.clientY - 34) + 'px';
    dragTip.textContent = fmtH(drag.val) + 'h × ' + (hi - lo + 1) + 'w';
  }

  function fillWeekHours(mids, w0, w1, val) {
    var lo = Math.min(w0, w1), hi = Math.max(w0, w1);
    commit('fill hours', function (s) {
      s.team.forEach(function (m) {
        if (mids.indexOf(m.id) === -1) return;
        m.weekHours = m.weekHours || {};
        for (var w = lo; w <= hi; w++) {
          var iso = RM.fmtISO(RM.weekStartDate(s.meta, w));
          if (val === RM.WEEK_HOURS) delete m.weekHours[iso];
          else m.weekHours[iso] = val;
        }
      });
    });
  }
  function rfillEnd(d) { fillWeekHours([d.mid], d.w0, d.w1, d.val); }

  // inline hour editor; Tab / Shift+Tab commits and hops to the next / previous cell
  function editHourCell(mid, w) {
    var row = resGrid.querySelector('.rrow[data-mid="' + mid + '"]');
    var cell = row && row.querySelector('.rh[data-w="' + w + '"]');
    if (!cell || cell.querySelector('input')) return;
    var m = null;
    state.team.forEach(function (x) { if (x.id === mid) m = x; });
    if (!m) return;
    var inp = document.createElement('input');
    inp.type = 'number';
    inp.min = '0';
    inp.step = '4';
    inp.value = fmtH(RM.memberHoursForWeek(state.meta, m, w));
    inp.className = 'rh-edit';
    cell.textContent = '';
    cell.appendChild(inp);
    inp.focus();
    inp.select();
    var done = false;
    function finish(saveIt) {
      if (done) return; done = true;
      var h = parseFloat(inp.value);
      if (saveIt && isFinite(h) && h >= 0) {
        commit('hours', function (s) {
          var m2 = null;
          s.team.forEach(function (x) { if (x.id === mid) m2 = x; });
          if (!m2) return;
          m2.weekHours = m2.weekHours || {};
          var iso = RM.fmtISO(RM.weekStartDate(s.meta, w));
          if (h === RM.WEEK_HOURS) delete m2.weekHours[iso];
          else m2.weekHours[iso] = h;
        });
      } else renderResources();
    }
    inp.addEventListener('blur', function () { finish(true); });
    inp.addEventListener('keydown', function (ev) {
      if (ev.key === 'Enter') finish(true);
      if (ev.key === 'Escape') finish(false);
      if (ev.key === 'Tab') {
        ev.preventDefault();
        finish(true);
        // hop across the spreadsheet: wrap between rows at either end
        var order = state.team.map(function (x) { return x.id; });
        var mi = order.indexOf(mid);
        var nw = w + (ev.shiftKey ? -1 : 1);
        if (nw >= state.meta.numWeeks && mi < order.length - 1) { mi += 1; nw = 0; }
        if (nw < 0 && mi > 0) { mi -= 1; nw = state.meta.numWeeks - 1; }
        if (nw >= 0 && nw < state.meta.numWeeks) {
          var nextMid = order[mi];
          requestAnimationFrame(function () { editHourCell(nextMid, nw); });
        }
      }
      ev.stopPropagation();
    });
  }

  // click (no drag) = inline edit
  resGrid.addEventListener('click', function (e) {
    if (dragConsumedClick) { dragConsumedClick = false; return; }
    var cell = e.target.closest('.rh');
    if (!cell || cell.querySelector('input')) return;
    editHourCell(cell.closest('.rrow').dataset.mid, parseInt(cell.dataset.w, 10));
  });

  // blank add row → type the new role in place
  resGrid.addEventListener('click', function (e) {
    var addRow = e.target.closest('[data-resadd]');
    if (!addRow || addRow.querySelector('input')) return;
    var lab = addRow.querySelector('.addrow-lab');
    var inp = document.createElement('input');
    inp.placeholder = 'Name — e.g. Alice…';
    inp.className = 'radd-input';
    lab.replaceWith(inp);
    inp.focus();
    var done = false;
    function finish(saveIt) {
      if (done) return; done = true;
      var v = inp.value.trim();
      if (saveIt && v) {
        commit('add person', function (s) {
          s.team.push({ id: RM.uid('t'), name: v, role: '', type: '', weekHours: {} });
        });
      } else renderResources();
    }
    inp.addEventListener('blur', function () { finish(true); });
    inp.addEventListener('keydown', function (ev) {
      if (ev.key === 'Enter') finish(true);
      if (ev.key === 'Escape') finish(false);
      ev.stopPropagation();
    });
  });

  function deleteRoleConfirm(mid) {
    var dm = null;
    state.team.forEach(function (x) { if (x.id === mid) dm = x; });
    if (!dm) return;
    confirmBox('Remove “' + esc(mLabel(dm)) + '”?', 'Their weekly hours disappear from capacity.', 'Remove', function () {
      commit('remove role', function (s) {
        s.team = s.team.filter(function (m) { return m.id !== mid; });
      });
    }, true);
  }

  // quick engagement window: weeks fully before the start / after the end go
  // to 0 h; zeroed weeks inside the window return to the default 40 h
  function roleDatesPopover(cx, cy, mid) {
    var m = null;
    state.team.forEach(function (x) { if (x.id === mid) m = x; });
    if (!m) return;
    var meta = state.meta;
    var first = -1, last = -1;
    for (var w = 0; w < meta.numWeeks; w++) {
      if (RM.memberHoursForWeek(meta, m, w) > 0) { if (first < 0) first = w; last = w; }
    }
    var sv = first >= 0 ? RM.fmtISO(RM.weekStartDate(meta, first)) : '';
    var ev = '';
    if (last >= 0) {
      var ld = new Date(RM.weekStartDate(meta, last));
      ld.setDate(ld.getDate() + 4);
      ev = RM.fmtISO(ld);
    }
    openPopover(cx, cy,
      '<div class="rd-pop">' +
      '<label>Start<input type="text" readonly class="cal-in" id="rdStart" value="' + sv + '"></label>' +
      '<label>End<input type="text" readonly class="cal-in" id="rdEnd" value="' + ev + '"></label>' +
      '<div class="po-actions"><button id="rdApply" class="primary">Apply</button></div>' +
      '</div>',
      function (host) {
        function apply() {
          var s0 = $('#rdStart', host).value, e0 = $('#rdEnd', host).value;
          closePopover();
          commit('role dates', function (s) {
            var t = null;
            s.team.forEach(function (x) { if (x.id === mid) t = x; });
            if (!t) return;
            t.weekHours = t.weekHours || {};
            for (var w2 = 0; w2 < s.meta.numWeeks; w2++) {
              var mon = RM.weekStartDate(s.meta, w2);
              var iso = RM.fmtISO(mon);
              var fri = new Date(mon);
              fri.setDate(fri.getDate() + 4);
              if ((s0 && RM.fmtISO(fri) < s0) || (e0 && iso > e0)) t.weekHours[iso] = 0;
              else if (t.weekHours[iso] === 0) delete t.weekHours[iso];
            }
          });
        }
        $('#rdApply', host).onclick = apply;
        host.addEventListener('keydown', function (ke) { if (ke.key === 'Enter') apply(); });
        $('#rdStart', host).focus();
      });
  }

  // right-click a role row for the full action menu
  resGrid.addEventListener('contextmenu', function (e) {
    var rrow = e.target.closest('.rrow[data-mid]');
    if (!rrow || e.target.closest('input')) return;
    e.preventDefault();
    e.stopPropagation(); // the document fallback must not replace this menu
    var mid = rrow.dataset.mid;
    var m = null;
    state.team.forEach(function (x) { if (x.id === mid) m = x; });
    if (!m) return;
    var cx = e.clientX, cy = e.clientY;
    openContextMenu(cx, cy, [
      { icon: 'pencil', label: 'Rename', fn: function () {
        var el = resGrid.querySelector('.rrow[data-mid="' + mid + '"] input[data-bud="name"]');
        if (el) { el.focus(); el.select(); }
      } },
      { icon: 'tags', label: 'Rate card…', fn: function () {
        var chipT = resGrid.querySelector('.rrow[data-mid="' + mid + '"] [data-bact="type"]');
        if (chipT) chipT.dispatchEvent(new MouseEvent('click', { bubbles: true }));
      } },
      { icon: 'layers', label: 'Workstream…', fn: function () {
        var chipW = resGrid.querySelector('.rrow[data-mid="' + mid + '"] [data-bact="ws"]');
        if (chipW) chipW.dispatchEvent(new MouseEvent('click', { bubbles: true }));
      } },
      state.meta.capacityEnabled ? { icon: 'gauge', label: 'Capacity…', fn: function () {
        // whichever capacity control the row shows: the × seat, the points
        // per sprint, or an untyped person's "set type" prompt
        var el = resGrid.querySelector('.rrow[data-mid="' + mid + '"] [data-rcap], .rrow[data-mid="' + mid + '"] [data-rpts], .rrow[data-mid="' + mid + '"] [data-runtyped]');
        if (el) el.dispatchEvent(new MouseEvent('click', { bubbles: true }));
      } } : null,
      { icon: 'calendar-range', label: 'Start / end dates…', fn: function () { roleDatesPopover(cx, cy, mid); } },
      { sep: true },
      { icon: 'trash-2', label: 'Remove role…', fn: function () { deleteRoleConfirm(mid); } }
    ].filter(Boolean));
  });

  // name / role text fields — same inline inputs as the budgeting rows
  resGrid.addEventListener('change', function (e) {
    var f = e.target.dataset && e.target.dataset.bud;
    var rrow = e.target.closest('.rrow[data-mid]');
    if (!f || !rrow || (f !== 'name' && f !== 'role')) return;
    var nmid = rrow.dataset.mid;
    var tv = e.target.value.trim();
    commit(f === 'name' ? 'rename person' : 'person role', function (s) {
      s.team.forEach(function (m) { if (m.id === nmid) m[f] = tv; });
    });
  });

  // capacity factor: click-to-type inline number
  resGrid.addEventListener('click', function (e) {
    var chip = e.target.closest('[data-rcap]');
    if (!chip || chip.querySelector('input')) return;
    var mid = chip.dataset.rcap;
    var m = null;
    state.team.forEach(function (x) { if (x.id === mid) m = x; });
    if (!m) return;
    var inp = document.createElement('input');
    inp.type = 'number';
    inp.min = '0';
    inp.step = '0.1';
    inp.value = m.capacity != null ? m.capacity : 1;
    inp.className = 'hc-edit';
    inp.style.width = '34px';
    chip.textContent = '';
    chip.appendChild(inp);
    inp.focus();
    inp.select();
    var done = false;
    function finish(saveIt) {
      if (done) return; done = true;
      var v = inp.value.trim() === '' ? 0 : parseFloat(inp.value); // blank = 0
      if (saveIt && isFinite(v) && v >= 0) {
        commit('role capacity', function (s) {
          s.team.forEach(function (x) { if (x.id === mid) x.capacity = v; });
        });
      } else renderResources();
    }
    inp.addEventListener('blur', function () { finish(true); });
    inp.addEventListener('keydown', function (ev) {
      ev.stopPropagation();
      if (ev.key === 'Enter') finish(true);
      if (ev.key === 'Escape') finish(false);
    });
  });

  // story points per sprint: same click-to-type chip, blank = document default
  resGrid.addEventListener('click', function (e) {
    var chip = e.target.closest('[data-rpts]');
    if (!chip || chip.querySelector('input')) return;
    var mid = chip.dataset.rpts;
    var m = null;
    state.team.forEach(function (x) { if (x.id === mid) m = x; });
    if (!m) return;
    var inp = document.createElement('input');
    inp.type = 'number';
    inp.min = '0';
    inp.step = '1';
    inp.value = m.points != null ? m.points : '';
    inp.placeholder = String(RM.memberPoints(state, m));
    inp.className = 'hc-edit';
    inp.style.width = '34px';
    chip.textContent = '';
    chip.appendChild(inp);
    inp.focus();
    inp.select();
    var done = false;
    function finish(saveIt) {
      if (done) return; done = true;
      var raw = inp.value.trim();
      var v = raw === '' ? null : parseFloat(raw); // blank = the document default
      if (saveIt && (v === null || (isFinite(v) && v >= 0))) {
        commit('role points', function (s) {
          s.team.forEach(function (x) { if (x.id === mid) x.points = v; });
        });
      } else renderResources();
    }
    inp.addEventListener('blur', function () { finish(true); });
    inp.addEventListener('keydown', function (ev) {
      ev.stopPropagation();
      if (ev.key === 'Enter') finish(true);
      if (ev.key === 'Escape') finish(false);
    });
  });

  // an untyped person's placeholder opens the capacity-type picker
  resGrid.addEventListener('click', function (e) {
    var chip = e.target.closest('[data-runtyped]');
    if (chip) openMemberCapDropdown(chip, chip.dataset.runtyped);
  });
  resGrid.addEventListener('keydown', function (e) {
    var chip = e.target.closest && e.target.closest('[data-runtyped]');
    if (chip && (e.key === 'Enter' || e.key === ' ')) {
      e.preventDefault();
      openMemberCapDropdown(chip, chip.dataset.runtyped);
    }
  });

  // rate-card / workstream chips → the SAME dropdowns as the budgeting rows
  resGrid.addEventListener('click', function (e) {
    var chip = e.target.closest('[data-bact]');
    var rrow = e.target.closest('.rrow[data-mid]');
    if (!chip || !rrow) return;
    if (chip.dataset.bact === 'type') openMemberTypeDropdown(chip, rrow.dataset.mid);
    else if (chip.dataset.bact === 'cap') openMemberCapDropdown(chip, rrow.dataset.mid);
    else openMemberWsDropdown(chip, rrow.dataset.mid);
  });

  // member reorder
  function rrowMove(e) {
    if (!drag.indicator) {
      drag.indicator = document.createElement('div');
      drag.indicator.className = 'rrow-indicator';
      resGrid.appendChild(drag.indicator);
      var src = resGrid.querySelector('.rrow[data-mid="' + drag.mid + '"]');
      if (src) src.classList.add('drag-row-ghost');
    }
    var rows = $$('.rrow[data-mid]', resGrid);
    drag.before = null;
    var g = resGrid.getBoundingClientRect();
    var y = null;
    for (var i = 0; i < rows.length; i++) {
      var r = rows[i].getBoundingClientRect();
      if (e.clientY < r.top + r.height / 2) { drag.before = rows[i].dataset.mid; y = r.top - g.top; break; }
    }
    if (y == null && rows.length) y = rows[rows.length - 1].getBoundingClientRect().bottom - g.top;
    drag.indicator.style.top = (y || 0) + 'px';
  }
  function rrowEnd(d) {
    var src = resGrid.querySelector('.rrow[data-mid="' + d.mid + '"]');
    if (src) src.classList.remove('drag-row-ghost');
    if (d.before === d.mid) { renderResources(); return; }
    commit('reorder member', function (s) {
      var m = null;
      s.team.forEach(function (x) { if (x.id === d.mid) m = x; });
      if (!m) return;
      s.team = s.team.filter(function (x) { return x.id !== d.mid; });
      var at = d.before ? s.team.map(function (x) { return x.id; }).indexOf(d.before) : s.team.length;
      if (at < 0) at = s.team.length;
      s.team.splice(at, 0, m);
    });
  }

  // ------------------------------------------------------------ personal settings
  // One set of controls, used by Setup → Personal and the start page modal.
  // Everything here is per-machine (UI_KEY / THEME_KEY), never document data.
  // part: 'appearance' | 'behavior' | undefined (both — the start page modal)
  function personalFieldsHtml(part) {
    var desktop = !!window.HeadwayDesktop;
    function chk(key, label, on) {
      return '<label class="p-check" style="margin-top:7px">' +
        '<input type="checkbox" data-pref="' + key + '"' + (on ? ' checked' : '') + '> ' + label + '</label>';
    }
    var themeSeg = '<div class="seg">' + THEME_CHOICES.map(function (t) {
      return '<button data-pref-theme="' + t[0] + '"' + (themePref === t[0] ? ' class="on"' : '') + '>' +
        t[1].replace(' theme', '') + '</button>';
    }).join('') + '</div>';
    function snapSeg(kind) {
      return '<div class="seg">' + SNAP_MODES.map(function (mode) {
        return '<button data-pref-snap="' + kind + ':' + mode + '"' + (snapModeFor(kind) === mode ? ' class="on"' : '') + '>' +
          SNAP_LABELS[mode].charAt(0).toUpperCase() + SNAP_LABELS[mode].slice(1) + '</button>';
      }).join('') + '</div>';
    }
    var appearance =
      '<div class="m-sec pf-name"><label>Your name</label>' +
      '<input type="text" data-pref-user maxlength="60" style="width:100%" placeholder="e.g. Alex Rivera" value="' + esc(userName()) + '">' +
      '<div class="m-hint">Recorded with your edits in Version history. Stored on this machine only.</div></div>' +
      '<div class="m-sec"><label>Theme</label>' + themeSeg + '</div>';
    var behavior = '<div class="m-sec"><label>Feature snap</label>' + snapSeg('feature') +
      '<div class="m-hint">Grid for dragging, resizing and placing feature bars.</div></div>' +
      '<div class="m-sec"><label>Story snap</label>' + snapSeg('story') +
      '<div class="m-hint">Same for story bars — a sprint by default.</div></div>' +
      '<div class="m-sec"><label>Timeline</label>' +
      chk('deps', 'Dependency arrows', depsMode === 'on') +
      chk('crit', 'Critical path highlight', showCrit) +
      chk('cap', 'Capacity rows', showCap) +
      chk('autoOrder', 'Auto-order rows by start', autoOrder) +
      '</div>' +
      '<div class="m-sec"><label>Grouping</label>' +
      (state && state.meta.workstreamsEnabled ? chk('groupWs', 'Group by workstream', groupWs) : '') +
      chk('groupEpic', 'Group by epic', groupEpic) +
      '</div>' +
      '<div class="m-sec"><label>Color bars by</label><div class="seg">' + COLOR_MODES.map(function (cm) {
        return '<button data-pref-color="' + cm[0] + '"' + (colorBy === cm[0] ? ' class="on"' : '') + '>' + cm[1] + '</button>';
      }).join('') + '</div></div>' +
      (desktop
        ? '<div class="m-sec"><label>Files</label>' + chk('autoSave', 'Auto-save to the open file', autoSave) + '</div>'
        : '');
    if (part === 'appearance') return appearance;
    if (part === 'behavior') return behavior;
    return appearance + behavior;
  }
  // delegated wiring; `after` refreshes whatever surface hosts the fields
  function wirePersonalFields(host, after) {
    host.addEventListener('click', function (e) {
      var tb = e.target.closest('[data-pref-theme]');
      if (tb) { setTheme(tb.dataset.prefTheme); after(); return; }
      var sb = e.target.closest('[data-pref-snap]');
      if (sb) { var sp = sb.dataset.prefSnap.split(':'); setSnapMode(sp[0], sp[1]); saveLocal(); renderTopbar(); after(); return; }
      var cb = e.target.closest('[data-pref-color]');
      if (cb) { setColorBy(cb.dataset.prefColor); saveLocal(); render(); after(); return; }
    });
    host.addEventListener('change', function (e) {
      if (e.target.dataset.prefUser != null) { setUserName(e.target.value); return; }
      var key = e.target.dataset.pref;
      if (!key) return;
      var on = e.target.checked;
      if (key === 'deps') depsMode = on ? 'on' : 'none';
      else if (key === 'crit') showCrit = on;
      else if (key === 'cap') showCap = on;
      else if (key === 'autoOrder') {
        autoOrder = on;
        if (on) commit('auto-order', function (s) { RM.sortItemsByStart(s); });
      }
      else if (key === 'groupWs') groupWs = on;
      else if (key === 'groupEpic') groupEpic = on;
      else if (key === 'autoSave') { autoSave = on; if (on) scheduleAutoSave(); }
      else return;
      saveLocal();
      render();
      after();
    });
  }

  function personalSettingsModal() {
    openModal(
      '<div class="modal" style="width:460px">' +
      '<div class="m-head"><h2>Personal settings</h2><button class="p-close" data-m="x"><i data-lucide="x"></i></button></div>' +
      '<div class="m-body"><div id="ppFields">' + personalFieldsHtml() + '</div></div>' +
      '<div class="m-foot"><button data-m="x2" class="primary">Done</button></div></div>',
      function (host) {
        $('[data-m=x]', host).onclick = closeModal;
        $('[data-m=x2]', host).onclick = closeModal;
        wirePersonalFields(host, function () {
          var f = $('#ppFields', host);
          if (f) { f.innerHTML = personalFieldsHtml(); if (window.lucide) lucide.createIcons(); }
        });
      });
  }

  // ------------------------------------------------------------ start page
  // The app opens here: pick a recent file, create a project (always saved
  // to disk first), open a file, or adjust personal settings. A tab refresh
  // mid-session skips it (sessionStorage flag); a fresh launch shows it.
  var RECENTS_KEY = 'headway-recents-v1';
  var SESSION_KEY = 'headway-in-editor';

  function loadRecents() {
    try {
      var r = JSON.parse(localStorage.getItem(RECENTS_KEY) || '[]');
      return Array.isArray(r) ? r : [];
    } catch (e) { return []; }
  }
  function saveRecents(list) {
    try { localStorage.setItem(RECENTS_KEY, JSON.stringify(list.slice(0, 12))); } catch (e) { /* storage optional */ }
  }
  // upsert a path at the top of the recents; title comes from the live doc
  function noteRecent(path) {
    if (!path) return;
    var list = loadRecents().filter(function (r) { return r.path !== path; });
    list.unshift({ path: path, title: (state && state.meta.title) || '', at: Date.now() });
    saveRecents(list);
    if (document.body.classList.contains('start')) renderStartPage();
  }
  function dropRecent(path) {
    saveRecents(loadRecents().filter(function (r) { return r.path !== path; }));
  }

  function relTime(t) {
    if (!t) return '';
    var s = (Date.now() - t) / 1000;
    if (s < 60) return 'just now';
    if (s < 3600) return Math.floor(s / 60) + ' min ago';
    if (s < 86400 * 2) return Math.floor(s / 3600) + ' h ago';
    if (s < 86400 * 14) return Math.floor(s / 86400) + ' days ago';
    return new Date(t).toLocaleDateString();
  }

  function enterEditor() {
    document.body.classList.remove('start');
    $('#startPage').hidden = true;
    try { sessionStorage.setItem(SESSION_KEY, '1'); } catch (e) { /* storage optional */ }
    render();
    if (view === 'planning') requestAnimationFrame(goToday);
  }
  function showStart() {
    document.body.classList.add('start');
    renderStartPage();
    $('#startPage').hidden = false;
    try { sessionStorage.removeItem(SESSION_KEY); } catch (e) { /* storage optional */ }
  }

  function renderStartPage() {
    var host = $('#startBody');
    if (!host) return;
    var desktop = !!window.HeadwayDesktop;
    var hasLocal = false;
    try { hasLocal = !!localStorage.getItem(LS_KEY); } catch (e) { /* storage optional */ }
    var recents = desktop ? loadRecents() : [];
    var rows = recents.map(function (r) {
      var base = String(r.path).replace(/^.*[\\/]/, '');
      return '<div class="sp-row" role="button" tabindex="0" data-sp-open="' + esc(r.path) + '">' +
        '<span class="sp-ico"><i data-lucide="file-spreadsheet"></i></span>' +
        '<span class="sp-rmain"><span class="sp-rtitle">' + esc(r.title || base) + '</span>' +
        '<span class="sp-rpath">' + esc(r.path) + '</span></span>' +
        '<span class="sp-rtime">' + esc(relTime(r.at)) + '</span>' +
        '<button class="sp-rx" data-sp-drop="' + esc(r.path) + '" title="Remove from this list (keeps the file)"><i data-lucide="x"></i></button>' +
        '</div>';
    }).join('');
    // a localStorage session without a recents entry (browser, or pre-file
    // desktop work) can be picked up where it left off
    var continueCard = hasLocal && (!desktop || !recents.length)
      ? '<div class="sp-row" role="button" tabindex="0" data-sp-continue>' +
        '<span class="sp-ico"><i data-lucide="history"></i></span>' +
        '<span class="sp-rmain"><span class="sp-rtitle">' + esc((state && state.meta.title) || 'Last session') + '</span>' +
        '<span class="sp-rpath">Continue where you left off — stored in this ' + (desktop ? 'app' : 'browser') + '</span></span></div>'
      : '';
    host.innerHTML =
      '<div class="sp-inner">' +
      '<div class="sp-hero">' +
      '<div class="tb-mark sp-mark" aria-hidden="true"><span></span><span></span><span></span></div>' +
      '<div class="sp-brand"><h1>Headway</h1><div class="sp-sub">Roadmap planner</div></div>' +
      '<div class="sp-hero-btns">' +
      '<button data-sp-settings title="Personal settings"><i data-lucide="settings"></i>Settings</button>' +
      (desktop ? updateButtonHtml() : '') +
      '</div></div>' +
      '<div class="sp-actions">' +
      '<button class="primary sp-big" data-sp-new><i data-lucide="file-plus-2"></i>New project…</button>' +
      '<button class="sp-big" data-sp-opendlg><i data-lucide="folder-open"></i>Open…</button>' +
      '</div>' +
      '<div class="sp-recent-hd">Recent</div>' +
      '<div class="sp-recents">' +
      ((continueCard + rows) || '<div class="sp-empty">Nothing yet — projects you create or open appear here.</div>') +
      '</div>' +
      (appVersion() ? '<div class="sp-version"><button class="sp-verbtn" data-sp-notes title="What\u2019s new in this version">Headway ' + esc(appVersion()) + '</button></div>' : '') +
      '</div>';
    if (window.lucide) lucide.createIcons();
  }

  // app version for the start page footer (desktop only — the browser build
  // has no version of its own)
  function appVersion() {
    return (window.HeadwayDesktop && window.HeadwayDesktop.appVersion) || '';
  }

  // ---- start page update button (desktop only). Mirrors the shared updater
  // in desktop.js: idle → "Check for updates", busy → "Checking…" /
  // "Downloading…", ready → a pulsing "Update to x.y.z" that installs.
  function updater() { return (window.HeadwayDesktop && window.HeadwayDesktop.updater) || null; }
  function updateButtonHtml() {
    var u = updater();
    var st = u ? u.state : 'idle';
    var label, icon = 'refresh-cw', extra = '';
    if (st === 'checking') { label = 'Checking…'; extra = ' disabled'; }
    else if (st === 'downloading') { label = 'Downloading ' + esc(u.version) + '…'; extra = ' disabled'; }
    else if (st === 'ready') { label = 'Update to ' + esc(u.version); extra = ' class="sp-update-ready"'; icon = 'arrow-down-circle'; }
    else if (st === 'installing') { label = 'Updating…'; extra = ' disabled'; }
    else { label = 'Check for updates'; }
    return '<button data-sp-update title="' + (st === 'ready' ? 'Version ' + esc(u.version) + ' downloaded \u2014 install and restart' : 'Check GitHub for a newer Headway') + '"' + extra + '>' +
      '<i data-lucide="' + icon + '"></i>' + label + '</button>';
  }
  function updateButtonClick() {
    var u = updater();
    if (!u) { toast('Updates are only available in the desktop app'); return; }
    if (u.state === 'ready') u.install();
    else u.check(true);
  }

  // ---- release notes: the bundled CHANGELOG.md, one "## <version>" section
  // per release. Shown once per version after an update (and on first
  // launch); the version footer reopens it any time.
  var NOTES_SEEN_KEY = 'headway-notes-seen-v1';
  var changelogText = null; // cached CHANGELOG.md
  function releaseNotesFor(md, version) {
    if (!md || !version) return '';
    var v = String(version).replace(/^v/, '').replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
    var m = String(md).replace(/\r\n?/g, '\n')
      .match(new RegExp('^## ' + v + '(?=\\s|$)[^\\n]*\\n([\\s\\S]*?)(?=^## |(?![\\s\\S]))', 'm'));
    return m ? m[1].trim() : '';
  }
  function loadChangelog() {
    if (changelogText !== null) return Promise.resolve(changelogText);
    if (typeof fetch !== 'function') return Promise.resolve('');
    return fetch('CHANGELOG.md').then(function (r) { return r.ok ? r.text() : ''; })
      .catch(function () { return ''; })
      .then(function (t) { changelogText = t || ''; return changelogText; });
  }
  function releaseNotesModal(version, notes) {
    var body = notes
      ? (window.HeadwayAI && HeadwayAI.md ? HeadwayAI.md(notes) : '<pre>' + esc(notes) + '</pre>')
      : '<p class="m-hint">No release notes for this version.</p>';
    openModal(
      '<div class="modal notes-modal" style="width:520px">' +
      '<div class="m-head"><h2>What\u2019s new in Headway ' + esc(version) + '</h2><button class="p-close" data-m="x"><i data-lucide="x"></i></button></div>' +
      '<div class="m-body"><div class="notes-md">' + body + '</div>' +
      '<div class="m-hint notes-all"><a href="https://github.com/smo-key/headway/releases" target="_blank" rel="noopener">All releases on GitHub</a></div></div>' +
      '<div class="m-foot"><button data-m="ok" class="primary">Got it</button></div></div>',
      function (host) {
        $('[data-m=x]', host).onclick = closeModal;
        $('[data-m=ok]', host).onclick = closeModal;
      });
  }
  // open the notes for the running version (md: optional preloaded text)
  function openReleaseNotes(md) {
    var v = appVersion();
    if (!v) return;
    var show = function (text) { releaseNotesModal(v, releaseNotesFor(text, v)); };
    if (typeof md === 'string') show(md);
    else loadChangelog().then(show);
  }
  // once per version: called by desktop.js when the app version is known
  function maybeShowReleaseNotes(md) {
    var v = appVersion();
    if (!v) return false;
    var seen = '';
    try { seen = localStorage.getItem(NOTES_SEEN_KEY) || ''; } catch (e) { /* storage optional */ }
    if (seen === v) return false;
    try { localStorage.setItem(NOTES_SEEN_KEY, v); } catch (e) { /* storage optional */ }
    var go = function (text) {
      var notes = releaseNotesFor(text, v);
      if (notes) releaseNotesModal(v, notes); // nothing to say → stay quiet
    };
    if (typeof md === 'string') go(md);
    else loadChangelog().then(go);
    return true;
  }

  $('#startPage').addEventListener('click', function (e) {
    var drop = e.target.closest('[data-sp-drop]');
    if (drop) { dropRecent(drop.dataset.spDrop); renderStartPage(); return; }
    var open = e.target.closest('[data-sp-open]');
    if (open) { guardUnsaved(function () { openRecent(open.dataset.spOpen); }); return; }
    if (e.target.closest('[data-sp-continue]')) { enterEditor(); return; }
    if (e.target.closest('[data-sp-new]')) { guardUnsaved(openWizard); return; }
    if (e.target.closest('[data-sp-opendlg]')) {
      guardUnsaved(function () {
        if (window.HeadwayDesktop) HeadwayDesktop.openDialog();
        else $('#filePick').click();
      });
      return;
    }
    if (e.target.closest('[data-sp-settings]')) { personalSettingsModal(); return; }
    if (e.target.closest('[data-sp-update]')) { updateButtonClick(); return; }
    if (e.target.closest('[data-sp-notes]')) { openReleaseNotes(); return; }
  });
  $('#startPage').addEventListener('keydown', function (e) {
    if (e.key !== 'Enter' && e.key !== ' ') return;
    var row = e.target.closest('.sp-row');
    if (row) { e.preventDefault(); row.click(); }
  });

  function openRecent(path) {
    if (!window.HeadwayDesktop || !HeadwayDesktop.openPath) return;
    HeadwayDesktop.openPath(path).catch(function (err) {
      toast('Could not open “' + String(path).replace(/^.*[\\/]/, '') + '” — ' +
        (err && err.message || err), 'err');
      dropRecent(path);
      renderStartPage();
    });
  }

  // every project starts life as a file on disk: desktop picks a location
  // first; the browser fires the .xlsx download the moment it's created
  // (the new-project wizard gathers the settings first — see openWizard)
  function adoptProject(st, savedPath) {
    replaceState('new project', st);
    selectedId = null;
    docSaved = true;
    updateSaveBtn();
    if (savedPath) noteRecent(savedPath);
    enterEditor();
  }

  // opts.onDone after the project is adopted; opts.onFail when no file was
  // chosen or the export failed (the wizard reopens with its draft)
  function createProjectOnDisk(st, opts) {
    opts = opts || {};
    var fname = ((st.meta.title || '').replace(/[\\/:*?"<>|]+/g, '').trim() || 'Roadmap') + '.xlsx';
    RMExcel.exportWorkbook(st, uiSnapshot()).then(function (blob) {
      if (window.HeadwayDesktop) {
        // the file must exist before the project does — Save dialog first
        return HeadwayDesktop.saveBlob(blob, fname, true).then(function (path) {
          if (!path) { toast('Project not created — no file chosen'); if (opts.onFail) opts.onFail(); return; }
          // the dialog may have picked a different name — the filename wins
          st.meta.title = titleFromFileName(HeadwayDesktop.basename(path));
          adoptProject(st, path);
          if (opts.onDone) opts.onDone();
          toast('Created “' + st.meta.title + '” — ' + HeadwayDesktop.basename(path));
        });
      }
      var a = document.createElement('a');
      a.href = URL.createObjectURL(blob);
      a.download = fname;
      a.click();
      setTimeout(function () { URL.revokeObjectURL(a.href); }, 4000);
      adoptProject(st, null);
      if (opts.onDone) opts.onDone();
      toast('Created “' + st.meta.title + '” — saved ' + fname + ' to your downloads');
    }).catch(function (err) {
      toast('Could not create the project: ' + (err && err.message || err), 'err');
      if (opts.onFail) opts.onFail();
    });
  }

  // ------------------------------------------------------------ files
  var savingNow = false;
  function updateSaveBtn() {
    var sb = $('#btnSave');
    if (savingNow) return;
    var mode = docSaved ? 'saved' : 'save';
    if (sb.dataset.mode === mode) return;
    sb.dataset.mode = mode;
    sb.disabled = docSaved;
    sb.innerHTML = docSaved
      ? '<i data-lucide="check"></i>Saved'
      : '<i data-lucide="download"></i>Save';
    if (window.lucide) lucide.createIcons();
  }

  function doSave(forceDialog, quiet) {
    if (wz) return null; // the open document is stashed behind the wizard
    var btn = $('#btnSave');
    savingNow = true;
    btn.disabled = true; btn.textContent = 'Saving…';
    function restoreBtn() {
      savingNow = false;
      btn.dataset.mode = '';
      updateSaveBtn();
    }
    // the exact document JSON this export embeds — the desktop shell keeps it
    // to recognize a sync client's rewrite of this very save (same document,
    // different bytes) and not reload over it
    var stateJson = RMExcel.stateJsonOf(state);
    return RMExcel.exportWorkbook(state, uiSnapshot()).then(function (blob) {
      var name = saveFileName();
      if (window.HeadwayDesktop) { // desktop: write straight to disk
        return HeadwayDesktop.saveBlob(blob, name, forceDialog, stateJson).then(function (path) {
          if (!path) return; // dialog canceled
          lastExport = new Date().toTimeString().slice(0, 5);
          docSaved = true;
          saveLocal();
          noteRecent(path); // keep the start page's title + timestamp fresh
          if (!quiet) toast('Saved ' + HeadwayDesktop.basename(path));
          // Save As under a different name: the filename wins — retitle the
          // doc (the autosave that follows rewrites the file to match)
          var ft = titleFromFileName(HeadwayDesktop.basename(path));
          if (ft !== state.meta.title) commit('title', function (s) { s.meta.title = ft; });
        });
      }
      var a = document.createElement('a');
      a.href = URL.createObjectURL(blob);
      a.download = name;
      a.click();
      setTimeout(function () { URL.revokeObjectURL(a.href); }, 4000);
      lastExport = new Date().toTimeString().slice(0, 5);
      docSaved = true;
      saveLocal();
      toast('Saved ' + name);
    }).catch(function (err) {
      toast('Export failed: ' + err.message, 'err');
    }).finally(function () {
      restoreBtn();
    });
  }
  $('#btnSave').addEventListener('click', function () { doSave(false); });

  // ------------------------------------------------- unsaved-work guard
  // Anything that would drop the open document — closing the window (the
  // desktop shell calls this from onCloseRequested), opening or creating
  // another project — runs through here first.
  function unsavedNow() { return sessionEdited && !docSaved; }
  function guardUnsaved(proceed) {
    // the new-project wizard goes first (its draft is never saved anywhere);
    // then the open document gets the usual question
    if (wz) {
      if (!wz.dirty) { closeWizard(true); guardUnsaved(proceed); return; }
      confirmBox('Discard this new project?', 'The settings you picked are not saved anywhere yet.', 'Discard',
        function () { closeWizard(true); guardUnsaved(proceed); }, true);
      return;
    }
    flushPanelEdit(); // typing still in the field counts as an edit too
    if (!unsavedNow()) { proceed(); return; }
    // autosave already owns this doc's file: flush the pending write instead
    // of asking a question the user has answered by turning autosave on
    if (autoSave && window.HeadwayDesktop && HeadwayDesktop.currentPath()) {
      clearTimeout(autoSaveTimer);
      (doSave(false, true) || Promise.resolve()).then(function () {
        if (docSaved) proceed();
      });
      return;
    }
    openModal(
      '<div class="modal" style="width:440px">' +
      '<div class="m-head"><h2>Unsaved changes</h2></div>' +
      '<div class="m-body"><div style="font-size:13px;color:var(--ink-2);line-height:1.5">' +
      '“' + esc(state.meta.title || 'This roadmap') + '” has changes that aren’t saved' +
      (window.HeadwayDesktop ? ' to a file' : ' to an .xlsx file') + ' yet.</div></div>' +
      '<div class="m-foot"><button data-m="gdiscard">Don’t save</button>' +
      '<button data-m="cancel">Cancel</button>' +
      '<button data-m="gsave" class="primary"><i data-lucide="download"></i>Save</button></div></div>',
      function (host) {
        $('[data-m=cancel]', host).onclick = closeModal;
        $('[data-m=gdiscard]', host).onclick = function () { closeModal(); proceed(); };
        $('[data-m=gsave]', host).onclick = function () {
          closeModal();
          (doSave(false) || Promise.resolve()).then(function () {
            // a canceled Save dialog leaves the doc unsaved — stay put
            if (docSaved) proceed();
          });
        };
      });
  }

  // browser build: the tab is the app — closing it with unsaved work gets the
  // native "leave site?" question (the desktop shell asks via onCloseRequested)
  window.addEventListener('beforeunload', function (e) {
    if (window.__TAURI__ || !unsavedNow()) return;
    e.preventDefault();
    e.returnValue = ''; // legacy engines need the assignment to show the prompt
  });

  // reload = the open file changed on disk underneath us: the DATA updates,
  // the screen does not — view, prefs, selection, scroll offsets, open
  // dialogs and the field being edited all stay exactly where they were
  function loadWorkbookBuffer(buf, name, reload) {
    return RMExcel.importWorkbook(buf).then(function (r) {
      var capTypeChanges = r.capTypeChanges || 0; // people moved to their role's type
      // the file's name IS the roadmap's title (minus .xlsx) — a rename on
      // disk or a differing embedded title resolves in the filename's favor
      if (name) r.state.meta.title = titleFromFileName(name);
      // the wizard holds a draft: a reload of the open file updates the
      // document stashed behind it; opening another file ends the wizard
      if (wz && reload) { wizardStashReload(r.state); return; }
      if (wz) closeWizard(true);
      if (reload) {
        // keep whatever selection still exists in the new document
        var keepSel = selectedId && RM.itemById(r.state, selectedId) ? selectedId : null;
        var keepStory = keepSel && selStory && storyById(RM.itemById(r.state, keepSel), selStory) ? selStory : null;
        var keepMulti = multiSel ? multiSel.filter(function (id) { return !!RM.itemById(r.state, id); }) : null;
        selectedId = keepSel;
        selStory = keepStory;
        reloadingDoc = true;
        try { adoptState(r.state); } finally { reloadingDoc = false; }
        if (keepMulti && keepMulti.length > 1) { multiSel = keepMulti; render(); }
        updateSaveBtn();
        return;
      }
      if (r.ui) applyUi(r.ui); // the file carries the browser prefs too
      adoptState(r.state); // fresh from disk — matches its file
      updateSaveBtn();
      selectedId = null;
      enterEditor(); // opening a file always lands in the editor
      // a clean tool-state load needs no announcement; parsing a foreign
      // template layout is worth a note
      if (r.source !== 'tool') {
        toast('Parsed “' + name + '” from the template layout');
      }
      // people now take their role's capacity type: say so when that moved anyone
      capTypeToast(capTypeChanges);
    });
  }

  $('#filePick').addEventListener('change', function (e) {
    var file = e.target.files[0];
    e.target.value = '';
    if (!file) return;
    file.arrayBuffer().then(function (buf) {
      return loadWorkbookBuffer(buf, file.name);
    }).catch(function (err) {
      toast('Could not open: ' + err.message, 'err');
    });
  });

  // hooks for the Tauri shell (js/desktop.js); harmless in a plain browser
  window.HeadwayApp = {
    toast: toast,
    loadBuffer: loadWorkbookBuffer,
    saveFileName: saveFileName,
    save: doSave,
    unsavedNow: unsavedNow,
    guardUnsaved: guardUnsaved,
    autoSaveOn: function () { return !!autoSave; },
    menuItems: menuItems,
    noteRecent: noteRecent,
    renderStartPage: renderStartPage,
    maybeShowReleaseNotes: maybeShowReleaseNotes,
    openReleaseNotes: openReleaseNotes,
    openModal: openModal,
    closeModal: closeModal,
    openDropdown: openDropdown,
    openSetup: function (tab) { setupTab = normSetupTab(tab); view = 'setup'; saveLocal(); render(); },
    wizard: {
      open: openWizard,
      go: function (k) { wizardGo(k); },
      create: wizardCreate,
      close: closeWizard,
      isOpen: function () { return !!wz; }
    },
    // hooks for the AI assistant (js/ai.js): reads are clones, every write
    // goes through commit() so it lands in undo + Version history (as
    // "<name> · AI")
    ai: {
      state: function () { return RM.clone(wz ? wz.stash.state : state); },
      hasDoc: function () { return !!state && !document.body.classList.contains('start'); },
      commit: function (label, mutate) {
        if (wz) { withStash(function () { aiActor = true; try { commit(label, mutate); } finally { aiActor = false; } }); return; }
        aiActor = true;
        try { commit(label, mutate); } finally { aiActor = false; }
      },
      // Auto timeline for one phase (id or name), or every real phase when
      // none is given; returns the number of changes
      autoTimelineNow: function (phase) {
        if (phase == null || phase === '') return autoTimelineAll();
        var ph = state.phases.filter(function (p) { return p.id === phase || p.name === phase; })[0];
        return ph && !ph.bucket ? autoTimelinePhase(ph.id) : 0;
      },
      validation: function () { return validation || RM.validate(state); },
      userName: userName,
      ui: function () {
        var sel = selectedId ? RM.itemById(state, selectedId) : null;
        return {
          view: view, selectedNum: sel ? sel.num : null, theme: themePref, snapFeat: snapFeat, snapStory: snapStory, weekPx: weekPx,
          deps: depsMode === 'on', crit: showCrit, cap: showCap, autoOrder: autoOrder,
          groupWs: groupWs, groupEpic: groupEpic, autoSave: autoSave, detailMode: detailMode,
          desktop: !!window.HeadwayDesktop, userName: userName()
        };
      },
      setPref: function (key, val) {
        if (key === 'theme') { if (['system', 'light', 'dark'].indexOf(val) !== -1) setTheme(val); return true; }
        if (key === 'userName') { setUserName(val); return true; }
        if (key === 'snapFeat' || key === 'snapStory') { if (SNAP_MODES.indexOf(val) === -1) return false; setSnapMode(key === 'snapStory' ? 'story' : 'feature', val); }
        else if (key === 'deps') depsMode = val ? 'on' : 'none';
        else if (key === 'crit') showCrit = !!val;
        else if (key === 'cap') showCap = !!val;
        else if (key === 'autoOrder') {
          autoOrder = !!val;
          if (autoOrder) commit('auto-order', function (s) { RM.sortItemsByStart(s); });
        }
        else if (key === 'groupWs') groupWs = !!val;
        else if (key === 'groupEpic') groupEpic = !!val;
        else if (key === 'colorBy') { if (RM.COLOR_MODES.indexOf(val) === -1) return false; setColorBy(val); }
        else if (key === 'autoSave') { autoSave = !!val; if (autoSave) scheduleAutoSave(); }
        else if (key === 'detailMode') { if (['feature', 'story'].indexOf(val) === -1) return false; detailMode = val; }
        else return false;
        saveLocal();
        render();
        return true;
      },
      setView: function (v) {
        if (['planning', 'scoping', 'setup', 'budget', 'reports', 'history', 'prio', 'sprints'].indexOf(v) === -1) return false;
        if (!RM.appEnabled(state, v)) return false; // switched off in Setup → Views
        flushPanelEdit();
        view = v;
        if (view === 'history') { vhSel = null; vhPick = []; vhTab = null; }
        saveLocal();
        render();
        return true;
      },
      selectNum: function (num) {
        var it = RM.itemByNum(state, num);
        if (!it) return false;
        if (view !== 'planning' && view !== 'scoping') { view = 'planning'; saveLocal(); }
        select(it.id, true);
        return true;
      },
      openSettings: function () {
        setupTab = 'ai';
        view = 'setup';
        saveLocal();
        render();
      },
      toast: toast
    }
  };

  // ------------------------------------------------------------ keyboard
  window.addEventListener('keydown', function (e) {
    var ae = document.activeElement;
    // contenteditable editors (rich description, scoping cells) count as fields too
    var inField = /INPUT|TEXTAREA|SELECT/.test(ae.tagName) ||
      ae.isContentEditable || (ae.closest && ae.closest('[contenteditable="true"]') !== null);
    // the new-project wizard takes no app shortcuts; Escape closes it
    if (wz) {
      if (e.key !== 'Escape') return;
      if (!popEl.hidden) { closePopover(); return; }
      if (!modalHost.hidden) { closeModal(); return; }
      if (!calPop.hidden) { closeCal(); return; }
      if (inField) { document.activeElement.blur(); return; }
      wizardTryClose();
      return;
    }
    // Esc/Delete while drawing a dependency line cancels the add
    if (drag && drag.kind === 'port' && (e.key === 'Escape' || e.key === 'Delete' || e.key === 'Backspace')) {
      e.preventDefault();
      cancelPortDrag();
      return;
    }
    if (e.key === 'Escape') {
      if (!popEl.hidden) { closePopover(); return; }
      if (!modalHost.hidden) { closeModal(); return; }
      if (inField) { document.activeElement.blur(); return; }
      if (selectedEdge) { selectedEdge = null; requestAnimationFrame(renderArrows); return; }
      if (multiSel) { multiSel = null; render(); return; } // collapse to the anchor first
      if (selectedId) { select(null); return; }
      return;
    }
    // on the start page only Escape (above, for its modals) applies
    if (document.body.classList.contains('start')) return;
    if ((e.metaKey || e.ctrlKey) && e.key.toLowerCase() === 'f' &&
        (view === 'planning' || view === 'scoping' || view === 'prio' || view === 'sprints')) {
      e.preventDefault();
      var ff = view === 'prio' ? $('#prFilter') : view === 'sprints' ? $('#spFilter') : $('#rowFilter');
      if (ff) { ff.focus(); ff.select(); }
      return;
    }
    if (inField) return;
    var mod = e.metaKey || e.ctrlKey;
    // [ and ] fold the left pane / the right panel (never while typing — above)
    if (!mod && !e.altKey && e.key === ']' && (view === 'planning' || view === 'scoping' || view === 'sprints')) {
      e.preventDefault(); togglePanel(); return;
    }
    if (!mod && !e.altKey && e.key === '[' && (view === 'planning' || view === 'scoping' || view === 'budget')) {
      e.preventDefault(); toggleLeftPane(); return;
    }
    if (mod && e.key.toLowerCase() === 'z') {
      e.preventDefault();
      if (e.shiftKey) redo(); else undo();
      return;
    }
    if (mod && e.key.toLowerCase() === 'y') { e.preventDefault(); redo(); return; }
    // ⌘+ / ⌘- (Ctrl on other platforms, numpad too) zoom the Planning
    // timeline instead of the page; never while a dialog is open
    if (mod && !e.altKey && view === 'planning' && modalHost.hidden) {
      var zIn = e.key === '=' || e.key === '+' || e.code === 'NumpadAdd';
      var zOut = e.key === '-' || e.key === '_' || e.code === 'NumpadSubtract';
      if (zIn || zOut) { e.preventDefault(); zoomBy(zIn ? 1.2 : 1 / 1.2); return; }
    }
    if (mod && e.key.toLowerCase() === 'a' && (view === 'planning' || view === 'scoping')) {
      e.preventDefault();
      var all = $$('#rows .row.item').map(function (r) { return r.dataset.id; });
      if (all.length) {
        flushPanelEdit();
        if (!selectedId || all.indexOf(selectedId) === -1) selectedId = all[0];
        multiSel = all;
        selStory = null;
        render();
      }
      return;
    }
    if (mod && e.key.toLowerCase() === 's') {
      e.preventDefault(); $('#btnSave').click(); return;
    }
    if ((e.key === 'Delete' || e.key === 'Backspace') && selectedEdge) {
      e.preventDefault();
      deleteSelectedEdge();
      return;
    }
    if ((e.key === 'Delete' || e.key === 'Backspace') && selectedId) {
      e.preventDefault();
      var del = selIds();
      if (del.length > 1) {
        confirmBox('Delete ' + del.length + ' items?', 'Every selected item is removed.', 'Delete', function () {
          commit('delete', function (s) {
            s.items = s.items.filter(function (x) { return del.indexOf(x.id) === -1; });
            selectedId = null;
            multiSel = null;
          });
        }, true);
        return;
      }
      var it = RM.itemById(state, selectedId);
      if (it) {
        confirmBox('Delete #' + it.num + '?', esc(it.feature), 'Delete', function () {
          commit('delete', function (s) {
            s.items = s.items.filter(function (x) { return x.id !== selectedId; });
            selectedId = null;
          });
        }, true);
      }
      return;
    }
    if ((e.key === 'ArrowLeft' || e.key === 'ArrowRight') && selectedId) {
      var nudgeIds = selIds().filter(function (id) {
        var t = RM.itemById(state, id);
        return t && isScheduled(t) && !t.locked;
      });
      if (nudgeIds.length) {
        e.preventDefault();
        var delta = (e.key === 'ArrowRight' ? 1 : -1) * (e.shiftKey ? 5 : 1);
        commit('nudge', function (s) {
          nudgeIds.forEach(function (id) {
            var t = RM.itemById(s, id);
            var ns2 = Math.max(0, t.startDay + delta);
            RM.shiftStories(t, ns2 - t.startDay);
            t.startDay = ns2;
          });
        });
      }
    }
  });

  // dirty indicator between commit & saveLocal isn't observable (sync), but flag edits in inputs
  document.addEventListener('input', function (e) {
    if (e.target.closest && (e.target.closest('#panel') || e.target.id === 'docTitle')) {
      var st = $('#saveStatus');
      if (st) {
        st.textContent = 'editing…';
        st.classList.add('dirty');
      }
    }
  });

  window.addEventListener('resize', function () { requestAnimationFrame(renderArrows); });

  // ------------------------------------------------------------ png export
  function exportModal() {
    var meta = state.meta;
    var sprints = [];
    for (var w = 0; w < meta.numWeeks; w++) {
      var n = RM.sprintNumForWeek(meta, w);
      if (!sprints.length || sprints[sprints.length - 1].num !== n)
        sprints.push({ num: n, date: RM.fmtShort(RM.dayToDate(meta, w * SPW())) });
    }
    function distinct(field) {
      var seen = {}, out = [];
      state.items.forEach(function (it) {
        var v = it[field];
        if (v && !seen[v]) { seen[v] = true; out.push(v); }
      });
      return out;
    }
    function sprOpts(selNum) {
      var withS = RM.sprintsEnabled(meta);
      return sprints.map(function (s) {
        return '<option value="' + s.num + '"' + (s.num === selNum ? ' selected' : '') +
          '>' + (withS ? 'S' + s.num + ' · ' : 'Week of ') + esc(s.date) + '</option>';
      }).join('');
    }
    function nameOpts(values, allLabel) {
      return '<option value="">' + allLabel + '</option>' + values.map(function (v) {
        return '<option value="' + esc(v) + '">' + esc(v) + '</option>';
      }).join('');
    }
    // the last-used settings come back; row grouping falls back to however
    // the timeline is grouped right now
    var pref = exportPrefs || {};
    var pWs = pref.groupWs != null ? !!pref.groupWs : groupWs;
    var pEpic = pref.groupEpic != null ? !!pref.groupEpic : groupEpic;
    function ck(on) { return on ? ' checked' : ''; }
    openModal(
      '<div class="modal" style="width:440px">' +
      '<div class="m-head"><h2>Export</h2>' +
      '<button class="p-close" data-m="x"><i data-lucide="x"></i></button></div>' +
      '<div class="m-body">' +
      '<div class="m-sec"><label>Format</label><div class="p-row">' +
      '<label class="p-check"><input type="radio" name="exFmt" id="exFmtPng"' + ck(pref.fmt !== 'pptx' && pref.fmt !== 'jira' && pref.fmt !== 'html') + '> PNG image</label>' +
      '<label class="p-check"><input type="radio" name="exFmt" id="exFmtPptx"' + ck(pref.fmt === 'pptx') + '> PowerPoint (editable)</label>' +
      '<label class="p-check"><input type="radio" name="exFmt" id="exFmtJira"' + ck(pref.fmt === 'jira') + '> Jira CSV</label>' +
      '<label class="p-check"><input type="radio" name="exFmt" id="exFmtHtml"' + ck(pref.fmt === 'html') + '> Standalone HTML (view-only)</label>' +
      '</div></div>' +
      '<div id="exHtml" class="m-sec"' + (pref.fmt === 'html' ? '' : ' hidden') + '><div class="m-hint">One self-contained .html file: every tab is there to browse, nothing can be edited, and Setup keeps only the theme. Open it in any browser \u2014 no install, no server.</div></div>' +
      '<div id="exTimeline"' + (pref.fmt === 'jira' || pref.fmt === 'html' ? ' hidden' : '') + '>' +
      '<div class="m-sec"><label>Date range</label><div class="p-grid2">' +
      '<div><label class="p-lab">From</label><select id="exFrom" style="width:100%">' + sprOpts(sprints[0].num) + '</select></div>' +
      '<div><label class="p-lab">To</label><select id="exTo" style="width:100%">' + sprOpts(sprints[sprints.length - 1].num) + '</select></div>' +
      '</div></div>' +
      '<div class="m-sec"><label>Filter</label><div class="p-grid2">' +
      '<div><label class="p-lab">Workstream</label><select id="exWs" style="width:100%">' + nameOpts(distinct('workstream'), 'All workstreams') + '</select></div>' +
      '<div><label class="p-lab">Epic</label><select id="exEpic" style="width:100%">' + nameOpts(distinct('epic'), 'All epics') + '</select></div>' +
      '</div>' +
      '<div style="margin-top:8px"><label class="p-lab">Phase</label><select id="exPhase" style="width:100%">' +
      '<option value="">All phases</option>' + state.phases.map(function (p) {
        return '<option value="' + esc(p.id) + '">' + esc(p.name) + '</option>';
      }).join('') + '</select></div></div>' +
      '<div class="m-sec"><label>Group rows</label><div class="p-row">' +
      (state.meta.workstreamsEnabled ?
        '<label class="p-check"><input type="checkbox" id="exGroupWs"' + ck(pWs) + '> By workstream</label>' : '') +
      '<label class="p-check"><input type="checkbox" id="exGroupEpic"' + ck(pEpic) + '> By epic</label>' +
      '</div></div>' +
      '<div class="m-sec"><label>Slides / files</label><div class="p-row">' +
      '<label class="p-check"><input type="checkbox" id="exSplitPhase"' + ck(pref.byPhase) + '> One per phase</label>' +
      (state.meta.workstreamsEnabled ?
        '<label class="p-check"><input type="checkbox" id="exSplitWs"' + ck(pref.byWs) + '> One per workstream</label>' : '') +
      '</div><div class="m-hint">Unchecked → everything on a single image or slide. PPTX keeps all slides in one file; a split PNG export saves one image per slice.</div></div>' +
      '<div class="m-sec"><label>Options</label><div class="p-row">' +
      '<select id="exScale"><option value="1"' + (pref.scale === 1 ? ' selected' : '') + '>1× scale</option>' +
      '<option value="2"' + (pref.scale !== 1 ? ' selected' : '') + '>2× scale</option></select>' +
      '<label class="p-check"><input type="checkbox" id="exArrows"' + ck(pref.arrows) + '> Dependency arrows</label>' +
      '</div><div class="m-hint">The exported dates are bounded by the range AND the filtered bars — empty weeks at either end are trimmed.</div></div>' +
      '</div>' +
      '<div id="exJira"' + (pref.fmt === 'jira' ? '' : ' hidden') + '>' + jiraOptionsHtml() + '</div>' +
      '</div>' +
      '<div class="m-foot"><button data-m="cancel">Cancel</button>' +
      '<button id="exGo" class="primary"><i data-lucide="image"></i>Export</button></div></div>',
      function (host) {
        $('[data-m=x]', host).onclick = closeModal;
        $('[data-m=cancel]', host).onclick = closeModal;
        // splitting the export by workstream hides the per-workstream grouping
        // (it would be a single-band no-op inside each slice)
        var splitWsChk = $('#exSplitWs', host), groupWsChk = $('#exGroupWs', host);
        if (splitWsChk && groupWsChk) splitWsChk.addEventListener('change', function () {
          groupWsChk.disabled = splitWsChk.checked;
        });
        // pixel scale only means something for the PNG raster; the Jira CSV
        // has its own options and none of the timeline ones
        function syncFmt() {
          var jira = $('#exFmtJira', host).checked;
          var htmlF = $('#exFmtHtml', host).checked;
          $('#exTimeline', host).hidden = jira || htmlF;
          $('#exJira', host).hidden = !jira;
          $('#exHtml', host).hidden = !htmlF;
          $('#exScale', host).disabled = $('#exFmtPptx', host).checked;
        }
        ['exFmtPng', 'exFmtPptx', 'exFmtJira', 'exFmtHtml'].forEach(function (id) {
          $('#' + id, host).addEventListener('change', syncFmt);
        });
        // restored settings get the same dependent-control states
        if (groupWsChk && splitWsChk) groupWsChk.disabled = splitWsChk.checked;
        syncFmt();
        $('#exGo', host).onclick = function () {
          if ($('#exFmtHtml', host).checked) {
            exportPrefs = Object.assign({}, exportPrefs || {}, { fmt: 'html' });
            saveLocal();
            closeModal();
            exportHtmlTo();
            return;
          }
          if ($('#exFmtJira', host).checked) {
            var jo = jiraOptionsOf(host);
            if (!jo) return;
            jiraPrefs = jo;
            exportPrefs = Object.assign({}, exportPrefs || {}, { fmt: 'jira' });
            saveLocal();
            closeModal();
            exportJiraTo(jo);
            return;
          }
          var exOpts = {
            fromSprint: parseInt($('#exFrom', host).value, 10),
            toSprint: parseInt($('#exTo', host).value, 10),
            ws: $('#exWs', host).value || null,
            epic: $('#exEpic', host).value || null,
            phaseId: $('#exPhase', host).value || null,
            scale: parseInt($('#exScale', host).value, 10),
            arrows: $('#exArrows', host).checked,
            groupWs: !!(groupWsChk && groupWsChk.checked && !(splitWsChk && splitWsChk.checked)),
            groupEpic: $('#exGroupEpic', host).checked,
            byPhase: $('#exSplitPhase', host).checked,
            byWs: !!(splitWsChk && splitWsChk.checked)
          };
          var fmt = $('#exFmtPptx', host).checked ? 'pptx' : 'png';
          // remember the settings for next time (persisted in the ui snapshot)
          exportPrefs = {
            fmt: fmt, scale: exOpts.scale, arrows: exOpts.arrows,
            groupWs: !!(groupWsChk && groupWsChk.checked), groupEpic: exOpts.groupEpic,
            byPhase: exOpts.byPhase, byWs: exOpts.byWs
          };
          saveLocal();
          closeModal();
          if (fmt === 'pptx') exportPptxTo(exOpts);
          else exportPngTo(exOpts);
        };
      });
  }
  $('#btnExport').addEventListener('click', exportModal);

  // Jira CSV: a file for Jira Cloud's user-level importer (work navigator →
  // Import issues from CSV). Rows parent to the Jira keys typed into
  // Headway; see export-jira.js. Lives in the Export dialog as a format.
  function jiraOptionsHtml() {
    var pref = jiraPrefs || {};
    var d = RM_JIRA.DEFAULTS;
    function ck(on) { return on ? ' checked' : ''; }
    return '<div class="m-sec"><label>Rows</label><div class="p-row">' +
      '<label class="p-check"><input type="checkbox" id="jxFeatures"' + ck(pref.features != null ? pref.features : d.features) + '> Features</label>' +
      '<label class="p-check"><input type="checkbox" id="jxStories"' + ck(pref.stories != null ? pref.stories : d.stories) + '> Stories</label>' +
      '</div><div class="m-hint">Without story rows, a feature\'s stories become a checklist in its description.</div>' +
      '<div class="m-hint">Issue types follow each item\'s type (Setup → Organization).</div></div>' +
      '<div class="m-sec"><label>In Jira</label><div class="m-hint">' +
      'Work navigator → ⋯ → Import issues from CSV (needs the Create work items and Make bulk changes permissions). ' +
      'Choose the date format <b>yyyy-MM-dd</b>. Parent and Blocked By carry the Jira keys entered in Headway; ' +
      'rows with a Jira Key column value can be mapped as updates.</div></div>';
  }
  // reads the Jira options out of the dialog; false when nothing is selected
  function jiraOptionsOf(host) {
    var opts = {
      features: $('#jxFeatures', host).checked,
      stories: $('#jxStories', host).checked,
    };
    if (!opts.features && !opts.stories) { toast('Pick features, stories or both'); return false; }
    return opts;
  }
  function exportJiraTo(opts) {
    var csv = RM_JIRA.csv(state, opts);
    saveExport({ blob: new Blob([csv], { type: 'text/csv' }), name: safeName(RM_JIRA.fileName(state)) }, CSV_KIND);
  }
  var CSV_KIND = { desc: 'CSV file', mime: 'text/csv', ext: 'csv' };
  var exportSink = null; // tests intercept exports here instead of the save path

  // Where an exported file lands: desktop asks for a destination via the
  // native save dialog and OPENS the file afterwards; browsers with a
  // save-file picker let the user choose the folder; anything else falls
  // back to a download. `kind` = { desc, mime, ext }.
  function saveExport(r, kind) {
    if (exportSink) return Promise.resolve(exportSink(r, kind));
    if (window.HeadwayDesktop && window.HeadwayDesktop.saveFileAndOpen) {
      return window.HeadwayDesktop.saveFileAndOpen(r.blob, r.name, kind.desc, kind.ext).then(function (p) {
        if (p) toast('Exported ' + p.replace(/^.*[\\/]/, ''));
      });
    }
    if (window.showSaveFilePicker) {
      var accept = {};
      accept[kind.mime] = ['.' + kind.ext];
      return window.showSaveFilePicker({
        suggestedName: r.name,
        types: [{ description: kind.desc, accept: accept }]
      }).then(function (handle) {
        return handle.createWritable().then(function (w) {
          return w.write(r.blob).then(function () { return w.close(); });
        }).then(function () { toast('Exported ' + handle.name); });
      }).catch(function (err) {
        if (err && err.name === 'AbortError') return; // user canceled
        throw err;
      });
    }
    var a = document.createElement('a');
    a.href = URL.createObjectURL(r.blob);
    a.download = r.name;
    a.click();
    setTimeout(function () { URL.revokeObjectURL(a.href); }, 5000);
    toast('Exported ' + r.name);
  }
  var PNG_KIND = { desc: 'PNG image', mime: 'image/png', ext: 'png' };
  var PPTX_KIND = {
    desc: 'PowerPoint presentation',
    mime: 'application/vnd.openxmlformats-officedocument.presentationml.presentation',
    ext: 'pptx'
  };
  function safeName(s) { return String(s).replace(/[\\/:*?"<>|]+/g, '').trim(); }

  // ---- standalone HTML: the app's own page with every stylesheet and
  // script inlined, the document and view state embedded, and the
  // desktop bridge + the Excel/PowerPoint engines left out (nothing to
  // save or export from a view-only copy). Opens from disk in any browser.
  var HTML_KIND = { desc: 'Standalone HTML page', mime: 'text/html', ext: 'html' };
  var HTML_BUNDLE_FILES = ['css/app.css', 'js/vendor/lucide.min.js', 'js/core.js', 'js/excel.js', 'js/export-png.js',
    'js/export-pptx.js', 'js/export-jira.js', 'js/jira.js', 'js/ai.js', 'js/app.js'];
  // a closing script tag inside inlined source would end the block early
  function inlineSafe(src) { return String(src).replace(/<\/(script)/gi, '<\\/$1'); }
  function buildStandaloneHtml(indexHtml, files, doc, ui) {
    var out = indexHtml;
    out = out.replace('<link rel="stylesheet" href="css/app.css">', function () {
      return '<style>' + inlineSafe(files['css/app.css'] || '').replace(/<\/style/gi, '<\\/style') + '</style>';
    });
    var payload = '<script>window.HEADWAY_VIEW = ' + inlineSafe(JSON.stringify({ doc: doc, ui: ui })) + ';</script>';
    out = out.replace(/<script src="([^"]+)"><\/script>/g, function (m0, src) {
      if (src === 'js/desktop.js' || src === 'js/vendor/exceljs.min.js' || src === 'js/vendor/pptxgen.bundle.js') return '';
      if (files[src] == null) return '';
      return (src === 'js/core.js' ? payload + '\n' : '') + '<script>' + inlineSafe(files[src]) + '</script>';
    });
    out = out.replace(/<title>[^<]*<\/title>/, function () { return '<title>' + esc((doc.meta && doc.meta.title) || 'Roadmap') + '</title>'; });
    return out;
  }
  function fetchText(path) {
    return fetch(path, { cache: 'no-store' }).then(function (r) {
      if (!r.ok) throw new Error('could not read ' + path);
      return r.text();
    });
  }
  function exportHtmlTo() {
    if (typeof fetch !== 'function') { toast('Standalone HTML needs a browser that can read the app files', 'err'); return; }
    var names = ['index.html'].concat(HTML_BUNDLE_FILES);
    Promise.all(names.map(fetchText)).then(function (texts) {
      var files = {};
      names.forEach(function (n, i) { files[n] = texts[i]; });
      var ui = uiSnapshot();
      ui.panelOpen = false;
      var html = buildStandaloneHtml(files['index.html'], files, RM.clone(state), ui);
      var blob = new Blob([html], { type: 'text/html' });
      return saveExport({ blob: blob, name: (safeName(state.meta.title) || 'Roadmap') + '.html' }, HTML_KIND);
    }).catch(function (err) { toast('Export failed: ' + (err && err.message || err), 'err'); });
  }

  function exportPngTo(exOpts) {
    var entries = RM_EXPORT.plan(state, exOpts);
    if (!entries.length) { toast('Nothing to export in this selection', 'err'); return; }
    var split = exOpts.byPhase || exOpts.byWs;
    if (!split) {
      RM_EXPORT.toBlob(state, entries[0].opts)
        .then(function (r) { return saveExport(r, PNG_KIND); })
        .catch(function (err) { toast('Export failed: ' + (err && err.message || err), 'err'); });
      return;
    }
    var docName = safeName(state.meta.title || '') || 'Roadmap';
    Promise.all(entries.map(function (e) {
      return RM_EXPORT.toBlob(state, e.opts).then(function (r) {
        return { blob: r.blob, name: docName + ' — ' + safeName(e.name) + '.png' };
      });
    })).then(function (files) {
      if (window.HeadwayDesktop && window.HeadwayDesktop.saveManyToFolder) {
        return window.HeadwayDesktop.saveManyToFolder(files).then(function (dir) {
          if (dir) toast('Exported ' + files.length + ' images to ' + dir.replace(/^.*[\\/]/, ''));
        });
      }
      // plain browser: a download per file, spaced so none get dropped
      files.forEach(function (f, i) {
        setTimeout(function () {
          var a = document.createElement('a');
          a.href = URL.createObjectURL(f.blob);
          a.download = f.name;
          a.click();
          setTimeout(function () { URL.revokeObjectURL(a.href); }, 5000);
        }, i * 350);
      });
      toast('Exporting ' + files.length + ' images');
    }).catch(function (err) {
      toast('Export failed: ' + (err && err.message || err), 'err');
    });
  }

  function exportPptxTo(exOpts) {
    RM_PPTX.toBlob(state, exOpts)
      .then(function (r) { return saveExport(r, PPTX_KIND); })
      .catch(function (err) { toast('Export failed: ' + (err && err.message || err), 'err'); });
  }

  // ------------------------------------------------------------ boot
  function boot() {
    if (readOnly) {
      applyUi(window.HEADWAY_VIEW.ui || null);
      state = RM.normalizeState(RM.clone(window.HEADWAY_VIEW.doc));
      document.body.classList.add('ro');
      setupTab = 'appearance';
    } else {
      state = loadLocal() || blankState(); // fresh installs start completely empty
    }
    validation = RM.validate(state);
    // features with stories start expanded so the story timelines are visible
    // (unless a saved expansion map — local or from an imported file — says otherwise)
    if (!uiExpandedLoaded) state.items.forEach(function (it) { if (it.stories.length) expanded[it.id] = true; });
    render();
    capTypeToast(localCapTypeChanges); // a restored local copy migrated people's types
    // fresh launches land on the start page; a mid-session reload (the
    // sessionStorage flag survives those, not app restarts) rejoins the editor
    var inEditor = readOnly;
    try { inEditor = inEditor || sessionStorage.getItem(SESSION_KEY) === '1'; } catch (e) { /* storage optional */ }
    if (inEditor) enterEditor();
    else showStart();
  }
  boot();

  // console/debug handle (read-only snapshot)
  // the exact project title, minus filesystem-hostile characters only —
  // no lowercasing, no hyphenation, no date
  function saveFileName() {
    return ((state.meta.title || '').replace(/[\\/:*?"<>|]+/g, '').trim() || 'Roadmap') + '.xlsx';
  }

  window.__headway = {
    getState: function () { return RM.clone(state); },
    saveFileName: saveFileName,
    getValidation: function () { return validation; },
    templateState: templateState,
    releaseNotesFor: releaseNotesFor,
    maybeShowReleaseNotes: maybeShowReleaseNotes,
    openReleaseNotes: openReleaseNotes,
    setExportSink: function (fn) { exportSink = typeof fn === 'function' ? fn : null; },
    buildStandaloneHtml: buildStandaloneHtml,
    setItemType: setItemType,
    selectItem: select
  };
})();
