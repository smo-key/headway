/* Headway UI. Vanilla JS, no build step; state logic lives in core.js. */
(function () {
  'use strict';

  var LS_KEY = 'headway-v1';
  var UI_KEY = 'headway-ui-v1';

  // ------------------------------------------------------------ app state
  var state;                 // document state (undo-tracked)
  var validation;            // RM.validate cache
  var undoStack = [], redoStack = [];
  var selectedId = null;
  var expanded = {};         // itemId -> true (stories open)
  var weekPx = 28;
  var view = 'planning';     // planning (timeline) | scoping (spreadsheet)
  var depsMode = 'on';       // on: selected item's explicit deps + violations + critical path | none
  var showCrit = true;       // orange critical-path highlight (bars + arrows)
  var showCap = true;        // weekly capacity row in the planning header
  var presentMode = false;   // timeline-only preview (hides topbar + left pane); transient
  var repCollapsed = true;   // bottom Reports drawer starts tucked away
  var leftWBudget = 696;     // frozen left-pane width, budgeting view
  var repMode = 'workstream'; // reports grouping: workstream | phase | phase-ws
  var panelW = 372;          // right edit-panel width (resizable)
  var leftWPlan = 538;       // frozen left-pane width, planning view
  var leftWScope = 538;      // …and scoping view (independently resizable)
  var groupWs = false;       // sub-group rows by workstream inside each phase
  var groupEpic = false;     // …and/or by epic (nested under workstream)
  var snapDays = 5;          // drag/resize snap: 1 (day) | 5 (week) | 10 (2 weeks)
  var autoOrder = true;      // after move/resize, reorder rows by start day (stable)
  var capType = 'Development'; // work type the header capacity row counts ('' = all)
  var filterText = '';       // planning/scoping row filter (⌘F); transient
  var docSaved = false;      // doc matches its last save/open (Save button shows ✓)
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
  function dayPx() { return weekPx / 5; }
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
    return { weekPx: weekPx, view: view, depsMode: depsMode, groupWs: groupWs, groupEpic: groupEpic, resCollapsed: resCollapsed, snapDays: snapDays, autoOrder: autoOrder, showCrit: showCrit, showCap: showCap, scopeColW: scopeColW, capType: capType, resPanelH: resPanelH, panelSec: panelSec, leftWPlan: leftWPlan, leftWScope: leftWScope, leftWBudget: leftWBudget, panelW: panelW, expanded: expanded, repCollapsed: repCollapsed, repMode: repMode, autoSave: autoSave };
  }
  var uiExpandedLoaded = false; // boot skips the auto-expand default when true

  function applyUi(ui) {
    if (!ui) return;
    weekPx = ui.weekPx || 28;
    view = ['scoping', 'setup', 'budget'].indexOf(ui.view) !== -1 ? ui.view : 'planning';
    depsMode = ui.depsMode === 'none' ? 'none' : 'on';
    groupWs = ui.groupWs != null ? !!ui.groupWs : ui.groupBy === 'ws';
    groupEpic = ui.groupEpic != null ? !!ui.groupEpic : (ui.groupBy === 'epic' || !!ui.groupByEpic);
    resCollapsed = !!ui.resCollapsed;
    snapDays = [1, 5, 10].indexOf(ui.snapDays) !== -1 ? ui.snapDays : 5;
    scopeColW = ui.scopeColW && typeof ui.scopeColW === 'object' ? ui.scopeColW : {};
    autoOrder = ui.autoOrder !== false; // default true
    showCrit = ui.showCrit !== false;   // default true
    showCap = ui.showCap !== false;     // default true
    capType = ui.capType != null ? ui.capType : 'Development';
    resPanelH = ui.resPanelH > 40 ? ui.resPanelH : 150;
    panelSec = ui.panelSec && typeof ui.panelSec === 'object' ? ui.panelSec : {};
    leftWPlan = ui.leftWPlan > 200 ? ui.leftWPlan : 538;
    leftWScope = ui.leftWScope > 200 ? ui.leftWScope : 538;
    // 660 was the pre-widened-columns default; bump those snapshots to 696
    leftWBudget = ui.leftWBudget > 300 && ui.leftWBudget !== 660 ? ui.leftWBudget : 696;
    panelW = ui.panelW > 280 ? ui.panelW : 372;
    repCollapsed = ui.repCollapsed !== false; // default collapsed
    repMode = ['workstream', 'phase', 'phase-ws'].indexOf(ui.repMode) !== -1 ? ui.repMode : 'workstream';
    autoSave = ui.autoSave !== false;   // default true (desktop writes to the open file)
    if (ui.expanded && typeof ui.expanded === 'object') {
      expanded = ui.expanded;
      uiExpandedLoaded = true;
    }
  }

  var localSaveBroken = false;
  function saveLocal() {
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

  function loadLocal() {
    try {
      applyUi(JSON.parse(localStorage.getItem(UI_KEY) || 'null'));
      var raw = JSON.parse(localStorage.getItem(LS_KEY) || 'null');
      if (raw && raw.items) return RM.normalizeState(raw);
    } catch (e) { /* corrupted local copy — fall back to seed */ }
    return null;
  }

  // ------------------------------------------------------------ commits
  function commit(label, mutate) {
    undoStack.push(JSON.stringify(state));
    if (undoStack.length > 120) undoStack.shift();
    redoStack.length = 0;
    if (mutate) mutate(state);
    afterChange();
  }
  function replaceState(label, next) {
    undoStack.push(JSON.stringify(state));
    if (undoStack.length > 120) undoStack.shift();
    redoStack.length = 0;
    state = next;
    afterChange();
  }
  function undo() {
    if (!undoStack.length) return;
    redoStack.push(JSON.stringify(state));
    state = JSON.parse(undoStack.pop());
    afterChange();
  }
  function redo() {
    if (!redoStack.length) return;
    undoStack.push(JSON.stringify(state));
    state = JSON.parse(redoStack.pop());
    afterChange();
  }
  function afterChange() {
    docSaved = false;
    validation = RM.validate(state);
    saveLocal();
    render();
    scheduleAutoSave();
  }

  var autoSaveTimer = null;
  function scheduleAutoSave() {
    // desktop only, and only once the doc lives in a real file
    if (!autoSave || !window.HeadwayDesktop || !HeadwayDesktop.currentPath()) return;
    clearTimeout(autoSaveTimer);
    autoSaveTimer = setTimeout(function () {
      if (!docSaved && !savingNow) doSave(false, true);
    }, 1500);
  }

  function matchesFilter(it) {
    if (!filterText) return true;
    var q = filterText.toLowerCase();
    return (it.feature || '').toLowerCase().indexOf(q) !== -1 ||
      (it.epic || '').toLowerCase().indexOf(q) !== -1 ||
      (it.workstream || '').toLowerCase().indexOf(q) !== -1 ||
      (it.stories || []).some(function (st) {
        return (st.title || '').toLowerCase().indexOf(q) !== -1;
      });
  }

  // ------------------------------------------------------------ toasts, popover, modal
  function toast(msg, kind) {
    var el = document.createElement('div');
    el.className = 'toast' + (kind === 'err' ? ' err' : '');
    el.textContent = msg;
    $('#toasts').appendChild(el);
    setTimeout(function () { el.style.opacity = '0'; el.style.transition = 'opacity .3s'; }, 3400);
    setTimeout(function () { el.remove(); }, 3800);
  }

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
    if (!popEl.hidden && !popEl.contains(e.target)) closePopover();
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
      sanitizeHtml(html) + '</div></div>';
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

  // ------------------------------------------------------------ render
  // scoping view columns come from the document (meta.scopeCols — orderable,
  // removable, plus custom ones); widths are user-resizable and remembered in
  // the browser (UI_KEY)
  var SCOPE_DEFAULT_W = { description: 320, enables: 260, outOfScope: 260, extDeps: 240, notes: 300 };
  // fixed lead columns (field chips, not text): [key, label, width, min width]
  var SCOPE_FIXED = [
    ['size', 'Size', 56, 44],
    ['risk', 'Risk', 56, 44],
    ['duration', 'Weeks', 62, 48],
    ['workstream', 'Workstream', 130, 80],
    ['epic', 'Epic', 150, 80]
  ];
  function allScopeCols() { return SCOPE_FIXED.concat(scopeCols()); }
  function scopeCols() {
    return state.meta.scopeCols.map(function (c) {
      return [c.key, RM.scopeColLabel(c), SCOPE_DEFAULT_W[c.key] || 240];
    });
  }
  var scopeColW = {}; // field -> px override

  // panel sections: key -> collapsed override (persisted in UI_KEY);
  // unlisted keys fall back to the defaults below
  var panelSec = {};
  var PANEL_SEC_CLOSED = { stories: 1, scope: 1 };
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
    var sx = board.scrollLeft, sy = board.scrollTop;
    critCache = RM.criticalPath(state);
    document.documentElement.style.setProperty('--week-px', weekPx + 'px');
    if (presentMode && view === 'setup') { presentMode = false; $('#btnPresentExit').hidden = true; }
    document.body.classList.toggle('no-cap', !showCap || !state.meta.capacityEnabled);
    document.body.classList.toggle('present', presentMode);
    document.body.classList.toggle('cap-off', !state.meta.capacityEnabled);
    document.documentElement.style.setProperty('--left-w',
      (presentMode && view === 'planning' ? 0 : (view === 'scoping' ? leftWScope : view === 'budget' ? leftWBudget : leftWPlan)) + 'px');
    document.documentElement.style.setProperty('--panel-w', panelW + 'px');
    document.body.dataset.view = view;

    renderTopbar();
    if (view === 'setup') {
      renderSetup();
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
      renderReports();
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
      $('#hdrCap').innerHTML = '';
      $('#capTypeCell').innerHTML = '';
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
    syncHdrH();
    board.scrollLeft = sx; board.scrollTop = sy;
    requestAnimationFrame(function () { renderArrows(); positionToday(); });
  }

  // ------------------------------------------------------------ reports drawer
  var REP_MODES = [['workstream', 'By workstream'], ['phase', 'By phase'], ['phase-ws', 'By phase × workstream']];
  function repModeLabel() {
    var out = 'By workstream';
    REP_MODES.forEach(function (m) { if (m[0] === repMode) out = m[1]; });
    return out;
  }
  function renderReports() {
    var panel = $('#repPanel');
    if (!panel) return;
    panel.classList.toggle('collapsed', repCollapsed);
    $('#repChev').classList.toggle('open', !repCollapsed);
    // the mode dropdown stays even while collapsed so the header keeps its height
    $('#repModeCell').innerHTML = ddButton('repmode', esc(repModeLabel()), null, 'Group the report');
    if (repCollapsed) { $('#repBody').innerHTML = ''; return; }
    var rep = RM.costReport(state, repMode);
    var wsMode = repMode === 'workstream';
    var head = '<tr><th class="bu-name">' + (repMode === 'phase' ? 'Phase' : wsMode ? 'Workstream' : 'Phase · Workstream') + '</th>' +
      '<th class="bu-num">Items</th><th class="bu-num">Effort (wks)</th><th class="bu-num">Hours</th><th class="bu-num">Est. cost</th>' +
      (wsMode ? '<th class="bu-num">Role hours</th><th class="bu-num">Role cost</th>' : '') + '</tr>';
    function row(r, cls) {
      return '<tr' + (cls ? ' class="' + cls + '"' : '') + '><td class="bu-name">' + esc(r.key) + '</td>' +
        '<td class="bu-num">' + r.items + '</td>' +
        '<td class="bu-num">' + (Math.round(r.days / 5 * 10) / 10) + '</td>' +
        '<td class="bu-num">' + Math.round(r.hours) + '</td>' +
        '<td class="bu-num">' + fmtMoney(r.cost) + '</td>' +
        (wsMode ? '<td class="bu-num">' + Math.round(r.roleHours || 0) + '</td>' +
          '<td class="bu-num">' + fmtMoney(r.roleCost || 0) + '</td>' : '') + '</tr>';
    }
    $('#repBody').innerHTML =
      '<table class="bu-table rep-table"><thead>' + head + '</thead><tbody>' +
      rep.rows.map(function (r) { return row(r); }).join('') +
      row(rep.total, 'bu-total') + '</tbody></table>';
    if (window.lucide) lucide.createIcons();
  }

  $('#repHead').addEventListener('click', function (e) {
    if (e.target.closest('#repModeCell')) return;
    repCollapsed = !repCollapsed;
    saveLocal();
    renderReports();
  });
  $('#repModeCell').addEventListener('click', function (e) {
    var dd = e.target.closest('[data-dd="repmode"]');
    if (!dd) return;
    openDropdown(dd, REP_MODES.map(function (m) {
      return { label: m[1], checked: repMode === m[0], fn: function () {
        repMode = m[0];
        saveLocal();
        renderReports();
      } };
    }));
  });

  function renderScopeHeader() {
    var out = ['<div class="sc-hrow">'];
    SCOPE_FIXED.forEach(function (c) {
      out.push('<div class="sc-hcell sc-fixh" data-col="' + c[0] + '" style="width:' + scopeColWidth(c) + 'px">' +
        '<span class="sc-hlab">' + c[1] + '</span>' +
        '<span class="sc-rz" data-rz="' + c[0] + '" title="Drag to resize column"></span></div>');
    });
    scopeCols().forEach(function (c) {
      out.push('<div class="sc-hcell" data-col="' + c[0] + '" style="width:' + scopeColWidth(c) + 'px">' +
        '<span class="sc-hlab">' + esc(c[1]) + '</span>' +
        '<button class="sc-hmenu" data-colmenu="' + c[0] + '" title="Column options"><i data-lucide="ellipsis"></i></button>' +
        '<span class="sc-rz" data-rz="' + c[0] + '" title="Drag to resize column"></span></div>');
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
    $$('#viewTabs button').forEach(function (b) {
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
  }

  function renderHeader(laneW) {
    var meta = state.meta;
    var si = RM.sprintInfo(meta);
    var wps = si.wps;
    var sprintW = wps * weekPx;
    var hset = RM.holidayDaySet(meta);
    var hs = [], hw = [], hc = [];
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
      var numTag = cellW >= 56 ? '<span class="sp-num">S' + num + '</span>' : '';
      var hnS = 0;
      for (var wv = w0v; wv < w1v; wv++) hnS += RM.holidaysInWeek(meta, wv, hset);
      hs.push('<div class="sprint-cell" style="left:' + (w0v * weekPx) + 'px;width:' + cellW +
        'px" title="Sprint ' + num + ' — starts week of ' + esc(RM.fmtShortYear(d)) +
        (hnS ? ' · ' + hnS + ' holiday day' + (hnS > 1 ? 's' : '') : '') + '">' +
        '<span class="sp-date">' + dateTxt + '</span>' + numTag + '</div>');
    }

    // capacity row: ≈ parallel work items available (people-equivalents of
    // the selected work type), colored by demand
    var cap = validation.capacity;
    for (var w = 0; w < meta.numWeeks; w++) {
      var cell = cap.weeks[w];
      var avail = capType ? (cell.capByType[capType] || 0) : cell.cap;
      var demand = capType ? (cell.demandByType[capType] || 0) : cell.demand;
      var hn = RM.holidaysInWeek(meta, w, hset);
      var cls, txt2 = '', title;
      if (cell.blackout) { cls = 'blackout'; txt2 = weekPx >= 24 ? '✕' : ''; title = 'Holiday week'; }
      else if (cap.teamTotal === 0) {
        cls = demand > 0 ? 'ok' : 'idle';
        txt2 = weekPx >= 22 && demand ? fmtPe(demand) : '';
        title = fmtPe(demand) + ' parallel item(s) needed (no roster yet)';
      } else if (avail === Infinity) {
        cls = 'idle'; title = 'No roster';
      } else {
        var over = demand > avail + 1e-9;
        cls = over ? 'over' : (avail > 0 && demand / avail > 0.85 ? 'mid' : (demand === 0 ? 'idle' : 'ok'));
        txt2 = weekPx >= 26 ? (fmtPe(demand) + '/' + fmtPe(avail)) : (weekPx >= 20 ? fmtPe(avail) : '');
        title = fmtPe(demand) + ' needed / ' + fmtPe(avail) + ' parallel ' + (capType || 'any-type') + ' capacity (hours ÷ 40)';
      }
      hc.push('<div class="cap-cell ' + cls + (hn && !cell.blackout ? ' part' : '') + '" tabindex="0" data-w="' + w +
        '" style="left:' + (w * weekPx + 1) + 'px;width:' + (weekPx - 2) + 'px" title="' +
        esc('Week of ' + RM.fmtShort(RM.weekStartDate(meta, w)) + ': ' + title +
          (hn ? ' · ' + hn + ' holiday day(s)' : '') + ' · click to toggle holiday week') + '">' + txt2 + '</div>');
    }
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
    phLine.style.height = (phLanes.length ? phLanes.length * PH_H + 4 : 22) + 'px';
    phLine.style.display = '';
    $('#hdrPhases').innerHTML = phCells.join('');
    $('#hdrPhases').style.width = laneW + 'px';

    $('#hdrSprints').innerHTML = hs.join('');
    $('#hdrSprints').style.width = laneW + 'px';
    $('#hdrCap').innerHTML = hc.join('');
    $('#capTypeCell').innerHTML = ddButton('captype',
      'capacity: ' + (capType ? esc(capType) : 'all types'), null, 'Which work type the capacity row counts');
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
    var lo = drag.lo0, hi = drag.hi0;
    if (drag.mode === 'move') {
      lo = Math.max(0, snapTo(drag.lo0 + dd));
      hi = lo + (drag.hi0 - drag.lo0);
    } else if (drag.mode === 'resize-l') {
      lo = Math.max(0, Math.min(snapTo(drag.lo0 + dd), hi - 1));
    } else {
      hi = Math.max(lo + 1, snapTo(drag.hi0 + dd));
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

  // capacity-row work-type filter dropdown
  $('#capTypeCell').addEventListener('click', function (e) {
    var dd = e.target.closest('[data-dd="captype"]');
    if (!dd) return;
    var items = [{ label: 'All types', checked: !capType, fn: function () { capType = ''; saveLocal(); render(); } }];
    state.teamTypes.forEach(function (t) {
      items.push({ label: esc(t), checked: capType === t, fn: function () { capType = t; saveLocal(); render(); } });
    });
    openDropdown(dd, items);
  });

  // click a capacity cell to toggle that week as a holiday week
  $('#hdrCap').addEventListener('click', function (e) {
    if (dragConsumedClick) { dragConsumedClick = false; return; }
    var cell = e.target.closest('[data-w]');
    if (!cell) return;
    var w = parseInt(cell.dataset.w, 10);
    var iso = RM.fmtISO(RM.weekStartDate(state.meta, w));
    var isFull = RM.holidaysInWeek(state.meta, w) === 5;
    commit('toggle holiday', function (s) {
      if (isFull) {
        // clear every holiday date falling in this week
        s.meta.holidays = s.meta.holidays.filter(function (h) {
          var d = RM.dateToDay(s.meta, RM.parseISO(h));
          return d == null || Math.floor(d / 5) !== w;
        });
      } else {
        var seen = {};
        s.meta.holidays.forEach(function (h) { seen[h] = true; });
        for (var i = 0; i < 5; i++) {
          var dIso = RM.fmtISO(RM.dayToDate(s.meta, w * 5 + i));
          if (dIso && !seen[dIso]) s.meta.holidays.push(dIso);
        }
        s.meta.holidays.sort();
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
    line.style.left = 'calc(var(--left-w) + ' + (d * dayPx()) + 'px)';
  }

  function warnBadge(it) {
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
    var list = validation.byItem[it.id] || [];
    if (!list.length) return;
    hoverTip.innerHTML = list.map(function (v) {
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

  function itemRowsHtml(html, it, cyclic) {
    var meta = state.meta;
    var color = '#' + RM.colorForItem(state, it);
    var vlist = validation.byItem[it.id] || [];
    var sizeCls = !it.size ? (isScheduled(it) ? ' missing' : '') : (sizeMatches(it) ? '' : ' custom');
    var rk = RM.depRisk(state, it, cyclic);
    var riskTitle = 'Risk — click to change' +
      (rk.level !== 'none' ? '\nDependency risk ' + rk.level + ':\n· ' + rk.reasons.join('\n· ') : '');
    var laneInner = '';
    var ports = '<span class="port p-in" data-port="in" title="Drag to another bar: this depends on it"></span>' +
      '<span class="port p-out" data-port="out" title="Drag to another bar: it depends on this"></span>';
    if (isScheduled(it)) {
      var left = it.startDay * dayPx();
      var workW = Math.max(6, it.durDays * dayPx());
      var riskW = (it.riskDays || 0) * dayPx();
      var width = workW + riskW;
      var hcTag = (state.meta.capacityEnabled && it.headcount > 1 && width > 90) ? '<span class="b-hc">×' + it.headcount + '</span>' : '';
      // label rides inside the bar when it fits (~6.3px/char at 10.5px bold);
      // otherwise it sits just right of the bar in ink
      var labelW = it.feature.length * 6.3 + 16 + (hcTag ? 30 : 0);
      var label = '<span class="b-label' + (labelW <= width ? '' : ' out') + '">' + esc(it.feature) + '</span>';
      var dates = RM.fmtShort(RM.dayToDate(meta, it.startDay)) + ' → ' +
        RM.fmtShort(RM.spanEndDate(meta, it.startDay, RM.itemSpan(it)));
      var tip = it.feature + '  ·  ' + dates + (it.size ? '  ·  ' + it.size : '') +
        (it.risk ? '  ·  risk ' + it.risk : '') + '  ·  ×' + it.headcount +
        (it.description ? '\n' + RM.htmlToText(it.description) : '');
      // one uniform duration on the timeline — the work/risk split lives in
      // the panel, not the paint
      laneInner =
        '<div class="bar' + (selectedId === it.id ? ' selected' : '') +
        (it.locked ? ' locked' : '') + (it.done ? ' done-bar' : '') + (width < 34 ? ' tiny' : '') +
        (showCrit && critCache && critCache.items[it.id] ? ' crit' : '') +
        '" data-bar="' + it.id + '"' +
        ' style="left:' + left + 'px;width:' + width + 'px;--bar-c:' + color + '"' +
        ' title="' + esc(tip) + '">' +
        '<div class="b-h l" data-act="bh-l"></div>' + label + hcTag +
        (it.locked ? '<span class="b-hc bi-lock" style="pointer-events:none"><i data-lucide="lock"></i></span>' : '') +
        '<div class="b-h r" data-act="bh-r"></div>' +
        ports +
        '</div>';
    }
    // unscheduled: the lane stays empty — hovering it previews the landing
    // slot (1 week unless sized), clicking places the item there
    if (view === 'scoping') {
      var cells = ['<div class="sc-row">'];
      var epIco2 = RM.iconForEpic(state, it.epic);
      var fixedContent = {
        size: '<span class="r-size' + sizeCls + '" tabindex="0" role="button" data-act="size" title="T-shirt size — click to change">' + (it.size || '·') + '</span>',
        risk: '<span class="r-risk rk-' + rk.level + (it.risk ? ' has-risk' : '') + '" tabindex="0" role="button" data-act="risk" title="' + esc(riskTitle) + '">' +
          (it.risk || (rk.level === 'none' ? '·' : rk.level.charAt(0).toUpperCase())) + '</span>',
        duration: (isScheduled(it)
          ? '<span class="r-wk editable" tabindex="0" role="button" data-act="wk" title="Weeks — click to edit">' + totalWeeks(it) + '</span>'
          : '<span class="r-wk" title="Estimated weeks from size">' + totalWeeks(it) + '</span>'),
        workstream: '<span class="r-ws sc-chip" tabindex="0" role="button" data-act="ws" title="Workstream — click to change">' +
          '<span class="dd-dot" style="background:#' + RM.colorForWs(state, it.workstream) + '"></span>' +
          (it.workstream ? esc(shorten(it.workstream, 18)) : '·') + '</span>',
        epic: '<span class="r-ws sc-chip" tabindex="0" role="button" data-act="epic" title="Epic — click to change">' +
          (epIco2 ? '<i data-lucide="' + epIco2 + '"></i>' : '') +
          (it.epic ? esc(shorten(it.epic, 20)) : '·') + '</span>'
      };
      SCOPE_FIXED.forEach(function (c) {
        cells.push('<div class="sc-cell sc-fix" data-col="' + c[0] + '" style="width:' + scopeColWidth(c) + 'px">' +
          fixedContent[c[0]] + '</div>');
      });
      scopeCols().forEach(function (c) {
        var sv = RM.scopeValue(it, c[0]);
        cells.push('<div class="sc-cell" data-col="' + c[0] + '" style="width:' + scopeColWidth(c) + 'px">' +
          (c[0] === 'description'
            // description holds rich text — edit it in place as such
            ? '<div class="sc-edit sc-rich" contenteditable="true" data-scope="description">' + sanitizeHtml(sv) + '</div>'
            : '<textarea class="sc-edit" data-scope="' + c[0] + '">' + esc(sv) + '</textarea>') + '</div>');
      });
      cells.push('</div>');
      laneInner = cells.join('');
    }
    html.push(
      '<div class="row item' + (view === 'scoping' ? ' scope' : '') +
      (selectedId === it.id ? ' selected' : '') + (it.done ? ' done' : '') +
      '" data-id="' + it.id + '">' +
      '<div class="row-left" title="Drag to reorder / move phase">' +
      '<span class="r-grip" data-act="grip"><i data-lucide="grip-vertical"></i></span>' +
      '<span class="r-chev' + (expanded[it.id] ? ' open' : '') + '" data-act="stories" title="Stories (' + it.stories.length + ')">' +
      (it.stories.length ? '<i data-lucide="chevron-right"></i>' : '<span style="opacity:.35"><i data-lucide="chevron-right"></i></span>') + '</span>' +
      '<span class="r-num">' + it.num + '</span>' +
      '<span class="r-dot" style="background:' + color + '"></span>' +
      '<div class="r-main">' +
      (it.locked ? '<span class="r-lock"><i data-lucide="lock"></i></span>' : '') +
      '<input class="r-name" data-rowname spellcheck="false" value="' + esc(it.feature) + '" placeholder="(untitled)" title="' + esc(it.feature + (it.description ? '\n' + RM.htmlToText(it.description) : '')) + '">' +
      (it.epic && !groupEpic ? '<span class="r-epic" title="' + esc(it.epic) + '">' +
        (RM.iconForEpic(state, it.epic) ? '<i data-lucide="' + RM.iconForEpic(state, it.epic) + '"></i>' : '') +
        esc(it.epic) + '</span>' : '') +
      '</div>' +
      (view === 'scoping' ? '' :
        '<span class="r-size' + sizeCls + '" tabindex="0" role="button" data-act="size" title="T-shirt size — click to change">' + (it.size || '·') + '</span>' +
        (isScheduled(it)
          ? '<span class="r-wk editable" tabindex="0" role="button" data-act="wk" title="Weeks — click to edit">' + totalWeeks(it) + '</span>'
          : '<span class="r-wk" title="Estimated weeks from size">' + totalWeeks(it) + '</span>') +
        (state.meta.capacityEnabled
          ? '<span class="r-hc' + (it.headcount > 1 ? ' multi' : '') + '" tabindex="0" role="button" data-act="hc" title="Headcount — click to edit">×' + it.headcount + '</span>'
          : '')) +
      warnBadge(it) +
      '</div>' +
      '<div class="row-lane">' + laneInner + '</div>' +
      '</div>');

    if (expanded[it.id] && view !== 'scoping') {
      it.stories.forEach(function (st) {
        var stSched = st.startDay != null && st.durDays != null;
        html.push(
          '<div class="row story" data-story="' + st.id + '" data-id="' + it.id + '">' +
          '<div class="row-left"><span class="st-pad"></span>' +
          '<span class="st-title' + (st.done ? ' done' : '') + '" data-act="st-title">' + esc(st.title) + '</span>' +
          '<button class="st-del" data-act="st-del" title="Delete story"><i data-lucide="x"></i></button>' +
          '</div>' +
          '<div class="row-lane"' + (!stSched && view === 'planning' ? ' title="Double-click to add a timeline"' : '') + '>' +
          (stSched
            ? (function () {
                var stW = Math.max(6, st.durDays * dayPx());
                // quiet label: inside when it fits (~5.5px/char at 9.5px), else right of the bar
                var stLabW = st.title.length * 5.5 + 12;
                return '<div class="st-bar" data-stbar="' + st.id + '" data-id="' + it.id + '" style="left:' + (st.startDay * dayPx()) +
                  'px;width:' + stW + 'px;--bar-c:' + color + '">' +
                  '<span class="stb-label' + (stLabW <= stW ? '' : ' out') + '">' + esc(st.title) + '</span>' +
                  '<span class="bh l" data-act="sh-l"></span><span class="bh r" data-act="sh-r"></span></div>';
              })()
            : '') +
          '</div></div>');
      });
      html.push(
        '<div class="row story story-add" data-id="' + it.id + '">' +
        '<div class="row-left"><span class="st-pad"></span>' +
        '<i data-lucide="plus" class="st-add-ico"></i>' +
        '<input class="st-add-input" data-act="st-add" placeholder="Add story…">' +
        '</div><div class="row-lane"></div></div>');
    }
  }

  function renderRows() {
    var html = [];
    var cyclic = RM.cycleMembers(state);
    state.phases.forEach(function (p) {
      var items = RM.itemsInPhase(state, p.id).filter(matchesFilter);
      html.push(
        '<div class="row band" data-kind="band" data-phase="' + p.id + '">' +
        '<div class="row-left">' +
        '<span class="band-chev' + (p.collapsed ? '' : ' open') + '" data-act="phase-toggle" title="Collapse / expand"><i data-lucide="chevron-right"></i></span>' +
        '<span class="band-name">' + esc(p.name) + '</span>' +
        '<span class="band-count">' + items.length + '</span>' +
        (p.bucket ? '<span class="band-bucket-tag">backlog</span>' : '') +
        '<button class="band-add" data-act="phase-additem" title="Add a feature to this phase">+ feature</button>' +
        '<button class="band-edit" data-act="phase-edit" title="Edit phase">edit</button>' +
        '</div>' +
        '<div class="row-lane">' + (p.description ? '<span class="band-desc" title="' + esc(p.description) + '">' + esc(p.description) + '</span>' : '') + '</div>' +
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
      function epicBands(list, sub) {
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
      if (groupWs) {
        var wg = partition(items, 'workstream', true);
        wg.keys.forEach(function (key) {
          html.push(
            '<div class="row eband wsband" data-kind="eband" data-phase="' + p.id +
            '" data-ws="' + esc(key) + '">' +
            '<div class="row-left">' +
            '<span class="r-dot" style="background:#' + RM.colorForWs(state, key) + '"></span>' +
            '<span class="eb-name">' + (key ? esc(key) : '<i>no workstream</i>') +
            '</span><span class="band-count">' + wg.by[key].length + '</span></div>' +
            '<div class="row-lane"></div></div>');
          if (groupEpic) epicBands(wg.by[key], true);
          else wg.by[key].forEach(function (it) { itemRowsHtml(html, it, cyclic); });
        });
      } else if (groupEpic) {
        epicBands(items, false);
      } else {
        items.forEach(function (it) { itemRowsHtml(html, it, cyclic); });
      }

      // blank click-to-add row at the bottom of every phase
      html.push(
        '<div class="row addrow" data-kind="addrow" data-phase="' + p.id + '" title="Add an item to ' + esc(p.name) + '">' +
        '<div class="row-left"><span class="addrow-lab"><i data-lucide="plus"></i> New item</span></div>' +
        '<div class="row-lane"></div></div>');
    });

    rowsEl.innerHTML = html.join('');
    if (window.lucide) lucide.createIcons();
  }

  // total effort (work + risk) in weeks for the row readout. Scheduled items
  // count WORKING days only — a bar stretched across a holiday still reads as
  // its clean size (e.g. 8w, not 8.2w)
  function totalWeeks(it) {
    var days = isScheduled(it)
      ? RM.workInSpan(state.meta, it.startDay, RM.itemSpan(it))
      : RM.effortDays(state, it) + RM.riskEffortDays(state, it);
    if (!days) return '·';
    return fmtDays(days);
  }

  function sizeHuman(size) {
    var d = RM.sizeDays(state, size);
    if (d == null) return '';
    if (d < 5) return d + 'd';
    return (d / 5) + 'w';
  }
  function sizeMatches(it) {
    if (!it.size || !isScheduled(it)) return true;
    var work = RM.workInSpan(state.meta, it.startDay, it.durDays);
    return work === RM.sizeDays(state, it.size);
  }

  // ------------------------------------------------------------ arrows
  function barRect(id) {
    var el = rowsEl.querySelector('[data-bar="' + id + '"]');
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
    gEl.innerHTML = out.join('');
  }

  // one smooth cubic curve; loops back naturally when the target is behind
  function curvePath(sx, sy, tx, ty) {
    var dx = Math.max(28, Math.min(90, Math.abs(tx - sx) * 0.5));
    return 'M' + sx + ',' + sy + ' C' + (sx + dx) + ',' + sy + ' ' + (tx - dx) + ',' + ty + ' ' + tx + ',' + ty;
  }

  // arrow click selects the edge; Delete removes it
  var selectedEdge = null; // { fromId, toId }
  $('#arrows').addEventListener('click', function (e) {
    var g = e.target.closest('g.edge');
    if (!g) return;
    selectedEdge = { fromId: g.dataset.from, toId: g.dataset.to };
    requestAnimationFrame(renderArrows);
  });
  function deleteSelectedEdge() {
    if (!selectedEdge) return false;
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
  function select(id, scrollTo) {
    selectedId = id;
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

  function renderPanel() {
    var panel = $('#panel');
    var it = selectedId ? RM.itemById(state, selectedId) : null;
    if (!it) { panel.hidden = true; panel.innerHTML = ''; return; }
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

    var sizeBtns = RM.SIZE_ORDER.map(function (s) {
      return '<button data-f="size" data-v="' + s + '"' + (it.size === s ? ' class="on"' : '') +
        ' title="' + fmtDays(RM.sizeDays(state, s)) + '">' + s + '</button>';
    }).join('');
    var riskBtns = ['<button data-f="riskSize" data-v=""' + (!it.risk ? ' class="on"' : '') + ' title="No risk">None</button>']
      .concat(RM.RISK_ORDER.map(function (s) {
        return '<button data-f="riskSize" data-v="' + s + '"' + (it.risk === s ? ' class="on"' : '') +
          ' title="' + (s === 'L' ? 'Low' : s === 'M' ? 'Medium' : 'High') + ' risk">' + s + '</button>';
      })).join('');

    var scheduleInfo = '';
    var startD = isScheduled(it) ? RM.dayToDate(meta, it.startDay) : null;
    if (isScheduled(it)) {
      scheduleInfo =
        '<div class="p-grid2">' +
        '<div><label class="p-lab">Start</label>' +
        '<input type="date" data-f="startDate" value="' + RM.fmtISO(startD) + '" style="width:100%"></div>' +
        '<div><label class="p-lab">Weeks</label>' +
        '<input type="number" data-f="durWeeks" min="0.2" step="0.2" value="' + (it.durDays / 5) + '" style="width:100%"></div>' +
        '</div>' +
        '<div class="p-row" style="margin-top:8px">' +
        '<button data-f="snap" title="Earliest slot after dependencies with free capacity">Snap earliest</button>' +
        '<button data-f="unschedule">Unschedule</button>' +
        '</div>';
    } else {
      scheduleInfo =
        '<div class="p-row" style="margin-top:8px">' +
        '<button data-f="schedule-now" class="primary">Place on timeline</button>' +
        '<button data-f="snap">Snap earliest</button>' +
        '</div>';
    }

    var riskInfo = '<div class="seg">' + riskBtns + '</div>';

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
      return '<div class="p-story">' +
        '<input type="text" data-pst-title="' + st.id + '" value="' + esc(st.title) + '">' +
        '<button class="st-del st-edit' + (hasBody ? ' has-body' : '') + '" style="opacity:1" data-pst-edit="' + st.id +
        '" title="Description &amp; acceptance criteria"><i data-lucide="pencil"></i></button>' +
        '<button class="st-del" style="opacity:1" data-pst-del="' + st.id + '"><i data-lucide="x"></i></button></div>';
    }).join('');

    // collapsible section card: the body always renders (collapsed hides via CSS)
    function sec(key, label, summary, body) {
      return '<div class="p-sec c' + (secOpen(key) ? ' open' : '') + '" data-sec="' + key + '">' +
        '<button class="p-sechead" data-sectoggle="' + key + '">' +
        '<i data-lucide="chevron-right"></i><span class="p-seclab">' + label + '</span>' +
        '</button><div class="p-secbody">' + body + '</div></div>';
    }

    var customCols = state.meta.scopeCols.filter(function (c) { return !RM.SCOPE_BUILTIN_LABELS[c.key]; });

    panel.innerHTML =
      '<div id="panelRz" title="Drag to resize"></div>' +
      '<div class="p-top"><span class="p-num">#<input class="p-num-edit" data-f="num" value="' + it.num +
      '" title="Item # — an invalid or taken number picks the next available one"></span>' +
      '<button class="p-close" data-f="close" title="Close (Esc)"><i data-lucide="x"></i></button></div>' +
      '<textarea class="p-name" data-f="feature" rows="1" placeholder="Feature name">' + esc(it.feature) + '</textarea>' +

      sec('desc', 'Description', '',
        wysHtml('description', it.description, 'What is this feature?')) +

      sec('scope', 'Scope', '',
        '<div class="p-fld"><label class="p-lab">Enables</label><textarea class="p-text" data-f="enables">' + esc(it.enables) + '</textarea></div>' +
        '<div class="p-fld"><label class="p-lab">Out of scope</label><textarea class="p-text" data-f="outOfScope">' + esc(it.outOfScope) + '</textarea></div>' +
        '<div class="p-fld"><label class="p-lab">Notes</label><textarea class="p-text" data-f="notes">' + esc(it.notes) + '</textarea></div>' +
        '<div class="p-fld"><label class="p-lab">External dependencies</label><textarea class="p-text" data-f="extDeps">' + esc(it.extDeps) + '</textarea></div>' +
        customCols.map(function (c) {
          return '<div class="p-fld"><label class="p-lab">' + esc(RM.scopeColLabel(c)) + '</label>' +
            '<textarea class="p-text" data-cf="' + c.key + '">' + esc(RM.scopeValue(it, c.key)) + '</textarea></div>';
        }).join('')) +

      sec('details', 'Details', '',
        '<div class="p-grid2">' +
        '<div><label class="p-lab">Phase</label>' + phaseDd + '</div>' +
        '<div><label class="p-lab">Workstream</label>' +
        '<input data-f="workstream" list="wsList" value="' + esc(it.workstream) + '" placeholder="Product / Data / Process" style="width:100%"></div>' +
        '</div>' +
        '<div style="margin-top:8px"><label class="p-lab">Epic</label>' + epicDd + '</div>') +

      sec('schedule', 'Size &amp; schedule', '',
        '<label class="p-lab">Size</label>' +
        '<div class="seg" style="margin-bottom:8px">' + sizeBtns +
        '<button data-f="size" data-v=""' + (!it.size ? ' class="on"' : '') + ' title="No size">—</button></div>' +
        scheduleInfo +
        '<label class="p-lab" style="margin-top:10px">Risk buffer</label>' + riskInfo +
        '<div class="p-row" style="margin-top:10px">' +
        '<label class="p-check fixed"><input type="checkbox" data-f="locked"' + (it.locked ? ' checked' : '') + '> Locked</label>' +
        '<label class="p-check fixed"><input type="checkbox" data-f="done"' + (it.done ? ' checked' : '') + '> Done</label>' +
        '</div>') +

      sec('people', 'People', '',
        '<div class="p-grid2">' +
        '<div><label class="p-lab">Headcount</label>' +
        '<input type="number" data-f="headcount" min="0.1" step="0.25" value="' + it.headcount + '" style="width:100%"></div>' +
        '<div><label class="p-lab">Role</label>' + typeDd + '</div>' +
        '</div>') +

      sec('deps', 'Dependencies', '',
        '<label class="p-lab">Depends on</label>' +
        '<div class="chips">' + depChips + depTextChips + (depChips || depTextChips ? '' : '<span class="p-none">none</span>') + '</div>' +
        '<div class="dep-search"><input data-f="depsearch" placeholder="+ add dependency — search by name…" autocomplete="off" style="width:100%;margin-top:7px">' +
        '<div class="dep-sug" hidden></div></div>' +
        '<label class="p-lab" style="margin-top:10px">Dependents (rely on this)</label>' +
        '<div class="chips">' + dependentChips + (dependentChips ? '' : '<span class="p-none">none</span>') + '</div>' +
        '<div class="m-hint">Tip: hover a bar and drag its edge circles to another bar to link.</div>') +

      sec('stories', 'Stories', '',
        '<div class="p-stories">' + storyRows + '</div>' +
        '<input data-f="storyadd" placeholder="+ add story…" style="width:100%;margin-top:6px">') +

      (vlist.length
        ? sec('checks', 'Checks', '',
          '<div class="p-warnlist">' + vlist.map(function (v) {
            var cls = v.level === 'error' ? 'err' : v.level;
            return '<div class="p-warnitem ' + cls + '">' + esc(v.msg) + '</div>';
          }).join('') + '</div>')
        : '') +

      '<div class="p-actions">' +
      '<button data-f="duplicate">Duplicate</button>' +
      '<button data-f="delete" class="danger">Delete</button>' +
      '</div>' +
      '<datalist id="wsList"><option>Product</option><option>Data</option><option>Process</option><option>Product / Process</option><option>All</option></datalist>';
    if (window.lucide) lucide.createIcons();
    var pn = $('.p-name', panel);
    if (pn) { pn.style.height = 'auto'; pn.style.height = pn.scrollHeight + 'px'; }
  }

  // ---- shared dropdown: a button that opens the same list UI the menu bar
  // uses (icons/dots, blue selected row, check on the right, optional edit)
  function ddButton(dd, label, dot, title) {
    return '<button class="dd-btn" data-dd="' + dd + '" title="' + esc(title || '') + '">' +
      (dot ? '<span class="dd-dot" style="background:' + dot + '"></span>' : '') +
      '<span class="dd-label">' + label + '</span><i data-lucide="chevron-down"></i></button>';
  }
  function openDropdown(anchor, items) {
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
    var html = '<div class="menu-list" style="min-width:' + Math.max(210, Math.round(r.width)) + 'px">' +
      items.map(function (m, i) {
        if (m.sep) return '<div class="menu-sep"></div>';
        return '<button data-mi="' + i + '"' + (m.checked ? ' class="on"' : '') + '>' +
          (m.dot ? '<span class="dd-dot" style="background:' + m.dot + '"></span>' : '') +
          (m.icon ? '<i data-lucide="' + m.icon + '"></i>' : '') +
          '<span>' + m.label + '</span>' +
          (m.edit ? '<span class="mi-edit" data-me="' + i + '" title="Edit"><i data-lucide="pencil"></i></span>' : '') +
          (m.checked ? '<i data-lucide="check" class="mi-check"></i>' : '') +
          '</button>';
      }).join('') + '</div>';
    openPopover(r.left, r.bottom + 4, html, function (host) {
      if (window.lucide) lucide.createIcons();
      // keyboard: focus lands on the current choice; arrows move; Enter picks
      var btns = $$('.menu-list [data-mi]', host);
      var start = host.querySelector('.menu-list button.on') || btns[0];
      if (start) start.focus({ preventScroll: true });
      host.addEventListener('keydown', function (ev) {
        if (ev.key !== 'ArrowDown' && ev.key !== 'ArrowUp') return;
        ev.preventDefault();
        var i2 = btns.indexOf(document.activeElement);
        var n = ev.key === 'ArrowDown' ? i2 + 1 : i2 - 1;
        if (n < 0) n = btns.length - 1;
        if (n >= btns.length) n = 0;
        btns[n].focus({ preventScroll: true });
      });
      host.addEventListener('click', function (ev) {
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
      .concat(st.team.map(function (x) { return x.workstream; }))
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
    out.sort(function (a, b) { return a.localeCompare(b); });
    return out;
  }

  // edit an epic's label + icon (applies to every item carrying it)
  var EPIC_ICONS = ['tag', 'star', 'flag', 'rocket', 'target', 'layers', 'database', 'shield',
    'zap', 'globe', 'users', 'wrench', 'chart-line', 'box', 'lightbulb', 'compass',
    'cpu', 'plug', 'bot', 'flask-conical', 'map', 'workflow', 'network', 'building'];
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
          closeModal();
          commit('edit epic', function (s) {
            var name2 = newName || epicName;
            if (name2 !== epicName) {
              s.items.forEach(function (x) { if (x.epic === epicName) x.epic = name2; });
              if (s.epicIcons[epicName] != null) { s.epicIcons[name2] = s.epicIcons[epicName]; delete s.epicIcons[epicName]; }
            }
            if (picked) s.epicIcons[name2] = picked;
            else delete s.epicIcons[name2];
          });
        };
      });
  }

  // edit a workstream's label + color (applies to items and roles carrying it)
  function wsEditModal(wsName) {
    var setting = state.wsColors[wsName] || null;
    var cur = '#' + RM.colorForWs(state, wsName);
    var count = state.items.filter(function (x) { return x.workstream === wsName; }).length;
    var sw = '<button class="swatch' + (!setting || setting === 'product' ? ' on' : '') + '" data-esw="product" style="background:#' + RM.PALETTE.product + '" title="Default blue"></button>' +
      RM.PALETTE_KEYS.slice(1).map(function (k) {
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
        var picked = setting; // palette key, hex, or null (default blue)
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
                s.team.forEach(function (m) { if (m.workstream === wsName) m.workstream = ''; });
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
              s.team.forEach(function (m) { if (m.workstream === wsName) m.workstream = name2; });
              if (s.wsColors[wsName] != null) { s.wsColors[name2] = s.wsColors[wsName]; delete s.wsColors[wsName]; }
            }
            // store the choice EXPLICITLY — deleting the entry would let the
            // known-workstream default (e.g. Product = plum) re-seed on the
            // next load and silently revert a "Default blue" pick
            s.wsColors[name2] = picked || 'product';
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
    if (d < 5) return d + 'd';
    var w = d / 5;
    return (Math.round(w * 10) / 10) + 'w';
  }
  function shorten(s, n) { return s.length > n ? s.slice(0, n - 1) + '…' : s; }

  // panel events (delegated)
  $('#panel').addEventListener('click', function (e) {
    var it = selectedId && RM.itemById(state, selectedId);
    if (!it) return;
    var secBtn = e.target.closest('[data-sectoggle]');
    if (secBtn) {
      var sk = secBtn.dataset.sectoggle;
      panelSec[sk] = secOpen(sk); // open → store collapsed=true, and vice versa
      saveLocal();
      renderPanel();
      return;
    }
    var btn = e.target.closest('[data-f]');
    var dep = e.target.closest('[data-deprm]');
    var depTxt = e.target.closest('[data-deptxtrm]');
    var rdep = e.target.closest('[data-rdep]');
    var depGo = e.target.closest('[data-depgo]');
    var stDel = e.target.closest('[data-pst-del]');
    var stEd = e.target.closest('[data-pst-edit]');
    if (stEd) { storyModal(it.id, stEd.dataset.pstEdit); return; }
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
    if (!btn) return;
    var f = btn.dataset.f;
    if (f === 'close') { select(null); return; }
    if (f === 'size') {
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
    if (f === 'snap') {
      var r = RM.snapEarliest(state, it.id);
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
        t.durDays = RM.stretchSpan(s.meta, t.startDay, RM.effortDays(s, t));
        t.riskDays = RM.stretchSpan(s.meta, t.startDay + t.durDays, RM.riskEffortDays(s, t));
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
    if (f === 'duplicate') {
      commit('duplicate', function (s) {
        var t = RM.itemById(s, it.id);
        var copy = RM.clone(t);
        copy.id = RM.uid('i');
        copy.num = RM.nextNum(s);
        copy.feature = t.feature + ' (copy)';
        copy.stories.forEach(function (st) { st.id = RM.uid('s'); });
        var idx = s.items.indexOf(t);
        s.items.splice(idx + 1, 0, copy);
        selectedId = copy.id;
      });
      return;
    }
    if (f === 'delete') {
      confirmBox('Delete #' + it.num + '?', esc(it.feature) +
        '<br><br>Items depending on it will keep a dangling reference (flagged by validation).',
        'Delete', function () {
          commit('delete', function (s) {
            s.items = s.items.filter(function (x) { return x.id !== it.id; });
            selectedId = null;
          });
        }, true);
      return;
    }
  });

  // story editor: title + rich Description / Acceptance Criteria
  function storyModal(itemId, stId) {
    var it = RM.itemById(state, itemId);
    var st = null;
    (it ? it.stories : []).forEach(function (x) { if (x.id === stId) st = x; });
    if (!st) return;
    openModal(
      '<div class="modal" style="width:560px">' +
      '<div class="m-head"><h2>Edit story</h2>' +
      '<button class="p-close" data-m="x"><i data-lucide="x"></i></button></div>' +
      '<div class="m-body">' +
      '<div class="m-sec"><label>Title</label><input id="stTitle" style="width:100%" value="' + esc(st.title) + '"></div>' +
      '<div class="m-sec"><label>Description</label>' + wysHtml('stDesc', st.description, 'What does this story cover?') + '</div>' +
      '<div class="m-sec"><label>Acceptance criteria</label>' + wysHtml('stAc', st.ac, 'When is it done?') + '</div>' +
      '</div>' +
      '<div class="m-foot"><button data-m="cancel">Cancel</button>' +
      '<button data-m="save" class="primary">Save</button></div></div>',
      function (host) {
        $('[data-m=x]', host).onclick = closeModal;
        $('[data-m=cancel]', host).onclick = closeModal;
        $('[data-m=save]', host).onclick = function () {
          var title = $('#stTitle', host).value.trim() || st.title;
          var dv = sanitizeHtml($('.wz-ed[data-f=stDesc]', host).innerHTML);
          var av = sanitizeHtml($('.wz-ed[data-f=stAc]', host).innerHTML);
          closeModal();
          commit('edit story', function (s) {
            RM.itemById(s, itemId).stories.forEach(function (x) {
              if (x.id === stId) { x.title = title; x.description = dv; x.ac = av; }
            });
          });
        };
      });
  }

  // the rich description commits on blur (contenteditable has no change event)
  $('#panel').addEventListener('focusout', function (e) {
    var ed = e.target.classList && e.target.classList.contains('wz-ed') ? e.target : null;
    if (!ed || ed.dataset.f !== 'description') return;
    var it = selectedId && RM.itemById(state, selectedId);
    if (!it) return;
    var v = sanitizeHtml(ed.innerHTML);
    if (v === it.description) return;
    commit('description', function (s) { RM.itemById(s, it.id).description = v; });
  });

  $('#panel').addEventListener('change', function (e) {
    var it = selectedId && RM.itemById(state, selectedId);
    if (!it) return;
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
    if (f === 'headcount') {
      // decimals are fine (0.5 = a half-time person); capacity math is hours-based
      var hv = parseFloat(val);
      hv = isFinite(hv) && hv > 0 ? Math.round(hv * 100) / 100 : 1;
      commit('headcount', function (s) { RM.itemById(s, it.id).headcount = hv; });
      return;
    }
    var simple = { feature: 1, description: 1, workstream: 1, enables: 1, outOfScope: 1, notes: 1, extDeps: 1 };
    if (simple[f]) {
      commit(f, function (s) { RM.itemById(s, it.id)[f] = val; });
      return;
    }
    if (f === 'num') {
      var newNum;
      commit('renumber', function (s) { newNum = RM.renumberItem(s, it.id, val); });
      if (newNum != null && String(newNum) !== String(val).trim()) {
        toast('#' + val + ' isn’t available — used #' + newNum + ' instead');
      }
      return;
    }
    if (f === 'locked') { commit('lock', function (s) { RM.itemById(s, it.id).locked = !!val; }); return; }
    if (f === 'done') { commit('done', function (s) { RM.itemById(s, it.id).done = !!val; }); return; }
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
      var wv = Math.max(0.2, parseFloat(val) || 1);
      commit('duration', function (s) {
        RM.itemById(s, it.id).durDays = Math.max(1, Math.round(wv * 5));
      });
      return;
    }
    if (f === 'storyadd') {
      var sv = val.trim();
      if (sv) {
        expanded[it.id] = true;
        commit('add story', function (s) {
          RM.itemById(s, it.id).stories.push({ id: RM.uid('s'), title: sv, done: false });
        });
      }
      return;
    }
  });

  // Enter in the story add input triggers change
  $('#panel').addEventListener('keydown', function (e) {
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
  $('#panel').addEventListener('input', function (e) {
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

  rowsEl.addEventListener('click', function (e) {
    if (dragConsumedClick) { dragConsumedClick = false; return; }
    // scoping text cells and inline editors edit in place — selecting would
    // re-render and steal their focus
    if (e.target.closest('.sc-edit,.st-add-input,.hc-edit,input.r-name')) return;
    var act = e.target.closest('[data-act]');
    var rowEl = e.target.closest('.row');
    if (!rowEl) return;

    if (rowEl.dataset.kind === 'addrow') {
      addFeature(rowEl.dataset.phase);
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
      if (act.dataset.act === 'st-del') {
        commit('delete story', function (s) {
          var t = RM.itemById(s, itemId);
          t.stories = t.stories.filter(function (st) { return st.id !== stId; });
        });
      } else if (act.dataset.act === 'st-title') {
        startInlineEdit(act, function (v) {
          commit('story title', function (s) {
            RM.itemById(s, itemId).stories.forEach(function (st) { if (st.id === stId) st.title = v; });
          });
        });
      }
      return;
    }

    if (act) {
      switch (act.dataset.act) {
        case 'wk': {
          if (act.querySelector('input') || !isScheduled(it)) return;
          var wkInp = document.createElement('input');
          wkInp.type = 'number';
          wkInp.min = '0.2';
          wkInp.step = '0.2';
          wkInp.value = Math.round((RM.workInSpan(state.meta, it.startDay, it.durDays) / 5) * 10) / 10;
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
            if (saveIt && isFinite(wv) && wv > 0) {
              commit('duration', function (s) {
                var t = RM.itemById(s, itemId);
                t.durDays = RM.stretchSpan(s.meta, t.startDay, Math.max(1, Math.round(wv * 5)));
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
        case 'hc': {
          if (act.querySelector('input')) return;
          var hcInp = document.createElement('input');
          hcInp.type = 'number';
          hcInp.min = '0.1';
          hcInp.step = '0.25';
          hcInp.value = it.headcount;
          hcInp.className = 'hc-edit';
          act.textContent = '×';
          act.appendChild(hcInp);
          hcInp.focus();
          hcInp.select();
          var hcDone = false;
          var hcFin = function (saveIt) {
            if (hcDone) return; hcDone = true;
            var n = parseFloat(hcInp.value);
            if (saveIt && isFinite(n) && n > 0) {
              n = Math.round(n * 100) / 100;
              commit('headcount', function (s) {
                RM.itemById(s, itemId).headcount = n;
              });
            } else render();
          };
          hcInp.addEventListener('blur', function () { hcFin(true); });
          hcInp.addEventListener('keydown', function (ev) {
            ev.stopPropagation();
            if (ev.key === 'Enter') hcFin(true);
            if (ev.key === 'Escape') hcFin(false);
          });
          return;
        }
        case 'size': {
          openDropdown(act, [{ label: '<i>no size</i>', checked: !it.size, fn: function () {
            setItemSize(itemId, null);
          } }].concat(RM.SIZE_ORDER.map(function (sz) {
            return { label: sz + ' <small>' + sizeHuman(sz) + '</small>', checked: it.size === sz, fn: function () {
              setItemSize(itemId, sz);
            } };
          })));
          return;
        }
        case 'risk': {
          openDropdown(act, [{ label: '<i>None</i>', checked: !it.risk, fn: function () {
            setItemRisk(itemId, null);
          } }].concat(RM.RISK_ORDER.map(function (sz) {
            return { label: sz, checked: it.risk === sz, fn: function () {
              setItemRisk(itemId, sz);
            } };
          })));
          return;
        }
        case 'epic': {
          openDropdown(act, setEpicMenu(itemId, true));
          return;
        }
        case 'ws': {
          var wsList = allWorkstreams();
          var wsItems = [{ label: '<i>— none —</i>', checked: !it.workstream, fn: function () {
            commit('workstream', function (s) { RM.itemById(s, itemId).workstream = ''; });
          } }];
          wsList.forEach(function (w) {
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
            var chip = rowsEl.querySelector('.row[data-id="' + itemId + '"] .r-ws');
            if (!chip) return;
            chip.textContent = it.workstream || '';
            startInlineEdit(chip, function (v) {
              commit('workstream', function (s) { RM.itemById(s, itemId).workstream = v; });
            });
          } });
          openDropdown(act, wsItems);
          return;
        }
        case 'stories': {
          expanded[itemId] = !expanded[itemId];
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
      select(null);
      return;
    }

    // empty lane space deselects; two quick clicks on an unscheduled row's
    // lane place the item (manual count — see lanePress above)
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
      select(null);
      return;
    }
    // clicking an already-open item closes the edit panel
    select(selectedId === itemId ? null : itemId);
  });

  // set size / risk from the chip dropdowns
  function setItemSize(itemId, sz) {
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

  // hovering an unscheduled row's lane previews where a click would place it
  var placePrev = null;
  function clearPlacePreview() {
    if (placePrev) { placePrev.remove(); placePrev = null; }
  }
  rowsEl.addEventListener('mousemove', function (e) {
    if (view !== 'planning' || drag) { clearPlacePreview(); return; }
    var lane = e.target.closest('.row-lane');
    // story lanes preview a mini bar where a double-click would land
    var stRow = e.target.closest('.row.story[data-story]');
    if (stRow && lane) {
      var pit = RM.itemById(state, stRow.dataset.id);
      var pst = pit && storyById(pit, stRow.dataset.story);
      if (!pst || pst.startDay != null || e.target.closest('[data-stbar]')) { clearPlacePreview(); return; }
      var sDay = Math.max(0, snapTo(laneDayAt(e.clientX)));
      var sDur = RM.stretchSpan(state.meta, sDay, 5);
      if (!placePrev || placePrev.parentNode !== lane || !placePrev.classList.contains('st-bar')) {
        clearPlacePreview();
        placePrev = document.createElement('div');
        placePrev.className = 'st-bar preview place-preview';
        lane.appendChild(placePrev);
      }
      placePrev.style.left = (sDay * dayPx()) + 'px';
      placePrev.style.width = (sDur * dayPx()) + 'px';
      placePrev.style.setProperty('--bar-c', '#' + RM.colorForItem(state, pit));
      placePrev.title = 'Double-click to place here';
      return;
    }
    var rowEl = e.target.closest('.row.item');
    var it = rowEl && RM.itemById(state, rowEl.dataset.id);
    if (!lane || !it || isScheduled(it) || e.target.closest('[data-bar],.port')) {
      clearPlacePreview();
      return;
    }
    var day = Math.max(0, snapTo(laneDayAt(e.clientX)));
    var dur = RM.stretchSpan(state.meta, day, RM.effortDays(state, it) || 5);
    if (!placePrev || placePrev.parentNode !== lane) {
      clearPlacePreview();
      placePrev = document.createElement('div');
      placePrev.className = 'bar preview place-preview';
      lane.appendChild(placePrev);
    }
    placePrev.style.left = (day * dayPx()) + 'px';
    placePrev.style.width = (dur * dayPx()) + 'px';
    placePrev.style.setProperty('--bar-c', '#' + RM.colorForItem(state, it));
    placePrev.title = 'Double-click to place here';
  });
  rowsEl.addEventListener('mouseleave', clearPlacePreview);

  // right-click context menus on rows
  rowsEl.addEventListener('contextmenu', function (e) {
    var rowEl = e.target.closest('.row');
    if (!rowEl || e.target.closest('input,textarea,select')) return;
    e.preventDefault();
    var cx = e.clientX, cy = e.clientY;
    var items = [];
    if (rowEl.dataset.kind === 'band') {
      var phaseId = rowEl.dataset.phase;
      items = [
        { icon: 'plus', label: 'Add feature here', fn: function () { addFeature(phaseId); } },
        { icon: 'plus', label: 'New phase…', fn: function () { phaseModal(null); } },
        { sep: true },
        { icon: 'pencil', label: 'Edit phase…', fn: function () { phaseModal(phaseId); } },
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
        stm.startDay != null ? { icon: 'calendar-off', label: 'Remove timeline', fn: function () {
          commit('story timeline', function (s) {
            var st = storyById(RM.itemById(s, stmItemId), stmId);
            if (st) { st.startDay = null; st.durDays = null; }
          });
        } } : null,
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
      items = [
        { icon: 'plus', label: 'Add story', fn: function () {
          expanded[itemId] = true;
          render();
          requestAnimationFrame(function () {
            var inp = rowsEl.querySelector('.row.story-add[data-id="' + itemId + '"] .st-add-input');
            if (inp) inp.focus();
          });
        } },
        { icon: 'plus', label: 'Insert feature above', fn: function () { addFeatureNear(itemId, 0); } },
        { icon: 'plus', label: 'Insert feature below', fn: function () { addFeatureNear(itemId, 1); } },
        { icon: 'plus', label: 'New phase…', fn: function () { phaseModal(null); } },
        { sep: true },
        { icon: 'folder-input', label: 'Move to phase…', fn: function () { openContextMenu(cx, cy, movePhaseMenu(itemId)); } },
        { icon: 'tag', label: 'Set epic…', fn: function () { openContextMenu(cx, cy, setEpicMenu(itemId, false)); } },
        { sep: true },
        isScheduled(it) ? { icon: 'calendar-off', label: 'Unschedule', fn: function () {
          commit('unschedule', function (s) {
            var t = RM.itemById(s, itemId);
            t.startDay = null; t.durDays = null; t.riskDays = 0;
          });
        } } : null,
        { icon: it.locked ? 'lock-open' : 'lock', label: it.locked ? 'Unlock' : 'Lock', fn: function () {
          commit('lock', function (s) { var t = RM.itemById(s, itemId); t.locked = !t.locked; });
        } },
        { icon: it.done ? 'circle' : 'circle-check', label: it.done ? 'Unmark as done' : 'Mark as done', fn: function () {
          commit('done', function (s) { var t = RM.itemById(s, itemId); t.done = !t.done; });
        } },
        { sep: true },
        { icon: 'copy', label: 'Duplicate', fn: function () { duplicateItem(itemId); } },
        { icon: 'trash-2', label: 'Delete #' + it.num + '…', fn: function () { deleteItemConfirm(itemId); } }
      ].filter(Boolean);
    } else return;
    openContextMenu(cx, cy, items);
  });

  function openContextMenu(x, y, items) {
    var html = '<div class="menu-list">' + items.map(function (m, i) {
      if (m.sep) return '<div class="menu-sep"></div>';
      return '<button data-mi="' + i + '"' + (m.checked ? ' class="on"' : '') + '>' +
        (m.dot ? '<span class="dd-dot" style="background:' + m.dot + '"></span>' : '') +
        (m.icon ? '<i data-lucide="' + m.icon + '"></i>' : '') +
        '<span>' + m.label + '</span>' +
        (m.checked ? '<i data-lucide="check" class="mi-check"></i>' : '') +
        '</button>';
    }).join('') + '</div>';
    openPopover(x, y, html, function (host) {
      if (window.lucide) lucide.createIcons();
      host.addEventListener('click', function (ev) {
        var mi = ev.target.closest('[data-mi]');
        if (!mi) return;
        closePopover();
        items[parseInt(mi.dataset.mi, 10)].fn();
      });
    });
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
          delete s.epicColors[epicName];
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
        items: [{ id: newId, num: RM.nextNum(s), phaseId: anchor.phaseId, feature: '', size: 'M', headcount: 1 }]
      }).items[0];
      it.id = newId;
      it.num = RM.nextNum(s);
      it.holdPos = true;
      var idx = s.items.indexOf(RM.itemById(s, itemId));
      s.items.splice(idx + offset, 0, it);
    });
    var row = rowsEl.querySelector('.row[data-id="' + newId + '"]');
    if (row && row.scrollIntoView) row.scrollIntoView({ block: 'nearest' });
    var inp = row && row.querySelector('input.r-name');
    if (inp) inp.focus();
  }

  function duplicateItem(itemId) {
    commit('duplicate', function (s) {
      var t = RM.itemById(s, itemId);
      if (!t) return;
      var copy = RM.clone(t);
      copy.id = RM.uid('i');
      copy.num = RM.nextNum(s);
      copy.feature = t.feature + ' (copy)';
      copy.stories.forEach(function (st) { st.id = RM.uid('s'); });
      s.items.splice(s.items.indexOf(t) + 1, 0, copy);
      selectedId = copy.id;
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

  // give a story its own little timeline where the pointer sits on its lane
  function placeStoryAt(itemId, stId, clientX) {
    var cur = storyById(RM.itemById(state, itemId) || {}, stId);
    if (!cur || cur.startDay != null) return;
    placedAt = Date.now(); placedKey = 'st:' + stId;
    var day = Math.max(0, snapTo(laneDayAt(clientX)));
    commit('place story', function (s) {
      var st = storyById(RM.itemById(s, itemId), stId);
      if (!st) return;
      st.startDay = day;
      st.durDays = RM.stretchSpan(s.meta, day, 5);
    });
  }

  // schedule an unscheduled item where the pointer sits on its lane
  var placedAt = 0, placedKey = null; // swallow the trailing native dblclick after a manual place
  function placeItemAt(itemId, clientX) {
    placedAt = Date.now(); placedKey = itemId;
    var pDay = Math.max(0, snapTo(laneDayAt(clientX)));
    commit('place item', function (s2) {
      var t2 = RM.itemById(s2, itemId);
      t2.startDay = pDay;
      t2.durDays = RM.stretchSpan(s2.meta, pDay, RM.effortDays(s2, t2) || 5);
      if (autoOrder) RM.sortItemsByStart(s2);
    });
    select(itemId);
  }

  function justPlaced(key) { return placedKey === key && Date.now() - placedAt < 500; }
  rowsEl.addEventListener('dblclick', function (e) {
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
    var max = 66;
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
  rowsEl.addEventListener('focusin', function (e) {
    if (!e.target.classList || !e.target.classList.contains('sc-rich')) return;
    var bar = ensureScFmtBar();
    var r = e.target.getBoundingClientRect();
    bar.style.left = Math.max(4, r.left) + 'px';
    bar.style.top = Math.max(4, r.top - 32) + 'px';
    bar.hidden = false;
  });
  rowsEl.addEventListener('focusout', function (e) {
    if (e.target.classList && e.target.classList.contains('sc-rich') && scFmtBar) {
      scFmtBar.hidden = true;
    }
  });

  // inline row-title rename
  rowsEl.addEventListener('change', function (e) {
    if (e.target.dataset && e.target.dataset.rowname != null) {
      var rowEl2 = e.target.closest('.row');
      var rid = rowEl2 && rowEl2.dataset.id;
      var nv = e.target.value;
      if (rid) commit('rename', function (s2) { RM.itemById(s2, rid).feature = nv; });
    }
  });

  // the rich scoping description cell commits on blur
  rowsEl.addEventListener('focusout', function (e) {
    if (!e.target.dataset || e.target.dataset.scope !== 'description') return;
    if (e.target.getAttribute('contenteditable') !== 'true') return;
    var rowEl = e.target.closest('.row');
    var itemId = rowEl && rowEl.dataset.id;
    if (!itemId) return;
    var val = sanitizeHtml(e.target.innerHTML);
    if (val === RM.scopeValue(RM.itemById(state, itemId), 'description')) return;
    commit('scope description', function (s) {
      var t = RM.itemById(s, itemId);
      if (t) RM.setScopeValue(t, 'description', val);
    });
  });

  // scoping view: spreadsheet cells commit on change (blur)
  rowsEl.addEventListener('change', function (e) {
    var f = e.target.dataset && e.target.dataset.scope;
    if (!f) return;
    var rowEl = e.target.closest('.row');
    var itemId = rowEl && rowEl.dataset.id;
    if (!itemId) return;
    var val = e.target.value;
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
        RM.itemById(s, itemId).stories.push({ id: RM.uid('s'), title: v, done: false });
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
    if (stBarEl) {
      var sh = e.target.closest('[data-act="sh-l"],[data-act="sh-r"]');
      startStoryBarDrag(e, stBarEl.dataset.id, stBarEl.dataset.stbar,
        sh ? (sh.dataset.act === 'sh-l' ? 'resize-l' : 'resize-r') : 'move', stBarEl);
      return;
    }
    if (leftEl && !e.target.closest('input,button,textarea,select')) {
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
  function startPortDrag(e, itemId, port) {
    drag = {
      kind: 'port', itemId: itemId, port: port,
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
    $$('.bar.link-target').forEach(function (el) { el.classList.remove('link-target'); });
    dragConsumedClick = true;
  }
  function portDragMove(e) {
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
  function portDragEnd(d) {
    $('#tempLink').setAttribute('hidden', '');
    $$('.bar.link-target').forEach(function (el) { el.classList.remove('link-target'); });
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

  // scoping column resize (widths remembered in the browser)
  $('#hdrSprints').addEventListener('pointerdown', function (e) {
    var rz = e.target.closest('[data-rz]');
    if (!rz) return;
    var field = rz.dataset.rz;
    var col = allScopeCols().filter(function (c) { return c[0] === field; })[0];
    if (!col) return;
    drag = { kind: 'scol', field: field, x0: e.clientX, y0: e.clientY, w0: scopeColWidth(col), moved: false };
    e.preventDefault();
    e.stopPropagation();
  });

  // scoping column menus: reorder / rename / remove, plus the trailing "+"
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
    var mb = e.target.closest('[data-colmenu]');
    if (!mb) return;
    var key = mb.dataset.colmenu;
    var cols = state.meta.scopeCols;
    var idx = -1;
    cols.forEach(function (c, i) { if (c.key === key) idx = i; });
    var items2 = [];
    if (idx > 0) items2.push({ icon: 'arrow-left', label: 'Move left', fn: function () {
      commit('move column', function (s) { RM.moveScopeCol(s, key, -1); });
    } });
    if (idx < cols.length - 1) items2.push({ icon: 'arrow-right', label: 'Move right', fn: function () {
      commit('move column', function (s) { RM.moveScopeCol(s, key, 1); });
    } });
    if (!RM.SCOPE_BUILTIN_LABELS[key]) {
      items2.push({ icon: 'pencil', label: 'Rename…', fn: function () { scopeColModal(key); } });
    }
    items2.push({ sep: true });
    items2.push({ icon: 'trash-2', label: 'Remove column', fn: function () {
      commit('remove column', function (s) { RM.removeScopeCol(s, key); });
    } });
    openDropdown(mb, items2);
  });

  // create (key == null) or rename a custom scoping column
  function scopeColModal(key) {
    var cur = '';
    if (key) {
      state.meta.scopeCols.forEach(function (c) { if (c.key === key) cur = c.label || ''; });
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
          if (!v) { closeModal(); return; }
          closeModal();
          commit(key ? 'rename column' : 'add column', function (s) {
            if (key) s.meta.scopeCols.forEach(function (c) { if (c.key === key) c.label = v; });
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
    drag = { kind: 'pan', x0: e.clientX, y0: e.clientY, sl: board.scrollLeft, st: board.scrollTop, moved: false };
  });

  function startBarDrag(e, itemId, mode, barEl) {
    var it = RM.itemById(state, itemId);
    drag = {
      kind: 'bar', mode: mode, itemId: itemId,
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
    else if (drag.kind === 'phspan') phSpanDragMove(e, dx);
    else if (drag.kind === 'port') portDragMove(e);
    else if (drag.kind === 'scol') scolDragMove(e);
    else if (drag.kind === 'rfill') rfillMove(e);
    else if (drag.kind === 'bfill') bfillMove(e);
    else if (drag.kind === 'rrow') rrowMove(e);
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
    if (d.kind === 'port') { $('#tempLink').setAttribute('hidden', ''); }
    if (!d.moved) { return; }
    dragConsumedClick = true;
    dragEndAt = Date.now();

    if (d.kind === 'bar') barDragEnd(d, e);
    else if (d.kind === 'stbar') stBarDragEnd(d);
    else if (d.kind === 'ghost') ghostDragEnd(d, e);
    else if (d.kind === 'row') rowDragEnd(d, e);
    else if (d.kind === 'phspan') phSpanDragEnd(d);
    else if (d.kind === 'port') portDragEnd(d);
    else if (d.kind === 'scol') saveLocal();
    else if (d.kind === 'rfill') rfillEnd(d);
    else if (d.kind === 'bfill') bfillEnd(d);
    else if (d.kind === 'rrow') rrowEnd(d);
    else if (d.kind === 'pan') requestAnimationFrame(renderArrows);
  });

  function daysFromDx(dx) { return Math.round(dx / dayPx()); }
  // touched items align to the snap grid: position AND width become multiples
  function snapTo(d) { return Math.round(d / snapDays) * snapDays; }

  function barDragMove(e, dx) {
    var it = RM.itemById(state, drag.itemId);
    var dd = daysFromDx(dx);
    var ns = drag.start0, nd = drag.dur0, nr = drag.risk0;
    if (drag.mode === 'move') ns = Math.max(0, snapTo(drag.start0 + dd));
    else if (drag.mode === 'resize-r') nd = Math.max(snapDays, snapTo(drag.dur0 + dd));
    else if (drag.mode === 'resize-l') {
      ns = Math.max(0, Math.min(snapTo(drag.start0 + dd), drag.start0 + drag.dur0 - snapDays));
      nd = drag.dur0 + (drag.start0 - ns);
      if (nd < 1) nd = Math.max(1, snapDays);
    }
    drag.el.classList.add('dragging');
    document.body.classList.add('dragging-x');
    drag.el.style.left = (ns * dayPx()) + 'px';
    drag.el.style.width = (Math.max(6, nd * dayPx()) + nr * dayPx()) + 'px';
    drag.ns = ns; drag.nd = nd; drag.nr = nr;

    // vertical: dragging the bar across other rows reorders / re-phases the
    // item (same drop logic as dragging the row's left pane)
    if (drag.mode === 'move') {
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
    var endD = RM.spanEndDate(meta, ns, nd + nr);
    var work = RM.workInSpan(meta, ns, nd);
    var riskWork = RM.workInSpan(meta, ns + nd, nr);
    dragTip.hidden = false;
    dragTip.style.left = (e.clientX + 14) + 'px';
    dragTip.style.top = (e.clientY - 34) + 'px';
    dragTip.innerHTML = '<b>' + RM.fmtShort(RM.dayToDate(meta, ns)) + '</b> → ' + RM.fmtShort(endD) +
      ' · ' + fmtDays(work) + (it.size ? ' <b>(' + it.size + ')</b>' : '') +
      (nr > 0 ? ' + ' + fmtDays(riskWork) + ' risk' : '');
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
        var endDelta = (d.ns + d.nd + d.nr) - endBefore;
        rippleMoved = RM.shiftDependents(s, t.id, endDelta);
      }
      if (vr) applyDrop(s, d.itemId, vr);
      if (autoOrder) RM.sortItemsByStart(s);
    });
    if (rippleMoved) toast('Pushed ' + rippleMoved + ' dependent item(s) along');
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
    var ns = drag.start0, nd = drag.dur0;
    if (drag.mode === 'move') ns = Math.max(0, snapTo(drag.start0 + dd));
    else if (drag.mode === 'resize-r') nd = Math.max(snapDays, snapTo(drag.dur0 + dd));
    else if (drag.mode === 'resize-l') {
      ns = Math.max(0, Math.min(snapTo(drag.start0 + dd), drag.start0 + drag.dur0 - snapDays));
      nd = drag.dur0 + (drag.start0 - ns);
      if (nd < 1) nd = Math.max(1, snapDays);
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
    var day = Math.max(0, snapTo(laneDayAt(e.clientX)));
    var dur = RM.stretchSpan(state.meta, day, RM.effortDays(state, it));
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
      t.durDays = d.dur;
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
  $('#docTitle').addEventListener('change', function (e) {
    commit('title', function (s) { s.meta.title = e.target.value || 'Roadmap'; });
  });
  // the title edits only via its pencil — readonly otherwise, so header
  // clicks can't accidentally start a rename (and can drag the window).
  // While editing, the pencil becomes a checkmark that commits the rename.
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
    var t = $('#docTitle'), b = $('#titleEdit');
    if (on) t.removeAttribute('readonly');
    else t.setAttribute('readonly', '');
    b.innerHTML = '<i data-lucide="' + (on ? 'check' : 'pencil') + '"></i>';
    b.title = on ? 'Save name' : 'Rename roadmap';
    if (window.lucide) lucide.createIcons();
    sizeTitle();
    if (on) { t.focus(); t.select(); }
  }
  // pointerdown is swallowed so clicking the checkmark doesn't blur first
  // (which would flip the button back to a pencil before the click lands)
  $('#titleEdit').addEventListener('pointerdown', function (e) { e.preventDefault(); });
  $('#titleEdit').addEventListener('click', function () {
    var t = $('#docTitle');
    if (t.hasAttribute('readonly')) setTitleEditing(true);
    else t.blur(); // fires change → commit; the blur handler restores the pencil
  });
  $('#docTitle').addEventListener('blur', function () {
    setTitleEditing(false);
  });
  $('#docTitle').addEventListener('keydown', function (e) {
    if (e.key === 'Enter') e.target.blur();
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

  $('#viewTabs').addEventListener('click', function (e) {
    var b = e.target.closest('[data-view]');
    if (!b || b.dataset.view === view) return;
    view = b.dataset.view;
    saveLocal();
    render();
    if (view === 'planning') requestAnimationFrame(goToday);
  });

  function goToday() {
    var now = new Date();
    var d = RM.dateToDay(state.meta, new Date(Date.UTC(now.getFullYear(), now.getMonth(), now.getDate())));
    if (d != null) scrollLaneTo(d * dayPx());
  }

  function doAuto() {
    var r = RM.autoSchedule(state);
    confirmBox('Auto-schedule?',
      'Reschedules every unlocked item in non-backlog phases:<br>' +
      '· dependency order first<br>' +
      '· then earliest start with free capacity (roster of ' + state.team.length + ')<br>' +
      '· holiday weeks stretch bars, they don’t consume them<br>' +
      '· risk buffers follow each bar<br><br>' +
      '<b>' + r.changed + '</b> item(s) would move.' +
      (r.notes.length ? '<br><br>' + r.notes.slice(0, 4).map(esc).join('<br>') : ''),
      'Apply', function () {
        replaceState('auto-schedule', r.state);
        toast('Auto-scheduled — ' + r.changed + ' item(s) moved');
      });
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
  function menuItems(name) {
    if (name === 'file') {
      return [
        { icon: 'file', label: 'New blank roadmap', fn: function () {
          confirmBox('Start a blank roadmap?', 'Current roadmap stays in undo history.', 'New roadmap', function () {
            replaceState('blank', blankState());
          });
        } },
        { icon: 'folder-open', label: 'Open .xlsx…', fn: function () {
          if (window.HeadwayDesktop) HeadwayDesktop.openDialog();
          else $('#filePick').click();
        } },
        { sep: true },
        { icon: 'download', label: 'Save .xlsx', kbd: '⌘S', fn: function () { $('#btnSave').click(); } },
        window.HeadwayDesktop
          ? { icon: 'save', label: 'Save As…', fn: function () { window.HeadwayApp.save(true); } }
          : null,
        window.HeadwayDesktop
          ? { icon: 'timer-reset', label: 'Auto-save', checked: autoSave, fn: function () {
              autoSave = !autoSave;
              saveLocal(); renderTopbar();
              if (autoSave) scheduleAutoSave();
              toast('Auto-save ' + (autoSave ? 'on — writes to the open file' : 'off'));
            } }
          : null,
        { sep: true },
        { icon: 'image', label: 'Export PNG…', fn: function () { $('#btnExport').click(); } },
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
        { icon: 'plus', label: 'Add feature', fn: doAddFeature },
        { icon: 'plus', label: 'Add phase…', fn: function () { phaseModal(null); } },
        { sep: true },
        { icon: 'zap', label: 'Auto-schedule…', fn: doAuto },
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
    var snapItems = [[1, 'day'], [5, 'week'], [10, '2 weeks']].map(function (sn) {
      return { icon: 'magnet', label: 'Snap to ' + sn[1], checked: snapDays === sn[0], fn: function () {
        snapDays = sn[0];
        saveLocal(); renderTopbar();
        toast('Drag snap: ' + sn[1]);
      } };
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
      state.meta.capacityEnabled ? { icon: 'gauge', label: 'Capacity row', checked: showCap, fn: function () {
        showCap = !showCap;
        saveLocal(); render();
        toast('Capacity row ' + (showCap ? 'shown' : 'hidden'));
      } } : null,
      { sep: true },
      { icon: 'layers', label: 'Group by workstream', checked: groupWs, fn: function () {
        groupWs = !groupWs;
        saveLocal(); render();
      } },
      { icon: 'layers', label: 'Group by epic', checked: groupEpic, fn: function () {
        groupEpic = !groupEpic;
        saveLocal(); render();
      } },
      { sep: true },
      { icon: 'chevrons-up-down', label: 'Expand all features', fn: function () {
        state.items.forEach(function (it) { if (it.stories.length) expanded[it.id] = true; });
        render();
      } },
      { icon: 'chevrons-down-up', label: 'Collapse all features', fn: function () {
        expanded = {};
        render();
      } },
      { icon: 'arrow-down-narrow-wide', label: 'Auto-order rows by start', checked: autoOrder, fn: function () {
        autoOrder = !autoOrder;
        saveLocal();
        if (autoOrder) commit('auto-order', function (s) { RM.sortItemsByStart(s); });
        else renderTopbar();
        toast('Auto-order ' + (autoOrder ? 'on — rows follow the timeline' : 'off'));
      } },
      { sep: true },
      { icon: 'zoom-in', label: 'Zoom in', kbd: '⌘scroll', fn: function () { zoomBy(1.2); } },
      { icon: 'zoom-out', label: 'Zoom out', fn: function () { zoomBy(1 / 1.2); } },
      { icon: 'crosshair', label: 'Scroll to today', fn: goToday },
    ].filter(Boolean));
  }

  var openMenuName = null;
  function openMenu(b) {
    var name = b.dataset.menu;
    var items = menuItems(name);
    var r = b.getBoundingClientRect();
    var html = '<div class="menu-list">' + items.map(function (m, i) {
      if (m.sep) return '<div class="menu-sep"></div>';
      return '<button data-mi="' + i + '"' + (m.disabled ? ' disabled' : '') + (m.checked ? ' class="on"' : '') + '>' +
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
  board.addEventListener('wheel', function (e) {
    if (view !== 'planning' || (!e.ctrlKey && !e.metaKey)) return;
    e.preventDefault();
    var r = board.getBoundingClientRect();
    var laneX = board.scrollLeft + (e.clientX - r.left) - leftW();
    var day = laneX / dayPx();
    weekPx = Math.max(14, Math.min(80, weekPx * (e.deltaY < 0 ? 1.12 : 1 / 1.12)));
    saveLocal();
    render();
    board.scrollLeft = Math.max(0, day * dayPx() - (e.clientX - r.left - leftW()));
    requestAnimationFrame(renderArrows);
  }, { passive: false });

  function addFeature(phaseId) {
    var newId = RM.uid('i');
    commit('add feature', function (s) {
      var it = RM.normalizeState({
        meta: s.meta, phases: s.phases,
        items: [{ id: newId, num: RM.nextNum(s), phaseId: phaseId, feature: '', size: 'M', headcount: 1 }]
      }).items[0];
      it.id = newId;
      it.num = RM.nextNum(s);
      var lastIdx = -1;
      s.items.forEach(function (x, i2) { if (x.phaseId === phaseId) lastIdx = i2; });
      s.items.splice(lastIdx === -1 ? s.items.length : lastIdx + 1, 0, it);
      s.phases.forEach(function (p) { if (p.id === phaseId) p.collapsed = false; });
      selectedId = newId;
    });
    select(newId, true);
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
      '<div class="modal" style="width:460px">' +
      '<div class="m-head"><h2>' + (isNew ? 'New phase' : 'Edit phase') + '</h2>' +
      '<button class="p-close" data-m="x"><i data-lucide="x"></i></button></div>' +
      '<div class="m-body">' +
      '<div class="m-sec"><label>Name</label><input id="phName" style="width:100%" value="' + esc(phase ? phase.name : '') + '" placeholder="MVP: Measurement"></div>' +
      '<div class="m-sec"><label>Description</label><input id="phDesc" style="width:100%" value="' + esc(phase ? phase.description : '') + '"></div>' +
      '<div class="m-sec"><label class="p-check"><input type="checkbox" id="phBucket"' + (phase && phase.bucket ? ' checked' : '') + '> Backlog bucket (items parked here aren’t auto-scheduled)</label></div>' +
      '<div class="m-sec"><label>Dates</label><div class="p-grid2">' +
      '<div><label class="p-lab">Start</label><input type="date" id="phStart" style="width:100%" value="' +
        (phase && phase.startDay != null ? esc(RM.fmtISO(RM.dayToDate(state.meta, phase.startDay))) : '') + '"></div>' +
      '<div><label class="p-lab">End</label><input type="date" id="phEnd" style="width:100%" value="' +
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
      '<button data-m="x2">Cancel</button><button id="phSave" class="primary">' + (isNew ? 'Add phase' : 'Save') + '</button>' +
      '</div></div>',
      function (host) {
        $('[data-m=x]', host).onclick = closeModal;
        $('[data-m=x2]', host).onclick = closeModal;
        $('#phSave', host).onclick = function () {
          var name = $('#phName', host).value.trim() || 'Phase';
          var desc = $('#phDesc', host).value.trim();
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
        };
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
          return '<div class="val-item ' + cls + '" data-goto="' + (x.it ? x.it.id : '') + '" data-week="' + (x.v.week != null ? x.v.week : '') + '">' +
            '<span class="vi-id">' + (x.it ? '#' + x.it.num : '⧗') + '</span><span>' + esc(x.v.msg) +
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

  // budgeting rows render into the shared board: frozen left pane columns
  // (name · type · workstream · rate · cost · margin · total) + week cells
  var BU_COLS = { type: 96, ws: 108, cost: 76, rate: 76, margin: 54, total: 80 };
  function renderBudgetRows() {
    var meta = state.meta;
    var html = [];
    var sumTotal = 0;
    state.team.forEach(function (m) {
      var wsHex = RM.colorForWs(state, m.workstream);
      var cr = parseInt(wsHex.slice(0, 2), 16), cg = parseInt(wsHex.slice(2, 4), 16), cb = parseInt(wsHex.slice(4, 6), 16);
      var hours = RM.roleTotalHours(state, m);
      var total = hours * (m.rate || 0); // billing total: actual hours × rate
      sumTotal += total;
      var margin = RM.roleMargin(m);
      var cells = [];
      for (var w = 0; w < meta.numWeeks; w++) {
        var wh = RM.roleWeekHours(state, m, w);
        var a = Math.max(0, Math.min(1, wh.actual / 40));
        // clipped weeks (holidays) show the ACTUAL hours below in small text —
        // that's what totals and cost are built from
        var clipped = wh.actual < wh.planned;
        cells.push('<div class="bu-cell' + (clipped ? ' clipped' : '') + '" tabindex="0" data-w="' + w + '" data-iso="' + wh.iso +
          '" style="left:' + (w * weekPx) + 'px;width:' + weekPx +
          'px;background:rgba(' + cr + ',' + cg + ',' + cb + ',' + (a * 0.30).toFixed(3) + ')">' +
          (weekPx >= 20
            ? '<span>' + (wh.planned || '') + '</span>' + (clipped ? '<span class="bu-sub">(' + wh.actual + ')</span>' : '')
            : '') + '</div>');
      }
      html.push('<div class="row brole" data-mid="' + m.id + '">' +
        '<div class="row-left">' +
        '<span class="r-dot" style="background:#' + wsHex + '"></span>' +
        '<span class="bu-nm">' + esc(m.name) + '</span>' +
        '<span class="r-ws sc-chip bu-col" style="width:' + BU_COLS.type + 'px" tabindex="0" role="button" data-bact="type" title="Type — click to change">' + esc(shorten(m.type || '·', 13)) + '</span>' +
        '<span class="r-ws sc-chip bu-col" style="width:' + BU_COLS.ws + 'px" tabindex="0" role="button" data-bact="ws" title="Workstream — click to change">' +
        '<span class="dd-dot" style="background:#' + wsHex + '"></span>' +
        (m.workstream ? esc(shorten(m.workstream, 14)) : '·') + '</span>' +
        '<input class="bu-in bu-col" style="width:' + BU_COLS.cost + 'px" type="number" min="0" data-bud="cost" value="' + (m.cost || 0) + '" title="Cost (hourly)">' +
        '<input class="bu-in bu-col" style="width:' + BU_COLS.rate + 'px" type="number" min="0" data-bud="rate" value="' + (m.rate || 0) + '" title="Rate (hourly)">' +
        '<span class="bu-col bu-ro" style="width:' + BU_COLS.margin + 'px" title="Margin">' + (margin == null ? '—' : Math.round(margin) + '%') + '</span>' +
        '<span class="bu-col bu-ro" style="width:' + BU_COLS.total + 'px" title="Total — actual hours × rate">' + fmtMoney(total) + '</span>' +
        '</div><div class="row-lane">' + cells.join('') + '</div></div>');
    });
    if (state.team.length) {
      html.push('<div class="row brole btotal"><div class="row-left">' +
        '<span class="r-dot" style="background:transparent"></span>' +
        '<span class="bu-nm">Total</span>' +
        '<span class="bu-col" style="width:' + BU_COLS.type + 'px"></span><span class="bu-col" style="width:' + BU_COLS.ws + 'px"></span>' +
        '<span class="bu-col" style="width:' + BU_COLS.cost + 'px"></span><span class="bu-col" style="width:' + BU_COLS.rate + 'px"></span>' +
        '<span class="bu-col" style="width:' + BU_COLS.margin + 'px"></span>' +
        '<span class="bu-col bu-ro" style="width:' + BU_COLS.total + 'px">' + fmtMoney(sumTotal) + '</span>' +
        '</div><div class="row-lane"></div></div>');
    } else {
      html.push('<div class="row brole"><div class="row-left">' +
        '<span class="bu-nm" style="color:var(--ink-3)">No roles yet — add people in the Resources panel</span>' +
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
    var v = Math.max(0, parseFloat(e.target.value) || 0);
    var roleId = rowEl2.dataset.mid;
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

  // budget chips: type and workstream pick from the shared dropdown
  rowsEl.addEventListener('click', function (e) {
    if (view !== 'budget') return;
    var chip = e.target.closest('[data-bact]');
    var rowEl2 = e.target.closest('[data-mid]');
    if (!chip || !rowEl2 || rowEl2.classList.contains('btotal')) return;
    var roleId = rowEl2.dataset.mid;
    var m = null;
    state.team.forEach(function (x) { if (x.id === roleId) m = x; });
    if (!m) return;
    if (chip.dataset.bact === 'type') {
      openDropdown(chip, state.teamTypes.map(function (t) {
        return { label: esc(t), checked: m.type === t, fn: function () {
          commit('role type', function (s) {
            s.team.forEach(function (x) { if (x.id === roleId) x.type = t; });
          });
        } };
      }));
    } else {
      var wsItems = [{ label: '<i>— none —</i>', checked: !m.workstream, fn: function () {
        commit('role workstream', function (s) {
          s.team.forEach(function (x) { if (x.id === roleId) x.workstream = ''; });
        });
      } }];
      allWorkstreams().forEach(function (wv) {
        wsItems.push({
          label: esc(wv), dot: '#' + RM.colorForWs(state, wv), checked: m.workstream === wv,
          fn: function () {
            commit('role workstream', function (s) {
              s.team.forEach(function (x) { if (x.id === roleId) x.workstream = wv; });
            });
          }
        });
      });
      openDropdown(chip, wsItems);
    }
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
    var cur = mm.weekHours[iso] != null ? mm.weekHours[iso] : 40;
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
            if (h === 40) delete m2.weekHours[iso]; else m2.weekHours[iso] = h;
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

  function renderSetup() {
    var m = state.meta;
    var host = $('#setupView');

    var sizeInputs = RM.SIZE_ORDER.map(function (s2) {
      return '<div><label>' + s2 + '</label><input type="number" min="1" data-susz="' + s2 + '" value="' + m.sizeDays[s2] + '"></div>';
    }).join('');

    var holChips = (m.holidays || []).map(function (iso) {
      return '<span class="type-chip bw-chip">' + esc(iso) +
        '<button data-suholrm="' + esc(iso) + '" title="Remove"><i data-lucide="x"></i></button></span>';
    }).join('');

    function grip() {
      return '<span class="su-grip" title="Drag to reorder"><i data-lucide="grip-vertical"></i></span>';
    }
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
    var typeRows = state.teamTypes.map(function (t) {
      return '<div class="su-row" data-key="' + esc(t) + '">' + grip() +
        '<span class="su-name">' + esc(t) + '</span>' +
        '<span class="band-count">' + (typeCounts[t] || 0) + '</span>' +
        '<button data-suttrm="' + esc(t) + '" class="danger" title="Remove type"><i data-lucide="x"></i></button>' +
        '</div>';
    }).join('');

    host.innerHTML =
      '<div class="su-wrap">' +
      '<h1 class="su-title">Project setup</h1>' +

      '<div class="su-grid">' +
      '<section class="su-card"><h2>Timeline</h2>' +
      '<div class="p-grid2">' +
      '<div><label class="p-lab">Start (Monday)</label><input type="date" id="suStart" value="' + esc(m.timelineStart) + '" style="width:100%"></div>' +
      '<div><label class="p-lab">End (last working day)</label><input type="date" id="suEnd" value="' + esc(m.endDate || '') + '" style="width:100%"></div>' +
      '</div>' +
      '</section>' +

      '<section class="su-card"><h2>Sprint numbering</h2>' +
      '<div class="p-grid2">' +
      '<div><label class="p-lab">Sprint starts on (Monday)</label><input type="date" id="suAnchor" value="' + esc(m.sprintAnchor || m.timelineStart) + '" style="width:100%"></div>' +
      '<div><label class="p-lab">…and is sprint #</label><input type="number" id="suAnchorNum" step="1" value="' + (m.sprintAnchorNum != null ? m.sprintAnchorNum : 1) + '" style="width:100%"></div>' +
      '</div>' +
      '</section>' +

      '<section class="su-card"><h2>Workstreams</h2>' +
      '<div class="su-rows" data-sulist="ws">' + (wsRows || '<div class="m-hint">none yet</div>') + '</div>' +
      '<div class="p-row" style="margin-top:8px"><input id="suWsAdd" placeholder="New workstream…"><button id="suWsAddBtn" class="fixed">Add</button></div>' +
      '</section>' +

      '<section class="su-card"><h2>Phases</h2>' +
      '<div class="su-rows" data-sulist="phase">' + phaseRows + '</div>' +
      '<button id="suPhAdd" style="margin-top:8px"><i data-lucide="plus"></i> Add phase</button>' +
      '</section>' +

      '<section class="su-card"><h2>Team types</h2>' +
      '<div class="su-rows" data-sulist="type">' + typeRows + '</div>' +
      '<div class="p-row" style="margin-top:8px"><input id="suTypeAdd" placeholder="New type, e.g. Data Scientist"><button id="suTypeAddBtn" class="fixed">Add</button></div>' +
      '</section>' +

      '<section class="su-card"><h2>Sizing rules</h2>' +
      '<div class="size-grid">' + sizeInputs + '</div>' +
      '</section>' +

      '<section class="su-card"><h2>Capacity</h2>' +
      '<label class="p-check" title="Roster limits scheduling and validation; shows the capacity row"><input type="checkbox" id="suCapEnable"' + (m.capacityEnabled ? ' checked' : '') + '> Enable capacity planning</label>' +
      '</section>' +

      '<section class="su-card su-span2"><h2>Holidays</h2>' +
      '<div class="su-chips">' + (holChips || '<span class="m-hint">none</span>') + '</div>' +
      '<div class="p-row" style="margin-top:8px"><input type="date" id="suHolAdd" class="fixed"><button id="suHolAddBtn" class="fixed">Add day</button></div>' +
      '</section>' +
      '</div></div>';
    if (window.lucide) lucide.createIcons();
  }

  $('#setupView').addEventListener('change', function (e) {
    var t = e.target;
    if (t.id === 'suCapEnable') {
      var on = t.checked;
      commit('capacity feature', function (s2) { s2.meta.capacityEnabled = on; });
      toast('Capacity planning ' + (on ? 'enabled' : 'disabled'));
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
      var sz = t.dataset.susz;
      var v = Math.max(1, parseInt(t.value, 10) || RM.DEFAULT_SIZE_DAYS[sz]);
      commit('size days', function (s2) { s2.meta.sizeDays[sz] = v; });
      return;
    }
  });

  $('#setupView').addEventListener('click', function (e) {
    var t = e.target.closest('button');
    if (!t) return;
    if (t.id === 'suHolAddBtn') {
      var v = $('#suHolAdd').value;
      if (!v) return;
      var d = RM.parseISO(v);
      if (!isFinite(d.getTime())) return;
      if (((d.getUTCDay() + 6) % 7) > 4) { toast('That date falls on a weekend — no working day to blank', 'err'); return; }
      var iso = RM.fmtISO(d);
      commit('holiday add', function (s2) {
        if (s2.meta.holidays.indexOf(iso) === -1) s2.meta.holidays.push(iso);
        s2.meta.holidays.sort();
      });
      return;
    }
    var holrm = t.dataset.suholrm;
    if (holrm) {
      commit('holiday rm', function (s2) {
        s2.meta.holidays = s2.meta.holidays.filter(function (x) { return x !== holrm; });
      });
      return;
    }
    if (t.id === 'suWsAddBtn') {
      var wv = $('#suWsAdd').value.trim();
      if (!wv) return;
      commit('add workstream', function (s2) {
        if (!s2.wsColors[wv]) s2.wsColors[wv] = RM.DEFAULT_WS_COLORS[wv] || RM.PALETTE.product;
      });
      return;
    }
    if (t.dataset.suwsedit) { wsEditModal(t.dataset.suwsedit); return; }
    if (t.id === 'suPhAdd') { phaseModal(null); return; }
    if (t.dataset.suphedit) { phaseModal(t.dataset.suphedit); return; }
    if (t.dataset.suphdel) { deletePhaseConfirm(t.dataset.suphdel); return; }
    if (t.id === 'suTypeAddBtn') {
      var tv = $('#suTypeAdd').value.trim();
      if (!tv) return;
      if (state.teamTypes.indexOf(tv) !== -1) { toast('Type already exists'); return; }
      commit('add type', function (s2) { s2.teamTypes.push(tv); });
      return;
    }
    var ttrm = t.dataset.suttrm;
    if (ttrm) {
      var used = state.team.some(function (mm) { return mm.type === ttrm; }) ||
        state.items.some(function (x) { return x.teamType === ttrm; });
      if (used) { toast('“' + ttrm + '” is in use by roles or features', 'err'); return; }
      commit('remove type', function (s2) {
        s2.teamTypes = s2.teamTypes.filter(function (x) { return x !== ttrm; });
      });
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
    $('#setupView').addEventListener('pointerdown', function (e) {
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
        } else if (dd.kind === 'ws') {
          s2.wsOrder = moveKeyBefore(allWorkstreams(s2), dd.key, dd.before);
        } else if (dd.kind === 'phase') {
          var ids = s2.phases.map(function (p2) { return p2.id; });
          var order = moveKeyBefore(ids, dd.key, dd.before);
          s2.phases.sort(function (a, b) { return order.indexOf(a.id) - order.indexOf(b.id); });
        }
      });
    });
  })();

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
      '<b>Bars</b> — drag to move · edges resize · <span class="kbd">⌘drag</span> pushes dependents along · <span class="kbd">←</span><span class="kbd">→</span> nudge (<span class="kbd">⇧</span> = 1 week) · snap grid in View<br>' +
      '<b>Sizes in weeks</b> — XS 2d · S 1w · M 2w · L 4w · XL 8w; the risk buffer uses the same scale (panel), shown as one duration<br>' +
      '<b>Rows</b> — drag the left pane to reorder / move phase (auto-order can re-sort by start) · right-click for insert/delete · chips cycle on click<br>' +
      '<b>Links</b> — drag a bar’s edge circles onto another row (left = depends ON it, right = dependency FOR it; Esc cancels) · click an arrow, Delete removes · orange = critical path<br>' +
      '<b>Timeline</b> — drag empty space to pan · <span class="kbd">⌘scroll</span> zooms · click a capacity cell to toggle a holiday week; single dates + sprint numbering (e.g. S1 = Sep 7) in Settings<br>' +
      '<b>Capacity row</b> — ≈ parallel work items (hours ÷ 40) for the selected work type; filter via its dropdown<br>' +
      '<b>Resources panel</b> — bottom, resizable/collapsible; hours per person per week (default 40) — click a cell to type, drag to fill; drag the grip to reorder people<br>' +
      '<b>Auto</b> — dependency-ordered, capacity-aware schedule; locked items stay put<br>' +
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
        var cls = h > RM.WEEK_HOURS ? ' rh-over' : '';
        // white at 0h → light blue at 20h → blue at 40h (kept light enough
        // that the black hour label stays readable)
        var t2 = Math.max(0, Math.min(1, h / RM.WEEK_HOURS));
        var bg = '#' + RM.tint('7FAEDD', 1 - t2);
        cells.push('<div class="rh' + cls + '" data-w="' + w + '" style="left:' + (w * weekPx) +
          'px;width:' + weekPx + 'px;background:' + bg + '" title="' + esc(m.name + ' — week of ' +
            RM.fmtShort(RM.weekStartDate(meta, w)) + ': ' + fmtH(h) + 'h') + '">' +
          (weekPx >= 20 ? fmtH(h) : '') + '</div>');
      }
      html.push(
        '<div class="rrow" data-mid="' + m.id + '">' +
        '<div class="rleft">' +
        '<span class="r-grip rr-grip" title="Drag to reorder"><i data-lucide="grip-vertical"></i></span>' +
        '<span class="res-name" data-rname="' + m.id + '" title="Role — click to rename">' + esc(m.name) + '</span>' +
        '<button class="res-type" data-rtype="' + m.id + '" title="Work type">' + esc(m.type) + '</button>' +
        '<button class="res-type res-wschip" data-rws="' + m.id + '" title="Workstream (optional)">' +
        (m.workstream ? esc(shorten(m.workstream, 12)) : '—') + '</button>' +
        (state.meta.capacityEnabled
          ? '<span class="res-cap" tabindex="0" role="button" data-rcap="' + m.id +
            '" title="Capacity at 40 h — click to edit">' + fmtPe(m.capacity != null ? m.capacity : 1) + '×</span>'
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

  $('#hdrCap').addEventListener('keydown', function (e) {
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
      lrs = { x0: e.clientX, w0: view === 'scoping' ? leftWScope : view === 'budget' ? leftWBudget : leftWPlan };
      e.preventDefault();
      e.stopPropagation();
    }
    $('#leftRz').addEventListener('pointerdown', startLeftRz);
    $('#leftRzLine').addEventListener('pointerdown', startLeftRz);
    window.addEventListener('pointermove', function (e) {
      if (!lrs) return;
      var w = Math.max(240, Math.min(window.innerWidth * 0.7, lrs.w0 + (e.clientX - lrs.x0)));
      if (view === 'scoping') leftWScope = Math.round(w);
      else if (view === 'budget') leftWBudget = Math.round(w);
      else leftWPlan = Math.round(w);
      document.documentElement.style.setProperty('--left-w', Math.round(w) + 'px');
    });
    window.addEventListener('pointerup', function () {
      if (lrs) { lrs = null; saveLocal(); render(); }
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
    inp.placeholder = 'Role — e.g. Senior Dev (Alice)…';
    inp.className = 'radd-input';
    lab.replaceWith(inp);
    inp.focus();
    var done = false;
    function finish(saveIt) {
      if (done) return; done = true;
      var v = inp.value.trim();
      if (saveIt && v) {
        commit('add person', function (s) {
          s.team.push({ id: RM.uid('t'), name: v, type: RM.DEFAULT_WORK_TYPE, weekHours: {} });
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
    confirmBox('Remove “' + esc(dm.name) + '”?', 'Their weekly hours disappear from capacity.', 'Remove', function () {
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
      '<label>Start<input type="date" id="rdStart" value="' + sv + '"></label>' +
      '<label>End<input type="date" id="rdEnd" value="' + ev + '"></label>' +
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
    var mid = rrow.dataset.mid;
    var m = null;
    state.team.forEach(function (x) { if (x.id === mid) m = x; });
    if (!m) return;
    var cx = e.clientX, cy = e.clientY;
    openContextMenu(cx, cy, [
      { icon: 'pencil', label: 'Rename', fn: function () {
        var el = resGrid.querySelector('.rrow[data-mid="' + mid + '"] [data-rname]');
        if (el) el.dispatchEvent(new MouseEvent('click', { bubbles: true }));
      } },
      { icon: 'tags', label: 'Work type…', fn: function () {
        openContextMenu(cx, cy, state.teamTypes.map(function (t) {
          return { label: esc(t), checked: m.type === t, fn: function () {
            commit('member type', function (s) {
              s.team.forEach(function (x) { if (x.id === mid) x.type = t; });
            });
          } };
        }));
      } },
      { icon: 'layers', label: 'Workstream…', fn: function () {
        var wsItems = [{ label: '<i>— none —</i>', checked: !m.workstream, fn: function () {
          commit('role workstream', function (s) {
            s.team.forEach(function (x) { if (x.id === mid) x.workstream = ''; });
          });
        } }];
        allWorkstreams().forEach(function (w) {
          wsItems.push({ label: esc(w), dot: '#' + RM.colorForWs(state, w), checked: m.workstream === w, fn: function () {
            commit('role workstream', function (s) {
              s.team.forEach(function (x) { if (x.id === mid) x.workstream = w; });
            });
          } });
        });
        openContextMenu(cx, cy, wsItems);
      } },
      state.meta.capacityEnabled ? { icon: 'gauge', label: 'Capacity…', fn: function () {
        var el = resGrid.querySelector('.rrow[data-mid="' + mid + '"] [data-rcap]');
        if (el) el.dispatchEvent(new MouseEvent('click', { bubbles: true }));
      } } : null,
      { icon: 'calendar-range', label: 'Start / end dates…', fn: function () { roleDatesPopover(cx, cy, mid); } },
      { sep: true },
      { icon: 'trash-2', label: 'Remove role…', fn: function () { deleteRoleConfirm(mid); } }
    ].filter(Boolean));
  });

  // role rename (inline)
  resGrid.addEventListener('click', function (e) {
    var nameEl = e.target.closest('[data-rname]');
    if (!nameEl || nameEl.querySelector('input')) return;
    var nmid = nameEl.dataset.rname;
    startInlineEdit(nameEl, function (v) {
      commit('rename role', function (s) {
        s.team.forEach(function (m) { if (m.id === nmid) m.name = v; });
      });
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

  // member workstream chip → shared dropdown
  resGrid.addEventListener('click', function (e) {
    var chip = e.target.closest('[data-rws]');
    if (!chip) return;
    var mid = chip.dataset.rws;
    var m = null;
    state.team.forEach(function (x) { if (x.id === mid) m = x; });
    if (!m) return;
    var wsList = allWorkstreams();
    var items = [{ label: '<i>— none —</i>', checked: !m.workstream, fn: function () {
      commit('role workstream', function (s) {
        s.team.forEach(function (x) { if (x.id === mid) x.workstream = ''; });
      });
    } }];
    wsList.forEach(function (w) {
      items.push({
        label: esc(w),
        dot: '#' + RM.colorForWs(state, w),
        checked: m.workstream === w,
        fn: function () {
          commit('role workstream', function (s) {
            s.team.forEach(function (x) { if (x.id === mid) x.workstream = w; });
          });
        },
        edit: function () { wsEditModal(w); }
      });
    });
    openDropdown(chip, items);
  });

  // member type chip → shared dropdown
  resGrid.addEventListener('click', function (e) {
    var chip = e.target.closest('[data-rtype]');
    if (!chip) return;
    var mid = chip.dataset.rtype;
    var m = null;
    state.team.forEach(function (x) { if (x.id === mid) m = x; });
    if (!m) return;
    openDropdown(chip, state.teamTypes.map(function (t) {
      return { label: esc(t), checked: m.type === t, fn: function () {
        commit('member type', function (s) {
          s.team.forEach(function (x) { if (x.id === mid) x.type = t; });
        });
      } };
    }));
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
    var btn = $('#btnSave');
    savingNow = true;
    btn.disabled = true; btn.textContent = 'Saving…';
    function restoreBtn() {
      savingNow = false;
      btn.dataset.mode = '';
      updateSaveBtn();
    }
    RMExcel.exportWorkbook(state, uiSnapshot()).then(function (blob) {
      var name = saveFileName();
      if (window.HeadwayDesktop) { // desktop: write straight to disk
        return HeadwayDesktop.saveBlob(blob, name, forceDialog).then(function (path) {
          if (!path) return; // dialog canceled
          lastExport = new Date().toTimeString().slice(0, 5);
          docSaved = true;
          saveLocal();
          if (!quiet) toast('Saved ' + HeadwayDesktop.basename(path));
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

  function loadWorkbookBuffer(buf, name, quiet) {
    return RMExcel.importWorkbook(buf).then(function (r) {
      if (r.ui) applyUi(r.ui); // the file carries the browser prefs too
      replaceState('open', r.state);
      docSaved = true; // fresh from disk — matches its file
      updateSaveBtn();
      selectedId = null;
      if (!quiet) {
        toast(r.source === 'tool'
          ? 'Loaded “' + name + '” (full tool state)'
          : 'Parsed “' + name + '” from the template layout');
      }
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
    menuItems: menuItems
  };

  // ------------------------------------------------------------ keyboard
  window.addEventListener('keydown', function (e) {
    var ae = document.activeElement;
    // contenteditable editors (rich description, scoping cells) count as fields too
    var inField = /INPUT|TEXTAREA|SELECT/.test(ae.tagName) ||
      ae.isContentEditable || (ae.closest && ae.closest('[contenteditable="true"]') !== null);
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
      if (selectedId) { select(null); return; }
      if (presentMode) { setPresent(false); return; }
      return;
    }
    if ((e.metaKey || e.ctrlKey) && e.key.toLowerCase() === 'f' &&
        (view === 'planning' || view === 'scoping')) {
      e.preventDefault();
      var ff = $('#rowFilter');
      ff.focus(); ff.select();
      return;
    }
    if (inField) return;
    var mod = e.metaKey || e.ctrlKey;
    if (mod && e.key.toLowerCase() === 'z') {
      e.preventDefault();
      if (e.shiftKey) redo(); else undo();
      return;
    }
    if (mod && e.key.toLowerCase() === 'y') { e.preventDefault(); redo(); return; }
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
      var it2 = RM.itemById(state, selectedId);
      if (it2 && isScheduled(it2) && !it2.locked) {
        e.preventDefault();
        var delta = (e.key === 'ArrowRight' ? 1 : -1) * (e.shiftKey ? 5 : 1);
        commit('nudge', function (s) {
          var t = RM.itemById(s, selectedId);
          var ns2 = Math.max(0, t.startDay + delta);
          RM.shiftStories(t, ns2 - t.startDay);
          t.startDay = ns2;
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

  // timeline-only preview: hides the topbar and the frozen left pane; the
  // floating minimize button (or Esc) restores the full UI
  function setPresent(on) {
    presentMode = on;
    $('#btnPresentExit').hidden = !on;
    render();
  }
  $('#btnPresent').addEventListener('click', function () { setPresent(true); });
  $('#btnPresentExit').addEventListener('click', function () { setPresent(false); });

  // ------------------------------------------------------------ png export
  function exportModal() {
    var meta = state.meta;
    var sprints = [];
    for (var w = 0; w < meta.numWeeks; w++) {
      var n = RM.sprintNumForWeek(meta, w);
      if (!sprints.length || sprints[sprints.length - 1].num !== n)
        sprints.push({ num: n, date: RM.fmtShort(RM.dayToDate(meta, w * 5)) });
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
      return sprints.map(function (s) {
        return '<option value="' + s.num + '"' + (s.num === selNum ? ' selected' : '') +
          '>S' + s.num + ' · ' + esc(s.date) + '</option>';
      }).join('');
    }
    function nameOpts(values, allLabel) {
      return '<option value="">' + allLabel + '</option>' + values.map(function (v) {
        return '<option value="' + esc(v) + '">' + esc(v) + '</option>';
      }).join('');
    }
    openModal(
      '<div class="modal" style="width:440px">' +
      '<div class="m-head"><h2>Export PNG</h2>' +
      '<button class="p-close" data-m="x"><i data-lucide="x"></i></button></div>' +
      '<div class="m-body">' +
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
      '<div class="m-sec"><label>Options</label><div class="p-row">' +
      '<select id="exScale"><option value="1">1× scale</option><option value="2" selected>2× scale</option></select>' +
      '<label class="p-check"><input type="checkbox" id="exArrows"> Dependency arrows</label>' +
      '</div><div class="m-hint">The exported dates are bounded by the range AND the filtered bars — empty weeks at either end are trimmed.</div></div>' +
      '</div>' +
      '<div class="m-foot"><button data-m="cancel">Cancel</button>' +
      '<button id="exGo" class="primary"><i data-lucide="image"></i>Export PNG</button></div></div>',
      function (host) {
        $('[data-m=x]', host).onclick = closeModal;
        $('[data-m=cancel]', host).onclick = closeModal;
        $('#exGo', host).onclick = function () {
          RM_EXPORT.download(state, {
            fromSprint: parseInt($('#exFrom', host).value, 10),
            toSprint: parseInt($('#exTo', host).value, 10),
            ws: $('#exWs', host).value || null,
            epic: $('#exEpic', host).value || null,
            phaseId: $('#exPhase', host).value || null,
            scale: parseInt($('#exScale', host).value, 10),
            arrows: $('#exArrows', host).checked
          });
          closeModal();
        };
      });
  }
  $('#btnExport').addEventListener('click', exportModal);

  // ------------------------------------------------------------ boot
  function boot() {
    state = loadLocal() || blankState(); // fresh installs start completely empty
    validation = RM.validate(state);
    // features with stories start expanded so the story timelines are visible
    // (unless a saved expansion map — local or from an imported file — says otherwise)
    if (!uiExpandedLoaded) state.items.forEach(function (it) { if (it.stories.length) expanded[it.id] = true; });
    render();
    // land at today on first paint
    if (view === 'planning') requestAnimationFrame(goToday);
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
    templateState: templateState
  };
})();
