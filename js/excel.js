/*
 * Headway Excel round-trip (ExcelJS).
 *  - exportWorkbook(state) -> Promise<Blob>: a "Roadmap" sheet in the source
 *    template's shape (band rows, sprint header, colored bar fills), plus
 *    Stories, Team, and a hidden _RoadmapTool sheet carrying lossless JSON.
 *  - importWorkbook(arrayBuffer) -> Promise<{state, source}>: prefers
 *    _RoadmapTool; otherwise parses a template-shaped Roadmap sheet.
 */
(function (root) {
  'use strict';

  var RM = root.RM || (typeof require !== 'undefined' ? require('./core.js') : null);
  var RMExcel = {};

  var FIRST_SPRINT_COL = 12; // L
  var GRAY = 'F2F2F2';
  var BAND = '1A1A1A';
  var HEADER_BLUES = { sprint: 'B8D0ED', mvp: '71A1DC', next: '3172C4', future: '265A99' };
  var STATE_VERSION = 'roadmapper-state-v1';
  var CHUNK = 30000;

  function ExcelJSRef() {
    var E = root.ExcelJS || (typeof require !== 'undefined' ? tryRequireExcel() : null);
    if (!E) throw new Error('ExcelJS is not loaded');
    return E;
  }
  function tryRequireExcel() {
    try { return require('exceljs'); } catch (e) { return null; }
  }

  function solid(argbHex) {
    return { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF' + argbHex } };
  }

  // \u-escape every non-ASCII char: the result is still valid JSON, and
  // pure-ASCII chunks are immune to the surrogate-pair corruption ExcelJS
  // exhibits at certain in-cell offsets (splitting mid-escape is fine —
  // concatenation restores it before JSON.parse).
  function asciiJson(obj) {
    return JSON.stringify(obj).replace(/[\u007F-\uFFFF]/g, function (ch) {
      return '\\u' + ('0000' + ch.charCodeAt(0).toString(16)).slice(-4);
    });
  }
  // The canonical serialized document: the EXACT string exportWorkbook embeds
  // in the hidden sheet. Two files carry the same document iff these match.
  RMExcel.stateJsonOf = function (state) { return asciiJson(state); };

  // The embedded document JSON of a workbook (the string stateJsonOf wrote),
  // or null for foreign/template files with no valid _RoadmapTool sheet.
  // Lets callers tell a container rewrite (sync clients re-zip xlsx files
  // without touching cell content) from a real document change.
  RMExcel.readStateJson = function (arrayBuffer) {
    var ExcelJS = ExcelJSRef();
    var wb = new ExcelJS.Workbook();
    return wb.xlsx.load(arrayBuffer).then(function () {
      var hws = wb.getWorksheet('_RoadmapTool');
      if (!hws || cellText(hws.getCell('A1')) !== STATE_VERSION) return null;
      var chunks = [];
      var count = parseInt(cellText(hws.getCell('A2')), 10) || 0;
      for (var i = 0; i < count; i++) chunks.push(cellRaw(hws.getCell(2 + i, 2)));
      var json = chunks.join('');
      try { JSON.parse(json); } catch (e) { return null; }
      return json;
    });
  };

  function depCellText(it) {
    var parts = it.deps.map(String);
    (it.depsText || []).forEach(function (t) { parts.push(t); });
    return parts.length ? parts.join(', ') : 'None';
  }

  // ------------------------------------------------------------------ export
  // ui (optional): the app's UI-preferences snapshot — zoom, view, grouping,
  // column widths, expansion… Stored beside the state so opening the file on
  // another machine restores the exact browser state.
  RMExcel.exportWorkbook = function (state, ui) {
    var ExcelJS = ExcelJSRef();
    var wb = new ExcelJS.Workbook();
    wb.creator = 'Headway';
    wb.created = new Date();

    // one column per WEEK (import infers the granularity from the 7-day gaps
    // between the header dates); row 2 still shows sprint numbers, merged
    // across each sprint's weeks
    var meta = state.meta;
    var wps = meta.weeksPerSprint || 2;
    var numWeeks = meta.numWeeks;
    var S5 = RM.slotsOf(meta); // slots (working days) per index week
    var lastWeekCol = FIRST_SPRINT_COL + numWeeks - 1;
    var nextCol = lastWeekCol + 1;
    var futureCol = lastWeekCol + 2;
    var extraCols = {
      headcount: futureCol + 1,
      teamType: futureCol + 2,
      start: futureCol + 3,
      end: futureCol + 4,
      status: futureCol + 5
    };
    var lastCol = extraCols.status;

    var ws = wb.addWorksheet('Roadmap', {
      views: [{ state: 'frozen', xSplit: 4, ySplit: 3 }]
    });

    var blackoutWks = {};
    for (var s = 0; s < numWeeks; s++) {
      if (RM.isBlackoutWeek(meta, s)) blackoutWks[s] = true;
    }

    // ---- row 1: phase coverage bands over sprint columns
    var r1 = ws.getRow(1);
    r1.getCell(1).value = 'Phase';
    var covered = {};
    state.phases.forEach(function (p) {
      if (p.bucket) return;
      var minW = null, maxW = null;
      RM.itemsInPhase(state, p.id).forEach(function (it) {
        if (it.startDay == null || it.durDays == null) return;
        var w0 = Math.floor(it.startDay / S5), w1 = Math.floor((it.startDay + it.durDays - 1) / S5);
        if (minW == null || w0 < minW) minW = w0;
        if (maxW == null || w1 > maxW) maxW = w1;
      });
      if (minW == null) return;
      var s0 = Math.max(0, minW), s1 = Math.min(numWeeks - 1, maxW);
      var free0 = null;
      for (var si = s0; si <= s1; si++) {
        if (!covered[si] && free0 == null) free0 = si;
        if (covered[si] && free0 != null) { s1 = si - 1; break; }
      }
      if (free0 == null) return;
      for (var sj = free0; sj <= s1; sj++) covered[sj] = true;
      var c0 = FIRST_SPRINT_COL + free0, c1 = FIRST_SPRINT_COL + s1;
      if (c1 > c0) ws.mergeCells(1, c0, 1, c1);
      var cell = ws.getCell(1, c0);
      cell.value = p.name;
      cell.fill = solid(HEADER_BLUES.mvp);
      cell.font = { bold: true, color: { argb: 'FF17324F' } };
      cell.alignment = { horizontal: 'center' };
    });
    ws.getCell(1, nextCol).fill = solid(HEADER_BLUES.next);
    ws.getCell(1, futureCol).fill = solid(HEADER_BLUES.future);

    // ---- row 2: sprint numbers, merged over each sprint's weeks
    var r2 = ws.getRow(2);
    r2.getCell(1).value = 'Sprint';
    r2.getCell(11).value = 'Sprint';
    for (s = 0; s * wps < numWeeks; s++) {
      var c0s = FIRST_SPRINT_COL + s * wps;
      var c1s = Math.min(FIRST_SPRINT_COL + (s + 1) * wps - 1, lastWeekCol);
      if (c1s > c0s) ws.mergeCells(2, c0s, 2, c1s);
      var c2 = r2.getCell(c0s);
      c2.value = s + 1;
      c2.fill = solid(HEADER_BLUES.sprint);
      c2.alignment = { horizontal: 'center' };
    }
    r2.getCell(nextCol).fill = solid(HEADER_BLUES.next);
    r2.getCell(futureCol).fill = solid(HEADER_BLUES.future);

    // ---- row 3: column headers + sprint start dates
    var headers = ['ID', 'Workstream', 'Epic', 'Feature', 'Enables',
      'Out of Scope (Non-Exhaustive)', 'Notes', 'Dependencies',
      'External Dependencies (High Level)', 'T-Shirt Size', 'Dependency Risk/Size'];
    var r3 = ws.getRow(3);
    headers.forEach(function (h, i) {
      var c = r3.getCell(i + 1);
      c.value = h;
      c.font = { bold: true };
      c.alignment = { wrapText: true, vertical: 'top' };
    });
    for (s = 0; s < numWeeks; s++) {
      var dc = r3.getCell(FIRST_SPRINT_COL + s);
      dc.value = RM.weekStartDate(meta, s);
      dc.numFmt = 'm/d/yy';
      dc.alignment = { textRotation: 45 };
      if (blackoutWks[s]) dc.fill = solid(GRAY);
    }
    r3.getCell(nextCol).value = 'Next';
    r3.getCell(futureCol).value = 'Future';
    r3.getCell(nextCol).font = { bold: true };
    r3.getCell(futureCol).font = { bold: true };
    r3.getCell(extraCols.headcount).value = 'Headcount';
    r3.getCell(extraCols.teamType).value = 'Team Type';
    r3.getCell(extraCols.start).value = 'Start';
    r3.getCell(extraCols.end).value = 'End';
    r3.getCell(extraCols.status).value = 'Status';
    Object.keys(extraCols).forEach(function (k) {
      r3.getCell(extraCols[k]).font = { bold: true, italic: true };
    });

    // ---- column widths
    ws.getColumn(1).width = 5;
    ws.getColumn(2).width = 12;
    ws.getColumn(3).width = 11;
    ws.getColumn(4).width = 42;
    ws.getColumn(5).width = 44;
    ws.getColumn(6).width = 32;
    ws.getColumn(7).width = 32;
    ws.getColumn(8).width = 13;
    ws.getColumn(9).width = 32;
    ws.getColumn(10).width = 9;
    ws.getColumn(11).width = 9;
    for (s = 0; s < numWeeks; s++) ws.getColumn(FIRST_SPRINT_COL + s).width = 4;
    ws.getColumn(nextCol).width = 6;
    ws.getColumn(futureCol).width = 7;
    Object.keys(extraCols).forEach(function (k) { ws.getColumn(extraCols[k]).width = 10; });

    // ---- body
    var rowIdx = 4;
    var phaseById = {};
    state.phases.forEach(function (p) { phaseById[p.id] = p; });

    state.phases.forEach(function (phase) {
      var band = ws.getRow(rowIdx);
      band.getCell(1).value = phase.name.toUpperCase();
      if (phase.description) band.getCell(4).value = phase.description;
      for (var c = 1; c <= lastCol; c++) {
        band.getCell(c).fill = solid(BAND);
        band.getCell(c).font = { bold: true, color: { argb: 'FFFFFFFF' } };
      }
      band.height = 18;
      rowIdx += 1;

      RM.itemsInPhase(state, phase.id).forEach(function (it) {
        var r = ws.getRow(rowIdx);
        r.getCell(1).value = it.num;
        r.getCell(2).value = it.workstream || null;
        r.getCell(3).value = it.epic || null;
        r.getCell(4).value = it.feature;
        // scope fields hold rich HTML — the sheet gets flattened plain text
        // (the hidden tool sheet keeps the formatted version losslessly)
        r.getCell(5).value = RM.htmlToText(it.enables || '') || null;
        r.getCell(6).value = RM.htmlToText(it.outOfScope || '') || null;
        r.getCell(7).value = RM.htmlToText(it.notes || '') || null;
        r.getCell(8).value = depCellText(it);
        r.getCell(9).value = RM.htmlToText(it.extDeps || '') || null;
        r.getCell(10).value = it.size || null;
        r.getCell(11).value = it.risk || null;
        [4, 5, 6, 7, 9].forEach(function (c) {
          r.getCell(c).alignment = { wrapText: true, vertical: 'top' };
        });

        var color = RM.colorForItem(state, it);
        var scheduled = it.startDay != null && it.durDays != null;
        if (scheduled && it.milestone) {
          // milestone: a diamond in its week, start = end = the fixed date
          var msWk = Math.max(0, Math.min(numWeeks - 1, Math.floor(it.startDay / S5)));
          var msCell = r.getCell(FIRST_SPRINT_COL + msWk);
          msCell.value = { diamond: '◆', star: '★', circle: '●' }[RM.msStyleOf(it)];
          msCell.font = { color: { argb: 'FF' + color }, bold: true };
          msCell.alignment = { horizontal: 'center' };
          r.getCell(extraCols.start).value = RM.dayToDate(meta, it.startDay);
          r.getCell(extraCols.start).numFmt = 'm/d/yy';
          r.getCell(extraCols.end).value = RM.dayToDate(meta, it.startDay);
          r.getCell(extraCols.end).numFmt = 'm/d/yy';
        } else if (scheduled) {
          // solid = work span; pale tint = trailing risk buffer (per WEEK)
          var totalSpan = it.durDays + (it.riskDays || 0);
          var s0 = Math.max(0, Math.floor(it.startDay / S5));
          var sWork = Math.min(numWeeks - 1, Math.floor((it.startDay + it.durDays - 1) / S5));
          var s1 = Math.min(numWeeks - 1, Math.floor((it.startDay + totalSpan - 1) / S5));
          for (var sp = s0; sp <= s1; sp++) {
            var cell = r.getCell(FIRST_SPRINT_COL + sp);
            cell.fill = solid(sp > sWork ? RM.tint(color, 0.75) : color);
          }
          r.getCell(extraCols.start).value = RM.dayToDate(meta, it.startDay);
          r.getCell(extraCols.start).numFmt = 'm/d/yy';
          r.getCell(extraCols.end).value = RM.spanEndDate(meta, it.startDay, totalSpan);
          r.getCell(extraCols.end).numFmt = 'm/d/yy';
        } else if (phase.bucket) {
          var mcol = /future/i.test(phase.name) ? futureCol : nextCol;
          r.getCell(mcol).fill = solid(mcol === futureCol ? HEADER_BLUES.future : HEADER_BLUES.next);
        }
        // gray blackout columns where no bar was painted
        for (sp = 0; sp < numWeeks; sp++) {
          if (!blackoutWks[sp]) continue;
          var bc = r.getCell(FIRST_SPRINT_COL + sp);
          if (!bc.fill || bc.fill.pattern !== 'solid') bc.fill = solid(GRAY);
        }
        r.getCell(extraCols.headcount).value = it.headcount;
        r.getCell(extraCols.teamType).value = it.teamType || null;
        r.getCell(extraCols.status).value = it.done ? 'Done' : (it.locked ? 'Locked' : null);
        rowIdx += 1;
      });
    });

    // ---- Stories sheet
    var sws = wb.addWorksheet('Stories');
    sws.getRow(1).values = ['Item #', 'Feature', 'Story', 'Done', 'Description', 'Acceptance Criteria'];
    sws.getRow(1).font = { bold: true };
    sws.getColumn(1).width = 8;
    sws.getColumn(2).width = 44;
    sws.getColumn(3).width = 60;
    sws.getColumn(4).width = 7;
    sws.getColumn(5).width = 50;
    sws.getColumn(6).width = 50;
    var srow = 2;
    state.items.forEach(function (it) {
      it.stories.forEach(function (st) {
        // rich text flattens to plain text in the sheet; the hidden tool
        // sheet keeps the formatted version losslessly
        sws.getRow(srow).values = [it.num, it.feature, st.title, st.done ? 'Yes' : 'No',
          RM.htmlToText(st.description || ''), RM.htmlToText(st.ac || '')];
        srow += 1;
      });
    });

    // ---- Team sheet
    var tws = wb.addWorksheet('Team');
    // Rate/Cost append AFTER the hours column so older importers (which read
    // col 5 as hours) still parse this layout
    // column 2 stays the rate-card role (older importers read it as the type);
    // the free-text Role/title appends at the end to keep the layout stable
    tws.getRow(1).values = ['Person', 'Rate card role', 'Workstream', 'Capacity (at full-time)', 'Week hours (overrides)', 'Rate (hourly)', 'Cost (hourly)', 'Role (title)'];
    tws.getRow(1).font = { bold: true };
    tws.getColumn(1).width = 28;
    tws.getColumn(2).width = 18;
    tws.getColumn(3).width = 20;
    tws.getColumn(4).width = 16;
    tws.getColumn(5).width = 60;
    tws.getColumn(6).width = 13;
    tws.getColumn(7).width = 13;
    tws.getColumn(8).width = 22;
    state.team.forEach(function (m, i) {
      var wh = m.weekHours || {};
      var txt = Object.keys(wh).sort().map(function (iso) { return iso + '=' + wh[iso]; }).join(', ');
      tws.getRow(i + 2).values = [m.name || null, m.type || null, m.workstream || null, m.capacity != null ? m.capacity : 1, txt,
        m.rate || 0, m.cost || 0, m.role || null];
    });

    // ---- hidden lossless state sheet
    var hws = wb.addWorksheet('_RoadmapTool');
    hws.getCell('A1').value = STATE_VERSION;
    var json = RMExcel.stateJsonOf(state);
    var n = 0;
    for (var off = 0; off < json.length;) {
      var len = Math.min(CHUNK, json.length - off);
      hws.getCell(2 + n, 2).value = json.slice(off, off + len);
      off += len;
      n += 1;
    }
    hws.getCell('A2').value = n; // chunk count
    if (ui) hws.getCell('A3').value = asciiJson(ui);
    hws.state = 'veryHidden';

    return wb.xlsx.writeBuffer().then(function (buf) {
      if (typeof window !== 'undefined' && typeof Blob !== 'undefined') {
        return new Blob([buf], {
          type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        });
      }
      return buf;
    });
  };

  // ------------------------------------------------------------------ import
  // Raw cell string — NO trimming (JSON chunk boundaries may fall on
  // whitespace inside user text; trimming would corrupt the reassembly).
  function cellRaw(cell) {
    if (!cell) return '';
    var v = cell.value;
    if (v == null) return '';
    if (v instanceof Date) return v.toISOString().slice(0, 10);
    if (typeof v === 'object') {
      if (v.richText) return v.richText.map(function (rt) { return rt.text; }).join('');
      if (v.text != null) return String(v.text);
      if (v.result != null) return String(v.result);
      return '';
    }
    return String(v);
  }

  function cellText(cell) {
    return cellRaw(cell).trim();
  }

  function cellDate(cell) {
    if (!cell) return null;
    var v = cell.value;
    if (v instanceof Date) return v;
    if (typeof v === 'object' && v && v.result instanceof Date) return v.result;
    var t = cellText(cell);
    if (/^\d{4}-\d{2}-\d{2}/.test(t)) return RM.parseISO(t);
    return null;
  }

  // Solid fill classification: returns 'bar' | 'shade' | null.
  function fillKind(cell) {
    var f = cell && cell.fill;
    if (!f || f.type !== 'pattern' || f.pattern !== 'solid' || !f.fgColor) return null;
    var fg = f.fgColor;
    if (fg.argb) {
      var hex = fg.argb.slice(-6).toUpperCase();
      if (hex === '000000' || hex === 'FFFFFF') return 'shade';
      // grays
      var r = parseInt(hex.slice(0, 2), 16), g = parseInt(hex.slice(2, 4), 16), b = parseInt(hex.slice(4, 6), 16);
      if (Math.abs(r - g) < 8 && Math.abs(g - b) < 8) return 'shade';
      return 'bar';
    }
    if (fg.theme != null) {
      // theme 0/1 are the bg/fg neutrals (grays); 2/3 (dk2/lt2) and the
      // accents (4+) are used as bar paint — the source workbook draws its
      // blue bars as dk2 tints, not accents
      return fg.theme >= 2 ? 'bar' : 'shade';
    }
    return null;
  }

  // Identity of a solid fill, for telling a pale lead-in tint apart from the
  // bar's dominant paint. Distinct source colors get distinct keys.
  function fillKey(cell) {
    var f = cell && cell.fill;
    if (!f || f.type !== 'pattern' || f.pattern !== 'solid' || !f.fgColor) return null;
    var fg = f.fgColor;
    if (fg.argb) return 'rgb:' + fg.argb.slice(-6).toUpperCase();
    if (fg.theme != null) return 't' + fg.theme + ':' + (Math.round((fg.tint || 0) * 100) / 100);
    return null;
  }

  // Approximate lightness of a fill key, to tell the pale risk area from the
  // solid work paint regardless of which end of the bar it sits on.
  function keyLightness(key) {
    if (!key) return 0;
    if (key.slice(0, 4) === 'rgb:') {
      var hex = key.slice(4);
      var r = parseInt(hex.slice(0, 2), 16), g = parseInt(hex.slice(2, 4), 16), b = parseInt(hex.slice(4, 6), 16);
      return (0.299 * r + 0.587 * g + 0.114 * b) / 255;
    }
    var m = /^t\d+:(-?[\d.]+)$/.exec(key);
    if (m) return 0.5 + parseFloat(m[1]) / 2; // tint -1..1 → 0..1 (same theme: lighter = higher)
    return 0;
  }

  function parseDeps(text) {
    var out = { deps: [], depsText: [] };
    if (!text || /^none$/i.test(text.trim())) return out;
    text.split(/[,\n]/).forEach(function (tok) {
      tok = tok.trim();
      if (!tok) return;
      if (/^\d+$/.test(tok)) out.deps.push(parseInt(tok, 10));
      else if (/^all above/i.test(tok)) { /* dropped — "All above" is no longer supported */ }
      else out.depsText.push(tok);
    });
    return out;
  }

  RMExcel.importWorkbook = function (arrayBuffer) {
    var ExcelJS = ExcelJSRef();
    var wb = new ExcelJS.Workbook();
    return wb.xlsx.load(arrayBuffer).then(function () {
      var hws = wb.getWorksheet('_RoadmapTool');
      if (hws && cellText(hws.getCell('A1')) === STATE_VERSION) {
        var chunks = [];
        var count = parseInt(cellText(hws.getCell('A2')), 10) || 0;
        for (var i = 0; i < count; i++) chunks.push(cellRaw(hws.getCell(2 + i, 2)));
        try {
          var parsed = JSON.parse(chunks.join(''));
          var ui = null;
          try { ui = JSON.parse(cellRaw(hws.getCell('A3')) || 'null'); } catch (e2) { /* prefs optional */ }
          var st = RM.normalizeState(parsed);
          try { reconcileVisibleEdits(wb, st); } catch (e3) { /* visible sheets optional */ }
          return { state: st, source: 'tool', ui: ui };
        } catch (e) { /* fall through to template parse */ }
      }
      return parseTemplate(wb);
    });
  };

  // Tool files load losslessly from the hidden JSON — but people also edit
  // the visible sheets in Excel. Where a visible text cell disagrees with
  // the hidden state, the visible edit wins (feature names, scope text,
  // story titles/done). Structural edits (rows added/removed, bars moved)
  // stay out of scope: the hidden state remains authoritative for those.
  function reconcileVisibleEdits(wb, state) {
    function norm(v) { return String(v == null ? '' : v).replace(/\r\n?/g, '\n').trim(); }
    var byNum = {};
    state.items.forEach(function (it) { byNum[it.num] = it; });

    var ws = wb.getWorksheet('Roadmap');
    if (ws) {
      // column layout matches the export: 1 ID · 4 Feature · 5 Enables ·
      // 6 Out of scope · 7 Notes · 9 External dependencies
      var FIELDS = [[4, 'feature'], [5, 'enables'], [6, 'outOfScope'], [7, 'notes'], [9, 'extDeps']];
      for (var r = 4; r <= ws.rowCount; r++) {
        var row = ws.getRow(r);
        var idTxt = cellText(row.getCell(1));
        if (!/^\d+$/.test(idTxt)) continue; // band/blank rows
        var it = byNum[parseInt(idTxt, 10)];
        if (!it) continue;
        FIELDS.forEach(function (f) {
          var v = cellText(row.getCell(f[0]));
          // rich fields export flattened — only a real Excel-side edit (vs the
          // flattened text) replaces the stored value, as plain text
          var stored = f[1] === 'feature' ? it[f[1]] : RM.htmlToText(it[f[1]] || '');
          if (norm(v) !== norm(stored)) it[f[1]] = norm(v);
        });
      }
    }

    var sws = wb.getWorksheet('Stories');
    if (sws) {
      var sheetStories = {}; // item num -> [{title, done}] in sheet order
      for (var sr = 2; sr <= sws.rowCount; sr++) {
        var srow = sws.getRow(sr);
        var sid = cellText(srow.getCell(1));
        if (!/^\d+$/.test(sid)) continue;
        (sheetStories[parseInt(sid, 10)] = sheetStories[parseInt(sid, 10)] || []).push({
          title: cellText(srow.getCell(3)),
          done: /^yes$/i.test(cellText(srow.getCell(4)))
        });
      }
      Object.keys(sheetStories).forEach(function (num) {
        var it = byNum[num];
        // only positional matching is safe — skip items whose story count
        // changed in the sheet
        if (!it || it.stories.length !== sheetStories[num].length) return;
        it.stories.forEach(function (st, i) {
          var sv = sheetStories[num][i];
          if (norm(sv.title) && norm(sv.title) !== norm(st.title)) st.title = norm(sv.title);
          st.done = sv.done;
        });
      });
    }
  }

  function parseTemplate(wb) {
    var ws = wb.getWorksheet('Roadmap');
    if (!ws) {
      wb.eachSheet(function (sheet) {
        if (!ws && /roadmap/i.test(sheet.name)) ws = sheet;
      });
    }
    if (!ws) ws = wb.worksheets[0];
    if (!ws) throw new Error('No worksheet found');

    // locate header row: contains both "ID" and "Feature"
    var headerRow = null;
    for (var r = 1; r <= Math.min(ws.rowCount, 12); r++) {
      var vals = [];
      ws.getRow(r).eachCell({ includeEmpty: false }, function (c) { vals.push(cellText(c).toLowerCase()); });
      if (vals.indexOf('id') !== -1 && vals.indexOf('feature') !== -1) { headerRow = r; break; }
    }
    if (!headerRow) throw new Error('Could not find the header row (needs ID + Feature columns)');

    var colMap = {};
    var sprintCols = [];
    var markerCols = {};
    var extraMap = {};
    var hr = ws.getRow(headerRow);
    hr.eachCell({ includeEmpty: false }, function (c, colNumber) {
      var t = cellText(c).toLowerCase();
      var d = cellDate(c);
      if (d) { sprintCols.push({ col: colNumber, date: d }); return; }
      if (t === 'id') colMap.num = colNumber;
      else if (t.indexOf('workstream') === 0) colMap.workstream = colNumber;
      else if (t === 'epic') colMap.epic = colNumber;
      else if (t === 'feature') colMap.feature = colNumber;
      else if (t === 'enables') colMap.enables = colNumber;
      else if (t.indexOf('out of scope') === 0) colMap.outOfScope = colNumber;
      else if (t === 'notes') colMap.notes = colNumber;
      else if (t === 'dependencies') colMap.deps = colNumber;
      else if (t.indexOf('external dep') === 0) colMap.extDeps = colNumber;
      else if (t.indexOf('t-shirt') === 0 || t.indexOf('size') === 0) colMap.size = colNumber;
      else if (t.indexOf('risk') !== -1) colMap.risk = colNumber;
      else if (t === 'next') markerCols.next = colNumber;
      else if (t === 'future') markerCols.future = colNumber;
      else if (t === 'headcount') extraMap.headcount = colNumber;
      else if (t === 'team type') extraMap.teamType = colNumber;
      else if (t === 'status') extraMap.status = colNumber;
    });
    if (!sprintCols.length) throw new Error('No sprint date columns found in the header row');
    sprintCols.sort(function (a, b) { return a.col - b.col; });

    // weeks per COLUMN, inferred from the header-date gaps: 7-day gaps mean
    // weekly columns (this tool's exports), 14-day gaps a legacy sprint
    // layout. This is grid granularity only — it does NOT set the sprint
    // length, which stays the app's own concept (default 2 weeks).
    var colWeeks = 2;
    if (sprintCols.length > 1) {
      var diffDaysGap = Math.round((sprintCols[1].date - sprintCols[0].date) / 86400000);
      colWeeks = Math.max(1, Math.round(diffDaysGap / 7));
    }
    var daysPerSprint = colWeeks * 5; // template layouts are Mon-Fri weeks
    var meta = {
      title: 'Imported Roadmap',
      timelineStart: RM.fmtISO(sprintCols[0].date),
      numWeeks: sprintCols.length * colWeeks,
      blackoutWeeks: [],
      sizeDays: RM.clone(RM.DEFAULT_SIZE_DAYS)
    };
    var colByIndex = {};
    sprintCols.forEach(function (sc, i) { colByIndex[sc.col] = i; });

    var phases = [];
    var items = [];
    var currentPhase = null;
    var shadeCounts = [];
    for (var sc0 = 0; sc0 < sprintCols.length; sc0++) shadeCounts.push(0);
    var itemRowCount = 0;

    function ensurePhase(name, description) {
      var bucket = /^(next|future|backlog|later)\b/i.test(name.trim());
      var existing = null;
      phases.forEach(function (p) { if (p.name.toLowerCase() === name.trim().toLowerCase()) existing = p; });
      if (existing) { currentPhase = existing; return existing; }
      var p = {
        id: 'ph' + (phases.length + 1),
        name: titleCase(name.trim()),
        description: description || '',
        bucket: bucket,
        collapsed: false
      };
      phases.push(p);
      currentPhase = p;
      return p;
    }

    var ACRONYMS = { mvp: 'MVP', os: 'OS', qa: 'QA', uat: 'UAT', bp1: 'BP1' };
    function titleCase(sname) {
      if (sname !== sname.toUpperCase()) return sname;
      return sname.toLowerCase()
        .replace(/(^|[\s:\-(])([a-z])/g, function (mm, a, b) { return a + b.toUpperCase(); })
        .replace(/[A-Za-z0-9]+/g, function (word) {
          return ACRONYMS[word.toLowerCase()] || word;
        });
    }

    for (var rr = headerRow + 1; rr <= ws.rowCount; rr++) {
      var row = ws.getRow(rr);
      var aText = cellText(row.getCell(colMap.num || 1));
      var epicText = colMap.epic ? cellText(row.getCell(colMap.epic)) : '';
      var featText = colMap.feature ? cellText(row.getCell(colMap.feature)) : '';

      var isBand = aText && !epicText && !/^\d+$/.test(aText);
      if (isBand) {
        ensurePhase(aText, featText);
        continue;
      }
      if (!featText) continue;
      if (!currentPhase) ensurePhase('Imported', '');

      var depInfo = parseDeps(colMap.deps ? cellText(row.getCell(colMap.deps)) : '');
      var barSprints = [];
      var barKeys = [];
      itemRowCount += 1;
      sprintCols.forEach(function (sc, i) {
        var kind = fillKind(row.getCell(sc.col));
        if (kind === 'bar') { barSprints.push(i); barKeys.push(fillKey(row.getCell(sc.col))); }
        else if (kind === 'shade') shadeCounts[i] += 1;
      });
      var startDay = null, durDays = null, riskDays = 0;
      if (barSprints.length) {
        var s0 = barSprints[0], s1 = barSprints[barSprints.length - 1];
        startDay = s0 * daysPerSprint;
        var totalDays = (s1 - s0 + 1) * daysPerSprint;
        // The pale run at either edge of the bar is the risk area. Identify
        // it by lightness (works for leading lead-ins in the source workbook
        // and the trailing buffers this tool writes); the buffer is stored
        // trailing either way — start and total span are preserved.
        var distinct = [];
        barKeys.forEach(function (k) { if (distinct.indexOf(k) === -1) distinct.push(k); });
        var riskCells = 0;
        if (distinct.length === 2) {
          var lightKey = keyLightness(distinct[0]) >= keyLightness(distinct[1]) ? distinct[0] : distinct[1];
          var bi;
          if (barKeys[barKeys.length - 1] === lightKey) {
            for (bi = barKeys.length - 1; bi > 0 && barKeys[bi] === lightKey; bi--) riskCells += 1;
          } else if (barKeys[0] === lightKey) {
            for (bi = 0; bi < barKeys.length - 1 && barKeys[bi] === lightKey; bi++) riskCells += 1;
          }
        }
        riskDays = riskCells * daysPerSprint;
        durDays = Math.max(daysPerSprint, totalDays - riskDays);
      }
      var sizeText = colMap.size ? cellText(row.getCell(colMap.size)).toUpperCase() : '';
      var statusText = extraMap.status ? cellText(row.getCell(extraMap.status)).toLowerCase() : '';
      items.push({
        num: /^\d+$/.test(aText) ? parseInt(aText, 10) : null,
        phaseId: currentPhase.id,
        workstream: colMap.workstream ? cellText(row.getCell(colMap.workstream)) : '',
        epic: epicText,
        feature: featText,
        enables: colMap.enables ? cellText(row.getCell(colMap.enables)) : '',
        outOfScope: colMap.outOfScope ? cellText(row.getCell(colMap.outOfScope)) : '',
        notes: colMap.notes ? cellText(row.getCell(colMap.notes)) : '',
        deps: depInfo.deps,
        depsText: depInfo.depsText,
        extDeps: colMap.extDeps ? cellText(row.getCell(colMap.extDeps)) : '',
        size: RM.SIZE_ORDER.indexOf(sizeText) !== -1 ? sizeText : null,
        risk: colMap.risk ? (cellText(row.getCell(colMap.risk)) || null) : null,
        headcount: extraMap.headcount ? (parseFloat(cellText(row.getCell(extraMap.headcount))) || 1) : 1,
        teamType: extraMap.teamType ? cellText(row.getCell(extraMap.teamType)) : '',
        startDay: startDay,
        durDays: durDays,
        riskDays: riskDays,
        locked: statusText === 'locked',
        done: statusText === 'done',
        stories: []
      });
    }

    // Stories sheet
    var sws = wb.getWorksheet('Stories');
    if (sws) {
      var byNum = {};
      items.forEach(function (it) { if (it.num != null) byNum[it.num] = it; });
      for (var sr = 2; sr <= sws.rowCount; sr++) {
        var numTxt = cellText(sws.getCell(sr, 1));
        var title = cellText(sws.getCell(sr, 3));
        if (!numTxt || !title) continue;
        var target = byNum[parseInt(numTxt, 10)];
        if (target) {
          target.stories.push({
            title: title, done: /^y(es)?$/i.test(cellText(sws.getCell(sr, 4))),
            description: cellText(sws.getCell(sr, 5)), ac: cellText(sws.getCell(sr, 6))
          });
        }
      }
    }

    // Team sheet
    var team = [];
    var tws = wb.getWorksheet('Team');
    if (tws) {
      for (var tr = 2; tr <= tws.rowCount; tr++) {
        var nm = cellText(tws.getCell(tr, 1));
        // names are optional — a row with just a role/title still counts
        if (nm || cellText(tws.getCell(tr, 8)) || cellText(tws.getCell(tr, 2))) {
          // new exports: 3=workstream, 4=capacity, 5=hours; legacy layouts had
          // hours in col 3 or 4 — hours cells are recognizable by their dates
          var c3 = cellText(tws.getCell(tr, 3));
          var c4 = cellText(tws.getCell(tr, 4));
          var c5 = cellText(tws.getCell(tr, 5));
          var isHours = function (t) { return /\d{4}-\d{2}-\d{2}/.test(t); };
          var hoursRaw = c5 || (isHours(c4) ? c4 : '') || (isHours(c3) ? c3 : '');
          var wsVal = isHours(c3) ? '' : c3;
          var capVal = !isHours(c4) && isFinite(parseFloat(c4)) && parseFloat(c4) >= 0 ? parseFloat(c4) : 1;
          var weekHours = {}, offWeeks = [];
          if (hoursRaw) {
            hoursRaw.split(/[,\s]+/).forEach(function (tok) {
              var mHours = /^(\d{4}-\d{2}-\d{2})=([\d.]+)$/.exec(tok);
              if (mHours) weekHours[mHours[1]] = parseFloat(mHours[2]);
              else if (/^\d{4}-\d{2}-\d{2}$/.test(tok)) offWeeks.push(tok); // legacy off-week list
            });
          }
          team.push({
            name: nm,
            role: cellText(tws.getCell(tr, 8)) || '',
            type: cellText(tws.getCell(tr, 2)) || 'Development',
            workstream: wsVal,
            capacity: capVal,
            weekHours: weekHours,
            offWeeks: offWeeks,
            rate: parseFloat(cellText(tws.getCell(tr, 6))) || 0,
            cost: parseFloat(cellText(tws.getCell(tr, 7))) || 0
          });
        }
      }
    }

    // Sprint columns gray-shaded on most item rows are holiday/blackout sprints.
    if (itemRowCount >= 4) {
      sprintCols.forEach(function (sc, i) {
        if (shadeCounts[i] >= Math.max(3, itemRowCount * 0.5)) {
          for (var wI = 0; wI < colWeeks; wI++) {
            var wd = RM.parseISO(meta.timelineStart);
            wd.setUTCDate(wd.getUTCDate() + (i * colWeeks + wI) * 7);
            var iso = RM.fmtISO(wd);
            if (meta.blackoutWeeks.indexOf(iso) === -1) meta.blackoutWeeks.push(iso);
          }
        }
      });
      meta.blackoutWeeks.sort();
    }

    var state = RM.normalizeState({
      meta: meta, phases: phases, items: items, team: team,
      teamTypes: null
    });
    return { state: state, source: 'template' };
  }

  if (typeof module !== 'undefined' && module.exports) module.exports = RMExcel;
  root.RMExcel = RMExcel;
})(typeof window !== 'undefined' ? window : globalThis);
