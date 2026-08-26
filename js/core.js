/*
 * Headway core — pure logic, no DOM. Loaded in the browser as window.RM and
 * in node (tests) via require. Time is measured in WORKING DAYS from
 * meta.timelineStart (a Monday): 5 per week, weekends don't exist in the index
 * space. week = floor(day / 5).
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
  RM.sizeOrderOf = function (state) {
    var m = state.meta || state;
    return m.sizeOrder || RM.SIZE_ORDER;
  };
  RM.sizingEnabled = function (state) {
    var m = state.meta || state;
    return m.sizeScheme !== 'none' && RM.sizeOrderOf(state).length > 0;
  };
  RM.setSizeScheme = function (state, scheme) {
    var def = RM.SIZE_SCHEMES[scheme];
    if (!def || scheme === 'custom') return;
    var m = state.meta;
    m.sizeScheme = scheme;
    m.sizeOrder = def.sizes.slice();
    m.sizeDays = RM.clone(def.days);
  };
  RM.renameSizeOption = function (state, oldLabel, newLabel) {
    var m = state.meta;
    if (!newLabel || oldLabel === newLabel || m.sizeOrder.indexOf(newLabel) !== -1) return;
    m.sizeOrder = m.sizeOrder.map(function (l) { return l === oldLabel ? newLabel : l; });
    m.sizeDays[newLabel] = m.sizeDays[oldLabel];
    delete m.sizeDays[oldLabel];
    state.items.forEach(function (it) { if (it.size === oldLabel) it.size = newLabel; });
    m.sizeScheme = 'custom';
  };
  RM.addSizeOption = function (state, label, days) {
    var m = state.meta;
    if (!label || m.sizeOrder.indexOf(label) !== -1) return;
    m.sizeOrder.push(label);
    m.sizeDays[label] = isFinite(+days) && +days > 0 ? +days : 5;
    m.sizeScheme = 'custom';
  };
  RM.removeSizeOption = function (state, label) {
    var m = state.meta;
    m.sizeOrder = m.sizeOrder.filter(function (l) { return l !== label; });
    delete m.sizeDays[label];
    state.items.forEach(function (it) { if (it.size === label) it.size = null; });
    m.sizeScheme = 'custom';
  };
  RM.RISK_ORDER = ['L', 'M', 'H']; // low / medium / high (severity, not a size)
  RM.DEFAULT_TEAM_TYPES = ['Development', 'Design', 'Product', 'Data', 'QA'];
  RM.DEFAULT_WORK_TYPE = 'Development';
  RM.WEEK_HOURS = 40; // one person's full week
  RM.ANY_TYPE = '';

  // Scoping-view columns. Built-ins map to fixed item fields; custom columns
  // ('c…' keys) store their text in item.custom.
  RM.SCOPE_BUILTIN_LABELS = {
    description: 'Description',
    enables: 'Enables',
    outOfScope: 'Out of scope',
    extDeps: 'External dependencies',
    notes: 'Notes'
  };
  // New documents start with Description only; the rest stay available in
  // the add-column menu. Legacy docs infer their list from actual content.
  RM.DEFAULT_SCOPE_COLS = ['description'];
  RM.SCOPE_BUILTIN_ORDER = ['description', 'enables', 'outOfScope', 'extDeps', 'notes'];

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
  // process green / mixed plum.
  RM.PALETTE = {
    product: '3273BD',
    data: 'C25E0E',
    process: '08875B',
    mixed: 'A14FBF'
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
  RM.colorForWs = function (state, ws) {
    var c = state && state.wsColors && ws ? resolveColor(state.wsColors[ws]) : null;
    return c || RM.PALETTE.product;
  };
  RM.colorForItem = function (state, it) {
    return RM.colorForWs(state, it.workstream);
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

  // ---------------------------------------------------------------- calendar
  RM.parseISO = function (iso) {
    var p = String(iso).slice(0, 10).split('-').map(Number);
    return new Date(Date.UTC(p[0], p[1] - 1, p[2]));
  };
  RM.fmtISO = function (dt) {
    if (!dt || !isFinite(dt.getTime())) return '';
    return dt.toISOString().slice(0, 10);
  };

  RM.numDays = function (meta) { return meta.numWeeks * 5; };

  // keep meta.endDate (last working day, a Friday) in sync with numWeeks;
  // call after any change to numWeeks or timelineStart
  RM.syncEndDate = function (meta) {
    meta.endDate = RM.fmtISO(RM.dayToDate(meta, meta.numWeeks * 5 - 1));
  };

  RM.weekStartDate = function (meta, week) {
    var d = RM.parseISO(meta.timelineStart);
    d.setUTCDate(d.getUTCDate() + week * 7);
    return d;
  };

  RM.dayToDate = function (meta, day) {
    var week = Math.floor(day / 5);
    var dow = ((day % 5) + 5) % 5;
    var d = RM.parseISO(meta.timelineStart);
    d.setUTCDate(d.getUTCDate() + week * 7 + dow);
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
    var dow = diff - week * 7;
    if (dow > 4) dow = 4; // weekend -> Friday of that week
    return week * 5 + dow;
  };

  var MONTHS = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];
  RM.fmtShort = function (dt) { return MONTHS[dt.getUTCMonth()] + ' ' + dt.getUTCDate(); };
  RM.fmtShortYear = function (dt) { return MONTHS[dt.getUTCMonth()] + ' ' + dt.getUTCDate() + ' ’' + String(dt.getUTCFullYear()).slice(2); };

  // Holidays are INDIVIDUAL dates (meta.holidays, ISO strings). This builds a
  // { workingDayIndex: true } lookup; weekend-dated holidays are ignored since
  // weekends don't exist in the index space.
  RM.holidayDaySet = function (meta) {
    var set = {};
    (meta.holidays || []).forEach(function (iso) {
      var d = RM.parseISO(iso);
      if (!d || !isFinite(d.getTime())) return;
      var dow = (d.getUTCDay() + 6) % 7;
      if (dow > 4) return;
      var day = RM.dateToDay(meta, d);
      if (day != null) set[day] = true;
    });
    return set;
  };

  RM.isHolidayDay = function (meta, day, set) {
    return (set || RM.holidayDaySet(meta))[day] === true;
  };

  // A week is "blacked out" only when ALL FIVE of its working days are
  // holidays — those weeks are excluded from capacity math entirely.
  RM.isBlackoutWeek = function (meta, week, set) {
    set = set || RM.holidayDaySet(meta);
    for (var i = 0; i < 5; i++) if (!set[week * 5 + i]) return false;
    return true;
  };

  RM.holidaysInWeek = function (meta, week, set) {
    set = set || RM.holidayDaySet(meta);
    var n = 0;
    for (var i = 0; i < 5; i++) if (set[week * 5 + i]) n += 1;
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
      if (!set[d]) remaining -= 1;
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
      if (!set[d]) n += 1;
    }
    return n;
  };

  // Sprint numbering anchor: sprint boundaries fall every weeksPerSprint weeks
  // aligned to meta.sprintAnchor (default: timelineStart), and the sprint that
  // starts there is numbered meta.sprintAnchorNum (default 1). Weeks before
  // the anchor number down through S0, S-1, …
  RM.sprintInfo = function (meta) {
    var wps = meta.weeksPerSprint || 2;
    var anchorWeek = 0;
    if (meta.sprintAnchor) {
      var d = RM.dateToDay(meta, RM.parseISO(meta.sprintAnchor));
      if (d != null) anchorWeek = Math.round(d / 5);
    }
    return { wps: wps, anchorWeek: anchorWeek, firstNum: meta.sprintAnchorNum != null ? meta.sprintAnchorNum : 1 };
  };

  RM.sprintNumForWeek = function (meta, week) {
    var si = RM.sprintInfo(meta);
    return si.firstNum + Math.floor((week - si.anchorWeek) / si.wps);
  };

  // plain-text projection of stored rich text (tooltips, Excel cells, search)
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

  RM.sizeDays = function (state, size) {
    var map = (state.meta && state.meta.sizeDays) || RM.DEFAULT_SIZE_DAYS;
    return size && map[size] != null ? map[size] : null;
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

  // ---------------------------------------------------------------- state
  RM.normalizeState = function (raw) {
    var state = RM.clone(raw || {});
    state.meta = state.meta || {};
    var m = state.meta;
    m.title = m.title || 'Roadmap';
    m.timelineStart = m.timelineStart || '2026-07-27';
    m.numWeeks = m.numWeeks || (m.numSprints ? m.numSprints * (m.weeksPerSprint || 2) : 48);
    m.weeksPerSprint = m.weeksPerSprint || 2;
    // capacity feature switch — roster-based scheduling constraints and the
    // capacity header row. OFF by default; enabled per-document in Setup.
    m.capacityEnabled = !!m.capacityEnabled;
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
        var week = Math.floor(day / 5);
        for (var i = 0; i < 5; i++) {
          var dIso = RM.fmtISO(RM.dayToDate(m, week * 5 + i));
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
        return k === 'description' || (state.items || []).some(function (it) { return it && it[k]; });
      }).map(function (k) { return { key: k }; });
    }
    m.scopeCols = inheritedCols
      .map(function (c) {
        if (typeof c === 'string') c = { key: c };
        if (!c || typeof c.key !== 'string' || !c.key) return null;
        if (RM.SCOPE_BUILTIN_LABELS[c.key]) {
          return c.label ? { key: c.key, label: String(c.label) } : { key: c.key };
        }
        return { key: c.key, label: String(c.label || 'Column') };
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
    // workstream feature switch — ON unless the project turned it off
    m.workstreamsEnabled = m.workstreamsEnabled !== false;
    delete m.sprintDates;

    state.phases = (state.phases || []).map(function (p) {
      // startDay/endDay: optional user-pinned phase window (working-day
      // indices); null = auto-derived from the phase's items
      var ps = p.startDay != null && isFinite(p.startDay) ? Math.max(0, Math.round(p.startDay)) : null;
      var pe = p.endDay != null && isFinite(p.endDay) ? Math.max(0, Math.round(p.endDay)) : null;
      if (ps != null && pe != null && pe <= ps) pe = ps + 1;
      return {
        id: p.id || RM.uid('p'),
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

    var phaseIds = {};
    state.phases.forEach(function (p) { phaseIds[p.id] = true; });
    var fallbackPhase = state.phases[0].id;

    state.items = (state.items || []).map(function (it) {
      return {
        id: it.id || RM.uid('i'),
        num: it.num != null ? it.num : null,
        phaseId: phaseIds[it.phaseId] ? it.phaseId : fallbackPhase,
        feature: it.feature || '',
        description: it.description || '',
        workstream: it.workstream || '',
        epic: it.epic || '',
        enables: it.enables || '',
        outOfScope: it.outOfScope || '',
        notes: it.notes || '',
        deps: (it.deps || []).map(Number).filter(function (n) { return !isNaN(n); }),
        depsText: it.depsText || [],
        extDeps: it.extDeps || '',
        size: it.size && RM.SIZE_ORDER.indexOf(it.size) !== -1 ? it.size : (it.size || null),
        // risk is a severity (None/L/M/H); legacy t-shirt values migrate:
        // XS/S → L, M → M, XL → H ('L' is low in both readings and stays L)
        risk: (function () {
          var rv = it.risk ? String(it.risk).toUpperCase() : null;
          if (!rv) return null;
          if (RM.RISK_ORDER.indexOf(rv) !== -1) return rv;
          if (rv === 'XS' || rv === 'S') return 'L';
          if (rv === 'XL') return 'H';
          return null;
        })(),
        headcount: it.headcount != null && it.headcount > 0 ? it.headcount : 1,
        // every work item defaults to 1 × Development
        teamType: it.teamType != null && it.teamType !== '' ? it.teamType : RM.DEFAULT_WORK_TYPE,
        // milestones are fixed dates: zero-duration diamonds on the timeline
        milestone: !!it.milestone,
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
        colorOverride: it.colorOverride || null,
        done: !!it.done,
        stories: (it.stories || []).map(function (s) {
          // stories may carry their own little timeline (startDay/durDays);
          // both null = no timeline (the default)
          var sched = s.startDay != null && isFinite(s.startDay) && s.durDays > 0;
          return {
            id: s.id || RM.uid('s'), title: s.title || '', done: !!s.done,
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
            durDays: sched ? Math.round(s.durDays) : null
          };
        })
      };
    });

    state.items.forEach(function (it) { delete it.leadDays; });

    // Unique nums: assign missing, renumber collisions (first occurrence wins —
    // that matches how deps on the duplicated number already resolved).
    var seen = {};
    var maxNum = 0;
    state.items.forEach(function (it) {
      if (it.num != null && !seen[it.num]) { seen[it.num] = true; maxNum = Math.max(maxNum, it.num); }
    });
    state.items.forEach(function (it) {
      if (it.num == null || seen[it.num] !== true) {
        maxNum += 1; it.num = maxNum; seen[maxNum] = true;
      } else {
        seen[it.num] = 'used';
      }
    });

    state.epicColors = state.epicColors || {}; // legacy — display now keys off workstream
    state.wsColors = state.wsColors && typeof state.wsColors === 'object' ? state.wsColors : {};
    state.epicIcons = state.epicIcons && typeof state.epicIcons === 'object' ? state.epicIcons : {};
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
      .concat((state.team || []).map(function (x) { return x.workstream || ''; }))
      .concat(Object.keys(state.wsColors))
      .forEach(function (w) { if (w && !wsRef[w]) { wsRef[w] = true; wsRefList.push(w); } });
    var wsOrder = [];
    (Array.isArray(state.wsOrder) ? state.wsOrder : []).forEach(function (w) {
      if (wsRef[w] && wsOrder.indexOf(w) === -1) wsOrder.push(w);
    });
    wsRefList.forEach(function (w) { if (wsOrder.indexOf(w) === -1) wsOrder.push(w); });
    state.wsOrder = wsOrder;
    state.teamTypes = state.teamTypes && state.teamTypes.length ? state.teamTypes : RM.clone(RM.DEFAULT_TEAM_TYPES);
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
      return {
        id: mbr.id || RM.uid('t'),
        name: mbr.name || 'Member',
        type: mbr.type || state.teamTypes[0],
        workstream: mbr.workstream || '', // optional workstream assignment
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
    // every referenced work type must exist in the list
    state.items.forEach(function (it) {
      if (it.teamType && state.teamTypes.indexOf(it.teamType) === -1) state.teamTypes.push(it.teamType);
    });
    state.team.forEach(function (mbr) {
      if (mbr.type && state.teamTypes.indexOf(mbr.type) === -1) state.teamTypes.push(mbr.type);
    });
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
  RM.addScopeCol = function (state, label, key) {
    var cols = state.meta.scopeCols;
    if (key && RM.SCOPE_BUILTIN_LABELS[key]) {
      if (!cols.some(function (c) { return c.key === key; })) cols.push({ key: key });
      return key;
    }
    var k = RM.uid('c');
    cols.push({ key: k, label: label || 'Column' });
    return k;
  };
  RM.removeScopeCol = function (state, key) {
    state.meta.scopeCols = state.meta.scopeCols.filter(function (c) { return c.key !== key; });
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
    return 5;
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
  // Concrete dependency item list from the explicit numbered deps.
  // ("All above" support was removed — only specifically-defined deps count.)
  RM.resolveDeps = function (state, it) {
    var out = { deps: [], unknown: [] };
    it.deps.forEach(function (num) {
      var dep = RM.itemByNum(state, num);
      if (!dep) out.unknown.push(num);
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
    if (!member.weekHours) return RM.WEEK_HOURS;
    var iso = RM.fmtISO(RM.weekStartDate(meta, week));
    var h = member.weekHours[iso];
    return h != null && isFinite(h) ? h : RM.WEEK_HOURS;
  };

  // Back-compat: "off" = a zero-hour week.
  RM.memberOffWeek = function (meta, member, week) {
    return RM.memberHoursForWeek(meta, member, week) === 0;
  };

  // Roster availability for one week, in PEOPLE-EQUIVALENTS (hours / 40):
  // the approximate number of parallel work items the roster can absorb.
  RM.availForWeek = function (state, week) {
    var total = 0, byType = {};
    state.team.forEach(function (m) {
      var pe = (RM.memberHoursForWeek(state.meta, m, week) / RM.WEEK_HOURS) * (m.capacity != null ? m.capacity : 1);
      if (pe <= 0) return;
      total += pe;
      byType[m.type] = (byType[m.type] || 0) + pe;
    });
    return { total: total, byType: byType };
  };

  // Weekly demand vs roster capacity. Blackout weeks carry no demand and no
  // check; member off-weeks lower that week's capacity.
  RM.capacity = function (state) {
    var meta = state.meta;
    var weeks = [];
    var teamTotal = state.team.length;
    var typeCounts = {};
    state.team.forEach(function (m) { typeCounts[m.type] = (typeCounts[m.type] || 0) + 1; });
    var w;
    for (w = 0; w < meta.numWeeks; w++) {
      var avail = RM.availForWeek(state, w);
      weeks.push({
        demand: 0, demandByType: {}, items: [],
        cap: teamTotal > 0 ? avail.total : Infinity,
        capByType: teamTotal > 0 ? avail.byType : typeCounts,
        blackout: RM.isBlackoutWeek(meta, w),
        over: false, overTypes: []
      });
    }
    state.items.forEach(function (it) {
      if (it.startDay == null || it.durDays == null || it.done || it.milestone) return;
      var w0 = Math.floor(it.startDay / 5);
      var w1 = Math.floor((it.startDay + it.durDays - 1) / 5);
      for (var wk = Math.max(0, w0); wk <= Math.min(meta.numWeeks - 1, w1); wk++) {
        var cell = weeks[wk];
        if (cell.blackout) continue;
        cell.demand += it.headcount;
        cell.items.push(it.id);
        if (it.teamType) {
          cell.demandByType[it.teamType] = (cell.demandByType[it.teamType] || 0) + it.headcount;
        }
      }
    });
    weeks.forEach(function (cell) {
      if (cell.demand > cell.cap + 1e-9) cell.over = true;
      if (teamTotal > 0) {
        Object.keys(cell.demandByType).forEach(function (t) {
          if (cell.demandByType[t] > (cell.capByType[t] || 0) + 1e-9) {
            cell.over = true;
            cell.overTypes.push(t);
          }
        });
      }
    });
    return { weeks: weeks, teamTotal: teamTotal, typeCounts: typeCounts };
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
      res.unknown.forEach(function (n) {
        add(it, 'warn', 'UNKNOWN_DEP', 'Depends on #' + n + ', which does not exist');
      });
      if (cyclic[it.id]) add(it, 'error', 'CYCLE', 'Part of a dependency cycle');
      if (it.deps.indexOf(it.num) !== -1) add(it, 'warn', 'SELF_DEP', 'Depends on itself (ignored)');
      if (!it.feature.trim()) add(it, 'warn', 'NO_TITLE', 'Feature has no title');

      var scheduled = it.startDay != null && it.durDays != null;
      var phase = phaseById[it.phaseId];
      if (scheduled) {
        if (!it.size && !it.milestone && RM.sizingEnabled(state)) add(it, 'warn', 'NO_SIZE', 'Scheduled without a size');
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

      if (state.meta.capacityEnabled && state.team.length > 0) {
        if (it.teamType) {
          var avail = state.team.filter(function (m) { return m.type === it.teamType; }).length;
          if (it.headcount > avail) {
            add(it, 'warn', 'HC_TYPE', 'Needs ' + it.headcount + ' × ' + it.teamType + ' but roster has ' + avail);
          }
        } else if (it.headcount > state.team.length) {
          add(it, 'warn', 'HC_TOTAL', 'Headcount ' + it.headcount + ' exceeds roster of ' + state.team.length);
        }
      }
    });

    var cap = RM.capacity(state);
    cap.weeks.forEach(function (cell, w) {
      if (!state.meta.capacityEnabled || !cell.over) return;
      var d = RM.weekStartDate(state.meta, w);
      var what = cell.overTypes.length
        ? cell.overTypes.map(function (t) { return t + ' ' + cell.demandByType[t] + '/' + (cap.typeCounts[t] || 0); }).join(', ')
        : cell.demand + ' needed / ' + cell.cap + ' on roster';
      global.push({
        level: 'warn', code: 'OVER_CAP', week: w,
        msg: 'Week of ' + RM.fmtShort(d) + ' over capacity (' + what + ')',
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
    var typeCounts = {};
    state.team.forEach(function (m) { typeCounts[m.type] = (typeCounts[m.type] || 0) + 1; });
    var HORIZON_WEEKS = meta.numWeeks + 104; // allow spill; UI can extend the grid
    var ledgerTotal = new Array(HORIZON_WEEKS);
    var ledgerType = {};
    for (var w = 0; w < HORIZON_WEEKS; w++) ledgerTotal[w] = 0;

    function occupy(it, startDay, durDays) {
      var w0 = Math.floor(startDay / 5);
      var w1 = Math.floor((startDay + durDays - 1) / 5);
      for (var wk = w0; wk <= w1 && wk < HORIZON_WEEKS; wk++) {
        if (RM.isBlackoutWeek(meta, wk)) continue;
        ledgerTotal[wk] += it.headcount;
        if (it.teamType) {
          if (!ledgerType[it.teamType]) {
            ledgerType[it.teamType] = new Array(HORIZON_WEEKS);
            for (var z = 0; z < HORIZON_WEEKS; z++) ledgerType[it.teamType][z] = 0;
          }
          ledgerType[it.teamType][wk] += it.headcount;
        }
      }
    }
    fixed.forEach(function (it) { occupy(it, it.startDay, it.durDays); });

    function fits(it, startDay, durDays) {
      if (teamTotal === 0) return true; // no roster -> no capacity constraint
      var w0 = Math.floor(startDay / 5);
      var w1 = Math.floor((startDay + durDays - 1) / 5);
      for (var wk = w0; wk <= w1; wk++) {
        if (wk >= HORIZON_WEEKS) return true;
        if (RM.isBlackoutWeek(meta, wk)) continue;
        var avail = RM.availForWeek(state, wk);
        if (ledgerTotal[wk] + it.headcount > avail.total + 1e-9) return false;
        if (it.teamType) {
          var typeCap = avail.byType[it.teamType] || 0;
          var used = ledgerType[it.teamType] ? ledgerType[it.teamType][wk] : 0;
          if (used + it.headcount > typeCap + 1e-9) return false;
        }
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

    // An item that can NEVER fit the roster — headcount above the PEAK weekly
    // availability (hours-based people-equivalents, total or for its type) —
    // must not trigger an endless capacity walk. Peaks are hours-aware, so a
    // roster of part-time roles counts fractionally.
    var peakTotal = 0, peakType = {};
    if (teamTotal > 0) {
      for (var pw = 0; pw < HORIZON_WEEKS; pw++) {
        var pa = RM.availForWeek(state, pw);
        if (pa.total > peakTotal) peakTotal = pa.total;
        Object.keys(pa.byType).forEach(function (tk) {
          if (pa.byType[tk] > (peakType[tk] || 0)) peakType[tk] = pa.byType[tk];
        });
        if (pw > meta.numWeeks && pa.total === peakTotal) break; // hours settle after overrides end
      }
    }
    function infeasible(it) {
      if (teamTotal === 0) return false;
      if (it.headcount > peakTotal + 1e-9) return true;
      if (it.teamType && it.headcount > (peakType[it.teamType] || 0) + 1e-9) return true;
      return false;
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
      while (!fits(it, s, dur) && guard < HORIZON_WEEKS * 5) {
        s += 1;
        dur = RM.stretchSpan(meta, s, work);
        guard += 1;
      }
      if (guard >= HORIZON_WEEKS * 5) {
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

    var neededWeeks = Math.ceil(maxDay / 5);
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
    var typeCounts = {};
    state.team.forEach(function (m) { typeCounts[m.type] = (typeCounts[m.type] || 0) + 1; });

    if (teamTotal > 0 && (it.headcount > teamTotal ||
      (it.teamType && it.headcount > (typeCounts[it.teamType] || 0)))) {
      var what = it.teamType
        ? it.headcount + ' × ' + it.teamType + ' (roster has ' + (typeCounts[it.teamType] || 0) + ')'
        : it.headcount + ' people (roster has ' + teamTotal + ')';
      return { state: state, changed: 0, note: 'Needs ' + what + ' — no slot can ever fit. Left unchanged.' };
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
    var LIMIT = (meta.numWeeks + 104) * 5;
    function fits(s, dur) {
      if (teamTotal === 0) return true;
      var w0 = Math.floor(s / 5), w1 = Math.floor((s + dur - 1) / 5);
      for (var wk = w0; wk <= w1; wk++) {
        if (RM.isBlackoutWeek(meta, wk)) continue;
        var cell = wk < capData.weeks.length ? capData.weeks[wk] : null;
        var avail = cell ? { total: cell.cap, byType: cell.capByType } : RM.availForWeek(state, wk);
        var demand = cell ? cell.demand : 0;
        if (demand + it.headcount > avail.total + 1e-9) return false;
        if (it.teamType) {
          var used = cell ? (cell.demandByType[it.teamType] || 0) : 0;
          if (used + it.headcount > (avail.byType[it.teamType] || 0) + 1e-9) return false;
        }
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
    var neededWeeks = Math.ceil((s + dur + riskSpan) / 5);
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
  // next available one. Dependency references follow the rename.
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
    state.items.forEach(function (x) {
      x.deps = x.deps.map(function (d) { return d === old ? n : d; });
    });
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

  RM.shiftDependents = function (state, itemId, delta) {
    if (!delta) return 0;
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
        if (ch.startDay == null || ch.locked || ch.done || ch.milestone) return;
        // never start before any scheduled dependency's buffered end
        var floor = 0;
        ch.deps.forEach(function (n) {
          var dp = RM.itemByNum(state, n);
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

  // ------------------------------------------------------------ budgeting
  // Planned vs ACTUAL hours for one role-week: actual clips the planned hours
  // (override or the 40 h default) to the week's workable days — 8 h per
  // non-holiday day. Actual is the basis for totals and cost.
  RM.roleWeekHours = function (state, m, w) {
    var iso = RM.fmtISO(RM.weekStartDate(state.meta, w));
    var planned = m.weekHours[iso] != null ? m.weekHours[iso] : 40;
    var workable = (5 - RM.holidaysInWeek(state.meta, w)) * 8;
    return { iso: iso, planned: planned, actual: Math.min(planned, workable) };
  };
  // Per-role ACTUAL hours over the whole project.
  RM.roleTotalHours = function (state, m) {
    var total = 0;
    for (var w = 0; w < state.meta.numWeeks; w++) total += RM.roleWeekHours(state, m, w).actual;
    return total;
  };
  // Margin: share of the bill rate kept after hourly cost, in %; null if no rate.
  RM.roleMargin = function (m) {
    if (!m.rate) return null;
    return (m.rate - (m.cost || 0)) / m.rate * 100;
  };
  // Average hourly COST of roster roles with the given team type; roles with
  // no cost set are ignored; falls back to the blended roster average.
  RM.avgCostRate = function (state, teamType) {
    var pool = state.team.filter(function (m) { return m.cost > 0 && (!teamType || m.type === teamType); });
    if (!pool.length && teamType) pool = state.team.filter(function (m) { return m.cost > 0; });
    if (!pool.length) return 0;
    var s = 0;
    pool.forEach(function (m) { s += m.cost; });
    return s / pool.length;
  };
  // Item effort: scheduled items use their painted working days, unscheduled
  // ones their size estimate. Hours = days × 8 × headcount; cost prices those
  // hours at the avg cost rate for the item's team type.
  RM.itemEffortInfo = function (state, it) {
    var days = it.startDay != null && it.durDays != null
      ? RM.workInSpan(state.meta, it.startDay, it.durDays)
      : (it.size ? RM.sizeDays(state, it.size) : 0);
    var hours = days * 8 * (it.headcount || 1);
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
        : mode === 'phase-ws' ? phaseName[it.phaseId] + ' · ' + (it.workstream || '(no workstream)')
        : (it.workstream || '(no workstream)');
      var b = bucket(key);
      b.items += 1; b.days += info.days; b.hours += info.hours; b.cost += info.cost;
    });
    if (mode === 'workstream') {
      state.team.forEach(function (m) {
        var b = bucket(m.workstream || '(no workstream)');
        var h = RM.roleTotalHours(state, m);
        b.roleHours = (b.roleHours || 0) + h;
        b.roleCost = (b.roleCost || 0) + h * (m.cost || 0);
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
