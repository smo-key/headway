/*
 * Headway PNG export.
 *  - layout(state, opts) -> pure geometry (node-testable): rows, bars, header
 *    cells for the selected sprint window and filters.
 *  - render(state, opts) -> <canvas> painted from that layout (browser only).
 *  - download(state, opts) -> renders and saves a .png named after the doc.
 *  opts: { weekPx, scale, fromSprint, toSprint, ws, epic, phaseId, arrows }
 *  Sprint bounds are inclusive sprint NUMBERS (anchor-aware, like the header).
 */
(function (root) {
  'use strict';

  var RM = root.RM || (typeof require !== 'undefined' ? require('./core.js') : null);
  var EX = {};

  var LEFT_W = 240;   // label column
  var TITLE_H = 46;   // title + date-range block
  var PH_H = 20;      // phase span lane
  var SPR_H = 22;     // sprint header lane
  var BAND_H = 30;    // phase band row
  var ROW_H = 26;     // item row
  var PAD = 14;       // bottom padding

  EX.layout = function (state, opts) {
    opts = opts || {};
    var meta = state.meta;
    var weekPx = opts.weekPx || 28;
    var dpx = weekPx / RM.slotsOf(meta);

    // week window from inclusive sprint numbers (anchor-aware)
    var w0 = 0, w1 = meta.numWeeks;
    if (opts.fromSprint != null || opts.toSprint != null) {
      var lo = null, hi = null;
      for (var w = 0; w < meta.numWeeks; w++) {
        var n = RM.sprintNumForWeek(meta, w);
        if (opts.fromSprint != null && n < opts.fromSprint) continue;
        if (opts.toSprint != null && n > opts.toSprint) continue;
        if (lo == null) lo = w;
        hi = w + 1;
      }
      if (lo != null) { w0 = lo; w1 = hi; }
    }
    var S = RM.slotsOf(meta);
    var d0 = w0 * S, d1 = w1 * S;

    // the selected range AND the filtered content together bound the window:
    // clamp to the span of visible bars (whole weeks); nothing visible → no
    // date columns at all
    var lo = null, hi = null;
    state.items.forEach(function (it) {
      if (opts.phaseId && it.phaseId !== opts.phaseId) return;
      if (!visibleIn(it, d0, d1)) return;
      var e = it.startDay + (it.durDays || 0) + (it.riskDays || 0);
      if (lo == null || it.startDay < lo) lo = it.startDay;
      if (hi == null || e > hi) hi = e;
    });
    if (lo == null) { w1 = w0; }
    else {
      w0 = Math.max(w0, Math.floor(lo / S));
      w1 = Math.min(w1, Math.ceil(hi / S));
    }
    d0 = w0 * S;
    d1 = w1 * S;

    function visibleIn(it, a, b) {
      if (it.startDay == null) return false;
      if (opts.ws && it.workstream !== opts.ws) return false;
      if (opts.epic && it.epic !== opts.epic) return false;
      var e = it.startDay + (it.durDays || 0) + (it.riskDays || 0);
      return e > a && it.startDay < b;
    }
    function visible(it) {
      if (it.startDay == null) return false;
      if (opts.ws && it.workstream !== opts.ws) return false;
      if (opts.epic && it.epic !== opts.epic) return false;
      var e = it.startDay + (it.durDays || 0) + (it.riskDays || 0);
      return e > d0 && it.startDay < d1;
    }

    var visPhases = [];
    state.phases.forEach(function (p) {
      if (opts.phaseId && p.id !== opts.phaseId) return;
      var items = RM.itemsInPhase(state, p.id).filter(visible);
      if (items.length) visPhases.push({ p: p, items: items });
    });

    // phase spans stack into lanes when they overlap (like the live header)
    var phaseSpans = [], laneEnds = [];
    visPhases.forEach(function (v) {
      var span = RM.phaseSpan(state, v.p);
      if (!span || span.hi <= d0 || span.lo >= d1) return;
      var lo2 = Math.max(span.lo, d0), hi2 = Math.min(span.hi, d1);
      var sp = { name: v.p.name, x: LEFT_W + (lo2 - d0) * dpx, w: (hi2 - lo2) * dpx, lane: 0 };
      while (sp.lane < laneEnds.length && laneEnds[sp.lane] > sp.x) sp.lane++;
      laneEnds[sp.lane] = sp.x + sp.w;
      phaseSpans.push(sp);
    });
    var phLanes = Math.max(1, laneEnds.length);

    var rows = [];
    var y = TITLE_H + phLanes * PH_H + SPR_H;
    visPhases.forEach(function (v) {
      rows.push({ kind: 'band', name: v.p.name, count: v.items.length, y: y, h: BAND_H });
      y += BAND_H;
      v.items.forEach(function (it) {
        var s = Math.max(it.startDay, d0);
        var e = Math.min(it.startDay + (it.durDays || 0) + (it.riskDays || 0), d1);
        rows.push({
          kind: 'item', id: it.id, feature: it.feature, y: y, h: ROW_H,
          bar: {
            x: LEFT_W + (s - d0) * dpx,
            w: Math.max(6, (e - s) * dpx),
            color: '#' + RM.colorForItem(state, it),
            done: !!it.done
          }
        });
        y += ROW_H;
      });
    });

    // sprint header cells across the window
    var sprints = [];
    for (var wk = w0; wk < w1; wk++) {
      var num = RM.sprintNumForWeek(meta, wk);
      var last = sprints[sprints.length - 1];
      if (last && last.num === num) last.w += weekPx;
      else sprints.push({
        num: num, label: RM.sprintsEnabled(meta) ? 'S' + num : '',
        x: LEFT_W + (wk - w0) * weekPx, w: weekPx,
        date: RM.fmtShort(RM.dayToDate(meta, wk * RM.slotsOf(meta)))
      });
    }

    // per-week shading (full holiday weeks)
    var hset = RM.holidayDaySet(meta);
    var weeks = [];
    for (var w2 = w0; w2 < w1; w2++) {
      var all = true;
      var S2 = RM.slotsOf(meta);
      for (var d = w2 * S2; d < w2 * S2 + S2; d++) if (!hset[d]) { all = false; break; }
      weeks.push({ x: LEFT_W + (w2 - w0) * weekPx, holiday: all });
    }

    return {
      weekPx: weekPx, laneX: LEFT_W, d0: d0, d1: d1, w0: w0, w1: w1, phLanes: phLanes,
      width: LEFT_W + (w1 - w0) * weekPx,
      height: y + PAD,
      rows: rows, sprints: sprints, phaseSpans: phaseSpans, weeks: weeks,
      title: meta.title || 'Roadmap',
      range: w1 > w0 ? RM.fmtShort(RM.dayToDate(meta, d0)) + ' → ' +
        RM.fmtShort(RM.dayToDate(meta, d1 - 1)) : ''
    };
  };

  // ------------------------------------------------------------ painting
  var INK = '#182430', INK3 = '#6E7883', LINE = '#E5E0D5', BAND = '#1A1F26';
  var FONT = '"Helvetica Neue", Arial, sans-serif';

  EX.render = function (state, opts) {
    opts = opts || {};
    var lay = EX.layout(state, opts);
    var scale = opts.scale || 2;
    var canvas = root.document.createElement('canvas');
    canvas.width = lay.width * scale;
    canvas.height = lay.height * scale;
    var ctx = canvas.getContext('2d');
    ctx.scale(scale, scale);

    ctx.fillStyle = '#FFFFFF';
    ctx.fillRect(0, 0, lay.width, lay.height);

    // title + range
    ctx.fillStyle = INK;
    ctx.font = '700 15px ' + FONT;
    ctx.fillText(lay.title, 12, 22);
    ctx.fillStyle = INK3;
    ctx.font = '11px ' + FONT;
    ctx.fillText(lay.range, 12, 38);

    var phBottom = TITLE_H + lay.phLanes * PH_H;
    var top = phBottom + SPR_H;

    // holiday-week shading under the rows
    lay.weeks.forEach(function (wcell) {
      if (!wcell.holiday) return;
      ctx.fillStyle = '#EFECE4';
      ctx.fillRect(wcell.x, top, lay.weekPx, lay.height - top - PAD);
    });

    // phase span lane
    lay.phaseSpans.forEach(function (sp) {
      var py = TITLE_H + sp.lane * PH_H;
      ctx.fillStyle = BAND;
      ctx.fillRect(sp.x + 1, py + 2, Math.max(2, sp.w - 2), PH_H - 6);
      ctx.fillStyle = '#F4F6F8';
      ctx.font = '700 10px ' + FONT;
      var name = sp.name;
      while (name && ctx.measureText(name).width > sp.w - 12) name = name.slice(0, -2);
      if (name) ctx.fillText(name === sp.name ? name : name + '…', sp.x + 6, py + PH_H - 9);
    });

    // sprint header
    lay.sprints.forEach(function (sp) {
      ctx.fillStyle = '#E3EDF9';
      ctx.fillRect(sp.x + 1, phBottom + 1, sp.w - 2, SPR_H - 2);
      ctx.fillStyle = INK;
      ctx.font = '700 10px ' + FONT;
      ctx.fillText(sp.label, sp.x + 5, phBottom + 14);
      if (sp.w >= 52) {
        ctx.fillStyle = INK3;
        ctx.font = '9px ' + FONT;
        ctx.fillText(sp.date, sp.x + 5 + ctx.measureText(sp.label).width + 14, phBottom + 14);
      }
    });

    // week gridlines
    for (var i = 0; i <= lay.w1 - lay.w0; i++) {
      var x = lay.laneX + i * lay.weekPx;
      ctx.strokeStyle = LINE;
      ctx.lineWidth = 1;
      ctx.beginPath();
      ctx.moveTo(x + 0.5, top);
      ctx.lineTo(x + 0.5, lay.height - PAD);
      ctx.stroke();
    }

    // rows
    var barByItem = {};
    lay.rows.forEach(function (r) {
      if (r.kind === 'band') {
        ctx.fillStyle = BAND;
        ctx.fillRect(0, r.y, lay.width, r.h);
        ctx.fillStyle = '#F4F6F8';
        ctx.font = '700 11px ' + FONT;
        ctx.fillText(r.name, 10, r.y + 19);
        ctx.fillStyle = '#8D97A1';
        ctx.font = '10px ' + FONT;
        ctx.fillText(String(r.count), 16 + ctx.measureText(r.name).width * 1.28, r.y + 19);
        return;
      }
      ctx.strokeStyle = LINE;
      ctx.beginPath();
      ctx.moveTo(0, r.y + r.h + 0.5);
      ctx.lineTo(lay.width, r.y + r.h + 0.5);
      ctx.stroke();

      ctx.fillStyle = INK;
      ctx.font = '11px ' + FONT;
      var label = r.feature;
      while (label && ctx.measureText(label).width > lay.laneX - 20) label = label.slice(0, -2);
      ctx.fillText(label === r.feature ? label : label + '…', 10, r.y + 17);

      var b = r.bar;
      barByItem[r.id] = { x: b.x, y: r.y + 4, w: b.w, h: r.h - 8 };
      ctx.fillStyle = b.color;
      ctx.globalAlpha = b.done ? 0.45 : 1;
      roundRect(ctx, b.x, r.y + 4, b.w, r.h - 8, 4);
      ctx.fill();
      ctx.globalAlpha = 1;
    });

    // dependency arrows between visible bars
    if (opts.arrows) {
      ctx.strokeStyle = '#8D97A1';
      ctx.fillStyle = '#8D97A1';
      ctx.lineWidth = 1;
      state.items.forEach(function (it) {
        var to = barByItem[it.id];
        if (!to) return;
        (it.deps || []).forEach(function (n) {
          var dep = null;
          state.items.forEach(function (o) { if (o.num === n) dep = o; });
          var from = dep && barByItem[dep.id];
          if (!from) return;
          var x0 = from.x + from.w, y0 = from.y + from.h / 2;
          var x1 = to.x, y1 = to.y + to.h / 2;
          var mid = Math.max(x0 + 6, x1 - 6);
          // plain connector — no arrowhead (matches the in-app style)
          ctx.beginPath();
          ctx.moveTo(x0, y0);
          ctx.lineTo(mid, y0);
          ctx.lineTo(mid, y1);
          ctx.lineTo(x1 - 2, y1);
          ctx.stroke();
        });
      });
    }

    // lane separator
    ctx.strokeStyle = '#CFC9BB';
    ctx.beginPath();
    ctx.moveTo(lay.laneX + 0.5, TITLE_H);
    ctx.lineTo(lay.laneX + 0.5, lay.height - PAD);
    ctx.stroke();

    return canvas;
  };

  function roundRect(ctx, x, y, w, h, r) {
    r = Math.min(r, w / 2, h / 2);
    ctx.beginPath();
    ctx.moveTo(x + r, y);
    ctx.arcTo(x + w, y, x + w, y + h, r);
    ctx.arcTo(x + w, y + h, x, y + h, r);
    ctx.arcTo(x, y + h, x, y, r);
    ctx.arcTo(x, y, x + w, y, r);
    ctx.closePath();
  }

  // the exact project title, minus filesystem-hostile characters only —
  // no lowercasing, no hyphenation, no date
  EX.fileName = function (state) {
    return ((state.meta.title || '').replace(/[\\/:*?"<>|]+/g, '').trim() || 'Roadmap') + '.png';
  };

  // render to a PNG blob; the caller decides where it goes
  EX.toBlob = function (state, opts) {
    var canvas = EX.render(state, opts);
    return new Promise(function (resolve, reject) {
      canvas.toBlob(function (blob) {
        if (blob) resolve({ blob: blob, name: EX.fileName(state) });
        else reject(new Error('Could not render the PNG'));
      }, 'image/png');
    });
  };

  // plain-browser fallback: a regular download
  EX.download = function (state, opts) {
    return EX.toBlob(state, opts).then(function (r) {
      var a = root.document.createElement('a');
      a.href = URL.createObjectURL(r.blob);
      a.download = r.name;
      a.click();
      setTimeout(function () { URL.revokeObjectURL(a.href); }, 5000);
      return null;
    });
  };

  root.RM_EXPORT = EX;
  if (typeof module !== 'undefined' && module.exports) module.exports = EX;
})(typeof window !== 'undefined' ? window : globalThis);
