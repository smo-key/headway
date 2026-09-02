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

  var LEFT_W = 0;     // no label rail — labels ride the bars
  var SPR_H = 22;     // date header lane (the only header row)
  var BAND_H = 30;    // phase band row
  var EBAND_H = 22;   // workstream / epic sub-band row
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

    // a locked window (a split export sharing one timeline across slices)
    // wins; otherwise the selected range AND the filtered content together
    // bound the window: clamp to the span of visible bars (whole weeks);
    // nothing visible → no date columns at all
    if (opts.lockW0 != null) {
      w0 = opts.lockW0;
      w1 = opts.lockW1;
    } else {
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
    }
    d0 = w0 * S;
    d1 = w1 * S;

    function passesFilters(it) {
      if (opts.ws && it.workstream !== opts.ws) return false;
      // wsKey is an exact workstream match where '' means "no workstream"
      if (opts.wsKey != null && (it.workstream || '') !== opts.wsKey) return false;
      if (opts.epic && it.epic !== opts.epic) return false;
      return true;
    }
    function visibleIn(it, a, b) {
      if (it.startDay == null || !passesFilters(it)) return false;
      var e = it.startDay + (it.durDays || 0) + (it.riskDays || 0);
      return e > a && it.startDay < b;
    }
    function visible(it) { return visibleIn(it, d0, d1); }

    var visPhases = [];
    state.phases.forEach(function (p) {
      if (opts.phaseId && p.id !== opts.phaseId) return;
      var items = RM.itemsInPhase(state, p.id).filter(visible);
      if (items.length) visPhases.push({ p: p, items: items });
    });

    // grouping hierarchy mirrors the timeline: phase > workstream > epic
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

    var rows = [];
    var y = SPR_H;
    var groupWs = !!opts.groupWs && meta.workstreamsEnabled !== false;
    visPhases.forEach(function (v) {
      rows.push({ kind: 'band', name: v.p.name, count: v.items.length, y: y, h: BAND_H });
      y += BAND_H;
      function pushItem(it) {
        var s = Math.max(it.startDay, d0);
        var e = Math.min(it.startDay + (it.durDays || 0) + (it.riskDays || 0), d1);
        rows.push({
          kind: 'item', id: it.id, feature: it.feature, y: y, h: ROW_H,
          bar: {
            x: LEFT_W + (s - d0) * dpx,
            w: Math.max(6, (e - s) * dpx),
            color: '#' + RM.colorForItem(state, it),
            done: !!it.done,
            ms: !!it.milestone,
            msStyle: RM.msStyleOf(it)
          }
        });
        y += ROW_H;
      }
      function pushEpicBands(list, sub) {
        var g = partition(list, 'epic', false);
        g.keys.forEach(function (key) {
          rows.push({ kind: 'eband', name: key, count: g.by[key].length, y: y, h: EBAND_H, sub: !!sub });
          y += EBAND_H;
          g.by[key].forEach(pushItem);
        });
      }
      if (groupWs) {
        var wg = partition(v.items, 'workstream', true);
        wg.keys.forEach(function (key) {
          rows.push({
            kind: 'wsband', name: key || RM.defaultWsName(state), count: wg.by[key].length,
            color: RM.colorForWs(state, key), y: y, h: EBAND_H
          });
          y += EBAND_H;
          if (opts.groupEpic) pushEpicBands(wg.by[key], true);
          else wg.by[key].forEach(pushItem);
        });
      } else if (opts.groupEpic) {
        pushEpicBands(v.items, false);
      } else {
        v.items.forEach(pushItem);
      }
    });

    // dependency connectors between visible bars (elbow segments, bar-relative)
    var links = [];
    if (opts.arrows) {
      var rowByItem = {};
      rows.forEach(function (r) { if (r.kind === 'item') rowByItem[r.id] = r; });
      var byNum = {};
      state.items.forEach(function (o) { byNum[o.num] = o; });
      state.items.forEach(function (it) {
        var to = rowByItem[it.id];
        if (!to) return;
        (it.deps || []).forEach(function (n) {
          var dep = byNum[n];
          var from = dep && rowByItem[dep.id];
          if (!from) return;
          var x0 = from.bar.x + from.bar.w, y0 = from.y + from.h / 2;
          var x1 = to.bar.x, y1 = to.y + to.h / 2;
          links.push({ x0: x0, y0: y0, x1: x1, y1: y1, mid: Math.max(x0 + 6, x1 - 6) });
        });
      });
    }

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

    var rowsBottom = y;

    // legend: one swatch per visible workstream (bar colors ARE workstream
    // colors), default workstream last. Entry positions are precomputed with
    // an approximate glyph width so both renderers lay them out identically.
    var legend = [];
    if (meta.workstreamsEnabled !== false) {
      var seenWs = {}, wsOrder = [];
      visPhases.forEach(function (v) {
        v.items.forEach(function (it) {
          var k = it.workstream || '';
          if (!seenWs[k]) { seenWs[k] = true; wsOrder.push(k); }
        });
      });
      if (wsOrder.indexOf('') !== -1) {
        wsOrder = wsOrder.filter(function (k) { return k !== ''; }).concat(['']);
      }
      var width = LEFT_W + (w1 - w0) * weekPx;
      var lx = 10, ly = y + 8, LG_H = 18;
      wsOrder.forEach(function (k) {
        var name = k || RM.defaultWsName(state);
        var w = 16 + name.length * 6.2 + 18; // swatch + gap + name + spacing
        if (lx > 10 && lx + w > width - 10) { lx = 10; ly += LG_H; }
        legend.push({ name: name, color: RM.colorForWs(state, k), x: lx, y: ly, h: LG_H });
        lx += w;
      });
      if (legend.length) y = ly + LG_H;
    }

    return {
      weekPx: weekPx, laneX: LEFT_W, d0: d0, d1: d1, w0: w0, w1: w1,
      width: LEFT_W + (w1 - w0) * weekPx,
      height: y + PAD, legend: legend, rowsBottom: rowsBottom,
      rows: rows, links: links, sprints: sprints, weeks: weeks,
      title: meta.title || 'Roadmap',
      range: w1 > w0 ? RM.fmtShort(RM.dayToDate(meta, d0)) + ' → ' +
        RM.fmtShort(RM.dayToDate(meta, d1 - 1)) : ''
    };
  };

  // ------------------------------------------------------------ export plan
  // How one export request becomes one or more images / slides.
  // opts.byPhase / opts.byWs split the export; every entry carries the layout
  // opts for that slice plus a human name. Empty slices are dropped.
  EX.plan = function (state, opts) {
    opts = opts || {};
    var base = {};
    ['weekPx', 'scale', 'fromSprint', 'toSprint', 'ws', 'wsKey', 'epic', 'phaseId',
      'arrows', 'groupWs', 'groupEpic'].forEach(function (k) {
      if (opts[k] != null) base[k] = opts[k];
    });
    function hasItems(o) {
      return EX.layout(state, o).rows.some(function (r) { return r.kind === 'item'; });
    }
    if (!opts.byPhase && !opts.byWs) {
      return [{ name: state.meta.title || 'Roadmap', opts: base }];
    }
    var phases = opts.byPhase
      ? state.phases.filter(function (p) { return !base.phaseId || p.id === base.phaseId; })
      : [null];
    var wsKeys = [null];
    if (opts.byWs) {
      // workstream keys in item order, the empty/default workstream last
      wsKeys = [];
      var seen = {};
      state.items.forEach(function (it) {
        var k = it.workstream || '';
        if (!seen[k]) { seen[k] = true; wsKeys.push(k); }
      });
      if (wsKeys.indexOf('') !== -1) {
        wsKeys = wsKeys.filter(function (k) { return k !== ''; }).concat(['']);
      }
      if (!wsKeys.length) wsKeys = [''];
    }
    // every slice shares the UNSPLIT export's window so rows and columns
    // line up across the slides / images
    var baseLay = EX.layout(state, base);
    var entries = [];
    phases.forEach(function (p) {
      wsKeys.forEach(function (wk) {
        var o = {};
        Object.keys(base).forEach(function (k) { o[k] = base[k]; });
        if (p) o.phaseId = p.id;
        if (wk != null) o.wsKey = wk;
        if (!hasItems(o)) return;
        o.lockW0 = baseLay.w0;
        o.lockW1 = baseLay.w1;
        var parts = [];
        if (p) parts.push(p.name);
        if (wk != null) parts.push(wk || RM.defaultWsName(state));
        entries.push({ name: parts.join(' — '), opts: o });
      });
    });
    return entries;
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
    ctx.textBaseline = 'middle';

    var phBottom = 0;
    var top = SPR_H;

    // holiday-week shading under the rows
    lay.weeks.forEach(function (wcell) {
      if (!wcell.holiday) return;
      ctx.fillStyle = '#EFECE4';
      ctx.fillRect(wcell.x, top, lay.weekPx, lay.rowsBottom - top);
    });

    // sprint header
    lay.sprints.forEach(function (sp) {
      ctx.fillStyle = '#E3EDF9';
      ctx.fillRect(sp.x + 1, phBottom + 1, sp.w - 2, SPR_H - 2);
      var cy = phBottom + SPR_H / 2;
      ctx.fillStyle = INK;
      ctx.font = '700 10px ' + FONT;
      ctx.fillText(sp.label, sp.x + 5, cy);
      if (sp.w >= 52) {
        ctx.fillStyle = INK3;
        ctx.font = '9px ' + FONT;
        ctx.fillText(sp.date, sp.x + 5 + ctx.measureText(sp.label).width + 14, cy);
      }
    });

    // week gridlines
    for (var i = 0; i <= lay.w1 - lay.w0; i++) {
      var x = lay.laneX + i * lay.weekPx;
      ctx.strokeStyle = LINE;
      ctx.lineWidth = 1;
      ctx.beginPath();
      ctx.moveTo(x + 0.5, top);
      ctx.lineTo(x + 0.5, lay.rowsBottom);
      ctx.stroke();
    }

    // rows
    lay.rows.forEach(function (r) {
      var cy = r.y + r.h / 2;
      if (r.kind === 'band') {
        ctx.fillStyle = BAND;
        ctx.fillRect(0, r.y, lay.width, r.h);
        ctx.fillStyle = '#F4F6F8';
        ctx.font = '700 11px ' + FONT;
        ctx.fillText(r.name, 10, cy);
        return;
      }
      if (r.kind === 'wsband' || r.kind === 'eband') {
        ctx.fillStyle = '#F2EFE8';
        ctx.fillRect(0, r.y, lay.width, r.h);
        var tx = 10 + (r.sub ? 14 : 0);
        if (r.kind === 'wsband') {
          ctx.fillStyle = '#' + r.color;
          ctx.beginPath();
          ctx.arc(tx + 4, cy, 4, 0, Math.PI * 2);
          ctx.fill();
          tx += 12;
        }
        ctx.fillStyle = INK;
        ctx.font = (r.kind === 'eband' && !r.name ? 'italic ' : '') + '700 10px ' + FONT;
        var nm = r.kind === 'eband' && !r.name ? 'no epic' : r.name;
        ctx.fillText(nm, tx, cy);
        ctx.fillStyle = INK3;
        ctx.font = '10px ' + FONT;
        ctx.fillText(String(r.count), tx + ctx.measureText(nm).width * 1.28 + 8, cy);
        return;
      }
      ctx.strokeStyle = LINE;
      ctx.beginPath();
      ctx.moveTo(0, r.y + r.h + 0.5);
      ctx.lineTo(lay.width, r.y + r.h + 0.5);
      ctx.stroke();

      var b = r.bar;
      ctx.fillStyle = b.color;
      ctx.globalAlpha = b.done ? 0.45 : 1;
      if (b.ms) {
        // a diamond on the start day, like the live timeline
        var mcx = b.x + 6, mcy = r.y + r.h / 2, mr = 6.5;
        ctx.beginPath();
        if (b.msStyle === 'circle') {
          ctx.arc(mcx, mcy, mr, 0, Math.PI * 2);
        } else if (b.msStyle === 'star') {
          for (var sp = 0; sp < 10; sp++) {
            var sr = sp % 2 ? mr * 0.45 : mr * 1.15, sa = -Math.PI / 2 + sp * Math.PI / 5;
            ctx[sp ? 'lineTo' : 'moveTo'](mcx + Math.cos(sa) * sr, mcy + Math.sin(sa) * sr);
          }
        } else {
          ctx.moveTo(mcx, mcy - mr);
          ctx.lineTo(mcx + mr, mcy);
          ctx.lineTo(mcx, mcy + mr);
          ctx.lineTo(mcx - mr, mcy);
        }
        ctx.closePath();
        ctx.fill();
        ctx.globalAlpha = 1;
        ctx.font = '600 10.5px ' + FONT;
        ctx.fillStyle = INK;
        var mtw = ctx.measureText(r.feature).width;
        if (mcx + mr + 7 + mtw <= lay.width - 4) ctx.fillText(r.feature, mcx + mr + 7, cy);
        else ctx.fillText(r.feature, Math.max(4, mcx - mr - 7 - mtw), cy);
        return;
      }
      roundRect(ctx, b.x, r.y + 4, b.w, r.h - 8, 4);
      ctx.fill();
      ctx.globalAlpha = 1;

      // the label rides the bar: inside when it fits, else just beside it
      ctx.font = '700 10.5px ' + FONT;
      var tw = ctx.measureText(r.feature).width;
      if (tw <= b.w - 16) {
        ctx.fillStyle = '#FFFFFF';
        ctx.fillText(r.feature, b.x + 8, cy);
      } else {
        ctx.font = '600 10.5px ' + FONT;
        tw = ctx.measureText(r.feature).width;
        ctx.fillStyle = INK;
        if (b.x + b.w + 7 + tw <= lay.width - 4) {
          ctx.fillText(r.feature, b.x + b.w + 7, cy);
        } else if (b.x - 7 - tw >= 4) {
          ctx.fillText(r.feature, b.x - 7 - tw, cy);
        } else {
          // no room either side — inside the bar, ellipsized
          ctx.font = '700 10.5px ' + FONT;
          ctx.fillStyle = '#FFFFFF';
          var label = r.feature;
          while (label && ctx.measureText(label + '…').width > b.w - 16) label = label.slice(0, -2);
          if (label) ctx.fillText(label + '…', b.x + 8, cy);
        }
      }
    });

    // workstream color legend
    (lay.legend || []).forEach(function (e) {
      var cy = e.y + e.h / 2;
      ctx.fillStyle = '#' + e.color;
      roundRect(ctx, e.x, cy - 5, 10, 10, 3);
      ctx.fill();
      ctx.fillStyle = INK;
      ctx.font = '10.5px ' + FONT;
      ctx.fillText(e.name, e.x + 16, cy);
    });

    // dependency connectors between visible bars (geometry from the layout)
    ctx.strokeStyle = '#8D97A1';
    ctx.lineWidth = 1;
    (lay.links || []).forEach(function (l) {
      // plain connector — no arrowhead (matches the in-app style)
      ctx.beginPath();
      ctx.moveTo(l.x0, l.y0);
      ctx.lineTo(l.mid, l.y0);
      ctx.lineTo(l.mid, l.y1);
      ctx.lineTo(l.x1 - 2, l.y1);
      ctx.stroke();
    });

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
