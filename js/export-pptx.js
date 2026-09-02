/*
 * Headway PPTX export — native, editable PowerPoint shapes.
 *  - slideShapes(lay) -> pure shape list in slide inches
 *    (node-testable) from an RM_EXPORT.layout() geometry.
 *  - toBlob(state, opts) -> .pptx blob via the vendored PptxGenJS
 *    (browser only). opts are RM_EXPORT.plan() opts: filters, grouping,
 *    byPhase / byWs splitting (one slide per plan entry).
 */
(function (root) {
  'use strict';

  var RM_EXPORT = root.RM_EXPORT ||
    (typeof require !== 'undefined' ? require('./export-png.js') : null);
  var PX = {};

  var SLIDE_W = 13.33, SLIDE_H = 7.5, MARGIN = 0.4;
  var INK = '182430', LINE = 'E5E0D5', BAND = '1A1F26';
  var SUBBAND = 'F2EFE8', HOLIDAY = 'EFECE4', SPRINT = 'E3EDF9';

  // approximate text width (px at the 10.5px bar-label font) — decides
  // whether a label fits inside its bar or sits beside it
  PX.labelPlacement = function (text, barWpx) {
    var wPx = String(text || '').length * 6.2 + 16;
    return { inside: wPx <= barWpx, wPx: wPx };
  };

  // the timeline always spreads to the full slide width; the vertical scale
  // shrinks independently when the roadmap is tall. A split export computes
  // ONE scale over all its slices so every slide shares the same row heights
  // and column widths.
  PX.slideScale = function (lays) {
    var maxW = 1, maxH = 1;
    lays.forEach(function (lay) {
      maxW = Math.max(maxW, lay.width);
      maxH = Math.max(maxH, lay.height);
    });
    var sx = (SLIDE_W - 2 * MARGIN) / maxW;
    return { sx: sx, sy: Math.min(sx, (SLIDE_H - 2 * MARGIN) / maxH) };
  };

  PX.slideShapes = function (lay, scale) {
    var sc = scale || PX.slideScale([lay]);
    var sx = sc.sx, sy = sc.sy;
    var ox = MARGIN;
    var oy = (SLIDE_H - lay.height * sy) / 2;
    function X(v) { return ox + v * sx; }
    function Y(v) { return oy + v * sy; }
    function fs(px) { return Math.max(6, Math.round(px * sy * 72)); }
    var shapes = [];

    var SPR_H = 22;
    var phBottom = 0;
    var top = SPR_H;

    // holiday-week shading under the rows
    lay.weeks.forEach(function (w) {
      if (!w.holiday) return;
      shapes.push({ type: 'rect', x: X(w.x), y: Y(top), w: lay.weekPx * sx,
        h: (lay.rowsBottom - top) * sy, color: HOLIDAY });
    });

    // sprint header
    lay.sprints.forEach(function (sp) {
      shapes.push({ type: 'rect', x: X(sp.x), y: Y(phBottom + 1), w: sp.w * sx,
        h: (SPR_H - 2) * sy, color: SPRINT });
      // dates only — sprint numbers add noise at slide scale
      shapes.push({ type: 'text', text: sp.date, x: X(sp.x + 4), y: Y(phBottom + 1),
        w: Math.max(4, sp.w - 8) * sx, h: (SPR_H - 2) * sy, size: fs(9.5),
        color: INK, align: 'left' });
    });

    // week gridlines
    function line(x1, y1, x2, y2, color, width) {
      shapes.push({ type: 'line', x: Math.min(x1, x2), y: Math.min(y1, y2),
        w: Math.abs(x2 - x1), h: Math.abs(y2 - y1),
        flipV: (x2 - x1) * (y2 - y1) < 0, color: color, width: width });
    }
    for (var i = 0; i <= lay.w1 - lay.w0; i++) {
      var gx = lay.laneX + i * lay.weekPx;
      line(X(gx), Y(top), X(gx), Y(lay.rowsBottom), LINE, 0.75);
    }

    // rows
    lay.rows.forEach(function (r) {
      if (r.kind === 'band') {
        shapes.push({ type: 'rect', x: X(0), y: Y(r.y), w: lay.width * sx, h: r.h * sy, color: BAND });
        shapes.push({ type: 'text', text: r.name, x: X(8), y: Y(r.y),
          w: lay.width * sx - 16 * sx, h: r.h * sy, size: fs(11), bold: true,
          color: 'F4F6F8', align: 'left' });
        return;
      }
      if (r.kind === 'wsband' || r.kind === 'eband') {
        var ind = r.kind === 'eband' && r.sub ? 14 : 0;
        shapes.push({ type: 'rect', x: X(0), y: Y(r.y), w: lay.width * sx, h: r.h * sy, color: SUBBAND });
        if (r.kind === 'wsband' && r.color)
          shapes.push({ type: 'rect', x: X(8), y: Y(r.y + r.h / 2 - 3.5), w: 7 * sy, h: 7 * sy, color: r.color });
        var nm = r.kind === 'eband' && !r.name ? 'no epic' : r.name;
        shapes.push({ type: 'text', text: nm + '   ' + r.count,
          x: X((r.kind === 'wsband' ? 20 : 8) + ind), y: Y(r.y),
          w: lay.width * sx - 30 * sx, h: r.h * sy, size: fs(10), bold: true,
          italic: r.kind === 'eband' && !r.name, color: INK, align: 'left' });
        return;
      }
      // item bar + label in / beside the bar; milestones are diamonds on
      // their start day, like the live timeline
      var b = r.bar;
      if (b.ms) {
        var md = 13 * sy;
        shapes.push({ type: b.msStyle === 'star' ? 'star' : b.msStyle === 'circle' ? 'circle' : 'diamond',
          x: X(b.x), y: Y(r.y + r.h / 2) - md / 2,
          w: md, h: md, color: b.color.replace(/^#/, ''), done: !!b.done });
        var mw = PX.labelPlacement(r.feature, 0).wPx * sy;
        var mx = X(b.x) + md + 4 * sy, malign = 'left';
        if (mx + mw > SLIDE_W - 0.05) { mx = Math.max(0.05, X(b.x) - 4 * sy - mw); malign = 'right'; }
        shapes.push({ type: 'text', text: r.feature, x: mx, y: Y(r.y + 4),
          w: mw, h: (r.h - 8) * sy, size: fs(10.5), color: INK, align: malign });
        return;
      }
      shapes.push({ type: 'bar', x: X(b.x), y: Y(r.y + 4), w: b.w * sx, h: (r.h - 8) * sy,
        color: b.color.replace(/^#/, ''), done: !!b.done });
      // fonts follow the vertical scale, bars the horizontal — compare in
      // a common unit when deciding whether the label fits inside the bar
      var pl = PX.labelPlacement(r.feature, b.w * sx / sy);
      if (pl.inside) {
        shapes.push({ type: 'text', text: r.feature, x: X(b.x + 4), y: Y(r.y + 4),
          w: b.w * sx - 8 * sy, h: (r.h - 8) * sy, size: fs(10.5), bold: true,
          color: 'FFFFFF', align: 'left' });
      } else {
        var lw = pl.wPx * sy;
        var lx = X(b.x) + b.w * sx + 5 * sy, align = 'left';
        if (lx + lw > SLIDE_W - 0.05) { lx = Math.max(0.05, X(b.x) - 5 * sy - lw); align = 'right'; }
        shapes.push({ type: 'text', text: r.feature, x: lx, y: Y(r.y + 4),
          w: lw, h: (r.h - 8) * sy, size: fs(10.5), color: INK, align: align });
      }
    });

    // workstream color legend
    (lay.legend || []).forEach(function (e) {
      shapes.push({ type: 'rect', x: X(e.x), y: Y(e.y + e.h / 2) - 5 * sy, w: 10 * sy, h: 10 * sy,
        color: e.color });
      shapes.push({ type: 'text', text: e.name, x: X(e.x + 16), y: Y(e.y),
        w: (e.name.length * 6.2 + 8) * sy, h: e.h * sy, size: fs(10.5), color: INK, align: 'left' });
    });

    // dependency connectors
    (lay.links || []).forEach(function (l) {
      line(X(l.x0), Y(l.y0), X(l.mid), Y(l.y0), '8D97A1', 0.75);
      line(X(l.mid), Y(l.y0), X(l.mid), Y(l.y1), '8D97A1', 0.75);
      line(X(l.mid), Y(l.y1), X(l.x1 - 2), Y(l.y1), '8D97A1', 0.75);
    });

    return { w: SLIDE_W, h: SLIDE_H, shapes: shapes };
  };

  PX.fileName = function (state) {
    return ((state.meta.title || '').replace(/[\\/:*?"<>|]+/g, '').trim() || 'Roadmap') + '.pptx';
  };

  // browser only: realize the slide plan with PptxGenJS and return a blob
  PX.toBlob = function (state, opts) {
    var Pptx = root.PptxGenJS;
    if (!Pptx) return Promise.reject(new Error('PowerPoint library not loaded'));
    var entries = RM_EXPORT.plan(state, opts);
    if (!entries.length) return Promise.reject(new Error('Nothing to export in this selection'));
    var pptx = new Pptx();
    pptx.defineLayout({ name: 'HW', width: SLIDE_W, height: SLIDE_H });
    pptx.layout = 'HW';
    var lays = entries.map(function (entry) { return RM_EXPORT.layout(state, entry.opts); });
    var scale = PX.slideScale(lays);
    lays.forEach(function (lay) {
      var sl = pptx.addSlide();
      sl.background = { color: 'FFFFFF' };
      PX.slideShapes(lay, scale).shapes.forEach(function (sh) {
        if (sh.type === 'line') {
          sl.addShape(pptx.ShapeType.line, {
            x: sh.x, y: sh.y, w: sh.w, h: sh.h, flipV: !!sh.flipV,
            line: { color: sh.color, width: sh.width }
          });
        } else if (sh.type === 'rect' || sh.type === 'bar' || sh.type === 'diamond' || sh.type === 'star' || sh.type === 'circle') {
          sl.addShape(sh.type === 'bar' ? pptx.ShapeType.roundRect :
            sh.type === 'diamond' ? pptx.ShapeType.diamond :
            sh.type === 'star' ? pptx.ShapeType.star5 :
            sh.type === 'circle' ? pptx.ShapeType.ellipse : pptx.ShapeType.rect, {
            x: sh.x, y: sh.y, w: Math.max(0.02, sh.w), h: Math.max(0.02, sh.h),
            fill: { color: sh.color, transparency: sh.done ? 55 : 0 },
            rectRadius: sh.type === 'bar' ? 0.03 : 0
          });
        } else if (sh.type === 'text') {
          sl.addText(sh.text, {
            x: sh.x, y: sh.y, w: Math.max(0.1, sh.w), h: Math.max(0.1, sh.h),
            fontSize: sh.size, bold: !!sh.bold, italic: !!sh.italic,
            color: sh.color, align: sh.align || 'left', valign: 'middle',
            margin: 0, wrap: false, fontFace: 'Helvetica Neue'
          });
        }
      });
    });
    return pptx.write({ outputType: 'blob' }).then(function (blob) {
      return { blob: blob, name: PX.fileName(state) };
    });
  };

  root.RM_PPTX = PX;
  if (typeof module !== 'undefined' && module.exports) module.exports = PX;
})(typeof window !== 'undefined' ? window : globalThis);
