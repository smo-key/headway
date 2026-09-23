/*
 * Headway — in-app "How to use" guide.
 *
 * Content for the start page's How-to area: a short tutorial (Start here),
 * Lean and literal: short imperative sentences, menu paths, no filler.
 * one line per view, the keyboard shortcuts and a features list.
 * Pure data + HTML builders, no DOM access — app.js renders it and wires the
 * tabs. Keep entries short: this is a field guide, not the manual (README.md
 * and DESIGN.md hold the detail).
 */
(function (root, factory) {
  if (typeof module === 'object' && module.exports) module.exports = factory();
  else root.HeadwayGuide = factory();
}(typeof self !== 'undefined' ? self : this, function () {
  'use strict';

  var mac = typeof navigator !== 'undefined' && /Mac|iPhone|iPad/.test(navigator.platform || '');
  var MOD = mac ? '⌘' : 'Ctrl';   // the app accepts both; show the viewer's own key

  function esc(s) {
    return String(s == null ? '' : s).replace(/[&<>"]/g, function (c) { return { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]; });
  }
  // inline `key` → <kbd>, **bold** → <b>
  function rich(s) {
    return esc(s)
      .replace(/`([^`]+)`/g, function (m, k) { return '<kbd class="kbd">' + k.replace(/Mod/g, MOD) + '</kbd>'; })
      .replace(/\*\*([^*]+)\*\*/g, '<b>$1</b>');
  }

  var TABS = [
    { id: 'start', label: 'Start here', icon: 'rocket' },
    { id: 'views', label: 'Views', icon: 'layout-grid' },
    { id: 'keys', label: 'Shortcuts', icon: 'keyboard' },
    { id: 'ideas', label: 'Features', icon: 'sparkles' }
  ];

  // Start here — the one path every new plan takes, in order
  var START = [
    { t: 'Create the project', d: '**New project…** sets the name and start date. **Setup**: end date, sprints, workstreams, phases, team, holidays.' },
    { t: 'Add features', d: 'Edit → **Add feature**, or the **Add feature** row under a phase. Click a row to see it in the panel: description, size, priority, tags, stories.' },
    { t: 'Schedule', d: '**Planning**: double-click the empty lane to place a bar. Drag to move, drag an edge to resize, or type the weeks in the chip.' },
    { t: 'Link dependencies', d: 'Hover a bar, drag its edge **circle** onto another bar. Orange = critical path, dashed amber = violation. ⚡ **Auto timeline** on a phase band lays the phase out; right-click a row → **Place at earliest slot** for one item.' },
    { t: 'Save and share', d: '`Mod+S` saves .xlsx. File → **Convert to shared folder…** makes a .headway folder for OneDrive. **Export…**: PNG, PowerPoint, Jira CSV.' }
  ];

  var VIEWS = [
    { t: 'Setup', d: 'Timeline, sprints, workstreams, phases, team, sizing, estimate mode, capacity, holidays, which tabs are on.' },
    { t: 'Planning', d: 'Gantt: bars, story bars, dependencies, phase lane, capacity rows, Resources panel. Drag empty space to pan, `Mod+scroll` to zoom.' },
    { t: 'Scoping', d: 'Spreadsheet of the rows: size, risk, weeks, workstream, epic, text columns. **+** adds columns.' },
    { t: 'Prioritizing', d: 'Kanban by phase, Priority, Size or Risk. Drag a card to set the field. **Story** level shows a story board.' },
    { t: 'Sprinting', d: 'Rows grouped by the sprint they start in. Drag to another sprint, or right-click → **Move to sprint…**' },
    { t: 'Budgeting', d: 'Cost and rate per role, margin, total, week hours. **Reports** drawer: effort and cost by workstream or phase.' }
  ];

  var KEYS = [
    { g: 'Files & history', k: [['Mod+S', 'Save'], ['Shift+Mod+S', 'Save as…'], ['Mod+Z', 'Undo'], ['Shift+Mod+Z', 'Redo']] },
    { g: 'Timeline', k: [['Double-click lane', 'Place a bar at that date'], ['Drag bar · drag edge', 'Move · resize'], ['Mod+drag bar', 'Move it and every dependent'], ['← →', 'Nudge the selected bar a day'], ['Shift+← →', 'Nudge five days'], ['Drag empty space', 'Pan'], ['Mod+scroll', 'Zoom around the cursor']] },
    { g: 'Selection', k: [['Click a row', 'Show it in the detail panel'], ['Right-click', 'Row, bar, band and card menus'], ['Delete', 'Delete the selected item or arrow'], ['Esc', 'Close a menu · cancel a dependency drag'], ['[ · ]', 'Collapse the left pane · the right panel']] },
    { g: 'Editing', k: [['Enter', 'Commit a field'], ['Mod+B · Mod+I', 'Bold · italic in rich text'], ['"- " or "1. "', 'Start a bullet or numbered list'], ['Tab · Shift+Tab', 'Hop resource and budget cells'], ['Shift+click ×N', 'Headcount −1 (click adds one)'], ['Mod+F', 'Filter the board or sprint list'], ['Mod+J', 'AI assistant']] }
  ];

  var FEATURES = [
    { t: 'Range estimates', d: 'Setup → Sizing → **Range estimate**: low and high per feature and story. The bar follows the basis (high or low); the hatched band shows the rest. View → **Estimate ranges** hides the bands. 4-day weeks: rate 1.25.' },
    { t: 'Plans', d: 'Plan menu next to the title: alternate versions of the roadmap. Switch, compare (ghost bars), add.' },
    { t: 'Shared roadmap folder', d: '.headway folder on OneDrive or SharePoint: one file per feature, phase and person. Presence shows who is in; Version history shows who changed what.' },
    { t: 'Colour and grouping', d: 'View → **Color by** workstream, epic, priority or type. **Group by** workstream or epic. Workstream colours: the pencil in any workstream dropdown.' },
    { t: 'Milestones, flags, locks', d: 'Zero-duration item = milestone (diamond, star, circle). Right-click → **Flag…** with a reason. **Lock** pins a bar; **Done** greys it out.' },
    { t: 'Capacity and cost', d: 'Setup → Capacity: hours per person per week in the Resources panel, capacity rows, Auto timeline that never overbooks. Budgeting prices it.' },
    { t: 'Checks', d: 'Preflight chip: cycles, unknown dependencies, starts inside a buffer, missing sizes, over-capacity weeks. Each row shows its own alert.' },
    { t: 'Sizes from stories', d: 'Setup → Sizing → **Roll up from stories**. Right-click a story → **Move to feature…**' },
    { t: 'Standalone HTML', d: 'Export → **Standalone HTML**: one view-only page with every tab.' },
    { t: 'AI assistant', d: '`Mod+J`. Ask about the plan, or tell it to re-tag, re-phase, re-size. Edits are undoable and logged as “you · AI”.' }
  ];


  function stepsHtml() {
    return '<ol class="hg-steps">' + START.map(function (s) {
      return '<li><b>' + rich(s.t) + '</b><span>' + rich(s.d) + '</span></li>';
    }).join('') + '</ol>';
  }
  function cardsHtml(list) {
    return '<div class="hg-cards">' + list.map(function (v) {
      return '<div class="hg-card"><b>' + rich(v.t) + '</b><span>' + rich(v.d) + '</span></div>';
    }).join('') + '</div>';
  }
  function keysHtml() {
    return '<div class="hg-keys">' + KEYS.map(function (g) {
      return '<div class="hg-keygrp"><div class="hg-keyhd">' + esc(g.g) + '</div>' + g.k.map(function (k) {
        var keys = k[0].split(' · ').map(function (part) {
          // a gesture ("Double-click lane", "Drag bar") is plain text; keys are chips
          if (/ /.test(part) && !/^(Mod|Shift|Ctrl)\+/.test(part) && !/^"/.test(part)) return '<span class="hg-kt">' + esc(part.replace(/Mod/g, MOD)) + '</span>';
          return part.split('+').map(function (p) { return '<kbd class="kbd">' + esc(p.replace(/Mod/g, MOD)) + '</kbd>'; }).join('<i>+</i>');
        }).join(' <i>·</i> ');
        return '<div class="hg-key"><span class="hg-kk">' + keys + '</span><span class="hg-kd">' + rich(k[1]) + '</span></div>';
      }).join('') + '</div>';
    }).join('') + '</div>' +
    '<div class="hg-note">' + (mac ? '⌘ is Ctrl on Windows.' : 'Ctrl is ⌘ on a Mac.') + ' Shortcuts are off while typing in a field.</div>';
  }

  var G = {};
  G.TABS = TABS;
  G.START = START; G.VIEWS = VIEWS; G.KEYS = KEYS; G.FEATURES = FEATURES;
  G.mod = MOD;
  G.bodyHtml = function (tab) {
    if (tab === 'views') return cardsHtml(VIEWS);
    if (tab === 'keys') return keysHtml();
    if (tab === 'ideas') return cardsHtml(FEATURES);
    return stepsHtml();
  };
  // the whole start-page area: header row (title + collapse), tab strip, body
  G.html = function (tab, open) {
    tab = TABS.some(function (t) { return t.id === tab; }) ? tab : 'start';
    return '<section class="hg' + (open ? ' open' : '') + '" data-hg>' +
      '<button class="hg-hd" data-hg-toggle aria-expanded="' + (open ? 'true' : 'false') + '">' +
      '<i data-lucide="book-open"></i><span>How to use Headway</span>' +
      '<span class="hg-hd-sub">' + (open ? 'start, views, shortcuts, features' : '') + '</span>' +
      '<i data-lucide="' + (open ? 'chevron-up' : 'chevron-down') + '" class="hg-chev"></i></button>' +
      (open
        ? '<div class="hg-tabs" role="tablist">' + TABS.map(function (t) {
            return '<button role="tab" aria-selected="' + (t.id === tab ? 'true' : 'false') + '" class="' + (t.id === tab ? 'on' : '') + '" data-hg-tab="' + t.id + '"><i data-lucide="' + t.icon + '"></i>' + esc(t.label) + '</button>';
          }).join('') + '</div>' +
          '<div class="hg-body" data-hg-body="' + tab + '">' + G.bodyHtml(tab) + '</div>'
        : '') +
      '</section>';
  };
  return G;
}));
