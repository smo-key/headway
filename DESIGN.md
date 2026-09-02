# Headway — design

A standalone browser roadmap planning tool. Open
`tools/roadmapping/index.html` directly (file://) — no server, no build step. Seeded from the
"Roadmap" tab of *Agents, Platforms, and Data Sources Inventory* and able to load/save
styled `.xlsx` workbooks in that same shape.

## Source-of-truth model

```
state = {
  meta:   { title, timelineStart (Mon ISO date), numWeeks, endDate,  // endDate (last working day) wins on load
            holidays: [ISO dates],                               // individual days, drawn as day-level segments
            sprintAnchor (Mon ISO), sprintAnchorNum,             // e.g. S1 starts Sep 7
            scopeCols: [ { key, label? } ],                      // scoping columns: built-ins + custom ('c…' keys)
            sizeDays: { XS:2, S:5, M:10, L:20, XL:40 } }        // working days per t-shirt size (week scale)
  phases: [ { id, name, description, bucket, collapsed,          // bucket = backlog shelf (Next/Future)
              startDay, endDay } ]                               // optional pinned window (null = auto from items)
  items:  [ { id, num, phaseId, feature, description, workstream, epic, enables, outOfScope, notes,
              deps: [num…], depsText: [str…], extDeps,          // explicit deps only ("All above" removed)
              custom: { colKey: text },                          // custom scoping-column values
              size,                                              // t-shirt size
              risk,                                              // severity None/L/M/H (metadata; legacy sizes migrate)
              headcount (default 1; decimals ok, 0.5 = half-time), teamType (optional),
              startDay, durDays,      // working-day index / span; null = unscheduled
              riskDays,               // always 0 — risk no longer pads the schedule
              locked, done,
              jiraKey,                // "HW-12" typed in after a Jira CSV import; null = none
              stories: [ { id, title, done, startDay, durDays, jiraKey } ] } ] // story timeline optional (both null = none)
  team:   [ { id, name, type, workstream?, capacity,             // "roles"; capacity = heads at 40 h (0.5 = half)
              rate, cost,                                        // hourly bill rate / hourly cost (budgeting; 0 = unset)
              weekHours: { isoMonday: hours } } ]                // default 40 h/week; 0 = off
  teamTypes:  [ str… ]
  wsColors:  { workstream: paletteKey | 6-hex }                  // COLOR follows the workstream (OS = blue;
                                                                  //  well-known ones get seeded defaults)
  epicIcons: { epicName: lucideIconName }                        // epics carry an ICON, shown on rows
  epicJira:  { epicName: jiraKey }                               // the Jira epic features parent to on export
  epicColors: { … }                                              // legacy, no longer drives display
}
```

Time is measured in **working days** (5 per week) from `timelineStart`; weekends never exist in
the index space. A sprint in the source workbook = 2 weeks = 10 working days. Holidays are individual
dates (Settings adds single days; clicking a capacity cell toggles the week's five): they count zero
progress in scheduling (bars stretch across them) and a fully-holiday week is excluded from
capacity checks. Legacy whole-week blackouts migrate to five dates on load.

**Weeks are weeks.** An item's span is exactly its durDays — the risk t-shirt is planning
metadata (shown on chips and in the panel) and adds no padding; `itemEnd = start + durDays`.
⌘-dragging a bar cascades the end-change through dependents iteratively — forward pushes move
each item only as far as its deps' ends require, backward pulls follow clamped by other
dependencies; locked/done items stop the chain. An auto-order option (default on) stable-sorts
rows by start day after moves/resizes.

## Views

- **Setup** — project configuration page (replaces the old Settings dialog): timeline start/end,
  sprint numbering, workstreams, phases and team types (each add/edit/delete plus drag-to-reorder
  via row grips; workstream order persists in state.wsOrder and drives every dropdown), sizing
  rules, and holidays. No explanatory footer text — controls speak for themselves. The 2026 US holiday calendar is merged once into every
  document (deletions stick). Edit menu → Setup and the resources "manage" button land here.
- **Planning** — the timeline. Sprint header shows dates first (sprint numbers secondary).
  Drag empty lane space to pan; ⌘/ctrl-scroll zooms around the cursor.
- **Budgeting** — planning's own board, resources only: same timeline header (phase lane +
  dates/sprints), frozen left pane, one scroll surface, no cards or rounded borders. One
  row per role: workstream color dot + name, then LEFT-aligned spelled-out columns —
  Type + Workstream chips (shared dropdown, like scoping), editable hourly **Cost** then
  **Rate** inputs, Margin % ((rate − cost) ÷ rate), and **Total = actual hours × rate**.
  The lane is the role's week-hours grid, tinted with the workstream color by load;
  holiday-clipped weeks show the ACTUAL hours in a small "(XX)" sub-line — actual hours
  (planned clipped to 8 h × non-holiday days, `RM.roleWeekHours`) are the basis for every
  total. Fully keyboard-navigable: chips and week cells are tabbable (Enter/Space opens),
  Tab in a cell editor commits and hops to the neighbouring week, and Tab out of a
  rate/cost input is steered across the commit re-render. Capacity is deliberately NOT
  shown here; the left pane resizes with its own remembered width. A Total row closes the
  roster.
- **Scoping** — a spreadsheet: same frozen left pane (chips hidden), right side starts with
  five fixed chip columns — Size, Risk, Weeks (editable), Workstream (color dot), Epic (icon) —
  each filling its whole cell (click anywhere in the cell), then the
  configurable text columns (default Enables / Out of scope / External dependencies / Notes;
  Description hidden by default). Header "+" adds hidden built-ins or custom columns; each
  text column's header menu moves / renames / removes it. Widths resize per browser; rows grow
  with their content (auto-sizing textareas, near-black text).
- Topbar: File / Edit / View menus (hover-switch when one is open), view tabs center; only
  Save .xlsx and the preflight chip remain as direct controls.

## Interactions

- **Timeline-only preview**: an expand button in the phase lane's left cell (Planning only;
  the lane renders even with no phase spans so the button is always reachable) hides the
  topbar, frozen left pane, edit panel and resources panel — just the planning area remains.
  A floating minimize button top-right (or Esc) restores everything; the mode is transient
  (never persisted) and leaves automatically on a view switch.
- **Header phase lane** (Planning only): above the sprint dates, one span per phase — the
  range auto-derives from its items (min start → max end) unless the phase carries pinned
  dates, which win side-by-side (set them in the phase modal, or drag the span: body moves,
  edge handles resize; each pins just that side; a pinned span gets a white outline).
  Overlapping phases stack on extra levels; click a span to edit the phase. Hovering a span shows a styled
  tooltip (name, description, date range) instead of a native title. Hidden when nothing is
  scheduled. Phase band descriptions in the timeline wrap onto multiple lines (bands grow).
- **Bars**: drag body = move (⌘ ripples dependents); left/right handles resize; the weeks chip
  is directly editable. All moves and sizes snap to the View → Snap grid (day / week / 2-weeks,
  default week). Unscheduled rows have no pill: hovering the empty lane previews the landing
  slot (1 week unless sized; low-opacity solid bar, no dashed outline) and a **double-click**
  places the item there (single click just deselects). Double-clicks are counted manually —
  the first click's deselect re-render replaces the lane node, so the native dblclick event
  never fires; the app treats two quick clicks on the same row's lane as the double-click. Size and risk chips open the
  shared dropdown (no more cycling). Bar labels that don't fit the bar spill to its right in ink instead of being clipped; the row weeks readout counts WORKING days (a bar stretched over holidays still reads its clean size), and the weeks editor takes effort weeks, stretching the span over holidays on commit. Dragging a bar VERTICALLY across other rows reorders / re-phases the item (same
  drop indicator and logic as dragging the row's left pane). No validation stripes on bars — checks live on the row's
  alert icon.
- **Stories** (expanded under a row; features with at least one story boot expanded, and
  View → Expand/Collapse all features flips them in bulk) are title-only lines — no checkboxes — that can carry
  their own mini timeline: hovering the empty story lane previews the landing slot and a
  double-click places a 1-week bar (nothing else in the lane — no start tick), draggable and
  edge-resizable like item bars but with no dependency ports. The story title rides on the
  bar as a smaller, more transparent label (spilling right of the bar when it doesn't fit).
  Story timelines ride along whenever their feature moves in time — bar drag, ⌘-ripple,
  keyboard nudge, panel start date, snap-earliest, and auto-schedule all shift them by the
  same delta (`RM.shiftStories`); resizes leave them put. Right-click a story row for
  Remove timeline / Delete story. The add row leads with a lucide plus aligned to the
  story titles. Scoping renders no story rows (its chevron goes invisible but keeps its
  space so columns stay aligned). Story timelines round-trip via the lossless
  `_RoadmapTool` sheet.
- **Rows**: the title is an inline-editable input; the epic chip after the title combines
  icon + label. Planning shows size · weeks · headcount chips — the risk chip lives only in
  Scoping (and the panel); risk values are None/L/M/H. The row context menu also offers Insert feature above/below (the new row
  focuses its inline title input instead of opening the edit panel, and a transient
  `holdPos` flag keeps auto-order from moving it until it gets a start date), Unschedule
  (scheduled items only), Lock/Unlock and Mark/Unmark as done. Drag the grip / number area to reorder or move
  between phases; a blank
  click-to-add row closes every phase (and the resources panel adds people the same way); near-edge
  auto-scroll during drags. Story chevron sits left of the ID. Chips: size · risk · total weeks ·
  headcount, aligned under header labels.
- **Workstreams, epics & grouping**: color follows the workstream (editable via any workstream
  dropdown's pencil or right-clicking a workstream band — rename + palette/custom color; known
  names seed defaults, OS = blue). Epics carry an icon instead (16-icon picker in the epic
  editor), shown beside the row dot. View → Group by workstream and/or Group by epic nest
  phase > workstream > epic ("no workstream" sorts last; drops adopt the target group's
  values). Right-clicking an epic band offers edit/delete; row context menus offer
  move-to-phase and set-epic; phase bands add delete-phase (items move to another phase).
  Edit menu can clear all dependencies (confirmed).
- **Dependencies**: hover a bar → drag its edge circles onto another bar (out-port: target
  depends on it; in-port: it depends on target). Arrows are curved and quiet: only the selected
  item's explicit deps render, plus violations (dashed amber) and the critical path (orange —
  longest scheduled chain incl. risk buffers; View → Critical path highlight toggles the
  orange on bars + arrows, default on). View → Capacity row shows/hides the weekly capacity
  header line (default on). Click an arrow to select it, Delete removes it.
  Dropdowns are keyboard-first: focus lands on the current choice, arrows move, Enter picks,
  and focus returns to the opening chip after the re-render. Panel adds deps by name search; a
  dashed blue preview with matching arrowhead follows the
  cursor while drawing, and every arrowhead matches its line's color. "All above" is not
  supported (dropped on import).
- **Resources**: bottom panel (resizable, collapsible, horizontal scroll synced with the
  timeline; spans exactly the project's weeks); one spreadsheet row per **role** with hours per
  week — click to edit, Tab/Shift+Tab hop cells, drag to fill, grip to reorder. The left pane is
  columnar: name · type chip · workstream chip · capacity (heads at 40 h, click to edit; 0 and
  blank are valid). Right-click a row for rename / type / workstream / capacity / start–end
  dates / remove. The date quick-set zeroes weeks fully outside the window and restores zeroed
  weeks inside it to the 40 h default. Hour cells shade white (0 h) → light blue (20 h) → blue
  (40 h), light enough for the black label; the panel doesn't rubber-band at scroll edges.
  Fill-drags clamp at the last project week, and drag auto-scroll never scrolls the board past
  the timeline's own width — nothing can be entered after the project end date.
  Weekly availability = Σ (hours ÷ 40 × capacity) in people-equivalents (total and per
  type); the header capacity row shows it for a selectable work type. Validation, auto-schedule
  and snap-earliest all use it. Clicking anywhere on the header bar toggles the panel ("manage" still jumps to Setup); every collapse chevron has an enlarged invisible hit area. Items default to 1 × Development.
- **Panel**: resizable via its left edge; # and close on one row, a wrapping auto-growing
  title below; then collapsible card sections (Description, Scope
  incl. custom columns, Details, Size & schedule, People, Dependencies, Stories, Checks) on a
  light background; open/closed state persists per browser. Headcount is a number input (also
  inline-editable on the row chip); clicking empty lane space deselects.
- **Left pane width**: drag the divider anywhere along its full height (a grab strip spans the
  boundary; the header handle still works) — remembered separately for Planning and
  Scoping. Item rows show the epic as a single chip after the title combining icon + label;
  epic bands keep their own leading icon. Settings takes a project END DATE (last working day) instead of a week count.
- **Reports drawer**: a collapsible bottom panel (Budgeting only; hidden elsewhere and
  in the timeline preview) — click the header to open; the grouping dropdown stays
  rendered while collapsed so the header height never changes. A grouping dropdown switches
  By workstream / By phase / By phase × workstream; rows show item count, effort weeks,
  hours and estimated cost (`RM.costReport`: effort hours × avg hourly cost of matching
  roles, blended fallback); workstream mode adds the roster's own booked hours × cost.
- **Capacity switch** (Setup, per-document `meta.capacityEnabled`, OFF by default): when
  off, the capacity header row, the View toggle for it, per-role capacity chips and menu
  entry, and the headcount chips on rows/bars (+ their header label) all disappear,
  headcount/over-capacity validation goes quiet, and auto-schedule /
  snap-earliest place items by dependencies alone.
- **Auto-schedule**: topological order, earliest-start greedy under weekly hours-based
  availability (total and per type; only when the capacity switch is on), blackout
  stretching, `locked`/`done` items pinned. With a
  roster set it NEVER overallocates: an item the roster can't absorb (hours-aware, so part-time
  roles count fractionally) is left unscheduled with a note. Bucket phases never auto-schedule.
  One-item variant: "Snap earliest".
- **Undo/redo**: JSON snapshot stack (session-only — not persisted or exported).
  **Autosave**: localStorage (full state + UI prefs); when a storage write fails (quota, or
  a browser blocking file:// storage) the topbar status flips to "local save unavailable —
  use Save to keep your work" instead of silently claiming success. Seed restore in the File menu.
  File → Download template writes `roadmap-template.xlsx` — a blank roadmap carrying one
  worked example feature (sized, scheduled, one story) so every sheet's shape is visible;
  it never touches the current document.

## Validation (badges on rows + preflight report)

unknown dep number · circular deps · item starts before a dependency's buffered end · scheduled
item whose dependency is unscheduled · missing size on a scheduled item · headcount exceeding
roster (total or type) · weekly over-capacity (respects time off) · bar outside the timeline.
A computed **dependency risk** estimate (none/low/med/high with reasons) shows on each row's
alert hover; the panel keeps only the actionable controls (start date, weeks, size, risk t-shirt).

## Excel round-trip (ExcelJS, vendored)

- **Export**: `Roadmap` sheet in the template's shape but at WEEK granularity — one column
  per week (7-day header dates; row 2 shows sprint numbers merged across each sprint's
  weeks), band rows, solid work cells per bar (legacy pale risk cells still import), gray
  blackout columns per week, Next/Future markers, trailing tool
  columns); `Stories`; `Team` (role/type/workstream + week-hour overrides); `_RoadmapTool` hidden sheet with lossless
  chunked state JSON **plus the UI-preferences snapshot** (zoom, view, grouping, column
  widths, feature expansion, capacity type…) so opening the file on another machine
  restores the exact browser state. `uiSnapshot()`/`applyUi()` in app.js are the single
  definition of that pref set — localStorage and the .xlsx both use them.
- **Import** prefers `_RoadmapTool`; otherwise parses the template — column granularity is
  inferred from the header-date gaps (7 days = weekly, this tool's exports; 14 = legacy
  sprint columns) and does NOT redefine the sprint length (`weeksPerSprint` stays the app
  default). Bars come from solid fills; the pale run at either bar edge (detected by
  lightness) becomes the trailing risk buffer.

## Jira CSV export (`js/export-jira.js`)

Targets Jira Cloud's **user-level** importer (work navigator → ⋯ → Import issues from CSV),
which needs only the Create work items + Make bulk changes permissions — not the admin-only
External System Import. That importer has no Issue Id / Parent Id columns, so one file cannot
create a parent and its children together: `Parent` and `Blocked By` carry the Jira keys typed
into Headway (`item.jiraKey`, `story.jiraKey`, `state.epicJira[epic]`) and stay blank until
then. Columns: Summary, Issue Type, Description (plain text: description, a story checklist
when stories aren't rows, then Enables / Out of scope / External dependencies / Notes),
Parent, Labels (`ws-…`, `phase-…`, `size-…`; story rows add `feature-…`), Priority, Due Date
(deadline), Start/End Date (ISO, from the schedule), Blocked By, Jira Key (lets the wizard map
the row as an update). It is the third format in the Export dialog (right of PowerPoint); choosing it swaps the timeline options for feature and/or story rows and the issue type names
(defaults Story / Sub-task); its settings persist in the UI snapshot (`jiraPrefs`).

## Files

`index.html` · `css/app.css` · `js/core.js` (pure logic, node-testable) · `js/excel.js` ·
`js/app.js` (UI) · `js/vendor/exceljs.min.js` · `js/vendor/lucide.min.js` ·
`tests/core.test.js` · `tests/smoke.test.js` (jsdom) · `tests/seed.fixture.js` (the old
workbook parse, now a TEST FIXTURE only — the app itself boots completely empty; there is
no embedded seed and no seed-restore menu).

## Look

Helvetica Neue, light drafting-paper surface, black phase
bands echoing the workbook, deep blue dominant. Lucide icons throughout. Epic colors draw from a
CVD-validated palette (blue / orange / green / plum) plus free custom hex.
