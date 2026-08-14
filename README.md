# Headway

Standalone roadmap planning tool. **Open
`index.html` directly in a browser** — no server, no build, everything is local
(vendored ExcelJS + Lucide, localStorage autosave).

It boots **completely empty** — start from scratch, open a saved `.xlsx`, or
File → Download template for a starter workbook with one worked example.

## Four views

Switch with the tab group in the top center:

- **Setup** — timeline start/end, sprint numbering, workstreams, phases, team
  types, sizing rules, holidays (2026 US calendar preloaded). Replaces the old
  Settings dialog.
- **Planning** — the timeline/gantt. Bars, dependencies, capacity, resources.
- **Budgeting** — planning's timeline board, resources only: editable hourly
  Cost and Rate per role (Type/Workstream chips edit like scoping), Margin %
  and Total (actual hours × rate), and each role's week-hours grid under the
  same phase/date header, tinted by workstream color. Holiday weeks clip to
  the workable hours, shown as a small "(XX)" — that's what totals price.
  Keyboard-friendly (tab across inputs, chips and week cells). The Reports
  drawer lives here.
- **Scoping** — a spreadsheet: the same rows on the left; the right side leads
  with fixed Size / Risk / Workstream / Epic chip columns, then editable text
  columns that grow with their content (Enables / Out of scope / External
  dependencies / Notes by default; Description exists but is hidden). The "+"
  header button adds hidden built-ins or new custom columns; each text column's
  header menu moves, renames (custom), or removes it. Columns — and the frozen
  left pane, per view — resize by dragging edges (remembered in the browser).

The menu bar (**File / Edit / View**) holds everything else; only **Save
.xlsx**, the view switch, and the preflight (validation) chip stay as direct
controls. Menus open on a 0.4s hover and switch instantly while one is open;
rows have right-click context menus; dropdowns share one list UI.

## What it does

| Area | How |
|---|---|
| Timeline | Biweekly-sprint grid (dates primary, sprint numbers secondary — numbering anchor configurable, e.g. S1 = Sep 7). Holidays are individual DATES drawn as day-level hatched segments; click a week header to toggle a whole week, single dates in Settings. Drag empty space to pan, `⌘scroll` to zoom |
| Bars | One uniform duration per item (the work/risk split lives in the panel). Drag to move, edges resize; `⌘-drag` pushes all downstream dependents along. Snap grid in View → Snap (day / week / 2 weeks, default week). View → Auto-order (default on) re-sorts rows by start date after a move |
| Sizes | Measured in weeks: XS 2d · S 1w · M 2w · L 4w · XL 8w (editable in Settings). Risk buffers use the same scale |
| Risk | Per-item severity — None / L / M / H (legacy t-shirt values migrate). Shown in Scoping and the panel only; Planning rows carry no risk chip. The panel also shows a computed dependency-risk estimate with reasons |
| Phases | Header phase lane: spans auto-derive from items, or pin explicit dates (phase modal, or drag the span — body moves, edges resize). Pinned spans show a white outline |
| Budgeting & reports | Budgeting view for costs/rates/margins per role; a collapsible bottom **Reports** drawer (Budgeting only) rolls up effort and estimated cost by workstream, phase, or phase × workstream |
| Capacity switch | Setup → Capacity: off by default. When off, capacity UI, validation and scheduling constraints all stand down |
| Rows | Drag **anywhere on the left pane** to reorder or move between phases; dragging near the top/bottom edge auto-scrolls. Story chevron sits left of the ID; size/risk/hc chips align under their header labels |
| Epics | Epic is a dropdown (with "＋ New epic…"); View → Group by epic groups rows under epic bands inside each phase (drops adopt the target group's epic) |
| Dependencies | Hover a bar → drag its edge **circles** onto another row (left = depends ON it, right = dependency FOR it; Esc cancels mid-draw). Curved arrows show every explicit dep when on; **critical path orange**, violations dashed amber. Click an arrow + Delete removes it. Panel search adds deps **by name** |
| Headcount | `×N` chip on each row (click +1, ⇧-click −1); every item defaults to 1 × Development, work type adjustable per item. Item #s are editable in the panel (invalid/taken numbers pick the next available; deps follow) |
| Team & resources | Roster with member types (reorderable). A **resizable, collapsible Resources panel** at the bottom is a spreadsheet of hours per person per week (default 40): click a cell to type, drag to fill. Weekly capacity = Σ hours ÷ 40 (people-equivalents); the capacity header row shows ≈ parallel work items for a selectable work type (default Development, or All) and everything schedules against it |
| Auto mode | Edit → Auto-schedule: dependency order, earliest start with free capacity, bars stretch across holiday weeks, risk buffers appended, locked items stay put; per-item "Snap earliest" in the panel |
| Validation | Preflight chip + report: cycles, unknown/self deps, starts inside a dependency's risk buffer, missing size, headcount vs roster, weekly over-capacity |
| Excel | **Save .xlsx** writes a styled workbook in the source template's layout at WEEK granularity (one column per week, sprint numbers merged above; solid work + pale risk cells, Next/Future markers) + Stories + Team (incl. off weeks) + a hidden `_RoadmapTool` sheet with lossless state. **Open** loads tool files losslessly and parses template-shaped workbooks — weekly or legacy sprint columns, inferred from the header dates; the pale run at either end of a bar is read as the risk area |
| Safety | Undo/redo (⌘Z / ⇧⌘Z), localStorage autosave (full state + UI prefs; a blocked/full storage now shows "local save unavailable" instead of failing silently), seed restore (File menu). Saved .xlsx files carry the UI prefs too — opening one on another machine restores the exact browser state |

## Files

- `index.html` — open this
- `js/core.js` — pure logic (calendar, deps, capacity, scheduler, risk, critical path); node-testable
- `js/excel.js` — ExcelJS import/export
- `js/app.js` — UI
- `tests/seed.fixture.js` — sample document used by the test suites only
- `js/vendor/exceljs.min.js`, `js/vendor/lucide.min.js` — vendored libraries
- `DESIGN.md` — data model + decisions

## Tests

```bash
# core + excel round-trip (needs exceljs resolvable)
NODE_PATH=<dir-with-node_modules> node tools/roadmapping/tests/core.test.js
# headless UI smoke (needs jsdom + exceljs; skips politely without them)
NODE_PATH=<dir-with-node_modules> node tools/roadmapping/tests/smoke.test.js
```

213 core assertions (calendar, deps/cycles, validation, capacity incl. time
off, auto-scheduler, risk buffers, iterative ripple, scope columns, end date, workstream colors, capacity-safe scheduling, critical
path, full export→import round-trips) + 214 UI smoke assertions (boot, menus,
both views, chips, panel sections, dep search, holiday toggle, resources,
grouping, context menus, column management, blank add rows, export).

## Known limitations

- The exported Roadmap sheet's row-1 phase ribbon allocates overlapping phase
  spans first-come-first-served, so a phase fully inside another's span gets no
  ribbon label (cosmetic only).
- Template import classifies a band row by "text in the ID column, empty Epic";
  a phase named with digits only would be misread — use at least one letter.
- Template import reads the risk area only when a bar uses exactly two fill
  colors; multi-color bars import as all-work.
