# Headway

Standalone roadmap planning tool. Runs two ways:

- **Desktop app** (Windows & macOS, via Tauri) — real Open/Save to disk,
  including synced folders like OneDrive, with auto-reload when the open file
  changes externally only while Auto save is on. Install below.
- **Web page** — open `index.html` directly in a browser: no server, no build,
  everything is local (vendored ExcelJS + Lucide, localStorage autosave).

## Install the desktop app

No admin rights needed — both installs are per-user.

**macOS** (installs to `~/Applications`):

```sh
curl -fsSL https://raw.githubusercontent.com/smo-key/headway/main/install.sh | sh
```

**Windows** (installs under `%LOCALAPPDATA%`; run in PowerShell or cmd):

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -Command "irm https://raw.githubusercontent.com/smo-key/headway/main/install.ps1 | iex"
```

Both scripts fetch the latest [GitHub release](https://github.com/smo-key/headway/releases)
(built by `.github/workflows/release.yml` on every `v*` tag). Re-run the same
one-liner to update, or let the app do it: it checks for a new release on
launch and every hour, downloads it in the background, and offers an Update
button (header and start page); the start page's **Check for updates** button
checks on demand. Release notes live in `CHANGELOG.md` — one `## <version>`
section per release, written as part of the release task — and become both the
GitHub release body and the in-app "What's new" dialog, shown once after each
update (click the version number on the start page to reopen it).

In the desktop app, File → Open… / Save / Save As… use native dialogs and write
straight to disk. On macOS the File / Edit / View menus live in the system menu
bar (with the standard clipboard, window and quit items); the in-window menu
buttons appear only on Windows and in the browser. The app header is the
titlebar: on macOS the native bar is hidden with overlay traffic lights, on
Windows the frame is custom with Windows-11-style caption buttons — empty
header space drags the window, double-click zooms/maximizes. Save any workbook into a OneDrive (or Dropbox, iCloud, …)
folder and it syncs like any other file; if the file changes on disk — another
machine syncs an edit, or Excel saves over it — Headway reloads it
automatically and shows a toast, only while Auto save is on.

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

Setup → **Apps** (right after Timeline) switches each of the header tabs on or off per project — hide Scoping or Budgeting for a plan that doesn't use them; Planning always stays. The data behind a hidden tab is kept.

The menu bar (**File / Edit / View**) holds everything else; only **Save
.xlsx**, the view switch, and the preflight (validation) chip stay as direct
controls. Menus open on a 0.4s hover and switch instantly while one is open;
rows have right-click context menus; dropdowns share one list UI.

## What it does

| Area | How |
|---|---|
| Timeline | Biweekly-sprint grid (dates primary, sprint numbers secondary — numbering anchor configurable, e.g. S1 = Sep 7). Holidays are individual DATES drawn as day-level hatched segments; click a week header to toggle a whole week, single dates in Settings. Drag empty space to pan, `⌘scroll` or `⌘+` / `⌘−` to zoom |
| Bars | One uniform duration per item (the work/risk split lives in the panel). Drag to move, edges resize; `⌘-drag` pushes all downstream dependents along. Snap grids in View → Snap, one for features (default week) and one for stories (default sprint); day / week / sprint each. View → Auto-order (default on) re-sorts rows by start date after a move |
| Sizes | Measured in weeks: XS 2d · S 1w · M 2w · L 4w · XL 8w (editable in Settings). Risk buffers use the same scale |
| Risk | Per-item severity — None / L / M / H (legacy t-shirt values migrate). Shown in Scoping and the panel only; Planning rows carry no risk chip. The panel also shows a computed dependency-risk estimate with reasons |
| Phases | Header phase lane: spans auto-derive from items, or pin explicit dates (phase modal, or drag the span — body moves, edges resize). Pinned spans show a white outline |
| Budgeting & reports | Budgeting view for costs/rates/margins per role; a collapsible bottom **Reports** drawer (Budgeting only) rolls up effort and estimated cost by workstream, phase, or phase × workstream |
| Capacity switch | Setup → Capacity: off by default. When off, capacity UI, validation and scheduling constraints all stand down |
| Rows | Drag **anywhere on the left pane** to reorder or move between phases; dragging near the top/bottom edge auto-scrolls. Story rows drag by their grip in Planning and Scoping alike — up/down inside a feature, or onto another feature. Story chevron sits left of the ID; size/risk/hc chips align under their header labels |
| Prioritizing | Kanban with feature cards: columns are the phases, or pick **Priority**, **Size** or **Risk** from the Columns dropdown to lay features out by that field (plus Unset) and drag between them to set it. The **Story** level turns it into a story board: one card per story with its feature named above the title, in columns of one story field — **Priority** (default), **Size** or **Risk**, picked from the Columns dropdown — plus an Unset column. Drag a story to a column to set that field; the **Fields** menu always offers the Priority and Risk chips (even when the columns are by that field); the toolbar filters by text, phase, epic and workstream. Milestones never appear on the Prioritizing board.; click a card to open the detail panel on it (click it again to close); **Columns → Unset column** hides the catch-all column |
| Sprinting | Sprint-by-sprint page: the sidebar lists every sprint (today's marked) with Unscheduled last, the main area is one scrolling list of sections with a row per feature — or per story, grouped under its feature (Features / Stories toggle). Drag a row to another sprint (a section or its sidebar entry) to start it on that sprint's first day with its span and stories riding along; drop it before another row to reorder — the order is the shared items order, so it changes in every view. Every row belongs to the sprint it starts in — an ⓘ glyph says how many sprints it carries over — and wears its sprint number, epic and workstream chips, size, priority, risk and assignees; story-view headings show the feature's priority, size and duration. Empty sprints at either end of the timeline are hidden. Right-click a row for **Move to sprint…**, phase, epic, workstream, assignees; story rows get their own timeline the same way. The toolbar filters by text, phase, epic and workstream; an empty section ends with **Add feature**, a filled one adds through the row context menu |
| Detail panel | Right-hand panel for the selected feature or story. Its header row (number, milestone chip, collapse) stays pinned while the body scrolls; **Fields** follow the Scoping column order and show only the columns scoped to that kind of row; **Tags** holds free-form labels (Enter or a comma adds one, Backspace on an empty box drops the last, the box suggests tags already used in the document); **Integrations** (Jira key) sits last. Stories get a **People** section for assignees. The header number is editable on features **and** stories — type a new one and it moves, or falls back to the next free number when that one is taken. Enter in any single-line field (the title included) leaves the field and saves it; clicking out always saves. The B / I / list bar of a rich field appears only while it has focus; URLs become links (⌘/Ctrl-click opens them) |
| Panes | Collapse the left pane from the button in its top-left corner and the right panel from its collapse button; each folds away completely, leaving only a reopen button in the same corner. Hotkeys `[` and `]` toggle them, never while typing |
| Columns | Setup → Columns (or a column header's right-click menu) picks whether each text column shows on **features, stories, or both**: excluded rows grey the cell out in Scoping and drop the field from the panel. **Acceptance criteria** is a built-in column that sits after Description and shows on stories only by default |
| Milestones | Zero-duration items pinned to a date. Each carries a marker style — **diamond** (default), **star** or **circle** — picked from the row's context menu or the panel's milestone chip; it shows on the timeline, row dots and in PNG, PowerPoint and Excel exports |
| Types & hierarchy | Every epic, feature and story has a **type** (Feature, Bug, Task, Story, Subtask, Epic by default — add your own). Setup → **Hierarchy** names the three levels, picks which types each level accepts (the first is the default for new items), and holds each type's **Jira issue type**. "Allow any type at any level" lifts the per-level restriction (off by default; a type outside its level only raises a validation warning). Pick a type from the panel chip, the row's context menu, or Edit epic; every row and card shows its type glyph (Feature = filled square, Story = bookmark, Bug, Task…) in the active color; each type has a color for **View → Color by item type** |
| Epics | Epic is a dropdown (with "＋ New epic…"); View → Group by epic groups rows under epic bands inside each phase (drops adopt the target group's epic); Insert feature above/below (row context menu) inherits the anchor's epic and workstream; the **Add feature** row appears only in a phase with no features yet |
| Dependencies | Hover a bar → drag its edge **circles** onto another row (left = depends ON it, right = dependency FOR it; Esc cancels mid-draw). Curved arrows show every explicit dep when on; **critical path orange**, violations dashed amber. Click an arrow + Delete removes it. Panel search adds deps **by name**. Stories depend on stories the same way — the story panel's **Dependencies** section, ports on story bars, and arrows between story bars (feature ↔ story links are not a thing) |
| Flags | Right-click a feature or story → **Flag…** (optional reason) puts an orange flag in the row's alert slot (hover for the reason); it also shows on cards, sprint rows and the panel header. **Edit flag…** / **Unflag** from the same menu |
| Headcount | `×N` chip on each row (click +1, ⇧-click −1); every item defaults to 1 × Development, work type adjustable per item. Item #s are editable in the panel (invalid/taken numbers pick the next available; deps follow) |
| Team & resources | Roster with member types (reorderable). A **resizable, collapsible Resources panel** at the bottom is a spreadsheet of hours per person per week (default 40): click a cell to type, drag to fill. Weekly capacity = Σ hours ÷ 40 (people-equivalents); the single capacity header row shows total demand / supply across every capacity type (red when any one type is over; the tooltip breaks it down per type; people without a capacity type supply nothing) — per week in people, per sprint in story points (points are a per-sprint budget spread over each unit's working days; two-week blocks with sprints off) — and everything schedules against it the same way |
| Auto timeline | The **⚡ button** on a phase band (or its right-click menu, or the phase dialog) lays that phase out once: features by dependency order at the earliest start with free capacity, bars stretch across holiday weeks, locked and done items never move (locked ones are booked in capacity as fixed, done ones anchor their dependents), and neither do features or stories you **Exclude from Auto timeline** from their right-click menu or the panel (booked like locked ones, but still draggable, and Place at earliest slot still moves them; Lock and Exclude clear each other), and work already under way keeps its start unless a dependency or capacity pushes it. At the Stories level the click also sizes each feature from the span its sized stories cover (an ordinary, editable size afterwards). One undo reverts it; the button is disabled when the phase is already in place or capacity planning is off. For a single item, **Place at earliest slot** / **Snap earliest** from the row's context menu or the panel |
| Validation | Preflight chip + report: cycles, unknown/self deps, starts inside a dependency's risk buffer, missing size, headcount vs roster, per-type over-capacity (per week, or per sprint in story points), capacity types nobody on the roster supplies |
| Excel | **Save .xlsx** writes a styled workbook in the source template's layout at WEEK granularity (one column per week, sprint numbers merged above; solid work + pale risk cells, Next/Future markers) + Stories + Team (incl. off weeks) + a hidden `_RoadmapTool` sheet with lossless state. **Open** loads tool files losslessly and parses template-shaped workbooks — weekly or legacy sprint columns, inferred from the header dates; the pale run at either end of a bar is read as the risk area |
| Jira | **Export → Jira CSV** writes a file for Jira Cloud's user-level CSV importer (work navigator → ⋯ → Import issues from CSV; needs only Create work items + Make bulk changes). Rows are features and/or stories with each item's type mapped to a Jira issue type in Setup → Hierarchy; Parent and Blocked By carry the Jira keys typed into Headway (panel "Jira key" on features and stories, Edit epic… for epics), so the first import creates issues, you paste the keys back, and later exports parent stories and re-map as updates. Dates are ISO — pick `yyyy-MM-dd` in the wizard |
| AI assistant | The **AI** toolbar button (⌘J) opens a chat drawer that answers questions about Headway, the open plan and project-management practice, and edits the document, your preferences or (when Setup → Jira is connected) Jira itself on request — every document edit is undoable and shows in Version history as “you · AI”. Set it up in Setup → Personal → AI assistant: a **LiteLLM gateway** (URL, API key, model picked from the gateway, optional extra headers) or, in the desktop app, your **Claude subscription** (runs `claude -p` from Claude Code, no API key). The drawer loads the gateway's models when it opens and the effort levels a model supports when you pick one, falling back to Medium. Pick the model and effort level in the drawer's compose bar; replies show their thinking; attach images, PDFs or text files with the paperclip, drag-drop or paste. Settings stay on this machine |
| Safety | Undo/redo (⌘Z / ⇧⌘Z), localStorage autosave (full state + UI prefs; a blocked/full storage now shows "local save unavailable" instead of failing silently), seed restore (File menu). Saved .xlsx files carry the UI prefs too — opening one on another machine restores the exact browser state |

## Files

- `index.html` — open this
- `js/core.js` — pure logic (calendar, deps, capacity, scheduler, risk, critical path); node-testable
- `js/excel.js` — ExcelJS import/export
- `js/app.js` — UI
- `js/export-jira.js` — Jira CSV export (user-level importer shape)
- `js/jira.js` — Jira Cloud sync
- `js/desktop.js` — Tauri desktop bridge (native dialogs, disk save/load, file watching); no-op in a browser
- `tests/seed.fixture.js` — sample document used by the test suites only
- `js/vendor/exceljs.min.js`, `js/vendor/lucide.min.js` — vendored libraries
- `src-tauri/` — Tauri shell (Rust); `scripts/copy-frontend.mjs` stages the static files into `dist/` for bundling
- `install.sh` / `install.ps1` — per-user desktop installers (fetch the latest GitHub release)
- `DESIGN.md` — data model + decisions

## Desktop development

```bash
make setup           # npm install — @tauri-apps/cli (plus jsdom/exceljs for the tests)
make dev             # run the desktop app with live frontend
make build           # build the platform installer locally (needs Rust)
make test            # core + headless UI smoke suites
```

Releases: push a tag like `v1.0.1` (matching `version` in
`src-tauri/tauri.conf.json`) and the GitHub Actions workflow builds the macOS
universal .app/.dmg and the Windows per-user NSIS installer and publishes them
to a GitHub release.

## Tests

```bash
# core + excel round-trip (needs exceljs resolvable)
NODE_PATH=./node_modules node tests/core.test.js
# headless UI smoke (needs jsdom + exceljs; skips politely without them)
NODE_PATH=./node_modules node tests/smoke.test.js
```

760 core assertions (calendar, deps/cycles, validation, capacity incl. time
off, autoTimeline/placeUnit, risk buffers, iterative ripple, scope columns, end date, workstream colors, capacity-safe scheduling, critical
path, full export→import round-trips) + 1150 UI smoke assertions (boot, menus,
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
