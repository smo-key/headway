# Milestones, rich fields everywhere, story details, panel & settings polish

Date: 2026-08-26. Follow-up batch to the 2026-08-25 dark-mode/start-page work.
Session runs autonomously; interpretation calls are recorded here and in the
final report.

## 1. Settings copy + shadcn-style controls

- Delete the sentence "Personal settings live on this machine — they never
  travel inside a project file." from the Setup → Appearance card and the
  personal-settings modal. Keep "System follows your OS."
- The app has no build step, so "use shadcn" means adopting shadcn/ui's design
  language as plain CSS components (no React):
  - Checkboxes: `appearance:none` square 16×16, radius 4, border `--line-2`,
    checked = `--blue` fill with an inline white check (SVG data URI);
    label text 13.5px, normal case, `--ink`.
  - Sub-headings inside settings (`.m-sec > label`, e.g. "Theme"): drop the
    uppercase treatment — 13px / 600, `--ink`.
  - Segmented pickers (`.seg`): shadcn "tabs" look — `--well` track, 3px
    padding, radius 8; active segment = `--surface` chip with shadow; normal
    UI font instead of mono.
  - Text/date/number inputs on settings surfaces: 32px height, radius 8,
    shadcn focus ring (`box-shadow: 0 0 0 3px` blue at ~20%).
- Typography hierarchy on settings pages: each Setup tab gets a page title
  (`h1`, ~19px/650) naming the tab; card headings (`h2`) move up to
  15px/650; field sub-headings 13px/600. Same hierarchy in the personal
  settings modal.

## 2. Milestones

- Model: `item.milestone` boolean (normalized in core). A scheduled milestone
  has `startDay` set and `durDays = 0`; `RM.itemSpan` returns 0 for
  milestones, so dependents may start the same day and capacity ignores them.
- Auto-schedule treats milestones as fixed dates (like locked items): never
  moved, dependents schedule after them.
- Validation: milestones skip the NO_SIZE warning.
- Planning view renders a diamond (rotated square, workstream color) centered
  on the start day instead of a bar; label sits to the right; move-drag works,
  no resize handles; dependency ports remain on both sides.
- Scoping/planning "weeks" cells show "—" for milestones; the headcount chip
  is hidden.
- Panel: schedule section shows a single Date field ("Fixed date"). The new
  "…" actions menu and the row context menu gain "Convert to milestone" /
  "Convert to feature" (converting back restores a duration from the item's
  size, else 1 week).
- Excel: visible sheet paints "◆" in the start week (workstream color),
  Start = End = the date. The hidden JSON sheet carries the flag losslessly.

## 3. Scope columns: rich text, removable Description, renaming

- All document scope columns (Description, Enables, Out of scope, External
  dependencies, Notes, custom) hold sanitized rich HTML. Scoping-view cells
  become contenteditable rich editors (the floating B/I/list toolbar extends
  to all of them); the panel edits every column with the WYSIWYG block.
  Legacy plain-text values render with newlines preserved.
- `RM.DEFAULT_SCOPE_COLS = ['description']` — new documents start with just
  Description. Docs saved with an explicit column list keep it. Legacy docs
  with no saved list get Description plus whichever built-ins actually hold
  content (so nothing disappears).
- Any column can be renamed (built-ins gain an optional `label` that wins
  over the canonical name) and removed — Description included. The
  add-column menu still restores hidden built-ins under canonical names.
- Excel: visible-sheet columns export as flattened plain text
  (`RM.htmlToText`); import reconcile compares against the flattened value
  and stores Excel edits as plain text.

## 4. Stories: first-class

- Story model gains `custom: {}` (values keyed by scope-column key, same
  columns as items).
- Stories render on the Scoping tab too: indented rows under the expanded
  feature (chevron already exists), with the add-story row; Planning
  behavior is unchanged.
- Clicking a story (row or bar, any tab; also the pencil in the item panel's
  story list) opens the story in the right panel: parent breadcrumb (click
  returns to the item), title, Done, read-only Workstream/Epic inherited
  from the parent ("rolls up"), rich Description + Acceptance criteria, all
  remaining scope columns as rich fields, timeline (date + weeks + remove)
  when scheduled, and a "…" menu with Delete. The old story modal is
  replaced by this panel; inline title editing in rows gives way to the
  panel.

## 5. Right panel cleanup

- One "Fields" section replaces Description + Scope: every scope column in
  document order, open by default.
- Section containers flatten: no background, border, or card padding —
  hierarchy comes from spacing and heading weight.
- Dropdown buttons (phase/epic/role) size to their content and align left
  instead of stretching.
- Duplicate/Delete leave the panel footer and join a "…" (ellipsis) button
  in the panel header, alongside Convert to milestone.
- Done items show a circled checkmark (lucide `circle-check`) before the
  name in the left pane (planning + scoping) and inside the timeline bar
  before the label.

## 6. macOS native menu

- The app-name menu ("Headway") gains the file actions, exactly:
  New project, Open project, Download template | Save, Save as…, Auto save
  (checkmark) | Export… | Help, followed by the standard Hide/Quit block.
  The native File submenu goes away; Edit/View/Window stay.
- Native icons where possible: `IconMenuItem` with macOS `NativeIcon`
  template images (Add, Folder, MultipleDocuments, Share, Info), falling
  back to a plain item if icon construction fails.
- "Start page" moves to the native View menu on macOS desktop (its old File
  slot no longer exists there). Browser/Windows keep the in-app File menu,
  with labels aligned to the new naming.

## 7. Project name in Setup

- Setup → Timeline gains a leading "Project" card with a Name field editing
  `meta.title` (same commit as the topbar title).

## 8. Interaction polish (added mid-session)

- Timeline drags must never select text: the board's row/lane surfaces get
  `user-select: none`; editable controls (inputs, contenteditable) opt back
  in.
- Dependency arrows lose their arrowheads entirely — plain curves, no SVG
  markers.
- The Save button keeps one fixed width across Save/Saving…/Saved states and
  is disabled while Saving and when already Saved.
- In planning view, clicking anywhere in a row (left pane or lane background)
  selects the item and opens the right panel; controls inside the row keep
  their own behavior.
- Remove the start-page blurb "Every project lives in an .xlsx file on disk —
  edits auto-save to the open file."

## 9. Second batch (added mid-session)

- **Transparent groupings**: workstream/epic band rows lose their fills so
  the sprint grid lines show through; they scroll with the rows instead of
  sticking (a stuck transparent band would collide with rows beneath it).
- **Persistent right panel**: a Planning-view fixture — always present,
  resizable, collapsible via a header button (an edge "peek" handle brings
  it back). With nothing selected it says "No item selected". It no longer
  renders on Scoping/Budget/Setup. Esc deselects; the open/collapsed state
  persists with the other UI prefs.
- **macOS chrome**: brand/title inset grows 86→96px past the traffic
  lights; the lights move down (y 22→26 in tauri.macos.conf.json and the
  Rust re-apply constant).
- **Sizing approaches** (researched: Scrum story points/Fibonacci, t-shirt
  buckets, simple 1–5 scales, Kanban no-estimates): `meta.sizeScheme` +
  ordered `meta.sizeOrder` with `meta.sizeDays` per label. Presets: T-shirt,
  Story points (Fibonacci), Points 1–5, No sizing; editing labels/days/rows
  flips to Custom. Items keep label strings; every option maps to working
  days so scheduling is scheme-agnostic. "No sizing" hides size chips,
  columns, and the NO_SIZE check.
- **Workstreams optional**: `meta.workstreamsEnabled` switch in Setup.
  While on, an item without a workstream renders as a transparent bar with
  a light outline (left-pane dot matches) and its Scoping chip stays empty.
  While off: no Workstream column/field/grouping, items keep the neutral
  default color.
- **"Weeks" → "Duration"** everywhere user-facing.
- **No placeholder dots** in empty Scoping cells/chips.
- **Duration without a schedule**: an unscheduled item accepts an explicit
  duration (panel field + Duration cell); placement uses, in order: preset
  duration → size estimate → 1 week. Empty by default.
- **Holidays as named ranges**: `meta.holidayRanges` [{name, start, end}]
  shown as a table in Setup → Timeline with a name + date-range add row;
  `meta.holidays` stays the derived flat list all calendar math reads.
  Legacy docs migrate by merging consecutive/weekend-bridged dates and
  naming known US observances; the capacity-header week toggle adds/carves
  ranges.

## 10. Third batch (added mid-session)

- **Work week**: `meta.weekHours` (full-time hours, default 40) and
  `meta.daysPerWeek` (1–5, default 5) in Setup → Team. Short weeks are
  modeled as implicit non-working day slots (day index % 5 ≥ daysPerWeek),
  so bar stretching, blackouts, and capacity all respect them; hours/day =
  weekHours ÷ daysPerWeek drives effort-hour and budget math.
- **Dark scrollbars**: explicit themed scrollbar styling (WebKit +
  scrollbar-color) on top of `color-scheme`.
- **Rate card**: `meta.rateCard[role] = {rate, cost}`. "Team types" become
  **Roles**; Setup → Team lists each role with its default hourly rate/cost.
  A person inherits their role's numbers unless overridden on the person
  (0/empty = inherit); Budgeting shows effective values.
- **Capacity rework** (researched: Kanban WIP limits & Little's Law, the
  "limit concurrent big bets to roughly team size ÷ 2" heuristic, flow
  focus over utilization): headcount disappears as an item field. The
  capacity row now shows **available people** per week (fractional FTE from
  the roster's hours ÷ full-time). "Too much work" is judged by
  size-weighted WIP: each active item costs focus by size (M≈1 unit at 2w,
  scaling by working days ÷ 10, clamped 0.3–2). A week is flagged when
  weighted WIP exceeds available FTE. Auto-schedule and snap use the same
  weights.
- **Budgeting**:
  - "+ Add role" click-to-add row at the bottom (same flow as Resources).
  - **Fixed & recurring costs**: `state.costs` [{name, amount, kind:
    fixed|weekly|monthly, startDay, endDay}] — their own section under the
    roles with occurrence markers on the timeline lane, included in the
    grand total and Reports.
  - Zoom works on Budgeting (wheel + buttons).
  - Role/Workstream chips restyle as full-height grid cells (scoping-like),
    keyboard-reachable.
- **Zoom & expand**: smoother wheel zoom (exponential per-delta instead of
  fixed 1.2× jumps); a floating bottom-right cluster on Planning/Budgeting
  with +, − and the Expand button (removed from the topbar, and Expand no
  longer exists on Scoping).
- **Reports tab** (replaces the reports drawer; researched: PMBOK-style
  status areas — schedule/scope/cost/risk, EVM-lite % complete, RAG
  flags, burnup): a dashboard view right of Budgeting with KPI cards
  (% complete by effort, items done, weeks left, effort, planned cost,
  billing & margin), progress-by-phase bars, effort/cost by workstream,
  a cumulative planned-cost curve including roster + fixed/recurring
  costs, upcoming milestones, and top validation flags.
- **Context-menu audit**: no default browser menu anywhere except inside
  editable text (inputs/rich editors keep the native clipboard menu).
  Added: start-page recents (Open/Remove), budget role rows (rename, role,
  workstream, remove), budget cost rows (remove), scoping column headers
  (right-click = column menu), plus the app-chrome theme fallback
  everywhere else.
- **Setup surface**: Epics editable from Setup (Workstreams tab card);
  Setup leaves the view-tab group and becomes a separate icon-only round
  button to the right of Scoping/Planning/Budgeting/Reports.
- **Milestones**: unfilled (outline) diamonds, ~30% narrower, same height.
- **macOS**: traffic lights up 2px (y 26 → 24).

## Testing

- Core tests: milestone span/schedule/validation, scope-column defaults and
  legacy inference, built-in rename normalization, story `custom`.
- Smoke tests: updated for the merged panel Fields section, "…" actions
  menu, story panel, scoping story rows, milestone rendering, settings
  markup changes.

## 11. Batch 4 (2026-08-26, later)

- **Toasts**: shadcn-style — surface background, border, shadow, icon
  (check / alert), bottom-center of the screen; the "Loaded … (full tool
  state)" toast is removed (opening a file needs no announcement).
- **Default workstream**: `null` workstream is a real, renamable workstream
  (`meta.defaultWsName`, default "General") with a customizable color
  (`meta.defaultWsColor`, default blue). Bars/dots/chips use it; the
  transparent "no workstream" outline treatment is gone. Editable at the
  top of Setup → Workstreams.
- **Scoping keyboard/tab**: title cell is full-height and in the same tab
  ring as every chip and rich cell.
- **Formatting toolbar**: the floating B/I/list bar follows its cell when
  the scoping board scrolls.
- **Work week**: `meta.workDays` — Sunday–Saturday checkboxes (up to 5
  working days; the 5-slot index space is kept, trailing slots read as
  off) plus `meta.weekStart` (first day of week, default Monday). Calendar
  mapping goes through per-week weekday offsets; timelineStart snaps to
  the chosen first day. Legacy `daysPerWeek` migrates to the first N
  weekdays from Monday.
- **Holiday quick-edit**: right-click an empty timeline slot on Planning/
  Budgeting → add that date / that week as a holiday, remove the range
  under the cursor, or jump to Setup → Timeline.
- **Export PNG**: picks the destination via the OS save dialog (Tauri) or
  the browser's save-file picker where available, then opens the exported
  file (Tauri opener plugin; browser downloads can't auto-open).
- **Expand fix**: Budgeting's frozen left columns/headers now hide in
  Expand mode like Planning's (they previously overlapped the timeline).
- **⌘-drag**: dragging a bar with ⌘ moves the whole transitive dependent
  chain rigidly by the same delta (push and pull; locked/done items stay).
- **Filter input**: standard input height, no "…", and a ⌘F keyboard-chip
  (kbd component) as an in-input right suffix.
- **Tooltips**: one styled tooltip component app-wide, fed by a delegated
  handler that lifts `title` attributes into a positioned tooltip (native
  browser tooltips no longer appear).
- **Roles**: defaults are real role names (Software Engineer, Product
  Designer, Product Manager, Data Scientist, QA Engineer); roles are
  renamable from Setup (propagates to people, items, and the rate card);
  capacity is fully decoupled from roles (no per-role availability filter,
  demand, or validation — WIP focus units vs total fractional people).
- **Risk column schemes** (researched: RAID/probability-impact risk logs,
  planning-poker estimation confidence, MoSCoW prioritization, dependency
  graph analysis): `meta.riskScheme` — `none` (default for new projects),
  `risk` (manual L/M/H), `auto` (computed dependency risk, read-only),
  `confidence` (H/M/L), `moscow` (M/S/C/W). Column label follows the
  scheme (Risk / Confidence / Priority); configured in Setup → Sizing.

## 12. Batch 5 (2026-08-26, evening)

- **Sprinting view** (new tab after Budgeting): current sprint by default,
  any sprint via selector, or All sprints grouped; kanban board (drag
  cards between status columns) or a fully tab-navigable editable grid
  (features + stories: title, status, assignees, size, duration).
- **Statuses**: configurable per document in Setup → Statuses, separate
  feature and story lists; the last status counts as done and syncs the
  done flag.
- **True 1–7 day work week**: the day-index space now has exactly as many
  slots per week as the checked working days; day indices re-encode
  through calendar dates on any work-week change; legacy 5-slot docs
  migrate once.
- **Sprint length**: Disabled (plain weeks) / 1 / 2 / 4 weeks.
- **Assignees**: features carry roster assignees; people render as
  deterministic-color initial avatars (panel, planning rows, Budgeting,
  Resources, sprint cards); Resources rows read as Name | Role columns.
- **Scoping**: one user-ordered column list across fixed chips and text
  columns (drag headers to reorder); new Start date column; new default
  order leads with Description and Epic. Milestone rows keep editable
  size/risk/duration chips; duration 0 ⇄ milestone, empty = unscheduled;
  milestone dots are diamonds.
- **Polish**: today line above the header (ducks under the frozen pane);
  filter input with search icon, right-aligned ⌘F kbd chip and far-right
  clear button; done text/checkmarks at full opacity; thicker milestone
  diamonds; expand keeps the app header and toggles from the in-place
  zoom cluster, which sits above the Resources panel; phase description
  is rich multiline; '- ' / '1. ' auto-lists in rich editors; panel
  section headers left-aligned with icons; NO_SIZE validation removed.

## 13. Batch 6 (2026-08-26, night)

- **Priority column** (separate from Risk): meta.priorityScheme — none
  (default) / MoSCoW (M/S/C/W) / Critical-High-Medium-Low; configured in
  Setup → Sizing next to the Risk card. Docs saved when MoSCoW lived
  under Risk migrate their values to item.priority.
- **Ladder glyphs**: L/M/H/C values draw as chevron-down / equal /
  chevron-up / arrow-up in chips, panel segs and dropdowns (risk,
  confidence, priority levels); MoSCoW and sizes stay letters.
- **Assignees column**: fixed scoping column before Size (avatar stack,
  click to assign); Planning rows get the same chip after Duration.
- **Stories on Scoping**: story rows share the grid — text columns (incl.
  Description) and a new story Size edit in place; non-applicable fixed
  cells gray out. Stories gain a size field.
- **Right panel on Scoping** (persistent, same as Planning).
- **Setup → Columns**: the full ordered column list — drag to reorder,
  rename/remove text columns, add new ones; hidden built-ins noted.
- **Snap fix**: snapDays is a mode (day/week/2-weeks); week snaps land on
  multiples of the real slots-per-week, holidays ignored.
- **Polish**: scoping header paints opaque (no rows showing through the
  chip-column headers); panel section icons no longer rotate with the
  chevron and Details/Checks got icons.
