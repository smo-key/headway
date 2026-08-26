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

## Testing

- Core tests: milestone span/schedule/validation, scope-column defaults and
  legacy inference, built-in rename normalization, story `custom`.
- Smoke tests: updated for the merged panel Fields section, "…" actions
  menu, story panel, scoping story rows, milestone rendering, settings
  markup changes.
