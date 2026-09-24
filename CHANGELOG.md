# Changelog

Release notes for Headway. Each version gets one `## <version>` section, written
during the release; the GitHub release body and the in-app "What's new" dialog
both come from the matching section here. Newest first.

Format: short sections, bullets, one bold title per notable change —
`**Feature title**: a short, impactful description and use case.`

## Unreleased

- **Each item type lives at one level**: Setup → Organization → Hierarchy lists every type once, under its level, with a Level picker to move it and an Add type per level; the first type at a level is the default. The Allow any type at any level switch is gone; older files keep each type at the first level it was allowed at.
- New projects use Epic › Feature › Story (Story and Task). Scrum presets add Bug at the story level; feature-organized presets (Rapid Delivery, Contract Negotiation, Innovation, Minimal) read Epic › Feature › Task, with Issue beside Feature.
- **Contract Negotiation and Innovation presets**: Rapid Delivery with MoSCoW priority and risk ratings on features, or with RICE priority and no sizes.
- New project setup: a progress bar sits under the top bar, the project name starts empty and is required, Continue stays grayed out until a step's required fields are filled, and switching preset after changing its settings asks first.
- The 2027 US holidays are on the calendar, including in files saved earlier.
- Setup → Organization shows Phases, Workstreams and Epics side by side; each Add box takes a pasted list (one per line), epics can be added before any item uses them, and Levels & item types fold away behind Customize. An edited size scale offers Reset to its scheme, each capacity type lists the roles that supply it with + Role, and the Jira mapping notes it is saved in the project.
- **New project setup**: New project opens a full-screen setup that walks the same sections as Setup — pick a preset first (Scrum, Advanced Scrum, Rapid Delivery or Minimal, each with a small timeline sketch), adjust anything, then review and create.
- The first project on a computer starts with a short Welcome that asks for your name and theme.
- Setup → Scheduling: pick the capacity type each role supplies, and a new How scheduling works card shows supply against demand. Setup → Custom columns: every column shows a delete button (disabled for built-ins).
- Setup → Budgeting has its own Track budget switch next to the rate card, and Setup → Team lists everyone in one editable table (name, title, role, workstreams, allocation, points, rate and cost); a person's capacity type now follows their role.
- Setup → Sizing, priority & risk: one grid picks the size, priority and risk scheme for features and for stories; priority and risk levels show as read-only chips, size scales stay editable.
- Setup is regrouped into Project, Sprints, Organization, Sizing, Budgeting, Team, Scheduling and Custom columns, then Views and your Personal settings (this computer only); Capacity is now Scheduling, and Jira is Jira Integration.
- Stories rate risk on their own scheme (none, Risk or Confidence), separate from features; older files keep what they had.
- Sprints can be 3 weeks long, and the Sprinting tab now simply follows whether sprints are on.
- Scheduling: each role now supplies one capacity type and its people take it; older files pick each role's most common type on open and say how many people moved.
- **Holiday weeks no longer block per-person capacity**: a unit now asks heads × multiplier × the share of the week's working days it covers (a full week asks one head, a week with a holiday 0.8, a two-day tail 0.4) — the same day weighting story points use. Before, it asked a whole head in every week it touched while a holiday cut the supply to 0.8, so a solo person could never take work in a week with a holiday and Auto timeline / Place at earliest slot jumped past such weeks. The capacity row reads the same way.
- **⌘+ / ⌘− zoom the Planning timeline** (Ctrl on Windows / Linux, numpad + / − too) instead of zooming the whole page; they do nothing while you type in a field or a dialog is open, and outside Planning. View → Zoom in / Zoom out show the keys; ⌘scroll over the timeline still zooms.
- **Exclude from Auto timeline**: right-click a feature or story (Planning, Scoping, Prioritizing, Sprinting, the panel or a multi-selection) → **Exclude from Auto timeline**, or tick *Excluded from Auto timeline* in the panel / story editor. The ⚡ Auto timeline leaves it where it sits (booked in capacity like a locked item) and never counts it, so a phase whose only out-of-place item is excluded reads as in place; a feature's flag covers its stories, and at the Features planning level a feature with any excluded story stays put as a whole (its stories ride with it); Auto no longer re-sizes an excluded feature. Unlike Lock it keeps the bar draggable and editable, and **Place at earliest slot** still moves it. Lock and Exclude are mutually exclusive — setting one clears the other. Rows and bars show a zap-off mark; the flag travels in the file, in the Excel sheets (an *Excluded from Auto* column) and through the AI assistant.
- Phase bands keep their + feature, edit and ⚡ Auto timeline buttons at the right edge, always visible.
- Sprinting: each sprint's total sits at the right edge of its side entry and heading as a larger number; with story-points capacity on (sprints on, numeric story sizes) it reads planned / available points (e.g. `18 / 20`), in red when the sprint is over — never red while a filter hides some of its stories.
- Sprinting: clicking a sprint in the side list lands on the sprint heading; it used to end up hidden under the sticky filter bar.
- Resources shows either the × seat multiplier (per-person demand) or points per sprint (story-points demand), never both; in story-points mode a person's points are scaled by their hours only. A seat of 0 still means "supplies nothing" in both modes. The row menu's Capacity… opens whichever of the two the row shows.
- The capacity type chip in the Planning left pane shows the whole type name once its column is wide enough (it was cut to eight characters regardless).
- Planning rows no longer show the epic tag beside the feature title; add the Epic column when you want it.
- The assistant's close button no longer sits under the right-panel toggle.
- **Sizes from stories on Auto timeline**: at the Stories planning level, clicking a phase's Auto timeline button also sizes each of its features that has at least one sized story — the nearest label on the feature scale for the working days its stories' timelines span (parallel stories do not add up), rounded up to the feature snap (day / week / sprint). The size is set once, on the click; afterwards it is an ordinary size you can edit.
- Stories snap to the story snap when created, auto-placed or placed at the earliest slot.
- **Planning columns**: the left pane's chip columns have headers, resize by dragging the header edge, reorder by dragging the label, and show/hide from the header's right-click menu or the + at its end — now including Workstream, Epic, Start and Deadline.
- Phase bands are light in light mode, and rows step down phase → workstream → epic → feature → story to white, across both the left pane and the timeline.
- Row titles no longer show an outline on hover; the box appears only while renaming.
- **Setup → Capacity**: capacity planning, the planning level, the demand model and capacity types now live on their own Setup tab (Team keeps roles and the work week).
- **Demand models**: Per person (a unit in flight uses one person of its type × its multiplier) or Story points (points spread over the unit's working days against each person's points per sprint, default 10).
- **Capacity row**: one header row, "Capacity (people)" or "Capacity (points)", reads total demand / supply across every capacity type — per week per person, and per sprint in story-points mode (one cell spanning each sprint; two-week blocks with sprints off). It turns red when any single type is over, even if the total fits, and its tooltip lists each type's numbers. The hand-typed weekly limit and the count/points basis are gone — the limit is the roster. Every capacity type counts: Auto timeline / Place at earliest slot constrain every type the roster supplies. The column legend now sits below the capacity row.
- **Story points are a per-sprint budget**: a unit's points follow its working days (a story starting mid-week puts most of its points where most of its days are), validation flags over-capacity once per sprint, and Auto timeline / Place at earliest slot fit work against the sprint's points. Work that can never fit a sprint is left where it is with a note instead of being pushed to the far end of the timeline.
- **Untyped people supply nothing**: in the Resources panel a person without a capacity type shows a "set type" prompt instead of the × seat / points chip; click it to pick a type (their seat and points come back).
- Features always carry a capacity type now (blank ones default to the first type, Development). A capacity-enabled document whose roster supplies no type will show a "nobody supplies" warning on open — give the people a capacity type in the Resources panel.
- **Auto timeline**: a ⚡ button on every phase band (also in the band's right-click menu and the phase dialog) lays that phase out in one go — its items follow their dependencies at the earliest start the roster's capacity allows. Locked and done items never move: locked items are booked in capacity as fixed, and done items anchor their dependents' starts. It is a one-shot action, not a mode: nothing re-runs on open or after other edits, and one undo takes the whole layout back. The button is disabled when everything in the phase is already in place, and when capacity planning is off.
- **Place at earliest slot**: right-click any feature or story → Place at earliest slot moves just that one to the first slot its dependencies and the roster's capacity allow (any phase, capacity planning on).
- Auto timeline and Place at earliest slot never move work before its phase begins (the phase's pinned start, or where its earliest item already sits); work already under way keeps its start unless a dependency or capacity pushes it (a push past today places it as new work, on the snap grid).
- Story rows in the Planning left pane show a capacity type chip and, in per-person mode, a × multiplier (feature rows too at the Features planning level); both are columns you can hide or reorder. The Resources panel gains a points column in story-points mode.
- Epic and workstream group rows now stay pinned under their phase band while you scroll, like the phase band itself.
- Auto-order rows by start now also applies when a document opens.
- Opening a document whose rows are out of start order (with auto-order on) re-sorts them there and then, says so ("Rows auto-ordered") and leaves the document unsaved.
- The Auto-schedule dialog is gone; Auto timeline and Place at earliest slot replace it.
- Flagged rows are tinted light orange in the left pane, darker when selected.
- The Planning / Scoping left pane no longer shows #numbers on rows (they stay in the panel header, on cards and on sprint rows).
- #numbers are gone from Sprinting rows and Prioritizing cards too (the panel header keeps them), and Sprinting rows drop the sprint-number bubble.
- Scoping: story glyphs and warning icons center on the first title line like the feature rows; story rows lose their left divider line and the epic tag beside the feature title (the Epic column already shows it).
- Workstream color marks are filled circles everywhere (rows, chips, dropdowns, the PNG legend).
- **Roll up from stories**: a new feature sizing option (Setup → Sizing, first in the list; T-shirt sizes stay the default). Each feature's size is the sum of its story points and its working days are the stories' days added up; feature sizes are not edited by hand in this mode. Switching back to a hand-picked scale clears the derived sizes.
- **Move a story to another feature**: right-click a story (Planning / Scoping rows, Sprinting rows, Prioritizing cards, the panel) → Move to feature… opens a searchable list of features.
- Assignee pickers are searchable and show each person's avatar; typing narrows the roster, Enter picks the first match.
- Prioritizing cards: text fields (Description etc.) are read-only on the card — edit them in the panel or the Scoping grid; the title still renames on double-click. Clicking the selected card keeps it selected.
- Story titles in the Planning left pane read a step darker (still lighter than features).
- **Standalone HTML export**: Export → Standalone HTML (view-only) saves the whole roadmap as one self-contained .html — every tab is there to browse, nothing edits, and Setup keeps only the theme. Opens from disk in any browser.
- **Planning level** (Setup → Capacity): Features (default) plans capacity on the features and stories need no details; Stories ignores feature weights and durations and plans the work on the stories and their capacity types.
- **Capacity types**: stories carry a capacity type (what they drain and who can take them) and people carry the type they supply — Development, Design, QA, … by default, editable and reorderable under Setup → Capacity. Story assignee pickers list only the people supplying the story's type. The Budgeting / Resources rows gain a Capacity column.
- Default roles now lead with Project Manager and Product Manager.

## 1.0.13 — 2026-09-09

### New

- **Story numbers**: every story has a # from the same pool as features; it shows on every row and card and edits in the story panel.
- **Story dependencies**: a story can depend on any other story. Add one in the story panel's Dependencies section (by # or name) or drag a story bar's edge circle onto another story; arrows and order warnings work like features; links push to Jira and the Stories sheet carries them.
- **Flags**: right-click any feature or story → **Flag…** (optional reason) to mark it for attention; an orange flag takes the alert slot on the row (hover for the reason) and shows on Prioritizing cards, Sprinting rows and the panel header. **Edit flag…** / **Unflag** from the same menu. The AI assistant can set `flag` too.
- **Titles rename on double-click**: on Planning and Prioritizing the feature and story titles are text, like Sprinting. Double-click one, or pick **Rename…** from its context menu, to edit; a single click just selects, and dragging from the title moves the row. Scoping keeps its spreadsheet cells.

### Improved

- AI assistant: Markdown tables in replies render as tables.
- ⌘B / ⌘I (Ctrl+B / Ctrl+I) bold and italicise inside every rich text field.
- Prioritizing: the board is full width; **Columns → Unset column** hides or shows the catch-all column; clicking a card opens the detail panel on it and clicking the same card again puts the panel away; an empty column or swimlane cell says "No items"; the catch-all swimlanes (the default workstream, "No epic") sit last.
- Sprinting: the main content area is capped at a readable width.
- Right-clicking the app background no longer opens the theme menu (the theme lives in Settings and the app menu).

## 1.0.12 — 2026-09-09

### New

- **Item types and hierarchy**: every epic, feature and story carries a type (Feature, Bug, Task, Story, Subtask, Epic by default). Add your own types and colors in Setup → Hierarchy, rename the three levels, choose which types each level accepts, or allow any type anywhere. Pick a type from the panel, the context menu or Edit epic; rows show an icon for non-default types.
- **Type glyphs and color by type**: the colored square left of every title is now the item's type icon (Feature keeps the filled square, Story is a bookmark) in the active color. **View → Color by item type** colors bars by type.
- **Tags**: features and stories carry free-form labels, edited in the panel's Tags section (Enter or a comma adds one; the box suggests tags already in the document). Filter by tag with `#tag`. Tags survive the Excel round trip, push to Jira as labels, and the AI assistant can read and set them.
- **Links in rich text**: any URL typed into a description or scope field renders as a link; ⌘-click (Ctrl-click on Windows/Linux) opens it in the browser. The stored text stays plain.
- **Insert story above / below**: right-click a story on the row, the card, the sprint row or the panel to insert a blank story next to it, ready to edit. Features had this already.
- **Duplicate story**: every story context menu can copy a story in place.
- **Zero and half-point stories**: the story Fibonacci scale starts at 0 and 0.5. A 0-point story is real work with no effort (a placeholder, a spike already done, tracking-only) — it counts 0 toward sprint totals but still takes a day on the timeline.
- **Jira issue types per Headway type**: sync and CSV export use each type's Jira issue type. The three fixed type fields in Setup → Jira become a per-type table; existing documents keep their previous names.

### Improved

- **Sprinting rows belong to one sprint**: a row sits only in the sprint it starts in, with an ⓘ glyph saying how many sprints it carries over. Rows wear sprint number, epic, workstream and assignee chips and drop the duration, date-range and phase columns; empty sprints at either end are hidden.
- **Sprinting context menu**: right-click a feature or story for **Move to sprint…** and **Assign…**. Titles are plain text; double-click or Rename to edit.
- **Add only where empty**: the Add feature / Add story buttons and rows show only in an empty phase, sprint section, Prioritizing column or story list. Everywhere else, right-click and use Insert above/below, so full lists stay compact.
- **Prioritizing board**: the Fields menu can hide any card chip (Size, Priority, Risk, Duration, Epic, Workstream) and always offers Priority and Risk even when the columns are by that field; the toolbar filters by phase like Sprinting; milestones no longer appear on the board.
- **AI assistant edits only what changed**: a write applies just the items the assistant added, edited, deleted or reordered, so the screen no longer flashes a reload and edits you are making are never overwritten. The assistant can also read and set item types.
- **AI model picker**: the drawer loads the gateway's models when opened and the effort levels a model supports when you pick one, switching to Medium when the current level is not offered.
- **Right panel order**: Estimate, then People, then Tags.
- **Milestones carry no size or priority**: converting an item to a milestone clears both, and the panel, timeline and Scoping table show no size or priority controls on milestone rows.
- **Desktop reload only with Auto save on**: the open file reloads after external changes only while Auto save is on, so unsaved edits are never replaced.
- Settings → Jira: issue-type icons match the rest of the UI.

### Fixed

- The story panel's Size, Priority and Risk buttons did nothing.
- Sprint totals show story counts until a story is sized, instead of a misleading 0.
- Reporting's Delivery-by-sprint counts a feature in every sprint it spans.
- Focusing and leaving a rich-text field that contains a URL no longer records a spurious edit.
- Story-view sprints whose features have no visible stories no longer render empty.
- Size day repair handles null values; story-only sizes fall back correctly in the size order.
- Excel export: the band fill no longer covers the Tags column.

## 1.0.11 — 2026-09-08

### New

- **Automatic updates**: the desktop app checks for a new release on launch and every hour, downloads it in the background, and offers an Update button in the header and on the start page — one click installs and relaunches. A **Check for updates** button on the start page checks on demand.
- **What's new dialog**: release notes open once after each update so you see what changed; click the version number on the start page to read them again.
- **Apps switch**: Setup → Apps turns each header tab on or off per project, so a plan that never budgets or scopes shows only the views it uses. Planning always stays and hidden data is kept.
- **Prioritizing by field**: a Columns dropdown lays feature cards out by Priority, Size or Risk instead of phase, plus an Unset column — drag a card between columns to set the value.
- **Story board**: the Prioritizing page's Story level shows one card per story, with its feature named above the title, in columns of Priority, Size or Risk. Drag to set the field; Group, Sort and the filters all work per story.
- **Sprinting filters**: narrow the sprint page by text, phase, epic and workstream; sidebar and section counts follow.
- **Acceptance criteria column**: a built-in column right after Description that shows on stories only by default, so criteria live next to the story instead of in a custom field.
- **Priority colors**: Must/Critical reads bright red, Should/High orange, Could/Medium green and Won't/Low gray — in chips, the panel's picker and dropdown dots, and on bars when coloring by priority.

### Improved

- **Disk reload keeps your place**: when the open file changes on disk, only the data refreshes — your view, selection, scroll position, open dialogs and focused field all survive.
- **AI effort follows the model**: the drawer offers only the effort levels the gateway says the model supports and hides the control for models without any; long gateway model ids show as short names with the full id in a tooltip.
- **Hollow milestone markers**: diamond, star and circle markers are now outlines that stay legible in any color, and milestone titles read bold in the left pane.
- **Unscheduled last**: the Sprinting sidebar lists every sprint first and Unscheduled at the end.

### Fixed

- **Corporate TLS inspection**: the desktop app now trusts the OS certificate store, so AI, Jira and update checks work behind Netskope or Zscaler proxies.

### Removed

- **Timeline-only preview**: the expand button in the phase lane is gone; collapse the left pane and the panel instead.

## 1.0.10 — 2026-09-07

### New

- **AI assistant**: ask questions about your roadmap and let the assistant draft, resize, and reschedule work for you from a chat panel — it reads the open plan and proposes edits you approve.
- **Jira Cloud sync**: link a plan to a Jira project so features and stories push to Jira issues and pull status back, keeping the roadmap and the tracker in step.
- **Per-kind sizing and priority schemes**: features and stories can carry their own size scales and priority ladders, so estimates read naturally at both levels.

### Improved

- **Role removal**: delete a team role from Setup and every assignment cleans up with it.
- **Inline holiday editing**: add and remove holiday dates straight from the Setup table.

## 1.0.9 — 2026-09-02

### New

- **Sprinting page**: sprint-by-sprint rows with drag to move and reorder work across sprints.
- **Group-level Add feature**: add a feature straight from a group row.

### Improved

- **Sticky panel header** and a **collapsible left pane** keep context while scrolling long plans.
- **Column scoping by kind**: Scoping columns can apply to features only or stories only.
- **Milestone marker styles** and filter chips that wear the epic icon or workstream dot.
