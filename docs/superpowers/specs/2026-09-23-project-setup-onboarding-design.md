# New-project onboarding and Setup redesign — design

Date: 2026-09-23
Mockup: https://claude.ai/artifact/FhsShx1Cj1dCkQXGwzgn4J (board 1 onboarding, board 2 Setup,
board 3 map of where today's settings go). Where this doc and the mockup disagree, this doc wins.

## Goal

Settings are hard to find: 12 Setup tabs, feature switches in Apps far from the settings they
control, people editable only in the Resources panel, and a New project dialog that asks for a
name and nothing else. Replace both with **one ordered set of project sections** that is shown
two ways:

- **Onboarding** — a full-screen, numbered, step-by-step wizard when a project is created.
- **Setup** — the same sections, unnumbered, behind a left rail, for editing later.

Each section's body is rendered by the same code in both places, so anything learned in
onboarding is found again in Setup under the same name.

## Non-goals

- No change to the Planning, Scoping, Prioritizing, Sprinting, Budgeting or Reporting views beyond
  what renamed or removed settings force.
- No new scheduling algorithm. Scheduling (today's capacity) keeps the ledger and Auto timeline
  from the 2026-09-16 auto-capacity design.
- No editing of priority or risk ladders. Only size scales are editable, as today.
- No change to how files are created, saved or watched (`createProjectOnDisk`, autosave).

## 1. Information architecture

Project sections, in this order, in both surfaces:

| # | Section | Contents |
|---|---|---|
| 1 | Project | name, start, end, preset (onboarding only), work week, holidays |
| 2 | Sprints | length (Off / 1 / 2 / 3 / 4 weeks), first sprint start, first sprint number |
| 3 | Organization | phases, workstreams, epics, levels & item types (today's Hierarchy card) |
| 4 | Sizing, priority & risk | feature × story grid of schemes; size scales |
| 5 | Budgeting | on/off; roles & rate card |
| 6 | Team | people: name, role, workstream, allocation; capacity type shown from role |
| 7 | Scheduling | on/off; planning level; demand; capacity types and the roles that supply them; explainer |
| 8 | Custom columns | built-in and custom scoping columns, which level each shows on, delete custom |

Setup adds, after the eight:

- **Views** — directly under the project sections, no group heading.
- **Personal · this computer only** — Appearance, Preferences, AI assistant, **Jira Integration**.

Every rail item and every header tab carries a Lucide icon. Setup rail items are not numbered;
onboarding steps are. Setup shows an on/off pill beside Sprints (`2 wk` / `Off`), Budgeting and
Scheduling.

Map from today (`SETUP_SECTIONS` in `js/app.js`):

| Today | New home |
|---|---|
| Timeline › name, start, end | Project |
| Timeline › Holidays; Team › Work week | Project |
| Timeline › Sprints | Sprints |
| Phases; Workstreams (+ Epics); Hierarchy card | Organization |
| Sizing (size, priority, risk stacked) | Sizing, priority & risk |
| Team › Roles & rate card; Apps › Budgeting | Budgeting |
| Resources panel (people) | Team (the panel keeps week-by-week hours) |
| Capacity | Scheduling (renamed) |
| Columns | Custom columns |
| Apps | Views (Sprinting and Budgeting switches removed from it) |
| Jira | Personal › Jira Integration |
| Appearance, Preferences, AI assistant | Personal |

"Capacity planning" is renamed **Scheduling** in all user-facing text (Setup, menus, hints,
tooltips such as "Setup → Capacity" in the capacity row, AI navigate targets). Internal names
(`capacityEnabled`, `capTypes`, `RM.capSupply`) stay.

## 2. Shared section rendering

Split today's `renderSetup` tab bodies into one function per section:

```
SECTIONS = [{ key, label, icon, sub, body(ctx) , summary(state) }]
ctx = { mode: 'wizard' | 'setup', state, draft? }
```

- `body(ctx)` returns the section's cards. `mode` only toggles the few documented differences
  (Project's preset card and work-week summary line).
- `summary(state)` returns the one-line text for the onboarding Review step.
- The existing delegated handlers on `#setupView` (`click`, `change`, drag-to-reorder) move to a
  shared host element so the same data attributes work in the wizard.
- `renderSetup` becomes: rail from `SECTIONS` + Views + Personal, then `body({ mode: 'setup' })`.
- `setupTab` keys are renamed to the section keys. A saved UI snapshot or `HeadwayApp.openSetup`
  call with an old key (`timeline`, `phases`, `workstreams`, `capacity`, `columns`, `sizing`,
  `apps`, `jira`, …) maps to its new section; unknown keys fall back to `project`.

In **Setup** every change commits immediately, as today. In the **wizard** the sections edit a
draft state built from `blankState()`; nothing touches disk until Create (section 4).

## 3. Welcome screen (first project only)

Shown before step 1 when `userName()` is empty and no project has ever been created on this
machine (new local flag `headway-onboarded-v1`, set on first Create).

- **Your name** → `setUserName` (`headway-user-v1`). Used by Version history, so it replaces the
  "Who's editing?" prompt for new users; `maybeAskName` still covers people who skip it.
- **Theme** — Light / Dark / Match system cards with a small swatch → `THEME_KEY`, applied live.
- Continue goes to Project. Back from Project returns here. Neither field is required.

## 4. Onboarding wizard

Replaces `newProjectModal`. Opened from File › New project, the start page's New project button,
and the native menu.

**Chrome.** Full window over the app. The real `#topbar` stays (traffic-light inset on macOS,
Windows caption buttons, drag region), with its content swapped for the logo mark, "New project",
"Step N of 9" and a close button. A 3px progress bar sits under it. Left: numbered step rail
(completed steps show a check). Right: the section body. Footer: Back, Skip (steps 2–8), Continue.
There is no "skip to review" and no footer note.

**Steps.** Welcome (first time only) → 1 Project … 8 Custom columns → 9 Review & create.

**Required preset.** Project's preset card must be chosen before Continue enables or any later
rail step can be opened. Choosing a preset writes its values into the draft (section 5); every
value remains editable on later steps. Changing preset again overwrites those values.

**Review & create.** One row per section: number, name, `summary()`, Edit (jumps to that step).
Views are not shown; a new project gets Scoping, Prioritizing and Reporting on. **Create project**
then runs today's flow: flush the open document if dirty, `createProjectOnDisk(draft)` (desktop
save dialog / browser download), `adoptProject`, set `headway-onboarded-v1`.

Closing the wizard discards the draft (confirm only if anything past step 1 was changed).

## 5. Presets

Shown as four selectable cards on Project, each with a name, one-sentence description and a small
static timeline sketch (sprint or month header plus feature/story bars; no capacity row, no tags).
Unselected: white, thin border. Selected: blue border, light blue fill, check badge top-right.

| Preset | Feature size | Story size | Feature prio | Story prio | Feature risk | Story risk | Sprints | Budget | Scheduling |
|---|---|---|---|---|---|---|---|---|---|
| Scrum | none | fibonacci | levels | moscow | none | none | 2 wk | on | on · story · points |
| Advanced Scrum | tshirt | fibonacci | levels | moscow | risk | risk | 2 wk | on | on · story · points |
| Rapid Delivery | tshirt | none | none | none | none | none | off | off | on · feature · person |
| Minimal | tshirt | none | none | none | none | none | off | off | off |

Keys: size `meta.sizeScheme` / `storySizeScheme`, priority `priorityScheme` /
`storyPriorityScheme`, risk `riskScheme` / new `storyRiskScheme`, sprints `weeksPerSprint`
(0 = off), budget `apps.budget`, scheduling `capacityEnabled` / `planLevel` / `capMode`.
Values the user did not specify and this doc chose: Scrum's priority schemes, Advanced Scrum's
budget and scheduling (copied from Scrum) and its hand-set feature sizes (not `rollup`), Rapid
Delivery's priority/risk off. Revisit if the user objects.

## 6. Sections in detail

### 1 Project
Name, Starts, Ends (two date fields; no weeks/working-days bubble). Onboarding: preset card, then
a one-line "Mon–Fri · 40 h per week · N US holidays in range" row with **Edit**, which expands the
Setup controls in place. Setup: full Work week and Holidays cards (today's controls, moved).

### 2 Sprints
Segmented length: Off, 1, 2, 3, 4 weeks. **3 weeks is new** (today: Disabled/1/2/4). When on:
first sprint start (date) and "…and is sprint number" stepper, plus a strip of the next six sprints
with dates. Off keeps today's `weeksPerSprint = 0` behaviour (week labels, no sprint numbers) and
**additionally hides the Sprinting tab** (section 7).

### 3 Organization
Three columns: Phases (with backlog buckets), Workstreams (color), Epics (icon), each with its
drag-to-reorder list and an Add box that accepts a pasted list (one per line). Below: a collapsed
"Levels & item types" row with **Customize**, which opens today's Hierarchy card.

### 4 Sizing, priority & risk
A grid: rows Size / Priority / Risk, columns Feature level / Story level; each cell picks a scheme
(RICE is not offered for stories, as today). Below, **Feature scales** and **Story scales** cards:
- Size: editable options (label, working days), remove, drag to reorder, Add option; editing marks
  the scheme Custom (today's behaviour) with a Reset to the scheme.
- Priority and risk: the scheme's levels as read-only chips; computed schemes (RICE, Risk (auto),
  Roll up from stories) and None show a one-line explanation.

**Story risk gets its own scheme.** Stories already carry `story.risk`, rated on the single
`meta.riskScheme`. Add `meta.storyRiskScheme` (`none` | `risk` | `confidence`; never `auto`) so
features and stories pick independently; `story.risk` is validated against it (`RM.setRiskScheme`
takes a `kind` like the size helpers). Migration: an old file gets `storyRiskScheme =
riskScheme` unless that is `auto`, in which case `none`, so nothing visible changes on open.
Excel already round-trips `story.risk`; the hidden JSON sheet carries the new meta field.

### 5 Budgeting
Switch "Track budget" = `apps.budget` (drives the Budgeting tab). On: Roles & rate card table
(role, bill rate/h, cost/h, margin), Add role. Off: a note that roles are named on the Team step.
Roles stay one list (`state.teamTypes`) whether budgeting is on or off.

### 6 Team
Table of people. Every field on a person (`normalizeState` member shape) is editable here except
capacity type:

| Column | Field | Notes |
|---|---|---|
| Name | `name` | |
| Title | `role` | free-text job title (optional) |
| Role | `type` | pick from `teamTypes`, or type a new role → added to `teamTypes` |
| Capacity type | — | read-only, "from role" (`m.capType`, kept in sync by `RM.syncMemberCapTypes`); shown only when Scheduling is on |
| Workstreams | `workstreams` | multi-pick; first one mirrors `workstream` as today |
| Allocation | `capacity` | percent in the UI, 100% = 1.0; 0 allowed |
| Points / sprint | `points` | shown only in points mode; blank = `meta.defaultPoints` |
| Rate, Cost | `rate`, `cost` | shown only when budgeting is on; blank = the role's rate card (shown as placeholder) |

Add person; delete person (×). No paste-from-spreadsheet. Note under the table: allocation is each
person's default; week-by-week hours (`weekHours`) stay in the Resources panel, which keeps its own
editing. Both edit `state.team`, and the Budgeting view's roster keeps working unchanged.

### 7 Scheduling
Switch "Schedule against capacity" = `capacityEnabled`. On: Planning level (Features / Stories),
Demand (Per person / Story points, with default points per person per sprint), Capacity types.

**Role → capacity type.** Roles and capacity types stay separate lists, but each role supplies at
most one capacity type and a person's type comes from their role:

- New `state.roleCapTypes = { roleName: capTypeName }`.
- Each Capacity types row shows "Supplied by" role chips and **+ Role**.
- `member.capType` stays as a stored copy so every existing reader keeps working, but it is
  no longer edited per person: `RM.syncMemberCapTypes(state)` sets it to
  `roleCapTypes[m.type] || ''` for every person with a role, and runs in `normalizeState` and
  after any change to roles, `roleCapTypes` or a person's role. A person with **no role** keeps
  the type they already had (read-only, hint "Set a role to change").
- The per-person capacity-type pickers (Budgeting roster chip, Resources panel menu) become
  read-only.
- Renaming/deleting a role or type updates `roleCapTypes` (extend `RM.renameCapType`,
  `RM.renameRole` and the matching delete paths).
- Migration on load: for each role, the capacity type most of its people carried wins; ties take
  the first type in `capTypes`. If any person's old type disagreed with their role's, show one
  `toast` after load naming how many people changed, so the user can check roles.

**All capacity types are tracked.** Remove `meta.capRowTypes` and the Tracked checkboxes; every
type gets a capacity row and constrains Auto timeline and Place. `normalizeState` deletes the
field.

Bottom card **How scheduling works**: three blurbs (Supply, Demand, Over capacity) and two small
weekly charts ("As drawn" with red overflow, "After Auto timeline"). Shown whether scheduling is on
or off. Static illustration, not live data.

### 8 Custom columns
Table: grip, column name, "Shows on" Feature / Story / Both (`scope`), kind pill (Built-in /
Custom), delete (×). Delete works on custom columns; on built-ins it is shown disabled with the
tooltip "Built-in columns can't be deleted". Add column.

### Views (Setup only)
Scoping, Prioritizing, Reporting switches (`meta.apps`). Planning shows "Always on"; Sprinting
shows "Follows Sprints"; Budgeting shows "Follows Budgeting". The Apps tab is removed.

### Personal · this computer only
Appearance, Preferences, AI assistant unchanged. **Jira Integration** moves here unchanged. Its
token is already per-machine; the project mapping still lives in `meta.jira` (per document), so the
page keeps a small "Saved in this project" label on the mapping block.

## 7. Apps model change

`RM.appEnabled(state, 'sprints')` becomes `RM.sprintsEnabled(state.meta)`; `apps.sprints` is
ignored and dropped by `normalizeState`. `apps.budget` stays but is edited only from Budgeting.
Everything else about hidden views (`view` fallback, AI `navigate` refusal) is unchanged.

## 8. Testing

- `tests/`: unit tests for preset application (each preset → exact meta values), `roleCapTypes`
  migration (majority, tie, disagreement count), `syncMemberCapTypes`, story risk scheme normalisation and
  Excel round-trip, `capRowTypes` removal, `apps.sprints` → `sprintsEnabled`, old `setupTab` key
  mapping.
- Headless Chrome + jsdom run (see the verification workflow memory): open the wizard, walk every
  step, pick each preset, Create in browser mode; open Setup and visit every rail item.
- Manual desktop check: the wizard's top bar drags the window on macOS and Windows.

## 9. Follow-ups (not in this change)

- Letting a person override their role's capacity type.
- A live preview of the project's own timeline in onboarding.
