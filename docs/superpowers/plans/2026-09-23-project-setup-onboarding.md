# Project Setup and Onboarding — Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Replace the 12-tab Setup and the one-field New project dialog with one ordered set of eight project sections, shown as an unnumbered Setup rail and as a full-screen numbered onboarding wizard with a required preset.

**Architecture:** Phase 1 changes the data model in `js/core.js` (role → capacity type, a story risk scheme, all capacity types tracked, 3-week sprints, Sprinting following sprints) and regroups `renderSetup` in `js/app.js` into section bodies keyed by the new sections. Phase 2 adds `RM.PRESETS` / `RM.applyPreset`, then a wizard that swaps the global `state` for a draft while open, so the very same Setup section bodies and delegated handlers edit the draft; Create hands the draft to the existing `createProjectOnDisk`.

**Tech Stack:** Vanilla ES5 browser JS (no build step); `js/core.js` is also a CommonJS module. Tests: `NODE_PATH=./node_modules node tests/core.test.js` and `NODE_PATH=./node_modules node tests/smoke.test.js` (jsdom). Lucide icons (`<i data-lucide="…">`).

**Spec:** `docs/superpowers/specs/2026-09-23-project-setup-onboarding-design.md` (mockup: https://claude.ai/artifact/FhsShx1Cj1dCkQXGwzgn4J)

## Global Constraints

- ES5 only in `js/*.js`: `var`, `function`, no arrow functions, template literals, `const`/`let`.
- Section order and labels, verbatim: `Project`, `Sprints`, `Organization`, `Sizing, priority & risk`, `Budgeting`, `Team`, `Scheduling`, `Custom columns`; then `Views`; then group `Personal · this computer only` with `Appearance`, `Preferences`, `AI assistant`, `Jira Integration`.
- Section keys (used in `data-sutab`, `setupTab`, the UI snapshot): `project`, `sprints`, `org`, `est`, `budget`, `team`, `scheduling`, `columns`, `views`, `appearance`, `prefs`, `ai`, `jira`.
- User-facing "Capacity" / "Capacity planning" becomes "Scheduling". Internal names (`capacityEnabled`, `capTypes`, `RM.capSupply`, …) do not change.
- Only size scales are editable; priority and risk levels are read-only chips.
- `weeksPerSprint` allowed values: `0` (off), `1`, `2`, `3`, `4`.
- `commit(label, fn)` labels are short lowercase; toasts are sentence case.
- Every task ends with both suites passing and a commit whose message ends with `Co-Authored-By: Claude Opus 5.5 (1M context) <noreply@anthropic.com>`.
- Other sessions edit this tree live: never `git stash`, never `git checkout` another branch in this tree; stage only the files the task names.
- One line per user-visible change under `## Unreleased` in `CHANGELOG.md`, in the task that ships it.

## Review Focus

- A person with no role but an old capacity type keeps that type after open and after any role edit (not blanked).
- The desktop file watcher reloading the open document while the wizard is open must update the stashed document, never the draft, and must not close the wizard.
- Closing the wizard (×, Escape) after edits restores the previous document exactly, including undo/redo and the saved flag; nothing is autosaved or written to version history from the draft.
- An old file with `meta.capRowTypes` as a list, `apps.sprints: false`, or `setupTab: 'capacity'` in its UI snapshot opens without errors and lands on the right section.
- Changing preset after editing later steps overwrites only the preset-owned fields (schemes, sprints, budget, scheduling), not name, dates, phases, people or columns.

---

## File map

| File | Responsibility in this plan |
|---|---|
| `js/core.js` | `roleCapTypes`, `RM.syncMemberCapTypes`, `storyRiskScheme`, `RM.setRiskScheme(state, key, kind)`, tracked types = all, `weeksPerSprint` 3, `RM.appEnabled('sprints')`, `RM.PRESETS`, `RM.applyPreset` |
| `tests/core.test.js` | new sections: `role capacity types`, `story risk scheme`, `sprints app`, `presets` |
| `js/app.js` | Setup rail + section bodies, Views, Team table, Scheduling card, column delete, read-only per-person capacity type, wizard (draft swap, chrome, welcome, presets, review, create) |
| `index.html` | `#wizard` host inside `body` after `#topbar` |
| `css/app.css` | `.su-pill`, `.su-grid`, `.su-chip`, `.su-team`, `.wz-*`, `.pcard`, `.pc-sketch` |
| `tests/smoke.test.js` | replace Setup-tab assertions (lines ~228–300, ~532–660) with the new keys; add wizard walk |
| `CHANGELOG.md`, `DESIGN.md` | Unreleased entries; the Views bullet and Setup bullet in DESIGN.md |

---

# Phase 1 — Data model and Setup redesign

### Task 1: Role → capacity type, and every type tracked

**Files:**
- Modify: `js/core.js` — `RM.trackedCapTypes` (~388), `RM.renameCapType` (~411), `RM.removeCapType` (~429), `RM.renameRole` (~3510), `RM.removeRole` (~3528), `normalizeState` meta block (~1407) and team/capTypes block (~2036–2064)
- Test: `tests/core.test.js`

**Interfaces:**
- Produces: `state.roleCapTypes` (`{ roleName: capTypeName }`), `RM.syncMemberCapTypes(state) → number` (people whose type changed), `RM.lastCapTypeChanges` (count from the latest `normalizeState`), `RM.setRoleCapType(state, role, capType)`, `RM.capTypeRoles(state, capType) → [role]`. `RM.trackedCapTypes(state, sup)` keeps its signature and returns every type. `meta.capRowTypes` no longer exists.

- [ ] **Step 1: Write the failing tests** (append before the `regressions` section)

```js
// ------------------------------------------------------------- role capacity types
section('role capacity types');
var sR = RM.normalizeState({
  meta: { timelineStart: '2026-07-27', numWeeks: 8, capacityEnabled: true, capRowTypes: ['Design'] },
  phases: [{ id: 'p1' }], items: [],
  capTypes: ['Development', 'Design'],
  teamTypes: ['Engineer', 'Designer', 'PM'],
  team: [
    { name: 'A', type: 'Engineer', capType: 'Development' },
    { name: 'B', type: 'Engineer', capType: 'Development' },
    { name: 'C', type: 'Engineer', capType: 'Design' },
    { name: 'D', type: 'Designer', capType: 'Design' },
    { name: 'E', type: '', capType: 'Design' },
    { name: 'F', type: 'PM', capType: '' }
  ]
});
eq(sR.roleCapTypes, { Engineer: 'Development', Designer: 'Design' }, 'majority type per role; roles with no typed people map to nothing');
eq(sR.team.map(function (m) { return m.capType; }), ['Development', 'Development', 'Development', 'Design', 'Design', ''],
  'people take their role\'s type; a person with no role keeps theirs');
ok(!('capRowTypes' in sR.meta), 'capRowTypes is dropped');
eq(RM.trackedCapTypes(sR), ['Development', 'Design'], 'every type is tracked');
var sTie = RM.normalizeState({ meta: { timelineStart: '2026-07-27', numWeeks: 8 }, phases: [{ id: 'p1' }], items: [],
  capTypes: ['Development', 'Design'], teamTypes: ['Eng'],
  team: [{ name: 'A', type: 'Eng', capType: 'Design' }, { name: 'B', type: 'Eng', capType: 'Development' }] });
eq(sTie.roleCapTypes, { Eng: 'Development' }, 'a tie takes the first type in capTypes');
var sKeep = RM.normalizeState(RM.clone(sR));
eq(sKeep.roleCapTypes, sR.roleCapTypes, 'an existing map is kept on reload, not re-derived');
RM.setRoleCapType(sR, 'PM', 'Design');
eq(sR.team[5].capType, 'Design', 'setting a role\'s type updates its people');
eq(RM.capTypeRoles(sR, 'Design'), ['Designer', 'PM'], 'roles supplying a type, in teamTypes order');
RM.renameCapType(sR, 'Design', 'UX');
eq([sR.roleCapTypes.Designer, sR.team[3].capType], ['UX', 'UX'], 'renaming a type follows into the map');
RM.renameRole(sR, 'Designer', 'Product designer');
eq(sR.roleCapTypes['Product designer'], 'UX', 'renaming a role moves its mapping');
RM.removeCapType(sR, 'UX');
ok(!('Product designer' in sR.roleCapTypes) && sR.team[3].capType === '', 'removing a type clears its roles and people');
RM.removeRole(sR, 'Engineer');
ok(!('Engineer' in sR.roleCapTypes), 'removing a role drops its mapping');
var sNoRole = RM.normalizeState({ meta: { timelineStart: '2026-07-27', numWeeks: 8 }, phases: [{ id: 'p1' }], items: [],
  capTypes: ['Design'], teamTypes: ['X'], team: [{ name: 'E', type: '', capType: 'Design' }] });
RM.setRoleCapType(sNoRole, 'X', 'Design'); RM.renameRole(sNoRole, 'X', 'Y');
eq(sNoRole.team[0].capType, 'Design', 'no-role person keeps an old type through role edits');
eq(RM.lastCapTypeChanges, 0, 'nothing to report for a clean file');
RM.normalizeState({ meta: { timelineStart: '2026-07-27', numWeeks: 8 }, phases: [{ id: 'p1' }], items: [],
  capTypes: ['A', 'B'], teamTypes: ['R'], team: [{ name: 'x', type: 'R', capType: 'A' }, { name: 'y', type: 'R', capType: 'A' }, { name: 'z', type: 'R', capType: 'B' }] });
eq(RM.lastCapTypeChanges, 1, 'normalize reports how many people changed type');
```

- [ ] **Step 2: Run to verify they fail**

Run: `NODE_PATH=./node_modules node tests/core.test.js`
Expected: FAIL — `roleCapTypes` undefined, `RM.setRoleCapType is not a function`.

- [ ] **Step 3: Implement**

In `normalizeState`, delete the `m.capRowTypes = …` lines (~1407–1408) and replace them with `delete m.capRowTypes;`. Delete the `if (Array.isArray(m.capRowTypes)) { … }` block (~2060–2064). After the block that collects `state.capTypes` from team/items (~2049), add:

```js
    // Each role supplies at most one capacity type; people inherit it.
    // Old files have no map: each role takes the type most of its people
    // carried (ties: first in capTypes). An existing map is kept.
    var rct = state.roleCapTypes && typeof state.roleCapTypes === 'object' ? state.roleCapTypes : null;
    if (!rct) {
      rct = {};
      state.teamTypes.forEach(function (role) {
        var counts = {};
        state.team.forEach(function (mbr) {
          if (mbr.type === role && mbr.capType) counts[mbr.capType] = (counts[mbr.capType] || 0) + 1;
        });
        var best = '', bestN = 0;
        state.capTypes.forEach(function (t) { if ((counts[t] || 0) > bestN) { best = t; bestN = counts[t]; } });
        if (best) rct[role] = best;
      });
    }
    state.roleCapTypes = {};
    Object.keys(rct).forEach(function (role) {
      if (state.teamTypes.indexOf(role) !== -1 && state.capTypes.indexOf(rct[role]) !== -1) state.roleCapTypes[role] = rct[role];
    });
    RM.lastCapTypeChanges = RM.syncMemberCapTypes(state);
```

Next to `RM.memberPoints`, add:

```js
  // people take their role's capacity type; a person with no role keeps
  // whatever type they already had. Returns how many people changed.
  RM.syncMemberCapTypes = function (state) {
    var map = state.roleCapTypes || {}, n = 0;
    state.team.forEach(function (m) {
      if (!m.type) return;
      var t = map[m.type] || '';
      if (m.capType !== t) { m.capType = t; n++; }
    });
    return n;
  };
  RM.setRoleCapType = function (state, role, capType) {
    state.roleCapTypes = state.roleCapTypes || {};
    if (capType && state.capTypes.indexOf(capType) !== -1) state.roleCapTypes[role] = capType;
    else delete state.roleCapTypes[role];
    RM.syncMemberCapTypes(state);
  };
  RM.capTypeRoles = function (state, capType) {
    var map = state.roleCapTypes || {};
    return state.teamTypes.filter(function (r) { return map[r] === capType; });
  };
```

Replace the body of `RM.trackedCapTypes` with `return RM.capTypesOf(state).slice();` (keep the `sup` parameter unused). In `RM.renameCapType`, after `state.capTypes[i] = newName;` add:

```js
    Object.keys(state.roleCapTypes || {}).forEach(function (r) {
      if (state.roleCapTypes[r] === oldName) state.roleCapTypes[r] = newName;
    });
```

and delete its `capRowTypes` block. In `RM.removeCapType` add the same loop with `delete state.roleCapTypes[r]`, and delete its `capRowTypes` block. In `RM.renameRole`, before `return true;`:

```js
    if (state.roleCapTypes && oldName in state.roleCapTypes) {
      state.roleCapTypes[newName] = state.roleCapTypes[oldName];
      delete state.roleCapTypes[oldName];
    }
```

In `RM.removeRole`, before its return: `if (state.roleCapTypes) delete state.roleCapTypes[name];`.

In `js/app.js`, `loadWorkbookBuffer`/`adoptState` path: right after `RM.normalizeState` runs on an opened file (not on reloads, not on local restore), read `var n = RM.lastCapTypeChanges;` and if `n > 0` call `toast(n + (n === 1 ? ' person now takes' : ' people now take') + ' the capacity type of their role \u2014 check Setup \u203a Scheduling')`. In `js/app.js` remove every write of `capRowTypes` (~10638–10641) — Task 6 replaces that UI.

- [ ] **Step 4: Run both suites**

Run: `NODE_PATH=./node_modules node tests/core.test.js && NODE_PATH=./node_modules node tests/smoke.test.js`
Expected: core PASS. Smoke: fix any assertion that reads `meta.capRowTypes` or the Tracked checkboxes by deleting it (Task 6 adds the replacements).

- [ ] **Step 5: Commit**

```bash
git add js/core.js js/app.js tests/core.test.js tests/smoke.test.js
git commit -m "feat(scheduling): roles supply capacity types; every type is tracked

Co-Authored-By: Claude Opus 5.5 (1M context) <noreply@anthropic.com>"
```

### Task 2: Story risk scheme, 3-week sprints, Sprinting follows sprints

**Files:**
- Modify: `js/core.js` — `RM.riskSchemeOf`/`riskEnabled`/`riskOrderOf`/`riskColLabel`/`setRiskScheme` (~333–354), story normalize (~1705, ~1818), `weeksPerSprint` normalize (~1397), `RM.appEnabled` (~292), apps normalize (~1412)
- Modify: `js/app.js` — story risk callers (~2077, ~4874) pass `'story'`
- Test: `tests/core.test.js`

**Interfaces:**
- Produces: `meta.storyRiskScheme`; `RM.riskSchemeOf(state, kind)`, `RM.riskEnabled(state, kind)`, `RM.riskOrderOf(state, kind)`, `RM.riskColLabel(state, kind)`, `RM.setRiskScheme(state, key, kind)` — `kind` `'feature'` (default) | `'story'`. `RM.STORY_RISK_SCHEMES = ['none', 'risk', 'confidence']`. `RM.appEnabled(state, 'sprints') === RM.sprintsEnabled(state.meta)`.

- [ ] **Step 1: Write the failing tests**

```js
// ------------------------------------------------------------- story risk scheme
section('story risk scheme');
function riskSt(meta, stories) {
  return RM.normalizeState({ meta: Object.assign({ timelineStart: '2026-07-27', numWeeks: 8 }, meta),
    phases: [{ id: 'p1' }], items: [{ num: 1, feature: 'f', risk: 'H', stories: stories }] });
}
var sOld = riskSt({ riskScheme: 'risk' }, [{ title: 's', risk: 'M' }]);
eq([sOld.meta.storyRiskScheme, sOld.items[0].stories[0].risk], ['risk', 'M'], 'old file: stories keep the feature scheme');
eq(riskSt({ riskScheme: 'auto' }, [{ title: 's' }]).meta.storyRiskScheme, 'none', 'auto never applies to stories');
var sConf = riskSt({ riskScheme: 'risk', storyRiskScheme: 'confidence' }, [{ title: 's', risk: 'H' }]);
eq(RM.riskOrderOf(sConf, 'story'), ['H', 'M', 'L'], 'story ladder follows the story scheme');
eq(RM.riskOrderOf(sConf), ['L', 'M', 'H'], 'feature ladder unchanged');
eq(riskSt({ storyRiskScheme: 'auto' }, []).meta.storyRiskScheme, 'none', 'auto is rejected for stories');
RM.setRiskScheme(sConf, 'none', 'story');
eq([sConf.items[0].stories[0].risk, sConf.items[0].risk], [null, 'H'], 'turning story risk off clears stories only');
RM.setRiskScheme(sConf, 'confidence');
eq(sConf.items[0].risk, 'H', 'feature scheme change keeps a value that exists on the new ladder');

section('sprints app');
var sp = RM.normalizeState({ meta: { timelineStart: '2026-07-27', numWeeks: 8, weeksPerSprint: 3, apps: { sprints: false } },
  phases: [{ id: 'p1' }], items: [] });
eq(sp.meta.weeksPerSprint, 3, '3-week sprints are allowed');
ok(RM.appEnabled(sp, 'sprints') && !('sprints' in sp.meta.apps), 'Sprinting follows sprints on, apps.sprints is dropped');
sp.meta.weeksPerSprint = 0;
ok(!RM.appEnabled(sp, 'sprints'), 'sprints off hides Sprinting');
```

- [ ] **Step 2: Run to verify they fail**

Run: `NODE_PATH=./node_modules node tests/core.test.js`
Expected: FAIL on `storyRiskScheme`, weeksPerSprint 3 → 2, appEnabled.

- [ ] **Step 3: Implement**

```js
  RM.STORY_RISK_SCHEMES = ['none', 'risk', 'confidence'];
  function riskKey(kind) { return kind === 'story' ? 'storyRiskScheme' : 'riskScheme'; }
  RM.riskSchemeOf = function (state, kind) {
    var s = state && state.meta && state.meta[riskKey(kind)];
    if (kind === 'story') return RM.STORY_RISK_SCHEMES.indexOf(s) !== -1 ? s : 'none';
    return RM.RISK_SCHEMES[s] ? s : 'none';
  };
  RM.riskEnabled = function (state, kind) { return RM.riskSchemeOf(state, kind) !== 'none'; };
  RM.riskOrderOf = function (state, kind) {
    return (RM.RISK_SCHEMES[RM.riskSchemeOf(state, kind)].order || []).slice();
  };
  RM.riskColLabel = function (state, kind) {
    return RM.RISK_SCHEMES[RM.riskSchemeOf(state, kind)].label;
  };
  RM.setRiskScheme = function (state, key, kind) {
    if (!RM.RISK_SCHEMES[key]) return;
    if (kind === 'story' && RM.STORY_RISK_SCHEMES.indexOf(key) === -1) return;
    state.meta[riskKey(kind)] = key;
    var order = RM.RISK_SCHEMES[key].order || [];
    state.items.forEach(function (it) {
      if (kind === 'story') {
        (it.stories || []).forEach(function (st) { if (st.risk && order.indexOf(st.risk) === -1) st.risk = null; });
      } else if (it.risk && order.indexOf(it.risk) === -1) it.risk = null;
    });
  };
```

In `normalizeState` after the `riskScheme` fallback (~1660), add:

```js
    if (RM.STORY_RISK_SCHEMES.indexOf(m.storyRiskScheme) === -1) {
      m.storyRiskScheme = RM.STORY_RISK_SCHEMES.indexOf(m.riskScheme) !== -1 ? m.riskScheme : 'none';
    }
    var storyRiskOrder = RM.RISK_SCHEMES[m.storyRiskScheme].order || [];
```

and change the story `risk:` line (~1819) to validate against `storyRiskOrder`. Change `[0, 1, 2, 4]` (~1397) to `[0, 1, 2, 3, 4]`. In the apps normalize (~1414) skip `sprints`: `if (a[0] === 'sprints') return;` so `m.apps.sprints` is never written. In `RM.appEnabled` add before the final return: `if (key === 'sprints') return RM.sprintsEnabled(state.meta);`.

In `js/app.js`, every story-risk reader passes `'story'`: `RM.riskColLabel(state, 'story')` (~2078), the story panel risk buttons (~4874, use `RM.riskOrderOf(state, 'story')` and hide the row when `!RM.riskEnabled(state, 'story')`), Prioritizing story Risk column option, Scoping story rows. Find them with `grep -n "st.risk\|stRisk\|'st-risk'" js/app.js`.

- [ ] **Step 4: Run both suites** — Expected: PASS (update the smoke Apps test that toggles `data-suapp="sprints"`: delete it; Task 3 replaces the Apps page).

- [ ] **Step 5: Commit** — `feat(setup): separate story risk scheme, 3-week sprints, Sprinting follows sprints` (files: `js/core.js js/app.js tests/core.test.js tests/smoke.test.js`), same trailer.

### Task 3: New Setup rail and section grouping

**Files:**
- Modify: `js/app.js` — `SETUP_SECTIONS` (~10510), `renderSetup` (~10064–10500), `setupTab` restore (~188), `HeadwayApp.openSetup` (~12106), every `setupTab = '…'` assignment (`grep -n "setupTab = '" js/app.js`), strings mentioning "Setup → Capacity" / "Setup → Apps" / "Setup → Hierarchy" / "Setup → Sizing"
- Modify: `css/app.css` — `.su-pill`, `.su-tab svg`
- Test: `tests/smoke.test.js`

**Interfaces:**
- Produces: `SETUP_SECTIONS` in the new shape (below); `sectionBody(key) → html string` (the per-section cards); `SETUP_KEY_MAP` for old keys; `setupSections()` returns the list filtered for read-only copies.

- [ ] **Step 1: Write the failing smoke test** (replace the block at ~532–551 that counts 12 tabs / 2 rail headers)

```js
{
  click(doc.querySelector('#btnSetup'));
  const keys = [...doc.querySelectorAll('#setupView .su-tab')].map(b => b.dataset.sutab);
  ok(keys.join() === 'project,sprints,org,est,budget,team,scheduling,columns,views,appearance,prefs,ai,jira',
    'Setup rail: eight sections, Views, then Personal');
  ok(doc.querySelectorAll('#setupView .su-rail-hd').length === 1 &&
    /this computer only/.test(doc.querySelector('#setupView .su-rail-hd').textContent), 'one group heading: Personal');
  ok([...doc.querySelectorAll('#setupView .su-tab')].every(b => b.querySelector('i[data-lucide]')), 'every rail item has an icon');
  ok(/Jira Integration/.test(doc.querySelector('#setupView [data-sutab="jira"]').textContent), 'Jira renamed');
  ok(!!doc.querySelector('#setupView [data-sutab="sprints"] .su-pill'), 'Sprints shows its on/off pill');
  window.__headway.openSetup && window.HeadwayApp.openSetup('capacity');
  ok(doc.querySelector('#setupView [data-sutab="scheduling"]').classList.contains('on'), 'old key capacity opens Scheduling');
  ok(!/Capacity planning/.test(doc.querySelector('#setupView').textContent), 'no "Capacity planning" copy left');
}
```

Update the other Setup smoke blocks' `data-sutab` values with this map: `timeline→project` (start/end/holidays/work week), `timeline→sprints` (sprint length), `phases→org`, `workstreams→org`, `sizing→est`, `team→budget` (rate card) / `project` (work week), `capacity→scheduling`, `apps→views`.

- [ ] **Step 2: Run** `NODE_PATH=./node_modules node tests/smoke.test.js` — Expected: FAIL on the rail order.

- [ ] **Step 3: Implement**

```js
  var SETUP_SECTIONS = [
    ['', [
      ['project', 'Project', 'calendar-range'],
      ['sprints', 'Sprints', 'repeat'],
      ['org', 'Organization', 'layers'],
      ['est', 'Sizing, priority & risk', 'ruler'],
      ['budget', 'Budgeting', 'wallet'],
      ['team', 'Team', 'users'],
      ['scheduling', 'Scheduling', 'gauge'],
      ['columns', 'Custom columns', 'columns-3'],
      ['views', 'Views', 'layout-grid']
    ]],
    ['Personal · this computer only', [
      ['appearance', 'Appearance', 'palette'],
      ['prefs', 'Preferences', 'sliders-horizontal'],
      ['ai', 'AI assistant', 'sparkles'],
      ['jira', 'Jira Integration', 'link']
    ]]
  ];
  var SETUP_KEY_MAP = { timeline: 'project', phases: 'org', workstreams: 'org', sizing: 'est',
    capacity: 'scheduling', apps: 'views' };
  function normSetupTab(k) {
    k = SETUP_KEY_MAP[k] || k;
    var all = [];
    SETUP_SECTIONS.forEach(function (s) { s[1].forEach(function (t) { all.push(t[0]); }); });
    return all.indexOf(k) !== -1 ? k : 'project';
  }
```

Use `normSetupTab` in the UI-snapshot restore (~188), in `openSetup`, and wherever `setupTab` is assigned from outside the rail. The rail renderer skips the heading when `sec[0] === ''` and appends a pill for three keys:

```js
  function setupPill(key) {
    if (key === 'sprints') return RM.sprintsEnabled(state.meta) ? state.meta.weeksPerSprint + ' wk' : 'Off';
    if (key === 'budget') return RM.appEnabled(state, 'budget') ? 'On' : 'Off';
    if (key === 'scheduling') return state.meta.capacityEnabled ? 'On' : 'Off';
    return '';
  }
  // in the rail: '<button class="su-tab…" data-sutab="…"><i data-lucide="' + t[2] + '"></i>' + t[1] +
  //   (setupPill(t[0]) ? '<span class="su-pill">' + setupPill(t[0]) + '</span>' : '') + '</button>'
```

Regroup the existing `tabBodies` cards (no new controls in this task — only moves):
- `project`: the `Project` + `Timeline` cards from `timeline`, then `Work week` from `team`, then `Holidays` from `timeline`.
- `sprints`: the `Sprints` card from `timeline`; add `[3, '3 weeks']` to the length options and relabel `[0, 'Disabled']` as `[0, 'Off']`.
- `org`: `Phases`, then `Workstreams` + `Epics`, then `hierCard`.
- `est`: the existing `sizing` body (Task 4 rebuilds it).
- `budget`: `Roles & rate card` from `team` (Task 5 adds the switch).
- `team`: placeholder of the Resources hint for now (Task 5 builds the table).
- `scheduling`: the existing `capacity` body with its h2 "Capacity planning" → "Schedule against capacity" and the enable label "Enable capacity planning" → "Schedule against capacity".
- `columns`, `appearance`, `prefs`, `ai`, `jira`: unchanged bodies.
- `views`: the existing `apps` card retitled "Views", rows for `sprints` and `budget` rendered as a non-toggle `<span class="su-pill">` "Follows Sprints" / "Follows Budgeting", Planning "Always on".

Replace user-facing strings: `grep -n "Setup → Capacity\|Setup → Capacity\|Capacity planning\|Enable capacity planning\|Setup → Apps\|Setup → Sizing\|Setup → Hierarchy" js/*.js` → "Setup → Scheduling", "Setup → Views", "Setup → Sizing, priority & risk", "Setup → Organization".

CSS:

```css
.su-pill { margin-left: auto; font-size: 11px; color: var(--ink-2); background: var(--well); border-radius: 10px; padding: 2px 8px; }
.su-tab.on .su-pill { color: var(--blue); background: var(--blue-soft); }
```

- [ ] **Step 4: Run both suites** — Expected: PASS.

- [ ] **Step 5: Commit** — `feat(setup): regroup Setup into eight project sections, Views and Personal` (files: `js/app.js css/app.css tests/smoke.test.js CHANGELOG.md`; changelog line "Setup is regrouped into Project, Sprints, Organization, Sizing, Budgeting, Team, Scheduling and Custom columns; Capacity is now Scheduling").

### Task 4: Sizing, priority & risk grid

**Files:**
- Modify: `js/app.js` — the `est` body (was `sizing`, built from `schemeRowsFor`, `sizeOptRowsFor`, `sizingCardsFor` at ~10064–10120) and its `change`/`click` handlers (`data-suscheme`, `data-suszlabel`, `data-susz`, `data-suszrm`, `#suSzAdd`)
- Modify: `css/app.css` — `.su-grid`, `.su-chip`
- Test: `tests/smoke.test.js`

**Interfaces:**
- Consumes: `RM.setRiskScheme(state, key, kind)` (Task 2), `RM.sizeSchemesFor(kind)`, `RM.PRIORITY_SCHEMES`, `RM.RISK_SCHEMES`.
- Produces: `<select data-suscheme-kind="size|prio|risk" data-kind="feature|story">` cells.

- [ ] **Step 1: Write the failing smoke test**

```js
{
  openSetupTab('est');
  const cells = doc.querySelectorAll('#setupView .su-grid select[data-suscheme-kind]');
  ok(cells.length === 6, 'grid: size, priority, risk × feature, story');
  const storyPrio = doc.querySelector('#setupView select[data-suscheme-kind="prio"][data-kind="story"]');
  ok(![...storyPrio.options].some(o => o.value === 'rice'), 'RICE is not offered for stories');
  const storyRisk = doc.querySelector('#setupView select[data-suscheme-kind="risk"][data-kind="story"]');
  ok(![...storyRisk.options].some(o => o.value === 'auto'), 'Risk (auto) is not offered for stories');
  storyRisk.value = 'confidence'; storyRisk.dispatchEvent(new window.Event('change', { bubbles: true }));
  ok(state().meta.storyRiskScheme === 'confidence', 'story risk scheme is set from the grid');
  ok(doc.querySelectorAll('#setupView .su-chip').length > 0, 'priority / risk levels show as read-only chips');
  ok(!doc.querySelector('#setupView .su-chip input'), 'chips are not editable');
  ok(!!doc.querySelector('#setupView [data-susz]'), 'size options stay editable');
}
```

(`openSetupTab(k)` is the existing helper at ~533; keep it.)

- [ ] **Step 2: Run** — Expected: FAIL (no `.su-grid`).

- [ ] **Step 3: Implement** — replace the scheme-button stacks with one grid card and keep the size option tables as the "Feature scales" / "Story scales" cards:

```js
  function schemeSelect(kindAttr, level, cur, options) {
    return '<select data-suscheme-kind="' + kindAttr + '" data-kind="' + level + '">' + options.map(function (o) {
      return '<option value="' + o[0] + '"' + (o[0] === cur ? ' selected' : '') + '>' + esc(o[1]) + '</option>';
    }).join('') + '</select>';
  }
  function estGridHtml() {
    var m = state.meta;
    function sizeOpts(kind) {
      var o = RM.sizeSchemesFor(kind).map(function (k) { return [k, RM.SIZE_SCHEMES[k].name]; });
      if (m[RM.sizeKeys(kind).scheme] === 'custom') o.push(['custom', 'Custom']);
      return o;
    }
    function prioOpts(kind) {
      return Object.keys(RM.PRIORITY_SCHEMES).filter(function (k) { return kind !== 'story' || k !== 'rice'; })
        .map(function (k) { return [k, RM.PRIORITY_SCHEMES[k].name]; });
    }
    function riskOpts(kind) {
      return RM.RISK_SCHEME_ORDER.filter(function (k) { return kind !== 'story' || RM.STORY_RISK_SCHEMES.indexOf(k) !== -1; })
        .map(function (k) { return [k, RM.RISK_SCHEMES[k].name]; });
    }
    var rows = [
      ['Size', 'Sets bar length', 'size', function (k) { return m[RM.sizeKeys(k).scheme]; }, sizeOpts],
      ['Priority', 'Ranks the backlog', 'prio', function (k) { return m[k === 'story' ? 'storyPriorityScheme' : 'priorityScheme']; }, prioOpts],
      ['Risk', 'Flags what may slip', 'risk', function (k) { return RM.riskSchemeOf(state, k); }, riskOpts]
    ];
    return '<section class="su-card"><div class="su-grid"><span></span><b>' + esc(lvl('feature')) + ' level</b><b>' + esc(lvl('story')) + ' level</b>' +
      rows.map(function (r) {
        return '<div><b>' + r[0] + '</b><div class="m-hint">' + r[1] + '</div></div>' +
          schemeSelect(r[2], 'feature', r[3]('feature'), r[4]('feature')) +
          schemeSelect(r[2], 'story', r[3]('story'), r[4]('story'));
      }).join('') + '</div></section>';
  }
  function ladderChips(order, labelFn) {
    return '<div class="su-chips">' + order.map(function (v) { return '<span class="su-chip">' + esc(labelFn(v)) + '</span>'; }).join('') + '</div>';
  }
```

The scales cards: for each level render "Size" (today's `sizingCardsFor(kind)` body), "Priority" (`ladderChips(RM.prioOrderOf(state, kind), prioValueLabel)` or the scheme's `desc` when it has no `order`), "Risk" (`ladderChips(RM.riskOrderOf(state, kind), riskValueLabel)` or `desc`). Use the existing label helpers (`grep -n "function prioValueLabel\|function riskValueLabel" js/app.js`; if `prioValueLabel` does not exist, use `RM.PRIORITY_SCHEMES[scheme].labels[v] || v`).

One `change` handler for the grid:

```js
    var sk = e.target.closest('[data-suscheme-kind]');
    if (sk) {
      var kind = sk.dataset.kind, v = sk.value;
      if (sk.dataset.suschemeKind === 'size') {
        commit('size scheme', function (s2) { RM.setSizeScheme(s2, v, kind); });
      } else if (sk.dataset.suschemeKind === 'prio') {
        commit('priority scheme', function (s2) { RM.setPriorityScheme(s2, v, kind); });
      } else {
        commit('risk scheme', function (s2) { RM.setRiskScheme(s2, v, kind); });
      }
      return;
    }
```

(Use the existing scheme setters the old buttons called — `grep -n "data-suscheme\|data-suprio\|data-surisk" js/app.js` for their names — and delete the old button handlers.)

CSS:

```css
.su-grid { display: grid; grid-template-columns: 1fr 1.6fr 1.6fr; gap: 10px 14px; align-items: center; }
.su-grid > b { font-size: 11.5px; font-weight: 500; color: var(--ink-3); }
.su-chips { display: flex; flex-wrap: wrap; gap: 6px; }
.su-chip { font-size: 12px; background: var(--well); color: var(--ink-2); border-radius: 10px; padding: 3px 9px; }
```

- [ ] **Step 4: Run both suites** — Expected: PASS. Remove smoke assertions that clicked the old `.su-scheme` buttons, replacing each with the equivalent `select` change.

- [ ] **Step 5: Commit** — `feat(setup): sizing, priority and risk as one feature × story grid` (files: `js/app.js css/app.css tests/smoke.test.js CHANGELOG.md`).

### Task 5: Budgeting switch and Team table

**Files:**
- Modify: `js/app.js` — `budget` and `team` bodies; setup `change`/`click` handlers; the Budgeting roster capacity-type chip (~9458) and the Resources panel capacity-type menu (~9814) become read-only
- Modify: `css/app.css` — `.su-team`
- Test: `tests/smoke.test.js`

**Interfaces:**
- Consumes: `RM.syncMemberCapTypes(state)` (Task 1), `RM.appEnabled`.
- Produces: data attributes `data-subudget` (switch), `data-sutm="<field>"` + `data-mid="<member id>"`, `data-sutmws` (workstream toggle), `data-sutmdel`, `#suTmAdd`.

- [ ] **Step 1: Write the failing smoke test**

```js
{
  openSetupTab('budget');
  const sw = doc.querySelector('#setupView [data-subudget]');
  ok(!!sw, 'Budgeting has a Track budget switch');
  const was = window.RM.appEnabled(state(), 'budget');
  click(sw);
  ok(window.RM.appEnabled(state(), 'budget') === !was, 'the switch drives the Budgeting tab');
  if (!window.RM.appEnabled(state(), 'budget')) click(doc.querySelector('#setupView [data-subudget]'));

  openSetupTab('team');
  const rows = doc.querySelectorAll('#setupView .su-team [data-mid]');
  ok(rows.length === state().team.length, 'one row per person');
  const m0 = state().team[0];
  const alloc = doc.querySelector('#setupView [data-sutm="capacity"][data-mid="' + m0.id + '"]');
  alloc.value = '50'; alloc.dispatchEvent(new window.Event('change', { bubbles: true }));
  ok(state().team[0].capacity === 0.5, 'allocation % edits capacity');
  const name = doc.querySelector('#setupView [data-sutm="name"][data-mid="' + m0.id + '"]');
  name.value = 'Renamed'; name.dispatchEvent(new window.Event('change', { bubbles: true }));
  ok(state().team[0].name === 'Renamed', 'name is editable');
  ok(!doc.querySelector('#setupView [data-sutm="capType"]'), 'capacity type is not editable here');
  ok(!/Paste/.test(doc.querySelector('#setupView .su-team').parentNode.textContent), 'no paste-from-spreadsheet');
  const n = state().team.length;
  click(doc.querySelector('#suTmAdd'));
  ok(state().team.length === n + 1, 'Add person');
  click(doc.querySelector('#setupView [data-sutmdel="' + state().team[n].id + '"]'));
  ok(state().team.length === n, 'delete person');
}
```

- [ ] **Step 2: Run** — Expected: FAIL.

- [ ] **Step 3: Implement**

Budget body: a first card with `<label class="p-check"><input type="checkbox" data-subudget' + (RM.appEnabled(state, 'budget') ? ' checked' : '') + '> Track budget</label>` and hint "Adds the Budgeting tab: hours and cost by role, week and phase.", then the Roles & rate card only when on; when off, `<div class="m-hint">Roles are named on the Team section as you add people.</div>`. Handler:

```js
    if (e.target.matches('[data-subudget]')) {
      var on = e.target.checked;
      commit('budgeting', function (s2) { s2.meta.apps.budget = on; });
      return;
    }
```

Team body:

```js
  function teamTableHtml() {
    var pts = state.meta.capMode === 'points' && state.meta.capacityEnabled;
    var bud = RM.appEnabled(state, 'budget'), cap = state.meta.capacityEnabled;
    var head = ['Name', 'Title', 'Role'].concat(cap ? ['Capacity type'] : [], ['Workstreams', 'Allocation'],
      pts ? ['Points / sprint'] : [], bud ? ['Rate', 'Cost'] : [], ['']);
    var roleOpts = function (cur) {
      return '<option value="">—</option>' + state.teamTypes.map(function (r) {
        return '<option' + (r === cur ? ' selected' : '') + '>' + esc(r) + '</option>';
      }).join('') + '<option value="__new">New role…</option>';
    };
    var rc = state.meta.rateCard || {};
    return '<table class="hol-table su-team"><thead><tr>' + head.map(function (h) { return '<th>' + h + '</th>'; }).join('') + '</tr></thead><tbody>' +
      state.team.map(function (m) {
        var id = ' data-mid="' + esc(m.id) + '"';
        var card = rc[m.type] || {};
        return '<tr' + id + '>' +
          '<td><input data-sutm="name"' + id + ' value="' + esc(m.name) + '" aria-label="Name"></td>' +
          '<td><input data-sutm="role"' + id + ' value="' + esc(m.role || '') + '" aria-label="Title"></td>' +
          '<td><select data-sutm="type"' + id + ' aria-label="Role">' + roleOpts(m.type) + '</select></td>' +
          (cap ? '<td class="m-hint">' + esc(m.capType || '—') + (m.type ? ' · from role' : ' · set a role to change') + '</td>' : '') +
          '<td>' + state.wsOrder.map(function (w) {
            var on = (m.workstreams || []).indexOf(w) !== -1;
            return '<button class="su-wsdot' + (on ? ' on' : '') + '" data-sutmws="' + esc(w) + '"' + id + ' title="' + esc(w) + '">' + esc(shorten(w, 10)) + '</button>';
          }).join('') + '</td>' +
          '<td><input type="number" min="0" step="5" data-sutm="capacity"' + id + ' value="' + Math.round(m.capacity * 100) + '" aria-label="Allocation percent"> %</td>' +
          (pts ? '<td><input type="number" min="0" data-sutm="points"' + id + ' value="' + (m.points == null ? '' : m.points) + '" placeholder="' + state.meta.defaultPoints + '"></td>' : '') +
          (bud ? '<td><input type="number" min="0" data-sutm="rate"' + id + ' value="' + (m.rate || '') + '" placeholder="' + (card.rate || '') + '"></td>' +
                 '<td><input type="number" min="0" data-sutm="cost"' + id + ' value="' + (m.cost || '') + '" placeholder="' + (card.cost || '') + '"></td>' : '') +
          '<td class="hol-x"><button data-sutmdel="' + esc(m.id) + '" title="Remove person"><i data-lucide="x"></i></button></td></tr>';
      }).join('') + '</tbody></table>' +
      '<button id="suTmAdd" style="margin-top:8px"><i data-lucide="plus"></i> Add person</button>' +
      '<div class="m-hint">Allocation here is each person’s default. You can change allocation and hours week by week later in the Resources panel under the timeline.</div>';
  }
```

(If `state.wsOrder` is not the workstream order list, use the helper the Workstreams card uses — `grep -n "wsOrder" js/app.js | head`.)

Handlers (in the setup `change` listener):

```js
    var tf = e.target.closest('[data-sutm]');
    if (tf) {
      var mid = tf.dataset.mid, f = tf.dataset.sutm, v = tf.value;
      if (f === 'type' && v === '__new') { teamNewRoleFor = mid; renderSetupHost(); return; }
      commit('team', function (s2) {
        var m = s2.team.filter(function (x) { return x.id === mid; })[0];
        if (!m) return;
        if (f === 'capacity') m.capacity = Math.max(0, (+v || 0) / 100);
        else if (f === 'points') m.points = v === '' ? null : Math.max(0, +v || 0);
        else if (f === 'rate' || f === 'cost') m[f] = Math.max(0, +v || 0);
        else m[f] = String(v);
        if (f === 'type') RM.syncMemberCapTypes(s2);
      });
      return;
    }
```

"New role…": `var teamNewRoleFor = null;` (module scope). When a row's id equals it, `teamTableHtml` renders `<input data-sutmnewrole data-mid="…" placeholder="New role name">` in the Role cell instead of the select, and focuses it after render. Its `change` handler: `var name = value.trim(); teamNewRoleFor = null; if (!name) { renderSetupHost(); return; } commit('add role', function (s2) { if (s2.teamTypes.indexOf(name) === -1) s2.teamTypes.push(name); var m = …by id…; m.type = name; RM.syncMemberCapTypes(s2); });`. `renderSetupHost()` = `wz ? renderWizard() : renderSetup()`. Never use `window.prompt` (it blocks the browser extension and jsdom).

Click handlers: `[data-sutmws]` toggles the workstream in `m.workstreams` and sets `m.workstream = m.workstreams[0] || ''`; `[data-sutmdel]` removes the person (use the same helper the Resources panel delete uses — `grep -n "removeMember\|team.splice" js/app.js`); `#suTmAdd` pushes `RM.newMember ? RM.newMember(state) : { id: RM.uid(), name: 'New person', type: '', capacity: 1, workstreams: [] }` through `commit('add person', …)` then `RM.normalizeState`-equivalent defaults via the same path the Resources "+" uses.

Read-only per-person capacity type: at ~9458 drop the chip's click attribute and add `title="From role"`; at ~9814 remove the capacity-type submenu from the person context menu.

CSS: `.su-team input, .su-team select { width: 100%; min-width: 0; } .su-team td { padding: 4px 6px; } .su-wsdot { font-size: 11px; padding: 1px 6px; margin: 1px; border-radius: 8px; } .su-wsdot.on { background: var(--blue-soft); color: var(--blue); }`.

- [ ] **Step 4: Run both suites** — Expected: PASS.

- [ ] **Step 5: Commit** — `feat(setup): Budgeting switch and an editable Team table` (files: `js/app.js css/app.css tests/smoke.test.js CHANGELOG.md`).

### Task 6: Scheduling roles, explainer, and column delete

**Files:**
- Modify: `js/app.js` — `scheduling` body (capacity types card ~10380–10400, tracked card), `columns` body and handlers
- Modify: `css/app.css` — `.su-sched-*`
- Test: `tests/smoke.test.js`

**Interfaces:**
- Consumes: `RM.setRoleCapType`, `RM.capTypeRoles` (Task 1).
- Produces: `data-surolect="<role>"` selects; `data-sucoldel="<col key>"` buttons.

- [ ] **Step 1: Write the failing smoke test**

```js
{
  openSetupTab('scheduling');
  ok(!doc.querySelector('#setupView [data-sucaprow]'), 'no Tracked checkboxes');
  ok(!/Tracked/.test(doc.querySelector('#setupView').textContent), 'no "Tracked" copy');
  const sel = doc.querySelector('#setupView [data-surolect]');
  ok(!!sel, 'each role picks the capacity type it supplies');
  const role = sel.dataset.surolect, t = state().capTypes[0];
  sel.value = t; sel.dispatchEvent(new window.Event('change', { bubbles: true }));
  ok(state().roleCapTypes[role] === t, 'role → type saved');
  ok(state().team.filter(m => m.type === role).every(m => m.capType === t), 'people follow their role');
  ok(/How scheduling works/.test(doc.querySelector('#setupView').textContent), 'explainer card present');

  openSetupTab('columns');
  const dels = [...doc.querySelectorAll('#setupView [data-sucoldel]')];
  ok(dels.length > 0 && dels.filter(b => b.disabled).length > 0, 'built-in columns show a disabled delete');
  const live = dels.find(b => !b.disabled);
  if (live) {
    const k = live.dataset.sucoldel;
    click(live);
    ok(!state().meta.scopeCols.some(c => c.key === k), 'deleting a custom column removes it');
  }
}
```

(The fixture has custom columns only if `seed.fixture.js` does; if not, the `if (live)` guard skips. Add one custom column at the top of the block via the existing Add column input to make it deterministic: set `#suColAdd` value and click its button — `grep -n "suColAdd\|data-sucoladd" js/app.js` for the real id.)

- [ ] **Step 2: Run** — Expected: FAIL.

- [ ] **Step 3: Implement**

Capacity types card: keep the type list (rename / drag / remove) and add below it a "Roles" table — one row per `state.teamTypes` role with `<select data-surolect="role"><option value="">None</option>…types…</select>`. Replace the Tracked card with nothing. Handler:

```js
    var rs = e.target.closest('[data-surolect]');
    if (rs) {
      var role = rs.dataset.surolect, t = rs.value;
      commit('role capacity type', function (s2) { RM.setRoleCapType(s2, role, t); });
      return;
    }
```

Explainer card (static HTML, always shown at the end of the section):

```js
  function schedExplainerHtml() {
    var SUP = [2, 2, 2, 1.2, 2, 2, 2, 2, 2, 2];
    function chart(title, note, dem) {
      return '<div class="su-sched-chart"><div><b>' + title + '</b> <span class="m-hint">' + note + '</span></div><div class="su-sched-bars">' +
        dem.map(function (d, i) {
          var ok = Math.min(d, SUP[i]), over = Math.max(0, d - SUP[i]);
          return '<div class="su-sched-wk"><i class="over" style="height:' + Math.round(over * 36) + 'px"></i><i class="ok" style="height:' + Math.round(ok * 36) + 'px"></i>' +
            '<i class="sup" style="bottom:' + Math.round(SUP[i] * 36) + 'px"></i></div>';
        }).join('') + '</div></div>';
    }
    return '<section class="su-card"><h2>How scheduling works</h2>' +
      '<div class="m-hint">Every week, for each capacity type, the work in flight has to fit inside what the people supply.</div>' +
      '<div class="su-sched-keys">' +
      '<div><b>Supply</b><div class="m-hint">Each person gives their role’s capacity type their weekly hours, or points per sprint. Holidays and time off lower it.</div></div>' +
      '<div><b>Demand</b><div class="m-hint">Each ' + esc(lvl(RM.planLevel(state)).toLowerCase()) + ' in flight asks its capacity type for one person × its multiplier, or its points spread over its weeks.</div></div>' +
      '<div><b>Over capacity</b><div class="m-hint">Demand above the line shows red on the timeline. Auto timeline moves work to the first week with room, after its dependencies.</div></div></div>' +
      '<div class="su-sched-charts">' + chart('As drawn', 'weeks 3–4 over capacity', [1, 2, 3, 3, 1, 1, 0, 0, 0, 0]) +
      chart('After Auto timeline', 'overflow moved to weeks 5–7', [1, 2, 2, 1.2, 2, 2, 0.8, 0, 0, 0]) + '</div></section>';
  }
```

CSS:

```css
.su-sched-keys, .su-sched-charts { display: grid; grid-template-columns: repeat(3, minmax(0, 1fr)); gap: 12px; }
.su-sched-charts { grid-template-columns: repeat(2, minmax(0, 1fr)); margin-top: 12px; }
.su-sched-bars { display: grid; grid-template-columns: repeat(10, 1fr); gap: 4px; height: 116px; align-items: end; border-bottom: 1px solid var(--line-2); }
.su-sched-wk { position: relative; height: 116px; display: flex; flex-direction: column; justify-content: flex-end; }
.su-sched-wk .ok { display: block; background: #7FAEDD; }
.su-sched-wk .over { display: block; background: var(--err); }
.su-sched-wk .sup { position: absolute; left: -2px; right: -2px; height: 2px; background: var(--ink); }
```

Columns: add a last cell per row:

```js
'<td class="hol-x"><button data-sucoldel="' + esc(c.key) + '"' + (isCustom ? ' title="Delete column"' : ' disabled title="Built-in columns can’t be deleted"') + '><i data-lucide="x"></i></button></td>'
```

where `isCustom = /^c/.test(c.key)` (custom keys start with `c`, per DESIGN.md). Click handler: `commit('delete column', function (s2) { s2.meta.scopeCols = s2.meta.scopeCols.filter(function (x) { return x.key !== key; }); (s2.meta.scopeColOrder || []).splice(…); s2.items.forEach(function (it) { if (it.custom) delete it.custom[key]; }); })` — mirror the existing custom-column removal if one exists (`grep -n "scopeCols.filter\|removeScopeCol" js/*.js`) and call it instead.

- [ ] **Step 4: Run both suites** — Expected: PASS.

- [ ] **Step 5: Commit** — `feat(setup): roles pick their capacity type, scheduling explainer, delete custom columns` (files: `js/app.js css/app.css tests/smoke.test.js CHANGELOG.md`).

---

# Phase 2 — Onboarding wizard

### Task 7: Presets in core

**Files:**
- Modify: `js/core.js` (after `RM.APPS`)
- Test: `tests/core.test.js`

**Interfaces:**
- Produces: `RM.PRESETS` (array of `{ key, name, desc, sprints, v }`), `RM.applyPreset(state, key) → boolean`. Keys: `scrum`, `ascrum`, `rapid`, `minimal`.

- [ ] **Step 1: Write the failing tests**

```js
// ------------------------------------------------------------- presets
section('presets');
function blankish() {
  return RM.normalizeState({ meta: { timelineStart: '2026-07-27', numWeeks: 8, title: 'Keep me' }, phases: [{ id: 'p1', name: 'Build' }], items: [] });
}
var want = {
  scrum:   { sizeScheme: 'none',   storySizeScheme: 'fibonacci', priorityScheme: 'levels', storyPriorityScheme: 'moscow', riskScheme: 'none', storyRiskScheme: 'none', weeksPerSprint: 2, budget: true,  capacityEnabled: true,  planLevel: 'story',   capMode: 'points' },
  ascrum:  { sizeScheme: 'tshirt', storySizeScheme: 'fibonacci', priorityScheme: 'levels', storyPriorityScheme: 'moscow', riskScheme: 'risk', storyRiskScheme: 'risk', weeksPerSprint: 2, budget: true,  capacityEnabled: true,  planLevel: 'story',   capMode: 'points' },
  rapid:   { sizeScheme: 'tshirt', storySizeScheme: 'none',      priorityScheme: 'none',   storyPriorityScheme: 'none',   riskScheme: 'none', storyRiskScheme: 'none', weeksPerSprint: 0, budget: false, capacityEnabled: true,  planLevel: 'feature', capMode: 'person' },
  minimal: { sizeScheme: 'tshirt', storySizeScheme: 'none',      priorityScheme: 'none',   storyPriorityScheme: 'none',   riskScheme: 'none', storyRiskScheme: 'none', weeksPerSprint: 0, budget: false, capacityEnabled: false }
};
Object.keys(want).forEach(function (k) {
  var s = blankish();
  ok(RM.applyPreset(s, k), k + ' applies');
  var m = s.meta, w = want[k];
  var got = { sizeScheme: m.sizeScheme, storySizeScheme: m.storySizeScheme, priorityScheme: m.priorityScheme,
    storyPriorityScheme: m.storyPriorityScheme, riskScheme: m.riskScheme, storyRiskScheme: m.storyRiskScheme,
    weeksPerSprint: m.weeksPerSprint, budget: m.apps.budget, capacityEnabled: m.capacityEnabled };
  if ('planLevel' in w) { got.planLevel = m.planLevel; got.capMode = m.capMode; }
  eq(got, w, k + ' writes exactly its settings');
  eq([m.title, s.phases[0].name], ['Keep me', 'Build'], k + ' leaves name and phases alone');
});
ok(!RM.applyPreset(blankish(), 'nope'), 'unknown preset is refused');
var sSw = blankish(); RM.applyPreset(sSw, 'ascrum'); RM.applyPreset(sSw, 'minimal');
eq([sSw.meta.riskScheme, sSw.meta.weeksPerSprint], ['none', 0], 'switching presets overwrites the preset-owned fields');
eq(RM.normalizeState(RM.clone(sSw)).meta.sizeScheme, 'tshirt', 'a preset survives normalize');
```

- [ ] **Step 2: Run** — Expected: FAIL (`RM.applyPreset is not a function`).

- [ ] **Step 3: Implement**

```js
  // Onboarding presets: each writes ONLY these fields (schemes, sprints,
  // budget, scheduling). Name, dates, phases, people and columns are never
  // touched, so switching presets is safe after later steps were edited.
  RM.PRESETS = [
    { key: 'scrum', name: 'Scrum', sprints: true,
      desc: 'Stories carry the points; a feature is the span of its stories. Work is scheduled into 2-week sprints by story points, with budget tracking.',
      v: { size: ['none', 'fibonacci'], prio: ['levels', 'moscow'], risk: ['none', 'none'], wps: 2, budget: true, cap: { on: true, plan: 'story', mode: 'points' } } },
    { key: 'ascrum', name: 'Advanced Scrum', sprints: true,
      desc: 'Scrum plus a size on every feature, and priority and risk on both features and stories.',
      v: { size: ['tshirt', 'fibonacci'], prio: ['levels', 'moscow'], risk: ['risk', 'risk'], wps: 2, budget: true, cap: { on: true, plan: 'story', mode: 'points' } } },
    { key: 'rapid', name: 'Rapid Delivery', sprints: false,
      desc: 'T-shirt sizes on features, no sprints. Each feature takes a person, and Auto timeline packs features in as people free up.',
      v: { size: ['tshirt', 'none'], prio: ['none', 'none'], risk: ['none', 'none'], wps: 0, budget: false, cap: { on: true, plan: 'feature', mode: 'person' } } },
    { key: 'minimal', name: 'Minimal', sprints: false,
      desc: 'Just features on a timeline with T-shirt sizes. You place and stretch the bars yourself.',
      v: { size: ['tshirt', 'none'], prio: ['none', 'none'], risk: ['none', 'none'], wps: 0, budget: false, cap: { on: false } } }
  ];
  RM.applyPreset = function (state, key) {
    var p = RM.PRESETS.filter(function (x) { return x.key === key; })[0];
    if (!p) return false;
    var v = p.v, m = state.meta;
    RM.setSizeScheme(state, v.size[0], 'feature');
    RM.setSizeScheme(state, v.size[1], 'story');
    RM.setPriorityScheme(state, v.prio[0], 'feature');
    RM.setPriorityScheme(state, v.prio[1], 'story');
    RM.setRiskScheme(state, v.risk[0], 'feature');
    RM.setRiskScheme(state, v.risk[1], 'story');
    m.weeksPerSprint = v.wps;
    m.apps = m.apps || {};
    m.apps.budget = v.budget;
    m.capacityEnabled = !!v.cap.on;
    if (v.cap.on) { m.planLevel = v.cap.plan; m.capMode = v.cap.mode; }
    m.preset = key;
    return true;
  };
```

Use the real setter names — `grep -n "RM.setSizeScheme\|RM.setPriorityScheme\|RM.setPrioScheme" js/core.js`. If a setter takes `(state, kind, key)` order, adapt the calls, not the setter. `m.preset` is informational (Review shows it); add `m.preset = typeof m.preset === 'string' ? m.preset : ''` to `normalizeState`.

- [ ] **Step 4: Run** `NODE_PATH=./node_modules node tests/core.test.js` — Expected: PASS.

- [ ] **Step 5: Commit** — `feat(onboarding): project presets` (files: `js/core.js tests/core.test.js`).

### Task 8: Wizard shell over a draft state

**Files:**
- Modify: `index.html` — add `<div id="wizard" hidden></div>` directly after `</header>` of `#topbar`
- Modify: `js/app.js` — replace `newProjectModal` (~11866) with `openWizard()`; guard `commit`, `replaceState`, `saveLocal`, `scheduleAutoSave`, `maybeAskName`, the disk-reload path of `loadBuffer`; rewire the three callers (~8736, ~8754, ~11836)
- Modify: `css/app.css` — `.wz-*`, `body.wizard-open`
- Test: `tests/smoke.test.js`

**Interfaces:**
- Consumes: `sectionBody(key)` and the `#setupView` delegated handlers (Tasks 3–6), `blankState()`, `createProjectOnDisk(st)`.
- Produces: `openWizard()`, `closeWizard(discard)`, `wizardCreate()`, `wz` = `{ step, stash: { state, undo, redo, docSaved, sessionEdited, view }, dirty, welcomed }`; `HeadwayApp.wizard = { open: openWizard, go: function (k) { wz.step = k; renderWizard(); }, create: wizardCreate, close: closeWizard, reloadForTest: function (st) { wz.stash.state = RM.normalizeState(st); wz.stash.docSaved = true; } }`.

Design: while the wizard is open the global `state` IS the draft. The Setup handlers already read `state` and call `commit`; `commit` detects `wz` and only mutates + re-renders the wizard. The section host is `#setupView`'s handlers bound to a shared element: move the three `$('#setupView').addEventListener(...)` registrations into `function bindSectionHandlers(host)` and call it for both `#setupView` and `#wizard`.

- [ ] **Step 1: Write the failing smoke test** (new block after the Setup blocks)

```js
{
  const before = JSON.stringify(state());
  window.HeadwayApp.wizard.open();
  ok(doc.body.classList.contains('wizard-open') && !doc.querySelector('#wizard').hidden, 'New project opens the full-screen wizard');
  ok(!!doc.querySelector('#topbar[data-tauri-drag-region], #topbar .tb-mark'), 'the app top bar stays (draggable)');
  ok(doc.querySelectorAll('#wizard .wz-step').length === 9, 'nine numbered steps');
  ok(doc.querySelector('#wizard [data-wz="next"]').disabled, 'Continue waits for a preset');
  // a disk reload of the open document while the wizard is up lands in the stash, not the draft
  const reloaded = JSON.parse(before); reloaded.meta.title = 'Changed on disk';
  window.HeadwayApp.wizard.reloadForTest(reloaded);
  ok(doc.body.classList.contains('wizard-open') && state().meta.title !== 'Changed on disk', 'reload leaves the wizard and draft alone');
  window.HeadwayApp.wizard.close(true);
  ok(state().meta.title === 'Changed on disk', 'and the reloaded document is what comes back');
  const beforeB = JSON.stringify(state());
  window.HeadwayApp.wizard.open();
  click(doc.querySelector('#wizard [data-wzpreset="minimal"]') || doc.body);
  window.HeadwayApp.wizard.close(true);
  ok(JSON.stringify(state()) === beforeB, 'closing after draft edits restores the open document exactly');
  ok(!doc.body.classList.contains('wizard-open'), 'wizard gone');
}
```

- [ ] **Step 2: Run** — Expected: FAIL (`HeadwayApp.wizard` undefined).

- [ ] **Step 3: Implement**

```js
  var wz = null;
  var WZ_STEPS = SETUP_SECTIONS[0][1].slice(0, 8).map(function (t) { return t[0]; }).concat(['review']);
  function openWizard() {
    wz = { step: 'project', dirty: false, stash: {
      state: state, undo: undoStack.slice(), redo: redoStack.slice(),
      docSaved: docSaved, sessionEdited: sessionEdited, view: view } };
    state = blankState();
    undoStack.length = 0; redoStack.length = 0;
    if (!userName() && !localStorageGet('headway-onboarded-v1')) wz.step = 'welcome';
    document.body.classList.add('wizard-open');
    $('#wizard').hidden = false;
    renderWizard();
  }
  function closeWizard(discard) {
    if (!wz) return;
    var s = wz.stash;
    var draft = state;
    state = s.state;
    undoStack.length = 0; Array.prototype.push.apply(undoStack, s.undo);
    redoStack.length = 0; Array.prototype.push.apply(redoStack, s.redo);
    docSaved = s.docSaved; sessionEdited = s.sessionEdited; view = s.view;
    wz = null;
    document.body.classList.remove('wizard-open');
    $('#wizard').hidden = true;
    $('#wizard').innerHTML = '';
    validation = RM.validate(state);
    render();
    return discard ? null : draft;
  }
  function wizardCreate() {
    var draft = closeWizard(false);
    try { localStorage.setItem('headway-onboarded-v1', '1'); } catch (e) { /* storage optional */ }
    var flush = (window.HeadwayDesktop && !docSaved && HeadwayDesktop.currentPath())
      ? (doSave(false, true) || Promise.resolve()) : Promise.resolve();
    flush.then(function () { createProjectOnDisk(draft); }, function () { createProjectOnDisk(draft); });
  }
  function localStorageGet(k) { try { return localStorage.getItem(k); } catch (e) { return null; } }
```

Guards (first line of each):

```js
  // commit:        if (wz) { if (mutate) mutate(state); wz.dirty = true; RM.applySizeRollup(state); renderWizard(); return; }
  // replaceState:  if (wz) { state = next; renderWizard(); return; }
  // saveLocal:     if (wz) return;
  // scheduleAutoSave: if (wz) return;
  // maybeAskName:  if (wz) return;
```

Disk reload while open (`loadBuffer(bytes, name, true)` path, `reloadingDoc`): parse and normalize as today, but when `wz` is set assign the result to `wz.stash.state` and set `wz.stash.docSaved = true`, then return without touching `state` or rendering.

`renderWizard()`:

```js
  function renderWizard() {
    var host = $('#wizard');
    var idx = WZ_STEPS.indexOf(wz.step);
    var need = wz.step === 'project' && !state.meta.preset;
    var rail = WZ_STEPS.map(function (k, i) {
      var t = k === 'review' ? ['review', 'Review & create'] : SETUP_SECTIONS[0][1][i];
      var done = i < idx;
      return '<button class="wz-step' + (k === wz.step ? ' on' : '') + (done ? ' done' : '') + '" data-wzgo="' + k + '"' +
        (!state.meta.preset && k !== 'project' ? ' disabled' : '') + '><span class="wz-num">' + (done ? '✓' : i + 1) + '</span>' + esc(t[1]) + '</button>';
    }).join('');
    var body = wz.step === 'welcome' ? welcomeHtml() : wz.step === 'review' ? reviewHtml() : sectionBody(wz.step, 'wizard');
    host.innerHTML = wz.step === 'welcome'
      ? '<div class="wz-welcome">' + body + '</div>'
      : '<div class="wz-progress"><i style="width:' + Math.round((idx + 1) / 9 * 100) + '%"></i></div>' +
        '<div class="wz-layout"><nav class="wz-rail" aria-label="New project steps">' + rail + '</nav>' +
        '<div class="wz-main"><div class="wz-content"><div class="wz-kicker">STEP ' + (idx + 1) + ' OF 9</div>' + body + '</div>' +
        '<div class="wz-foot"><span></span><button data-wz="back"' + (idx === 0 && !wz.welcomed ? ' disabled' : '') + '>Back</button>' +
        (idx > 0 && wz.step !== 'review' ? '<button data-wz="skip">Skip</button>' : '') +
        '<button class="primary" data-wz="next"' + (need ? ' disabled' : '') + '>' + (wz.step === 'review' ? 'Create project' : 'Continue') + '</button></div></div></div>';
    if (window.lucide) lucide.createIcons();
  }
```

Topbar while open: `body.wizard-open #menus, body.wizard-open .tabs-center, body.wizard-open #optBtn, body.wizard-open .tb-right { display: none; }` and set `#docTitle` value to "New project" while open (restore in `closeWizard`). The × lives at the right of the topbar: append `<button id="wzClose" aria-label="Close">` into `#topbar` on open and remove it on close (keep `.win-caption` visible on Windows).

Click handling on `#wizard`: `data-wzgo` → `wz.step = k; renderWizard()` (ignored when disabled); `data-wz="next"` → next step or `wizardCreate()` on review; `skip` → next step; `back` → previous (from `project` to `welcome` only if `wz.welcomed`); `#wzClose` and Escape (only when no modal is open) → `if (!wz.dirty || confirmDiscard()) closeWizard(true)`, where `confirmDiscard` uses the app's existing in-app confirm modal (`grep -n "function confirmModal\|function askConfirm" js/app.js`) — never `window.confirm`.

Rewire: the three `newProjectModal` references → `openWizard` (keep `guardUnsaved(openWizard)` on the start page). Delete `newProjectModal`.

`sectionBody(key, mode)`: extract from `renderSetup` the `tabBodies[key]` lookup into this function so both callers share it; `mode === 'wizard'` hides the Project section's Work week and Holidays cards behind a summary row (Task 9) and adds the preset card.

CSS:

```css
body.wizard-open #main, body.wizard-open #startPage, body.wizard-open #panelPeek, body.wizard-open #leftPeek { display: none !important; }
#wizard { position: fixed; inset: var(--topbar-h, 46px) 0 0 0; background: var(--paper); z-index: 35; display: flex; flex-direction: column; }
.wz-progress { height: 3px; background: var(--well); } .wz-progress i { display: block; height: 3px; background: var(--blue); }
.wz-layout { flex: 1; display: flex; min-height: 0; }
.wz-rail { width: 264px; padding: 16px 12px; border-right: 1px solid var(--line); background: var(--paper-2); display: flex; flex-direction: column; gap: 2px; }
.wz-step { display: flex; gap: 10px; align-items: center; min-height: 36px; padding: 0 10px; border: 0; background: transparent; border-radius: 6px; text-align: left; }
.wz-step.on { background: var(--blue-soft); color: var(--blue); font-weight: 600; }
.wz-num { width: 22px; height: 22px; border-radius: 11px; border: 1px solid var(--line-2); display: grid; place-items: center; font: 11px var(--mono); }
.wz-step.done .wz-num { background: var(--ok); border-color: var(--ok); color: #fff; }
.wz-main { flex: 1; display: flex; flex-direction: column; min-width: 0; }
.wz-content { flex: 1; overflow: auto; padding: 32px 56px; max-width: 920px; }
.wz-kicker { font: 11.5px var(--mono); color: var(--ink-3); letter-spacing: .04em; }
.wz-foot { display: flex; gap: 10px; align-items: center; padding: 14px 56px; border-top: 1px solid var(--line); background: var(--surface); }
.wz-foot span { flex: 1; }
```

(Measure `--topbar-h` from `#topbar.offsetHeight` on open if the variable does not exist.)

- [ ] **Step 4: Run both suites** — Expected: PASS.

- [ ] **Step 5: Commit** — `feat(onboarding): full-screen new-project wizard over a draft` (files: `index.html js/app.js css/app.css tests/smoke.test.js`).

### Task 9: Welcome, presets, Project step, Review, Create

**Files:**
- Modify: `js/app.js` — `welcomeHtml`, `presetCardsHtml`, `reviewHtml`, section `summary` strings, Project body in wizard mode, Sprints preview strip, wizard click handlers for presets and welcome
- Modify: `css/app.css` — `.pcard`, `.pc-sketch`, `.wz-welcome`, `.wz-sum`
- Modify: `CHANGELOG.md`, `DESIGN.md` (Setup + Views bullets)
- Test: `tests/smoke.test.js`

**Interfaces:**
- Consumes: `RM.PRESETS`, `RM.applyPreset` (Task 7), wizard API (Task 8), `setUserName`, `applyTheme` / `THEME_KEY` setter (`grep -n "function setTheme\|THEME_KEY" js/app.js`).
- Produces: `data-wzpreset="<key>"` cards; `data-wztheme="light|dark|system"`; `#wzName`.

- [ ] **Step 1: Write the failing smoke test**

```js
{
  window.localStorage.removeItem('headway-onboarded-v1');
  const savedName = window.localStorage.getItem('headway-user-v1');
  window.localStorage.removeItem('headway-user-v1');
  window.HeadwayApp.wizard.open();
  ok(/Welcome to Headway/.test(doc.querySelector('#wizard').textContent), 'first project starts on Welcome');
  const nm = doc.querySelector('#wzName'); nm.value = 'Sam'; nm.dispatchEvent(new window.Event('change', { bubbles: true }));
  click(doc.querySelector('#wizard [data-wztheme="dark"]'));
  ok(window.localStorage.getItem('headway-user-v1') === 'Sam', 'name saved per machine');
  click(doc.querySelector('#wizard [data-wz="next"]'));
  const cards = doc.querySelectorAll('#wizard [data-wzpreset]');
  ok(cards.length === 4 && [...cards].map(c => c.dataset.wzpreset).join() === 'scrum,ascrum,rapid,minimal', 'four presets');
  ok([...cards].every(c => c.querySelector('.pc-sketch')), 'each card has a sketch');
  click(doc.querySelector('#wizard [data-wzpreset="rapid"]'));
  ok(doc.querySelector('#wizard [data-wzpreset="rapid"]').classList.contains('on') &&
     !doc.querySelector('#wizard [data-wz="next"]').disabled, 'picking a preset selects it and enables Continue');
  ok(state().meta.weeksPerSprint === 0 && state().meta.capacityEnabled, 'the draft took the preset');
  window.HeadwayApp.wizard.go('review');
  const rows = doc.querySelectorAll('#wizard .wz-sum [data-wzgo]');
  ok(rows.length === 8, 'review lists the eight sections with Edit');
  ok(!/Views/.test(doc.querySelector('#wizard .wz-sum').textContent), 'no Views on review');
  window.HeadwayApp.wizard.close(true);
  if (savedName) window.localStorage.setItem('headway-user-v1', savedName);
}
```

- [ ] **Step 2: Run** — Expected: FAIL.

- [ ] **Step 3: Implement**

```js
  function welcomeHtml() {
    var cur = themePref || 'system';
    var sw = { light: ['#F5F2EC', '#FFFFFF', '#0057B8'], dark: ['#161B21', '#242C35', '#5B9BE0'], system: ['linear-gradient(90deg,#F5F2EC 50%,#161B21 50%)', 'rgba(255,255,255,.55)', '#0057B8'] };
    return '<section class="su-card wz-hello"><div class="wz-kicker">FIRST TIME ONLY</div><h1 class="su-page">Welcome to Headway</h1>' +
      '<div class="m-hint">Two things about you before the project. They’re saved on this computer, not in project files.</div>' +
      '<label class="p-lab">Your name<input id="wzName" maxlength="60" placeholder="e.g. Sam Rivera" value="' + esc(userName()) + '"></label>' +
      '<div class="p-lab">Theme</div><div class="wz-themes">' + ['light', 'dark', 'system'].map(function (k) {
        var c = sw[k];
        return '<button class="pcard' + (cur === k ? ' on' : '') + '" data-wztheme="' + k + '"><span class="wz-sw" style="background:' + c[0] + '"><i style="background:' + c[1] + '"></i><i style="background:' + c[2] + '"></i></span>' +
          (k === 'system' ? 'Match system' : k.charAt(0).toUpperCase() + k.slice(1)) + '</button>';
      }).join('') + '</div>' +
      '<div class="wz-hello-foot"><button class="primary" data-wz="next">Continue</button></div></section>';
  }
```

Wizard handlers: `#wzName` change → `setUserName(value)`; `[data-wztheme]` → the existing theme setter (the one `data-pref-theme` uses); `next` from welcome → `wz.welcomed = true; wz.step = 'project'`.

Preset cards (Project, wizard mode only, after the name/date card):

```js
  var PC_SKETCH = {
    scrum: pcSprints() + pcHull(1, 44, 26, 'Checkout') + pcStory(1, 13, 46, '3') + pcStory(15, 11, 46, '5') + pcStory(27, 18, 46, '2') + pcHull(40, 56, 66, 'Search') + pcStory(40, 22, 86, '8') + pcStory(63, 16, 86, '3'),
    ascrum: pcSprints() + pcFeat(1, 44, 26, 'Checkout · L · Must') + pcRisk(42, 30) + pcStory(1, 13, 46, '3 H') + pcStory(15, 11, 46, '5 M') + pcStory(27, 18, 46, '2 L') + pcFeat(40, 56, 66, 'Search · XL · Should') + pcRisk(93, 70) + pcStory(40, 22, 86, '8 M') + pcStory(63, 16, 86, '3 L'),
    rapid: pcMonths() + pcFeat(1, 30, 26, 'Login · M') + pcFeat(32, 44, 26, 'Billing · L') + pcFeat(1, 18, 48, 'Export · S') + pcFeat(20, 30, 48, 'Search · M') + pcFeat(51, 26, 48, 'Alerts · M') + pcFeat(78, 20, 70, 'Admin · S'),
    minimal: pcMonths() + pcFeat(1, 30, 26, 'Login · M') + pcFeat(24, 44, 46, 'Billing · L') + pcFeat(8, 18, 66, 'Export · S') + pcFeat(60, 36, 86, 'Search · M')
  };
  function pcBar(cls, x, w, y, tx) { return '<span class="pc-' + cls + '" style="left:' + x + '%;width:' + w + '%;top:' + y + 'px">' + esc(tx || '') + '</span>'; }
  function pcFeat(x, w, y, tx) { return pcBar('feat', x, w, y, tx); }
  function pcHull(x, w, y, tx) { return pcBar('hull', x, w, y, tx); }
  function pcStory(x, w, y, tx) { return pcBar('story', x, w, y, tx); }
  function pcRisk(x, y) { return pcBar('risk', x, 2.2, y, ''); }
  function pcSprints() { return [0, 1, 2, 3, 4].map(function (i) { return pcBar('hdr', 1 + i * 19.6, 19, 4, 'S' + (i + 1)); }).join(''); }
  function pcMonths() { return ['Oct', 'Nov', 'Dec', 'Jan', 'Feb'].map(function (m, i) { return pcBar('mon', 1 + i * 19.6, 19, 4, m); }).join(''); }
  function presetCardsHtml() {
    var cur = state.meta.preset;
    return '<section class="su-card"><h2>Start from a preset <span class="req" aria-hidden="true">*</span></h2>' +
      '<div class="m-hint">Pick how you work. It fills in the next steps, and you can change any of it on the way.</div>' +
      '<div class="pc-grid" role="radiogroup" aria-label="Preset">' + RM.PRESETS.map(function (p) {
        var on = cur === p.key;
        return '<button class="pcard' + (on ? ' on' : '') + '" role="radio" aria-checked="' + on + '" data-wzpreset="' + p.key + '">' +
          (on ? '<span class="pc-check" aria-hidden="true">✓</span>' : '') +
          '<b>' + esc(p.name) + '</b><span class="m-hint">' + esc(p.desc) + '</span>' +
          '<span class="pc-sketch" aria-hidden="true">' + PC_SKETCH[p.key] + '</span></button>';
      }).join('') + '</div>' + (cur ? '' : '<div class="m-hint req-hint">Choose a preset to continue.</div>') + '</section>';
  }
```

Handler: `[data-wzpreset]` → `commit('preset', function (s2) { RM.applyPreset(s2, key); })` (wizard `commit` just mutates and re-renders).

Project in wizard mode: replace the Work week + Holidays cards with one row: `'<section class="su-card su-row"><div><h2>Work week &amp; holidays</h2><div class="m-hint">' + days + ' · ' + hours + ' h per week · ' + n + ' holidays in range</div></div><button data-wzexpand="workweek">Edit</button></section>'`; `data-wzexpand` sets `wz.expandWorkWeek = true` and re-renders with the full cards.

Sprints (both modes): below the start/number fields, a strip of the next six sprints: `RM.sprintNumForWeek` + `RM.weekStartDate` for `anchorWeek + i * wps`, rendered as `<span class="su-spr"><b>S' + n + '</b>' + fmtShort(start) + ' – ' + fmtShort(endFriday) + '</span>`; when off, the hint "No sprint numbers on the timeline, and the Sprinting tab is hidden."

Review:

```js
  var WZ_SUMMARY = {
    project: function () { var p = RM.PRESETS.filter(function (x) { return x.key === state.meta.preset; })[0];
      return (state.meta.title || 'Untitled') + ' · ' + RM.fmtShort(RM.parseISO(state.meta.timelineStart)) + ' – ' + (state.meta.endDate || '') + (p ? ' · ' + p.name + ' preset' : ''); },
    sprints: function () { return RM.sprintsEnabled(state.meta) ? state.meta.weeksPerSprint + '-week sprints from S' + RM.sprintInfo(state.meta).firstNum : 'No sprints'; },
    org: function () { return state.phases.filter(function (p) { return !p.bucket; }).length + ' phases · ' + state.wsOrder.length + ' workstreams'; },
    est: function () { var m = state.meta; return 'Features: ' + RM.SIZE_SCHEMES[m.sizeScheme].name + ' · Stories: ' + RM.SIZE_SCHEMES[m.storySizeScheme].name; },
    budget: function () { return RM.appEnabled(state, 'budget') ? state.teamTypes.length + ' roles on the rate card' : 'Off'; },
    team: function () { return state.team.length + (state.team.length === 1 ? ' person' : ' people'); },
    scheduling: function () { var m = state.meta; return m.capacityEnabled ? (m.planLevel === 'story' ? 'Stories' : 'Features') + ', ' + (m.capMode === 'points' ? 'story points' : 'per person') + ', ' + state.capTypes.length + ' capacity types' : 'Off'; },
    columns: function () { return state.meta.scopeCols.filter(function (c) { return /^c/.test(c.key); }).length + ' custom columns'; }
  };
  function reviewHtml() {
    return '<h1 class="su-page">Review &amp; create</h1><div class="m-hint">Check the essentials. Anything can be changed later in Setup.</div>' +
      '<section class="su-card wz-sum">' + WZ_STEPS.slice(0, 8).map(function (k, i) {
        return '<div class="wz-sum-row"><span class="wz-num">' + (i + 1) + '</span><b>' + esc(SETUP_SECTIONS[0][1][i][1]) + '</b><span>' + esc(WZ_SUMMARY[k]()) + '</span>' +
          '<button data-wzgo="' + k + '">Edit</button></div>';
      }).join('') + '</section>';
  }
```

(Adjust `state.wsOrder`, `RM.fmtShort`, `RM.parseISO` to the real helper names — each already exists per DESIGN.md; confirm with grep.)

CSS:

```css
.pc-grid, .wz-themes { display: grid; grid-template-columns: repeat(2, minmax(0, 1fr)); gap: 10px; }
.wz-themes { grid-template-columns: repeat(3, minmax(0, 1fr)); }
.pcard { position: relative; display: flex; flex-direction: column; align-items: stretch; gap: 6px; text-align: left; padding: 14px; border: 1.5px solid var(--line); background: var(--surface); border-radius: 8px; cursor: pointer; }
.pcard:hover { border-color: var(--line-3); }
.pcard.on { border-color: var(--blue); background: var(--blue-soft); }
.pc-check { position: absolute; top: 10px; right: 10px; width: 22px; height: 22px; border-radius: 11px; background: var(--blue); color: #fff; font-weight: 700; display: grid; place-items: center; }
.pc-sketch { position: relative; display: block; height: 110px; margin-top: 8px; border: 1px solid var(--line); border-radius: 6px; overflow: hidden; pointer-events: none; }
.pcard.on .pc-sketch { background: var(--surface); }
.pc-sketch span { position: absolute; box-sizing: border-box; height: 16px; line-height: 16px; padding: 0 5px; border-radius: 3px; font: 600 9.5px var(--mono); white-space: nowrap; overflow: hidden; }
.pc-hdr { background: var(--well); color: var(--ink-2); text-align: center; height: 14px !important; line-height: 14px !important; font-weight: 400 !important; }
.pc-mon { color: var(--ink-3); border-left: 1px solid var(--line); border-radius: 0 !important; height: 14px !important; line-height: 14px !important; font-weight: 400 !important; }
.pc-feat { background: var(--c-product); color: #fff; }
.pc-hull { border: 1.5px dashed var(--c-product); color: var(--c-product); line-height: 13px !important; }
.pc-story { background: #7FAEDD; color: var(--cell-ink); height: 12px !important; line-height: 12px !important; font-weight: 400 !important; }
.pc-risk { background: var(--flag); height: 8px !important; border-radius: 4px !important; padding: 0 !important; }
.wz-welcome { flex: 1; display: grid; place-items: center; padding: 40px; overflow: auto; }
.wz-hello { width: 640px; padding: 40px 44px; gap: 18px; }
.wz-sw { display: flex; flex-direction: column; gap: 8px; height: 84px; border-radius: 6px; padding: 12px; border: 1px solid var(--line); }
.wz-sw i { display: block; height: 14px; border-radius: 3px; width: 70%; } .wz-sw i + i { width: 45%; }
.wz-sum-row { display: flex; align-items: center; gap: 14px; padding: 12px 0; border-bottom: 1px solid var(--line); }
.wz-sum-row b { width: 190px; } .wz-sum-row span:not(.wz-num) { flex: 1; color: var(--ink-2); }
```

CHANGELOG (Unreleased): "New projects open a full-screen setup: pick a preset (Scrum, Advanced Scrum, Rapid Delivery, Minimal) and walk the same sections Setup uses", "The first project asks for your name and theme". DESIGN.md: replace the **Apps** and **Setup** bullets with the new Setup rail (sections, Views, Personal) and add an **Onboarding** bullet (draft swap, guards, `headway-onboarded-v1`).

- [ ] **Step 4: Run both suites** — Expected: PASS. Then a headless Chrome pass per the verification workflow memory: open `index.html`, File › New project, walk every step, pick each preset, Create (browser download path), then open Setup and visit every rail item; no console errors.

- [ ] **Step 5: Commit** — `feat(onboarding): welcome, presets with sketches, review and create` (files: `js/app.js css/app.css tests/smoke.test.js CHANGELOG.md DESIGN.md`).

---

## Self-review notes

- Spec §1 IA → Task 3; §2 shared rendering → Tasks 3 + 8 (`sectionBody`, shared handlers); §3 Welcome → Task 9; §4 wizard → Tasks 8–9; §5 presets → Tasks 7 + 9; §6 sections: Project/Sprints → 3 + 9, Organization → 3, Sizing → 2 + 4, Budgeting/Team → 5, Scheduling → 1 + 6, Columns → 6, Views/Personal → 3; §7 apps model → Task 2; §8 testing → every task.
- Review Focus lines are pinned in: Task 1 (no-role person), Task 8 (close restores exactly; reload goes to the stash — the reload path must call the same code as `reloadForTest`), Task 3 (old `setupTab` key), Task 2 (`apps.sprints`), Task 7 (preset switching leaves name/phases).
