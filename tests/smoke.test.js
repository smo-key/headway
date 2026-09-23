/* Headless boot + interaction smoke test for the Headway UI (jsdom).
 * Run: NODE_PATH=<dir-with-node_modules> node tools/roadmapping/tests/smoke.test.js
 * Needs jsdom + exceljs resolvable; skips politely when they aren't. */
'use strict';
const fs = require('fs');
const path = require('path');

let JSDOM, ExcelJS;
try {
  JSDOM = require('jsdom').JSDOM;
  ExcelJS = require('exceljs');
} catch (e) {
  console.log('(skipped — jsdom/exceljs not resolvable in NODE_PATH)');
  process.exit(0);
}

const ROOT = path.join(__dirname, '..');
const html = fs.readFileSync(path.join(ROOT, 'index.html'), 'utf8');

const dom = new JSDOM(html, {
  url: 'http://localhost/roadmapping/index.html',
  runScripts: 'outside-only',
  pretendToBeVisual: true,
});
const { window } = dom;
window.ExcelJS = ExcelJS; // stand-in for the vendored browser build
// lucide is not loaded here — app guards every createIcons call

let failed = 0, passed = 0;
function ok(cond, name) {
  if (cond) { passed++; console.log('  ✓ ' + name); }
  else { failed++; console.error('  ✗ ' + name); }
}

const errors = [];
window.addEventListener('error', (e) => errors.push(e.message));

// The app boots empty now (no embedded seed). Preload the workbook fixture
// into localStorage so the suite still exercises a fully-populated document;
// capacity is enabled here because several tests exercise the capacity row.
{
  const fixture = require('./seed.fixture.js');
  const seeded = JSON.parse(JSON.stringify(fixture));
  seeded.meta.capacityEnabled = true;
  window.localStorage.setItem('headway-v1', JSON.stringify(seeded));
}

for (const f of ['js/core.js', 'js/excel.js', 'js/export-png.js', 'js/export-pptx.js', 'js/export-jira.js', 'js/jira.js', 'js/ai.js', 'js/app.js']) {
  try {
    window.eval(fs.readFileSync(path.join(ROOT, f), 'utf8'));
  } catch (e) {
    failed++;
    console.error('  ✗ ' + f + ' threw on load: ' + e.message + '\n' + (e.stack || '').split('\n').slice(0, 4).join('\n'));
  }
}

const doc = window.document;
const state = () => window.__headway.getState();
const click = (el) => el.dispatchEvent(new window.MouseEvent('click', { bubbles: true }));

// ---------------------------------------------------------------- boot
ok(errors.length === 0, 'no window errors during boot' + (errors.length ? ' — ' + errors.join('; ') : ''));

// ------------------------------------------------------------ start page
// a fresh launch lands on the start page; entering via Continue restores
// the localStorage session and reveals the editor
ok(doc.documentElement.dataset.theme === 'light' || doc.documentElement.dataset.theme === 'dark',
  'theme stamped on <html> (' + doc.documentElement.dataset.theme + ')');
ok(doc.body.classList.contains('start') && !doc.querySelector('#startPage').hidden,
  'boot shows the start page');
ok(doc.querySelector('#startBody [data-sp-new]') && doc.querySelector('#startBody [data-sp-opendlg]'),
  'start page offers New project and Open');
{
  click(doc.querySelector('#startBody [data-sp-settings]'));
  ok(!doc.querySelector('#modalHost').hidden &&
    doc.querySelectorAll('#modalHost [data-pref-theme]').length === 3,
    'start page settings modal offers the three themes');
  click(doc.querySelector('#modalHost [data-pref-theme="dark"]'));
  ok(doc.documentElement.dataset.theme === 'dark', 'picking Dark stamps data-theme=dark');
  click(doc.querySelector('#modalHost [data-pref-theme="system"]'));
  ok(window.localStorage.getItem('headway-theme-v1') === 'system', 'theme choice persists');
  click(doc.querySelector('#modalHost [data-m="x2"]'));
}
// ------------------------------------------------------ release notes
// the browser build has no version → no footer, no update button
ok(!doc.querySelector('#startBody [data-sp-notes]') && !doc.querySelector('#startBody [data-sp-update]'),
  'browser build shows neither the version button nor the update button');
{
  const md = fs.readFileSync(path.join(ROOT, 'CHANGELOG.md'), 'utf8');
  const { releaseNotesFor } = window.__headway;
  const notes = releaseNotesFor(md, '1.0.10');
  ok(/\*\*AI assistant\*\*/.test(notes) && !/Sprinting page/.test(notes),
    'releaseNotesFor picks exactly the 1.0.10 section');
  ok(releaseNotesFor(md, 'v1.0.9') === releaseNotesFor(md, '1.0.9') && /Sprinting page/.test(releaseNotesFor(md, '1.0.9')),
    'releaseNotesFor accepts a v-prefix and finds an inner section');
  ok(releaseNotesFor(md, '1.0.1') === '' && releaseNotesFor(md, '9.9.9') === '',
    'releaseNotesFor: "1.0.1" does not match "1.0.10"; unknown versions are empty');
  ok(releaseNotesFor('## 2.0.0\r\n- **A**: b\r\n## 1.0.0\r\n- c', '2.0.0') === '- **A**: b',
    'releaseNotesFor tolerates CRLF');

  // fake desktop version so the start page grows its footer + update button
  window.HeadwayDesktop = { appVersion: '1.0.10', updater: { state: 'idle', version: '', check() { this.checked = true; }, install() {} } };
  window.HeadwayApp.renderStartPage();
  const verBtn = doc.querySelector('#startBody [data-sp-notes]');
  ok(verBtn && /1\.0\.10/.test(verBtn.textContent), 'desktop start page footer is a version button');
  const upBtn = doc.querySelector('#startBody [data-sp-update]');
  ok(upBtn && /Check for updates/.test(upBtn.textContent) &&
    upBtn.previousElementSibling === doc.querySelector('#startBody [data-sp-settings]'),
    'update button sits right of Settings and offers a check');
  click(upBtn);
  ok(window.HeadwayDesktop.updater.checked === true, 'clicking Check for updates asks the updater');
  window.HeadwayDesktop.updater.state = 'ready';
  window.HeadwayDesktop.updater.version = '1.0.11';
  window.HeadwayApp.renderStartPage();
  const readyBtn = doc.querySelector('#startBody [data-sp-update]');
  ok(readyBtn && /Update to 1\.0\.11/.test(readyBtn.textContent) && readyBtn.classList.contains('sp-update-ready'),
    'a downloaded update turns the button into "Update to 1.0.11"');

  // once per version: first call shows, second is silent, new version shows again
  window.localStorage.removeItem('headway-notes-seen-v1');
  ok(window.__headway.maybeShowReleaseNotes(md) === true && !doc.querySelector('#modalHost').hidden &&
    /What.s new in Headway 1\.0\.10/.test(doc.querySelector('#modalHost h2').textContent) &&
    doc.querySelectorAll('#modalHost .notes-md li').length >= 4 &&
    doc.querySelector('#modalHost .notes-md b') && /AI assistant/.test(doc.querySelector('#modalHost .notes-md b').textContent),
    'first launch on a version opens What\'s new with rendered bullets and bold titles');
  click(doc.querySelector('#modalHost [data-m="ok"]'));
  ok(doc.querySelector('#modalHost').hidden, 'Got it closes the notes');
  ok(window.__headway.maybeShowReleaseNotes(md) === false && doc.querySelector('#modalHost').hidden,
    'the same version never shows the notes twice');
  ok(window.localStorage.getItem('headway-notes-seen-v1') === '1.0.10', 'seen version is remembered');
  window.HeadwayDesktop.appVersion = '1.0.9';
  ok(window.__headway.maybeShowReleaseNotes(md) === true && /1\.0\.9/.test(doc.querySelector('#modalHost h2').textContent),
    'a different version shows its own notes');
  click(doc.querySelector('#modalHost [data-m="x"]'));
  window.HeadwayDesktop.appVersion = '1.0.10';
  // a version with no section stays quiet but is still marked seen
  window.HeadwayDesktop.appVersion = '0.0.1';
  ok(window.__headway.maybeShowReleaseNotes(md) === true && doc.querySelector('#modalHost').hidden &&
    window.localStorage.getItem('headway-notes-seen-v1') === '0.0.1',
    'a version without notes shows nothing');
  window.HeadwayDesktop.appVersion = '1.0.10';

  // the footer button reopens the notes any time
  window.__headway.openReleaseNotes(md);
  ok(!doc.querySelector('#modalHost').hidden && /1\.0\.10/.test(doc.querySelector('#modalHost h2').textContent),
    'openReleaseNotes reopens the current version\'s notes');
  click(doc.querySelector('#modalHost [data-m="ok"]'));

  delete window.HeadwayDesktop;
  window.HeadwayApp.renderStartPage();
}

const contBtn = doc.querySelector('#startBody [data-sp-continue]');
ok(!!contBtn, 'browser session offers Continue where you left off');
click(contBtn);
ok(!doc.body.classList.contains('start') && doc.querySelector('#startPage').hidden,
  'Continue enters the editor');

// -------------------------------------------- author required on first change
{
  ok(!window.localStorage.getItem('headway-user-v1'), 'fresh session starts with no author name');
  const t = doc.querySelector('#docTitle');
  const origTitle = state().meta.title;
  ok(t.hasAttribute('readonly'), 'the title rests readonly');
  click(t);
  ok(!t.hasAttribute('readonly'), 'clicking the title starts a rename in place (no pencil)');
  ok(!doc.querySelector('#titleEdit'), 'the pencil button is gone');
  t.value = 'Renamed Roadmap.xlsx';
  t.dispatchEvent(new window.Event('change', { bubbles: true }));
  ok(state().meta.title === 'Renamed Roadmap', 'renaming strips a typed .xlsx extension');
  ok(window.__headway.saveFileName() === 'Renamed Roadmap.xlsx', 'the title maps 1:1 onto the file name');
  const mh = doc.querySelector('#modalHost');
  ok(!mh.hidden && /Who’s editing\?/.test(mh.textContent), 'the first saved change asks for your name');
  t.value = origTitle; // put the suite's document title back
  t.dispatchEvent(new window.Event('change', { bubbles: true }));
  const nameIn = doc.querySelector('#vhNameIn');
  nameIn.value = 'Test User';
  nameIn.dispatchEvent(new window.Event('input', { bubbles: true }));
  click(doc.querySelector('#vhNameSave'));
  ok(mh.hidden, 'saving the name closes the prompt');
  ok(window.localStorage.getItem('headway-user-v1') === 'Test User', 'the name persists on this machine');
  const hist = state().history;
  ok(hist.length && hist[hist.length - 1].u === 'Test User', 'the anonymous change is stamped with the new author');
}

// ------------------------------------------------------------ AI assistant
{
  const AI = window.HeadwayAI;
  ok(!!AI && typeof AI.send === 'function', 'js/ai.js registers window.HeadwayAI');
  const drawer = doc.querySelector('#aiDrawer');
  ok(drawer && drawer.hidden, 'the assistant drawer starts closed');
  click(doc.querySelector('#btnAI'));
  ok(!drawer.hidden && doc.body.classList.contains('ai-open'), 'the toolbar AI button opens the drawer');
  ok(doc.querySelector('#aiDrawer #aiInput') && doc.querySelector('#aiDrawer #aiCompose #aiEffortSel') && doc.querySelector('#aiDrawer #aiCompose #aiModelSel') && doc.querySelector('#aiDrawer textarea#aiInput'),
    'the composer is a textarea with the model and effort pickers beneath it');
  ok(/Open AI settings/.test(drawer.textContent), 'an unconfigured provider points at the settings');
  click(doc.querySelector('#aiDrawer #aiClose'));
  ok(drawer.hidden, 'the close button hides the drawer');
  // settings tab renders the provider fields
  window.HeadwayApp.ai.openSettings();
  ok(doc.querySelector('#aiSettingsCard #aiBase') && doc.querySelector('#aiSettingsCard [data-aiprov="claude"]'),
    'Setup → AI assistant shows the gateway fields and the provider switch');
  click(doc.querySelector('#aiSettingsCard [data-aiprov="claude"]'));
  ok(AI.loadSettings().provider === 'claude' && !doc.querySelector('#aiSettingsCard #aiClaude').hidden,
    'picking the Claude provider persists and reveals its fields');
  click(doc.querySelector('#aiSettingsCard [data-aiprov="litellm"]'));
  ok(!doc.querySelector('#aiSettingsCard [data-aieffort]'), 'the settings tab carries no effort control (it lives in the drawer)');
  {
    const effSel = doc.querySelector('#aiDrawer #aiEffortSel');
    effSel.value = 'high';
    effSel.dispatchEvent(new window.Event('change', { bubbles: true }));
    ok(AI.loadSettings().effort === 'high', 'effort persists from the drawer select');
    ok([...doc.querySelectorAll('#aiDrawer #aiModelSel option')].every(o => o.textContent.indexOf('/') === -1),
      'model labels in the drawer drop provider prefixes');
  }
  // tools run against the live app bridge and land in history as "· AI"
  const before = state().items.length;
  const res = AI.runTool('add_items', { items: [{ feature: 'AI-made feature', start: state().meta.timelineStart, durDays: 5 }] }, window.HeadwayApp);
  ok(res.created.length === 1 && state().items.length === before + 1, 'add_items creates a feature through the app');
  const h = state().history;
  ok(h[h.length - 1].u === 'Test User · AI', 'AI edits are attributed to "<name> · AI" in Version history');
  ok(AI.runTool('navigate', { view: 'planning', num: res.created[0].num }, window.HeadwayApp).selected === true, 'navigate selects the new feature');
  AI.runTool('update_items', { updates: [{ num: res.created[0].num, delete: true }] }, window.HeadwayApp);
  ok(state().items.length === before, 'update_items can delete it again');
  window.HeadwayApp.ai.setView('planning');
}

// ---------------------------------------------------------------- Setup → Views (per-project tab switch)
{
  const openApps = () => { click(doc.querySelector('#btnSetup')); click(doc.querySelector('#setupView [data-sutab="views"]')); };
  openApps();
  const tabs = [...doc.querySelectorAll('#setupView .su-tab')].map(b => b.dataset.sutab);
  ok(tabs.indexOf('views') === tabs.indexOf('columns') + 1, 'Views sits right after the project sections');
  const boxes = doc.querySelectorAll('#setupView [data-suapp]');
  ok([...boxes].map(b => b.dataset.suapp).join() === 'scoping,prio,reports' && [...boxes].every(b => b.checked),
    'Scoping, Prioritizing and Reporting have switches, on by default');
  ok(!doc.querySelector('#setupView [data-suapp="planning"]') && /Always on/.test(doc.querySelector('#setupView [data-suview="planning"]').textContent),
    'Planning cannot be switched off');
  ok(/Follows Sprints/.test(doc.querySelector('#setupView [data-suview="sprints"]').textContent) &&
    /Follows Budgeting/.test(doc.querySelector('#setupView [data-suview="budget"]').textContent), 'Sprinting and Budgeting follow their sections');
  const sc = doc.querySelector('#setupView [data-suapp="scoping"]');
  sc.checked = false; sc.dispatchEvent(new window.Event('change', { bubbles: true }));
  ok(state().meta.apps.scoping === false, 'unchecking Scoping commits to the document');
  ok(doc.querySelector('#viewTabs [data-view="scoping"]').hidden && !doc.querySelector('#viewTabs [data-view="planning"]').hidden,
    'the Scoping tab hides from the header while the others stay');
  ok(window.HeadwayApp.ai.setView('scoping') === false && doc.body.dataset.view !== 'scoping', 'the AI cannot open a hidden app');
  window.HeadwayApp.ai.setView('sprints');
  ok(doc.body.dataset.view === 'sprints', 'Sprinting still opens while on');
  window.dispatchEvent(new window.KeyboardEvent('keydown', { key: 'z', metaKey: true, bubbles: true }));
  ok(state().meta.apps.scoping === true && !doc.querySelector('#viewTabs [data-view="scoping"]').hidden,
    'undo restores the app and its tab');
  window.HeadwayApp.ai.setView('planning');
}
ok(doc.querySelectorAll('#rows .row.band').length === 6, 'six phase bands rendered');
const itemRows = doc.querySelectorAll('#rows .row.item').length;
ok(itemRows > 100, 'item rows rendered (' + itemRows + ')');
const visibleSched = state().items.filter(i => i.startDay != null &&
  !state().phases.find(p => p.id === i.phaseId).collapsed).length;
ok(doc.querySelectorAll('#rows .bar').length === visibleSched,
  'bars rendered for every visible scheduled item (' + doc.querySelectorAll('#rows .bar').length + ')');
ok(doc.querySelectorAll('#hdrCapRows .hdr-cap').length === 1, 'one summed header capacity row');
ok(doc.querySelectorAll('#hdrCapRows .hdr-cap:first-child .cap-cell').length === 48, 'capacity strip has 48 week cells');
ok(state() && state().items.length > 100, 'debug state handle live (' + state().items.length + ' items)');
ok(doc.querySelector('#resPanel') !== null && doc.querySelector('#resGrid') !== null, 'resources panel present');
ok(/^Capacity \((people|points)\)$/.test(doc.querySelector('#hdrCapRows .hdr-cap .cap-row-lab').textContent),
  'the capacity row is labelled Capacity, with its unit');
ok(!doc.querySelector('#hdrCapRows [data-captype]'), 'the row belongs to no single type');
ok(doc.querySelector('#hdrCapRows .dd-btn') === null, 'capacity is role-agnostic: no role filter dropdown');
ok(doc.querySelectorAll('#hdrCapRows .hdr-cap:first-child .cap-cell').length === 48, 'capacity row spans all weeks');
ok(doc.querySelector('.hdr-legend .hdr-left.corner #hlCols') !== null,
  'the column legend sits on its own header line below the capacity rows');
ok(doc.querySelector('.hdr-legend #leftRz') !== null, 'and keeps the left-pane resize handle');
ok(doc.querySelectorAll('#rows .bar .port').length === visibleSched * 2, 'link ports rendered on bars');
ok(doc.querySelectorAll('#rows .bar .b-label').length === visibleSched,
  'every bar carries a label (inside or spilled right)');
ok(doc.querySelectorAll('#rows .bar .b-label.out').length > 0,
  'overflowing labels spill to the right of the bar');
ok(doc.querySelectorAll('#hdrSprints .sprint-cell .sp-date').length > 20, 'sprint header shows dates as primary label');
ok(doc.querySelector('#leftRzLine') !== null, 'full-height left-pane resize line present');
{
  const phCells = doc.querySelectorAll('#hdrPhases .ph-cell');
  const withScheduled = state().phases.filter(p =>
    state().items.some(i => i.phaseId === p.id && i.startDay != null)).length;
  ok(phCells.length === withScheduled,
    'header phase lane shows one span per phase with scheduled items (' + phCells.length + ')');
  ok(!phCells[0].getAttribute('title'), 'phase spans carry no native title');
  phCells[0].dispatchEvent(new window.MouseEvent('mouseover', { bubbles: true, clientX: 300, clientY: 20 }));
  const tip = doc.querySelector('#phTip');
  ok(tip && !tip.hidden && /→/.test(tip.querySelector('.pht-range').textContent),
    'hovering a phase span shows the tooltip with a date range');
  phCells[0].dispatchEvent(new window.MouseEvent('mouseout', { bubbles: true }));
  ok(tip.hidden, 'tooltip hides on mouseout');
  click(phCells[0]);
  ok(!doc.querySelector('#modalHost').hidden, 'clicking a header phase span opens the phase editor');
  click(doc.querySelector('#modalHost [data-m=cancel], #modalHost [data-m=x]'));
}
ok(!doc.querySelector('#rows .row.item .r-ico'), 'item rows carry no standalone epic-icon slot');
ok(!doc.querySelector('#rows .row.item .r-epic'), 'item rows carry no epic tag beside the title (the Epic column does)');

// ---------------------------------------------------------------- menus
click(doc.querySelector('[data-menu="file"]'));
ok(!doc.querySelector('#popover').hidden && doc.querySelectorAll('#popover .menu-list button').length >= 5, 'File menu opens with items');
ok(Array.from(doc.querySelectorAll('#popover .menu-list button')).some(b => /Download template/.test(b.textContent)),
  'File menu offers Download template');
ok(Array.from(doc.querySelectorAll('#popover .menu-list button')).some(b => /Sync with Jira/.test(b.textContent)),
  'File menu offers Sync with Jira (the CSV export stays in the Export dialog)');
{
  const tpl = window.__headway.templateState();
  ok(tpl.items.length === 1 && /Example/.test(tpl.items[0].feature) &&
    tpl.items[0].stories.length === 1 && tpl.items[0].stories[0].startDay != null &&
    tpl.team.length === 0,
    'template state is empty but for one example feature (with a story)');
}
click(doc.querySelector('[data-menu="edit"]'));
ok(Array.from(doc.querySelectorAll('#popover .menu-list button')).some(b => /Undo/.test(b.textContent)), 'Edit menu holds Undo');
doc.querySelector('#popover').hidden = true;

// ---------------------------------------------------------------- select + panel
const firstItem = doc.querySelector('#rows .row.item');
const itId = firstItem.getAttribute('data-id');
click(firstItem.querySelector('.r-num'));
ok(!doc.querySelector('#panel').hidden, 'clicking a row opens the detail panel');
ok(doc.querySelector('#panel .p-name').value.length > 0, 'panel shows the feature name');
{
  const rowOf = () => doc.querySelector('#rows .row.item[data-id="' + itId + '"]');
  const nameEl = rowOf().querySelector('.r-name');
  ok(nameEl && nameEl.tagName !== 'INPUT' && nameEl.textContent.length > 0, 'row title is text, not an input');
  click(nameEl);
  ok(!rowOf().querySelector('input'), 'a single click on the title does not start an edit');
  rowOf().querySelector('.r-name').dispatchEvent(new window.MouseEvent('dblclick', { bubbles: true }));
  const nameInp = rowOf().querySelector('input.r-name');
  ok(!!nameInp, 'double-click turns the row title into an input');
  nameInp.value = 'Renamed inline';
  nameInp.dispatchEvent(new window.KeyboardEvent('keydown', { key: 'Enter', bubbles: true }));
  ok(state().items.find(i => i.id === itId).feature === 'Renamed inline', 'Enter commits the row title rename');
  // the row menu starts the same edit
  rowOf().dispatchEvent(new window.MouseEvent('contextmenu', { bubbles: true, cancelable: true, clientX: 120, clientY: 120 }));
  const rnBtn = [...doc.querySelectorAll('#popover .menu-list button')].find(b => /Rename/.test(b.textContent));
  ok(!!rnBtn, 'the row context menu offers Rename…');
  click(rnBtn);
  const ed2 = rowOf().querySelector('input.r-name');
  ok(!!ed2, 'Rename… starts the same inline edit');
  ed2.dispatchEvent(new window.KeyboardEvent('keydown', { key: 'Escape', bubbles: true }));
  ok(state().items.find(i => i.id === itId).feature === 'Renamed inline', 'Escape leaves the title alone');
  click(rowOf().querySelector('.r-num')); // Escape cleared the selection — put it back
}
ok(doc.querySelector('#panel .wz-ed[data-f="col:description"]') !== null, 'panel has a description field');
{
  const jk = doc.querySelector('#panel input[data-f="jiraKey"]');
  ok(!!jk, 'panel has a Jira key input');
  jk.value = ' hw-7 ';
  jk.dispatchEvent(new window.Event('change', { bubbles: true }));
  ok(state().items.find(i => i.id === itId).jiraKey === 'HW-7', 'Jira key commits normalized');
  ok(doc.querySelector('#panel input[data-f="jiraKey"]').value === 'HW-7', 'panel re-renders the normalized key');
}
ok(doc.querySelector('#panel .p-sec[data-sec="fields"]').classList.contains('open'), 'Fields section is open by default');
ok(doc.querySelector('#panel .p-actions') === null, 'footer Duplicate/Delete buttons are gone');
ok(doc.querySelector('#panel .p-more') === null, 'the … actions button is gone from the panel');
{
  doc.querySelector('#panel .p-top').dispatchEvent(
    new window.MouseEvent('contextmenu', { bubbles: true, cancelable: true, clientX: 500, clientY: 100 }));
  ok(!doc.querySelector('#popover').hidden &&
    Array.from(doc.querySelectorAll('#popover .menu-list button')).some(b => /Duplicate/.test(b.textContent)),
    'right-clicking the panel opens the item actions menu');
  doc.querySelector('#popover').hidden = true;
}
ok(doc.querySelector('#panel [data-dd=epic]') !== null, 'epic is a dropdown button');
click(doc.querySelector('#panel [data-dd=epic]'));
ok(!doc.querySelector('#popover').hidden && doc.querySelectorAll('#popover .menu-list [data-mi]').length > 2, 'epic dropdown opens the shared list UI');
ok(doc.querySelectorAll('#popover .menu-list .mi-edit').length > 0, 'epic options carry an edit affordance');
doc.querySelector('#popover').hidden = true;
ok(doc.querySelector('#panel [data-f=riskSize]') !== null, 'risk size segment present');
ok(doc.querySelector('#panel [data-f=allabove]') === null, '"all items above" checkbox is gone');

// description commit (rich editor — commits on blur)
{
  const descEd0 = doc.querySelector('#panel .wz-ed[data-f="col:description"]');
  descEd0.focus();
  descEd0.innerHTML = 'A crisp description';
  descEd0.dispatchEvent(new window.FocusEvent('focusout', { bubbles: true }));
  ok(state().items.find(i => i.id === itId).description === 'A crisp description', 'description saves');
}

// Backspace while typing in the rich description edits text — it must not
// trigger the delete-item shortcut (Mac delete key sends "Backspace")
{
  const before = state().items.length;
  const descEd = doc.querySelector('#panel .wz-ed[data-f="col:description"]');
  ok(!!descEd && descEd.getAttribute('contenteditable') === 'true', 'panel description is a rich contenteditable editor');
  descEd.focus();
  window.dispatchEvent(new window.KeyboardEvent('keydown', { key: 'Backspace', bubbles: true }));
  ok(doc.querySelector('#modalHost').hidden, 'Backspace in the description does not open the delete confirm');
  ok(state().items.length === before, 'Backspace in the description does not delete the item');
  descEd.blur();
}

// ---------------------------------------------------------------- sticky bands
// phase / workstream / epic band rows freeze under the header while scrolling;
// the offsets come from a measured --hdr-h (header height varies per view)
{
  ok((doc.documentElement.style.getPropertyValue('--hdr-h') || '').endsWith('px'),
    'render measures the header and syncs --hdr-h');
  const css = fs.readFileSync(path.join(ROOT, 'css/app.css'), 'utf8');
  const decl = (sel) => {
    const m = css.match(new RegExp(sel.replace(/[.\\]/g, '\\$&') + '\\s*{([^}]*)}', 'g')) || [];
    return m.join(' ');
  };
  ok(/body\[data-view="scoping"\] \.hdr-legend,/.test(css),
    'scoping hides the column-legend line (it draws its own column header)');
  ok(/position:\s*sticky/.test(decl('.row.band')) && /top:\s*var\(--hdr-h\)/.test(decl('.row.band')),
    'phase bands are sticky below the header');
  // second-level group bands pin under the phase band; nested epics one lower
  ok(/position:\s*sticky/.test(decl('.row.eband')) &&
     /top:\s*calc\(var\(--hdr-h\)\s*\+\s*var\(--band-real-h/.test(decl('.row.eband')),
    'epic/workstream bands are sticky under the phase band');
  ok(/\.row\.eband\.sub\s*{[^}]*var\(--eband-real-h/.test(css),
    'nested epic bands stack one band lower');
  ok(/\.row\.eband \.row-lane\s*{[^}]*background-color:\s*var\(--lvl-epic\)/.test(css) &&
     /\.row\.eband \.row-left\s*{[^}]*background:\s*var\(--lvl-epic\)/.test(css),
    'group band cells are opaque so rows do not show through when pinned');
  ok(/\.row\.eband \.row-lane\s*{[^}]*repeating-linear-gradient\([^)]*\)[^}]*var\(--sprint-px/.test(css),
    'the pinned band lane redraws the sprint grid on the sprint pitch');
  // one period per tile, or a non-zero --sprint-off leaves the last stripe off-lattice
  ok(/\.row\.eband \.row-lane\s*{[^}]*background-size:\s*var\(--sprint-px[^}]*}/.test(css),
    'the band lane tiles exactly one sprint period');
  // Scoping has no time axis (#bgcols is empty there) — no grid on the band
  ok(/body\[data-view="scoping"\] \.row\.eband \.row-lane\s*{[^}]*background-image:\s*none/.test(css) &&
     /body\[data-view="scoping"\] \.row\.eband \.row-lane\s*{[^}]*background-color:\s*var\(--surface\)/.test(css),
    'Scoping band lanes drop the sprint grid and match the Scoping item lanes');
}

// ---------------------------------------------------------------- chips
click(doc.querySelector('#rows .row.item .r-size'));
ok(!doc.querySelector('#popover').hidden, 'size chip opens a dropdown');
{
  const xl = Array.from(doc.querySelectorAll('#popover .menu-list button')).find(b => /^XL/.test(b.textContent.trim()));
  click(xl);
  ok(state().items.find(i => i.id === itId).size === 'XL', 'picking a size commits (XL)');
}
// no risk scheme yet: planning rows carry no risk chip; scoping keeps its chip with None/L/M/H options
ok(!!doc.querySelector('#rows .row.item [data-act="risk"]') === window.RM.riskEnabled(state()),
  'planning rows carry a risk chip exactly when a risk scheme is on');
click(doc.querySelector('#viewTabs [data-view="scoping"]'));
click(doc.querySelector('#rows .row.item[data-id="' + itId + '"] .r-risk'));
ok(!doc.querySelector('#popover').hidden, 'scoping risk chip opens a dropdown');
{
  const labels = Array.from(doc.querySelectorAll('#popover .menu-list button')).map(b => b.textContent.trim());
  ok(labels.includes('None') && ['Low', 'Medium', 'High'].every(v => labels.includes(v)) && !labels.some(l => /^XL/.test(l)),
    'risk options are None / Low / Medium / High (glyph + label, no t-shirt sizes)');
  click(Array.from(doc.querySelectorAll('#popover .menu-list button')).find(b => b.textContent.trim() === 'Low'));
  ok(state().items.find(i => i.id === itId).risk === 'L', 'picking a risk commits (L)');
}
click(doc.querySelector('#viewTabs [data-view="planning"]'));

// headcount is gone as an item field — availability is roster-driven now
ok(!doc.querySelector('#rows .row.item .r-hc'), 'no headcount chip on planning rows');
ok(!doc.querySelector('#panel [data-f=headcount]'), 'no headcount field in the panel');

// undo via keyboard (toolbar buttons moved into the Edit menu)
{
  const before = state().items.find(i => i.id === itId).risk;
  window.dispatchEvent(new window.KeyboardEvent('keydown', { key: 'z', metaKey: true, bubbles: true }));
  ok(state().items.find(i => i.id === itId).risk !== before || before == null,
    'undo (⌘Z) reverts the last commit');
  window.dispatchEvent(new window.KeyboardEvent('keydown', { key: 'z', metaKey: true, shiftKey: true, bubbles: true }));
}

// ---------------------------------------------------------------- risk is metadata only
const scheduled = state().items.find(i => i.startDay != null && !i.locked);
window.__headway.getState(); // no-op, keep linear
ok(state().items.every(i => (i.riskDays || 0) === 0), 'no schedule padding from risk anywhere');
ok(state().items.some(i => i.risk), 'risk t-shirts survive as metadata');

// ---------------------------------------------------------------- dep search by name
const target = state().items.find(i => i.id !== itId && i.feature && i.feature.length > 6);
const depInput = doc.querySelector('#panel [data-f=depsearch]');
depInput.value = target.feature.slice(0, 6);
depInput.dispatchEvent(new window.Event('input', { bubbles: true }));
const sugBtn = doc.querySelector('#panel .dep-sug [data-addep]');
ok(!!sugBtn, 'dep search suggests matches by name');
if (sugBtn) {
  const num = parseInt(sugBtn.getAttribute('data-addep'), 10);
  click(sugBtn);
  ok(state().items.find(i => i.id === itId).deps.indexOf(num) !== -1, 'clicking a suggestion adds the dependency');
}

// ---------------------------------------------------------------- holidays are day-granular
ok(Array.isArray(state().meta.holidays) && state().meta.blackoutWeeks === undefined,
  'holidays are individual dates (' + state().meta.holidays.length + '); blackoutWeeks migrated away');
ok(doc.querySelectorAll('#bgcols .bg-blackout').length >= 1, 'holiday segments drawn on the timeline');
const boBefore = state().meta.holidays.length;
click(doc.querySelector('#hdrCapRows [data-w="3"]'));
ok(state().meta.holidays.length === boBefore + 5, 'clicking a capacity cell adds that week\'s five holiday days');
click(doc.querySelector('#hdrCapRows [data-w="3"]'));
ok(state().meta.holidays.length === boBefore, 'clicking again removes them');

// ---------------------------------------------------------------- context menu
const cmRow = doc.querySelectorAll('#rows .row.item')[3];
cmRow.dispatchEvent(new window.MouseEvent('contextmenu', { bubbles: true, clientX: 200, clientY: 200 }));
{
  const labels = Array.from(doc.querySelectorAll('#popover .menu-list button')).map(b => b.textContent);
  ok(!doc.querySelector('#popover').hidden && labels.some(l => /Insert feature below/.test(l)),
    'right-click opens a row context menu');
  ok(labels.some(l => /Insert feature above/.test(l)), 'context menu also offers insert above');
  ok(labels.some(l => /Move to phase/.test(l)) && labels.some(l => /Set epic/.test(l)),
    'row context menu offers move-to-phase and set-epic');
  ok(labels.some(l => /^Lock$|^Unlock$/.test(l)) && labels.some(l => /Mark as done|Unmark as done/.test(l)),
    'row context menu offers lock and done toggles');
  const cmIt = state().items.find(i => i.id === cmRow.dataset.id);
  ok(labels.some(l => /Unschedule/.test(l)) === (cmIt.startDay != null),
    'Unschedule shown only for scheduled items');
}
doc.querySelector('#popover').hidden = true;
// done toggle round-trips through the menu
{
  const before = !!state().items.find(i => i.id === cmRow.dataset.id).done;
  cmRow.dispatchEvent(new window.MouseEvent('contextmenu', { bubbles: true, clientX: 200, clientY: 200 }));
  click(Array.from(doc.querySelectorAll('#popover .menu-list button'))
    .find(b => /Mark as done|Unmark as done/.test(b.textContent)));
  ok(!!state().items.find(i => i.id === cmRow.dataset.id).done === !before,
    'context-menu done toggle commits');
  window.dispatchEvent(new window.KeyboardEvent('keydown', { key: 'z', metaKey: true, bubbles: true }));
}
{
  const band = doc.querySelector('#rows .row.band');
  band.dispatchEvent(new window.MouseEvent('contextmenu', { bubbles: true, clientX: 180, clientY: 120 }));
  const labels = Array.from(doc.querySelectorAll('#popover .menu-list button')).map(b => b.textContent);
  ok(labels.some(l => /Delete phase/.test(l)), 'phase band context menu offers delete');
  doc.querySelector('#popover').hidden = true;
}

// ---------------------------------------------------------------- setup view
const suTab = (k) => {
  // entering Setup first keeps the tab click live (a hidden #setupView keeps
  // stale DOM from its last visit)
  if (doc.body.dataset.view !== 'setup') click(doc.querySelector('#btnSetup'));
  click(doc.querySelector('#setupView [data-sutab="' + k + '"]'));
};
// Setup → Sizing, priority & risk: one select per field × level
const schemeSel = (what, level) => doc.querySelector('#setupView select[data-suscheme-kind="' + what + '"][data-kind="' + level + '"]');
const schemeOpts = (what, level) => [...schemeSel(what, level).options].map(o => o.value);
const pickScheme = (what, level, v) => {
  const sel = schemeSel(what, level);
  sel.value = v;
  sel.dispatchEvent(new window.Event('change', { bubbles: true }));
};
click(doc.querySelector('#resManage'));
ok(doc.body.dataset.view === 'setup', 'resources "manage" jumps to the Setup view');
ok(doc.querySelector('#setupView [data-sutab="team"]').classList.contains('on'),
  'resources "manage" lands on the Team tab');
{
  const keys = [...doc.querySelectorAll('#setupView .su-tab')].map(b => b.dataset.sutab);
  ok(keys.join() === 'project,sprints,org,est,budget,team,scheduling,columns,views,appearance,prefs,ai,jira',
    'Setup rail: eight sections, Views, then Personal');
  ok(doc.querySelectorAll('#setupView .su-rail-hd').length === 1 &&
    /this computer only/.test(doc.querySelector('#setupView .su-rail-hd').textContent), 'one group heading: Personal');
  ok([...doc.querySelectorAll('#setupView .su-tab')].every(b => b.querySelector('i[data-lucide]')), 'every rail item has an icon');
  ok(/Jira Integration/.test(doc.querySelector('#setupView [data-sutab="jira"]').textContent), 'Jira renamed');
  ok(!!doc.querySelector('#setupView [data-sutab="sprints"] .su-pill'), 'Sprints shows its on/off pill');
  window.HeadwayApp.openSetup('capacity');
  ok(doc.querySelector('#setupView [data-sutab="scheduling"]').classList.contains('on'), 'old key capacity opens Scheduling');
  ok(!/Capacity planning/.test(doc.querySelector('#setupView').textContent), 'no "Capacity planning" copy left');
  suTab('sprints');
  ok([...doc.querySelectorAll('#setupView [data-suwps]')].map(b => b.textContent).join() === 'Off,1 week,2 weeks,3 weeks,4 weeks',
    'Sprints offers Off and 1–4 weeks');
}
suTab('budget');
ok(/Roles/.test(doc.querySelector('#setupView .su-card h2').textContent), 'Budgeting holds the roles and rate card');
ok(!!doc.querySelector('#setupView [data-rcrate]') && !!doc.querySelector('#setupView [data-rccost]'),
  'rate card inputs per role');
suTab('project');
ok(!!doc.querySelector('#suWeekHours') && doc.querySelectorAll('#setupView [data-suwday]').length === 7 &&
  !!doc.querySelector('#suWeekStart'),
  'work week card offers full-time hours, Sun-Sat day checkboxes and a first-day select');
{
  // Mon-Fri checked by default; up to all 7 days can be selected
  const wdBoxes = Array.from(doc.querySelectorAll('#setupView [data-suwday]'));
  ok(wdBoxes.filter(b => b.checked).length === 5 && wdBoxes.filter(b => b.disabled).length === 0,
    'five working days checked, none disabled (7-day weeks allowed)');
  // unchecking Friday commits a 4-day week and re-encodes the day space
  const fri = wdBoxes.find(b => b.dataset.suwday === '5');
  fri.checked = false;
  fri.dispatchEvent(new window.Event('change', { bubbles: true }));
  ok(state().meta.workDays.join(',') === '1,2,3,4', 'unchecking Friday leaves Mon-Thu');
  ok(window.RM.slotsOf(state().meta) === 4, 'the index week now has 4 slots');
  window.dispatchEvent(new window.KeyboardEvent('keydown', { key: 'z', metaKey: true, bubbles: true }));
  ok(window.RM.slotsOf(state().meta) === 5, 'undo restores the 5-slot week exactly');
  // a 6th day can be checked (Saturday) — and undone
  const sat = doc.querySelector('#setupView [data-suwday="6"]');
  sat.checked = true;
  sat.dispatchEvent(new window.Event('change', { bubbles: true }));
  ok(window.RM.slotsOf(state().meta) === 6, 'checking Saturday makes a 6-slot week');
  window.dispatchEvent(new window.KeyboardEvent('keydown', { key: 'z', metaKey: true, bubbles: true }));
  ok(window.RM.slotsOf(state().meta) === 5, 'undo restores Mon-Fri');
}
{
  // rate card commit + inheritance shows up in core helpers
  suTab('budget');
  const rateInp = doc.querySelector('#setupView [data-rcrate]');
  const role0 = rateInp.dataset.rcrate;
  rateInp.value = '175';
  rateInp.dispatchEvent(new window.Event('change', { bubbles: true }));
  ok(state().meta.rateCard[role0] && state().meta.rateCard[role0].rate === 175,
    'rate card edit commits to meta.rateCard');
  ok(window.RM.memberRate(state(), { type: role0, rate: 0, cost: 0 }) === 175,
    'a person with no override inherits the card rate');
  window.dispatchEvent(new window.KeyboardEvent('keydown', { key: 'z', metaKey: true, bubbles: true }));
}
{
  // work week commit
  suTab('project');
  const wh = doc.querySelector('#suWeekHours');
  wh.value = '32';
  wh.dispatchEvent(new window.Event('change', { bubbles: true }));
  ok(state().meta.weekHours === 32, 'full-time hours commit');
  const wsSel = doc.querySelector('#suWeekStart');
  wsSel.value = '0';
  wsSel.dispatchEvent(new window.Event('change', { bubbles: true }));
  ok(state().meta.weekStart === 0, 'first day of week commits');
  window.dispatchEvent(new window.KeyboardEvent('keydown', { key: 'z', metaKey: true, bubbles: true }));
  window.dispatchEvent(new window.KeyboardEvent('keydown', { key: 'z', metaKey: true, bubbles: true }));
  window.dispatchEvent(new window.KeyboardEvent('keydown', { key: 'z', metaKey: true, bubbles: true }));
}
suTab('appearance');
ok(doc.querySelectorAll('#setupView [data-pref-theme]').length === 3,
  'Personal → Appearance offers the three themes');
suTab('prefs');
ok(!!doc.querySelector('#setupView [data-pref="crit"]') &&
  doc.querySelectorAll('#setupView [data-pref-snap]').length === 6,
  'Personal → Preferences holds the view options');
{
  const cb = doc.querySelector('#setupView [data-pref="crit"]');
  cb.checked = false;
  cb.dispatchEvent(new window.Event('change', { bubbles: true }));
  ok(window.localStorage.getItem('headway-ui-v1').includes('"showCrit":false'),
    'unticking critical path persists to the UI prefs');
  const cb2 = doc.querySelector('#setupView [data-pref="crit"]');
  cb2.checked = true;
  cb2.dispatchEvent(new window.Event('change', { bubbles: true }));
}
suTab('project');
ok(!!doc.querySelector('#suStart') && !!doc.querySelector('#suEnd'), 'timeline start/end editable in setup');
{
  const end = doc.querySelector('#suEnd');
  const startIso = state().meta.timelineStart;
  end.value = '2027-01-15';
  end.dispatchEvent(new window.Event('change', { bubbles: true }));
  ok(state().meta.numWeeks === 26 || Math.abs(state().meta.numWeeks - 26) <= 1,
    'end date drives numWeeks (' + state().meta.numWeeks + ' from ' + startIso + ')');
  window.dispatchEvent(new window.KeyboardEvent('keydown', { key: 'z', metaKey: true, bubbles: true }));
}
ok(doc.querySelectorAll('#setupView [data-suholrm]').length === state().meta.holidayRanges.length &&
  state().meta.holidayRanges.length > 0,
  'holidays listed as a removable named-range table (Project section)');
{
  // add a named range and remove it again
  const before = state().meta.holidayRanges.length;
  doc.querySelector('#suHolName').value = 'Offsite';
  doc.querySelector('#suHolStart').value = '2026-10-07';
  doc.querySelector('#suHolEnd').value = '2026-10-08';
  click(doc.querySelector('#suHolAddBtn'));
  ok(state().meta.holidayRanges.length === before + 1 &&
    state().meta.holidays.indexOf('2026-10-07') !== -1 &&
    state().meta.holidays.indexOf('2026-10-08') !== -1,
    'adding a named range expands into holiday dates');
  const idx = state().meta.holidayRanges.findIndex(r => r.name === 'Offsite');
  click(doc.querySelector('#setupView [data-suholrm="' + idx + '"]'));
  ok(state().meta.holidayRanges.length === before &&
    state().meta.holidays.indexOf('2026-10-07') === -1,
    'removing the range removes its dates');
}
suTab('budget');
{
  const inp = doc.querySelector('#suTypeAdd');
  inp.value = 'Data Scientist';
  click(doc.querySelector('#suTypeAddBtn'));
  ok(state().teamTypes.indexOf('Data Scientist') !== -1, 'setup adds a team type');
}
suTab('org');
ok(doc.querySelectorAll('#setupView [data-suphedit]').length === state().phases.length, 'phases listed with edit controls');
suTab('org');
ok(doc.querySelectorAll('#setupView [data-suwsedit]').length > 0, 'workstreams listed with edit controls');

// picking "Default gray" for a workstream with a seeded default must SURVIVE
// a reload — the choice is stored explicitly, not deleted (regression: the
// known-default re-seeded on load and reverted the color)
{
  const editBtn = Array.from(doc.querySelectorAll('#setupView [data-suwsedit]'))
    .find(b => b.dataset.suwsedit === 'Product') || doc.querySelector('#setupView [data-suwsedit]');
  const wsName = editBtn.dataset.suwsedit;
  click(editBtn);
  click(doc.querySelector('.swatch[data-esw="neutral"]'));
  click(doc.querySelector('#wsSave'));
  ok(state().wsColors[wsName] === 'neutral', 'Default gray stored explicitly for ' + wsName);
  const reloaded = window.RM.normalizeState(JSON.parse(JSON.stringify(state())));
  ok(window.RM.colorForWs(reloaded, wsName) === window.RM.PALETTE.neutral,
    'color survives a reload (normalize does not re-seed the default)');
}
ok(doc.querySelectorAll('#setupView .su-grip').length ===
  doc.querySelectorAll('#setupView [data-sulist] .su-row').length, 'every reorderable row has a drag grip');
// drag the first phase's grip to the bottom of its list (jsdom rects are all
// zero, so a large clientY resolves to "after the last row")
{
  suTab('org');
  const firstId = state().phases[0].id;
  const grip = doc.querySelector('#setupView [data-sulist="phase"] .su-row .su-grip');
  grip.dispatchEvent(new window.MouseEvent('pointerdown', { bubbles: true }));
  window.dispatchEvent(new window.MouseEvent('pointermove', { clientY: 999 }));
  window.dispatchEvent(new window.MouseEvent('pointerup'));
  ok(state().phases[state().phases.length - 1].id === firstId, 'dragging a phase grip reorders phases');
  window.dispatchEvent(new window.KeyboardEvent('keydown', { key: 'z', metaKey: true, bubbles: true }));
  ok(state().phases[0].id === firstId, 'phase reorder undoes');
}
// same machinery drives team types
{
  suTab('budget');
  const firstType = state().teamTypes[0];
  const grip = doc.querySelector('#setupView [data-sulist="type"] .su-row .su-grip');
  grip.dispatchEvent(new window.MouseEvent('pointerdown', { bubbles: true }));
  window.dispatchEvent(new window.MouseEvent('pointermove', { clientY: 999 }));
  window.dispatchEvent(new window.MouseEvent('pointerup'));
  ok(state().teamTypes[state().teamTypes.length - 1] === firstType, 'dragging a type grip reorders types');
}
// and workstreams (order persists in state.wsOrder)
{
  suTab('org');
  const firstWs = doc.querySelector('#setupView [data-sulist="ws"] .su-row').dataset.key;
  const grip = doc.querySelector('#setupView [data-sulist="ws"] .su-row .su-grip');
  grip.dispatchEvent(new window.MouseEvent('pointerdown', { bubbles: true }));
  window.dispatchEvent(new window.MouseEvent('pointermove', { clientY: 999 }));
  window.dispatchEvent(new window.MouseEvent('pointerup'));
  const order = state().wsOrder;
  ok(order[order.length - 1] === firstWs, 'dragging a workstream grip reorders wsOrder');
}
click(doc.querySelector('#viewTabs [data-view="planning"]'));
ok(doc.body.dataset.view === 'planning', 'back to planning after setup');

// ---------------------------------------------------------------- team + resources
{
  click(doc.querySelector('#resGrid [data-resadd]'));
  const nameInp = doc.querySelector('#resGrid [data-resadd] input');
  nameInp.value = 'Senior Dev (Kim)';
  nameInp.dispatchEvent(new window.KeyboardEvent('keydown', { key: 'Enter', bubbles: true }));
}
ok(state().team.length === 1, 'role added via the blank add row');
ok(doc.querySelectorAll('#resGrid .rrow[data-mid]').length === 1, 'resource row rendered for the member');
ok(doc.querySelectorAll('#resGrid .rh').length === 48, 'hour cells for every week (default 40h)');
// spreadsheet edit: click a cell, type 24, commit
{
  const cell = doc.querySelector('#resGrid .rh[data-w="2"]');
  click(cell);
  const inp = cell.querySelector('input.rh-edit');
  ok(!!inp, 'clicking an hour cell opens an inline editor');
  if (inp) {
    inp.value = '24';
    inp.dispatchEvent(new window.KeyboardEvent('keydown', { key: 'Enter', bubbles: true }));
    const m = state().team[0];
    const iso = Object.keys(m.weekHours)[0];
    ok(m.weekHours && Object.values(m.weekHours)[0] === 24, 'hours committed (' + JSON.stringify(m.weekHours) + ')');
  }
}
// right-click role row → context menu (delete lives here now)
{
  ok(!doc.querySelector('#resGrid .rr-del'), 'dedicated role delete button removed');
  const rrow = doc.querySelector('#resGrid .rrow[data-mid]');
  rrow.dispatchEvent(new window.MouseEvent('contextmenu', { bubbles: true, clientX: 150, clientY: 400 }));
  const labels = Array.from(doc.querySelectorAll('#popover .menu-list button')).map(b => b.textContent);
  ok(labels.some(l => /Remove role/.test(l)) && labels.some(l => /Rate card/.test(l)) &&
    labels.some(l => /Workstream/.test(l)) && labels.some(l => /Capacity/.test(l)),
    'role context menu offers rename/rate card/workstream/capacity/remove');
  ok(labels.some(l => /Start \/ end dates/.test(l)), 'role context menu offers start/end dates');
  doc.querySelector('#popover').hidden = true;
}

// start/end date quick update: zeroes weeks outside the window
{
  const rrow = doc.querySelector('#resGrid .rrow[data-mid]');
  const mid = rrow.dataset.mid;
  rrow.dispatchEvent(new window.MouseEvent('contextmenu', { bubbles: true, clientX: 150, clientY: 400 }));
  const btn = Array.from(doc.querySelectorAll('#popover .menu-list button'))
    .find(b => /Start \/ end dates/.test(b.textContent));
  click(btn);
  const sInp = doc.querySelector('#popover #rdStart'), eInp = doc.querySelector('#popover #rdEnd');
  ok(!!sInp && !!eInp, 'date popover shows start and end inputs');
  const meta = state().meta;
  const wk2 = window.RM.fmtISO(window.RM.weekStartDate(meta, 2));
  sInp.value = wk2; // start at week 2
  eInp.value = '';
  click(doc.querySelector('#popover #rdApply'));
  const m = state().team.find(x => x.id === mid);
  const iso0 = window.RM.fmtISO(window.RM.weekStartDate(meta, 0));
  const iso1 = window.RM.fmtISO(window.RM.weekStartDate(meta, 1));
  ok(m.weekHours[iso0] === 0 && m.weekHours[iso1] === 0,
    'weeks before the start date drop to 0 h');
  ok(window.RM.memberHoursForWeek(meta, m, 2) > 0, 'weeks from the start date keep their hours');
  window.dispatchEvent(new window.KeyboardEvent('keydown', { key: 'z', metaKey: true, bubbles: true }));
}

// header bar toggles the panel; Save button label
{
  ok(doc.querySelector('#btnSave').textContent.trim() === 'Save', 'save button says just "Save"');
  const collapsedBefore = doc.querySelector('#resPanel').classList.contains('collapsed');
  click(doc.querySelector('#resHead .rp-title'));
  ok(doc.querySelector('#resPanel').classList.contains('collapsed') !== collapsedBefore,
    'clicking the Resources header text toggles the panel');
  click(doc.querySelector('#resHead .rp-title'));
  ok(doc.querySelector('#resPanel').classList.contains('collapsed') === collapsedBefore,
    'clicking again restores it');
}

// dragging hours far right can never write past the project end date
{
  const rrow = doc.querySelector('#resGrid .rrow[data-mid]');
  const mid = rrow.dataset.mid;
  const cell = rrow.querySelector('.rh[data-w="1"]');
  cell.dispatchEvent(new window.MouseEvent('pointerdown', { bubbles: true, clientX: 40, clientY: 500, button: 0 }));
  window.dispatchEvent(new window.MouseEvent('pointermove', { bubbles: true, clientX: 99999, clientY: 500 }));
  window.dispatchEvent(new window.MouseEvent('pointerup', { bubbles: true, clientX: 99999, clientY: 500 }));
  const m = state().team.find(x => x.id === mid);
  const meta = state().meta;
  const lastISO = window.RM.fmtISO(window.RM.weekStartDate(meta, meta.numWeeks - 1));
  const past = Object.keys(m.weekHours || {}).filter(iso => iso > lastISO);
  ok(past.length === 0, 'hour fill drag clamps at the last project week (no keys past ' + lastISO + ')');
  window.dispatchEvent(new window.KeyboardEvent('keydown', { key: 'z', metaKey: true, bubbles: true }));
}

// capacity factor: editable, feeds availability (typed people only — an
// untyped person supplies nothing and shows a "set type" prompt instead)
{
  window.HeadwayApp.ai.commit('type first person', (s) => { s.team[0].capType = s.capTypes[0]; });
  const capChip = doc.querySelector('#resGrid .rrow[data-mid="' + state().team[0].id + '"] [data-rcap]');
  ok(!!capChip, 'resource rows show a capacity column');
  click(capChip);
  const capInp = doc.querySelector('#resGrid [data-rcap] input');
  ok(!!capInp, 'capacity chip opens an inline editor');
  capInp.value = '0.5';
  capInp.dispatchEvent(new window.KeyboardEvent('keydown', { key: 'Enter', bubbles: true }));
  ok(state().team[0].capacity === 0.5, 'capacity commits (0.5)');
  const heads = window.RM.memberHeads(state(), state().team[0], 0);
  ok(Math.abs(heads - 0.5) < 1e-9, 'availability scales by the capacity factor (' + heads + ')');
  window.dispatchEvent(new window.KeyboardEvent('keydown', { key: 'z', metaKey: true, bubbles: true }));
  window.dispatchEvent(new window.KeyboardEvent('keydown', { key: 'z', metaKey: true, bubbles: true }));
}

// ---------------------------------------------------------------- validation modal
click(doc.querySelector('#btnValidation'));
ok(!doc.querySelector('#modalHost').hidden, 'preflight modal opens');
click(doc.querySelector('#modalHost [data-m=x]'));

// ---------------------------------------------------------------- scoping view
click(doc.querySelector('#viewTabs [data-view="scoping"]'));
ok(doc.body.dataset.view === 'scoping', 'view switches to scoping');
ok(doc.querySelectorAll('#rows .sc-cell').length > 400, 'scoping cells rendered (' + doc.querySelectorAll('#rows .sc-cell').length + ')');
ok(doc.querySelectorAll('#rows .bar').length === 0, 'no bars in scoping view');
ok(doc.querySelectorAll('#hdrSprints .sc-hcell.sc-fixh').length === 8, 'fixed columns: assignees/size/risk/priority/duration/start/deadline/workstream/epic');
ok(doc.querySelectorAll('#hdrSprints .sc-hcell[data-col]').length === 14, 'default columns are 8 fixed + 6 text (incl. Description, Acceptance criteria)');
ok(!!doc.querySelector('#hdrSprints [data-col="description"]'), 'Description column shown by default');
{
  const hdrOrder = Array.from(doc.querySelectorAll('#hdrSprints .sc-hcell[data-col]')).map(c => c.dataset.col);
  ok(hdrOrder.indexOf('description') !== -1 && hdrOrder.indexOf('description') < hdrOrder.indexOf('enables'),
    'Description sits left of Enables');
  ok(hdrOrder.indexOf('description') === 0 && hdrOrder.indexOf('ac') === 1 && hdrOrder.indexOf('epic') === 2 && hdrOrder.indexOf('epic') < hdrOrder.indexOf('size'),
    'default order leads with Description and Epic before Size');
  ok(hdrOrder.indexOf('start') === hdrOrder.indexOf('duration') + 1, 'Start column follows Duration');
}
ok(doc.querySelectorAll('#hdrSprints .sc-rz').length === 14, 'column resize handles present');
const cell = doc.querySelector('#rows .row.item[data-id="' + itId + '"] [data-scope="notes"]');
ok(cell.getAttribute('contenteditable') === 'true', 'scoping cells are rich editors');
cell.innerHTML = 'noted in the grid';
cell.dispatchEvent(new window.FocusEvent('focusout', { bubbles: true }));
ok(state().items.find(i => i.id === itId).notes === 'noted in the grid', 'scoping cell edit commits');

// column management: right-click a header cell for its column menu (the ⋯
// button is gone) — remove the Description built-in…
ok(!doc.querySelector('#hdrSprints .sc-hmenu'), 'header cells carry no ⋯ menu button');
doc.querySelector('#hdrSprints .sc-hcell[data-col="description"]').dispatchEvent(
  new window.MouseEvent('contextmenu', { bubbles: true, cancelable: true, clientX: 120, clientY: 60 }));
click(Array.from(doc.querySelectorAll('#popover .menu-list button')).find(b => /Remove column/.test(b.textContent)));
ok(!doc.querySelector('#hdrSprints [data-col="description"]'), 'column removed via its menu');
// …then re-add it (now hidden) through the "+" menu
click(doc.querySelector('#hdrSprints [data-coladd]'));
const descAdd = Array.from(doc.querySelectorAll('#popover .menu-list button')).find(b => /^Description$/.test(b.textContent.trim()));
click(descAdd);
ok(!!doc.querySelector('#hdrSprints [data-col="description"]'), 'hidden built-in column re-added via + menu');
// leave the grid as it started: Description back in front of Enables
doc.querySelector('#hdrSprints .sc-hcell[data-col="description"]').dispatchEvent(
  new window.MouseEvent('contextmenu', { bubbles: true, cancelable: true, clientX: 120, clientY: 60 }));
{
  const mv = Array.from(doc.querySelectorAll('#popover .menu-list button')).find(b => /Move left/.test(b.textContent));
  if (mv) click(mv);
}

// custom column: create, edit a cell, move it left
click(doc.querySelector('#hdrSprints [data-coladd]'));
click(Array.from(doc.querySelectorAll('#popover .menu-list button')).find(b => /New custom column/.test(b.textContent)));
ok(!doc.querySelector('#modalHost').hidden, 'new-column modal opens');
doc.querySelector('#colName').value = 'Owner';
click(doc.querySelector('#colSave'));
const ownerCol = state().meta.scopeCols.find(c => c.label === 'Owner');
ok(!!ownerCol, 'custom column added to the document');
const ownerCell = doc.querySelector('#rows .row.item[data-id="' + itId + '"] [data-scope="' + (ownerCol && ownerCol.key) + '"]');
ok(!!ownerCell, 'custom column cells rendered');
if (ownerCell) {
  ownerCell.innerHTML = 'Rita';
  ownerCell.dispatchEvent(new window.FocusEvent('focusout', { bubbles: true }));
  ok(state().items.find(i => i.id === itId).custom[ownerCol.key] === 'Rita', 'custom cell edit commits to item.custom');
}
doc.querySelector('#hdrSprints .sc-hcell[data-col="' + ownerCol.key + '"]').dispatchEvent(
  new window.MouseEvent('contextmenu', { bubbles: true, cancelable: true, clientX: 120, clientY: 60 }));
click(Array.from(doc.querySelectorAll('#popover .menu-list button')).find(b => /Move left/.test(b.textContent)));
{
  const ord = state().meta.scopeColOrder;
  ok(ord.indexOf(ownerCol.key) === ord.length - 2, 'column moved left in the full order');
}

// scoping swaps the wks/headcount chips for a workstream dropdown chip
ok(!!doc.querySelector('#rows .row.item .r-ws') && !doc.querySelector('#rows .row.item .r-hc'),
  'scoping shows workstream chips instead of wks/headcount');
{
  const wsChip = doc.querySelector('#rows .row.item[data-id="' + itId + '"] [data-act="ws"]');
  click(wsChip);
  const prodOpt = Array.from(doc.querySelectorAll('#popover .menu-list button')).find(b => /^Product$/.test(b.textContent.trim()));
  ok(!!prodOpt, 'workstream chip opens the shared dropdown');
  click(prodOpt);
  ok(state().items.find(i => i.id === itId).workstream === 'Product', 'workstream commits from the dropdown');
}

click(doc.querySelector('#viewTabs [data-view="planning"]'));
ok(doc.body.dataset.view === 'planning', 'view switches back to planning');

// ---------------------------------------------------------------- blank add rows
{
  // only a phase with no features shows the add row (filled phases add via
  // Insert above/below): give the document one
  window.HeadwayApp.ai.commit('blank phase', (s) => {
    s.phases.push({ id: 'ph_blank_probe', name: 'Blank probe', description: '', bucket: false, collapsed: false });
  });
  const before = state().items.length;
  const addRow = doc.querySelector('#rows .row.addrow');
  ok(!!addRow && addRow.dataset.phase === 'ph_blank_probe', 'an empty phase ends with a blank add row');
  click(addRow.querySelector('.row-lane'));
  ok(state().items.length === before, 'the add row\'s timeline lane is inert');
  click(addRow.querySelector('.row-left'));
  ok(state().items.length === before + 1, 'clicking the add row creates an item in that phase');
  ok(state().items.some(i => i.phaseId === addRow.dataset.phase && i.feature === ''), 'new item lands in the clicked phase');
  ok(!doc.querySelector('#panel .p-dates') && !doc.querySelector('#panel .p-risknote'),
    'panel has no date-range row and no dependency-risk note');
  // collapsible sections
  ok(doc.querySelectorAll('#panel .p-sechead').length >= 6, 'panel renders collapsible section headers');
  const stSec = doc.querySelector('#panel .p-sec[data-sec="stories"]');
  ok(stSec && !stSec.classList.contains('open'), 'stories section starts collapsed');
  click(doc.querySelector('#panel [data-sectoggle="stories"]'));
  ok(doc.querySelector('#panel .p-sec[data-sec="stories"]').classList.contains('open'), 'section header toggles open');
  click(doc.querySelector('#panel [data-sectoggle="stories"]'));
  ok(!doc.querySelector('#panel .p-sec[data-sec="stories"]').classList.contains('open'), 'and toggles closed again');
  ok(!doc.querySelector('#panel input[data-f="headcount"]'), 'headcount input is gone from the panel');
  // fold the probe phase away: its new feature moves to the first phase
  window.HeadwayApp.ai.commit('blank phase cleanup', (s) => {
    s.items.forEach((i) => { if (i.phaseId === 'ph_blank_probe') i.phaseId = s.phases[0].id; });
    s.phases = s.phases.filter(p => p.id !== 'ph_blank_probe');
  });
}
{
  ok(!/No people yet/.test(doc.querySelector('#resGrid').textContent), 'no "No people yet" message');
  const resAdd = doc.querySelector('#resGrid [data-resadd]');
  ok(!!resAdd, 'resources panel ends with a blank add row');
  click(resAdd);
  const nameInp = doc.querySelector('#resGrid [data-resadd] input');
  ok(!!nameInp, 'clicking it opens an inline name input');
  const teamBefore = state().team.length;
  nameInp.value = 'Rita';
  nameInp.dispatchEvent(new window.KeyboardEvent('keydown', { key: 'Enter', bubbles: true }));
  ok(state().team.length === teamBefore + 1 && state().team.some(m => m.name === 'Rita'), 'Enter adds the person (40 h/week default)');
}

// ---------------------------------------------------------------- group by epic
window.eval("document.querySelector('[data-menu=\"view\"]').click()");
const groupBtn = Array.from(doc.querySelectorAll('#popover .menu-list button')).find(b => /Group by epic/.test(b.textContent));
click(groupBtn);
ok(doc.querySelectorAll('#rows .row.eband').length > 3, 'epic group bands rendered (' + doc.querySelectorAll('#rows .row.eband').length + ')');
{
  const eb = doc.querySelector('#rows .row.eband');
  ok(!!eb, 'an epic band renders with group-by-epic on');
  // the sprint pitch/phase the pinned band's gradient uses must match #bgcols
  const rs = doc.documentElement.style;
  const px = rs.getPropertyValue('--sprint-px'), off = rs.getPropertyValue('--sprint-off');
  ok(/px$/.test(px) && parseFloat(px) > 0, 'render publishes --sprint-px (' + px + ')');
  ok(/px$/.test(off), 'render publishes --sprint-off (' + off + ')');
  const line = doc.querySelector('#bgcols .bg-week.sprint');
  const at = line && parseFloat((line.getAttribute('style').match(/\+\s*(-?[\d.]+)px/) || [])[1]);
  const pxN = parseFloat(px), r = (((at - parseFloat(off)) % pxN) + pxN) % pxN;
  ok(typeof at === 'number' && !isNaN(at) && Math.min(r, pxN - r) < 0.01,
    'the drawn sprint lines land on the --sprint-off/--sprint-px lattice (' + at + ')');
}
{
  // groups always hold at least one feature, so no group carries an
  // "Add feature" row: the context menu (Insert above/below) adds features
  ok(doc.querySelectorAll('#rows .row.addrow.sub').length === 0, 'non-empty epic groups have no Add feature row');
  ok([...doc.querySelectorAll('#rows .addrow-lab')].every(l => /Add Feature/.test(l.textContent)), 'add rows say "Add Feature"');
  // Insert feature above inherits the anchor's epic while grouped
  const nBefore = state().items.length;
  const anchorRow = [...doc.querySelectorAll('#rows .row.item')].find(r => (state().items.find(i => i.id === r.dataset.id) || {}).epic);
  const anchorIt = state().items.find(i => i.id === anchorRow.dataset.id);
  anchorRow.dispatchEvent(new window.MouseEvent('contextmenu', { bubbles: true, clientX: 220, clientY: 220 }));
  click([...doc.querySelectorAll('#popover .menu-list button')].find(b => /Insert feature above/.test(b.textContent)));
  const s3 = state();
  const ins = s3.items[s3.items.findIndex(i => i.id === anchorIt.id) - 1];
  ok(ins && ins.feature === '' && ins.epic === anchorIt.epic && (ins.workstream || '') === (anchorIt.workstream || ''),
    'insert above inherits the anchor\'s epic and workstream');
  if (doc.activeElement && doc.activeElement.blur) doc.activeElement.blur();
  window.dispatchEvent(new window.KeyboardEvent('keydown', { key: 'z', metaKey: true, bubbles: true }));
  ok(state().items.length === nBefore, 'undo removes the inserted feature');
}
{
  const eb = Array.from(doc.querySelectorAll('#rows .row.eband')).find(r => r.dataset.epic);
  eb.dispatchEvent(new window.MouseEvent('contextmenu', { bubbles: true, clientX: 220, clientY: 220 }));
  const labels = Array.from(doc.querySelectorAll('#popover .menu-list button')).map(b => b.textContent);
  ok(labels.some(l => /Edit epic/.test(l)) && labels.some(l => /Delete epic/.test(l)),
    'right-clicking an epic band offers edit/delete');
  const epName = eb.dataset.epic;
  click(Array.from(doc.querySelectorAll('#popover .menu-list button')).find(b => /Edit epic/.test(b.textContent)));
  const epJira = doc.querySelector('#modalHost #epJira');
  ok(!!epJira, 'Edit epic dialog has a Jira key field');
  epJira.value = 'hw-1';
  click(doc.querySelector('#modalHost #epSave'));
  ok(state().epicJira[epName] === 'HW-1', 'epic Jira key saves to state.epicJira');
  doc.querySelector('#popover').hidden = true;
}
window.eval("document.querySelector('[data-menu=\"view\"]').click()");
const groupBtn2 = Array.from(doc.querySelectorAll('#popover .menu-list button')).find(b => /Group by epic/.test(b.textContent));
click(groupBtn2);
ok(doc.querySelectorAll('#rows .row.eband').length === 0, 'grouping toggles back off');

// ---------------------------------------------------------------- group by workstream
window.eval("document.querySelector('[data-menu=\"view\"]').click()");
click(Array.from(doc.querySelectorAll('#popover .menu-list button')).find(b => /Group by workstream/.test(b.textContent)));
ok(doc.querySelectorAll('#rows .row.eband[data-ws]').length > 1,
  'workstream group bands rendered (' + doc.querySelectorAll('#rows .row.eband[data-ws]').length + ')');
ok(doc.querySelectorAll('#rows .row.addrow.sub[data-ws]').length === 0,
  'non-empty workstream groups have no Add feature row');
window.eval("document.querySelector('[data-menu=\"view\"]').click()");
click(Array.from(doc.querySelectorAll('#popover .menu-list button')).find(b => /Group by workstream/.test(b.textContent)));
ok(doc.querySelectorAll('#rows .row.eband').length === 0, 'workstream grouping toggles back off');

// resources rows carry a workstream chip
ok(!!doc.querySelector('#resGrid [data-bact="ws"]'), 'resource rows have a workstream chip');

// ---------------------------------------------------------------- budgeting view
{
  click(doc.querySelector('#viewTabs [data-view="budget"]'));
  ok(doc.body.dataset.view === 'budget', 'Budgeting tab switches the view');
  ok(!!doc.querySelector('#hdrSprints .sp-date'), 'budgeting shows the planning timeline header (dates/sprints)');
  const roleRow = doc.querySelector('#rows .row.brole[data-mid]');
  ok(!!roleRow, 'role rows render on the board');
  ok(roleRow.querySelectorAll('.bu-cell').length === state().meta.numWeeks, 'one week cell per project week');
  ok(!roleRow.querySelector('.res-cap'), 'no capacity column in budgeting');
  const rateInp = roleRow.querySelector('input[data-bud="rate"]');
  rateInp.value = '200';
  rateInp.dispatchEvent(new window.Event('change', { bubbles: true }));
  const c2 = doc.querySelector('#rows .row.brole input[data-bud="cost"]');
  c2.value = '120';
  c2.dispatchEvent(new window.Event('change', { bubbles: true }));
  const m = state().team[0];
  ok(m.rate === 200 && m.cost === 120, 'rate and cost commit');
  ok(/40%/.test(doc.querySelector('#rows .row.brole').textContent), 'margin computed (40%)');
  // cost column precedes rate; both left pane and header agree
  {
    const inps = Array.from(doc.querySelectorAll('#rows .row.brole[data-mid] input[data-bud]')).filter(i => i.dataset.bud !== 'name' && i.dataset.bud !== 'role');
    ok(inps[0].dataset.bud === 'cost' && inps[1].dataset.bud === 'rate', 'Cost input comes before Rate');
    const labels = Array.from(doc.querySelectorAll('.hl-cols .bu-only')).map(i => i.textContent);
    ok(labels.join(',') === 'Role,Rate card,Capacity,Workstream,Cost,Rate,Margin,Total', 'header labels spelled out, cost before rate');
  }
  // total = actual hours × RATE
  {
    const hrs = window.RM.roleTotalHours(window.__headway.getState(), state().team[0]);
    const exp = '$' + Math.round(hrs * 200).toString().replace(/\B(?=(\d{3})+(?!\d))/g, ',');
    ok(doc.querySelector('#rows .row.btotal').textContent.includes(exp),
      'total row = hours × rate (' + exp + ')');
  }
  // holiday-clipped weeks show the actual hours below in small text
  ok(!!doc.querySelector('#rows .bu-cell.clipped .bu-sub'), 'holiday weeks show (actual hours) sub-line');
  // week cells are keyboard-reachable
  ok(doc.querySelector('#rows .bu-cell').getAttribute('tabindex') === '0', 'week cells are tabbable');
  // the reports drawer is gone — Reports is a full tab now
  ok(!doc.querySelector('#repPanel'), 'reports drawer removed from the budget view');
  // add-role and add-cost rows
  ok(!!doc.querySelector('#rows .row.addrow[data-kind="baddrole"]'), 'budget ends the roster with an Add role row');
  ok(!!doc.querySelector('#rows .row.addrow[data-kind="baddcost"]'), 'budget offers an Add cost row');
  // role + workstream chips edit via the shared dropdown
  click(doc.querySelector('#rows .row.brole [data-bact="type"]'));
  ok(!doc.querySelector('#popover').hidden, 'role chip opens the shared dropdown');
  const typePick = Array.from(doc.querySelectorAll('#popover .menu-list button'))[1];
  const typeName = typePick.textContent.trim();
  click(typePick);
  ok(state().team[0].type === typeName, 'picking a role commits');
  click(doc.querySelector('#rows .row.brole [data-bact="ws"]'));
  ok(!doc.querySelector('#popover').hidden, 'workstream chip opens the shared dropdown');
  doc.querySelector('#popover').hidden = true;

  // money cells: wide enough to edit, contents selected on focus
  const costInp = doc.querySelector('#rows .row.brole[data-mid] input[data-bud="cost"]');
  ok(costInp.style.width === 'var(--bu-w-cost)' &&
    parseInt(doc.documentElement.style.getPropertyValue('--bu-w-cost'), 10) >= 72,
    'budget money cells are wide enough to edit');
  let selected = false;
  costInp.select = () => { selected = true; };
  costInp.dispatchEvent(new window.FocusEvent('focusin', { bubbles: true }));
  ok(selected, 'focusing a budget money cell selects its contents');

  // clicking another money cell while one is mid-edit lands focus on the
  // fresh element even though the commit re-render replaced it
  {
    const row = doc.querySelector('#rows .row.brole[data-mid]');
    const cost2 = row.querySelector('input[data-bud="cost"]');
    cost2.focus();
    cost2.value = '111';
    const rate2 = row.querySelector('input[data-bud="rate"]');
    rate2.dispatchEvent(new window.MouseEvent('pointerdown', { bubbles: true, button: 0 }));
    cost2.dispatchEvent(new window.Event('change', { bubbles: true })); // blur-commit → re-render
    window.dispatchEvent(new window.MouseEvent('pointerup', { bubbles: true }));
    const active = doc.activeElement;
    ok(active && active.dataset && active.dataset.bud === 'rate' && active !== rate2,
      'clicking Rate while Cost is mid-edit refocuses the fresh Rate cell');
  }

  // spreadsheet-style fill handle on week-hour cells
  {
    const row = doc.querySelector('#rows .row.brole[data-mid]');
    const mid = row.dataset.mid;
    const cell0 = row.querySelector('.bu-cell');
    click(cell0);
    const ed = cell0.querySelector('input');
    ok(!!ed, 'clicking a week cell opens its editor');
    ed.value = '20';
    ed.dispatchEvent(new window.FocusEvent('blur')); // commits → re-render
    const row2 = doc.querySelector('#rows .row.brole[data-mid="' + mid + '"]');
    const c0 = row2.querySelector('.bu-cell');
    c0.focus();
    c0.dispatchEvent(new window.FocusEvent('focusin', { bubbles: true }));
    const handle = c0.querySelector('.bu-fill');
    ok(!!handle, 'focused week cell grows a fill handle');
    handle.dispatchEvent(new window.MouseEvent('pointerdown', { bubbles: true, button: 0, clientX: 0, clientY: 0 }));
    window.dispatchEvent(new window.MouseEvent('pointermove', { clientX: 28 * 4 + 3, clientY: 0 }));
    window.dispatchEvent(new window.MouseEvent('pointerup', { bubbles: true }));
    const member = state().team.find(t => t.id === mid);
    const isoOf = w => window.RM.fmtISO(window.RM.weekStartDate(state().meta, w));
    ok([0, 1, 2, 3, 4].every(w => member.weekHours[isoOf(w)] === 20),
      'dragging the fill handle spreads the value across the crossed weeks');
  }

  // vertical fill: dragging the handle down spreads across roles too
  {
    if (state().team.length < 2) {
      click(doc.querySelector('#viewTabs [data-view="planning"]'));
      click(doc.querySelector('#resGrid [data-resadd]'));
      const ni = doc.querySelector('#resGrid [data-resadd] input');
      ni.value = 'Second Role';
      ni.dispatchEvent(new window.Event('change', { bubbles: true }));
      click(doc.querySelector('#viewTabs [data-view="budget"]'));
    }
    const vRows = Array.from(doc.querySelectorAll('#rows .row.brole[data-mid]'));
    ok(vRows.length >= 2, 'two roles available for vertical fill');
    vRows.forEach((r, i) => {
      const rect = { top: i * 28, bottom: i * 28 + 28, left: 0, right: 9999, width: 9999, height: 28 };
      r.getBoundingClientRect = () => rect;
      r.querySelector('.row-lane').getBoundingClientRect = () => rect;
    });
    const srcCell = vRows[0].querySelector('.bu-cell'); // week 0 holds 20h from the previous test
    srcCell.focus();
    srcCell.dispatchEvent(new window.FocusEvent('focusin', { bubbles: true }));
    const h2 = srcCell.querySelector('.bu-fill');
    h2.dispatchEvent(new window.MouseEvent('pointerdown', { bubbles: true, button: 0, clientX: 3, clientY: 5 }));
    window.dispatchEvent(new window.MouseEvent('pointermove', { clientX: 28 * 2 + 3, clientY: 28 + 14 }));
    window.dispatchEvent(new window.MouseEvent('pointerup', { bubbles: true }));
    const below = state().team[1];
    const isoOf2 = w => window.RM.fmtISO(window.RM.weekStartDate(state().meta, w));
    ok([0, 1, 2].every(w => below.weekHours[isoOf2(w)] === 20),
      'dragging the fill handle downward spreads the value to the roles below');
  }

  // zoom cluster: shown on Budgeting
  ok(!doc.querySelector('#zoomCtl').hidden, 'zoom cluster visible on Budgeting');
  {
    const wpx0 = window.__headway.getState() && doc.documentElement.style.getPropertyValue('--week-px');
    window.eval("document.querySelector('#zoomInBtn').click()");
    ok(doc.documentElement.style.getPropertyValue('--week-px') !== wpx0, 'zoom + changes the week width on Budgeting');
    window.eval("document.querySelector('#zoomOutBtn').click()");
  }
  ok(!doc.querySelector('#btnPresent') && doc.querySelectorAll('#zoomCtl button').length === 2,
    'the cluster is zoom in / zoom out only (no expand button)');
  // the cluster only exists on Planning/Budgeting
  click(doc.querySelector('#viewTabs [data-view="scoping"]'));
  ok(doc.querySelector('#zoomCtl').hidden, 'no zoom cluster on Scoping');
  click(doc.querySelector('#viewTabs [data-view="planning"]'));
}

// ---------------------------------------------------------------- budget: costs + add role
{
  click(doc.querySelector('#viewTabs [data-view="budget"]'));
  // add a fixed cost via its addrow
  const cBefore = (state().costs || []).length;
  click(doc.querySelector('#rows .row.addrow[data-kind="baddcost"]'));
  ok(state().costs.length === cBefore + 1, 'Add cost creates a cost row');
  const costRow = doc.querySelector('#rows .row.bcost[data-cost]');
  ok(!!costRow, 'cost rows render in the Costs band');
  const amt = costRow.querySelector('[data-cf="amount"]');
  amt.value = '2500';
  amt.dispatchEvent(new window.Event('change', { bubbles: true }));
  ok(state().costs[0].amount === 2500, 'cost amount commits');
  ok(!!doc.querySelector('#rows .row.bcost .bu-costmark'), 'cost occurrences mark the timeline lane');
  // kind → weekly multiplies occurrences
  click(doc.querySelector('#rows .row.bcost [data-cact="kind"]'));
  click(Array.from(doc.querySelectorAll('#popover .menu-list button')).find(b => /Weekly/.test(b.textContent)));
  ok(state().costs[0].kind === 'weekly', 'cost cadence commits');
  ok(doc.querySelectorAll('#rows .row.bcost .bu-costmark').length > 3, 'recurring costs mark every occurrence');
  ok(/Costs \$/.test(doc.querySelector('#rows .row.btotal').textContent), 'total row includes the cost spend');
  // remove via the cost context menu
  doc.querySelector('#rows .row.bcost[data-cost]')
    .dispatchEvent(new window.MouseEvent('contextmenu', { bubbles: true, clientX: 200, clientY: 300 }));
  click(Array.from(doc.querySelectorAll('#popover .menu-list button')).find(b => /Remove cost/.test(b.textContent)));
  ok(state().costs.length === cBefore, 'context menu removes the cost');

  // add-role row spawns an inline name input and creates the person
  const tBefore = state().team.length;
  click(doc.querySelector('#rows .row.addrow[data-kind="baddrole"]'));
  const roleInp = doc.querySelector('#rows .row.addrow[data-kind="baddrole"] input');
  ok(!!roleInp, 'Add role opens an inline name input');
  roleInp.value = 'Norah';
  roleInp.dispatchEvent(new window.KeyboardEvent('keydown', { key: 'Enter', bubbles: true }));
  ok(state().team.length === tBefore + 1 && state().team.some(m => m.name === 'Norah'),
    'Enter adds the person from Budgeting');
  // budget person row context menu
  doc.querySelector('#rows .row.brole[data-mid]')
    .dispatchEvent(new window.MouseEvent('contextmenu', { bubbles: true, clientX: 200, clientY: 200 }));
  const rl = Array.from(doc.querySelectorAll('#popover .menu-list button')).map(b => b.textContent);
  ok(rl.some(l => /Rename/.test(l)) && rl.some(l => /Remove person/.test(l)),
    'budget person rows have a context menu');
  doc.querySelector('#popover').hidden = true;
  click(doc.querySelector('#viewTabs [data-view="planning"]'));
}

// ---------------------------------------------------------------- epics in Setup
{
  click(doc.querySelector('#btnSetup'));
  suTab('org');
  ok(doc.querySelectorAll('#setupView [data-suepedit]').length > 0, 'Setup lists epics with edit controls');
  click(doc.querySelector('#setupView [data-suepedit]'));
  ok(!doc.querySelector('#modalHost').hidden, 'epic edit modal opens from Setup');
  window.dispatchEvent(new window.KeyboardEvent('keydown', { key: 'Escape', bubbles: true }));
  click(doc.querySelector('#viewTabs [data-view="planning"]'));
}

// ---------------------------------------------------------------- hierarchy card in Setup
{
  click(doc.querySelector('#btnSetup'));
  suTab('org');
  const card = [...doc.querySelectorAll('#setupView .su-card h2')].find(h => h.textContent === 'Hierarchy');
  ok(!!card, 'Hierarchy card renders in the Workstreams tab');
  const bugChip = doc.querySelector('button[data-suhtype="feature:bug"]');
  ok(bugChip && bugChip.classList.contains('on'), 'Bug is allowed at the Feature level by default');
  click(bugChip);
  ok(state().meta.hierarchy.levels[1].types.indexOf('bug') === -1, 'clicking a chip disallows the type');
  click(doc.querySelector('button[data-suhtype="feature:bug"]'));
  ok(state().meta.hierarchy.levels[1].types.indexOf('bug') !== -1, 'clicking again re-allows it');
  const lbl = doc.querySelector('input[data-suhlabel="story"]');
  lbl.value = 'Task'; lbl.dispatchEvent(new window.Event('change', { bubbles: true }));
  ok(state().meta.hierarchy.levels[2].label === 'Task', 'level label edit commits');
  // the focus-the-new-row call lands inside a requestAnimationFrame; run it
  // synchronously here so the assertion below doesn't need to wait a real frame
  {
    const realRaf = window.requestAnimationFrame;
    window.requestAnimationFrame = (cb) => cb();
    click(doc.querySelector('#suHierAdd'));
    window.requestAnimationFrame = realRaf;
  }
  ok(state().meta.itemTypes.some(t => t.label === 'New type'), 'Add type appends a record');
  const hierAddLabels = [...doc.querySelectorAll('#setupView input[data-suhtlabel]')];
  const lastHierAddLabel = hierAddLabels[hierAddLabels.length - 1];
  ok(doc.activeElement === lastHierAddLabel,
    'Add type focuses the newly added type\'s label input, not the first (Epic) one');
  const any = doc.querySelector('#suHierAny');
  any.checked = true; any.dispatchEvent(new window.Event('change', { bubbles: true }));
  ok(state().meta.hierarchy.anyTypeAnyLevel === true && !doc.querySelector('button[data-suhtype]'), 'the switch hides the chips');
  any.checked = false; any.dispatchEvent(new window.Event('change', { bubbles: true }));
  click(doc.querySelector('#viewTabs [data-view="planning"]'));
}

// ---------------------------------------------------------------- level labels drive prominent UI strings
{
  click(doc.querySelector('#btnSetup'));
  suTab('org');
  const sl = doc.querySelector('input[data-suhlabel="story"]');
  sl.value = 'Task'; sl.dispatchEvent(new window.Event('change', { bubbles: true }));
  const fl = doc.querySelector('input[data-suhlabel="feature"]');
  fl.value = 'Capability'; fl.dispatchEvent(new window.Event('change', { bubbles: true }));
  window.HeadwayApp.ai.commit('empty phase probe', (s) => {
    s.phases.push({ id: 'ph_empty_probe', name: 'Empty probe', description: '', bucket: false, collapsed: false });
  });
  click(doc.querySelector('.tab[data-view="planning"]') || doc.querySelector('[data-view="planning"]'));
  ok([...doc.querySelectorAll('.addrow-lab')].some(el => /Add Capability/.test(el.textContent)), 'Add-row wording follows the level label');
  window.HeadwayApp.ai.commit('empty phase probe cleanup', (s) => {
    s.phases = s.phases.filter(p => p.id !== 'ph_empty_probe');
  });
  // reset the labels back to their defaults through Setup, since the inputs
  // do not survive the view switch
  click(doc.querySelector('#btnSetup'));
  suTab('org');
  const fl2 = doc.querySelector('input[data-suhlabel="feature"]');
  fl2.value = 'Feature'; fl2.dispatchEvent(new window.Event('change', { bubbles: true }));
  const sl2 = doc.querySelector('input[data-suhlabel="story"]');
  sl2.value = 'Story'; sl2.dispatchEvent(new window.Event('change', { bubbles: true }));
  ok(state().meta.hierarchy.levels[1].label === 'Feature' && state().meta.hierarchy.levels[2].label === 'Story',
    'level labels reset back to their defaults');
  click(doc.querySelector('#viewTabs [data-view="planning"]'));
}

// ---------------------------------------------------------------- reports tab
{
  ok(!doc.querySelector('#repPanel'), 'the old reports drawer is gone');
  click(doc.querySelector('#viewTabs [data-view="reports"]'));
  ok(doc.body.dataset.view === 'reports', 'Reports is a full view tab');
  ok(doc.querySelectorAll('#reportsView .rp-kpi').length >= 5, 'dashboard leads with KPI cards');
  ok(doc.querySelectorAll('#reportsView .rp-card').length >= 4,
    'dashboard renders phase progress, workstream costs, spend curve, milestones, flags');
  ok(/\$/.test(doc.querySelector('#reportsView').textContent), 'reports price in dollars');
  ok(doc.querySelector('#panel').hidden, 'no edit panel on Reports');
  click(doc.querySelector('#viewTabs [data-view="planning"]'));
}

// ---------------------------------------------------------------- capacity row toggle
{
  ok(!doc.body.classList.contains('no-cap'), 'capacity row shown by default');
  window.eval("document.querySelector('[data-menu=\"view\"]').click()");
  click(Array.from(doc.querySelectorAll('#popover .menu-list button')).find(b => /Capacity row/.test(b.textContent)));
  ok(doc.body.classList.contains('no-cap'), 'View → Capacity row hides the capacity header');
  window.eval("document.querySelector('[data-menu=\"view\"]').click()");
  click(Array.from(doc.querySelectorAll('#popover .menu-list button')).find(b => /Capacity row/.test(b.textContent)));
  ok(!doc.body.classList.contains('no-cap'), 'toggling again shows it');
}

// ---------------------------------------------------------------- capacity feature switch (Setup)
{
  window.eval("document.querySelector('#btnSetup').click()");
  suTab('scheduling'); // the capacity switch lives on the Capacity tab
  const capChk = doc.querySelector('#suCapEnable');
  ok(capChk && capChk.checked, 'Setup capacity checkbox reflects the enabled fixture');
  capChk.checked = false;
  capChk.dispatchEvent(new window.Event('change', { bubbles: true }));
  window.eval("document.querySelector('#viewTabs [data-view=\"planning\"]').click()");
  ok(doc.body.classList.contains('no-cap'), 'disabling the capacity feature hides the capacity row');
  ok(!doc.querySelector('#resGrid .res-cap'), 'per-role capacity chips hidden too');
  window.eval("document.querySelector('[data-menu=\"view\"]').click()");
  ok(!Array.from(doc.querySelectorAll('#popover .menu-list button')).some(b => /Capacity row/.test(b.textContent)),
    'View menu drops its capacity-row toggle');
  doc.querySelector('#popover').hidden = true;
  window.eval("document.querySelector('#btnSetup').click()");
  const capChk2 = doc.querySelector('#suCapEnable');
  capChk2.checked = true;
  capChk2.dispatchEvent(new window.Event('change', { bubbles: true }));
  window.eval("document.querySelector('#viewTabs [data-view=\"planning\"]').click()");
  ok(!doc.body.classList.contains('no-cap'), 're-enabling restores the capacity row');
}

// ---------------------------------------------------------------- critical path toggle
{
  const critOn = doc.querySelectorAll('#rows .bar.crit').length;
  ok(critOn > 0, 'critical-path bars highlighted by default (' + critOn + ')');
  window.eval("document.querySelector('[data-menu=\"view\"]').click()");
  const critBtn = Array.from(doc.querySelectorAll('#popover .menu-list button')).find(b => /Critical path highlight/.test(b.textContent));
  ok(!!critBtn, 'View menu offers a critical-path toggle');
  click(critBtn);
  ok(doc.querySelectorAll('#rows .bar.crit').length === 0 &&
    doc.querySelectorAll('#arrows g.edge.crit').length === 0,
    'toggle off removes the orange highlight from bars and arrows');
  window.eval("document.querySelector('[data-menu=\"view\"]').click()");
  click(Array.from(doc.querySelectorAll('#popover .menu-list button')).find(b => /Critical path highlight/.test(b.textContent)));
  ok(doc.querySelectorAll('#rows .bar.crit').length === critOn, 'toggle back on restores it');
}

// ---------------------------------------------------------------- lane click selects
{
  click(doc.querySelector('#rows .row.item .r-num'));
  ok(!doc.querySelector('#panel').hidden, 'clicking a row number opens the panel');
  // any part of the row selects it — empty lane space included
  const laneRow = doc.querySelectorAll('#rows .row.item')[1];
  const laneRowId = laneRow.dataset.id;
  click(laneRow.querySelector('.row-lane'));
  ok(!doc.querySelector('#panel').hidden, 'clicking lane space keeps the panel open');
  ok(doc.querySelector('#rows .row.item.selected') &&
    doc.querySelector('#rows .row.item.selected').dataset.id === laneRowId,
    'clicking lane space selects that row');
  window.dispatchEvent(new window.KeyboardEvent('keydown', { key: 'Escape', bubbles: true }));
}

// ---------------------------------------------------------------- story add keeps focus
{
  const chev = doc.querySelector('#rows .row.item .r-chev');
  click(chev); // expand stories
  ok(/Feature/.test(doc.querySelector('#detailBtn').textContent),
    'expanding an item stays in Feature view');
  ok(doc.querySelectorAll('#rows .row.story-add').length === 1,
    'only the expanded item opens its stories');
  const inp = doc.querySelector('#rows .row.story-add .st-add-input');
  ok(!!inp, 'story add input rendered');
  click(inp);
  ok(inp.isConnected, 'clicking the story add input does not re-render it away');
  click(doc.querySelector('#rows .row.item .r-chev')); // collapse again
}

// ---------------------------------------------------------------- story timelines
{
  // create a story on the first item via the quick-add input, then test
  let withStories = state().items.find(i => i.stories.length > 0);
  if (!withStories) {
    const firstId = doc.querySelector('#rows .row.item').dataset.id;
    click(doc.querySelector('#rows .row.item[data-id="' + firstId + '"] .r-chev'));
    const addInp = doc.querySelector('#rows .row.story-add[data-id="' + firstId + '"] .st-add-input');
    addInp.value = 'test story';
    addInp.dispatchEvent(new window.KeyboardEvent('keydown', { key: 'Enter', bubbles: true }));
    click(doc.querySelector('#rows .row.item[data-id="' + firstId + '"] .r-chev')); // collapse; re-expanded below
    withStories = state().items.find(i => i.stories.length > 0);
  }
  if (withStories) {
    click(doc.querySelector('#rows .row.item[data-id="' + withStories.id + '"] .r-chev'));
    const stRow = doc.querySelector('#rows .row.story[data-story]');
    ok(!!stRow, 'story rows render when expanded');
    ok(!stRow.querySelector('input[type="checkbox"]'), 'story rows have no checkbox');
    ok(!doc.querySelector('#panel input[data-pst-done]'), 'panel story list has no checkboxes either');
    const stId = stRow.dataset.story;
    ok(!stRow.querySelector('.st-bar'), 'story starts without a timeline bar');
    ok(!stRow.querySelector('.st-tick'), 'no start tick on story lanes — one rectangle only');
    // hovering the empty lane previews the landing slot with the dashed
    // ghost alone (the old low-opacity preview bar is gone)
    stRow.querySelector('.row-lane')
      .dispatchEvent(new window.PointerEvent('pointermove', { bubbles: true, clientX: 600, clientY: 300 }));
    ok(!!stRow.querySelector('.place-ghost'), 'hovering an empty story lane shows the dashed place ghost');
    ok(!stRow.querySelector('.place-preview'), 'no low-opacity preview bar on story lanes');
    // double-click the story lane → timeline appears
    stRow.querySelector('.row-lane')
      .dispatchEvent(new window.MouseEvent('dblclick', { bubbles: true, clientX: 600, clientY: 300 }));
    let st = state().items.find(i => i.id === withStories.id).stories.find(s => s.id === stId);
    ok(st.startDay != null && st.durDays >= 5, 'double-clicking the story lane gives it a timeline');
    const stBar = doc.querySelector('#rows .st-bar[data-stbar="' + stId + '"]');
    ok(!!stBar, 'story timeline renders as a mini bar');
    ok(!!stBar.querySelector('.port[data-port="in"]') && !!stBar.querySelector('.port[data-port="out"]'),
      'story bars carry the same in/out dependency ports as feature bars');
    ok(!!stBar.querySelector('.stb-label'), 'story bar carries its title as a quiet label');
    // context menu offers Remove timeline; removing clears it
    doc.querySelector('#rows .row.story[data-story="' + stId + '"]')
      .dispatchEvent(new window.MouseEvent('contextmenu', { bubbles: true, clientX: 300, clientY: 300 }));
    const rmBtn = Array.from(doc.querySelectorAll('#popover .menu-list button')).find(b => /Remove timeline/.test(b.textContent));
    ok(!!rmBtn, 'story context menu offers Remove timeline');
    click(rmBtn);
    st = state().items.find(i => i.id === withStories.id).stories.find(s => s.id === stId);
    ok(st.startDay == null && st.durDays == null, 'Remove timeline clears the story schedule');
    // View → Collapse all / Expand all features
    window.eval("document.querySelector('[data-menu=\"view\"]').click()");
    click(Array.from(doc.querySelectorAll('#popover .menu-list button')).find(b => /Collapse all features/.test(b.textContent)));
    ok(!doc.querySelector('#rows .row.story[data-story]'), 'Collapse all features hides story rows');
    window.eval("document.querySelector('[data-menu=\"view\"]').click()");
    click(Array.from(doc.querySelectorAll('#popover .menu-list button')).find(b => /Expand all features/.test(b.textContent)));
    ok(!!doc.querySelector('#rows .row.story[data-story="' + stId + '"]'), 'Expand all features shows story rows again');
    // story-add "+" is a lucide icon, not placeholder text
    const addRow = doc.querySelector('#rows .row.story-add');
    ok(!!addRow.querySelector('.st-add-ico') &&
      !/\+/.test(addRow.querySelector('.st-add-input').placeholder),
      'story add row leads with a lucide plus icon');
    // scoping shows story rows too while expanded
    click(doc.querySelector('#viewTabs [data-view="scoping"]'));
    ok(!!doc.querySelector('#rows .row.story[data-story="' + stId + '"]'), 'scoping renders story rows too');
    click(doc.querySelector('#viewTabs [data-view="planning"]'));
    window.eval("document.querySelector('[data-menu=\"view\"]').click()");
    click(Array.from(doc.querySelectorAll('#popover .menu-list button')).find(b => /Collapse all features/.test(b.textContent)));
  } else ok(true, '(no item with stories in seed)');
}

// ---------------------------------------------------------------- no validation stripes on bars
ok(!doc.querySelector('#rows .bar.warnbar') && !doc.querySelector('#rows .bar.errbar'),
  'bars carry no warning/error top stripes');

// ---------------------------------------------------------------- risk adds no padding
{
  ok(state().items.every(i => (i.riskDays || 0) === 0), 'no item carries risk padding');
  const bar = doc.querySelector('#rows .bar[data-bar]');
  const barIt = bar && state().items.find(i => i.id === bar.dataset.bar);
  ok(bar && Math.round(parseFloat(bar.style.width)) === Math.round(barIt.durDays * (28 / 5)),
    'bar width = durDays exactly (no risk padding)');
}

// ---------------------------------------------------------------- weeks chip inline edit
{
  const wk = doc.querySelector('#rows .row.item .r-wk.editable');
  ok(!!wk, 'scheduled rows show an editable weeks chip');
  click(wk);
  const wkInp = doc.querySelector('#rows .row.item .r-wk input');
  ok(!!wkInp, 'weeks chip opens an inline editor');
  const wkRow = wkInp.closest('.row.item');
  const wkId = wkRow.dataset.id;
  wkInp.value = '3';
  wkInp.dispatchEvent(new window.KeyboardEvent('keydown', { key: 'Enter', bubbles: true }));
  const wkIt = state().items.find(i => i.id === wkId);
  ok(window.RM.workInSpan(state().meta, wkIt.startDay, wkIt.durDays) === 15,
    'weeks edit commits 3w = 15 working days (span stretches over holidays, got ' + wkIt.durDays + ')');
}

// ---------------------------------------------------------------- double-click-to-place unscheduled
ok(!doc.querySelector('#rows .ghost-pill'), 'no ghost pill on unscheduled rows');
{
  const unsched = state().items.find(i => i.startDay == null);
  const row = doc.querySelector('#rows .row.item[data-id="' + unsched.id + '"]');
  if (row) {
    click(row.querySelector('.row-lane'));
    let after2 = state().items.find(i => i.id === unsched.id);
    ok(after2.startDay == null, 'a single click no longer places the item');
    doc.querySelector('#rows .row.item[data-id="' + unsched.id + '"] .row-lane')
      .dispatchEvent(new window.MouseEvent('dblclick', { bubbles: true, clientX: 600, clientY: 300 }));
    after2 = state().items.find(i => i.id === unsched.id);
    ok(after2.startDay != null && after2.durDays >= 5, 'double-clicking the empty lane places the item');
    window.dispatchEvent(new window.KeyboardEvent('keydown', { key: 'z', metaKey: true, bubbles: true }));
    // two quick single clicks also place it — the first click's re-render
    // replaces the lane node, so the app counts clicks itself
    click(doc.querySelector('#rows .row.item[data-id="' + unsched.id + '"] .row-lane'));
    click(doc.querySelector('#rows .row.item[data-id="' + unsched.id + '"] .row-lane'));
    after2 = state().items.find(i => i.id === unsched.id);
    ok(after2.startDay != null, 'two quick clicks on the lane place the item (manual dblclick counter)');
    window.dispatchEvent(new window.KeyboardEvent('keydown', { key: 'z', metaKey: true, bubbles: true }));
  } else ok(true, '(no visible unscheduled row to place)');
}

// ---------------------------------------------------------------- insert above/below stays inline
{
  // deselect first (Escape) so the panel is shut
  window.dispatchEvent(new window.KeyboardEvent('keydown', { key: 'Escape', bubbles: true }));
  const anchorRow = doc.querySelectorAll('#rows .row.item')[2];
  const anchorId = anchorRow.dataset.id;
  const before = state().items.length;
  anchorRow.dispatchEvent(new window.MouseEvent('contextmenu', { bubbles: true, clientX: 220, clientY: 220 }));
  const aboveBtn = Array.from(doc.querySelectorAll('#popover .menu-list button')).find(b => /Insert feature above/.test(b.textContent));
  click(aboveBtn);
  const s3 = state();
  ok(s3.items.length === before + 1, 'insert above adds an item');
  const anchorIdx = s3.items.findIndex(i => i.id === anchorId);
  const fresh = s3.items[anchorIdx - 1];
  ok(fresh && fresh.feature === '' && fresh.startDay == null, 'new item sits immediately above the anchor');
  ok(fresh.holdPos === true, 'inserted item carries holdPos until a date is set');
  ok(!doc.querySelector('#panel .p-name'), 'insert does not open an item in the edit panel');
  ok(!!doc.querySelector('#panel .p-empty'), 'persistent panel shows its no-selection state');
  const focused = doc.activeElement;
  ok(focused && focused.classList.contains('r-name') &&
    focused.closest('.row.item').dataset.id === fresh.id,
    'focus lands on the new row\'s inline title input');
  window.dispatchEvent(new window.KeyboardEvent('keydown', { key: 'z', metaKey: true, bubbles: true }));
}

// ---------------------------------------------------------------- clear all deps (Edit menu)
{
  window.eval("document.querySelector('[data-menu=\"edit\"]').click()");
  const clearBtn = Array.from(doc.querySelectorAll('#popover .menu-list button')).find(b => /Clear all dependencies/.test(b.textContent));
  ok(!!clearBtn, 'Edit menu offers Clear all dependencies');
  click(clearBtn);
  const confirmBtn = Array.from(doc.querySelectorAll('#modalHost button')).find(b => /^Clear$/.test(b.textContent.trim()));
  ok(!!confirmBtn, 'clearing asks for confirmation');
  click(confirmBtn);
  ok(state().items.every(i => i.deps.length === 0 && (i.depsText || []).length === 0), 'all dependencies cleared');
  window.dispatchEvent(new window.KeyboardEvent('keydown', { key: 'z', metaKey: true, bubbles: true }));
  ok(state().items.some(i => i.deps.length > 0), 'undo restores the dependency links');
}

// ---------------------------------------------------------------- rich text editors
{
  const withStories = state().items.find(i => i.stories.length);
  click(doc.querySelector('#rows .row.item[data-id="' + withStories.id + '"] .row-left'));
  ok(!doc.querySelector('#panel').hidden, 'panel opens for a storied item');

  // feature description is a WYSIWYG editor committing sanitized HTML
  const ed = doc.querySelector('#panel .wz-ed[data-f="col:description"]');
  ok(!!ed && ed.getAttribute('contenteditable') === 'true', 'feature description is a rich editor');
  ed.innerHTML = 'Hello <i>world</i><script>evil()</script>';
  ed.dispatchEvent(new window.FocusEvent('focusout', { bubbles: true }));
  ok(state().items.find(i => i.id === withStories.id).description === 'Hello <i>world</i>',
    'rich description commits sanitized on blur');

  // stories: the edit affordance opens the story panel with rich fields
  const stBtn = doc.querySelector('#panel [data-pst-edit]');
  ok(!!stBtn, 'panel stories offer an edit affordance');
  click(stBtn);
  ok(!!doc.querySelector('#panel .p-crumb'), 'story panel opens with a parent breadcrumb');
  ok(!!doc.querySelector('#panel .p-rollup'), 'story panel shows rolled-up workstream/epic');
  {
    const sjk = doc.querySelector('#panel input[data-stf="jiraKey"]');
    ok(!!sjk, 'story panel has a Jira key input');
    sjk.value = 'hw-8';
    sjk.dispatchEvent(new window.Event('change', { bubbles: true }));
    ok(state().items.find(i => i.id === withStories.id).stories[0].jiraKey === 'HW-8', 'story Jira key commits');
  }
  const sd = doc.querySelector('#panel .wz-ed[data-f="stcol:description"]');
  const sa = doc.querySelector('#panel .wz-ed[data-f="stcol:ac"]');
  ok(!!sd && !!sa && sd.getAttribute('contenteditable') === 'true' && sa.getAttribute('contenteditable') === 'true',
    'story panel has rich Description and Acceptance criteria editors');
  sd.innerHTML = 'Does <b>things</b>';
  sd.dispatchEvent(new window.FocusEvent('focusout', { bubbles: true }));
  // the commit re-rendered the panel — re-query the AC editor
  const sa2 = doc.querySelector('#panel .wz-ed[data-f="stcol:ac"]');
  sa2.innerHTML = '<ul><li>works offline</li></ul>';
  sa2.dispatchEvent(new window.FocusEvent('focusout', { bubbles: true }));
  const st0 = state().items.find(i => i.id === withStories.id).stories[0];
  ok(st0.description === 'Does <b>things</b>', 'story description saves');
  ok(st0.ac === '<ul><li>works offline</li></ul>', 'story acceptance criteria save');
  // a custom scope column edits into st.custom
  const scEd = doc.querySelector('#panel .wz-ed[data-f^="stcol:"]:not([data-f="stcol:description"]):not([data-f="stcol:ac"])');
  if (scEd) {
    scEd.innerHTML = 'story-scoped value';
    scEd.dispatchEvent(new window.FocusEvent('focusout', { bubbles: true }));
    const key = scEd.dataset.f.slice(6);
    ok(state().items.find(i => i.id === withStories.id).stories[0].custom[key] === 'story-scoped value',
      'story custom field commits to st.custom');
  }
  // breadcrumb returns to the parent item
  click(doc.querySelector('#panel .p-crumb'));
  ok(!doc.querySelector('#panel .p-crumb') && !!doc.querySelector('#panel .p-num'),
    'breadcrumb returns to the item panel');

  // scoping description cell renders/edits the rich value (column added on demand)
  click(doc.querySelector('#viewTabs [data-view="scoping"]'));
  if (!doc.querySelector('#rows [data-scope="description"]')) {
    click(doc.querySelector('#hdrSprints [data-coladd]'));
    click(Array.from(doc.querySelectorAll('#popover .menu-list button')).find(b => /Description/.test(b.textContent)));
  }
  const cell = doc.querySelector('#rows .row.item[data-id="' + withStories.id + '"] [data-scope="description"]');
  ok(!!cell && cell.getAttribute('contenteditable') === 'true', 'scoping description cell is rich');
  ok(/Hello/.test(cell.textContent), 'scoping description shows the committed rich text');
  cell.innerHTML = 'From <b>scoping</b>';
  cell.dispatchEvent(new window.FocusEvent('focusout', { bubbles: true }));
  ok(state().items.find(i => i.id === withStories.id).description === 'From <b>scoping</b>',
    'scoping description commits rich text');
  click(doc.querySelector('#viewTabs [data-view="planning"]'));
  window.__headway.getState && doc.querySelector('#panel [data-f=close]') && click(doc.querySelector('#panel [data-f=close]'));
}

// ---------------------------------------------------------------- persistent panel
{
  click(doc.querySelector('#viewTabs [data-view="planning"]'));
  window.dispatchEvent(new window.KeyboardEvent('keydown', { key: 'Escape', bubbles: true }));
  ok(!doc.querySelector('#panel').hidden && !!doc.querySelector('#panel .p-empty'),
    'panel persists on Planning with a no-selection state');
  click(doc.querySelector('#panel [data-f="collapse"]'));
  ok(doc.querySelector('#panel').hidden && !doc.querySelector('#panelPeek').hidden,
    'collapsing hides the panel and shows the peek handle');
  click(doc.querySelector('#panelPeek'));
  ok(!doc.querySelector('#panel').hidden, 'the peek handle reopens the panel');
  click(doc.querySelector('#viewTabs [data-view="scoping"]'));
  ok(!doc.querySelector('#panel').hidden, 'the panel is available on Scoping too');
  click(doc.querySelector('#viewTabs [data-view="budget"]'));
  ok(doc.querySelector('#panel').hidden, 'the panel does not render on Budgeting');
  click(doc.querySelector('#viewTabs [data-view="planning"]'));
}

// ---------------------------------------------------------------- duration preset off-timeline
{
  // give an unscheduled item an explicit duration, then place it — the
  // preset must win over the size estimate
  const unsched = state().items.find(i => i.startDay == null && !i.milestone);
  if (unsched) {
    click(doc.querySelector('#rows .row.item[data-id="' + unsched.id + '"] .r-num'));
    const durInp = doc.querySelector('#panel [data-f="durWeeks"]');
    ok(!!durInp && durInp.value === '', 'unscheduled items offer an empty Duration field');
    durInp.value = '3';
    durInp.dispatchEvent(new window.Event('change', { bubbles: true }));
    ok(state().items.find(i => i.id === unsched.id).durDays === 15,
      'duration can be set while off the timeline');
    click(doc.querySelector('#panel [data-f="schedule-now"]'));
    const placed = state().items.find(i => i.id === unsched.id);
    ok(placed.startDay != null && placed.durDays === 15, 'placing respects the preset duration');
    window.dispatchEvent(new window.KeyboardEvent('keydown', { key: 'z', metaKey: true, bubbles: true }));
    window.dispatchEvent(new window.KeyboardEvent('keydown', { key: 'z', metaKey: true, bubbles: true }));
    window.dispatchEvent(new window.KeyboardEvent('keydown', { key: 'Escape', bubbles: true }));
  } else ok(true, '(no unscheduled item in seed)');
}

// ---------------------------------------------------------------- milestones (UI)
{
  const anyRow = doc.querySelector('#rows .row.item');
  const msId = anyRow.dataset.id;
  anyRow.dispatchEvent(new window.MouseEvent('contextmenu', { bubbles: true, clientX: 240, clientY: 240 }));
  const convBtn = Array.from(doc.querySelectorAll('#popover .menu-list button')).find(b => /Convert to milestone/.test(b.textContent));
  ok(!!convBtn, 'row context menu offers Convert to milestone');
  click(convBtn);
  const msIt = state().items.find(i => i.id === msId);
  ok(msIt.milestone === true && (msIt.startDay == null || msIt.durDays === 0),
    'converting makes a zero-duration milestone');
  if (msIt.startDay != null) {
    ok(!!doc.querySelector('#rows .bar.ms[data-bar="' + msId + '"] .ms-diamond'),
      'milestones render as diamonds on the timeline');
  }
  window.dispatchEvent(new window.KeyboardEvent('keydown', { key: 'z', metaKey: true, bubbles: true }));
  ok(state().items.find(i => i.id === msId).milestone === false, 'undo restores the feature');
}

// ---------------------------------------------------------------- sizing approaches
{
  click(doc.querySelector('#btnSetup'));
  suTab('est');
  ok(schemeOpts('size', 'feature').length >= 4, 'Sizing offers approach presets');
  pickScheme('size', 'feature', 'fibonacci');
  ok(state().meta.sizeScheme === 'fibonacci' &&
    state().meta.sizeOrder.join(',') === '0.5,1,2,3,5,8,13',
    'Story points preset applies its scale');
  ok(doc.querySelectorAll('#setupView [data-susz]:not([data-kind])').length === 7, 'option table lists the seven point values');
  // rename an option — items follow, scheme flips to custom
  const lblInp = doc.querySelector('#setupView [data-suszlabel="13"]');
  lblInp.value = '21';
  lblInp.dispatchEvent(new window.Event('change', { bubbles: true }));
  ok(state().meta.sizeScheme === 'custom' && state().meta.sizeOrder.indexOf('21') !== -1,
    'editing options flips the approach to Custom');
  pickScheme('size', 'feature', 'none');
  ok(state().meta.sizeScheme === 'none' && state().meta.sizeOrder.length === 0, 'No sizing empties the scale');
  click(doc.querySelector('#viewTabs [data-view="planning"]'));
  ok(!doc.querySelector('#rows .row.item [data-act="size"]'), 'no size chips while sizing is off');
  ok(doc.body.classList.contains('no-size'), 'body carries the no-size flag');
  click(doc.querySelector('#btnSetup'));
  pickScheme('size', 'feature', 'tshirt');
  ok(state().meta.sizeOrder.join(',') === 'XS,S,M,L,XL', 'T-shirt preset restores the classic scale');
}
{
  suTab('est');
  const cells = doc.querySelectorAll('#setupView .su-grid select[data-suscheme-kind]');
  ok(cells.length === 6, 'grid: size, priority, risk × feature, story');
  ok(schemeOpts('prio', 'story').indexOf('rice') === -1, 'RICE is not offered for stories');
  ok(schemeOpts('risk', 'story').indexOf('auto') === -1, 'Risk (auto) is not offered for stories');
  const riskBefore = state().meta.storyRiskScheme;
  pickScheme('risk', 'story', 'confidence');
  ok(state().meta.storyRiskScheme === 'confidence', 'story risk scheme is set from the grid');
  ok(doc.querySelectorAll('#setupView .su-chip').length > 0, 'priority / risk levels show as read-only chips');
  ok(!doc.querySelector('#setupView .su-chip input'), 'chips are not editable');
  ok(!!doc.querySelector('#setupView [data-susz]'), 'size options stay editable');
  pickScheme('risk', 'story', riskBefore);
}

// ---------------------------------------------------------------- workstream feature toggle
{
  suTab('org');
  const wsChk = doc.querySelector('#suWsEnable');
  ok(!!wsChk && wsChk.checked, 'Workstreams tab offers the feature switch (on by default)');
  wsChk.checked = false;
  wsChk.dispatchEvent(new window.Event('change', { bubbles: true }));
  ok(state().meta.workstreamsEnabled === false, 'workstreams can be disabled per project');
  click(doc.querySelector('#viewTabs [data-view="scoping"]'));
  ok(!doc.querySelector('#hdrSprints [data-col="workstream"]'), 'Scoping hides the Workstream column when off');
  click(doc.querySelector('#btnSetup'));
  suTab('org');
  const wsChk2 = doc.querySelector('#suWsEnable');
  wsChk2.checked = true;
  wsChk2.dispatchEvent(new window.Event('change', { bubbles: true }));
  ok(state().meta.workstreamsEnabled === true, 'workstreams re-enable');
  click(doc.querySelector('#viewTabs [data-view="planning"]'));
}

// ---------------------------------------------------------------- empty workstream = default workstream
{
  const sched = state().items.find(i => i.startDay != null && i.workstream && !i.milestone);
  if (sched) {
    click(doc.querySelector('#rows .row.item[data-id="' + sched.id + '"] .r-num'));
    // the panel's workstream control is a dropdown now — pick the default row
    const wsBtn = doc.querySelector('#panel [data-dd="ws"]');
    ok(!!wsBtn, 'the panel offers a Workstream dropdown (color dot, not a text input)');
    click(wsBtn);
    const wsMenu = doc.querySelector('#popover .menu-list');
    ok(!!wsMenu && /default/.test(wsMenu.textContent), 'the workstream dropdown lists the default stream');
    click(wsMenu.querySelector('[data-mi="0"]'));
    const bar = doc.querySelector('#rows .bar[data-bar="' + sched.id + '"]');
    ok(bar && !bar.classList.contains('nows') &&
      bar.getAttribute('style').includes(window.RM.defaultWsColor(state())),
      'a bar with no workstream paints in the default workstream color');
    const dot = doc.querySelector('#rows .row.item[data-id="' + sched.id + '"] .r-dot');
    ok(dot && !dot.classList.contains('nodot'), 'its left-pane dot stays filled');
    window.dispatchEvent(new window.KeyboardEvent('keydown', { key: 'z', metaKey: true, bubbles: true }));
    window.dispatchEvent(new window.KeyboardEvent('keydown', { key: 'Escape', bubbles: true }));
  } else ok(true, '(no scheduled workstreamed item in seed)');
}

// ---------------------------------------------------------------- scoping: no placeholder dots
{
  click(doc.querySelector('#viewTabs [data-view="scoping"]'));
  const empties = Array.from(doc.querySelectorAll('#rows .sc-fix .r-ws, #rows .sc-fix .r-size'))
    .filter(el => el.textContent.trim() === '·');
  ok(empties.length === 0, 'empty scoping chips show no placeholder dots');
  click(doc.querySelector('#viewTabs [data-view="planning"]'));
}

// ---------------------------------------------------------------- png export UI
{
  const tbRight = doc.querySelector('#topbar .tb-right');
  const kids = Array.from(tbRight.querySelectorAll('button')).map(b => b.id);
  ok(kids.indexOf('btnExport') !== -1 && kids.indexOf('btnExport') < kids.indexOf('btnSave'),
    'Export sits before Save');
  ok(!!doc.querySelector('#btnSetup') && !doc.querySelector('#viewTabs [data-view="setup"]'),
    'Setup is its own icon button beside the view tabs');
  ok(!!doc.querySelector('#viewTabs [data-view="reports"]'), 'Reports tab sits in the view tabs');

  window.eval("document.querySelector('#btnExport').click()");
  ok(!doc.querySelector('#modalHost').hidden, 'Export opens a dialog');
  ok(!!doc.querySelector('#modalHost #exFrom') && !!doc.querySelector('#modalHost #exTo'),
    'dialog offers a range');
  ok(/Date range/.test(doc.querySelector('#modalHost .modal').textContent) &&
    !/Sprint range/.test(doc.querySelector('#modalHost .modal').textContent),
    'range section is titled Date range');
  const wsOpts = Array.from(doc.querySelectorAll('#modalHost #exWs option')).map(o => o.textContent);
  ok(wsOpts.length > 1 && wsOpts.some(t => /Product/.test(t)), 'dialog lists workstreams to filter by');
  ok(!!doc.querySelector('#modalHost #exEpic') && !!doc.querySelector('#modalHost #exPhase'),
    'dialog offers epic and phase filters');
  ok(!!doc.querySelector('#modalHost #exFmtPng') && !!doc.querySelector('#modalHost #exFmtPptx'),
    'dialog offers PNG and PPTX formats');
  ok(!!doc.querySelector('#modalHost #exSplitPhase') && !!doc.querySelector('#modalHost #exSplitWs'),
    'dialog offers per-phase and per-workstream splitting');
  ok(!!doc.querySelector('#modalHost #exGroupWs') && !!doc.querySelector('#modalHost #exGroupEpic'),
    'dialog offers timeline-style grouping');
  ok(/Export/.test(doc.querySelector('#modalHost .m-head h2').textContent) &&
    !/Export PNG/.test(doc.querySelector('#modalHost .m-head h2').textContent),
    'dialog is titled Export, format is a choice inside it');
  ok(typeof window.RM_PPTX === 'object' && typeof window.RM_PPTX.toBlob === 'function',
    'PPTX export module is loaded in the page');
  ok(typeof window.RM_EXPORT.plan === 'function' &&
    window.RM_EXPORT.plan(state(), { byPhase: true }).length >= 1,
    'export plan splits the live document');
  // export settings persist between sessions (via the ui snapshot)
  doc.querySelector('#modalHost #exFmtPptx').checked = true;
  doc.querySelector('#modalHost #exSplitPhase').checked = true;
  doc.querySelector('#modalHost #exArrows').checked = true;
  click(doc.querySelector('#modalHost #exGo'));
  const savedPrefs = JSON.parse(window.localStorage.getItem('headway-ui-v1')).exportPrefs;
  ok(savedPrefs && savedPrefs.fmt === 'pptx' && savedPrefs.byPhase === true && savedPrefs.arrows === true,
    'exporting writes the chosen settings to the ui snapshot');
  window.eval("document.querySelector('#btnExport').click()");
  ok(doc.querySelector('#modalHost #exFmtPptx').checked &&
    doc.querySelector('#modalHost #exSplitPhase').checked &&
    doc.querySelector('#modalHost #exArrows').checked,
    'reopening the export dialog restores the saved settings');
  const lay = window.RM_EXPORT.layout(state(), {});
  ok(lay.rows.filter(r => r.kind === 'item').length ===
    state().items.filter(i => i.startDay != null).length,
    'export layout covers every scheduled item of the live document');
  click(doc.querySelector('#modalHost [data-m=cancel], #modalHost [data-m=x]'));
  ok(doc.querySelector('#modalHost').hidden, 'export dialog closes');
}

// hovering an unscheduled row's lane shows ONLY the dashed place ghost —
// the old low-opacity preview bar is gone
{
  const unRow = [...doc.querySelectorAll('#rows .row.item')].find(r => {
    const it = state().items.find(i => i.id === r.dataset.id);
    return it && it.startDay == null;
  });
  if (unRow) {
    const lane = unRow.querySelector('.row-lane');
    lane.dispatchEvent(new window.MouseEvent('mousemove', { bubbles: true, clientX: 400 }));
    ok(!doc.querySelector('.place-preview'),
      'hovering an unscheduled lane paints no low-opacity preview bar');
  } else {
    ok(true, 'no unscheduled row rendered to hover (skipped)');
  }
}

// ---------------------------------------------------------------- batch 4: toasts, tooltips, filter, titles, schemes
{
  // shadcn-style toast: icon + message span, in the #toasts stack
  window.eval("window.HeadwayApp.toast ? window.HeadwayApp.toast('hello toast') : (function(){ })()");
  // fall back to triggering one through a real action if not exposed
  const anyToast = doc.querySelector('#toasts .toast');
  if (anyToast) {
    ok(!!anyToast.querySelector('.toast-msg'), 'toasts carry a message span (icon + text layout)');
  } else ok(true, '(no toast surfaced to inspect)');

  // filter input: no dots in the label, kbd suffix chip present
  ok(doc.querySelector('#rowFilter').placeholder === 'Filter rows', 'filter placeholder has no ellipsis');
  ok(!!doc.querySelector('#filterCell kbd.filter-kbd'), 'kbd-style shortcut chip inside the filter input');

  // tooltip component: hovering a [title] element lifts it to data-tip
  const titled = doc.querySelector('#btnExport');
  ok(titled.hasAttribute('title'), 'buttons still author titles');
  titled.dispatchEvent(new window.Event('pointerover', { bubbles: true }));
  ok(!titled.hasAttribute('title') && !!titled.getAttribute('data-tip'),
    'hover lifts title into data-tip so the native tooltip never shows');
}

// scoping: the title is a full-height editable cell that commits on blur
{
  click(doc.querySelector('#viewTabs [data-view="scoping"]'));
  const nameCell = doc.querySelector('#rows .row.item .sc-name');
  ok(!!nameCell && nameCell.getAttribute('contenteditable') === 'true',
    'scoping title renders as a contenteditable cell');
  const rowId = nameCell.closest('.row').dataset.id;
  nameCell.textContent = 'Renamed via cell';
  nameCell.dispatchEvent(new window.FocusEvent('focusout', { bubbles: true }));
  ok(state().items.find(i => i.id === rowId).feature === 'Renamed via cell',
    'blurring the title cell commits the rename');
  window.dispatchEvent(new window.KeyboardEvent('keydown', { key: 'z', metaKey: true, bubbles: true }));
}

// default workstream: setup row + modal rename/recolor
{
  suTab('org');
  ok(!!doc.querySelector('#setupView .su-defws'), 'default workstream row leads the Workstreams tab');
  click(doc.querySelector('#setupView [data-sudefws]'));
  ok(!!doc.querySelector('#modalHost #dwsName'), 'default workstream modal opens');
  ok(!!doc.querySelector('#modalHost .swatch[data-esw="6E7883"]'),
    'default workstream picker offers gray');
  doc.querySelector('#modalHost #dwsName').value = 'Core';
  click(doc.querySelector('#modalHost #dwsSave'));
  ok(state().meta.defaultWsName === 'Core', 'default workstream renames');
  ok(window.RM.defaultWsName(state()) === 'Core', 'core helper reads the rename');
  window.dispatchEvent(new window.KeyboardEvent('keydown', { key: 'z', metaKey: true, bubbles: true }));
}

// risk scheme card: switching to MoSCoW relabels the scoping column
{
  suTab('est');
  ok(schemeOpts('risk', 'feature').length === 4, 'four risk schemes offered (none, risk, auto, confidence)');
  pickScheme('risk', 'feature', 'confidence');
  ok(state().meta.riskScheme === 'confidence', 'Confidence scheme commits');
  window.dispatchEvent(new window.KeyboardEvent('keydown', { key: 'z', metaKey: true, bubbles: true }));
  // scheme none removes the column
  suTab('est');
  pickScheme('risk', 'feature', 'none');
  ok(state().meta.riskScheme === 'none', 'scheme none commits');
  click(doc.querySelector('#viewTabs [data-view="scoping"]'));
  ok(!doc.querySelector('#rows .row.item .r-risk'), 'no assessment chips when the scheme is none');
  window.dispatchEvent(new window.KeyboardEvent('keydown', { key: 'z', metaKey: true, bubbles: true }));
}

// role rename from Setup propagates everywhere
{
  suTab('budget');
  const nameInp = doc.querySelector('#setupView [data-rcname]');
  ok(!!nameInp, 'role names are editable inputs');
  const oldRole = nameInp.dataset.rcname;
  nameInp.value = 'Renamed Role';
  nameInp.dispatchEvent(new window.Event('change', { bubbles: true }));
  ok(state().teamTypes.includes('Renamed Role') && !state().teamTypes.includes(oldRole),
    'renaming a role updates the role list');
  ok(!state().team.some(m => m.type === oldRole) && !state().items.some(i => i.teamType === oldRole),
    'people and items follow the role rename');
  window.dispatchEvent(new window.KeyboardEvent('keydown', { key: 'z', metaKey: true, bubbles: true }));
}

// holiday quick-edit: right-clicking an empty planning lane offers holiday actions
{
  click(doc.querySelector('#viewTabs [data-view="planning"]'));
  const lane = doc.querySelector('#rows .row.item .row-lane');
  lane.dispatchEvent(new window.MouseEvent('contextmenu', { bubbles: true, clientX: 400, clientY: 120 }));
  const menuText = doc.querySelector('#popover').textContent;
  ok(/holiday/i.test(menuText) && /Holiday settings/.test(menuText),
    'lane right-click opens the holiday quick-edit menu');
  window.dispatchEvent(new window.KeyboardEvent('keydown', { key: 'Escape', bubbles: true }));
}

// export renders through toBlob (folder choice happens in the picker/dialog)
ok(typeof window.RM_EXPORT.toBlob === 'function', 'PNG export exposes a blob renderer for save-to-folder flows');

// ---------------------------------------------------------------- prioritizing view
{
  const tabs = [...doc.querySelectorAll('#viewTabs [data-view]')].map(b => b.dataset.view);
  ok(tabs.indexOf('sprints') === tabs.indexOf('planning') + 1, 'Sprinting sits right of Planning');
  ok(tabs.indexOf('prio') === tabs.indexOf('scoping') + 1, 'Prioritizing sits right of Scoping');
  click(doc.querySelector('#viewTabs [data-view="prio"]'));
  ok(doc.body.dataset.view === 'prio', 'view switches to Prioritizing');
  ok(doc.querySelectorAll('#prioView .sp-col').length === state().phases.length,
    'kanban columns are the phases');
  ok(!doc.querySelector('#prVision'), 'the vision line is gone');
  ok(doc.querySelectorAll('#prioView .pr-phhd').length === state().phases.length,
    'the header row names every phase');
  // cards are compact by default: title + chips, no # label, no extra fields
  const card = doc.querySelector('#prioView .pr-card');
  ok(!!card, 'cards render on the board');
  ok(!card.querySelector('.pr-rich') && !card.querySelector('.r-num'),
    'cards are compact by default — no fields, no # label');
  ok(!!card.querySelector('[data-pract="size"]') && !!card.querySelector('[data-pract="epic"]') &&
    !!card.querySelector('[data-pract="ws"]'),
    'cards carry size, epic and workstream chips');
  ok((card.getAttribute('style') || '').indexOf('--ws-c') !== -1, 'cards are tinted by workstream');
  // filter input matches the header filter: search icon + ⌘F chip
  const pf = doc.querySelector('#prFilter');
  ok(!!pf && !!pf.closest('.filter-wrap') &&
    !!pf.closest('.filter-wrap').querySelector('.filter-kbd'),
    'the card filter carries the search icon and ⌘F chip');
  // right-clicking a card offers move / epic / workstream / delete
  card.dispatchEvent(new window.MouseEvent('contextmenu', { bubbles: true, clientX: 10, clientY: 10 }));
  const ctxLabels = [...doc.querySelectorAll('#popover .menu-list button')].map(b => b.textContent);
  ok(ctxLabels.some(l => /Move to phase/.test(l)) && ctxLabels.some(l => /Set workstream/.test(l)) &&
    ctxLabels.some(l => /Set epic/.test(l)) && ctxLabels.some(l => /Delete/.test(l)),
    'right-clicking a card offers move, epic, workstream and delete');
  window.dispatchEvent(new window.KeyboardEvent('keydown', { key: 'Escape', bubbles: true }));
  // the Fields menu opts cards into scope columns; Description gets the editor
  click(doc.querySelector('#prFieldsBtn'));
  const descOpt = [...doc.querySelectorAll('#popover .menu-list button')].find(b => /Description/.test(b.textContent));
  ok(!!descOpt && !descOpt.classList.contains('on'), 'Fields menu lists Description, off by default');
  click(descOpt);
  const richEd = doc.querySelector('#prioView .pr-rich[data-prsc="description"]');
  ok(!!richEd && richEd.dataset.ph === 'No description',
    'checking Description adds a label-free field with a quiet "No description" hint');
  ok(!richEd.isContentEditable && richEd.getAttribute('contenteditable') == null, 'card fields are read-only (the panel edits them)');
  click(doc.querySelector('#prFieldsBtn'));
  click([...doc.querySelectorAll('#popover .menu-list button')].find(b => /Description/.test(b.textContent)));
  ok(!doc.querySelector('#prioView .pr-rich'), 'unchecking returns cards to compact');
  // the same menu hides the card chips (duration, epic, workstream, ...)
  ok(!!doc.querySelector('#prioView .pr-card [data-pract="dur"]') && !!doc.querySelector('#prioView .pr-card [data-pract="epic"]'),
    'cards show the duration and epic chips by default');
  click(doc.querySelector('#prFieldsBtn'));
  {
    const labels = [...doc.querySelectorAll('#popover .menu-list button')].map(b => b.textContent.trim());
    ok(['Duration', 'Epic', 'Workstream'].every(l => labels.includes(l)) && !!doc.querySelector('#popover .menu-sep'),
      'Fields menu lists Duration, Epic and Workstream after a separator');
    const durOpt = [...doc.querySelectorAll('#popover .menu-list button')].find(b => b.textContent.trim() === 'Duration');
    ok(durOpt.classList.contains('on'), 'chip entries start checked');
    click(durOpt);
  }
  ok(!doc.querySelector('#prioView .pr-card [data-pract="dur"]') && !!doc.querySelector('#prioView .pr-card [data-pract="epic"]') &&
    JSON.parse(window.localStorage.getItem('headway-ui-v1')).prioChipHide.includes('dur'),
    'unchecking Duration drops that chip from every card and persists');
  click(doc.querySelector('#prFieldsBtn'));
  click([...doc.querySelectorAll('#popover .menu-list button')].find(b => b.textContent.trim() === 'Epic'));
  ok(!doc.querySelector('#prioView .pr-card [data-pract="epic"]'), 'unchecking Epic hides the epic chip too');
  click(doc.querySelector('#prFieldsBtn'));
  click([...doc.querySelectorAll('#popover .menu-list button')].find(b => b.textContent.trim() === 'Epic'));
  click(doc.querySelector('#prFieldsBtn'));
  click([...doc.querySelectorAll('#popover .menu-list button')].find(b => b.textContent.trim() === 'Duration'));
  ok(!!doc.querySelector('#prioView .pr-card [data-pract="dur"]') && !!doc.querySelector('#prioView .pr-card [data-pract="epic"]'),
    'rechecking restores both chips');
  // feature level: the Columns dropdown swaps phases for a feature field's ladder
  {
    const heads = () => [...doc.querySelectorAll('#prioView .pr-phhd')].map(h => h.firstChild.textContent.trim());
    const colsBtn = () => doc.querySelector('#prioView [data-prdd="cols"]');
    ok(!!colsBtn() && /Phase/.test(colsBtn().textContent), 'feature level offers a Columns dropdown reading Phase');
    click(colsBtn());
    click([...doc.querySelectorAll('#popover .menu-list button')].find(b => /Size/.test(b.textContent)));
    const so = window.RM.sizeOrderOf(state());
    ok(heads().length === so.length + 1 && heads()[0] === so[0] && heads()[heads().length - 1] === 'Unset' &&
      JSON.parse(window.localStorage.getItem('headway-ui-v1')).prioFeatCol === 'size',
      'Columns → Size lays features out by size plus Unset and persists');
    ok(doc.querySelectorAll('#prioView .pr-card').length > 0 && !doc.querySelector('#prioView .pr-card [data-pract="size"]'),
      'the size chip leaves the cards');
    ok(!doc.querySelector('#prioView .pr-add'), 'Add hides while the columns are not phases');
    const fcard = doc.querySelector('#prioView .pr-card');
    const fid = fcard.dataset.prcard;
    const sizeOf = () => state().items.find(i => i.id === fid).size || '';
    const was = sizeOf();
    const tgtSize = so.find(v => v !== was);
    fcard.dispatchEvent(new window.MouseEvent('pointerdown', { bubbles: true, button: 0, clientX: 10, clientY: 10 }));
    doc.querySelector('#prioView .sp-col[data-prcol="' + tgtSize + '"]').dispatchEvent(new window.MouseEvent('pointermove', { bubbles: true, clientX: 80, clientY: 80 }));
    window.dispatchEvent(new window.MouseEvent('pointerup', { bubbles: true, clientX: 80, clientY: 80 }));
    doc.querySelector('#prioView .pr-table').dispatchEvent(new window.MouseEvent('click', { bubbles: true })); // the browser's trailing click
    ok(sizeOf() === tgtSize && !!doc.querySelector('#prioView .sp-col[data-prcol="' + tgtSize + '"] [data-prcard="' + fid + '"]'),
      'dragging a feature card to a size column sets its size');
    window.dispatchEvent(new window.KeyboardEvent('keydown', { key: 'z', metaKey: true, bubbles: true }));
    ok(sizeOf() === was, 'undo restores the size');
    click(colsBtn());
    click([...doc.querySelectorAll('#popover .menu-list button')].find(b => /Phase/.test(b.textContent)));
    ok(heads().length === state().phases.length, 'Columns → Phase brings the phase board back');
    ok([...doc.querySelectorAll('#prioView .sp-col')].every(c => !!c.querySelector('.pr-add') === !c.querySelector('.pr-card')),
      'the Add button sits only under empty phase columns');
  }
  // the Feature / Story dropdown leads the toolbar, like every other view
  {
    const prBar = doc.querySelector('#prioView .pr-bar');
    ok(prBar.firstElementChild.matches('.dm-btn[data-prdd="level"]') && /Feature/.test(prBar.firstElementChild.textContent),
      'Prioritizing: level dropdown sits at the left of the toolbar, reading Feature');
    ok(!doc.querySelector('#prioView .pr-story'), 'Feature level shows no story rows');
    click(doc.querySelector('#prioView [data-prdd="level"]'));
    click([...doc.querySelectorAll('#popover .menu-list button')].find(b => /Story/.test(b.textContent)));
    const withStories = state().items.find(i => (i.stories || []).length && !i.milestone);
    ok(/Story/.test(doc.querySelector('#prioView [data-prdd="level"]').textContent) &&
      JSON.parse(window.localStorage.getItem('headway-ui-v1')).prioLevel === 'story',
      'Story level toggles and persists with the view prefs');
    // story level is a board of STORY cards in priority columns
    const stCards = () => [...doc.querySelectorAll('#prioView .pr-stcard')];
    const allStories = state().items.filter(i => !i.milestone).reduce((a, i) => a.concat((i.stories || []).map(st => ({ it: i, st }))), []);
    ok(!doc.querySelector('#prioView .pr-card:not(.pr-stcard)') && allStories.length > 0 && stCards().length === allStories.length,
      'Story level shows one card per story and no feature cards (' + stCards().length + ')');
    const priOrder = window.RM.priorityOrderOf(state(), 'story');
    const heads = () => [...doc.querySelectorAll('#prioView .pr-phhd')].map(h => h.firstChild.textContent.trim());
    ok(heads().length === priOrder.length + 1 && heads()[0] === 'Critical' && heads()[heads().length - 1] === 'Unset',
      'columns are the story priority ladder plus Unset');
    const card0 = stCards()[0];
    const fid = card0.dataset.prcard, sid = card0.dataset.prst;
    const feat0 = state().items.find(i => i.id === fid);
    ok(!!card0.querySelector('.pr-stfeat') && card0.querySelector('.pr-stfeat').textContent.indexOf(feat0.feature) !== -1,
      'each story card names its feature');
    ok(!!card0.querySelector('[data-prstact="st-pri"]') && !!card0.querySelector('[data-prstact="st-size"]'),
      'Priority and Risk chips stay on the card even when the columns are that field');
    ok(!doc.querySelector('#prFieldsBtn') && !doc.querySelector('#prioView .pr-add') && !doc.querySelector('#prioView [data-prstadd]'),
      'Fields, Add and Add story leave in story mode');
    const stCardOf = () => doc.querySelector('#prioView .pr-stcard[data-prst="' + sid + '"]');
    ok(stCardOf().querySelector('.pr-title').tagName !== 'INPUT', 'a story card title is text, not an input');
    click(stCardOf().querySelector('.pr-title'));
    ok(!stCardOf().querySelector('input'), 'a single click on a story card title does not start an edit');
    stCardOf().querySelector('.pr-title').dispatchEvent(new window.MouseEvent('dblclick', { bubbles: true }));
    const stIn = stCardOf().querySelector('input[data-prstf="title"]');
    ok(!!stIn, 'double-click makes the story card title editable');
    stIn.value = 'Renamed on the board';
    stIn.dispatchEvent(new window.KeyboardEvent('keydown', { key: 'Enter', bubbles: true }));
    ok(state().items.find(i => i.id === fid).stories.find(x => x.id === sid).title === 'Renamed on the board',
      'a story title edits inline on the card');
    // the card menu starts the same edit
    stCardOf().dispatchEvent(new window.MouseEvent('contextmenu', { bubbles: true, cancelable: true, clientX: 90, clientY: 90 }));
    const stRn = [...doc.querySelectorAll('#popover .menu-list button')].find(b => /Rename/.test(b.textContent));
    ok(!!stRn, 'a story card context menu offers Rename…');
    click(stRn);
    ok(!!stCardOf().querySelector('input[data-prstf="title"]'), 'Rename… opens the story card editor');
    stCardOf().querySelector('input[data-prstf="title"]')
      .dispatchEvent(new window.KeyboardEvent('keydown', { key: 'Escape' }));
    window.dispatchEvent(new window.KeyboardEvent('keydown', { key: 'z', metaKey: true, bubbles: true }));
    // drag a story card into the High column: its priority follows
    const storyOf = () => state().items.find(i => i.id === fid).stories.find(x => x.id === sid);
    const dragCardTo = (colKey) => {
      const c = doc.querySelector('#prioView .pr-stcard[data-prst="' + sid + '"]');
      const col = doc.querySelector('#prioView .sp-col[data-prcol="' + colKey + '"]');
      c.dispatchEvent(new window.MouseEvent('pointerdown', { bubbles: true, button: 0, clientX: 10, clientY: 10 }));
      col.dispatchEvent(new window.MouseEvent('pointermove', { bubbles: true, clientX: 80, clientY: 80 }));
      window.dispatchEvent(new window.MouseEvent('pointerup', { bubbles: true, clientX: 80, clientY: 80 }));
      // a real browser fires a click after the pointerup; the view swallows it
      doc.querySelector('#prioView .pr-table').dispatchEvent(new window.MouseEvent('click', { bubbles: true }));
    };
    const before = storyOf().priority || null;
    const tgt = priOrder.find(v => v !== before); // a column the story is not in, so the drop commits
    dragCardTo(tgt);
    ok(storyOf().priority === tgt, 'dragging a story card to a column sets its priority');
    ok(!!doc.querySelector('#prioView .sp-col[data-prcol="' + tgt + '"] .pr-stcard[data-prst="' + sid + '"]'), 'and the card now sits in that column');
    dragCardTo('');
    ok(storyOf().priority == null && !!doc.querySelector('#prioView .sp-col[data-prcol=""] .pr-stcard[data-prst="' + sid + '"]'),
      'dropping on Unset clears the priority');
    window.dispatchEvent(new window.KeyboardEvent('keydown', { key: 'z', metaKey: true, bubbles: true }));
    ok(storyOf().priority === tgt, 'undo steps the drop back');
    window.dispatchEvent(new window.KeyboardEvent('keydown', { key: 'z', metaKey: true, bubbles: true }));
    ok((storyOf().priority || null) === before, 'and again back to the original priority');
    // the Columns dropdown swaps the field the columns represent
    click(doc.querySelector('#prioView [data-prdd="cols"]'));
    click([...doc.querySelectorAll('#popover .menu-list button')].find(b => /Size/.test(b.textContent)));
    const sizeOrder = window.RM.sizeOrderOf(state(), 'story');
    ok(heads().length === sizeOrder.length + 1 && heads()[0] === sizeOrder[0] &&
      JSON.parse(window.localStorage.getItem('headway-ui-v1')).prioStoryCol === 'size',
      'Columns → Size lays the board out by story size and persists');
    ok(!!stCards()[0].querySelector('[data-prstact="st-pri"]') && !stCards()[0].querySelector('[data-prstact="st-size"]'),
      'the size chip leaves the cards and the priority chip returns');
    click(doc.querySelector('#prioView [data-prdd="cols"]'));
    click([...doc.querySelectorAll('#popover .menu-list button')].find(b => /Priority/.test(b.textContent)));
    ok(heads()[0] === 'Critical', 'Columns → Priority restores the priority ladder');
    click(doc.querySelector('#prioView [data-prdd="level"]'));
    click([...doc.querySelectorAll('#popover .menu-list button')].find(b => /Feature/.test(b.textContent)));
    ok(!doc.querySelector('#prioView .pr-stcard') && !!doc.querySelector('#prioView .pr-card'), 'back at Feature level the feature cards return');
  }
  // sort control: Title ordering applies inside a column
  click(doc.querySelector('#prioView [data-prdd="sort"]'));
  click([...doc.querySelectorAll('#popover .menu-list button')].find(b => /Title/.test(b.textContent)));
  const colTitles = [...doc.querySelector('#prioView .sp-col').querySelectorAll('.pr-title')]
    .map(i => i.classList.contains('pr-ph') ? '' : i.textContent); // an untitled card shows its placeholder
  ok(colTitles.length < 2 || colTitles.every((t, i) => i === 0 || colTitles[i - 1].localeCompare(t) <= 0),
    'Sort → Title orders a column alphabetically');
  click(doc.querySelector('#prioView [data-prdd="sort"]'));
  click([...doc.querySelectorAll('#popover .menu-list button')].find(b => /Priority/.test(b.textContent)));
  // epic filter narrows the board
  click(doc.querySelector('#prioView [data-prdd="fepic"]'));
  const epPick = [...doc.querySelectorAll('#popover .menu-list button')].filter(b => !/All epics|no epic/.test(b.textContent))[0];
  const epName = epPick.textContent.trim();
  click(epPick);
  const visibleCards = [...doc.querySelectorAll('#prioView .pr-card')];
  ok(visibleCards.length > 0 && visibleCards.every(c =>
    state().items.find(i => i.id === c.dataset.prcard).epic === epName),
    'the epic filter narrows the board to that epic');
  const fepicBtn = doc.querySelector('#prioView [data-prdd="fepic"]');
  ok(fepicBtn.classList.contains('pr-on') && fepicBtn.querySelector('.dd-label i, .dd-label svg') &&
    (fepicBtn.querySelector('.dd-label i, .dd-label svg').getAttribute('data-lucide') === (state().epicIcons[epName] || 'tag')),
    'an active epic filter lights up and wears the epic\'s icon');
  click(doc.querySelector('#prioView [data-prdd="fws"]'));
  const wsPick = [...doc.querySelectorAll('#popover .menu-list button')].filter(b => !/All workstreams|default/.test(b.textContent))[0];
  click(wsPick);
  const fwsBtn = doc.querySelector('#prioView [data-prdd="fws"]');
  ok(fwsBtn.classList.contains('pr-on') && !!fwsBtn.querySelector('.dd-label .dd-dot'),
    'an active workstream filter lights up and shows its color dot');
  click(doc.querySelector('#prioView [data-prdd="fws"]'));
  click([...doc.querySelectorAll('#popover .menu-list button')].find(b => /All workstreams/.test(b.textContent)));
  ok(!doc.querySelector('#prioView [data-prdd="fws"]').classList.contains('pr-on'), 'clearing the filter dims the button');
  click(doc.querySelector('#prioView [data-prdd="fepic"]'));
  click([...doc.querySelectorAll('#popover .menu-list button')].find(b => /All epics/.test(b.textContent)));
  // swimlanes: Group dropdown, labels in the leftmost column of one table
  click(doc.querySelector('#prioView [data-prdd="group"]'));
  click([...doc.querySelectorAll('#popover .menu-list button')].find(b => /Epics/.test(b.textContent)));
  const laneCount = doc.querySelectorAll('#prioView .pr-lanecell').length;
  ok(laneCount > 1, 'grouping by epic renders swimlane label cells (' + laneCount + ')');
  ok(doc.querySelectorAll('#prioView .sp-col').length === laneCount * state().phases.length,
    'every swimlane row spans all phase columns');
  ok(JSON.parse(window.localStorage.getItem('headway-ui-v1')).prioGroup === 'epic',
    'the grouping choice persists with the view prefs');
  // grouped by epic, cards drop the redundant epic chip
  ok(!doc.querySelector('#prioView .pr-card [data-pract="epic"]'),
    'epic swimlanes hide the duplicate epic chip on cards');
  click(doc.querySelector('#prioView [data-prdd="group"]'));
  click([...doc.querySelectorAll('#popover .menu-list button')].find(b => /None/.test(b.textContent)));
  click(doc.querySelector('#viewTabs [data-view="planning"]'));
}

// ---------------------------------------------------------------- sprinting view
{
  click(doc.querySelector('#viewTabs [data-view="sprints"]'));
  ok(doc.body.dataset.view === 'sprints', 'view switches to Sprinting');
  const side = doc.querySelectorAll('#sprintView .spv-sbtn');
  const secs = doc.querySelectorAll('#sprintView .spv-sec');
  ok(side.length > 2 && side.length === secs.length,
    'sidebar lists Unscheduled plus every sprint, one page section each (' + side.length + ')');
  ok(side[side.length - 1].dataset.spside === 'u' && /Unscheduled/.test(side[side.length - 1].textContent) &&
    secs[secs.length - 1].dataset.spsec === 'u' && side[0].dataset.spside !== 'u',
    'Unscheduled comes last in the sidebar and on the page');
  const unsched = state().items.filter(i => i.startDay == null).length;
  ok(doc.querySelectorAll('#sprintView .spv-sec[data-spsec="u"] .spv-row').length === unsched,
    'the Unscheduled section lists every unscheduled item');
  const spf = doc.querySelector('#spFilter');
  ok(!!spf && !!spf.closest('.filter-wrap').querySelector('.filter-ico'), 'the row filter carries the search icon');
  // phase / epic / workstream dropdowns narrow the rows, like the prioritizing board
  {
    const rowsOf = () => [...doc.querySelectorAll('#sprintView .spv-row')];
    const itemOf = r => state().items.find(i => i.id === r.dataset.spid);
    const pick = re => click([...doc.querySelectorAll('#popover .menu-list button')].find(b => re.test(b.textContent)));
    const before = rowsOf().length;
    const spBar = doc.querySelector('#sprintView .pr-bar');
    ok(!!spBar.querySelector('[data-spdd="fphase"]') && !!spBar.querySelector('[data-spdd="fepic"]') && !!spBar.querySelector('[data-spdd="fws"]'),
      'the sprinting toolbar offers phase, epic and workstream filters');
    const ph = state().phases[1];
    click(spBar.querySelector('[data-spdd="fphase"]'));
    pick(new RegExp('^' + ph.name.replace(/[.*+?^${}()|[\]\\]/g, '\\$&') + '$'));
    ok(rowsOf().length > 0 && rowsOf().length < before && rowsOf().every(r => itemOf(r).phaseId === ph.id),
      'the phase filter narrows the sprint rows to that phase');
    const fphBtn = doc.querySelector('#sprintView [data-spdd="fphase"]');
    ok(fphBtn.classList.contains('pr-on') && fphBtn.textContent.indexOf(ph.name) !== -1,
      'an active phase filter lights up and names the phase');
    // the sidebar mirrors the section header (a plain row count, or a story-point total)
    const uCt = doc.querySelector('#sprintView .spv-sbtn[data-spside="u"] .pr-lanect').textContent.trim();
    const uHd = doc.querySelector('#sprintView .spv-sec[data-spsec="u"] .spv-sechd .pr-lanect').textContent.trim();
    ok(uCt === uHd && (/^\d+$/.test(uCt)
      ? Number(uCt) === doc.querySelectorAll('#sprintView .spv-sec[data-spsec="u"] .spv-row').length
      : / pt$/.test(uCt)),
      'sidebar counts follow the filtered rows');
    click(fphBtn); pick(/All phases/);
    ok(rowsOf().length === before && !doc.querySelector('#sprintView [data-spdd="fphase"]').classList.contains('pr-on'),
      'clearing the phase filter restores every row and dims the button');
    click(doc.querySelector('#sprintView [data-spdd="fepic"]'));
    const epBtn = [...doc.querySelectorAll('#popover .menu-list button')].filter(b => !/All epics|no epic/.test(b.textContent))[0];
    const epName = epBtn.textContent.trim();
    click(epBtn);
    ok(rowsOf().length > 0 && rowsOf().every(r => itemOf(r).epic === epName), 'the epic filter narrows the sprint rows to that epic');
    ok(doc.querySelector('#sprintView [data-spdd="fepic"]').classList.contains('pr-on'), 'an active epic filter lights up');
    click(doc.querySelector('#sprintView [data-spdd="fepic"]')); pick(/All epics/);
    click(doc.querySelector('#sprintView [data-spdd="fws"]'));
    const wsBtn = [...doc.querySelectorAll('#popover .menu-list button')].filter(b => !/All workstreams|default/.test(b.textContent))[0];
    const wsName = wsBtn.textContent.trim();
    click(wsBtn);
    ok(rowsOf().length > 0 && rowsOf().every(r => itemOf(r).workstream === wsName), 'the workstream filter narrows the sprint rows to that stream');
    ok(!!doc.querySelector('#sprintView [data-spdd="fws"].pr-on .dd-label .dd-dot'), 'an active workstream filter shows its color dot');
    click(doc.querySelector('#sprintView [data-spdd="fws"]')); pick(/All workstreams/);
    ok(rowsOf().length === before, 'clearing every dropdown restores the full list');
    ok(!doc.querySelector('#prioView [data-prdd="fepic"].pr-on'), 'sprinting filters do not leak into the prioritizing board');
  }
  const sched = state().items.find(i => i.startDay != null);
  const schedRow = doc.querySelector('#sprintView .spv-row[data-spid="' + sched.id + '"]');
  ok(!!schedRow && schedRow.dataset.spsec !== 'u', 'scheduled items list under a sprint');
  ok(!!schedRow.querySelector('[data-spact="epic"]') && !!schedRow.querySelector('[data-spact="size"]') &&
    !!schedRow.querySelector('[data-spf="feature"]'), 'rows carry an inline title, epic and size chips');
  ok([...doc.querySelectorAll('#sprintView .spv-sec')].every(sec =>
    !!sec.querySelector('.spv-add') === !sec.querySelector('.spv-rows .spv-row')),
    'only an empty section ends with an Add feature button');
  // drag an unscheduled row onto a sprint in the sidebar: it gets that
  // sprint's start day and a span, and lists under that sprint
  const uRow = doc.querySelector('#sprintView .spv-sec[data-spsec="u"] .spv-row');
  const uId = uRow.dataset.spid;
  const target = doc.querySelectorAll('#sprintView .spv-sbtn')[2]; // re-queried: the filters above re-rendered
  const tnum = Number(target.dataset.spside);
  uRow.dispatchEvent(new window.MouseEvent('pointerdown', { bubbles: true, button: 0, clientX: 10, clientY: 10 }));
  target.dispatchEvent(new window.MouseEvent('pointermove', { bubbles: true, clientX: 60, clientY: 60 }));
  window.dispatchEvent(new window.MouseEvent('pointerup', { bubbles: true, clientX: 60, clientY: 60 }));
  const moved = state().items.find(i => i.id === uId);
  ok(moved.startDay != null && moved.durDays != null, 'dropping on a sidebar sprint schedules the item');
  ok(doc.querySelector('#sprintView .spv-row[data-spid="' + uId + '"]').dataset.spsec === String(tnum),
    'and it now lists under that sprint');
  // clicking a side entry jumps to the sprint heading (jsdom has no layout,
  // so this only proves the jump code runs; the offset math is visual)
  {
    let threw = false;
    try { click(doc.querySelector('#sprintView .spv-sbtn[data-spside="' + tnum + '"]')); } catch (e) { threw = true; }
    ok(!threw, 'clicking a side-list sprint runs the jump without error');
  }
  // drop the row before another row of the same section: document order
  // changes (the shared items array), which every view reads
  const secRows = doc.querySelectorAll('#sprintView .spv-sec[data-spsec="' + tnum + '"] .spv-row');
  ok(secRows.length >= 2, 'the target sprint has rows to reorder among');
  const first = secRows[0], mine = doc.querySelector('#sprintView .spv-row[data-spid="' + uId + '"]');
  mine.dispatchEvent(new window.MouseEvent('pointerdown', { bubbles: true, button: 0, clientX: 10, clientY: 10 }));
  first.dispatchEvent(new window.MouseEvent('pointermove', { bubbles: true, clientX: 60, clientY: -5 }));
  window.dispatchEvent(new window.MouseEvent('pointerup', { bubbles: true, clientX: 60, clientY: -5 }));
  const ids = state().items.map(i => i.id);
  ok(ids.indexOf(uId) === ids.indexOf(first.dataset.spid) - 1, 'dropping before a row reorders the items array');
  ok(state().items.find(i => i.id === uId).startDay === moved.startDay, 'a same-sprint reorder leaves the start alone');
  // dragging to another sprint's section (empty space) moves the start
  const other = [...doc.querySelectorAll('#sprintView .spv-sec')].find(sc => sc.dataset.spsec !== 'u' && sc.dataset.spsec !== String(tnum));
  const onum = Number(other.dataset.spsec);
  const mine2 = doc.querySelector('#sprintView .spv-row[data-spid="' + uId + '"]');
  mine2.dispatchEvent(new window.MouseEvent('pointerdown', { bubbles: true, button: 0, clientX: 10, clientY: 10 }));
  other.querySelector('.spv-sechd').dispatchEvent(new window.MouseEvent('pointermove', { bubbles: true, clientX: 60, clientY: 60 }));
  window.dispatchEvent(new window.MouseEvent('pointerup', { bubbles: true, clientX: 60, clientY: 60 }));
  ok(doc.querySelector('#sprintView .spv-row[data-spid="' + uId + '"]').dataset.spsec === String(onum) &&
    state().items.find(i => i.id === uId).durDays === moved.durDays,
    'dropping into another section moves the start and keeps the span');
  click(doc.querySelector('#sprintView')); // the click a browser fires after the drag's pointerup
  // stories level: the Feature / Story dropdown is the first control in the bar
  // (same place and look as Planning / Scoping / Prioritizing)
  const spBar = doc.querySelector('#sprintView .pr-bar');
  ok(spBar.firstElementChild.matches('.dm-btn[data-spdd="level"]') && /Feature/.test(spBar.firstElementChild.textContent),
    'Sprinting: level dropdown sits at the left of the toolbar, reading Feature');
  click(doc.querySelector('#sprintView [data-spdd="level"]'));
  click([...doc.querySelectorAll('#popover .menu-list button')].find(b => /Story/.test(b.textContent)));
  ok(/Story/.test(doc.querySelector('#sprintView [data-spdd="level"]').textContent), 'the Stories level toggles');
  ok(JSON.parse(window.localStorage.getItem('headway-ui-v1')).sprLevel === 'story', 'the level persists with the view prefs');
  ok(!doc.querySelector('#sprintView .spv-row[data-spid]:not([data-spst])'), 'story level lists no feature rows');
  click(doc.querySelector('#sprintView [data-spdd="level"]'));
  click([...doc.querySelectorAll('#popover .menu-list button')].find(b => /Feature/.test(b.textContent)));
  // undo the three moves so later blocks see the fixture's schedule
  for (let k = 0; k < 3; k++) window.dispatchEvent(new window.KeyboardEvent('keydown', { key: 'z', metaKey: true, bubbles: true }));
  ok(state().items.find(i => i.id === uId).startDay == null, 'undo restores the item to Unscheduled');
  click(doc.querySelector('#viewTabs [data-view="planning"]'));
}
// ------------------------------------------------------- RICE priority scheme
{
  suTab('est');
  ok(schemeOpts('prio', 'feature').length === 4, 'four priority schemes (none, MoSCoW, levels, RICE)');
  ok(schemeOpts('prio', 'story').length === 3, 'stories offer three (no RICE)');
  pickScheme('prio', 'feature', 'rice');
  ok(state().meta.priorityScheme === 'rice', 'RICE scheme commits');
  // prio cards now carry the priority chip; clicking it opens RICE dropdowns
  click(doc.querySelector('#viewTabs [data-view="prio"]'));
  // pick a card whose phase is expanded so its scoping row exists later
  const openPhases = state().phases.filter(ph => !ph.collapsed).map(ph => ph.id);
  const chip = [...doc.querySelectorAll('#prioView .pr-card [data-pract="priority"]')].find(c => {
    const iid = c.closest('[data-prcard]').dataset.prcard;
    const item = state().items.find(i => i.id === iid);
    return item && openPhases.indexOf(item.phaseId) !== -1;
  });
  ok(!!chip, 'cards carry the priority chip under RICE');
  const cid = chip.closest('[data-prcard]').dataset.prcard;
  click(chip);
  const sels = doc.querySelectorAll('#popover .rice-pop select[data-rk]');
  ok(sels.length === 4, 'the priority chip opens four RICE dropdowns');
  const pick = (k, v) => {
    const sel = doc.querySelector('#popover select[data-rk="' + k + '"]');
    sel.value = String(v);
    sel.dispatchEvent(new window.Event('change', { bubbles: true }));
  };
  pick('reach', 100); pick('impact', 2); pick('confidence', 80); pick('effort', 4);
  const scored = state().items.find(i => i.id === cid);
  ok(window.RM.riceScore(scored) === 40, 'dropdown picks score R × I × C% ÷ E');
  window.dispatchEvent(new window.KeyboardEvent('keydown', { key: 'Escape', bubbles: true }));
  const chip2 = doc.querySelector('#prioView [data-prcard="' + cid + '"] [data-pract="priority"]');
  ok(/40/.test(chip2.textContent), 'the chip shows the computed score');
  ok(doc.querySelector('#prioView [data-prcol="' + scored.phaseId + '"] .pr-card').dataset.prcard === cid,
    'the scored card auto-sorts to the top of its column');
  // the scoping Priority column shows the identical chip
  click(doc.querySelector('#viewTabs [data-view="scoping"]'));
  const scChip = doc.querySelector('#rows .row.item[data-id="' + cid + '"] [data-act="priority"]');
  ok(!!scChip && /40/.test(scChip.textContent), 'the scoping Priority column shows the same score');
  {
    const stRowR = doc.querySelector('#rows .row.story[data-story]');
    ok(!stRowR || !!stRowR.querySelector('[data-act="st-pri"]') === window.RM.priorityEnabled(state(), 'story'),
      'story priority chips follow the STORY scheme, untouched by feature RICE');
  }
  // undo the scheme switch + four RICE picks so later suites see the default state
  for (let u = 0; u < 5; u++) window.dispatchEvent(new window.KeyboardEvent('keydown', { key: 'z', metaKey: true, bubbles: true }));
  click(doc.querySelector('#viewTabs [data-view="planning"]'));
}

// (workflow statuses were removed — features and stories carry only Done)
{
  click(doc.querySelector('#viewTabs [data-view="planning"]'));
}

// ---------------------------------------------------------------- batch 6: priority, story cells, columns tab
{
  // priority column: enable MoSCoW in Setup, chip appears in Scoping
  suTab('est');
  ok(schemeOpts('prio', 'feature').length === 4, 'four priority schemes (none, MoSCoW, levels, RICE)');
  pickScheme('prio', 'feature', 'moscow');
  ok(state().meta.priorityScheme === 'moscow', 'MoSCoW priority commits');
  click(doc.querySelector('#viewTabs [data-view="scoping"]'));
  ok(!!doc.querySelector('#hdrSprints [data-col="priority"]'), 'Priority column renders when enabled');
  const priChip = doc.querySelector('#rows .row.item [data-act="priority"]');
  click(priChip);
  const mustOpt = Array.from(doc.querySelectorAll('#popover .menu-list button')).find(b => /Must/.test(b.textContent));
  ok(!!mustOpt, 'priority dropdown offers MoSCoW values');
  click(mustOpt);
  const priRow = priChip.closest('.row');
  ok(state().items.find(i => i.id === priRow.dataset.id).priority === 'M', 'priority commits to item.priority');
  window.dispatchEvent(new window.KeyboardEvent('keydown', { key: 'z', metaKey: true, bubbles: true }));
  window.dispatchEvent(new window.KeyboardEvent('keydown', { key: 'z', metaKey: true, bubbles: true }));

  // assignees is a fixed scoping column before Size
  const ordCols = Array.from(doc.querySelectorAll('#hdrSprints .sc-hcell[data-col]')).map(c => c.dataset.col);
  ok(ordCols.indexOf('assignees') !== -1 && ordCols.indexOf('assignees') < ordCols.indexOf('size'),
    'Assignees column sits before Size');

  // stories: text cells + size editable, other fixed cells n/a
  // (expand a feature with stories first — collapsed features hide them)
  {
    const withStories = state().items.find(i => i.stories.length);
    click(doc.querySelector('#viewTabs [data-view="planning"]'));
    const chev = doc.querySelector('#rows .row.item[data-id="' + withStories.id + '"] [data-act="stories"]');
    if (!doc.querySelector('#rows .row.story[data-story]')) click(chev);
    click(doc.querySelector('#viewTabs [data-view="scoping"]'));
  }
  const stRow = doc.querySelector('#rows .row.story[data-story]');
  ok(!!stRow && !!stRow.querySelector('[data-stscope="description"]'), 'story rows carry editable text cells');
  ok(!!stRow.querySelector('[data-act="st-size"]'), 'story rows carry an editable Size chip');
  ok(stRow.querySelectorAll('.sc-cell.sc-na').length > 0, 'non-applicable story cells gray out');
  const stCell = stRow.querySelector('[data-stscope="description"]');
  stCell.innerHTML = 'story desc from grid';
  stCell.dispatchEvent(new window.FocusEvent('focusout', { bubbles: true }));
  const stParent = state().items.find(i => i.id === stRow.dataset.id);
  ok(stParent.stories.find(s2 => s2.id === stRow.dataset.story).description === 'story desc from grid',
    'story cell edit commits to the story');
  window.dispatchEvent(new window.KeyboardEvent('keydown', { key: 'z', metaKey: true, bubbles: true }));
}

// columns tab in Setup: full ordered list, add/remove text columns
{
  suTab('columns');
  ok(doc.querySelectorAll('#setupView [data-sulist="scol"] .su-row').length ===
    doc.querySelectorAll('#hdrSprints .sc-hcell[data-col]').length ||
    doc.querySelectorAll('#setupView [data-sulist="scol"] .su-row').length > 5,
    'Columns tab lists the scoping columns');
  const addInp2 = doc.querySelector('#suColAdd');
  addInp2.value = 'Reviewer';
  click(doc.querySelector('#suColAddBtn'));
  ok(state().meta.scopeCols.some(c => c.label === 'Reviewer'), 'column added from Setup');
  const revKey = state().meta.scopeCols.find(c => c.label === 'Reviewer').key;
  click(doc.querySelector('#setupView [data-sucolrm="' + revKey + '"]'));
  ok(!state().meta.scopeCols.some(c => c.label === 'Reviewer'), 'column removed from Setup');
  click(doc.querySelector('#viewTabs [data-view="planning"]'));
}

// ---------------------------------------------------------------- batch 7: deadlines, calendar, stories everywhere
{
  // Deadline is a fixed scoping column right after Start
  click(doc.querySelector('#viewTabs [data-view="scoping"]'));
  const ord7 = Array.from(doc.querySelectorAll('#hdrSprints .sc-hcell[data-col]')).map(c => c.dataset.col);
  ok(ord7.indexOf('deadline') === ord7.indexOf('start') + 1, 'Deadline column follows Start');

  // the deadline chip opens the shared calendar; picking a day commits
  const dlChip = doc.querySelector('#rows .row.item [data-act="deadline"]');
  const dlRowId = dlChip.closest('.row').dataset.id;
  click(dlChip);
  const calPop = doc.querySelector('#calPop');
  ok(!!calPop && !calPop.hidden && calPop.querySelectorAll('.cal-day').length === 42,
    'deadline chip opens the calendar popover');
  const dayBtn = calPop.querySelector('.cal-day:not(.out)');
  const pickedIso = dayBtn.dataset.iso;
  click(dayBtn);
  ok(state().items.find(i => i.id === dlRowId).deadline === pickedIso, 'calendar pick commits the deadline');

  // planning paints a deadline tick for the item
  click(doc.querySelector('#viewTabs [data-view="planning"]'));
  ok(doc.querySelectorAll('#rows .r-dl-mark').length > 0, 'deadline tick renders on the timeline');
  window.dispatchEvent(new window.KeyboardEvent('keydown', { key: 'z', metaKey: true, bubbles: true }));

  // the Start chip opens the calendar too (no more native date input)
  click(doc.querySelector('#viewTabs [data-view="scoping"]'));
  click(doc.querySelector('#rows .row.item [data-act="startd"]'));
  ok(!doc.querySelector('#calPop').hidden && !!doc.querySelector('#calPop [data-cal="clear"]'),
    'Start chip opens the calendar with an Unschedule action');
  doc.dispatchEvent(new window.Event('pointerdown', { bubbles: true }));
  // the panel's date fields became calendar inputs (select a row first)
  click(doc.querySelector('#rows .row.item .r-num'));
  ok(!!doc.querySelector('#panel input.cal-in[data-f="deadline"]'), 'panel carries a Deadline calendar field');

  // story rows: assignees + duration chips and italic rolled-up values
  const stRow7 = doc.querySelector('#rows .row.story[data-story]');
  ok(!!stRow7.querySelector('[data-act="st-asg"]'), 'story rows carry an Assignees chip');
  ok(!!stRow7.querySelector('[data-act="st-wk"]'), 'story rows carry a Duration chip');
  ok(stRow7.querySelectorAll('.sc-roll').length > 0, 'rolled-up feature values show dimmed in story rows');
  // duration commits in story units (weeks × slots)
  click(stRow7.querySelector('[data-act="st-wk"]'));
  const stwInp = stRow7.querySelector('[data-act="st-wk"] input') || doc.querySelector('#rows .hc-edit');
  ok(!!stwInp, 'story duration chip opens an inline editor');
  stwInp.value = '2';
  stwInp.dispatchEvent(new window.FocusEvent('blur'));
  const stRowP = state().items.find(i => i.id === stRow7.dataset.id);
  ok(stRowP.stories.find(s2 => s2.id === stRow7.dataset.story).durDays === 10, 'story duration commits (2w = 10 days)');
  window.dispatchEvent(new window.KeyboardEvent('keydown', { key: 'z', metaKey: true, bubbles: true }));
  // assignees dropdown assigns a person to the story
  const stAsgChip = doc.querySelector('#rows .row.story[data-story] [data-act="st-asg"]');
  click(stAsgChip);
  const asgOpt = doc.querySelector('#popover .menu-list button');
  ok(!!asgOpt, 'story assignee dropdown lists the roster');
  click(asgOpt);
  const stRowP2 = state().items.find(i => i.id === stRow7.dataset.id);
  ok(stRowP2.stories.find(s2 => s2.id === stRow7.dataset.story).assignees.length === 1, 'story assignee commits');
  window.dispatchEvent(new window.KeyboardEvent('keydown', { key: 'z', metaKey: true, bubbles: true }));

  // budgeting: names edit in place, people join several workstreams
  click(doc.querySelector('#viewTabs [data-view="budget"]'));
  const nmInp = doc.querySelector('#rows .row.brole input[data-bud="name"]');
  ok(!!nmInp, 'budget rows carry an editable name input');
  const oldName = nmInp.value;
  nmInp.value = 'Renamed Person';
  nmInp.dispatchEvent(new window.Event('change', { bubbles: true }));
  ok(state().team[0].name === 'Renamed Person', 'budget name edit renames the person');
  window.dispatchEvent(new window.KeyboardEvent('keydown', { key: 'z', metaKey: true, bubbles: true }));
  ok(state().team[0].name === oldName, 'rename undoes cleanly');
  // toggle two workstreams onto the first person
  const wsChip7 = doc.querySelector('#rows .row.brole [data-bact="ws"]');
  click(wsChip7);
  const wsOpts = Array.from(doc.querySelectorAll('#popover .menu-list button')).slice(1); // skip "none"
  ok(wsOpts.length >= 2, 'workstream menu lists the project workstreams');
  click(wsOpts[0]);
  click(doc.querySelector('#rows .row.brole [data-bact="ws"]'));
  click(Array.from(doc.querySelectorAll('#popover .menu-list button')).slice(1)[1]);
  ok((state().team[0].workstreams || []).length === 2, 'a person can join two workstreams');
  ok(/\+1/.test(doc.querySelector('#rows .row.brole [data-bact="ws"]').textContent), 'chip shows the primary +N');
  window.dispatchEvent(new window.KeyboardEvent('keydown', { key: 'z', metaKey: true, bubbles: true }));
  window.dispatchEvent(new window.KeyboardEvent('keydown', { key: 'z', metaKey: true, bubbles: true }));
  click(doc.querySelector('#viewTabs [data-view="planning"]'));
}

// ---------------------------------------------------------------- batch 8: story rows fill out
{
  click(doc.querySelector('#viewTabs [data-view="scoping"]'));
  const stRow8 = doc.querySelector('#rows .row.story[data-story]');
  // start and deadline are the story's own, always editable
  ok(!!stRow8.querySelector('[data-act="st-startd"]'), 'story Start chip is always editable (no roll-up)');
  ok(!!stRow8.querySelector('[data-act="st-dl"]'), 'story Deadline chip is its own field');
  ok(!stRow8.querySelector('.sc-cell[data-col="start"].sc-na') && !stRow8.querySelector('.sc-cell[data-col="deadline"].sc-na'),
    'start/deadline story cells are not grayed out');
  // the left-pane title edits in place; the delete button is gone
  const stName = stRow8.querySelector('.st-name');
  ok(!!stName && stName.getAttribute('contenteditable') === 'true', 'story title is editable in the left pane');
  ok(!stRow8.querySelector('.row-left .st-del'), 'no delete button in the story left pane');
  stName.textContent = 'Renamed story inline';
  stName.dispatchEvent(new window.FocusEvent('focusout', { bubbles: true }));
  const stP8 = state().items.find(i => i.id === stRow8.dataset.id);
  ok(stP8.stories.find(s2 => s2.id === stRow8.dataset.story).title === 'Renamed story inline',
    'story title edit commits');
  window.dispatchEvent(new window.KeyboardEvent('keydown', { key: 'z', metaKey: true, bubbles: true }));
  // story deadline commits through the calendar
  click(doc.querySelector('#rows .row.story[data-story] [data-act="st-dl"]'));
  const calDay8 = doc.querySelector('#calPop .cal-day:not(.out)');
  const iso8 = calDay8.dataset.iso;
  click(calDay8);
  const stP8b = state().items.find(i => i.id === stRow8.dataset.id);
  ok(stP8b.stories.find(s2 => s2.id === stRow8.dataset.story).deadline === iso8, 'story deadline commits');
  window.dispatchEvent(new window.KeyboardEvent('keydown', { key: 'z', metaKey: true, bubbles: true }));
  click(doc.querySelector('#viewTabs [data-view="planning"]'));
}

// ---------------------------------------------------------------- batch 9: version history, budgeting columns, resources parity
{
  // the filter cell must keep .hdr-left's sticky (an own `position` broke it)
  const css = fs.readFileSync(path.join(ROOT, 'css/app.css'), 'utf8');
  const cellRule = /#filterCell\s*{[^}]*}/.exec(css);
  ok(cellRule && !/position\s*:/.test(cellRule[0]), 'filter cell has no position of its own (stays sticky in the left pane)');

  // floating B/I toolbar closes when the view changes
  click(doc.querySelector('#viewTabs [data-view="scoping"]'));
  const rich = doc.querySelector('#rows .sc-rich');
  if (rich) {
    rich.dispatchEvent(new window.FocusEvent('focusin', { bubbles: true }));
    ok(doc.querySelector('#scFmtBar') && !doc.querySelector('#scFmtBar').hidden, 'format toolbar opens on rich-cell focus');
    click(doc.querySelector('#viewTabs [data-view="planning"]'));
    ok(doc.querySelector('#scFmtBar').hidden, 'format toolbar closes when the view changes');
  } else {
    ok(false, 'scoping view has a rich cell to test the format toolbar with');
  }

  // header: Setup / Version history split button
  ok(!!doc.querySelector('#setupSplit #btnSetup') && !!doc.querySelector('#setupSplit #btnHistory'),
    'setup split button carries Setup and Version history halves');

  // budgeting: Name · Role · Rate card · Workstream columns + reorder grips
  click(doc.querySelector('#viewTabs [data-view="budget"]'));
  const brow = doc.querySelector('#rows .row.brole[data-mid]');
  ok(!!brow.querySelector('input[data-bud="name"]') && !!brow.querySelector('input[data-bud="role"]'),
    'budget rows have Name and free-text Role inputs');
  ok(!!brow.querySelector('[data-bact="type"]') && !!brow.querySelector('[data-bact="ws"]'),
    'budget rows have Rate card and Workstream chips');
  ok(!!brow.querySelector('.bu-grip'), 'budget person rows have a reorder grip');
  // costs reorder too
  click(doc.querySelector('#rows .row.addrow[data-kind="baddcost"]'));
  const crow = doc.querySelector('#rows .row.bcost[data-cost]');
  ok(!!crow && !!crow.querySelector('.bu-grip'), 'budget cost rows have a reorder grip');

  // free-text role commits; empty name is allowed (re-query — the add-cost
  // commit above re-rendered the rows)
  const brow2 = doc.querySelector('#rows .row.brole[data-mid]');
  const roleInp = brow2.querySelector('input[data-bud="role"]');
  roleInp.value = 'Senior Backend Dev';
  roleInp.dispatchEvent(new window.Event('change', { bubbles: true }));
  ok(state().team.find(m => m.id === brow2.dataset.mid).role === 'Senior Backend Dev', 'free-text role commits');
  const nameInp = doc.querySelector('#rows .row.brole input[data-bud="name"]');
  nameInp.value = '';
  nameInp.dispatchEvent(new window.Event('change', { bubbles: true }));
  ok(state().team[0].name === '', 'name can be cleared (optional)');
  // rate-card dropdown offers a "none" option that clears the assignment
  const typeChip = doc.querySelector('#rows .row.brole [data-bact="type"]');
  click(typeChip);
  const noneOpt = Array.from(doc.querySelectorAll('#popover button')).find(b => /none/.test(b.textContent));
  ok(!!noneOpt, 'rate-card dropdown offers — none —');
  if (noneOpt) { click(noneOpt); ok(state().team[0].type === '', 'rate card can be unassigned'); }
  doc.querySelector('#popover').hidden = true;

  // resources panel mirrors the budgeting columns (minus the money ones)
  click(doc.querySelector('#viewTabs [data-view="planning"]'));
  const rrow2 = doc.querySelector('#resGrid .rrow[data-mid]');
  ok(!!rrow2.querySelector('input[data-bud="name"]') && !!rrow2.querySelector('input[data-bud="role"]') &&
    !!rrow2.querySelector('[data-bact="type"]') && !!rrow2.querySelector('[data-bact="ws"]'),
    'resources rows carry the same name/role/rate-card/workstream columns');
  ok(!rrow2.querySelector('[data-bud="cost"]') && !rrow2.querySelector('[data-bud="rate"]'),
    'resources rows carry no money columns');
  const rRole = rrow2.querySelector('input[data-bud="role"]');
  rRole.value = 'QA Lead';
  rRole.dispatchEvent(new window.Event('change', { bubbles: true }));
  ok(state().team[0].role === 'QA Lead', 'role edits commit from the resources panel too');

  // version history: commits record who + what, coalescing rapid same-kind edits
  window.localStorage.setItem('headway-user-v1', 'Test User');
  // re-query between edits: each commit re-renders the panel
  let rr = doc.querySelector('#resGrid .rrow[data-mid] input[data-bud="role"]');
  rr.value = 'QA Lead II';
  rr.dispatchEvent(new window.Event('change', { bubbles: true }));
  rr = doc.querySelector('#resGrid .rrow[data-mid] input[data-bud="role"]');
  rr.value = 'QA Lead III';
  rr.dispatchEvent(new window.Event('change', { bubbles: true }));
  const hist = state().history || [];
  const lastH = hist[hist.length - 1];
  ok(lastH && lastH.label === 'person role' && lastH.u === 'Test User' && lastH.n >= 2,
    'history records user + change and coalesces rapid edits (' + JSON.stringify(lastH) + ')');
  // Version History is a full page now — the split button's right half opens it
  click(doc.querySelector('#btnHistory'));
  ok(doc.body.dataset.view === 'history', 'the History button switches to the Version History page');
  const hv = doc.querySelector('#historyView');
  ok(/Version history/.test(hv.textContent) && /Test User/.test(hv.textContent) && /Person role/.test(hv.textContent),
    'version history page shows who and what');
  const firstItem = hv.querySelector('.vh-item');
  ok(firstItem && /Test User/.test(firstItem.textContent), 'newest entry listed first');
  click(doc.querySelector('#viewTabs [data-view="planning"]'));
}

// ---------------------------------------------------------------- detail level dropdown
{
  click(doc.querySelector('#viewTabs [data-view="planning"]'));
  const dmBtn = doc.querySelector('#detailBtn');
  ok(!!dmBtn && !!dmBtn.querySelector('svg,[data-lucide]'), 'a detail-level dropdown sits left of the Feature header');
  click(dmBtn);
  const dmMenu = doc.querySelector('#popover .menu-list');
  ok(!!dmMenu && dmMenu.querySelectorAll('[data-mi]').length === 2 &&
    !/Phase/.test(dmMenu.textContent) && /Feature/.test(dmMenu.textContent) && /Story/.test(dmMenu.textContent),
    'the dropdown offers Feature / Story only, each with an icon');
  ok(dmMenu.querySelectorAll('[data-mi] svg, [data-mi] [data-lucide]').length >= 2, 'detail options carry icons');
  ok(!dmMenu.querySelector('[data-mi] small'), 'detail options carry no explanatory text');
  click(dmMenu.querySelector('[data-mi="0"]')); // Feature
  ok(doc.querySelectorAll('#rows .row.story').length === 0,
    'feature detail keeps story rows tucked away');
  click(dmBtn);
  click(doc.querySelector('#popover .menu-list [data-mi="1"]')); // Story
  ok(doc.querySelectorAll('#rows .row.story').length > 0, 'story detail opens every story row');
  click(doc.querySelector('#viewTabs [data-view="scoping"]'));
  click(doc.querySelector('#detailBtn'));
  click(doc.querySelector('#popover .menu-list [data-mi="1"]')); // Story
  ok(doc.querySelectorAll('#rows .row.story').length > 0, 'story detail applies on Scoping too');
  click(doc.querySelector('#detailBtn'));
  click(doc.querySelector('#popover .menu-list [data-mi="0"]')); // back to Feature
  click(doc.querySelector('#viewTabs [data-view="planning"]'));
}

// ---------------------------------------------------------------- Jira login stays on this machine
{
  window.HeadwayJira.saveCreds({ email: 'me@example.com', token: 'SECRET-TOKEN-123' });
  window.eval("HeadwayApp.ai.commit('jira settings', function (s) { s.meta.jira = { site: 'https://x.atlassian.net', project: 'HW' }; })");
  const docJson = JSON.stringify(state());
  const uiJson = window.localStorage.getItem('headway-ui-v1') || '';
  ok(docJson.indexOf('SECRET-TOKEN-123') === -1 && docJson.indexOf('me@example.com') === -1,
    'the Jira token and email never enter the document');
  ok(uiJson.indexOf('SECRET-TOKEN-123') === -1, 'the Jira token is not in the UI snapshot that the file carries');
  ok(docJson.indexOf('x.atlassian.net') !== -1, 'the site and project are shared through the document');
  ok(window.HeadwayJira.credsFor(state()).site === 'https://x.atlassian.net', 'the login pairs with the document site');
  window.localStorage.removeItem('headway-jira-v1');
}

// ---------------------------------------------------------------- workstream via context menu
{
  const itRow = doc.querySelector('#rows .row.item');
  itRow.dispatchEvent(new window.MouseEvent('contextmenu', { bubbles: true, cancelable: true, clientX: 60, clientY: 60 }));
  const ctx = doc.querySelector('#popover .menu-list');
  ok(!!ctx && /Set workstream…/.test(ctx.textContent), 'item context menu offers Set workstream…');
  doc.querySelector('#popover').hidden = true;
  doc.querySelector('#popover').innerHTML = '';
}

// ---------------------------------------------------------------- version history diffs
{
  window.localStorage.setItem('headway-user-v1', 'Test User');
  const nameSpan = doc.querySelector('#rows .row.item [data-rowname]');
  const rowId = nameSpan.closest('.row').dataset.id;
  const it = window.RM.itemById(state(), rowId);
  const oldTitle = it.feature;
  nameSpan.dispatchEvent(new window.MouseEvent('dblclick', { bubbles: true }));
  const inp = doc.querySelector('#rows .row.item[data-id="' + rowId + '"] input[data-rowname]');
  inp.value = 'Diffed Feature Title';
  inp.dispatchEvent(new window.KeyboardEvent('keydown', { key: 'Enter', bubbles: true }));
  const en = state().history[state().history.length - 1];
  ok(Array.isArray(en.d) && en.d.some(op => op[0] === 'scope' && /Title/.test(op[1]) &&
    op[2] === oldTitle && op[3] === 'Diffed Feature Title'),
    'a commit records field-level old → new detail (' + JSON.stringify(en.d && en.d[0]) + ')');

  click(doc.querySelector('#btnHistory'));
  const hv = doc.querySelector('#historyView');
  ok(!!hv.querySelector('.hv-side') && !!hv.querySelector('.hv-detail'),
    'Version History renders as a page: change feed left, diff right');
  ok(!!hv.querySelector('.hv-tab'), 'the diff groups changes into tabs');
  const oldEl = Array.from(hv.querySelectorAll('.vd-old')).find(el => el.textContent === oldTitle);
  const newEl = Array.from(hv.querySelectorAll('.vd-new')).find(el => el.textContent === 'Diffed Feature Title');
  ok(!!oldEl && !!newEl, 'the newest change shows its old (red) and new (green) values');

  // schedule a change too, then check the Timeline tab appears for it
  click(doc.querySelector('#viewTabs [data-view="planning"]'));
  const bar2 = doc.querySelector('#rows .bar:not(.ms)[data-bar]');
  click(doc.querySelector('#rows .row.item[data-id="' + bar2.getAttribute('data-bar') + '"] .r-num'));
  const durInp = doc.querySelector('#panel input[data-f="durWeeks"]');
  durInp.value = String(parseFloat(durInp.value) + 1);
  durInp.dispatchEvent(new window.Event('change', { bubbles: true }));
  click(doc.querySelector('#btnHistory'));
  const tlTab = Array.from(doc.querySelectorAll('#historyView .hv-tab')).find(t => /Timeline/.test(t.textContent));
  ok(!!tlTab, 'timeline edits appear under a Timeline tab');

  // tick two entries → the page diffs ACROSS them
  const cks = doc.querySelectorAll('#historyView [data-vhck]');
  ok(cks.length >= 2, 'entries offer compare checkboxes');
  click(cks[0]);
  click(doc.querySelectorAll('#historyView [data-vhck]')[1]);
  ok(/Comparing/.test(doc.querySelector('#historyView .hv-dhead h2').textContent),
    'ticking two versions diffs between them');
  click(doc.querySelector('#viewTabs [data-view="planning"]'));
}

// ---------------------------------------------------------------- reporting page
{
  ok(/Reporting/.test(doc.querySelector('#viewTabs [data-view="reports"]').textContent),
    'the Reports tab reads Reporting');
  click(doc.querySelector('#viewTabs [data-view="reports"]'));
  const rp = doc.querySelector('#reportsView');
  ok(/Reporting/.test(rp.querySelector('h1').textContent), 'the page heading reads Reporting');
  ok(/Delivery by (sprint|week)/.test(rp.textContent), 'reporting includes sprint-level delivery');
  click(doc.querySelector('#viewTabs [data-view="planning"]'));
}

// ---------------------------------------------------------------- budgeting column resize
{
  click(doc.querySelector('#viewTabs [data-view="budget"]'));
  const rz = doc.querySelector('#hdr [data-burz="role"]');
  ok(!!rz, 'budget header columns grow resize handles');
  const before = parseInt(doc.documentElement.style.getPropertyValue('--bu-w-role'), 10);
  rz.dispatchEvent(new window.MouseEvent('pointerdown', { bubbles: true, cancelable: true, clientX: 200, button: 0 }));
  window.dispatchEvent(new window.MouseEvent('pointermove', { clientX: 260 }));
  window.dispatchEvent(new window.MouseEvent('pointerup', {}));
  const after = parseInt(doc.documentElement.style.getPropertyValue('--bu-w-role'), 10);
  ok(after === before + 60, 'dragging a handle widens the column (' + before + ' → ' + after + ')');
  ok(JSON.parse(window.localStorage.getItem('headway-ui-v1')).buColW.role === after,
    'budget column widths persist');
  const roleInp = doc.querySelector('#rows .row.brole[data-mid] input[data-bud="role"]');
  ok(roleInp.style.width === 'var(--bu-w-role)', 'row cells track the resized header width');
  click(doc.querySelector('#viewTabs [data-view="planning"]'));
}

// ------------------------------------------------- batch 11: header columns
{
  // detail selector lives in the top header line, left of Filter rows,
  // and carries a text label
  const dmBtn = doc.querySelector('#filterCell #detailBtn');
  ok(!!dmBtn, 'the detail selector sits inside the filter header cell');
  ok(/Feature|Story/.test(dmBtn.textContent), 'the detail selector shows a text label');

  // budget headers are rendered dynamically, aligned to the same visible set
  click(doc.querySelector('#viewTabs [data-view="budget"]'));
  const hdrKeys = Array.from(doc.querySelectorAll('#hlCols i[data-bucol]')).map(i => i.dataset.bucol);
  ok(hdrKeys.join(',') === 'role,type,cap,ws,cost,rate,margin,total', 'budget header renders all columns in order');
  ok(doc.querySelectorAll('#hlCols i[data-bucol][title]').length === hdrKeys.length,
    'every budget header carries an explanatory tooltip');
  const roleCell = doc.querySelector('#rows .row.brole[data-mid] input[data-bud="role"]');
  ok(roleCell && !roleCell.title, 'budget row cells carry no per-cell tooltips');
  const typeChip = doc.querySelector('#rows .row.brole[data-mid] [data-bact="type"]');
  ok(typeChip && !typeChip.title, 'rate-card chips carry no per-cell tooltips');

  // hide a column from the corner context menu — header AND rows drop it
  doc.querySelector('.hdr-left.corner').dispatchEvent(
    new window.MouseEvent('contextmenu', { bubbles: true, cancelable: true, clientX: 40, clientY: 40 }));
  const marginItem = Array.from(doc.querySelectorAll('#popover .menu-list button')).find(b => /Margin/.test(b.textContent));
  ok(!!marginItem, 'right-clicking the corner lists budget columns to show/hide');
  click(marginItem);
  ok(!doc.querySelector('#hlCols i[data-bucol="margin"]'), 'hiding Margin removes its header');
  ok(!doc.querySelector('#rows .row.brole[data-mid] [style*="--bu-w-margin"]'), 'hiding Margin removes its row cells');

  // drag a header to reorder (zero-rect jsdom drops it at the end)
  const roleHdr = doc.querySelector('#hlCols i[data-bucol="role"]');
  roleHdr.dispatchEvent(new window.MouseEvent('pointerdown', { bubbles: true, cancelable: true, clientX: 100, button: 0 }));
  window.dispatchEvent(new window.MouseEvent('pointermove', { clientX: 400 }));
  window.dispatchEvent(new window.MouseEvent('pointerup', { clientX: 400 }));
  const keysAfter = Array.from(doc.querySelectorAll('#hlCols i[data-bucol]')).map(i => i.dataset.bucol);
  ok(keysAfter[keysAfter.length - 1] === 'role', 'dragging a header reorders the column (' + keysAfter.join(',') + ')');
  ok(JSON.parse(window.localStorage.getItem('headway-ui-v1')).buColOrder.length > 0, 'column order persists');

  // reset restores everything
  doc.querySelector('.hdr-left.corner').dispatchEvent(
    new window.MouseEvent('contextmenu', { bubbles: true, cancelable: true, clientX: 40, clientY: 40 }));
  click(Array.from(doc.querySelectorAll('#popover .menu-list button')).find(b => /Reset columns/.test(b.textContent)));
  ok(Array.from(doc.querySelectorAll('#hlCols i[data-bucol]')).map(i => i.dataset.bucol).join(',') ===
    'role,type,cap,ws,cost,rate,margin,total', 'Reset columns restores the default set');

  // planning columns hide too
  click(doc.querySelector('#viewTabs [data-view="planning"]'));
  {
    const plKeys = Array.from(doc.querySelectorAll('#hlCols i[data-plcol]')).map(i => i.dataset.plcol);
    ok(['size', 'dur', 'asg'].every(k => plKeys.indexOf(k) !== -1), 'planning header lists size/dur/ppl (' + plKeys.join(',') + ')');
  }
  doc.querySelector('.hdr-left.corner').dispatchEvent(
    new window.MouseEvent('contextmenu', { bubbles: true, cancelable: true, clientX: 40, clientY: 40 }));
  click(Array.from(doc.querySelectorAll('#popover .menu-list button')).find(b => /^Dur/.test(b.textContent.trim())));
  ok(!doc.querySelector('#hlCols i[data-plcol="dur"]'), 'hiding Duration removes its planning header');
  ok(!doc.querySelector('#rows .row.item .r-wk'), 'hiding Duration removes the row chips');
  doc.querySelector('.hdr-left.corner').dispatchEvent(
    new window.MouseEvent('contextmenu', { bubbles: true, cancelable: true, clientX: 40, clientY: 40 }));
  click(Array.from(doc.querySelectorAll('#popover .menu-list button')).find(b => /Reset columns/.test(b.textContent)));
  ok(!!doc.querySelector('#rows .row.item .r-wk'), 'reset brings Duration back');
}

// ------------------------------------- planning columns: labels, resize, add/remove
{
  click(doc.querySelector('#viewTabs [data-view="planning"]'));
  const plLeftW = () => parseInt(doc.documentElement.style.getPropertyValue('--left-w'), 10);
  const plBaseLeftW = plLeftW();
  const plOrder = () => Array.from(doc.querySelectorAll('#hlCols i[data-plcol]')).map(i => i.dataset.plcol).join(',');
  const plMenu = (re) => {
    click(doc.querySelector('#plColsAdd'));
    click(Array.from(doc.querySelectorAll('#popover .menu-list button')).find(b => re.test(b.textContent.trim())));
  };
  const plDrag = (k, from, to) => {
    doc.querySelector('#hlCols [data-plrz="' + k + '"]').dispatchEvent(
      new window.MouseEvent('pointerdown', { bubbles: true, cancelable: true, clientX: from, button: 0 }));
    window.dispatchEvent(new window.MouseEvent('pointermove', { clientX: to }));
    window.dispatchEvent(new window.MouseEvent('pointerup', { clientX: to }));
  };
  // every visible planning column now carries its short label + tooltip
  const plHdrs = Array.from(doc.querySelectorAll('#hlCols i[data-plcol]'));
  ok(plHdrs.length > 0 && plHdrs.every(i => i.textContent.replace(/\s/g, '').length > 0),
    'planning headers show a label for every visible column');
  ok(plHdrs.every(i => !!i.title), 'planning headers carry their tooltips');

  // drag the header edge to resize
  const plrz = doc.querySelector('#hlCols [data-plrz="size"]');
  ok(!!plrz, 'planning header columns grow resize handles');
  const plBefore = parseInt(doc.documentElement.style.getPropertyValue('--pl-w-size'), 10);
  plrz.dispatchEvent(new window.MouseEvent('pointerdown', { bubbles: true, cancelable: true, clientX: 100, button: 0 }));
  window.dispatchEvent(new window.MouseEvent('pointermove', { clientX: 140 }));
  window.dispatchEvent(new window.MouseEvent('pointerup', {}));
  const plAfter = parseInt(doc.documentElement.style.getPropertyValue('--pl-w-size'), 10);
  ok(plAfter === plBefore + 40, 'dragging a planning handle widens the column (' + plBefore + ' → ' + plAfter + ')');
  ok(JSON.parse(window.localStorage.getItem('headway-ui-v1')).plColW.size === plAfter,
    'planning column widths persist');

  // the handle resizes and nothing else: a long drag on it must not reorder
  const plOrderBefore = plOrder();
  plDrag('pri', 100, 400);
  ok(plOrder() === plOrderBefore, 'dragging the resize handle never reorders the column');
  ok(parseInt(doc.documentElement.style.getPropertyValue('--pl-w-pri'), 10) === 240,
    'column widths clamp at the 240px maximum');
  plDrag('pri', 400, 0);
  ok(parseInt(doc.documentElement.style.getPropertyValue('--pl-w-pri'), 10) === 22,
    'and never shrink past the column minimum');

  // the + at the end of the strip opens the same columns menu
  ok(!!doc.querySelector('#plColsAdd'), 'the header strip ends with a + columns button');
  click(doc.querySelector('#plColsAdd'));
  const plDl = Array.from(doc.querySelectorAll('#popover .menu-list button'))
    .find(b => /^Deadline/.test(b.textContent.trim()));
  ok(!!plDl, 'the + button opens a columns menu offering Deadline');
  click(plDl);
  ok(!!doc.querySelector('#rows .row.item .r-date-col.dl-chip'),
    'turning Deadline on renders a deadline chip on feature rows');

  // the new columns are hidden until asked for, and stories inherit ws/epic
  click(doc.querySelector('#detailBtn'));
  click(doc.querySelector('#popover .menu-list [data-mi="1"]')); // Story detail
  plMenu(/^Workstream/);
  ok(!!doc.querySelector('#rows .row.item .r-ws-col'),
    'turning Workstream on renders a workstream chip on feature rows');
  ok(JSON.parse(window.localStorage.getItem('headway-ui-v1')).plColHide.ws === false,
    'a column switched on is recorded explicitly, so it survives a prefs reload');
  const plStWs = doc.querySelector('#rows .row.story .r-ws-col');
  ok(plStWs && plStWs.classList.contains('roll'),
    'story rows show the feature workstream rolled up, dimmed and non-clickable');

  // the Workstream column follows the project-wide switch, like every other view
  click(doc.querySelector('#btnSetup'));
  click(doc.querySelector('#setupView [data-sutab="org"]'));
  const plWsSw = doc.querySelector('#suWsEnable');
  plWsSw.checked = false;
  plWsSw.dispatchEvent(new window.Event('change', { bubbles: true }));
  click(doc.querySelector('#viewTabs [data-view="planning"]'));
  ok(!doc.querySelector('#hlCols i[data-plcol="ws"]') && !doc.querySelector('#rows .row.item .r-ws-col'),
    'Workstreams off hides the column even though it is switched on');
  click(doc.querySelector('#btnSetup'));
  click(doc.querySelector('#setupView [data-sutab="org"]'));
  const plWsSw2 = doc.querySelector('#suWsEnable');
  plWsSw2.checked = true;
  plWsSw2.dispatchEvent(new window.Event('change', { bubbles: true }));
  click(doc.querySelector('#viewTabs [data-view="planning"]'));
  ok(!!doc.querySelector('#hlCols i[data-plcol="ws"]'), 'and it returns when workstreams are back on');

  // feature and story date chips are separate fields on separate handlers
  plMenu(/^Start/);
  ok(!!doc.querySelector('#rows .row.item .r-date-col[data-act="startd"]') &&
     !!doc.querySelector('#rows .row.story .r-date-col[data-act="st-startd"]'),
    'feature and story date chips route to their own handlers');

  // the pane grows so the title keeps its floor when every column is on
  plMenu(/^Epic/);
  ok(plLeftW() > plBaseLeftW,
    'switching every column on widens the pane past the stored width (' + plBaseLeftW + ' → ' + plLeftW() + ')');

  // reset puts the defaults (and the default-hidden new columns) back
  plMenu(/Reset columns/);
  ok(!doc.querySelector('#rows .row.item .r-ws-col') && !doc.querySelector('#rows .row.item .r-date-col'),
    'Workstream, Epic, Start and Deadline are hidden by default');
  ok(!JSON.parse(window.localStorage.getItem('headway-ui-v1')).plColW.size,
    'Reset columns clears the planning widths too');
  ok(plLeftW() === plBaseLeftW, 'hiding them again restores the stored pane width');

  // dragging the pane must continue from the width actually on screen, and a
  // drag narrower than the columns need must preview the floor it will keep
  plMenu(/^Workstream/);
  plMenu(/^Epic/);
  const plDerived = plLeftW();
  ok(plDerived > plBaseLeftW, 'two wide columns push the pane past the stored width');
  const plRz = doc.querySelector('#leftRz');
  const paneDrag = (from, to) => {
    plRz.dispatchEvent(new window.MouseEvent('pointerdown', { bubbles: true, cancelable: true, clientX: from }));
    window.dispatchEvent(new window.MouseEvent('pointermove', { clientX: to }));
    return plLeftW();
  };
  const plWider = paneDrag(700, 710);
  ok(plWider === plDerived + 10,
    'a pane drag continues from the derived width (' + plDerived + ' + 10 → ' + plWider + ')');
  window.dispatchEvent(new window.MouseEvent('pointerup', {}));
  ok(plLeftW() === plDerived + 10, 'and the dragged width survives pointerup');
  ok(paneDrag(700, 700 - (plDerived + 10 - plBaseLeftW)) === plDerived,
    'dragging narrower than the columns need previews the floor, not a squeeze');
  window.dispatchEvent(new window.MouseEvent('pointerup', {}));
  plMenu(/Reset columns/);
  ok(plLeftW() === plBaseLeftW, 'and the stored width returns once the columns are hidden');
}


// ------------------------------------- row depth ramp + quiet title hover (stylesheet)
{
  const cssRamp = fs.readFileSync(path.join(ROOT, 'css/app.css'), 'utf8');
  const lightRoot = cssRamp.slice(cssRamp.indexOf(':root'), cssRamp.indexOf('html[data-theme="dark"]'));
  const darkRoot = cssRamp.slice(cssRamp.indexOf('html[data-theme="dark"]'), cssRamp.indexOf('* { box-sizing'));
  ['--lvl-phase', '--lvl-ws', '--lvl-epic', '--lvl-feature', '--lvl-story'].forEach((t) => {
    ok(lightRoot.indexOf(t + ':') !== -1 && darkRoot.indexOf(t + ':') !== -1,
      'both palettes define ' + t);
  });
  ok(/--band:\s*var\(--lvl-phase\)/.test(lightRoot) && !/--band:\s*#1A1F26/.test(lightRoot),
    'the light phase band is no longer a dark literal');
  ok(/--band-ink:\s*var\(--ink\)/.test(lightRoot), 'light bands carry dark ink');
  ok(/\.row\.story \.row-left[^{}]*\{[^}]*var\(--lvl-story\)/.test(cssRamp),
    'story rows paint from --lvl-story');
  ok(/\.row\.item \.row-left[^{}]*\{[^}]*var\(--lvl-feature\)/.test(cssRamp),
    'feature rows paint from --lvl-feature');
  ok(/\.row\.eband\.wsband[^{}]*\{[^}]*var\(--lvl-ws\)/.test(cssRamp),
    'workstream bands paint from --lvl-ws');

  // hover still says "this row" on both row kinds, one clear step off the ramp
  ['--lvl-feature-hover', '--lvl-story-hover', '--tip-bg', '--tip-ink'].forEach((t) => {
    ok(lightRoot.indexOf(t + ':') !== -1 && darkRoot.indexOf(t + ':') !== -1, 'both palettes define ' + t);
  });
  ok(/\.row\.item:not\(\.selected\)[^{]*:hover \.row-left\s*{[^}]*var\(--lvl-feature-hover\)/.test(cssRamp) &&
     /\.row\.story:not\(\.selected\)[^{]*:hover \.row-left\s*{[^}]*var\(--lvl-story-hover\)/.test(cssRamp),
    'feature and story rows both light up on hover');
  // overlays keep their own colours: a tooltip is a dark chip, the drag tip
  // takes the band's ink rather than a hard-coded white
  ok(/#uiTip\s*{[^}]*var\(--tip-bg\)[^}]*var\(--tip-ink\)/.test(cssRamp), 'tooltips stay dark chips');
  ok(!/#dragTip\s*{[^}]*color:\s*#fff/.test(cssRamp) &&
     /#dragTip\s*{[^}]*color:\s*var\(--band-ink\)/.test(cssRamp),
    'the drag tip reads on a light band');
  // things that now sit under the raised grid need a layer of their own
  ['.st-ghost', '.bu-cell', '.bu-costmark'].forEach((sel) => {
    ok(new RegExp(sel.replace('.', '\\.') + '\\s*{[^}]*z-index:\\s*2').test(cssRamp),
      sel + ' rises above the grid layer');
  });
  // Scoping is a grid, not a timeline: its left pane keeps the old tones
  ok(/body\[data-view="scoping"\] \.row\.item \.row-left\s*{[^}]*var\(--surface\)/.test(cssRamp) &&
     /body\[data-view="scoping"\] \.row\.story \.row-left\s*{[^}]*var\(--paper-2\)/.test(cssRamp),
    'Scoping keeps its own left-pane tones');
  // the exported board follows the light band too
  const pngSrc = fs.readFileSync(path.join(ROOT, 'js/export-png.js'), 'utf8');
  const pptxSrc = fs.readFileSync(path.join(ROOT, 'js/export-pptx.js'), 'utf8');
  ok(/BAND = '#E3DFD5'/.test(pngSrc) && !/#F4F6F8/.test(pngSrc), 'the PNG export paints a light phase band');
  ok(/BAND = 'E3DFD5'/.test(pptxSrc) && !/'F4F6F8'/.test(pptxSrc), 'the PPTX export paints a light phase band');

  // the title floor and the hover tints are Planning's, not Scoping's
  ok(/body:not\(\[data-view="scoping"\]\) #rows \.r-main,\s*\n\s*body:not\(\[data-view="scoping"\]\) \.hl-title\.pl-title\s*{[^}]*min-width:\s*120px/.test(cssRamp),
    'the title floor is scoped to Planning');
  ok(!/\n\.row\.item:not\(\.selected\)/.test(cssRamp) &&
     /body:not\(\[data-view="scoping"\]\) \.row\.item:not\(\.selected\)/.test(cssRamp) &&
     /body:not\(\[data-view="scoping"\]\) \.row\.story:not\(\.selected\)/.test(cssRamp),
    'the ramp hover tints skip Scoping');
  // one definition of Planning's pane width, shared by render, the drag and resize
  const appSrc = fs.readFileSync(path.join(ROOT, 'js/app.js'), 'utf8');
  ok((appSrc.match(/planLeftW\(\)/g) || []).length >= 4, 'planLeftW is the single source of the pane width');
  ok(/addEventListener\('resize'[^}]*planLeftW\(\)/.test(appSrc.replace(/\n/g, ' ')) ||
     /setTimeout[\s\S]{0,240}planLeftW\(\)/.test(appSrc),
    'a debounced resize listener re-derives the pane width');

  // titles show no box until they are being renamed
  ok(!/span\.r-name:hover\s*\{[^}]*border-color/.test(cssRamp),
    'plain-text row titles draw no hover outline');
  ok(!/\.st-title-txt:hover\s*\{[^}]*border-color/.test(cssRamp),
    'story titles draw no hover outline either');
}

// ------------------------------------------------- batch 11: detail modes
{
  // story detail: everything expanded, add-story everywhere, milestones bare
  click(doc.querySelector('#detailBtn'));
  click(doc.querySelector('#popover .menu-list [data-mi="1"]')); // Story
  const visItems = Array.from(doc.querySelectorAll('#rows .row.item'));
  const visNonMs = visItems.filter(r => {
    const it = state().items.find(i => i.id === r.dataset.id);
    return it && !it.milestone;
  });
  const visNoStories = visNonMs.filter(r => !(state().items.find(i => i.id === r.dataset.id).stories || []).length);
  ok(doc.querySelectorAll('#rows .row.story-add').length === visNoStories.length,
    'story detail gives every story-less non-milestone feature an add-story row (' +
    doc.querySelectorAll('#rows .row.story-add').length + '/' + visNoStories.length + ')');
  ok(!doc.querySelector('#rows .row.item .r-chev[data-act="stories"]'),
    'story detail disables the per-row expand toggles');
  ok(![...doc.querySelectorAll('#rows .row.item .r-chev svg')].length,
    'story detail shows no collapse chevrons at all');
  const msRow = visItems.find(r => {
    const it = state().items.find(i => i.id === r.dataset.id);
    return it && it.milestone;
  });
  if (msRow) {
    ok(!msRow.querySelector('.r-chev svg') && !doc.querySelector('#rows .row.story-add[data-id="' + msRow.dataset.id + '"]'),
      'milestones carry no story chevron and no add-story row');
  } else ok(true, '(no milestone visible)');

  // stories reorder / move across features by dragging their row
  const stRows = doc.querySelectorAll('#rows .row.story[data-story]');
  ok(stRows.length >= 1, 'story rows present for the drag test (' + stRows.length + ')');
  const src = stRows[0];
  const srcSt = src.dataset.story, srcItem = src.dataset.id;
  const allStoryRows = doc.querySelectorAll('#rows .row.story');
  const lastRow = allStoryRows[allStoryRows.length - 1];
  const destItem = lastRow.dataset.id;
  src.querySelector('.row-left').dispatchEvent(
    new window.MouseEvent('pointerdown', { bubbles: true, cancelable: true, clientX: 30, clientY: 30, button: 0 }));
  window.dispatchEvent(new window.MouseEvent('pointermove', { clientX: 30, clientY: 300 }));
  window.dispatchEvent(new window.MouseEvent('pointerup', { clientX: 30, clientY: 300 }));
  const moved = state().items.find(i => i.id === destItem).stories.some(s => s.id === srcSt);
  ok(moved, 'dragging a story row drops it into the targeted feature (' + srcItem + ' → ' + destItem + ')');
  window.dispatchEvent(new window.KeyboardEvent('keydown', { key: 'z', metaKey: true, bubbles: true }));
  click(doc.querySelector('#detailBtn'));
  click(doc.querySelector('#popover .menu-list [data-mi="0"]')); // back to Feature
}

// ------------------------------------------------- batch 11: history timeline
{
  const bar = doc.querySelector('#rows .bar:not(.ms)[data-bar]');
  click(doc.querySelector('#rows .row.item[data-id="' + bar.getAttribute('data-bar') + '"] .r-num'));
  const durInp = doc.querySelector('#panel input[data-f="durWeeks"]');
  durInp.value = String(parseFloat(durInp.value) + 2);
  durInp.dispatchEvent(new window.Event('change', { bubbles: true }));
  const lastEn = state().history[state().history.length - 1];
  ok(Array.isArray(lastEn.tl) && lastEn.tl.length > 0 && lastEn.tl[0].s1 != null,
    'schedule changes record machine-readable before/after positions');
  click(doc.querySelector('#btnHistory'));
  const tlTab = Array.from(doc.querySelectorAll('#historyView .hv-tab')).find(t => /Timeline/.test(t.textContent));
  ok(!!tlTab, 'the change lands under a Timeline tab');
  click(tlTab);
  ok(!!doc.querySelector('#historyView .vt-wrap'), 'the Timeline tab renders a mini timeline');
  ok(!!doc.querySelector('#historyView .vt-bar.new'), 'the new position paints green');
  const vtRow = doc.querySelector('#historyView .vt-row');
  ok(vtRow && /Duration/.test(vtRow.title), 'hovering a row lists the changed fields');
  click(doc.querySelector('#viewTabs [data-view="planning"]'));
}

// ---------------------------------------------------------------- options
// alternate plan versions: a dropdown right of the title. "Default" alone at
// first; New option duplicates the current plan and switches to it; switching
// swaps whole documents; Compare ghosts the other option's bars on the
// timeline; Close removes a finished version.
{
  click(doc.querySelector('#viewTabs [data-view="planning"]'));
  const optBtn = doc.querySelector('#optBtn');
  ok(!!optBtn && /Default/.test(optBtn.textContent), 'options button sits by the title and reads "Default"');
  ok(state().optName === 'Default' && (state().options || []).length === 0,
    'a fresh document holds only the Default option');
  const menuBtn = (re) =>
    [...doc.querySelectorAll('#popover .menu-list [data-mi]')].find(b => re.test(b.textContent.trim()));

  const rowAct = (re, title) => {
    const row = [...doc.querySelectorAll('#popover .menu-list [data-mi]')].find(b => re.test(b.textContent.trim()));
    return row && row.querySelector('.mi-act[title="' + title + '"]');
  };
  click(optBtn);
  ok(!!menuBtn(/New option/), 'options menu offers New option');
  ok(!/copy of the current plan/.test(doc.querySelector('#popover').textContent),
    'New option carries no explainer text');
  ok(!doc.querySelector('#popover .mi-act[title="Compare"]') &&
    !doc.querySelector('#popover .mi-act[title="Delete"]'),
    'no Compare or Delete actions while Default is alone');
  ok(!!rowAct(/^Default$/, 'Rename'), 'the option row carries a Rename action');
  click(menuBtn(/New option/));
  const nameIn = doc.querySelector('#modalHost #optNameIn');
  ok(!!nameIn, 'New option asks for a name');
  nameIn.value = 'Plan B';
  click(doc.querySelector('#modalHost #optNameOk'));
  ok(state().optName === 'Plan B', 'creating an option switches to it');
  ok(state().options.length === 1 && state().options[0].name === 'Default',
    'the previous plan is parked as Default');
  ok(/Plan B/.test(doc.querySelector('#optBtn').textContent), 'the button shows the active option');

  // diverge Plan B: push a scheduled feature's duration up by two weeks
  const bar = doc.querySelector('#rows .bar:not(.ms)[data-bar]');
  const vId = bar.getAttribute('data-bar');
  click(doc.querySelector('#rows .row.item[data-id="' + vId + '"] .r-num'));
  const durInp = doc.querySelector('#panel input[data-f="durWeeks"]');

  durInp.value = String(parseFloat(durInp.value) + 2);
  durInp.dispatchEvent(new window.Event('change', { bubbles: true }));
  const editedDays = state().items.find(i => i.id === vId).durDays;

  // compare with Default: the untouched schedule ghosts behind the live bars
  click(doc.querySelector('#optBtn'));
  ok(!!doc.querySelector('#popover .menu-list button.on [data-lucide="check"]'),
    'the active option marks itself with a leading check');
  ok(!doc.querySelector('#popover .menu-list .mi-check'),
    'no trailing checkmark on option rows');
  ok(!!rowAct(/^Default$/, 'Compare'), 'the other option row carries a Compare action');
  ok(!!rowAct(/^Default$/, 'Delete') && !!rowAct(/Plan B/, 'Delete'),
    'every option row carries a Delete action');
  click(rowAct(/^Default$/, 'Compare'));
  ok(doc.querySelectorAll('#rows .bar.cmp').length > 0, 'comparing paints ghost bars for differing items');
  const ghost = [...doc.querySelectorAll('#rows .row.item[data-id="' + vId + '"] .bar.cmp')][0];
  ok(!!ghost && /^Default: /.test(ghost.title), 'the edited feature carries a ghost titled with the option name');
  ok(!!doc.querySelector('#cmpPill') && /Default/.test(doc.querySelector('#cmpPill').textContent),
    'a pill announces what is being compared');
  click(doc.querySelector('#cmpPill [data-cmpx]'));
  ok(!doc.querySelector('#cmpPill') && !doc.querySelector('#rows .bar.cmp'),
    'dismissing the pill clears the overlay');

  // switch back to Default: the duration edit stays behind in Plan B
  click(doc.querySelector('#optBtn'));
  click(menuBtn(/^Default$/));
  ok(state().optName === 'Default', 'switching activates the picked option');
  const backDur = doc.querySelector('#rows .row.item[data-id="' + vId + '"] [data-act="wk"]');
  ok(state().items.find(i => i.id === vId).durDays !== editedDays,
    'Default still has the original schedule (' + backDur.textContent + 'w vs Plan B\'s edit)');
  ok(state().options.length === 1 && state().options[0].name === 'Plan B' &&
    state().options[0].doc.items.find(i => i.id === vId).durDays === editedDays,
    'Plan B is parked with its own edited schedule');
  // delete Plan B out via its row action
  click(doc.querySelector('#optBtn'));
  click(rowAct(/Plan B/, 'Delete'));
  ok(!!doc.querySelector('#modalHost [data-m="ok"]'), 'closing an option asks for confirmation');
  click(doc.querySelector('#modalHost [data-m="ok"]'));
  ok((state().options || []).length === 0 && state().optName === 'Default',
    'closing removes the option and keeps the current one');
  ok(JSON.parse(window.localStorage.getItem('headway-v1')).optName === 'Default',
    'the active option name autosaves with the document');
}

// ------------------------------------------------- prio view survives reload
// regression guard: applyUi's saved-view whitelist must include every tab, or
// reloading on it silently falls back to Planning. Boot a second window with
// the current session storage and assert Prioritizing is restored.
{
  // pick Story detail first — the level chosen must come back after a reload
  // (and default to Feature when nothing is stored)
  click(doc.querySelector('#detailBtn'));
  click([...doc.querySelectorAll('#popover .menu-list [data-mi]')].find(b => /Story/.test(b.textContent)));
  ok(JSON.parse(window.localStorage.getItem('headway-ui-v1')).detailMode === 'story',
    'the detail level is written to the ui snapshot');
  click(doc.querySelector('#viewTabs [data-view="prio"]'));
  ok(JSON.parse(window.localStorage.getItem('headway-ui-v1')).view === 'prio',
    'prio is written to the ui snapshot');
  const dom2 = new JSDOM(html, {
    url: 'http://localhost/roadmapping/index.html',
    runScripts: 'outside-only',
    pretendToBeVisual: true,
  });
  const w2 = dom2.window;
  w2.ExcelJS = ExcelJS;
  w2.localStorage.setItem('headway-v1', window.localStorage.getItem('headway-v1'));
  w2.localStorage.setItem('headway-ui-v1', window.localStorage.getItem('headway-ui-v1'));
  for (const f of ['js/core.js', 'js/excel.js', 'js/export-png.js', 'js/export-pptx.js', 'js/export-jira.js', 'js/jira.js', 'js/ai.js', 'js/app.js']) {
    w2.eval(fs.readFileSync(path.join(ROOT, f), 'utf8'));
  }
  const d2 = w2.document;
  d2.querySelector('#startBody [data-sp-continue]')
    .dispatchEvent(new w2.MouseEvent('click', { bubbles: true }));
  const onTab = d2.querySelector('#viewTabs button.on');
  ok(onTab && onTab.dataset.view === 'prio',
    'reload restores the Prioritizing tab (got ' + (onTab ? onTab.dataset.view : 'none') + ')');
  ok(!!d2.querySelector('#prioView .pr-table'), 'the prioritizing board renders after the reload');
  ok(/Story/.test(d2.querySelector('#detailBtn').textContent),
    'reload restores the chosen detail level (Story)');
  dom2.window.close();
  // a boot with no stored ui must land on Feature detail
  const dom3 = new JSDOM(html, { url: 'http://localhost/roadmapping/index.html', runScripts: 'outside-only', pretendToBeVisual: true });
  dom3.window.ExcelJS = ExcelJS;
  dom3.window.localStorage.setItem('headway-v1', window.localStorage.getItem('headway-v1'));
  for (const f of ['js/core.js', 'js/excel.js', 'js/export-png.js', 'js/export-pptx.js', 'js/export-jira.js', 'js/jira.js', 'js/ai.js', 'js/app.js']) {
    dom3.window.eval(fs.readFileSync(path.join(ROOT, f), 'utf8'));
  }
  dom3.window.document.querySelector('#startBody [data-sp-continue]')
    .dispatchEvent(new dom3.window.MouseEvent('click', { bubbles: true }));
  ok(/Feature/.test(dom3.window.document.querySelector('#detailBtn').textContent),
    'with nothing stored the detail level defaults to Feature');
  dom3.window.close();
  click(doc.querySelector('#viewTabs [data-view="planning"]'));
  // leave the long-lived suite window back on Feature detail
  click(doc.querySelector('#detailBtn'));
  click([...doc.querySelectorAll('#popover .menu-list [data-mi]')].find(b => /Feature/.test(b.textContent)));
}

// ---------------------------------------------- panel edits flush on click-out
// Click-out targets can suppress the blur that normally commits a panel field
// (drag surfaces preventDefault on pointerdown; macOS WebKit buttons never
// take focus) — re-rendering the panel must flush the pending edit first.
{
  // render() rebuilds #rows, so address rows by item id, freshly each time
  const scheduled = state().items.filter((i) => i.startDay != null &&
    !state().phases.find((p) => p.id === i.phaseId).collapsed);
  const [id1, id2] = [scheduled[0].id, scheduled[1].id];
  const rowFor = (id) => doc.querySelector('#rows .row.item[data-id="' + id + '"] .row-left');
  const itemById = (id) => state().items.find((i) => i.id === id);
  click(rowFor(id1));
  const ed = doc.querySelector('#panel .wz-ed[data-f="col:description"]');
  ok(!!ed, 'panel shows the rich description editor');
  ed.focus();
  ed.innerHTML = '<p>typed but never blurred</p>';
  click(rowFor(id2)); // select another row without a blur
  ok(/typed but never blurred/.test(window.RM.scopeValue(itemById(id1), 'description')),
    'switching selection commits a rich edit that never blurred');
  const din = doc.querySelector('#panel input[data-f="durWeeks"]');
  ok(!!din, 'panel shows the duration input for the second item');
  din.focus();
  din.value = '7';
  click(rowFor(id1)); // again: no blur, no change event
  const spw = window.RM.slotsOf(state().meta);
  ok(itemById(id2).durDays === Math.max(1, Math.round(7 * spw)),
    'switching selection commits a modified input that never fired change');
}

// ------------------------------------------------- story selection highlight
// A selected story carries the same selected look a selected feature does;
// its parent feature drops to a distinct, quieter marker — never the same
// background as the selection itself.
{
  click(doc.querySelector('#detailBtn'));
  click([...doc.querySelectorAll('#popover .menu-list [data-mi]')].find((b) => /Story/.test(b.textContent)));
  const stRow = doc.querySelector('#rows .row.story[data-story]');
  ok(!!stRow, 'a story row renders in Story detail');
  const host = state().items.find((i) => i.id === stRow.dataset.id);
  const stId = stRow.dataset.story;
  click(stRow.querySelector('.st-pad') || stRow);
  const stRow2 = doc.querySelector('#rows .row.story[data-story="' + stId + '"]');
  const itRow2 = doc.querySelector('#rows .row.item[data-id="' + host.id + '"]');
  ok(stRow2.classList.contains('selected'), 'clicking a story selects the story row');
  ok(!itRow2.classList.contains('selected') && itRow2.classList.contains('sel-parent'),
    'the parent feature shows the quiet parent marker, not the selected background');
  click(itRow2.querySelector('.row-left'));
  const itRow3 = doc.querySelector('#rows .row.item[data-id="' + host.id + '"]');
  ok(itRow3.classList.contains('selected') && !itRow3.classList.contains('sel-parent'),
    'selecting the feature itself restores the full selected background');
  click(doc.querySelector('#detailBtn'));
  click([...doc.querySelectorAll('#popover .menu-list [data-mi]')].find((b) => /Feature/.test(b.textContent)));
}

// ------------------------------------------------------------- multi-select
{
  const visIds = () => [...doc.querySelectorAll('#rows .row.item')].map((r) => r.dataset.id);
  const rowFor = (id) => doc.querySelector('#rows .row.item[data-id="' + id + '"] .row-left');
  const selRows = () => [...doc.querySelectorAll('#rows .row.item.selected')].map((r) => r.dataset.id);
  const key = (opts) => window.dispatchEvent(new window.KeyboardEvent('keydown', Object.assign({ bubbles: true }, opts)));
  const ids = visIds();
  click(rowFor(ids[0]));
  rowFor(ids[2]).dispatchEvent(new window.MouseEvent('click', { bubbles: true, shiftKey: true }));
  ok(selRows().join() === ids.slice(0, 3).join(),
    'shift-click selects the contiguous range from the anchor');
  key({ key: 'a', metaKey: true });
  ok(selRows().length === visIds().length && selRows().length > 3, 'Cmd+A selects all visible items');
  // group nudge: every selected, scheduled, unlocked item moves one day
  const before = {};
  state().items.forEach((i) => { before[i.id] = i.startDay; });
  key({ key: 'ArrowRight' });
  const movable = (i) => selRows().indexOf(i.id) !== -1 && before[i.id] != null && !i.locked;
  ok(state().items.filter(movable).every((i) => i.startDay === before[i.id] + 1),
    'ArrowRight nudges the whole selection');
  ok(state().items.filter((i) => !movable(i) && before[i.id] != null)
    .every((i) => i.startDay === before[i.id]), 'unselected items hold their start');
  key({ key: 'z', metaKey: true }); // undo the nudge
  key({ key: 'Escape' });
  ok(selRows().length === 1, 'Escape collapses a multi-selection to the anchor');
  // bulk edit through the context menu: set the epic on a range of two
  click(rowFor(ids[0]));
  rowFor(ids[1]).dispatchEvent(new window.MouseEvent('click', { bubbles: true, shiftKey: true }));
  rowFor(ids[0]).dispatchEvent(new window.MouseEvent('contextmenu', { bubbles: true, clientX: 10, clientY: 10 }));
  const epicBtn = [...doc.querySelectorAll('#popover .menu-list [data-mi]')].find((b) => /Set epic/.test(b.textContent));
  ok(!!epicBtn && /2 selected/.test(doc.querySelector('#popover').textContent),
    'right-click on a multi-selection offers bulk actions labeled with the count');
  click(epicBtn);
  const epicPick = [...doc.querySelectorAll('#popover .menu-list [data-mi]')].filter((b) => !/none|New epic/.test(b.textContent))[0];
  click(epicPick);
  const epicName = state().items.find((i) => i.id === ids[0]).epic;
  ok(!!epicName && state().items.find((i) => i.id === ids[1]).epic === epicName,
    'the picked epic lands on every selected item');
  key({ key: 'z', metaKey: true }); // undo the bulk epic
  // bulk delete asks once with the count, then removes the whole selection
  click(rowFor(ids[0]));
  rowFor(ids[1]).dispatchEvent(new window.MouseEvent('click', { bubbles: true, shiftKey: true }));
  key({ key: 'Delete' });
  const mh = () => doc.querySelector('#modalHost');
  ok(!mh().hidden && /2 items/.test(mh().textContent), 'Delete on a multi-selection confirms with the count');
  click(mh().querySelector('[data-m="ok"]'));
  ok(!state().items.some((i) => i.id === ids[0] || i.id === ids[1]),
    'confirming deletes every selected item');
  key({ key: 'z', metaKey: true }); // put them back
  ok(state().items.some((i) => i.id === ids[0]), 'undo restores the deleted selection');
}

// ------------------------------------------------------------ unsaved guard
{
  // the suite has edited the doc and never saved it — the guard must be live
  ok(window.HeadwayApp.unsavedNow() === true, 'edited, never-saved doc counts as unsaved');
  const mh = () => doc.querySelector('#modalHost'); // openModal swaps the node
  let proceeded = 0;
  window.HeadwayApp.guardUnsaved(() => proceeded++);
  ok(!mh().hidden && /Unsaved changes/.test(mh().textContent),
    'guard on an unsaved doc opens the warning dialog');
  ok(mh().querySelector('[data-m="gsave"]') && mh().querySelector('[data-m="gdiscard"]') &&
    mh().querySelector('[data-m="cancel"]'),
    'dialog offers Save, Don’t save and Cancel');
  click(mh().querySelector('[data-m="cancel"]'));
  ok(mh().hidden && proceeded === 0, 'Cancel keeps the doc and does not proceed');
  window.HeadwayApp.guardUnsaved(() => proceeded++);
  click(mh().querySelector('[data-m="gdiscard"]'));
  ok(mh().hidden && proceeded === 1, 'Don’t save closes the dialog and proceeds');
  // browser build: closing the tab with unsaved work asks the native question
  const bu = new window.Event('beforeunload', { cancelable: true });
  window.dispatchEvent(bu);
  ok(bu.defaultPrevented, 'beforeunload is blocked while work is unsaved');
}

// ---------------------------------------------------------------- story dependency arrows and ports
{
  click(doc.querySelector('#viewTabs [data-view="planning"]'));
  // arrows are drawn inside a requestAnimationFrame; jsdom's rAF is a timer,
  // so run callbacks inline for the length of this block (same trick the
  // focus-restore tests use)
  const realRaf = window.requestAnimationFrame;
  window.requestAnimationFrame = (cb) => cb();
  click(doc.querySelector('#detailBtn'));
  click(doc.querySelector('#popover .menu-list [data-mi="1"]')); // Story detail
  // a filter from an earlier block may hide some features: host the fixture on
  // a scheduled feature whose row is actually on screen
  const host = Array.from(doc.querySelectorAll('#rows .row.item[data-id]'))
    .map(r => state().items.find(i => i.id === r.dataset.id))
    .find(i => i && !i.milestone && i.startDay != null);
  window.HeadwayApp.ai.commit('story arrow fixture', (s) => {
    const f = window.RM.itemById(s, host.id);
    const n = window.RM.nextNum(s);
    f.stories = f.stories || [];
    f.stories.push({ id: 'ar_1', title: 'arrow one', done: false, num: n, startDay: f.startDay, durDays: 2 },
      { id: 'ar_2', title: 'arrow two', done: false, num: n + 1, startDay: f.startDay, durDays: 2, deps: [n] });
  });
  ok(!!doc.querySelector('#rows .st-bar[data-stbar="ar_1"] .port[data-port="out"]') && !!doc.querySelector('#rows .st-bar[data-stbar="ar_2"] .port[data-port="in"]'), 'story bars carry in/out ports');
  const edge = doc.querySelector('#arrowPaths g.edge[data-sfrom="ar_1"][data-sto="ar_2"]');
  ok(!!edge, 'a story dependency draws an arrow between the two story bars');
  ok(edge && edge.classList.contains('viol'), 'starting on the same day as the dependency marks the arrow as a violation');
  click(edge.querySelector('path.hit'));
  window.dispatchEvent(new window.KeyboardEvent('keydown', { key: 'Delete', bubbles: true }));
  ok(window.RM.storyRef(state(), 'ar_2').st.deps.length === 0, 'selecting the arrow and pressing Delete removes the story dependency');
  window.HeadwayApp.ai.commit('story arrow cleanup', (s) => { const f = window.RM.itemById(s, host.id); f.stories = f.stories.filter(x => x.id !== 'ar_1' && x.id !== 'ar_2'); });
  click(doc.querySelector('#detailBtn'));
  click(doc.querySelector('#popover .menu-list [data-mi="0"]')); // back to Feature detail
  window.requestAnimationFrame = realRaf;
}

// ---------------------------------------------------------------- Add buttons only where nothing exists yet
{
  click(doc.querySelector('#viewTabs [data-view="planning"]'));
  const withSt = state().items.find(i => i.stories && i.stories.length && !i.milestone);
  const noSt = state().items.find(i => (!i.stories || !i.stories.length) && !i.milestone);
  window.HeadwayApp.ai.commit('add-row probe', (s) => {
    s.phases.push({ id: 'ph_add_probe', name: 'Add probe', description: '', bucket: false, collapsed: false });
  });
  const addRows = [...doc.querySelectorAll('#rows .row.addrow')];
  ok(addRows.length === 1 && addRows[0].dataset.phase === 'ph_add_probe', 'only the empty phase shows an Add feature row');
  window.HeadwayApp.ai.commit('add-row probe cleanup', (s) => { s.phases = s.phases.filter(p => p.id !== 'ph_add_probe'); });
  // story add rows: only under a feature without stories
  window.__headway.selectItem(withSt.id);
  ok(!doc.querySelector('#panel [data-f="storyadd"]'), 'the panel of a feature with stories has no add-story box');
  if (noSt) {
    window.__headway.selectItem(noSt.id);
    ok(!!doc.querySelector('#panel [data-f="storyadd"]'), 'the panel of a feature without stories offers the add-story box');
  }
  click(doc.querySelector('#viewTabs [data-view="scoping"]'));
  const stAddRows = [...doc.querySelectorAll('#rows .row.story-add')];
  ok(stAddRows.every(r => !(state().items.find(i => i.id === r.dataset.id).stories || []).length),
    'story add rows appear only under features without stories');
  // Prioritizing: Add story button only on cards without stories
  click(doc.querySelector('#viewTabs [data-view="prio"]'));
  ok([...doc.querySelectorAll('#prioView [data-prstadd]')].every(b => {
    const card = b.closest('[data-prcard]');
    return !(state().items.find(i => i.id === card.dataset.prcard).stories || []).length;
  }), 'Prioritizing cards with stories have no Add story button');
  click(doc.querySelector('#viewTabs [data-view="planning"]'));
}

// ---------------------------------------------------------------- story context menu: insert above / below
{
  const host = state().items.find(i => i.stories && i.stories.length >= 1 && !i.milestone);
  window.__headway.selectItem(host.id);
  const first = host.stories[0];
  click(doc.querySelector('#panel [data-pst-edit="' + first.id + '"]'));
  doc.querySelector('#panel').dispatchEvent(new window.MouseEvent('contextmenu', { bubbles: true, cancelable: true, clientX: 300, clientY: 300 }));
  const below = [...doc.querySelectorAll('#popover .menu-list button')].find(b => /Insert story below/.test(b.textContent));
  const above = [...doc.querySelectorAll('#popover .menu-list button')].find(b => /Insert story above/.test(b.textContent));
  ok(!!below && !!above, 'the story context menu offers Insert story above and below');
  const nBefore = host.stories.length;
  click(below);
  let stories = window.RM.itemById(state(), host.id).stories;
  ok(stories.length === nBefore + 1 && stories[1].title === '' && stories[0].id === first.id, 'Insert below adds a blank story right after the anchor');
  const inserted = stories[1];
  ok(!!doc.querySelector('#panel textarea[data-stf="title"]') && doc.querySelector('#panel .p-crumb'), 'the new story opens in the panel');
  ok(doc.activeElement === doc.querySelector('#panel textarea[data-stf="title"]'), 'its title is focused for typing');
  doc.querySelector('#panel').dispatchEvent(new window.MouseEvent('contextmenu', { bubbles: true, cancelable: true, clientX: 300, clientY: 300 }));
  click([...doc.querySelectorAll('#popover .menu-list button')].find(b => /Insert story above/.test(b.textContent)));
  stories = window.RM.itemById(state(), host.id).stories;
  ok(stories.length === nBefore + 2 && stories[1].title === '' && stories[2].id === inserted.id, 'Insert above adds a blank story right before the anchor');
  // the Planning story row menu offers the same entries
  const row = doc.querySelector('#rows .row.story[data-story="' + first.id + '"]');
  if (row) {
    row.dispatchEvent(new window.MouseEvent('contextmenu', { bubbles: true, cancelable: true, clientX: 200, clientY: 200 }));
    ok(!![...doc.querySelectorAll('#popover .menu-list button')].find(b => /Insert story below/.test(b.textContent)), 'the story row menu offers Insert story below');
    window.dispatchEvent(new window.KeyboardEvent('keydown', { key: 'Escape', bubbles: true }));
  }
  window.HeadwayApp.ai.commit('insert probe cleanup', (s) => {
    const f = window.RM.itemById(s, host.id);
    f.stories = f.stories.filter(x => x.title !== '' || x.id === first.id);
  });
}

// ---------------------------------------------------------------- story numbers everywhere
// Every list that shows a story shows its #number, and the story panel edits
// it exactly like the feature number (shared pool, so a feature's number is
// refused and falls back to the next free one).
{
  click(doc.querySelector('#viewTabs [data-view="planning"]'));
  const host = state().items.find(i => !i.milestone && (i.stories || []).length);
  const st0 = host.stories[0];
  ok(typeof st0.num === 'number' && !state().items.some(i => i.num === st0.num), 'stories carry a number that no feature uses');
  click(doc.querySelector('#detailBtn'));
  click(doc.querySelector('#popover .menu-list [data-mi="1"]')); // Story detail
  const rowNum = doc.querySelector('#rows .row.story[data-story="' + st0.id + '"] .r-num');
  ok(!!rowNum && rowNum.textContent.trim() === '#' + st0.num, 'Planning story rows show the story number');
  click(doc.querySelector('#detailBtn'));
  click(doc.querySelector('#popover .menu-list [data-mi="0"]')); // back to Feature detail
  window.__headway.selectItem(host.id);
  const listNum = doc.querySelector('#panel .p-story[data-pst="' + st0.id + '"] .r-num');
  ok(!!listNum && listNum.textContent.trim() === '#' + st0.num, 'the feature panel story list shows numbers');
  click(doc.querySelector('#panel [data-pst-edit="' + st0.id + '"]'));
  const numIn = doc.querySelector('#panel input.p-num-edit[data-stf="num"]');
  ok(!!numIn && numIn.value === String(st0.num), 'the story panel header shows the editable number');
  const free = window.RM.nextNum(state()) + 5;
  numIn.value = String(free); numIn.dispatchEvent(new window.Event('change', { bubbles: true }));
  ok(window.RM.storyRef(state(), st0.id).st.num === free, 'changing the number renumbers the story');
  const feat = state().items.find(i => i.id !== host.id && !i.milestone);
  const numIn2 = doc.querySelector('#panel input.p-num-edit[data-stf="num"]');
  numIn2.value = String(feat.num); numIn2.dispatchEvent(new window.Event('change', { bubbles: true }));
  ok(window.RM.storyRef(state(), st0.id).st.num !== feat.num, 'a feature number is refused (falls back to a free one)');
  // Sprinting and Prioritizing: switch to the story level through the real
  // Feature / Story dropdown each view puts first in its toolbar
  click(doc.querySelector('#viewTabs [data-view="sprints"]'));
  click(doc.querySelector('#sprintView [data-spdd="level"]'));
  click([...doc.querySelectorAll('#popover .menu-list button')].find(b => /Story/.test(b.textContent)));
  ok(doc.querySelectorAll('#sprintView .spv-row.spv-st').length > 0 && !doc.querySelector('#sprintView .spv-row .r-num, #sprintView .spv-feat .r-num'), 'Sprinting rows show no numbers');
  click(doc.querySelector('#sprintView [data-spdd="level"]'));
  click([...doc.querySelectorAll('#popover .menu-list button')].find(b => /Feature/.test(b.textContent)));
  click(doc.querySelector('#viewTabs [data-view="prio"]'));
  click(doc.querySelector('#prioView [data-prdd="level"]'));
  click([...doc.querySelectorAll('#popover .menu-list button')].find(b => /Story/.test(b.textContent)));
  ok(doc.querySelectorAll('#prioView .pr-stcard, #prioView .pr-story').length > 0 && !doc.querySelector('#prioView .pr-card .r-num'), 'Prioritizing cards show no numbers');
  click(doc.querySelector('#prioView [data-prdd="level"]'));
  click([...doc.querySelectorAll('#popover .menu-list button')].find(b => /Feature/.test(b.textContent)));
  click(doc.querySelector('#viewTabs [data-view="planning"]'));
}
{
  // duplicating a feature gives the copied stories fresh numbers and remaps deps among them
  // pick a feature the Planning grid is actually showing (earlier blocks may
  // leave a filter on)
  const rowEl = [...doc.querySelectorAll('#rows .row.item[data-id]')]
    .find(r => { const it = state().items.find(i => i.id === r.dataset.id); return it && !it.milestone; });
  const host = state().items.find(i => i.id === rowEl.dataset.id);
  window.HeadwayApp.ai.commit('dup num fixture', (s) => {
    const f = window.RM.itemById(s, host.id);
    const n1 = window.RM.nextNum(s);
    f.stories.push({ id: 'dn_1', title: 'first', done: false, num: n1 }, { id: 'dn_2', title: 'second', done: false, num: n1 + 1, deps: [n1] });
  });
  const nBefore = state().items.length;
  const row = doc.querySelector('#rows .row.item[data-id="' + host.id + '"]');
  row.dispatchEvent(new window.MouseEvent('contextmenu', { bubbles: true, cancelable: true, clientX: 200, clientY: 200 }));
  click([...doc.querySelectorAll('#popover .menu-list button')].find(b => /^Duplicate$/.test(b.textContent.trim())));
  const copy = state().items[state().items.findIndex(i => i.id === host.id) + 1];
  ok(state().items.length === nBefore + 1 && copy && copy.id !== host.id, 'the feature is duplicated');
  const c1 = copy.stories.find(s => s.title === 'first'), c2 = copy.stories.find(s => s.title === 'second');
  const orig1 = window.RM.storyRef(state(), 'dn_1').st;
  ok(c1 && c2 && c1.num !== orig1.num && String(c2.deps) === String(c1.num), 'the copied story depends on the copied story’s new number');
  window.HeadwayApp.ai.commit('dup num cleanup', (s) => {
    s.items = s.items.filter(i => i.id !== copy.id);
    const f = window.RM.itemById(s, host.id); f.stories = f.stories.filter(x => x.id !== 'dn_1' && x.id !== 'dn_2');
  });
}

// ---------------------------------------------------------------- story dependencies in the panel
{
  const hosts = state().items.filter(i => !i.milestone).slice(0, 2);
  window.HeadwayApp.ai.commit('story dep fixture', (s) => {
    const f0 = window.RM.itemById(s, hosts[0].id), f1 = window.RM.itemById(s, hosts[1].id);
    f0.stories.push({ id: 'sd_a', title: 'Dep story alpha', done: false, num: window.RM.nextNum(s) });
    f1.stories.push({ id: 'sd_b', title: 'Dep story beta', done: false, num: window.RM.nextNum(s) + 1 });
  });
  const numA = window.RM.storyRef(state(), 'sd_a').st.num;
  window.__headway.selectItem(hosts[1].id);
  click(doc.querySelector('#panel [data-pst-edit="sd_b"]'));
  const sec = doc.querySelector('#panel .p-sec[data-sec="st-deps"]');
  ok(!!sec, 'the story panel has a Dependencies section');
  const secs = [...doc.querySelectorAll('#panel .p-sec')].map(e => e.dataset.sec);
  ok(secs.indexOf('st-deps') > secs.indexOf('st-schedule') && secs.indexOf('st-deps') < secs.indexOf('st-integrations'), 'it sits between Timeline and Integrations');
  const search = doc.querySelector('#panel input[data-stf="stdepsearch"]');
  ok(!!search, 'the section offers a search box');
  search.value = '#' + numA; search.dispatchEvent(new window.Event('input', { bubbles: true }));
  ok(!!doc.querySelector('#panel .dep-sug button[data-addstdep="' + numA + '"]'), 'typing #number suggests that story');
  search.value = 'alpha'; search.dispatchEvent(new window.Event('input', { bubbles: true }));
  const hit = doc.querySelector('#panel .dep-sug button[data-addstdep="' + numA + '"]');
  ok(!!hit && hit.textContent.indexOf('#' + numA) !== -1, 'typing a title suggests it with its number');
  click(hit);
  const depsOf = (id) => window.RM.storyRef(state(), id).st.deps;
  ok(String(depsOf('sd_b')) === String(numA), 'picking the suggestion adds the dependency');
  ok(!!doc.querySelector('#panel .dep-chip[data-stdepgo="sd_a"]'), 'the dependency shows as a chip');
  click(doc.querySelector('#panel .dep-chip[data-stdepgo="sd_a"]'));
  ok(!!doc.querySelector('#panel .p-crumb') && /alpha/.test(doc.querySelector('#panel textarea[data-stf="title"]').value), 'clicking the chip opens that story');
  ok(!!doc.querySelector('#panel .dep-chip button[data-strdep="sd_b"]'), 'the other side lists the dependent');
  click(doc.querySelector('#panel .dep-chip button[data-strdep="sd_b"]'));
  ok(depsOf('sd_b').length === 0, 'removing from the dependent side clears the link');
  // version history records story numbers and dependency changes
  const enDep = state().history[state().history.length - 1];
  ok(Array.isArray(enDep.d) && enDep.d.some(op => /Depends on/.test(op[1]) && /#/.test(op[2]) && op[3] === ''),
    'removing a story dependency records a Depends on diff row (' + JSON.stringify(enDep.d && enDep.d[0]) + ')');
  ok(enDep.d.some(op => /› #\d+ /.test(op[1])), 'story diff rows carry the story number');
  window.HeadwayApp.ai.commit('story dep fixture cleanup', (s) => {
    [hosts[0].id, hosts[1].id].forEach((id) => { const f = window.RM.itemById(s, id); f.stories = f.stories.filter(x => x.id !== 'sd_a' && x.id !== 'sd_b'); });
  });
}

// ---------------------------------------------------------------- Prioritizing: Unset column toggle, widths, card click → panel
{
  click(doc.querySelector('#viewTabs [data-view="prio"]'));
  const heads = () => [...doc.querySelectorAll('#prioView .pr-phhd')].map(h => h.childNodes[0].textContent.trim());
  const pick = (re) => click([...doc.querySelectorAll('#popover .menu-list button')].find(b => re.test(b.textContent)));
  // feature level, Columns → Size: the ladder plus Unset
  click(doc.querySelector('#prioView [data-prdd="level"]')); pick(/Feature/);
  click(doc.querySelector('#prioView [data-prdd="cols"]')); pick(/Size/);
  ok(heads()[heads().length - 1] === 'Unset', 'Size columns end with Unset by default');
  click(doc.querySelector('#prioView [data-prdd="cols"]'));
  const unsetItem = [...doc.querySelectorAll('#popover .menu-list button')].find(b => /Unset column/.test(b.textContent));
  ok(!!unsetItem, 'the Columns menu offers the Unset column toggle');
  click(unsetItem);
  ok(heads().indexOf('Unset') === -1, 'hiding the Unset column removes it from the board');
  ok(JSON.parse(window.localStorage.getItem('headway-ui-v1') || '{}').prioHideUnset === true, 'the choice persists in the browser prefs');
  click(doc.querySelector('#prioView [data-prdd="cols"]')); pick(/Unset column/);
  ok(heads()[heads().length - 1] === 'Unset', 'showing it again brings the column back');
  click(doc.querySelector('#prioView [data-prdd="cols"]')); pick(/Phase/);
  click(doc.querySelector('#prioView [data-prdd="cols"]'));
  ok(![...doc.querySelectorAll('#popover .menu-list button')].some(b => /Unset column/.test(b.textContent)), 'phase columns have no Unset, so no toggle');
  window.dispatchEvent(new window.KeyboardEvent('keydown', { key: 'Escape', bubbles: true }));
  // widths
  const cssW = fs.readFileSync(path.join(ROOT, 'css/app.css'), 'utf8');
  ok(!/\.sp-page\s*\{[^}]*max-width/.test(cssW), 'the Prioritizing page is full width (no max-width on .sp-page)');
  ok(/\.spv-main\s*\{[^}]*max-width:\s*\d+px/.test(cssW), 'the Sprinting main content area has a max-width');
  // clicking a card shows the panel for that item; clicking it again hides the panel
  window.__headway.selectItem(null); // start with nothing selected (an earlier test may have left a selection)
  const card = doc.querySelector('#prioView .pr-card[data-prcard]:not([data-prst])');
  const cid = card.dataset.prcard;
  click(card.querySelector('.pr-head') || card);
  ok(!doc.querySelector('#panel').hidden && !!doc.querySelector('#panel input.p-num-edit[data-f="num"]') &&
    doc.querySelector('#panel input.p-num-edit[data-f="num"]').value === String(state().items.find(i => i.id === cid).num),
    'clicking a card opens the panel on that feature');
  ok(doc.querySelector('#prioView .pr-card[data-prcard="' + cid + '"]').classList.contains('selected'), 'the selected card is marked');
  click(doc.querySelector('#prioView .pr-card[data-prcard="' + cid + '"] .pr-head'));
  ok(!doc.querySelector('#panel').hidden && doc.querySelector('#prioView .pr-card[data-prcard="' + cid + '"]').classList.contains('selected'),
    'clicking the selected card again keeps it selected');
  // story level: a story card opens the story panel
  click(doc.querySelector('#prioView [data-prdd="level"]')); pick(/Stor/);
  const sc = doc.querySelector('#prioView .pr-stcard[data-prst]');
  if (sc) {
    click(sc.querySelector('.pr-head'));
    ok(!doc.querySelector('#panel').hidden && !!doc.querySelector('#panel .p-crumb') &&
      doc.querySelector('#panel input.p-num-edit[data-stf="num"]').value === String(window.RM.storyRef(state(), sc.dataset.prst).st.num),
      'clicking a story card opens the story panel');
    click(doc.querySelector('#prioView .pr-stcard[data-prst="' + sc.dataset.prst + '"] .pr-head'));
    ok(!doc.querySelector('#panel').hidden && doc.querySelector('#prioView .pr-stcard[data-prst="' + sc.dataset.prst + '"]').classList.contains('selected'),
      'clicking it again keeps the story selected');
  }
  click(doc.querySelector('#prioView [data-prdd="level"]')); pick(/Feature/);
  click(doc.querySelector('#viewTabs [data-view="planning"]'));
}

// ---------------------------------------------------------------- Prioritizing swimlanes: the catch-all lane sits last
{
  click(doc.querySelector('#viewTabs [data-view="prio"]'));
  const pick = (re) => click([...doc.querySelectorAll('#popover .menu-list button')].find(b => re.test(b.textContent)));
  const lanes = () => [...doc.querySelectorAll('#prioView .pr-lanename')].map(e => e.textContent.trim());
  click(doc.querySelector('#prioView [data-prdd="group"]')); pick(/Epics/);
  ok(lanes().length > 1 && lanes()[lanes().length - 1] === 'No epic' && lanes()[0] !== 'No epic', 'grouped by epic, "No epic" is the last lane (' + lanes().join(' | ') + ')');
  if (state().meta.workstreamsEnabled) {
    click(doc.querySelector('#prioView [data-prdd="group"]')); pick(/Workstreams/);
    const dflt = window.RM.defaultWsName(state());
    ok(lanes().length > 1 && lanes()[lanes().length - 1] === dflt && lanes()[0] !== dflt, 'grouped by workstream, the default workstream is the last lane (' + lanes().join(' | ') + ')');
  }
  click(doc.querySelector('#prioView [data-prdd="level"]')); pick(/Stor/);
  click(doc.querySelector('#prioView [data-prdd="group"]')); pick(/Epics/);
  ok(lanes().length < 2 || lanes()[lanes().length - 1] === 'No epic', 'story level: "No epic" is last too');
  click(doc.querySelector('#prioView [data-prdd="level"]')); pick(/Feature/);
  click(doc.querySelector('#prioView [data-prdd="group"]')); pick(/None/);
  click(doc.querySelector('#viewTabs [data-view="planning"]'));
}

// ---------------------------------------------------------------- background right-click: no theme menu; empty Prioritizing cells say so
{
  click(doc.querySelector('#viewTabs [data-view="prio"]'));
  const ev = new window.MouseEvent('contextmenu', { bubbles: true, cancelable: true, clientX: 30, clientY: 30 });
  doc.querySelector('#prioView .sp-page').dispatchEvent(ev);
  ok(ev.defaultPrevented && doc.querySelector('#popover').hidden, 'right-clicking the app background opens nothing (and not the browser menu either)');
  const pick = (re) => click([...doc.querySelectorAll('#popover .menu-list button')].find(b => re.test(b.textContent)));
  click(doc.querySelector('#prioView [data-prdd="level"]')); pick(/Feature/);
  window.HeadwayApp.ai.commit('empty phase probe', (s) => { s.phases.push({ id: 'ph_pr_empty', name: 'Nobody here', description: '', bucket: false, collapsed: false }); });
  const emptyCol = doc.querySelector('#prioView .sp-col[data-prcol="ph_pr_empty"]');
  ok(!!emptyCol && emptyCol.querySelector('.sp-colbody .pr-empty') && /No items/.test(emptyCol.textContent), 'an empty column shows a "No items" placeholder');
  ok([...doc.querySelectorAll('#prioView .sp-col')].every(c => !!c.querySelector('.pr-card') !== !!c.querySelector('.pr-empty')), 'the placeholder appears exactly where there are no cards');
  window.HeadwayApp.ai.commit('empty phase probe cleanup', (s) => { s.phases = s.phases.filter(p => p.id !== 'ph_pr_empty'); });
  click(doc.querySelector('#viewTabs [data-view="planning"]'));
}

// ---------------------------------------------------------------- flags: context menu → dialog → orange flag in the alert slot
{
  click(doc.querySelector('#viewTabs [data-view="planning"]'));
  const ctxOn = (el) => el.dispatchEvent(new window.MouseEvent('contextmenu', { bubbles: true, cancelable: true, clientX: 200, clientY: 200 }));
  const menuBtn = (re) => [...doc.querySelectorAll('#popover .menu-list button')].find(b => re.test(b.textContent));
  const host = [...doc.querySelectorAll('#rows .row.item')].map(r => state().items.find(i => i.id === r.dataset.id)).find(i => i && !i.milestone && !i.flag);
  const row = () => doc.querySelector('#rows .row.item[data-id="' + host.id + '"]');
  ctxOn(row());
  ok(!!menuBtn(/^Flag…$/) && !menuBtn(/Unflag/), 'an unflagged row offers Flag…');
  click(menuBtn(/^Flag…$/));
  const ta = doc.querySelector('#modalHost textarea#flagReason');
  ok(!doc.querySelector('#modalHost').hidden && !!ta, 'Flag… opens a dialog with a reason box');
  ta.value = 'Waiting on security sign-off';
  click([...doc.querySelectorAll('#modalHost button')].find(b => /^Flag$/.test(b.textContent.trim())));
  const flagged = state().items.find(i => i.id === host.id);
  ok(flagged.flag && flagged.flag.reason === 'Waiting on security sign-off', 'confirming stores the flag and its reason');
  const badge = row().querySelector('.r-warn.flag');
  ok(!!badge && /Waiting on security/.test(badge.title) && !!badge.querySelector('[data-lucide="flag"]'), 'the row shows an orange flag in the alert slot with the reason as its tooltip');
  ok(row().classList.contains('flagged'), 'a flagged row carries the flagged class (orange left pane)');
  window.__headway.selectItem(host.id);
  ok(row().classList.contains('flagged') && row().classList.contains('selected'), 'selected + flagged combine (darker orange)');
  const cssF = fs.readFileSync(path.join(ROOT, 'css/app.css'), 'utf8');
  ok(/#rows \.row\.item \.row-left \.r-num, #rows \.row\.story \.row-left \.r-num \{ display: none; \}/.test(cssF), 'the left pane hides item and story numbers');
  ok(/\.row\.flagged \.row-left[^{]*\{[^}]*--flag-soft/.test(cssF) && /\.row\.flagged\.selected \.row-left[^{]*\{[^}]*--flag-sel/.test(cssF), 'flagged rows use the light orange at rest and the darker one when selected');
  ok(/--flag-sel:/.test(cssF.split('html[data-theme="dark"]')[1] || ''), 'the dark theme defines its own flag tints');
  ctxOn(row());
  ok(!!menuBtn(/Edit flag/) && !!menuBtn(/Unflag/), 'a flagged row offers Edit flag… and Unflag');
  click(menuBtn(/Unflag/));
  ok(!state().items.find(i => i.id === host.id).flag && !row().querySelector('.r-warn.flag'), 'Unflag clears it');
  // a story, from its row menu; an empty reason is fine
  window.HeadwayApp.ai.commit('flag probe story', (s) => { const f = window.RM.itemById(s, host.id); f.stories.push({ id: 'fl_st', title: 'flag me', done: false, num: window.RM.nextNum(s) }); });
  click(doc.querySelector('#detailBtn'));
  click(doc.querySelector('#popover .menu-list [data-mi="1"]'));
  const stRow = () => doc.querySelector('#rows .row.story[data-story="fl_st"]');
  ctxOn(stRow());
  click(menuBtn(/^Flag…$/));
  click([...doc.querySelectorAll('#modalHost button')].find(b => /^Flag$/.test(b.textContent.trim())));
  ok(window.RM.storyRef(state(), 'fl_st').st.flag && window.RM.storyRef(state(), 'fl_st').st.flag.reason === '', 'a story flags with an empty reason');
  ok(!!stRow().querySelector('.r-warn.flag') && stRow().querySelector('.r-warn.flag').title === 'Flagged', 'the story row shows the flag');
  ok(stRow().classList.contains('flagged'), 'a flagged story row is tinted too');
  click(doc.querySelector('#detailBtn'));
  click(doc.querySelector('#popover .menu-list [data-mi="0"]'));
  // Prioritizing card and Sprinting row carry the flag too
  window.HeadwayApp.ai.commit('flag probe feature', (s) => { window.RM.itemById(s, host.id).flag = { reason: 'card flag' }; });
  click(doc.querySelector('#viewTabs [data-view="prio"]'));
  ok(!!doc.querySelector('#prioView .pr-card[data-prcard="' + host.id + '"] .pr-flag'), 'the Prioritizing card shows the flag');
  click(doc.querySelector('#viewTabs [data-view="sprints"]'));
  ok(!!doc.querySelector('#sprintView .spv-row[data-spid="' + host.id + '"] .spv-flag') || !window.RM.itemById(state(), host.id).startDay, 'the Sprinting row shows the flag');
  window.__headway.selectItem(host.id);
  ok(!!doc.querySelector('#panel .p-flag') && /card flag/.test(doc.querySelector('#panel .p-flag').textContent), 'the panel header shows the flag and its reason');
  window.HeadwayApp.ai.commit('flag probe cleanup', (s) => { const f = window.RM.itemById(s, host.id); f.flag = null; f.stories = f.stories.filter(x => x.id !== 'fl_st'); });
  click(doc.querySelector('#viewTabs [data-view="planning"]'));
}

// ---------------------------------------------------------------- title double-click counted by hand; titles never select text
{
  click(doc.querySelector('#viewTabs [data-view="planning"]'));
  window.__headway.selectItem(null);
  const hostRow = [...doc.querySelectorAll('#rows .row.item')].find(r => r.querySelector('.r-name-txt'));
  const hid = hostRow.dataset.id;
  const titleOf = () => doc.querySelector('#rows .row.item[data-id="' + hid + '"] .r-name-txt');
  click(titleOf()); // selects (re-renders the row)
  ok(!doc.querySelector('#rows .row.item[data-id="' + hid + '"] input.r-name'), 'a single click on a title does not open the editor');
  click(titleOf()); // second click within the double-click window, on the re-rendered node
  const inp = doc.querySelector('#rows .row.item[data-id="' + hid + '"] input.r-name');
  ok(!!inp, 'two quick clicks on a title open the rename editor even though the row re-rendered in between');
  inp.value = 'Renamed by two clicks';
  inp.dispatchEvent(new window.KeyboardEvent('keydown', { key: 'Enter', bubbles: true }));
  ok(state().items.find(i => i.id === hid).feature === 'Renamed by two clicks', 'Enter commits the rename');
  window.dispatchEvent(new window.KeyboardEvent('keydown', { key: 'z', metaKey: true, bubbles: true }));
  const cssT = fs.readFileSync(path.join(ROOT, 'css/app.css'), 'utf8');
  ok(/\.r-name-txt,\s*\.st-title-txt\s*\{[^}]*user-select:\s*none/.test(cssT), 'title spans never select text (a selected word would hijack the row drag)');
}

// ---------------------------------------------------------------- ⌘B / ⌘I in rich editors
{
  click(doc.querySelector('#viewTabs [data-view="planning"]'));
  const hostK = [...doc.querySelectorAll('#rows .row.item')].map(r => state().items.find(i => i.id === r.dataset.id)).find(i => i && !i.milestone);
  window.__headway.selectItem(hostK.id);
  const ed = doc.querySelector('#panel .wz-ed[data-f="col:description"]');
  const calls = [];
  const orig = doc.execCommand;
  doc.execCommand = (c) => { calls.push(c); return true; };
  const key = (el, k, mods) => { const ev = new window.KeyboardEvent('keydown', Object.assign({ key: k, bubbles: true, cancelable: true }, mods)); el.dispatchEvent(ev); return ev; };
  ed.focus();
  let ev = key(ed, 'b', { metaKey: true });
  ok(ev.defaultPrevented && calls[calls.length - 1] === 'bold', '⌘B in a rich editor bolds');
  ev = key(ed, 'i', { ctrlKey: true });
  ok(ev.defaultPrevented && calls[calls.length - 1] === 'italic', 'Ctrl+I in a rich editor italicises');
  const n = calls.length;
  ev = key(ed, 'b', {});
  ok(!ev.defaultPrevented && calls.length === n, 'a plain b just types');
  const plain = doc.querySelector('#rows .sc-name') || doc.createElement('div');
  ev = key(plain, 'b', { metaKey: true });
  ok(calls.length === n, '⌘B outside a rich editor does nothing');
  doc.execCommand = orig;
}

// ---------------------------------------------------------------- story panel estimate buttons
// Size / Priority / Risk in the story panel are buttons with data-stf AND
// data-v; the generic data-stf handler must not swallow them.
{
  const hostIt = state().items.find(i => !i.milestone);
  window.HeadwayApp.ai.commit('estimate probe story', (s) => {
    const f = window.RM.itemById(s, hostIt.id);
    f.stories = f.stories || [];
    f.stories.push({ id: 'st_est_probe', title: 'estimate probe', done: false });
  });
  window.__headway.selectItem(hostIt.id);
  click(doc.querySelector('#panel [data-pst-edit="st_est_probe"]'));
  const stOf = () => window.RM.itemById(state(), hostIt.id).stories.find(x => x.id === 'st_est_probe');
  const sizeBtn = Array.from(doc.querySelectorAll('#panel button[data-stf="size"][data-v]')).find(b => b.dataset.v);
  ok(!!sizeBtn, 'story panel offers size buttons');
  if (sizeBtn) {
    click(sizeBtn);
    ok(stOf().size === sizeBtn.dataset.v, 'clicking a story size button sets the story size (' + sizeBtn.dataset.v + ')');
  }
  const riskBtn = Array.from(doc.querySelectorAll('#panel button[data-stf="risk"][data-v]')).find(b => b.dataset.v);
  ok(!!riskBtn, 'story panel offers risk buttons');
  if (riskBtn) {
    click(riskBtn);
    ok(stOf().risk === riskBtn.dataset.v, 'clicking a story risk button sets the story risk (' + riskBtn.dataset.v + ')');
  }
  click(doc.querySelector('#panel button[data-stf="size"][data-v=""]'));
  ok(!stOf().size, 'the — button clears the story size');
  window.HeadwayApp.ai.commit('estimate probe cleanup', (s) => {
    const f = window.RM.itemById(s, hostIt.id);
    f.stories = f.stories.filter(x => x.id !== 'st_est_probe');
  });
}

// ---------------------------------------------------------------- jira csv in the export dialog
{
  window.eval("document.querySelector('#btnExport').click()");
  const fmts = Array.from(doc.querySelectorAll('#modalHost input[name="exFmt"]')).map(r => r.id);
  ok(fmts.join(',') === 'exFmtPng,exFmtPptx,exFmtJira,exFmtHtml', 'Jira CSV is the third format, right of PowerPoint; standalone HTML is last');
  ok(/Jira CSV/.test(doc.querySelector('#modalHost label[for="exFmtJira"], #modalHost #exFmtJira').closest('label').textContent),
    'the third format is labeled Jira CSV');
  ok(doc.querySelector('#modalHost #exJira').hidden && !doc.querySelector('#modalHost #exTimeline').hidden,
    'Jira options stay hidden while a timeline format is chosen');
  doc.querySelector('#modalHost #exFmtJira').checked = true;
  doc.querySelector('#modalHost #exFmtJira').dispatchEvent(new window.Event('change', { bubbles: true }));
  ok(!doc.querySelector('#modalHost #exJira').hidden && doc.querySelector('#modalHost #exTimeline').hidden,
    'choosing Jira CSV swaps the timeline options for the Jira options');
  ok(!!doc.querySelector('#modalHost #jxFeatures') && !!doc.querySelector('#modalHost #jxStories'),
    'dialog offers feature and story rows');
  ok(/yyyy-MM-dd/.test(doc.querySelector('#modalHost #exJira').textContent), 'dialog names the wizard date format');
  ok(typeof window.RM_JIRA === 'object' && typeof window.RM_JIRA.csv === 'function', 'Jira export module is loaded');
  doc.querySelector('#modalHost #jxStories').checked = true;
  let exported = null;
  window.__headway.setExportSink((r) => { exported = r; });
  click(doc.querySelector('#modalHost #exGo'));
  ok(doc.querySelector('#modalHost').hidden, 'Export closes the dialog');
  ok(exported && /-jira\.csv$/.test(exported.name) && exported.blob && exported.blob.size > 100,
    'Export hands a CSV blob to the save path');
  const ui = JSON.parse(window.localStorage.getItem('headway-ui-v1'));
  ok(ui.exportPrefs.fmt === 'jira' && ui.jiraPrefs && ui.jiraPrefs.stories === true,
    'jira export settings persist in the ui snapshot');
  window.eval("document.querySelector('#btnExport').click()");
  ok(doc.querySelector('#modalHost #exFmtJira').checked && !doc.querySelector('#modalHost #exJira').hidden &&
    doc.querySelector('#modalHost #jxStories').checked,
    'reopening restores the Jira format and its settings');
  window.eval("document.querySelector('#modalHost [data-m=x]').click()");
  window.__headway.setExportSink(null);
}

// ---------------------------------------------------------------- export smoke
ok(window.__headway.saveFileName() === state().meta.title + '.xlsx',
  'save uses the exact project title as the filename (no slug, no date)');
// ---------------------------------------------------------------- panel field order + Integrations + column scope
{
  click(doc.querySelector('#btnSetup'));
  suTab('columns');
  doc.querySelector('#suColAdd').value = 'Zed';
  click(doc.querySelector('#suColAddBtn'));
  const zed = state().meta.scopeCols.find(c => c.label === 'Zed');
  click(doc.querySelector('#viewTabs [data-view="scoping"]'));
  const hdrOrder = () => Array.from(doc.querySelectorAll('#hdrSprints .sc-hcell[data-col]')).map(h => h.dataset.col);
  const zedHdr = () => doc.querySelector('#hdrSprints .sc-hcell[data-col="' + zed.key + '"]');
  const colMenu = (re) => {
    zedHdr().dispatchEvent(new window.MouseEvent('contextmenu', { bubbles: true, clientX: 300, clientY: 40 }));
    return Array.from(doc.querySelectorAll('#popover .menu-list button')).find(b => re.test(b.textContent));
  };
  let guard = 0;
  while (hdrOrder().indexOf(zed.key) > 0 && guard++ < 30) click(colMenu(/Move left/));
  ok(hdrOrder()[0] === zed.key, 'custom column moved to the first scoping column');
  // the feature panel lists only the columns that show on features (Acceptance criteria is stories-only by default)
  const textOrder = () => state().meta.scopeColOrder.filter(k => state().meta.scopeCols.some(c => c.key === k && window.RM.scopeColShows(c, 'feature')));
  const feat = doc.querySelector('#rows .row.item');
  click(feat.querySelector('.r-num'));
  const panelFields = () => Array.from(doc.querySelectorAll('#panel .p-sec[data-sec="fields"] .wz-ed[data-f^="col:"]')).map(e => e.dataset.f.slice(4));
  ok(panelFields().join(',') === textOrder().join(','),
    'panel Fields follow the scoping column order (' + panelFields().join(',') + ')');
  const secs = Array.from(doc.querySelectorAll('#panel .p-sec')).map(e => e.dataset.sec);
  ok(secs[secs.length - 1] === 'integrations' || (secs[secs.length - 1] === 'checks' && secs[secs.length - 2] === 'integrations'),
    'Integrations is the last panel section (' + secs.join(',') + ')');
  ok(!!doc.querySelector('#panel .p-sec[data-sec="integrations"] input[data-f="jiraKey"]'), 'Jira key lives under Integrations');
  ok(!doc.querySelector('#panel .p-sec[data-sec="details"] input[data-f="jiraKey"]'), 'Details no longer carries the Jira key');

  // story-only: the feature cell greys out, the panel omits the field
  click(colMenu(/Stories only/));
  ok(state().meta.scopeCols.find(c => c.key === zed.key).scope === 'story', 'column menu sets Stories only');
  ok(doc.querySelector('#rows .row.item .sc-cell.sc-na[data-col="' + zed.key + '"]') !== null, 'story-only column greys out on feature rows');
  ok(panelFields().indexOf(zed.key) === -1, 'story-only column leaves the feature panel');
  const withSt = state().items.find(i => i.stories && i.stories.length && !i.milestone);
  const chev = doc.querySelector('#rows .row.item[data-id="' + withSt.id + '"] .r-chev[data-act="stories"]');
  if (chev && !doc.querySelector('#rows .row.story[data-id="' + withSt.id + '"]')) click(chev);
  const stRow2 = doc.querySelector('#rows .row.story[data-id="' + withSt.id + '"]');
  ok(!!stRow2 && !!stRow2.querySelector('[data-stscope="' + zed.key + '"]'), 'story-only column still edits on story rows');
  click(doc.querySelector('#rows .row.item[data-id="' + withSt.id + '"] .r-num'));
  click(doc.querySelector('#panel [data-pst-edit="' + withSt.stories[0].id + '"]'));
  const stPanelFields = () => Array.from(doc.querySelectorAll('#panel .wz-ed[data-f^="stcol:"]')).map(e => e.dataset.f.slice(6));
  ok(stPanelFields()[0] === zed.key, 'story panel Fields follow the scoping order too');
  const stSecs = Array.from(doc.querySelectorAll('#panel .p-sec')).map(e => e.dataset.sec);
  ok(stSecs[stSecs.length - 1] === 'st-integrations' && !!doc.querySelector('#panel .p-sec[data-sec="st-integrations"] input[data-stf="jiraKey"]'),
    'story panel ends with Integrations holding the Jira key');

  // feature-only: story rows and the story panel drop it
  click(colMenu(/Features only/));
  ok(state().meta.scopeCols.find(c => c.key === zed.key).scope === 'feature', 'column menu sets Features only');
  const stRow3 = doc.querySelector('#rows .row.story[data-id="' + withSt.id + '"]');
  ok(!!stRow3 && !!stRow3.querySelector('.sc-cell.sc-na[data-col="' + zed.key + '"]'), 'feature-only column greys out on story rows');
  ok(stPanelFields().indexOf(zed.key) === -1, 'feature-only column leaves the story panel');
  click(colMenu(/Features and stories/));
  ok(state().meta.scopeCols.find(c => c.key === zed.key).scope == null, 'Features and stories clears the scope');
  // Setup lists the scope too
  click(doc.querySelector('#btnSetup'));
  suTab('columns');
  const scopeSel = doc.querySelector('#setupView [data-sucolscope="' + zed.key + '"]');
  ok(!!scopeSel, 'Columns tab offers a scope control per text column');
  scopeSel.value = 'story';
  scopeSel.dispatchEvent(new window.Event('change', { bubbles: true }));
  ok(state().meta.scopeCols.find(c => c.key === zed.key).scope === 'story', 'Setup scope control commits');
  click(doc.querySelector('#setupView [data-sucolrm="' + zed.key + '"]'));
  click(doc.querySelector('#viewTabs [data-view="planning"]'));
}

// ---------------------------------------------------------------- Enter commits + blurs; format bar hides until focus
{
  const enter = (el) => el.dispatchEvent(new window.KeyboardEvent('keydown', { key: 'Enter', bubbles: true, cancelable: true }));
  const row0 = Array.from(doc.querySelectorAll('#rows .row.item')).find(r => r.querySelector('.r-name-txt'));
  const it0 = state().items.find(i => i.id === row0.dataset.id);
  click(row0.querySelector('.r-num'));
  const pn = doc.querySelector('#panel .p-name');
  pn.focus();
  pn.value = 'Enter titled';
  const ev = enter(pn);
  ok(ev === false, 'Enter on the panel title is swallowed (no newline)');
  ok(doc.activeElement !== pn, 'Enter on the panel title leaves the field');
  ok(state().items.find(i => i.id === it0.id).feature === 'Enter titled', 'Enter on the panel title saves it');
  const jk2 = doc.querySelector('#panel input[data-f="jiraKey"]');
  jk2.focus();
  jk2.value = 'hw-77';
  enter(jk2);
  ok(doc.activeElement !== jk2, 'Enter on a panel input leaves the field');
  ok(state().items.find(i => i.id === it0.id).jiraKey === 'HW-77', 'Enter on a panel input saves it');
  doc.querySelector('#rows .row.item[data-id="' + it0.id + '"] .r-name-txt')
    .dispatchEvent(new window.MouseEvent('dblclick', { bubbles: true }));
  const rn = doc.querySelector('#rows .row.item[data-id="' + it0.id + '"] input.r-name');
  rn.focus();
  rn.value = 'Enter row';
  enter(rn);
  ok(state().items.find(i => i.id === it0.id).feature === 'Enter row', 'Enter on a left-pane name input saves it');
  ok(!doc.querySelector('#rows .row.item[data-id="' + it0.id + '"] input.r-name'),
    'and the input gives way to the text title again');
  const css = fs.readFileSync(path.join(ROOT, 'css/app.css'), 'utf8');
  ok(/\.wz-bar\s*\{[^}]*display:\s*none/.test(css) && /\.wz:focus-within\s*(>\s*)?\.wz-bar\s*\{[^}]*display:\s*flex/.test(css),
    'panel format bar stays hidden until its editor has focus');
}

// ---------------------------------------------------------------- stories drag in Planning too
{
  click(doc.querySelector('#viewTabs [data-view="planning"]'));
  click(doc.querySelector('#detailBtn'));
  click(doc.querySelector('#popover .menu-list [data-mi="1"]')); // Story detail: every feature open
  const seedRow = doc.querySelector('#rows .row.story[data-story]');
  // a feature that already has a story adds the next one from the story's context menu
  seedRow.dispatchEvent(new window.MouseEvent('contextmenu', { bubbles: true, cancelable: true, clientX: 200, clientY: 200 }));
  click([...doc.querySelectorAll('#popover .menu-list button')].find(b => /Insert story below/.test(b.textContent)));
  window.HeadwayApp.ai.commit('name story', (s) => {
    const f = window.RM.itemById(s, seedRow.dataset.id);
    f.stories.forEach((st) => { if (st.title === '') st.title = 'Second story'; });
  });
  const host = state().items.find(i => i.id === seedRow.dataset.id);
  ok(host.stories.length >= 2, 'feature now has two stories for the drag');
  const stRows = doc.querySelectorAll('#rows .row.story[data-story][data-id="' + host.id + '"]');
  ok(stRows.length === host.stories.length, 'planning lists the feature\'s story rows');
  ok(!!stRows[0].querySelector('.st-grip'), 'planning story rows carry a drag grip');
  const firstId = host.stories[0].id, lastId = host.stories[host.stories.length - 1].id;
  const left = stRows[0].querySelector('.row-left');
  left.dispatchEvent(new window.MouseEvent('pointerdown', { bubbles: true, button: 0, clientX: 10, clientY: 10 }));
  window.dispatchEvent(new window.MouseEvent('pointermove', { bubbles: true, clientX: 10, clientY: 60 }));
  window.dispatchEvent(new window.MouseEvent('pointerup', { bubbles: true, clientX: 10, clientY: 60 }));
  const after = state().items.find(i => i.id === host.id).stories.map(s2 => s2.id);
  // jsdom has no layout, so the drop resolves to the last story row on the page
  const lastStoryRow = Array.from(doc.querySelectorAll('#rows .row.story')).pop();
  const dest = state().items.find(i => i.stories.some(x => x.id === firstId));
  ok(!!dest && dest.id === lastStoryRow.dataset.id && (dest.id !== host.id || after[0] !== firstId),
    'dragging a story row in Planning moves it to the drop row\'s feature (' + (dest && dest.id) + ')');
  window.dispatchEvent(new window.KeyboardEvent('keydown', { key: 'z', metaKey: true, bubbles: true }));
  ok(state().items.find(i => i.id === host.id).stories[0].id === firstId, 'undo restores the story order');
  window.dispatchEvent(new window.KeyboardEvent('keydown', { key: 'z', metaKey: true, bubbles: true })); // drop the story's name
  window.dispatchEvent(new window.KeyboardEvent('keydown', { key: 'z', metaKey: true, bubbles: true })); // and the inserted story
  ok(state().items.find(i => i.id === host.id).stories.length === host.stories.length - 1, 'undo drops the inserted story');
  const upBtn = doc.querySelector('#panel [data-stf="up"]');
  if (upBtn) click(upBtn); // the insert opened the story in the panel; back to the feature
  click(doc.querySelector('#detailBtn'));
  click(doc.querySelector('#popover .menu-list [data-mi="0"]')); // back to Feature detail
}

// ---------------------------------------------------------------- titles rename on double-click only
{
  click(doc.querySelector('#viewTabs [data-view="planning"]'));
  click(doc.querySelector('#detailBtn'));
  click(doc.querySelector('#popover .menu-list [data-mi="1"]')); // Story detail: every feature open
  const stRow0 = doc.querySelector('#rows .row.story[data-story]');
  const sFid = stRow0.dataset.id, sSid = stRow0.dataset.story;
  const rowOf = () => doc.querySelector('#rows .row.story[data-story="' + sSid + '"]');
  const storyOf = () => state().items.find(i => i.id === sFid).stories.find(s2 => s2.id === sSid);
  const wasTitle = storyOf().title;
  const title0 = rowOf().querySelector('.st-title');
  ok(title0 && title0.tagName !== 'INPUT', 'a planning story title is text');
  click(title0);
  ok(!rowOf().querySelector('input'), 'a single click on a story title opens it, it does not edit it');
  rowOf().querySelector('.st-title').dispatchEvent(new window.MouseEvent('dblclick', { bubbles: true }));
  const sEd = rowOf().querySelector('input.st-add-input');
  ok(!!sEd, 'double-click makes a planning story title editable');
  sEd.value = 'Renamed story row';
  sEd.dispatchEvent(new window.KeyboardEvent('keydown', { key: 'Enter', bubbles: true }));
  ok(storyOf().title === 'Renamed story row', 'Enter commits the story rename');
  rowOf().dispatchEvent(new window.MouseEvent('contextmenu', { bubbles: true, cancelable: true, clientX: 200, clientY: 200 }));
  const sRn = [...doc.querySelectorAll('#popover .menu-list button')].find(b => /Rename/.test(b.textContent));
  ok(!!sRn, 'the story row context menu offers Rename…');
  click(sRn);
  const sEd2 = rowOf().querySelector('input.st-add-input');
  ok(!!sEd2, 'Rename… starts the same inline edit on a story row');
  sEd2.dispatchEvent(new window.KeyboardEvent('keydown', { key: 'Escape' }));
  window.dispatchEvent(new window.KeyboardEvent('keydown', { key: 'z', metaKey: true, bubbles: true }));
  ok(storyOf().title === wasTitle, 'undo puts the old story title back');
  click(doc.querySelector('#detailBtn'));
  click(doc.querySelector('#popover .menu-list [data-mi="0"]')); // back to Feature detail
  // the same on a Prioritizing feature card
  click(doc.querySelector('#viewTabs [data-view="prio"]'));
  const cid = doc.querySelector('#prioView .pr-card:not(.pr-stcard)').dataset.prcard;
  const cardOf = () => doc.querySelector('#prioView .pr-card[data-prcard="' + cid + '"]:not(.pr-stcard)');
  const featOf = () => state().items.find(i => i.id === cid);
  const wasFeat = featOf().feature;
  const ct = cardOf().querySelector('.pr-head .pr-title');
  ok(ct && ct.tagName !== 'INPUT' && ct.dataset.prf === 'feature', 'a prioritizing card title is text, not an input');
  click(ct);
  ok(!cardOf().querySelector('.pr-head input'), 'a single click on a card title does not start an edit');
  cardOf().querySelector('.pr-head .pr-title').dispatchEvent(new window.MouseEvent('dblclick', { bubbles: true }));
  const cEd = cardOf().querySelector('.pr-head input[data-prf="feature"]');
  ok(!!cEd, 'double-click makes the card title editable');
  cEd.value = 'Renamed on a card';
  cEd.dispatchEvent(new window.KeyboardEvent('keydown', { key: 'Enter', bubbles: true }));
  ok(featOf().feature === 'Renamed on a card', 'Enter commits the card rename');
  cardOf().dispatchEvent(new window.MouseEvent('contextmenu', { bubbles: true, cancelable: true, clientX: 90, clientY: 90 }));
  const cRn = [...doc.querySelectorAll('#popover .menu-list button')].find(b => /Rename/.test(b.textContent));
  ok(!!cRn, 'a prioritizing card context menu offers Rename…');
  click(cRn);
  ok(!!cardOf().querySelector('.pr-head input[data-prf="feature"]'), 'Rename… opens the card editor');
  cardOf().querySelector('.pr-head input[data-prf="feature"]')
    .dispatchEvent(new window.KeyboardEvent('keydown', { key: 'Escape' }));
  window.dispatchEvent(new window.KeyboardEvent('keydown', { key: 'z', metaKey: true, bubbles: true }));
  ok(featOf().feature === wasFeat, 'undo puts the old card title back');
  click(doc.querySelector('#viewTabs [data-view="planning"]'));
}

// ---------------------------------------------------------------- searchable menus, move to feature, rolled-up sizes
{
  const undo = () => window.dispatchEvent(new window.KeyboardEvent('keydown', { key: 'z', metaKey: true, bubbles: true }));
  const menuBtns = () => [...doc.querySelectorAll('#popover .menu-list button')];
  const typeIn = (inp, v) => { inp.value = v; inp.dispatchEvent(new window.Event('input', { bubbles: true })); };
  click(doc.querySelector('#viewTabs [data-view="planning"]'));
  const hostF = state().items.find(i => !i.milestone && (i.stories || []).length >= 1);
  // assignee picker: a search box and an avatar per person; typing narrows, Enter picks
  window.__headway.selectItem(hostF.id);
  click(doc.querySelector('#panel [data-dd="assign"]'));
  const srch = doc.querySelector('#popover .menu-search');
  ok(!!srch && doc.querySelectorAll('#popover .menu-list button .avatar').length === state().team.length,
    'the assignee picker is searchable and shows an avatar per person');
  const target = state().team[state().team.length - 1];
  typeIn(srch, window.RM.memberLabel(target));
  const vis = menuBtns().filter(b => !b.hidden);
  ok(vis.length >= 1 && vis.every(b => b.textContent.includes(window.RM.memberLabel(target))), 'typing narrows the roster');
  srch.dispatchEvent(new window.KeyboardEvent('keydown', { key: 'Enter', bubbles: true, cancelable: true }));
  ok((state().items.find(i => i.id === hostF.id).assignees || []).includes(target.id), 'Enter picks the first match');
  undo();
  // right-click a story row → Move to feature… → a searchable list of the other features
  click(doc.querySelector('#detailBtn'));
  click(doc.querySelector('#popover .menu-list [data-mi="1"]')); // Story detail
  const stRow = doc.querySelector('#rows .row.story[data-story]');
  const fromId = stRow.dataset.id, stId = stRow.dataset.story;
  stRow.dispatchEvent(new window.MouseEvent('contextmenu', { bubbles: true, cancelable: true, clientX: 200, clientY: 200 }));
  const mv = menuBtns().find(b => /^Move to feature/i.test(b.textContent.trim()));
  ok(!!mv, 'story rows offer Move to feature…');
  click(mv);
  const fsrch = doc.querySelector('#popover .menu-search');
  const dest = state().items.find(i => !i.milestone && i.id !== fromId && i.feature);
  ok(!!fsrch && menuBtns().some(b => b.textContent.includes(dest.feature)) && !menuBtns().some(b => b.textContent.includes(state().items.find(i => i.id === fromId).feature)),
    'the feature list is searchable, lists the other features and not the current one');
  typeIn(fsrch, dest.feature);
  click(menuBtns().find(b => !b.hidden));
  ok(!state().items.find(i => i.id === fromId).stories.some(s => s.id === stId) &&
    state().items.find(i => i.id === dest.id).stories.some(s => s.id === stId), 'the story moves to the chosen feature');
  undo();
  click(doc.querySelector('#detailBtn'));
  click(doc.querySelector('#popover .menu-list [data-mi="0"]')); // back to Feature detail
  // Setup → Sizing: "Roll up from stories" leads the feature list, never the story list
  click(doc.querySelector('#btnSetup'));
  suTab('est');
  const featSchemes = schemeOpts('size', 'feature');
  ok(featSchemes[0] === 'rollup' && schemeOpts('size', 'story').indexOf('rollup') === -1,
    'Roll up from stories is the first feature option and absent for stories');
  ok(state().meta.sizeScheme === 'tshirt', 'T-shirt sizes stay the default');
  pickScheme('size', 'feature', 'rollup');
  ok(state().meta.sizeScheme === 'rollup' && !!doc.querySelector('#setupView .m-hint') && !doc.querySelector('#setupView [data-susz]:not([data-kind="story"])'),
    'picking it explains the rollup instead of a size-options table');
  window.HeadwayApp.ai.commit('size stories', (s) => {
    const t = s.items.find(i => i.id === hostF.id);
    t.stories[0].size = '3';
    if (t.stories[1]) t.stories[1].size = '5';
  });
  const hf = () => state().items.find(i => i.id === hostF.id);
  const expectPts = String(hf().stories.reduce((a, st) => a + (isNaN(Number(st.size)) || !st.size ? 0 : Number(st.size)), 0));
  const expectDays = hf().stories.reduce((a, st) => a + (window.RM.sizeDays(state(), st.size, 'story') || 0), 0);
  ok(hf().size === expectPts, 'a feature size is the sum of its story points (' + expectPts + ')');
  ok(window.RM.itemSizeDays(state(), hf()) === expectDays, 'its working days are the sized stories\' days added up (' + expectDays + ')');
  click(doc.querySelector('#viewTabs [data-view="planning"]'));
  const szChip = doc.querySelector('#rows .row.item[data-id="' + hostF.id + '"] [data-act="size"]');
  ok(!!szChip && szChip.classList.contains('ro') && szChip.textContent.trim() === expectPts, 'the size chip shows the total and reads as read-only');
  ok(szChip.getAttribute('title') === 'Size (sum of story points)', 'its tooltip explains the rollup');
  click(szChip);
  ok(!doc.querySelector('#popover .menu-list'), 'clicking it opens no size menu');
  window.__headway.selectItem(hostF.id);
  ok(!!doc.querySelector('#panel .p-rollup') && !doc.querySelector('#panel [data-f="size"]'), 'the panel shows the rolled-up total instead of size buttons');
  click(doc.querySelector('#btnSetup'));
  suTab('est');
  pickScheme('size', 'feature', 'tshirt');
  ok(state().meta.sizeScheme === 'tshirt' && state().items.every(i => !i.size), 'back on T-shirt sizes the derived sizes are cleared');
  undo(); undo(); undo();
  ok(state().meta.sizeScheme === 'tshirt' && state().items.some(i => i.size), 'undo restores the hand-picked sizes');
  click(doc.querySelector('#viewTabs [data-view="planning"]'));
}

// ---------------------------------------------------------------- capacity types, planning level, capacity row, standalone HTML
{
  const undo = () => window.dispatchEvent(new window.KeyboardEvent('keydown', { key: 'z', metaKey: true, bubbles: true }));
  const menuBtns = () => [...doc.querySelectorAll('#popover .menu-list button')];
  ok(state().capTypes.slice(0, 3).join(',') === 'Development,Design,QA', 'a document starts with the default capacity types');
  // Setup → Capacity: capacity types list (rename / add / remove / reorder) + planning level + capacity row options
  click(doc.querySelector('#btnSetup'));
  suTab('scheduling');
  ok(doc.querySelectorAll('#setupView [data-sulist="captype"] .su-row').length === state().capTypes.length &&
    !!doc.querySelector('#setupView [data-sulist="captype"] .su-grip'), 'the Capacity tab lists the capacity types with reorder grips');
  doc.querySelector('#suCapTypeAdd').value = 'Research';
  click(doc.querySelector('#suCapTypeAddBtn'));
  ok(state().capTypes.indexOf('Research') !== -1, 'a capacity type can be added');
  const capIn = doc.querySelector('#setupView [data-capname="Research"]');
  capIn.value = 'Discovery'; capIn.dispatchEvent(new window.Event('change', { bubbles: true }));
  ok(state().capTypes.indexOf('Discovery') !== -1 && state().capTypes.indexOf('Research') === -1, 'a capacity type can be renamed');
  click(doc.querySelector('#setupView [data-sucaprm="Discovery"]'));
  ok(state().capTypes.indexOf('Discovery') === -1, 'a capacity type can be removed');
  undo(); undo(); undo();
  ok(doc.querySelectorAll('#setupView [data-suplan]').length === 2 && doc.querySelector('#setupView [data-suplan="feature"]').classList.contains('on'),
    'planning level offers Features (default) and Stories');
  click(doc.querySelector('#setupView [data-suplan="story"]'));
  ok(state().meta.planLevel === 'story' && window.RM.planLevel(state()) === 'story', 'picking Stories sets the planning level');
  undo();
  // Setup → Capacity tab
  window.HeadwayApp.openSetup ? window.HeadwayApp.openSetup('capacity') : window.HeadwayApp.ai.openSetup('capacity');
  ok(!!doc.querySelector('#suCapEnable') && !!doc.querySelector('[data-sucapmode="points"]'), 'the Capacity tab carries the enable switch and the demand picker');
  ok(!doc.querySelector('#suCapLimit') && !doc.querySelector('#suCapBasis'), 'the weekly limit and row basis controls are gone');
  ok(!doc.querySelector('#suDefPoints'), 'per-person demand hides the default-points input');
  click(doc.querySelector('#setupView [data-sucapmode="points"]'));
  ok(state().meta.capMode === 'points' && !!doc.querySelector('#suDefPoints'), 'picking story points switches the demand model and reveals the default points');
  click(doc.querySelector('#setupView [data-sucapmode="person"]'));
  ok(state().meta.capMode === 'person', 'and back to per person');
  // every capacity type counts: there is nothing to track or untrack
  ok(!doc.querySelector('#setupView [data-sucaprow]') && !/Tracked capacity types/.test(doc.querySelector('#setupView').textContent),
    'the Capacity tab has no tracked-type checkboxes');
  undo(); undo(); // the demand-model clicks
  undo(); undo(); undo();
  // people and stories carry a capacity type; assignability follows it
  const person = state().team[0], person2 = state().team[1];
  window.HeadwayApp.ai.commit('cap types', (s) => {
    s.team[0].capType = 'Design';
    s.team[1].capType = 'Development';
    const t = s.items.find(i => !i.milestone && (i.stories || []).length);
    t.stories[0].capType = 'Design';
  });
  const hostC = state().items.find(i => !i.milestone && (i.stories || []).length);
  const pool = window.RM.assignableFor(state(), hostC.stories[0]).map(m => m.id);
  ok(pool.length === 1 && pool[0] === person.id, 'a story with a capacity type is assignable only to people supplying it');
  ok(window.RM.assignableFor(state(), { capType: '' }).length === state().team.length, 'a story without a type is open to everyone');
  ok(window.RM.assignableFor(state(), { capType: 'Nobody' }).length === state().team.length, 'a type nobody supplies falls back to everyone');
  click(doc.querySelector('#viewTabs [data-view="budget"]'));
  const capChip = doc.querySelector('#rows [data-mid="' + person2.id + '"] [data-bact="cap"]');
  ok(!!capChip && capChip.textContent.trim() === 'Development', 'the Budgeting / Resources rows show a Capacity type chip');
  click(capChip);
  ok(menuBtns().some(b => /Design/.test(b.textContent)) && menuBtns()[0].textContent.indexOf('general') !== -1, 'the chip picks from the capacity types (or general)');
  window.dispatchEvent(new window.KeyboardEvent('keydown', { key: 'Escape', bubbles: true }));
  click(doc.querySelector('#viewTabs [data-view="planning"]'));
  window.__headway.selectItem(hostC.id);
  click(doc.querySelector('#panel [data-pst-edit="' + hostC.stories[0].id + '"]'));
  const stCapBtn = doc.querySelector('#panel [data-dd="stcap"]');
  ok(!!stCapBtn && /Design/.test(stCapBtn.textContent), 'the story panel shows its capacity type');
  click(doc.querySelector('#panel [data-dd="stassign"]'));
  const asgNames = menuBtns().map(b => b.textContent);
  ok(asgNames.length === 1 && asgNames[0].indexOf(window.RM.memberLabel(person)) !== -1, 'the story assignee picker lists only the people supplying its type');
  window.dispatchEvent(new window.KeyboardEvent('keydown', { key: 'Escape', bubbles: true }));
  undo();
  // capacity: story level plans on the stories and drains per type
  {
    const s = window.RM.clone(window.HeadwayApp.ai.state());
    s.meta.capacityEnabled = true; s.meta.planLevel = 'story'; s.meta.capMode = 'person';
    s.team = [{ id: 'p1', name: 'A', capType: 'Design', weekHours: {}, capacity: 1 }];
    const it = s.items.find((i) => i.startDay != null && !i.milestone);
    it.stories = [{ title: 'design it', startDay: it.startDay, durDays: 5, capType: 'Design' }, { title: 'later', capType: 'Design' }];
    const capS = window.RM.capacity(window.RM.normalizeState(s));
    const wk = Math.floor(it.startDay / 5);
    ok(capS.rows.Design[wk].demand === 1 && capS.rows.Design[wk].supply === 1, 'story level: one Design story vs one Design person');
    s.items.find((i) => i.id === it.id).stories[0].capMult = 2;
    ok(window.RM.capacity(window.RM.normalizeState(s)).weeks[wk].over, 'a ×2 story over-asks a one-person pool');
    s.meta.capRowTypes = ['Development']; // an old document's selection
    ok(window.RM.capacity(window.RM.normalizeState(s)).weeks[wk].over, 'an old tracked-type selection is ignored: every type counts');
  }
  // the header's single row sums demand / supply over every type
  window.HeadwayApp.ai.commit('cap row', (s) => {
    s.meta.capacityEnabled = true; s.meta.planLevel = 'feature'; s.meta.capMode = 'person';
    s.team = [{ id: 'p1', name: 'A', capType: 'Development', weekHours: {}, capacity: 1 }];
    s.items.forEach((i) => { if (!i.milestone) { i.capType = 'Development'; i.capMult = 1; } });
  });
  const capRow = () => doc.querySelectorAll('#hdrCapRows .hdr-cap');
  ok(capRow().length === 1 && capRow()[0].querySelector('.cap-row-lab').textContent === 'Capacity (people)',
    'one capacity row, labelled Capacity (people)');
  const cells = [...capRow()[0].querySelectorAll('.cap-cell')];
  ok(cells.length === 48, 'per person: 48 week cells');
  const busy = cells.find((c) => c.classList.contains('over'));
  ok(!!busy && /Development \d+(\.\d)? \/ \d+(\.\d)? — over/.test(busy.getAttribute('title')), 'an over-asked week reads over and lists the type in its tooltip');
  ok(/\d+(\.\d)? ?\/ ?\d+(\.\d)?/.test(busy.textContent), 'the cell reads demand / supply');
  ok(/people|points/.test(busy.getAttribute('title')), 'the tooltip says the unit');
  ok([...doc.querySelectorAll('#hdrCapRows .cap-cell')].every((c) => {
    const t = c.textContent.trim();
    return t === '' || t === '\u2715' || /^\d+(\.\d)?( ?\/ ?\d+(\.\d)?)?$/.test(t);
  }), 'no capacity cell renders a clipped fragment like “4 / 0.”');
  // a two-digit ask against a fractional supply cannot fit "14 / 0.9" in a
  // 28px week: the cell drops to the ask alone rather than clipping it
  window.HeadwayApp.ai.commit('fractional supply', (s) => { s.team[0].capacity = 0.9; });
  const frac = [...doc.querySelectorAll('#hdrCapRows .hdr-cap .cap-cell')]
    .find((c) => /^\d\d/.test(c.textContent.trim()));
  ok(!!frac && /^\d+$/.test(frac.textContent.trim()) && /0\.9 available/.test(frac.getAttribute('title')),
    'a two-digit ask against 0.9 available shows the ask alone, with both numbers in the tooltip');
  undo();
  // the sum can hide one type overflowing: the cell still reads over
  window.HeadwayApp.ai.commit('design bench', (s) => {
    s.team.push({ id: 'p2', name: 'B', capType: 'Design', weekHours: {}, capacity: 20 });
  });
  {
    const capH = window.RM.capacity(window.HeadwayApp.ai.state());
    const hi = capH.weeks.findIndex((c) => !c.blackout && c.demand > 1 && c.demand <= c.supply);
    const hc = [...doc.querySelectorAll('#hdrCapRows .hdr-cap .cap-cell')][hi];
    ok(hi !== -1 && capH.rows.Development[hi].over && hc.classList.contains('over'),
      'a week under its summed supply still reads over when Development alone is over');
    ok(hi !== -1 && /Development \d+(\.\d)? \/ \d+(\.\d)? — over/.test(hc.getAttribute('title')) && /Design 0 \/ \d+/.test(hc.getAttribute('title')),
      'and its tooltip lists each type\'s demand / supply (' + (hc && hc.getAttribute('title')) + ')');
  }
  undo();
  // a type nobody supplies can never be done: that reads over, not ok
  window.HeadwayApp.ai.commit('no supply', (s) => {
    s.items.forEach((i) => { if (!i.milestone) i.capType = 'Design'; });
  });
  const dry = [...doc.querySelectorAll('#hdrCapRows .hdr-cap .cap-cell')].find((c) => c.classList.contains('over'));
  ok(!!dry && / ?\/ ?1/.test(dry.textContent) && /no supply/.test(dry.getAttribute('title')),
    'a week asking a type nobody supplies reads over and says “no supply”');
  undo();
  // story-points mode: one header cell per sprint, lined up under the sprint cells
  {
    window.HeadwayApp.ai.commit('points header', (s) => {
      s.meta.capMode = 'points'; s.team[0].points = 10;
      // feature sizes are labels here: give one scheduled feature a points size
      const f = s.items.find((i) => !i.milestone && !i.done && i.startDay != null && i.startDay >= 0);
      f.size = '5';
    });
    const sprCells = [...doc.querySelectorAll('#hdrSprints .sprint-cell')];
    const pCells = [...doc.querySelectorAll('#hdrCapRows .hdr-cap .cap-cell')];
    const px = (el, k) => parseFloat(el.style[k]);
    ok(pCells.length === sprCells.length && pCells.length < 48, 'points mode: one capacity cell per sprint (' + pCells.length + ' vs ' + sprCells.length + ' sprints)');
    ok(pCells.every((c, i) => Math.abs(px(c, 'left') - 1 - px(sprCells[i], 'left')) < 0.01 &&
      Math.abs(px(c, 'width') + 2 - px(sprCells[i], 'width')) < 0.01), 'each sprint cell spans its sprint\'s weeks');
    const capP = window.RM.capacity(window.HeadwayApp.ai.state());
    const i0 = capP.weeks.findIndex((c) => c.demand > 0 && !c.blackout);
    ok(i0 !== -1 && pCells[i0].textContent.trim().indexOf(String(Math.round(capP.weeks[i0].supply))) !== -1 &&
      /points/.test(pCells[i0].getAttribute('title')) && /Sprint \d+/.test(pCells[i0].getAttribute('title')),
      'a sprint cell reads demand / supply for the sprint and names it in the tooltip (' + (i0 !== -1 ? pCells[i0].textContent : '') + ')');
    ok(pCells.every((c, i) => c.classList.contains('over') === capP.weeks[i].overAny || capP.weeks[i].blackout),
      'sprint cells read over exactly when the sprint is over');
    // a click lands on the week under the pointer: week 2 of the sprint
    {
      const m1 = state().meta;
      const ci = capP.periods.findIndex((p) => p.w1 - p.w0 === 2 && p.w0 > 0 &&
        window.RM.holidaysInWeek(m1, p.w0) === 0 && window.RM.holidaysInWeek(m1, p.w0 + 1) === 0);
      const w0 = capP.periods[ci].w0;
      const wkPx = px(sprCells[ci], 'width') / 2;
      const cell = doc.querySelectorAll('#hdrCapRows .hdr-cap .cap-cell')[ci];
      const r = cell.getBoundingClientRect();
      cell.dispatchEvent(new window.MouseEvent('click', { bubbles: true, clientX: r.left + wkPx * 1.5, clientY: 5 }));
      const m2 = state().meta, full = window.RM.slotsOf(m2);
      ok(window.RM.holidaysInWeek(m2, w0 + 1) === full && window.RM.holidaysInWeek(m2, w0) === 0,
        'clicking the second half of a sprint cell toggles the sprint\'s second week as a holiday week');
      undo();
    }
    undo();
    ok(doc.querySelectorAll('#hdrCapRows .hdr-cap .cap-cell').length === 48, 'back per person: 48 week cells');
  }
  undo();
  // Auto timeline: a one-shot ⚡ button on every real phase band
  {
    const zapOf = (pid) => doc.querySelector('#rows .row.band[data-phase="' + pid + '"] .band-zap');
    const snap0 = JSON.stringify(window.HeadwayApp.ai.state());
    window.HeadwayApp.ai.commit('cap off', (s) => { s.meta.capacityEnabled = false; });
    const realPhases = window.HeadwayApp.ai.state().phases.filter((p) => !p.bucket);
    const bucketPhases = window.HeadwayApp.ai.state().phases.filter((p) => p.bucket);
    ok(realPhases.every((p) => !!zapOf(p.id)), 'every real phase band has an Auto timeline button');
    ok(bucketPhases.every((p) => !zapOf(p.id)), 'a backlog bucket has none');
    ok(realPhases.every((p) => zapOf(p.id).disabled && /scheduling/.test(zapOf(p.id).getAttribute('title') || zapOf(p.id).dataset.tip || '')),
      'with scheduling off every button is disabled and says why');
    ok(!!zapOf(realPhases[0].id).querySelector('[data-lucide="zap"], svg'), 'the button carries the zap icon');
    const ph0 = realPhases[0].id;
    let lockedId = null, lockedStart = null;
    window.HeadwayApp.ai.commit('auto setup', (s) => {
      s.meta.capacityEnabled = true; s.meta.planLevel = 'feature'; s.meta.capMode = 'person';
      s.team = [{ id: 'solo', name: 'Solo', capType: 'Development', weekHours: {}, capacity: 1 }];
      // every feature now carries a capacity type, so park the other phases'
      // work as done — this check is about what the phase's button lays out
      s.items.forEach((it) => {
        if (it.phaseId === ph0) { it.locked = false; it.done = false; it.capType = 'Development'; it.capMult = 1; }
        else it.done = true;
      });
      // one scheduled item is locked: a fixed point the layout works around
      const lk = s.items.find((it) => it.phaseId === ph0 && !it.milestone && it.startDay != null);
      lk.locked = true; lockedId = lk.id; lockedStart = lk.startDay;
    });
    const dry = window.RM.autoPhase(window.HeadwayApp.ai.state(), ph0, { snap: { feature: window.HeadwayApp.ai.ui().snapFeat, story: window.HeadwayApp.ai.ui().snapStory } });
    ok(dry.changed > 0, 'the seeded phase is out of place (a dry run would move ' + dry.changed + ')');
    ok(!zapOf(ph0).disabled, 'so its button is enabled');
    ok(realPhases.slice(1).every((p) => zapOf(p.id).disabled ===
      (window.RM.autoPhase(window.HeadwayApp.ai.state(), p.id, { snap: { feature: window.HeadwayApp.ai.ui().snapFeat, story: window.HeadwayApp.ai.ui().snapStory } }).changed === 0)),
      'every other band: disabled exactly when a dry run changes nothing');
    // the band menu and the phase dialog offer the same action
    const band0 = doc.querySelector('#rows .row.band[data-phase="' + ph0 + '"]');
    band0.dispatchEvent(new window.MouseEvent('contextmenu', { bubbles: true, cancelable: true, clientX: 10, clientY: 10 }));
    const mAuto = [...doc.querySelectorAll('#popover .menu-list button[data-mi]')].find((b) => /Auto timeline/.test(b.textContent));
    ok(!!mAuto && !mAuto.disabled, 'the band menu has an enabled Auto timeline entry');
    doc.querySelector('#popover').hidden = true;
    click(band0.querySelector('[data-act="phase-edit"]'));
    const dlgAuto = doc.querySelector('#phAutoRun');
    ok(!!dlgAuto && !dlgAuto.disabled && /Auto timeline/.test(dlgAuto.textContent), 'the phase dialog has an enabled Auto timeline button');
    ok(!doc.querySelector('#phAuto'), 'and no Auto timeline checkbox any more');
    click(doc.querySelector('[data-m="x2"]'));
    // click: one commit lays the phase out
    const before = JSON.stringify(window.HeadwayApp.ai.state().items);
    const hLen = window.HeadwayApp.ai.state().history.length;
    [...doc.querySelectorAll('#toasts .toast')].forEach((t) => t.remove());
    click(zapOf(ph0));
    const after = window.HeadwayApp.ai.state();
    ok(JSON.stringify(after.items) !== before, 'clicking the button moves the phase’s items');
    ok(!window.RM.capacity(after).weeks.some((c) => c.over), 'after the click no week is over capacity');
    ok(window.RM.itemById(after, lockedId).startDay === lockedStart, 'the locked item keeps its start');
    ok(after.history.length === hLen + 1 && after.history[after.history.length - 1].label === 'auto timeline',
      'the click is one version-history entry, labelled auto timeline');
    ok([...doc.querySelectorAll('#toasts .toast')].some((t) => /Auto timeline moved \d+ item/.test(t.textContent)), 'and toasts how many items moved');
    {
      const was = JSON.parse(before);
      const nMoved = after.items.filter((b) => { const a = was.find((x) => x.id === b.id); return a && (a.startDay !== b.startDay || a.durDays !== b.durDays); }).length;
      ok([...doc.querySelectorAll('#toasts .toast')].some((t) => t.textContent.indexOf('moved ' + nMoved + ' item') !== -1),
        'the toast counts the units that moved (' + nMoved + ')');
    }
    ok(zapOf(ph0).disabled && /already in place/.test(zapOf(ph0).getAttribute('title') || zapOf(ph0).dataset.tip || ''),
      'right after the click the button is disabled: everything is in place');
    ok(!after.phases.some((p) => 'auto' in p), 'no phase carries an auto flag');
    undo();
    ok(JSON.stringify(window.HeadwayApp.ai.state().items) === before, 'one undo restores every item');
    ok(!zapOf(ph0).disabled, 'and the button is enabled again');
    // the assistant hook: a named phase, or every real phase
    ok(window.HeadwayApp.ai.autoTimelineNow(realPhases[0].name) > 0, 'autoTimelineNow runs a phase by name');
    window.HeadwayApp.ai.autoTimelineNow();
    ok(realPhases.every((p) => zapOf(p.id).disabled), 'autoTimelineNow() runs every real phase: every button is disabled after');
    ok(window.HeadwayApp.ai.autoTimelineNow(ph0) === 0, 'and a second run of the same phase changes nothing');
    // a snap picked in the View menu re-renders the rows: the buttons follow the new snap
    {
      const zOld = zapOf(ph0);
      const curFeat = window.HeadwayApp.ai.ui().snapFeat;
      const other = curFeat === 'day' ? 'sprint' : 'day';
      const pickSnap = (mode) => {
        click(doc.querySelector('.menu-btn[data-menu="view"]'));
        const want = new RegExp('snap to ' + { day: 'day', week: 'week', sprint: 'sprint' }[mode], 'i');
        const b = [...doc.querySelectorAll('#popover .menu-list button[data-mi]')].filter((x) => want.test(x.textContent))[0];
        if (b) click(b);
        return !!b;
      };
      ok(pickSnap(other), 'the View menu offers a feature snap');
      ok(window.HeadwayApp.ai.ui().snapFeat === other, 'and picking it switches the feature snap');
      ok(zapOf(ph0) !== zOld, 'the rows re-render, so the band button is rebuilt');
      const dryO = window.RM.autoPhase(window.HeadwayApp.ai.state(), ph0, { autoOrder: window.HeadwayApp.ai.ui().autoOrder, snap: { feature: other, story: window.HeadwayApp.ai.ui().snapStory } });
      ok(zapOf(ph0).disabled === (dryO.changed === 0), 'and its disabled state matches a dry run under the new snap');
      pickSnap(curFeat);
      ok(window.HeadwayApp.ai.ui().snapFeat === curFeat, 'the snap is back');
    }
    // the phase dialog's Auto timeline button saves the dialog's edits first
    {
      const st0 = window.HeadwayApp.ai.state();
      const open = st0.items.filter((i) => i.phaseId === ph0 && !i.milestone && !i.locked && !i.done && i.startDay != null);
      window.HeadwayApp.ai.commit('overlap', (s) => { s.items.find((i) => i.id === open[1].id).startDay = open[0].startDay; });
      click(doc.querySelector('#rows .row.band[data-phase="' + ph0 + '"] [data-act="phase-edit"]'));
      const nameIn = doc.querySelector('#phName');
      const oldName = nameIn.value;
      nameIn.value = oldName + ' renamed';
      const run = doc.querySelector('#phAutoRun');
      ok(!!run && !run.disabled, 'the dialog button is enabled for an overlapping phase');
      click(run);
      const st1 = window.HeadwayApp.ai.state();
      ok(st1.phases.find((p) => p.id === ph0).name === oldName + ' renamed', 'the dialog\u2019s pending rename was saved, not dropped');
      ok(st1.history[st1.history.length - 1].label === 'auto timeline', 'and then the action ran');
      ok(zapOf(ph0).disabled, 'leaving the phase in place');
      // a unit the roster can never fit is left where it is, and the toast says so
      window.HeadwayApp.ai.commit('crowd', (s) => {
        const a = s.items.find((i) => i.id === open[0].id), b = s.items.find((i) => i.id === open[1].id);
        a.capMult = 3; b.startDay = a.startDay;
      });
      [...doc.querySelectorAll('#toasts .toast')].forEach((t) => t.remove());
      if (!zapOf(ph0).disabled) {
        click(zapOf(ph0));
        ok([...doc.querySelectorAll('#toasts .toast')].some((t) => / \u00b7 1 could not be placed/.test(t.textContent)),
          'the toast says a unit could not be placed');
      } else ok(false, 'a crowded phase enables the button');
    }
    ok(/body\.ro[^{]*\.band-zap/.test(fs.readFileSync(require('path').join(__dirname, '..', 'css', 'app.css'), 'utf8')),
      'read-only documents hide the Auto timeline button');
    // editing capacity no longer toasts about a switched-off Auto flag
    window.HeadwayApp.ai.commit('cap off', (s) => { s.meta.capacityEnabled = false; });
    ok(zapOf(ph0).disabled, 'turning capacity planning off disables the button');
    window.HeadwayApp.ai.commit('restore', (s) => {
      const d = JSON.parse(snap0);
      s.meta = d.meta; s.phases = d.phases; s.items = d.items; s.team = d.team;
    });
  }
  // Place at earliest slot: right-click a feature row
  {
    window.HeadwayApp.ai.commit('cap off', (s) => { s.meta.capacityEnabled = false; });
    const plRow = doc.querySelector('#rows .row.item');
    plRow.dispatchEvent(new window.MouseEvent('contextmenu', { bubbles: true, cancelable: true, clientX: 10, clientY: 10 }));
    ok(![...doc.querySelectorAll('#popover .menu-list button[data-mi]')].some((b) => /Place at earliest slot/.test(b.textContent)),
      'with capacity planning off the row menu has no Place at earliest slot');
    doc.querySelector('#popover').hidden = true;
    window.HeadwayApp.ai.commit('cap on', (s) => { s.meta.capacityEnabled = true; });
    const plRow2 = doc.querySelector('#rows .row.item');
    plRow2.dispatchEvent(new window.MouseEvent('contextmenu', { bubbles: true, cancelable: true, clientX: 10, clientY: 10 }));
    const entry = [...doc.querySelectorAll('#popover .menu-list button[data-mi]')].find((b) => /Place at earliest slot/.test(b.textContent));
    ok(!!entry, 'feature context menu offers Place at earliest slot when capacity planning is on');
    doc.querySelector('#popover').hidden = true;
    const withSt = window.HeadwayApp.ai.state().items.find((i) => i.stories.length);
    if (!doc.querySelector('#rows .row.story[data-story]')) {
      click(doc.querySelector('#rows .row.item[data-id="' + withSt.id + '"] [data-act="stories"]'));
    }
    const plStRow = doc.querySelector('#rows .row.story[data-story]');
    plStRow.dispatchEvent(new window.MouseEvent('contextmenu', { bubbles: true, cancelable: true, clientX: 10, clientY: 10 }));
    ok([...doc.querySelectorAll('#popover .menu-list button[data-mi]')].some((b) => /Place at earliest slot/.test(b.textContent)),
      'the story context menu offers it too');
    doc.querySelector('#popover').hidden = true;
    // a locked feature keeps the entry, greyed out
    const lockId = doc.querySelector('#rows .row.item').dataset.id;
    window.HeadwayApp.ai.commit('lock it', (s) => { s.items.forEach((i) => { if (i.id === lockId) i.locked = true; }); });
    doc.querySelector('#rows .row.item[data-id="' + lockId + '"]')
      .dispatchEvent(new window.MouseEvent('contextmenu', { bubbles: true, cancelable: true, clientX: 10, clientY: 10 }));
    const lockedEntry = [...doc.querySelectorAll('#popover .menu-list button[data-mi]')].find((b) => /Place at earliest slot/.test(b.textContent));
    ok(!!lockedEntry && lockedEntry.disabled, 'a locked feature shows the entry disabled');
    doc.querySelector('#popover').hidden = true;
    undo();
    // clicking it reports back — placed, already there, or why it could not
    [...doc.querySelectorAll('#toasts .toast')].forEach((t) => t.remove());
    const beforePlace = JSON.stringify(window.HeadwayApp.ai.state());
    doc.querySelector('#rows .row.item[data-id="' + lockId + '"]')
      .dispatchEvent(new window.MouseEvent('contextmenu', { bubbles: true, cancelable: true, clientX: 10, clientY: 10 }));
    click([...doc.querySelectorAll('#popover .menu-list button[data-mi]')].find((b) => /Place at earliest slot/.test(b.textContent)));
    const placeToast = doc.querySelector('#toasts .toast');
    ok(!!placeToast && /earliest slot/.test(placeToast.textContent), 'clicking it reports back with a toast');
    if (JSON.stringify(window.HeadwayApp.ai.state()) !== beforePlace) undo();
    undo(); undo();
  }
  // the panel's own Snap earliest refuses a locked row the same way
  {
    window.HeadwayApp.ai.commit('cap on for snap', (s) => { s.meta.capacityEnabled = true; });
    const snapIt = state().items.find((i) => !i.milestone && i.startDay != null && !i.done);
    window.HeadwayApp.ai.commit('lock for snap', (s) => {
      s.items.forEach((i) => { if (i.id === snapIt.id) { i.locked = true; i.done = false; } });
    });
    window.__headway.selectItem(snapIt.id);
    const snapBtn = doc.querySelector('#panel [data-f="snap"]');
    ok(!!snapBtn, 'the panel offers Snap earliest for a scheduled feature');
    [...doc.querySelectorAll('#toasts .toast')].forEach((t) => t.remove());
    const snapBefore = state().items.find((i) => i.id === snapIt.id).startDay;
    click(snapBtn);
    const snapToast = doc.querySelector('#toasts .toast');
    ok(state().items.find((i) => i.id === snapIt.id).startDay === snapBefore,
      'the panel Snap earliest leaves a locked row where it is');
    ok(!!snapToast && /unlock it to place it/.test(snapToast.textContent),
      'and the panel guard says so in its own words');
    window.__headway.selectItem(null);
    undo(); undo();
  }
  // capacity chips in the Planning left pane
  {
    window.HeadwayApp.ai.commit('chips on', (s) => { s.meta.capacityEnabled = true; s.meta.planLevel = 'story'; s.meta.capMode = 'person'; });
    window.eval("document.querySelector('[data-menu=\"view\"]').click()");
    click([...doc.querySelectorAll('#popover .menu-list button')].find(b => /Expand all features/.test(b.textContent)));
    const stRow = doc.querySelector('#rows .row.story');
    ok(!!stRow && !!stRow.querySelector('.r-cap[data-act="st-cap"]'), 'story rows show a capacity type chip');
    ok(!!stRow.querySelector('.r-mult[data-act="st-mult"]'), 'and a multiplier chip in person mode');
    window.HeadwayApp.ai.commit('feature level', (s) => { s.meta.planLevel = 'feature'; });
    const itRow = [...doc.querySelectorAll('#rows .row.item[data-id]')]
      .find((r) => !(state().items.find((i) => i.id === r.dataset.id) || {}).milestone);
    ok(!!itRow && !!itRow.querySelector('.r-cap[data-act="cap"]'), 'feature rows show the chip at Features level');
    // a feature always plans as one of the types — no "general" entry
    click(itRow.querySelector('.r-cap[data-act="cap"]'));
    ok(menuBtns().length > 0 && !menuBtns().some((b) => /general/.test(b.textContent)),
      'the feature capacity picker offers the types only, with no “general” entry');
    window.dispatchEvent(new window.KeyboardEvent('keydown', { key: 'Escape', bubbles: true }));
    // the chip renders the whole type name; a narrow column clips it in CSS,
    // so widening the column reveals the rest instead of a baked-in ellipsis
    window.HeadwayApp.ai.commit('long type', (s) => {
      s.capTypes.push('Quality engineering');
      s.items.forEach((i) => { if (i.id === itRow.dataset.id) i.capType = 'Quality engineering'; });
    });
    ok(doc.querySelector('#rows .row.item[data-id="' + itRow.dataset.id + '"] .r-cap[data-act="cap"]').textContent === 'Quality engineering',
      'the capacity chip carries the full type name (no JS truncation)');
    undo();
    // a feature whose stories all agree shows the type as inherited
    const inhId = [...doc.querySelectorAll('#rows .row.item[data-id]')]
      .map((r) => state().items.find((i) => i.id === r.dataset.id))
      .find((i) => i && !i.milestone && (i.stories || []).length).id;
    window.HeadwayApp.ai.commit('stories agree', (s) => {
      s.items.forEach((i) => { if (i.id === inhId) { i.capType = ''; i.stories.forEach((st) => { st.capType = 'Design'; }); } });
    });
    const inhChip = doc.querySelector('#rows .row.item[data-id="' + inhId + '"] .r-cap[data-act="cap"]');
    ok(!!inhChip && inhChip.classList.contains('inherited') && /Design/.test(inhChip.textContent),
      'a feature whose stories all share a type shows it dimmed as inherited');
    undo();
    // story points mode: the Resources rows gain a points column
    window.HeadwayApp.ai.commit('points mode', (s) => { s.meta.capMode = 'points'; s.team[0].capType = 'Development'; });
    ok(!!doc.querySelector('#resGrid .rrow[data-mid] .res-pts[data-rpts]'), 'story-points mode gives each typed Resources row a points column');
    ok(!doc.querySelector('#resGrid .rrow[data-mid] [data-rcap]'), 'and hides the × seat chip — it is points OR the multiplier, never both');
    window.HeadwayApp.ai.commit('person mode', (s) => { s.meta.capMode = 'person'; });
    ok(!!doc.querySelector('#resGrid .rrow[data-mid] [data-rcap]') && !doc.querySelector('#resGrid .res-pts'),
      'per-person mode shows the × seat chip and no points column');
    // untyped people supply nothing: a "set type" prompt instead of the chips
    window.HeadwayApp.ai.commit('untyped person', (s) => {
      s.team[0].capType = ''; s.team[0].capacity = 2; s.team[0].points = 50;
      if (s.team[1]) s.team[1].capType = 'Development';
    });
    const utId = state().team[0].id;
    const utRow = () => doc.querySelector('#resGrid .rrow[data-mid="' + utId + '"]');
    const ph = utRow().querySelector('.res-cap.res-untyped');
    ok(!!ph && /set type/.test(ph.textContent) && ph.getAttribute('tabindex') === '0' && ph.getAttribute('role') === 'button',
      'an untyped person shows a quiet "set type" placeholder');
    ok(!utRow().querySelector('[data-rcap]') && !utRow().querySelector('[data-rpts]'), 'and no × seat or points chip');
    ok(state().team[0].capacity === 2 && state().team[0].points === 50, 'their seat and points are kept');
    window.HeadwayApp.ai.commit('points for untyped', (s) => { s.meta.capMode = 'points'; });
    ok(!!utRow().querySelector('.res-untyped') && !utRow().querySelector('[data-rpts]'), 'points mode: still the placeholder, no points chip');
    const esc = () => window.dispatchEvent(new window.KeyboardEvent('keydown', { key: 'Escape', bubbles: true }));
    const rowMenu = (row, re) => {
      row.dispatchEvent(new window.MouseEvent('contextmenu', { bubbles: true, cancelable: true, clientX: 120, clientY: 300 }));
      return menuBtns().find((b) => re.test(b.textContent));
    };
    // right-click → Capacity… reaches the points chip in points mode
    const t1Row = doc.querySelector('#resGrid .rrow[data-mid="' + state().team[1].id + '"]');
    click(rowMenu(t1Row, /Capacity/));
    const ptsInp = doc.querySelector('#resGrid .rrow[data-mid="' + state().team[1].id + '"] [data-rpts] input');
    ok(!!ptsInp, 'points mode: the row menu\'s Capacity… opens the points editor');
    if (ptsInp) ptsInp.dispatchEvent(new window.KeyboardEvent('keydown', { key: 'Escape', bubbles: true }));
    // … and the type picker for an untyped person
    click(rowMenu(utRow(), /Capacity/));
    ok(menuBtns().some((b) => /Development/.test(b.textContent)), 'for an untyped person Capacity… opens the type picker');
    esc();
    undo();
    // the placeholder answers the keyboard too
    utRow().querySelector('.res-untyped').dispatchEvent(new window.KeyboardEvent('keydown', { key: 'Enter', bubbles: true }));
    ok(menuBtns().some((b) => /Development/.test(b.textContent)), 'Enter on the placeholder opens the type picker');
    esc();
    utRow().querySelector('.res-untyped').dispatchEvent(new window.KeyboardEvent('keydown', { key: ' ', bubbles: true }));
    ok(menuBtns().some((b) => /Development/.test(b.textContent)), 'and so does Space');
    esc();
    click(utRow().querySelector('.res-untyped'));
    ok(menuBtns().some((b) => /Development/.test(b.textContent)), 'the placeholder opens the capacity-type picker');
    click(menuBtns().find((b) => /Development/.test(b.textContent)));
    ok(state().team[0].capType === 'Development' && !!utRow().querySelector('[data-rcap]') &&
      /^2×$/.test(utRow().querySelector('[data-rcap]').textContent.trim()), 'picking a type brings back the kept × seat');
    undo(); undo();
    undo();
    window.HeadwayApp.ai.commit('cap off', (s) => { s.meta.capacityEnabled = false; });
    ok(!doc.querySelector('#rows .r-cap'), 'no chips with capacity planning off');
    undo(); undo(); undo(); undo();
  }
  // a feature created from the UI starts on the first capacity type
  {
    window.HeadwayApp.ai.commit('cap default probe', (s) => {
      s.phases.push({ id: 'ph_capdflt_probe', name: 'Cap default probe', description: '', bucket: false, collapsed: false });
    });
    const capAddRow = doc.querySelector('#rows .row.addrow[data-phase="ph_capdflt_probe"]');
    click(capAddRow.querySelector('.row-left'));
    const made = state().items.find((i) => i.phaseId === 'ph_capdflt_probe');
    ok(!!made && made.capType === state().capTypes[0], 'a new feature carries the first capacity type');
    undo(); undo();
  }
  // standalone HTML: one file with the styles and scripts inlined and the document embedded
  const idx = fs.readFileSync(path.join(ROOT, 'index.html'), 'utf8');
  const files = { 'css/app.css': 'body{}', 'js/vendor/lucide.min.js': 'L', 'js/core.js': 'C', 'js/excel.js': 'X', 'js/export-png.js': 'P',
    'js/export-pptx.js': 'T', 'js/export-jira.js': 'J', 'js/jira.js': 'JJ', 'js/ai.js': 'A', 'js/app.js': 'var s = "</script>";' };
  const html = window.__headway.buildStandaloneHtml(idx, files, { meta: { title: 'My Plan' }, items: [] }, { view: 'sprints' });
  ok(html.indexOf('<link rel="stylesheet"') === -1 && html.indexOf('<style>body{}</style>') !== -1, 'the stylesheet is inlined');
  ok(!/<script src=/.test(html), 'no script tag points at a file');
  ok(html.indexOf('window.HEADWAY_VIEW = {"doc":{"meta":{"title":"My Plan"},"items":[]},"ui":{"view":"sprints"}}') !== -1, 'the document and view state ride along');
  ok(html.indexOf('<script>C</script>') !== -1 && html.indexOf('<script>A</script>') !== -1 && html.indexOf('exceljs') === -1 && html.indexOf('desktop') === -1,
    'app scripts are inlined; the desktop bridge and the Excel/PowerPoint engines stay out');
  ok(html.indexOf('var s = "<\\/script>";') !== -1, 'a closing script tag inside a source is escaped');
  ok(html.indexOf('<title>My Plan</title>') !== -1, 'the page is titled after the roadmap');
  ok(html.indexOf('data:image/svg+xml') !== -1 && html.indexOf('headway-theme-v1') !== -1, 'the favicon and the theme stamp survive');
  window.eval("document.querySelector('#btnExport').click()");
  ok(!!doc.querySelector('#modalHost #exFmtHtml') && !!doc.querySelector('#modalHost #exHtml'), 'the Export dialog offers Standalone HTML');
  click(doc.querySelector('#modalHost #exFmtHtml'));
  ok(doc.querySelector('#modalHost #exTimeline').hidden && !doc.querySelector('#modalHost #exHtml').hidden, 'picking it hides the timeline options');
  click(doc.querySelector('#modalHost [data-m=x]'));
}

// ---------------------------------------------------------------- story assignees in the story panel
{
  const hostA = state().items.find(i => !i.milestone && i.stories && i.stories.length &&
    doc.querySelector('#rows .row.item[data-id="' + i.id + '"] .r-num'));
  click(doc.querySelector('#rows .row.item[data-id="' + hostA.id + '"] .r-num'));
  click(doc.querySelector('#panel [data-pst-edit="' + hostA.stories[0].id + '"]'));
  const stSecsA = Array.from(doc.querySelectorAll('#panel .p-sec')).map(e => e.dataset.sec);
  ok(stSecsA.indexOf('st-people') !== -1 && stSecsA.indexOf('st-people') < stSecsA.indexOf('st-integrations'),
    'story panel has a People section (' + stSecsA.join(',') + ')');
  ok(!doc.querySelector('#panel .p-sec[data-sec="st-people"] [data-dd="teamType"]'), 'story People carries no role picker');
  const asgBtn = doc.querySelector('#panel [data-dd="stassign"]');
  ok(!!asgBtn, 'story panel offers Assign…');
  click(asgBtn);
  const firstPerson = doc.querySelector('#popover .menu-list button');
  ok(!!firstPerson, 'assign dropdown lists the roster');
  click(firstPerson);
  const stA = () => state().items.find(i => i.id === hostA.id).stories[0];
  ok((stA().assignees || []).length === 1, 'picking a person assigns them to the story');
  const rmBtn = doc.querySelector('#panel .p-sec[data-sec="st-people"] [data-stasgrm]');
  ok(!!rmBtn, 'assigned person shows as a chip with a remove button');
  click(rmBtn);
  ok((stA().assignees || []).length === 0, 'the chip\'s x unassigns');
  click(doc.querySelector('#panel [data-stf="up"]'));
}

// ---------------------------------------------------------------- pane collapse: sticky header, left toggle, [ ] hotkeys
{
  click(doc.querySelector('#viewTabs [data-view="planning"]'));
  const css2 = fs.readFileSync(path.join(ROOT, 'css/app.css'), 'utf8');
  ok(/\.p-top\s*\{[^}]*position:\s*sticky/.test(css2), 'panel header row is sticky');
  ok(!/#panelPeek\s*\{[^}]*top:\s*50%/.test(css2), 'panel peek button no longer floats mid-screen');
  const key = (k, target) => (target || doc.body).dispatchEvent(new window.KeyboardEvent('keydown', { key: k, bubbles: true, cancelable: true }));
  // right pane: ] toggles, peek button sits where the close button was
  click(doc.querySelector('#rows .row.item .r-num'));
  ok(!doc.querySelector('#panel').hidden, 'panel open before the hotkey');
  key(']');
  ok(doc.querySelector('#panel').hidden && !doc.querySelector('#panelPeek').hidden, '] collapses the right panel');
  key(']');
  ok(!doc.querySelector('#panel').hidden, '] again reopens it');
  // left pane: a collapse button in the header corner, [ toggles
  const lc = doc.querySelector('#leftCollapse');
  ok(!!lc && !!lc.closest('.hdr-left'), 'left pane has a collapse button in its header');
  ok(doc.querySelector('#leftPeek').hidden, 'left peek button is hidden while the pane is open');
  click(lc);
  ok(doc.body.classList.contains('left-collapsed'), 'clicking it collapses the left pane');
  ok(parseInt(doc.documentElement.style.getPropertyValue('--left-w'), 10) === 0, 'collapsed left pane takes no width');
  ok(/body\.left-collapsed[^{]*\.hdr-left[^{]*\{[^}]*display:\s*none/.test(css2) && /body\.left-collapsed[^{]*\.row-left[^{]*\{[^}]*display:\s*none/.test(css2),
    'collapsed left pane hides its header and row cells entirely');
  ok(!doc.querySelector('#leftPeek').hidden && doc.querySelector('#leftPeek').parentElement === doc.querySelector('#panelPeek').parentElement,
    'only a floating button remains, anchored like the right peek button');
  ok(/#leftPeek\s*\{[^}]*left:\s*12px/.test(css2) && /#leftPeek\s*\{[^}]*top:\s*12px/.test(css2), 'left peek sits in the top-left corner');
  click(doc.querySelector('#leftPeek'));
  ok(!doc.body.classList.contains('left-collapsed') && parseInt(doc.documentElement.style.getPropertyValue('--left-w'), 10) > 200,
    'clicking again restores the left pane');
  key('[');
  ok(doc.body.classList.contains('left-collapsed'), '[ collapses the left pane');
  key('[');
  ok(!doc.body.classList.contains('left-collapsed'), '[ again restores it');
  // never while typing
  const rf = doc.querySelector('#rowFilter');
  rf.focus();
  key(']', rf);
  key('[', rf);
  ok(!doc.querySelector('#panel').hidden && !doc.body.classList.contains('left-collapsed'), 'brackets are ignored inside a text box');
  rf.blur();
  const ce = doc.querySelector('#panel .wz-ed');
  if (ce) { ce.focus(); key(']', ce); ok(!doc.querySelector('#panel').hidden, 'brackets are ignored inside a rich editor'); ce.blur(); }
}

// ------------------------------------------- assistant close clears the panel peek toggle
// #panelPeek floats over the drawer's top-right corner, so #aiClose would sit
// under it. jsdom has no layout engine — the two boxes both measure 0 — so this
// asserts the state class the fix keys on plus the CSS rule that reads it.
{
  click(doc.querySelector('#viewTabs [data-view="planning"]'));
  const cssPk = fs.readFileSync(path.join(ROOT, 'css/app.css'), 'utf8');
  const bracket = () => doc.body.dispatchEvent(new window.KeyboardEvent('keydown', { key: ']', bubbles: true, cancelable: true }));
  ok(!doc.body.classList.contains('peek-on'), 'no peek-on class while the right panel is open');
  bracket();
  ok(!doc.querySelector('#panelPeek').hidden && doc.body.classList.contains('peek-on'),
    'collapsing the right panel marks the body peek-on');
  click(doc.querySelector('#btnAI'));
  ok(!doc.querySelector('#aiDrawer').hidden && doc.body.classList.contains('peek-on'),
    'opening the assistant leaves the peek toggle (and the class) in place');
  const mPk = cssPk.match(/body\.peek-on\s+\.ai-head\s*\{([^}]*)\}/);
  ok(!!mPk && /padding-right:\s*(4[0-9]|[5-9]\d)px/.test(mPk[1]),
    'body.peek-on .ai-head reserves 40px+ on the right so #aiClose lands left of #panelPeek');
  click(doc.querySelector('#aiDrawer #aiClose'));
  if (doc.activeElement && doc.activeElement.blur) doc.activeElement.blur(); // the drawer focused its composer
  bracket();
  ok(!doc.body.classList.contains('peek-on'), 'reopening the right panel drops the class again');
}
// ---------------------------------------------------------------- milestone styles (UI)
{
  const schedRow = Array.from(doc.querySelectorAll('#rows .row.item')).find(r => {
    const i = state().items.find(x => x.id === r.dataset.id);
    return i && i.startDay != null && !i.milestone && !i.locked;
  });
  schedRow.dispatchEvent(new window.MouseEvent('contextmenu', { bubbles: true, clientX: 240, clientY: 240 }));
  click(Array.from(doc.querySelectorAll('#popover .menu-list button')).find(b => /Convert to milestone/.test(b.textContent)));
  const msIt = state().items.find(i => i.id === schedRow.dataset.id);
  const msRow = doc.querySelector('#rows .row.item[data-id="' + msIt.id + '"]');
  ok(!!msRow && msIt.milestone && msIt.startDay != null, 'a scheduled milestone row is visible');
  ok(msIt.size == null && msIt.priority == null, 'converting to a milestone clears size and priority');
  ok(!msRow.querySelector('[data-act="size"]') && !msRow.querySelector('[data-act="priority"]'), 'milestone rows have no size or priority chips');
  ok(!!doc.querySelector('#rows .bar.ms.ms-diamond[data-bar="' + msIt.id + '"]'), 'milestones default to the diamond');
  ok(!!doc.querySelector('#rows .row.item.ms[data-id="' + msIt.id + '"] .r-name'), 'a milestone row carries the ms class (bold title)');
  msRow.dispatchEvent(new window.MouseEvent('contextmenu', { bubbles: true, clientX: 240, clientY: 240 }));
  const starBtn = Array.from(doc.querySelectorAll('#popover .menu-list button')).find(b => /Star/.test(b.textContent));
  ok(!!starBtn, 'milestone context menu offers the Star style');
  click(starBtn);
  ok(state().items.find(i => i.id === msIt.id).msStyle === 'star', 'picking Star stores the style');
  ok(!!doc.querySelector('#rows .bar.ms.ms-star[data-bar="' + msIt.id + '"]'), 'timeline marks the star milestone');
  ok(!!doc.querySelector('#rows .row.item[data-id="' + msIt.id + '"] .r-dot.msdot.star'), 'row dot takes the star style');
  click(msRow.querySelector('.r-num'));
  const chip = doc.querySelector('#panel .p-mschip');
  ok(!!chip && /Star/.test(chip.textContent), 'panel milestone chip names the style');
  ok(!doc.querySelector('#panel [data-f="priSet"]') && !doc.querySelector('#panel [data-f="size"]'), 'milestone panel shows neither size nor priority');
  click(chip);
  const circBtn = Array.from(doc.querySelectorAll('#popover .menu-list button')).find(b => /Circle/.test(b.textContent));
  ok(!!circBtn, 'chip opens the style picker');
  click(circBtn);
  ok(state().items.find(i => i.id === msIt.id).msStyle === 'circle', 'picker sets Circle');
  ok(!!doc.querySelector('#rows .bar.ms.ms-circle[data-bar="' + msIt.id + '"]'), 'timeline marks the circle milestone');
  const png = window.RM_EXPORT.layout(state(), {});
  const lr = png.rows.find(r => r.kind === 'item' && r.id === msIt.id);
  ok(!!lr && lr.bar.msStyle === 'circle', 'PNG layout carries the milestone style');
  window.dispatchEvent(new window.KeyboardEvent('keydown', { key: 'z', metaKey: true, bubbles: true }));
  window.dispatchEvent(new window.KeyboardEvent('keydown', { key: 'z', metaKey: true, bubbles: true }));
  ok(state().items.find(i => i.id === msIt.id).msStyle == null, 'undo restores the diamond');
  window.dispatchEvent(new window.KeyboardEvent('keydown', { key: 'z', metaKey: true, bubbles: true }));
  ok(state().items.find(i => i.id === msIt.id).milestone === false, 'undo restores the feature');
}

// ---------------------------------------------------------------- story scales, snap prefs, holidays, roles
{
  // stories estimate on their own scale: story points + severity ladder by default
  ok(state().meta.storySizeScheme === 'fibonacci' && state().meta.storyPriorityScheme === 'levels',
    'stories default to story points and the severity ladder');
  ok(window.RM.sizeOrderOf(state(), 'story').join(',') === '0,0.5,1,2,3,5,8,13' &&
    window.RM.sizeOrderOf(state()).join(',') !== window.RM.sizeOrderOf(state(), 'story').join(','),
    'the story scale is separate from the feature scale');
  click(doc.querySelector('#btnSetup'));
  suTab('est');
  ok(schemeOpts('size', 'story').indexOf('tshirt') !== -1, 'Setup → Sizing offers a story scale picker');
  pickScheme('size', 'story', 'tshirt');
  ok(state().meta.storySizeScheme === 'tshirt' && state().meta.sizeScheme !== 'none' &&
    window.RM.sizeOrderOf(state(), 'story').indexOf('XL') !== -1, 'picking a story scale leaves the feature scale alone');
  pickScheme('prio', 'story', 'moscow');
  ok(state().meta.storyPriorityScheme === 'moscow' && state().meta.priorityScheme !== 'moscow',
    'story priority scheme is independent of the feature scheme');
  window.dispatchEvent(new window.KeyboardEvent('keydown', { key: 'z', metaKey: true, bubbles: true }));
  window.dispatchEvent(new window.KeyboardEvent('keydown', { key: 'z', metaKey: true, bubbles: true }));

  // story chips on every board: Planning rows, Sprinting rows, Prioritizing cards
  click(doc.querySelector('#viewTabs [data-view="planning"]'));
  click(doc.querySelector('#detailBtn'));
  click(doc.querySelector('#popover .menu-list [data-mi="1"]')); // Story detail
  const stRowP = doc.querySelector('#rows .row.story[data-story]');
  ok(!!stRowP && !!stRowP.querySelector('[data-act="st-size"]') && !!stRowP.querySelector('[data-act="st-wk"]'),
    'planning story rows carry size + duration chips');
  click(stRowP.querySelector('[data-act="st-size"]'));
  const pt5 = Array.from(doc.querySelectorAll('#popover .menu-list button')).find(b => /^5\b/.test(b.textContent.trim()));
  ok(!!pt5, 'the story size menu lists story points');
  click(pt5);
  const hostP = state().items.find(i => i.id === stRowP.dataset.id);
  ok(hostP.stories.find(x => x.id === stRowP.dataset.story).size === '5', 'picking a story size commits');
  window.dispatchEvent(new window.KeyboardEvent('keydown', { key: 'z', metaKey: true, bubbles: true }));
  // an unscheduled story shows its name faintly where the feature starts
  {
    const ghostRow = Array.from(doc.querySelectorAll('#rows .row.story[data-story]')).find(r => {
      const h = state().items.find(i => i.id === r.dataset.id);
      const st = h && h.stories.find(x => x.id === r.dataset.story);
      return h && h.startDay != null && st && st.startDay == null;
    });
    ok(!ghostRow || !!ghostRow.querySelector('.st-ghost'), 'unscheduled stories show a faint name on the lane');
  }
  // selecting a scheduled story outlines the story bar, not the feature bar
  {
    const schedRow = Array.from(doc.querySelectorAll('#rows .row.story[data-story]')).find(r => r.querySelector('[data-stbar]'));
    if (schedRow) {
      click(schedRow.querySelector('[data-act="st-open"]'));
      ok(!!doc.querySelector('#rows .st-bar.selected') && !doc.querySelector('#rows .row.item[data-id="' + schedRow.dataset.id + '"] .bar.selected'),
        'a selected story owns the selection outline');
    }
  }
  click(doc.querySelector('#detailBtn'));
  click(doc.querySelector('#popover .menu-list [data-mi="0"]')); // back to Feature detail
  click(doc.querySelector('#viewTabs [data-view="sprints"]'));
  ok(!!doc.querySelector('#sprintsView .spv-row [data-spact="asg"]') || !!doc.querySelector('.spv-row [data-spact="asg"]'),
    'sprinting rows carry an assignee chip');
  click(doc.querySelector('#viewTabs [data-view="planning"]'));

  // one Add feature row per group when grouped
  {
    const groupedBefore = doc.querySelectorAll('#rows .row.addrow.sub').length;
    if (groupedBefore) {
      const phasesWithGroups = new Set(Array.from(doc.querySelectorAll('#rows .row.addrow.sub')).map(r => r.dataset.phase));
      const dup = Array.from(doc.querySelectorAll('#rows .row.addrow:not(.sub)')).filter(r => phasesWithGroups.has(r.dataset.phase));
      ok(dup.length === 0, 'grouped phases have no duplicate phase-level Add feature row');
    }
  }

  // snap: features and stories each keep their own grid
  click(doc.querySelector('#btnSetup'));
  suTab('prefs');
  ok(doc.querySelectorAll('#setupView [data-pref-snap="feature:week"].on').length === 1 &&
    doc.querySelectorAll('#setupView [data-pref-snap="story:sprint"].on').length === 1,
    'defaults: features snap to the week, stories to the sprint');
  click(doc.querySelector('#setupView [data-pref-snap="story:day"]'));
  ok(JSON.parse(window.localStorage.getItem('headway-ui-v1')).snapStory === 'day' &&
    JSON.parse(window.localStorage.getItem('headway-ui-v1')).snapFeat === 'week', 'story snap persists on its own');
  click(doc.querySelector('#setupView [data-pref-snap="story:sprint"]'));

  // holidays edit in place
  suTab('project');
  {
    const nm = doc.querySelector('#setupView [data-suholname="0"]');
    ok(!!nm && !!doc.querySelector('#setupView [data-suholstart="0"]'), 'holiday rows expose name + date inputs');
    nm.value = 'Renamed day';
    nm.dispatchEvent(new window.Event('change', { bubbles: true }));
    ok(state().meta.holidayRanges.some(r => r.name === 'Renamed day'), 'renaming a holiday commits');
    const idx = state().meta.holidayRanges.findIndex(r => r.name === 'Renamed day');
    const r0 = state().meta.holidayRanges[idx];
    const endIn = doc.querySelector('#setupView [data-suholend="' + idx + '"]');
    const nextDay = (iso) => { const d = new Date(iso + 'T00:00:00Z'); d.setUTCDate(d.getUTCDate() + 1); return d.toISOString().slice(0, 10); };
    const newEnd = nextDay(r0.end);
    endIn.value = newEnd;
    endIn.dispatchEvent(new window.Event('change', { bubbles: true }));
    const r1 = state().meta.holidayRanges.find(r => r.name === 'Renamed day');
    ok(r1.end === newEnd && state().meta.holidays.indexOf(newEnd) !== -1, 'moving a holiday\'s last day extends its dates');
    window.dispatchEvent(new window.KeyboardEvent('keydown', { key: 'z', metaKey: true, bubbles: true }));
    window.dispatchEvent(new window.KeyboardEvent('keydown', { key: 'z', metaKey: true, bubbles: true }));
  }

  // roles delete even when in use — people and features just lose the role
  {
    // give a feature a role through the panel first
    click(doc.querySelector('#viewTabs [data-view="planning"]'));
    const roleRow = doc.querySelector('#rows .row.item');
    click(roleRow.querySelector('.r-num'));
    click(doc.querySelector('#panel [data-dd="teamType"]'));
    const roleOpt = Array.from(doc.querySelectorAll('#popover .menu-list button')).find(b => b.textContent.trim() === state().teamTypes[0]);
    click(roleOpt);
    const usedRole = state().teamTypes[0];
    ok(state().items.find(i => i.id === roleRow.dataset.id).teamType === usedRole, 'a feature now carries a role');
    click(doc.querySelector('#btnSetup'));
    suTab('budget');
    click(doc.querySelector('#setupView [data-suttrm="' + usedRole + '"]'));
    ok(state().teamTypes.indexOf(usedRole) === -1, 'an in-use role can be removed');
    ok(!state().team.some(m => m.type === usedRole) && !state().items.some(i => i.teamType === usedRole),
      'people and features that had it now have no role');
    window.dispatchEvent(new window.KeyboardEvent('keydown', { key: 'z', metaKey: true, bubbles: true }));
    ok(state().teamTypes.indexOf(usedRole) !== -1, 'undo restores the role');
    window.dispatchEvent(new window.KeyboardEvent('keydown', { key: 'z', metaKey: true, bubbles: true }));
  }
  click(doc.querySelector('#viewTabs [data-view="planning"]'));

  // type pickers, row icons, epic type
  {
    // pick a currently-rendered row's item — not state().items[0], which
    // may sit in a phase collapsed by default in the fixture
    const firstRow = doc.querySelector('#rows .row.item[data-id]');
    const first = state().items.find(i => i.id === firstRow.dataset.id);
    window.__headway.selectItem ? window.__headway.selectItem(first.id) : click(firstRow);
    const chip = doc.querySelector('#panel [data-act="itype"]');
    ok(chip && /Feature/.test(chip.textContent), 'panel shows the type chip');
    click(chip);
    const bug = [...doc.querySelectorAll('.menu-list [data-mi]')].find(el => /^Bug$/.test(el.textContent.trim()));
    ok(!!bug, 'type dropdown lists Bug');
    click(bug);
    ok(state().items.find(i => i.id === first.id).type === 'bug', 'picking Bug sets the item type');
    ok(!!doc.querySelector('#rows .row.item[data-id="' + first.id + '"] .r-type'), 'a non-default type shows its icon on the row');
    ok(!doc.querySelector('#rows .row.item[data-id="' + first.id + '"] .r-dot'), 'the type icon replaces the colored square');
    window.__headway.setItemType(first.id, 'feature');
    ok(!doc.querySelector('#rows .row.item[data-id="' + first.id + '"] .r-type') &&
      !!doc.querySelector('#rows .row.item[data-id="' + first.id + '"] .r-dot'), 'the default Feature type draws the filled square');

    // a type label containing markup renders as text, not HTML, in the
    // panel dropdown (typeMenuItems must esc() the label)
    window.HeadwayApp.ai.commit('rename type', (s) => {
      const bugType = s.meta.itemTypes.find(t => t.key === 'bug');
      bugType.label = '<b>Bug</b>';
    });
    const chip2 = doc.querySelector('#panel [data-act="itype"]');
    click(chip2);
    const bugAfterRename = [...doc.querySelectorAll('.menu-list [data-mi]')].find(el => el.textContent.trim().indexOf('<b>Bug</b>') !== -1);
    ok(!!bugAfterRename && bugAfterRename.innerHTML.indexOf('&lt;b&gt;') !== -1,
      'a type label with markup renders as escaped text in the dropdown');
    window.HeadwayApp.ai.commit('rename type', (s) => {
      const bugType = s.meta.itemTypes.find(t => t.key === 'bug');
      bugType.label = 'Bug';
    });

    // a type icon containing attribute-breaking markup renders as a single
    // escaped data-lucide attribute, not injected markup (menu renderers
    // must esc() m.icon)
    window.HeadwayApp.ai.commit('rename type icon', (s) => {
      const bugType = s.meta.itemTypes.find(t => t.key === 'bug');
      bugType.icon = 'tag" data-x="y';
    });
    const chip3 = doc.querySelector('#panel [data-act="itype"]');
    click(chip3);
    const bugIconRow = [...doc.querySelectorAll('.menu-list [data-mi]')].find(el => /^Bug$/.test(el.textContent.trim()));
    ok(!!bugIconRow, 'type dropdown still lists Bug after icon change');
    const bugIcon = bugIconRow.querySelector('i');
    ok(!!bugIcon && bugIcon.getAttribute('data-lucide') === 'tag" data-x="y',
      'the icon renders with the full unbroken-out string as data-lucide');
    ok(!bugIcon.hasAttribute('data-x'), 'no stray data-x attribute was injected from the icon value');
    window.HeadwayApp.ai.commit('rename type icon', (s) => {
      const bugType = s.meta.itemTypes.find(t => t.key === 'bug');
      bugType.icon = 'bug';
    });
  }

  // type glyphs everywhere + color by item type
  {
    const firstRow = doc.querySelector('#rows .row.item[data-id]');
    const first = state().items.find(i => i.id === firstRow.dataset.id);
    window.__headway.setItemType(first.id, 'bug');
    const glyph = doc.querySelector('#rows .row.item[data-id="' + first.id + '"] .r-type');
    ok(glyph && glyph.querySelector('[data-lucide="bug"]') && /color:\s*#/.test(glyph.getAttribute('style') || ''),
      'a Bug row draws the bug icon in a color');
    ok(/#/.test(glyph.getAttribute('style')) && glyph.getAttribute('style').toUpperCase().indexOf(window.RM.colorForItem(state(), first).toUpperCase()) !== -1,
      'the glyph wears the item color of the active color mode');
    // a story row draws the bookmark, filled
    window.HeadwayApp.ai.commit('story', (s) => { const it = window.RM.itemById(s, first.id); it.stories = it.stories || []; it.stories.unshift({ id: 'glyph-st', title: 'Glyph story' }); });
    window.HeadwayApp.ai.setView('planning');
    click(doc.querySelector('#rows .row.item[data-id="' + first.id + '"] [data-act="stories"]'));
    const stGlyph = doc.querySelector('#rows .row.story[data-story="glyph-st"] .r-type');
    ok(stGlyph && stGlyph.classList.contains('fill') && stGlyph.querySelector('[data-lucide="bookmark"]'), 'a story row draws the filled bookmark glyph');
    // prioritizing cards carry the glyph in a head row beside the title
    click(doc.querySelector('#viewTabs [data-view="prio"]'));
    const card = doc.querySelector('#prioView .pr-card[data-prcard="' + first.id + '"]');
    ok(card && card.querySelector('.pr-head .r-type [data-lucide="bug"]') && card.querySelector('.pr-head .pr-title'),
      'a prioritizing card shows the type glyph beside its title');
    // sprinting rows and story-view headings
    click(doc.querySelector('#viewTabs [data-view="sprints"]'));
    ok(!!doc.querySelector('#sprintView .spv-row[data-spid="' + first.id + '"] .r-type [data-lucide="bug"]'), 'a sprinting row shows the type glyph');
    ok(!doc.querySelector('#sprintView .spv-row .r-dot:not(.msdot) + .r-type'), 'sprinting rows no longer stack a square before the glyph');
    window.__headway.setItemType(first.id, 'feature');
    ok(!!doc.querySelector('#sprintView .spv-row[data-spid="' + first.id + '"] .r-dot'), 'a Feature sprinting row draws the square');
    // View menu offers Color by item type; picking it recolors the glyph by type
    click(doc.querySelector('#viewTabs [data-view="planning"]'));
    click(doc.querySelector('.menu-btn[data-menu="view"]'));
    const cbt = [...doc.querySelectorAll('#popover .menu-list button')].find(b => /Color by item type/.test(b.textContent));
    ok(!!cbt, 'View menu offers Color by item type');
    click(cbt);
    ok(window.RM.colorMode() === 'type', 'picking it switches the color mode');
    const sq = doc.querySelector('#rows .row.item[data-id="' + first.id + '"] .r-dot');
    ok(sq && sq.getAttribute('style').toUpperCase().indexOf(window.RM.colorForType(state(), 'feature')) !== -1, 'the square takes the Feature type color');
    // prefs segment carries the mode too
    click(doc.querySelector('#btnSetup'));
    click(doc.querySelector('#setupView [data-sutab="prefs"]'));
    ok(!!doc.querySelector('#setupView [data-pref-color="type"]'), 'Settings offers Color bars by item type');
    click(doc.querySelector('#setupView [data-pref-color="workstream"]'));
    ok(window.RM.colorMode() === 'workstream', 'the prefs segment switches back');
    // Hierarchy card: a color input per type commits the type color
    click(doc.querySelector('#btnSetup'));
    click(doc.querySelector('#setupView [data-sutab="org"]'));
    const cin = doc.querySelector('#setupView input[type="color"][data-suhtcolor="bug"]');
    ok(!!cin, 'Hierarchy types table has a color input per type');
    cin.value = '#112233';
    cin.dispatchEvent(new window.Event('change', { bubbles: true }));
    ok(state().meta.itemTypes.find(t => t.key === 'bug').color === '112233', 'changing the color input sets the type color');
    window.HeadwayApp.ai.setView('planning');
  }

  // milestones stay off the prioritizing board; Jira issue-type icons are 14px
  {
    const feat = state().items.find(i => !i.milestone && i.stories && i.stories.length === 0) || state().items.find(i => !i.milestone);
    window.HeadwayApp.ai.commit('milestone', (s) => { const it = window.RM.itemById(s, feat.id); it.milestone = true; it.durDays = it.startDay == null ? null : 1; });
    click(doc.querySelector('#viewTabs [data-view="prio"]'));
    ok(!doc.querySelector('#prioView [data-prcard="' + feat.id + '"]'), 'a milestone has no card on the prioritizing board');
    window.HeadwayApp.ai.commit('milestone', (s) => { window.RM.itemById(s, feat.id).milestone = false; });
    ok(!!doc.querySelector('#prioView [data-prcard="' + feat.id + '"]'), 'clearing the flag brings the card back');
    const css = fs.readFileSync(path.join(ROOT, 'css/app.css'), 'utf8');
    ok(/\.jr-types td svg\.lucide\s*\{[^}]*width:\s*14px[^}]*height:\s*14px/.test(css), 'Jira issue-type table icons are sized like every other icon');
  }

  // sprinting: one sprint per row, carry-over glyph, sprint-number tag, trimmed ends
  {
    click(doc.querySelector('#viewTabs [data-view="sprints"]'));
    const meta = state().meta;
    const nums = [...doc.querySelectorAll('#sprintView .spv-sec')].map(s => s.dataset.spsec).filter(k => k !== 'u').map(Number);
    const startNum = nums[0];
    const it = state().items.find(i => !i.milestone);
    const perSprint = window.RM.sprintDays(meta);
    window.HeadwayApp.ai.commit('span', (s) => {
      const t = window.RM.itemById(s, it.id);
      t.startDay = window.RM.sprintStartDay(s.meta, startNum); t.durDays = perSprint * 3; t.riskDays = 0; t.locked = false;
    });
    const rows = doc.querySelectorAll('#sprintView .spv-row[data-spid="' + it.id + '"]');
    ok(rows.length === 1 && rows[0].dataset.spsec === String(startNum), 'a three-sprint item lists once, under the sprint it starts in');
    const info = rows[0].querySelector('.spv-info');
    ok(info && info.getAttribute('title') === 'Expecting to carryover for 2 sprints (through sprint ' + (startNum + 2) + ')',
      'the carry-over glyph says how many sprints it runs on and the last one');
    ok(!rows[0].querySelector('.spv-tag') && !rows[0].querySelector('.r-num'),
      'rows carry no sprint-number bubble and no #number');
    ok(!rows[0].querySelector('.spv-dates') && !rows[0].querySelector('.spv-phase'), 'rows carry no date range or phase column');
    window.HeadwayApp.ai.commit('span', (s) => { const t = window.RM.itemById(s, it.id); t.durDays = perSprint; });
    ok(!doc.querySelector('#sprintView .spv-row[data-spid="' + it.id + '"] .spv-info'), 'an item inside one sprint has no carry-over glyph');
    // empty sprints at either end are hidden (the current sprint always stays)
    const secs = [...doc.querySelectorAll('#sprintView .spv-sec')].filter(s => s.dataset.spsec !== 'u');
    const sideKeys = [...doc.querySelectorAll('#sprintView .spv-sbtn')].map(b => b.dataset.spside);
    const nonEmpty = s => s.querySelectorAll('.spv-row').length > 0 || doc.querySelector('.spv-sbtn.today[data-spside="' + s.dataset.spsec + '"]');
    ok(nonEmpty(secs[0]) && nonEmpty(secs[secs.length - 1]), 'the first and last sprint sections shown have rows (or are today)');
    ok(sideKeys.length === secs.length + 1 && sideKeys[sideKeys.length - 1] === 'u', 'the sidebar mirrors the trimmed sections plus Unscheduled');
    // push one item far out: the tail extends only to that sprint
    const lastShown = Number(secs[secs.length - 1].dataset.spsec);
    const allNums = [];
    for (let w = 0; w < meta.numWeeks; w++) { const n = window.RM.sprintNumForWeek(meta, w); if (allNums.indexOf(n) === -1) allNums.push(n); }
    const far = allNums[allNums.length - 1];
    if (far > lastShown) {
      window.HeadwayApp.ai.commit('span', (s) => { const t = window.RM.itemById(s, it.id); t.startDay = window.RM.sprintStartDay(s.meta, far); t.durDays = perSprint; });
      const secs2 = [...doc.querySelectorAll('#sprintView .spv-sec')].filter(s => s.dataset.spsec !== 'u');
      ok(secs2[secs2.length - 1].dataset.spsec === String(far), 'scheduling into the last sprint reveals it');
      window.HeadwayApp.ai.commit('span', (s) => { const t = window.RM.itemById(s, it.id); t.startDay = window.RM.sprintStartDay(s.meta, startNum); });
    } else ok(true, 'timeline already ends at the last shown sprint');
  }

  // sprinting chips: assignees replace duration, workstream chip, heading chips
  {
    click(doc.querySelector('#viewTabs [data-view="sprints"]'));
    const row = doc.querySelector('#sprintView .spv-sec:not([data-spsec="u"]) .spv-row');
    const id = row.dataset.spid;
    ok(!!row.querySelector('[data-spact="asg"]') && !row.querySelector('[data-spact="dur"]'), 'feature rows show an assignee chip instead of duration');
    ok(!!row.querySelector('[data-spact="ws"]'), 'feature rows carry a workstream chip');
    click(row.querySelector('[data-spact="asg"]'));
    const member = state().team[0];
    const mBtn = [...doc.querySelectorAll('#popover .menu-list button')].find(b => b.textContent.indexOf(window.RM.memberLabel(member)) !== -1);
    ok(!!mBtn, 'the assignee chip opens the roster');
    click(mBtn);
    ok((state().items.find(i => i.id === id).assignees || []).indexOf(member.id) !== -1, 'picking a person assigns them');
    ok(!!doc.querySelector('#sprintView .spv-row[data-spid="' + id + '"] [data-spact="asg"] .avatar'), 'the chip shows their avatar');
    const wsChip = doc.querySelector('#sprintView .spv-row[data-spid="' + id + '"] [data-spact="ws"]');
    click(wsChip);
    const wsBtn = [...doc.querySelectorAll('#popover .menu-list button')].filter(b => !/default|Other/.test(b.textContent))[0];
    const wsName = wsBtn.textContent.trim();
    click(wsBtn);
    ok(state().items.find(i => i.id === id).workstream === wsName, 'the workstream chip sets the workstream');
    // story level: heading chips and story assignee chip
    // (priority is 'none' at this point in the suite from an earlier undo — turn
    // it on so the heading's priority chip has something to render)
    const prevPriScheme = state().meta.priorityScheme;
    window.HeadwayApp.ai.commit('priority scheme', (s) => { s.meta.priorityScheme = 'moscow'; });
    click(doc.querySelector('#sprintView [data-spdd="level"]'));
    click([...doc.querySelectorAll('#popover .menu-list button')].find(b => /Story/.test(b.textContent)));
    const head = doc.querySelector('#sprintView .spv-feat');
    ok(head && head.querySelector('.spv-est [data-spact="priority"]') && head.querySelector('.spv-est [data-spact="size"]') && head.querySelector('.spv-est [data-spact="dur"]'),
      'story-view feature headings show priority, size and duration chips');
    ok(!head.querySelector('.spv-phase'), 'headings drop the phase column');
    window.HeadwayApp.ai.commit('priority scheme', (s) => { s.meta.priorityScheme = prevPriScheme; });
    const stRow = doc.querySelector('#sprintView .spv-row.spv-st');
    ok(stRow && stRow.querySelector('[data-spact="st-asg"]') && !stRow.querySelector('[data-spact="st-wk"]'), 'story rows show a story assignee chip instead of duration');
    click(stRow.querySelector('[data-spact="st-asg"]'));
    const mBtn2 = [...doc.querySelectorAll('#popover .menu-list button')].find(b => b.textContent.indexOf(window.RM.memberLabel(member)) !== -1);
    click(mBtn2);
    const stIt = state().items.find(i => i.id === stRow.dataset.spid);
    ok((stIt.stories.find(s => s.id === stRow.dataset.spst).assignees || []).indexOf(member.id) !== -1, 'picking a person assigns them to the story');
    click(doc.querySelector('#sprintView [data-spdd="level"]'));
    click([...doc.querySelectorAll('#popover .menu-list button')].find(b => /Feature/.test(b.textContent)));
  }

  // sprinting context menu: Move to sprint for features and stories, Assign
  {
    click(doc.querySelector('#viewTabs [data-view="sprints"]'));
    const menu = () => [...doc.querySelectorAll('#popover .menu-list button')];
    const row = doc.querySelector('#sprintView .spv-sec:not([data-spsec="u"]) .spv-row');
    const id = row.dataset.spid;
    const secNums = [...doc.querySelectorAll('#sprintView .spv-sec')].map(s => s.dataset.spsec).filter(k => k !== 'u').map(Number);
    row.dispatchEvent(new window.MouseEvent('contextmenu', { bubbles: true, cancelable: true, clientX: 120, clientY: 120 }));
    const mv = menu().find(b => /Move to sprint/.test(b.textContent));
    ok(!!mv && menu().some(b => /Assign/.test(b.textContent)), 'the row menu offers Move to sprint and Assign');
    click(mv);
    const target = secNums.find(n => n !== Number(row.dataset.spsec)) || secNums[0] + 1;
    const tBtn = menu().find(b => new RegExp('^Sprint ' + target + '\\b').test(b.textContent.trim()));
    ok(!!tBtn && menu().some(b => /Unscheduled/.test(b.textContent)), 'the submenu lists the sprints and Unscheduled');
    click(tBtn);
    const moved = state().items.find(i => i.id === id);
    ok(moved.startDay === window.RM.sprintStartDay(state().meta, target), 'picking a sprint moves the item there');
    ok(doc.querySelector('#sprintView .spv-row[data-spid="' + id + '"]').dataset.spsec === String(target), 'and it lists under that sprint');
    // story rows: Move to sprint gives the story its own timeline
    click(doc.querySelector('#sprintView [data-spdd="level"]'));
    click(menu().find(b => /Story/.test(b.textContent)));
    const stRow = doc.querySelector('#sprintView .spv-row.spv-st');
    stRow.dispatchEvent(new window.MouseEvent('contextmenu', { bubbles: true, cancelable: true, clientX: 120, clientY: 160 }));
    const mv2 = menu().find(b => /Move to sprint/.test(b.textContent));
    ok(!!mv2, 'story rows have a context menu with Move to sprint');
    click(mv2);
    const t2 = secNums[0];
    click(menu().find(b => new RegExp('^Sprint ' + t2 + '\\b').test(b.textContent.trim())));
    const st = state().items.find(i => i.id === stRow.dataset.spid).stories.find(s => s.id === stRow.dataset.spst);
    ok(st.startDay === window.RM.sprintStartDay(state().meta, t2) && st.durDays != null, 'the story gets its own timeline in that sprint');
    const stRow2 = doc.querySelector('#sprintView .spv-row[data-spst="' + st.id + '"]');
    stRow2.dispatchEvent(new window.MouseEvent('contextmenu', { bubbles: true, cancelable: true, clientX: 120, clientY: 160 }));
    click(menu().find(b => /With feature/.test(b.textContent)));
    const st2 = state().items.find(i => i.id === stRow.dataset.spid).stories.find(s => s.id === st.id);
    ok(st2.startDay == null, 'With feature drops the story timeline again');
    click(doc.querySelector('#sprintView [data-spdd="level"]'));
    click(menu().find(b => /Feature/.test(b.textContent)));
  }
}


// stories: Duplicate story in every story context menu
{
  const menu = () => [...doc.querySelectorAll('#popover .menu-list button')];
  const ctx = (el) => el.dispatchEvent(new window.MouseEvent('contextmenu', { bubbles: true, cancelable: true, clientX: 150, clientY: 150 }));
  click(doc.querySelector('#viewTabs [data-view="planning"]'));
  const it = state().items.find(i => !i.milestone && i.stories && i.stories.length);
  window.__headway.selectItem(it.id);
  const chev = doc.querySelector('#rows .row.item[data-id="' + it.id + '"] [data-act="stories"]');
  if (chev && !chev.classList.contains('open')) click(chev);
  const st = it.stories[0];
  const stRow = doc.querySelector('#rows .row.story[data-story="' + st.id + '"]');
  ok(!!stRow, 'the feature shows its story row');
  ctx(stRow);
  const dup = menu().find(b => /Duplicate story/.test(b.textContent));
  ok(!!dup, 'planning story rows offer Duplicate story');
  const nBefore = it.stories.length;
  click(dup);
  const after = state().items.find(i => i.id === it.id);
  const idx = after.stories.findIndex(s => s.id === st.id);
  const copy = after.stories[idx + 1];
  ok(after.stories.length === nBefore + 1 && copy && copy.id !== st.id && copy.title === st.title + ' (copy)', 'the copy sits right after the original with a fresh id and "(copy)" title');
  ok(copy.size === st.size && copy.priority === st.priority && JSON.stringify(copy.assignees || []) === JSON.stringify(st.assignees || []), 'fields carry over');
  ok(!copy.jiraKey, 'the copy does not point at the original Jira issue');
  ok(new Set(after.stories.map(s => s.id)).size === after.stories.length, 'story ids stay unique');
  // the other three menus offer it too
  ctx(doc.querySelector('#panel'));
  ok(menu().some(b => /Duplicate story/.test(b.textContent)), 'the panel story menu offers Duplicate story');
  // separator placement matches the other story menus: Duplicate sits directly above Delete
  {
    const kids = [...doc.querySelectorAll('#popover .menu-list > *')];
    const di = kids.findIndex(n => /Duplicate story/.test(n.textContent));
    ok(di !== -1 && kids[di + 1] && kids[di + 1].tagName === 'BUTTON' && /Delete story/.test(kids[di + 1].textContent),
      'the panel story menu puts Delete story directly under Duplicate story, with no separator between');
  }
  click(doc.querySelector('#viewTabs [data-view="prio"]'));
  click(doc.querySelector('#prioView [data-prdd="level"]'));
  click(menu().find(b => /Story/.test(b.textContent)));
  const prCard = doc.querySelector('#prioView .pr-card[data-prst]');
  if (prCard) { ctx(prCard); ok(menu().some(b => /Duplicate story/.test(b.textContent)), 'prioritizing story cards offer Duplicate story'); }
  else ok(true, 'no story cards on the board in this fixture state');
  click(doc.querySelector('#prioView [data-prdd="level"]'));
  click(menu().find(b => /Feature/.test(b.textContent)));
  click(doc.querySelector('#viewTabs [data-view="sprints"]'));
  click(doc.querySelector('#sprintView [data-spdd="level"]'));
  click(menu().find(b => /Story/.test(b.textContent)));
  const spRow = doc.querySelector('#sprintView .spv-row.spv-st');
  ctx(spRow);
  ok(menu().some(b => /Duplicate story/.test(b.textContent)), 'sprinting story rows offer Duplicate story');
  click(doc.querySelector('#sprintView [data-spdd="level"]'));
  click(menu().find(b => /Feature/.test(b.textContent)));
  // restore
  window.HeadwayApp.ai.commit('undo dup', (s) => { const t = window.RM.itemById(s, it.id); t.stories = t.stories.filter(x => x.id !== copy.id); });
  click(doc.querySelector('#viewTabs [data-view="planning"]'));
}
// sprinting: titles are text until double-click / Rename; right panel like Scoping
{
  const menu = () => [...doc.querySelectorAll('#popover .menu-list button')];
  click(doc.querySelector('#viewTabs [data-view="sprints"]'));
  const row = doc.querySelector('#sprintView .spv-sec:not([data-spsec="u"]) .spv-row');
  const id = row.dataset.spid;
  const it = state().items.find(i => i.id === id);
  const title = row.querySelector('[data-spf="feature"]');
  ok(title && title.tagName !== 'INPUT' && title.textContent === it.feature, 'the row title is text, not an input');
  ok(!!row.querySelector('.spv-fill'), 'a filler pushes the chips right of the narrow title');
  // click selects and opens the panel
  click(title);
  const pName = doc.querySelector('#panel textarea[data-f="feature"]');
  ok(!doc.querySelector('#panel').hidden && pName && pName.value === it.feature,
    'clicking a row shows the item in the right panel');
  // double-click edits; Enter commits (re-query: selecting re-rendered the row)
  const rowSel = () => doc.querySelector('#sprintView .spv-row[data-spid="' + id + '"]');
  rowSel().querySelector('[data-spf="feature"]').dispatchEvent(new window.MouseEvent('dblclick', { bubbles: true }));
  const ed = rowSel().querySelector('input.st-add-input, [data-spf="feature"]');
  ok(ed && (ed.tagName === 'INPUT' || ed.isContentEditable), 'double-click makes the title editable');
  ok(ed.dataset.spf === 'feature', 'the editor keeps the field name');
  ed.value = 'Renamed via sprint';
  ed.dispatchEvent(new window.KeyboardEvent('keydown', { key: 'Enter', bubbles: true }));
  ed.dispatchEvent(new window.FocusEvent('blur', { bubbles: true }));
  ok(state().items.find(i => i.id === id).feature === 'Renamed via sprint', 'Enter commits the new title');
  // context menu offers Rename…
  rowSel().dispatchEvent(new window.MouseEvent('contextmenu', { bubbles: true, cancelable: true, clientX: 120, clientY: 120 }));
  ok(menu().some(b => /Rename/.test(b.textContent)), 'the row context menu offers Rename…');
  click(menu().find(b => /Rename/.test(b.textContent)));
  const ed2 = rowSel().querySelector('input.st-add-input');
  ok(!!ed2, 'Rename… starts the same inline edit');
  ed2.dispatchEvent(new window.KeyboardEvent('keydown', { key: 'Escape', bubbles: true }));
  ok(state().items.find(i => i.id === id).feature === 'Renamed via sprint', 'Escape leaves the title alone');
  // panel peek / toggle on Sprinting
  window.dispatchEvent(new window.KeyboardEvent('keydown', { key: ']', bubbles: true }));
  ok(doc.querySelector('#panel').hidden && !doc.querySelector('#panelPeek').hidden,
    '] hides the panel and shows the peek button on Sprinting');
  window.dispatchEvent(new window.KeyboardEvent('keydown', { key: ']', bubbles: true }));
  ok(!doc.querySelector('#panel').hidden, '] brings it back');
  // story rows: text title, story panel on select
  click(doc.querySelector('#sprintView [data-spdd="level"]'));
  click(menu().find(b => /Story/.test(b.textContent)));
  const stRow = doc.querySelector('#sprintView .spv-row.spv-st');
  ok(stRow && stRow.querySelector('[data-spf="story"]').tagName !== 'INPUT', 'story titles are text too');
  click(stRow.querySelector('[data-spf="story"]'));
  ok(!doc.querySelector('#panel').hidden && !!doc.querySelector('#panel textarea[data-stf="title"]'),
    'selecting a story row shows the story panel');
  click(doc.querySelector('#sprintView [data-spdd="level"]'));
  click(menu().find(b => /Feature/.test(b.textContent)));
  window.HeadwayApp.ai.commit('restore', (s) => { window.RM.itemById(s, id).feature = it.feature; });
  // sprint totals in story points when the story size scheme is numeric
  {
    const scheme0 = state().meta.storySizeScheme;
    const fRow = doc.querySelector('#sprintView .spv-sec:not([data-spsec="u"]) .spv-row');
    const fid = fRow.dataset.spid, secKey = fRow.dataset.spsec;
    const hd = () => doc.querySelector('#sprintView .spv-sec[data-spsec="' + secKey + '"] .spv-sechd .pr-lanect');
    const side = () => doc.querySelector('#sprintView .spv-sbtn[data-spside="' + secKey + '"] .pr-lanect');
    window.HeadwayApp.ai.commit('story points setup', (s) => {
      window.RM.setSizeScheme(s, 'fibonacci', 'story');
      const f = window.RM.itemById(s, fid);
      f.stories = f.stories || [];
      while (f.stories.length < 3) f.stories.push({ id: window.RM.uid('s'), title: 'pt story', done: false });
      f.stories.slice(0, 3).forEach((st) => { st.size = ''; st.startDay = null; st.durDays = null; });
    });
    ok(hd().textContent.trim() === String(doc.querySelectorAll('#sprintView .spv-sec[data-spsec="' + secKey + '"] .spv-row').length),
      'a section whose stories are all unsized keeps the plain row count, not "0 pt"');
    ok(side().textContent.trim() === hd().textContent.trim(), 'the sidebar entry shows the same count');
    window.HeadwayApp.ai.commit('story points', (s) => {
      const f = window.RM.itemById(s, fid);
      f.stories[0].size = '3'; f.stories[1].size = '5'; f.stories[2].size = '';
    });
    ok(/^\d+ pt$/.test(hd().textContent.trim()), 'a numeric story scheme totals the sprint in points once a story is sized');
    ok(Number(hd().textContent.replace(' pt', '')) === 8 && side().textContent.trim() === hd().textContent.trim(),
      'sizing two stories 3 and 5 adds 8 points to the sprint total (an unsized story adds none)');
    window.HeadwayApp.ai.commit('tshirt stories', (s) => { window.RM.setSizeScheme(s, 'tshirt', 'story'); });
    ok(hd().textContent.trim() === String(doc.querySelectorAll('#sprintView .spv-sec[data-spsec="' + secKey + '"] .spv-row').length),
      'a non-numeric story scheme keeps the plain item count');
    ok(side().classList.contains('spv-ct') && side().parentNode.lastElementChild === side() &&
      hd().classList.contains('spv-ct') && hd().parentNode.lastElementChild === hd(),
      'the total sits last in the side entry and the section heading (right edge)');
    // story-points capacity: planned points against the sprint's supply
    const capPrev = { on: state().meta.capacityEnabled, mode: state().meta.capMode };
    window.HeadwayApp.ai.commit('points capacity', (s) => {
      window.RM.setSizeScheme(s, 'fibonacci', 'story');
      const f = window.RM.itemById(s, fid);
      f.stories[0].size = '3'; f.stories[1].size = '5';
      s.meta.capacityEnabled = true; s.meta.capMode = 'points';
    });
    const capX = window.RM.capacity(state());
    const perX = capX.periods.findIndex((p) => p.num === Number(secKey));
    const yExp = perX === -1 ? NaN : capX.types.reduce((a, t) => a + capX.rows[t][perX].supply, 0);
    const xy = /^(\d+(?:\.\d)?) \/ (\d+(?:\.\d)?)$/.exec(hd().textContent.trim());
    ok(!!xy && Number(xy[1]) === 8 && Math.abs(Number(xy[2]) - yExp) < 0.06,
      'points capacity: the heading reads planned / available points for the sprint (' + hd().textContent.trim() + ' vs ' + yExp + ')');
    ok(side().textContent.trim() === hd().textContent.trim(), 'the side entry reads the same X / Y');
    ok(side().parentNode.lastElementChild === side(), 'and the X / Y stays last in the side entry');
    ok(!hd().classList.contains('over') === !(8 > yExp + 1e-9), 'over only when the sprint asks more than it has');
    const uSide = doc.querySelector('#sprintView .spv-sbtn[data-spside="u"] .spv-ct');
    ok(!!uSide && uSide.textContent.indexOf('/') === -1, 'Unscheduled shows its points alone (' + (uSide && uSide.textContent) + ')');
    window.HeadwayApp.ai.commit('overbook', (s) => { window.RM.itemById(s, fid).stories[0].size = '9999'; });
    ok(hd().classList.contains('over') && side().classList.contains('over'), 'a sprint asking more points than it has reads over');
    // a filter hides some of the sprint's stories: X is a partial count, so no red
    {
      const phName = state().phases.find((p) => p.id === state().items.find((i) => i.id === fid).phaseId).name;
      const pickPh = (re) => {
        click(doc.querySelector('#sprintView [data-spdd="fphase"]'));
        click([...doc.querySelectorAll('#popover .menu-list button')].find((b) => re.test(b.textContent.trim())));
      };
      pickPh(new RegExp('^' + phName.replace(/[.*+?^${}()|[\]\\]/g, '\\$&') + '$'));
      ok(!!hd() && !hd().classList.contains('over') && /\(filtered\)/.test(hd().getAttribute('title') || ''),
        'with a filter on the total is not red and its tooltip says (filtered) (' + (hd() && hd().getAttribute('title')) + ')');
      pickPh(/All phases/);
      ok(hd().classList.contains('over'), 'clearing the filter brings the red back');
    }
    // no story points to add up, or no sprints: the plain count as before
    window.HeadwayApp.ai.commit('tshirt under points', (s) => { window.RM.setSizeScheme(s, 'tshirt', 'story'); });
    ok(hd().textContent.trim() === String(doc.querySelectorAll('#sprintView .spv-sec[data-spsec="' + secKey + '"] .spv-row').length) &&
      !hd().classList.contains('over'), 'points capacity with a non-numeric story scheme: the plain count, no "0 / Y"');
    window.dispatchEvent(new window.KeyboardEvent('keydown', { key: 'z', metaKey: true, bubbles: true }));
    window.HeadwayApp.ai.commit('sprints off', (s) => { s.meta.weeksPerSprint = 0; });
    ok(doc.body.dataset.view === 'planning' && doc.querySelector('#viewTabs [data-view="sprints"]').hidden,
      'with sprints off Sprinting hides and the view falls back to Planning');
    window.dispatchEvent(new window.KeyboardEvent('keydown', { key: 'z', metaKey: true, bubbles: true }));
    window.HeadwayApp.ai.setView('sprints');
    window.HeadwayApp.ai.commit('person mode', (s) => { s.meta.capMode = 'person'; });
    ok(/^\d+(\.\d)? pt$/.test(hd().textContent.trim()) && !hd().classList.contains('over'),
      'per-person mode: the heading shows just the points total');
    window.HeadwayApp.ai.commit('restore capacity', (s) => { s.meta.capacityEnabled = capPrev.on; s.meta.capMode = capPrev.mode; });
    window.HeadwayApp.ai.commit('restore sizing', (s) => {
      window.RM.setSizeScheme(s, scheme0, 'story');
      const f = window.RM.itemById(s, fid);
      f.stories.slice(0, 3).forEach((st) => { st.size = ''; });
    });
  }
  click(doc.querySelector('#viewTabs [data-view="planning"]'));
}

ok(JSON.parse(window.localStorage.getItem('headway-v1')).items.length > 100, 'commits autosave to localStorage');
// desktop: reload from disk only while auto-save is on
{
  ok(typeof window.HeadwayApp.autoSaveOn === 'function', 'HeadwayApp exposes autoSaveOn()');
  const was = window.HeadwayApp.autoSaveOn();
  click(doc.querySelector('.menu-btn[data-menu="file"]'));
  const tog = [...doc.querySelectorAll('#popover .menu-list button')].find(b => /Auto.?save/i.test(b.textContent));
  if (tog) {
    click(tog);
    ok(window.HeadwayApp.autoSaveOn() === !was, 'autoSaveOn() follows the File menu toggle');
    click(doc.querySelector('.menu-btn[data-menu="file"]'));
    click([...doc.querySelectorAll('#popover .menu-list button')].find(b => /Auto.?save/i.test(b.textContent)));
    ok(window.HeadwayApp.autoSaveOn() === was, 'toggling back restores it');
  } else ok(true, 'auto-save toggle is desktop-only in this build');
  const desk = fs.readFileSync(path.join(ROOT, 'js/desktop.js'), 'utf8');
  ok(/autoSaveOn\(\)/.test(desk) && /if \(!hit \|\| reloading \|\| !app\(\)\.autoSaveOn\(\)\) return;/.test(desk),
    'the disk watcher skips reloads while auto-save is off');
}

// rich text: URLs render as links; ⌘-click opens them; storage stays plain
{
  const url = 'https://example.com/path?q=1';
  const firstRow = doc.querySelector('#rows .row.item[data-id]');
  const it = state().items.find(i => i.id === firstRow.dataset.id);
  window.HeadwayApp.ai.commit('desc', (s) => { window.RM.itemById(s, it.id).description = '<p>See ' + url + ', then (https://b.example/x). Not a link: example.com</p>'; });
  // scoping tab renders the anchors
  click(doc.querySelector('#viewTabs [data-view="scoping"]'));
  const cell = doc.querySelector('#rows .row.item[data-id="' + it.id + '"] .sc-rich[data-scope="description"]');
  ok(!!cell, 'scoping shows the description cell');
  const links = cell ? [...cell.querySelectorAll('a.auto-link')] : [];
  ok(links.length === 2 && links[0].getAttribute('href') === url && links[0].textContent === url, 'a URL in a scoping cell renders as an auto-link');
  ok(links[1] && links[1].getAttribute('href') === 'https://b.example/x' && /\)\./.test(cell.textContent), 'trailing punctuation stays outside the link');
  ok(!/example\.com<\/a>/.test(cell.innerHTML.replace(url, '')), 'a bare domain without a scheme is not linked');
  ok(state().items.find(i => i.id === it.id).description.indexOf('<a') === -1, 'the stored value carries no anchor');
  // ⌘-click opens; a plain click does not
  const opened = [];
  const realOpen = window.open;
  window.open = (u) => { opened.push(u); return null; };
  links[0].dispatchEvent(new window.MouseEvent('click', { bubbles: true, cancelable: true }));
  ok(opened.length === 0, 'a plain click does not open the link');
  links[0].dispatchEvent(new window.MouseEvent('click', { bubbles: true, cancelable: true, metaKey: true }));
  ok(opened.length === 1 && opened[0] === url, '⌘-click opens the URL in a new tab');
  links[0].dispatchEvent(new window.MouseEvent('click', { bubbles: true, cancelable: true, ctrlKey: true }));
  ok(opened.length === 2, 'Ctrl-click opens it too');
  window.open = realOpen;
  // editing the cell and blurring stores plain text, not the anchor markup
  cell.innerHTML = '<p>Go to <a class="auto-link" href="' + url + '">' + url + '</a> now</p>';
  cell.dispatchEvent(new window.Event('input', { bubbles: true }));
  cell.dispatchEvent(new window.FocusEvent('focusout', { bubbles: true }));
  const stored = state().items.find(i => i.id === it.id).description;
  ok(stored.indexOf('<a') === -1 && stored.indexOf(url) !== -1, 'reading an edited cell back drops the anchor and keeps the URL text');
  // the right panel's description editor links too
  click(doc.querySelector('#viewTabs [data-view="planning"]'));
  window.__headway.selectItem(it.id);
  const ed = doc.querySelector('#panel .wz-ed[data-f="col:description"]');
  ok(!!ed && !!ed.querySelector('a.auto-link[href="' + url + '"]'), 'the panel description editor renders the URL as a link');
  // desktop shell: capability and opener wiring exist
  const cap = fs.readFileSync(path.join(ROOT, 'src-tauri/capabilities/default.json'), 'utf8');
  const desk12 = fs.readFileSync(path.join(ROOT, 'js/desktop.js'), 'utf8');
  ok(/opener:allow-open-url/.test(cap) && /openUrl/.test(desk12), 'the desktop shell can open URLs');
}

// reporting: Delivery-by-sprint counts an item in every sprint it is in flight
{
  const meta0 = state().meta;
  const it = state().items.find(i => !i.milestone && i.startDay != null);
  const prev = { startDay: it.startDay, durDays: it.durDays, riskDays: it.riskDays, locked: it.locked };
  const perSprint = window.RM.sprintDays(meta0);
  const nums = [];
  for (let w = 0; w < meta0.numWeeks; w++) { const n = window.RM.sprintNumForWeek(meta0, w); if (nums.indexOf(n) === -1) nums.push(n); }
  const startNum = nums[0];
  // feature counts per delivery bar, keyed by the bar's label
  const delivCounts = () => {
    click(doc.querySelector('#viewTabs [data-view="reports"]'));
    const card = [...doc.querySelectorAll('#reportsView .rp-card, .rp-card')].find(c => c.querySelector('h2') && /Delivery by/.test(c.querySelector('h2').textContent));
    const out = {};
    if (!card) return out;
    [...card.querySelectorAll('.rp-bar-row')].forEach((r) => {
      const m = r.querySelector('.rp-bar-val').textContent.match(/(\d+) feature/);
      out[r.querySelector('.rp-bar-label').textContent] = m ? Number(m[1]) : 0;
    });
    return out;
  };
  window.HeadwayApp.ai.commit('span', (s) => { const t = window.RM.itemById(s, it.id); t.startDay = null; t.durDays = perSprint; t.riskDays = 0; t.locked = false; });
  const base = delivCounts();
  window.HeadwayApp.ai.commit('span', (s) => { const t = window.RM.itemById(s, it.id); t.startDay = window.RM.sprintStartDay(s.meta, startNum); t.durDays = perSprint * 3; });
  const spanned = delivCounts();
  const grew = Object.keys(spanned).filter(k => spanned[k] - (base[k] || 0) === 1);
  ok(grew.length === 3, 'a three-sprint feature counts in three delivery bars (grew in ' + grew.length + ')');
  // sprinting still lists that row exactly once, under the sprint it starts in
  click(doc.querySelector('#viewTabs [data-view="sprints"]'));
  const rows = doc.querySelectorAll('#sprintView .spv-row[data-spid="' + it.id + '"]');
  ok(rows.length === 1 && rows[0].dataset.spsec === String(startNum), 'sprinting still lists the spanning row once, in its start sprint');
  window.HeadwayApp.ai.commit('span', (s) => {
    const t = window.RM.itemById(s, it.id);
    t.startDay = prev.startDay; t.durDays = prev.durDays; t.riskDays = prev.riskDays; t.locked = prev.locked;
  });
  const back = state().items.find(i => i.id === it.id);
  ok(back.startDay === prev.startDay && back.durDays === prev.durDays, 'the item is restored to its original schedule');
}

// rich text: focusing and blurring an unchanged legacy value commits nothing
{
  const firstRow = doc.querySelector('#rows .row.item[data-id]');
  const it = state().items.find(i => i.id === firstRow.dataset.id);
  const legacy = 'See https://example.com/a\nline two';
  window.HeadwayApp.ai.commit('desc', (s) => { window.RM.itemById(s, it.id).description = legacy; });
  click(doc.querySelector('#viewTabs [data-view="scoping"]'));
  const cell = doc.querySelector('#rows .row.item[data-id="' + it.id + '"] .sc-rich[data-scope="description"]');
  ok(!!cell && !!cell.querySelector('a.auto-link'), 'the legacy value renders with a link');
  const before = JSON.stringify(window.__headway.getState());
  cell.dispatchEvent(new window.FocusEvent('focusin', { bubbles: true }));
  cell.dispatchEvent(new window.FocusEvent('focusout', { bubbles: true }));
  const after = JSON.stringify(window.__headway.getState());
  ok(before === after, 'a no-op focus/blur on a linkified cell commits nothing (state and history unchanged)');
  ok(state().items.find(i => i.id === it.id).description === legacy, 'the stored description is byte-identical');
}

// sprinting: story level hides sprints whose features have no visible stories
{
  click(doc.querySelector('#viewTabs [data-view="sprints"]'));
  click(doc.querySelector('#sprintView [data-spdd="level"]'));
  click([...doc.querySelectorAll('#popover .menu-list button')].find(b => /Story/.test(b.textContent)));
  const btns = [...doc.querySelectorAll('#sprintView .spv-sbtn')].filter(b => b.dataset.spside !== 'u');
  const emptyEnd = b => b && parseInt(b.querySelector('.pr-lanect').textContent.trim(), 10) === 0 && !b.classList.contains('today');
  ok(btns.length === 0 || (!emptyEnd(btns[0]) && !emptyEnd(btns[btns.length - 1])),
    'story level trims empty sprints at both ends');
  click(doc.querySelector('#sprintView [data-spdd="level"]'));
  click([...doc.querySelectorAll('#popover .menu-list button')].find(b => /Feature/.test(b.textContent)));
}

// 0.5 story points: option offered, totals sum fractions
{
  click(doc.querySelector('#viewTabs [data-view="sprints"]'));
  const prev = JSON.parse(JSON.stringify({ s: state().meta.storySizeScheme, o: state().meta.storySizeOrder, d: state().meta.storySizeDays }));
  window.HeadwayApp.ai.commit('fib', (s) => { window.RM.setSizeScheme(s, 'fibonacci', 'story'); });
  ok(state().meta.storySizeOrder.join(',') === '0,0.5,1,2,3,5,8,13', 'the Fibonacci story scale offers 0 and 0.5');
  const row = doc.querySelector('#sprintView .spv-sec:not([data-spsec="u"]) .spv-row');
  const id = row.dataset.spid, sec = row.dataset.spsec;
  const it = state().items.find(i => i.id === id);
  if (it.stories.length >= 2) {
    const before = doc.querySelector('#sprintView .spv-sec[data-spsec="' + sec + '"] .pr-lanect').textContent;
    window.HeadwayApp.ai.commit('half', (s) => { const t = window.RM.itemById(s, id); t.stories[0].size = '0.5'; t.stories[1].size = '3'; });
    const txt = doc.querySelector('#sprintView .spv-sec[data-spsec="' + sec + '"] .pr-lanect').textContent;
    ok(/^\d+(\.5)? pt$/.test(txt) && txt !== before, 'sprint totals include half points and show one decimal at most (' + txt + ')');
    window.HeadwayApp.ai.commit('half', (s) => { const t = window.RM.itemById(s, id); t.stories[0].size = ''; t.stories[1].size = ''; });
  } else ok(true, 'first sprint row has fewer than two stories');
  window.HeadwayApp.ai.commit('restore', (s) => { s.meta.storySizeScheme = prev.s; s.meta.storySizeOrder = prev.o; s.meta.storySizeDays = prev.d; });
  click(doc.querySelector('#viewTabs [data-view="planning"]'));
}

// ------------------------------------------------ auto-sized features + story snap (Stories level)
{
  const restore = JSON.stringify(state());
  const prevSnapStory = window.HeadwayApp.ai.ui().snapStory;
  const prevSnapFeat = window.HeadwayApp.ai.ui().snapFeat;
  window.HeadwayApp.ai.setPref('snapFeat', 'day');
  window.HeadwayApp.ai.setPref('snapStory', 'day');
  click(doc.querySelector('#viewTabs [data-view="planning"]'));
  const f0 = [...doc.querySelectorAll('#rows .row.item[data-id]')]
    .map(r => state().items.find(i => i.id === r.dataset.id))
    .find(i => i && !i.milestone && (i.stories || []).length >= 1);
  const stSizes = window.RM.sizeOrderOf(state(), 'story').filter(l => window.RM.sizeDays(state(), l, 'story') >= 3);
  const stSize = stSizes[0];
  window.HeadwayApp.ai.commit('auto-sized setup', (s) => {
    s.meta.planLevel = 'story';
    s.meta.capacityEnabled = true;
    const t = s.items.find(i => i.id === f0.id);
    t.size = 'XL';
    t.stories.forEach((st) => { st.size = null; });
    t.stories[0].size = stSize;
  });
  const ff = () => state().items.find(i => i.id === f0.id);
  ok(ff().size === 'XL', 'nothing re-sizes a feature on commit: the hand-picked XL stands');
  click(doc.querySelector('#viewTabs [data-view="planning"]'));
  const szA = doc.querySelector('#rows .row.item[data-id="' + f0.id + '"] [data-act="size"]');
  ok(!!szA && !szA.classList.contains('ro'), 'the size chip is an ordinary, editable size');
  // the Auto timeline button sizes the phase's features from their stories, once
  const zapF = doc.querySelector('#rows .row.band[data-phase="' + f0.phaseId + '"] .band-zap');
  ok(!!zapF && !zapF.disabled, 'the feature’s phase has an enabled Auto timeline button');
  click(zapF);
  const derived = window.RM.sizeForDays(state(), window.RM.autoSizeDays(state(), ff(), { snap: { feature: 'day' } }));
  ok(ff().size === derived && derived !== 'XL',
    'the click sizes the feature from the span its stories cover (' + ff().size + '), not the hand-picked XL');
  const zapF2 = doc.querySelector('#rows .row.band[data-phase="' + f0.phaseId + '"] .band-zap');
  ok(zapF2.disabled, 'and right after, the button is disabled');
  window.__headway.selectItem(f0.id);
  ok(!!doc.querySelector('#panel [data-f="size"]'), 'the panel keeps its size buttons: the size stays editable');
  window.HeadwayApp.ai.commit('hand size', (s) => { s.items.find(i => i.id === f0.id).size = 'XS'; });
  ok(ff().size === 'XS', 'a hand-written size after the click sticks');
  ok(!doc.querySelector('#rows .row.band[data-phase="' + f0.phaseId + '"] .band-zap').disabled,
    'and the button offers to re-derive it');

  // a 3-day story under a week feature snap derives from 5 days
  const s3 = window.RM.sizeOrderOf(state(), 'story').find(l => window.RM.sizeDays(state(), l, 'story') === 3);
  window.HeadwayApp.ai.commit('three-day story', (s) => {
    const t = s.items.find(i => i.id === f0.id);
    t.stories.forEach((st, ix) => {
      st.size = ix ? null : s3; st.startDay = 0; st.durDays = ix ? 1 : 3;
    });
  });
  const rawH = window.RM.autoSizeDays(state(), ff());
  const roundH = window.RM.autoSizeDays(state(), ff(), { snap: { feature: 'week' } });
  ok(roundH > rawH, 'the week snap rounds the short hull up (' + rawH + ' → ' + roundH + ')');
  // the click sizes the feature from that rounded span, and a size set from
  // stories never reads as a size/bar mismatch: the bar is the stories' hull
  window.HeadwayApp.ai.setPref('snapFeat', 'week');
  const zapW = doc.querySelector('#rows .row.band[data-phase="' + f0.phaseId + '"] .band-zap');
  if (zapW && !zapW.disabled) click(zapW);
  const szB = doc.querySelector('#rows .row.item[data-id="' + f0.id + '"] [data-act="size"]');
  ok(!!szB && !szB.classList.contains('custom'),
    'a derived size never reads as a size/bar mismatch under a week feature snap');
  window.HeadwayApp.ai.setPref('snapFeat', 'day');


  // a story buys whole snap units: 3 days under a week snap stores 5
  window.HeadwayApp.ai.setPref('snapStory', 'week');
  const stId0 = ff().stories[0].id;
  if (!doc.querySelector('#rows .row.story[data-story="' + stId0 + '"]')) {
    click(doc.querySelector('#rows .row.item[data-id="' + f0.id + '"] .r-chev'));
  }
  const stRowA = doc.querySelector('#rows .row.story[data-story="' + stId0 + '"]');
  if (stRowA) click(stRowA.querySelector('.row-left') || stRowA);
  const durIn = doc.querySelector('#panel [data-stf="durWeeks"]');
  const stNow = () => ff().stories.find(st => st.id === stId0);
  if (durIn && stNow().startDay != null) {
    durIn.value = '0.6'; // 3 days
    durIn.dispatchEvent(new window.Event('change', { bubbles: true }));
    ok(window.RM.workInSpan(state().meta, stNow().startDay, stNow().durDays) === 5,
      'a 3-day story under the week snap buys a whole week (' + stNow().durDays + ')');
  } else {
    ok(false, 'the story panel offers a duration field for a scheduled story');
  }
  // …and so does a 3-day story SIZE
  const szChipS = doc.querySelector('#rows .row.story[data-story="' + stId0 + '"] [data-act="st-size"]');
  if (s3 && szChipS && stNow().startDay != null) {
    click(szChipS);
    const opt = [...doc.querySelectorAll('#popover .menu-list button')].find(b => b.textContent.trim().indexOf(s3) === 0);
    click(opt);
    ok(window.RM.workInSpan(state().meta, stNow().startDay, stNow().durDays) === 5,
      'a 3-day story size under the week snap buys a whole week too (' + stNow().durDays + ')');
  } else {
    ok(false, 'the story row offers a size chip with a 3-day option');
  }
  window.HeadwayApp.ai.setPref('snapStory', prevSnapStory);
  window.HeadwayApp.ai.setPref('snapFeat', prevSnapFeat);
  window.HeadwayApp.ai.commit('restore', (s) => {
    const d = JSON.parse(restore);
    s.meta = d.meta; s.phases = d.phases; s.items = d.items;
  });
  window.__headway.selectItem(null);
  ok(state().items.find(i => i.id === f0.id).size === JSON.parse(restore).items.find(i => i.id === f0.id).size,
    'the document is back the way the suite found it');
}

// ---------------------------------------------------------------- tags (panel editor + filter)
let tagFilterChecks = () => Promise.resolve();
let taggedForXlsx = null;
{
  click(doc.querySelector('#viewTabs [data-view="planning"]'));
  const visibleIds = [...doc.querySelectorAll('#rows .row.item[data-id]')].map(r => r.dataset.id);
  const visible = state().items.filter(i => visibleIds.indexOf(i.id) !== -1 && !i.milestone);
  const tagged = visible.find(i => i.stories.length) || visible[0];
  const other = visible.find(i => i.id !== tagged.id);
  // a second item carries a tag already, so the datalist has something to offer
  window.HeadwayApp.ai.commit('tags', (s) => { window.RM.setTags(s, window.RM.itemById(s, other.id), ['beta']); });

  click(doc.querySelector('#rows .row.item[data-id="' + tagged.id + '"] .row-left'));
  ok(!!doc.querySelector('#panel .p-sec[data-sec="tags"]'), 'the feature panel has a Tags section');
  const tagIn = () => doc.querySelector('#panel .p-tag-in');
  ok(!!tagIn(), 'the Tags section offers an input');
  const typeTag = (v, k) => {
    const inp = tagIn();
    inp.value = v;
    inp.dispatchEvent(new window.KeyboardEvent('keydown', { key: k || 'Enter', bubbles: true, cancelable: true }));
  };
  // the focus-restore call lands inside a requestAnimationFrame; run it
  // synchronously here so the assertion below doesn't need to wait a real frame
  {
    const realRaf = window.requestAnimationFrame;
    window.requestAnimationFrame = (cb) => cb();
    typeTag('alpha');
    window.requestAnimationFrame = realRaf;
  }
  const itemTags = id => state().items.find(i => i.id === id).tags;
  ok(String(itemTags(tagged.id)) === 'alpha', 'Enter in the tag input stores the tag (' + itemTags(tagged.id) + ')');
  ok(!!doc.querySelector('#panel .tag-chip [data-tagrm="alpha"]'), 'the tag renders as a removable chip');
  ok(doc.querySelector('#panel .p-tag-in') === doc.activeElement, 'the input survives the re-render');
  // a comma commits too, and duplicates collapse
  typeTag('gamma', ',');
  ok(String(itemTags(tagged.id)) === 'alpha,gamma', 'a comma commits the tag as well');
  typeTag('ALPHA');
  ok(String(itemTags(tagged.id)) === 'alpha,gamma', 'a case-insensitive duplicate is ignored');
  const opts = [...doc.querySelectorAll('#panel #tagOptions option')].map(o => o.value);
  ok(opts.indexOf('beta') !== -1 && opts.indexOf('alpha') === -1,
    'the datalist offers other documents tags but not the ones already on this item (' + opts.join(',') + ')');
  // Backspace on an empty input drops the last tag
  const bsInp = tagIn();
  bsInp.value = '';
  bsInp.dispatchEvent(new window.KeyboardEvent('keydown', { key: 'Backspace', bubbles: true, cancelable: true }));
  ok(String(itemTags(tagged.id)) === 'alpha', 'Backspace in an empty input removes the last tag');
  // blur with text commits
  const blurInp = tagIn();
  blurInp.value = 'delta';
  blurInp.dispatchEvent(new window.FocusEvent('focusout', { bubbles: true }));
  ok(String(itemTags(tagged.id)) === 'alpha,delta', 'blurring with text adds the tag');
  click(doc.querySelector('#panel .tag-chip [data-tagrm="delta"]'));
  ok(String(itemTags(tagged.id)) === 'alpha', 'clicking a chips x removes that tag');

  // the story panel edits story tags the same way
  click(doc.querySelector('#panel [data-pst-edit]'));
  ok(!!doc.querySelector('#panel .p-crumb') && !!doc.querySelector('#panel .p-sec[data-sec="st-tags"]'),
    'the story panel has a Tags section too');
  typeTag('sigma');
  const storyTags = () => state().items.find(i => i.id === tagged.id).stories[0].tags;
  ok(String(storyTags()) === 'sigma', 'the story panel stores tags on the story');
  click(doc.querySelector('#panel .p-crumb'));

  // a story tag makes its feature match; #alpha searches tags only
  const storyTagged = visible.find(i => i.id !== tagged.id && i.stories.length && i.id !== other.id);
  if (storyTagged) {
    window.HeadwayApp.ai.commit('tags', (s) => {
      window.RM.setTags(s, window.RM.itemById(s, storyTagged.id).stories[0], ['alpha']);
    });
  }
  // the row filter is debounced (120 ms), so these run from the async chain below
  const rowIdsNow = () => [...doc.querySelectorAll('#rows .row.item[data-id]')].map(r => r.dataset.id);
  const allRows = rowIdsNow().length;
  const typeFilter = (v) => new Promise((res) => {
    const rf2 = doc.querySelector('#rowFilter');
    rf2.value = v;
    rf2.dispatchEvent(new window.Event('input', { bubbles: true }));
    window.setTimeout(res, 220);
  });
  tagFilterChecks = () => typeFilter('#alpha').then(() => {
    const hits = rowIdsNow();
    ok(hits.length < allRows && hits.indexOf(tagged.id) !== -1,
      '#alpha narrows the rows to tagged items (' + hits.length + ' of ' + allRows + ')');
    ok(!storyTagged || hits.indexOf(storyTagged.id) !== -1, 'a feature whose story carries the tag matches too');
    ok(hits.indexOf(other.id) === -1, 'an item tagged beta is filtered out by a tag-only query');
    return typeFilter('');
  }).then(() => {
    ok(rowIdsNow().length === allRows, 'clearing the filter restores every row');
    // the tags stay on the document so the export chain below can round-trip them
    taggedForXlsx = { id: tagged.id, storyId: tagged.stories[0] && tagged.stories[0].id };
  });
}

// explicit 0-point stories: sized, worth nothing, offered in the size picker
{
  click(doc.querySelector('#viewTabs [data-view="sprints"]'));
  const prevZ = JSON.parse(JSON.stringify({ s: state().meta.storySizeScheme, o: state().meta.storySizeOrder, d: state().meta.storySizeDays }));
  window.HeadwayApp.ai.commit('fib0', (s) => { window.RM.setSizeScheme(s, 'fibonacci', 'story'); });
  const rowZ = doc.querySelector('#sprintView .spv-sec:not([data-spsec="u"]) .spv-row');
  const idZ = rowZ.dataset.spid, secZ = rowZ.dataset.spsec;
  // every story in this section: only the chosen feature's two get a size
  window.HeadwayApp.ai.commit('zero', (s) => {
    s.items.forEach((it) => (it.stories || []).forEach((st) => { st.size = ''; }));
    const t = window.RM.itemById(s, idZ);
    t.stories[0].size = '0';
    if (t.stories[1]) t.stories[1].size = '0';
  });
  const hdZ = doc.querySelector('#sprintView .spv-sec[data-spsec="' + secZ + '"] .pr-lanect').textContent.trim();
  ok(hdZ === '0 pt', 'a section of only 0-point stories reads "0 pt", not a count (' + hdZ + ')');
  // the story size picker offers 0 (story level shows the per-story chips)
  click(doc.querySelector('#sprintView [data-spdd="level"]'));
  click([...doc.querySelectorAll('#popover .menu-list button')].find((b) => /Story/.test(b.textContent)));
  const stId0 = state().items.find((i) => i.id === idZ).stories[0].id;
  const stRow = doc.querySelector('#sprintView .spv-row[data-spst="' + stId0 + '"]');
  const szBtn = stRow && stRow.querySelector('[data-spact="st-size"]');
  ok(!!szBtn, 'the 0-point story shows a size chip');
  if (szBtn) {
    click(szBtn);
    const opts = [...doc.querySelectorAll('#popover .menu-list button')].map((b) => b.textContent.trim());
    ok(opts.some((o) => /^0\b/.test(o)), 'the story size picker offers 0 (' + opts.slice(0, 3).join(' | ') + ')');
    window.dispatchEvent(new window.KeyboardEvent('keydown', { key: 'Escape', bubbles: true }));
  } else ok(window.RM.sizeOrderOf(state(), 'story')[0] === '0', 'the story size scale offers 0');
  click(doc.querySelector('#sprintView [data-spdd="level"]'));
  click([...doc.querySelectorAll('#popover .menu-list button')].find((b) => /Feature/.test(b.textContent)));
  window.HeadwayApp.ai.commit('restore0', (s) => {
    s.items.forEach((it) => (it.stories || []).forEach((st) => { st.size = ''; }));
    s.meta.storySizeScheme = prevZ.s; s.meta.storySizeOrder = prevZ.o; s.meta.storySizeDays = prevZ.d;
  });
  click(doc.querySelector('#viewTabs [data-view="planning"]'));
}

// NOTE: this export promise chain must stay LAST in this file — its .then /
// .catch bodies run after every synchronous block, and the .then calls
// process.exit. New blocks go ABOVE this line.
tagFilterChecks().then(() => window.RMExcel.exportWorkbook(state())).then((buf) => {
  const bytes = buf.size != null ? buf.size : buf.byteLength;
  ok(buf && bytes > 20000, 'xlsx export produced a workbook (' + bytes + ' bytes)');
  // a disk reload (third arg) swaps the data under the screen: the view, the
  // selection and the file's saved UI prefs are left exactly as they were
  const firstNum = state().items[0].num;
  window.HeadwayAI.runTool('navigate', { view: 'scoping', num: firstNum }, window.HeadwayApp);
  ok(window.HeadwayApp.ai.ui().view === 'scoping' && window.HeadwayApp.ai.ui().selectedNum === firstNum, 'setup: Scoping view with a selection');
  const edited = state();
  edited.items[0].feature = 'Renamed on disk';
  // the file carries another machine's prefs (Planning view) — ignored on reload
  // auto-order runs when a document opens and would legitimately dirty a file
  // whose rows are not in start order — switch it off to judge the reload alone
  window.HeadwayApp.ai.setPref('autoOrder', false);
  return window.RMExcel.exportWorkbook(edited, { view: 'planning' }).then((b1) => b1.arrayBuffer ? b1.arrayBuffer() : b1).then((ab1) =>
    window.HeadwayApp.loadBuffer(Buffer.from(new Uint8Array(ab1)), 'Roadmap.xlsx', true)
  ).then(() => {
    ok(state().items[0].feature === 'Renamed on disk', 'disk reload: the data updated');
    if (taggedForXlsx) {
      const rtIt = state().items.find(i => i.id === taggedForXlsx.id);
      ok(rtIt && String(rtIt.tags) === 'alpha', 'xlsx round trip: the feature keeps its tag (' + (rtIt && rtIt.tags) + ')');
      const rtSt = rtIt && rtIt.stories.find(x => x.id === taggedForXlsx.storyId);
      ok(!rtSt || String(rtSt.tags) === 'sigma', 'xlsx round trip: the story keeps its tag');
    }
    ok(window.HeadwayApp.ai.ui().view === 'scoping', 'disk reload: the view did not change');
    ok(window.HeadwayApp.ai.ui().selectedNum === firstNum, 'disk reload: the selection survived');
    ok(window.HeadwayApp.unsavedNow() === false, 'disk reload: nothing to save');
    window.HeadwayApp.ai.setPref('autoOrder', true);
    // priority chips and the panel picker carry a color tier class
    const rowIds = [].slice.call(doc.querySelectorAll('#rows .row.item[data-id]')).map((r) => r.dataset.id);
    ok(rowIds.length > 1, 'setup: rows on screen to color');
    window.HeadwayApp.ai.commit('priority', (s) => {
      window.RM.setPriorityScheme(s, 'moscow');
      window.RM.itemById(s, rowIds[0]).priority = 'M';
      window.RM.itemById(s, rowIds[1]).priority = 'W';
    });
    ok(!!doc.querySelector('.row[data-id="' + rowIds[0] + '"] .r-risk.pri.pt-crit'), 'a Must chip is tier crit');
    ok(!!doc.querySelector('.row[data-id="' + rowIds[1] + '"] .r-risk.pri.pt-low'), 'a Won’t chip is tier low');
    window.HeadwayAI.runTool('navigate', { view: 'scoping', num: state().items.find((i) => i.id === rowIds[0]).num }, window.HeadwayApp);
    ok(!!doc.querySelector('#panel .seg button.pt-crit.on'), 'the panel picker lights Must in tier crit');
    ok(!!doc.querySelector('#panel .seg button.pt-high'), 'Should sits in tier high');
    // prioritizing: Priority/Risk chips stay available whatever the columns are; phase filter
    {
      click(doc.querySelector('#viewTabs [data-view="prio"]'));
      const pick = re => click([...doc.querySelectorAll('#popover .menu-list button')].find(b => re.test(b.textContent)));
      // columns by priority: the Fields menu still offers Priority and Risk, and cards show them
      click(doc.querySelector('#prioView [data-prdd="cols"]'));
      pick(/^Priority$/);
      click(doc.querySelector('#prFieldsBtn'));
      const labels = [...doc.querySelectorAll('#popover .menu-list button')].map(b => b.textContent.trim());
      ok(labels.indexOf('Priority') !== -1 && labels.indexOf('Risk') !== -1, 'Fields menu offers Priority and Risk even when the columns are by priority');
      click(doc.querySelector('#prioView'));
      ok(!!doc.querySelector('#prioView .pr-card [data-pract="priority"]') && !!doc.querySelector('#prioView .pr-card [data-pract="risk"]'),
        'cards show the priority and risk chips in the priority columns');
      click(doc.querySelector('#prioView [data-prdd="cols"]'));
      pick(/^Phase$/);
      // phase filter, same as Sprinting
      const cardsOf = () => [...doc.querySelectorAll('#prioView .pr-card[data-prcard]')];
      const itemOf = c => state().items.find(i => i.id === c.dataset.prcard);
      const before = cardsOf().length;
      const ph = state().phases[1];
      const fph = doc.querySelector('#prioView [data-prdd="fphase"]');
      ok(!!fph, 'the prioritizing toolbar offers a phase filter');
      click(fph);
      pick(new RegExp('^' + ph.name.replace(/[.*+?^${}()|[\]\\]/g, '\\$&') + '$'));
      ok(cardsOf().length > 0 && cardsOf().length < before && cardsOf().every(c => itemOf(c).phaseId === ph.id),
        'the phase filter narrows the cards to that phase');
      const fph2 = doc.querySelector('#prioView [data-prdd="fphase"]');
      ok(fph2.classList.contains('pr-on') && fph2.textContent.indexOf(ph.name) !== -1, 'an active phase filter lights up and names the phase');
      click(fph2); pick(/All phases/);
      ok(cardsOf().length === before && !doc.querySelector('#prioView [data-prdd="fphase"]').classList.contains('pr-on'),
        'clearing the phase filter restores every card');
      click(doc.querySelector('#viewTabs [data-view="sprints"]'));
      ok(!doc.querySelector('#sprintView [data-spdd="fphase"].pr-on'), 'the prioritizing phase filter does not leak into Sprinting');
      click(doc.querySelector('#viewTabs [data-view="planning"]'));
    }
  }).then(() => {
  // opening a file is not an edit: no version-history entry, nothing unsaved
  const tpl = window.__headway.templateState();
  const tplHist = (tpl.history || []).length;
  return window.RMExcel.exportWorkbook(tpl).then((b2) => b2.arrayBuffer ? b2.arrayBuffer() : b2).then((ab) =>
    window.HeadwayApp.loadBuffer(Buffer.from(new Uint8Array(ab)), 'Template.xlsx', true)
  ).then(() => {
    ok(state().meta.title === 'Template', 'the template workbook opened');
    ok((state().history || []).length === tplHist,
      'opening a file adds no version-history entry (' + (state().history || []).length + ' vs ' + tplHist + ')');
    ok(!(state().history || []).some(h => h.label === 'open'), 'no "open" entry in the history');
    ok(window.HeadwayApp.unsavedNow() === false, 'a freshly opened doc has nothing to save');
  });
  });
}).then(() => {
  // ---- story numbers and story-to-story dependencies survive a save/reload
  // (the Stories sheet's own "#" / "Depends on" columns are asserted in
  // tests/core.test.js — ExcelJS writes whole rows through `row.values =
  // [...]`, which silently drops arrays built in this jsdom realm)
  const dependentId = 'story-dep-probe';
  window.HeadwayApp.ai.commit('add a dependent story', (s) => {
    const it = s.items[0];
    it.stories.push({ id: dependentId, title: 'Dependent story', num: window.RM.nextNum(s), deps: [it.stories[0].num] });
  });
  const host0 = state().items[0];
  ok(host0.stories.length >= 2 && host0.stories.every((x) => x.num > 0),
    'every story carries a number (' + host0.stories.map((x) => x.num).join(',') + ')');
  const depNum = host0.stories[0].num;
  ok(String(window.RM.storyRef(state(), dependentId).st.deps) === String(depNum),
    'the dependency is stored as a story number');
  return window.RMExcel.exportWorkbook(state())
    .then((b) => (b.arrayBuffer ? b.arrayBuffer() : b))
    .then((ab) => window.HeadwayApp.loadBuffer(Buffer.from(new Uint8Array(ab)), 'Deps.xlsx', true))
    .then(() => {
      const back = window.RM.storyRef(state(), dependentId);
      ok(!!back && String(back.st.deps) === String(depNum),
        'story dependencies survive an xlsx round trip (got ' + (back && back.st.deps) + ')');
      ok(!!back && back.st.num > 0 && state().items[0].stories[0].num === depNum,
        'story numbers survive an xlsx round trip');
    });
}).then(() => {
  // opening a document runs auto-order only: the rows come back in start
  // order, but nothing is laid out until someone clicks Auto timeline
  const undo = () => window.dispatchEvent(new window.KeyboardEvent('keydown', { key: 'z', metaKey: true, bubbles: true }));
  window.HeadwayApp.ai.setPref('autoOrder', true);
  const doc0 = window.RM.clone(state());
  doc0.meta.capacityEnabled = true; doc0.meta.planLevel = 'feature'; doc0.meta.capMode = 'person';
  doc0.meta.capRowTypes = ['Design']; // an old document's tracked-type selection
  doc0.team = [{ id: 'solo', name: 'Solo', capType: 'Development', weekHours: {}, capacity: 1 }];
  const autoPh = doc0.phases.find((p) => !p.bucket);
  autoPh.auto = true; // an older document's flag: ignored now
  // three features of the Auto phase pile into the same week: the phase floor
  // holds them at that week, so capacity has to spread them forward from it
  const seed = doc0.items.filter((i) => !i.milestone)[0];
  [1, 2].forEach((n) => {
    const copy = window.RM.clone(seed);
    copy.id = 'pile' + n; copy.num = 900 + n; copy.feature = 'Pile ' + n; copy.stories = []; copy.deps = [];
    doc0.items.push(copy);
  });
  const piled = doc0.items.filter((i) => !i.milestone).slice(0, 3);
  piled.forEach((it) => { it.phaseId = autoPh.id; it.locked = false; it.done = false; it.capType = 'Development'; it.capMult = 1; it.startDay = 10; it.durDays = 5; it.deps = []; });
  doc0.items.reverse(); // and the rows arrive out of start order
  return window.RMExcel.exportWorkbook(doc0)
    .then((b) => (b.arrayBuffer ? b.arrayBuffer() : b))
    .then((ab) => window.HeadwayApp.loadBuffer(Buffer.from(new Uint8Array(ab)), 'Auto.xlsx'))
    .then(() => {
      const opened = state();
      ok(!opened.phases.some((p) => 'auto' in p), 'on open: an old per-phase Auto flag is dropped');
      ok(window.RM.capacity(opened).weeks.some((c) => c.over), 'on open: nothing is laid out — the piled week is still over capacity');
      ok(opened.meta.capRowTypes === undefined, 'on open: an old capRowTypes selection is dropped');
      const sorted = window.RM.clone(opened);
      window.RM.sortItemsByStart(sorted);
      ok(sorted.items.map((i) => i.id).join() === opened.items.map((i) => i.id).join(), 'on open: the rows are in start order');
      const before = JSON.stringify(state().items);
      undo(); // the open cleared the stack — nothing to step back to
      ok(JSON.stringify(state().items) === before, 'on open: undo cannot step back past the open');
      ok(window.HeadwayApp.ai.autoTimelineNow(autoPh.id) > 0, 'Auto timeline on the phase moves the piled work');
      ok(!window.RM.capacity(state()).weeks.some((c) => c.over), 'and no week is over capacity after');
      ok(state().history[state().history.length - 1].label === 'auto timeline', 'the newest version-history entry is labelled auto timeline');
      const hLen = state().history.length;
      ok(window.HeadwayApp.ai.autoTimelineNow(autoPh.id) === 0, 'a second pass over a laid-out phase moves nothing');
      ok(state().history.length === hLen, 'and a pass that moves nothing adds no version-history entry');
    });
}).then(() => {
  console.log('\n' + passed + ' passed, ' + failed + ' failed');
  process.exit(failed ? 1 : 0);
}).catch((e) => {
  failed++;
  console.error('  ✗ export threw: ' + e.message);
  console.log('\n' + passed + ' passed, ' + failed + ' failed');
  process.exit(1);
});
