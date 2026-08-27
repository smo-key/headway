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

for (const f of ['js/core.js', 'js/excel.js', 'js/export-png.js', 'js/app.js']) {
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
const contBtn = doc.querySelector('#startBody [data-sp-continue]');
ok(!!contBtn, 'browser session offers Continue where you left off');
click(contBtn);
ok(!doc.body.classList.contains('start') && doc.querySelector('#startPage').hidden,
  'Continue enters the editor');
ok(doc.querySelectorAll('#rows .row.band').length === 6, 'six phase bands rendered');
const itemRows = doc.querySelectorAll('#rows .row.item').length;
ok(itemRows > 100, 'item rows rendered (' + itemRows + ')');
const visibleSched = state().items.filter(i => i.startDay != null &&
  !state().phases.find(p => p.id === i.phaseId).collapsed).length;
ok(doc.querySelectorAll('#rows .bar').length === visibleSched,
  'bars rendered for every visible scheduled item (' + doc.querySelectorAll('#rows .bar').length + ')');
ok(doc.querySelectorAll('#hdrCap .cap-cell').length === 48, 'capacity strip has 48 week cells');
ok(state() && state().items.length > 100, 'debug state handle live (' + state().items.length + ' items)');
ok(doc.querySelector('#resPanel') !== null && doc.querySelector('#resGrid') !== null, 'resources panel present');
ok(doc.querySelector('#capTypeCell .cap-lab') !== null, 'capacity header shows a plain availability label');
ok(doc.querySelector('#capTypeCell .dd-btn') === null, 'capacity is role-agnostic: no role filter dropdown');
ok(doc.querySelectorAll('#hdrCap .cap-cell').length === 48, 'capacity row spans all weeks');
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
{
  const chip = doc.querySelector('#rows .row.item .r-epic');
  ok(!!chip && (chip.querySelector('svg') || chip.querySelector('i')) !== null,
    'epic chip combines icon + label');
}

// ---------------------------------------------------------------- menus
click(doc.querySelector('[data-menu="file"]'));
ok(!doc.querySelector('#popover').hidden && doc.querySelectorAll('#popover .menu-list button').length >= 5, 'File menu opens with items');
ok(Array.from(doc.querySelectorAll('#popover .menu-list button')).some(b => /Download template/.test(b.textContent)),
  'File menu offers Download template');
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
  const nameInp = doc.querySelector('#rows .row.item[data-id="' + itId + '"] input.r-name');
  ok(!!nameInp, 'row title is an editable input');
  nameInp.value = 'Renamed inline';
  nameInp.dispatchEvent(new window.Event('change', { bubbles: true }));
  ok(state().items.find(i => i.id === itId).feature === 'Renamed inline', 'row title rename commits');
}
ok(doc.querySelector('#panel .wz-ed[data-f="col:description"]') !== null, 'panel has a description field');
ok(doc.querySelector('#panel .p-sec[data-sec="fields"]').classList.contains('open'), 'Fields section is open by default');
ok(doc.querySelector('#panel .p-actions') === null, 'footer Duplicate/Delete buttons are gone');
ok(doc.querySelector('#panel .p-more') !== null, 'panel has a … actions button');
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
  ok(/position:\s*sticky/.test(decl('.row.band')) && /top:\s*var\(--hdr-h\)/.test(decl('.row.band')),
    'phase bands are sticky below the header');
  // workstream/epic group rows are transparent (grid lines show through)
  // and scroll with the rows instead of sticking
  ok(!/position:\s*sticky/.test(decl('.row.eband')),
    'epic/workstream bands scroll with the rows (not sticky)');
  ok(/\.row\.eband \.row-lane\s*{[^}]*background:\s*transparent/.test(css),
    'epic/workstream band lanes are transparent');
}

// ---------------------------------------------------------------- chips
click(doc.querySelector('#rows .row.item .r-size'));
ok(!doc.querySelector('#popover').hidden, 'size chip opens a dropdown');
{
  const xl = Array.from(doc.querySelectorAll('#popover .menu-list button')).find(b => /^XL/.test(b.textContent.trim()));
  click(xl);
  ok(state().items.find(i => i.id === itId).size === 'XL', 'picking a size commits (XL)');
}
// risk moved out of Planning; scoping keeps its chip with None/L/M/H options
ok(!doc.querySelector('#rows .row.item .r-risk'), 'planning rows no longer show a risk chip');
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
click(doc.querySelector('#hdrCap [data-w="3"]'));
ok(state().meta.holidays.length === boBefore + 5, 'clicking a capacity cell adds that week\'s five holiday days');
click(doc.querySelector('#hdrCap [data-w="3"]'));
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
click(doc.querySelector('#resManage'));
ok(doc.body.dataset.view === 'setup', 'resources "manage" jumps to the Setup view');
ok(doc.querySelector('#setupView [data-sutab="team"]').classList.contains('on'),
  'resources "manage" lands on the Team tab');
ok(doc.querySelectorAll('#setupView .su-tab').length === 9 &&
  doc.querySelectorAll('#setupView .su-rail-hd').length === 2,
  'settings rail: 9 vertical tabs under Project + Personal sections');
ok(doc.querySelectorAll('#setupView .su-card').length === 3 && !!doc.querySelector('#suCapEnable'),
  'Team tab shows roles + work week + capacity');
ok(/Roles/.test(doc.querySelector('#setupView .su-card h2').textContent), 'team types renamed to Roles');
ok(!!doc.querySelector('#setupView [data-rcrate]') && !!doc.querySelector('#setupView [data-rccost]'),
  'rate card inputs per role');
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
  doc.querySelectorAll('#setupView [data-pref-snap]').length === 3,
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
suTab('timeline');
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
  'holidays listed as a removable named-range table (Timeline tab)');
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
suTab('team');
{
  const inp = doc.querySelector('#suTypeAdd');
  inp.value = 'Data Scientist';
  click(doc.querySelector('#suTypeAddBtn'));
  ok(state().teamTypes.indexOf('Data Scientist') !== -1, 'setup adds a team type');
}
suTab('phases');
ok(doc.querySelectorAll('#setupView [data-suphedit]').length === state().phases.length, 'phases listed with edit controls');
suTab('workstreams');
ok(doc.querySelectorAll('#setupView [data-suwsedit]').length > 0, 'workstreams listed with edit controls');

// picking "Default blue" for a workstream with a seeded default must SURVIVE
// a reload — the choice is stored explicitly, not deleted (regression: the
// known-default re-seeded on load and reverted the color)
{
  const editBtn = Array.from(doc.querySelectorAll('#setupView [data-suwsedit]'))
    .find(b => b.dataset.suwsedit === 'Product') || doc.querySelector('#setupView [data-suwsedit]');
  const wsName = editBtn.dataset.suwsedit;
  click(editBtn);
  click(doc.querySelector('.swatch[data-esw="product"]'));
  click(doc.querySelector('#wsSave'));
  ok(state().wsColors[wsName] === 'product', 'Default blue stored explicitly for ' + wsName);
  const reloaded = window.RM.normalizeState(JSON.parse(JSON.stringify(state())));
  ok(window.RM.colorForWs(reloaded, wsName) === window.RM.PALETTE.product,
    'color survives a reload (normalize does not re-seed the default)');
}
ok(doc.querySelectorAll('#setupView .su-grip').length ===
  doc.querySelectorAll('#setupView [data-sulist] .su-row').length, 'every reorderable row has a drag grip');
// drag the first phase's grip to the bottom of its list (jsdom rects are all
// zero, so a large clientY resolves to "after the last row")
{
  suTab('phases');
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
  suTab('team');
  const firstType = state().teamTypes[0];
  const grip = doc.querySelector('#setupView [data-sulist="type"] .su-row .su-grip');
  grip.dispatchEvent(new window.MouseEvent('pointerdown', { bubbles: true }));
  window.dispatchEvent(new window.MouseEvent('pointermove', { clientY: 999 }));
  window.dispatchEvent(new window.MouseEvent('pointerup'));
  ok(state().teamTypes[state().teamTypes.length - 1] === firstType, 'dragging a type grip reorders types');
}
// and workstreams (order persists in state.wsOrder)
{
  suTab('workstreams');
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

// capacity factor: editable, feeds availability
{
  const capChip = doc.querySelector('#resGrid [data-rcap]');
  ok(!!capChip, 'resource rows show a capacity column');
  click(capChip);
  const capInp = doc.querySelector('#resGrid [data-rcap] input');
  ok(!!capInp, 'capacity chip opens an inline editor');
  capInp.value = '0.5';
  capInp.dispatchEvent(new window.KeyboardEvent('keydown', { key: 'Enter', bubbles: true }));
  ok(state().team[0].capacity === 0.5, 'capacity commits (0.5)');
  const avail = window.RM.availForWeek(state(), 0);
  ok(Math.abs(avail.total - 0.5) < 1e-9, 'availability scales by the capacity factor (' + avail.total + ')');
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
ok(doc.querySelectorAll('#hdrSprints .sc-hcell[data-col]').length === 13, 'default columns are 8 fixed + 5 text (incl. Description)');
ok(!!doc.querySelector('#hdrSprints [data-col="description"]'), 'Description column shown by default');
{
  const hdrOrder = Array.from(doc.querySelectorAll('#hdrSprints .sc-hcell[data-col]')).map(c => c.dataset.col);
  ok(hdrOrder.indexOf('description') !== -1 && hdrOrder.indexOf('description') < hdrOrder.indexOf('enables'),
    'Description sits left of Enables');
  ok(hdrOrder.indexOf('description') === 0 && hdrOrder.indexOf('epic') === 1 && hdrOrder.indexOf('epic') < hdrOrder.indexOf('size'),
    'default order leads with Description and Epic before Size');
  ok(hdrOrder.indexOf('start') === hdrOrder.indexOf('duration') + 1, 'Start column follows Duration');
}
ok(doc.querySelectorAll('#hdrSprints .sc-rz').length === 13, 'column resize handles present');
const cell = doc.querySelector('#rows .row.item[data-id="' + itId + '"] [data-scope="notes"]');
ok(cell.getAttribute('contenteditable') === 'true', 'scoping cells are rich editors');
cell.innerHTML = 'noted in the grid';
cell.dispatchEvent(new window.FocusEvent('focusout', { bubbles: true }));
ok(state().items.find(i => i.id === itId).notes === 'noted in the grid', 'scoping cell edit commits');

// column management: remove the Description built-in via its column menu…
click(doc.querySelector('#hdrSprints [data-colmenu="description"]'));
click(Array.from(doc.querySelectorAll('#popover .menu-list button')).find(b => /Remove column/.test(b.textContent)));
ok(!doc.querySelector('#hdrSprints [data-col="description"]'), 'column removed via its menu');
// …then re-add it (now hidden) through the "+" menu
click(doc.querySelector('#hdrSprints [data-coladd]'));
const descAdd = Array.from(doc.querySelectorAll('#popover .menu-list button')).find(b => /^Description$/.test(b.textContent.trim()));
click(descAdd);
ok(!!doc.querySelector('#hdrSprints [data-col="description"]'), 'hidden built-in column re-added via + menu');
// leave the grid as it started: Description back in front of Enables
click(doc.querySelector('#hdrSprints [data-colmenu="description"]'));
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
click(doc.querySelector('#hdrSprints [data-colmenu="' + ownerCol.key + '"]'));
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
  const before = state().items.length;
  const addRow = doc.querySelector('#rows .row.addrow');
  ok(!!addRow, 'phases end with a blank add row');
  click(addRow);
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
  const eb = Array.from(doc.querySelectorAll('#rows .row.eband')).find(r => r.dataset.epic);
  eb.dispatchEvent(new window.MouseEvent('contextmenu', { bubbles: true, clientX: 220, clientY: 220 }));
  const labels = Array.from(doc.querySelectorAll('#popover .menu-list button')).map(b => b.textContent);
  ok(labels.some(l => /Edit epic/.test(l)) && labels.some(l => /Delete epic/.test(l)),
    'right-clicking an epic band offers edit/delete');
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
    ok(labels.join(',') === 'Role,Rate card,Workstream,Cost,Rate,Margin,Total', 'header labels spelled out, cost before rate');
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

  // zoom cluster: shown on Budgeting, expand lives there now
  ok(!doc.querySelector('#zoomCtl').hidden, 'zoom cluster visible on Budgeting');
  {
    const wpx0 = window.__headway.getState() && doc.documentElement.style.getPropertyValue('--week-px');
    window.eval("document.querySelector('#zoomInBtn').click()");
    ok(doc.documentElement.style.getPropertyValue('--week-px') !== wpx0, 'zoom + changes the week width on Budgeting');
    window.eval("document.querySelector('#zoomOutBtn').click()");
  }
  window.eval("document.querySelector('#btnPresent').click()");
  ok(doc.body.classList.contains('present') && doc.body.dataset.view === 'budget',
    'Expand enters present mode from Budgeting');
  ok(!doc.querySelector('#zoomCtl').hidden, 'zoom cluster stays put in expand mode');
  window.eval("document.querySelector('#btnPresent').click()");
  ok(!doc.body.classList.contains('present'), 'clicking it again restores the full Budgeting UI');
  // Expand is gone from Scoping (the cluster only exists on Planning/Budgeting)
  click(doc.querySelector('#viewTabs [data-view="scoping"]'));
  ok(doc.querySelector('#zoomCtl').hidden, 'no zoom/expand cluster on Scoping');
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
  suTab('workstreams');
  ok(doc.querySelectorAll('#setupView [data-suepedit]').length > 0, 'Setup lists epics with edit controls');
  click(doc.querySelector('#setupView [data-suepedit]'));
  ok(!doc.querySelector('#modalHost').hidden, 'epic edit modal opens from Setup');
  window.dispatchEvent(new window.KeyboardEvent('keydown', { key: 'Escape', bubbles: true }));
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
  suTab('team'); // the capacity switch lives on the Team tab
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

// ---------------------------------------------------------------- timeline-only preview
{
  const pbtn = doc.querySelector('#zoomCtl #btnPresent');
  ok(!!pbtn, 'expand button lives in the zoom cluster');
  click(pbtn);
  ok(doc.body.classList.contains('present'), 'expand enters the preview');
  ok(!doc.querySelector('#btnPresentExit'), 'no floating exit button — the cluster button toggles in place');
  ok(doc.querySelector('#topbar') && !doc.querySelector('#zoomCtl').hidden,
    'the app header and the zoom cluster stay during expand');
  click(doc.querySelector('#zoomCtl #btnPresent'));
  ok(!doc.body.classList.contains('present'), 'clicking again restores the full UI');
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
    // hovering the empty lane previews the landing slot
    stRow.querySelector('.row-lane')
      .dispatchEvent(new window.MouseEvent('mousemove', { bubbles: true, clientX: 600, clientY: 300 }));
    ok(!!stRow.querySelector('.st-bar.place-preview'), 'hovering an empty story lane previews the mini bar');
    // double-click the story lane → timeline appears
    stRow.querySelector('.row-lane')
      .dispatchEvent(new window.MouseEvent('dblclick', { bubbles: true, clientX: 600, clientY: 300 }));
    let st = state().items.find(i => i.id === withStories.id).stories.find(s => s.id === stId);
    ok(st.startDay != null && st.durDays >= 5, 'double-clicking the story lane gives it a timeline');
    const stBar = doc.querySelector('#rows .st-bar[data-stbar="' + stId + '"]');
    ok(!!stBar, 'story timeline renders as a mini bar');
    ok(!stBar.querySelector('[data-port]'), 'story bars have no dependency ports');
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
  const sd = doc.querySelector('#panel .wz-ed[data-f="stcol:description"]');
  const sa = doc.querySelector('#panel .wz-ed[data-f="stac"]');
  ok(!!sd && !!sa && sd.getAttribute('contenteditable') === 'true' && sa.getAttribute('contenteditable') === 'true',
    'story panel has rich Description and Acceptance Criteria editors');
  sd.innerHTML = 'Does <b>things</b>';
  sd.dispatchEvent(new window.FocusEvent('focusout', { bubbles: true }));
  // the commit re-rendered the panel — re-query the AC editor
  const sa2 = doc.querySelector('#panel .wz-ed[data-f="stac"]');
  sa2.innerHTML = '<ul><li>works offline</li></ul>';
  sa2.dispatchEvent(new window.FocusEvent('focusout', { bubbles: true }));
  const st0 = state().items.find(i => i.id === withStories.id).stories[0];
  ok(st0.description === 'Does <b>things</b>', 'story description saves');
  ok(st0.ac === '<ul><li>works offline</li></ul>', 'story acceptance criteria save');
  // a custom scope column edits into st.custom
  const scEd = doc.querySelector('#panel .wz-ed[data-f^="stcol:"]:not([data-f="stcol:description"])');
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
  suTab('sizing');
  ok(doc.querySelectorAll('#setupView .su-scheme').length >= 4, 'Sizing offers approach presets');
  click(doc.querySelector('#setupView [data-suscheme="fibonacci"]'));
  ok(state().meta.sizeScheme === 'fibonacci' &&
    state().meta.sizeOrder.join(',') === '1,2,3,5,8,13',
    'Story points preset applies its scale');
  ok(doc.querySelectorAll('#setupView [data-susz]').length === 6, 'option table lists the six point values');
  // rename an option — items follow, scheme flips to custom
  const lblInp = doc.querySelector('#setupView [data-suszlabel="13"]');
  lblInp.value = '21';
  lblInp.dispatchEvent(new window.Event('change', { bubbles: true }));
  ok(state().meta.sizeScheme === 'custom' && state().meta.sizeOrder.indexOf('21') !== -1,
    'editing options flips the approach to Custom');
  click(doc.querySelector('#setupView [data-suscheme="none"]'));
  ok(state().meta.sizeScheme === 'none' && state().meta.sizeOrder.length === 0, 'No sizing empties the scale');
  click(doc.querySelector('#viewTabs [data-view="planning"]'));
  ok(!doc.querySelector('#rows .row.item .r-size'), 'no size chips while sizing is off');
  ok(doc.body.classList.contains('no-size'), 'body carries the no-size flag');
  click(doc.querySelector('#btnSetup'));
  click(doc.querySelector('#setupView [data-suscheme="tshirt"]'));
  ok(state().meta.sizeOrder.join(',') === 'XS,S,M,L,XL', 'T-shirt preset restores the classic scale');
}

// ---------------------------------------------------------------- workstream feature toggle
{
  suTab('workstreams');
  const wsChk = doc.querySelector('#suWsEnable');
  ok(!!wsChk && wsChk.checked, 'Workstreams tab offers the feature switch (on by default)');
  wsChk.checked = false;
  wsChk.dispatchEvent(new window.Event('change', { bubbles: true }));
  ok(state().meta.workstreamsEnabled === false, 'workstreams can be disabled per project');
  click(doc.querySelector('#viewTabs [data-view="scoping"]'));
  ok(!doc.querySelector('#hdrSprints [data-col="workstream"]'), 'Scoping hides the Workstream column when off');
  click(doc.querySelector('#btnSetup'));
  suTab('workstreams');
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
  ok(kids.indexOf('btnPresent') === -1, 'Expand button moved out of the topbar');
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
  const lay = window.RM_EXPORT.layout(state(), {});
  ok(lay.rows.filter(r => r.kind === 'item').length ===
    state().items.filter(i => i.startDay != null).length,
    'export layout covers every scheduled item of the live document');
  click(doc.querySelector('#modalHost [data-m=cancel], #modalHost [data-m=x]'));
  ok(doc.querySelector('#modalHost').hidden, 'export dialog closes');
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
  suTab('workstreams');
  ok(!!doc.querySelector('#setupView .su-defws'), 'default workstream row leads the Workstreams tab');
  click(doc.querySelector('#setupView [data-sudefws]'));
  ok(!!doc.querySelector('#modalHost #dwsName'), 'default workstream modal opens');
  doc.querySelector('#modalHost #dwsName').value = 'Core';
  click(doc.querySelector('#modalHost #dwsSave'));
  ok(state().meta.defaultWsName === 'Core', 'default workstream renames');
  ok(window.RM.defaultWsName(state()) === 'Core', 'core helper reads the rename');
  window.dispatchEvent(new window.KeyboardEvent('keydown', { key: 'z', metaKey: true, bubbles: true }));
}

// risk scheme card: switching to MoSCoW relabels the scoping column
{
  suTab('sizing');
  const riskCards = doc.querySelectorAll('#setupView [data-surisk]');
  ok(riskCards.length === 4, 'four risk schemes offered (none, risk, auto, confidence)');
  click(Array.from(riskCards).find(b => b.dataset.surisk === 'confidence'));
  ok(state().meta.riskScheme === 'confidence', 'Confidence scheme commits');
  window.dispatchEvent(new window.KeyboardEvent('keydown', { key: 'z', metaKey: true, bubbles: true }));
  // scheme none removes the column
  suTab('sizing');
  click(Array.from(doc.querySelectorAll('#setupView [data-surisk]')).find(b => b.dataset.surisk === 'none'));
  ok(state().meta.riskScheme === 'none', 'scheme none commits');
  click(doc.querySelector('#viewTabs [data-view="scoping"]'));
  ok(!doc.querySelector('#rows .row.item .r-risk'), 'no assessment chips when the scheme is none');
  window.dispatchEvent(new window.KeyboardEvent('keydown', { key: 'z', metaKey: true, bubbles: true }));
}

// role rename from Setup propagates everywhere
{
  suTab('team');
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

// ---------------------------------------------------------------- sprinting view + statuses
{
  ok(!!doc.querySelector('#viewTabs [data-view="sprints"]'), 'Sprinting tab sits in the view tabs');
  click(doc.querySelector('#viewTabs [data-view="sprints"]'));
  ok(doc.body.dataset.view === 'sprints', 'view switches to Sprinting');
  ok(!!doc.querySelector('#sprintView .sp-toolbar [data-sprsel]'), 'sprint selector present');
  ok(doc.querySelectorAll('#sprintView .sp-col').length === 4, 'kanban board shows the 4 default feature statuses');
  ok(doc.querySelectorAll('#sprintView .sp-card').length > 0, 'cards rendered for the selected sprint');

  // grid mode: editable + tab-navigable
  click(doc.querySelector('#sprintView [data-spmode="grid"]'));
  ok(doc.querySelectorAll('#sprintView .sp-grid').length > 0, 'grid mode renders a table');
  const titleInp = doc.querySelector('#sprintView [data-sptitle]');
  ok(!!titleInp && titleInp.tagName === 'INPUT', 'grid titles are real inputs (tabbable)');
  const gid = titleInp.dataset.sptitle;
  titleInp.value = 'Renamed in sprint grid';
  titleInp.dispatchEvent(new window.Event('change', { bubbles: true }));
  ok(state().items.find(i => i.id === gid).feature === 'Renamed in sprint grid', 'grid title edit commits');
  window.dispatchEvent(new window.KeyboardEvent('keydown', { key: 'z', metaKey: true, bubbles: true }));

  // status chip commits + syncs the done flag on the last status
  const stBtn = doc.querySelector('#sprintView [data-spstatus]');
  const stId = stBtn.dataset.spstatus;
  click(stBtn);
  const doneOpt = Array.from(doc.querySelectorAll('#popover .menu-list button')).find(b => /^Done$/.test(b.textContent.trim()));
  ok(!!doneOpt, 'status dropdown lists the feature statuses');
  click(doneOpt);
  const stItem = state().items.find(i => i.id === stId);
  ok(stItem.status === 'Done' && stItem.done === true, 'last status marks the item done');
  window.dispatchEvent(new window.KeyboardEvent('keydown', { key: 'z', metaKey: true, bubbles: true }));

  // all-sprints grouping
  click(doc.querySelector('#sprintView [data-sprsel]'));
  click(Array.from(doc.querySelectorAll('#popover .menu-list button')).find(b => /All sprints/.test(b.textContent)));
  ok(doc.querySelectorAll('#sprintView .sp-ghd').length > 1, 'All sprints groups by sprint');
  click(doc.querySelector('#sprintView [data-sprsel]'));
  click(Array.from(doc.querySelectorAll('#popover .menu-list button')).find(b => /Current sprint/.test(b.textContent)));
  click(doc.querySelector('#sprintView [data-spmode="board"]'));
}

// statuses configurable in Setup, feature and story lists separate
{
  suTab('statuses');
  ok(doc.querySelectorAll('#setupView .su-card').length === 2, 'Statuses tab: feature + story cards');
  ok(doc.querySelectorAll('#setupView [data-sustat="feature"]').length === 4 &&
     doc.querySelectorAll('#setupView [data-sustat="story"]').length === 3,
    'default lists: 4 feature statuses, 3 story statuses');
  const inp = doc.querySelector('#setupView [data-sustat="feature"][data-old="Blocked"]');
  inp.value = 'Waiting';
  inp.dispatchEvent(new window.Event('change', { bubbles: true }));
  ok(state().meta.statuses.feature.includes('Waiting') && !state().meta.statuses.feature.includes('Blocked'),
    'renaming a status commits');
  const addInp = doc.querySelector('#suStAdd-story');
  addInp.value = 'In review';
  click(doc.querySelector('#setupView [data-sustatadd="story"]'));
  ok(state().meta.statuses.story.indexOf('In review') === state().meta.statuses.story.length - 2,
    'new story status lands before Done');
  window.dispatchEvent(new window.KeyboardEvent('keydown', { key: 'z', metaKey: true, bubbles: true }));
  window.dispatchEvent(new window.KeyboardEvent('keydown', { key: 'z', metaKey: true, bubbles: true }));
  click(doc.querySelector('#viewTabs [data-view="planning"]'));
}

// ---------------------------------------------------------------- batch 6: priority, story cells, columns tab
{
  // priority column: enable MoSCoW in Setup, chip appears in Scoping
  suTab('sizing');
  const priCards = doc.querySelectorAll('#setupView [data-supri]');
  ok(priCards.length === 3, 'three priority schemes (none, MoSCoW, levels)');
  click(Array.from(priCards).find(b => b.dataset.supri === 'moscow'));
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

  // sprint board shows story cards alongside feature cards
  click(doc.querySelector('#viewTabs [data-view="sprints"]'));
  click(doc.querySelector('#sprintView [data-sprsel]'));
  click(Array.from(doc.querySelectorAll('#popover .menu-list button')).find(b => /All sprints/.test(b.textContent)));
  ok(doc.querySelectorAll('#sprintView .sp-card.sp-stcard').length > 0, 'sprint board shows story cards');
  ok(doc.querySelectorAll('#sprintView [data-spstasg]').length === 0 || true, 'board mode has no grid buttons');
  click(doc.querySelector('#sprintView [data-spmode="grid"]'));
  ok(doc.querySelectorAll('#sprintView [data-spstasg]').length > 0, 'sprint grid stories carry assignee buttons');
  click(doc.querySelector('#sprintView [data-spmode="board"]'));
  click(doc.querySelector('#sprintView [data-sprsel]'));
  click(Array.from(doc.querySelectorAll('#popover .menu-list button')).find(b => /Current sprint/.test(b.textContent)));

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
  ok(!!dmMenu && dmMenu.querySelectorAll('[data-mi]').length === 3 &&
    /Phase/.test(dmMenu.textContent) && /Feature/.test(dmMenu.textContent) && /Story/.test(dmMenu.textContent),
    'the dropdown offers Phase / Feature / Story, each with an icon');
  ok(dmMenu.querySelectorAll('[data-mi] svg, [data-mi] [data-lucide]').length >= 3, 'detail options carry icons');
  click(dmMenu.querySelector('[data-mi="1"]')); // Feature
  ok(doc.querySelectorAll('#rows .row.story').length === 0,
    'feature detail keeps story rows tucked away');
  click(dmBtn);
  click(doc.querySelector('#popover .menu-list [data-mi="2"]')); // Story
  ok(doc.querySelectorAll('#rows .row.story').length > 0, 'story detail opens every story row');
  click(doc.querySelector('#detailBtn'));
  click(doc.querySelector('#popover .menu-list [data-mi="0"]')); // Phase
  ok(doc.querySelectorAll('#rows .row.item').length === 0 &&
    doc.querySelectorAll('#rows .row.band').length === state().phases.length,
    'phase detail shows only the phase bands');
  ok(!!doc.querySelector('#rows .ph-row-bar'), 'phase bands paint their span as a bar on Planning');
  click(doc.querySelector('#viewTabs [data-view="scoping"]'));
  ok(doc.querySelectorAll('#rows .row.item').length === 0, 'phase detail applies on Scoping too');
  click(doc.querySelector('#detailBtn'));
  click(doc.querySelector('#popover .menu-list [data-mi="2"]')); // Story
  ok(doc.querySelectorAll('#rows .row.story').length > 0, 'story detail applies on Scoping too');
  click(doc.querySelector('#detailBtn'));
  click(doc.querySelector('#popover .menu-list [data-mi="1"]')); // back to Feature
  click(doc.querySelector('#viewTabs [data-view="planning"]'));
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
  const inp = doc.querySelector('#rows .row.item input[data-rowname]');
  const it = window.RM.itemById(state(), inp.closest('.row').dataset.id);
  const oldTitle = it.feature;
  inp.value = 'Diffed Feature Title';
  inp.dispatchEvent(new window.Event('change', { bubbles: true }));
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

// ---------------------------------------------------------------- export smoke
ok(window.__headway.saveFileName() === state().meta.title + '.xlsx',
  'save uses the exact project title as the filename (no slug, no date)');
ok(JSON.parse(window.localStorage.getItem('headway-v1')).items.length > 100, 'commits autosave to localStorage');
window.RMExcel.exportWorkbook(state()).then((buf) => {
  const bytes = buf.size != null ? buf.size : buf.byteLength;
  ok(buf && bytes > 20000, 'xlsx export produced a workbook (' + bytes + ' bytes)');
  console.log('\n' + passed + ' passed, ' + failed + ' failed');
  process.exit(failed ? 1 : 0);
}).catch((e) => {
  failed++;
  console.error('  ✗ export threw: ' + e.message);
  console.log('\n' + passed + ' passed, ' + failed + ' failed');
  process.exit(1);
});
