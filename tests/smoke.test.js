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

for (const f of ['js/core.js', 'js/excel.js', 'js/app.js']) {
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
ok(doc.querySelector('#capTypeCell .dd-btn') !== null, 'capacity work-type selector rendered');
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
ok(doc.querySelector('#panel [data-f=description]') !== null, 'panel has a description field');
ok(doc.querySelector('#panel [data-dd=epic]') !== null, 'epic is a dropdown button');
click(doc.querySelector('#panel [data-dd=epic]'));
ok(!doc.querySelector('#popover').hidden && doc.querySelectorAll('#popover .menu-list [data-mi]').length > 2, 'epic dropdown opens the shared list UI');
ok(doc.querySelectorAll('#popover .menu-list .mi-edit').length > 0, 'epic options carry an edit affordance');
doc.querySelector('#popover').hidden = true;
ok(doc.querySelector('#panel [data-f=riskSize]') !== null, 'risk size segment present');
ok(doc.querySelector('#panel [data-f=allabove]') === null, '"all items above" checkbox is gone');

// description commit
const descTa = doc.querySelector('#panel [data-f=description]');
descTa.value = 'A crisp description';
descTa.dispatchEvent(new window.Event('change', { bubbles: true }));
ok(state().items.find(i => i.id === itId).description === 'A crisp description', 'description saves');

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
  ok(labels.includes('None') && ['L', 'M', 'H'].every(v => labels.includes(v)) && !labels.includes('XL'),
    'risk options are None / L / M / H (no t-shirt sizes)');
  click(Array.from(doc.querySelectorAll('#popover .menu-list button')).find(b => b.textContent.trim() === 'L'));
  ok(state().items.find(i => i.id === itId).risk === 'L', 'picking a risk commits (L)');
}
click(doc.querySelector('#viewTabs [data-view="planning"]'));

click(doc.querySelector('#rows .row.item[data-id="' + itId + '"] .r-hc'));
{
  const hcInp = doc.querySelector('#rows .row.item[data-id="' + itId + '"] .r-hc input');
  ok(!!hcInp, 'headcount chip opens an inline number editor');
  hcInp.value = '3';
  hcInp.dispatchEvent(new window.KeyboardEvent('keydown', { key: 'Enter', bubbles: true }));
  ok(state().items.find(i => i.id === itId).headcount === 3, 'headcount edit commits (3)');
}

// decimals are valid effort levels (0.5 = half a person)
click(doc.querySelector('#rows .row.item[data-id="' + itId + '"] .r-hc'));
{
  const hcInp = doc.querySelector('#rows .row.item[data-id="' + itId + '"] .r-hc input');
  hcInp.value = '0.5';
  hcInp.dispatchEvent(new window.KeyboardEvent('keydown', { key: 'Enter', bubbles: true }));
  ok(state().items.find(i => i.id === itId).headcount === 0.5, 'headcount accepts a decimal (0.5)');
  window.dispatchEvent(new window.KeyboardEvent('keydown', { key: 'z', metaKey: true, bubbles: true }));
}

// undo via keyboard (toolbar buttons moved into the Edit menu)
window.dispatchEvent(new window.KeyboardEvent('keydown', { key: 'z', metaKey: true, bubbles: true }));
ok(state().items.find(i => i.id === itId).headcount === 1, 'undo (⌘Z) restores headcount');

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
click(doc.querySelector('#resManage'));
ok(doc.body.dataset.view === 'setup', 'resources "manage" jumps to the Setup view');
ok(doc.querySelectorAll('#setupView .su-card').length === 8, 'setup shows all eight cards (incl. Capacity)');
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
{
  const inp = doc.querySelector('#suTypeAdd');
  inp.value = 'Data Scientist';
  click(doc.querySelector('#suTypeAddBtn'));
  ok(state().teamTypes.indexOf('Data Scientist') !== -1, 'setup adds a team type');
}
ok(doc.querySelectorAll('#setupView [data-suphedit]').length === state().phases.length, 'phases listed with edit controls');
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
ok(doc.querySelectorAll('#setupView [data-suholrm]').length === state().meta.holidays.length, 'holidays listed as removable chips');
ok(doc.querySelectorAll('#setupView .su-grip').length ===
  doc.querySelectorAll('#setupView [data-sulist] .su-row').length, 'every reorderable row has a drag grip');
ok(!doc.querySelector('#setupView .su-card .m-hint') ||
  doc.querySelectorAll('#setupView .su-card .m-hint').length === 0, 'no explanatory grey text on setup cards');
// drag the first phase's grip to the bottom of its list (jsdom rects are all
// zero, so a large clientY resolves to "after the last row")
{
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
  const firstType = state().teamTypes[0];
  const grip = doc.querySelector('#setupView [data-sulist="type"] .su-row .su-grip');
  grip.dispatchEvent(new window.MouseEvent('pointerdown', { bubbles: true }));
  window.dispatchEvent(new window.MouseEvent('pointermove', { clientY: 999 }));
  window.dispatchEvent(new window.MouseEvent('pointerup'));
  ok(state().teamTypes[state().teamTypes.length - 1] === firstType, 'dragging a type grip reorders types');
}
// and workstreams (order persists in state.wsOrder)
{
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
  ok(labels.some(l => /Remove role/.test(l)) && labels.some(l => /Work type/.test(l)) &&
    labels.some(l => /Workstream/.test(l)) && labels.some(l => /Capacity/.test(l)),
    'role context menu offers rename/type/workstream/capacity/remove');
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
ok(doc.querySelectorAll('#hdrSprints .sc-hcell.sc-fixh').length === 5, 'fixed lead columns: size/risk/weeks/workstream/epic');
ok(doc.querySelectorAll('#hdrSprints .sc-hcell[data-col]').length === 9, 'default columns are 5 fixed + 4 text (no Description)');
ok(!doc.querySelector('#hdrSprints [data-col="description"]'), 'Description column hidden by default');
ok(doc.querySelectorAll('#hdrSprints .sc-rz').length === 9, 'column resize handles present');
const cell = doc.querySelector('#rows .row.item[data-id="' + itId + '"] [data-scope="notes"]');
cell.value = 'noted in the grid';
cell.dispatchEvent(new window.Event('change', { bubbles: true }));
ok(state().items.find(i => i.id === itId).notes === 'noted in the grid', 'scoping cell edit commits');

// column management: re-add the hidden Description built-in via the "+" menu
click(doc.querySelector('#hdrSprints [data-coladd]'));
const descAdd = Array.from(doc.querySelectorAll('#popover .menu-list button')).find(b => /^Description$/.test(b.textContent.trim()));
click(descAdd);
ok(!!doc.querySelector('#hdrSprints [data-col="description"]'), 'hidden built-in column re-added via + menu');
// …and remove it again through its column menu
click(doc.querySelector('#hdrSprints [data-colmenu="description"]'));
click(Array.from(doc.querySelectorAll('#popover .menu-list button')).find(b => /Remove column/.test(b.textContent)));
ok(!doc.querySelector('#hdrSprints [data-col="description"]'), 'column removed via its menu');

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
  ownerCell.value = 'Rita';
  ownerCell.dispatchEvent(new window.Event('change', { bubbles: true }));
  ok(state().items.find(i => i.id === itId).custom[ownerCol.key] === 'Rita', 'custom cell edit commits to item.custom');
}
click(doc.querySelector('#hdrSprints [data-colmenu="' + ownerCol.key + '"]'));
click(Array.from(doc.querySelectorAll('#popover .menu-list button')).find(b => /Move left/.test(b.textContent)));
ok(state().meta.scopeCols[state().meta.scopeCols.length - 2].key === ownerCol.key, 'column moved left');

// scoping swaps the wks/headcount chips for a workstream dropdown chip
ok(!!doc.querySelector('#rows .row.item .r-ws') && !doc.querySelector('#rows .row.item .r-hc'),
  'scoping shows workstream chips instead of wks/headcount');
{
  const wsChip = doc.querySelector('#rows .row.item[data-id="' + itId + '"] .r-ws');
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
  const hcNum = doc.querySelector('#panel input[data-f="headcount"]');
  ok(!!hcNum && hcNum.type === 'number', 'panel headcount is a number input');
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
ok(!!doc.querySelector('#resGrid [data-rws]'), 'resource rows have a workstream chip');

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
    const inps = Array.from(doc.querySelectorAll('#rows .row.brole[data-mid] input[data-bud]'));
    ok(inps[0].dataset.bud === 'cost' && inps[1].dataset.bud === 'rate', 'Cost input comes before Rate');
    const labels = Array.from(doc.querySelectorAll('.hl-cols .bu-only')).map(i => i.textContent);
    ok(labels.join(',') === 'Type,Workstream,Cost,Rate,Margin,Total', 'header labels spelled out, cost before rate');
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
  // reports header keeps its dropdown while collapsed (height stays put)
  ok(!!doc.querySelector('#repModeCell [data-dd="repmode"]'), 'reports header keeps the grouping dropdown when collapsed');
  // type + workstream chips edit via the shared dropdown
  click(doc.querySelector('#rows .row.brole [data-bact="type"]'));
  ok(!doc.querySelector('#popover').hidden, 'type chip opens the shared dropdown');
  const typePick = Array.from(doc.querySelectorAll('#popover .menu-list button'))[1];
  const typeName = typePick.textContent.trim();
  click(typePick);
  ok(state().team[0].type === typeName, 'picking a type commits');
  click(doc.querySelector('#rows .row.brole [data-bact="ws"]'));
  ok(!doc.querySelector('#popover').hidden, 'workstream chip opens the shared dropdown');
  doc.querySelector('#popover').hidden = true;
  click(doc.querySelector('#viewTabs [data-view="planning"]'));
}

// ---------------------------------------------------------------- reports drawer
{
  ok(doc.querySelector('#repPanel').classList.contains('collapsed'), 'reports drawer starts collapsed');
  click(doc.querySelector('#repHead .rp-title'));
  ok(!doc.querySelector('#repPanel').classList.contains('collapsed'), 'header click opens the drawer');
  const t = doc.querySelector('#repBody table');
  ok(!!t && t.querySelectorAll('tbody tr').length > 1, 'workstream report renders group rows');
  ok(/\$/.test(t.textContent), 'report prices in dollars');
  click(doc.querySelector('#repModeCell [data-dd="repmode"]'));
  click(Array.from(doc.querySelectorAll('#popover .menu-list button')).find(b => b.textContent.trim() === 'By phase'));
  ok(/^Phase/.test(doc.querySelector('#repBody thead').textContent.trim()), 'grouping switches to phase');
  click(doc.querySelector('#repHead .rp-title'));
  ok(doc.querySelector('#repPanel').classList.contains('collapsed'), 'drawer collapses again');
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
  window.eval("document.querySelector('#viewTabs [data-view=\"setup\"]').click()");
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
  window.eval("document.querySelector('#viewTabs [data-view=\"setup\"]').click()");
  const capChk2 = doc.querySelector('#suCapEnable');
  capChk2.checked = true;
  capChk2.dispatchEvent(new window.Event('change', { bubbles: true }));
  window.eval("document.querySelector('#viewTabs [data-view=\"planning\"]').click()");
  ok(!doc.body.classList.contains('no-cap'), 're-enabling restores the capacity row');
}

// ---------------------------------------------------------------- timeline-only preview
{
  const pbtn = doc.querySelector('#btnPresent');
  ok(!!pbtn, 'expand button lives in the phase lane\'s left cell');
  click(pbtn);
  ok(doc.body.classList.contains('present') && !doc.querySelector('#btnPresentExit').hidden,
    'expand enters the preview and shows the floating minimize button');
  click(doc.querySelector('#btnPresentExit'));
  ok(!doc.body.classList.contains('present') && doc.querySelector('#btnPresentExit').hidden,
    'minimize restores the full UI');
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

// ---------------------------------------------------------------- lane click deselects
{
  click(doc.querySelector('#rows .row.item .r-num'));
  ok(!doc.querySelector('#panel').hidden, 'clicking a row number opens the panel');
  // re-query: selection re-rendered the rows
  click(doc.querySelector('#rows .row.item .row-lane'));
  ok(doc.querySelector('#panel').hidden, 'clicking empty lane space deselects');
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
    // scoping shows no story information even while expanded
    click(doc.querySelector('#viewTabs [data-view="scoping"]'));
    ok(!doc.querySelector('#rows .row.story'), 'scoping renders no story rows');
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
  // deselect first (click a scheduled row's empty lane) so the panel is shut
  const schedRow = Array.from(doc.querySelectorAll('#rows .row.item')).find(r => {
    const i = state().items.find(x => x.id === r.dataset.id);
    return i && i.startDay != null;
  });
  if (schedRow) click(schedRow.querySelector('.row-lane'));
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
  ok(doc.querySelector('#panel').hidden, 'insert does not open the edit panel');
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

// ---------------------------------------------------------------- export smoke
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
