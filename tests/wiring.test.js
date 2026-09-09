/* App ↔ desktop wiring for the shared-bundle (folder) document, end to end in
 * jsdom against the in-memory fake Tauri: open a bundle in the real editor,
 * commit through the UI, watch the flush, feed peer shards through the
 * watcher, undo, plans, export, convert, menus, recents, resume, web build,
 * presence (heartbeat file, peer chips, nudge toast, cleanup).
 * Run: NODE_PATH=./node_modules node tests/wiring.test.js */
'use strict';
const fs = require('fs');
const path = require('path');

let JSDOM, VirtualConsole, ExcelJS;
try {
  ({ JSDOM, VirtualConsole } = require('jsdom'));
  ExcelJS = require('exceljs');
} catch (e) { console.log('(skipped — jsdom/exceljs not resolvable in NODE_PATH)'); process.exit(0); }

const makeFakeTauri = require('./fake-tauri.js');
const ROOT = path.join(__dirname, '..');
const html = fs.readFileSync(path.join(ROOT, 'index.html'), 'utf8');

let passed = 0, failed = 0;
function ok(cond, name) {
  if (cond) { passed++; }
  else { failed++; console.error('  ✗ ' + name); }
}
function eq(a, b, name) {
  const good = JSON.stringify(a) === JSON.stringify(b);
  if (good) passed++;
  else { failed++; console.error('  ✗ ' + name + '\n      got:  ' + JSON.stringify(a) + '\n      want: ' + JSON.stringify(b)); }
}
function section(name) { console.log('— ' + name); }
const tick = (n) => new Promise((res) => setTimeout(res, n || 5));
// the app's debounce is 1.5 s in production; boot() caps every window timer
// at 8 ms, so a few ticks let a flush (timer + promise chain) land
async function settle(n) { for (let i = 0; i < (n || 4); i++) await tick(12); }
async function until(fn, ms) {
  const t0 = Date.now();
  while (!fn()) { if (Date.now() - t0 > (ms || 5000)) return false; await tick(10); }
  return true;
}

const T0 = '2026-09-01T10:00:00.000Z';
const T2 = '2026-09-01T10:10:00.000Z';
const DIR = 'C:/Users/me/OneDrive/Roadmap.headway';
const SEED_USER = 'seed-user-00000';
const SCRIPTS = ['js/core.js', 'js/bundle.js', 'js/excel.js', 'js/export-png.js', 'js/export-pptx.js', 'js/app.js'];

// the full index.html in jsdom; tauri = a fake from fake-tauri.js, or null for
// the browser build (no desktop.js, no __TAURI__)
function boot(tauri, opts) {
  opts = opts || {};
  const vc = new VirtualConsole();
  // canvas getContext noise is a jsdom error, not ours (sendTo → forwardTo in jsdom 27)
  if (typeof vc.forwardTo === 'function') vc.forwardTo(console, { jsdomErrors: 'none' });
  else vc.sendTo(console, { omitJSDOMErrors: true });
  const dom = new JSDOM(html, { url: 'http://localhost/index.html', runScripts: 'outside-only', pretendToBeVisual: true, virtualConsole: vc });
  const { window } = dom;
  const realSetTimeout = window.setTimeout.bind(window);
  window.setTimeout = function (fn, ms) { return realSetTimeout(fn, Math.min(+ms || 0, 8)); };
  window.ExcelJS = ExcelJS;
  if (!window.Blob.prototype.arrayBuffer) { // older jsdom: exportBlob needs it
    window.Blob.prototype.arrayBuffer = function () {
      return new Promise((res, rej) => {
        const fr = new window.FileReader();
        fr.onload = () => res(fr.result);
        fr.onerror = () => rej(fr.error);
        fr.readAsArrayBuffer(this);
      });
    };
  }
  Object.keys(opts.localStorage || {}).forEach((k) => window.localStorage.setItem(k, opts.localStorage[k]));
  Object.keys(opts.sessionStorage || {}).forEach((k) => window.sessionStorage.setItem(k, opts.sessionStorage[k]));
  if (tauri) window.__TAURI__ = { core: {}, fs: tauri.fs, dialog: tauri.dialog, window: tauri.window };
  const errors = [];
  window.addEventListener('error', (e) => errors.push(e.message));
  for (const f of SCRIPTS.concat(tauri ? ['js/desktop.js'] : [])) {
    try { window.eval(fs.readFileSync(path.join(ROOT, f), 'utf8')); }
    catch (e) { failed++; console.error('  ✗ ' + f + ' threw on load: ' + e.message + '\n' + (e.stack || '').split('\n').slice(0, 3).join('\n')); }
  }
  const doc = window.document;
  return {
    window, doc, errors, RM: window.RM, RB: window.RMBundle, HD: window.HeadwayDesktop, HA: window.HeadwayApp,
    state: () => window.__headway.getState(),
    info: () => window.__headway.getInfo(),
    click: (el) => el.dispatchEvent(new window.MouseEvent('click', { bubbles: true })),
    key: (key, mods) => window.dispatchEvent(new window.KeyboardEvent('keydown', Object.assign({ key, bubbles: true }, mods || {}))),
    // the menu button toggles: dismiss any open popover first, as the smoke suite does
    menuLabels: (name) => { doc.querySelector('#popover').hidden = true; doc.querySelector('[data-menu="' + name + '"]').click(); return [...doc.querySelectorAll('#popover .menu-list button')].map((b) => b.textContent.trim()); },
    menuClick: (name, re) => {
      doc.querySelector('#popover').hidden = true;
      doc.querySelector('[data-menu="' + name + '"]').click();
      const b = [...doc.querySelectorAll('#popover .menu-list button')].find((x) => re.test(x.textContent));
      if (!b) throw new Error('no menu item ' + re);
      b.dispatchEvent(new window.MouseEvent('click', { bubbles: true }));
    },
    toasts: () => [...doc.querySelectorAll('#toasts .toast')].map((t) => t.textContent).join(' | '),
    optRow: (re) => [...doc.querySelectorAll('#popover .menu-list [data-mi]')].find((b) => re.test(b.textContent.trim())),
    optAct: (re, title) => {
      const row = [...doc.querySelectorAll('#popover .menu-list [data-mi]')].find((b) => re.test(b.textContent.trim()));
      return row && row.querySelector('.mi-act[title="' + title + '"]');
    },
  };
}
const shardWrites = (tauri, from) => tauri.log.slice(from).filter((l) => l.op === 'writeTextFile' && l.path.indexOf('/plans/') >= 0);

async function main() {
  let nextOpenDir = null, nextSavePath = null;
  const tauri = makeFakeTauri({ dialogOpen: () => nextOpenDir, dialogSave: () => nextSavePath });
  const b = boot(tauri);
  const { doc, window } = b;

  section('boot');
  ok(b.errors.length === 0, 'no window errors during boot' + (b.errors.length ? ' — ' + b.errors.join('; ') : ''));
  ['openBundleDoc', 'applyExternalEntities', 'presenceChanged', 'plansChanged', 'planShardChanged', 'resumeBundle', 'beforeClose', 'editingIds', 'flushBundle', 'closeBundleSession', 'userId']
    .forEach((m) => ok(typeof b.HA[m] === 'function', 'HeadwayApp.' + m + ' exists'));
  ['pickFolder', 'pathExists', 'exportBlob', 'readHeadway', 'writeHeadway'].forEach((m) => ok(typeof b.HD[m] === 'function', 'HeadwayDesktop.' + m + ' exists'));
  eq(b.info().docKind, 'xlsx', 'boots as a standalone document');
  ok(doc.body.classList.contains('start'), 'fresh launch shows the start page');
  ok(!window.localStorage.getItem('headway-user-v2'), 'no identity minted before it is needed');

  section('open bundle → editor shows the document');
  const fixture = b.RM.normalizeState(JSON.parse(JSON.stringify(require('./seed.fixture.js'))));
  const contents = b.RB.migrateFromState(fixture, SEED_USER, T0);
  const pid = contents.headway.plans[0].id;
  const seedLines = contents.history[SEED_USER].length;
  await b.HD.createBundle(DIR, contents);
  await b.HA.openBundleDoc(DIR);
  let info = b.info();
  eq(info.docKind, 'bundle', 'docKind is bundle');
  eq(info.bundleDir, DIR, 'bundleDir set');
  eq(info.activePlanId, pid, 'activePlanId = the first plan');
  eq(info.docSaved, true, 'fresh from disk = synced');
  eq(b.state().items.length, fixture.items.length, 'item count matches (' + fixture.items.length + ')');
  eq(b.state().phases.length, fixture.phases.length, 'phase count matches');
  eq(doc.querySelector('#docTitle').value, fixture.meta.title, 'title shown');
  ok(!doc.body.classList.contains('start'), 'editor entered');
  ok(doc.querySelectorAll('#rows .row.item').length > 50, 'rows rendered (' + doc.querySelectorAll('#rows .row.item').length + ')');
  eq(doc.querySelector('#btnSave').textContent.trim(), 'Synced ✓', 'save button reads Synced ✓');
  ok(/Default/.test(doc.querySelector('#optBtn').textContent), 'plan button shows the plan name');
  const myId = info.userId;
  ok(/^anon-[a-z0-9]{5}$/.test(myId), 'user id minted on open (no name yet): ' + myId);
  eq(b.HD.userId(), myId, 'desktop shell knows our id');
  eq(JSON.parse(window.localStorage.getItem('headway-user-v2')).id, myId, 'identity persisted');
  const rec0 = JSON.parse(window.localStorage.getItem('headway-recents-v1'));
  eq([rec0[0].path, rec0[0].kind, rec0[0].title], [DIR, 'bundle', fixture.meta.title], 'recents entry carries kind + live title');
  const uiSnap = JSON.parse(window.localStorage.getItem('headway-ui-v1'));
  eq([uiSnap.docKind, uiSnap.bundleDir, uiSnap.activePlanId], ['bundle', DIR, pid], 'ui snapshot carries the session for a reload');
  ok(tauri.watching() && tauri.watchOpts().recursive === true, 'recursive watch active');

  section('a commit flushes ONLY the changed shard + one history line');
  const bar = doc.querySelector('#rows .bar:not(.ms)[data-bar]');
  const vId = bar.getAttribute('data-bar');
  const itemPath = DIR + '/plans/' + pid + '/items/' + vId + '.json';
  b.click(doc.querySelector('#rows .row.item[data-id="' + vId + '"] .r-num'));
  let durInp = doc.querySelector('#panel input[data-f="durWeeks"]');
  ok(!!durInp, 'the panel shows the selected item');
  const dur0 = b.state().items.find((i) => i.id === vId).durDays;
  durInp.value = String(parseFloat(durInp.value) + 2);
  durInp.dispatchEvent(new window.Event('change', { bubbles: true }));
  const dur1 = b.state().items.find((i) => i.id === vId).durDays;
  ok(dur1 !== dur0, 'the commit changed the duration');
  eq(b.info().docSaved, false, 'dirty right after the commit');
  eq(doc.querySelector('#btnSave').textContent.trim(), 'Sync', 'button reads Sync while dirty');
  eq(b.info().undoLen, 1, 'one undo entry');
  let mark = tauri.log.length;
  await settle();
  let writes = shardWrites(tauri, mark);
  eq(writes.length, 1, 'exactly one shard write');
  ok(writes[0] && writes[0].path === itemPath + '.tmp', 'and it is the edited item (via tmp + rename)');
  let disk = JSON.parse(tauri.files.get(itemPath));
  eq(disk.fields.durDays, dur1, 'shard on disk carries the new duration');
  eq(disk.updatedBy, myId, 'stamped with our id');
  eq(disk.rev, 2, 'rev bumped from the seed');
  eq(disk.fieldsAt.durDays > T0, true, 'durDays stamp moved');
  eq(disk.fieldsAt.feature, T0, 'untouched field keeps its seed stamp');
  eq(b.info().docSaved, true, 'synced after the flush');
  eq(doc.querySelector('#btnSave').textContent.trim(), 'Synced ✓', 'button back to Synced ✓');
  const histPath = DIR + '/history/' + myId + '.jsonl';
  ok(tauri.files.has(histPath), 'history/<userId>.jsonl written');
  let lines = b.RB.parseHistory(tauri.files.get(histPath));
  eq(lines.length, 1, 'one history line');
  eq([lines[0].planId, lines[0].userId, lines[0].n], [pid, myId, 1], 'line carries planId + userId');
  ok(lines[0].label && lines[0].d && lines[0].d.length > 0, 'line has a label and a diff');
  eq(b.state().history, [], 'state.history stays empty in bundle mode');
  let hv = window.__headway.getHistory();
  eq(hv.length, Math.min(b.RM.HISTORY_MAX, seedLines + 1), 'getHistory merges our line with the seed user\'s file');
  eq(hv[hv.length - 1].userId, myId, 'newest merged line is ours');

  section('the name prompt stamps our own file');
  const mh = doc.querySelector('#modalHost');
  ok(!mh.hidden && /Who’s editing\?/.test(mh.textContent), 'first change asks for a name');
  const nameIn = doc.querySelector('#vhNameIn');
  nameIn.value = 'Wiring Tester';
  nameIn.dispatchEvent(new window.Event('input', { bubbles: true }));
  b.click(doc.querySelector('#vhNameSave'));
  await settle();
  lines = b.RB.parseHistory(tauri.files.get(histPath));
  eq(lines.length, 1, 'still one line');
  eq(lines[0].u, 'Wiring Tester', 'our anonymous line now carries the name (file rewritten)');
  eq(JSON.parse(window.localStorage.getItem('headway-user-v2')).id, myId, 'the id did not change with the name');
  eq(window.localStorage.getItem('headway-user-v1'), 'Wiring Tester', 'legacy name key kept in step');

  section('coalescing rewrites our tail line');
  durInp = doc.querySelector('#panel input[data-f="durWeeks"]');
  durInp.value = String(parseFloat(durInp.value) + 1);
  durInp.dispatchEvent(new window.Event('change', { bubbles: true }));
  await settle();
  lines = b.RB.parseHistory(tauri.files.get(histPath));
  eq(lines.length, 1, 'same label within the window → one line');
  eq(lines[0].n, 2, '…counting both edits');

  section('peer shard: their field lands, my unsaved edit survives, ONE render, no undo push');
  disk = JSON.parse(tauri.files.get(itemPath));
  const peer = JSON.parse(JSON.stringify(disk));
  peer.fields.notes = 'peer note'; peer.fieldsAt.notes = T2; peer.updatedAt = T2; peer.updatedBy = 'peer-zz999'; peer.rev = disk.rev + 1;
  // my edit first (unsaved: the debounce has not fired)…
  durInp = doc.querySelector('#panel input[data-f="durWeeks"]');
  durInp.value = String(parseFloat(durInp.value) + 1);
  durInp.dispatchEvent(new window.Event('change', { bubbles: true }));
  const myDur = b.state().items.find((i) => i.id === vId).durDays;
  const undoLen = b.info().undoLen, rc = b.info().renderCount;
  // …then the peer's shard arrives before it
  tauri.files.set(itemPath, JSON.stringify(peer));
  await tauri.emitPaths(itemPath);
  let it = b.state().items.find((i) => i.id === vId);
  eq(it.notes, 'peer note', 'the peer\'s field is in the live document');
  eq(it.durDays, myDur, 'my unsaved duration survives');
  eq(b.info().renderCount, rc + 1, 'exactly one render');
  eq(b.info().undoLen, undoLen, 'no undo entry pushed');
  eq(doc.querySelector('#rows .row.item[data-id="' + vId + '"]') != null, true, 'row still rendered');
  await settle();
  disk = JSON.parse(tauri.files.get(itemPath));
  eq(disk.fields.notes, 'peer note', 'my flush kept the peer\'s notes');
  eq(disk.fields.durDays, myDur, '…and wrote my duration');
  eq(disk.rev, peer.rev + 1, 'based on the peer\'s rev — no clobber');
  eq(disk.fieldsAt.notes, T2, 'the peer\'s stamp carried forward');
  // a repeat event with the same content is an echo (nothing re-applied)
  const rc2 = b.info().renderCount;
  await tauri.emitPaths(itemPath);
  eq(b.info().renderCount, rc2, 'our own write echoing back does not render');

  section('undo after a peer change reverts only my edit');
  b.key('z', { metaKey: true });
  it = b.state().items.find((i) => i.id === vId);
  eq(it.notes, 'peer note', 'undo kept the peer\'s field');
  ok(it.durDays !== myDur, 'undo reverted my duration');
  await settle();
  disk = JSON.parse(tauri.files.get(itemPath));
  eq(disk.fields.notes, 'peer note', 'the flush after undo still carries the peer\'s notes');
  eq(disk.fields.durDays, it.durDays, '…and my reverted duration');
  b.key('z', { metaKey: true, shiftKey: true }); // redo
  eq(b.state().items.find((i) => i.id === vId).durDays, myDur, 'redo restores my duration');
  eq(b.state().items.find((i) => i.id === vId).notes, 'peer note', '…with the peer\'s field intact (redo snapshot rebased too)');
  await settle();

  section('peer tombstone removes the item, and undo does not resurrect it');
  const shownIds = () => new Set([...doc.querySelectorAll('#rows .row.item')].map((r) => r.dataset.id));
  const gone = b.state().items.find((i) => i.id !== vId && shownIds().has(i.id) && (i.deps || []).length === 0);
  ok(!!doc.querySelector('#rows .row.item[data-id="' + gone.id + '"]'), 'the victim is on screen before the tombstone');
  const gonePath = DIR + '/plans/' + pid + '/items/' + gone.id + '.json';
  const tomb = b.RB.tombstone(JSON.parse(tauri.files.get(gonePath)), 'peer-zz999', T2);
  tauri.files.set(gonePath, JSON.stringify(tomb));
  const nBefore = b.state().items.length;
  await tauri.emitPaths(gonePath);
  eq(b.state().items.length, nBefore - 1, 'one item fewer');
  ok(!b.state().items.some((i) => i.id === gone.id), 'the tombstoned item is gone');
  ok(!doc.querySelector('#rows .row.item[data-id="' + gone.id + '"]'), '…and off the screen');
  b.key('z', { metaKey: true });
  ok(!b.state().items.some((i) => i.id === gone.id), 'undo does not bring the peer-deleted item back');
  b.key('z', { metaKey: true, shiftKey: true });
  await settle();
  ok(tauri.files.has(gonePath) && JSON.parse(tauri.files.get(gonePath)).deleted === true, 'the tombstone stays on disk (no resurrecting write)');

  section('a local delete writes a tombstone');
  const victim = b.state().items.find((i) => i.id !== vId && shownIds().has(i.id) && (i.deps || []).length === 0);
  b.key('Escape'); // leave the panel field so Delete reaches the selection
  b.click(doc.querySelector('#rows .row.item[data-id="' + victim.id + '"] .r-num'));
  b.key('Delete');
  b.click(doc.querySelector('#modalHost [data-m="ok"]'));
  ok(!b.state().items.some((i) => i.id === victim.id), 'deleted locally');
  mark = tauri.log.length;
  await settle();
  writes = shardWrites(tauri, mark);
  eq(writes.length, 1, 'one shard write for the delete');
  const vdisk = JSON.parse(tauri.files.get(DIR + '/plans/' + pid + '/items/' + victim.id + '.json'));
  eq([vdisk.deleted, vdisk.updatedBy], [true, myId], 'a tombstone envelope, not a file removal');
  b.key('z', { metaKey: true }); // undo the delete → re-created from the tombstone
  await settle();
  const back = JSON.parse(tauri.files.get(DIR + '/plans/' + pid + '/items/' + victim.id + '.json'));
  eq([back.deleted, back.rev], [false, vdisk.rev + 1], 'undoing the delete revives the shard with a higher rev');

  section('a title edit lands in meta.json');
  const t = doc.querySelector('#docTitle');
  b.click(t);
  t.value = 'Shared Title';
  t.dispatchEvent(new window.Event('change', { bubbles: true }));
  eq(b.state().meta.title, 'Shared Title', 'title committed');
  mark = tauri.log.length;
  await settle();
  writes = shardWrites(tauri, mark);
  eq(writes.map((w) => w.path), [DIR + '/plans/' + pid + '/meta.json.tmp'], 'only meta.json written');
  eq(JSON.parse(tauri.files.get(DIR + '/plans/' + pid + '/meta.json')).fields.meta.title, 'Shared Title', 'meta shard carries the title');
  eq(b.HD.bundleDir(), DIR, 'the folder is not renamed');
  eq(b.HD.currentPath(), null, 'no xlsx path adopted');

  section('presence + editingIds');
  b.HA.presenceChanged('peer-zz999', { name: 'Peer', planId: pid, editing: [vId], ts: 1 });
  eq(Object.keys(b.info().peers), ['peer-zz999'], 'peer stored');
  b.HA.presenceChanged('peer-zz999', null);
  eq(Object.keys(b.info().peers), [], 'peer dropped on null');
  b.click(doc.querySelector('#rows .row.item[data-id="' + vId + '"] .r-num'));
  eq(b.HA.editingIds(), [vId], 'editingIds = the selection');

  section('plans: New writes plans/<id>/ + headway.json and switches; switch back writes no shard');
  b.click(doc.querySelector('#optBtn'));
  b.click(b.optRow(/New plan/));
  const pn = doc.querySelector('#modalHost #optNameIn');
  ok(!!pn, 'New plan asks for a name');
  pn.value = 'Plan B';
  b.click(doc.querySelector('#modalHost #optNameOk'));
  ok(await until(() => b.info().activePlanId !== pid), 'switched to the new plan — toasts: ' + b.toasts());
  await settle();
  let hw = JSON.parse(tauri.files.get(DIR + '/headway.json'));
  eq(hw.plans.length, 2, 'headway.json lists two plans');
  const pB = hw.plans.find((p) => p.name === 'Plan B');
  ok(!!pB && pB.id === b.info().activePlanId, 'the new entry is the active plan');
  ok(tauri.files.has(DIR + '/plans/' + pB.id + '/meta.json'), 'plans/<id>/meta.json written');
  const bItems = [...tauri.files.keys()].filter((f) => f.indexOf(DIR + '/plans/' + pB.id + '/items/') === 0).length;
  eq(bItems, b.state().items.length, 'one item shard per item in the new plan');
  eq(b.state().meta.title, 'Shared Title', 'the copy carries the document');
  ok(/Plan B/.test(doc.querySelector('#optBtn').textContent), 'button shows Plan B');
  eq(b.info().undoLen, 0, 'undo cleared across plans');
  // an edit in Plan B touches only Plan B's folder
  b.click(doc.querySelector('#rows .row.item[data-id="' + vId + '"] .r-num'));
  durInp = doc.querySelector('#panel input[data-f="durWeeks"]');
  durInp.value = String(parseFloat(durInp.value) + 1);
  durInp.dispatchEvent(new window.Event('change', { bubbles: true }));
  const bDur = b.state().items.find((i) => i.id === vId).durDays;
  mark = tauri.log.length;
  await settle();
  writes = shardWrites(tauri, mark);
  eq(writes.map((w) => w.path), [DIR + '/plans/' + pB.id + '/items/' + vId + '.json.tmp'], 'only Plan B\'s shard written');
  lines = b.RB.parseHistory(tauri.files.get(histPath));
  eq(lines[lines.length - 1].planId, pB.id, 'history line carries Plan B\'s id');
  hv = window.__headway.getHistory();
  ok(hv.every((l) => l.planId === pB.id), 'history view is filtered to the active plan');
  // switch back to Default
  mark = tauri.log.length;
  b.click(doc.querySelector('#optBtn'));
  b.click(b.optRow(/^Default$/));
  ok(await until(() => b.info().activePlanId === pid), 'switched back to Default');
  eq(shardWrites(tauri, mark).length, 0, 'a switch writes no shard');
  eq(b.state().items.find((i) => i.id === vId).durDays, myDur, 'Default kept its own duration');
  ok(b.state().items.find((i) => i.id === vId).durDays !== bDur, '(Plan B\'s edit stayed in Plan B)');
  eq(b.info().undoLen, 0, 'undo cleared');
  // compare loads the other plan from disk
  b.click(doc.querySelector('#optBtn'));
  b.click(b.optAct(/Plan B/, 'Compare'));
  ok(await until(() => doc.querySelectorAll('#rows .bar.cmp').length > 0), 'Compare ghosts Plan B\'s differing bars');
  ok(!!doc.querySelector('#cmpPill') && /Plan B/.test(doc.querySelector('#cmpPill').textContent), 'pill names the compared plan');
  b.click(doc.querySelector('#cmpPill [data-cmpx]'));
  // rename + delete (tombstone; files stay)
  b.click(doc.querySelector('#optBtn'));
  b.click(b.optAct(/Plan B/, 'Rename'));
  doc.querySelector('#modalHost #optNameIn').value = 'Plan C';
  b.click(doc.querySelector('#modalHost #optNameOk'));
  ok(await until(() => /Plan C/.test(JSON.parse(tauri.files.get(DIR + '/headway.json')).plans.find((p) => p.id === pB.id).name)), 'rename lands in headway.json');
  b.click(doc.querySelector('#optBtn'));
  b.click(b.optAct(/Plan C/, 'Delete'));
  ok(!!doc.querySelector('#modalHost [data-m="ok"]'), 'delete asks for confirmation');
  b.click(doc.querySelector('#modalHost [data-m="ok"]'));
  ok(await until(() => JSON.parse(tauri.files.get(DIR + '/headway.json')).plans.find((p) => p.id === pB.id).deleted === true), 'delete = tombstone entry');
  ok(tauri.files.has(DIR + '/plans/' + pB.id + '/meta.json'), 'the plan\'s files are kept');
  eq(b.info().plans.filter((p) => !p.deleted).length, 1, 'one live plan');
  // a peer adds a plan: plansChanged via the watcher
  hw = JSON.parse(tauri.files.get(DIR + '/headway.json'));
  hw.plans.push(b.RB.newPlanEntry('plan-peer', 'Peer plan', T2));
  tauri.files.set(DIR + '/headway.json', JSON.stringify(hw));
  await tauri.emitPaths(DIR + '/headway.json');
  await settle(1);
  ok(b.info().plans.some((p) => p.id === 'plan-peer'), 'a peer\'s plan entry arrives through the watcher');
  ok(/2 open plans/.test(doc.querySelector('#optBtn').title), 'plan button counts it');

  section('Export .xlsx… writes a standalone workbook and adopts nothing');
  nextSavePath = 'C:/tmp/export.xlsx';
  tauri.dirs.add('C:/tmp'); // the dialog only returns folders that exist
  let labels = b.menuLabels('file');
  ok(labels.some((l) => /^Export \.xlsx…/.test(l)), 'File menu offers Export .xlsx… in bundle mode');
  ok(!labels.some((l) => /^Save/.test(l)), '…and no Save / Save as');
  ok(!labels.some((l) => /Auto save/.test(l)), 'Auto save hidden');
  ok(!labels.some((l) => /Convert to shared/.test(l)), 'Convert hidden while already shared');
  b.menuClick('file', /Export \.xlsx…/);
  ok(await until(() => tauri.files.has('C:/tmp/export.xlsx')), 'workbook written where the dialog said — toasts: ' + b.toasts());
  const xbytes = tauri.files.get('C:/tmp/export.xlsx');
  ok(xbytes && xbytes.length > 5000, 'a real workbook (' + (xbytes && xbytes.length) + ' bytes)');
  const xbuf = Buffer.from(xbytes); // node realm for ExcelJS (a jsdom ArrayBuffer fails its instanceof)
  const imp = await window.RMExcel.importWorkbook(xbuf);
  eq(imp.source, 'tool', 'the export re-imports through the tool sheet');
  ok(!('docId' in imp.state) && !('bundle' in imp.state) && !('planId' in imp.state), 'exported state carries no bundle markers');
  eq(imp.state.items.length, b.state().items.length, 'same items');
  ok(imp.ui && imp.ui.docKind === undefined && imp.ui.bundleDir === undefined, 'the workbook\'s ui snapshot has no folder link');
  eq(b.HD.currentPath(), null, 'currentPath untouched by the export');
  eq([b.info().docKind, b.HD.bundleDir()], ['bundle', DIR], 'bundle session untouched');

  section('menus with __TAURI__: the shared-roadmap items');
  ok(labels.some((l) => /New shared roadmap…/.test(l)) && labels.some((l) => /Open shared roadmap…/.test(l)), 'File: New/Open shared roadmap present');
  const mac = b.HA.menuItems('macApp').map((m) => m.label || '');
  ok(mac.some((l) => /New shared roadmap…/.test(l)) && mac.some((l) => /Open shared roadmap…/.test(l)) && mac.some((l) => /Export \.xlsx…/.test(l)), 'macApp list mirrors them');
  doc.querySelector('#popover').hidden = true;

  section('open an .xlsx → standalone; Convert to shared folder… → a new bundle');
  await b.HA.loadBuffer(xbuf, 'Exported.xlsx');
  await settle();
  eq(b.info().docKind, 'xlsx', 'an .xlsx is a standalone document');
  eq(b.HD.bundleDir(), null, 'the bundle session closed');
  eq(b.state().meta.title, 'Exported', 'title from the file name');
  labels = b.menuLabels('file');
  ok(labels.some((l) => /Convert to shared folder…/.test(l)), 'File menu offers Convert for an .xlsx');
  ok(labels.some((l) => /^Save/.test(l)) && !labels.some((l) => /Export \.xlsx…/.test(l)), 'Save is Save again');
  nextOpenDir = 'C:/Users/me/OneDrive';
  b.menuClick('file', /Convert to shared folder…/);
  const CDIR = 'C:/Users/me/OneDrive/Exported.headway';
  ok(await until(() => b.info().docKind === 'bundle' && b.info().bundleDir === CDIR), 'converted and opened as a bundle — toasts: ' + b.toasts());
  ok(tauri.files.has(CDIR + '/headway.json'), 'headway.json written');
  const cPid = b.info().activePlanId;
  eq([...tauri.files.keys()].filter((f) => f.indexOf(CDIR + '/plans/' + cPid + '/items/') === 0).length, b.state().items.length, 'one shard per item');
  ok(tauri.files.has(CDIR + '/history/' + myId + '.jsonl'), 'legacy history landed in the converting user\'s file');
  eq(b.state().meta.title, 'Exported', 'same document');

  section('Convert refuses a folder that already holds a bundle');
  await b.HA.loadBuffer(xbuf, 'Exported.xlsx');
  await settle();
  eq(b.info().docKind, 'xlsx', 'back on the standalone document');
  const filesBefore = [...tauri.files.keys()].sort().join('|');
  // toasts auto-dismiss within milliseconds under the compressed timers, so catch them as they are added
  const seenToasts = [];
  const mo = new window.MutationObserver((muts) => muts.forEach((m) => m.addedNodes.forEach((n) => { if (n.textContent) seenToasts.push(n.textContent); })));
  mo.observe(doc.querySelector('#toasts'), { childList: true });
  nextOpenDir = 'C:/Users/me/OneDrive'; // same parent → Exported.headway already exists
  b.menuClick('file', /Convert to shared folder…/);
  ok(await until(() => seenToasts.length > 0), 'a toast appears');
  mo.disconnect();
  ok(seenToasts.some((t) => /already exists/.test(t)), 'it says the folder already exists: ' + seenToasts.join(' | '));
  eq(b.info().docKind, 'xlsx', 'still a standalone document — nothing was opened');
  eq([...tauri.files.keys()].sort().join('|'), filesBefore, 'not a single file written or changed');

  section('recents: kind → folder icon; start page has Open shared roadmap…');
  b.HA.noteRecent('C:/x/Plain.xlsx');
  b.HA.renderStartPage();
  const rows = [...doc.querySelectorAll('#startBody [data-sp-open]')];
  const bRow = rows.find((r) => r.dataset.spOpen === CDIR);
  const xRow = rows.find((r) => r.dataset.spOpen === 'C:/x/Plain.xlsx');
  ok(bRow && bRow.querySelector('i[data-lucide="folder-open"]'), 'a bundle recent shows the folder icon');
  ok(xRow && xRow.querySelector('i[data-lucide="file-spreadsheet"]'), 'an xlsx recent shows the spreadsheet icon');
  ok(!!doc.querySelector('#startBody [data-sp-openbundle]'), 'Open shared roadmap… button present (desktop)');

  section('New shared roadmap… from the menu');
  nextOpenDir = 'C:/Users/me/OneDrive';
  b.menuClick('file', /New shared roadmap…/);
  const nn = doc.querySelector('#modalHost #optNameIn');
  ok(!!nn, 'asks for a name');
  nn.value = 'Fresh';
  b.click(doc.querySelector('#modalHost #optNameOk'));
  const FDIR = 'C:/Users/me/OneDrive/Fresh.headway';
  ok(await until(() => b.info().bundleDir === FDIR), 'created and opened — toasts: ' + b.toasts());
  eq(b.state().meta.title, 'Fresh', 'blank roadmap titled from the prompt');
  eq(b.state().items.length, 0, 'empty');
  ok(tauri.files.has(FDIR + '/headway.json'), 'headway.json written');

  section('resume after a reload (ui snapshot) + beforeClose');
  const ls = {};
  ['headway-v1', 'headway-ui-v1', 'headway-user-v2', 'headway-user-v1', 'headway-recents-v1'].forEach((k) => { const v = window.localStorage.getItem(k); if (v != null) ls[k] = v; });
  const b2 = boot(tauri, { localStorage: ls, sessionStorage: { 'headway-in-editor': '1' } });
  ok(b2.errors.length === 0, 'reload boots clean');
  ok(await until(() => b2.info().docKind === 'bundle'), 'the folder is re-linked after the reload — toasts: ' + b2.toasts());
  eq(b2.info().bundleDir, FDIR, 'same folder');
  eq(b2.info().userId, myId, 'same identity');
  ok(!b2.doc.body.classList.contains('start'), 'editor active (mid-session reload)');
  await b2.HD.writePresence(FDIR, myId, { name: 'Wiring Tester', planId: b2.info().activePlanId, editing: [], ts: 1 });
  await b2.HA.beforeClose();
  ok(!tauri.files.has(FDIR + '/presence/' + myId + '.json'), 'beforeClose drops our presence file');
  eq(b2.HD.bundleDir(), null, '…and leaves the folder');
  eq(b2.info().docKind, 'xlsx', 'app-side session closed');

  section('web build: no shared-roadmap surfaces, no resume');
  const w = boot(null, { localStorage: ls, sessionStorage: { 'headway-in-editor': '1' } });
  ok(w.errors.length === 0, 'browser build boots clean');
  ok(!w.HD, 'no HeadwayDesktop');
  eq(w.info().docKind, 'xlsx', 'a browser never re-links a folder');
  const wl = w.HA.menuItems('file').map((m) => m.label || '');
  ok(!wl.some((l) => /shared roadmap|Convert to shared|Export \.xlsx/.test(l)), 'File menu has none of the bundle items');
  ok(wl.some((l) => l === 'Save'), 'Save stays Save');
  w.HA.renderStartPage();
  ok(!w.doc.querySelector('#startBody [data-sp-openbundle]'), 'no Open shared roadmap… button');
  w.HA.applyExternalEntities([{ kind: 'items', id: 'x', env: { id: 'x', rev: 1, fields: {}, fieldsAt: {} } }]);
  ok(true, 'applyExternalEntities is a no-op outside a bundle');

  await fixes();
  await presence();
  await importFlow();

  console.log('\n' + passed + ' passed, ' + failed + ' failed');
  process.exit(failed ? 1 : 0);
}

// ---- review fixes F1–F7 / N1–N2: each group boots a fresh app over a fresh
// fake fs so a race in one cannot bleed into the next
const FIXER = { name: 'Fixer', id: 'fixer-00001' };
async function openFresh(name, tauriOpts) {
  const tauri = makeFakeTauri(tauriOpts || {});
  // a stored identity: no name prompt in the way of the panel edits
  const b = boot(tauri, { localStorage: { 'headway-user-v2': JSON.stringify(FIXER), 'headway-user-v1': FIXER.name } });
  const fixture = b.RM.normalizeState(JSON.parse(JSON.stringify(require('./seed.fixture.js'))));
  const contents = b.RB.migrateFromState(fixture, SEED_USER, T0);
  const dir = 'C:/Users/me/OneDrive/' + name + '.headway';
  await b.HD.createBundle(dir, contents);
  await b.HA.openBundleDoc(dir);
  const pid = contents.headway.plans[0].id;
  const vId = b.doc.querySelector('#rows .bar:not(.ms)[data-bar]').getAttribute('data-bar');
  const S = {
    tauri, b, dir, pid, vId, fixture,
    path: (id, plan) => dir + '/plans/' + (plan || pid) + '/items/' + id + '.json',
    disk: (id, plan) => JSON.parse(tauri.files.get(S.path(id, plan))),
    item: (id) => b.state().items.find((i) => i.id === (id || vId)),
    select: (id) => b.click(b.doc.querySelector('#rows .row.item[data-id="' + (id || vId) + '"] .r-num')),
    // drive a commit through the panel's change handler (a field the panel
    // does not render gets a throwaway input — the handler reads data-f + value)
    set: (f, v) => {
      let el = b.doc.querySelector('#panel [data-f="' + f + '"]');
      if (!el) { el = b.doc.createElement('input'); el.dataset.f = f; b.doc.querySelector('#panel').appendChild(el); }
      el.value = v;
      el.dispatchEvent(new b.window.Event('change', { bubbles: true }));
    },
    // a peer's version of the shard now on disk: one field changed, stamped `at`
    peerEnv: (id, field, value, at) => {
      const env = S.disk(id);
      env.fields[field] = value; env.fieldsAt[field] = at; env.updatedAt = at; env.updatedBy = 'peer-zz999'; env.rev = env.rev + 1;
      return env;
    },
    emitPeer: async (id, field, value, at) => {
      const env = S.peerEnv(id, field, value, at);
      tauri.files.set(S.path(id), JSON.stringify(env));
      await tauri.emitPaths(S.path(id));
      return env;
    },
    renames: (id) => tauri.log.filter((l) => l.op === 'rename' && l.to === S.path(id)).length,
  };
  return S;
}
// hold one fs op for matching paths until release(); other calls pass through
function stall(tauri, op, match) {
  const orig = tauri.fs[op];
  let release;
  const gate = new Promise((r) => { release = r; });
  const hit = [];
  tauri.fs[op] = function () {
    const args = [...arguments];
    const p = String(args[0]).replace(/\\/g, '/');
    if (!match(p)) return orig.apply(null, args);
    hit.push(p);
    return gate.then(() => orig.apply(null, args));
  };
  return { hit, release: () => release(), restore: () => { tauri.fs[op] = orig; } };
}
// the app reads the clock through window.Date: shift it to make an edit "old"
function shiftClock(win, ms) {
  const Real = win.Date;
  function Shifted(...a) { return a.length ? new Real(...a) : new Real(Real.now() + ms); }
  Shifted.prototype = Real.prototype;
  Shifted.now = () => Real.now() + ms;
  Shifted.parse = Real.parse; Shifted.UTC = Real.UTC;
  win.Date = Shifted;
  return () => { win.Date = Real; };
}
const isoIn = (ms) => new Date(Date.now() + ms).toISOString();

async function fixes() {
  section('F2/T1: peer envelope applied WHILE our flush is in flight → both fields land, one follow-up flush');
  {
    const S = await openFresh('T1');
    const key = 'items/' + S.vId;
    S.select();
    const r0 = S.renames(S.vId);
    const st = stall(S.tauri, 'rename', (p) => p.indexOf('/items/' + S.vId + '.json') >= 0);
    S.set('feature', 'Local feature');
    ok(await until(() => st.hit.length > 0), 'flush in flight (rename stalled)');
    eq(S.b.info().flushing, true, 'flushing flag up');
    const peer = await S.emitPeer(S.vId, 'notes', 'peer note', isoIn(60000));
    eq(S.item().notes, 'peer note', 'peer field applied mid-flight');
    eq(S.item().feature, 'Local feature', 'local field kept mid-flight');
    eq(S.b.info().lastEnv[key].rev, peer.rev, 'baseline moved to the peer envelope');
    st.release();
    ok(await until(() => S.b.info().docSaved && !S.b.info().flushing), 'flush chain settled — toasts: ' + S.b.toasts());
    st.restore();
    const disk = S.disk(S.vId);
    eq([disk.fields.feature, disk.fields.notes], ['Local feature', 'peer note'], 'disk shard carries BOTH the local and the peer field');
    eq(disk.rev, peer.rev + 1, 'written on top of the peer rev');
    eq(disk.fieldsAt.notes, peer.fieldsAt.notes, 'peer stamp carried, not re-stamped');
    eq(S.b.RB.canonicalize(S.b.info().lastEnv[key]), S.b.RB.canonicalize(disk), 'lastEnv baseline == the newer (disk) env');
    eq(S.renames(S.vId) - r0, 2, 'exactly one follow-up flush (two shard renames in total)');
    ok(S.b.errors.length === 0, 'no window errors');
  }

  section('F1/T2a: edit during an in-flight flush, then switchPlan → both edits on disk under the ORIGINAL plan');
  {
    const S = await openFresh('T2a');
    // a second plan straight on disk (files + list entry), registered via the watcher hook
    const planB = 'plan-b';
    const envsB = S.b.RB.wrapState(S.b.RM.clone(S.b.state()), {}, 'seed', T0);
    const chB = [{ kind: 'meta', id: 'meta', env: S.b.RB.wrapMeta(S.b.state(), null, 'seed', T0), baseRev: 0 }];
    S.b.RB.KINDS.forEach((k) => envsB[k].forEach((env) => chB.push({ kind: k, id: env.id, env, baseRev: 0 })));
    await S.b.HD.flushShards(S.dir, planB, chB);
    const hw = await S.b.HD.writeHeadway(S.dir, { plans: [S.b.RB.newPlanEntry(planB, 'Plan B', T0)] });
    S.b.HA.plansChanged(hw);
    eq(S.b.info().plans.filter((p) => !p.deleted).length, 2, 'two live plans');
    S.select();
    const r0 = S.renames(S.vId);
    const st = stall(S.tauri, 'rename', (p) => p.indexOf('/items/' + S.vId + '.json') >= 0);
    S.set('feature', 'A1');
    ok(await until(() => st.hit.length > 0), 'first flush in flight');
    S.set('notes', 'N2'); // lands while the first flush is stalled → queued
    S.b.click(S.b.doc.querySelector('#optBtn'));
    S.b.click(S.b.optRow(/Plan B/)); // switchPlan awaits the whole chain
    await tick(20);
    eq(S.b.info().activePlanId, S.pid, 'still on the original plan while its flush is pending');
    st.release();
    ok(await until(() => S.b.info().activePlanId === planB), 'switched to Plan B — toasts: ' + S.b.toasts());
    st.restore();
    const d = S.disk(S.vId);
    eq([d.fields.feature, d.fields.notes], ['A1', 'N2'], 'both edits on disk under the original plan');
    eq(S.renames(S.vId) - r0, 2, 'the queued re-flush ran (two renames)');
    const dB = S.disk(S.vId, planB);
    ok(dB.fields.feature !== 'A1' && dB.fields.notes !== 'N2', 'Plan B untouched');
    eq(S.item().feature === 'A1', false, 'the document is Plan B\'s');
  }

  section('F1+F7/T2b+T8: edit during an in-flight flush, then open an .xlsx (openPath) → edit on disk, xlsx watcher alive');
  {
    const S = await openFresh('T2b');
    const blob = await S.b.window.RMExcel.exportWorkbook(S.b.RB.exportableState(S.b.state()), {});
    const ab = await blob.arrayBuffer();
    S.tauri.dirs.add('C:/tmp');
    S.tauri.files.set('C:/tmp/Solo.xlsx', Uint8Array.from(new Uint8Array(ab))); // own buffer: bytes.buffer is the file
    S.select();
    const st = stall(S.tauri, 'rename', (p) => p.indexOf('/items/' + S.vId + '.json') >= 0);
    S.set('feature', 'A1');
    ok(await until(() => st.hit.length > 0), 'flush in flight');
    S.set('notes', 'N2');
    const opening = S.b.HD.openPath('C:/tmp/Solo.xlsx');
    // give the workbook import every chance to finish first: the fix holds it
    // behind the stalled flush, the bug let it adopt the path (and its watcher) early
    const early = await until(() => S.b.HD.currentPath() != null, 1500);
    eq(early, false, 'the xlsx open waits for the bundle flush');
    eq(S.b.info().docKind, 'bundle', 'still the bundle document meanwhile');
    st.release();
    await opening;
    st.restore();
    await settle();
    eq(S.b.info().docKind, 'xlsx', 'xlsx open');
    eq(S.b.HD.currentPath(), 'C:/tmp/Solo.xlsx', 'path adopted');
    eq(S.b.HD.bundleDir(), null, 'folder left');
    const d = S.disk(S.vId);
    eq([d.fields.feature, d.fields.notes], ['A1', 'N2'], 'both edits on disk under the bundle plan');
    ok(S.tauri.watching(), 'the xlsx watcher is active after bundle → xlsx');
    ok(!(S.tauri.watchOpts() || {}).recursive, '…non-recursive (a file watch on the parent)');
    const lastWatch = S.tauri.log.filter((l) => l.op === 'watch').pop();
    eq(lastWatch && lastWatch.path, 'C:/tmp', '…on the file\'s parent folder');
    const ops = S.tauri.log.map((l) => l.op);
    ok(ops.lastIndexOf('unwatch') < ops.lastIndexOf('watch'), 'no unwatch after the xlsx watch (leave ran first)');
    eq(S.b.state().meta.title, 'Solo', 'the workbook is the document');
  }

  section('F1/T2c: edit during an in-flight flush, then beforeClose → nothing dropped');
  {
    const S = await openFresh('T2c');
    S.select();
    const st = stall(S.tauri, 'rename', (p) => p.indexOf('/items/' + S.vId + '.json') >= 0);
    S.set('feature', 'A1');
    ok(await until(() => st.hit.length > 0), 'flush in flight');
    S.set('notes', 'N2');
    await S.b.HD.writePresence(S.dir, FIXER.id, { name: FIXER.name, planId: S.pid, editing: [], ts: 1 });
    const closing = S.b.HA.beforeClose();
    await tick(20);
    eq(S.b.info().docKind, 'bundle', 'close waits for the flush');
    st.release();
    await closing;
    st.restore();
    const d = S.disk(S.vId);
    eq([d.fields.feature, d.fields.notes], ['A1', 'N2'], 'both edits on disk');
    eq(S.b.info().docKind, 'xlsx', 'session closed');
    ok(!S.tauri.files.has(S.dir + '/presence/' + FIXER.id + '.json'), 'presence dropped after the flush');
    eq(S.b.HD.bundleDir(), null, 'folder left');
  }

  section('F6/T3: a coalesce rewrite before readHistory resolves keeps the file\'s past lines');
  {
    const S = await openFresh('T3');
    S.select();
    S.set('deadline', '2026-10-01');
    await settle();
    const hp = S.dir + '/history/' + FIXER.id + '.jsonl';
    const before = S.b.RB.parseHistory(S.tauri.files.get(hp));
    eq(before.length, 1, 'one line on disk from the first session');
    // second session on the same folder: our history file read is held
    const b2 = boot(S.tauri, { localStorage: { 'headway-user-v2': JSON.stringify(FIXER), 'headway-user-v1': FIXER.name } });
    const st = stall(S.tauri, 'readTextFile', (p) => p === hp);
    await b2.HA.openBundleDoc(S.dir);
    ok(await until(() => st.hit.length > 0), 'history read in flight (held)');
    b2.click(b2.doc.querySelector('#rows .row.item[data-id="' + S.vId + '"] .r-num'));
    const setB2 = (f, v) => {
      let el = b2.doc.querySelector('#panel [data-f="' + f + '"]');
      if (!el) { el = b2.doc.createElement('input'); el.dataset.f = f; b2.doc.querySelector('#panel').appendChild(el); }
      el.value = v; el.dispatchEvent(new b2.window.Event('change', { bubbles: true }));
    };
    setB2('notes', 'n1');
    setB2('notes', 'n2'); // same label inside the window → coalesce → rewrite op
    await settle();
    eq(b2.info().pendingHistory, 0, 'history ops taken by the flush');
    eq(S.b.RB.parseHistory(S.tauri.files.get(hp)).length, 1, 'nothing rewritten while the read is pending');
    st.release();
    ok(await until(() => S.b.RB.parseHistory(S.tauri.files.get(hp)).length === 2), 'file rewritten once the read landed');
    st.restore();
    const after = S.b.RB.parseHistory(S.tauri.files.get(hp));
    eq(after[0], before[0], 'the pre-existing line survives, first');
    eq([after[1].label, after[1].n], ['notes', 2], 'the coalesced new line follows');
    const hv = b2.window.__headway.getHistory();
    ok(hv.some((l) => l.t === before[0].t && l.userId === FIXER.id), 'history view (cache) includes the past line');
  }

  section('F5: coalescing never merges across plans');
  {
    const S = await openFresh('F5');
    S.select();
    S.set('notes', 'plan A note');
    await settle();
    // the same label right away, but on another plan (simulated: the line's planId differs)
    const hp = S.dir + '/history/' + FIXER.id + '.jsonl';
    const lines0 = S.b.RB.parseHistory(S.tauri.files.get(hp));
    eq(lines0.length, 1, 'one line');
    // switch plan by creating Plan B via the UI, then the same edit there
    S.b.click(S.b.doc.querySelector('#optBtn'));
    S.b.click(S.b.optRow(/New plan/));
    S.b.doc.querySelector('#modalHost #optNameIn').value = 'Plan B';
    S.b.click(S.b.doc.querySelector('#modalHost #optNameOk'));
    ok(await until(() => S.b.info().activePlanId !== S.pid), 'on Plan B — toasts: ' + S.b.toasts());
    S.select();
    S.set('notes', 'plan B note');
    await settle();
    const lines1 = S.b.RB.parseHistory(S.tauri.files.get(hp));
    eq(lines1.length, 2, 'two lines: the plan-B edit did not coalesce into plan A\'s');
    eq([lines1[0].planId, lines1[1].planId], [S.pid, S.b.info().activePlanId], 'each carries its own plan');
  }

  section('F4/T4: peer changes one field → undo ×2 reverts both earlier local edits in order; redo restores');
  {
    const S = await openFresh('T4');
    const it0 = JSON.parse(JSON.stringify(S.item()));
    S.select();
    S.set('feature', 'F1');
    S.set('notes', 'N1');
    await settle();
    eq(S.b.info().undoLen, 2, 'two undo entries');
    eq([S.disk(S.vId).fields.feature, S.disk(S.vId).fields.notes], ['F1', 'N1'], 'both flushed');
    await S.emitPeer(S.vId, 'description', 'peer desc', isoIn(60000));
    eq(S.item().description, 'peer desc', 'peer field applied');
    S.b.key('z', { metaKey: true });
    eq([S.item().feature, S.item().notes, S.item().description], ['F1', it0.notes, 'peer desc'], 'undo 1: notes reverted, feature + peer field stay');
    S.b.key('z', { metaKey: true });
    eq([S.item().feature, S.item().notes, S.item().description], [it0.feature, it0.notes, 'peer desc'], 'undo 2: feature reverted, peer field stays');
    S.b.key('z', { metaKey: true, shiftKey: true });
    S.b.key('z', { metaKey: true, shiftKey: true });
    eq([S.item().feature, S.item().notes, S.item().description], ['F1', 'N1', 'peer desc'], 'redo ×2 restores both, peer field intact');
    await settle();
    const d = S.disk(S.vId);
    eq([d.fields.feature, d.fields.notes, d.fields.description], ['F1', 'N1', 'peer desc'], 'disk agrees');
  }

  section('T5: a denied rename fails the flush, keeps history pending; Sync retries once allowed');
  {
    const S = await openFresh('T5');
    S.select();
    const seen = [];
    const mo = new S.b.window.MutationObserver((muts) => muts.forEach((m) => m.addedNodes.forEach((n) => { if (n.textContent) seen.push(n.textContent); })));
    mo.observe(S.b.doc.querySelector('#toasts'), { childList: true });
    S.tauri.deny.add('rename');
    S.set('feature', 'Denied');
    ok(await until(() => seen.some((t) => /Sync failed/.test(t))), 'Sync failed toast: ' + seen.join(' | '));
    await settle();
    eq(S.b.info().docSaved, false, 'still dirty');
    eq(S.b.info().pendingHistory, 1, 'history line retained');
    eq(S.disk(S.vId).fields.feature !== 'Denied', true, 'shard not written');
    ok(!S.tauri.files.has(S.dir + '/history/' + FIXER.id + '.jsonl'), 'no history line written');
    eq(S.b.doc.querySelector('#btnSave').textContent.trim(), 'Sync', 'button offers Sync');
    S.tauri.deny.delete('rename');
    S.b.click(S.b.doc.querySelector('#btnSave'));
    ok(await until(() => S.b.info().docSaved), 'synced after the retry — toasts: ' + seen.join(' | '));
    mo.disconnect();
    eq(S.disk(S.vId).fields.feature, 'Denied', 'shard landed');
    const lines = S.b.RB.parseHistory(S.tauri.files.get(S.dir + '/history/' + FIXER.id + '.jsonl'));
    eq(lines.length, 1, 'history line landed');
    eq(S.b.info().pendingHistory, 0, 'nothing pending');
  }

  section('T6a: merged fold-back — the peer moved the shard before our flush read it');
  {
    const S = await openFresh('T6a');
    S.select();
    // the peer's shard is on disk but its watch event has not reached us
    const peer = S.peerEnv(S.vId, 'notes', 'peer note', isoIn(60000));
    S.tauri.files.set(S.path(S.vId), JSON.stringify(peer));
    S.set('feature', 'F6');
    await settle();
    eq(S.item().notes, 'peer note', 'the peer\'s field folded back into the document');
    eq(S.item().feature, 'F6', 'local field kept');
    const d = S.disk(S.vId);
    eq([d.fields.feature, d.fields.notes], ['F6', 'peer note'], 'disk holds the merge');
    eq(S.b.info().docSaved, true, 'synced');
    eq(S.b.RB.canonicalize(S.b.info().lastEnv['items/' + S.vId]), S.b.RB.canonicalize(d), 'baseline == disk');
    // the fold is a peer change for undo purposes: undo reverts only the local field
    S.b.key('z', { metaKey: true });
    eq([S.item().feature === 'F6', S.item().notes], [false, 'peer note'], 'undo reverts F6, keeps the peer note');
  }

  section('N1/T6b: peer envelope for the dragged item waits for pointerup — and for pointercancel');
  {
    const S = await openFresh('T6b');
    const bar = () => S.b.doc.querySelector('#rows .bar[data-bar="' + S.vId + '"]');
    const down = () => bar().dispatchEvent(new S.b.window.MouseEvent('pointerdown', { bubbles: true, cancelable: true, button: 0, clientX: 50, clientY: 50 }));
    down();
    ok(S.b.HA.editingIds().indexOf(S.vId) !== -1, 'drag on the item started');
    await S.emitPeer(S.vId, 'notes', 'held', isoIn(60000));
    eq(S.item().notes !== 'held', true, 'not applied under the pointer');
    eq(S.b.info().deferred, 1, 'held in deferredExternal');
    S.b.window.dispatchEvent(new S.b.window.MouseEvent('pointerup', { bubbles: true, clientX: 50, clientY: 50 }));
    ok(await until(() => S.item().notes === 'held'), 'applied after pointerup');
    eq(S.b.info().deferred, 0, 'queue drained');
    down();
    ok(S.b.HA.editingIds().indexOf(S.vId) !== -1, 'second drag started');
    await S.emitPeer(S.vId, 'notes', 'held2', isoIn(120000));
    eq(S.b.info().deferred, 1, 'held again');
    S.b.window.dispatchEvent(new S.b.window.Event('pointercancel', { bubbles: true }));
    ok(await until(() => S.item().notes === 'held2'), 'applied after pointercancel');
    eq(S.b.HA.editingIds().indexOf(S.vId), -1, 'the cancelled drag is over');
    ok(S.b.errors.length === 0, 'no window errors');
  }

  section('F3/T7: a local edit keeps its commit-time stamp — an older offline edit loses to a newer peer edit');
  {
    const S = await openFresh('T7');
    S.select();
    const unshift = shiftClock(S.b.window, -10 * 60000); // this machine's clock reads 10 min ago
    S.set('notes', 'stale local');
    unshift();
    eq(S.b.info().docSaved, false, 'dirty, flush not yet started');
    const la = S.b.info().localAt['items/' + S.vId];
    ok(la && la.notes && la.notes < isoIn(-9 * 60000), 'notes stamped at its (shifted) commit time: ' + (la && la.notes));
    // the peer edited notes 5 minutes ago — newer than our 10-minutes-ago edit — and it arrives before our flush
    const peer = await S.emitPeer(S.vId, 'notes', 'peer newer', isoIn(-5 * 60000));
    eq(S.item().notes, 'peer newer', 'the peer\'s newer value wins in the document');
    await settle();
    const d = S.disk(S.vId);
    eq(d.fields.notes, 'peer newer', 'disk notes = the peer\'s (ours was older)');
    eq(d.fieldsAt.notes, peer.fieldsAt.notes, 'peer stamp intact — nothing re-stamped');
    eq(S.b.info().docSaved, true, 'synced (nothing of ours left to send)');
    // and the reverse: our edit newer than the peer's → ours wins and carries its own stamp
    const unshift2 = shiftClock(S.b.window, -2 * 60000);
    S.set('notes', 'fresher local');
    unshift2();
    const mine = S.b.info().localAt['items/' + S.vId].notes;
    await S.emitPeer(S.vId, 'notes', 'peer older', isoIn(-4 * 60000));
    eq(S.item().notes, 'fresher local', 'our newer edit beats the peer\'s older one');
    await settle();
    const d2 = S.disk(S.vId);
    eq([d2.fields.notes, d2.fieldsAt.notes], ['fresher local', mine], 'disk carries our value with its commit-time stamp');
    eq(Object.keys(S.b.info().localAt['items/' + S.vId] || {}), [], 'local stamps cleared once written');
  }

  section('N2: the active plan tombstoned with no live plan left → toast, flushing stops');
  {
    const S = await openFresh('N2');
    const seen = [];
    const mo = new S.b.window.MutationObserver((muts) => muts.forEach((m) => m.addedNodes.forEach((n) => { if (n.textContent) seen.push(n.textContent); })));
    mo.observe(S.b.doc.querySelector('#toasts'), { childList: true });
    const hw = JSON.parse(S.tauri.files.get(S.dir + '/headway.json'));
    hw.plans[0].deleted = true; hw.plans[0].updatedAt = isoIn(1000);
    S.tauri.files.set(S.dir + '/headway.json', JSON.stringify(hw));
    await S.tauri.emitPaths(S.dir + '/headway.json');
    ok(seen.some((t) => /deleted by someone else/.test(t)), 'toast says the plan was deleted: ' + seen.join(' | '));
    eq(S.b.info().planGone, true, 'planGone set');
    S.select();
    const mark = S.tauri.log.length;
    S.set('feature', 'into the void');
    await settle();
    eq(shardWrites(S.tauri, mark).length, 0, 'no shard written into the deleted plan');
    eq(S.b.info().docSaved, false, 'stays dirty');
    S.b.click(S.b.doc.querySelector('#btnSave')); // an explicit Sync is refused too
    await settle();
    eq(shardWrites(S.tauri, mark).length, 0, 'Sync writes nothing into the deleted plan either');
    mo.disconnect();
  }
}

// ---- presence P1–P7: heartbeat file, peers → chips (patched, not rendered),
// ignored entries, removal, nudge toast, cleanup on close, web build
const presenceWrites = (tauri, from) => tauri.log.slice(from || 0).filter((l) => l.op === 'writeTextFile' && l.path.indexOf('/presence/') >= 0);
async function presence() {
  const chipsOf = (S, id) => [...S.b.doc.querySelectorAll('.row.item[data-id="' + id + '"] .r-presence .avatar.presence')];
  const PEER = 'peer-zz999';

  section('P1: heartbeat — presence/<userId>.json, never a commit');
  const S = await openFresh('P1');
  const mark0 = S.tauri.log.length; // createBundle's own shard writes are before this
  const own = S.dir + '/presence/' + FIXER.id + '.json';
  eq(S.b.info().presenceOn, true, 'heartbeat armed on open');
  ok(await until(() => S.tauri.files.has(own)), 'presence/<userId>.json written on open');
  let pres = JSON.parse(S.tauri.files.get(own));
  eq([pres.name, pres.planId, pres.editing], [FIXER.name, S.pid, []], 'carries {name, planId, editing}');
  ok(typeof pres.ts === 'number' && Math.abs(Date.now() - pres.ts) < 5000, 'ts is a fresh Date.now(): ' + pres.ts);
  const ren0 = S.b.info().renderCount, hist0 = S.b.HA.userId && S.b.window.__headway.getHistory().length;
  const pw0 = presenceWrites(S.tauri).length;
  ok(await until(() => presenceWrites(S.tauri).length >= pw0 + 3), 'heartbeat keeps writing (' + (presenceWrites(S.tauri).length - pw0) + ' more writes)');
  eq(S.b.info().docSaved, true, 'heartbeats never dirty the document');
  eq(S.b.info().pendingHistory, 0, 'no pending history');
  eq(S.b.window.__headway.getHistory().length, hist0, 'no history line');
  eq(S.b.info().renderCount, ren0, 'no render from heartbeats');
  ok(!S.tauri.files.has(S.dir + '/history/' + FIXER.id + '.jsonl'), 'no history file appeared');
  eq(shardWrites(S.tauri, mark0).length, 0, 'no shard written');
  S.select();
  ok(await until(() => { const p = JSON.parse(S.tauri.files.get(own)); return p.editing.length === 1 && p.editing[0] === S.vId; }), 'selection lands in editing on the next write');
  eq(S.b.info().docSaved, true, 'still clean after the selection write');

  section('P2: a peer file → peerByItem + one ringed chip, patched in place (no full render)');
  const peerPath = S.dir + '/presence/' + PEER + '.json';
  S.tauri.files.set(peerPath, JSON.stringify({ name: 'Zoe Quinn', planId: S.pid, editing: [S.vId], ts: Date.now() }));
  const ren1 = S.b.info().renderCount;
  await S.tauri.emitPaths(peerPath, 'create');
  ok(await until(() => (S.b.info().peerByItem[S.vId] || []).length === 1), 'peerByItem[X] has the peer');
  eq(S.b.info().peerByItem[S.vId][0], { id: PEER, name: 'Zoe Quinn' }, 'peer entry = {id, name}');
  let chips = chipsOf(S, S.vId);
  eq(chips.length, 1, 'one presence chip on the row');
  eq(chips[0].textContent, S.b.RM.initialsOf('Zoe Quinn'), 'chip shows the initials (' + chips[0].textContent + ')');
  ok(chips[0].textContent.length >= 1 && chips[0].textContent.length <= 3, 'initials are 1–3 chars');
  eq(chips[0].getAttribute('title'), 'Zoe Quinn is here', 'chip title names the peer');
  ok(chips[0].classList.contains('sm'), 'reuses the .avatar.sm look');
  eq(S.b.info().renderCount, ren1, 'no full render for a presence change');
  eq(S.b.doc.querySelectorAll('#rows .row.item .r-presence .avatar.presence').length, 1, 'no chip on any other row');
  S.select(); // a full render also produces the chips
  ok(S.b.info().renderCount > ren1, 'select() rendered');
  eq(chipsOf(S, S.vId).length, 1, 'full render carries the chip too');
  // more than three peers on one row → two chips + a +N bubble
  const others = ['peer-a1111', 'peer-b2222', 'peer-c3333'];
  others.forEach((id, i) => S.b.HA.presenceChanged(id, { name: 'Peer ' + 'ABC'[i], planId: S.pid, editing: [S.vId], ts: Date.now() }));
  chips = chipsOf(S, S.vId);
  eq(chips.length, 3, 'capped at three bubbles');
  eq(chips.filter((c) => c.classList.contains('more')).length, 1, 'the last one is the +N bubble');
  eq(chips[2].textContent, '+2', 'reads +2 (4 peers, 2 shown)');
  others.forEach((id) => S.b.HA.presenceChanged(id, null));
  eq(chipsOf(S, S.vId).length, 1, 'back to one chip');

  section('P3: ignored — own id, another plan, a stale heartbeat');
  // a second RENDERED row (an item in a collapsed phase has no row to chip)
  const otherId = [...S.b.doc.querySelectorAll('#rows .row.item[data-id]')].map((r) => r.dataset.id).find((id) => id !== S.vId);
  S.b.HA.presenceChanged(FIXER.id, { name: 'Me', planId: S.pid, editing: [otherId], ts: Date.now() });
  eq(S.b.info().peerByItem[otherId], undefined, 'own userId ignored');
  S.b.HA.presenceChanged('peer-plan2', { name: 'Elsewhere', planId: 'plan-other', editing: [otherId], ts: Date.now() });
  eq(S.b.info().peerByItem[otherId], undefined, 'another planId ignored');
  S.b.HA.presenceChanged('peer-old', { name: 'Ghost', planId: S.pid, editing: [otherId], ts: Date.now() - 100000 });
  eq(S.b.info().peerByItem[otherId], undefined, 'stale ts (> 90 s) ignored');
  eq(chipsOf(S, otherId).length, 0, 'no chip for any of them');
  // stale on disk: the tick's readPresence refresh drops it too
  const stalePath = S.dir + '/presence/peer-old.json';
  S.tauri.files.set(stalePath, JSON.stringify({ name: 'Ghost', planId: S.pid, editing: [otherId], ts: Date.now() - 100000 }));
  await S.tauri.emitPaths(stalePath, 'create');
  await settle();
  eq(S.b.info().peerByItem[otherId], undefined, 'stale file on disk → still no peer on the row');
  eq(chipsOf(S, otherId).length, 0, 'no chip from the stale file');
  S.tauri.files.delete(stalePath);

  section('P4: removal clears the chip');
  S.tauri.files.delete(peerPath);
  await S.tauri.emitPaths(peerPath, 'remove');
  ok(await until(() => chipsOf(S, S.vId).length === 0), 'remove event → chip gone');
  eq(S.b.info().peerByItem[S.vId], undefined, 'peerByItem cleared');
  eq(S.b.info().peers[PEER], undefined, 'peer dropped from peers');
  // and the null form of presenceChanged
  S.b.HA.presenceChanged(PEER, { name: 'Zoe Quinn', planId: S.pid, editing: [S.vId], ts: Date.now() });
  eq(chipsOf(S, S.vId).length, 1, 'chip back');
  S.b.HA.presenceChanged(PEER, null);
  eq(chipsOf(S, S.vId).length, 0, 'presenceChanged(id, null) clears it');

  section('P5: nudge — a flush touching a peer\'s row toasts once per item per minute');
  S.tauri.files.set(peerPath, JSON.stringify({ name: 'Zoe Quinn', planId: S.pid, editing: [S.vId], ts: Date.now() }));
  await S.tauri.emitPaths(peerPath, 'create');
  ok(await until(() => chipsOf(S, S.vId).length === 1), 'peer back on the row');
  const seen = [];
  const mo = new S.b.window.MutationObserver((muts) => muts.forEach((m) => m.addedNodes.forEach((n) => { if (n.textContent) seen.push(n.textContent); })));
  mo.observe(S.b.doc.querySelector('#toasts'), { childList: true });
  const num = S.item().num;
  S.select();
  S.set('notes', 'nudge me');
  await settle();
  const nudges = () => seen.filter((t) => /also editing/.test(t));
  eq(nudges().length, 1, 'exactly one nudge toast: ' + seen.join(' | '));
  eq(nudges()[0], 'Zoe Quinn is also editing #' + num, 'names the peer and the item number');
  eq(S.disk(S.vId).fields.notes, 'nudge me', 'the write still went ahead (advisory only)');
  S.set('notes', 'nudge again');
  await settle();
  eq(nudges().length, 1, 'a second commit inside the window → no second toast');
  eq(S.disk(S.vId).fields.notes, 'nudge again', 'second write landed too');
  eq(S.b.info().docSaved, true, 'synced');
  // another item with no peer on it → no toast
  S.select(otherId);
  S.set('notes', 'quiet');
  await settle();
  eq(nudges().length, 1, 'an item nobody else is on → no nudge');
  mo.disconnect();

  section('P5b: deleting a row a peer is on is not a nudge; peer names are escaped');
  // a row that is actually rendered (not every item has a row in this view)
  const delId = [...S.b.doc.querySelectorAll('#rows .row.item')].map((r) => r.dataset.id).find((id) => id && id !== S.vId && id !== otherId);
  ok(!!delId, 'found a third rendered row to delete');
  S.b.HA.presenceChanged(PEER, { name: 'Zoe Quinn', planId: S.pid, editing: [delId], ts: Date.now() });
  ok(await until(() => (S.b.info().peerByItem[delId] || []).length === 1), 'peer sits on the row we are about to delete');
  const seenDel = [];
  const moDel = new S.b.window.MutationObserver((muts) => muts.forEach((m) => m.addedNodes.forEach((n) => { if (n.textContent) seenDel.push(n.textContent); })));
  moDel.observe(S.b.doc.querySelector('#toasts'), { childList: true });
  S.select(delId);
  S.b.key('Delete');
  const okBtn = S.b.doc.querySelector('#modalHost [data-m="ok"]');
  ok(!!okBtn, 'delete asks for confirmation');
  if (okBtn) okBtn.click();
  await settle();
  ok(!S.item(delId), 'the row is gone locally');
  ok(await until(() => { try { return S.disk(delId).deleted === true; } catch (e) { return false; } }), 'a tombstone reached the disk');
  eq(seenDel.filter((t) => /also editing/.test(t)).length, 0, 'no "also editing" nudge for a delete: ' + seenDel.join(' | '));
  moDel.disconnect();
  // a hostile peer name renders as text, never markup
  S.b.HA.presenceChanged(PEER, { name: 'Zoe <b>Q</b> & "Co"', planId: S.pid, editing: [S.vId], ts: Date.now() });
  const hot = chipsOf(S, S.vId);
  eq(hot.length, 1, 'chip for the odd-named peer');
  ok(hot[0].innerHTML.indexOf('<b>') === -1 && hot[0].querySelector('b') === null, 'name markup is escaped in the chip');
  eq(hot[0].getAttribute('title'), 'Zoe <b>Q</b> & "Co" is here', 'title carries the raw name as text');
  S.b.HA.presenceChanged(PEER, { name: 'Zoe Quinn', planId: S.pid, editing: [S.vId], ts: Date.now() });

  section('P6: closeBundleSession / beforeClose remove our file and stop the heartbeat');
  ok(S.tauri.files.has(own), 'our presence file is there before close');
  await S.b.HA.closeBundleSession();
  ok(!S.tauri.files.has(own), 'closeBundleSession removed presence/<userId>.json');
  eq(S.b.info().presenceOn, false, 'heartbeat off');
  eq(S.b.info().peerByItem, {}, 'peerByItem cleared');
  let mark = presenceWrites(S.tauri).length;
  await settle(6);
  eq(presenceWrites(S.tauri).length, mark, 'no presence write after close');
  ok(!S.tauri.files.has(own), 'file stays gone');
  const S2 = await openFresh('P6b');
  const own2 = S2.dir + '/presence/' + FIXER.id + '.json';
  ok(await until(() => S2.tauri.files.has(own2)), 'second session heartbeats');
  await S2.b.HA.beforeClose();
  ok(!S2.tauri.files.has(own2), 'beforeClose removed the file');
  eq(S2.b.info().presenceOn, false, 'heartbeat off after beforeClose');
  mark = presenceWrites(S2.tauri).length;
  await settle(6);
  eq(presenceWrites(S2.tauri).length, mark, 'no presence write after beforeClose');
  // reopening re-arms it (the same id, a fresh file)
  await S2.b.HA.openBundleDoc(S2.dir);
  ok(await until(() => S2.tauri.files.has(own2)), 'reopen writes the heartbeat again');
  eq(S2.b.info().presenceOn, true, 're-armed');
  await S2.b.HA.closeBundleSession();

  section('P7: web build — no heartbeat, no errors');
  const w = boot(null, { localStorage: { 'headway-user-v2': JSON.stringify(FIXER) } });
  ok(w.errors.length === 0, 'browser build boots clean' + (w.errors.length ? ' — ' + w.errors.join('; ') : ''));
  eq(w.info().presenceOn, false, 'no heartbeat without HeadwayDesktop');
  w.HA.presenceChanged(PEER, { name: 'Zoe Quinn', planId: null, editing: ['x'], ts: Date.now() });
  eq(w.info().peerByItem, {}, 'presenceChanged is inert outside a bundle (planId never matches)');
  eq(w.HA.editingIds(), [], 'editingIds empty');
  await settle(3);
  ok(w.errors.length === 0, 'still no errors after a few ticks');
}

// ---- Import from Excel… (phase 7): a workbook merged INTO the open bundle,
// add-only, previewed, applied through commit
async function importFlow() {
  section('Import from Excel…: menu item only in bundle mode on the desktop');
  let nextPick = null;
  const S = await openFresh('IMP', { dialogOpen: () => nextPick });
  const { b, tauri, dir } = S;
  const doc = b.doc, window = b.window;
  ok(b.menuLabels('file').some((l) => /^Import from Excel…/.test(l)), 'File menu offers Import from Excel… in bundle mode');
  ok(b.HA.menuItems('macApp').some((m) => /Import from Excel…/.test(m.label || '')), 'macApp list mirrors it');
  doc.querySelector('#popover').hidden = true;

  section('Import from Excel…: preview counts, add-only apply, exact shard writes');
  // the workbook: this roadmap exported elsewhere and edited — one new
  // feature, one new story under M, M's empty Enables filled, M's Notes changed
  const base = b.RB.exportableState(b.state());
  const M = base.items.find((i) => i.notes && !i.enables);
  ok(!!M, 'a fixture feature with Notes set and Enables empty');
  const mod = JSON.parse(JSON.stringify(base));
  const mM = mod.items.find((i) => i.id === M.id);
  mM.enables = 'Filled from Excel';
  mM.notes = 'CONFLICT — the roadmap keeps its own';
  mM.stories.push({ id: 's-imp-new', title: 'Imported story' });
  mod.items.push({ id: 'i-imp-new', num: 999, feature: 'Imported feature', phaseId: M.phaseId, workstream: M.workstream, teamType: M.teamType, stories: [] });
  const blob = await window.RMExcel.exportWorkbook(b.RM.normalizeState(mod));
  const u8 = new Uint8Array(Buffer.from(await blob.arrayBuffer())); // node realm for ExcelJS
  tauri.dirs.add('C:/tmp');
  tauri.files.set('C:/tmp/Team edits.xlsx', u8);
  nextPick = 'C:/tmp/Team edits.xlsx';
  const nItems = b.state().items.length, nStories = S.item(M.id).stories.length;
  const undoLen = b.info().undoLen;
  const histPath = dir + '/history/' + FIXER.id + '.jsonl';
  const histLen = () => b.RB.parseHistory(tauri.files.get(histPath) || '').length;
  const hist0 = histLen();
  // openModal swaps the #modalHost node (resetNode) — always re-query it
  const mh = () => doc.querySelector('#modalHost');
  // toasts auto-dismiss within milliseconds under the compressed timers: catch them as they are added
  const seenToasts = [];
  const mo = new window.MutationObserver((muts) => muts.forEach((m) => m.addedNodes.forEach((n) => { if (n.textContent) seenToasts.push(n.textContent); })));
  mo.observe(doc.querySelector('#toasts'), { childList: true });
  b.menuClick('file', /Import from Excel…/);
  ok(await until(() => !mh().hidden && /Import from/.test(mh().textContent)), 'the preview modal opens — toasts: ' + seenToasts.join(' | ') + ' errors: ' + b.errors.join('; '));
  eq(seenToasts.length, 0, 'no error toast on the way to the preview');
  const txt = mh().textContent;
  ok(/Import from “Team edits\.xlsx”/.test(txt), 'titled after the file');
  ok(/1 new feature(?!s)/.test(txt), 'says 1 new feature'); // li texts run together in textContent
  ok(/1 new story/.test(txt), 'says 1 new story');
  ok(/1 field filled in/.test(txt), 'says 1 field filled in');
  ok(/0 team members/.test(txt) && /0 phases/.test(txt), 'says 0 team members, 0 phases');
  ok(/1 difference left alone \(the shared roadmap wins\)/.test(txt), 'says 1 difference left alone: ' + txt.replace(/\s+/g, ' ').slice(0, 200));
  ok(/Imported feature/.test(txt), 'lists the new feature title');
  ok(!/template layout/.test(txt), 'a Headway workbook: no template note');
  eq(b.state().items.length, nItems, 'nothing applied before Import');
  let mark = tauri.log.length;
  b.click(mh().querySelector('[data-m="ok"]'));
  ok(mh().hidden, 'modal closed');
  ok(await until(() => seenToasts.some((t) => /^Imported 1 feature, 1 story, 1 field filled in/.test(t))), 'summary toast: ' + seenToasts.join(' | '));
  mo.disconnect();
  eq(b.state().items.length, nItems + 1, 'one feature added');
  const added = b.state().items.find((i) => i.feature === 'Imported feature');
  ok(!!added && added.id === 'i-imp-new', 'the new feature keeps its workbook id (free in the roadmap)');
  eq(added.num, b.RM.nextNum({ items: b.state().items.filter((i) => i.id !== added.id) }), '…and takes the next num');
  eq(S.item(M.id).stories.length, nStories + 1, 'one story added under M');
  eq(S.item(M.id).enables, 'Filled from Excel', 'the empty field is filled');
  eq(S.item(M.id).notes, M.notes, 'the conflicting field is untouched');
  eq(b.info().undoLen, undoLen + 1, 'one undo entry');
  await settle();
  let writes = shardWrites(tauri, mark);
  eq(writes.map((w) => w.path).sort(), [S.path(added.id) + '.tmp', S.path(M.id) + '.tmp'].sort(), 'exactly the added item and M written — nothing else');
  const mDisk = S.disk(M.id);
  eq(mDisk.fields.enables, 'Filled from Excel', 'M on disk: filled');
  eq(mDisk.fields.notes, M.notes, 'M on disk: the conflict left alone');
  ok(mDisk.fields.stories.some((s) => s.id === 's-imp-new'), 'M on disk: the new story');
  eq(S.disk(added.id).fields.feature, 'Imported feature', 'the new shard on disk');
  eq(histLen(), hist0 + 1, 'one history line');
  const lines = b.RB.parseHistory(tauri.files.get(histPath));
  eq(lines[lines.length - 1].label, 'import from Excel', '…labelled import from Excel');

  section('Import from Excel…: undo reverts it; the same workbook again has nothing to import');
  b.key('z', { metaKey: true });
  eq(b.state().items.length, nItems, 'undo removed the added feature');
  eq(S.item(M.id).enables, '', 'undo cleared the fill');
  eq(S.item(M.id).stories.length, nStories, 'undo removed the story');
  mark = tauri.log.length;
  await settle();
  writes = shardWrites(tauri, mark);
  eq(writes.map((w) => w.path).sort(), [S.path(added.id) + '.tmp', S.path(M.id) + '.tmp'].sort(), 'the flush after undo touches the same two shards');
  eq(S.disk(added.id).deleted, true, 'the added feature is tombstoned');
  eq(S.disk(M.id).fields.enables, '', 'M on disk: fill reverted');
  ok(!S.disk(M.id).fields.stories.some((s) => s.id === 's-imp-new'), 'M on disk: story gone');
  b.key('z', { metaKey: true, shiftKey: true }); // redo → the import is back
  eq(b.state().items.length, nItems + 1, 'redo restores the import');
  await settle();
  eq(S.disk(added.id).deleted, false, 'the shard is live again');
  b.menuClick('file', /Import from Excel…/);
  ok(await until(() => !mh().hidden && /Import from/.test(mh().textContent)), 'the preview opens again');
  ok(/Nothing new to import/.test(mh().textContent), 'says Nothing new to import');
  ok(/1 difference left alone/.test(mh().textContent), '…still reporting the difference left alone');
  ok(!mh().querySelector('[data-m="ok"]'), 'no Import button');
  eq(mh().querySelector('[data-m="cancel"]').textContent.trim(), 'Close', 'just Close');
  b.click(mh().querySelector('[data-m="cancel"]'));
  ok(mh().hidden, 'closed');
  eq(b.state().items.length, nItems + 1, 'nothing changed');
  nextPick = null; // dialog cancelled
  mark = tauri.log.length;
  b.menuClick('file', /Import from Excel…/);
  await settle();
  ok(mh().hidden && shardWrites(tauri, mark).length === 0, 'dialog cancel: no modal, no shard writes');
  eq([b.info().docKind, b.HD.currentPath()], ['bundle', null], 'the workbook was never adopted');

  section('Import from Excel…: absent for an .xlsx document and in the web build');
  await b.HA.loadBuffer(u8.buffer, 'Team edits.xlsx');
  await settle();
  eq(b.info().docKind, 'xlsx', 'standalone document');
  ok(!b.menuLabels('file').some((l) => /Import from Excel/.test(l)), 'File menu: no Import for an .xlsx');
  ok(!b.HA.menuItems('macApp').some((m) => /Import from Excel/.test(m.label || '')), 'macApp: none either');
  doc.querySelector('#popover').hidden = true;
  const w = boot(null, { localStorage: { 'headway-user-v2': JSON.stringify(FIXER) } });
  ok(!w.HA.menuItems('file').some((m) => /Import from Excel/.test(m.label || '')), 'web build: no Import');
  ok(w.errors.length === 0, 'web build boots clean');
}

main().catch((e) => {
  console.error('  ✗ suite threw: ' + (e && e.stack || e));
  process.exit(1);
});
