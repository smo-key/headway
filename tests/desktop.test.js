/* Desktop (Tauri) folder backend against an in-memory fake Tauri fs (jsdom).
 * Run: NODE_PATH=./node_modules node tests/desktop.test.js */
'use strict';
const fs = require('fs');
const path = require('path');

let JSDOM;
try { JSDOM = require('jsdom').JSDOM; }
catch (e) { console.log('(skipped — jsdom not resolvable in NODE_PATH)'); process.exit(0); }

const makeFakeTauri = require('./fake-tauri.js');
const ROOT = path.join(__dirname, '..');

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

// boot core + bundle + desktop in a bare jsdom against one fake Tauri
function boot(tauri, extra) {
  const dom = new JSDOM(
    '<!doctype html><html><body><div id="topbar"><div class="tb-right"></div></div>' +
    '<span id="docTitle"></span><div id="toasts"></div></body></html>',
    { url: 'http://localhost/index.html', runScripts: 'outside-only', pretendToBeVisual: true });
  const { window } = dom;
  // the read-retry ladder is 1.2 s / 3 s / 8 s in production; compress every
  // timer so a corrupt-shard case does not take 12 s
  const realSetTimeout = window.setTimeout.bind(window);
  window.setTimeout = function (fn, ms) { return realSetTimeout(fn, Math.min(+ms || 0, 8)); };
  window.__TAURI__ = { core: {}, fs: tauri.fs, dialog: tauri.dialog, window: tauri.window };
  const calls = [];
  const rec = (name) => function () { calls.push({ name, args: Array.prototype.slice.call(arguments) }); return Promise.resolve(); };
  window.HeadwayApp = {
    toast: rec('toast'), noteRecent: rec('noteRecent'), applyExternalEntities: rec('applyExternalEntities'),
    presenceChanged: rec('presenceChanged'), plansChanged: rec('plansChanged'), loadBuffer: rec('loadBuffer'),
    beforeClose: rec('beforeClose'), planShardChanged: rec('planShardChanged'),
  };
  (extra || []).forEach((n) => { window.HeadwayApp[n] = rec(n); });
  const errors = [];
  window.addEventListener('error', (e) => errors.push(e.message));
  for (const f of ['js/core.js', 'js/bundle.js', 'js/desktop.js']) {
    try { window.eval(fs.readFileSync(path.join(ROOT, f), 'utf8')); }
    catch (e) { failed++; console.error('  ✗ ' + f + ' threw on load: ' + e.message); }
  }
  const named = (n) => calls.filter((c) => c.name === n);
  return { window, calls, named, errors, RM: window.RM, RB: window.RMBundle, HD: window.HeadwayDesktop };
}

const USER = 'tester-abc12';
const T0 = '2026-09-01T10:00:00.000Z';
const T1 = '2026-09-01T10:05:00.000Z';
const T2 = '2026-09-01T10:10:00.000Z';
const DIR = 'C:/Users/me/OneDrive/Roadmap.headway';

async function main() {
  const tauri = makeFakeTauri();
  const { window, named, errors, RM, RB, HD } = boot(tauri);

  section('boot');
  ok(errors.length === 0, 'no window errors during boot' + (errors.length ? ' — ' + errors.join('; ') : ''));
  ok(HD && typeof HD.openBundle === 'function' && typeof HD.openDialog === 'function', 'HeadwayDesktop has bundle + xlsx methods');
  ['openBundle', 'readPlan', 'flushShards', 'appendHistory', 'rewriteHistory', 'readHistory', 'writePresence', 'readPresence',
    'removePresence', 'removeShard', 'createBundle', 'openBundleDialog', 'closeBundle', 'setUserId', 'bundleDir',
    'activePlanId', 'bundleWarnings', 'friendlyFsError', 'classify'].forEach((m) => {
    ok(typeof HD[m] === 'function', 'HeadwayDesktop.' + m + ' exists');
  });

  section('helpers');
  eq(HD.classify('headway.json').kind, 'plans', 'classify headway.json');
  eq(HD.classify('plans/p1/meta.json').kind, 'meta', 'classify meta.json');
  eq(HD.classify('plans/p1/items/i1-2-abc.json'), { kind: 'item', planId: 'p1', file: 'i1-2-abc.json', stem: 'i1-2-abc', entityKind: 'items', userId: null }, 'classify item shard');
  eq(HD.classify('plans/p1/phases/x.json').kind, 'phase', 'classify phase');
  eq(HD.classify('plans/p1/team/x.json').kind, 'team', 'classify team');
  eq(HD.classify('plans/p1/costs/x.json').kind, 'cost', 'classify cost');
  eq(HD.classify('plans/p1/items/x.json.tmp').kind, 'tmp', 'classify .tmp');
  eq(HD.classify('history/bob-1.jsonl'), { kind: 'history', planId: null, file: 'bob-1.jsonl', stem: 'bob-1', entityKind: null, userId: 'bob-1' }, 'classify history');
  eq(HD.classify('presence/bob-1.json').kind, 'presence', 'classify presence');
  eq(HD.classify('plans/p1/items/nested/x.json').kind, 'other', 'classify unknown depth');
  eq(HD.classify('desktop.ini').kind, 'other', 'classify stray file');
  eq(HD.friendlyFsError('fs.rename not allowed. Permissions associated with this command: fs:allow-rename'),
    'Headway is missing capability fs:allow-rename', 'friendlyFsError maps a plugin denial');
  eq(HD.friendlyFsError('fs.write_text_file not allowed'), 'Headway is missing capability fs:allow-write-text-file', 'friendlyFsError derives the cap from the command');
  eq(HD.friendlyFsError(new Error('disk full')), 'disk full', 'friendlyFsError passes other errors through');

  section('createBundle + openBundle');
  const fixture = RM.normalizeState(JSON.parse(JSON.stringify(require('./seed.fixture.js'))));
  const contents = RB.migrateFromState(fixture, USER, T0);
  const pid = contents.headway.plans[0].id;
  const out = await HD.createBundle(DIR, contents);
  eq(out, DIR, 'createBundle resolves the dir');
  ok(tauri.files.has(DIR + '/headway.json'), 'headway.json written');
  eq(tauri.files.has(DIR + '/plans/' + pid + '/meta.json'), true, 'meta.json written');
  const shardCount = [...tauri.files.keys()].filter((f) => f.indexOf(DIR + '/plans/' + pid + '/items/') === 0).length;
  eq(shardCount, fixture.items.length, 'one item shard per item');
  ok([...tauri.files.keys()].every((f) => !/\.tmp$/.test(f)), 'no .tmp left behind');

  let opened = await HD.openBundle(DIR);
  eq(opened.planId, pid, 'openBundle picks the first plan');
  eq(opened.doc.items.length, fixture.items.length, 'item count matches the fixture');
  eq(opened.doc.phases.length, fixture.phases.length, 'phase count matches the fixture');
  eq(opened.doc.team.length, fixture.team.length, 'team count matches the fixture');
  eq(opened.doc.meta.title, fixture.meta.title, 'meta.title survives');
  eq(opened.warnings, [], 'no warnings on a clean bundle');
  eq(opened.envs.items.length, fixture.items.length, 'envelopes handed back for lastCanon priming');
  eq(HD.bundleDir(), DIR, 'bundleDir set');
  eq(HD.activePlanId(), pid, 'activePlanId set');
  eq(HD.currentPath(), null, 'currentPath stays null in bundle mode');
  eq(window.document.title, 'Roadmap — Headway', 'title from the folder name');
  ok(tauri.watching() && tauri.watchOpts().recursive === true, 'recursive watch installed');
  eq(named('noteRecent').slice(-1)[0].args, [DIR, 'bundle'], 'noteRecent(dir, "bundle")');

  section('corrupt shard is skipped with a warning');
  const itemFiles = [...tauri.files.keys()].filter((f) => f.indexOf(DIR + '/plans/' + pid + '/items/') === 0).sort();
  const badPath = itemFiles[0];
  const badEnv = JSON.parse(tauri.files.get(badPath));
  const goodText = tauri.files.get(badPath);
  tauri.files.set(badPath, '{');
  opened = await HD.openBundle(DIR);
  eq(opened.doc.items.length, fixture.items.length - 1, 'the corrupt item is missing');
  ok(!opened.doc.items.some((it) => it.id === badEnv.id), 'and it is exactly the corrupt one');
  eq(opened.warnings.length, 1, 'exactly one warning');
  ok(opened.warnings[0].path.indexOf(badPath) >= 0, 'warning names the shard');
  tauri.files.set(badPath, goodText);
  opened = await HD.openBundle(DIR);
  eq(opened.warnings.length, 0, 'clean again');

  section('flushShards writes only the changed shards, atomically');
  const envsById = {};
  opened.envs.items.forEach((e) => { envsById[e.id] = e; });
  const [idA, idB] = opened.doc.items.slice(0, 2).map((it) => it.id);
  const entA = Object.assign({}, RB.unwrap(envsById[idA]), { notes: 'flushed A' });
  const entB = Object.assign({}, RB.unwrap(envsById[idB]), { size: 'XL' });
  const envA = RB.wrap(entA, envsById[idA], USER, T1);
  const envB = RB.wrap(entB, envsById[idB], USER, T1);
  const mark = tauri.log.length;
  const res = await HD.flushShards(DIR, pid, [
    { kind: 'items', id: idA, env: envA, baseRev: envsById[idA].rev },
    { kind: 'items', id: idB, env: envB, baseRev: envsById[idB].rev },
  ]);
  const writes = tauri.log.slice(mark).filter((l) => l.op === 'writeTextFile' && l.path.indexOf('/plans/') >= 0);
  eq(writes.length, 2, 'exactly 2 shard writes');
  eq(res.written.map((w) => w.id), [idA, idB], 'written ids');
  eq(res.merged, [], 'nothing needed a disk merge');
  const iA = tauri.log.findIndex((l, i) => i >= mark && l.op === 'writeTextFile' && l.path === DIR + '/plans/' + pid + '/items/' + idA + '.json.tmp');
  ok(iA >= 0, 'writeTextFile to <id>.json.tmp logged');
  ok(tauri.log[iA + 1] && tauri.log[iA + 1].op === 'rename' && tauri.log[iA + 1].to === DIR + '/plans/' + pid + '/items/' + idA + '.json',
    'immediately followed by rename → <id>.json');
  const diskA = JSON.parse(tauri.files.get(DIR + '/plans/' + pid + '/items/' + idA + '.json'));
  eq(diskA.fields.notes, 'flushed A', 'shard on disk carries the new value');
  eq(diskA.rev, envsById[idA].rev + 1, 'rev bumped');
  ok(!tauri.files.has(DIR + '/plans/' + pid + '/items/' + idA + '.json.tmp'), 'tmp gone after rename');
  ok(/\n  "fields"/.test(tauri.files.get(DIR + '/plans/' + pid + '/items/' + idA + '.json')), 'shards are pretty-printed');
  eq(RB.canonicalize(diskA), res.written[0].canon, 'returned canon = canonical of what is on disk');

  section('echo: our own write is not applied');
  let before = named('applyExternalEntities').length;
  await tauri.emitPaths(DIR + '/plans/' + pid + '/items/' + idA + '.json');
  await tauri.emitPaths(DIR.replace(/\//g, '\\') + '\\plans\\' + pid + '\\items\\' + idB + '.json');
  await tick();
  eq(named('applyExternalEntities').length, before, 'applyExternalEntities not called for own writes (either separator)');
  await tauri.emitPaths(DIR + '/plans/' + pid + '/items/' + idA + '.json.tmp');
  await tauri.emitPaths(DIR + '/history/' + USER + '.jsonl');
  await tick();
  eq(named('applyExternalEntities').length, before, '.tmp and history events are ignored');

  section('peer change is applied once with the disk envelope');
  const peer = JSON.parse(JSON.stringify(diskA));
  peer.fields.notes = 'peer note';
  peer.fieldsAt.notes = T2;
  peer.updatedAt = T2; peer.updatedBy = 'peer-zz999'; peer.rev = diskA.rev + 1;
  tauri.files.set(DIR + '/plans/' + pid + '/items/' + idA + '.json', JSON.stringify(peer));
  before = named('applyExternalEntities').length;
  await tauri.emitPaths(DIR + '/plans/' + pid + '/items/' + idA + '.json');
  await tick();
  eq(named('applyExternalEntities').length, before + 1, 'applyExternalEntities called once');
  const batch = named('applyExternalEntities').slice(-1)[0].args[0];
  eq(batch.length, 1, 'one entity in the batch');
  eq(batch[0].kind, 'items', 'kind is the RMBundle kind');
  eq(batch[0].id, idA, 'id is the item id');
  eq(batch[0].env.fields.notes, 'peer note', 'env carries the peer notes');
  eq(batch[0].env.fields.size, diskA.fields.size, '…and the other fields from disk');
  eq(batch[0].env.fields.name, diskA.fields.name, '…including name');
  await tauri.emitPaths(DIR + '/plans/' + pid + '/items/' + idA + '.json');
  await tick();
  eq(named('applyExternalEntities').length, before + 1, 'a repeat event for the same content is an echo now');
  // a change in another plan is not applied live
  tauri.files.set(DIR + '/plans/other-plan/items/' + idA + '.json', JSON.stringify(peer));
  await tauri.emitPaths(DIR + '/plans/other-plan/items/' + idA + '.json');
  await tick();
  eq(named('applyExternalEntities').length, before + 1, 'other plan ignored');
  // batch: two peer files in one event → one call
  const peerB = JSON.parse(tauri.files.get(DIR + '/plans/' + pid + '/items/' + idB + '.json'));
  peerB.fields.notes = 'peer B'; peerB.fieldsAt.notes = T2; peerB.updatedAt = T2; peerB.updatedBy = 'peer-zz999'; peerB.rev++;
  const peerA2 = JSON.parse(JSON.stringify(peer)); peerA2.fields.notes = 'peer note 2'; peerA2.fieldsAt.notes = '2026-09-01T10:11:00.000Z'; peerA2.updatedAt = peerA2.fieldsAt.notes; peerA2.rev++;
  tauri.files.set(DIR + '/plans/' + pid + '/items/' + idB + '.json', JSON.stringify(peerB));
  tauri.files.set(DIR + '/plans/' + pid + '/items/' + idA + '.json', JSON.stringify(peerA2));
  await tauri.emitPaths([DIR + '/plans/' + pid + '/items/' + idA + '.json', DIR + '/plans/' + pid + '/items/' + idB + '.json']);
  await tick();
  eq(named('applyExternalEntities').length, before + 2, 'two paths in one event → one call');
  eq(named('applyExternalEntities').slice(-1)[0].args[0].map((x) => x.id).sort(), [idA, idB].sort(), '…carrying both ids');

  section('flushShards merges a shard a peer moved on');
  // our env was based on rev r, but disk is now at peerA2.rev with a newer notes stamp
  const mine = RB.wrap(Object.assign({}, RB.unwrap(diskA), { size: 'S' }), diskA, USER, T1);
  const res2 = await HD.flushShards(DIR, pid, [{ kind: 'items', id: idA, env: mine, baseRev: diskA.rev }]);
  eq(res2.merged, [idA], 'merge reported');
  const afterMerge = JSON.parse(tauri.files.get(DIR + '/plans/' + pid + '/items/' + idA + '.json'));
  eq(afterMerge.fields.size, 'S', 'my size landed');
  eq(afterMerge.fields.notes, 'peer note 2', 'peer notes preserved');
  eq(res2.written[0].env.fields.notes, 'peer note 2', 'returned env is the merged one');

  section('conflict sibling absorbed and removed');
  const canonPath = DIR + '/plans/' + pid + '/items/' + idB + '.json';
  const canonEnv = JSON.parse(tauri.files.get(canonPath));
  const sib = JSON.parse(JSON.stringify(canonEnv));
  sib.fields.size = 'XS'; sib.fieldsAt.size = T2; sib.updatedAt = T2; sib.updatedBy = 'peer-zz999';
  const sibPath = DIR + '/plans/' + pid + '/items/' + idB + '-OMC-HOST.json';
  tauri.files.set(sibPath, JSON.stringify(sib));
  before = named('applyExternalEntities').length;
  await tauri.emitPaths(sibPath, 'create');
  await tick();
  const merged = JSON.parse(tauri.files.get(canonPath));
  eq(merged.fields.size, 'XS', 'canonical shard has the sibling size');
  eq(merged.fields.notes, canonEnv.fields.notes, 'canonical keeps its own newer fields');
  ok(!tauri.files.has(sibPath), 'sibling file removed');
  eq(named('applyExternalEntities').length, before + 1, 'merged result applied once');
  eq(named('applyExternalEntities').slice(-1)[0].args[0][0].env.fields.size, 'XS', '…with the merged env');
  // the canonical's own watch event after the absorb is an echo
  await tauri.emitPaths(canonPath);
  await tick();
  eq(named('applyExternalEntities').length, before + 1, 'rewrite of the canonical is an echo');
  // a sibling found on open is absorbed too
  const sib2 = JSON.parse(JSON.stringify(merged));
  sib2.fields.notes = 'sibling on open'; sib2.fieldsAt.notes = '2026-09-01T10:20:00.000Z'; sib2.updatedAt = sib2.fieldsAt.notes; sib2.updatedBy = 'peer-zz999';
  tauri.files.set(sibPath, JSON.stringify(sib2));
  opened = await HD.openBundle(DIR);
  ok(!tauri.files.has(sibPath), 'sibling present at open is removed');
  eq(JSON.parse(tauri.files.get(canonPath)).fields.notes, 'sibling on open', 'canonical shard absorbed it');
  eq(opened.doc.items.filter((it) => it.id === idB)[0].notes, 'sibling on open', 'assembled doc reflects the merge');
  eq(opened.doc.items.length, fixture.items.length, 'no duplicate item from the sibling');
  eq(opened.warnings, [], 'no warnings');
  // removeShard refuses the canonical
  let refused = null;
  await HD.removeShard(DIR, 'plans/' + pid + '/items/' + idB + '.json').catch((e) => { refused = e; });
  ok(refused && /refusing/.test(refused.message), 'removeShard refuses a canonical shard');
  ok(tauri.files.has(canonPath), '…and the file is still there');
  refused = null;
  await HD.removeShard(DIR, 'headway.json').catch((e) => { refused = e; });
  ok(refused && /refusing/.test(refused.message), 'removeShard refuses headway.json');

  section('plans + presence events');
  before = named('plansChanged').length;
  await tauri.emitPaths(DIR + '/headway.json');
  await tick();
  eq(named('plansChanged').length, before + 1, 'plansChanged called');
  eq(named('plansChanged').slice(-1)[0].args[0].plans[0].id, pid, '…with the parsed headway.json');
  HD.setUserId(USER);
  await HD.writePresence(DIR, 'peer-zz999', { name: 'Peer', planId: pid, editing: [idA], ts: 1 });
  await tauri.emitPaths(DIR + '/presence/peer-zz999.json');
  await tick();
  eq(named('presenceChanged').slice(-1)[0].args, ['peer-zz999', { name: 'Peer', planId: pid, editing: [idA], ts: 1 }], 'presenceChanged(userId, obj)');
  before = named('presenceChanged').length;
  await HD.writePresence(DIR, USER, { name: 'Me', planId: pid, editing: [], ts: 2 });
  await tauri.emitPaths(DIR + '/presence/' + USER + '.json');
  await tick();
  eq(named('presenceChanged').length, before, 'own presence event skipped');
  await tauri.emitPaths(DIR + '/presence/peer-zz999.json', 'remove');
  await tick();
  eq(named('presenceChanged').slice(-1)[0].args, ['peer-zz999', null], 'remove → presenceChanged(userId, null)');

  section('presence round-trip');
  const pres = await HD.readPresence(DIR);
  eq(Object.keys(pres).sort(), ['peer-zz999', USER].sort(), 'readPresence lists both');
  eq(pres[USER].name, 'Me', '…with parsed content');
  tauri.files.set(DIR + '/presence/broken.json', '{nope');
  eq(Object.keys(await HD.readPresence(DIR)).indexOf('broken'), -1, 'unparsable presence skipped');
  await HD.removePresence(DIR, 'peer-zz999');
  ok(!tauri.files.has(DIR + '/presence/peer-zz999.json'), 'removePresence deletes');
  await HD.removePresence(DIR, 'peer-zz999');
  ok(true, 'removePresence of a missing file resolves');

  section('history');
  const line = RB.historyLine({ t: 1, u: 'Me', label: 'Edit', n: 1 }, USER, pid);
  await HD.appendHistory(DIR, USER, line);
  await HD.appendHistory(DIR, USER, Object.assign({}, line, { t: 2 }));
  const hist = await HD.readHistory(DIR);
  eq(hist[USER].slice(-2).map((l) => l.t), [1, 2], 'appendHistory appends, readHistory parses');
  eq(hist[USER].length, contents.history[USER].length + 2, 'legacy lines from the migration kept');
  await HD.rewriteHistory(DIR, USER, [line]);
  eq((await HD.readHistory(DIR))[USER].length, 1, 'rewriteHistory replaces the file');

  section('closeBundle');
  await HD.closeBundle();
  eq(HD.bundleDir(), null, 'bundleDir cleared');
  eq(HD.activePlanId(), null, 'activePlanId cleared');
  ok(!tauri.files.has(DIR + '/presence/' + USER + '.json'), 'own presence removed on close');
  ok(!tauri.watching(), 'watcher stopped');
  eq(window.document.title, 'Headway — Roadmap Planner', 'title reset');

  section('openBundle failures');
  let err = null;
  await HD.openBundle('C:/nowhere/Nope.headway').catch((e) => { err = e; });
  ok(err && /headway\.json|No such file/i.test(err.message), 'missing headway.json rejects: ' + (err && err.message));
  tauri.dirs.add('C:/x/Bad.headway');
  tauri.files.set('C:/x/Bad.headway/headway.json', '{"format":"headway-bundle-v1"}');
  err = null;
  await HD.openBundle('C:/x/Bad.headway').catch((e) => { err = e; });
  ok(err && /plan list/.test(err.message), 'headway.json without plans rejects');
  eq(HD.bundleDir(), null, 'a failed open leaves no bundle state');

  section('close hook');
  const closeCalls = named('beforeClose').length;
  const prevented = await tauri.requestClose();
  ok(prevented === true, 'close request is intercepted');
  eq(named('beforeClose').length, closeCalls + 1, 'beforeClose called');
  ok(tauri.log.some((l) => l.op === 'destroy'), 'window destroyed after beforeClose settled');

  section('denied capability surfaces, never masked');
  const t2 = makeFakeTauri({ deny: ['rename'] });
  const b2 = boot(t2);
  const c2 = b2.RB.migrateFromState(b2.RM.normalizeState(JSON.parse(JSON.stringify(require('./seed.fixture.js')))), USER, T0);
  err = null;
  await b2.HD.createBundle(DIR, c2).catch((e) => { err = e; });
  ok(err && err.message.indexOf('fs:allow-rename') >= 0, 'createBundle rejects naming fs:allow-rename: ' + (err && err.message));
  // a bundle that already exists (written by a peer), then a flush from a build missing the cap
  const t3 = makeFakeTauri();
  const b3 = boot(t3);
  await b3.HD.createBundle(DIR, c2);
  const o3 = await b3.HD.openBundle(DIR);
  const id3 = o3.doc.items[0].id;
  const e3 = o3.envs.items.filter((e) => e.id === id3)[0];
  t3.deny.add('rename');
  const shardPath = DIR + '/plans/' + o3.planId + '/items/' + id3 + '.json';
  const textBefore = t3.files.get(shardPath);
  const mark3 = t3.log.length;
  err = null;
  await b3.HD.flushShards(DIR, o3.planId, [{ kind: 'items', id: id3, env: b3.RB.wrap(Object.assign({}, b3.RB.unwrap(e3), { notes: 'nope' }), e3, USER, T1), baseRev: e3.rev }])
    .catch((e) => { err = e; });
  ok(err && err.message.indexOf('fs:allow-rename') >= 0, 'flushShards rejects naming fs:allow-rename');
  eq(t3.files.get(shardPath), textBefore, 'shard on disk untouched — direct-write fallback did NOT run');
  ok(!t3.log.slice(mark3).some((l) => l.op === 'writeTextFile' && l.path === shardPath), 'no direct write of the shard logged');
  ok(!t3.files.has(shardPath + '.tmp'), 'tmp cleaned up');
  before = b3.named('applyExternalEntities').length;
  await t3.emitPaths(shardPath);
  await tick();
  eq(b3.named('applyExternalEntities').length, before, 'lastShardJson rolled back: the unchanged shard is still an echo');

  section('rename retry then in-place fallback on a sharing violation');
  const t4 = makeFakeTauri();
  const b4 = boot(t4);
  await b4.HD.createBundle(DIR, c2);
  const o4 = await b4.HD.openBundle(DIR);
  const id4 = o4.doc.items[0].id;
  const e4 = o4.envs.items.filter((e) => e.id === id4)[0];
  const realRename = t4.fs.rename;
  let renameTries = 0;
  t4.fs.rename = () => { renameTries++; return Promise.reject('The process cannot access the file because it is being used by another process. (os error 32)'); };
  const mark4 = t4.log.length;
  const r4 = await b4.HD.flushShards(DIR, o4.planId, [{ kind: 'items', id: id4, env: b4.RB.wrap(Object.assign({}, b4.RB.unwrap(e4), { notes: 'busy' }), e4, USER, T1), baseRev: e4.rev }]);
  eq(renameTries, 4, 'rename attempted 1 + 3 retries');
  eq(r4.written.length, 1, 'flush still resolves');
  const p4 = DIR + '/plans/' + o4.planId + '/items/' + id4 + '.json';
  eq(JSON.parse(t4.files.get(p4)).fields.notes, 'busy', 'in-place fallback wrote the shard');
  ok(!t4.files.has(p4 + '.tmp'), 'tmp removed after fallback');
  ok(t4.log.slice(mark4).some((l) => l.op === 'writeTextFile' && l.path === p4), 'direct write logged');
  t4.fs.rename = realRename;

  section('in-place fallback re-merges a peer shard that landed during the rename retries');
  const t5 = makeFakeTauri();
  const b5 = boot(t5);
  await b5.HD.createBundle(DIR, c2);
  const o5 = await b5.HD.openBundle(DIR);
  const id5 = o5.doc.items[0].id;
  const e5 = o5.envs.items.filter((e) => e.id === id5)[0];
  const p5 = DIR + '/plans/' + o5.planId + '/items/' + id5 + '.json';
  // the peer edits a DIFFERENT field and its write is what holds the target
  const peer5 = b5.RB.wrap(Object.assign({}, b5.RB.unwrap(e5), { size: 'L' }), e5, 'peer-zz999', T1);
  const realRename5 = t5.fs.rename;
  let tries5 = 0;
  t5.fs.rename = () => {
    tries5++;
    if (tries5 === 2) t5.files.set(p5, JSON.stringify(peer5)); // lands mid-retry
    return Promise.reject('The process cannot access the file because it is being used by another process. (os error 32)');
  };
  const mine5 = b5.RB.wrap(Object.assign({}, b5.RB.unwrap(e5), { notes: 'mine' }), e5, USER, T1);
  const r5 = await b5.HD.flushShards(DIR, o5.planId, [{ kind: 'items', id: id5, env: mine5, baseRev: e5.rev }]);
  const disk5 = JSON.parse(t5.files.get(p5));
  eq(disk5.fields.notes, 'mine', 'our field is on disk');
  eq(disk5.fields.size, 'L', 'the peer field that landed mid-retry is NOT lost');
  eq(r5.written[0].env.fields.size, 'L', 'the returned env is what landed (merged), not the stale local');
  eq(r5.written[0].canon, b5.RB.canonicalize(disk5), 'the echo key matches the merged shard on disk');
  t5.fs.rename = realRename5;
  before = b5.named('applyExternalEntities').length;
  await t5.emitPaths(p5);
  await tick();
  eq(b5.named('applyExternalEntities').length, before, 'the merged write is recognised as our own echo');

  section('pathExists');
  eq(await HD.pathExists(DIR), true, 'pathExists: the bundle folder exists');
  eq(await HD.pathExists(DIR + '/'), true, 'pathExists tolerates a trailing slash');
  eq(await HD.pathExists('C:/Users/me/OneDrive/Nope.headway'), false, 'pathExists: a missing folder is false');

  section('readPresence skips one unreadable file instead of failing the scan');
  const t6 = makeFakeTauri();
  const b6 = boot(t6);
  await b6.HD.createBundle(DIR, c2);
  await b6.HD.writePresence(DIR, 'good-aaa11', { name: 'Good', planId: 'p', editing: [], ts: 1 });
  await b6.HD.writePresence(DIR, 'bad-bbb22', { name: 'Bad', planId: 'p', editing: [], ts: 1 });
  const realRead6 = t6.fs.readTextFile;
  t6.fs.readTextFile = (p) => /bad-bbb22\.json$/.test(p) ? Promise.reject('The system cannot find the file specified. (os error 2)') : realRead6(p);
  const pres6 = await b6.HD.readPresence(DIR);
  ok(pres6['good-aaa11'] && pres6['good-aaa11'].name === 'Good', 'the readable presence file is returned');
  ok(!('bad-bbb22' in pres6), 'the unreadable one is skipped');
  t6.fs.readTextFile = realRead6;

  section('a denied read surfaces once, without burning the retry ladder');
  const t7 = makeFakeTauri();
  const b7 = boot(t7);
  await b7.HD.createBundle(DIR, c2);
  const o7 = await b7.HD.openBundle(DIR);
  const id7 = o7.doc.items[0].id;
  const p7 = DIR + '/plans/' + o7.planId + '/items/' + id7 + '.json';
  t7.deny.add('readTextFile');
  const toasts7 = b7.named('toast').length;
  const t7start = Date.now();
  await t7.emitPaths(p7);
  await tick();
  await t7.emitPaths(p7);
  await tick();
  ok(Date.now() - t7start < 1000, 'no 12 s ladder on a capability gap (took ' + (Date.now() - t7start) + ' ms)');
  const newToasts7 = b7.named('toast').slice(toasts7);
  eq(newToasts7.length, 1, 'exactly one toast for two denied events');
  ok(newToasts7.length && String(newToasts7[0].args[0]).indexOf('fs:allow-read-text-file') >= 0, 'the toast names the missing capability: ' + (newToasts7[0] && newToasts7[0].args[0]));
  ok(b7.HD.bundleWarnings().some((w) => /fs:allow-read-text-file/.test(w.err)), 'and the warning records it too');
  t7.deny.delete('readTextFile');

  section('writeHeadway / readHeadway: the plan list merges by id');
  const t8 = makeFakeTauri({ dialogOpen: () => 'C:/picked/Parent', dialogSave: (o) => 'C:/picked/' + ((o && o.defaultPath) || 'out') });
  const b8 = boot(t8);
  t8.dirs.add('C:/picked'); // the dialog only ever returns folders that exist
  ['pickFolder', 'exportBlob', 'readHeadway', 'writeHeadway'].forEach((m) => ok(typeof b8.HD[m] === 'function', 'HeadwayDesktop.' + m + ' exists'));
  await b8.HD.createBundle(DIR, c2);
  const hw0 = await b8.HD.readHeadway(DIR);
  eq(hw0.plans.length, c2.headway.plans.length, 'readHeadway parses the plan list');
  eq(await b8.HD.readHeadway('C:/nowhere/None.headway'), null, 'readHeadway of a missing folder is null');
  const entry = b8.RB.newPlanEntry('plan-b', 'Plan B', T1);
  const hw1 = await b8.HD.writeHeadway(DIR, { plans: [entry] });
  eq(hw1.plans.map((p) => p.name), c2.headway.plans.map((p) => p.name).concat('Plan B'), 'a new entry is appended, existing ones kept');
  eq([hw1.format, hw1.docId, hw1.title], ['headway-bundle-v1', c2.headway.docId, c2.headway.title], 'other keys kept');
  ok(/\n  "plans"/.test(t8.files.get(DIR + '/headway.json')), 'pretty-printed');
  ok(!t8.files.has(DIR + '/headway.json.tmp'), 'written atomically (no tmp left)');
  ok(t8.log.some((l) => l.op === 'rename' && l.to === DIR + '/headway.json'), 'tmp → rename observed');
  // a peer's concurrent entry on disk survives our write
  const disk8 = JSON.parse(t8.files.get(DIR + '/headway.json'));
  disk8.plans.push(b8.RB.newPlanEntry('plan-c', 'Peer plan', T1));
  t8.files.set(DIR + '/headway.json', JSON.stringify(disk8));
  const hw2 = await b8.HD.writeHeadway(DIR, { plans: [Object.assign({}, entry, { name: 'Plan B2', updatedAt: T2 })] });
  eq(hw2.plans.filter((p) => p.id === 'plan-b')[0].name, 'Plan B2', 'rename lands (later updatedAt wins)');
  ok(hw2.plans.some((p) => p.id === 'plan-c'), 'the peer\'s entry survives our write');
  const hw3 = await b8.HD.writeHeadway(DIR, { plans: [{ id: 'plan-b', name: 'Plan B2', deleted: true, createdAt: T1, updatedAt: T2 }] });
  eq(hw3.plans.filter((p) => p.id === 'plan-b')[0].deleted, true, 'a tombstone entry lands');
  const hw4 = await b8.HD.writeHeadway(DIR, { plans: [Object.assign({}, entry, { name: 'Late rename', updatedAt: '2026-09-01T11:00:00.000Z' })] });
  eq(hw4.plans.filter((p) => p.id === 'plan-b')[0].deleted, true, 'a later rename does not undelete');
  eq(JSON.parse(t8.files.get(DIR + '/headway.json')).plans.length, hw4.plans.length, 'what resolved is what is on disk');

  section('pickFolder / exportBlob adopt nothing');
  eq(await b8.HD.pickFolder(), 'C:/picked/Parent', 'pickFolder resolves the chosen folder');
  const o8 = await b8.HD.openBundle(DIR);
  const blob8 = { arrayBuffer: () => Promise.resolve(new Uint8Array([1, 2, 3]).buffer) };
  const ep = await b8.HD.exportBlob(blob8, 'Roadmap.xlsx', 'xlsx', 'Excel workbook');
  eq(ep, 'C:/picked/Roadmap.xlsx', 'exportBlob resolves the written path');
  eq(Array.from(t8.files.get('C:/picked/Roadmap.xlsx')), [1, 2, 3], 'bytes written');
  eq(await b8.HD.exportBlob(blob8, 'Plain', 'xlsx'), 'C:/picked/Plain.xlsx', 'the extension is appended when missing');
  eq(b8.HD.currentPath(), null, 'currentPath untouched');
  eq(b8.HD.bundleDir(), DIR, 'bundle session untouched');
  ok(t8.watching() && t8.watchOpts().recursive === true, 'the recursive bundle watch is still the active one');
  eq(b8.named('noteRecent').filter((c) => c.args[0] === 'C:/picked/Roadmap.xlsx').length, 0, 'an export is not a recent');
  const t8b = makeFakeTauri({ dialogSave: () => null, dialogOpen: () => null });
  const b8b = boot(t8b);
  eq(await b8b.HD.exportBlob(blob8, 'x.xlsx'), null, 'cancelled save → null');
  eq(await b8b.HD.pickFolder(), null, 'cancelled folder pick → null');

  section('planShardChanged for a shard of a plan we are not viewing');
  const idX = o8.doc.items[0].id;
  const otherPath = DIR + '/plans/other-plan/items/' + idX + '.json';
  t8.files.set(otherPath, t8.files.get(DIR + '/plans/' + o8.planId + '/items/' + idX + '.json'));
  before = b8.named('planShardChanged').length;
  const applied8 = b8.named('applyExternalEntities').length;
  await t8.emitPaths(otherPath);
  await tick();
  eq(b8.named('planShardChanged').length, before + 1, 'planShardChanged called once');
  eq(b8.named('planShardChanged').slice(-1)[0].args, ['other-plan'], '…with that plan\'s id');
  eq(b8.named('applyExternalEntities').length, applied8, 'and nothing applied live');
  await t8.emitPaths(DIR + '/plans/other-plan/meta.json');
  await tick();
  eq(b8.named('planShardChanged').length, before + 2, 'a meta shard of another plan reports too');
  await t8.emitPaths(DIR + '/headway.json');
  await tick();
  eq(b8.named('planShardChanged').length, before + 2, 'headway.json is not a plan shard');

  section('resumeBundle is called once desktop.js has loaded');
  const b9 = boot(makeFakeTauri(), ['resumeBundle']);
  eq(b9.named('resumeBundle').length, 1, 'HeadwayApp.resumeBundle() called at load');
  ok(b9.errors.length === 0, 'no window errors');
  const b9b = boot(makeFakeTauri());
  ok(b9b.errors.length === 0 && b9b.HD, 'an app without resumeBundle still loads (optional contract)');

  console.log('\n' + passed + ' passed, ' + failed + ' failed');
  process.exit(failed ? 1 : 0);
}

main().catch((e) => {
  console.error('  ✗ suite threw: ' + (e && e.stack || e));
  process.exit(1);
});
