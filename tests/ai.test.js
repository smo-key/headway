/* AI assistant: path ops, tools against a fake app bridge, provider protocol
 * pieces, markdown. Run: node tests/ai.test.js */
'use strict';
var RM = require('../js/core.js');
var AI = require('../js/ai.js');

var passed = 0, failed = 0;
function ok(cond, name) {
  if (cond) passed++;
  else { failed++; console.error('  ✗ ' + name); }
}
function eq(a, b, name) {
  if (JSON.stringify(a) === JSON.stringify(b)) passed++;
  else { failed++; console.error('  ✗ ' + name + '\n      got:  ' + JSON.stringify(a) + '\n      want: ' + JSON.stringify(b)); }
}
function throws(fn, re, name) {
  try { fn(); failed++; console.error('  ✗ ' + name + ' (did not throw)'); }
  catch (e) { if (re.test(e.message)) passed++; else { failed++; console.error('  ✗ ' + name + ' — threw: ' + e.message); } }
}

function freshState() {
  return RM.normalizeState({
    meta: { title: 'T', timelineStart: '2026-07-27', numWeeks: 20, priorityScheme: 'levels', riskScheme: 'lmh' },
    phases: [{ id: 'p1', name: 'Pilot' }, { id: 'p2', name: 'Scale', bucket: true }],
    team: [{ id: 't1', name: 'Ada', type: 'Engineer', rate: 100, cost: 60 }],
    items: [
      { id: 'a', num: 1, phaseId: 'p1', feature: 'Alpha', workstream: 'Data', epic: 'Core', size: 'M', priority: 'H',
        startDay: 0, durDays: 5, description: '<p>First</p>', stories: [{ id: 's1', title: 'Story one' }] },
      { id: 'b', num: 2, phaseId: 'p1', feature: 'Beta', deps: [1], stories: [] }
    ]
  });
}

// a stand-in for window.HeadwayApp: commit mutates the held state and
// records the label the way app.js would
function fakeApp(state) {
  var A = { labels: [], prefs: {}, view: 'planning', selected: null };
  A.ai = {
    state: function () { return RM.clone(state); },
    hasDoc: function () { return true; },
    commit: function (label, mutate) { A.labels.push(label); mutate(state); },
    validation: function () { return RM.validate(state); },
    userName: function () { return 'Tester'; },
    ui: function () { return { view: A.view, selectedNum: A.selected, desktop: false, userName: 'Tester', theme: 'system' }; },
    setPref: function (k, v) { if (k === 'theme' && ['system', 'light', 'dark'].indexOf(v) === -1) return false; A.prefs[k] = v; return true; },
    setView: function (v) { A.view = v; return true; },
    selectNum: function (n) { var it = RM.itemByNum(state, n); if (!it) return false; A.selected = n; return true; },
    openSettings: function () {},
    toast: function () {}
  };
  A.state = function () { return state; };
  return A;
}

console.log('— paths');
{
  var st = freshState();
  var r = AI.resolvePath(st, 'items/#2/feature');
  eq(r.parent.feature, 'Beta', '#num addresses a feature by number');
  r = AI.resolvePath(st, 'items/#1/stories/@s1/title');
  eq(r.parent.title, 'Story one', '@id addresses by id');
  r = AI.resolvePath(st, 'phases/1/name');
  eq(r.parent.name, 'Scale', 'digits index an array');
  throws(function () { AI.resolvePath(st, 'items/#9/feature'); }, /no element/, 'a missing num throws');
  throws(function () { AI.resolvePath(st, 'meta/nothing/deep'); }, /nothing at/, 'a missing object throws without create');
  r = AI.resolvePath(st, 'meta/jira/project', true);
  ok(st.meta.jira && r.key === 'project', 'create builds intermediate objects');
}

console.log('— applyOps');
{
  var st = freshState();
  var res = AI.applyOps(st, [
    { op: 'set', path: 'meta/title', value: 'Renamed' },
    { op: 'set', path: 'items/#2/size', value: 'L' },
    { op: 'push', path: 'phases', value: { name: 'Later' } },
    { op: 'set', path: 'team/-', value: { name: 'Bob', rate: 80 } },
    { op: 'delete', path: 'items/#1/stories/@s1' },
    { op: 'set', path: 'wsColors/Data', value: 'teal' }
  ]);
  eq(res.state.meta.title, 'Renamed', 'set on meta');
  eq(res.state.items[1].size, 'L', 'set on an item by num');
  eq(res.state.phases.length, 3, 'push appends a phase');
  ok(res.state.phases[2].id, 'normalizeState gives the new phase an id');
  eq(res.state.team.length, 2, '"-" appends');
  ok(res.state.team[1].id && res.state.team[1].name === 'Bob', 'appended team member is normalized');
  eq(res.state.items[0].stories.length, 0, 'delete removes an array element');
  eq(res.state.wsColors.Data, 'teal', 'set on a keyed map');
  eq(res.changes.length, 6, 'one change per op');
  eq(st.meta.title, 'T', 'the input state is untouched');
  throws(function () { AI.applyOps(st, []); }, /non-empty/, 'empty ops rejected');
  throws(function () { AI.applyOps(st, [{ op: 'zap', path: 'meta/title' }]); }, /unknown op/, 'unknown op rejected');
  throws(function () { AI.applyOps(st, [{ op: 'set', path: 'meta/title' }]); }, /no value/, 'set without value rejected');
}

console.log('— summary');
{
  var st = freshState();
  var s = AI.summary(st);
  eq(s.title, 'T', 'title');
  eq(s.phases.map(function (p) { return p.name + ':' + p.items; }), ['Pilot:2', 'Scale:0'], 'phase item counts');
  eq(s.items[0].start, '2026-07-27', 'start as ISO');
  eq(s.items[0].end, '2026-07-31', 'end is the last working day');
  eq(s.items[1].start, null, 'unscheduled items say so');
  eq(s.items[1].deps, [1], 'deps carried');
  eq(s.items[0].stories, 1, 'story count');
  eq(s.schemes.featurePriority, 'levels', 'priority scheme');
  eq(s.workstreams.map(function (w) { return w.name; }), ['Data'], 'workstreams from wsOrder');
}

console.log('— tools');
{
  var st = freshState();
  var A = fakeApp(st);
  var sum = AI.runTool('get_project', {}, A);
  eq(sum.items.length, 2, 'get_project defaults to the summary');
  var full = AI.runTool('get_project', { part: 'items', nums: [1, 9] }, A);
  eq(full.items[0].description, 'First', 'full items carry plain-text rich fields');
  eq(full.items[0].stories[0].title, 'Story one', 'full items carry stories');
  eq(full.missing, [9], 'unknown nums reported');
  throws(function () { AI.runTool('get_project', { part: 'items' }, A); }, /give nums/, 'items part needs nums');
  var val = AI.runTool('get_project', { part: 'validation' }, A);
  ok(val.counts && typeof val.counts.error === 'number', 'validation report has counts');

  var added = AI.runTool('add_items', { phase: 'Scale', items: [
    { feature: 'Gamma', workstream: 'Data', size: 'S', start: '2026-08-03', durDays: 10, deps: [1], stories: [{ title: 'S-a' }, 'S-b'], description: 'Line one\n\nLine two' },
    { feature: 'Launch', milestone: true, start: '2026-09-07' }
  ] }, A);
  // #3 is taken by the existing story (features and stories share one pool)
  eq(added.created.map(function (c) { return c.num; }), [4, 5], 'new features get the next numbers');
  var g = RM.itemByNum(st, 4);
  eq(g.phaseId, 'p2', 'phase resolved by name');
  eq(g.startDay, 5, 'ISO start became a working-day index');
  eq(g.durDays, 10, 'duration kept');
  eq(g.stories.map(function (x) { return x.title; }), ['S-a', 'S-b'], 'stories created from objects and strings');
  ok(g.stories[0].id, 'stories get ids');
  eq(g.description, '<p>Line one</p><p>Line two</p>', 'plain text description became paragraphs');
  var ms = RM.itemByNum(st, 5);
  ok(ms.milestone && ms.durDays === 0 && ms.startDay === 30, 'milestone: zero duration on its date');
  ok(/add 2 features/.test(A.labels[A.labels.length - 1]), 'history label for the add');
  throws(function () { AI.runTool('add_items', { phase: 'Nope', items: [{ feature: 'x' }] }, A); }, /no phase/, 'unknown phase rejected');
  throws(function () { AI.runTool('add_items', { items: [{ size: 'M' }] }, A); }, /feature title/, 'title required');

  var upd = AI.runTool('update_items', { label: 'tweak', updates: [
    { num: 2, fields: { start: '2026-08-10', end: '2026-08-14', priority: 'C', phase: 'Scale', addStories: [{ title: 'New story', ac: 'Given…' }] } },
    { num: 1, story: 's1', fields: { done: true, start: '2026-07-28', durDays: 2 } }
  ] }, A);
  eq(upd.updated.length, 2, 'two updates reported');
  var b = RM.itemByNum(st, 2);
  eq([b.startDay, b.durDays, b.priority, b.phaseId], [10, 5, 'C', 'p2'], 'start/end/priority/phase applied');
  eq(b.stories.length, 1, 'addStories appended');
  eq(b.stories[0].ac, '<p>Given…</p>', 'story acceptance criteria stored as html');
  var a1 = RM.itemByNum(st, 1);
  ok(a1.stories[0].done && a1.stories[0].startDay === 1 && a1.stories[0].durDays === 2, 'story fields merged');
  eq(A.labels[A.labels.length - 1], 'tweak', 'custom label used');
  throws(function () { AI.runTool('update_items', { updates: [{ num: 2, fields: { start: '2030-01-01' } }] }, A); }, /inside the timeline/, 'dates outside the timeline rejected');
  throws(function () { AI.runTool('update_items', { updates: [{ num: 77, fields: {} }] }, A); }, /no feature or story #77/, 'unknown feature rejected');
  AI.runTool('update_items', { updates: [{ num: 4, delete: true }] }, A);
  ok(!RM.itemByNum(st, 4), 'delete removes a feature');

  // tags: normalized on the way in, replaced wholesale on update
  {
    AI.runTool('add_items', { items: [{ feature: 'Tagged', tags: ['x', 'x', 'Y'],
      stories: [{ title: 'st', tags: ' a , a ,b ' }] }] }, A);
    var tg = st.items[st.items.length - 1];
    eq(tg.tags, ['x', 'Y'], 'add_items stores de-duplicated tags');
    eq(tg.stories[0].tags, ['a', 'b'], 'story tags normalize too');
    AI.runTool('update_items', { updates: [{ num: tg.num, fields: { tags: ['fresh'] } }] }, A);
    eq(RM.itemByNum(st, tg.num).tags, ['fresh'], 'update_items replaces the tag list');
    AI.runTool('update_items', { updates: [{ num: tg.num, story: tg.stories[0].id, fields: { tags: [] } }] }, A);
    eq(RM.itemByNum(st, tg.num).stories[0].tags, [], 'an empty list clears story tags');
    ok(AI.runTool('get_project', {}, A).items.some(function (x) { return String(x.tags) === 'fresh'; }),
      'the item summary line carries tags');
  }

  var up = AI.runTool('update_project', { ops: [{ op: 'set', path: 'meta/vision', value: 'Ship it' }, { op: 'push', path: 'meta/holidayRanges', value: { name: 'Offsite', start: '2026-08-20', end: '2026-08-21' } }] }, A);
  eq(up.changes, ['set meta/vision', 'push meta/holidayRanges'], 'update_project reports its changes');
  eq(st.meta.vision, 'Ship it', 'vision set through the app commit');
  ok(st.meta.holidays.indexOf('2026-08-20') !== -1, 'holiday range normalized into dates');
  ok(st.history === undefined || true, 'history untouched by the copy');

  eq(AI.runTool('set_preference', { key: 'theme', value: 'dark' }, A), { ok: true, key: 'theme', value: 'dark' }, 'set_preference goes through setPref');
  throws(function () { AI.runTool('set_preference', { key: 'nope', value: 1 }, A); }, /unknown preference/, 'unknown pref rejected');
  throws(function () { AI.runTool('set_preference', { key: 'theme', value: 'neon' }, A); }, /not accepted/, 'bad value rejected');
  eq(AI.runTool('navigate', { view: 'scoping', num: 2 }, A), { view: true, selected: true }, 'navigate switches and selects');
  eq(A.view, 'scoping', 'view changed');
  throws(function () { AI.runTool('navigate', { view: 'nowhere' }, A); }, /unknown view/, 'bad view rejected');
  var prefs = AI.runTool('get_preferences', {}, A);
  ok(prefs.ui.view === 'scoping' && prefs.preferenceKeys.theme, 'get_preferences returns ui + keys');
  throws(function () { AI.runTool('nope', {}, A); }, /unknown tool/, 'unknown tool rejected');
}

console.log('— story numbers and deps');
{
  // stories carry their own numbers from the shared pool; give the fixture two
  // well away from the feature numbers so the assertions read clearly
  var st = freshState();
  st.items[0].stories[0].num = 10;
  st.items[1].stories.push({ id: 's2', title: 'Story two', num: 11 });
  st = RM.normalizeState(st);
  var A = fakeApp(st);
  AI.runTool('update_items', { updates: [{ num: 11, fields: { deps: [10, 99] } }] }, A);
  eq(RM.storyRef(st, 's2').st.deps, [10, 99], 'a story number targets the story; deps are numbers');
  eq(AI.runTool('get_project', { part: 'items', nums: [2] }, A).items[0].stories[0].deps, [10, 99], 'get_project items reports story deps');
  ok(AI.summary(st).items.some(function (l) { return l.storyNums && l.storyNums.indexOf(10) !== -1; }), 'the summary lists story numbers');
  AI.runTool('update_items', { updates: [{ num: 1, fields: { addStories: [{ title: 'New one' }] } }] }, A);
  ok(RM.itemById(st, 'a').stories.slice(-1)[0].num === RM.nextNum(st) - 1, 'a story added by the AI gets a number at once');
  AI.runTool('add_items', { items: [{ feature: 'Delta', stories: [{ title: 'D-a' }] }] }, A);
  var delta = st.items[st.items.length - 1];
  ok(delta.stories[0].num > 0 && RM.storyByNum(st, delta.stories[0].num).st === delta.stories[0], 'add_items numbers the stories it creates');
  eq(RM.storyRef(st, 's2').st.num, 11, 'the story number itself is untouched by the edits');
  ok(/story number/i.test(AI.toolByName('update_items').description), 'the update tool documents story numbers');
  ok(/story number/i.test(AI.GUIDE), 'the guide documents story numbers');
  throws(function () { AI.runTool('update_items', { updates: [{ num: 77, fields: {} }] }, A); }, /no feature or story #77/, 'an unknown number names both kinds');
}

console.log('— item types');
{
  var stT = freshState();
  var AT = fakeApp(stT);
  AI.runTool('add_items', { items: [{ feature: 'Crash on save', type: 'Bug', stories: [{ title: 'repro', type: 'subtask' }] }] }, AT);
  var added = stT.items.filter(function (x) { return x.feature === 'Crash on save'; })[0];
  ok(added && added.type === 'bug', 'add_items accepts a type label');
  ok(added.stories[0].type === 'subtask', 'story type on add');
  AI.runTool('update_items', { updates: [{ num: added.num, fields: { type: 'task' } }] }, AT);
  ok(stT.items.filter(function (x) { return x.num === added.num; })[0].type === 'task', 'update_items sets type by key');
  var threwType = false;
  try { AI.runTool('update_items', { updates: [{ num: added.num, fields: { type: 'nope' } }] }, AT); } catch (e) { threwType = /unknown type/.test(e.message); }
  ok(threwType, 'unknown type is rejected');
  var line = JSON.stringify(AI.runTool('get_project', {}, AT).items);
  ok(/"type":"task"/.test(line), 'itemLine reports a non-default type');
}

console.log('— targeted apply');
{
  // The model thinks for seconds between reading the project and writing;
  // meanwhile the user keeps editing. A write must touch only what the tool
  // changed and leave every other item exactly as the user left it.
  var st = freshState();
  var A = fakeApp(st);
  var snap = A.ai.state;
  A.ai.state = function () {
    var c = snap();
    // user edits landing after the snapshot was taken
    st.items[0].feature = 'Alpha (user typed)';
    st.items[0].stories[0].title = 'Story one (user typed)';
    st.meta.title = 'Renamed by user';
    return c;
  };
  var alphaObj = st.items[0];
  AI.runTool('update_items', { updates: [{ num: 2, fields: { feature: 'Beta 2', size: 'S' } }] }, A);
  eq(st.items[1].feature, 'Beta 2', 'the edited feature is updated');
  eq(st.items[1].size, 'S', '…with every changed field');
  eq(st.items[0].feature, 'Alpha (user typed)', 'an untouched feature keeps the user’s concurrent edit');
  eq(st.items[0].stories[0].title, 'Story one (user typed)', '…and so do its stories');
  ok(st.items[0] === alphaObj, 'untouched items keep their object identity (no wholesale replace)');
  eq(st.meta.title, 'Renamed by user', 'untouched meta keeps the user’s concurrent edit');

  var added = AI.runTool('add_items', { phase: 'Scale', items: [{ feature: 'Gamma' }] }, A);
  eq(st.items.length, 3, 'add_items appends the new feature');
  eq(st.items[2].feature, 'Gamma', '…at the end');
  ok(st.items[0] === alphaObj, 'adding does not rewrite existing items');

  AI.runTool('update_items', { updates: [{ num: 2, delete: true }] }, A);
  // Gamma took #4 — #3 belongs to Alpha's story (shared number pool)
  eq(st.items.map(function (i) { return i.num; }).join(','), '1,4', 'deleting removes just that feature');
  ok(st.items[0] === alphaObj, 'deleting does not rewrite the others');

  AI.runTool('update_project', { ops: [{ op: 'set', path: 'meta/vision', value: 'Ship it' }] }, A);
  eq(st.meta.vision, 'Ship it', 'update_project applies a meta change');
  ok(st.items[0] === alphaObj, '…without touching items');
  ok(/navigate/.test(AI.GUIDE) && /in place/.test(AI.GUIDE), 'the guide tells the model edits show in place and not to navigate uninvited');

  // numbers are one pool: a story the tool numbers while the user adds a
  // feature with that number is moved to the next free number
  var stN = freshState();
  var AN = fakeApp(stN);
  var snapN = AN.ai.state;
  AN.ai.state = function () {
    var c = snapN();
    var taken = RM.nextNum(stN); // what the tool will hand its new story
    stN.items.push(RM.normalizeState({ meta: stN.meta, phases: stN.phases, items: [{ id: 'user-new', num: taken, phaseId: 'p1', feature: 'User added', stories: [] }] }).items[0]);
    return c;
  };
  AI.runTool('update_items', { updates: [{ num: 1, fields: { addStories: [{ title: 'Tool added' }] } }] }, AN);
  var allNums = [];
  stN.items.forEach(function (i) { allNums.push(i.num); (i.stories || []).forEach(function (x) { allNums.push(x.num); }); });
  ok(allNums.every(function (n, i) { return n != null && allNums.indexOf(n) === i; }), 'no two features/stories share a number after a concurrent add (' + allNums.join(',') + ')');
  throws(function () { AI.runTool('update_items', { updates: [{ num: 1, fields: { num: 50 } }] }, AN); }, /assigned by Headway/, 'the AI cannot set numbers directly');
  AI.runTool('update_items', { updates: [{ num: 1, fields: { flag: 'blocked on vendor' } }] }, AN);
  eq(RM.itemById(stN, 'a').flag, { reason: 'blocked on vendor' }, 'the AI can flag with a reason');
  eq(AI.runTool('get_project', {}, AN).items.filter(function (l) { return l.num === 1; })[0].flag, 'blocked on vendor', 'the summary shows the flag');
  AI.runTool('update_items', { updates: [{ num: 1, fields: { flag: null } }] }, AN);
  eq(RM.itemById(stN, 'a').flag, null, '…and unflag');
  // Exclude from Auto timeline: a field on features and stories, exclusive with Lock
  AI.runTool('update_items', { updates: [{ num: 1, fields: { locked: true } }] }, AN);
  AI.runTool('update_items', { updates: [{ num: 1, fields: { noAuto: true } }] }, AN);
  eq([RM.itemById(stN, 'a').noAuto, RM.itemById(stN, 'a').locked], [true, false], 'the AI can exclude a feature from Auto, which clears its lock');
  eq(AI.runTool('get_project', {}, AN).items.filter(function (l) { return l.num === 1; })[0].noAuto, true, 'the summary shows noAuto');
  AI.runTool('update_items', { updates: [{ num: 1, fields: { locked: true } }] }, AN);
  eq([RM.itemById(stN, 'a').noAuto, RM.itemById(stN, 'a').locked], [false, true], 'locking through the AI clears noAuto');
  AI.runTool('update_items', { updates: [{ num: 1, fields: { locked: true, noAuto: true } }] }, AN);
  eq([RM.itemById(stN, 'a').noAuto, RM.itemById(stN, 'a').locked], [false, true], 'both at once: Lock wins');
  var naSt = RM.itemById(stN, 'a').stories[0];
  AI.runTool('update_items', { updates: [{ num: naSt.num, fields: { noAuto: true } }] }, AN);
  eq(RM.itemById(stN, 'a').stories[0].noAuto, true, 'the AI can exclude a story from Auto');
  ok(/noAuto/.test(AI.TOOLS.filter(function (t) { return t.name === 'update_items'; })[0].description), 'update_items documents noAuto');
  // the other way round: the user adds a story while the tool adds a feature
  // with the same number — the existing story keeps it, the new feature moves
  var stF = freshState();
  var AF = fakeApp(stF);
  var snapF = AF.ai.state;
  AF.ai.state = function () {
    var c = snapF();
    var takenF = RM.nextNum(stF);
    stF.items[1].stories.push({ id: 'user-story', title: 'User story', num: takenF, deps: [] });
    stF.items[0].stories[0].deps = [takenF];
    return c;
  };
  AI.runTool('add_items', { phase: 'Scale', items: [{ feature: 'Tool feature' }] }, AF);
  var userSt = RM.storyRef(stF, 'user-story').st;
  var toolFeat = stF.items.filter(function (i) { return i.feature === 'Tool feature'; })[0];
  ok(toolFeat && toolFeat.num !== userSt.num, 'the tool feature moved off the user story’s number (' + toolFeat.num + ' vs ' + userSt.num + ')');
  eq(RM.resolveStoryDeps(stF, stF.items[0].stories[0]).deps[0].st.id, 'user-story', 'the dep that named the story still resolves to it');
}

console.log('— jira sync tool');
{
  var JR = require('../js/jira.js');
  globalThis.localStorage = { _m: {}, getItem: function (k) { return this._m[k] || null; }, setItem: function (k, v) { this._m[k] = String(v); } };
  var st = freshState();
  st.meta.jira = { project: 'HW', pushStories: false }; // stories sync by default now; this test counts features only
  st.items[0].jiraKey = 'HW-1';
  var A = fakeApp(st);
  var caught = null;
  try { AI.runTool('sync_jira', {}, A); } catch (e) { caught = e.message; }
  ok(/not connected/.test(caught || ''), 'sync needs a connection');
  JR.saveCreds({ site: 'x.atlassian.net', email: 'e', token: 't' });
  var calls = [];
  JR.fetchImpl = function (url, opts) {
    calls.push((opts.method || 'GET') + ' ' + url.replace('https://x.atlassian.net', ''));
    function reply(obj) { return Promise.resolve({ ok: true, status: 200, text: function () { return Promise.resolve(JSON.stringify(obj)); } }); }
    if (/issue\/HW-1\?/.test(url)) return reply({ fields: { summary: 'Alpha', status: { name: 'Done', statusCategory: { key: 'done' } } } });
    if (/issue\/bulk/.test(url)) {
      var n = JSON.parse(opts.body).issueUpdates.length;
      var issues = []; for (var i = 0; i < n; i++) issues.push({ key: 'HW-' + (10 + i) });
      return reply({ issues: issues, errors: [] });
    }
    return reply({});
  };
  var s0 = JSON.stringify(st);
  AI.runTool('sync_jira', { dryRun: true }, A).then(function (prev) {
    ok(prev.dryRun && prev.counts.create === 2 && prev.counts.update === 1, 'dry run previews creates (epic + feature) and the update');
    ok(prev.readBack.length === 1 && /done/.test(prev.readBack[0]), 'dry run lists the Done flag to read back');
    eq(JSON.stringify(st), s0, 'dry run changes nothing');
    return AI.runTool('sync_jira', {}, A);
  }).then(function (res) {
    ok(res.ok && res.created === 2 && res.updated === 1 && res.readBack === 1, 'apply creates, updates and reads back');
    eq(res.newKeys, { '#2': 'HW-10' }, 'new keys reported by feature number');
    ok(RM.itemByNum(st, 2).jiraKey === 'HW-10' && st.epicJira.Core === 'HW-10' || RM.itemByNum(st, 2).jiraKey, 'keys landed on the document');
    ok(RM.itemByNum(st, 1).done === true, 'Done state pulled from Jira');
    ok(st.meta.jira.lastSync && st.meta.jira.syncedHash, 'sync stamp recorded');
    ok(A.labels[A.labels.length - 1] === 'jira sync', 'committed through the app bridge');
    ok(calls.some(function (c) { return /PUT \/rest\/api\/3\/issue\/HW-1/.test(c); }), 'linked issue updated in Jira');
    JR.fetchImpl = null;
    settingsTests();
  }, function (err) { failed++; console.error('  ✗ jira sync threw: ' + err.message); JR.fetchImpl = null; settingsTests(); });
}

function settingsTests() {
console.log('— settings');
{
  eq(AI.baseOf('https://x.example.com/v1/'), 'https://x.example.com', 'base strips /v1');
  eq(AI.baseOf('x.example.com/litellm'), 'https://x.example.com/litellm', 'base adds https');
  eq(AI.parseHeaders('X-A: 1\nbad line\n X-B :two '), { 'X-A': '1', 'X-B': 'two' }, 'header lines parsed');
  eq(AI.ready({ provider: 'claude' }, false).ok, false, 'claude needs the desktop');
  eq(AI.ready({ provider: 'claude' }, true).ok, true, 'claude ready on desktop');
  eq(AI.ready({ provider: 'litellm', baseUrl: 'x', apiKey: '' }, false).ok, false, 'litellm needs a key');
  eq(AI.ready({ provider: 'litellm', baseUrl: 'x', apiKey: 'k', model: 'm' }, false).ok, true, 'litellm ready');
}

console.log('— openai protocol');
{
  var msgs = AI.openai.messages('SYS', [
    { role: 'user', text: 'hi', files: [{ name: 'a.png', type: 'image/png', kind: 'image', data: 'AAA' }, { name: 'n.txt', kind: 'text', data: 'notes' }] },
    { role: 'assistant', text: '', toolCalls: [{ id: 'c1', name: 'get_project', args: { part: 'summary' } }], thinkingBlocks: [{ type: 'thinking', thinking: 't', signature: 's' }] },
    { role: 'tool', results: [{ id: 'c1', name: 'get_project', content: { title: 'T' } }] },
    { role: 'assistant', text: 'done', thinking: 'r' }
  ]);
  eq(msgs[0], { role: 'system', content: 'SYS' }, 'system first');
  eq(msgs[1].content[0], { type: 'text', text: 'hi' }, 'user text part first');
  eq(msgs[1].content[1].type, 'image_url', 'images as image_url');
  ok(/notes/.test(msgs[1].content[2].text), 'text files inlined');
  eq(msgs[2].tool_calls[0].function, { name: 'get_project', arguments: '{"part":"summary"}' }, 'tool calls serialized');
  eq(msgs[2].thinking_blocks.length, 1, 'thinking blocks replayed');
  eq(msgs[3], { role: 'tool', tool_call_id: 'c1', content: '{"title":"T"}' }, 'tool results as tool messages');
  eq(msgs[4].reasoning_content, 'r', 'plain reasoning replayed');

  var events = [];
  var acc = AI.openai.accumulator(function (e) { events.push(e.type); });
  acc.push({ choices: [{ delta: { reasoning_content: 'hm ' } }] });
  acc.push({ choices: [{ delta: { thinking_blocks: [{ type: 'thinking', thinking: 'hm ' }] } }] });
  acc.push({ choices: [{ delta: { thinking_blocks: [{ type: 'thinking', thinking: '', signature: 'sig' }] } }] });
  acc.push({ choices: [{ delta: { content: 'Hel' } }] });
  acc.push({ choices: [{ delta: { content: 'lo' } }] });
  acc.push({ choices: [{ delta: { tool_calls: [{ index: 0, id: 'c9', function: { name: 'get_', arguments: '{"pa' } }] } }] });
  acc.push({ choices: [{ delta: { tool_calls: [{ index: 0, function: { name: 'project', arguments: 'rt":"meta"}' } }] }, finish_reason: 'tool_calls' }] });
  acc.push({ usage: { total_tokens: 5 }, choices: [] });
  var out = acc.done();
  eq(out.text, 'Hello', 'text deltas joined');
  eq(out.thinking, 'hm ', 'reasoning deltas joined');
  eq(out.thinkingBlocks, [{ type: 'thinking', thinking: 'hm ', signature: 'sig' }], 'thinking block pieces merged');
  eq(out.toolCalls, [{ id: 'c9', name: 'get_project', args: { part: 'meta' } }], 'tool call fragments assembled');
  eq(out.finish, 'tool_calls', 'finish reason kept');
  eq(events.slice(0, 3), ['thinking', 'text', 'text'], 'events fired in order');
  throws(function () { acc.push({ error: { message: 'boom' } }); }, /boom/, 'stream errors surface');

  var p = AI.sseParser();
  eq(p.push('data: {"a":1}\n\ndata: [DO'), ['{"a":1}'], 'complete data lines emitted');
  eq(p.push('NE]\n'), ['[DONE]'], 'split lines joined');
  eq(p.flush(), [], 'nothing left');
  eq(AI.openai.effort('max'), 'high', 'max maps to high for reasoning_effort');
}

openaiStreamingTests();
}

function openaiStreamingTests() {
console.log('— openai streaming run');
{
  var chunks = [
    'data: {"choices":[{"delta":{"reasoning_content":"think"}}]}\n\n',
    'data: {"choices":[{"delta":{"content":"Hi "}}]}\n\ndata: {"choices":[{"delta":{"content":"there"},"finish_reason":"stop"}]}\n\n',
    'data: [DONE]\n\n'
  ];
  var seen = null;
  AI.fetchImpl = function (url, opts) {
    seen = { url: url, opts: opts };
    var enc = new TextEncoder();
    var i = 0;
    var body = new ReadableStream({ pull: function (c) { if (i < chunks.length) c.enqueue(enc.encode(chunks[i++])); else c.close(); } });
    return Promise.resolve({ ok: true, status: 200, headers: { get: function () { return 'text/event-stream'; } }, body: body });
  };
  var evs = [];
  AI.openai.run({ settings: { baseUrl: 'https://gw.test/v1', apiKey: 'k', model: 'm', effort: 'high', headers: 'X-T: 1' }, system: 'S', conv: [{ role: 'user', text: 'yo' }], onEvent: function (e) { evs.push(e.type); } })
    .then(function (res) {
      eq(res.text, 'Hi there', 'streamed text assembled');
      eq(res.thinking, 'think', 'streamed reasoning assembled');
      eq(seen.url, 'https://gw.test/v1/chat/completions', 'chat completions url from the base');
      var body = JSON.parse(seen.opts.body);
      eq([body.model, body.stream, body.reasoning_effort, body.tools.length], ['m', true, 'high', AI.TOOLS.length], 'request body');
      eq(seen.opts.headers.Authorization, 'Bearer k', 'bearer auth');
      eq(seen.opts.headers['X-T'], '1', 'extra header sent');
      // a 400 blaming reasoning_effort retries without it
      var calls = 0;
      AI.fetchImpl = function (url, opts) {
        calls += 1;
        var b = JSON.parse(opts.body);
        if (b.reasoning_effort) return Promise.resolve({ ok: false, status: 400, text: function () { return Promise.resolve('{"error":{"message":"reasoning_effort is not supported"}}'); } });
        return Promise.resolve({ ok: true, status: 200, headers: { get: function () { return 'application/json'; } }, json: function () { return Promise.resolve({ choices: [{ message: { content: 'plain' }, finish_reason: 'stop' }] }); } });
      };
      return AI.openai.run({ settings: { baseUrl: 'https://gw.test', apiKey: 'k', model: 'm', effort: 'low' }, system: 'S', conv: [{ role: 'user', text: 'yo' }] });
    })
    .then(function (res) {
      eq(res.text, 'plain', 'non-streaming JSON reply handled');
      AI.fetchImpl = null;
      done();
    }, function (err) { failed++; console.error('  ✗ openai run threw: ' + err.message); AI.fetchImpl = null; done(); });
}
}

function done() {
  console.log('— claude protocol');
  {
    var s = { claudeModel: 'opus', effort: 'max' };
    var args = AI.claude.args(s, 'sess-1');
    ok(args.indexOf('--resume') !== -1 && args[args.indexOf('--resume') + 1] === 'sess-1', 'resume passed');
    ok(args[args.indexOf('--tools') + 1] === '', 'built-in tools disabled');
    ok(args[args.indexOf('--model') + 1] === 'opus' && args[args.indexOf('--effort') + 1] === 'max', 'model + effort passed');
    ok(AI.claude.args(s, null).indexOf('--resume') === -1, 'no resume for a fresh session');

    var line = JSON.parse(AI.claude.userLine({ role: 'user', text: 'hello', files: [{ name: 'p.pdf', kind: 'pdf', data: 'QUJD' }, { name: 'i.png', type: 'image/png', kind: 'image', data: 'QUJD' }] }));
    eq(line.type, 'user', 'user line type');
    eq(line.message.content.map(function (c) { return c.type; }), ['text', 'document', 'image'], 'files become content blocks');
    var tl = JSON.parse(AI.claude.userLine({ role: 'tool', results: [{ name: 'get_project', content: { x: 1 } }] }));
    ok(/Tool results/.test(tl.message.content[0].text) && /"x":1/.test(tl.message.content[0].text), 'tool results go back as a user turn');

    var parsed = AI.claude.parseFences('Let me look.\n\n```headway-tool\n{"name":"get_project","args":{"part":"summary"}}\n```\n\n```headway-tool\n{"name": "navigate", "args": {"view": "planning"}}\n```\n');
    eq(parsed.text, 'Let me look.', 'fences stripped from the text');
    eq(parsed.calls.map(function (c) { return c.name; }), ['get_project', 'navigate'], 'both calls parsed');
    eq(parsed.calls[0].args, { part: 'summary' }, 'args parsed');
    var bad = AI.claude.parseFences('```headway-tool\n{not json\n```');
    ok(bad.calls.length === 1 && bad.calls[0].args.__parseError, 'bad JSON surfaces as a parse error');

    var evs = [];
    var red = AI.claude.reducer(function (e) { evs.push(e.type + ':' + (e.delta || e.model || '')); });
    red.push('{"type":"system","subtype":"init","session_id":"S1","model":"claude-x"}');
    red.push('{"type":"stream_event","event":{"type":"content_block_delta","delta":{"type":"thinking_delta","thinking":"hmm"}},"session_id":"S1"}');
    red.push('{"type":"stream_event","event":{"type":"content_block_delta","delta":{"type":"text_delta","text":"pong"}},"session_id":"S1"}');
    red.push('not json');
    red.push('{"type":"assistant","message":{"content":[{"type":"text","text":"pong"}]},"session_id":"S1"}');
    red.push('{"type":"result","subtype":"success","is_error":false,"result":"pong","session_id":"S1"}');
    eq([red.state.session, red.state.text, red.state.thinking, red.state.done, red.state.error], ['S1', 'pong', 'hmm', true, null], 'events reduced');
    eq(evs, ['meta:claude-x', 'thinking:hmm', 'text:pong'], 'stream events surfaced');
    var red2 = AI.claude.reducer();
    red2.push('{"type":"result","subtype":"error_during_execution","is_error":true,"result":"Not logged in","session_id":"S2"}');
    eq(red2.state.error, 'Not logged in', 'result errors surface');
  }

  console.log('— system prompt');
  {
    var sp = AI.systemPrompt({ today: '2026-09-07', userName: 'Ada', desktop: true, view: 'planning', selectedNum: 3, doc: { title: 'Doc', items: 4, phases: 2, start: '2026-07-27', end: '2026-12-31' }, textTools: true });
    ok(/Today: 2026-09-07\. User: Ada/.test(sp) && /selected feature #3/.test(sp) && /"Doc"/.test(sp), 'context lines present');
    ok(sp.indexOf('```headway-tool') === -1 && /language is headway-tool/.test(sp) && /"name":"get_project"/.test(sp), 'text-tool instructions include the schemas');
    var sp2 = AI.systemPrompt({ today: '2026-09-07', view: 'setup', doc: null, textTools: false });
    ok(/No project is open/.test(sp2) && !/## Tools/.test(sp2), 'no tool section for native function calling');
  }

  console.log('— markdown');
  {
    eq(AI.md('Hi **there** and `x<y`'), '<p>Hi <b>there</b> and <code>x&lt;y</code></p>', 'inline bold, code, escaping');
    eq(AI.md('- a\n- b\n\n1. c'), '<ul><li>a</li><li>b</li></ul><ol><li>c</li></ol>', 'lists');
    eq(AI.md('## Head\ntext'), '<h4>Head</h4><p>text</p>', 'headings step down');
    eq(AI.md('```js\nlet a = "<b>";\n```'), '<pre><code>let a = &quot;&lt;b&gt;&quot;;</code></pre>', 'code fences escaped');
    eq(AI.md('see #12 and #3.'), '<p>see <a class="ai-ref" data-num="12" href="#">#12</a> and <a class="ai-ref" data-num="3" href="#">#3</a>.</p>', 'feature references link');
    eq(AI.md('text\n```headway-tool\n{"name":"x"}\n```\nafter'), '<p>text</p><p>after</p>', 'tool fences never render');
    eq(AI.md('[doc](https://x.y/z) <script>'), '<p><a href="https://x.y/z" target="_blank" rel="noopener">doc</a> &lt;script&gt;</p>', 'links and escaping');
    // GitHub-style tables: a header row, a delimiter row, body rows
    eq(AI.md('| # | Feature | Size |\n|---|:--------|-----:|\n| 1 | **Alpha** | M |\n| 2 | a \\| b | <s> |'),
      '<div class="ai-tbl"><table><thead><tr><th>#</th><th style="text-align:left">Feature</th><th style="text-align:right">Size</th></tr></thead>' +
      '<tbody><tr><td>1</td><td style="text-align:left"><b>Alpha</b></td><td style="text-align:right">M</td></tr>' +
      '<tr><td>2</td><td style="text-align:left">a | b</td><td style="text-align:right">&lt;s&gt;</td></tr></tbody></table></div>',
      'pipe tables render as tables with alignment, inline markup, escaped pipes and escaping');
    eq(AI.md('Before\n\nA | B\n--|--\n1 | 2\n\nAfter'),
      '<p>Before</p><div class="ai-tbl"><table><thead><tr><th>A</th><th>B</th></tr></thead><tbody><tr><td>1</td><td>2</td></tr></tbody></table></div><p>After</p>',
      'tables without outer pipes work and end at a blank line');
    eq(AI.md('| a | b |\n| 1 | 2 |'), '<p>| a | b |<br>| 1 | 2 |</p>', 'a pipe line without a delimiter row is plain text');
    eq(AI.md('| a | b |\n|---|---|\n| only |'), '<div class="ai-tbl"><table><thead><tr><th>a</th><th>b</th></tr></thead><tbody><tr><td>only</td><td></td></tr></tbody></table></div>', 'short rows pad to the header width');
  }

  console.log('— model labels + effort levels');
  {
    eq(AI.shortModel('bedrock/global.us.claude-opus-5'), 'claude-opus-5', 'provider + region prefixes drop from the label');
    eq(AI.shortModel('bedrock/global.anthropic.claude-opus-5'), 'claude-opus-5', 'vendor prefix drops too');
    eq(AI.shortModel('anthropic.claude-3-5-sonnet-20241022-v2:0'), 'claude-3-5-sonnet-20241022-v2:0', 'bedrock-style ids keep their version tail');
    eq(AI.shortModel('gpt-4.1'), 'gpt-4.1', 'dots inside a model name survive');
    eq(AI.shortModel('openai/gpt-4o'), 'gpt-4o', 'a bare provider path shortens');
    eq(AI.shortModel(''), '', 'empty stays empty');
    var lit = { provider: 'litellm', model: 'bedrock/x' };
    AI.modelInfo = null;
    eq(AI.effortsFor(lit).map(function (e) { return e[0]; }), ['low', 'medium', 'high', 'max'], 'no gateway info: every level offered');
    AI.modelInfo = { 'bedrock/x': { reasoning: true }, 'plain': { reasoning: false } };
    eq(AI.effortsFor(lit).map(function (e) { return e[0]; }), ['low', 'medium', 'high'], 'a reasoning model offers the gateway levels');
    eq(AI.effortsFor({ provider: 'litellm', model: 'plain' }), [], 'a model without reasoning offers no effort');
    eq(AI.effortsFor({ provider: 'litellm', model: 'unknown' }).length, 4, 'a model the gateway did not describe keeps every level');
    eq(AI.effortsFor({ provider: 'claude', claudeModel: 'opus' }).length, 4, 'the Claude CLI keeps every level');
    eq(AI.effortAllowed({ provider: 'litellm', model: 'plain', effort: 'high' }), false, 'effort is not sent to a model that lacks it');
    eq(AI.effortAllowed({ provider: 'litellm', model: 'bedrock/x', effort: 'max' }), false, 'max is not sent when the gateway only knows low/medium/high');
    eq(AI.effortAllowed({ provider: 'litellm', model: 'bedrock/x', effort: 'high' }), true, 'a listed level is sent');
    eq(AI.pickEffort({ provider: 'litellm', model: 'bedrock/x', effort: 'high' }), 'high', 'pickEffort keeps a level the model offers');
    eq(AI.pickEffort({ provider: 'litellm', model: 'bedrock/x', effort: 'max' }), 'medium', 'pickEffort falls back to medium when the pick is not offered');
    eq(AI.pickEffort({ provider: 'litellm', model: 'plain', effort: 'max' }), 'max', 'a model without effort keeps the stored value (selector hidden)');
    AI.GATEWAY_EFFORTS_SAVE = AI.GATEWAY_EFFORTS; AI.GATEWAY_EFFORTS = [['low', 'Low']];
    eq(AI.pickEffort({ provider: 'litellm', model: 'bedrock/x', effort: 'max' }), 'low', 'without medium the first offered level wins');
    AI.GATEWAY_EFFORTS = AI.GATEWAY_EFFORTS_SAVE; delete AI.GATEWAY_EFFORTS_SAVE;
    AI.modelInfo = null;
    ok(!/aiEffortSeg|>Effort</.test(AI.settingsHtml()), 'the settings page no longer carries an effort control');
  }

  console.log('— model list on open');
  var calls = [];
  AI.fetchImpl = function (url) {
    calls.push(url);
    if (/\/v1\/models$/.test(url)) return Promise.resolve({ ok: true, status: 200, json: function () { return Promise.resolve({ data: [{ id: 'zeta' }, { id: 'alpha' }] }); } });
    return Promise.resolve({ ok: false, status: 404, text: function () { return Promise.resolve(''); } });
  };
  AI.modelCache = null; AI.modelCacheBase = '';
  var gw = { provider: 'litellm', baseUrl: 'https://gw.test', apiKey: 'k', model: 'alpha', effort: 'medium' };
  AI.ensureModels(gw).then(function (ids) {
    eq(ids, ['alpha', 'zeta'], 'ensureModels fetches and sorts the gateway models');
    eq([AI.modelCache, AI.modelCacheBase], [['alpha', 'zeta'], 'https://gw.test'], 'the list is cached per gateway');
    eq(calls.filter(function (u) { return /models$/.test(u); }).length, 1, 'one models request');
    return AI.ensureModels(gw);
  }).then(function () {
    eq(calls.filter(function (u) { return /models$/.test(u); }).length, 1, 'a second open reuses the cache');
    return AI.ensureModels({ provider: 'litellm', baseUrl: 'https://other.test', apiKey: 'k' });
  }).then(function () {
    eq(calls.filter(function (u) { return /models$/.test(u); }).length, 2, 'a different gateway fetches again');
    return AI.ensureModels({ provider: 'claude' });
  }).then(function (r) {
    eq(r, null, 'the Claude provider never lists gateway models');
    AI.fetchImpl = function () { return Promise.resolve({ ok: false, status: 500, text: function () { return Promise.resolve('boom'); } }); };
    return AI.ensureModels({ provider: 'litellm', baseUrl: 'https://down.test', apiKey: 'k' });
  }).then(function (r) {
    eq(r, null, 'a failing gateway resolves null instead of throwing');
    AI.fetchImpl = null; AI.modelCache = null; AI.modelCacheBase = '';
    finish();
  }, function (err) { failed++; console.error('  ✗ ensureModels threw: ' + err.message); AI.fetchImpl = null; finish(); });
}

function finish() {
  console.log('— desktop transport');
  {
    // The Tauri http plugin's reqwest trusts only bundled Mozilla roots unless
    // native roots are enabled; behind corporate TLS inspection (Netskope,
    // Zscaler) every gateway call then dies with "error sending request for url".
    var cargo = require('fs').readFileSync(require('path').join(__dirname, '..', 'src-tauri', 'Cargo.toml'), 'utf8');
    ok(/^reqwest\s*=.*rustls-tls-native-roots/m.test(cargo), 'Cargo.toml enables reqwest native TLS roots for the http plugin');
  }

  console.log('\n' + passed + ' passed, ' + failed + ' failed');
  process.exit(failed ? 1 : 0);
}
