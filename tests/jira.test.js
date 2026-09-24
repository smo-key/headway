/* Jira sync: discover, plan and apply against a fake client. Run: node tests/jira.test.js */
'use strict';
var RM = require('../js/core.js');
var JR = require('../js/jira.js');

var passed = 0, failed = 0;
function ok(cond, name) {
  if (cond) passed++;
  else { failed++; console.error('  ✗ ' + name); }
}
function eq(a, b, name) {
  if (JSON.stringify(a) === JSON.stringify(b)) passed++;
  else { failed++; console.error('  ✗ ' + name + '\n      got:  ' + JSON.stringify(a) + '\n      want: ' + JSON.stringify(b)); }
}

var state = RM.normalizeState({
  meta: { title: 'T', timelineStart: '2026-07-27', numWeeks: 20, weeksPerSprint: 2, sprintAnchor: '2026-07-27', sprintAnchorNum: 1,
    // a project saved with the older type set (Bug and Subtask at the story level)
    itemTypes: [{ key: 'epic', label: 'Epic', icon: 'layers', jira: 'Epic' }, { key: 'feature', label: 'Feature', icon: 'square', jira: 'Story' },
      { key: 'bug', label: 'Bug', icon: 'bug', jira: 'Bug' }, { key: 'task', label: 'Task', icon: 'check-square', jira: 'Task' },
      { key: 'story', label: 'Story', icon: 'bookmark', jira: 'Sub-task' }, { key: 'subtask', label: 'Subtask', icon: 'corner-down-right', jira: 'Sub-task' }],
    hierarchy: { levels: [{ key: 'epic', types: ['epic'] }, { key: 'feature', types: ['feature', 'task'] }, { key: 'story', types: ['story', 'subtask', 'bug'] }] },
    priorityScheme: 'levels', jira: { project: 'HW', pushStories: true } },
  phases: [{ id: 'p1', name: 'Pilot Phase' }],
  team: [{ id: 'm1', name: 'Alice Rivera' }, { id: 'm2', name: 'Bob Stone' }],
  epicJira: { 'Known Epic': 'HW-1' },
  items: [
    { id: 'a', num: 1, phaseId: 'p1', feature: 'Alpha', epic: 'Known Epic', workstream: 'Data', size: 'M', priority: 'H', type: 'bug',
      startDay: 0, durDays: 5, assignees: ['m1'], description: '<p>First</p><p>Second line</p>',
      stories: [{ id: 's1', title: 'Story one', num: 10, done: false, assignees: ['m2'], type: 'subtask' }, { id: 's2', title: 'Story two', num: 11, jiraKey: 'HW-20', done: false, type: 'bug' }] },
    { id: 'b', num: 2, phaseId: 'p1', feature: 'Beta', epic: 'New Epic', jiraKey: 'HW-10', deps: [1], deadline: '2026-09-01',
      startDay: 12, durDays: 5, done: false, stories: [] },
    { id: 'c', num: 3, phaseId: 'p1', feature: 'Gamma', jiraKey: 'HW-99', stories: [] },
    { id: 'd', num: 5, phaseId: 'p1', feature: 'Delta', jiraKey: 'HW-30', done: true, startDay: 5, durDays: 5, epic: 'Known Epic',
      stories: [{ id: 's3', title: 'Story three', num: 12, deps: [11], done: true }] },
    { id: 'm', num: 4, phaseId: 'p1', feature: 'Launch', milestone: true, startDay: 15, durDays: 0, deps: [1], jiraKey: 'HW-77', stories: [] }
  ]
});
var cfg = JR.cfgOf(state, { site: 'x.atlassian.net', email: 'e', token: 't' });
var S = RM.slotsOf(state.meta);

console.log('— adf');
var doc = JR.adf('First\nsecond\n\nThird');
eq(doc.content.length, 2, 'blank lines split paragraphs');
eq(doc.content[0].content.map(function (n) { return n.type; }), ['text', 'hardBreak', 'text'], 'single newlines become hard breaks');
eq(JR.adf('').content, [], 'empty text is an empty doc');

console.log('— resolveTypes');
var projTypes = [{ name: 'Epic', hierarchyLevel: 1 }, { name: 'Story' }, { name: 'Bug' }, { name: 'Sub-task', subtask: true }];
var rt = JR.resolveTypes(projTypes, state);
eq(rt.byKey.epic.name, 'Epic', 'epic resolves by name');
eq(rt.byKey.bug.name, 'Bug', 'bug resolves by name');
eq(rt.byKey.story, { name: 'Sub-task', subtask: true }, 'story type resolves to Sub-task and is flagged subtask');
eq(rt.byKey.task.name, 'Story', 'a missing feature-level type falls back to Story');
ok(rt.notes.some(function (n) { return /Task.*not in the project.*Story/.test(n); }), 'fallback is noted');
var rt2 = JR.resolveTypes([{ name: 'Epic', hierarchyLevel: 1 }, { name: 'Task' }], state);
eq(rt2.byKey.subtask, { name: 'Task', subtask: false }, 'no subtask type in the project: story-level types fall back to Task');
eq(JR.resolveTypes([], state).known, false, 'empty project type list is unknown');

console.log('— discovery helpers');
var sf = JR.findStartField([
  { id: 'duedate', name: 'Due date', schema: { type: 'date' } },
  { id: 'customfield_10015', name: 'Start date', custom: 'com.atlassian.jira.plugin.system.customfieldtypes:datepicker', schema: { type: 'date' } }
], '');
eq(sf, { id: 'customfield_10015', name: 'Start date' }, 'Start date is found by name');
eq(JR.findStartField([], 'customfield_1'), { id: 'customfield_1', name: 'customfield_1' }, 'a configured field id wins');
eq(JR.findStartField([{ id: 'x', name: 'Other', schema: { type: 'date' } }], ''), null, 'no Start date field reads as null');
eq(JR.pickUser([{ accountId: 'u1', displayName: 'Alice Rivera' }, { accountId: 'u2', displayName: 'Alice River' }], 'alice rivera').accountId, 'u1', 'exact display-name match wins');
eq(JR.pickUser([{ accountId: 'u3', displayName: 'Bobby Stone' }], 'Bob Stone').accountId, 'u3', 'a sole candidate is accepted');
eq(JR.pickUser([{ accountId: 'u4', displayName: 'A' }, { accountId: 'u5', displayName: 'B' }], 'Zed'), null, 'ambiguous results resolve to nobody');
eq(JR.pickUser([{ accountId: 'u6', displayName: 'Rivera, Alice (Contractor)' }, { accountId: 'u7', displayName: 'Alice Park' }], 'Alice Rivera').accountId, 'u6', 'every word of the name present in one candidate wins');
eq(JR.pickUser([{ accountId: 'u8', displayName: 'A. Rivera', emailAddress: 'alice.rivera@x.com' }, { accountId: 'u9', displayName: 'Bo' }], 'alice.rivera@x.com').accountId, 'u8', 'an email matches too');
eq(JR.peopleOf(state, cfg), ['Alice Rivera', 'Bob Stone'], 'people are the assignees of features and stories, milestones excluded');

console.log('— sprints');
eq(JR.sprintOfDay(state.meta, 0), 1, 'day 0 is in sprint 1');
eq(JR.sprintOfDay(state.meta, 12), 2, 'day 12 (week 3 of 5-day weeks) is in sprint 2');
eq(JR.sprintSpan(state.meta, 2), { num: 2, name: 'Sprint 2', start: '2026-08-10', end: '2026-08-23' }, 'a sprint spans two calendar weeks');

console.log('— fields');
var info = { remote: {}, types: rt, startField: 'customfield_10015',
  accounts: { 'Alice Rivera': 'acc-alice' }, board: { id: 7, name: 'HW board', sprints: [{ id: 501, name: 'Sprint 1', startDate: '2026-07-27T09:00:00.000Z', endDate: '2026-08-09T17:00:00.000Z', state: 'active' }] }, notes: [] };
var fa = JR.featureFields(state, RM.itemById(state, 'a'), cfg, 'HW-1', info);
eq(fa.project, { key: 'HW' }, 'project key from the document mapping');
eq(fa.issuetype, { name: 'Bug' }, 'a Bug item resolves to the Bug issue type');
eq(fa.parent, { key: 'HW-1' }, 'parent is the epic key');
eq(fa.priority, { name: 'High' }, 'levels priority maps to Jira names');
ok(fa.labels.indexOf('ws-data') !== -1 && fa.labels.indexOf('phase-pilot-phase') !== -1 && fa.labels.indexOf('size-m') !== -1, 'labels carry workstream, phase and size');
eq(fa.customfield_10015, '2026-07-27', 'start date lands in the discovered Start date field');
eq(fa.duedate, '2026-07-31', 'end of the bar is the due date');
eq(fa.assignee, { accountId: 'acc-alice' }, 'the first assignee becomes the Jira assignee');
ok(!fa.description.content.some(function (p) { return p.content.some(function (n) { return /Stories:/.test(n.text); }); }),
  'story checklist stays out of the description when stories sync on their own');
var fb = JR.featureFields(state, RM.itemById(state, 'b'), cfg, null, info);
eq(fb.issuetype, { name: 'Story' }, 'a default-type feature resolves to Story');
eq(fb.duedate, '2026-09-01', 'a deadline wins as the due date');
ok(!fb.parent, 'no epic key, no parent');
ok(!fb.assignee, 'no assignee, no Jira assignee');
// tags become plain slugged labels on both levels
{
  var tSt = RM.clone(state);
  RM.setTags(tSt, RM.itemById(tSt, 'a'), ['Tech Debt', 'q3']);
  RM.setTags(tSt, RM.itemById(tSt, 'a').stories[0], ['Story Tag']);
  var tf = JR.featureFields(tSt, RM.itemById(tSt, 'a'), cfg, 'HW-1', info);
  ok(tf.labels.indexOf('tech-debt') !== -1 && tf.labels.indexOf('q3') !== -1, 'feature tags push as slugged labels');
  var ts = JR.storyFields(tSt, RM.itemById(tSt, 'a'), RM.itemById(tSt, 'a').stories[0], cfg, 'HW-5', info);
  ok(ts.labels.indexOf('story-tag') !== -1, 'story tags push as slugged labels');
}
var fs1 = JR.storyFields(state, RM.itemById(state, 'a'), RM.itemById(state, 'a').stories[0], cfg, 'HW-5', info);
eq(fs1.issuetype, { name: 'Sub-task' }, 'a Subtask story resolves to Sub-task');
eq(fs1.customfield_10015, '2026-07-27', 'a story without its own timeline takes the feature’s start');
eq(fs1.duedate, '2026-07-31', '…and the feature’s end');
ok(!fs1.assignee, 'a story assignee with no Jira match stays unassigned');
var fs2 = JR.storyFields(state, RM.itemById(state, 'a'), RM.itemById(state, 'a').stories[1], cfg, 'HW-5', info);
eq(fs2.issuetype, { name: 'Bug' }, 'a Bug story resolves to Bug');
eq(fs2.assignee, { accountId: 'acc-alice' }, 'a story with no assignee of its own inherits the feature’s');

console.log('— plan');
info.remote = { 'HW-10': { summary: 'Beta', done: true, status: 'Done' }, 'HW-20': { summary: 'Story two', done: false, status: 'To Do' }, 'HW-99': null,
  'HW-30': { summary: 'Delta', done: false, status: 'In Progress' } };
eq(JR.epicSpan(state, 'Known Epic'), { start: '2026-07-27', end: '2026-08-07' }, 'an epic spans its features’ first start to last end');
eq(JR.epicSpan(state, 'Nope'), null, 'an epic with nothing scheduled has no span');
var fe = JR.epicFields('Known Epic', cfg, info, state);
eq([fe.customfield_10015, fe.duedate], ['2026-07-27', '2026-08-07'], 'epic fields carry the rolled-up dates');
var plan = JR.plan(state, cfg, info);
eq(plan.epics.map(function (e) { return e.name; }), ['New Epic'], 'only epics without a key are created');
eq(plan.features.map(function (f) { return f.id; }), ['a'], 'features without a key are created; the milestone is not');
eq(plan.milestones.map(function (m) { return m.num + ':' + m.key; }), ['4:HW-77'], 'milestones are listed, never synced');
eq(plan.stories.map(function (s) { return s.id; }), ['s1', 's3'], 'stories without a key are created');
eq(plan.updates.map(function (u) { return u.key; }).sort(), ['HW-1', 'HW-10', 'HW-20', 'HW-30'], 'keyed items, stories and epics are updated; the milestone’s key is left alone');
var epicUpd = plan.updates.filter(function (u) { return u.kind === 'epic'; })[0];
eq([epicUpd.key, epicUpd.fields.duedate], ['HW-1', '2026-08-07'], 'a known epic gets its rolled-up dates refreshed');
eq(plan.transitions.map(function (t) { return t.kind + ':' + (t.key || t.id); }), ['feature:HW-30', 'story:s3'],
  'Headway-done issues Jira still shows open are transitioned; a new done story is too');
ok(!plan.updates[0].fields.project && !plan.updates[0].fields.issuetype, 'updates drop project and issue type');
eq(plan.pulls.map(function (p) { return p.key + ':' + p.done; }), ['HW-10:true'], 'a Done status in Jira pulls back as done');
eq(plan.missing.map(function (m) { return m.key; }), ['HW-99'], 'keys Jira does not know are reported, not updated');
eq(plan.links.filter(function (l) { return !l.story; }).length, 1, 'one feature dependency link planned (the milestone’s dependency is dropped)');
ok(plan.links.some(function (l) { return l.story && l.blockerId === 's2' && l.blockedId === 's3'; }),
  'a story dependency is planned as a link between the two stories');
eq(plan.people, { total: 2, unresolved: ['Bob Stone'] }, 'people summary names who did not match');
eq(plan.sprints.create, [{ num: 2, name: 'Sprint 2', start: '2026-08-10', end: '2026-08-23' }], 'a Headway sprint with no Jira sprint is created');
eq(plan.sprints.assign.map(function (a) { return a.ref.kind + ':' + a.ref.id + '→' + a.num + (a.sprintId ? '#' + a.sprintId : ''); }),
  ['feature:a→1#501', 'story:s2→1#501', 'feature:b→2', 'feature:d→1#501'],
  'scheduled features are placed; sub-task stories follow their parent; the Bug story (not a subtask) is placed by its feature’s day; the milestone is skipped');
eq(plan.counts, { create: 4, update: 4, link: 2, pull: 1, missing: 1, done: 2, sprintsCreate: 1, sprintAssign: 4 }, 'counts summarise the plan');
eq(JR.keysOf(state, cfg), ['HW-20', 'HW-10', 'HW-99', 'HW-30'], 'keysOf lists every linked key except milestones');
var planNoBoard = JR.plan(state, cfg, { remote: info.remote, types: info.types, startField: null, accounts: {}, board: null });
eq(planNoBoard.sprints, null, 'no board, no sprint work');
ok(!planNoBoard.features[0].fields.customfield_10015, 'no Start date field, no start date');

var pa = plan.features.filter(function (f) { return f.id === 'a'; })[0];
eq(pa.fields.issuetype.name, 'Bug', 'a Bug feature is created as Bug');
var ps1 = plan.stories.filter(function (s) { return s.id === 's1'; })[0];
eq(ps1.fields.issuetype.name, 'Sub-task', 'a Subtask story is created as Sub-task');
ok(ps1.nest === true, 'subtask stories nest');

console.log('— apply');
var calls = [];
var nextKey = 100;
var client = {
  get: function (p) {
    calls.push(['GET', p]);
    if (/\/transitions$/.test(p)) return Promise.resolve({ transitions: [{ id: '11', to: { statusCategory: { key: 'indeterminate' } } }, { id: '31', to: { statusCategory: { key: 'done' } } }] });
    return Promise.resolve({});
  },
  post: function (p, b) {
    calls.push(['POST', p, b]);
    if (p === '/rest/api/3/issue/bulk') {
      return Promise.resolve({ issues: b.issueUpdates.map(function () { return { key: 'HW-' + (nextKey++) }; }), errors: [] });
    }
    if (p === '/rest/agile/1.0/sprint') return Promise.resolve({ id: 777, name: b.name });
    return Promise.resolve({});
  },
  put: function (p, b) { calls.push(['PUT', p, b]); return Promise.resolve(null); }
};
JR.apply(plan, client, null).then(function (result) {
  eq(result.epicKeys, { 'New Epic': 'HW-100' }, 'epic created first');
  eq(result.itemKeys, { a: 'HW-101' }, 'feature keyed');
  eq(result.storyKeys, { s1: 'HW-102', s3: 'HW-103' }, 'stories keyed');
  var creates = calls.filter(function (c) { return c[1] === '/rest/api/3/issue/bulk'; });
  eq(creates.length, 3, 'three bulk-create rounds: epics, features, stories');
  eq(creates[2][2].issueUpdates[0].fields.parent, { key: 'HW-101' }, 'the new story parents to the freshly created feature');
  var upd = calls.filter(function (c) { return c[0] === 'PUT'; });
  eq(upd.map(function (c) { return c[1]; }).sort(), ['/rest/api/3/issue/HW-1', '/rest/api/3/issue/HW-10', '/rest/api/3/issue/HW-20', '/rest/api/3/issue/HW-30'], 'linked issues and the epic are updated in place');
  var trans = calls.filter(function (c) { return c[0] === 'POST' && /\/transitions$/.test(c[1]); }).map(function (c) { return c[1].replace(/.*issue\//, '') + '=' + c[2].transition.id; }).sort();
  eq(trans, ['HW-103/transitions=31', 'HW-30/transitions=31'], 'done issues take the transition that leads to the Done category');
  eq(result.transitioned, 2, 'transition count');
  var updB = upd.filter(function (c) { return c[1] === '/rest/api/3/issue/HW-10'; })[0];
  eq(updB[2].fields.parent, { key: 'HW-100' }, 'an updated feature gets its newly created epic as parent');
  var links = calls.filter(function (c) { return c[1] === '/rest/api/3/issueLink'; });
  eq(links.length, 2, 'the feature dependency and the story dependency each became a link');
  eq(links[0][2].outwardIssue, { key: 'HW-101' }, 'the dependency (blocker) is the outward issue');
  eq(links[0][2].inwardIssue, { key: 'HW-10' }, 'the dependent is the inward issue');
  ok(links.some(function (c) { return JSON.stringify(c[2].outwardIssue) === '{"key":"HW-20"}' && JSON.stringify(c[2].inwardIssue) === '{"key":"HW-103"}'; }),
    'the story link resolves the keyed blocker and the freshly created dependent');
  var sprintCreate = calls.filter(function (c) { return c[1] === '/rest/agile/1.0/sprint'; });
  eq(sprintCreate.length, 1, 'one sprint created');
  eq(sprintCreate[0][2].originBoardId, 7, 'on the project’s board');
  ok(/^2026-08-10T/.test(sprintCreate[0][2].startDate) && /^2026-08-23T/.test(sprintCreate[0][2].endDate), 'with Headway’s sprint dates');
  var moves = calls.filter(function (c) { return /\/rest\/agile\/1.0\/sprint\/\d+\/issue$/.test(c[1]); }).map(function (c) { return c[1].replace(/.*sprint\//, '') + ':' + c[2].issues.join(','); }).sort();
  eq(moves, ['501/issue:HW-101,HW-20,HW-30', '777/issue:HW-10'], 'issues land in the matching existing sprint and the newly created one');
  eq(result.sprintsCreated, 1, 'sprint created count');
  eq(result.sprintAssigned, 4, 'sprint placement count');
  eq(result.created, 4, 'created count');
  eq(result.updated, 4, 'updated count');
  eq(result.errors, [], 'no errors');

  var s2 = RM.clone(state);
  JR.applyToState(s2, plan, result, info);
  eq(RM.itemById(s2, 'a').jiraKey, 'HW-101', 'feature key lands on the item');
  eq(RM.itemById(s2, 'a').stories[0].jiraKey, 'HW-102', 'story key lands on the story');
  eq(s2.epicJira['New Epic'], 'HW-100', 'epic key lands in epicJira');
  eq(RM.itemById(s2, 'b').done, true, 'done pulled from Jira');
  eq(s2.meta.jira.accounts, { 'Alice Rivera': 'acc-alice' }, 'matched accounts are remembered in the document');
  var s6 = RM.clone(state); s6.meta.jira.accounts = { 'Bob Stone': 'acc-bob' };
  eq(JR.plan(s6, cfg, { remote: {}, types: info.types, startField: null, accounts: { 'Bob Stone': 'acc-bob', 'Alice Rivera': null }, board: null }).stories[0].fields.assignee,
    { accountId: 'acc-bob' }, 'a hand-picked account from Setup is used as the assignee');

  // a field Jira refuses is dropped and the batch retried once
  var calls3 = [];
  var client3 = {
    post: function (p, b) {
      calls3.push([p, b]);
      if (p === '/rest/api/3/issue/bulk') {
        var bad = b.issueUpdates.some(function (u) { return u.fields.customfield_10015; });
        if (bad) {
          return Promise.resolve({ issues: [], errors: b.issueUpdates.map(function (u, i) {
            return { failedElementNumber: i, elementErrors: { errorMessages: [], errors: { customfield_10015: "Field 'customfield_10015' cannot be set. It is not on the appropriate screen, or unknown." } } };
          }) });
        }
        return Promise.resolve({ issues: b.issueUpdates.map(function (x, i) { return { key: 'Y-' + i }; }), errors: [] });
      }
      return Promise.resolve({ id: 1 });
    },
    put: function () { return Promise.resolve(null); },
    get: function (p) { return Promise.resolve(/transitions$/.test(p) ? { transitions: [{ id: '31', to: { statusCategory: { key: 'done' } } }] } : {}); }
  };
  var plan3 = JR.plan(state, cfg, info);
  return JR.apply(plan3, client3, null).then(function (r3) {
    eq(r3.droppedFields, ['customfield_10015'], 'the refused field is reported');
    eq(r3.itemKeys, { a: 'Y-0' }, 'the feature is created on the retry without it');
    ok(r3.errors.length === 0, 'a dropped field is not an error by itself');
    var bulk = calls3.filter(function (c) { return c[0] === '/rest/api/3/issue/bulk'; });
    ok(bulk.length === 4, 'epics, features (twice: refused, then retried), stories');
    ok(!bulk[3][1].issueUpdates[0].fields.customfield_10015, 'later rounds skip the refused field too');
  }).then(function () {
    // a failed element in a bulk response is reported and skipped
    var client2 = {
      post: function (p, b) {
        if (p === '/rest/api/3/issue/bulk') {
          if (b.issueUpdates[0].fields.issuetype.name === 'Epic') {
            return Promise.resolve({ issues: [], errors: [{ failedElementNumber: 0, elementErrors: { errorMessages: ['nope'], errors: {} } }] });
          }
          return Promise.resolve({ issues: b.issueUpdates.map(function (x, i) { return { key: 'X-' + i }; }), errors: [] });
        }
        return Promise.resolve({ id: 2 });
      },
      put: function () { return Promise.resolve(null); }, get: function () { return Promise.resolve({}); }
    };
    var plan2 = JR.plan(state, cfg, info);
    return JR.apply(plan2, client2, null).then(function (r2) {
      ok(r2.errors.some(function (e) { return /New Epic: nope/.test(e); }), 'bulk element errors name the failed entry');
      eq(r2.epicKeys, {}, 'the failed epic gets no key');
      eq(r2.itemKeys, { a: 'X-0' }, 'later rounds still create');
    });
  });
}).then(function () {
  console.log('— discover');
  var seen = [];
  var dclient = {
    get: function (p) {
      seen.push(p);
      if (/^\/rest\/api\/3\/project\//.test(p)) return Promise.resolve({ issueTypes: [{ name: 'Epic', hierarchyLevel: 1 }, { name: 'Story' }, { name: 'Bug' }, { name: 'Sub-task', subtask: true }] });
      if (p === '/rest/api/3/field') return Promise.resolve([{ id: 'customfield_10015', name: 'Start date', custom: 'x', schema: { type: 'date' } }]);
      if (/user\/assignable/.test(p)) return Promise.resolve(/Alice/.test(decodeURIComponent(p)) ? [{ accountId: 'acc-alice', displayName: 'Alice Rivera' }] : []);
      if (/\/rest\/agile\/1.0\/board\?/.test(p)) return Promise.resolve({ values: [{ id: 7, name: 'HW board' }] });
      if (/\/board\/7\/sprint/.test(p)) return Promise.resolve({ values: [{ id: 501, name: 'Sprint 1', startDate: '2026-07-27T00:00:00Z', endDate: '2026-08-09T00:00:00Z' }], isLast: true });
      if (/\/rest\/api\/3\/issue\/HW-99/.test(p)) { var e = new Error('nf'); e.status = 404; return Promise.reject(e); }
      if (/\/rest\/api\/3\/issue\//.test(p)) return Promise.resolve({ fields: { summary: 's', status: { name: 'Done', statusCategory: { key: 'done' } } } });
      return Promise.resolve({});
    }
  };
  return JR.discover(dclient, state, cfg, null).then(function (inf) {
    eq(inf.types.byKey.story.name, 'Sub-task', 'types come from the project');
    eq(inf.startField, 'customfield_10015', 'the Start date field is found');
    eq(inf.accounts, { 'Alice Rivera': 'acc-alice', 'Bob Stone': null }, 'people are looked up per name');
    eq(inf.accountNames['Alice Rivera'], 'Alice Rivera', 'the matched display name is kept for the settings page');
    eq(inf.board.id, 7, 'the Scrum board is found');
    eq(inf.board.sprints.length, 1, 'its sprints are listed');
    eq(inf.remote['HW-99'], null, 'a missing key reads as null');
    eq(inf.remote['HW-10'].done, true, 'a linked issue’s done state is read');
    ok(!seen.some(function (p) { return /HW-77/.test(p); }), 'the milestone’s key is never fetched');
  });
}).then(function () {
  console.log('— client');
  eq(JR.siteUrl('x.atlassian.net/'), 'https://x.atlassian.net', 'site url gets a scheme and loses the slash');
  var seen = [];
  JR.fetchImpl = function (url, opts) {
    seen.push([url, opts]);
    return Promise.resolve({ ok: false, status: 400, text: function () { return Promise.resolve('{"errorMessages":["bad"],"errors":{"summary":"required"}}'); } });
  };
  return JR.client({ site: 'x.atlassian.net', email: 'e', token: 't' }).post('/rest/api/3/issue', { a: 1 }).then(function () {
    ok(false, 'a 400 rejects');
  }, function (err) {
    eq(err.message, 'bad; summary: required', 'error messages come from the Jira body');
    eq(err.fields, ['summary'], 'the failing fields are named');
    eq(seen[0][0], 'https://x.atlassian.net/rest/api/3/issue', 'request goes to the site');
    ok(/^Basic /.test(seen[0][1].headers.Authorization), 'basic auth header');
  });
}).then(function () {
  console.log('— stories without a sub-task type');
  var flat = { remote: info.remote, types: { byKey: { epic: { name: 'Epic', subtask: false }, feature: { name: 'Story', subtask: false }, bug: { name: 'Bug', subtask: false }, task: { name: 'Task', subtask: false }, story: { name: 'Story', subtask: false }, subtask: { name: 'Story', subtask: false } }, known: true, notes: [] }, startField: null, accounts: {}, board: null, notes: [] };
  var pf = JR.plan(state, cfg, flat);
  ok(pf.stories.every(function (st) { return !st.fields.parent && st.nest === false; }), 'stories beside their feature carry no parent');
  eq(pf.storyLinks.length, 3, 'every unnested story (new or already keyed) is noted for a link to its feature');
  ok(pf.notes.some(function (n) { return /not a sub-task/.test(n); }), 'the preview explains why');

  console.log('— mixed nesting');
  var mixed = { remote: info.remote, types: rt, startField: null, accounts: {}, board: null, notes: [] };
  var pm = JR.plan(state, cfg, mixed);
  var bugStory = pm.updates.filter(function (u) { return u.id === 's2'; })[0];
  ok(bugStory && bugStory.kind === 'story', 'the Bug story (non-subtask) is planned');
  ok(pm.storyLinks.some(function (l) { return l.storyId === 's2'; }), 'a non-subtask story gets a feature link instead of a parent');
  ok(!pm.storyLinks.some(function (l) { return l.storyId === 's1'; }), 'a subtask story does not');

  var callsF = [];
  var nkF = 700;
  var clientF = {
    get: function (p) { return Promise.resolve(/transitions$/.test(p) ? { transitions: [{ id: '31', to: { statusCategory: { key: 'done' } } }] } : {}); },
    post: function (p, b) {
      callsF.push([p, b]);
      if (p === '/rest/api/3/issue/bulk') return Promise.resolve({ issues: b.issueUpdates.map(function () { return { key: 'F-' + (nkF++) }; }), errors: [] });
      return Promise.resolve({ id: 1 });
    },
    put: function () { return Promise.resolve(null); }
  };
  return JR.apply(pf, clientF, null).then(function (rf) {
    ok(Object.keys(rf.storyKeys).length === 2, 'both stories are created without a parent');
    var rel = callsF.filter(function (c) { return c[0] === '/rest/api/3/issueLink' && c[1].type.name === 'Relates'; });
    eq(rel.length, 2, 'each story is linked to its feature with Relates');
    eq(rf.storyLinked, 2, 'story link count');
    ok(!rf.errors.length, 'no errors');
  }).then(function () {
    // Jira refusing the parent at create time: retried without it, then linked
    var callsR = [];
    var nkR = 800;
    var clientR = {
      get: function (p) { return Promise.resolve(/transitions$/.test(p) ? { transitions: [{ id: '31', to: { statusCategory: { key: 'done' } } }] } : {}); },
      post: function (p, b) {
        callsR.push([p, b]);
        if (p === '/rest/api/3/issue/bulk') {
          var issues = [], errors = [];
          b.issueUpdates.forEach(function (x, i) {
            if (x.fields.parent && x.fields.issuetype.name === 'Sub-task') errors.push({ status: 400, failedElementNumber: i, elementErrors: { errorMessages: [], errors: { parent: 'Please select valid parent issue.', parentId: 'Please select valid parent issue.' } } });
            else issues.push({ key: 'R-' + (nkR++) });
          });
          return Promise.resolve({ issues: issues, errors: errors });
        }
        return Promise.resolve({ id: 1 });
      },
      put: function () { return Promise.resolve(null); }
    };
    var pr = JR.plan(state, cfg, info);
    return JR.apply(pr, clientR, null).then(function (rr) {
      ok(Object.keys(rr.storyKeys).length === 2, 'stories Jira would not nest are created on a second try without the parent');
      var rel = callsR.filter(function (c) { return c[0] === '/rest/api/3/issueLink' && c[1].type.name === 'Relates'; });
      eq(rel.length, 2, 'and linked to their features');
      ok(rr.errors.length === 1 && /could not nest/.test(rr.errors[0]), 'one summary problem explains the fallback');
      ok(!rr.errors.some(function (e) { return /Please select valid parent/.test(e); }), 'the raw parent errors are not repeated per story');
    });
  });
}).then(function () {
  console.log('— error text');
  eq(JR.errText({ a: 1 }), '{"a":1}', 'an object error is shown as JSON, never [object Object]');
  eq(JR.errText('plain'), 'plain', 'a string error passes through');
  eq(JR.errText(new Error('boom')), 'boom', 'an Error shows its message');
  eq(JR.bodyMessages({ errorMessages: [], errors: [{ status: 400, failedElementNumber: 0, elementErrors: { errorMessages: [], errors: { customfield_10015: 'Field cannot be set' } } }] }),
    ['item 1: customfield_10015: Field cannot be set'], 'a bulk 400 body lists each element’s errors');
  eq(JR.bodyMessages({ errors: { summary: { nested: true } } }), ['summary: {"nested":true}'], 'object-valued errors are still readable');
  // a transport that rejects with a bare string (the Tauri plugin does) becomes an Error
  JR.fetchImpl = function () { return Promise.reject('url not allowed on the configured scope: https://x'); };
  return JR.client({ site: 'x.atlassian.net', email: 'e', token: 't' }).get('/rest/api/3/myself').then(function () {
    ok(false, 'a transport rejection rejects');
  }, function (err) {
    ok(err instanceof Error && /not allowed/.test(err.message), 'transport strings become Errors with their text');
  }).then(function () {
    // a bulk create where every element fails comes back as a 400 with the
    // errors[] body: the refused field is dropped and the batch retried
    var seenBulk = 0;
    JR.fetchImpl = function (url, opts) {
      if (/issue\/bulk$/.test(url)) {
        seenBulk += 1;
        var b = JSON.parse(opts.body);
        var bad = b.issueUpdates.some(function (x) { return x.fields.customfield_10015; });
        if (bad) {
          return Promise.resolve({ ok: false, status: 400, text: function () { return Promise.resolve(JSON.stringify({ issues: [], errors: b.issueUpdates.map(function (x, i) {
            return { status: 400, failedElementNumber: i, elementErrors: { errorMessages: [], errors: { customfield_10015: "Field 'customfield_10015' cannot be set. It is not on the appropriate screen, or unknown." } } };
          }) })); } });
        }
        return Promise.resolve({ ok: true, status: 201, text: function () { return Promise.resolve(JSON.stringify({ issues: b.issueUpdates.map(function (x, i) { return { key: 'Z-' + i }; }), errors: [] })); } });
      }
      if (/\/transitions$/.test(url) && opts.method === 'GET') return Promise.resolve({ ok: true, status: 200, text: function () { return Promise.resolve('{"transitions":[{"id":"31","to":{"statusCategory":{"key":"done"}}}]}'); } });
      return Promise.resolve({ ok: true, status: 200, text: function () { return Promise.resolve('{"id":5}'); } });
    };
    var c = JR.client({ site: 'x.atlassian.net', email: 'e', token: 't' });
    var p4 = JR.plan(state, cfg, info);
    return JR.apply(p4, c, null).then(function (r4) {
      eq(r4.droppedFields, ['customfield_10015'], 'a 400-with-elements bulk failure still drops the refused field');
      ok(r4.itemKeys.a === 'Z-0', 'and the retry creates the feature');
      ok(!r4.errors.some(function (e) { return /object Object/.test(e); }), 'no [object Object] anywhere in the problems');
      ok(seenBulk >= 3, 'epics, features (refused then retried), stories');
      JR.fetchImpl = null;
    });
  });
}).then(function () {
  console.log('— status');
  var f1 = JR.fingerprint(state);
  eq(JR.fingerprint(RM.clone(state)), f1, 'fingerprint is deterministic');
  var s3 = RM.clone(state); s3.items[0].feature = 'Alpha renamed';
  ok(JR.fingerprint(s3) !== f1, 'a title change alters the fingerprint');
  var s5 = RM.clone(state); s5.items[4].feature = 'Launch moved';
  eq(JR.fingerprint(s5), f1, 'milestone edits do not count as sync changes');
  var s4 = RM.clone(state); s4.meta.jira.lastSync = '2026-09-07T00:00:00Z';
  eq(JR.fingerprint(s4), f1, 'sync bookkeeping in meta.jira does not alter the fingerprint');
  eq(JR.status(state).kind, 'off', 'no credentials on this machine reads as off');
  JR.hasCreds = function () { return true; };
  eq(JR.status(state).kind, 'never', 'credentials + project but no sync yet');
  s4.meta.jira.syncedHash = f1;
  eq(JR.status(s4).kind, 'clean', 'matching fingerprint reads as clean');
  s4.items[0].done = true;
  eq(JR.status(s4).kind, 'dirty', 'an edit after the sync reads as dirty');
  s4.meta.jira.lastErrors = 2;
  eq(JR.status(s4).kind, 'error', 'problems in the last sync read as error');
  eq(JR.status(s4).linked, 3, 'linked feature count excludes the milestone');
}).then(function () {
  console.log('— jira csv story deps');
  var JX = require('../js/export-jira.js');
  var xrows = JX.rows(state, { features: true, stories: true });
  var xs3 = xrows.filter(function (r) { return r['Summary'] === 'Story three'; })[0];
  eq(xs3['Blocked By'], 'HW-20', 'a story row lists the Jira keys of its story dependencies');
  var xs1 = xrows.filter(function (r) { return r['Summary'] === 'Story one'; })[0];
  eq(xs1['Blocked By'], '', 'a story with no dependencies leaves Blocked By blank');
}).then(function () {
  console.log('— auto-sync');
  var base = RM.clone(state);
  base.meta.jira = { project: 'HW', pushStories: true, site: 'https://x.atlassian.net' };
  JR.hasCreds = function () { return true; };
  var unlinked = RM.clone(base);
  unlinked.epicJira = {};
  unlinked.items.forEach(function (it) { it.jiraKey = null; it.stories.forEach(function (st) { st.jiraKey = null; }); });
  eq(JR.autoDue(unlinked), null, 'nothing linked to Jira yet: nothing is due');
  ok(!!JR.autoDue(base), 'one linked item is enough, even without a completed sync');
  base.meta.jira.lastSync = '2026-09-07T00:00:00Z';
  base.meta.jira.syncedHash = JR.fingerprint(base);
  eq(JR.autoDue(base), null, 'in sync: nothing is due');
  base.items[0].feature = 'Alpha v2';
  var due = JR.autoDue(base);
  ok(due && due === JR.fingerprint(base), 'a change after a successful sync is due');
  base.meta.jira.auto = false;
  eq(JR.autoDue(base), null, 'the toggle turns it off');
  base.meta.jira.auto = true;
  JR.autoFailedHash = due;
  eq(JR.autoDue(base), null, 'a failed automatic run waits for the next change');
  JR.autoFailedHash = null;
  JR.job = { progress: 'x' };
  eq(JR.autoDue(base), null, 'nothing is due while a sync runs');
  JR.job = null;

  // a whole automatic run through a fake app + fake Jira
  var doc = base, labels = [], toasts = [];
  globalThis.HeadwayApp = { ai: { hasDoc: function () { return true; }, state: function () { return RM.clone(doc); },
    commit: function (label, fn) { labels.push(label); var s = RM.clone(doc); fn(s); doc = s; } },
    toast: function (m) { toasts.push(m); } };
  JR.loadCreds = function () { return { email: 'e', token: 't' }; };
  var nk = 300;
  JR.fetchImpl = function (url, opts) {
    var path = url.replace('https://x.atlassian.net', '');
    function reply(status, body) { return Promise.resolve({ ok: status < 300, status: status, text: function () { return Promise.resolve(body == null ? '' : JSON.stringify(body)); } }); }
    if (/^\/rest\/api\/3\/project\//.test(path)) return reply(200, { issueTypes: [{ name: 'Epic', hierarchyLevel: 1 }, { name: 'Story' }, { name: 'Bug' }, { name: 'Sub-task', subtask: true }] });
    if (path === '/rest/api/3/field') return reply(200, [{ id: 'customfield_10015', name: 'Start date', schema: { type: 'date' } }]);
    if (/user\/assignable/.test(path)) return reply(200, []);
    if (/\/rest\/agile\/1.0\/board\?/.test(path)) return reply(200, { values: [] });
    if (/\/rest\/api\/3\/issue\/HW-99/.test(path)) return reply(404, { errorMessages: ['nope'] });
    if (/\/transitions$/.test(path) && opts.method === 'GET') return reply(200, { transitions: [{ id: '31', to: { statusCategory: { key: 'done' } } }] });
    if (opts.method === 'GET' && /\/rest\/api\/3\/issue\//.test(path)) return reply(200, { fields: { summary: 's', status: { name: 'To Do', statusCategory: { key: 'new' } } } });
    if (/issue\/bulk$/.test(path)) { var b = JSON.parse(opts.body); return reply(201, { issues: b.issueUpdates.map(function () { return { key: 'HW-' + (nk++) }; }), errors: [] }); }
    if (opts.method === 'PUT') return reply(204, null);
    return reply(201, null);
  };
  return JR.autoTick().then(function (ran) {
    ok(ran === true, 'the tick ran a sync');
    ok(labels[labels.length - 1] === 'jira sync', 'the result was committed');
    ok(RM.itemById(doc, 'a').jiraKey, 'the unlinked feature got its key');
    eq(doc.meta.jira.syncedHash, JR.fingerprint(doc), 'the document is in sync afterwards');
    eq(toasts.length, 0, 'a clean automatic run says nothing');
    eq(JR.autoDue(doc), null, 'nothing more is due');
    return JR.autoTick();
  }).then(function (ran2) {
    ok(ran2 === false, 'the next tick does nothing');
    JR.fetchImpl = function () { return Promise.reject('network down'); };
    doc.items[0].feature = 'Alpha v3';
    return JR.autoTick();
  }).then(function (ran3) {
    ok(ran3 === false && /auto-sync failed: network down/.test(toasts[toasts.length - 1] || ''), 'a failed automatic run reports the real error');
    eq(JR.autoDue(doc), null, '…and does not retry until the next change');
    doc.items[0].feature = 'Alpha v4';
    ok(!!JR.autoDue(doc), 'a new change makes it due again');
    JR.fetchImpl = null;
    delete globalThis.HeadwayApp;
  });
}).then(function () {
  console.log(passed + ' passed, ' + failed + ' failed');
  process.exit(failed ? 1 : 0);
}, function (e) { console.error(e); process.exit(1); });
