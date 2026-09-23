/* Node test suite for Headway core + Excel round-trip.
 * Run:  node tools/roadmapping/tests/core.test.js
 * (Excel tests need `exceljs` resolvable via NODE_PATH; they self-skip otherwise.)
 */
'use strict';
var RM = require('../js/core.js');

var passed = 0, failed = 0, skipped = 0;
function ok(cond, name) {
  if (cond) { passed++; }
  else { failed++; console.error('  ✗ ' + name); }
}
function eq(a, b, name) {
  var good = JSON.stringify(a) === JSON.stringify(b);
  if (good) passed++;
  else { failed++; console.error('  ✗ ' + name + '\n      got:  ' + JSON.stringify(a) + '\n      want: ' + JSON.stringify(b)); }
}
function section(name) { console.log('— ' + name); }

var META = {
  title: 'T', timelineStart: '2026-07-27', numWeeks: 48,
  weeksPerSprint: 2,
  capacityEnabled: true, // capacity tests exercise the roster constraints
  // two full holiday weeks (Nov 16–20 and Nov 23–27), as individual dates
  holidays: [
    '2026-11-16', '2026-11-17', '2026-11-18', '2026-11-19', '2026-11-20',
    '2026-11-23', '2026-11-24', '2026-11-25', '2026-11-26', '2026-11-27'
  ],
  sizeDays: { XS: 2, S: 3, M: 5, L: 10, XL: 20 }
};

function mkState(items, extras) {
  var base = {
    meta: JSON.parse(JSON.stringify(META)),
    phases: [
      { id: 'p1', name: 'Alpha', bucket: false },
      { id: 'p2', name: 'Next', bucket: true }
    ],
    items: items,
    team: [],
    teamTypes: ['Development', 'Data']
  };
  if (extras && extras.meta) { Object.keys(extras.meta).forEach(function (k) { base.meta[k] = extras.meta[k]; }); delete extras.meta; }
  if (extras) Object.keys(extras).forEach(function (k) { base[k] = extras[k]; });
  return RM.normalizeState(base);
}

// ------------------------------------------------------------- calendar
section('calendar');
ok(RM.fmtISO(RM.dayToDate(META, 0)) === '2026-07-27', 'day 0 is timeline start');
ok(RM.fmtISO(RM.dayToDate(META, 4)) === '2026-07-31', 'day 4 is Friday of week 0');
ok(RM.fmtISO(RM.dayToDate(META, 5)) === '2026-08-03', 'day 5 skips the weekend');
eq(RM.dateToDay(META, RM.parseISO('2026-08-03')), 5, 'dateToDay Monday week 1');
eq(RM.dateToDay(META, RM.parseISO('2026-08-01')), 4, 'Saturday snaps to Friday');
eq(RM.dateToDay(META, RM.parseISO('2026-07-27')), 0, 'roundtrip day 0');
for (var d = 0; d < 60; d += 7) {
  var dd = RM.dateToDay(META, RM.dayToDate(META, d));
  if (dd !== d) { ok(false, 'roundtrip day ' + d); break; }
}
ok(true, 'roundtrip 0..60');
ok(RM.isBlackoutWeek(META, 16), 'week of 2026-11-16 is blackout');
ok(RM.isBlackoutWeek(META, 17), 'week of 2026-11-23 is blackout');
ok(!RM.isBlackoutWeek(META, 15), 'week 15 is not blackout');

// stretchSpan: 10 working days starting week 15 must stretch over 2 blackout weeks
eq(RM.stretchSpan(META, 15 * 5, 10), 20, 'span stretches across blackout weeks');
eq(RM.stretchSpan(META, 0, 10), 10, 'span with no blackout is exact');
eq(RM.workInSpan(META, 15 * 5, 20), 10, 'workInSpan inverse of stretchSpan');

// ------------------------------------------------------------- sizes
section('sizes');
// week-based scale: XS 2d · S 1w · M 2w · L 4w · XL 8w
var s0 = mkState([]);
eq(RM.sizeDays(s0, 'S'), 5, 'S = 1 week');
eq(RM.sizeDays(s0, 'M'), 10, 'M = 2 weeks');
eq(RM.sizeDays(s0, 'L'), 20, 'L = 4 weeks');
eq(RM.sizeDays(s0, 'XL'), 40, 'XL = 8 weeks');
eq(RM.sizeForDays(s0, 18), 'L', '18 days ≈ L');
eq(RM.sizeForDays(s0, 4), 'S', '4 days ≈ S');
eq(RM.sizeForDays(s0, 45), 'XL', '45 days ≈ XL');
// legacy size map migrates to the week scale
var sLeg = RM.normalizeState({ meta: { sizeDays: { XS: 2, S: 3, M: 5, L: 10, XL: 20 } }, phases: [{ id: 'p' }], items: [] });
eq(sLeg.meta.sizeDays.L, 20, 'legacy default size map migrated');
// null/empty day values are healed back to the scheme default; 0 stays valid
var sSzHeal = RM.normalizeState({
  meta: { sizeDays: { XS: 2, S: null, M: 0, L: '' } }, phases: [{ id: 'p' }], items: []
});
eq(sSzHeal.meta.sizeDays.S, RM.DEFAULT_SIZE_DAYS.S, 'null S repaired to scheme default');
eq(sSzHeal.meta.sizeDays.M, 0, 'M stays 0');
eq(sSzHeal.meta.sizeDays.L, RM.DEFAULT_SIZE_DAYS.L, 'empty-string L repaired to scheme default');
// the sizeOrder fallback filters storyOnly sizes for the feature scale, same as setSizeScheme
var sSzOrderFib = RM.normalizeState({ meta: { sizeScheme: 'fibonacci' }, phases: [{ id: 'p' }], items: [] });
eq(sSzOrderFib.meta.sizeOrder.join(','), '0.5,1,2,3,5,8,13', 'feature sizeOrder fallback drops the story-only 0');

// ------------------------------------------------------------- normalize
section('normalizeState');
var sN = mkState([
  { feature: 'A' },
  { num: 7, feature: 'B', headcount: 0, phaseId: 'nope' }
]);
eq(sN.items[0].headcount, 1, 'default headcount 1');
ok(sN.items[0].num != null && sN.items[0].num !== 7, 'auto num assigned, no collision');
eq(sN.items[1].headcount, 1, 'headcount floor 1');
var sMs = mkState([{ feature: 'M', milestone: true, size: 'L', priority: 'P1', startDay: 0, durDays: 0 }]);
sMs.meta.priorityScheme = 'levels'; sMs.items[0].priority = 'P1'; sMs = RM.normalizeState(sMs);
eq(sMs.items[0].size, null, 'milestones carry no size');
eq(sMs.items[0].priority, null, 'milestones carry no priority');
eq(sN.items[1].phaseId, 'p1', 'bad phase falls back to first');
ok(Array.isArray(sN.items[0].stories), 'stories default []');

// ------------------------------------------------------------- deps
section('dependencies');
var sD = mkState([
  { num: 1, feature: 'one', phaseId: 'p1' },
  { num: 2, feature: 'two', phaseId: 'p1', deps: [1] },
  { num: 3, feature: 'three', phaseId: 'p1', depsAllAbove: true },
  { num: 4, feature: 'four', phaseId: 'p1', deps: [99] }
]);
eq(RM.resolveDeps(sD, sD.items[1]).deps.map(function (x) { return x.num; }), [1], 'numbered dep resolves');
eq(RM.resolveDeps(sD, sD.items[2]).deps.length, 0, '"All above" is dropped — only explicit deps count');
ok(sD.items[2].depsAllAbove === undefined, 'depsAllAbove stripped by normalize');
eq(RM.resolveDeps(sD, sD.items[3]).unknown, [99], 'unknown dep reported');

var sC = mkState([
  { num: 1, feature: 'a', deps: [2] },
  { num: 2, feature: 'b', deps: [1] },
  { num: 3, feature: 'c', deps: [2] }
]);
var cyc = RM.cycleMembers(sC);
ok(cyc[sC.items[0].id] && cyc[sC.items[1].id], 'cycle detected for 1<->2');
ok(!cyc[sC.items[2].id], 'downstream of a cycle is not itself cyclic');

// ------------------------------------------------------------- validation
section('validation');
var sV = mkState([
  { num: 1, feature: 'base', startDay: 0, durDays: 10, size: 'L', capType: 'Development' },
  { num: 2, feature: 'early bird', deps: [1], startDay: 5, durDays: 5, size: 'M' },
  { num: 3, feature: 'no size', startDay: 0, durDays: 5 },
  { num: 4, feature: 'ghost dep', deps: [42], startDay: 20, durDays: 5, size: 'M' },
  { num: 5, feature: 'big ask', startDay: 0, durDays: 25, size: 'XL', teamType: 'Data', capType: 'Development' }
], { team: [{ name: 'X', type: 'Development', capType: 'Development' }, { name: 'Y', type: 'Data' }] });
var v = RM.validate(sV);
function codes(state, i) { return (v.byItem[state.items[i].id] || []).map(function (x) { return x.code; }); }
ok(codes(sV, 1).indexOf('DEP_ORDER') !== -1, 'DEP_ORDER: starts before dep ends');
ok(codes(sV, 2).indexOf('NO_SIZE') === -1, 'no NO_SIZE nag — sizing is optional');
ok(codes(sV, 3).indexOf('UNKNOWN_DEP') !== -1, 'UNKNOWN_DEP flagged');
ok(v.global.some(function (g) { return g.code === 'OVER_CAP'; }), 'OVER_CAP: weekly WIP over what the team can focus on');
ok(!v.global.some(function (g) { return /Data/.test(g.msg); }), 'capacity messages are role-agnostic now');
ok(v.counts.warn > 0, 'counts aggregated');

// a feature always carries a capacity type, so a clean document needs somebody
// supplying it (an unsupplied type is a warning of its own)
var vClean = RM.validate(mkState([{ num: 1, feature: 'solo', startDay: 0, durDays: 5, size: 'M' }],
  { team: [{ name: 'X', capType: 'Development' }] }));
eq(vClean.counts.error + vClean.counts.warn, 0, 'clean state has no errors/warnings');

// done-dep suppression
var sDone = mkState([
  { num: 1, feature: 'shipped', startDay: 10, durDays: 10, size: 'L', done: true },
  { num: 2, feature: 'after', deps: [1], startDay: 0, durDays: 5, size: 'M' }
]);
var vd = RM.validate(sDone);
ok(!(vd.byItem[sDone.items[1].id] || []).some(function (x) { return x.code === 'DEP_ORDER'; }),
  'done dependency does not trigger DEP_ORDER');

// ------------------------------------------------------------- capacity
section('capacity');
// person mode: a unit in flight costs one person of its type × multiplier;
// supply is the typed roster in people-equivalents, cut by holidays
var sCap = mkState([
  { num: 1, feature: 'big', startDay: 0, durDays: 20, capType: 'Development' },
  { num: 2, feature: 'mid', startDay: 5, durDays: 10, capType: 'Development', capMult: 2 },
  { num: 3, feature: 'design', startDay: 5, durDays: 5, capType: 'Design' }
], { team: [{ name: 'X', capType: 'Development' }, { name: 'Y', capType: 'Development', capacity: 0.5 }] });
var sup = RM.capSupply(sCap);
eq(sup.types, ['Development'], 'supplied types come from the roster');
eq(sup.byType.Development[0], 1.5, 'week 0 supply = 1 + 0.5 heads');
eq(RM.holidayFactor(META, 16), 0, 'blackout week factor 0');
var METAF = JSON.parse(JSON.stringify(META)); METAF.holidays = ['2026-08-05'];
eq(RM.holidayFactor(METAF, 1), 0.8, 'one holiday in a 5-day week = 0.8');
eq(RM.capSupply(mkState([], { meta: METAF, team: [{ name: 'X', capType: 'Development' }] })).byType.Development[1], 0.8, 'supply scales by the holiday factor');
var cap = RM.capacity(sCap);
eq(cap.weeks[0].demand, 1, 'week 0: one Development unit');
eq(cap.weeks[1].demand, 4, 'week 1: 1 + 2 (multiplier) Development + 1 Design — every type counts');
eq(cap.weeks[0].supply, 1.5, 'row supply aggregates all supplied types');
ok(cap.weeks[1].over, 'week 1 over: Development asks 3 of its 1.5 (4 asked in all)');
ok(!cap.weeks[0].over, 'week 0 fits');
ok(cap.weeks[1].byType.Design.demand === 1 && cap.weeks[1].byType.Design.supply === 0, 'a type nobody supplies is listed with zero supply but never marks over');
ok(cap.weeks[1].items.indexOf(sCap.items[2].id) !== -1, 'items in flight listed');
// every capacity type counts: one row per type, the aggregate sums them all
var capTwo = RM.capacity(sCap);
eq(capTwo.types, RM.capTypesOf(sCap), 'one row per capacity type, in capacity-type order');
eq(capTwo.rows.Development[1].demand, 3, 'the Development row carries only Development demand');
eq(capTwo.rows.Development[1].supply, 1.5, 'and only Development supply');
eq(capTwo.rows.Design[1].demand, 1, 'the Design row carries only Design demand');
eq(capTwo.rows.Design[1].supply, 0, 'nobody supplies Design');
ok(capTwo.rows.Design[1].over, 'a type nobody supplies is over as soon as anything asks for it');
ok(capTwo.rows.Development[1].over && !capTwo.rows.Development[0].over, 'a supplied row is over only in the week that over-asks');
ok(capTwo.rows.Design[1].items.indexOf(sCap.items[2].id) !== -1, 'each row lists its own items');
eq(capTwo.weeks[1].demand, 4, 'the aggregate week sums every type');
ok(capTwo.rows.Development[16].blackout, 'rows carry the blackout weeks too');
ok(capTwo.weeks[1].overAny && !capTwo.weeks[0].overAny, 'overAny: some type is over that week');
eq(RM.trackedCapTypes, undefined, 'there is no tracked-type filter any more');
// the sum can hide one type overflowing: over / overAny read per type
{
  var sHide = mkState([
    { num: 1, feature: 'dev a', startDay: 0, durDays: 5, capType: 'Development' },
    { num: 2, feature: 'dev b', startDay: 0, durDays: 5, capType: 'Development' }
  ], { team: [{ name: 'D', capType: 'Development' }, { name: 'S1', capType: 'Design' }, { name: 'S2', capType: 'Design' }] });
  var cHide = RM.capacity(sHide);
  ok(cHide.weeks[0].demand === 2 && cHide.weeks[0].supply === 3, 'the week sums 2 asked of 3');
  ok(cHide.weeks[0].over && cHide.weeks[0].overAny, 'yet it is over: Development asks 2 of its 1');
  // work of a type nobody supplies: overAny (the header shows it), not over
  // (validation says CAP_TYPE_UNSUPPLIED instead, the scheduler ignores it)
  var sDry = mkState([{ num: 1, feature: 'qa', startDay: 0, durDays: 5, capType: 'QA' }],
    { team: [{ name: 'D', capType: 'Development' }] });
  var cDry = RM.capacity(sDry);
  ok(cDry.weeks[0].overAny && !cDry.weeks[0].over, 'unsupplied work: overAny but not over');
}
// untyped people supply nothing, whatever their seat or points
{
  var sUt = mkState([], { team: [{ name: 'U', capType: '', capacity: 2, points: 50 }, { name: 'D', capType: 'Development' }] });
  eq(RM.capSupply(sUt).types, ['Development'], 'an untyped person supplies no type');
  eq(RM.capSupply(sUt).byType.Development[0], 1, 'and adds nothing to anyone else\'s supply');
  sUt.meta.capMode = 'points';
  eq(RM.capacity(sUt).weeks[0].supply, 10, 'points mode: only the typed person\'s 10 per sprint');
  ok(sUt.team[0].capacity === 2 && sUt.team[0].points === 50, 'normalize keeps an untyped person\'s seat and points for later');
}
// story level: stories carry the demand; the feature bar is ignored
var sCapS = mkState([
  { num: 1, feature: 'f', startDay: 0, durDays: 20, capType: 'Development', stories: [
    { title: 'a', startDay: 0, durDays: 5, capType: 'Design' },
    { title: 'b', capType: 'Development' }
  ] }
], { team: [{ name: 'D', capType: 'Design' }] });
sCapS.meta.planLevel = 'story';
var capS = RM.capacity(sCapS);
eq(capS.weeks[0].demand, 1, 'story level: only the scheduled story counts');
eq(capS.weeks[1].demand, 0, 'unscheduled stories and the feature bar add nothing');
eq(RM.capUnits(sCapS).length, 2, 'one unit per story');
eq(RM.capUnits(mkState([{ num: 1, feature: 'lonely', startDay: 0, durDays: 5 }], { meta: Object.assign({}, META, { planLevel: 'story' }) })).length, 1, 'a story-less feature is one unit of its own at story level');
// points mode: points spread over the unit's working weeks vs points per sprint
var sPts = mkState([
  { num: 1, feature: 'eight', startDay: 0, durDays: 10, size: 8, capType: 'Development' }
], { team: [{ name: 'X', capType: 'Development' }, { name: 'Y', capType: 'Development', points: 6 }] });
sPts.meta.capMode = 'points';
var capP = RM.capacity(sPts);
eq(capP.period, 'sprint', 'points mode summarizes per sprint');
eq(capP.periods.slice(0, 2), [{ w0: 0, w1: 2, num: 1 }, { w0: 2, w1: 4, num: 2 }], 'the summary carries its periods');
eq(capP.weeks.length, capP.periods.length, 'one aggregate cell per period');
eq(capP.weeks[0].demand, 8, '8 points over the 2-week sprint = 8 in that sprint');
eq(capP.weeks[0].supply, 16, '(10 + 6) points per 2-week sprint');
eq(capP.rows.Development[0].demand, 8, 'the type row is per sprint too');
eq(capP.rows.Development[1].demand, 0, 'nothing asked in sprint 2');
sPts.meta.weeksPerSprint = 0;
var capP0 = RM.capacity(sPts);
eq(capP0.weeks[0].supply, 16, 'sprints off: points are per two weeks, summed over two-week blocks');
eq(capP0.periods[0], { w0: 0, w1: 2, num: null }, 'sprints off: the blocks carry no sprint number');
// per person the summary stays weekly
var capW = RM.capacity(sCap);
ok(capW.period === 'week' && capW.periods.length === sCap.meta.numWeeks && capW.weeks.length === sCap.meta.numWeeks,
  'per-person mode: one period (and cell) per week');
// a sprint total over supply reads over even when one of its weeks is under;
// a single week over its share does not, while the sprint holds
{
  var sOv = mkState([
    { num: 1, feature: 'big', startDay: 0, durDays: 5, size: 8, capType: 'Development' },
    { num: 2, feature: 'small', startDay: 5, durDays: 5, size: 4, capType: 'Development' },
    { num: 3, feature: 'lone', startDay: 10, durDays: 5, size: 8, capType: 'Development' }
  ], { team: [{ name: 'X', capType: 'Development', points: 10 }] });
  sOv.meta.capMode = 'points';
  var cOv = RM.capacity(sOv);
  eq(cOv.rows.Development[0].demand, 12, 'sprint 1 asks 8 + 4');
  eq(cOv.rows.Development[0].supply, 10, 'sprint 1 supplies 10');
  ok(cOv.rows.Development[0].over && cOv.weeks[0].over, 'week 2 asks 4 of its 5, but the sprint asks 12 of 10: over');
  eq(cOv.rows.Development[0].items.sort(), [sOv.items[0].id, sOv.items[1].id].sort(), 'the sprint cell lists both items');
  ok(!cOv.rows.Development[1].over, 'sprint 2: 8 points in its first week is within the sprint\'s 10');
  var vOv = RM.validate(sOv).global.filter(function (v) { return v.code === 'OVER_CAP'; });
  eq(vOv.length, 1, 'validation reports once per over sprint');
  ok(/Sprint 1\b/.test(vOv[0].msg) && vOv[0].week === 0, 'the message names the sprint and points at its first week (' + vOv[0].msg + ')');
  // a fully blacked-out sprint reads blackout
  var sBo = mkState([], { team: [{ name: 'X', capType: 'Development', points: 10 }] });
  sBo.meta.capMode = 'points';
  var cBo = RM.capacity(sBo);
  var boIdx = -1;
  cBo.periods.forEach(function (p, i) { if (p.w0 === 16 && p.w1 === 18) boIdx = i; });
  ok(boIdx !== -1 && cBo.weeks[boIdx].blackout && cBo.rows.Development[boIdx].blackout, 'a sprint of two holiday weeks is blackout');
}
// feature type rolls up from stories when they agree
var sRoll = mkState([{ num: 1, feature: 'f', capType: 'Development', stories: [{ title: 'a', capType: 'Design' }, { title: 'b', capType: 'Design' }] }]);
eq(RM.itemCapType(sRoll, sRoll.items[0]), 'Design', 'all stories Design → feature plans as Design');
sRoll.items[0].stories[1].capType = 'QA';
eq(RM.itemCapType(sRoll, sRoll.items[0]), 'Development', 'mixed stories → the feature keeps its own type');
// validation names the type
var vCap = RM.validate(sCap).global.filter(function (v) { return v.code === 'OVER_CAP'; });
ok(vCap.length && /Development/.test(vCap[0].msg), 'OVER_CAP names the type');
ok(RM.validate(sCap).global.some(function (v) { return v.code === 'CAP_TYPE_UNSUPPLIED' && /Design/.test(v.msg); }), 'CAP_TYPE_UNSUPPLIED for a type nobody supplies');

// ------------------------------------------------------------- capacity fields
section('capacity periods');
{
  var cpTrip = function (ps) { return ps.map(function (p) { return [p.w0, p.w1, p.num]; }); };
  var mW = JSON.parse(JSON.stringify(META));
  eq(cpTrip(RM.capPeriods(mW, 4)), [[0, 1, null], [1, 2, null], [2, 3, null], [3, 4, null]], 'per person: one period per week');
  var mS = JSON.parse(JSON.stringify(META)); mS.capMode = 'points';
  eq(cpTrip(RM.capPeriods(mS, 5)), [[0, 2, 1], [2, 4, 2], [4, 5, 3]], 'points: one period per sprint, the last one partial');
  mS.weeksPerSprint = 3; mS.sprintAnchor = '2026-08-03'; mS.sprintAnchorNum = 4; // anchored at week 1
  eq(cpTrip(RM.capPeriods(mS, 8)), [[0, 1, 3], [1, 4, 4], [4, 7, 5], [7, 8, 6]], 'points: sprints follow the numbering anchor (a partial first sprint)');
  var mO = JSON.parse(JSON.stringify(META)); mO.capMode = 'points'; mO.weeksPerSprint = 0;
  eq(cpTrip(RM.capPeriods(mO, 5)), [[0, 2, null], [2, 4, null], [4, 5, null]], 'points with sprints off: two-week blocks from the timeline start');
  eq(RM.capPeriods(mS).slice(-1)[0].w1, mS.numWeeks, 'horizon defaults to the timeline');
}

section('capacity fields');
var sF = RM.normalizeState({
  meta: { timelineStart: '2026-07-27', numWeeks: 8, capLimit: 3, capBasis: 'stories', capUnit: 'points', capacityEnabled: true, capMode: 'points', defaultPoints: 8, capRowTypes: ['Design'] }, // an old doc's row selection
  phases: [{ id: 'p1', auto: true }, { id: 'p2', bucket: true, auto: true }],
  items: [{ num: 1, feature: 'f', capType: 'Design', capMult: 2, stories: [{ title: 's', capMult: 0 }] }],
  team: [{ name: 'A', points: 12 }, { name: 'B' }]
});
ok(sF.meta.capLimit === undefined && sF.meta.capBasis === undefined && sF.meta.capUnit === undefined, 'weekly limit and row basis/unit fields are gone');
eq(sF.meta.capMode, 'points', 'capMode kept');
eq(sF.meta.defaultPoints, 8, 'defaultPoints kept');
eq(sF.meta.capRowTypes, undefined, 'an old document\'s capRowTypes is dropped: every type counts');
eq(RM.normalizeState({ meta: {}, phases: [{ id: 'p' }], items: [] }).meta.capMode, 'person', 'capMode defaults to person');
eq(RM.normalizeState({ meta: {}, phases: [{ id: 'p' }], items: [] }).meta.defaultPoints, 10, 'defaultPoints defaults to 10');
eq(RM.normalizeState({ meta: {}, phases: [{ id: 'p' }], items: [] }).meta.capRowTypes, undefined, 'no capRowTypes on a new document');
eq(sF.items[0].capType, 'Design', 'feature capType kept');
eq(sF.items[0].capMult, 2, 'feature capMult kept');
eq(sF.items[0].stories[0].capMult, 1, 'story capMult below or at 0 falls back to 1');
eq(sF.team[0].points, 12, 'member points kept');
eq(sF.team[1].points, null, 'member points default null');
eq(RM.memberPoints(sF, sF.team[1]), 8, 'memberPoints falls back to the document default');
// Auto timeline is a one-shot action now: an old document's per-phase flag is dropped
ok(!('auto' in sF.phases[0]) && !('auto' in sF.phases[1]), 'an old phase.auto flag is dropped on normalize');
eq(RM.normalizeState({ meta: { defaultPoints: '' }, phases: [{ id: 'p' }], items: [] }).meta.defaultPoints, 10, 'a blank defaultPoints falls back to 10');
// a capacity type used by a feature follows a rename / removal
var sCT = RM.normalizeState({ meta: {}, phases: [{ id: 'p' }], items: [{ num: 1, feature: 'f', capType: 'Design' }] });
ok(RM.renameCapType(sCT, 'Design', 'UX'), 'renameCapType reports the rename');
eq(sCT.items[0].capType, 'UX', 'a feature capType follows the rename');
eq(RM.normalizeState(sCT).capTypes.indexOf('Design'), -1, 'renormalizing does not resurrect the old type');
ok(RM.removeCapType(sCT, 'UX'), 'removeCapType reports the removal');
eq(sCT.items[0].capType, '', 'a feature capType is cleared on removal');
eq(RM.normalizeState(sCT).capTypes.indexOf('UX'), -1, 'renormalizing does not resurrect the removed type');
// documents written before capacity types: a blank feature type takes the
// document's first type so it is never invisible to the capacity row
var sDefCT = RM.normalizeState({ meta: {}, capTypes: ['Development', 'Design'], phases: [{ id: 'p' }], items: [
  { num: 1, feature: 'old', phaseId: 'p', stories: [{ num: 101, title: 'st' }] },
  { num: 2, feature: 'gate', phaseId: 'p', milestone: true }
] });
eq(sDefCT.items[0].capType, 'Development', 'a blank feature capType defaults to the first capacity type');
eq(sDefCT.items[0].stories[0].capType, '', 'a story may still be general');
eq(sDefCT.items[1].capType, '', 'a milestone stays untyped');

// the span helpers take an optional prebuilt holiday set — hot loops (the
// auto timeline) pass one instead of rebuilding it per call
var hsSet = RM.holidayDaySet(META);
eq(RM.stretchSpan(META, 70, 12, hsSet), RM.stretchSpan(META, 70, 12), 'stretchSpan with a passed holiday set matches');
eq(RM.workInSpan(META, 70, 30, hsSet), RM.workInSpan(META, 70, 30), 'workInSpan with a passed holiday set matches');
eq(RM.workingWeeksInSpan(META, 70, 30, hsSet), RM.workingWeeksInSpan(META, 70, 30), 'workingWeeksInSpan with a passed holiday set matches');

// ------------------------------------------------------------- auto timeline
section('autoTimeline');
function autoMeta(extra) {
  var m = JSON.parse(JSON.stringify(META));
  m.holidays = []; m.capacityEnabled = true;
  if (extra) Object.keys(extra).forEach(function (k) { m[k] = extra[k]; });
  return m;
}
function autoState(items, team, extraMeta, phases) {
  return mkState(items, { meta: autoMeta(extraMeta), team: team || [],
    phases: phases || [{ id: 'p1', name: 'Alpha', bucket: false }, { id: 'p2', name: 'Later', bucket: false }, { id: 'p3', name: 'Next', bucket: true }] });
}
function byNum(st) { var o = {}; st.items.forEach(function (it) { o[it.num] = it; }); return o; }
// feature level, person mode, one Development head: three features serialize, deps hold
var sT = autoState([
  { num: 1, feature: 'a', phaseId: 'p1', durDays: 5, capType: 'Development' },
  { num: 2, feature: 'b', phaseId: 'p1', durDays: 5, capType: 'Development', deps: [1] },
  { num: 3, feature: 'c', phaseId: 'p1', durDays: 5, capType: 'Development' },
  { num: 8, feature: 'other phase', phaseId: 'p2', startDay: 40, durDays: 5, capType: 'Development' },
  { num: 9, feature: 'parked', phaseId: 'p3', durDays: 5 }
], [{ name: 'Solo', capType: 'Development' }]);
var rT = RM.autoTimeline(sT, { phaseIds: ['p1'], today: 0 });
var T = byNum(rT.state);
eq(T[1].startDay, 0, 'first unit starts today');
eq(T[3].startDay, 5, 'third slides to week 1 (cap 1)');
eq(T[2].startDay, 10, 'dependent lands after its dep and after the busy week');
eq(T[8].startDay, 40, 'non-auto phase untouched');
ok(T[9].startDay == null, 'bucket untouched');
ok(!RM.capacity(rT.state).weeks.some(function (c) { return c.over; }), 'never overallocates');
eq(RM.autoTimeline(sT, { today: 0, phaseIds: [] }).changed, 0, 'no target phases → nothing changes');
// no phaseIds: every non-bucket phase is laid out (there is no per-phase flag any more)
var sTAll = autoState([
  { num: 1, feature: 'a', phaseId: 'p1', durDays: 5, capType: 'Development' },
  { num: 2, feature: 'b', phaseId: 'p2', durDays: 5, capType: 'Development' },
  { num: 3, feature: 'parked', phaseId: 'p3', durDays: 5 }
], [{ name: 'Solo', capType: 'Development' }]);
var TAll = byNum(RM.autoTimeline(sTAll, { today: 0 }).state);
ok(TAll[1].startDay === 0 && TAll[2].startDay === 5, 'without phaseIds every real phase is laid out');
ok(TAll[3].startDay == null, 'but a bucket never is');
var sOff = autoState([{ num: 1, feature: 'a', phaseId: 'p1', durDays: 5 }], [], { capacityEnabled: false });
eq(RM.autoTimeline(sOff, { phaseIds: ['p1'], today: 0 }).changed, 0, 'capacity off → nothing changes');
// every supplied type is constrained; a type nobody supplies is dependency-only
function trackState(team) {
  return autoState([
    { num: 1, feature: 'dsn a', phaseId: 'p1', durDays: 5, capType: 'Design' },
    { num: 2, feature: 'dsn b', phaseId: 'p1', durDays: 5, capType: 'Design' }
  ], team, { capRowTypes: ['Development'] }); // a stale old-doc selection is ignored
}
eq(byNum(RM.autoTimeline(trackState([{ name: 'D', capType: 'Development' }]), { phaseIds: ['p1'], today: 0 }).state)[2].startDay, 0,
  'a type nobody supplies is dependency-only: both Design units start in the same week');
eq(byNum(RM.autoTimeline(trackState([{ name: 'D', capType: 'Development' }, { name: 'S', capType: 'Design' }]), { phaseIds: ['p1'], today: 0 }).state)[2].startDay, 5,
  'once someone supplies Design the second Design unit waits for the week to free up');
// cap 2: two run together, third waits
var s2 = autoState([
  { num: 1, feature: 'a', phaseId: 'p1', durDays: 5, capType: 'Development' },
  { num: 2, feature: 'b', phaseId: 'p1', durDays: 5, capType: 'Development' },
  { num: 3, feature: 'c', phaseId: 'p1', durDays: 5, capType: 'Development' }
], [{ name: 'X', capType: 'Development' }, { name: 'Y', capType: 'Development' }]);
var R2 = byNum(RM.autoTimeline(s2, { phaseIds: ['p1'], today: 0 }).state);
ok(R2[1].startDay === 0 && R2[2].startDay === 0 && R2[3].startDay === 5, 'cap 2 lets two run at once');
// locked pre-books; unlocked flows around; the phase floor holds the rest
var sLk = autoState([
  { num: 1, feature: 'rock', phaseId: 'p1', startDay: 5, durDays: 5, locked: true, capType: 'Development' },
  { num: 2, feature: 'water', phaseId: 'p1', durDays: 10, capType: 'Development' },
  { num: 3, feature: 'free', phaseId: 'p1', durDays: 5, capType: '' }
], [{ name: 'Solo', capType: 'Development' }]);
var Lk = byNum(RM.autoTimeline(sLk, { phaseIds: ['p1'], today: 0 }).state);
eq(Lk[1].startDay, 5, 'locked stays');
eq(Lk[2].startDay, 10, 'a 2-week unit cannot straddle the locked week, so it waits');
eq(Lk[3].startDay, 20, 'the third unit waits for a free week (the phase floor only rules out weeks before day 5)');
// a locked item is a fixed point: its dependent is placed after it, the locked dates never move
var sLkD = autoState([
  { num: 1, feature: 'anchor', phaseId: 'p1', startDay: 20, durDays: 5, locked: true, capType: 'Development' },
  { num: 2, feature: 'follows', phaseId: 'p1', startDay: 0, durDays: 5, capType: 'Development', deps: [1] }
], [{ name: 'X', capType: 'Development' }, { name: 'Y', capType: 'Development' }]);
var LkD = byNum(RM.autoTimeline(sLkD, { phaseIds: ['p1'], today: 0 }).state);
eq([LkD[1].startDay, LkD[1].durDays], [20, 5], 'the locked item keeps its start and length');
eq(LkD[2].startDay, 25, 'and its dependent lands right after it');
// milestones: dependency-free stays; with deps lands at the dep end
var sMs = autoState([
  { num: 1, feature: 'work', phaseId: 'p1', durDays: 5, capType: 'Development' },
  { num: 2, feature: 'gate', phaseId: 'p1', milestone: true, startDay: 30, durDays: 0, deps: [1] },
  { num: 3, feature: 'fixed date', phaseId: 'p1', milestone: true, startDay: 30, durDays: 0 },
  { num: 4, feature: 'after gate', phaseId: 'p1', durDays: 5, capType: 'Development', deps: [2] }
], [{ name: 'Solo', capType: 'Development' }]);
var Ms = byNum(RM.autoTimeline(sMs, { phaseIds: ['p1'], today: 0 }).state);
eq(Ms[2].startDay, 5, 'milestone with deps moves to the dep end');
eq(Ms[3].startDay, 30, 'dependency-free milestone stays');
eq(Ms[4].startDay, 5, 'dependents may start on the milestone day');
// floor: started work keeps its start; future work may move earlier
var sFl = autoState([
  { num: 1, feature: 'in progress', phaseId: 'p1', startDay: 2, durDays: 5, capType: 'Development' },
  { num: 2, feature: 'far future', phaseId: 'p1', startDay: 60, durDays: 5, capType: 'Development' }
], [{ name: 'X', capType: 'Development' }, { name: 'Y', capType: 'Development' }]);
var Fl = byNum(RM.autoTimeline(sFl, { phaseIds: ['p1'], today: 4 }).state);
eq(Fl[1].startDay, 2, 'started work keeps its start');
eq(Fl[2].startDay, 4, 'future work pulls in to today');
// holidays stretch, working days preserved
var sHo = autoState([
  // unscheduled, 10 working days of effort; the floor (today) is day 75
  { num: 1, feature: 'spanner', phaseId: 'p1', durDays: 10, capType: 'Development' }
], [{ name: 'Solo', capType: 'Development' }], { holidays: META.holidays });
var rHo = RM.autoTimeline(sHo, { phaseIds: ['p1'], today: 75 });
var Ho = byNum(rHo.state);
eq(Ho[1].startDay, 75, 'starts at the floor');
eq(Ho[1].durDays, 20, '10 working days stretch over two blackout weeks');
eq(RM.workInSpan(rHo.state.meta, Ho[1].startDay, Ho[1].durDays), 10, 'net work preserved');
// infeasible: demand above peak supply stays put with a note
var sInf = autoState([
  { num: 1, feature: 'crowd', phaseId: 'p1', startDay: 10, durDays: 5, capType: 'Development', capMult: 3 }
], [{ name: 'X', capType: 'Development' }]);
var rInf = RM.autoTimeline(sInf, { phaseIds: ['p1'], today: 0 });
eq(byNum(rInf.state)[1].startDay, 10, 'infeasible unit keeps its start');
ok(rInf.notes.length === 1 && /never/.test(rInf.notes[0]), 'and explains itself');
// left where it is means still booked there: later work goes around it
var sInfB = autoState([
  { num: 1, feature: 'crowd', phaseId: 'p1', startDay: 10, durDays: 5, capType: 'Development', capMult: 3 },
  { num: 2, feature: 'normal', phaseId: 'p1', durDays: 5, capType: 'Development' }
], [{ name: 'X', capType: 'Development' }]);
var InfB = byNum(RM.autoTimeline(sInfB, { phaseIds: ['p1'], today: 10 }).state);
eq(InfB[1].startDay, 10, 'the infeasible unit keeps week 2');
eq(InfB[2].startDay, 15, 'and feasible work is not piled onto the week it still occupies');
// a capacity type named like an Object.prototype key behaves like any other
var sProto = autoState([
  { num: 1, feature: 'crowd', phaseId: 'p1', startDay: 10, durDays: 5, capType: 'constructor', capMult: 3 },
  { num: 2, feature: 'normal', phaseId: 'p1', durDays: 5, capType: 'constructor' }
], [{ name: 'X', capType: 'constructor' }]);
var rProto = RM.autoTimeline(sProto, { phaseIds: ['p1'], today: 10 });
var Proto = byNum(rProto.state);
eq(Proto[1].startDay, 10, 'an infeasible unit of a type named "constructor" is left where it is');
ok(rProto.notes.length === 1 && /never/.test(rProto.notes[0]), 'and the never note still fires for it');
eq(Proto[2].startDay, 15, 'while a feasible unit of that type is placed around it');
// ... and the capacity summary carries that type without writing onto Object
var sProtoRow = autoState([
  { num: 1, feature: 'crowd', phaseId: 'p1', startDay: 10, durDays: 5, capType: 'constructor', capMult: 3 }
], [{ name: 'X', capType: 'Development' }]);
var protoCell = RM.capacity(sProtoRow).weeks[2];
ok(Object.prototype.hasOwnProperty.call(protoCell.byType, 'constructor') && protoCell.byType.constructor.demand === 3,
  'an unsupplied type named "constructor" carries its own demand in byType');
eq(Object.demand, undefined, 'and nothing was written onto Object itself');
// story level: stories move, feature bar becomes their hull
var sSt = autoState([
  { num: 1, feature: 'f', phaseId: 'p1', startDay: 0, durDays: 40, capType: 'Development', stories: [
    { num: 101, title: 'a', durDays: 5, capType: 'Development' },
    { num: 102, title: 'b', durDays: 5, capType: 'Development', deps: [101] },
    { num: 103, title: 'c', durDays: 5, capType: 'Development' }
  ] },
  { num: 2, feature: 'g', phaseId: 'p1', capType: 'Development', deps: [1], stories: [{ num: 201, title: 'd', durDays: 5, capType: 'Development' }] }
], [{ name: 'Solo', capType: 'Development' }], { planLevel: 'story' });
var rSt = RM.autoTimeline(sSt, { phaseIds: ['p1'], today: 0 });
var St = byNum(rSt.state);
var sts = {}; St[1].stories.forEach(function (s) { sts[s.num] = s; });
eq(sts[101].startDay, 0, 'story a first');
eq(sts[103].startDay, 5, 'story c waits for capacity');
eq(sts[102].startDay, 10, 'story b after a and after c took week 1');
eq(St[1].startDay, 0, 'feature hull start');
eq(St[1].durDays, 15, 'feature hull spans its stories');
eq(St[2].stories[0].startDay, 15, 'feature dep expands to every story of the dependency');
ok(rSt.state.items.every(function (it) { return it.stories.every(function (s) { return s.startDay != null; }); }), 'every story placed');
// an untyped unit carries no demand: a week that is full for everyone else
// never holds it back (only its dependencies and its phase floor do)
var sUn = autoState([
  { num: 1, feature: 'f', phaseId: 'p1', startDay: 0, durDays: 5, stories: [
    { num: 101, title: 'typed', startDay: 0, durDays: 5, capType: 'Development' },
    { num: 102, title: 'untyped', durDays: 5, capType: '' }
  ] }
], [{ name: 'Solo', capType: 'Development' }], { planLevel: 'story' });
var Un = RM.itemByNum(RM.autoTimeline(sUn, { phaseIds: ['p1'], today: 0 }).state, 1);
eq(Un.stories[0].startDay, 0, 'the typed story fills the only Development week there is');
eq(Un.stories[1].startDay, 0, 'and an untyped story sits right beside it — placed by dependencies only');
// points mode
var sPm = autoState([
  { num: 1, feature: 'ten', phaseId: 'p1', durDays: 5, size: 10, capType: 'Development' },
  { num: 2, feature: 'six', phaseId: 'p1', durDays: 5, size: 6, capType: 'Development' }
], [{ name: 'X', capType: 'Development', points: 20 }], { capMode: 'points', sizeScheme: 'points' });
var Pm = byNum(RM.autoTimeline(sPm, { phaseIds: ['p1'], today: 0 }).state);
ok(Pm[1].startDay === 0 && Pm[2].startDay === 0, '10 + 6 points fit one 20-point sprint, even in the same week');
var sPm3 = autoState([
  { num: 1, feature: 'ten', phaseId: 'p1', durDays: 5, size: 10, capType: 'Development' },
  { num: 2, feature: 'six', phaseId: 'p1', durDays: 5, size: 6, capType: 'Development' },
  { num: 3, feature: 'six more', phaseId: 'p1', durDays: 5, size: 6, capType: 'Development' }
], [{ name: 'X', capType: 'Development', points: 20 }], { capMode: 'points', sizeScheme: 'points' });
eq(byNum(RM.autoTimeline(sPm3, { phaseIds: ['p1'], today: 0 }).state)[3].startDay, 7,
  '10 + 6 + 6 exceed the 20-point sprint → the third starts at day 7: points follow working days, so only 3 of its 5 days (3.6 of its 6 points) land in sprint 1');
eq(byNum(RM.autoTimeline(sPm3, { phaseIds: ['p1'], today: 0, snap: { feature: 'sprint' } }).state)[3].startDay, 10,
  'snapped to sprints, the third waits for the next sprint');
// the ledger checks the sprint, not the week: 8 booked in week 1 leaves 2 of
// the sprint's 10, so a 4-point story does not fit in week 2 either (week 2
// alone would give it 5)
var sPl8 = autoState([
  { num: 1, feature: 'booked', phaseId: 'p2', startDay: 0, durDays: 5, size: 8, capType: 'Development' },
  { num: 2, feature: 'four', phaseId: 'p1', durDays: 5, size: 4, capType: 'Development' }
], [{ name: 'X', capType: 'Development', points: 10 }], { capMode: 'points', sizeScheme: 'points' });
eq(byNum(RM.autoTimeline(sPl8, { phaseIds: ['p1'], today: 5 }).state)[2].startDay, 8,
  'week 2 alone has room, but the sprint does not → not at day 5; at day 8 only 2 of its 5 days (1.6 points) land in sprint 1');
eq(byNum(RM.placeUnit(sPl8, sPl8.items[1].id, null, { today: 5 }).state)[2].startDay, 8, 'Place at earliest slot checks the sprint too');
eq(byNum(RM.placeUnit(sPl8, sPl8.items[1].id, null, { today: 5, snap: { feature: 'sprint' } }).state)[2].startDay, 10,
  'snapped to sprints, it takes the next sprint');
// the never-fits guard measures a sprint: 8 points in one week asks more than
// the week's 5 but the sprint gives 10
var sBig = autoState([
  { num: 1, feature: 'eight', phaseId: 'p1', durDays: 5, size: 8, capType: 'Development' },
  { num: 2, feature: 'two', phaseId: 'p1', durDays: 5, size: 2, capType: 'Development' }
], [{ name: 'X', capType: 'Development', points: 10 }], { capMode: 'points', sizeScheme: 'points' });
var rBig = RM.autoTimeline(sBig, { phaseIds: ['p1'], today: 0 });
var Big = byNum(rBig.state);
ok(Big[1].startDay === 0 && !rBig.notes.some(function (n) { return /never fits/.test(n); }), '8 points in one week fits a 10-point sprint');
eq(Big[2].startDay, 0, 'and 2 more fill the sprint exactly');
// never fits, story points: a 30-point, 15-day feature against 10 points a
// sprint puts at least 8 of its days (16 points) in one sprint wherever it
// starts — left where it is, with the note, and the timeline not stretched
{
  var nfTeam = [{ name: 'X', capType: 'Development', points: 10 }];
  var sNf = autoState([{ num: 1, feature: 'huge', phaseId: 'p1', durDays: 15, size: 30, capType: 'Development' }],
    nfTeam, { capMode: 'points', sizeScheme: 'points' });
  var rNf = RM.autoTimeline(sNf, { phaseIds: ['p1'], today: 0 });
  ok(rNf.state.items[0].startDay == null, 'the never-fitting feature is not moved');
  ok(rNf.notes.some(function (n) { return /never fits/.test(n) && /in a sprint/.test(n); }), 'the never-fits note says "in a sprint" (' + rNf.notes.join(' | ') + ')');
  eq(rNf.state.meta.numWeeks, sNf.meta.numWeeks, 'and the timeline is not stretched to a far horizon');
  var pNf = RM.placeUnit(sNf, sNf.items[0].id, null, { today: 0 });
  ok(pNf.state.items[0].startDay == null && /never fits/.test(pNf.note || '') && /in a sprint/.test(pNf.note || ''),
    'Place at earliest slot leaves it too, with the note (' + pNf.note + ')');
  eq(pNf.state.meta.numWeeks, sNf.meta.numWeeks, 'Place does not stretch the timeline either');
  // 18 points over 15 days fits once it straddles two sprints 8 / 7 days (9.6 + 8.4 points)
  var sNf18 = autoState([{ num: 1, feature: 'long', phaseId: 'p1', durDays: 15, size: 18, capType: 'Development' }],
    nfTeam, { capMode: 'points', sizeScheme: 'points' });
  var rNf18 = RM.autoTimeline(sNf18, { phaseIds: ['p1'], today: 0 });
  eq(rNf18.state.items[0].startDay, 2, 'an 18-point 3-week feature fits from day 2 (8 days in sprint 1, 7 in sprint 2)');
  ok(!rNf18.notes.some(function (n) { return /never fits/.test(n); }), 'with no never-fits note');
  // the search itself gives up at the horizon: supply only in week 5 (5
  // points in sprint 3) — a 12-day unit always touches an empty sprint
  var onlyWk4 = {};
  for (var hw = 0; hw < 200; hw++) onlyWk4[RM.fmtISO(RM.weekStartDate(sNf.meta, hw))] = hw === 4 ? 40 : 0;
  var hzTeam = [{ name: 'X', capType: 'Development', points: 10, weekHours: onlyWk4 }];
  var sHzn = autoState([{ num: 1, feature: 'twelve', phaseId: 'p1', durDays: 12, size: 4, capType: 'Development' }],
    hzTeam, { capMode: 'points', sizeScheme: 'points' });
  var rHzn = RM.autoTimeline(sHzn, { phaseIds: ['p1'], today: 0 });
  ok(rHzn.state.items[0].startDay == null && rHzn.notes.some(function (n) { return /never fits/.test(n); }),
    'a search that reaches the horizon without a fit is never-fits, not a slot at the horizon');
  eq(rHzn.state.meta.numWeeks, sHzn.meta.numWeeks, 'and the timeline stays put');
  ok(RM.placeUnit(sHzn, sHzn.items[0].id, null, { today: 0 }).state.items[0].startDay == null, 'Place at earliest slot gives up the same way');
  var sHz5 = autoState([{ num: 1, feature: 'five', phaseId: 'p1', durDays: 5, size: 4, capType: 'Development' }],
    hzTeam, { capMode: 'points', sizeScheme: 'points' });
  eq(RM.autoTimeline(sHz5, { phaseIds: ['p1'], today: 0 }).state.items[0].startDay, 20, 'while a 5-day unit fits the one supplied week');
}
// per person the note still says "in a week"
{
  var sNfP = autoState([{ num: 1, feature: 'crowd', phaseId: 'p1', durDays: 5, capType: 'Development', capMult: 3 }],
    [{ name: 'X', capType: 'Development' }]);
  ok(RM.autoTimeline(sNfP, { phaseIds: ['p1'], today: 0 }).notes.some(function (n) { return /in a week/.test(n); }), 'per person the note says "in a week"');
}
// points follow working days, not weeks touched: a 5-day, 10-point story
// starting on the Wednesday of week 2 puts 3 days (6 points) in sprint 1 and
// 2 days (4 points) in sprint 2 — not 5 / 5 by weeks touched
{
  var sDw = mkState([{ num: 1, feature: 'wed', startDay: 7, durDays: 5, size: 10, capType: 'Development' }],
    { team: [{ name: 'X', capType: 'Development', points: 10 }] });
  sDw.meta.capMode = 'points'; sDw.meta.holidays = [];
  var dw = RM.unitDemandByWeek(sDw, RM.capUnits(sDw)[0], 7, 5);
  eq([dw.w0, dw.byWeek], [1, [6, 4]], 'week 2 gets 3 of 5 days (6 points), week 3 gets 2 (4 points)');
  var cDw = RM.capacity(sDw);
  eq([cDw.rows.Development[0].demand, cDw.rows.Development[1].demand], [6, 4], 'the sprint cells follow the days');
  sDw.meta.capMode = 'person';
  eq(RM.unitDemandByWeek(sDw, RM.capUnits(sDw)[0], 7, 5).byWeek, [1, 1], 'per person: one head in each week touched');
  // a holiday day carries no points
  sDw.meta.capMode = 'points'; sDw.meta.holidays = ['2026-08-07']; // Friday of week 2 (day 9)
  eq(RM.unitDemandByWeek(sDw, RM.capUnits(sDw)[0], 7, 6).byWeek, [4, 6], 'a holiday inside the span carries none of the points (2 + 3 working days)');
}
// cycles do not hang
var sCy2 = autoState([
  { num: 1, feature: 'a', phaseId: 'p1', durDays: 5, deps: [2], capType: 'Development' },
  { num: 2, feature: 'b', phaseId: 'p1', durDays: 5, deps: [1], capType: 'Development' }
], [{ name: 'Solo', capType: 'Development' }]);
var rCy2 = RM.autoTimeline(sCy2, { phaseIds: ['p1'], today: 0 });
ok(rCy2.state.items.every(function (it) { return it.startDay != null; }) && rCy2.notes.some(function (n) { return /cycle/.test(n); }), 'cycle members placed, note added');
// horizon grows
var sHz = autoState([{ num: 1, feature: 'late', phaseId: 'p1', startDay: 48 * 5 + 5, durDays: 5, capType: 'Development' }], [{ name: 'Solo', capType: 'Development' }]);
ok(RM.autoTimeline(sHz, { phaseIds: ['p1'], today: 48 * 5 + 5 }).state.meta.numWeeks >= 50, 'timeline extends to fit');
// a phase's start is a floor: nothing in it is ever placed before it begins
var phPin20 = [{ id: 'p1', name: 'Alpha', bucket: false, startDay: 20 },
  { id: 'p2', name: 'Later', bucket: false }, { id: 'p3', name: 'Next', bucket: true }];
var sPh1 = autoState([{ num: 1, feature: 'a', phaseId: 'p1', durDays: 5, capType: 'Development' }],
  [{ name: 'Solo', capType: 'Development' }], null, phPin20);
eq(byNum(RM.autoTimeline(sPh1, { phaseIds: ['p1'], today: 0 }).state)[1].startDay, 20, 'a pinned phase start floors the auto timeline');
var sPh2 = autoState([
  { num: 1, feature: 'anchor', phaseId: 'p1', startDay: 15, durDays: 5, locked: true, capType: 'Development' },
  { num: 2, feature: 'later', phaseId: 'p1', startDay: 40, durDays: 5, capType: 'Development' }
], [{ name: 'X', capType: 'Development' }, { name: 'Y', capType: 'Development' }]);
eq(byNum(RM.autoTimeline(sPh2, { phaseIds: ['p1'], today: 0 }).state)[2].startDay, 15,
  'an unpinned phase floors where its earliest item already sits, not at today');
var phPin2 = [{ id: 'p1', name: 'Alpha', bucket: false, startDay: 2 },
  { id: 'p2', name: 'Later', bucket: false }, { id: 'p3', name: 'Next', bucket: true }];
var sPh3 = autoState([{ num: 1, feature: 'a', phaseId: 'p1', durDays: 5, capType: 'Development' }],
  [{ name: 'Solo', capType: 'Development' }], null, phPin2);
eq(byNum(RM.autoTimeline(sPh3, { phaseIds: ['p1'], today: 10 }).state)[1].startDay, 10, 'a phase pinned before today does not drag work into the past');
var phPin30 = [{ id: 'p1', name: 'Alpha', bucket: false, startDay: 30 },
  { id: 'p2', name: 'Later', bucket: false }, { id: 'p3', name: 'Next', bucket: true }];
var sPh4 = autoState([{ num: 1, feature: 'under way', phaseId: 'p1', startDay: 5, durDays: 5, capType: 'Development' }],
  [{ name: 'Solo', capType: 'Development' }], null, phPin30);
eq(byNum(RM.autoTimeline(sPh4, { phaseIds: ['p1'], today: 10 }).state)[1].startDay, 5,
  'work already under way keeps its start even when its phase is pinned later');

// placeUnit
section('placeUnit');
var sPl = autoState([
  { num: 1, feature: 'base', phaseId: 'p2', startDay: 0, durDays: 10, capType: 'Development' },
  { num: 2, feature: 'placeme', phaseId: 'p2', deps: [1], startDay: 0, durDays: 5, capType: 'Development' }
], [{ name: 'Solo', capType: 'Development' }]);
var rPl = RM.placeUnit(sPl, sPl.items[1].id, null, { today: 0 });
eq(RM.itemByNum(rPl.state, 2).startDay, 10, 'placeUnit lands right after its dep, in a non-auto phase');
eq(rPl.changed, 1, 'one change');
var rPl2 = RM.placeUnit(rPl.state, sPl.items[1].id, null, { today: 0 });
eq(rPl2.changed, 0, 'already there → no change (its own booking is released first)');
var sPlS = autoState([
  { num: 1, feature: 'f', phaseId: 'p2', startDay: 0, durDays: 10, stories: [
    { num: 101, title: 'a', startDay: 0, durDays: 5, capType: 'Development' },
    { num: 102, title: 'b', startDay: 0, durDays: 5, capType: 'Development' }
  ] }
], [{ name: 'Solo', capType: 'Development' }], { planLevel: 'story' });
var rPlS = RM.placeUnit(sPlS, sPlS.items[0].id, sPlS.items[0].stories[1].id, { today: 0 });
eq(RM.itemByNum(rPlS.state, 1).stories[1].startDay, 5, 'a story places after the week its sibling fills');
var rPlI = RM.placeUnit(autoState([{ num: 1, feature: 'crowd', phaseId: 'p2', startDay: 10, durDays: 5, capType: 'Development', capMult: 3 }], [{ name: 'X', capType: 'Development' }]), null, null, { today: 0 });
eq(rPlI.changed, 0, 'missing item → no change');
// story level: asking for the feature places every one of its stories
var sPlF = autoState([
  { num: 1, feature: 'f', phaseId: 'p2', startDay: 0, durDays: 5, stories: [
    { num: 101, title: 'a', durDays: 5, capType: 'Development' },
    { num: 102, title: 'b', durDays: 5, capType: 'Development' }
  ] }
], [{ name: 'Solo', capType: 'Development' }], { planLevel: 'story' });
var rPlF = RM.placeUnit(sPlF, sPlF.items[0].id, null, { today: 0 });
var plF = RM.itemByNum(rPlF.state, 1);
eq(plF.stories[0].startDay, 0, 'the first story takes week 0');
eq(plF.stories[1].startDay, 5, 'the second waits for week 1');
eq([plF.startDay, plF.durDays], [0, 10], 'and the feature hull covers both');
ok(rPlF.changed > 0, 'placing a story-level feature reports its changes');
// locked and finished work never moves
var sPlL = autoState([
  { num: 1, feature: 'locked', phaseId: 'p2', startDay: 30, durDays: 5, capType: 'Development', locked: true },
  { num: 2, feature: 'finished', phaseId: 'p2', startDay: 30, durDays: 5, capType: 'Development', done: true }
], [{ name: 'Solo', capType: 'Development' }]);
var rPlLk = RM.placeUnit(sPlL, sPlL.items[0].id, null, { today: 0 });
eq(rPlLk.changed, 0, 'a locked feature is not placed');
eq(RM.itemByNum(rPlLk.state, 1).startDay, 30, 'and keeps its dates');
eq(rPlLk.note, 'Locked', 'placeUnit says the feature is locked');
var rPlDn = RM.placeUnit(sPlL, sPlL.items[1].id, null, { today: 0 });
eq(rPlDn.changed, 0, 'a done feature is not placed');
eq(RM.itemByNum(rPlDn.state, 2).startDay, 30, 'and keeps its dates');
ok(/done/.test(rPlDn.note || ''), 'placeUnit says it is already done');
// story level: the feature fan-out skips locked and done stories
var sPlFD = autoState([
  { num: 1, feature: 'f', phaseId: 'p2', startDay: 20, durDays: 5, stories: [
    { num: 101, title: 'shipped', startDay: 20, durDays: 5, capType: 'Development', done: true },
    { num: 102, title: 'todo', startDay: 40, durDays: 5, capType: 'Development' }
  ] }
], [{ name: 'Solo', capType: 'Development' }], { planLevel: 'story' });
var rPlFD = RM.placeUnit(sPlFD, sPlFD.items[0].id, null, { today: 0 });
var plFD = RM.itemByNum(rPlFD.state, 1);
eq(plFD.stories[0].startDay, 20, 'a done story keeps its dates while its feature fans out');
eq(plFD.stories[1].startDay, 20, 'and the open story pulls back to the earliest slot, which is where its phase begins');
var sPlFL = autoState([
  { num: 1, feature: 'f', phaseId: 'p2', startDay: 30, durDays: 5, locked: true, stories: [
    { num: 101, title: 'a', durDays: 5, capType: 'Development' },
    { num: 102, title: 'b', durDays: 5, capType: 'Development' }
  ] }
], [{ name: 'Solo', capType: 'Development' }], { planLevel: 'story' });
var rPlFL = RM.placeUnit(sPlFL, sPlFL.items[0].id, null, { today: 0 });
eq(rPlFL.changed, 0, 'a locked feature places none of its stories');
eq(rPlFL.note, 'Locked', 'and says why');
eq(RM.itemByNum(rPlFL.state, 1).stories[0].startDay, null, 'its stories are left unscheduled');
// a dependency-free milestone is a fixed date, not something to place
var sPlM = autoState([{ num: 1, feature: 'gate', phaseId: 'p2', milestone: true, startDay: 30, durDays: 0 }], [{ name: 'Solo', capType: 'Development' }]);
var rPlM = RM.placeUnit(sPlM, sPlM.items[0].id, null, { today: 0 });
eq(rPlM.changed, 0, 'a dependency-free milestone stays put');
eq(RM.itemByNum(rPlM.state, 1).startDay, 30, 'on its own date');
ok(/dependencies/.test(rPlM.note || ''), 'and placeUnit says why');
// placeUnit honours the same phase floor
var phPin20b = [{ id: 'p1', name: 'Alpha', bucket: false },
  { id: 'p2', name: 'Later', bucket: false, startDay: 20 }, { id: 'p3', name: 'Next', bucket: true }];
var sPhP = autoState([{ num: 1, feature: 'a', phaseId: 'p2', startDay: 60, durDays: 5, capType: 'Development' }],
  [{ name: 'Solo', capType: 'Development' }], null, phPin20b);
eq(RM.itemByNum(RM.placeUnit(sPhP, sPhP.items[0].id, null, { today: 0 }).state, 1).startDay, 20,
  'placeUnit never moves a unit before its phase begins');
var phPinM30 = [{ id: 'p1', name: 'Alpha', bucket: false },
  { id: 'p2', name: 'Later', bucket: false, startDay: 30 }, { id: 'p3', name: 'Next', bucket: true }];
var sPhM = autoState([
  { num: 1, feature: 'work', phaseId: 'p2', startDay: 0, durDays: 5, capType: 'Development' },
  { num: 2, feature: 'gate', phaseId: 'p2', milestone: true, startDay: 60, durDays: 0, deps: [1] }
], [{ name: 'Solo', capType: 'Development' }], null, phPinM30);
eq(RM.itemByNum(RM.placeUnit(sPhM, sPhM.items[1].id, null, { today: 0 }).state, 2).startDay, 30,
  'and a milestone lands at its phase start rather than at its earlier dependency end');
eq(RM.todayDay(META, new Date(Date.UTC(2026, 6, 20))), 0, 'today before the timeline clamps to 0');
eq(RM.todayDay(META, new Date(Date.UTC(2026, 7, 4))), 6, 'today maps to its working-day index');

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
eq(riskSt({ riskScheme: 'none', storyRiskScheme: 'auto' }, []).meta.storyRiskScheme, 'none', 'auto is rejected for stories');
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

// ------------------------------------------------------------- regressions (adversarial review)
section('regressions');
// total calendar helpers
eq(RM.dateToDay(META, RM.parseISO('')), null, 'dateToDay(invalid) is null, not NaN');
eq(RM.fmtISO(new Date(NaN)), '', 'fmtISO(invalid) is empty, not a throw');
// NaN schedule fields sanitized
var sNaN = mkState([{ num: 1, feature: 'x', startDay: NaN, durDays: 5 }]);
eq(sNaN.items[0].startDay, null, 'NaN startDay normalized to null');
// duplicate nums renumbered (first occurrence keeps the number)
var sDup = mkState([
  { num: 5, feature: 'first five' },
  { num: 5, feature: 'second five' },
  { num: 9, feature: 'nine' }
]);
eq(sDup.items[0].num, 5, 'first duplicate keeps its num');
ok(sDup.items[1].num !== 5, 'second duplicate renumbered (' + sDup.items[1].num + ')');
ok(sDup.items[1].num > 9, 'renumber does not collide with existing nums');
// day-granular holidays: a single date blanks exactly one working day
var METAH = JSON.parse(JSON.stringify(META));
METAH.holidays = ['2026-08-05']; // Wednesday of week 1
eq(RM.stretchSpan(METAH, 5, 5), 6, 'single holiday stretches a week-long span by 1 day');
eq(RM.workInSpan(METAH, 5, 6), 5, 'workInSpan skips just that day');
ok(!RM.isBlackoutWeek(METAH, 1), 'a partial-holiday week is NOT a blackout week');
METAH.holidays = ['2026-08-01']; // Saturday
eq(RM.stretchSpan(METAH, 0, 5), 5, 'weekend-dated holiday is ignored');
// legacy whole-week blackouts migrate to five holiday dates
var sMigW = RM.normalizeState({ meta: { timelineStart: '2026-07-27', numWeeks: 8, blackoutWeeks: ['2026-08-03'], holidaysV2026: true }, phases: [{ id: 'p' }], items: [] });
eq(sMigW.meta.holidays.length, 5, 'blackout week migrated to 5 holiday dates');
eq(sMigW.meta.holidays[0], '2026-08-03', 'migration starts at the week\'s Monday');
ok(sMigW.meta.blackoutWeeks === undefined, 'blackoutWeeks field removed');
ok(RM.isBlackoutWeek(sMigW.meta, 1), 'migrated week is fully blacked out');
// 2026 US calendar merges once, then user deletions stick
var sHol = RM.normalizeState({ meta: { timelineStart: '2026-07-27', numWeeks: 8 }, phases: [{ id: 'p' }], items: [] });
ok(sHol.meta.holidays.indexOf('2026-09-07') !== -1 && sHol.meta.holidays.indexOf('2026-12-25') !== -1,
  '2026 US holidays loaded into a fresh document');
RM.clipHolidayRanges(sHol.meta, '2026-09-07', '2026-09-07');
ok(sHol.meta.holidays.indexOf('2026-09-07') === -1 &&
  RM.normalizeState(sHol).meta.holidays.indexOf('2026-09-07') === -1,
  'a deleted holiday stays deleted (merge is one-time)');
// named ranges: migration groups weekend-bridged observances and names them
var thx = sHol.meta.holidayRanges.find(function (r) { return r.start === '2026-11-25'; });
ok(!!thx && thx.end === '2026-11-27' && thx.name === 'Thanksgiving',
  'flat dates migrate into a named Thanksgiving range');
RM.addHolidayRange(sHol.meta, 'Offsite', '2026-08-05', '2026-08-06');
ok(sHol.meta.holidays.indexOf('2026-08-05') !== -1 && sHol.meta.holidays.indexOf('2026-08-06') !== -1,
  'added range expands into the flat date list');
var offIdx = sHol.meta.holidayRanges.findIndex(function (r) { return r.name === 'Offsite'; });
RM.removeHolidayRange(sHol.meta, offIdx);
ok(sHol.meta.holidays.indexOf('2026-08-05') === -1, 'removing a range removes its dates');
// self-dependency flagged, not silent
var sSelf = mkState([{ num: 1, feature: 'ouroboros', deps: [1] }]);
ok((RM.validate(sSelf).byItem[sSelf.items[0].id] || []).some(function (x) { return x.code === 'SELF_DEP'; }),
  'self-dependency produces SELF_DEP warning');
// ------------------------------------------------------------- color
section('colors');
var sCol = mkState([{ num: 1, feature: 'x', workstream: 'Custom A' }, { num: 2, feature: 'y', workstream: 'Custom B' }, { num: 3, feature: 'z' }]);
eq(RM.colorForItem(sCol, sCol.items[0]), RM.PALETTE.neutral, 'unknown workstream without a color is gray');
eq(RM.colorForItem(sCol, sCol.items[2]), RM.PALETTE.neutral, 'no workstream is gray');
sCol.wsColors['Custom B'] = 'process';
eq(RM.colorForItem(sCol, sCol.items[1]), RM.PALETTE.process, 'palette-key workstream color applies');
sCol.wsColors['Custom A'] = '#A14FBF';
eq(RM.colorForItem(sCol, sCol.items[0]), 'A14FBF', 'custom hex workstream color applies');
sCol.wsColors['Custom A'] = 'not-a-color';
eq(RM.colorForItem(sCol, sCol.items[0]), RM.PALETTE.neutral, 'invalid custom color falls back to gray');
var sCol2 = mkState([{ num: 1, feature: 'x', workstream: 'OS' }, { num: 2, feature: 'y', workstream: 'Data' }]);
eq(RM.colorForWs(sCol2, 'OS'), '3273BD', 'seeded default: OS is blue');
eq(RM.colorForWs(sCol2, 'Data'), RM.DEFAULT_WS_COLORS['Data'], 'seeded default: known workstreams get their own color');
eq(RM.iconForEpic(mkState([{ num: 1, feature: 'x', epic: 'OS' }]), 'OS'), 'cpu', 'seeded default: OS epic gets its icon');
sCol2.epicIcons['Search'] = 'rocket';
eq(RM.iconForEpic(sCol2, 'Search'), 'rocket', 'epic icon lookup');
eq(RM.iconForEpic(sCol2, 'Other'), null, 'no icon by default');

// ------------------------------------------------------------- budgeting & reports
section('budgeting & reports');
var sBud = mkState([
  { num: 1, feature: 'a', phaseId: 'p1', workstream: 'WS1', teamType: 'Development', startDay: 0, durDays: 10, headcount: 2 },
  { num: 2, feature: 'b', phaseId: 'p1', workstream: 'WS2', teamType: 'Data', size: 'M' }
], { team: [
  { name: 'Dev', type: 'Development', workstream: 'WS1', rate: 200, cost: 100 },
  { name: 'Analyst', type: 'Data', workstream: 'WS2', rate: 150, cost: 50 }
] });
eq(RM.roleMargin(sBud, sBud.team[0]), 50, 'margin 50% at rate 200 / cost 100');
ok(RM.roleMargin(sBud, { rate: 0, cost: 10, type: 'Nope' }) === null, 'no rate → no margin');
// rate card: a person with no override inherits their role's numbers
sBud.meta.rateCard = { Design: { rate: 180, cost: 90 } };
var rcM = { name: 'D', type: 'Design', rate: 0, cost: 0 };
eq(RM.memberRate(sBud, rcM), 180, 'rate card supplies the role rate');
eq(RM.memberCost(sBud, rcM), 90, 'rate card supplies the role cost');
eq(RM.memberRate(sBud, { type: 'Design', rate: 250, cost: 0 }), 250, 'a personal override beats the card');
eq(RM.roleMargin(sBud, rcM), 50, 'margin computes from effective card numbers');
// actual hours clip each week to its workable (non-holiday) days
var expT = 0;
for (var wq = 0; wq < sBud.meta.numWeeks; wq++) expT += Math.min(40, (5 - RM.holidaysInWeek(sBud.meta, wq)) * 8);
eq(RM.roleTotalHours(sBud, sBud.team[0]), expT, 'role hours = Σ min(planned, workable) per week');
eq(RM.roleWeekHours(sBud, sBud.team[0], 16).planned, 40, 'full-holiday week still plans 40 h');
eq(RM.roleWeekHours(sBud, sBud.team[0], 16).actual, 0, 'full-holiday week yields 0 actual hours');
var w0Act = RM.roleWeekHours(sBud, sBud.team[0], 0).actual;
sBud.team[0].weekHours['2026-07-27'] = 0;
eq(RM.roleTotalHours(sBud, sBud.team[0]), expT - w0Act, 'week-hour override subtracts');
eq(RM.avgCostRate(sBud, 'Data'), 50, 'avg cost rate by team type');
var inf1 = RM.itemEffortInfo(sBud, sBud.items[0]);
eq(inf1.hours, 10 * 8, 'scheduled effort hours = days × hours-per-day');
eq(inf1.cost, inf1.hours * 100, 'item cost priced at the type cost rate');
var repWs = RM.costReport(sBud, 'workstream');
eq(repWs.rows.length, 2, 'workstream report covers both groups');
var inf2 = RM.itemEffortInfo(sBud, sBud.items[1]);
ok(inf2.days > 0 && inf2.cost === inf2.hours * 50, 'unscheduled item priced from its size estimate');
eq(repWs.total.cost, inf1.cost + inf2.cost, 'total item cost sums the groups');
ok(repWs.total.roleCost > 0, 'workstream mode includes roster spend');
eq(RM.costReport(sBud, 'phase').rows[0].items, 2, 'phase report groups items');
eq(RM.costReport(sBud, 'phase-ws').rows.length, 2, 'phase×workstream splits per pair');

// phase window overrides
var sPhw = mkState([{ num: 1, feature: 'x', phaseId: 'p1', startDay: 10, durDays: 5 }]);
eq(RM.phaseSpan(sPhw, sPhw.phases[0]).lo, 10, 'phase span auto-derives lo');
eq(RM.phaseSpan(sPhw, sPhw.phases[0]).hi, 15, 'phase span auto-derives hi');
sPhw.phases[0].startDay = 5;
sPhw.phases[0].endDay = 30;
eq(RM.phaseSpan(sPhw, sPhw.phases[0]).lo, 5, 'pinned phase start wins');
eq(RM.phaseSpan(sPhw, sPhw.phases[0]).hi, 30, 'pinned phase end wins');

// capacity feature off (the default) → roster never constrains anything
var sOff = mkState([{ num: 1, feature: 'big', phaseId: 'p1', size: 'M', headcount: 9 }],
  { team: [{ name: 'solo', type: 'Development' }] });
sOff.meta.capacityEnabled = false;
var vOff = RM.validate(sOff);
ok(!Object.keys(vOff.byItem).some(function (id) {
  return vOff.byItem[id].some(function (w) { return /^HC_|^OVER_CAP/.test(w.code); });
}), 'capacity off: no headcount/over-capacity warnings');

// ------------------------------------------------------------- risk (metadata only)
section('risk metadata');
// risk never pads the schedule: riskDays is always zero, weeks are weeks
var sMig = mkState([{ num: 1, feature: 'legacy', startDay: 0, durDays: 10, leadDays: 5, riskDays: 5, risk: 'S' }]);
eq(sMig.items[0].durDays, 10, 'durDays stays exactly as set');
eq(sMig.items[0].riskDays, 0, 'legacy riskDays zeroed on load');
eq(sMig.items[0].risk, 'L', 'legacy risk t-shirt migrates to severity (S → L)');
eq(RM.itemEnd(sMig.items[0]), 10, 'itemEnd = start + durDays, no padding');
eq(RM.riskEffortDays(sMig, sMig.items[0]), 0, 'risk contributes no effort days');

// DEP_ORDER keys off the plain end
var sRk2 = mkState([
  { num: 1, feature: 'A', startDay: 0, durDays: 10 },
  { num: 2, feature: 'B', startDay: 8, durDays: 5, size: 'S', deps: [1] }
]);
var vRk = RM.validate(sRk2);
ok((vRk.byItem[sRk2.items[1].id] || []).some(function (v) { return v.code === 'DEP_ORDER'; }),
  'starting before a dependency ends is flagged');

// ------------------------------------------------------------- resources / time off
section('time off');
var sOff = mkState([
  { num: 1, feature: 'chunky work', startDay: 0, durDays: 15, capType: 'Development', capMult: 2 }
], {
  team: [
    { id: 'm1', name: 'Ada', capType: 'Development', offWeeks: ['2026-07-27'] },
    { id: 'm2', name: 'Grace', capType: 'Development' }
  ]
});
var capOff = RM.capacity(sOff);
eq(capOff.weeks[0].supply, 1, 'off member lowers week 0 supply');
eq(capOff.weeks[1].supply, 2, 'week 1 back to full roster');
ok(capOff.weeks[0].over, '2 people asked in the short-handed week is over');
ok(!capOff.weeks[1].over, 'the full roster absorbs the same ask');

// ------------------------------------------------------------- hours model
section('hours');
var sHr = mkState([
  { num: 1, feature: 'w', startDay: 0, durDays: 10, capType: 'Development' }
], {
  team: [
    { id: 'h1', name: 'Ada', capType: 'Development', weekHours: { '2026-07-27': 20 } },
    { id: 'h2', name: 'Grace', capType: 'Development' }
  ]
});
eq(RM.memberHoursForWeek(sHr.meta, sHr.team[0], 0), 20, 'explicit week hours read back');
eq(RM.memberHoursForWeek(sHr.meta, sHr.team[0], 1), 40, 'unlisted weeks default to 40');
eq(RM.memberHeads(sHr, sHr.team[0], 0), 0.5, 'people-equivalents: 20 h of a 40 h week is half a head');
eq(RM.capSupply(sHr).byType.Development[0], 1.5, 'the typed pool sums to 1.5 people');
ok(!RM.capacity(sHr).weeks[0].over, 'one person asked of a 1.5-person pool is fine');
var sHrSolo = mkState([
  { num: 1, feature: 'w', startDay: 0, durDays: 10, capType: 'Development' }
], { team: [{ id: 'h1', name: 'Ada', capType: 'Development', weekHours: { '2026-07-27': 20 } }] });
ok(RM.capacity(sHrSolo).weeks[0].over, 'one person asked of 0.5 available is over');
// legacy offWeeks migrate to zero-hour weeks
var sOffMig = RM.normalizeState({
  meta: { timelineStart: '2026-07-27', numWeeks: 8 }, phases: [{ id: 'p' }], items: [],
  team: [{ name: 'Ada', type: 'Development', offWeeks: ['2026-08-03'] }]
});
eq(sOffMig.team[0].weekHours['2026-08-03'], 0, 'offWeeks -> 0-hour week');
ok(RM.memberOffWeek(sOffMig.meta, sOffMig.team[0], 1), 'memberOffWeek still answers via hours');
// items default to 1 × any role (empty) — a removed role leaves nothing behind
eq(mkState([{ num: 1, feature: 'x' }]).items[0].teamType, '', 'default work type is none');

// capacity factor: PE at 40h scales availability
var sCf = mkState([], { team: [{ name: 'Half', capType: 'Development', capacity: 0.5 }] });
eq(RM.capSupply(sCf).byType.Development[0], 0.5, 'capacity 0.5 at 40h = half a head');
eq(sCf.team[0].capacity, 0.5, 'capacity survives normalize');
eq(mkState([], { team: [{ name: 'X', type: 'Development' }] }).team[0].capacity, 1, 'capacity defaults to 1');
eq(mkState([], { team: [{ name: 'Z', type: 'Development', capacity: 0 }] }).team[0].capacity, 0, 'capacity 0 is allowed');
eq(RM.capSupply(mkState([], { team: [{ name: 'Z', capType: 'Development', capacity: 0 }] })).types, [],
  'a zero-capacity person supplies nothing — like an untyped one');
// story-points mode: the seat multiplier is a per-person concept and does not
// scale points (the Resources panel shows one or the other); fewer hours still do
{
  var mP = JSON.parse(JSON.stringify(META)); mP.capMode = 'points'; mP.defaultPoints = 10; mP.weeksPerSprint = 2;
  var sP = mkState([], { meta: mP, team: [{ name: 'Half', capType: 'Development', capacity: 0.5 }] });
  eq(RM.capSupply(sP).byType.Development[0], 5, 'points mode: 10 pt / 2 weeks = 5 pt a week, ignoring the 0.5 seat');
  var sP0 = mkState([], { meta: mP, team: [{ name: 'Off', capType: 'Development', capacity: 0 }, { name: 'On', capType: 'Development' }] });
  eq(RM.capSupply(sP0).byType.Development[0], 5, 'points mode: a seat of 0 still means "supplies nothing" (only On\'s 5 a week)');
  var sPh = mkState([], { meta: mP, team: [{ name: 'PT', capType: 'Development', weekHours: { '2026-07-27': 20 } }] });
  eq(RM.capSupply(sPh).byType.Development[0], 2.5, 'points mode: 20 of 40 hours halves the points supply');
}

// ------------------------------------------------------------- renumbering
section('renumber');
var sRn = mkState([
  { num: 1, feature: 'a' },
  { num: 2, feature: 'b', deps: [1] },
  { num: 5, feature: 'c' }
]);
eq(RM.renumberItem(sRn, sRn.items[0].id, 9), 9, 'free number accepted');
eq(sRn.items[1].deps, [9], 'deps follow the rename');
eq(RM.renumberItem(sRn, sRn.items[2].id, 9), 10, 'taken number falls back to next available');
eq(RM.renumberItem(sRn, sRn.items[2].id, 'zap'), 11, 'invalid number falls back to next available');

// ------------------------------------------------------------- project end date
section('end date');
var sEd = mkState([]);
eq(sEd.meta.endDate, RM.fmtISO(RM.dayToDate(sEd.meta, sEd.meta.numWeeks * 5 - 1)),
  'endDate synced to the last working day');
var sEd2 = RM.normalizeState({
  meta: { timelineStart: '2026-07-27', numWeeks: 48, endDate: '2026-10-16' }, // Fri of week 11
  phases: [{ id: 'p' }], items: []
});
eq(sEd2.meta.numWeeks, 12, 'saved endDate wins over numWeeks');
eq(sEd2.meta.endDate, '2026-10-16', 'endDate normalizes to that week\'s Friday');

// ------------------------------------------------------------- scoping columns
section('scope columns');
var sSc = mkState([{ num: 1, feature: 'a' }]);
eq(sSc.meta.scopeCols.map(function (c) { return c.key; }),
  ['description', 'ac'], 'new documents start with Description and Acceptance criteria');
// legacy docs without a saved column list surface built-ins that hold content
var sScLegacy = mkState([{ num: 1, feature: 'a', enables: 'x', notes: 'y' }]);
eq(sScLegacy.meta.scopeCols.map(function (c) { return c.key; }),
  ['description', 'ac', 'enables', 'notes'], 'legacy content infers its columns');
eq(sSc.meta.scopeCols[1].scope, 'story', 'Acceptance criteria defaults to stories only');
eq(sSc.meta.scopeColOrder.indexOf('ac'), sSc.meta.scopeColOrder.indexOf('description') + 1, 'Acceptance criteria orders right after Description');
RM.removeScopeCol(sSc, 'ac');
var ck = RM.addScopeCol(sSc, 'Owner');
eq(sSc.meta.scopeCols.length, 2, 'custom column appended');
eq(RM.scopeColLabel(sSc.meta.scopeCols[1]), 'Owner', 'custom label kept');
RM.setScopeValue(sSc.items[0], ck, 'Rita');
eq(RM.scopeValue(sSc.items[0], ck), 'Rita', 'custom value stored in item.custom');
RM.setScopeValue(sSc.items[0], 'notes', 'n1');
eq(sSc.items[0].notes, 'n1', 'built-in key routes to the item field');
RM.addScopeCol(sSc, null, 'notes');
RM.moveScopeCol(sSc, ck, 1);
eq(sSc.meta.scopeCols[2].key, ck, 'column moved');
RM.removeScopeCol(sSc, 'description');
eq(sSc.meta.scopeCols.length, 2, 'Description is removable like any column');
RM.addScopeCol(sSc, null, 'description');
eq(sSc.meta.scopeCols[2].key, 'description', 'hidden built-in re-added');
RM.addScopeCol(sSc, null, 'description');
eq(sSc.meta.scopeCols.length, 3, 're-adding a visible built-in is a no-op');
RM.removeScopeCol(sSc, ck);
eq(sSc.meta.scopeCols.length, 2, 'column removed');
eq(RM.scopeValue(sSc.items[0], ck), '', 'custom values cleaned up on remove');
RM.renameScopeCol(sSc, 'notes', 'Field notes');
eq(RM.scopeColLabel(sSc.meta.scopeCols[0]), 'Field notes', 'built-in columns are renamable');
var sSc2 = RM.normalizeState(sSc);
eq(sSc2.meta.scopeCols.map(function (c) { return c.key; }),
  sSc.meta.scopeCols.map(function (c) { return c.key; }), 'scopeCols survive normalize');
eq(RM.scopeColLabel(sSc2.meta.scopeCols[0]), 'Field notes', 'built-in rename survives normalize');
RM.renameScopeCol(sSc, 'notes', '');
eq(RM.scopeColLabel(sSc.meta.scopeCols[0]), 'Notes', 'clearing a rename restores the canonical name');
RM.setScopeValue(sSc.items[0], 'x9', 'keep');
// a saved doc keeps its column list: removing Acceptance criteria sticks
var sScNoAc = RM.normalizeState(RM.normalizeState(sSc));
ok(!sScNoAc.meta.scopeCols.some(function (c) { return c.key === 'ac'; }), 'a removed Acceptance criteria column stays removed on reload');
// pre-migration docs with a custom "Acceptance Criteria" column hand it over to the built-in
var sAcMig = RM.normalizeState({
  meta: { timelineStart: '2026-07-27', scopeCols: [{ key: 'description' }, { key: 'cx', label: 'Acceptance Criteria', scope: 'story' }, { key: 'notes' }],
    scopeColOrder: ['description', 'notes', 'cx', 'epic'] },
  phases: [{ id: 'p' }],
  items: [{ num: 1, feature: 'a', custom: { cx: 'feat crit' }, stories: [{ title: 's', custom: { cx: '<ul><li>crit</li></ul>' } }] }]
});
eq(sAcMig.meta.scopeCols.map(function (c) { return c.key; }), ['description', 'ac', 'notes'], 'custom Acceptance Criteria column becomes the built-in in place');
eq(sAcMig.meta.scopeCols[1].scope, 'story', 'migrated column keeps its scope');
eq(sAcMig.meta.scopeColOrder.indexOf('ac'), sAcMig.meta.scopeColOrder.indexOf('notes') + 1, 'migrated column keeps its position');
eq(sAcMig.items[0].stories[0].ac, '<ul><li>crit</li></ul>', 'story values move to story.ac');
eq(sAcMig.items[0].stories[0].custom.cx, undefined, 'old custom story value removed');
eq(sAcMig.items[0].ac, 'feat crit', 'feature values move to item.ac');
eq(RM.storyScopeValue(sAcMig.items[0].stories[0], 'ac'), '<ul><li>crit</li></ul>', 'storyScopeValue reads ac as a story field');
RM.setStoryScopeValue(sAcMig.items[0].stories[0], 'notes', 'n');
eq(sAcMig.items[0].stories[0].custom.notes, 'n', 'other columns land in story.custom');
// a doc saved before the built-in existed gets it after Description, stories only
var sAcNew = RM.normalizeState({ meta: { timelineStart: '2026-07-27', scopeCols: [{ key: 'notes' }, { key: 'description' }], scopeColOrder: ['notes', 'description', 'epic'] }, phases: [{ id: 'p' }], items: [] });
eq(sAcNew.meta.scopeCols.map(function (c) { return c.key; }), ['notes', 'description', 'ac'], 'existing docs gain Acceptance criteria after Description');
eq(sAcNew.meta.scopeColOrder.slice(0, 3), ['notes', 'description', 'ac'], 'saved column order slots it after Description');
RM.setScopeColScope(sAcNew, 'ac', 'both');
eq(RM.normalizeState(sAcNew).meta.scopeCols[2].scope, 'both', 'choosing both for Acceptance criteria survives reload');
ok(RM.scopeColShows({ key: 'ac', scope: 'both' }, 'feature'), 'an explicit both shows on features');
eq(RM.normalizeState(sSc).items[0].custom.x9, 'keep', 'custom values survive normalize');

// ------------------------------------------------------------- column scope
section('column scope');
var sCol = mkState([{ num: 1, feature: 'a' }]);
RM.removeScopeCol(sCol, 'ac');
RM.addScopeCol(sCol, null, 'notes');
RM.addScopeCol(sCol, null, 'enables');
sCol.meta.scopeCols[0].scope = 'story';
sCol.meta.scopeCols[1].scope = 'bogus';
var sCol2 = RM.normalizeState(sCol);
eq(sCol2.meta.scopeCols[0].scope, 'story', 'a column scope survives normalize');
eq(sCol2.meta.scopeCols[1].scope, undefined, 'an unknown scope drops back to both');
eq(sCol2.meta.scopeCols[2].scope, undefined, 'missing scope stays absent (both)');
eq(RM.scopeColShows({ key: 'x' }, 'feature'), true, 'no scope shows on features');
eq(RM.scopeColShows({ key: 'x' }, 'story'), true, 'no scope shows on stories');
eq(RM.scopeColShows({ key: 'x', scope: 'story' }, 'feature'), false, 'story-only hides on features');
eq(RM.scopeColShows({ key: 'x', scope: 'story' }, 'story'), true, 'story-only shows on stories');
eq(RM.scopeColShows({ key: 'x', scope: 'feature' }, 'story'), false, 'feature-only hides on stories');
RM.setScopeColScope(sCol2, 'notes', 'feature');
eq(sCol2.meta.scopeCols[1].scope, 'feature', 'setScopeColScope stores the scope');
RM.setScopeColScope(sCol2, 'notes', 'both');
eq(sCol2.meta.scopeCols[1].scope, undefined, 'both clears the scope');
RM.setScopeColScope(sCol2, 'notes', 'nope');
eq(sCol2.meta.scopeCols[1].scope, undefined, 'an unknown scope is ignored');

// ------------------------------------------------------------- milestone styles
section('milestone styles');
eq(RM.MS_STYLES, ['diamond', 'star', 'circle'], 'three milestone styles');
var sSty = mkState([
  { num: 1, feature: 'star', milestone: true, startDay: 0, durDays: 0, msStyle: 'star' },
  { num: 2, feature: 'odd', milestone: true, startDay: 0, durDays: 0, msStyle: 'hexagon' },
  { num: 3, feature: 'plain', milestone: true, startDay: 0, durDays: 0 },
  { num: 4, feature: 'bar', startDay: 0, durDays: 5, msStyle: 'circle' }
]);
eq(sSty.items[0].msStyle, 'star', 'star style survives normalize');
eq(sSty.items[1].msStyle, undefined, 'an unknown style is dropped');
eq(sSty.items[2].msStyle, undefined, 'no style stays absent');
eq(sSty.items[3].msStyle, 'circle', 'a bar remembers its style for when it becomes a milestone');
eq(RM.msStyleOf(sSty.items[0]), 'star', 'msStyleOf reads the style');
eq(RM.msStyleOf(sSty.items[2]), 'diamond', 'msStyleOf defaults to diamond');

// ------------------------------------------------------------- milestones
section('milestones');
var sMs = mkState([
  { num: 1, feature: 'work', startDay: 0, durDays: 10, capType: 'Development' },
  { num: 2, feature: 'launch', milestone: true, startDay: 10, durDays: 0, deps: [1] },
  { num: 3, feature: 'after', startDay: 10, durDays: 5, deps: [2], capType: 'Development' }
], { team: [{ name: 'X', capType: 'Development' }] });
eq(sMs.items[1].milestone, true, 'milestone flag survives normalize');
eq(sMs.items[1].durDays, 0, 'milestone keeps zero duration');
eq(RM.itemSpan(sMs.items[1]), 0, 'milestone span is zero');
eq(RM.itemEnd(sMs.items[1]), 10, 'milestone end equals its day');
var vMs = RM.validate(sMs);
var msWarns = (vMs.byItem[sMs.items[1].id] || []).map(function (v) { return v.code; });
eq(msWarns.indexOf('NO_SIZE'), -1, 'no size warnings at all');
var depWarns = (vMs.byItem[sMs.items[2].id] || []).map(function (v) { return v.code; });
eq(depWarns.indexOf('DEP_ORDER'), -1, 'dependent may start on the milestone day');
var msCap = RM.capacity(sMs);
eq(msCap.weeks[2].demand, 1, 'milestone consumes no capacity (only the 5-day item registers)');
var sMsStory = mkState([{ num: 1, feature: 'a', stories: [{ title: 's', custom: { c1: 'v' } }] }]);
eq(sMsStory.items[0].stories[0].custom.c1, 'v', 'story custom fields survive normalize');

// ------------------------------------------------------------- work week & costs
section('work week & costs');
var sWw = mkState([{ num: 1, feature: 'x' }]);
sWw.meta.workDays = [1, 2, 3, 4]; // Mon-Thu -> 4 slots per index week
sWw.meta.weekHours = 32;
eq(RM.slotsOf(sWw.meta), 4, 'a 4-day week has 4 slots per index week');
eq(RM.workInSpan(sWw.meta, 0, 4), 4, 'one index week holds exactly its working days');
eq(RM.stretchSpan(sWw.meta, 0, 5), 5, 'no phantom off slots — 5 work days = 5 slots');
eq(RM.fmtISO(RM.dayToDate(sWw.meta, 4)), '2026-08-03', 'slot 4 = the second week Monday (Mon-Thu week)');
eq(RM.hoursPerDay(sWw.meta), 8, '32 h over 4 days = 8 h/day');
eq(RM.memberHoursForWeek(sWw.meta, {}, 0), 32, 'default member week = the project full-time week');
sWw.team = [{ id: 't1', name: 'A', capType: 'Development', capacity: 1, weekHours: {} }];
eq(RM.capSupply(sWw).byType.Development[0], 1, '32 h at a 32 h full-time week = 1 person');
var sWw5 = RM.normalizeState({ meta: { timelineStart: '2026-07-27', numWeeks: 8, weekHours: 32, daysPerWeek: 4 }, phases: [{ id: 'p' }], items: [] });
eq(sWw5.meta.weekHours, 32, 'weekHours survives normalize');
eq(sWw5.meta.daysPerWeek, 4, 'daysPerWeek survives normalize');
eq(sWw5.meta.workDays.join(','), '1,2,3,4', 'legacy daysPerWeek 4 migrates to Mon-Thu');
eq(sWw5.meta.weekStart, 1, 'week starts Monday by default');

// which weekdays work + first day of week
var sWd = RM.normalizeState({
  meta: { timelineStart: '2026-07-28', numWeeks: 8, weekStart: 0, workDays: [0, 1, 2, 3, 4] },
  phases: [{ id: 'p' }], items: []
});
eq(sWd.meta.timelineStart, '2026-07-26', 'timeline start snaps back to the first day of the week (Sunday)');
eq(RM.fmtISO(RM.dayToDate(sWd.meta, 0)), '2026-07-26', 'slot 0 = Sunday');
eq(RM.fmtISO(RM.dayToDate(sWd.meta, 4)), '2026-07-30', 'slot 4 = Thursday (Sun-Thu week)');
eq(RM.dateToDay(sWd.meta, RM.parseISO('2026-07-31')), 4, 'a non-working Friday maps back to the last working slot');
var sWd2 = RM.normalizeState({
  meta: { timelineStart: '2026-07-27', numWeeks: 8, workDays: [2, 3, 4, 5, 6] }, // Tue-Sat
  phases: [{ id: 'p' }], items: []
});
eq(RM.fmtISO(RM.dayToDate(sWd2.meta, 0)), '2026-07-28', 'Tue-Sat week: slot 0 lands on Tuesday');
eq(RM.fmtISO(RM.dayToDate(sWd2.meta, 4)), '2026-08-01', 'Tue-Sat week: slot 4 lands on Saturday');
// a holiday that falls on a non-working weekday is ignored
var sWd3 = RM.normalizeState({
  meta: { timelineStart: '2026-07-27', numWeeks: 8, workDays: [1, 2, 3, 4], holidaysV2026: true,
    holidayRanges: [{ name: 'F', start: '2026-07-31', end: '2026-07-31' }, { name: 'T', start: '2026-07-28', end: '2026-07-28' }] },
  phases: [{ id: 'p' }], items: []
});
var wd3set = RM.holidayDaySet(sWd3.meta);
ok(!wd3set[4], 'a Friday holiday in a Mon-Thu week does not exist in the index space');
ok(wd3set[1], 'a Tuesday holiday registers on its slot');

var sCst = mkState([{ num: 1, feature: 'x' }]);
sCst.costs = [
  { id: 'c1', name: 'License', amount: 1000, kind: 'fixed', startDay: 10 },
  { id: 'c2', name: 'Cloud', amount: 50, kind: 'weekly', startDay: 0, endDay: 20 },
  { id: 'c3', name: 'Rent', amount: 300, kind: 'monthly', startDay: 0, endDay: null }
];
sCst = RM.normalizeState(sCst);
eq(RM.costOccurrences(sCst, sCst.costs[0]).length, 1, 'fixed cost hits once');
eq(RM.costOccurrences(sCst, sCst.costs[1]).length, 5, 'weekly cost repeats to its end day');
eq(RM.costTotal(sCst, sCst.costs[1]), 250, 'weekly total sums occurrences');
ok(RM.costOccurrences(sCst, sCst.costs[2]).length >= 3, 'monthly cost repeats to the timeline end');
eq(RM.costsTotal(sCst), 1000 + 250 + RM.costTotal(sCst, sCst.costs[2]), 'grand total sums all costs');

// ------------------------------------------------------------- dependency risk
section('depRisk');
var sDr = mkState([
  { num: 1, feature: 'free' },
  { num: 2, feature: 'waiting', deps: [1] },
  { num: 3, feature: 'cyc a', deps: [4] },
  { num: 4, feature: 'cyc b', deps: [3] }
]);
eq(RM.depRisk(sDr, sDr.items[0]).level, 'none', 'no deps → none');
ok(RM.depRisk(sDr, sDr.items[1]).level !== 'none', 'unscheduled dep raises risk');
eq(RM.depRisk(sDr, sDr.items[2]).level, 'high', 'cycle membership → high');

// ------------------------------------------------------------- sprint anchor
section('sprint numbering');
var sSp = mkState([]);
sSp.meta.sprintAnchor = '2026-09-07'; // 6 weeks after timeline start
sSp.meta.sprintAnchorNum = 1;
eq(RM.sprintNumForWeek(sSp.meta, 6), 1, 'anchor week is S1');
eq(RM.sprintNumForWeek(sSp.meta, 8), 2, 'next sprint is S2');
eq(RM.sprintNumForWeek(sSp.meta, 4), 0, 'sprint before the anchor is S0');
eq(RM.sprintNumForWeek(sSp.meta, 0), -2, 'weeks count down before the anchor');
eq(RM.sprintInfo(mkState([]).meta).anchorWeek, 0, 'default anchor is the timeline start');

// ------------------------------------------------------------- ordering & ripple
section('ordering & ripple');
var sOrd = mkState([
  { num: 1, feature: 'late', phaseId: 'p1', startDay: 10, durDays: 5 },
  { num: 2, feature: 'early', phaseId: 'p1', startDay: 0, durDays: 5 },
  { num: 3, feature: 'backlog', phaseId: 'p1' },
  { num: 4, feature: 'early too', phaseId: 'p1', startDay: 0, durDays: 5 }
]);
RM.sortItemsByStart(sOrd);
eq(sOrd.items.map(function (x) { return x.num; }), [2, 4, 1, 3], 'stable sort by start; unscheduled last');

// holdPos: a freshly-inserted unscheduled row keeps its slot through sorts
var sHold = mkState([
  { num: 1, feature: 'late', phaseId: 'p1', startDay: 10, durDays: 5 },
  { num: 2, feature: 'inserted', phaseId: 'p1' },
  { num: 3, feature: 'early', phaseId: 'p1', startDay: 0, durDays: 5 }
]);
sHold.items[1].holdPos = true;
RM.sortItemsByStart(sHold);
eq(sHold.items.map(function (x) { return x.num; }), [3, 2, 1], 'holdPos row stays at its slot instead of sorting to the bottom');
RM.sortItemsByStart(sHold);
eq(sHold.items.map(function (x) { return x.num; }), [3, 2, 1], 'holdPos survives repeated sorts while undated');
sHold.items[1].startDay = 20; sHold.items[1].durDays = 5;
RM.sortItemsByStart(sHold);
eq(sHold.items.map(function (x) { return x.num; }), [3, 1, 2], 'once dated the row sorts normally');
ok(!sHold.items[2].holdPos, 'holdPos clears itself after a date is set');

var sRip = mkState([
  { num: 1, feature: 'root', startDay: 0, durDays: 5 },
  { num: 2, feature: 'child', deps: [1], startDay: 5, durDays: 5 },
  { num: 3, feature: 'grandchild', deps: [2], startDay: 10, durDays: 5 },
  { num: 4, feature: 'pinned', deps: [1], startDay: 5, durDays: 5, locked: true },
  { num: 5, feature: 'unrelated', startDay: 0, durDays: 5 }
]);
// the caller (drag end) moves the root first; ripple then chains the push down
sRip.items[0].startDay = 5;
var movedN = RM.shiftDependents(sRip, sRip.items[0].id, 5);
eq(movedN, 2, 'ripple moves the two unlocked dependents');

// story timelines ride along when their feature moves in time
var sSt = mkState([
  { num: 1, feature: 'root', startDay: 0, durDays: 5 },
  { num: 2, feature: 'child', deps: [1], startDay: 5, durDays: 5,
    stories: [{ id: 'sx', title: 'story', startDay: 6, durDays: 5 }] }
]);
sSt.items[0].startDay = 5;
RM.shiftDependents(sSt, sSt.items[0].id, 5);
eq(sSt.items[1].startDay, 10, 'ripple moved the child');
eq(sSt.items[1].stories[0].startDay, 11, "the child's story rode along");
RM.shiftStories(sSt.items[1], -20);
eq(sSt.items[1].stories[0].startDay, 0, 'shiftStories clamps at day 0');
eq(sRip.items[1].startDay, 10, 'child pushed forward');
eq(sRip.items[2].startDay, 15, 'grandchild pushed to follow the child');
eq(sRip.items[3].startDay, 5, 'locked dependent stays');
eq(sRip.items[4].startDay, 0, 'unrelated item stays');

// slack absorbs part of the push: dependents move only as far as needed
var sRip2 = mkState([
  { num: 1, feature: 'root', startDay: 5, durDays: 5 },      // already moved +5
  { num: 2, feature: 'slack', deps: [1], startDay: 12, durDays: 5 },
  { num: 3, feature: 'tail', deps: [2], startDay: 17, durDays: 5 }
]);
eq(RM.shiftDependents(sRip2, sRip2.items[0].id, 5), 0, 'push inside slack moves nothing');
sRip2.items[0].startDay = 10; // +5 more — now 2 days past the slack
eq(RM.shiftDependents(sRip2, sRip2.items[0].id, 5), 2, 'push past slack chains down');
eq(sRip2.items[1].startDay, 15, 'dependent pushed just to the buffered end');
eq(sRip2.items[2].startDay, 20, 'its dependent pushed by the same applied amount');

// pulling back: dependents follow, clamped by their other dependencies
var sRip3 = mkState([
  { num: 1, feature: 'root', startDay: 0, durDays: 5 },      // already moved −5 (was 5)
  { num: 2, feature: 'other', startDay: 0, durDays: 8 },
  { num: 3, feature: 'child', deps: [1, 2], startDay: 10, durDays: 5 }
]);
eq(RM.shiftDependents(sRip3, sRip3.items[0].id, -5), 1, 'pull moves the dependent back');
eq(sRip3.items[2].startDay, 8, 'clamped at the other dependency\'s end');

// ------------------------------------------------------------- critical path
section('critical path');
var sCp = mkState([
  { num: 1, feature: 'A', startDay: 0, durDays: 10 },
  { num: 2, feature: 'B', startDay: 10, durDays: 10, deps: [1] },
  { num: 3, feature: 'C', startDay: 0, durDays: 2 }
]);
var cp = RM.criticalPath(sCp);
ok(cp.edges[sCp.items[0].id + '>' + sCp.items[1].id], 'A→B is on the critical path');
ok(cp.items[sCp.items[0].id] && cp.items[sCp.items[1].id], 'A and B are critical');
ok(!cp.items[sCp.items[2].id], 'short independent C is not critical');
eq(cp.total, 20, 'critical chain length');

// ------------------------------------------------------------- rich text
section('rich text');
eq(RM.htmlToText('<b>Hi</b><br>there &amp; <i>more</i>'), 'Hi\nthere & more', 'htmlToText strips tags and decodes entities');
eq(RM.htmlToText('<ul><li>a</li><li>b</li></ul>'), 'a\nb', 'htmlToText keeps list items on their own lines');
eq(RM.htmlToText(''), '', 'htmlToText of empty is empty');
var sRich = mkState([{ num: 1, feature: 'F', stories: [
  { title: 'S1', description: '<b>body</b>', ac: '<ul><li>crit</li></ul>' },
  { title: 'S2' }
] }]);
eq(sRich.items[0].stories[0].description, '<b>body</b>', 'story description survives normalize');
eq(sRich.items[0].stories[0].ac, '<ul><li>crit</li></ul>', 'story acceptance criteria survive normalize');
eq(sRich.items[0].stories[1].description, '', 'missing story description defaults empty');
eq(sRich.items[0].stories[1].ac, '', 'missing story acceptance criteria default empty');

// ------------------------------------------------------------- png export layout
section('png export layout');
var RMExport = require('../js/export-png.js');
(function () {
  var sEx = mkState([
    { num: 1, feature: 'Alpha item', workstream: 'Product', epic: 'OS', startDay: 0, durDays: 10 },
    { num: 2, feature: 'Beta item', workstream: 'Data', epic: 'OS', startDay: 20, durDays: 5, riskDays: 5 },
    { num: 3, feature: 'Backlog idea', workstream: 'Product', epic: 'OS' }, // unscheduled
    { num: 4, feature: 'Next thing', phaseId: 'p2', workstream: 'Process', epic: 'Ops', startDay: 40, durDays: 10 }
  ]);
  // normalize zeroes riskDays; only the auto-scheduler sets them at runtime
  sEx.items.forEach(function (it) { if (it.num === 2) it.riskDays = 5; });
  var itemRows = function (lay) {
    return lay.rows.filter(function (r) { return r.kind === 'item'; });
  };

  var full = RMExport.layout(sEx, {});
  // content spans days 0..50 → weeks 0..10; the window clamps to that span
  eq(full.d0, 0, 'window starts at the first visible bar');
  eq(full.d1, 50, 'window ends at the last visible bar');
  eq(full.width, full.laneX + 10 * full.weekPx, 'export width clamps to the content span');
  eq(itemRows(full).map(function (r) { return r.feature; }),
    ['Alpha item', 'Beta item', 'Next thing'], 'scheduled items export; unscheduled are omitted');
  eq(full.rows.filter(function (r) { return r.kind === 'band'; }).map(function (r) { return r.name; }),
    ['Alpha', 'Next'], 'each phase with visible items contributes a band row');
  var bar1 = itemRows(full)[0].bar, bar2 = itemRows(full)[1].bar;
  eq(bar1.x, full.laneX, 'day 0 bar starts at the lane origin');
  eq(bar1.w, 10 * full.weekPx / 5, 'bar width is durDays at a fifth of weekPx per day');
  eq(bar2.w, (5 + 5) * full.weekPx / 5, 'risk days extend the painted bar');
  ok(full.height >= full.rows[full.rows.length - 1].y + full.rows[full.rows.length - 1].h,
    'canvas height covers the last row');

  var ws = RMExport.layout(sEx, { ws: 'Data' });
  eq(itemRows(ws).map(function (r) { return r.feature; }), ['Beta item'], 'workstream filter keeps only matching items');
  eq(ws.d0, 20, 'filters inform the exported window start');
  eq(ws.d1, 30, 'filters inform the exported window end');
  eq(ws.rows.filter(function (r) { return r.kind === 'band'; }).map(function (r) { return r.name; }),
    ['Alpha'], 'phases emptied by a filter drop their band row');

  var ep = RMExport.layout(sEx, { epic: 'Ops' });
  eq(itemRows(ep).map(function (r) { return r.feature; }), ['Next thing'], 'epic filter keeps only matching items');

  var ph = RMExport.layout(sEx, { phaseId: 'p1' });
  eq(itemRows(ph).map(function (r) { return r.feature; }), ['Alpha item', 'Beta item'], 'phase filter keeps only that phase');

  // sprints are 2 weeks here: sprint 3 covers weeks 4–5, days 20–29
  var rng = RMExport.layout(sEx, { fromSprint: 3, toSprint: 3 });
  eq(rng.d0, 20, 'sprint range sets the day window start');
  eq(rng.d1, 30, 'sprint range sets the day window end');
  eq(rng.width, rng.laneX + 2 * rng.weekPx, 'ranged export width covers only the selected weeks');
  eq(itemRows(rng).map(function (r) { return r.feature; }), ['Beta item'], 'bars outside the window are dropped');
  eq(itemRows(rng)[0].bar.x, rng.laneX + (20 - 20) * rng.weekPx / 5, 'ranged bar x is relative to the window');
  var clip = RMExport.layout(sEx, { fromSprint: 1, toSprint: 1 }); // days 0–19
  eq(itemRows(clip).map(function (r) { return r.feature; }), ['Alpha item'], 'window keeps overlapping bars only');
  var rng2 = RMExport.layout(sEx, { fromSprint: 5, toSprint: 6 }); // days 40–59
  var b4 = itemRows(rng2)[0].bar;
  eq(b4.w, 10 * rng2.weekPx / 5, 'bar fully inside the window keeps its width');
  var part = RMExport.layout(sEx, { fromSprint: 2, toSprint: 2 }); // days 10–19: Alpha item runs 0–9, Beta starts day 20
  eq(itemRows(part).length, 0, 'bars touching neither side of the window are dropped');
  eq(part.width, part.laneX, 'an empty window shows no date columns');
  eq(part.sprints.length, 0, 'an empty window has no sprint header');
  eq(part.range, '', 'an empty window shows no date range');
  var edge = RMExport.layout(sEx, { fromSprint: 1, toSprint: 1 });
  // Beta item (20..29) is out; Alpha (0..9) fully in
  eq(itemRows(edge)[0].bar.w, 10 * edge.weekPx / 5, 'unclipped bar keeps full width at window edge');
  var clip2 = RMExport.layout(mkState([
    { num: 9, feature: 'Spans', workstream: 'Product', epic: 'OS', startDay: 5, durDays: 30 }
  ]), { fromSprint: 2, toSprint: 2 }); // window days 10–19, bar runs 5..34
  eq(itemRows(clip2)[0].bar.x, clip2.laneX, 'bar entering from the left clips to the window start');
  eq(itemRows(clip2)[0].bar.w, 10 * clip2.weekPx / 5, 'clipped bar width covers only the visible days');

  // sprints 3–4 select weeks 4–7, but content (Beta, days 20–29) ends at week 6
  var spr = RMExport.layout(sEx, { fromSprint: 3, toSprint: 4 });
  eq(spr.d1, 30, 'range and content together bound the window');
  eq(spr.sprints.length, 1, 'sprint header cells cover only the clamped window');
  eq(spr.sprints[0].x, spr.laneX, 'first sprint cell starts at the lane origin');

  // filenames keep the title verbatim, minus filesystem-hostile characters
  eq(RMExport.fileName(sEx), 'T.png', 'png filename is the exact title');
  eq(RMExport.fileName({ meta: { title: 'My Plan: Q3/Q4?' } }), 'My Plan Q3Q4.png',
    'png filename strips only invalid filename characters');

  // the header is the date row alone — no title / range / phase-span lanes
  eq(RMExport.layout(sEx, {}).rows[0].y, 22, 'rows start right under the date header');
})();

// ------------------------------------------------------------- export grouping + plan
section('export grouping');
(function () {
  var sG = mkState([
    { num: 1, feature: 'Alpha item', workstream: 'Product', epic: 'OS', startDay: 0, durDays: 10 },
    { num: 2, feature: 'Beta item', workstream: 'Data', epic: 'ML', startDay: 20, durDays: 5 },
    { num: 3, feature: 'Gamma item', workstream: 'Product', epic: 'ML', startDay: 5, durDays: 5 },
    { num: 4, feature: 'Next thing', phaseId: 'p2', workstream: '', epic: '', startDay: 40, durDays: 10 }
  ]);
  function names(lay, kind) {
    return lay.rows.filter(function (r) { return r.kind === kind; })
      .map(function (r) { return r.name; });
  }
  function feats(lay) {
    return lay.rows.filter(function (r) { return r.kind === 'item'; })
      .map(function (r) { return r.feature; });
  }

  // the left label rail is gone: bars start at x 0 and labels ride the bars
  var base = RMExport.layout(sG, {});
  eq(base.laneX, 0, 'export lane starts at 0 — no left rail');
  eq(base.width, (base.w1 - base.w0) * base.weekPx, 'width is the date lane alone');

  // group by workstream inside each phase, default/empty workstream last
  var gw = RMExport.layout(sG, { groupWs: true });
  eq(names(gw, 'wsband'), ['Product', 'Data', 'General'],
    'workstream sub-bands per phase; empty workstream shows the default name last');
  eq(feats(gw), ['Alpha item', 'Gamma item', 'Beta item', 'Next thing'],
    'items regroup under their workstream band');
  ok(/^[0-9A-Fa-f]{6}$/.test(gw.rows.filter(function (r) { return r.kind === 'wsband'; })[0].color),
    'workstream bands carry the workstream color');
  eq(gw.rows.filter(function (r) { return r.kind === 'wsband'; }).map(function (r) { return r.count; }),
    [2, 1, 1], 'workstream bands count their items');

  // group by epic (no workstream grouping)
  var ge = RMExport.layout(sG, { groupEpic: true });
  eq(names(ge, 'eband'), ['OS', 'ML', ''], 'epic sub-bands in encounter order; empty epic key kept raw');
  eq(feats(ge), ['Alpha item', 'Beta item', 'Gamma item', 'Next thing'],
    'items regroup under their epic band');

  // both: phase > workstream > epic
  var gb = RMExport.layout(sG, { groupWs: true, groupEpic: true });
  eq(gb.rows.map(function (r) { return r.kind; }),
    ['band', 'wsband', 'eband', 'item', 'eband', 'item', 'wsband', 'eband', 'item',
      'band', 'wsband', 'eband', 'item'],
    'nested grouping emits phase > workstream > epic > items');
  var lastRow = gb.rows[gb.rows.length - 1];
  ok(gb.height >= lastRow.y + lastRow.h, 'grouped canvas height covers the last row');

  // grouping respects workstreamsEnabled, like the timeline
  var sNoWs = mkState([{ num: 1, feature: 'A', workstream: 'Product', startDay: 0, durDays: 5 }]);
  sNoWs.meta.workstreamsEnabled = false;
  eq(names(RMExport.layout(sNoWs, { groupWs: true }), 'wsband'), [],
    'groupWs is a no-op when workstreams are disabled');

  // exact workstream filter ('' = the default workstream) for per-workstream slides
  var wk = RMExport.layout(sG, { wsKey: '' });
  eq(feats(wk), ['Next thing'], 'wsKey "" selects items with no workstream');
  eq(feats(RMExport.layout(sG, { wsKey: 'Product' })), ['Alpha item', 'Gamma item'],
    'wsKey selects exactly one workstream');

  // ---- export plan: how a split turns into multiple images / slides
  var p1 = RMExport.plan(sG, {});
  eq(p1.map(function (e) { return e.name; }), ['T'], 'no split → one export named after the doc');
  var pp = RMExport.plan(sG, { byPhase: true });
  eq(pp.map(function (e) { return e.name; }), ['Alpha', 'Next'], 'phase split → one export per visible phase');
  eq(pp[0].opts.phaseId, 'p1', 'phase split entries carry the phase filter');
  var pw = RMExport.plan(sG, { byWs: true });
  eq(pw.map(function (e) { return e.name; }), ['Product', 'Data', 'General'],
    'workstream split → one export per workstream, default last');
  eq(pw[2].opts.wsKey, '', 'default-workstream entry filters on the empty key');
  var pb = RMExport.plan(sG, { byPhase: true, byWs: true });
  eq(pb.map(function (e) { return e.name; }),
    ['Alpha — Product', 'Alpha — Data', 'Next — General'],
    'phase+workstream split → non-empty combinations only');
  eq(RMExport.plan(sG, { byPhase: true, phaseId: 'p2' }).map(function (e) { return e.name; }),
    ['Next'], 'a phase filter narrows the split');
  var pwFil = RMExport.plan(sG, { byWs: true, ws: 'Data' });
  eq(pwFil.map(function (e) { return e.name; }), ['Data'], 'a workstream filter narrows the split');

  // every slice of one split shares the SAME date window and column grid
  var ppL = RMExport.plan(sG, { byPhase: true }).map(function (e) {
    return RMExport.layout(sG, e.opts);
  });
  eq(ppL[0].d0, ppL[1].d0, 'split slices share the window start');
  eq(ppL[0].d1, ppL[1].d1, 'split slices share the window end');
  eq(ppL[0].width, ppL[1].width, 'split slices share the lane width');
  var full2 = RMExport.layout(sG, {});
  eq(ppL[0].d0, full2.d0, 'the shared window is the unsplit window');
  eq(ppL[0].d1, full2.d1, 'the shared window spans all slices');

  // ---- legend: one swatch per visible workstream, default last
  var lg = RMExport.layout(sG, {});
  eq(lg.legend.map(function (e) { return e.name; }), ['Product', 'Data', 'General'],
    'legend lists the visible workstreams, default last');
  ok(lg.legend.every(function (e) { return /^[0-9A-Fa-f]{6}$/.test(e.color); }),
    'legend entries carry the workstream colors');
  var lgLast = lg.rows[lg.rows.length - 1];
  ok(lg.legend[0].y >= lgLast.y + lgLast.h, 'legend sits below the last row');
  ok(lg.height >= lg.legend[lg.legend.length - 1].y + 16, 'canvas height covers the legend');
  eq(RMExport.layout(sG, { wsKey: 'Data' }).legend.map(function (e) { return e.name; }),
    ['Data'], 'a workstream slice keeps only its own legend entry');
  var sNoWs2 = mkState([{ num: 1, feature: 'A', workstream: 'Product', startDay: 0, durDays: 5 }]);
  sNoWs2.meta.workstreamsEnabled = false;
  eq(RMExport.layout(sNoWs2, {}).legend, [], 'no legend when workstreams are disabled');

// ------------------------------------------------------------- pptx export
section('pptx export');
  var RMPptx = require('../js/export-pptx.js');
  eq(RMPptx.fileName(sG), 'T.pptx', 'pptx filename is the exact title');

  var lay = RMExport.layout(sG, { groupWs: true });
  var slide = RMPptx.slideShapes(lay);
  ok(slide.shapes.length > 0, 'slideShapes emits shapes');
  // everything must land inside a 13.33" × 7.5" slide
  var inBounds = slide.shapes.every(function (s) {
    return s.x >= -0.01 && s.y >= -0.01 && s.x + (s.w || 0) <= 13.34 && s.y + (s.h || 0) <= 7.51;
  });
  ok(inBounds, 'all shapes fit the 16:9 slide');
  var texts = slide.shapes.filter(function (s) { return s.type === 'text'; })
    .map(function (s) { return s.text; });
  ok(texts.indexOf('T') === -1, 'slides carry no title block');
  ok(!texts.some(function (t) { return /^S\d+/.test(t); }), 'the date header shows dates only, no sprint numbers');
  ok(texts.indexOf('Alpha') !== -1, 'phase band names carry no item count');
  ok(texts.some(function (t) { return /Alpha item/.test(t); }), 'item labels are text shapes');
  ok(texts.some(function (t) { return /Product/.test(t); }), 'workstream band names are text shapes');
  var flat = RMPptx.slideShapes(RMExport.layout(sG, {}));
  var flatTexts = flat.shapes.filter(function (s) { return s.type === 'text'; })
    .map(function (s) { return s.text; });
  ok(flatTexts.indexOf('Product') !== -1 && flatTexts.indexOf('General') !== -1,
    'the slide carries a workstream legend');
  var flatRight = Math.max.apply(null, flat.shapes.map(function (s) { return s.x + (s.w || 0); }));
  ok(flatRight > 13.33 - 0.5, 'the timeline always spreads to the slide width');

  // split slides share one scale so rows read the same size on every slide
  var slLays = RMExport.plan(sG, { byPhase: true }).map(function (e) {
    return RMExport.layout(sG, e.opts);
  });
  var shared = RMPptx.slideScale(slLays);
  function bandH(sl) {
    return sl.shapes.filter(function (s) { return s.type === 'rect' && s.color === 'E3DFD5'; })[0].h;
  }
  var shA = RMPptx.slideShapes(slLays[0], shared);
  var shB = RMPptx.slideShapes(slLays[1], shared);
  ok(Math.abs(bandH(shA) - bandH(shB)) < 1e-9,
    'phase band rows are the same height on every slide of a split');
  var bars = slide.shapes.filter(function (s) { return s.type === 'bar'; });
  eq(bars.length, 4, 'each visible item paints one bar shape');
  ok(bars.every(function (b) { return /^[0-9A-Fa-f]{6}$/.test(b.color); }), 'bars carry item colors');
  // a bar wide enough for its label puts the label inside; a narrow one puts it beside
  var wide = RMPptx.labelPlacement('Hi', 200), narrow = RMPptx.labelPlacement('A very long feature label', 30);
  eq(wide.inside, true, 'short label on a wide bar sits inside the bar');
  eq(narrow.inside, false, 'long label on a narrow bar sits beside the bar');

  // milestones render as diamonds, like the live timeline
  var sM = mkState([
    { num: 1, feature: 'Launch', milestone: true, startDay: 10, durDays: 1 },
    { num: 2, feature: 'Work', workstream: 'Product', startDay: 0, durDays: 10 }
  ]);
  var lm = RMExport.layout(sM, {});
  var msRow = lm.rows.filter(function (r) { return r.kind === 'item' && r.feature === 'Launch'; })[0];
  eq(msRow.bar.ms, true, 'milestone rows are flagged in the layout');
  var mShapes = RMPptx.slideShapes(lm).shapes;
  ok(mShapes.some(function (s) { return s.type === 'diamond'; }), 'milestones become diamond shapes');
  eq(mShapes.filter(function (s) { return s.type === 'bar'; }).length, 1,
    'a milestone paints no plain bar');
})();

// ------------------------------------------------------------- excel round-trip
section('excel round-trip');
var ExcelJS = null;
try { ExcelJS = require('exceljs'); } catch (e) { /* not installed here */ }
if (!ExcelJS) {
  skipped++;
  console.log('  (skipped — exceljs not resolvable in NODE_PATH)');
  finish();
} else {
  global.ExcelJS = ExcelJS;
  var RMExcel = require('../js/excel.js');
  var seed = require('./seed.fixture.js');
  var st = RM.normalizeState(seed);
  st.team = [{ id: 't1', name: 'Ada', type: 'Development', rate: 210, cost: 95 }, { id: 't2', name: 'Grace', type: 'Data' }];
  st = RM.normalizeState(st);
  st.items[0].stories = [
    { id: 's1', title: 'story one', done: false, description: '<b>rich</b> body', ac: '<ul><li>crit one</li></ul>' },
    { id: 's2', title: 'story two', done: true, description: '', ac: '' }
  ];
  st.items[0].headcount = 3;
  // a wall of emoji guarantees several chunk boundaries fall inside surrogate pairs
  st.items[1].notes = new Array(20001).join('🚀');
  st.items[1].feature = 'emoji stress 🎯';

  var uiPrefs = { weekPx: 41, view: 'scoping', capType: 'Data Science 🧪', groupEpic: true, expanded: { i1: true } };
  RMExcel.exportWorkbook(st, uiPrefs).then(function (buf) {
    return RMExcel.importWorkbook(buf).then(function (r1) {
      ok(r1.source === 'tool', 'reimport hits the lossless path');
      ok(r1.ui && r1.ui.weekPx === 41 && r1.ui.view === 'scoping' &&
        r1.ui.capType === uiPrefs.capType && r1.ui.groupEpic === true && r1.ui.expanded.i1 === true,
        'UI prefs (zoom/view/grouping/expansion, incl. emoji) ride in the file');
      eq(r1.state.items.length, st.items.length, 'item count survives');
      eq(r1.state.team.length, 2, 'team survives');
      eq(r1.state.items[0].stories.length, 2, 'stories survive');
      eq(r1.state.items[0].stories[0].description, '<b>rich</b> body', 'story rich description survives losslessly');
      eq(r1.state.items[0].stories[0].ac, '<ul><li>crit one</li></ul>', 'story acceptance criteria survive losslessly');
      eq(r1.state.items[0].headcount, 3, 'headcount survives');
      ok(r1.state.items[1].notes === st.items[1].notes, 'emoji notes survive chunk boundaries losslessly');
      var origSched = st.items.filter(function (i2) { return i2.startDay != null; }).length;
      var newSched = r1.state.items.filter(function (i2) { return i2.startDay != null; }).length;
      eq(newSched, origSched, 'schedules survive');
      eq(r1.state.meta.holidays, st.meta.holidays, 'holidays survive');
      eq(r1.state.meta.sprintAnchor, st.meta.sprintAnchor, 'sprint anchor survives');

      // strip the tool sheet -> template parsing path
      var wb2 = new ExcelJS.Workbook();
      return wb2.xlsx.load(buf).then(function () {
        wb2.removeWorksheet(wb2.getWorksheet('_RoadmapTool').id);
        return wb2.xlsx.writeBuffer();
      }).then(function (buf2) {
        return RMExcel.importWorkbook(buf2);
      }).then(function (r2) {
        ok(r2.source === 'template', 'without tool sheet, template parse engages');
        eq(r2.state.items.length, st.items.length, 'template parse finds every item row');
        eq(r2.state.phases.length, st.phases.length, 'template parse finds every phase band');
        var sc2 = r2.state.items.filter(function (i2) { return i2.startDay != null; }).length;
        ok(Math.abs(sc2 - origSched) <= 2, 'template parse recovers bars (±sprint rounding), got ' + sc2 + ' vs ' + origSched);
        eq(r2.state.meta.numWeeks, st.meta.numWeeks, 'weekly columns keep the exact week count');
        eq(r2.state.meta.weeksPerSprint, 2, 'column granularity does not redefine the sprint length');
        var wkOk = st.items.every(function (o) {
          if (o.startDay == null) return true;
          var m2 = r2.state.items.filter(function (i2) { return i2.num === o.num; })[0];
          return m2 && m2.startDay === Math.floor(o.startDay / 5) * 5;
        });
        ok(wkOk, 'bar starts re-import at week precision (not rounded to sprints)');
        var it2 = r2.state.items.filter(function (i2) { return i2.num === 11; })[0];
        ok(it2 && it2.deps.indexOf(1) !== -1 && it2.deps.indexOf(2) !== -1, 'deps re-parsed from cell text');
        var withStories = r2.state.items.filter(function (i2) { return i2.stories.length === 2; });
        ok(withStories.length === 1, 'stories re-attached via Stories sheet');
        eq(withStories[0].stories[0].description, 'rich body', 'template path keeps story description as text');
        eq(withStories[0].stories[0].ac, 'crit one', 'template path keeps acceptance criteria as text');
        eq(r2.state.team.length, 2, 'team re-parsed from Team sheet');
        eq(r2.state.team[0].rate, 210, 'role rate survives the template path');
        eq(r2.state.team[0].cost, 95, 'role cost survives the template path');

        // Excel-style edits to the visible sheets must win over hidden state
        var wb3 = new ExcelJS.Workbook();
        var target = st.items[0].num;
        return wb3.xlsx.load(buf).then(function () {
          var rws = wb3.getWorksheet('Roadmap');
          for (var rr = 4; rr <= rws.rowCount; rr++) {
            if (String(rws.getRow(rr).getCell(1).value) === String(target)) {
              rws.getRow(rr).getCell(4).value = 'Renamed in Excel';
              rws.getRow(rr).getCell(7).value = 'note from excel';
              break;
            }
          }
          var stws = wb3.getWorksheet('Stories');
          for (var sr = 2; sr <= stws.rowCount; sr++) {
            if (String(stws.getRow(sr).getCell(1).value) === String(target)) {
              stws.getRow(sr).getCell(3).value = 'Story renamed in Excel';
              stws.getRow(sr).getCell(4).value = 'Yes';
              break;
            }
          }
          return wb3.xlsx.writeBuffer();
        }).then(function (buf3) {
          return RMExcel.importWorkbook(buf3);
        }).then(function (r3) {
          ok(r3.source === 'tool', 'excel-edited file still loads via the lossless path');
          var e1 = r3.state.items.filter(function (i2) { return i2.num === target; })[0];
          eq(e1.feature, 'Renamed in Excel', 'visible Feature edit wins over hidden state');
          eq(e1.notes, 'note from excel', 'visible Notes edit wins over hidden state');
          eq(e1.stories[0].title, 'Story renamed in Excel', 'visible story rename wins');
          ok(e1.stories[0].done === true, 'story Done toggled from the sheet');
          var un = r3.state.items.filter(function (i2) { return i2.num === st.items[1].num; })[0];
          ok(un.feature === st.items[1].feature, 'untouched rows keep their hidden-state values');

          // sync-client echo detection: a container rewrite (re-zip of the
          // same workbook — what OneDrive does after upload) changes the
          // bytes but NOT the embedded document JSON. The desktop shell
          // compares readStateJson against the stateJsonOf it last wrote to
          // swallow echoes of its own save instead of announcing a reload.
          var jsonOut = RMExcel.stateJsonOf(st);
          return RMExcel.readStateJson(buf).then(function (jsonBack) {
            ok(typeof jsonBack === 'string' && jsonBack === jsonOut,
              'readStateJson returns the exact string stateJsonOf embedded');
            var wb4 = new ExcelJS.Workbook();
            return wb4.xlsx.load(buf).then(function () {
              return wb4.xlsx.writeBuffer();
            });
          }).then(function (buf4) {
            return RMExcel.readStateJson(buf4);
          }).then(function (json4) {
            ok(json4 === jsonOut, 'embedded JSON survives a container rewrite unchanged');
            var wb5 = new ExcelJS.Workbook();
            wb5.addWorksheet('Sheet1').getCell('A1').value = 'not a headway file';
            return wb5.xlsx.writeBuffer();
          }).then(function (buf5) {
            return RMExcel.readStateJson(buf5);
          }).then(function (json5) {
            ok(json5 === null, 'foreign workbooks read as null (no echo match possible)');
            // ---- tags round trip, on its own document (the emoji fixture
            // above is byte-sensitive, so tags get their own workbook)
            var stT = RM.normalizeState({
              meta: JSON.parse(JSON.stringify(META)),
              phases: [{ id: 'p1', name: 'Alpha', bucket: false }],
              items: [{ id: 'i1', num: 1, phaseId: 'p1', feature: 'Tagged', tags: ['tech debt', 'q3'],
                stories: [{ id: 's1', title: 'story one', num: 2, tags: ['spike'] }, { id: 's2', title: 'story two', num: 40, deps: [2] }] }],
              team: []
            });
            return RMExcel.exportWorkbook(stT).then(function (bufT) {
              return RMExcel.importWorkbook(bufT).then(function (rt1) {
                eq(rt1.state.items[0].tags, ['tech debt', 'q3'], 'feature tags survive the lossless path');
                eq(rt1.state.items[0].stories[0].tags, ['spike'], 'story tags survive the lossless path');
                var wbT = new ExcelJS.Workbook();
                return wbT.xlsx.load(bufT).then(function () {
                  wbT.removeWorksheet(wbT.getWorksheet('_RoadmapTool').id);
                  return wbT.xlsx.writeBuffer();
                }).then(function (bufT2) {
                  return RMExcel.importWorkbook(bufT2);
                }).then(function (rt2) {
                  ok(rt2.source === 'template', 'tags doc without the tool sheet parses as a template');
                  eq(rt2.state.items[0].tags, ['tech debt', 'q3'], 'template path re-reads the Roadmap Tags column');
                  eq(rt2.state.items[0].stories[0].tags, ['spike'], 'template path re-reads the Stories Tags column');
                  eq(rt2.state.items[0].stories.map(function (x) { return x.num; }),
                    stT.items[0].stories.map(function (x) { return x.num; }), 'template path re-reads the Stories # column');
                  eq(rt2.state.items[0].stories[1].deps, [stT.items[0].stories[0].num], 'template path re-reads the Stories Depends on column');
                  // an older workbook has no Tags columns at all
                  var wbT3 = new ExcelJS.Workbook();
                  return wbT3.xlsx.load(bufT).then(function () {
                    wbT3.removeWorksheet(wbT3.getWorksheet('_RoadmapTool').id);
                    var rws3 = wbT3.getWorksheet('Roadmap');
                    rws3.spliceColumns(rws3.columnCount, 1);
                    var sws3 = wbT3.getWorksheet('Stories');
                    sws3.spliceColumns(7, 3);
                    return wbT3.xlsx.writeBuffer();
                  }).then(function (bufT3) {
                    return RMExcel.importWorkbook(bufT3);
                  }).then(function (rt3) {
                    eq(rt3.state.items[0].tags, [], 'a pre-tags workbook still imports, with no tags');
                    eq(rt3.state.items[0].stories[0].tags, [], 'and its stories carry no tags');
                    eq(rt3.state.items[0].stories[1].deps, [], 'a workbook without the Depends on column imports with no story deps');
                    ok(rt3.state.items[0].stories[0].num > 0, 'and normalize assigns fresh story numbers');

                    // ---- capacity fields round trip (lossless path + the
                    // visible Stories/Team columns on the template path)
                    var sX = mkState([{ num: 1, feature: 'f', stories: [{ title: 'a', capType: 'Design', capMult: 2 }] }],
                      { team: [{ name: 'P', capType: 'Design', points: 7 }] });
                    sX.meta.capMode = 'points';
                    return RMExcel.exportWorkbook(sX).then(function (bufX) {
                      return RMExcel.importWorkbook(bufX).then(function (rx1) {
                        eq(rx1.state.items[0].stories[0].capMult, 2, 'story multiplier survives the round trip');
                        eq(rx1.state.team[0].points, 7, 'member points survive the round trip');
                        eq(rx1.state.meta.capMode, 'points', 'capMode survives the round trip');
                        var cellStr = function (c) { return c && c.value != null ? String(c.value) : ''; };
                        var wbX = new ExcelJS.Workbook();
                        return wbX.xlsx.load(bufX).then(function () {
                          var swsX = wbX.getWorksheet('Stories');
                          eq(cellStr(swsX.getCell(1, 10)), 'Capacity type', 'Stories sheet shows a Capacity type column');
                          eq(cellStr(swsX.getCell(1, 11)), 'Multiplier', 'Stories sheet shows a Multiplier column');
                          eq(cellStr(swsX.getCell(2, 10)), 'Design', 'the story capacity type is written out');
                          eq(cellStr(swsX.getCell(2, 11)), '2', 'the story multiplier is written out when it is not 1');
                          var twsX = wbX.getWorksheet('Team');
                          eq(cellStr(twsX.getCell(1, 9)), 'Points per sprint', 'Team sheet shows a Points per sprint column');
                          eq(cellStr(twsX.getCell(2, 9)), '7', 'the member points are written out');
                          wbX.removeWorksheet(wbX.getWorksheet('_RoadmapTool').id);
                          return wbX.xlsx.writeBuffer();
                        }).then(function (bufX2) {
                          return RMExcel.importWorkbook(bufX2);
                        }).then(function (rx2) {
                          ok(rx2.source === 'template', 'the capacity doc without the tool sheet parses as a template');
                          eq(rx2.state.team[0].points, 7, 'template path re-reads Points per sprint');
                        });
                      });
                    }).then(function () {
                      // a member with no points cell inherits meta.defaultPoints
                      var sY = mkState([{ num: 1, feature: 'f' }], { team: [{ name: 'Q' }] });
                      return RMExcel.exportWorkbook(sY).then(function (bufY) {
                        var wbY = new ExcelJS.Workbook();
                        return wbY.xlsx.load(bufY).then(function () {
                          wbY.removeWorksheet(wbY.getWorksheet('_RoadmapTool').id);
                          return wbY.xlsx.writeBuffer();
                        });
                      }).then(function (bufY2) {
                        return RMExcel.importWorkbook(bufY2);
                      }).then(function (ry) {
                        ok(ry.state.team[0].points === null, 'a blank Points per sprint cell reads back as null (inherit)');
                        finish();
                      });
                    });
                  });
                });
              });
            });
          });
        });
      });
    });
  }).catch(function (err) {
    failed++;
    console.error('  ✗ excel round-trip threw: ' + (err && err.stack || err));
    finish();
  });
}

function finish() {
section('apps switch');
{
  var ap = RM.normalizeState({ meta: { title: 'A', timelineStart: '2026-07-27', numWeeks: 8 }, phases: [], items: [] });
  ok(RM.APPS.every(function (a) { return a[0] === 'sprints' ? RM.appEnabled(ap, 'sprints') : ap.meta.apps[a[0]] === true; }), 'a fresh document has every app on (Sprinting via sprints on)');
  var ap2 = RM.normalizeState({ meta: { title: 'A', timelineStart: '2026-07-27', numWeeks: 8, apps: { scoping: false, planning: false, bogus: false } }, phases: [], items: [] });
  eq([ap2.meta.apps.scoping, ap2.meta.apps.planning, ap2.meta.apps.prio, 'bogus' in ap2.meta.apps], [false, true, true, false],
    'off flags stick, Planning is forced on, unknown keys drop');
  eq([RM.appEnabled(ap2, 'scoping'), RM.appEnabled(ap2, 'planning'), RM.appEnabled(ap2, 'setup'), RM.appEnabled(ap2, 'history')], [false, true, true, true],
    'appEnabled: off app, forced Planning, and non-apps always reachable');
}

section('item types & hierarchy');
{
  var sT = mkState([
    { num: 1, feature: 'A', epic: 'E1', stories: [{ id: 'sa', title: 'x' }, { id: 'sb', title: 'y', type: 'bug' }] },
    { num: 2, feature: 'B', type: 'bug' },
    { num: 3, feature: 'C', type: 'nope' }
  ]);
  eq(sT.meta.itemTypes.map(function (t) { return t.key; }), ['epic', 'feature', 'bug', 'task', 'story', 'subtask'], 'default types seeded');
  eq(sT.meta.hierarchy.levels.map(function (l) { return l.key; }), ['epic', 'feature', 'story'], 'three fixed levels');
  eq(sT.meta.hierarchy.levels[1].types, ['feature', 'bug', 'task'], 'feature level default types');
  eq(sT.meta.hierarchy.anyTypeAnyLevel, false, 'switch off by default');
  eq(sT.items[0].type, 'feature', 'missing item type -> level default');
  eq(sT.items[1].type, 'bug', 'known item type kept');
  eq(sT.items[2].type, 'feature', 'unknown item type -> level default');
  eq(sT.items[0].stories[0].type, 'story', 'missing story type -> level default');
  eq(sT.items[0].stories[1].type, 'bug', 'story type kept');
  eq(sT.epicTypes, {}, 'epicTypes seeded empty');
  eq(RM.typeOf(sT, 'E1', 'epic').key, 'epic', 'epic without a stored type resolves to epic');
  eq(RM.levelLabel(sT, 'feature'), 'Feature', 'level label');
  eq(RM.levelLabel(sT, 'story', true), 'Stories', 'plural label');
  eq(RM.typesFor(sT, 'story').map(function (t) { return t.key; }), ['story', 'subtask', 'bug'], 'allowed types at story level');
  eq(RM.defaultTypeFor(sT, 'epic'), 'epic', 'default type for epic level');
  eq(RM.jiraTypeName(sT, 'story'), 'Sub-task', 'story type maps to Sub-task by default');
  eq(RM.typeOf(sT, sT.items[1], 'feature').icon, 'bug', 'typeOf returns the record');

  // a disallowed stored type survives normalize
  var sT2 = mkState([{ num: 1, feature: 'A', type: 'subtask' }]);
  eq(sT2.items[0].type, 'subtask', 'disallowed type kept on normalize');
  // the switch opens every type at every level
  sT2.meta.hierarchy.anyTypeAnyLevel = true;
  eq(RM.typesFor(sT2, 'epic').length, 6, 'any type any level lists all types');

  // legacy Jira names migrate into the type records once
  var sT3 = mkState([{ num: 1, feature: 'A' }], { meta: Object.assign(JSON.parse(JSON.stringify(META)), { jira: { epicType: 'Initiative', featureType: 'Task', storyType: 'Subtask' } }) });
  eq(RM.jiraTypeName(sT3, 'epic'), 'Initiative', 'legacy epicType migrates');
  eq(RM.jiraTypeName(sT3, 'feature'), 'Task', 'legacy featureType migrates');
  eq(RM.jiraTypeName(sT3, 'story'), 'Subtask', 'legacy storyType migrates');
  sT3.meta.jira.featureType = 'Bug';
  eq(RM.jiraTypeName(RM.normalizeState(sT3), 'feature'), 'Task', 'legacy names are read only when itemTypes is absent');

  // custom labels and lists round-trip; empty level list falls back
  var sT4 = mkState([{ num: 1, feature: 'A' }], { meta: Object.assign(JSON.parse(JSON.stringify(META)), {
    itemTypes: [{ key: 'epic', label: 'Theme', icon: 'layers', jira: 'Epic' }, { key: 'feature', label: 'Feature', icon: 'rows-3', jira: 'Story' }, { key: 'story', label: 'Story', icon: 'list-tree', jira: 'Sub-task' }],
    hierarchy: { levels: [{ key: 'feature', label: 'Capability', types: ['feature', 'ghost'] }, { key: 'story', label: 'Task', types: [] }], anyTypeAnyLevel: true } }) });
  eq(sT4.meta.hierarchy.levels.map(function (l) { return l.label; }), ['Epic', 'Capability', 'Task'], 'missing level gets default label, order fixed');
  eq(sT4.meta.hierarchy.levels[1].types, ['feature'], 'unknown type keys are dropped from a level');
  eq(sT4.meta.hierarchy.levels[2].types, ['story'], 'empty level list falls back to defaults filtered to existing types');
  eq(sT4.meta.hierarchy.anyTypeAnyLevel, true, 'switch round-trips');
  eq(RM.levelLabel(sT4, 'story', true), 'Tasks', 'plural of a custom label');
}
{
  var sC = RM.normalizeState({ meta: { title: 'C', timelineStart: '2026-07-27', numWeeks: 8 }, phases: [{ id: 'p', name: 'P' }], team: [], items: [] });
  var tF = RM.itemType(sC, 'feature'), tS = RM.itemType(sC, 'story');
  eq(tF.icon, 'square', 'feature default icon is the filled square glyph');
  eq(tS.icon, 'bookmark', 'story default icon is the bookmark');
  eq(RM.itemType(sC, 'task').icon, 'check-square', 'task icon unchanged');
  ok(RM.itemTypes(sC).every(function (t) { return /^[0-9A-F]{6}$/.test(t.color); }), 'every type carries a resolved 6-hex color');
  eq(RM.itemTypes(sC).map(function (t) { return t.color; }), RM.HASH_PALETTE.slice(0, RM.itemTypes(sC).length), 'default colors follow the hash palette in list order');
  eq(RM.colorForType(sC, 'bug'), RM.HASH_PALETTE[2], 'colorForType reads the record');
  eq(RM.colorForType(sC, 'nope'), RM.PALETTE.neutral, 'unknown type is neutral');
  RM.setItemTypeColor(sC, 'bug', '#ff0000');
  eq(RM.itemType(sC, 'bug').color, 'FF0000', 'setItemTypeColor stores an upper-case hex without #');
  RM.setItemTypeColor(sC, 'bug', 'not a color');
  eq(RM.itemType(sC, 'bug').color, 'FF0000', 'a bad color is ignored');
  var kNew = RM.addItemType(sC, 'Spike', 'zap', 'Spike');
  ok(/^[0-9A-F]{6}$/.test(RM.itemType(sC, kNew).color), 'a new type gets a palette color at once');
  // stored documents on the old default icons migrate; custom icons stay
  var sM = RM.normalizeState({ meta: { title: 'M', timelineStart: '2026-07-27', numWeeks: 8,
    itemTypes: [{ key: 'feature', label: 'Feature', icon: 'rows-3', jira: 'Story' }, { key: 'story', label: 'Story', icon: 'list-tree', jira: 'Sub-task' }, { key: 'bug', label: 'Bug', icon: 'flame', jira: 'Bug', color: '123456' }] },
    phases: [{ id: 'p', name: 'P' }], team: [], items: [] });
  eq(RM.itemType(sM, 'feature').icon, 'square', 'old feature default icon migrates to square');
  eq(RM.itemType(sM, 'story').icon, 'bookmark', 'old story default icon migrates to bookmark');
  eq(RM.itemType(sM, 'bug').icon, 'flame', 'a custom icon survives');
  eq(RM.itemType(sM, 'bug').color, '123456', 'a stored color survives');
  // color mode 'type'
  ok(RM.COLOR_MODES.indexOf('type') !== -1, "'type' is a color mode");
  var sT3col = RM.normalizeState({ meta: { title: 'T', timelineStart: '2026-07-27', numWeeks: 8 }, phases: [{ id: 'p', name: 'P' }], team: [],
    items: [{ id: 'x', num: 1, phaseId: 'p', feature: 'X', type: 'bug' }, { id: 'y', num: 2, phaseId: 'p', feature: 'Y' }] });
  RM.setColorMode('type');
  eq(RM.colorForItem(sT3col, sT3col.items[0]), RM.colorForType(sT3col, 'bug'), 'type mode colors an item by its type');
  eq(RM.colorForItem(sT3col, sT3col.items[1]), RM.colorForType(sT3col, 'feature'), 'an item without a type takes the level default type color');
  eq(RM.colorLegend(sT3col, sT3col.items).map(function (e) { return e.name; }), ['Bug', 'Feature'], 'type legend names the types in first-seen order');
  RM.setColorMode('workstream');
}

section('item type mutations & validation');
{
  var sM = mkState([
    { num: 1, feature: 'A', type: 'bug', epic: 'E', stories: [{ id: 's1', title: 'x', type: 'bug' }] }
  ]);
  var nk = RM.addItemType(sM, 'Spike', 'zap', 'Spike');
  eq(nk, 'spike', 'addItemType slugs the label into a key');
  eq(RM.addItemType(sM, 'Spike', 'zap', 'Spike'), 'spike-2', 'duplicate labels get a suffixed key');
  ok(RM.setTypeAllowed(sM, 'feature', 'spike', true), 'allow a type at a level');
  eq(RM.levelOf(sM, 'feature').types.slice(-1)[0], 'spike', 'allowed list grows');
  ok(!RM.setTypeAllowed(sM, 'epic', 'epic', false), 'cannot remove the last type of a level');
  ok(RM.setTypeAllowed(sM, 'feature', 'spike', false), 'disallow again');
  RM.renameItemType(sM, 'spike', 'Research');
  eq(RM.itemType(sM, 'spike').label, 'Research', 'rename keeps the key');
  RM.setItemTypeJira(sM, 'spike', 'Research task');
  eq(RM.jiraTypeName(sM, 'spike'), 'Research task', 'jira name edit');
  RM.setItemTypeIcon(sM, 'spike', 'flask-conical');
  eq(RM.itemType(sM, 'spike').icon, 'flask-conical', 'icon edit');
  RM.setLevelLabel(sM, 'story', 'Task');
  eq(RM.levelLabel(sM, 'story'), 'Task', 'level label edit');
  ok(!RM.removeItemType(sM, 'epic'), 'cannot remove the only type of a level');
  ok(RM.removeItemType(sM, 'bug'), 'remove a type');
  eq(sM.items[0].type, 'feature', 'items of the removed type fall back to the level default');
  eq(sM.items[0].stories[0].type, 'story', 'stories too');
  eq(RM.levelOf(sM, 'story').types, ['story', 'subtask'], 'removed key leaves every level list');
  ok(!RM.itemType(sM, 'bug'), 'record gone');

  var sV2 = mkState([{ num: 1, feature: 'A', type: 'subtask', epic: 'E', stories: [{ id: 's1', title: 'x', type: 'task' }] }], { epicTypes: { E: 'feature' } });
  var vv = RM.validate(sV2);
  ok((vv.byItem[sV2.items[0].id] || []).some(function (f) { return f.code === 'TYPE_LEVEL' && /Subtask/.test(f.msg); }), 'item with a disallowed type warns');
  ok(vv.global.some(function (f) { return f.code === 'TYPE_LEVEL' && /story/i.test(f.msg) && /Task/.test(f.msg); }), 'story with a disallowed type warns globally');
  ok(vv.global.some(function (f) { return f.code === 'TYPE_LEVEL' && /Epic/.test(f.msg) && /Feature/.test(f.msg); }), 'epic with a disallowed type warns globally');
  RM.setAnyTypeAnyLevel(sV2, true);
  var vv2 = RM.validate(sV2);
  ok(!(vv2.byItem[sV2.items[0].id] || []).some(function (f) { return f.code === 'TYPE_LEVEL'; }) && !vv2.global.some(function (f) { return f.code === 'TYPE_LEVEL'; }), 'switch on silences TYPE_LEVEL');

  // TYPE_LEVEL must resolve missing/raw type fields (e.g. items pushed without
  // a `type` after commit(), before the next normalizeState) via RM.typeOf,
  // not compare the raw field directly.
  var sV3 = mkState([{ num: 1, feature: 'A', epic: 'E', stories: [{ id: 's1', title: 'x' }] }], { epicTypes: { E: 'epic' } });
  delete sV3.items[0].type;
  delete sV3.items[0].stories[0].type;
  var vv3 = RM.validate(sV3);
  ok(!(vv3.byItem[sV3.items[0].id] || []).some(function (f) { return f.code === 'TYPE_LEVEL'; }), 'item pushed with no type resolves via RM.typeOf and warns nothing');
  ok(!vv3.global.some(function (f) { return f.code === 'TYPE_LEVEL'; }), 'story pushed with no type resolves via RM.typeOf and warns nothing');
}

section('flags');
{
  var sF = mkState([{ id: 'f', num: 1, phaseId: 'p1', feature: 'F', flag: 'needs legal review',
    stories: [{ id: 'a', title: 'A', flag: true }, { id: 'b', title: 'B', flag: { reason: '  late  ' } }, { id: 'c', title: 'C', flag: false }] }]);
  eq(sF.items[0].flag, { reason: 'needs legal review' }, 'a string flag becomes { reason }');
  eq(sF.items[0].stories[0].flag, { reason: '' }, 'true flags with an empty reason');
  eq(sF.items[0].stories[1].flag, { reason: 'late' }, 'an object flag keeps a trimmed reason');
  eq(sF.items[0].stories[2].flag, null, 'false / missing means not flagged');
  eq(RM.normalizeFlag(undefined), null, 'normalizeFlag: nothing → null');
}
section('story numbers');
{
  var sN = mkState([
    { id: 'f1', num: 1, phaseId: 'p1', feature: 'One', stories: [{ id: 'a', title: 'A' }, { id: 'b', title: 'B', num: 7 }] },
    { id: 'f2', num: 3, phaseId: 'p1', feature: 'Two', stories: [{ id: 'c', title: 'C', num: 3 }] }
  ]);
  eq(sN.items[0].stories.map(function (s) { return s.num; }), [8, 7], 'a numbered story keeps its number; a blank one gets the next past everything in use');
  eq(sN.items[1].stories[0].num, 9, 'a story number that collides with a feature is reassigned');
  eq(RM.nextNum(sN), 10, 'nextNum spans features and stories');
  // a blank story must never take a number another story holds (that would
  // silently re-point every dep naming it)
  var sKeep = mkState([{ id: 'f', num: 1, phaseId: 'p1', feature: 'F', stories: [{ id: 'a', title: 'A' }, { id: 'b', title: 'B', num: 2 }, { id: 'c', title: 'C', num: 3, deps: [2] }] }]);
  eq(sKeep.items[0].stories.map(function (x) { return x.id + ':' + x.num; }), ['a:4', 'b:2', 'c:3'], 'blank story is numbered past the held ones');
  eq(RM.resolveStoryDeps(sKeep, sKeep.items[0].stories[2]).deps[0].st.id, 'b', 'the dep still points at the story it named');
  var sDup = RM.normalizeState(mkState([{ id: 'f', num: 1, phaseId: 'p1', feature: 'F', stories: [{ id: 'a', title: 'A', num: 5 }] }, { id: 'g', num: 2, phaseId: 'p1', feature: 'G', stories: [] }]));
  sDup.items[1].num = 5; // forced collision, as the targeted apply could leave transiently
  ok(RM.validate(sDup).global.some(function (g) { return g.code === 'DUP_NUM' && g.storyId === 'a'; }) &&
    (RM.validate(sDup).byItem['g'] || []).some(function (x) { return x.code === 'DUP_NUM'; }), 'DUP_NUM covers a feature/story collision on both sides');
  eq(RM.storyByNum(sN, 7).st.id, 'b', 'storyByNum finds a story');
  eq(RM.storyByNum(sN, 1), null, 'a feature number is not a story');
  eq(RM.byNum(sN, 1).kind, 'feature', 'byNum: feature');
  eq(RM.byNum(sN, 9).kind + ':' + RM.byNum(sN, 9).st.id, 'story:c', 'byNum: story');
  eq(RM.byNum(sN, 99), null, 'byNum: nothing');
  eq(RM.renumberItem(sN, 'f2', 7), 10, 'renumbering a feature onto a story number falls back to the next free one');
  sN.items[1].stories[0].deps = [7];
  eq(RM.renumberStory(sN, 'f1', 'b', 20), 20, 'renumberStory takes a free number');
  eq(sN.items[1].stories[0].deps, [20], 'story deps follow the renumbered story');
  eq(RM.renumberStory(sN, 'f1', 'a', 20), 21, 'a taken number falls back to nextNum');
}
section('story deps');
{
  var sSD = mkState([
    { id: 'f1', num: 1, phaseId: 'p1', feature: 'One', startDay: 0, durDays: 5,
      stories: [{ id: 'a', title: 'A', num: 10 }, { id: 'b', title: 'B', num: 11, deps: [10, '10', 11, 99, 1] }] },
    { id: 'f2', num: 2, phaseId: 'p1', feature: 'Two', startDay: 10, durDays: 5,
      stories: [{ id: 'c', title: 'C', num: 12, deps: [11], startDay: 2, durDays: 3 }] }
  ]);
  eq(sSD.items[0].stories[1].deps, [10, 99, 1], 'normalize keeps numeric unique deps, drops self, keeps unknown');
  var rb = RM.resolveStoryDeps(sSD, sSD.items[0].stories[1]);
  eq(rb.deps.map(function (r) { return r.st.id; }), ['a'], 'resolveStoryDeps returns refs');
  eq(rb.unknown, [99, 1], 'unknown numbers — including a feature number — are reported');
  eq(RM.storyLabel(sSD, RM.storyRef(sSD, 'c')), '#12 · C', 'label is the story number and title');
  eq(RM.storyWindow(sSD, sSD.items[0], sSD.items[0].stories[0]), { startDay: 0, endDay: 5 }, 'a story without a timeline takes the feature bar');
  eq(RM.storyWindow(sSD, sSD.items[1], sSD.items[1].stories[0]), { startDay: 2, endDay: 5 }, 'a story with a timeline uses it');
  eq(RM.storyDepEdges(sSD).map(function (e) { return e[0].st.id + '>' + e[1].st.id; }), ['a>b', 'b>c'], 'every explicit edge, dep first');
  eq(RM.storyDependents(sSD, sSD.items[0].stories[1]).map(function (r) { return r.st.id; }), ['c'], 'dependents list this story');
  var vSD = RM.validate(sSD);
  ok(vSD.global.some(function (g) { return g.code === 'STORY_DEP_ORDER' && g.storyId === 'c' && /#12 "C" starts before #11 "B"/.test(g.msg); }), 'STORY_DEP_ORDER names both ends by number');
  ok(vSD.global.filter(function (g) { return g.code === 'STORY_UNKNOWN_DEP' && g.storyId === 'b'; }).length === 2, 'each unknown dep warns');
  ok(vSD.global.some(function (g) { return g.code === 'STORY_UNKNOWN_DEP' && /#1 is a feature/.test(g.msg); }), 'a feature number says so');
  var sCyc = mkState([{ id: 'f', num: 1, phaseId: 'p1', feature: 'F', stories: [{ id: 'x', title: 'X', num: 2, deps: [3] }, { id: 'y', title: 'Y', num: 3, deps: [2] }, { id: 'z', title: 'Z', num: 4, deps: [2] }] }]);
  eq(RM.storyCycleMembers(sCyc), { x: true, y: true }, 'cycle members (z hangs off the cycle, not in it)');
  ok(RM.validate(sCyc).global.filter(function (g) { return g.code === 'STORY_CYCLE'; }).length === 2, 'both cycle members get STORY_CYCLE');
  var sUn = mkState([{ id: 'f', num: 1, phaseId: 'p1', feature: 'F', stories: [{ id: 'p', title: 'P', num: 5 }] },
    { id: 'g', num: 2, phaseId: 'p1', feature: 'G', startDay: 0, durDays: 5, stories: [{ id: 'q', title: 'Q', num: 6, deps: [5], startDay: 3, durDays: 2 }] }]);
  ok(RM.validate(sUn).global.some(function (g) { return g.code === 'STORY_DEP_UNSCHEDULED' && g.storyId === 'q'; }), 'a scheduled story depending on an unscheduled one gets the info');
  var sDone = mkState([{ id: 'f', num: 1, phaseId: 'p1', feature: 'F', startDay: 0, durDays: 5, stories: [{ id: 'd', title: 'D', num: 2, done: true }, { id: 'e', title: 'E', num: 3, deps: [2], startDay: 1, durDays: 1 }] }]);
  ok(!RM.validate(sDone).global.some(function (g) { return g.code === 'STORY_DEP_ORDER'; }), 'a done dependency never violates order');
  var mapped = [{ id: 'n1', title: 'A', num: 30, deps: [21, 5] }, { id: 'n2', title: 'B', num: 31, deps: [] }];
  RM.remapStoryDeps(mapped, { 20: 30, 21: 31 });
  eq(mapped[0].deps, [31, 5], 'remapStoryDeps rewrites numbers inside the copy and keeps outside ones');
}
  console.log('\n' + passed + ' passed, ' + failed + ' failed' + (skipped ? ', ' + skipped + ' skipped' : ''));
  process.exit(failed ? 1 : 0);
}

// ------------------------------------------------------------- risk schemes
section('risk schemes');
var sRs = mkState([{ num: 1, feature: 'a' }]);
eq(RM.riskSchemeOf(sRs), 'none', 'new documents track no assessment column');
ok(!RM.riskEnabled(sRs), 'riskEnabled false when none');
var sRs2 = mkState([{ num: 1, feature: 'a', risk: 'H' }]);
eq(RM.riskSchemeOf(sRs2), 'risk', 'legacy docs that used Risk keep the risk scheme');
eq(sRs2.items[0].risk, 'H', 'risk value survives');
RM.setRiskScheme(sRs2, 'confidence');
eq(RM.riskColLabel(sRs2), 'Confidence', 'confidence relabels the column');

// priority is its own column now
var sPr = mkState([{ num: 1, feature: 'a' }]);
eq(RM.prioritySchemeOf(sPr), 'none', 'new docs have no priority column');
RM.setPriorityScheme(sPr, 'moscow');
eq(RM.priorityOrderOf(sPr).join(''), 'MSCW', 'MoSCoW priority options');
RM.setPriorityScheme(sPr, 'levels');
eq(RM.priorityOrderOf(sPr).join(''), 'CHML', 'Critical/High/Medium/Low options');
// every value maps to a color tier: critical red, high orange, medium green, low gray
eq(['C', 'H', 'M', 'L'].map(function (v) { return RM.priorityTier(sPr, v); }).join(' '), 'crit high med low', 'levels ladder tiers');
eq(RM.priorityTier(sPr, null), null, 'no value has no tier');
eq(RM.priorityTier(sPr, 'W'), null, 'a value outside the scheme has no tier');
RM.setPriorityScheme(sPr, 'moscow');
eq(['M', 'S', 'C', 'W'].map(function (v) { return RM.priorityTier(sPr, v); }).join(' '), 'crit high med low', 'MoSCoW ladder tiers');
sPr.items[0].priority = 'M';
eq(RM.colorForPriority(sPr, sPr.items[0]), RM.PRIORITY_RAMP[0], 'Must paints the hottest ramp color');
eq(RM.PRIORITY_RAMP.length, 4, 'four ramp colors, one per tier');
// a doc saved when MoSCoW lived under Risk migrates its values across
var sPrMig = RM.normalizeState({
  meta: { timelineStart: '2026-07-27', numWeeks: 8, riskScheme: 'moscow', holidaysV2026: true },
  phases: [{ id: 'p' }], items: [{ num: 1, feature: 'a', risk: 'W' }]
});
eq(sPrMig.meta.priorityScheme, 'moscow', 'moscow risk scheme migrates to the Priority column');
eq(sPrMig.items[0].priority, 'W', 'the value moves to item.priority');
eq(sPrMig.items[0].risk, null, 'risk is cleared');
ok(RM.SCOPE_FIXED_KEYS.indexOf('priority') !== -1, 'priority is a fixed scoping column');
var sRs3 = RM.normalizeState({
  meta: { timelineStart: '2026-07-27', numWeeks: 8, riskScheme: 'confidence', holidaysV2026: true },
  phases: [{ id: 'p' }], items: [{ num: 1, feature: 'a', risk: 'H' }, { num: 2, feature: 'b', risk: 'W' }]
});
eq(sRs3.items[0].risk, 'H', 'confidence keeps H');
eq(sRs3.items[1].risk, null, 'a value outside the scheme is dropped');
eq(RM.riskColLabel(sRs3), 'Confidence', 'confidence label');

// ------------------------------------------------------------- default workstream
section('default workstream');
var sDw = mkState([{ num: 1, feature: 'a' }]);
eq(RM.defaultWsName(sDw), 'General', 'default workstream is General');
eq(RM.colorForWs(sDw, ''), RM.PALETTE.neutral, 'null workstream paints the default gray');
sDw.meta.defaultWsName = 'Core';
sDw.meta.defaultWsColor = '08875B';
eq(RM.defaultWsName(sDw), 'Core', 'default workstream renames');
eq(RM.colorForWs(sDw, ''), '08875B', 'default workstream color follows the setting');
var sDwN = RM.normalizeState({ meta: { timelineStart: '2026-07-27', numWeeks: 8, defaultWsColor: 'nope', holidaysV2026: true }, phases: [{ id: 'p' }], items: [] });
eq(sDwN.meta.defaultWsColor, RM.PALETTE.neutral, 'bad default color falls back to gray');

// ------------------------------------------------------------- role rename
section('role rename');
var sRr = mkState([{ num: 1, feature: 'a', teamType: 'Development' }], {
  team: [{ name: 'Ada', type: 'Development' }]
});
sRr.meta.rateCard = { Development: { rate: 200, cost: 90 } };
ok(RM.renameRole(sRr, 'Development', 'Engineer'), 'rename succeeds');
eq(sRr.teamTypes.indexOf('Engineer') !== -1, true, 'role list renamed');
eq(sRr.team[0].type, 'Engineer', 'people follow the rename');
eq(sRr.items[0].teamType, 'Engineer', 'items follow the rename');
eq(sRr.meta.rateCard.Engineer.rate, 200, 'rate card key follows the rename');
ok(!RM.renameRole(sRr, 'Engineer', 'Data'), 'rename refuses an existing name');
eq(RM.DEFAULT_TEAM_TYPES[0], 'Project Manager', 'default roles lead with Project Manager and use real role names');
ok(RM.DEFAULT_TEAM_TYPES.indexOf('Software Engineer') !== -1 && RM.DEFAULT_CAP_TYPES.slice(0, 3).join(',') === 'Development,Design,QA',
  'default capacity types are Development, Design, QA, …');

// ------------------------------------------------------------- rigid chain drag
section('rigid chain shift');
var sRg = mkState([
  { num: 1, feature: 'root', startDay: 0, durDays: 5 },
  { num: 2, feature: 'mid', deps: [1], startDay: 20, durDays: 5 },
  { num: 3, feature: 'leaf', deps: [2], startDay: 40, durDays: 5 },
  { num: 4, feature: 'pin', deps: [1], startDay: 30, durDays: 5, milestone: true }
]);
eq(RM.shiftDependents(sRg, sRg.items[0].id, 5, { rigid: true }), 3, 'rigid push moves the whole chain, slack or not');
eq(sRg.items[1].startDay, 25, 'direct dependent moved by the full delta');
eq(sRg.items[2].startDay, 45, 'transitive dependent moved by the full delta');
eq(sRg.items[3].startDay, 35, 'milestones ride along under a rigid drag');
RM.shiftDependents(sRg, sRg.items[0].id, -5, { rigid: true });
eq(sRg.items[2].startDay, 40, 'rigid pull brings the chain back');

// ------------------------------------------------------------- deadlines
section('deadlines');
var sDl = mkState([
  { num: 1, feature: 'a', startDay: 0, durDays: 10, deadline: '2026-08-06' },
  { num: 2, feature: 'b', startDay: 0, durDays: 5, deadline: '2026-08-07' },
  { num: 3, feature: 'c', deadline: 'garbage' },
  { num: 4, feature: 'ms', milestone: true, startDay: 12, durDays: 0, deadline: '2026-08-07' }
]);
eq(sDl.items[0].deadline, '2026-08-06', 'valid ISO deadline kept');
eq(sDl.items[2].deadline, null, 'junk deadline dropped');
eq(RM.deadlineDay(sDl.meta, sDl.items[1]), 9, 'deadline maps to its day index');
ok(RM.pastDeadline(sDl.meta, sDl.items[0]), '2-week bar past a Thursday-week-2 deadline');
ok(!RM.pastDeadline(sDl.meta, sDl.items[1]), '1-week bar within the deadline');
ok(!RM.pastDeadline(sDl.meta, sDl.items[2]), 'no deadline, never late');
ok(RM.pastDeadline(sDl.meta, sDl.items[3]), 'milestone after its deadline is late');
ok(RM.SCOPE_FIXED_KEYS.indexOf('deadline') !== -1, 'deadline is a fixed scoping column');
eq(RM.SCOPE_DEFAULT_ORDER.indexOf('deadline'), RM.SCOPE_DEFAULT_ORDER.indexOf('start') + 1,
  'deadline defaults to just after Start');
// docs saved before the column slot it after Start, not at the end
var sDlOrd = RM.normalizeState({
  meta: { timelineStart: '2026-07-27', numWeeks: 8, holidaysV2026: true,
    scopeColOrder: ['description', 'epic', 'assignees', 'size', 'risk', 'priority', 'duration', 'start', 'workstream'] },
  phases: [{ id: 'p' }], items: []
});
eq(sDlOrd.meta.scopeColOrder.indexOf('deadline'), sDlOrd.meta.scopeColOrder.indexOf('start') + 1,
  'saved order gains Deadline right after Start');

// ------------------------------------------------------------- multi-workstream people
section('multi-workstream people');
var sMw = mkState([{ num: 1, feature: 'a', startDay: 0, durDays: 5, size: 'M', workstream: 'OS' }], {
  team: [
    { name: 'Ada', type: 'Development', workstream: 'OS' },
    { name: 'Grace', type: 'Development', workstreams: ['OS', 'Apps', 'OS'] }
  ]
});
eq(RM.memberWorkstreams(sMw.team[0]), ['OS'], 'legacy single workstream migrates to a list');
eq(RM.memberWorkstreams(sMw.team[1]), ['OS', 'Apps'], 'workstream list dedupes');
eq(sMw.team[1].workstream, 'OS', 'primary workstream mirrors the first entry');
RM.setMemberWorkstreams(sMw.team[0], ['Apps', 'OS']);
eq(sMw.team[0].workstream, 'Apps', 'setMemberWorkstreams keeps the primary in sync');
RM.setMemberWorkstreams(sMw.team[0], []);
eq(sMw.team[0].workstream, '', 'clearing the list clears the primary');
// a person on two workstreams splits hours and cost between them
var sMwR = mkState([], {
  team: [{ name: 'Ada', type: 'Development', workstreams: ['OS', 'Apps'], cost: 100 }]
});
var repMw = RM.costReport(sMwR, 'workstream');
var rowOS = repMw.rows.filter(function (r) { return r.key === 'OS'; })[0];
var rowApps = repMw.rows.filter(function (r) { return r.key === 'Apps'; })[0];
ok(rowOS && rowApps && Math.abs(rowOS.roleCost - rowApps.roleCost) < 0.01,
  'roster cost splits evenly across the two workstreams');
ok(Math.abs((rowOS.roleHours + rowApps.roleHours) - RM.roleTotalHours(sMwR, sMwR.team[0])) < 0.01,
  'split hours add back up to the person\'s total');

// ------------------------------------------------------------- story fields
section('story fields');
var sSfMeta = JSON.parse(JSON.stringify(META));
sSfMeta.priorityScheme = 'levels';
var sSf = RM.normalizeState({
  meta: sSfMeta,
  phases: [{ id: 'p1', name: 'Alpha', bucket: false }],
  items: [{ num: 1, feature: 'a', stories: [
    { title: 's1', priority: 'h', assignees: ['ghost'], durDays: 5, deadline: '2026-09-04' },
    { title: 's2', priority: 'zz', deadline: 'soonish' }
  ] }],
  team: [{ id: 'tm1', name: 'Ada', type: 'Development' }],
  teamTypes: ['Development']
});
eq(sSf.items[0].stories[0].priority, 'H', 'story priority validated + uppercased');
eq(sSf.items[0].stories[1].priority, null, 'unknown story priority dropped');
eq(sSf.items[0].stories[0].assignees, [], 'story assignees validated against the roster');
eq(sSf.items[0].stories[0].durDays, 5, 'unscheduled story keeps its duration');
eq(sSf.items[0].stories[0].deadline, '2026-09-04', 'story deadline kept');
eq(sSf.items[0].stories[1].deadline, null, 'junk story deadline dropped');
var sSf2 = mkState([{ num: 1, feature: 'a', assignees: ['tm1'], stories: [{ title: 's', assignees: ['tm1', 'tm1', 'nope'] }] }], {
  team: [{ id: 'tm1', name: 'Ada', type: 'Development' }]
});
eq(sSf2.items[0].stories[0].assignees, ['tm1'], 'story assignees dedupe and keep real people');

// ------------------------------------------------------------- jira keys
section('jira keys');
var sJk = mkState([{ num: 1, feature: 'a', epic: 'Login', jiraKey: ' hw-12 ',
  stories: [{ title: 's', jiraKey: 'HW-13' }, { title: 't', jiraKey: 42 }] }],
  { epicJira: { Login: 'HW-1', Stale: '', Junk: 7 } });
eq(sJk.items[0].jiraKey, 'HW-12', 'item jira key trimmed and uppercased');
eq(sJk.items[0].stories[0].jiraKey, 'HW-13', 'story jira key kept');
eq(sJk.items[0].stories[1].jiraKey, null, 'non-string story jira key dropped');
eq(sJk.epicJira, { Login: 'HW-1' }, 'epicJira keeps only non-empty string keys');
var sJk2 = mkState([{ num: 1, feature: 'a' }]);
eq(sJk2.items[0].jiraKey, null, 'item jira key defaults to null');
eq(sJk2.epicJira, {}, 'epicJira defaults to an empty map');

// ------------------------------------------------------------------- tags
section('tags');
eq(RM.normalizeTags(' a , b,, A ,c'), ['a', 'b', 'c'], 'a comma string splits, trims, drops empties and dedupes case-insensitively');
eq(RM.normalizeTags(['x', ' y ', 'X', '']), ['x', 'y'], 'array input normalizes the same way');
eq(RM.normalizeTags('Alpha, alpha'), ['Alpha'], 'the first spelling wins');
eq(RM.normalizeTags([new Array(60).join('z')])[0].length, 40, 'each tag caps at 40 chars');
eq(RM.normalizeTags(null), [], 'nothing in, nothing out');
eq(RM.normalizeTags(7), [], 'a non-string, non-array value yields no tags');
{
  var sTg = mkState([
    { num: 1, feature: 'a', tags: ['red', 7, null, 'Blue'], stories: [{ title: 's', tags: 'green, red' }] },
    { num: 2, feature: 'b', tags: 'ZED' }
  ]);
  eq(sTg.items[0].tags, ['red', 'Blue'], 'normalizeState keeps item tags and drops non-strings');
  eq(sTg.items[0].stories[0].tags, ['green', 'red'], 'story tags normalize from a comma string');
  eq(sTg.items[1].tags, ['ZED'], 'a plain string tag field becomes a one-tag array');
  eq(mkState([{ num: 1, feature: 'a' }]).items[0].tags, [], 'tags default to an empty array');
  eq(RM.allTags(sTg), ['Blue', 'green', 'red', 'ZED'], 'allTags unions features and stories, sorted case-insensitively');
  RM.setTags(sTg, sTg.items[0], ' one , one , two ');
  eq(sTg.items[0].tags, ['one', 'two'], 'setTags normalizes what it stores');
}

// ------------------------------------------------------------- jira csv export
section('jira csv export');
// ------------------------------------------------------------ sprint moves
section('sprint moves');
{
  var sSp = mkState([
    { id: 'a', num: 1, phaseId: 'p1', feature: 'A', size: 'M', startDay: 0, durDays: 5,
      stories: [{ id: 'a1', title: 's1', startDay: 2, durDays: 3 }, { id: 'a2', title: 's2' }] },
    { id: 'b', num: 2, phaseId: 'p1', feature: 'B', size: 'M', startDay: 0, durDays: 5 },
    { id: 'c', num: 3, phaseId: 'p1', feature: 'C', size: 'L' },
    { id: 'd', num: 4, phaseId: 'p2', feature: 'D', size: 'S', startDay: 20, durDays: 3 }
  ]);
  eq(RM.sprintStartDay(sSp.meta, 1), 0, 'sprint 1 starts on day 0');
  eq(RM.sprintStartDay(sSp.meta, 3), 20, 'sprint 3 starts on day 20 (2-week sprints)');
  eq(RM.sprintStartDay(sSp.meta, -5), 0, 'sprints before the timeline clamp to day 0');

  RM.moveItemToSprint(sSp, 'a', 3, null);
  var a = RM.itemById(sSp, 'a');
  eq(a.startDay, 20, 'moving to sprint 3 lands the item on its first day');
  eq(a.durDays, 5, 'the duration is kept');
  eq(a.stories[0].startDay, 22, 'stories with a timeline ride along by the same delta');
  ok(a.stories[1].startDay == null, 'stories without a timeline stay put');
  eq(sSp.items.map(function (x) { return x.id; }), ['b', 'c', 'a', 'd'], 'no before-item: it moves to the end of its phase');

  RM.moveItemToSprint(sSp, 'c', 2, 'a');
  var c = RM.itemById(sSp, 'c');
  eq(c.startDay, 10, 'an unscheduled item dropped in sprint 2 gets that start');
  eq(c.durDays, RM.stretchSpan(sSp.meta, 10, RM.effortDays(sSp, c)), 'and a span from its size');
  eq(sSp.items.map(function (x) { return x.id; }), ['b', 'c', 'a', 'd'], 'reordered before the given item');

  RM.moveItemToSprint(sSp, 'b', 1, 'd');
  var b = RM.itemById(sSp, 'b');
  eq(b.phaseId, 'p2', 'dropping before an item in another phase adopts that phase');
  eq(sSp.items.map(function (x) { return x.id; }), ['c', 'a', 'b', 'd'], 'and sits right before it');
  eq(b.startDay, 0, 'same sprint: the start does not change');

  RM.moveItemToSprint(sSp, 'd', null, null);
  var d = RM.itemById(sSp, 'd');
  ok(d.startDay == null && d.durDays == null, 'dropping on Unscheduled clears the timeline');

  ok(RM.reorderItem(sSp, 'a', 'a') === false, 'reordering before itself is a no-op');
  a.holdPos = true;
  RM.reorderItem(sSp, 'a', 'c');
  ok(!a.holdPos, 'a manual reorder clears holdPos');
  eq(sSp.items.map(function (x) { return x.id; }), ['a', 'c', 'b', 'd'], 'reorder before another item');

  RM.moveStoryToSprint(sSp, 'a', 'a2', 2, 'a1');
  eq(a.stories.map(function (x) { return x.id; }), ['a2', 'a1'], 'story reorders before its sibling');
  eq(a.stories[0].startDay, 10, 'story lands on the sprint start');
  eq(a.stories[0].durDays, RM.sprintDays(sSp.meta), 'a story without a span gets one sprint');
  RM.moveStoryToSprint(sSp, 'a', 'a2', null, null);
  ok(a.stories[1].id === 'a2' && a.stories[1].startDay == null && a.stories[1].durDays == null,
    'unscheduling a story clears its timeline and moves it last');
}

var RMJira = require('../js/export-jira.js');
var sJc = mkState([
  { num: 1, feature: 'Login page', epic: 'Login', workstream: 'Product', size: 'M',
    startDay: 0, durDays: 5, deadline: '2026-09-04', jiraKey: 'HW-12',
    description: '<p>Hi <b>there</b></p>', enables: 'Checkout', notes: '',
    stories: [{ title: 's1', done: true }, { title: 's2', jiraKey: 'HW-13', type: 'bug' }] },
  { num: 2, feature: 'Search, "fast"', deps: [1], phaseId: 'p2', type: 'bug' },
  { num: 3, feature: 'Orphan', deps: [2] }
], { epicJira: { Login: 'HW-1' } });
eq(RMJira.fileName(sJc), 'T-jira.csv', 'jira csv filename');
var jr = RMJira.rows(sJc, { features: true, stories: false });
eq(jr.length, 3, 'features only: one row per feature');
var r1 = jr[0];
eq(r1['Summary'], 'Login page', 'summary is the feature name');
eq(r1['Issue Type'], 'Story', 'feature type Feature maps to Jira Story');
eq(jr[1]['Issue Type'], 'Bug', 'a Bug feature maps to Jira Bug');
eq(r1['Parent'], 'HW-1', 'parent is the epic jira key');
eq(r1['Labels'], 'ws-product phase-alpha size-m', 'labels are slugged workstream, phase and size');
eq(r1['Due Date'], '2026-09-04', 'due date is the deadline');
eq(r1['Start Date'], '2026-07-27', 'start date from the schedule');
eq(r1['End Date'], '2026-07-31', 'inclusive end date from the schedule');
eq(r1['Jira Key'], 'HW-12', 'jira key column carries the existing key');
ok(r1['Description'].indexOf('Hi there') === 0, 'description leads with the plain-text description');
ok(r1['Description'].indexOf('[x] s1') !== -1 && r1['Description'].indexOf('[ ] s2') !== -1,
  'stories render as a checklist when not exported as rows');
ok(r1['Description'].indexOf('Enables:\nCheckout') !== -1, 'non-empty scope fields become sections');
ok(r1['Description'].indexOf('Notes') === -1, 'empty scope fields are skipped');
var r2 = jr[1];
eq(r2['Parent'], '', 'no epic key: parent blank');
eq(r2['Labels'], 'phase-next', 'no workstream or size: only the phase label');
eq(r2['Blocked By'], 'HW-12', 'dependencies with keys list the key');
eq(r2['Start Date'], '', 'unscheduled: blank dates');
eq(jr[2]['Blocked By'], '', 'dependencies without keys are left out');

var jrs = RMJira.rows(sJc, { features: true, stories: true });
eq(jrs.length, 5, 'features and stories: a row per story too');
eq(jrs[0]['Issue Type'], 'Story', 'feature row type from the type record');
ok(jrs[0]['Description'].indexOf('[x]') === -1, 'checklist omitted when stories are rows');
eq(jrs[1]['Summary'], 's1', 'story row summary');
eq(jrs[1]['Issue Type'], 'Sub-task', 'story issue type from the type record');
eq(jrs[2]['Issue Type'], 'Bug', 'a Bug story maps to Jira Bug');
eq(jrs[1]['Parent'], 'HW-12', 'story parents to the feature key');
eq(jrs[1]['Labels'], 'feature-login-page ws-product phase-alpha', 'story labels name the feature');
eq(jrs[2]['Jira Key'], 'HW-13', 'story jira key');
var jro = RMJira.rows(sJc, { features: false, stories: true });
eq(jro.length, 2, 'stories only');

var csv = RMJira.csv(sJc, { features: true, stories: false });
ok(csv.charCodeAt(0) === 0xFEFF, 'csv starts with a UTF-8 BOM');
var lines = csv.slice(1).split('\r\n');
eq(lines[0], 'Summary,Issue Type,Description,Parent,Labels,Priority,Due Date,Start Date,End Date,Blocked By,Jira Key', 'header row');
ok(lines.some(function (l) { return l.indexOf('"Search, ""fast"""') === 0; }), 'commas and quotes are escaped');
ok(/"Hi there\n/.test(csv), 'newlines stay inside a quoted cell');
eq(RMJira.csv(mkState([]), { features: true }).slice(1).split('\r\n').length, 2, 'empty doc: header plus trailing newline');

// ------------------------------------------------------------ half points
section('half points');
{
  eq(RM.SIZE_SCHEMES.fibonacci.sizes.indexOf('0.5'), 1, 'story points offer 0.5 right after 0');
  eq(RM.SIZE_SCHEMES.fibonacci.days['0.5'], 0.5, '0.5 points = half a day');
  var sH = RM.normalizeState({ meta: { title: 'H', timelineStart: '2026-07-27', numWeeks: 8, storySizeScheme: 'fibonacci', storySizeOrder: ['1', '2', '3', '5', '8', '13'], storySizeDays: { 1: 1, 2: 2, 3: 3, 5: 5, 8: 10, 13: 20 } },
    phases: [{ id: 'p', name: 'P' }], team: [], items: [{ id: 'x', num: 1, phaseId: 'p', feature: 'X', startDay: 0, durDays: 5, stories: [{ id: 's', title: 'S', size: '0.5' }] }] });
  eq(RM.sizeOrderOf(sH, 'story').indexOf('0.5'), 1, 'an older Fibonacci document gains the 0.5 option');
  eq(RM.sizeDays(sH, '0.5', 'story'), 0.5, 'and its day value');
  eq(RM.storyEffortDays(sH, sH.items[0].stories[0]), 0.5, 'a 0.5-point story is half a day of effort');
  ok(RM.stretchSpan(sH.meta, 0, 0.5) >= 1, 'a half-day span still occupies one working day on the grid');
  var sC = RM.normalizeState({ meta: { title: 'C', timelineStart: '2026-07-27', numWeeks: 8, storySizeScheme: 'custom', storySizeOrder: ['1', '2', '3', '5', '8', '13'], storySizeDays: { 1: 1, 2: 2, 3: 3, 5: 5, 8: 10, 13: 20 } }, phases: [{ id: 'p', name: 'P' }], team: [], items: [] });
  eq(RM.sizeOrderOf(sC, 'story').indexOf('0.5'), -1, 'a custom scale is left alone');
  RM.addSizeOption(sC, '0.5', 0.5, 'story');
  eq(RM.sizeDays(sC, '0.5', 'story'), 0.5, 'addSizeOption keeps a fractional day value');
}

// tags ride along as slugged Jira labels
{
  var sTl = mkState([{ num: 1, feature: 'Login page', workstream: 'Product', tags: ['Tech Debt', 'q3'],
    stories: [{ title: 's1', tags: ['Story Tag'] }] }]);
  var tr1 = RMJira.rows(sTl, { features: true, stories: true });
  ok(tr1[0]['Labels'].split(' ').indexOf('tech-debt') !== -1 && tr1[0]['Labels'].split(' ').indexOf('q3') !== -1,
    'feature labels include the slugged tags');
  ok(tr1[1]['Labels'].split(' ').indexOf('story-tag') !== -1, 'story labels include the slugged story tags');
}

// ------------------------------------------------------------ zero points
section('zero points');
{
  eq(RM.SIZE_SCHEMES.fibonacci.sizes[0], '0', 'story points start at 0');
  eq(RM.SIZE_SCHEMES.fibonacci.days['0'], 0, '0 points = no effort');
  // features never gain the 0 step: picking Story points for the feature
  // scale still starts at 0.5
  var sZF = mkState([]);
  RM.setSizeScheme(sZF, 'fibonacci', 'feature');
  eq(RM.sizeOrderOf(sZF).join(','), '0.5,1,2,3,5,8,13', 'the feature Fibonacci scale skips 0');
  eq(RM.sizeDays(sZF, '0'), null, 'and carries no day value for it');
  RM.setSizeScheme(sZF, 'fibonacci', 'story');
  eq(RM.sizeOrderOf(sZF, 'story').join(','), '0,0.5,1,2,3,5,8,13', 'the story Fibonacci scale offers 0');

  // a size option may be worth zero days
  var sZA = mkState([]);
  RM.addSizeOption(sZA, '0', 0, 'story');
  eq(RM.sizeDays(sZA, '0', 'story'), 0, 'addSizeOption stores a zero day value');

  // a 0-point story: no effort, but still a day on the grid once scheduled
  var sZ = RM.normalizeState({ meta: { title: 'Z', timelineStart: '2026-07-27', numWeeks: 8, storySizeScheme: 'fibonacci' },
    phases: [{ id: 'p', name: 'P' }], team: [],
    items: [{ id: 'x', num: 1, phaseId: 'p', feature: 'X', startDay: 0, durDays: 5, stories: [{ id: 's', title: 'S', size: '0' }] }] });
  eq(RM.storyEffortDays(sZ, sZ.items[0].stories[0]), 0, 'a 0-point story is zero days of effort');
  ok(RM.moveStoryToSprint(sZ, 'x', 's', 0), 'a 0-point story moves onto a sprint');
  ok(sZ.items[0].stories[0].durDays >= 1, 'and occupies at least one working day (' + sZ.items[0].stories[0].durDays + ')');
  var sZ2 = RM.normalizeState(sZ);
  ok(sZ2.items[0].stories[0].startDay != null && sZ2.items[0].stories[0].durDays >= 1,
    'the 0-point story survives normalize still scheduled');

  // migration: the Task 15 default gains 0; the pre-0.5 default gains both
  var sZ15 = RM.normalizeState({ meta: { title: 'Z15', timelineStart: '2026-07-27', numWeeks: 8, storySizeScheme: 'fibonacci', storySizeOrder: ['0.5', '1', '2', '3', '5', '8', '13'], storySizeDays: { '0.5': 0.5, 1: 1, 2: 2, 3: 3, 5: 5, 8: 10, 13: 20 } },
    phases: [{ id: 'p', name: 'P' }], team: [], items: [] });
  eq(RM.sizeOrderOf(sZ15, 'story').join(','), '0,0.5,1,2,3,5,8,13', 'a 0.5-era document gains the 0 option');
  eq(RM.sizeDays(sZ15, '0', 'story'), 0, 'and its zero day value survives normalize');
  var sZOld = RM.normalizeState({ meta: { title: 'ZO', timelineStart: '2026-07-27', numWeeks: 8, storySizeScheme: 'fibonacci', storySizeOrder: ['1', '2', '3', '5', '8', '13'], storySizeDays: { 1: 1, 2: 2, 3: 3, 5: 5, 8: 10, 13: 20 } },
    phases: [{ id: 'p', name: 'P' }], team: [], items: [] });
  eq(RM.sizeOrderOf(sZOld, 'story').join(','), '0,0.5,1,2,3,5,8,13', 'a pre-0.5 document gains both 0 and 0.5');
  var sZC = RM.normalizeState({ meta: { title: 'ZC', timelineStart: '2026-07-27', numWeeks: 8, storySizeScheme: 'custom', storySizeOrder: ['0.5', '1', '2', '3', '5', '8', '13'], storySizeDays: { '0.5': 0.5, 1: 1, 2: 2, 3: 3, 5: 5, 8: 10, 13: 20 } },
    phases: [{ id: 'p', name: 'P' }], team: [], items: [] });
  eq(RM.sizeOrderOf(sZC, 'story').indexOf('0'), -1, 'a custom story scale is left alone');
  var sZFO = RM.normalizeState({ meta: { title: 'ZF', timelineStart: '2026-07-27', numWeeks: 8, sizeScheme: 'fibonacci', sizeOrder: ['1', '2', '3', '5', '8', '13'], sizeDays: { 1: 1, 2: 2, 3: 3, 5: 5, 8: 10, 13: 20 } },
    phases: [{ id: 'p', name: 'P' }], team: [], items: [] });
  eq(RM.sizeOrderOf(sZFO).join(','), '0.5,1,2,3,5,8,13', 'an older feature Fibonacci scale gains 0.5 only');
}

// -------------------------------------------- Auto timeline sizes (Stories level)
section('auto timeline sizes');
{
  var TSHIRT = { XS: 2, S: 5, M: 10, L: 20, XL: 40 };
  function asMeta(over) {
    var m = { planLevel: 'story', sizeScheme: 'tshirt', sizeOrder: ['XS', 'S', 'M', 'L', 'XL'], sizeDays: JSON.parse(JSON.stringify(TSHIRT)),
      storySizeScheme: 'fibonacci' };
    if (over) Object.keys(over).forEach(function (k) { m[k] = over[k]; });
    return m;
  }
  function asState(stories, over, phases) {
    return autoState([{ num: 1, feature: 'F', phaseId: 'p1', capType: 'Development', size: 'XL', stories: stories }],
      [{ name: 'Solo', capType: 'Development' }], asMeta(over), phases);
  }
  var f1 = function (s) { return RM.itemByNum(s, 1); };
  function sizeOf(changes, s) {
    var c = changes.filter(function (x) { return x.itemId === f1(s).id; })[0];
    return c ? c.size : null;
  }

  // parallel stories don't add up: the size follows the span their bars cover
  var sPar = asState([
    { num: 101, title: 'a', size: '5', startDay: 0, durDays: 5, capType: 'Development' },
    { num: 102, title: 'b', size: '5', startDay: 0, durDays: 5, capType: 'Development' }
  ]);
  eq(RM.autoSizeDays(sPar, f1(sPar)), 5, 'two parallel 5-day stories span 5 working days');
  eq(sizeOf(RM.autoSizeChanges(sPar, 'p1'), sPar), 'S', 'and the derived size is the nearest label (S)');
  eq(f1(sPar).size, 'XL', 'nothing is written until the Auto action applies it — the size stays a hand size');
  eq(RM.itemSizeDays(sPar, f1(sPar)), 40, 'and itemSizeDays reads the stored size, never a derived one');

  var sSeq = asState([
    { num: 101, title: 'a', size: '5', startDay: 0, durDays: 5, capType: 'Development' },
    { num: 102, title: 'b', size: '5', startDay: 5, durDays: 5, capType: 'Development' }
  ]);
  eq(sizeOf(RM.autoSizeChanges(sSeq, 'p1'), sSeq), 'M', 'two sequential 5-day stories span 10 working days → M');

  // nothing scheduled: fall back to the sized stories added up
  var sUns = asState([
    { num: 101, title: 'a', size: '5', capType: 'Development' },
    { num: 102, title: 'b', size: '5', capType: 'Development' }
  ]);
  eq(sizeOf(RM.autoSizeChanges(sUns, 'p1'), sUns), 'M', 'unscheduled sized stories size from their days summed');

  // no sized story, Features level, another phase, the none / rollup schemes: nothing to write
  var sNo = asState([{ num: 101, title: 'a', startDay: 0, durDays: 5, capType: 'Development' }]);
  eq(RM.autoSizeChanges(sNo, 'p1'), [], 'no sized story means no size change');
  var sFeat = asState([{ num: 101, title: 'a', size: '5', startDay: 0, durDays: 5, capType: 'Development' }], { planLevel: 'feature' });
  eq(RM.autoSizeChanges(sFeat, 'p1'), [], 'the Features level never sizes from stories');
  eq(RM.autoSizeChanges(sPar, 'p2'), [], 'only the asked phase is sized');
  var sNone = asState([{ num: 101, title: 'a', size: '5', startDay: 0, durDays: 5, capType: 'Development' }],
    { sizeScheme: 'none', sizeOrder: [], sizeDays: {} });
  eq(RM.autoSizeChanges(sNone, 'p1'), [], 'the none scheme is never sized');
  var sRl = asState([{ num: 101, title: 'a', size: '5', startDay: 0, durDays: 5, capType: 'Development' }], { sizeScheme: 'rollup' });
  eq(RM.autoSizeChanges(sRl, 'p1'), [], 'the rollup scheme derives sizes on its own');
  var sSame = asState([{ num: 101, title: 'a', size: '5', startDay: 0, durDays: 5, capType: 'Development' }]);
  f1(sSame).size = 'S';
  eq(RM.autoSizeChanges(sSame, 'p1'), [], 'a feature already at its derived size is not a change');

  // the derived day count rounds UP to the feature snap
  var sSnap = asState([{ num: 101, title: 'a', size: '5', startDay: 0, durDays: 6, capType: 'Development' }]);
  eq(RM.autoSizeDays(sSnap, f1(sSnap)), 6, 'day snap leaves the 6-day hull alone');
  eq(sizeOf(RM.autoSizeChanges(sSnap, 'p1'), sSnap), 'S', 'and 6 days reads as S');
  eq(RM.autoSizeDays(sSnap, f1(sSnap), { snap: { feature: 'week' } }), 10, 'week snap rounds the 6-day hull up to 10');
  eq(sizeOf(RM.autoSizeChanges(sSnap, 'p1', { snap: { feature: 'week' } }), sSnap), 'M', 'so the feature reads as the two-week label');
  eq(RM.snapUnitDays(sSnap.meta, 'day'), 1, 'snapUnitDays: a day is one slot');
  eq(RM.snapUnitDays(sSnap.meta, 'week'), 5, 'snapUnitDays: a week is the work week');
  eq(RM.snapUnitDays(sSnap.meta, 'sprint'), 10, 'snapUnitDays: a sprint is its weeks');

  // autoPhase: the one-shot action — lay the phase out, then size its features
  var sAP = asState([
    { num: 101, title: 'a', size: '5', durDays: 5, capType: 'Development' },
    { num: 102, title: 'b', size: '5', durDays: 5, capType: 'Development', deps: [101] }
  ]);
  var rAP = RM.autoPhase(sAP, 'p1', { today: 0 });
  var apF = f1(rAP.state);
  eq([apF.stories[0].startDay, apF.stories[1].startDay], [0, 5], 'autoPhase lays the stories out');
  eq(apF.size, 'M', 'and sizes the feature from the span they now cover');
  ok(rAP.moved > 0 && rAP.sized.length === 1, 'it reports the moves and the sizes it wrote');
  eq(rAP.changed, rAP.moved + rAP.sized.length + 1, 'changed counts the moves, the sizes and the rebuilt feature hull');
  eq(f1(sAP).size, 'XL', 'the input state is left untouched');
  // dry-run equivalence: a second pass over its own result changes nothing
  var rAP2 = RM.autoPhase(rAP.state, 'p1', { today: 0 });
  eq(rAP2.changed, 0, 'a pass over an already arranged phase changes nothing (the button disables)');
  // after the action the size is an ordinary hand size again
  var sHand = RM.clone(rAP.state);
  f1(sHand).size = 'L';
  RM.applySizeRollup(sHand);
  eq(f1(sHand).size, 'L', 'a hand edit after the action sticks');
  eq(RM.autoPhase(sHand, 'p1', { today: 0 }).changed, 1, 'until the next Auto click, which re-derives it');
  // capacity off / a bucket phase: nothing
  var sAPoff = asState([{ num: 101, title: 'a', size: '5', durDays: 5, capType: 'Development' }], { capacityEnabled: false });
  eq(RM.autoPhase(sAPoff, 'p1', { today: 0 }).changed, 0, 'capacity off → autoPhase changes nothing');
  eq(RM.autoPhase(sAP, 'p3', { today: 0 }).changed, 0, 'a bucket phase → nothing');
  // locked items are fixed points under autoPhase too
  var sAPL = autoState([
    { num: 1, feature: 'anchor', phaseId: 'p1', startDay: 20, durDays: 5, locked: true, capType: 'Development' },
    { num: 2, feature: 'follows', phaseId: 'p1', startDay: 0, durDays: 5, capType: 'Development', deps: [1] }
  ], [{ name: 'X', capType: 'Development' }, { name: 'Y', capType: 'Development' }]);
  var rAPL = RM.autoPhase(sAPL, 'p1', { today: 0 });
  eq(RM.itemByNum(rAPL.state, 1).startDay, 20, 'autoPhase leaves the locked item where it is');
  eq(RM.itemByNum(rAPL.state, 2).startDay, 25, 'and puts its dependent after it');
  eq(RM.autoPhase(rAPL.state, 'p1', { today: 0 }).changed, 0, 'then nothing is left to move');

  // (a) at the Stories level the phase floor reads story bars too, so a
  // feature hull rebuilt after the pass can never lower it
  var sFlS = autoState([{ num: 1, feature: 'F', phaseId: 'p1', startDay: 20, durDays: 5, capType: 'Development',
    stories: [{ num: 101, title: 'a', startDay: 12, durDays: 3, capType: 'Development' }] }],
    [{ name: 'Solo', capType: 'Development' }], { planLevel: 'story' });
  eq(RM.phaseFloorDay(sFlS, sFlS.phases[0]), 12, 'the derived phase floor is the earliest of story and feature bars');

  // (b) auto-order: the action sorts the rows by start itself and lays out
  // again until nothing moves, so one click settles
  var sAO = autoState([
    { num: 1, feature: 'late', phaseId: 'p1', durDays: 5, capType: 'Development', deps: [3] },
    { num: 2, feature: 'mid', phaseId: 'p1', durDays: 5, capType: 'Development' },
    { num: 3, feature: 'first', phaseId: 'p1', durDays: 5, capType: 'Development' }
  ], [{ name: 'Solo', capType: 'Development' }]);
  var rAO = RM.autoPhase(sAO, 'p1', { today: 0, autoOrder: true });
  var aoStarts = rAO.state.items.filter(function (i) { return i.phaseId === 'p1'; }).map(function (i) { return i.startDay; });
  eq(aoStarts, aoStarts.slice().sort(function (a, b) { return a - b; }), 'with autoOrder the rows come back in start order');
  eq(RM.autoPhase(rAO.state, 'p1', { today: 0, autoOrder: true }).changed, 0, 'and a second click finds nothing to do');
  eq(rAO.moved, 3, 'moved counts the three units once each, however many passes it took');

  // moved counts units, not the feature hulls rebuilt around them
  var sMvS = autoState([{ num: 1, feature: 'F', phaseId: 'p1', startDay: 0, durDays: 5, capType: 'Development', stories: [
    { num: 101, title: 'a', startDay: 0, durDays: 5, capType: 'Development' },
    { num: 102, title: 'b', startDay: 0, durDays: 5, capType: 'Development' }] }],
    [{ name: 'Solo', capType: 'Development' }], { planLevel: 'story' });
  eq(RM.autoPhase(sMvS, 'p1', { today: 0 }).moved, 1, 'one story moved reads as one, not story + hull');
}

// -------------------------------------------- snapped placement
section('snapped placement');
{
  function snapSt() {
    return autoState([{ num: 1, feature: 'F', phaseId: 'p1', capType: 'Development',
      stories: [{ num: 101, title: 'a', durDays: 3, capType: 'Development' }] }],
      [{ name: 'Solo', capType: 'Development' }], { planLevel: 'story' });
  }
  var rPlain = RM.autoTimeline(snapSt(), { phaseIds: ['p1'], today: 2 });
  eq(RM.itemByNum(rPlain.state, 1).stories[0].startDay, 2, 'without a snap the story starts the day it is ready');
  var rWeek = RM.autoTimeline(snapSt(), { phaseIds: ['p1'], today: 2, snap: { story: 'week', feature: 'week' } });
  var stW = RM.itemByNum(rWeek.state, 1).stories[0];
  eq(stW.startDay, 5, 'week snap pushes a story ready on day 2 to the next week');
  eq(stW.durDays, 5, 'and rounds its 3 days up to a whole week');
  var rSpr = RM.autoTimeline(snapSt(), { phaseIds: ['p1'], today: 2, snap: { story: 'sprint', feature: 'sprint' } });
  var stS = RM.itemByNum(rSpr.state, 1).stories[0];
  eq(stS.startDay, 10, 'sprint snap lands on the next two-week boundary');
  eq(stS.durDays, 10, 'and fills the sprint');
  // work already under way is left alone by the auto pass
  var sIn = autoState([{ num: 1, feature: 'F', phaseId: 'p1', capType: 'Development',
    stories: [{ num: 101, title: 'a', startDay: 3, durDays: 5, capType: 'Development' }] }],
    [{ name: 'Solo', capType: 'Development' }], { planLevel: 'story' });
  var stIn = RM.itemByNum(RM.autoTimeline(sIn, { phaseIds: ['p1'], today: 6, snap: { story: 'week', feature: 'week' } }).state, 1).stories[0];
  eq([stIn.startDay, stIn.durDays], [3, 5], 'an in-flight story keeps its start and duration under a snap');
  // started work that capacity pushes off its start takes the snap where it
  // lands, so a second pass finds nothing left to do
  var sPush = autoState([
    { num: 1, feature: 'busy', phaseId: 'p1', startDay: 0, durDays: 15, locked: true, capType: 'Development' },
    { num: 2, feature: 'begun', phaseId: 'p1', startDay: 1, durDays: 3, capType: 'Development' }
  ], [{ name: 'Solo', capType: 'Development' }]);
  var wk = { phaseIds: ['p1'], today: 2, snap: { story: 'week', feature: 'week' } };
  var rPush = RM.autoTimeline(sPush, wk);
  var pushed = RM.itemByNum(rPush.state, 2);
  eq([pushed.startDay, pushed.durDays], [15, 5], 'begun work pushed by capacity lands on the week grid with whole weeks');
  eq(RM.autoTimeline(rPush.state, wk).changed, 0, 'and a second pass leaves it there');
  // (c) …and when it is pushed to today or later it is new work: the phase floor holds it too
  var phPin30b = [{ id: 'p1', name: 'Alpha', bucket: false, startDay: 20 }, { id: 'p3', name: 'Next', bucket: true }];
  ['day', 'week'].forEach(function (m) {
    var sPin = autoState([
      { num: 1, feature: 'busy', phaseId: 'p1', startDay: 0, durDays: 15, locked: true, capType: 'Development' },
      { num: 2, feature: 'begun', phaseId: 'p1', startDay: 5, durDays: 3, capType: 'Development' }
    ], [{ name: 'Solo', capType: 'Development' }], null, phPin30b);
    var oPin = { today: 10, snap: { feature: m, story: 'day' } };
    var rPin = RM.autoPhase(sPin, 'p1', oPin);
    eq(RM.itemByNum(rPin.state, 2).startDay, 20, 'begun work pushed past today lands at its phase floor (' + m + ' snap)');
    eq(RM.autoPhase(rPin.state, 'p1', oPin).changed, 0, 'and one click settles it (' + m + ' snap)');
  });
  // (4) a holiday on an in-flight start is not a reason to move it
  ['day', 'sprint'].forEach(function (m) {
    // two heads: the holiday week's 0.8 × 2 supply still covers it, so only the day itself is in question
    var sHol = autoState([{ num: 1, feature: 'begun', phaseId: 'p1', startDay: 5, durDays: 3, capType: 'Development' }],
      [{ name: 'X', capType: 'Development' }, { name: 'Y', capType: 'Development' }], { holidays: ['2026-08-03'] });
    var hb = RM.itemByNum(RM.autoTimeline(sHol, { phaseIds: ['p1'], today: 6, snap: { feature: m, story: 'day' } }).state, 1);
    eq([hb.startDay, hb.durDays], [5, 3], 'in-flight work starting on a holiday keeps its start and length (' + m + ' snap)');
  });
  // a dependency that now ends later pushes in-flight work, minimally while it is still before today
  var sDepIn = autoState([
    { num: 1, feature: 'dep', phaseId: 'p1', startDay: 0, durDays: 4, locked: true, capType: '' },
    { num: 2, feature: 'begun', phaseId: 'p1', startDay: 2, durDays: 3, capType: '', deps: [1] }
  ], [{ name: 'X', capType: 'Development' }, { name: 'Y', capType: 'Development' }]);
  var dIn = RM.itemByNum(RM.autoTimeline(sDepIn, { phaseIds: ['p1'], today: 10, snap: { feature: 'week', story: 'day' } }).state, 2);
  eq([dIn.startDay, dIn.durDays], [4, 3], 'pushed but still before today: day grid, same length');
  var sPl = snapSt();
  var stId = sPl.items[0].stories[0].id;
  var rPl = RM.placeUnit(sPl, sPl.items[0].id, stId, { today: 2, snap: { story: 'week', feature: 'week' } });
  var stP = RM.itemByNum(rPl.state, 1).stories[0];
  eq(stP.startDay, 5, 'Place at earliest slot snaps too');
  eq(stP.durDays, 5, 'and rounds the duration up');
}
