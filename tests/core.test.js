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

// ------------------------------------------------------------- normalize
section('normalizeState');
var sN = mkState([
  { feature: 'A' },
  { num: 7, feature: 'B', headcount: 0, phaseId: 'nope' }
]);
eq(sN.items[0].headcount, 1, 'default headcount 1');
ok(sN.items[0].num != null && sN.items[0].num !== 7, 'auto num assigned, no collision');
eq(sN.items[1].headcount, 1, 'headcount floor 1');
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
  { num: 1, feature: 'base', startDay: 0, durDays: 10, size: 'L' },
  { num: 2, feature: 'early bird', deps: [1], startDay: 5, durDays: 5, size: 'M' },
  { num: 3, feature: 'no size', startDay: 0, durDays: 5 },
  { num: 4, feature: 'ghost dep', deps: [42], startDay: 20, durDays: 5, size: 'M' },
  { num: 5, feature: 'big ask', startDay: 0, durDays: 25, size: 'XL', teamType: 'Data' }
], { team: [{ name: 'X', type: 'Development' }, { name: 'Y', type: 'Data' }] });
var v = RM.validate(sV);
function codes(state, i) { return (v.byItem[state.items[i].id] || []).map(function (x) { return x.code; }); }
ok(codes(sV, 1).indexOf('DEP_ORDER') !== -1, 'DEP_ORDER: starts before dep ends');
ok(codes(sV, 2).indexOf('NO_SIZE') === -1, 'no NO_SIZE nag — sizing is optional');
ok(codes(sV, 3).indexOf('UNKNOWN_DEP') !== -1, 'UNKNOWN_DEP flagged');
ok(v.global.some(function (g) { return g.code === 'OVER_CAP'; }), 'OVER_CAP: weekly WIP over what the team can focus on');
ok(!v.global.some(function (g) { return /Data/.test(g.msg); }), 'capacity messages are role-agnostic now');
ok(v.counts.warn > 0, 'counts aggregated');

var vClean = RM.validate(mkState([{ num: 1, feature: 'solo', startDay: 0, durDays: 5, size: 'M' }]));
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
// WIP model: an active item costs focus by its working days / 10 (M = 1,
// clamped 0.3..2) — 'demand' is size-weighted concurrent work, 'cap' the
// fractional people available
var sCap = mkState([
  { num: 1, feature: 'big', startDay: 0, durDays: 20, size: 'XL' },   // weight 2
  { num: 2, feature: 'mid', startDay: 5, durDays: 10, size: 'M' }     // weight 1
], { team: [{ name: 'X', type: 'Development' }, { name: 'Y', type: 'Development' }] });
var cap = RM.capacity(sCap);
eq(cap.weeks[0].demand, 2, 'week 0 WIP = 2 focus units (one XL)');
eq(cap.weeks[1].demand, 3, 'week 1 WIP = 3 (overlap)');
eq(cap.weeks[0].cap, 2, 'availability = fractional people on the roster');
ok(cap.weeks[1].over, 'week 1: more WIP than the 2-person team can focus on');
ok(!cap.weeks[0].over, 'week 0 fits');
eq(RM.wipWeight(sCap, { size: 'XS' }), 0.3, 'tiny items still register (clamped floor)');
eq(RM.wipWeight(sCap, { size: 'XL' }), 2, 'XL clamps at 2 focus units');

// ------------------------------------------------------------- auto-schedule
section('autoSchedule');
// dep chain: 2 after 1, capacity 1 person forces serialization of parallel items
var sA = mkState([
  { num: 1, feature: 'first', phaseId: 'p1', size: 'M' },
  { num: 2, feature: 'second', phaseId: 'p1', size: 'M', deps: [1] },
  { num: 3, feature: 'third', phaseId: 'p1', size: 'M' },
  { num: 9, feature: 'parked', phaseId: 'p2', size: 'M' }
], { team: [{ name: 'Solo', type: 'Development' }] });
var rA = RM.autoSchedule(sA);
var A = {};
rA.state.items.forEach(function (it) { A[it.num] = it; });
eq(A[1].startDay, 0, 'item 1 starts at 0');
ok(A[2].startDay >= A[1].startDay + A[1].durDays, 'item 2 starts after dep 1 ends');
// capacity 1: items 1,2,3 cannot overlap at week granularity
var spans = [A[1], A[2], A[3]].map(function (it) { return [it.startDay, it.startDay + it.durDays]; });
var overlapWeeks = false;
for (var i = 0; i < 3; i++) for (var j = i + 1; j < 3; j++) {
  var wA = [Math.floor(spans[i][0] / 5), Math.floor((spans[i][1] - 1) / 5)];
  var wB = [Math.floor(spans[j][0] / 5), Math.floor((spans[j][1] - 1) / 5)];
  if (wA[0] <= wB[1] && wB[0] <= wA[1]) overlapWeeks = true;
}
ok(!overlapWeeks, '1-person roster serializes the three items');
ok(A[9].startDay == null, 'bucket-phase item stays unscheduled');

// locked stays put and is scheduled around
var sL = mkState([
  { num: 1, feature: 'locked rock', phaseId: 'p1', size: 'M', startDay: 5, durDays: 5, locked: true },
  { num: 2, feature: 'flows around', phaseId: 'p1', size: 'M' }
], { team: [{ name: 'Solo', type: 'Development' }] });
var rL = RM.autoSchedule(sL);
var L = {};
rL.state.items.forEach(function (it) { L[it.num] = it; });
eq(L[1].startDay, 5, 'locked item did not move');
var lockWeeks = [1];
var w2 = [Math.floor(L[2].startDay / 5), Math.floor((L[2].startDay + L[2].durDays - 1) / 5)];
ok(!(w2[0] <= 1 && 1 <= w2[1]), 'unlocked item avoids the locked week');

// blackout stretch: force start before blackout, size L must stretch
var sB = mkState([
  { num: 1, feature: 'pre', phaseId: 'p1', size: 'M', startDay: 70, durDays: 5, locked: true },
  { num: 2, feature: 'holiday spanner', phaseId: 'p1', size: 'L', deps: [1] }
]);
var rB = RM.autoSchedule(sB);
var B = {};
rB.state.items.forEach(function (it) { B[it.num] = it; });
eq(B[2].startDay, 75, 'starts right after dep at day 75 (week 15)');
eq(B[2].durDays, 30, 'L(20d) stretches to 30 slots across two blackout weeks');
eq(RM.workInSpan(rB.state.meta, B[2].startDay, B[2].durDays), 20, 'net work still 20 days');

// cycles do not hang
var sCy = mkState([
  { num: 1, feature: 'a', deps: [2], size: 'M' },
  { num: 2, feature: 'b', deps: [1], size: 'M' }
]);
var rCy = RM.autoSchedule(sCy);
ok(rCy.state.items.every(function (it) { return it.startDay != null; }), 'cycle members still get scheduled');
ok(rCy.notes.length > 0, 'cycle break noted');

// snapEarliest
var sS = mkState([
  { num: 1, feature: 'base', startDay: 0, durDays: 10, size: 'L' },
  { num: 2, feature: 'snapme', deps: [1], startDay: 0, durDays: 5, size: 'M' }
]);
var rS = RM.snapEarliest(sS, sS.items[1].id);
eq(RM.itemByNum(rS.state, 2).startDay, 10, 'snapEarliest lands right after dep');

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
// snapEarliest: infeasible item (more focus than the roster has) stays put
var sInf = mkState([
  { num: 1, feature: 'crowd', startDay: 10, durDays: 15, size: 'L' }
], { team: [{ name: 'X', type: 'Development', capacity: 0.5 }] });
var rInf = RM.snapEarliest(sInf, sInf.items[0].id);
eq(rInf.changed, 0, 'infeasible snap changes nothing');
eq(RM.itemByNum(rInf.state, 1).startDay, 10, 'infeasible snap keeps the old start');
ok(!!rInf.note, 'infeasible snap explains itself');
// snapEarliest: feasible item beyond horizon extends the timeline instead of vanishing
var packed = [];
for (var pk = 0; pk < 48; pk++) {
  packed.push({ num: pk + 1, feature: 'wk' + pk, startDay: pk * 5, durDays: 5, size: 'M', locked: true });
}
packed.push({ num: 99, feature: 'late arrival', size: 'M' });
var sPack = mkState(packed, { team: [{ name: 'Solo', type: 'Development' }] });
sPack.meta.blackoutWeeks = [];
var rPack = RM.snapEarliest(sPack, sPack.items[48].id);
var late = RM.itemByNum(rPack.state, 99);
eq(late.startDay, 240, 'feasible overflow schedules right after the packed horizon');
ok(rPack.state.meta.numWeeks >= 49, 'timeline extended to fit (' + rPack.state.meta.numWeeks + 'w)');
// autoSchedule: infeasible item gets a note and no endless walk
var sInf2 = mkState([
  { num: 1, feature: 'crowd', size: 'XL' }
], { team: [{ name: 'X', type: 'Development', capacity: 0.5 }] });
var rInf2 = RM.autoSchedule(sInf2);
ok(rInf2.notes.some(function (nn) { return nn.indexOf('left unscheduled') !== -1; }),
  'autoSchedule notes the infeasible item');
eq(RM.itemByNum(rInf2.state, 1).startDay, null, 'infeasible item is left unscheduled, never overallocated');

// hours-aware: a part-time roster (20h = 0.5 people) can't absorb an M's focus
var sInf3 = mkState([
  { num: 1, feature: 'full-time ask', size: 'M' }
], { team: [{ name: 'Half', type: 'Development', weekHours: (function () {
  var wh = {};
  for (var w = 0; w < 60; w++) wh[RM.fmtISO(RM.weekStartDate(META, w))] = 20;
  return wh;
})() }] });
var rInf3 = RM.autoSchedule(sInf3);
eq(RM.itemByNum(rInf3.state, 1).startDay, null, 'part-time-only roster leaves a full-focus item unscheduled');
ok(rInf3.notes.length > 0, 'shortfall is noted');

// a sufficient roster schedules without tripping the over-capacity check
var sOk = mkState([
  { num: 1, feature: 'a', size: 'M' },
  { num: 2, feature: 'b', size: 'M' }
], { team: [{ name: 'X', type: 'Development' }] });
var rOk = RM.autoSchedule(sOk);
ok(rOk.state.items.every(function (it) { return it.startDay != null; }), 'feasible items all scheduled');
ok(!RM.validate(rOk.state).global.some(function (g) { return g.code === 'OVER_CAP'; }),
  'auto-schedule result has no over-capacity week');

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
ok(RM.autoSchedule(sOff).state.items[0].startDay != null, 'capacity off: item schedules despite a tiny roster');
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

// auto-schedule: dependents start right after the work span
var sRk = mkState([
  { num: 1, feature: 'A', size: 'M', risk: 'S' },
  { num: 2, feature: 'B', size: 'S', deps: [1] }
]);
var rRk = RM.autoSchedule(sRk);
var K = {};
rRk.state.items.forEach(function (it) { K[it.num] = it; });
eq(K[1].durDays, 10, 'A work = M = 2w');
eq(K[1].riskDays, 0, 'no buffer appended despite risk S');
eq(K[2].startDay, 10, 'B starts right after A\'s work');

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
  { num: 1, feature: 'chunky work', startDay: 0, durDays: 15 }
], {
  team: [
    { id: 'm1', name: 'Ada', type: 'Development', offWeeks: ['2026-07-27'] },
    { id: 'm2', name: 'Grace', type: 'Development' }
  ]
});
var capOff = RM.capacity(sOff);
eq(capOff.weeks[0].cap, 1, 'off member lowers week 0 capacity');
ok(capOff.weeks[0].over, '1.5 focus units vs 1 person available is over');
eq(capOff.weeks[1].cap, 2, 'week 1 back to full roster');
var snapOff = RM.snapEarliest(sOff, sOff.items[0].id);
eq(snapOff.state.items[0].startDay, 5, 'snap skips the short-handed week');

// ------------------------------------------------------------- hours model
section('hours');
var sHr = mkState([
  { num: 1, feature: 'w', startDay: 0, durDays: 10, teamType: 'Development' }
], {
  team: [
    { id: 'h1', name: 'Ada', type: 'Development', weekHours: { '2026-07-27': 20 } },
    { id: 'h2', name: 'Grace', type: 'Data' }
  ]
});
eq(RM.memberHoursForWeek(sHr.meta, sHr.team[0], 0), 20, 'explicit week hours read back');
eq(RM.memberHoursForWeek(sHr.meta, sHr.team[0], 1), 40, 'unlisted weeks default to 40');
eq(RM.availForWeek(sHr, 0).total, 1.5, 'total people-equivalents (20h counts as 0.5)');
ok(!RM.capacity(sHr).weeks[0].over, 'capacity is role-agnostic: 1 focus unit vs 1.5 people is fine');
var sHrSolo = mkState([
  { num: 1, feature: 'w', startDay: 0, durDays: 10, teamType: 'Development' }
], { team: [{ id: 'h1', name: 'Ada', type: 'Development', weekHours: { '2026-07-27': 20 } }] });
ok(RM.capacity(sHrSolo).weeks[0].over, 'a 1-focus item vs 0.5 available people is over');
// legacy offWeeks migrate to zero-hour weeks
var sOffMig = RM.normalizeState({
  meta: { timelineStart: '2026-07-27', numWeeks: 8 }, phases: [{ id: 'p' }], items: [],
  team: [{ name: 'Ada', type: 'Development', offWeeks: ['2026-08-03'] }]
});
eq(sOffMig.team[0].weekHours['2026-08-03'], 0, 'offWeeks -> 0-hour week');
ok(RM.memberOffWeek(sOffMig.meta, sOffMig.team[0], 1), 'memberOffWeek still answers via hours');
// items default to 1 × Development
eq(mkState([{ num: 1, feature: 'x' }]).items[0].teamType, 'Development', 'default work type');

// capacity factor: PE at 40h scales availability
var sCf = mkState([], { team: [{ name: 'Half', type: 'Development', capacity: 0.5 }] });
eq(RM.availForWeek(sCf, 0).total, 0.5, 'capacity 0.5 at 40h = half a head');
eq(sCf.team[0].capacity, 0.5, 'capacity survives normalize');
eq(mkState([], { team: [{ name: 'X', type: 'Development' }] }).team[0].capacity, 1, 'capacity defaults to 1');
eq(mkState([], { team: [{ name: 'Z', type: 'Development', capacity: 0 }] }).team[0].capacity, 0, 'capacity 0 is allowed');
eq(RM.availForWeek(mkState([], { team: [{ name: 'Z', type: 'Development', capacity: 0 }] }), 0).total, 0,
  'a zero-capacity role contributes nothing');

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
  ['description'], 'new documents start with Description only');
// legacy docs without a saved column list surface built-ins that hold content
var sScLegacy = mkState([{ num: 1, feature: 'a', enables: 'x', notes: 'y' }]);
eq(sScLegacy.meta.scopeCols.map(function (c) { return c.key; }),
  ['description', 'enables', 'notes'], 'legacy content infers its columns');
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
eq(RM.normalizeState(sSc).items[0].custom.x9, 'keep', 'custom values survive normalize');

// ------------------------------------------------------------- milestones
section('milestones');
var sMs = mkState([
  { num: 1, feature: 'work', startDay: 0, durDays: 10 },
  { num: 2, feature: 'launch', milestone: true, startDay: 10, durDays: 0, deps: [1] },
  { num: 3, feature: 'after', startDay: 10, durDays: 5, deps: [2] }
]);
eq(sMs.items[1].milestone, true, 'milestone flag survives normalize');
eq(sMs.items[1].durDays, 0, 'milestone keeps zero duration');
eq(RM.itemSpan(sMs.items[1]), 0, 'milestone span is zero');
eq(RM.itemEnd(sMs.items[1]), 10, 'milestone end equals its day');
var vMs = RM.validate(sMs);
var msWarns = (vMs.byItem[sMs.items[1].id] || []).map(function (v) { return v.code; });
eq(msWarns.indexOf('NO_SIZE'), -1, 'no size warnings at all');
var depWarns = (vMs.byItem[sMs.items[2].id] || []).map(function (v) { return v.code; });
eq(depWarns.indexOf('DEP_ORDER'), -1, 'dependent may start on the milestone day');
var msAuto = RM.autoSchedule(sMs);
eq(RM.itemById(msAuto.state, sMs.items[1].id).startDay, 10, 'auto-schedule leaves milestones fixed');
var msCap = RM.capacity(sMs);
eq(msCap.weeks[2].demand, 0.5, 'milestone consumes no capacity (only the 5-day item registers)');
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
sWw.team = [{ id: 't1', name: 'A', type: 'Development', capacity: 1, weekHours: {} }];
eq(RM.availForWeek(sWw, 0).total, 1, '32 h at a 32 h full-time week = 1 person');
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
    return sl.shapes.filter(function (s) { return s.type === 'rect' && s.color === '1A1F26'; })[0].h;
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
            finish();
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
eq(RM.DEFAULT_TEAM_TYPES[0], 'Software Engineer', 'default roles use real role names');

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
  eq(a.stories[0].durDays, 5, 'a story without a span gets one week');
  RM.moveStoryToSprint(sSp, 'a', 'a2', null, null);
  ok(a.stories[1].id === 'a2' && a.stories[1].startDay == null && a.stories[1].durDays == null,
    'unscheduling a story clears its timeline and moves it last');
}

var RMJira = require('../js/export-jira.js');
var sJc = mkState([
  { num: 1, feature: 'Login page', epic: 'Login', workstream: 'Product', size: 'M',
    startDay: 0, durDays: 5, deadline: '2026-09-04', jiraKey: 'HW-12',
    description: '<p>Hi <b>there</b></p>', enables: 'Checkout', notes: '',
    stories: [{ title: 's1', done: true }, { title: 's2', jiraKey: 'HW-13' }] },
  { num: 2, feature: 'Search, "fast"', deps: [1], phaseId: 'p2' },
  { num: 3, feature: 'Orphan', deps: [2] }
], { epicJira: { Login: 'HW-1' } });
eq(RMJira.fileName(sJc), 'T-jira.csv', 'jira csv filename');
var jr = RMJira.rows(sJc, { features: true, stories: false });
eq(jr.length, 3, 'features only: one row per feature');
var r1 = jr[0];
eq(r1['Summary'], 'Login page', 'summary is the feature name');
eq(r1['Issue Type'], 'Story', 'feature issue type defaults to Story');
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

var jrs = RMJira.rows(sJc, { features: true, stories: true, featureType: 'Task', storyType: 'Sub-task' });
eq(jrs.length, 5, 'features and stories: a row per story too');
eq(jrs[0]['Issue Type'], 'Task', 'custom feature issue type');
ok(jrs[0]['Description'].indexOf('[x]') === -1, 'checklist omitted when stories are rows');
eq(jrs[1]['Summary'], 's1', 'story row summary');
eq(jrs[1]['Issue Type'], 'Sub-task', 'story issue type');
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
