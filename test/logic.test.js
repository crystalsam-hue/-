// npm test  (node --test test/logic.test.js)
const test = require('node:test');
const assert = require('node:assert');
const L = require('../src/Logic.gs');

const { ACTIONS: A, STATES: S } = L;

test('시각 파싱/포맷', () => {
  assert.strictEqual(L.parseHm('09:30'), 570);
  assert.strictEqual(L.formatHm(570), '09:30');
  assert.deepStrictEqual(L.parseHmList('18:30, 09:30 ,13:30'), [570, 810, 1110]);
  assert.throws(() => L.parseHm('9시'));
});

test('근무요일', () => {
  assert.ok(L.isWorkday(1, '월,화,수,목,금'));
  assert.ok(!L.isWorkday(0, '월,화,수,목,금'));
  assert.ok(L.isWorkday(6, '월,화,수,목,금,토'));
});

test('상태 전이 검증', () => {
  assert.strictEqual(L.validateAction(S.NONE, A.IN), null);
  assert.ok(L.validateAction(S.NONE, A.OUT));
  assert.ok(L.validateAction(S.NONE, A.LEAVE));
  assert.ok(L.validateAction(S.WORKING, A.IN));
  assert.strictEqual(L.validateAction(S.WORKING, A.OUT), null);
  assert.ok(L.validateAction(S.WORKING, A.BACK));
  assert.strictEqual(L.validateAction(S.AWAY, A.BACK), null);
  assert.strictEqual(L.validateAction(S.AWAY, A.LEAVE), null);
  assert.ok(L.validateAction(S.LEFT, A.IN));
  assert.ok(L.validateAction(S.LEFT, A.LEAVE));
});

test('직원 하루 계산', () => {
  const d = L.computeEmployeeDay([
    { action: A.IN, minutes: 545 },
    { action: A.OUT, minutes: 700 },
    { action: A.BACK, minutes: 760 },
    { action: A.OUT, minutes: 900 },
  ]);
  assert.strictEqual(d.state, S.AWAY);
  assert.strictEqual(d.inAt, 545);
  assert.strictEqual(d.outCount, 2);
  assert.strictEqual(d.outAt, 900);
});

test('지각 판정', () => {
  assert.ok(!L.isLate(550, 540, 10));
  assert.ok(L.isLate(551, 540, 10));
  assert.ok(!L.isLate(null, 540, 10));
});

test('전체 현황 집계', () => {
  const centers = [{ code: 'C01', name: '서울', managerName: '김' }, { code: 'C02', name: '부산', managerName: '이' }];
  const employees = [
    { id: 'E1', name: 'a', centerCode: 'C01' },
    { id: 'E2', name: 'b', centerCode: 'C01' },
    { id: 'E3', name: 'c', centerCode: 'C01' },
    { id: 'E4', name: 'd', centerCode: 'C02' },
  ];
  const logs = [
    { employeeId: 'E1', action: A.IN, minutes: 530 },
    { employeeId: 'E2', action: A.IN, minutes: 600 },
    { employeeId: 'E1', action: A.OUT, minutes: 700 },
    { employeeId: 'E4', action: A.IN, minutes: 535 },
    { employeeId: 'E4', action: A.LEAVE, minutes: 1090 },
  ];
  const r = L.computeDailyStatus(centers, employees, logs, { startMinutes: 540, graceMinutes: 10 });
  assert.deepStrictEqual(
    { ...r.summaries[0] },
    { centerCode: 'C01', centerName: '서울', managerName: '김', total: 3, present: 2, late: 1, working: 1, away: 1, left: 0, absent: 1 });
  assert.strictEqual(r.summaries[1].left, 1);
  assert.strictEqual(r.total.total, 4);
  assert.strictEqual(r.total.absent, 1);
  assert.strictEqual(r.total.present, 3);
});

test('보고 시각 도래', () => {
  const slots = [570, 810, 1110];
  assert.strictEqual(L.dueReportSlot(560, slots, []), null);
  assert.strictEqual(L.dueReportSlot(575, slots, []), 570);
  assert.strictEqual(L.dueReportSlot(575, slots, [570]), null);
  assert.strictEqual(L.dueReportSlot(900, slots, []), 810);
});

test('거리 계산', () => {
  // 서울시청 ↔ 광화문 약 1km 내외
  const d = L.haversineMeters(37.5663, 126.9779, 37.5759, 126.9768);
  assert.ok(d > 900 && d < 1200, String(d));
});
