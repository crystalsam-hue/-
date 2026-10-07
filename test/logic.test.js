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
    { centerCode: 'C01', centerName: '서울', managerName: '김', total: 3, present: 2, late: 1, working: 1, away: 1, left: 0, vacation: 0, absent: 1 });
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

/* ───── 직급 · 휴가 ───── */

const RULE = { startMinutes: 600, graceMinutes: 0, endMinutes: 1140, halfAmStartMinutes: 900, halfPmEndMinutes: 840,
  rankOrder: ['센터장', '팀장', '사원'] };

test('출근 10:00 기준 지각', () => {
  assert.ok(!L.isLate(600, RULE.startMinutes, RULE.graceMinutes)); // 10:00 정상
  assert.ok(L.isLate(601, RULE.startMinutes, RULE.graceMinutes));  // 10:01 지각
});

test('휴가자는 미출근에서 제외, 오전반차는 지각 기준 변경, 직급순 정렬', () => {
  const centers = [{ code: 'C01', name: '서울', managerName: '김' }];
  const employees = [
    { id: 'E9', name: '사원1', rank: '사원', centerCode: 'C01' },
    { id: 'E2', name: '팀장1', rank: '팀장', centerCode: 'C01' },
    { id: 'E1', name: '센터장', rank: '센터장', centerCode: 'C01' },
    { id: 'E3', name: '오후반차', rank: '사원', centerCode: 'C01' },
  ];
  const logs = [
    { employeeId: 'E2', action: A.IN, minutes: 870 }, // 14:30 출근, 오전반차 → 지각 아님
  ];
  const leaves = { E1: '연차', E2: '오전반차', E3: '오후반차' };
  const r = L.computeDailyStatus(centers, employees, logs, RULE, leaves);
  assert.deepStrictEqual(r.rows.map(x => x.employeeId), ['E1', 'E2', 'E3', 'E9']);
  const by = Object.fromEntries(r.rows.map(x => [x.employeeId, x]));
  assert.strictEqual(by.E1.state, S.VACATION);
  assert.strictEqual(by.E2.state, S.WORKING);
  assert.strictEqual(by.E2.late, false);
  assert.strictEqual(by.E3.state, S.NONE); // 오후반차는 오전 출근 필요
  assert.strictEqual(r.summaries[0].vacation, 3);
  assert.strictEqual(r.summaries[0].absent, 2);
});

test('날짜 정규화/계산', () => {
  assert.strictEqual(L.normalizeYmd('2024.3.1'), '2024-03-01');
  assert.strictEqual(L.normalizeYmd('2024/03/01'), '2024-03-01');
  assert.strictEqual(L.normalizeYmd('2023-02-29'), null);
  assert.strictEqual(L.addMonths('2026-01-31', 1), '2026-02-28');
  assert.strictEqual(L.addYears('2024-02-29', 1), '2025-02-28');
  assert.strictEqual(L.weekdayOfYmd('2026-10-07'), 3); // 수
});

test('휴가 일수: 주말·공휴일 제외, 반차 0.5', () => {
  const wd = '월,화,수,목,금';
  // 2026-10-05(월) ~ 10-11(일), 10-09 한글날
  assert.strictEqual(L.countLeaveDays('연차', '2026-10-05', '2026-10-11', wd, ['2026-10-09']), 4);
  assert.strictEqual(L.countLeaveDays('오전반차', '2026-10-07', '2026-10-07', wd, []), 0.5);
  assert.strictEqual(L.countLeaveDays('연차', '2026-10-10', '2026-10-11', wd, []), 0);
  assert.throws(() => L.countLeaveDays('오후반차', '2026-10-07', '2026-10-08', wd, []));
});

test('연차 부여 (근로기준법 입사일 기준)', () => {
  let e = L.leaveEntitlement('2026-03-15', '2026-10-07', '');
  assert.strictEqual(e.days, 6); // 1년 미만: 만근 개월수
  assert.strictEqual(e.periodStart, '2026-03-15');
  assert.strictEqual(e.periodEnd, '2027-03-14');
  assert.strictEqual(L.leaveEntitlement('2025-01-01', '2026-10-07', '').days, 15);  // 1년
  assert.strictEqual(L.leaveEntitlement('2023-01-01', '2026-10-07', '').days, 16);  // 3년
  assert.strictEqual(L.leaveEntitlement('2021-01-01', '2026-10-07', '').days, 17);  // 5년
  assert.strictEqual(L.leaveEntitlement('1990-01-01', '2026-10-07', '').days, 25);  // 상한
  e = L.leaveEntitlement('2023-05-10', '2026-10-07', '');
  assert.strictEqual(e.periodStart, '2026-05-10');
  assert.strictEqual(e.periodEnd, '2027-05-09');
  assert.strictEqual(L.leaveEntitlement('2023-05-10', '2026-10-07', '20').days, 20); // 수동
  assert.strictEqual(L.leaveEntitlement(null, '2026-10-07', '').days, null);
  assert.strictEqual(L.leaveEntitlement(null, '2026-10-07', '15').periodStart, '2026-01-01');
});

test('연차 잔여 · 신청 검증', () => {
  const ent = L.leaveEntitlement('2025-01-01', '2026-10-07', ''); // 15일, 2026-01-01~
  const requests = [
    { employeeId: 'E1', type: '연차', start: '2026-02-02', end: '2026-02-06', days: 5, status: '승인' },
    { employeeId: 'E1', type: '오전반차', start: '2026-10-08', end: '2026-10-08', days: 0.5, status: '신청' },
    { employeeId: 'E1', type: '병가', start: '2026-03-02', end: '2026-03-03', days: 2, status: '승인' }, // 미차감
    { employeeId: 'E1', type: '연차', start: '2025-12-01', end: '2025-12-01', days: 1, status: '승인' }, // 이전 기간
    { employeeId: 'E1', type: '연차', start: '2026-04-01', end: '2026-04-01', days: 1, status: '반려' },
  ];
  const bal = L.leaveBalance(requests, 'E1', ent);
  assert.deepStrictEqual(bal, { granted: 15, used: 5, pending: 0.5, remaining: 9.5 });

  const v = (req) => L.validateLeaveRequest({ employeeId: 'E1', ...req }, { requests, balance: bal });
  assert.strictEqual(v({ type: '연차', start: '2026-10-12', end: '2026-10-16', days: 5 }), null);
  assert.match(v({ type: '연차', start: '2026-10-12', end: '2026-10-30', days: 15 }), /잔여 연차가 부족/);
  assert.match(v({ type: '연차', start: '2026-10-08', end: '2026-10-08', days: 1 }), /겹칩니다/);
  assert.strictEqual(v({ type: '오후반차', start: '2026-10-08', end: '2026-10-08', days: 0.5 }), null); // 오전+오후 허용
  assert.strictEqual(v({ type: '병가', start: '2026-10-12', end: '2026-10-30', days: 15 }), null); // 미차감 종류
  assert.match(v({ type: '연차', start: '2026-10-10', end: '2026-10-11', days: 0 }), /근무일이 없습니다/);
  assert.match(v({ type: '연차', start: '2026-10-12', end: '2026-10-11', days: 0 }), /종료일/);
});

test('특정일 승인 휴가', () => {
  const requests = [
    { employeeId: 'E1', type: '연차', start: '2026-10-07', end: '2026-10-08', status: '승인' },
    { employeeId: 'E2', type: '연차', start: '2026-10-07', end: '2026-10-07', status: '신청' },
    { employeeId: 'E3', type: '오전반차', start: '2026-10-07', end: '2026-10-07', status: '승인' },
    { employeeId: 'E3', type: '오후반차', start: '2026-10-07', end: '2026-10-07', status: '승인' },
  ];
  assert.deepStrictEqual(L.leavesOn(requests, '2026-10-07'), { E1: '연차', E3: '연차' });
});
