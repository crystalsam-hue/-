/**
 * 순수 로직 (Google 서비스에 의존하지 않음 → Node 로 테스트 가능)
 */

var ACTIONS = {
  IN: '출근',
  OUT: '외출',
  BACK: '복귀',
  LEAVE: '퇴근'
};

var STATES = {
  NONE: '미출근',
  WORKING: '근무중',
  AWAY: '외출중',
  LEFT: '퇴근',
  VACATION: '휴가'
};

var LEAVE_TYPES = ['연차', '오전반차', '오후반차', '병가', '경조사', '공가', '기타'];
/** 연차에서 차감되는 종류 */
var LEAVE_DEDUCT = ['연차', '오전반차', '오후반차'];
var LEAVE_STATUS = { REQ: '신청', OK: '승인', NO: '반려', CANCEL: '취소' };

function isHalfDay(type) {
  return type === '오전반차' || type === '오후반차';
}

/* ───────────── 시각 ───────────── */

/** "09:30" → 570 (분) */
function parseHm(text) {
  var m = String(text).trim().match(/^(\d{1,2}):(\d{2})$/);
  if (!m) throw new Error('시각 형식 오류(HH:MM): ' + text);
  return Number(m[1]) * 60 + Number(m[2]);
}

/** 570 → "09:30" */
function formatHm(minutes) {
  var h = Math.floor(minutes / 60);
  var m = minutes % 60;
  return (h < 10 ? '0' : '') + h + ':' + (m < 10 ? '0' : '') + m;
}

/** "09:30, 14:00 ,19:00" → [570, 840, 1140] (정렬) */
function parseHmList(text) {
  return String(text || '')
    .split(/[,\s]+/)
    .filter(function (s) { return s; })
    .map(parseHm)
    .sort(function (a, b) { return a - b; });
}

/* ───────────── 날짜 (yyyy-MM-dd 문자열) ───────────── */

var WEEKDAY_NAMES = ['일', '월', '화', '수', '목', '금', '토'];

function pad2_(n) { return (n < 10 ? '0' : '') + n; }

/** "2024.3.1", "2024/03/01", "2024-03-01" → "2024-03-01". 잘못된 값이면 null */
function normalizeYmd(text) {
  var m = String(text || '').trim().match(/^(\d{4})\D+(\d{1,2})\D+(\d{1,2})\D*$/);
  if (!m) return null;
  var y = Number(m[1]), mo = Number(m[2]), d = Number(m[3]);
  var t = new Date(Date.UTC(y, mo - 1, d));
  if (t.getUTCFullYear() !== y || t.getUTCMonth() !== mo - 1 || t.getUTCDate() !== d) return null;
  return y + '-' + pad2_(mo) + '-' + pad2_(d);
}

function ymdParts_(s) {
  var p = s.split('-');
  return { y: Number(p[0]), m: Number(p[1]), d: Number(p[2]) };
}

function utcToYmd_(t) {
  return t.getUTCFullYear() + '-' + pad2_(t.getUTCMonth() + 1) + '-' + pad2_(t.getUTCDate());
}

function addDays(s, n) {
  var p = ymdParts_(s);
  return utcToYmd_(new Date(Date.UTC(p.y, p.m - 1, p.d + n)));
}

/** 월 더하기 (말일 보정: 1/31 + 1개월 = 2/28) */
function addMonths(s, n) {
  var p = ymdParts_(s);
  var first = new Date(Date.UTC(p.y, p.m - 1 + n, 1));
  var lastDay = new Date(Date.UTC(first.getUTCFullYear(), first.getUTCMonth() + 1, 0)).getUTCDate();
  first.setUTCDate(Math.min(p.d, lastDay));
  return utcToYmd_(first);
}

function addYears(s, n) {
  return addMonths(s, n * 12);
}

function weekdayOfYmd(s) {
  var p = ymdParts_(s);
  return new Date(Date.UTC(p.y, p.m - 1, p.d)).getUTCDay();
}

/** 근무요일 문자열("월,화,수,목,금")에 해당 요일(0=일)이 포함되는지 */
function isWorkday(weekdayIndex, workdaysText) {
  var list = String(workdaysText || '월,화,수,목,금').split(/[,\s]+/);
  return list.indexOf(WEEKDAY_NAMES[weekdayIndex]) !== -1;
}

/** 근무일인지 (근무요일이면서 공휴일이 아님) */
function isWorkdate(ymd, workdaysText, holidays) {
  return isWorkday(weekdayOfYmd(ymd), workdaysText) && (holidays || []).indexOf(ymd) === -1;
}

/* ───────────── 위치 ───────────── */

/** 두 좌표 간 거리(m) */
function haversineMeters(lat1, lng1, lat2, lng2) {
  var R = 6371000;
  var toRad = function (d) { return d * Math.PI / 180; };
  var dLat = toRad(lat2 - lat1);
  var dLng = toRad(lng2 - lng1);
  var a = Math.sin(dLat / 2) * Math.sin(dLat / 2) +
    Math.cos(toRad(lat1)) * Math.cos(toRad(lat2)) *
    Math.sin(dLng / 2) * Math.sin(dLng / 2);
  return Math.round(2 * R * Math.asin(Math.sqrt(a)));
}

/* ───────────── 출퇴근 ───────────── */

/**
 * 오늘 기록(시간순)으로 직원의 현재 상태 계산
 * @param {Array<{action:string, minutes:number}>} events 해당 직원의 오늘 기록 (시간순)
 */
function computeEmployeeDay(events) {
  var day = {
    state: STATES.NONE,
    inAt: null,
    outAt: null,
    backAt: null,
    leaveAt: null,
    outCount: 0
  };
  events.forEach(function (e) {
    if (e.action === ACTIONS.IN) {
      if (day.inAt === null) day.inAt = e.minutes;
      day.state = STATES.WORKING;
    } else if (e.action === ACTIONS.OUT) {
      day.outAt = e.minutes;
      day.outCount++;
      day.state = STATES.AWAY;
    } else if (e.action === ACTIONS.BACK) {
      day.backAt = e.minutes;
      day.state = STATES.WORKING;
    } else if (e.action === ACTIONS.LEAVE) {
      day.leaveAt = e.minutes;
      day.state = STATES.LEFT;
    }
  });
  return day;
}

/** 현재 상태에서 요청한 구분이 가능한지. 불가하면 사유 문자열, 가능하면 null */
function validateAction(state, action) {
  switch (action) {
    case ACTIONS.IN:
      if (state === STATES.NONE) return null;
      return state === STATES.LEFT ? '이미 퇴근 처리되었습니다.' : '이미 출근 처리되었습니다.';
    case ACTIONS.OUT:
      if (state === STATES.WORKING) return null;
      if (state === STATES.NONE) return '출근 기록이 없습니다. 먼저 출근을 눌러주세요.';
      if (state === STATES.AWAY) return '이미 외출 중입니다. 복귀를 눌러주세요.';
      return '이미 퇴근 처리되었습니다.';
    case ACTIONS.BACK:
      if (state === STATES.AWAY) return null;
      return '외출 중이 아닙니다.';
    case ACTIONS.LEAVE:
      if (state === STATES.WORKING || state === STATES.AWAY) return null;
      if (state === STATES.NONE) return '출근 기록이 없습니다. 먼저 출근을 눌러주세요.';
      return '이미 퇴근 처리되었습니다.';
    default:
      return '알 수 없는 구분입니다: ' + action;
  }
}

/** 출근 시각이 지각인지 */
function isLate(inMinutes, startMinutes, graceMinutes) {
  return inMinutes !== null && inMinutes > startMinutes + (graceMinutes || 0);
}

/** 오늘 휴가 종류에 따른 출근/퇴근 기준시각 */
function workRuleFor(leaveType, rule) {
  return {
    startMinutes: leaveType === '오전반차' && rule.halfAmStartMinutes != null ? rule.halfAmStartMinutes : rule.startMinutes,
    endMinutes: leaveType === '오후반차' && rule.halfPmEndMinutes != null ? rule.halfPmEndMinutes : rule.endMinutes,
    graceMinutes: rule.graceMinutes || 0
  };
}

/** 직급 순서 정렬용 인덱스 (목록에 없으면 맨 뒤) */
function rankIndex(rank, rankOrder) {
  var i = (rankOrder || []).indexOf(rank);
  return i === -1 ? 999 : i;
}

/**
 * 전체 현황 계산
 * @param {Array} centers  [{code, name, managerName, ...}]
 * @param {Array} employees [{id, name, rank, centerCode, ...}] (재직자만)
 * @param {Array} logs [{employeeId, action, minutes}] 오늘 기록 (시간순)
 * @param {{startMinutes, graceMinutes, halfAmStartMinutes?, rankOrder?}} rule
 * @param {Object<string,string>} [leavesToday] 사번 → 오늘 승인된 휴가 종류
 */
function computeDailyStatus(centers, employees, logs, rule, leavesToday) {
  leavesToday = leavesToday || {};
  var byEmployee = {};
  logs.forEach(function (l) {
    (byEmployee[l.employeeId] = byEmployee[l.employeeId] || []).push(l);
  });

  var rows = employees.map(function (emp) {
    var day = computeEmployeeDay(byEmployee[emp.id] || []);
    var leaveType = leavesToday[emp.id] || '';
    var r = workRuleFor(leaveType, rule);
    var fullDayLeave = leaveType && !isHalfDay(leaveType);
    var state = day.state;
    // 휴가자는 출근 전이면 '휴가' (오후반차는 오전에 출근해야 하므로 제외)
    if (state === STATES.NONE && leaveType && leaveType !== '오후반차') state = STATES.VACATION;
    return {
      centerCode: emp.centerCode,
      employeeId: emp.id,
      name: emp.name,
      rank: emp.rank || '',
      state: state,
      leaveType: leaveType,
      inAt: day.inAt,
      outAt: day.outAt,
      backAt: day.backAt,
      leaveAt: day.leaveAt,
      outCount: day.outCount,
      late: !fullDayLeave && isLate(day.inAt, r.startMinutes, r.graceMinutes)
    };
  });

  rows.sort(function (a, b) {
    if (a.centerCode !== b.centerCode) return a.centerCode < b.centerCode ? -1 : 1;
    var ra = rankIndex(a.rank, rule.rankOrder), rb = rankIndex(b.rank, rule.rankOrder);
    if (ra !== rb) return ra - rb;
    return a.employeeId < b.employeeId ? -1 : a.employeeId > b.employeeId ? 1 : 0;
  });

  var KEYS = ['total', 'present', 'late', 'working', 'away', 'left', 'vacation', 'absent'];
  var summaries = centers.map(function (c) {
    var mine = rows.filter(function (r) { return r.centerCode === c.code; });
    var count = function (fn) { return mine.filter(fn).length; };
    return {
      centerCode: c.code,
      centerName: c.name,
      managerName: c.managerName,
      total: mine.length,
      present: count(function (r) { return r.inAt !== null; }),
      late: count(function (r) { return r.late; }),
      working: count(function (r) { return r.state === STATES.WORKING; }),
      away: count(function (r) { return r.state === STATES.AWAY; }),
      left: count(function (r) { return r.state === STATES.LEFT; }),
      vacation: count(function (r) { return !!r.leaveType; }),
      absent: count(function (r) { return r.state === STATES.NONE; })
    };
  });

  var total = { centerCode: '', centerName: '합계', managerName: '' };
  KEYS.forEach(function (k) {
    total[k] = summaries.reduce(function (acc, s) { return acc + s[k]; }, 0);
  });

  return { rows: rows, summaries: summaries, total: total };
}

/** 보고 시각 중 지금 보내야 할 것 (아직 안 보낸 것 중 가장 최근에 지난 시각). 없으면 null */
function dueReportSlot(nowMinutes, slots, sentSlots) {
  var due = null;
  slots.forEach(function (s) {
    if (s <= nowMinutes && sentSlots.indexOf(s) === -1) due = s;
  });
  return due;
}

/* ───────────── 휴가 · 연차 ───────────── */

/** 휴가 사용일수 (근무일만 계산, 반차는 0.5) */
function countLeaveDays(type, start, end, workdaysText, holidays) {
  if (isHalfDay(type)) {
    if (start !== end) throw new Error('반차는 하루만 신청할 수 있습니다.');
    return isWorkdate(start, workdaysText, holidays) ? 0.5 : 0;
  }
  var n = 0;
  for (var d = start; d <= end; d = addDays(d, 1)) {
    if (isWorkdate(d, workdaysText, holidays)) n++;
  }
  return n;
}

/**
 * 연차 부여일수와 산정 기간 (근로기준법 제60조, 입사일 기준)
 *  - 1년 미만: 1개월 개근 시 1일 (최대 11일)
 *  - 1년 이상: 15일, 3년 이상부터 2년마다 1일 가산 (최대 25일)
 * 입사일이 없으면 달력연도 기준. manualDays 가 있으면 일수만 그 값으로 대체.
 * @return {{days:number|null, periodStart:string, periodEnd:string, years:number}}
 */
function leaveEntitlement(hireDate, today, manualDays) {
  var manual = manualDays === '' || manualDays === null || manualDays === undefined || isNaN(Number(manualDays))
    ? null : Number(manualDays);
  if (!hireDate) {
    var y = today.slice(0, 4);
    return { days: manual, periodStart: y + '-01-01', periodEnd: y + '-12-31', years: 0 };
  }
  if (today < hireDate) {
    return { days: manual === null ? 0 : manual, periodStart: hireDate, periodEnd: addDays(addYears(hireDate, 1), -1), years: 0 };
  }
  var years = 0;
  while (addYears(hireDate, years + 1) <= today) years++;

  var days;
  if (years === 0) {
    var months = 0;
    while (months < 11 && addMonths(hireDate, months + 1) <= today) months++;
    days = months;
  } else {
    days = Math.min(25, 15 + Math.floor((years - 1) / 2));
  }
  return {
    days: manual === null ? days : manual,
    periodStart: addYears(hireDate, years),
    periodEnd: addDays(addYears(hireDate, years + 1), -1),
    years: years
  };
}

/**
 * 연차 사용 현황
 * @param {Array<{employeeId,type,start,days,status}>} requests
 */
function leaveBalance(requests, employeeId, entitlement) {
  var used = 0, pending = 0;
  requests.forEach(function (r) {
    if (r.employeeId !== employeeId || LEAVE_DEDUCT.indexOf(r.type) === -1) return;
    if (r.start < entitlement.periodStart || r.start > entitlement.periodEnd) return;
    if (r.status === LEAVE_STATUS.OK) used += r.days;
    else if (r.status === LEAVE_STATUS.REQ) pending += r.days;
  });
  return {
    granted: entitlement.days,
    used: used,
    pending: pending,
    remaining: entitlement.days === null ? null : entitlement.days - used - pending
  };
}

/** 해당 날짜에 승인된 휴가: 사번 → 종류 */
function leavesOn(requests, ymd) {
  var map = {};
  requests.forEach(function (r) {
    if (r.status === LEAVE_STATUS.OK && r.start <= ymd && ymd <= r.end) {
      // 오전반차+오후반차가 같은 날 겹치면 종일 휴가로 취급
      map[r.employeeId] = map[r.employeeId] && map[r.employeeId] !== r.type ? '연차' : r.type;
    }
  });
  return map;
}

/** 기존 신청(신청/승인)과 기간이 겹치는 건. 오전반차/오후반차 조합은 허용. */
function findOverlap(requests, employeeId, type, start, end) {
  return requests.filter(function (r) {
    if (r.employeeId !== employeeId) return false;
    if (r.status !== LEAVE_STATUS.REQ && r.status !== LEAVE_STATUS.OK) return false;
    if (r.end < start || r.start > end) return false;
    return !(isHalfDay(type) && isHalfDay(r.type) && type !== r.type);
  })[0] || null;
}

/** 휴가 신청 검증. 문제가 있으면 사유 문자열, 없으면 null */
function validateLeaveRequest(req, opts) {
  if (LEAVE_TYPES.indexOf(req.type) === -1) return '휴가 종류를 선택하세요.';
  if (!req.start || !req.end) return '날짜를 확인하세요.';
  if (req.end < req.start) return '종료일이 시작일보다 빠릅니다.';
  if (isHalfDay(req.type) && req.start !== req.end) return '반차는 하루만 신청할 수 있습니다.';
  if (req.days <= 0) return '선택한 기간에 근무일이 없습니다 (주말·공휴일 제외).';
  var dup = findOverlap(opts.requests, req.employeeId, req.type, req.start, req.end);
  if (dup) return '이미 신청한 휴가와 기간이 겹칩니다 (' + dup.start + ' ' + dup.type + ', ' + dup.status + ').';
  if (LEAVE_DEDUCT.indexOf(req.type) !== -1 && opts.balance.remaining !== null && req.days > opts.balance.remaining) {
    return '잔여 연차가 부족합니다 (잔여 ' + opts.balance.remaining + '일, 신청 ' + req.days + '일).';
  }
  return null;
}

if (typeof module !== 'undefined') {
  module.exports = {
    ACTIONS: ACTIONS, STATES: STATES, LEAVE_TYPES: LEAVE_TYPES, LEAVE_STATUS: LEAVE_STATUS,
    parseHm: parseHm, formatHm: formatHm, parseHmList: parseHmList,
    normalizeYmd: normalizeYmd, addDays: addDays, addMonths: addMonths, addYears: addYears,
    weekdayOfYmd: weekdayOfYmd, isWorkday: isWorkday, isWorkdate: isWorkdate,
    haversineMeters: haversineMeters,
    computeEmployeeDay: computeEmployeeDay, validateAction: validateAction, isLate: isLate,
    workRuleFor: workRuleFor, rankIndex: rankIndex,
    computeDailyStatus: computeDailyStatus, dueReportSlot: dueReportSlot,
    countLeaveDays: countLeaveDays, leaveEntitlement: leaveEntitlement, leaveBalance: leaveBalance,
    leavesOn: leavesOn, findOverlap: findOverlap, validateLeaveRequest: validateLeaveRequest
  };
}
