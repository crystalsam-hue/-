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
  LEFT: '퇴근'
};

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

var WEEKDAY_NAMES = ['일', '월', '화', '수', '목', '금', '토'];

/** 근무요일 문자열("월,화,수,목,금")에 해당 요일(0=일)이 포함되는지 */
function isWorkday(weekdayIndex, workdaysText) {
  var list = String(workdaysText || '월,화,수,목,금').split(/[,\s]+/);
  return list.indexOf(WEEKDAY_NAMES[weekdayIndex]) !== -1;
}

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

/**
 * 전체 현황 계산
 * @param {Array} centers  [{code, name, managerName, ...}]
 * @param {Array} employees [{id, name, centerCode, ...}] (재직자만)
 * @param {Array} logs [{employeeId, action, minutes}] 오늘 기록 (시간순)
 * @param {{startMinutes:number, graceMinutes:number}} rule
 */
function computeDailyStatus(centers, employees, logs, rule) {
  var byEmployee = {};
  logs.forEach(function (l) {
    (byEmployee[l.employeeId] = byEmployee[l.employeeId] || []).push(l);
  });

  var rows = employees.map(function (emp) {
    var day = computeEmployeeDay(byEmployee[emp.id] || []);
    return {
      centerCode: emp.centerCode,
      employeeId: emp.id,
      name: emp.name,
      state: day.state,
      inAt: day.inAt,
      outAt: day.outAt,
      backAt: day.backAt,
      leaveAt: day.leaveAt,
      outCount: day.outCount,
      late: isLate(day.inAt, rule.startMinutes, rule.graceMinutes)
    };
  });

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
      absent: count(function (r) { return r.state === STATES.NONE; })
    };
  });

  var total = summaries.reduce(function (acc, s) {
    ['total', 'present', 'late', 'working', 'away', 'left', 'absent'].forEach(function (k) {
      acc[k] += s[k];
    });
    return acc;
  }, { centerCode: '', centerName: '합계', managerName: '', total: 0, present: 0, late: 0, working: 0, away: 0, left: 0, absent: 0 });

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

if (typeof module !== 'undefined') {
  module.exports = {
    ACTIONS: ACTIONS, STATES: STATES, parseHm: parseHm, formatHm: formatHm,
    parseHmList: parseHmList, isWorkday: isWorkday, haversineMeters: haversineMeters,
    computeEmployeeDay: computeEmployeeDay, validateAction: validateAction, isLate: isLate,
    computeDailyStatus: computeDailyStatus, dueReportSlot: dueReportSlot
  };
}
