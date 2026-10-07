/**
 * 웹앱(직원용 출퇴근 · 휴가 화면) 및 출퇴근 기록 처리
 */

function doGet(e) {
  var p = (e && e.parameter) || {};
  if (p.leave && p.t) return leaveDecisionPage_(p.leave, p.t);
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('출퇴근 체크')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

var MAX_PIN_FAILS = 5;

function findEmployee_(employeeId, pin) {
  var id = String(employeeId).trim();
  var cache = CacheService.getScriptCache();
  var failKey = 'PINFAIL_' + id;
  var fails = Number(cache.get(failKey) || 0);
  if (fails >= MAX_PIN_FAILS) throw new Error('PIN을 ' + MAX_PIN_FAILS + '회 틀렸습니다. 10분 후 다시 시도하세요.');

  var emp = getEmployees().filter(function (e) { return e.id === id; })[0];
  if (!emp || !emp.pin || emp.pin !== String(pin).trim()) {
    cache.put(failKey, String(fails + 1), 600);
    throw new Error('사번 또는 PIN이 올바르지 않습니다.');
  }
  if (fails) cache.remove(failKey);
  return emp;
}

function stateOf_(employeeId, dateStr) {
  var events = getLogsForDate(dateStr).filter(function (l) { return l.employeeId === employeeId; });
  return computeEmployeeDay(events);
}

function todayLeaveOf_(employeeId, dateStr) {
  return leavesOn(getLeaves(), dateStr)[employeeId] || '';
}

/** 웹앱: 로그인 후 현재 상태 조회 */
function apiStatus(employeeId, pin) {
  var emp = findEmployee_(employeeId, pin);
  var center = getCenters().filter(function (c) { return c.code === emp.centerCode; })[0];
  var today = todayStr_();
  var day = stateOf_(emp.id, today);
  return {
    name: displayName_(emp),
    centerName: center ? center.name : emp.centerCode,
    state: day.state,
    todayLeave: todayLeaveOf_(emp.id, today),
    inAt: day.inAt === null ? '' : formatHm(day.inAt),
    leaveAt: day.leaveAt === null ? '' : formatHm(day.leaveAt),
    geoRequired: getSettings().geoOn
  };
}

/** 웹앱: 출근/외출/복귀/퇴근 기록 */
function apiRecord(employeeId, pin, action, note, lat, lng) {
  var emp = findEmployee_(employeeId, pin);
  var settings = getSettings();
  var center = getCenters().filter(function (c) { return c.code === emp.centerCode; })[0];
  if (!center) throw new Error('센터코드(' + emp.centerCode + ')가 [센터] 시트에 없습니다. 관리자에게 문의하세요.');

  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  var record;
  try {
    var now = new Date();
    var dateStr = todayStr_(now);
    var day = stateOf_(emp.id, dateStr);
    var reason = validateAction(day.state, action);
    if (reason) throw new Error(reason);

    var distance = '';
    var geoJudge = '';
    var hasPos = lat !== null && lat !== undefined && lat !== '' && lng !== null && lng !== undefined && lng !== '';
    if (hasPos && center.lat !== null && center.lng !== null) {
      distance = haversineMeters(Number(lat), Number(lng), center.lat, center.lng);
      geoJudge = distance <= center.radius ? '정상' : '범위밖';
    } else if (settings.geoOn) {
      geoJudge = hasPos ? '센터좌표없음' : '위치없음';
    }

    var timeStr = Utilities.formatDate(now, 'Asia/Seoul', 'HH:mm:ss');
    var minutes = nowMinutes_(now);
    var leaveType = todayLeaveOf_(emp.id, dateStr);
    var rule = workRuleFor(leaveType, workRule_(settings));
    var tags = [];
    if (leaveType) tags.push(leaveType);
    if (action === ACTIONS.IN && (!leaveType || isHalfDay(leaveType)) && isLate(minutes, rule.startMinutes, rule.graceMinutes)) tags.push('지각');
    if (action === ACTIONS.LEAVE && minutes < rule.endMinutes) tags.push('조퇴');
    if (geoJudge === '범위밖') tags.push('위치 범위밖 ' + distance + 'm');

    record = {
      id: Utilities.getUuid().slice(0, 8),
      date: dateStr,
      time: timeStr,
      employeeId: emp.id,
      name: displayName_(emp),
      centerCode: center.code,
      centerName: center.name,
      action: action,
      note: String(note || '').slice(0, 200),
      tags: tags
    };

    sheet_(SHEET.LOG).appendRow([
      record.id, record.date, record.time, record.employeeId, emp.name,
      record.centerCode, record.centerName, record.action,
      [record.note].concat(tags).filter(String).join(' / '),
      hasPos ? Number(lat) : '', hasPos ? Number(lng) : '', distance, geoJudge
    ]);
    SpreadsheetApp.flush();
  } finally {
    lock.releaseLock();
  }

  // 알림/현황 갱신은 실패해도 기록 자체는 유지
  try { notifyEvent_(record, center, settings); } catch (e) { console.error('알림 실패', e); }
  try { refreshLiveSheet(); } catch (e) { console.error('현황 갱신 실패', e); }

  return {
    message: record.name + '님 ' + record.action + ' 완료 (' + record.time.slice(0, 5) + ')' +
      (record.tags.length ? ' · ' + record.tags.join(', ') : '')
  };
}
