/**
 * 휴가 · 연차 관리
 *  - 직원: 웹앱 [휴가] 탭에서 신청 / 취소, 잔여 연차 확인
 *  - 센터장: 이메일 링크(또는 [휴가] 시트 상태 변경)로 승인 / 반려
 *    (센터장 본인의 휴가는 회장에게 승인 요청)
 *  - [연차현황] 시트: 전 직원 부여 / 사용 / 잔여 자동 집계
 */

function leaveContext_(emp, today) {
  var requests = getLeaves();
  var ent = leaveEntitlement(emp.hireDate, today, emp.manualLeaveDays);
  return { requests: requests, entitlement: ent, balance: leaveBalance(requests, emp.id, ent) };
}

/** 웹앱: 내 연차 현황 + 신청 내역 */
function apiLeaveInfo(employeeId, pin) {
  var emp = findEmployee_(employeeId, pin);
  var today = todayStr_();
  var ctx = leaveContext_(emp, today);
  var mine = ctx.requests
    .filter(function (r) { return r.employeeId === emp.id && (r.end >= ctx.entitlement.periodStart); })
    .sort(function (a, b) { return a.start < b.start ? 1 : -1; })
    .map(function (r) {
      return { id: r.id, type: r.type, start: r.start, end: r.end, days: r.days, reason: r.reason, status: r.status, memo: r.memo };
    });
  return {
    types: LEAVE_TYPES,
    today: today,
    periodStart: ctx.entitlement.periodStart,
    periodEnd: ctx.entitlement.periodEnd,
    granted: ctx.balance.granted,
    used: ctx.balance.used,
    pending: ctx.balance.pending,
    remaining: ctx.balance.remaining,
    requests: mine
  };
}

/** 웹앱: 휴가 신청 */
function apiRequestLeave(employeeId, pin, type, start, end, reason) {
  var emp = findEmployee_(employeeId, pin);
  var settings = getSettings();
  var center = getCenters().filter(function (c) { return c.code === emp.centerCode; })[0];
  if (!center) throw new Error('센터 정보가 없습니다. 관리자에게 문의하세요.');

  start = normalizeYmd(start);
  end = normalizeYmd(isHalfDay(type) ? start : end);
  var req;
  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    var ctx = leaveContext_(emp, todayStr_());
    var days = start && end && end >= start ? countLeaveDays(type, start, end, settings.workdays, getHolidays()) : 0;
    req = {
      id: 'L' + Utilities.getUuid().slice(0, 7),
      employeeId: emp.id, name: emp.name, rank: emp.rank, centerCode: emp.centerCode,
      type: type, start: start, end: end, days: days,
      reason: String(reason || '').slice(0, 200),
      status: LEAVE_STATUS.REQ,
      requestedAt: nowStr_()
    };
    var problem = validateLeaveRequest(req, { requests: ctx.requests, balance: ctx.balance });
    if (problem) throw new Error(problem);

    sheet_(SHEET.LEAVE).appendRow([
      req.id, req.employeeId, req.name, req.rank, req.centerCode, req.type, req.start, req.end,
      req.days, req.reason, req.status, req.requestedAt, '', '', ''
    ]);
    SpreadsheetApp.flush();
  } finally {
    lock.releaseLock();
  }

  try { notifyLeaveRequest_(req, emp, center, settings); } catch (e) { console.error('휴가 알림 실패', e); }
  try { refreshLeaveSheet(); } catch (e) { console.error('연차현황 갱신 실패', e); }
  return { message: req.type + ' ' + periodText_(req) + ' (' + req.days + '일) 신청 완료. 승인을 기다려 주세요.' };
}

/** 웹앱: 승인 전 신청 취소 */
function apiCancelLeave(employeeId, pin, leaveId) {
  var emp = findEmployee_(employeeId, pin);
  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  var req;
  try {
    req = getLeaves().filter(function (r) { return r.id === leaveId && r.employeeId === emp.id; })[0];
    if (!req) throw new Error('신청 내역을 찾을 수 없습니다.');
    if (req.status !== LEAVE_STATUS.REQ) throw new Error('이미 ' + req.status + '된 휴가는 취소할 수 없습니다. 센터장에게 문의하세요.');
    writeLeaveDecision_(req, LEAVE_STATUS.CANCEL, '본인', '');
  } finally {
    lock.releaseLock();
  }
  try { refreshLeaveSheet(); } catch (e) { console.error('연차현황 갱신 실패', e); }
  return { message: '신청을 취소했습니다.' };
}

function writeLeaveDecision_(req, status, by, memo) {
  var sh = sheet_(SHEET.LEAVE);
  sh.getRange(req.row, 11).setValue(status);
  sh.getRange(req.row, 13, 1, 3).setValues([[by, nowStr_(), String(memo || '').slice(0, 200)]]);
  SpreadsheetApp.flush();
}

/* ───────────── 승인 링크 ───────────── */

function approvalSecret_() {
  var props = PropertiesService.getScriptProperties();
  var s = props.getProperty('APPROVAL_SECRET');
  if (!s) {
    s = Utilities.getUuid() + Utilities.getUuid();
    props.setProperty('APPROVAL_SECRET', s);
  }
  return s;
}

function leaveToken_(leaveId) {
  return Utilities.computeHmacSha256Signature(leaveId, approvalSecret_())
    .slice(0, 12)
    .map(function (b) { var h = (b & 0xff).toString(16); return h.length === 1 ? '0' + h : h; })
    .join('');
}

function approvalUrl_(leaveId) {
  var base = ScriptApp.getService().getUrl();
  return base ? base + '?leave=' + encodeURIComponent(leaveId) + '&t=' + leaveToken_(leaveId) : '';
}

/** 승인자: 일반 직원 → 센터장, 센터장 본인 → 회장 */
function approverOf_(emp, center, settings) {
  var isManager = emp.rank === '센터장' || (center.managerName && emp.name === center.managerName);
  return isManager
    ? { label: settings.chairmanName, email: settings.chairmanEmail, phone: settings.chairmanPhone }
    : { label: center.name + ' 센터장 ' + center.managerName, email: center.managerEmail, phone: center.managerPhone };
}

/** 이메일 링크 → 승인/반려 화면 (링크만 열어서는 처리되지 않음 — 메일 보안검사의 자동 클릭 방지) */
function leaveDecisionPage_(leaveId, token) {
  var t = HtmlService.createTemplateFromFile('Approve');
  t.token = token;
  t.req = null;
  t.error = '';
  if (token !== leaveToken_(leaveId)) {
    t.error = '유효하지 않은 링크입니다.';
  } else {
    var req = getLeaves().filter(function (r) { return r.id === leaveId; })[0];
    if (!req) t.error = '신청 내역을 찾을 수 없습니다.';
    else {
      var center = getCenters().filter(function (c) { return c.code === req.centerCode; })[0];
      t.req = {
        id: req.id, name: req.rank ? req.name + ' ' + req.rank : req.name,
        center: center ? center.name : req.centerCode, type: req.type,
        period: periodText_(req), days: req.days, reason: req.reason || '-',
        status: req.status, decidedBy: req.decidedBy, memo: req.memo
      };
    }
  }
  return t.evaluate()
    .setTitle('휴가 승인')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

/** 승인 화면: 승인/반려 처리 */
function apiDecideLeave(leaveId, token, decision, memo) {
  if (token !== leaveToken_(leaveId)) throw new Error('유효하지 않은 링크입니다.');
  if (decision !== LEAVE_STATUS.OK && decision !== LEAVE_STATUS.NO) throw new Error('잘못된 요청입니다.');
  var settings = getSettings();
  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  var req, emp, center;
  try {
    req = getLeaves().filter(function (r) { return r.id === leaveId; })[0];
    if (!req) throw new Error('신청 내역을 찾을 수 없습니다.');
    if (req.status !== LEAVE_STATUS.REQ) throw new Error('이미 처리된 신청입니다 (' + req.status + ').');
    emp = getEmployees().filter(function (e) { return e.id === req.employeeId; })[0] || { name: req.name, rank: req.rank, phone: '' };
    center = getCenters().filter(function (c) { return c.code === req.centerCode; })[0] || { name: req.centerCode, managerName: '' };
    writeLeaveDecision_(req, decision, approverOf_(emp, center, settings).label, memo);
    req.status = decision;
    req.memo = memo;
  } finally {
    lock.releaseLock();
  }
  try { notifyLeaveDecision_(req, emp, center, settings); } catch (e) { console.error('휴가 처리 알림 실패', e); }
  try { refreshLeaveSheet(); refreshLiveSheet(); } catch (e) { console.error('현황 갱신 실패', e); }
  return { message: req.name + ' ' + req.type + ' (' + periodText_(req) + ') ' + decision + ' 처리했습니다.' };
}

/* ───────────── 알림 ───────────── */

function periodText_(req) {
  return req.start === req.end ? req.start : req.start + ' ~ ' + req.end;
}

function leaveVariables_(req, centerName, status) {
  return {
    '#{센터}': centerName,
    '#{이름}': req.rank ? req.name + ' ' + req.rank : req.name,
    '#{종류}': req.type,
    '#{기간}': periodText_(req) + ' (' + req.days + '일)',
    '#{상태}': status,
    '#{비고}': req.memo || req.reason || '-'
  };
}

function notifyLeaveRequest_(req, emp, center, settings) {
  var approver = approverOf_(emp, center, settings);
  var link = approvalUrl_(req.id);
  var subject = '[휴가신청] ' + center.name + ' ' + displayName_(emp) + ' ' + req.type + ' ' + periodText_(req);
  var lines = [
    '센터: ' + center.name,
    '직원: ' + displayName_(emp) + ' (' + emp.id + ')',
    '종류: ' + req.type,
    '기간: ' + periodText_(req) + ' (' + req.days + '일)',
    '사유: ' + (req.reason || '-')
  ];
  if (approver.email) {
    MailApp.sendEmail({
      to: approver.email,
      subject: subject,
      body: lines.join('\n') + (link ? '\n\n승인/반려: ' + link : ''),
      htmlBody: '<div style="font-family:sans-serif;font-size:14px">' + lines.map(escapeHtml_).join('<br>') +
        (link ? '<br><br><a href="' + link + '" style="display:inline-block;padding:10px 18px;background:#1f4e79;color:#fff;border-radius:8px;text-decoration:none">승인 / 반려 하기</a>' : '') +
        '<br><br><span style="color:#888">[휴가] 시트의 상태 칸을 "승인"/"반려"로 바꿔도 처리됩니다.</span></div>'
    });
  }
  if (settings.alimtalkOn && settings.leaveTemplateId && approver.phone) {
    sendAlimtalk_(settings, approver.phone, settings.leaveTemplateId,
      leaveVariables_(req, center.name, LEAVE_STATUS.REQ + ' (승인 필요)'), subject + '\n' + lines.join('\n'));
  }
}

function notifyLeaveDecision_(req, emp, center, settings) {
  if (settings.alimtalkOn && settings.leaveTemplateId && emp.phone) {
    sendAlimtalk_(settings, emp.phone, settings.leaveTemplateId,
      leaveVariables_(req, center.name, req.status),
      '[휴가 ' + req.status + '] ' + req.type + ' ' + periodText_(req));
  }
}

/* ───────────── 연차현황 시트 ───────────── */

var LEAVE_STATUS_HEADER = ['센터', '직급', '사번', '이름', '입사일', '근속(년)', '산정기간', '부여', '사용', '승인대기', '잔여'];

function leaveStatusValues_(today) {
  var settings = getSettings();
  var centers = getCenters();
  var centerName = {};
  centers.forEach(function (c) { centerName[c.code] = c.name; });
  var requests = getLeaves();
  return getEmployees()
    .sort(function (a, b) {
      if (a.centerCode !== b.centerCode) return a.centerCode < b.centerCode ? -1 : 1;
      return rankIndex(a.rank, settings.rankOrder) - rankIndex(b.rank, settings.rankOrder);
    })
    .map(function (emp) {
      var ent = leaveEntitlement(emp.hireDate, today, emp.manualLeaveDays);
      var b = leaveBalance(requests, emp.id, ent);
      return [centerName[emp.centerCode] || emp.centerCode, emp.rank, emp.id, emp.name, emp.hireDate || '(미입력)',
        emp.hireDate ? ent.years : '', ent.periodStart + ' ~ ' + ent.periodEnd,
        b.granted === null ? '미설정' : b.granted, b.used, b.pending, b.remaining === null ? '' : b.remaining];
    });
}

function writeLeaveStatus_(sh, today) {
  sh.clear();
  var rows = leaveStatusValues_(today);
  sh.getRange(1, 1).setValue(today + ' 기준 연차 현황').setFontSize(14).setFontWeight('bold');
  sh.getRange(2, 1).setValue('근로기준법 입사일 기준 자동 계산 · 회계연도 기준이면 [직원] 시트 "연차 부여일수(수동)"에 입력');
  sh.getRange(4, 1, 1, LEAVE_STATUS_HEADER.length).setValues([LEAVE_STATUS_HEADER])
    .setFontWeight('bold').setBackground('#1f4e79').setFontColor('#ffffff');
  if (rows.length) sh.getRange(5, 1, rows.length, LEAVE_STATUS_HEADER.length).setValues(rows);
  sh.autoResizeColumns(1, LEAVE_STATUS_HEADER.length);
}

function refreshLeaveSheet() {
  var ss = ss_();
  var sh = ss.getSheetByName(SHEET.LEAVE_STATUS) || ss.insertSheet(SHEET.LEAVE_STATUS);
  writeLeaveStatus_(sh, todayStr_());
}
