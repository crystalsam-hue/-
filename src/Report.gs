/**
 * 실시간 현황 시트 + 정기 보고(회장: 엑셀 첨부 이메일/알림톡, 센터장: 자기 센터 현황)
 */

var SUMMARY_HEADER = ['센터', '센터장', '전체', '출근', '지각', '근무중', '외출중', '퇴근', '미출근'];
var DETAIL_HEADER = ['센터', '사번', '이름', '현재상태', '출근', '지각', '최근 외출', '최근 복귀', '외출횟수', '퇴근'];

function buildStatus_(date) {
  var settings = getSettings();
  var centers = getCenters();
  var employees = getEmployees();
  var dateStr = todayStr_(date);
  var status = computeDailyStatus(centers, employees, getLogsForDate(dateStr), {
    startMinutes: settings.startMinutes,
    graceMinutes: settings.graceMinutes
  });
  var centerName = {};
  centers.forEach(function (c) { centerName[c.code] = c.name; });
  status.rows.sort(function (a, b) {
    return a.centerCode === b.centerCode ? a.employeeId.localeCompare(b.employeeId) : a.centerCode.localeCompare(b.centerCode);
  });
  return { settings: settings, centers: centers, dateStr: dateStr, status: status, centerName: centerName };
}

function summaryValues_(s) {
  return [s.centerName, s.managerName, s.total, s.present, s.late, s.working, s.away, s.left, s.absent];
}

function detailValues_(r, centerName) {
  var t = function (m) { return m === null ? '' : formatHm(m); };
  return [centerName[r.centerCode] || r.centerCode, r.employeeId, r.name, r.state,
    t(r.inAt), r.late ? '지각' : '', t(r.outAt), t(r.backAt), r.outCount || '', t(r.leaveAt)];
}

/** 시트에 요약표 + 상세표를 씀 (실시간현황 시트, 엑셀 보고서 공용) */
function writeReport_(sh, ctx, title) {
  sh.clear();
  var summary = [SUMMARY_HEADER]
    .concat(ctx.status.summaries.map(summaryValues_))
    .concat([summaryValues_(ctx.status.total)]);
  var detail = [DETAIL_HEADER].concat(ctx.status.rows.map(function (r) { return detailValues_(r, ctx.centerName); }));

  sh.getRange(1, 1).setValue(title).setFontSize(14).setFontWeight('bold');
  sh.getRange(2, 1).setValue('갱신: ' + Utilities.formatDate(new Date(), 'Asia/Seoul', 'yyyy-MM-dd HH:mm:ss'));

  var sTop = 4;
  sh.getRange(sTop, 1, summary.length, SUMMARY_HEADER.length).setValues(summary);
  sh.getRange(sTop, 1, 1, SUMMARY_HEADER.length).setFontWeight('bold').setBackground('#1f4e79').setFontColor('#ffffff');
  sh.getRange(sTop + summary.length - 1, 1, 1, SUMMARY_HEADER.length).setFontWeight('bold').setBackground('#ddebf7');

  var dTop = sTop + summary.length + 2;
  sh.getRange(dTop - 1, 1).setValue('직원별 상세').setFontWeight('bold');
  sh.getRange(dTop, 1, detail.length, DETAIL_HEADER.length).setValues(detail);
  sh.getRange(dTop, 1, 1, DETAIL_HEADER.length).setFontWeight('bold').setBackground('#1f4e79').setFontColor('#ffffff');

  // 상태 강조 (미출근: 빨강, 외출중: 노랑, 지각: 붉은 글씨) — 한 번에 적용
  if (detail.length > 1) {
    var body = detail.slice(1);
    sh.getRange(dTop + 1, 1, body.length, DETAIL_HEADER.length).setBackgrounds(body.map(function (r) {
      var color = r[3] === STATES.NONE ? '#fde2e1' : r[3] === STATES.AWAY ? '#fff2cc' : null;
      return DETAIL_HEADER.map(function () { return color; });
    }));
    sh.getRange(dTop + 1, 6, body.length, 1).setFontColor('#c00000').setFontWeight('bold');
  }
  sh.autoResizeColumns(1, DETAIL_HEADER.length);
}

/** [실시간현황] 시트 갱신 — 기록이 들어올 때마다 호출 */
function refreshLiveSheet() {
  var lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    var ctx = buildStatus_();
    var ss = ss_();
    var sh = ss.getSheetByName(SHEET.LIVE) || ss.insertSheet(SHEET.LIVE, 0);
    writeReport_(sh, ctx, ctx.dateStr + ' (' + WEEKDAY_NAMES[weekdayIndex_()] + ') 전 센터 근태 현황');
  } finally {
    lock.releaseLock();
  }
}

/** 5분마다 실행: 보고 시각이 지났고 아직 안 보냈으면 발송 */
function scheduledReportTick() {
  var now = new Date();
  var settings = getSettings();
  if (!isWorkday(weekdayIndex_(now), settings.workdays)) return;

  var props = PropertiesService.getScriptProperties();
  var key = 'SENT_' + todayStr_(now);
  var sent = JSON.parse(props.getProperty(key) || '[]');
  var slot = dueReportSlot(nowMinutes_(now), settings.reportSlots, sent);
  if (slot === null) return;

  // 이 시각 이전의 미발송 슬롯은 함께 처리된 것으로 간주 (중복 발송 방지)
  settings.reportSlots.forEach(function (s) { if (s <= slot && sent.indexOf(s) === -1) sent.push(s); });
  props.setProperty(key, JSON.stringify(sent));

  sendReport(formatHm(slot));
  cleanupSentKeys_(props, key);
}

function cleanupSentKeys_(props, keepKey) {
  Object.keys(props.getProperties()).forEach(function (k) {
    if (k.indexOf('SENT_') === 0 && k !== keepKey) props.deleteProperty(k);
  });
}

/** 회장 보고서 발송 (메뉴에서 수동 실행 가능) */
function sendReport(slotLabel) {
  var ctx = buildStatus_();
  var settings = ctx.settings;
  var label = typeof slotLabel === 'string' ? slotLabel : Utilities.formatDate(new Date(), 'Asia/Seoul', 'HH:mm');
  var title = ctx.dateStr + ' ' + label + ' 전 센터 근태 현황';
  var t = ctx.status.total;

  refreshLiveSheet();

  // 1) 회장: 이메일 + 엑셀 첨부
  if (settings.chairmanEmail) {
    var xlsx = exportXlsx_(ctx, title);
    MailApp.sendEmail({
      to: settings.chairmanEmail,
      subject: '[근태보고] ' + title + ' - 출근 ' + t.present + '/' + t.total + ', 지각 ' + t.late + ', 미출근 ' + t.absent,
      htmlBody: reportHtml_(ctx, title, ctx.status.summaries.concat([t]), ctx.status.rows),
      attachments: [xlsx]
    });
  }

  // 2) 회장: 알림톡 요약
  if (settings.alimtalkOn && settings.reportTemplateId && settings.chairmanPhone) {
    sendAlimtalk_(settings, settings.chairmanPhone, settings.reportTemplateId,
      reportVariables_(ctx.dateStr, label, '전체 센터', t),
      title + '\n전체 ' + t.total + ' / 출근 ' + t.present + ' / 지각 ' + t.late + ' / 미출근 ' + t.absent + ' / 외출중 ' + t.away);
  }

  // 3) 센터장: 자기 센터 현황
  if (settings.managerReports) {
    ctx.centers.forEach(function (c) {
      var s = ctx.status.summaries.filter(function (x) { return x.centerCode === c.code; })[0];
      var rows = ctx.status.rows.filter(function (r) { return r.centerCode === c.code; });
      if (!s || !s.total) return;
      var cTitle = ctx.dateStr + ' ' + label + ' ' + c.name + ' 근태 현황';
      if (settings.emailOn && c.managerEmail) {
        MailApp.sendEmail({
          to: c.managerEmail,
          subject: '[근태보고] ' + cTitle + ' - 미출근 ' + s.absent + '명, 외출중 ' + s.away + '명',
          htmlBody: reportHtml_(ctx, cTitle, [s], rows)
        });
      }
      if (settings.alimtalkOn && settings.reportTemplateId && c.managerPhone) {
        sendAlimtalk_(settings, c.managerPhone, settings.reportTemplateId,
          reportVariables_(ctx.dateStr, label, c.name, s),
          cTitle + '\n전체 ' + s.total + ' / 출근 ' + s.present + ' / 지각 ' + s.late + ' / 미출근 ' + s.absent + ' / 외출중 ' + s.away);
      }
    });
  }
}

function reportVariables_(dateStr, label, scope, s) {
  return {
    '#{일자}': dateStr,
    '#{시각}': label,
    '#{범위}': scope,
    '#{전체}': String(s.total),
    '#{출근}': String(s.present),
    '#{지각}': String(s.late),
    '#{미출근}': String(s.absent),
    '#{외출}': String(s.away),
    '#{퇴근}': String(s.left)
  };
}

function reportHtml_(ctx, title, summaries, rows) {
  var table = function (header, body) {
    var th = header.map(function (h) {
      return '<th style="background:#1f4e79;color:#fff;padding:4px 8px;border:1px solid #ccc">' + escapeHtml_(h) + '</th>';
    }).join('');
    var trs = body.map(function (r) {
      var bg = r[3] === STATES.NONE ? '#fde2e1' : r[3] === STATES.AWAY ? '#fff2cc' : '#fff';
      return '<tr style="background:' + bg + '">' + r.map(function (v) {
        return '<td style="padding:4px 8px;border:1px solid #ccc">' + escapeHtml_(v) + '</td>';
      }).join('') + '</tr>';
    }).join('');
    return '<table style="border-collapse:collapse;font-size:13px">' + '<tr>' + th + '</tr>' + trs + '</table>';
  };
  var attention = rows.filter(function (r) { return r.state === STATES.NONE || r.state === STATES.AWAY || r.late; });
  return '<div style="font-family:sans-serif">' +
    '<h3>' + escapeHtml_(title) + '</h3>' +
    table(SUMMARY_HEADER, summaries.map(summaryValues_)) +
    '<h4>확인 필요 (미출근 · 외출중 · 지각)</h4>' +
    (attention.length
      ? table(DETAIL_HEADER, attention.map(function (r) { return detailValues_(r, ctx.centerName); }))
      : '<p>없음</p>') +
    '<p><a href="' + ss_().getUrl() + '">실시간 현황 시트 열기</a> (전체 직원 상세는 첨부 엑셀 참고)</p>' +
    '</div>';
}

/** 보고서만 담은 임시 스프레드시트를 만들어 .xlsx 로 내보낸 뒤 삭제 */
function exportXlsx_(ctx, title) {
  var tmp = SpreadsheetApp.create('근태보고_' + ctx.dateStr);
  try {
    var sh = tmp.getSheets()[0].setName('근태현황');
    writeReport_(sh, ctx, title);
    SpreadsheetApp.flush();
    var url = 'https://docs.google.com/spreadsheets/d/' + tmp.getId() + '/export?format=xlsx';
    var blob = UrlFetchApp.fetch(url, {
      headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() }
    }).getBlob();
    return blob.setName('근태보고_' + ctx.dateStr + '_' + title.split(' ')[1].replace(':', '') + '.xlsx');
  } finally {
    DriveApp.getFileById(tmp.getId()).setTrashed(true);
  }
}
