/**
 * 메뉴 · 초기 설정 · 트리거
 */

function onOpen() {
  SpreadsheetApp.getUi().createMenu('근태관리')
    .addItem('1. 초기 설정 (시트 생성)', 'setup')
    .addItem('2. 자동 보고 시작 (트리거 등록)', 'installTriggers')
    .addSeparator()
    .addItem('현황 지금 새로고침', 'periodicRefresh')
    .addItem('회장 보고서 지금 발송', 'sendReportNow')
    .addItem('알림 테스트 (회장에게 발송)', 'testNotification')
    .addSeparator()
    .addItem('알림톡 API 키 등록 (솔라피)', 'setSolapiKeys')
    .addItem('자동 보고 중지', 'removeTriggers')
    .addToUi();
}

function setup() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  PropertiesService.getScriptProperties().setProperty('SPREADSHEET_ID', ss.getId());
  ss.setSpreadsheetTimeZone('Asia/Seoul');

  var settings = ensureSheet_(ss, SHEET.SETTINGS);
  if (settings.getLastRow() < 2) {
    settings.getRange(2, 1, DEFAULT_SETTINGS.length, 3).setNumberFormat('@').setValues(DEFAULT_SETTINGS);
  } else {
    // 새 버전에서 추가된 항목만 보충
    var existing = dataRows_(SHEET.SETTINGS).map(function (r) { return String(r[0]).trim(); });
    DEFAULT_SETTINGS.forEach(function (row) {
      if (existing.indexOf(row[0]) === -1) settings.appendRow(row);
    });
  }

  var centers = ensureSheet_(ss, SHEET.CENTERS);
  centers.getRange('A:E').setNumberFormat('@');
  if (centers.getLastRow() < 2) {
    var rows = [];
    for (var i = 1; i <= 12; i++) {
      var code = 'C' + (i < 10 ? '0' : '') + i;
      rows.push([code, '센터' + i, '', '', '', '', '', 300]);
    }
    centers.getRange(2, 1, rows.length, rows[0].length).setValues(rows);
  }

  var employees = ensureSheet_(ss, SHEET.EMPLOYEES);
  employees.getRange('A:H').setNumberFormat('@');
  if (employees.getLastRow() < 2) {
    employees.getRange(2, 1, 1, 9).setValues([['E0001', '홍길동(예시)', '사원', 'C01', '01000000000', '1234', 'Y', '2024-03-01', '']]);
  }
  var rankList = getSettings().rankOrder;
  employees.getRange('C2:C').setDataValidation(SpreadsheetApp.newDataValidation()
    .requireValueInList(rankList, true).setAllowInvalid(true).build());
  employees.getRange('G2:G').setDataValidation(SpreadsheetApp.newDataValidation()
    .requireValueInList(['Y', 'N'], true).build());
  employees.getRange(1, 8).setNote('예: 2024-03-01. 연차가 입사일 기준으로 자동 계산됩니다.');
  employees.getRange(1, 9).setNote('비워 두면 자동 계산. 회계연도 기준 등으로 직접 정하려면 일수 입력.');

  var leave = ensureSheet_(ss, SHEET.LEAVE);
  leave.getRange('A:H').setNumberFormat('@');
  leave.getRange('J:O').setNumberFormat('@');
  leave.getRange('F2:F').setDataValidation(SpreadsheetApp.newDataValidation()
    .requireValueInList(LEAVE_TYPES, true).build());
  leave.getRange('K2:K').setDataValidation(SpreadsheetApp.newDataValidation()
    .requireValueInList([LEAVE_STATUS.REQ, LEAVE_STATUS.OK, LEAVE_STATUS.NO, LEAVE_STATUS.CANCEL], true).build());
  leave.getRange(1, 11).setNote('관리자가 직접 "승인"/"반려"로 바꿔도 됩니다. (10분 내 현황 반영)');

  var holidays = ensureSheet_(ss, SHEET.HOLIDAYS);
  holidays.getRange('A:A').setNumberFormat('@');
  if (holidays.getLastRow() < 2) {
    holidays.getRange(2, 1, DEFAULT_HOLIDAYS.length, 2).setValues(DEFAULT_HOLIDAYS);
  }

  var log = ensureSheet_(ss, SHEET.LOG);
  log.getRange('A:I').setNumberFormat('@');
  log.getRange('L:M').setNumberFormat('@');

  if (!ss.getSheetByName(SHEET.LIVE)) ss.insertSheet(SHEET.LIVE, 0);
  var blank = ss.getSheetByName('시트1') || ss.getSheetByName('Sheet1');
  if (blank && blank.getLastRow() === 0) ss.deleteSheet(blank);

  periodicRefresh();
  SpreadsheetApp.getUi().alert(
    '초기 설정 완료',
    '[설정] [센터] [직원] [공휴일] 시트를 채운 뒤, 메뉴 [2. 자동 보고 시작]을 실행하세요.\n' +
    '직원용 출퇴근 화면은 [배포 > 새 배포 > 웹 앱]으로 만든 주소를 직원들에게 공유합니다.',
    SpreadsheetApp.getUi().ButtonSet.OK);
}

function ensureSheet_(ss, name) {
  var sh = ss.getSheetByName(name) || ss.insertSheet(name);
  var header = HEADERS[name];
  if (header) {
    sh.getRange(1, 1, 1, header.length).setValues([header])
      .setFontWeight('bold').setBackground('#1f4e79').setFontColor('#ffffff');
    sh.setFrozenRows(1);
  }
  return sh;
}

function installTriggers() {
  removeTriggers(true);
  ScriptApp.newTrigger('scheduledReportTick').timeBased().everyMinutes(5).create();
  ScriptApp.newTrigger('periodicRefresh').timeBased().everyMinutes(10).create();
  var s = getSettings();
  SpreadsheetApp.getUi().alert('자동 보고 시작',
    '보고 시각: ' + s.reportSlots.map(formatHm).join(', ') + ' (근무 요일: ' + s.workdays + ')\n' +
    '각 시각 이후 5분 이내에 발송됩니다.',
    SpreadsheetApp.getUi().ButtonSet.OK);
}

function removeTriggers(silent) {
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (['scheduledReportTick', 'refreshLiveSheet', 'periodicRefresh'].indexOf(t.getHandlerFunction()) !== -1) ScriptApp.deleteTrigger(t);
  });
  if (silent !== true) SpreadsheetApp.getUi().alert('자동 보고를 중지했습니다.');
}

/** 10분마다: 날짜 변경 · 시트에서 직접 바꾼 휴가 승인 등을 현황에 반영 */
function periodicRefresh() {
  refreshLiveSheet();
  refreshLeaveSheet();
}

function sendReportNow() {
  sendReport();
  SpreadsheetApp.getUi().alert('보고서를 발송했습니다.');
}

function setSolapiKeys() {
  var ui = SpreadsheetApp.getUi();
  var key = ui.prompt('솔라피 API Key 입력').getResponseText().trim();
  if (!key) return;
  var secret = ui.prompt('솔라피 API Secret 입력').getResponseText().trim();
  if (!secret) return;
  var props = PropertiesService.getScriptProperties();
  props.setProperty('SOLAPI_API_KEY', key);
  props.setProperty('SOLAPI_API_SECRET', secret);
  ui.alert('저장했습니다. (시트에는 저장되지 않고 스크립트 속성에 안전하게 보관됩니다)');
}

function testNotification() {
  var s = getSettings();
  var ui = SpreadsheetApp.getUi();
  var result = [];
  if (s.chairmanEmail) {
    MailApp.sendEmail(s.chairmanEmail, '[근태관리] 이메일 알림 테스트', '이메일 알림이 정상적으로 설정되었습니다.');
    result.push('이메일 → ' + s.chairmanEmail);
  }
  if (s.alimtalkOn && s.chairmanPhone && s.eventTemplateId) {
    var ok = sendAlimtalk_(s, s.chairmanPhone, s.eventTemplateId, {
      '#{센터}': '테스트센터', '#{이름}': '테스트', '#{구분}': ACTIONS.IN, '#{시각}': formatHm(nowMinutes_()), '#{비고}': '알림 테스트'
    }, '[근태관리] 알림톡 테스트');
    result.push('알림톡 → ' + s.chairmanPhone + (ok ? ' (성공)' : ' (실패: 실행 로그 확인)'));
  }
  ui.alert(result.length ? result.join('\n') : '[설정]에 회장 이메일/휴대폰을 입력하세요.');
}
