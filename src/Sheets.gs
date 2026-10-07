/**
 * 시트 이름·구조 정의 및 데이터 읽기
 */

var SHEET = {
  SETTINGS: '설정',
  CENTERS: '센터',
  EMPLOYEES: '직원',
  LOG: '기록',
  LIVE: '실시간현황'
};

var HEADERS = {};
HEADERS[SHEET.SETTINGS] = ['항목', '값', '설명'];
HEADERS[SHEET.CENTERS] = ['센터코드', '센터명', '센터장 이름', '센터장 이메일', '센터장 휴대폰', '위도', '경도', '허용반경(m)'];
HEADERS[SHEET.EMPLOYEES] = ['사번', '이름', '센터코드', '휴대폰', 'PIN(4~6자리)', '재직(Y/N)'];
HEADERS[SHEET.LOG] = ['기록ID', '일자', '시각', '사번', '이름', '센터코드', '센터명', '구분', '비고', '위도', '경도', '센터거리(m)', '위치판정'];

/** [항목, 기본값, 설명] */
var DEFAULT_SETTINGS = [
  ['회장 이름', '회장님', ''],
  ['회장 이메일', '', '보고서를 받을 이메일 (여러 명이면 쉼표로 구분)'],
  ['회장 휴대폰', '', '알림톡 보고를 받을 번호 (예: 01012345678)'],
  ['출근 기준시각', '09:00', 'HH:MM'],
  ['지각 허용(분)', '10', '기준시각 + 허용분 이후 출근은 지각'],
  ['퇴근 기준시각', '18:00', 'HH:MM (기준 이전 퇴근은 "조퇴"로 표시)'],
  ['보고 시각', '09:30, 13:30, 18:30', '회장 보고서(엑셀 첨부) 발송 시각. 쉼표로 여러 개'],
  ['근무 요일', '월,화,수,목,금', '이 요일에만 정기 보고 발송'],
  ['이메일 알림', 'Y', 'Y/N - 센터장에게 실시간 이메일'],
  ['알림톡 알림', 'N', 'Y/N - 센터장에게 실시간 알림톡 (솔라피 설정 필요)'],
  ['회장 실시간 알림', 'N', 'Y/N - 모든 출퇴근 이벤트를 회장에게도 즉시 전송'],
  ['센터장 정기보고', 'Y', 'Y/N - 보고 시각마다 각 센터장에게 자기 센터 현황 전송'],
  ['위치 확인', 'N', 'Y/N - 휴대폰 GPS로 센터 반경 확인 (반경 밖이면 "범위밖" 표시)'],
  ['알림톡 발신번호', '', '솔라피에 등록된 발신번호'],
  ['카카오 채널 pfId', '', '솔라피에 연동한 카카오 비즈니스 채널 ID'],
  ['알림톡 템플릿(이벤트)', '', '검수 완료된 출퇴근 알림 템플릿 ID'],
  ['알림톡 템플릿(보고)', '', '검수 완료된 현황 보고 템플릿 ID']
];

function ss_() {
  return SpreadsheetApp.getActiveSpreadsheet() ||
    SpreadsheetApp.openById(PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID'));
}

function sheet_(name) {
  var sh = ss_().getSheetByName(name);
  if (!sh) throw new Error('"' + name + '" 시트가 없습니다. 메뉴 [근태관리 > 초기 설정]을 먼저 실행하세요.');
  return sh;
}

/** 헤더 제외 데이터 행 */
function dataRows_(name) {
  var sh = sheet_(name);
  var last = sh.getLastRow();
  if (last < 2) return [];
  return sh.getRange(2, 1, last - 1, sh.getLastColumn()).getValues();
}

function yn_(v) {
  return String(v).trim().toUpperCase() === 'Y';
}

function digits_(v) {
  return String(v || '').replace(/\D/g, '');
}

function getSettings() {
  var map = {};
  dataRows_(SHEET.SETTINGS).forEach(function (r) { map[String(r[0]).trim()] = r[1]; });
  var get = function (k) {
    var v = map[k];
    if (v instanceof Date) return Utilities.formatDate(v, 'Asia/Seoul', 'HH:mm');
    return v === undefined || v === null ? '' : String(v).trim();
  };
  return {
    chairmanName: get('회장 이름') || '회장님',
    chairmanEmail: get('회장 이메일'),
    chairmanPhone: digits_(get('회장 휴대폰')),
    startMinutes: parseHm(get('출근 기준시각') || '09:00'),
    graceMinutes: Number(get('지각 허용(분)') || 0),
    endMinutes: parseHm(get('퇴근 기준시각') || '18:00'),
    reportSlots: parseHmList(get('보고 시각')),
    workdays: get('근무 요일'),
    emailOn: yn_(get('이메일 알림')),
    alimtalkOn: yn_(get('알림톡 알림')),
    chairmanRealtime: yn_(get('회장 실시간 알림')),
    managerReports: yn_(get('센터장 정기보고')),
    geoOn: yn_(get('위치 확인')),
    senderPhone: digits_(get('알림톡 발신번호')),
    pfId: get('카카오 채널 pfId'),
    eventTemplateId: get('알림톡 템플릿(이벤트)'),
    reportTemplateId: get('알림톡 템플릿(보고)')
  };
}

function getCenters() {
  return dataRows_(SHEET.CENTERS)
    .filter(function (r) { return String(r[0]).trim(); })
    .map(function (r) {
      return {
        code: String(r[0]).trim(),
        name: String(r[1]).trim(),
        managerName: String(r[2]).trim(),
        managerEmail: String(r[3]).trim(),
        managerPhone: digits_(r[4]),
        lat: r[5] === '' ? null : Number(r[5]),
        lng: r[6] === '' ? null : Number(r[6]),
        radius: Number(r[7] || 300)
      };
    });
}

function getEmployees() {
  return dataRows_(SHEET.EMPLOYEES)
    .filter(function (r) { return String(r[0]).trim() && String(r[5]).trim().toUpperCase() !== 'N'; })
    .map(function (r) {
      return {
        id: String(r[0]).trim(),
        name: String(r[1]).trim(),
        centerCode: String(r[2]).trim(),
        phone: digits_(r[3]),
        pin: String(r[4]).trim()
      };
    });
}

/** 오늘(또는 지정일) 기록. 기록은 시간순으로 추가되므로 마지막 N행만 읽는다. */
var LOG_SCAN_ROWS = 6000;

function getLogsForDate(dateStr) {
  var sh = sheet_(SHEET.LOG);
  var last = sh.getLastRow();
  if (last < 2) return [];
  var start = Math.max(2, last - LOG_SCAN_ROWS + 1);
  var values = sh.getRange(start, 1, last - start + 1, HEADERS[SHEET.LOG].length).getDisplayValues();
  return values
    .filter(function (r) { return r[1] === dateStr; })
    .map(function (r) {
      return {
        employeeId: String(r[3]).trim(),
        action: r[7],
        minutes: parseHm(r[2].slice(0, 5)),
        time: r[2]
      };
    });
}

function todayStr_(d) {
  return Utilities.formatDate(d || new Date(), 'Asia/Seoul', 'yyyy-MM-dd');
}

function nowMinutes_(d) {
  var hm = Utilities.formatDate(d || new Date(), 'Asia/Seoul', 'HH:mm');
  return parseHm(hm);
}

function weekdayIndex_(d) {
  // 'u' = 1(월)~7(일)
  return Number(Utilities.formatDate(d || new Date(), 'Asia/Seoul', 'u')) % 7;
}
