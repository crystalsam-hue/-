/**
 * 시트 이름·구조 정의 및 데이터 읽기
 */

var SHEET = {
  SETTINGS: '설정',
  CENTERS: '센터',
  EMPLOYEES: '직원',
  LOG: '기록',
  LIVE: '실시간현황',
  LEAVE: '휴가',
  LEAVE_STATUS: '연차현황',
  HOLIDAYS: '공휴일'
};

var HEADERS = {};
HEADERS[SHEET.SETTINGS] = ['항목', '값', '설명'];
HEADERS[SHEET.CENTERS] = ['센터코드', '센터명', '센터장 이름', '센터장 이메일', '센터장 휴대폰', '위도', '경도', '허용반경(m)'];
HEADERS[SHEET.EMPLOYEES] = ['사번', '이름', '직급', '센터코드', '휴대폰', 'PIN(4~6자리)', '재직(Y/N)', '입사일', '연차 부여일수(수동)'];
HEADERS[SHEET.LOG] = ['기록ID', '일자', '시각', '사번', '이름', '센터코드', '센터명', '구분', '비고', '위도', '경도', '센터거리(m)', '위치판정'];
HEADERS[SHEET.LEAVE] = ['신청ID', '사번', '이름', '직급', '센터코드', '종류', '시작일', '종료일', '일수', '사유', '상태', '신청일시', '처리자', '처리일시', '처리메모'];
HEADERS[SHEET.HOLIDAYS] = ['날짜', '이름'];

/** [항목, 기본값, 설명] */
var DEFAULT_SETTINGS = [
  ['회장 이름', '회장님', ''],
  ['회장 이메일', '', '보고서를 받을 이메일 (여러 명이면 쉼표로 구분)'],
  ['회장 휴대폰', '', '알림톡 보고를 받을 번호 (예: 01012345678)'],
  ['출근 기준시각', '10:00', 'HH:MM. 이 시각까지 출근하면 정상'],
  ['지각 허용(분)', '0', '기준시각 + 허용분 이후 출근은 지각'],
  ['퇴근 기준시각', '19:00', 'HH:MM (기준 이전 퇴근은 "조퇴"로 표시)'],
  ['오전반차 출근기준', '15:00', '오전반차인 날의 출근 기준시각'],
  ['오후반차 퇴근기준', '14:00', '오후반차인 날의 퇴근 기준시각'],
  ['보고 시각', '10:10, 14:00, 19:30', '회장 보고서(엑셀 첨부) 발송 시각. 쉼표로 여러 개, 언제든 변경 가능'],
  ['근무 요일', '월,화,수,목,금', '이 요일에만 정기 보고 발송 · 휴가일수 계산 ([공휴일] 시트 날짜 제외)'],
  ['직급 순서', '센터장,부센터장,실장,팀장,과장,대리,주임,사원', '보고서 정렬 순서 · 직급 선택 목록'],
  ['이메일 알림', 'Y', 'Y/N - 센터장에게 실시간 이메일'],
  ['알림톡 알림', 'N', 'Y/N - 센터장에게 실시간 알림톡 (솔라피 설정 필요)'],
  ['회장 실시간 알림', 'N', 'Y/N - 모든 출퇴근 이벤트를 회장에게도 즉시 전송'],
  ['센터장 정기보고', 'Y', 'Y/N - 보고 시각마다 각 센터장에게 자기 센터 현황 전송'],
  ['위치 확인', 'N', 'Y/N - 휴대폰 GPS로 센터 반경 확인 (반경 밖이면 "범위밖" 표시)'],
  ['알림톡 발신번호', '', '솔라피에 등록된 발신번호'],
  ['카카오 채널 pfId', '', '솔라피에 연동한 카카오 비즈니스 채널 ID'],
  ['알림톡 템플릿(이벤트)', '', '검수 완료된 출퇴근 알림 템플릿 ID'],
  ['알림톡 템플릿(보고)', '', '검수 완료된 현황 보고 템플릿 ID'],
  ['알림톡 템플릿(휴가)', '', '검수 완료된 휴가 신청/처리 알림 템플릿 ID']
];

/** 초기 설정 시 미리 넣는 공휴일 (매년 [공휴일] 시트에 추가하세요) */
var DEFAULT_HOLIDAYS = [
  ['2026-10-09', '한글날'],
  ['2026-12-25', '성탄절']
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

/** 셀 값(문자열 또는 Date) → "yyyy-MM-dd" 또는 null */
function cellYmd_(v) {
  if (v instanceof Date) return Utilities.formatDate(v, 'Asia/Seoul', 'yyyy-MM-dd');
  return normalizeYmd(v);
}

function getSettings() {
  var map = {};
  dataRows_(SHEET.SETTINGS).forEach(function (r) { map[String(r[0]).trim()] = r[1]; });
  var get = function (k) {
    var v = map[k];
    if (v instanceof Date) return Utilities.formatDate(v, 'Asia/Seoul', 'HH:mm');
    return v === undefined || v === null ? '' : String(v).trim();
  };
  var dflt = {};
  DEFAULT_SETTINGS.forEach(function (r) { dflt[r[0]] = r[1]; });
  var getOr = function (k) { return get(k) || dflt[k]; };
  return {
    chairmanName: getOr('회장 이름'),
    chairmanEmail: get('회장 이메일'),
    chairmanPhone: digits_(get('회장 휴대폰')),
    startMinutes: parseHm(getOr('출근 기준시각')),
    graceMinutes: Number(get('지각 허용(분)') || 0),
    endMinutes: parseHm(getOr('퇴근 기준시각')),
    halfAmStartMinutes: parseHm(getOr('오전반차 출근기준')),
    halfPmEndMinutes: parseHm(getOr('오후반차 퇴근기준')),
    reportSlots: parseHmList(getOr('보고 시각')),
    workdays: getOr('근무 요일'),
    rankOrder: getOr('직급 순서').split(/\s*,\s*/).filter(String),
    emailOn: yn_(get('이메일 알림')),
    alimtalkOn: yn_(get('알림톡 알림')),
    chairmanRealtime: yn_(get('회장 실시간 알림')),
    managerReports: yn_(get('센터장 정기보고')),
    geoOn: yn_(get('위치 확인')),
    senderPhone: digits_(get('알림톡 발신번호')),
    pfId: get('카카오 채널 pfId'),
    eventTemplateId: get('알림톡 템플릿(이벤트)'),
    reportTemplateId: get('알림톡 템플릿(보고)'),
    leaveTemplateId: get('알림톡 템플릿(휴가)')
  };
}

/** 출근 판정 규칙 (Logic.computeDailyStatus / workRuleFor 용) */
function workRule_(settings) {
  return {
    startMinutes: settings.startMinutes,
    graceMinutes: settings.graceMinutes,
    endMinutes: settings.endMinutes,
    halfAmStartMinutes: settings.halfAmStartMinutes,
    halfPmEndMinutes: settings.halfPmEndMinutes,
    rankOrder: settings.rankOrder
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
    .filter(function (r) { return String(r[0]).trim() && String(r[6]).trim().toUpperCase() !== 'N'; })
    .map(function (r) {
      return {
        id: String(r[0]).trim(),
        name: String(r[1]).trim(),
        rank: String(r[2]).trim(),
        centerCode: String(r[3]).trim(),
        phone: digits_(r[4]),
        pin: String(r[5]).trim(),
        hireDate: cellYmd_(r[7]),
        manualLeaveDays: String(r[8]).trim()
      };
    });
}

/** 이름 + 직급 (예: "홍길동 과장") */
function displayName_(emp) {
  return emp.rank ? emp.name + ' ' + emp.rank : emp.name;
}

/** 휴가 신청 목록 (row: 시트 행 번호) */
function getLeaves() {
  return dataRows_(SHEET.LEAVE)
    .map(function (r, i) {
      return {
        row: i + 2,
        id: String(r[0]).trim(),
        employeeId: String(r[1]).trim(),
        name: String(r[2]).trim(),
        rank: String(r[3]).trim(),
        centerCode: String(r[4]).trim(),
        type: String(r[5]).trim(),
        start: cellYmd_(r[6]),
        end: cellYmd_(r[7]),
        days: Number(r[8]) || 0,
        reason: String(r[9]),
        status: String(r[10]).trim(),
        requestedAt: String(r[11]),
        decidedBy: String(r[12]),
        decidedAt: String(r[13]),
        memo: String(r[14])
      };
    })
    .filter(function (l) { return l.id && l.start && l.end; });
}

function getHolidays() {
  var sh = ss_().getSheetByName(SHEET.HOLIDAYS);
  if (!sh || sh.getLastRow() < 2) return [];
  return sh.getRange(2, 1, sh.getLastRow() - 1, 1).getValues()
    .map(function (r) { return cellYmd_(r[0]); })
    .filter(String);
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

function nowStr_(d) {
  return Utilities.formatDate(d || new Date(), 'Asia/Seoul', 'yyyy-MM-dd HH:mm');
}

function nowMinutes_(d) {
  var hm = Utilities.formatDate(d || new Date(), 'Asia/Seoul', 'HH:mm');
  return parseHm(hm);
}

function weekdayIndex_(d) {
  // 'u' = 1(월)~7(일)
  return Number(Utilities.formatDate(d || new Date(), 'Asia/Seoul', 'u')) % 7;
}
