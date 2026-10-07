// Apps Script 서비스(시트·메일·락 등)를 메모리로 흉내 내어 실제 흐름을 검증
const test = require('node:test');
const assert = require('node:assert');
const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');
const crypto = require('node:crypto');

function a1ToRange(a1) {
  // "A:H", "C2:C" 만 지원
  const m = a1.match(/^([A-Z])(\d*):([A-Z])(\d*)$/);
  const col = (c) => c.charCodeAt(0) - 64;
  return { row: Number(m[2] || 1), col: col(m[1]), rows: 1000, cols: col(m[3]) - col(m[1]) + 1 };
}

class Sheet {
  constructor(name) { this.name = name; this.data = []; }
  getName() { return this.name; }
  setName(n) { this.name = n; return this; }
  getLastRow() {
    for (let i = this.data.length - 1; i >= 0; i--) if (this.data[i] && this.data[i].some((v) => v !== '' && v !== undefined)) return i + 1;
    return 0;
  }
  getLastColumn() { return Math.max(0, ...this.data.map((r) => (r ? r.length : 0))); }
  getRange(r, c, nr, nc) {
    if (typeof r === 'string') { const x = a1ToRange(r); return new Range(this, x.row, x.col, x.rows, x.cols); }
    return new Range(this, r, c, nr || 1, nc || 1);
  }
  appendRow(row) { this.data[this.getLastRow()] = Array.from(row); return this; }
  clear() { this.data = []; return this; }
  setFrozenRows() { return this; }
  autoResizeColumns() { return this; }
}

class Range {
  constructor(sh, r, c, nr, nc) { Object.assign(this, { sh, r, c, nr, nc }); }
  getValues() {
    const out = [];
    for (let i = 0; i < this.nr; i++) {
      const row = this.sh.data[this.r - 1 + i] || [];
      out.push(Array.from({ length: this.nc }, (_, j) => (row[this.c - 1 + j] === undefined ? '' : row[this.c - 1 + j])));
    }
    return out;
  }
  getDisplayValues() { return this.getValues().map((r) => r.map(String)); }
  setValues(v) {
    v.forEach((row, i) => {
      const target = (this.sh.data[this.r - 1 + i] = this.sh.data[this.r - 1 + i] || []);
      row.forEach((x, j) => { target[this.c - 1 + j] = x; });
    });
    return this;
  }
  setValue(x) { return this.setValues([[x]]); }
}
['setFontSize', 'setFontWeight', 'setBackground', 'setFontColor', 'setBackgrounds', 'setNumberFormat',
  'setDataValidation', 'setNote'].forEach((m) => { Range.prototype[m] = function () { return this; }; });

function makeEnv() {
  const sheets = {};
  const mails = [];
  const fetches = [];
  const props = {};
  const cache = {};
  const ss = {
    getSheetByName: (n) => sheets[n] || null,
    insertSheet: (n) => (sheets[n] = new Sheet(n)),
    getUrl: () => 'https://sheet.example/ss',
    getId: () => 'SSID',
    getSheets: () => Object.values(sheets),
  };
  const fmt = (d, tz, pattern) => {
    const k = new Date(d.getTime() + 9 * 3600 * 1000); // Asia/Seoul
    const p = (n) => String(n).padStart(2, '0');
    const map = {
      yyyy: k.getUTCFullYear(), MM: p(k.getUTCMonth() + 1), dd: p(k.getUTCDate()),
      HH: p(k.getUTCHours()), mm: p(k.getUTCMinutes()), ss: p(k.getUTCSeconds()), u: k.getUTCDay() || 7,
    };
    return pattern.replace(/yyyy|MM|dd|HH|mm|ss|u/g, (t) => map[t]);
  };
  const ctx = {
    console: { log() {}, warn() {}, error() {} },
    SpreadsheetApp: {
      getActiveSpreadsheet: () => ss,
      flush() {},
      newDataValidation: () => {
        const b = { requireValueInList: () => b, setAllowInvalid: () => b, build: () => ({}) };
        return b;
      },
    },
    Utilities: {
      formatDate: fmt,
      getUuid: () => crypto.randomUUID(),
      computeHmacSha256Signature: (v, k) => Array.from(crypto.createHmac('sha256', k).update(v).digest()).map((b) => (b > 127 ? b - 256 : b)),
    },
    LockService: { getScriptLock: () => ({ waitLock() {}, releaseLock() {} }) },
    CacheService: { getScriptCache: () => ({ get: (k) => cache[k] || null, put: (k, v) => { cache[k] = v; }, remove: (k) => { delete cache[k]; } }) },
    PropertiesService: {
      getScriptProperties: () => ({
        getProperty: (k) => props[k] || null, setProperty: (k, v) => { props[k] = v; },
        getProperties: () => ({ ...props }), deleteProperty: (k) => { delete props[k]; },
      }),
    },
    MailApp: { sendEmail: (o) => mails.push(o) },
    UrlFetchApp: { fetch: (url, o) => { fetches.push({ url, o }); return { getResponseCode: () => 200, getContentText: () => '{}' }; } },
    ScriptApp: { getService: () => ({ getUrl: () => 'https://script.example/exec' }) },
  };
  vm.createContext(ctx);
  const dir = path.join(__dirname, '..', 'src');
  const code = fs.readdirSync(dir).filter((f) => f.endsWith('.gs')).map((f) => fs.readFileSync(path.join(dir, f), 'utf8')).join('\n');
  vm.runInContext(code, ctx);

  // 시트 준비 (setup() 은 UI 를 쓰므로 직접 구성)
  const mk = (name, rows) => { const sh = ss.insertSheet(name); sh.data = [vm.runInContext(`HEADERS['${name}']`, ctx) || []].concat(rows); return sh; };
  mk('설정', vm.runInContext('DEFAULT_SETTINGS', ctx).map((r) => r.slice()));
  const set = (k, v) => { sheets['설정'].data.find((r) => r[0] === k)[1] = v; };
  set('회장 이메일', 'chair@x.com');
  set('회장 휴대폰', '010-9999-9999');
  set('근무 요일', '월,화,수,목,금,토,일'); // 테스트가 요일에 영향받지 않게
  mk('센터', [['C01', '서울센터', '김센터', 'kim@x.com', '01011112222', '', '', 300],
    ['C02', '부산센터', '이센터', 'lee@x.com', '01033334444', '', '', 300]]);
  mk('직원', [
    ['E1', '김센터', '센터장', 'C01', '01011112222', '1111', 'Y', '2018-01-01', ''],
    ['E2', '박사원', '사원', 'C01', '01055556666', '2222', 'Y', '2025-01-01', ''],
    ['E3', '최팀장', '팀장', 'C02', '01077778888', '3333', 'Y', '2020.5.1', ''],
    ['E4', '퇴사자', '사원', 'C02', '', '4444', 'N', '', ''],
  ]);
  mk('기록', []);
  mk('휴가', []);
  mk('공휴일', []);
  return { ctx, sheets, mails, fetches, set, call: (expr) => vm.runInContext(expr, ctx) };
}

test('출근 기록 → 센터장 메일 → 실시간현황 갱신', () => {
  const env = makeEnv();
  const st = env.ctx.apiStatus('E2', '2222');
  assert.strictEqual(st.name, '박사원 사원');
  assert.strictEqual(st.state, '미출근');

  const r = env.ctx.apiRecord('E2', '2222', '출근', '', null, null);
  assert.match(r.message, /박사원 사원님 출근 완료/);
  assert.strictEqual(env.sheets['기록'].data.length, 2);
  assert.strictEqual(env.sheets['기록'].data[1][7], '출근');
  assert.strictEqual(env.mails[0].to, 'kim@x.com');
  assert.match(env.mails[0].subject, /\[서울센터\] 박사원 사원 출근/);

  assert.throws(() => env.ctx.apiRecord('E2', '2222', '출근', '', null, null), /이미 출근/);
  assert.throws(() => env.ctx.apiRecord('E2', '0000', '퇴근', '', null, null), /PIN/);

  const live = env.sheets['실시간현황'].data;
  const header = live.find((row) => row && row[0] === '센터' && row[1] === '센터장');
  assert.deepStrictEqual(header, ['센터', '센터장', '전체', '출근', '지각', '근무중', '외출중', '퇴근', '휴가', '미출근']);
  const seoul = live.find((row) => row && row[0] === '서울센터');
  assert.strictEqual(seoul[2], 2); // 전체 (퇴사자 제외)
  assert.strictEqual(seoul[3], 1); // 출근
  const total = live.find((row) => row && row[0] === '합계');
  assert.strictEqual(total[2], 3);
});

test('휴가 신청 → 센터장 승인 링크 → 승인 → 현황/연차 반영', () => {
  const env = makeEnv();
  const today = env.call('todayStr_()');

  const info = env.ctx.apiLeaveInfo('E2', '2222');
  assert.strictEqual(info.granted, 15);
  assert.strictEqual(info.remaining, 15);

  const res = env.ctx.apiRequestLeave('E2', '2222', '연차', today, today, '개인 사정');
  assert.match(res.message, /신청 완료/);
  const row = env.sheets['휴가'].data[1];
  assert.deepStrictEqual(row.slice(1, 11), ['E2', '박사원', '사원', 'C01', '연차', today, today, 1, '개인 사정', '신청']);

  const mail = env.mails.find((m) => m.subject.startsWith('[휴가신청]'));
  assert.strictEqual(mail.to, 'kim@x.com'); // 사원 → 센터장
  const link = mail.body.match(/https:\/\/script\.example\/exec\?leave=([^&]+)&t=(\w+)/);
  assert.ok(link, mail.body);

  assert.throws(() => env.ctx.apiRequestLeave('E2', '2222', '연차', today, today, ''), /겹칩니다/);
  assert.strictEqual(env.ctx.apiLeaveInfo('E2', '2222').pending, 1);

  assert.throws(() => env.ctx.apiDecideLeave(link[1], 'bad', '승인', ''), /유효하지 않은/);
  const d = env.ctx.apiDecideLeave(decodeURIComponent(link[1]), link[2], '승인', '');
  assert.match(d.message, /승인 처리/);
  assert.strictEqual(env.sheets['휴가'].data[1][10], '승인');
  assert.match(env.sheets['휴가'].data[1][12], /서울센터 센터장 김센터/);
  assert.throws(() => env.ctx.apiDecideLeave(decodeURIComponent(link[1]), link[2], '반려', ''), /이미 처리/);

  const after = env.ctx.apiLeaveInfo('E2', '2222');
  assert.deepStrictEqual([after.used, after.pending, after.remaining], [1, 0, 14]);
  assert.strictEqual(env.ctx.apiStatus('E2', '2222').todayLeave, '연차');

  const live = env.sheets['실시간현황'].data;
  const seoul = live.find((r) => r && r[0] === '서울센터' && typeof r[2] === 'number');
  assert.strictEqual(seoul[8], 1); // 휴가
  assert.strictEqual(seoul[9], 1); // 미출근 (센터장만)
  const detail = live.find((r) => r && r[2] === 'E2');
  assert.strictEqual(detail[4], '휴가');
  assert.strictEqual(detail[5], '연차');

  const status = env.sheets['연차현황'].data;
  const e3 = status.find((r) => r && r[2] === 'E3');
  assert.strictEqual(e3[4], '2020-05-01'); // "2020.5.1" 정규화
  assert.strictEqual(e3[7], 17); // 근속 6년: 15 + floor(5/2)
});

test('센터장 본인 휴가는 회장에게 승인 요청, 취소 가능', () => {
  const env = makeEnv();
  const today = env.call('todayStr_()');
  env.ctx.apiRequestLeave('E1', '1111', '오전반차', today, today, '');
  const mail = env.mails.find((m) => m.subject.startsWith('[휴가신청]'));
  assert.strictEqual(mail.to, 'chair@x.com');
  const id = env.sheets['휴가'].data[1][0];
  env.ctx.apiCancelLeave('E1', '1111', id);
  assert.strictEqual(env.sheets['휴가'].data[1][10], '취소');
  assert.throws(() => env.ctx.apiCancelLeave('E1', '1111', id), /취소할 수 없습니다/);
});

test('정기 보고: 회장 엑셀 첨부 + 센터장 보고, 하루 한 번만', () => {
  const env = makeEnv();
  env.set('보고 시각', '00:00');
  // 엑셀 내보내기는 Drive 가 필요하므로 대체
  env.call(`exportXlsx_ = function (ctx, title) { return { name: 'x.xlsx', title: title }; }`);
  env.ctx.scheduledReportTick();
  const chair = env.mails.find((m) => m.to === 'chair@x.com');
  assert.ok(chair, 'chairman mail');
  assert.match(chair.subject, /\[근태보고\] .* 00:00 전 센터 근태 현황 - 출근 0\/3, 지각 0, 휴가 0, 미출근 3/);
  assert.strictEqual(chair.attachments.length, 1);
  assert.ok(env.mails.some((m) => m.to === 'kim@x.com' && /서울센터 근태 현황/.test(m.subject)));
  const count = env.mails.length;
  env.ctx.scheduledReportTick();
  assert.strictEqual(env.mails.length, count);
});

test('알림톡: 솔라피 요청 형식', () => {
  const env = makeEnv();
  env.set('알림톡 알림', 'Y');
  env.set('알림톡 발신번호', '0212345678');
  env.set('카카오 채널 pfId', 'PF1');
  env.set('알림톡 템플릿(이벤트)', 'TPL_EVT');
  env.call(`PropertiesService.getScriptProperties().setProperty('SOLAPI_API_KEY', 'K'); PropertiesService.getScriptProperties().setProperty('SOLAPI_API_SECRET', 'S');`);
  env.ctx.apiRecord('E3', '3333', '출근', '', null, null);
  assert.strictEqual(env.fetches.length, 1);
  const { url, o } = env.fetches[0];
  assert.strictEqual(url, 'https://api.solapi.com/messages/v4/send');
  assert.match(o.headers.Authorization, /^HMAC-SHA256 apiKey=K, date=\S+, salt=\w+, signature=[0-9a-f]{64}$/);
  const msg = JSON.parse(o.payload).message;
  assert.strictEqual(msg.to, '01033334444');
  assert.strictEqual(msg.kakaoOptions.templateId, 'TPL_EVT');
  assert.strictEqual(msg.kakaoOptions.variables['#{이름}'], '최팀장 팀장');
});
