// 핵심 로직 스모크 테스트: node test/logic-smoke.js
// 시트/Gmail/Properties/Cache 를 목으로 대체해 인증 토큰과 점주 링크 확정(confirmBookingByLink) 규칙을 검증한다.
const fs = require('fs');
const path = require('path');
const vm = require('vm');
const assert = require('assert');

// ---- fakes -----------------------------------------------------------------
function fakeSheet(rows) { // rows: 1-based sheet rows as arrays (index 0 = column A)
  const cells = rows.map(r => r.slice());
  const get = (r, c) => (cells[r - 1] || [])[c - 1] ?? '';
  const set = (r, c, v) => { while (cells.length < r) cells.push([]); cells[r - 1][c - 1] = v; };
  return {
    cells,
    getDataRange: () => ({ getValues: () => cells.map(r => r.slice()) }),
    getLastRow: () => cells.length,
    appendRow: row => { cells.push(row.slice()); },
    getRange: (r, c, nr = 1, nc = 1) => ({
      getValue: () => get(r, c),
      setValue: v => set(r, c, v),
      clearContent: () => { for (let i = 0; i < nr; i++) for (let j = 0; j < nc; j++) set(r + i, c + j, ''); },
      getValues: () => Array.from({ length: nr }, (_, i) => Array.from({ length: nc }, (_, j) => get(r + i, c + j))),
    }),
  };
}
const store = {};
const props = { getProperties: () => ({ ...store }), setProperty: (k, v) => { store[k] = v; }, getProperty: k => store[k] ?? null, deleteProperty: k => { delete store[k]; } };
const cacheStore = {};
const cache = { get: k => cacheStore[k] ?? null, put: (k, v) => { cacheStore[k] = v; }, remove: k => { delete cacheStore[k]; } };

let sheets, sent;
const sandbox = {
  console,
  Date,
  GmailApp: { sendEmail: (to, subject, body, opts) => sent.push({ to, subject, opts }) },
  PropertiesService: { getScriptProperties: () => props },
  CacheService: { getScriptCache: () => cache },
  Utilities: {
    getUuid: () => require('crypto').randomUUID(),
    formatDate: (d, tz, f) => {
      const p = n => String(n).padStart(2, '0');
      return f.includes('HH') ? `${d.getFullYear()}-${p(d.getMonth() + 1)}-${p(d.getDate())} ${p(d.getHours())}:${p(d.getMinutes())}` : `${d.getFullYear()}-${p(d.getMonth() + 1)}-${p(d.getDate())}`;
    },
  },
  Session: { getScriptTimeZone: () => 'Asia/Seoul' },
  SpreadsheetApp: { getActiveSpreadsheet: () => ({ getSheetByName: n => sheets[n] || null }), flush: () => {} },
  LockService: { getScriptLock: () => ({ tryLock: () => true, releaseLock: () => {} }) },
  ContentService: { MimeType: { JSON: 'json' }, createTextOutput: s => ({ text: s, setMimeType() { return this; } }) },
  Logger: { log: () => {} },
  ScriptApp: { getService: () => ({ getUrl: () => 'https://example.test/exec' }) },
};
vm.createContext(sandbox);
['Config.js', 'Utils.js', 'Service_Auth.js', 'Service_Mail.js', 'Service_Store.js', 'Service_Admin.js', 'Service_Creator.js', 'Service_Imweb.js'].forEach(f =>
  vm.runInContext(fs.readFileSync(path.join(__dirname, '..', f), 'utf8'), sandbox, { filename: f }));
const run = code => vm.runInContext(code, sandbox);

// ---- 1. auth tokens --------------------------------------------------------
const tok = run(`issueToken_('admin','admin')`);
assert.strictEqual(run(`requireToken_('${tok}','admin')`), 'admin', 'valid token');
assert.throws(() => run(`requireToken_('${tok}','store')`), /AUTH_REQUIRED/, 'scope mismatch');
assert.throws(() => run(`requireToken_('nope','admin')`), /AUTH_REQUIRED/, 'unknown token');
assert.throws(() => run(`requireToken_('','admin')`), /AUTH_REQUIRED/, 'empty token');
store['TOKEN_' + tok] = JSON.stringify({ scope: 'admin', subject: 'admin', exp: Date.now() - 1 });
assert.throws(() => run(`requireToken_('${tok}','admin')`), /AUTH_REQUIRED/, 'expired token');
run(`issueToken_('admin','admin')`);
assert.ok(!('TOKEN_' + tok in store), 'expired token purged on next issue');
for (let i = 0; i < 10; i++) run(`recordLoginFailure_('admin')`);
assert.strictEqual(run(`isLoginLocked_('admin')`), true, 'admin lock after 10 failures');
run(`clearLoginFailures_('admin')`);
assert.strictEqual(run(`isLoginLocked_('admin')`), false);

const storeTok = run(`issueToken_('store','bookmark-0001')`);
assert.doesNotThrow(() => run(`requireStoreAccess_('${storeTok}','BOOKMARK-0001')`), 'store id compare is case-insensitive');
assert.throws(() => run(`requireStoreAccess_('${storeTok}','BOOKMARK-0002')`), /AUTH_REQUIRED/, 'other store rejected');

// ---- 2. confirmBookingByLink ----------------------------------------------
const fmt = d => sandbox.Utilities.formatDate(d, 'Asia/Seoul', 'yyyy-MM-dd HH:mm');
const hoursAgo = h => fmt(new Date(Date.now() - h * 3600e3));
function setup(status, logAge) {
  sent = [];
  const master = [[], []]; // 2 header rows
  const row = new Array(25).fill('');
  row[0] = "'1001"; row[1] = 'mc1'; row[2] = 'Alice'; row[4] = 'S1'; row[5] = '테스트매장';
  row[7] = '2026-12-01 12:00'; row[8] = '2명'; row[11] = status;
  row[24] = logAge === null ? '' : `[${hoursAgo(logAge)}] 점주메일 발송완료`;
  master.push(row);
  const restRow = new Array(30).fill(''); restRow[0] = 'S1'; restRow[1] = '테스트매장'; restRow[2] = 'テスト店'; restRow[25] = 0;
  const userRow = new Array(16).fill(''); userRow[0] = 'mc1'; userRow[13] = 'alice@example.com';
  sheets = { Master_Log: fakeSheet(master), Restaurant_List: fakeSheet([[], [], restRow]), User_DB: fakeSheet([[], userRow]), Settings: null };
}
const call = () => run(`confirmBookingByLink({ row: '3', o: '1001' })`);
const statusOf = () => sheets.Master_Log.cells[2][11];

setup('예약확인중', 1);
let r = call();
assert.deepStrictEqual(JSON.parse(JSON.stringify(r)), { isExpired: false, isAlreadyRequested: false });
assert.strictEqual(statusOf(), '방문전', 'fresh pending → confirmed');
assert.ok(sent.some(m => m.to === 'alice@example.com'), 'creator mail sent');

setup('예약확인중', 30);
r = call();
assert.strictEqual(r.isExpired, true);
assert.strictEqual(statusOf(), '취소(기한만료)', '30h-old pending → expired');
assert.strictEqual(sheets.Restaurant_List.cells[2][25], 1, 'no-response counter +1');
assert.strictEqual(sent.length, 0, 'no mail on expiry');

for (const st of ['방문전', '방문완료', '제출완료']) {
  setup(st, 100);
  r = call();
  assert.strictEqual(r.isExpired, false, st + ' must not expire');
  assert.strictEqual(statusOf(), st, st + ' must stay unchanged');
  assert.strictEqual(sent.length, 0, st + ' must not re-send mail');
}

setup('일정조율필요', 1);
assert.strictEqual(call().isAlreadyRequested, true);
setup('취소(운영자)', 1);
assert.strictEqual(call().isAlreadyRequested, true);

// ---- 3. sendDailyStoreReminders -------------------------------------------
const todayDate = fmt(new Date()).substring(0, 10);
function remindSetup(rows) { // rows: [storeId, status, visit, people, name, log]
  sent = [];
  const master = [[], []];
  rows.forEach(([sid, st, visit, ppl, name, log], n) => {
    const row = new Array(25).fill('');
    row[0] = "'" + (2000 + n); row[2] = name; row[4] = sid; row[7] = visit; row[8] = ppl; row[11] = st; row[24] = log || '';
    master.push(row);
  });
  const restRow = (id, jp, email) => { const r = new Array(30).fill(''); r[0] = id; r[2] = jp; r[10] = email; return r; };
  sheets = { Master_Log: fakeSheet(master), Restaurant_List: fakeSheet([[], [], restRow('S1', 'テスト店', 's1@example.com'), restRow('S2', '店2', ''), restRow('S3', '店3', 's3@example.com')]), Settings: null };
}
remindSetup([
  ['S1', '방문전', `${todayDate} 18:30`, '4명', 'Bob'],
  ['S1', '방문전', `${todayDate} 12:00`, '2명', 'Alice'],
  ['S1', '예약확인중', `${todayDate} 13:00`, '2명', 'Pending'],
  ['S1', '방문전', '2000-01-01 12:00', '2명', 'OtherDay'],
  ['S2', '방문전', `${todayDate} 12:00`, '2명', 'NoEmail'],
  ['S3', '방문전', `${todayDate} 12:00`, '1명', 'Done', `[${todayDate} 09:00] 당일 리마인드 발송`],
]);
run(`sendDailyStoreReminders()`);
assert.strictEqual(sent.length, 1, 'S1 1통만 발송 (S2 이메일 없음, S3 이미 발송)');
assert.strictEqual(sent[0].to, 's1@example.com');
assert.ok(/2件: 12:00 2名 \/ 18:30 4名/.test(sent[0].subject), '시간순 정렬 + 제목에 시간·인원: ' + sent[0].subject);
assert.ok(!/Pending|OtherDay/.test(sent[0].opts.htmlBody), '방문전·오늘만 포함');
const logAt = n => sheets.Master_Log.cells[n][24];
assert.ok(/당일 리마인드 발송/.test(logAt(2)) && /당일 리마인드 발송/.test(logAt(3)), '발송 건 로그 기록');
assert.ok(/실패: 점주 이메일 주소 없음/.test(logAt(6)), '이메일 없는 점포 실패 로그');
assert.strictEqual(logAt(4), '', '예약확인중 건은 건드리지 않음');
sent = [];
run(`sendDailyStoreReminders()`);
assert.strictEqual(sent.length, 0, '재실행 시 중복 발송 없음');

// ---- 4. 숙박(STAY) 예약 ----------------------------------------------------
const ymd = n => fmt(new Date(Date.now() + n * 864e5)).substring(0, 10);
function staySetup(blackouts, bookings) { // bookings: [status, checkIn, checkOut]
  sent = [];
  const master = [[], []];
  const base = new Array(25).fill('');
  base[0] = "'3001"; base[1] = 'mc1'; base[2] = 'Alice'; base[4] = 'H1'; base[11] = '예약대기';
  master.push(base);
  bookings.forEach(([st, cin, cout], n) => {
    const row = new Array(25).fill('');
    row[0] = "'" + (3100 + n); row[4] = 'H1'; row[7] = cin + ' 15:00'; row[11] = st; row[21] = cout;
    master.push(row);
  });
  const rest = new Array(35).fill('');
  rest[0] = 'H1'; rest[1] = '호텔'; rest[2] = 'ホテル'; rest[7] = 'STAY'; rest[10] = 'h1@example.com';
  rest[13] = blackouts; rest[21] = 3; rest[23] = 1; rest[24] = 90; rest[30] = 1; rest[33] = 3;
  const userRow = new Array(16).fill(''); userRow[0] = 'mc1'; userRow[13] = 'alice@example.com';
  sheets = { Master_Log: fakeSheet(master), Restaurant_List: fakeSheet([[], [], rest]), User_DB: fakeSheet([[], userRow]), Settings: null };
}
const stay = (a, b) => JSON.parse(JSON.stringify(run(`checkStayAvailability('H1', '${a}', '${b}')`)));

staySetup('', []);
let sr = stay(ymd(10), ymd(12));
assert.strictEqual(sr.isAvailable, true, '기본 2박 가능');
assert.strictEqual(sr.nights, 2);
assert.strictEqual(sr.checkinTime, '15:00', '체크인 기본 시간');
assert.strictEqual(stay(ymd(10), ymd(14)).isAvailable, false, '최대 3박 초과 거절');
assert.strictEqual(stay(ymd(10), ymd(10)).isAvailable, false, '0박 거절');
assert.strictEqual(stay(ymd(0), ymd(2)).isAvailable, false, '최소 사전 예약일(1일) 미달 거절');

staySetup(ymd(11), []);
assert.strictEqual(stay(ymd(10), ymd(12)).isAvailable, false, '숙박하는 밤이 휴무일이면 거절');
assert.strictEqual(stay(ymd(10), ymd(11)).isAvailable, true, '체크아웃일이 휴무일이어도 허용');

staySetup('', [['방문전', ymd(10), ymd(12)], ['취소(운영자)', ymd(20), ymd(22)], ['노쇼', ymd(30), ymd(32)]]);
assert.strictEqual(stay(ymd(11), ymd(13)).isAvailable, false, 'ROOMS=1 이고 겹치는 밤이 있으면 거절');
assert.strictEqual(stay(ymd(12), ymd(14)).isAvailable, true, '앞 예약의 체크아웃일 = 새 체크인일 허용');
assert.strictEqual(stay(ymd(20), ymd(22)).isAvailable, true, '취소 건은 객실 점유하지 않음');
assert.strictEqual(stay(ymd(30), ymd(32)).isAvailable, true, '노쇼 건은 객실 점유하지 않음');

staySetup('', []);
let br = JSON.parse(JSON.stringify(run(`bookStay(3, '${ymd(10)}', '${ymd(12)}', 4)`)));
assert.strictEqual(br.success, false, '최대 인원(3명) 초과 거절');
assert.strictEqual(sheets.Master_Log.cells[2][11], '예약대기', '거절 시 상태 불변');
br = JSON.parse(JSON.stringify(run(`bookStay(3, '${ymd(10)}', '${ymd(12)}', 2)`)));
assert.strictEqual(br.success, true, 'bookStay 성공: ' + JSON.stringify(br));
const stayRow = sheets.Master_Log.cells[2];
assert.strictEqual(stayRow[11], '예약확인중');
assert.strictEqual(stayRow[7], ymd(10) + ' 15:00');
assert.strictEqual(stayRow[8], '2명');
assert.strictEqual(stayRow[21], ymd(12));
assert.strictEqual(stayRow[22], 2);
assert.ok(sent.some(m => m.to === 'h1@example.com' && /ご宿泊予約のご確認 \(.*~.*\(2泊3日\) \/ 2名\)/.test(m.subject)), '점주 메일 제목에 숙박 일정: ' + sent.map(m => m.subject));
assert.strictEqual(JSON.parse(JSON.stringify(run(`bookStay(3, '${ymd(10)}', '${ymd(12)}', 1)`))).success, false, '이미 예약확인중인 행은 재예약 불가');

// ---- 5. 기존 시간 슬롯 예약 회귀 (휴무 판정 공통화 / 최대 인원 검증 / getMissions) ----
function slotSetup(blackouts) {
  sent = [];
  const master = [[], []];
  const base = new Array(25).fill('');
  base[0] = "'4001"; base[1] = 'mc1'; base[2] = 'Alice'; base[4] = 'R1'; base[11] = '예약대기';
  master.push(base);
  const rest = new Array(35).fill('');
  rest[0] = 'R1'; rest[1] = '식당'; rest[2] = '食堂'; rest[7] = 'RESTAURANT'; rest[10] = 'r1@example.com';
  rest[11] = 1; rest[12] = 2; rest[13] = blackouts; rest[14] = '12:00-14:00'; rest[15] = 60; rest[21] = 2;
  const userRow = new Array(16).fill(''); userRow[0] = 'mc1'; userRow[3] = '01012345678'; userRow[13] = 'alice@example.com';
  sheets = { Master_Log: fakeSheet(master), Restaurant_List: fakeSheet([[], [], rest]), User_DB: fakeSheet([[], userRow]), Settings: null };
}
const slotDate = ymd(10);
const dow = new Date(slotDate + 'T00:00:00').getDay();
const dowJa = ['日', '月', '火', '水', '木', '金', '土'][dow];
const slots = d => JSON.parse(JSON.stringify(run(`checkAvailability('R1', '${d}', '')`)));

slotSetup('');
assert.deepStrictEqual(slots(slotDate).availableSlots, ['12:00', '13:00'], '휴무 없으면 슬롯 생성');
slotSetup(slotDate);
assert.strictEqual(slots(slotDate).isAvailable, false, '날짜 지정 휴무');
slotSetup(dowJa);
assert.strictEqual(slots(slotDate).isAvailable, false, '요일 휴무');
slotSetup(`第${Math.ceil(new Date(slotDate + 'T00:00:00').getDate() / 7)}${dowJa}`);
assert.strictEqual(slots(slotDate).isAvailable, false, '第N요일 휴무');
slotSetup(`第${Math.ceil(new Date(slotDate + 'T00:00:00').getDate() / 7) % 5 + 1}${dowJa}`);
assert.strictEqual(slots(slotDate).isAvailable, true, '다른 주차 第N요일은 영업');

// 예약 불가 시간: 요일 접두는 해당 요일에만, 접두 없으면 매일
const otherJa = ['日', '月', '火', '水', '木', '金', '土'][(dow + 1) % 7];
slotSetup(''); sheets.Restaurant_List.cells[2][16] = `${dowJa}12:00`;
assert.deepStrictEqual(slots(slotDate).availableSlots, ['13:00'], '해당 요일 접두 차단');
slotSetup(''); sheets.Restaurant_List.cells[2][16] = `${otherJa}12:00`;
assert.deepStrictEqual(slots(slotDate).availableSlots, ['12:00', '13:00'], '다른 요일 접두는 차단 안 함');
slotSetup(''); sheets.Restaurant_List.cells[2][16] = `${otherJa}${dowJa}12:00-12:30,13:00`;
assert.ok(slots(slotDate).error, '여러 요일 접두 범위 + 접두 없는 항목(매일)으로 전 슬롯 차단');
slotSetup(''); sheets.Restaurant_List.cells[2][16] = `${otherJa}${dowJa}12:00-12:30,${otherJa}13:00`;
assert.deepStrictEqual(slots(slotDate).availableSlots, ['13:00'], '여러 요일 접두 범위(해당 요일 포함)');
slotSetup(''); sheets.Restaurant_List.cells[2][16] = '13:00';
assert.deepStrictEqual(slots(slotDate).availableSlots, ['12:00'], '접두 없는 기존 표기 유지');

slotSetup('');
let sb = JSON.parse(JSON.stringify(run(`bookTimeSlot(3, '${slotDate}', '12:00', 3)`)));
assert.strictEqual(sb.success, false, '최대 인원(2명) 초과 거절');
assert.strictEqual(sheets.Master_Log.cells[2][11], '예약대기');
sb = JSON.parse(JSON.stringify(run(`bookTimeSlot(3, '${slotDate}', '12:00', 2)`)));
assert.strictEqual(sb.success, true, 'bookTimeSlot 정상 예약: ' + JSON.stringify(sb));
assert.strictEqual(sheets.Master_Log.cells[2][11], '예약확인중');
assert.ok(sent.some(m => m.to === 'r1@example.com' && /ご来店予約のご確認 \(.* 12:00 \/ 2名\)/.test(m.subject)), '기존 점주 메일 제목 유지: ' + sent.map(m => m.subject));
assert.ok(!sheets.Master_Log.cells[2][21] && !sheets.Master_Log.cells[2][22], '일반 예약은 V·W열을 건드리지 않음');

const missions = JSON.parse(JSON.stringify(run(`getMissions('4001', '5678')`)));
assert.strictEqual(missions.length, 1, 'getMissions 조회');
assert.strictEqual(missions[0].info.storeType, 'RESTAURANT');
assert.strictEqual(missions[0].checkoutDate, '', '일반 예약은 checkoutDate 비어 있음');
assert.strictEqual(missions[0].info.checkinTime, undefined, '일반 매장 info 에 STAY 필드 없음');

staySetup('', []);
sheets.User_DB = fakeSheet([[], (() => { const u = new Array(16).fill(''); u[0] = 'mc1'; u[3] = '01012345678'; return u; })()]);
sheets.Master_Log.cells[2][0] = '5001';
run(`bookStay(3, '${ymd(10)}', '${ymd(12)}', 2)`);
const sm = JSON.parse(JSON.stringify(run(`getMissions('5001', '5678')`)))[0];
assert.strictEqual(sm.checkoutDate, ymd(12), 'STAY getMissions 체크아웃일');
assert.strictEqual(sm.nights, 2);
assert.strictEqual(sm.info.checkoutTime, '11:00');
assert.strictEqual(sm.info.bookingUrl, '', 'STAY 는 외부 예약 URL 없음');

// ---- 6. 숙박 확정 메일 언어 / 금·토 휴무 ----------------------------------
staySetup('', []);
run(`bookStay(3, '${ymd(10)}', '${ymd(11)}', 2)`);
sent = [];
run(`confirmBookingByLink({ row: '3', o: '3001' })`);
const creatorMail = sent.find(m => m.to === 'alice@example.com');
const adminMail = sent.find(m => m.to !== 'alice@example.com');
assert.ok(creatorMail && creatorMail.opts.htmlBody.includes(`${ymd(10)} ~ ${ymd(11)} (1박 2일)`), '크리에이터 확정 메일은 한국어 박수 표기');
assert.ok(!/泊/.test(creatorMail.opts.htmlBody), '크리에이터 확정 메일에 일본어 박수 표기 없음');
assert.ok(adminMail && adminMail.opts.htmlBody.includes('(1泊2日)'), '운영진 확정 메일은 일본어 박수 표기');

let sun = 3; while (new Date(ymd(sun) + 'T00:00:00').getDay() !== 0) sun++;
staySetup('金,土', []);
sheets.Restaurant_List.cells[2][33] = 7;
assert.strictEqual(stay(ymd(sun), ymd(sun + 5)).isAvailable, true, '일요일 체크인 ~ 금요일 체크아웃 가능');
assert.strictEqual(stay(ymd(sun), ymd(sun + 6)).isAvailable, false, '토요일 체크아웃은 금요일 밤이 휴무라 불가');
assert.strictEqual(stay(ymd(sun - 2), ymd(sun)).isAvailable, false, '금요일 체크인은 불가');
assert.strictEqual(stay(ymd(sun - 1), ymd(sun + 1)).isAvailable, false, '토요일 밤 포함 불가');

// ---- 7. 아임웹 세트(1+1) 주문 → 매장별 행 확장 ----------------------------
function imwebRest(id, ko, jp, pair) { const r = new Array(35).fill(''); r[0] = id; r[1] = ko; r[2] = jp; r[35] = pair; return r; }
const MAGU = '마구로토 고항 쿠로다한 후쿠오카 본점', TORI = '토리카와 중화소바 카와노소바';
function imwebSetup() {
  sheets = {
    Master_Log: fakeSheet([[], []]),
    Restaurant_List: fakeSheet([[], [], imwebRest('BOOKMARK-0001', MAGU, 'まぐろ', 'BOOKMARK-0003'), imwebRest('BOOKMARK-0003', TORI, '鶏川', ''), imwebRest('BOOKMARK-0004', '단품 식당', '', '')]),
    User_DB: fakeSheet([[]]), Settings: null, Log: null,
  };
}
function imwebPayload(type, prodName, optionInfo, total) {
  return { postData: { contents: JSON.stringify({ eventType: type, data: { orderNo: 202610303865991, memberUid: 'm@example.com', ordererName: '황상훈', ordererCall: '01012345678', totalPaymentPrice: total,
    sections: [{ sectionItems: [{ productInfo: { prodName, optionInfo } }] }] } }) } };
}
const doubleName = '[더블 체험] ' + MAGU + ' + ' + TORI;
const doubleOpt = { '[필수] 추가 매장 선택': TORI };
const logRows = () => sheets.Master_Log.cells.slice(2);

imwebSetup();
run(`handleImwebWebhook(${JSON.stringify(imwebPayload('ORDER_CREATE', doubleName, doubleOpt, 20000))})`);
let rows = logRows();
assert.strictEqual(rows.length, 2, '세트 주문은 2행');
assert.deepStrictEqual(rows.map(r => r[4]), ['BOOKMARK-0001', 'BOOKMARK-0003'], 'A 다음 B');
assert.deepStrictEqual(rows.map(r => r[10]), [10000, 10000], '보증금은 행 수로 나눔');
assert.deepStrictEqual(rows.map(r => r[6]), ['', ''], '매장 선택 옵션은 구매옵션에서 제외');
assert.deepStrictEqual(rows.map(r => r[11]), ['예약대기', '예약대기']);

run(`handleImwebWebhook(${JSON.stringify(imwebPayload('ORDER_PAY_COMPLETE', doubleName, doubleOpt, 20000))})`);
assert.strictEqual(logRows().length, 2, '같은 주문 재전송은 행이 늘지 않음');

run(`handleImwebWebhook(${JSON.stringify(imwebPayload('ORDER_CANCEL', doubleName, doubleOpt, 20000))})`);
assert.deepStrictEqual(logRows().map(r => r[11]), ['취소(아임웹)', '취소(아임웹)'], '취소는 두 행 모두');

imwebSetup(); // A 단품명이어도 짝 열로 B 생성, 런치/디너 옵션은 보존
run(`handleImwebWebhook(${JSON.stringify(imwebPayload('ORDER_CREATE', MAGU, { 코스: '런치' }, 20000))})`);
rows = logRows();
assert.deepStrictEqual(rows.map(r => r[4]), ['BOOKMARK-0001', 'BOOKMARK-0003'], '짝 열로 B 보충');
assert.deepStrictEqual(rows.map(r => r[6]), ['런치', '런치']);

imwebSetup(); // 짝 없는 매장은 1행 / B 상품은 A 를 끌고 오지 않음 / 매칭 실패는 상품명 그대로
run(`handleImwebWebhook(${JSON.stringify(imwebPayload('ORDER_CREATE', '단품 식당', {}, 10000))})`);
run(`handleImwebWebhook(${JSON.stringify(imwebPayload('ORDER_CREATE', TORI, {}, 10000))})`);
assert.deepStrictEqual(logRows().map(r => r[4]), ['BOOKMARK-0004', 'BOOKMARK-0003']);
const lst = [{ id: 'S', nameKo: '가나다 본점', nameJp: '', pairId: '' }, { id: 'T', nameKo: '가나다', nameJp: '', pairId: '' }];
assert.deepStrictEqual(Array.from(run(`resolveImwebStores_('가나다 본점', ${JSON.stringify(lst)})`), s => s.id), ['S'], '긴 이름 안의 짧은 이름은 중복 매칭하지 않음');
assert.deepStrictEqual(Array.from(run(`resolveImwebStores_('없는 매장', ${JSON.stringify(lst)})`)), [], '매칭 실패는 빈 배열');

console.log('OK: logic-smoke (auth tokens, store access, confirmBookingByLink, daily reminders, stay booking, slot regression, stay mail language, imweb bundle)');
