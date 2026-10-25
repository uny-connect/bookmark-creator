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
  SpreadsheetApp: { getActiveSpreadsheet: () => ({ getSheetByName: n => sheets[n] || null }) },
  ScriptApp: { getService: () => ({ getUrl: () => 'https://example.test/exec' }) },
};
vm.createContext(sandbox);
['Config.js', 'Utils.js', 'Service_Auth.js', 'Service_Mail.js', 'Service_Store.js'].forEach(f =>
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

console.log('OK: logic-smoke (auth tokens, store access, confirmBookingByLink)');
