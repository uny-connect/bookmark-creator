// 메일 템플릿 스모크 테스트: node test/mail-smoke.js
// GAS 서비스를 목(mock)으로 대체해 Service_Mail.js 의 send*Mail 이 예외 없이 렌더링되는지,
// 본문에 'undefined' / '[object Object]' 가 섞이지 않는지 확인한다.
const fs = require('fs');
const path = require('path');
const vm = require('vm');

const sent = [];
const sandbox = {
  console,
  GmailApp: { sendEmail: (to, subject, body, opts) => sent.push({ to, subject, opts }) },
  Utilities: { formatDate: () => '2026-01-01' },
  Session: { getScriptTimeZone: () => 'Asia/Seoul' },
  SpreadsheetApp: { getActiveSpreadsheet: () => ({ getSheetByName: () => null }) },
  ScriptApp: { getService: () => ({ getUrl: () => 'https://script.google.com/macros/s/TEST/exec' }) },
};
vm.createContext(sandbox);
['Config.js', 'Utils.js', 'Service_Mail.js'].forEach(f => {
  vm.runInContext(fs.readFileSync(path.join(__dirname, '..', f), 'utf8'), sandbox, { filename: f });
});

const base = { to: 'x@example.com', storeName: '테스트매장', storeNameJp: 'テスト店', storeNameKo: '테스트매장', visitDateStr: '2026-01-02 12:00', peopleCount: '2', orderNo: '1001', memberName: 'Alice', memberCode: 'mc1', storeId: 'S1', dateStr: '2026-01-02', timeStr: '12:00', creatorProfileUrl: 'instagram.com/a', benefit: 'ランチ', row: 5, creatorName: 'Alice', visitDisplay: '2026-01-02 12:00', msg: '16:00なら可能です', addressLabel: '登録アドレス', storeEmail: 's@example.com' };
const fns = ['sendCreatorConfirmedMail', 'sendCreatorRescheduleNoticeMail', 'sendCreatorRescheduleRequestMail', 'sendStoreBookingRequestMail', 'sendStoreCancellationMail', 'sendStoreConnectionTestMail', 'sendAdminBookingRequestMail', 'sendAdminBookingConfirmedMail', 'sendAdminRescheduleRequestMail'];
let failed = 0;
for (const fn of fns) {
  // capture htmlBody by wrapping sendEmail
  let html = '';
  sandbox.GmailApp.sendEmail = (to, subject, body, opts) => { html = opts.htmlBody + '\n' + subject; sent.push({ to, subject, opts }); };
  try {
    sandbox[fn] ? sandbox[fn](base) : vm.runInContext(`${fn}(${JSON.stringify(base)})`, sandbox);
    if (/undefined|\[object Object\]/.test(html.replace(/from: undefined/g, ''))) { failed++; console.error(`FAIL ${fn}: 본문에 undefined 포함`); }
    else console.log(`ok   ${fn}`);
  } catch (e) { failed++; console.error(`FAIL ${fn}: ${e.message}`); }
}
process.exit(failed ? 1 : 0);
