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

const base = { to: 'x@example.com', storeName: '테스트매장', storeNameJp: 'テスト店', storeNameKo: '테스트매장', visitDateStr: '2026-01-02 12:00', peopleCount: '2', orderNo: '1001', memberName: 'Alice', memberCode: 'mc1', storeId: 'S1', dateStr: '2026-01-02', timeStr: '12:00', creatorProfileUrl: 'instagram.com/a', benefit: 'ランチ', row: 5, creatorName: 'Alice', visitDisplay: '2026-01-02 12:00', msg: '16:00なら可能です', addressLabel: '登録アドレス', storeEmail: 's@example.com', reservations: [{ timeStr: '12:00', peopleCount: '2', memberName: 'Alice', benefit: 'ランチ', creatorProfileUrl: 'instagram.com/a' }, { timeStr: '18:30', peopleCount: '4', memberName: 'Bob<b>', benefit: '' }] };
const fns = ['sendStoreDailyReminderMail','sendCreatorConfirmedMail', 'sendCreatorRescheduleNoticeMail', 'sendCreatorRescheduleRequestMail', 'sendStoreBookingRequestMail', 'sendStoreCancellationMail', 'sendStoreConnectionTestMail', 'sendAdminBookingRequestMail', 'sendAdminBookingConfirmedMail', 'sendAdminRescheduleRequestMail'];
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

// 제목만 봐도 일시/인원을 알 수 있어야 한다
const subjectOf = (fn, p) => { let s = ''; sandbox.GmailApp.sendEmail = (to, subject) => { s = subject; }; sandbox[fn](p); return s; };
const expectSubject = (fn, p, re) => { const s = subjectOf(fn, p); if (!re.test(s)) { failed++; console.error(`FAIL ${fn} 제목: ${s}`); } else console.log(`ok   ${fn} 제목: ${s}`); };
expectSubject('sendStoreBookingRequestMail', base, /2026-01-02 12:00 \/ 2名/);
expectSubject('sendAdminBookingRequestMail', base, /2026-01-02 12:00 \/ 2名/);
expectSubject('sendStoreDailyReminderMail', base, /2件: 12:00 2名 \/ 18:30 4名/);
expectSubject('sendStoreDailyReminderMail', Object.assign({}, base, { reservations: [base.reservations[0]] }), /\(12:00 \/ 2名\)/);
// 일정 토큰/변경 안내: 링크에 t 포함, isUpdate 일 때만 제목 접두·안내 배너
const bodyOf = (fn, p) => { let h = ''; sandbox.GmailApp.sendEmail = (to, subject, body, opts) => { h = opts.htmlBody; }; sandbox[fn](p); return h; };
const expectTrue = (name, cond) => { if (!cond) { failed++; console.error(`FAIL ${name}`); } else console.log(`ok   ${name}`); };
const updBase = Object.assign({}, base, { token: '202601021200', isUpdate: true });
expectSubject('sendStoreBookingRequestMail', updBase, /^【日時変更のお知らせ】ご来店予約のご確認 \(2026-01-02 12:00 \/ 2名\)$/);
expectTrue('갱신 메일: 확정/변경요청 링크에 t 포함', (bodyOf('sendStoreBookingRequestMail', updBase).match(/&t=202601021200/g) || []).length === 2);
expectTrue('갱신 메일: 안내 배너 포함', bodyOf('sendStoreBookingRequestMail', updBase).includes('先にお送りしたメールの内容が更新されました'));
const firstBase = Object.assign({}, base, { token: '202601021200' });
expectTrue('최초 메일: 제목에 변경 접두 없음', !/【日時変更/.test(subjectOf('sendStoreBookingRequestMail', firstBase)));
expectTrue('최초 메일: 배너 없음, t 포함', !bodyOf('sendStoreBookingRequestMail', firstBase).includes('更新されました') && bodyOf('sendStoreBookingRequestMail', firstBase).includes('&t=202601021200'));
expectTrue('토큰 없는 호출은 t 파라미터 생략', !bodyOf('sendStoreBookingRequestMail', base).includes('&t='));
process.exit(failed ? 1 : 0);
