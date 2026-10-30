// 월간 리포트 집계 스모크 테스트: node test/report-smoke.js
// buildReportRows_ (순수 함수) 의 월 필터 / 상태 분기 / 매칭 실패 / 멱등성을 점검한다.
const fs = require('fs');
const path = require('path');
const vm = require('vm');

const sandbox = {
  console,
  Logger: { log: () => {} },
  Utilities: { formatDate: (d, tz, f) => { const s = d.toISOString(); return f === 'yyyy-MM' ? s.substring(0, 7) : f === 'HH:mm' ? s.substring(11, 16) : s.substring(0, 10); } },
  Session: { getScriptTimeZone: () => 'UTC' },
};
vm.createContext(sandbox);
['Config.js', 'Utils.js', 'Service_Report.js'].forEach(f => {
  vm.runInContext(fs.readFileSync(path.join(__dirname, '..', f), 'utf8'), sandbox, { filename: f });
});

vm.runInContext('this.IDX_MASTER = IDX_MASTER; this.IDX_REST = IDX_REST;', sandbox);
let failed = 0;
const eq = (label, a, b) => {
  if (JSON.stringify(a) !== JSON.stringify(b)) { failed++; console.error(`FAIL ${label}\n  got:  ${JSON.stringify(a)}\n  want: ${JSON.stringify(b)}`); }
  else console.log(`ok   ${label}`);
};

const mrow = (o) => {
  const r = new Array(25).fill('');
  r[sandbox.IDX_MASTER.STORE_ID] = o.id; r[sandbox.IDX_MASTER.VISIT] = o.visit; r[sandbox.IDX_MASTER.STATUS] = o.status;
  r[sandbox.IDX_MASTER.SUBMIT_DATE] = o.sub || ''; r[sandbox.IDX_MASTER.REVIEW_URL] = o.q || ''; r[sandbox.IDX_MASTER.GOOGLE_URL] = o.r || ''; r[sandbox.IDX_MASTER.SHORT_REVIEW_JP] = o.t || '';
  return r;
};
const rrow = (id, ko, jp) => { const r = new Array(35).fill(''); r[sandbox.IDX_REST.ID] = id; r[sandbox.IDX_REST.NAME_KO] = ko; r[sandbox.IDX_REST.NAME_JP] = jp; return r; };

const master = [[], [],
  mrow({ id: 'a1', visit: '2026-09-30 (수) 18:00', status: '방문완료' }),                                  // 말일 방문, 미제출
  mrow({ id: 'A1', visit: '2026-09-05 (토) 12:00', status: '제출완료', sub: '2026-09-12', q: 'http://q', r: 'http://r', t: '美味しい' }),
  mrow({ id: 'A1', visit: '2026-09-03 (목) 12:00', status: '제출완료', sub: '2026-09-08', q: 'http://q2' }),
  mrow({ id: 'A1', visit: '2026-09-10 (목) 12:00', status: '취소(운영자)' }),                              // 제외
  mrow({ id: 'A1', visit: '2026-09-11 (금) 12:00', status: '노쇼' }),                                     // 제외
  mrow({ id: 'A1', visit: '2026-10-01 (목) 12:00', status: '제출완료', sub: '2026-10-05' }),              // 다른 달
  mrow({ id: 'ZZ', visit: '2026-09-15 (화) 12:00', status: '제출완료', sub: '2026-09-20' }),              // 매칭 실패
  mrow({ id: '', visit: '2026-09-16 (수) 12:00', status: '방문완료' }),                                    // 빈 점포ID
  mrow({ id: 'B2', visit: '', status: '제출완료' }),                                                       // 방문일 없음
];
const rest = [[], [], rrow('A1', '가게', 'カフェ'), rrow('B2', '두번째', '')];

const rows = sandbox.buildReportRows_(master, rest, '2026-09', 'UTC');
eq('행 수', rows.length, 5);
eq('A1 제출완료 2건은 제출일 순', rows.filter(r => r[1] === 'A1').map(r => [r[3], r[4], r[8]]), [[2, '2026-09-08', '제출완료'], [2, '2026-09-12', '제출완료'], [2, '', '제출 대기']]);
eq('업체명 JP', rows.find(r => r[1] === 'A1')[2], 'カフェ');
eq('링크/후기 매핑', rows.find(r => r[4] === '2026-09-12').slice(5, 8), ['http://q', 'http://r', '美味しい']);
eq('매칭 실패 표시', rows.find(r => r[1] === 'ZZ')[2], '(店舗ID不明: ZZ)');
eq('빈 점포ID 표시', rows.find(r => r[1] === '')[2], '(店舗ID未入力)');
eq('방문일 매핑', rows.filter(r => r[1] === 'A1').map(r => r[9]), ['2026-09-03', '2026-09-05', '2026-09-30']);
eq('월 라벨 JP', sandbox._monthLabelJp_('2026-09'), '2026年9月');
eq('멱등(같은 입력 같은 결과)', sandbox.buildReportRows_(master, rest, '2026-09', 'UTC'), rows);
eq('다른 달 분리', sandbox.buildReportRows_(master, rest, '2026-10', 'UTC').length, 1);
eq('문서 URL → ID', sandbox._extractDriveId_('https://docs.google.com/document/d/1JUv4F7BczxTA9VF8uqssuAbcfHVnEWtpfT0uD6jBr5s/edit?tab=t.0'), '1JUv4F7BczxTA9VF8uqssuAbcfHVnEWtpfT0uD6jBr5s');
eq('폴더 URL → ID', sandbox._extractDriveId_('https://drive.google.com/drive/folders/abc_123-X?usp=sharing'), 'abc_123-X');
eq('ID 그대로', sandbox._extractDriveId_(' abc123 '), 'abc123');
eq('전월 계산', sandbox._prevMonth_('UTC', new Date('2026-01-15T00:00:00Z')), '2025-12');

process.exit(failed ? 1 : 0);
