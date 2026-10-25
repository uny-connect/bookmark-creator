// 정적 검증: node test/run-checks.js
// GAS 런타임을 쓰지 않고 구문 / 전역 이름 중복 / 화면↔서버 함수 연결 / 템플릿 파일 존재를 점검한다.
const fs = require('fs');
const path = require('path');
const vm = require('vm');

const root = path.join(__dirname, '..');
const read = f => fs.readFileSync(path.join(root, f), 'utf8');
const jsFiles = fs.readdirSync(root).filter(f => f.endsWith('.js'));
const htmlFiles = fs.readdirSync(root).filter(f => f.endsWith('.html'));
const errors = [];
const fail = msg => errors.push(msg);

// 1. 구문
jsFiles.forEach(f => {
  try { new vm.Script(read(f), { filename: f }); } catch (e) { fail(`[syntax] ${f}: ${e.message}`); }
});

// 2. 전역 이름 중복 (GAS는 모든 .js 가 하나의 전역 스코프)
const declared = new Map(); // name -> file
const serverFns = new Set();
jsFiles.forEach(f => {
  const src = read(f);
  for (const m of src.matchAll(/^(?:async\s+)?function\s+(\w+)\s*\(/gm)) {
    serverFns.add(m[1]);
    if (declared.has(m[1])) fail(`[duplicate] ${m[1]} in ${declared.get(m[1])} and ${f}`);
    else declared.set(m[1], f);
  }
  for (const m of src.matchAll(/^(?:const|let|var)\s+(\w+)\s*=/gm)) {
    if (declared.has(m[1])) fail(`[duplicate] ${m[1]} in ${declared.get(m[1])} and ${f}`);
    else declared.set(m[1], f);
  }
});

// 3. createTemplateFromFile / createHtmlOutputFromFile 대상 파일 존재
jsFiles.forEach(f => {
  for (const m of read(f).matchAll(/create(?:Template|HtmlOutput)FromFile\(\s*'([^']+)'/g)) {
    if (!htmlFiles.includes(m[1] + '.html')) fail(`[template] ${f}: '${m[1]}.html' 없음`);
  }
});

// 4. 화면의 google.script.run.<fn>() 이 서버 함수로 존재하는지
const HANDLERS = new Set(['withSuccessHandler', 'withFailureHandler', 'withUserObject']);
function runChainNames(src) {
  const names = [];
  let from = 0;
  for (;;) {
    const at = src.indexOf('google.script.run', from);
    if (at < 0) break;
    let i = at + 'google.script.run'.length;
    for (;;) {
      const m = /^\s*\.\s*(\w+)\s*\(/.exec(src.slice(i));
      if (!m) break;
      names.push(m[1]);
      i += m[0].length;
      let depth = 1, q = null;
      while (i < src.length && depth > 0) {
        const c = src[i];
        if (q) { if (c === '\\') i++; else if (c === q) q = null; }
        else if (c === '"' || c === "'" || c === '`') q = c;
        else if (c === '(') depth++;
        else if (c === ')') depth--;
        i++;
      }
    }
    from = Math.max(i, at + 1);
  }
  return names;
}
htmlFiles.forEach(f => {
  new Set(runChainNames(read(f))).forEach(n => {
    if (!HANDLERS.has(n) && !serverFns.has(n)) fail(`[run] ${f}: 서버에 없는 함수 호출 ${n}()`);
  });
});

// 5. 시트 메뉴 핸들러 존재
for (const m of read('Triggers.js').matchAll(/\.addItem\(\s*'[^']*'\s*,\s*'(\w+)'\s*\)/g)) {
  if (!serverFns.has(m[1])) fail(`[menu] onOpen 메뉴 핸들러 없음: ${m[1]}`);
}

if (errors.length) {
  console.error(errors.join('\n'));
  console.error(`\nFAILED: ${errors.length} issue(s)`);
  process.exit(1);
}
console.log(`OK: ${jsFiles.length} js, ${htmlFiles.length} html, ${serverFns.size} global functions`);
