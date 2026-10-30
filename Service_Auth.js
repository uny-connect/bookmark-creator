/*****************************************************************************************
 * 웹앱 세션 토큰 / 로그인 시도 제한
 * - 웹앱은 ANYONE_ANONYMOUS 로 배포되어 google.script.run 으로 공개 함수를 누구나 호출할 수 있다.
 *   → 로그인(관리자 비밀번호 / 점주 PIN) 성공 시 서버가 토큰을 발급하고, 보호 대상 API는 토큰을 검증한다.
 * - 끝이 `_` 인 함수는 google.script.run 에서 호출할 수 없다(비공개). 인증 내부 함수는 모두 비공개로 둔다.
 ****************************************************************************************/

const AUTH_TTL_MS = 7 * 24 * 60 * 60 * 1000; // 클라이언트 세션 유지 기간과 동일(7일)
const AUTH_FAIL_WINDOW_SEC = 600;            // 로그인 실패 집계 구간(10분)
const AUTH_MAX_FAILS = { admin: 10, store: 30 }; // 구간 내 허용 실패 횟수 (범위별 전역 카운터)
const AUTH_TOKEN_PREFIX = 'TOKEN_';

/** 로그인 성공 시 호출: 범위(scope)와 주체(subject)에 묶인 토큰 발급. 만료된 토큰은 함께 정리한다. */
function issueToken_(scope, subject) {
  const props = PropertiesService.getScriptProperties();
  const now = Date.now();
  const all = props.getProperties();
  Object.keys(all).forEach(function (k) {
    if (k.indexOf(AUTH_TOKEN_PREFIX) !== 0) return;
    try { if (Number(JSON.parse(all[k]).exp) < now) props.deleteProperty(k); } catch (e) { props.deleteProperty(k); }
  });
  const token = Utilities.getUuid().replace(/-/g, '') + Utilities.getUuid().replace(/-/g, '');
  props.setProperty(AUTH_TOKEN_PREFIX + token, JSON.stringify({ scope: scope, subject: String(subject || ''), exp: now + AUTH_TTL_MS }));
  return token;
}

/** 토큰 검증. 실패 시 Error('AUTH_REQUIRED') — 클라이언트는 이 메시지로 재로그인을 유도한다. 성공 시 subject 반환. */
function requireToken_(token, scope) {
  const raw = token ? PropertiesService.getScriptProperties().getProperty(AUTH_TOKEN_PREFIX + String(token)) : null;
  if (raw) {
    try {
      const info = JSON.parse(raw);
      if (info.scope === scope && Number(info.exp) >= Date.now()) return info.subject;
    } catch (e) {}
  }
  throw new Error('AUTH_REQUIRED');
}

/** 로그아웃: 토큰 폐기 */
function revokeToken_(token) {
  if (token) PropertiesService.getScriptProperties().deleteProperty(AUTH_TOKEN_PREFIX + String(token));
}

/** 로그인 시도 제한: 실패가 한도를 넘으면 true */
function isLoginLocked_(scope) {
  return Number(CacheService.getScriptCache().get('authfail_' + scope) || 0) >= AUTH_MAX_FAILS[scope];
}
function recordLoginFailure_(scope) {
  const cache = CacheService.getScriptCache();
  const n = Number(cache.get('authfail_' + scope) || 0) + 1;
  cache.put('authfail_' + scope, String(n), AUTH_FAIL_WINDOW_SEC);
}
function clearLoginFailures_(scope) {
  CacheService.getScriptCache().remove('authfail_' + scope);
}
