/** 순수 헬퍼: 날짜/문자 정규화, 시트 일괄 조회, Settings 조회 */

/** 🎯 현재 활성화된 웹 앱 URL 안전 취득 (Settings 시트 우선 참조 및 ScriptApp 폴백) */
function getActiveWebAppUrl() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(SHEET.SETTINGS);
    if (sheet) {
      const data = sheet.getDataRange().getValues();
      for (let i = 1; i < data.length; i++) {
        const key = String(data[i][IDX_SETTINGS.KEY] || '').trim().toUpperCase();
        if (key === "WEB_APP_URL") {
          const customUrl = String(data[i][IDX_SETTINGS.VALUE] || '').trim();
          if (customUrl && customUrl.startsWith("http")) return customUrl;
        }
      }
    }
  } catch(e) {}
  
  return ScriptApp.getService().getUrl();
}

/** [2] 등급 이모지 변환 (객체 매핑 구조로 압축) */
function parseTierEmoji(tierStr) {
  if (!tierStr) return "🟡";
  const normalized = String(tierStr).trim().toUpperCase();
  const emojiMap = { "YELLOW": "🟡", "RED": "🔴", "BLACK": "⚫" };
  const key = Object.keys(emojiMap).find(k => normalized.includes(k));
  return emojiMap[key] || "🟡";
}

/** [휴무일 안전 변환] (인라인 조건문 최적화) */
function getSafeBlackouts(raw) {
  if (!raw) return [];
  if (raw instanceof Date) return [Utilities.formatDate(raw, Session.getScriptTimeZone(), 'yyyy-MM-dd')];
  return String(raw).split(',').map(s => s.trim()).filter(Boolean);
}

/** [추가] 전각 문자를 반각 문자로 변환하는 공통 세탁 함수 */
function toHalfWidth(str) {
  if (!str) return "";
  return String(str)
    .replace(/[！-～]/g, function(s) {
      return String.fromCharCode(s.charCodeAt(0) - 0xFEE0);
    })
    .replace(/：/g, ":")
    .replace(/[－─〜~ー━]/g, "-")
    .replace(/[，、]/g, ",")
    .replace(/[\s\u3000]+/g, "");
}

/** [공통 헬퍼] 시트 전체 데이터 일괄 수집 (null 방어 포함) */
function _getSheetsData(sheetNames) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  return sheetNames.reduce((acc, name) => {
    const sheet = ss.getSheetByName(name);
    acc[name] = sheet ? sheet.getDataRange().getValues() : [];
    return acc;
  }, {});
}

/** Settings 시트에서 운영진 알림 수신 이메일 목록 불러오기 */
function getAdminAlertEmails() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(SHEET.SETTINGS);
    if (!sheet) return MAIL.DEFAULT_ADMIN;

    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      const key = String(data[i][IDX_SETTINGS.KEY] || '').trim().toUpperCase();
      if (key === "ADMIN_NOTIFICATION_EMAILS") {
        const emails = String(data[i][IDX_SETTINGS.VALUE] || '').trim();
        return emails || MAIL.DEFAULT_ADMIN;
      }
    }
  } catch (e) {
    Logger.log("getAdminAlertEmails 에러: " + e.toString());
  }
  return MAIL.DEFAULT_ADMIN;
}

/** 메일/HTML 본문에 사용자 입력을 넣을 때 사용하는 이스케이프 */
function _escapeHtml(v) {
  return String(v === undefined || v === null ? '' : v)
    .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;').replace(/'/g, '&#39;');
}

/** H열 방문일시(Date 또는 문자열) → { dateStr:'yyyy-MM-dd', timeStr:'HH:mm' }. "2026-10-27 (화) 12:00" 처럼 요일이 섞인 문자열도 정규식으로 분리 */
function _splitVisit_(v, timeZone) {
  if (v instanceof Date) return { dateStr: Utilities.formatDate(v, timeZone, 'yyyy-MM-dd'), timeStr: Utilities.formatDate(v, timeZone, 'HH:mm') };
  const s = String(v || '');
  const d = s.match(/\d{4}-\d{2}-\d{2}/), t = s.match(/(\d{1,2}):(\d{2})/);
  return { dateStr: d ? d[0] : s.substring(0, 10), timeStr: t ? ('0' + t[1]).slice(-2) + ':' + t[2] : '' };
}

/** 메일 링크용 일정 토큰: 방문일시(+STAY 체크아웃일)를 숫자열로 정규화 ("2026-11-02 12:00" → "202611021200"). 일정이 바뀌면 값이 달라져 옛 링크를 가려낸다. 일정이 비면 '' */
function _visitToken_(rawVisit, timeZone, checkoutYmd) {
  const sp = _splitVisit_(rawVisit, timeZone);
  if (!sp.dateStr) return '';
  return (sp.dateStr + sp.timeStr).replace(/\D/g, '') + (checkoutYmd ? '-' + String(checkoutYmd).replace(/\D/g, '') : '');
}
