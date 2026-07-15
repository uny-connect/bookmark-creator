/** [2] 등급 이모지 변환 (객체 매핑 구조로 압축) */
function parseTierEmoji(tierStr) {
  if (!tierStr) return "🟡";
  const normalized = String(tierStr).trim().toUpperCase();
  const emojiMap = { "YELLOW": "🟡", "RED": "🔴", "BLACK": "⚫" };
  
  // 등급명이 포함되어 있는지 검사 후 매칭, 없으면 기본값 '🟡'
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
