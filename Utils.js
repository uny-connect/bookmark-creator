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

function syncImwebUsers() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const rawSheet = ss.getSheetByName("Imweb_Raw");
  const userSheet = ss.getSheetByName("User_DB");
  
  if (!rawSheet || !userSheet) return SpreadsheetApp.getUi().alert("Imweb_Raw 또는 User_DB 시트가 없습니다.");

  const rawData = rawSheet.getDataRange().getValues();
  if (rawData.length <= 1) return SpreadsheetApp.getUi().alert("가져올 데이터가 없습니다.");

  // 헤더 텍스트로 엑셀의 열 인덱스 찾기
  const headers = rawData[0];
  const col = (name) => headers.indexOf(name);

  // 기존 User_DB에 있는 고유키(A열) 목록 (중복 가입 방지용)
  const existingUsers = userSheet.getDataRange().getValues().map(row => String(row[0]).trim());
  let addedCount = 0;

  for (let i = 1; i < rawData.length; i++) {
    const row = rawData[i];
    
    // 고유키 검증
    const uniqueKey = String(row[col("고유키")] || "").trim();
    if (!uniqueKey || existingUsers.includes(uniqueKey)) continue; 

    // 📱 [연락처 포맷팅 로직] 엑셀에서 0이 날아간 번호 복구 및 하이픈 자동 삽입
    let rawPhone = String(row[col("연락처")] || "").replace(/[^0-9]/g, ""); // 숫자만 추출
    if (rawPhone.startsWith("10") && (rawPhone.length === 9 || rawPhone.length === 10)) {
      rawPhone = "0" + rawPhone; // 앞에 0이 날아간 경우 복구
    }
    
    let formattedPhone = rawPhone;
    if (rawPhone.length === 11) {
      formattedPhone = rawPhone.replace(/(\d{3})(\d{4})(\d{4})/, "$1-$2-$3"); // 010-XXXX-XXXX
    } else if (rawPhone.length === 10) {
      formattedPhone = rawPhone.replace(/(\d{3})(\d{3})(\d{4})/, "$1-$2-$3"); // 02-XXX-XXXX 등
    }

    // 🎯 [시트 구조 변경 반영] 총 16열(A~P)에 맞춘 데이터 배열 생성
    const newRow = new Array(16).fill(""); 
    
    newRow[0] = uniqueKey;                                // A열: 멤버코드(고유키)
    newRow[1] = row[col("영문명")] || "";                   // B열: 영문 성함
    newRow[2] = row[col("이름")] || "";                     // C열: 한국어 실명
    newRow[3] = formattedPhone;                           // D열: 연락처 (하이픈 처리 완)
    newRow[4] = row[col("네이버 블로그")] || "";             // E열: 블로그 URL
    newRow[5] = row[col("인스타그램(릴스)")] || "";          // F열: 인스타그램 URL
    newRow[6] = row[col("유튜브(youtube)")] || "";          // G열: 유튜브 URL
    newRow[7] = row[col("틱톡(TicTok)")] || "";             // H열: 틱톡 URL
    newRow[8] = row[col("구글 로컬 가이드(프로필캡쳐)")] || ""; // I열: 구글 로컬 가이드 (신규행)
    newRow[9] = row[col("회원 그룹")] || "";                // J열: 현재 등급
    newRow[10] = "";                                      // K열: 누적 패널티 (요청하신 대로 빈칸 처리)
    newRow[11] = "";                                      // L열: 미션 완료수 (요청하신 대로 빈칸 처리)
    newRow[12] = row[col("가입일")] || "";                  // M열: 가입 승인일
    newRow[13] = row[col("이메일")] || "";                  // N열: 이메일 주소
    newRow[14] = "";                                      // O열: 환불용 계좌 정보
    newRow[15] = row[col("관리자 메모")] || "";             // P열: 관리자 메모
    
    userSheet.appendRow(newRow);
    addedCount++;
  }
  
  SpreadsheetApp.getUi().alert(`✅ 동기화 완료! 총 ${addedCount}명의 신규 유저가 User_DB에 등록되었습니다.`);
}
