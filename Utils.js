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

/** [아임웹 유저 동기화 - 신규 등록 + 기존 회원 정보 업데이트(업서트)] */
function syncImwebUsers() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const rawSheet = ss.getSheetByName("Imweb_Raw");
  const userSheet = ss.getSheetByName("User_DB");
  
  if (!rawSheet || !userSheet) return SpreadsheetApp.getUi().alert("Imweb_Raw 또는 User_DB 시트가 없습니다.");

  const rawData = rawSheet.getDataRange().getValues();
  if (rawData.length <= 1) return SpreadsheetApp.getUi().alert("가져올 데이터가 없습니다.");

  // 🎯 키워드로 유연하게 엑셀 헤더 열 번호 찾기
  const headers = rawData[0].map(h => String(h || "").trim().toLowerCase());
  const col = (...keywords) => {
    return headers.findIndex(h => keywords.some(k => h.includes(k.toLowerCase())));
  };

  // 기존 User_DB의 전체 데이터 및 행 위치 매핑 (고유키 -> 행 번호)
  const userData = userSheet.getDataRange().getValues();
  const existingUserRowMap = new Map();
  for (let r = 1; r < userData.length; r++) {
    const key = String(userData[r][0] || "").trim();
    if (key) {
      existingUserRowMap.set(key, r + 1); // 1-based 시트 행 번호
    }
  }

  let addedCount = 0;
  let updatedCount = 0;

  for (let i = 1; i < rawData.length; i++) {
    const row = rawData[i];
    
    // 고유키 검증
    const keyIdx = col("고유키", "회원코드", "member_code");
    const uniqueKey = keyIdx !== -1 ? String(row[keyIdx] || "").trim() : "";
    if (!uniqueKey) continue; 

    // 연락처 포맷팅
    const phoneIdx = col("연락처", "휴대폰", "전화번호");
    let rawPhone = phoneIdx !== -1 ? String(row[phoneIdx] || "").replace(/[^0-9]/g, "") : "";
    if (rawPhone.startsWith("10") && (rawPhone.length === 9 || rawPhone.length === 10)) {
      rawPhone = "0" + rawPhone;
    }
    
    let formattedPhone = rawPhone;
    if (rawPhone.length === 11) {
      formattedPhone = rawPhone.replace(/(\d{3})(\d{4})(\d{4})/, "$1-$2-$3");
    } else if (rawPhone.length === 10) {
      formattedPhone = rawPhone.replace(/(\d{3})(\d{3})(\d{4})/, "$1-$2-$3");
    }

    const getVal = (...keys) => {
      const idx = col(...keys);
      return idx !== -1 ? (row[idx] || "") : "";
    };

    const englishName = getVal("영문명", "영문이름", "english");
    const koreanName = getVal("이름", "성명", "실명");
    const blogUrl = getVal("블로그", "blog", "네이버");
    const instaUrl = getVal("인스타", "insta", "instagram", "릴스");
    const youtubeUrl = getVal("유튜브", "youtube");
    const tiktokUrl = getVal("틱톡", "tiktok", "tictok");
    const googleGuide = getVal("구글", "로컬", "가이드");
    const userGroup = getVal("회원 그룹", "그룹", "등급");
    const signupDate = getVal("가입일", "가입 승인일");
    const email = getVal("이메일", "email");
    const adminMemo = getVal("관리자 메모", "메모");

    // 🔄 [케이스 1: 기존 유저가 있는 경우 -> 누락/최신 정보 업데이트]
    if (existingUserRowMap.has(uniqueKey)) {
      const targetRow = existingUserRowMap.get(uniqueKey);
      
      // B열~I열: 기본 인적사항 및 SNS 채널 URL 업데이트
      if (englishName) userSheet.getRange(targetRow, 2).setValue(englishName);       // B열: 영문명
      if (koreanName) userSheet.getRange(targetRow, 3).setValue(koreanName);         // C열: 이름
      if (formattedPhone) userSheet.getRange(targetRow, 4).setValue(formattedPhone); // D열: 연락처
      if (blogUrl) userSheet.getRange(targetRow, 5).setValue(blogUrl);               // E열: 블로그
      if (instaUrl) userSheet.getRange(targetRow, 6).setValue(instaUrl);             // F열: 인스타
      if (youtubeUrl) userSheet.getRange(targetRow, 7).setValue(youtubeUrl);         // G열: 유튜브
      if (tiktokUrl) userSheet.getRange(targetRow, 8).setValue(tiktokUrl);           // H열: 틱톡
      if (googleGuide) userSheet.getRange(targetRow, 9).setValue(googleGuide);       // I열: 구글가이드
      
      // J열(등급), K열(패널티), L열(완료수), O열(계좌)은 시트 수동 관리 데이터를 위해 보존하고
      // N열(이메일), P열(메모)만 최신화
      if (email) userSheet.getRange(targetRow, 14).setValue(email);                  // N열: 이메일
      if (adminMemo) userSheet.getRange(targetRow, 16).setValue(adminMemo);          // P열: 관리자메모
      
      updatedCount++;
    } 
    // ➕ [케이스 2: 완전 신규 유저인 경우 -> 새로운 행 추가]
    else {
      const newRow = new Array(16).fill("");
      newRow[0] = uniqueKey;        // A열: 멤버코드
      newRow[1] = englishName;      // B열: 영문명
      newRow[2] = koreanName;       // C열: 한국어 실명
      newRow[3] = formattedPhone;   // D열: 연락처
      newRow[4] = blogUrl;          // E열: 블로그
      newRow[5] = instaUrl;         // F열: 인스타
      newRow[6] = youtubeUrl;       // G열: 유튜브
      newRow[7] = tiktokUrl;        // H열: 틱톡
      newRow[8] = googleGuide;      // I열: 구글 로컬 가이드
      newRow[9] = userGroup;        // J열: 회원 그룹
      newRow[10] = "";              // K열: 누적 패널티
      newRow[11] = "";              // L열: 미션 완료수
      newRow[12] = signupDate;      // M열: 가입일
      newRow[13] = email;           // N열: 이메일
      newRow[14] = "";              // O열: 환불 계좌
      newRow[15] = adminMemo;       // P열: 관리자 메모

      userSheet.appendRow(newRow);
      existingUserRowMap.set(uniqueKey, userSheet.getLastRow());
      addedCount++;
    }
  }

  SpreadsheetApp.getUi().alert(`✅ 동기화 완료!\n\n• 신규 등록: ${addedCount}명\n• 정보 업데이트: ${updatedCount}명`);
}

/**
 * 🚨 [자동 노쇼 일괄 처리 엔진 - 최종 비즈니스 세이프가드 적용]
 * - 대상: 상태(L열)가 '방문전'이고, 방문 예약 일시(H열)로부터 익일 04:00가 경과한 미방문 건
 * - 보호 대상: '예약대기', '일정조율필요', '예약확인중', '방문완료', '제출완료', '취소완료' 등은 절대 노쇼 처리하지 않음
 * - 시스템 로그는 Y열(25번째 열)에 기록
 */
function checkAndMarkNoShow() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Master_Log');
  if (!sheet) {
    Logger.log("❌ Master_Log 시트를 찾을 수 없습니다.");
    return;
  }

  try {
    const data = sheet.getDataRange().getValues();
    if (data.length < 3) return; // 헤더 제외 3행부터 시작

    const now = new Date();
    const timeZone = Session.getScriptTimeZone() || "Asia/Seoul";
    let updatedCount = 0;

    // Master_Log 데이터 시작 행: 3행 (배열 인덱스 2부터)
    for (let i = 2; i < data.length; i++) {
      const row = data[i];
      const rowNum = i + 1; // 실제 스프레드시트 행 번호

      const orderNo = String(row[0] || '').trim().replace(/'/g, ''); // A열: 주문번호
      const status = String(row[11] || '').trim();                  // L열: 진행 상태 (인덱스 11)
      const visitDateRaw = row[7];                                  // H열: 방문예정일시 (인덱스 7)

      // 1️⃣ 주문번호가 없거나, '방문전'이 아닌 다른 모든 상태는 건너뜀 (안전 격리)
      if (!orderNo || status !== '방문전') {
        continue;
      }

      // 2️⃣ 날짜 파싱 (Date 객체 및 YYYY-MM-DD HH:mm 문자열 안전 파싱)
      let visitDate = null;
      if (visitDateRaw instanceof Date) {
        visitDate = new Date(visitDateRaw.getTime());
      } else if (typeof visitDateRaw === 'string' && visitDateRaw.trim() !== '') {
        const match = visitDateRaw.trim().match(/^(\d{4})[-/.](\d{1,2})[-/.](\d{1,2})(?:\s+(\d{1,2}):(\d{1,2}))?/);
        if (match) {
          const year = parseInt(match[1], 10);
          const month = parseInt(match[2], 10) - 1;
          const day = parseInt(match[3], 10);
          const hour = match[4] ? parseInt(match[4], 10) : 23;
          const min = match[5] ? parseInt(match[5], 10) : 59;
          visitDate = new Date(year, month, day, hour, min, 0);
        }
      }

      if (!visitDate || isNaN(visitDate.getTime())) {
        continue;
      }

      // 3️⃣ [비즈니스 버퍼] 방문일 익일 새벽 04:00 이후에만 노쇼 처리
      const safeDeadline = new Date(visitDate.getTime());
      safeDeadline.setDate(safeDeadline.getDate() + 1);
      safeDeadline.setHours(4, 0, 0, 0);

      if (now > safeDeadline) {
        // L열 (12번째 열) 진행 상태를 '노쇼'로 변경
        sheet.getRange(rowNum, 12).setValue('노쇼');
        
        // 🎯 M열은 건드리지 않고, 시스템 로그 전용인 Y열(25번째 열)에만 기록
        const currentSysLog = String(sheet.getRange(rowNum, 25).getValue() || '').trim();
        const autoLog = `[시스템] ${Utilities.formatDate(now, timeZone, 'yyyy-MM-dd HH:mm')} 자동 노쇼 처리`;
        const updatedSysLog = currentSysLog ? `${currentSysLog} | ${autoLog}` : autoLog;
        sheet.getRange(rowNum, 25).setValue(updatedSysLog);

        updatedCount++;
        Logger.log(`[노쇼 처리] Row: ${rowNum} | 주문번호: ${orderNo} | 방문예정일: ${visitDateRaw}`);
      }
    }

    Logger.log(`✅ [checkAndMarkNoShow] 총 ${updatedCount}건 노쇼 처리 완료`);
    
    // 수동 메뉴 클릭 시에만 UI 알림창 출력 (새벽 자동 트리거 시 팝업 에러 방어)
    if (SpreadsheetApp.getUi) {
      try {
        SpreadsheetApp.getUi().alert(`✅ 노쇼 처리 완료\n\n총 ${updatedCount}건이 노쇼 처리되었습니다.`);
      } catch(uiErr) {}
    }

  } catch (err) {
    Logger.log(`❌ [checkAndMarkNoShow 오류] ${err.toString()}`);
  }
}
