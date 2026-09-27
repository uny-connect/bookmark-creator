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

/** [아임웹 유저 동기화] */
function syncImwebUsers() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const rawSheet = ss.getSheetByName("Imweb_Raw");
  const userSheet = ss.getSheetByName("User_DB");
  
  if (!rawSheet || !userSheet) return SpreadsheetApp.getUi().alert("Imweb_Raw 또는 User_DB 시트가 없습니다.");

  const rawData = rawSheet.getDataRange().getValues();
  if (rawData.length <= 1) return SpreadsheetApp.getUi().alert("가져올 데이터가 없습니다.");

  // 🎯 [핵심 수정 1] 헤더명에 띄어쓰기나 대소문자 차이가 있어도 키워드로 열을 찾는 함수
  const headers = rawData[0].map(h => String(h || "").trim().toLowerCase());
  const col = (...keywords) => {
    return headers.findIndex(h => keywords.some(k => h.includes(k.toLowerCase())));
  };

  // 기존 User_DB에 있는 고유키(A열) 목록 (중복 가입 방지용)
  const existingUsers = userSheet.getDataRange().getValues().map(row => String(row[0]).trim());
  let addedCount = 0;

  for (let i = 1; i < rawData.length; i++) {
    const row = rawData[i];
    
    // 고유키 검증
    const keyIdx = col("고유키", "회원코드", "member_code");
    const uniqueKey = keyIdx !== -1 ? String(row[keyIdx] || "").trim() : "";
    if (!uniqueKey || existingUsers.includes(uniqueKey)) continue; 

    // 연락처 포맷팅 로직
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

    // 🎯 [핵심 수정 2] 헤더 이름 오타/띄어쓰기 상관없이 안전하게 값을 가져오도록 매핑
    const getVal = (...keys) => {
      const idx = col(...keys);
      return idx !== -1 ? (row[idx] || "") : "";
    };

    // 총 16열(A~P)에 맞춘 데이터 배열 생성
    const newRow = new Array(16).fill(""); 
    
    newRow[0] = uniqueKey;                                             // A열: 멤버코드(고유키)
    newRow[1] = getVal("영문명", "영문이름", "english");               // B열: 영문 성함
    newRow[2] = getVal("이름", "성명", "실명");                         // C열: 한국어 실명
    newRow[3] = formattedPhone;                                        // D열: 연락처
    newRow[4] = getVal("블로그", "blog", "네이버");                     // E열: 블로그 URL
    newRow[5] = getVal("인스타", "insta", "instagram", "릴스");         // F열: 인스타그램 URL
    newRow[6] = getVal("유튜브", "youtube");                           // G열: 유튜브 URL
    newRow[7] = getVal("틱톡", "tiktok", "tictok");                    // H열: 틱톡 URL
    newRow[8] = getVal("구글", "로컬", "가이드");                       // I열: 구글 로컬 가이드
    newRow[9] = getVal("회원 그룹", "그룹", "등급");                    // J열: 현재 등급
    newRow[10] = "";                                                   // K열: 누적 패널티
    newRow[11] = "";                                                   // L열: 미션 완료수
    newRow[12] = getVal("가입일", "가입 승인일");                       // M열: 가입 승인일
    newRow[13] = getVal("이메일", "email");                            // N열: 이메일 주소
    newRow[14] = "";                                                   // O열: 환불용 계좌 정보
    newRow[15] = getVal("관리자 메모", "메모");                         // P열: 관리자 메모
    
    userSheet.appendRow(newRow);
    addedCount++;
  }

  SpreadsheetApp.getUi().alert(`✅ 동기화 완료! 총 ${addedCount}명의 신규 유저가 User_DB에 등록되었습니다.`);
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
