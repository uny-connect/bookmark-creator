/** 🎯 현재 활성화된 웹 앱 URL 안전 취득 (Settings 시트 우선 참조 및 ScriptApp 폴백) */
function getActiveWebAppUrl() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName("Settings");
    if (sheet) {
      const data = sheet.getDataRange().getValues();
      for (let i = 1; i < data.length; i++) {
        const key = String(data[i][0] || '').trim().toUpperCase();
        if (key === "WEB_APP_URL") {
          const customUrl = String(data[i][1] || '').trim();
          if (customUrl && customUrl.startsWith("http")) return customUrl;
        }
      }
    }
  } catch(e) {}
  
  // Settings에 설정이 없으면 기본 ScriptApp URL 사용
  return ScriptApp.getService().getUrl();
}

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

/** [아임웹 유저 동기화 - User_DB 업서트 + Master_Log 과거 이메일 식별자 자동 승격] */
function syncImwebUsers() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const rawSheet = ss.getSheetByName("Imweb_Raw");
  const userSheet = ss.getSheetByName("User_DB");
  const logSheet = ss.getSheetByName("Master_Log");
  
  if (!rawSheet || !userSheet) return SpreadsheetApp.getUi().alert("Imweb_Raw 또는 User_DB 시트가 없습니다.");

  const rawData = rawSheet.getDataRange().getValues();
  if (rawData.length <= 1) return SpreadsheetApp.getUi().alert("가져올 데이터가 없습니다.");

  const headers = rawData[0].map(h => String(h || "").trim().toLowerCase());
  const col = (...keywords) => {
    return headers.findIndex(h => keywords.some(k => h.includes(k.toLowerCase())));
  };

  const userData = userSheet.getDataRange().getValues();
  const existingUserRowMap = new Map();
  for (let r = 1; r < userData.length; r++) {
    const key = String(userData[r][0] || "").trim();
    if (key) {
      existingUserRowMap.set(key, r + 1);
    }
  }

  let addedCount = 0;
  let updatedCount = 0;

  for (let i = 1; i < rawData.length; i++) {
    const row = rawData[i];
    
    const keyIdx = col("고유키", "회원코드", "member_code");
    const uniqueKey = keyIdx !== -1 ? String(row[keyIdx] || "").trim() : "";
    if (!uniqueKey) continue; 

    // 전화번호 하이픈 포맷팅
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

    // 케이스 1: 기존 유저 정보 업데이트
    if (existingUserRowMap.has(uniqueKey)) {
      const targetRow = existingUserRowMap.get(uniqueKey);
      if (englishName) userSheet.getRange(targetRow, 2).setValue(englishName);
      if (koreanName) userSheet.getRange(targetRow, 3).setValue(koreanName);
      if (formattedPhone) userSheet.getRange(targetRow, 4).setValue(formattedPhone);
      if (blogUrl) userSheet.getRange(targetRow, 5).setValue(blogUrl);
      if (instaUrl) userSheet.getRange(targetRow, 6).setValue(instaUrl);
      if (youtubeUrl) userSheet.getRange(targetRow, 7).setValue(youtubeUrl);
      if (tiktokUrl) userSheet.getRange(targetRow, 8).setValue(tiktokUrl);
      if (googleGuide) userSheet.getRange(targetRow, 9).setValue(googleGuide);
      if (userGroup) userSheet.getRange(targetRow, 10).setValue(userGroup);
      if (email) userSheet.getRange(targetRow, 14).setValue(email);
      if (adminMemo) userSheet.getRange(targetRow, 16).setValue(adminMemo);
      updatedCount++;
    } 
    // 케이스 2: 신규 유저 등록
    else {
      const newRow = new Array(16).fill("");
      newRow[0] = uniqueKey;
      newRow[1] = englishName;
      newRow[2] = koreanName;
      newRow[3] = formattedPhone;
      newRow[4] = blogUrl;
      newRow[5] = instaUrl;
      newRow[6] = youtubeUrl;
      newRow[7] = tiktokUrl;
      newRow[8] = googleGuide;
      newRow[9] = userGroup;
      newRow[10] = "";
      newRow[11] = "";
      newRow[12] = signupDate;
      newRow[13] = email;
      newRow[14] = "";
      newRow[15] = adminMemo;

      userSheet.appendRow(newRow);
      existingUserRowMap.set(uniqueKey, userSheet.getLastRow());
      addedCount++;
    }
  }

  // 🎯 [핵심 연동] Master_Log의 과거 주문 중 B열이 이메일로 남아있는 건들을 A열 고유키로 일괄 자동 승격!
  let fixedLogCount = 0;
  if (logSheet) {
    const refreshedUserData = userSheet.getDataRange().getValues();
    const emailToKeyMap = new Map();
    const keyToNameMap = new Map();
    for (let u = 1; u < refreshedUserData.length; u++) {
      const uKey = String(refreshedUserData[u][0] || '').trim();
      const uName = String(refreshedUserData[u][1] || '').trim();
      const uMail = String(refreshedUserData[u][13] || '').trim().toLowerCase();
      if (uKey && uMail) emailToKeyMap.set(uMail, uKey);
      if (uKey && uName) keyToNameMap.set(uKey, uName);
    }

    const logData = logSheet.getDataRange().getValues();
    for (let i = 2; i < logData.length; i++) {
      const rowNum = i + 1;
      const currentCode = String(logData[i][1] || '').trim();
      const currentEngName = String(logData[i][2] || '').trim();

      // B열이 이메일 형식인 경우 진짜 고유키로 승격
      if (currentCode.includes("@") && emailToKeyMap.has(currentCode.toLowerCase())) {
        const correctKey = emailToKeyMap.get(currentCode.toLowerCase());
        logSheet.getRange(rowNum, 2).setValue(correctKey); // B열: 진짜 고유키로 변경
        if (!currentEngName && keyToNameMap.has(correctKey)) {
          logSheet.getRange(rowNum, 3).setValue(keyToNameMap.get(correctKey)); // C열: 이름 채우기
        }
        fixedLogCount++;
      }
    }
  }

  SpreadsheetApp.getUi().alert(
    `✅ 동기화 완료!\n\n` +
    `• 신규 등록: ${addedCount}명\n` +
    `• 정보 업데이트: ${updatedCount}명\n` +
    `• 기존 주문의 고유키 자동 복구: ${fixedLogCount}건`
  );
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
/** [비상용] Master_Log 누락 데이터 일괄 복구 및 자동 완성 */
function fillMissingData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const logSheet = ss.getSheetByName("Master_Log");
  const userSheet = ss.getSheetByName("User_DB");
  const restSheet = ss.getSheetByName("Restaurant_List");
  
  if (!logSheet || !userSheet || !restSheet) {
    SpreadsheetApp.getUi().alert("❌ 시트 확인 필요: Master_Log, User_DB, Restaurant_List 중 누락된 시트가 있습니다.");
    return;
  }

  const logData = logSheet.getDataRange().getValues();
  const userData = userSheet.getDataRange().getValues();
  const restData = restSheet.getDataRange().getValues();

  // 1️⃣ User_DB 맵핑 (멤버코드 -> 영문 성함: B열)
  const userMap = new Map();
  for (let j = 1; j < userData.length; j++) {
    const code = String(userData[j][0] || '').trim();
    const engName = String(userData[j][1] || '').trim();
    if (code && engName) {
      userMap.set(code.toLowerCase(), engName);
    }
  }

  // 2️⃣ Restaurant_List 맵핑 (점포 ID -> 한국어 점포명: B열)
  const restMap = new Map();
  for (let k = 2; k < restData.length; k++) {
    const sId = String(restData[k][0] || '').trim().toUpperCase();
    const sName = String(restData[k][1] || '').trim();
    if (sId && sName) {
      restMap.set(sId, sName);
    }
  }

  let updated = 0;

  // 3️⃣ Master_Log 3행(인덱스 2)부터 전체 행 검사
  for (let i = 2; i < logData.length; i++) {
    const rowNum = i + 1; // 1-based 행 번호
    const memberCode = String(logData[i][1] || '').trim().toLowerCase();
    const currentEngName = String(logData[i][2] || '').trim();
    const storeId = String(logData[i][4] || '').trim().toUpperCase();
    const currentStoreName = String(logData[i][5] || '').trim();
    const visitDateVal = logData[i][7];

    // ① 영문 이름(C열) 복구
    if ((!currentEngName || currentEngName === "미승인/정보없음" || currentEngName === "undefined") && userMap.has(memberCode)) {
      logSheet.getRange(rowNum, 3).setValue(userMap.get(memberCode));
      updated++;
    }

    // ② 점포명(F열) 복구
    if ((!currentStoreName || currentStoreName === "식당명 없음" || currentStoreName === "undefined") && restMap.has(storeId)) {
      logSheet.getRange(rowNum, 6).setValue(restMap.get(storeId));
      updated++;
    }

    // ③ 방문일시(H열)가 있는 경우: 마감일(J열) 및 기본 보증금(K열) 복구
    if (visitDateVal instanceof Date && !isNaN(visitDateVal.getTime())) {
      const deadlineVal = logData[i][9];
      const depositVal = logData[i][10];

      // J열 마감일 누락 시 (방문일 + 10일)
      if (!deadlineVal || String(deadlineVal).trim() === '') {
        const d = new Date(visitDateVal.getTime());
        d.setDate(d.getDate() + 10);
        logSheet.getRange(rowNum, 10).setValue(d);
        updated++;
      }

      // K열 보증금 누락/0원 시 (10,000원 기본 세팅)
      if (!depositVal || depositVal === 0 || String(depositVal).trim() === '') {
        logSheet.getRange(rowNum, 11).setValue(10000);
        updated++;
      }
    }
  }

  SpreadsheetApp.getUi().alert(`✨ 빈칸 채우기 완료\n\n총 ${updated}건의 누락 데이터가 성공적으로 보완되었습니다.`);
}


/** [관리자 기능] 선택된 행 또는 주문번호로 점주 예약 안내 메일 재발송 */
function resendStoreBookingEmail() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const logSheet = ss.getSheetByName("Master_Log");
  const restSheet = ss.getSheetByName("Restaurant_List");
  const userSheet = ss.getSheetByName("User_DB");

  if (!logSheet || !restSheet || !userSheet) return ui.alert("❌ 시트 확인 필요");

  // 1. 현재 커서가 위치한 행 번호 확인 (또는 직접 입력 프롬프트)
  let targetRow = logSheet.getActiveRange().getRow();
  
  // 만약 선택된 행이 3행 미만이거나 Master_Log 시트가 아니면 주문번호 입력받기
  if (ss.getActiveSheet().getName() !== "Master_Log" || targetRow < 3) {
    const promptRes = ui.prompt(
      "📧 점주 메일 재발송",
      "메일을 재발송할 [행 번호] 또는 [주문번호]를 입력하세요:\n(예: 3 또는 202609297282617)",
      ui.ButtonSet.OK_CANCEL
    );
    if (promptRes.getSelectedButton() !== ui.Button.OK) return;
    const inputVal = promptRes.getResponseText().trim();
    if (!inputVal) return ui.alert("입력값이 없습니다.");

    if (!isNaN(inputVal) && parseInt(inputVal, 10) < 1000) {
      targetRow = parseInt(inputVal, 10);
    } else {
      const data = logSheet.getDataRange().getValues();
      for (let i = 2; i < data.length; i++) {
        if (String(data[i][0]).replace(/['\s]/g, '') === inputVal.replace(/['\s]/g, '')) {
          targetRow = i + 1;
          break;
        }
      }
    }
  }

  // 2. 해당 행의 예약 데이터 로드
  const rowData = logSheet.getRange(targetRow, 1, 1, 25).getValues()[0];
  const orderNo = String(rowData[0] || '').replace(/'/g, '').trim();
  const memberCode = String(rowData[1] || '').trim();
  const memberName = String(rowData[2] || '').trim();
  const storeId = String(rowData[4] || '').trim().toUpperCase();
  const rawVisitDate = rowData[7];
  const peopleStr = String(rowData[8] || '1명').replace(/[^0-9]/g, '') || '1';

  if (!orderNo || !rawVisitDate) {
    return ui.alert(`❌ [재발송 불가]\n\n${targetRow}행에 주문번호나 방문예정일시가 없습니다.`);
  }

  const timeZone = Session.getScriptTimeZone() || "Asia/Seoul";
  const dateStr = (rawVisitDate instanceof Date) ? Utilities.formatDate(rawVisitDate, timeZone, "yyyy-MM-dd") : String(rawVisitDate).substring(0, 10);
  const timeStr = (rawVisitDate instanceof Date) ? Utilities.formatDate(rawVisitDate, timeZone, "HH:mm") : String(rawVisitDate).substring(11, 16);

  // 3. 매장 정보(이메일, 매장명) 찾기
  let storeEmail = "", storeNameJp = "";
  const restData = restSheet.getDataRange().getValues();
  for (let k = 2; k < restData.length; k++) {
    if (String(restData[k][0] || '').trim().toUpperCase() === storeId) {
      storeNameJp = String(restData[k][2] || '').trim() || String(restData[k][1] || '').trim();
      storeEmail = String(restData[k][10] || '').trim();
      break;
    }
  }

  if (!storeEmail || !storeEmail.includes("@")) {
    return ui.alert(`❌ 매장 이메일 주소를 찾을 수 없습니다.\nRestaurant_List 시트의 [${storeId}] 매장 K열을 확인해 주세요.`);
  }

  // 4. 재발송 확인 창
  const confirm = ui.alert(
    "📧 점주 예약 메일 재발송",
    `[재발송 대상 정보]\n• 주문번호: #${orderNo}\n• 크리에이터: ${memberName}\n• 매장: ${storeNameJp} (${storeId})\n• 예약일시: ${dateStr} ${timeStr} (${peopleStr}명)\n• 수신처: ${storeEmail}\n\n지금 점주에게 예약 안내 메일을 재발송하시겠습니까?`,
    ui.ButtonSet.YES_NO
  );
  if (confirm !== ui.Button.YES) return;

  // 5. 메일 본문 생성 및 발송
  try {
    const scriptUrl = getActiveWebAppUrl();
    const confirmUrl = `${scriptUrl}?mode=store_confirm&row=${targetRow}&o=${encodeURIComponent(orderNo)}`;
    const feedbackUrl = `${scriptUrl}?mode=feedback&row=${targetRow}&o=${encodeURIComponent(orderNo)}`;




// 🎯 재발송 시에도 SNS 링크 순차 자동 탐색 (블로그 -> 인스타 -> 유튜브 -> 틱톡 -> 구글 로컬 가이드)
    let creatorProfileUrl = "";
    const userData = userSheet.getDataRange().getValues();
    for (let u = 1; u < userData.length; u++) {
      if (String(userData[u][0] || '').trim().toLowerCase() === memberCode.toLowerCase()) {
        const blogUrl = String(userData[u][4] || '').trim();       // E열 (블로그)
        const instaUrl = String(userData[u][5] || '').trim();      // F열 (인스타)
        const youtubeUrl = String(userData[u][6] || '').trim();    // G열 (유튜브)
        const tiktokUrl = String(userData[u][7] || '').trim();     // H열 (틱톡)
        const googleGuide = String(userData[u][8] || '').trim();   // I열 (구글 로컬 가이드)

        creatorProfileUrl = blogUrl || instaUrl || youtubeUrl || tiktokUrl || googleGuide || "";
        break;
      }
    }

    let profileHtml = `<p style="margin: 5px 0; font-size: 15px; color: #8b95a1;"><strong>&#128279; <span>SNS:</span></strong> <span>当日確認</span></p>`;

    if (creatorProfileUrl) {
      const isHttp = /^https?:\/\//i.test(creatorProfileUrl);
      const isDomainLike = /^[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}/.test(creatorProfileUrl);

      if (isHttp || isDomainLike) {
        const fullUrl = isHttp ? creatorProfileUrl : "https://" + creatorProfileUrl;
        const btnText = (creatorProfileUrl.includes('drive.google.com') || creatorProfileUrl.includes('imweb') || creatorProfileUrl.match(/\.(jpg|jpeg|png|webp|gif)/i))
          ? "プロフィール確認❯" 
          : "SNSを見る❯";
        profileHtml = `<p style="margin: 5px 0; font-size: 15px;"><strong>&#128279; <span>SNS:</span></strong> <a href="${fullUrl}" target="_blank" style="color: #1a73e8; font-weight: bold; text-decoration: underline;"><span>${btnText}</span></a></p>`;
      } else {
        profileHtml = `<p style="margin: 5px 0; font-size: 15px;"><strong>&#128279; <span>SNS:</span></strong> <span style="color: #1A2B49; font-weight: bold;">Google Local Guides (${creatorProfileUrl})</span></p>`;
      }
    }



// 운영진 이메일 목록
    const adminAlertEmails = (typeof getAdminAlertEmails === 'function') ? getAdminAlertEmails() : "bookmarkjapan.info@gmail.com";

    const subject = `【BOOKMARK CREATORS】 クリエイター来店予約の確認(${dateStr})`;
    const htmlBody = `
      <meta charset="UTF-8">
      <div style="font-family: 'Helvetica Neue', Arial, sans-serif; padding: 20px; background-color: #f4f5f7;">
        <div style="max-width: 500px; margin: 0 auto; background: #ffffff; border-radius: 12px; padding: 30px; box-shadow: 0 4px 10px rgba(0,0,0,0.05);">
          <div style="text-align: center; margin-bottom: 20px;">
            <h1 style="color: #1A2B49; margin: 0; font-size: 24px; font-weight: 900; letter-spacing: -0.5px;">BOOKMARK CREATORS</h1>
            <div style="width: 40px; height: 3px; background: #C5A358; margin: 10px auto;"></div>
          </div>
          <h2 style="color: #1A2B49; margin-top: 0; font-size: 18px; border-bottom: 2px solid #f1f3f5; padding-bottom: 15px; text-align: center;">&#128197; 来店予約の依頼</h2>
          <div style="margin-top: 20px;">
            <p style="color: #1A2B49; font-size: 16px; font-weight: bold; margin-bottom: 5px;">${storeNameJp}</p>
            <p style="color: #495057; font-size: 14px; margin-top: 0;">店舗管理者様</p>
          </div>
          <p style="color: #495057; font-size: 14px; line-height: 1.6;">BOOKMARK CREATORSより、クリエイターの訪問予約申請が届きました。内容をご確認の上、以下のボタンより確定または日時変更のご対応をお願いいたします。</p>
          <div style="background-color: #f8f9fa; border-left: 4px solid #C5A358; padding: 15px; border-radius: 4px; margin: 20px 0;">
            <p style="margin: 5px 0; font-size: 15px;"><strong>&#128100; クリエイター名:</strong> ${memberName}</p>
            <p style="margin: 5px 0; font-size: 15px;"><strong>&#9200; 訪問日時:</strong> <span style="color: #d63384; font-weight: bold;">${dateStr} ${timeStr}</span></p>
            <p style="margin: 5px 0; font-size: 15px;"><strong>&#128101; 訪問人数:</strong> <span style="color: #1A2B49; font-weight: bold;">${peopleStr}名</span></p>
            ${profileHtml}
          </div>
          <div style="margin: 30px 0; text-align: center;">
            <a href="${confirmUrl}" target="_blank" style="background-color: #2D6A4F; color: #ffffff; padding: 14px 20px; font-size: 14px; font-weight: bold; text-decoration: none; border-radius: 10px; display: inline-block; margin-right: 10px; box-shadow: 0 4px 12px rgba(45,106,79,0.2);">&#9989; 予約を確定する</a>
            <a href="${feedbackUrl}" target="_blank" style="background-color: #1A2B49; color: #ffffff; padding: 14px 20px; font-size: 14px; font-weight: bold; text-decoration: none; border-radius: 10px; display: inline-block; box-shadow: 0 4px 12px rgba(26,43,73,0.15);">&#128260; 日時変更をリクエスト</a>
          </div>
        </div>
      </div>`;

    GmailApp.sendEmail(storeEmail, subject, "", {
      htmlBody: htmlBody,
      name: "BOOKMARK CREATORS",
      from: "info@bookmarkfukuoka.jp"
    });

    logSheet.getRange(targetRow, 25).setValue("점주메일 재발송완료");
    ui.alert(`✅ [재발송 성공]\n\n• 매장: ${storeNameJp}\n• 수신 메일: ${storeEmail}\n\n점주에게 예약 승인 메일이 성공적으로 전송되었습니다!`);
  } catch (err) {
    logSheet.getRange(targetRow, 25).setValue("❌ 재발송 실패: " + err.toString());
    ui.alert(`❌ 재발송 실패: ${err.toString()}`);
  }
}
