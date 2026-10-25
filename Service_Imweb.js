/** 아임웹 연동: 웹훅(주문/취소/가입) 처리, 회원 동기화 */

/*****************************************************************************************
 * ⚡ [엔진 완전 최적화] 아임웹 실제 수신 페이로드 정합성 100% 동기화 시스템 (주문/가입/취소/구매옵션 통합)
 ****************************************************************************************/
function handleImwebWebhook(e) {
  // 🔒 웹훅 동시 인입(ORDER_CREATE & ORDER_PAY_COMPLETE 동시 도달) 시 중복 등록 방지 Lock
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(10000)) {
    return ContentService.createTextOutput(JSON.stringify({ result: "busy_retry" })).setMimeType(ContentService.MimeType.JSON);
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const logSheet = ss.getSheetByName("Master_Log");
  const userSheet = ss.getSheetByName("User_DB");
  const restSheet = ss.getSheetByName("Restaurant_List"); 
  const debugSheet = ss.getSheetByName("Log"); 

  try {
    if (!e || !e.postData || !e.postData.contents) {
      return ContentService.createTextOutput("NO_DATA");
    }

    const rawData = e.postData.contents;
    const postData = JSON.parse(rawData);
    if (debugSheet) debugSheet.appendRow([new Date(), "수신: " + rawData]);

    const eventType = String(postData.eventType || postData.event_type || "").trim().toUpperCase();
    const dataObj = postData.data || postData;

    // 🚨 0️⃣ 아임웹 주문 취소 및 환불 웹훅 (CANCEL 또는 REFUND 키워드 포괄 방어)
    if (eventType.includes("CANCEL") || eventType.includes("REFUND")) {
      const orderNo = dataObj.orderNo || dataObj.order_no || postData.orderNo || postData.order_no;
      const cancelReason = dataObj?.section?.cancelInfo?.cancelReason || dataObj?.cancelReason || "아임웹 취소/환불";
      const result = handleImwebCancellation(orderNo, cancelReason);
      return ContentService.createTextOutput(JSON.stringify(result)).setMimeType(ContentService.MimeType.JSON);
    }

    // 1️⃣ 아임웹 회원가입 이벤트 (END_USER_SIGN_UP)
    if (eventType === "END_USER_SIGN_UP") {
      const newMemberUid = String(dataObj.memberUid || dataObj.member_code || "").trim();
      if (newMemberUid && userSheet) {
        const uValues = userSheet.getDataRange().getValues();
        let exists = false;
        for (let u = 1; u < uValues.length; u++) {
          if (String(uValues[u][0]).trim().toLowerCase() === newMemberUid.toLowerCase() || 
              String(uValues[u][13]).trim().toLowerCase() === newMemberUid.toLowerCase()) {
            exists = true; 
            break;
          }
        }
        if (!exists) {
          const newUserRow = new Array(16).fill("");
          newUserRow[0] = newMemberUid;                  // A: 고유키(이메일UID)
          newUserRow[1] = "新規会員";                      // B: 이름
          newUserRow[13] = newMemberUid;                  // N: 이메일
          userSheet.appendRow(newUserRow);
        }
      }
      return ContentService.createTextOutput(JSON.stringify({ result: "success_signup" })).setMimeType(ContentService.MimeType.JSON);
    }

    // 2️⃣ 아임웹 주문 생성/결제완료 이벤트 (ORDER_CREATE / ORDER_PAY_COMPLETE)
    const rawOrderNo = String(dataObj.orderNo || dataObj.order_no || "번호없음").trim();
    const cleanOrderNo = rawOrderNo.replace(/['"\s]/g, '');
    const totalPrice = Number(dataObj.totalPaymentPrice || (dataObj.payment && dataObj.payment.paidPrice) || 0);
    
    const memberUid = String(dataObj.memberUid || dataObj.member_code || "").trim();
    const ordererName = String(dataObj.ordererName || dataObj.orderer_name || "").trim();
    const ordererCall = String(dataObj.ordererCall || dataObj.orderer_phone || "").replace(/[^0-9]/g, "");

    // 3️⃣ User_DB 대조 및 회원 정보 동기화 (전화번호/이름 보정)
    let englishName = ordererName;
    let finalMemberKey = memberUid;

    if (userSheet) {
      const uData = userSheet.getDataRange().getValues();
      let foundUser = false;
      for (let j = 1; j < uData.length; j++) {
        const dbKey = String(uData[j][0] || '').trim();
        const dbEmail = String(uData[j][13] || '').trim();
        
        if ((memberUid && dbKey.toLowerCase() === memberUid.toLowerCase()) || 
            (memberUid && dbEmail.toLowerCase() === memberUid.toLowerCase())) {
          finalMemberKey = dbKey;
          englishName = String(uData[j][1] || ordererName).trim();
          
          if (!uData[j][3] && ordererCall) {
            userSheet.getRange(j + 1, 4).setValue(ordererCall);
          }
          foundUser = true;
          break;
        }
      }

      if (!foundUser && memberUid) {
        const newUserRow = new Array(16).fill("");
        newUserRow[0] = memberUid;
        newUserRow[1] = ordererName;
        newUserRow[2] = ordererName;
        newUserRow[3] = ordererCall;
        newUserRow[13] = memberUid;
        newUserRow[15] = "주문시 자동안착";
        userSheet.appendRow(newUserRow);
      }
    }

    // 4️⃣ 상품 목록(sections / items) 정규화 추출
    let items = [];
    if (dataObj.sections && Array.isArray(dataObj.sections)) {
      dataObj.sections.forEach(sec => {
        if (sec.sectionItems && Array.isArray(sec.sectionItems)) {
          items = items.concat(sec.sectionItems);
        }
      });
    } else if (dataObj.items && Array.isArray(dataObj.items)) {
      items = dataObj.items;
    }

    // 5️⃣ Restaurant_List 사전 구축 (긴 이름 우선 매칭)
    const restData = restSheet ? restSheet.getDataRange().getValues() : [];
    const restIdList = [];
    for (let k = 2; k < restData.length; k++) {
      const rId = String(restData[k][0] || '').trim().toUpperCase();
      const rNameKo = String(restData[k][1] || '').trim();
      const rNameJp = String(restData[k][2] || '').trim();
      if (rId && (rNameKo || rNameJp)) {
        restIdList.push({ id: rId, nameKo: rNameKo, nameJp: rNameJp });
      }
    }
    restIdList.sort((a, b) => b.nameKo.length - a.nameKo.length);

    // 6️⃣ Master_Log 기존 데이터를 단 1번만 로드하여 고속 매핑 배열 구축
    const logData = logSheet.getDataRange().getValues();
    const existingOrderRows = [];

    for (let i = 2; i < logData.length; i++) {
      const sheetOrder = String(logData[i][0] || '').replace(/['"\s]/g, '');
      const sheetStoreId = String(logData[i][4] || '').trim().toUpperCase();
      const sheetStoreName = String(logData[i][5] || '').trim();
      const sheetOption = String(logData[i][6] || '').trim(); // 🎯 G열 구매옵션

      if (sheetOrder === cleanOrderNo) {
        existingOrderRows.push({
          rowNum: i + 1,
          storeId: sheetStoreId,
          storeName: sheetStoreName,
          option: sheetOption,
          status: String(logData[i][11] || '').trim(),
          used: false
        });
      }
    }

    const itemDeposit = (items.length > 0 && totalPrice > 0) ? Math.floor(totalPrice / items.length) : 10000;

    items.forEach(function(item) {
      const rawProdName = (item.productInfo && item.productInfo.prodName) || item.product_name || item.name || "식당명 없음";
      
      // 🎯 [수정 후] 아임웹 optionInfo 객체 파싱 및 원본 옵션명 보존
      let rawOption = "";
      if (item.productInfo && item.productInfo.optionInfo) {
        rawOption = Object.values(item.productInfo.optionInfo).join(' ');
      } else if (item.optionInfo) {
        rawOption = Object.values(item.optionInfo).join(' ');
      } else if (item.selectedOption && item.selectedOption.name) {
        rawOption = item.selectedOption.name;
      } else if (item.options && Array.isArray(item.options)) {
        rawOption = item.options.map(o => o.value || o.name || '').join(' ');
      }

      // 원본 옵션명 우선 보존 (비어있을 때만 상품명에서 런치/디너 추출)
      let mealOption = rawOption.trim();
      if (!mealOption) {
        const lowerProd = rawProdName.toLowerCase();
        if (lowerProd.includes("런치") || lowerProd.includes("lunch")) mealOption = "런치";
        else if (lowerProd.includes("디너") || lowerProd.includes("dinner")) mealOption = "디너";
      }

      const cleanProd = rawProdName.replace(/\[.*?\]/g, '').replace(/\s+/g, '').toLowerCase();

      let targetStoreId = "";
      let matchedStoreName = rawProdName;

      for (let m = 0; m < restIdList.length; m++) {
        const cleanKo = restIdList[m].nameKo.replace(/\s+/g, '').toLowerCase();
        const cleanJp = restIdList[m].nameJp.replace(/\s+/g, '').toLowerCase();

        if (cleanKo === cleanProd || (cleanJp && cleanJp === cleanProd)) {
          targetStoreId = restIdList[m].id;
          matchedStoreName = restIdList[m].nameKo;
          break;
        }
      }

      if (!targetStoreId) {
        for (let m = 0; m < restIdList.length; m++) {
          const cleanKo = restIdList[m].nameKo.replace(/\s+/g, '').toLowerCase();
          if (cleanProd.includes(cleanKo) || cleanKo.includes(cleanProd)) {
            targetStoreId = restIdList[m].id;
            matchedStoreName = restIdList[m].nameKo;
            break;
          }
        }
      }

      // 🔒 [중복 검사] 주문번호 + 매장 + 구매옵션까지 일치해야 동일 건으로 인식
      let matchedExisting = null;
      for (let ex of existingOrderRows) {
        if (!ex.used) {
          const isStoreMatch = (targetStoreId && ex.storeId === targetStoreId) || 
                               (matchedStoreName && ex.storeName === matchedStoreName);
          const isOptionMatch = (!mealOption && !ex.option) || (mealOption === ex.option);

          if (isStoreMatch && isOptionMatch) {
            matchedExisting = ex;
            ex.used = true;
            break;
          }
        }
      }

      let status = "예약대기";
      if (rawData.includes("CANCEL") || rawData.includes("REFUND") || eventType.includes("CANCEL")) {
        status = "취소완료";
      }

      if (matchedExisting) {
        // 🔒 [치명적 덮어쓰기 방어 락] 이미 방문완료/제출완료이거나 이미 취소된 건은 건드리지 않음
        const currentSheetStatus = matchedExisting.status;
        if (currentSheetStatus === '방문완료' || currentSheetStatus === '제출완료' || currentSheetStatus.includes('취소')) {
          return;
        }

        // 방문 전이나 미신청 대기 상태인 행만 안전하게 취소 처리
        if (status === "취소완료") {
          logSheet.getRange(matchedExisting.rowNum, 12).setValue("취소(아임웹)");
          logSheet.getRange(matchedExisting.rowNum, 8, 1, 3).clearContent();
          logSheet.getRange(matchedExisting.rowNum, 21).setValue("취소환불");
        }
      } else {
        // 신규 주문 행 추가
        const newRow = new Array(25).fill(""); 
        newRow[0] = "'" + cleanOrderNo;      // A: 주문번호
        newRow[1] = finalMemberKey;          // B: 멤버코드
        newRow[2] = englishName;             // C: 성함
        newRow[3] = "";                      // D: 참여 채널
        newRow[4] = targetStoreId;           // E: 점포 ID
        newRow[5] = matchedStoreName;        // F: 점포명
        newRow[6] = mealOption;              // 🎯 G: 구매옵션 (A코스, 런치, B코스 등 있는 그대로 기록)
        newRow[7] = "";                      // H: 방문예정일시
        newRow[8] = "";                      // I: 방문인원
        newRow[9] = "";                      // J: 리뷰 마감 기한
        newRow[10] = itemDeposit || 10000;   // K: 보증금액
        newRow[11] = status;                 // L: 진행 상태 ("예약대기")
        newRow[12] = "";                     // M: 점주 피드백
        newRow[13] = "";                     // N: 방문 확인
        newRow[14] = "";                     // O: 매장 영수증
        newRow[15] = "";                     // P: 콘텐츠 제출일
        newRow[16] = "";                     // Q: 제출 콘텐츠 URL
        newRow[17] = "";                     // R: 구글맵 리뷰 URL
        newRow[18] = "";                     // S: 한 줄 후기
        newRow[19] = "";                     // T: (JP)한 줄 후기
        newRow[20] = "";                     // U: 환불 처리

        logSheet.appendRow(newRow);
        
        existingOrderRows.push({
          rowNum: logSheet.getLastRow(),
          storeId: targetStoreId,
          storeName: matchedStoreName,
          option: mealOption,
          status: status,
          used: true
        });
      }
    });

    SpreadsheetApp.flush();
    return ContentService.createTextOutput(JSON.stringify({ result: "success" })).setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    if (debugSheet) debugSheet.appendRow([new Date(), "에러: " + err.toString()]);
    return ContentService.createTextOutput(JSON.stringify({ result: "error", error: err.toString() })).setMimeType(ContentService.MimeType.JSON);
  } finally {
    lock.releaseLock();
  }
}

/** 🎯 아임웹 취소/환불 연동 (체험 완료 건 보호 락 탑재) */
function handleImwebCancellation(orderNo, reason) {
  try {
    if (!orderNo) return { success: false, error: "주문번호 누락" };

    const targetOrderNo = String(orderNo).trim().replace(/['"\s]/g, '');
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('Master_Log');
    if (!sheet) return { success: false, error: "Master_Log 시트 없음" };

    const lastRow = sheet.getLastRow();
    if (lastRow < 3) return { success: false, error: "검사할 데이터 행이 없습니다." };

    const totalDataRows = lastRow - 2;
    // A열(주문번호), L열(상태), M열(메모) 데이터 로드
    const logData = sheet.getRange(3, 1, totalDataRows, 13).getValues();
    const timeZone = Session.getScriptTimeZone() || "Asia/Seoul";
    const timeStr = Utilities.formatDate(new Date(), timeZone, "yyyy-MM-dd HH:mm");
    let canceledCount = 0;

    for (let i = 0; i < totalDataRows; i++) {
      const sheetOrder = String(logData[i][0] || '').trim().replace(/['"\s]/g, '');
      const currentStatus = String(logData[i][11] || '').trim();

      if (sheetOrder === targetOrderNo) {
        // 🔒 [치명적 버그 방어 락] 이미 방문완료/제출완료/취소 건 방어
        if (currentStatus === '방문완료' || currentStatus === '제출완료' || currentStatus.includes('취소')) {
          continue; 
        }

        // 방문 전이거나 승인 대기 중인 행만 안전하게 취소 처리
        const targetRow = i + 3;
        sheet.getRange(targetRow, 12).setValue('취소(아임웹)');
        sheet.getRange(targetRow, 8, 1, 3).clearContent(); // H~J열 슬롯 반환

        const prevMemo = String(logData[i][12] || '').trim();
        const cancelNote = `[${timeStr} 아임웹 취소: ${reason}]`;
        sheet.getRange(targetRow, 13).setValue(prevMemo ? `${prevMemo} | ${cancelNote}` : cancelNote);

        canceledCount++;
      }
    }

    SpreadsheetApp.flush();
    return { success: true, count: canceledCount, orderNo: targetOrderNo };

  } catch (e) {
    console.error("handleImwebCancellation 실패: " + e.toString());
    return { success: false, error: e.toString() };
  }
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

    if (existingUserRowMap.has(uniqueKey)) {
      const targetRow = existingUserRowMap.get(uniqueKey);
      if (englishName) userSheet.getRange(targetRow, 2).setValue(englishName);
      if (koreanName) userSheet.getRange(targetRow, 3).setValue(koreanName);
      if (formattedPhone) userSheet.getRange(targetRow, 4).setValue(formattedPhone);

      if (col("블로그", "blog", "네이버") !== -1) userSheet.getRange(targetRow, 5).setValue(blogUrl);
      if (col("인스타", "insta", "instagram", "릴스") !== -1) userSheet.getRange(targetRow, 6).setValue(instaUrl);
      if (col("유튜브", "youtube") !== -1) userSheet.getRange(targetRow, 7).setValue(youtubeUrl);
      if (col("틱톡", "tiktok", "tictok") !== -1) userSheet.getRange(targetRow, 8).setValue(tiktokUrl);
      if (col("구글", "로컬", "가이드") !== -1) userSheet.getRange(targetRow, 9).setValue(googleGuide);

      if (userGroup) userSheet.getRange(targetRow, 10).setValue(userGroup);
      if (email) userSheet.getRange(targetRow, 14).setValue(email);
      if (adminMemo) userSheet.getRange(targetRow, 16).setValue(adminMemo);
      updatedCount++;
    } else {
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

      if (currentCode.includes("@") && emailToKeyMap.has(currentCode.toLowerCase())) {
        const correctKey = emailToKeyMap.get(currentCode.toLowerCase());
        logSheet.getRange(rowNum, 2).setValue(correctKey);
        if (!currentEngName && keyToNameMap.has(correctKey)) {
          logSheet.getRange(rowNum, 3).setValue(keyToNameMap.get(correctKey));
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

/** 🧪 아임웹 주문 취소 웹훅 가상 테스트 함수 */
function testCancelWebhook() {
  // ⚠️ 중요: 현재 Master_Log 시트 A열에 실제로 존재하는 주문번호 하나를 입력하세요 (예: 2222 또는 실제 주문번호)
  const testOrderNo = "2222"; 

  // 아임웹에서 실제로 보내는 취소 웹훅 규격 가상 데이터
  const samplePayload = {
    eventType: "ORDER_CANCEL_COMPLETE",
    eventTime: 1791274411890,
    data: {
      orderNo: testOrderNo,
      section: {
        cancelInfo: {
          cancelReason: "테스트 단순 변심 취소"
        }
      }
    }
  };

  const mockEvent = {
    postData: {
      contents: JSON.stringify(samplePayload)
    }
  };

  // doPost 직접 호출 실행
  const response = doPost(mockEvent);
  Logger.log("실행 결과: " + response.getContent());
}
