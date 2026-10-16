/*****************************************************************************************
 * ⚡ [엔진 완전 최적화] 아임웹 실제 수신 페이로드 정합성 100% 동기화 시스템 (주문/가입/취소 통합)
 ****************************************************************************************/
function doPost(e) {
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

    const eventType = String(postData.eventType || postData.event_type || "").trim();
    const dataObj = postData.data || postData;

    // 🚨 0️⃣ 아임웹 주문 취소/환불 웹훅 (ORDER_CANCEL_COMPLETE)
    if (eventType === "ORDER_CANCEL_COMPLETE" || eventType.includes("CANCEL")) {
      const orderNo = dataObj.orderNo || dataObj.order_no || postData.orderNo || postData.order_no;
      const cancelReason = dataObj?.section?.cancelInfo?.cancelReason || "아임웹 취소";
      const result = handleImwebCancellation(orderNo, cancelReason);
      return ContentService.createTextOutput(JSON.stringify(result)).setMimeType(ContentService.MimeType.JSON);
    }

    // 1️⃣ 아임웹 회원가입 이벤트 (END_USER_SIGN_UP)
    if (eventType === "END_USER_SIGN_UP") {
      const newMemberUid = String(dataObj.memberUid || "").trim();
      if (newMemberUid && userSheet) {
        const uValues = userSheet.getDataRange().getValues();
        let exists = false;
        for (let u = 1; u < uValues.length; u++) {
          if (String(uValues[u][0]).trim() === newMemberUid || String(uValues[u][13]).trim() === newMemberUid) {
            exists = true; break;
          }
        }
        if (!exists) {
          const newUserRow = new Array(16).fill("");
          newUserRow[0] = newMemberUid;                   // A: 고유키(이메일UID)
          newUserRow[1] = "신규가입(정보수집필요)";         // B: 이름
          newUserRow[13] = newMemberUid;                  // N: 이메일
          userSheet.appendRow(newUserRow);
        }
      }
      return ContentService.createTextOutput(JSON.stringify({"result": "success_signup"})).setMimeType(ContentService.MimeType.JSON);
    }

    // 2️⃣ 아임웹 주문 생성 이벤트 (ORDER_CREATE / ORDER_PAY_COMPLETE)
    const rawOrderNo = String(dataObj.orderNo || dataObj.order_no || "번호없음").trim();
    const cleanOrderNo = rawOrderNo.replace(/['"\s]/g, '');
    const totalPrice = Number(dataObj.totalPaymentPrice || (dataObj.payment && dataObj.payment.paidPrice) || 0);
    
    const memberUid = String(dataObj.memberUid || dataObj.member_code || "").trim();
    const ordererName = String(dataObj.ordererName || dataObj.orderer_name || "").trim();
    const ordererCall = String(dataObj.ordererCall || dataObj.orderer_phone || "").replace(/[^0-9]/g, "");

    // 3️⃣ User_DB 대조 및 회원 정보 확보
    let englishName = ordererName;
    let finalMemberKey = memberUid;

    if (userSheet) {
      const uData = userSheet.getDataRange().getValues();
      let foundUser = false;
      for (let j = 1; j < uData.length; j++) {
        const dbKey = String(uData[j][0] || '').trim();
        const dbEmail = String(uData[j][13] || '').trim();
        
        if ((memberUid && dbKey === memberUid) || (memberUid && dbEmail === memberUid)) {
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

    // 4️⃣ 상품 목록(sections) 안전 추출
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

    // 5️⃣ Restaurant_List 사전 구축 (정확한 매칭을 위해 매장명 길이순 정렬)
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
    // 긴 이름부터 매칭되도록 정렬 (오매칭 원천 방지)
    restIdList.sort((a, b) => b.nameKo.length - a.nameKo.length);

    // 6️⃣ items 순회하며 Master_Log 기록 (복수 주문 완전 독립 행 등록)
    const itemDeposit = (items.length > 0 && totalPrice > 0) ? Math.floor(totalPrice / items.length) : 10000;
    
    // Master_Log 기존 데이터 1회 로드
    let logValues = logSheet.getDataRange().getValues();

    items.forEach(function(item) {
      const rawProdName = (item.productInfo && item.productInfo.prodName) || item.product_name || item.name || "식당명 없음";
      
      // 대괄호 수식어구 정리 (예: "[하카타 모츠나베와 야키토리] 하카타후도 본점" -> "하카타후도 본점")
      const cleanProd = rawProdName.replace(/\[.*?\]/g, '').replace(/\s+/g, '').toLowerCase();

      let targetStoreId = "";
      let matchedStoreName = rawProdName;

      // 1순위: 완전 일치 검색 / 2순위: 부분 일치
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

      // 🔒 [복수 상품 판별 락] 주문번호 AND (점포ID OR 매장명) 모두 일치해야 기존 행으로 인식
      let existingRowIndex = -1;
      for (let i = 2; i < logValues.length; i++) {
        const sheetOrder = String(logValues[i][0] || '').replace(/['"\s]/g, '');
        const sheetStoreId = String(logValues[i][4] || '').trim().toUpperCase();
        const sheetStoreName = String(logValues[i][5] || '').trim();

        if (sheetOrder === cleanOrderNo) {
          if ((targetStoreId && sheetStoreId === targetStoreId) || 
              (matchedStoreName && sheetStoreName === matchedStoreName)) {
            existingRowIndex = i + 1;
            break;
          }
        }
      }

      let status = "예약대기"; 
      if (rawData.includes("CANCEL") || rawData.includes("REFUND") || eventType === "ORDER_CANCEL") {
        status = "취소완료";
      }

      if (existingRowIndex > 0) {
        // 이미 등록된 '해당 식당'인 경우 상태만 업데이트
        logSheet.getRange(existingRowIndex, 12).setValue(status); 
        if (status === "취소완료") {
          logSheet.getRange(existingRowIndex, 21).setValue("취소환불"); 
        }
      } else {
        // 신규 행 추가
        const newRow = new Array(25).fill(""); 
        newRow[0] = "'" + cleanOrderNo;      // A: 주문번호
        newRow[1] = finalMemberKey;          // B: 멤버코드
        newRow[2] = englishName;             // C: 성함
        newRow[3] = "";                      // D: 참여 채널
        newRow[4] = targetStoreId;           // E: 점포 ID
        newRow[5] = matchedStoreName;        // F: 점포명
        newRow[6] = "";                      // G: 예약 캡처 URL
        newRow[7] = "";                      // H: 방문예정일시
        newRow[8] = "";                      // I: 방문인원
        newRow[9] = "";                      // J: 리뷰 마감 기한
        newRow[10] = itemDeposit || 10000;   // K: 보증금액 (분할 분배)
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

        // 방금 추가된 행을 메모리 상의 logValues에도 푸시하여 중복 방어
        logValues.push(newRow);
      }
    });

    return ContentService.createTextOutput(JSON.stringify({"result": "success"})).setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    if (debugSheet) debugSheet.appendRow([new Date(), "에러: " + err.toString()]);
    return ContentService.createTextOutput(JSON.stringify({"result": "error", "error": err.toString()})).setMimeType(ContentService.MimeType.JSON);
  }
}

/** 🎯 아임웹 주문 취소 시트 반영 엔진 */
function handleImwebCancellation(orderNo, reason) {
  try {
    if (!orderNo) return { success: false, error: "주문번호 누락" };

    const targetOrderNo = String(orderNo).trim().replace(/['"\s]/g, '');
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('Master_Log');
    if (!sheet) return { success: false, error: "Master_Log 시트 없음" };

    const lastRow = sheet.getLastRow();
    if (lastRow < 3) return { success: false, error: "데이터 없음" };

    const orderRange = sheet.getRange(3, 1, lastRow - 2, 1).getValues();
    const timeZone = Session.getScriptTimeZone();
    const timeStr = Utilities.formatDate(new Date(), timeZone, "yyyy-MM-dd HH:mm");
    let canceledCount = 0;

    // 🎯 하나의 주문번호에 묶인 모든 매장 행을 일괄 취소 처리
    for (let i = 0; i < orderRange.length; i++) {
      const sheetOrder = String(orderRange[i][0] || '').trim().replace(/['"\s]/g, '');
      if (sheetOrder === targetOrderNo) {
        const targetRow = i + 3;
        sheet.getRange(targetRow, 12).setValue('취소(아임웹)');
        sheet.getRange(targetRow, 8, 1, 3).clearContent(); // 슬롯 반환

        const prevMemo = String(sheet.getRange(targetRow, 13).getValue() || '').trim();
        const cancelNote = `[${timeStr} 아임웹 취소: ${reason}]`;
        sheet.getRange(targetRow, 13).setValue(prevMemo ? `${prevMemo} | ${cancelNote}` : cancelNote);
        canceledCount++;
      }
    }

    if (canceledCount === 0) {
      console.warn(`[취소 반영 실패] Master_Log에서 주문번호 #${targetOrderNo} 를 찾을 수 없음`);
      return { success: false, error: `주문번호 #${targetOrderNo} 없음` };
    }

    return { success: true, count: canceledCount, orderNo: targetOrderNo };

  } catch (e) {
    console.error("handleImwebCancellation 실패: " + e.toString());
    return { success: false, error: e.toString() };
  }
}
