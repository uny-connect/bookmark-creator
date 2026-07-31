/*****************************************************************************************
 * ⚡ [엔진 완전 최적화] 아임웹 신규 주문 및 취소/환불 웹훅 실시간 시트 적재 시스템 (최종)
 ****************************************************************************************/
function doPost(e) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const logSheet = ss.getSheetByName("Master_Log");
  const userSheet = ss.getSheetByName("User_DB");
  const restSheet = ss.getSheetByName("Restaurant_List"); 
  const debugSheet = ss.getSheetByName("Log"); 

  try {
    const rawData = e.postData.contents;
    const postData = JSON.parse(rawData);
    if (debugSheet) debugSheet.appendRow([new Date(), "수신: " + rawData]);

    const eventType = postData.eventType || "";
    
    // 1️⃣ 아임웹 회원가입 이벤트(END_USER_SIGN_UP) 캐치
    if (eventType === "END_USER_SIGN_UP" || rawData.includes("END_USER_SIGN_UP")) {
      const newMemberUid = (postData.data && postData.data.memberUid) ? postData.data.memberUid : "";
      if (newMemberUid && userSheet) {
        const newUserRow = new Array(15).fill("");
        newUserRow[0] = newMemberUid;
        newUserRow[1] = "신규가입(정보수집필요)"; 
        userSheet.appendRow(newUserRow);
      }
      return ContentService.createTextOutput(JSON.stringify({"result": "success_signup"})).setMimeType(ContentService.MimeType.JSON);
    }

    // 2️⃣ 아임웹 주문 이벤트 처리 (v2 구조 대응)
    const dataObj = postData.data || postData; 

    const orderNo = dataObj.orderNo || dataObj.order_no || "번호없음";
    const memberCode = dataObj.memberUid || dataObj.member_code || (dataObj.member && dataObj.member.code) || "";
    const totalPrice = dataObj.totalPaymentPrice || (dataObj.payment && dataObj.payment.paidPrice) || 0;
    
    // 3️⃣ 상품(items) 배열 추출 
    let items = dataObj.items || [];
    if (items.length === 0 && dataObj.sections && Array.isArray(dataObj.sections)) {
      dataObj.sections.forEach(section => {
        if (section.sectionItems && Array.isArray(section.sectionItems)) {
          items = items.concat(section.sectionItems);
        }
      });
    } else if (items.length === 0 && dataObj.section && dataObj.section.sectionItems) {
      items = dataObj.section.sectionItems;
    }

    // 4️⃣ 매장 ID 추적용 사전 맵 구축
    const restData = restSheet ? restSheet.getDataRange().getValues() : [];
    const restIdMap = {};
    for (let k = 2; k < restData.length; k++) {
      const rId = String(restData[k][0]).trim();
      const rNameKo = String(restData[k][1]).trim();
      if (rNameKo && rId) {
        restIdMap[rNameKo] = rId; 
      }
    }

    // 5️⃣ 추출된 아이템들을 순회하며 Master_Log에 기록
    items.forEach(function(item) {
      const productName = item.product_name || item.name || (item.productInfo && item.productInfo.prodName) || "식당명 없음";
      
      const values = logSheet.getDataRange().getValues();
      let rowIndex = -1;
      for (let i = 1; i < values.length; i++) {
        if (String(values[i][0]).replace(/'/g, '') === String(orderNo)) { // 상품명 비교 제거(수식 충돌 방지)
          rowIndex = i + 1;
          break;
        }
      }

      let status = "예약대기"; 
      if (rawData.includes("CANCEL") || rawData.includes("REFUND") || eventType === "ORDER_CANCEL") {
        status = "취소완료";
      }

      if (rowIndex > 0) {
        // 기존 주문 업데이트
        logSheet.getRange(rowIndex, 12).setValue(status); 
        if (status === "취소완료") {
          logSheet.getRange(rowIndex, 21).setValue("취소환불"); 
        }
      } else {
        // 신규 주문 추가
        
        // 🎯 [신규 로직] User_DB를 검색하여 이메일 대신 '고유키(m202...)' 찾아오기
        let uniqueKey = memberCode; // 기본값은 넘어온 이메일
        if (memberCode && userSheet) {
          const userData = userSheet.getDataRange().getValues();
          for (let j = 1; j < userData.length; j++) {
            // N열(인덱스 13)에 있는 이메일과 비교하여 일치하면 A열(인덱스 0)의 고유키 추출
            if (String(userData[j][13]) === String(memberCode) || String(userData[j][0]) === String(memberCode)) {
              uniqueKey = userData[j][0];
              break;
            }
          }
        }

        const targetStoreId = restIdMap[productName] || "";

        const newRow = new Array(24).fill(""); 
        newRow[0] = "'" + orderNo;   // A: 주문번호
        newRow[1] = uniqueKey;       // B: 멤버코드 (고유키로 자동 변환됨!)
        newRow[2] = "";              // C: 영문 성함 (요청하신 대로 공백 처리)
        newRow[3] = "";              // D: 참여 채널
        newRow[4] = targetStoreId;   // E: 점포 ID
        newRow[5] = "";              // F: 점포명 (시트 수식이 작동하도록 공백 처리)
        newRow[6] = "";              // G: 예약 캡처 URL
        newRow[7] = "";              // H: 방문예정일시
        newRow[8] = "";              // I: 방문인원
        newRow[9] = "";              // J: 리뷰 마감 기한
        newRow[10] = totalPrice;     // K: 보증금액
        newRow[11] = status;         // L: 진행 상태
        newRow[12] = "";             // M: 점주 피드백
        newRow[13] = "";             // N: 방문 확인
        newRow[14] = "";             // O: 매장 영수증
        newRow[15] = "";             // P: 콘텐츠 제출일
        newRow[16] = "";             // Q: 제출 콘텐츠 URL
        newRow[17] = "";             // R: 구글맵 리뷰 URL
        newRow[18] = "";             // S: 한 줄 후기
        newRow[19] = "";             // T: (JP)한 줄 후기
        newRow[20] = "";             // U: 환불 처리

        logSheet.appendRow(newRow);
      }
    });

    return ContentService.createTextOutput(JSON.stringify({"result": "success"})).setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    if (debugSheet) debugSheet.appendRow([new Date(), "에러: " + err.toString()]);
    return ContentService.createTextOutput(JSON.stringify({"result": "error", "error": err.toString()})).setMimeType(ContentService.MimeType.JSON);
  }
}
