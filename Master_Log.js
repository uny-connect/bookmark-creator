/*****************************************************************************************
 * ⚡ [엔진 완전 최적화] 아임웹 실제 수신 페이로드 정합성 100% 동기화 시스템
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
    const dataObj = postData.data || postData;

    // 1️⃣ 아임웹 회원가입 이벤트(END_USER_SIGN_UP)
    if (eventType === "END_USER_SIGN_UP") {
      const newMemberUid = String(dataObj.memberUid || "").trim();
      if (newMemberUid && userSheet) {
        // 이미 등록된 회원인지 확인
        const uValues = userSheet.getDataRange().getValues();
        let exists = false;
        for (let u = 1; u < uValues.length; u++) {
          if (String(uValues[u][0]).trim() === newMemberUid || String(uValues[u][13]).trim() === newMemberUid) {
            exists = true; break;
          }
        }
        if (!exists) {
          const newUserRow = new Array(16).fill("");
          newUserRow[0] = newMemberUid;                  // A: 고유키(이메일UID)
          newUserRow[1] = "신규가입(정보수집필요)";        // B: 이름
          newUserRow[13] = newMemberUid;                 // N: 이메일
          userSheet.appendRow(newUserRow);
        }
      }
      return ContentService.createTextOutput(JSON.stringify({"result": "success_signup"})).setMimeType(ContentService.MimeType.JSON);
    }

    // 2️⃣ 아임웹 주문 이벤트 (ORDER_CREATE)
    const orderNo = String(dataObj.orderNo || dataObj.order_no || "번호없음").trim();
    const totalPrice = dataObj.totalPaymentPrice || (dataObj.payment && dataObj.payment.paidPrice) || 0;
    
    // 🎯 실제 아임웹 JSON 키값 완벽 타겟팅
    const memberUid = String(dataObj.memberUid || dataObj.member_code || "").trim();
    const ordererName = String(dataObj.ordererName || dataObj.orderer_name || "").trim();
    const ordererCall = String(dataObj.ordererCall || dataObj.orderer_phone || "").replace(/[^0-9]/g, "");

    // 3️⃣ User_DB 대조 (기존 유저의 영문명/고유키 확인 및 없을 시 연락처 선기록)
    let englishName = ordererName; // 기본값은 주문자 실명
    let finalMemberKey = memberUid;

    if (userSheet) {
      const uData = userSheet.getDataRange().getValues();
      let foundUser = false;
      for (let j = 1; j < uData.length; j++) {
        const dbKey = String(uData[j][0] || '').trim();
        const dbEmail = String(uData[j][13] || '').trim();
        
        if ((memberUid && dbKey === memberUid) || (memberUid && dbEmail === memberUid)) {
          finalMemberKey = dbKey;
          englishName = String(uData[j][1] || ordererName).trim(); // User_DB B열 이름
          // 혹시 D열 연락처가 비어있다면 이번 주문서의 전화번호로 보완
          if (!uData[j][3] && ordererCall) {
            userSheet.getRange(j + 1, 4).setValue(ordererCall);
          }
          foundUser = true;
          break;
        }
      }

      // User_DB에 아예 없는 신규 유저가 바로 주문한 경우: 로그인용 최소 정보 자동 안착
      if (!foundUser && memberUid) {
        const newUserRow = new Array(16).fill("");
        newUserRow[0] = memberUid;       // A: 고유키 (아임웹 memberUid)
        newUserRow[1] = ordererName;     // B: 이름
        newUserRow[2] = ordererName;     // C: 실명
        newUserRow[3] = ordererCall;     // D: 연락처 (로그인에 필수!)
        newUserRow[13] = memberUid;      // N: 이메일
        newUserRow[15] = "주문시 자동안착";
        userSheet.appendRow(newUserRow);
      }
    }

    // 4️⃣ 상품명(sections) 배열 추출
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

    // 5️⃣ 매장 리스트(Restaurant_List) 사전 구축
    const restData = restSheet ? restSheet.getDataRange().getValues() : [];
    const restIdList = [];
    for (let k = 2; k < restData.length; k++) {
      const rId = String(restData[k][0] || '').trim();
      const rNameKo = String(restData[k][1] || '').trim();
      if (rNameKo && rId) {
        restIdList.push({ name: rNameKo, id: rId });
      }
    }

    // 6️⃣ items 순회하며 Master_Log 기록
    items.forEach(function(item) {
      const rawProdName = (item.productInfo && item.productInfo.prodName) || item.product_name || item.name || "식당명 없음";
      
      // 상품명("미이로 miiro")에서 매장명 매칭
      let targetStoreId = "";
      let matchedStoreName = rawProdName;
      for (let m = 0; m < restIdList.length; m++) {
        if (rawProdName.toLowerCase().includes(restIdList[m].name.toLowerCase())) {
          targetStoreId = restIdList[m].id;
          matchedStoreName = restIdList[m].name;
          break;
        }
      }

      const values = logSheet.getDataRange().getValues();
      let rowIndex = -1;
      for (let i = 1; i < values.length; i++) {
        if (String(values[i][0]).replace(/'/g, '').trim() === orderNo) {
          rowIndex = i + 1;
          break;
        }
      }

      let status = "예약대기"; 
      if (rawData.includes("CANCEL") || rawData.includes("REFUND") || eventType === "ORDER_CANCEL") {
        status = "취소완료";
      }

      if (rowIndex > 0) {
        logSheet.getRange(rowIndex, 12).setValue(status); 
        if (status === "취소완료") {
          logSheet.getRange(rowIndex, 21).setValue("취소환불"); 
        }
      } else {
        const newRow = new Array(24).fill(""); 
        newRow[0] = "'" + orderNo;           // A: 주문번호
        newRow[1] = finalMemberKey;          // B: 멤버코드
        newRow[2] = englishName;             // C: 성함
        newRow[3] = "";                      // D: 참여 채널
        newRow[4] = targetStoreId;           // E: 점포 ID
        newRow[5] = matchedStoreName;        // F: 점포명 ("미이로" 자동 매칭)
        newRow[6] = "";                      // G: 예약 캡처 URL
        newRow[7] = "";                      // H: 방문예정일시
        newRow[8] = "";                      // I: 방문인원
        newRow[9] = "";                      // J: 리뷰 마감 기한
        newRow[10] = totalPrice;             // K: 보증금액 (10000)
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
      }
    });

    return ContentService.createTextOutput(JSON.stringify({"result": "success"})).setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    if (debugSheet) debugSheet.appendRow([new Date(), "에러: " + err.toString()]);
    return ContentService.createTextOutput(JSON.stringify({"result": "error", "error": err.toString()})).setMimeType(ContentService.MimeType.JSON);
  }
}
