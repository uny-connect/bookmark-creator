/*****************************************************************************************
 * ⚡ [엔진 완전 최적화] 아임웹 신규 주문 및 취소/환불 웹훅 실시간 시트 적재 시스템
 ****************************************************************************************/
function doPost(e) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const logSheet = ss.getSheetByName("Master_Log");
  const userSheet = ss.getSheetByName("User_DB");
  const restSheet = ss.getSheetByName("Restaurant_List"); // 🎯 매장 ID 추적용 시트 바인딩
  const debugSheet = ss.getSheetByName("Log"); 

  try {
    const rawData = e.postData.contents;
    const postData = JSON.parse(rawData);
    if (debugSheet) debugSheet.appendRow([new Date(), "수신: " + rawData]);

    const orderNo = postData.orderNo || postData.order_no || "번호없음";
    const memberCode = postData.member_code || (postData.member && postData.member.code) || "";
    const totalPrice = (postData.payment && postData.payment.paidPrice) ? postData.payment.paidPrice : 0;
    
    let items = postData.items || [];
    if (items.length === 0 && postData.section && postData.section.sectionItems) {
      items = postData.section.sectionItems;
    }

    // 🎯 [사전 데이터 스캔] Restaurant_List를 메모리에 올려 상품명-매장ID 매핑용 사전 맵 구축
    const restData = restSheet ? restSheet.getDataRange().getValues() : [];
    const restIdMap = {};
    for (let k = 2; k < restData.length; k++) {
      const rId = String(restData[k][0]).trim();
      const rNameKo = String(restData[k][1]).trim();
      if (rNameKo && rId) {
        restIdMap[rNameKo] = rId; // 한국어 매장명을 키값으로 고유 ID 저장
      }
    }

    items.forEach(function(item) {
      const productName = item.product_name || item.name || (item.productInfo && item.productInfo.prodName) || "식당명 없음";
      
      const values = logSheet.getDataRange().getValues();
      let rowIndex = -1;
      for (let i = 1; i < values.length; i++) {
        if (String(values[i][0]).replace(/'/g, '') === String(orderNo) && values[i][5] === productName) {
          rowIndex = i + 1;
          break;
        }
      }

      // 최초 기본 상태를 '예약대기'로 세팅 (크리에이터가 날짜 고르기 전 상태)
      let status = "예약대기"; 
      if (rawData.includes("CANCEL") || rawData.includes("REFUND")) {
        status = "취소완료";
      }

      if (rowIndex > 0) {
        // 1️⃣ 기존 데이터 존재 시: 상태(L열)만 정밀 업데이트 후 마감
        logSheet.getRange(rowIndex, 12).setValue(status); 
        if (status === "취소완료") {
          logSheet.getRange(rowIndex, 21).setValue("취소환불"); // U열 환불 처리 플래그 적재
        }
      } else {
        // 2️⃣ 신규 주문 발생 시: 크리에이터 및 매장 정보 교차 대조 후 구조화 배열 생성
        let englishName = "미승인/정보없음";
        if (memberCode) {
          const userData = userSheet.getDataRange().getValues();
          for (let j = 1; j < userData.length; j++) {
            if (String(userData[j][0]) === String(memberCode)) {
              englishName = userData[j][1];
              break;
            }
          }
        }

        // 🎯 [버그 해결] 사전 구축한 맵에서 상품명과 일치하는 매장 고유 ID 자동 스캔
        const targetStoreId = restIdMap[productName] || "";

        // 실제 크리에이터 마스터 로그 구조 (A~X열) 스펙에 완벽 동기화 (총 24칸)
        const newRow = new Array(24).fill(""); 
        newRow[0] = "'" + orderNo;   // A: 주문번호
        newRow[1] = memberCode;      // B: 멤버코드
        newRow[2] = englishName;     // C: 영문 성함
        newRow[3] = "";              // D: 참여 채널
        newRow[4] = targetStoreId;   // E: 점포 ID (★공백 유실 버그 완벽 수리)
        newRow[5] = productName;     // F: 점포명
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
