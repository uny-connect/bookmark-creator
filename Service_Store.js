/*****************************************************************************************
 * ⚡ [엔진 완전 통합] 점주 일정 변경 리퀘스트 처리 및 크리에이터+운영진 동시 노티 엔진
 ****************************************************************************************/
function submitStoreFeedback(row, orderNo, msg) {
  try {
    const safeRow = parseInt(row, 10);
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('Master_Log');
    
    // 데이터(A~N열까지 총 14칸) 캐싱
    const rowData = sheet.getRange(safeRow, 1, 1, 14).getValues()[0];
    
    const sheetOrderNo = String(rowData[0]).trim().replace(/'/g, '');
    const cleanOrderNo = String(orderNo).trim().replace(/'/g, '');
    
    if (!safeRow || sheetOrderNo !== cleanOrderNo) {
      return { success: false, error: "올바르지 않은 요청 정보입니다." };
    }
    
    // 1️⃣ L열(12번째): 진행 상태를 '일정조율필요'로 변경
    sheet.getRange(safeRow, 12).setValue('일정조율필요');
    
    // 2️⃣ H열~J열 슬롯 영역 자동 리셋 처리
    sheet.getRange(safeRow, 8, 1, 3).clearContent();
    
    // 3️⃣ M열(13번째): 피드백 기록
    const timestamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "MM/dd HH:mm");
    const currentNote = String(rowData[12] || '').trim();
    const newNote = (currentNote ? currentNote + "\n" : "") + `[점주요청 ${timestamp}] ${msg}`;
    sheet.getRange(safeRow, 13).setValue(newNote); 
    
    // 4️⃣ 변수 바인딩
    const storeNameKo = String(rowData[5]).trim();     // F열 기본 한국어 매장명
    const creatorName = String(rowData[2]).trim();   // C열 크리에이터 성함
    const currentMemberCode = String(rowData[1]).trim(); // B열 크리에이터 코드

    // Restaurant_List 시트를 크로스 대조하여 일본어 매장명(storeNameJp)을 실시간 파싱합니다.
    let storeNameJp = "";
    const restSheet = ss.getSheetByName('Restaurant_List');
    if (restSheet) {
      const restData = restSheet.getDataRange().getValues();
      const storeId = String(rowData[4] || '').trim().toUpperCase(); // E열: 매장 ID
      
      for (let k = 2; k < restData.length; k++) {
        if (String(restData[k][0]).trim().toUpperCase() === storeId) {
          storeNameJp = String(restData[k][2] || '').trim(); // C열: 일본어 매장명
          break;
        }
      }
    }
    const finalStoreName = storeNameJp ? storeNameJp : storeNameKo;

    // ✉️ [1번 슈팅] 내부 운영진 담당자 알림 메일 발송 (디자인 커스텀 및 Y열 로그 레이어 연동)
    const adminEmail = "bookmarkjapan.info@gmail.com"; 
    const alertSubject = `[NOTICE] [${finalStoreName}] 店舗からの日程変更リクエスト`;
    
    // 🎯 하단 버튼을 거두어내고, 회색 카드 내부에 '■ ログ位置 (Row 번호)' 항목을 정밀 주입
    const alertBody = `
      <meta charset="UTF-8">
      <div style="background: #f4f5f7; padding: 30px 10px; font-family: 'Helvetica Neue', Arial, 'Hiragino Kaku Gothic ProN', sans-serif;">
        <div style="max-width: 500px; margin: 0 auto; padding: 32px 20px; background: #ffffff; border: 1px solid #eef0f2; border-radius: 20px; box-shadow: 0 4px 12px rgba(0,0,0,0.02);">
          
          <!-- 상단 브랜딩 영역 -->
          <div style="margin-bottom: 24px; text-align: left;">
            <span style="font-size: 11px; font-weight: 800; letter-spacing: 1px; color: #ffffff; background: #7048e8; padding: 4px 10px; border-radius: 6px; display: inline-block;">ADMIN NOTICE</span>
            <h2 style="font-size: 20px; font-weight: 800; color: #1A2B49; margin: 12px 0 0 0;">BOOKMARK CREATORS</h2>
          </div>
          
          <!-- 헤더 타이틀 및 안내 문구 -->
          <div style="border-top: 2px solid #7048e8; padding-top: 24px; margin-bottom: 24px;">
            <p style="font-size: 14.5px; font-weight: 700; color: #7048e8; margin: 0 0 12px 0;">日程変更リクエスト受付</p>
            <p style="font-size: 13.5px; line-height: 1.6; color: #495057; margin: 0;">
              パートナーセンター大画面より、店舗（加盟店）からの日程調整リクエストが送信されました。<br>
              マスターログのステータスは自動的に <span style="font-weight: 700; color: #7048e8;">調整必要</span> に変更されました。クリエイターとの迅速な再スケジュール調整をお願いいたします。
            </p>
          </div>
          
          <!-- 정밀 카드 명세 바디 (로그 위치 로우 항목 추가 탑재 완료) -->
          <div style="background: #f8f9fa; border-radius: 14px; padding: 18px; margin-bottom: 10px;">
            <table style="width: 100%; border-collapse: collapse; font-size: 13.5px;">
              <tr>
                <td style="width: 110px; color: #8b95a1; font-weight: 700; padding: 8px 0; vertical-align: top;">■ 店舗名</td>
                <td style="color: #1A2B49; font-weight: 700; padding: 8px 0;">${finalStoreName}</td>
              </tr>
              <tr>
                <td style="color: #8b95a1; font-weight: 700; padding: 8px 0; vertical-align: top;">■ クレイター</td>
                <td style="color: #495057; font-weight: 700; padding: 8px 0;">${creatorName} <span style="font-size:11.5px; color:#8b95a1; font-weight:500;">(ID: ${currentMemberCode})</span></td>
              </tr>
              <tr>
                <td style="color: #8b95a1; font-weight: 700; padding: 8px 0; vertical-align: top;">■ 注文番号</td>
                <td style="color: #495057; font-weight: 700; padding: 8px 0;">#${cleanOrderNo}</td>
              </tr>
              <tr>
                <td style="color: #8b95a1; font-weight: 700; padding: 8px 0; vertical-align: top;">■ メッセージ</td>
                <td style="color: #7048e8; font-weight: 800; padding: 8px 0; line-height: 1.5; word-break: break-all;">${msg}</td>
              </tr>
              <!-- 🎯 운영 편의성 고도화: 해당 주문이 마스터로그 몇 번째 행에 있는지 직관적으로 명시 -->
              <tr>
                <td style="color: #8b95a1; font-weight: 700; padding: 8px 0; vertical-align: top; border-top: 1px dashed #dee2e6; padding-top: 12px;">■ ログ位置</td>
                <td style="color: #d63384; font-weight: 800; padding: 8px 0; border-top: 1px dashed #dee2e6; padding-top: 12px;">Master_Log Sheet ${safeRow}番目行 (Row ${safeRow})</td>
              </tr>
            </table>
          </div>
          
        </div>
      </div>`;

    if (adminEmail && adminEmail.includes('@')) {
      try {
        GmailApp.sendEmail(adminEmail, alertSubject, "", { htmlBody: alertBody, name: "BOOKMARK CREATORS" });
        sheet.getRange(safeRow, 25).setValue("운영진알림 발송완료"); 
      } catch (adminMailErr) {
        console.error("❌ 운영진 메일 알림 셧다운 낚아챔: " + adminMailErr.toString());
        sheet.getRange(safeRow, 25).setValue("❌ 운영진알림 실패: " + adminMailErr.toString());
      }
    } else {
      sheet.getRange(safeRow, 25).setValue("❌ 실패: 운영진 이메일 주소 포맷 오류");
    }

    // ✉️ [2번 슈팅] 크리에이터향 알림 이메일 발송
    if (currentMemberCode) {
      const userSheet = ss.getSheetByName('User_DB');
      if (userSheet) {
        const userData = userSheet.getDataRange().getValues();
        let creatorEmail = "";
        
        for (let j = 1; j < userData.length; j++) {
          if (String(userData[j][0]).trim() === currentMemberCode) {
            creatorEmail = String(userData[j][12] || '').trim();
            break;
          }
        }

        if (creatorEmail && creatorEmail.includes("@")) {
          const timeZone = Session.getScriptTimeZone();
          const rawVisitDate = rowData[7]; 
          const pCount = rowData[8] ? String(rowData[8]).replace(/[^0-9]/g, '') : '1'; 
          const visitDateStr = (rawVisitDate instanceof Date) ? Utilities.formatDate(rawVisitDate, timeZone, 'yyyy-MM-dd HH:mm') : String(rawVisitDate || '-');

          const creatorSubject = "[NOTICE] [BOOKMARK CREATORS] 방문 일정 변경 조율 안내";
          const creatorHtmlBody = `
            <meta charset="UTF-8">
            <div style="background: #f4f5f7; padding: 30px 10px; font-family: 'Apple SD Gothic Neo', 'Malgun Gothic', sans-serif;">
              <div style="max-width: 500px; margin: 0 auto; padding: 32px 20px; background: #ffffff; border: 1px solid #eef0f2; border-radius: 20px; box-shadow: 0 4px 12px rgba(0,0,0,0.02);">
                <div style="margin-bottom: 24px; text-align: left;">
                  <span style="font-size: 11px; font-weight: 800; letter-spacing: 1px; color: #ffffff; background: #e03131; padding: 4px 10px; border-radius: 6px; display: inline-block;">NOTICE</span>
                  <h2 style="font-size: 20px; font-weight: 800; color: #1A2B49; margin: 12px 0 0 0;">BOOKMARK CREATORS</h2>
                </div>
                <div style="border-top: 2px solid #e03131; padding-top: 24px; margin-bottom: 24px;">
                  <p style="font-size: 14.5px; font-weight: 700; color: #e03131; margin: 0 0 12px 0;">🟣 방문 일정 변경 요청 알림</p>
                  <p style="font-size: 13.5px; line-height: 1.6; color: #495057; margin: 0;">
                    안녕하세요 크리에이터님, 신청하신 매장의 예약이 현지 사정으로 인해 <span style="font-weight: 700; color: #e03131;">일정 조율이 필요한 상태</span>로 변경되었습니다.<br><br>
                    아래의 신청 내역을 바탕으로 운영 담당자가 신속하게 연락드려 일정 재조율을 도와드리겠습니다.
                  </p>
                </div>
                <div style="background: #f8f9fa; border-radius: 14px; padding: 18px; margin-bottom: 28px;">
                  <table style="width: 100%; border-collapse: collapse; font-size: 13px;">
                    <tr>
                      <td style="width: 85px; color: #8b95a1; font-weight: 700; padding: 6px 0;">방문 매장</td>
                      <td style="color: #1A2B49; font-weight: 700; padding: 6px 0;">${storeNameKo}</td>
                    </tr>
                    <tr>
                      <td style="color: #8b95a1; font-weight: 700; padding: 6px 0;">기존 일시</td>
                      <td style="color: #e03131; font-weight: 700; padding: 6px 0;">${visitDateStr}</td>
                    </tr>
                    <tr>
                      <td style="color: #8b95a1; font-weight: 700; padding: 6px 0;">방문 인원</td>
                      <td style="color: #495057; font-weight: 700; padding: 6px 0;">${pCount}명</td>
                    </tr>
                  </table>
                </div>
                <div style="text-align: center;">
                  <a href="http://pf.kakao.com/_vFSxfX/chat" target="_blank" style="display: block; background: #e03131; color: #ffffff; font-size: 14px; font-weight: 700; text-decoration: none; padding: 14px; border-radius: 12px; box-shadow: 0 4px 12px rgba(224,49,49,0.15);">담당자에게 카톡 문의하기</a>
                </div>
              </div>
            </div>`;

          try {
            GmailApp.sendEmail(creatorEmail, creatorSubject, "", { htmlBody: creatorHtmlBody, name: "BOOKMARK CREATORS" });
            const prevLog = String(sheet.getRange(safeRow, 25).getValue() || '');
            sheet.getRange(safeRow, 25).setValue(prevLog.includes("완료") ? "조율 노티 전체메일 발송완료" : "크리에이터 노티 메일 발송완료");
          } catch (creatorMailErr) {
            console.error("❌ 크리에이터 변경 안내 메일 실패: " + creatorMailErr.toString());
            const prevLog = String(sheet.getRange(safeRow, 25).getValue() || '');
            sheet.getRange(safeRow, 25).setValue(prevLog + " | ❌ 크리에이터메일 실패: " + creatorMailErr.toString());
          }
        }
      }
    }
    return { success: true };
  } catch (e) { return { success: false, error: e.toString() }; }
}


/*****************************************************************************************
 * 🎯 [파트너센터 로그인 & 대시보드 연동 코어 엔진]
 ****************************************************************************************/

/** [1] 매장 고유 PIN 번호로 로그인 검증 */
function loginStoreByPin(pin) {
  try {
    const pinClean = String(pin).trim().toUpperCase();
    if (!pinClean) return { success: false, error: "コードを入力してください。" };

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const restSheet = ss.getSheetByName('Restaurant_List');
    if (!restSheet) return { success: false, error: "시스템 오류: Restaurant_List 시트가 없습니다." };

    const restData = restSheet.getDataRange().getValues();
    
    for (let i = 2; i < restData.length; i++) {
      const sheetPin = String(restData[i][8] || '').trim().toUpperCase();
      if (sheetPin === pinClean) {
        const storeId = String(restData[i][0]).trim(); 
        return { success: true, storeId: storeId };
      }
    }
    
    return { success: false, error: "無効なパートナーコードです。再度コードを確認してください。" };
  } catch (e) {
    return { success: false, error: "서버 오류: " + e.toString() };
  }
}

/** [2] 특정 매장의 대시보드 스탯 및 할당 크리에이터 미션 리스트 조회 */
function getStoreData(storeId) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const data = _getSheetsData(['Master_Log', 'Restaurant_List', 'User_DB']);
    
    let storeName = storeId;
    let storeNameJp = "";
    
    const restList = data.Restaurant_List;
    for (let k = 2; k < restList.length; k++) {
      if (String(restList[k][0]).trim().toUpperCase() === String(storeId).toUpperCase()) {
        storeName = String(restList[k][1] || '').trim(); 
        storeNameJp = String(restList[k][2] || '').trim(); 
        break;
      }
    }
    const finalStoreName = storeNameJp ? storeNameJp : storeName;

    const userTierMap = {};
    const userDB = data.User_DB;
    for (let u = 1; u < userDB.length; u++) {
      if (userDB[u][0]) {
        let emoji = '🟡';
        try { emoji = parseTierEmoji(userDB[u][8]); } catch(ex) {}
        userTierMap[String(userDB[u][0]).trim()] = emoji;
      }
    }

    let stats = { total: 0, pending: 0, verified: 0, submitted: 0 };
    let storeMissions = [];
    const timeZone = Session.getScriptTimeZone();
    const masterLog = data.Master_Log;

    const webAppUrl = ScriptApp.getService().getUrl();

    for (let i = masterLog.length - 1; i >= 2; i--) {
      const mRow = masterLog[i];
      const mStoreId = String(mRow[4] || '').trim().toUpperCase(); 
      const status = String(mRow[11] || '').trim(); 

      if (!mRow[0] || status.includes('취소') || mStoreId !== String(storeId).toUpperCase()) continue;

      stats.total++;
      if (status === '방문전' || status === '예약확인중' || status === '일정조율필요' || status === '예약대기') stats.pending++;
      else if (status === '방문완료') stats.verified++;
      else if (status === '제출완료') stats.submitted++;

      const formatDate = (val, format) => (val instanceof Date) ? Utilities.formatDate(val, timeZone, format) : String(val || '-');
      const memberCode = String(mRow[1]).trim();
      const orderNoRaw = String(mRow[0]).replace(/'/g, ''); 

      storeMissions.push({
        row: i + 1,
        orderNo: orderNoRaw,
        memberCode: memberCode,
        member: String(mRow[2] || '').trim(), 
        tierEmoji: userTierMap[memberCode] || '🟡',
        visitDate: formatDate(mRow[7], 'yyyy-MM-dd HH:mm'), 
        peopleCount: mRow[8] ? String(mRow[8]).replace(/[^0-9]/g, '') : '1', 
        status: status,
        submitDate: formatDate(mRow[15], 'yyyy-MM-dd'), 
        reviewUrl: String(mRow[16] || ''), 
        googleMapUrl: String(mRow[17] || ''), 
        shortReview: String(mRow[19] || mRow[18] || '')
      });
    }

    storeMissions.forEach(m => {
      m.confirmUrl = `${webAppUrl}?mode=store_confirm&row=${m.row}&o=${encodeURIComponent(m.orderNo)}`;
      m.feedbackUrl = `${webAppUrl}?mode=feedback&row=${m.row}&o=${encodeURIComponent(m.orderNo)}`;
    });

    return { storeName: finalStoreName, stats: stats, missions: storeMissions };
  } catch (e) {
    throw new Error("매장 데이터 로드 실패: " + e.toString());
  }
}

/** [3] 매장 예약 환경설정 파일 불러오기 */
function getStoreConfig(storeId) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('Restaurant_List');
    const data = sheet.getDataRange().getValues();

    for (let i = 2; i < data.length; i++) {
      if (String(data[i][0]).trim().toUpperCase() === String(storeId).toUpperCase()) {
        return {
          success: true,
          row: i + 1,
          storeId: String(data[i][0]).trim(),
          email: String(data[i][10] || '').trim(),      
          storeType: String(data[i][7] || '').trim(),   
          maxPerSlot: data[i][11] ? parseInt(data[i][11], 10) : 1, 
          dailyCap: data[i][12] ? parseInt(data[i][12], 10) : 999,  
          blackouts: String(data[i][13] || '').trim(),   
          bizHours: String(data[i][14] || '').trim(),    
          intervalMin: data[i][15] ? String(data[i][15]).trim() : "30", 
          blockedTimes: String(data[i][16] || '').trim() 
        };
      }
    }
    return { success: false, error: "店舗設定情報が見つかりません。" };
  } catch (e) {
    return { success: false, error: e.toString() };
  }
}

/** [4] 매장 예약 환경설정 파일 수정 저장하기 */
function updateStoreConfig(row, email, storeType, maxPerSlot, dailyCap, blackouts, bizHours, intervalMin, blockedTimes) {
  try {
    const safeRow = parseInt(row, 10);
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Restaurant_List');
    
    sheet.getRange(safeRow, 11).setValue(email);          
    sheet.getRange(safeRow, 8).setValue(storeType);        
    sheet.getRange(safeRow, 12).setValue(parseInt(maxPerSlot, 10) || 1);  
    sheet.getRange(safeRow, 13).setValue(parseInt(dailyCap, 10) || 999);   
    sheet.getRange(safeRow, 14).setValue(blackouts);     
    sheet.getRange(safeRow, 15).setValue(bizHours);      
    sheet.getRange(safeRow, 16).setValue(parseInt(intervalMin, 10) || 30); 
    sheet.getRange(safeRow, 17).setValue(blockedTimes);  

    return { success: true };
  } catch (e) {
    return { success: false, error: e.toString() };
  }
}

/** [신설] 점주가 대시보드에서 즉시 예약을 확정(승인)하는 함수 */
function storeDirectConfirm(row, orderNo) {
  try {
    const safeRow = parseInt(row, 10);
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Master_Log');
    const sheetOrderNo = String(sheet.getRange(safeRow, 1).getValue()).trim().replace(/'/g, '');
    
    if (sheetOrderNo !== String(orderNo).trim().replace(/'/g, '')) {
      return { success: false, error: "주문번호가 일치하지 않습니다." };
    }
    
    sheet.getRange(safeRow, 12).setValue('방문전');
    if (String(sheet.getRange(safeRow, 7).getValue() || '').trim() === '') {
      sheet.getRange(safeRow, 7).setValue('점주_대시보드_직접확정');
    }

    try {
      const currentRestaurantName = String(sheet.getRange(safeRow, 6).getValue() || '').trim(); 
      const rawMemberCode = String(sheet.getRange(safeRow, 2).getValue() || '').trim();
      const currentMemberCode = rawMemberCode.replace(/['"\s]/g, '').toLowerCase();
      
      const timeZone = Session.getScriptTimeZone();
      const rawVisitDate = sheet.getRange(safeRow, 8).getValue();
      const pCount = sheet.getRange(safeRow, 9).getValue() ? String(sheet.getRange(safeRow, 9).getValue()).replace(/[^0-9]/g, '') : '1'; 
      const visitDateStr = (rawVisitDate instanceof Date) ? Utilities.formatDate(rawVisitDate, timeZone, 'yyyy-MM-dd HH:mm') : String(rawVisitDate || '-');

      if (currentMemberCode) {
        const userSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('User_DB');
        const userData = userSheet.getDataRange().getValues();
        let creatorEmail = "";
        
        for (let j = 1; j < userData.length; j++) {
          const dbMemberCode = String(userData[j][0] || '').trim().replace(/['"\s]/g, '').toLowerCase();
          if (dbMemberCode === currentMemberCode) {
            creatorEmail = String(userData[j][12] || '').trim();
            break;
          }
        }

        if (creatorEmail && creatorEmail.includes("@")) {
          const subject = "[BOOKMARK CREATORS] 予約確定のご案内"; // 🎯 메일 제목 인코딩 리스크 방어 우회식 처리 완료
          const htmlBody = `
            <meta charset="UTF-8">
            <div style="max-width: 500px; margin: 0 auto; padding: 32px 20px; background: #ffffff; font-family: 'Apple SD Gothic Neo', 'Malgun Gothic', sans-serif; border: 1px solid #eef0f2; border-radius: 20px; box-shadow: 0 4px 12px rgba(0,0,0,0.02);">
              <div style="margin-bottom: 24px; text-align: left;">
                <span style="font-size: 11px; font-weight: 800; letter-spacing: 1px; color: #ffffff; background: #1A2B49; padding: 4px 10px; border-radius: 6px; display: inline-block;">NOTICE</span>
                <h2 style="font-size: 20px; font-weight: 800; color: #1A2B49; margin: 12px 0 0 0;">BOOKMARK CREATORS</h2>
              </div>
              <div style="border-top: 2px solid #1A2B49; padding-top: 24px; margin-bottom: 24px;">
                <p style="font-size: 14.5px; font-weight: 700; color: #2D6A4F; margin: 0 0 12px 0;">✅ 예약 확정 알림</p>
                <p style="font-size: 13.5px; line-height: 1.6; color: #495057; margin: 0;">
                  매장에서 방문 예약이 확정되었습니다.<br>
                  날짜와 시간을 다시 한번 확인 후 늦지 않게 방문해주세요! <br>
                  <span style="font-weight: 700; color: #dc3545;">혹시라도 늦는다면 미리 말씀 부탁드립니다.</span>
                </p>
              </div>
              <div style="background: #f8f9fa; border-radius: 14px; padding: 18px; margin-bottom: 28px;">
                <table style="width: 100%; border-collapse: collapse; font-size: 13px;">
                  <tr>
                    <td style="width: 85px; color: #8b95a1; font-weight: 700; padding: 6px 0;">방문 매장</td>
                    <td style="color: #1A2B49; font-weight: 700; padding: 6px 0;">${currentRestaurantName}</td>
                  </tr>
                  <tr>
                    <td style="color: #8b95a1; font-weight: 700; padding: 6px 0;">예약 일시</td>
                    <td style="color: #1a73e8; font-weight: 700; padding: 6px 0;">${visitDateStr}</td>
                  </tr>
                  <tr>
                    <td style="color: #8b95a1; font-weight: 700; padding: 6px 0;">방문 인원</td>
                    <td style="color: #495057; font-weight: 700; padding: 6px 0;">${pCount}명</td>
                  </tr>
                </table>
              </div>
              <div style="text-align: center;">
                <a href="http://pf.kakao.com/_vFSxfX/chat" target="_blank" style="display: block; background: #1A2B49; color: #ffffff; font-size: 14px; font-weight: 700; text-decoration: none; padding: 14px; border-radius: 12px; box-shadow: 0 4px 12px rgba(26,43,73,0.15);">카카오톡 채팅하기</a>
              </div>
            </div>`;
          
          try {
            GmailApp.sendEmail(creatorEmail, subject, "", { htmlBody: htmlBody, name: "BOOKMARK CREATORS" });
            sheet.getRange(safeRow, 25).setValue("크리에이터확정메일 발송완료"); 
          } catch (directMailErr) {
            console.error("❌ 점주 직접 확정 메일 발송 실패: " + directMailErr.toString());
            sheet.getRange(safeRow, 25).setValue("❌ 크리에이터확정메일 실패: " + directMailErr.toString());
          }
        }
      }
    } catch (mailErr) {}
    return { success: true };
  } catch (e) { return { success: false, error: e.toString() }; }
}

/** [신설] 점주가 대시보드에서 일정 변경 팝업창을 통해 리퀘스트를 보낼 때 처리 함수 */
function storeDirectFeedback(row, orderNo, msg) {
  return submitStoreFeedback(row, orderNo, msg);
}
