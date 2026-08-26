/********************************************************************
 * [1] 웹 앱 진입점 컨트롤러 (doGet 엔진)
 ********************************************************************/
function doGet(e) {
  const mode = e?.parameter?.mode;
  
  // 1️⃣ 어드민 대시보드 진입 분기
  if (mode === 'admin') {
    return HtmlService.createTemplateFromFile('Admin').evaluate()
      .setTitle('BOOKMARK CREATORS | Admin')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
      .addMetaTag('viewport', 'width=device-width, initial-scale=1, maximum-scale=1');
  }
  
  // 2️⃣ 점주 파트너 레포트 대시보드 진입 분기
  if (mode === 'store' || e?.parameter?.id) {
    const template = HtmlService.createTemplateFromFile('Store'); 
    template.storeId = e?.parameter?.id || ''; 
    return template.evaluate()
      .setTitle('BOOKMARK CREATORS | パートナーレポート')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
      .addMetaTag('viewport', 'width=device-width, initial-scale=1, maximum-scale=1');
  }
  
  // 3️⃣ 메일 직통 링크 예약 승인 엔진 분기
  if (mode === 'store_confirm') {
    let isAlreadyRequested = false; 
    
    try {
      const safeRow = parseInt(e.parameter.row, 10);
      const orderNo = e.parameter.o;
      const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Master_Log');
      
      const sheetOrderNo = String(sheet.getRange(safeRow, 1).getValue()).trim().replace(/'/g, '');
      const currentStatus = String(sheet.getRange(safeRow, 12).getValue()).trim();

      if (currentStatus === '일정조율필요' || currentStatus.includes('취소') || currentStatus.includes('노쇼')) {
        isAlreadyRequested = true; 
      } 
      else if (safeRow && sheetOrderNo === String(orderNo).trim().replace(/'/g, '')) {
        // 시트 진행 상태를 '방문전'으로 변경 및 마킹 보완
        sheet.getRange(safeRow, 12).setValue('방문전'); 
        if (String(sheet.getRange(safeRow, 7).getValue() || '').trim() === '') {
          sheet.getRange(safeRow, 7).setValue('점주_직접_링크확정');
        }

        // 크리에이터향 확정 알림 메일 슈팅 (NOTICE 뱃지 가동)
        try {
          const currentRestaurantName = String(sheet.getRange(safeRow, 6).getValue() || '').trim(); 
          const rawMemberCode = String(sheet.getRange(safeRow, 2).getValue() || '').trim();
          const currentMemberCode = rawMemberCode.replace(/['"\s]/g, '').toLowerCase();
          
          const timeZone = Session.getScriptTimeZone();
          const rawVisitDate = sheet.getRange(safeRow, 8).getValue();
          // 🎯 버그 보정: 이스케이프 없이 안전하게 변수로 분리 선언 후 치환
          const rawPeopleStr = String(sheet.getRange(safeRow, 9).getValue() || '1');
          const pCount = rawPeopleStr.replace(/[^0-9]/g, '');
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
              const subject = "🗓️ [BOOKMARK CREATORS] 방문 예약 확정 안내";
              const htmlBody = `
                <div style="max-width: 500px; margin: 0 auto; padding: 32px 20px; background: #ffffff; font-family: 'Apple SD Gothic Neo', 'Malgun Gothic', sans-serif; border: 1px solid #eef0f2; border-radius: 20px; box-shadow: 0 4px 12px rgba(0,0,0,0.02);">
                  <div style="margin-bottom: 24px; text-align: left;">
                    <span style="font-size: 11px; font-weight: 800; letter-spacing: 1px; color: #ffffff; background: #1A2B49; padding: 4px 10px; border-radius: 6px; display: inline-block;">NOTICE</span>
                    <h2 style="font-size: 20px; font-weight: 800; color: #1A2B49; margin: 12px 0 0 0;">BOOKMARK CREATORS</h2>
                  </div>
                  <div style="border-top: 2px solid #1A2B49; padding-top: 24px; margin-bottom: 24px;">
                    <p style="font-size: 14.5px; font-weight: 700; color: #2D6A4F; margin: 0 0 12px 0;">✅ 예약 확정 안내</p>
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
              
              GmailApp.sendEmail(creatorEmail, subject, "", { htmlBody: htmlBody, name: "BOOKMARK CREATORS" });
            }
          }
        } catch (mailErr) {
          console.error("❌ doGet 메일 엔진 연산 실패: " + mailErr.toString());
        }
      }
    } catch(err) {
      console.error("❌ store_confirm 코어 에러: " + err.toString());
    }
    
    // 🚨 [1] 중복 확정 불가 실패 화면 (하단 회색 박스 스펙 통일)
    if (isAlreadyRequested) {
      return HtmlService.createHtmlOutput(`
        <!DOCTYPE html>
        <html>
        <head>
          <base target="_top">
          <meta charset="UTF-8">
          <meta name="viewport" content="width=device-width, initial-scale=1, maximum-scale=1, user-scalable=no">
          <link rel="stylesheet" href="https://cdn.jsdelivr.net/gh/orioncactus/pretendard@v1.3.9/dist/web/static/pretendard.min.css">
          <link href="https://fonts.googleapis.com/css2?family=Noto+Sans+JP:wght@400;500;700&display=swap" rel="stylesheet">
          <style>
            body { background: #f4f5f7; font-family: 'Pretendard', 'Noto Sans JP', sans-serif; margin: 0; padding: 0 !important; color: #333; }
            .container { max-width: 380px; padding: 40px 20px; margin: auto; text-align: center; }
            .card { border-radius: 18px; border: none; box-shadow: 0 8px 24px rgba(0,0,0,0.02); background: #ffffff; padding: 24px 20px; margin-top: 5px; }
            .badge-partner { font-size: 10px; letter-spacing: 0.8px; padding: 5px 10px; border-radius: 6px; background-color: #1A2B49 !important; color: #fff; display: inline-block; font-weight: 700; }
            .brand-title { color: #1A2B49; font-weight: 700; margin-top: 10px; margin-bottom: 20px; font-size: 22px; letter-spacing: -0.5px; }
            .fail-status { color: #dc3545; font-weight: 700; font-size: 15px; margin-top: 0; margin-bottom: 16px; display: block; text-align: center; }
            .desc-text { color: #495057; font-size: 12.5px; line-height: 1.6; margin: 0; font-weight: 500; text-align: center; word-break: break-all; overflow-wrap: break-word; padding: 0 4px; }
            .notice-box { margin-top: 24px; padding: 14px; background-color: #f8f9fa; border-radius: 12px; border: 1px dashed #cfd4da; color: #495057; font-size: 13px; font-weight: 700; text-align: center; }
          </style>
        </head>
        <body>
          <div class="container">
            <div style="margin-bottom: 10px; margin-top: 10px;"><span class="badge-partner">PARTNER CENTER</span></div>
            <h2 class="brand-title">BOOKMARK CREATORS</h2>
            <div class="card">
              <span class="fail-status">❌ 確定不可</span>
              <p class="desc-text">
                この件は店舗から<b>'日時変更リクエスト'</b>が送信されている状況です。<br><br>
                店舗のご要望に合わせた新たな日時をご提案いたしますので<br>
                今しばらくお待ちください。🙇‍♂️
              </p>
              <div class="notice-box">このままページを閉じてください。</div>
            </div>
          </div>
        </body>
        </html>
      `).setTitle('BOOKMARK CREATORS | 確定不可').addMetaTag('viewport', 'width=device-width, initial-scale=1');
    }

    // 🚨 [2] 정상 확정 완료 시 화면 (하단 회색 박스 스펙 통일)
    return HtmlService.createHtmlOutput(`
        <!DOCTYPE html>
        <html>
        <head>
          <base target="_top">
          <meta charset="UTF-8">
          <meta name="viewport" content="width=device-width, initial-scale=1, maximum-scale=1, user-scalable=no">
          <link rel="stylesheet" href="https://cdn.jsdelivr.net/gh/orioncactus/pretendard@v1.3.9/dist/web/static/pretendard.min.css">
          <link href="https://fonts.googleapis.com/css2?family=Noto+Sans+JP:wght@400;500;700&display=swap" rel="stylesheet">
          <style>
            body { background: #f4f5f7; font-family: 'Pretendard', 'Noto Sans JP', sans-serif; margin: 0; padding: 0 !important; color: #333; }
            .container { max-width: 380px; padding: 40px 20px; margin: auto; text-align: center; }
            .card { border-radius: 18px; border: none; box-shadow: 0 8px 24px rgba(0,0,0,0.02); background: #ffffff; padding: 24px 20px; margin-top: 5px; }
            .badge-partner { font-size: 10px; letter-spacing: 0.8px; padding: 5px 10px; border-radius: 6px; background-color: #1A2B49 !important; color: #fff; display: inline-block; font-weight: 700; }
            .brand-title { color: #1A2B49; font-weight: 700; margin-top: 10px; margin-bottom: 20px; font-size: 22px; letter-spacing: -0.5px; }
            .success-status { color: #2D6A4F; font-weight: 700; font-size: 15px; margin-top: 0; margin-bottom: 16px; display: block; text-align: center; }
            .desc-text { color: #495057; font-size: 12.5px; line-height: 1.6; margin: 0; font-weight: 500; text-align: center; word-break: break-all; overflow-wrap: break-word; padding: 0 4px; }
            .notice-box { margin-top: 24px; padding: 14px; background-color: #f8f9fa; border-radius: 12px; border: 1px dashed #cfd4da; color: #495057; font-size: 13px; font-weight: 700; text-align: center; }
          </style>
        </head>
        <body>
          <div class="container">
            <div style="margin-bottom: 10px; margin-top: 10px;"><span class="badge-partner">PARTNER CENTER</span></div>
            <h2 class="brand-title">BOOKMARK CREATORS</h2>
            <div class="card">
              <span class="success-status">✅ 予約確定完了</span>
              <p class="desc-text">
                来店予約が確定しました。<br>
                クリエイターの訪問日時にあわせて、<br>
                ご準備をお願いいたします。🙏
              </p>
              <div class="notice-box">このままページを閉じてください。</div>
            </div>
          </div>
        </body>
        </html>
      `).setTitle('BOOKMARK CREATORS | 確定完了').addMetaTag('viewport', 'width=device-width, initial-scale=1');
  }

  // 4️⃣ 점주 피드백 입력 분기
  if (mode === 'feedback') {
    const template = HtmlService.createTemplateFromFile('Feedback');
    template.row = e?.parameter?.row || '';
    template.orderNo = e?.parameter?.o || '';
    return template.evaluate().setTitle('BOOKMARK CREATORS | 리퀘스트').setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
      .addMetaTag('viewport', 'width=device-width, initial-scale=1, maximum-scale=1, user-scalable=no');
  }

  // 5️⃣ 기본 디폴트 화면 (크리에이터용 PIN 화면)
  const template = HtmlService.createTemplateFromFile('PIN화면2');
  template.orderNo = (e?.parameter?.o) || ''; 
  template.phoneLast4 = (e?.parameter?.p) || '';
  return template.evaluate().setTitle('BOOKMARK CREATORS | Creator').setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, maximum-scale=1, user-scalable=no');
}

/********************************************************************
 * [2] 시트 어시스턴트 유틸리티 매뉴얼 (onOpen / onEdit)
 ********************************************************************/
function onOpen() {
  SpreadsheetApp.getUi().createMenu('⚙️ BOOKMARK CREATORS 관리')
    .addItem('✨ 빈칸 자동 채우기 (이름/매장/마감일/보증금)', 'fillMissingData')
    .addItem('🔑 매장별 고유 PIN 6자리 생성', 'generateStorePins') 
    .addItem('🚨 자동 노쇼 일괄 처리 (과거 날짜)', 'checkAndMarkNoShow')
    .addItem('👥 아임웹 신규 회원 동기화', 'syncImwebUsers')
    .addItem('📧 매장 메일 수신 테스트 발송', 'sendTestEmailToStore') // 🎯 신규 추가
    .addToUi();
}

function onEdit(e) {
  if (!e || !e.range) return; 
  const sheet = e.range.getSheet(); 
  if (sheet.getName() !== 'Master_Log') return;
  
  const row = e.range.getRow(); 
  const col = e.range.getColumn();
  const val = String(e.range.getValue()).trim();

// 🎯 [완벽 보정] E열(5번째 열) 점포 ID 입력 시 F열(6번째 열) 점포명 실시간 자동 기입
  if (col === 5 && row >= 3) { 
    const storeId = val.toUpperCase();
    const storeNameCell = sheet.getRange(row, 6); // F열 (점포명)
    const statusCell = sheet.getRange(row, 12);   // L열 (진행 상태)
    
    if (storeId === "") {
      storeNameCell.clearContent();
    } else {
      const ss = SpreadsheetApp.getActiveSpreadsheet();
      const restSheet = ss.getSheetByName('Restaurant_List');
      if (restSheet) {
        const restData = restSheet.getDataRange().getValues();
        let matchedName = "식당명 없음"; // 매칭 실패 시 기본 텍스트
        
        for (let k = 2; k < restData.length; k++) {
          const sheetStoreId = String(restData[k][0]).trim().toUpperCase();
          if (sheetStoreId === storeId) {
            matchedName = String(restData[k][1]).trim(); // B열: 한국어 매장명
            break;
          }
        }
        storeNameCell.setValue(matchedName); // F열에 최종 기입
      }
      
      // 🎯 [신규] L열이 비어있으면 '예약대기'로 자동 초기화 (수기 입력 시 상태 누락 방지)
      if (String(statusCell.getValue()).trim() === "") {
        statusCell.setValue("예약대기");
      }
    }
  }

  // 🗓️ H열(8번째 열: 방문예정일시) 편집 시 J열(마감일), K열(보증금) 자동 연산
  if (col === 8 && row >= 3) { 
    const deadlineCell = sheet.getRange(row, 10); // J열 (리뷰 마감 기한)
    const depositCell = sheet.getRange(row, 11);  // K열 (보증금액)
    const cellValue = e.range.getValue();
    
    if (cellValue instanceof Date) { 
      const d = new Date(cellValue); 
      d.setDate(d.getDate() + 10); 
      deadlineCell.setValue(d); 
      if (!depositCell.getValue()) depositCell.setValue(50000); 
    } else if (!cellValue) { 
      deadlineCell.clearContent(); 
    } 
  }

  // 🔄 L열(12번째 열: 진행 상태) 변경 시 슬롯 리셋
  if (col === 12 && row >= 3) {
    if (val === '예약대기' || val === '일정조율필요') {
      sheet.getRange(row, 8, 1, 3).clearContent();
    }
  }
}

/********************************************************************
 * [3] 어드민 코어 백엔드 API 연산 엔진 (adminUpdateMission 정밀 수리 완료)
 ********************************************************************/
function getAdminData() {
  try {
    const data = _getSheetsData(['Master_Log', 'Restaurant_List', 'User_DB']);
    let stats = { total: 0, pending: 0, verified: 0, submitted: 0, waitingReview: 0 };
    let missions = [], restList = [], restNameMap = {}, userTierMap = {};

    const userDB = data.User_DB;
    for (let u = 2; u < userDB.length; u++) {
      if(userDB[u][0]) {
        try {
          userTierMap[String(userDB[u][0]).trim()] = parseTierEmoji(userDB[u][8]);
        } catch(tierErr) {
          userTierMap[String(userDB[u][0]).trim()] = '🟡'; 
        }
      }
    }

    const restListRaw = data.Restaurant_List;
    for (let k = 2; k < restListRaw.length; k++) {
      if(!restListRaw[k][0]) continue; 
      let rId = String(restListRaw[k][0]).trim(); 
      restNameMap[rId] = String(restListRaw[k][2] || '').trim();
      restList.push({ row: k + 1, id: rId, name: String(restListRaw[k][1] || '').trim(), nameJp: restNameMap[rId], pin: restListRaw[k][8] });
    }

    const timeZone = Session.getScriptTimeZone();
    const formatDate = (val, format) => (val instanceof Date) ? Utilities.formatDate(val, timeZone, format) : String(val || '-');

    const masterLog = data.Master_Log;
    for (let i = masterLog.length - 1; i >= 2; i--) {
      const mRow = masterLog[i];
      const status = String(mRow[11] || '').trim();
      if (!mRow[0] || status.includes('취소')) continue; 
      stats.total++;
      
      if (status === '방문전' || status === '예약확인중' || status === '일정조율필요') stats.pending++; 
      else if (status === '방문완료') stats.verified++; 
      else if (status === '제출완료') { stats.submitted++; stats.waitingReview++; }

      let depRaw = String(mRow[10] || '').replace(/[^0-9.-]/g, ''); 
      let mRestId = String(mRow[4]).trim(), mMemberCode = String(mRow[1]).trim();
      
      missions.push({
        row: i + 1, orderNo: mRow[0], memberCode: mMemberCode, member: mRow[2], 
        tierEmoji: userTierMap[mMemberCode] || '🟡', restId: mRestId, restaurant: String(mRow[5] || '').trim(), 
        restaurantJp: restNameMap[mRestId] || String(mRow[5] || '').trim(), status: status,
        feedback: String(mRow[12] || ''),
        visitDate: formatDate(mRow[7], 'yyyy-MM-dd HH:mm'),
        deadline: formatDate(mRow[9], 'yyyy-MM-dd'), deposit: depRaw ? Number(depRaw) : 0, 
        receiptUrl: String(mRow[14] || ''),
        submitDate: formatDate(mRow[15], 'yyyy-MM-dd'),
        reviewUrl: String(mRow[16] || ''),
        googleMapUrl: String(mRow[17] || ''),
        shortReview: String(mRow[18] || ''),
        refundStatus: String(mRow[20] || '').trim()
      });
    }

    return { stats, missions, restList, challengeSettings: getChallengeSettings() };
  } catch (e) { throw new Error(e.toString()); }
}

function adminUpdateMission(row, newStatus, newLink, newRefundStatus, newVisitDate, newGoogleLink, newShortReview) {
  try {
    const safeRow = parseInt(row, 10); 
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Master_Log');
    
    const currentRestaurantName = String(sheet.getRange(safeRow, 6).getValue() || '').trim(); 
    const currentMemberCode = String(sheet.getRange(safeRow, 2).getValue() || '').trim(); 
    
    sheet.getRange(safeRow, 12).setValue(newStatus);
    
    if (newStatus === '방문완료' || newStatus === '제출완료') {
      sheet.getRange(safeRow, 14).setValue('Y');
    } else if (newStatus === '방문전' || newStatus === '예약확인중' || newStatus === '일정조율필요' || newStatus === '예약대기') {
      sheet.getRange(safeRow, 14).clearContent();
    }
    
    if (newStatus === '예약대기' || newStatus === '일정조율필요') {
      sheet.getRange(safeRow, 8, 1, 3).clearContent();
      sheet.getRange(safeRow, 15, 1, 7).clearContent();
    } else {
      if (newVisitDate && newVisitDate.trim() !== '') {
        const vDate = new Date(newVisitDate), dDate = new Date(vDate); 
        dDate.setDate(dDate.getDate() + 10);
        sheet.getRange(safeRow, 8).setValue(vDate);   
        sheet.getRange(safeRow, 10).setValue(dDate);  
        if(!sheet.getRange(safeRow, 11).getValue()) sheet.getRange(safeRow, 11).setValue(50000); 
      }
      if (newStatus === '방문전' && String(sheet.getRange(safeRow, 7).getValue() || '').trim() === '') {
        sheet.getRange(safeRow, 7).setValue('어드민_강제승인_패스');
      }
    }
    
    if (newStatus !== '예약확인중' && newStatus !== '일정조율필요' && newStatus !== '예약대기') {
      const fixUrl = (u) => { const s = String(u || '').trim(); return (s && !/^https?:\/\//i.test(s)) ? 'https://' + s : s; };
      sheet.getRange(safeRow, 17).setValue(fixUrl(newLink));
      sheet.getRange(safeRow, 18).setValue(fixUrl(newGoogleLink));

      if(newShortReview !== undefined && String(newShortReview).trim() !== '') {
        const cleanRev = String(newShortReview).trim();
        sheet.getRange(safeRow, 19).setValue(cleanRev);
        try { sheet.getRange(safeRow, 20).setValue(LanguageApp.translate(cleanRev, 'ko', 'ja')); } catch(e) {}
      }
      sheet.getRange(safeRow, 21).setValue((newStatus === '제출완료' && (!newRefundStatus || newRefundStatus.trim() === '')) ? '환급대기' : newRefundStatus);
    }
    
    // ✉️ 수동 어드민 패널 상태 제어 알림 (NOTICE 뱃지 및 인원수 보정 통합)
    if ((newStatus === '방문전' || newStatus === '일정조율필요') && currentMemberCode) {
      try {
        const userSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('User_DB');
        const userData = userSheet.getDataRange().getValues();
        let creatorEmail = "";
        
        for (let j = 1; j < userData.length; j++) {
          if (String(userData[j][0]).trim() === currentMemberCode) {
            creatorEmail = String(userData[j][12] || '').trim();
            break;
          }
        }

        if (creatorEmail && creatorEmail.includes("@")) {
          let subject = "";
          let htmlBody = "";
          
          if (newStatus === '방문전') {
            const rawPeopleStr = String(sheet.getRange(safeRow, 9).getValue() || '1');
            const pCount = rawPeopleStr.replace(/[^0-9]/g, '');

            subject = "🗓️ [BOOKMARK CREATORS] 방문 예약 확정 안내";
            htmlBody = `
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
                      <td style="color: #1a73e8; font-weight: 700; padding: 6px 0;">${newVisitDate || '-'}</td>
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
          } else if (newStatus === '일정조율필요') {
            subject = "🚨 [BOOKMARK CREATORS] 일정 조율 요청 안내";
            htmlBody = `
              <div style="max-width: 500px; margin: 0 auto; padding: 32px 20px; background: #ffffff; font-family: 'Apple SD Gothic Neo', 'Malgun Gothic', sans-serif; border: 1px solid #eef0f2; border-radius: 20px; box-shadow: 0 4px 12px rgba(0,0,0,0.02);">
                <div style="margin-bottom: 24px; text-align: left;">
                  <span style="font-size: 11px; font-weight: 800; letter-spacing: 1px; color: #ffffff; background: #e03131; padding: 4px 10px; border-radius: 6px; display: inline-block;">STATUS NOTICE</span>
                  <h2 style="font-size: 20px; font-weight: 800; color: #1A2B49; margin: 12px 0 0 0;">BOOKMARK CREATORS</h2>
                </div>
                <div style="border-top: 2px solid #e03131; padding-top: 24px; margin-bottom: 28px;">
                  <p style="font-size: 14.5px; font-weight: 700; color: #e03131; margin: 0 0 12px 0;">🟣 일정 변경 요청 알림</p>
                  <p style="font-size: 13.5px; line-height: 1.6; color: #495057; margin: 0;">
                    매장상황으로 일정 변경을 요청했습니다.<br>
                    <span style="font-weight: 700; color: #1A2B49;">담당자가 연락드리며 조율된 날짜로 다시 일정 예약해주세요.</span>
                  </p>
                </div>
                <div style="text-align: center;">
                  <a href="http://pf.kakao.com/_vFSxfX/chat" target="_blank" style="display: block; background: #e03131; color: #ffffff; font-size: 14px; font-weight: 700; text-decoration: none; padding: 14px; border-radius: 12px; box-shadow: 0 4px 12px rgba(224,49,49,0.15);">담당자에게 미리 연락하기</a>
                </div>
              </div>`;
          }
          GmailApp.sendEmail(creatorEmail, subject, "", { htmlBody: htmlBody, name: "BOOKMARK CREATORS" });
        }
      } catch (mailErr) {}
    }
    return { success: true };
  } catch(e) { return { success: false, error: e.toString() }; }
}

/********************************************************************
 * [4] 백그라운드 데이터 자동 배치 & 정합성 보완 필터
 ********************************************************************/
function checkAndMarkNoShow() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var logSheet = ss.getSheetByName("Master_Log");
  if (!logSheet) return;
  
  var range = logSheet.getDataRange();
  var values = range.getValues();
  var now = new Date();
  var todayStart = new Date(now.getFullYear(), now.getMonth(), now.getDate());
  
  for (var i = 2; i < values.length; i++) {
    var visitDateVal = values[i][7];  
    var currentStatus = String(values[i][11]).trim(); 
    
    if (visitDateVal instanceof Date) {
      var visitDate = new Date(visitDateVal);
      if (visitDate < todayStart && (currentStatus === "방문전" || currentStatus === "예약확정")) {
        logSheet.getRange(i + 1, 12).setValue("노쇼");       
        logSheet.getRange(i + 1, 21).setValue("보증금몰수");   
      }
    }
  }
}

function fillMissingData() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var logSheet = ss.getSheetByName("Master_Log");
  var userSheet = ss.getSheetByName("User_DB");
  if (!logSheet || !userSheet) return;
  
  var logRange = logSheet.getDataRange();
  var logData = logRange.getValues();
  var userData = userSheet.getDataRange().getValues();
  
  var userMap = {};
  for (var j = 1; j < userData.length; j++) {
    var mCode = String(userData[j][0]).trim(); 
    if (mCode && mCode !== "" && mCode !== "undefined") {
      userMap[mCode] = userData[j][1]; 
    }
  }
  
  var updatedRows = 0;
  for (var i = 2; i < logData.length; i++) {
    var row = i + 1;
    var memberCode = String(logData[i][1]).trim();         
    var currentEnglishName = String(logData[i][2]).trim(); 
    var visitDate = logData[i][7];                         
    
    if ((currentEnglishName === "" || currentEnglishName === "undefined" || currentEnglishName === "미승인/정보없음") && memberCode && userMap[memberCode]) {
      logSheet.getRange(row, 3).setValue(userMap[memberCode]);
      updatedRows++;
    }
    
    if (visitDate instanceof Date) {
      var deadlineVal = logData[i][9];
      var depositVal = logData[i][10];
      
      if (!deadlineVal || deadlineVal === "") {
        var d = new Date(visitDate);
        d.setDate(d.getDate() + 10); 
        logSheet.getRange(row, 10).setValue(d);
        updatedRows++;
      }
      if (!depositVal || depositVal === 0 || depositVal === "") {
        logSheet.getRange(row, 11).setValue(50000);
        updatedRows++;
      }
    }
  }
  SpreadsheetApp.getUi().alert("✨ 빈칸 자동 채우기가 완료되었습니다.\n(총 " + updatedRows + "개의 데이터 보완됨)");
}

function generateStorePins() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var storeSheet = ss.getSheetByName("Restaurant_List"); 
  if (!storeSheet) return;
  
  var lastRow = storeSheet.getLastRow();
  if (lastRow <= 1) return;
  
  var data = storeSheet.getRange(1, 1, lastRow, 9).getValues();
  var createdCount = 0;
  var charPool = "23456789ABCDEFGHJKLMNPQRSTUVWXYZ";
  var pinLength = 6; 
  
  for (var i = 1; i < data.length; i++) {
    var storeId = String(data[i][0]).trim();        
    var currentPin = String(data[i][8]).trim();     
    
    if (storeId !== "" && storeId !== "undefined" && (!currentPin || currentPin === "")) {
      var randPin = "";
      for (var j = 0; j < pinLength; j++) {
        var randomIndex = Math.floor(Math.random() * charPool.length);
        randPin += charPool.charAt(randomIndex);
      }
      storeSheet.getRange(i + 1, 9).setValue("'" + randPin);
      createdCount++;
    }
  }
  SpreadsheetApp.getUi().alert("🔒 PIN 생성 완료: 총 " + createdCount + "개 발급됨.");
}

function getChallengeSettings() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Settings");
  
  if (!sheet) {
    sheet = ss.insertSheet("Settings");
    sheet.appendRow(["항목", "설정값"]);
    sheet.appendRow(["START_DATE", "2026-06-01"]);
    sheet.appendRow(["END_DATE", "2026-11-30"]); 
    sheet.appendRow(["TARGET_COUNT", "20"]);        
  }
  
  var data = sheet.getDataRange().getValues();
  var settings = { startDate: "", endDate: "", targetCount: "20" };
  
  for (var i = 1; i < data.length; i++) {
    var key = String(data[i][0]).trim();
    var val = data[i][1];
    if (val instanceof Date) {
      val = Utilities.formatDate(val, Session.getScriptTimeZone(), "yyyy-MM-dd");
    }
    val = String(val).trim();
    
    if (key === "START_DATE") settings.startDate = val;
    if (key === "END_DATE") settings.endDate = val;
    if (key === "TARGET_COUNT") settings.targetCount = val;
  }
  return settings;
}

function saveChallengeSettings(startDate, endDate, targetCount) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Settings");
  if (!sheet) sheet = ss.insertSheet("Settings");
  
  sheet.clear();
  sheet.appendRow(["항목", "설정값"]);
  sheet.appendRow(["START_DATE", startDate]);
  sheet.appendRow(["END_DATE", endDate]);
  sheet.appendRow(["TARGET_COUNT", targetCount]);
  
  return "✨ 챌린지 시즌 설정이 성공적으로 저장되었습니다.";
}

function verifyAdminPassword(inputPw) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("Settings");
    if (!sheet) return false;
    
    var data = sheet.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][0]).trim() === "ADMIN_PASSWORD") {
        return String(data[i][1]).trim() === String(inputPw).trim();
      }
    }
    return false;
  } catch(e) { return false; }
}
