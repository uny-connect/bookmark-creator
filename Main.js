/********************************************************************
 * [1] 웹 앱 진입점 컨트롤러 (doGet 엔진)
 ********************************************************************/

/** 🌐 웹 앱 화면 라우팅 (GET) */
function doGet(e) {
  const mode = e?.parameter?.mode;
  
  // 🎯 [신규] 영업/담당자용 웹 테스트 발송 툴 분기 (store 분기보다 앞에 배치하여 파라미터 충돌 방지)
  if (mode === 'test_tool') {
    const template = HtmlService.createTemplateFromFile('TestTool');
    template.initialStoreId = e?.parameter?.id || '';
    return template.evaluate()
      .setTitle('BOOKMARK CREATORS | 店舗メール登録・テスト')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
      .addMetaTag('viewport', 'width=device-width, initial-scale=1, maximum-scale=1, user-scalable=no');
  }

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
        // 🔒 [중복 발송 방어 락] 최초 승인일 때만 메일을 발송하도록 분기 (브라우저 prefetch 등 2회 호출 원천 차단)
        const isFirstConfirm = (currentStatus !== '방문전');

        sheet.getRange(safeRow, 12).setValue('방문전'); 
        if (String(sheet.getRange(safeRow, 7).getValue() || '').trim() === '') {
          sheet.getRange(safeRow, 7).setValue('점주_직접_링크확정');
        }

        // 최초 1회 승인 처리일 때만 메일 엔진 동작
        if (isFirstConfirm) {
          try {
            const currentRestaurantName = String(sheet.getRange(safeRow, 6).getValue() || '').trim(); 
            const storeId = String(sheet.getRange(safeRow, 5).getValue() || '').trim().toUpperCase();
            const rawMemberCode = String(sheet.getRange(safeRow, 2).getValue() || '').trim();
            const currentMemberCode = rawMemberCode.replace(/['"\s]/g, '').toLowerCase();
            
            // 🎯 일본어 점포명 조회 (Restaurant_List C열)
            let storeNameJp = "";
            const restSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Restaurant_List');
            if (restSheet && storeId) {
              const restData = restSheet.getDataRange().getValues();
              for (let k = 2; k < restData.length; k++) {
                if (String(restData[k][0] || '').trim().toUpperCase() === storeId) {
                  storeNameJp = String(restData[k][2] || '').trim();
                  break;
                }
              }
            }
            const finalStoreNameJp = storeNameJp || currentRestaurantName;

            const timeZone = Session.getScriptTimeZone();
            const rawVisitDate = sheet.getRange(safeRow, 8).getValue();
            const rawPeopleStr = String(sheet.getRange(safeRow, 9).getValue() || '1');
            const pCount = rawPeopleStr.replace(/[^0-9]/g, '') || '1';
            const visitDateStr = (rawVisitDate instanceof Date) ? Utilities.formatDate(rawVisitDate, timeZone, 'yyyy-MM-dd HH:mm') : String(rawVisitDate || '-');

            if (currentMemberCode) {
              const userSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('User_DB');
              const userData = userSheet.getDataRange().getValues();
              let creatorEmail = "";
              
              for (let j = 1; j < userData.length; j++) {
                const dbMemberCode = String(userData[j][0] || '').trim().replace(/['"\s]/g, '').toLowerCase();
                if (dbMemberCode === currentMemberCode) {
                  creatorEmail = String(userData[j][13] || '').trim();
                  break;
                }
              }

              // ✉ 1. 크리에이터 대상 한국어 확정 안내 발송 (bcc 없이 단독 발송)
              if (creatorEmail && creatorEmail.includes("@")) {
                const subject = "[BOOKMARK CREATORS] 방문 예약 확정 안내";
                const htmlBody = `
                  <meta charset="UTF-8">
                  <div style="max-width: 500px; margin: 0 auto; padding: 32px 20px; background: #ffffff; font-family: 'Apple SD Gothic Neo', 'Malgun Gothic', sans-serif; border: 1px solid #eef0f2; border-radius: 20px; box-shadow: 0 4px 12px rgba(0,0,0,0.02);">
                    <div style="margin-bottom: 24px; text-align: left;">
                      <span style="font-size: 11px; font-weight: 800; letter-spacing: 1px; color: #ffffff; background: #1A2B49; padding: 4px 10px; border-radius: 6px; display: inline-block;">NOTICE</span>
                      <h2 style="font-size: 20px; font-weight: 800; color: #1A2B49; margin: 12px 0 0 0;">BOOKMARK CREATORS</h2>
                    </div>
                    <div style="border-top: 2px solid #1A2B49; padding-top: 24px; margin-bottom: 24px;">
                      <p style="font-size: 14.5px; font-weight: 700; color: #2D6A4F; margin: 0 0 12px 0;">[예약 확정 안내]</p>
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

                GmailApp.sendEmail(creatorEmail, subject, "", { 
                  htmlBody: htmlBody, 
                  name: "BOOKMARK CREATORS"
                });
              }

              // ✉️ 2. 관리자 대상 실시간 일본어 모니터링 알림 발송 (점포명 일본어 & 인원수 '名' 적용)
              try {
                const adminAlertEmails = getAdminAlertEmails();
                if (adminAlertEmails) {
                  const adminSubject = `【予約確定】店舗が予約を確定しました - #${orderNo}`;
                  const adminHtml = `
                    <meta charset="UTF-8">
                    <div style="font-family: 'Helvetica Neue', Arial, sans-serif; padding: 24px; background: #f8f9fa; border-radius: 16px; border: 1px solid #e9ecef; max-width: 520px; margin: 0 auto; color: #333; line-height: 1.6;">
                      <div style="margin-bottom: 18px;">
                        <span style="background: #2D6A4F; color: #fff; font-size: 11px; font-weight: bold; padding: 4px 10px; border-radius: 6px;">CONFIRMATION NOTICE</span>
                        <h3 style="color: #1A2B49; margin: 10px 0 0 0; font-size: 18px; font-weight: 800;">店舗による予約確定完了</h3>
                      </div>
                      <p style="font-size: 13.5px; color: #495057; margin: 0 0 16px 0;">
                        店舗管理者がメールリンクより来店予約を【確定】しました。<br>
                        クリエイターへ予約確定の案内メールが送信されました。
                      </p>
                      <div style="background: #ffffff; border-radius: 12px; padding: 18px; margin: 16px 0; border: 1px solid #eef0f2; font-size: 13.5px;">
                        <p style="margin: 6px 0;"><b>・注文番号:</b> #${orderNo}</p>
                        <p style="margin: 6px 0;"><b>・店舗名:</b> ${finalStoreNameJp}</p>
                        <p style="margin: 6px 0;"><b>・クリエイター:</b> ${currentMemberCode}</p>
                        <p style="margin: 6px 0;"><b>・予約日時:</b> <span style="color: #2D6A4F; font-weight: bold;">${visitDateStr} (${pCount}名)</span></p>
                      </div>
                    </div>
                  `;
                  GmailApp.sendEmail(adminAlertEmails, adminSubject, "", { 
                    htmlBody: adminHtml, 
                    name: "BOOKMARK NOTI" 
                  });
                }
              } catch(e) {}
            }
          } catch (mailErr) {
            console.error("❌ doGet 메일 엔진 연산 실패: " + mailErr.toString());
          }
        }
      }
    } catch(err) {
      console.error("❌ store_confirm 코어 에러: " + err.toString());
    }
    
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
    return template.evaluate().setTitle('BOOKMARK CREATORS | リクエスト').setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
      .addMetaTag('viewport', 'width=device-width, initial-scale=1, maximum-scale=1, user-scalable=no');
  }

  // 5️⃣ 기본 디폴트 화면 (크리에이터용 PIN 화면)
  const template = HtmlService.createTemplateFromFile('PIN화면2');
  template.orderNo = (e?.parameter?.o) || ''; 
  template.phoneLast4 = (e?.parameter?.p) || '';
  return template.evaluate().setTitle('BOOKMARK CREATORS | Creator').setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, maximum-scale=1, user-scalable=no');
}

/*******************************************************************************************************
 * [2] 시트 어시스턴트 유틸리티 매뉴얼 (onOpen / onEdit)
 *******************************************************************************************************/
function onOpen() {
  SpreadsheetApp.getUi().createMenu('⚙️ BOOKMARK CREATORS 관리')
    .addItem('✨ 빈칸 자동 채우기 (점포ID/이름/마감일/보증금)', 'fillMissingData')
    .addItem('📩 선택한 예약 건 점주 메일 재발송', 'resendStoreBookingEmail')
    .addItem('🔑 매장별 고유 PIN 6자리 생성', 'generateStorePins') 
    .addItem('🚨 자동 노쇼 일괄 처리 (과거 날짜)', 'checkAndMarkNoShow')
    .addItem('👥 아임웹 신규 회원 동기화', 'syncImwebUsers')
    .addItem('📧 매장 메일 수신 테스트 발송', 'sendTestEmailToStore')
    .addToUi();
}

function onEdit(e) {
  if (!e || !e.range) return; 
  const sheet = e.range.getSheet(); 
  if (sheet.getName() !== 'Master_Log') return;
  
  const row = e.range.getRow(); 
  const col = e.range.getColumn();
  const val = String(e.range.getValue()).trim();

  // 1️⃣ E열(5번째 열) 점포 ID 입력 시 ➔ F열(6번째 열) 점포명 실시간 자동 기입
  if (col === 5 && row >= 3) { 
    const storeId = val.toUpperCase();
    const storeNameCell = sheet.getRange(row, 6);
    const statusCell = sheet.getRange(row, 12);
    
    if (storeId === "") {
      storeNameCell.clearContent();
    } else {
      const ss = SpreadsheetApp.getActiveSpreadsheet();
      const restSheet = ss.getSheetByName('Restaurant_List');
      if (restSheet) {
        const restData = restSheet.getDataRange().getValues();
        let matchedName = "식당명 없음";
        
        for (let k = 2; k < restData.length; k++) {
          const sheetStoreId = String(restData[k][0] || '').trim().toUpperCase();
          if (sheetStoreId === storeId) {
            matchedName = String(restData[k][1] || '').trim();
            break;
          }
        }
        storeNameCell.setValue(matchedName);
      }
      
      if (String(statusCell.getValue()).trim() === "") {
        statusCell.setValue("예약대기");
      }
    }
  }

  // 2️⃣ 🎯 F열(6번째 열) 점포명 직접 입력 시 ➔ E열(5번째 열) 점포 ID 실시간 역방향 기입 + F열 공식 명칭 보정
  if (col === 6 && row >= 3) {
    const cleanStoreName = val.replace(/\s+/g, '');
    const storeIdCell = sheet.getRange(row, 5);
    const storeNameCell = sheet.getRange(row, 6);
    const statusCell = sheet.getRange(row, 12);

    if (cleanStoreName === "") {
      storeIdCell.clearContent();
    } else {
      const ss = SpreadsheetApp.getActiveSpreadsheet();
      const restSheet = ss.getSheetByName('Restaurant_List');
      if (restSheet) {
        const restData = restSheet.getDataRange().getValues();
        let matchedId = "";
        let matchedOfficialName = "";

        for (let k = 2; k < restData.length; k++) {
          const sheetStoreId = String(restData[k][0] || '').trim().toUpperCase();
          const sheetNameKo = String(restData[k][1] || '').trim();
          const sheetNameJp = String(restData[k][2] || '').trim();
          const cleanKo = sheetNameKo.replace(/\s+/g, '');
          const cleanJp = sheetNameJp.replace(/\s+/g, '');

          if ((cleanKo && (cleanKo === cleanStoreName || cleanStoreName.includes(cleanKo) || cleanKo.includes(cleanStoreName))) ||
              (cleanJp && (cleanJp === cleanStoreName || cleanStoreName.includes(cleanJp)))) {
            matchedId = sheetStoreId;
            matchedOfficialName = sheetNameKo;
            break;
          }
        }

        if (matchedId) {
          storeIdCell.setValue(matchedId);
          if (matchedOfficialName && matchedOfficialName !== val) {
            storeNameCell.setValue(matchedOfficialName); // 공식 명칭으로 자동 보정
          }
        }
      }

      if (String(statusCell.getValue()).trim() === "") {
        statusCell.setValue("예약대기");
      }
    }
  }

  // 3️⃣ H열(8번째 열) 방문예정일시 편집 시 J열(마감일), K열(보증금) 자동 연산
  if (col === 8 && row >= 3) { 
    const deadlineCell = sheet.getRange(row, 10);
    const depositCell = sheet.getRange(row, 11);
    const cellValue = e.range.getValue();
    
    if (cellValue instanceof Date) { 
      const d = new Date(cellValue); 
      d.setDate(d.getDate() + 10); 
      deadlineCell.setValue(d); 
      if (!depositCell.getValue()) depositCell.setValue(10000); // 🎯 10,000원 보정
    } else if (!cellValue) { 
      deadlineCell.clearContent(); 
    } 
  }

  // 4️⃣ L열(12번째 열) 진행 상태 변경 시 슬롯 리셋
  if (col === 12 && row >= 3) { 
    if (val === '예약대기' || val === '일정조율필요') {
      sheet.getRange(row, 8, 1, 3).clearContent();
    }
  }
}

/********************************************************************
 * [3] 어드민 코어 백엔드 API 연산 엔진
 ********************************************************************/
function getAdminData() {
  try {
    const data = _getSheetsData(['Master_Log', 'Restaurant_List', 'User_DB']);
    let stats = { total: 0, pending: 0, verified: 0, submitted: 0, waitingReview: 0 };
    let missions = [], restList = [], restNameMap = {}, userTierMap = {};

    const userDB = data.User_DB;
    for (let u = 1; u < userDB.length; u++) {
      if(userDB[u][0]) {
        try {
          userTierMap[String(userDB[u][0]).trim()] = parseTierEmoji(userDB[u][9]);
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
        if(!sheet.getRange(safeRow, 11).getValue()) sheet.getRange(safeRow, 11).setValue(10000); 
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
    
    if ((newStatus === '방문전' || newStatus === '일정조율필요') && currentMemberCode) {
      try {
        const userSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('User_DB');
        const userData = userSheet.getDataRange().getValues();
        let creatorEmail = "";
        
        for (let j = 1; j < userData.length; j++) {
          if (String(userData[j][0]).trim() === currentMemberCode) {
            creatorEmail = String(userData[j][13] || '').trim();
            break;
          }
        }

        if (creatorEmail && creatorEmail.includes("@")) {
          let subject = "";
          let htmlBody = "";
          
          if (newStatus === '방문전') {
            const rawPeopleStr = String(sheet.getRange(safeRow, 9).getValue() || '1');
            const pCount = rawPeopleStr.replace(/[^0-9]/g, '') || '1';

            subject = "[BOOKMARK CREATORS] 방문 예약 확정 안내";
            htmlBody = `
              <meta charset="UTF-8">
              <div style="max-width: 500px; margin: 0 auto; padding: 32px 20px; background: #ffffff; font-family: 'Apple SD Gothic Neo', 'Malgun Gothic', sans-serif; border: 1px solid #eef0f2; border-radius: 20px; box-shadow: 0 4px 12px rgba(0,0,0,0.02);">
                <div style="margin-bottom: 24px; text-align: left;">
                  <span style="font-size: 11px; font-weight: 800; letter-spacing: 1px; color: #ffffff; background: #1A2B49; padding: 4px 10px; border-radius: 6px; display: inline-block;">NOTICE</span>
                  <h2 style="font-size: 20px; font-weight: 800; color: #1A2B49; margin: 12px 0 0 0;">BOOKMARK CREATORS</h2>
                </div>
                <div style="border-top: 2px solid #1A2B49; padding-top: 24px; margin-bottom: 24px;">
                  <p style="font-size: 14.5px; font-weight: 700; color: #2D6A4F; margin: 0 0 12px 0;">[예약 확정 알림]</p>
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
            subject = "[BOOKMARK CREATORS] 일정 조율 요청 안내";
            htmlBody = `
              <meta charset="UTF-8">
              <div style="max-width: 500px; margin: 0 auto; padding: 32px 20px; background: #ffffff; font-family: 'Apple SD Gothic Neo', 'Malgun Gothic', sans-serif; border: 1px solid #eef0f2; border-radius: 20px; box-shadow: 0 4px 12px rgba(0,0,0,0.02);">
                <div style="margin-bottom: 24px; text-align: left;">
                  <span style="font-size: 11px; font-weight: 800; letter-spacing: 1px; color: #ffffff; background: #e03131; padding: 4px 10px; border-radius: 6px; display: inline-block;">STATUS NOTICE</span>
                  <h2 style="font-size: 20px; font-weight: 800; color: #1A2B49; margin: 12px 0 0 0;">BOOKMARK CREATORS</h2>
                </div>
                <div style="border-top: 2px solid #e03131; padding-top: 24px; margin-bottom: 28px;">
                  <p style="font-size: 14.5px; font-weight: 700; color: #e03131; margin: 0 0 12px 0;">[일정 변경 요청 알림]</p>
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

          GmailApp.sendEmail(creatorEmail, subject, "", { 
            htmlBody: htmlBody, 
            name: "BOOKMARK CREATORS"
          });
        }
      } catch (mailErr) {}
    }
    return { success: true };
  } catch(e) { return { success: false, error: e.toString() }; }
}

function sendTestEmailToStore() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const activeSheet = ss.getActiveSheet();
  const restSheet = ss.getSheetByName('Restaurant_List');
  
  if (!restSheet) return ui.alert('❌ Restaurant_List 시트를 찾을 수 없습니다.');
  
  let targetEmail = "";
  let storeNameJp = "";
  let storeId = "";

  if (activeSheet.getName().trim() === 'Restaurant_List') {
    const selectedRow = activeSheet.getActiveRange().getRow();
    
    if (selectedRow >= 2) {
      const rowValues = restSheet.getRange(selectedRow, 1, 1, 11).getValues()[0];
      storeId = String(rowValues[0] || '').trim();
      const nameKo = String(rowValues[1] || '').trim();
      const nameJp = String(rowValues[2] || '').trim();
      const email = String(rowValues[10] || '').trim();

      if (storeId && email && email.includes('@')) {
        storeNameJp = nameJp || nameKo;
        targetEmail = email;

        const confirm = ui.alert(
          '📧 선택된 매장 테스트 발송',
          `[선택된 매장 정보]\n• 점포 ID: ${storeId}\n• 매장명: ${storeNameJp}\n• 수신 이메일: ${targetEmail}\n\n이 매장으로 테스트 메일을 발송하시겠습니까?`,
          ui.ButtonSet.YES_NO
        );

        if (confirm !== ui.Button.YES) return;
      }
    }
  }

  if (!targetEmail) {
    const response = ui.prompt(
      '📧 매장 메일 수신 테스트',
      '테스트할 매장의 [점포 ID]를 입력하세요:\n(예: FUK01)',
      ui.ButtonSet.OK_CANCEL
    );

    if (response.getSelectedButton() !== ui.Button.OK) return;
    const inputVal = response.getResponseText().trim();
    if (!inputVal) return ui.alert('⚠️ 점포 ID를 입력해주세요.');

    const restData = restSheet.getDataRange().getValues();
    for (let i = 1; i < restData.length; i++) {
      const sheetStoreId = String(restData[i][0] || '').trim().toUpperCase();
      if (sheetStoreId === inputVal.toUpperCase()) {
        storeId = sheetStoreId;
        targetEmail = String(restData[i][10] || '').trim();
        storeNameJp = String(restData[i][2] || '').trim() || String(restData[i][1] || '').trim();
        break;
      }
    }
  }

  if (!targetEmail || !targetEmail.includes('@')) {
    return ui.alert(`❌ 유효한 이메일 주소를 찾을 수 없습니다.\nRestaurant_List 시트의 K열 이메일을 확인해 주세요.`);
  }

  try {
    const todayStr = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd");
    const subject = `【TEST / 接続テスト】 BOOKMARK CREATORS 予約通知テスト送信 (${storeNameJp})`;
    
    const htmlBody = `
      <meta charset="UTF-8">
      <div style="font-family: 'Helvetica Neue', Arial, sans-serif; padding: 20px; background-color: #f4f5f7;">
        <div style="max-width: 500px; margin: 0 auto; background: #ffffff; border-radius: 12px; padding: 30px; box-shadow: 0 4px 10px rgba(0,0,0,0.05);">
          <div style="text-align: center; margin-bottom: 20px;">
            <span style="font-size: 11px; font-weight: 800; letter-spacing: 1px; color: #ffffff; background: #e03131; padding: 4px 10px; border-radius: 6px; display: inline-block;">CONNECTION TEST</span>
            <h1 style="color: #1A2B49; margin: 10px 0 0 0; font-size: 22px; font-weight: 900; letter-spacing: -0.5px;">BOOKMARK CREATORS</h1>
            <div style="width: 40px; height: 3px; background: #C5A358; margin: 10px auto;"></div>
          </div>
          <h2 style="color: #1A2B49; margin-top: 0; font-size: 16px; border-bottom: 2px solid #f1f3f5; padding-bottom: 15px; text-align: center;">&#128236; メール受信テスト</h2>
          <div style="margin-top: 20px;">
            <p style="color: #1A2B49; font-size: 15px; font-weight: bold; margin-bottom: 5px;">${storeNameJp}</p>
            <p style="color: #495057; font-size: 13.5px; margin-top: 0;">店舗管理者様</p>
          </div>
          <p style="color: #495057; font-size: 13.5px; line-height: 1.6;">
            本メールは、BOOKMARK CREATORSからの予約受付通知が正常に届くか確認するための<b>【テストメール】</b>です。<br>
            このメールが確認できましたら、今後のクリエイター来店予約通知も問題なく受信いただけます。
          </p>
          <div style="background-color: #f8f9fa; border-left: 4px solid #2D6A4F; padding: 15px; border-radius: 4px; margin: 20px 0;">
            <p style="margin: 4px 0; font-size: 13.5px;"><strong>■ 店舗ID:</strong> ${storeId || '-'}</p>
            <p style="margin: 4px 0; font-size: 13.5px;"><strong>■ 受信状態:</strong> <span style="color: #2D6A4F; font-weight: bold;">正常受信完了 (Success)</span></p>
            <p style="margin: 4px 0; font-size: 13.5px;"><strong>■ テスト日時:</strong> ${todayStr}</p>
            <p style="margin: 4px 0; font-size: 13.5px;"><strong>■ 対象アドレス:</strong> ${targetEmail}</p>
          </div>
          <div style="text-align: center; margin-top: 24px; padding: 12px; background-color: #f1f3f5; border-radius: 8px; font-size: 12px; color: #6c757d; font-weight: 600;">
            ※ 本メールに対する返信や確定手続き는 不要です。
          </div>
        </div>
      </div>`;

    GmailApp.sendEmail(targetEmail, subject, "", {
      htmlBody: htmlBody,
      name: "BOOKMARK CREATORS",
      from: "info@bookmarkfukuoka.jp"
    });

    ui.alert(`✅ [테스트 발송 성공]\n\n• 매장명: ${storeNameJp}\n• 수신 이메일: ${targetEmail}\n\n해당 매장 메일함을 확인해 주세요.`);
  } catch (err) {
    ui.alert(`❌ [발송 실패 에러]\n\n${err.toString()}`);
  }
}

/** [매장별 고유 PIN 6자리 생성 엔진] */
function generateStorePins() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const storeSheet = ss.getSheetByName("Restaurant_List"); 
  if (!storeSheet) {
    SpreadsheetApp.getUi().alert("❌ Restaurant_List 시트를 찾을 수 없습니다.");
    return;
  }
  
  const lastRow = storeSheet.getLastRow();
  if (lastRow <= 2) {
    SpreadsheetApp.getUi().alert("⚠️ 등록된 매장 데이터가 없습니다.");
    return;
  }
  
  // A열부터 I열(9번째 열)까지 데이터 로드 (헤더 제외 3번째 행부터)
  const data = storeSheet.getRange(3, 1, lastRow - 2, 9).getValues();
  let createdCount = 0;
  const charPool = "23456789ABCDEFGHJKLMNPQRSTUVWXYZ"; // 헷갈리는 0, 1, I, O 제외
  const pinLength = 6; 
  
  for (let i = 0; i < data.length; i++) {
    const storeId = String(data[i][0] || "").trim(); // A열: 매장 ID
    const currentPin = String(data[i][8] || "").trim(); // I열: PIN
    
    // 매장 ID는 있는데 PIN 번호가 비어있는 경우에만 신규 발급
    if (storeId !== "" && (!currentPin || currentPin === "undefined")) {
      let randPin = "";
      for (let j = 0; j < pinLength; j++) {
        const randomIndex = Math.floor(Math.random() * charPool.length);
        randPin += charPool.charAt(randomIndex);
      }
      
      // I열(9번째 열)에 텍스트 형태('PIN)로 안전하게 주입
      storeSheet.getRange(i + 3, 9).setValue("'" + randPin);
      createdCount++;
    }
  }
  
  SpreadsheetApp.getUi().alert(`🔒 PIN 생성 완료: 총 ${createdCount}개의 매장 PIN이 새로 발급되었습니다.`);
}

/********************************************************************
 * 🔐 [어드민 인증 및 챌린지 설정 엔진 + 알림 이메일 로더]
 ********************************************************************/

/** 🎯 Settings 시트에서 운영진 알림 수신 이메일 목록 불러오기 도우미 */
function getAdminAlertEmails() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName("Settings");
    if (!sheet) return "bookmarkjapan.info@gmail.com";

    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      const key = String(data[i][0] || '').trim().toUpperCase();
      if (key === "ADMIN_NOTIFICATION_EMAILS") {
        const emails = String(data[i][1] || '').trim();
        return emails || "bookmarkjapan.info@gmail.com";
      }
    }
  } catch (e) {
    Logger.log("getAdminAlertEmails 에러: " + e.toString());
  }
  return "bookmarkjapan.info@gmail.com";
}

/** [1] 관리자 패스워드 검증 함수 (공백 무시 및 강력한 정규화) */
function verifyAdminPassword(inputPw) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName("Settings");
    if (!sheet) {
      Logger.log("❌ Settings 시트를 찾을 수 없습니다.");
      return false;
    }

    const cleanInput = String(inputPw || '').replace(/[\s\uFEFF\xA0]+/g, ''); // 모든 공백 제거
    const data = sheet.getDataRange().getValues();

    for (let i = 1; i < data.length; i++) {
      const key = String(data[i][0] || '').replace(/[\s_]+/g, '').toUpperCase();
      // 'ADMINPASSWORD' 또는 'ADMINPW' 키 매칭
      if (key === "ADMINPASSWORD" || key === "ADMINPW") {
        const savedPw = String(data[i][1] || '').replace(/[\s\uFEFF\xA0]+/g, '');
        Logger.log(`[비밀번호 검증] 시트저장값: [${savedPw}] vs 입력값: [${cleanInput}]`);
        return savedPw === cleanInput;
      }
    }
    
    Logger.log("❌ Settings 시트에서 ADMIN_PASSWORD 항목을 찾지 못했습니다.");
    return false;
  } catch (e) {
    Logger.log("verifyAdminPassword 에러: " + e.toString());
    return false;
  }
}

/** [2] 챌린지 시즌 설정 불러오기 */
function getChallengeSettings() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName("Settings");
    if (!sheet) return { startDate: "2026-06-01", endDate: "2026-12-31", targetCount: "20" };

    const data = sheet.getDataRange().getValues();
    const settings = { startDate: "2026-06-01", endDate: "2026-12-31", targetCount: "20" };
    const timeZone = Session.getScriptTimeZone() || "Asia/Seoul";

    for (let i = 1; i < data.length; i++) {
      const key = String(data[i][0] || '').trim().toUpperCase();
      let val = data[i][1];
      if (val instanceof Date) {
        val = Utilities.formatDate(val, timeZone, "yyyy-MM-dd");
      }
      val = String(val || '').trim();

      if (key === "START_DATE") settings.startDate = val;
      if (key === "END_DATE") settings.endDate = val;
      if (key === "TARGET_COUNT") settings.targetCount = val;
    }
    return settings;
  } catch (e) {
    return { startDate: "2026-06-01", endDate: "2026-12-31", targetCount: "20" };
  }
}

/** [3] 챌린지 시즌 설정 저장하기 */
function saveChallengeSettings(startDate, endDate, targetCount) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName("Settings");
    if (!sheet) return "Settings 시트가 없습니다.";

    const data = sheet.getDataRange().getValues();
    let startRow = -1, endRow = -1, targetRow = -1;

    for (let i = 1; i < data.length; i++) {
      const key = String(data[i][0] || '').trim().toUpperCase();
      if (key === "START_DATE") startRow = i + 1;
      if (key === "END_DATE") endRow = i + 1;
      if (key === "TARGET_COUNT") targetRow = i + 1;
    }

    if (startRow > 0) sheet.getRange(startRow, 2).setValue(startDate);
    if (endRow > 0) sheet.getRange(endRow, 2).setValue(endDate);
    if (targetRow > 0) sheet.getRange(targetRow, 2).setValue(targetCount);

    return "기준 저장 완료";
  } catch (e) {
    return "저장 실패: " + e.toString();
  }
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
