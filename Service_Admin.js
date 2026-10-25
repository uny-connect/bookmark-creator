/** 관리자 도메인: 대시보드 API, 인증/챌린지 설정, 시트 메뉴 작업(노쇼·빈칸 채우기·PIN·취소·재발송) */

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
      if (newStatus === '방문전') {
        // 🎯 [G열 비우기] Y열(25열)에 기존 로그 보존하며 누적 기록
        const prevLogY = String(sheet.getRange(safeRow, 25).getValue() || '').trim();
        const adminTime = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm");
        const appendLog = `[${adminTime}] 어드민_강제승인_패스`;
        sheet.getRange(safeRow, 25).setValue(prevLogY ? `${prevLogY} | ${appendLog}` : appendLog);
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

/** 🚨 [자동 노쇼 일괄 처리 엔진] */
function checkAndMarkNoShow() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Master_Log');
  if (!sheet) return;

  try {
    const data = sheet.getDataRange().getValues();
    if (data.length < 3) return;

    const now = new Date();
    const timeZone = Session.getScriptTimeZone() || "Asia/Seoul";
    let updatedCount = 0;

    for (let i = 2; i < data.length; i++) {
      const row = data[i];
      const rowNum = i + 1;

      const orderNo = String(row[0] || '').trim().replace(/'/g, ''); 
      const status = String(row[11] || '').trim();                    
      const visitDateRaw = row[7];                                    

      if (!orderNo || status !== '방문전') continue;

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

      if (!visitDate || isNaN(visitDate.getTime())) continue;

      const safeDeadline = new Date(visitDate.getTime());
      safeDeadline.setDate(safeDeadline.getDate() + 1);
      safeDeadline.setHours(4, 0, 0, 0);

      if (now > safeDeadline) {
        sheet.getRange(rowNum, 12).setValue('노쇼');
        
        const currentSysLog = String(sheet.getRange(rowNum, 25).getValue() || '').trim();
        const autoLog = `[시스템] ${Utilities.formatDate(now, timeZone, 'yyyy-MM-dd HH:mm')} 자동 노쇼 처리`;
        const updatedSysLog = currentSysLog ? `${currentSysLog} | ${autoLog}` : autoLog;
        sheet.getRange(rowNum, 25).setValue(updatedSysLog);

        updatedCount++;
      }
    }
    
    if (SpreadsheetApp.getUi) {
      try {
        SpreadsheetApp.getUi().alert(`✅ 노쇼 처리 완료\n\n총 ${updatedCount}건이 노쇼 처리되었습니다.`);
      } catch(uiErr) {}
    }
  } catch (err) {
    Logger.log(`❌ [checkAndMarkNoShow 오류] ${err.toString()}`);
  }
}

/** 🎯 [통합 완결본] Master_Log 누락 데이터 일괄 복구 및 자동 완성 */
function fillMissingData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const logSheet = ss.getSheetByName("Master_Log");
  const userSheet = ss.getSheetByName("User_DB");
  const restSheet = ss.getSheetByName("Restaurant_List");
  
  if (!logSheet || !userSheet || !restSheet) {
    SpreadsheetApp.getUi().alert("❌ 시트 확인 필요: Master_Log, User_DB, Restaurant_List 중 누락된 시트가 있습니다.");
    return;
  }

  const lastRow = logSheet.getLastRow();
  if (lastRow < MASTER_LOG_START_ROW) {
    SpreadsheetApp.getUi().alert("⚠️ 검사할 데이터가 없습니다.");
    return;
  }

  const logData = logSheet.getRange(3, 1, lastRow - 2, 25).getValues();
  const userData = userSheet.getDataRange().getValues();
  const restData = restSheet.getDataRange().getValues();

  const userMap = new Map();
  for (let j = 1; j < userData.length; j++) {
    const code = String(userData[j][0] || '').trim().toLowerCase();
    const engName = String(userData[j][1] || '').trim();
    if (code && engName) userMap.set(code, engName);
  }

  const idToNameMap = new Map();
  const nameToIdMap = new Map();

  for (let k = REST_DATA_START_INDEX; k < restData.length; k++) {
    const sId = String(restData[k][0] || '').trim().toUpperCase();
    const sNameKo = String(restData[k][1] || '').trim();
    const sNameJp = String(restData[k][2] || '').trim();

    if (sId && sNameKo) {
      idToNameMap.set(sId, sNameKo);
      nameToIdMap.set(sNameKo.replace(/\s+/g, ''), { id: sId, name: sNameKo });
    }
    if (sId && sNameJp) {
      nameToIdMap.set(sNameJp.replace(/\s+/g, ''), { id: sId, name: sNameKo });
    }
  }

  let updatedCount = 0;

  for (let i = 0; i < logData.length; i++) {
    const rowNum = i + 3;
    const memberCode = String(logData[i][1] || '').trim().toLowerCase();
    const currentEngName = String(logData[i][2] || '').trim();
    const currentStoreId = String(logData[i][4] || '').trim().toUpperCase();
    const currentStoreName = String(logData[i][5] || '').trim();
    const cleanStoreName = currentStoreName.replace(/\s+/g, '');
    const visitDateVal = logData[i][7];
    const deadlineVal = logData[i][9];
    const depositVal = logData[i][10];

    let rowChanged = false;

    if ((!currentEngName || currentEngName === "미승인/정보없음" || currentEngName === "undefined") && userMap.has(memberCode)) {
      logSheet.getRange(rowNum, 3).setValue(userMap.get(memberCode));
      rowChanged = true;
    }

    if (!currentStoreId && cleanStoreName) {
      for (let [cleanName, info] of nameToIdMap.entries()) {
        if (cleanStoreName.includes(cleanName) || cleanName.includes(cleanStoreName)) {
          logSheet.getRange(rowNum, 5).setValue(info.id);
          logSheet.getRange(rowNum, 6).setValue(info.name);
          rowChanged = true;
          break;
        }
      }
    } else if (currentStoreId && (!currentStoreName || currentStoreName === "식당명 없음") && idToNameMap.has(currentStoreId)) {
      logSheet.getRange(rowNum, 6).setValue(idToNameMap.get(currentStoreId));
      rowChanged = true;
    }

    if (visitDateVal instanceof Date && !isNaN(visitDateVal.getTime())) {
      if (!deadlineVal || String(deadlineVal).trim() === '') {
        const d = new Date(visitDateVal.getTime());
        d.setDate(d.getDate() + 10);
        logSheet.getRange(rowNum, 10).setValue(d);
        rowChanged = true;
      }
    }

    if (!depositVal || depositVal === 0 || String(depositVal).trim() === '') {
      logSheet.getRange(rowNum, 11).setValue(10000);
      rowChanged = true;
    }

    if (rowChanged) updatedCount++;
  }

  SpreadsheetApp.getUi().alert(`✨ 자동 채우기 완료!\n\n총 ${updatedCount}개 행의 누락 데이터가 성공적으로 보정되었습니다.`);
}

/** [관리자 기능] 선택된 행(또는 입력)으로 점주 예약 안내 메일 재발송 (구매옵션 포함 완결본) */
function resendStoreBookingEmail() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const activeSheet = ss.getActiveSheet();
  const logSheet = ss.getSheetByName("Master_Log");
  const restSheet = ss.getSheetByName("Restaurant_List");
  const userSheet = ss.getSheetByName("User_DB");

  if (!logSheet || !restSheet || !userSheet) return ui.alert("❌ 시트 확인 필요: Master_Log, Restaurant_List, User_DB");

  let targetRow = activeSheet.getActiveRange().getRow();

  if (activeSheet.getName() !== "Master_Log" || targetRow < MASTER_LOG_START_ROW) {
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

  if (isNaN(targetRow) || targetRow < MASTER_LOG_START_ROW) {
    return ui.alert("⚠️ 올바른 데이터 행 번호를 선택하거나 입력해주세요.");
  }

  const rowData = logSheet.getRange(targetRow, 1, 1, 25).getValues()[0];
  const orderNo = String(rowData[0] || '').replace(/'/g, '').trim();
  const memberCode = String(rowData[1] || '').trim();
  const memberName = String(rowData[2] || '').trim();
  const storeId = String(rowData[4] || '').trim().toUpperCase();
  const mealOption = String(rowData[6] || '').trim(); // 🎯 G열 구매옵션 연동
  const rawVisitDate = rowData[7];
  const peopleStr = String(rowData[8] || '1명').replace(/[^0-9]/g, '') || '1';

  if (!orderNo || !rawVisitDate) {
    return ui.alert(`❌ [재발송 불가]\n\n${targetRow}행에 주문번호나 방문예정일시가 없습니다.`);
  }

  const timeZone = Session.getScriptTimeZone() || "Asia/Seoul";
  const dateStr = (rawVisitDate instanceof Date) ? Utilities.formatDate(rawVisitDate, timeZone, "yyyy-MM-dd") : String(rawVisitDate).substring(0, 10);
  const timeStr = (rawVisitDate instanceof Date) ? Utilities.formatDate(rawVisitDate, timeZone, "HH:mm") : String(rawVisitDate).substring(11, 16);

  let storeEmail = "", storeNameJp = "", storeBenefit = "";
  const restData = restSheet.getDataRange().getValues();
  
  for (let k = REST_DATA_START_INDEX; k < restData.length; k++) {
    if (String(restData[k][0] || '').trim().toUpperCase() === storeId) {
      storeNameJp = String(restData[k][2] || '').trim() || String(restData[k][1] || '').trim();
      storeEmail = String(restData[k][10] || '').trim();
      storeBenefit = String(restData[k][22] || '').trim();
      break;
    }
  }

  if (!storeEmail || !storeEmail.includes("@")) {
    return ui.alert(`❌ 매장 이메일 주소를 찾을 수 없습니다.\nRestaurant_List 시트의 [${storeId}] 매장 K열을 확인해 주세요.`);
  }

  const confirm = ui.alert(
    "📧 점주 예약 메일 재발송",
    `[재발송 대상 정보]\n• 행 번호: ${targetRow}행\n• 주문번호: #${orderNo}\n• 크리에이터: ${memberName}\n• 플랜: ${mealOption || '기본'}\n• 매장: ${storeNameJp} (${storeId})\n• 예약일시: ${dateStr} ${timeStr} (${peopleStr}명)\n• 수신처: ${storeEmail}\n\n지금 점주에게 예약 안내 메일을 재발송하시겠습니까?`,
    ui.ButtonSet.YES_NO
  );
  if (confirm !== ui.Button.YES) return;

  try {
    const scriptUrl = getActiveWebAppUrl();
    const confirmUrl = `${scriptUrl}?mode=store_confirm&row=${targetRow}&o=${encodeURIComponent(orderNo)}`;
    const feedbackUrl = `${scriptUrl}?mode=feedback&row=${targetRow}&o=${encodeURIComponent(orderNo)}`;

    let creatorProfileUrl = "";
    const userData = userSheet.getDataRange().getValues();
    for (let u = 1; u < userData.length; u++) {
      if (String(userData[u][0] || '').trim().toLowerCase() === memberCode.toLowerCase()) {
        const blogUrl = String(userData[u][4] || '').trim();        
        const instaUrl = String(userData[u][5] || '').trim();       
        const youtubeUrl = String(userData[u][6] || '').trim();     
        const tiktokUrl = String(userData[u][7] || '').trim();      
        const googleGuide = String(userData[u][8] || '').trim();    

        creatorProfileUrl = blogUrl || instaUrl || youtubeUrl || tiktokUrl || googleGuide || "";
        break;
      }
    }

    let profileHtmlClean = `<span style="color: #8b95a1; font-weight: 500;">当日確認</span>`;
    if (creatorProfileUrl) {
      const isHttp = /^https?:\/\//i.test(creatorProfileUrl);
      const isDomainLike = /^[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}/.test(creatorProfileUrl);
      if (isHttp || isDomainLike) {
        const fullUrl = isHttp ? creatorProfileUrl : "https://" + creatorProfileUrl;
        const btnText = (creatorProfileUrl.includes('drive.google.com') || creatorProfileUrl.includes('imweb') || creatorProfileUrl.match(/\.(jpg|jpeg|png|webp|gif)/i))
          ? "プロフィール確認 ↗" : "SNSを見る ↗";
        profileHtmlClean = `<a href="${fullUrl}" target="_blank" style="color: #1a73e8; font-weight: bold; text-decoration: underline;">${btnText}</a>`;
      } else {
        profileHtmlClean = `<span style="color: #1A2B49; font-weight: bold;">Google Local Guides (${creatorProfileUrl})</span>`;
      }
    }

    const benefitDisplay = storeBenefit || "店舗指定のクリエイター向け提供メニュー";
    const subject = `【BOOKMARK CREATORS】 ご来店予約のご確認 (${dateStr})`;
    const htmlBody = `
      <meta charset="UTF-8">
      <div style="font-family: -apple-system, BlinkMacSystemFont, 'Hiragino Kaku Gothic ProN', 'Hiragino Sans', Meiryo, sans-serif; padding: 20px 10px; background-color: #f4f5f7; -webkit-font-smoothing: antialiased;">
        <table align="center" border="0" cellpadding="0" cellspacing="0" width="100%" style="max-width: 500px; background-color: #ffffff; border-radius: 16px; border: 1px solid #edf0f2; box-shadow: 0 4px 16px rgba(0,0,0,0.04); overflow: hidden;">
          <tr>
            <td style="padding: 32px 24px;">
              
              <div style="text-align: center; margin-bottom: 20px;">
                <span style="font-size: 11px; font-weight: 800; letter-spacing: 1px; color: #ffffff; background-color: #1A2B49; padding: 4px 10px; border-radius: 6px; display: inline-block;">RESERVATION REQUEST</span>
                <h1 style="color: #1A2B49; margin: 10px 0 0 0; font-size: 22px; font-weight: 900; letter-spacing: -0.5px;">BOOKMARK CREATORS</h1>
                <div style="width: 36px; height: 3px; background-color: #C5A358; margin: 10px auto 0;"></div>
              </div>

              <h2 style="color: #1A2B49; margin: 18px 0 0 0; font-size: 16.5px; border-bottom: 1.5px solid #f1f3f5; padding-bottom: 14px; text-align: center; font-weight: 800;">来店予約の確認</h2>

              <div style="margin-top: 18px;">
                <p style="color: #1A2B49; font-size: 15px; font-weight: bold; margin: 0 0 2px 0;">${storeNameJp}</p>
                <p style="color: #6c757d; font-size: 14px; margin: 0 0 6px 0;">店舗管理者様</p>
              </div>

              <div style="margin: 6px 0 16px 0; color: #333333; font-size: 14px; line-height: 1.55;">
                <p style="margin: 0 0 3px 0;">BOOKMARK CREATORSより来店予約の申請が届きました。</p>
                <p style="margin: 0;">下記の内容をご確認いただき、<b>【確定】</b>または<b>【日時変更】</b>のご対応をお願い申し上げます。</p>
              </div>

              <div style="background-color: #fbfbfc; border: 1px solid #edf0f2; border-radius: 10px; padding: 12px 16px; margin-bottom: 20px; text-align: left;">
                <p style="margin: 0; font-size: 13px; color: #c92a2a; line-height: 1.55; font-weight: 700;">
                  ※ 本メール送信後24時間以内にご回答がない場合、自動的にキャンセル処理が行われます。
                </p>
              </div>

              <table border="0" cellpadding="0" cellspacing="0" width="100%" style="background-color: #ffffff; border: 1.5px solid #edf0f2; border-radius: 12px; padding: 14px 16px; margin-bottom: 24px; font-size: 13.5px;">
                <tr>
                  <td style="padding: 6px 0; color: #333333;">
                    <strong>■ クリエイター名:</strong> ${memberName}
                  </td>
                </tr>
                <tr>
                  <td style="padding: 6px 0; color: #333333;">
                    <strong>■ 訪問日時:</strong> <span style="color: #d63384; font-weight: 800;">${dateStr} ${timeStr}</span>
                  </td>
                </tr>
                <tr>
                  <td style="padding: 6px 0; color: #333333;">
                    <strong>■ 訪問人数:</strong> ${peopleStr}名
                  </td>
                </tr>
                <tr>
                  <td style="padding: 6px 0; color: #333333;">
                    <strong>■ 提供内容:</strong> <span style="color: #2D6A4F; font-weight: 700;">${benefitDisplay}</span>
                  </td>
                </tr>
                <tr>
                  <td style="padding: 6px 0; color: #333333;">
                    <strong>■ SNS:</strong> ${profileHtmlClean}
                  </td>
                </tr>
              </table>

              <table border="0" cellpadding="0" cellspacing="0" width="100%" style="margin: 0 auto;">
                <tr>
                  <td align="center" width="50%" style="padding-right: 6px;">
                    <a href="${confirmUrl}" target="_blank" style="display: block; background-color: #2D6A4F; color: #ffffff; padding: 13px 0; font-size: 13.5px; font-weight: 800; text-decoration: none; border-radius: 10px; text-align: center; box-shadow: 0 3px 8px rgba(45,106,79,0.18);">予約を確定する</a>
                  </td>
                  <td align="center" width="50%" style="padding-left: 6px;">
                    <a href="${feedbackUrl}" target="_blank" style="display: block; background-color: #1A2B49; color: #ffffff; padding: 13px 0; font-size: 13.5px; font-weight: 800; text-decoration: none; border-radius: 10px; text-align: center; box-shadow: 0 3px 8px rgba(26,43,73,0.15);">日時変更リクエスト</a>
                  </td>
                </tr>
              </table>

            </td>
          </tr>
        </table>
      </div>
    `;

    const adminAlertEmails = (typeof getAdminAlertEmails === 'function') ? getAdminAlertEmails() : "bookmarkjapan.info@gmail.com";

    GmailApp.sendEmail(storeEmail, subject, "", {
      htmlBody: htmlBody,
      name: "BOOKMARK CREATORS",
      from: "info@bookmarkfukuoka.jp",
      bcc: adminAlertEmails
    });

    const nowTimeStr = Utilities.formatDate(new Date(), timeZone, "yyyy-MM-dd HH:mm");
    logSheet.getRange(targetRow, 25).setValue(`[${nowTimeStr}] 점주메일 재발송완료`);
    
    ui.alert(`✅ [재발송 성공]\n\n• 매장: ${storeNameJp}\n• 수신 메일: ${storeEmail}\n• 플랜: ${mealOption || '기본'}\n\n점주에게 예약 안내 메일이 재전송되었습니다!`);
  } catch (err) {
    logSheet.getRange(targetRow, 25).setValue("❌ 재발송 실패: " + err.toString());
    ui.alert(`❌ 재발송 실패: ${err.toString()}`);
  }
}

/** [9] 📱 [모바일 어드민 겸용] 점주 취소 알림 메일 발송 및 마스터로그 슬롯 반환 엔진 */
function adminCancelAndNotifyStore(row) {
  try {
    const safeRow = parseInt(row, 10);
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const logSheet = ss.getSheetByName('Master_Log');
    const restSheet = ss.getSheetByName('Restaurant_List');

    if (!logSheet || !restSheet) return { success: false, error: "필수 시트 누락" };

    const rowData = logSheet.getRange(safeRow, 1, 1, 25).getValues()[0];
    const orderNo = String(rowData[0] || '').replace(/'/g, '').trim();
    const creatorName = String(rowData[2] || '').trim();
    const storeId = String(rowData[4] || '').trim().toUpperCase();
    const storeNameKo = String(rowData[5] || '').trim();
    const rawVisitDate = rowData[7];
    const peopleStr = String(rowData[8] || '1명').replace(/[^0-9]/g, '') || '1';

    if (!orderNo) return { success: false, error: "주문번호가 없습니다." };

    const timeZone = Session.getScriptTimeZone() || "Asia/Seoul";
    const dateStr = (rawVisitDate instanceof Date) 
      ? Utilities.formatDate(rawVisitDate, timeZone, "yyyy-MM-dd") 
      : String(rawVisitDate || '').substring(0, 10);
    const timeStr = (rawVisitDate instanceof Date) 
      ? Utilities.formatDate(rawVisitDate, timeZone, "HH:mm") 
      : String(rawVisitDate || '').substring(11, 16);
    const visitDisplay = (dateStr && timeStr) ? `${dateStr} ${timeStr}` : (dateStr || "未定 (미정)");

    let storeEmail = "", storeNameJp = "";
    const restData = restSheet.getDataRange().getValues();
    
    // 3행부터 순회 (k=2)
    for (let k = REST_DATA_START_INDEX; k < restData.length; k++) {
      if (String(restData[k][0] || '').trim().toUpperCase() === storeId) {
        storeNameJp = String(restData[k][2] || '').trim() || String(restData[k][1] || '').trim();
        storeEmail = String(restData[k][10] || '').trim();
        break;
      }
    }
    const finalStoreName = storeNameJp || storeNameKo || storeId;

    if (!storeEmail || !storeEmail.includes("@")) {
      return { success: false, error: `매장 이메일이 등록되어 있지 않습니다. (${storeId})` };
    }

    const todayStr = Utilities.formatDate(new Date(), timeZone, "yyyy-MM-dd HH:mm");
    const subject = `【予約キャンセル】 ご来店予約キャンセルのご案内 (${finalStoreName})`;
    
    const htmlBody = `
      <meta charset="UTF-8">
      <div style="font-family: 'Helvetica Neue', Arial, 'Hiragino Kaku Gothic ProN', Meiryo, sans-serif; padding: 20px; background-color: #f4f5f7;">
        <div style="max-width: 500px; margin: 0 auto; background: #ffffff; border-radius: 14px; padding: 30px; box-shadow: 0 4px 12px rgba(0,0,0,0.05); border: 1px solid #edf0f2;">
          
          <div style="text-align: center; margin-bottom: 22px;">
            <span style="font-size: 11px; font-weight: 800; letter-spacing: 1px; color: #ffffff; background: #dc3545; padding: 4px 10px; border-radius: 6px; display: inline-block;">CANCELLATION</span>
            <h1 style="color: #1A2B49; margin: 10px 0 0 0; font-size: 22px; font-weight: 900; letter-spacing: -0.5px;">BOOKMARK CREATORS</h1>
            <div style="width: 36px; height: 3px; background: #dc3545; margin: 10px auto;"></div>
          </div>

          <h2 style="color: #dc3545; margin-top: 0; font-size: 17px; border-bottom: 2px solid #f1f3f5; padding-bottom: 15px; text-align: center; font-weight: 800;">&#10060; ご来店予約キャンセルのご連絡</h2>

          <div style="margin-top: 20px;">
            <p style="color: #1A2B49; font-size: 15.5px; font-weight: bold; margin-bottom: 4px;">${finalStoreName}</p>
            <p style="color: #6c757d; font-size: 13px; margin-top: 0;">店舗管理者様</p>
          </div>

          <p style="color: #495057; font-size: 13.5px; line-height: 1.65; margin: 16px 0;">
            いつもお世話になっております。BOOKMARK CREATORS運営事務局です。<br><br>
            予定されておりました下記の来店予約につきまして、<b>クリエイター側の都合によりキャンセルとなりました</b>のでご案内申し上げます。<br>
            お席の確保やご準備は不要となります。
          </p>

          <div style="background-color: #f8f9fa; border: 1px solid #e9ecef; border-left: 4px solid #dc3545; padding: 16px; border-radius: 8px; margin: 20px 0; font-size: 13.5px;">
            <p style="margin: 5px 0;"><strong>■ 注文番号:</strong> #${orderNo}</p>
            <p style="margin: 5px 0;"><strong>■ クリエ이터名:</strong> ${creatorName}</p>
            <p style="margin: 5px 0;"><strong>■ 予約日時:</strong> <span style="color: #dc3545; font-weight: 800;">${visitDisplay}</span></p>
            <p style="margin: 5px 0;"><strong>■ 予約人数:</strong> ${peopleStr}名</p>
          </div>

          <p style="color: #8b95a1; font-size: 12px; line-height: 1.6; margin: 18px 0 0 0;">
            ※ 本件に関する店舗様側のご対応（確認ボタンの押下等）は不要です。<br>
            ※ 新たな日程で再予約が入りましたら、改めてご案内いたします。
          </p>

          <div style="text-align: center; margin-top: 26px; padding: 12px; background-color: #f1f3f5; border-radius: 8px; font-size: 12px; color: #6c757d; font-weight: 600;">
            BOOKMARK CREATORS 運営事務局 (<a href="mailto:info@bookmarkfukuoka.jp" style="color: #1a73e8; text-decoration: none;">info@bookmarkfukuoka.jp</a>)
          </div>

        </div>
      </div>
    `;

    const adminAlertEmails = (typeof getAdminAlertEmails === 'function') ? getAdminAlertEmails() : "bookmarkjapan.info@gmail.com";

    GmailApp.sendEmail(storeEmail, subject, "", {
      htmlBody: htmlBody,
      name: "BOOKMARK CREATORS",
      from: "info@bookmarkfukuoka.jp",
      bcc: adminAlertEmails
    });

    // 🎯 [정합성 보정] L열: 취소(운영자), H~J열 슬롯 비움, U열(21번째): 취소환불
    logSheet.getRange(safeRow, 12).setValue("취소(운영자)");
    logSheet.getRange(safeRow, 8, 1, 3).clearContent();
    logSheet.getRange(safeRow, 21).setValue("취소환불");

    const currentLog = String(logSheet.getRange(safeRow, 25).getValue() || '').trim();
    const newLog = `[${todayStr}] 점주 취소메일 발송완료(어드민)`;
    logSheet.getRange(safeRow, 25).setValue(currentLog ? `${currentLog} | ${newLog}` : newLog);

    return { success: true };
  } catch (err) {
    return { success: false, error: err.toString() };
  }
}

/**
 * ❌ [시트 메뉴용] 마우스로 선택한 행 자동 감지 + 취소 통지 메일 발송
 */
function sendStoreCancellationEmail() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const activeSheet = ss.getActiveSheet();

  // 1. 현재 시트가 Master_Log인지 확인
  if (activeSheet.getName() !== "Master_Log") {
    return ui.alert("⚠️ 'Master_Log' 시트로 이동한 후 취소할 행을 클릭하고 실행해주세요.");
  }

  // 2. 현재 마우스 커서가 올라가 있는 행 번호 취득
  const targetRow = activeSheet.getActiveRange().getRow();

  // 3. 헤더 행(1, 2행) 클릭 방어
  if (targetRow < 3) {
    return ui.alert("⚠️ 취소하려는 예약 데이터 행(3행 이후)을 마우스로 클릭한 뒤 실행해주세요.");
  }

  const orderNo = String(activeSheet.getRange(targetRow, 1).getValue() || '').replace(/'/g, '').trim();
  const creatorName = String(activeSheet.getRange(targetRow, 3).getValue() || '').trim();
  const storeName = String(activeSheet.getRange(targetRow, 6).getValue() || '').trim();
  const currentStatus = String(activeSheet.getRange(targetRow, 12).getValue() || '').trim();

  if (!orderNo) {
    return ui.alert(`⚠️ 선택하신 ${targetRow}행에는 주문 데이터가 없습니다.`);
  }

  // 4. 클릭한 대상 정보 확인창 띄우기 (실수 방지)
  const confirm = ui.alert(
    "⚠️ 점주 취소 통지 확인",
    `[선택된 예약 정보]\n` +
    `• 행 번호: ${targetRow}행\n` +
    `• 주문번호: #${orderNo}\n` +
    `• 크리에이터: ${creatorName}\n` +
    `• 매장명: ${storeName}\n` +
    `• 현재상태: ${currentStatus}\n\n` +
    `점주에게 예약 취소 안내 메일을 발송하고 상태를 '취소(운영자)'로 변경하시겠습니까?`,
    ui.ButtonSet.YES_NO
  );

  if (confirm !== ui.Button.YES) return;

  const result = adminCancelAndNotifyStore(targetRow);

  if (result.success) {
    ui.alert(`✅ [취소 완료]\n\n#${orderNo} (${storeName}) 예약 취소 메일이 점주에게 발송되었습니다.`);
  } else {
    ui.alert(`❌ [처리 실패]\n\n사유: ${result.error}`);
  }
}
