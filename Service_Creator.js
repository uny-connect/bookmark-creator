/** [3] 미션 목록 조회 (A~U열 변경 구조 반영 - 안정화 버전) */
function getMissions(orderNo, phoneLast4) {
  try {
    const data = _getSheetsData(['Master_Log', 'User_DB', 'Restaurant_List']);
    const inputOrderNo = String(orderNo).trim().replace(/'/g, '');
    const inputPhone = String(phoneLast4).trim();

    const userMap = new Map(); const userTierMap = new Map(); 
    for (let j = 1; j < data.User_DB.length; j++) { // 🚨 인덱스 1(2번째 행)부터 안전하게 스캔
      const uRow = data.User_DB[j]; if (!uRow || !uRow[0]) continue;
      const mCode = String(uRow[0]).trim();
      // uRow[3]이 전화번호 열이 맞는지 반드시 시트 확인 필요!
      const phoneStr = uRow[3] ? String(uRow[3]).replace(/[^0-9]/g, '') : '';
      userMap.set(mCode, phoneStr);
      userTierMap.set(mCode, parseTierEmoji(uRow[8])); 
    }
    
// 매장 ID 맵과 매장 이름 맵을 동시에 구축 (ID가 없을 때 이름으로 찾기 위함)
    const restMap = new Map();
    const restNameMap = new Map();
    for (let k = 1; k < data.Restaurant_List.length; k++) {
      const rRow = data.Restaurant_List[k]; if (!rRow || !rRow[0]) continue;
      const rId = String(rRow[0]).trim().toUpperCase();
      const rName = String(rRow[1] || '').trim();
      const storeType = String(rRow[7] || '').trim().toUpperCase();
      
      // 🎯 V열(22번째 열, 인덱스 21)에서 최대인원 파싱 (미입력 또는 오류 시 기본값 4)
      const maxPeople = parseInt(rRow[21], 10) || 4;

      const rDetails = { 
        map: String(rRow[17] || '#').trim(),
        guide: String(rRow[18] || '#').trim(),
        bookingUrl: (storeType === 'RETAIL') ? '' : String(rRow[20] || '').trim(), 
        storeType: storeType, 
        maxPeople: maxPeople, // 🎯 추가됨
        blackouts: getSafeBlackouts(rRow[13])
      };
      
      if (rId) restMap.set(rId, rDetails);
      if (rName) restNameMap.set(rName, rDetails);
    }

    const mList = []; const timeZone = Session.getScriptTimeZone();

    for (let i = 1; i < data.Master_Log.length; i++) {
      const mRow = data.Master_Log[i]; if (!mRow || !mRow[0]) continue;
      
      const status = String(mRow[11] || '').trim(); 
      if (status.includes('취소') || String(mRow[0]).replace(/'/g, '').trim() !== inputOrderNo) continue; 

      const memberCode = String(mRow[1]).trim();
      const userPhone = userMap.get(memberCode) || '';
      if (!userPhone.endsWith(inputPhone)) continue; // 폰 번호 뒷자리 대조

      const restId = String(mRow[4] || '').trim().toUpperCase();
      const restaurantName = String(mRow[5] || '').trim();
      
      // 🚨 [방어막] ID로 못 찾으면 매장 이름으로 한 번 더 가이드/맵 링크 매칭 시도
      let restInfo = restMap.get(restId);
      if (!restInfo && restaurantName) {
        restInfo = restNameMap.get(restaurantName);
      }
      // 둘 다 없으면 기본값 유지 (maxPeople: 4 추가 방어)
      if (!restInfo) {
        restInfo = { map: '#', guide: '#', bookingUrl: '', storeType: '', maxPeople: 4, blackouts: [] };
      }
      
      const formatDate = (val, format) => (val instanceof Date) ? Utilities.formatDate(val, timeZone, format) : String(val || '');
      const deadline = formatDate(mRow[9], 'yyyy-MM-dd');
      let vDate = formatDate(mRow[7], 'yyyy-MM-dd HH:mm');
      if (vDate === '-' || vDate === '') vDate = '';

      mList.push({
        row: i + 1, orderNo: inputOrderNo, restId, memberCode,
        member: String(mRow[2] || ''), restaurant: restaurantName, status,
        captureUrl: String(mRow[6] || ''), 
        verified: String(mRow[13] || ''),  
        receiptUrl: String(mRow[14] || ''), 
        submitDate: formatDate(mRow[15], 'yyyy-MM-dd'), 
        reviewUrl: String(mRow[16] || ''),  
        googleMapUrl: String(mRow[17] || ''), 
        shortReview: String(mRow[18] || ''), 
        deadline, visitDate: vDate, tierEmoji: userTierMap.get(memberCode) || '🟡',
        info: { name: restaurantName, ...restInfo }
      });
    }
    return mList;
  } catch (e) { 
    // 🚨 에러가 나면 숨기지 않고 로그를 찍어 원인을 찾을 수 있게 양보
    Logger.log('getMissions 에러 발생: ' + e.toString());
    throw new Error('데이터 통신 오류: ' + e.toString()); 
  }
}

/** 캡처본 업로드 */
function uploadCapture(row, base64Data, filename, orderNo, memberName) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Master_Log');
    const folder = DriveApp.getFoldersByName("bookmark_Captures").hasNext() ? DriveApp.getFoldersByName("bookmark_Captures").next() : DriveApp.createFolder("bookmark_Captures");
    const splitData = base64Data.split(',');
    const blob = Utilities.newBlob(Utilities.base64Decode(splitData[1]), splitData[0].split(':')[1].split(';')[0], `${orderNo}_${memberName}_${filename}`);
    const file = folder.createFile(blob).setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    sheet.getRange(parseInt(row, 10), 7).setValue(file.getUrl()); 
    sheet.getRange(parseInt(row, 10), 12).setValue('예약확인중'); 
    return { success: true };
  } catch (e) { return { success: false, error: e.toString() }; }
}

/** 현장 방문 확정 클릭 버튼 리스너 */
/** [방문 확인 기능] 크리에이터가 현장에서 버튼을 클릭했을 때 실행 */
function confirmVisit(row) {
  try {
    const safeRow = parseInt(row, 10); 
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Master_Log');
    if (String(sheet.getRange(safeRow, 12).getValue()).trim() !== '방문전') {
      return { success: false, error: "既に訪問確認済みか、対象外のステータスです。" };
    } 
    sheet.getRange(safeRow, 12).setValue('방문완료'); 
    sheet.getRange(safeRow, 14).setValue('Y'); // N열방문확인
    return { success: true };
  } catch (e) { return { success: false, error: e.toString() }; }
}

/** 최종 리뷰 콘텐츠 및 영수증 제출 */
function submitContent(row, reviewUrl, receiptBase64, receiptName, googleMapUrl, shortReview) {
  try {
    const safeRow = parseInt(row, 10); 
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Master_Log');
    
    if (receiptBase64) {
      const folder = DriveApp.getFoldersByName("bookmark_Receipts").hasNext() ? DriveApp.getFoldersByName("bookmark_Receipts").next() : DriveApp.createFolder("bookmark_Receipts");
      const finalFileName = `${String(sheet.getRange(safeRow, 1).getValue()).trim()}_${String(sheet.getRange(safeRow, 5).getValue()).trim()}_${receiptName}`;
      const splitData = receiptBase64.split(',');
      const file = folder.createFile(Utilities.newBlob(Utilities.base64Decode(splitData[1]), splitData[0].split(':')[1].split(';')[0], finalFileName));
      file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW); 
      sheet.getRange(safeRow, 15).setValue(file.getUrl()); // O열 매장 영수증 적재 [15번째 열]
    }
    
    const fixUrl = (url) => { const u = String(url || '').trim(); return (u && !/^https?:\/\//i.test(u)) ? 'https://' + u : u; };
    sheet.getRange(safeRow, 16).setValue(Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd")); // P열 제출일 [16]
    sheet.getRange(safeRow, 17).setValue(fixUrl(reviewUrl)); // Q열 콘텐츠 링크 [17]
    if (googleMapUrl) sheet.getRange(safeRow, 18).setValue(fixUrl(googleMapUrl)); // R열 구글맵 [18]
    
    const safeReview = String(shortReview || '').trim();
    if (safeReview) {
      sheet.getRange(safeRow, 19).setValue(safeReview); // S열 한 줄 후기 [19]
      try { sheet.getRange(safeRow, 20).setValue(LanguageApp.translate(safeReview, 'ko', 'ja')); } catch(txtErr) {} // T열 JP후기 [20]
    }
    sheet.getRange(safeRow, 12).setValue('제출완료'); // L열 상태변경
    sheet.getRange(safeRow, 21).setValue('환급대기'); // U열 보증금 세팅
    return { success: true };
  } catch (e) { return { success: false, error: e.toString() }; }
}

/** 수정 제출 처리 */
function updateSubmission(row, newUrl, newReceiptBase64, newReceiptName, newGoogleUrl, newShortReview) {
  try {
    const safeRow = parseInt(row, 10); 
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Master_Log');
    const fixUrl = (url) => { const u = String(url || '').trim(); return (u && !/^https?:\/\//i.test(u)) ? 'https://' + u : u; };

    sheet.getRange(safeRow, 17).setValue(fixUrl(newUrl)); // Q열 [17]
    sheet.getRange(safeRow, 18).setValue(fixUrl(newGoogleUrl)); // R열 [18]
    if(newShortReview !== undefined) sheet.getRange(safeRow, 19).setValue(String(newShortReview).trim()); // S열 [19]
    
    if (newReceiptBase64) {
      const folder = DriveApp.getFoldersByName("bookmark_Receipts").hasNext() ? DriveApp.getFoldersByName("bookmark_Receipts").next() : DriveApp.createFolder("bookmark_Receipts");
      const splitData = newReceiptBase64.split(',');
      const file = folder.createFile(Utilities.newBlob(Utilities.base64Decode(splitData[1]), splitData[0].split(':')[1].split(';')[0], "temp"));
      file.setName(`${String(sheet.getRange(safeRow, 1).getValue()).trim()}_${String(sheet.getRange(safeRow, 5).getValue()).trim()}_Edit_${newReceiptName}`);
      file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW); 
      sheet.getRange(safeRow, 15).setValue(file.getUrl()); // O열 [15]
    }
    return { success: true };
  } catch (e) { return { success: false, error: e.toString() }; }
}

/* ⚡ [속도 대폭 튜닝] 시트 스캔 횟수를 단 1회로 줄인 고속 슬롯 필터 제너레이터 */
function checkAvailability(storeId, targetDateStr) {
  try {
    const data = _getSheetsData(['Restaurant_List', 'Master_Log']);
    let config = null;

    const restList = data.Restaurant_List;
    for (let i = 2; i < restList.length; i++) {
      if (String(restList[i][0]).trim() === storeId) {
        config = {
          maxPerSlot: parseInt(restList[i][11], 10) || 1,
          dailyCap: parseInt(restList[i][12], 10) || 999,
          blackouts: getSafeBlackouts(restList[i][13]),
          bizHours: String(restList[i][14] || '').trim(),
          intervalMin: parseInt(restList[i][15], 10) || 30,
          blockedTimes: String(restList[i][16] || '').split(',').map(s => s.trim()).filter(Boolean)
        };
        break;
      }
    }

    if (!config) return { error: "오류: 매장 확인 불가" };
    if (!config.bizHours) return { error: "오류: 매장 영업시간 설정 미비" };

    const targetDateObj = new Date(targetDateStr);
    const tDayJa = ['日', '月', '火', '水', '木', '金', '土'][targetDateObj.getDay()]; 
    const tWeekNum = Math.ceil(targetDateObj.getDate() / 7); 

    const isBlackout = config.blackouts.some(b => {
      if (!b) return false;
      if (b === targetDateStr || b === tDayJa) return true;
      if (b.includes('第')) {
        const match = b.match(/第([1-5])/);
        return match && parseInt(match[1], 10) === tWeekNum && b.includes(tDayJa);
      }
      return false;
    });

    if (isBlackout) return { isAvailable: false, reason: "정기 휴무일", availableSlots: [] };

    let dailyTotal = 0; 
    let timeSlotCounts = {};
    const timeZone = Session.getScriptTimeZone();

    // 🚨 [핵심 변경] getValue() 추방 ❯ 전체 2차원 메모리 배열 내부 스캔으로 전면 대체
    const masterLog = data.Master_Log;
    for (let i = 2; i < masterLog.length; i++) {
      const mRow = masterLog[i];
      if (String(mRow[4]).trim() === storeId && !String(mRow[11]).includes('취소') && !String(mRow[11]).includes('노쇼')) {
        let v = mRow[7]; // 메모리에 이미 들어와 있는 컬럼 데이터 직접 참조
        let dateStr = (v instanceof Date) ? Utilities.formatDate(v, timeZone, 'yyyy-MM-dd') : String(v || '').substring(0, 10);
        let timeStr = (v instanceof Date) ? Utilities.formatDate(v, timeZone, 'HH:mm') : String(v || '').substring(11, 16).replace(/\s+/g, '');
        
        if (dateStr === targetDateStr) {
          dailyTotal++;
          if (timeStr) timeSlotCounts[timeStr] = (timeSlotCounts[timeStr] || 0) + 1;
        }
      }
    }

    if (dailyTotal >= config.dailyCap) return { isAvailable: false, reason: "당일 체험 마감", availableSlots: [] };

    let allSlots = [];
    try {
      const parseTime = (t) => { const p = t.split(':'); return parseInt(p[0], 10) * 60 + parseInt(p[1], 10); };
      const times = config.bizHours.split('-');
      const startMin = parseTime(times[0].trim());
      let endMin = parseTime(times[1].trim());
      if (endMin < startMin) endMin += 1440; 
      
      const blocks = config.blockedTimes.map(b => {
        if (!b.includes('-')) return { type: 'exact', time: parseTime(b) };
        const bp = b.split('-');
        let bs = parseTime(bp[0].trim()), be = parseTime(bp[1].trim());
        if (endMin < bs) be += 1440;
        return { type: 'range', start: bs, end: be };
      });

      for (let curr = startMin; curr < endMin; curr += config.intervalMin) {
        const hit = blocks.some(b => b.type === 'exact' ? curr === b.time : (curr >= b.start && curr <= b.end));
        if (!hit) {
          const h = Math.floor(curr / 60) % 24, m = curr % 60;
          allSlots.push((h < 10 ? '0' + h : h) + ':' + (m < 10 ? '0' + m : m));
        }
      }
    } catch(e) { return { error: "営業時間の入力形式が正しくありません。(例: 11:00-22:00)" }; }

    const availableSlots = allSlots.filter(time => (timeSlotCounts[time] || 0) < config.maxPerSlot);
    return availableSlots.length ? { isAvailable: true, availableSlots } : { isAvailable: false, reason: "すべての枠が埋まっているか、予約不可の日です。", availableSlots: [] };
  } catch (e) { return { error: e.toString() }; }
}

/** 확정 슬롯 타임 데이터베이스 픽싱 및 점주 메일 인터랙션 노티 */
function bookTimeSlot(row, dateStr, timeStr, peopleCount) {
  try {
    const safeRow = parseInt(row, 10);
    const safePeopleCount = parseInt(peopleCount, 10) || 1; 
    const data = _getSheetsData(['Restaurant_List', 'User_DB', 'Master_Log']);
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Master_Log');
    
    // 🔒 [실시간 중복 예약 원천 차단 락]
    const realTimeStatus = String(sheet.getRange(safeRow, 12).getValue()).trim();
    if (realTimeStatus !== '예약대기' && realTimeStatus !== '일정조율필요') {
      return { 
        success: false, 
        error: `이미 처리 중이거나 확정된 예약입니다. (현재 상태: ${realTimeStatus})\n새로고침 후 다시 확인해 주세요.` 
      };
    }
    
    const visitDateTime = dateStr + " " + timeStr;
    sheet.getRange(safeRow, 8).setValue(visitDateTime); 
    sheet.getRange(safeRow, 9).setValue(safePeopleCount + '명'); 
    if (!sheet.getRange(safeRow, 11).getValue()) sheet.getRange(safeRow, 11).setValue(50000); 
    
    // 상태 즉시 변경하여 중복 진입 차단
    sheet.getRange(safeRow, 12).setValue('예약확인중'); 
    
    const dDate = new Date(visitDateTime.replace(' ', 'T'));
    dDate.setDate(dDate.getDate() + 10);
    sheet.getRange(safeRow, 10).setValue(dDate); 
    
    const storeId = String(sheet.getRange(safeRow, 5).getValue()).trim(); 
    const memberCode = String(sheet.getRange(safeRow, 2).getValue()).trim(); 
    const memberName = String(sheet.getRange(safeRow, 3).getValue()).trim(); 
    
    let storeEmail = "", storeNameJp = "";
    for (let i = 2; i < data.Restaurant_List.length; i++) {
      if (String(data.Restaurant_List[i][0]).trim() === storeId) {
        storeNameJp = String(data.Restaurant_List[i][2]).trim() || String(data.Restaurant_List[i][1]).trim();
        storeEmail = String(data.Restaurant_List[i][10]).trim(); 
        break;
      }
    }

    let creatorProfileUrl = "";
    for (let u = 1; u < data.User_DB.length; u++) {
      if (String(data.User_DB[u][0]).trim() === memberCode) {
        creatorProfileUrl = String(data.User_DB[u][4] || '').trim(); 
        break;
      }
    }

    const profileHtml = (creatorProfileUrl.length > 5) 
      ? `<p style="margin: 5px 0; font-size: 15px;"><strong>&#128279; <span>SNS:</span></strong> <a href="${creatorProfileUrl}" target="_blank" style="color: #1a73e8; font-weight: bold; text-decoration: underline;"><span>SNSを見る❯</span></a></p>`
      : `<p style="margin: 5px 0; font-size: 15px; color: #8b95a1;"><strong>&#128279; <span>SNS:</span></strong> <span>当日確認</span></p>`;
    
    // 🎯 [대안 1 적용] 이메일 발송 트라이캐치 격리 및 Y열(25번째 열) 실시간 추적 로그 작성
    if (storeEmail && storeEmail.includes("@")) {
      const scriptUrl = ScriptApp.getService().getUrl();
      const orderNo = String(sheet.getRange(safeRow, 1).getValue()).trim().replace(/'/g, '');
      
      const confirmUrl = `${scriptUrl}?mode=store_confirm&row=${safeRow}&o=${encodeURIComponent(orderNo)}`;
      const feedbackUrl = `${scriptUrl}?mode=feedback&row=${safeRow}&o=${encodeURIComponent(orderNo)}`;

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
              <p style="margin: 5px 0; font-size: 15px;"><strong>&#128101; 訪問人数:</strong> <span style="color: #1A2B49; font-weight: bold;">${safePeopleCount}名</span></p>
              ${profileHtml}
            </div>
            <div style="margin: 30px 0; text-align: center;">
              <a href="${confirmUrl}" target="_blank" style="background-color: #2D6A4F; color: #ffffff; padding: 14px 20px; font-size: 14px; font-weight: bold; text-decoration: none; border-radius: 10px; display: inline-block; margin-right: 10px; box-shadow: 0 4px 12px rgba(45,106,79,0.2);">&#9989; 予約を確定する</a>
              <a href="${feedbackUrl}" target="_blank" style="background-color: #1A2B49; color: #ffffff; padding: 14px 20px; font-size: 14px; font-weight: bold; text-decoration: none; border-radius: 10px; display: inline-block; box-shadow: 0 4px 12px rgba(26,43,73,0.15);">&#128260; 日時変更をリクエスト</a>
            </div>
          </div>
        </div>`;
      
      try {
        GmailApp.sendEmail(storeEmail, subject, "", { htmlBody: htmlBody, name: "BOOKMARK CREATORS" });
        sheet.getRange(safeRow, 25).setValue("점주메일 발송완료"); // Y열[25]에 기록
      } catch (mailErr) {
        console.error("점주 이메일 슈팅 실패: " + mailErr.toString());
        sheet.getRange(safeRow, 25).setValue("❌ 점주메일 실패: " + mailErr.toString()); // 에러 내용 시트에 바인딩
      }
    } else {
      sheet.getRange(safeRow, 25).setValue("❌ 실패: 점주 이메일 주소 없음");
    }
    return { success: true };
  } catch (e) { return { success: false, error: e.toString() }; }
}
