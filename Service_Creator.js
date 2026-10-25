/** [3] 미션 목록 조회 (승인 채널 목록 approvedChannels, minDays/maxDays, G열 구매옵션 연동) */
function getMissions(orderNo, phoneLast4) {
  try {
    const data = _getSheetsData(['Master_Log', 'User_DB', 'Restaurant_List']);
    const inputOrderNo = String(orderNo || '').trim().replace(/['"\s]/g, '');
    const inputPhone = String(phoneLast4 || '').trim().replace(/[^0-9]/g, '');

    const userMap = new Map(); 
    const userTierMap = new Map(); 
    const userChannelsMap = new Map();

    for (let j = 1; j < data.User_DB.length; j++) {
      const uRow = data.User_DB[j]; 
      if (!uRow || !uRow[0]) continue; 
      
      const mCode = String(uRow[0]).trim();
      const mEmail = String(uRow[13] || '').trim();
      const phoneStr = uRow[3] ? String(uRow[3]).replace(/[^0-9]/g, '') : '';
      
      if (mCode) userMap.set(mCode.toLowerCase(), phoneStr);
      if (mEmail) userMap.set(mEmail.toLowerCase(), phoneStr);
      userTierMap.set(mCode.toLowerCase(), parseTierEmoji(uRow[9])); 

      const channels = [];
      if (String(uRow[4] || '').trim().length > 3) channels.push({ type: 'blog', name: '네이버 블로그', icon: '📝', placeholder: '블로그 리뷰 포스팅 URL' });
      if (String(uRow[5] || '').trim().length > 3) channels.push({ type: 'instagram', name: '인스타그램(릴스)', icon: '📸', placeholder: '인스타그램 릴스/게시물 URL' });
      if (String(uRow[6] || '').trim().length > 3) channels.push({ type: 'youtube', name: '유튜브', icon: '🎥', placeholder: '유튜브 영상/쇼츠 URL' });
      if (String(uRow[7] || '').trim().length > 3) channels.push({ type: 'tiktok', name: '틱톡', icon: '🎵', placeholder: '틱톡 영상 URL' });
      if (String(uRow[8] || '').trim().length > 3) channels.push({ type: 'google', name: '구글 로컬 가이드', icon: '🗺️', placeholder: '로컬 가이드 리뷰 링크' });

      if (channels.length === 0) {
        channels.push({ type: 'instagram', name: '인스타그램(릴스)', icon: '📸', placeholder: '인스타그램 릴스/게시물 URL' });
      }

      if (mCode) userChannelsMap.set(mCode.toLowerCase(), channels);
      if (mEmail) userChannelsMap.set(mEmail.toLowerCase(), channels);
    }
    
    const restMap = new Map();
    const restNameMap = new Map();
    for (let k = 1; k < data.Restaurant_List.length; k++) {
      const rRow = data.Restaurant_List[k]; 
      if (!rRow || !rRow[0]) continue; 
      const rId = String(rRow[0]).trim().toUpperCase();
      const rName = String(rRow[1] || '').trim();
      const storeType = String(rRow[7] || '').trim().toUpperCase();
      const maxPeople = parseInt(rRow[21], 10) || 4;

      const minLeadDays = parseInt(rRow[23], 10) || 1;
      const maxLeadDays = parseInt(rRow[24], 10) || 90;

      const rDetails = { 
        map: String(rRow[17] || '#').trim(),
        guide: String(rRow[18] || '#').trim(),
        bookingUrl: (storeType === 'RETAIL') ? '' : String(rRow[20] || '').trim(), 
        storeType: storeType, 
        maxPeople: maxPeople,
        minDays: minLeadDays,
        maxDays: maxLeadDays,
        blackouts: getSafeBlackouts(rRow[13]),
        benefit: String(rRow[22] || '').trim() // 🎯 [추가] W열(23번째 열, index 22) 제공 메뉴/혜택
      };
      
      if (rId) restMap.set(rId, rDetails);
      if (rName) restNameMap.set(rName, rDetails);
    }

    const mList = []; 
    const timeZone = Session.getScriptTimeZone();

    for (let i = 1; i < data.Master_Log.length; i++) {
      const mRow = data.Master_Log[i]; 
      if (!mRow || !mRow[0]) continue; 
      
      const status = String(mRow[11] || '').trim(); 
      const sheetOrderNo = String(mRow[0]).replace(/['"\s]/g, '').trim(); 
      
      if (status.includes('취소') || sheetOrderNo !== inputOrderNo) continue; 

      const memberCode = String(mRow[1] || '').trim().toLowerCase();
      const userPhone = userMap.get(memberCode) || '';
      
      if (!userPhone.endsWith(inputPhone)) continue;

      const restId = String(mRow[4] || '').trim().toUpperCase();
      const restaurantName = String(mRow[5] || '').trim();
      
      let restInfo = restMap.get(restId);
      if (!restInfo && restaurantName) {
        restInfo = restNameMap.get(restaurantName);
      }
      if (!restInfo) {
        restInfo = { map: '#', guide: '#', bookingUrl: '', storeType: '', maxPeople: 4, minDays: 1, maxDays: 90, blackouts: [] };
      }
      
      const formatDate = (val, format) => (val instanceof Date) ? Utilities.formatDate(val, timeZone, format) : String(val || '');
      const deadline = formatDate(mRow[9], 'yyyy-MM-dd');
      let vDate = formatDate(mRow[7], 'yyyy-MM-dd HH:mm');
      if (vDate === '-' || vDate === '') vDate = '';

      mList.push({
        row: i + 1, orderNo: inputOrderNo, restId, memberCode,
        member: String(mRow[2] || ''), restaurant: restaurantName, status,
        mealOption: String(mRow[6] || '').trim(), // 🎯 G열 구매옵션 (인덱스 6)
        captureUrl: String(mRow[6] || ''), 
        verified: String(mRow[13] || ''),  
        receiptUrl: String(mRow[14] || ''), 
        submitDate: formatDate(mRow[15], 'yyyy-MM-dd'), 
        reviewUrl: String(mRow[16] || ''),  
        googleMapUrl: String(mRow[17] || ''), 
        shortReview: String(mRow[18] || ''), 
        deadline, visitDate: vDate, tierEmoji: userTierMap.get(memberCode) || '🟡',
        approvedChannels: userChannelsMap.get(memberCode) || [{ type: 'instagram', name: '인스타그램(릴스)', icon: '📸', placeholder: '인스타그램 릴스/게시물 URL' }],
        info: { name: restaurantName, ...restInfo }
      });
    }
    return mList;
  } catch (e) { 
    Logger.log('getMissions 에러 발생: ' + e.toString());
    throw new Error('데이터 통신 오류: ' + e.toString()); 
  }
}

/** 현장 방문 확정 클릭 버튼 리스너 */
function confirmVisit(row) {
  try {
    const safeRow = parseInt(row, 10); 
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Master_Log');
    if (String(sheet.getRange(safeRow, 12).getValue()).trim() !== '방문전') {
      return { success: false, error: "既に訪問確認済みか、対象外のステータスです。" };
    } 
    sheet.getRange(safeRow, 12).setValue('방문완료'); 
    sheet.getRange(safeRow, 14).setValue('Y');
    return { success: true };
  } catch (e) { return { success: false, error: e.toString() }; }
}

/** 최종 리뷰 콘텐츠 제출 */
function submitContent(row, reviewUrl, googleMapUrl, shortReview, channelType) {
  try {
    const safeRow = parseInt(row, 10); 
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Master_Log');
    const fixUrl = (url) => { 
      const u = String(url || '').trim(); 
      return (u && !/^https?:\/\//i.test(u)) ? 'https://' + u : u; 
    };

    const cleanReviewUrl = fixUrl(reviewUrl);
    const cleanGoogleUrl = fixUrl(googleMapUrl);
    const lowerRev = cleanReviewUrl.toLowerCase();
    const lowerGoo = cleanGoogleUrl.toLowerCase();

    if (channelType === 'instagram' && !lowerRev.includes('instagram.com')) {
      return { success: false, error: "인스타그램 리뷰 링크는 instagram.com 도메인이어야 합니다." };
    }
    if (channelType === 'blog' && !lowerRev.includes('naver.com')) {
      return { success: false, error: "블로그 리뷰 링크는 naver.com 도메인이어야 합니다." };
    }
    if (cleanGoogleUrl && !lowerGoo.includes('google.') && !lowerGoo.includes('goo.gl')) {
      return { success: false, error: "구글맵 리뷰 링크 형식이 올바르지 않습니다." };
    }
    if (cleanReviewUrl === cleanGoogleUrl) {
      return { success: false, error: "콘텐츠 링크와 구글맵 리뷰 링크가 동일합니다." };
    }

    sheet.getRange(safeRow, 16).setValue(Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd"));
    sheet.getRange(safeRow, 17).setValue(cleanReviewUrl);
    if (cleanGoogleUrl) sheet.getRange(safeRow, 18).setValue(cleanGoogleUrl); 
    
    const safeReview = String(shortReview || '').trim();
    if (safeReview) {
      sheet.getRange(safeRow, 19).setValue(safeReview);
      try { sheet.getRange(safeRow, 20).setValue(LanguageApp.translate(safeReview, 'ko', 'ja')); } catch(txtErr) {}
    }
    sheet.getRange(safeRow, 12).setValue('제출완료');
    sheet.getRange(safeRow, 21).setValue('환급대기');
    return { success: true };
  } catch (e) { return { success: false, error: e.toString() }; }
}

/** 수정 제출 처리 */
function updateSubmission(row, newUrl, newGoogleUrl, newShortReview) {
  try {
    const safeRow = parseInt(row, 10); 
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Master_Log');
    const fixUrl = (url) => { const u = String(url || '').trim(); return (u && !/^https?:\/\//i.test(u)) ? 'https://' + u : u; };

    sheet.getRange(safeRow, 17).setValue(fixUrl(newUrl));
    sheet.getRange(safeRow, 18).setValue(fixUrl(newGoogleUrl));
    if(newShortReview !== undefined) sheet.getRange(safeRow, 19).setValue(String(newShortReview).trim());
    return { success: true };
  } catch (e) { return { success: false, error: e.toString() }; }
}

/* 🎯 [동적 키워드 매칭] 고속 슬롯 필터 제너레이터 */
function checkAvailability(storeId, targetDateStr, mealOption) {
  try {
    const data = _getSheetsData(['Restaurant_List', 'Master_Log']);
    let config = null;

    const cleanInputId = String(storeId || '').trim().toUpperCase().replace(/['"]/g, '');
    const userOption = String(mealOption || '').trim().toLowerCase();
    const restList = data.Restaurant_List;

    for (let i = 1; i < restList.length; i++) {
      const rowId = String(restList[i][0] || '').trim().toUpperCase().replace(/['"]/g, '');
      const rowNameKo = String(restList[i][1] || '').trim();
      const rowNameJp = String(restList[i][2] || '').trim();

      if (rowId === cleanInputId || (cleanInputId && (cleanInputId === rowNameKo || cleanInputId === rowNameJp))) {
        // 1. 기본 운영시간 (O열: 인덱스 14)
        let rawBiz = toHalfWidth(restList[i][14] || '');

        // 2. 동적 옵션 시간대 1 (AA: 매칭키워드[인덱스 26], AB: 운영시간[인덱스 27])
        const opt1Keywords = String(restList[i][26] || '').trim().toLowerCase();
        const opt1Hours    = String(restList[i][27] || '').replace(/[~～－—]/g, '-').trim();

        // 3. 동적 옵션 시간대 2 (AC: 매칭키워드[인덱스 28], AD: 운영시간[인덱스 29])
        const opt2Keywords = String(restList[i][28] || '').trim().toLowerCase();
        const opt2Hours    = String(restList[i][29] || '').replace(/[~～－—]/g, '-').trim();

        // 🎯 [핵심] 유저 옵션 문자열이 시트 키워드 목록 중 하나라도 포함하는지 대조
        if (userOption) {
          const match1 = opt1Keywords && opt1Keywords.split(',').some(k => k.trim() && userOption.includes(k.trim()));
          const match2 = opt2Keywords && opt2Keywords.split(',').some(k => k.trim() && userOption.includes(k.trim()));

          if (match1 && opt1Hours) {
            rawBiz = opt1Hours;
          } else if (match2 && opt2Hours) {
            rawBiz = opt2Hours;
          }
        }

        config = {
          maxPerSlot: parseInt(restList[i][11], 10) || 1,
          rawDailyCap: String(restList[i][12] || '').trim(),
          blackouts: getSafeBlackouts(restList[i][13]),
          bizHours: rawBiz,
          intervalMin: parseInt(restList[i][15], 10) || 30,
          blockedTimes: String(restList[i][16] || '').split(',').map(s => s.trim()).filter(Boolean)
        };
        break;
      }
    }

    if (!config) return { error: "오류: 매장 확인 불가" };
    if (!config.bizHours) return { error: "오류: 매장 영업시간 설정 미비" };

    const targetDateObj = new Date(targetDateStr);
    const dayJaMap = ['日', '月', '火', '水', '木', '金', '土'];
    const dayKoMap = ['일', '월', '화', '수', '목', '금', '토'];
    const tDayJa = dayJaMap[targetDateObj.getDay()];
    const tDayKo = dayKoMap[targetDateObj.getDay()];
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

    let effectiveDailyCap = 999;
    const rawCap = config.rawDailyCap;

    if (!rawCap || /^\d+$/.test(rawCap)) {
      effectiveDailyCap = rawCap ? parseInt(rawCap, 10) : 999;
    } else {
      const capRules = rawCap.split(/[,/]/).map(s => s.trim());
      let defaultCap = 999;
      let specificCap = null;

      for (let rule of capRules) {
        if (!rule) continue;
        if (/^\d+$/.test(rule)) {
          defaultCap = parseInt(rule, 10);
          continue;
        }

        const parts = rule.split(/[:=]/).map(s => s.trim());
        if (parts.length === 2) {
          const key = parts[0];
          const val = parseInt(parts[1], 10);

          if (/^(기본|default|全体|基本)$/i.test(key)) {
            defaultCap = isNaN(val) ? 999 : val;
          } else if (key.includes(tDayKo) || key.includes(tDayJa)) {
            specificCap = isNaN(val) ? 0 : val;
          }
        }
      }
      effectiveDailyCap = (specificCap !== null) ? specificCap : defaultCap;
    }

    if (effectiveDailyCap <= 0) {
      return { isAvailable: false, reason: "당일 체험 마감", availableSlots: [] };
    }

    let dailyTotal = 0;
    let timeSlotCounts = {};
    const timeZone = Session.getScriptTimeZone();
    const masterLog = data.Master_Log;

    for (let i = 2; i < masterLog.length; i++) {
      const mRow = masterLog[i];
      const mStoreId = String(mRow[4] || '').trim().toUpperCase();
      if (mStoreId === cleanInputId && !String(mRow[11]).includes('취소') && !String(mRow[11]).includes('노쇼')) {
        let v = mRow[7];
        let dateStr = (v instanceof Date) ? Utilities.formatDate(v, timeZone, 'yyyy-MM-dd') : String(v || '').substring(0, 10);
        let timeStr = (v instanceof Date) ? Utilities.formatDate(v, timeZone, 'HH:mm') : String(v || '').substring(11, 16).replace(/\s+/g, '');
        
        if (dateStr === targetDateStr) {
          dailyTotal++;
          if (timeStr) timeSlotCounts[timeStr] = (timeSlotCounts[timeStr] || 0) + 1;
        }
      }
    }

    if (dailyTotal >= effectiveDailyCap) {
      return { isAvailable: false, reason: "당일 체험 마감", availableSlots: [] };
    }

    let allSlots = [];
    try {
      const parseTime = (t) => {
        const clean = String(t || '').trim();
        const p = clean.split(':');
        return parseInt(p[0], 10) * 60 + (p[1] ? parseInt(p[1], 10) : 0);
      };

      const timeRanges = config.bizHours.split(',');

      const blocks = config.blockedTimes.map(b => {
        if (!b.includes('-')) return { type: 'exact', time: parseTime(b) };
        const bp = b.split('-');
        let bs = parseTime(bp[0]), be = parseTime(bp[1]);
        return { type: 'range', start: bs, end: be };
      });

      timeRanges.forEach(rangeStr => {
        const times = rangeStr.trim().split('-');
        if (times.length === 2) {
          const startMin = parseTime(times[0]);
          let endMin = parseTime(times[1]);
          if (endMin < startMin) endMin += 1440;

          // 🎯 curr < endMin (< 미만 적용: 마감 정각 슬롯 생성 제외)
          for (let curr = startMin; curr < endMin; curr += config.intervalMin) {
            const hit = blocks.some(b => b.type === 'exact' ? curr === b.time : (curr >= b.start && curr <= b.end));
            if (!hit) {
              const h = Math.floor(curr / 60) % 24;
              const m = curr % 60;
              allSlots.push((h < 10 ? '0' + h : h) + ':' + (m < 10 ? '0' + m : m));
            }
          }
        }
      });
    } catch(e) {
      return { error: "영업시간 입력 형식 오류: " + e.toString() };
    }

    if (allSlots.length === 0) {
      return { error: "오류: 매장 영업시간 설정 미비" };
    }

    const availableSlots = allSlots.filter(time => (timeSlotCounts[time] || 0) < config.maxPerSlot);
    return availableSlots.length 
      ? { isAvailable: true, availableSlots } 
      : { isAvailable: false, reason: "모든 시간대 예약이 마감되었습니다.", availableSlots: [] };

  } catch (e) {
    return { error: e.toString() };
  }
}

/** Settings 시트에서 운영진 알림 수신 이메일 목록 불러오기 */
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

/** 확정 슬롯 타임 데이터베이스 픽싱 및 점주/운영진 동시 노티 엔진 */
function bookTimeSlot(row, dateStr, timeStr, peopleCount) {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(8000)) {
    return { success: false, error: "현재 접속자가 몰려 처리 중입니다. 잠시 후 다시 눌러주세요." };
  }

  const safeRow = parseInt(row, 10);
  const safePeopleCount = parseInt(peopleCount, 10) || 1;
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Master_Log');

  let mailPayload = null;

  try {
    const realTimeStatus = String(sheet.getRange(safeRow, 12).getValue()).trim();
    if (realTimeStatus !== '예약대기' && realTimeStatus !== '일정조율필요') {
      return { 
        success: false, 
        error: `이미 처리 중이거나 확정된 예약입니다. (현재 상태: ${realTimeStatus})\n새로고침 후 다시 확인해 주세요.` 
      };
    }

    const storeId = String(sheet.getRange(safeRow, 5).getValue()).trim();
    // 🎯 G열(7번째 열) 구매옵션 읽기
    const mealOption = String(sheet.getRange(safeRow, 7).getValue() || '').trim();

    // 🎯 옵션을 포함하여 슬롯 유효성 검사 수행
    const avail = checkAvailability(storeId, dateStr, mealOption);
    if (!avail.isAvailable || !avail.availableSlots.includes(timeStr)) {
      return { 
        success: false, 
        error: "선택하신 시간대의 예약이 방금 마감되었습니다. 다른 시간을 선택해 주세요." 
      };
    }

    const visitDateTime = dateStr + " " + timeStr;
    const dateParts = dateStr.split('-');
    const baseDate = new Date(parseInt(dateParts[0], 10), parseInt(dateParts[1], 10) - 1, parseInt(dateParts[2], 10));
    baseDate.setDate(baseDate.getDate() + 10);

    sheet.getRange(safeRow, 8).setValue(visitDateTime);
    sheet.getRange(safeRow, 9).setValue(safePeopleCount + '명');
    sheet.getRange(safeRow, 10).setValue(baseDate);
    if (!sheet.getRange(safeRow, 11).getValue()) sheet.getRange(safeRow, 11).setValue(10000);
    sheet.getRange(safeRow, 12).setValue('예약확인중');

    SpreadsheetApp.flush();

    mailPayload = {
      storeId: storeId,
      memberCode: String(sheet.getRange(safeRow, 2).getValue()).trim(),
      memberName: String(sheet.getRange(safeRow, 3).getValue()).trim(),
      orderNo: String(sheet.getRange(safeRow, 1).getValue()).trim().replace(/'/g, ''),
      mealOption: mealOption, // 🎯 메일 발송 페이로드에 구매옵션 포함
      dateStr: dateStr,
      timeStr: timeStr,
      safePeopleCount: safePeopleCount,
      safeRow: safeRow
    };

  } catch (err) {
    return { success: false, error: "예약 처리 중 오류가 발생했습니다: " + err.toString() };
  } finally {
    lock.releaseLock();
  }

  try {
    if (mailPayload) {
      _sendBookingNotifications(mailPayload);
    }
  } catch (mailErr) {
    console.error("메일 발송 에러: " + mailErr.toString());
  }

  return { success: true };
}

/** 메일 발송 분리 함수 (주문번호 제거, 5대 항목 고정, 24시간 답변기한 경고, 깨짐 방지 table 구조, Y열 타임스탬프 기록) */
function _sendBookingNotifications(p) {
  const data = _getSheetsData(['Restaurant_List', 'User_DB']);
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Master_Log');

  let storeEmail = "", storeNameJp = "", storeNameKo = "", storeBenefit = "";
  for (let i = 2; i < data.Restaurant_List.length; i++) {
    if (String(data.Restaurant_List[i][0] || '').trim().toUpperCase() === p.storeId.toUpperCase()) {
      storeNameKo = String(data.Restaurant_List[i][1] || '').trim();
      storeNameJp = String(data.Restaurant_List[i][2] || '').trim() || storeNameKo;
      storeEmail = String(data.Restaurant_List[i][10] || '').trim();
      storeBenefit = String(data.Restaurant_List[i][22] || '').trim();
      break;
    }
  }

  let creatorProfileUrl = "";
  for (let u = 1; u < data.User_DB.length; u++) {
    if (String(data.User_DB[u][0] || '').trim().toLowerCase() === p.memberCode.toLowerCase()) {
      creatorProfileUrl = String(data.User_DB[u][4] || '').trim() || 
                          String(data.User_DB[u][5] || '').trim() || 
                          String(data.User_DB[u][6] || '').trim() || 
                          String(data.User_DB[u][7] || '').trim() || 
                          String(data.User_DB[u][8] || '').trim() || "";
      break;
    }
  }

  let profileHtmlClean = `<span style="color: #8b95a1; font-weight: 500;">当日確認</span>`;
  let adminSnsHtml = `<p style="margin: 6px 0;"><b>・SNS:</b> <span style="color: #8b95a1;">未登録 (当日確認)</span></p>`;

  if (creatorProfileUrl) {
    const isHttp = /^https?:\/\//i.test(creatorProfileUrl);
    const isDomainLike = /^[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}/.test(creatorProfileUrl);
    if (isHttp || isDomainLike) {
      const fullUrl = isHttp ? creatorProfileUrl : "https://" + creatorProfileUrl;
      const btnText = (creatorProfileUrl.includes('drive.google.com') || creatorProfileUrl.includes('imweb') || creatorProfileUrl.match(/\.(jpg|jpeg|png|webp|gif)/i))
        ? "プロフィール確認 ↗" : "SNSを見る ↗";
      profileHtmlClean = `<a href="${fullUrl}" target="_blank" style="color: #1a73e8; font-weight: bold; text-decoration: underline;">${btnText}</a>`;
      adminSnsHtml = `<p style="margin: 6px 0;"><b>・SNS:</b> <a href="${fullUrl}" target="_blank" style="color: #1a73e8; font-weight: bold; text-decoration: underline;">${fullUrl}</a></p>`;
    } else {
      profileHtmlClean = `<span style="color: #1A2B49; font-weight: bold;">Google Local Guides (${creatorProfileUrl})</span>`;
      adminSnsHtml = `<p style="margin: 6px 0;"><b>・SNS:</b> <span style="color: #1A2B49; font-weight: bold;">Google Local Guides (${creatorProfileUrl})</span></p>`;
    }
  }

  const benefitDisplay = storeBenefit || "店舗指定のクリエイター向け提供メニュー";
  // 🎯 구매 옵션 문자열을 있는 그대로 플랜명에 반영
  const optionDisplayJp = p.mealOption ? `${p.mealOption} プラン` : '通常プラン';

  // 1. 점주에게 일본어 예약 신청 메일 발송
  if (storeEmail && storeEmail.includes("@")) {
    const scriptUrl = getActiveWebAppUrl();
    const confirmUrl = `${scriptUrl}?mode=store_confirm&row=${p.safeRow}&o=${encodeURIComponent(p.orderNo)}`;
    const feedbackUrl = `${scriptUrl}?mode=feedback&row=${p.safeRow}&o=${encodeURIComponent(p.orderNo)}`;

    const subject = `【BOOKMARK CREATORS】 ご来店予約のご確認 (${p.dateStr})`;
    const htmlBody = `
      <meta charset="UTF-8">
      <div style="font-family: -apple-system, BlinkMacSystemFont, 'Hiragino Kaku Gothic ProN', 'Hiragino Sans', Meiryo, sans-serif; padding: 20px 10px; background-color: #f4f5f7; -webkit-font-smoothing: antialiased;">
        <table align="center" border="0" cellpadding="0" cellspacing="0" width="100%" style="max-width: 500px; background-color: #ffffff; border-radius: 16px; border: 1px solid #edf0f2; box-shadow: 0 4px 16px rgba(0,0,0,0.04); overflow: hidden;">
          <tr>
            <td style="padding: 32px 24px;">
              
              <!-- 상단 브랜딩 -->
              <div style="text-align: center; margin-bottom: 20px;">
                <span style="font-size: 11px; font-weight: 800; letter-spacing: 1px; color: #ffffff; background-color: #1A2B49; padding: 4px 10px; border-radius: 6px; display: inline-block;">RESERVATION REQUEST</span>
                <h1 style="color: #1A2B49; margin: 10px 0 0 0; font-size: 22px; font-weight: 900; letter-spacing: -0.5px;">BOOKMARK CREATORS</h1>
                <div style="width: 36px; height: 3px; background-color: #C5A358; margin: 10px auto 0;"></div>
              </div>

              <h2 style="color: #1A2B49; margin: 18px 0 0 0; font-size: 16.5px; border-bottom: 1.5px solid #f1f3f5; padding-bottom: 14px; text-align: center; font-weight: 800;">来店予約の確認</h2>

              <!-- 점포명 & 담당자 -->
              <div style="margin-top: 18px;">
                <p style="color: #1A2B49; font-size: 15px; font-weight: bold; margin: 0 0 2px 0;">${storeNameJp}</p>
                <p style="color: #6c757d; font-size: 14px; margin: 0 0 6px 0;">店舗管理者様</p>
              </div>

              <!-- 안내 문구 -->
              <div style="margin: 6px 0 16px 0; color: #333333; font-size: 14px; line-height: 1.55;">
                <p style="margin: 0 0 3px 0;">BOOKMARK CREATORSより来店予約の申請が届きました。</p>
                <p style="margin: 0;">下記の内容をご確認いただき、<b>【確定】</b>または<b>【日時変更】</b>のご対応をお願い申し上げます。</p>
              </div>

              <!-- 24시간 기한 안내 -->
              <div style="background-color: #fbfbfc; border: 1px solid #edf0f2; border-radius: 10px; padding: 12px 16px; margin-bottom: 20px; text-align: left;">
                <p style="margin: 0; font-size: 13px; color: #c92a2a; line-height: 1.55; font-weight: 700;">
                  ※ 本メール送信後24時間以内にご回答がない場合、自動的にキャンセル処理が行われます。
                </p>
              </div>

              <!-- 예약 상세 정보 테이블 (옵션 플랜 행 추가) -->
              <table border="0" cellpadding="0" cellspacing="0" width="100%" style="background-color: #ffffff; border: 1.5px solid #edf0f2; border-radius: 12px; padding: 14px 16px; margin-bottom: 24px; font-size: 13.5px;">
                <tr>
                  <td style="padding: 6px 0; color: #333333;">
                    <strong>■ クリエイター名:</strong> ${p.memberName}
                  </td>
                </tr>
                <tr>
                  <td style="padding: 6px 0; color: #333333;">
                    <strong>■ 訪問日時:</strong> <span style="color: #d63384; font-weight: 800;">${p.dateStr} ${p.timeStr}</span>
                  </td>
                </tr>
                <tr>
                  <td style="padding: 6px 0; color: #333333;">
                    <strong>■ 訪問人数:</strong> ${p.safePeopleCount}名
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

              <!-- 버튼 영역 (2열 정렬) -->
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

    try {
      const adminAlertEmails = getAdminAlertEmails();
      GmailApp.sendEmail(storeEmail, subject, "", { 
        htmlBody: htmlBody, 
        name: "BOOKMARK CREATORS", 
        from: "info@bookmarkfukuoka.jp",
        bcc: adminAlertEmails
      });
      const sentTimeStr = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm");
      sheet.getRange(p.safeRow, 25).setValue(`[${sentTimeStr}] 점주메일 발송완료`);
    } catch (mailErr) {
      console.error("점주 이메일 발송 실패: " + mailErr.toString());
      sheet.getRange(p.safeRow, 25).setValue("❌ 점주메일 실패: " + mailErr.toString());
    }
  } else {
    sheet.getRange(p.safeRow, 25).setValue("❌ 실패: 점주 이메일 주소 없음");
  }

  // 2. 운영진 실시간 알림 발송
  try {
    const adminAlertEmails = getAdminAlertEmails();
    if (adminAlertEmails) {
      const adminSubject = `【予約受付】${p.memberName}様 ➔ ${storeNameJp || storeNameKo} (${p.dateStr} ${p.timeStr})`;
      const adminHtml = `
        <meta charset="UTF-8">
        <div style="font-family: 'Helvetica Neue', Arial, 'Hiragino Kaku Gothic ProN', Meiryo, sans-serif; padding: 24px; background: #f8f9fa; border-radius: 14px; border: 1px solid #e9ecef; max-width: 520px; margin: 0 auto; color: #333; line-height: 1.6;">
          <div style="margin-bottom: 18px;">
            <span style="background: #1A2B49; color: #fff; font-size: 11px; font-weight: bold; padding: 4px 10px; border-radius: 6px; letter-spacing: 0.5px;">ADMIN NOTIFICATION</span>
            <h3 style="color: #1A2B49; margin: 10px 0 0 0; font-size: 18px; font-weight: 800;">クリエイター来店予約の受付</h3>
          </div>
          <p style="font-size: 13.5px; color: #495057; margin: 0 0 16px 0;">
            クリエイターより来店予約申請が届き、店舗管理者へ確認メールを送信しました。<br>
            現在、店舗側による【確定】または【日程変更】の対応待ちステータスです。
          </p>
          <div style="background: #ffffff; border-radius: 10px; padding: 18px; margin: 16px 0; border: 1px solid #eef0f2; font-size: 13.5px;">
            <p style="margin: 6px 0;"><b>・注文번호:</b> #${p.orderNo}</p>
            <p style="margin: 6px 0;"><b>・クリエイター:</b> ${p.memberName} <span style="color:#8b95a1; font-size: 12px;">(${p.memberCode})</span></p>
            <p style="margin: 6px 0;"><b>・訪問店舗:</b> ${storeNameJp || storeNameKo} <span style="color:#8b95a1; font-size: 12px;">(${p.storeId})</span></p>
            <p style="margin: 6px 0;"><b>・提供内容:</b> <span style="color: #2D6A4F; font-weight: bold;">${storeBenefit || '店舗指定のクリエイター向け提供メニュー'}</span></p>
            <p style="margin: 6px 0;"><b>・予約日時:</b> <span style="color: #d63384; font-weight: bold;">${p.dateStr} ${p.timeStr}</span></p>
            <p style="margin: 6px 0;"><b>・訪問人数:</b> <span style="font-weight: bold;">${p.safePeopleCount}人</span></p>
            ${adminSnsHtml}
            <p style="margin: 6px 0;"><b>・店舗受信メール:</b> <span style="color: #1a73e8;">${storeEmail || '未登録'}</span></p>
          </div>
        </div>
      `;
      GmailApp.sendEmail(adminAlertEmails, adminSubject, "", { 
        htmlBody: adminHtml, 
        name: "BOOKMARK NOTI" 
      });
    }
  } catch (adminErr) {
    console.error("운영진 알림 발송 실패: " + adminErr.toString());
  }
}
