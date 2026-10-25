/** 크리에이터 도메인: 미션 조회, 슬롯 검증/예약, 방문 확인, 리뷰 제출 */

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
