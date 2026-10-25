/** 크리에이터 도메인: 미션 조회, 슬롯 검증/예약, 방문 확인, 리뷰 제출 */

/** [3] 미션 목록 조회 (승인 채널 목록 approvedChannels, minDays/maxDays, G열 구매옵션 연동) */
function getMissions(orderNo, phoneLast4) {
  try {
    const data = _getSheetsData([SHEET.MASTER, SHEET.USER, SHEET.REST]);
    const inputOrderNo = String(orderNo || '').trim().replace(/['"\s]/g, '');
    const inputPhone = String(phoneLast4 || '').trim().replace(/[^0-9]/g, '');

    const userMap = new Map(); 
    const userTierMap = new Map(); 
    const userChannelsMap = new Map();

    for (let j = 1; j < data[SHEET.USER].length; j++) {
      const uRow = data[SHEET.USER][j]; 
      if (!uRow || !uRow[IDX_USER.KEY]) continue; 
      
      const mCode = String(uRow[IDX_USER.KEY]).trim();
      const mEmail = String(uRow[IDX_USER.EMAIL] || '').trim();
      const phoneStr = uRow[IDX_USER.PHONE] ? String(uRow[IDX_USER.PHONE]).replace(/[^0-9]/g, '') : '';
      
      if (mCode) userMap.set(mCode.toLowerCase(), phoneStr);
      if (mEmail) userMap.set(mEmail.toLowerCase(), phoneStr);
      userTierMap.set(mCode.toLowerCase(), parseTierEmoji(uRow[IDX_USER.GROUP])); 

      const channels = [];
      if (String(uRow[IDX_USER.BLOG] || '').trim().length > 3) channels.push({ type: 'blog', name: '네이버 블로그', icon: '📝', placeholder: '블로그 리뷰 포스팅 URL' });
      if (String(uRow[IDX_USER.INSTA] || '').trim().length > 3) channels.push({ type: 'instagram', name: '인스타그램(릴스)', icon: '📸', placeholder: '인스타그램 릴스/게시물 URL' });
      if (String(uRow[IDX_USER.YOUTUBE] || '').trim().length > 3) channels.push({ type: 'youtube', name: '유튜브', icon: '🎥', placeholder: '유튜브 영상/쇼츠 URL' });
      if (String(uRow[IDX_USER.TIKTOK] || '').trim().length > 3) channels.push({ type: 'tiktok', name: '틱톡', icon: '🎵', placeholder: '틱톡 영상 URL' });
      if (String(uRow[IDX_USER.GOOGLE_GUIDE] || '').trim().length > 3) channels.push({ type: 'google', name: '구글 로컬 가이드', icon: '🗺️', placeholder: '로컬 가이드 리뷰 링크' });

      if (channels.length === 0) {
        channels.push({ type: 'instagram', name: '인스타그램(릴스)', icon: '📸', placeholder: '인스타그램 릴스/게시물 URL' });
      }

      if (mCode) userChannelsMap.set(mCode.toLowerCase(), channels);
      if (mEmail) userChannelsMap.set(mEmail.toLowerCase(), channels);
    }
    
    const restMap = new Map();
    const restNameMap = new Map();
    for (let k = 1; k < data[SHEET.REST].length; k++) {
      const rRow = data[SHEET.REST][k]; 
      if (!rRow || !rRow[IDX_REST.ID]) continue; 
      const rId = String(rRow[IDX_REST.ID]).trim().toUpperCase();
      const rName = String(rRow[IDX_REST.NAME_KO] || '').trim();
      const storeType = String(rRow[IDX_REST.STORE_TYPE] || '').trim().toUpperCase();
      const maxPeople = parseInt(rRow[IDX_REST.MAX_PEOPLE], 10) || 4;

      const minLeadDays = parseInt(rRow[IDX_REST.MIN_DAYS], 10) || 1;
      const maxLeadDays = parseInt(rRow[IDX_REST.MAX_DAYS], 10) || 90;

      const rDetails = { 
        map: String(rRow[IDX_REST.MAP_URL] || '#').trim(),
        guide: String(rRow[IDX_REST.GUIDE_URL] || '#').trim(),
        bookingUrl: (storeType === 'RETAIL') ? '' : String(rRow[IDX_REST.BOOKING_URL] || '').trim(), 
        storeType: storeType, 
        maxPeople: maxPeople,
        minDays: minLeadDays,
        maxDays: maxLeadDays,
        blackouts: getSafeBlackouts(rRow[IDX_REST.BLACKOUTS]),
        benefit: String(rRow[IDX_REST.BENEFIT] || '').trim() // 🎯 [추가] W열(23번째 열, index 22) 제공 메뉴/혜택
      };
      
      if (rId) restMap.set(rId, rDetails);
      if (rName) restNameMap.set(rName, rDetails);
    }

    const mList = []; 
    const timeZone = Session.getScriptTimeZone();

    for (let i = 1; i < data[SHEET.MASTER].length; i++) {
      const mRow = data[SHEET.MASTER][i]; 
      if (!mRow || !mRow[IDX_MASTER.ORDER]) continue; 
      
      const status = String(mRow[IDX_MASTER.STATUS] || '').trim(); 
      const sheetOrderNo = String(mRow[IDX_MASTER.ORDER]).replace(/['"\s]/g, '').trim(); 
      
      if (status.includes('취소') || sheetOrderNo !== inputOrderNo) continue; 

      const memberCode = String(mRow[IDX_MASTER.MEMBER] || '').trim().toLowerCase();
      const userPhone = userMap.get(memberCode) || '';
      
      if (!userPhone.endsWith(inputPhone)) continue;

      const restId = String(mRow[IDX_MASTER.STORE_ID] || '').trim().toUpperCase();
      const restaurantName = String(mRow[IDX_MASTER.STORE_NAME] || '').trim();
      
      let restInfo = restMap.get(restId);
      if (!restInfo && restaurantName) {
        restInfo = restNameMap.get(restaurantName);
      }
      if (!restInfo) {
        restInfo = { map: '#', guide: '#', bookingUrl: '', storeType: '', maxPeople: 4, minDays: 1, maxDays: 90, blackouts: [] };
      }
      
      const formatDate = (val, format) => (val instanceof Date) ? Utilities.formatDate(val, timeZone, format) : String(val || '');
      const deadline = formatDate(mRow[IDX_MASTER.DEADLINE], 'yyyy-MM-dd');
      let vDate = formatDate(mRow[IDX_MASTER.VISIT], 'yyyy-MM-dd HH:mm');
      if (vDate === '-' || vDate === '') vDate = '';

      mList.push({
        row: i + 1, orderNo: inputOrderNo, restId, memberCode,
        member: String(mRow[IDX_MASTER.NAME] || ''), restaurant: restaurantName, status,
        mealOption: String(mRow[IDX_MASTER.OPTION] || '').trim(), // 🎯 G열 구매옵션 (인덱스 6)
        captureUrl: String(mRow[IDX_MASTER.OPTION] || ''), 
        verified: String(mRow[IDX_MASTER.VISITED] || ''),  
        receiptUrl: String(mRow[IDX_MASTER.RECEIPT] || ''), 
        submitDate: formatDate(mRow[IDX_MASTER.SUBMIT_DATE], 'yyyy-MM-dd'), 
        reviewUrl: String(mRow[IDX_MASTER.REVIEW_URL] || ''),  
        googleMapUrl: String(mRow[IDX_MASTER.GOOGLE_URL] || ''), 
        shortReview: String(mRow[IDX_MASTER.SHORT_REVIEW] || ''), 
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
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET.MASTER);
    if (String(sheet.getRange(safeRow, COL_MASTER.STATUS).getValue()).trim() !== '방문전') {
      return { success: false, error: "既に訪問確認済みか、対象外のステータスです。" };
    } 
    sheet.getRange(safeRow, COL_MASTER.STATUS).setValue('방문완료'); 
    sheet.getRange(safeRow, COL_MASTER.VISITED).setValue('Y');
    return { success: true };
  } catch (e) { return { success: false, error: e.toString() }; }
}

/** 최종 리뷰 콘텐츠 제출 */
function submitContent(row, reviewUrl, googleMapUrl, shortReview, channelType) {
  try {
    const safeRow = parseInt(row, 10); 
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET.MASTER);
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

    sheet.getRange(safeRow, COL_MASTER.SUBMIT_DATE).setValue(Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd"));
    sheet.getRange(safeRow, COL_MASTER.REVIEW_URL).setValue(cleanReviewUrl);
    if (cleanGoogleUrl) sheet.getRange(safeRow, COL_MASTER.GOOGLE_URL).setValue(cleanGoogleUrl); 
    
    const safeReview = String(shortReview || '').trim();
    if (safeReview) {
      sheet.getRange(safeRow, COL_MASTER.SHORT_REVIEW).setValue(safeReview);
      try { sheet.getRange(safeRow, COL_MASTER.SHORT_REVIEW_JP).setValue(LanguageApp.translate(safeReview, 'ko', 'ja')); } catch(txtErr) {}
    }
    sheet.getRange(safeRow, COL_MASTER.STATUS).setValue('제출완료');
    sheet.getRange(safeRow, COL_MASTER.REFUND).setValue('환급대기');
    return { success: true };
  } catch (e) { return { success: false, error: e.toString() }; }
}

/** 수정 제출 처리 */
function updateSubmission(row, newUrl, newGoogleUrl, newShortReview) {
  try {
    const safeRow = parseInt(row, 10); 
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET.MASTER);
    const fixUrl = (url) => { const u = String(url || '').trim(); return (u && !/^https?:\/\//i.test(u)) ? 'https://' + u : u; };

    sheet.getRange(safeRow, COL_MASTER.REVIEW_URL).setValue(fixUrl(newUrl));
    sheet.getRange(safeRow, COL_MASTER.GOOGLE_URL).setValue(fixUrl(newGoogleUrl));
    if(newShortReview !== undefined) sheet.getRange(safeRow, COL_MASTER.SHORT_REVIEW).setValue(String(newShortReview).trim());
    return { success: true };
  } catch (e) { return { success: false, error: e.toString() }; }
}

/* 🎯 [동적 키워드 매칭] 고속 슬롯 필터 제너레이터 */
function checkAvailability(storeId, targetDateStr, mealOption) {
  try {
    const data = _getSheetsData([SHEET.REST, SHEET.MASTER]);
    let config = null;

    const cleanInputId = String(storeId || '').trim().toUpperCase().replace(/['"]/g, '');
    const userOption = String(mealOption || '').trim().toLowerCase();
    const restList = data[SHEET.REST];

    for (let i = 1; i < restList.length; i++) {
      const rowId = String(restList[i][IDX_REST.ID] || '').trim().toUpperCase().replace(/['"]/g, '');
      const rowNameKo = String(restList[i][IDX_REST.NAME_KO] || '').trim();
      const rowNameJp = String(restList[i][IDX_REST.NAME_JP] || '').trim();

      if (rowId === cleanInputId || (cleanInputId && (cleanInputId === rowNameKo || cleanInputId === rowNameJp))) {
        // 1. 기본 운영시간 (O열: 인덱스 14)
        let rawBiz = toHalfWidth(restList[i][IDX_REST.BIZ_HOURS] || '');

        // 2. 동적 옵션 시간대 1 (AA: 매칭키워드[인덱스 26], AB: 운영시간[인덱스 27])
        const opt1Keywords = String(restList[i][IDX_REST.OPT1_KEYWORDS] || '').trim().toLowerCase();
        const opt1Hours    = String(restList[i][IDX_REST.OPT1_HOURS] || '').replace(/[~～－—]/g, '-').trim();

        // 3. 동적 옵션 시간대 2 (AC: 매칭키워드[인덱스 28], AD: 운영시간[인덱스 29])
        const opt2Keywords = String(restList[i][IDX_REST.OPT2_KEYWORDS] || '').trim().toLowerCase();
        const opt2Hours    = String(restList[i][IDX_REST.OPT2_HOURS] || '').replace(/[~～－—]/g, '-').trim();

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
          maxPerSlot: parseInt(restList[i][IDX_REST.MAX_PER_SLOT], 10) || 1,
          rawDailyCap: String(restList[i][IDX_REST.DAILY_CAP] || '').trim(),
          blackouts: getSafeBlackouts(restList[i][IDX_REST.BLACKOUTS]),
          bizHours: rawBiz,
          intervalMin: parseInt(restList[i][IDX_REST.INTERVAL_MIN], 10) || 30,
          blockedTimes: String(restList[i][IDX_REST.BLOCKED_TIMES] || '').split(',').map(s => s.trim()).filter(Boolean)
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
    const masterLog = data[SHEET.MASTER];

    for (let i = 2; i < masterLog.length; i++) {
      const mRow = masterLog[i];
      const mStoreId = String(mRow[IDX_MASTER.STORE_ID] || '').trim().toUpperCase();
      if (mStoreId === cleanInputId && !String(mRow[IDX_MASTER.STATUS]).includes('취소') && !String(mRow[IDX_MASTER.STATUS]).includes('노쇼')) {
        let v = mRow[IDX_MASTER.VISIT];
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
  const sheet = ss.getSheetByName(SHEET.MASTER);

  let mailPayload = null;

  try {
    const realTimeStatus = String(sheet.getRange(safeRow, COL_MASTER.STATUS).getValue()).trim();
    if (realTimeStatus !== '예약대기' && realTimeStatus !== '일정조율필요') {
      return { 
        success: false, 
        error: `이미 처리 중이거나 확정된 예약입니다. (현재 상태: ${realTimeStatus})\n새로고침 후 다시 확인해 주세요.` 
      };
    }

    const storeId = String(sheet.getRange(safeRow, COL_MASTER.STORE_ID).getValue()).trim();
    // 🎯 G열(7번째 열) 구매옵션 읽기
    const mealOption = String(sheet.getRange(safeRow, COL_MASTER.OPTION).getValue() || '').trim();

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

    sheet.getRange(safeRow, COL_MASTER.VISIT).setValue(visitDateTime);
    sheet.getRange(safeRow, COL_MASTER.PEOPLE).setValue(safePeopleCount + '명');
    sheet.getRange(safeRow, COL_MASTER.DEADLINE).setValue(baseDate);
    if (!sheet.getRange(safeRow, COL_MASTER.DEPOSIT).getValue()) sheet.getRange(safeRow, COL_MASTER.DEPOSIT).setValue(10000);
    sheet.getRange(safeRow, COL_MASTER.STATUS).setValue('예약확인중');

    SpreadsheetApp.flush();

    mailPayload = {
      storeId: storeId,
      memberCode: String(sheet.getRange(safeRow, COL_MASTER.MEMBER).getValue()).trim(),
      memberName: String(sheet.getRange(safeRow, COL_MASTER.NAME).getValue()).trim(),
      orderNo: String(sheet.getRange(safeRow, COL_MASTER.ORDER).getValue()).trim().replace(/'/g, ''),
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
