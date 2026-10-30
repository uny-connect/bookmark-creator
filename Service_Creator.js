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
    for (let k = REST_DATA_START_INDEX; k < data[SHEET.REST].length; k++) {
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
        bookingUrl: (storeType === 'RETAIL' || storeType === STORE_TYPE_STAY) ? '' : String(rRow[IDX_REST.BOOKING_URL] || '').trim(),
        storeType: storeType,
        maxPeople: maxPeople,
        minDays: minLeadDays,
        maxDays: maxLeadDays,
        blackouts: getSafeBlackouts(rRow[IDX_REST.BLACKOUTS]),
        benefit: String(rRow[IDX_REST.BENEFIT] || '').trim() // 🎯 [추가] W열(23번째 열, index 22) 제공 메뉴/혜택
      };
      if (storeType === STORE_TYPE_STAY) {
        const stay = _getStayConfig_(rRow);
        rDetails.checkinTime = stay.checkinTime;
        rDetails.checkoutTime = stay.checkoutTime;
        rDetails.maxNights = stay.maxNights;
        rDetails.autoCheckin = stay.autoCheckin;
      }
      
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
        deadline, visitDate: vDate,
        checkoutDate: vDate ? formatDate(mRow[IDX_MASTER.CHECKOUT], 'yyyy-MM-dd') : '', // STAY 전용(체크인 H 값이 있을 때만)
        nights: vDate ? (parseInt(mRow[IDX_MASTER.NIGHTS], 10) || 0) : 0,
        tierEmoji: userTierMap.get(memberCode) || '🟡',
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

    for (let i = REST_DATA_START_INDEX; i < restList.length; i++) {
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
    const dayKoMap = ['일', '월', '화', '수', '목', '금', '토'];
    const tDayJa = _DAY_JA[targetDateObj.getDay()];
    const tDayKo = dayKoMap[targetDateObj.getDay()];

    const isBlackout = _isBlackoutDate_(config.blackouts, targetDateStr, targetDateObj.getDay(), targetDateObj.getDate());

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
        const { dateStr, timeStr } = _splitVisit_(v, timeZone);
        
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

      // 🎯 요일 접두(예: 火木日11:00-13:00)가 있으면 해당 요일에만 차단, 없으면 매일
      const blocks = config.blockedTimes.map(raw => {
        const dm = String(raw).trim().match(/^([日月火水木金土]+)\s*(.*)$/);
        const days = dm ? dm[1] : '';
        const b = dm ? dm[2] : String(raw).trim();
        if (!b.includes('-')) return { type: 'exact', time: parseTime(b), days };
        const bp = b.split('-');
        let bs = parseTime(bp[0]), be = parseTime(bp[1]);
        return { type: 'range', start: bs, end: be, days };
      }).filter(b => !b.days || b.days.includes(tDayJa));

      timeRanges.forEach(rangeStr => {
        const times = rangeStr.trim().split('-');
        if (times.length === 2) {
          const startMin = parseTime(times[0]);
          let endMin = parseTime(times[1]);
          if (endMin < startMin) endMin += 1440;

          // 🎯 curr < endMin (< 미만 적용: 마감 정각 슬롯 생성 제외)
          for (let curr = startMin; curr < endMin; curr += config.intervalMin) {
            // 🌙 자정 넘는 영업시간(18:00-03:00)은 curr 가 1440 이상이므로 24h 를 뺀 시각(01:30 등)으로도 차단 대조
            const inBlock = (b, t) => b.type === 'exact' ? t === b.time : (t >= b.start && t <= b.end);
            const hit = blocks.some(b => inBlock(b, curr) || (curr >= 1440 && inBlock(b, curr - 1440)));
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

    const maxPeople = _getStoreMaxPeople_(storeId);
    if (safePeopleCount > maxPeople) {
      return { success: false, error: `이 매장은 최대 ${maxPeople}명까지 예약할 수 있습니다.` };
    }

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

/*********************** 🏨 숙박(STAY) 예약 ***********************/

const _DAY_JA = ['日', '月', '火', '水', '木', '金', '土'];

/** 휴무 규칙(날짜/요일/第N요일) 일치 여부. checkAvailability 와 checkStayAvailability 공용 */
function _isBlackoutDate_(blackouts, dateStr, dayIdx, dayOfMonth) {
  const dayJa = _DAY_JA[dayIdx];
  const weekNum = Math.ceil(dayOfMonth / 7);
  return blackouts.some(b => {
    if (!b) return false;
    if (b === dateStr || b === dayJa) return true;
    if (b.includes('第')) {
      const match = b.match(/第([1-5])/);
      return match && parseInt(match[1], 10) === weekNum && b.includes(dayJa);
    }
    return false;
  });
}

/** 'yyyy-MM-dd' → UTC 자정 ms (형식 오류 시 null). 박수 계산을 타임존과 무관하게 하기 위해 UTC 사용 */
function _ymdToUtc_(s) {
  const m = String(s || '').match(/^(\d{4})-(\d{2})-(\d{2})/);
  return m ? Date.UTC(parseInt(m[1], 10), parseInt(m[2], 10) - 1, parseInt(m[3], 10)) : null;
}
function _utcToYmd_(ms) { return new Date(ms).toISOString().substring(0, 10); }
const _MS_DAY = 86400000;

/** 시트 셀(Date 또는 문자열) → 'yyyy-MM-dd' */
function _cellToYmd_(v) {
  if (v instanceof Date) return Utilities.formatDate(v, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  return String(v || '').substring(0, 10);
}

/** 점포 행 찾기: ID 또는 KO/JP 매장명 일치 */
function _findRestRow_(restList, storeId) {
  const cleanId = String(storeId || '').trim().toUpperCase().replace(/['"]/g, '');
  for (let i = REST_DATA_START_INDEX; i < restList.length; i++) {
    const r = restList[i];
    const rowId = String(r[IDX_REST.ID] || '').trim().toUpperCase().replace(/['"]/g, '');
    if (rowId === cleanId || (cleanId && (cleanId === String(r[IDX_REST.NAME_KO] || '').trim() || cleanId === String(r[IDX_REST.NAME_JP] || '').trim()))) return r;
  }
  return null;
}

/** 서버 측 최대 인원 검증용 (시트 값이 비어 있으면 UI 기본값 4) */
function _getStoreMaxPeople_(storeId) {
  const row = _findRestRow_(_getSheetsData([SHEET.REST])[SHEET.REST], storeId);
  return row ? (parseInt(row[IDX_REST.MAX_PEOPLE], 10) || 4) : 4;
}

/** STAY 전용 설정 (빈 값은 기본값) */
function _getStayConfig_(restRow) {
  // 시트에서 15:00 을 직접 입력하면 시간(Date) 셀이 되므로 HH:mm 으로 되돌린다
  const hm = v => (v instanceof Date) ? Utilities.formatDate(v, Session.getScriptTimeZone(), 'HH:mm') : String(v || '').trim();
  return {
    rooms: parseInt(restRow[IDX_REST.ROOMS], 10) || 1,
    checkinTime: hm(restRow[IDX_REST.CHECKIN_TIME]) || '15:00',
    checkoutTime: hm(restRow[IDX_REST.CHECKOUT_TIME]) || '11:00',
    maxNights: parseInt(restRow[IDX_REST.MAX_NIGHTS], 10) || 7,
    autoCheckin: String(restRow[IDX_REST.AUTO_CHECKIN] || '').trim().toUpperCase() === 'Y'
  };
}

/** 숙박 가능 여부: 박수/리드타임/휴무(숙박하는 밤 기준, 체크아웃일 제외)/일일 객실수 */
function checkStayAvailability(storeId, checkInStr, checkOutStr) {
  try {
    const data = _getSheetsData([SHEET.REST, SHEET.MASTER]);
    const restRow = _findRestRow_(data[SHEET.REST], storeId);
    if (!restRow) return { error: "오류: 매장 확인 불가" };
    if (String(restRow[IDX_REST.STORE_TYPE] || '').trim().toUpperCase() !== STORE_TYPE_STAY) {
      return { error: "오류: 숙박 매장이 아닙니다" };
    }

    const inMs = _ymdToUtc_(checkInStr), outMs = _ymdToUtc_(checkOutStr);
    if (inMs === null || outMs === null) return { error: "오류: 날짜 형식이 올바르지 않습니다" };

    const stay = _getStayConfig_(restRow);
    const nights = Math.round((outMs - inMs) / _MS_DAY);
    if (nights < 1 || nights > stay.maxNights) {
      return { isAvailable: false, reason: `숙박은 1박~${stay.maxNights}박까지 가능합니다.` };
    }

    const todayMs = _ymdToUtc_(Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd'));
    const minDays = parseInt(restRow[IDX_REST.MIN_DAYS], 10) || 1;
    const maxDays = parseInt(restRow[IDX_REST.MAX_DAYS], 10) || 90;
    if (inMs < todayMs + minDays * _MS_DAY || outMs > todayMs + maxDays * _MS_DAY) {
      return { isAvailable: false, reason: "예약 가능 기간을 벗어났습니다." };
    }

    const blackouts = getSafeBlackouts(restRow[IDX_REST.BLACKOUTS]);
    const cleanId = String(restRow[IDX_REST.ID] || '').trim().toUpperCase();
    const occupied = {};
    const masterLog = data[SHEET.MASTER];
    for (let i = 2; i < masterLog.length; i++) {
      const mRow = masterLog[i];
      const status = String(mRow[IDX_MASTER.STATUS] || '');
      if (String(mRow[IDX_MASTER.STORE_ID] || '').trim().toUpperCase() !== cleanId || status.includes('취소') || status.includes('노쇼')) continue;
      const s = _ymdToUtc_(_cellToYmd_(mRow[IDX_MASTER.VISIT]));
      if (s === null) continue;
      const eRaw = _ymdToUtc_(_cellToYmd_(mRow[IDX_MASTER.CHECKOUT]));
      const e = (eRaw !== null && eRaw > s) ? eRaw : s + _MS_DAY;
      for (let d = s; d < e; d += _MS_DAY) occupied[d] = (occupied[d] || 0) + 1;
    }

    for (let d = inMs; d < outMs; d += _MS_DAY) {
      const dt = new Date(d);
      if (_isBlackoutDate_(blackouts, _utcToYmd_(d), dt.getUTCDay(), dt.getUTCDate())) {
        return { isAvailable: false, reason: `${_utcToYmd_(d)} 은(는) 예약 불가일입니다.` };
      }
      if ((occupied[d] || 0) >= stay.rooms) {
        return { isAvailable: false, reason: `${_utcToYmd_(d)} 은(는) 객실이 마감되었습니다.` };
      }
    }

    return {
      isAvailable: true, nights: nights, checkinTime: stay.checkinTime, checkoutTime: stay.checkoutTime,
      maxPeople: parseInt(restRow[IDX_REST.MAX_PEOPLE], 10) || 4
    };
  } catch (e) {
    return { error: e.toString() };
  }
}

/** 숙박 예약 신청: bookTimeSlot 과 동일한 흐름(상태 예약확인중 → 점주 확인 메일) */
function bookStay(row, checkInStr, checkOutStr, peopleCount) {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(8000)) {
    return { success: false, error: "현재 접속자가 몰려 처리 중입니다. 잠시 후 다시 눌러주세요." };
  }

  const safeRow = parseInt(row, 10);
  const safePeopleCount = parseInt(peopleCount, 10) || 1;
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET.MASTER);

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
    const avail = checkStayAvailability(storeId, checkInStr, checkOutStr);
    if (avail.error) return { success: false, error: avail.error };
    if (!avail.isAvailable) return { success: false, error: avail.reason || "선택하신 일정은 예약할 수 없습니다." };
    if (safePeopleCount > avail.maxPeople) {
      return { success: false, error: `이 숙소는 최대 ${avail.maxPeople}명까지 예약할 수 있습니다.` };
    }

    const outMs = _ymdToUtc_(checkOutStr);
    sheet.getRange(safeRow, COL_MASTER.VISIT).setValue(checkInStr + " " + avail.checkinTime);
    sheet.getRange(safeRow, COL_MASTER.PEOPLE).setValue(safePeopleCount + '명');
    sheet.getRange(safeRow, COL_MASTER.DEADLINE).setValue(new Date(outMs + 10 * _MS_DAY)); // 마감일: 체크아웃일 + 10일
    sheet.getRange(safeRow, COL_MASTER.CHECKOUT).setValue(checkOutStr);
    sheet.getRange(safeRow, COL_MASTER.NIGHTS).setValue(avail.nights);
    if (!sheet.getRange(safeRow, COL_MASTER.DEPOSIT).getValue()) sheet.getRange(safeRow, COL_MASTER.DEPOSIT).setValue(10000);
    sheet.getRange(safeRow, COL_MASTER.STATUS).setValue('예약확인중');

    SpreadsheetApp.flush();

    mailPayload = {
      storeId: storeId,
      memberCode: String(sheet.getRange(safeRow, COL_MASTER.MEMBER).getValue()).trim(),
      memberName: String(sheet.getRange(safeRow, COL_MASTER.NAME).getValue()).trim(),
      orderNo: String(sheet.getRange(safeRow, COL_MASTER.ORDER).getValue()).trim().replace(/'/g, ''),
      mealOption: String(sheet.getRange(safeRow, COL_MASTER.OPTION).getValue() || '').trim(),
      dateStr: _stayRangeText_(checkInStr, checkOutStr, avail.nights),
      timeStr: '',
      isStay: true,
      safePeopleCount: safePeopleCount,
      safeRow: safeRow
    };
  } catch (err) {
    return { success: false, error: "예약 처리 중 오류가 발생했습니다: " + err.toString() };
  } finally {
    lock.releaseLock();
  }

  try {
    if (mailPayload) _sendBookingNotifications(mailPayload);
  } catch (mailErr) {
    console.error("메일 발송 에러: " + mailErr.toString());
  }

  return { success: true };
}

/** 메일용 숙박 일정 문구: 기본(일본어, 점주·운영진) "2026-11-05 ~ 2026-11-07 (2泊3日)", lang='ko'(크리에이터) "… (2박 3일)" */
function _stayRangeText_(checkInStr, checkOutStr, nights, lang) {
  const span = (lang === 'ko') ? `${nights}박 ${nights + 1}일` : `${nights}泊${nights + 1}日`;
  return `${checkInStr} ~ ${checkOutStr} (${span})`;
}

/** Master_Log 한 행 배열(IDX_MASTER 기준) → 메일 링크용 일정 토큰. H가 비면 V·W 는 stale 이므로 무시 */
function _visitTokenFromRow_(rowArr, timeZone) {
  const nights = parseInt(rowArr[IDX_MASTER.NIGHTS], 10) || 0;
  const checkout = nights > 0 ? _cellToYmd_(rowArr[IDX_MASTER.CHECKOUT]) : '';
  return _visitToken_(rowArr[IDX_MASTER.VISIT], timeZone, checkout);
}

/** 메일용 일정 표기: 숙박 행(V·W열 값 있음)이면 "체크인 ~ 체크아웃 (N박 N+1일)", 아니면 방문일시 그대로. lang 은 _stayRangeText_ 참고 */
function _visitDisplayForRow_(sheet, safeRow, visitDateStr, lang) {
  const nights = parseInt(sheet.getRange(safeRow, COL_MASTER.NIGHTS).getValue(), 10) || 0;
  const checkout = _cellToYmd_(sheet.getRange(safeRow, COL_MASTER.CHECKOUT).getValue());
  // 일정이 비어 있는 행(H열 없음)에 남은 V·W 값은 무시
  return (nights > 0 && checkout && /^\d{4}-\d{2}-\d{2}/.test(String(visitDateStr))) ? _stayRangeText_(String(visitDateStr).substring(0, 10), checkout, nights, lang) : visitDateStr;
}
