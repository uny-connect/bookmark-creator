/*****************************************************************************************
 * 점주(파트너) 도메인: 예약 승인/일정변경 요청, PIN 로그인, 대시보드, 매장 설정, 간편 등록
 ****************************************************************************************/
/** 📩 메일 직통 링크 예약 승인 엔진 (doGet mode=store_confirm 에서 위임) → { isExpired, isAlreadyRequested } */
function confirmBookingByLink(params) {
    let isAlreadyRequested = false; 
    let isExpired = false; // 24시간 만료 플래그
    
    try {
      const safeRow = parseInt(params.row, 10);
      const orderNo = params.o;
      const ss = SpreadsheetApp.getActiveSpreadsheet();
      const sheet = ss.getSheetByName(SHEET.MASTER);
      const restSheet = ss.getSheetByName(SHEET.REST);
      
      const sheetOrderNo = String(sheet.getRange(safeRow, COL_MASTER.ORDER).getValue()).trim().replace(/'/g, '');
      const currentStatus = String(sheet.getRange(safeRow, COL_MASTER.STATUS).getValue()).trim();
      const storeId = String(sheet.getRange(safeRow, COL_MASTER.STORE_ID).getValue() || '').trim().toUpperCase();

      // 🚨 1. 이미 다른 상태로 처리된 건인지 검사
      if (currentStatus === '일정조율필요' || currentStatus.includes('취소') || currentStatus.includes('노쇼')) {
        isAlreadyRequested = true;
      }
      // 🔒 이미 확정(방문전)·방문완료·제출완료된 건은 링크를 다시 눌러도 상태/메일/만료 처리를 하지 않는다 (멱등)
      else if (currentStatus === '방문전' || currentStatus === '방문완료' || currentStatus === '제출완료') {
        // 변경 없음 → 확정 완료 화면만 표시
      }
      else if (safeRow && sheetOrderNo === String(orderNo).trim().replace(/'/g, '')) {

        // 🚨 2. Y열 시스템 로그의 타임스탬프 검사 ([YYYY-MM-DD HH:mm] 형태 추출) — 승인 대기(예약확인중) 건만 만료 대상
        const sysLog = String(sheet.getRange(safeRow, COL_MASTER.SYS_LOG).getValue() || '');
        const logMatch = sysLog.match(/(\d{4}-\d{2}-\d{2}\s+\d{2}:\d{2})/);

        if (logMatch && currentStatus === '예약확인중') {
          const sentTime = new Date(logMatch[1].replace(/-/g, '/'));
          const now = new Date();
          const diffHours = (now.getTime() - sentTime.getTime()) / (1000 * 60 * 60);
          
          // 24시간 초과 시 만료 처리
          if (diffHours >= 24) {
            isExpired = true;
            
            // ① Master_Log 슬롯 회수 및 상태 취소 변경
            sheet.getRange(safeRow, COL_MASTER.STATUS).setValue('취소(기한만료)');
            sheet.getRange(safeRow, COL_MASTER.VISIT, 1, MASTER_SLOT_COLS).clearContent(); // H~J열 슬롯 반환
            sheet.getRange(safeRow, COL_MASTER.SYS_LOG).setValue(sysLog + " | [시스템] 점주 24시간 미답변 만료");

            // ② 🎯 Restaurant_List Z열(26번째 열) 무응답 카운트 +1 누적
            if (restSheet && storeId) {
              const restData = restSheet.getDataRange().getValues();
              for (let k = REST_DATA_START_INDEX; k < restData.length; k++) {
                if (String(restData[k][IDX_REST.ID] || '').trim().toUpperCase() === storeId) {
                  const currentCount = parseInt(restData[k][IDX_REST.NORESPONSE_COUNT], 10) || 0; // Z열(index 25)
                  restSheet.getRange(k + 1, COL_REST.NORESPONSE_COUNT).setValue(currentCount + 1);
                  break;
                }
              }
            }
          }
        }

        // 🔒 [중복 발송 방어 락] 기한 내이고 최초 1회 승인일 때만 메일 엔진 동작
        const isFirstConfirm = (currentStatus !== '방문전');

        if (!isExpired) {
          sheet.getRange(safeRow, COL_MASTER.STATUS).setValue('방문전'); 
          
          // 🎯 [G열 비우기] Y열(25열)에 기존 로그 보존하며 누적 기록
          const prevLogY = String(sheet.getRange(safeRow, COL_MASTER.SYS_LOG).getValue() || '').trim();
          const confirmTime = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm");
          const appendLog = `[${confirmTime}] 점주_직접_링크확정`;
          sheet.getRange(safeRow, COL_MASTER.SYS_LOG).setValue(prevLogY ? `${prevLogY} | ${appendLog}` : appendLog);

          if (isFirstConfirm) {
            try {
              const currentRestaurantName = String(sheet.getRange(safeRow, COL_MASTER.STORE_NAME).getValue() || '').trim(); 
              const rawMemberCode = String(sheet.getRange(safeRow, COL_MASTER.MEMBER).getValue() || '').trim();
              const currentMemberCode = rawMemberCode.replace(/['"\s]/g, '').toLowerCase();
              
              let storeNameJp = "";
              if (restSheet && storeId) {
                const restData = restSheet.getDataRange().getValues();
                for (let k = REST_DATA_START_INDEX; k < restData.length; k++) {
                  if (String(restData[k][IDX_REST.ID] || '').trim().toUpperCase() === storeId) {
                    storeNameJp = String(restData[k][IDX_REST.NAME_JP] || '').trim();
                    break;
                  }
                }
              }
              const finalStoreNameJp = storeNameJp || currentRestaurantName;

              const timeZone = Session.getScriptTimeZone();
              const rawVisitDate = sheet.getRange(safeRow, COL_MASTER.VISIT).getValue();
              const rawPeopleStr = String(sheet.getRange(safeRow, COL_MASTER.PEOPLE).getValue() || '1');
              const pCount = rawPeopleStr.replace(/[^0-9]/g, '') || '1';
              const rawVisitStr = (rawVisitDate instanceof Date) ? Utilities.formatDate(rawVisitDate, timeZone, 'yyyy-MM-dd HH:mm') : String(rawVisitDate || '-');
              const visitDateStr = _visitDisplayForRow_(sheet, safeRow, rawVisitStr);           // 운영진(일본어)
              const creatorVisitStr = _visitDisplayForRow_(sheet, safeRow, rawVisitStr, 'ko');   // 크리에이터(한국어)

              if (currentMemberCode) {
                // ✉️ 크리에이터 확정 안내 메일
                const creatorEmail = getCreatorEmail(currentMemberCode);
                if (creatorEmail.includes("@")) {
                  sendCreatorConfirmedMail({ to: creatorEmail, storeName: currentRestaurantName, visitDateStr: creatorVisitStr, peopleCount: pCount });
                }

                // ✉️ 관리자 알림 발송
                try {
                  sendAdminBookingConfirmedMail({ orderNo: orderNo, storeNameJp: finalStoreNameJp, memberCode: currentMemberCode, visitDateStr: visitDateStr, peopleCount: pCount });
                } catch(e) {}
              }
            } catch (mailErr) {
              console.error("❌ doGet 메일 엔진 연산 실패: " + mailErr.toString());
            }
          }
        }
      }
    } catch(err) {
      console.error("❌ store_confirm 코어 에러: " + err.toString());
    }

  return { isExpired: isExpired, isAlreadyRequested: isAlreadyRequested };
}

/** 🔒 일정변경 요청 화면 진입 가능 여부 (24시간 만료 및 취소/노쇼/조율 중 건 차단) → true면 차단 */
function isFeedbackBlocked(params) {
  const safeRow = parseInt(params?.row, 10);
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET.MASTER);

  if (safeRow && sheet) {
    const currentStatus = String(sheet.getRange(safeRow, COL_MASTER.STATUS).getValue() || '').trim();
    const sysLog = String(sheet.getRange(safeRow, COL_MASTER.SYS_LOG).getValue() || '');
    const logMatch = sysLog.match(/(\d{4}-\d{2}-\d{2}\s+\d{2}:\d{2})/);

    let isExpired = false;
    if (logMatch) {
      const sentTime = new Date(logMatch[1].replace(/-/g, '/'));
      const now = new Date();
      const diffHours = (now.getTime() - sentTime.getTime()) / (1000 * 60 * 60);
      if (diffHours >= 24) isExpired = true;
    }

    return isExpired || currentStatus.includes('취소') || currentStatus.includes('노쇼') || currentStatus === '일정조율필요';
  }
  return false;
}

/*****************************************************************************************
 * ⚡ [엔진 완전 통합] 점주 일정 변경 리퀘스트 처리 및 크리에이터+운영진 동시 노티 엔진
 ****************************************************************************************/
function submitStoreFeedback(row, orderNo, msg) {
  try {
    const safeRow = parseInt(row, 10);
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(SHEET.MASTER);
    
    // 데이터(A~N열까지 총 14칸) 캐싱
    const rowData = sheet.getRange(safeRow, COL_MASTER.ORDER, 1, COL_MASTER.VISITED).getValues()[0];
    
    const sheetOrderNo = String(rowData[IDX_MASTER.ORDER]).trim().replace(/'/g, '');
    const cleanOrderNo = String(orderNo).trim().replace(/'/g, '');
    
    if (!safeRow || sheetOrderNo !== cleanOrderNo) {
      return { success: false, error: "올바르지 않은 요청 정보입니다." };
    }
    
    // 1️⃣ L열(12번째): 진행 상태를 '일정조율필요'로 변경
    sheet.getRange(safeRow, COL_MASTER.STATUS).setValue('일정조율필요');
    
    // 2️⃣ H열~J열 슬롯 영역 자동 리셋 처리
    sheet.getRange(safeRow, COL_MASTER.VISIT, 1, MASTER_SLOT_COLS).clearContent();
    
    // 3️⃣ M열(13번째): 피드백 기록
    const timestamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "MM/dd HH:mm");
    const currentNote = String(rowData[IDX_MASTER.FEEDBACK] || '').trim();
    const newNote = (currentNote ? currentNote + "\n" : "") + `[점주요청 ${timestamp}] ${msg}`;
    sheet.getRange(safeRow, COL_MASTER.FEEDBACK).setValue(newNote); 
    
    // 4️⃣ 변수 바인딩
    const storeNameKo = String(rowData[IDX_MASTER.STORE_NAME]).trim();      // F열 기본 한국어 매장명
    const creatorName = String(rowData[IDX_MASTER.NAME]).trim();      // C열 크리에이터 성함
    const currentMemberCode = String(rowData[IDX_MASTER.MEMBER]).trim(); // B열 크리에이터 코드

    // Restaurant_List 시트 3행(k=REST_DATA_START_INDEX)부터 대조하여 일본어 매장명(storeNameJp) 파싱
    let storeNameJp = "";
    const restSheet = ss.getSheetByName(SHEET.REST);
    if (restSheet) {
      const restData = restSheet.getDataRange().getValues();
      const storeId = String(rowData[IDX_MASTER.STORE_ID] || '').trim().toUpperCase(); // E열: 매장 ID
      
      for (let k = REST_DATA_START_INDEX; k < restData.length; k++) {
        if (String(restData[k][IDX_REST.ID]).trim().toUpperCase() === storeId) {
          storeNameJp = String(restData[k][IDX_REST.NAME_JP] || '').trim(); // C열: 일본어 매장명
          break;
        }
      }
    }
    const finalStoreName = storeNameJp ? storeNameJp : storeNameKo;

    // ✉️ [1번 슈팅] 내부 운영진 담당자 일본어 알림 메일 발송
    const adminEmail = getAdminAlertEmails();


    if (adminEmail && adminEmail.includes('@')) {
      try {
        sendAdminRescheduleRequestMail({ to: adminEmail, storeName: finalStoreName, creatorName: creatorName, memberCode: currentMemberCode, orderNo: cleanOrderNo, msg: msg, row: safeRow });
        sheet.getRange(safeRow, COL_MASTER.SYS_LOG).setValue("운영진알림 발송완료"); 
      } catch (adminMailErr) {
        console.error("❌ 운영진 메일 알림 실패: " + adminMailErr.toString());
        sheet.getRange(safeRow, COL_MASTER.SYS_LOG).setValue("❌ 운영진알림 실패: " + adminMailErr.toString());
      }
    } else {
      sheet.getRange(safeRow, COL_MASTER.SYS_LOG).setValue("❌ 실패: 운영진 이메일 주소 포맷 오류");
    }

    // ✉️ [2번 슈팅] 크리에이터향 알림 이메일 발송
    if (currentMemberCode) {
      const creatorEmail = getCreatorEmail(currentMemberCode);

      if (creatorEmail.includes("@")) {
        const timeZone = Session.getScriptTimeZone();
        const rawVisitDate = rowData[IDX_MASTER.VISIT]; 
        const pCount = rowData[IDX_MASTER.PEOPLE] ? String(rowData[IDX_MASTER.PEOPLE]).replace(/[^0-9]/g, '') : '1'; 
        const visitDateStr = _visitDisplayForRow_(sheet, safeRow, (rawVisitDate instanceof Date) ? Utilities.formatDate(rawVisitDate, timeZone, 'yyyy-MM-dd HH:mm') : String(rawVisitDate || '-'), 'ko');

        try {
          sendCreatorRescheduleRequestMail({ to: creatorEmail, storeNameKo: storeNameKo, visitDateStr: visitDateStr, peopleCount: pCount });
          const prevLog = String(sheet.getRange(safeRow, COL_MASTER.SYS_LOG).getValue() || '');
          sheet.getRange(safeRow, COL_MASTER.SYS_LOG).setValue(prevLog.includes("완료") ? "조율 노티 전체메일 발송완료" : "크리에이터 노티 메일 발송완료");
        } catch (creatorMailErr) {
          console.error("❌ 크리에이터 변경 안내 메일 실패: " + creatorMailErr.toString());
          const prevLog = String(sheet.getRange(safeRow, COL_MASTER.SYS_LOG).getValue() || '');
          sheet.getRange(safeRow, COL_MASTER.SYS_LOG).setValue(prevLog + " | ❌ 크리에이터메일 실패: " + creatorMailErr.toString());
        }
      }
    }
    return { success: true };
  } catch (e) { return { success: false, error: e.toString() }; }
}

/*****************************************************************************************
 * 🎯 [파트너센터 로그인 & 대시보드 연동 코어 엔진]
 ****************************************************************************************/

/** 점주 토큰 검증 + 토큰 주체(storeId)와 요청 대상 매장 일치 확인. 실패 시 Error('AUTH_REQUIRED') */
function requireStoreAccess_(token, storeId) {
  const subject = requireToken_(token, 'store');
  if (String(subject).toUpperCase() !== String(storeId || '').trim().toUpperCase()) throw new Error('AUTH_REQUIRED');
}

/** [1] 매장 고유 PIN 번호로 로그인 검증 → 성공 시 { storeId, token } */
function loginStoreByPin(pin) {
  try {
    const pinClean = String(pin).trim().toUpperCase();
    if (!pinClean) return { success: false, error: "コードを入力してください。" };
    if (isLoginLocked_('store')) return { success: false, error: "試行回数が上限に達しました。しばらくしてからもう一度お試しください。" };

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const restSheet = ss.getSheetByName(SHEET.REST);
    if (!restSheet) return { success: false, error: "시스템 오류: Restaurant_List 시트가 없습니다." };

    const restData = restSheet.getDataRange().getValues();
    
    // 3행부터 순회
    for (let i = REST_DATA_START_INDEX; i < restData.length; i++) {
      const sheetPin = String(restData[i][IDX_REST.PIN] || '').trim().toUpperCase();
      if (sheetPin === pinClean) {
        const storeId = String(restData[i][IDX_REST.ID]).trim(); 
        clearLoginFailures_('store');
        return { success: true, storeId: storeId, token: issueToken_('store', storeId) };
      }
    }
    
    recordLoginFailure_('store');
    return { success: false, error: "無効なパートナーコードです。再度コードを確認してください。" };
  } catch (e) {
    return { success: false, error: "서버 오류: " + e.toString() };
  }
}

/** [2] 특정 매장의 대시보드 스탯 및 할당 크리에이터 미션 리스트 조회 */
function getStoreData(storeId, token) {
  requireStoreAccess_(token, storeId);
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const data = _getSheetsData([SHEET.MASTER, SHEET.REST, SHEET.USER]);
    
    let storeName = storeId;
    let storeNameJp = "";
    
    const restList = data[SHEET.REST];
    // 3행부터 순회
    for (let k = REST_DATA_START_INDEX; k < restList.length; k++) {
      if (String(restList[k][IDX_REST.ID]).trim().toUpperCase() === String(storeId).toUpperCase()) {
        storeName = String(restList[k][IDX_REST.NAME_KO] || '').trim(); 
        storeNameJp = String(restList[k][IDX_REST.NAME_JP] || '').trim(); 
        break;
      }
    }
    const finalStoreName = storeNameJp ? storeNameJp : storeName;

    const userTierMap = {};
    const userDB = data[SHEET.USER];
    for (let u = 1; u < userDB.length; u++) {
      if (userDB[u][IDX_USER.KEY]) {
        let emoji = '🟡';
        try { emoji = parseTierEmoji(userDB[u][IDX_USER.GROUP]); } catch(ex) {}
        userTierMap[String(userDB[u][IDX_USER.KEY]).trim()] = emoji;
      }
    }

    let stats = { total: 0, pending: 0, verified: 0, submitted: 0 };
    let storeMissions = [];
    const timeZone = Session.getScriptTimeZone();
    const masterLog = data[SHEET.MASTER];

    const webAppUrl = getActiveWebAppUrl();

    for (let i = masterLog.length - 1; i >= 2; i--) {
      const mRow = masterLog[i];
      const mStoreId = String(mRow[IDX_MASTER.STORE_ID] || '').trim().toUpperCase(); 
      const status = String(mRow[IDX_MASTER.STATUS] || '').trim(); 

      if (!mRow[IDX_MASTER.ORDER] || status.includes('취소') || mStoreId !== String(storeId).toUpperCase()) continue;

      stats.total++;
      if (status === '방문전' || status === '예약확인중' || status === '일정조율필요' || status === '예약대기') stats.pending++;
      else if (status === '방문완료') stats.verified++;
      else if (status === '제출완료') stats.submitted++;

      const formatDate = (val, format) => (val instanceof Date) ? Utilities.formatDate(val, timeZone, format) : String(val || '-');
      const memberCode = String(mRow[IDX_MASTER.MEMBER]).trim();
      const orderNoRaw = String(mRow[IDX_MASTER.ORDER]).replace(/'/g, ''); 

      storeMissions.push({
        row: i + 1,
        orderNo: orderNoRaw,
        memberCode: memberCode,
        member: String(mRow[IDX_MASTER.NAME] || '').trim(), 
        tierEmoji: userTierMap[memberCode] || '🟡',
        visitDate: formatDate(mRow[IDX_MASTER.VISIT], 'yyyy-MM-dd HH:mm'), 
        peopleCount: mRow[IDX_MASTER.PEOPLE] ? String(mRow[IDX_MASTER.PEOPLE]).replace(/[^0-9]/g, '') : '1', 
        status: status,
        submitDate: formatDate(mRow[IDX_MASTER.SUBMIT_DATE], 'yyyy-MM-dd'), 
        reviewUrl: String(mRow[IDX_MASTER.REVIEW_URL] || ''), 
        googleMapUrl: String(mRow[IDX_MASTER.GOOGLE_URL] || ''), 
        shortReview: String(mRow[IDX_MASTER.SHORT_REVIEW_JP] || mRow[IDX_MASTER.SHORT_REVIEW] || '')
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
function getStoreConfig(storeId, token) {
  requireStoreAccess_(token, storeId);
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(SHEET.REST);
    const data = sheet.getDataRange().getValues();

    // 3행부터 순회
    for (let i = REST_DATA_START_INDEX; i < data.length; i++) {
      if (String(data[i][IDX_REST.ID]).trim().toUpperCase() === String(storeId).toUpperCase()) {
        return {
          success: true,
          row: i + 1,
          storeId: String(data[i][IDX_REST.ID]).trim(),
          email: String(data[i][IDX_REST.EMAIL] || '').trim(),      
          storeType: String(data[i][IDX_REST.STORE_TYPE] || '').trim(),   
          maxPerSlot: data[i][IDX_REST.MAX_PER_SLOT] ? parseInt(data[i][IDX_REST.MAX_PER_SLOT], 10) : 1, 
          dailyCap: data[i][IDX_REST.DAILY_CAP] ? parseInt(data[i][IDX_REST.DAILY_CAP], 10) : 999,  
          blackouts: String(data[i][IDX_REST.BLACKOUTS] || '').trim(),   
          bizHours: String(data[i][IDX_REST.BIZ_HOURS] || '').trim(),    
          intervalMin: data[i][IDX_REST.INTERVAL_MIN] ? String(data[i][IDX_REST.INTERVAL_MIN]).trim() : "30", 
          blockedTimes: String(data[i][IDX_REST.BLOCKED_TIMES] || '').trim(),
          ...(String(data[i][IDX_REST.STORE_TYPE] || '').trim().toUpperCase() === STORE_TYPE_STAY ? _getStayConfig_(data[i]) : {})
        };
      }
    }
    return { success: false, error: "店舗設定情報が見つかりません。" };
  } catch (e) {
    return { success: false, error: e.toString() };
  }
}

/** [4] 매장 예약 환경설정 파일 수정 저장하기 */
function updateStoreConfig(token, row, email, storeType, maxPerSlot, dailyCap, blackouts, bizHours, intervalMin, blockedTimes, rooms, checkinTime, checkoutTime, maxNights) {
  try {
    const safeRow = parseInt(row, 10);
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET.REST);
    // 토큰 주체(매장)의 행만 수정 가능
    requireStoreAccess_(token, sheet.getRange(safeRow, COL_REST.ID).getValue());
    
    sheet.getRange(safeRow, COL_REST.EMAIL).setValue(email);          
    sheet.getRange(safeRow, COL_REST.STORE_TYPE).setValue(storeType);        
    sheet.getRange(safeRow, COL_REST.MAX_PER_SLOT).setValue(parseInt(maxPerSlot, 10) || 1);  
    sheet.getRange(safeRow, COL_REST.DAILY_CAP).setValue(parseInt(dailyCap, 10) || 999);   
    sheet.getRange(safeRow, COL_REST.BLACKOUTS).setValue(blackouts);     
    sheet.getRange(safeRow, COL_REST.BIZ_HOURS).setValue(bizHours);      
    sheet.getRange(safeRow, COL_REST.INTERVAL_MIN).setValue(parseInt(intervalMin, 10) || 30); 
    sheet.getRange(safeRow, COL_REST.BLOCKED_TIMES).setValue(blockedTimes);

    // 숙박 전용 항목 (객실수/체크인·아웃 시간/최대 박수). 시간은 HH:mm 형식일 때만 저장
    if (String(storeType || '').trim().toUpperCase() === STORE_TYPE_STAY) {
      const isHm = t => /^\d{1,2}:\d{2}$/.test(String(t || '').trim());
      sheet.getRange(safeRow, COL_REST.ROOMS).setValue(parseInt(rooms, 10) || 1);
      if (isHm(checkinTime)) sheet.getRange(safeRow, COL_REST.CHECKIN_TIME).setValue("'" + String(checkinTime).trim());
      if (isHm(checkoutTime)) sheet.getRange(safeRow, COL_REST.CHECKOUT_TIME).setValue("'" + String(checkoutTime).trim());
      sheet.getRange(safeRow, COL_REST.MAX_NIGHTS).setValue(parseInt(maxNights, 10) || 7);
    }

    return { success: true };
  } catch (e) {
    return { success: false, error: e.toString() };
  }
}

/** [5] 점주가 대시보드에서 즉시 예약을 확정(승인)하는 함수 */
function storeDirectConfirm_(row, orderNo) {
  try {
    const safeRow = parseInt(row, 10);
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET.MASTER);
    const sheetOrderNo = String(sheet.getRange(safeRow, COL_MASTER.ORDER).getValue()).trim().replace(/'/g, '');
    
    if (sheetOrderNo !== String(orderNo).trim().replace(/'/g, '')) {
      return { success: false, error: "주문번호가 일치하지 않습니다." };
    }
    
    sheet.getRange(safeRow, COL_MASTER.STATUS).setValue('방문전');
    
    // 🎯 [G열 비우기] Y열(25열)에 기존 로그 보존하며 누적 기록
    const prevLogY = String(sheet.getRange(safeRow, COL_MASTER.SYS_LOG).getValue() || '').trim();
    const directTime = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm");
    const appendLog = `[${directTime}] 점주_대시보드_직접확정`;
    sheet.getRange(safeRow, COL_MASTER.SYS_LOG).setValue(prevLogY ? `${prevLogY} | ${appendLog}` : appendLog);

    try {
      const currentRestaurantName = String(sheet.getRange(safeRow, COL_MASTER.STORE_NAME).getValue() || '').trim(); 
      const rawMemberCode = String(sheet.getRange(safeRow, COL_MASTER.MEMBER).getValue() || '').trim();
      const currentMemberCode = rawMemberCode.replace(/['"\s]/g, '').toLowerCase();
      
      const timeZone = Session.getScriptTimeZone();
      const rawVisitDate = sheet.getRange(safeRow, COL_MASTER.VISIT).getValue();
      const pCount = sheet.getRange(safeRow, COL_MASTER.PEOPLE).getValue() ? String(sheet.getRange(safeRow, COL_MASTER.PEOPLE).getValue()).replace(/[^0-9]/g, '') : '1'; 
      const visitDateStr = _visitDisplayForRow_(sheet, safeRow, (rawVisitDate instanceof Date) ? Utilities.formatDate(rawVisitDate, timeZone, 'yyyy-MM-dd HH:mm') : String(rawVisitDate || '-'), 'ko');

      if (currentMemberCode) {
        const creatorEmail = getCreatorEmail(currentMemberCode);

        if (creatorEmail.includes("@")) {
          try {
            sendCreatorConfirmedMail({ to: creatorEmail, storeName: currentRestaurantName, visitDateStr: visitDateStr, peopleCount: pCount });
            sheet.getRange(safeRow, COL_MASTER.SYS_LOG).setValue("크리에이터확정메일 발송완료"); 
          } catch (directMailErr) {
            console.error("❌ 점주 직접 확정 메일 발송 실패: " + directMailErr.toString());
            sheet.getRange(safeRow, COL_MASTER.SYS_LOG).setValue("❌ 크리에이터확정메일 실패: " + directMailErr.toString());
          }
        }
      }
    } catch (mailErr) {}
    return { success: true };
  } catch (e) { return { success: false, error: e.toString() }; }
}

/** [6] 점주가 대시보드에서 일정 변경 팝업창을 통해 리퀘스트를 보낼 때 처리 함수 */
function storeDirectFeedback(row, orderNo, msg) {
  return submitStoreFeedback(row, orderNo, msg);
}

/** [7] [간편 등록 툴] 점포 ID로 매장 기본 정보 조회 */
function findStoreSimple(storeId) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const restSheet = ss.getSheetByName(SHEET.REST);
    if (!restSheet) return { success: false, error: "Restaurant_List 시트를 찾을 수 없습니다." };

    const data = restSheet.getDataRange().getValues();
    const cleanId = String(storeId).trim().toUpperCase();

    // 3행부터 순회
    for (let i = REST_DATA_START_INDEX; i < data.length; i++) {
      if (String(data[i][IDX_REST.ID]).trim().toUpperCase() === cleanId) {
        return {
          success: true,
          data: {
            storeId: cleanId,
            nameKo: String(data[i][IDX_REST.NAME_KO] || '').trim(),
            nameJp: String(data[i][IDX_REST.NAME_JP] || '').trim(),
            email: String(data[i][IDX_REST.EMAIL] || '').trim()
          }
        };
      }
    }
    return { success: false, error: "無効な店舗コードです。" };
  } catch (e) {
    return { success: false, error: e.toString() };
  }
}

/** [8] [간편 등록 툴] 시트 K열에 이메일 저장 후 테스트 메일 즉시 발송 */
function saveStoreEmailAndSendTest(storeId, email) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const restSheet = ss.getSheetByName(SHEET.REST);
    if (!restSheet) return { success: false, error: "Restaurant_List シートを見つかりません。" };

    const data = restSheet.getDataRange().getValues();
    const cleanId = String(storeId).trim().toUpperCase();
    let rowIndex = -1;
    let storeNameJp = "";

    // 3행부터 순회
    for (let i = REST_DATA_START_INDEX; i < data.length; i++) {
      if (String(data[i][IDX_REST.ID]).trim().toUpperCase() === cleanId) {
        rowIndex = i + 1;
        storeNameJp = String(data[i][IDX_REST.NAME_JP] || '').trim() || String(data[i][IDX_REST.NAME_KO] || '').trim();
        break;
      }
    }

    if (rowIndex === -1) return { success: false, error: "店舗コードが見つかりません。" };

    // 시트 K열(11번째 열) 이메일 갱신
    restSheet.getRange(rowIndex, COL_REST.EMAIL).setValue(email.trim());

    // 테스트 메일 발송
    sendStoreConnectionTestMail({ to: email.trim(), storeId: cleanId, storeNameJp: storeNameJp, addressLabel: '登録アドレス' });

    return { success: true };
  } catch (err) {
    return { success: false, error: err.toString() };
  }
}
