/** 메일 템플릿·발송 (C단계에서 전 메일 로직을 이 파일로 일원화) */

/** 메일 발송 분리 함수 (주문번호 제거, 5대 항목 고정, 24시간 답변기한 경고, 깨짐 방지 table 구조, Y열 타임스탬프 기록) */
function _sendBookingNotifications(p) {
  const data = _getSheetsData([SHEET.REST, SHEET.USER]);
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET.MASTER);

  let storeEmail = "", storeNameJp = "", storeNameKo = "", storeBenefit = "";
  for (let i = 2; i < data[SHEET.REST].length; i++) {
    if (String(data[SHEET.REST][i][IDX_REST.ID] || '').trim().toUpperCase() === p.storeId.toUpperCase()) {
      storeNameKo = String(data[SHEET.REST][i][IDX_REST.NAME_KO] || '').trim();
      storeNameJp = String(data[SHEET.REST][i][IDX_REST.NAME_JP] || '').trim() || storeNameKo;
      storeEmail = String(data[SHEET.REST][i][IDX_REST.EMAIL] || '').trim();
      storeBenefit = String(data[SHEET.REST][i][IDX_REST.BENEFIT] || '').trim();
      break;
    }
  }

  let creatorProfileUrl = "";
  for (let u = 1; u < data[SHEET.USER].length; u++) {
    if (String(data[SHEET.USER][u][IDX_USER.KEY] || '').trim().toLowerCase() === p.memberCode.toLowerCase()) {
      creatorProfileUrl = String(data[SHEET.USER][u][IDX_USER.BLOG] || '').trim() || 
                          String(data[SHEET.USER][u][IDX_USER.INSTA] || '').trim() || 
                          String(data[SHEET.USER][u][IDX_USER.YOUTUBE] || '').trim() || 
                          String(data[SHEET.USER][u][IDX_USER.TIKTOK] || '').trim() || 
                          String(data[SHEET.USER][u][IDX_USER.GOOGLE_GUIDE] || '').trim() || "";
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
      sheet.getRange(p.safeRow, COL_MASTER.SYS_LOG).setValue(`[${sentTimeStr}] 점주메일 발송완료`);
    } catch (mailErr) {
      console.error("점주 이메일 발송 실패: " + mailErr.toString());
      sheet.getRange(p.safeRow, COL_MASTER.SYS_LOG).setValue("❌ 점주메일 실패: " + mailErr.toString());
    }
  } else {
    sheet.getRange(p.safeRow, COL_MASTER.SYS_LOG).setValue("❌ 실패: 점주 이메일 주소 없음");
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
