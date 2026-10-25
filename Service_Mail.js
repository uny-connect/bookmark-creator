/** 메일 템플릿·발송 — 모든 메일은 이 파일의 send*Mail 함수로만 발송한다. (호출부는 조회/로그 기록만 담당) */

const _DEFAULT_BENEFIT_JP = "店舗指定のクリエイター向け提供メニュー";

/** 공통 발송 래퍼. 실패 시 예외를 던지므로 호출부가 로그/UI 처리를 한다. */
function _sendMail(to, subject, htmlBody, opts) {
  const options = Object.assign({ htmlBody: htmlBody, name: MAIL.NAME, from: MAIL.FROM }, opts || {});
  if (!options.from) delete options.from; // 별칭 없이 기본 발신 주소로 발송
  GmailApp.sendEmail(to, subject, "", options);
}

/** 멤버코드 비교용 정규화 (따옴표/공백 제거 + 소문자) */
function _normMemberCode(v) {
  return String(v || '').trim().replace(/['"\s]/g, '').toLowerCase();
}

/** User_DB 행 배열에서 크리에이터 이메일 조회 (없으면 '') */
function findCreatorEmail(userRows, memberCode) {
  const key = _normMemberCode(memberCode);
  if (!key) return '';
  for (let j = 1; j < userRows.length; j++) {
    if (_normMemberCode(userRows[j][IDX_USER.KEY]) === key) return String(userRows[j][IDX_USER.EMAIL] || '').trim();
  }
  return '';
}

/** User_DB 를 읽어 크리에이터 이메일 조회 */
function getCreatorEmail(memberCode) {
  const userSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET.USER);
  return userSheet ? findCreatorEmail(userSheet.getDataRange().getValues(), memberCode) : '';
}

/** User_DB 행 배열에서 SNS/프로필 URL 1개 선택 (블로그 > 인스타 > 유튜브 > 틱톡 > 구글 가이드) */
function findCreatorProfileUrl(userRows, memberCode) {
  const key = String(memberCode || '').trim().toLowerCase();
  for (let u = 1; u < userRows.length; u++) {
    if (String(userRows[u][IDX_USER.KEY] || '').trim().toLowerCase() === key) {
      const r = userRows[u];
      return String(r[IDX_USER.BLOG] || '').trim() ||
             String(r[IDX_USER.INSTA] || '').trim() ||
             String(r[IDX_USER.YOUTUBE] || '').trim() ||
             String(r[IDX_USER.TIKTOK] || '').trim() ||
             String(r[IDX_USER.GOOGLE_GUIDE] || '').trim() || "";
    }
  }
  return "";
}

/** 프로필 URL → 점주 메일용 링크 HTML / 운영진 메일용 SNS 행 HTML */
function _creatorProfileLinks(creatorProfileUrl) {
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
  return { profileHtmlClean: profileHtmlClean, adminSnsHtml: adminSnsHtml };
}

/*********************** 크리에이터 수신 ***********************/

/** ✉️ 예약 확정 안내 (점주 링크 확정 / 점주 대시보드 확정 / 어드민 강제 승인 공용) */
function sendCreatorConfirmedMail(p) {
  const storeName = p.storeName, visitDateStr = p.visitDateStr, peopleCount = p.peopleCount;
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
                            <td style="color: #1A2B49; font-weight: 700; padding: 6px 0;">${storeName}</td>
                          </tr>
                          <tr>
                            <td style="color: #8b95a1; font-weight: 700; padding: 6px 0;">예약 일시</td>
                            <td style="color: #1a73e8; font-weight: 700; padding: 6px 0;">${visitDateStr}</td>
                          </tr>
                          <tr>
                            <td style="color: #8b95a1; font-weight: 700; padding: 6px 0;">방문 인원</td>
                            <td style="color: #495057; font-weight: 700; padding: 6px 0;">${peopleCount}명</td>
                          </tr>
                        </table>
                      </div>
                      <div style="text-align: center;">
                        <a href="${MAIL.KAKAO_URL}" target="_blank" style="display: block; background: #1A2B49; color: #ffffff; font-size: 14px; font-weight: 700; text-decoration: none; padding: 14px; border-radius: 12px; box-shadow: 0 4px 12px rgba(26,43,73,0.15);">카카오톡 채팅하기</a>
                      </div>
                    </div>`;
  _sendMail(p.to, subject, htmlBody);
}

/** ✉️ 어드민이 '일정조율필요'로 변경했을 때 안내 */
function sendCreatorRescheduleNoticeMail(p) {
  const subject = "[BOOKMARK CREATORS] 일정 조율 요청 안내";
  const htmlBody = `
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
                  <a href="${MAIL.KAKAO_URL}" target="_blank" style="display: block; background: #e03131; color: #ffffff; font-size: 14px; font-weight: 700; text-decoration: none; padding: 14px; border-radius: 12px; box-shadow: 0 4px 12px rgba(224,49,49,0.15);">담당자에게 미리 연락하기</a>
                </div>
              </div>`;
  _sendMail(p.to, subject, htmlBody);
}

/** ✉️ 점주가 일정 변경을 요청했을 때 안내 */
function sendCreatorRescheduleRequestMail(p) {
  const storeNameKo = p.storeNameKo, visitDateStr = p.visitDateStr, peopleCount = p.peopleCount;
  const subject = "[BOOKMARK CREATORS] 방문 일정 변경 조율 안내";
  const htmlBody = `
            <meta charset="UTF-8">
            <div style="background: #f4f5f7; padding: 24px 10px; font-family: 'Apple SD Gothic Neo', 'Malgun Gothic', sans-serif;">
              <div style="max-width: 500px; margin: 0 auto; padding: 32px 24px; background: #ffffff; border: 1px solid #eef0f2; border-radius: 14px; box-shadow: 0 4px 12px rgba(0,0,0,0.03);">
                
                <div style="margin-bottom: 20px; text-align: left;">
                  <span style="font-size: 11px; font-weight: 800; letter-spacing: 1px; color: #ffffff; background: #e03131; padding: 4px 10px; border-radius: 6px; display: inline-block;">NOTICE</span>
                  <h2 style="font-size: 20px; font-weight: 800; color: #1A2B49; margin: 10px 0 0 0;">BOOKMARK CREATORS</h2>
                </div>
                
                <div style="border-top: 2px solid #e03131; padding-top: 20px; margin-bottom: 20px;">
                  <p style="font-size: 14.5px; font-weight: 800; color: #e03131; margin: 0 0 10px 0;">[방문 일정 변경 요청 알림]</p>
                  <p style="font-size: 13.5px; line-height: 1.65; color: #495057; margin: 0;">
                    안녕하세요 크리에이터님, 신청하신 매장의 사정으로 인해 <b>일정 조율이 필요한 상태</b>로 변경되었습니다.<br><br>
                    신청 내역을 바탕으로 운영 담당자가 빠르게 연락드려 일정 재조율을 도와드리겠습니다.
                  </p>
                </div>
                
                <div style="background: #f8f9fa; border: 1px solid #e9ecef; border-radius: 10px; padding: 16px 18px; margin-bottom: 24px;">
                  <table style="width: 100%; border-collapse: collapse; font-size: 13.5px;">
                    <tr>
                      <td style="width: 85px; color: #8b95a1; font-weight: 700; padding: 6px 0;">방문 매장</td>
                      <td style="color: #1A2B49; font-weight: 700; padding: 6px 0;">${storeNameKo}</td>
                    </tr>
                    <tr>
                      <td style="color: #8b95a1; font-weight: 700; padding: 6px 0;">기존 일시</td>
                      <td style="color: #e03131; font-weight: 800; padding: 6px 0;">${visitDateStr}</td>
                    </tr>
                    <tr>
                      <td style="color: #8b95a1; font-weight: 700; padding: 6px 0;">방문 인원</td>
                      <td style="color: #495057; font-weight: 700; padding: 6px 0;">${peopleCount}명</td>
                    </tr>
                  </table>
                </div>
                
                <div style="text-align: center;">
                  <a href="${MAIL.KAKAO_URL}" target="_blank" style="display: block; background: #1A2B49; color: #ffffff; font-size: 14px; font-weight: 800; text-decoration: none; padding: 13px; border-radius: 10px; box-shadow: 0 4px 12px rgba(26,43,73,0.15);">담당자에게 문의하기</a>
                </div>
              </div>
            </div>`;
  _sendMail(p.to, subject, htmlBody);
}

/*********************** 점주 수신 ***********************/

/** ✉️ 점주에게 방문 예약 확인 요청 (확정/일시변경 링크 포함, 운영진 BCC) */
function sendStoreBookingRequestMail(p) {
  const storeNameJp = p.storeNameJp, memberName = _escapeHtml(p.memberName), dateStr = p.dateStr, timeStr = p.timeStr, peopleCount = p.peopleCount;
  const links = _creatorProfileLinks(p.creatorProfileUrl);
  const profileHtmlClean = links.profileHtmlClean;
  const benefitDisplay = p.benefit || _DEFAULT_BENEFIT_JP;
  const scriptUrl = getActiveWebAppUrl();
  const confirmUrl = `${scriptUrl}?mode=store_confirm&row=${p.row}&o=${encodeURIComponent(p.orderNo)}`;
  const feedbackUrl = `${scriptUrl}?mode=feedback&row=${p.row}&o=${encodeURIComponent(p.orderNo)}`;

  const subject = `【BOOKMARK CREATORS】 ご来店予約のご確認 (${dateStr})`;
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
                    <strong>■ 訪問人数:</strong> ${peopleCount}名
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
  _sendMail(p.to, subject, htmlBody, { bcc: getAdminAlertEmails() });
}

/** ✉️ 점주에게 예약 취소 통지 (운영진 BCC) */
function sendStoreCancellationMail(p) {
  const storeNameJp = p.storeNameJp, orderNo = p.orderNo, creatorName = _escapeHtml(p.creatorName), visitDisplay = p.visitDisplay, peopleCount = p.peopleCount;
  const subject = `【予約キャンセル】 ご来店予約キャンセルのご案内 (${storeNameJp})`;
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
            <p style="color: #1A2B49; font-size: 15.5px; font-weight: bold; margin-bottom: 4px;">${storeNameJp}</p>
            <p style="color: #6c757d; font-size: 13px; margin-top: 0;">店舗管理者様</p>
          </div>

          <p style="color: #495057; font-size: 13.5px; line-height: 1.65; margin: 16px 0;">
            いつもお世話になっております。BOOKMARK CREATORS運営事務局です。<br><br>
            予定されておりました下記の来店予約につきまして、<b>クリエイター側の都合によりキャンセルとなりました</b>のでご案内申し上げます。<br>
            お席の確保やご準備は不要となります。
          </p>

          <div style="background-color: #f8f9fa; border: 1px solid #e9ecef; border-left: 4px solid #dc3545; padding: 16px; border-radius: 8px; margin: 20px 0; font-size: 13.5px;">
            <p style="margin: 5px 0;"><strong>■ 注文番号:</strong> #${orderNo}</p>
            <p style="margin: 5px 0;"><strong>■ クリエイター名:</strong> ${creatorName}</p>
            <p style="margin: 5px 0;"><strong>■ 予約日時:</strong> <span style="color: #dc3545; font-weight: 800;">${visitDisplay}</span></p>
            <p style="margin: 5px 0;"><strong>■ 予約人数:</strong> ${peopleCount}名</p>
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
  _sendMail(p.to, subject, htmlBody, { bcc: getAdminAlertEmails() });
}

/** ✉️ 점주 메일 수신 테스트 (addressLabel: 본문의 주소 항목명) */
function sendStoreConnectionTestMail(p) {
  const storeNameJp = p.storeNameJp, storeId = p.storeId || '-', email = p.to, addressLabel = p.addressLabel;
  const todayStr = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd");
  const subject = `【TEST / 接続テスト】 BOOKMARK CREATORS 予約通知テスト送信 (${storeNameJp})`;
  const htmlBody = `
      <meta charset="UTF-8">
      <div style="font-family: 'Helvetica Neue', Arial, 'Hiragino Kaku Gothic ProN', Meiryo, sans-serif; padding: 20px; background-color: #f4f5f7;">
        <div style="max-width: 500px; margin: 0 auto; background: #ffffff; border-radius: 14px; padding: 30px; box-shadow: 0 4px 12px rgba(0,0,0,0.05); border: 1px solid #edf0f2;">
          <div style="text-align: center; margin-bottom: 20px;">
            <span style="font-size: 11px; font-weight: 800; letter-spacing: 1px; color: #ffffff; background: #1A2B49; padding: 4px 10px; border-radius: 6px; display: inline-block;">CONNECTION TEST</span>
            <h1 style="color: #1A2B49; margin: 10px 0 0 0; font-size: 22px; font-weight: 900; letter-spacing: -0.5px;">BOOKMARK CREATORS</h1>
            <div style="width: 36px; height: 3px; background: #C5A358; margin: 10px auto;"></div>
          </div>
          <h2 style="color: #1A2B49; margin-top: 0; font-size: 16px; border-bottom: 2px solid #f1f3f5; padding-bottom: 14px; text-align: center; font-weight: 800;">メール受信テスト</h2>
          <div style="margin-top: 18px;">
            <p style="color: #1A2B49; font-size: 15px; font-weight: bold; margin-bottom: 4px;">${storeNameJp}</p>
            <p style="color: #6c757d; font-size: 13px; margin-top: 0;">店舗管理者様</p>
          </div>
          <p style="color: #495057; font-size: 13.5px; line-height: 1.65;">
            本メールは、BOOKMARK CREATORSからの予約受付通知が正常に届くか確認するための<b>【テストメール】</b>です。<br>
            このメールが確認できましたら、今後のクリエイター来店予約通知も問題なく受信いただけます。
          </p>
          <div style="background-color: #f8f9fa; border: 1px solid #e9ecef; border-radius: 8px; padding: 14px 16px; margin: 20px 0; font-size: 13px;">
            <p style="margin: 4px 0;"><strong>■ 店舗ID:</strong> ${storeId}</p>
            <p style="margin: 4px 0;"><strong>■ 受信状態:</strong> <span style="color: #2D6A4F; font-weight: bold;">正常受信完了 (Success)</span></p>
            <p style="margin: 4px 0;"><strong>■ テスト日時:</strong> ${todayStr}</p>
            <p style="margin: 4px 0;"><strong>■ ${addressLabel}:</strong> ${email}</p>
          </div>
          <div style="text-align: center; margin-top: 24px; padding: 10px; background-color: #f1f3f5; border-radius: 8px; font-size: 12px; color: #6c757d; font-weight: 600;">
            ※ 本メールに対する返信や確定手続きは不要です。
          </div>
        </div>
      </div>`;
  _sendMail(p.to, subject, htmlBody);
}

/*********************** 운영진 수신 ***********************/

/** ✉️ 운영진: 크리에이터 예약 신청 접수 알림 */
function sendAdminBookingRequestMail(p) {
  const to = getAdminAlertEmails();
  if (!to) return;
  const orderNo = p.orderNo, memberName = _escapeHtml(p.memberName), memberCode = p.memberCode, storeId = p.storeId, dateStr = p.dateStr, timeStr = p.timeStr, peopleCount = p.peopleCount;
  const storeName = p.storeName, storeEmail = p.storeEmail;
  const benefitDisplay = p.benefit || _DEFAULT_BENEFIT_JP;
  const adminSnsHtml = _creatorProfileLinks(p.creatorProfileUrl).adminSnsHtml;
  const adminSubject = `【予約受付】${p.memberName}様 ➔ ${storeName} (${dateStr} ${timeStr})`;
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
            <p style="margin: 6px 0;"><b>・注文番号:</b> #${orderNo}</p>
            <p style="margin: 6px 0;"><b>・クリエイター:</b> ${memberName} <span style="color:#8b95a1; font-size: 12px;">(${memberCode})</span></p>
            <p style="margin: 6px 0;"><b>・訪問店舗:</b> ${storeName} <span style="color:#8b95a1; font-size: 12px;">(${storeId})</span></p>
            <p style="margin: 6px 0;"><b>・提供内容:</b> <span style="color: #2D6A4F; font-weight: bold;">${benefitDisplay}</span></p>
            <p style="margin: 6px 0;"><b>・予約日時:</b> <span style="color: #d63384; font-weight: bold;">${dateStr} ${timeStr}</span></p>
            <p style="margin: 6px 0;"><b>・訪問人数:</b> <span style="font-weight: bold;">${peopleCount}人</span></p>
            ${adminSnsHtml}
            <p style="margin: 6px 0;"><b>・店舗受信メール:</b> <span style="color: #1a73e8;">${storeEmail || '未登録'}</span></p>
          </div>
        </div>
      `;
  _sendMail(to, adminSubject, adminHtml, { name: MAIL.NOTI_NAME, from: null });
}

/** ✉️ 운영진: 점주가 예약을 확정했음을 알림 */
function sendAdminBookingConfirmedMail(p) {
  const to = getAdminAlertEmails();
  if (!to) return;
  const orderNo = p.orderNo, storeNameJp = p.storeNameJp, memberCode = p.memberCode, visitDateStr = p.visitDateStr, peopleCount = p.peopleCount;
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
                          <p style="margin: 6px 0;"><b>・店舗名:</b> ${storeNameJp}</p>
                          <p style="margin: 6px 0;"><b>・クリエイター:</b> ${memberCode}</p>
                          <p style="margin: 6px 0;"><b>・予約日時:</b> <span style="color: #2D6A4F; font-weight: bold;">${visitDateStr} (${peopleCount}名)</span></p>
                        </div>
                      </div>
                    `;
  _sendMail(to, adminSubject, adminHtml, { name: MAIL.NOTI_NAME, from: null });
}

/** ✉️ 운영진: 점주가 일정 변경을 요청했음을 알림 */
function sendAdminRescheduleRequestMail(p) {
  const storeName = p.storeName, creatorName = _escapeHtml(p.creatorName), memberCode = p.memberCode, orderNo = p.orderNo, msg = _escapeHtml(p.msg), row = p.row;
  const alertSubject = `【日程変更要請】[${storeName}] 店舗からの日時変更リクエスト - #${orderNo}`;
  const alertBody = `
      <meta charset="UTF-8">
      <div style="background: #f4f5f7; padding: 24px 10px; font-family: 'Helvetica Neue', Arial, 'Hiragino Kaku Gothic ProN', Meiryo, sans-serif;">
        <div style="max-width: 500px; margin: 0 auto; padding: 32px 24px; background: #ffffff; border: 1px solid #eef0f2; border-radius: 14px; box-shadow: 0 4px 12px rgba(0,0,0,0.03);">
          
          <div style="margin-bottom: 20px; text-align: left;">
            <span style="font-size: 11px; font-weight: 800; letter-spacing: 1px; color: #ffffff; background: #7048e8; padding: 4px 10px; border-radius: 6px; display: inline-block;">ADMIN NOTICE</span>
            <h2 style="font-size: 20px; font-weight: 800; color: #1A2B49; margin: 10px 0 0 0;">BOOKMARK CREATORS</h2>
          </div>
          
          <div style="border-top: 2px solid #7048e8; padding-top: 20px; margin-bottom: 20px;">
            <p style="font-size: 14.5px; font-weight: 800; color: #7048e8; margin: 0 0 10px 0;">日程変更リクエスト受付</p>
            <p style="font-size: 13.5px; line-height: 1.6; color: #495057; margin: 0;">
              店舗（加盟店）より日程調整のリクエストが送信されました。<br>
              マスターログのステータスは自動的に <b>【調整必要】</b> に変更されました。クリエイターとの再スケジュール調整をお願いいたします。
            </p>
          </div>
          
          <div style="background: #f8f9fa; border: 1px solid #e9ecef; border-radius: 10px; padding: 16px 18px; margin-bottom: 16px;">
            <table style="width: 100%; border-collapse: collapse; font-size: 13.5px;">
              <tr>
                <td style="width: 105px; color: #8b95a1; font-weight: 700; padding: 6px 0; vertical-align: top;">■ 店舗名</td>
                <td style="color: #1A2B49; font-weight: 700; padding: 6px 0;">${storeName}</td>
              </tr>
              <tr>
                <td style="color: #8b95a1; font-weight: 700; padding: 6px 0; vertical-align: top;">■ クリエイター</td>
                <td style="color: #495057; font-weight: 700; padding: 6px 0;">${creatorName} <span style="font-size:11.5px; color:#8b95a1; font-weight:500;">(ID: ${memberCode})</span></td>
              </tr>
              <tr>
                <td style="color: #8b95a1; font-weight: 700; padding: 6px 0; vertical-align: top;">■ 注文番号</td>
                <td style="color: #495057; font-weight: 700; padding: 6px 0;">#${orderNo}</td>
              </tr>
              <tr>
                <td style="color: #8b95a1; font-weight: 700; padding: 6px 0; vertical-align: top;">■ 店舗メッセージ</td>
                <td style="color: #7048e8; font-weight: 800; padding: 6px 0; line-height: 1.5; word-break: break-all;">${msg}</td>
              </tr>
              <tr>
                <td style="color: #8b95a1; font-weight: 700; padding: 8px 0 0 0; vertical-align: top; border-top: 1px dashed #dee2e6;">■ ログ位置</td>
                <td style="color: #d63384; font-weight: 800; padding: 8px 0 0 0; border-top: 1px dashed #dee2e6;">Master_Log Sheet ${row}行目 (Row ${row})</td>
              </tr>
            </table>
          </div>
          
        </div>
      </div>`;
  _sendMail(p.to, alertSubject, alertBody);
}

/*********************** 예약 신청 알림 오케스트레이터 ***********************/

/** 예약 신청 직후: 점주 확인 메일 + 운영진 알림 발송, Master_Log Y열에 결과 기록 */
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

  const creatorProfileUrl = findCreatorProfileUrl(data[SHEET.USER], p.memberCode);

  // 1. 점주에게 일본어 예약 신청 메일 발송
  if (storeEmail && storeEmail.includes("@")) {
    try {
      sendStoreBookingRequestMail({
        to: storeEmail, storeNameJp: storeNameJp, memberName: p.memberName, dateStr: p.dateStr, timeStr: p.timeStr,
        peopleCount: p.safePeopleCount, benefit: storeBenefit, creatorProfileUrl: creatorProfileUrl, row: p.safeRow, orderNo: p.orderNo
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
    sendAdminBookingRequestMail({
      orderNo: p.orderNo, memberName: p.memberName, memberCode: p.memberCode, storeId: p.storeId,
      storeName: storeNameJp || storeNameKo, storeEmail: storeEmail, benefit: storeBenefit,
      dateStr: p.dateStr, timeStr: p.timeStr, peopleCount: p.safePeopleCount, creatorProfileUrl: creatorProfileUrl
    });
  } catch (adminErr) {
    console.error("운영진 알림 발송 실패: " + adminErr.toString());
  }
}
