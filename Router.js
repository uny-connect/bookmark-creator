/*********************************************************************
 * 웹 앱 진입점 (doGet / doPost) — 라우팅만 담당, 비즈니스 로직은 Service_*.js
 ********************************************************************/

/** 점주 안내용 결과 화면 공통 템플릿 (tone: 'fail' | 'success') */
function _renderNoticePage(tone, statusText, bodyHtml, pageTitle) {
  const statusClass = tone === 'success' ? 'success-status' : 'fail-status';
  const statusColor = tone === 'success' ? '#2D6A4F' : '#dc3545';
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
        .${statusClass} { color: ${statusColor}; font-weight: 700; font-size: 15px; margin-top: 0; margin-bottom: 16px; display: block; text-align: center; }
        .desc-text { color: #495057; font-size: 12.5px; line-height: 1.6; margin: 0; font-weight: 500; text-align: center; word-break: break-all; overflow-wrap: break-word; padding: 0 4px; }
        .notice-box { margin-top: 24px; padding: 14px; background-color: #f8f9fa; border-radius: 12px; border: 1px dashed #cfd4da; color: #495057; font-size: 13px; font-weight: 700; text-align: center; }
      </style>
    </head>
    <body>
      <div class="container">
        <div style="margin-bottom: 10px; margin-top: 10px;"><span class="badge-partner">PARTNER CENTER</span></div>
        <h2 class="brand-title">BOOKMARK CREATORS</h2>
        <div class="card">
          <span class="${statusClass}">${statusText}</span>
          <p class="desc-text">
            ${bodyHtml}
          </p>
          <div class="notice-box">このままページを閉じてください。</div>
        </div>
      </div>
    </body>
    </html>
  `).setTitle('BOOKMARK CREATORS | ' + pageTitle).addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

/** 🌐 웹 앱 화면 라우팅 (GET) */
function doGet(e) {
  const mode = e?.parameter?.mode;

  // 🎯 영업/담당자용 웹 테스트 발송 툴 (store 분기보다 앞에 배치하여 파라미터 충돌 방지)
  if (mode === 'test_tool') {
    const template = HtmlService.createTemplateFromFile('View_TestTool');
    template.initialStoreId = e?.parameter?.id || '';
    return template.evaluate()
      .setTitle('BOOKMARK CREATORS | 店舗メール登録・テスト')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
      .addMetaTag('viewport', 'width=device-width, initial-scale=1, maximum-scale=1, user-scalable=no');
  }

  // 1️⃣ 어드민 대시보드
  if (mode === 'admin') {
    return HtmlService.createTemplateFromFile('View_Admin').evaluate()
      .setTitle('BOOKMARK CREATORS | Admin')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
      .addMetaTag('viewport', 'width=device-width, initial-scale=1, maximum-scale=1');
  }

  // 2️⃣ 점주 파트너 레포트 대시보드
  if (mode === 'store' || e?.parameter?.id) {
    const template = HtmlService.createTemplateFromFile('View_Store');
    template.storeId = e?.parameter?.id || '';
    return template.evaluate()
      .setTitle('BOOKMARK CREATORS | パートナーレポート')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
      .addMetaTag('viewport', 'width=device-width, initial-scale=1, maximum-scale=1');
  }

  // 3️⃣ 메일 직통 링크 예약 승인
  if (mode === 'store_confirm') {
    const result = confirmBookingByLink(e.parameter);
    if (result.isExpired) {
      return _renderNoticePage('fail', '⚠️ 回答期限切れ',
        '本予約はご回答期限（<b>24時間</b>）を経過したため、<br>自動的にキャンセル処理が行われました。<br><br>クリエイターへの予約確定は行われませんでした。',
        '期限切れ');
    }
    if (result.isAlreadyRequested) {
      return _renderNoticePage('fail', '❌ 確定不可',
        "この件は店舗から<b>'日時変更リクエスト'</b>が送信されているか、<br>すでにキャンセルされたリクエストです。<br><br>新たな日程調整のご連絡をお待ちください。🙇‍♂️",
        '確定不可');
    }
    return _renderNoticePage('success', '✅ 予約確定完了',
      '来店予約が確定しました。<br>クリエイターの訪問日時にあわせて、<br>ご準備をお願いいたします。🙏',
      '確定完了');
  }

  // 4️⃣ 점주 일정변경 요청 입력 (24시간 만료 및 취소 건 차단)
  if (mode === 'feedback') {
    if (isFeedbackBlocked(e.parameter)) {
      return _renderNoticePage('fail', '⚠️ 回答期限切れ',
        '本予約はご回答期限（<b>24時間</b>）を経過したため、<br>自動的にキャンセル処理が行われました。<br><br>日時変更リクエストの送信は行えません。',
        '期限切れ');
    }
    const template = HtmlService.createTemplateFromFile('View_Feedback');
    template.row = e?.parameter?.row || '';
    template.orderNo = e?.parameter?.o || '';
    return template.evaluate().setTitle('BOOKMARK CREATORS | リクエスト').setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
      .addMetaTag('viewport', 'width=device-width, initial-scale=1, maximum-scale=1, user-scalable=no');
  }

  // 5️⃣ 기본 화면 (크리에이터용 PIN 화면)
  const template = HtmlService.createTemplateFromFile('View_Creator');
  template.orderNo = (e?.parameter?.o) || '';
  template.phoneLast4 = (e?.parameter?.p) || '';
  return template.evaluate().setTitle('BOOKMARK CREATORS | Creator').setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, maximum-scale=1, user-scalable=no');
}

/** 🔌 아임웹 웹훅 진입점 (POST) */
function doPost(e) {
  return handleImwebWebhook(e);
}
