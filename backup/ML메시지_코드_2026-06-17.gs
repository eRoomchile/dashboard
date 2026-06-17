// ==========================================
// ML 고객 메시지 반자동화 시스템 v2.2
// ==========================================

const CONFIG_SHEET = '⚙️설정';
const PRE_SHEET = '📨구매전문의';
const POST_SHEET = '📦구매후메시지';
const CLAIM_SHEET = '⚠️분쟁';

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('🤖 ML 메시지')
    .addItem('1️⃣ 초기 설정 실행', 'setupSheets')
    .addSeparator()
    .addItem('2️⃣ 토큰 발급 (URL 열기)', 'getAuthUrl')
    .addItem('3️⃣ 토큰 저장 (code 입력)', 'saveToken')
    .addSeparator()
    .addItem('4️⃣ 구매전 문의 가져오기', 'fetchPreSalesMessages')
    .addItem('5️⃣ 구매후 메시지 가져오기', 'fetchPostSalesMessages')
    .addItem('6️⃣ 분쟁 가져오기', 'fetchClaims')
    .addSeparator()
    .addItem('7️⃣ 스페인어 미리보기 (구매전)', 'previewPreReplies')
    .addItem('8️⃣ 스페인어 미리보기 (구매후)', 'previewPostReplies')
    .addSeparator()
    .addItem('9️⃣ 답변 전송 (구매전)', 'sendPreReplies')
    .addItem('🔟 답변 전송 (구매후)', 'sendPostReplies')
    .addSeparator()
    .addItem('⏱️ 자동실행 시작 (30분마다)', 'setupTriggers')
    .addItem('⏹️ 자동실행 중지', 'removeTriggers')
    .addSeparator()
    .addItem('📦 반품/취소 가져오기', 'fetchReturns')
    .addSeparator()
    .addItem('🛒 ML판매 초기설정', 'setupMlSalesSheets')
    .addItem('📊 ML판매내역 가져오기', 'fetchMlSales')
    .addItem('✅ 판매완료 처리', 'completeMlSales')
    .addItem('🔄 SKU불일치 재처리', 'reprocessSkuMismatch')
    .addItem('📋 판매기록 날짜순 정렬', 'sortRecordSheet')
    .addToUi();
}

// ==========================================
// showAlert
// ==========================================
function showAlert(msg) {
  try {
    SpreadsheetApp.getUi().alert(msg);
  } catch (e) {
    console.log(msg);
  }
}

// ==========================================
// 초기 설정
// ==========================================
function setupSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // ⚙️설정 시트
  let configSheet = ss.getSheetByName(CONFIG_SHEET);
  if (!configSheet) configSheet = ss.insertSheet(CONFIG_SHEET);
  configSheet.clearContents();
  configSheet.getRange('A1:B1').setValues([['항목', '값']]);
  configSheet.getRange('A2:B6').setValues([
    ['CLIENT_ID', ''],
    ['SECRET_KEY', ''],
    ['GOOGLE_API_KEY', ''],
    ['ACCESS_TOKEN', ''],
    ['REFRESH_TOKEN', ''],
  ]);
  configSheet.getRange('A1:B1').setBackground('#34495e').setFontColor('white').setFontWeight('bold');
  configSheet.setColumnWidth(1, 150);
  configSheet.setColumnWidth(2, 400);

  // 📨구매전문의 시트
  let preSheet = ss.getSheetByName(PRE_SHEET);
  if (!preSheet) preSheet = ss.insertSheet(PRE_SHEET);
  preSheet.clearContents();
  const preHeaders = ['메시지ID', '상품명', '상품링크', '구매자', '날짜', '유형', '스페인어 원문', '한글 번역', '한글 답변', '스페인어 번역(미리보기)', '전송상태'];
  preSheet.getRange(1, 1, 1, preHeaders.length).setValues([preHeaders]);
  preSheet.getRange(1, 1, 1, preHeaders.length).setBackground('#2980b9').setFontColor('white').setFontWeight('bold');
  preSheet.setColumnWidth(2, 200);
  preSheet.setColumnWidth(3, 300);
  preSheet.setColumnWidth(7, 300);
  preSheet.setColumnWidth(8, 300);
  preSheet.setColumnWidth(9, 300);
  preSheet.setColumnWidth(10, 300);

  // 📦구매후메시지 시트
  let postSheet = ss.getSheetByName(POST_SHEET);
  if (!postSheet) postSheet = ss.insertSheet(POST_SHEET);
  postSheet.clearContents();
  const postHeaders = ['메시지ID', '주문번호', '구매자', '날짜', '주문링크', '상품명', '주문금액', '배송상태', '유형', '스페인어 원문', '한글 번역', '한글 답변', '스페인어 번역(미리보기)', '전송상태'];
  postSheet.getRange(1, 1, 1, postHeaders.length).setValues([postHeaders]);
  postSheet.getRange(1, 1, 1, postHeaders.length).setBackground('#27ae60').setFontColor('white').setFontWeight('bold');
  postSheet.setColumnWidth(5, 200);
  postSheet.setColumnWidth(9, 300);
  postSheet.setColumnWidth(10, 300);
  postSheet.setColumnWidth(11, 300);
  postSheet.setColumnWidth(12, 300);

  // ⚠️분쟁 시트
  let claimSheet = ss.getSheetByName(CLAIM_SHEET);
  if (!claimSheet) claimSheet = ss.insertSheet(CLAIM_SHEET);
  claimSheet.clearContents();
  const claimHeaders = ['분쟁ID', '주문번호', '구매자', '날짜', '분쟁유형', '분쟁유형(한글)', '상태', '스페인어 내용', '한글 번역', '비고'];
  claimSheet.getRange(1, 1, 1, claimHeaders.length).setValues([claimHeaders]);
  claimSheet.getRange(1, 1, 1, claimHeaders.length).setBackground('#c0392b').setFontColor('white').setFontWeight('bold');
  claimSheet.setColumnWidth(8, 300);
  claimSheet.setColumnWidth(9, 300);

  showAlert('✅ 초기 설정 완료!\n\n⚙️설정 시트에 아래 3가지를 입력하세요:\n- CLIENT_ID\n- SECRET_KEY\n- GOOGLE_API_KEY');
}

// ==========================================
// 설정값 읽기/쓰기
// ==========================================
function getConfig(key) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG_SHEET);
  const data = sheet.getRange('A2:B6').getValues();
  for (const row of data) {
    if (row[0] === key) return row[1];
  }
  return '';
}

function setConfig(key, value) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG_SHEET);
  const data = sheet.getRange('A2:B6').getValues();
  for (let i = 0; i < data.length; i++) {
    if (data[i][0] === key) {
      sheet.getRange(i + 2, 2).setValue(value);
      return;
    }
  }
}

// ==========================================
// 토큰 관련
// ==========================================
function getAuthUrl() {
  const clientId = getConfig('CLIENT_ID');
  if (!clientId) { showAlert('❌ 설정 시트에 CLIENT_ID를 먼저 입력하세요!'); return; }
  const redirectUri = 'https://www.google.com';
  const url = `https://auth.mercadolibre.cl/authorization?response_type=code&client_id=${clientId}&redirect_uri=${encodeURIComponent(redirectUri)}`;
  showAlert('📋 아래 URL을 브라우저에서 열고\n승인 후 나오는 code= 값을 복사하세요:\n\n' + url);
}

function saveToken() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.prompt('🔑 토큰 저장', 'URL에서 복사한 code 값을 입력하세요:', ui.ButtonSet.OK_CANCEL);
  if (response.getSelectedButton() !== ui.Button.OK) return;

  const code = response.getResponseText().trim();
  const clientId = getConfig('CLIENT_ID');
  const secretKey = getConfig('SECRET_KEY');

  const res = UrlFetchApp.fetch('https://api.mercadolibre.com/oauth/token', {
    method: 'post',
    payload: {
      grant_type: 'authorization_code',
      client_id: clientId,
      client_secret: secretKey,
      code: code,
      redirect_uri: 'https://www.google.com'
    },
    muteHttpExceptions: true
  });

  const data = JSON.parse(res.getContentText());
  if (data.access_token) {
    setConfig('ACCESS_TOKEN', data.access_token);
    setConfig('REFRESH_TOKEN', data.refresh_token);
    ui.alert('✅ 토큰 저장 완료!\n이제 메시지 가져오기를 클릭하세요.');
  } else {
    ui.alert('❌ 토큰 발급 실패:\n' + res.getContentText());
  }
}

function refreshAccessToken() {
  const clientId = getConfig('CLIENT_ID');
  const secretKey = getConfig('SECRET_KEY');
  const refreshToken = getConfig('REFRESH_TOKEN');
  if (!refreshToken) return false;

  const res = UrlFetchApp.fetch('https://api.mercadolibre.com/oauth/token', {
    method: 'post',
    payload: {
      grant_type: 'refresh_token',
      client_id: clientId,
      client_secret: secretKey,
      refresh_token: refreshToken
    },
    muteHttpExceptions: true
  });

  const data = JSON.parse(res.getContentText());
  if (data.access_token) {
    setConfig('ACCESS_TOKEN', data.access_token);
    setConfig('REFRESH_TOKEN', data.refresh_token);
    return true;
  }
  return false;
}

function getHeaders() {
  return { Authorization: 'Bearer ' + getConfig('ACCESS_TOKEN') };
}

function getMyUserId() {
  const res = UrlFetchApp.fetch('https://api.mercadolibre.com/users/me', {
    headers: getHeaders(), muteHttpExceptions: true
  });
  return JSON.parse(res.getContentText()).id;
}

// ==========================================
// 번역 함수
// ==========================================
function translateText(text, targetLang) {
  if (!text) return '';
  const apiKey = getConfig('GOOGLE_API_KEY');
  if (!apiKey) return text;

  const url = `https://translation.googleapis.com/language/translate/v2?key=${apiKey}`;
  const res = UrlFetchApp.fetch(url, {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify({ q: text, target: targetLang }),
    muteHttpExceptions: true
  });

  try {
    return JSON.parse(res.getContentText()).data.translations[0].translatedText;
  } catch (e) {
    return text;
  }
}

// ==========================================
// 유형 감지 / 코드 변환
// ==========================================
function detectMessageType(originalText, translatedText) {
  const text = (originalText + ' ' + translatedText).toLowerCase();
  if (text.match(/boleta|factura|recibo|영수증|invoice|receipt/)) return '🧾영수증요청';
  if (text.match(/defecto|roto|malo|equivocado|incorrecto|불량|파손|오배송|다른제품/)) return '⚠️불량오배송';
  if (text.match(/reembolso|devolucion|devolución|cancelar|환불|취소|반품/)) return '💸환불취소';
  if (text.match(/cantidad|falta|faltó|수량|부족|빠진/)) return '📦수량문제';
  if (text.match(/envío|envio|despacho|llegó|llego|no llegó|배송|도착|미도착/)) return '🚚배송문의';
  return '💬일반문의';
}

function translateClaimReason(reason) {
  const map = {
    'PDD9939': '상품이 설명과 다름',
    'PNR9945': '상품 미도착',
    'PDD9941': '상품 파손',
    'PDD9940': '수량 부족',
    'PNR9944': '배송 지연',
  };
  return map[reason] || reason;
}

function translateClaimStatus(status) {
  const map = {
    'opened': '🚨진행중',
    'closed': '✅종료',
    'resolved': '✅해결완료',
  };
  return map[status] || status;
}

function translateShipStatus(status) {
  const map = {
    'pending': '결제대기',
    'handling': '준비중',
    'ready_to_ship': '발송준비완료',
    'shipped': '배송중',
    'delivered': '배송완료',
    'not_delivered': '미배송',
    'cancelled': '취소됨'
  };
  return map[status] || status;
}

// ==========================================
// 주문/상품 정보
// ==========================================
function getOrderInfo(packId) {
  try {
    const res = UrlFetchApp.fetch(
      `https://api.mercadolibre.com/orders/search?pack=${packId}`,
      { headers: getHeaders(), muteHttpExceptions: true }
    );
    const data = JSON.parse(res.getContentText());
    const order = data.results?.[0];
    if (!order) return { itemTitle: '', amount: '', shipStatus: '' };
    return {
      itemTitle: order.order_items?.[0]?.item?.title || '',
      amount: order.total_amount || '',
      shipStatus: translateShipStatus(order.shipping?.status || '')
    };
  } catch (e) {
    return { itemTitle: '', amount: '', shipStatus: '' };
  }
}

function getItemTitle(itemId) {
  try {
    const res = UrlFetchApp.fetch(
      `https://api.mercadolibre.com/items/${itemId}`,
      { headers: getHeaders(), muteHttpExceptions: true }
    );
    return JSON.parse(res.getContentText()).title || '';
  } catch (e) {
    return '';
  }
}

// ==========================================
// 4️⃣ 구매전 문의 가져오기
// ==========================================
function fetchPreSalesMessages() {
  const token = getConfig('ACCESS_TOKEN');
  if (!token) { showAlert('❌ 먼저 토큰 발급을 완료하세요!'); return; }

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(PRE_SHEET);
  const lastRow = sheet.getLastRow();
  const existingIds = lastRow > 1
    ? sheet.getRange(2, 1, lastRow - 1, 1).getValues().flat().map(String) : [];

  const userId = getMyUserId();
  const res = UrlFetchApp.fetch(
   `https://api.mercadolibre.com/questions/search?seller_id=${userId}&sort_fields=date_created&sort_types=DESC&limit=15`, 
    { headers: getHeaders(), muteHttpExceptions: true }
  );

  const data = JSON.parse(res.getContentText());
  let newCount = 0;

  // 최근 7일 이내 질문만
  const cutoffDate = new Date();
  cutoffDate.setDate(cutoffDate.getDate() - 7);

  if (data.questions) {
    for (const q of data.questions) {
      const qId = String(q.id || '');
      if (existingIds.includes(qId)) continue;

      // 7일 이전 질문 건너뜀
      const qDate = new Date(q.date_created);
      if (qDate < cutoffDate) continue;

      const text = q.text || '';
      const translated = translateText(text, 'ko');
      const itemTitle = q.item_id ? getItemTitle(q.item_id) : '';
      const buyer = q.from?.id ? String(q.from.id) : '';
      const date = q.date_created || '';
      const msgType = detectMessageType(text, translated);

      const itemLink = q.item_id ? `https://articulo.mercadolibre.cl/MLC-${q.item_id.replace('MLC','')}-_JM` : '';
sheet.appendRow([qId, itemTitle, itemLink, buyer, date, msgType, text, translated, '', '', '미전송']);
      newCount++;
    }
  }
  showAlert(`✅ 구매전 문의 ${newCount}개 새로 가져왔습니다.`);
}
// ==========================================
// 5️⃣ 구매후 메시지 가져오기
// ==========================================
function fetchPostSalesMessages() {
  const token = getConfig('ACCESS_TOKEN');
  if (!token) { showAlert('❌ 먼저 토큰 발급을 완료하세요!'); return; }

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(POST_SHEET);
  const lastRow = sheet.getLastRow();
  const existingMsgIds = lastRow > 1
    ? sheet.getRange(2, 1, lastRow - 1, 1).getValues().flat().map(String) : [];

  const userId = getMyUserId();

  // 최근 주문 50개 가져오기
  const ordersRes = UrlFetchApp.fetch(
    `https://api.mercadolibre.com/orders/search?seller=${userId}&sort=date_desc&limit=50`,
    { headers: getHeaders(), muteHttpExceptions: true }
  );
  const ordersData = JSON.parse(ordersRes.getContentText());
  const orders = ordersData.results || [];

  let newCount = 0;
  const processedPacks = [];

  for (const order of orders) {
    // pack_id가 있으면 pack_id 사용, 없으면 order_id 사용
    const packId = String(order.pack_id || order.id);

    // 이미 처리한 pack_id 중복 방지
    if (processedPacks.includes(packId)) continue;
    processedPacks.push(packId);

    // 해당 주문 메시지 가져오기
    const msgRes = UrlFetchApp.fetch(
      `https://api.mercadolibre.com/messages/packs/${packId}/sellers/${userId}?tag=post_sale`,
      { headers: getHeaders(), muteHttpExceptions: true }
    );
    const msgData = JSON.parse(msgRes.getContentText());

    // active 상태이고 메시지 있는 것만 처리
    if (msgData.conversation_status?.status !== 'active') continue;
    if (!msgData.messages || msgData.messages.length === 0) continue;

    // 주문 정보
    const itemTitle = order.order_items?.[0]?.item?.title || '';
    const orderAmount = order.total_amount || '';
    const shipStatus = translateShipStatus(order.shipping?.status || '');
    const buyer = order.buyer?.nickname || '';

    // 구매자가 보낸 메시지만 처리
    for (const msg of msgData.messages) {
      const msgId = String(msg.id || '');
      if (existingMsgIds.includes(msgId)) continue;

      // 구매자 메시지만 (판매자 메시지 제외)
      if (String(msg.from?.user_id) === String(userId)) continue;

      // 빈 텍스트 제외
      const text = msg.text || '';
      if (!text.trim()) continue;

      const date = msg.message_date?.created || '';
      const translated = translateText(text, 'ko');
      const msgType = detectMessageType(text, translated);

      const orderLink = `https://www.mercadolibre.cl/ventas/${packId}/detalle`;
sheet.appendRow([msgId, packId, buyer, date, orderLink, itemTitle, orderAmount, shipStatus, msgType, text, translated, '', '', '미전송']);
      newCount++;
    }
  }

  showAlert(`✅ 구매후 메시지 ${newCount}개 새로 가져왔습니다.`);
}
// ==========================================
// 6️⃣ 분쟁 가져오기
// ==========================================
function fetchClaims() {
  const token = getConfig('ACCESS_TOKEN');
  if (!token) { showAlert('❌ 먼저 토큰 발급을 완료하세요!'); return; }

  const userId = getMyUserId();
  const res = UrlFetchApp.fetch(
    `https://api.mercadolibre.com/post-purchase/v1/claims/search?seller_id=${userId}&status=opened`,
    { headers: getHeaders(), muteHttpExceptions: true }
  );

  const claimSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CLAIM_SHEET);
  const lastRow = claimSheet.getLastRow();
  const existingIds = lastRow > 1
    ? claimSheet.getRange(2, 1, lastRow - 1, 1).getValues().flat().map(String) : [];

  const data = JSON.parse(res.getContentText());
  let newCount = 0;
  const list = data.data || data.results || [];

  for (const claim of list) {
    const claimId = String(claim.id || '');
    if (existingIds.includes(claimId)) continue;
    const reason = claim.reason_id || '';
    const reasonKo = translateClaimReason(reason);
    const status = translateClaimStatus(claim.status || '');
    const date = claim.date_created || '';
    const orderId = claim.resource_id || '';
    const buyer = claim.players?.find(p => p.role === 'complainant')?.user_id || '';
    const detail = claim.resolution?.reason || reason;
    const detailKo = translateText(detail, 'ko');
    claimSheet.appendRow([claimId, orderId, buyer, date, reason, reasonKo, status, detail, detailKo, '']);
    newCount++;
  }
  showAlert(`✅ 분쟁 ${newCount}개 새로 가져왔습니다.`);
}

// ==========================================
// 7️⃣ 스페인어 미리보기 (구매전)
// ==========================================
function previewPreReplies() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(PRE_SHEET);
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) { showAlert('데이터가 없습니다.'); return; }

  const data = sheet.getRange(2, 1, lastRow - 1, 11).getValues();
  let count = 0;

  Logger.log('총 행수: ' + data.length);
  for (let i = 0; i < data.length; i++) {
    Logger.log(`행${i+2} / H열(한글답변): "${data[i][7]}" / I열(한글답변2): "${data[i][8]}" / J열(미리보기): "${data[i][9]}"`);
  }

  for (let i = 0; i < data.length; i++) {
    const koreanReply = data[i][8]; // I열 한글답변
    const preview = data[i][9];    // J열 미리보기
    if (!koreanReply) continue;    // 한글답변 없으면 건너뜀
    if (preview) continue;         // 미리보기 이미 있으면 건너뜀
    
    const spanishPreview = translateText(koreanReply, 'es');
    sheet.getRange(i + 2, 10).setValue(spanishPreview);
    count++;
  }
  showAlert(`✅ 스페인어 미리보기 ${count}개 생성 완료!\n확인 후 답변 전송을 클릭하세요.`);
}
// ==========================================
// 8️⃣ 스페인어 미리보기 (구매후)
// ==========================================
function previewPostReplies() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(POST_SHEET);
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) { showAlert('데이터가 없습니다.'); return; }

  const data = sheet.getRange(2, 1, lastRow - 1, 13).getValues();
  let count = 0;

  for (let i = 0; i < data.length; i++) {
    const koreanReply = data[i][11]; // 열 한글답변
    const preview = data[i][12];     //열 미리보기
    if (!koreanReply || preview) continue;
    const spanishPreview = translateText(koreanReply, 'es');
    sheet.getRange(i + 2, 13).setValue(spanishPreview);
    count++;
  }
  showAlert(`✅ 스페인어 미리보기 ${count}개 생성 완료!\n확인 후 답변 전송을 클릭하세요.`);
}

// ==========================================
// 9️⃣ 답변 전송 (구매전)
// ==========================================
function sendPreReplies() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(PRE_SHEET);
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) { showAlert('전송할 메시지가 없습니다.'); return; }

  const data = sheet.getRange(2, 1, lastRow - 1, 11).getValues();
  let sentCount = 0;

  for (let i = 0; i < data.length; i++) {
    const qId = data[i][0];
    const koreanReply = data[i][8];
    const status = data[i][10];
    if (!koreanReply || status === '✅전송완료') continue;

    const spanishReply = data[i][9] || translateText(koreanReply, 'es');

    const res = UrlFetchApp.fetch(
      `https://api.mercadolibre.com/answers`,
      {
        method: 'post',
        headers: { ...getHeaders(), 'Content-Type': 'application/json' },
        payload: JSON.stringify({
          question_id: Number(qId),
          text: spanishReply
        }),
        muteHttpExceptions: true
      }
    );

    const result = JSON.parse(res.getContentText());
    if (result.id) {
      sheet.getRange(i + 2, 11).setValue('✅전송완료');
      sentCount++;
    } else if (result.error === 'not_unanswered_question') {
      sheet.getRange(i + 2, 11).setValue('✅이미답변됨');
    } else {
      sheet.getRange(i + 2, 11).setValue('❌실패: ' + result.error);
    }
  }
  showAlert(`✅ 구매전 답변 ${sentCount}개 전송 완료!`);
}

// ==========================================
// 🔟 답변 전송 (구매후)
// ==========================================
function sendPostReplies() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(POST_SHEET);
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) { showAlert('전송할 메시지가 없습니다.'); return; }

  const data = sheet.getRange(2, 1, lastRow - 1, 14).getValues();
  let sentCount = 0;

  for (let i = 0; i < data.length; i++) {
    const msgId = data[i][0];
    const packId = data[i][1];
    const koreanReply = data[i][11]; // K열 한글답변
    const status = data[i][13];      // M열 전송상태
    if (!koreanReply || status === '✅전송완료') continue;

    // 미리보기가 있으면 그걸 사용, 없으면 새로 번역
    const spanishReply = data[i][11] || translateText(koreanReply, 'es');

    const res = UrlFetchApp.fetch(
      `https://api.mercadolibre.com/messages/packs/${packId}/sellers/me`,
      {
        method: 'post',
        headers: { ...getHeaders(), 'Content-Type': 'application/json' },
        payload: JSON.stringify({ text: spanishReply }),
        muteHttpExceptions: true
      }
    );
    const result = JSON.parse(res.getContentText());
    sheet.getRange(i + 2, 14).setValue(result.id ? '✅전송완료' : '❌실패');
    if (result.id) sentCount++;
  }
  showAlert(`✅ 구매후 답변 ${sentCount}개 전송 완료!`);
}

// ==========================================
// ⏱️ 자동 트리거
// ==========================================
function setupTriggers() {
  const triggers = ScriptApp.getProjectTriggers();
  for (const trigger of triggers) ScriptApp.deleteTrigger(trigger);
  ScriptApp.newTrigger('autoFetchAll').timeBased().everyMinutes(5).create();
  showAlert('✅ 트리거 설정 완료!\n30분마다 자동으로 메시지를 가져옵니다.');
}

function removeTriggers() {
  const triggers = ScriptApp.getProjectTriggers();
  for (const trigger of triggers) ScriptApp.deleteTrigger(trigger);
  showAlert('✅ 자동 트리거가 중지되었습니다.');
}

function autoFetchAll() {
  try {
    refreshAccessToken();
    fetchPreSalesMessages();
    fetchPostSalesMessages();
    fetchClaims();
    fetchReturns();
    fetchMlSales();
    syncTokenToNaos();
    logAutoRun('✅ 성공');
  } catch (e) {
    logAutoRun('❌ 오류: ' + e.message);
  }
}

function logAutoRun(status) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let logSheet = ss.getSheetByName('📋실행로그');
  if (!logSheet) {
    logSheet = ss.insertSheet('📋실행로그');
    logSheet.getRange(1, 1, 1, 3).setValues([['실행시간', '상태', '비고']]);
    logSheet.getRange(1, 1, 1, 3).setBackground('#7f8c8d').setFontColor('white').setFontWeight('bold');
  }
  logSheet.appendRow([new Date(), status, '자동실행']);
  const lastRow = logSheet.getLastRow();
  if (lastRow > 101) logSheet.deleteRows(2, lastRow - 101);
}



function getSharedToken() {
  // ML_고객메시지 파일의 설정 시트에서 토큰 읽기
  const ML_FILE_ID = SpreadsheetApp.getActiveSpreadsheet().getId();
  const configSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG_SHEET);
  const data = configSheet.getRange('A2:B6').getValues();
  for (const row of data) {
    if (row[0] === 'ACCESS_TOKEN') return row[1];
  }
  return '';
}

// ==========================================
// 🔗 NAOS재고 파일에 토큰 공유
// ==========================================
function syncTokenToNaos() {
  try {
    const token = getConfig('ACCESS_TOKEN');
    const refreshToken = getConfig('REFRESH_TOKEN');
    if (!token) return;

    const naosFile = SpreadsheetApp.openById('1hw8M2WsmmTY-Obz12N3WUBi7Fv1tlrs9tGBUGvwHWrk');
    const naosConfig = naosFile.getSheetByName('⚙️ML설정');
    
    if (!naosConfig) {
      // 설정 시트 없으면 생성
      const newSheet = naosFile.insertSheet('⚙️ML설정');
      newSheet.getRange('A1:B1').setValues([['항목', '값']]);
      newSheet.getRange('A2:B3').setValues([
        ['ACCESS_TOKEN', token],
        ['REFRESH_TOKEN', refreshToken]
      ]);
      newSheet.getRange('A1:B1').setBackground('#34495e').setFontColor('white').setFontWeight('bold');
      newSheet.setColumnWidth(1, 150);
      newSheet.setColumnWidth(2, 400);
    } else {
      const data = naosConfig.getRange('A2:B3').getValues();
      for (let i = 0; i < data.length; i++) {
        if (data[i][0] === 'ACCESS_TOKEN') naosConfig.getRange(i + 2, 2).setValue(token);
        if (data[i][0] === 'REFRESH_TOKEN') naosConfig.getRange(i + 2, 2).setValue(refreshToken);
      }
    }
  } catch(e) {
    console.log('토큰 동기화 오류: ' + e.message);
  }
}

// ==========================================
// 📦 반품/취소 가져오기
// ==========================================
function fetchReturns() {
  const token = getConfig('ACCESS_TOKEN');
  if (!token) { showAlert('❌ 먼저 토큰 발급을 완료하세요!'); return; }

  const userId = getMyUserId();

  // 취소 주문 가져오기
  const res = UrlFetchApp.fetch(
    `https://api.mercadolibre.com/orders/search?seller=${userId}&order.status=cancelled&sort=date_desc&limit=50`,
    { headers: getHeaders(), muteHttpExceptions: true }
  );
  const data = JSON.parse(res.getContentText());

  // 반품 진행중 가져오기
  const res2 = UrlFetchApp.fetch(
    `https://api.mercadolibre.com/orders/search?seller=${userId}&order.status=invalid&sort=date_desc&limit=50`,
    { headers: getHeaders(), muteHttpExceptions: true }
  );
  const data2 = JSON.parse(res2.getContentText());

  const allOrders = [
    ...(data.results || []),
    ...(data2.results || [])
  ];

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let returnSheet = ss.getSheetByName('📦반품취소');

  // 시트 없으면 새로 생성
  if (!returnSheet) {
    returnSheet = ss.insertSheet('📦반품취소');
  }

  // 헤더 항상 업데이트
  const headers = ['주문번호', '구매자', '상품명', 'SKU', '수량', '금액', '상태', '날짜', '반품사유(스페인어)', '반품사유(한글)', '주문링크', '비고'];
  returnSheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  returnSheet.getRange(1, 1, 1, headers.length)
    .setBackground('#922b21').setFontColor('white').setFontWeight('bold');
  returnSheet.setColumnWidth(3, 250);
  returnSheet.setColumnWidth(9, 250);
  returnSheet.setColumnWidth(10, 250);
  returnSheet.setColumnWidth(11, 80);

  const lastRow = returnSheet.getLastRow();
  const existingIds = lastRow > 1
    ? returnSheet.getRange(2, 1, lastRow - 1, 1).getValues().flat().map(String) : [];

  let newCount = 0;

  for (const order of allOrders) {
    const orderId = String(order.id || '');
    if (existingIds.includes(orderId)) continue;

    const buyer = order.buyer?.nickname || '';
    const itemTitle = order.order_items?.[0]?.item?.title || '';
    const sku = order.order_items?.[0]?.item?.seller_sku || '';
    const qty = order.order_items?.[0]?.quantity || '';
    const amount = order.total_amount || '';
    const status = translateReturnStatus(order.status || '');
    const date = order.date_created || '';

    // 반품 사유 (스페인어)
    const cancelDetail = order.cancel_detail || {};
    const reasonCode = cancelDetail.code || '';
    const reasonDesc = cancelDetail.description || '';
    const requestedBy = cancelDetail.requested_by || '';
    const reasonEs = reasonDesc || reasonCode || '';

    // 사유 코드 한글 변환
    const codeMap = {
      'shipment_not_delivered': '배송 미완료',
      'feedback_out_of_stock': '재고 없음 (판매자)',
      'pack_unknown': '알 수 없는 사유',
      'buyer_cancel': '구매자 취소',
      'seller_cancel': '판매자 취소',
      'out_of_stock': '재고 없음',
      'wrong_product': '잘못된 상품',
      'damaged_product': '상품 파손',
    };

    const requestedByKo = requestedBy === 'buyer' ? '구매자요청' :
                          requestedBy === 'seller' ? '판매자요청' :
                          requestedBy === 'meli' ? 'ML자동처리' : requestedBy;

    const reasonKo = (codeMap[reasonCode] || translateText(reasonEs, 'ko')) +
                     (requestedByKo ? ` (${requestedByKo})` : '');

    // 주문 링크
    const orderLink = `https://www.mercadolibre.cl/ventas/${orderId}/detalle`;

    returnSheet.appendRow([
      orderId, buyer, itemTitle, sku, qty, amount,
      status, date, reasonEs, reasonKo, '', ''
    ]);

    // 링크 하이퍼링크 설정
    const newRow = returnSheet.getLastRow();
    returnSheet.getRange(newRow, 11).setFormula(`=HYPERLINK("${orderLink}","링크")`);

    newCount++;
  }

  showAlert(`✅ 반품/취소 ${newCount}개 새로 가져왔습니다.`);
}

function translateReturnStatus(status) {
  const map = {
    'cancelled': '🔴취소됨',
    'invalid': '⚠️반품진행중',
    'paid': '✅결제완료',
    'partially_refunded': '💸부분환불',
    'refunded': '💸전액환불',
  };
  return map[status] || status;
}

// ==========================================
// 📊 ML판매내역 + 📋ML판매기록 초기 설정
// ==========================================
function setupMlSalesSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // 📊ML판매내역 시트
  let salesSheet = ss.getSheetByName('📊ML판매내역');
  if (!salesSheet) salesSheet = ss.insertSheet('📊ML판매내역');
  salesSheet.clearContents();
  const salesHeaders = [
    '주문번호', '날짜', '배송유형', '구매자',
    'SKU', '상품명', '판매수량', '묶음수량', '차감수량',
    '단가', '총금액', '처리상태', '주문링크'
  ];
  salesSheet.getRange(1, 1, 1, salesHeaders.length).setValues([salesHeaders]);
  salesSheet.getRange(1, 1, 1, salesHeaders.length)
    .setBackground('#1a5276').setFontColor('white').setFontWeight('bold');
  salesSheet.setColumnWidth(1, 180);
  salesSheet.setColumnWidth(4, 150);
  salesSheet.setColumnWidth(5, 150);
  salesSheet.setColumnWidth(6, 250);
  salesSheet.setColumnWidth(12, 100);
  salesSheet.setColumnWidth(13, 60);

  // 📋ML판매기록 시트
  let recordSheet = ss.getSheetByName('📋ML판매기록');
  if (!recordSheet) recordSheet = ss.insertSheet('📋ML판매기록');
  recordSheet.clearContents();
  const recordHeaders = [
    '날짜', '주문번호', '배송유형', '구매자',
    'SKU', '상품명', '판매수량', '묶음수량', '차감수량',
    '단가', '총금액', '처리일시', '재고차감여부', 'URL'
  ];
  recordSheet.getRange(1, 1, 1, recordHeaders.length).setValues([recordHeaders]);
  recordSheet.getRange(1, 1, 1, recordHeaders.length)
    .setBackground('#1e8449').setFontColor('white').setFontWeight('bold');
  recordSheet.setColumnWidth(2, 180);
  recordSheet.setColumnWidth(5, 150);
  recordSheet.setColumnWidth(6, 250);
  recordSheet.setColumnWidth(14, 60);

  showAlert('✅ ML판매내역, ML판매기록 시트 생성 완료!');
}

// ==========================================
// 📊 ML판매내역 가져오기
// ==========================================
function fetchMlSales() {
  const token = getConfig('ACCESS_TOKEN');
  if (!token) { showAlert('❌ 먼저 토큰 발급을 완료하세요!'); return; }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const salesSheet = ss.getSheetByName('📊ML판매내역');
  if (!salesSheet) { showAlert('❌ 먼저 초기 설정을 실행하세요!'); return; }

  // 1. 판매내역 기존 주문번호
  const lastRow = salesSheet.getLastRow();
  const existingIds = [];

  if (lastRow > 1) {
    salesSheet.getRange(2, 1, lastRow - 1, 1).getValues().flat()
      .forEach(id => {
        const clean = String(id).replace(/'/g, '').trim();
        if (clean && !existingIds.includes(clean)) existingIds.push(clean);
      });
  }

  // 2. ✅완료주문번호 시트에서 중복체크
  const doneSheet = ss.getSheetByName('✅완료주문번호');
  if (doneSheet && doneSheet.getLastRow() > 1) {
    doneSheet.getRange(2, 1, doneSheet.getLastRow() - 1, 1)
      .getValues().flat()
      .forEach(id => {
        const clean = String(id).replace(/'/g, '').trim();
        if (clean && !existingIds.includes(clean)) existingIds.push(clean);
      });
  }

  // 3. 판매기록 주문번호도 포함
  const recordSheet = ss.getSheetByName('📋ML판매기록');
  if (recordSheet && recordSheet.getLastRow() > 1) {
    recordSheet.getRange(2, 2, recordSheet.getLastRow() - 1, 1)
      .getValues().flat()
      .forEach(id => {
        const clean = String(id).replace(/'/g, '').trim();
        if (clean && !existingIds.includes(clean)) existingIds.push(clean);
      });
  }

  Logger.log('중복체크 ID수: ' + existingIds.length);

  // 묶음수량 맵
  const bundleMap = getBundleMap();

  // 최근 3일 주문
  const today = new Date();
  const from = new Date(today);
  from.setDate(from.getDate() - 3);
  const fromStr = from.toISOString().split('T')[0] + 'T00:00:00.000-04:00';

  // 주문 목록을 텍스트로 직접 가져오기 (숫자 정밀도 손실 방지)
  const res = UrlFetchApp.fetch(
    `https://api.mercadolibre.com/orders/search?seller=${getMyUserId()}&order.status=paid&sort=date_desc&limit=50&order.date_created.from=${fromStr}`,
    { headers: getHeaders(), muteHttpExceptions: true }
  );

  // JSON 파싱 전에 텍스트에서 주문번호 추출 (숫자 정밀도 손실 방지)
  const rawText = res.getContentText();
  const data = JSON.parse(rawText);
  const orders = data.results || [];

  let newCount = 0;

  for (let oi = 0; oi < orders.length; oi++) {
    const order = orders[oi];

    // 주문번호 추출 - rawText에서 직접 파싱
    const orderStr = rawText.match(new RegExp('"id":(\\d{16,})[^}]*"date_created":"' + (order.date_created || '').substring(0,10)));
    const orderId = orderStr ? orderStr[1] : String(order.id);

    if (existingIds.includes(orderId)) continue;

    const buyer = order.buyer?.nickname || '';
    const date = (order.date_created || '').substring(0, 10);
    const totalAmount = order.total_amount || 0;
    const shipId = order.shipping?.id || '';

    // 배송유형
    let deliveryType = 'ML';
    if (shipId) {
      try {
        const shipRes = UrlFetchApp.fetch(
          `https://api.mercadolibre.com/shipments/${shipId}`,
          { headers: getHeaders(), muteHttpExceptions: true }
        );
        const shipData = JSON.parse(shipRes.getContentText());
        const logisticType = shipData.logistic_type || '';
        if (logisticType === 'self_service') deliveryType = 'FLEX';
        else if (logisticType === 'fulfillment') deliveryType = 'ML풀필먼트';
      } catch(e) {
        deliveryType = 'ML';
      }
    }

    for (const item of (order.order_items || [])) {
      const sku = (item.item?.seller_sku || '').trim();
      const itemTitle = item.item?.title || '';
      const qty = item.quantity || 1;
      const unitPrice = item.unit_price || 0;
      const bundleQty = Number(bundleMap[sku]) || 1;
      const orderLink = `https://www.mercadolibre.cl/ventas/${orderId}/detalle`;

      const newRow = salesSheet.getLastRow() + 1;
      salesSheet.getRange(newRow, 1, 1, 13).setValues([[
        orderId, date, deliveryType, buyer,
        sku, itemTitle, qty, bundleQty, '',
        unitPrice, totalAmount, '⏳미처리', ''
      ]]);
      salesSheet.getRange(newRow, 9).setFormula(`=G${newRow}*H${newRow}`);
      salesSheet.getRange(newRow, 13).setFormula(`=HYPERLINK("${orderLink}","링크")`);
      newCount++;
    }
  }

  // 날짜순 정렬
  if (salesSheet.getLastRow() > 2) {
    salesSheet.getRange(2, 1, salesSheet.getLastRow() - 1, 12)
      .sort({column: 2, ascending: false});
  }

  showAlert(`✅ ML판매내역 ${newCount}개 새로 가져왔습니다.`);
}
// ==========================================
// 묶음수량 맵 가져오기 (NAOS재고 파일)
// ==========================================
function getBundleMap() {
  try {
    const naosFile = SpreadsheetApp.openById('1hw8M2WsmmTY-Obz12N3WUBi7Fv1tlrs9tGBUGvwHWrk');
    const resultSheet = naosFile.getSheetByName('🔍ML비교결과');
    if (!resultSheet) return {};

    const lastRow = resultSheet.getLastRow();
    if (lastRow < 2) return {};

    // A열=SKU, D열=묶음수량
    const data = resultSheet.getRange(2, 1, lastRow - 1, 4).getValues();
    const bundleMap = {};
    for (const row of data) {
      const sku = String(row[0] || '').trim();
      const bundleVal = row[3]; // D열 묶음수량
      const bundleQty = (bundleVal && !isNaN(Number(bundleVal))) ? Number(bundleVal) : 1;
      if (sku) bundleMap[sku] = bundleQty;
    }
    return bundleMap;
  } catch(e) {
    console.log('묶음수량 로드 오류: ' + e.message);
    return {};
  }
}
// ==========================================
// ✅ 판매완료 처리
// ==========================================
function completeMlSales() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const salesSheet = ss.getSheetByName('📊ML판매내역');
  const recordSheet = ss.getSheetByName('📋ML판매기록');

  if (!salesSheet || !recordSheet) {
    showAlert('❌ 초기 설정을 먼저 실행하세요!');
    return;
  }

  const lastRow = salesSheet.getLastRow();
  if (lastRow < 2) { showAlert('처리할 판매내역이 없습니다.'); return; }

  const data = salesSheet.getRange(2, 1, lastRow - 1, 12).getValues();

  // 미처리 항목 확인
  const pendingItems = data.filter(row => row[11] === '⏳미처리');
  if (pendingItems.length === 0) {
    showAlert('⏳미처리 항목이 없습니다.');
    return;
  }

  // 확인 팝업
  const ui = SpreadsheetApp.getUi();
  const confirm = ui.alert(
    '✅ 판매완료 처리',
    `미처리 항목 ${pendingItems.length}개를\n판매완료 처리하시겠습니까?\n\n✅ 재고 차감\n✅ 판매기록 저장\n✅ 판매내역 전체 삭제`,
    ui.ButtonSet.YES_NO
  );
  if (confirm !== ui.Button.YES) return;

  // NAOS재고 파일 열기
  const naosFile = SpreadsheetApp.openById('1hw8M2WsmmTY-Obz12N3WUBi7Fv1tlrs9tGBUGvwHWrk');
  const naosSheet = naosFile.getSheetByName('Inventario');
  if (!naosSheet) {
    showAlert('❌ Inventario 시트를 찾을 수 없습니다!');
    return;
  }
  const naosLastRow = naosSheet.getLastRow();
  const naosData = naosSheet.getRange(2, 1, naosLastRow - 1, 7).getValues();

  // 🔍ML비교결과 시트 연결
  const mlResultSheet = naosFile.getSheetByName('🔍ML비교결과');
  const mlResultSkuMap = {};
  if (mlResultSheet && mlResultSheet.getLastRow() > 1) {
    const mlResultData = mlResultSheet.getRange(2, 1, mlResultSheet.getLastRow()-1, 1).getValues();
    for (let i = 0; i < mlResultData.length; i++) {
      const sku = String(mlResultData[i][0] || '').trim();
      if (sku) mlResultSkuMap[sku] = i + 2;
    }
  }

  // NAOS SKU → 행번호 맵
  const skuRowMap = {};
  for (let i = 0; i < naosData.length; i++) {
    const sku = (naosData[i][1] || '').trim();
    if (sku) skuRowMap[sku] = i + 2;
  }

  let completedCount = 0;
  let skippedCount = 0;
  let notFoundSkus = [];
  const now = new Date();
  const recordRows = [];

  for (const row of data) {
    const orderId = String(row[0]).trim();
    const date = row[1];
    const deliveryType = row[2];
    const buyer = row[3];
    const sku = (row[4] || '').trim();
    const itemTitle = row[5];
    const qty = Number(row[6]) || 1;
    const bundleQty = Number(row[7]) || 1;
    // I열 수식 결과값 읽기
    const deductQty = Number(salesSheet.getRange(data.indexOf(row) + 2, 9).getValue()) || qty;
    const unitPrice = row[9];
    const totalAmount = row[10];
    const status = row[11];

    // ✅완료 이미 처리된것도 판매기록에 포함
    // ⏳미처리만 재고 차감
    if (status === '⏳미처리') {
      if (!sku) {
        skippedCount++;
      } else {
        // 1. NAOS재고 Inventario G열 차감 (차감수량 기준)
        if (skuRowMap[sku]) {
          const naosRow = skuRowMap[sku];
          const currentStock = Number(naosSheet.getRange(naosRow, 7).getValue()) || 0;
          const newStock = currentStock - deductQty;
          naosSheet.getRange(naosRow, 7).setValue(newStock);
          completedCount++;
        } else {
          notFoundSkus.push(`${sku} (${itemTitle})`);
          completedCount++;
        }

        // 2. 🔍ML비교결과 G열 ML재고 차감 (판매수량 기준)
        // 3. 🔍ML비교결과 I열 NAOS재고 차감 (차감수량 기준)
        if (mlResultSkuMap[sku]) {
          const mlResultRow = mlResultSkuMap[sku];

          // G열 ML재고 차감 (판매수량)
          const currentMlStock = Number(mlResultSheet.getRange(mlResultRow, 7).getValue()) || 0;
          const newMlStock = currentMlStock - qty;
          mlResultSheet.getRange(mlResultRow, 7).setValue(newMlStock < 0 ? 0 : newMlStock);

          // I열 NAOS재고 차감 (차감수량)
          const currentNaosStock = Number(mlResultSheet.getRange(mlResultRow, 9).getValue()) || 0;
          const newNaosStock = currentNaosStock - deductQty;
          mlResultSheet.getRange(mlResultRow, 9).setValue(newNaosStock);
        }
      }
    }

    // 판매기록에 모두 추가 (SKU 없는 것도 포함)
    // 재고차감여부 결정
    let deductStatus = '';
    if (status === '⏳미처리') {
      if (!sku) {
        deductStatus = '❌SKU없음';
      } else if (skuRowMap[sku]) {
        deductStatus = '✅차감완료';
      } else {
        deductStatus = '⚠️SKU불일치';
      }
    } else {
      deductStatus = '➖이미처리됨';
    }

    // 판매기록에 모두 추가
    const orderLink = `https://www.mercadolibre.cl/ventas/${orderId}/detalle`;
    recordRows.push([
      date, orderId, deliveryType, buyer,
      sku, itemTitle, qty, bundleQty, deductQty,
      unitPrice, totalAmount, now, deductStatus, orderLink
    ]);
  }

  // 판매기록 시트에 한번에 저장
  // 날짜순 정렬 후 저장
  recordRows.sort((a, b) => {
    const dateA = new Date(a[0]);
    const dateB = new Date(b[0]);
    return dateB - dateA; // 최신날짜 먼저
  });

  // 판매기록 시트에 한번에 저장
  if (recordRows.length > 0) {
    const recLastRow = recordSheet.getLastRow();
    const startRow = recLastRow + 1;
    recordSheet.getRange(startRow, 1, recordRows.length, 14).setValues(recordRows);
    // URL 하이퍼링크 설정
    for (let i = 0; i < recordRows.length; i++) {
      const url = recordRows[i][13];
      if (url) {
        recordSheet.getRange(startRow + i, 14)
          .setFormula(`=HYPERLINK("${url}","링크")`);
      }
    }

    // 색상 표시
    for (let i = 0; i < recordRows.length; i++) {
      const deductStatus = recordRows[i][12];
      let color = null;
      if (deductStatus === '✅차감완료') color = '#eafaf1';
      else if (deductStatus === '❌SKU없음') color = '#fff176';
      else if (deductStatus === '⚠️SKU불일치') color = '#fde8e8';
      else if (deductStatus === '➖이미처리됨') color = '#f2f3f4';
      if (color) {
        recordSheet.getRange(startRow + i, 1, 1, 14).setBackground(color);
      }
    }

    // 전체 판매기록 날짜순 정렬
    if (recordSheet.getLastRow() > 2) {
      recordSheet.getRange(2, 1, recordSheet.getLastRow()-1, 14)
        .sort({column: 1, ascending: false});
    }
  }

  // 완료된 주문번호 별도 저장 (중복방지용)
  let doneSheet = ss.getSheetByName('✅완료주문번호');
  if (!doneSheet) {
    doneSheet = ss.insertSheet('✅완료주문번호');
    doneSheet.getRange(1, 1).setValue('주문번호');
    doneSheet.getRange(1, 1).setBackground('#34495e').setFontColor('white').setFontWeight('bold');
  }
  
  // 완료된 주문번호 텍스트로 저장 (중복 제거)
  const existingDoneIds = doneSheet.getLastRow() > 1
    ? doneSheet.getRange(2, 1, doneSheet.getLastRow() - 1, 1)
        .getValues().flat().map(id => String(id).trim())
    : [];

  const newDoneIds = [];
  for (const row of data) {
    const orderId = String(row[0]).trim();
    if (orderId && !existingDoneIds.includes(orderId) && !newDoneIds.find(r => r[0] === orderId)) {
      newDoneIds.push([orderId]);
    }
  }

  if (newDoneIds.length > 0) {
    const doneLastRow = doneSheet.getLastRow();
    const doneRange = doneSheet.getRange(doneLastRow + 1, 1, newDoneIds.length, 1);
    doneRange.setNumberFormat('@');
    doneRange.setValues(newDoneIds);
  }

  // 📊ML판매내역 시트 전체 삭제 (헤더 제외)
  const salesLastRow = salesSheet.getLastRow();
  if (salesLastRow > 1) {
    salesSheet.deleteRows(2, salesLastRow - 1);
  }

  // 결과 메시지
  let msg = `✅ 판매완료 처리!\n\n재고 차감: ${completedCount}개\n판매기록 저장: ${recordRows.length}개\n판매내역 삭제 완료!`;
  if (notFoundSkus.length > 0) {
    msg += `\n\n⚠️ SKU 없음 (수동차감 필요):\n${notFoundSkus.join('\n')}`;
  }
  if (skippedCount > 0) {
    msg += `\n\n➖ SKU 미입력: ${skippedCount}개`;
  }
  showAlert(msg);
}

function testSalesData() {
  const userId = getMyUserId();
  const res = UrlFetchApp.fetch(
    `https://api.mercadolibre.com/orders/search?seller=${userId}&order.status=paid&sort=date_desc&limit=3`,
    { headers: getHeaders(), muteHttpExceptions: true }
  );
  const orders = JSON.parse(res.getContentText()).results || [];
  
  for (const order of orders) {
    Logger.log('=== 주문 ' + order.id + ' ===');
    Logger.log('tags: ' + JSON.stringify(order.tags));
    Logger.log('shipping: ' + JSON.stringify(order.shipping));
    Logger.log('context: ' + JSON.stringify(order.context));
    for (const item of (order.order_items || [])) {
      Logger.log('SKU: ' + item.item?.seller_sku);
    }
  }

  // 묶음수량 맵 확인
  const bundleMap = getBundleMap();
  Logger.log('\n=== 묶음맵 샘플 ===');
  let count = 0;
  for (const key of Object.keys(bundleMap)) {
    Logger.log(`SKU: "${key}" → 묶음수량: ${bundleMap[key]}`);
    if (++count >= 5) break;
  }
}

// ==========================================
// 🔄 SKU불일치 재처리
// ==========================================
function reprocessSkuMismatch() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const recordSheet = ss.getSheetByName('📋ML판매기록');

  if (!recordSheet || recordSheet.getLastRow() < 2) {
    showAlert('📋ML판매기록에 데이터가 없습니다.');
    return;
  }

  const lastRow = recordSheet.getLastRow();
  const data = recordSheet.getRange(2, 1, lastRow - 1, 13).getValues();

  // ⚠️SKU불일치 항목 찾기
  const mismatchRows = [];
  for (let i = 0; i < data.length; i++) {
    if (data[i][12] === '⚠️SKU불일치') {
      mismatchRows.push({ rowIndex: i + 2, data: data[i] });
    }
  }

  if (mismatchRows.length === 0) {
    showAlert('⚠️SKU불일치 항목이 없습니다.');
    return;
  }

  // NAOS재고 연결
  const naosFile = SpreadsheetApp.openById('1hw8M2WsmmTY-Obz12N3WUBi7Fv1tlrs9tGBUGvwHWrk');
  const naosSheet = naosFile.getSheetByName('Inventario');
  const naosLastRow = naosSheet.getLastRow();
  const naosData = naosSheet.getRange(2, 1, naosLastRow - 1, 7).getValues();

  // SKU → 행번호 맵 (최신 SKU 기준)
  const skuRowMap = {};
  for (let i = 0; i < naosData.length; i++) {
    const sku = (naosData[i][1] || '').trim();
    if (sku) skuRowMap[sku] = i + 2;
  }

  let fixedCount = 0;
  let stillMismatch = [];

  for (const item of mismatchRows) {
    const sku = (item.data[4] || '').trim();
    const deductQty = Number(item.data[8]) || 1;
    const itemTitle = item.data[5];

    if (sku && skuRowMap[sku]) {
      // 재고 차감
      const naosRow = skuRowMap[sku];
      const currentStock = Number(naosSheet.getRange(naosRow, 7).getValue()) || 0;
      const newStock = currentStock - deductQty;
      naosSheet.getRange(naosRow, 7).setValue(newStock);

      // 재고차감여부 업데이트
      recordSheet.getRange(item.rowIndex, 13).setValue('✅차감완료(재처리)');
      recordSheet.getRange(item.rowIndex, 1, 1, 13).setBackground('#eafaf1');
      fixedCount++;
    } else {
      stillMismatch.push(`${sku} (${itemTitle})`);
    }
  }

  let msg = `✅ 재처리 완료!\n\n차감완료: ${fixedCount}개`;
  if (stillMismatch.length > 0) {
    msg += `\n\n⚠️ 여전히 SKU불일치 (NAOS재고 SKU 확인필요):\n${stillMismatch.join('\n')}`;
  }
  showAlert(msg);
}


function registerExistingOrders() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const salesSheet = ss.getSheetByName('📊ML판매내역');
  
  if (!salesSheet || salesSheet.getLastRow() < 2) {
    showAlert('판매내역이 없습니다.');
    return;
  }

  let doneSheet = ss.getSheetByName('✅완료주문번호');
  if (!doneSheet) {
    doneSheet = ss.insertSheet('✅완료주문번호');
    doneSheet.getRange(1, 1).setValue('주문번호');
    doneSheet.getRange(1, 1).setBackground('#34495e').setFontColor('white').setFontWeight('bold');
  }

  // 기존 완료주문번호 목록
  const doneLastRow = doneSheet.getLastRow();
  const existingDoneIds = doneLastRow > 1
    ? doneSheet.getRange(2, 1, doneLastRow - 1, 1).getValues().flat()
        .map(id => String(id).replace(/'/g,'').trim())
    : [];

  // 판매내역 전체 행 읽기
  const totalRows = salesSheet.getLastRow() - 1;
  const data = salesSheet.getRange(2, 1, totalRows, 2).getValues();
  
  Logger.log('판매내역 총 행수: ' + totalRows);
  
  const newIds = [];
  for (const row of data) {
    const rawId = row[0];
    const date = String(row[1]).trim();
    
    // 주문번호 추출 (숫자 정밀도 손실 고려)
    const orderId = String(rawId).replace(/'/g,'').trim();
    
    Logger.log(`주문번호: "${orderId}" / 날짜: "${date}"`);
    
    // 모든 주문번호 등록 (날짜 무관)
    if (orderId && orderId !== '' && !existingDoneIds.includes(orderId) && 
        !newIds.find(r => r[0] === orderId)) {
      newIds.push([orderId]);
    }
  }

  if (newIds.length > 0) {
    const range = doneSheet.getRange(doneLastRow + 1, 1, newIds.length, 1);
    range.setNumberFormat('@');
    range.setValues(newIds);
    showAlert(`✅ ${newIds.length}개 주문번호 등록완료!`);
  } else {
    showAlert('등록할 새 주문번호가 없습니다.\n(이미 모두 등록되어 있거나 판매내역이 비어있습니다)');
  }
}

function checkDoneSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // 완료주문번호 확인
  const doneSheet = ss.getSheetByName('✅완료주문번호');
  if (!doneSheet || doneSheet.getLastRow() < 2) {
    Logger.log('✅완료주문번호 시트 비어있음!');
  } else {
    Logger.log('완료주문번호 수: ' + (doneSheet.getLastRow() - 1));
    const ids = doneSheet.getRange(2, 1, Math.min(5, doneSheet.getLastRow()-1), 1).getValues();
    Logger.log('샘플: ' + JSON.stringify(ids));
  }

  // 판매기록 확인
  const recordSheet = ss.getSheetByName('📋ML판매기록');
  if (!recordSheet || recordSheet.getLastRow() < 2) {
    Logger.log('판매기록 비어있음!');
  } else {
    Logger.log('판매기록 수: ' + (recordSheet.getLastRow() - 1));
    const ids = recordSheet.getRange(2, 2, Math.min(5, recordSheet.getLastRow()-1), 1).getValues();
    Logger.log('판매기록 주문번호 샘플: ' + JSON.stringify(ids));
  }
}

function checkMismatch() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // 판매내역에 새로 들어온 2개 주문번호 확인
  const salesSheet = ss.getSheetByName('📊ML판매내역');
  if (!salesSheet || salesSheet.getLastRow() < 2) {
    Logger.log('판매내역 없음');
    return;
  }
  
  const salesData = salesSheet.getRange(2, 1, salesSheet.getLastRow()-1, 2).getValues();
  Logger.log('=== 판매내역 주문번호 ===');
  for (const row of salesData) {
    Logger.log(`"${row[0]}" / 날짜: ${row[1]}`);
  }

  // 완료주문번호 확인
  const doneSheet = ss.getSheetByName('✅완료주문번호');
  const doneData = doneSheet.getRange(2, 1, doneSheet.getLastRow()-1, 1).getValues();
  Logger.log('\n=== 완료주문번호 마지막 10개 ===');
  const last10 = doneData.slice(-10);
  for (const row of last10) {
    Logger.log(`"${row[0]}"`);
  }
}


function sortRecordSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const recordSheet = ss.getSheetByName('📋ML판매기록');
  
  if (!recordSheet || recordSheet.getLastRow() < 2) {
    showAlert('판매기록이 없습니다.');
    return;
  }

  const lastRow = recordSheet.getLastRow();
  recordSheet.getRange(2, 1, lastRow - 1, 14)
    .sort({column: 1, ascending: false});

  SpreadsheetApp.getUi().alert('✅ 날짜순 정렬 완료!');
}

function testRecentOrders() {
  const userId = getMyUserId();
  const today = new Date();
  const from = new Date(today);
  from.setDate(from.getDate() - 7);
  const fromStr = from.toISOString().split('T')[0] + 'T00:00:00.000-04:00';

  const res = UrlFetchApp.fetch(
    `https://api.mercadolibre.com/orders/search?seller=${userId}&order.status=paid&sort=date_desc&limit=50&order.date_created.from=${fromStr}`,
    { headers: getHeaders(), muteHttpExceptions: true }
  );

  const rawText = res.getContentText();
  const orders = JSON.parse(rawText).results || [];

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const doneSheet = ss.getSheetByName('✅완료주문번호');
  const doneIds = doneSheet && doneSheet.getLastRow() > 1
    ? doneSheet.getRange(2, 1, doneSheet.getLastRow()-1, 1).getValues().flat().map(id => String(id).trim())
    : [];

  Logger.log('최근 7일 주문 총: ' + orders.length);
  let newCount = 0;
  for (const order of orders) {
    const orderRaw = JSON.stringify(order);
    const idMatch = orderRaw.match(/"id":(\d{16,})/);
    const orderId = idMatch ? idMatch[1] : String(order.id);
    const date = (order.date_created || '').substring(0, 10);
    const isDone = doneIds.includes(orderId);
    if (!isDone) {
      Logger.log(`⏳미처리 : ${orderId} / ${date}`);
      newCount++;
    }
  }
  Logger.log('미처리 주문 수: ' + newCount);
}

function checkToken() {
  const token = getConfig('ACCESS_TOKEN');
  Logger.log('토큰 앞 30자: ' + String(token).substring(0, 30));
  Logger.log('토큰 길이: ' + String(token).length);
}

function testHighlightsV4() {
  const accessToken = getConfig('ACCESS_TOKEN');
  const results = [];

  // 테스트 1: 토큰 유효성
  const meRes = UrlFetchApp.fetch('https://api.mercadolibre.com/users/me', {
    headers: { 'Authorization': 'Bearer ' + accessToken },
    muteHttpExceptions: true
  });
  results.push('테스트1 토큰: ' + meRes.getResponseCode());

  Utilities.sleep(300);

  // 테스트 2: /highlights 베스트셀러
  const hlRes = UrlFetchApp.fetch('https://api.mercadolibre.com/highlights/MLC/category/MLC1574', {
    headers: { 'Authorization': 'Bearer ' + accessToken },
    muteHttpExceptions: true
  });
  results.push('테스트2 highlights: ' + hlRes.getResponseCode());
  if (hlRes.getResponseCode() === 200) {
    results.push('✅ 내용: ' + hlRes.getContentText().substring(0, 200));
  } else {
    results.push('❌ ' + hlRes.getContentText().substring(0, 200));
  }

  Utilities.sleep(300);

  // 테스트 3: /trends
  const trRes = UrlFetchApp.fetch('https://api.mercadolibre.com/trends/MLC', {
    headers: { 'Authorization': 'Bearer ' + accessToken },
    muteHttpExceptions: true
  });
  results.push('테스트3 trends: ' + trRes.getResponseCode());

  Utilities.sleep(300);

  // 테스트 4: /search
  const srRes = UrlFetchApp.fetch('https://api.mercadolibre.com/sites/MLC/search?q=bisagra&limit=3', {
    headers: { 'Authorization': 'Bearer ' + accessToken },
    muteHttpExceptions: true
  });
  results.push('테스트4 search: ' + srRes.getResponseCode());

  Logger.log(results.join('\n'));
  SpreadsheetApp.getUi().alert(results.join('\n'));
}

function debugRefresh() {
  const clientId = getConfig('CLIENT_ID');
  const secretKey = getConfig('SECRET_KEY');
  const refreshToken = getConfig('REFRESH_TOKEN');
  
  Logger.log('CLIENT_ID: ' + String(clientId).substring(0, 15));
  Logger.log('SECRET_KEY: ' + String(secretKey).substring(0, 10));
  Logger.log('REFRESH_TOKEN 앞 20자: ' + String(refreshToken).substring(0, 20));
  Logger.log('REFRESH_TOKEN 길이: ' + String(refreshToken).length);
  
  const res = UrlFetchApp.fetch('https://api.mercadolibre.com/oauth/token', {
    method: 'post',
    payload: {
      grant_type: 'refresh_token',
      client_id: clientId,
      client_secret: secretKey,
      refresh_token: refreshToken
    },
    muteHttpExceptions: true
  });
  
  Logger.log('갱신 응답코드: ' + res.getResponseCode());
  Logger.log('갱신 응답내용: ' + res.getContentText());
  
  SpreadsheetApp.getUi().alert(
    'CLIENT_ID: ' + String(clientId).substring(0, 15) + '...\n' +
    'REFRESH_TOKEN 길이: ' + String(refreshToken).length + '\n\n' +
    '갱신 응답코드: ' + res.getResponseCode() + '\n' +
    '갱신 응답내용:\n' + res.getContentText().substring(0, 300)
  );
}

function testItemDetail() {
  const token = getConfig('ACCESS_TOKEN');
  
  // 베스트셀러에서 나온 실제 ID로 테스트
  const testIds = ['MLC57718763', 'MLCU3057614156', 'MLC32171120'];
  
  let msg = '';
  
  // 방법1: 배치
  const batchRes = UrlFetchApp.fetch(
    'https://api.mercadolibre.com/items?ids=' + testIds.join(',') + '&attributes=id,title,price,permalink',
    { headers: { 'Authorization': 'Bearer ' + token }, muteHttpExceptions: true }
  );
  msg += '배치 응답코드: ' + batchRes.getResponseCode() + '\n';
  msg += '배치 내용: ' + batchRes.getContentText().substring(0, 300) + '\n\n';
  
  Utilities.sleep(500);
  
  // 방법2: 단건
  const singleRes = UrlFetchApp.fetch(
    'https://api.mercadolibre.com/items/MLC57718763?attributes=id,title,price,permalink',
    { headers: { 'Authorization': 'Bearer ' + token }, muteHttpExceptions: true }
  );
  msg += '단건 응답코드: ' + singleRes.getResponseCode() + '\n';
  msg += '단건 내용: ' + singleRes.getContentText().substring(0, 300) + '\n';
  
  Logger.log(msg);
  SpreadsheetApp.getUi().alert(msg);
}

function testHighlightsRaw() {
  const token = getConfig('ACCESS_TOKEN');
  
  const res = UrlFetchApp.fetch(
    'https://api.mercadolibre.com/highlights/MLC/category/MLC175547',
    { headers: { 'Authorization': 'Bearer ' + token }, muteHttpExceptions: true }
  );
  
  const body = res.getContentText();
  Logger.log(body);
  
  // 첫 3개 항목만 표시
  const json = JSON.parse(body);
  const items = json.content || [];
  let msg = '전체 항목수: ' + items.length + '\n\n';
  msg += '첫 3개 원본:\n' + JSON.stringify(items.slice(0, 3), null, 2);
  
  SpreadsheetApp.getUi().alert(msg.substring(0, 1000));
}

function testUserProduct() {
  const token = getConfig('ACCESS_TOKEN');
  
  // MLCU 타입 테스트
  const id = 'MLCU3057614156';
  let msg = '';
  
  // 방법1: /products
  const r1 = UrlFetchApp.fetch(
    `https://api.mercadolibre.com/products/${id}?attributes=id,name,buy_box_winner`,
    { headers: { 'Authorization': 'Bearer ' + token }, muteHttpExceptions: true }
  );
  msg += '/products 응답: ' + r1.getResponseCode() + '\n';
  msg += r1.getContentText().substring(0, 400) + '\n\n';
  
  Utilities.sleep(300);
  
  // 방법2: /items (catalog item)
  const r2 = UrlFetchApp.fetch(
    `https://api.mercadolibre.com/items?ids=${id}`,
    { headers: { 'Authorization': 'Bearer ' + token }, muteHttpExceptions: true }
  );
  msg += '/items 응답: ' + r2.getResponseCode() + '\n';
  msg += r2.getContentText().substring(0, 400);
  
  Logger.log(msg);
  SpreadsheetApp.getUi().alert(msg.substring(0, 800));
}

function testUserProduct2() {
  const token = getConfig('ACCESS_TOKEN');
  const id = 'MLCU3057614156';
  let msg = '';

  // 방법1: catalog_product_id로 검색
  const r1 = UrlFetchApp.fetch(
    `https://api.mercadolibre.com/sites/MLC/search?catalog_product_id=${id}&limit=1`,
    { headers: { 'Authorization': 'Bearer ' + token }, muteHttpExceptions: true }
  );
  msg += '방법1 catalog검색: ' + r1.getResponseCode() + '\n';
  msg += r1.getContentText().substring(0, 400) + '\n\n';

  Utilities.sleep(300);

  // 방법2: highlights에서 item_id 필드 확인
  const r2 = UrlFetchApp.fetch(
    `https://api.mercadolibre.com/highlights/MLC/category/MLC175547`,
    { headers: { 'Authorization': 'Bearer ' + token }, muteHttpExceptions: true }
  );
  const body = JSON.parse(r2.getContentText());
  const mlcu = body.content.find(i => i.id === id);
  msg += '방법2 highlights 원본:\n' + JSON.stringify(mlcu) + '\n\n';

  Utilities.sleep(300);

  // 방법3: product catalog
  const r3 = UrlFetchApp.fetch(
    `https://api.mercadolibre.com/products/search?site_id=MLC&catalog_product_id=${id}`,
    { headers: { 'Authorization': 'Bearer ' + token }, muteHttpExceptions: true }
  );
  msg += '방법3 products검색: ' + r3.getResponseCode() + '\n';
  msg += r3.getContentText().substring(0, 300);

  Logger.log(msg);
  SpreadsheetApp.getUi().alert(msg.substring(0, 900));
}
