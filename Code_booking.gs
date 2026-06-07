/**
 * 《超限戰》台南訂票系統 - Google Apps Script
 *
 * 使用方式：
 * 1. 建立一個新的 Google Sheet
 * 2. 擴充功能 → Apps Script
 * 3. 貼上本檔內容
 * 4. 修改 CONFIG.spreadsheetId 為您的 Google Sheet ID
 * 5. 執行 setup()
 * 6. 部署 → 新增部署作業 → 網頁應用程式
 *    - 執行身分：我
 *    - 存取權：任何人
 * 7. 複製 Web App URL，貼到 index.html 的 GAS_WEBAPP_URL
 */

const CONFIG = {
  spreadsheetId: '請貼上您的 Google Sheet ID',
  sheetName: '訂票資料',
  unitPrice: 150,
  maxTicketsPerOrder: 2,
  eventName: '《超限戰》台南放映',
  venue: '勞工育樂中心',
  address: '台南市南區南門路261號',
  contactPerson: '陳小姐',
  contactPhone: '0930-885-835'
};

const HEADERS = [
  '建立時間',
  '訂票編號',
  '訂票人姓名',
  '聯絡電話',
  '票數',
  '單價',
  '總金額',
  'LINE 或 Email',
  '訊息來源',
  '備註',
  '確認狀態',
  '付款狀態',
  '活動名稱',
  '場地',
  '地址',
  '頁面網址',
  'User Agent',
  '處理備註'
];

/**
 * 第一次使用請先執行 setup()
 * 會自動建立訂票資料分頁與欄位。
 */
function setup() {
  const ss = SpreadsheetApp.openById(CONFIG.spreadsheetId);
  let sheet = ss.getSheetByName(CONFIG.sheetName);

  if (!sheet) {
    sheet = ss.insertSheet(CONFIG.sheetName);
  }

  sheet.clear();
  sheet.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]);

  const headerRange = sheet.getRange(1, 1, 1, HEADERS.length);
  headerRange
    .setFontWeight('bold')
    .setBackground('#111827')
    .setFontColor('#ffffff')
    .setHorizontalAlignment('center');

  sheet.setFrozenRows(1);
  sheet.autoResizeColumns(1, HEADERS.length);

  // 常用欄寬
  sheet.setColumnWidth(1, 160); // 建立時間
  sheet.setColumnWidth(2, 150); // 訂票編號
  sheet.setColumnWidth(3, 140); // 姓名
  sheet.setColumnWidth(4, 140); // 電話
  sheet.setColumnWidth(8, 180); // LINE 或 Email
  sheet.setColumnWidth(10, 240); // 備註
  sheet.setColumnWidth(18, 240); // 處理備註

  // 狀態欄預設格式
  const statusRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['新訂票', '已確認', '取消', '未聯絡上'], true)
    .setAllowInvalid(false)
    .build();

  const paymentRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['未付款', '已付款', '現場付款', '不需付款'], true)
    .setAllowInvalid(false)
    .build();

  sheet.getRange(2, 11, 500, 1).setDataValidation(statusRule);
  sheet.getRange(2, 12, 500, 1).setDataValidation(paymentRule);

  return 'setup完成：已建立「' + CONFIG.sheetName + '」欄位。';
}

/**
 * 測試 Web App 是否正常。
 */
function doGet(e) {
  return jsonOutput({
    ok: true,
    message: '《超限戰》訂票 API 正常運作',
    eventName: CONFIG.eventName,
    venue: CONFIG.venue,
    unitPrice: CONFIG.unitPrice
  });
}

/**
 * 接收 GitHub Pages 網頁送出的訂票資料。
 */
function doPost(e) {
  try {
    const data = parseRequestBody(e);
    const cleaned = validateAndNormalize(data);
    const bookingId = createBookingId();

    const ss = SpreadsheetApp.openById(CONFIG.spreadsheetId);
    let sheet = ss.getSheetByName(CONFIG.sheetName);
    if (!sheet) {
      setup();
      sheet = ss.getSheetByName(CONFIG.sheetName);
    }

    const now = new Date();

    const row = [
      now,
      bookingId,
      cleaned.name,
      cleaned.phone,
      cleaned.ticketQty,
      CONFIG.unitPrice,
      cleaned.totalAmount,
      cleaned.contact,
      cleaned.source,
      cleaned.note,
      '新訂票',
      '未付款',
      CONFIG.eventName,
      CONFIG.venue,
      CONFIG.address,
      cleaned.pageUrl,
      cleaned.userAgent,
      ''
    ];

    sheet.appendRow(row);

    return jsonOutput({
      ok: true,
      message: '訂票成功',
      bookingId: bookingId,
      ticketQty: cleaned.ticketQty,
      totalAmount: cleaned.totalAmount
    });
  } catch (err) {
    return jsonOutput({
      ok: false,
      message: err.message || '系統錯誤'
    });
  }
}

/**
 * 解析前端送來的 JSON。
 * 前端用 text/plain 送出，是為了降低 CORS 預檢問題。
 */
function parseRequestBody(e) {
  if (!e || !e.postData || !e.postData.contents) {
    throw new Error('沒有收到資料');
  }

  try {
    return JSON.parse(e.postData.contents);
  } catch (err) {
    throw new Error('資料格式錯誤');
  }
}

/**
 * 檢查與整理資料。
 */
function validateAndNormalize(data) {
  const name = String(data.name || '').trim();
  const phone = String(data.phone || '').trim();
  const contact = String(data.contact || '').trim();
  const source = String(data.source || '').trim();
  const note = String(data.note || '').trim();
  const pageUrl = String(data.pageUrl || '').trim();
  const userAgent = String(data.userAgent || '').trim();

  const ticketQty = Number(data.ticketQty);
  const unitPrice = CONFIG.unitPrice;
  const totalAmount = ticketQty * unitPrice;

  if (!name) throw new Error('請填寫訂票人姓名');
  if (!phone) throw new Error('請填寫聯絡電話');
  if (!contact) throw new Error('請填寫 LINE 或 Email');
  if (!source) throw new Error('請選擇訊息來源');

  if (!Number.isInteger(ticketQty) || ticketQty < 1 || ticketQty > CONFIG.maxTicketsPerOrder) {
    throw new Error('票數需為 1 到 ' + CONFIG.maxTicketsPerOrder + ' 張');
  }

  // 簡單電話檢查：保留彈性，不過度限制格式。
  const phoneDigits = phone.replace(/\D/g, '');
  if (phoneDigits.length < 8) {
    throw new Error('聯絡電話格式可能不正確');
  }

  return {
    name,
    phone,
    ticketQty,
    totalAmount,
    contact,
    source,
    note,
    pageUrl,
    userAgent
  };
}

/**
 * 建立訂票編號。
 * 格式：TUW-年月日-時分秒-亂數
 */
function createBookingId() {
  const now = new Date();
  const tz = Session.getScriptTimeZone() || 'Asia/Taipei';
  const stamp = Utilities.formatDate(now, tz, 'yyyyMMdd-HHmmss');
  const random = Math.floor(Math.random() * 900 + 100);
  return 'TUW-' + stamp + '-' + random;
}

/**
 * JSON 回傳。
 */
function jsonOutput(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
