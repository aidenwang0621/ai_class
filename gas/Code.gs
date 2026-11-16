/**
 * 金邊鍋貼 — 訂單接收（Google Apps Script，綁定在 Google 試算表上）
 *
 * 安裝步驟見 gas/README.md。重點：
 *   1. 在試算表選「擴充功能 → Apps Script」，把這個檔案整份貼上。
 *   2. 執行一次 setup()（會建立「訂單」「訂單明細」兩張工作表）。
 *   3. 部署 → 新增部署作業 → 網路應用程式；執行身分「我」，存取權「所有人」。
 *   4. 把得到的 /exec 網址貼到 index.html 的 GAS_URL。
 */

// 想在每筆新訂單時收到 Email 通知，就填入信箱；留空則不寄信。
const NOTIFY_EMAIL = '';

const SHEET_ORDERS = '訂單';
const SHEET_ITEMS = '訂單明細';
const TIMEZONE = 'Asia/Taipei';

// 價格以這裡為準（不採用網頁傳來的金額）。改菜單時，index.html 的 catalog 也要一起改。
const CATALOG = {
  original: ['招牌鮮肉鍋貼', 88],
  chive: ['韭香鮮肉鍋貼', 96],
  chili: ['椒麻紅油鍋貼', 108],
  shrimp: ['鮮蝦豬肉水餃', 128],
  cabbage: ['高麗菜鮮肉水餃', 88],
  soup: ['酸辣湯', 55],
  tea: ['無糖焙香烏龍', 40],
  cucumber: ['蒜香脆小黃瓜', 48],
  tofu: ['胡麻時蔬', 60],
  gift: ['招牌冷凍鍋貼禮盒', 580]
};
const STORES = ['中山旗艦店', '信義店'];
const MODES = ['內用', '外帶', '宅配'];

const ORDER_HEADERS = ['訂單編號', '下單時間', '取餐方式', '門市', '姓名', '電話', '收件地址', '品項', '件數', '總計', '狀態', '備註', '請求代碼'];
const ITEM_HEADERS = ['訂單編號', '下單時間', '商品代碼', '品名', '單價', '數量', '小計'];

/** 第一次使用時手動執行：建立工作表、標題列與格式。重複執行不會清掉資料。 */
function setup() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  prepareSheet_(ss, SHEET_ORDERS, ORDER_HEADERS, [6, 7]);
  prepareSheet_(ss, SHEET_ITEMS, ITEM_HEADERS, []);
  const orders = ss.getSheetByName(SHEET_ORDERS);
  const rule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['新訂單', '製作中', '已完成', '已取消'], true).build();
  orders.getRange(2, 11, orders.getMaxRows() - 1, 1).setDataValidation(rule);
}

function prepareSheet_(ss, name, headers, textColumns) {
  const sheet = ss.getSheetByName(name) || ss.insertSheet(name);
  sheet.getRange(1, 1, 1, headers.length).setValues([headers])
    .setFontWeight('bold').setBackground('#2d211b').setFontColor('#ffc591');
  sheet.setFrozenRows(1);
  // 電話、地址用純文字格式，避免開頭的 0 被吃掉。
  textColumns.forEach(function (col) { sheet.getRange(1, col, sheet.getMaxRows(), 1).setNumberFormat('@'); });
  return sheet;
}

/** 用瀏覽器直接開 /exec 網址時的健康檢查。 */
function doGet() {
  return json_({ ok: true, service: 'golden-edge-orders' });
}

/** 網站結帳時呼叫：驗證 → 寫入試算表 → 回傳訂單編號。 */
function doPost(e) {
  let data;
  try {
    data = JSON.parse(e.postData.contents);
  } catch (err) {
    return json_({ ok: false, error: '資料格式錯誤' });
  }
  // 機器人常會填滿所有欄位；真人看不到這個欄位。
  if (data.website) return json_({ ok: true, orderId: '-' });

  let order;
  try {
    order = validate_(data);
  } catch (err) {
    return json_({ ok: false, error: err.message });
  }

  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(20000);
  } catch (err) {
    return json_({ ok: false, error: '系統忙線中，請稍後再試' });
  }
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(SHEET_ORDERS) || prepareSheet_(ss, SHEET_ORDERS, ORDER_HEADERS, [6, 7]);
    const itemSheet = ss.getSheetByName(SHEET_ITEMS) || prepareSheet_(ss, SHEET_ITEMS, ITEM_HEADERS, []);

    // 同一筆訂單因網路重試而重送時，回傳原本的編號，不重複寫入。
    const last = sheet.getLastRow();
    if (order.clientId && last > 1) {
      const from = Math.max(2, last - 199);
      const recent = sheet.getRange(from, 1, last - from + 1, ORDER_HEADERS.length).getValues();
      for (let i = recent.length - 1; i >= 0; i--) {
        if (recent[i][12] === order.clientId) return json_({ ok: true, orderId: recent[i][0], total: recent[i][9], duplicate: true });
      }
    }

    const now = new Date();
    const orderId = 'GE' + Utilities.formatDate(now, TIMEZONE, 'yyMMdd') + '-' + ('000' + last).slice(-4);
    sheet.appendRow([
      orderId, now, order.mode, order.store, order.name, order.phone, order.address,
      order.lines.map(function (l) { return l.name + ' × ' + l.qty; }).join('、'),
      order.count, order.total, '新訂單', '', order.clientId
    ]);
    const rows = order.lines.map(function (l) { return [orderId, now, l.id, l.name, l.price, l.qty, l.price * l.qty]; });
    itemSheet.getRange(itemSheet.getLastRow() + 1, 1, rows.length, ITEM_HEADERS.length).setValues(rows);
    SpreadsheetApp.flush();

    if (NOTIFY_EMAIL) {
      try {
        MailApp.sendEmail(NOTIFY_EMAIL, '【金邊鍋貼】新訂單 ' + orderId,
          [orderId, order.mode + (order.store ? ' · ' + order.store : ''), order.name + ' ' + order.phone, order.address,
            order.lines.map(function (l) { return l.name + ' × ' + l.qty; }).join('\n'), '總計 NT$ ' + order.total]
            .filter(String).join('\n'));
      } catch (mailErr) { /* 寄信失敗不影響訂單 */ }
    }
    return json_({ ok: true, orderId: orderId, total: order.total });
  } catch (err) {
    return json_({ ok: false, error: '訂單寫入失敗，請稍後再試' });
  } finally {
    lock.releaseLock();
  }
}

/** 檢查欄位並用伺服器端的價格重新計算；有問題就丟出給顧客看的訊息。 */
function validate_(data) {
  const mode = String(data.mode || '');
  if (MODES.indexOf(mode) < 0) throw new Error('取餐方式不正確');

  if (!Array.isArray(data.items) || !data.items.length || data.items.length > 20) throw new Error('購物車是空的');
  const seen = {};
  let count = 0, total = 0;
  const lines = data.items.map(function (row) {
    const id = String(row && row.id), qty = Number(row && row.qty);
    if (!Object.prototype.hasOwnProperty.call(CATALOG, id) || seen[id]) throw new Error('商品不存在');
    if (!(qty >= 1 && qty <= 99) || Math.floor(qty) !== qty) throw new Error('數量不正確');
    seen[id] = true;
    count += qty;
    total += CATALOG[id][1] * qty;
    return { id: id, name: CATALOG[id][0], price: CATALOG[id][1], qty: qty };
  });

  // 與網站相同的規則：禮盒不能內用；宅配只限單買冷凍禮盒。
  const hasGift = !!seen.gift, onlyGift = hasGift && lines.length === 1;
  if (mode === '內用' && hasGift) throw new Error('冷凍禮盒無法內用');
  if (mode === '宅配' && !onlyGift) throw new Error('宅配僅適用冷凍禮盒');

  const name = clean_(data.name, 50), phone = clean_(data.phone, 24);
  if (!name) throw new Error('請填寫聯絡姓名');
  if (!/^[0-9+\-() #]{8,24}$/.test(phone)) throw new Error('請填寫正確的聯絡電話');

  let store = '', address = '';
  if (mode === '宅配') {
    address = clean_(data.address, 160);
    if (!address) throw new Error('請填寫收件地址');
  } else {
    store = STORES[Number(data.store)];
    if (!store) throw new Error('請選擇取餐門市');
  }
  return {
    mode: mode, store: store, address: safe_(address), name: safe_(name), phone: "'" + phone,
    lines: lines, count: count, total: total, clientId: clean_(data.clientId, 64).replace(/[^A-Za-z0-9-]/g, '')
  };
}

function clean_(value, max) {
  return String(value == null ? '' : value).replace(/[\u0000-\u001f\u007f]/g, ' ').trim().slice(0, max);
}

/** 避免顧客輸入的文字被試算表當成公式執行。 */
function safe_(text) {
  return /^[=+\-@]/.test(text) ? "'" + text : text;
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
