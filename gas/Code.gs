/**
 * 金煎 GINJIAN 線上訂單後端（Google Apps Script）
 *
 * 功能：
 *  - 接收官網結帳送出的訂單（doPost），寫入 Google 試算表
 *  - 依「商品」工作表重新計算金額，不信任前端傳來的價格
 *  - 產生訂單編號（GJ + 日期 + 流水號）
 *  - 新訂單 Email 通知店家；可選擇寄確認信給顧客
 *  - 「訂單」工作表的「狀態」欄有下拉選單，方便門市管理
 *
 * 第一次使用：在 Apps Script 編輯器執行 setup()，再依 README 部署為網頁應用程式。
 */

// ===== 設定 =====
const CONFIG = {
  SHEET_ORDERS: "訂單",
  SHEET_ITEMS: "訂單明細",
  SHEET_PRODUCTS: "商品",
  SHEET_STORES: "門市",
  DELIVERY_FEE: 160,          // 冷凍宅配運費
  MAX_LINE_ITEMS: 30,         // 單筆訂單最多品項數
  MAX_QTY: 99,                // 單一品項最多數量
  STATUSES: ["新訂單", "製作中", "可取餐", "已出貨", "已完成", "已取消"],
  TIMEZONE: "Asia/Taipei"
};

// 指令碼屬性（專案設定 → 指令碼屬性）：
//   NOTIFY_EMAIL      新訂單通知寄到這個信箱（可用逗號分隔多個）；不填就不寄
//   SEND_CUSTOMER_EMAIL  填 true 才會寄訂單確認信給顧客
function prop_(key) {
  return PropertiesService.getScriptProperties().getProperty(key) || "";
}

const ORDER_HEADERS = ["訂單編號", "建立時間", "狀態", "取餐方式", "門市", "宅配地址",
  "姓名", "電話", "Email", "備註", "商品明細", "商品金額", "運費", "合計"];
const ITEM_HEADERS = ["訂單編號", "建立時間", "商品ID", "品名", "單價", "數量", "小計"];
const PRODUCT_HEADERS = ["商品ID", "品名", "單價", "冷凍宅配", "上架"];
const STORE_HEADERS = ["門市ID", "門市名稱"];

const SEED_PRODUCTS = [
  [1, "經典鮮肉鍋貼", 85, false, true],
  [2, "松露野菇鍋貼", 120, false, true],
  [3, "韓式泡菜鍋貼", 95, false, true],
  [4, "鮮蝦韭黃鍋貼", 130, false, true],
  [5, "鮮肉高麗菜水餃", 80, false, true],
  [6, "鮮蝦玉米水餃", 110, false, true],
  [7, "暖心酸辣湯", 45, false, true],
  [8, "清香無糖烏龍", 40, false, true],
  [9, "職人秘製泡菜", 45, false, true],
  [10, "香蒜涼拌小黃瓜", 40, false, true],
  [11, "招牌招財鍋貼禮盒", 599, true, true]
];
const SEED_STORES = [
  ["flagship", "台北旗艦店"],
  ["xinyi", "信義概念店"],
  ["taichung", "台中勤美店"]
];
const MODE_NAMES = { pickup: "外帶自取", dinein: "門市內用", delivery: "冷凍宅配" };

// ===== 初始化：建立工作表、標題列、商品與門市資料 =====
function setup() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const orders = ensureSheet_(ss, CONFIG.SHEET_ORDERS, ORDER_HEADERS);
  ensureSheet_(ss, CONFIG.SHEET_ITEMS, ITEM_HEADERS);
  const products = ensureSheet_(ss, CONFIG.SHEET_PRODUCTS, PRODUCT_HEADERS);
  const stores = ensureSheet_(ss, CONFIG.SHEET_STORES, STORE_HEADERS);

  if (products.getLastRow() < 2) {
    products.getRange(2, 1, SEED_PRODUCTS.length, PRODUCT_HEADERS.length).setValues(SEED_PRODUCTS);
  }
  if (stores.getLastRow() < 2) {
    stores.getRange(2, 1, SEED_STORES.length, STORE_HEADERS.length).setValues(SEED_STORES);
  }

  // 狀態欄下拉選單
  const statusRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(CONFIG.STATUSES, true).setAllowInvalid(false).build();
  orders.getRange(2, 3, orders.getMaxRows() - 1, 1).setDataValidation(statusRule);
  orders.getRange(2, 2, orders.getMaxRows() - 1, 1).setNumberFormat("yyyy/mm/dd hh:mm");
  orders.getRange(2, 12, orders.getMaxRows() - 1, 3).setNumberFormat("#,##0");

  ss.setSpreadsheetTimeZone(CONFIG.TIMEZONE);
  safeToast_(ss, "金煎訂單系統初始化完成");
}

function ensureSheet_(ss, name, headers) {
  const sheet = ss.getSheetByName(name) || ss.insertSheet(name);
  if (sheet.getLastRow() === 0) {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers])
      .setFontWeight("bold").setBackground("#30251f").setFontColor("#ffffff");
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function safeToast_(ss, msg) {
  try { ss.toast(msg); } catch (err) { /* 從觸發器執行時沒有 UI */ }
}

// ===== API =====
function doGet() {
  return json_({ ok: true, service: "ginjian-orders" });
}

function doPost(e) {
  try {
    const body = JSON.parse((e && e.postData && e.postData.contents) || "{}");
    if (body.website) return json_({ ok: false, error: "invalid" }); // 防機器人欄位
    const order = validateOrder_(body);

    const lock = LockService.getScriptLock();
    lock.waitLock(20000);
    let saved;
    try {
      saved = saveOrder_(order);
    } finally {
      lock.releaseLock();
    }

    notify_(saved);
    return json_({ ok: true, orderId: saved.orderId, subtotal: saved.subtotal, fee: saved.fee, total: saved.total });
  } catch (err) {
    const message = err && err.userMessage ? err.userMessage : "系統忙碌中，請稍後再試";
    if (!(err && err.userMessage)) console.error(err);
    return json_({ ok: false, error: message });
  }
}

// ===== 驗證與計價 =====
function userError_(message) {
  const err = new Error(message);
  err.userMessage = message;
  return err;
}

function text_(value, max) {
  return String(value == null ? "" : value).replace(/[\u0000-\u001f]/g, " ").trim().slice(0, max);
}

function validateOrder_(body) {
  const mode = text_(body.fulfillment, 20);
  if (!MODE_NAMES[mode]) throw userError_("取餐方式不正確");

  const customer = body.customer || {};
  const name = text_(customer.name, 60);
  const phone = text_(customer.phone, 20);
  const email = text_(customer.email, 120);
  const note = text_(customer.note, 300);
  if (!name) throw userError_("請填寫姓名");
  if (!/^[0-9+\-\s()#]{8,20}$/.test(phone)) throw userError_("電話格式不正確");
  if (!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email)) throw userError_("Email 格式不正確");

  const productMap = loadProducts_();
  const rawItems = Array.isArray(body.items) ? body.items : [];
  if (!rawItems.length) throw userError_("購物車是空的");
  if (rawItems.length > CONFIG.MAX_LINE_ITEMS) throw userError_("品項過多");

  const merged = {};
  rawItems.forEach(item => {
    const id = Number(item && item.id);
    const qty = Number(item && item.qty);
    const product = productMap[id];
    if (!product || !product.active) throw userError_("有商品已下架，請重新整理頁面");
    if (!Number.isInteger(qty) || qty < 1 || qty > CONFIG.MAX_QTY) throw userError_("商品數量不正確");
    merged[id] = Math.min(CONFIG.MAX_QTY, (merged[id] || 0) + qty);
  });

  const items = Object.keys(merged).map(id => {
    const p = productMap[id];
    return { id: p.id, name: p.name, price: p.price, qty: merged[id], amount: p.price * merged[id], frozen: p.frozen };
  });

  let store = "";
  let address = "";
  if (mode === "delivery") {
    if (items.some(i => !i.frozen)) throw userError_("宅配僅支援冷凍禮盒商品");
    address = text_(body.deliveryAddress, 200);
    if (address.length < 6) throw userError_("請填寫完整宅配地址");
  } else {
    const stores = loadStores_();
    store = stores[text_(body.storeId, 40)];
    if (!store) throw userError_("請選擇門市");
  }

  const subtotal = items.reduce((sum, i) => sum + i.amount, 0);
  const fee = mode === "delivery" ? CONFIG.DELIVERY_FEE : 0;
  return { mode, store, address, name, phone, email, note, items, subtotal, fee, total: subtotal + fee };
}

function loadProducts_() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEET_PRODUCTS);
  if (!sheet || sheet.getLastRow() < 2) throw new Error("找不到商品工作表，請先執行 setup()");
  const rows = sheet.getRange(2, 1, sheet.getLastRow() - 1, PRODUCT_HEADERS.length).getValues();
  const map = {};
  rows.forEach(r => {
    const id = Number(r[0]);
    if (!id) return;
    map[id] = { id, name: String(r[1]), price: Number(r[2]) || 0, frozen: r[3] === true, active: r[4] !== false };
  });
  return map;
}

function loadStores_() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEET_STORES);
  const map = {};
  if (!sheet || sheet.getLastRow() < 2) return map;
  sheet.getRange(2, 1, sheet.getLastRow() - 1, 2).getValues().forEach(r => {
    if (r[0]) map[String(r[0])] = String(r[1]);
  });
  return map;
}

// ===== 寫入試算表 =====
// 避免顧客輸入 =、+、-、@ 開頭的文字被試算表當成公式執行
function safeCell_(value) {
  const s = String(value == null ? "" : value);
  return /^[=+\-@]/.test(s) ? "'" + s : s;
}

function nextOrderId_(now) {
  const day = Utilities.formatDate(now, CONFIG.TIMEZONE, "yyyyMMdd");
  const props = PropertiesService.getScriptProperties();
  const key = "SEQ_" + day;
  const seq = Number(props.getProperty(key) || 0) + 1;
  props.setProperty(key, String(seq));
  return "GJ" + day + "-" + ("000" + seq).slice(-3);
}

function saveOrder_(order) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const now = new Date();
  const orderId = nextOrderId_(now);
  const summary = order.items.map(i => i.name + " x" + i.qty).join("、");

  ss.getSheetByName(CONFIG.SHEET_ORDERS).appendRow([
    orderId, now, CONFIG.STATUSES[0], MODE_NAMES[order.mode], order.store, safeCell_(order.address),
    safeCell_(order.name), "'" + order.phone, safeCell_(order.email), safeCell_(order.note), summary,
    order.subtotal, order.fee, order.total
  ]);

  const itemRows = order.items.map(i => [orderId, now, i.id, i.name, i.price, i.qty, i.amount]);
  const items = ss.getSheetByName(CONFIG.SHEET_ITEMS);
  items.getRange(items.getLastRow() + 1, 1, itemRows.length, ITEM_HEADERS.length).setValues(itemRows);

  return Object.assign({ orderId, createdAt: now, summary }, order);
}

// ===== Email 通知 =====
function notify_(o) {
  const lines = o.items.map(i => `・${i.name} × ${i.qty}　NT$ ${i.amount}`).join("\n");
  const where = o.mode === "delivery" ? `宅配地址：${o.address}` : `門市：${o.store}`;
  const detail = `訂單編號：${o.orderId}
取餐方式：${MODE_NAMES[o.mode]}
${where}

${lines}

商品金額：NT$ ${o.subtotal}
運費：NT$ ${o.fee}
合計：NT$ ${o.total}`;

  try {
    const shop = prop_("NOTIFY_EMAIL");
    if (shop) {
      MailApp.sendEmail({
        to: shop,
        subject: `【金煎新訂單】${o.orderId}｜${o.name}｜NT$ ${o.total}`,
        body: `${detail}\n\n顧客：${o.name}\n電話：${o.phone}\nEmail：${o.email}\n備註：${o.note || "（無）"}`
      });
    }
    if (prop_("SEND_CUSTOMER_EMAIL") === "true") {
      MailApp.sendEmail({
        to: o.email,
        subject: `金煎 GINJIAN 訂單確認 ${o.orderId}`,
        body: `${o.name} 您好，\n\n感謝您的訂購！我們已收到您的訂單：\n\n${detail}\n\n金煎 GINJIAN 敬上`
      });
    }
  } catch (err) {
    console.error("寄信失敗", err); // 寄信失敗不影響訂單成立
  }
}

function json_(data) {
  return ContentService.createTextOutput(JSON.stringify(data)).setMimeType(ContentService.MimeType.JSON);
}
