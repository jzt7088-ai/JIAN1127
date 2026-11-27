/**
 * 煎記 JIAN JI — 訂單後端 (Google Apps Script)
 * 網站結帳 → POST JSON → 這支程式 → 寫入 Google 試算表（＋寄通知信）
 *
 * 使用步驟：
 *   1. 把本檔全部貼到 Apps Script 的 Code.gs
 *   2. 上方函式選單選「setup」→ 執行 → 授權（會建立/連結試算表並建好工作表）
 *   3. 部署 → 新增部署作業 → 網頁應用程式：執行身分「我」、存取權「所有人」
 *   4. 把部署得到的 /exec 網址貼給網站（index.html 的 ORDER_API）
 */

// ===== 設定 =====
const CONFIG = {
  SPREADSHEET_ID: '',          // 留空：setup() 會自動建立新試算表；也可填入既有試算表 ID
  NOTIFY_EMAIL: '',            // 填店家 Email 就會在每張新訂單寄通知信；留空不寄
  SEND_CUSTOMER_EMAIL: true,   // 顧客有填 Email 時寄訂單確認信
  FREE_SHIPPING: 1500,         // 宅配滿額免運
  SHIPPING_FEE: 160,
  TIMEZONE: 'Asia/Taipei'
};

// 價目表（以伺服器端為準，避免前端竄改金額）— 新增品項時請與網站同步
const CATALOG = {
  p1: ['招牌豬肉鍋貼（10 顆）', 90], p2: ['韭黃鮮蝦鍋貼（10 顆）', 130], p3: ['麻辣鍋貼（10 顆）', 100],
  p4: ['玉米豬肉鍋貼（10 顆）', 95], d1: ['高麗菜水餃（10 顆）', 75], d2: ['紅油抄手（8 顆）', 85],
  d3: ['香菇雞肉水餃（10 顆）', 85], s1: ['酸辣湯', 45], s2: ['玉米濃湯', 40], s3: ['古早味紅茶', 30],
  c1: ['涼拌小黃瓜', 40], c2: ['滷味拼盤', 60],
  g1: ['招牌招財鍋貼禮盒（50 入）', 680], g2: ['家庭團圓水餃組（100 入）', 990]
};

const ORDER_HEADERS = ['時間', '訂單編號', '取餐方式', '姓名', '手機', 'Email', '地址', '備註', '品項', '小計', '運費', '總計', '狀態'];
const ITEM_HEADERS = ['訂單編號', '時間', '品項代碼', '品名', '單價', '數量', '金額'];
const MODES = ['內用', '外帶', '宅配'];

// ===== 初始化：建立/連結試算表與工作表 =====
function setup() {
  const ss = getSpreadsheet_();
  ensureSheet_(ss, '訂單', ORDER_HEADERS);
  ensureSheet_(ss, '訂單明細', ITEM_HEADERS);
  Logger.log('試算表網址：' + ss.getUrl());
}

// ===== API =====
function doGet() {
  return json_({ ok: true, service: 'jianji-orders', time: new Date().toISOString() });
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(20000);
    const data = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    if (data.website) return json_({ ok: true });               // 防機器人陷阱欄位

    // --- 驗證 ---
    const name = clean_(data.name, 40), phone = clean_(data.phone, 20).replace(/\D/g, '');
    const mode = MODES.indexOf(data.mode) >= 0 ? data.mode : '';
    const address = clean_(data.address, 200), note = clean_(data.note, 300), email = clean_(data.email, 100);
    if (!name) return json_({ ok: false, error: '請填寫姓名' });
    if (!/^09\d{8}$/.test(phone)) return json_({ ok: false, error: '手機格式不正確' });
    if (!mode) return json_({ ok: false, error: '請選擇取餐方式' });
    if (mode === '宅配' && !address) return json_({ ok: false, error: '宅配請填寫地址' });

    // --- 以伺服器價目表重新計價 ---
    const lines = [];
    (Array.isArray(data.items) ? data.items : []).forEach(it => {
      const p = CATALOG[it && it.id], qty = Math.floor(Number(it && it.qty));
      if (p && qty > 0 && qty <= 99) lines.push({ id: it.id, name: p[0], price: p[1], qty: qty, amount: p[1] * qty });
    });
    if (!lines.length) return json_({ ok: false, error: '購物車是空的' });
    const subtotal = lines.reduce((s, l) => s + l.amount, 0);
    const shipping = mode === '宅配' && subtotal < CONFIG.FREE_SHIPPING ? CONFIG.SHIPPING_FEE : 0;
    const total = subtotal + shipping;

    // --- 寫入 ---
    const ss = getSpreadsheet_();
    const orders = ensureSheet_(ss, '訂單', ORDER_HEADERS);
    const items = ensureSheet_(ss, '訂單明細', ITEM_HEADERS);
    const now = new Date();
    const orderNo = nextOrderNo_(now);
    const summary = lines.map(l => l.name + ' ×' + l.qty).join('、');
    orders.appendRow([now, orderNo, mode, name, "'" + phone, email, address, note, summary, subtotal, shipping, total, '新訂單']);
    items.getRange(items.getLastRow() + 1, 1, lines.length, ITEM_HEADERS.length)
      .setValues(lines.map(l => [orderNo, now, l.id, l.name, l.price, l.qty, l.amount]));

    if (CONFIG.NOTIFY_EMAIL) {
      MailApp.sendEmail(CONFIG.NOTIFY_EMAIL, '【煎記】新訂單 ' + orderNo + '（' + mode + '）$' + total,
        ['訂單編號：' + orderNo, '取餐方式：' + mode, '姓名：' + name, '手機：' + phone, address ? '地址：' + address : '',
         '品項：' + summary, '小計：$' + subtotal, '運費：$' + shipping, '總計：$' + total, note ? '備註：' + note : '',
         '', '試算表：' + ss.getUrl()].filter(String).join('\n'));
    }
    if (CONFIG.SEND_CUSTOMER_EMAIL && /^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(email)) {
      MailApp.sendEmail(email, '【煎記】訂單確認 ' + orderNo,
        [name + ' 您好，感謝您的訂購！', '', '訂單編號：' + orderNo, '取餐方式：' + mode, '品項：' + summary,
         '總計：$' + total, '', mode === '宅配' ? '我們將於 2 個工作天內出貨。' : '餐點約 15–20 分鐘完成。', '', '煎記 JIAN JI'].join('\n'));
    }
    return json_({ ok: true, orderNo: orderNo, subtotal: subtotal, shipping: shipping, total: total });
  } catch (err) {
    return json_({ ok: false, error: '伺服器忙碌，請稍後再試' , detail: String(err) });
  } finally {
    try { lock.releaseLock(); } catch (e) {}
  }
}

// ===== 工具 =====
function getSpreadsheet_() {
  const props = PropertiesService.getScriptProperties();
  const id = CONFIG.SPREADSHEET_ID || props.getProperty('SPREADSHEET_ID');
  if (id) return SpreadsheetApp.openById(id);
  const ss = SpreadsheetApp.create('煎記 JIAN JI 訂單');
  props.setProperty('SPREADSHEET_ID', ss.getId());
  return ss;
}

function ensureSheet_(ss, name, headers) {
  let sh = ss.getSheetByName(name);
  if (!sh) {
    const blank = ss.getSheets().length === 1 && ss.getSheets()[0].getLastRow() === 0 ? ss.getSheets()[0] : null;
    sh = blank ? blank.setName(name) : ss.insertSheet(name);
  }
  if (sh.getLastRow() === 0) {
    sh.appendRow(headers);
    sh.getRange(1, 1, 1, headers.length).setFontWeight('bold').setBackground('#4a2e1f').setFontColor('#f7f1e6');
    sh.setFrozenRows(1);
    if (name === '訂單') {
      sh.getRange('A:A').setNumberFormat('yyyy/mm/dd hh:mm');
      sh.getRange('J:L').setNumberFormat('$#,##0');
      sh.getRange('M2:M').setDataValidation(SpreadsheetApp.newDataValidation()
        .requireValueInList(['新訂單', '製作中', '可取餐', '已出貨', '已完成', '已取消'], true).build());
    } else {
      sh.getRange('B:B').setNumberFormat('yyyy/mm/dd hh:mm');
      sh.getRange('E:E').setNumberFormat('$#,##0'); sh.getRange('G:G').setNumberFormat('$#,##0');
    }
  }
  return sh;
}

function nextOrderNo_(now) {
  const props = PropertiesService.getScriptProperties();
  const day = Utilities.formatDate(now, CONFIG.TIMEZONE, 'yyMMdd');
  const seq = props.getProperty('SEQ_DATE') === day ? Number(props.getProperty('SEQ') || 0) + 1 : 1;
  props.setProperties({ SEQ_DATE: day, SEQ: String(seq) });
  return 'JJ' + day + ('000' + seq).slice(-3);
}

function clean_(v, max) {
  return String(v == null ? '' : v).replace(/[\u0000-\u001f]/g, ' ').replace(/^[=+\-@]/, "'$&").trim().slice(0, max);
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
