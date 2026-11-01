/**
 * YUMU 報名表 → Google Sheet
 * 貼到「YUMU 客戶檔」試算表的 擴充功能 > Apps Script
 */
const SHEET_ID   = '1IE35o-_whUkiXT95wHncVQTVc9le3gCnl0fXgXIDTss';
const SHEET_NAME = '工作表1';
// 新報名通知信收件人（可填多個，用逗號分隔）
const NOTIFY_TO = '858377355tw@gmail.com';
const HEADERS = ['時間','姓名','公司／單位','職稱','Email','電話','想了解的項目','問題或需求','個資同意','來源頁面','處理狀態','備註'];

function doPost(e) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(20000);
    const p = (e && e.parameter) || {};

    // 必填檢查
    if (!clean_(p.name) || !clean_(p.org) || !clean_(p.email) || p.agree !== '是') {
      return json_({ ok: false, error: 'missing' });
    }
    if (!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(p.email)) {
      return json_({ ok: false, error: 'email' });
    }

    const ss = SpreadsheetApp.openById(SHEET_ID);
    const sh = ss.getSheetByName(SHEET_NAME) || ss.insertSheet(SHEET_NAME);
    if (sh.getLastRow() === 0) {
      sh.appendRow(HEADERS);
      sh.getRange(1, 1, 1, HEADERS.length).setFontWeight('bold').setBackground('#EEF2EC');
      sh.setFrozenRows(1);
    }

    sh.appendRow([
      Utilities.formatDate(new Date(), 'Asia/Taipei', 'yyyy-MM-dd HH:mm:ss'),
      clean_(p.name), clean_(p.org), clean_(p.title), clean_(p.email), clean_(p.phone),
      clean_(p.interest), clean_(p.message), '是', clean_(p.source), '待處理', ''
    ]);
    try { notify_(p); } catch (mailErr) { /* 寄信失敗不影響資料寫入 */ }
    return json_({ ok: true });
  } catch (err) {
    return json_({ ok: false, error: String(err) });
  } finally {
    try { lock.releaseLock(); } catch (x) {}
  }
}

// 新報名通知信（純文字）
function notify_(p) {
  const oneLine = function (v) { return clean_(v).replace(/[\r\n]+/g, ' '); };
  const body = [
    '收到一筆新的報名／洽詢：', '',
    '姓名：' + oneLine(p.name),
    '公司／單位：' + oneLine(p.org),
    '職稱：' + oneLine(p.title),
    'Email：' + oneLine(p.email),
    '電話：' + oneLine(p.phone),
    '想了解的項目：' + oneLine(p.interest), '',
    '問題或需求：', clean_(p.message), '',
    '試算表：https://docs.google.com/spreadsheets/d/' + SHEET_ID + '/edit'
  ].join('\n');
  MailApp.sendEmail({
    to: NOTIFY_TO,
    subject: '【YUMU 新報名】' + oneLine(p.name) + '｜' + oneLine(p.org),
    body: body,
    replyTo: oneLine(p.email),
    name: 'YUMU 報名表'
  });
}

// 開網址測試用：看到 {"ok":true,...} 代表部署成功
function doGet() { return json_({ ok: true, service: 'YUMU signup' }); }

// 防公式注入 + 限長；電話等文字一律存成純文字
function clean_(v) {
  let s = String(v == null ? '' : v).trim().slice(0, 2000);
  if (/^[=+\-@]/.test(s)) s = "'" + s;
  return s;
}
function json_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}

// 可選：在編輯器手動執行一次，先確認試算表有權限、表頭會建立
function testWrite() {
  doPost({ parameter: { name: '測試', org: '測試公司', email: 'test@example.com', agree: '是', message: '這是一筆測試，可刪除' } });
}
