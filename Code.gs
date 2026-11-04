// 校長 Coach 工具箱｜教案雲端備份（Google Apps Script）
// 用法：開啟一份 Google 試算表 → 擴充功能 → Apps Script → 貼上本檔
// → 部署 → 新增部署作業 → 類型：網頁應用程式 → 執行身分：我 → 存取：所有人 → 複製網址貼到工具箱的雲端同步設定
const SHEET = 'lessons';

function sheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let s = ss.getSheetByName(SHEET);
  if (!s) { s = ss.insertSheet(SHEET); s.appendRow(['id', '日期', '時間', '名稱', 'json']); }
  return s;
}
function out_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}
function doGet(e) {
  const rows = sheet_().getDataRange().getValues().slice(1);
  const items = rows.map(function (r) { try { return JSON.parse(r.filter(function (c) { return c !== ''; }).pop()); } catch (x) { return null; } })
                    .filter(function (x) { return x; });
  return out_({ ok: true, items: items });
}
function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const b = JSON.parse(e.postData.contents);
    const s = sheet_();
    const ids = s.getRange(1, 1, Math.max(s.getLastRow(), 1), 1).getValues().map(function (r) { return String(r[0]); });
    if (b.action === 'save') {
      const en = b.entry;
      let js = JSON.stringify(en);
      if (js.length > 45000) { en.form.text = (en.form.text || '').slice(0, 3000); js = JSON.stringify(en); }
      if (js.length > 49000) return out_({ ok: false, error: '內容太大' });
      const dt = new Date(en.date);
      const row = [String(en.id), Utilities.formatDate(dt, 'Asia/Taipei', 'yyyy-MM-dd'), Utilities.formatDate(dt, 'Asia/Taipei', 'HH:mm'), en.name, js];
      const i = ids.indexOf(String(en.id));
      if (i > 0) s.getRange(i + 1, 1, 1, 5).setValues([row]); else s.appendRow(row);
      return out_({ ok: true });
    }
    if (b.action === 'delete') {
      const i = ids.indexOf(String(b.id));
      if (i > 0) s.deleteRow(i + 1);
      return out_({ ok: true });
    }
    return out_({ ok: false, error: '未知操作' });
  } finally { lock.releaseLock(); }
}
