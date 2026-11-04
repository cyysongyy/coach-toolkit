// 校長 Coach 工具箱｜教案雲端備份（Google Apps Script）
// 用法：開啟一份 Google 試算表 → 擴充功能 → Apps Script → 貼上本檔 → 把 TOKEN 改成自訂密碼
// → 部署 → 新增部署作業 → 類型：網頁應用程式 → 執行身分：我 → 存取：所有人 → 複製網址貼到工具箱的雲端同步設定
const TOKEN = '請改成你自己的密碼';
const SHEET = 'lessons';

function sheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let s = ss.getSheetByName(SHEET);
  if (!s) { s = ss.insertSheet(SHEET); s.appendRow(['id', 'name', 'date', 'json']); }
  return s;
}
function out_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}
function doGet(e) {
  if (!e.parameter || e.parameter.token !== TOKEN) return out_({ ok: false, error: '密碼錯誤' });
  const rows = sheet_().getDataRange().getValues().slice(1);
  const items = rows.map(function (r) { try { return JSON.parse(r[3]); } catch (x) { return null; } })
                    .filter(function (x) { return x; });
  return out_({ ok: true, items: items });
}
function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const b = JSON.parse(e.postData.contents);
    if (b.token !== TOKEN) return out_({ ok: false, error: '密碼錯誤' });
    const s = sheet_();
    const ids = s.getRange(1, 1, Math.max(s.getLastRow(), 1), 1).getValues().map(function (r) { return String(r[0]); });
    if (b.action === 'save') {
      const en = b.entry;
      let js = JSON.stringify(en);
      if (js.length > 45000) { en.form.text = (en.form.text || '').slice(0, 3000); js = JSON.stringify(en); }
      if (js.length > 49000) return out_({ ok: false, error: '內容太大' });
      const row = [String(en.id), en.name, en.date, js];
      const i = ids.indexOf(String(en.id));
      if (i > 0) s.getRange(i + 1, 1, 1, 4).setValues([row]); else s.appendRow(row);
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
