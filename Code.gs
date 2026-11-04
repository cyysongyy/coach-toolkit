// 校長 Coach 工具箱｜雲端後端（Google Apps Script）
// 功能：1) 教案庫雲端備份（試算表 lessons 分頁）  2) 讀取 Drive 教案資料夾（Google 文件、純文字）
// 用法：開啟一份 Google 試算表 → 擴充功能 → Apps Script → 貼上本檔 → 部署為「網頁應用程式」
//       （執行身分：我；存取：所有人）。首次部署與每次改程式後，都要「管理部署作業 → 編輯 → 新版本」。
// 注意：此網址沒有密碼，知道網址的人可以讀寫教案庫並讀取下方資料夾內的文件，請勿公開網址。
const SHEET = 'lessons';
const FOLDER_ID = '1QIaM3Clcp-tBEwvhzhFo78jtBMvXosAZ'; // 教案資料夾

function sheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let s = ss.getSheetByName(SHEET);
  if (!s) { s = ss.insertSheet(SHEET); s.appendRow(['id', '日期', '時間', '名稱', 'json']); }
  return s;
}
function out_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}
// 檔案是否在教案資料夾（含子資料夾）內
function inTree_(file) {
  let level = [];
  const p0 = file.getParents();
  while (p0.hasNext()) level.push(p0.next());
  for (let d = 0; d < 4 && level.length; d++) {
    const next = [];
    for (let i = 0; i < level.length; i++) {
      if (level[i].getId() === FOLDER_ID) return true;
      const p = level[i].getParents();
      while (p.hasNext()) next.push(p.next());
    }
    level = next;
  }
  return false;
}
function collect_(folder, depth, acc) {
  const files = folder.getFiles();
  while (files.hasNext()) {
    const f = files.next();
    const m = f.getMimeType();
    if (m === MimeType.GOOGLE_DOCS || m === MimeType.PLAIN_TEXT) {
      acc.push({ id: f.getId(), name: f.getName(), updated: f.getLastUpdated().toISOString(), mime: m });
    }
  }
  if (depth > 0) {
    const subs = folder.getFolders();
    while (subs.hasNext()) collect_(subs.next(), depth - 1, acc);
  }
}
function doGet(e) {
  const a = e.parameter && e.parameter.action;
  if (a === 'files') {
    const acc = [];
    collect_(DriveApp.getFolderById(FOLDER_ID), 1, acc);
    acc.sort(function (x, y) { return y.updated.localeCompare(x.updated); });
    return out_({ ok: true, items: acc.slice(0, 100) });
  }
  if (a === 'file') {
    const f = DriveApp.getFileById(e.parameter.id);
    if (!inTree_(f)) return out_({ ok: false, error: '不在教案資料夾內' });
    const m = f.getMimeType();
    let text = '';
    if (m === MimeType.GOOGLE_DOCS) text = DocumentApp.openById(f.getId()).getBody().getText();
    else if (m === MimeType.PLAIN_TEXT) text = f.getBlob().getDataAsString('UTF-8');
    else return out_({ ok: false, error: '只支援 Google 文件與純文字' });
    return out_({ ok: true, name: f.getName(), text: text.slice(0, 40000) });
  }
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
