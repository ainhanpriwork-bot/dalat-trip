// Apps Script cho trang "Đà Lạt tháng 12".
// Dán toàn bộ file này vào Tiện ích mở rộng > Apps Script của Google Sheet,
// rồi Triển khai > Lần triển khai mới > Ứng dụng web
// (Thực thi với tư cách: Tôi, Ai có quyền truy cập: Bất kỳ ai).

const SHEET_ID = '1v6N65ipkcueUdUQYw3Z9GAnngBmC5q6IzVZlj8Y_d6o';

const TABS = {
  items:    { name: 'LichTrinh', cols: ['id','day','time','title','place','note','link','done','suggestion','createdAt'] },
  expenses: { name: 'ChiPhi',    cols: ['id','day','title','amount','category','payer','status','note','createdAt'] },
  memories: { name: 'KyNiem',    cols: ['id','day','title','text','createdAt','photo'] },
  fund:     { name: 'QuyChung',  cols: ['id','day','payer','amount','note','createdAt'] },
};
const PHOTO_FOLDER = 'Đà Lạt tháng 12 – ảnh';

function photoFolder_() {
  const it = DriveApp.getFoldersByName(PHOTO_FOLDER);
  return it.hasNext() ? it.next() : DriveApp.createFolder(PHOTO_FOLDER);
}
function upload_(d) {
  const mime = /^image\/(jpeg|png|webp|gif)$/.test(d.mime) ? d.mime : 'image/jpeg';
  const bytes = Utilities.base64Decode(d.b64);
  if (bytes.length > 10 * 1024 * 1024) throw new Error('too large');
  const blob = Utilities.newBlob(bytes, mime, 'kyniem-' + Date.now() + (mime === 'image/png' ? '.png' : '.jpg'));
  const file = photoFolder_().createFile(blob);
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  return file.getId();
}
const BOOL = ['done','suggestion'];
const NUM = ['amount','createdAt'];
const TZ = 'Asia/Ho_Chi_Minh';

function sheet_(name) { return SpreadsheetApp.openById(SHEET_ID).getSheetByName(name); }

function cellOut_(col, v) {
  if (v instanceof Date) return Utilities.formatDate(v, TZ, col === 'time' ? 'HH:mm' : 'yyyy-MM-dd');
  if (BOOL.indexOf(col) >= 0) return v === true || String(v).toUpperCase() === 'TRUE';
  if (NUM.indexOf(col) >= 0) { const n = Number(String(v).replace(/[^\d.-]/g, '')); return isNaN(n) ? 0 : n; }
  if (col === 'payer') return v === '' || v === null ? null : (String(v) === 'fund' ? 'fund' : Number(v));
  return v === null || v === undefined ? '' : String(v);
}
function cellIn_(col, v) {
  if (v === null || v === undefined) return '';
  if (typeof v === 'boolean') return v ? 'TRUE' : 'FALSE';
  return String(v);
}

function readTab_(key) {
  const t = TABS[key], sh = sheet_(t.name);
  const last = sh.getLastRow();
  if (last < 2) return [];
  const rows = sh.getRange(2, 1, last - 1, t.cols.length).getValues();
  return rows.filter(r => String(r[0]).trim() !== '').map(r => {
    const o = {};
    t.cols.forEach((c, i) => { o[c] = cellOut_(c, r[i]); });
    return o;
  });
}
function readMembers_() {
  const sh = sheet_('ThanhVien');
  const names = ['Người 1','Người 2','Người 3','Người 4'];
  const last = sh.getLastRow();
  if (last >= 2) sh.getRange(2, 1, last - 1, 2).getValues().forEach(r => {
    const i = Number(r[0]); if (i >= 0 && i < 4 && String(r[1]).trim()) names[i] = String(r[1]).trim();
  });
  return names;
}
function readAll_() {
  return { items: readTab_('items'), expenses: readTab_('expenses'), memories: readTab_('memories'), fund: readTab_('fund'), members: readMembers_() };
}
function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
function findRow_(sh, id) {
  const last = sh.getLastRow();
  if (last < 2) return -1;
  const ids = sh.getRange(2, 1, last - 1, 1).getValues();
  for (let i = 0; i < ids.length; i++) if (String(ids[i][0]) === String(id)) return i + 2;
  return -1;
}

function doGet() { return json_({ ok: true, data: readAll_() }); }

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const req = JSON.parse(e.postData.contents);
    const a = req.action;
    if (a === 'upload') return json_({ ok: true, id: upload_(req.data || {}) });
    if (a === 'setMembers') {
      const sh = sheet_('ThanhVien');
      const vals = (req.data || []).slice(0, 4).map((n, i) => [String(i), String(n || ('Người ' + (i + 1))).slice(0, 40)]);
      sh.getRange(2, 1, vals.length, 2).setNumberFormat('@').setValues(vals);
      return json_({ ok: true, data: readAll_() });
    }
    const t = TABS[req.col];
    if (!t) return json_({ ok: false, error: 'unknown collection' });
    const sh = sheet_(t.name);
    if (a === 'add') {
      const d = req.data || {};
      d.id = d.id || Utilities.getUuid().slice(0, 12);
      const row = t.cols.map(c => cellIn_(c, d[c]));
      const r = sh.getLastRow() + 1;
      sh.getRange(r, 1, 1, row.length).setNumberFormat('@').setValues([row]);
    } else if (a === 'update') {
      const r = findRow_(sh, req.id);
      if (r < 0) return json_({ ok: false, error: 'not found' });
      const cur = sh.getRange(r, 1, 1, t.cols.length).getValues()[0];
      const d = req.data || {};
      const row = t.cols.map((c, i) => c === 'id' ? cur[i] : (c in d ? cellIn_(c, d[c]) : cur[i]));
      sh.getRange(r, 1, 1, row.length).setNumberFormat('@').setValues([row]);
    } else if (a === 'remove') {
      const r = findRow_(sh, req.id);
      if (r > 0) sh.deleteRow(r);
    } else {
      return json_({ ok: false, error: 'unknown action' });
    }
    return json_({ ok: true, data: readAll_() });
  } catch (err) {
    return json_({ ok: false, error: String(err) });
  } finally {
    lock.releaseLock();
  }
}
