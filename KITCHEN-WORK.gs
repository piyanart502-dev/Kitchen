/**
 * KITCHEN WORK — หลังบ้าน (Google Apps Script + Google Sheets)
 *
 * ไฟล์นี้รับข้อมูลจากหน้าแอปบน GitHub Pages แล้วบันทึกลง Google Sheets
 * ส่วนรูปบ่อดักไขมันเก็บไว้ใน Google Drive โฟลเดอร์ "KITCHEN WORK - รูปบ่อดักไขมัน"
 *
 * ติดตั้ง (ทำครั้งเดียว)
 *   1) เปิด Google Sheets ไฟล์ใหม่ > ส่วนขยาย > Apps Script
 *   2) วางโค้ดนี้ทับของเดิมทั้งหมดใน code.gs แล้วกด Ctrl+S
 *   3) เลือกฟังก์ชัน setup แล้วกด ▷ เรียกใช้ (ครั้งแรกจะขอสิทธิ์ ให้กด ขั้นสูง > ไปที่ KITCHEN WORK > อนุญาต)
 *   4) การทำให้ใช้งานได้ > การทำให้ใช้งานได้รายการใหม่ > ประเภท เว็บแอป
 *        เรียกใช้งานในฐานะ: ฉัน
 *        ผู้ที่มีสิทธิ์เข้าถึง: ทุกคน
 *   5) คัดลอก URL ที่ลงท้ายด้วย /exec ไปใส่ใน index.html บรรทัด const API_URL = '...';
 *
 * อัปเดตโค้ดภายหลัง: วางทับแล้วไปที่ การทำให้ใช้งานได้ > จัดการการทำให้ใช้งานได้ > ✏️ > เวอร์ชันใหม่ > ทำให้ใช้งานได้
 * (ลิงก์ /exec จะเหมือนเดิม)
 *
 * ตั้ง PIN: แก้ตัวเลขใน PIN ด้านล่าง > Ctrl+S > เลือกฟังก์ชัน setPin > ▷ เรียกใช้
 *          จากนั้นแก้ตัวเลขในโค้ดกลับเป็น 1234 ได้ เพราะระบบเก็บ PIN จริงไว้แยกแล้ว
 * ยกเลิก PIN: เลือกฟังก์ชัน clearPin > ▷ เรียกใช้
 *
 * ห้ามสลับ แทรก หรือลบคอลัมน์ในชีต เพราะระบบอ่านข้อมูลตามลำดับคอลัมน์
 */

var PIN = '1234';

var API_VERSION = 1;
var FOLDER_NAME = 'KITCHEN WORK - รูปบ่อดักไขมัน';
var MAX_ROWS = 1000; // จำนวนแถวล่าสุดที่ส่งให้แอปต่อชีต

// ชนิดคอลัมน์: s = ข้อความ, n = ตัวเลข, b = ใช่/ไม่ใช่, a = รายการคั่นด้วยจุลภาค, j = JSON
var SCHEMA = {
  temps: {sheet: 'Temps', cols: [
    ['id', 'รหัส', 's'], ['date', 'วันที่', 's'], ['time', 'เวลา', 's'], ['slot', 'รอบ', 's'],
    ['cab', 'รหัสตู้', 's'], ['cabName', 'ตู้', 's'], ['val', 'อุณหภูมิ °C', 'n'],
    ['min', 'เกณฑ์ต่ำสุด', 'n'], ['max', 'เกณฑ์สูงสุด', 'n'], ['who', 'ผู้บันทึก', 's'],
    ['note', 'หมายเหตุ', 's'], ['ts', 'ลำดับเวลา', 's'], ['createdAt', 'บันทึกเมื่อ', 's'],
    ['_extra', 'ข้อมูลอื่น (JSON)', 'j']]},
  grease_weeks: {sheet: 'GreaseWeeks', cols: [
    ['id', 'สัปดาห์ (วันอาทิตย์)', 's'], ['week', 'สัปดาห์', 's'], ['cleanDue', 'กำหนดล้าง', 's'],
    ['cleanBy', 'ผู้ทำความสะอาด', 's'], ['cleanAt', 'วันที่ล้าง', 's'], ['cleanTime', 'เวลาที่ล้าง', 's'],
    ['cleanSigned', 'ลงชื่อล้างเมื่อ', 's'], ['hasPhoto', 'มีรูป', 'b'], ['photoUrl', 'ลิงก์รูป', 's'],
    ['photoId', 'รหัสไฟล์รูป', 's'], ['inspDue', 'กำหนดตรวจ', 's'], ['inspBy', 'ผู้ตรวจ', 's'],
    ['inspPos', 'ตำแหน่งผู้ตรวจ', 's'], ['inspAt', 'วันที่ตรวจ', 's'], ['inspResult', 'ผลตรวจ', 's'],
    ['inspNote', 'สิ่งที่ต้องแก้ไข', 's'], ['inspSigned', 'ลงชื่อตรวจเมื่อ', 's'], ['ts', 'ลำดับเวลา', 's'],
    ['_extra', 'ข้อมูลอื่น (JSON)', 'j']]},
  stock_items: {sheet: 'StockItems', cols: [
    ['id', 'รหัส', 's'], ['name', 'เครื่องปรุง', 's'], ['unit', 'หน่วย', 's'], ['bal', 'คงเหลือ', 'n'],
    ['min', 'ขั้นต่ำ', 'n'], ['codes', 'บาร์โค้ด', 'a'], ['order', 'ลำดับ', 'n'], ['ts', 'ลำดับเรียง', 's'],
    ['_extra', 'ข้อมูลอื่น (JSON)', 'j']]},
  stock_moves: {sheet: 'StockMoves', cols: [
    ['id', 'รหัส', 's'], ['date', 'วันที่', 's'], ['ts', 'ลำดับเวลา', 's'], ['type', 'ประเภท (in/out/adj)', 's'],
    ['itemId', 'รหัสเครื่องปรุง', 's'], ['item', 'เครื่องปรุง', 's'], ['unit', 'หน่วย', 's'], ['qty', 'จำนวน', 'n'],
    ['balAfter', 'คงเหลือหลังทำรายการ', 'n'], ['diff', 'ส่วนต่าง (ปรับยอด)', 'n'], ['prev', 'ยอดเดิม (ปรับยอด)', 'n'],
    ['ref', 'ใช้สำหรับ', 's'], ['who', 'ผู้ทำรายการ', 's'], ['note', 'หมายเหตุ', 's'], ['createdAt', 'บันทึกเมื่อ', 's'],
    ['_extra', 'ข้อมูลอื่น (JSON)', 'j']]},
  settings: {sheet: 'Settings', whole: true, cols: [['id', 'รหัส', 's'], ['_all', 'ค่าตั้ง (JSON)', 'j']]}
};

/* ===================== เรียกใช้เองจากหน้า Apps Script ===================== */

function setup() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) throw new Error('ให้เปิด Apps Script จากเมนู ส่วนขยาย > Apps Script ของไฟล์ Google Sheets');
  PropertiesService.getScriptProperties().setProperty('SHEET_ID', ss.getId());
  Object.keys(SCHEMA).forEach(function (col) {
    var s = SCHEMA[col];
    prepSheet_(ss.getSheetByName(s.sheet) || ss.insertSheet(s.sheet), s);
  });
  ss.getSheets().forEach(function (sh) {
    var ours = Object.keys(SCHEMA).some(function (c) { return SCHEMA[c].sheet === sh.getName(); });
    if (!ours && sh.getLastRow() === 0 && ss.getSheets().length > 1) ss.deleteSheet(sh);
  });
  folder_();
  Logger.log('ติดตั้งเสร็จแล้ว ไฟล์ข้อมูล: ' + ss.getUrl());
}

function setPin() {
  var pin = String(PIN || '').trim();
  if (!/^\d{4,6}$/.test(pin)) throw new Error('PIN ต้องเป็นตัวเลข 4–6 หลัก');
  PropertiesService.getScriptProperties().setProperty('PIN', pin);
  Logger.log('ตั้ง PIN แล้ว ทุกเครื่องต้องใส่ PIN นี้ครั้งแรกที่เปิดแอป');
}

function clearPin() {
  PropertiesService.getScriptProperties().deleteProperty('PIN');
  Logger.log('ยกเลิก PIN แล้ว');
}

/* ===================== รับคำขอจากหน้าแอป ===================== */

function doGet() {
  return json_({ok: true, app: 'KITCHEN WORK', version: API_VERSION, message: 'หลังบ้าน KITCHEN WORK พร้อมใช้งาน'});
}

function doPost(e) {
  var req;
  try { req = JSON.parse((e && e.postData && e.postData.contents) || '{}'); }
  catch (err) { return json_({ok: false, error: 'bad_request'}); }
  try {
    var pin = PropertiesService.getScriptProperties().getProperty('PIN');
    if (pin && String(req.pin || '') !== pin) return json_({ok: false, error: 'pin'});
    return json_(route_(req));
  } catch (err) {
    return json_({ok: false, error: 'server', message: String((err && err.message) || err)});
  }
}

function route_(req) {
  var op = req.op, col = req.col, id = req.id == null ? '' : String(req.id);
  if (op === 'ping') return {ok: true, version: API_VERSION};
  if (op === 'sync') return sync_();
  if (['get', 'set', 'update', 'delete'].indexOf(op) < 0) return {ok: false, error: 'bad_op'};
  if (!id) return {ok: false, error: 'bad_id'};
  if (col === 'grease_photos') return photo_(op, id, req.data || {});
  if (!SCHEMA[col]) return {ok: false, error: 'bad_col'};
  if (op === 'get') { var doc = findDoc_(col, id); return {ok: true, exists: !!doc, data: doc}; }
  return withLock_(function () {
    if (op === 'delete') deleteDoc_(col, id);
    else writeDoc_(col, id, req.data || {}, op === 'update');
    return {ok: true};
  });
}

function sync_() {
  var data = {};
  ['temps', 'grease_weeks', 'stock_items', 'stock_moves'].forEach(function (col) {
    var rows = readAll_(col);
    if (col !== 'stock_items') {
      rows.sort(function (a, b) { var x = a[1].ts || '', y = b[1].ts || ''; return x < y ? 1 : x > y ? -1 : 0; });
      rows = rows.slice(0, MAX_ROWS);
    }
    data[col] = rows;
  });
  return {ok: true, version: API_VERSION, sheetUrl: ss_().getUrl(), data: data, settings: findDoc_('settings', 'main')};
}

/* รูปบ่อดักไขมัน: เก็บไฟล์ใน Drive แล้วจดรหัสไฟล์ไว้ในชีต GreaseWeeks */
function photo_(op, week, data) {
  if (op === 'get') {
    var w = findDoc_('grease_weeks', week);
    if (!w || !w.photoId) return {ok: true, exists: false};
    try {
      var b = DriveApp.getFileById(w.photoId).getBlob();
      return {ok: true, exists: true, data: {week: week, img: 'data:' + b.getContentType() + ';base64,' + Utilities.base64Encode(b.getBytes())}};
    } catch (err) { return {ok: true, exists: false}; }
  }
  return withLock_(function () {
    var cur = findDoc_('grease_weeks', week);
    if (cur && cur.photoId) { try { DriveApp.getFileById(cur.photoId).setTrashed(true); } catch (err) {} }
    if (op === 'delete') {
      if (cur) writeDoc_('grease_weeks', week, {photoId: '', photoUrl: '', hasPhoto: false}, true);
      return {ok: true};
    }
    var m = String(data.img || '').match(/^data:([^;]+);base64,(.+)$/);
    if (!m) return {ok: false, error: 'bad_image', message: 'รูปไม่ถูกต้อง'};
    var blob = Utilities.newBlob(Utilities.base64Decode(m[2]), m[1], 'บ่อดักไขมัน_' + week + '.jpg');
    var file = folder_().createFile(blob);
    try { file.setDescription('ทำความสะอาดโดย ' + (data.by || '') + ' เมื่อ ' + (data.at || '')); } catch (err) {}
    var patch = {photoId: file.getId(), photoUrl: file.getUrl(), hasPhoto: true};
    if (cur) writeDoc_('grease_weeks', week, patch, true);
    else writeDoc_('grease_weeks', week, merge_({week: week, ts: week, cleanDue: week}, patch), false);
    return {ok: true};
  });
}

/* ===================== อ่าน/เขียนชีต ===================== */

function ss_() {
  var id = PropertiesService.getScriptProperties().getProperty('SHEET_ID');
  return id ? SpreadsheetApp.openById(id) : SpreadsheetApp.getActiveSpreadsheet();
}

function sheet_(col) {
  var s = SCHEMA[col], ss = ss_(), sh = ss.getSheetByName(s.sheet);
  if (!sh) { sh = ss.insertSheet(s.sheet); prepSheet_(sh, s); }
  return sh;
}

function fmt_(t) { return t === 'n' || t === 'b' ? 'General' : '@'; }

function prepSheet_(sh, s) {
  var n = s.cols.length;
  sh.getRange(1, 1, 1, n).setNumberFormat('@')
    .setValues([s.cols.map(function (c) { return c[1]; })])
    .setFontWeight('bold').setBackground('#E2EDF7');
  sh.setFrozenRows(1);
  var rows = Math.max(sh.getMaxRows() - 1, 1);
  s.cols.forEach(function (c, i) { sh.getRange(2, i + 1, rows, 1).setNumberFormat(fmt_(c[2])); });
}

function findRow_(sh, id) {
  var last = sh.getLastRow();
  if (last < 2) return -1;
  var ids = sh.getRange(2, 1, last - 1, 1).getDisplayValues();
  for (var i = 0; i < ids.length; i++) if (ids[i][0] === id) return i + 2;
  return -1;
}

function findDoc_(col, id) {
  var s = SCHEMA[col], sh = sheet_(col), r = findRow_(sh, id);
  return r < 0 ? null : fromRow_(s, sh.getRange(r, 1, 1, s.cols.length).getValues()[0]);
}

function readAll_(col) {
  var s = SCHEMA[col], sh = sheet_(col), last = sh.getLastRow(), out = [];
  if (last < 2) return out;
  sh.getRange(2, 1, last - 1, s.cols.length).getValues().forEach(function (row) {
    var id = cellText_(row[0], 'id');
    if (id) out.push([id, fromRow_(s, row)]);
  });
  return out;
}

function writeDoc_(col, id, data, isMerge) {
  var s = SCHEMA[col], sh = sheet_(col), n = s.cols.length, r = findRow_(sh, id), doc = data;
  if (isMerge && r > 0) doc = merge_(fromRow_(s, sh.getRange(r, 1, 1, n).getValues()[0]), data);
  if (r < 0) r = Math.max(sh.getLastRow(), 1) + 1;
  if (r > sh.getMaxRows()) sh.insertRowsAfter(sh.getMaxRows(), r - sh.getMaxRows() + 50);
  var rg = sh.getRange(r, 1, 1, n);
  rg.setNumberFormats([s.cols.map(function (c) { return fmt_(c[2]); })]);
  rg.setValues([toRow_(s, id, doc)]);
}

function deleteDoc_(col, id) {
  var s = SCHEMA[col], sh = sheet_(col), r = findRow_(sh, id);
  if (r < 0) return;
  if (sh.getMaxRows() <= 2) sh.getRange(r, 1, 1, s.cols.length).clearContent();
  else sh.deleteRow(r);
}

function toRow_(s, id, doc) {
  if (s.whole) return [id, JSON.stringify(doc || {})];
  var known = {}, extra = {}, hasExtra = false;
  s.cols.forEach(function (c) { known[c[0]] = true; });
  Object.keys(doc || {}).forEach(function (k) {
    if (!known[k] && k !== '_id' && doc[k] !== undefined) { extra[k] = doc[k]; hasExtra = true; }
  });
  return s.cols.map(function (c) {
    var k = c[0], t = c[2];
    var v = k === 'id' ? id : k === '_extra' ? (hasExtra ? extra : '') : doc[k];
    if (v === undefined || v === null || v === '') return '';
    if (t === 'n') { var num = Number(v); return isFinite(num) ? num : ''; }
    if (t === 'b') return !!v;
    if (t === 'a') return (Array.isArray(v) ? v : [v]).join(', ');
    if (t === 'j') return JSON.stringify(v);
    return String(v);
  });
}

function fromRow_(s, row) {
  if (s.whole) { try { return JSON.parse(row[1] || '{}'); } catch (err) { return {}; } }
  var doc = {};
  s.cols.forEach(function (c, i) {
    var k = c[0], t = c[2], v = row[i];
    if (k === 'id' || v === '' || v === null || v === undefined) return;
    if (k === '_extra') {
      try { var x = JSON.parse(v); for (var kk in x) if (!(kk in doc)) doc[kk] = x[kk]; } catch (err) {}
    } else if (t === 'n') {
      var num = Number(v); if (v !== '' && isFinite(num)) doc[k] = num;
    } else if (t === 'b') {
      doc[k] = v === true || String(v).toUpperCase() === 'TRUE';
    } else if (t === 'a') {
      doc[k] = String(v).split(/[,\s]+/).filter(function (x) { return x; });
    } else {
      doc[k] = cellText_(v, k);
    }
  });
  return doc;
}

/* เซลล์ที่ถูกแก้ด้วยมือในชีตอาจกลายเป็นวันที่ ให้แปลงกลับเป็นข้อความรูปแบบเดิม */
function cellText_(v, k) {
  if (v instanceof Date) {
    var tz = Session.getScriptTimeZone();
    return Utilities.formatDate(v, tz, /time/i.test(k) ? 'HH:mm' : 'yyyy-MM-dd');
  }
  return String(v);
}

function merge_(a, b) {
  var o = {}, k;
  for (k in a) o[k] = a[k];
  for (k in b) o[k] = b[k];
  return o;
}

function folder_() {
  var p = PropertiesService.getScriptProperties(), id = p.getProperty('FOLDER_ID');
  if (id) { try { return DriveApp.getFolderById(id); } catch (err) {} }
  var it = DriveApp.getFoldersByName(FOLDER_NAME), f = it.hasNext() ? it.next() : DriveApp.createFolder(FOLDER_NAME);
  p.setProperty('FOLDER_ID', f.getId());
  return f;
}

function withLock_(fn) {
  var lock = LockService.getScriptLock();
  lock.waitLock(25000);
  try { return fn(); } finally { lock.releaseLock(); }
}

function json_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}
