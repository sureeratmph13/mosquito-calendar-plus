/**
 * ปฏิทินกันยุงออนไลน์ หมู่ 5 บ้านจีต — Google Apps Script
 *
 * วิธีติดตั้ง (ทำครั้งเดียว)
 * 1. สร้าง Google Sheets ใหม่ 1 ไฟล์ > เมนู ส่วนขยาย > Apps Script
 * 2. วางโค้ดนี้ใน Code.gs และสร้างไฟล์ HTML ชื่อ Index แล้ววางโค้ด Index.html
 * 3. ตั้งค่า ADMIN_PIN และ STAFF_PASSWORD ใน Project Settings > Script Properties
 * 4. เลือกฟังก์ชัน setup แล้วกด "เรียกใช้" 1 ครั้ง (อนุญาตสิทธิ์)
 * 5. การทำให้ใช้งานได้ > การทำให้ใช้งานได้รายการใหม่ > ประเภท: เว็บแอป
 *    ดำเนินการในฐานะ: ฉัน · ผู้ที่มีสิทธิ์เข้าถึง: ทุกคน > ทำให้ใช้งานได้ > คัดลอก URL เว็บแอป
 */

// รหัสเข้าโหมด อสม. — ดูภาพรวม บันทึกผล จัดการรายชื่อบ้าน (ไม่เห็นลิงก์ Google Sheets)
// ตั้งค่า ADMIN_PIN ใน Script Properties เท่านั้น
// รหัสเจ้าหน้าที่ — ทำได้ทุกอย่างเหมือน อสม. และเปิดข้อมูลใน Google Sheets ได้
// ตั้งค่า STAFF_PASSWORD ใน Script Properties เท่านั้น

var SHEET_HOUSES = 'รายชื่อบ้าน';
var SHEET_CHECKS = 'ผลสำรวจ';
var HOUSE_HEAD = ['บ้านเลขที่', 'ชื่อเจ้าบ้าน', 'อสม. ผู้ดูแล', 'วันที่เพิ่ม'];
var CHECK_HEAD = ['รหัส', 'วันศุกร์', 'บ้านเลขที่', 'ปิด', 'เปลี่ยน', 'ปล่อย', 'ปรับ', 'ปฏิบัติ', 'ขัด',
                  'ภาชนะที่สำรวจ', 'ภาชนะพบลูกน้ำ', 'อสม. ตรวจแล้ว', 'ผู้บันทึก', 'เวลาบันทึก'];
var MEAS = ['pid', 'plian', 'ploi', 'prab', 'patibat', 'khat'];

function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('ปฏิทินกันยุงออนไลน์ หมู่ 5 บ้านจีต')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.DEFAULT);
}

/** เรียกใช้ 1 ครั้งหลังวางโค้ด เพื่อสร้างแผ่นงาน */
function setup() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) {
    ss = SpreadsheetApp.create('ปฏิทินกันยุง หมู่ 5 บ้านจีต (ข้อมูล)');
  }
  PropertiesService.getScriptProperties().setProperty('SHEET_ID', ss.getId());
  var h = ss.getSheetByName(SHEET_HOUSES) || ss.insertSheet(SHEET_HOUSES);
  if (h.getLastRow() === 0) h.appendRow(HOUSE_HEAD);
  h.getRange('A:A').setNumberFormat('@');
  h.setFrozenRows(1);
  var c = ss.getSheetByName(SHEET_CHECKS) || ss.insertSheet(SHEET_CHECKS);
  if (c.getLastRow() === 0) c.appendRow(CHECK_HEAD);
  c.getRange('A:C').setNumberFormat('@');
  c.setFrozenRows(1);
  [h, c].forEach(function (s) {
    s.getRange(1, 1, 1, s.getLastColumn()).setFontWeight('bold').setBackground('#0A62C9').setFontColor('#FFFFFF');
  });
  Logger.log('พร้อมใช้งาน ไฟล์ข้อมูล: ' + ss.getUrl());
}

/* ---------------- helpers ---------------- */
function ss_() {
  var id = PropertiesService.getScriptProperties().getProperty('SHEET_ID');
  var ss = id ? SpreadsheetApp.openById(id) : SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) throw new Error('ยังไม่ได้ตั้งค่า ให้เรียกใช้ฟังก์ชัน setup ก่อน');
  return ss;
}
function sheet_(name) {
  var s = ss_().getSheetByName(name);
  if (!s) throw new Error('ไม่พบแผ่นงาน "' + name + '" ให้เรียกใช้ฟังก์ชัน setup อีกครั้ง');
  return s;
}
function norm_(no) { return String(no == null ? '' : no).trim(); }
/** คืนบทบาท 'staff' หรือ 'vhv' ถ้ารหัสไม่ถูกต้องจะหยุดทำงาน */
function checkPin_(pin) {
  pin = String(pin == null ? '' : pin);
  var properties = PropertiesService.getScriptProperties();
  var staffPassword = properties.getProperty('STAFF_PASSWORD');
  var adminPin = properties.getProperty('ADMIN_PIN');
  if (!staffPassword || !adminPin || !staffPassword.trim() || !adminPin.trim() || staffPassword === adminPin) {
    throw new Error('กรุณาตั้งค่า ADMIN_PIN และ STAFF_PASSWORD ให้ต่างกันใน Script Properties');
  }
  if (!pin) throw new Error('รหัสผ่านไม่ถูกต้อง');
  if (pin === staffPassword) return 'staff';
  if (pin === adminPin) return 'vhv';
  throw new Error('รหัสผ่านไม่ถูกต้อง');
}
function validFriday_(f) {
  if (!/^\d{4}-\d{2}-\d{2}$/.test(f)) throw new Error('วันที่ไม่ถูกต้อง');
  var p = f.split('-'); var d = new Date(+p[0], +p[1] - 1, +p[2]);
  if (d.getDay() !== 5) throw new Error('ต้องเป็นวันศุกร์');
  return f;
}
function readHouses_() {
  var s = sheet_(SHEET_HOUSES); var n = s.getLastRow();
  if (n < 2) return [];
  return s.getRange(2, 1, n - 1, 3).getDisplayValues()
    .filter(function (r) { return norm_(r[0]); })
    .map(function (r) { return { no: norm_(r[0]), owner: r[1], vhv: r[2] }; });
}
function readChecks_(monthKey, onlyHouse) {
  var s = sheet_(SHEET_CHECKS); var n = s.getLastRow();
  if (n < 2) return [];
  var rows = s.getRange(2, 1, n - 1, CHECK_HEAD.length).getValues();
  var out = [];
  rows.forEach(function (r) {
    var fri = norm_(r[1]);
    if (fri.indexOf(monthKey) !== 0) return;
    var no = norm_(r[2]);
    if (onlyHouse != null && no !== onlyHouse) return;
    var e = { friday: fri, no: no, containers: Number(r[9]) || 0, larvae: Number(r[10]) || 0, vhv: r[11] === true, by: r[12] };
    MEAS.forEach(function (k, i) { e[k] = r[3 + i] === true; });
    out.push(e);
  });
  return out;
}
function writeCheck_(no, friday, entry, by, allowVhv) {
  var s = sheet_(SHEET_CHECKS);
  var key = friday + '|' + no;
  var found = s.getRange('A:A').createTextFinder(key).matchEntireCell(true).findNext();
  var oldVhv = false;
  if (found) oldVhv = s.getRange(found.getRow(), 12).getValue() === true;
  var row = [key, friday, no];
  MEAS.forEach(function (k) { row.push(entry[k] === true); });
  row.push(Math.max(0, parseInt(entry.containers, 10) || 0));
  row.push(Math.max(0, parseInt(entry.larvae, 10) || 0));
  row.push(allowVhv ? entry.vhv === true : oldVhv);
  row.push(by);
  row.push(new Date());
  if (found) s.getRange(found.getRow(), 1, 1, row.length).setValues([row]);
  else s.appendRow(row);
}

/* ---------------- สำหรับเจ้าของบ้าน (ไม่ต้องใช้รหัส) ---------------- */

/** ตรวจว่ามีบ้านเลขที่นี้ และส่งผลของบ้านนี้ในเดือนที่เลือก */
function homeData(no, monthKey) {
  no = norm_(no);
  if (!/^\d{4}-\d{2}$/.test(monthKey)) throw new Error('เดือนไม่ถูกต้อง');
  var exists = readHouses_().some(function (h) { return h.no === no; });
  if (!exists) return { found: false };
  return { found: true, no: no, checks: readChecks_(monthKey, no) };
}

function homeSave(no, friday, entry) {
  no = norm_(no); validFriday_(friday);
  var lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    var exists = readHouses_().some(function (h) { return h.no === no; });
    if (!exists) throw new Error('ไม่พบบ้านเลขที่ ' + no + ' ในรายชื่อ แจ้ง อสม. ให้เพิ่มก่อน');
    writeCheck_(no, friday, entry || {}, 'เจ้าของบ้าน', false);
  } finally { lock.releaseLock(); }
  return true;
}

/* ---------------- สำหรับ อสม. (ต้องใช้รหัส) ---------------- */

function adminData(pin, monthKey) {
  var role = checkPin_(pin);
  if (!/^\d{4}-\d{2}$/.test(monthKey)) throw new Error('เดือนไม่ถูกต้อง');
  return {
    role: role,
    houses: readHouses_(),
    checks: readChecks_(monthKey, null),
    sheetUrl: role === 'staff' ? ss_().getUrl() : ''
  };
}

function adminSave(pin, no, friday, entry) {
  var role = checkPin_(pin); no = norm_(no); validFriday_(friday);
  var lock = LockService.getScriptLock(); lock.waitLock(20000);
  try { writeCheck_(no, friday, entry || {}, role === 'staff' ? 'เจ้าหน้าที่' : 'อสม.', true); } finally { lock.releaseLock(); }
  return true;
}

function adminAddHouses(pin, rows) {
  checkPin_(pin);
  var lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    var have = {}; readHouses_().forEach(function (h) { have[h.no] = true; });
    var add = [];
    (rows || []).forEach(function (r) {
      var no = norm_(r.no); if (!no || have[no]) return; have[no] = true;
      add.push([no, String(r.owner || '').trim(), String(r.vhv || '').trim(), new Date()]);
    });
    if (add.length) {
      var s = sheet_(SHEET_HOUSES);
      s.getRange(s.getLastRow() + 1, 1, add.length, 4).setValues(add);
    }
    return { added: add.length, skipped: (rows || []).length - add.length };
  } finally { lock.releaseLock(); }
}

function adminDeleteHouse(pin, no) {
  checkPin_(pin); no = norm_(no);
  var lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    var s = sheet_(SHEET_HOUSES);
    var f = s.getRange('A:A').createTextFinder(no).matchEntireCell(true).findNext();
    if (f && f.getRow() > 1) s.deleteRow(f.getRow());
  } finally { lock.releaseLock(); }
  return true;
}
