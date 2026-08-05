/**
 * Classroom Duty Wizard — Google Sheets Backend (Google Apps Script)
 * ------------------------------------------------------------------
 * ทำหน้าที่เป็น "คลาวด์ฐานข้อมูล" แบบง่ายๆ (Key-Value store) เก็บไว้ใน Google Sheet
 * ให้แอป Classroom Duty Wizard เรียกใช้แทน/เสริมจาก localStorage
 * โครงสร้าง key ใช้รูปแบบเดียวกับที่แอปใช้ใน localStorage อยู่แล้ว เช่น
 *   wizard_registered_accounts
 *   wizard_admin_passcode
 *   wizard_<classroomKey>_students
 *   wizard_<classroomKey>_assignments
 *   wizard_<classroomKey>_attendance
 *   wizard_<classroomKey>_types
 *   wizard_<classroomKey>_notes
 *
 * วิธีติดตั้ง:
 * 1. สร้าง Google Sheet ใหม่ (ชื่ออะไรก็ได้)
 * 2. เมนู Extensions > Apps Script
 * 3. ลบโค้ดตัวอย่างเดิมทั้งหมด แล้ววางโค้ดไฟล์นี้ทับ
 * 4. กด Deploy > New deployment
 *      - เลือกประเภท (Select type): Web app
 *      - Execute as: Me
 *      - Who has access: Anyone
 * 5. กด Deploy แล้วอนุญาตสิทธิ์ (Authorize) ตามที่ Google ถาม
 * 6. คัดลอก "Web app URL" ที่ได้ (ลงท้ายด้วย /exec) ไปใช้ในหน้าเว็บ (ดูตัวอย่างการเรียกใช้ด้านล่างในแชท)
 *
 * หมายเหตุ: ทุกครั้งที่แก้โค้ดนี้ ต้องกด Deploy > Manage deployments > แก้ไข (ไอคอนดินสอ)
 * แล้วเลือก "New version" ใหม่ ไม่งั้น Web app URL จะยังใช้โค้ดเวอร์ชันเก่าอยู่
 */

const SHEET_NAME = 'KV_Store';

// ---------- Sheet helpers ----------

function getSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    sheet.appendRow(['key', 'value', 'updatedAt']);
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function findRow_(sheet, key) {
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === key) return i + 1; // เลขแถวจริงใน sheet (1-indexed)
  }
  return -1;
}

function setKey_(sheet, key, value) {
  const row = findRow_(sheet, key);
  const json = JSON.stringify(value);
  const now = new Date();
  if (row === -1) {
    sheet.appendRow([key, json, now]);
  } else {
    sheet.getRange(row, 2, 1, 2).setValues([[json, now]]);
  }
}

function jsonResponse_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

// ---------- GET: อ่านข้อมูล ----------
// ?action=get&key=wizard_admin_passcode
// ?action=list&prefix=wizard_

function doGet(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const sheet = getSheet_();
    const action = e.parameter.action || 'get';

    if (action === 'get') {
      const key = e.parameter.key;
      if (!key) return jsonResponse_({ ok: false, error: 'missing key' });
      const row = findRow_(sheet, key);
      if (row === -1) return jsonResponse_({ ok: true, key: key, value: null });
      const raw = sheet.getRange(row, 2).getValue();
      return jsonResponse_({ ok: true, key: key, value: raw ? JSON.parse(raw) : null });
    }

    if (action === 'list') {
      const prefix = e.parameter.prefix || '';
      const data = sheet.getDataRange().getValues();
      const keys = [];
      for (let i = 1; i < data.length; i++) {
        if (String(data[i][0]).indexOf(prefix) === 0) keys.push(data[i][0]);
      }
      return jsonResponse_({ ok: true, keys: keys });
    }

    return jsonResponse_({ ok: false, error: 'unknown action: ' + action });
  } catch (err) {
    return jsonResponse_({ ok: false, error: String(err) });
  } finally {
    lock.releaseLock();
  }
}

// ---------- POST: เขียน/ลบข้อมูล ----------
// body: { action: 'set', key: '...', value: {...} }
// body: { action: 'bulkSet', items: [{key, value}, ...] }
// body: { action: 'delete', key: '...' }

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const body = JSON.parse(e.postData.contents);
    const sheet = getSheet_();
    const action = body.action;

    if (action === 'set') {
      setKey_(sheet, body.key, body.value);
      return jsonResponse_({ ok: true });
    }

    if (action === 'bulkSet') {
      const items = body.items || [];
      items.forEach(function (item) { setKey_(sheet, item.key, item.value); });
      return jsonResponse_({ ok: true, count: items.length });
    }

    if (action === 'delete') {
      const row = findRow_(sheet, body.key);
      if (row !== -1) sheet.deleteRow(row);
      return jsonResponse_({ ok: true });
    }

    if (action === 'clearAll') {
      const lastRow = sheet.getLastRow();
      if (lastRow > 1) sheet.deleteRows(2, lastRow - 1); // เก็บแถวหัวตาราง (header) ไว้
      return jsonResponse_({ ok: true });
    }

    return jsonResponse_({ ok: false, error: 'unknown action: ' + action });
  } catch (err) {
    return jsonResponse_({ ok: false, error: String(err) });
  } finally {
    lock.releaseLock();
  }
}
