/**
 * ==============================================================================
 * Google Apps Script: ฐานข้อมูลออนไลน์ส่วนกลางสำหรับระบบลงทะเบียนเข้าใช้งานอาคาร
 * อาคารอเนกประสงค์ ต.0025 การไฟฟ้าฝ่ายผลิตแห่งประเทศไทย สำนักงานไทรน้อย
 * ==============================================================================
 *
 * วิธีการติดตั้ง / อัปเดต (ทำครั้งเดียว ไม่เกิน 2 นาที):
 * 1. เปิด Google Sheets เปล่า (หรือไฟล์ชีตเดิมที่ใช้อยู่)
 * 2. เมนูด้านบน เลือก "ส่วนขยาย" (Extensions) > "Apps Script"
 * 3. ลบโค้ดเดิมออกทั้งหมด แล้วคัดลอกโค้ดในไฟล์นี้ไปวางแทนที่ทั้งหมด จากนั้นกดไอคอน "บันทึก" (Save 💾)
 * 4. คลิกปุ่มสีน้ำเงินมุมขวาบน "ทำให้ใช้งานได้" (Deploy) > "การจัดการทำให้ใช้งานได้" (Manage deployments)
 *    - หรือเลือก "การทำให้ใช้งานได้ใหม่" (New deployment)
 * 5. หากเป็นการอัปเดต: คลิกไอคอนรูปดินสอ ✏️ ด้านบนขวา เลือกเวอร์ชัน "เวอร์ชันใหม่" (New version)
 *    หากสร้างใหม่: คลิก ⚙️ เลือก "เว็บแอป" (Web app)
 *    - คำอธิบาย: ระบบลงทะเบียนอาคารและกำหนดการชมรม กฟผ.
 *    - ดำเนินการในฐานะ (Execute as): "ฉัน" (Me)
 *    - ผู้ที่มีสิทธิ์เข้าถึง (Who has access): "ทุกคน" (Anyone)  <--- สำคัญมาก ต้องเลือกทุกคน
 * 6. กดปุ่ม "ทำให้ใช้งานได้" (Deploy) แล้วอนุมัติสิทธิ์ (Review permissions)
 * 7. คัดลอก "URL ของเว็บแอป" (ลงท้ายด้วย /exec)
 * 8. นำ URL นั้นมาใส่ในตัวแปร DEFAULT_CLOUD_API_URL หรือวางในช่อง Cloud Webhook URL
 *
 * ผลลัพธ์:
 * - ข้อมูลผู้เข้าใช้งานทั้งหมด (AccessLogs) จะซิงก์กันทุกเครื่องแบบเรียลไทม์
 * - กำหนดการใช้งาน รายชื่อ และการตั้งค่าของทุกชมรม (ClubConfig) จะซิงก์ตรงกันทุกคอมพิวเตอร์และมือถือทันที!
 */

function doGet(e) {
  var action = (e && e.parameter) ? e.parameter.action : '';
  
  if (action === 'getLogs') {
    var logs = getAllRecords();
    var storedClubs = getStoredClubs();
    var res = { success: true, records: logs };
    if (storedClubs && storedClubs.length > 0) {
      res.clubs = storedClubs;
    }
    return responseJSON(res);
  }
  
  if (action === 'getClubs') {
    var clubs = getStoredClubs();
    return responseJSON({ success: true, clubs: clubs || [] });
  }

  return responseJSON({ 
    status: 'ok', 
    message: 'EGAT Building Access API is running successfully',
    totalRecords: getAllRecords().length,
    hasClubsConfig: !!getStoredClubs()
  });
}

function doPost(e) {
  try {
    var rawData = e.postData ? e.postData.contents : '';
    var data = {};
    if (rawData) {
      try {
        data = JSON.parse(rawData);
      } catch (err) {
        data = e.parameter || {};
      }
    } else {
      data = e.parameter || {};
    }

    var action = data.action || 'checkIn';

    // ─── ACTION: บันทึกข้อมูลการตั้งค่าชมรมและกำหนดการ (Club Configuration) ───
    if (action === 'saveClubs') {
      var clubs = data.clubs;
      if (typeof clubs === 'string') {
        try { clubs = JSON.parse(clubs); } catch(err) {}
      }
      if (clubs && Array.isArray(clubs)) {
        saveStoredClubs(clubs);
        return responseJSON({ success: true, count: clubs.length, message: 'Clubs configuration saved successfully' });
      }
      return responseJSON({ success: false, message: 'Invalid clubs data' });
    }

    var sheet = getOrCreateSheet();

    // ─── ACTION: เช็คอิน (Check-in) ───
    if (action === 'checkIn') {
      var id = data.id || ('REC-' + Utilities.formatDate(new Date(), 'GMT+7', 'yyyyMMdd-HHmmss'));
      sheet.appendRow([
        id,
        data.userName || '',
        data.department || '',
        data.empId || '',
        data.clubId || '',
        data.clubName || '',
        data.checkInTime || new Date().toISOString(),
        '', // checkOutTime (เว้นว่างไว้จนกว่าจะเช็คเอาต์)
        0,  // durationMinutes
        'active',
        data.note || '',
        new Date().toISOString()
      ]);
      return responseJSON({ success: true, id: id, message: 'Check-in recorded' });
    }

    // ─── ACTION: เช็คเอาต์ (Check-out) ───
    if (action === 'checkOut') {
      var id = data.id;
      var values = sheet.getDataRange().getValues();
      var found = false;
      for (var i = 1; i < values.length; i++) {
        // ค้นหาตาม ID หรือชื่อผู้ใช้ที่มีสถานะ active
        if ((id && values[i][0] == id) || (values[i][1] == data.userName && values[i][9] == 'active')) {
          var rowNum = i + 1;
          sheet.getRange(rowNum, 8).setValue(data.checkOutTime || new Date().toISOString());
          sheet.getRange(rowNum, 9).setValue(data.durationMinutes || 0);
          sheet.getRange(rowNum, 10).setValue('completed');
          if (data.note) {
            var oldNote = sheet.getRange(rowNum, 11).getValue();
            sheet.getRange(rowNum, 11).setValue(oldNote ? (oldNote + ' | ' + data.note) : data.note);
          }
          found = true;
          break;
        }
      }
      return responseJSON({ success: true, updated: found });
    }

    return responseJSON({ success: true, received: data });
  } catch (error) {
    return responseJSON({ success: false, error: error.toString() });
  }
}

// ─── ACCESS LOGS SHEET MANAGEMENT ───
function getOrCreateSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('AccessLogs');
  if (!sheet) {
    sheet = ss.insertSheet('AccessLogs');
    sheet.appendRow([
      'รหัสบันทึก', 'ชื่อ-นามสกุล', 'สังกัด/ฝ่าย', 'รหัสพนักงาน',
      'รหัสกิจกรรม', 'ชื่อกิจกรรม/ชมรม', 'เวลาเข้า', 'เวลาออก',
      'ระยะเวลา (นาที)', 'สถานะ', 'หมายเหตุ', 'บันทึกเมื่อ'
    ]);
    sheet.getRange(1, 1, 1, 12)
      .setBackground('#0f172a')
      .setFontColor('#ffffff')
      .setFontWeight('bold')
      .setHorizontalAlignment('center');
    sheet.setFrozenRows(1);
    sheet.setRowHeight(1, 35);
  }
  return sheet;
}

function getAllRecords() {
  var sheet = getOrCreateSheet();
  var rows = sheet.getDataRange().getValues();
  if (rows.length <= 1) return [];

  var list = [];
  // วนลูปจากแถวล่าสุดขึ้นไป (เพื่อให้ได้ข้อมูลใหม่ล่าสุดอยู่ด้านบน)
  for (var i = rows.length - 1; i >= 1; i--) {
    var r = rows[i];
    if (!r[0] && !r[1]) continue;

    var checkInVal = r[6];
    if (checkInVal instanceof Date) {
      checkInVal = checkInVal.toISOString();
    } else {
      checkInVal = String(checkInVal || '');
    }

    var checkOutVal = r[7];
    if (checkOutVal instanceof Date) {
      checkOutVal = checkOutVal.toISOString();
    } else {
      checkOutVal = checkOutVal ? String(checkOutVal) : null;
    }

    list.push({
      id: String(r[0] || ''),
      userName: String(r[1] || ''),
      department: String(r[2] || ''),
      empId: String(r[3] || ''),
      clubId: String(r[4] || ''),
      clubName: String(r[5] || ''),
      checkInTime: checkInVal,
      checkOutTime: checkOutVal,
      durationMinutes: Number(r[8]) || 0,
      status: String(r[9] || 'completed'),
      note: String(r[10] || ''),
      createdAt: r[11] instanceof Date ? r[11].toISOString() : String(r[11] || '')
    });
  }
  return list;
}

// ─── CLUB CONFIGURATION SHEET & SCRIPT PROPERTIES ───
function getOrCreateClubsSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('ClubConfig');
  if (!sheet) {
    sheet = ss.insertSheet('ClubConfig');
    sheet.appendRow([
      'รหัสชมรม (ID)', 'ชื่อชมรม', 'ชื่อภาษาอังกฤษ', 'ไอคอน', 'สีประจำชมรม',
      'โซน/สนาม', 'รายละเอียด', 'วันเปิดใช้งาน (0=อาทิตย์..6=เสาร์)', 
      'เวลาเปิด', 'เวลาปิด', 'โหมดสถานะ', 'หมายเหตุ', 'อัปเดตล่าสุด'
    ]);
    sheet.getRange(1, 1, 1, 13)
      .setBackground('#1e293b')
      .setFontColor('#38bdf8')
      .setFontWeight('bold')
      .setHorizontalAlignment('center');
    sheet.setFrozenRows(1);
    sheet.setRowHeight(1, 35);
  }
  return sheet;
}

function saveStoredClubs(clubs) {
  if (!clubs || !Array.isArray(clubs)) return;
  
  // 1. บันทึกลง ScriptProperties เพื่อความรวดเร็วและโครงสร้าง JSON ที่สมบูรณ์ 100%
  try {
    PropertiesService.getScriptProperties().setProperty('CLUBS_CONFIG_JSON', JSON.stringify(clubs));
  } catch (err) {
    Logger.log('Error saving to ScriptProperties: ' + err);
  }

  // 2. บันทึกเป็นตารางที่อ่านง่ายลงชีต ClubConfig เพื่อให้ Admin สามารถเปิดดูใน Google Sheets ได้
  try {
    var sheet = getOrCreateClubsSheet();
    var lastRow = sheet.getLastRow();
    if (lastRow > 1) {
      sheet.getRange(2, 1, lastRow - 1, 13).clearContent();
    }
    
    var rows = [];
    var nowStr = Utilities.formatDate(new Date(), 'GMT+7', 'yyyy-MM-dd HH:mm:ss');
    for (var i = 0; i < clubs.length; i++) {
      var c = clubs[i];
      var openDaysStr = Array.isArray(c.openDays) ? c.openDays.join(',') : (c.openDays || '');
      rows.push([
        c.id || '',
        c.name || '',
        c.enName || '',
        c.icon || '',
        c.color || '',
        c.zone || '',
        c.desc || '',
        openDaysStr,
        c.startTime || '',
        c.endTime || '',
        c.statusMode || 'auto',
        c.note || '',
        nowStr
      ]);
    }
    if (rows.length > 0) {
      sheet.getRange(2, 1, rows.length, 13).setValues(rows);
    }
  } catch (err) {
    Logger.log('Error saving to ClubConfig sheet: ' + err);
  }
}

function getStoredClubs() {
  // 1. ลองอ่านจาก ScriptProperties ก่อน
  try {
    var prop = PropertiesService.getScriptProperties().getProperty('CLUBS_CONFIG_JSON');
    if (prop) {
      var parsed = JSON.parse(prop);
      if (Array.isArray(parsed) && parsed.length > 0) {
        return parsed;
      }
    }
  } catch (err) {
    Logger.log('Error reading ScriptProperties: ' + err);
  }

  // 2. Fallback: อ่านจากแถวในชีต ClubConfig
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName('ClubConfig');
    if (sheet && sheet.getLastRow() > 1) {
      var values = sheet.getRange(2, 1, sheet.getLastRow() - 1, 13).getValues();
      var list = [];
      for (var i = 0; i < values.length; i++) {
        var r = values[i];
        if (!r[0] && !r[1]) continue;
        var openDays = [];
        if (r[7] !== '' && r[7] !== null && r[7] !== undefined) {
          openDays = String(r[7]).split(',').map(function(n) { return parseInt(n.trim(), 10); }).filter(function(n) { return !isNaN(n); });
        }
        list.push({
          id: String(r[0] || ''),
          name: String(r[1] || ''),
          enName: String(r[2] || ''),
          icon: String(r[3] || '🏆'),
          color: String(r[4] || '#3b82f6'),
          zone: String(r[5] || ''),
          desc: String(r[6] || ''),
          openDays: openDays,
          startTime: String(r[8] || ''),
          endTime: String(r[9] || ''),
          statusMode: String(r[10] || 'auto'),
          note: String(r[11] || '')
        });
      }
      if (list.length > 0) return list;
    }
  } catch (err) {
    Logger.log('Error reading ClubConfig sheet: ' + err);
  }

  return null;
}

function responseJSON(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
