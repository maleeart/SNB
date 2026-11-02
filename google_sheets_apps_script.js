/**
 * ==============================================================================
 * Google Apps Script: ฐานข้อมูลออนไลน์ส่วนกลางสำหรับระบบลงทะเบียนเข้าใช้งานอาคาร
 * อาคารอเนกประสงค์ ต.0025 การไฟฟ้าฝ่ายผลิตแห่งประเทศไทย สำนักงานไทรน้อย
 * ==============================================================================
 *
 * วิธีการติดตั้ง (ทำครั้งเดียว ไม่เกิน 2 นาที):
 * 1. เปิด Google Sheets เปล่า (สร้างใหม่หรือใช้ไฟล์เดิม)
 * 2. เมนูด้านบน เลือก "ส่วนขยาย" (Extensions) > "Apps Script"
 * 3. ลบโค้ดเดิมออกทั้งหมด แล้ววางโค้ดในไฟล์นี้ลงไปทั้งหมด จากนั้นกดไอคอน "บันทึก" (Save)
 * 4. คลิกปุ่มสีน้ำเงินมุมขวาบน "ทำให้ใช้งานได้" (Deploy) > "การทำให้ใช้งานได้ใหม่" (New deployment)
 * 5. คลิกไอคอนรูปเฟือง ⚙️ ข้าง "เลือกประเภท" แล้วเลือก "เว็บแอป" (Web app)
 *    - คำอธิบาย: ระบบลงทะเบียนอาคาร กฟผ.
 *    - ดำเนินการในฐานะ (Execute as): "ฉัน" (Me)
 *    - ผู้ที่มีสิทธิ์เข้าถึง (Who has access): "ทุกคน" (Anyone)  <--- สำคัญมาก ต้องเลือกทุกคน
 * 6. กดปุ่ม "ทำให้ใช้งานได้" (Deploy) แล้วอนุมัติสิทธิ์ (Review permissions)
 * 7. คัดลอก "URL ของเว็บแอป" (ลงท้ายด้วย /exec)
 * 8. นำ URL นั้นมาวางในช่อง "Cloud Webhook URL" ในหน้าตั้งค่าของแอพ หรือใส่ในตัวแปร DEFAULT_CLOUD_API_URL
 *
 * ผลลัพธ์: มือถือทุกเครื่องที่สแกน QR และคอมพิวเตอร์ทุกเครื่องจะเชื่อมต่อและเห็นประวัติร่วมกันทันที!
 */

function doGet(e) {
  var action = (e && e.parameter) ? e.parameter.action : '';
  if (action === 'getLogs') {
    var logs = getAllRecords();
    return responseJSON({ success: true, records: logs });
  }
  return responseJSON({ 
    status: 'ok', 
    message: 'EGAT Building Access API is running successfully',
    totalRecords: getAllRecords().length
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

    var sheet = getOrCreateSheet();
    var action = data.action || 'checkIn';

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

function responseJSON(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
