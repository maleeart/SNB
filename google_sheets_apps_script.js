/**
 * Google Apps Script สำหรับเชื่อมโยงระบบลงทะเบียนเข้าใช้งานอาคารกับ Google Sheets
 * วิธีการติดตั้ง:
 * 1. เปิด Google Sheet ที่ต้องการเก็บข้อมูล
 * 2. ไปที่ Extensions (ส่วนขยาย) > Apps Script
 * 3. วางโค้ดนี้ลงใน Apps Script แล้วกด Save
 * 4. ไปที่ Deploy > New Deployment > เลือก Web App
 *    - Execute as: Me (ฉัน)
 *    - Who has access: Anyone (ทุกคน)
 * 5. คัดลอก Web App URL ที่ได้ นำมาใส่ในช่อง "Cloud Webhook URL" ในหน้าตั้งค่าของแอพ
 */

function doGet(e) {
  var action = e.parameter.action;
  if (action === 'getLogs') {
    return ContentService.createTextOutput(JSON.stringify(getAllRecords()))
      .setMimeType(ContentService.MimeType.JSON);
  }
  return ContentService.createTextOutput(JSON.stringify({ status: 'ok', message: 'API ready' }))
    .setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  try {
    var rawData = e.postData ? e.postData.contents : '';
    var data = {};
    if (rawData) {
      try {
        data = JSON.parse(rawData);
      } catch (err) {
        data = e.parameter;
      }
    } else {
      data = e.parameter;
    }

    var sheet = getOrCreateSheet();
    var action = data.action || 'checkIn';

    if (action === 'checkIn') {
      sheet.appendRow([
        data.id || Utilities.getUuid(),
        data.userName || '',
        data.department || '',
        data.empId || '',
        data.clubId || '',
        data.clubName || '',
        data.checkInTime || new Date().toISOString(),
        '', // checkOutTime
        '', // durationMinutes
        'active',
        data.note || '',
        new Date().toISOString()
      ]);
      return responseJSON({ success: true, message: 'Check-in recorded' });
    }

    if (action === 'checkOut') {
      var id = data.id;
      var values = sheet.getDataRange().getValues();
      var found = false;
      for (var i = 1; i < values.length; i++) {
        if (values[i][0] == id || (values[i][1] == data.userName && values[i][9] == 'active')) {
          sheet.getRange(i + 1, 8).setValue(data.checkOutTime || new Date().toISOString());
          sheet.getRange(i + 1, 9).setValue(data.durationMinutes || 0);
          sheet.getRange(i + 1, 10).setValue('completed');
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
      'ID', 'ชื่อ-นามสกุล', 'สังกัด', 'รหัสพนักงาน',
      'รหัสชมรม', 'ชื่อชมรม', 'เวลาเข้า', 'เวลาออก',
      'ระยะเวลา(นาที)', 'สถานะ', 'หมายเหตุ', 'บันทึกเมื่อ'
    ]);
    sheet.getRange(1, 1, 1, 12).setBackground('#1e293b').setFontColor('#ffffff').setFontWeight('bold');
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function getAllRecords() {
  var sheet = getOrCreateSheet();
  var rows = sheet.getDataRange().getValues();
  var headers = rows[0];
  var list = [];
  for (var i = 1; i < rows.length; i++) {
    var item = {};
    for (var j = 0; j < headers.length; j++) {
      item[headers[j]] = rows[i][j];
    }
    list.push(item);
  }
  return list;
}

function responseJSON(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON)
    .setHeader('Access-Control-Allow-Origin', '*');
}
