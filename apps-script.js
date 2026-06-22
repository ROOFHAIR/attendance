// ================================================================
// Google Apps Script — ระบบลงเวลา ROOFHAIR
// วิธีติดตั้ง:
//   1. ไปที่ https://script.google.com → New project
//   2. วางโค้ดนี้ทั้งหมดแทนที่โค้ดเดิม
//   3. แก้ SHEET_ID และ FOLDER_NAME ตามต้องการ
//   4. Deploy → New deployment → Web app
//      - Execute as: Me
//      - Who has access: Anyone
//   5. Copy Web app URL ไปวางใน index.html ที่ APPS_SCRIPT_URL
// ================================================================

const SHEET_NAME = 'ลงเวลา';       // ชื่อ sheet (แก้ได้)
const FOLDER_NAME = 'ROOFHAIR-Attendance-Photos'; // โฟลเดอร์ใน Drive

function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    const { name, photo, timestamp } = data;

    if (!name || !photo || !timestamp) {
      return jsonResponse({ status: 'error', message: 'ข้อมูลไม่ครบ' });
    }

    // 1. บันทึกรูปลง Google Drive
    const folder = getOrCreateFolder(FOLDER_NAME);
    const filename = `${timestamp.replace(/[/:]/g, '-')}_${name}.jpg`;
    const blob = Utilities.newBlob(
      Utilities.base64Decode(photo.replace(/^data:image\/\w+;base64,/, '')),
      'image/jpeg',
      filename
    );
    const file = folder.createFile(blob);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    const photoUrl = `https://drive.google.com/uc?id=${file.getId()}`;

    // 2. บันทึกลง Google Sheets
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName(SHEET_NAME);
    if (!sheet) {
      sheet = ss.insertSheet(SHEET_NAME);
      sheet.appendRow(['ลำดับ', 'ชื่อ-นามสกุล', 'วันที่-เวลา', 'ลิงก์รูปภาพ']);
      sheet.getRange(1, 1, 1, 4).setFontWeight('bold').setBackground('#1a1a2e').setFontColor('white');
    }

    const lastRow = sheet.getLastRow();
    sheet.appendRow([lastRow, name, timestamp, photoUrl]);

    return jsonResponse({ status: 'ok', message: 'บันทึกสำเร็จ', photoUrl });

  } catch (err) {
    return jsonResponse({ status: 'error', message: err.toString() });
  }
}

function getOrCreateFolder(name) {
  const folders = DriveApp.getFoldersByName(name);
  return folders.hasNext() ? folders.next() : DriveApp.createFolder(name);
}

function jsonResponse(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
