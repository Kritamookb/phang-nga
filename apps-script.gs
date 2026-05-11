// Google Apps Script — Phang Nga Open 2026 Registration
// วิธีติดตั้ง:
//   1. เปิด Google Sheets ใหม่ → Extensions → Apps Script
//   2. วางโค้ดนี้แทนที่โค้ดเดิม
//   3. Deploy → New deployment → Web app
//      - Execute as: Me
//      - Who has access: Anyone
//   4. คัดลอก URL แล้วนำไปแทนที่ APPS_SCRIPT_URL ใน index.html

const SHEET_NAME = 'สมัคร';
const HEADERS = [
  'วันเวลาสมัคร', 'ประเภท',
  'ผู้เล่นที่ 1 – ชื่อ', 'ผู้เล่นที่ 1 – เบอร์',
  'ผู้เล่นที่ 2 – ชื่อ', 'ผู้เล่นที่ 2 – เบอร์',
];

function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    const ss   = SpreadsheetApp.getActiveSpreadsheet();
    let sheet  = ss.getSheetByName(SHEET_NAME);

    if (!sheet) {
      sheet = ss.insertSheet(SHEET_NAME);
      sheet.appendRow(HEADERS);
      sheet.getRange(1, 1, 1, HEADERS.length)
        .setFontWeight('bold')
        .setBackground('#1a56b0')
        .setFontColor('#ffffff');
      sheet.setFrozenRows(1);
    }

    sheet.appendRow([
      new Date(),
      data.category || '',
      data.p1_name  || '',
      data.p1_phone || '',
      data.p2_name  || '',
      data.p2_phone || '',
    ]);

    return ContentService
      .createTextOutput(JSON.stringify({ success: true }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ success: false, error: err.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// ทดสอบเพิ่มข้อมูลตัวอย่าง (รันจากใน Apps Script editor)
function testInsert() {
  const mock = {
    postData: {
      contents: JSON.stringify({
        category: 'NS',
        p1_name:  'สมชาย ทดสอบ',
        p1_phone: '0811234567',
        p2_name:  'สมหญิง ทดสอบ',
        p2_phone: '0819876543',
      }),
    },
  };
  const result = doPost(mock);
  Logger.log(result.getContent());
}
