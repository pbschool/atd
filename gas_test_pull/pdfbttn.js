// =========================================================================
// สคริปต์: pdfbtn (สำหรับปุ่ม Export เฉพาะแผ่นงาน)
// =========================================================================

// --------------------------------------------------------------------------
// ภาคผนวก: ฟังก์ชันตั้งค่าชื่อไฟล์ (ใช้ตัวช่วยเดิม)
// --------------------------------------------------------------------------

function getClassLevel_All(ss) {
  const classSheet = ss.getSheetByName('ชั้น');
  if (!classSheet) return ''; 
  const classLevel = classSheet.getRange('B6').getDisplayValue();
  return classLevel ? `${classLevel} ` : ''; 
}

function getFinalFileName(classLevel, sheetName) {
  const currentYear = Utilities.formatDate(new Date(), SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone(), "yyyy");
  return `${classLevel}${sheetName} (${currentYear}).pdf`;
}

// --------------------------------------------------------------------------
// ฟังก์ชันหลัก: การสร้าง PDF (สำหรับ นน.สส.)
// --------------------------------------------------------------------------

/**
 * 1. ฟังก์ชันหลักในการสร้างไฟล์ PDF สำหรับแผ่นงาน "นน.สส."
 */
function pdfbtnWH() { 
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();
  
  // *** 1. ตัวแปรหลักสำหรับแผ่นงาน นน.สส. ***
  const MASTER_SHEET_NAME = "นน.สส."; 
  
  const sheet = ss.getSheetByName(MASTER_SHEET_NAME);
  
  if (!sheet) {
      ui.alert(`⚠️ ไม่พบแผ่นงานมาสเตอร์ชื่อ: "${MASTER_SHEET_NAME}"`);
      return;
  }
  
  SpreadsheetApp.getActiveSpreadsheet().toast(`กำลังสร้างไฟล์ PDF สำหรับ ${MASTER_SHEET_NAME}...`, 'เริ่มต้น', 5);

  try {
    const sheetId = sheet.getSheetId();
    const token = ScriptApp.getOAuthToken();
    const headers = { 'Authorization': 'Bearer ' + token };
    const baseUrl = 'https://docs.google.com/spreadsheets/d/' + ss.getId() + '/export?';

    // 2. กำหนดพารามิเตอร์การส่งออกหลัก (การตั้งค่าเดียวกับ pdfMaAll2_Export.gs)
    let exportOptions = {
      format: 'pdf',
      size: 'A4', 
      portrait: false, // แนวนอน
      scale: 3, // Fit to Height
      gridlines: false, 
      top_margin: 0.3, 
      bottom_margin: 0.2, 
      left_margin: 0.5, 
      right_margin: 0.2,
      gid: sheetId,
      delay: 5, 
      printBackground: true
    };
    
    // 3. สร้าง URL และ Export
    const params = Object.keys(exportOptions).map(key => `${key}=${exportOptions[key]}`).join('&');
    const finalUrl = baseUrl + params;
    
    // 4. เพิ่มการรอคอยใน Apps Script 5 วินาที
    Utilities.sleep(5000); 

    const response = UrlFetchApp.fetch(finalUrl, { headers: headers });
    const pdfBlob = response.getBlob();
    
    // 5. บันทึกไฟล์
    const classLevel = getClassLevel_All(ss);
    const fileName = getFinalFileName(classLevel, MASTER_SHEET_NAME);
    
    pdfBlob.setName(fileName);
    const pdfFile = DriveApp.createFile(pdfBlob);
    pdfFile.setTrashed(true); // ย้ายไปถังขยะทันที

    // 6. แสดงผลและเปิดไฟล์ในเบราว์เซอร์
    const openUrl = pdfFile.getUrl(); 
    
    const htmlOutput = HtmlService
      .createHtmlOutput(`<script>window.open('${openUrl}', '_blank'); google.script.host.close();</script>`)
      .setWidth(100)
      .setHeight(1); 
    
    ui.showModalDialog(htmlOutput, 'เปิดไฟล์ PDF...');
    
    SpreadsheetApp.getActiveSpreadsheet().toast(`✅ ไฟล์ถูกสร้างแล้ว (อยู่ในถังขยะ)`, 'สำเร็จ', 5);

  } catch (error) {
    ui.alert(`❌ เกิดข้อผิดพลาดในการสร้าง PDF: ${error.message}`);
  }
}

// --------------------------------------------------------------------------
// ฟังก์ชันหลัก: การสร้าง PDF (สำหรับ สรุปมา)
// --------------------------------------------------------------------------

/**
 * 1. ฟังก์ชันหลักในการสร้างไฟล์ PDF สำหรับแผ่นงาน "สรุปมา"
 */
function pdfbtnMaResult() { 
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();
  
  // *** 1. ตัวแปรหลักสำหรับแผ่นงาน สรุปมา ***
  const MASTER_SHEET_NAME = "สรุปมา"; 
  
  const sheet = ss.getSheetByName(MASTER_SHEET_NAME);
  
  if (!sheet) {
      ui.alert(`⚠️ ไม่พบแผ่นงานมาสเตอร์ชื่อ: "${MASTER_SHEET_NAME}"`);
      return;
  }
  
  SpreadsheetApp.getActiveSpreadsheet().toast(`กำลังสร้างไฟล์ PDF สำหรับ ${MASTER_SHEET_NAME}...`, 'เริ่มต้น', 5);

  try {
    const sheetId = sheet.getSheetId();
    const token = ScriptApp.getOAuthToken();
    const headers = { 'Authorization': 'Bearer ' + token };
    const baseUrl = 'https://docs.google.com/spreadsheets/d/' + ss.getId() + '/export?';

    // 2. กำหนดพารามิเตอร์การส่งออกหลัก (การตั้งค่าเดียวกับ pdfMaAll2_Export.gs)
    let exportOptions = {
      format: 'pdf',
      size: 'A4', 
      portrait: false, // แนวนอน
      scale: 3, // Fit to Height
      gridlines: false, 
      top_margin: 0.3, 
      bottom_margin: 0.2, 
      left_margin: 0.5, 
      right_margin: 0.2,
      gid: sheetId,
      delay: 5, 
      printBackground: true
    };
    
    // 3. สร้าง URL และ Export
    const params = Object.keys(exportOptions).map(key => `${key}=${exportOptions[key]}`).join('&');
    const finalUrl = baseUrl + params;
    
    // 4. เพิ่มการรอคอยใน Apps Script 5 วินาที
    Utilities.sleep(5000); 

    const response = UrlFetchApp.fetch(finalUrl, { headers: headers });
    const pdfBlob = response.getBlob();
    
    // 5. บันทึกไฟล์
    const classLevel = getClassLevel_All(ss);
    const fileName = getFinalFileName(classLevel, MASTER_SHEET_NAME);
    
    pdfBlob.setName(fileName);
    const pdfFile = DriveApp.createFile(pdfBlob);
    pdfFile.setTrashed(true); // ย้ายไปถังขยะทันที

    // 6. แสดงผลและเปิดไฟล์ในเบราว์เซอร์
    const openUrl = pdfFile.getUrl(); 
    
    const htmlOutput = HtmlService
      .createHtmlOutput(`<script>window.open('${openUrl}', '_blank'); google.script.host.close();</script>`)
      .setWidth(100)
      .setHeight(1); 
    
    ui.showModalDialog(htmlOutput, 'เปิดไฟล์ PDF...');
    
    SpreadsheetApp.getActiveSpreadsheet().toast(`✅ ไฟล์ถูกสร้างแล้ว (อยู่ในถังขยะ)`, 'สำเร็จ', 5);

  } catch (error) {
    ui.alert(`❌ เกิดข้อผิดพลาดในการสร้าง PDF: ${error.message}`);
  }
}

// --------------------------------------------------------------------------
// ฟังก์ชันหลัก: การสร้าง PDF (สำหรับ สรุปกิจฯ)
// --------------------------------------------------------------------------

/**
 * 1. ฟังก์ชันหลักในการสร้างไฟล์ PDF สำหรับแผ่นงาน "สรุปกิจฯ"
 */
function pdfbtnStuActive() { 
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();
  
  // *** 1. ตัวแปรหลักสำหรับแผ่นงาน สรุปกิจฯ ***
  const MASTER_SHEET_NAME = "สรุปกิจฯ"; 
  
  const sheet = ss.getSheetByName(MASTER_SHEET_NAME);
  
  if (!sheet) {
      ui.alert(`⚠️ ไม่พบแผ่นงานมาสเตอร์ชื่อ: "${MASTER_SHEET_NAME}"`);
      return;
  }
  
  SpreadsheetApp.getActiveSpreadsheet().toast(`กำลังสร้างไฟล์ PDF สำหรับ ${MASTER_SHEET_NAME}...`, 'เริ่มต้น', 5);

  try {
    const sheetId = sheet.getSheetId();
    const token = ScriptApp.getOAuthToken();
    const headers = { 'Authorization': 'Bearer ' + token };
    const baseUrl = 'https://docs.google.com/spreadsheets/d/' + ss.getId() + '/export?';

    // 2. กำหนดพารามิเตอร์การส่งออกหลัก (การตั้งค่าเดียวกับ pdfMaAll2_Export.gs)
    let exportOptions = {
      format: 'pdf',
      size: 'A4', 
      portrait: false, // แนวนอน
      scale: 3, // Fit to Height
      gridlines: false, 
      top_margin: 0.3, 
      bottom_margin: 0.2, 
      left_margin: 0.5, 
      right_margin: 0.2,
      gid: sheetId,
      delay: 5, 
      printBackground: true
    };
    
    // 3. สร้าง URL และ Export
    const params = Object.keys(exportOptions).map(key => `${key}=${exportOptions[key]}`).join('&');
    const finalUrl = baseUrl + params;
    
    // 4. เพิ่มการรอคอยใน Apps Script 5 วินาที
    Utilities.sleep(5000); 

    const response = UrlFetchApp.fetch(finalUrl, { headers: headers });
    const pdfBlob = response.getBlob();
    
    // 5. บันทึกไฟล์
    const classLevel = getClassLevel_All(ss);
    const fileName = getFinalFileName(classLevel, MASTER_SHEET_NAME);
    
    pdfBlob.setName(fileName);
    const pdfFile = DriveApp.createFile(pdfBlob);
    pdfFile.setTrashed(true); // ย้ายไปถังขยะทันที

    // 6. แสดงผลและเปิดไฟล์ในเบราว์เซอร์
    const openUrl = pdfFile.getUrl(); 
    
    const htmlOutput = HtmlService
      .createHtmlOutput(`<script>window.open('${openUrl}', '_blank'); google.script.host.close();</script>`)
      .setWidth(100)
      .setHeight(1); 
    
    ui.showModalDialog(htmlOutput, 'เปิดไฟล์ PDF...');
    
    SpreadsheetApp.getActiveSpreadsheet().toast(`✅ ไฟล์ถูกสร้างแล้ว (อยู่ในถังขยะ)`, 'สำเร็จ', 5);

  } catch (error) {
    ui.alert(`❌ เกิดข้อผิดพลาดในการสร้าง PDF: ${error.message}`);
  }
}