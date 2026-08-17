// /**
//  * unCheckMaAll - ล้างข้อมูลการเช็คชื่อทั้งหมดในแผ่นงานปัจจุบัน
//  * ปรับปรุงให้ตรงกับชื่อตัวแปรใน Constants.gs (R_START_COL, R_END_COL, R_START_ROW, R_END_ROW)
//  */
// function unCheckMaAll() {
//   const ss = SpreadsheetApp.getActiveSpreadsheet();
//   const sheet = ss.getActiveSheet();
//   const sheetName = sheet.getName().trim();

//   // 1. ตรวจสอบก่อนว่าอยู่ในชีทรายเดือนหรือไม่ (เช็คจากรายการใน Constants.gs)
//   if (!MONTHLY_SHEETS.includes(sheetName)) {
//     ss.toast('⚠️ ฟังก์ชันนี้ใช้ได้เฉพาะในชีทรายเดือนเท่านั้น', 'ระงับการทำงาน');
//     return;
//   }

//   // 2. ดึงค่าขอบเขตจาก Constants.gs (ใช้ชื่อใหม่ที่ครูตั้งไว้)
//   const startCol = R_START_COL; // 11 (K)
//   const endCol = R_END_COL;     // 41 (AO)
//   const numCols = endCol - startCol + 1;
//   const numRows = R_END_ROW - R_START_ROW + 1; // แถว 7 ถึง 26

//   try {
//     // 3. ล้างข้อมูลสถานะนักเรียน (แถว 7-26)
//     // ใช้ clearContent เพื่อลบทั้งคำว่า "มา" และ "ค่าล่องหน"
//     sheet.getRange(R_START_ROW, startCol, numRows, numCols).clearContent();
    
//     // 4. ล้างค่า Checkbox ในแถวตัวควบคุม (แถว 6 และ 27)
//     // การ clearContent จะทำให้ Checkbox กลับไปเป็นค่าว่าง (Unchecked)
//     sheet.getRange(6, startCol, 1, numCols).clearContent();
//     sheet.getRange(27, startCol, 1, numCols).clearContent();
    
//     // แจ้งเตือนสถานะ
//     ss.toast('🗑️ ล้างข้อมูลการเช็คชื่อเดือน ' + sheetName + ' เรียบร้อยแล้ว!', 'สำเร็จ', 3);
    
//   } catch (e) {
//     Logger.log("Error in unCheckMaAll: " + e.message);
//     ss.toast('🚫 ไม่สามารถล้างข้อมูลได้: ' + e.message, 'ข้อผิดพลาด');
//   }
// }