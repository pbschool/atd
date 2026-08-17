function checkMaAll() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getActiveSheet(); 
  
  // ตรวจสอบว่าเป็นชีทรายเดือนที่อนุญาตให้รันหรือไม่
  const sheetName = sheet.getName().trim();
  if (!MONTHLY_SHEETS.includes(sheetName)) {
    ss.toast("⚠️ ฟังก์ชันนี้ใช้ได้เฉพาะในแผ่นงานรายเดือนเท่านั้น", "ระงับการทำงาน");
    return;
  }

  const numRows = DATA_END_ROW - DATA_START_ROW + 1;
  const numCols = STATUS_END_COL - STATUS_START_COL + 1;

  // ดึงข้อมูลแบบ Bulk
  const nameValues = sheet.getRange(DATA_START_ROW, NAMES_COL, numRows, 1).getValues(); 
  const dayOrderValues = sheet.getRange(DAY_ORDER_ROW, STATUS_START_COL, 1, numCols).getValues()[0]; 
  const currentData = sheet.getRange(DATA_START_ROW, STATUS_START_COL, numRows, numCols).getValues(); 
  
  const newValues = [];
  let countMa = 0;

  for (let i = 0; i < numRows; i++) {
    const rowOutput = [];
    const hasName = nameValues[i][0] && String(nameValues[i][0]).trim() !== ""; 

    for (let j = 0; j < numCols; j++) {
      const hasDayOrder = dayOrderValues[j] && String(dayOrderValues[j]).trim() !== "";
      
      if (hasName && hasDayOrder) {
        rowOutput.push("มา");
        countMa++;
      } else {
        rowOutput.push(currentData[i][j]); 
      }
    }
    newValues.push(rowOutput);
  }

  if (countMa > 0) {
    sheet.getRange(DATA_START_ROW, STATUS_START_COL, numRows, numCols).setValues(newValues);
    ss.toast(`✅ เติมสถานะ "มา" สำเร็จทั้งหมด ${countMa} จุด`, 'ระบบอัตโนมัติ');
  }
}

/**
 * unCheckMaAll - ล้างข้อมูลการเช็คชื่อทั้งหมดในแผ่นงานปัจจุบัน
 * ปรับปรุงให้ตรงกับชื่อตัวแปรใน Constants.gs (R_START_COL, R_END_COL, R_START_ROW, R_END_ROW)
 */
function unCheckMaAll() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getActiveSheet();
  const sheetName = sheet.getName().trim();

  // 1. ตรวจสอบก่อนว่าอยู่ในชีทรายเดือนหรือไม่ (เช็คจากรายการใน Constants.gs)
  if (!MONTHLY_SHEETS.includes(sheetName)) {
    ss.toast('⚠️ ฟังก์ชันนี้ใช้ได้เฉพาะในชีทรายเดือนเท่านั้น', 'ระงับการทำงาน');
    return;
  }

  // 2. ดึงค่าขอบเขตจาก Constants.gs (ใช้ชื่อใหม่ที่ครูตั้งไว้)
  const startCol = R_START_COL; // 11 (K)
  const endCol = R_END_COL;     // 41 (AO)
  const numCols = endCol - startCol + 1;
  const numRows = R_END_ROW - R_START_ROW + 1; // แถว 7 ถึง 26

  try {
    // 3. ล้างข้อมูลสถานะนักเรียน (แถว 7-26)
    // ใช้ clearContent เพื่อลบทั้งคำว่า "มา" และ "ค่าล่องหน"
    sheet.getRange(R_START_ROW, startCol, numRows, numCols).clearContent();
    
    // 4. ล้างค่า Checkbox ในแถวตัวควบคุม (แถว 6 และ 27)
    // การ clearContent จะทำให้ Checkbox กลับไปเป็นค่าว่าง (Unchecked)
    sheet.getRange(6, startCol, 1, numCols).clearContent();
    sheet.getRange(27, startCol, 1, numCols).clearContent();
    
    // แจ้งเตือนสถานะ
    ss.toast('🗑️ ล้างข้อมูลการเช็คชื่อเดือน ' + sheetName + ' เรียบร้อยแล้ว!', 'สำเร็จ', 3);
    
  } catch (e) {
    Logger.log("Error in unCheckMaAll: " + e.message);
    ss.toast('🚫 ไม่สามารถล้างข้อมูลได้: ' + e.message, 'ข้อผิดพลาด');
  }
}