// onOpenTrigger.gs

/**
 * This function is automatically triggered when the spreadsheet is opened.
 * It checks if the active sheet is "สรุปมา" and calls reloadRefErrorsInActiveSheet if true.
 * Additionally, if the active sheet is a month sheet (01-12), it always calls Ad_reMonth
 * to pull data from 'rec', showing a modal during the process.
 *
 * ฟังก์ชันนี้จะทำงานโดยอัตโนมัติเมื่อเปิดสเปรดชีต
 * มันจะตรวจสอบว่าชีตที่เปิดอยู่คือ "สรุปมา" หรือไม่ และเรียกใช้ reloadRefErrorsInActiveSheet หากใช่
 * นอกจากนี้ หากชีตที่เปิดอยู่เป็นชีทเดือน (01-12) จะเรียกใช้ Ad_reMonth เสมอ
 * เพื่อดึงข้อมูลจาก 'rec' โดยจะแสดง modal ระหว่างการทำงาน
 */
/*
function onOpenTrigger() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const activeSheet = ss.getActiveSheet();
  const sheetName = activeSheet.getName();

  Logger.log(`Spreadsheet opened. Active sheet: "${sheetName}"`);

  // ตรวจสอบว่าชีทที่เปิดอยู่คือ "สรุปมา"
  if (sheetName === "สรุปมา") {
    Logger.log('Active sheet is "สรุปมา", calling reloadRefErrorsInActiveSheet...');
    reloadRefErrorsInActiveSheet(); // เรียกใช้ฟังก์ชันเพื่อโหลดสูตรใหม่
    ss.toast('✅ ข้อมูลในชีต "สรุปมา" ถูกโหลดใหม่แล้ว', 'โหลดข้อมูล', 10);
    return; // จบการทำงานหลังจากจัดการชีท "สรุปมา" แล้ว
  }

  // ตรวจสอบว่าเป็นชีทเดือน (01-12) หรือไม่
  const isMonthSheet = /^(0[1-9]|1[0-2])$/.test(sheetName);

  if (isMonthSheet) {
    Logger.log(`Active sheet "${sheetName}" is a month sheet. Calling Ad_reMonth via showAdReMonthModal...`);
    // เมื่อเปิดชีทเดือน (01-12) จะเรียก Ad_reMonth เสมอ
    Ad_reMonth();//showAdReMonthModal(); // ฟังก์ชันนี้จะแสดง Modal และเรียก Ad_reMonth()
  } else {
    Logger.log('Active sheet is not a month sheet or "สรุปมา", skipping specific actions.');
  }
}
*/