// copyCompactToRec.gs

/**
 * คัดลอกข้อมูลจาก recCompact!G3:K ไปยัง rec!A3:E
 * โดยจะล้างข้อมูลเดิมใน rec!A3:E ก่อน
 */
function copy2_CompactToRec() {
  const recSheetName = REC_SHEET_NAME;        // ใช้ REC_SHEET_NAME
  const compactSheetName = COMPACT_SHEET_NAME; // ใช้ COMPACT_SHEET_NAME
  const recHeaderRow = REC_HEADER_ROW;         // ใช้ REC_HEADER_ROW (สำคัญ)
  
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const recCompactSheet = ss.getSheetByName(compactSheetName);
  if (!recCompactSheet) {
    Logger.log(`Error: Source sheet '${compactSheetName}' not found. Please make sure the sheet exists.`);
    return;
  }

  const recSheet = ss.getSheetByName(recSheetName);
  if (!recSheet) {
    Logger.log(`Error: Destination sheet '${recSheetName}' not found. Please make sure the sheet exists.`);
    return;
  }

  Logger.log("--- เริ่มการทำงาน: คัดลอกข้อมูลจาก recCompact (G:K) ไปยัง rec (A:E) ---");

  const START_ROW = 3;
  // NUM_COLS_TO_COPY จะอ้างอิง REC_HEADER_ROW.length ซึ่งเราได้ปรับใน Constants.gs แล้ว
  const NUM_COLS_TO_COPY = recHeaderRow.length; 

  const SOURCE_START_COL = 7; // คอลัมน์ G
  const sourceLastRow = recCompactSheet.getLastRow();
  const numRowsToRead = Math.min(sourceLastRow - START_ROW + 1, 90000 - START_ROW + 1); 
  
  let dataToCopy = [];
  if (numRowsToRead > 0) {
    dataToCopy = recCompactSheet.getRange(START_ROW, SOURCE_START_COL, numRowsToRead, NUM_COLS_TO_COPY).getValues();
    Logger.log(`อ่านข้อมูล ${dataToCopy.length} แถว (G:K) จาก '${compactSheetName}'`);
  } else {
    Logger.log(`ชีท '${compactSheetName}' ไม่มีข้อมูลตั้งแต่แถว ${START_ROW} ในคอลัมน์ G:K`);
  }

  const DEST_START_COL = 1; // คอลัมน์ A
  const recSheetLastRow = recSheet.getLastRow();
  const numRowsToClear = Math.min(Math.max(0, recSheetLastRow - START_ROW + 1), 90000 - START_ROW + 1); 
  
  if (numRowsToClear > 0) {
    recSheet.getRange(START_ROW, DEST_START_COL, numRowsToClear, NUM_COLS_TO_COPY).clearContent();
    Logger.log(`ล้างข้อมูล ${numRowsToClear} แถวใน '${recSheetName}' (A${START_ROW}:E) สำเร็จ`);
  } else {
    Logger.log(`ไม่มีข้อมูลใน '${recSheetName}' (A${START_ROW}:E) ที่จะล้าง`);
  }

  if (dataToCopy.length > 0) {
    recSheet.getRange(START_ROW, DEST_START_COL, dataToCopy.length, dataToCopy[0].length).setValues(dataToCopy);
    Logger.log(`คัดลอกข้อมูล ${dataToCopy.length} แถวไปยัง '${recSheetName}' (A${START_ROW}:E) สำเร็จ`);
  } else {
    Logger.log("ไม่มีข้อมูลที่จะคัดลอก.");
  }

  Logger.log("--- สิ้นสุดการทำงาน: คัดลอกข้อมูล ---");
}