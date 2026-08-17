// copyDataWithinRecCompact.gs

/**
 * คัดลอกข้อมูลจาก recCompact!A3:E ไปยัง recCompact!G3:L
 */
function copy1_inRecCompact() {
  const sheetName = "recCompact";
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(sheetName);

  if (!sheet) {
    Logger.log(`Error: Sheet '${sheetName}' not found. Please ensure the sheet exists.`);
    return;
  }

  Logger.log("--- เริ่มการทำงาน: คัดลอกข้อมูลภายในชีท recCompact ---");

  const SOURCE_START_ROW = 3;
  const SOURCE_START_COL = 1; // คอลัมน์ A
  const NUM_COLS_TO_COPY = 5; // A, B, C, D, E (5 คอลัมน์)

  const DEST_START_ROW = 3;
  const DEST_START_COL = 7; // คอลัมน์ G
  const MAX_ROWS_TO_HANDLE = 90000;

  const lastRowWithData = sheet.getLastRow();
  const numRowsToRead = Math.min(lastRowWithData - SOURCE_START_ROW + 1, MAX_ROWS_TO_HANDLE - SOURCE_START_ROW + 1);

  let dataToCopy = [];
  if (numRowsToRead > 0) {
    dataToCopy = sheet.getRange(SOURCE_START_ROW, SOURCE_START_COL, numRowsToRead, NUM_COLS_TO_COPY).getValues();
    Logger.log(`อ่านข้อมูล ${dataToCopy.length} แถว (A:E) จาก '${sheetName}'`);
  } else {
    Logger.log(`ไม่มีข้อมูลใน '${sheetName}' ตั้งแต่แถว ${SOURCE_START_ROW} ถึง E`);
  }

  const numRowsToClear = Math.min(sheet.getLastRow() - DEST_START_ROW + 1, MAX_ROWS_TO_HANDLE - DEST_START_ROW + 1);
  
  if (numRowsToClear > 0) {
      sheet.getRange(DEST_START_ROW, DEST_START_COL, numRowsToClear, NUM_COLS_TO_COPY).clearContent();
      Logger.log(`ล้างข้อมูล ${numRowsToClear} แถวใน '${sheetName}' (G:L) สำเร็จ`);
  } else {
      Logger.log(`ไม่มีข้อมูลใน '${sheetName}' (G:L) ที่จะล้าง`);
  }

  if (dataToCopy.length > 0) {
    sheet.getRange(DEST_START_ROW, DEST_START_COL, dataToCopy.length, dataToCopy[0].length).setValues(dataToCopy);
    Logger.log(`คัดลอกข้อมูล ${dataToCopy.length} แถวไปยัง '${sheetName}' (G:L) สำเร็จ`);
  } else {
    Logger.log("ไม่มีข้อมูลที่จะคัดลอก.");
  }

  Logger.log("--- สิ้นสุดการทำงาน: คัดลอกข้อมูลภายในชีท recCompact ---");
}