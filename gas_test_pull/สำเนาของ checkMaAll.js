function checkMaAll() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getActiveSheet(); 
  
  const START_ROW = 6;
  const END_ROW = 25;
  const START_COL = 11; // K
  const END_COL = 41;   // AO
  const CHECKBOX_ROW = 27;

  const numRows = END_ROW - START_ROW + 1;
  const numCols = END_COL - START_COL + 1;

  // 1. ดึงข้อมูลมาทั้งหมด
  const cValues = sheet.getRange(START_ROW, 3, numRows, 1).getValues(); // Column C
  const header5 = sheet.getRange(5, START_COL, 1, numCols).getValues()[0]; // แถว 5 (K-AO)
  const currentData = sheet.getRange(START_ROW, START_COL, numRows, numCols).getValues(); 
  
  const newValues = [];
  const checkboxStates = new Array(numCols).fill(false);
  let countMa = 0;

  // วนลูปตามแถว (6-25)
  for (let i = 0; i < numRows; i++) {
    const rowValues = [];
    const hasName = cValues[i][0].toString().trim() !== ""; // คอลัมน์ C ไม่ว่าง

    // วนลูปตามคอลัมน์ (K-AO)
    for (let j = 0; j < numCols; j++) {
      const hasHeader = header5[j].toString().trim() !== ""; // แถว 5 ไม่ว่าง
      
      // เงื่อนไข: ถ้ามีชื่อใน C และ แถว 5 มีข้อมูล ให้ใส่ "มา"
      if (hasName && hasHeader) {
        rowValues.push("มา");
        checkboxStates[j] = true; // คอลัมน์นี้มีการลงข้อมูล ให้ Checkbox เป็น True
        countMa++;
      } else {
        // ถ้าเงื่อนไขไม่ครบ (เช่น แถว 5 ว่าง) ให้เอาค่าเดิมมาใส่ (ซึ่งมักจะเป็นค่าว่าง)
        rowValues.push(currentData[i][j]); 
      }
    }
    newValues.push(rowValues);
  }

  // 2. เขียนข้อมูลลงไป
  if (countMa > 0) {
    sheet.getRange(START_ROW, START_COL, numRows, numCols).setValues(newValues);
    
    // ตั้งค่า Checkbox แถวที่ 27
    // หมายเหตุ: checkboxStates จะเป็น True ถ้าคอลัมน์นั้นมี "มา" อย่างน้อย 1 แถว
    //sheet.getRange(CHECKBOX_ROW, START_COL, 1, numCols).setValues([checkboxStates]);
    
    ss.toast('✅ เติมคำว่า "มา" สำเร็จ ' + countMa + ' จุด', 'สถานะ');
  } else {
    ss.toast('⚠️ ไม่มีการเติมข้อมูล กรุณาเช็คว่าแถว 5 และคอลัมน์ C มีข้อมูลหรือไม่', 'คำเตือน');
  }
}