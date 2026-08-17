/**
 * onEdit.gs - ฉบับสมบูรณ์ที่สุด
 * รวมระบบ: เช็คชื่อรายเดือน (แถว 6) + ระบบสุขภาพ (Full) + ระบบอนุบาล
 * อ้างอิงตัวแปรจาก Constants.gs ล่าสุดของคุณครู
 */

function onEdit(e) {
  if (!e || !e.range) return;
  
  const range = e.range;
  const sheet = range.getSheet();
  const ss = sheet.getParent(); 
  const editedRow = range.getRow();
  const editedCol = range.getColumn();
  const sheetName = sheet.getName().trim();
  const newValue = e.value ? String(e.value).toUpperCase() : ""; 

  // --- [1] กลุ่มชีทรายเดือน (05, 06, 07, ...) ---
  if (MONTHLY_SHEETS.includes(sheetName)) {
    // ดักจับแถว 6 (ตัวควบคุมมา/ล่องหน) ช่วงคอลัมน์ K(11) ถึง AO(41)
    if (editedRow === 6 && (editedCol >= R_START_COL && editedCol <= R_END_COL)) {
      const numRows = R_END_ROW - R_START_ROW + 1; 
      const names = sheet.getRange(R_START_ROW, R_NAME_COL, numRows, 1).getValues();
      const targetRange = sheet.getRange(R_START_ROW, editedCol, numRows, 1);

      if (newValue === 'TRUE') {
        const results = names.map(r => [(r[0] && String(r[0]).trim() !== "") ? "มา" : ""]);
        targetRange.setValues(results);
        ss.toast("✅ บันทึก 'มา' (แถว 6)", "ระบบเช็คชื่อ");
      } 
      else if (newValue === 'FALSE' || newValue === "") {
        const results = names.map(r => [(r[0] && String(r[0]).trim() !== "") ? ZERO_WIDTH_SPACE : ""]);
        targetRange.setValues(results);
        ss.toast("🗑️ ล้างค่าล่องหน (แถว 6)", "ระบบเช็คชื่อ");
      }
    }
  }

  // --- [2] กลุ่มชีทสุขภาพ ---
  else if (sheetName === "สุขภาพ") {
    try {
      Health_Logic_Main(sheet, ss, editedRow, editedCol, e.value);
    } catch(err) {
      Logger.log("Health Error: " + err.message);
    }
  }

  // --- [3] กลุ่มชีทประเมินคุณธรรม อนุบาล (ถ้ามี) ---
  else if (sheetName === "ประเมินคุณธรรม อนุบาล") {
    if (typeof setCheckboxValues === "function") {
       setCheckboxValues(e); 
    }
  }
}

/**
 * =========================================================================
 * === ฟังก์ชันระบบสุขภาพ (อ้างอิงชื่อตัวแปรจาก Constants.gs คุณครูเป๊ะๆ) ===
 * =========================================================================
 */
function Health_Logic_Main(sheet, ss, editedRow, editedCol, val) {
  const newValue = val ? String(val).toUpperCase() : "";

  // 1. Checkbox ควบคุมการเลือกเดือน (คอลัมน์ C)
  const monthBlock = HEALTH_TAB_CONFIG.find(b => (editedRow === b.controlRow && editedCol === 3) || (editedRow === b.controlRow + 1 && editedCol === 3));
  if (editedCol === 3 && monthBlock && newValue !== 'FALSE') {
    const isC4Type = (editedRow === monthBlock.controlRow);
    const cond = sheet.getRange(monthBlock.dataStartRow, 2, monthBlock.dataEndRow - monthBlock.dataStartRow + 1, 1).getValues();
    let ranges = [];
    let template = sheet.getRange(isC4Type ? 'AW8' : 'AX8');
    
    if (isC4Type) {
      ranges.push(sheet.getRange(monthBlock.checkboxRow, 3, 1, 45).getA1Notation());
      for (let i = 0; i < cond.length; i++) if (cond[i][0]) ranges.push(sheet.getRange(monthBlock.dataStartRow + i, 3, 1, 45).getA1Notation());
    } else {
      const letters = ['C', 'H', 'M', 'R', 'W', 'AB', 'AG', 'AL', 'AQ'];
      letters.forEach(col => ranges.push(col + monthBlock.checkboxRow));
      for (let i = 0; i < cond.length; i++) if (cond[i][0]) letters.forEach(col => ranges.push(col + (monthBlock.dataStartRow + i)));
    }
    batchCopy(sheet, template, ranges);
    sheet.getRange(editedRow, editedCol).setValue(false);
    return;
  }

  // 2. Checkbox หัวตาราง (แถว 10, 39, 68...)
  const headerBlock = HEALTH_TAB_CONFIG.find(b => editedRow === b.checkboxRow);
  if (headerBlock && editedCol >= 3 && editedCol <= 47) {
    const cond = sheet.getRange(headerBlock.dataStartRow, 2, headerBlock.dataEndRow - headerBlock.dataStartRow + 1, 1).getValues();
    let rs = [];
    for (let i = 0; i < cond.length; i++) if (cond[i][0]) rs.push(sheet.getRange(headerBlock.dataStartRow + i, editedCol).getA1Notation());
    batchCopy(sheet, sheet.getRange(editedRow, editedCol), rs);
    return;
  }

  // 3. Checkbox ควบคุมชุดคอลัมน์ (สุขภาพรายบุคคล)
  const targetCols = HEALTH_COL_MAPPING[editedCol];
  const activeBlock = HEALTH_TAB_CONFIG.find(b => editedRow >= b.dataStartRow && editedRow <= b.dataEndRow);
  if (targetCols && activeBlock) {
    // ถ้าไม่มีชื่อเด็ก (คอลัมน์ B ว่าง) ให้ดีด Checkbox ออก
    if (!sheet.getRange(editedRow, 2).getValue()) {
      sheet.getRange(editedRow, editedCol).setValue(newValue === 'TRUE' ? false : true);
      return;
    }
    let rs = [];
    targetCols.forEach(c => rs.push(sheet.getRange(editedRow, c).getA1Notation()));
    batchCopy(sheet, sheet.getRange(newValue === 'TRUE' ? 'AX8' : 'AW8'), rs);
  }
}

/**
 * ฟังก์ชันช่วย Copy ค่าแบบกลุ่ม
 */
function batchCopy(sheet, temp, ranges) {
  if (ranges.length === 0) return;
  const list = sheet.getRangeList(ranges);
  list.getRanges().forEach(r => temp.copyTo(r, SpreadsheetApp.CopyPasteType.PASTE_VALUES, false));
}