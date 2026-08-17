function reloadRefErrorsInActiveSheet() {
  const sheet = SpreadsheetApp.getActiveSheet();
  const range = sheet.getRange("A1:BY25");
  const formulas = range.getFormulas();
  const values = range.getValues();

  for (let row = 0; row < values.length; row++) {
    for (let col = 0; col < values[0].length; col++) {
      const val = values[row][col];
      if (String(val).includes("#REF!")) {
        const formula = formulas[row][col];
        if (formula) {
          range.getCell(row + 1, col + 1).setFormula(formula);
        }
      }else{
          SpreadsheetApp.getActiveSpreadsheet().toast('✅ ข้อมูลในชีต "สรุปมา" ไม่มีการเปลี่ยนแปลง', 10);
          return;
        }
    }
  }
}
