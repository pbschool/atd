/**
 * Utils.gs
 */

/**
 * Normalize cell value to a clean string.
 * - Converts null/undefined to ''
 * - Trims whitespace
 * - Removes ZERO_WIDTH_SPACE placeholders
 * - Keeps non-string primitives readable
 *
 * @param {any} value
 * @returns {string}
 */
function cleanStatus(value) {
  if (value === null || value === undefined) return "";
  let s = String(value);
  if (typeof ZERO_WIDTH_SPACE === "string" && ZERO_WIDTH_SPACE) {
    s = s.split(ZERO_WIDTH_SPACE).join("");
  }
  return s.trim();
}

/**
 * Find the last row index (1-based) with non-empty value in a specific column.
 *
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet
 * @param {number} column 1-based column index
 * @param {number} startRow 1-based row index
 * @param {number} maxRow 1-based row index (upper bound)
 * @returns {number} last row with data, or startRow-1 if none
 */
function getLastRowWithDataInColumn(sheet, column, startRow, maxRow) {
  const lastPossible = Math.min(maxRow, sheet.getMaxRows());
  if (lastPossible < startRow) return startRow - 1;

  const numRows = lastPossible - startRow + 1;
  const values = sheet.getRange(startRow, column, numRows, 1).getValues();

  let lastWithData = startRow - 1;
  for (let i = 0; i < values.length; i++) {
    const v = cleanStatus(values[i][0]);
    if (v !== "") lastWithData = startRow + i;
  }
  return lastWithData;
}

