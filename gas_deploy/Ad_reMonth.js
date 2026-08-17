// Ad_reMonth.gs

/**
 * ฟังก์ชัน Ad_reMonth:
 * อ่านข้อมูลจากชีต 'rec' และนำมาอัปเดตข้อมูลสถานะในชีทเดือนปัจจุบัน
 * โดยจะกรองข้อมูลที่เกี่ยวข้องกับรหัสนักเรียนและวันที่ในชีทเดือนนั้นๆ
 *
 * (สมมติว่าฟังก์ชันนี้ถูกเรียกใช้เพื่อรีเฟรชข้อมูลในชีทเดือน)
 */
/*
function Ad_reMonth() {
  Logger.log('--- เริ่มการทำงาน Ad_reMonth (ปรับปรุงการอ่านจาก rec) ---');
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const activeSheet = ss.getActiveSheet();
  const activeSheetName = activeSheet.getName();
  // const ui = SpreadsheetApp.getUi(); // ลบการประกาศ ui object ออกไป

  try {
    const allowedMonthSheetNames = Array.from({ length: 12 }, (_, i) => String(i + 1).padStart(2, '0'));

    if (!allowedMonthSheetNames.includes(activeSheetName)) {
      // ไม่ใช้ ui.alert แต่ใช้ ss.toast แทน
      ss.toast('🚫 กรุณาทำงานบนชีตเดือน (เช่น 01, 02, ...) เพื่อดึงข้อมูล', 'ข้อผิดพลาด', 5);
      throw new Error('🚫 กรุณาทำงานบนชีตเดือน (เช่น 01, 02, ...) เพื่อดึงข้อมูล');
    }

    const recSheet = ss.getSheetByName(REC_SHEET_NAME);
    if (!recSheet) { 
      ss.toast(`ไม่พบชีต "${REC_SHEET_NAME}"`, 'ข้อผิดพลาด', 5); // เพิ่ม toast ตรงนี้
      throw new Error(`ไม่พบชีต "${REC_SHEET_NAME}"`); 
    }

    // --- ส่วนการกำหนดช่วงข้อมูลปัจจุบันในชีทเดือน ---
    const lastRowWithDataInC = getLastRowWithDataInColumn(activeSheet, CODE_COLUMN, STATUS_START_ROW, STATUS_MAX_SEARCH_ROW);
    if (lastRowWithDataInC < STATUS_START_ROW) {
      ss.toast('ℹ️ ไม่พบข้อมูลรหัสนักเรียนในคอลัมน์ C (ตั้งแต่ C' + STATUS_START_ROW + ' ถึง C' + STATUS_MAX_SEARCH_ROW + ') ไม่มีข้อมูลให้อัปเดต', 'สถานะ', 5);
      Logger.log('ไม่พบข้อมูลรหัสนักเรียนในคอลัมน์ C ในชีทเดือน');
      return;
    }
    const actualNumRowsToProcess = lastRowWithDataInC - STATUS_START_ROW + 1; // จำนวนแถวของข้อมูลสถานะจริง
    const numCols = STATUS_END_COL - STATUS_START_COL + 1; // จำนวนคอลัมน์ของข้อมูลสถานะ (K ถึง AO)

    // อ่านรหัสนักเรียนทั้งหมดจากชีทเดือน
    const studentsCodeColValues = activeSheet.getRange(STATUS_START_ROW, CODE_COLUMN, actualNumRowsToProcess, 1).getValues().flat();
    // อ่านวันที่จากแถวที่ 2 ของช่วงสถานะ (K2:AO2)
    const datesRow = activeSheet.getRange(DATE_ROW_FOR_STATUS_SHEET, STATUS_START_COL, 1, numCols).getValues().flat();

    Logger.log(`Processing range C${STATUS_START_ROW}:C${lastRowWithDataInC} and K${STATUS_START_ROW}:AO${lastRowWithDataInC}`);
    Logger.log(`Last row with student code in C: ${lastRowWithDataInC}`);
    Logger.log(`Actual rows to process: ${actualNumRowsToProcess}`);

    // --- ส่วนการอ่านข้อมูลจากชีต 'rec' ที่ปรับปรุงแล้ว ---
    // อ่านข้อมูลทั้งหมดจากชีต 'rec' ที่มีข้อมูลอยู่ (ตั้งแต่ A1 ไปจนถึงคอลัมน์สุดท้ายที่มีข้อมูล)
    // สมมติว่าข้อมูลที่เกี่ยวข้องอยู่ในคอลัมน์ A (StudentCode), B (Date), D (Status)
    // ตรวจสอบให้แน่ใจว่าค่าคงที่ REC_SHEET_LAST_COL_WITH_DATA ถูกกำหนดไว้อย่างถูกต้องใน Constants.gs
    
    // ตรวจสอบว่า recSheet มีข้อมูลหรือไม่
    if (recSheet.getLastRow() < 1) {
      ss.toast('ℹ️ ชีต "rec" ไม่มีข้อมูล ไม่มีอะไรให้อัปเดต', 'สถานะ', 5);
      Logger.log('ชีต "rec" ไม่มีข้อมูล');
      return;
    }

    const recDataRange = recSheet.getRange(1, 1, recSheet.getLastRow(), REC_SHEET_LAST_COL_WITH_DATA); 
    const allRecValues = recDataRange.getValues();
    Logger.log(`Finished reading ${allRecValues.length} rows from 'rec' sheet.`);

    // เตรียม Map สำหรับการเข้าถึงข้อมูล rec ได้อย่างรวดเร็ว
    // Key: `studentCode|formattedDate` Value: `status`
    const recDataMap = new Map();

    // วนลูปผ่านข้อมูลทั้งหมดจากชีต 'rec' เพื่อสร้าง Map
    // (สมมติว่า Col A คือ StudentCode, Col B คือ Date, Col D คือ Status)
    // ตรวจสอบ Index ของคอลัมน์ในชีต rec
    // Col A -> Index 0
    // Col B -> Index 1
    // Col D -> Index 3
    for (let i = 0; i < allRecValues.length; i++) {
        const row = allRecValues[i];
        const recStudentCode = cleanStatus(row[REC_COL_STUDENT_CODE_INDEX]); // A
        const recDate = row[REC_COL_DATE_INDEX]; // B
        const recStatus = cleanStatus(row[REC_COL_STATUS_INDEX]); // D

        if (recStudentCode && recDate instanceof Date && !isNaN(recDate.getTime()) && recStatus) {
            const formattedRecDate = Utilities.formatDate(recDate, ss.getSpreadsheetTimeZone(), "yyyy-MM-dd");
            const key = `${recStudentCode}|${formattedRecDate}`;
            recDataMap.set(key, recStatus);
        } else {
            Logger.log(`Skipping invalid rec record at row ${i + 1}: Code='${recStudentCode}', Date='${recDate}', Status='${recStatus}'`);
        }
    }
    Logger.log(`Created recDataMap with ${recDataMap.size} entries.`);

    // --- ส่วนการอัปเดตชีทเดือนปัจจุบัน ---
    const updatedStatuses = []; // Array สำหรับเก็บค่าสถานะใหม่ที่จะเขียนลงชีทเดือน
    let updatesCount = 0;

    for (let r = 0; r < actualNumRowsToProcess; r++) {
        const currentStudentCode = cleanStatus(studentsCodeColValues[r]);
        const rowStatuses = []; // สถานะของแถวปัจจุบันในชีทเดือน

        if (!currentStudentCode) {
            // ถ้าไม่มีรหัสนักเรียนในแถวนี้ ให้คงค่าเดิมไว้ (หรือเติมค่าว่าง)
            // หรือถ้าคุณต้องการให้ข้อมูลสถานะของนักเรียนที่ไม่มีรหัสถูกล้าง ให้ใช้ Array(numCols).fill('')
            const existingRow = activeSheet.getRange(STATUS_START_ROW + r, STATUS_START_COL, 1, numCols).getValues().flat();
            updatedStatuses.push(existingRow);
            continue;
        }

        for (let c = 0; c < numCols; c++) {
            const date = datesRow[c];
            let dateObj;
            try {
                if (date instanceof Date && !isNaN(date.getTime())) {
                    dateObj = date;
                } else if (typeof date === 'string' && date.trim() !== '') {
                    dateObj = Utilities.parseDate(date.trim(), ss.getSpreadsheetTimeZone(), 'yyyy-MM-dd');
                    if (isNaN(dateObj.getTime())) { 
                        dateObj = Utilities.parseDate(date.trim(), ss.getSpreadsheetTimeZone(), 'dd/MM/yyyy');
                    }
                    if (isNaN(dateObj.getTime())) {
                        dateObj = new Date(date.trim()); 
                    }
                } else if (typeof date === 'number') {
                    dateObj = new Date((date - (25569)) * 86400 * 1000);
                }
                
                if (!dateObj || isNaN(dateObj.getTime())) {
                  throw new Error('Invalid date value after parsing attempts');
                }
            } catch (e) {
                Logger.log(`Warning: Failed to parse date '${date}' from R${DATE_ROW_FOR_STATUS_SHEET}C${STATUS_START_COL + c}. Error: ${e.message}. Will keep existing status for this cell.`);
                // ใช้ค่าสถานะเดิมหากวันที่ไม่ถูกต้อง
                rowStatuses.push(activeSheet.getRange(STATUS_START_ROW + r, STATUS_START_COL + c).getValue());
                continue;
            }

            const formattedDate = Utilities.formatDate(dateObj, ss.getSpreadsheetTimeZone(), "yyyy-MM-dd");
            const key = `${currentStudentCode}|${formattedDate}`;
            const newStatus = recDataMap.get(key) || ''; // ดึงสถานะจาก Map, ถ้าไม่พบให้เป็นค่าว่าง

            const currentCellStatus = activeSheet.getRange(STATUS_START_ROW + r, STATUS_START_COL + c).getValue();
            const cleanedCurrentCellStatus = cleanStatus(currentCellStatus);

            if (newStatus !== cleanedCurrentCellStatus) {
                rowStatuses.push(newStatus);
                updatesCount++;
                Logger.log(`Found update for Student Code: ${currentStudentCode}, Date: ${formattedDate}. Changing status from '${cleanedCurrentCellStatus}' to '${newStatus}'`);
            } else {
                rowStatuses.push(currentCellStatus); // ถ้าไม่เปลี่ยน ให้คงค่าเดิมไว้ (รวมถึงรูปแบบเซลล์)
            }
        }
        updatedStatuses.push(rowStatuses);
    }

    // เขียนข้อมูลสถานะที่อัปเดตแล้วกลับลงในชีทเดือน
    if (updatesCount > 0) {
      activeSheet.getRange(STATUS_START_ROW, STATUS_START_COL, actualNumRowsToProcess, numCols).setValues(updatedStatuses);
      ss.toast(`✅ อัปเดตข้อมูลสถานะในชีทเดือน ${updatesCount} รายการเรียบร้อยแล้ว`, 'รีเฟรชสำเร็จ', 5);
      // ui.alert('รีเฟรชข้อมูลสำเร็จ', `อัปเดตข้อมูลสถานะในชีทเดือนปัจจุบัน ${updatesCount} รายการ`, ui.ButtonSet.OK); // ลบ Alert นี้
      Logger.log(`Finished updating ${updatesCount} statuses in active sheet.`);
    } else {
      ss.toast('ℹ️ ไม่มีข้อมูลสถานะใดๆ ที่ต้องอัปเดต', 'รีเฟรช', 3);
      Logger.log('ไม่มีข้อมูลสถานะใดๆ ที่ต้องอัปเดต');
    }

    SpreadsheetApp.flush();
    Logger.log('--- สิ้นสุดการทำงาน Ad_reMonth ---');

  } catch (err) {
    Logger.log(`Error in Ad_reMonth: ${err.message}. Stack: ${err.stack}`);
    // ui.alert('เกิดข้อผิดพลาด', `🚫 เกิดข้อผิดพลาดในการรีเฟรชข้อมูล: ${err.message}\nโปรดตรวจสอบ Log เพิ่มเติม`, ui.ButtonSet.OK); // ลบ Alert นี้
    ss.toast(`🚫 เกิดข้อผิดพลาดในการรีเฟรช: ${err.message}`, 'ข้อผิดพลาด', 8); // เพิ่ม Toast ตรงนี้
    throw new Error(`🚫 เกิดข้อผิดพลาดในการโหลดข้อมูล: ${err.message}`);
  }
}
*/