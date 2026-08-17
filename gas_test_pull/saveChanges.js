// saveChanges.gs

/**
 * ฟังก์ชัน saveChanges:
 * จะบันทึกข้อมูลสถานะปัจจุบันจากชีตเดือน (K6:AO_last_row_with_data) ลงในชีท 'rec' แบบ Append-Only
 * โดยจะบันทึกเฉพาะเซลล์ที่มีข้อมูล (ไม่ว่างเปล่า) เท่านั้น พร้อม Timestamp ปัจจุบัน
 * และจะอัปเดตข้อมูลสำรองใน PropertiesService ให้เป็นสถานะปัจจุบันเสมอ
 * - แสดง Toast message แจ้งสถานะการบันทึก
 *
 * *** ฟังก์ชันนี้จะถูกเรียกโดยตรงจาก Server-side หรือจาก Client-side (google.script.run) ***
 */
function saveChanges() {
  Logger.log('--- เริ่มการทำงาน saveChanges (บันทึกเฉพาะเซลล์ที่ไม่ว่างเปล่า) ---'); 
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const activeSheet = ss.getActiveSheet();
  const activeSheetName = activeSheet.getName();

  try {
    const allowedMonthSheetNames = Array.from({ length: 12 }, (_, i) => String(i + 1).padStart(2, '0'));

    if (!allowedMonthSheetNames.includes(activeSheetName)) {
      ss.toast('🚫 กรุณาทำงานบนชีตเดือน (เช่น 01, 02, ...) เพื่อบันทึกข้อมูล', 'ข้อผิดพลาด', 5);
      throw new Error('🚫 กรุณาทำงานบนชีตเดือน (เช่น 01, 02, ...) เพื่อบันทึกข้อมูล');
    }

    const recSheet = ss.getSheetByName(REC_SHEET_NAME);
    
    if (!recSheet) { 
      ss.toast(`ไม่พบชีต "${REC_SHEET_NAME}"`, 'ข้อผิดพลาด', 5);
      throw new Error(`ไม่พบชีต "${REC_SHEET_NAME}"`); 
    }

    // กำหนดแถวสุดท้ายที่จะประมวลผลเป็น 26 ตามที่คุณระบุ
    // หรือใช้ STATUS_MAX_SEARCH_ROW ถ้าคุณต้องการให้ยืดหยุ่นกว่า
    const lastRowToProcess = Math.min(26, getLastRowWithDataInColumn(activeSheet, CODE_COLUMN, STATUS_START_ROW, STATUS_MAX_SEARCH_ROW)); 
    
    if (lastRowToProcess < STATUS_START_ROW) {
      ss.toast('ℹ️ ไม่พบข้อมูลรหัสนักเรียนในคอลัมน์ C (ตั้งแต่ C' + STATUS_START_ROW + ' ถึง C' + STATUS_MAX_SEARCH_ROW + ') ไม่มีอะไรต้องบันทึก', 'สถานะ', 5);
      Logger.log('ไม่พบข้อมูลรหัสนักเรียนในคอลัมน์ C');
      return 'no_data';
    }
    
    const actualNumRowsToProcess = lastRowToProcess - STATUS_START_ROW + 1;
    const numCols = STATUS_END_COL - STATUS_START_COL + 1;

    const currentStatusesRange = activeSheet.getRange(STATUS_START_ROW, STATUS_START_COL, actualNumRowsToProcess, numCols);
    const currentStatuses = currentStatusesRange.getValues();

    const studentsCodeColValues = activeSheet.getRange(STATUS_START_ROW, CODE_COLUMN, actualNumRowsToProcess, 1).getValues();
    const studentLevelColValues = activeSheet.getRange(STATUS_START_ROW, LEVEL_COLUMN, actualNumRowsToProcess, 1).getValues();
    const datesRow = activeSheet.getRange(DATE_ROW_FOR_STATUS_SHEET, STATUS_START_COL, 1, numCols).getValues().flat();

    const scriptProperties = PropertiesService.getScriptProperties();
    const backupKey = `backup_${activeSheetName}`;
    
    const rowsToAppendToRec = [];
    const currentTime = new Date();

    const cleanedCurrentStatusesForBackup = []; // ใช้เก็บสถานะปัจจุบันเพื่ออัปเดต backup

    for (let r = 0; r < actualNumRowsToProcess; r++) {
      const studentCode = cleanStatus(studentsCodeColValues[r][0]);
      const studentLevel = cleanStatus(studentLevelColValues[r][0]);

      if (!studentCode) {
        Logger.log(`Skipping row ${STATUS_START_ROW + r} due to empty student code. No record for rec sheet.`);
        cleanedCurrentStatusesForBackup.push(Array(numCols).fill('')); // ยังคงเพิ่มแถวว่างสำหรับ backup
        continue;
      }

      const rowForBackup = []; // เก็บสถานะของแต่ละคอลัมน์ในแถวปัจจุบันเพื่อใช้สำหรับ backup
      for (let c = 0; c < numCols; c++) {
        const currentStatus = currentStatuses[r][c];
        const finalCurrentStatus = cleanStatus(currentStatus);
        
        rowForBackup.push(finalCurrentStatus); // เพิ่มสถานะปัจจุบันลงในแถวสำหรับ backup

        // *** ตรวจสอบว่า finalCurrentStatus ไม่ว่างเปล่าก่อนบันทึก ***
        if (finalCurrentStatus === '') {
          continue; // ถ้าสถานะว่างเปล่า ให้ข้ามการบันทึกรายการนี้ลงใน recSheet
        }

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
                dateObj = new Date(date.trim()); // ลองแปลงแบบ Generic
            }
          } else if (typeof date === 'number') {
            dateObj = new Date((date - (25569)) * 86400 * 1000); // แปลง Serial Number เป็น Date
          }
          
          if (!dateObj || isNaN(dateObj.getTime())) {
            throw new Error('Invalid date value after parsing attempts');
          }
        } catch (e) {
          Logger.log(`Warning: Failed to parse date '${date}' from R${DATE_ROW_FOR_STATUS_SHEET}C${STATUS_START_COL + c}. Error: ${e.message}. Skipping this record for rec sheet.`);
          continue; // ข้ามการบันทึกข้อมูลนี้หากวันที่ผิดพลาด
        }

        const formattedDateForRecB = Utilities.formatDate(dateObj, ss.getSpreadsheetTimeZone(), "yyyy-MM-dd");
        
        // ปรับปรุงการ push ข้อมูลให้มี 5 คอลัมน์ (A, B, C, D, E) ตาม Constants.gs ล่าสุด
        // ลำดับข้อมูล: รหัสนักเรียน, วันที่มาเรียน (เป็น Date object), ระดับชั้น, สถานะมาเรียน, Timestamp การบันทึก
        rowsToAppendToRec.push([
          studentCode,
          new Date(formattedDateForRecB), // คอลัมน์ B: วันที่ (Date Object)
          studentLevel,                   // คอลัมน์ C: ระดับชั้น
          finalCurrentStatus,             // คอลัมน์ D: สถานะมาเรียน
          currentTime                     // คอลัมน์ E: Timestamp
        ]);

        Logger.log(`เตรียมบันทึก: Code=${studentCode}, Date(B)=${formattedDateForRecB}, Status=${finalCurrentStatus}, Timestamp(E)=${currentTime}`);
      }
      cleanedCurrentStatusesForBackup.push(rowForBackup); // เพิ่มแถวสถานะปัจจุบันทั้งหมดลงใน cleanedCurrentStatusesForBackup
    }

    if (rowsToAppendToRec.length > 0) {
      recSheet.getRange(recSheet.getLastRow() + 1, 1, rowsToAppendToRec.length, rowsToAppendToRec[0].length).setValues(rowsToAppendToRec);
      Logger.log(`บันทึกข้อมูลสถานะ ${rowsToAppendToRec.length} รายการลงในชีต "${REC_SHEET_NAME}" สำเร็จ`);
      
      ss.toast(`✅ บันทึก ${rowsToAppendToRec.length} รายการแล้ว`, 'เรียบร้อย', 5);
    } else {
        ss.toast('⚠️ ไม่พบรายการสถานะที่ไม่ว่างเปล่าเพื่อบันทึก', 'แจ้งเตือน', 5);
        Logger.log('No non-empty status records prepared for saving.');
    }

    // อัปเดต PropertiesService ด้วยสถานะปัจจุบันเสมอ
    if (cleanedCurrentStatusesForBackup && cleanedCurrentStatusesForBackup.length > 0) {
      scriptProperties.setProperty(backupKey, JSON.stringify(cleanedCurrentStatusesForBackup));
      Logger.log(`Updated backup data in PropertiesService for sheet '${activeSheetName}'.`);
    } else {
      Logger.log('Warning: cleanedCurrentStatusesForBackup is empty or invalid. PropertiesService not updated.');
    }

    SpreadsheetApp.flush(); // บังคับให้การเปลี่ยนแปลงในชีทถูกเขียนลงทันที

    Logger.log('--- สิ้นสุดการทำงาน saveChanges (บันทึกเฉพาะเซลล์ที่ไม่ว่างเปล่า) ---'); 

    return 'success';

  } catch (err) {
    Logger.log(`Error in saveChanges: ${err.message}. Stack: ${err.stack}`);
    ss.toast(`🚫 เกิดข้อผิดพลาดในการบันทึก: ${err.message}`, 'ข้อผิดพลาด', 8); 
    throw new Error(`🚫 เกิดข้อผิดพลาดในการบันทึกข้อมูล: ${err.message}`);
  }
}