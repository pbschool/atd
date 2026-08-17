// // saveChanges.gs

// /**
//  * ฟังก์ชัน saveChanges:
//  * ตรวจสอบการเปลี่ยนแปลงข้อมูลในชีตเดือนปัจจุบัน (K6:AO_last_row_with_data) เทียบกับข้อมูลสำรองใน PropertiesService
//  * - หากมีการเปลี่ยนแปลง จะบันทึกข้อมูลใหม่ลงในชีต 'rec' แบบ Append-Only พร้อม Timestamp ปัจจุบัน
//  * - อัปเดตข้อมูลสำรองใน PropertiesService ให้เป็นสถานะปัจจุบัน
//  * - แสดง Toast message แจ้งสถานะการบันทึก
//  * - หากไม่มีการเปลี่ยนแปลง จะไม่ดำเนินการใดๆ และแสดง Toast message แจ้งว่าไม่มีการเปลี่ยนแปลง
//  *
//  * *** ฟังก์ชันนี้จะถูกเรียกโดยตรงจาก Server-side หรือจาก Client-side (google.script.run) ***
//  */
// function saveChanges() {
//   Logger.log('--- เริ่มการทำงาน saveChanges (ไม่มี UI Modal) ---'); 
//   const ss = SpreadsheetApp.getActiveSpreadsheet();
//   const activeSheet = ss.getActiveSheet();
//   const activeSheetName = activeSheet.getName();

//   try {
//     const allowedMonthSheetNames = Array.from({ length: 12 }, (_, i) => String(i + 1).padStart(2, '0'));

//     if (!allowedMonthSheetNames.includes(activeSheetName)) {
//       ss.toast('🚫 กรุณาทำงานบนชีตเดือน (เช่น 01, 02, ...) เพื่อบันทึกข้อมูล', 'ข้อผิดพลาด', 5);
//       throw new Error('🚫 กรุณาทำงานบนชีตเดือน (เช่น 01, 02, ...) เพื่อบันทึกข้อมูล');
//     }

//     const recSheet = ss.getSheetByName(REC_SHEET_NAME);
    
//     if (!recSheet) { 
//       ss.toast(`ไม่พบชีต "${REC_SHEET_NAME}"`, 'ข้อผิดพลาด', 5);
//       throw new Error(`ไม่พบชีต "${REC_SHEET_NAME}"`); 
//     }

//     const lastRowWithDataInC = getLastRowWithDataInColumn(activeSheet, CODE_COLUMN, STATUS_START_ROW, STATUS_MAX_SEARCH_ROW);
//     if (lastRowWithDataInC < STATUS_START_ROW) {
//       ss.toast('ℹ️ ไม่พบข้อมูลรหัสนักเรียนในคอลัมน์ C (ตั้งแต่ C' + STATUS_START_ROW + ' ถึง C' + STATUS_MAX_SEARCH_ROW + ') ไม่มีอะไรต้องบันทึก', 'สถานะ', 5);
//       Logger.log('ไม่พบข้อมูลรหัสนักเรียนในคอลัมน์ C');
//       return 'no_data';
//     }
//     const actualNumRowsToProcess = lastRowWithDataInC - STATUS_START_ROW + 1;
//     const numCols = STATUS_END_COL - STATUS_START_COL + 1;

//     const currentStatusesRange = activeSheet.getRange(STATUS_START_ROW, STATUS_START_COL, actualNumRowsToProcess, numCols);
//     const currentStatuses = currentStatusesRange.getValues();

//     const studentsCodeColValues = activeSheet.getRange(STATUS_START_ROW, CODE_COLUMN, actualNumRowsToProcess, 1).getValues();
//     const studentLevelColValues = activeSheet.getRange(STATUS_START_ROW, LEVEL_COLUMN, actualNumRowsToProcess, 1).getValues();
//     const datesRow = activeSheet.getRange(DATE_ROW_FOR_STATUS_SHEET, STATUS_START_COL, 1, numCols).getValues().flat();

//     const scriptProperties = PropertiesService.getScriptProperties();
//     const backupKey = `backup_${activeSheetName}`;

//     let backupValues = [];
//     const storedBackupJson = scriptProperties.getProperty(backupKey);
//     Logger.log(`Attempting to load backup for '${backupKey}'. Stored JSON: ${storedBackupJson ? 'Found' : 'Not found/null'}`);

//     if (storedBackupJson) {
//       try {
//         const parsedBackup = JSON.parse(storedBackupJson);
//         if (Array.isArray(parsedBackup) && parsedBackup.length > 0 && Array.isArray(parsedBackup[0])) {
//             backupValues = parsedBackup;
//             Logger.log(`Loaded backup data from PropertiesService for sheet '${activeSheetName}'. Backup rows: ${backupValues.length}, cols: ${backupValues[0].length}`);
//         } else {
//             throw new Error('Parsed backup data is not in expected array of arrays format.');
//         }
//       } catch (parseError) {
//         Logger.log(`Error parsing/validating JSON from PropertiesService for sheet '${activeSheetName}': ${parseError.message}. Will use empty backup for comparison.`);
//       }
//     } else {
//         Logger.log('No existing backup found in PropertiesService. Will use empty backup for comparison.');
//     }

//     if (backupValues.length < actualNumRowsToProcess) {
//         Logger.log(`Backup data has ${backupValues.length} rows, but current sheet has ${actualNumRowsToProcess} rows. Extending backup.`);
//         for (let r = backupValues.length; r < actualNumRowsToProcess; r++) {
//             backupValues.push(Array(numCols).fill(''));
//         }
//     } else if (backupValues.length > actualNumRowsToProcess) {
//         Logger.log(`Backup data has ${backupValues.length} rows, but current sheet has ${actualNumRowsToProcess} rows. Truncating backup.`);
//         backupValues = backupValues.slice(0, actualNumRowsToProcess);
//     }

//     for (let r = 0; r < actualNumRowsToProcess; r++) {
//         if (!Array.isArray(backupValues[r])) {
//             backupValues[r] = [];
//         }
//         if (backupValues[r].length < numCols) {
//             Logger.log(`Row ${r} in backup has ${backupValues[r].length} columns, but current sheet has ${numCols} columns. Extending row.`);
//             for (let c = backupValues[r].length; c < numCols; c++) {
//                 backupValues[r].push('');
//             }
//         } else if (backupValues[r].length > numCols) {
//             Logger.log(`Row ${r} in backup has ${backupValues[r].length} columns, but current sheet has ${numCols} columns. Truncating columns.`);
//             backupValues[r] = backupValues[r].slice(0, numCols);
//         }
//     }

//     const rowsToAppendToRec = [];
//     const currentTime = new Date();

//     const cleanedCurrentStatusesForBackup = [];
//     let changesDetected = false;

//     for (let r = 0; r < actualNumRowsToProcess; r++) {
//       const studentCode = cleanStatus(studentsCodeColValues[r][0]);
//       const studentLevel = cleanStatus(studentLevelColValues[r][0]);

//       if (!studentCode) {
//         Logger.log(`Skipping row ${STATUS_START_ROW + r} due to empty student code. No record for rec sheet.`);
//         cleanedCurrentStatusesForBackup.push(Array(numCols).fill(''));
//         continue;
//       }

//       const rowForBackup = [];
//       for (let c = 0; c < numCols; c++) {
//         const currentStatus = currentStatuses[r][c];
//         const backupStatus = (backupValues[r] && typeof backupValues[r][c] !== 'undefined') ? backupValues[r][c] : '';

//         const finalCurrentStatus = cleanStatus(currentStatus);
//         const finalBackupStatus = cleanStatus(backupStatus);
        
//         Logger.log(`Row ${STATUS_START_ROW + r}, Col ${STATUS_START_COL + c} (Data): Current='${finalCurrentStatus}', Backup='${finalBackupStatus}'. Is different? ${finalCurrentStatus !== finalBackupStatus}`);

//         rowForBackup.push(finalCurrentStatus);

//         if (finalCurrentStatus !== finalBackupStatus) {
//           changesDetected = true;

//           const date = datesRow[c];
//           let dateObj;
//           try {
//               if (date instanceof Date && !isNaN(date.getTime())) {
//                   dateObj = date;
//               } else if (typeof date === 'string' && date.trim() !== '') {
//                   dateObj = Utilities.parseDate(date.trim(), ss.getSpreadsheetTimeZone(), 'yyyy-MM-dd');
//                   if (isNaN(dateObj.getTime())) { 
//                       dateObj = Utilities.parseDate(date.trim(), ss.getSpreadsheetTimeZone(), 'dd/MM/yyyy');
//                   }
//                   if (isNaN(dateObj.getTime())) {
//                       dateObj = new Date(date.trim());
//                   }
//               } else if (typeof date === 'number') {
//                   dateObj = new Date((date - (25569)) * 86400 * 1000);
//               }
              
//               if (!dateObj || isNaN(dateObj.getTime())) {
//                 throw new Error('Invalid date value after parsing attempts');
//               }
//           } catch (e) {
//               Logger.log(`Warning: Failed to parse date '${date}' from R${DATE_ROW_FOR_STATUS_SHEET}C${STATUS_START_COL + c}. Error: ${e.message}. Skipping this record for rec sheet.`);
//               continue;
//           }

//           const formattedDateForRecB = Utilities.formatDate(dateObj, ss.getSpreadsheetTimeZone(), "yyyy-MM-dd");
//           // *** ลบบรรทัด formattedDateForRecE เนื่องจากชีต rec ไม่มีคอลัมน์ E แล้ว ***
//           // const formattedDateForRecE = Utilities.formatDate(dateObj, ss.getSpreadsheetTimeZone(), "d/M/yyyy"); 

//           // ปรับปรุงการ push ข้อมูลให้เหลือ 5 คอลัมน์ (A, B, C, D, E)
//           // ลำดับข้อมูล: รหัสนักเรียน, วันที่มาเรียน (เป็น Date object), ระดับชั้น, สถานะมาเรียน, Timestamp การบันทึก
//           rowsToAppendToRec.push([
//             studentCode,
//             new Date(formattedDateForRecB), // คอลัมน์ B: วันที่ (Date Object)
//             studentLevel,                   // คอลัมน์ C: ระดับชั้น
//             finalCurrentStatus,             // คอลัมน์ D: สถานะมาเรียน
//             currentTime                     // คอลัมน์ E: Timestamp
//           ]);

//           Logger.log(`เตรียมบันทึก: Code=${studentCode}, Date(B)=${formattedDateForRecB}, Status=${finalCurrentStatus}, Timestamp(E)=${currentTime}`);
//         }
//       }
//       cleanedCurrentStatusesForBackup.push(rowForBackup);
//     }

//     if (!changesDetected) {
//         ss.toast('ℹ️ ไม่มีข้อมูลที่เปลี่ยนแปลง ไม่จำเป็นต้องบันทึก', 'สถานะ', 3);
//         Logger.log('ไม่มีการเปลี่ยนแปลงข้อมูล ไม่มีอะไรต้องบันทึก');
//         return 'no_changes';
//     }

//     if (rowsToAppendToRec.length > 0) {
//       recSheet.getRange(recSheet.getLastRow() + 1, 1, rowsToAppendToRec.length, rowsToAppendToRec[0].length).setValues(rowsToAppendToRec);
//       Logger.log(`บันทึกข้อมูลที่เปลี่ยนแปลง ${rowsToAppendToRec.length} แถวลงในชีต "${REC_SHEET_NAME}" สำเร็จ`);
      
//       ss.toast(`✅ บันทึก ${rowsToAppendToRec.length} รายการแล้ว`, 'เรียบร้อย', 5);
//     } else {
//         ss.toast('⚠️ ตรวจพบการเปลี่ยนแปลง แต่ไม่สามารถเตรียมรายการบันทึกได้ (ตรวจสอบการแปลงวันที่)', 'แจ้งเตือน', 5);
//         Logger.log('Changes detected, but rowsToAppendToRec is empty. Check date parsing or other filtering logic.');
//     }

//     if (cleanedCurrentStatusesForBackup && cleanedCurrentStatusesForBackup.length > 0) {
//         scriptProperties.setProperty(backupKey, JSON.stringify(cleanedCurrentStatusesForBackup));
//         Logger.log(`Updated backup data in PropertiesService for sheet '${activeSheetName}'.`);
//     } else {
//         Logger.log('Warning: cleanedCurrentStatusesForBackup is empty or invalid. PropertiesService not updated.');
//     }

//     SpreadsheetApp.flush();

//     Logger.log('--- สิ้นสุดการทำงาน saveChanges (ไม่มี UI Modal) ---'); 

//     return 'success';

//   } catch (err) {
//     Logger.log(`Error in saveChanges: ${err.message}. Stack: ${err.stack}`);
//     ss.toast(`🚫 เกิดข้อผิดพลาดในการบันทึก: ${err.message}`, 'ข้อผิดพลาด', 8); 
//     throw new Error(`🚫 เกิดข้อผิดพลาดในการบันทึกข้อมูล: ${err.message}`);
//   }
// }