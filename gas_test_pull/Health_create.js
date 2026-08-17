// -------------------------------------------------------------------------
// *** ตารางกำหนดช่วงเดือนสำหรับชีท "สุขภาพ" (Health Monthly Blocks) ***
// -------------------------------------------------------------------------

/**
 * โครงสร้างข้อมูลสำหรับแต่ละบล็อกเดือนในชีท "สุขภาพ"
 */
const HEALTH_MONTHLY_BLOCKS_CREATE = [
    { name: 'พฤษภาคม', dataStartRow: 11, dataEndRow: 30 },
    { name: 'มิถุนายน', dataStartRow: 40, dataEndRow: 59 },
    { name: 'กรกฎาคม', dataStartRow: 69, dataEndRow: 88 },
    { name: 'สิงหาคม', dataStartRow: 98, dataEndRow: 117 },
    { name: 'กันยายน', dataStartRow: 127, dataEndRow: 146 },
    { name: 'ตุลาคม', dataStartRow: 156, dataEndRow: 175 },
    { name: 'พฤศจิกายน', dataStartRow: 185, dataEndRow: 204 },
    { name: 'ธันวาคม', dataStartRow: 214, dataEndRow: 233 },
    { name: 'มกราคม', dataStartRow: 243, dataEndRow: 262 },
    { name: 'กุมภาพันธ์', dataStartRow: 272, dataEndRow: 291 },
    { name: 'มีนาคม', dataStartRow: 301, dataEndRow: 320 } 
];

// แก้ไขชื่อตัวแปรเพื่อป้องกันการชนกัน (SyntaxError)
const START_COL_CREATE = 3;   // คอลัมน์ C
const END_COL_CAU_CREATE = 47; // คอลัมน์ AU (คอลัมน์ที่ 47)
const START_COL_BG_CREATE = 59; // คอลัมน์ BG (คอลัมน์ที่ 59)
const END_COL_BK_CREATE = 63;  // คอลัมน์ BK (คอลัมน์ที่ 63)
const CONDITION_COL_CREATE = 2; // คอลัมน์ B


/**
 * ตรวจสอบแถวที่มีข้อมูลในคอลัมน์ B และคัดลอก Checkbox จาก AW8 
 * ไปวางในช่วง C:AU และ BG:BK สำหรับแถวเหล่านั้น ในทุกเดือน โดยใช้ Batch Processing
 */
function Health_create() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getActiveSheet(); 
    const ui = SpreadsheetApp.getUi();
    
    if (sheet.getName().trim() !== "สุขภาพ") {
        ui.alert('❌ ข้อผิดพลาด', 'ฟังก์ชันนี้ใช้งานได้เฉพาะในแผ่นงาน "สุขภาพ" เท่านั้น', ui.ButtonSet.OK);
        return;
    }
    
    // *** เซลล์ต้นแบบ Checkbox เปล่า (Checkbox ที่ AW8) ***
    const TEMPLATE_CELL = sheet.getRange("AW8");
    
    let allRangesToClear = [];      // รวบรวมช่วงทั้งหมดเพื่อ Clear (C:AU ทั้งหมด + BG:BK ข้อมูล)
    let allRangesToPasteFromAW8 = []; // รวบรวมช่วง A1 Notation ทั้งหมดที่ต้องวาง Checkbox

    try {
        // --- ส่วนที่ 1: รวบรวมช่วงที่ต้องทำงานทั้งหมด ---
        HEALTH_MONTHLY_BLOCKS_CREATE.forEach(block => {
            const START_ROW = block.dataStartRow;
            const END_ROW = block.dataEndRow;
            const NUM_ROWS = END_ROW - START_ROW + 1;
            const NUM_COLS_CAU = END_COL_CAU_CREATE - START_COL_CREATE + 1;
            const NUM_COLS_BG_BK = END_COL_BK_CREATE - START_COL_BG_CREATE + 1;
            
            // แถว Checkbox ควบคุม (Row 10, 39, 68, ...)
            const CHECKBOX_ROW = START_ROW - 1;
            
            // 1.1 รวบรวมช่วงสำหรับ CLEAR 
            
            // - C:AU: ช่วง C:AU สำหรับ Clear (แถวควบคุม + แถวข้อมูลทั้งหมด)
            const fullCAUBlockRangeA1 = sheet.getRange(CHECKBOX_ROW, START_COL_CREATE, NUM_ROWS + 1, NUM_COLS_CAU).getA1Notation();
            allRangesToClear.push(fullCAUBlockRangeA1);
            
            // *** การแก้ไข: BG:BK Clear เฉพาะแถวข้อมูลเท่านั้น (START_ROW ถึง END_ROW) ***
            // นับจำนวนแถวจาก START_ROW
            const dataOnlyBGBKRangeA1 = sheet.getRange(START_ROW, START_COL_BG_CREATE, NUM_ROWS, NUM_COLS_BG_BK).getA1Notation();
            allRangesToClear.push(dataOnlyBGBKRangeA1);
            
            
            // 1.2 รวบรวมช่วงสำหรับ PASTE (AW8 - Checkbox เปล่า)
            
            // *** A. แถวควบคุม (Row 10, 39, ...): วาง Checkbox เฉพาะช่วง C:AU ***
            const checkboxRowCAURangeA1 = sheet.getRange(CHECKBOX_ROW, START_COL_CREATE, 1, NUM_COLS_CAU).getA1Notation();
            allRangesToPasteFromAW8.push(checkboxRowCAURangeA1); 
            // *** BG:BK ในแถวควบคุม (CHECKBOX_ROW) จะถูกเว้นไว้ ไม่มีการ Paste ***
            
            
            // B. จัดการแถวข้อมูล (Row 11-30, 40-59, ...) (กรอง B)
            const conditionValues = sheet.getRange(START_ROW, CONDITION_COL_CREATE, NUM_ROWS, 1).getValues();

            for (let i = 0; i < NUM_ROWS; i++) {
                // ตรวจสอบว่าคอลัมน์ B มีข้อมูล
                if (conditionValues[i][0] && String(conditionValues[i][0]).trim() !== '') {
                    const row = START_ROW + i;
                    
                    // แถวที่มีข้อมูลใน B: เตรียมสำหรับคัดลอก Checkbox
                    
                    // - ช่วง C:AU (วางเต็มช่วง)
                    const targetCAURangeA1 = sheet.getRange(row, START_COL_CREATE, 1, NUM_COLS_CAU).getA1Notation();
                    allRangesToPasteFromAW8.push(targetCAURangeA1);
                    
                    // - ช่วง BG:BK (วางเต็มช่วง)
                    const targetBGBKRangeA1 = sheet.getRange(row, START_COL_BG_CREATE, 1, NUM_COLS_BG_BK).getA1Notation();
                    allRangesToPasteFromAW8.push(targetBGBKRangeA1);
                }
            }
        }); 
        
        // --- ส่วนที่ 2: ดำเนินการ Batch Processing ---
        
        if (allRangesToClear.length === 0) {
            ui.alert('ℹ️ ไม่มีข้อมูล', `ไม่พบช่วงเดือนที่กำหนด`, ui.ButtonSet.OK);
            return;
        }

        // 2.1 ทำการ Clear Data Validation และค่า ของทุกช่วงในครั้งเดียว
        sheet.getRangeList(allRangesToClear).getRanges().forEach(range => {
            range.clearDataValidations();
            range.clearContent();
        });
        
        // 2.2 ทำการ Paste Checkbox และค่า (FALSE) ของทุกช่วงที่คัดสรรแล้วในครั้งเดียว
        if (allRangesToPasteFromAW8.length > 0) {
            const combinedPasteRangeList = sheet.getRangeList(allRangesToPasteFromAW8);

            combinedPasteRangeList.getRanges().forEach(destinationRange => {
                // คัดลอกเฉพาะ Data Validation (Checkbox)
                TEMPLATE_CELL.copyTo(destinationRange, SpreadsheetApp.CopyPasteType.PASTE_DATA_VALIDATION, false);
                
                // คัดลอกค่า (FALSE)
                TEMPLATE_CELL.copyTo(destinationRange, SpreadsheetApp.CopyPasteType.PASTE_VALUES, false);
            });
        }
        
        // --- ส่วนที่ 3: สรุปผล ---
        const totalRangesProcessed = allRangesToPasteFromAW8.length;
        const controlRowsCount = HEALTH_MONTHLY_BLOCKS_CREATE.length;
        const dataRowsRangesCount = totalRangesProcessed - controlRowsCount;
        const totalDataRowsProcessed = dataRowsRangesCount / 2;
        
        if (totalRangesProcessed > 0) {
            ui.alert('✅ สร้าง Checkbox สำเร็จ!', 
                `ดำเนินการสร้าง Checkbox ที่ถูกต้องเรียบร้อยแล้ว:\n` +
                ` - แถวควบคุม (C:AU): ${controlRowsCount} แถว\n` +
                ` - แถวข้อมูล (C:AU และ BG:BK กรอง B): ${totalDataRowsProcessed} แถว\n` +
                `ใช้ Batch Processing อย่างรวดเร็วแล้ว`, ui.ButtonSet.OK);
        } else {
            ui.alert('ℹ️ ไม่มีข้อมูล', `ไม่พบแถวข้อมูลหรือแถวควบคุมที่ต้องสร้าง Checkbox`, ui.ButtonSet.OK);
        }

    } catch (error) {
        ui.alert('❌ ข้อผิดพลาด', `เกิดข้อผิดพลาดในการสร้าง Checkbox: ${error.message}`, ui.ButtonSet.OK);
    }
}