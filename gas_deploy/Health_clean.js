// -------------------------------------------------------------------------
// *** ตารางกำหนดช่วงเดือนสำหรับชีท "สุขภาพ" (Health Monthly Blocks) ***
// -------------------------------------------------------------------------

/**
 * โครงสร้างข้อมูลสำหรับแต่ละบล็อกเดือนในชีท "สุขภาพ" (ใช้สำหรับ Health_clean)
 */
const HEALTH_MONTHLY_BLOCKS_clean = [
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

// แก้ไขชื่อตัวแปรเพื่อป้องกันการชนกับโค้ดอื่น:
const START_COL_clean = 3;   // คอลัมน์ C
const END_COL_CAU_clean = 47;  // คอลัมน์ AU (คอลัมน์ที่ 47)
const START_COL_BG_clean = 59; // คอลัมน์ BG (คอลัมน์ที่ 59)
const END_COL_BK_clean = 63;  // คอลัมน์ BK (คอลัมน์ที่ 63)
const CONDITION_COL_clean = 2; // คอลัมน์ B


/**
 * ตรวจสอบแถวที่มีข้อมูลในคอลัมน์ B และคัดลอก Checkbox FALSE จาก AW8 
 * ไปวางในช่วง C:AU และ BG:BK เพื่อทำการล้างค่าในแถวควบคุมและแถวข้อมูล (Batch Processing)
 */
function Health_clean() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getActiveSheet(); 
    const ui = SpreadsheetApp.getUi();
    
    if (sheet.getName().trim() !== "สุขภาพ") {
        ui.alert('❌ ข้อผิดพลาด', 'ฟังก์ชันนี้ใช้งานได้เฉพาะในแผ่นงาน "สุขภาพ" เท่านั้น', ui.ButtonSet.OK);
        return;
    }

    // *** เซลล์ต้นแบบ Checkbox FALSE (Checkbox ที่ AW8) ***
    const TEMPLATE_CELL_FALSE = sheet.getRange("AW8");
    
    let allRangesToClear = []; // รวบรวมช่วง A1 Notation ทั้งหมดที่ต้องวางค่า FALSE/ว่างเปล่า

    try {
        
        // วนลูปทำงานสำหรับทุกบล็อกเดือน (11 เดือน)
        HEALTH_MONTHLY_BLOCKS_clean.forEach(block => {
            const START_ROW = block.dataStartRow;
            const END_ROW = block.dataEndRow;
            const NUM_ROWS = END_ROW - START_ROW + 1;
            const NUM_COLS_CAU = END_COL_CAU_clean - START_COL_clean + 1;
            const NUM_COLS_BGBK = END_COL_BK_clean - START_COL_BG_clean + 1;
            
            // แถวควบคุม (Row 10, 39, 68, ...)
            const CHECKBOX_ROW = START_ROW - 1; 

            // ** 1. เพิ่มแถวควบคุม (Row 10/39/...) **
            // C:AU: แถวควบคุมต้องล้าง/รีเซ็ตค่าเสมอ
            const checkboxRowCAURangeA1 = sheet.getRange(CHECKBOX_ROW, START_COL_clean, 1, NUM_COLS_CAU).getA1Notation();
            allRangesToClear.push(checkboxRowCAURangeA1);
            
            // BG:BK: ไม่ต้องเพิ่มแถวควบคุม BG:BK เพราะต้องการคงหัวข้อไว้

            // ** 2. เพิ่มแถวข้อมูล (Conditional - กรอง B) **
            
            // ดึงข้อมูลจากคอลัมน์ B ในช่วงของเดือนปัจจุบัน
            const conditionValues = sheet.getRange(START_ROW, CONDITION_COL_clean, NUM_ROWS, 1).getValues();

            // วนลูปตรวจสอบแต่ละแถวข้อมูล และสร้าง List ของช่วงที่ต้องการล้างค่า
            for (let i = 0; i < conditionValues.length; i++) {
                
                // ตรวจสอบว่าคอลัมน์ B มีข้อมูล
                if (conditionValues[i][0] && String(conditionValues[i][0]).trim() !== '') {
                    const row = START_ROW + i;
                    
                    // A. เพิ่มช่วง Checkbox ของแถวนั้นๆ (C:AU)
                    const targetCAURangeA1 = sheet.getRange(row, START_COL_clean, 1, NUM_COLS_CAU).getA1Notation();
                    allRangesToClear.push(targetCAURangeA1);

                    // B. เพิ่มช่วง Checkbox ของแถวนั้นๆ (BG:BK)
                    const targetBGBKRangeA1 = sheet.getRange(row, START_COL_BG_clean, 1, NUM_COLS_BGBK).getA1Notation();
                    allRangesToClear.push(targetBGBKRangeA1);
                }
            }
        }); 
        
        // 3. ทำการคัดลอก/วางค่า FALSE/ว่างเปล่า ทีเดียวทั้งหมด (Batch Processing)
        if (allRangesToClear.length > 0) {
            const combinedRangeList = sheet.getRangeList(allRangesToClear);
            let rangesCleanedCount = 0;

            combinedRangeList.getRanges().forEach(destinationRange => {
                
                // คัดลอกเฉพาะค่า (FALSE หรือค่าว่าง) จาก AW8 เพื่อรีเซ็ต Checkbox
                TEMPLATE_CELL_FALSE.copyTo(destinationRange, SpreadsheetApp.CopyPasteType.PASTE_VALUES, false);
                rangesCleanedCount++;
            });
            
            ui.alert('✅ ล้างค่า Checkbox สำเร็จ!', `รีเซ็ต Checkbox ใน ${rangesCleanedCount} ช่วง (C:AU และ BG:BK) ของทั้ง 11 เดือนให้เป็น FALSE/ว่างเปล่า โดยใช้ Batch Processing เรียบร้อยแล้ว`, ui.ButtonSet.OK);

        } else {
            ui.alert('ℹ️ ไม่มีข้อมูล', 'ไม่พบแถวที่ต้องล้างค่า Checkbox', ui.ButtonSet.OK);
        }

    } catch (error) {
        ui.alert('❌ ข้อผิดพลาด', `เกิดข้อผิดพลาดในการล้างค่า Checkbox: ${error.message}`, ui.ButtonSet.OK);
    }
}