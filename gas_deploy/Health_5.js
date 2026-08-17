// -------------------------------------------------------------------------
// *** ตารางกำหนดช่วงเดือนสำหรับชีท "สุขภาพ" (Health Monthly Blocks) ***
// -------------------------------------------------------------------------

/**
 * โครงสร้างข้อมูลสำหรับแต่ละบล็อกเดือนในชีท "สุขภาพ" (ใช้สำหรับ Health_5)
 */
const HEALTH_MONTHLY_BLOCKS_5 = [
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

const CONDITION_COL_5 = 2; // คอลัมน์ B
// คอลัมน์ Checkbox ที่ต้องการตั้งค่าเป็น TRUE (C, H, M, R, W, AB, AG, AL, AQ)
const CHECKBOX_COLS_5 = [3, 8, 13, 18, 23, 28, 33, 38, 43]; 


/**
 * ตั้งค่า Checkbox ที่อยู่ในคอลัมน์ควบคุม (C, H, M, ...) 
 * ให้เป็น TRUE (ทำเครื่องหมายถูก) ของทุกแถวควบคุม และเฉพาะแถวที่มีข้อมูลในคอลัมน์ B ของแถวข้อมูล
 */
function Health_5() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getActiveSheet(); 
    const ui = SpreadsheetApp.getUi();
    
    if (sheet.getName().trim() !== "สุขภาพ") {
        ui.alert('❌ ข้อผิดพลาด', 'ฟังก์ชันนี้ใช้งานได้เฉพาะในแผ่นงาน "สุขภาพ" เท่านั้น', ui.ButtonSet.OK);
        return;
    }

    // *** เซลล์ต้นแบบ Checkbox TRUE (Checkbox ที่ AX8) ***
    const TEMPLATE_CELL_TRUE = sheet.getRange("AX8");
    
    let allRangesToPaste = []; // รวบรวมช่วง A1 Notation ทั้งหมดที่ต้องวาง TRUE

    try {
        
        // วนลูปทำงานสำหรับทุกบล็อกเดือน (11 เดือน)
        HEALTH_MONTHLY_BLOCKS_5.forEach(block => {
            const START_ROW = block.dataStartRow;
            const END_ROW = block.dataEndRow;
            const NUM_ROWS = END_ROW - START_ROW + 1;
            
            // แถวควบคุม (Row 10, 39, 68, ...)
            const CHECKBOX_ROW = START_ROW - 1; 

            // ** 1. เพิ่มแถวควบคุม (Unconditional) **
            CHECKBOX_COLS_5.forEach(colIndex => {
                // เพิ่มเซลล์ Checkbox ในแถวควบคุมเข้าไปใน List (เช่น C10, H10, C39, H39, ...)
                allRangesToPaste.push(sheet.getRange(CHECKBOX_ROW, colIndex).getA1Notation());
            });

            // ** 2. เพิ่มแถวข้อมูล (Conditional - กรอง B) **
            
            // ดึงข้อมูลจากคอลัมน์ B ในช่วงของเดือนปัจจุบัน
            const conditionValues = sheet.getRange(START_ROW, CONDITION_COL_5, NUM_ROWS, 1).getValues();

            // วนลูปตรวจสอบแต่ละแถวข้อมูล และสร้าง List ของเซลล์ที่ต้องการวาง TRUE
            for (let i = 0; i < conditionValues.length; i++) {
                
                // ตรวจสอบว่าคอลัมน์ B มีข้อมูล
                if (conditionValues[i][0] && String(conditionValues[i][0]).trim() !== '') {
                    const row = START_ROW + i;
                    
                    // วนลูปคอลัมน์ Checkbox ที่ต้องการ (C, H, M, ...)
                    CHECKBOX_COLS_5.forEach(colIndex => {
                        // เพิ่มเซลล์ Checkbox ในแถวข้อมูลเข้าไปใน List
                        allRangesToPaste.push(sheet.getRange(row, colIndex).getA1Notation());
                    });
                }
            }
        }); 
        
        // 3. ทำการคัดลอก/วาง Checkbox TRUE ทีเดียวทั้งหมด (Batch Processing)
        if (allRangesToPaste.length > 0) {
            const combinedRangeList = sheet.getRangeList(allRangesToPaste);

            combinedRangeList.getRanges().forEach(destinationCell => {
                
                // คัดลอกเฉพาะค่า (TRUE) และ Data Validation (Checkbox)
                TEMPLATE_CELL_TRUE.copyTo(destinationCell, SpreadsheetApp.CopyPasteType.PASTE_DATA_VALIDATION, false);
                TEMPLATE_CELL_TRUE.copyTo(destinationCell, SpreadsheetApp.CopyPasteType.PASTE_VALUES, false);
            });
            
            ui.alert('✅ ตั้งค่า Checkbox สำเร็จ!', `ตั้งค่า Checkbox ${allRangesToPaste.length} เซลล์ (รวมแถวควบคุมและแถวข้อมูลที่กรอง B) ของทั้ง 11 เดือนเป็น TRUE โดยใช้ Batch Processing เรียบร้อยแล้ว`, ui.ButtonSet.OK);

        } else {
            ui.alert('ℹ️ ไม่มีข้อมูล', 'ไม่พบแถวที่ต้องตั้งค่า Checkbox', ui.ButtonSet.OK);
        }

    } catch (error) {
        ui.alert('❌ ข้อผิดพลาด', `เกิดข้อผิดพลาดในการตั้งค่า Checkbox: ${error.message}`, ui.ButtonSet.OK);
    }
}