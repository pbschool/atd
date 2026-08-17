// ลำดับของแผ่นงานเดือนตามที่ผู้ใช้กำหนด (04 ถูกละไว้)
const MONTH_SHEET_SEQUENCE = ["05", "06", "07", "08", "09", "10", "11", "12", "01", "02", "03"];

/**
 * A universal function to switch to a specified sheet and stabilize the view
 * by forcing a quick switch-back-and-forth between adjacent sheets.
 * @param {string} sheetName The name of the target sheet.
 */
function goToSheet(sheetName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const targetSheet = ss.getSheetByName(sheetName);

  if (!targetSheet) {
    ss.toast(`⚠️ ไม่พบแผ่นงานชื่อ "${sheetName}"`, 'ข้อผิดพลาด', 5);
    return;
  }
  
  // 1. เปิดใช้งานแผ่นงานเป้าหมาย (ครั้งที่ 1)
  targetSheet.activate(); 

  let intermediateSheetName = null;
  const idx = MONTH_SHEET_SEQUENCE.indexOf(sheetName);

  if (idx !== -1) {
    // เป็นแผ่นงานเดือน
    if (sheetName === "03") {
      // กรณีพิเศษตามคำขอ: 03 -> 02 -> 03 (สลับไปเดือนก่อนหน้า)
      intermediateSheetName = "02";
    } else {
      // กรณีทั่วไป: N -> N+1 -> N (สลับไปเดือนถัดไป)
      const nextIdx = (idx + 1) % MONTH_SHEET_SEQUENCE.length;
      intermediateSheetName = MONTH_SHEET_SEQUENCE[nextIdx];
    }
  } else {
    // เป็นแผ่นงานสรุป/เมนู (ใช้ "05" เป็นแผ่นงานกลางที่เสถียร)
    intermediateSheetName = "05";
  }
  
  // 2. ดำเนินการสลับ: Target -> Intermediate -> Target
  const intermediateSheet = intermediateSheetName ? ss.getSheetByName(intermediateSheetName) : null;
  
  // ตรวจสอบว่าแผ่นงานกลางมีอยู่จริงและไม่ใช่แผ่นงานเป้าหมายเดียวกัน
  if (intermediateSheet && intermediateSheet.getName() !== sheetName) {
    intermediateSheet.activate(); // สลับไปแผ่นงานกลาง
    targetSheet.activate(); // สลับกลับมาแผ่นงานเป้าหมาย (เพื่อยืนยันการวาดซ้ำ)
  }
  
  ss.toast(`✅ ไปยังแผ่นงาน "${sheetName}" เรียบร้อยแล้ว`, 'สำเร็จ', 3);
}

// --- ฟังก์ชันสำหรับแต่ละปุ่ม (Functions for each button) ---

function btnMenu() {
  goToSheet("เมนู"); 
}

function btnLinkm05() {
  goToSheet("05");
}

function btnLinkm06() {
  goToSheet("06");
}

function btnLinkm07() {
  goToSheet("07");
}

function btnLinkm08() {
  goToSheet("08");
}

function btnLinkm09() {
  goToSheet("09");
}

function btnLinkm10() {
  goToSheet("10");
}

function btnLinkm11() {
  goToSheet("11");
}

function btnLinkm12() {
  goToSheet("12");
}

function btnLinkm01() {
  goToSheet("01");
}

function btnLinkm02() {
  goToSheet("02");
}

function btnLinkm03() {
  goToSheet("03");
}

function btnResultMa() {
  goToSheet("สรุปมา");
}

function btnWH() {
  goToSheet("นน.สส.");
}

function btnActivities() {
  goToSheet("สรุปกิจฯ");
}

function btnHealth() {
  goToSheet("สุขภาพ");
}