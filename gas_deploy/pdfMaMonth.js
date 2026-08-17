// =========================================================================
// โค้ดสำหรับส่งออก PDF: เพิ่มระดับชั้น ('ชั้น'!B6) ในชื่อไฟล์ และย้ายไปถังขยะทันที
// =========================================================================

// ลำดับของแผ่นงานเดือนตามที่ผู้ใช้กำหนด
const ALL_MONTHS_SHEETS = ["05", "06", "07", "08", "09", "10", "11", "12", "01", "02", "03"];
const RANGE_TO_PRINT = 'A1:AO35'; 
const BASE_SLEEP_MS = 3000; // ฐานเวลาหน่วง 3 วินาที

/**
 * ดึงระดับชั้นจากชีต 'ชั้น' ที่เซลล์ B6
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} ss Spreadsheet Object
 * @returns {string} ระดับชั้น (e.g., "ป.1 ")
 */
function getClassLevel(ss) {
  const classSheet = ss.getSheetByName('ชั้น');
  if (!classSheet) {
    // หากไม่พบชีต 'ชั้น' ให้ส่งค่าว่างกลับไป
    return ''; 
  }
  // อ่านค่าจากเซลล์ B6
  const classLevel = classSheet.getRange('B6').getDisplayValue();
  // คืนค่าพร้อมเว้นวรรคท้าย เพื่อนำไปต่อกับชื่อไฟล์
  return classLevel ? `${classLevel} ` : ''; 
}

/**
 * กำหนดชื่อไฟล์ PDF ตามชื่อแผ่นงาน (เดือน)
 * @param {string} sheetName ชื่อแผ่นงาน (e.g., "05", "01")
 * @param {string} classLevel ระดับชั้นที่ดึงมา (e.g., "ป.1 ")
 * @returns {string} ชื่อไฟล์ PDF (e.g., "ป.1 มาเรียน 68 05.pdf")
 */
function getFileName(sheetName, classLevel) {
  const month = parseInt(sheetName, 10);
  let prefix;

  if (month >= 5 && month <= 12) {
    prefix = 'มาเรียน 68';
  } else if (month >= 1 && month <= 3) {
    prefix = 'มาเรียน 69';
  } else {
    prefix = 'รายงานเดือน';
  }
  
  // นำระดับชั้น (ClassLevel) มาต่อหน้าสุด
  return `${classLevel}${prefix} ${sheetName}.pdf`;
}

/**
 * ฟังก์ชันหลักในการนำข้อมูลจากแผ่นงานที่กำหนด
 * ไปสร้างเป็นไฟล์ PDF และเปิดในเบราว์เซอร์โดยตรง หรือเปิดโฟลเดอร์
 * @param {string[]} sheetNames Array ของชื่อแผ่นงานที่จะนำข้อมูลออก
 */
function exportSheetsToPDF(sheetNames) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();
  
  const isMultiSheetExport = sheetNames.length > 1;

  let sheets = []; 
  let openUrl;    // URL สุดท้ายที่จะเปิดในป๊อปอัพ

  try {
    // 1. ตรวจสอบและดึง Sheets
    sheets = sheetNames.map(name => ss.getSheetByName(name)).filter(sheet => sheet !== null);
    if (sheets.length === 0) {
      ss.toast(`⚠️ ไม่พบแผ่นงานใดๆ ตามรายการที่ระบุ`, 'ข้อผิดพลาด', 5);
      return;
    }

    // *** ดึงระดับชั้นจากชีต 'ชั้น' ***
    const classLevel = getClassLevel(ss);

    ui.showModalDialog(HtmlService.createHtmlOutput('<b>กำลังสร้างไฟล์ PDF โปรดรอสักครู่... (ประมวลผลทีละชีต)</b>'), 'กำลังดำเนินการ');

    // 2. กำหนด URL พื้นฐานและพารามิเตอร์การส่งออกหลัก
    const baseUrl = 'https://docs.google.com/spreadsheets/d/' + ss.getId() + '/export?';
    
    // พารามิเตอร์หลักที่ใช้ร่วมกัน
    let baseExportOptions = {
      format: 'pdf',
      size: 'A4',
      portrait: false,
      fitw: true,
      printtitle: false,
      gridlines: false,
      fzr: false,
      top_margin: 0.3937,
      bottom_margin: 0,
      left_margin: 0.3937,
      right_margin: 0.1969,
      
      // กำหนดขอบเขตการพิมพ์ A1:AO35 ด้วย Cell Coordinates (0-indexed)
      r1: 0,   
      c1: 0,   
      r2: 34,  
      c2: 40   
    };
    
    const token = ScriptApp.getOAuthToken();
    const headers = { 'Authorization': 'Bearer ' + token };

    if (isMultiSheetExport) {
        // *** 3. โหมดรวมหลายชีต: สร้างโฟลเดอร์และบันทึกทีละไฟล์ ***
        
        const FOLDER_NAME = `${classLevel}รายงานรวมหลายเดือน_${Utilities.formatDate(new Date(), ss.getSpreadsheetTimeZone(), "yyyyMMdd_HHmmss")}`;
        const tempFolder = DriveApp.createFolder(FOLDER_NAME);
        
        // ย้ายโฟลเดอร์ไปถังขยะทันที
        tempFolder.setTrashed(true);
        
        openUrl = tempFolder.getUrl(); // URL ที่จะใช้เปิด
        
        const MAX_RETRIES = 3;
        
        sheets.forEach((sheet) => {
            const sheetId = sheet.getSheetId();
            
            let exportOptions = { ...baseExportOptions, gid: sheetId };
            
            const params = Object.keys(exportOptions)
              .map(key => `${key}=${exportOptions[key]}`)
              .join('&');

            const finalUrl = baseUrl + params;
            
            let response;
            let retryCount = 0;
            
            // ใช้ Loop เพื่อพยายามเรียกซ้ำ (Retry Logic) หากเจอ 429
            do {
                try {
                    response = UrlFetchApp.fetch(finalUrl, { headers: headers });
                    break; 
                } catch (e) {
                    if (e.message.includes('429') && retryCount < MAX_RETRIES) {
                        retryCount++;
                        const delay = BASE_SLEEP_MS * retryCount; 
                        
                        ui.showModalDialog(HtmlService.createHtmlOutput(`<b>การเชื่อมต่อถูกจำกัด! กำลังลองใหม่ใน ${delay / 1000} วินาที... (ครั้งที่ ${retryCount}/${MAX_RETRIES})</b>`), 'กำลังดำเนินการ');
                        Utilities.sleep(delay);
                    } else {
                        throw e; 
                    }
                }
            } while (retryCount < MAX_RETRIES);

            if (!response) {
                throw new Error(`ไม่สามารถดึงข้อมูล PDF ของชีต ${sheet.getName()} ได้หลังจากลองใหม่ ${MAX_RETRIES} ครั้ง`);
            }

            // ใช้ฟังก์ชัน getFileName เพื่อตั้งชื่อไฟล์
            const sheetName = sheet.getName();
            const fileName = getFileName(sheetName, classLevel);
            const pdfBlob = response.getBlob().setName(fileName);
            tempFolder.createFile(pdfBlob); 
            
            Utilities.sleep(BASE_SLEEP_MS); // หน่วงเวลาก่อนไปชีตถัดไป

            ui.showModalDialog(HtmlService.createHtmlOutput(`<b>บันทึกไฟล์ PDF สำเร็จ: ${fileName}</b>`), 'กำลังดำเนินการ');
        });
        
    } else {
        // 3. โหมดชีตเดียว: ใช้ URL Export ปกติ
        
        const sheetName = sheets[0].getName();
        let exportOptions = { ...baseExportOptions, gid: sheets[0].getSheetId() };
        
        const params = Object.keys(exportOptions)
          .map(key => `${key}=${exportOptions[key]}`)
          .join('&');

        const finalUrl = baseUrl + params;
        const response = UrlFetchApp.fetch(finalUrl, { headers: headers });
        
        // ใช้ฟังก์ชัน getFileName เพื่อตั้งชื่อไฟล์
        const fileName = getFileName(sheetName, classLevel);
        
        // บันทึกไฟล์ PDF และย้ายไปถังขยะทันที
        const pdfFile = DriveApp.createFile(response.getBlob().setName(fileName));
        pdfFile.setTrashed(true);
        
        openUrl = pdfFile.getUrl(); 
    }

    // 4. สร้าง HTML เพื่อเปิด URL (ไฟล์เดียวหรือโฟลเดอร์)
    const htmlOutput = HtmlService
      .createHtmlOutput(`<script>window.open('${openUrl}', '_blank'); google.script.host.close();</script>`)
      .setWidth(100)
      .setHeight(1);
    
    ui.showModalDialog(htmlOutput, 'กำลังเปิดไฟล์/โฟลเดอร์...');
    ss.toast(`✅ สร้างไฟล์ PDF (หรือโฟลเดอร์) สำเร็จแล้ว (อยู่ในถังขยะ)`, 'สำเร็จ', 5);

  } catch (error) {
    ui.alert(`❌ เกิดข้อผิดพลาดในการสร้าง PDF: ${error.message}\n\nStack:\n${error.stack}`);
  } finally {
    // ไม่มีโค้ดลบไฟล์ใน finally เพื่อแก้ปัญหา Timing
  }
}

// --------------------------------------------------------------------------
// ฟังก์ชันสำหรับผูกกับปุ่ม (Button Functions) (ส่วนนี้ไม่ต้องเปลี่ยนแปลง)
// --------------------------------------------------------------------------

function exPDFm05() { exportSheetsToPDF(["05"]); }
function exPDFm06() { exportSheetsToPDF(["06"]); }
function exPDFm07() { exportSheetsToPDF(["07"]); }
function exPDFm08() { exportSheetsToPDF(["08"]); }
function exPDFm09() { exportSheetsToPDF(["09"]); }
function exPDFm10() { exportSheetsToPDF(["10"]); }
function exPDFm11() { exportSheetsToPDF(["11"]); }
function exPDFm12() { exportSheetsToPDF(["12"]); }
function exPDFm01() { exportSheetsToPDF(["01"]); }
function exPDFm02() { exportSheetsToPDF(["02"]); }
function exPDFm03() { exportSheetsToPDF(["03"]); }

function exPDFmAll() {
  exportSheetsToPDF(ALL_MONTHS_SHEETS);
}