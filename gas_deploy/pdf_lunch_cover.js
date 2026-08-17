// /**
//  * ฟังก์ชันที่จะทำงานอัตโนมัติเมื่อเปิด Google Sheet เพื่อสร้างเมนู
//  */
// function onOpen() {
//   SpreadsheetApp.getUi()
//       .createMenu('เมนู PDF')
//       // เรียกใช้ฟังก์ชันหลัก
//       .addItem('เปิด PDF (A1:Z34, ขอบกำหนดเอง)', 'pdf_lunch_cover') 
//       .addToUi();
// }


// /**
//  * ชื่อฟังก์ชัน: pdf_lunch_cover
//  * วัตถุประสงค์: สร้าง PDF จาก Pอาหาร1/A1:Z34, บันทึกชั่วคราวลง Drive, 
//  * เปิดในเบราว์เซอร์ทันที, และย้ายไฟล์ไปถังขยะ
//  * **ใช้การตั้งค่าขอบกำหนดเองตามที่ผู้ใช้ระบุ**
//  */
// function pdf_lunch_cover() {
//   const sheetName = "Pอาหาร1";
//   const printRange = "A1:Z34"; 
//   const pdfFileName = "อาหารกลางวัน_ปก_กำหนดขอบเอง"; 

//   const ss = SpreadsheetApp.getActiveSpreadsheet();
//   const sheet = ss.getSheetByName(sheetName);
//   const ui = SpreadsheetApp.getUi();

//   if (!sheet) {
//     ui.alert(`ไม่พบแผ่นงานชื่อ "${sheetName}" กรุณาตรวจสอบชื่อแผ่นงานอีกครั้ง`);
//     return;
//   }

//   const ssId = ss.getId();
//   const sheetId = sheet.getSheetId();

//   // 1. ตั้งค่า URL สำหรับ Export PDF
//   const baseUrl = 'https://docs.google.com/spreadsheets/d/' + ssId + '/export?';
  
//   // ค่าขอบ
//   const tiny = 0.01; // ค่าต่ำสุดแทน 0 (แต่ในโค้ดใช้ 0.1 อยู่แล้ว)

//   let exportOptions = {
//     format: 'pdf',
//     gid: sheetId,
//     size: 'A4',
//     portrait: false,       // แนวนอน
//     fitw: true,            // พอดีความกว้าง
    
//     // ⭐ การตั้งค่าขอบกำหนดเอง (หน่วยเป็นนิ้ว) ⭐
//     top_margin: 0.4,       // บน
//     bottom_margin: 0.1,    // ล่าง
//     left_margin: 0.7,      // ซ้าย
//     right_margin: 0.1,     // ขวา
    
//     // ⭐ การจัดแนว: เปิด 'true' คือ จัดกึ่งกลาง
//     horizontal_center: true, // จัดกึ่งกลางแนวนอน
//     vertical_center: true,   // จัดกึ่งกลางแนวตั้ง
    
//     range: encodeURIComponent(printRange), 
//     page: 1, 
//     pagestart: 1,    
//     printtitle: false,
//     gridlines: false,
//     fzr: false,
//   };

//   const params = Object.keys(exportOptions)
//     .map(key => `${key}=${exportOptions[key]}`)
//     .join('&');

//   const finalUrl = baseUrl + params;
  
//   const token = ScriptApp.getOAuthToken();
//   const headers = { 'Authorization': 'Bearer ' + token };

//   try {
//     ui.showModalDialog(HtmlService.createHtmlOutput('<b>กำลังสร้างไฟล์ PDF ชั่วคราวและเปิดในเบราว์เซอร์...</b>'), 'กำลังดำเนินการ');
    
//     // 2. ดึงข้อมูล PDF และสร้างไฟล์ชั่วคราวใน Drive
//     const response = UrlFetchApp.fetch(finalUrl, { headers: headers });
    
//     if (response.getResponseCode() !== 200) {
//       throw new Error(`เกิดข้อผิดพลาดในการดึง PDF: ${response.getResponseCode()}`);
//     }

//     const tempFileName = `${pdfFileName}_${Utilities.getUuid()}.pdf`;
//     const pdfBlob = response.getBlob().setName(tempFileName);
    
//     const pdfFile = DriveApp.createFile(pdfBlob);
//     const openUrl = pdfFile.getUrl(); 
    
//     pdfFile.setTrashed(true); // ย้ายไฟล์ไปถังขยะ

//     // 3. สร้าง HTML เพื่อเปิด URL ในเบราว์เซอร์
//     const htmlOutput = HtmlService
//       .createHtmlOutput(`<script>window.open('${openUrl}', '_blank'); google.script.host.close();</script>`)
//       .setWidth(100)
//       .setHeight(1);
    
//     ui.showModalDialog(htmlOutput, 'กำลังเปิดไฟล์ในแท็บใหม่...');
//     SpreadsheetApp.getActiveSpreadsheet().toast(`✅ ไฟล์ถูกเปิดในเบราว์เซอร์แล้ว`, 'สำเร็จ', 5);

//   } catch (error) {
//     ui.alert(`❌ เกิดข้อผิดพลาดในการสร้าง PDF: ${error.message}`);
//     Logger.log(`PDF Export Error: ${error.message}`);
//   }
// }