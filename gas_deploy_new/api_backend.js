const SPREADSHEET_ID_SETTINGS = '1YnLdYnylhp8YKC4Hf6Q3Gf56X93LoW9NVTNdM5Sn8B8';

function doGet(e) {
  if (e && e.parameter && e.parameter.action) {
    return handleApiRequest(e.parameter.action, e.parameter);
  }
  return ContentService.createTextOutput("Attendance API is running").setMimeType(ContentService.MimeType.TEXT);
}

function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    return handleApiRequest(data.action, data.payload);
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ success: false, error: err.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function handleApiRequest(action, payload) {
  let result;
  try {
    switch (action) {
      case 'getSettings':
        result = getSettings();
        break;
      case 'getLinks':
        result = getLinks();
        break;
      case 'getStudents':
        // Expecting url, month, date, year
        result = getStudents(payload.url, payload.month, payload.date, payload.year);
        break;
      case 'saveAttendance':
        result = saveAttendance(payload.url, payload.month, payload.date, payload.data);
        break;
      default:
        result = { success: false, error: 'Unknown action: ' + action };
    }
  } catch (e) {
    result = { success: false, error: e.toString() };
  }
  
  return ContentService.createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}

function getSettings() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID_SETTINGS);
  const sheet = ss.getSheetByName('Menu');
  if (!sheet) throw new Error("Menu sheet not found in central file");
  
  const logoUrl = sheet.getRange('G2').getValue();
  const bgUrl = sheet.getRange('G3').getValue();
  
  return { success: true, data: { logo: logoUrl, background: bgUrl } };
}

function getLinks() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID_SETTINGS);
  const sheet = ss.getSheetByName('LINK');
  if (!sheet) throw new Error("LINK sheet not found in central file");
  
  const yearsRange = sheet.getRange('C19:Z19').getValues()[0];
  const gradesRange = sheet.getRange('A20:A28').getValues();
  const urlsRange = sheet.getRange('C20:Z28').getValues();
  
  let years = [];
  yearsRange.forEach((y, idx) => {
    if (y) years.push({ year: y.toString(), colIndex: idx });
  });
  
  let links = {};
  gradesRange.forEach((row, rIdx) => {
    let grade = row[0];
    if (grade) {
      links[grade] = {};
      years.forEach(y => {
        let url = urlsRange[rIdx][y.colIndex];
        if (url) {
          links[grade][y.year] = url;
        }
      });
    }
  });
  
  let uniqueYears = years.map(y => y.year).sort((a,b) => parseInt(b) - parseInt(a));
  
  return { success: true, data: { years: uniqueYears, links: links } };
}

function getStudents(url, monthTab, dateStr, yearStr) {
  // 1. Fetch Students from destination URL
  const idMatch = url.match(/[-\w]{25,}/);
  if (!idMatch) throw new Error("Invalid URL");
  const ss = SpreadsheetApp.openById(idMatch[0]);
  const sheet = ss.getSheetByName(monthTab);
  if (!sheet) throw new Error("Month tab '" + monthTab + "' not found in destination sheet");
  
  const dates = sheet.getRange('K2:AO2').getDisplayValues()[0];
  let dateCol = -1;
  for (let i = 0; i < dates.length; i++) {
    if (dates[i] === dateStr || dates[i].includes(dateStr)) {
      dateCol = 11 + i; // K is 11th col
      break;
    }
  }
  
  const ids = sheet.getRange('C7:C26').getValues();
  const names = sheet.getRange('E7:E26').getValues();
  const nicknames = sheet.getRange('F7:F26').getValues();
  
  let statuses = [];
  if (dateCol !== -1) {
    statuses = sheet.getRange(7, dateCol, 20, 1).getValues();
  }
  
  // 2. Fetch Photos from settings sheet url
  let photoMap = {};
  try {
    const settingsSs = SpreadsheetApp.openById(SPREADSHEET_ID_SETTINGS);
    const settingSheet = settingsSs.getSheetByName('setting');
    let photoUrlOrId = '';
    if (settingSheet) {
      photoUrlOrId = settingSheet.getRange('C5').getValue();
    }
    
    if (photoUrlOrId) {
      const idMatch = String(photoUrlOrId).match(/[-\w]{25,}/);
      if (idMatch) {
        const photoSs = SpreadsheetApp.openById(idMatch[0]);
        const photoSheet = photoSs.getSheetByName('infopic');
        if (photoSheet) {
          // Assuming data from row 2 downwards
          const lastRow = photoSheet.getLastRow();
          if (lastRow >= 2) {
            // Read Col A (Year) and Col C (Student ID) and Col I (File_ID)
            // A=1, C=3, I=9
            const yearsData = photoSheet.getRange(2, 1, lastRow - 1, 1).getValues();
            const idsData = photoSheet.getRange(2, 3, lastRow - 1, 1).getValues();
            const filesData = photoSheet.getRange(2, 9, lastRow - 1, 1).getValues();
            
            for (let i = 0; i < yearsData.length; i++) {
              if (yearsData[i][0] && String(yearsData[i][0]) === String(yearStr)) {
                let sId = String(idsData[i][0]).trim();
                let fileId = filesData[i][0];
                if (sId && fileId) {
                  photoMap[sId] = fileId;
                }
              }
            }
          }
        }
      }
    }
  } catch (e) {
    // If photo sheet errors out, we just continue without photos
    console.error("Error reading photo sheet: " + e.toString());
  }
  
  // 3. Combine Data
  let students = [];
  for (let i = 0; i < ids.length; i++) {
    let sId = ids[i][0] ? String(ids[i][0]).trim() : '';
    if (sId) {
      students.push({
        id: sId,
        name: names[i][0],
        nickname: nicknames[i][0],
        status: (dateCol !== -1 && statuses[i] && statuses[i][0]) ? statuses[i][0] : '-',
        photoFileId: photoMap[sId] || ''
      });
    }
  }
  
  return { success: true, data: students };
}

function saveAttendance(url, monthTab, dateStr, data) {
  const idMatch = url.match(/[-\w]{25,}/);
  if (!idMatch) throw new Error("Invalid URL");
  const ss = SpreadsheetApp.openById(idMatch[0]);
  const sheet = ss.getSheetByName(monthTab);
  if (!sheet) throw new Error("Month tab '" + monthTab + "' not found");
  
  const dates = sheet.getRange('K2:AO2').getDisplayValues()[0];
  let dateCol = -1;
  for (let i = 0; i < dates.length; i++) {
    if (dates[i] === dateStr || dates[i].includes(dateStr)) {
      dateCol = 11 + i;
      break;
    }
  }
  
  if (dateCol === -1) {
    throw new Error("Date " + dateStr + " not found in row 2");
  }
  
  const ids = sheet.getRange('C7:C26').getValues();
  let currentStatuses = sheet.getRange(7, dateCol, 20, 1).getValues();
  
  let updated = false;
  for (let i = 0; i < ids.length; i++) {
    let studentId = ids[i][0];
    if (studentId && data[studentId] !== undefined) {
       currentStatuses[i][0] = data[studentId];
       updated = true;
    }
  }
  
  if (updated) {
    sheet.getRange(7, dateCol, 20, 1).setValues(currentStatuses);
  }
  
  return { success: true };
}
