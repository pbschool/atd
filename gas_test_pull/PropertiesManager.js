// PropertiesManager.gs

/**
 * ฟังก์ชันหลักสำหรับจัดการ PropertiesService
 * รวมถึงการตั้งค่า, ดึงค่า, และลบค่าต่างๆ
 * มี 3 ขอบเขตให้เลือกใช้งาน: Script, User, Document
 */

/**
 * Retrieves a specific property value from Script Properties.
 * Properties in this scope are shared across all users of the script.
 * @param {string} key The key of the property to retrieve.
 * @returns {string|null} The value of the property, or null if not found.
 */
function getScriptProperty(key) {
  const properties = PropertiesService.getScriptProperties();
  const value = properties.getProperty(key);
  if (value === null) {
    Logger.log(`Script property '${key}' not found.`);
  } else {
    Logger.log(`Retrieved script property '${key}': ${value}`);
  }
  return value;
}

/**
 * Sets a specific property value in Script Properties.
 * Values are stored as strings. If you pass an object/array, it will be stringified.
 * @param {string} key The key of the property to set.
 * @param {any} value The value to set for the property.
 */
function setScriptProperty(key, value) {
  const properties = PropertiesService.getScriptProperties();
  try {
    // If value is an object or array, convert it to JSON string
    const stringValue = (typeof value === 'object' && value !== null) ? JSON.stringify(value) : String(value);
    properties.setProperty(key, stringValue);
    Logger.log(`Set script property '${key}' to '${stringValue}'.`);
  } catch (e) {
    Logger.log(`Error setting script property '${key}': ${e.message}`);
    // *** บรรทัด SpreadsheetApp.getUi().alert ถูกลบออกแล้ว ***
  }
}

/**
 * Deletes a specific property from Script Properties.
 * @param {string} key The key of the property to delete.
 */
function deleteScriptProperty(key) {
  const properties = PropertiesService.getScriptProperties();
  properties.deleteProperty(key);
  Logger.log(`Deleted script property '${key}'.`);
}

/**
 * Retrieves all properties from Script Properties.
 * Useful for inspecting all stored configurations.
 * @returns {Object} An object containing all script properties.
 */
function getAllScriptProperties() {
  const properties = PropertiesService.getScriptProperties();
  const allProps = properties.getProperties();
  Logger.log('All Script Properties:', allProps);
  return allProps;
}

/**
 * Deletes all properties from Script Properties.
 * Use with caution, as this will clear all shared configurations.
 */
function deleteAllScriptProperties() {
  const properties = PropertiesService.getScriptProperties();
  properties.deleteAllProperties();
  Logger.log('All Script Properties deleted.');
}

// --- ตัวอย่างการใช้งานสำหรับ User Properties (เฉพาะผู้ใช้ปัจจุบัน) ---

/**
 * Retrieves a specific property value from User Properties.
 * Properties in this scope are specific to the current user.
 * @param {string} key The key of the property to retrieve.
 * @returns {string|null} The value of the property, or null if not found.
 */
function getUserProperty(key) {
  const properties = PropertiesService.getUserProperties();
  const value = properties.getProperty(key);
  if (value === null) {
    Logger.log(`User property '${key}' not found.`);
  } else {
    Logger.log(`Retrieved user property '${key}': ${value}`);
  }
  return value;
}

/**
 * Sets a specific property value in User Properties.
 * @param {string} key The key of the property to set.
 * @param {any} value The value to set for the property.
 */
function setUserProperty(key, value) {
  const properties = PropertiesService.getUserProperties();
  try {
    const stringValue = (typeof value === 'object' && value !== null) ? JSON.stringify(value) : String(value);
    properties.setProperty(key, stringValue);
    Logger.log(`Set user property '${key}' to '${stringValue}'.`);
  } catch (e) {
    Logger.log(`Error setting user property '${key}': ${e.message}`);
    // *** บรรทัด SpreadsheetApp.getUi().alert ถูกลบออกแล้ว ***
  }
}

/**
 * Deletes a specific property from User Properties.
 * @param {string} key The key of the property to delete.
 */
function deleteUserProperty(key) {
  const properties = PropertiesService.getUserProperties();
  properties.deleteProperty(key);
  Logger.log(`Deleted user property '${key}'.`);
}

// --- ตัวอย่างการใช้งานสำหรับ Document Properties (เฉพาะเอกสารนี้, ใช้กับ Script ที่ผูกกับเอกสารเท่านั้น) ---

/**
 * Retrieves a specific property value from Document Properties.
 * Properties in this scope are specific to the document the script is bound to.
 * @param {string} key The key of the property to retrieve.
 * @returns {string|null} The value of the property, or null if not found.
 */
function getDocumentProperty(key) {
  // Document Properties can only be used in scripts bound to a document (e.g., Spreadsheet, Doc).
  // If running in a standalone script, this will throw an error.
  if (SpreadsheetApp.getActiveSpreadsheet()) { 
    const properties = PropertiesService.getDocumentProperties();
    const value = properties.getProperty(key);
    if (value === null) {
      Logger.log(`Document property '${key}' not found.`);
    } else {
      Logger.log(`Retrieved document property '${key}': ${value}`);
    }
    return value;
  } else {
    Logger.log("Cannot access Document Properties from a standalone script.");
    return null;
  }
}

/**
 * Sets a specific property value in Document Properties.
 * @param {string} key The key of the property to set.
 * @param {any} value The value to set for the property.
 */
function setDocumentProperty(key, value) {
  if (SpreadsheetApp.getActiveSpreadsheet()) {
    const properties = PropertiesService.getDocumentProperties();
    try {
      const stringValue = (typeof value === 'object' && value !== null) ? JSON.stringify(value) : String(value);
      properties.setProperty(key, stringValue);
      Logger.log(`Set document property '${key}' to '${stringValue}'.`);
    } catch (e) {
      Logger.log(`Error setting document property '${key}': ${e.message}`);
      // *** บรรทัด SpreadsheetApp.getUi().alert ถูกลบออกแล้ว ***
    }
  } else {
    Logger.log("Cannot set Document Properties from a standalone script.");
  }
}

/**
 * Example function to demonstrate how to use PropertiesService.
 * You can run this function to see the logging output.
 */
function testPropertiesService() {
  Logger.log('--- Testing PropertiesService ---');

  // Test Script Properties
  Logger.log('Testing Script Properties:');
  setScriptProperty('myGlobalSetting', 'active');
  const globalSetting = getScriptProperty('myGlobalSetting');
  Logger.log(`Current global setting: ${globalSetting}`);

  // Storing an object/array
  const configData = {
    sheets: { rec: 'RecSheet', backup: '_BackupData' },
    version: '1.0'
  };
  setScriptProperty('appConfig', configData);
  const retrievedConfigJson = getScriptProperty('appConfig');
  if (retrievedConfigJson) {
    const retrievedConfig = JSON.parse(retrievedConfigJson);
    Logger.log('Retrieved appConfig:', retrievedConfig.sheets.rec);
  }

  // Get all properties
  getAllScriptProperties();

  // Test User Properties
  Logger.log('\nTesting User Properties:');
  setUserProperty('myUserPreference', 'dark_mode');
  const userPref = getUserProperty('myUserPreference');
  Logger.log(`Current user preference: ${userPref}`);

  // Test Document Properties (if script is bound to a spreadsheet)
  Logger.log('\nTesting Document Properties:');
  if (SpreadsheetApp.getActiveSpreadsheet()) { 
    setDocumentProperty('sheetId', SpreadsheetApp.getActiveSpreadsheet().getId());
    const docId = getDocumentProperty('sheetId');
    Logger.log(`Current document ID: ${docId}`);
  } else {
    Logger.log('Skipping Document Properties test as this is not a document-bound script.');
  }

  // Clean up (optional)
  // deleteScriptProperty('myGlobalSetting');
  // deleteScriptProperty('appConfig');
  // deleteUserProperty('myUserPreference');
  // if (SpreadsheetApp.getActiveSpreadsheet()) {
  //   deleteDocumentProperty('sheetId');
  // }
  Logger.log('--- PropertiesService Test Complete ---');
}