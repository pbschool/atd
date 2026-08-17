/**
 * Constants.gs
 */
const MONTHLY_SHEETS = ["05", "06", "07", "08", "09", "10", "11", "12", "01", "02", "03"];
const ZERO_WIDTH_SPACE = String.fromCharCode(8203);

// ตำแหน่งต่างๆ ในชีทรายเดือน
const R_START_ROW = 7;
const R_END_ROW = 26;
const R_HEADER_ROW = 5;
const R_START_COL = 11; // คอลัมน์ K
const R_END_COL = 41;   // คอลัมน์ AO
const R_NAME_COL = 3;   // คอลัมน์ C

// ============================================================================
// === Attendance (rec / recCompact) ==========================================
// ============================================================================
// ชื่อชีตมาตรฐาน (ใช้ร่วมกับ AppSheet)
const REC_SHEET_NAME = "rec";
const COMPACT_SHEET_NAME = "recCompact";

// โครงสร้างคอลัมน์ในชีต rec (ต้องให้สอดคล้องกับการใช้งานใน saveChanges / Ad_reMonth / copy2)
// A: studentCode, B: date, C: level, D: status, E: timestamp
const REC_HEADER_ROW = ["studentCode", "date", "level", "status", "timestamp"];
const REC_SHEET_LAST_COL_WITH_DATA = REC_HEADER_ROW.length;

// Index สำหรับอ่านข้อมูลจากแถว rec (0-based, ใช้กับ getValues())
const REC_COL_STUDENT_CODE_INDEX = 0; // A
const REC_COL_DATE_INDEX = 1;         // B
const REC_COL_STATUS_INDEX = 3;       // D

// ============================================================================
// === Attendance sheet layout (เดือน) ========================================
// ============================================================================
// ชุด constants นี้ถูกเรียกใช้โดย saveChanges / Ad_reMonth / checkMaAll
// เพื่อให้เข้ากับ layout เดียวกับ R_* (K:AO, แถวข้อมูล 7-26, ชื่อ/รหัสอยู่คอลัมน์ C)
const STATUS_START_ROW = R_START_ROW; // 7
const STATUS_MAX_SEARCH_ROW = R_END_ROW; // 26

const STATUS_START_COL = R_START_COL; // 11 (K)
const STATUS_END_COL = R_END_COL;     // 41 (AO)

// แถวหัวตารางที่มี "วัน/วันที่" เพื่อใช้ประกอบการกรอกอัตโนมัติ/บันทึกลง rec
const DAY_ORDER_ROW = R_HEADER_ROW; // 5
const DATE_ROW_FOR_STATUS_SHEET = R_HEADER_ROW; // 5

// คอลัมน์ข้อมูลนักเรียนในชีทรายเดือน
const CODE_COLUMN = R_NAME_COL; // 3 (C)
const NAMES_COL = R_NAME_COL;   // 3 (C) (เข้ากันกับ checkMaAll)
const LEVEL_COLUMN = 4;         // 4 (D) ค่าเริ่มต้น: ระดับชั้น

// Alias สำหรับสคริปต์เก่าบางตัว
const DATA_START_ROW = R_START_ROW; // 7
const DATA_END_ROW = R_END_ROW;     // 26

// คอนฟิกสุขภาพ (สำหรับฟังก์ชัน Health)
const HEALTH_TAB_CONFIG = [
    { name: 'พฤษภาคม', controlRow: 4, dataStartRow: 11, dataEndRow: 30, checkboxRow: 10 },
    { name: 'มิถุนายน', controlRow: 33, dataStartRow: 40, dataEndRow: 59, checkboxRow: 39 },
    { name: 'กรกฎาคม', controlRow: 62, dataStartRow: 69, dataEndRow: 88, checkboxRow: 68 },
    { name: 'สิงหาคม', controlRow: 91, dataStartRow: 98, dataEndRow: 117, checkboxRow: 97 },
    { name: 'กันยายน', controlRow: 120, dataStartRow: 127, dataEndRow: 146, checkboxRow: 126 },
    { name: 'ตุลาคม', controlRow: 149, dataStartRow: 156, dataEndRow: 175, checkboxRow: 155 },
    { name: 'พฤศจิกายน', controlRow: 178, dataStartRow: 185, dataEndRow: 204, checkboxRow: 184 },
    { name: 'ธันวาคม', controlRow: 207, dataStartRow: 214, dataEndRow: 233, checkboxRow: 213 },
    { name: 'มกราคม', controlRow: 236, dataStartRow: 243, dataEndRow: 262, checkboxRow: 242 },
    { name: 'กุมภาพันธ์', controlRow: 265, dataStartRow: 272, dataEndRow: 291, checkboxRow: 271 },
    { name: 'มีนาคม', controlRow: 294, dataStartRow: 301, dataEndRow: 320, checkboxRow: 300 } 
];

const HEALTH_COL_MAPPING = {
    59: [3, 8, 13, 18, 23, 28, 33, 38, 43], 60: [4, 9, 14, 19, 24, 29, 34, 39, 44],
    61: [5, 10, 15, 20, 25, 30, 35, 40, 45], 62: [6, 11, 16, 21, 26, 31, 36, 41, 46],
    63: [7, 12, 17, 22, 27, 32, 37, 42, 47]
};