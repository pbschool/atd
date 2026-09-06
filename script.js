const DEFAULT_SPREADSHEET_ID = '1r9BWbErUqeZexwrvKPhve0zA6cm7vEl01CMme6_hbi4';
const DEFAULT_API_URL = 'https://script.google.com/macros/s/AKfycbwjReuLuLZ7UL5XJKdclBgrIMiWIV0wWtdWaIhCp4AfpW4ZWe4sDw0q_P4_LyTWmqcQ/exec';

const state = {
    apiUrl: localStorage.getItem('apiUrl') || DEFAULT_API_URL,
    spreadsheetId: localStorage.getItem('spreadsheetId') || DEFAULT_SPREADSHEET_ID,
    logoUrl: localStorage.getItem('logoUrl') || '',
    bgUrl: localStorage.getItem('bgUrl') || '',
    config: JSON.parse(localStorage.getItem('config')) || { years: [], links: {} },
    lastYear: localStorage.getItem('lastYear') || '',
    lastGrade: localStorage.getItem('lastGrade') || 'อ.1',
    studentsData: [],
    attendanceData: {},
    currentSelectedUrl: '',
    currentSelectedMonthTab: '',
    currentSelectedDateStr: ''
};

// UI Elements
const els = {
    navBtns: document.querySelectorAll('.nav-btn'),
    pages: document.querySelectorAll('.page'),
    appLogo: document.getElementById('app-logo'),
    bgOverlay: document.getElementById('bg-overlay'),
    
    // Record Page
    selectYear: document.getElementById('select-year'),
    selectGrade: document.getElementById('select-grade'),
    selectDate: document.getElementById('select-date'),
    btnNextStep: document.getElementById('btn-next-step'),
    btnBackStep: document.getElementById('btn-back-step'),
    step1: document.getElementById('record-step-1'),
    step2: document.getElementById('record-step-2'),
    loadingConfig: document.getElementById('loading-config'),
    loadingStudents: document.getElementById('loading-students'),
    studentsList: document.getElementById('students-list'),
    currentClassInfo: document.getElementById('current-class-info'),
    saveRow: document.getElementById('save-row'),
    btnSaveAttendance: document.getElementById('btn-save-attendance'),
    bulkActions: document.getElementById('bulk-actions'),
    btnSelectAllPresent: document.getElementById('btn-select-all-present'),
    btnSelectAllAbsent: document.getElementById('btn-select-all-absent'),
    btnClearAll: document.getElementById('btn-clear-all'),
    
    // Home Page Elements
    homeSelectYear: document.getElementById('home-select-year'),
    homeGradeButtons: document.querySelectorAll('.btn-grade-opt'),
    homeCurrentGradeDisplay: document.getElementById('home-current-grade-display'),
    btnHomeGoRecord: document.getElementById('btn-home-go-record'),
    btnHomeGoReport: document.getElementById('btn-home-go-report'),
    reportSelectedGradeLabel: document.getElementById('report-selected-grade-label'),
    dynGradeTexts: document.querySelectorAll('.dyn-grade-text'),

    // Settings Page
    adminAuthSection: document.getElementById('admin-auth-section'),
    adminPassword: document.getElementById('admin-password'),
    btnAdminUnlock: document.getElementById('btn-admin-unlock'),
    settingsContentSection: document.getElementById('settings-content-section'),
    apiUrlInput: document.getElementById('api-url'),
    btnSyncSettings: document.getElementById('btn-sync-settings'),
    syncStatus: document.getElementById('sync-status'),

    // Daily Report Elements
    btnOpenDailyReport: document.getElementById('btn-open-daily-report'),
    reportMenuView: document.getElementById('report-menu-view'),
    dailyReportView: document.getElementById('daily-report-view'),
    btnDailyBack: document.getElementById('btn-daily-back'),
    btnDailyRefresh: document.getElementById('btn-daily-refresh'),
    btnDailyCopyImage: document.getElementById('btn-daily-copy-image'),
    btnDailySaveImage: document.getElementById('btn-daily-save-image'),
    btnDailyCopyText: document.getElementById('btn-daily-copy-text'),
    dailyToast: document.getElementById('daily-toast'),
    dailyLoading: document.getElementById('daily-loading'),
    dailyReportCard: document.getElementById('daily-report-card'),
    dailyReportDateText: document.getElementById('daily-report-date-text'),
    dailyStatsTbody: document.getElementById('daily-stats-tbody')
};

// Initialize
function init() {
    applyBranding();
    setupNavigation();
    
    // Set Default Date
    const today = new Date();
    const yyyy = today.getFullYear();
    const mm = String(today.getMonth() + 1).padStart(2, '0');
    const dd = String(today.getDate()).padStart(2, '0');
    els.selectDate.value = `${yyyy}-${mm}-${dd}`;
    
    populateYearDropdown();
    setSelectedGrade(state.lastGrade || 'อ.1');
    
    // Setup Settings initial state
    els.apiUrlInput.value = state.apiUrl;
    
    // Auto sync on load if API URL is provided
    if (state.apiUrl) {
        autoSyncSettings();
    }
    
    // Event Listeners
    els.btnAdminUnlock.addEventListener('click', handleAdminUnlock);
    els.btnSyncSettings.addEventListener('click', handleSyncSettings);
    els.btnNextStep.addEventListener('click', handleNextStep);
    els.btnBackStep.addEventListener('click', handleBackStep);
    els.btnSaveAttendance.addEventListener('click', handleSaveAttendance);
    els.btnSelectAllPresent.addEventListener('click', handleSelectAllPresent);
    if (els.btnSelectAllAbsent) {
        els.btnSelectAllAbsent.addEventListener('click', handleSelectAllAbsent);
    }
    els.btnClearAll.addEventListener('click', handleClearAllStatus);
    
    // Home Page Grade Selector Listeners
    if (els.homeGradeButtons) {
        els.homeGradeButtons.forEach(btn => {
            btn.addEventListener('click', () => {
                const grade = btn.getAttribute('data-grade');
                if (grade) setSelectedGrade(grade);
            });
        });
    }
    
    if (els.btnHomeGoRecord) {
        els.btnHomeGoRecord.addEventListener('click', () => {
            const recordNav = document.querySelector('.nav-btn[data-target="page-record"]');
            if (recordNav) recordNav.click();
        });
    }
    
    if (els.btnHomeGoReport) {
        els.btnHomeGoReport.addEventListener('click', () => {
            const reportNav = document.querySelector('.nav-btn[data-target="page-report"]');
            if (reportNav) reportNav.click();
        });
    }
    
    // Daily Report Listeners
    if (els.btnOpenDailyReport) {
        els.btnOpenDailyReport.addEventListener('click', handleOpenDailyReport);
    }
    if (els.btnDailyBack) {
        els.btnDailyBack.addEventListener('click', handleDailyBack);
    }
    if (els.btnDailyRefresh) {
        els.btnDailyRefresh.addEventListener('click', () => loadDailyAttendanceStats(true));
    }
    if (els.btnDailyCopyImage) {
        els.btnDailyCopyImage.addEventListener('click', handleCopyDailyReportImage);
    }
    if (els.btnDailySaveImage) {
        els.btnDailySaveImage.addEventListener('click', handleSaveDailyReportImage);
    }
    if (els.btnDailyCopyText) {
        els.btnDailyCopyText.addEventListener('click', handleCopyDailySummaryText);
    }
    
    els.selectYear.addEventListener('change', (e) => {
        state.lastYear = e.target.value;
        localStorage.setItem('lastYear', e.target.value);
        if (els.homeSelectYear) els.homeSelectYear.value = e.target.value;
    });
    
    if (els.homeSelectYear) {
        els.homeSelectYear.addEventListener('change', (e) => {
            state.lastYear = e.target.value;
            localStorage.setItem('lastYear', e.target.value);
            els.selectYear.value = e.target.value;
        });
    }
    
    els.selectGrade.addEventListener('change', (e) => {
        setSelectedGrade(e.target.value);
    });
}

function setSelectedGrade(grade) {
    if (!grade) return;
    state.lastGrade = grade;
    localStorage.setItem('lastGrade', grade);
    
    // Update Home page buttons
    if (els.homeGradeButtons) {
        els.homeGradeButtons.forEach(b => {
            if (b.getAttribute('data-grade') === grade) {
                b.classList.add('active');
            } else {
                b.classList.remove('active');
            }
        });
    }
    
    // Update Home page displays
    if (els.homeCurrentGradeDisplay) {
        els.homeCurrentGradeDisplay.textContent = grade;
    }
    if (els.dynGradeTexts) {
        els.dynGradeTexts.forEach(el => el.textContent = grade);
    }
    
    // Update Record page select
    if (els.selectGrade) {
        els.selectGrade.value = grade;
    }
    
    // Update Report page label
    if (els.reportSelectedGradeLabel) {
        els.reportSelectedGradeLabel.textContent = grade;
    }
}

function applyBranding() {
    if (state.logoUrl) {
        els.appLogo.src = state.logoUrl;
        els.appLogo.style.display = 'block';
    }
    if (state.bgUrl) {
        els.bgOverlay.style.backgroundImage = `url('${state.bgUrl}')`;
    }
}

function setupNavigation() {
    els.navBtns.forEach(btn => {
        btn.addEventListener('click', (e) => {
            // Remove active from all
            els.navBtns.forEach(b => b.classList.remove('active'));
            els.pages.forEach(p => p.classList.remove('active'));
            
            // Add active to clicked
            e.target.classList.add('active');
            const targetId = e.target.getAttribute('data-target');
            document.getElementById(targetId).classList.add('active');
        });
    });
}

function populateYearDropdown() {
    const currentVal = els.selectYear.value || state.lastYear;
    els.selectYear.innerHTML = '';
    if (els.homeSelectYear) els.homeSelectYear.innerHTML = '';
    
    const today = new Date();
    let thaiYear = today.getFullYear() + 543;
    if (today.getMonth() < 4) thaiYear -= 1; // Academic year typically starts in May
    
    // Force exactly 2 years: Current and Previous
    let yearsToShow = [thaiYear.toString(), (thaiYear - 1).toString()];

    yearsToShow.forEach(y => {
        const option1 = document.createElement('option');
        option1.value = y;
        option1.textContent = y;
        els.selectYear.appendChild(option1);

        if (els.homeSelectYear) {
            const option2 = document.createElement('option');
            option2.value = y;
            option2.textContent = y;
            els.homeSelectYear.appendChild(option2);
        }
    });
    
    if (currentVal && yearsToShow.includes(currentVal)) {
        els.selectYear.value = currentVal;
        if (els.homeSelectYear) els.homeSelectYear.value = currentVal;
    } else {
        els.selectYear.value = yearsToShow[0];
        if (els.homeSelectYear) els.homeSelectYear.value = yearsToShow[0];
    }
    state.lastYear = els.selectYear.value;
    localStorage.setItem('lastYear', els.selectYear.value);
}

async function apiRequest(action, payload = {}) {
    if (!state.apiUrl) {
        throw new Error('กรุณาตั้งค่า API URL ในหน้าตั้งค่าก่อน');
    }
    const response = await fetch(state.apiUrl, {
        method: 'POST',
        headers: {
            'Content-Type': 'text/plain;charset=utf-8', 
        },
        body: JSON.stringify({ action, payload })
    });
    const text = await response.text();
    let result;
    try {
        result = JSON.parse(text);
    } catch (e) {
        if (text.includes('<!DOCTYPE') || text.includes('<html')) {
            throw new Error('Web App ส่งกลับเป็นหน้าเว็บ HTML (อาจเกิดจากยังไม่ได้ตั้งค่าสิทธิ์เข้าถึงเป็น "Anyone / ทุกคน" ตอน Deploy หรือติดหน้าขอสิทธิ์ Login ของ Google)');
        }
        throw new Error('ไม่สามารถอ่านข้อมูลผลลัพธ์จาก API ได้: ' + text.substring(0, 100));
    }
    if (!result.success) {
        throw new Error(result.error || 'API Error');
    }
    return result.data;
}

// --- Settings ---
function handleAdminUnlock() {
    if (els.adminPassword.value === '9919') {
        els.adminAuthSection.style.display = 'none';
        els.settingsContentSection.style.display = 'block';
        els.adminPassword.value = '';
    } else {
        alert('รหัสผ่านผู้ดูแลระบบไม่ถูกต้อง');
    }
}

async function handleSyncSettings() {
    const url = els.apiUrlInput.value.trim();
    if (!url) {
        els.syncStatus.innerHTML = '<span style="color:var(--st-absent)">กรุณาระบุ URL</span>';
        return;
    }
    state.apiUrl = url;
    localStorage.setItem('apiUrl', url);
    els.syncStatus.innerHTML = '<span style="color:var(--st-leave)">กำลังซิงค์...</span>';
    els.btnSyncSettings.disabled = true;
    
    try {
        const [settingsRes, linksRes] = await Promise.all([
            apiRequest('getSettings'),
            apiRequest('getLinks')
        ]);
        
        state.logoUrl = settingsRes.logo || '';
        state.bgUrl = settingsRes.background || '';
        localStorage.setItem('logoUrl', state.logoUrl);
        localStorage.setItem('bgUrl', state.bgUrl);
        
        state.config = linksRes;
        localStorage.setItem('config', JSON.stringify(linksRes));
        
        applyBranding();
        populateYearDropdown();
        
        els.syncStatus.innerHTML = '<span style="color:var(--st-present)">ซิงค์สำเร็จ!</span>';
    } catch (err) {
        els.syncStatus.innerHTML = `<span style="color:var(--st-absent)">เกิดข้อผิดพลาด: ${err.message}</span>`;
    } finally {
        els.btnSyncSettings.disabled = false;
    }
}

// --- Auto Sync ---
async function autoSyncSettings() {
    try {
        const [settingsRes, linksRes] = await Promise.all([
            apiRequest('getSettings'),
            apiRequest('getLinks')
        ]);
        
        state.logoUrl = settingsRes.logo || '';
        state.bgUrl = settingsRes.background || '';
        localStorage.setItem('logoUrl', state.logoUrl);
        localStorage.setItem('bgUrl', state.bgUrl);
        
        state.config = linksRes;
        localStorage.setItem('config', JSON.stringify(linksRes));
        
        applyBranding();
        populateYearDropdown();
        console.log('ซิงค์ข้อมูลสำเร็จ');
    } catch (err) {
        console.error('ซิงค์ข้อมูลไม่สำเร็จ:', err);
    }
}

// --- Record Attendance ---
async function handleNextStep() {
    const year = els.selectYear.value;
    const grade = els.selectGrade.value;
    const dateStr = els.selectDate.value; // YYYY-MM-DD
    
    if (!state.apiUrl) {
        alert('ยังไม่ได้ระบุ Web App URL สำหรับเชื่อมต่อระบบ\nกรุณาไปที่เมนู "ตั้งค่า" (รหัสผ่าน 9919) เพื่อระบุ API URL หรือใส่ค่าใน DEFAULT_API_URL ใน script.js');
        return;
    }
    
    if (!state.config || !state.config.links || Object.keys(state.config.links).length === 0) {
        alert('ระบบยังโหลดข้อมูลห้องเรียนไม่สำเร็จ\nกรุณาไปที่เมนู "ตั้งค่า" แล้วกดปุ่ม "ซิงค์การตั้งค่าล่าสุด" อีกครั้งครับ');
        return;
    }
    
    if (!year) {
        alert('กรุณาเลือกปีการศึกษา');
        return;
    }
    
    const targetUrl = state.config.links[grade] && state.config.links[grade][year];
    if (!targetUrl) {
        alert(`ไม่พบลิงก์สำหรับชั้น ${grade} ปีการศึกษา ${year} ในระบบ (อาจจะยังไม่มีข้อมูลใน Google Sheets หรือยังไม่ได้ซิงค์การตั้งค่า)`);
        return;
    }
    
    // Extract month (01-12) from date without timezone offset
    const monthTab = dateStr ? dateStr.split('-')[1] : '05';
    
    state.currentSelectedUrl = targetUrl;
    state.currentSelectedMonthTab = monthTab;
    state.currentSelectedDateStr = dateStr;
    state.lastYear = year;
    state.lastGrade = grade;
    
    els.currentClassInfo.textContent = `ชั้น ${grade} (ปี ${year}) | วันที่ ${dateStr}`;
    els.step1.classList.remove('active');
    els.step2.classList.add('active');
    els.loadingStudents.style.display = 'block';
    els.studentsList.innerHTML = '';
    els.saveRow.style.display = 'none';
    if (els.bulkActions) els.bulkActions.style.display = 'none';
    
    try {
        const students = await apiRequest('getStudents', {
            url: targetUrl,
            month: monthTab,
            date: dateStr,
            year: year
        });
        
        state.studentsData = students;
        state.attendanceData = {};
        renderStudents(students);
        
    } catch (err) {
        els.loadingStudents.innerHTML = `<span style="color:var(--st-absent)">เกิดข้อผิดพลาด: ${err.message}</span>`;
    }
}

function handleBackStep() {
    els.step2.classList.remove('active');
    els.step1.classList.add('active');
    els.loadingStudents.innerHTML = 'กำลังโหลดรายชื่อนักเรียน...';
    if (els.bulkActions) els.bulkActions.style.display = 'none';
    if (state.lastGrade) els.selectGrade.value = state.lastGrade;
    if (state.lastYear) els.selectYear.value = state.lastYear;
}

function handleSelectAllPresent() {
    const rows = document.querySelectorAll('.student-row:not(.is-transferred)');
    rows.forEach(row => {
        const studentId = row.getAttribute('data-id');
        if (studentId) {
            state.attendanceData[studentId] = 'มา';
            
            // Update UI
            const btns = row.querySelectorAll('.status-btn');
            btns.forEach(b => b.classList.remove('selected'));
            const presentBtn = row.querySelector('.status-btn[data-status="มา"]');
            if (presentBtn) presentBtn.classList.add('selected');
        }
    });
}

function handleSelectAllAbsent() {
    const rows = document.querySelectorAll('.student-row:not(.is-transferred)');
    rows.forEach(row => {
        const studentId = row.getAttribute('data-id');
        if (studentId) {
            state.attendanceData[studentId] = 'ขาด';
            
            // Update UI
            const btns = row.querySelectorAll('.status-btn');
            btns.forEach(b => b.classList.remove('selected'));
            const absentBtn = row.querySelector('.status-btn[data-status="ขาด"]');
            if (absentBtn) absentBtn.classList.add('selected');
        }
    });
}

function handleClearAllStatus() {
    const rows = document.querySelectorAll('.student-row:not(.is-transferred)');
    rows.forEach(row => {
        const studentId = row.getAttribute('data-id');
        if (studentId) {
            state.attendanceData[studentId] = '';
            
            // Update UI
            const btns = row.querySelectorAll('.status-btn');
            btns.forEach(b => b.classList.remove('selected'));
            const blankBtn = row.querySelector('.status-btn[data-status=""]');
            if (blankBtn) blankBtn.classList.add('selected');
        }
    });
}

function renderStudents(students) {
    els.loadingStudents.style.display = 'none';
    els.studentsList.innerHTML = '';
    
    if (students.length === 0) {
        els.studentsList.innerHTML = '<p class="text-center text-muted">ไม่พบรายชื่อนักเรียนในแท็บเดือนนี้</p>';
        return;
    }
    
    const template = document.getElementById('student-row-template');
    
    students.forEach((student, index) => {
        const clone = template.content.cloneNode(true);
        const row = clone.querySelector('.student-row');
        row.setAttribute('data-id', student.id);
        
        row.querySelector('.student-no').textContent = index + 1;
        row.querySelector('.student-fullname').textContent = student.name || 'ไม่มีชื่อ';
        row.querySelector('.student-nickname').textContent = student.nickname ? `(${student.nickname})` : '';
        
        const rawStatus = (student.status !== undefined && student.status !== null) ? String(student.status).trim() : '';
        const isTransferred = (rawStatus === '-');
        
        if (isTransferred) {
            row.classList.add('is-transferred');
            const badge = document.createElement('span');
            badge.className = 'badge-transferred';
            badge.textContent = '(ย้ายแล้ว)';
            row.querySelector('.student-names').appendChild(badge);
        }
        
        const imgPlaceholder = row.querySelector('.student-img-placeholder');
        if (student.photoFileId) {
            imgPlaceholder.innerHTML = `<img src="https://drive.google.com/thumbnail?id=${student.photoFileId}&sz=w200" style="width:100%; height:100%; border-radius:50%; object-fit:cover;" referrerpolicy="no-referrer" />`;
            // Remove the default avatar icon
            const sheetStyle = document.createElement('style');
            sheetStyle.innerHTML = `.student-row[data-id="${student.id}"] .student-img-placeholder::after { display: none; }`;
            document.head.appendChild(sheetStyle);
        }
        
        // Status buttons setup
        const statusBtns = row.querySelectorAll('.status-btn');
        state.attendanceData[student.id] = rawStatus;
        
        statusBtns.forEach(btn => {
            const btnStatus = btn.getAttribute('data-status');
            
            if (isTransferred) {
                btn.disabled = true;
            } else {
                if (btnStatus === rawStatus) {
                    btn.classList.add('selected');
                }
                
                btn.addEventListener('click', (e) => {
                    // Remove selected from all in this row
                    statusBtns.forEach(b => b.classList.remove('selected'));
                    e.target.classList.add('selected');
                    
                    const newStatus = e.target.getAttribute('data-status');
                    state.attendanceData[student.id] = newStatus;
                });
            }
        });
        
        els.studentsList.appendChild(clone);
    });
    
    els.saveRow.style.display = 'block';
    if (els.bulkActions) els.bulkActions.style.display = 'flex';
}

async function handleSaveAttendance() {
    els.btnSaveAttendance.disabled = true;
    els.btnSaveAttendance.textContent = 'กำลังบันทึก...';
    
    try {
        await apiRequest('saveAttendance', {
            url: state.currentSelectedUrl,
            month: state.currentSelectedMonthTab,
            date: state.currentSelectedDateStr,
            data: state.attendanceData
        });
        
        alert('บันทึกข้อมูลเรียบร้อยแล้ว!');
    } catch (err) {
        alert('เกิดข้อผิดพลาดในการบันทึก: ' + err.message);
    } finally {
        els.btnSaveAttendance.disabled = false;
        els.btnSaveAttendance.textContent = 'บันทึกข้อมูล';
    }
}

// --- Daily Attendance Report ---
let currentDailyData = null;

function handleOpenDailyReport() {
    els.reportMenuView.style.display = 'none';
    els.dailyReportView.style.display = 'block';
    loadDailyAttendanceStats();
}

function handleDailyBack() {
    els.dailyReportView.style.display = 'none';
    els.reportMenuView.style.display = 'block';
    hideDailyToast();
}

function showDailyToast(message, type = 'success', duration = 4000) {
    if (!els.dailyToast) return;
    els.dailyToast.textContent = message;
    els.dailyToast.className = `daily-toast-msg toast-${type}`;
    els.dailyToast.style.display = 'block';
    
    if (window.dailyToastTimer) clearTimeout(window.dailyToastTimer);
    window.dailyToastTimer = setTimeout(() => {
        hideDailyToast();
    }, duration);
}

function hideDailyToast() {
    if (els.dailyToast) els.dailyToast.style.display = 'none';
}

async function loadDailyAttendanceStats(isRefresh = false) {
    els.dailyLoading.style.display = 'block';
    els.dailyReportCard.style.display = 'none';
    if (isRefresh) {
        showDailyToast('กำลังดึงข้อมูลล่าสุดจาก Google Sheets...', 'info', 2000);
    }

    try {
        const response = await apiRequest('getDailyAttendanceStats');
        currentDailyData = response;
        renderDailyStats(response);
        els.dailyReportCard.style.display = 'block';
        if (isRefresh) {
            showDailyToast('อัปเดตข้อมูลล่าสุดเรียบร้อยแล้ว!', 'success');
        }
    } catch (err) {
        els.dailyStatsTbody.innerHTML = `<tr><td colspan="12" style="color:var(--st-absent); padding:20px;">เกิดข้อผิดพลาด: ${err.message}</td></tr>`;
        els.dailyReportCard.style.display = 'block';
        showDailyToast(`เกิดข้อผิดพลาด: ${err.message}`, 'error', 5000);
    } finally {
        els.dailyLoading.style.display = 'none';
    }
}

function renderDailyStats(data) {
    els.dailyStatsTbody.innerHTML = '';
    
    if (data.updatedAt) {
        els.dailyReportDateText.textContent = `ข้อมูล ณ วันที่ ${data.updatedAt}`;
    }
    
    if (!data.rows || data.rows.length === 0) {
        els.dailyStatsTbody.innerHTML = '<tr><td colspan="12" style="padding:20px; text-align:center; color:var(--text-muted);">ไม่พบข้อมูลสถิติในแท็บ DayNow</td></tr>';
        return;
    }
    
    data.rows.forEach(row => {
        const tr = document.createElement('tr');
        
        let rowClass = 'row-regular';
        if (row.isGrandTotal || row.grade.includes('รวมทั้งหมด')) {
            rowClass = 'row-summary row-summary-grand-total';
        } else if (row.grade.includes('รวมอนุบาล')) {
            rowClass = 'row-summary row-summary-kinder';
        } else if (row.grade.includes('รวมประถมต้น')) {
            rowClass = 'row-summary row-summary-primary-lower';
        } else if (row.grade.includes('รวมประถมปลาย')) {
            rowClass = 'row-summary row-summary-primary-upper';
        } else if (row.grade.includes('รวมประถม')) {
            rowClass = 'row-summary row-summary-primary';
        }
        
        tr.className = rowClass;
        
        // Note Badge
        let noteHtml = '';
        const rawNote = (row.note || '').trim();
        if (rawNote === 'ผ่าน') {
            noteHtml = `<span class="note-badge note-badge-pass"><svg viewBox="0 0 24 24" width="13" height="13" fill="none" stroke="currentColor" stroke-width="3" stroke-linecap="round" stroke-linejoin="round"><polyline points="20 6 9 17 4 12"></polyline></svg> ผ่าน</span>`;
        } else if (rawNote === 'รอเช็คชื่อ') {
            noteHtml = `<span class="note-badge note-badge-pending"><svg viewBox="0 0 24 24" width="13" height="13" fill="none" stroke="currentColor" stroke-width="3" stroke-linecap="round" stroke-linejoin="round"><line x1="18" y1="6" x2="6" y2="18"></line><line x1="6" y1="6" x2="18" y2="18"></line></svg> รอเช็คชื่อ</span>`;
        } else if (rawNote === 'ไม่ถูกต้อง') {
            noteHtml = `<span class="note-badge note-badge-fail"><svg viewBox="0 0 24 24" width="13" height="13" fill="none" stroke="currentColor" stroke-width="3" stroke-linecap="round" stroke-linejoin="round"><line x1="18" y1="6" x2="6" y2="18"></line><line x1="6" y1="6" x2="18" y2="18"></line></svg> ไม่ถูกต้อง</span>`;
        } else if (rawNote) {
            noteHtml = `<span>${rawNote}</span>`;
        }
        
        tr.innerHTML = `
            <td class="col-grade">${row.grade}</td>
            <td class="col-all-m">${row.allMale}</td>
            <td class="col-all-f">${row.allFemale}</td>
            <td class="col-all-total">${row.allTotal}</td>
            <td class="col-present-m">${row.presentMale}</td>
            <td class="col-present-f">${row.presentFemale}</td>
            <td class="col-present-total">${row.presentTotal}</td>
            <td class="col-present-pct">${row.presentPercent}</td>
            <td class="col-absent-m">${row.absentMale}</td>
            <td class="col-absent-f">${row.absentFemale}</td>
            <td class="col-absent-total">${row.absentTotal}</td>
            <td class="col-absent-pct">${row.absentPercent}</td>
            <td class="col-note">${noteHtml}</td>
        `;
        
        els.dailyStatsTbody.appendChild(tr);
    });
}

async function handleCopyDailyReportImage() {
    if (typeof html2canvas === 'undefined') {
        showDailyToast('ไม่สามารถใช้งานฟังก์ชันรูปภาพได้ (ไม่พบไลบรารี html2canvas)', 'error');
        return;
    }
    
    showDailyToast('กำลังสร้างรูปภาพ...', 'info', 3000);
    els.btnDailyCopyImage.disabled = true;
    
    try {
        const canvas = await html2canvas(els.dailyReportCard, {
            scale: 2,
            backgroundColor: '#ffffff',
            useCORS: true,
            logging: false
        });
        
        canvas.toBlob(async (blob) => {
            if (!blob) {
                showDailyToast('สร้างรูปภาพไม่สำเร็จ กรุณาลองใหม่อีกครั้ง', 'error');
                els.btnDailyCopyImage.disabled = false;
                return;
            }
            
            try {
                if (navigator.clipboard && window.ClipboardItem) {
                    await navigator.clipboard.write([
                        new ClipboardItem({ 'image/png': blob })
                    ]);
                    showDailyToast('📋 คัดลอกรูปภาพแล้ว! สามารถกดวาง (Ctrl+V / Paste) ใน LINE หรือแชทได้ทันที', 'success', 5000);
                } else {
                    downloadCanvasImage(canvas);
                    showDailyToast('อุปกรณ์นี้ไม่รองรับการคัดลอกรูปภาพลง Clipboard โดยตรง ระบบจึงบันทึกไฟล์ PNG ลงเครื่องให้แทนครับ', 'info', 5000);
                }
            } catch (clipErr) {
                console.warn('Clipboard write failed, downloading instead:', clipErr);
                downloadCanvasImage(canvas);
                showDailyToast('เบราว์เซอร์ไม่อนุญาตให้เขียน Clipboard โดยตรง ระบบจึงบันทึกไฟล์ PNG ให้แทนครับ', 'info', 5000);
            } finally {
                els.btnDailyCopyImage.disabled = false;
            }
        }, 'image/png');
    } catch (err) {
        showDailyToast('เกิดข้อผิดพลาด: ' + err.message, 'error');
        els.btnDailyCopyImage.disabled = false;
    }
}

async function handleSaveDailyReportImage() {
    if (typeof html2canvas === 'undefined') {
        showDailyToast('ไม่สามารถใช้งานฟังก์ชันรูปภาพได้ (ไม่พบไลบรารี html2canvas)', 'error');
        return;
    }
    
    showDailyToast('กำลังเตรียมไฟล์รูปภาพ...', 'info', 2000);
    els.btnDailySaveImage.disabled = true;
    
    try {
        const canvas = await html2canvas(els.dailyReportCard, {
            scale: 2,
            backgroundColor: '#ffffff',
            useCORS: true,
            logging: false
        });
        
        downloadCanvasImage(canvas);
        showDailyToast('💾 บันทึกไฟล์รูปภาพสำเร็จ!', 'success');
    } catch (err) {
        showDailyToast('เกิดข้อผิดพลาดในการบันทึกรูปภาพ: ' + err.message, 'error');
    } finally {
        els.btnDailySaveImage.disabled = false;
    }
}

function downloadCanvasImage(canvas) {
    const todayStr = new Date().toISOString().slice(0, 10);
    const link = document.createElement('a');
    link.download = `สถิตินักเรียนมาเรียนประจำวัน_${todayStr}.png`;
    link.href = canvas.toDataURL('image/png');
    link.click();
}

function handleCopyDailySummaryText() {
    if (!currentDailyData || !currentDailyData.rows || currentDailyData.rows.length === 0) {
        showDailyToast('ยังไม่มีข้อมูลสำหรับสรุปข้อความ', 'error');
        return;
    }
    
    const grandTotalRow = currentDailyData.rows.find(r => r.isGrandTotal || r.grade.includes('รวมทั้งหมด'));
    const dateText = currentDailyData.updatedAt || new Date().toLocaleDateString('th-TH');
    
    let text = `📊 สถิตินักเรียนมาเรียนประจำวัน\n📅 ข้อมูล ณ ${dateText}\n━━━━━━━━━━━━━━━━━━\n`;
    
    if (grandTotalRow) {
        text += `👥 นักเรียนทั้งหมด: ${grandTotalRow.allTotal} คน (ช ${grandTotalRow.allMale} / ญ ${grandTotalRow.allFemale})\n`;
        text += `✅ มาเรียน: ${grandTotalRow.presentTotal} คน (${grandTotalRow.presentPercent}%)\n`;
        text += `❌ ขาดเรียน: ${grandTotalRow.absentTotal} คน (${grandTotalRow.absentPercent}%)\n`;
    }
    
    text += `━━━━━━━━━━━━━━━━━━\n`;
    
    const kinderRow = currentDailyData.rows.find(r => r.grade.includes('รวมอนุบาล'));
    const primaryRow = currentDailyData.rows.find(r => r.grade.includes('รวมประถม') && !r.grade.includes('ต้น') && !r.grade.includes('ปลาย'));
    
    if (kinderRow) {
        text += `• รวมอนุบาล: มา ${kinderRow.presentTotal}/${kinderRow.allTotal} คน (${kinderRow.presentPercent}%) | ขาด ${kinderRow.absentTotal} คน\n`;
    }
    if (primaryRow) {
        text += `• รวมประถม: มา ${primaryRow.presentTotal}/${primaryRow.allTotal} คน (${primaryRow.presentPercent}%) | ขาด ${primaryRow.absentTotal} คน\n`;
    }
    
    if (navigator.clipboard && navigator.clipboard.writeText) {
        navigator.clipboard.writeText(text).then(() => {
            showDailyToast('💬 คัดลอกข้อความสรุปแล้ว! สามารถวางในแชท LINE ได้ทันที', 'success');
        }).catch(err => {
            showDailyToast('ไม่สามารถคัดลอกข้อความได้: ' + err.message, 'error');
        });
    } else {
        // Fallback prompt for older browsers
        window.prompt('คัดลอกข้อความด้านล่าง (Ctrl+C):', text);
    }
}


// Start
document.addEventListener('DOMContentLoaded', init);
