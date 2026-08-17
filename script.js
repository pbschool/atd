const state = {
    apiUrl: localStorage.getItem('apiUrl') || '',
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
    
    // Settings Page
    adminAuthSection: document.getElementById('admin-auth-section'),
    adminPassword: document.getElementById('admin-password'),
    btnAdminUnlock: document.getElementById('btn-admin-unlock'),
    settingsContentSection: document.getElementById('settings-content-section'),
    apiUrlInput: document.getElementById('api-url'),
    btnSyncSettings: document.getElementById('btn-sync-settings'),
    syncStatus: document.getElementById('sync-status')
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
    els.selectGrade.value = state.lastGrade;
    
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
    
    els.selectYear.addEventListener('change', (e) => localStorage.setItem('lastYear', e.target.value));
    els.selectGrade.addEventListener('change', (e) => localStorage.setItem('lastGrade', e.target.value));
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
    els.selectYear.innerHTML = '';
    
    const today = new Date();
    let thaiYear = today.getFullYear() + 543;
    if (today.getMonth() < 4) thaiYear -= 1; // Academic year typically starts in May
    
    // Force exactly 2 years: Current and Previous
    let yearsToShow = [thaiYear.toString(), (thaiYear - 1).toString()];

    yearsToShow.forEach(y => {
        const option = document.createElement('option');
        option.value = y;
        option.textContent = y;
        els.selectYear.appendChild(option);
    });
    
    if (state.lastYear && yearsToShow.includes(state.lastYear)) {
        els.selectYear.value = state.lastYear;
    } else {
        els.selectYear.value = yearsToShow[0];
        localStorage.setItem('lastYear', yearsToShow[0]);
    }
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
    const result = await response.json();
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
    
    if (!state.config || !state.config.links || Object.keys(state.config.links).length === 0) {
        alert('ข้อมูลกำลังโหลด หรือยังไม่ได้ตั้งค่า API_URL ใน script.js');
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
    
    // Extract month (01-12) from date
    const d = new Date(dateStr);
    const monthTab = String(d.getMonth() + 1).padStart(2, '0');
    
    state.currentSelectedUrl = targetUrl;
    state.currentSelectedMonthTab = monthTab;
    state.currentSelectedDateStr = dateStr;
    
    els.currentClassInfo.textContent = `ชั้น ${grade} | วันที่ ${dateStr}`;
    els.step1.classList.remove('active');
    els.step2.classList.add('active');
    els.loadingStudents.style.display = 'block';
    els.studentsList.innerHTML = '';
    els.saveRow.style.display = 'none';
    
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
        let currentStatus = student.status || '-';
        state.attendanceData[student.id] = currentStatus;
        
        statusBtns.forEach(btn => {
            if (btn.getAttribute('data-status') === currentStatus) {
                btn.classList.add('selected');
            }
            
            btn.addEventListener('click', (e) => {
                // Remove selected from all in this row
                statusBtns.forEach(b => b.classList.remove('selected'));
                e.target.classList.add('selected');
                
                const newStatus = e.target.getAttribute('data-status');
                state.attendanceData[student.id] = newStatus;
            });
        });
        
        els.studentsList.appendChild(clone);
    });
    
    els.saveRow.style.display = 'block';
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

// Start
document.addEventListener('DOMContentLoaded', init);
