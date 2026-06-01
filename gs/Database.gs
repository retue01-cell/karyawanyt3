/**
 * Portal Karyawan - Database Helper
 * Generic CRUD operations for Google Sheets
 * 
 * PENTING: Ganti SPREADSHEET_ID dengan ID spreadsheet kamu
 */

// ========== KONFIGURASI ==========
const SPREADSHEET_ID = '1K8ZogDZS96LlSPSqf7J0eH25uvSyq6oHxgzlWWkBJ_4';

// Cache spreadsheet reference
let _spreadsheet = null;

function getSpreadsheet() {
  if (!_spreadsheet) {
    _spreadsheet = SpreadsheetApp.openById(SPREADSHEET_ID);
  }
  return _spreadsheet;
}

function getSheet(sheetName) {
  const ss = getSpreadsheet();
  let sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
  }
  return sheet;
}

// ========== INIT DATABASE ==========
function initDatabase() {
  const sheets = {
    'Users': ['id', 'name', 'email', 'password', 'role', 'avatar', 'createdAt'],
    'Employees': ['id', 'name', 'email', 'department', 'position', 'shift', 'status', 'joinDate', 'avatar', 'password'],
    'Attendance': ['id', 'userId', 'date', 'shift', 'clockIn', 'clockOut', 'breakStart', 'breakEnd', 'overtimeStart', 'status', 'verificationPhoto', 'verificationLocation', 'verificationTimestamp'],
    'Journals': ['id', 'userId', 'date', 'tasks', 'achievements', 'obstacles', 'plan', 'photo', 'updatedAt'],
    'Leaves': ['id', 'userId', 'type', 'typeLabel', 'startDate', 'endDate', 'duration', 'reason', 'status', 'appliedAt'],
    'Izin': ['id', 'userId', 'type', 'typeLabel', 'date', 'duration', 'reason', 'status', 'hasAttachment', 'verificationPhoto', 'verificationLocation', 'verificationTimestamp', 'appliedAt'],
    'Settings': ['key', 'value'],
    'Shifts': ['id', 'name', 'startTime', 'endTime']
  };

  const ss = getSpreadsheet();
  
  Object.entries(sheets).forEach(([sheetName, headers]) => {
    let sheet = ss.getSheetByName(sheetName);
    if (!sheet) {
      sheet = ss.insertSheet(sheetName);
      sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
      sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');
      sheet.setFrozenRows(1);
    } else {
      // Periksa dan perbaiki header jika perlu
      const existingHeaders = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
      // Jika header tidak lengkap, timpa dengan headers yang benar
      if (existingHeaders.length !== headers.length || !existingHeaders.includes('plan')) {
        // Hapus isi baris pertama, tulis ulang headers
        sheet.getRange(1, 1, 1, sheet.getLastColumn()).clear();
        sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
        sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');
        sheet.setFrozenRows(1);
      }
    }
  });

  seedDefaultData();
  
  try {
    setupDailyTrigger();
  } catch (e) {
    console.error("Gagal menginisialisasi trigger harian:", e);
  }
  
  return { success: true, message: 'Database initialized successfully' };
}

function repairDatabase() {
  const ss = getSpreadsheet();
  
  // Fix Employees headers
  const empSheet = getSheet('Employees');
  const expectedEmpHeaders = ['id', 'name', 'email', 'department', 'position', 'shift', 'status', 'joinDate', 'avatar', 'password'];
  empSheet.getRange(1, 1, 1, expectedEmpHeaders.length).setValues([expectedEmpHeaders]);
  empSheet.getRange(1, 1, 1, expectedEmpHeaders.length).setFontWeight('bold');
  empSheet.setFrozenRows(1);

  // Fix Attendance headers
  const attSheet = getSheet('Attendance');
  const attHeaders = ['id', 'userId', 'date', 'shift', 'clockIn', 'clockOut', 'breakStart', 'breakEnd', 'overtimeStart', 'status', 'verificationPhoto', 'verificationLocation', 'verificationTimestamp'];
  attSheet.getRange(1, 1, 1, attHeaders.length).setValues([attHeaders]);
  attSheet.getRange(1, 1, 1, attHeaders.length).setFontWeight('bold');
  attSheet.setFrozenRows(1);
  
  // Fix Journals headers - pastikan ada kolom 'plan'
  const journalSheet = getSheet('Journals');
  const journalHeaders = ['id', 'userId', 'date', 'tasks', 'achievements', 'obstacles', 'plan', 'photo', 'updatedAt'];
  journalSheet.getRange(1, 1, 1, journalHeaders.length).setValues([journalHeaders]);
  journalSheet.getRange(1, 1, 1, journalHeaders.length).setFontWeight('bold');
  journalSheet.setFrozenRows(1);
  
  return { success: true, message: 'Database headers repaired successfully.' };
}

function seedDefaultData() {
  // Users
  const usersSheet = getSheet('Users');
  if (usersSheet.getLastRow() <= 1) {
    usersSheet.appendRow([1, 'Admin User', 'admin@company.com', 'admin123', 'admin', 'https://ui-avatars.com/api/?name=Admin&background=F59E0B&color=fff', new Date().toISOString()]);
    usersSheet.appendRow([2, 'Dewi Karyawan', 'karyawan@company.com', 'karyawan123', 'karyawan', 'https://ui-avatars.com/api/?name=Dewi&background=3B82F6&color=fff', new Date().toISOString()]);
  }

  // Shifts
  const shiftsSheet = getSheet('Shifts');
  if (shiftsSheet.getLastRow() <= 1) {
    shiftsSheet.appendRow([1, 'Pagi', '08:00', '17:00']);
    shiftsSheet.appendRow([2, 'Siang', '14:00', '23:00']);
    shiftsSheet.appendRow([3, 'Malam', '23:00', '08:00']);
  }

  // Settings
  const settingsSheet = getSheet('Settings');
  if (settingsSheet.getLastRow() <= 1) {
    settingsSheet.appendRow(['company_name', 'Portal Karyawan']);
    settingsSheet.appendRow(['company_logo', '']);
  }

  // Employees
  const empSheet = getSheet('Employees');
  if (empSheet.getLastRow() <= 1) {
    const employees = [
      [1, 'Ahmad Rizky', 'ahmad@company.com', 'IT', 'Developer', 'Pagi', 'active', '2024-01-15', 'https://ui-avatars.com/api/?name=Ahmad&background=3B82F6&color=fff'],
      [2, 'Budi Santoso', 'budi@company.com', 'HR', 'HR Manager', 'Pagi', 'active', '2023-06-01', 'https://ui-avatars.com/api/?name=Budi&background=10B981&color=fff'],
      [3, 'Citra Dewi', 'citra@company.com', 'Finance', 'Accountant', 'Pagi', 'on-leave', '2024-03-10', 'https://ui-avatars.com/api/?name=Citra&background=F59E0B&color=fff'],
      [4, 'Dedi Pratama', 'dedi@company.com', 'Marketing', 'Marketing Staff', 'Siang', 'active', '2024-02-20', 'https://ui-avatars.com/api/?name=Dedi&background=EF4444&color=fff'],
      [5, 'Eka Putri', 'eka@company.com', 'IT', 'UI/UX Designer', 'Pagi', 'active', '2024-01-05', 'https://ui-avatars.com/api/?name=Eka&background=8B5CF6&color=fff'],
      [6, 'Fajar Nugraha', 'fajar@company.com', 'Operations', 'Supervisor', 'Malam', 'inactive', '2023-09-12', 'https://ui-avatars.com/api/?name=Fajar&background=6B7280&color=fff']
    ];
    employees.forEach(emp => empSheet.appendRow(emp));
  }
}

// ========== GENERIC CRUD ==========

function getAllRows(sheetName) {
  const sheet = getSheet(sheetName);
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return [];
  const data = sheet.getRange(1, 1, lastRow, sheet.getLastColumn()).getDisplayValues();
  const headers = data[0];
  const rows = [];
  for (let i = 1; i < data.length; i++) {
    const obj = {};
    headers.forEach((header, j) => { obj[header] = data[i][j]; });
    rows.push(obj);
  }
  return rows;
}

function findRows(sheetName, column, value) {
  const allRows = getAllRows(sheetName);
  return allRows.filter(row => String(row[column]) === String(value));
}

function findRow(sheetName, column, value) {
  const rows = findRows(sheetName, column, value);
  return rows.length > 0 ? rows[0] : null;
}

function addRow(sheetName, data) {
  const sheet = getSheet(sheetName);
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const row = headers.map(header => data[header] !== undefined ? data[header] : '');
  sheet.appendRow(row);
  return data;
}

function updateRow(sheetName, id, data) {
  const sheet = getSheet(sheetName);
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return null;
  const allData = sheet.getRange(1, 1, lastRow, sheet.getLastColumn()).getDisplayValues();
  const headers = allData[0];
  const idColIndex = headers.indexOf('id');
  if (idColIndex === -1) return null;
  for (let i = 1; i < allData.length; i++) {
    if (String(allData[i][idColIndex]) === String(id)) {
      headers.forEach((header, j) => {
        if (data[header] !== undefined && header !== 'id') {
          sheet.getRange(i + 1, j + 1).setValue(data[header]);
        }
      });
      return { ...rowToObject(headers, allData[i]), ...data };
    }
  }
  return null;
}

function deleteRow(sheetName, id) {
  const sheet = getSheet(sheetName);
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return false;
  const allData = sheet.getRange(1, 1, lastRow, sheet.getLastColumn()).getDisplayValues();
  const headers = allData[0];
  const idColIndex = headers.indexOf('id');
  if (idColIndex === -1) return false;
  for (let i = 1; i < allData.length; i++) {
    if (String(allData[i][idColIndex]) === String(id)) {
      sheet.deleteRow(i + 1);
      return true;
    }
  }
  return false;
}

function updateRowByColumn(sheetName, column, value, data) {
  const sheet = getSheet(sheetName);
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return null;
  const allData = sheet.getRange(1, 1, lastRow, sheet.getLastColumn()).getDisplayValues();
  const headers = allData[0];
  const colIndex = headers.indexOf(column);
  if (colIndex === -1) return null;
  for (let i = 1; i < allData.length; i++) {
    if (String(allData[i][colIndex]) === String(value)) {
      headers.forEach((header, j) => {
        if (data[header] !== undefined) sheet.getRange(i + 1, j + 1).setValue(data[header]);
      });
      return { ...rowToObject(headers, allData[i]), ...data };
    }
  }
  return null;
}

function getNextId(sheetName) {
  const sheet = getSheet(sheetName);
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return 1;
  const allData = sheet.getRange(1, 1, lastRow, sheet.getLastColumn()).getValues();
  const headers = allData[0];
  const idColIndex = headers.indexOf('id');
  if (idColIndex === -1) return Date.now();
  let maxId = 0;
  for (let i = 1; i < allData.length; i++) {
    const id = Number(allData[i][idColIndex]);
    if (id > maxId) maxId = id;
  }
  return maxId + 1;
}

function rowToObject(headers, row) {
  const obj = {};
  headers.forEach((header, i) => { obj[header] = row[i]; });
  return obj;
}