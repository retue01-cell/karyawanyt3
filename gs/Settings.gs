/**
 * Portal Karyawan - Settings
 * Company settings, shifts, and schedule endpoints
 */

// ========== SETTINGS ==========

function getSettingsData() {
  const rows = getAllRows('Settings');
  // Convert to key-value object
  const settings = {};
  rows.forEach(row => {
    settings[row.key] = row.value;
  });
  return { success: true, data: settings };
}

function saveSettingData(key, value) {
  if (!key) {
    return { success: false, error: 'key is required' };
  }
  
  const sheet = getSheet('Settings');
  let lastRow = sheet.getLastRow();
  
  // Reconstruct headers if the sheet was accidentally wiped
  if (lastRow === 0 || sheet.getRange(1, 1).getValue() !== 'key') {
    sheet.insertRowBefore(1);
    sheet.getRange(1, 1, 1, 2).setValues([['key', 'value']]);
    sheet.getRange(1, 1, 1, 2).setFontWeight('bold');
    sheet.setFrozenRows(1);
    lastRow = sheet.getLastRow();
  }
  
  if (lastRow > 1) {
    const data = sheet.getRange(1, 1, lastRow, 2).getValues();
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]) === String(key)) {
        // Update existing setting
        sheet.getRange(i + 1, 2).setValue(value);
        if (key.startsWith('shift_schedule_')) {
           autoUpdateDailyShifts();
        }
        return { success: true, data: { key: key, value: value } };
      }
    }
  }
  
  // Add new setting
  sheet.appendRow([key, value]);
  
  if (key.startsWith('shift_schedule_')) {
     autoUpdateDailyShifts();
  }
  
  return { success: true, data: { key: key, value: value } };
}

/**
 * Automatisasi: Membaca jadwal Admin untuk hari ini dan memaksanya ke kolom 'shift' di tabel Employees.
 * Fungsi ini bisa dipanggil manual, via trigger jam 00:00, atau otomatis sesaat setelah Simpan Jadwal.
 */
function autoUpdateDailyShifts() {
  const jakartaDateStr = Utilities.formatDate(new Date(), "Asia/Jakarta", "yyyy-MM-dd");
  const dateParts = jakartaDateStr.split('-');
  const currentYear = parseInt(dateParts[0], 10);
  
  // NOTE: JS frontend used `date.getMonth()` which is 0-indexed.
  // Utilities output `03` for March, so we subtract 1 safely:
  const currentMonth = parseInt(dateParts[1], 10) - 1; 
  const currentDay = parseInt(dateParts[2], 10);
  
  const key = `shift_schedule_${currentYear}-${currentMonth}`;
  console.log(`[ShiftSync] Mulai sinkronisasi untuk Hari ini: ${currentDay}, Kunci Bulan: ${key}`);
  
  const settingsRows = getAllRows('Settings');
  let monthScheduleStr = null;
  settingsRows.forEach(row => {
    if (String(row.key) === key) {
      monthScheduleStr = row.value;
    }
  });
  
  if (!monthScheduleStr) {
    console.log(`[ShiftSync] Gagal: Tidak ada string pengaturan untuk ${key}`);
    return { success: false, error: 'Belum ada jadwal bulan ini' };
  }
  
  let schedules;
  try {
    schedules = JSON.parse(monthScheduleStr);
    console.log(`[ShiftSync] Berhasil parsing JSON jadwal. Mengandung ID karyawan:`, Object.keys(schedules));
  } catch (e) {
    console.log(`[ShiftSync] Gagal JSON parse:`, e);
    return { success: false, error: 'Gagal membaca format jadwal' };
  }
  
  const employeesRows = getAllRows('Employees');
  let updatedCount = 0;
  
  employeesRows.forEach(emp => {
    const stringId = String(emp.id);
    console.log(`[ShiftSync] Mengecek Karyawan baris Excel ID ${stringId} | Nama: ${emp.name} | Shift di Excel saat ini: '${emp.shift}'`);
    if (schedules[stringId] && schedules[stringId][currentDay]) {
      const assignedShift = schedules[stringId][currentDay];
      console.log(`[ShiftSync] >> Jadwal di Kalender Admin: '${assignedShift}'`);
      if (assignedShift && String(emp.shift).trim() !== String(assignedShift).trim()) {
        console.log(`[ShiftSync] >> BEDA! Menimpa baris Excel ID ${emp.id} ke shift '${assignedShift}' via updateRow...`);
        updateRow('Employees', emp.id, { shift: assignedShift });
        updatedCount++;
      } else {
        console.log(`[ShiftSync] >> SAMA (atau kosong). Dilewati.`);
      }
    } else {
      console.log(`[ShiftSync] >> Tidak ada jadwal khusus untuk karyawan ini di tanggal ${currentDay}.`);
    }
  });
  
  console.log(`[ShiftSync] SELESAI. Total Karyawan Diupdate: ${updatedCount}`);
  return { success: true, message: `Berhasil sinkronisasi fisik shift hari ini untuk ${updatedCount} karyawan` };
}

/**
 * Automatisasi: Memasang Trigger Time-Driven eksternal untuk autoUpdateDailyShifts
 * agar skrip tersebut berlari sendiri secara rahasia tiap tengah malam (00:00).
 */
function setupDailyTrigger() {
  // Hapus semua trigger autoUpdateDailyShifts yang sudah ada agar tidak bentrok
  const triggers = ScriptApp.getProjectTriggers();
  for (let i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === 'autoUpdateDailyShifts') {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }

  // Bikin trigger baru untuk menyala setiap hari di jam 00:00 (Tengah Malam)
  ScriptApp.newTrigger('autoUpdateDailyShifts')
           .timeBased()
           .everyDays(1)
           .atHour(0)
           .create();

  return { success: true, message: 'Weker Shift Otomatis (Trigger 00:00) berhasil dipasang!' };
}

// ========== SHIFTS ==========

function getShiftsData() {
  const rows = getAllRows('Shifts');
  // Google Sheets converts "08:00" to Date objects.
  // Convert them back to "HH:mm" strings.
  const fixed = rows.map(row => {
    return {
      ...row,
      startTime: formatTimeValue(row.startTime),
      endTime: formatTimeValue(row.endTime)
    };
  });
  return { success: true, data: fixed };
}

/**
 * Convert a Sheets time value (Date or string) to "HH:mm" format.
 */
function formatTimeValue(val) {
  if (!val || val === '' || val === null) return '';
  // Already a proper string
  if (typeof val === 'string' && /^\d{2}:\d{2}$/.test(val)) return val;
  // If it's a Date object from Sheets
  if (val instanceof Date) {
    const h = String(val.getHours()).padStart(2, '0');
    const m = String(val.getMinutes()).padStart(2, '0');
    return h + ':' + m;
  }
  // ISO string fallback
  const str = String(val);
  if (str.includes('T')) {
    try {
      const d = new Date(str);
      const h = String(d.getHours()).padStart(2, '0');
      const m = String(d.getMinutes()).padStart(2, '0');
      return h + ':' + m;
    } catch(e) { return ''; }
  }
  return str;
}

function addShiftData(data) {
  if (!data.name) {
    return { success: false, error: 'Shift name is required' };
  }
  
  data.id = getNextId('Shifts');
  if (!data.startTime) data.startTime = '09:00';
  if (!data.endTime) data.endTime = '18:00';
  
  addRow('Shifts', data);
  return { success: true, data: data };
}

function updateShiftData(id, data) {
  if (!id) {
    return { success: false, error: 'id is required' };
  }
  
  const updated = updateRow('Shifts', id, data);
  if (updated) {
    return { success: true, data: updated };
  }
  return { success: false, error: 'Shift not found' };
}

function deleteShiftData(id) {
  if (!id) {
    return { success: false, error: 'id is required' };
  }
  
  const deleted = deleteRow('Shifts', id);
  if (deleted) {
    return { success: true, data: { id: id } };
  }
  return { success: false, error: 'Shift not found' };
}

// ========== SCHEDULE ==========

function getScheduleData(month, year) {
  const key = `schedule_${year}_${month}`;
  const settings = getAllRows('Settings');
  const entry = settings.find(s => String(s.key) === key);
  
  if (entry) {
    try {
      return { success: true, data: JSON.parse(entry.value) };
    } catch (e) {
      return { success: true, data: {} };
    }
  }
  
  return { success: true, data: {} };
}

function saveScheduleData(data) {
  if (!data.month || !data.year) {
    return { success: false, error: 'month and year are required' };
  }
  
  const key = `schedule_${data.year}_${data.month}`;
  const scheduleData = data.schedule || {};
  
  return saveSettingData(key, JSON.stringify(scheduleData));
}
