/**
 * Portal Karyawan - Attendance
 * Attendance/Clock In-Out endpoints
 */

function _parseDateToYMD(val) {
  if (!val) return '';
  if (val instanceof Date) {
    return Utilities.formatDate(val, 'Asia/Jakarta', 'yyyy-MM-dd');
  }
  if (typeof val === 'string' && val.length >= 10) {
    return val.substring(0, 10);
  }
  return String(val);
}

function getAttendance(userId) {
  if (!userId) {
    return { success: false, error: 'userId is required' };
  }
  
  const rows = findRows('Attendance', 'userId', userId);
  rows.forEach(r => r.date = _parseDateToYMD(r.date));
  
  // Sort by date descending
  rows.sort((a, b) => String(b.date).localeCompare(String(a.date)));
  
  return { success: true, data: rows };
}

function getTodayAttendance(userId) {
  if (!userId) {
    return { success: false, error: 'userId is required' };
  }
  
  const today = Utilities.formatDate(new Date(), 'Asia/Jakarta', 'yyyy-MM-dd');
  const allRows = getAllRows('Attendance');
  
  const todayRecord = allRows.find(row => 
    String(row.userId) === String(userId) && _parseDateToYMD(row.date) === today
  );
  
  if (todayRecord) {
    todayRecord.date = _parseDateToYMD(todayRecord.date);
    return { success: true, data: todayRecord };
  }
  
  // Return empty template
  return { 
    success: true, 
    data: {
      id: null,
      userId: userId,
      date: today,
      shift: 'Pagi',
      clockIn: '',
      clockOut: '',
      breakStart: '',
      breakEnd: '',
      overtimeStart: '',
      status: 'waiting',
      verificationPhoto: '',
      verificationLocation: '',
      verificationTimestamp: ''
    }
  };
}

function saveAttendanceData(data) {
  if (!data.userId || !data.date) {
    return { success: false, error: 'userId and date are required' };
  }
  
  // If clocking in, determine if ontime or late
  if (data.clockIn && !data.clockOut && !data.breakStart && !data.breakEnd && !data.overtimeStart) {
      // Get settings tolerance
      let tolerance = 15; // default 15 mins
      const settingsRows = getAllRows('Settings');
      const toleranceSetting = settingsRows.find(s => String(s.key) === 'lateTolerance');
      if (toleranceSetting) {
          tolerance = parseInt(toleranceSetting.value, 10) || 15;
      }
      
      // Get shift start time
      let shiftStartTimeStr = "08:00"; // fallback
      const shifts = getAllRows('Shifts');
      const userShift = shifts.find(s => String(s.name) === String(data.shift));
      
      if (userShift && userShift.startTime) {
          // Format from Date or String
          if (userShift.startTime instanceof Date) {
              const h = String(userShift.startTime.getHours()).padStart(2, '0');
              const m = String(userShift.startTime.getMinutes()).padStart(2, '0');
              shiftStartTimeStr = h + ':' + m;
          } else {
              shiftStartTimeStr = String(userShift.startTime).substring(0, 5);
          }
      }
      
      // Compare times
      const safeClockIn = String(data.clockIn).replace('.', ':');
      const safeShiftStart = String(shiftStartTimeStr).replace('.', ':');
      
      const [inH, inM] = safeClockIn.split(':').map(Number);
      const [startH, startM] = safeShiftStart.split(':').map(Number);
      
      const inMinutes = (inH || 0) * 60 + (inM || 0);
      const expectedMinutes = (startH || 0) * 60 + (startM || 0);
      
      if (inMinutes > expectedMinutes + tolerance) {
          data.status = 'Terlambat';
      } else {
          data.status = 'ontime';
      }
  }
  
  // Check if record exists for this user+date
  const allRows = getAllRows('Attendance');
  const existing = allRows.find(row => 
    String(row.userId) === String(data.userId) && _parseDateToYMD(row.date) === String(data.date)
  );
  
  if (existing && existing.id) {
    // Update existing record
    const updated = updateRow('Attendance', existing.id, data);
    return { success: true, data: updated };
  } else {
    // Create new record
    data.id = getNextId('Attendance');
    addRow('Attendance', data);
    return { success: true, data: data };
  }
}

function getAllAttendanceData() {
  const rows = getAllRows('Attendance');
  rows.forEach(r => r.date = _parseDateToYMD(r.date));
  rows.sort((a, b) => String(b.date).localeCompare(String(a.date)));
  return { success: true, data: rows };
}
