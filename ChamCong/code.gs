/**
 * ============================================================
 * CHẤM CÔNG & NHẬT KÝ ĐI LÀM - GOOGLE APPS SCRIPT BACKEND
 * ============================================================
 * Quy chuẩn giờ làm:
 * - Giờ làm việc tiêu chuẩn: 08:00 - 17:00 (1 công)
 * - Tăng ca (OT) mặc định: Đến 20:00 (3 giờ OT)
 */

const SHEET_CHAM_CONG = 'CHAM_CONG';
const SHEET_CONFIG = 'CONFIG';

// Tiêu đề các cột trong Sheet CHAM_CONG
const HEADERS = [
  'ID',
  'Ngày',
  'Thứ',
  'Trạng thái',
  'Số công',
  'Giờ OT',
  'Giờ vào',
  'Giờ ra',
  'Ghi chú',
  'Thời gian cập nhật'
];

const APP_VERSION = 'v1.1.0';
const SCHEMA_VERSION_PROP = 'chamcong_schema_version';

/**
 * Phục vụ ứng dụng Web App
 */
function doGet(e) {
  return HtmlService
    .createHtmlOutputFromFile('index')
    .setTitle('Chấm Công & Lịch Đi Làm')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, maximum-scale=1, user-scalable=no')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/**
 * Khởi tạo cấu hình và sheet ban đầu
 */
function setup() {
  try {
    const sheet = getOrCreateLogSheet_();
    getConfigSheet_();
    return {
      success: true,
      sheet: sheet.getName(),
      timezone: SpreadsheetApp.getActive().getSpreadsheetTimeZone(),
      version: APP_VERSION
    };
  } catch (err) {
    return {
      success: false,
      error: err.toString()
    };
  }
}

/**
 * Lấy hoặc tạo sheet CHAM_CONG với định dạng đẹp
 */
function getOrCreateLogSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_CHAM_CONG);

  if (!sheet) {
    sheet = ss.insertSheet(SHEET_CHAM_CONG);
    sheet.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]);
    formatLogSheetHeader_(sheet);
    markSchemaUpToDate_();
    return sheet;
  }

  if (!isSchemaUpToDate_()) {
    checkAndMigrateSchema_(sheet);
    markSchemaUpToDate_();
  }

  return sheet;
}

function isSchemaUpToDate_() {
  return PropertiesService.getDocumentProperties().getProperty(SCHEMA_VERSION_PROP) === APP_VERSION;
}

function markSchemaUpToDate_() {
  PropertiesService.getDocumentProperties().setProperty(SCHEMA_VERSION_PROP, APP_VERSION);
}

function formatLogSheetHeader_(sheet) {
  const headerRange = sheet.getRange(1, 1, 1, HEADERS.length);
  headerRange
    .setFontWeight('bold')
    .setBackground('#1e293b')
    .setFontColor('#ffffff')
    .setHorizontalAlignment('center');
  sheet.setFrozenRows(1);

  // Chiều rộng các cột
  const colWidths = [140, 110, 100, 140, 90, 90, 90, 90, 280, 160];
  colWidths.forEach((width, i) => {
    sheet.setColumnWidth(i + 1, width);
  });

  if (sheet.getMaxRows() > 1) {
    // Format Ngày
    sheet.getRange(2, 2, sheet.getMaxRows() - 1, 1).setNumberFormat('@'); // Text YYYY-MM-DD
    // Format Căn giữa Thứ, Trạng thái, Giờ
    sheet.getRange(2, 3, sheet.getMaxRows() - 1, 1).setHorizontalAlignment('center');
    sheet.getRange(2, 4, sheet.getMaxRows() - 1, 1).setHorizontalAlignment('center');
    // Format Số công, Giờ OT
    sheet.getRange(2, 5, sheet.getMaxRows() - 1, 2).setNumberFormat('0.0#').setHorizontalAlignment('center');
    // Format Giờ vào, Giờ ra
    sheet.getRange(2, 7, sheet.getMaxRows() - 1, 2).setHorizontalAlignment('center');
  }
}

function checkAndMigrateSchema_(sheet) {
  const lastCol = Math.max(sheet.getLastColumn(), 1);
  const currentHeaders = sheet.getRange(1, 1, 1, lastCol).getValues()[0].map(v => String(v || '').trim());
  if (sheet.getLastRow() === 0 || currentHeaders.every(h => !h)) {
    sheet.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]);
    formatLogSheetHeader_(sheet);
  }
}

/**
 * Lấy hoặc tạo sheet CONFIG
 */
function getConfigSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_CONFIG);

  if (!sheet) {
    sheet = ss.insertSheet(SHEET_CONFIG);
    sheet.getRange('A1:B7').setValues([
      ['Cấu hình', 'Giá trị'],
      ['Công chuẩn / tháng', 22],
      ['Giờ bắt đầu làm', '08:00'],
      ['Giờ kết thúc làm', '17:00'],
      ['Giờ kết thúc OT mặc định', '20:00 (3h OT)'],
      ['Tên người dùng', 'Nhân viên'],
      ['Cập nhật lần cuối', new Date().toISOString()]
    ]);
    sheet.getRange('A1:B1').setFontWeight('bold').setBackground('#334155').setFontColor('#ffffff');
    sheet.setColumnWidth(1, 220);
    sheet.setColumnWidth(2, 180);
  }
  return sheet;
}

/**
 * Lấy danh sách thứ trong tuần dạng tiếng Việt
 */
function getDayOfWeekVN_(dateStr) {
  const parts = dateStr.split('-');
  const d = new Date(parseInt(parts[0], 10), parseInt(parts[1], 10) - 1, parseInt(parts[2], 10));
  const days = ['Chủ Nhật', 'Thứ Hai', 'Thứ Ba', 'Thứ Tư', 'Thứ Năm', 'Thứ Sáu', 'Thứ Bảy'];
  return days[d.getDay()];
}

/**
 * Định dạng ngày YYYY-MM-DD
 */
function formatDateKey_(d) {
  if (typeof d === 'string') {
    if (/^\d{4}-\d{2}-\d{2}$/.test(d.trim())) return d.trim();
    d = new Date(d);
  }
  const year = d.getFullYear();
  const month = String(d.getMonth() + 1).padStart(2, '0');
  const day = String(d.getDate()).padStart(2, '0');
  return `${year}-${month}-${day}`;
}

/**
 * Lấy dữ liệu chấm công trong 1 tháng kèm thống kê
 * TỐI ƯU HÓA HIỆU NĂNG: Không đọc toàn bộ bảng tính, chỉ quét cột Ngày và đọc đúng các dòng thuộc tháng
 * @param {number} year - Ví dụ 2026
 * @param {number} month - 1 đến 12
 */
function getMonthAttendance(year, month) {
  try {
    const sheet = getOrCreateLogSheet_();
    const lastRow = sheet.getLastRow();
    
    const yearNum = parseInt(year, 10);
    const monthNum = parseInt(month, 10);
    const monthPrefix = `${yearNum}-${String(monthNum).padStart(2, '0')}`;

    const items = {};
    let totalWorkUnits = 0;
    let totalOtHours = 0;
    let workDaysCount = 0;
    let otDaysCount = 0;
    let leavePaidDays = 0;
    let leaveUnpaidDays = 0;
    let holidayDays = 0;

    if (lastRow >= 2) {
      // 1. Chỉ đọc duy nhất 1 cột Ngày (Cột B: từ dòng 2 đến lastRow) để tìm các dòng khớp
      const dateColumnValues = sheet.getRange(2, 2, lastRow - 1, 1).getValues();
      const matchedRows = [];

      for (let i = 0; i < dateColumnValues.length; i++) {
        let dateVal = dateColumnValues[i][0];
        if (!dateVal) continue;

        let dateStr = '';
        if (dateVal instanceof Date) {
          dateStr = formatDateKey_(dateVal);
        } else {
          dateStr = String(dateVal).trim();
          if (dateStr.length > 10) dateStr = dateStr.substring(0, 10);
        }

        if (dateStr.startsWith(monthPrefix)) {
          matchedRows.push({
            rowIndex: i + 2,
            dateStr: dateStr
          });
        }
      }

      // 2. Chỉ đọc dữ liệu của các dòng thuộc tháng được chọn
      if (matchedRows.length > 0) {
        const minRow = matchedRows[0].rowIndex;
        const maxRow = matchedRows[matchedRows.length - 1].rowIndex;
        const rowSpan = maxRow - minRow + 1;

        // Nếu các dòng nằm trong khoảng hẹp (<= 60 dòng): Đọc duy nhất 1 range từ minRow đến maxRow
        if (rowSpan <= Math.max(matchedRows.length * 2, 60)) {
          const rangeData = sheet.getRange(minRow, 1, rowSpan, HEADERS.length).getValues();
          const targetSet = new Set(matchedRows.map(m => m.rowIndex));

          for (let k = 0; k < rangeData.length; k++) {
            const currentRowIndex = minRow + k;
            if (targetSet.has(currentRowIndex)) {
              processRow_(rangeData[k], currentRowIndex);
            }
          }
        } else {
          // Trường hợp các dòng rải rác cách xa nhau: đọc từng dòng cần thiết
          matchedRows.forEach(m => {
            const rowData = sheet.getRange(m.rowIndex, 1, 1, HEADERS.length).getValues()[0];
            processRow_(rowData, m.rowIndex);
          });
        }
      }
    }

    function processRow_(row, rowIndex) {
      let dateVal = row[1];
      let dateStr = (dateVal instanceof Date) ? formatDateKey_(dateVal) : String(dateVal).trim().substring(0, 10);

      const status = String(row[3] || '').trim();
      const workUnits = parseFloat(row[4]) || 0;
      const otHours = parseFloat(row[5]) || 0;
      const checkIn = String(row[6] || '').trim();
      const checkOut = String(row[7] || '').trim();
      const note = String(row[8] || '').trim();
      const updatedAt = String(row[9] || '').trim();
      const id = String(row[0] || `cc_${dateStr}`);

      items[dateStr] = {
        id: id,
        date: dateStr,
        dayOfWeek: row[2] || getDayOfWeekVN_(dateStr),
        status: status,
        workUnits: workUnits,
        otHours: otHours,
        checkIn: checkIn,
        checkOut: checkOut,
        note: note,
        updatedAt: updatedAt,
        rowIndex: rowIndex
      };

      totalWorkUnits += workUnits;
      totalOtHours += otHours;
      if (workUnits > 0) workDaysCount++;
      if (otHours > 0) otDaysCount++;

      if (status.includes('Nghỉ phép') || status === 'Nghỉ phép') {
        leavePaidDays++;
      } else if (status.includes('Nghỉ không lương') || status === 'Nghỉ không lương') {
        leaveUnpaidDays++;
      } else if (status.includes('Lễ') || status.includes('Nghỉ tuần') || status === 'Nghỉ') {
        holidayDays++;
      }
    }

    // Tính tổng số ngày trong tháng
    const daysInMonth = new Date(yearNum, monthNum, 0).getDate();

    return {
      success: true,
      year: yearNum,
      month: monthNum,
      daysInMonth: daysInMonth,
      records: items,
      summary: {
        totalWorkUnits: Math.round(totalWorkUnits * 100) / 100,
        totalOtHours: Math.round(totalOtHours * 100) / 100,
        workDaysCount: workDaysCount,
        otDaysCount: otDaysCount,
        leavePaidDays: leavePaidDays,
        leaveUnpaidDays: leaveUnpaidDays,
        holidayDays: holidayDays,
        recordedDays: Object.keys(items).length
      }
    };
  } catch (err) {
    return {
      success: false,
      error: err.toString()
    };
  }
}

/**
 * Lưu hoặc cập nhật chấm công của 1 ngày (Upsert)
 * @param {Object} entry - { date, status, workUnits, otHours, checkIn, checkOut, note }
 */
function saveAttendance(entry) {
  try {
    if (!entry || !entry.date) {
      return { success: false, error: 'Thiếu thông tin ngày chấm công' };
    }

    const sheet = getOrCreateLogSheet_();
    const dateStr = formatDateKey_(entry.date);
    const dayOfWeek = entry.dayOfWeek || getDayOfWeekVN_(dateStr);
    const status = entry.status || 'Đi làm';
    const workUnits = entry.workUnits !== undefined ? Number(entry.workUnits) : (status.includes('Nghỉ') ? 0 : 1);
    const otHours = Number(entry.otHours || 0);
    
    // Mặc định giờ vào 08:00, giờ ra 17:00 nếu đi làm bình thường, 20:00 nếu OT 3h
    let checkIn = entry.checkIn;
    let checkOut = entry.checkOut;
    if (checkIn === undefined || checkIn === null) {
      checkIn = (workUnits > 0) ? '08:00' : '';
    }
    if (checkOut === undefined || checkOut === null) {
      if (otHours >= 3) checkOut = '20:00';
      else if (otHours > 0) checkOut = `${17 + Math.floor(otHours)}:${(otHours % 1) * 60 === 30 ? '30' : '00'}`;
      else if (workUnits === 0.5) checkOut = '12:00';
      else if (workUnits > 0) checkOut = '17:00';
      else checkOut = '';
    }

    const note = entry.note || '';
    const nowStr = Utilities.formatDate(new Date(), SpreadsheetApp.getActive().getSpreadsheetTimeZone(), 'yyyy-MM-dd HH:mm:ss');
    const id = entry.id || `cc_${dateStr}`;

    const lastRow = sheet.getLastRow();
    let targetRow = -1;

    if (lastRow >= 2) {
      const dates = sheet.getRange(2, 2, lastRow - 1, 1).getValues();
      for (let i = 0; i < dates.length; i++) {
        let curDate = dates[i][0];
        let curDateStr = (curDate instanceof Date) ? formatDateKey_(curDate) : String(curDate).trim();
        if (curDateStr === dateStr) {
          targetRow = i + 2;
          break;
        }
      }
    }

    const rowData = [
      id,
      dateStr,
      dayOfWeek,
      status,
      workUnits,
      otHours,
      checkIn,
      checkOut,
      note,
      nowStr
    ];

    if (targetRow > 0) {
      // Cập nhật dòng hiện có
      sheet.getRange(targetRow, 1, 1, HEADERS.length).setValues([rowData]);
    } else {
      // Thêm dòng mới
      sheet.appendRow(rowData);
      targetRow = sheet.getLastRow();
    }

    return {
      success: true,
      record: {
        id: id,
        date: dateStr,
        dayOfWeek: dayOfWeek,
        status: status,
        workUnits: workUnits,
        otHours: otHours,
        checkIn: checkIn,
        checkOut: checkOut,
        note: note,
        updatedAt: nowStr,
        rowIndex: targetRow
      }
    };
  } catch (err) {
    return {
      success: false,
      error: err.toString()
    };
  }
}

/**
 * Lưu hàng loạt bản ghi chấm công (Batch Save)
 * @param {Array<Object>} entries
 */
function batchSaveAttendance(entries) {
  try {
    if (!entries || !Array.isArray(entries) || entries.length === 0) {
      return { success: false, error: 'Không có dữ liệu chấm công để lưu' };
    }

    const sheet = getOrCreateLogSheet_();
    const lastRow = sheet.getLastRow();
    const dateMap = {};

    if (lastRow >= 2) {
      const dates = sheet.getRange(2, 2, lastRow - 1, 1).getValues();
      for (let i = 0; i < dates.length; i++) {
        let curDate = dates[i][0];
        let curDateStr = (curDate instanceof Date) ? formatDateKey_(curDate) : String(curDate).trim();
        if (curDateStr) {
          dateMap[curDateStr] = i + 2;
        }
      }
    }

    const nowStr = Utilities.formatDate(new Date(), SpreadsheetApp.getActive().getSpreadsheetTimeZone(), 'yyyy-MM-dd HH:mm:ss');
    const newRows = [];
    const updatedRecords = [];

    entries.forEach(entry => {
      const dateStr = formatDateKey_(entry.date);
      const dayOfWeek = entry.dayOfWeek || getDayOfWeekVN_(dateStr);
      const status = entry.status || 'Đi làm';
      const workUnits = entry.workUnits !== undefined ? Number(entry.workUnits) : (status.includes('Nghỉ') ? 0 : 1);
      const otHours = Number(entry.otHours || 0);
      const checkIn = entry.checkIn || '08:00';
      const checkOut = entry.checkOut || (otHours > 0 ? '20:00' : '17:00');
      const note = entry.note || '';
      const id = entry.id || `cc_${dateStr}`;

      const rowData = [
        id,
        dateStr,
        dayOfWeek,
        status,
        workUnits,
        otHours,
        checkIn,
        checkOut,
        note,
        nowStr
      ];

      if (dateMap[dateStr]) {
        // Cập nhật dòng cũ
        sheet.getRange(dateMap[dateStr], 1, 1, HEADERS.length).setValues([rowData]);
      } else {
        // Thêm mới
        newRows.push(rowData);
      }

      updatedRecords.push({
        id: id,
        date: dateStr,
        dayOfWeek: dayOfWeek,
        status: status,
        workUnits: workUnits,
        otHours: otHours,
        checkIn: checkIn,
        checkOut: checkOut,
        note: note,
        updatedAt: nowStr
      });
    });

    if (newRows.length > 0) {
      sheet.getRange(sheet.getLastRow() + 1, 1, newRows.length, HEADERS.length).setValues(newRows);
    }

    return {
      success: true,
      count: entries.length,
      records: updatedRecords
    };
  } catch (err) {
    return {
      success: false,
      error: err.toString()
    };
  }
}

/**
 * Xóa bản ghi chấm công của 1 ngày
 * @param {string} dateStr - 'YYYY-MM-DD'
 */
function deleteAttendance(dateStr) {
  try {
    const cleanDate = formatDateKey_(dateStr);
    const sheet = getOrCreateLogSheet_();
    const lastRow = sheet.getLastRow();

    if (lastRow < 2) {
      return { success: true, message: 'Bảng tính rỗng' };
    }

    const dates = sheet.getRange(2, 2, lastRow - 1, 1).getValues();
    for (let i = 0; i < dates.length; i++) {
      let curDate = dates[i][0];
      let curDateStr = (curDate instanceof Date) ? formatDateKey_(curDate) : String(curDate).trim();
      if (curDateStr === cleanDate) {
        sheet.deleteRow(i + 2);
        return { success: true, date: cleanDate };
      }
    }

    return { success: true, message: 'Không tìm thấy bản ghi để xóa' };
  } catch (err) {
    return {
      success: false,
      error: err.toString()
    };
  }
}

/**
 * Lấy báo cáo thống kê cả năm (12 tháng)
 * @param {number} year
 */
function getYearStats(year) {
  try {
    const yearNum = parseInt(year, 10);
    const sheet = getOrCreateLogSheet_();
    const lastRow = sheet.getLastRow();

    const monthlyStats = [];
    for (let m = 1; m <= 12; m++) {
      monthlyStats.push({
        month: m,
        workUnits: 0,
        otHours: 0,
        workDays: 0,
        leaveDays: 0
      });
    }

    if (lastRow >= 2) {
      const data = sheet.getRange(2, 1, lastRow - 1, HEADERS.length).getValues();
      data.forEach(row => {
        let dateVal = row[1];
        if (!dateVal) return;

        let dateStr = (dateVal instanceof Date) ? formatDateKey_(dateVal) : String(dateVal).trim();
        if (dateStr.startsWith(`${yearNum}-`)) {
          const mIndex = parseInt(dateStr.substring(5, 7), 10) - 1;
          if (mIndex >= 0 && mIndex < 12) {
            const workUnits = parseFloat(row[4]) || 0;
            const otHours = parseFloat(row[5]) || 0;
            const status = String(row[3] || '');

            monthlyStats[mIndex].workUnits += workUnits;
            monthlyStats[mIndex].otHours += otHours;
            if (workUnits > 0) monthlyStats[mIndex].workDays++;
            if (status.includes('Nghỉ')) monthlyStats[mIndex].leaveDays++;
          }
        }
      });
    }

    return {
      success: true,
      year: yearNum,
      months: monthlyStats
    };
  } catch (err) {
    return {
      success: false,
      error: err.toString()
    };
  }
}
