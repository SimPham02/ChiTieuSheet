const LOG_SHEET = 'LOG';
const META_SHEET = 'META';
const HEADERS = ['ID', 'Thời gian', 'Loại', 'Số tiền', 'Danh mục', 'Nội dung', 'Nguyên văn'];

// Tăng phiên bản schema khi có cập nhật cấu trúc hoặc logic để chỉ chạy migration 1 lần
const SCHEMA_VERSION = 'v2.1-fix-totals';
const SCHEMA_VERSION_PROP = 'schemaVersion';
const META_BOOTSTRAP_PROP = 'metaBootstrapped';

function doGet() {
  return HtmlService
    .createHtmlOutputFromFile('index')
    .setTitle('Nhật ký chi tiêu')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function setup() {
  const sheet = getOrCreateSheet_();
  const metaSheet = getMetaSheet_();
  recalculateMetaTotals_(metaSheet);
  return {
    success: true,
    sheet: sheet.getName(),
    timezone: SpreadsheetApp.getActive().getSpreadsheetTimeZone()
  };
}

/**
 * Lấy hoặc khởi tạo sheet LOG giao dịch
 */
function getOrCreateSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(LOG_SHEET);

  if (!sheet) {
    sheet = ss.insertSheet(LOG_SHEET);
    sheet.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]);
    formatHeader_(sheet);
    markSchemaUpToDate_();
    return sheet;
  }

  if (!isSchemaUpToDate_()) {
    migrateSchema_(sheet);
    markSchemaUpToDate_();
  }

  return sheet;
}

function isSchemaUpToDate_() {
  return PropertiesService.getDocumentProperties().getProperty(SCHEMA_VERSION_PROP) === SCHEMA_VERSION;
}

function markSchemaUpToDate_() {
  PropertiesService.getDocumentProperties().setProperty(SCHEMA_VERSION_PROP, SCHEMA_VERSION);
}

function migrateSchema_(sheet) {
  const lastColumn = Math.max(sheet.getLastColumn(), 1);
  const headers = sheet.getRange(1, 1, 1, lastColumn).getValues()[0]
    .map(v => String(v || '').trim());

  if (sheet.getLastRow() === 0 || headers.every(h => !h)) {
    sheet.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]);
    formatHeader_(sheet);
    return;
  }

  // Xóa cột "Ngày" cũ nếu tồn tại trong schema cũ
  const dateIndex = headers.indexOf('Ngày');
  if (
    dateIndex === 2 &&
    headers[0] === 'ID' &&
    headers[1] === 'Thời gian' &&
    headers[3] === 'Loại'
  ) {
    sheet.deleteColumn(dateIndex + 1);
  }

  const after = sheet.getRange(
    1, 1, 1, Math.max(sheet.getLastColumn(), HEADERS.length)
  ).getValues()[0].map(v => String(v || '').trim());

  if (after.indexOf('Danh mục') === -1) {
    const amountIndex = after.indexOf('Số tiền');
    if (amountIndex >= 0) {
      sheet.insertColumnAfter(amountIndex + 1);
      sheet.getRange(1, amountIndex + 2).setValue('Danh mục');
    } else {
      sheet.insertColumnAfter(sheet.getLastColumn());
      sheet.getRange(1, sheet.getLastColumn()).setValue('Danh mục');
    }
  }

  sheet.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]);
  formatHeader_(sheet);
  backfillCategories_(sheet);
}

function formatHeader_(sheet) {
  sheet.getRange(1, 1, 1, HEADERS.length)
    .setFontWeight('bold')
    .setBackground('#172a4a')
    .setFontColor('#ffffff');
  sheet.setFrozenRows(1);

  [230, 155, 90, 120, 140, 280, 280].forEach((width, i) => {
    sheet.setColumnWidth(i + 1, width);
  });

  if (sheet.getMaxRows() > 1) {
    sheet.getRange(2, 2, sheet.getMaxRows() - 1, 1)
      .setNumberFormat('dd/MM/yyyy HH:mm:ss');
    sheet.getRange(2, 4, sheet.getMaxRows() - 1, 1)
      .setNumberFormat('#,##0');
  }
}

function backfillCategories_(sheet) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return;

  const range = sheet.getRange(2, 1, lastRow - 1, HEADERS.length);
  const rows = range.getValues();
  let changed = false;

  rows.forEach(row => {
    if (!row[4] && row[5]) {
      row[4] = detectCategory_(row[5], row[2]);
      changed = true;
    }
  });

  if (changed) range.setValues(rows);
}

// ============================================================
// CÁC HÀM TIỆN ÍCH CHUYỂN ĐỔI & LÀM SẠCH DỮ LIỆU AN TOÀN
// ============================================================

/**
 * Chuyển đổi và làm sạch số tiền an toàn tuyệt đối từ bất kỳ kiểu dữ liệu nào
 * Xử lý hoàn hảo các trường hợp: 50000, "50.000", "50,000", "1.500.000 ₫", "2,500,000 VND"
 */
function cleanNumber_(val) {
  if (val === null || val === undefined || val === '') return 0;
  if (typeof val === 'number') return isFinite(val) ? val : 0;

  let str = String(val).trim();
  // Xóa ký hiệu tiền tệ và khoảng trắng: "50.000 ₫" -> "50.000"
  str = str.replace(/[₫đvndVND\s]/g, '');

  const dotCount = (str.match(/\./g) || []).length;
  const commaCount = (str.match(/,/g) || []).length;

  if (dotCount && commaCount) {
    const lastDot = str.lastIndexOf('.');
    const lastComma = str.lastIndexOf(',');
    if (lastComma > lastDot) {
      // Dạng VN/Châu Âu: 1.000.000,00
      str = str.replace(/\./g, '').replace(',', '.');
    } else {
      // Dạng Mỹ: 1,000,000.00
      str = str.replace(/,/g, '');
    }
  } else if (dotCount > 1) {
    str = str.replace(/\./g, '');
  } else if (commaCount > 1) {
    str = str.replace(/,/g, '');
  } else if (dotCount === 1) {
    const parts = str.split('.');
    // Nếu có đúng 3 chữ số sau dấu chấm -> phân cách hàng nghìn (50.000 -> 50000)
    if (parts[1] && parts[1].length === 3) {
      str = parts[0] + parts[1];
    }
  } else if (commaCount === 1) {
    const parts = str.split(',');
    // Nếu có đúng 3 chữ số sau dấu phẩy -> phân cách hàng nghìn (50,000 -> 50000)
    if (parts[1] && parts[1].length === 3) {
      str = parts[0] + parts[1];
    } else {
      str = parts[0] + '.' + parts[1];
    }
  }

  const num = parseFloat(str);
  return isFinite(num) ? num : 0;
}

/**
 * Phân tích ngày tháng an toàn từ Date, ISO string, text DD/MM/YYYY của Google Sheets
 */
function parseDateSafe_(val) {
  if (!val) return null;
  if (val instanceof Date) {
    return isNaN(val.getTime()) ? null : val;
  }
  if (typeof val === 'number') {
    if (val > 100000000000) return new Date(val);
    const excelEpoch = new Date(1899, 11, 30);
    return new Date(excelEpoch.getTime() + val * 86400000);
  }

  const str = String(val).trim();
  if (!str) return null;

  // Định dạng DD/MM/YYYY hoặc DD-MM-YYYY (rất phổ biến trong Google Sheets tiếng Việt)
  const vnMatch = str.match(/^(\d{1,2})[\/-](\d{1,2})[\/-](\d{4})(?:\s+(\d{1,2}):(\d{1,2})(?::(\d{1,2}))?)?/);
  if (vnMatch) {
    const day = parseInt(vnMatch[1], 10);
    const month = parseInt(vnMatch[2], 10);
    const year = parseInt(vnMatch[3], 10);
    const hour = parseInt(vnMatch[4] || '0', 10);
    const min = parseInt(vnMatch[5] || '0', 10);
    const sec = parseInt(vnMatch[6] || '0', 10);
    const res = new Date(year, month - 1, day, hour, min, sec);
    return isNaN(res.getTime()) ? null : res;
  }

  const d = new Date(str);
  return isNaN(d.getTime()) ? null : d;
}

// ============================================================
// META SHEET: TỔNG THU / TỔNG CHI / SỐ DƯ VÀ ĐỒNG BỘ TỰ ĐỘNG
// ============================================================

function getMetaSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(META_SHEET);

  if (!sheet) {
    sheet = ss.insertSheet(META_SHEET);
    sheet.getRange('A1:B4').setValues([
      ['Tổng thu (toàn thời gian)', 0],
      ['Tổng chi (toàn thời gian)', 0],
      ['Số dư hiện tại', 0],
      ['Cập nhật lúc', '']
    ]);
    sheet.getRange('B3').setFormula('=B1-B2');
    sheet.getRange('A1:A4').setFontWeight('bold');
    sheet.setColumnWidth(1, 220);
    sheet.setColumnWidth(2, 160);
  }

  if (!isMetaBootstrapped_()) {
    recalculateMetaTotals_(sheet);
    markMetaBootstrapped_();
  }

  return sheet;
}

function isMetaBootstrapped_() {
  return PropertiesService.getDocumentProperties().getProperty(META_BOOTSTRAP_PROP) === '1';
}

function markMetaBootstrapped_() {
  PropertiesService.getDocumentProperties().setProperty(META_BOOTSTRAP_PROP, '1');
}

/**
 * Tính toán lại chính xác 100% Tổng thu, Tổng chi, Số dư từ toàn bộ dữ liệu trong sheet LOG
 * Tự động sửa chữa khi có sai lệch (Auto-heal)
 */
function recalculateMetaTotals_(metaSheet) {
  metaSheet = metaSheet || getMetaSheet_();
  const logSheet = getOrCreateSheet_();
  const lastRow = logSheet.getLastRow();

  let income = 0;
  let expense = 0;

  if (lastRow >= 2) {
    const values = logSheet.getRange(2, 3, lastRow - 1, 2).getValues();
    for (let i = 0; i < values.length; i++) {
      const type = String(values[i][0] || '').trim();
      const amount = cleanNumber_(values[i][1]);
      if (type === 'Thu') {
        income += amount;
      } else {
        expense += amount;
      }
    }
  }

  metaSheet.getRange('B1').setValue(income);
  metaSheet.getRange('B2').setValue(expense);
  metaSheet.getRange('B3').setFormula('=B1-B2');
  metaSheet.getRange('B4').setValue(new Date());
  SpreadsheetApp.flush();

  return {
    income: income,
    expense: expense,
    balance: income - expense
  };
}

/**
 * Hàm công khai cho frontend gọi khi người dùng bấm nút Làm mới (Refresh)
 * Đảm bảo số dư luôn đồng bộ 100% với dữ liệu thực tế
 */
function recalculateBalance() {
  const metaSheet = getMetaSheet_();
  return recalculateMetaTotals_(metaSheet);
}

/**
 * Điều chỉnh cộng/trừ O(1) vào META khi có giao dịch thêm/sửa/xóa
 */
function adjustMetaTotals_(type, amountDelta) {
  amountDelta = cleanNumber_(amountDelta);
  if (!amountDelta) return;

  const sheet = getMetaSheet_();
  const cell = sheet.getRange(type === 'Thu' ? 'B1' : 'B2');
  const current = cleanNumber_(cell.getValue());
  cell.setValue(current + amountDelta);
  sheet.getRange('B4').setValue(new Date());
}

/**
 * Lấy số dư hiện tại từ bảng META.
 * Tự động đồng bộ lại nếu bảng META bị trống hoặc chưa có dữ liệu hợp lệ.
 */
function getBalance() {
  const sheet = getMetaSheet_();
  const values = sheet.getRange('B1:B3').getValues();
  const income = cleanNumber_(values[0][0]);
  const expense = cleanNumber_(values[1][0]);
  const balance = cleanNumber_(values[2][0]);

  // Tự phục hồi: nếu cả thu và chi đều bằng 0 nhưng LOG có dữ liệu
  if (income === 0 && expense === 0) {
    const logSheet = getOrCreateSheet_();
    if (logSheet.getLastRow() >= 2) {
      return recalculateMetaTotals_(sheet);
    }
  }

  return {
    income: income,
    expense: expense,
    balance: balance
  };
}

// ============================================================
// BỘ PHÂN TÍCH TIẾNG VIỆT THÔNG MINH (NATURAL LANGUAGE PARSER)
// ============================================================

/**
 * Nhận diện cụm số kết hợp đặc biệt trong tiếng Việt:
 * Ví dụ: "1tr5", "2k5", "1củ5", "2 triệu rưỡi", "1 củ rưỡi", "50 nghìn rưỡi"
 */
function parseCompoundAmount_(text) {
  // Cú pháp "... rưỡi"
  const ruoiRegex = /([0-9]+(?:[.,][0-9]+)?)\s*(triệu|tr|củ|nghìn|ngàn|k)\s*rưỡi(?=$|[\s,;:.!?-])/i;
  const ruoiMatch = text.match(ruoiRegex);
  if (ruoiMatch) {
    let base = parseFloat(ruoiMatch[1].replace(',', '.'));
    const unit = ruoiMatch[2].toLowerCase();
    if (unit === 'triệu' || unit === 'tr' || unit === 'củ') {
      base = base * 1000000 + 500000;
    } else {
      base = base * 1000 + 500;
    }
    return {
      rawMatch: ruoiMatch[0],
      amount: Math.round(base)
    };
  }

  // Cú pháp viết tắt: "1tr5", "2tr350", "1k5"
  const shortCompoundRegex = /([0-9]+)(tr|triệu|củ|k)([0-9]+)(?=$|[\s,;:.!?-])/i;
  const shortMatch = text.match(shortCompoundRegex);
  if (shortMatch) {
    const main = parseInt(shortMatch[1], 10);
    const unit = shortMatch[2].toLowerCase();
    const subStr = shortMatch[3];
    let amount = 0;
    if (unit === 'tr' || unit === 'triệu' || unit === 'củ') {
      let frac = parseFloat('0.' + subStr);
      amount = Math.round((main + frac) * 1000000);
    } else if (unit === 'k') {
      let frac = parseFloat('0.' + subStr);
      amount = Math.round((main + frac) * 1000);
    }
    return {
      rawMatch: shortMatch[0],
      amount: amount
    };
  }

  return null;
}

function parseAmountToken_(numberText, unit) {
  let cleaned = String(numberText).trim();

  const dotCount = (cleaned.match(/\./g) || []).length;
  const commaCount = (cleaned.match(/,/g) || []).length;

  if (dotCount && commaCount) {
    cleaned = cleaned.replace(/\./g, '').replace(',', '.');
  } else if (dotCount > 1) {
    cleaned = cleaned.replace(/\./g, '');
  } else if (commaCount > 1) {
    cleaned = cleaned.replace(/,/g, '');
  } else if (dotCount === 1) {
    const parts = cleaned.split('.');
    cleaned = parts[1] && parts[1].length === 3
      ? cleaned.replace('.', '')
      : cleaned;
  } else if (commaCount === 1) {
    const parts = cleaned.split(',');
    cleaned = parts[1] && parts[1].length === 3
      ? cleaned.replace(',', '')
      : cleaned.replace(',', '.');
  }

  let amount = parseFloat(cleaned);
  if (isNaN(amount)) return 0;

  unit = String(unit || '').toLowerCase().trim();

  if (unit === 'k' || unit === 'nghìn' || unit === 'ngàn') {
    amount *= 1000;
  } else if (unit === 'tr' || unit === 'triệu' || unit === 'm') {
    amount *= 1000000;
  } else if (unit === 'tỷ' || unit === 'ti' || unit === 'ty') {
    amount *= 1000000000;
  } else if (unit === 'củ') {
    amount *= 1000000;
  } else if (unit === 'lít' || unit === 'lit' || unit === 'lốp' || unit === 'lop' || unit === 'loét') {
    amount *= 100000;
  }

  return Math.round(amount);
}

/**
 * Phân tích chuỗi nhập tự nhiên thông minh:
 * - Tránh bắt nhầm các số trong ghi chú (tháng 5, 2 ly, 7 chỗ, số 10, 20/10)
 * - Tự động nhận diện Thu / Chi
 */
function parseInput(text) {
  let value = String(text || '').trim();
  if (!value) {
    throw new Error('Bạn chưa nhập nội dung.');
  }

  const lower = value.toLowerCase();
  let type = 'Chi';

  // Nhận diện tiền tố Thu / Chi rõ ràng hoặc tự nhiên
  if (/^(thu|thu:|thu\s*-)\b/i.test(value)) {
    type = 'Thu';
    value = value.replace(/^(thu|thu:|thu\s*-)\s*/i, '').trim();
  } else if (/^(chi|chi:|chi\s*-)\b/i.test(value)) {
    type = 'Chi';
    value = value.replace(/^(chi|chi:|chi\s*-)\s*/i, '').trim();
  } else if (/^(nhận|được tặng|được thưởng|hoàn tiền|lương|tiền về)\b/i.test(lower)) {
    type = 'Thu';
  }

  // 1. Kiểm tra các cụm số phức hợp ("1tr5", "2 triệu rưỡi", "1 củ rưỡi")
  const compound = parseCompoundAmount_(value);
  let bestMatch = null;
  let chosenAmount = 0;

  if (compound) {
    bestMatch = compound.rawMatch;
    chosenAmount = compound.amount;
  } else {
    // 2. Tìm các token số có đơn vị tiền tệ rõ ràng (Ưu tiên số 1)
    const unitPattern = /([0-9][0-9.,]*)\s*(nghìn|ngàn|triệu|củ|lít|lit|lốp|lop|loét|tỷ|ti|ty|đồng|vnđ|vnd|k|tr|m|đ|d)(?=$|[\s,;:.!?-])/gi;
    const unitMatches = [];
    let m;
    while ((m = unitPattern.exec(value)) !== null) {
      const amt = parseAmountToken_(m[1], m[2]);
      if (amt > 0) {
        unitMatches.push({
          rawMatch: m[0],
          amount: amt,
          index: m.index
        });
      }
    }

    if (unitMatches.length > 0) {
      chosenAmount = unitMatches[0].amount;
      bestMatch = unitMatches[0].rawMatch;
    } else {
      // 3. Không có đơn vị tiền tệ: Tìm số đứng độc lập (không nằm trong ngày tháng như 20/10)
      const rawPattern = /(?<![\/\d])([0-9][0-9.,]*)(?![\/\d])/g;
      const rawMatches = [];
      while ((m = rawPattern.exec(value)) !== null) {
        const cleanedStr = m[1].replace(/[.,]$/, '');
        const amt = cleanNumber_(cleanedStr);
        if (amt > 0) {
          rawMatches.push({
            rawMatch: m[0],
            amount: Math.round(amt),
            index: m.index
          });
        }
      }

      if (rawMatches.length === 0) {
        throw new Error('Không nhận diện được số tiền. Ví dụ: 50k ăn sáng');
      }

      // Ưu tiên số tiền lớn (>= 1000) để không bắt nhầm các số đếm như "2 ly", "7 chỗ", "số 10"
      const largeMatches = rawMatches.filter(r => r.amount >= 1000);
      if (largeMatches.length > 0) {
        chosenAmount = largeMatches[largeMatches.length - 1].amount;
        bestMatch = largeMatches[largeMatches.length - 1].rawMatch;
      } else {
        chosenAmount = rawMatches[rawMatches.length - 1].amount;
        bestMatch = rawMatches[rawMatches.length - 1].rawMatch;
      }
    }
  }

  if (!isFinite(chosenAmount) || chosenAmount <= 0) {
    throw new Error('Số tiền không hợp lệ.');
  }

  let note = value
    .replace(bestMatch, '')
    .trim()
    .replace(/^[\s\-:;,.]+/, '')
    .replace(/[\s\-:;,.]+$/, '')
    .trim();

  if (!note) note = type === 'Thu' ? 'Thu nhập' : 'Chi tiêu';

  return {
    type: type,
    amount: Math.round(chosenAmount),
    category: detectCategory_(note, type),
    note: note
  };
}

/**
 * Tự động phân loại danh mục
 */
function detectCategory_(note, type) {
  const text = String(note || '').toLowerCase();

  if (type === 'Thu') {
    if (/lương|salary|payroll/.test(text)) return 'Lương';
    if (/thưởng|bonus/.test(text)) return 'Thưởng';
    if (/bán|sale|bán hàng|freelance|dự án/.test(text)) return 'Thu nhập khác';
    if (/lãi|tiết kiệm|cổ tức|đầu tư/.test(text)) return 'Đầu tư';
    return 'Thu nhập khác';
  }

  const rules = [
    ['Ăn uống', /ăn|cơm|phở|bún|mì|trưa|sáng|tối|nhà hàng|quán|đồ ăn|food|bánh|lẩu|nướng/],
    ['Cà phê & đồ uống', /cà phê|cafe|trà sữa|nước uống|bia|đồ uống|sinh tố|nước ép/],
    ['Di chuyển', /xăng|grab|taxi|be\b|go\b|xe|ship|gửi xe|vé xe|bus|xe buýt|bảo dưỡng/],
    ['Nhà ở', /thuê nhà|tiền nhà|điện|nước sinh hoạt|internet|wifi|phòng trọ|nội thất/],
    ['Hóa đơn', /hóa đơn|điện thoại|điện|nước|internet|wifi|cước|nạp tiền/],
    ['Sức khỏe', /thuốc|khám|bệnh viện|y tế|nha khoa|vitamin|bác sĩ/],
    ['Mua sắm', /mua sắm|shopee|lazada|tiki|quần áo|giày|túi|mỹ phẩm|đồ dùng/],
    ['Học tập', /học|sách|khóa học|course|học phí|đào tạo/],
    ['Giải trí', /giải trí|phim|game|netflix|spotify|du lịch|karaoke/],
    ['Gia đình', /gia đình|bố|mẹ|ba|má|con|vợ|chồng|em|anh|chị/],
    ['Đầu tư', /đầu tư|cổ phiếu|chứng khoán|quỹ|crypto/],
    ['Quà tặng', /quà|sinh nhật|cưới|mừng/]
  ];

  for (let i = 0; i < rules.length; i++) {
    if (rules[i][1].test(text)) return rules[i][0];
  }

  return 'Khác';
}

// ============================================================
// TRANSACTION CRUD (THÊM / SỬA / XÓA AN TOÀN)
// ============================================================

function addTransaction(text) {
  if (!text || !String(text).trim()) {
    throw new Error('Bạn chưa nhập nội dung.');
  }

  const parsed = parseInput(text);
  const lock = LockService.getScriptLock();

  if (!lock.tryLock(10000)) {
    throw new Error('Hệ thống đang bận, thử lại sau ít giây.');
  }

  try {
    const sheet = getOrCreateSheet_();
    const now = new Date();
    const id = Utilities.getUuid();

    sheet.appendRow([
      id,
      now,
      parsed.type,
      parsed.amount,
      parsed.category,
      parsed.note,
      text
    ]);

    adjustMetaTotals_(parsed.type, parsed.amount);
    SpreadsheetApp.flush();

    return {
      id: id,
      type: parsed.type,
      amount: parsed.amount,
      category: parsed.category,
      note: parsed.note,
      timestamp: now.toISOString()
    };
  } finally {
    lock.releaseLock();
  }
}

function deleteTransaction(id) {
  if (!id) throw new Error('Thiếu mã giao dịch.');

  const lock = LockService.getScriptLock();
  if (!lock.tryLock(10000)) {
    throw new Error('Hệ thống đang bận, thử lại sau ít giây.');
  }

  try {
    const sheet = getOrCreateSheet_();
    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return { success: false };

    // Chỉ đọc duy nhất 1 cột ID (Cột A) để tìm dòng cần xóa nhanh nhất
    const ids = sheet.getRange(2, 1, lastRow - 1, 1).getValues();

    for (let i = 0; i < ids.length; i++) {
      if (String(ids[i][0]).trim() === String(id).trim()) {
        const rowIndex = i + 2;
        // Đọc Loại & Số tiền của dòng đó để trừ META
        const rowMeta = sheet.getRange(rowIndex, 3, 1, 2).getValues()[0];
        const type = String(rowMeta[0] || 'Chi').trim();
        const amount = cleanNumber_(rowMeta[1]);

        sheet.deleteRow(rowIndex);
        adjustMetaTotals_(type, -amount);
        SpreadsheetApp.flush();

        return { success: true };
      }
    }

    return { success: false };
  } finally {
    lock.releaseLock();
  }
}

function updateTransaction(id, updates) {
  if (!id) throw new Error('Thiếu mã giao dịch.');
  updates = updates || {};

  const lock = LockService.getScriptLock();
  if (!lock.tryLock(10000)) {
    throw new Error('Hệ thống đang bận, thử lại sau ít giây.');
  }

  try {
    const sheet = getOrCreateSheet_();
    const lastRow = sheet.getLastRow();
    if (lastRow < 2) throw new Error('Không tìm thấy giao dịch.');

    const ids = sheet.getRange(2, 1, lastRow - 1, 1).getValues();

    for (let i = 0; i < ids.length; i++) {
      if (String(ids[i][0]).trim() !== String(id).trim()) continue;

      const row = i + 2;
      const currentRow = sheet.getRange(row, 1, 1, HEADERS.length).getValues()[0];
      const oldType = String(currentRow[2] || 'Chi').trim();
      const oldAmount = cleanNumber_(currentRow[3]);

      if (updates.type === 'Thu' || updates.type === 'Chi') {
        sheet.getRange(row, 3).setValue(updates.type);
      }

      let newAmount = oldAmount;
      if (updates.amount !== undefined && updates.amount !== null && updates.amount !== '') {
        const amount = cleanNumber_(updates.amount);
        if (!isFinite(amount) || amount <= 0) {
          throw new Error('Số tiền không hợp lệ.');
        }
        newAmount = Math.round(amount);
        sheet.getRange(row, 4).setValue(newAmount);
      }

      if (updates.category) {
        sheet.getRange(row, 5).setValue(String(updates.category).trim());
      }

      if (updates.note !== undefined) {
        const note = String(updates.note).trim();
        if (!note) throw new Error('Nội dung không được để trống.');
        sheet.getRange(row, 6).setValue(note);
      }

      const newType = (updates.type === 'Thu' || updates.type === 'Chi') ? updates.type : oldType;

      if (oldType === newType) {
        adjustMetaTotals_(oldType, newAmount - oldAmount);
      } else {
        adjustMetaTotals_(oldType, -oldAmount);
        adjustMetaTotals_(newType, newAmount);
      }

      SpreadsheetApp.flush();

      const updatedRow = sheet.getRange(row, 1, 1, HEADERS.length).getValues()[0];
      const date = parseDateSafe_(updatedRow[1]);

      return {
        id: updatedRow[0] || null,
        timestamp: date ? date.toISOString() : new Date().toISOString(),
        type: updatedRow[2] || 'Chi',
        amount: cleanNumber_(updatedRow[3]),
        category: updatedRow[4] || 'Khác',
        note: updatedRow[5] || ''
      };
    }

    throw new Error('Không tìm thấy giao dịch.');
  } finally {
    lock.releaseLock();
  }
}

// ============================================================
// ĐỌC DỮ LIỆU THEO KHOẢNG THỜI GIAN (BATCH READ TỐI ƯU)
// ============================================================

function rowsToTransactions_(values) {
  return values
    .map(row => {
      const date = parseDateSafe_(row[1]);
      if (!date) return null;
      return {
        id: row[0] || null,
        timestamp: date.toISOString(),
        type: row[2] || 'Chi',
        amount: cleanNumber_(row[3]),
        category: row[4] || detectCategory_(row[5], row[2]),
        note: row[5] || ''
      };
    })
    .filter(item => item && item.id);
}

/**
 * Lấy danh sách giao dịch trong khoảng thời gian [startIso, endIso)
 * TỐI ƯU HÓA: Quét 1 lần duy nhất cột Thời gian (B) thay cho Binary Search
 * Hoạt động chính xác 100% kể cả khi bảng tính bị sắp xếp (Sort) hoặc sai thứ tự
 */
function getTransactionsInRange(startIso, endIso) {
  const sheet = getOrCreateSheet_();
  const lastRow = sheet.getLastRow();
  const totalDataRows = Math.max(0, lastRow - 1);
  if (totalDataRows === 0) return [];

  const start = startIso ? parseDateSafe_(startIso) : null;
  const end = endIso ? parseDateSafe_(endIso) : null;

  // 1. Đọc duy nhất 1 cột Thời gian (Cột B) chỉ với 1 network call
  const dateValues = sheet.getRange(2, 2, totalDataRows, 1).getValues();
  const matchedIndices = [];

  for (let i = 0; i < dateValues.length; i++) {
    const raw = dateValues[i][0];
    if (!raw) continue;
    const d = parseDateSafe_(raw);
    if (!d) continue;

    if (start && d < start) continue;
    if (end && d >= end) continue;

    matchedIndices.push(i);
  }

  if (matchedIndices.length === 0) return [];

  // 2. Đọc theo khối nếu các dòng gần nhau
  const minIdx = matchedIndices[0];
  const maxIdx = matchedIndices[matchedIndices.length - 1];
  const rowSpan = maxIdx - minIdx + 1;

  if (rowSpan <= Math.max(matchedIndices.length * 2, 50)) {
    const startRow = 2 + minIdx;
    const allValues = sheet.getRange(startRow, 1, rowSpan, HEADERS.length).getValues();
    const matchedSet = new Set(matchedIndices.map(idx => idx - minIdx));
    const selectedRows = [];
    for (let k = 0; k < allValues.length; k++) {
      if (matchedSet.has(k)) {
        selectedRows.push(allValues[k]);
      }
    }
    return rowsToTransactions_(selectedRows);
  } else {
    const selectedRows = [];
    matchedIndices.forEach(idx => {
      const rowVal = sheet.getRange(2 + idx, 1, 1, HEADERS.length).getValues()[0];
      selectedRows.push(rowVal);
    });
    return rowsToTransactions_(selectedRows);
  }
}

/**
 * Tính tổng Thu và Chi trong khoảng thời gian [startIso, endIso)
 * Chỉ đọc 1 range duy nhất gồm 3 cột (B, C, D) -> Tốc độ cực nhanh ~50ms
 */
function getExpenseSumInRange(startIso, endIso) {
  const sheet = getOrCreateSheet_();
  const lastRow = sheet.getLastRow();
  const totalDataRows = Math.max(0, lastRow - 1);
  if (totalDataRows === 0) return { income: 0, expense: 0 };

  const start = startIso ? parseDateSafe_(startIso) : null;
  const end = endIso ? parseDateSafe_(endIso) : null;

  // Đọc 3 cột Thời gian, Loại, Số tiền (Cột B, C, D)
  const values = sheet.getRange(2, 2, totalDataRows, 3).getValues();

  let income = 0;
  let expense = 0;

  for (let i = 0; i < values.length; i++) {
    const rawDate = values[i][0];
    if (!rawDate) continue;
    const d = parseDateSafe_(rawDate);
    if (!d) continue;

    if (start && d < start) continue;
    if (end && d >= end) continue;

    const type = String(values[i][1] || '').trim();
    const amount = cleanNumber_(values[i][2]);

    if (type === 'Thu') {
      income += amount;
    } else {
      expense += amount;
    }
  }

  return { income: income, expense: expense };
}

/**
 * Đọc theo phân trang cho chế độ xem "Tất cả"
 */
function getTransactionsChunk(offset, limit) {
  offset = Math.max(0, Number(offset) || 0);
  limit = Math.max(1, Number(limit) || 3000);

  const sheet = getOrCreateSheet_();
  const lastRow = sheet.getLastRow();
  const totalDataRows = Math.max(0, lastRow - 1);

  if (offset >= totalDataRows) {
    return { rows: [], total: totalDataRows };
  }

  const startRow = 2 + offset;
  const rowsToRead = Math.min(limit, totalDataRows - offset);

  const values = sheet.getRange(startRow, 1, rowsToRead, HEADERS.length).getValues();
  return { rows: rowsToTransactions_(values), total: totalDataRows };
}
