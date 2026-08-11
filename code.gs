const LOG_SHEET = 'LOG';
const HEADERS = ['ID', 'Thời gian', 'Loại', 'Số tiền', 'Danh mục', 'Nội dung', 'Nguyên văn'];
const CACHE_KEY = 'transactions_v2';
const CACHE_SECONDS = 30;

function doGet() {
  return HtmlService
    .createHtmlOutputFromFile('Index')
    .setTitle('Nhật ký chi tiêu')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function setup() {
  const sheet = getOrCreateSheet_();
  return {
    success: true,
    sheet: sheet.getName(),
    timezone: SpreadsheetApp.getActive().getSpreadsheetTimeZone()
  };
}

/**
 * Current schema:
 * ID | Thời gian | Loại | Số tiền | Danh mục | Nội dung | Nguyên văn
 *
 * Legacy schema had an extra "Ngày" column. It is removed because
 * "Thời gian" is the single source of truth for date + time.
 */
function getOrCreateSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(LOG_SHEET);

  if (!sheet) {
    sheet = ss.insertSheet(LOG_SHEET);
    sheet.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]);
    formatHeader_(sheet);
    return sheet;
  }

  migrateSchema_(sheet);
  return sheet;
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

  // Remove the legacy duplicate date column only when the old schema is detected.
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

    invalidateCache_();

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

    const ids = sheet.getRange(2, 1, lastRow - 1, 1).getValues();

    for (let i = 0; i < ids.length; i++) {
      if (String(ids[i][0]) === String(id)) {
        sheet.deleteRow(i + 2);
        invalidateCache_();
        return { success: true };
      }
    }

    return { success: false };
  } finally {
    lock.releaseLock();
  }
}

function parseInput(text) {
  let value = String(text).trim();
  const lower = value.toLowerCase();

  let type = 'Chi';

  if (lower === 'thu' || lower.startsWith('thu ')) {
    type = 'Thu';
    value = value.substring(3).trim();
  } else if (lower === 'chi' || lower.startsWith('chi ')) {
    type = 'Chi';
    value = value.substring(3).trim();
  }

  const regex = /([0-9][0-9.,]*)\s*(nghìn|ngàn|triệu|đồng|vnđ|vnd|k|tr|m|đ|d)?/i;
  const match = value.match(regex);

  if (!match || !match[1]) {
    throw new Error('Không nhận diện được số tiền. Ví dụ: 50k ăn sáng');
  }

  const amount = parseAmountToken_(match[1], (match[2] || '').toLowerCase());

  if (!isFinite(amount) || amount <= 0) {
    throw new Error('Số tiền không hợp lệ.');
  }

  let note = value
    .replace(match[0], '')
    .trim()
    .replace(/^[\s\-:;,.]+/, '');

  if (!note) note = type === 'Thu' ? 'Thu nhập' : 'Chi tiêu';

  return {
    type: type,
    amount: Math.round(amount),
    category: detectCategory_(note, type),
    note: note
  };
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

  if (unit === 'k' || unit === 'nghìn' || unit === 'ngàn') {
    amount *= 1000;
  } else if (unit === 'tr' || unit === 'triệu' || unit === 'm') {
    amount *= 1000000;
  }

  return amount;
}

/**
 * Automatic categorization suitable for personal-finance tracking.
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
    ['Ăn uống', /ăn|cơm|phở|bún|mì|trưa|sáng|tối|nhà hàng|quán|đồ ăn|food/],
    ['Cà phê & đồ uống', /cà phê|cafe|trà sữa|nước uống|bia|đồ uống/],
    ['Di chuyển', /xăng|grab|taxi|be\b|go\b|xe|ship|gửi xe|vé xe|bus|xe buýt/],
    ['Nhà ở', /thuê nhà|tiền nhà|điện|nước sinh hoạt|internet|wifi|phòng trọ|nội thất/],
    ['Hóa đơn', /hóa đơn|điện thoại|điện|nước|internet|wifi|cước/],
    ['Sức khỏe', /thuốc|khám|bệnh viện|y tế|nha khoa|vitamin/],
    ['Mua sắm', /mua sắm|shopee|lazada|tiki|quần áo|giày|túi|mỹ phẩm/],
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

function getTransactions() {
  const cache = CacheService.getScriptCache();
  const cached = cache.get(CACHE_KEY);

  if (cached) return JSON.parse(cached);

  const sheet = getOrCreateSheet_();
  const lastRow = sheet.getLastRow();

  if (lastRow < 2) return [];

  const values = sheet
    .getRange(2, 1, lastRow - 1, HEADERS.length)
    .getValues();

  const result = values
    .map(row => {
      const date = row[1] ? new Date(row[1]) : null;

      return {
        id: row[0] || null,
        timestamp: date && !isNaN(date.getTime()) ? date.toISOString() : null,
        type: row[2] || 'Chi',
        amount: Number(row[3]) || 0,
        category: row[4] || detectCategory_(row[5], row[2]),
        note: row[5] || '',
        original: row[6] || ''
      };
    })
    .filter(item => item.id && item.timestamp);

  try {
    cache.put(CACHE_KEY, JSON.stringify(result), CACHE_SECONDS);
  } catch (e) {
    // Cache is optional; large datasets can exceed Apps Script limits.
  }

  return result;
}

function getStatistics() {
  return buildStatistics_(getTransactions());
}

function buildStatistics_(transactions) {
  const expenseByCategory = {};
  const incomeByCategory = {};
  const daily = {};

  let income = 0;
  let expense = 0;

  transactions.forEach(item => {
    const amount = Number(item.amount) || 0;
    const category = item.category || 'Khác';
    const date = new Date(item.timestamp);

    if (item.type === 'Thu') {
      income += amount;
      incomeByCategory[category] = (incomeByCategory[category] || 0) + amount;
    } else {
      expense += amount;
      expenseByCategory[category] = (expenseByCategory[category] || 0) + amount;
    }

    if (!isNaN(date.getTime())) {
      const key = Utilities.formatDate(
        date,
        Session.getScriptTimeZone(),
        'yyyy-MM-dd'
      );

      if (!daily[key]) daily[key] = { income: 0, expense: 0 };
      if (item.type === 'Thu') daily[key].income += amount;
      else daily[key].expense += amount;
    }
  });

  return {
    income: income,
    expense: expense,
    balance: income - expense,
    transactionCount: transactions.length,
    expenseByCategory: expenseByCategory,
    incomeByCategory: incomeByCategory,
    daily: daily
  };
}

function invalidateCache_() {
  CacheService.getScriptCache().remove(CACHE_KEY);
}
