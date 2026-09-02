const LOG_SHEET = 'LOG';
const META_SHEET = 'META';
const HEADERS = ['ID', 'Thời gian', 'Loại', 'Số tiền', 'Danh mục', 'Nội dung', 'Nguyên văn'];

// Bump this whenever migrateSchema_/backfillCategories_ logic changes so it
// re-runs exactly once for existing sheets, instead of on every request.
const SCHEMA_VERSION = 'v2-danh-muc';
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
  getMetaSheet_();
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
 *
 * IMPORTANT: this function is called on every addTransaction/
 * deleteTransaction/updateTransaction/getTransactionsInRange request.
 * Schema migration scans (and can rewrite) the whole sheet, so it must
 * only run once — never on the hot path — or every single click in the
 * app pays for a full-sheet read/write.
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

// ============================================================
// META sheet: running income/expense/balance totals.
//
// Instead of recomputing the balance by scanning every transaction on
// every request, the totals live in three cells and are nudged by a
// small delta whenever a transaction is added/edited/deleted — O(1)
// regardless of how many rows LOG has. Balance is a formula (income -
// expense) so it can never drift out of sync with the two totals that
// actually get written.
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
    bootstrapMetaTotals_(sheet);
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

// One-time full scan (guarded so it only ever runs once) to seed META
// from whatever rows already exist in LOG at the time this ships.
function bootstrapMetaTotals_(metaSheet) {
  const logSheet = getOrCreateSheet_();
  const lastRow = logSheet.getLastRow();

  let income = 0;
  let expense = 0;

  if (lastRow >= 2) {
    // Only the Loại + Số tiền columns are needed to sum totals.
    const values = logSheet.getRange(2, 3, lastRow - 1, 2).getValues();
    values.forEach(pair => {
      const type = pair[0];
      const amount = Number(pair[1]) || 0;
      if (type === 'Thu') income += amount; else expense += amount;
    });
  }

  metaSheet.getRange('B1').setValue(income);
  metaSheet.getRange('B2').setValue(expense);
  metaSheet.getRange('B4').setValue(new Date());
}

// O(1): reads one cell, writes it back with the delta applied.
function adjustMetaTotals_(type, amountDelta) {
  if (!amountDelta) return;
  const sheet = getMetaSheet_();
  const cell = sheet.getRange(type === 'Thu' ? 'B1' : 'B2');
  const current = Number(cell.getValue()) || 0;
  cell.setValue(current + amountDelta);
  sheet.getRange('B4').setValue(new Date());
}

/**
 * Returns { income, expense, balance } straight from the META cells —
 * three cell reads, no matter how many rows LOG has. This is what the
 * "CÒN LẠI" card uses, so the all-time balance is always instantly
 * available even when the browser has only loaded a single day's worth
 * of transactions.
 */
function getBalance() {
  const sheet = getMetaSheet_();
  const values = sheet.getRange('B1:B3').getValues();
  return {
    income: Number(values[0][0]) || 0,
    expense: Number(values[1][0]) || 0,
    balance: Number(values[2][0]) || 0
  };
}

// ============================================================
// Transaction CRUD
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

    // ID, Thời gian, Loại, Số tiền — need type+amount to undo the META totals.
    const rows = sheet.getRange(2, 1, lastRow - 1, 4).getValues();

    for (let i = 0; i < rows.length; i++) {
      if (String(rows[i][0]) === String(id)) {
        const type = rows[i][2] || 'Chi';
        const amount = Number(rows[i][3]) || 0;

        sheet.deleteRow(i + 2);
        adjustMetaTotals_(type, -amount);

        return { success: true };
      }
    }

    return { success: false };
  } finally {
    lock.releaseLock();
  }
}

/**
 * Updates an existing transaction's type/amount/category/note in place.
 * "Nguyên văn" (the original raw text) is preserved untouched so the
 * edit history of what the user actually typed is never lost.
 */
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

    const rows = sheet.getRange(2, 1, lastRow - 1, 4).getValues();

    for (let i = 0; i < rows.length; i++) {
      if (String(rows[i][0]) !== String(id)) continue;

      const row = i + 2;
      const oldType = rows[i][2] || 'Chi';
      const oldAmount = Number(rows[i][3]) || 0;

      if (updates.type === 'Thu' || updates.type === 'Chi') {
        sheet.getRange(row, 3).setValue(updates.type);
      }

      let newAmount = oldAmount;
      if (updates.amount !== undefined && updates.amount !== null && updates.amount !== '') {
        const amount = Number(updates.amount);
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

      const values = sheet.getRange(row, 1, 1, HEADERS.length).getValues()[0];
      const date = values[1] ? new Date(values[1]) : null;

      return {
        id: values[0] || null,
        timestamp: date && !isNaN(date.getTime()) ? date.toISOString() : null,
        type: values[2] || 'Chi',
        amount: Number(values[3]) || 0,
        category: values[4] || 'Khác',
        note: values[5] || ''
      };
    }

    throw new Error('Không tìm thấy giao dịch.');
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

// ============================================================
// Range-scoped reads
//
// LOG rows are always appended in chronological order (appendRow), so
// column B (Thời gian) is non-decreasing top to bottom. That means the
// start/end row of any date range can be located with a binary search —
// O(log n) sheet reads — instead of scanning every row to find matches.
// Only the located slice is then read in full. This is what lets the
// frontend load "just today", "just this month", etc. at roughly
// constant cost regardless of how many years of history the sheet holds.
// ============================================================

/**
 * Returns the 0-based data-row index (0 = first data row) of the first
 * row whose timestamp is >= targetDate.
 */
function locateBoundaryRow_(sheet, totalDataRows, targetDate) {
  if (totalDataRows <= 0) return 0;

  let lo = 0;
  let hi = totalDataRows;

  while (lo < hi) {
    const mid = (lo + hi) >> 1;
    const raw = sheet.getRange(2 + mid, 2).getValue();
    const d = raw ? new Date(raw) : null;
    const valid = d && !isNaN(d.getTime());

    if (!valid || d < targetDate) {
      lo = mid + 1;
    } else {
      hi = mid;
    }
  }

  return lo;
}

function rowsToTransactions_(values) {
  return values
    .map(row => {
      const date = row[1] ? new Date(row[1]) : null;
      if (!date || isNaN(date.getTime())) return null;
      return {
        id: row[0] || null,
        timestamp: date.toISOString(),
        type: row[2] || 'Chi',
        amount: Number(row[3]) || 0,
        category: row[4] || detectCategory_(row[5], row[2]),
        note: row[5] || ''
      };
    })
    .filter(item => item && item.id);
}

/**
 * Fetches only the transactions whose timestamp falls in [startIso, endIso).
 * Either bound may be null/omitted for an open end. This is the main
 * entry point the frontend uses for every period filter (today, this
 * week, a picked month, a custom date range) — it never reads more of
 * the sheet than the requested window actually spans.
 */
function getTransactionsInRange(startIso, endIso) {
  const sheet = getOrCreateSheet_();
  const lastRow = sheet.getLastRow();
  const totalDataRows = Math.max(0, lastRow - 1);
  if (totalDataRows === 0) return [];

  const start = startIso ? new Date(startIso) : null;
  const end = endIso ? new Date(endIso) : null;

  const fromIdx = start ? locateBoundaryRow_(sheet, totalDataRows, start) : 0;
  const toIdx = end ? locateBoundaryRow_(sheet, totalDataRows, end) : totalDataRows;

  if (toIdx <= fromIdx) return [];

  const startRow = 2 + fromIdx;
  const rowCount = toIdx - fromIdx;

  const values = sheet.getRange(startRow, 1, rowCount, HEADERS.length).getValues();
  return rowsToTransactions_(values);
}

/**
 * Same range-location strategy as getTransactionsInRange, but only sums
 * income/expense instead of returning full rows — used for the "so với
 * kỳ trước" comparison, where a scalar is all that's needed.
 */
function getExpenseSumInRange(startIso, endIso) {
  const sheet = getOrCreateSheet_();
  const lastRow = sheet.getLastRow();
  const totalDataRows = Math.max(0, lastRow - 1);
  if (totalDataRows === 0) return { income: 0, expense: 0 };

  const start = startIso ? new Date(startIso) : null;
  const end = endIso ? new Date(endIso) : null;

  const fromIdx = start ? locateBoundaryRow_(sheet, totalDataRows, start) : 0;
  const toIdx = end ? locateBoundaryRow_(sheet, totalDataRows, end) : totalDataRows;

  if (toIdx <= fromIdx) return { income: 0, expense: 0 };

  const startRow = 2 + fromIdx;
  const rowCount = toIdx - fromIdx;

  // Only Loại + Số tiền are needed to sum.
  const values = sheet.getRange(startRow, 3, rowCount, 2).getValues();

  let income = 0;
  let expense = 0;
  values.forEach(pair => {
    const type = pair[0];
    const amount = Number(pair[1]) || 0;
    if (type === 'Thu') income += amount; else expense += amount;
  });

  return { income: income, expense: expense };
}

/**
 * Reads one fixed-size slice of the sheet by row offset, for the "Tất cả"
 * (all time) view. Even with range-scoped reads for normal filters, "all"
 * has no bound to binary-search against — it always means every row — so
 * it's still pulled in chunks rather than one giant request, keeping any
 * single call small regardless of how large the log has grown.
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
