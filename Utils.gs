/**
 * ============================================================
 * UTILS
 * ============================================================
 */

function getSpreadsheet_() {
  return SpreadsheetApp.getActiveSpreadsheet();
}


function getSheet_(sheetName) {
  const sheet = getSpreadsheet_().getSheetByName(sheetName);

  if (!sheet) {
    throw new Error(`شیت "${sheetName}" پیدا نشد.`);
  }

  return sheet;
}


function normalizeText_(value) {

  if (value === null || value === undefined) {
    return '';
  }

  return String(value)
    .trim()
    .replace(/\s+/g, ' ')
    .toUpperCase();
}

function toNumber_(value, fieldName) {
  const number = Number(value);

  if (!Number.isFinite(number)) {
    throw new Error(`${fieldName} باید عدد باشد.`);
  }

  return number;
}


function roundNumber_(value, decimals = CONFIG.PRECISION) {
  const factor = Math.pow(10, decimals);
  return Math.round((value + Number.EPSILON) * factor) / factor;
}


function approximatelyEqual_(a, b, tolerance = 0.000001) {
  return Math.abs(Number(a) - Number(b)) <= tolerance;
}


function now_() {
  return new Date();
}


function currentUser_() {
  try {
    return Session.getActiveUser().getEmail() || 'کاربر نامشخص';
  } catch (e) {
    return 'کاربر نامشخص';
  }
}


function generateId_(prefix) {
  return `${prefix}-${Utilities.getUuid()}`;
}


function generateProductionNumber_() {

  const lock =
    LockService.getScriptLock();

  if (!lock.tryLock(30000)) {
    throw new Error(
      'سیستم در حال تولید شماره تولید است. لطفاً دوباره تلاش کنید.'
    );
  }

  try {

    const properties =
      PropertiesService.getScriptProperties();

    const key =
      'LAST_PRODUCTION_SEQUENCE';

    const currentYear =
      Utilities.formatDate(
        new Date(),
        Session.getScriptTimeZone(),
        'yyyy'
      );

    const stored =
      properties.getProperty(key);

    let sequence = 0;

    if (stored) {

      const parts =
        stored.split('-');

      const storedYear =
        parts[0];

      const storedSequence =
        Number(parts[1]);

      if (
        storedYear === currentYear &&
        Number.isFinite(storedSequence)
      ) {
        sequence = storedSequence;
      }
    }

    sequence++;

    properties.setProperty(
      key,
      `${currentYear}-${sequence}`
    );

    return (
      `P-${currentYear}-` +
      String(sequence).padStart(6, '0')
    );

  } finally {

    lock.releaseLock();
  }
}

function generateTransactionNumber_() {
  return `TRX-${Utilities.formatDate(
    new Date(),
    Session.getScriptTimeZone(),
    'yyyyMMdd-HHmmss'
  )}-${Math.floor(Math.random() * 10000)}`;
}


function getData_(sheet) {
  const lastRow = sheet.getLastRow();
  const lastColumn = sheet.getLastColumn();

  if (lastRow < 2 || lastColumn < 1) {
    return [];
  }

  return sheet
    .getRange(2, 1, lastRow - 1, lastColumn)
    .getValues();
}


function getHeaderMap_(sheet) {
  const headers = sheet
    .getRange(1, 1, 1, sheet.getLastColumn())
    .getValues()[0];

  const map = {};

  headers.forEach((header, index) => {
    map[normalizeText_(header)] = index;
  });

  return map;
}


function findRowByValue_(sheet, columnIndex, value) {
  const lastRow = sheet.getLastRow();

  if (lastRow < 2) {
    return -1;
  }

  const values = sheet
    .getRange(2, columnIndex + 1, lastRow - 1, 1)
    .getValues();

  const target = normalizeText_(value);

  for (let i = 0; i < values.length; i++) {
    if (normalizeText_(values[i][0]) === target) {
      return i + 2;
    }
  }

  return -1;
}


function assertSheetHeaders_(sheetName, expectedHeaders) {
  const sheet = getSheet_(sheetName);

  const actual = sheet
    .getRange(1, 1, 1, expectedHeaders.length)
    .getValues()[0]
    .map(normalizeText_);

  expectedHeaders.forEach((header, index) => {
    if (actual[index] !== header) {
      throw new Error(
        `ساختار شیت "${sheetName}" صحیح نیست. ستون "${header}" پیدا نشد.`
      );
    }
  });
}


function setStatusCell_(sheet, row, headerName, value) {
  const map = getHeaderMap_(sheet);

  if (map[headerName] === undefined) {
    throw new Error(`ستون "${headerName}" پیدا نشد.`);
  }

  sheet.getRange(row, map[headerName] + 1).setValue(value);
}


function getCellByHeader_(sheet, row, headerName) {
  const map = getHeaderMap_(sheet);

  if (map[headerName] === undefined) {
    throw new Error(`ستون "${headerName}" پیدا نشد.`);
  }

  return sheet.getRange(row, map[headerName] + 1).getValue();
}


function setCellByHeader_(sheet, row, headerName, value) {
  const map = getHeaderMap_(sheet);

  if (map[headerName] === undefined) {
    throw new Error(`ستون "${headerName}" پیدا نشد.`);
  }

  sheet.getRange(row, map[headerName] + 1).setValue(value);
}


function showMessage_(message) {
  SpreadsheetApp.getActiveSpreadsheet().toast(
    message,
    'مدیریت تولید و انبار',
    8
  );
}


