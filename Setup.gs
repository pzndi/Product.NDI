/**
 * ============================================================
 * SETUP
 * ایجاد و آماده‌سازی شیت‌ها
 * ============================================================
 */

function setupSystem() {

  const ss = getSpreadsheet_();

  const definitions = [
    [CONFIG.SHEETS.SETTINGS, CONFIG.HEADERS.SETTINGS],
    [CONFIG.SHEETS.MATERIALS, CONFIG.HEADERS.MATERIALS],
    [CONFIG.SHEETS.FORMULAS, CONFIG.HEADERS.FORMULAS],
    [CONFIG.SHEETS.PRODUCTION, CONFIG.HEADERS.PRODUCTION],
    [CONFIG.SHEETS.INVENTORY, CONFIG.HEADERS.INVENTORY],
    [CONFIG.SHEETS.TRANSACTIONS, CONFIG.HEADERS.TRANSACTIONS],
    [CONFIG.SHEETS.ERRORS, CONFIG.HEADERS.ERRORS]
  ];

  definitions.forEach(([sheetName, headers]) => {

    let sheet = ss.getSheetByName(sheetName);

    if (!sheet) {
      sheet = ss.insertSheet(sheetName);
    }

    initializeSheet_(sheet, headers);
  });

  initializeSettings_();
  initializeInventoryFormulas_();
  formatSystemColumns_();
  protectSystemColumns_();
  SpreadsheetApp.flush();

  showMessage_('سیستم با موفقیت آماده شد.');
}


function initializeSheet_(sheet, headers) {

  if (sheet.getLastRow() === 0) {
    sheet.getRange(1, 1, 1, headers.length)
      .setValues([headers]);
  } else {
    const currentHeaders = sheet
      .getRange(1, 1, 1, headers.length)
      .getValues()[0];

    let changed = false;

    headers.forEach((header, index) => {
      if (normalizeText_(currentHeaders[index]) !== header) {
        currentHeaders[index] = header;
        changed = true;
      }
    });

    if (changed) {
      sheet
        .getRange(1, 1, 1, headers.length)
        .setValues([currentHeaders]);
    }
  }

  sheet.setFrozenRows(1);

  const headerRange = sheet.getRange(
    1,
    1,
    1,
    headers.length
  );

  headerRange.setFontWeight('bold');

  if (sheet.getLastRow() === 1) {
    sheet.autoResizeColumns(1, headers.length);
  }
}


function initializeSettings_() {

  const sheet = getSheet_(CONFIG.SHEETS.SETTINGS);

  if (sheet.getLastRow() > 1) {
    return;
  }

  const rows = [
    ['نام سیستم', 'مدیریت تولید و انبار', 'نام سیستم'],
    ['وزن پایه فرمول', CONFIG.FORMULA_TOTAL_KG, 'تمام فرمول‌ها باید بر اساس این مقدار باشند'],
    ['واحد پایه', 'کیلوگرم', 'واحد استاندارد سیستم']
  ];

  sheet
    .getRange(2, 1, rows.length, rows[0].length)
    .setValues(rows);
}


function initializeInventoryFormulas_() {

  const sheet =
    getSheet_(CONFIG.SHEETS.INVENTORY);

  const lastRow =
    sheet.getLastRow();

  if (lastRow < 2) {
    return;
  }

  const headers =
    getHeaderMap_(sheet);

  const initialCol =
    headers['موجودی اولیه'] + 1;

  const inputCol =
    headers['ورودی'] + 1;

  const outputCol =
    headers['مصرف'] + 1;

  const currentCol =
    headers['موجودی فعلی'] + 1;

  const minCol =
    headers['حداقل موجودی'] + 1;

  const statusCol =
    headers['وضعیت'] + 1;

  for (
    let row = 2;
    row <= lastRow;
    row++
  ) {

    const materialCode =
      normalizeText_(
        sheet.getRange(
          row,
          headers['کد ماده'] + 1
        ).getValue()
      );

    if (!materialCode) {
      continue;
    }

    const formula =
      `=N(${columnLetter_(initialCol)}${row})` +
      `+N(${columnLetter_(inputCol)}${row})` +
      `-N(${columnLetter_(outputCol)}${row})`;

    sheet
      .getRange(row, currentCol)
      .setFormula(formula);

    sheet
      .getRange(row, statusCol)
      .setFormula(
        `=IF(${columnLetter_(currentCol)}${row}` +
        `<=${columnLetter_(minCol)}${row},` +
        `"نیاز به تأمین","عادی")`
      );
  }

  SpreadsheetApp.flush();
}

function formatSystemColumns_() {

  const sheet =
    getSheet_(CONFIG.SHEETS.PRODUCTION);

  const headers =
    getHeaderMap_(sheet);

  const systemColumns = [
    'شماره تولید',
    'تاریخ',
    'وضعیت',
    'زمان پردازش',
    'پیام سیستم'
  ];

  // قرمز بسیار کم‌رنگ
  const systemColor = '#FCE8E6';

  systemColumns.forEach(function(columnName) {

    const columnIndex =
      headers[columnName];

    if (columnIndex === undefined) {
      return;
    }

    const range =
      sheet.getRange(
        1,
        columnIndex + 1,
        sheet.getMaxRows(),
        1
      );

    range.setBackground(systemColor);
  });
}


function columnLetter_(column) {

  let letter = '';

  while (column > 0) {
    const temp = (column - 1) % 26;

    letter = String.fromCharCode(temp + 65) + letter;

    column = Math.floor((column - temp) / 26);
  }

  return letter;
}

/**
 * ============================================================
 * PROTECTION
 * حفاظت از ستون‌های سیستمی
 * ============================================================
 */

function protectSystemColumns_() {

  const sheet =
    getSheet_(CONFIG.SHEETS.PRODUCTION);

  const headers =
    getHeaderMap_(sheet);

  const systemColumns = [
    'شماره تولید',
    'تاریخ',
    'وضعیت',
    'زمان پردازش',
    'پیام سیستم'
  ];

  /*
   * حذف Protectionهای قبلی این شیت
   * تا تنظیمات جدید بدون تداخل اعمال شوند.
   */
  const protections =
    sheet.getProtections(
      SpreadsheetApp.ProtectionType.RANGE
    );

  protections.forEach(function(protection) {

    const range =
      protection.getRange();

    if (
      range.getSheet().getSheetId() ===
      sheet.getSheetId()
    ) {
      protection.remove();
    }
  });

  systemColumns.forEach(function(columnName) {

    const columnIndex =
      headers[columnName];

    if (columnIndex === undefined) {
      return;
    }

    const range =
      sheet.getRange(
        2,
        columnIndex + 1,
        Math.max(
          sheet.getMaxRows() - 1,
          1
        ),
        1
      );

    const protection =
      range.protect();

    protection.setDescription(
      `سیستمی - ${columnName}`
    );

    /*
     * فقط مالک فایل اجازه ویرایش دارد.
     */
    protection.setWarningOnly(false);

    const editors =
      protection.getEditors();

    if (editors.length > 0) {
      protection.removeEditors(editors);
    }

    if (
      protection.canDomainEdit()
    ) {
      protection.setDomainEdit(false);
    }
  });
}

function syncFormulaMaterialColumns_() {

  const inventorySheet =
    getSheet_(CONFIG.SHEETS.INVENTORY);

  const formulaSheet =
    getSheet_(CONFIG.SHEETS.FORMULAS);

  const inventoryHeaders =
    getHeaderMap_(inventorySheet);

  const formulaHeaders =
    getHeaderMap_(formulaSheet);

  const inventoryData =
    getData_(inventorySheet);

  /*
   * مواد اولیه را از انبار می‌خوانیم.
   */
  const materialNames = [];
  const materialCodes = {};

  inventoryData.forEach(function(row) {

    const code =
      normalizeText_(
        row[inventoryHeaders['کد ماده']]
      );

    const name =
      String(
        row[inventoryHeaders['نام ماده']] || ''
      ).trim();

    if (!code || !name) {
      return;
    }

    /*
     * جلوگیری از ایجاد دو ستون با نام یکسان
     */
    if (materialCodes[code]) {
      return;
    }

    materialCodes[code] = true;

    materialNames.push(name);
  });

  /*
   * ستون‌های ثابت فرمول
   */
  const fixedColumns = [
    'کد فرمول',
    'نام فرمول',
    'مشتری',
    'محصول'
  ];

  /*
   * نام تمام ستون‌های فعلی
   */
  const lastColumn =
    formulaSheet.getLastColumn();

  const currentHeaders =
    formulaSheet
      .getRange(
        1,
        1,
        1,
        Math.max(lastColumn, 1)
      )
      .getValues()[0];

  /*
   * فرمول‌های موجود را بر اساس نام ستون نگه می‌داریم
   * تا هنگام جابه‌جایی ستون‌ها مقدارها از بین نروند.
   */
  const data =
    formulaSheet.getLastRow() >= 2
      ? formulaSheet
          .getRange(
            2,
            1,
            formulaSheet.getLastRow() - 1,
            Math.max(lastColumn, 1)
          )
          .getValues()
      : [];

  const oldMap = {};

  currentHeaders.forEach(function(header, index) {

    const name =
      String(header || '').trim();

    if (!name) {
      return;
    }

    oldMap[name] = index;
  });

  /*
   * ساختار جدید
   */
  const newHeaders =
    fixedColumns.concat(materialNames);

  /*
   * اگر تعداد ستون‌ها کمتر است، اضافه می‌کنیم.
   */
  if (
    formulaSheet.getMaxColumns() <
    newHeaders.length
  ) {

    formulaSheet.insertColumnsAfter(
      formulaSheet.getMaxColumns(),
      newHeaders.length -
      formulaSheet.getMaxColumns()
    );
  }

  /*
   * سرستون‌های جدید
   */
  formulaSheet
    .getRange(
      1,
      1,
      1,
      newHeaders.length
    )
    .setValues([newHeaders]);

  /*
   * بازسازی داده‌ها بدون از دست دادن مقادیر قبلی
   */
  if (data.length > 0) {

    const newData =
      data.map(function(oldRow) {

        return newHeaders.map(function(header) {

          const oldIndex =
            oldMap[header];

          if (
            oldIndex === undefined ||
            oldIndex >= oldRow.length
          ) {
            return '';
          }

          return oldRow[oldIndex];
        });

      });

    formulaSheet
      .getRange(
        2,
        1,
        newData.length,
        newHeaders.length
      )
      .setValues(newData);
  }

  /*
   * ستون‌های اضافی قدیمی را پاک می‌کنیم.
   */
  const currentLastColumn =
    formulaSheet.getLastColumn();

  if (
    currentLastColumn >
    newHeaders.length
  ) {

    formulaSheet
      .deleteColumns(
        newHeaders.length + 1,
        currentLastColumn -
        newHeaders.length
      );
  }

  /*
   * قالب‌بندی
   */
  formulaSheet
    .getRange(
      1,
      1,
      1,
      newHeaders.length
    )
    .setFontWeight('bold');

  SpreadsheetApp.flush();
}
