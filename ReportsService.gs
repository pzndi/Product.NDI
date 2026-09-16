/**
 * ============================================================
 * REPORTS & LOGGING
 * ============================================================
 */

function logError_(
  productionNumber,
  errorType,
  message
) {

  const sheet =
    getSheet_(CONFIG.SHEETS.ERRORS);

  sheet.appendRow([
    generateId_('ERR'),
    productionNumber,
    now_(),
    errorType,
    message,
    currentUser_(),
    now_()
  ]);
}


function showErrorReport() {

  const sheet =
    getSheet_(CONFIG.SHEETS.ERRORS);

  const lastRow =
    sheet.getLastRow();

  if (lastRow <= 1) {

    SpreadsheetApp.getUi().alert(
      'گزارش خطا',
      'هیچ خطایی ثبت نشده است.',
      SpreadsheetApp.getUi().ButtonSet.OK
    );

    return;
  }

  const errors = sheet
    .getRange(
      Math.max(2, lastRow - 9),
      1,
      Math.min(10, lastRow - 1),
      sheet.getLastColumn()
    )
    .getDisplayValues();

  const text = errors
    .map(row => {
      return `${row[1]} - ${row[3]}\n${row[4]}`;
    })
    .join('\n\n');

  SpreadsheetApp.getUi().alert(
    'آخرین خطاها',
    text,
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}


function showInventoryReport() {

  const sheet =
    getSheet_(CONFIG.SHEETS.INVENTORY);

  const headers =
    getHeaderMap_(sheet);

  const data =
    getData_(sheet);

  const warnings = [];

  data.forEach(row => {

    const name =
      row[headers['نام ماده']];

    const current =
      Number(
        row[headers['موجودی فعلی']] || 0
      );

    const minimum =
      Number(
        row[headers['حداقل موجودی']] || 0
      );

    if (current <= minimum) {

      warnings.push(
        `${name}: موجودی ${current} کیلوگرم`
      );
    }
  });

  if (warnings.length === 0) {

    SpreadsheetApp.getUi().alert(
      'گزارش موجودی',
      'تمام مواد اولیه بالاتر از حداقل موجودی هستند.',
      SpreadsheetApp.getUi().ButtonSet.OK
    );

    return;
  }

  SpreadsheetApp.getUi().alert(
    'مواد نیازمند تأمین',
    warnings.join('\n'),
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}


/**
 * این تابع در نسخه پایه برای بازسازی مصرف طراحی شده است.
 * برای جلوگیری از تغییر ناخواسته موجودی، عمداً نیازمند
 * تأیید کاربر است.
 */
function rebuildInventoryConsumption() {

  const ui = SpreadsheetApp.getUi();

  const response = ui.alert(
    'هشدار',
    'بازسازی مصرف می‌تواند مقادیر فعلی انبار را تغییر دهد.\n\nآیا مطمئن هستید؟',
    ui.ButtonSet.YES_NO
  );

  if (response !== ui.Button.YES) {
    return;
  }

  showMessage_(
    'این عملیات در نسخه پایه غیرفعال است. برای اصلاح گردش‌ها از ثبت برگشت استفاده کنید.'
  );
}
