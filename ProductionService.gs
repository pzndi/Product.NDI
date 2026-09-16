/**
 * ============================================================
 * PRODUCTION SERVICE
 * پردازش تولید
 * ============================================================
 */

function processNewProductions() {

  const lock = LockService.getScriptLock();

  if (!lock.tryLock(30000)) {

    SpreadsheetApp.getUi().alert(
      'سیستم مشغول است',
      'یک عملیات دیگر در حال اجراست. لطفاً چند لحظه بعد دوباره تلاش کنید.',
      SpreadsheetApp.getUi().ButtonSet.OK
    );

    return;
  }

  try {

    const productionSheet =
      getSheet_(CONFIG.SHEETS.PRODUCTION);

    assertSheetHeaders_(
      CONFIG.SHEETS.PRODUCTION,
      CONFIG.HEADERS.PRODUCTION
    );

    const data = getData_(productionSheet);

    if (data.length === 0) {

      showMessage_(
        'هیچ رکورد تولیدی برای پردازش وجود ندارد.'
      );

      return;
    }

    const headers = getHeaderMap_(
      productionSheet
    );

    let processed = 0;
    let errors = 0;
    let skipped = 0;

    data.forEach((row, index) => {

      const rowNumber = index + 2;

      const productionNumber =
        normalizeText_(
          row[headers['شماره تولید']]
        );

      const status =
        normalizeText_(
          row[headers['وضعیت']]
        );

      if (!productionNumber) {
        return;
      }

      if (
        status === CONFIG.STATUS.PROCESSED
      ) {
        skipped++;
        return;
      }

      if (
        status === CONFIG.STATUS.CANCELLED
      ) {
        skipped++;
        return;
      }

      try {

        processProductionRow_(
          productionSheet,
          rowNumber
        );

        processed++;

      } catch (error) {

        errors++;

        handleProductionError_(
          productionSheet,
          rowNumber,
          error
        );
      }
    });

    SpreadsheetApp.flush();

    SpreadsheetApp.getUi().alert(
      'نتیجه پردازش',
      `پردازش با موفقیت انجام شد.\n\n` +
      `تعداد پردازش موفق: ${processed}\n` +
      `تعداد خطا: ${errors}\n` +
      `تعداد رد شده/قبلی: ${skipped}`,
      SpreadsheetApp.getUi().ButtonSet.OK
    );

  } finally {

    lock.releaseLock();
  }
}


function processProductionRow_(
  productionSheet,
  rowNumber
) {

  const productionNumber =
    normalizeText_(
      getCellByHeader_(
        productionSheet,
        rowNumber,
        'شماره تولید'
      )
    );

  const formulaCode =
    normalizeText_(
      getCellByHeader_(
        productionSheet,
        rowNumber,
        'کد فرمول'
      )
    );

  const productionQuantity =
    toNumber_(
      getCellByHeader_(
        productionSheet,
        rowNumber,
        'مقدار تولید'
      ),
      'مقدار تولید'
    );

  const customer =
    normalizeText_(
      getCellByHeader_(
        productionSheet,
        rowNumber,
        'مشتری'
      )
    );

  if (!productionNumber) {
    throw new Error(
      'شماره تولید خالی است.'
    );
  }

  if (!formulaCode) {
    throw new Error(
      'کد فرمول خالی است.'
    );
  }

  if (productionQuantity <= 0) {
    throw new Error(
      'مقدار تولید باید بیشتر از صفر باشد.'
    );
  }

  if (isProductionAlreadyProcessed_(
    productionNumber
  )) {

    throw new Error(
      `تولید ${productionNumber} قبلاً پردازش شده است.`
    );
  }

  const formulaRows =
    getFormula_(formulaCode);

  const consumption =
    calculateProductionConsumption_(
      formulaRows,
      productionQuantity
    );

  validateInventoryAvailability_(
    consumption
  );

  /*
   * در این نقطه تمام اعتبارسنجی‌ها انجام شده‌اند.
   * سپس موجودی مصرف می‌شود.
   */

  applyInventoryConsumption_(
    productionNumber,
    consumption
  );

  setCellByHeader_(
    productionSheet,
    rowNumber,
    'وضعیت',
    CONFIG.STATUS.PROCESSED
  );

  setCellByHeader_(
    productionSheet,
    rowNumber,
    'زمان پردازش',
    now_()
  );

  setCellByHeader_(
    productionSheet,
    rowNumber,
    'پیام سیستم',
    `تولید ${productionQuantity} کیلوگرم برای ${customer} با موفقیت پردازش شد.`
  );
}


function isProductionAlreadyProcessed_(
  productionNumber
) {

  const sheet =
    getSheet_(CONFIG.SHEETS.PRODUCTION);

  const headers =
    getHeaderMap_(sheet);

  const data =
    getData_(sheet);

  const target =
    normalizeText_(productionNumber);

  for (let i = 0; i < data.length; i++) {

    const number =
      normalizeText_(
        data[i][headers['شماره تولید']]
      );

    const status =
      normalizeText_(
        data[i][headers['وضعیت']]
      );

    if (
      number === target &&
      (
        status === CONFIG.STATUS.PROCESSED ||
        status === CONFIG.STATUS.REVERSED
      )
    ) {
      return true;
    }
  }

  return false;
}

function handleProductionError_(
  productionSheet,
  rowNumber,
  error
) {

  const message =
    error && error.message
      ? error.message
      : String(error);

  setCellByHeader_(
    productionSheet,
    rowNumber,
    'وضعیت',
    CONFIG.STATUS.ERROR
  );

  setCellByHeader_(
    productionSheet,
    rowNumber,
    'زمان پردازش',
    now_()
  );

  setCellByHeader_(
    productionSheet,
    rowNumber,
    'پیام سیستم',
    message
  );

  logError_(
    getCellByHeader_(
      productionSheet,
      rowNumber,
      'شماره تولید'
    ),
    'خطای پردازش تولید',
    message
  );
}

/**
 * ============================================================
 * ثبت تولید جدید
 * شماره تولید به صورت خودکار ایجاد می‌شود.
 * ============================================================
 */

function registerProduction() {

  const ui =
    SpreadsheetApp.getUi();

  const customerResponse =
    ui.prompt(
      'ثبت تولید',
      'نام مشتری را وارد کنید:',
      ui.ButtonSet.OK_CANCEL
    );

  if (
    customerResponse.getSelectedButton() !==
    ui.Button.OK
  ) {
    return;
  }

  const customer =
    normalizeText_(
      customerResponse.getResponseText()
    );

  if (!customer) {

    ui.alert(
      'خطا',
      'نام مشتری وارد نشده است.',
      ui.ButtonSet.OK
    );

    return;
  }

  const formulaResponse =
    ui.prompt(
      'ثبت تولید',
      'کد فرمول را وارد کنید:',
      ui.ButtonSet.OK_CANCEL
    );

  if (
    formulaResponse.getSelectedButton() !==
    ui.Button.OK
  ) {
    return;
  }

  const formulaCode =
    normalizeText_(
      formulaResponse.getResponseText()
    );

  if (!formulaCode) {

    ui.alert(
      'خطا',
      'کد فرمول وارد نشده است.',
      ui.ButtonSet.OK
    );

    return;
  }

  /*
   * اعتبارسنجی فرمول قبل از ثبت تولید
   */
  getFormula_(formulaCode);

  const quantityResponse =
    ui.prompt(
      'ثبت تولید',
      'مقدار تولید را به کیلوگرم وارد کنید:',
      ui.ButtonSet.OK_CANCEL
    );

  if (
    quantityResponse.getSelectedButton() !==
    ui.Button.OK
  ) {
    return;
  }

  const quantity =
    toNumber_(
      quantityResponse.getResponseText(),
      'مقدار تولید'
    );

  if (quantity <= 0) {

    ui.alert(
      'خطا',
      'مقدار تولید باید بیشتر از صفر باشد.',
      ui.ButtonSet.OK
    );

    return;
  }

  const lock =
    LockService.getScriptLock();

  if (!lock.tryLock(30000)) {

    ui.alert(
      'سیستم مشغول است',
      'یک عملیات دیگر در حال اجراست. لطفاً دوباره تلاش کنید.',
      ui.ButtonSet.OK
    );

    return;
  }

  try {

    const sheet =
      getSheet_(CONFIG.SHEETS.PRODUCTION);

    const headers =
      getHeaderMap_(sheet);

    /*
     * تولید شماره یکتا
     */
    const productionNumber =
      generateProductionNumber_();

    /*
     * کنترل نهایی یکتا بودن
     */
    const existingRow =
      findRowByValue_(
        sheet,
        headers['شماره تولید'],
        productionNumber
      );

    if (existingRow !== -1) {

      throw new Error(
        `شماره تولید "${productionNumber}" قبلاً وجود دارد.`
      );
    }

    const row =
      sheet.getLastRow() + 1;

    sheet
      .getRange(row, 1, 1, CONFIG.HEADERS.PRODUCTION.length)
      .setValues([[
        productionNumber,
        now_(),
        customer,
        formulaCode,
        quantity,
        CONFIG.STATUS.NEW,
        '',
        'آماده پردازش'
      ]]);

    SpreadsheetApp.flush();

    ui.alert(
      'تولید ثبت شد',
      `شماره تولید: ${productionNumber}\n` +
      `مشتری: ${customer}\n` +
      `کد فرمول: ${formulaCode}\n` +
      `مقدار تولید: ${quantity} کیلوگرم\n\n` +
      `وضعیت: ${CONFIG.STATUS.NEW}`,
      ui.ButtonSet.OK
    );

  } catch (error) {

    logError_(
      '',
      'خطای ثبت تولید',
      error.message || String(error)
    );

    ui.alert(
      'خطا',
      error.message || String(error),
      ui.ButtonSet.OK
    );

  } finally {

    lock.releaseLock();
  }
}
