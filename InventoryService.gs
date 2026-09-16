/**
 * ============================================================
 * INVENTORY SERVICE
 * ============================================================
 */

function getInventoryMap_() {

  const sheet = getSheet_(CONFIG.SHEETS.INVENTORY);

  const headers = getHeaderMap_(sheet);
  const data = getData_(sheet);

  const inventory = {};

  data.forEach((row, index) => {

    const code = normalizeText_(
      row[headers['کد ماده']]
    );

    if (!code) {
      return;
    }

    const rowNumber = index + 2;

    inventory[code] = {
      row: rowNumber,

      code: code,

      name: row[headers['نام ماده']],

      unit: row[headers['واحد']],

      initial: Number(
        row[headers['موجودی اولیه']
        ] || 0
      ),

      input: Number(
        row[headers['ورودی']] || 0
      ),

      output: Number(
        row[headers['مصرف']] || 0
      ),

      current: Number(
        row[headers['موجودی فعلی']] || 0
      ),

      minimum: Number(
        row[headers['حداقل موجودی']] || 0
      )
    };
  });

  return inventory;
}


function calculateProductionConsumption_(
  formulaRows,
  productionQuantity
) {

  const quantity = toNumber_(
    productionQuantity,
    'مقدار تولید'
  );

  if (quantity <= 0) {
    throw new Error(
      'مقدار تولید باید بیشتر از صفر باشد.'
    );
  }

  return formulaRows.map(item => {

    const consumption = roundNumber_(
      quantity *
      item.quantity /
      CONFIG.FORMULA_TOTAL_KG
    );

    return {
      materialCode: item.materialCode,
      materialName: item.materialName,
      formulaQuantity: item.quantity,
      productionQuantity: quantity,
      consumption: consumption
    };
  });
}


function validateInventoryAvailability_(
  consumption
) {

  const inventory = getInventoryMap_();

  const shortages = [];

  consumption.forEach(item => {

    const material = inventory[item.materialCode];

    if (!material) {

      shortages.push(
        `ماده "${item.materialCode} - ${item.materialName}" در انبار تعریف نشده است.`
      );

      return;
    }

    if (material.current < item.consumption) {

      shortages.push(
        `ماده: ${item.materialName}\n` +
        `موجودی: ${material.current} کیلوگرم\n` +
        `مورد نیاز: ${item.consumption} کیلوگرم\n` +
        `کسری: ${roundNumber_(
          item.consumption - material.current
        )} کیلوگرم`
      );
    }
  });

  if (shortages.length > 0) {

    throw new Error(
      'موجودی برای تولید کافی نیست:\n\n' +
      shortages.join('\n\n')
    );
  }

  return true;
}


function applyInventoryConsumption_(
  productionNumber,
  consumption
) {

  const sheet = getSheet_(CONFIG.SHEETS.INVENTORY);

  const headers = getHeaderMap_(sheet);

  const transactionsSheet =
    getSheet_(CONFIG.SHEETS.TRANSACTIONS);

  const inventory = getInventoryMap_();

  const user = currentUser_();
  const timestamp = now_();

  const transactionRows = [];

  consumption.forEach(item => {

    const material = inventory[item.materialCode];

    if (!material) {
      throw new Error(
        `ماده ${item.materialCode} در انبار پیدا نشد.`
      );
    }

    const before = material.current;

    const after = roundNumber_(
      before - item.consumption
    );

    const outputCell = sheet.getRange(
      material.row,
      headers['مصرف'] + 1
    );

    const oldOutput = Number(
      outputCell.getValue() || 0
    );

    outputCell.setValue(
      roundNumber_(
        oldOutput + item.consumption
      )
    );

    transactionRows.push([
      generateTransactionNumber_(),
      productionNumber,
      timestamp,
      item.materialCode,
      item.materialName,
      CONFIG.TRANSACTION_TYPES.PRODUCTION_CONSUMPTION,
      0,
      item.consumption,
      before,
      after,
      user,
      timestamp
    ]);
  });

  if (transactionRows.length > 0) {

    transactionsSheet
      .getRange(
        transactionsSheet.getLastRow() + 1,
        1,
        transactionRows.length,
        transactionRows[0].length
      )
      .setValues(transactionRows);
  }

  SpreadsheetApp.flush();
}

/**
 * ============================================================
 * ثبت ورود مواد اولیه به انبار
 * ============================================================
 */

function registerInventoryEntry() {

  const ui = SpreadsheetApp.getUi();

  const materialResponse = ui.prompt(
    'ثبت ورود مواد اولیه',
    'کد ماده اولیه را وارد کنید:',
    ui.ButtonSet.OK_CANCEL
  );

  if (materialResponse.getSelectedButton() !== ui.Button.OK) {
    return;
  }

  const materialCode =
    normalizeText_(
      materialResponse.getResponseText()
    );

  if (!materialCode) {
    ui.alert('خطا', 'کد ماده وارد نشده است.', ui.ButtonSet.OK);
    return;
  }

  const quantityResponse = ui.prompt(
    'ثبت ورود مواد اولیه',
    'مقدار ورود را به کیلوگرم وارد کنید:',
    ui.ButtonSet.OK_CANCEL
  );

  if (quantityResponse.getSelectedButton() !== ui.Button.OK) {
    return;
  }

  const quantity =
    toNumber_(
      quantityResponse.getResponseText(),
      'مقدار ورود'
    );

  if (quantity <= 0) {
    ui.alert(
      'خطا',
      'مقدار ورود باید بیشتر از صفر باشد.',
      ui.ButtonSet.OK
    );

    return;
  }

  const lock =
    LockService.getScriptLock();

  if (!lock.tryLock(30000)) {

    ui.alert(
      'سیستم مشغول است',
      'یک عملیات دیگر در حال اجراست. لطفاً چند لحظه بعد دوباره تلاش کنید.',
      ui.ButtonSet.OK
    );

    return;
  }

  try {

    const inventorySheet =
      getSheet_(CONFIG.SHEETS.INVENTORY);

    const headers =
      getHeaderMap_(inventorySheet);

    const row =
      findRowByValue_(
        inventorySheet,
        headers['کد ماده'],
        materialCode
      );

    if (row === -1) {

      throw new Error(
        `ماده با کد "${materialCode}" در انبار تعریف نشده است.`
      );
    }

    const materialName =
      getCellByHeader_(
        inventorySheet,
        row,
        'نام ماده'
      );

    const unit =
      getCellByHeader_(
        inventorySheet,
        row,
        'واحد'
      );

    const inputCell =
      inventorySheet.getRange(
        row,
        headers['ورودی'] + 1
      );

    const oldInput =
      Number(
        inputCell.getValue() || 0
      );

    const before =
      Number(
        getCellByHeader_(
          inventorySheet,
          row,
          'موجودی فعلی'
        ) || 0
      );

    const newInput =
      roundNumber_(
        oldInput + quantity
      );

    const after =
      roundNumber_(
        before + quantity
      );

    const timestamp = now_();
    const user = currentUser_();

    /*
     * ابتدا ورودی را ثبت می‌کنیم.
     * موجودی فعلی از فرمول شیت محاسبه می‌شود.
     */
    inputCell.setValue(newInput);

    SpreadsheetApp.flush();

    const calculatedAfter =
      Number(
        getCellByHeader_(
          inventorySheet,
          row,
          'موجودی فعلی'
        ) || 0
      );

    /*
     * کنترل نهایی:
     * موجودی بعد باید دقیقاً برابر
     * موجودی قبل + مقدار ورود باشد.
     */
    if (
      !approximatelyEqual_(
        calculatedAfter,
        after
      )
    ) {

      throw new Error(
        `محاسبه موجودی پس از ورود صحیح نیست. ` +
        `موجودی مورد انتظار: ${after}، ` +
        `موجودی محاسبه‌شده: ${calculatedAfter}`
      );
    }

    const transactionsSheet =
      getSheet_(CONFIG.SHEETS.TRANSACTIONS);

    transactionsSheet.appendRow([
      generateTransactionNumber_(),
      '',
      timestamp,
      materialCode,
      materialName,
      CONFIG.TRANSACTION_TYPES.INVENTORY_ENTRY,
      quantity,
      0,
      before,
      calculatedAfter,
      user,
      timestamp
    ]);

    SpreadsheetApp.flush();

    ui.alert(
      'ورود انبار ثبت شد',
      `ماده: ${materialName}\n` +
      `کد: ${materialCode}\n` +
      `مقدار ورود: ${quantity} ${unit}\n` +
      `موجودی قبل: ${before} ${unit}\n` +
      `موجودی بعد: ${calculatedAfter} ${unit}`,
      ui.ButtonSet.OK
    );

  } catch (error) {

    logError_(
      '',
      'خطای ورود انبار',
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

/**
 * ============================================================
 * برگشت کامل مصرف یک تولید
 * ============================================================
 */

function reverseProduction() {

  const ui =
    SpreadsheetApp.getUi();

  const response =
    ui.prompt(
      'برگشت تولید',
      'شماره تولیدی که قبلاً پردازش شده را وارد کنید:',
      ui.ButtonSet.OK_CANCEL
    );

  if (
    response.getSelectedButton() !==
    ui.Button.OK
  ) {
    return;
  }

  const productionNumber =
    normalizeText_(
      response.getResponseText()
    );

  if (!productionNumber) {

    ui.alert(
      'خطا',
      'شماره تولید وارد نشده است.',
      ui.ButtonSet.OK
    );

    return;
  }

  const lock =
    LockService.getScriptLock();

  if (!lock.tryLock(30000)) {

    ui.alert(
      'سیستم مشغول است',
      'یک عملیات دیگر در حال اجراست. لطفاً چند لحظه بعد دوباره تلاش کنید.',
      ui.ButtonSet.OK
    );

    return;
  }

  try {

    const productionSheet =
      getSheet_(CONFIG.SHEETS.PRODUCTION);

    const productionHeaders =
      getHeaderMap_(productionSheet);

    const productionRow =
      findRowByValue_(
        productionSheet,
        productionHeaders['شماره تولید'],
        productionNumber
      );

    if (productionRow === -1) {

      throw new Error(
        `تولید "${productionNumber}" پیدا نشد.`
      );
    }

    const productionStatus =
      normalizeText_(
        getCellByHeader_(
          productionSheet,
          productionRow,
          'وضعیت'
        )
      );

    if (
      productionStatus !==
      CONFIG.STATUS.PROCESSED
    ) {

      if (
        productionStatus ===
        CONFIG.STATUS.REVERSED
      ) {

        throw new Error(
          `تولید "${productionNumber}" قبلاً برگشت داده شده است.`
        );
      }

      throw new Error(
        `تولید "${productionNumber}" قابل برگشت نیست. ` +
        `وضعیت فعلی: ${productionStatus || 'خالی'}`
      );
    }

    /*
     * دریافت تمام تراکنش‌های مصرف مربوط به این تولید
     */
    const transactionsSheet =
      getSheet_(CONFIG.SHEETS.TRANSACTIONS);

    const transactionHeaders =
      getHeaderMap_(transactionsSheet);

    const transactionData =
      getData_(transactionsSheet);

    const consumptionTransactions = [];

    transactionData.forEach(function(row, index) {

      const rowNumber =
        index + 2;

      const trxProduction =
        normalizeText_(
          row[
            transactionHeaders['شماره تولید']
          ]
        );

      const operation =
        normalizeText_(
          row[
            transactionHeaders['نوع عملیات']
          ]
        );

      if (
        trxProduction === productionNumber &&
        operation ===
        normalizeText_(
          CONFIG.TRANSACTION_TYPES
            .PRODUCTION_CONSUMPTION
        )
      ) {

        const materialCode =
          normalizeText_(
            row[
              transactionHeaders['کد ماده']
            ]
          );

        const materialName =
          row[
            transactionHeaders['نام ماده']
          ];

        const quantity =
          Number(
            row[
              transactionHeaders['مقدار خروج']
            ] || 0
          );

        if (!materialCode) {

          throw new Error(
            `در تراکنش تولید "${productionNumber}" کد ماده خالی است.`
          );
        }

        if (
          !Number.isFinite(quantity) ||
          quantity <= 0
        ) {

          throw new Error(
            `مقدار مصرف ماده "${materialName}" معتبر نیست.`
          );
        }

        consumptionTransactions.push({
          row: rowNumber,
          materialCode: materialCode,
          materialName: materialName,
          quantity: quantity
        });
      }
    });

    if (
      consumptionTransactions.length === 0
    ) {

      throw new Error(
        `برای تولید "${productionNumber}" هیچ تراکنش مصرفی پیدا نشد.`
      );
    }

    /*
     * جلوگیری از برگشت دوباره:
     * اگر قبلاً تراکنش برگشت برای این تولید وجود داشته باشد،
     * عملیات متوقف می‌شود.
     */
    const reversalOperation =
      normalizeText_(
        CONFIG.TRANSACTION_TYPES.REVERSAL
      );

    const alreadyReversed =
      transactionData.some(function(row) {

        const trxProduction =
          normalizeText_(
            row[
              transactionHeaders['شماره تولید']
            ]
          );

        const operation =
          normalizeText_(
            row[
              transactionHeaders['نوع عملیات']
            ]
          );

        return (
          trxProduction === productionNumber &&
          operation === reversalOperation
        );
      });

    if (alreadyReversed) {

      throw new Error(
        `برای تولید "${productionNumber}" قبلاً تراکنش برگشت ثبت شده است.`
      );
    }

    /*
     * دریافت موجودی فعلی تمام مواد
     */
    const inventorySheet =
      getSheet_(CONFIG.SHEETS.INVENTORY);

    const inventoryHeaders =
      getHeaderMap_(inventorySheet);

    const inventoryData =
      getData_(inventorySheet);

    const inventoryMap = {};

    inventoryData.forEach(function(row, index) {

      const code =
        normalizeText_(
          row[
            inventoryHeaders['کد ماده']
          ]
        );

      if (!code) {
        return;
      }

      inventoryMap[code] = {
        row: index + 2,
        code: code,
        name:
          row[
            inventoryHeaders['نام ماده']
          ],
        current:
          Number(
            row[
              inventoryHeaders['موجودی فعلی']
            ] || 0
          ),
        output:
          Number(
            row[
              inventoryHeaders['مصرف']
            ] || 0
          )
      };
    });

    /*
     * مرحله اعتبارسنجی کامل قبل از هر تغییر
     */
    const changes = [];

    consumptionTransactions.forEach(
      function(item) {

        const material =
          inventoryMap[item.materialCode];

        if (!material) {

          throw new Error(
            `ماده "${item.materialCode}" در انبار پیدا نشد.`
          );
        }

        /*
         * برای برگشت مصرف:
         * موجودی فعلی + مقدار مصرف برگشتی
         *
         * و ستون مصرف:
         * مصرف قبلی - مقدار برگشتی
         */
        const newOutput =
          roundNumber_(
            material.output -
            item.quantity
          );

        if (newOutput < -0.000001) {

          throw new Error(
            `مقدار مصرف ماده "${material.name}" ` +
            `نمی‌تواند پس از برگشت منفی شود.`
          );
        }

        const before =
          material.current;

        const after =
          roundNumber_(
            before +
            item.quantity
          );

        changes.push({
          materialCode:
            item.materialCode,

          materialName:
            item.materialName,

          row:
            material.row,

          quantity:
            item.quantity,

          before:
            before,

          after:
            after,

          oldOutput:
            material.output,

          newOutput:
            Math.max(0, newOutput)
        });
      }
    );

    /*
     * تأیید نهایی کاربر
     */
    const confirmation =
      ui.alert(
        'تأیید برگشت تولید',
        `تولید "${productionNumber}" برگشت داده شود؟\n\n` +
        `تعداد مواد: ${changes.length}\n\n` +
        `این عملیات موجودی مواد را افزایش می‌دهد و قابل تکرار نیست.`,
        ui.ButtonSet.YES_NO
      );

    if (
      confirmation !==
      ui.Button.YES
    ) {
      return;
    }

    const timestamp =
      now_();

    const user =
      currentUser_();

    /*
     * نگهداری وضعیت قبلی برای Rollback
     */
    const rollbackData =
      changes.map(function(change) {

        return {
          row: change.row,
          oldOutput: change.oldOutput
        };

      });

    let transactionStartRow = null;

    try {

      /*
       * اعمال تغییرات موجودی
       */
      changes.forEach(
        function(change) {

          inventorySheet
            .getRange(
              change.row,
              inventoryHeaders['مصرف'] + 1
            )
            .setValue(
              change.newOutput
            );
        }
      );

      SpreadsheetApp.flush();

      /*
       * کنترل موجودی پس از تغییر
       */
      changes.forEach(
        function(change) {

          const calculatedCurrent =
            Number(
              getCellByHeader_(
                inventorySheet,
                change.row,
                'موجودی فعلی'
              ) || 0
            );

          if (
            !approximatelyEqual_(
              calculatedCurrent,
              change.after
            )
          ) {

            throw new Error(
              `موجودی ماده "${change.materialName}" ` +
              `پس از برگشت صحیح محاسبه نشده است.`
            );
          }
        }
      );

      /*
       * ثبت تراکنش‌های برگشت
       */
      const transactionRows =
        changes.map(function(change) {

          return [
            generateTransactionNumber_(),
            productionNumber,
            timestamp,
            change.materialCode,
            change.materialName,
            CONFIG.TRANSACTION_TYPES.REVERSAL,
            change.quantity,
            0,
            change.before,
            change.after,
            user,
            timestamp
          ];

        });

      transactionStartRow =
        transactionsSheet.getLastRow() + 1;

      transactionsSheet
        .getRange(
          transactionStartRow,
          1,
          transactionRows.length,
          transactionRows[0].length
        )
        .setValues(transactionRows);

      /*
       * تغییر وضعیت تولید
       */
      setCellByHeader_(
        productionSheet,
        productionRow,
        'وضعیت',
        CONFIG.STATUS.REVERSED
      );

      setCellByHeader_(
        productionSheet,
        productionRow,
        'زمان پردازش',
        timestamp
      );

      setCellByHeader_(
        productionSheet,
        productionRow,
        'پیام سیستم',
        `مصرف تولید ${productionNumber} با موفقیت برگشت داده شد.`
      );

      SpreadsheetApp.flush();

    } catch (operationError) {

      /*
       * Rollback موجودی
       */
      rollbackData.forEach(
        function(item) {

          inventorySheet
            .getRange(
              item.row,
              inventoryHeaders['مصرف'] + 1
            )
            .setValue(
              item.oldOutput
            );
        }
      );

      /*
       * حذف تراکنش‌هایی که ممکن است ثبت شده باشند
       */
      if (
        transactionStartRow !== null &&
        transactionsSheet.getLastRow() >=
        transactionStartRow
      ) {

        const rowsToDelete =
          Math.min(
            changes.length,
            transactionsSheet.getLastRow() -
            transactionStartRow +
            1
          );

        if (rowsToDelete > 0) {

          transactionsSheet.deleteRows(
            transactionStartRow,
            rowsToDelete
          );
        }
      }

      SpreadsheetApp.flush();

      throw operationError;
    }

    ui.alert(
      'برگشت موفق',
      `تولید "${productionNumber}" با موفقیت برگشت داده شد.\n\n` +
      `تعداد مواد برگشتی: ${changes.length}`,
      ui.ButtonSet.OK
    );

  } catch (error) {

    const message =
      error && error.message
        ? error.message
        : String(error);

    logError_(
      productionNumber,
      'خطای برگشت تولید',
      message
    );

    ui.alert(
      'خطا در برگشت تولید',
      message,
      ui.ButtonSet.OK
    );

  } finally {

    lock.releaseLock();
  }
}

function syncMaterialsToInventory() {

  const lock = LockService.getScriptLock();

  if (!lock.tryLock(30000)) {
    SpreadsheetApp.getUi().alert(
      'سیستم مشغول است',
      'یک عملیات دیگر در حال اجراست. لطفاً دوباره تلاش کنید.',
      SpreadsheetApp.getUi().ButtonSet.OK
    );
    return;
  }

  try {

    const materialsSheet =
      getSheet_(CONFIG.SHEETS.MATERIALS);

    const inventorySheet =
      getSheet_(CONFIG.SHEETS.INVENTORY);

    const materialHeaders =
      getHeaderMap_(materialsSheet);

    const inventoryHeaders =
      getHeaderMap_(inventorySheet);

    const materialData =
      getData_(materialsSheet);

    const inventoryData =
      getData_(inventorySheet);

    /*
     * مواد موجود در انبار را بر اساس کد ماده
     * در یک Map نگهداری می‌کنیم.
     */
    const inventoryMap = {};

    inventoryData.forEach(function(row, index) {

      const code =
        normalizeText_(
          row[inventoryHeaders['کد ماده']]
        );

      if (!code) {
        return;
      }

      inventoryMap[code] = {
        rowNumber: index + 2,
        row: row
      };
    });

    const rowsToAdd = [];

    let addedCount = 0;
    let updatedCount = 0;

    /*
     * مواد اولیه را می‌خوانیم.
     */
    materialData.forEach(function(row) {

      const code =
        normalizeText_(
          row[materialHeaders['کد ماده']]
        );

      if (!code) {
        return;
      }

      const name =
        row[materialHeaders['نام ماده']];

      const unit =
        row[materialHeaders['واحد']];

      const minimumStock =
        toNumber_(
          row[materialHeaders['حداقل موجودی']],
          'حداقل موجودی'
        );

      /*
       * اگر ماده قبلاً در انبار وجود دارد،
       * فقط اطلاعات تعریفی آن را به‌روزرسانی می‌کنیم.
       *
       * موجودی اولیه، ورودی و مصرف دست‌نخورده می‌مانند.
       */
      if (inventoryMap[code]) {

        const inventoryRow =
          inventoryMap[code].rowNumber;

        inventorySheet
          .getRange(
            inventoryRow,
            inventoryHeaders['نام ماده'] + 1
          )
          .setValue(name);

        inventorySheet
          .getRange(
            inventoryRow,
            inventoryHeaders['واحد'] + 1
          )
          .setValue(unit);

        inventorySheet
          .getRange(
            inventoryRow,
            inventoryHeaders['حداقل موجودی'] + 1
          )
          .setValue(minimumStock);

        updatedCount++;

        return;
      }

      /*
       * ماده جدید است.
       * یک ردیف جدید برای انبار ایجاد می‌کنیم.
       */
      const newRow =
        Array(CONFIG.HEADERS.INVENTORY.length)
          .fill('');

      newRow[
        inventoryHeaders['کد ماده']
      ] = code;

      newRow[
        inventoryHeaders['نام ماده']
      ] = name;

      newRow[
        inventoryHeaders['واحد']
      ] = unit;

      newRow[
        inventoryHeaders['موجودی اولیه']
      ] = 0;

      newRow[
        inventoryHeaders['ورودی']
      ] = 0;

      newRow[
        inventoryHeaders['مصرف']
      ] = 0;

      newRow[
        inventoryHeaders['حداقل موجودی']
      ] = minimumStock;

      rowsToAdd.push(newRow);

      addedCount++;
    });

    /*
     * اضافه کردن مواد جدید
     */
    if (rowsToAdd.length > 0) {

      const startRow =
        inventorySheet.getLastRow() + 1;

      inventorySheet
        .getRange(
          startRow,
          1,
          rowsToAdd.length,
          CONFIG.HEADERS.INVENTORY.length
        )
        .setValues(rowsToAdd);
    }

    /*
     * فرمول‌های موجودی فعلی و وضعیت را
     * برای تمام ردیف‌های انبار بازسازی می‌کنیم.
     *
     * این کار موجودی‌های قبلی را تغییر نمی‌دهد.
     */
    initializeInventoryFormulas_();

    SpreadsheetApp.flush();

    SpreadsheetApp.getUi().alert(
      'همگام‌سازی انجام شد',
      'مواد اولیه و انبار با موفقیت هماهنگ شدند.\n\n' +
      'مواد جدید اضافه‌شده: ' + addedCount + '\n' +
      'مواد موجود به‌روزرسانی‌شده: ' + updatedCount,
      SpreadsheetApp.getUi().ButtonSet.OK
    );

  } catch (error) {

    logError_(
      '',
      'خطای همگام‌سازی مواد اولیه با انبار',
      error.message || String(error)
    );

    SpreadsheetApp.getUi().alert(
      'خطا',
      error.message || String(error),
      SpreadsheetApp.getUi().ButtonSet.OK
    );

  } finally {

    lock.releaseLock();
  }
}
