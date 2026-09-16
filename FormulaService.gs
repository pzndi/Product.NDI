/**
 * ============================================================
 * FORMULA SERVICE
 * مدیریت و اعتبارسنجی فرمول‌های تولید
 * ============================================================
 */


/**
 * دریافت یک فرمول بر اساس کد فرمول
 */
function getFormula_(formulaCode) {

  const sheet = getSheet_(CONFIG.SHEETS.FORMULAS);

  const headers = getHeaderMap_(sheet);
  const data = getData_(sheet);

  const result = [];

  const targetCode = normalizeText_(formulaCode);

  data.forEach(function(row) {

    const code = normalizeText_(
      row[headers['کد فرمول']]
    );

    const active = normalizeText_(
      row[headers['فعال']]
    );

    if (
      code === targetCode &&
      active !== 'خیر'
    ) {

      result.push({
        formulaCode: code,

        formulaName: row[headers['نام فرمول']],

        customer: row[headers['مشتری']],

        product: row[headers['محصول']],

        materialCode: normalizeText_(
          row[headers['کد ماده']]
        ),

        materialName: row[headers['نام ماده']],

        quantity: toNumber_(
          row[headers['مقدار در 1000 کیلو']],
          'مقدار فرمول'
        )
      });

    }

  });


  if (result.length === 0) {

    throw new Error(
      'فرمول فعال با کد "' +
      formulaCode +
      '" پیدا نشد.'
    );

  }


  validateFormula_(result);

  return result;
}


/**
 * اعتبارسنجی یک فرمول
 *
 * شرط اصلی:
 * مجموع مواد اولیه باید دقیقاً 1000 کیلوگرم باشد.
 */
function validateFormula_(formulaRows) {

  if (!formulaRows || formulaRows.length === 0) {

    throw new Error(
      'فرمول فاقد مواد اولیه است.'
    );

  }


  let total = 0;

  const materials = {};


  formulaRows.forEach(function(item) {

    if (!item.materialCode) {

      throw new Error(
        'در فرمول "' +
        item.formulaCode +
        '" کد ماده اولیه خالی است.'
      );

    }


    if (!Number.isFinite(item.quantity)) {

      throw new Error(
        'مقدار ماده "' +
        item.materialName +
        '" عدد معتبر نیست.'
      );

    }


    if (item.quantity < 0) {

      throw new Error(
        'مقدار ماده "' +
        item.materialName +
        '" نمی‌تواند منفی باشد.'
      );

    }


    if (materials[item.materialCode]) {

      throw new Error(
        'ماده "' +
        item.materialCode +
        '" در فرمول "' +
        item.formulaCode +
        '" بیش از یک بار ثبت شده است.'
      );

    }


    materials[item.materialCode] = true;

    total += item.quantity;

  });


  total = roundNumber_(total);


  if (
    !approximatelyEqual_(
      total,
      CONFIG.FORMULA_TOTAL_KG
    )
  ) {

    throw new Error(
      'فرمول "' +
      formulaRows[0].formulaCode +
      '" معتبر نیست. ' +
      'مجموع فرمول برابر ' +
      total +
      ' کیلوگرم است؛ ' +
      'در حالی که باید دقیقاً ' +
      CONFIG.FORMULA_TOTAL_KG +
      ' کیلوگرم باشد.'
    );

  }


  return true;
}


/**
 * بررسی تمام فرمول‌های موجود در شیت
 */
function validateAllFormulas() {

  const sheet =
    getSheet_(CONFIG.SHEETS.FORMULAS);

  const headers =
    getHeaderMap_(sheet);

  const data =
    getData_(sheet);


  const formulas = {};


  data.forEach(function(row) {

    const code = normalizeText_(
      row[headers['کد فرمول']]
    );


    if (!code) {
      return;
    }


    if (!formulas[code]) {
      formulas[code] = [];
    }


    formulas[code].push({

      formulaCode: code,

      formulaName:
        row[headers['نام فرمول']],

      customer:
        row[headers['مشتری']],

      product:
        row[headers['محصول']],

      materialCode:
        normalizeText_(
          row[headers['کد ماده']]
        ),

      materialName:
        row[headers['نام ماده']],

      quantity:
        Number(
          row[headers['مقدار در 1000 کیلو']]
        )

    });

  });


  const formulaCodes =
    Object.keys(formulas);


  if (formulaCodes.length === 0) {

    SpreadsheetApp.getUi().alert(
      'بررسی فرمول‌ها',
      'هیچ فرمولی در شیت "فرمول تولید" پیدا نشد.',
      SpreadsheetApp.getUi().ButtonSet.OK
    );

    return;
  }


  const errors = [];


  formulaCodes.forEach(function(code) {

    try {

      validateFormula_(
        formulas[code]
      );

    } catch (error) {

      errors.push(
        error.message
      );

    }

  });


  if (errors.length === 0) {

    SpreadsheetApp.getUi().alert(
      'بررسی فرمول‌ها',
      'تمام ' +
      formulaCodes.length +
      ' فرمول موجود معتبر هستند.',
      SpreadsheetApp.getUi().ButtonSet.OK
    );

    return;
  }


  SpreadsheetApp.getUi().alert(
    'فرمول‌های دارای خطا',
    errors.join('\n\n'),
    SpreadsheetApp.getUi().ButtonSet.OK
  );

}
