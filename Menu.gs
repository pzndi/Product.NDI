function onOpen() {

  SpreadsheetApp.getUi()
    .createMenu('مدیریت تولید و انبار')
    .addItem(
      '⚙ آماده‌سازی سیستم',
      'setupSystem'
    )
    .addSeparator()
    .addItem(
      '➕ ثبت تولید جدید',
      'registerProduction'
    )
    .addItem(
      '▶ پردازش تولیدهای جدید',
      'processNewProductions'
    )
    .addItem(
      '↩ برگشت تولید',
      'reverseProduction'
    )
    .addItem(
      '➕ ثبت ورود مواد اولیه',
      'registerInventoryEntry'
    )
    .addItem(
      '🔄 همگام‌سازی مواد اولیه با انبار',
      'syncMaterialsToInventory'
    )
    .addItem(
      '🔍 بررسی فرمول‌ها',
      'validateAllFormulas'
    )
    .addItem(
      '📦 بروزرسانی انبار',
      'rebuildInventoryConsumption'
    )
    .addSeparator()
    .addItem(
      '📊 گزارش موجودی',
      'showInventoryReport'
    )
    .addItem(
      '⚠ گزارش خطاها',
      'showErrorReport'
    )
    .addToUi();
}
