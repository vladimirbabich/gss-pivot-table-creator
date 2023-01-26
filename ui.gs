function createUi(autoUpdateText) {
  SpreadsheetApp.getUi()
    .createMenu('Функции')
    .addItem('Создать сводную автоматически', 'createPivotTable')
    .addSeparator()
    .addItem(autoUpdateText, 'toggleIsAutoUpdate')
    .addToUi();
}
