
function onOpen(e) {
  putDefaultParamsOnTechDataSheet();
  createUi('☑ Остановить авто-обновление');
}


//triggers
function onChange(e) {
  let techSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('techdata');
  if (!techSheet) {
    Logger.log('Error: Лист "techdata" не обнаружен');
    return;
  }
  let techData = techSheet.getRange(1, 1, techSheet.getLastRow(), 2).getValues();
  let isAutoUpdate = {
    value: null
  };
  for (let i = 0; i < techData.length; i++) {
    if (techData[i][0].toString().startsWith('isAutoUpdate')) {
      isAutoUpdate.value = Boolean(techData[i][1]);
    }
  }
  if (isAutoUpdate.value === null) {
    Logger.log('Error: параметр "isAutoUpdate" не обнаружен в листе "techdata"');
    return;
  }

  if (isAutoUpdate.value) {
    let pivotSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Сводная');
    if (!pivotSheet) {
      Logger.log('Error: Листа "Сводная" не существует');
      return;
    }
    // create pivotTable
    let pivotTable = pivotSheet.getPivotTables();
    pivotTable = pivotTable[0];

    if (!pivotSheet) {
      Logger.log('Error: Сводной таблицы в листе "Сводная" не существует');
      return;
    }
    formatPivotTable(pivotSheet, pivotTable);
  }
}
