var scriptProperties = PropertiesService.getScriptProperties();
var userProperties = PropertiesService.getUserProperties();
var documentProperties = PropertiesService.getDocumentProperties();

function createPivotTable() {
  let sheet = SpreadsheetApp.getActive().getSheetByName('export FB');
  let lastRow = sheet.getLastRow();
  let lastColumn = sheet.getLastColumn();
  let allRange = sheet.getRange(1, 1, lastRow, lastColumn);
  let dataArr = allRange.getValues();

  let columnNamesForRowGroup = {//key - column name, value - index in table(not in data)
    'Ad set name': -1,
    'Reporting starts': -1
  };
  let columnNamesForPivotValue = {
    'Amount spent (RUB)': -1,
    'Link clicks': -1,
    'Leads': -1
  };
  let metricsForCalcPivotValue = {
    'CPC': `=IFERROR('Amount spent (RUB)'/'Link clicks'; "-")`,
    'CPL': `=IFERROR('Amount spent (RUB)'/'Leads';"-")`
  }
  findIndexes(sheet, columnNamesForRowGroup);
  findIndexes(sheet, columnNamesForPivotValue);

  let stringsToReplaceArr = {
    ' ₽': '',
    '(Пусто)': 0
  }
  for (const el in stringsToReplaceArr) {
    replaceSymbolsInArray(dataArr, el, stringsToReplaceArr[el]);
  }

  allRange.setValues(dataArr);

  //create or get pivotSheet and clear it
  let pivotSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Сводная');
  pivotSheet = pivotSheet != null ? pivotSheet : SpreadsheetApp.getActive().insertSheet('Сводная', 1);
  pivotSheet.clear();
  // create pivotTable
  let pivotTable = pivotSheet.getRange(1, 1).createPivotTable(allRange);

  //create every row/column/values
  for (el in columnNamesForRowGroup) {
    pivotTable.addRowGroup(columnNamesForRowGroup[el]);
  }
  for (el in columnNamesForPivotValue) {
    pivotTable.addPivotValue(columnNamesForPivotValue[el], SpreadsheetApp.PivotTableSummarizeFunction.SUM).setDisplayName(el);
  }
  for (el in metricsForCalcPivotValue) {
    pivotTable.addCalculatedPivotValue(el, metricsForCalcPivotValue[el]);
  }

  //formatting
  formatPivotTable(pivotSheet, pivotTable);

}

function findIndexes(pivotSheet, columnNames) {
  //need to find indexes of all columns for pivottable
  let firstRowArr = pivotSheet.getRange(1, 1, 1, pivotSheet.getLastColumn()).getValues()[0];
  let columnIndexes = {};
  // for(let j=0; j<columnNames.length; j++){
  for (let columnNameEl in columnNames) {
    for (let i = 0; i < firstRowArr.length; i++) {
      if (firstRowArr[i].startsWith(columnNameEl)) {
        columnNames[columnNameEl] = i + 1;
      }
    }
  }
}

function formatPivotTable(pivotSheet, pivotTable) {
  let anchorCellRowIndex = pivotTable.getAnchorCell().getRowIndex();
  let anchorCellColumnIndex = pivotTable.getAnchorCell().getColumnIndex();
  let rowGroupLength = pivotTable.getRowGroups().length;//column offset

  //format first row
  pivotSheet
    .getRange(anchorCellRowIndex, anchorCellColumnIndex, 1, pivotSheet.getLastColumn())
    .setWrap(true)
    .setVerticalAlignment('middle')
    .setHorizontalAlignment('center'); // 1st row outline

  // format all rowGroups
  pivotSheet
    .getRange(anchorCellRowIndex + 1, 1, pivotSheet.getLastRow() - anchorCellRowIndex - 1 + 1, rowGroupLength)
    .setHorizontalAlignment('left')
    .setNumberFormat('yyyy.MM');
  // format all pivotValue & calcPivotValue
  pivotSheet
    .getRange(anchorCellRowIndex + 1, rowGroupLength + 1, pivotSheet.getLastRow() - anchorCellRowIndex - 1 + 1, pivotSheet.getLastColumn() - rowGroupLength)
    .setHorizontalAlignment('right')
    .setNumberFormat('#,##0'); // ##,##### -> ## : 11,14

  //freeze first row of the pivotTable
  pivotSheet.setFrozenRows(anchorCellRowIndex);
  //freeze columns(rowgroups) of the pivotTable
  pivotSheet.setFrozenColumns(rowGroupLength);
}

function replaceSymbolsInArray(arr, find, replace) {
  for (let i = 0; i < arr.length; i++) {
    for (let j = 0; j < arr[0].length; j++) {
      if (typeof arr[i][j] === 'string' || arr[i][j] instanceof String) {
        if (arr[i][j].includes(find)) {
          arr[i][j] = arr[i][j].replace(find, replace);
        }
      }
    }
  }
}

function toggleIsAutoUpdate() {
  let techSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('techdata');
  if (!techSheet) {
    Logger.log('Error: Лист "techdata" не обнаружен');
  }
  let techData = techSheet.getRange(1, 1, techSheet.getLastRow(), 2).getValues();
  let isAutoUpdate = {
    value: null,
    xIndex: -1,
    yIndex: -1
  };

  for (let i = 0; i < techData.length; i++) {
    if (techData[i][0].toString().startsWith('isAutoUpdate')) {
      isAutoUpdate.value = Boolean(techData[i][1]);
      isAutoUpdate.xIndex = i;
      isAutoUpdate.yIndex = 1;
    }
  }
  if (isAutoUpdate.value === null) {
    Logger.log('Error: параметр "isAutoUpdate" не обнаружен в листе "techdata"');
    return;
  }
  let menuItemName;
  if (isAutoUpdate.value) {
    menuItemName = '☐ Включить авто-обновление сводной таблицы';
  } else {
    menuItemName = '☑ Остановить авто-обновление сводной таблицы';
  }
  isAutoUpdate.value = !isAutoUpdate.value;
  techSheet.getRange(isAutoUpdate.xIndex + 1, isAutoUpdate.yIndex + 1).setValue(isAutoUpdate.value);//+1 because numeration in sheet
  SpreadsheetApp.getUi()
    .createMenu('Функции')
    .addItem('Создать сводную автоматически', 'createPivotTable')
    .addSeparator()
    .addItem(menuItemName, 'toggleIsAutoUpdate')
    .addToUi();
}


function putDefaultParamsOnTechDataSheet() {
  //create or get techSheet
  let techSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('techdata');
  techSheet = techSheet != null ? techSheet : SpreadsheetApp.getActive().insertSheet('techdata', 1);
  // techSheet.hideSheet();

  let defaultParams = [
    ['variable', 'value'],//not parameter, only to name columns
    ['isAutoUpdate', true],
    ['test', 123]
  ];
  techSheet.getRange(1, 1, defaultParams.length, defaultParams[0].length).setValues(defaultParams);//+1 because numeration in sheet
}
