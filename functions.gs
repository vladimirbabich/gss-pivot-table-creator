function getColumnName(sheet1, column) {
  if (!sheet1) return;
  return column ? sheet1.getRange(`${column}1`).getValue() : null;
}

function getDataFromSheet(sheet) {
  const lastRow = sheet.getLastRow();
  const lastColumn = sheet.getLastColumn();
  const columnsRange = sheet.getRange(1, 1, 1, lastColumn);
  const columns = columnsRange.getValues()[0];
  Logger.log(columns)
  const dataRange = sheet.getRange(2, 1, lastRow - 1, lastColumn);
  const dataArr = dataRange.getValues();
  const resArray = dataArr.map(record => {
    const recObject = {};
    for (let i = 0; i < columns.length; i++) {
      recObject[columns[i]] = record[i];
    }
    return recObject;
  })
  return resArray;
}

class PivotTableSetup {
  constructor(rows) {//, cols, values, valuesFromFormulas) {
    this.rows = rows;
    // this.cols = cols;
    // this.values = values;
    // this.valuesFromFormulas = valuesFromFormulas;
  }

}
