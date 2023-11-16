function FormatCalReport(sheet) {
  DeleteUnusedRowsAndColumns(sheet);
  
  let SRC_CHARGING_DATE_COLUMN = 5;
  let DST_CHARGING_COLUMN = 1;
  let BUYING_DATE_COLUMN = 2;

  MoveColumn(sheet, SRC_CHARGING_DATE_COLUMN, DST_CHARGING_COLUMN);
  FormatDate(sheet, DST_CHARGING_COLUMN);
  SortTableByAscendingDate(sheet, BUYING_DATE_COLUMN);
  SetTableStyle(sheet)

  Logger.log("FormatCalReport() is finished.");
}

function DeleteUnusedRowsAndColumns(sheet) {
  sheet.deleteRows(1, 2);
  sheet.deleteRow(sheet.getLastRow());
  sheet.deleteColumns(sheet.getLastColumn() - 1, 2);
}

function SetTableStyle(sheet) {
  let table = sheet.getDataRange();
  
  table.setFontFamily("Ariel");
  table.setFontSize(10);
  table.setBorder(false, false, false, false, false, false);
}

function FormatDate(sheet, dateColumn) {
  let data = sheet.getRange(dateColumn, 1, sheet.getLastRow(), 1).getValues();

  for (let i = 0; i < data.length; i++) {
    let dateValue = data[i][0];
    
    if (dateValue instanceof Date) {
      let year = dateValue.getFullYear();
      let month = dateValue.getMonth() + 1; // Month is zero-based, so we add 1.
      let formattedDate = year + (month < 10 ? '0' : '') + month;
      sheet.getRange(i + 1, dateColumn).setValue(formattedDate);
    }
  }
}

function SortTableByAscendingDate(sheet, dateColumn) {
  let TABLE_INDEX = dateColumn - 1;
  let range = sheet.getDataRange();
  let data = range.getValues();

  data.sort(function (a, b) {
    let dateA = new Date(a[TABLE_INDEX].split('/').reverse().join('/'));
    let dateB = new Date(b[TABLE_INDEX].split('/').reverse().join('/'));
    return dateA - dateB;
  });

  range.setValues(data);
}
