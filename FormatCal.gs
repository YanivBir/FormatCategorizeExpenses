let SRC_CHARGING_DATE_COLUMN = 5;
let DST_CHARGING_COLUMN = 1;
let BUYING_DATE_COLUMN = 2;
let BUSSNIES_COLUMN = 3;
let CATEGORY_COLUMN = 6;
let SUB_CATEGORY_COLUMN = 7;

function FormatCalReport(sheet, categories, desiredChargingDate) {
  if(!desiredChargingDate)
    throw new Error("Desired charging date is undifened.");
  DeleteUnusedRowsAndColumns(sheet);
  MoveColumn(sheet, SRC_CHARGING_DATE_COLUMN, DST_CHARGING_COLUMN);
  ReverseRows(sheet);
  FormatTable(sheet, categories, desiredChargingDate);
  SetTableStyle(sheet)
  Logger.log("FormatCalReport() is finished.");
}

function DeleteUnusedRowsAndColumns(sheet) {
  sheet.deleteRows(1, 2);
  sheet.deleteRow(sheet.getLastRow());
  sheet.deleteColumns(sheet.getLastColumn() - 2, 3);
}

function SetTableStyle(sheet) {
  let table = sheet.getDataRange();
  
  table.setFontFamily("Ariel");
  table.setFontSize(10);
  table.setBorder(false, false, false, false, false, false);
}

function FormatTable(sheet, categories, desiredChargingDate) {
  let data = sheet.getDataRange().getValues();
  for (let i = sheet.getLastRow() - 1; i >= 0; i--) {
    let charging = data[i][DST_CHARGING_COLUMN - 1];
    let formattedCharging;
    if (charging != "")
      formattedCharging = FormatChargingData(charging);
    if((charging == "") || (formattedCharging != desiredChargingDate))
    {
      sheet.deleteRow(i + 1);
      continue;
    }
    sheet.getRange(i + 1, DST_CHARGING_COLUMN).setValue(formattedCharging);
    SetCategory(sheet, data, categories, i);
  }
}

function SetCategory(sheet, data, categories, i) {
  let bussniesName = data[i][BUSSNIES_COLUMN - 1];
  let category = categories[bussniesName];
  if(!category)
    return;
  sheet.getRange(i + 1, CATEGORY_COLUMN).setValue(category.name);
  sheet.getRange(i + 1, CATEGORY_COLUMN).setBackground(category.background);
  if(category.subName) {
      sheet.getRange(i + 1, SUB_CATEGORY_COLUMN).setValue(category.subName);
      sheet.getRange(i + 1, SUB_CATEGORY_COLUMN).setBackground(category.background);
  }
}

function FormatChargingData(dateValue) {
 if (dateValue instanceof Date) {
      let year = dateValue.getFullYear();
      let month = dateValue.getMonth() + 1; // Month is zero-based, so we add 1.
      let formattedDate = year + (month < 10 ? '0' : '') + month;
      return formattedDate;
    }
    throw new Error("An error occurred while processing charging date, got: \"" +  dateValue + "\" as input.");
}

function ReverseRows(sheet) {
  let lastRow = sheet.getLastRow();
  let lastColumn = sheet.getLastColumn();
  let range = sheet.getRange(1, 1, lastRow, lastColumn);
  let data = range.getValues();

  data.reverse();
  range.clearContent();
  sheet.getRange(1, 1, data.length, data[0].length).setValues(data);
}
