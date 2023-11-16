let CAL_SPREADSHEET = "פירוט עסקאות וזיכויים.xlsx";
let CAL_SHEET = "פירוט עסקאות וזיכויים";

let CONVERTED_SPREADSHEET = "ConvertedFile";
let CONVERTED_SHEET = CAL_SHEET;

let MONEY_SPREADSHEET = "Money";
let MONEY_SHEET = "Expenses";

function Main() {
}

function FormatAndAppendNewMonth() {
  ConvertExcelToGoogleSheets(CAL_SPREADSHEET, CONVERTED_SPREADSHEET);
  let convertedSheet = OpenExternalSpreadsheet(CONVERTED_SPREADSHEET, CONVERTED_SHEET);
  FormatCalReport(convertedSheet);

  let moneySheet = OpenExternalSpreadsheet(MONEY_SPREADSHEET, MONEY_SHEET);
  AppendDataToSheet(convertedSheet, moneySheet)

  DeleteFileIfExists(CONVERTED_SPREADSHEET);
}
