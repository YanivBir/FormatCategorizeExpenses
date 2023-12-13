let CAL_SPREADSHEET = "פירוט עסקאות וזיכויים.xlsx";
let CAL_SHEET = "פירוט עסקאות וזיכויים";

let CONVERTED_SPREADSHEET = "ConvertedFile";
let CONVERTED_SHEET = CAL_SHEET;

let MONEY_SPREADSHEET = "Money";
let MONEY_SHEET = "Expenses";
let CATEGORIES_SHEET = "Categories";

let CHARGING_DATE_DBG = null;

function FormatAndAppendNewMonth(desiredDate) {
  if(!desiredDate)
    desiredDate = CHARGING_DATE_DBG;
  ConvertExcelToGoogleSheets(CAL_SPREADSHEET, CONVERTED_SPREADSHEET);
  let convertedSheet = OpenExternalSpreadsheet(CONVERTED_SPREADSHEET, CONVERTED_SHEET);
  let categories = LoadCategories();
  FormatCalReport(convertedSheet, categories, desiredDate);

  let moneySheet = OpenExternalSpreadsheet(MONEY_SPREADSHEET, MONEY_SHEET);
  AppendDataToSheet(convertedSheet, moneySheet)
  
  DeleteFileIfExists(CONVERTED_SPREADSHEET);
}

function LoadCategories() {
  let categorySheet = OpenExternalSpreadsheet(MONEY_SPREADSHEET, CATEGORIES_SHEET);
  let CATEGORY_STARTING_ROW = 3;

  let dictionary = {};
  let data = categorySheet.getDataRange().getValues();
  let backgrounds = categorySheet.getDataRange().getBackgrounds();

  for (var i = CATEGORY_STARTING_ROW-1; i < data.length; i++) {
    let key = data[i][0];
    let category = { 
      name: data[i][1],
      subName: data[i][2],
      background: backgrounds[i][1]
    };
    dictionary[key] = category;
  }

  return dictionary;
}
