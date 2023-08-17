const profileTest = SpreadsheetApp.openById("1EkNlZo_cky-q76O5tbKQeIPUyU-c4CY4UeRa37bnZaM");

// return the 'information' sheet
const informationSheet = profileTest.getSheetByName('information');
const information = informationSheet.getRange(1, 1, informationSheet.getLastRow(), informationSheet.getLastColumn()).getValues().filter(r => r[0] != '');
const informationHeader = informationSheet.getRange(1, 1, 1, informationSheet.getLastColumn()).getValue()[0];

// return value index's 'information' sheet
const newSheetNameIndex = findIndex(information, "new_sheet_name");
const profileTestURLIndex = findIndex(information, "link_result_sheet");
const newScaleIndex = findIndex(information, "scale_new");
const formNameIndex = findIndex(information, "form_name");
const formEditURLIndex = findIndex(information, "link_google_form");
const formViewURLIndex = findIndex(information, "link_google_form_view");

// return values of 'information' sheet
const newSheetName = information[newSheetNameIndex - 1][1];
const profileTestURL = information[profileTestURLIndex - 1][1];
const newScale = information[newScaleIndex - 1][1];
const formName = information[formNameIndex - 1][1];
const formEditURL = information[formEditURLIndex - 1][1].toString();
const formViewURL = information[formViewURLIndex - 1][1].toString();

const profileTestID = SpreadsheetApp.openByUrl(profileTestURL).getId();

// return the 'mo_ta' sheet
const newSheet = profileTest.getSheetByName(newSheetName);
const newSheetValue = newSheet.getRange(1, 1, newSheet.getLastRow(), newSheet.getLastColumn()).getValues();

// return the 'employee' sheet
const employeeSheet = profileTest.getSheetByName('employee');
const employee = employeeSheet.getRange(2, 1, employeeSheet.getLastRow() - 1, employeeSheet.getLastColumn()).getValues().filter(r => r[0] != '');

// return the 'mapping' sheet
const mappingSheet = profileTest.getSheetByName('mapping');
const mappingHeader = mappingSheet.getRange(1, 1, 1, mappingSheet.getLastColumn()).getValues()[0];
const mapping = mappingSheet.getRange(2, 1, mappingSheet.getLastRow() - 1,mappingSheet.getLastColumn()).getValues();

// return the 'Form Response' sheet
const formResponseSheet = profileTest.getSheetByName('Form Responses 1');
const formResponseLastRow = formResponseSheet.getLastRow();

if (formResponseLastRow > 1)
{
  var formResponse = formResponseSheet.getRange(2, 1, formResponseSheet.getLastRow() - 1, formResponseSheet.getLastColumn()).getValues();
}

const formResponseColumn = formResponseSheet.getRange(1, 1, 1, formResponseSheet.getLastColumn()).getValues()[0];

// return value index's 'Form Response' sheet
const timeIndex = findIndex(formResponseColumn, "Timestamp");
const employeeCodeIndex = findIndex(formResponseColumn, "MSNV");
const codeIndex = findIndex(formResponseColumn, "mẫu");
const noteIndex = findIndex(formResponseColumn, "off-notes");
const sensoryAttributeLength = formResponseSheet.getLastColumn() - codeIndex;

// return the 'cache' sheet
const cacheSheet = profileTest.getSheetByName('cache');
const cacheLastRow = cacheSheet.getLastRow();

if (cacheLastRow > 1)
{
  var cacheRange = cacheSheet.getRange(2, 1, cacheSheet.getLastRow() - 1, cacheSheet.getLastColumn());
  var cache = cacheRange.getValues();
}
const cacheHeader = cacheSheet.getRange(1, 1, 1, cacheSheet.getLastColumn()).getValues()[0];

// return value index's 'cache' sheet
const cacheSubjectIDIndex = findIndex(cacheHeader, "subject_id");
const cacheSampleIDIndex = findIndex(cacheHeader, "sample_id");

// return the 'data' sheet
const dataSheet = profileTest.getSheetByName('data');
const dataLastRow = dataSheet.getLastRow();

if (dataLastRow > 1)
{
  var dataRange = dataSheet.getRange(2, 1, dataSheet.getLastRow() - 1, dataSheet.getLastColumn());
  var data = dataRange.getValues();
}

// return 'Profile Test' form
const form = FormApp.openByUrl(formEditURL);
