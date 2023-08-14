/**
 * This file contains the main function that imports data of item codes and item descriptions from the latest modified file in the
 * 'MATERIAL_PACKAGING_MASTTER_DATA' folder to the destination file titled "Packaging Code"
 */

/**
 * Imports data of item codes and item descriptions from the latest modified file in the
 * 'MATERIAL_PACKAGING_MASTTER_DATA' folder to the destination file titled "Packaging Code"
 */
function importItemDescription()
{
  // get the ID of the latest modified file utilizing a getLatestModifiedFileId function
  var latestModifiedFileId = getLatestModifiedFileId();

  // open the latest modified file using its id
  var latestModifiedFile = SpreadsheetApp.openById(latestModifiedFileId);

  // return a sheet named "Sheet 1" of the latest modified file
  var sheet = latestModifiedFile.getSheetByName('Sheet 1');

  // read data of a column named "item_code" on a sheet titled "Sheet 1"
  var itemCode = sheet.getRange(2, itemCodeIndex, sheet.getLastRow() - 1, 1).getValues();

  // read data of a column named "item_description" on a sheet titled "Sheet 1"
  var itemDecription = sheet.getRange(2, itemDecriptionIndex, sheet.getLastRow() -1, 1).getValues();

  // write data of a column titled "item_code" on the a sheet named "Sheet 1" to a column with the same name on the new sheet titled "item_description"
  itemDescriptionSheet.getRange(2, itemCodeIndex, itemCode.length, 1).setValues(itemCode);

  // write data of a column titled "item_description" on the a sheet named "Sheet 1" to a column with the same name on the new sheet titled "item_description"
  itemDescriptionSheet.getRange(2, itemDecriptionIndex, itemDecription.length, 1).setValues(itemDecription);
}
