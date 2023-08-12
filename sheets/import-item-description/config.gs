/**
 * This file contains application constants
 */
const FOLDER_ID = '1Fo91Fnm--3FddcwrGj9qNannwyoPbu9y'; // the ID of the 'MATERIAL_PACKAGING_MASTTER_DATA' folder
const folder = DriveApp.getFolderById(FOLDER_ID); // open the 'MATERIAL_PACKAGING_MASTTER_DATA' folder by utilizing the ID of this folder

const ID = '1WrkHXf2QqRzcU68JqB5-oZPcemQ-RvbGwwtAZGgJ5CY'; // the ID of the 'Packaging Code' file
const file = SpreadsheetApp.openById(ID); // open the 'Packaging Code' file by utilizing the ID of this file

// Open the 'item_description' sheet in the 'Packaging Code' file
const itemDescriptionSheet = file.getSheetByName('item_description');
const itemDescriptionHeader = itemDescriptionSheet.getRange(1, 1, 1, itemDescriptionSheet.getLastColumn()).getValues()[0];

const itemCodeIndex = findIndex(itemDescriptionHeader, 'item_code'); // initialize a variable to contain the index of the heading column called 'item_code'
const itemDecriptionIndex = findIndex(itemDescriptionHeader, 'item_description'); // initialize a variable to store the index of the heading column called 'item_description'

// Open the 'packaging_code' sheet in the 'Packaging Code' file
const packagingCodeSheet = file.getSheetByName('packaging_code');
const packagingCodeHeader = packagingCodeSheet.getRange(1, 1, 1, packagingCodeSheet.getLastColumn()).getValues()[0];
const packagingCode = packagingCodeSheet.getRange(2, 1, packagingCodeSheet.getLastRow() - 1, packagingCodeSheet.getLastColumn()).getValues();

const validDateIndex = findIndex(packagingCodeHeader, 'ngay_hieu_luc'); // initialize a variable to contain the index of the heading column called 'ngay_hieu_luc'
const validDateMBOIndex = findIndex(packagingCodeHeader, 'ngay_hieu_luc_mpo'); // initialize a variable to contain the index of the heading column called 'ngay_hieu_luc_mpo'
