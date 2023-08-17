// get 'INPUT-FISH-SAUCE-FILE' folder
const folder = DriveApp.getFolderById('17o4PRVzrhb213vxI66y7-KZ6G-WwtUxN');

// get '[APP SHEET] THÔNG TIN MẪU' file
const sampleInformationFile = SpreadsheetApp.openById('1dXN0wGbtoewGcjDxZIHNBTbIMVTduOxfjtrUnh_sj7o');

// get '3-digit-code' file
const codeFile = SpreadsheetApp.openById('1BT9ZS8NVY1al9B8HkSy-8aNCDIJvLaZH1DLP3-0Z4Sw');

// get 'mau_chao' sheet
const sampleInformationSheet = sampleInformationFile.getSheetByName('mau_chao');
const columnHeaderSampleInformation = sampleInformationSheet.getRange(1, 1, 1, sampleInformationSheet.getLastColumn()).getValues()[0];
const sampleInformationValue = sampleInformationSheet.getRange(2, 1, sampleInformationSheet.getLastRow() - 1, sampleInformationSheet.getLastColumn()).getValues();

// get 'mau_bon' sheet
const sensorySampleInformationSheet = sampleInformationFile.getSheetByName('mau_bon');
const sensorySampleInformation = sensorySampleInformationSheet.getRange(2, 1, sensorySampleInformationSheet.getLastRow() - 1, sensorySampleInformationSheet.getLastColumn()).getValues();
const columnHeaderSensorySampleInformation = sensorySampleInformationSheet.getRange(1, 1, 1, sensorySampleInformationSheet.getLastColumn()).getValues()[0];

// get 'mau_bon' sheet's column indexes
const sensorySampleInformationSubmitTimeIndex = findIndex(columnHeaderSensorySampleInformation, 'submit_form');
const sensorySampleInformationLotNumberIndex = findIndex(columnHeaderSensorySampleInformation, 'lot_number');

//get sample_name column's index
const sampleNameSampleInformationIndex = findIndex(columnHeaderSampleInformation, 'sample_name');

// get 'supplier' sheet
const supplierSheet = sampleInformationFile.getSheetByName('supplier');
const supplierValue = supplierSheet.getRange(2, 1, supplierSheet.getLastRow() - 1, supplierSheet.getLastColumn()).getValues();

// get ' sample_id' sheet
const sampleIDSheet = sampleInformationFile.getSheetByName('sample_id');
const columnHeaderSampleID = sampleIDSheet.getRange(1, 1, 1, sampleIDSheet.getLastColumn()).getValues()[0];
const sampleIDValue = sampleIDSheet.getRange(2, 1, sampleIDSheet.getLastRow() - 1, sampleIDSheet.getLastColumn()).getValues();

// get 'sample_id_nhap' sheet
const sampleIDInputSheet = sampleInformationFile.getSheetByName('sample_id_nhap');
const columnHeaderSampleIDInput = sampleIDInputSheet.getRange(1, 1, 1, sampleIDInputSheet.getLastColumn()).getValues()[0];
const sampleIDInputLastRow = sampleIDInputSheet.getLastRow();

if (sampleIDInputLastRow > 1)
{
  var sampleIDInputValue = sampleIDInputSheet.getRange(2, 1, sampleIDInputSheet.getLastRow() - 1, sampleIDInputSheet.getLastColumn()).getValues();
}

// get 'sample_id_bon' sheet
const sensorySampleIDSheet = sampleInformationFile.getSheetByName('sample_id_bon');
const columnHeaderSensorySampleID = sensorySampleIDSheet.getRange(1, 1, 1, sensorySampleIDSheet.getLastColumn()).getValues()[0];
const sensorySampleIDLastRow = sensorySampleIDSheet.getLastRow();

if (sensorySampleIDLastRow > 1)
{
  var sensorySampleIDValue = sensorySampleIDSheet.getRange(2, 1, sensorySampleIDSheet.getLastRow() - 1, sensorySampleIDSheet.getLastColumn()).getValues();
}

// get 'qlone_id' sheet
const qloneIDSheet = sampleInformationFile.getSheetByName('qlone_id');

// get 'Sheet1' sheet
const codeSheet = codeFile.getSheetByName('Sheet1');
const codeValue = codeSheet.getRange(2, 1, codeSheet.getLastRow() - 1, codeSheet.getLastColumn()).getValues();
