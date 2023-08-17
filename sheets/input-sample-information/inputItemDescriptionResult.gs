const testDescriptionArray = ['Cảm quan - Cảm giác châm chích trên lưỡi',	'Cảm quan - Mùi oxi hóa',	'Cảm quan - Mùi mắm tôm',	'Cảm quan - vị',	'Cảm quan - Mùi lòng đỏ trứng muối',	'Cảm quan - Cảm giác ngọt đạm',	'Cảm quan - Cảm giác chát',	'Cảm quan - Mùi gỗ',	'Cảm quan - Mùi thối',	'Cảm quan - mùi',	'Cảm quan - Vị ngọt',	'Cảm quan - Mùi nhiệt',	'Cảm quan - Vị mặn',	'Cảm quan - màu',	'Cảm quan - Mùi phô mai lên men',	'Cảm quan - Mùi lòng trắng trứng muối',	'Cảm quan - Hâu vị umami',	'Cảm quan - Vị đắng',	'Cảm quan - Mùi cá khô',	'Vị chua',	'Cảm quan - Hậu vị đắng',	'Cảm quan - Mùi muối biển',	'Cảm quan - Vị umami', 'Cảm quan - Mùi khắm', 'Tỷ lệ Nito amin/ Nitơ tổng', 'Tỷ lệ Nito amon/ Niơ tổng', 'Hàm lượng Nitơ tổng số', 'Giá trị màu a', 'Giá trị màu L', 'Giá trị màu b', 'Hàm lượng Histamine', 'Hàm lượng muối (thực phẩm)', 'Hàm lượng ion Mg2+',	'Hàm lượng ion Ca2+'];

function createMarkUpQlone()
{
  // get 'df_nhap_qlone' file
  var qloneID = qloneIDSheet.getRange(2, 1).getValue();
  var qloneFile = SpreadsheetApp.openById(qloneID);

  // get 'Sheet 1' file
  var qloneSheet = qloneFile.getSheetByName('Sheet 1');

  // get column headers
  var columnHeaderQlone = qloneSheet.getRange(1, 1, 1, qloneSheet.getLastColumn()).getValues()[0];

  // get actual_result and note2's indexes
  var actualResultIndex = findIndex(columnHeaderQlone, 'actual_result');
  var note2Index = findIndex(columnHeaderQlone, 'note2');

  // get qlone values
  var qlone = qloneSheet.getRange(1, 1, qloneSheet.getLastRow(), qloneSheet.getLastColumn()).getValues();

  // create 'markup_qlone' sheet
  var qloneMarkUpSheet = qloneFile.insertSheet('markup_qlone');
  qloneMarkUpSheet.getRange(1, 1, qlone.length, qlone[0].length).setValues(qlone);

  // check if an actual result row contains a 'None' value, this value will be replaced by an equipvalent value in note2 row
  for (var i = 0; i < qlone.length; ++i)
  {
    var actualResult = qlone[i][actualResultIndex - 1];
    var note2 = qlone[i][note2Index - 1];

    if (actualResult == 'None')
    {
      qloneMarkUpSheet.getRange(i + 1, actualResultIndex).setValue(note2);
    }
  }
}

function calculateTestDescriptionAverage()
{
  // get 'df_nhap_qlone' file
  var qloneID = qloneIDSheet.getRange(2, 1).getValue();
  var qloneFile = SpreadsheetApp.openById(qloneID);

  // get 'markup_qlone' sheet
  var qloneMarkUpSheet = qloneFile.getSheetByName('markup_qlone');

  // get 'markup_qlone' values
  var qloneMarkUp = qloneMarkUpSheet.getRange(1, 1, qloneMarkUpSheet.getLastRow(), qloneMarkUpSheet.getLastColumn()).getValues();

  // convert an array of test descriptions' values to an array of objects
  var qloneToArrayObject = convertArrayOfObject(qloneMarkUp);

  // calculate test descriptions' average values by sample name
  var qloneAverageArray = averageByGroup(qloneToArrayObject);

  // convert an array of objects to an array of test descriptions' average values
  var qloneAverageResult = qloneAverageArray.map(({sample_name, test_description, actual_result}) => [sample_name, test_description, actual_result])

  // create 'average_qlone' sheet and paste test descriptions' average values into this sheet
  var qloneAverageSheet = qloneFile.insertSheet('average_qlone');
  const columnHeaderQloneAverageArray = [['sample_name', 'test_description', 'actual_result']];
  qloneAverageSheet.getRange(1, 1, 1, columnHeaderQloneAverageArray[0].length).setValues(columnHeaderQloneAverageArray);
  qloneAverageSheet.getRange(2, 1, qloneAverageResult.length, qloneAverageResult[0].length).setValues(qloneAverageResult);

  // remain important test descriptions to analyze
  const qloneAverageRange = qloneAverageSheet.getRange(2, 1, qloneAverageSheet.getLastRow() - 1, qloneAverageSheet.getLastColumn());
  const qloneAverage = qloneAverageRange.getValues();
  deleteElementOfArray(qloneAverageSheet, qloneAverageRange, qloneAverage, testDescriptionArray);
}

function createMarkUpQloneAverage()
{
  // get 'df_nhap_qlone' file
  var qloneID = qloneIDSheet.getRange(2, 1).getValue();
  var qloneFile = SpreadsheetApp.openById(qloneID);

  // get 'markup_qlone' sheet
  var qloneAverageSheet = qloneFile.getSheetByName('average_qlone');
  var qloneAverage = qloneAverageSheet.getRange(2, 1, qloneAverageSheet.getLastRow(), qloneAverageSheet.getLastColumn()).getValues();

  // transpose an array of test descriptions' values
  var transposeQloneAverage = transposeArray(qloneAverage, 0, 1, 2);

  // create 'markup_average_qlone' sheet
  var qloneMarkUpAverageSheet = qloneFile.insertSheet('markup_average_qlone');
  qloneMarkUpAverageSheet.getRange(1, 1, transposeQloneAverage.length, transposeQloneAverage[0].length).setValues(transposeQloneAverage);
}

function inputTestDescriptionResult()
{ 
  // get 'df_nhap_qlone' file
  var qloneID = qloneIDSheet.getRange(2, 1).getValue();
  var qloneFile = SpreadsheetApp.openById(qloneID);

  // get 'markup_qlone' sheet
  var qloneMarkUpAverageSheet = qloneFile.getSheetByName('markup_average_qlone');

  qloneMarkUpAverageSheet.deleteColumn(2);

  // get 'markup_average_qlone''s column headers and test escriptions' average values
  var columnHeaderQloneAverage = qloneMarkUpAverageSheet.getRange(1, 1, 1, qloneMarkUpAverageSheet.getLastColumn()).getValues()[0];
  var qloneMarkUpAverage = qloneMarkUpAverageSheet.getRange(2, 1, qloneMarkUpAverageSheet.getLastRow() - 1, qloneMarkUpAverageSheet.getLastColumn()).getValues();

  // get the 'markup_average_qlone'sheet's column headers
  var sampleNameIndex = findIndex(columnHeaderQloneAverage, 'sample_name');
  var columnHeaderTestDescription = qloneMarkUpAverageSheet.getRange(1, sampleNameIndex + 1, 1, qloneMarkUpAverageSheet.getLastColumn()).getValues();
  var classificationIndex = findIndex(columnHeaderSampleInformation, 'phan_loai_cam_quan');

  // clear previous test descriptions' average values in the 'mau_chao' sheet before pasting new ones
  sampleInformationSheet.getRange(1, classificationIndex + 1, sampleInformationSheet.getLastRow (), sampleInformationSheet.getLastColumn()).clear({contentsOnly: true});

  // input the 'markup_average_qlone'sheet's column headers into the 'mau_chao' sheet
  sampleInformationSheet.getRange(1, classificationIndex + 1, 1, columnHeaderTestDescription[0].length).setValues(columnHeaderTestDescription);

  // input test descriptions' average values into 'mau_chao' sheet
  var testDescriptionAverage = qloneMarkUpAverageSheet.getRange(2, sampleNameIndex + 1, qloneMarkUpAverageSheet.getLastRow() - 1, qloneMarkUpAverageSheet.getLastColumn()).getValues();

  for (var x = 0; x < qloneMarkUpAverage.length; ++x)
  {
    var sampleName = qloneMarkUpAverage[x][sampleNameIndex - 1].toString().trim();

    for (var y = 0; y < sampleInformationValue.length; ++y)
    {
      var sampleNameSampleInformation = sampleInformationValue[y][sampleNameSampleInformationIndex - 1].toString().trim();
      var testDescriptionResult = [];
      
      if (sampleName == sampleNameSampleInformation)
      {
        testDescriptionResult.push([testDescriptionAverage[x]])

        sampleInformationSheet.getRange(y + 2, classificationIndex + 1, 1, columnHeaderTestDescription[0].length).setValues(testDescriptionResult[0])
      }
    }
  }
}
