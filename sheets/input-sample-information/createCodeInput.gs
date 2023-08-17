// get 'sample_id_nhap' sheet's column indexes
const sampleIDInputIndex = findIndex(columnHeaderSampleIDInput, 'sample_id');
const codeInputIndex = findIndex(columnHeaderSampleIDInput, 'code');
const statusInputIndex = findIndex(columnHeaderSampleIDInput, 'status');

function createCodeInput()
{
  for (var j = 0; j < sampleIDInputValue.length; ++j)
  {
    var sampleIDInput = sampleIDInputValue[j][sampleIDInputIndex - 1];
    var statusInput = sampleIDInputValue[j][statusInputIndex - 1];
    
    if (sampleIDInput != '' & statusInput != 'Done')
    {
      sampleIDInputSheet.getRange(j + 2, codeInputIndex).setValue(sampleIDInput);
      sampleIDInputSheet.getRange(j + 2, statusInputIndex).setValue('Done');
    }
  }
}
