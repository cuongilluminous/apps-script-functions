// get 'sample_id' sheet's column indexes
const sampleNameSampleIDIndex = findIndex(columnHeaderSampleID, 'sample_name');
const sampleIDIndex = findIndex(columnHeaderSampleID, 'sample_id');
const codeIndex = findIndex(columnHeaderSampleID, 'code');
const statusIndex = findIndex(columnHeaderSampleID, 'status');

function createCode()
{
  for (var j = 0; j < sampleIDValue.length; ++j)
  {
    var sampleID = sampleIDValue[j][sampleIDIndex - 1];
    var status = sampleIDValue[j][statusIndex - 1];
    
    if (sampleID != '' & status != 'Done')
    {
      sampleIDSheet.getRange(j + 2, codeIndex).setValue(sampleID);
      sampleIDSheet.getRange(j + 2, statusIndex).setValue('Done');
    }
  }
}
