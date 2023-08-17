// get 'sample_id_bon' sheet's column indexes
const sensorySampleIDIndex = findIndex(columnHeaderSensorySampleID, 'sample_id');
const sensoryCodeIndex = findIndex(columnHeaderSensorySampleID, 'code');
const sensoryStatusIndex = findIndex(columnHeaderSensorySampleID, 'status');

function createSensoryCodeInput()
{
  for (var j = 0; j < sensorySampleIDValue.length; ++j)
  {
    var sensorySampleID = sensorySampleIDValue[j][sensorySampleIDIndex - 1];
    var sensoryStatus = sensorySampleIDValue[j][sensoryStatusIndex - 1];
    
    if (sensorySampleID != '' & sensoryStatus != 'Done')
    {
      sensorySampleIDSheet.getRange(j + 2, sensoryCodeIndex).setValue(sensorySampleID);
      sensorySampleIDSheet.getRange(j + 2, sensoryStatusIndex).setValue('Done');
    }
  }
}
