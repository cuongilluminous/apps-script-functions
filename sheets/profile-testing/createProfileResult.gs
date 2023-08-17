function createProfileCache()
{
  cacheRange.clear({contentsOnly: true});

  for (var i = 0; i < formResponse.length; ++i)
  {
    var time = formResponse[i][timeIndex - 1];
    var formatTime = Utilities.formatDate(new Date(time), "GMT+7", "yyyy-MM-dd");
    var employeeCode = formResponse[i][employeeCodeIndex - 1];
    var code = formResponse[i][codeIndex - 1];
    var note = formResponse[i][noteIndex - 1];

    var informationTest = [[formatTime, employeeCode, code, note]];
    var sensoryAttribute = formResponseSheet.getRange(i + 2, codeIndex + 1, 1, sensoryAttributeLength).getValues();

    cacheSheet.getRange(i + 2, 1, informationTest.length, informationTest[0].length).setValues(informationTest);
    cacheSheet.getRange(i + 2, informationTest[0].length + 1, sensoryAttribute.length, sensoryAttribute[0].length).setValues(sensoryAttribute);
  }
}

function createProfileData()
{
  var cacheDuplication = cacheSheet.getRange(2, cacheSubjectIDIndex, cacheSheet.getLastRow() - 1, cacheSampleIDIndex - cacheSubjectIDIndex + 1).getValues();
  var uniqueCache = findUniqueValue(cacheDuplication);
  var profileData = removeDuplication(uniqueCache, cacheDuplication, cache);

  if (profileData.join() == "")
  {
    dataSheet.getRange(2, 1, cache.length, cache[0].length).setValues(cache)
  }
  else
  {
    dataSheet.getRange(2, 1, profileData.length, profileData[0].length).setValues(profileData);
  }
}
