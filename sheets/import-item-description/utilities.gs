/**
 * This file contains utility function that work with the 'MATERIAL_PACKAGING_MASTTER_DATA' folder and the 'Packaging Code' file
 */

/**
 * Gets the column index of a given file with a matching keyword (case insensitive) .
 * 
 * @param {object} value - a file's given value.
 * @param {string} text - a text string contains the name of a table column in which to find.
 * @return {object} the index object of the column.
 */
function findIndex(value, text)
{
  for (var i = 0; i < value.length; ++i)
  {
    if (value[i].indexOf(text) > -1)
    {
      return i + 1;
    }
  }
}

/**
 * Gets the previous date from the current date.
 * 
 * @return {object} the previous date
 */
function getPreviousDate()
{
  // get the current date
  var currentDate = new Date();

  // get the previous date by subtracting a day from the current date
  var previousDate = new Date(currentDate.getFullYear(), currentDate.getMonth(), currentDate.getDate() - 1);

  return previousDate;
}

/**
 * Gets the ID of the latest modified file from the 'MATERIAL_PACKAGING_MASTTER_DATA' folder.
 * 
 * @return {object} the ID of the latest modified file.
 */
function getLatestModifiedFileId()
{
  // get all files from the 'MATERIAL_PACKAGING_MASTTER_DATA' folder
  var allFile = folder.getFiles();

  // initialize a variable to store the ID of the latest modified file
  var latestModifiedFileId = '';
  var latestModifiedDate = getPreviousDate(); // initialize a variable to capture the latest modified date's previous date
  console.log('Latest modified day: ' + latestModifiedDate);

  // iterate through each file to get each file's modified time and compare this date with the defined the latest modified date's previous date variable
  while (allFile.hasNext())
  {
    var file = allFile.next();

    // check the file is in a Google Sheet
    if (file.getMimeType() == "application/vnd.google-apps.spreadsheet")
    {
      var modifiedTime = file.getLastUpdated();

      // compare the modified time with the latest modified time
      if (modifiedTime >= latestModifiedDate)
      {
        modifiedTime = latestModifiedDate
        latestModifiedFileId = file.getId() ;
      }
    }
  }
  console.log('File Id: ' + latestModifiedFileId);

  return latestModifiedFileId;
}
