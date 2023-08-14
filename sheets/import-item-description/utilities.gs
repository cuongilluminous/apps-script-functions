/**
 * This file contains utility functions that work with the 'MATERIAL_PACKAGING_MASTTER_DATA' folder and the 'Packaging Code' file
 */

/**
 * Gets the column index of a given file with a matching keyword (case insensitive).
 * 
 * @param {object} value - data of a given file.
 * @param {string} text - a text string contains the name of a table column in which to find.
 * @returns {object} the index object of the column.
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
 * Gets the previous date of the current date.
 * 
 * @returns {object} the previous date
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
 * Gets the ID of the latest modified file in the 'MATERIAL_PACKAGING_MASTTER_DATA' folder.
 * 
 * @returns {object} the ID of the latest modified file.
 */
function getLatestModifiedFileId()
{
  // get all files in the 'MATERIAL_PACKAGING_MASTTER_DATA' folder
  var allFile = folder.getFiles();

  // initialize a variable to store the ID of the latest modified file
  var latestModifiedFileId = '';
  var latestModifiedDate = getPreviousDate(); // initialize a variable to capture the latest modified date's previous date
  console.log('Latest modified date: ' + latestModifiedDate);

  // iterate through each file to get each file's a modified date and compare this date with the defined the latest modified date's previous date
  while (allFile.hasNext())
  {
    var file = allFile.next();

    // check a file is in a Google Sheet
    if (file.getMimeType() == "application/vnd.google-apps.spreadsheet")
    {
      var modifiedDate = file.getLastUpdated();

      // compare the modified date with the latest modified date
      if (modifiedDate >= latestModifiedDate)
      {
        modifiedDate = latestModifiedDate
        latestModifiedFileId = file.getId() ;
      }
    }
  }
  console.log('File Id: ' + latestModifiedFileId);

  return latestModifiedFileId;
}
