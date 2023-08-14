/**
 * This file contains the main function that convert a datetime object in the given table to the desired datetime format
 */
function formatDateUsingRegex()
{
  // initialize a regular expression to match ' ' or 'Updating'
  var regex = /^\d{2}\/\d{2}\/\d{2}$/;
  
  for (var i = 0; i < packagingCode.length; ++i)
  {
    var validDateMBO = packagingCode[i][validDateMBOIndex - 1];
    
    if (regex.test(validDateMBO)) {
      // split a given date into a array with three elements; including a year, a month and a date, using split() method based on a seperator ("/")
      var splittedvalidDateMBO = validDateMBO.split('/');

      // return a year, a month and a date
      var year = '20' + splittedvalidDateMBO[2];
      var month = splittedvalidDateMBO[1];
      var day = splittedvalidDateMBO[0];

      // format a given date in the the format year-month-date 
      var formattedvalidDateMBO = Utilities.formatDate(new Date(year, month - 1, day), 'GMT+7', 'yyyy-MM-dd');

      // write a converted date time value to each column cell if it does not contain ' ' or 'Updating'
      packagingCodeSheet.getRange(i + 2, validDateMBOIndex).setValue(formattedvalidDateMBO);
    } else {
      // write an unconverted date time value to each column cell if it contains ' ' or 'Updating'
      packagingCodeSheet.getRange(i + 2, validDateMBOIndex).setValue(validDateMBO);
    }
  }
}
