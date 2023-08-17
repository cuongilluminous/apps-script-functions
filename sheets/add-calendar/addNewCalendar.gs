function runCreateNewEvent()
{
  try
  {
    createGoogleSheet()
    Utilities.sleep(2000)
    var spreadSheetId = getSpreadSheet()
    Logger.log(spreadSheetId)
    var ss = SpreadsheetApp.openById(spreadSheetId)
    var sheet = ss.getSheetByName('calendar')
    var data = sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).getValues().filter(r => r[0] != '')
    Logger.log(data);
    createNewEvent(data);
  }

  catch (f)
  {
    Logger.log(f.toString())
  }
  DriveApp.getFileById(spreadSheetId).setTrashed(true)
}

function createNewEvent(data) 
{
  for (var i = 0; i < data.length; ++i)
  {
    var projectId = data[i][0]
    var n = data[i][1]
    var	productName	= data[i][2] 
    var startDate	= data[i][3]
    var milestone	= data[i][4]
    var type = data[i][5]
    var	color =	parseFloat(data[i][6])

    if (type.trim() == 'new')
    {
      var dayList = milestone.toString().split(',')
      Logger.log(dayList)
      for (var j = 0; j < dayList.length; ++j)
      {
        var mailTitle = projectId + ' - ' + n + ' - ' + productName + ' - D' + parseInt(dayList[j])
        var day = startDate.getDate() + parseFloat(dayList[j])
        var month = startDate.getMonth()
        var year = startDate.getFullYear();
        var date = new Date(year, month, day)
        var cal = CalendarApp.getCalendarById(calendarId);
        var event = cal.createAllDayEvent(mailTitle, 
                        date,
                        ).setVisibility(CalendarApp.Visibility.PUBLIC);                        
                        // .addGuest(calendarId)
                                
        event.setColor(color)
      }
    }
  }
}
