function runDeleteEvent()
{
  try
  {
    createGoogleSheet()
    Utilities.sleep(2000)
    var spreadSheetId = getSpreadSheet()
    Logger.log(spreadSheetId)
    var ss = SpreadsheetApp.openById(spreadSheetId)
    var sheet = ss.getSheetByName('calendar')
    var data = sheet.getRange(2, 1, sheet.getLastRow(), sheet.getLastColumn()).getValues().filter(r => r[0] != '')
    Logger.log(data)
    deleteEvent(data)
  }

  catch (f)
  {
    Logger.log(f.toString())
  }
  DriveApp.getFileById(spreadSheetId).setTrashed(true)
}

function deleteEvent(data) 
{
  var fromDate = new Date("2020-01-01"); 
  var toDate = new Date("2030-12-31");

  for (var i = 0; i < data.length; ++i)
  {
    var projectId = data[i][0]
    var n = data[i][1]
    var	productName	= data[i][2] 
    var milestone	= data[i][4].toString()
    var type = data[i][5]
    
    if (type.trim() == 'delete')
    {
      var dayList = milestone.split(',')
      for (var j = 0; j < dayList.length; ++j)
      {
        var mailTitle = projectId + ' - ' + n + ' - ' + productName + ' - D' + parseInt(dayList[j])
        var cal = CalendarApp.getCalendarById(calendarId);
        var events = cal.getEvents(fromDate, toDate);

        for(var k = 0; k < events.length; k++)
        {
          var event = events[k];
          var eventName = event.getTitle()

          if (eventName === mailTitle)
          { 
            Logger.log("Yeah: " + eventName)
            event.deleteEvent();
            break
          }
        }
      }
    }
  }
}
