var folderId = '16w1zfYkVgSTMFxaFdAquD-DOTlxEoffh'
var name = 'CHALLENGE-TEST-SCHEDULE-SPLIT.xlsx'
var folder = DriveApp.getFolderById(folderId)
var destFolderId = '1hglrMELZuuiRedRDswkOMrWVMfC8iD3F'
var calendarId = 'duclt@msc.masangroup.com'

function createGoogleSheet()
{  
  var files = folder.getFiles()
  var standardFileName = name

  while (files.hasNext())
  {
    var file = files.next()
    var fileName = file.getName()
    if (fileName == standardFileName)
    { 
      let blob = file.getBlob();

      let config = 
      {
        title: fileName.replace('.xlsx', ''),
        parents: [{id: destFolderId}],
        mimeType: MimeType.GOOGLE_SHEETS
      };

      Drive.Files.insert(config, blob);
      return
    }
  }
}

function getSpreadSheet()
{
  var spreadSheetName = name.replaceAll('.xlsx', '')
  var destFolder = DriveApp.getFolderById(destFolderId)
  var files = destFolder.getFiles()

  while (files.hasNext())
  {
    var file = files.next()
    var fileName = file.getName()
    if (fileName === spreadSheetName)
    {
      return file.getId()
    }
  }
}
