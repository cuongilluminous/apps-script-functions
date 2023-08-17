function getQloneID()
{
  var allFile = folder.getFiles();

  if (!allFile.hasNext())
  {
    var id = convertExcel('df_nhap_qlone.xlsx');
    DriveApp.getFileById(id).moveTo(folder);
    qloneIDSheet.getRange(2, 1).setValue(id);
  }
  else
  {
    while (allFile.hasNext())
    {
      var file = allFile.next();
      var fileName = file.getName();

      if (fileName == 'df_nhap_qlone')
      {
        file.setTrashed(true);
        var id = convertExcel('df_nhap_qlone.xlsx');
        DriveApp.getFileById(id).moveTo(folder);
        qloneIDSheet.getRange(2, 1).setValue(id);
      }
      else
      {
        var id = convertExcel('df_nhap_qlone.xlsx');
        DriveApp.getFileById(id).moveTo(folder);
        qloneIDSheet.getRange(2, 1).setValue(id);
      }
    }
  }  
}
