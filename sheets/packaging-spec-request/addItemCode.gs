// open a sheet titled "item_code" by its name in the spec file
const itemCodeSpecSheet = specFile.getSheetByName('item_code');
const itemCodeSpecHeader = itemCodeSpecSheet.getRange(1, 1, 1, itemCodeSpecSheet.getLastColumn()).getValues()[0]; // get values for the specified range on the sheet named "item_code"
const itemCodeSpecData = itemCodeSpecSheet.getRange(2, 1, itemCodeSpecSheet.getLastRow() - 1, itemCodeSpecSheet.getLastColumn()).getValues(); // get values for the specified range on the sheet named "item_code"

function addItemCode() {
  for (var i = 0; i < itemCodeSpecData.length; ++i)
  {
    var submitTime = itemCodeSpecData[i][submitTimeIndex - 1];
    var formattedSubmitTime = Utilities.formatDate(new Date(submitTime), 'GMT+7', 'yyyy-MM-dd');
    var mailingStatus = itemCodeSpecData[i][findColumnIndex(itemCodeSpecHeader, 'mailing_status') - 1];

    if (submitTime != '')
    {
      if (mailingStatus != 'Sent')
      {
        var requesterEmail = itemCodeSpecData[i][requesterEmailIndex  - 1];
        console.log('Requester email: ' + requesterEmail);
        var picEmail = itemCodeSpecData[i][picEmailIndex  - 1];
        var extractedPICEmail = extractValueFromString(picEmail, /([a-zA-Z0-9._-]+@[a-zA-Z0-9._-]+\.[a-zA-Z0-9_-]+)/, 'No email address found: ');
        var standardName = itemCodeSpecData[i][findColumnIndex(itemCodeSpecHeader, 'ma_tieu_chuan') - 1];
        console.log('The standard name is: ' + standardName)
        var itemCode = itemCodeSpecData[i][findColumnIndex(itemCodeSpecHeader, 'item_code') - 1];

        var version = createVersion(standardName, packagingCode, packagingCodeHeader);

        var condition = 0;

        for (var j = 0; j < packagingCode.length; ++j)
        {
          if (packagingCode[j][standardCodeIndex - 1].toString().trim() == standardName.toString().trim() &
              packagingCode[j][findColumnIndex(packagingCodeHeader, 'phien_ban') - 1] == version[1])
          {
            condition = 1;

            var itemCodeList = itemCode.split(',');
            console.log('The item code list is: ' + itemCodeList);
            addPackagingCode(standardName, itemCodeList);

            var subject = '[ADD-ITEM-CODE] Yêu Cầu Thêm Item Code ' + itemCode + '_' + formattedSubmitTime;

            var body = "<div style=\"text-align:left;display: inline-block\">";
            body += "<p>" + "Dear anh/chị RD Packaging," + "</p>"; 
            body += "<p>" + "Team SPEC đã nhận được yêu cầu thêm item code " + itemCode + " vào thời gian " + formattedSubmitTime + "</p>";
            body += "<p>" + "Anh/chị vui lòng truy cập vào đường dẫn đính kèm tại đây để kiểm tra thông tin: " + "<a class=\"standard-template-folder-link\", href="+"https://docs.google.com/spreadsheets/d/"+PACKAGING_CODE_ID+">" + "Truy cập tại đây" + "</a>" + "</p>";

            var tableHeader = ["Thông tin yêu cầu", "Mã tiêu chuẩn", "Item code"];

            var tableData = ["Thêm item code", standardName, itemCode];

            sendSPECEmail(subject, body, tableHeader, tableData, requesterEmail, extractedPICEmail);

            break;
          }
        }
          
        if (condition == 0)
        {
          var subject = '[ADD-ITEM-CODE] Không Tìm Thấy Item Code ' + itemCode + '_' + formattedSubmitTime;

          var body = "<div style=\"text-align:left;display: inline-block\">";
          body += "<p>" + "Dear anh/chị RD Packaging," + "</p>"; 
          body += "<p>" + "Team SPEC đã nhận được yêu cầu thêm item code, nhưng chúng tôi không tìm thấy item code " + itemCode + " phù hợp yêu cầu vào thời gian " + formattedSubmitTime + "</p>";
          body += "<p>" + "Anh/chị vui lòng truy cập vào đường dẫn đính kèm tại đây để kiểm tra thông tin: " + "<a class=\"standard-template-folder-link\", href="+"https://docs.google.com/spreadsheets/d/"+PACKAGING_CODE_ID+">" + "Truy cập tại đây" + "</a>" + "</p>";

          var tableHeader = ["Thông tin yêu cầu", "Mã tiêu chuẩn", "Item code"];

          var tableData = ["Thêm item code", standardName, itemCode];

          sendSPECEmail(subject, body, tableHeader, tableData, requesterEmail, extractedPICEmail);

          break;
        }

        itemCodeSpecSheet.getRange(i + 2, findColumnIndex(itemCodeSpecHeader, 'mailing_status')).setValue('Sent');
      }
    }
  }
}
