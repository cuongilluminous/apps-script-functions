function createOrderForm()
{
  if (responseLastRow < 1)
  {
    return;
  }
  else
  {
    responseSheet.getRange(1, statusIndex).setValue('Tình trạng gửi'); // input 'Tình trạng gửi' statement into the 4th-positioned column
    
    for (var i = 0; i < response.length; ++i)
    {
      console.log(i);

      var status = response[i][statusIndex - 1];

      if (status != 'Sent')
      {
        var submitTime = response[i][submitTimeIndex - 1];
        var dateTime = Utilities.formatDate(new Date(submitTime), "GMT+7", 'dd-MM-yyyy HH:m:s'); // format the submitting time into an ISO time format
        var link = response[i][linkIndex - 1];
        var email = response[i][emailOrderIndex - 1];
        console.log(email);
        var emailLineManager = findEmail(email);
        console.log(emailLineManager);

        checkFile(i, responseSheet, dateTime, link, email, formURL); // check an order file submitted by an order form whether is in the correct format or not. In case of the wrong format, a notification email will be sent to alert a requester

        var file = SpreadsheetApp.openByUrl(link);

        if (file.getSheetByName('order_template') == null) continue;

        var orderData = file.getSheetByName('order_template').getRange(2, 1, file.getLastRow() - 1, file.getLastColumn()).getValues().filter(r => r != '');

        var orderList = [];

        for (var j = 0; j < orderData.length; ++j)
        {
          var itemName = orderData[j][itemNameIndex - 1];
          var quantity = orderData[j][quantityIndex - 1];
          var unit = orderData[j][unitIndex - 1];
          var idScheme = removeSpecialCharacterFromString(orderData[j][idSchemeIndex - 1]);
          var category = removeSpecialCharacterFromString(orderData[j][categoryIndex - 1]);
          var requestDate = orderData[j][requestDateIndex - 1];
          var note = orderData[j][noteIndex - 1];

          orderList.push([link, dateTime, itemName, quantity, unit, idScheme, category, requestDate, note, email, emailLineManager]);
        }

        var newOrderFileID = createOrderFile(dateTime + " - " + email);
        var newOrderFile = SpreadsheetApp.openById(newOrderFileID);
        var orderSheet = newOrderFile.getSheets()[0].setName('order_information');

        var columnHeader = ['Link','Thời gian', 'Tên mặt hàng', 'Số lượng', 'Đơn vị tính', 'ID Scheme', 'Ngành hàng', 'Ngày cần', 'Chú ý', 'Email', 'Email (Line Manager)'];
        orderSheet.getRange(1, 1, 1, orderList[0].length).setValues([columnHeader]);

        orderSheet.getRange(orderSheet.getLastRow() + 1, 1, orderList.length, orderList[0].length).setValues(orderList);

        var approvalFormURL = createApprovalForm(dateTime, email);

        // send an order email when order information is valid
        var htmlTable = "<table cellspacing=0 cellpadding=\"5px\" style=\"border-collapse:collapse\", border=\"1px\">";
        htmlTable += "<tr>" + "<th style=\"background-color: yellow\", align=\"center\">" + "Thời gian" + "</th>";
        htmlTable += "<th style=\"background-color: yellow\", align=\"center\">" + "Email" + "</th>";
        htmlTable += "<th style=\"background-color: yellow\", align=\"center\">" + "Email (Line Manager)" + "</th>";
        htmlTable += "<th style=\"background-color: yellow\", align=\"center\">" + "Tên mặt hàng" + "</th>";
        htmlTable += "<th style=\"background-color: yellow\", align=\"center\">" + "Số lượng" + "</th>";
        htmlTable += "<th style=\"background-color: yellow\", align=\"center\">" + "Đơn vị tính" + "</th>";
        htmlTable += "<th style=\"background-color: yellow\", align=\"center\">" + "ID Scheme" + "</th>";
        htmlTable += "<th style=\"background-color: yellow\", align=\"center\">" + "Ngành hàng" + "</th>";
        htmlTable += "<th style=\"background-color: yellow\", align=\"center\">" + "Ngày cần" + "</th>";
        htmlTable += "<th style=\"background-color: yellow\", align=\"center\">" + "Ghi chú" + "</th>" + "</tr>";

        for (var k = 0; k < orderList.length; ++k)
        {
          htmlTable += "<tr>" + "<td align=\"center\">" + orderList[k][1] + "</td>";
          htmlTable += "<td align=\"center\">" + orderList[k][9] + "</td>";
          htmlTable += "<td align=\"center\">" + orderList[k][10] + "</td>";
          htmlTable += "<td align=\"center\">" + orderList[k][2] + "</td>";
          htmlTable += "<td align=\"center\">" + orderList[k][3] + "</td>";
          htmlTable += "<td align=\"center\">" + orderList[k][4] + "</td>";
          htmlTable += "<td align=\"center\">" + orderList[k][5] + "</td>";
          htmlTable += "<td align=\"center\">" + orderList[k][6] + "</td>";
          htmlTable += "<td align=\"center\">" + orderList[k][7] + "</td>";
          htmlTable += "<td align=\"center\">" + orderList[k][8] + "</td>";
          htmlTable += "</tr>";
        }

        htmlTable += "</table>";
        var subject = 'ĐỀ NGHỊ MUA HÀNG ' + " - " + dateTime.toString();
        var body = "<div style=\"text-align:left;display: inline-block\">";
        body += "<p>" + "Dear anh/ chị " + emailLineManager + "," + "</p>"; 
        body += "<p>" + "Có một đề nghị mua hàng từ " + email + " đang chờ anh/ chị phê duyệt." + "</p>";
        body += "<p>" + "Nếu có thay đổi thông tin, anh/ chị vui lòng vào đây để thay đổi: " +  "<a class=\"gmail-in-cell-link\", href="+link+">" + "Đơn mua hàng" + "</a>" + "</p>";
        body += "<p>" + "Anh/ chị phê duyệt tại đây: " + "<a class=\"gmail-in-cell-link\", href="+approvalFormURL+">" + "Form phê duyệt" + "</a>" + "</p>";
        body += "<p>" + "Thông tin của các sản phẩm cần mua ở bảng phía dưới." + "</p>";
        var ending = "<div style=\"text-align:left;display: inline-block\">";
        ending += "<p>" + "Best regards," + "</p>";
        ending += "<p>" + "Admin RD" + "</p>";

        var emailCC = email;

        var emailBCC = 'cuongtq4@msc.masangroup.com, thond@msc.masangroup.com, anhntp@msc.masangroup.com';

        MailApp.sendEmail(emailLineManager, subject, body, {htmlBody: body + htmlTable + ending, cc: emailCC, bcc: emailBCC});

        responseSheet.getRange(i + 2, statusIndex).setValue('Sent');
      }
      else if (status == 'Yes')
      {
        break;
      }
    }
  }
}

function sendApprovalMail()
{
  var sheetList = approvalFile.getSheets();
  
  for (var i = 0; i < sheetList.length; ++i)
  {
    console.log(i);

    var approvalSheet = sheetList[i];
    var approvalData = approvalSheet.getRange(2, 1, 1, approvalSheet.getLastColumn()).getValues();

    var approvalStatus = approvalData[0][statusIndex - 2];
    var mailingStatus = approvalData[0][statusIndex - 1];

    var newOrderFileID = findOrderFile(approvalSheet.getName());

    console.log(checkID(newOrderFileID));

    if (checkID(newOrderFileID) == true)
    {
      var newOrderFile = SpreadsheetApp.openById(newOrderFileID);

      var orderSheet = newOrderFile.getSheetByName('order_information');
      var orderInformation = orderSheet.getRange(2, 1, orderSheet.getLastRow() - 1, orderSheet.getLastColumn()).getValues().filter(r => r != '');

      var submitTime = orderInformation[0][submitTimeIndex];
      var link = orderInformation[0][linkIndex - 2];
      var email = orderInformation[0][noteIndex + submitTimeIndex + 1];
      console.log(email);
      var emailLineManager = orderInformation[0][noteIndex + submitTimeIndex + 2];
      console.log(emailLineManager);
      var orderList = [];

      for (var j = 0; j < orderInformation.length; ++j)
      {
        var itemName = orderInformation[j][itemNameIndex + submitTimeIndex];
        var quantity = orderInformation[j][quantityIndex + submitTimeIndex];
        var unit = orderInformation[j][unitIndex + submitTimeIndex];
        var idScheme = orderInformation[j][idSchemeIndex + submitTimeIndex]
        var category = orderInformation[j][categoryIndex + submitTimeIndex];
        var requestDate = orderInformation[j][requestDateIndex + submitTimeIndex];
        var note = orderInformation[j][noteIndex + submitTimeIndex];

        orderList.push([submitTime, itemName, quantity, unit, idScheme, category, requestDate, note, email, emailLineManager]);
      }

      if (approvalStatus == 'Yes' && mailingStatus != 'Sent')
      {
        // send an order email when order information is valid
        var htmlTable = "<table cellspacing=0 cellpadding=\"5px\" style=\"border-collapse:collapse\", border=\"1px\">";
        htmlTable += "<tr>" + "<th style=\"background-color: yellow\", align=\"center\">" + "Thời gian" + "</th>";
        htmlTable += "<th style=\"background-color: yellow\", align=\"center\">" + "Email" + "</th>";
        htmlTable += "<th style=\"background-color: yellow\", align=\"center\">" + "Email (Line Manager)" + "</th>";
        htmlTable += "<th style=\"background-color: yellow\", align=\"center\">" + "Tên mặt hàng" + "</th>";
        htmlTable += "<th style=\"background-color: yellow\", align=\"center\">" + "Số lượng" + "</th>";
        htmlTable += "<th style=\"background-color: yellow\", align=\"center\">" + "Đơn vị tính" + "</th>";
        htmlTable += "<th style=\"background-color: yellow\", align=\"center\">" + "ID Scheme" + "</th>";
        htmlTable += "<th style=\"background-color: yellow\", align=\"center\">" + "Ngành hàng" + "</th>";
        htmlTable += "<th style=\"background-color: yellow\", align=\"center\">" + "Ngày cần" + "</th>";
        htmlTable += "<th style=\"background-color: yellow\", align=\"center\">" + "Ghi chú" + "</th>" + "</tr>";

        for (var k = 0; k < orderList.length; ++k)
        {
          htmlTable += "<tr>" + "<td align=\"center\">" + orderList[k][0] + "</td>";
          htmlTable += "<td align=\"center\">" + orderList[k][8] + "</td>";
          htmlTable += "<td align=\"center\">" + orderList[k][9] + "</td>";
          htmlTable += "<td align=\"center\">" + orderList[k][1] + "</td>";
          htmlTable += "<td align=\"center\">" + orderList[k][2] + "</td>";
          htmlTable += "<td align=\"center\">" + orderList[k][3] + "</td>";
          htmlTable += "<td align=\"center\">" + orderList[k][4] + "</td>";
          htmlTable += "<td align=\"center\">" + orderList[k][5] + "</td>";
          htmlTable += "<td align=\"center\">" + orderList[k][6] + "</td>";
          htmlTable += "<td align=\"center\">" + orderList[k][7] + "</td>";
          htmlTable += "</tr>";
        }

        htmlTable += "</table>";
        var subject = 'ĐỀ NGHỊ MUA HÀNG ĐƯỢC PHÊ DUYỆT ' + " - " + submitTime.toString();
        var body = "<div style=\"text-align:left;display: inline-block\">";
        body += "<p>" + "Dear anh/ chị " + email + "," + "</p>"; 
        body += "<p>" + "Đề nghị mua hàng lúc " + submitTime + " đã được phê duyệt bởi " + emailLineManager + "</p>";
        var ending = "<div style=\"text-align:left;display: inline-block\">";
        ending += "<p>" + "Best regards," + "</p>";
        ending += "<p>" + "Admin RD" + "</p>";

        var emailCC = emailLineManager;

        var emailBCC = 'cuongtq4@msc.masangroup.com, thond@msc.masangroup.com, anhntp@msc.masangroup.com';

        MailApp.sendEmail(email, subject, body, {htmlBody: body + htmlTable + ending, cc: emailCC, bcc: emailBCC});

        approvalSheet.getRange(1, statusIndex, 2).setValue('Sent');
        approvalSheet.getRange(1, submitTimeIndex, 1).setValue('Thời gian');
        approvalSheet.getRange(1, emailOrderIndex - 1, 1).setValue('Email');
        approvalSheet.getRange(1, statusIndex, 1).setValue('Tình trạng gửi');

        dataSheet.getRange(dataSheet.getLastRow() + 1, 1, orderList.length, orderList[0].length).setValues(orderList);
      }
      else if (approvalStatus == 'No' && mailingStatus != 'Sent')
      {
        // send an order email when order information is valid
        var htmlTable = "<table cellspacing=0 cellpadding=\"5px\" style=\"border-collapse:collapse\", border=\"1px\">";
        htmlTable += "<tr>" + "<th style=\"background-color: yellow\", align=\"center\">" + "Thời gian" + "</th>";
        htmlTable += "<th style=\"background-color: yellow\", align=\"center\">" + "Email" + "</th>";
        htmlTable += "<th style=\"background-color: yellow\", align=\"center\">" + "Email (Line Manager)" + "</th>";
        htmlTable += "<th style=\"background-color: yellow\", align=\"center\">" + "Tên mặt hàng" + "</th>";
        htmlTable += "<th style=\"background-color: yellow\", align=\"center\">" + "Số lượng" + "</th>";
        htmlTable += "<th style=\"background-color: yellow\", align=\"center\">" + "Đơn vị tính" + "</th>";
        htmlTable += "<th style=\"background-color: yellow\", align=\"center\">" + "ID Scheme" + "</th>";
        htmlTable += "<th style=\"background-color: yellow\", align=\"center\">" + "Ngành hàng" + "</th>";
        htmlTable += "<th style=\"background-color: yellow\", align=\"center\">" + "Ngày cần" + "</th>";
        htmlTable += "<th style=\"background-color: yellow\", align=\"center\">" + "Ghi chú" + "</th>" + "</tr>";

        for (var h = 0; h < orderList.length; ++h)
        {
          htmlTable += "<tr>" + "<td align=\"center\">" + orderList[h][0] + "</td>";
          htmlTable += "<td align=\"center\">" + orderList[h][8] + "</td>";
          htmlTable += "<td align=\"center\">" + orderList[h][9] + "</td>";
          htmlTable += "<td align=\"center\">" + orderList[h][1] + "</td>";
          htmlTable += "<td align=\"center\">" + orderList[h][2] + "</td>";
          htmlTable += "<td align=\"center\">" + orderList[h][3] + "</td>";
          htmlTable += "<td align=\"center\">" + orderList[h][4] + "</td>";
          htmlTable += "<td align=\"center\">" + orderList[h][5] + "</td>";
          htmlTable += "<td align=\"center\">" + orderList[h][6] + "</td>";
          htmlTable += "<td align=\"center\">" + orderList[h][7] + "</td>";
          htmlTable += "</tr>";
        }

        htmlTable += "</table>";
        var subject = 'ĐỀ NGHỊ MUA HÀNG KHÔNG ĐƯỢC PHÊ DUYỆT ' + " - " + submitTime.toString();
        var body = "<div style=\"text-align:left;display: inline-block\">";
        body += "<p>" + "Dear anh/ chị " + email + "," + "</p>"; 
        body += "<p>" + "Đề nghị mua hàng lúc " + submitTime + " không được phê duyệt bởi " + emailLineManager + "</p>";
        body += "<p>" + "Anh/ chị vui lòng thay đổi thông tin đơn đặt hàng tại đây: " +  "<a class=\"gmail-in-cell-link\", href="+link+">" + "Đơn mua hàng" + "</a>" + "</p>";
        var ending = "<div style=\"text-align:left;display: inline-block\">" ;
        ending += "<p>" + "Best regards," + "</p>";
        ending += "<p>" + "Admin RD" + "</p>";

        var emailCC = emailLineManager;
        var emailBCC = 'cuongtq4@msc.masangroup.com, thond@msc.masangroup.com, anhntp@msc.masangroup.com';

        MailApp.sendEmail(email, subject, body, {htmlBody: body + htmlTable + ending, cc: emailCC, bcc: emailBCC});       
        approvalSheet.getRange(1, statusIndex, 2).setValue('Sent');
        approvalSheet.getRange(1, submitTimeIndex, 1).setValue('Thời gian');
        approvalSheet.getRange(1, emailOrderIndex - 1, 1).setValue('Email');
        approvalSheet.getRange(1, statusIndex, 1).setValue('Tình trạng gửi');
      }
    }    
  }
}
