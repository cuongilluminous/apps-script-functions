function sendSensoryMail()
{
  var response = responseSheet.getRange(2, 1, sensoryFile.getLastRow() - 1, sensoryFile.getLastColumn()).getValues();

  for (var i = 0; i < response.length; ++i)
  {
    var submitTime = response[i][submitTimeIndex - 1];
    var formatSubmitTime = Utilities.formatDate(new Date(submitTime), 'GMT+7', 'yyyy-MM-dd HH-mm-ss');

    if (submitTime != '')
    {
      var status = response[i][statusIndex - 1];

      if (status == '')
      {
        var employeeCode = response[i][employeeCodeIndex - 1];
        var categoryCode = response[i][categoryCodeIndex - 1];
        console.log(categoryCode);
        var productName = response[i][productNameIndex - 1];
        var projectID = response[i][projectIDIndex - 1];
        var itemCode = response[i][itemCodeIndex - 1];
        var formulationCode = response[i][formulationCodeIndex - 1];
        var objective = response[i][objectiveIndex - 1];
        var note = response[i][noteIndex - 1];
        var requesterEmail = response[i][requesterEmailIndex - 1];
        var context = response[i][contextIndex - 1];
        var expectedDate = response[i][expectedDateIndex - 1];
        var formatExpectedDate = Utilities.formatDate(new Date(expectedDate), 'GMT+7', 'yyyy-MM-dd');
        var productionSite = response[i][productionSiteIndex - 1];

        var projectEmail = findEmail(categoryCode);
        console.log(projectEmail);

        var htmlTable = "<table cellspacing=0 cellpadding=\"5px\" style=\"border-collapse:collapse\", border=\"1px\">";
        htmlTable += "<tr>" + "<th style=\"background-color: yellow\", align=\"center\">" + "Mã số nhân viên" + "</th>";
        htmlTable += "<th style=\"background-color: yellow\", align=\"center\">" + "Ngành hàng" + "</th>";
        htmlTable += "<th style=\"background-color: yellow\", align=\"center\">" + "Tên sản phẩm" + "</th>";
        htmlTable += "<th style=\"background-color: yellow\", align=\"center\">" + "Project Id" + "</th>";
        htmlTable += "<th style=\"background-color: yellow\", align=\"center\">" + "Item code" + "</th>";
        htmlTable += "<th style=\"background-color: yellow\", align=\"center\">" + "Mã công thức" + "</th>" + "</tr>";

        htmlTable += "<tr>" + "<td align=\"left\">" + employeeCode + "</td>";
        htmlTable += "<td align=\"left\">" + categoryCode + "</td>";
        htmlTable += "<td align=\"left\">" + productName + "</td>";
        htmlTable += "<td align=\"left\">" + projectID + "</td>";
        htmlTable += "<td align=\"left\">" + itemCode + "</td>";
        htmlTable += "<td align=\"left\">" + formulationCode + "</td>";
        htmlTable += "</tr>";

        htmlTable += "</table>";

        var subject = '[YÊU CẦU TEST SENSORY]' + '_' + formatSubmitTime;

        var body = "<div style=\"text-align:left;display: inline-block\">";
        body += "<p>" + "Dear anh/chị " + requesterEmail + ","; 
        body += "<p>" + "Team SCI xác nhận anh/chị vừa gửi yêu cầu test Sensory." + "</p>";
        body += "<p>" + "Thông tin bối cảnh thí nghiệm được thể hiện ở phía dưới: " + "</p>";
        body += "<ol>" + "<li>" + "Thời gian: " + formatSubmitTime + "</li>";
        body += "<li>" + "Bối cảnh thí nghiệm: " + context + "</li>";
        body += "<li>" + "Mục đích test: " + objective + "</li>";
        body += "<li>" + "Lưu ý: " + note + "</li>";
        body += "<li>" + "Ngày dự kiến giao mẫu: " + formatExpectedDate + "</li>";
        body += "<li>" + "Địa điểm sản xuất: " + productionSite + "</li>" + "</ol>";
        body += "<p>" + "Ngoài ra, thông tin sản phẩm yêu cầu test được mô tả ở bảng phía dưới." + "</p>";
        body += "<p>" + "Best regards," + "</p>";
        body += "<p>" + "SCI" + "</p>";

        var ccEmail = projectEmail +", " + 'thond@msc.masangroup.com, cuongtq4@msc.masangroup.com, longnn@msc.masangroup.com, aivd@msc.masangroup.com, uyenptx@msc.masangroup.com, myntn@msc.masangroup.com, anhlnd@msc.masangroup.com, tamlm@msc.masangroup.com, Sarun.phaosathienpan@msc.masangroup.com';
        MailApp.sendEmail(requesterEmail, subject, body,  {htmlBody: body + htmlTable, cc: ccEmail});

        responseSheet.getRange(i + 2, statusIndex).setValue('Sent');
      }
    }
    else
    {
      break;
    }
  }
}
