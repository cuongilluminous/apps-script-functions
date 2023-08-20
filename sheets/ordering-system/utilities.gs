function createOrderFile(name)
{
  var formFolder = DriveApp.getFolderById(formFolderID);
  var newOrderFile = SpreadsheetApp.create(name);
  DriveApp.getFileById(newOrderFile.getId()).moveTo(formFolder);

  return newOrderFile.getId();
}

function createApprovalForm(dateTime, email) 
{
  var name = dateTime + " - " + email;

  var approvalForm = FormApp.create(name);

  DriveApp.getFileById(approvalForm.getId()).moveTo(fileFolder);

  var question = approvalForm.addMultipleChoiceItem();
  question.setTitle('Anh/ chị vui lòng phê duyệt.')
    .setChoices([
        question.createChoice('Yes'),
        question.createChoice('No'),
  ])
  .setRequired(true);    
  approvalForm.setCollectEmail(true);
  approvalForm.setLimitOneResponsePerUser(true);
  
  createApprovalFile(approvalForm, name);

  var approvalFormURL = approvalForm.getPublishedUrl();

  return approvalFormURL;
}

function createApprovalFile(form, name) 
{
  form.setDestination(FormApp.DestinationType.SPREADSHEET, approvalFileID);
  SpreadsheetApp.openById(approvalFileID).getSheets()[0].setName(name);
}

function findIndex(value, string)
{
  for (var i = 0; i < value.length; ++i)
  {
    if (value[i].indexOf(string) > -1)
    {
      return i + 1;
    }
  }
}

function findOrderFile(name)
{
  var orderFileList = formFolder.getFiles();

  while (orderFileList.hasNext())
  {
    var orderFile = orderFileList.next();

    if (orderFile.getName().toString() == name.toString())
    {
      var id = orderFile.getId();

      return id;
    }
  }
}

function checkFile(i, sheet, time, url, email, formURL)
{
  var file = SpreadsheetApp.openByUrl(url);
  var allSheet = file.getSheets();

  if (allSheet[0].getName().toString() != 'order_template')
  {
    var subject = 'LINK ĐỀ NGHỊ MUA HÀNG KHÔNG ĐÚNG ' + " - " + time.toString();
    var body = "<div style=\"text-align:left;display: inline-block\">";
    body += "<p>" + "Dear anh/ chị " + email + "," + "</p>"; 
    body += "<p>" + "Link đề nghị mua hàng nhập không đúng theo yêu cầu." + "</p>";
    body += "<p>" + "Anh/ chị vui lòng nhập lại link đề nghị mua hàng tại đây: " +  "<a class=\"gmail-in-cell-link\", href="+formURL+">" + "Form đặt hàng" + "</a>" + "</p>";
    body += "<p>" + "Best regards," + "</p>";
    body += "<p>" + "Admin RD" + "</p>";

    var emailCC = 'cuongtq4@msc.masangroup.com, thond@msc.masangroup.com, anhntp@msc.masangroup.com';

    MailApp.sendEmail(email, subject, body, {htmlBody: body, cc: emailCC});

    sheet.getRange(i + 2, statusIndex).setValue('Sent');
  }
}

function checkID(id)
{
  return UrlFetchApp.fetch(`https://drive.google.com/thumbnail?id=${id}`, { muteHttpExceptions: true }).getResponseCode() == 200 ? true : false;
}

function findEmail(emailCheck)
{
  for (var i = 0; i < email.length; ++i)
  {
    if (email[i][emailIndex - 1] == emailCheck)
    {
      if (email[i][emailLineManagerIndex - 1] != 'ngalt@msc.masangroup.com')
      {
        return email[i][emailLineManagerIndex - 1];
      }
      else
      {
        return 'anhntp@msc.masangroup.com';
      }
    }
  }
}

function removeSpecialCharacterFromString(string)
{
  var string = string.toString();
  var newString = string.split('-')[0];

  return newString;
}
