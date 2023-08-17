function createProfileTest()
{
  var currentDate = new Date();
  var formatCurrentDate = Utilities.formatDate(new Date(currentDate), 'UTC', 'yyyy-MM-dd');

  // return the 'Profile Test' form
  var formTest = 'Profile Test' + ' - ' + formName + ' - ' + formatCurrentDate;
  var form = FormApp.openByUrl(formEditURL).setTitle(formTest);
  form.removeDestination();

  // create full name question
  var employeeCode = form.addListItem().setTitle('Anh/Chị vui lòng chọn MSNV của mình:').setHelpText('VD: 21SF15443').setRequired(true);
  var employeeCodeList = [];
  for (var i = 0; i < employee.length; ++i)
  {
    employeeCodeList.push(employeeCode.createChoice(employee[i]))
  }
  employeeCode.setChoices(employeeCodeList);

  createProfileQuestion(form, newSheet, newSheetValue, mapping, commentSection);

  // create 'Comment' section
  var commentSection = form.addPageBreakItem().setTitle('Nhận xét')
  form.addParagraphTextItem().setTitle('Các off-notes khác (nếu không có thì Anh/Chị nhập KHÔNG):').setRequired(true);

  // create 'Result' section
  form.addPageBreakItem().setTitle('Kết quả');

  form.addPageBreakItem().setGoToPage(FormApp.PageNavigationType.SUBMIT);

  form.addPageBreakItem().setGoToPage(commentSection);

  form.setDestination(FormApp.DestinationType.SPREADSHEET, profileTestID);
}

function createProfileQuestion(form, sheet, value, mapping, commentSection)
{
  var testSection = form.addPageBreakItem().setTitle(formName);
  var sampleIDIndex = findIndex(mappingHeader, "sample_id");
  var updateStatusIndex = findIndex(mappingHeader, "update_status");

  if (sheet.getName().toString() == 'mo_ta')
  {
    var questionCode = form.addListItem().setTitle('Anh/Chị vui lòng chọn mã số mẫu Anh/Chị nhận được:');

    var codeList = [];

    for (var i = 0; i < mapping.length; ++i)
    {
      var code = mapping[i][sampleIDIndex - 1];
      var updateStatus = mapping[i][updateStatusIndex - 1];

      if (updateStatus != "Done")
      {
        codeList.push(questionCode.createChoice(code));
        questionCode.setChoices(codeList).setRequired(true);
      }
    }
    
    var numberValidation = FormApp.createTextValidation()
    .setHelpText('Anh/Chị vui lòng nhập đúng với điều kiện thang đo.')
    .requireTextContainsPattern('^10$|^[0-9](\.5){0,1}$')
    .build();

    for (var j = 0; j < value[0].length; ++j)
    {
      var testValue = sheet.getRange(1, j + 1, sheet.getLastRow(), 1).getValues().filter(r => r != '');
      if (testValue.length == 1)
      {
        var textQuestion = testValue[0];
        form.addTextItem().setTitle(textQuestion).setHelpText(newScale).setRequired(true).setValidation(numberValidation);
      }
      else
      {
        for (var k = 1; k < testValue.length; ++k)
        {
          var textQuestion = testValue[0] + ' ' + testValue[k];
          form.addTextItem().setTitle(textQuestion).setHelpText(newScale).setRequired(true).setValidation(numberValidation);
        }
      }
    }
  }

  return testSection;
}
