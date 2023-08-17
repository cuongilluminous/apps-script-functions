function updateForm()
{
  var form = FormApp.openByUrl(formEditURL);
  var question = form.getItems();

  // update list code for new method
  updateCodeList(question, "Anh/Chị vui lòng chọn mã số mẫu Anh/Chị nhận được:", 3);
}

function updateCodeList(question, string, index)
{
  var questionCodeListID = [];

  for (var i = 0; i < question.length; ++i)
  {
    if (question[i].getTitle() == string)
    {
      questionCodeListID.push(question[i].getId())
    }
  }

  var codeList = [];

  var updateStatusIndex = findIndex(mappingHeader, "update_status");
  
  for (var j = 0; j < mapping.length; ++j)
  {
    var updateStatus = mapping[j][updateStatusIndex - 1];
    if (updateStatus != "Done")
    {
      codeList.push(mapping[j][index]);
    }
  }

  if (codeList.length != 0)
  {
    var questionCodeList = form.getItemById(questionCodeListID[0]);

    questionCodeList.asListItem().setChoiceValues(codeList);
  }
  else
  {
    console.log('The array is empty');
  }
}
