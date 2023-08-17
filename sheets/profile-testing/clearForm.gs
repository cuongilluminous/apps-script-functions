function runClearForm()
{
  form.removeDestination();
  deleteItems(form);
  clearForm(form);
}

function clearForm(form)
{
  var items = form.getItems();
  while (items.length > 0)
  {
    form.deleteItem(items.pop());
  }
}

function deleteItems(form)
{
  var items = form.getItems();

  for (var i = items.length - 1; i >= 0; --i)
  {
    var item = form.getItems()[i]
    var itemType = item.getType();

    if (itemType != "PAGE_BREAK")
    {
      form.deleteItem(i)      
    }
  } 
}
