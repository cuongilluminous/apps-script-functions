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

function findEmail(categoryCode)
{
  for (var i = 0; i < mapping.length; ++i)
  {
    var categoryCodeMapping = mapping[i][categoryCodeMappingIndex - 1];
    var email = mapping[i][emailIndex - 1];

    if (categoryCodeMapping.toString() == categoryCode.toString())
    {
      return email;
    }
  }
}
