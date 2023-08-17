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

function deleteEmptyRow(cacheSheet, cacheRange, cacheValue)
{
  // Checks if array is all empty values.
  const valueFilter = value => value !== '';
  const isRowEmpty = (row) => {
    return row.filter(valueFilter).length === 0;
  }

  // Maps the range values as an object with value (to test) and corresponding row index (with offset from selection).
  const rowToDelete = cacheValue.map((row, index) => ({ row, offset: index + cacheRange.getRowIndex() }))
    .filter(item => isRowEmpty(item.row)) // Test to filter out non-empty rows.
    .map(item => item.offset); //Remap to include just the row indexes that will be removed.

  // Combines a sorted, ascending list of indexes into a set of ranges capturing consecutive values as start/end ranges.
  // Combines sequential empty rows for faster processing.
  const rangeToDelete = rowToDelete.reduce((range, index) => {
    const currentRange = range[range.length - 1];
    if (currentRange) 
    {
      if (index === currentRange[1] + 1)
      {
        currentRange[1] = index;
        return range;
      }
    }
    range.push([index, index]);
    return range;
  }, []);

  // Deletes the rows using REVERSE order to ensure proper indexing is used.
  rangeToDelete.reverse().forEach(([start, end]) => cacheSheet.deleteRows(start, end - start + 1));
}

function findUniqueValue(data)
{
  var uniqueData = new Array();

  for (var i = 0; i < data.length; ++i)
  {
    duplicate = false;
    for (var j = 0; j < uniqueData.length; ++j)
    {
      if (data[i].join() == uniqueData[j].join())
      {
        duplicate = true;
      }
    }

    if (!duplicate)
    {
      uniqueData.push(data[i])
    }
  }

  return uniqueData;
}

function removeDuplication(uniqueData, data, input)
{
  var array = [];

  for (var i = 0; i < uniqueData.length; ++i)
  {
    var countDuplicate = 0;
    var index = [];

    for (var j = 0; j < data.length; ++j)
    {
      if (uniqueData[i].join() == data[j].join())
      {
        countDuplicate += 1;
        index.push(j);
      }
    }

    if (countDuplicate == 1)
    {
      array.push(input[index[0]]);
    }
    else if (countDuplicate > 1)
    {
      array.push(input[index[index.length - 1]]);
    }
  }

  return array;
}
