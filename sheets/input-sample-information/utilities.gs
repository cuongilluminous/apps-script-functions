// find supplier name
function findSupplierName(code)
{
  for (var i = 0; i < supplierValue.length; ++i)
  {
    var supplierCode = supplierValue[i][0];
    var supplierName = supplierValue[i][1];

    if (supplierCode.toString().trim().normalize('NFKC') == code.toString().trim().normalize('NFKC'))
    {
      return supplierName;
    }
  }
}

// randomize code
function randomeCode(array) {
  var codeList = array.slice(0);

  return function()
  {
    if (codeList.length < 1)
    { 
      codeList = array.slice(0);
    }
    var random = Math.floor(Math.random() * codeList.length);
    var item = codeList[random];
    var index = codeList.indexOf(item);
    codeList.splice(index, 1);
    return item;
  };
}

// remove underscore characters from string
function removeUnderscoreFromString(text)
{
  var result = text.split('-')[1];
  return result;
}

// convert an Excel file to Google Sheets
function convertExcel(name) 
{
  var file = DriveApp.getFilesByName(name);
  var excelFile = null;

  if (file.hasNext())
  {
    excelFile = file.next();
  }
  else
  {
    return null;
  }

  var blob = excelFile.getBlob();
  var config = 
  {
    title: excelFile.getName(),
    parents: [{id: excelFile.getParents().next().getId()}],
    mimeType: MimeType.GOOGLE_SHEETS
  };

  var spreadsheet = Drive.Files.insert(config, blob);
  return spreadsheet.id;
}

// find an index of a array
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

// convert an array to an array of objects
function convertArrayOfObject(array)
{
  var key = array.shift();
  var object = null;
  var output = [];

  for (var i = 0; i < array.length; ++i)
  {
    object = {};

    for (var j = 0; j < key.length; ++j)
    {
      object[key[j]] = array[i][j];
    }

    output.push(object)
  }

  return output
}

// group objects by multiple properties in an array then calculate their average values
function averageByGroup(array)
{
  var copy = {};
  var arrayGroupBy = array.reduce(function(object, data) {
    var key = data.test_decription + '-' + data.sample_name;
  
    if(!copy[key]) {
      copy[key] = {...data, count: 1}
      object.push(copy[key])
    } else {
      copy[key].actual_result += data.actual_result;
      copy[key].count += 1;
    }

    return object;
}, []);

  var averageArrayByGroup = Object.keys(arrayGroupBy).map(function(i) {
    var item = arrayGroupBy[i];
    return {
      sample_name: item.sample_name,
      test_description: item.test_decription,
      actual_result: item.actual_result/ item.count
    }
  })

  return averageArrayByGroup;
}

function deleteElementOfArray(sheet, range, value, array)
{ 
  const rowOffset = value.map(([sample_name, test_description, actual_result], index) => ({row: {sample_name, test_description, actual_result}, offset: index + range.getRowIndex()}))
    .filter((item) => !array.includes(item.row.test_description))
    .map((item) => item.offset)
  
  const rowToDelete = rowOffset.reduce((row, offset) => {
    const currentRow = row[row.length - 1];
    if (currentRow)
    {
      if (offset == currentRow[1] + 1)
      {
        currentRow[1] = offset;
        return row;
      }
    }
    row.push([offset, offset]);
    return row;
  }, [])

  // Deletes the rows using REVERSE order to ensure proper indexing is used.
    rowToDelete.reverse().forEach(([start, end]) => sheet.deleteRows(start, end - start + 1));
}


// transpose an array of values
function transposeArray(dataArray, rowIndex, columnIndex, dataIndex) 
{
  var result = {};
  var array = [];
  var newColumn = [];

  for (var i = 0; i < dataArray.length; i++)
  {
    if (!result[dataArray[i][rowIndex]])
    {
      result[dataArray[i][rowIndex]] = {};
    }
    result[dataArray[i][rowIndex]][dataArray[i][columnIndex]] = dataArray[i][dataIndex];
 
    if (newColumn.indexOf(dataArray[i][columnIndex]) == -1)
    {
      newColumn.push(dataArray[i][columnIndex]);
    }
  }
 
  newColumn.sort();
  var item = [];

  item.push('sample_name');
  item.push.apply(item, newColumn);
  array.push(item);
  
  for (var key in result)
  {
    item = [];
    item.push(key);
    for (var i = 0; i < newColumn.length; i++)
    {
      item.push(result[key][newColumn[i]] || "");
    }
    array.push(item);
  }

  return array;
}

function findSubmitTime(value)
{
  for (var i = 0; i < sensorySampleInformation.length; ++i)
  {
    var submitTime = sensorySampleInformation[i][sensorySampleInformationSubmitTimeIndex - 1];
    var formatSubmitTime = Utilities.formatDate(new Date(submitTime), 'GMT+7', 'ddMMYY');
    var lotNumber = sensorySampleInformation[i][sensorySampleInformationLotNumberIndex - 1].toString().trim().normalize('NFKC');

    if (lotNumber == value.toString().trim().normalize('NFKC'))
    {
      return formatSubmitTime
    }
  }
}
