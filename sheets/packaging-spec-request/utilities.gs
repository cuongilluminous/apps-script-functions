/**
 * This file contains utility functions that work with the file titled "Packaging Spec Data 2.0" and the file named "Packaging Code"
 */

// open a sheet titled "requester_email" by its name in the spec file
const requesterEmailSheet = specFile.getSheetByName('requester_email');
const requesterEmailHeader = requesterEmailSheet.getRange(1, 1, 1, requesterEmailSheet.getLastColumn()).getValues()[0]; // get values for the specified range on the sheet named "requester_email"
const requesterEmailData = requesterEmailSheet.getRange(2, 1, requesterEmailSheet.getLastRow() - 1, requesterEmailSheet.getLastColumn()).getValues(); // get values for the specified range on the sheet named "requester_email"

// return the column index of the the sheet named "requester_email" using findComlunIndex() method with a specified string text
const lineManagerIndex = findColumnIndex(requesterEmailHeader, 'line_manager');
const lineManagerEmailIndex = findColumnIndex(requesterEmailHeader, 'line_manager_email');
const requesterIndex = findColumnIndex(requesterEmailHeader, 'requester');
const folderIdIndex = findColumnIndex(requesterEmailHeader, 'folder_id');

// open a sheet titled "logo" by its name in the spec file
const logoSheet = specFile.getSheetByName('logo');
const logoHeader = logoSheet.getRange(1, 1, 1, logoSheet.getLastColumn()).getValues()[0];
const logo = logoSheet.getRange(2, 1, logoSheet.getLastRow() - 1, logoSheet.getLastColumn()).getValues();

// return the column index of the the sheet named "logo" using findComlunIndex() method with a specified string text
const logoLinkIndex = findColumnIndex(logoHeader, 'link');

// open a sheet titled "packaging_code" by its name in the file named "Packaging Code"
const packagingCodeSheet = packagingCodeFile.getSheetByName('packaging_code');
const packagingCodeHeader = packagingCodeSheet.getRange(1, 1, 1, packagingCodeSheet.getLastColumn()).getValues()[0]; // get values for the specified range on the sheet named "packaging_code"
const packagingCode = packagingCodeSheet.getRange(2, 1, packagingCodeSheet.getLastRow() - 1, packagingCodeSheet.getLastColumn()).getValues(); // get values for the specified range on the sheet named "packaging_code"

const standardCodeIndex = findColumnIndex(packagingCodeHeader, 'ma_tieu_chuan'); // return the column index of the the sheet named "packaging_code" using findComlunIndex() method with a specified string text

// open a sheet titled "item_code" by its name in the file named "Packaging Code"
const itemCodeSheet = packagingCodeFile.getSheetByName('item_code');
const itemCodeHeader = itemCodeSheet.getRange(1, 1, 1, itemCodeSheet.getLastColumn()).getValues()[0]; // get values for the specified range on the sheet named "item_code"
const itemCode = itemCodeSheet.getRange(2, 1, itemCodeSheet.getLastRow() - 1, itemCodeSheet.getLastColumn()).getValues(); // get values for the specified range on the sheet named "item_code"

/**
 * Gets the column index of a given file with a matching keyword (case insensitive).
 * 
 * @param {object} value - data of a given file.
 * @param {string} text - a text string contains the name of a table column in which to find.
 * @returns {object} the column index.
 */
function findColumnIndex(value, string)
{
  for (var i = 0; i < value.length; ++i)
  {
    if (value[i].toString().toLowerCase().indexOf(string) > -1)
    {
      return i + 1;
    }
  }
}

/**
 * Gets the row index of a given file with a matching keyword (case insensitive).
 * 
 * @param {object} value - data of a given file.
 * @param {string} text - a text string contains the cell value in which to find.
 * @returns {object} the row index.
 */
function findRowIndex(value, string)
{
  for (var i = 0; i < value.length; ++i)
  {
    var row = value[i];

    for (j = 0; j < row.length; ++j)
    {
      if (row[j].toString().toLowerCase().indexOf(string) > -1)
      {
        return i + 1
      }
    }
  }
}

/**
 * Gets the information contained on sheet titled "requester_email" based on requester's email
 * 
 * @param {string} email - a requester's email.
 * @returns {object} information - an array contains three elements, including a line manager's name, a line manager's email and a requester's name.
 */
function getRequesterInformation(email)
{
  var information = [];

  for (i = 0; i < requesterEmailData.length; ++i)
  {
    var requesterEmail = requesterEmailData[i][folderIdIndex - 2];

    if (requesterEmail == email)
    {
      var lineManager = requesterEmailData[i][lineManagerIndex - 1];
      var lineManagerEmail = requesterEmailData[i][lineManagerEmailIndex - 1];
      var requester = requesterEmailData[i][requesterIndex - 1];

      information.push([lineManager, lineManagerEmail, requester]);
      break;
    }
  }

  return information;
}

/**
 * Gets the information contained on sheet titled "requester_email" based on requester's email
 * 
 * @param {string} email - a requester's email.
 * @returns {object} information - an array contains three elements, including a line manager's name, a line manager's email and a requester's name.
 */
function standardizeOrganization(string)
{
  var extractedOrganization = string.split(',');

  var organization = [];

  for (var i = 0; i < extractedOrganization.length; ++i)
  {
    organization.push(extractedOrganization[i].split('-')[0].trim());
  }

  organization = 'MSC và các nhà máy thành viên: ' + organization.join(', ')
  
  return organization;
}

/**
 * Gets the information contained on sheet titled "requester_email" based on requester's email
 * 
 * @param {string} email - a requester's email.
 * @returns {object} information - an array contains three elements, including a line manager's name, a line manager's email and a requester's name.
 */
function extractValueFromString(string, regex, errorMessage)
{
  var valueMatch = string.match(regex);

  if (valueMatch)
  {
    var extractedValue = valueMatch[0];
    return extractedValue
  }
  else
  {
    console.log(errorMessage + string)
    return null;
  }
}

/**
 * Gets the information contained on sheet titled "requester_email" based on requester's email
 * 
 * @param {string} email - a requester's email.
 * @returns {object} information - an array contains three elements, including a line manager's name, a line manager's email and a requester's name.
 */
function createNewStandardCode(category, subCategory, documentClassification)
{
  return 'TC' + extractValueFromString(category, /\d+/, 'No number found in the string: ') + extractValueFromString(subCategory, /\d+/, 'No number found in the string: ') + "B" + extractValueFromString(documentClassification, /\d+/, 'No number found in the string: ');
}

/**
 * Gets the information contained on sheet titled "requester_email" based on requester's email
 * 
 * @param {string} email - a requester's email.
 * @returns {object} information - an array contains three elements, including a line manager's name, a line manager's email and a requester's name.
 */
function sendEmailMessageForNoFileFound(folderId, code, picEmail, requesterEmail)
{
  var subject = '[ERROR-IN-ADDING-NEW-SPEC] Không Tìm Thấy File Tiêu Chuẩn Trong Hệ Thống Lưu Trữ';

  var body = "<div style=\"text-align:left;display: inline-block\">";
  body += "<p>" + "Dear anh/chị SPEC," + "</p>"; 
  body += "<p>" + "Hệ thống phát hiện không có file tiêu chuẩn tương ứng với thông tin code " + code + " để tạo SPEC mới." + "</p>";
  body += "<p>" + "Anh chị vui lòng tạo file tiêu chuẩn và cập nhập vào link đính kèm tại đây: " + "<a class=\"standard-template-folder-link\", href="+"https://drive.google.com/drive/folders/"+folderId+">" + "Standard Template Folder" + "</a>" + "</p>";

  var ending = "<div style=\"text-align:left;display: inline-block\">";
  ending += "<p>" + "Best regards," + "</p>";
  ending += "<p>" + "TDM" + "</p>";

  MailApp.sendEmail(picEmail, subject, body,  {htmlBody: body + ending, cc: requesterEmail});
}

/**
 * Gets the information contained on sheet titled "requester_email" based on requester's email
 * 
 * @param {string} email - a requester's email.
 * @returns {object} information - an array contains three elements, including a line manager's name, a line manager's email and a requester's name.
 */
function findStandardFile(string)
{
  var string = string.toString().trim().normalize('NFKC');

  var allFile = standardTemplateFolder.getFiles();

  while (allFile.hasNext())
  {
    var file = allFile.next();
    var fileName = file.getName().toString().trim().normalize('NFKC');

    if (string == fileName)
    {
      var standardFile = file;
      var standardFileId = standardFile.getId();
      return standardFileId;
    }
  }
}

/**
 * Gets the information contained on sheet titled "requester_email" based on requester's email
 * 
 * @param {string} email - a requester's email.
 * @returns {object} information - an array contains three elements, including a line manager's name, a line manager's email and a requester's name.
 */
function createIncrementalNumber(code)
{
  var maxIncrementalNumber = 1;

  for (var i = 0; i < packagingCode.length; ++i)
  {
    var standardCode = packagingCode[i][standardCodeIndex - 1];

    var extractedStandardCode = extractValueFromString(standardCode, /([A-Z]{2}\d{2}[A-Z]\d{2})/, 'No standard code found in the string: ');

    if (extractedStandardCode == code)
    {
      var number = parseInt(standardCode.replace(extractedStandardCode, '').trim());

      if (number > maxIncrementalNumber)
      {
        maxIncrementalNumber = number;
      }
      else if (number == maxIncrementalNumber)
      {
        maxIncrementalNumber = parseInt(maxIncrementalNumber) + 1;
      }
    }
  }

  if (maxIncrementalNumber < 10)
  {
    maxIncrementalNumber = `00${maxIncrementalNumber}`;
  }
  else if (maxIncrementalNumber < 100)
  {
    maxIncrementalNumber = `0${maxIncrementalNumber}`;
  }
  else
  {
    maxIncrementalNumber = maxIncrementalNumber;
  }

  return maxIncrementalNumber;
}

/**
 * Gets the information contained on sheet titled "requester_email" based on requester's email
 * 
 * @param {string} email - a requester's email.
 * @returns {object} information - an array contains three elements, including a line manager's name, a line manager's email and a requester's name.
 */
function getRequesterFolderID(email)
{
  for (var i = 0; i < requesterEmailData.length; ++i)
  {
    var requesterEmail = requesterEmailData[i][folderIdIndex - 2];
    
    if (requesterEmail == email)
    {
      var folderId = requesterEmailData[i][folderIdIndex - 1];
    }
  }

  return folderId;
}

/**
 * Gets the information contained on sheet titled "requester_email" based on requester's email
 * 
 * @param {string} email - a requester's email.
 * @returns {object} information - an array contains three elements, including a line manager's name, a line manager's email and a requester's name.
 */
function findLogoLink(string)
{
  for (var i = 0; i < logo.length; ++i)
  {
    var organization = logo[i][logoLinkIndex - 2];

    if (organization == string)
    {
      var link = logo[i][logoLinkIndex - 1]
      break;
    }
  }

  return link;
}

/**
 * Gets the information contained on sheet titled "requester_email" based on requester's email
 * 
 * @param {string} email - a requester's email.
 * @returns {object} information - an array contains three elements, including a line manager's name, a line manager's email and a requester's name.
 */
function addLogo(sheet, organization)
{
  // input logo into a template file' s sheet containing the information of packaging specification
  if (organization.includes('Phúc Long'))
  {
    var logoLink = findLogoLink('Phúc Long');
  }
  else if (organization.includes('VHP-Vĩnh Hảo'))
  {
    var logoLink = findLogoLink('Vĩnh Hảo');
  }
  else if (organization.includes('VNP-Vinacafe Biên Hòa') || organization.includes('VFF-Vinacafé chi nhánh MSI') || organization.includes('LTP-Vinacafe Long Thành'))
  {
    var logoLink = findLogoLink('Vinacafe');
  }
  else
  {
    var logoLink = findLogoLink('Masan');
  }

  var allImage = sheet.getImages();
          
  if (allImage.length > 0 &
      sheet.getRange(1, 1).getValue() == '')
  {
    allImage[0].remove();
  }

  sheet.getRange(1, 1).setValue('=IMAGE("' + logoLink + '")');
}

/**
 * Gets the information contained on sheet titled "requester_email" based on requester's email
 * 
 * @param {string} email - a requester's email.
 * @returns {object} information - an array contains three elements, including a line manager's name, a line manager's email and a requester's name.
 */
function addSpecInformation(sheet, data, header, specInformation)
{
  var standardNameSpecIndex = findColumnIndex(header, 'tc');
  var investigationIndex = findRowIndex(data, 'xem xét');
  var approvalData = sheet.getRange(investigationIndex, 1, 1, sheet.getLastColumn()).getValues()[0]
  
  // input the standard name packaging into a template file' s sheet containing the information of packaging specification
  sheet.getRange(findRowIndex(data, 'tiêu chuẩn kỹ thuật bao bì') + 2, findColumnIndex(header, 'tiêu chuẩn kỹ thuật bao bì')).setValue(specInformation[0]);

  // input the new standard name into a template file' s sheet containing the information of packaging specification
  sheet.getRange(findRowIndex(data, 'tc'), standardNameSpecIndex).setValue(specInformation[1]);

  // input the spec date into a template file' s sheet containing the information of packaging specification
  sheet.getRange(findRowIndex(data, 'ngày hiệu lực'), standardNameSpecIndex).setNumberFormat('dd/mm/yy');
  sheet.getRange(findRowIndex(data, 'ngày hiệu lực'), standardNameSpecIndex).setValue(specInformation[2]);

  // input the version into a template file' s sheet containing the information of packaging specification
  sheet.getRange(findRowIndex(data, 'phiên bản'), standardNameSpecIndex).setNumberFormat("@");
  sheet.getRange(findRowIndex(data, 'phiên bản'), standardNameSpecIndex).setValue('01'.toString());

  // input the status change into a template file' s sheet containing the information of packaging specification
  sheet.getRange(findRowIndex(data, 'thay thế'), standardNameSpecIndex).setValue('N/A');

  sheet.getRange(findRowIndex(data, 'phê duyệt') + 2, findColumnIndex(approvalData, 'xem xét')).setValue(specInformation[3]);
  sheet.getRange(findRowIndex(data, 'phê duyệt') + 2, findColumnIndex(approvalData, 'soạn thảo')).setValue(specInformation[4]);
}

/**
 * Gets the information contained on sheet titled "requester_email" based on requester's email
 * 
 * @param {string} email - a requester's email.
 * @returns {object} information - an array contains three elements, including a line manager's name, a line manager's email and a requester's name.
 */
function updateSpecInformation(sheet, data, header, updatingSpecInformation)
{
  var standardNameSpecIndex = findColumnIndex(header, 'tc');
  var investigationIndex = findRowIndex(data, 'xem xét');
  var approvalData = sheet.getRange(investigationIndex, 1, 1, sheet.getLastColumn()).getValues()[0];

  // input the new standard name into a template file' s sheet containing the information of packaging specification
  sheet.getRange(findRowIndex(data, 'tc'), standardNameSpecIndex).setValue(updatingSpecInformation[0]);

  // input the spec date into a template file' s sheet containing the information of packaging specification
  sheet.getRange(findRowIndex(data, 'ngày hiệu lực'), standardNameSpecIndex).setNumberFormat('dd/mm/yy');
  sheet.getRange(findRowIndex(data, 'ngày hiệu lực'), standardNameSpecIndex).setValue(updatingSpecInformation[1]);

  // input the next version into a template file' s sheet containing the information of packaging specification
  sheet.getRange(findRowIndex(data, 'phiên bản'), standardNameSpecIndex).setNumberFormat("@");
  sheet.getRange(findRowIndex(data, 'phiên bản'), standardNameSpecIndex).setValue(updatingSpecInformation[2]);

  // input the status change into a template file' s sheet containing the information of packaging specification
  sheet.getRange(findRowIndex(data, 'thay thế'), standardNameSpecIndex).setValue(updatingSpecInformation[0] + '_' + updatingSpecInformation[3]);

  sheet.getRange(findRowIndex(data, 'phê duyệt') + 2, findColumnIndex(approvalData, 'xem xét')).setValue(updatingSpecInformation[4]);
  sheet.getRange(findRowIndex(data, 'phê duyệt') + 2, findColumnIndex(approvalData, 'soạn thảo')).setValue(updatingSpecInformation[5]);
}

/**
 * Gets the information contained on sheet titled "requester_email" based on requester's email
 * 
 * @param {string} email - a requester's email.
 * @returns {object} information - an array contains three elements, including a line manager's name, a line manager's email and a requester's name.
 */
function changeTemplateFileInformation(sheet, data, version, date, note)
{
  var versionStatusIndex = findRowIndex(data, 'lịch sử sửa đổi');

  var versionData = sheet.getRange(versionStatusIndex + 1, 1, 1, sheet.getLastColumn()).getValues()[0];

  sheet.getRange(versionStatusIndex + 2, 1).setValue(version);
  sheet.getRange(versionStatusIndex + 2, findColumnIndex(versionData, 'ngày hiệu lực')).setValue(date);
  sheet.getRange(versionStatusIndex + 2, findColumnIndex(versionData, 'nội dung thay đổi')).setValue(note);

  return sheet;
}

/**
 * Gets the information contained on sheet titled "requester_email" based on requester's email
 * 
 * @param {string} email - a requester's email.
 * @returns {object} information - an array contains three elements, including a line manager's name, a line manager's email and a requester's name.
 */
function sendErrorEmail(error, picEmail, requesterEmail) 
{
  var subject = '[ERROR-IN-ADDING-NEW-SPEC] Lỗi Hệ Thống Trong Quá Trình Tạo SPEC Mới';

  var body = "<div style=\"text-align:left;display: inline-block\">";
  body += "<p>" + "Dear anh/chị SPEC," + "</p>"; 
  body += "<p>" + "Trong quá trình tạo SPEC mơi, hệ thống phát hiện lỗi: " + error + "</p>";
  body += "<p>" + "Anh/chị vui lòng thông báo team TDM để vào kiểm tra và khắc phục lỗi." + "</p>";

  var ending = "<div style=\"text-align:left;display: inline-block\">";
  ending += "<p>" + "Best regards," + "</p>";
  ending += "<p>" + "TDM" + "</p>";

  MailApp.sendEmail(picEmail, subject, body,  {htmlBody: body + ending, cc: requesterEmail});
}

/**
 * Gets the information contained on sheet titled "requester_email" based on requester's email
 * 
 * @param {string} email - a requester's email.
 * @returns {object} information - an array contains three elements, including a line manager's name, a line manager's email and a requester's name.
 */
function sendSPECEmail(title, body, tableHeader, tableData, requesterEmail, picEmail)
{
  var subject = title;

  var end = "<div style=\"text-align:left;display: inline-block\">";
  end += "<p>" + "Best regards," + "</p>";
  end += "<p>" + "RD SPEC" + "</p>";

  var htmlTable = "<table cellspacing=0 cellpadding=\"5px\" style=\"border-collapse:collapse\", border=\"1px\">";
  for (i = 0; i < tableHeader.length; ++i)
  {
    htmlTable += "<th style=\"background-color: yellow\", align=\"center\">" + tableHeader[i] + "</th>";
  }

  htmlTable += "<tr>";
  for (j = 0; j < tableData.length; ++j)
  {
    htmlTable += "<td align=\"center\">" + tableData[j] + "</td>";
  }
  htmlTable += "</td>";
  htmlTable += "</table>";

  var ccEmail = picEmail + ", cuongtq4@msc.masangroup.com";
  MailApp.sendEmail(requesterEmail, subject, body,  {htmlBody: body + htmlTable + end, cc: ccEmail});
}

/**
 * Gets the information contained on sheet titled "requester_email" based on requester's email
 * 
 * @param {string} email - a requester's email.
 * @returns {object} information - an array contains three elements, including a line manager's name, a line manager's email and a requester's name.
 */
function createVersion(string, data, header) 
{
  var nextVersion = 1;

  for (var i = 0; i < data.length; ++i)
  {
    var standardName = data[i][standardCodeIndex - 1];

    if (string.toString().trim() == standardName.toString().trim()) 
    {
      var specVersion = data[i][findColumnIndex(header, 'phien_ban') - 1];
      console.log('The current version is: ' + specVersion);

      if (nextVersion <= specVersion)
      {
        nextVersion = parseInt(specVersion) + 1;
      }
    }
  }

  if (nextVersion < 10) 
  {
    nextVersion = `0${nextVersion}`;
  }

  var currentVersion = parseInt(nextVersion) - 1;
          
  if (currentVersion < 10) 
  {
    currentVersion = `0${currentVersion}`;
  }

  var version = [nextVersion, currentVersion]

  return version;
}

/**
 * Gets the information contained on sheet titled "requester_email" based on requester's email
 * 
 * @param {string} email - a requester's email.
 * @returns {object} information - an array contains three elements, including a line manager's name, a line manager's email and a requester's name.
 */
function findSpecFile(string, picEmail, requesterEmail) 
{
  var allFolder = packagingFolder.getFolders();

  while (allFolder.hasNext()) 
  {
    var folder = allFolder.next();
    var folderName = folder.getName();

    if (string.toString().trim() == folderName.toString().trim()) 
    {
      var specFolder = folder;
      break;
    }
  }

  var allFile = specFolder.getFiles();

  var startingVersion = 1;

  while (allFile.hasNext())
  {
    var file = allFile.next();

    if (file.getMimeType() == "application/vnd.google-apps.spreadsheet")
    {
      var fileName = file.getName();

      var standardName = extractValueFromString(fileName, /([A-Z]{2}\d{2}[A-Z]\d+)/, 'No standard name found in the file name: ');

      if (standardName.toString().trim() == string.toString().trim()) 
      {
        var versionText = fileName.split("_")[2];

        var version = parseInt(extractValueFromString(versionText, /\d+/, 'No number found in the string: '));
        
        if (version >= startingVersion) 
        {
          var specFile = file;

          return specFile;
        }
      }
      else
      {
        console.log('No spec file corresponding to the provided standard name ' + string + ' found in the folder');
        sendEmailMessageForNoFileFound(packagingFolderId, string, picEmail, requesterEmail);
        updatingSpecSheet.getRange(rowIndex + 2, findColumnIndex(updatingSpecData, 'spec_status')).setValue('Deleted');
      } 
    }  
  }
}

/**
 * Gets the information contained on sheet titled "requester_email" based on requester's email
 * 
 * @param {string} email - a requester's email.
 * @returns {object} information - an array contains three elements, including a line manager's name, a line manager's email and a requester's name.
 */
function addPackagingCode(string, list)
{
  for (var i = 0; i < list.length; ++i)
  {
    var itemCode = itemCodeSheet.getRange(2, 1, itemCodeSheet.getLastRow(), itemCodeSheet.getLastColumn()).getValues();

    for (var j = 0; j < itemCode.length; ++j)
    {
      if (itemCode[j][findColumnIndex(itemCodeHeader, 'item_code') - 1] == '')
      {
        itemCodeSheet.getRange(j + 2, findColumnIndex(itemCodeHeader, 'item_code')).setValue(list[i].trim());
        itemCodeSheet.getRange(j + 2, findColumnIndex(itemCodeHeader, 'ma_tieu_chuan')).setValue(string);
        break;
      }
    }
  }
}
