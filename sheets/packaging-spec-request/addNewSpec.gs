/**
 * This file contains a main function the writes new spec data to the sheet titled "new_spec_data" in the file named "Packaging Spec Data 2.0"
 * 
 */

// open a sheet titled "new_spec_data" by its name in the spec file
const newSpecSheet = specFile.getSheetByName('new_spec_data');
const newSpecHeader = newSpecSheet.getRange(1, 1, 1, newSpecSheet.getLastColumn()).getValues()[0]; // get values for the specified range on the sheet named "new_spec_data"
const newSpecData = newSpecSheet.getRange(2, 1, newSpecSheet.getLastRow() - 1, newSpecSheet.getLastColumn()).getValues(); // get values for the specified range on the sheet named "new_spec_data"

// return the column index of the the sheet named "new_spec_data" using findComlunIndex() method with a specified string text
const submitTimeIndex = findColumnIndex(newSpecHeader, 'submit_time');
const requesterEmailIndex = findColumnIndex(newSpecHeader, 'requester_email');
const picEmailIndex = findColumnIndex(newSpecHeader, 'pic_email');
const categoryIndex = findColumnIndex(newSpecHeader, 'category');
const subCategoryIndex = findColumnIndex(newSpecHeader, 'sub_category');
const documentClassificationIndex = findColumnIndex(newSpecHeader, 'document_classification');
const standardNamePackagingIndex = findColumnIndex(newSpecHeader, 'ten_tieu_chuan_bao_bi');
const packagingTypeIndex = findColumnIndex(newSpecHeader, 'packaging_type');
const categoryGroupIndex = findColumnIndex(newSpecHeader, 'nganh_hang_ap_dung');
const materialStructureIndex = findColumnIndex(newSpecHeader, 'cau_truc_vat_lieu_thay_doi');
const organizationIndex = findColumnIndex(newSpecHeader, 'organization');
const mailingStatusIndex = findColumnIndex(newSpecHeader, 'mailing_status');
const statusAddedToPackagingCodeIndex = findColumnIndex(newSpecHeader, 'adding_to_control_list_status');
const specStatusIndex = findColumnIndex(newSpecHeader, 'spec_status');

function addNewSpec() {
  // iterate through each row of the sheet named "new_spec_data"
  for (var i = 0; i < newSpecData.length; ++i)
  {
    var submitTime = newSpecData[i][submitTimeIndex  - 1];
    var formattedSubmitTime = Utilities.formatDate(new Date(submitTime), 'GMT+7', 'yyyy-MM-dd');
    var mailingStatus = newSpecData[i][mailingStatusIndex  - 1];

    if (submitTime != '')
    {
      if (mailingStatus != 'Sent')
      {
        // return requester's email value for each row and log this information
        var requesterEmail = newSpecData[i][requesterEmailIndex  - 1];
        console.log('Requester email: ' + requesterEmail);
        var requesterFolderId = getRequesterFolderID(requesterEmail); // return the id of the requester's folder based on requester's email and log this information
        console.log('Folder id: ' + requesterFolderId);

        // open the requester's folder by its id
        var requesterFolder = DriveApp.getFolderById(requesterFolderId);
        var requesterInformation = getRequesterInformation(requesterEmail); // get the information contained on sheet titled "requester_email" based on requester's email

        // return PIC's email using extractedPICEmail() method
        var picEmail = newSpecData[i][picEmailIndex  - 1];
        var extractedPICEmail = extractValueFromString(picEmail, /([a-zA-Z0-9._-]+@[a-zA-Z0-9._-]+\.[a-zA-Z0-9_-]+)/, 'No email address found: ');

        var category = newSpecData[i][categoryIndex  - 1];
        var subCategory = newSpecData[i][subCategoryIndex  - 1];
        var documentClassification = newSpecData[i][documentClassificationIndex  - 1];
        var standardNamePackaging = newSpecData[i][standardNamePackagingIndex  - 1];
        var packagingType = newSpecData[i][packagingTypeIndex  - 1];
        var categoryGroup = newSpecData[i][categoryGroupIndex  - 1];
        var materialStructure = newSpecData[i][materialStructureIndex  - 1];
        var organization = newSpecData[i][organizationIndex  - 1];
        
        // initialize a variable to store the information of the spec date by formatting the submit time in the format date/month/year
        var specDate = Utilities.formatDate(new Date(submitTime), 'GMT+7', 'dd/MM/yyyy');

        // create the new standard code based on the information of three elements, including the categoty, the sub-category and the document classification
        var newStandardCode = createNewStandardCode(category, subCategory, documentClassification);
        console.log('The new standard code is: ' + newStandardCode);

        // create a max incremental number for the new standard code, then create a new standard name by concatenating the new standard code and a max incremental number just generated
        var newStandardName = newStandardCode + createIncrementalNumber(newStandardCode);
        console.log('The new standard name is: ' + newStandardName);

        // find the standard file by matching the new standard code with each file name in the folder named "Standard Template"
        var standardFileId = findStandardFile(newStandardCode);

        if (standardFileId)
        {
          // create a new file name for the corresponding requester and create a copy of this file in the corresponding requester's folder
          var requesterFileName = newStandardName + '_' + standardNamePackaging + '_' + 'Ver01' + '_' + specDate;
          DriveApp.getFileById(standardFileId).makeCopy(requesterFileName, requesterFolder);

          // get the template file in the corresponding requester's folder
          var templateFile = SpreadsheetApp.open(requesterFolder.getFilesByName(requesterFileName).next());
          console.log('The template file id is: ' + templateFile.getId());

          // get the first sheet of the template file
          var templateSheet = templateFile.getSheets()[0];
          var templateHeader = templateSheet.getRange(1, 1, 1, templateSheet.getLastColumn()).getValues()[0];
          var templateData = templateSheet.getDataRange().getValues();

        
          if (templateSheet.getRange(1, findColumnIndex(templateHeader, 'mã tài liệu:')).getValue().toString().trim().toLowerCase() == 'mã tài liệu:')
          {
            // add a logo to the first sheet of the template file
            addLogo(templateSheet, organization);

            var specInformation = [standardNamePackaging, newStandardName, specDate, requesterInformation[0][0], requesterInformation[0][2]];

            // add the spec information to the first sheet of the template file
            addSpecInformation(templateSheet, templateData, templateHeader, specInformation);

            var standardizedOrganization = standardizeOrganization(organization);
            templateSheet.getRange(findRowIndex(templateData, 'nơi áp dụng') + 1, 1).setValue(standardizedOrganization);
            
            changeTemplateFileInformation(templateSheet, templateData, '01', specDate, 'Ban hành văn bản lần 01');
          }

          var subject = '[ADD-NEW-SPEC] Yêu Cầu Khởi Tạo SPEC Mới ' + newStandardName + '_' + formattedSubmitTime;

          var body = "<div style=\"text-align:left;display: inline-block\">";
          body += "<p>" + "Dear anh/chị RD Packaging," + "</p>"; 
          body += "<p>" + "Team SPEC đã nhận được yêu cầu khởi tạo SPEC mới " + newStandardName + " vào thời gian " + formattedSubmitTime + "</p>";
          body += "<p>" + "Anh/chị vui lòng truy cập vào đường dẫn đính kèm tại đây để kiểm tra thông tin: " + "<a class=\"standard-template-folder-link\", href="+"https://drive.google.com/drive/folders/"+requesterFolderId+">" + "Truy cập tại đây" + "</a>" + "</p>";

          var tableHeader = ["Thông tin yêu cầu", "Category", "Sub-category", "Phân cấp tài liệu", "Tên tiêu chuẩn bao bì", "Loại bao bì", "Ngành hàng áp dụng", "Tên tổ chức", "Cấu trúc vật liệu thay đổi", "Mã tiêu chuẩn"];

          var tableData = ["Khởi tạo tiêu chuẩn mới", category, subCategory, documentClassification, standardNamePackaging, packagingType, categoryGroup, organization, materialStructure, newStandardName];

          sendSPECEmail(subject, body, tableHeader, tableData, requesterEmail, extractedPICEmail);

          newSpecSheet.getRange(i + 2, mailingStatusIndex).setValue('Sent');
        
          if (newSpecSheet.getRange(i + 2, statusAddedToPackagingCodeIndex).getValue() != 'Added') 
          {
            console.log('The index in the new spec sheet: ' + i);

            for (var j = 0; j < packagingCode.length; ++j)
            {
              if (packagingCode[j][standardCodeIndex - 1] == '') 
              {
                console.log('The index in the packaging code sheet: ' + j);
              
                packagingCodeSheet.getRange(j + 2, standardCodeIndex).setValue(newStandardName);	
                packagingCodeSheet.getRange(j + 2, findColumnIndex(packagingCodeHeader, 'ten_tai_lieu')).setValue(standardNamePackaging);
                packagingCodeSheet.getRange(j + 2, findColumnIndex(packagingCodeHeader, 'phien_ban')).setValue('01');
                packagingCodeSheet.getRange(j + 2, findColumnIndex(packagingCodeHeader, 'ngay_hieu_luc')).setValue(formattedSubmitTime);
                packagingCodeSheet.getRange(j + 2, findColumnIndex(packagingCodeHeader, 'loai')).setValue(packagingType);
                packagingCodeSheet.getRange(j + 2, findColumnIndex(packagingCodeHeader, 'cau_truc_thay_doi')).setValue(materialStructure);	
                packagingCodeSheet.getRange(j + 2, findColumnIndex(packagingCodeHeader, 'status')).setValue('Updating');	
                packagingCodeSheet.getRange(j + 2, findColumnIndex(packagingCodeHeader, 'organization')).setValue(standardizedOrganization.replace("MSC và các nhà máy thành viên: ", "").trim());
                packagingCodeSheet.getRange(j + 2, findColumnIndex(packagingCodeHeader, 'nganh_hang')).setValue(categoryGroup);	
                packagingCodeSheet.getRange(j + 2, findColumnIndex(packagingCodeHeader, 'requester_email')).setValue(requesterEmail);
                break;
              }
            }

            newSpecSheet.getRange(i + 2, statusAddedToPackagingCodeIndex).setValue('Added');
            newSpecSheet.getRange(i + 2, specStatusIndex).setValue('Updating');
          }
        }
        else
        {
          console.log('No standard file corresponding to the provided standard code ' + newStandardCode + ' found in the folder');
          sendEmailMessageForNoFileFound(STANDARD_TEMPLATE_FOLDER_ID, newStandardCode, extractedPICEmail, requesterEmail);
          newSpecSheet.getRange(rowIndex + 2, mailingStatusIndex).setValue('Sent');
          newSpecSheet.getRange(rowIndex + 2, specStatusIndex).setValue('Deleted');
        }
      }
    }
  }
}
