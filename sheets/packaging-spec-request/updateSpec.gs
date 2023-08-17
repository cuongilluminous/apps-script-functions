// open a sheet titled "updating_spec_data" by its name in the spec file
const updatingSpecSheet = specFile.getSheetByName('updating_spec_data');
const updatingSpecHeader = updatingSpecSheet.getRange(1, 1, 1, updatingSpecSheet.getLastColumn()).getValues()[0]; // get values for the specified range on the sheet named "updating_spec_data"
const updatingSpecData = updatingSpecSheet.getRange(2, 1, updatingSpecSheet.getLastRow() - 1, updatingSpecSheet.getLastColumn()).getValues(); // get values for the specified range on the sheet named "updating_spec_data"

function updateSpec() {
  for (var i = 0; i < updatingSpecData.length; ++i)
  {
    var submitTime = updatingSpecData[i][submitTimeIndex - 1];
    var formattedSubmitTime = Utilities.formatDate(new Date(submitTime), 'GMT+7', 'yyyy-MM-dd');
    var mailingStatus = updatingSpecData[i][findColumnIndex(updatingSpecHeader, 'mailing_status') - 1];

    if (submitTime != '')
    {
      if (mailingStatus != 'Sent')
      {
        var requesterEmail = updatingSpecData[i][requesterEmailIndex  - 1];
        console.log('Requester email: ' + requesterEmail);
        var requesterFolderId = getRequesterFolderID(requesterEmail);
        console.log('Folder id: ' + requesterFolderId);
        var requesterFolder = DriveApp.getFolderById(requesterFolderId);
        var requesterInformation = getRequesterInformation(requesterEmail);
        var picEmail = updatingSpecData[i][picEmailIndex  - 1];
        var extractedPICEmail = extractValueFromString(picEmail, /([a-zA-Z0-9._-]+@[a-zA-Z0-9._-]+\.[a-zA-Z0-9_-]+)/, 'No email address found: ');
        var standardName = updatingSpecData[i][findColumnIndex(updatingSpecHeader, 'ma_tieu_chuan') - 1];
        var materialStructure = updatingSpecData[i][findColumnIndex(updatingSpecHeader, 'cau_truc_vat_lieu_thay_doi') - 1];
        var note = updatingSpecData[i][findColumnIndex(updatingSpecHeader, 'note') - 1];
        var organization = updatingSpecData[i][findColumnIndex(updatingSpecHeader, 'organization') - 1];

        // initialize a variable to store the information of the spec date by converting submit time to the format 'dd/MM/yyyy'
        var specDate = Utilities.formatDate(new Date(submitTime), 'GMT+7', 'dd/MM/yyyy');

        var version = createVersion(standardName, packagingCode, packagingCodeHeader);
        console.log('The next version is: ' + version[0]);

        var specFile = findSpecFile(standardName, extractedPICEmail, requesterEmail);

        if (specFile)
        {
          var specFileId = specFile.getId();
          var specFileName = specFile.getName().trim();
          console.log('The spec file name is: ' + specFileName)

          var requesterFileName = specFileName.replace(extractValueFromString(specFileName, /_[A-Za-z\s]+\d+_(19[0-9]{2}|2[0-9]{3})-(0[1-9]|1[012])-([123]0|[012][1-9]|31)/, 'No version and spec date found in the spec file name: '), '').trim() + '_' + 'Ver' + version[0] + '_' + formattedSubmitTime;
          console.log('The requester file name is: ' + requesterFileName);

          DriveApp.getFileById(specFileId).makeCopy(requesterFileName, requesterFolder);

          // get the template file in the corresponding requester folder
          var templateFile = SpreadsheetApp.open(requesterFolder.getFilesByName(requesterFileName).next());
          console.log('The template file id is: ' + templateFile.getId());

          // get a template file' s sheet containing the information of packaging specification
          var templateSheet = templateFile.getSheets()[0];
          var templateHeader = templateSheet.getRange(1, 1, 1, templateSheet.getLastColumn()).getValues()[0];
          var templateData = templateSheet.getDataRange().getValues();

          if (templateSheet.getRange(1, findColumnIndex(templateHeader, 'mã tài liệu:')).getValue().toString().trim().toLowerCase() == 'mã tài liệu:')
          {
            // add logo to a template file' s sheet containing the information of packaging specification
            addLogo(templateSheet, organization);

            var updatingSpecInformation = [standardName, specDate, version[0], version[1], requesterInformation[0][0], requesterInformation[0][2]];

            // update spec information to a template file' s sheet containing the information of packaging specification
            updateSpecInformation(templateSheet, templateData, templateHeader, updatingSpecInformation);

            var standardizedOrganization = standardizeOrganization(organization);
            templateSheet.getRange(findRowIndex(templateData, 'nơi áp dụng') + 1, 1).setValue(standardizedOrganization);
            
            templateSheet.getRange(findRowIndex(templateData, 'nơi áp dụng') + 1, findColumnIndex(templateHeader, 'tc')).setValue(standardName + '_M');
            
            changeTemplateFileInformation(templateSheet, templateData, version[0], specDate, note);
          }

          var subject = '[UPDATE-SPEC] Yêu Cầu Thay Đổi SPEC Đã Ban Hành ' + standardName + '_' + formattedSubmitTime;

          var body = "<div style=\"text-align:left;display: inline-block\">";
          body += "<p>" + "Dear anh/chị RD Packaging," + "</p>"; 
          body += "<p>" + "Team SPEC đã nhận được yêu cầu thay đổi thông tin SPEC đã ban hành " + standardName + " vào thời gian " + formattedSubmitTime + "</p>";
          body += "<p>" + "Anh/chị vui lòng truy cập vào đường dẫn đính kèm tại đây để kiểm tra thông tin: " + "<a class=\"standard-template-folder-link\", href="+"https://drive.google.com/drive/folders/"+requesterFolderId+">" + "Truy cập tại đây" + "</a>" + "</p>";

          var tableHeader = ["Thông tin yêu cầu", "Mã tiêu chuẩn", "Cấu trúc vật liệu thay đổi", "Thông tin thay đổi", "Tên tổ chức"];

          var tableData = ["Thay đổi thông tin SPEC đã ban hành", standardName, materialStructure, note, organization];

          sendSPECEmail(subject, body, tableHeader, tableData, requesterEmail, extractedPICEmail);

          updatingSpecSheet.getRange(i + 2, findColumnIndex(updatingSpecHeader, 'mailing_status')).setValue('Sent');
        
          if (updatingSpecSheet.getRange(i + 2, findColumnIndex(updatingSpecHeader, 'adding_to_control_list_status')).getValue() != 'Added') 
          {
            console.log('The index in the updating spec sheet: ' + i);

            for (var j = 0; j < packagingCode.length; ++j)
            {
              if (packagingCode[j][standardCodeIndex - 1] == '') 
              {
                console.log('The index in the packaging code sheet: ' + j);
              
                packagingCodeSheet.getRange(j + 2, standardCodeIndex).setValue(standardName);
                var document = specFileName.split('_')[1];
                packagingCodeSheet.getRange(j + 2, findColumnIndex(packagingCodeHeader, 'ten_tai_lieu')).setValue(document);
                packagingCodeSheet.getRange(j + 2, findColumnIndex(packagingCodeHeader, 'phien_ban')).setValue(version[0]);
                packagingCodeSheet.getRange(j + 2, findColumnIndex(packagingCodeHeader, 'ngay_hieu_luc')).setValue(formattedSubmitTime);
                packagingCodeSheet.getRange(j + 2, findColumnIndex(packagingCodeHeader, 'cau_truc_thay_doi')).setValue(materialStructure);	
                packagingCodeSheet.getRange(j + 2, findColumnIndex(packagingCodeHeader, 'status')).setValue('Updating');	
                packagingCodeSheet.getRange(j + 2, findColumnIndex(packagingCodeHeader, 'organization')).setValue(standardizedOrganization.replace("MSC và các nhà máy thành viên: ", "").trim());
                packagingCodeSheet.getRange(j + 2, findColumnIndex(packagingCodeHeader, 'requester_email')).setValue(requesterEmail);
                break;
              }
            }

            updatingSpecSheet.getRange(i + 2, findColumnIndex(updatingSpecHeader, 'adding_to_control_list_status')).setValue('Added');
            updatingSpecSheet.getRange(i + 2, findColumnIndex(updatingSpecHeader, 'spec_status')).setValue('Updating');
          }
        }
        else
        {
          break;
        }      
      }
    }
  }
}
