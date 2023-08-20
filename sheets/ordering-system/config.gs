const orderFileID = '1_Te6gjg_MpJ1PFdfUoh_DfmbDLSdKo1UJijF1GPLJnY'; // declare an order file id
const orderFile = SpreadsheetApp.openById(orderFileID); // get an order file by utilizing an order file id

const responseSheet = orderFile.getSheetByName('form_response') // get a response sheet by using the getSheetByName function
const responseColumnHeader = responseSheet.getRange(1, 1, 1, responseSheet.getLastColumn()).getValues()[0]; // get response sheet's column headers
const response = responseSheet.getRange(2, 1, responseSheet.getLastRow() - 1, responseSheet.getLastColumn()).getValues(); // get response sheet's data
const responseLastRow = responseSheet.getLastRow() // get response sheet's last row

// define response sheet's column indexes
const submitTimeIndex = findIndex(responseColumnHeader, 'Timestamp');
const linkIndex = findIndex(responseColumnHeader, 'link');
const emailOrderIndex = findIndex(responseColumnHeader, 'Email');
const statusIndex = findIndex(responseColumnHeader, 'Tình trạng gửi');

const dataSheet = orderFile.getSheetByName('order_data'); // get a data sheet by using the getSheetByName function

const emailSheet = orderFile.getSheetByName('email'); // get a email sheet by using the getSheetByName function
const emailColumnHeader = emailSheet.getRange(1, 1, 1, emailSheet.getLastColumn()).getValues()[0]; // get email sheet's column headers
const email = emailSheet.getRange(2, 1, emailSheet.getLastRow() - 1, emailSheet.getLastColumn()).getValues(); // get email sheet's data

// define email sheet's column indexes
const emailIndex = findIndex(emailColumnHeader, 'Email');
const emailLineManagerIndex = findIndex(emailColumnHeader, 'Line Manager');

const formID = '1FAIpQLScjDaIlMEQrhY-ILNsirs0f1HJSH-hUEpg1zhWWeB_UXOXNGA'; // declare a form id
const formURL = 'https://docs.google.com/forms/d/e/' + formID + '/viewform';

const formFolderID = '1TD0EzDZvdRq5jMsb1y_smVfeH_ko8Acq'; // declare a form folder id id
const formFolder = DriveApp.getFolderById(formFolderID);

const fileFolderID = '1ea2vh-tJP0u7DWaCATlnKghQlGm5Us22'; // declare a file folder id
const fileFolder = DriveApp.getFolderById(fileFolderID);

const approvalFileID = '1c0ZJR8zLIvV5_5WuRI2zb0-MyQ1Z7Y5Gv2Hmy2tqtFE'; // declare an approval file id
const approvalFile = SpreadsheetApp.openById(approvalFileID);

const templateID = '1Q2e73nC2Fg6srBAmAML3aRr_QePVB2eGbepwXdugt7E';
const templateFile = SpreadsheetApp.openById(templateID);
const orderTemplateSheet = templateFile.getSheetByName('order_template');
const orderTemplateColumnHeader = orderTemplateSheet.getRange(1, 1, 1, orderTemplateSheet.getLastColumn()).getValues()[0];

// define order template sheet's column indexes
const itemNameIndex = findIndex(orderTemplateColumnHeader, 'Tên mặt hàng');
const quantityIndex = findIndex(orderTemplateColumnHeader, 'Số lượng');
const unitIndex = findIndex(orderTemplateColumnHeader, 'Đơn vị tính');
const idSchemeIndex = findIndex(orderTemplateColumnHeader, 'ID');
const categoryIndex = findIndex(orderTemplateColumnHeader, 'Ngành hàng');
const requestDateIndex = findIndex(orderTemplateColumnHeader, 'Ngày cần');
const noteIndex = findIndex(orderTemplateColumnHeader, 'Ghi chú');
