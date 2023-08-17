const id = "1urO9j9aBRBdKchiZ9WyuM-_24vA9xOznEdR25wvY5OY";
const sensoryFile = SpreadsheetApp.openById(id);

const responseSheet = sensoryFile.getSheetByName('Câu trả lời biểu mẫu 1');
const columnNameResponse = responseSheet.getRange(1, 1, 1, responseSheet.getLastColumn()).getValues()[0];
const response = responseSheet.getRange(2, 1, responseSheet.getLastRow() - 1, responseSheet.getLastColumn()).getValues();

const submitTimeIndex = findIndex(columnNameResponse, 'Dấu thời gian');
const employeeCodeIndex = findIndex(columnNameResponse, 'mã số nhân viên');
const categoryCodeIndex = findIndex(columnNameResponse, 'MÃ NGÀNH HÀNG');
const productNameIndex = findIndex(columnNameResponse, 'TÊN SẢN PHẨM');
const projectIDIndex = findIndex(columnNameResponse, 'project_id');
const itemCodeIndex = findIndex(columnNameResponse, 'item_code');
const formulationCodeIndex = findIndex(columnNameResponse, 'mã công thức');
const objectiveIndex = findIndex(columnNameResponse, 'MỤC ĐÍCH THÍ NGHIỆM');
const noteIndex = findIndex(columnNameResponse, 'lưu ý');
const requesterEmailIndex = findIndex(columnNameResponse, 'email');
const contextIndex = findIndex(columnNameResponse, 'BỐI CẢNH THÍ NGHIỆM');
const expectedDateIndex = findIndex(columnNameResponse, 'Ngày dự kiến');
const productionSiteIndex = findIndex(columnNameResponse, 'Địa điểm sản xuất');
const statusIndex = findIndex(columnNameResponse, 'Mailing');

const mappingSheet = sensoryFile.getSheetByName('mapping');
const columnNameMapping = mappingSheet.getRange(1, 1, 1, mappingSheet.getLastColumn()).getValues()[0];
const mapping = mappingSheet.getRange(2, 1, mappingSheet.getLastRow() - 1, mappingSheet.getLastColumn()).getValues().filter(r => r[0] != '');

const categoryCodeMappingIndex = findIndex(columnNameMapping, 'Category_code');
const emailIndex = findIndex(columnNameMapping, 'email');
