/**
 * This file contains application constants
 */
// open the folder titled 'Standard Template' by the its id
const STANDARD_TEMPLATE_FOLDER_ID = '1xyC1Kg4aZ9ZsuTRtH3yPfq_4XyKjNy7v';
const standardTemplateFolder = DriveApp.getFolderById(STANDARD_TEMPLATE_FOLDER_ID);

// open the folder named "Packaging" by its id
const PACKAGING_FOLDER_ID = '1dY7kIYfOGLS2YSNBCYgkanGzGVMEBeAm';
const packagingFolder = DriveApp.getFolderById(PACKAGING_FOLDER_ID);

// open the spec file by its id
const SPEC_ID = '1ekGR-D-STpubv9uBCQnqCeXWY0Cbd6_vsNdGifPHp9g';
const specFile = SpreadsheetApp.openById(SPEC_ID);

// open the file titled 'Packaging Code' by the its id
const PACKAGING_CODE_ID = '1WrkHXf2QqRzcU68JqB5-oZPcemQ-RvbGwwtAZGgJ5CY';
const packagingCodeFile = SpreadsheetApp.openById(PACKAGING_CODE_ID);
