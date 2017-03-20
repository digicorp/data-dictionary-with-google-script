/**
 * To get sheets of the current active spreadsheet
 * @param spreadsheet Spreadsheet
 */
function CreatNewDocument()
{
  var sheet = GetActiveSheet("Metadata");
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var documentname = GetCurrentSpreadSheetName(spreadsheet)
  var document=DocumentApp.create(documentname);
  var documentid = document.getId();
  SetCellValue(sheet,1,1,documentid);
  return documentid;
}

/**
 * To get sheets of the current active spreadsheet
 * @param spreadsheet Spreadsheet
 */
function GetExistingDocument()
{
  var sheet = GetActiveSheet("Metadata");
  var documentid = GetCellValue(sheet,1,1);
  return documentid;
}





function GetDocumentFile()
{
    var documentid = GetExistingDocument();
    var documentfile = DriveApp.getFileById(documentid);
    return documentfile;
}


