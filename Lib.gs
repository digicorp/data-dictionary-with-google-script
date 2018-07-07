/**
 * To get the name of current active spreadsheet
 * @param spreadsheet Spreadsheet
 */
function GetCurrentSpreadSheetName(spreadsheet)
{
  return spreadsheet.getName();
}

/**
 * To get sheets of the current active spreadsheet
 * @param spreadsheet Spreadsheet
 */
function GetSheets(spreadsheet)
{
  return spreadsheet.getSheets();
}


/**
 * Get active sheet OR Get sheet by name
 * @param sheetname Name of the sheet.
 */
function GetActiveSheet(sheetname) {
    var sheet;

    if (sheetname) {
        sheet = SpreadsheetApp.getActive().getSheetByName(sheetname);
    } else {
        sheet = SpreadsheetApp.getActive().getActiveSheet();
    }
    return sheet;
}

/**
 * To get name of sheet
 * @param sheets Array of sheets.
 * @param index Index number for item of arrary 
 */
function GetSheetName(sheets,index)
{
    var sheetname;
    if (sheets.length>0)
    {
       sheetname = sheets[index].getName();
    }
    else
    {
      sheetname = ''
    }
    return sheetname;
}


/**
 * To set value of a cell
 * @param sheet Sheet contain cell on which value should set.
 * @param row Row on which value should set
 * @param column Column on which value should set
 * @param value Value which should set in cell
 */
function SetCellValue(sheet,row,column,value)
{
  sheet.getRange(row, column).setValue(value);
}


/**
 * To get value of a cell
 * @param sheet Sheet contain cell of which value should get.
 * @param row Row on which value should get
 * @param column Column on which value should get
 */
function GetCellValue(sheet,row,column)
{
  return sheet.getRange(row, column).getValue();
}


/**
 * Will clear sheet.
 * @param sheet Sheet that needs to be cleared
 */
function ClearSheet(sheet)
{
  sheet.clear({
        formatOnly: true,
        contentsOnly: true
    });
    sheet.clearNotes();
}







