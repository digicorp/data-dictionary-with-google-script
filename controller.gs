function onOpen() {
     var spreadsheet = SpreadsheetApp.getActive();
    var menuItems = [{
        name: 'Generate Dictionary',
        functionName: 'CreateDataDictionary'
    },{
        name: 'Email Document as PDF',
        functionName: 'EmailDataDictionary'
    }];
    spreadsheet.addMenu('Digicorp Data Dictionary', menuItems);
}


function CreateDataDictionary()
{
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var documentid = GetExistingDocument();
  if (documentid == '')
  {
      documentid = CreatNewDocument()
  }
  var body = DocumentApp.openById(documentid).getBody();
  var sheets = GetSheets(spreadsheet);
  var sheetscount = sheets.length;
  body.clear();
  
  for (count = 0; count<sheetscount; count++)
  {
    var sheetname = GetSheetName(sheets,count);
    if (sheetname != "Metadata")
    {
       var sheet = GetActiveSheet(sheetname);
       CreateTable(sheet,body);
    }    
  }
    
  Browser.msgBox("Data dictionary is generated.");
  var to =  Browser.inputBox("TO", "Please Enter Email Address", Browser.Buttons.OK_CANCEL);
  
  if (to!="cancel")
  {
       var sheet = GetActiveSheet("Metadata");
       SetCellValue(sheet,1,2,to);
       Browser.msgBox("Thank you. We will send you data dictionary on " + to + " in few minutes");
  }
 
}

function CreateTable(sheet,body) { 
    
    //Style for header text
    var headerStyle = {};
    headerStyle[DocumentApp.Attribute.BACKGROUND_COLOR] = '#336600';
    headerStyle[DocumentApp.Attribute.BOLD] = true;
    headerStyle[DocumentApp.Attribute.FOREGROUND_COLOR] = '#FFFFFF';
    
    //Style for the cells for primary keys
    var cellStylePK = {};
    cellStylePK[DocumentApp.Attribute.BOLD] = false;
    cellStylePK[DocumentApp.Attribute.FOREGROUND_COLOR] = '#000000';
    cellStylePK[DocumentApp.Attribute.BACKGROUND_COLOR] = '#ff9900';
    
    //Style for the cells for foriegn keys
    var cellStyleFK = {};
    cellStyleFK[DocumentApp.Attribute.BOLD] = false;
    cellStyleFK[DocumentApp.Attribute.FOREGROUND_COLOR] = '#000000';
    cellStyleFK[DocumentApp.Attribute.BACKGROUND_COLOR] = '#ffcc99';
    
    //Style for the cells other than header row
    var cellStyle = {};
    cellStyle[DocumentApp.Attribute.BOLD] = false;
    cellStyle[DocumentApp.Attribute.FOREGROUND_COLOR] = '#000000';
    
    var bold = {};
    bold[DocumentApp.Attribute.BOLD] = true;
    bold[DocumentApp.Attribute.ITALIC] = false;
  
    var italic = {};
    italic[DocumentApp.Attribute.BOLD] = false;
    italic[DocumentApp.Attribute.ITALIC] = true;
  
    var normal = {};
    normal[DocumentApp.Attribute.BOLD] = false;
    normal[DocumentApp.Attribute.ITALIC] = false;

    body.appendParagraph(GetCellValue(sheet,1,1)).setAttributes(bold);
    body.appendParagraph(GetCellValue(sheet,1,2)).setAttributes(italic);
    body.appendParagraph("");
        
    body.appendParagraph(GetCellValue(sheet,2,1)).setAttributes(bold);
    body.appendParagraph(GetCellValue(sheet,2,2)).setAttributes(italic);
    body.appendParagraph("");
    body.appendParagraph("Dictionary").setAttributes(bold);
    body.appendHorizontalRule();
  

    var table = body.appendTable();   
    for(var i=1; i<=sheet.getLastColumn(); i++){
    var tr = table.appendTableRow();
    
    for(var j=3; j<=6; j++){
      var td = tr.appendTableCell(GetCellValue(sheet,j,i));
      if(i == 1) td.setAttributes(headerStyle);
      else if (GetCellValue(sheet,7,i) =="YES") td.setAttributes(cellStylePK);
      else if (GetCellValue(sheet,8,i) =="YES") td.setAttributes(cellStyleFK);
      else td.setAttributes(cellStyle);
 
    }
  }
 
  body.appendPageBreak()
  body.appendParagraph("Sample Values");
  body.appendHorizontalRule()
    
  var values = body.appendTable(); 
  for(var j=1; j<=sheet.getLastColumn(); j++)
  {
      var tr = values.appendTableRow();
      var td = tr.appendTableCell(GetCellValue(sheet,3,j)).setAttributes(normal);
      
      for(var i=1; i<=sheet.getLastRow(); i++)
      {
      if (GetCellValue(sheet,i,1)=="Value")
      {
        td = tr.appendTableCell(GetCellValue(sheet,i,j)).setAttributes(normal);
      }
      }
  }
    body.appendPageBreak()
}


function EmailDataDictionary()
{
    var documentfile = GetDocumentFile();
    var content = "Please find attached PDF containing data dictionary in meaningful format"
    var sheet =  GetActiveSheet("Metadata");
    var toemailaddress = GetCellValue(sheet,1,2);
    MailApp.sendEmail(toemailaddress, documentfile.getName(),content,{ name: Session.getActiveUser(), attachments: [documentfile.getAs(MimeType.PDF)]});
    SetCellValue(sheet,1,2,'');
}
