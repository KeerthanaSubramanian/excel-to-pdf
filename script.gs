
n onOpen() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var csvMenuEntries = [{name: "Generate PDF for first 15 days", functionName: "generatePDFFirst15Days"},
                        {name: "Generate PDF for last 15 days", functionName: "generatePDFLast15Days"},
                        {name: "Delete employee record with zero payment", functionName: "deleteEmployeeWithZeroPayment"}];
  ss.addMenu("Tasks", csvMenuEntries);
}

function deleteEmployeeWithZeroPayment() {
  var ss = SpreadsheetApp.getActive();
  var sheets = SpreadsheetApp.getActiveSpreadsheet().getSheets();
  for(var sheetNumber = 3; sheetNumber < sheets.length; sheetNumber++) {
    var sheet = sheets[sheetNumber];
    var J29Cell = Number(sheet.getRange(29, 10).getValue());
    var J61Cell = Number(sheet.getRange(61, 10).getValue());
    if(J29Cell + J61Cell <= 0) {
      var sheetToDeleted = ss.getSheetByName(sheet.getName()); 
      ss.deleteSheet(sheetToDeleted);
    }
  }  
}

function generatePDFFirst15Days() {    
  var documentName = "NYC_" + Utilities.formatDate(new Date(), "GMT+9:00", "yyyy_dd_MMMM") + "_UC_Payslip";
  var doc = DocumentApp.create(documentName);  
  var body = doc.getBody();
  var pointsInInch = 72;
  body.setPageHeight(5.83 * pointsInInch);  
  body.setPageWidth(8.27 * pointsInInch); 
  body.setMarginTop(15);
  body.setMarginBottom(0);
  body.setMarginLeft(0);
  body.setMarginRight(0);
  
  var tableStyle = {};
  tableStyle[DocumentApp.Attribute.FONT_SIZE] = 8;
  tableStyle[DocumentApp.Attribute.BOLD] = false;
  tableStyle[DocumentApp.Attribute.FONT_FAMILY] = 'Calibri';
  tableStyle[DocumentApp.Attribute.HORIZONTAL_ALIGNMENT] = DocumentApp.HorizontalAlignment.CENTER;
  tableStyle[DocumentApp.Attribute.VERTICAL_ALIGNMENT] = DocumentApp.VerticalAlignment.CENTER;
  
  var headerStyle = {};
  headerStyle[DocumentApp.Attribute.BOLD] = true;
  
  var sheets = SpreadsheetApp.getActiveSpreadsheet().getSheets();
  
  for(var sheetNumber = 3; sheetNumber < sheets.length; sheetNumber++){ 
    var sheet = sheets[sheetNumber];   
    var J29Cell = Number(sheet.getRange(29, 10).getValue());
    if(J29Cell > 0) {
      if(sheetNumber != 3)
      {
        body.appendParagraph("");
      }
      var table = body.appendTable();
      table.setAttributes(tableStyle);
      table.setBorderColor("#FFFFFF");      

      for(var row = 1; row < 32; row++) {
        if(row == 8|| row == 16 || row == 23 || row == 27)
          continue;
        var tr = table.appendTableRow();
        if(row == 7 || row == 15 || row == 22 || row == 26) 
          tr.setMinimumHeight(17);
        else
          tr.setMinimumHeight(10);
        for(var col = 1; col < 12; col++) {          
          var cellValue = sheet.getRange(row, col).getDisplayValue();            
          var cell = tr.appendTableCell(cellValue);   
          if(col == 1)
            cell.setWidth(65);
          else            
            cell.setWidth(55);
          
          cell.setPaddingBottom(0);
          cell.setPaddingTop(0);
          cell.setPaddingLeft(0);
          cell.setPaddingRight(0);
          cell.setAttributes(tableStyle);
          cell.getChild(0).asParagraph().setAlignment(DocumentApp.HorizontalAlignment.CENTER);
          if((row == 20 || row == 21 || row == 28 || row == 29) && col == 10) {
            cell.setAttributes(headerStyle);            
          }
        }
      }
      body.appendPageBreak();           
    }    
  }
  doc.saveAndClose();
  var docBlob = doc.getAs('application/pdf');
  docBlob.setName(doc.getName() + ".pdf");
  var file = DriveApp.createFile(docBlob);
}

function generatePDFLast15Days() {  
  var documentName = "NYC_" + Utilities.formatDate(new Date(), "GMT+9:00", "yyyy_dd_MMMM") + "_UC_Payslip";
  var doc = DocumentApp.create(documentName);  
  var body = doc.getBody();
  var pointsInInch = 72;
  body.setPageHeight(5.83 * pointsInInch);  
  body.setPageWidth(8.27 * pointsInInch); 
  body.setMarginTop(15);
  body.setMarginBottom(0);
  body.setMarginLeft(0);
  body.setMarginRight(0);
  
  var tableStyle = {};
  tableStyle[DocumentApp.Attribute.FONT_SIZE] = 8;
  tableStyle[DocumentApp.Attribute.BOLD] = false;
  tableStyle[DocumentApp.Attribute.FONT_FAMILY] = 'Calibri';
  tableStyle[DocumentApp.Attribute.HORIZONTAL_ALIGNMENT] = DocumentApp.HorizontalAlignment.CENTER;
  tableStyle[DocumentApp.Attribute.VERTICAL_ALIGNMENT] = DocumentApp.VerticalAlignment.CENTER;
  
  var headerStyle = {};
  headerStyle[DocumentApp.Attribute.BOLD] = true;
  
  var sheets = SpreadsheetApp.getActiveSpreadsheet().getSheets();
  
  for(var sheetNumber = 3; sheetNumber < sheets.length; sheetNumber++){ 
    var sheet = sheets[sheetNumber];
    var J61Cell = Number(sheet.getRange(61, 10).getValue());
    if(J61Cell > 0) {
      if(sheetNumber != 3)
      {
        body.appendParagraph("");
      }
      var table = body.appendTable();
      table.setAttributes(tableStyle);
      table.setBorderColor("#FFFFFF");      
      
      for(var row = 33; row < 64; row++) {          
        if(row == 40|| row == 48 || row == 55 || row == 59)
          continue;
        var tr = table.appendTableRow();
        if(row == 39 || row == 47 || row == 54 || row == 58) 
          tr.setMinimumHeight(17);
        else
          tr.setMinimumHeight(10);
        
        for(var col = 1; col < 12; col++) {
          var cellValue = sheet.getRange(row, col).getDisplayValue();            
          var cell = tr.appendTableCell(cellValue);   
          if(col == 1)
            cell.setWidth(65);
          else            
            cell.setWidth(55);
          cell.setPaddingBottom(0);
          cell.setPaddingTop(0);
          cell.setPaddingLeft(0);
          cell.setPaddingRight(0);
          cell.setAttributes(tableStyle);
          cell.getChild(0).asParagraph().setAlignment(DocumentApp.HorizontalAlignment.CENTER);
          if((row == 52 || row == 53 || row == 60 || row == 61) && col == 10) {
            cell.setAttributes(headerStyle);
          }
        }
      }
      body.appendPageBreak();
    }    
  }
  doc.saveAndClose();
  var docBlob = doc.getAs('application/pdf');
  docBlob.setName(doc.getName() + ".pdf");
  var file = DriveApp.createFile(docBlob);
}
