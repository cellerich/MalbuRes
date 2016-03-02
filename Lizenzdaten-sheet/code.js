function myTestFunction() {
  var sh = getLizenzCheckSheet();
  var ro = new Object();
  
 
  //for (var row = 2; row<=sh.getLastRow(); row++){
    for (var row = 2; row<=20; row++){
    //get the row object of the current row
    ro = getRowObject(sh, row);
    
    //check if we have to send any update notification
    if (ro.Check > 0) {

      //check if there is a notification to be sent "Check" > 0
      if (ro.Check > 0) {

        //check if the "mailCheck" is greater then 0 
        //means we have remaining days until we send another notice
        if (ro.mailCheck >= 0) {

          //its time to send another notification
          sendUpdateNotification(ro);
          
          //update "lastMail" cell with actual date so we do not send a mail again to early
          updateLastMailCell(sh, row);
          
        }
      }
    }
  }
}

function updateLastMailCell(sheet, row) {
  //get the column type object to identify the correct cells
  var ct = getLicenseCheckColumnTypes();
  
  updateCell(sheet, row, ct.lastMail, Utilities.formatDate(new Date(), "CET", "dd.MM.yyyy"));
}

function updateCell(sheet, row, column, value) {
  var mc = sheet.getRange(row, column);
  mc.setValue(value);
}


function getRowObject(sheet, row){
  //get the name value pairs of the requested row
  //column names are in row 1
  
  var numCol = sheet.getLastColumn() + 1;
  var rowObj = new Object();    
  
  for (var i=1; i<numCol; i++){
    rowObj[sheet.getRange(1, i).getValue()] = sheet.getRange(row, i).getValue();
    }
  
  return rowObj;
}

function getLizenzCheckSheet() {
  var mySS = SpreadsheetApp.openById("1OmNwva41km5n78c5hZ4DopE96Gr6DtEEobXeHS1fUy0");
  SpreadsheetApp.setActiveSpreadsheet(mySS);
  return mySS.getSheetByName("LizenzCheck");  
}



function getLicenseCheckColumnTypes () {
  //function to name the columns of the sheet
  //has to be adapted if the sheet format changes
  
  var CT = {
    LicNr: 1,
    Type: 2,
    Name: 3,
    Prename: 4,
    Check: 5,
    SEP: 6,
    MEP: 9,
    IR: 12,
    FI: 15,
    IRI: 18,
    MedI: 21,
    MedII: 24,
    LPenglish: 28,
    LPunlimited: 29,
    lastMail: 33,
    mailInterval: 34
  }
  return CT
}



function sendUpdateNotification(rowObject){
  Logger.log("send Upadte E-mail to %s %s", rowObject.Name, rowObject.Vorname);

  GmailApp.sendEmail(rowObject.eMail, "MALBUWIT Ausweisüberwachung", "Ausweisüberwachung: Aktualisierung ist fällig", {
    htmlBody: getConfirmPage(rowObject,"UpdateNotification").getContent(),
    name: "Malbuwit License Service",
    replyTo: "info@malbuwit.ch"})
    
    
}

function getConfirmPage(rowObject, page) {
  var t = HtmlService.createTemplateFromFile(page);   
  t.data = rowObject;
  return t.evaluate().setSandboxMode(HtmlService.SandboxMode.IFRAME);
}
