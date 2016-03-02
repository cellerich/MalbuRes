function onBookingChange(){
  
  // Iterate trough entries and find a FormState = New, Edit or Delete
  var MySheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  Logger.log(MySheet.getLastRow());
  
  for (var row = 2; row<=MySheet.getLastRow(); row++){
    var bd;
    var formState = MySheet.getRange(row, 12).getValue();
    var typ = MySheet.getRange(row, 8).getValue();
    
    // select calendar from entry     
    var cn = "ReservationTest";    
    switch(typ) {      
      case "Flugzeug Seneca III - HB-LQY":         
        cn = "Seneca III - HB-LQY";       
        break;       
      case "Simulator FNTP II - Seneca III":         
        cn = "Elite FNTP-II";       
        break;       
      case "Simulator FNTP II - Beech":         
        cn = "Elite FNTP-II";       
        break;       
      default:     
    } 
    
    switch (formState) {
        
        
      case "New":         
      //CASE New        
        Logger.log("New Booking"); 
        updatePhoneCell(row);
        bd = getBookingDetails(row);
        // check if appropriate Calendar has no entry for selected date/time        
        if (checkConcurrentBooking(cn,bd)) {          
          // if concurrent booking, notify user trough email and do noting
          bd.mailImage = PropertiesService.getScriptProperties().getProperty("concurrentImage");          
          sendConcurrentNotification(bd);   
          // update Jotform submission FormState to "Denied"
          updateFormFormState(bd["Submission Id"],"Denied");
        }         
        else {          
          // update calendar with new booking entry           
          var eventId = createEvent(cn, bd);         
          // update JotForm submission FormState and EventId        
          updateFormFormState(bd["Submission ID"],"Edit");   
          updateFormEventId(bd["Submission ID"], eventId);
          // enter EventId into SpreadSheet   
          updateSheetEventId(row, eventId);          
          // send notification about successful booking to user        
          sendSuccessNotification(bd);           
          // if FI was booked, send notification about booking to FI          
          if (bd.Fluglehrer != "")             
            sendFINotification(bd);          
        }  
        // clear FormState in SpreadSheet
        clearSheetFormState(row);
        
        //send Event Log via email
        GmailApp.sendEmail("cello@cello.ch", "Reservation entry log", Logger.getLog());
        break;
       
      case "Edit":
        /*
        Logger.log("Edited Booking")
        bd = getBookingDetails(row);
        // check if appropriate Calendar has no entry for selected date/time     
        if (checkConcurrentBooking(cn,bd)) {         
          // if concurrent booking, notify user trough email and leave current booking    
          bd.mailImage = PropertiesService.getScriptProperties().getProperty("concurrentImage");          
          sendUpdateConcurrentNotification(bd);     
          // ToDo: -- ursprüngliche Daten aus Kalender holen und im SpreadSheet und Formular wieder einsetzen
        }        
        else {        
          // update calendar with changed booking entry     
          var eventId = updateEvent(cn, bd);          
          // update Jotform submission EventId
          updateFormEventId(bd["Submission ID"],eventId);
          // update EventId in Spreadsheet
          updateSheetEventId(row, eventId);
          // send notification about successful updated booking to user      
          sendUpdateSuccessNotification(bd);         
          // if FI was booked, send notification about updated booking to FI      
          if (bd.Fluglehrer != "")          
            sendFIUpdateNotification(bd); 
        }
        // clear FormState in Spreadsheet
        clearSheetFormState(row);
        */
        break;
        
      case "Delete":
        Logger.log("Deleted Booking")
        bd = getBookingDetails(row);
        
        // clear calendar booking   
        deleteEvent(cn, bd);      
        // update FormState and EventId in SpreadSheet       
        updateSheetFormState(row, "Deleted");   
        updateSheetEventId(row,"");
        // delete JotForm submission    
        clearJotForm(bd["Submission ID"]);    
        // send notification about canceled booking to user     
        bd.mailImage = PropertiesService.getScriptProperties().getProperty("cancelImage");
        sendCancelNotification(bd);      
        // if FI was booked, send notification about cancellation to FI       
        if (bd.Fluglehrer != "")          
          sendFICancelNotification(bd); 
        
        //send Event Log via email
        GmailApp.sendEmail("cello@cello.ch", "Reservation entry log", Logger.getLog());
        break;
        
      default:
        Logger.log("Unchanged Booking")
        break;
        
    }   
  }
}

function doGet(e) {   
  //check if submissionId is valid and set the FormState to delete and call onChange function
  
  // Iterate trough entries and find the correct submissionId 
  var mySS = SpreadsheetApp.openById(PropertiesService.getScriptProperties().getProperty("ssId"));
  SpreadsheetApp.setActiveSpreadsheet(mySS);
  var MySheet = mySS.getSheetByName("Submissions");
  SpreadsheetApp.setActiveSheet(MySheet);
  Logger.log("request to delete submissionId %s",e.parameter.submissionId);   
  
  var requestedRow = 0;
  var bd;
  for (var row = 2; row<=MySheet.getLastRow(); row++){   
    bd = getBookingDetails(row);
    if (bd["Submission ID"] == e.parameter.submissionId) {
      requestedRow = row;
      break;
    }
  }
  
  Logger.log("row, formState: ", requestedRow, bd.FormState);
  
  if ((requestedRow > 0) && (bd.FormState != "Deleted")) {
    updateSheetFormState(requestedRow,"Delete");
    bd.mailImage = PropertiesService.getScriptProperties().getProperty("cancelImage");
    onBookingChange();
    return getConfirmPage(bd, "cancelPage");
  }
  else {
    bd.mailImage = PropertiesService.getScriptProperties().getProperty("errorImage");
    return getConfirmPage(bd, "notfoundPage");
  }
  
}

function getConfirmPage(bd, page) {
  var t = HtmlService.createTemplateFromFile(page);   
  t.data = bd;
  return t.evaluate().setSandboxMode(HtmlService.SandboxMode.IFRAME);
}


function updateSheetFormState(row, value) {
  Logger.log("Formstate in row " + row + " updated");
  //get the active Sheet ´submissions`    
  var mySS = SpreadsheetApp.getActive();   
  var myS = mySS.getActiveSheet();
  var myCell = myS.getRange(row, 12);
  myCell.setValue(value);
  
}

function updatePhoneCell(row) {
  //update the phone Telefonnummer (5) to correctly format it
  var mySS = SpreadsheetApp.getActive();   
  var myS = mySS.getActiveSheet();
  //Telefonnummer
  var myCell = myS.getRange(row, 5);
  var myCV = myCell.getValue();
  if (myCV == "#ERROR!")
    myCell.setValue(myCell.getFormula().replace("=",""));
}

function updateDateCells(row) {
  //update the phone Telefonnummer (5) to correctly format it
  var mySS = SpreadsheetApp.getActive();   
  var myS = mySS.getActiveSheet();
  //Von 
  var myCell = myS.getRange(row, 9);
  Logger.log(myCell.getValue());
  Logger.log(myCell.getFormula());


//  myCell.setNumberFormat(numberFormat)
//  var myCV = myCell.getValue();
//  myCell.setValue(myCell.getFormula().replace("'",""));
}

function clearSheetFormState(row){
  updateSheetFormState(row,"");
}


function updateSheetEventId(row, eventId) {
  Logger.log("Eventid - " + eventId + " - inserted in row " + row);
  //get the active Sheet ´submissions`     
  var mySS = SpreadsheetApp.getActive();    
  var myS = mySS.getActiveSheet();
  var myCell = myS.getRange(row, 13)
  myCell.setValue(eventId);
  
}

function myTestFunction() {
  getFIdata("Armand Baccalá")
}

function updateFormFormState(submissionId, newState) {
  updateJotForm(submissionId, "submission[19_FormState]", newState);
}

function updateFormEventId(submissionId, eventId) {
  updateJotForm(submissionId, "submission[22_EventId]", eventId);
}

function updateJotForm(submissionId, field, value) {
  /*
  
  var apiUrl = "https://api.jotform.com/submission/"
  var apiKey = PropertiesService.getScriptProperties().getProperty("apiKey");
  var options = {'method': 'post'};     
           //    'payload': payload }; 
  
  var response = UrlFetchApp.fetch(apiUrl + submissionId + "?apiKey=" + apiKey + "&" + field + "=" + value, options);
  
  */
  
}

function clearJotForm(submissionId) {
  /*
  var apiUrl = "https://api.jotform.com/submission/"  
  var apiKey = PropertiesService.getScriptProperties().getProperty("apiKey");
   
  var options = {'method': 'delete'};    
  
  var response = UrlFetchApp.fetch(apiUrl + submissionId + "?apiKey=" + apiKey, options);  
  */
}

 function createEvent(calendarName, data) {   
   var myCal = CalendarApp.getCalendarsByName(calendarName)[0];   
   var myEvent = myCal.createEvent(data["Reservations Art"] + " - " + 
                                   data.Vorname + " " + 
                                   data.Nachname, 
                                   data.Von, 
                                   data.Bis, 
                                   {location:"Airport Bern", description:getEventDescription(data)});   

  Logger.log("Event created, Id:" + myEvent.getId());
  return myEvent.getId();
}  

function updateEvent(calendarName, data) {
  var myCal = CalendarApp.getCalendarsByName(calendarName)[0];
  var myE = myCal.getEventSeriesById(data.EventId);
  
  myE.deleteEventSeries();
  var myEvent = createEvent(calendarName, data);
  
  Logger.log("Event updated, Id:" + myEvent);
  return myEvent;
  
}

function deleteEvent(calendarName, data) {
  var myCal = CalendarApp.getCalendarsByName(calendarName)[0];
  var myE = myCal.getEventSeriesById(data.EventId);
  
  myE.deleteEventSeries();
  Logger.log("Event deleted, Id:" + data.EventId);
}

function getEventDescription(data) {   
  var desc = data.Typ + "\n";   
  if (data.Fluglehrer != "")     
    desc = desc + "Fluglehrer: " + data.Fluglehrer + "\n";   
  desc = desc + data.Bemerkungen;      
  
  return desc; 
}  

function getDate(dateString){
  if (dateString instanceof Date)
    return dateString;
  var parts = dateString.split(" ");
  var date = parts[0].split(".");
  var time = parts[1].split(":");
  
  return new Date(date[2], date[1] - 1, date[0], time[0], time[1]);
}


function getBookingDetails(RowNumber){
    //get the active Sheet ´submissions`
    var mySS = SpreadsheetApp.getActive();
    var myS = mySS.getActiveSheet();
    
    var columns = myS.getLastColumn() + 1;
  var data = new Object();    
  
  for (var i=1; i<columns; i++){
    data[myS.getRange(1, i).getValue()] = myS.getRange(RowNumber, i).getValue();
    }
  
  if(data.FormState != "Delete") {
    data.Von = getDate(data.Von);
    data.Bis = getDate(data.Bis);
  }
  
  data.submissionId = data["Submission ID"];
  data.editLink = "http://jotformpro.com/form.php?formID=43586374390968&sid=" + data.submissionId + "&mode=edit";
  data.cancelLink = PropertiesService.getScriptProperties().getProperty("cancelLink") + data.submissionId;
  data.reservationType = data["Reservations Art"];
  data.email = data["E-Mail"];
  data.phone = data.Telefonnummer;
  var FIdata = getFIdata(data.Fluglehrer)
  data.FIemail = FIdata.Email;
  data.FIphone = FIdata.Telefon;
  data.FIlastname = FIdata.Name;
  data.FIfirstname = FIdata.Vorname;
  data.hasFI = data.Fluglehrer != "";
  var Ddata = getDeviceData(data.Typ);
  data.mailImage = Ddata.image;
  data.deviceText = Ddata.text;
    
  
  return data;
    
}

function getDeviceData(typ){
  //ToDo Bilder für die verschiedenen Reservationen anzeigen
  if (typ.indexOf("Simulator") >= 0)
    return {image:PropertiesService.getScriptProperties().getProperty("simImage"), text:"Der Simulator"};
  if (typ.indexOf("LQY") >= 0)
    return {image:PropertiesService.getScriptProperties().getProperty("senecaImage"), text:"Die Seneca III"};
  return {image:"http://placehold.it/280x100", text:"Die Mietsache"};
}

function getFIdata(FIName){
 
  var myS = SpreadsheetApp.getActive().getSheetByName("FI");
  var row;
  for (row = 2; row<=myS.getLastRow(); row++) {
    if(myS.getRange(row,1).getValue() == FIName)
      break;
  }
  
  var data = new Object();
  for (var i=1; i<=myS.getLastColumn(); i++) {
    data[myS.getRange(1,i).getValue()] = myS.getRange(row,i).getValue();
  }
  
  return data;
}


function checkConcurrentBooking(calendarName, bookingDetails) {
  //if no EventId (new event) check if no other events are conflicting
  var myCal = CalendarApp.getCalendarsByName(calendarName)[0];
  var myEvents = myCal.getEvents(bookingDetails.Von, bookingDetails.Bis);
  
  if (bookingDetails.EventId == "")
    return myEvents.length > 0;
  else {
    if (myEvents.length > 1)
      return true;
    else {
      if (myEvents.length == 0)
        return false;
      else {
        if (myEvents[0].getId() == bookingDetails.EventId)
          return false;
        else
          return true;
      }
    }
  }
      
}

function sendSuccessNotification(bd){
  Logger.log("send Success E-mail");
  GmailApp.sendEmail(bd.email, "MALBUWIT Buchung", "Buchung bestätigt", {
    htmlBody: getConfirmPage(bd,"ReservationAcceptedIL").getContent(),
    name: "Malbuwit Booking Service",
    replyTo: "info@malbuwit.ch"})
}

function sendConcurrentNotification(bd){
  Logger.log("send Concurrent E-mail");
  GmailApp.sendEmail(bd.email, "MALBUWIT Buchung", "Buchung nicht möglich", {
    htmlBody: getConfirmPage(bd,"ReservationDeclinedIL").getContent(),
    name: "Malbuwit Booking Service",
    replyTo: "info@malbuwit.ch"})
}

function sendFINotification(bd){
  Logger.log("send FI Notification");
  GmailApp.sendEmail(bd.FIemail, "MALBUWIT Buchung", "Buchung durch Schüler erstellt", {
    htmlBody: getConfirmPage(bd,"FIBookingNotificationIL").getContent(),
    name: "Malbuwit Booking Service",
    replyTo: "info@malbuwit.ch"})

}

/*
function sendUpdateSuccessNotification(bd){
  Logger.log("send Update Success E-mail");
  GmailApp.sendEmail(bd.email, "MALBUWIT Buchung", "Buchungsänderung bestätigt", {
    htmlBody: getConfirmPage(bd,"updateSuccessEmail").getContent(),
    name: "Malbuwit Booking Service",
    replyTo: "info@malbuwit.ch"})

}

function sendFIUpdateNotification (bd){
  Logger.log("send FI Update Notification");
  GmailApp.sendEmail(bd.FIemail, "MALBUWIT Buchung", "Buchung durch Schüler geändert", {
    htmlBody: getConfirmPage(bd,"FIupdateEmail").getContent(),
    name: "Malbuwit Booking Service",
    replyTo: "info@malbuwit.ch"})

}

function sendUpdateConcurrentNotification(bd){
  Logger.log("send Update Concurrent E-mail");
  GmailApp.sendEmail(bd.email, "MALBUWIT Buchung", "Buchungsänderung nicht möglich", {
    htmlBody: getConfirmPage(bd,"updateConcurrentEmail").getContent(),
    name: "Malbuwit Booking Service",
    replyTo: "info@malbuwit.ch"})

}
*/

function sendCancelNotification(bd){
  Logger.log("send Cancel E-mail");
  GmailApp.sendEmail(bd.email, "MALBUWIT Buchung", "Buchung annuliert", {
    htmlBody: getConfirmPage(bd,"ReservationCanceledIL").getContent(),
    name: "Malbuwit Booking Service",
    replyTo: "info@malbuwit.ch"})

}

function sendFICancelNotification(bd){
  Logger.log("send FI Cancel Notification");
  GmailApp.sendEmail(bd.FIemail, "MALBUWIT Buchung", "Buchung durch Schüler geändert", {
    htmlBody: getConfirmPage(bd,"FICancelNotificationIL").getContent(),
    name: "Malbuwit Booking Service",
    replyTo: "info@malbuwit.ch"})

}

/*
  ToDo's
  - Test Delete Funktionen
  - Final code publizieren und /new entfernen
  
*/

