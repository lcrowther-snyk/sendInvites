var ss = SpreadsheetApp.getActiveSpreadsheet();
var workingCalendar ="XXXXX@EMAIL.COM"
var sheet = ss.getActiveSheet();
var rangeData = sheet.getDataRange();
var lastColumn = rangeData.getLastColumn();
var lastRow = rangeData.getLastRow();
var emailColumnNum = 3;
var startRow = 35;
var searchRange = sheet.getRange(startRow,2, lastRow-1,2);
var calendar = CalendarApp.getOwnedCalendarById(workingCalendar);
var sixteenHours = 16 * 60 * 60 * 1000; 
var thirtyMins =  30 * 60 * 1000; 
var fiveHours = 5 * 60 * 60 * 1000; 
var fortyEightHours = 48 * 60 * 60 * 1000; 
var reminderInviteDesc = "Reminder to host the SE weekly call this week - get content ready etc.."

function sendInvites() {
    for (var i = 1; i < lastRow ; i++){
      var emailAddr = searchRange.getCell(i,2).getValue();
      if(emailAddr == "End") break;
      var wedDate = new Date(searchRange.getCell(i,1).getValue());
      var mondayDate = new Date(wedDate.getTime()-fortyEightHours);
      var currentRow = (startRow + i)-1;
      Logger.log('about to handle row ' + currentRow);
      for(j=0;j<2;j++) {
        var date = new Date();
        if(j==0) date=mondayDate;
        if(j==1) date=wedDate
        date = new Date(date.getTime()+sixteenHours); //make it 4pm AEST
        var event = getEvent(date,reminderInviteDesc);
        if(event !=null) {
          event.deleteEvent(); //if event already exits then delete 
        }
        event = calendar.createEvent(reminderInviteDesc,date,new Date(date.getTime()+thirtyMins));
        event.addGuest(emailAddr);
      }
    };
}


function getEvent(date,title) {
   var events = calendar.getEventsForDay(date);
   for(var i=0;i<events.length;i++) {
      var tTitle = events[i].getTitle();
      if(tTitle == title) {
        return events[i];
      }
   }
}



