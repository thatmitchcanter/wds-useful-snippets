/**
  PTO Google Calendar Sync
  Source: https://gist.github.com/williamsba/2d6dd4aed51cf81fef688f69598acb24
**/

function PTOApproved() {
/**
  Task 1) Open the Event Calendar
**/
var spreadsheet = SpreadsheetApp.getActiveSheet();
var calendarId = spreadsheet.getRange("G4").getValue();
var eventCal = CalendarApp.getCalendarById("c_5j8ik8td2798vss0vmtssqgckk@group.calendar.google.com");

/**
Task 2) Pull PTO information into the code, in a form that the code can understand.
**/
  
  var requests = spreadsheet.getRange("A7:C500").getValues();
  
 /**
 Task 3) Do the work
 **/
  for (x=0; x<requests.length; x++) {
    
    var shift = requests[x];
    
    var name = shift [0]
    var getAllDayStartDate = shift[1]
    var options = shift[2]
    
    eventCal.createAllDayEvent(name + " " + options, getAllDayStartDate);
    
    //pause for 1 seconds
    Utilities.sleep(500);
    
  }
}


function clearCalendar2(){
  
  var fromDate = new Date();

  //Get all events through 2024
  //The year needs to be set to the next calendar year for this to work
  var toDate = new Date(2024,0,0,0,0,0);
  var calendar = CalendarApp.getCalendarById("c_5j8ik8td2798vss0vmtssqgckk@group.calendar.google.com");
  var events = calendar.getEvents(fromDate, toDate);
  for(var i=0; i<events.length;i++){
    var ev = events[i];
    Logger.log(ev.getTitle()); // show event name in log
    ev.deleteEvent();
    
    //pause for .5 seconds
    Utilities.sleep(500);
    
  }
  
}
  

/** Task 4) Make it easy
**/
function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Sync to Calendar')
  .addItem('Sync PTO', 'PTOApproved')
      .addItem('Clear all events from Calendar', 'clearCalendar2')
      .addToUi();
   
}
