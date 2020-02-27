function triggerNewSheet() {
  var id = "1Ff_oDkKvcajigONMEPpkonVVlhYVyUJJVOd-ZEEUMcw";  
  createSpreadsheetEditTrigger(id);
}


/**
 * Creates a trigger for when a spreadsheet opens.
 */
function createSpreadsheetEditTrigger(spreadsheetId) {
  
  /*
  Opens the spreadsheet with the given ID.
  A spreadsheet ID can be extracted from its URL.
  For example, the spreadsheet ID in the URL https://docs.google.com/spreadsheets/d/abc1234567/edit#gid=0 is "abc1234567". 
  */
  
     
  var ss = SpreadsheetApp.getActive();
  ScriptApp.newTrigger('createEvents')
      .forSpreadsheet(ss)
      .onEdit()
      .create();
  
  ScriptApp.newTrigger('myFunction')
  .forSpreadsheet(SpreadsheetApp.getActive())
  .onEdit()
  .create();
}

/**
 * Deletes a trigger.
 * @param {string} triggerId The Trigger ID.
 */
function deleteTrigger(triggerId) {
  // Loop over all triggers.
  var allTriggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < allTriggers.length; i++) {
    // If the current trigger is the correct one, delete it.
    if (allTriggers[i].getUniqueId() === triggerId) {
      ScriptApp.deleteTrigger(allTriggers[i]);
      break;
    }
  }
}

function createEvents() {

  // Step 1:
//  var spreadsheet = SpreadsheetApp.getActiveSheet();
  var sheetName = "Copy of [2020] [23-29/03]";
  var spreadsheet = SpreadsheetApp.getSheetByName(sheetname);
  var calendarId = 'qvplpq5euuqfupgfj6aqd0nnuk@group.calendar.google.com';
  var eventCal = CalendarApp.getCalendarById(calendarId);

  //Step 2: Collect data from Spread sheet
  var recordings = spreadsheet.getRange("E10:Y18").getValues();
  var volunteers = spreadsheet.getRange("B10:D12").getValues();
  var shifts = spreadsheet.getRange("E8:Y9").getValues();

  var eventRange = spreadsheet.getRange("AE10:AY18");
  var eventIds = spreadsheet.getRange("AE10:AY18").getValues();


  //Step 3: Detect whether or not there is an event at a specific time range

  //  var row = 2;
  //  var column = 2;

  for (row = 0; row < recordings.length; row++) {

    for (column = 0; column < recordings[0].length; column++) {
      var isSelected = recordings[row][column];

      // if a cell is selected

      if (isSelected) {

        // Colect data from the changed position

        var startTime = shifts[0][column];
        var endTime = shifts[1][column];

        var eventTitle = "";

        //Name
        var name;
        if (volunteers[row][0] != null && volunteers[row][0] != "") {
          name = volunteers[row][0];
        } else {
          name = "Không rõ tên";
        }

        //Book
        var book;
        if (volunteers[row][1] != null && volunteers[row][1] != "") {
          book = " - " + volunteers[row][1];
        } else {
          book = " - " + "Không rõ sách";
        }

        //Author
        var author;
        if (volunteers[row][2] != null && volunteers[row][2] != "") {
          author = " - " + volunteers[row][2];
        } else {
          author = " - " + "Không rõ tác giả";
        }

        eventTitle += name + book + author;

        // check if there is already an event

        var oldEvent = eventCal.getEventById(eventIds[row][column]);

        if (oldEvent != null) {

          // Edit the event
          oldEvent.setTitle(eventTitle);
          oldEvent.setTime(startTime, endTime);
          Logger.log('Edited an event');

        } else {
          // add new event
          var event = eventCal.createEvent(eventTitle, startTime, endTime);

          // edit event id
          eventIds[row][column] = event.getId();
          eventRange.setValues(eventIds);
          Logger.log('Added an event');

        }

      } else {

        // remove event if there is an event
        // check if there is already an event

        var oldEvent = eventCal.getEventById(eventIds[row][column]);

        if (oldEvent != null) {

          //Remove event
          oldEvent.deleteEvent();

          // Remove event id
          eventIds[row][column] = null;
          eventRange.setValues(eventIds);

          Logger.log('Remove an event');

        }

      }

    }

  }

}
