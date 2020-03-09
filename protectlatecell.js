/**
 * Creates a trigger for when a the spreadsheet is editted
 alert and protect if the time period is too close.
 */

function createProtectLateCellTrigger() {
  var spreadsheetId = '1Ff_oDkKvcajigONMEPpkonVVlhYVyUJJVOd-ZEEUMcw';
  ScriptApp.newTrigger('protectLateCell')
  .forSpreadsheet(spreadsheetId)
  .onOpen()
  .create();
}

function protectLateCell() {
  var hourLeftRange = "AE8:AY8";
  var recordingsRange = "E10:Y18";  
  createCellProtectionAndFormatColor(hourLeftRange, recordingsRange);
}

function createCellProtectionAndFormatColor(hourLeftRange, recordingsRange) {
  
  // Get the active sheet object
  var spreadsheet = SpreadsheetApp.getActiveSheet();

  // Hours left
  var hourLefts = spreadsheet.getRange(hourLeftRange).getValues()[0];
  var recordings = spreadsheet.getRange(recordingsRange).getValues();
  var recordingRange = spreadsheet.getRange(recordingsRange);
  recordingRange.setBackgroundColor("white");
  var firstRowPosition = recordingRange.getRow();
  var firstColumnPosition = recordingRange.getColumn();
  var numOfRows = recordingRange.getNumRows();
  
  // Loop through the shift starting times to get the last column to protect
  var numColumnProtect = 0;      
  for(column = 0; column<hourLefts.length; column++) {
    numColumnProtect = column;
    // Check whether the starting time is too close
    if(hourLefts[column] > 18) {      
      break;            
    }        
  }
  
  // Protect and format color the right range
  var rangeProtect = spreadsheet.getRange(firstRowPosition, firstColumnPosition, numOfRows, numColumnProtect);
  
  rangeProtect.setBackground("#757575");
        
  // Protect cell
  var protection = rangeProtect.protect();
  protection.setDescription("You choose this shift too late.")
        
  // Allow owner to edit
  var owner = SpreadsheetApp.getActiveSpreadsheet().getOwner();
  protection.addEditor(owner);
        
  // Prevent other collaborators from editting
  protection.removeEditors(protection.getEditors());
  if (protection.canDomainEdit()) {
    protection.setDomainEdit(false);
  }  
  
}