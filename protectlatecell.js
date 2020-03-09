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

  if(numOfRows>0 && numColumnProtect>0) {
    // Protect and format color the right range
  var rangeProtect = spreadsheet.getRange(firstRowPosition, firstColumnPosition, numOfRows, numColumnProtect);
  rangeProtect.setBackground("#757575");

  // Protect cell
  var protection;
  var protectionDescription = "You choose this shift too late.";
  if(getProtectionByDescription() != null) {
    protection = getProtectionByDescription(protectionDescription);
    Logger.log("get the already created protection");
  } else {
    var protection = rangeProtect.protect();
    protection.setDescription(protectionDescription);
    Logger.log("create new protection.")
  }

  // Get the active sheet object
  var spreadsheet = SpreadsheetApp.getActiveSheet();
  var protections = spreadsheet.getProtections(SpreadsheetApp.ProtectionType.RANGE);
  Logger.log("protections.length = " + protections.length);

  // Allow owner to edit
  var owner = SpreadsheetApp.getActiveSpreadsheet().getOwner();
  protection.addEditor(owner);

  // Prevent other collaborators from editting
  protection.removeEditors(protection.getEditors());
  if (protection.canDomainEdit()) {
    protection.setDomainEdit(false);
  }
  }

}

function getProtectionByDescription(description) {
  var spreadsheet = SpreadsheetApp.getActiveSheet();
  var protections = spreadsheet.getProtections(SpreadsheetApp.ProtectionType.RANGE);
  for (var i = 0; i < protections.length; i++) {
    var protection = protections[i];
    if (protection.getDescription()) {
      return protection;
    } else {
      return null;
    }
  }
}


function deleteAllProtections() {
  // Get the active sheet object
  var spreadsheet = SpreadsheetApp.getActiveSheet();
  var protections = spreadsheet.getProtections(SpreadsheetApp.ProtectionType.RANGE);
  for (var i = 0; i < protections.length; i++) {
    var protection = protections[i];
    if (protection.canEdit()) {
      protection.remove();
    }
  }
}

function getAllProtections() {
  // Get the active sheet object
  var spreadsheet = SpreadsheetApp.getActiveSheet();
  var protections = spreadsheet.getProtections(SpreadsheetApp.ProtectionType.RANGE);
  Logger.log("protections.lenght = " + protections.length);
}


// Check if current user is the owner
function isOwner(e) {

  try {

    var currentUser = e.user;
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var owner = ss.getOwner();


    if(currentUser.getEmail() != owner.getEmail()) {
      // Different user logged in
      Logger.log("1: You are not the owner");
      return false;
    } else {
      // Owner logged in
      Logger.log("Hi owner");
      return true;
    }
  } catch(e) {
    //anonymous user
    Logger.log("2: You are not the owner");
    return false;
  }
}
