function AnnounceToAllLSAs()
{
  var spreadsheet = SpreadsheetApp.getActive();
  var ui = SpreadsheetApp.getUi();  
  var masterAnnounceSheet = spreadsheet.getSheetByName( SHEETS.MASTER_ANNOUNCEMENTS.NAME );

  var approvedText = PreviewAnnouncement( spreadsheet, masterAnnounceSheet );
  if( approvedText != null ) {
    PushToChildWorksheets( null, null, approvedText, null, false, false, false );
    ClearDownProposedAnnouncement( spreadsheet, masterAnnounceSheet, true );
  }
}

//returns announcementText if preview was approved and null if not
function PreviewAnnouncement( spreadsheet, masterAnnounceSheet )
{
  spreadsheet = ( spreadsheet ) ? spreadsheet : SpreadsheetApp.getActive();
  masterAnnounceSheet = ( masterAnnounceSheet ) ? masterAnnounceSheet : spreadsheet.getSheetByName( SHEETS.MASTER_ANNOUNCEMENTS.NAME );
  var ui = SpreadsheetApp.getUi();

  var pendingAnnouncement = masterAnnounceSheet.getRange(
      SHEETS.MASTER_ANNOUNCEMENTS.REFS.ROW_NO_PROPOSED_ANCMNT, 
      SHEETS.MASTER_ANNOUNCEMENTS.REFS.COL_NO_PROPOSED_ANCMNT 
    ).getValue();

  var alertBoxText = pendingAnnouncement + 
              "\n\n______________________________________________________________________\n\n" +
              "ARE YOU SURE YOU WANT TO SEND THIS ANNOUNCEMENT?"
Logger.log( pendingAnnouncement );
  if( pendingAnnouncement != "" ) {
    var announcementResponse = ui.alert ( "Preview Announcement...  SEND IT NOW?", alertBoxText, ui.ButtonSet.OK_CANCEL );
    if( announcementResponse == ui.Button.OK ) {
      return pendingAnnouncement;
    }
  }
  else {
    ui.alert ( "No announcement composed", 
        "Please go to sheet '" + SHEETS.MASTER_ANNOUNCEMENTS.NAME + "' and compose an announcement before attempting to send one.", 
        ui.ButtonSet.OK );
  }

  return null;
}

function ClearDownProposedAnnouncement( spreadsheet, masterAnnounceSheet, confirmFirst ) {
  spreadsheet = ( spreadsheet ) ? spreadsheet : SpreadsheetApp.getActive();
  masterAnnounceSheet = ( masterAnnounceSheet ) ? masterAnnounceSheet : spreadsheet.getSheetByName( SHEETS.MASTER_ANNOUNCEMENTS.NAME );
  confirmFirst = ( confirmFirst ) ? true : false;
  var actuallyRemove = !confirmFirst;

  if( confirmFirst != "" ) {
    var ui = SpreadsheetApp.getUi();
    var confirmResponse = ui.alert ( "Clear down your Announcement Text?", 
          "We recommend that your announcement text from the sheet '" + SHEETS.MASTER_ANNOUNCEMENTS.NAME + "' be removed,\n" +
          "that way no-one else can send your announcement again thinking it hadn't been sent already.\n\n" +
          "Click 'Yes' to clear down your announcement text, or click 'No' to keep it.",
          ui.ButtonSet.YES_NO );
    actuallyRemove = ( confirmResponse == ui.Button.YES );
  }

  if( actuallyRemove ) {
    masterAnnounceSheet.getRange(
        SHEETS.MASTER_ANNOUNCEMENTS.REFS.ROW_NO_PROPOSED_ANCMNT, 
        SHEETS.MASTER_ANNOUNCEMENTS.REFS.COL_NO_PROPOSED_ANCMNT 
      ).setValue("");
  }
}


function AddAnnouncementToTheQueue( announcementText, globalSettingsSheet, topRow ) {

  //parse 3rd param which might be missing - it is used to override the rownum at top (front) of queue as you might be in
  //a older version of the spreadsheet
  topRow = (topRow) ? topRow : SHEETS.GLOBAL_SETTINGS.REFS.ROW_NO_PENDING_ANNOUNCEMENT;

  //append some text to the start of the message so the recipient knows when it was sent.
  announcementText = "You were sent this message on " + Utilities.formatDate(new Date(), GLOBAL_CONSTANTS.TIMEZONE, "EEEE, d MMMM") + ":\n" +
                     "___________________\n\n" +
                     announcementText

  //start at the top announcement in global settings then go down until a black cell is found, and put the announcement in there
  var foundRow = findFirstEmptyCellInColumn( topRow, SHEETS.GLOBAL_SETTINGS.REFS.COL_NO, globalSettingsSheet, null );

  //if it couldnt find one then the queue is full so remove the oldest one from the top by moving all cells below it up one.
  //thus creating a space at the bottom (back of the queue) to make space for our announcement
  if( foundRow == null ) {
    deleteCellsOnRowAndMoveUp( topRow, [ SHEETS.GLOBAL_SETTINGS.REFS.COL_NO ], globalSettingsSheet, null );
    foundRow = globalSettingsSheet.getMaxRows();
  }
  
  //finally save our announcement in our first blank cell in the queue
  globalSettingsSheet.getRange( foundRow, SHEETS.GLOBAL_SETTINGS.REFS.COL_NO ).setValue( announcementText );
}
