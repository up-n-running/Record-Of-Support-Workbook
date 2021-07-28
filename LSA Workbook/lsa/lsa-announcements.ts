function CheckForAccouncements( spreadsheet, globalSettingsSheet )
{
  Logger.log( "CheckForAccouncements Started" );  
  spreadsheet = ( spreadsheet ) ? spreadsheet : SpreadsheetApp.getActive();
  globalSettingsSheet = ( globalSettingsSheet ) ? globalSettingsSheet : spreadsheet.getSheetByName( SHEETS.GLOBAL_SETTINGS.NAME );

  if( isAMasterNotAChild(spreadsheet, globalSettingsSheet, true ) ) {
    return;
  }

  var alertBoxText = null;
  var queueOffset = 0;
  var pendingAnnouncement = globalSettingsSheet.getRange(
      SHEETS.GLOBAL_SETTINGS.REFS.ROW_NO_PENDING_ANNOUNCEMENT, 
      SHEETS.GLOBAL_SETTINGS.REFS.COL_NO 
  ).getValue();

  //keep checking for messages until there are none left to check
  while( pendingAnnouncement != "" ) {
    alertBoxText = pendingAnnouncement + 
              "\n------------------------------\n" +
              "Do you want to mark this message as 'read' so you will not see it again?"


    var ui = SpreadsheetApp.getUi();
    var announcementResponse = ui.alert ( "You have a new message...   mark it as 'read'?", alertBoxText, ui.ButtonSet.YES_NO );

    if( announcementResponse == ui.Button.YES ) {
      //if they mark as read we can delete it, so move queue up and keep pointer still
      deleteCellsOnRowAndMoveUp( SHEETS.GLOBAL_SETTINGS.REFS.ROW_NO_PENDING_ANNOUNCEMENT, 
                                 [ SHEETS.GLOBAL_SETTINGS.REFS.COL_NO ],
                                 globalSettingsSheet, null );
    }
    else {
      //if they dont mark as read we're keeping it, so keep queue still and move pointer down
      queueOffset++;
    }

    pendingAnnouncement = ( queueOffset >= SHEETS.GLOBAL_SETTINGS.REFS.ANNOUNCEMENT_QUEUE_LENGTH ) ? "" : globalSettingsSheet.getRange(
        SHEETS.GLOBAL_SETTINGS.REFS.ROW_NO_PENDING_ANNOUNCEMENT + queueOffset, 
        SHEETS.GLOBAL_SETTINGS.REFS.COL_NO 
    ).getValue();
  }
  Logger.log( "CheckForAccouncements Finished" ); 
}