const USING_INSTALLED_MOBILE_ONEDIT_TRIGGER = true;

function onEdit( event ) {

  var startTime = new Date();

  //find which sheet was edited and see if it's one of the tracked sheets
  let sheetEditedName = event.range.getSheet().getName();
  Logger.log( "ONEDIT: Checking Sheet Name: " + sheetEditedName );

  //check if it's the mobile single record sheet
  if( sheetEditedName == SHEETS.MOBILE_THIS_RECORD.NAME ) {
    let range = event.range;
    const row   = range.getRow(); 
    const col   = range.getColumn();
    const isSingleCell = ( ( range.getHeight() + range.getWidth() ) == 2 );
    mobile_thisRecord_onedit_( event, row, col, isSingleCell, sheetEditedName, range.getSheet(), null );
  }
  //check if it's the mobile main sheet
  if( sheetEditedName == SHEETS.MOBILE_MAIN.NAME ) {
    let range = event.range;
    onEditCheck_FromAlertSheetDef( range.getSheet(), SHEETS.MOBILE_MAIN.REFS.ALERT_BOX, event );
    if( !USING_INSTALLED_MOBILE_ONEDIT_TRIGGER ) {
      const row   = range.getRow(); 
      const col   = range.getColumn();
      const isSingleCell = ( ( range.getHeight() + range.getWidth() ) == 2 );
      mobile_main_onedit_( event, row, col, isSingleCell, sheetEditedName, range.getSheet(), null );
    }
    
  }
  //check if it's one of the the vital fields that trigger the pre-population rules on the input sheet
  //or mobile single record sheet
  else if( sheetEditedName == SHEETS.INPUT.NAME ) {
    Logger.log( "Starting Autocomplete Check" );
    let range = event.range;
    const row   = range.getRow(); const col   = range.getColumn();
    const isSingleCell = ( ( range.getHeight() + range.getWidth() ) == 2 );

    if( isSingleCell && col >= SHEETS.INPUT.REFS.COL_NO_RECORD_1 && col <= SHEETS.INPUT.REFS.COL_NO_RECORD_LAST && 
        ( row == SHEETS.INPUT.REFS.ROW_NO_LESSON_NAME ||
          row == SHEETS.INPUT.REFS.ROW_NO_LEARNER_NAME || row == SHEETS.INPUT.REFS.ROW_NO_ATTENDED ) ) {

      let sheetEdited = range.getSheet();
      let autoCompleteResults = trigger_getDataAndDefsForRecordInputAutocomplete( 
            event, row, col, isSingleCell, sheetEditedName, range.getSheet(), 0 );

      saveDataRangeValues( autoCompleteResults.DEFS, autoCompleteResults.DATA, sheetEdited, col, null, 0 );
      SpreadsheetApp.flush();
    }
  }

  //***  CHECK IF THE SHEET BEING EDITED IS TARGET GRADES SHEET OR MASTER FORM TEMPLATE SHEET AND IF SO
  //***  THEN MAINTAIN THE SHEETs LAST EDITED DATE' IN GLOBAL SETTINGS  
  if( sheetEditedName == SHEETS.MASTER_TEMPLATE.NAME || sheetEditedName == SHEETS.SETTINGS_LESSON_TARGETS.NAME ||
      sheetEditedName == SHEETS.MASTER_LEARNERS.NAME ) {

    //work out which row of global settings stores the last edited date for this sheet.
    var globalSettingsSheet = null, globalSettingsRowNum = null,  exitingDate = null, today = null;
    if( sheetEditedName == SHEETS.MASTER_TEMPLATE.NAME ) {
      globalSettingsRowNum = SHEETS.GLOBAL_SETTINGS.REFS.ROW_NO_LAST_TEMPLATE_EDIT;
    }
    if( sheetEditedName == SHEETS.MASTER_LEARNERS.NAME ) {
      globalSettingsRowNum = SHEETS.GLOBAL_SETTINGS.REFS.ROW_NO_LAST_MTR_LRNR_EDIT;
    }
    else if ( sheetEditedName == SHEETS.SETTINGS_LESSON_TARGETS.NAME ) {
      globalSettingsRowNum = SHEETS.GLOBAL_SETTINGS.REFS.ROW_NO_LAST_TARGETS_EDIT;
    }

    //get the date this sheet was last edited from the global settings
    globalSettingsSheet = SpreadsheetApp.getActive().getSheetByName( SHEETS.GLOBAL_SETTINGS.NAME );
    exitingDate = globalSettingsSheet.getRange( globalSettingsRowNum, SHEETS.GLOBAL_SETTINGS.REFS.COL_NO ).getValue();
    today = new Date(new Date().setHours(0,0,0,0));

    //if it was last edited before today then update the last edited datr to today
    if( today != exitingDate ) {
      globalSettingsSheet.getRange( globalSettingsRowNum, SHEETS.GLOBAL_SETTINGS.REFS.COL_NO ).setValue( today );
    }
  }

  //2-Way Record Date Sync
  if( sheetEditedName == SHEETS.INPUT.NAME && 
      event.range.getRow()    == SHEETS.INPUT.REFS.LESSON_DATE.ROW_NO &&
      event.range.getColumn() == SHEETS.INPUT.REFS.LESSON_DATE.COL_NO ) {
    setRecordDateFields(
          event.range.getSheet().getRange( 
                SHEETS.INPUT.REFS.LESSON_DATE.ROW_NO, 
                SHEETS.INPUT.REFS.LESSON_DATE.COL_NO 
          ).getValue(),
          null,
          SpreadsheetApp.getActive().getSheetByName( SHEETS.MOBILE_MAIN.NAME )
    );
  }
  if( sheetEditedName == SHEETS.MOBILE_MAIN.NAME && 
      event.range.getRow()    == SHEETS.MOBILE_MAIN.REFS.LESSON_DATE.ROW_NO &&
      event.range.getColumn() == SHEETS.MOBILE_MAIN.REFS.LESSON_DATE.COL_NO ) {
    setRecordDateFields(
          event.range.getSheet().getRange( 
                SHEETS.MOBILE_MAIN.REFS.LESSON_DATE.ROW_NO, 
                SHEETS.MOBILE_MAIN.REFS.LESSON_DATE.COL_NO 
          ).getValue(),
          SpreadsheetApp.getActive().getSheetByName( SHEETS.INPUT.NAME ),
          null
    );
  }

  //check if its a checkbox cell being updated as its ridiculously easy to delete checkboxes by accident in
  //google sheets at present
  renewCheckboxes( sheetEditedName, event );

  var endTime = new Date();
  Logger.log( "onEdit Finished in " + ( endTime.getTime() - startTime.getTime() ) + " ms" );
}