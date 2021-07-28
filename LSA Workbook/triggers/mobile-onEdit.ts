function mobile_thisRecord_onedit_( event, row: number, col: number, isSingleCell: boolean, 
      sheetEditedName: string, singleRecordSheet: GoogleAppsScript.Spreadsheet.Sheet, 
      inputSheet: GoogleAppsScript.Spreadsheet.Sheet ) {

  let rowOffset: number = SHEETS.MOBILE_THIS_RECORD.REFS.RECORD_DATA_OFFSET;

Logger.log( "mobile_thisRecord_onedit called");
Logger.log( "row = '" + row + "'");
Logger.log( "col = '" + col + "'");
Logger.log( "isSingleCell = '" + isSingleCell + "'");
Logger.log( "sheetEditedName = '" + sheetEditedName + "'");
Logger.log( "singleRecordSheet = '" + singleRecordSheet + "'");
Logger.log( "inputSheet = '" + inputSheet + "'");
Logger.log( "rowOffset = '" + rowOffset + "'");

  let recordNo: number = -1;

  //autocomplete based on lesson, learner and attended
  if( isSingleCell && col == SHEETS.MOBILE_THIS_RECORD.REFS.COL_NO_RECORD_DATA && 
    ( row == SHEETS.INPUT.REFS.ROW_NO_LESSON_NAME + rowOffset ||
      row == SHEETS.INPUT.REFS.ROW_NO_LEARNER_NAME + rowOffset || 
      row == SHEETS.INPUT.REFS.ROW_NO_ATTENDED + rowOffset ) ) {

    let autoCompleteResults = trigger_getDataAndDefsForRecordInputAutocomplete( 
          event, row, col, isSingleCell, sheetEditedName, singleRecordSheet, rowOffset );

    saveDataRangeValues( autoCompleteResults.DEFS, autoCompleteResults.DATA, singleRecordSheet, col, null, rowOffset );
    SpreadsheetApp.flush();
  }

  //save record to input sheet
  if( isSingleCell && col == SHEETS.MOBILE_THIS_RECORD.REFS.COL_NO_CHECKBOXES && 
      row == SHEETS.MOBILE_THIS_RECORD.REFS.ROW_NO_CBX_SAVE ) {

    Logger.log( "Save Checkbox Ticked" );
    let unsavedChangesStatus = readHasUnsavedChangesCell( singleRecordSheet )
    Logger.log( "unsavedChangesStatus = " + unsavedChangesStatus );

    if( unsavedChangesStatus != "" ) {

      //initialise values
      inputSheet = ( inputSheet ) ? inputSheet : SpreadsheetApp.getActive().getSheetByName( SHEETS.INPUT.NAME );
      let recordNo = readRecordNoCell( singleRecordSheet );
      Logger.log( "recordNo = '" + recordNo + "'" );

      //Get data from Mobile sheet
      let defaultDefinitionsSectionArray = SHEETS.INPUT.MOBILE_COPY;
      let rangesValuesArray = loadDataArraysFromSheet( SHEETS.INPUT.REFS, singleRecordSheet, 
              SHEETS.MOBILE_THIS_RECORD.REFS.COL_NO_RECORD_DATA, defaultDefinitionsSectionArray, rowOffset );
      
      //save values on Input sheet
      saveDataRangeValues( defaultDefinitionsSectionArray, rangesValuesArray, inputSheet, 
        SHEETS.INPUT.REFS.COL_NO_RECORD_1 + recordNo - 1, null, 0 );

      //save values on column next to mobile form - along with the record number on the top
      saveDataRangeValues( defaultDefinitionsSectionArray, rangesValuesArray, singleRecordSheet, 
            SHEETS.MOBILE_THIS_RECORD.REFS.COL_NO_SNAPSHOT, null, rowOffset );

    }

    //uncheck the checkbox now we're finished
    singleRecordSheet.getRange( 
          SHEETS.MOBILE_THIS_RECORD.REFS.ROW_NO_CBX_SAVE,
          SHEETS.MOBILE_THIS_RECORD.REFS.COL_NO_CHECKBOXES 
    ).setValue( false );

    SpreadsheetApp.flush();
  }

  try {
    //load record from input sheet
    if( isSingleCell && ( ( col == SHEETS.MOBILE_THIS_RECORD.REFS.COL_NO_RECORD_DATA && 
        row == SHEETS.MOBILE_THIS_RECORD.REFS.ROW_NO_RECORD_NO )  || 
        ( col == SHEETS.MOBILE_THIS_RECORD.REFS.COL_NO_CHECKBOXES && 
          row == SHEETS.MOBILE_THIS_RECORD.REFS.ROW_NO_CBX_RELOAD ) ) ) {

      //initialise values
      inputSheet = ( inputSheet ) ? inputSheet : SpreadsheetApp.getActive().getSheetByName( SHEETS.INPUT.NAME );
      let recordNo = readRecordNoCell( singleRecordSheet );
      Logger.log( "recordNo = '" + recordNo + "'" );

      //copy values from input sheet to Mobile Single Record Sheet
      let fullRecordDefinitionsSectionArray: Array<Array<Array<string>>> = SHEETS.INPUT.MOBILE_COPY;
      let defaultDefinitionsSectionArray = SHEETS.INPUT.MOBILE_COPY;
      let rangesValuesArray = loadDataArraysFromSheet( SHEETS.INPUT.REFS, inputSheet, 
            SHEETS.INPUT.REFS.COL_NO_RECORD_1 + recordNo - 1, defaultDefinitionsSectionArray, 0 );
      
      //save values on mobile form
      saveDataRangeValues( defaultDefinitionsSectionArray, rangesValuesArray, singleRecordSheet, 
            SHEETS.MOBILE_THIS_RECORD.REFS.COL_NO_RECORD_DATA, null, rowOffset );

      //save values on column next to mobile form - along with the record number on the top
      saveDataRangeValues( defaultDefinitionsSectionArray, rangesValuesArray, singleRecordSheet, 
            SHEETS.MOBILE_THIS_RECORD.REFS.COL_NO_SNAPSHOT, null, rowOffset );

      //clear out any error messages from previous load
      singleRecordSheet.getRange( 
            SHEETS.MOBILE_THIS_RECORD.REFS.ROW_NO_STATUS_OVERRIDE,
            SHEETS.MOBILE_THIS_RECORD.REFS.COL_NO_SNAPSHOT 
      ).setValue( "" );

      //mark the loading as finished
      if( ( col == SHEETS.MOBILE_THIS_RECORD.REFS.COL_NO_CHECKBOXES && 
        row == SHEETS.MOBILE_THIS_RECORD.REFS.ROW_NO_CBX_RELOAD ) ) {
        //untick the checkboxso we know its finished loading
        singleRecordSheet.getRange(
              SHEETS.MOBILE_THIS_RECORD.REFS.ROW_NO_CBX_RELOAD, 
              SHEETS.MOBILE_THIS_RECORD.REFS.COL_NO_CHECKBOXES 
        ).setValue( false );
      }
      else {
        //update the loaded record number shapshot hidden field so we know its finished loading
        singleRecordSheet.getRange(
              SHEETS.MOBILE_THIS_RECORD.REFS.ROW_NO_RECORD_NO, 
              SHEETS.MOBILE_THIS_RECORD.REFS.COL_NO_SNAPSHOT 
        ).setValue( recordNo );
      }
    }
    SpreadsheetApp.flush();
  }
  catch ( e ) {
    singleRecordSheet.getRange(
      SHEETS.MOBILE_THIS_RECORD.REFS.ROW_NO_STATUS_OVERRIDE, 
      SHEETS.MOBILE_THIS_RECORD.REFS.COL_NO_SNAPSHOT 
    ).setValue( "" + e + ( ( recordNo<0 ) ? "" : " Please go to the Input Sheet and fix the error on record " + 
          recordNo + " then reload record " + recordNo + " here." ) );
    SpreadsheetApp.flush(); 
    throw e;
  }
}

function readRecordNoCell( singleRecordSheet ): number {
  let recordNoString = singleRecordSheet.getRange( 
        SHEETS.MOBILE_THIS_RECORD.REFS.ROW_NO_RECORD_NO, 
        SHEETS.MOBILE_THIS_RECORD.REFS.COL_NO_RECORD_DATA 
  ).getValue();
  let recordNo: number = parseInt( recordNoString, 10 );
  //google sheets lets you clear the value to empty string annoyingly
  if (isNaN(recordNo)) {
    recordNoString = singleRecordSheet.getRange( 
          SHEETS.MOBILE_THIS_RECORD.REFS.ROW_NO_RECORD_NO, 
          SHEETS.MOBILE_THIS_RECORD.REFS.COL_NO_SNAPSHOT
    ).getValue();
    Logger.log( "Record No is NaN so using the snapshot one of '" + recordNoString + "'" );
    singleRecordSheet.getRange( 
      SHEETS.MOBILE_THIS_RECORD.REFS.ROW_NO_RECORD_NO, 
      SHEETS.MOBILE_THIS_RECORD.REFS.COL_NO_RECORD_DATA 
    ).setValue( recordNoString );
    recordNo = parseInt( recordNoString, 10 );
  }
  return recordNo;
}

function readHasUnsavedChangesCell( singleRecordSheet ): string {
  return singleRecordSheet.getRange( 
        SHEETS.MOBILE_THIS_RECORD.REFS.ROW_NO_CBX_SAVE, 
        SHEETS.MOBILE_THIS_RECORD.REFS.COL_NO_CHECKBOXES + 1 
  ).getValue();
}

function mobile_main_onedit_( event, row: number, col: number, isSingleCell: boolean, 
  sheetEditedName: string, mobileMainSheet: GoogleAppsScript.Spreadsheet.Sheet,
  inputSheet: GoogleAppsScript.Spreadsheet.Sheet ) {

Logger.log( "mobile_main_onedit called");
Logger.log( "row = '" + row + "'");
Logger.log( "col = '" + col + "'");
Logger.log( "isSingleCell = '" + isSingleCell + "'");
Logger.log( "sheetEditedName = '" + sheetEditedName + "'");
Logger.log( "mobileMainSheet = '" + mobileMainSheet + "'");
Logger.log( "inputSheet = '" + inputSheet + "'"); 

  let clearDownCheckboxAfterFinished = false;

  //autocomplete based on lesson, learner and attended
  if( isSingleCell && col == SHEETS.MOBILE_MAIN.REFS.COL_NO_CHECKBOXES && 
      row == SHEETS.MOBILE_MAIN.REFS.ROW_NO_GENERATE_ROS ) {

    if( justTicked( mobileMainSheet, row ) && getLockOnMainSheet_( mobileMainSheet, 1 ) ) {
      Mobile_SendSelectedRecords_DontEmail( mobileMainSheet );
      releaseLockOnMainSheet_( mobileMainSheet, 1 );
    }
    clearDownCheckboxAfterFinished = true;
  }

  if( isSingleCell && col == SHEETS.MOBILE_MAIN.REFS.COL_NO_CHECKBOXES && 
    row == SHEETS.MOBILE_MAIN.REFS.ROW_NO_NEW_DAY_CLEAN ) {

    if( justTicked( mobileMainSheet, row ) && getLockOnMainSheet_( mobileMainSheet, 1 ) ) {
      Mobile_ResetInputData( mobileMainSheet );
      releaseLockOnMainSheet_( mobileMainSheet, 1 );
    }
    clearDownCheckboxAfterFinished = true;
  }
  
  //check all / uncheck all checkbox
  if( isSingleCell && col == SHEETS.MOBILE_MAIN.REFS.COL_NO_CHECKBOXES && 
      row == SHEETS.MOBILE_MAIN.REFS.ROW_NO_UN_TICK_ALL ) {

    mobileMainSheet.getRange( 
          SHEETS.MOBILE_MAIN.REFS.ROW_NO_FIRST_RECORD, 
          SHEETS.MOBILE_MAIN.REFS.COL_NO_CHECKBOXES,
          SHEETS.MOBILE_MAIN.REFS.ROW_NO_LAST_RECORD - SHEETS.MOBILE_MAIN.REFS.ROW_NO_FIRST_RECORD + 1,
          1
    ).setValue( justTicked( mobileMainSheet, row ) );

    clearDownCheckboxAfterFinished = false;
  }

  if( clearDownCheckboxAfterFinished ) {
    mobileMainSheet.getRange(
      row, 
      col
    ).setValue( false );
  } 
}

function getLockOnMainSheet_( mobileMainSheet: GoogleAppsScript.Spreadsheet.Sheet, row: number ) {
  let lockCell = mobileMainSheet.getRange( row, SHEETS.MOBILE_MAIN.REFS.COL_NO_HIDDEN );
  let lockValue = lockCell.getValue();
  if( lockValue == "" || lockValue < ( (new Date()).getTime() - 1000*60*5 ) ) {
    lockCell.setValue( (new Date()).getTime() );
    return true;
  }
}

function releaseLockOnMainSheet_( mobileMainSheet: GoogleAppsScript.Spreadsheet.Sheet, row: number ) {
  let lockCell = mobileMainSheet.getRange( row, SHEETS.MOBILE_MAIN.REFS.COL_NO_HIDDEN );
  lockCell.setValue( "" );
}

function justTicked( mobileMainSheet: GoogleAppsScript.Spreadsheet.Sheet, row: number ) {
  return mobileMainSheet.getRange( row, SHEETS.MOBILE_MAIN.REFS.COL_NO_CHECKBOXES ).getValue();
}



function InstalledTrigger_MainWorkbook_OnEdit( event ) {

  var startTime = new Date();

  //find which sheet was edited and see if it's one of the tracked sheets
  let sheetEditedName = event.range.getSheet().getName();
  Logger.log( "INSTALLED ONEDIT (installedTrigger_MobileOnEdit): Checking Sheet Name: " + sheetEditedName );

  //check if it's the mobile main sheet
  if( sheetEditedName == SHEETS.MOBILE_MAIN.NAME ) {
    let range = event.range;
    const row   = range.getRow(); 
    const col   = range.getColumn();
    const isSingleCell = ( ( range.getHeight() + range.getWidth() ) == 2 );
    mobile_main_onedit_( event, row, col, isSingleCell, sheetEditedName, range.getSheet(), null );
  }
}