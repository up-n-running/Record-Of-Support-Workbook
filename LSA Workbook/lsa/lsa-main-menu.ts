function ExportAllToSendRoSs() {
  let ui = SpreadsheetApp.getUi();
  
  let spreadsheet   = SpreadsheetApp.getActive();
  let inputSheet    = spreadsheet.setActiveSheet(spreadsheet.getSheetByName(SHEETS.INPUT.NAME), true);
  let myFilesSheet  = spreadsheet.getSheetByName(SHEETS.MY_FILES.NAME);
  
  ExportToSendRoSsFromList(null, spreadsheet, inputSheet, myFilesSheet, ui );
}

function RefreshAllFileModifiedDates_InputSheet() {
  var spreadsheet = SpreadsheetApp.getActive();
  var ui = SpreadsheetApp.getUi();

  let fileIdsChecked = CheckForSignatures_InputSheet( spreadsheet, null, ui, null, null );

  //FEEDBACK ONCE FILES HAVE BEEN CHECKED
  ui.alert(
    'Finished Checking for Signatures', 
    'A total of ' + fileIdsChecked.length + ' file' + ( fileIdsChecked.length == 1 ? ' was' : 's were' ) + 
    ' checked for updates.', 
    ui.ButtonSet.OK);
}

function RefreshAllFileModifiedDates_MyFilesSheet() {
  var spreadsheet = SpreadsheetApp.getActive();
  var ui = SpreadsheetApp.getUi();

  let fileIdsChecked = CheckForSignatures_MyFilesSheet( spreadsheet, null, null, ui, null, null );

  //FEEDBACK ONCE FILES HAVE BEEN CHECKED
  ui.alert(
    'Finished Checking for Signatures', 
    'A total of ' + fileIdsChecked.length + ' file' + ( fileIdsChecked.length == 1 ? ' was' : 's were' ) + 
    ' checked for updates.', 
    ui.ButtonSet.OK);
}

function ResetInputData( spreadsheet?: GoogleAppsScript.Spreadsheet.Spreadsheet, ui?: any ) {

  //parse params
  spreadsheet = ( spreadsheet ) ? spreadsheet : SpreadsheetApp.getActive();
  ui = ( ui ) ? ui : SpreadsheetApp.getUi();

  if( !isAuthorised_( spreadsheet, ui, null, true, true, true, false ) ) { return false; }

  //are you sure dialogue?
  var areYouSureResponse = ui.alert("Are you sure?", "This will remove all of the record of support information from this Sheet\n so you can start a new day with a fresh sheet.\n\nAre you sure?", ui.ButtonSet.OK_CANCEL);
  
  if( areYouSureResponse == ui.Button.OK )
  { 
    let inputSheet = spreadsheet.getSheetByName( SHEETS.INPUT.NAME );
    let mobileMainSheet = spreadsheet.getSheetByName( SHEETS.MOBILE_MAIN.NAME );
    let mobileInputSheet = spreadsheet.getSheetByName( SHEETS.MOBILE_THIS_RECORD.NAME );

    if( !ui.FAKE_MODE ) { spreadsheet.setActiveSheet( inputSheet, true ); }
    clearDownRoSRecordsOnInputSheet( inputSheet );
    clearDownRoSRecordsOnMobileInputSheet( mobileInputSheet );
    populateDefaultsFromTimetable( spreadsheet, inputSheet, null, null );
    setRecordDateFields( null, inputSheet, mobileMainSheet );
    
    //leave on date cell to prompt user to double check date is correct
    if( !ui.FAKE_MODE ) {
      inputSheet.getRange( SHEETS.INPUT.REFS.LESSON_DATE.ROW_NO,SHEETS.INPUT.REFS.LESSON_DATE.COL_NO ).activate();
      spreadsheet.toast( 'Please wait while the colours update behind the scenes', 'Please wait', 3 );
    }
    
    SpreadsheetApp.flush();
  }
};

function clearDownRoSRecordsOnInputSheet( inputSheet )
{
  let defaultDefinitionsSectionArray = SHEETS.INPUT.DEFAULTS;
  let rangesValuesArray = getDataArraysFromDefaultsDef( SHEETS.INPUT.REFS, defaultDefinitionsSectionArray );
  saveDataRangeValues( defaultDefinitionsSectionArray, rangesValuesArray, inputSheet, 
                      SHEETS.INPUT.REFS.COL_NO_RECORD_1, SHEETS.INPUT.REFS.COL_NO_RECORD_LAST, 0 );

}

function clearDownRoSRecordsOnMobileInputSheet( mobileInputSheet )
{
  let defaultDefinitionsSectionArray = SHEETS.INPUT.MOBILE_COPY;
  let rangesValuesArray = getDataArraysFromDefaultsDef( SHEETS.INPUT.REFS, defaultDefinitionsSectionArray );
  saveDataRangeValues( defaultDefinitionsSectionArray, rangesValuesArray, mobileInputSheet, 
        SHEETS.MOBILE_THIS_RECORD.REFS.COL_NO_RECORD_DATA, null,
        SHEETS.MOBILE_THIS_RECORD.REFS.RECORD_DATA_OFFSET );
  saveDataRangeValues( defaultDefinitionsSectionArray, rangesValuesArray, mobileInputSheet, 
        SHEETS.MOBILE_THIS_RECORD.REFS.COL_NO_SNAPSHOT, null,
        SHEETS.MOBILE_THIS_RECORD.REFS.RECORD_DATA_OFFSET );

  //now set the record number drop down to empty string
  mobileInputSheet.getRange( 
        SHEETS.MOBILE_THIS_RECORD.REFS.ROW_NO_RECORD_NO, 
        SHEETS.MOBILE_THIS_RECORD.REFS.COL_NO_RECORD_DATA 
  ).setValue( "" ); 
  mobileInputSheet.getRange( 
        SHEETS.MOBILE_THIS_RECORD.REFS.ROW_NO_RECORD_NO, 
        SHEETS.MOBILE_THIS_RECORD.REFS.COL_NO_SNAPSHOT 
  ).setValue( "" ); 
}

function setRecordDateFields( dateString: string|null, 
      inputSheet: GoogleAppsScript.Spreadsheet.Sheet|null, mobileMainSheet: GoogleAppsScript.Spreadsheet.Sheet|null ) {

  //parse params (is either sheet is null then we leave it as null and dont update that sheet)
  dateString = ( dateString != null ) ? dateString :
        Utilities.formatDate( new Date( new Date().setHours(0,0,0,0) ), GLOBAL_CONSTANTS.TIMEZONE, "dd MMMM yyyy" );
  
  if( inputSheet != null ) {
    inputSheet.getRange( SHEETS.INPUT.REFS.LESSON_DATE.ROW_NO,SHEETS.INPUT.REFS.LESSON_DATE.COL_NO )
              .setValue( dateString );
  }
  if( mobileMainSheet != null ) {
    mobileMainSheet.getRange( SHEETS.MOBILE_MAIN.REFS.LESSON_DATE.ROW_NO,SHEETS.MOBILE_MAIN.REFS.LESSON_DATE.COL_NO )
                   .setValue( dateString );
  }
}