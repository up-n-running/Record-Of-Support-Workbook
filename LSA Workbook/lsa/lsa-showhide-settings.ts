function HideSettingsWorksheets() {
  var spreadsheet = SpreadsheetApp.getActive();
  var ui = SpreadsheetApp.getUi();
  var inputSheet = spreadsheet.setActiveSheet(spreadsheet.getSheetByName(SHEETS.INPUT.NAME), true);

  //get list of record statuses
  let numberOfRecords = SHEETS.INPUT.REFS.COL_NO_RECORD_LAST - SHEETS.INPUT.REFS.COL_NO_RECORD_1 + 1;
  var toPrintRange = inputSheet.getRange( SHEETS.INPUT.REFS.ROW_NO_STATUS_BAR, SHEETS.INPUT.REFS.COL_NO_RECORD_1, 1, numberOfRecords );
  var toPrintRangeValues = toPrintRange.getValues();

  spreadsheet.toast("Hiding the sheets that we don't need, please be patient", "Hiding Settings Sheets");
  
  spreadsheet.getSheetByName(SHEETS.SETTINGS_LEARNERS.NAME       ).hideSheet();
  spreadsheet.getSheetByName(SHEETS.SETTINGS_LESSONS.NAME        ).hideSheet();
  spreadsheet.getSheetByName(SHEETS.SETTINGS_LESSON_TARGETS.NAME ).hideSheet();
  spreadsheet.getSheetByName(SHEETS.SETTINGS_TARGET_GRADES.NAME  ).hideSheet();
  
  //loop through record of support tabs 1 to 25 hiding and unhiding as appropriate
  for ( let i = 0; i < numberOfRecords; i++) {
    if( toPrintRangeValues[0][i] == SHEETS.INPUT.STATUSES.UNSENT || 
        toPrintRangeValues[0][i] == SHEETS.INPUT.STATUSES.UNSIGNED || 
        toPrintRangeValues[0][i] == SHEETS.INPUT.STATUSES.SIGNED ) {

      if( spreadsheet.getSheetByName(""+(i+1)).isSheetHidden() ) {
        spreadsheet.getSheetByName(""+(i+1)).showSheet();
      }
    } else {
      if( !spreadsheet.getSheetByName(""+(i+1)).isSheetHidden() ) {
        spreadsheet.getSheetByName(""+(i+1)).hideSheet();
      }
    } 
  }
}

function ShowSettingsWorksheets() {
  var spreadsheet = SpreadsheetApp.getActive();
  var ui = SpreadsheetApp.getUi();
  var inputSheet = spreadsheet.setActiveSheet(spreadsheet.getSheetByName(SHEETS.INPUT.NAME), true);

  //get list of record statuses
  let numberOfRecords = SHEETS.INPUT.REFS.COL_NO_RECORD_LAST - SHEETS.INPUT.REFS.COL_NO_RECORD_1 + 1;
  var toPrintRange = inputSheet.getRange( SHEETS.INPUT.REFS.ROW_NO_STATUS_BAR, SHEETS.INPUT.REFS.COL_NO_RECORD_1, 1, numberOfRecords );
  var toPrintRangeValues = toPrintRange.getValues();

  spreadsheet.toast("Showing the sheets that we need, please be patient", "Showing Settings Sheets");
  
  spreadsheet.getSheetByName(SHEETS.SETTINGS_LEARNERS.NAME       ).showSheet();
  spreadsheet.getSheetByName(SHEETS.SETTINGS_LESSONS.NAME        ).showSheet();
  spreadsheet.getSheetByName(SHEETS.SETTINGS_LESSON_TARGETS.NAME ).showSheet();
  spreadsheet.getSheetByName(SHEETS.SETTINGS_TARGET_GRADES.NAME  ).showSheet();

  //loop through record of support tabs 1 to 25 hiding and unhiding as appropriate
  for ( let i = 0; i < numberOfRecords; i++) {
    if( toPrintRangeValues[0][i] == SHEETS.INPUT.STATUSES.UNSENT || 
        toPrintRangeValues[0][i] == SHEETS.INPUT.STATUSES.UNSIGNED || 
        toPrintRangeValues[0][i] == SHEETS.INPUT.STATUSES.SIGNED ) {

      if( spreadsheet.getSheetByName(""+(i+1)).isSheetHidden() ) {
        spreadsheet.getSheetByName(""+(i+1)).showSheet();
      }
    } else {
      if( !spreadsheet.getSheetByName(""+(i+1)).isSheetHidden() ) {
        spreadsheet.getSheetByName(""+(i+1)).hideSheet();
      }
    } 
  }
}
