function trigger_getDataAndDefsForRecordInputAutocomplete( event, row: number, col: number, 
      isSingleCell: boolean, sheetEditedName: string, 
      recordSheetEdited: GoogleAppsScript.Spreadsheet.Sheet, rowOffset: number ) {

Logger.log( "trigger_getDataAndDefsForRecordInputAutocomplete called");
Logger.log( "row = '" + row + "'");
Logger.log( "col = '" + col + "'");
Logger.log( "isSingleCell = '" + isSingleCell + "'");
Logger.log( "sheetEditedName = '" + sheetEditedName + "'");
Logger.log( "recordSheetEdited = '" + recordSheetEdited + "'");
Logger.log( "rowOffset = '" + rowOffset + "'");

  let keyData: Array<Array<String>> = recordSheetEdited.getRange( 
        SHEETS.INPUT.REFS.ROW_NO_LESSON_NAME+rowOffset, col, 3, 1 
  ).getValues();
  const lesson   = keyData[0][0];
  const learner  = keyData[1][0];
  const attended = keyData[2][0].toUpperCase();

  let rangesValuesArray = new Array();
  let defaultDefinitionsSectionArray = new Array();

  if( lesson!="" && learner!="" && attended != "" ) {

    if( attended == "YES" ) {

      //tick checkboxes and default resources and other text
      if( row == SHEETS.INPUT.REFS.ROW_NO_ATTENDED + rowOffset ) {
        defaultDefinitionsSectionArray = SHEETS.INPUT.DEFAULTS_ATTENDANCE.YES;
        rangesValuesArray = getDataArraysFromDefaultsDef( 
              SHEETS.INPUT.REFS, defaultDefinitionsSectionArray );
      }
      else {
        //dont update all attended fields but do still override autosign comments as it's no longer the same record
        defaultDefinitionsSectionArray.push( {  ROW: SHEETS.INPUT.REFS.ROW_NO_AUTOSIGN_MANUALENTRY, HEIGHT: 1 } );
        rangesValuesArray = [ [ [ "" ] ] ]; //array of 2d arrays with just 1 2d array in it, consisting of just 1 row and 1 column!
      }

      //now pupulate the checkboxes and support strategy notes and resources used from lookups based on selected learner and lesson
      let spreadsheet = SpreadsheetApp.getActive();
      let learnerSheet = spreadsheet.getSheetByName( SHEETS.SETTINGS_LEARNERS.NAME );
      let lessonSheet = spreadsheet.getSheetByName( SHEETS.SETTINGS_LESSONS.NAME );
      
      //find the learner and the lesson
      let learnerRow = findInColumn( learnerSheet, learner, SHEETS.SETTINGS_LEARNERS.REFS.COL_NO_LEARNER_NAME, 
            SHEETS.SETTINGS_LEARNERS.REFS.ROW_NO_FIRST_LEARNER, SHEETS.SETTINGS_LEARNERS.REFS.ROW_NO_LAST_LEARNER );
      let lessonRow = findInColumn( lessonSheet, lesson, SHEETS.SETTINGS_LESSONS.REFS.COL_NO_LESSON_NAME_READONLY, 
            SHEETS.SETTINGS_LESSONS.REFS.ROW_NO_FIRST_LESSON, SHEETS.SETTINGS_LESSONS.REFS.ROW_NO_LAST_LESSON );
      if( learnerRow > 0 && lessonRow > 0 ) {

        Logger.log( "learnerRow = " + learnerRow );
        Logger.log( "lessonRow = " + learnerRow );

        //create section definition to store values in
        defaultDefinitionsSectionArray.push( { 
              ROW: SHEETS.INPUT.REFS.ROW_NO_SUPPORT_STRAT_FIRST, 
              HEIGHT: SHEETS.INPUT.REFS.ROW_NO_RESOURCES_USED - SHEETS.INPUT.REFS.ROW_NO_SUPPORT_STRAT_FIRST + 1 } );
        
        //now get the data in the 2d array format needed
        //get lesson info first
        let tempSectionValuesColumn = new Array();
        let tempSectionValuesRow = lessonSheet.getRange( 
            lessonRow, SHEETS.SETTINGS_LESSONS.REFS.COL_NO_EQUIPMENT_USED, 1, 
            SHEETS.SETTINGS_LESSONS.REFS.COL_NO_SUPPORT_STRAT_LAST - SHEETS.SETTINGS_LESSONS.REFS.COL_NO_EQUIPMENT_USED + 1 
        ).getValues();
        //append the learner data to the end of the first row
        tempSectionValuesRow[0] = tempSectionValuesRow[0].concat( learnerSheet.getRange( 
            learnerRow, SHEETS.SETTINGS_LEARNERS.REFS.COL_NO_SUPPORT_STRAT_FIRST, 1, 
            SHEETS.SETTINGS_LEARNERS.REFS.COL_NO_EXTRA_SUPPORT_TEXT - SHEETS.SETTINGS_LEARNERS.REFS.COL_NO_SUPPORT_STRAT_FIRST + 1 
        ).getValues()[0] );
        //now we've got all the data in one row and it happens to be in the right order except the first column 
        //(lesson equipment used) needs to be removed and moved to the end.
        //while were doing this we transpose from row to column!
        for( let i=1; i < tempSectionValuesRow[0].length; i++ ) {
          tempSectionValuesColumn.push( [ tempSectionValuesRow[0][i] ] )
        }
        tempSectionValuesColumn.push( [ tempSectionValuesRow[0][0] ] )

        rangesValuesArray.push( tempSectionValuesColumn );
      }
    }
    else {
      if( row == SHEETS.INPUT.REFS.ROW_NO_ATTENDED + rowOffset ) {
        defaultDefinitionsSectionArray = SHEETS.INPUT.DEFAULTS_ATTENDANCE.NO;
        rangesValuesArray = getDataArraysFromDefaultsDef( SHEETS.INPUT.REFS, defaultDefinitionsSectionArray );
      }
      else {
        //dont update all attended fields but do still override autosign comments as it's no longer the same record
        defaultDefinitionsSectionArray.push( {  ROW: SHEETS.INPUT.REFS.ROW_NO_AUTOSIGN_MANUALENTRY, HEIGHT: 1 } );
        rangesValuesArray = [ [ [ "" ] ] ]; //array of 2d arrays with just 1 2d array in it, consisting of just 1 row and 1 column!
      }
    }

    Logger.log( "DATA:" );
    Logger.log( rangesValuesArray );
    Logger.log( "DEFINITIONS:" );
    Logger.log( defaultDefinitionsSectionArray );
  }

  return {
    DATA: rangesValuesArray,
    DEFS: defaultDefinitionsSectionArray
  };
}

