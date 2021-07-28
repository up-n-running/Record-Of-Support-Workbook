function populateDefaultsFromTimetable( spreadsheet: GoogleAppsScript.Spreadsheet.Spreadsheet , 
                                        inputSheet: GoogleAppsScript.Spreadsheet.Sheet, 
                                        timetableSheet: GoogleAppsScript.Spreadsheet.Sheet,
                                        lessonDate: Date )
{
  //parse params
  spreadsheet    = spreadsheet    ? spreadsheet    : SpreadsheetApp.getActive();
  inputSheet     = inputSheet     ? inputSheet     : spreadsheet.getSheetByName( SHEETS.INPUT.NAME );
  timetableSheet = timetableSheet ? timetableSheet : spreadsheet.getSheetByName( SHEETS.TIMETABLE.NAME );
  lessonDate     = lessonDate     ? lessonDate     : new Date();

  //create filter for timetable days so we get just the day (of week) that we want
  let dayNameFilter = dayNumberToName( lessonDate.getDay() + 1, 1 );
  Logger.log( "dayNameFilter = '" + dayNameFilter + "'" );

  //get timetable rows as 2d array with blank ones filtered, ordered by day and time
  let allTimeTableRowsRawData: Array<Array<string>> = getFilteredAndOrderedTimeTableData_( spreadsheet, inputSheet, timetableSheet, dayNameFilter, true );

  //start to build the data to add to the input tab. start with 4 data arrays for the 4 rows we will populate
  let noOfRecords = SHEETS.INPUT.REFS.COL_NO_RECORD_LAST - SHEETS.INPUT.REFS.COL_NO_RECORD_1 + 1
  let rowLessonTimeData     = create2DPrePopulatedArray( 1, noOfRecords, "" );
  let rowLessonDurationData = create2DPrePopulatedArray( 1, noOfRecords, "" );
  let rowLessonNameData     = create2DPrePopulatedArray( 1, noOfRecords, "" );
  let rowLearnerNameData    = create2DPrePopulatedArray( 1, noOfRecords, "" );

  //loop through each of the timetable rows and copy data to input sheet data row arrays
  let noOfTimetableRows: number = allTimeTableRowsRawData.length;
  let timeTableRowData: Array<string>|null = null;
  let learnerName: string = "";
  let currentInputColumnIdx = 0;
  for( let ttRowIdx: number = 0; ttRowIdx < noOfTimetableRows && currentInputColumnIdx < noOfRecords; ttRowIdx++ ) {
    timeTableRowData = allTimeTableRowsRawData[ttRowIdx];
    //now loop through each of the learners from the row
    for( let learnerIdx: number = SHEETS.TIMETABLE.REFS.COL_NO_LEARNER_FIRST - 1; 
        learnerIdx < SHEETS.TIMETABLE.REFS.COL_NO_LEARNER_LAST && currentInputColumnIdx < noOfRecords; 
        learnerIdx++ ) {
      learnerName = timeTableRowData[ learnerIdx ];
      if( learnerName != "" ) {
        //populate Input Sheet Column Data
        rowLessonTimeData[0][ currentInputColumnIdx ]     = timeTableRowData[ SHEETS.TIMETABLE.REFS.COL_NO_TIME-1 ];
        rowLessonDurationData[0][ currentInputColumnIdx ] = timeTableRowData[ SHEETS.TIMETABLE.REFS.COL_NO_DURATION-1 ];
        rowLessonNameData[0][ currentInputColumnIdx ]     = timeTableRowData[ SHEETS.TIMETABLE.REFS.COL_NO_LESSON-1 ];
        rowLearnerNameData[0][ currentInputColumnIdx ]    = learnerName;
        currentInputColumnIdx++;
      }
    }
  }

  /*
  let defaultDefinitionsSectionArray = SHEETS.INPUT.DEFAULTS;
  let rangesValuesArray = getDataArraysFromDefaultsDef( SHEETS.INPUT.REFS, defaultDefinitionsSectionArray );
  saveDataRangeValues( defaultDefinitionsSectionArray, rangesValuesArray, inputSheet, 
                      SHEETS.INPUT.REFS.COL_NO_RECORD_1, SHEETS.INPUT.REFS.COL_NO_RECORD_LAST, 0 );
  */
  Logger.log( "SORTED TIMETABLE:" );
  Logger.log( allTimeTableRowsRawData );

  Logger.log( "INPUT DATA" );
  Logger.log( rowLessonTimeData );
  Logger.log( rowLessonDurationData );
  Logger.log( rowLessonNameData );
  Logger.log( rowLearnerNameData );

  //save the input data back to the input sheet
  inputSheet.getRange( SHEETS.INPUT.REFS.ROW_NO_START_TIME, SHEETS.INPUT.REFS.COL_NO_RECORD_1, 1, noOfRecords )
            .setValues( rowLessonTimeData );
  inputSheet.getRange( SHEETS.INPUT.REFS.ROW_NO_DURATION, SHEETS.INPUT.REFS.COL_NO_RECORD_1, 1, noOfRecords )
            .setValues( rowLessonDurationData );
  inputSheet.getRange( SHEETS.INPUT.REFS.ROW_NO_LESSON_NAME, SHEETS.INPUT.REFS.COL_NO_RECORD_1, 1, noOfRecords )
            .setValues( rowLessonNameData );
  inputSheet.getRange( SHEETS.INPUT.REFS.ROW_NO_LEARNER_NAME, SHEETS.INPUT.REFS.COL_NO_RECORD_1, 1, noOfRecords )
            .setValues( rowLearnerNameData );

  SpreadsheetApp.flush();

  //store necessary onedit changes one by one
  let autoCompleteResultsByColumnArray = new Array();
  for( let iCol: number = SHEETS.INPUT.REFS.COL_NO_RECORD_1; iCol <= SHEETS.INPUT.REFS.COL_NO_RECORD_LAST; iCol++ ) {
    autoCompleteResultsByColumnArray.push( trigger_getDataAndDefsForRecordInputAutocomplete( 
          null, SHEETS.INPUT.REFS.ROW_NO_LEARNER_NAME, iCol, true, SHEETS.INPUT.NAME, inputSheet, 0 
    ) );
  }

  //save sotred onedit changes one by one
  let autoCompleteResults = null;
  for( let iCol: number = SHEETS.INPUT.REFS.COL_NO_RECORD_1; iCol <= SHEETS.INPUT.REFS.COL_NO_RECORD_LAST; iCol++ ) {
    autoCompleteResults = autoCompleteResultsByColumnArray[ iCol - SHEETS.INPUT.REFS.COL_NO_RECORD_1 ];
    saveDataRangeValues( autoCompleteResults.DEFS, autoCompleteResults.DATA, inputSheet, iCol, iCol, 0 );
  }

  SpreadsheetApp.flush();
}


function SortTimetableRecords( spreadsheet: GoogleAppsScript.Spreadsheet.Spreadsheet , 
                               inputSheet: GoogleAppsScript.Spreadsheet.Sheet, 
                               timetableSheet: GoogleAppsScript.Spreadsheet.Sheet,
                               dayNameFilter: string|null )
{

  //parse params
  spreadsheet    = spreadsheet    ? spreadsheet    : SpreadsheetApp.getActive();
  inputSheet     = inputSheet     ? inputSheet     : spreadsheet.getSheetByName( SHEETS.INPUT.NAME );
  timetableSheet = timetableSheet ? timetableSheet : spreadsheet.getSheetByName( SHEETS.TIMETABLE.NAME );
  dayNameFilter  = dayNameFilter  ? dayNameFilter  : null;

  //get timetable rows as 2d array with blank ones filtered, ordered by day and time
  let allTimeTableRowsRawData: Array<Array<string>> = getFilteredAndOrderedTimeTableData_( spreadsheet, inputSheet, timetableSheet, dayNameFilter, false );

  //clear rows then
  let timetableDataRange: GoogleAppsScript.Spreadsheet.Range = timetableSheet.getRange( 
    SHEETS.TIMETABLE.REFS.ROW_NO_FIRST_ENTRY,
    SHEETS.TIMETABLE.REFS.COL_NO_DAY_OF_WEEK,
    SHEETS.TIMETABLE.REFS.ROW_NO_LAST_ENTRY - SHEETS.TIMETABLE.REFS.ROW_NO_FIRST_ENTRY + 1,
    SHEETS.TIMETABLE.REFS.COL_NO_LEARNER_LAST - SHEETS.TIMETABLE.REFS.COL_NO_DAY_OF_WEEK + 1
  ).clearContent();

  //write reordered rows over top
  timetableSheet.getRange( 
        SHEETS.TIMETABLE.REFS.ROW_NO_FIRST_ENTRY,
        SHEETS.TIMETABLE.REFS.COL_NO_DAY_OF_WEEK,
        allTimeTableRowsRawData.length,
        SHEETS.TIMETABLE.REFS.COL_NO_LEARNER_LAST - SHEETS.TIMETABLE.REFS.COL_NO_DAY_OF_WEEK + 1
  ).setValues( allTimeTableRowsRawData );

  //filter out incomplete/blank rows
  allTimeTableRowsRawData = allTimeTableRowsRawData.filter( 
    function( row ) {
      return row[SHEETS.TIMETABLE.REFS.COL_NO_DAY_OF_WEEK-1] != "" &&
            row[SHEETS.TIMETABLE.REFS.COL_NO_TIME-1] != "" &&
            row[SHEETS.TIMETABLE.REFS.COL_NO_DURATION-1] != "" &&
            row[SHEETS.TIMETABLE.REFS.COL_NO_LESSON-1] != "" &&
            ( !dayNameFilter || row[SHEETS.TIMETABLE.REFS.COL_NO_DAY_OF_WEEK-1] == dayNameFilter );
    }
  );

  return allTimeTableRowsRawData;
}

function getFilteredAndOrderedTimeTableData_( spreadsheet: GoogleAppsScript.Spreadsheet.Spreadsheet , 
                                             inputSheet: GoogleAppsScript.Spreadsheet.Sheet, 
                                             timetableSheet: GoogleAppsScript.Spreadsheet.Sheet,
                                             dayNameFilter: string|null,
                                             onlyCompleteRecordsFilter: boolean )
{

  //parse params
  spreadsheet    = spreadsheet    ? spreadsheet    : SpreadsheetApp.getActive();
  inputSheet     = inputSheet     ? inputSheet     : spreadsheet.getSheetByName( SHEETS.INPUT.NAME );
  timetableSheet = timetableSheet ? timetableSheet : spreadsheet.getSheetByName( SHEETS.TIMETABLE.NAME );
  dayNameFilter  = dayNameFilter  ? dayNameFilter  : null;

  //get all timetable rows
  let allTimeTableRowsRawData: Array<Array<string>> = timetableSheet.getRange( 
        SHEETS.TIMETABLE.REFS.ROW_NO_FIRST_ENTRY,
        SHEETS.TIMETABLE.REFS.COL_NO_DAY_OF_WEEK,
        SHEETS.TIMETABLE.REFS.ROW_NO_LAST_ENTRY - SHEETS.TIMETABLE.REFS.ROW_NO_FIRST_ENTRY + 1,
        SHEETS.TIMETABLE.REFS.COL_NO_LEARNER_LAST - SHEETS.TIMETABLE.REFS.COL_NO_DAY_OF_WEEK + 1
  ).getValues();

  //filter out incomplete/blank rows
  allTimeTableRowsRawData = allTimeTableRowsRawData.filter( 
    function( row ) {
      return ( !onlyCompleteRecordsFilter || ( 
                 row[SHEETS.TIMETABLE.REFS.COL_NO_DAY_OF_WEEK-1] != "" &&
                 row[SHEETS.TIMETABLE.REFS.COL_NO_TIME-1] != "" &&
                 row[SHEETS.TIMETABLE.REFS.COL_NO_DURATION-1] != "" &&
                 row[SHEETS.TIMETABLE.REFS.COL_NO_LESSON-1] != "" ) 
             ) &&
             ( !dayNameFilter || row[SHEETS.TIMETABLE.REFS.COL_NO_DAY_OF_WEEK-1] == dayNameFilter );
    }
  );

  //order timetable rows by lesson day (of week) and time
  allTimeTableRowsRawData.sort( 
    function( rowA, rowB ) {

      //compare days
      let dayNumberA = dayNameToNumber( rowA[SHEETS.TIMETABLE.REFS.COL_NO_DAY_OF_WEEK-1] );
      let dayNumberB = dayNameToNumber( rowB[SHEETS.TIMETABLE.REFS.COL_NO_DAY_OF_WEEK-1] );
      dayNumberA = dayNumberA < 0 ? 10 : dayNumberA; //if value was not filled in on spreadsheet
      dayNumberB = dayNumberB < 0 ? 10 : dayNumberB;       
      if( dayNumberA != dayNumberB ) {
        return dayNumberA - dayNumberB;
      }

      //if days the dame compare times
      let baseDate = new Date(0);
      let timeStringA = rowA[SHEETS.TIMETABLE.REFS.COL_NO_TIME-1];
      let timeStringB = rowB[SHEETS.TIMETABLE.REFS.COL_NO_TIME-1];
      timeStringA = timeStringA == "" ? "11:59 P" : timeStringA; //if value was not filled in on spreadsheet
      timeStringB = timeStringB == "" ? "11:59 P" : timeStringB;      
      let rowALessonTimeObj = getlessonTimesObjectFromInputSheetValues( baseDate, timeStringA, "0m" );
      let rowBLessonTimeObj = getlessonTimesObjectFromInputSheetValues( baseDate, timeStringB, "0m" );
      let timeDiff = rowALessonTimeObj.start.getTime() - rowBLessonTimeObj.start.getTime();
      return timeDiff;
    }
  );

  return allTimeTableRowsRawData;
}