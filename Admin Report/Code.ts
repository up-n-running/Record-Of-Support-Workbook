var props: GoogleAppsScript.Properties.Properties = PropertiesService.getScriptProperties();
var SHEETS: any = {};

const GLOBAL_CONSTANTS = {
  LSA_ADMINS_GOOGLE_GROUP_EMAIL: "lsa-administrators@wlc.ac.uk",
  TIMEZONE: "Europe/London"
};

function CreateChildDailyReport() {

  let rootSnapshotDirectoryId = props.getProperty( "DailyReportDirectoryId" );
  let snapshotDate = new Date();

  let thisReportTemplateSpreadsheet: GoogleAppsScript.Spreadsheet.Spreadsheet = SpreadsheetApp.getActive();
  let rootSnapshotDirectory: GoogleAppsScript.Drive.Folder = DriveApp.getFolderById( rootSnapshotDirectoryId );

  //make copy of template into reports directory
  let thisReportTemplateFile: GoogleAppsScript.Drive.File = DriveApp.getFileById( thisReportTemplateSpreadsheet.getId() );
  let childReportFile = thisReportTemplateFile.makeCopy(
        thisReportTemplateFile.getName() + Utilities.formatDate( snapshotDate, GLOBAL_CONSTANTS.TIMEZONE, "__yyyy-MM-dd" ),
        rootSnapshotDirectory
  );

  let childReportSpreadsheet = SpreadsheetApp.open( childReportFile );
  RefreshRoSRawDataSnapshot( childReportSpreadsheet );
  //return the file
}

function RefreshRoSRawDataSnapshot( reportSpreadsheet: GoogleAppsScript.Spreadsheet.Spreadsheet ) {

  //parse params
  reportSpreadsheet = reportSpreadsheet ? reportSpreadsheet : SpreadsheetApp.getActive();

  let masterFileId = props.getProperty( "MasterFileId" );
  Logger.log( "Getting Master Spreadsheet with ID: " + masterFileId );
  let masterSpreadsheet: GoogleAppsScript.Spreadsheet.Spreadsheet = SpreadsheetApp.openById( masterFileId );
  let ui = SpreadsheetApp.getUi();

  let latestVersion: string = "" + props.getProperty( "MasterVersion" );
  const LATEST_SHEET: any = SHEETS[ latestVersion ];

  let mastersGlobalSettingsSheet: GoogleAppsScript.Spreadsheet.Sheet = masterSpreadsheet.getSheetByName( LATEST_SHEET.GLOBAL_SETTINGS.NAME );
  //let reportSettingsSheet: GoogleAppsScript.Spreadsheet.Sheet = reportSpreadsheet.getSheetByName( "Settings" );
  let masterLearnerWorksheet: GoogleAppsScript.Spreadsheet.Sheet = masterSpreadsheet.getSheetByName( LATEST_SHEET.MASTER_LEARNERS.NAME );
  let lsaWorksheet: GoogleAppsScript.Spreadsheet.Sheet           = masterSpreadsheet.getSheetByName( LATEST_SHEET.MASTER_LSAS.NAME );
  let reportInputSheet: GoogleAppsScript.Spreadsheet.Sheet       = reportSpreadsheet.getSheetByName( "Filter" );
  let reportRoSRawDataSheet: GoogleAppsScript.Spreadsheet.Sheet  = reportSpreadsheet.getSheetByName( "Raw Data - RoSs" );
  let reportLearnerRawDataSheet: GoogleAppsScript.Spreadsheet.Sheet = reportSpreadsheet.getSheetByName( "Raw Data - Learners" );
  let reportLSARawDataSheet: GoogleAppsScript.Spreadsheet.Sheet  = reportSpreadsheet.getSheetByName( "LSAs" );

  /* READ GLOBAL SETTINGS FROM MASTER AS SANITY CHECK */
  Logger.log( "Checking Master Spreadsheet has Version: " + latestVersion );
  let mastersVersion: string = "" + mastersGlobalSettingsSheet.getRange( 
    LATEST_SHEET.GLOBAL_SETTINGS.REFS.ROW_NO_VERSION_NO,
    LATEST_SHEET.GLOBAL_SETTINGS.REFS.COL_NO 
  ).getValue();
  if( mastersVersion !== latestVersion ) {
    ui.alert( "Incorrect Version Number", "The Version Number from the Document Property 'MasterVersion' is: '" + latestVersion + "'\n" +
          "Whereas the version from the Master Spreadsheet is: '" + mastersVersion + "'.\n" + 
          "The Document Property: 'MasterFileId' defines the current master file to be: '" + masterFileId + "'\n\nCannot generate Report please contact Support",
          ui.ButtonSet.OK );
    return;
  }

  //clear down report ready for new snapshot
  ClearReportReadyForNewSnapshot( reportSpreadsheet );
  
  /* MASTER LEARNER DATABASE COLUMNS - READ LEARNER INFO FROM MASTER */
  let learnerIdColumnValues:             Array<Array<any>>|null = null;
  let learnerForeNameColumnValues:       Array<Array<any>>|null = null;
  let learnerNickNameColumnValues:       Array<Array<any>>|null = null;
  let learnerSurNameColumnValues:        Array<Array<any>>|null = null;
  let learnerDirectoryIdColumnValues:    Array<Array<any>>|null = null;
  let learnerSignatureIdColumnValues:    Array<Array<any>>|null = null;
  //read learnerId column
  let lastLearnerRowNo = masterLearnerWorksheet.getMaxRows();
  learnerIdColumnValues = masterLearnerWorksheet.getRange( 
    LATEST_SHEET.MASTER_LEARNERS.REFS.ROW_NO_FIRST_LEARNER, 
    LATEST_SHEET.MASTER_LEARNERS.REFS.COL_NO_LEARNER_ID,
    lastLearnerRowNo - LATEST_SHEET.MASTER_LEARNERS.REFS.ROW_NO_FIRST_LEARNER + 1,
    1 ).getValues();
  //read learner First Name column
  learnerForeNameColumnValues = masterLearnerWorksheet.getRange( 
    LATEST_SHEET.MASTER_LEARNERS.REFS.ROW_NO_FIRST_LEARNER, 
    LATEST_SHEET.MASTER_LEARNERS.REFS.COL_NO_FORENAME,
    lastLearnerRowNo - LATEST_SHEET.MASTER_LEARNERS.REFS.ROW_NO_FIRST_LEARNER + 1,
    1 ).getValues();
  //read learner Nick Name column
  learnerNickNameColumnValues = masterLearnerWorksheet.getRange( 
    LATEST_SHEET.MASTER_LEARNERS.REFS.ROW_NO_FIRST_LEARNER, 
    LATEST_SHEET.MASTER_LEARNERS.REFS.COL_NO_NICKNAME,
    lastLearnerRowNo - LATEST_SHEET.MASTER_LEARNERS.REFS.ROW_NO_FIRST_LEARNER + 1,
    1 ).getValues();
  //read learner Last Name column
  learnerSurNameColumnValues = masterLearnerWorksheet.getRange( 
    LATEST_SHEET.MASTER_LEARNERS.REFS.ROW_NO_FIRST_LEARNER, 
    LATEST_SHEET.MASTER_LEARNERS.REFS.COL_NO_SURNAME,
    lastLearnerRowNo - LATEST_SHEET.MASTER_LEARNERS.REFS.ROW_NO_FIRST_LEARNER + 1,
    1 ).getValues();
  //read Directory ID column
  learnerDirectoryIdColumnValues = masterLearnerWorksheet.getRange( 
    LATEST_SHEET.MASTER_LEARNERS.REFS.ROW_NO_FIRST_LEARNER, 
    LATEST_SHEET.MASTER_LEARNERS.REFS.COL_NO_LEARNER_DIR,
    lastLearnerRowNo - LATEST_SHEET.MASTER_LEARNERS.REFS.ROW_NO_FIRST_LEARNER + 1,
    1 ).getValues();
  //read lsa workbook file Id Column
  learnerSignatureIdColumnValues = masterLearnerWorksheet.getRange( 
    LATEST_SHEET.MASTER_LEARNERS.REFS.ROW_NO_FIRST_LEARNER, 
    LATEST_SHEET.MASTER_LEARNERS.REFS.COL_NO_SIGNATURE_FILE_ID,
    lastLearnerRowNo - LATEST_SHEET.MASTER_LEARNERS.REFS.ROW_NO_FIRST_LEARNER + 1,
    1 ).getValues();
  //filter blank learners and create 2d array of final Report Learner Sheet rows
  let finalReportLearnerData2dArray: Array<Array<any>> = new Array();
  let finalReportLearnerDataRow: Array<any>|null = null;
  let learnerRowsByLearnerId: any = {};
  let tempLearnerId: number = -1;
  for( let lri = 0 ; lri < learnerIdColumnValues.length ; lri++ ) {
    tempLearnerId = learnerIdColumnValues[ lri ][ 0 ];
    if( tempLearnerId && tempLearnerId > 10 && tempLearnerId < 99999990 ) {
      finalReportLearnerDataRow = [
        tempLearnerId,
        learnerForeNameColumnValues[ lri ][0] + 
              ( ( learnerNickNameColumnValues[ lri ][0] ) ? " (" + learnerNickNameColumnValues[ lri ][0] + ") " : " " ) +
              learnerSurNameColumnValues[ lri ][0],
        learnerDirectoryIdColumnValues[ lri ][0],
        learnerSignatureIdColumnValues[ lri ][0]
      ];
      finalReportLearnerData2dArray.push( finalReportLearnerDataRow );
      learnerRowsByLearnerId[ ""+tempLearnerId ] = finalReportLearnerDataRow;
    }
  }
  //finalReportLearnerData2dArray.sort( function( rowA, rowB ) { return rowA[0] - rowB[0]; } );
  finalReportLearnerData2dArray.sort( function( rowA, rowB ) { return rowA[1].localeCompare(rowB[1]); } );


  /* LSA COLUMNS - READ LSA INFO FROM MASTER */
  let lsaNameColumnValues:           Array<Array<any>>|null = null;
  let lsaVersionColumnValues:        Array<Array<any>>|null = null;
  let lsaFileIdColumnValues:         Array<Array<any>>|null = null;
  let mostRecentLsaValidFromDate:    Date = new Date( 0 );
  //read LSA Name column
  lsaNameColumnValues = lsaWorksheet.getRange( 
    LATEST_SHEET.MASTER_LSAS.REFS.ROW_NO_FIRST_LSA, 
    LATEST_SHEET.MASTER_LSAS.REFS.COL_NO_LSA_NAME,
    LATEST_SHEET.MASTER_LSAS.REFS.ROW_NO_LAST_LSA - LATEST_SHEET.MASTER_LSAS.REFS.ROW_NO_FIRST_LSA + 1,
    1 ).getValues();
  //read LSA Name column
  lsaVersionColumnValues = lsaWorksheet.getRange( 
    LATEST_SHEET.MASTER_LSAS.REFS.ROW_NO_FIRST_LSA, 
    LATEST_SHEET.MASTER_LSAS.REFS.COL_NO_USERS_VERSION,
    LATEST_SHEET.MASTER_LSAS.REFS.ROW_NO_LAST_LSA - LATEST_SHEET.MASTER_LSAS.REFS.ROW_NO_FIRST_LSA + 1,
    1 ).getValues();
  //read lsa workbook file Id Column
  lsaFileIdColumnValues = lsaWorksheet.getRange( 
    LATEST_SHEET.MASTER_LSAS.REFS.ROW_NO_FIRST_LSA, 
    LATEST_SHEET.MASTER_LSAS.REFS.COL_NO_USERS_FILE_ID,
    LATEST_SHEET.MASTER_LSAS.REFS.ROW_NO_LAST_LSA - LATEST_SHEET.MASTER_LSAS.REFS.ROW_NO_FIRST_LSA + 1,
    1 ).getValues();
  //filter blank LSAs and LSAs with no Workbook and create 2d array of final Report LSA Sheet rows
  const IDX_HYPERLINK      = 0;
  const IDX_LSA_NAME       = 1;
  const IDX_LSA_FILE_ID    = 2;
  const IDX_LSA_VERSION    = 3;
  const IDX_LSA_OLDEST_ROS = 4;
  const IDX_LSA_VALID_FROM = 5;

  let finalReportLSAData2dArray: Array<Array<any>> = new Array();
  let finalReportLSADataRow: Array<any>|null = null;
  let tempFileId: string = "";
  for( let lsai = 0 ; lsai < lsaFileIdColumnValues.length ; lsai++ ) {
    tempFileId = lsaFileIdColumnValues[ lsai ][ 0 ];
    if( tempFileId ) {
      finalReportLSADataRow = [
        '=HYPERLINK( "https://docs.google.com/spreadsheets/d/'+tempFileId+'/edit", "'+lsaNameColumnValues[ lsai ][0]+'" )',
        lsaNameColumnValues[ lsai ][0],
        tempFileId,
        lsaVersionColumnValues[ lsai ][0],
        "",
        ""
      ];
      finalReportLSAData2dArray.push( finalReportLSADataRow );
    }
  }
  finalReportLSAData2dArray.sort( function( rowA, rowB ) { return rowA[IDX_LSA_NAME].localeCompare(rowB[IDX_LSA_NAME]); } );  


  /* READ SETTINGS FROM ADMIN SPREADSHEET AND CREATE THE COLUMN STORES WE NEED */
  let ReportSettingsMyFilesColumnHandles = reportSpreadsheet.getRangeByName( "MY_FILES_COLUMN_DEFS" ).getValues();
  let myFilesColumnNameHandleArray: Array<string> = new Array();
  let myFilesReportColNoArray: Array<number> = new Array();
  let myFilesColumnValuesArray: Array<Array<Array<any>>> = new Array();
  let myFilesLSANameColumnValueArray: Array<Array<any>> = new Array();
  for( let mfi=0; mfi < ReportSettingsMyFilesColumnHandles.length; mfi++ ) {
    if( ReportSettingsMyFilesColumnHandles[mfi][0] ) {
      myFilesReportColNoArray.push( ReportSettingsMyFilesColumnHandles[mfi][0] )
      myFilesColumnNameHandleArray.push( ReportSettingsMyFilesColumnHandles[mfi][2] );
      myFilesColumnValuesArray.push( new Array() );
    }
  }


  //prepare to loop through each of the LSA numbers in the array and check if their LSA Workbook has already been generated
  //these are used to manage the iteration through the loop
  let noOfWorksheetsFound: number= 0;
  let lsaName: string = ""; let lsaVersion: string = ""; let lsaFileId: string = "";
  let childWorkbook: GoogleAppsScript.Spreadsheet.Spreadsheet = null;
  let childMyFilesSheet: GoogleAppsScript.Spreadsheet.Sheet = null;
  let LSAS_SHEETS = null;
  let myFilesFirstRowNo: number|undefined = undefined;
  let myFilesLastRowNo: number|undefined = undefined;
  let myFilesNumRows: number = -1;

  //prepare for internal loop where we loop through the columns of the my files sheet.
  let columnHandle: string = ""; 
  let columnNo: number|undefined = undefined;
  let colValuesToSaveForThisWorkbook: Array<Array<any>>|null = null;
  let lsaOldestRoSOnRecord: Date|null = null;
  let lsaValidFromDate: Date|null = null;


  //actually loop through each of the LSA numbers in the array and check if their LSA Workbook has already been generated
  for ( var lsai=0; lsai<finalReportLSAData2dArray.length; lsai++ ) {
    lsaName    = finalReportLSAData2dArray[ lsai ][ IDX_LSA_NAME ];
    lsaVersion = finalReportLSAData2dArray[ lsai ][ IDX_LSA_VERSION ];
    lsaFileId  = finalReportLSAData2dArray[ lsai ][ IDX_LSA_FILE_ID ];
      
    LSAS_SHEETS = SHEETS[ lsaVersion ];
    myFilesFirstRowNo = LSAS_SHEETS.MY_FILES.REFS.ROW_NO_FIRST_FILE;
    myFilesLastRowNo  = LSAS_SHEETS.MY_FILES.REFS.ROW_NO_LAST_FILE;
    myFilesNumRows    = myFilesLastRowNo - myFilesFirstRowNo + 1;

    noOfWorksheetsFound++;
    childWorkbook = null;
    Logger.log( "Workbook " + (lsai+1) + " Found with ID '"+lsaFileId+"': " + childWorkbook ); 
    childWorkbook = SpreadsheetApp.openById( lsaFileId );
    Logger.log( "Opened Workbook with ID '"+lsaFileId+"': " + childWorkbook );   
    childMyFilesSheet = childWorkbook.getSheetByName( LSAS_SHEETS.MY_FILES.NAME );

    //add this My Files Sheet data column by column to the existing column stores.
    myFilesLSANameColumnValueArray.push.apply( myFilesLSANameColumnValueArray, create2DPrePopulatedArray( myFilesNumRows, 1, lsaName ) );
//Logger.log( "  myFilesLSANameColumnValueArray = " + myFilesLSANameColumnValueArray );
    //copy columns of data into the master myFilesColumnValuesArray arrays
    for( let ci = 0; ci < myFilesColumnNameHandleArray.length; ci++ ) {
      columnHandle = myFilesColumnNameHandleArray[ ci ];
      if( columnHandle ) {
        columnNo = LSAS_SHEETS.MY_FILES.REFS[ "COL_NO_" + columnHandle ];
        colValuesToSaveForThisWorkbook = null; 
        if( columnNo ) {
          colValuesToSaveForThisWorkbook = childMyFilesSheet.getRange( myFilesFirstRowNo, columnNo, myFilesNumRows, 1 ).getValues();
        }
        else {
          colValuesToSaveForThisWorkbook = create2DPrePopulatedArray( myFilesNumRows, 1, "" );
        }
//Logger.log( "  columnHandle = '"+columnHandle+"', columnNo = " + columnNo );
//Logger.log( "  colValuesToSaveForThisWorkbook = " + colValuesToSaveForThisWorkbook );
        //concatenate array 2 to array 1
        myFilesColumnValuesArray[ci].push.apply( myFilesColumnValuesArray[ci], colValuesToSaveForThisWorkbook );

        if( columnHandle === "CREATED_DATE" ) {
          lsaValidFromDate = colValuesToSaveForThisWorkbook[ colValuesToSaveForThisWorkbook.length -1 ][ 0 ] ? 
                colValuesToSaveForThisWorkbook[ colValuesToSaveForThisWorkbook.length -1 ][ 0 ] : null;
          lsaOldestRoSOnRecord = lsaValidFromDate;
          for( let di = colValuesToSaveForThisWorkbook.length -2; !lsaOldestRoSOnRecord && di>=0; di-- ) {
            if( colValuesToSaveForThisWorkbook[ di ][ 0 ] ) {
              lsaOldestRoSOnRecord = colValuesToSaveForThisWorkbook[ di ][ 0 ];
            }
          }

          lsaValidFromDate = truncateTimeFromDateObject( lsaValidFromDate );
          lsaOldestRoSOnRecord = truncateTimeFromDateObject( lsaOldestRoSOnRecord );

          //if MyFiles Store is Full
          finalReportLSAData2dArray[lsai][IDX_LSA_OLDEST_ROS] = ( lsaOldestRoSOnRecord ) ? lsaOldestRoSOnRecord : "";
          finalReportLSAData2dArray[lsai][IDX_LSA_VALID_FROM] = ( lsaValidFromDate ) ? lsaValidFromDate : "";
          if( lsaValidFromDate && lsaValidFromDate > mostRecentLsaValidFromDate ) {
            mostRecentLsaValidFromDate = lsaValidFromDate;
          }
        }
      }
    }
  }

  Logger.log( "Filtering and Restructuring data and evaluating calculated fields ready for writing");

  //collate all the column arrays into one big 2d array.
  //at the same time as filtering out the blank rows and the test learners
  //and populating the Calculated Columns
  const IDX_LESSON_NAME   = 0;
  const IDX_FILE_NAME     = 1;
  const IDX_USERS_FILE_ID = 2;
  const IDX_CREATED_DATE  = 3;
  const IDX_UPDATED_DATE  = 4;
  const IDX_DELETED_DATE  = 5;
  const IDX_LEARNER_NAME  = 6;
  const IDX_LEARNER_ID    = 7;
  const IDX_LEARNER_EMAIL = 8;
  const IDX_LESSON_DATE   = 9;
  const IDX_START_TIME    = 10;
  const IDX_DURATION      = 11;
  const IDX_SIGN_TYPE     = 12;
  const IDX_AUTOSIGN_CMTS = 13;

  let finalReportROSData2dArray: Array<Array<any>> = new Array();
  let finalReportROSDataRow: Array<any>|null = null;
  let fileId: string = "";
  let fileName: string = "";
  let learnerId: number = 99999999;
  let deletedDate: Date|null = null;
  let learnerName_CALC: string = "";
  let lessonStart_CALC: Date|null|string = null;
  let lessonEnd_CALC: Date|null|string = null;
  let authorisedAbsence_CALC: number = 0;
  let autoSignOverride_CALC: number = 0;
  let signed_CALC: number = 0;
  let signatureStatus_CALC: string = "";
  let hyperlink_CALC: string = "";
  let autosignOverrideComments_CALC: string = "";

  for( let fri = 0 ; fri < myFilesColumnValuesArray[ 1 ].length ; fri++ ) {
    fileId      = myFilesColumnValuesArray[ IDX_USERS_FILE_ID ][ fri ][ 0 ];
    learnerId   = myFilesColumnValuesArray[ IDX_LEARNER_ID    ][ fri ][ 0 ];
    deletedDate = myFilesColumnValuesArray[ IDX_DELETED_DATE  ][ fri ][ 0 ];
//Logger.log( "fileId = " + fileId + ", learnerId = " + learnerId + ", deletedDate = " + deletedDate );
    if( fileId && learnerId && learnerId < 99999990 && !deletedDate ) {
      learnerName_CALC = ( ( learnerRowsByLearnerId[""+learnerId] ) ? 
            learnerRowsByLearnerId[""+learnerId][1] : myFilesColumnValuesArray[ IDX_LEARNER_NAME ][ fri ][0] );
      try { 
        lessonStart_CALC = new Date( myFilesColumnValuesArray[ IDX_LESSON_DATE ][ fri ][0].setHours(
              myFilesColumnValuesArray[ IDX_START_TIME ][ fri ][0].getHours(),
              myFilesColumnValuesArray[ IDX_START_TIME ][ fri ][0].getMinutes(),
              0,
              0 ) 
        );
        lessonEnd_CALC = new Date( lessonStart_CALC.getTime() + 1000*60*myFilesColumnValuesArray[ IDX_DURATION ][ fri ][0] );
      } catch( e ) {
        lessonStart_CALC= "ERROR";
        lessonEnd_CALC = "ERROR";
      }

      //signature type calculated fields
      authorisedAbsence_CALC = ( myFilesColumnValuesArray[ IDX_AUTOSIGN_CMTS ][ fri ][0] == "Absence was Authorised" ) ? 1 : 0;
      autoSignOverride_CALC = ( authorisedAbsence_CALC == 0 && myFilesColumnValuesArray[ IDX_AUTOSIGN_CMTS ][ fri ][0] ) ? 1 : 0;
      signed_CALC = ( 
                    myFilesColumnValuesArray[ IDX_AUTOSIGN_CMTS ][ fri ][0] || 
                  ( myFilesColumnValuesArray[ IDX_UPDATED_DATE ][ fri ][0].getTime() - 
                    myFilesColumnValuesArray[ IDX_CREATED_DATE ][ fri ][0].getTime() > 86400 ) 
            ) ? 1 : 0;
      signatureStatus_CALC = ( authorisedAbsence_CALC ) ? "Auto - Auth Absence" : 
                               ( ( autoSignOverride_CALC ) ? "Auto - See Comments" : 
                                 ( ( signed_CALC ) ? "Signed" : "Unsigned" ) );

      fileName = myFilesColumnValuesArray[ IDX_FILE_NAME     ][ fri ][0];
      hyperlink_CALC = '=HYPERLINK("https://drive.google.com/file/d/' + fileId + '", "'+fileName+'" )';
      autosignOverrideComments_CALC = ( autoSignOverride_CALC ) ? myFilesColumnValuesArray[ IDX_AUTOSIGN_CMTS ][ fri ][0] : "";

      finalReportROSDataRow = [
        myFilesLSANameColumnValueArray[ fri ][0],
        learnerName_CALC,
        myFilesColumnValuesArray[ IDX_LESSON_NAME   ][ fri ][0],
        fileName,
        fileId,
        myFilesColumnValuesArray[ IDX_CREATED_DATE  ][ fri ][0],
        myFilesColumnValuesArray[ IDX_UPDATED_DATE  ][ fri ][0],
        myFilesColumnValuesArray[ IDX_DELETED_DATE  ][ fri ][0],
        myFilesColumnValuesArray[ IDX_LEARNER_NAME  ][ fri ][0],
        learnerId,
        myFilesColumnValuesArray[ IDX_LEARNER_EMAIL ][ fri ][0],
        myFilesColumnValuesArray[ IDX_LESSON_DATE   ][ fri ][0],
        myFilesColumnValuesArray[ IDX_START_TIME    ][ fri ][0],
        myFilesColumnValuesArray[ IDX_DURATION      ][ fri ][0],
        myFilesColumnValuesArray[ IDX_SIGN_TYPE     ][ fri ][0],
        myFilesColumnValuesArray[ IDX_AUTOSIGN_CMTS ][ fri ][0],
        lessonStart_CALC,
        lessonEnd_CALC,
        signed_CALC,
        authorisedAbsence_CALC,
        autoSignOverride_CALC,
        signatureStatus_CALC,
        hyperlink_CALC,
        autosignOverrideComments_CALC
      ];

      finalReportROSData2dArray.push( finalReportROSDataRow );
    }
  }

  Logger.log( "Writing Data to Report RoS Raw Data Sheet");
  reportRoSRawDataSheet.getRange( 2, 1, finalReportROSData2dArray.length, finalReportROSData2dArray[0].length )
  .setValues( finalReportROSData2dArray );

  Logger.log( "Writing Data to Report Learner Raw Data Sheet");
  //save results on the Learner Data Sheet
  reportLearnerRawDataSheet.getRange( 2, 1, finalReportLearnerData2dArray.length, finalReportLearnerData2dArray[0].length ).setValues( finalReportLearnerData2dArray );

  Logger.log( "Writing Data to Report LSA Raw Data Sheet");
  //save results on the ROS Data Sheet
  //lsa name col
  reportLSARawDataSheet.getRange( 2, 1, finalReportLSAData2dArray.length, finalReportLSAData2dArray[0].length ).setValues( finalReportLSAData2dArray );


  Logger.log( "Writing Start and End Valid Dates to Input Sheet");
  //save results on the ROS Data Sheet
  //lsa name col
  reportInputSheet.getRange( 2, 4, 2, 1 ).setValues( [ [ mostRecentLsaValidFromDate ], [ truncateTimeFromDateObject( new Date() ) ] ] );

  Logger.log( "Flushing Write Buffer");
  SpreadsheetApp.flush();
  Logger.log( "Done");
}


function ClearReportReadyForNewSnapshot( reportSpreadsheet: GoogleAppsScript.Spreadsheet.Spreadsheet ) {

  //parse params
  reportSpreadsheet = reportSpreadsheet ? reportSpreadsheet : SpreadsheetApp.getActive();

  Logger.log( "ClearReportReadyForNewSnapshot called with reportSpreadsheet with ID: " + reportSpreadsheet.getId() );

  let reportRoSRawDataSheet: GoogleAppsScript.Spreadsheet.Sheet  = reportSpreadsheet.getSheetByName( "Raw Data - RoSs" );
  let reportLearnerRawDataSheet: GoogleAppsScript.Spreadsheet.Sheet = reportSpreadsheet.getSheetByName( "Raw Data - Learners" );
  let reportLSARawDataSheet: GoogleAppsScript.Spreadsheet.Sheet  = reportSpreadsheet.getSheetByName( "LSAs" );

  

  Logger.log( "Clearing learner data on report down ready for repopulation at the end");
  if( reportLearnerRawDataSheet.getMaxRows() >= 2) {
    reportLearnerRawDataSheet.getRange( 2, 1, 1, reportLearnerRawDataSheet.getMaxColumns() ).clear();
  }
  if( reportLearnerRawDataSheet.getMaxRows() >= 3) {
    reportLearnerRawDataSheet.deleteRows( 3, reportLearnerRawDataSheet.getMaxRows()-2 );
  }

  Logger.log( "Clearing LSA data on report down ready for repopulation at the end");
  if( reportLSARawDataSheet.getMaxRows() >= 2) {
    reportLSARawDataSheet.getRange( 2, 1, 1, reportLSARawDataSheet.getMaxColumns() ).clear();
  }
  if( reportLSARawDataSheet.getMaxRows() >= 3) {
    reportLSARawDataSheet.deleteRows( 3, reportLSARawDataSheet.getMaxRows()-2 );
  }  

  Logger.log( "Clearing ROS data on report down ready for repopulation at the end");
  if( reportRoSRawDataSheet.getMaxRows() >= 2) {
    reportRoSRawDataSheet.getRange( 2, 1, 1, reportRoSRawDataSheet.getMaxColumns() ).clear();
  }
  if( reportRoSRawDataSheet.getMaxRows() >= 3) {
    reportRoSRawDataSheet.deleteRows( 3, reportRoSRawDataSheet.getMaxRows()-2 );
  }

  Logger.log( "Flushing Write Buffer");
  SpreadsheetApp.flush();
  Logger.log( "ClearReportReadyForNewSnapshot Done");
}


function create2DPrePopulatedArray( rows: number, cols: number, value: string): any[][] {
  var arr = Array(rows);
  for (let ri = 0; ri < rows; ri++) {
      arr[ri] = Array(cols).fill(value);
  }
  return arr;
}

function truncateTimeFromDateObject( dateObj: Date|null ) {
  if( !dateObj ) {
    return null;
  }
  return new Date( dateObj.toDateString() );
}
