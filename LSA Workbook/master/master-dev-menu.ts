function TestMasterRoSAPIWebapp() {
  CheckForUpdates();
}

function TestPushToChildWorksheets() {
  PushToChildWorksheets( null, SpreadsheetApp.getActive().getId(), "TEST ANNOUNCEMENT", null, false, false, false );
}

function CopyProtectionsFromAnotherWorkbook() {
  var destSpreadsheet = SpreadsheetApp.getActive();
  var ui = SpreadsheetApp.getUi();
  var promptResponse = ui.prompt( "Which Master Spreadsheet To Copy From", 
                                "Enter the fileId of the source Workbook here:", ui.ButtonSet.OK_CANCEL );
  if( promptResponse.getSelectedButton() == ui.Button.OK) {
    var sourceFileId = promptResponse.getResponseText();
    Logger.log( "CopyProtectionsFromAnotherWorkbook, sourceFileId = " + sourceFileId );
    var sourceSpreadsheet = SpreadsheetApp.openById( sourceFileId );
    Logger.log( "sourceSpreadsheet = " + sourceSpreadsheet );

    //get a list of Sheet Names
    var sheetNameArray = [];
    for (const property in SHEETS) {
      Logger.log(`${property}: ${SHEETS[property]}`);
      sheetNameArray.push( SHEETS[property].NAME );
    }
    Logger.log( "sheetNameArray = " + sheetNameArray );
    
    //loop through and copy protections sheet by sheet
    var sourceSheet = null, destSheet = null;
    for( var i=0; i<sheetNameArray.length; i++ ) {
      sourceSheet = sourceSpreadsheet.getSheetByName( sheetNameArray[i] );
      destSheet   = destSpreadsheet.getSheetByName(   sheetNameArray[i] );
      DuplicateSheetLevelProtection( sourceSheet, destSheet );
      DuplicateRangeLevelProtection( sourceSheet, destSheet );
    }
  }
}


function TestLearnerEmail() {
  SpreadsheetApp.getUi().alert( "Function deprecated, please correct code in master-dev-menu.ts" );
  /*
  SendLearnerSignEmail( ["18kaSt6qo1jTN4K9mS-993FO_8g-KcEVE", "122Bas36WUNYKAnGH0OK05B4ykzG8weFs"], 
                        null,
                        "Test - No Record Number",
                        "learner.test@college.wlc.ac.uk",
                        "Yasser",
                        "John Milner",
                        true,
                        null
                      ) */
}

function CopyJohnMilnersWorkbooksToTest() {
  return CopyLiveWorkbooksToTest( null, "John Milner", null )
}

function CopyLiveWorkbooksToTest( liveLSAParentDirId, folderNameFilter, cleanExitAfterSeconds: number ) {

  let lsaDirEmailsToGrantEditAccess = [ "J.MILNER@WLC.AC.UK" ];

  //parse params
  liveLSAParentDirId = (liveLSAParentDirId) ? liveLSAParentDirId : "1kJ_0TDOn7S-cXn1xspSx5XMS67MQ1K3u";
  cleanExitAfterSeconds = ( cleanExitAfterSeconds ) ? cleanExitAfterSeconds : 60*4;

  //start timer
  let startTimeMillis = ( new Date() ).getTime();

  let spreadsheet = SpreadsheetApp.getActive();
  let settingsSheet = spreadsheet.getSheetByName( SHEETS.GLOBAL_SETTINGS.NAME);
  let masterLSAsSheet = spreadsheet.getSheetByName( SHEETS.MASTER_LSAS.NAME );

  if( !isAuthorised_( spreadsheet, null, settingsSheet, true, false, true, false ) ) { return false; }

  if( !spreadsheet.getName().toUpperCase().includes("TEST") ) {
    SpreadsheetApp.getUi().alert( "CANNOT CONTINUE AS THIS MASTER FILE DOES NOT HAVE TEST IN THE NAME" );
  }

  let testAreaLSAParentDirId = settingsSheet.getRange( 
        SHEETS.GLOBAL_SETTINGS.REFS.ROW_NO_DIR_ID_LSAS,
        SHEETS.GLOBAL_SETTINGS.REFS.COL_NO
  ).getValue();
  let testAreaRoSParentDirId = settingsSheet.getRange( 
    SHEETS.GLOBAL_SETTINGS.REFS.ROW_NO_DIR_ID_ROS,
    SHEETS.GLOBAL_SETTINGS.REFS.COL_NO
).getValue();
  let testAreaMastersVersion = settingsSheet.getRange( 
        SHEETS.GLOBAL_SETTINGS.REFS.ROW_NO_VERSION_NO,
        SHEETS.GLOBAL_SETTINGS.REFS.COL_NO
  ).getValue();
  let testAreaMastersFileId = spreadsheet.getId();

  let lsaSubfolders = DriveApp.getFolderById( liveLSAParentDirId ).getFolders();
  let liveSubFolder: GoogleAppsScript.Drive.Folder = null;
  let testSubFolder: GoogleAppsScript.Drive.Folder = null;
  let testSubFolderUsers = null;
  let foundLSAUserAsEditor = false;
  let testSubFolderFiles = null;
  let testSubFolderFileNames: Array<string> = new Array();
  let feedbackText = "";
  while (lsaSubfolders.hasNext() && ( cleanExitAfterSeconds <= 0 || 
          ( ( new Date() ).getTime() - startTimeMillis ) < ( cleanExitAfterSeconds * 1000 ) ) ) {
    liveSubFolder = lsaSubfolders.next();
    Logger.log( "LSA Dir: "+ liveSubFolder.getName())

    if( !folderNameFilter || liveSubFolder.getName()==folderNameFilter ) {
      testSubFolder = createOrGetChildFolder( testAreaLSAParentDirId, liveSubFolder.getName()+"_TEST_AREA", null );

      //get array of file names of pre-existing files to avoid duplicates
      testSubFolderFiles = testSubFolder.getFiles();
      testSubFolderFileNames = new Array();
      while( testSubFolderFiles.hasNext() ) {
        testSubFolderFileNames.push( testSubFolderFiles.next().getName() )
      }
      Logger.log( "Existing contents: " + testSubFolderFileNames );

      let testGlobalSettingsSheet: GoogleAppsScript.Spreadsheet.Sheet = null;
      let liveWorkbookFiles = liveSubFolder.getFiles();
      let liveWorkbookFile: GoogleAppsScript.Drive.File = null;
      let testWorkbookFileName: string = null;
      let testWorkbookFileNamePartsArray: Array<string> = null;
      let testWorkbookFile: GoogleAppsScript.Drive.File = null;
      let testWorkbook = null;
      while (liveWorkbookFiles.hasNext()) {
        liveWorkbookFile = liveWorkbookFiles.next();
        Logger.log( "  LSA FILE: "+ liveWorkbookFile.getName());
        testWorkbookFileName = liveWorkbookFile.getName()+"_TEST_AREA";
        Logger.log( "  IMPORTING AND RENAMING TO: " + testWorkbookFileName );

        //see if this file already exists cos if it does we'll skip the import
        if( testSubFolderFileNames.includes( testWorkbookFileName ) ) {
          Logger.log( "  ALREADY EXISTS - SKIPPING IMPORT" );
        }
        else {
          testWorkbookFileNamePartsArray = testWorkbookFileName.replace( " RoS Workbook v", "|").
                replace( "_TEST_AREA", "" ).split( "|" );
          testWorkbookFile = liveWorkbookFile.makeCopy( testWorkbookFileName, testSubFolder );
          testWorkbook = SpreadsheetApp.openById( testWorkbookFile.getId() );
          testGlobalSettingsSheet = testWorkbook.getSheetByName( SHEETS.GLOBAL_SETTINGS.NAME );

          //link the test copy up to the new master
          Logger.log( "  LSA FILE: "+ testWorkbookFile.getName() + " - Linking to Test Master" );
          let testsThisFilesIdCell = GetMasterSettingsCellFromOldVersionWorkbook( testWorkbook, testGlobalSettingsSheet, 
                  SHEETS.GLOBAL_SETTINGS.REFS.HANDLE_THIS_FILES_ID, SHEETS.GLOBAL_SETTINGS.REFS.ROW_NO_THIS_FILES_ID );
          let testsMasterFileIdCell = GetMasterSettingsCellFromOldVersionWorkbook( testWorkbook, testGlobalSettingsSheet, 
                  SHEETS.GLOBAL_SETTINGS.REFS.HANDLE_MASTER_FILE_ID, SHEETS.GLOBAL_SETTINGS.REFS.ROW_NO_MASTER_FILE_ID );
          let testsMastersLinkToThisFilesIdCell = GetMasterSettingsCellFromOldVersionWorkbook( testWorkbook, testGlobalSettingsSheet, 
                  SHEETS.GLOBAL_SETTINGS.REFS.HANDLE_MASTERS_LINK_TO_THIS_CHILD, 
                  SHEETS.GLOBAL_SETTINGS.REFS.ROW_NO_MASTERS_LINK_TO_THIS_CHILD );
          let testsMastersVersionCell = GetMasterSettingsCellFromOldVersionWorkbook( testWorkbook, testGlobalSettingsSheet, 
                  SHEETS.GLOBAL_SETTINGS.REFS.HANDLE_MASTERS_VERSION, 
                  SHEETS.GLOBAL_SETTINGS.REFS.ROW_NO_MASTERS_VERSION );
          testsThisFilesIdCell.setValue( testWorkbook.getId() );
          testsMasterFileIdCell.setValue( testAreaMastersFileId );
          testsMastersLinkToThisFilesIdCell.setValue( testWorkbook.getId() );
          testsMastersVersionCell.setValue( testAreaMastersVersion );

          //find the email address of the master user from this new test workbook
          let testsMainUsersEmail = GetMasterSettingsCellFromOldVersionWorkbook( testWorkbook, testGlobalSettingsSheet, 
                  SHEETS.GLOBAL_SETTINGS.REFS.HANDLE_MAIN_USERS_EMAIL, 
                  SHEETS.GLOBAL_SETTINGS.REFS.ROW_NO_MAIN_USERS_EMAIL 
          ).getValue();

          //make sure the test areas LSAs dir has edit access for the lsa's email's user
          if( lsaDirEmailsToGrantEditAccess.includes( testsMainUsersEmail.trim().toUpperCase() ) ) {
            testSubFolderUsers = testSubFolder.getEditors();
            foundLSAUserAsEditor = false;
            for( let u = 0; !foundLSAUserAsEditor && u < testSubFolderUsers.length; u++ ) {
              foundLSAUserAsEditor = ( testSubFolderUsers[u].getEmail().trim().toUpperCase() === testsMainUsersEmail.trim().toUpperCase() );
            }
            if( !foundLSAUserAsEditor ) {
              testSubFolder.addEditor( testsMainUsersEmail.trim() );
            }
          }

          //now do the settings that dont have setting handles - namely the LSA dir id and the RoS dir Id
          let testWorkbookLsaDirRowNum = findInColumn( testGlobalSettingsSheet, liveLSAParentDirId, 
                SHEETS.GLOBAL_SETTINGS.REFS.COL_NO, 1 );
          Logger.log( "testWorkbookLsaDirRowNum = " + testWorkbookLsaDirRowNum );
          testGlobalSettingsSheet.getRange( testWorkbookLsaDirRowNum, SHEETS.GLOBAL_SETTINGS.REFS.COL_NO ).setValue(
                testAreaLSAParentDirId );
          testGlobalSettingsSheet.getRange( testWorkbookLsaDirRowNum-1, SHEETS.GLOBAL_SETTINGS.REFS.COL_NO ).setValue(
                testAreaRoSParentDirId );
                

          //add test copy to master's LSAs sheet
          let firstBlankRowNum = findInColumn( masterLSAsSheet, "", SHEETS.MASTER_LSAS.REFS.COL_NO_USERS_FILE_ID, 
                SHEETS.MASTER_LSAS.REFS.ROW_NO_FIRST_LSA, SHEETS.MASTER_LSAS.REFS.ROW_NO_LAST_LSA );
          Logger.log( "Saving LSA Child on Master LSA Sheet Row: " + firstBlankRowNum );
          masterLSAsSheet.getRange( 
                firstBlankRowNum,
                SHEETS.MASTER_LSAS.REFS.COL_NO_USERS_FILE_ID 
          ).setValue( testWorkbookFile.getId() );
          masterLSAsSheet.getRange( 
                firstBlankRowNum,
                SHEETS.MASTER_LSAS.REFS.COL_NO_LSA_NAME 
          ).setValue( testWorkbookFileNamePartsArray[0] + "_TEST_AREA" );
          masterLSAsSheet.getRange( 
                firstBlankRowNum,
                SHEETS.MASTER_LSAS.REFS.COL_NO_LSA_EMAIL 
          ).setValue( testsMainUsersEmail );
          masterLSAsSheet.getRange( 
                firstBlankRowNum,
                SHEETS.MASTER_LSAS.REFS.COL_NO_USERS_VERSION 
          ).setValue( testWorkbookFileNamePartsArray[1] );

          feedbackText += testWorkbookFile.getId() + "|" + testWorkbookFileNamePartsArray[0] + 
                "|" +  testWorkbookFileNamePartsArray[1] + "\n";
        }
      }
    }
  }
  SpreadsheetApp.flush();
  Logger.log( feedbackText ); 
  SpreadsheetApp.getUi().alert( feedbackText );
  return;
}

function TestCreateLearnerRoSFromSheetsTemplate() {
  let previewSheet = SpreadsheetApp.getActive().getSheetByName( "1" );
  let folder = DriveApp.getFolderById( "1RRcRF-fYaGA1aXMmgtgj8qeIrKTRrJ16" );
  createLearnerRoSFromSheetsTemplate_( previewSheet, folder, "Test Learner_2021-04-12_1_John Milner",
                                       //"learner.test@college.wlc.ac.uk", 
                                       "johndev@gmail.com",
                                       "1kFqsiuPpzMqevWHGoIKkyLsLAWdzLS4U" );
}