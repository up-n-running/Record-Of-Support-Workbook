function isAMasterNotAChild( spreadSheet, globalSettingsWorkSheet, alertIfMaster ) {
Logger.log( "isAMasterNotAChild called" )
    //get default values for params if any params are missing
    spreadSheet = (spreadSheet) ? spreadSheet : SpreadsheetApp.getActiveSpreadsheet();

    globalSettingsWorkSheet = (globalSettingsWorkSheet) 
                              ? globalSettingsWorkSheet 
                              : spreadSheet.getSheetByName( SHEETS.GLOBAL_SETTINGS.NAME );
Logger.log( "globalSettingsWorkSheet = " + globalSettingsWorkSheet );
    alertIfMaster = ( alertIfMaster ) ? alertIfMaster : false;
  
    var masterFileIdFromSettings = globalSettingsWorkSheet.getRange(
          SHEETS.GLOBAL_SETTINGS.REFS.ROW_NO_MASTER_FILE_ID, 
          SHEETS.GLOBAL_SETTINGS.REFS.COL_NO 
        ).getValue(); //if this field is blank it's a master

    var isMaster = masterFileIdFromSettings == "";

    if( isMaster && alertIfMaster ) {
      var ui = SpreadsheetApp.getUi();
      ui.alert ( "Action not available", "Sorry this action is not available on the Master Workbook", ui.ButtonSet.OK );
    }

    return ( isMaster );
}


function currentUserIsADev( spreadSheet, globalSettingsWorkSheet ) {

    spreadSheet = (spreadSheet) ? spreadSheet : SpreadsheetApp.getActiveSpreadsheet();

    globalSettingsWorkSheet = (globalSettingsWorkSheet) 
                              ? globalSettingsWorkSheet 
                              : spreadSheet.getSheetByName( SHEETS.GLOBAL_SETTINGS.NAME );
  
    var devEmailAddressFromSettings = globalSettingsWorkSheet.getRange(
          SHEETS.GLOBAL_SETTINGS.REFS.ROW_NO_DEVS_EMAIL, 
          SHEETS.GLOBAL_SETTINGS.REFS.COL_NO 
        ).getValue(); //if this field is blank it's a master

    //just one of the quirks of google scripts - theres three ways to get the user and they dont all return something!
    var userEmail = Session.getActiveUser().getEmail();
    userEmail = (userEmail) ? userEmail : Session.getEffectiveUser().getEmail();
    userEmail = (userEmail) ? userEmail : Session.getUser().getEmail();

    return ( devEmailAddressFromSettings.trim() == userEmail );
}

/**
 * [lsa-utils.gs]
 * Check to see if a particular action is available on this particular Workbook
 * @param spreadsheet {Spreadsheet=} The Spreadsheet Object we are checking
 * @param ui {UserInterface=} The UserInterface object we are using to make alerts
 * @param settingsSheet {Sheet=} The spreadsheet's global settings sheet
 * @param alertIfFail {bool=} Whether to throw an alert if this check fails
 * @param allowLSAMode {bool=} Whether to allow this action on Child LSA Workbooks
 * @param allowMasterMode {bool=} Whether to allow this action on the Master Workbook
 * @param allowMasterMode {bool=} Whether to allow this action on decommissioned child Workbooks
 * @return {bool} true iff we are allowed to perform the action on this workbook
 */
function isAuthorised_( spreadsheet, ui, settingsSheet, alertIfFail, allowLSAMode, allowMasterMode, allowDecomissionedMode ) {
  //parse params
  spreadsheet = ( spreadsheet ) ? spreadsheet : SpreadsheetApp.getActive();
  alertIfFail = ( alertIfFail ) ? true : false;
  ui = ( ui ) ? ui : ( alertIfFail ? SpreadsheetApp.getUi() : null );
  settingsSheet = ( settingsSheet ) ? settingsSheet : spreadsheet.getSheetByName( SHEETS.GLOBAL_SETTINGS.NAME );
  allowLSAMode = ( allowLSAMode ) ? true : false;
  allowMasterMode = ( allowMasterMode ) ? true : false;
  allowDecomissionedMode = ( allowDecomissionedMode ) ? true : false;

  if( !allowLSAMode || ! allowMasterMode ) {

    let masterFileIdFromSettings = settingsSheet.getRange(
          SHEETS.GLOBAL_SETTINGS.REFS.ROW_NO_MASTER_FILE_ID, 
          SHEETS.GLOBAL_SETTINGS.REFS.COL_NO 
        ).getValue(); //if this field is blank it's a master
    let isMaster = masterFileIdFromSettings == "";

    if( isMaster && !allowMasterMode ) {
      if( alertIfFail ) {
        ui.alert ( "Action not available", "Sorry this action is not available on the Master Workbook", ui.ButtonSet.OK );
      }
      return false;
    }

    if( !isMaster && !allowLSAMode ) {
      if( alertIfFail ) {
        ui.alert ( "Action not available", "Sorry this action is only available on the Master Workbook", ui.ButtonSet.OK );
      }
      return false;
    }
  }

  if( !allowDecomissionedMode && !isCorrectInstance( settingsSheet ) ) {
    if( alertIfFail ) {
      ui.alert ( "Action not available", "Sorry this action is not available on a decomissioned Workbook", ui.ButtonSet.OK );
    }
    return false;
  }

  return true;
}


/**
 * [lsa-utils.gs]
 * Find a particular LSAs folder given the LSAs workbook file object
 * 
 * In google drive a file can have >1 parent folder so if there is more than one it will look for the one:
 *  - whose parent folder in turn matches the id passed in rootAllLSAsFolderId (if rootAllLSAsFolderId is set)
 *  - and / or with matching folderName ( if folderName parameter is set )
 * 
 * If there is at least one parent but neither match then it will return the first parent folder found 
 * If there is no parent then something is very wrong so log this then cleanly return the root folder
 * 
 * @param workBookFile {File} The File object of the LSAs Workbook
 * @param folderName {String=} The Directory name of the LSAs Directory
 * @param rootAllLSAsFolderId {String=} The Id the LSAs Directory's root folder in turn
 * @return {Folder} The found folder based on the best guess algorithm described or the Users Root folder if not found
 */
function getLSAsFolderFromWorkbookFile( workBookFile, folderName, rootAllLSAsFolderId ) {

  //parse optional param
  folderName = (folderName) ? folderName : null;
  rootAllLSAsFolderId = (rootAllLSAsFolderId) ? rootAllLSAsFolderId : null;

  Logger.log( "getLSAsFolderFromWorkbookFile called, folderName = '"+folderName+"', rootAllLSAsFolderId = '"+rootAllLSAsFolderId+"'" );

  var parents = workBookFile.getParents();
  var usersFolder = null, tempFolder = null;
  let keepLooking = true;
  while ( keepLooking && parents.hasNext() ) {
    tempFolder = parents.next();
    if( folderName == null && rootAllLSAsFolderId == null ) {
        //its a match
        Logger.log( "ITS A MATCH - RETURNING FIRST PARENT FOUND AS NO FILTERS USED" );        
        usersFolder = tempFolder;
        keepLooking = false;
    }
    else if( folderName == null || tempFolder.getName() == folderName ) {
      if( rootAllLSAsFolderId == null ) {
        //its a match
        Logger.log( "ITS A MATCH - NO rootAllLSAsFolderId FILTER APPLIED" );        
        usersFolder = tempFolder;
        keepLooking = false;
      }
      else {
        let parentsParents = tempFolder.getParents();
        let parentsParent
        while( parentsParents.hasNext() ) {
          parentsParent = parentsParents.next();
          if( rootAllLSAsFolderId == parentsParent.getId() ) {
            //its a match
            Logger.log( "ITS A MATCH: rootAllLSAsFolderId MATCHES & folderName == null || tempFolder.getName() == folderName" );        
            usersFolder = tempFolder;
            keepLooking = false;
          }
        }
      }
    }
  }

  return usersFolder;
}


/**
 * Check that the LSA Name and the Record Date fields are populated on the Input SHeet
 */
function validateInputSheetFields( inputSheet, ui, settingsSheet ) {

  Logger.log( "inputSheet.getRange( SHEETS.INPUT.REFS.LESSON_DATE.ROW_NO, SHEETS.INPUT.REFS.LESSON_DATE.COL_NO ).getValue() = " + inputSheet.getRange( SHEETS.INPUT.REFS.LESSON_DATE.ROW_NO, SHEETS.INPUT.REFS.LESSON_DATE.COL_NO ).getValue() );
  let recordDate: Date = inputSheet.getRange( SHEETS.INPUT.REFS.LESSON_DATE.ROW_NO, SHEETS.INPUT.REFS.LESSON_DATE.COL_NO ).getValue();
  recordDate = new Date( recordDate );
  let today = new Date( new Date().setHours(0,0,0,0) );
  let endOfToday = new Date( new Date().setHours(23,59,59,59) );

  if( inputSheet.getRange( SHEETS.INPUT.REFS.LESSON_DATE.ROW_NO, SHEETS.INPUT.REFS.LESSON_DATE.COL_NO ).getValue() == "" ) {
    ui.alert( "What Date Are These Records For?", 
              "Please enter the date at the top of the '"+SHEETS.INPUT.NAME+"' Sheet and then try again.", 
              ui.ButtonSet.OK );
    return false;
  }
  else if( !isValidDate( recordDate ) ) {
    ui.alert( "Invalid Record Date", 
              "The 'Date of Lessons' entered at the top of the '"+SHEETS.INPUT.NAME+"' Sheet is not a valid date.\n\n" +
              "The 'New Day Cleanup' button will always enter todays date for you.\n\n" +
              "Please enter a valid date manually however in this instance, to avoid losing your RoS data.", 
              ui.ButtonSet.OK );
    return false;
  }
  else if( recordDate > endOfToday ) {
    ui.alert( "Future Record Date", 
              "The 'Date of Lessons' entered at the top of the '"+SHEETS.INPUT.NAME+"' Sheet is in the future.\n\n" +
              "You cannot generate Records of Support for future dates, sorry.", 
              ui.ButtonSet.OK );
    return false;
  }
  else if( recordDate < today ) {
    settingsSheet = ( settingsSheet ) ? settingsSheet : SpreadsheetApp.getActive().getSheetByName( SHEETS.GLOBAL_SETTINGS.NAME );
    let noOfDaysAgoAllowed = settingsSheet.getRange( SHEETS.GLOBAL_SETTINGS.REFS.ROW_NO_MAX_PAST_DAYS_ROS, 
                                                     SHEETS.GLOBAL_SETTINGS.REFS.COL_NO ).getValue();
    if( today.getTime() - recordDate.getTime() >  noOfDaysAgoAllowed*(24*3600*1000) ) {
      ui.alert( "Record Date too far in the past", 
                "The 'Date of Lessons' entered at the top of the '"+SHEETS.INPUT.NAME+"' Sheet is too far in the past.\n\n" +
                "You can only generate Records of Support up to "+noOfDaysAgoAllowed+" days in the past, sorry.", 
                ui.ButtonSet.OK );
      return false;
    }
  }
  if( recordDate < today ) {
    let confirmResponse = ui.alert( "Generate Historic Records of Support?", 
              "The 'Date of Lessons' entered at the top of the '"+SHEETS.INPUT.NAME+"' Sheet is in the past.\n\n" +
              "Are you sure you want to generate Records of Support for a previous date?", 
              ui.ButtonSet.YES_NO );
    if( confirmResponse != ui.Button.YES ) {
      return false;
    }
  }
  if( inputSheet.getRange( SHEETS.INPUT.REFS.LSA_NAME.ROW_NO, SHEETS.INPUT.REFS.LSA_NAME.COL_NO ).getValue() == "" ) {
    ui.alert( "What is your name?", 
              "Please enter your name (LSA Name) at the top of '"+SHEETS.INPUT.NAME+"' Sheet and then try again.", 
              ui.ButtonSet.OK );
    return false;
  }
  return true;
}


/**
 * Returns true if it worked, false if it didn't
 */
function ExportOneRoSRecord( recordNo, ui, learnerNamesToEmail, learnerNamesToSkip ) {
  //if last param is missing
  ui = (ui) ? ui : SpreadsheetApp.getUi();

  //find which column letter we should be accessing
  let recordColumnLetter = columnToLetter( recordNo + SHEETS.INPUT.REFS.COL_NO_RECORD_1 - 1 );
  
  let spreadsheet = SpreadsheetApp.getActive();

  //get RoS Sheet
  let recordSheet = spreadsheet.getSheetByName(recordNo);

  uiSensitiveToast( spreadsheet, ui, "Generating Record " + recordNo, "Saving Record " + recordNo );

  //if( recordSheet ) {
  refreshRoSSheetsLinkToInputSheet_( recordNo, recordSheet, spreadsheet );

  if( recordSheet.isSheetHidden() ) {
      recordSheet.showSheet();
  }
  let globalSettingsSheet = spreadsheet.getSheetByName( SHEETS.GLOBAL_SETTINGS.NAME );
  let rootRoSDirectoryId = globalSettingsSheet.getRange( 
        SHEETS.GLOBAL_SETTINGS.REFS.ROW_NO_DIR_ID_ROS,
        SHEETS.GLOBAL_SETTINGS.REFS.COL_NO
      ).getValue();
  
  let inputSheet          = spreadsheet.getSheetByName( SHEETS.INPUT.NAME);
  let recordDate          = inputSheet.getRange( SHEETS.INPUT.REFS.REF_DATE ).getValue();
  let recordLSAName       = inputSheet.getRange( SHEETS.INPUT.REFS.REF_LSA_NAME ).getValue();
  let recordStartTime     = inputSheet.getRange( "" + recordColumnLetter + SHEETS.INPUT.REFS.ROW_NO_START_TIME ).getValue();
  let recordDuration      = inputSheet.getRange( "" + recordColumnLetter + SHEETS.INPUT.REFS.ROW_NO_DURATION ).getValue();
  let recordLearnerName   = inputSheet.getRange( "" + recordColumnLetter + SHEETS.INPUT.REFS.ROW_NO_LEARNER_NAME ).getValue();
  let recordLessonName    = inputSheet.getRange( "" + recordColumnLetter + SHEETS.INPUT.REFS.ROW_NO_LESSON_NAME ).getValue();
  let recordEmailAddress  = inputSheet.getRange( "" + recordColumnLetter + SHEETS.INPUT.REFS.ROW_NO_EMAIL_ADDRESS ).getValue();
  recordEmailAddress      = substututeIfPlaceholderEmailAddress_( recordEmailAddress );
  let recordAutoSignComments = inputSheet.getRange( "" + recordColumnLetter + SHEETS.INPUT.REFS.ROW_NO_AUTOSIGN_COMMENTS ).getValue();
  let recordLearnerId     = inputSheet.getRange( "" + recordColumnLetter + SHEETS.INPUT.REFS.ROW_NO_LEARNER_ID ).getValue();
  let learnerObj: any = getChildLearnerObjByLearnerIdFromSameVersionSource_( spreadsheet, null, recordLearnerId, null );
  
Logger.log("3");
Logger.log( "learnerNamesToSkip.includes( '"+recordLearnerName+"' ) = " + learnerNamesToSkip.includes( recordLearnerName ) );

  //only export if not a skipped name
  if( learnerNamesToSkip.includes( recordLearnerName ) )
  {
    learnerObj.SIGN_TYPE = "Skip";
  }

  Logger.log( "learnerObj.SIGN_TYPE = '" + learnerObj.SIGN_TYPE + "'" );
  Logger.log( "learnerObj.SIGNATURE_ID = '" + learnerObj.SIGNATURE_ID + "'" );

  //perform a sanity check to make sure the learner data is okay as mey have have to adjust it or skip it
  if( learnerObj.SIGN_TYPE == "Stored" && !learnerObj.SIGNATURE_ID ) {
    learnerObj.SIGN_TYPE = "PDF";
    if( !learnerNamesToEmail.includes( recordLearnerName ) ) {
      let selectedBtn = ui.alert( "No Stored Signature for " + recordLearnerName, 
            recordLearnerName + " is set to " +
            "use 'Stored Signature' type Records of Support.\n\n" + "However, they have not yet saved " + 
            "their signature so we can only send the PDF version instead which is a bit " +
            "more complex for them to sign.\n\n" +
            "After you have exported your records of support please go to the 'Settings - Learners' sheet and " +
            "use the 'Stored Signature' column to generate a reminder email for them to prompt them to save their Signature with us.\n\n" +
            "Press 'OK' to send their Records of Support in the PDF format instead (you may have to talk your learner through the signing process).\n" +
            "Or press 'Cancel' to skip generating Records of Support for " + recordLearnerName,
            ui.ButtonSet.OK_CANCEL );
      if( selectedBtn == ui.Button.CANCEL ) {
        learnerObj.SIGN_TYPE = "Skip";
        learnerNamesToSkip.push( recordLearnerName );
      }
    }
  }

Logger.log("3a");

  let rosFile = null;
  
  if( learnerObj.SIGN_TYPE == "Stored" ) {
    rosFile = saveRosSpreadsheetToGivenFolder(
                  spreadsheet,
                  rootRoSDirectoryId,
                  learnerObj, 
                  recordLearnerName, 
                  recordStartTime, 
                  recordDuration,
                  recordLSAName, 
                  recordDate,
                  recordNo,
                  ( recordAutoSignComments != "" ),
                  recordSheet,
                  recordEmailAddress,
                  ui );
  }
  else if( learnerObj.SIGN_TYPE == "PDF" ) {
    rosFile = savePDFsToGivenFolder(
                  rootRoSDirectoryId,
                  learnerObj, 
                  recordLearnerName, 
                  recordStartTime, 
                  recordDuration,
                  recordLSAName, 
                  recordDate,
                  recordNo,
                  ( recordAutoSignComments != "" ),
                  null, 
                  recordSheet.getSheetId(),
                  ui );
  }

Logger.log("4");
  if( rosFile ) { //if it succeeded
    //give file permission to learner if they dont have it already
    let existingEditors    = rosFile.getEditors();
    existingEditors.push( { getEmail: function() { return "lsa.admin@wlc.ac.uk"; } } ); //TO DO: CHANGE THIS HACK TO GET VALUE FROM GLOBAL SETTINGS
    let needToAddEditor    = true;
    for( var e=0; needToAddEditor && e<existingEditors.length; e++ ) {
      needToAddEditor = !( recordEmailAddress == existingEditors[e].getEmail() );
    }
    if( needToAddEditor ) {
      rosFile.addEditor(recordEmailAddress);
    }
    
    //save file info in spreadsheet to activate hyperlinks to see file on Input and 'My RoS Files' Sheets
    AddFileToWorkbook( spreadsheet, inputSheet, null, recordNo, rosFile, recordLearnerName, recordLearnerId, 
          recordEmailAddress, recordLessonName, recordDate, recordStartTime, recordDuration, learnerObj.SIGN_TYPE,
          recordAutoSignComments );
Logger.log("5a");
    return rosFile.getId();
  }
Logger.log("5b");
  return null; //if it failed
}


function getOrCreateLearnerYearAndMonthSubDir( spreadsheet: GoogleAppsScript.Spreadsheet.Spreadsheet,
  recordDate: any, learnerName: string, learnerObj: any, rootRoSDirectoryId: string ) { 

  //learner folder id should be passed in on learner record but if learner folder id is "" then it's probably a short term learner
  //so just use the learner name and create (or get) the directory
  //if learnerObj not found (edge case - should never happen) then just use learner name as folder name - just in case
  let learnerFolder = null;
  let learnerFolderId = ""
  let learnerFolderToCreateName = learnerName;
  if( learnerObj != null ) {
    learnerFolderId = learnerObj.LEARNER_DIR; //still might be empty string
    learnerFolderToCreateName = learnerFolderToCreateName + " (" + learnerObj.LEARNER_ID + ")" + 
          ( ( learnerObj.CATEGORY != "Long Term" ) ? " " + learnerObj.CATEGORY : "" );
  }
  if( learnerFolderId != "" ) {
    learnerFolder = DriveApp.getFolderById( learnerFolderId );
  }
  if( learnerFolder == null ) {
    learnerFolder = createOrGetChildFolder( rootRoSDirectoryId, learnerFolderToCreateName, null );
  }

  //if there was no learner folder stored in the Settings - Learners sheet ( probably because its a short term learner )
  //then save the link to it here
  if( learnerObj != null && learnerObj.LEARNER_DIR == "" ) {
    let learnerSheet = spreadsheet.getSheetByName( SHEETS.SETTINGS_LEARNERS.NAME );
    let learnerRowNum = findInColumn( learnerSheet, learnerObj.LEARNER_ID, SHEETS.SETTINGS_LEARNERS.REFS.COL_NO_LEARNER_ID,
          SHEETS.SETTINGS_LEARNERS.REFS.ROW_NO_FIRST_LEARNER, SHEETS.SETTINGS_LEARNERS.REFS.ROW_NO_LAST_LEARNER );
    if( learnerRowNum > 0 ) {
      learnerSheet.getRange( learnerRowNum, SHEETS.SETTINGS_LEARNERS.REFS.COL_NO_LEARNER_DIR ).setValue( learnerFolder.getId() );
      learnerObj.LEARNER_DIR = learnerFolder.getId();
    }
  }

  let learnerMonthFolder =  createOrGetChildFolder( 
        learnerFolder.getId(), Utilities.formatDate(new Date( recordDate ), GLOBAL_CONSTANTS.TIMEZONE, "yyyy-MM" ), null 
  );

  return learnerMonthFolder;
}

function generateRosFileName(learnerName: string, lessonTimeObj: any, lsaName: string, isAutoSign: boolean, isPDF: boolean ) {
  return learnerName + '_' + 
          Utilities.formatDate( lessonTimeObj.start, GLOBAL_CONSTANTS.TIMEZONE, "yyyy-MM-dd" ) + '_' + 
          Utilities.formatDate( lessonTimeObj.start, GLOBAL_CONSTANTS.TIMEZONE, "HH:mm" ) + '-' +
          Utilities.formatDate( lessonTimeObj.end, GLOBAL_CONSTANTS.TIMEZONE, "HH:mm" ) + '_' +
          lsaName + 
          ( ( isAutoSign ) ? "_AUTO_SIGNED" : "") +
          ( ( isPDF ) ? ".pdf" : "");
}

/**
 * Save brand new file info in spreadsheet to activate hyperlinks to see file on Input and 'My RoS Files' Sheets
 * Creates new My RoS Files record if there isn't one already for that fileId (which their should be with an add)
 */
function AddFileToWorkbook( spreadsheet, inputSheet, myFilesSheet, recordNo, pdfFile, learnerName, learnerId, learnerEmail,
                            lessonName, lessonDate, startTime, duration, signType, autoSignComments ) {
Logger.log( "AddFileToWorkbook called, learnerEmail = "+learnerEmail+", startTime = "+startTime+", autoSignComments = " + autoSignComments );
  //parse params
  spreadsheet   = ( spreadsheet )   ? spreadsheet   : SpreadsheetApp.getActive();
  inputSheet    = ( inputSheet )    ? inputSheet    : spreadsheet.getSheetByName( SHEETS.INPUT.NAME     );
  myFilesSheet  = ( myFilesSheet )  ? myFilesSheet  : spreadsheet.getSheetByName( SHEETS.MY_FILES.NAME  );
  
  //get file info
  var fileId   = pdfFile.getId();
  var fileName = pdfFile.getName();
  var fileLastUpdated = pdfFile.getLastUpdated();

  //parse lesson times information
  let lessonTimeObj = getlessonTimesObjectFromInputSheetValues( lessonDate, startTime, duration );


  //Find the file in the My RoS Files Sheet
  var fileRowNo  = getRowNumOnFilesSheetReadyForInsertOrUpdate( myFilesSheet, fileId, null );

  //now make the actual update to the record on the My RoS Files Sheet
  myFilesSheet.getRange( fileRowNo, SHEETS.MY_FILES.REFS.COL_NO_FILE_NAME    ).setValue( fileName  );
  myFilesSheet.getRange( fileRowNo, SHEETS.MY_FILES.REFS.COL_NO_USERS_FILE_ID).setValue( fileId    );
  myFilesSheet.getRange( fileRowNo, SHEETS.MY_FILES.REFS.COL_NO_CREATED_DATE ).setValue( fileLastUpdated );
  myFilesSheet.getRange( fileRowNo, SHEETS.MY_FILES.REFS.COL_NO_UPDATED_DATE ).setValue( fileLastUpdated );
  myFilesSheet.getRange( fileRowNo, SHEETS.MY_FILES.REFS.COL_NO_DELETED_DATE ).setValue( "" );
  myFilesSheet.getRange( fileRowNo, SHEETS.MY_FILES.REFS.COL_NO_LEARNER_NAME ).setValue( learnerName );
  myFilesSheet.getRange( fileRowNo, SHEETS.MY_FILES.REFS.COL_NO_LEARNER_ID   ).setValue( learnerId );
  myFilesSheet.getRange( fileRowNo, SHEETS.MY_FILES.REFS.COL_NO_LEARNER_EMAIL).setValue( learnerEmail );
  myFilesSheet.getRange( fileRowNo, SHEETS.MY_FILES.REFS.COL_NO_LESSON_NAME  ).setValue( lessonName );
  myFilesSheet.getRange( fileRowNo, SHEETS.MY_FILES.REFS.COL_NO_LESSON_DATE  ).setValue( lessonDate );
  myFilesSheet.getRange( fileRowNo, SHEETS.MY_FILES.REFS.COL_NO_START_TIME   ).setValue( Utilities.formatDate( lessonTimeObj.start, GLOBAL_CONSTANTS.TIMEZONE, "HH:mm" ) );
  myFilesSheet.getRange( fileRowNo, SHEETS.MY_FILES.REFS.COL_NO_DURATION     ).setValue( lessonTimeObj.duration );
  myFilesSheet.getRange( fileRowNo, SHEETS.MY_FILES.REFS.COL_NO_SIGN_TYPE    ).setValue( signType );
  myFilesSheet.getRange( fileRowNo, SHEETS.MY_FILES.REFS.COL_NO_AUTOSIGN_CMTS).setValue( autoSignComments );
  

  //now make the actual update to the record on the Input Files Sheet
  var recordColumnNo  = recordNo + SHEETS.INPUT.REFS.COL_NO_RECORD_1 - 1;
  inputSheet.getRange( SHEETS.INPUT.REFS.ROW_NO_FILE_ID       , recordColumnNo ).setValue( fileId    );
  inputSheet.getRange( SHEETS.INPUT.REFS.ROW_NO_FILE_NAME     , recordColumnNo ).setValue( fileName  );
  inputSheet.getRange( SHEETS.INPUT.REFS.ROW_NO_FILE_CREATED  , recordColumnNo ).setValue( fileLastUpdated );
  inputSheet.getRange( SHEETS.INPUT.REFS.ROW_NO_FILE_UPDATED  , recordColumnNo ).setValue( fileLastUpdated );
}

/**
 * Update existing file's updated date in spreadsheet to activate hyperlinks to see file on Input and 'My RoS Files' Sheets
 * Creates new My RoS Files record if there isn't one already for that fileId (which their should be with an add)
 */
function UpdateFileInfoInWorkbook( spreadsheet, inputSheet, myFilesSheet, recordNo, rosFile, flagAsJustTrashed ) {
  //parse params
  spreadsheet   = ( spreadsheet )   ? spreadsheet   : SpreadsheetApp.getActive();
  inputSheet    = ( inputSheet )    ? inputSheet    : spreadsheet.getSheetByName( SHEETS.INPUT.NAME     );
  myFilesSheet  = ( myFilesSheet )  ? myFilesSheet  : spreadsheet.getSheetByName( SHEETS.MY_FILES.NAME  );
  
  //get file info
  var fileId   = rosFile.getId();
  var fileLastUpdated = rosFile.getLastUpdated();

  //Find the file in the My RoS Files Sheet
  var fileRowNo  = getRowNumOnFilesSheetReadyForInsertOrUpdate( myFilesSheet, fileId, true );

  //now make the actual update to the record on the My RoS Files Sheet
  //assuming row was found above - it SHOULD ALWAYS be found above but in reality it isnt
  //when input data is being imported from a previous version of a workbook after an upgrade
  //so handle this elegantly
  if( fileRowNo > 0 ) {
    myFilesSheet.getRange( fileRowNo, SHEETS.MY_FILES.REFS.COL_NO_UPDATED_DATE ).setValue( fileLastUpdated );
    if( flagAsJustTrashed ) {
      myFilesSheet.getRange( fileRowNo, SHEETS.MY_FILES.REFS.COL_NO_DELETED_DATE ).setValue( new Date() );
    }
  }

  //if we dont pass in a recordNo, then we must find it (if it exists - only todays files will exist)
  if( !recordNo || recordNo < 0 ) {
    let foundColNo = findInRow( inputSheet, fileId, SHEETS.INPUT.REFS.ROW_NO_FILE_ID, 
                                SHEETS.INPUT.REFS.COL_NO_RECORD_1, SHEETS.INPUT.REFS.COL_NO_RECORD_LAST );
    recordNo = ( foundColNo >= 0 ) ? ( foundColNo - SHEETS.INPUT.REFS.COL_NO_RECORD_1 + 1 ) : null;
  }

  //now make the actual update to the record on the Input Files Sheet
  if( recordNo && recordNo >= 0 ) {
    let recordColumnNo  = recordNo + SHEETS.INPUT.REFS.COL_NO_RECORD_1 - 1;
    inputSheet.getRange( SHEETS.INPUT.REFS.ROW_NO_FILE_UPDATED  , recordColumnNo ).setValue( fileLastUpdated );
    if( flagAsJustTrashed ) {
      inputSheet.getRange( SHEETS.INPUT.REFS.ROW_NO_FILE_ID       , recordColumnNo ).setValue( "" );
      inputSheet.getRange( SHEETS.INPUT.REFS.ROW_NO_FILE_NAME     , recordColumnNo ).setValue( "" );
      inputSheet.getRange( SHEETS.INPUT.REFS.ROW_NO_FILE_CREATED  , recordColumnNo ).setValue( "" );
      inputSheet.getRange( SHEETS.INPUT.REFS.ROW_NO_FILE_UPDATED  , recordColumnNo ).setValue( "" );
    }
  }
}


function getRowNumOnFilesSheetReadyForInsertOrUpdate( myFilesSheet, fileId, dontAllowCreateNewRow )
{
  dontAllowCreateNewRow = ( dontAllowCreateNewRow ) ? true : false;

    //Find the file in the My RoS Files Sheet
  var fileRowNo  = findInColumn( myFilesSheet, fileId, SHEETS.MY_FILES.REFS.COL_NO_USERS_FILE_ID, 
                                 SHEETS.MY_FILES.REFS.ROW_NO_FIRST_FILE, SHEETS.MY_FILES.REFS.ROW_NO_LAST_FILE );
  
  //if file doesnt exist in the list the move all rows down one, losing the last row in the process to make space in the top row
  if( fileRowNo < 0 && !dontAllowCreateNewRow ) {
    fileRowNo = SHEETS.MY_FILES.REFS.ROW_NO_FIRST_FILE;

    //check if the top row is blank
    var preExistingFileName = myFilesSheet.getRange( fileRowNo, SHEETS.MY_FILES.REFS.COL_NO_FILE_NAME ).getValue();
    if( preExistingFileName != "" ) { 
      //copy all bar the bottom row down one
      myFilesSheet.getRange( 
                      SHEETS.MY_FILES.REFS.ROW_NO_FIRST_FILE, SHEETS.MY_FILES.REFS.COL_NO_FILE_NAME,
                      SHEETS.MY_FILES.REFS.ROW_NO_LAST_FILE - SHEETS.MY_FILES.REFS.ROW_NO_FIRST_FILE,
                      SHEETS.MY_FILES.REFS.COL_NO_AUTOSIGN_CMTS - SHEETS.MY_FILES.REFS.COL_NO_FILE_NAME + 1 
                  ).copyTo ( 
                      myFilesSheet.getRange( SHEETS.MY_FILES.REFS.ROW_NO_FIRST_FILE + 1, SHEETS.MY_FILES.REFS.COL_NO_FILE_NAME ),
                      SpreadsheetApp.CopyPasteType.PASTE_VALUES, false
                  );
      //delete top row
      myFilesSheet.getRange( SHEETS.MY_FILES.REFS.ROW_NO_FIRST_FILE, SHEETS.MY_FILES.REFS.COL_NO_FILE_NAME,
                             1, SHEETS.MY_FILES.REFS.COL_NO_AUTOSIGN_CMTS - SHEETS.MY_FILES.REFS.COL_NO_FILE_NAME + 1 )
                  .clear({contentsOnly: true, skipFilteredRows: false});
    }
  }
  return fileRowNo;
}

/**
 * [lsa-utils.gs]
 * Loops through each of the 25 records on the Input sheet that have status of Sent to Learner, 
 * except for those records who'se fileid is in fileIdsToSkip.
 * keeps track of which are signed and which are still unsigned in the optional array parms, stillUnsignedFileIds and newlySignedFileIds
 * Also updates the spreadsheet to save the file info on both Input and My Files Sheets which in turn refreshes status bar to be up to 
 * date
 * @param spreadsheet {Spreadsheet=} The spreadsheet containing the input sheet
 * @param inputSheet {Sheet=} The input sheet to check
 * @param ui {UserInterface=} The ui opject if there is one
 * @param fileIdsToSkip {String[]=} Array of fileids not to bother checking
 * @param onlyTheseLearnerNames {String[]=} Array of Learner Names to filter by. pass null to apply no filter and check all learners
 * @return filesCheckedFileIDs {String[]} The File IDs of all the files actually checked
 */
function CheckForSignatures_InputSheet( spreadsheet, inputSheet, ui, fileIdsToSkip, onlyTheseLearnerNames ) {

  //parse params
  spreadsheet = (spreadsheet) ? spreadsheet : SpreadsheetApp.getActive();
  inputSheet = (inputSheet) ? inputSheet : spreadsheet.getSheetByName( SHEETS.INPUT.NAME );
  ui = (ui) ? ui : SpreadsheetApp.getUi();
  fileIdsToSkip = (fileIdsToSkip) ? fileIdsToSkip : new Array();

  //return array
  var fileIdsChecked = new Array();
  
  //switch UI to Input Sheet
  if( !ui.FAKE_MODE ) { spreadsheet.setActiveSheet(inputSheet, true); }

  let noOfRecordsOnInputSheet = SHEETS.INPUT.REFS.COL_NO_RECORD_LAST-SHEETS.INPUT.REFS.COL_NO_RECORD_1 + 1;
  for( let recordNo = 1 ; recordNo <=noOfRecordsOnInputSheet ; recordNo++ ) {
    let recordColumnNo = recordNo + SHEETS.INPUT.REFS.COL_NO_RECORD_1 - 1;
    let fileId = inputSheet.getRange( SHEETS.INPUT.REFS.ROW_NO_FILE_ID, recordColumnNo ).getValue();
    let recordStatus = inputSheet.getRange( SHEETS.INPUT.REFS.ROW_NO_STATUS_BAR, recordColumnNo ).getValue();
    let recordLearnerName = inputSheet.getRange( SHEETS.INPUT.REFS.ROW_NO_LEARNER_NAME, recordColumnNo ).getValue();

    if( fileId!=="" && 
        recordStatus != SHEETS.INPUT.STATUSES.SIGNED && 
        recordStatus != SHEETS.INPUT.STATUSES.SAVED_EMAILWAIT && 
        recordStatus != SHEETS.INPUT.STATUSES.SIGNED_AUTOSIGN && 
        !fileIdsToSkip.includes(fileId) && 
        ( !onlyTheseLearnerNames || onlyTheseLearnerNames.includes( recordLearnerName ) ) ) {

      //select the cell so user can follow where we are
      if( !ui.FAKE_MODE ) { 
        inputSheet.setActiveSelection( inputSheet.getRange( 
          SHEETS.INPUT.REFS.ROW_NO_STATUS_BAR, 
          recordColumnNo ) 
        );
      }

      //refresh the updated date on file data on both sheets
      UpdateFileInfoInWorkbook( spreadsheet, inputSheet, null, recordNo, DriveApp.getFileById( fileId ), false );
      fileIdsChecked.push( fileId );
    }
  }

  return fileIdsChecked;
}


/**
 * [lsa-utils.gs]
 * Loops through each of the 25 records on the Input sheet that have status of Sent to Learner, 
 * except for those records who'se fileid is in fileIdsToSkip.
 * keeps track of which are signed and which are still unsigned in the optional array parms, stillUnsignedFileIds and newlySignedFileIds
 * Also updates the spreadsheet to save the file info on both Input and My Files Sheets which in turn refreshes status bar to be up to 
 * date
 * @param spreadsheet {Spreadsheet=} The spreadsheet containing the My Files sheet
 * @param myFilesSheet {Sheet=} The My Files sheet to check
 * @param ui {UserInterface=} The ui opject if there is one
 * @param fileIdsToSkip {String[]=} Array of fileids not to bother checking
 * @param onlyTheseLearnerIds {String[]=} Array of Learner IDs to filter by. pass null to apply no filter and check all learners
 * @return filesCheckedFileIDs {String[]} The File IDs of all the files actually checked
 */
function CheckForSignatures_MyFilesSheet( spreadsheet, inputSheet, myFilesSheet, ui, fileIdsToSkip, onlyTheseLearnerIds ) {

  //parse params
  spreadsheet = (spreadsheet) ? spreadsheet : SpreadsheetApp.getActive();
  inputSheet = ( inputSheet ) ? inputSheet : spreadsheet.getSheetByName( SHEETS.INPUT.NAME );
  myFilesSheet = (myFilesSheet) ? myFilesSheet : spreadsheet.getSheetByName( SHEETS.MY_FILES.NAME );
  ui = (ui) ? ui : SpreadsheetApp.getUi();
  fileIdsToSkip = (fileIdsToSkip) ? fileIdsToSkip : new Array();

  //return array
  var fileIdsChecked = new Array();
  
  //switch UI to My Files Sheet
  if( !ui.FAKE_MODE ) { spreadsheet.setActiveSheet(myFilesSheet, true); }

  //find the row numbers where the status = unsigned
  let unsignedRowNums = findAllInColumn( myFilesSheet,
        SHEETS.MY_FILES.STATUSES.UNSIGNED, 
        SHEETS.MY_FILES.REFS.COL_NO_STATUS_BAR,
        SHEETS.MY_FILES.REFS.ROW_NO_FIRST_FILE,
        SHEETS.MY_FILES.REFS.ROW_NO_LAST_FILE );

  let rowFileId = null;
  let rowLearnerId = null;
  for (const rowNum of unsignedRowNums) {
    rowFileId = myFilesSheet.getRange( rowNum, SHEETS.MY_FILES.REFS.COL_NO_USERS_FILE_ID ).getValue();
    rowLearnerId = myFilesSheet.getRange( rowNum, SHEETS.MY_FILES.REFS.COL_NO_LEARNER_ID ).getValue();

    if( !fileIdsToSkip.includes(rowFileId) && ( !onlyTheseLearnerIds || onlyTheseLearnerIds.includes( rowLearnerId ) ) ) {
      
      //select the cell so user can follow where we are
      if( !ui.FAKE_MODE ) { 
        myFilesSheet.setActiveSelection( myFilesSheet.getRange( 
          rowNum,
          SHEETS.MY_FILES.REFS.COL_NO_STATUS_BAR ) 
        );
      }
      //spreadsheet.toast( "Checking for signatures from "+rowLearnerName+" on row " + rowNum, 'Checking Row ' + rowNum  );

      //refresh the updated date on file data on both sheets
      UpdateFileInfoInWorkbook( spreadsheet, inputSheet, myFilesSheet, null, DriveApp.getFileById( rowFileId ), false );
      fileIdsChecked.push( rowFileId );
    }
  }

  return fileIdsChecked;
}

/**
 * [lsa-utils.gs]
 * Generate PDFs for all 'To Send'/'To Save' Records from the list of record numbers passed in. For all of the Learners that we generate
 * Records for, check all of their historic unsigned files to see if they have recently been signed. and then get a list of
 * historic still-unsigned files for each learner then use this info to send 1email to each learner
 * each email lists the newly generated files and the still-unsigned historic files to prompt them to sign all
 * @param recordNumList {Array[Integer]=} Array of record numbers (or null for ALL record numbers)
 * @param spreadsheet {Spreadsheet=} The Spreadsheet
 * @param inputSheet {Sheet=} The Input Sheet for the Spreadsheet
 * @param myFilesSheet {Sheet=} The My Files Sheet for the Spreadsheet
 * @param ui {UserInterface=} The ui object to use
 */

function ExportToSendRoSsFromList( recordNumList, spreadsheet: GoogleAppsScript.Spreadsheet.Spreadsheet, 
        inputSheet, myFilesSheet, ui ) {

  //parse params
  spreadsheet = ( spreadsheet ) ? spreadsheet : SpreadsheetApp.getActive();
  inputSheet = ( inputSheet ) ? inputSheet : spreadsheet.getSheetByName(SHEETS.INPUT.NAME);
  myFilesSheet = ( myFilesSheet ) ? myFilesSheet : spreadsheet.getSheetByName(SHEETS.MY_FILES.NAME);
  ui = ( ui ) ? ui :  SpreadsheetApp.getUi();

  if( !isAuthorised_( spreadsheet, ui, null, true, true, true, false ) ) { return false; }
Logger.log( "isAuthorised_" );
  if( !recordNumList ) {
    //empty array means dont export any, null or undefined means export all of them
    recordNumList = new Array();
    for( let i = 1; i <= SHEETS.INPUT.REFS.COL_NO_RECORD_LAST - SHEETS.INPUT.REFS.COL_NO_RECORD_1 + 1 ; i++ ) {
      recordNumList.push( i );
    }
  }
Logger.log( "recordNumList = " + recordNumList );
  //switch to input sheet so user can see which record is being worked on
  if( !ui.FAKE_MODE ){ spreadsheet.setActiveSheet(inputSheet, true); }
Logger.log( "spreadsheet.getActiveSheet().getName() = " + spreadsheet.getActiveSheet().getName() ); //
  if( validateInputSheetFields( inputSheet, ui, null ) ) {

    // ***** FIRST GENERATE THE RECORDS OF SUPPORT ON THE INPUT SHEET AND KEEP TRACK OF WHCH ONES WE'VE GENRATED *****
    //keep track of what happened as we look through each of the Records of Support
    let couldntExportErrors = "";
    let learnerNamesToEmail: Array<string> = new Array();
    let learnerNamesToSkip: Array<string> = new Array();
    let learnerIdsToEmail = new Array(); //mirrors the above array
    let learnerEmailAddressesUsed = new Array();  //mirrors the above array but gets populated right at the end
    let newFileIDsToEmailFlatList = Array(); //1d array [i]
    let newFileRecordNosToEmailByFileId = {};
    let noOfRecordsGenerated = 0;
    let alreadyRefreshedLearnerSheetFromMaster = false;
    
    //loop through each of the records 1 through 25
    let recordNo = -1;
    let recordStatus = null;
    let recordLearnerName = null;
    let recordLearnerId = null;
    let exportedFileId = null;

    for( var i = 0 ; i < recordNumList.length ; i++ ) {
      recordNo = recordNumList[i];
      
      //get info about record that we need from Input Sheet
      recordStatus = inputSheet.getRange( 
            SHEETS.INPUT.REFS.ROW_NO_STATUS_BAR, 
            SHEETS.INPUT.REFS.COL_NO_RECORD_1 + recordNo - 1
          ).getValue();
      recordLearnerName = inputSheet.getRange( 
            SHEETS.INPUT.REFS.ROW_NO_LEARNER_NAME, 
            SHEETS.INPUT.REFS.COL_NO_RECORD_1 + recordNo - 1
          ).getValue();
      recordLearnerId = inputSheet.getRange( 
            SHEETS.INPUT.REFS.ROW_NO_LEARNER_ID, 
            SHEETS.INPUT.REFS.COL_NO_RECORD_1 + recordNo - 1
          ).getValue();
      
      //either export it or add a message to the error feedback for the user
      if( recordStatus == SHEETS.INPUT.STATUSES.UNSENT || recordStatus == SHEETS.INPUT.STATUSES.UNSENT_AUTOSIGN ) {

        //select the cell so user can follow where we are
        if( !ui.FAKE_MODE ) { 
          inputSheet.setActiveSelection( inputSheet.getRange( 
              SHEETS.INPUT.REFS.ROW_NO_STATUS_BAR, 
              SHEETS.INPUT.REFS.COL_NO_RECORD_1 + recordNo - 1 ) 
          );
        }

        //if first one, refresh learner data in case sign types have changed or in case sorted signature 
        //images have just come in
        if( !alreadyRefreshedLearnerSheetFromMaster ) {
          uiSensitiveToast( spreadsheet, ui, 
                "Checking for recent Stored Signature updates", 
                "Refreshng Learners" );
          RefreshSettingsLearnerSheetDataFromMasterDatabaseFromChild( spreadsheet, ui, null, false );
          alreadyRefreshedLearnerSheetFromMaster = true;
        }

        //geenrate PDF onto shared drive
        exportedFileId = null;
        try{  
          exportedFileId = ExportOneRoSRecord( recordNo, ui, learnerNamesToEmail, learnerNamesToSkip );
        }
        catch( e ) {
          debugCatchError( e );
          ui.alert( "Issue exporting Record" + recordNo, "Export failed for record " + recordNo + "\n\nThis is normally a one-off and nothing to worry about." + 
              "The record will be exported next time you Generate Records of Support.\n\nPress OK to continue.\n\n\n\n" +
              "Error Details:\n" + e,
              ui.ButtonSet.OK );
        }

        //if it worked, update the tracking variables.
        if( exportedFileId != null ) {
          noOfRecordsGenerated++;
          //add the learner and the file to to the lists of files by learner
          if( recordStatus != SHEETS.INPUT.STATUSES.UNSENT_AUTOSIGN ) {
            if( !learnerNamesToEmail.includes( recordLearnerName ) ) { 
              learnerNamesToEmail.push( recordLearnerName );
              learnerIdsToEmail.push( recordLearnerId );
            }
            newFileIDsToEmailFlatList.push( exportedFileId );
            newFileRecordNosToEmailByFileId[ exportedFileId ] = recordNo;
          }
        }
        else
        {
          couldntExportErrors += " - Record " + recordNo + " - EXPORT " + 
          ( learnerNamesToSkip.includes( recordLearnerName ) ? "SKIPPED\n" : "FAILED - please try again\n" );
        }
      }
      else if( recordStatus == SHEETS.INPUT.STATUSES.SAVED_EMAILWAIT ) {
        //this is a previously saved record where the email failed to send so we'll pick this up as if weve just saved it
        //adding it to the list of files and learners that we're going to email

        //add the learner and the file to to the lists of files by learner
        if( !learnerNamesToEmail.includes( recordLearnerName ) ) { 
              learnerNamesToEmail.push( recordLearnerName );
              learnerIdsToEmail.push( recordLearnerId );
            }
        let tempFileId = inputSheet.getRange( 
            SHEETS.INPUT.REFS.ROW_NO_FILE_ID, 
            SHEETS.INPUT.REFS.COL_NO_RECORD_1 + recordNo - 1
          ).getValue();
        newFileIDsToEmailFlatList.push( tempFileId );
        newFileRecordNosToEmailByFileId[ tempFileId ] = recordNo;
      }
      else if(  recordStatus != SHEETS.INPUT.STATUSES.SIGNED && recordStatus != SHEETS.INPUT.STATUSES.SIGNED_AUTOSIGN &&
                recordStatus != SHEETS.INPUT.STATUSES.UNSIGNED && recordStatus != SHEETS.INPUT.STATUSES.BLANK ) {
        //if it's none of the usual statuses then it must be an error message
        couldntExportErrors += " - Record " + recordNo + " - " + recordStatus + "\n";
      }
    }

    //VARIABLES PASSED ON TO NEXT SECTION
    //couldntExportErrors           //Logger.log( "couldntExportErrors = " + couldntExportErrors );
    //learnerNamesFound             //Logger.log( "learnerNamesFound = " + learnerNamesFound );
    //newFileIDsFlatList            //Logger.log( "newFileIDsFlatList = " + newFileIDsFlatList );
    //newFileRecordNosByFileId      //Logger.log( "newFileRecordNosByFileId = " + newFileRecordNosByFileId );

    if( learnerNamesToEmail && learnerNamesToEmail.length > 0 ) {

      // ***** NEXT CHECK FOR SIGNATURES ON INPUT SHEET (SKIP THE JUST GENERATED FILES TO SAVE TIME) *****
      uiSensitiveToast( spreadsheet, ui, 
              "Checking for today's signatures from these learners:\n"+learnerNamesToEmail, 
              "Checking Today's Signatures" );
      let fileIdsCheckedForSignatures = CheckForSignatures_InputSheet( spreadsheet, inputSheet, ui, 
                                                                       newFileIDsToEmailFlatList, learnerNamesToEmail);

      //VARIABLES PASSED ON TO NEXT SECTION
      //couldntExportErrors           //Logger.log( "couldntExportErrors = " + couldntExportErrors );
      //learnerNamesFound             //Logger.log( "learnerNamesFound = " + learnerNamesFound );
      //newFileIDsFlatList            //Logger.log( "newFileIDsFlatList = " + newFileIDsFlatList );
      //newFileRecordNosByFileId      //Logger.log( "newFileRecordNosByFileId = " + newFileRecordNosByFileId );
      //fileIdsCheckedForSignatures   //Logger.log( "fileIdsCheckedForSignatures = " + fileIdsCheckedForSignatures );

      // ***** NEXT CHECK FOR SIGNATURES ON MY FILES SHEET (SKIP THE JUST GENERATED FILES AND JUST CHECKED FILES TO SAVE TIME) *****
      uiSensitiveToast( spreadsheet, ui, 
        "Checking for signatures on old Records from these learners:\n"+learnerNamesToEmail, 
        "Checking Historic Signatures", 10 );
      let fileIdsToExclude = CloneArray_ShallowCopy( newFileIDsToEmailFlatList ).concat( fileIdsCheckedForSignatures );
      fileIdsCheckedForSignatures = fileIdsCheckedForSignatures.concat( 
            CheckForSignatures_MyFilesSheet( spreadsheet, inputSheet, myFilesSheet, ui, fileIdsToExclude, learnerIdsToEmail ) 
      );

      //VARIABLES PASSED ON TO NEXT SECTION
      //couldntExportErrors           
      Logger.log( "couldntExportErrors = " + couldntExportErrors );
      //learnerNamesToEmail             
      Logger.log( "learnerNamesToEmail = " + learnerNamesToEmail );
       //learnerIdsToEmail             
      Logger.log( "learnerIdsToEmail = " + learnerIdsToEmail );
      //newFileIDsFlatList            
      Logger.log( "newFileIDsFlatList = " + newFileIDsToEmailFlatList );
      //newFileRecordNosByFileId      
      Logger.log( "newFileRecordNosByFileId = " + newFileRecordNosToEmailByFileId );
      //fileIdsCheckedForSignatures   
      Logger.log( "fileIdsCheckedForSignatures = " + fileIdsCheckedForSignatures );

      // ***** NOW QUERY THE MY FILES SHEET STATUS COLUMN TO POPULATE ARRAYS OF ONES ARE STILL UNSIGNED BY LEAERNER ID *****
      uiSensitiveToast( spreadsheet, ui,  "Collecting data ready for sending emails", "Preparing Emails", 15 );

      let newFilesByLearnerId = new Array(); //2d array, [learner name][i]
      let oldUnsignedFilesByLearnerId = new Array(); //2d array, [learner name][i]

      const dataRowOffset = SHEETS.MY_FILES.REFS.COL_NO_FILE_NAME;
      let myFilesRow = null;
      let thisStatus = null;
      let thisFileId = null;
      let thisLearnerId = null;
      let thisFileStruct = null;
      let mostRecentSignType = null;

      var allMyFilesData = myFilesSheet.getRange( SHEETS.MY_FILES.REFS.ROW_NO_FIRST_FILE, SHEETS.MY_FILES.REFS.COL_NO_FILE_NAME,
                                                  SHEETS.MY_FILES.REFS.ROW_NO_LAST_FILE - SHEETS.MY_FILES.REFS.ROW_NO_FIRST_FILE + 1,
                                                  SHEETS.MY_FILES.REFS.COL_NO_SIGN_TYPE - SHEETS.MY_FILES.REFS.COL_NO_FILE_NAME + 1 
                                                ).getValues(); 

      for ( let rowIndex = 0; rowIndex < allMyFilesData.length; rowIndex++ ) {
        myFilesRow = allMyFilesData[rowIndex];
        thisFileId = myFilesRow[ SHEETS.MY_FILES.REFS.COL_NO_USERS_FILE_ID - dataRowOffset ];

        if( newFileIDsToEmailFlatList.includes( thisFileId ) ) {
          //WE'VE FOUND A FILE THAT WAS JUST CREATED SO IT GOES IN THE JUST CREATED LIST (WE KNOW ITS UNSIGNED AS ITS JUST CREATED)
          //create an object that represents the file
          thisFileStruct = {
            id             : thisFileId,
            lessonName     : myFilesRow[ SHEETS.MY_FILES.REFS.COL_NO_LESSON_NAME - dataRowOffset ],
            lessonTimesObj : getlessonTimesObjectFromMyFilesSheetValues( 
              myFilesRow[ SHEETS.MY_FILES.REFS.COL_NO_LESSON_DATE - dataRowOffset ],
              myFilesRow[ SHEETS.MY_FILES.REFS.COL_NO_START_TIME - dataRowOffset ],
              myFilesRow[ SHEETS.MY_FILES.REFS.COL_NO_DURATION - dataRowOffset ]
            ),
            signType       : myFilesRow[ SHEETS.MY_FILES.REFS.COL_NO_SIGN_TYPE - dataRowOffset ],
            email          : myFilesRow[ SHEETS.MY_FILES.REFS.COL_NO_LEARNER_EMAIL  - dataRowOffset ],
            url            : DriveApp.getFileById( thisFileId ).getUrl(),
            recordNo       : newFileRecordNosToEmailByFileId[ thisFileId ]
          };
          //add the file object to the 2D list of files by Learner Name
          thisLearnerId = myFilesRow[ SHEETS.MY_FILES.REFS.COL_NO_LEARNER_ID - dataRowOffset ];
          if( !newFilesByLearnerId[ ""+thisLearnerId ] ) { newFilesByLearnerId[ ""+thisLearnerId ] = new Array(); }
          newFilesByLearnerId[ ""+thisLearnerId ].push( thisFileStruct );
        }
        else if( fileIdsCheckedForSignatures.includes( thisFileId ) ) {
          //WE'VE FOUND A FILE THAT WAS JUST CHECKED FOR UPDATES SO IT GOES IN THE OLD FILES LIST (BUT ONLY IF STILL UNSIGNED)
          thisStatus = myFilesSheet.getRange( SHEETS.MY_FILES.REFS.ROW_NO_FIRST_FILE + rowIndex, 
                                              SHEETS.MY_FILES.REFS.COL_NO_STATUS_BAR ).getValue();
          if( thisStatus == SHEETS.MY_FILES.STATUSES.UNSIGNED ) {
            //create an object that represents the file
            thisFileStruct = {
              id             : thisFileId,
              lessonName     : myFilesRow[ SHEETS.MY_FILES.REFS.COL_NO_LESSON_NAME - dataRowOffset ],
              lessonTimesObj : getlessonTimesObjectFromMyFilesSheetValues( 
                myFilesRow[ SHEETS.MY_FILES.REFS.COL_NO_LESSON_DATE - dataRowOffset ],
                myFilesRow[ SHEETS.MY_FILES.REFS.COL_NO_START_TIME - dataRowOffset ],
                myFilesRow[ SHEETS.MY_FILES.REFS.COL_NO_DURATION - dataRowOffset ]
              ),
              signType       : myFilesRow[ SHEETS.MY_FILES.REFS.COL_NO_SIGN_TYPE - dataRowOffset ],
              email          : myFilesRow[ SHEETS.MY_FILES.REFS.COL_NO_LEARNER_EMAIL  - dataRowOffset ],
              url            : DriveApp.getFileById( thisFileId ).getUrl()
            };
            //add the file object to the 2D list of files by Learner Name
            thisLearnerId = myFilesRow[ SHEETS.MY_FILES.REFS.COL_NO_LEARNER_ID - dataRowOffset ];
            if( !oldUnsignedFilesByLearnerId[ ""+thisLearnerId ] ) { oldUnsignedFilesByLearnerId[ ""+thisLearnerId ] = new Array(); }
            oldUnsignedFilesByLearnerId[ ""+thisLearnerId ].push( thisFileStruct );
          }
        }
      }


      //VARIABLES PASSED ON TO NEXT SECTION
      //couldntExportErrors           //Logger.log( "couldntExportErrors = " + couldntExportErrors );
      //learnerNamesFound             //Logger.log( "learnerNamesFound = " + learnerNamesFound );
      //newFileIDsFlatList            //Logger.log( "newFileIDsFlatList = " + newFileIDsFlatList );
      //newFileRecordNosByFileId      //Logger.log( "newFileRecordNosByFileId = " + newFileRecordNosByFileId );
      //fileIdsCheckedForSignatures   //Logger.log( "fileIdsCheckedForSignatures = " + fileIdsCheckedForSignatures );

      // ***** NOW SEND EMAILS TO EACH OF THE LEARNERS TOUCHED ABOVE *****
      //before we switch back to input sheet load the email blobs into memory as this takes a while and the previous toast is still
      //relevant here
      EnsureLearnerEmailBlobsLoaded_(); //takes a while to run
      SpreadsheetApp.flush();  //takes a while and ensures that the spreadsheet is up to date when the user is staring at input sheet


      //Switch back view to the Input Sheet
      if( !ui.FAKE_MODE ) { spreadsheet.setActiveSheet( inputSheet, true ); };
      uiSensitiveToast( spreadsheet, ui, "Emailing the following Learners:\n" + learnerNamesToEmail, "Emailing Learners" );

      //Send Learner Emails
      learnerEmailAddressesUsed = new Array();
      var emailTemplate = HtmlService.createTemplateFromFile('html/html-email-learner-tosign');
      const lSAName = inputSheet.getRange( SHEETS.INPUT.REFS.LSA_NAME.ROW_NO,SHEETS.INPUT.REFS.LSA_NAME.COL_NO ).getValue();
      let learnerSheet = spreadsheet.getSheetByName( SHEETS.SETTINGS_LEARNERS.NAME );
      let globalSettingsSheet = spreadsheet.getSheetByName( SHEETS.GLOBAL_SETTINGS.NAME );
      let learnerWebappDeployId = globalSettingsSheet.getRange (
            SHEETS.GLOBAL_SETTINGS.REFS.ROW_NO_WEBAPP_LEARNER_DEPLOYID,
            SHEETS.GLOBAL_SETTINGS.REFS.COL_NO
      ).getValue();
      let learnerObj = null;
      let thisSendEmailAddress = null;
      let foundLearnerName = null; let foundLearnerId = null;
      let learnerFriendlyName = null;
      for( let li=0; li < learnerNamesToEmail.length; li++ ) {
        foundLearnerName = learnerNamesToEmail[ li ];
        learnerFriendlyName = foundLearnerName;
        foundLearnerId   = learnerIdsToEmail[ li ];
        //find learner email address from learner Sheet - if it doesnt exist (ie learner record deleted), then revert back 
        //to the most recent email address from the My Files sheet
        learnerObj = getChildLearnerObjByLearnerIdFromSameVersionSource_( spreadsheet, learnerSheet, foundLearnerId, null );
        if( learnerObj != null ) {
          thisSendEmailAddress = learnerObj.EMAIL_ADDRESS;
          learnerFriendlyName = ( learnerObj.NICKNAME ) ? learnerObj.NICKNAME : learnerObj.FORENAME;
        }
        else {
          //no need to check newFilesByLearner as if we've just geenrated a new file for this learner
          //then the email address will DEFFO have been found above.
          thisSendEmailAddress = oldUnsignedFilesByLearnerId[foundLearnerId][0].email;
        }
        learnerEmailAddressesUsed.push( thisSendEmailAddress );

        //send email
        SendLearnerSignEmail( spreadsheet, newFilesByLearnerId[""+foundLearnerId], oldUnsignedFilesByLearnerId[""+foundLearnerId],
                              thisSendEmailAddress, foundLearnerName, learnerFriendlyName, lSAName, ""+foundLearnerId, 
                              learnerObj.SIGN_TYPE, learnerObj.SIGNATURE_ID, false, learnerWebappDeployId, emailTemplate );

        //flag email as sent in the Input Sheet
        //FIND ALL newFilesByLearner FILES for this Learner on the Input Sheet and mark them as email sent
        let newFilesEmailed = newFilesByLearnerId[""+foundLearnerId];
        for( let nfi=0; nfi < newFilesEmailed.length; nfi++ ) {
          inputSheet.getRange( SHEETS.INPUT.REFS.ROW_NO_LEARNER_EMAILED, 
                               SHEETS.INPUT.REFS.COL_NO_RECORD_1 + newFilesEmailed[nfi].recordNo - 1 ).setValue( "Yes" );
        }
        SpreadsheetApp.flush();
      }
    }

    //VARIABLES PASSED ON TO NEXT SECTION
    //couldntExportErrors           //Logger.log( "couldntExportErrors = " + couldntExportErrors );
    //learnerNamesFound             //Logger.log( "learnerNamesFound = " + learnerNamesFound );
    //learnerEmailAddressesUsed     //Logger.log( "learnerEmailAddressesUsed = " + learnerEmailAddressesUsed );

    // ***** FEEDBACK ONCE FILES HAVE BEEN SENT *****
    //emailed lerners feedback
    let learnersEmailedFeedback = ( noOfRecordsGenerated == 0 ) ? "" : "\nThe following learners have been emailed:\n";
    for( let i=0; i<learnerNamesToEmail.length ; i++ ) {
      learnersEmailedFeedback += " - " + learnerNamesToEmail[i] + " (" + 
                                 substututeIfPlaceholderEmailAddress_( learnerEmailAddressesUsed[i] ) + ")\n";
    }

    //records skipped feedback
    let recordErrorsFeedback = ( couldntExportErrors=="" ) 
              ? "" : ( "\nThe following Records could not be generated because of missing information:\n" + couldntExportErrors );

    //feedback to user
    ui.alert(
      "Finished Sending Records", 
      "A total of " + noOfRecordsGenerated + " Record" + ( noOfRecordsGenerated == 1 ? "" : "s" ) + " of Support " +
      ( noOfRecordsGenerated == 1 ? "was" : "were" ) + " saved on the shared drive.\n" +
      learnersEmailedFeedback + 
      recordErrorsFeedback,
      ui.ButtonSet.OK);
  }
}

/** 
 * [lsa-utils.gs]
 * sheet-defs.gs contains data about default field values for some of the sheets. Use this function when you want to set the values of
 * some of the cells on a sheet to these defult values. The default field defs are split up into seperate sections to allow updating
 * around locked cells whilst writing all ranges in one go for efficiency
 * This function appends row and height data to each of the sections (each defaultDefList) passed in defaultDefListArray
 * It then returns the default data array ready to be saved later using saveDataRangeValues
 * @param sheetRefsObject {Object} from the sheet-defs listing which rows are on which row numbers
 * @param defaultDefListArray {Object} An Array of Section Field Lists, Each Section Field Lists gets annotated with row and height data
 * @return rangesArray {Array} An array of arrays of objects (one per section) holding the default values ready to write to spreadsheet
*/
function getDataArraysFromDefaultsDef( sheetRefsObject, defaultDefListArray ) {
  let defaultDefList = null;
  let defaultFieldDef = null;
  let fieldName = null;;
  let fieldDefaultValue = null;
  let row = null;
  let height = null;

  //loop through each def list and then each field in the list creating an array of default values as we go
  let rangesArray = new Array();
  let rangeFieldsArray = null;
  for( let d=0; d < defaultDefListArray.length; d++ ) {
    rangeFieldsArray = new Array();
    defaultDefList = defaultDefListArray[d];
    height = 0;
    for( let i=0; i < defaultDefList.length; i++ ) {
      defaultFieldDef = defaultDefList[i];
      fieldName = defaultFieldDef[0];
      fieldDefaultValue = defaultFieldDef[1];
      if( i==0 ) {
        row = sheetRefsObject[ 'ROW_NO_'+fieldName ];
        defaultDefList.ROW = row;
      }

      if( fieldDefaultValue === null ) { //its a list of checkboxes
        height += sheetRefsObject[ 'ROW_NO_'+fieldName.substring(0, fieldName.length-5)+"LAST" ] - row + 1;
        for( let j=0; j<height; j++ ) { rangeFieldsArray.push( [ false ] ); }
      }
      else { //its just a single field
        height += 1;
        rangeFieldsArray.push( [ fieldDefaultValue ] );
      }
    }
    defaultDefList.HEIGHT = height;
    rangesArray.push( rangeFieldsArray )
  }

  return rangesArray;
}

/** 
 * [lsa-utils.gs]
 * See: getDataArraysFromDefaultsDef documentation for more information
 * This function saves the default data array sections back to the sheet at column columnNo using saveDataRangeValues
 * @param defaultDefListArray {Object} An Array of Section Field Lists, Each Section Field Lists gets annotated with row and height data
 * @param rangesValuesArray {Array<Array<Array<Object>>>} An array of arrays of objects (one per section) holding the default values ready to write
 * @param sheet {GoogleAppsScript.Spreadsheet.Sheet} the sheet to write to
 * @param columnNo {number} the column number to write to
 * @param upUntilThisColumnNo {number=} OPTIONAL If you use this parameter it will update a 2d range of many columns all at once
 * @param rowOffset {number} If the form you are saving to is lower (+) or higher (-) than the data definitions, 0 if not.
*/
function saveDataRangeValues( defaultDefListArray, 
        rangesValuesArray: Array<Array<Array<Object>>>, sheet: GoogleAppsScript.Spreadsheet.Sheet, 
        columnNo: number, upUntilThisColumnNo: number|null, rowOffset: number ) {
  //Parse Params
  upUntilThisColumnNo = (upUntilThisColumnNo) ? upUntilThisColumnNo : columnNo;

  let defaultDefList = null;
  for( let d=0; d < defaultDefListArray.length; d++ ) {
    defaultDefList = defaultDefListArray[d];
    for( let i=columnNo; i < upUntilThisColumnNo; i++ ) {  //duplicate values if >1 column
      for( let f=0; f<rangesValuesArray[d].length; f++ ) {
        rangesValuesArray[d][f].push( rangesValuesArray[d][f][0] );
      }
    }

    sheet.getRange( 
          defaultDefList.ROW+rowOffset, columnNo, 
          defaultDefList.HEIGHT, upUntilThisColumnNo-columnNo + 1 
    ).setValues( rangesValuesArray[d] ); 
  }
}

/** 
 * [lsa-utils.gs]
 * sheet-defs.gs contains data about default field values for some of the sheets. Use this function when you want to copy the values
 * for these fields from one sheet to another (or from one column to another)
 * The default field defs are split up into seperate sections to allow updating
 * around locked cells whilst writing all ranges in one go for efficiency. The calso contain default value definitions
 * but these are ignored as we are not using default values we are reading real values
 * This function appends row and height data to each of the sections (each defaultDefList) passed in defaultDefListArray
 * It then returns the data array array ready to be saved elsewhere later using saveDataRangeValues
 * @param sheetRefsObject {Object} from the sheet-defs listing which rows are on which row numbers
 * @param sheet {GoogleAppsScript.Spreadsheet.Sheet} The sheet to load the data from
 * @param columnNo {number} The number fo the column to load the data from
 * @param defaultDefListArray {Array} An Array of Section Field Lists, Each Section Field Lists gets annotated with row and height data
 * @param rowOffset {number} If the form you are loading from is lower (+) or higher (-) than the data definitions, 0 if not.
 * @return {Array<Array<Array<object>>>} An array of 2D arrays of objects (one per section) holding the default values ready to write to spreadsheet
*/
function loadDataArraysFromSheet( sheetRefsObject: Object, sheet: GoogleAppsScript.Spreadsheet.Sheet, 
      columnNo: number, defaultDefListArray, rowOffset: number ) {
  let thisDefList = null;
  let thisFieldDef: Array<string> = null;
  let fieldName: string = "";
  let fieldDefaultValue: Object|null = null;
  let row: number = -1;
  let height: number = -1;

  //loop through each def list and then each field in the list creating an array of default values as we go
  let valueRangesArray: Array<Array<Array<object>>> = new Array();
  let sheetRangeArray: Array<GoogleAppsScript.Spreadsheet.Range> = new Array();
  let rangeFieldsArray = null;
  for( let d=0; d < defaultDefListArray.length; d++ ) {
    rangeFieldsArray = new Array();
    thisDefList = defaultDefListArray[d];
    height = 0;
    for( let i=0; i < thisDefList.length; i++ ) {
      thisFieldDef = thisDefList[i];
      fieldName = thisFieldDef[0];
      fieldDefaultValue = thisFieldDef[1];
      if( i==0 ) {
        row = sheetRefsObject[ 'ROW_NO_'+fieldName ];
        thisDefList.ROW = row;
      }

      if( fieldDefaultValue === null ) { //its a list of checkboxes
        height += sheetRefsObject[ 'ROW_NO_'+fieldName.substring(0, fieldName.length-5)+"LAST" ] - row + 1;
      }
      else { //its just a single field
        height += 1;
      }
    }
    thisDefList.HEIGHT = height;
    sheetRangeArray.push( sheet.getRange( thisDefList.ROW + rowOffset, columnNo, height, 1 ) );
  }

  //now read all data in one go
  for( let r=0; r<sheetRangeArray.length; r++ ) {
    valueRangesArray.push( sheetRangeArray[r].getValues() );
  }

  return valueRangesArray;
}

interface lessonTimes {
  start: Date,
  end: Date,
  duration: number
};



function getlessonTimesObjectFromInputSheetValues( lessonDate: Date, startTime: string, duration: string ) {
  let lessonTimes: lessonTimes = {
      start: null ,
      end: null,
      duration: 0,
  }

  let dateString = Utilities.formatDate( lessonDate, GLOBAL_CONSTANTS.TIMEZONE, "dd MMMM yyyy" );

  //start time
  lessonTimes.start = new Date( dateString + " " + startTime.slice(0, -1) + " " + startTime.slice(-1)+"M" );

  //duration
  let indexOfH = duration.indexOf( "h" );
  let indexOfM = duration.indexOf( "m" );
  if( indexOfH >= 0) {
    lessonTimes.duration += 60 * parseInt( duration.substr( indexOfH-1, 1 ), 10 );
  }
  if( indexOfM >= 0) {
    lessonTimes.duration += parseInt( duration.substr( indexOfM-2, 2 ), 10 );
  }

  //end time
  lessonTimes.end = new Date( lessonTimes.start.getTime() + lessonTimes.duration * 60000 );

  return lessonTimes;
}


function getlessonTimesObjectFromMyFilesSheetValues( lessonDate: Date, startTime: Date, duration ) {

  let startTimeMillis = lessonDate.getTime() + startTime.getHours() * 1000*60*60 + startTime.getMinutes() * 1000*60 +
        startTime.getSeconds() * 1000 + startTime.getMilliseconds();
  let endTimeMillis   = startTimeMillis + duration * 1000 * 60;

  let lessonTimes: lessonTimes = {
    start: new Date( startTimeMillis ),
    end: new Date( endTimeMillis ),
    duration: duration,
  };

  return lessonTimes;
}