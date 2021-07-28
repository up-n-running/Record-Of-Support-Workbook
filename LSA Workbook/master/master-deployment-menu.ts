function DeployNewVersion() {
  var spreadsheet = SpreadsheetApp.getActive();
  var ui = SpreadsheetApp.getUi();
  var globalSettingsSheet = spreadsheet.getSheetByName( SHEETS.GLOBAL_SETTINGS.NAME );
  var lsasSheet = spreadsheet.getSheetByName( SHEETS.MASTER_LSAS.NAME );
  var upgradeScriptSheet = spreadsheet.getSheetByName( SHEETS.MASTER_UPGRADE_SCRIPT.NAME );
  var learnersSheet = spreadsheet.getSheetByName( SHEETS.SETTINGS_LEARNERS.NAME );
  var lessonsSheet = spreadsheet.getSheetByName( SHEETS.SETTINGS_LESSONS.NAME );

  //RUN DIAGNOSTICS AND CONFIRM WITH THE USER THAT THEY ARE HAPPY TO DEPLOY

  //get some diagnostics data ready for the delopment checklist alert to let the dev user perform a sanity check before deciding
  //to commit to the deployment
  var thisMasterFile = DriveApp.getFileById( SpreadsheetApp.getActiveSpreadsheet().getId() );
  var masterFileCreatedDate = thisMasterFile.getDateCreated();
  masterFileCreatedDate = new Date(masterFileCreatedDate.setHours(0,0,0,0));
  
  //get proposed new version
  var proposedNewVersion = globalSettingsSheet.getRange( 
      SHEETS.GLOBAL_SETTINGS.REFS.ROW_NO_VERSION_NO, 
      SHEETS.GLOBAL_SETTINGS.REFS.COL_NO
    ).getValue();

  //get a list of all version numbers currently in use by all the LSAs
  var allVersionNumbers2D = lsasSheet.getRange( 
        SHEETS.MASTER_LSAS.REFS.ROW_NO_FIRST_LSA,
        SHEETS.MASTER_LSAS.REFS.COL_NO_USERS_VERSION,
        SHEETS.MASTER_LSAS.REFS.ROW_NO_LAST_LSA - SHEETS.MASTER_LSAS.REFS.ROW_NO_FIRST_LSA + 1,
        1 ).getValues();
  //convert 2d array to 1d array and remove blanks and deduplicate, all at the same time
  let allVersionNumbers = [];
  allVersionNumbers2D.sort();
  for(var n in allVersionNumbers2D) {
    if( allVersionNumbers2D[n][0] != "" && ( allVersionNumbers.length == 0 || 
        allVersionNumbers[allVersionNumbers.length-1]!=allVersionNumbers2D[n][0] ) ) {
          allVersionNumbers.push(allVersionNumbers2D[n][0]);
    }
  }
  //make sure the new version is not in the list of existing versions
  var warning_updateScript = (allVersionNumbers.indexOf(proposedNewVersion)+1) ? 
        ". . . *********** PROPOSED VERSION NUMBER ALREADY IN USE ***********\n" : "";
  //find out how many upgrade script records there are for each of the necessary version numbers
  var upgradeScriptHelpText = "";
  var linesInScript = -1
  var minLinesPerVersion = 999;
  var versionWithMinLines = "";
  allVersionNumbers.push( proposedNewVersion );
  for(var n in allVersionNumbers) {
    linesInScript = findAllInColumn( upgradeScriptSheet, allVersionNumbers[n], SHEETS.MASTER_UPGRADE_SCRIPT.REFS.COL_NO_SOURCE_VERSION,
          SHEETS.MASTER_UPGRADE_SCRIPT.REFS.ROW_NO_FIRST_SCRIPT_ROW, SHEETS.MASTER_UPGRADE_SCRIPT.REFS.ROW_NO_LAST_SCRIPT_ROW ).length;
    upgradeScriptHelpText += ". . . v" + allVersionNumbers[n] + " script: " + linesInScript + " lines.\n";
    if( linesInScript < minLinesPerVersion ) {
      versionWithMinLines = allVersionNumbers[n];
      minLinesPerVersion = linesInScript;
    }
  }
  allVersionNumbers.pop();
  warning_updateScript += ( minLinesPerVersion < 5 ) ? 
        ". . . *********** ONLY " + minLinesPerVersion + " LINES IN UPGRADE SCRIPT FOR V" + versionWithMinLines + " ***********\n" : "";

  //get a list of the owner and all the editors for the master spreadsheet
  var ownerAndEditorsHelpText = "";
  var fileEditors = thisMasterFile.getEditors();
  for( var i = 0; i<fileEditors.length; i++ ) {
    ownerAndEditorsHelpText += ". . . Editor " + (i+1) + ": " + fileEditors[i].getName() + "\n";
  }
 
  //work out when the RoS Form Template was last updated and see if it was updated since file was created
  var lastEdit_MasterTemplate = globalSettingsSheet.getRange( 
        SHEETS.GLOBAL_SETTINGS.REFS.ROW_NO_LAST_TEMPLATE_EDIT, 
        SHEETS.GLOBAL_SETTINGS.REFS.COL_NO
      ).getValue();
  var warning_MasterTemplate = masterFileCreatedDate <= lastEdit_MasterTemplate ? 
        ". . . **** MASTER HAS BEEN UPDATED - BE SURE TO CASCADE CHANGES ****\n" : "";

  //ensure no test data is there in the spreadsheet
  //i.e. ensure there is only 1 line on Settings - Learners and Settings - Lessons sheets 
  var noOfBlankLearners = findAllInColumn( learnersSheet, "", SHEETS.SETTINGS_LEARNERS.REFS.COL_NO_LEARNER_NAME,
          SHEETS.SETTINGS_LEARNERS.REFS.ROW_NO_FIRST_LEARNER, SHEETS.SETTINGS_LEARNERS.REFS.ROW_NO_LAST_LEARNER ).length;
  var noOfLearners = SHEETS.SETTINGS_LEARNERS.REFS.ROW_NO_LAST_LEARNER -  SHEETS.SETTINGS_LEARNERS.REFS.ROW_NO_FIRST_LEARNER -     
                 noOfBlankLearners + 1;
  var noOfBlankLessons = findAllInColumn( lessonsSheet, "", SHEETS.SETTINGS_LESSONS.REFS.COL_NO_LESSON_NAME,
          SHEETS.SETTINGS_LESSONS.REFS.ROW_NO_FIRST_LESSON, SHEETS.SETTINGS_LESSONS.REFS.ROW_NO_LAST_LESSON ).length;
  var noOfLessons = SHEETS.SETTINGS_LESSONS.REFS.ROW_NO_LAST_LESSON -  SHEETS.SETTINGS_LESSONS.REFS.ROW_NO_FIRST_LESSON -     
                 noOfBlankLessons + 1;
  var warning_TestData = ( noOfLearners != 0 || noOfLessons != 0 ) ? 
        ". . . *********** YOU MUST REMOVE TEST DATA AND REPLACE WITH SAMPLE DATA ***********\n" : "";

  //NOW BUILD A MASSIVE DEPLOYMENT CHECKLIST TOGETHER WITH DIAGNOSTIC TEST RESULTS AND WARNINGS
  //READY TO SHOW THE USER AND CONFIRM IF THEY WANT TO DEPLOY
  var deploymentPrepTxt = "Please review the following checklist before deploying a new version:\n\n" +
          "- Have you entered the Release Notes & Version Number in the '"+SHEETS.GLOBAL_SETTINGS.NAME+"' sheet?\n" +
          "\n\n- Have you added the new version's Update Script to the '"+SHEETS.MASTER_UPGRADE_SCRIPT.NAME+"' sheet?\n" +
          "\n. . . New version proposed: " + proposedNewVersion + "\n" +
          ". . . Current versions in use: " + allVersionNumbers + "\n" +
          "\n" + upgradeScriptHelpText +
          warning_updateScript +
          "\n\n- If you have made a new copy of the Master Workbook file to make your changes then:\n" +
          "- - Have you RECENTLY copied across the '"+SHEETS.MASTER_LSAS.NAME+"' sheet from the current Master (Do Not Leave Test Area IDs)?\n" +
          "- - Have you RECENTLY copied across the '"+SHEETS.MASTER_LEARNERS.NAME+"' sheet from the current Master (Do Not Leave Test Area IDs)?\n" +
          "- - Do you need to copy across any of the '"+SHEETS.MASTER_HELP.NAME+"' sheet from the current Master?\n" +
          "- - Is the file owned by the dev account?\n" +
          "- - Have you changed the 'ID for RoS Folder in Shared Drive:', 'Deployment ID of the Web App accessible by all LSAs:' and the 'ID for LSA Worksheet Folder in Shared Drive:' " +
                  "settings back to the live directories? (should not be '1nHEVWajwavTeu7Fi7I8JCvGqrT-341fA', 'AKfycbwb3ojVvCJGxfG5UaTFoba1q1qiudMoUn4X3hZMDXKoEyYC_vgtI_W3MGi28RMFuscd' or " + 
                  "'1Cr86CaeXqFDTLaNEmVkgcIKPEJt2fg2e'.\n" +
          "- - Have you shared the file to the LSA Admins in G Drive file sharing?\n" +
          "\n" + ownerAndEditorsHelpText +
          "\n\n- Remember to hide hidden Rows on "+SHEETS.INPUT.NAME+" Sheet and hidden columns on "+SHEETS.MY_FILES.NAME+" Sheet, " +
          "on "+SHEETS.SETTINGS_LEARNERS.NAME+" Sheet, the "+SHEETS.SETTINGS_LESSONS.NAME+" Sheet and the 2 Mobile Sheets." +
          "\n\n- If you have made changes to the '"+SHEETS.MASTER_TEMPLATE.NAME+
                  "' sheet, have you cascaded these changes down to the 25 RoS worksheets?\n" +
          "\n. . . Last  '"+SHEETS.MASTER_TEMPLATE.NAME+"' edit: " + 
                  Utilities.formatDate(lastEdit_MasterTemplate, GLOBAL_CONSTANTS.TIMEZONE, "EEEE, d MMMM") + "\n" +
          ". . . Master file created: " + Utilities.formatDate(masterFileCreatedDate, GLOBAL_CONSTANTS.TIMEZONE, "EEEE, d MMMM") + "\n" +
          warning_MasterTemplate +
          "\n\n- Have you removed any test data and replaced with sample data ready for deployment?\n" +
          "\n. . . No of Learner Records: " + noOfLearners + "\n" +
          ". . . No of Lesson Records: " + noOfLessons + "\n" +
          warning_TestData +
          "\n\n- Have you updated the script properties of the Learner webapp (esp the master File ID)?" +
          "\n\n- Have you deployed the LIVE (not just TEST) version of the Staff Webapp to the latest code?" +
          "\n\n- Have you uninstalled all instances of the mobile service so that the Global Setting reads: FALSE?" +
          "\n_____________________________\n\nARE YOU READY TO DEPLOY?";
  var confirmButton = ui.alert( "Are you ready to deploy version "+proposedNewVersion+"?", deploymentPrepTxt, ui.ButtonSet.OK_CANCEL );
  if( confirmButton == ui.Button.OK ) {
    //last check, confirm they are happy with the release notes
    var upgradeAvailablePrefix = globalSettingsSheet.getRange( 
        SHEETS.GLOBAL_SETTINGS.REFS.ROW_NO_UPDATE_AVAILABLE_TXT, 
        SHEETS.GLOBAL_SETTINGS.REFS.COL_NO
      ).getValue();
    var releaseNotes = globalSettingsSheet.getRange( 
        SHEETS.GLOBAL_SETTINGS.REFS.ROW_NO_RELEASE_NOTES, 
        SHEETS.GLOBAL_SETTINGS.REFS.COL_NO
      ).getValue();
    var announcementText = upgradeAvailablePrefix + "\n" + releaseNotes;
    confirmButton = ui.alert( "Approve 'Upgrade Available' Announcement?", announcementText + "\n_____________________________\n\n" +
                                  "PRESS OK TO DEPLOY NOW", ui.ButtonSet.OK_CANCEL );
    if( confirmButton == ui.Button.OK ) {
      //USER HAS CONFIRMED THEY ARE HAPPY TO DEPLOY SO LETS GO...
      PushToChildWorksheets( null, thisMasterFile.getId(), announcementText, null, false, false, false );
    }
  
  }
}

function RegenerateFromMasterTemplate() {
  
  let spreadsheet = SpreadsheetApp.getActive();  
  let masterTemplateSheet = spreadsheet.getSheetByName( SHEETS.MASTER_TEMPLATE.NAME );
  
  //Make sure left 2 cols are hidden on master Sheet
  masterTemplateSheet.hideColumns( 1, SHEETS.INPUT.COPY_TO_ROS_SHEETS.START_COPY_COL );
  
  //make sure all children are deleted before we regenarate them
  for(var i = 1; i <= 25; i++) {
    var tmpSheetToDelete = spreadsheet.getSheetByName(''+i);
    if( tmpSheetToDelete !== null ) {
      spreadsheet.deleteSheet( tmpSheetToDelete );
    }
  }
  
  //Regenrate sheets 1 thru 25
  var newSheet = null;
  for(var i = 1; i <= 25; i++) {
    newSheet = createRoSSheetFromMasterTemplate_( spreadsheet, masterTemplateSheet, i );
    if( i>5 ) {
      newSheet.hideSheet();
    }
  }

  SpreadsheetApp.flush();
}

function createRoSSheetFromMasterTemplate_( 
      spreadsheet: GoogleAppsScript.Spreadsheet.Spreadsheet, 
      masterTemplateSheet: GoogleAppsScript.Spreadsheet.Sheet, 
      recordNo: number ) {
  
  //parse params
  spreadsheet = ( spreadsheet ) ? spreadsheet : SpreadsheetApp.getActive();
  masterTemplateSheet = ( masterTemplateSheet ) ? masterTemplateSheet : spreadsheet.getSheetByName(SHEETS.MASTER_TEMPLATE.NAME);
  
  //toast
  spreadsheet.toast( "Generating Missing RoS Preview Sheet '" + recordNo + "' for Record "+recordNo, "Missing Sheet: " + recordNo, 8 );

  //duplicate sheet
  let newSheet = masterTemplateSheet.copyTo( spreadsheet ).setName(''+recordNo);
  
  //create formulas for fields to link back to Input Sheet
  refreshRoSSheetsLinkToInputSheet_( recordNo, newSheet, spreadsheet );

  DuplicateRangeLevelProtection( masterTemplateSheet, newSheet );
  //manually do a SpreadsheetApp.flush() after calling this function to ensure up-to-date data

  return newSheet;
}


function refreshRoSSheetsLinkToInputSheet_( recordNo: number, sheet: GoogleAppsScript.Spreadsheet.Sheet|null,
        spreadsheet: GoogleAppsScript.Spreadsheet.Spreadsheet|null ) {

  //parse params
  spreadsheet = ( spreadsheet ) ? spreadsheet : SpreadsheetApp.getActive();
  sheet = ( sheet ) ? sheet : spreadsheet.getSheetByName( "" + recordNo );

  //prepare to create formulas for fields to link back to Input Sheet
  let formulaPrefix = SHEETS.INPUT.NAME+"!$"+columnToLetter( recordNo + SHEETS.INPUT.REFS.COL_NO_RECORD_1 - 1 )+"$"; //eg Input!$B$
  const copyDefsStruct = SHEETS.INPUT.COPY_TO_ROS_SHEETS;
  let col = copyDefsStruct.START_COPY_COL;
  let row = copyDefsStruct.START_COPY_ROW;

  //copy individual fields
  for( let i=0; i < copyDefsStruct.FIELDS.length; i++ ) {
    Logger.log( "Sheet "+recordNo+", Ros Row " + row + " ROW_NO_"+copyDefsStruct.FIELDS[i] + " is " + formulaPrefix+SHEETS.INPUT.REFS[ "ROW_NO_"+copyDefsStruct.FIELDS[i] ] );
    sheet.getRange( row, col ).setFormula( formulaPrefix+SHEETS.INPUT.REFS[ "ROW_NO_"+copyDefsStruct.FIELDS[i] ] );
    row++;
  }

  //copy range fields
  for( let i=0; i < copyDefsStruct.RANGES.length; i++ ) {
    let startRow = SHEETS.INPUT.REFS[ "ROW_NO_"+copyDefsStruct.RANGES[i]+"_FIRST" ];
    let endRow   = SHEETS.INPUT.REFS[ "ROW_NO_"+copyDefsStruct.RANGES[i]+"_LAST"  ];
    for( let r=startRow; r<=endRow; r++ ) {
      sheet.getRange( row, col ).setFormula( formulaPrefix+r );
      row++;
    }
  }

}


function ClearTestDataReadyForDeployment() {
  let spreadsheet = SpreadsheetApp.getActive();

  let inputSheet = spreadsheet.setActiveSheet(spreadsheet.getSheetByName(SHEETS.INPUT.NAME), true);
  let mobileMainSheet = spreadsheet.getSheetByName( SHEETS.MOBILE_MAIN.NAME );
  let mobileInputSheet = spreadsheet.getSheetByName( SHEETS.MOBILE_THIS_RECORD.NAME );

  clearDownRoSRecordsOnInputSheet( inputSheet );
  clearDownRoSRecordsOnMobileInputSheet( mobileInputSheet );
  setRecordDateFields( "", inputSheet, mobileMainSheet );
  inputSheet.getRange( 
        SHEETS.INPUT.REFS.LSA_NAME.ROW_NO, 
        SHEETS.INPUT.REFS.LSA_NAME.COL_NO 
  ).setValue( "MASTER SPREADSHEET" );
  inputSheet.getRange( 
        SHEETS.INPUT.REFS.LESSON_DATE.ROW_NO, 
        SHEETS.INPUT.REFS.LESSON_DATE.COL_NO 
  ).activateAsCurrentCell();

  let myFilesSheet = spreadsheet.setActiveSheet(spreadsheet.getSheetByName(SHEETS.MY_FILES.NAME), true);
  myFilesSheet.getRange( 
        SHEETS.MY_FILES.REFS.ROW_NO_FIRST_FILE, 
        SHEETS.MY_FILES.REFS.COL_NO_FILE_NAME,
        SHEETS.MY_FILES.REFS.ROW_NO_LAST_FILE     - SHEETS.MY_FILES.REFS.ROW_NO_FIRST_FILE + 1,
        SHEETS.MY_FILES.REFS.COL_NO_AUTOSIGN_CMTS - SHEETS.MY_FILES.REFS.COL_NO_FILE_NAME  + 1 
  ).clear({contentsOnly: true, skipFilteredRows: false}); 
  myFilesSheet.getRange( 
        SHEETS.MY_FILES.REFS.ROW_NO_FIRST_FILE, 
        SHEETS.MY_FILES.REFS.COL_NO_FILE_NAME,
  ).activateAsCurrentCell();

  let learnersSheet = spreadsheet.setActiveSheet(spreadsheet.getSheetByName(SHEETS.SETTINGS_LEARNERS.NAME), true);
  //between readonly grey cols and checkbox cols
  learnersSheet.getRange( 
        SHEETS.SETTINGS_LEARNERS.REFS.ROW_NO_FIRST_LEARNER,
        SHEETS.SETTINGS_LEARNERS.REFS.COL_NO_EDITABLE_NICKNAME,
        SHEETS.SETTINGS_LEARNERS.REFS.ROW_NO_LAST_LEARNER            - SHEETS.SETTINGS_LEARNERS.REFS.ROW_NO_FIRST_LEARNER     + 1,
        (SHEETS.SETTINGS_LEARNERS.REFS.COL_NO_SUPPORT_STRAT_FIRST-1)  - SHEETS.SETTINGS_LEARNERS.REFS.COL_NO_EDITABLE_NICKNAME + 1
  ).clear({contentsOnly: true, skipFilteredRows: false})
  //checkbox columns
  learnersSheet.getRange( 
        SHEETS.SETTINGS_LEARNERS.REFS.ROW_NO_FIRST_LEARNER,
        SHEETS.SETTINGS_LEARNERS.REFS.COL_NO_SUPPORT_STRAT_FIRST,
        SHEETS.SETTINGS_LEARNERS.REFS.ROW_NO_LAST_LEARNER       - SHEETS.SETTINGS_LEARNERS.REFS.ROW_NO_FIRST_LEARNER       + 1,
        SHEETS.SETTINGS_LEARNERS.REFS.COL_NO_SUPPORT_STRAT_LAST - SHEETS.SETTINGS_LEARNERS.REFS.COL_NO_SUPPORT_STRAT_FIRST + 1
  ).setValue(false).insertCheckboxes();
  //between last checkbox cols and last col
  learnersSheet.getRange( 
      SHEETS.SETTINGS_LEARNERS.REFS.ROW_NO_FIRST_LEARNER,
      SHEETS.SETTINGS_LEARNERS.REFS.COL_NO_SUPPORT_STRAT_LAST+1,
      SHEETS.SETTINGS_LEARNERS.REFS.ROW_NO_LAST_LEARNER - SHEETS.SETTINGS_LEARNERS.REFS.ROW_NO_FIRST_LEARNER           + 1,
      SHEETS.SETTINGS_LEARNERS.REFS.COL_NO_SIGNATURE_FILE_ID - (SHEETS.SETTINGS_LEARNERS.REFS.COL_NO_SUPPORT_STRAT_LAST+1) + 1
  ).clear({contentsOnly: true, skipFilteredRows: false});
  learnersSheet.getRange(
        SHEETS.SETTINGS_LEARNERS.REFS.ROW_NO_FIRST_LEARNER,
        SHEETS.SETTINGS_LEARNERS.REFS.COL_NO_EDITABLE_NICKNAME
  ).activateAsCurrentCell();

  let lessonsSheet = spreadsheet.setActiveSheet(spreadsheet.getSheetByName(SHEETS.SETTINGS_LESSONS.NAME), true);
  //between readonly grey cols and checkbox cols
  lessonsSheet.getRange( 
        SHEETS.SETTINGS_LESSONS.REFS.ROW_NO_FIRST_LESSON,
        SHEETS.SETTINGS_LESSONS.REFS.COL_NO_EQUIPMENT_USED,
        SHEETS.SETTINGS_LESSONS.REFS.ROW_NO_LAST_LESSON             - SHEETS.SETTINGS_LESSONS.REFS.ROW_NO_FIRST_LESSON   + 1,
        (SHEETS.SETTINGS_LESSONS.REFS.COL_NO_SUPPORT_STRAT_FIRST-1) - SHEETS.SETTINGS_LESSONS.REFS.COL_NO_EQUIPMENT_USED + 1
  ).clear({contentsOnly: true, skipFilteredRows: false});
  //checkbox columns
  lessonsSheet.getRange( 
        SHEETS.SETTINGS_LESSONS.REFS.ROW_NO_FIRST_LESSON,
        SHEETS.SETTINGS_LESSONS.REFS.COL_NO_SUPPORT_STRAT_FIRST,
        SHEETS.SETTINGS_LESSONS.REFS.ROW_NO_LAST_LESSON        - SHEETS.SETTINGS_LESSONS.REFS.ROW_NO_FIRST_LESSON        + 1,
        SHEETS.SETTINGS_LESSONS.REFS.COL_NO_SUPPORT_STRAT_LAST - SHEETS.SETTINGS_LESSONS.REFS.COL_NO_SUPPORT_STRAT_FIRST + 1
  ).setValue(false).insertCheckboxes();
  //between last checkbox col and last col
  lessonsSheet.getRange( 
        SHEETS.SETTINGS_LESSONS.REFS.ROW_NO_FIRST_LESSON,
        SHEETS.SETTINGS_LESSONS.REFS.COL_NO_SUPPORT_STRAT_LAST+1,
        SHEETS.SETTINGS_LESSONS.REFS.ROW_NO_LAST_LESSON - SHEETS.SETTINGS_LESSONS.REFS.ROW_NO_FIRST_LESSON           + 1,
        SHEETS.SETTINGS_LESSONS.REFS.COL_NO_LESSON_NAME - (SHEETS.SETTINGS_LESSONS.REFS.COL_NO_SUPPORT_STRAT_LAST+1) + 1
  ).clear({contentsOnly: true, skipFilteredRows: false});
  lessonsSheet.getRange(
        SHEETS.SETTINGS_LESSONS.REFS.ROW_NO_FIRST_LESSON,
        SHEETS.SETTINGS_LESSONS.REFS.COL_NO_EQUIPMENT_USED
  ).activateAsCurrentCell();

  let targetGradesSheet = spreadsheet.setActiveSheet(spreadsheet.getSheetByName(SHEETS.SETTINGS_TARGET_GRADES.NAME), true);
  targetGradesSheet.getRange( 
        SHEETS.SETTINGS_TARGET_GRADES.REFS.ROW_NO_LESSON_NAMES  + 1,
        SHEETS.SETTINGS_TARGET_GRADES.REFS.COL_NO_LEARNER_NAMES + 1,
        SHEETS.SETTINGS_TARGET_GRADES.REFS.ROW_NO_LESSON_NAMES  + 
              SHEETS.SETTINGS_LEARNERS.REFS.ROW_NO_LAST_LEARNER - SHEETS.SETTINGS_LEARNERS.REFS.ROW_NO_FIRST_LEARNER,
        SHEETS.SETTINGS_TARGET_GRADES.REFS.COL_NO_LEARNER_NAMES + 
              SHEETS.SETTINGS_LESSONS.REFS.ROW_NO_LAST_LESSON   - SHEETS.SETTINGS_LESSONS.REFS.ROW_NO_FIRST_LESSON
  ).clear({contentsOnly: true, skipFilteredRows: false});
  targetGradesSheet.getRange(
        SHEETS.SETTINGS_TARGET_GRADES.REFS.ROW_NO_LESSON_NAMES  + 1,
        SHEETS.SETTINGS_TARGET_GRADES.REFS.COL_NO_LEARNER_NAMES + 1
  ).activateAsCurrentCell();  

  let lessonTargetsSheet = spreadsheet.setActiveSheet(spreadsheet.getSheetByName(SHEETS.SETTINGS_LESSON_TARGETS.NAME), true);
  lessonTargetsSheet.getRange( 
        SHEETS.SETTINGS_LESSON_TARGETS.REFS.ROW_NO_LESSON_NAMES  + 1,
        SHEETS.SETTINGS_LESSON_TARGETS.REFS.COL_NO_LEARNER_NAMES + 1,
        SHEETS.SETTINGS_LESSON_TARGETS.REFS.ROW_NO_LESSON_NAMES  + 
              SHEETS.SETTINGS_LEARNERS.REFS.ROW_NO_LAST_LEARNER  - SHEETS.SETTINGS_LEARNERS.REFS.ROW_NO_FIRST_LEARNER,
        SHEETS.SETTINGS_LESSON_TARGETS.REFS.COL_NO_LEARNER_NAMES + 
              SHEETS.SETTINGS_LESSONS.REFS.ROW_NO_LAST_LESSON    - SHEETS.SETTINGS_LESSONS.REFS.ROW_NO_FIRST_LESSON
  ).clear({contentsOnly: true, skipFilteredRows: false});
  lessonTargetsSheet.getRange(
        SHEETS.SETTINGS_LESSON_TARGETS.REFS.ROW_NO_LESSON_NAMES  + 1,
        SHEETS.SETTINGS_LESSON_TARGETS.REFS.COL_NO_LEARNER_NAMES + 1
  ).activateAsCurrentCell();  

  spreadsheet.setActiveSheet( inputSheet, true );
};

function HideNonMasterSheetsReadyForDeployment() {
  let spreadsheet = SpreadsheetApp.getActive();

  spreadsheet.setActiveSheet( spreadsheet.getSheetByName( SHEETS.GLOBAL_SETTINGS.NAME ), true );

  if( !spreadsheet.getSheetByName( SHEETS.INPUT.NAME ).isSheetHidden() ) {
    spreadsheet.getSheetByName( SHEETS.INPUT.NAME ).hideSheet();
  }
  if( !spreadsheet.getSheetByName( SHEETS.MY_FILES.NAME ).isSheetHidden() ) {
    spreadsheet.getSheetByName( SHEETS.MY_FILES.NAME ).hideSheet();
  }
  if( !spreadsheet.getSheetByName( SHEETS.SETTINGS_LEARNERS.NAME ).isSheetHidden() ) {
    spreadsheet.getSheetByName( SHEETS.SETTINGS_LEARNERS.NAME ).hideSheet();
  }
  if( !spreadsheet.getSheetByName( SHEETS.SETTINGS_LESSONS.NAME ).isSheetHidden() ) {
    spreadsheet.getSheetByName( SHEETS.SETTINGS_LESSONS.NAME ).hideSheet();
  }
  if( !spreadsheet.getSheetByName( SHEETS.SETTINGS_TARGET_GRADES.NAME ).isSheetHidden() ) {
    spreadsheet.getSheetByName( SHEETS.SETTINGS_TARGET_GRADES.NAME ).hideSheet();
  }
  if( !spreadsheet.getSheetByName( SHEETS.SETTINGS_LESSON_TARGETS.NAME ).isSheetHidden() ) {
    spreadsheet.getSheetByName( SHEETS.SETTINGS_LESSON_TARGETS.NAME ).hideSheet();
  }
  if( !spreadsheet.getSheetByName( SHEETS.MOBILE_MAIN.NAME ).isSheetHidden() ) {
    spreadsheet.getSheetByName( SHEETS.MOBILE_MAIN.NAME ).hideSheet();
  }
  if( !spreadsheet.getSheetByName( SHEETS.MOBILE_THIS_RECORD.NAME ).isSheetHidden() ) {
    spreadsheet.getSheetByName( SHEETS.MOBILE_THIS_RECORD.NAME ).hideSheet();
  }
  for ( let i = 0; i < 25; i++) {
    if( i < 5 && spreadsheet.getSheetByName(""+(i+1)).isSheetHidden() ) {
      spreadsheet.getSheetByName(""+(i+1)).showSheet();
    } 
    else if( i >= 5 && !spreadsheet.getSheetByName(""+(i+1)).isSheetHidden() ) {
      spreadsheet.getSheetByName(""+(i+1)).hideSheet();
    } 
  }
};

function PushHelpLinkChangesToAllChildren() {
      PushToChildWorksheets( null, null, null, false, true, false, false  );
}

function PushDeployIdChangesToAllChildren() {
      PushToChildWorksheets( null, null, null, false, false, true, false  );
}

function PushAdminGlobalSettingsChangesToAllChildren() {
  PushToChildWorksheets( null, null, null, false, false, false, true );
}
