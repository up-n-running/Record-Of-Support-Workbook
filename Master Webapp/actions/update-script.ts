const MODES = {
  CREATE : { 
    CODE: 1,
    DESCRIPTION: 'Create'
  },
  UPDATE : { 
    CODE: 2,
    DESCRIPTION: 'Update'
  },
  REPAIR : { 
    CODE: 3,
    DESCRIPTION: 'Repair'
  },
};

//KEEP PERFORMING UPDATE TILL THE NEW MASTER IS PERFECT
//THEN TRY DELETING THE DIR AND CREATING AGAIN FROM SCRATCH TO SEE IF frolder.addEditor MANAGES TO ADD THE LSA AS A CONTENT MANAGER
//ADD A SHARED DRIVE PARAM TO SKIP THE TRY CATCH BLOCK WHERE IT SETS FILE EDITORS AND VIEWERS ON NEW FILES
//AND THE CHANGE OWNER LINE AT THE BOTTOM FRO UPGRADES AND REPAIRS
//AND THE createOrGetChildFolder FUNCTION
function PerformUpdates( modeCode, update_usersSpreadsheetId /* update only param */, returnData, runningAsUser,
                         //these params are for create only
                         create_userEmail, create_masterFileId ) { 

  //repair means perform update even if we're on latest version
  //null spreadsheetid means create a brand new spreadsheet from scratch


  var modeDescription =  ( modeCode == MODES.CREATE.CODE ) ? MODES.CREATE.DESCRIPTION : 
                       ( ( modeCode == MODES.UPDATE.CODE ) ? MODES.UPDATE.DESCRIPTION : MODES.REPAIR.DESCRIPTION );

  let usersSpreadsheetId = update_usersSpreadsheetId;
  let usersSpreadsheet = null;
  let users_SettingsSheet = null; 
  let usersOrigVersionNumber = -1;
  let usersEmailAddress = create_userEmail; 
  let usersMasterFileId = create_masterFileId;

  if( modeCode != MODES.CREATE.CODE ) {
    //WE ARE DOING AN UPDATE FOR AN EXISTING USER  
    //GET DETAILS FROM users SPREADSHEET's global settings sheet
    Logger.log("Fetching data from old user workbooks Global Settings. FileId: " + usersSpreadsheetId );
    usersSpreadsheet = SpreadsheetApp.openById(usersSpreadsheetId);
    users_SettingsSheet = usersSpreadsheet.getSheetByName( SHEETS.GLOBAL_SETTINGS.NAME );

    usersOrigVersionNumber = GetMasterSettingsCellFromOldVersionWorkbook( usersSpreadsheet, users_SettingsSheet, 
          SHEETS.GLOBAL_SETTINGS.REFS.HANDLE_VERSION_NO, SHEETS.GLOBAL_SETTINGS.REFS.ROW_NO_VERSION_NO ).getValue();
    
    usersMasterFileId = GetMasterSettingsCellFromOldVersionWorkbook( usersSpreadsheet, users_SettingsSheet, 
          SHEETS.GLOBAL_SETTINGS.REFS.HANDLE_MASTER_FILE_ID, SHEETS.GLOBAL_SETTINGS.REFS.ROW_NO_MASTER_FILE_ID ).getValue();

    usersEmailAddress = GetMasterSettingsCellFromOldVersionWorkbook( usersSpreadsheet, users_SettingsSheet, 
          SHEETS.GLOBAL_SETTINGS.REFS.HANDLE_MAIN_USERS_EMAIL, SHEETS.GLOBAL_SETTINGS.REFS.ROW_NO_MAIN_USERS_EMAIL ).getValue().trim();
  }

  Logger.log("Using Master File with id: " + usersMasterFileId );

  //get users child spreadsheet's linked master (should be the latest version)
  var masterFile = DriveApp.getFileById(usersMasterFileId);
  var masterSpreadsheet = SpreadsheetApp.open( masterFile );

  //now we have the latest version master open we can get the version number to display to the user to see if they want to upgrade
  //and the dev email address which we'll use later
  var master_SettingsSheet = masterSpreadsheet.getSheetByName( SHEETS.GLOBAL_SETTINGS.NAME );
  var mastersVersionNumber = master_SettingsSheet.getRange(
        SHEETS.GLOBAL_SETTINGS.REFS.ROW_NO_VERSION_NO,
        SHEETS.GLOBAL_SETTINGS.REFS.COL_NO
      ).getValue();

  if( modeCode == MODES.UPDATE.CODE && mastersVersionNumber == usersOrigVersionNumber ) {
    returnData.success = 0;
    returnData.errorMessage = "Your version is: " + usersOrigVersionNumber + "\n" +
               "The latest version is: " + mastersVersionNumber + "\n\n" +
               "There is no newer version available.";
    Logger.log( returnData.errorMessage );
  }
  else {
    
    Logger.log( modeDescription + " Starting: " +
                "User's version is: " + usersOrigVersionNumber + ", latest version is: " + mastersVersionNumber );
    
    //////////////// START UPGRADE ////////////////////

    //find the LSAs record in the latest master sheet
    var master_LSAsSheet = masterSpreadsheet.getSheetByName( SHEETS.MASTER_LSAS.NAME );
    var foundLSARowNum = findInColumn( master_LSAsSheet, usersEmailAddress, 
                                        SHEETS.MASTER_LSAS.REFS.COL_NO_LSA_EMAIL, 
                                        SHEETS.MASTER_LSAS.REFS.ROW_NO_FIRST_LSA, 
                                        SHEETS.MASTER_LSAS.REFS.ROW_NO_LAST_LSA );
    if( foundLSARowNum < 0 )
    {
      returnData.success = 0;
      returnData.errorMessage = "The Users Email Address '" + usersEmailAddress + "' was not found in Latest Master LSA List. If this email address is correct please speak to your manager to ask them to correct this.";
      Logger.log( returnData.errorMessage );
      return;
    }
    let usersLSANameFromMaster = master_LSAsSheet.getRange(
        foundLSARowNum,
        SHEETS.MASTER_LSAS.REFS.COL_NO_LSA_NAME
      ).getValue();
    Logger.log( "Found LSA Record (" + usersEmailAddress + ") in Master on row : " + foundLSARowNum +
                ", LSA's name in master is: " + usersLSANameFromMaster );

    //Get devs eaail address and the LSAs Workbook Folder from the master's global settings
    var devEmailAddress = master_SettingsSheet.getRange(
      SHEETS.GLOBAL_SETTINGS.REFS.ROW_NO_DEVS_EMAIL, 
      SHEETS.GLOBAL_SETTINGS.REFS.COL_NO 
    ).getValue();
    let lsasWorkbookRootDirId = master_SettingsSheet.getRange(
      SHEETS.GLOBAL_SETTINGS.REFS.ROW_NO_DIR_ID_LSAS, 
      SHEETS.GLOBAL_SETTINGS.REFS.COL_NO 
    ).getValue();

    // Get folder containing users spreadsheet - and rename it if it is no longer called the same thing as the LSAs Name from Master
    let usersFolder = null;
    if( usersSpreadsheet != null ) { //this means its an update or a repair not a create
      let usersSpreadsheetFile = DriveApp.getFileById( usersSpreadsheet.getId() );
      usersFolder = getLSAsFolderFromWorkbookFile( usersSpreadsheetFile, usersLSANameFromMaster, lsasWorkbookRootDirId );
      if( usersFolder != null) {
        Logger.log( "Found Existing LSA Folder With Correct Name '"+usersFolder.getName()+"' and id: " + usersFolder.getId() );
      }
      else {
        usersFolder = getLSAsFolderFromWorkbookFile( usersSpreadsheetFile, null, lsasWorkbookRootDirId );
        if( usersFolder != null) {
          Logger.log( "Found Existing LSA Folder With Wrong Name '"+usersFolder.getName()+"' and id: " + usersFolder.getId() );
          usersFolder.setName( usersLSANameFromMaster );
          Logger.log( "Renamed To '"+usersFolder.getName()+"', id still = " + usersFolder.getId() );
        }
        else {
          Logger.log( "IMPORTANT: COULD NOT FIND LSAs DIRECTORY UNDER ROOT LSAS DIR WITH ID: '"+lsasWorkbookRootDirId+"'" );
        }
      }
    }

    if( usersFolder == null ) {
      usersFolder = createOrGetChildFolder( lsasWorkbookRootDirId, usersLSANameFromMaster, usersEmailAddress );
      Logger.log( "Created Brand New LSA Folder With Name '"+usersFolder.getName()+"' and id: " + usersFolder.getId() );
    }

    //Make a copy of the master file in the folder containing users spreadsheet
    //with the new version number reflected in the filename
    let usersNewFileFilename = usersLSANameFromMaster + " RoS Workbook v" + mastersVersionNumber;
    Logger.log( "Creating new file: '" + usersNewFileFilename + "'");
    let usersNewFile = masterFile.makeCopy(usersNewFileFilename + "_INCOMPLETE", usersFolder);
    Logger.log( "New File's ID: '" + usersNewFile + "'");

    //now set the file permissions - it's going to be owned by the dev account (which is that account this WebApp project runs as)
    //but we need to copy the editors and viewers from the master so that the master admins also have access to go into
    //everyone's spreadhseet and make changes and help them with their settings etc.
    //we lso need to give the user access of course - but as an editor not an owner
    Logger.log( "Setting File Permissions for: " + usersNewFileFilename );

    var mastersViewers    = masterFile.getViewers();
    var mastersEditors    = masterFile.getEditors();
    var usersNewFileViewers = usersNewFile.getViewers();
    var usersNewFileEditors = usersNewFile.getEditors();

    // **** NEXT WORK OUT WHICH VIEWERS EXIST ON MASTER THAT DONT ALREADY EXIST ON THE NEW FILE
    var newViewEmails = [], newEditEmails = [];
    var missingViewEmails = [], missingEditEmails = []; 
    var tempEmail = null;

    //convert usersNewFileViewers and usersNewFileEditors to arrays of strings containing email addresses
    for( var nV=0; nV<usersNewFileViewers.length; nV++ ) { newViewEmails.push( usersNewFileViewers[nV].getEmail() ); }
    for( var nE=0; nE<usersNewFileEditors.length; nE++ ) { newEditEmails.push( usersNewFileEditors[nE].getEmail() ); }
    Logger.log( "New file already has these Viewers: " + newViewEmails );
    Logger.log( "New file already has these Editors: " + newEditEmails );

    //loop through mastersViewers and mastersEditors and populate arrays of missing email addresses
    for( var mV=0; mV<mastersViewers.length; mV++ ) {
      tempEmail = mastersViewers[mV].getEmail();
      Logger.log( "Found Master Viewer: " + tempEmail );
      if( !newViewEmails.includes( tempEmail ) ) { missingViewEmails.push( tempEmail ); }
    }
    Logger.log( "Missing Viewers: " + missingViewEmails );

    for( var mE=0; mE<mastersEditors.length; mE++ ) {
      tempEmail = mastersEditors[mE].getEmail();
      Logger.log( "Found Master Editor: " + tempEmail );
      if( !newEditEmails.includes( tempEmail ) ) { missingEditEmails.push( tempEmail ); }
    }
    Logger.log( "Missing Editors: " + missingEditEmails );

    //now add the LSA to the list of missing Editors if it's not already in there
    if( !newEditEmails.includes( usersEmailAddress ) && !missingEditEmails.includes( usersEmailAddress ) ) {
      missingEditEmails.push( usersEmailAddress );
      Logger.log( "Adding LSA User as editor, Missing Editors now = " + missingEditEmails );
    }

    //now actually make the updates to the file's editor and view permissions - if any are necessary.
    //on a shared drive these shouldnt be necessary as the containing folder should already be shared with the right poeple
    try {
      for (var v=0; v<missingViewEmails.length; v++) {
        usersNewFile.addViewer(missingViewEmails[v]);
        Logger.log( "UPDATING PERMISSIONS - Added Viewer: " + missingViewEmails[v]);
      }
      for (var ed=0; ed<missingEditEmails.length; ed++) { 
        usersNewFile.addEditor(missingEditEmails[ed]);
        Logger.log( "UPDATING PERMISSIONS - Added Editor: " + missingEditEmails[ed] );
      }
    }
    catch ( e ) {
      returnData.success = 0;
      returnData.errorMessage = "Could not set editor and viewer permissions on the new workbook: " + usersNewFileFilename + 
                                "\n" + e;
      Logger.log( returnData.errorMessage );
      usersNewFile.setTrashed( true );
      return;
    }
    
    //get details from newly copied file being upgraded/repaired/created
    var usersNewSpreadsheet = SpreadsheetApp.open(usersNewFile);
    var usersNewSpreadsheetId = usersNewSpreadsheet.getId();
    Logger.log( "Your New File has been created: " + usersNewFileFilename + "      file id = " + usersNewSpreadsheetId );
    var usersNewSpreadsheet_SettingsSheet = usersNewSpreadsheet.getSheetByName( SHEETS.GLOBAL_SETTINGS.NAME );
    var usersNewSpreadsheet_InputSheet    = usersNewSpreadsheet.getSheetByName( SHEETS.INPUT.NAME           );

    // ... parse upgrade script tab on master spreadsheet to see which ranges from the old sheet need to
    // be copied across and where to copy them to. This is the most time consuming bit.
    if( modeCode != MODES.CREATE.CODE ) {
      Logger.log( "Starting to copy data ranges from old spreadsheet using the upgrade script filtered by source version: " +
                  usersOrigVersionNumber );
      let master_UpgradeScriptSheet = masterSpreadsheet.getSheetByName( SHEETS.MASTER_UPGRADE_SCRIPT.NAME );
      //find all records from the script for the source version in question, in the right order
      let oldVersionRowIds = findAllInColumn( master_UpgradeScriptSheet, usersOrigVersionNumber, 
                                              SHEETS.MASTER_UPGRADE_SCRIPT.REFS.COL_NO_SOURCE_VERSION, 
                                              SHEETS.MASTER_UPGRADE_SCRIPT.REFS.ROW_NO_FIRST_SCRIPT_ROW, 
                                              SHEETS.MASTER_UPGRADE_SCRIPT.REFS.ROW_NO_LAST_SCRIPT_ROW );
      //setup vars to loop through source ranges and paste to dest sheet
      let sourceSheetName = null; let destSheetName = null; let sourceRangeDef = null;
      let findAndReplaceJSON = null;
      let destinationPasteRow = null; let destinationPasteCol = null;
      let upgradeScriptRow = null;
      let tempSourceSheet = null, tempDestinationSheet = null;
      let tempRangeOfData = null; let tempRangeValues2DArray = null; let tempDestinationRange = null;

      //loop through each record of script in order
      let offset = SHEETS.MASTER_UPGRADE_SCRIPT.REFS.COL_NO_SHEET_NAME;
      for( var i = 0; i < oldVersionRowIds.length; i++ ) {
        //get the upgrade script record in question
        upgradeScriptRow    = master_UpgradeScriptSheet.getRange( 
              oldVersionRowIds[i], SHEETS.MASTER_UPGRADE_SCRIPT.REFS.COL_NO_SHEET_NAME, 
              1, SHEETS.MASTER_UPGRADE_SCRIPT.REFS.COL_NO_FIND_REPLACE_JSON 
        ).getValues();
        sourceSheetName     = upgradeScriptRow[0][SHEETS.MASTER_UPGRADE_SCRIPT.REFS.COL_NO_SHEET_NAME - offset];
        sourceRangeDef      = upgradeScriptRow[0][SHEETS.MASTER_UPGRADE_SCRIPT.REFS.COL_NO_SOURCE_RANGE - offset];
        destinationPasteRow = upgradeScriptRow[0][SHEETS.MASTER_UPGRADE_SCRIPT.REFS.COL_NO_DEST_PASTE_AT_ROW - offset];
        destinationPasteCol = upgradeScriptRow[0][SHEETS.MASTER_UPGRADE_SCRIPT.REFS.COL_NO_DEST_PASTE_AT_COL - offset ];
        destSheetName       = upgradeScriptRow[0][SHEETS.MASTER_UPGRADE_SCRIPT.REFS.COL_NO_RENAME_SHEET_TO - offset ];
        findAndReplaceJSON  = upgradeScriptRow[0][SHEETS.MASTER_UPGRADE_SCRIPT.REFS.COL_NO_FIND_REPLACE_JSON - offset ];
        findAndReplaceJSON  = ( findAndReplaceJSON ) ? findAndReplaceJSON : "[]";

        //destination sheet will be blank if we're not renaming the sheet mid upgrade
        destSheetName = ( !destSheetName || destSheetName == "" ) ? sourceSheetName : destSheetName;

        Logger.log( "  Script Row: " + oldVersionRowIds[i] + " - Source Sheet: " + sourceSheetName + ", Source Range: " + 
                    sourceRangeDef + ", Dest Cell: row " + destinationPasteRow + ", col " + destinationPasteCol + ", Dest Sheet: " +
                    destSheetName + ", findAndReplaceJSON:\n" + findAndReplaceJSON );

        let jsonValidationError = validateFindAndReplaceJSON( findAndReplaceJSON );
        if( jsonValidationError !== null ) {
          returnData.success = 0;
          returnData.errorMessage = "Sorry, there was an issue while trying to copy the data from this Workbook to the new" +
                " upgraded version of the Workbook\n"+
                "We tried to run a Search and Replace operation on the data for this copy but the config was malformed:\n" + 
                "Source Sheet: " + sourceSheetName + "\n" +
                "Source Range: " + sourceRangeDef + "\n" +
                "\nAnd the error message generated was:\n" + jsonValidationError + "\n\n" +
                "Please copy this error message and then email it to support using the menu at the top by clicking\n" +
                "[LSA Menu] --> [Help & Support] --> [Email Support]";
          Logger.log( returnData.errorMessage );
          usersNewFile.setTrashed( true );
          return;
        }

        //get the actual data from the source sheet range and copy to dest sheet at the top left cell ref
        //specified in the script
        try {
          tempSourceSheet = usersSpreadsheet.getSheetByName( sourceSheetName );
          tempDestinationSheet = usersNewSpreadsheet.getSheetByName( destSheetName );
          tempRangeOfData = tempSourceSheet.getRange( sourceRangeDef );
          tempDestinationRange = tempDestinationSheet.getRange( destinationPasteRow, destinationPasteCol, 
                                                                tempRangeOfData.getHeight(), tempRangeOfData.getWidth() );

          try {
            tempRangeValues2DArray = tempRangeOfData.getValues();
            findAndReplaceExactCellValuesInRange_FromJSON( tempRangeValues2DArray, findAndReplaceJSON );
            tempDestinationRange.setValues( tempRangeValues2DArray );
            SpreadsheetApp.flush();
          }
          catch( eInner ) {
            Logger.log( "  COPY FAILED SO TEMPORARILY ALLOWING INVALID VALUES ON ALL VALIDATION RULES AND TRYING AGAIN" );
            //if the copy fails it is almost always because of Validation rules on the cells
            //Edit the range cells Allow Invalid Values on all Validation Rules in the Range
            //TO DO: CHeck if any o fthe validation rules already have allow setAllowInvalid = true
            //AND IF THEY DO NONT SET IT BACK TO FALSE AFTER - NOT NECESSARY AT TIME OF WRITING AS I DONT USE
            //ALLOW INVALID CHECKS
            let validationRuleMatrix = tempDestinationRange.getDataValidations();
            let rule = null;
            for (let yi = 0; yi < validationRuleMatrix.length; yi++) {
              for (let xi = 0; xi < validationRuleMatrix[yi].length; xi++) {
                rule = validationRuleMatrix[yi][xi];
                if (rule != null) {
                  validationRuleMatrix[yi][xi] = rule.copy().setAllowInvalid(true).build();
                }
              }
            }
            tempDestinationRange.setDataValidations(validationRuleMatrix);
            SpreadsheetApp.flush();

            //try the copy again but this time dont catch any errors, we need to report these errors
            //as something has gone wrong
            tempRangeValues2DArray = tempRangeOfData.getValues();
            findAndReplaceExactCellValuesInRange_FromJSON( tempRangeValues2DArray, findAndReplaceJSON );
            tempDestinationRange.setValues( tempRangeValues2DArray );
            SpreadsheetApp.flush();

            //Disallow Invalid Values on all Validation Rules in the Range
            for (let yi = 0; yi < validationRuleMatrix.length; yi++) {
              for (let xi = 0; xi < validationRuleMatrix[yi].length; xi++) {
                rule = validationRuleMatrix[yi][xi];
                if (rule != null) {
                  validationRuleMatrix[yi][xi] = rule.copy().setAllowInvalid(false).build();
                }
              }
            }
            tempDestinationRange.setDataValidations(validationRuleMatrix);
          }
        }
        catch( e ) {
          returnData.success = 0;
          returnData.errorMessage = "Sorry, there was an issue while trying to copy the data from this Workbook to the new" +
                " upgraded version of the Workbook\n"+
                "The data being copied was as follows:\n" + 
                "Source Sheet: " + sourceSheetName + "\n" +
                "Source Range: " + sourceRangeDef + "\n" +
                "\nAnd the error message generated was:\n" + e + "\n\n" +
                "Please copy this error message and then email it to support using the menu at the top by clicking\n" +
                "[LSA Menu] --> [Help & Support] --> [Email Support]";
          Logger.log( returnData.errorMessage );
          usersNewFile.setTrashed( true );
          return;
        }

        //log that the copy is completed
        Logger.log( "  Script Row: " + oldVersionRowIds[i] + " - range copied" );
      }
    }

    //set the LSA Name after all the data has been copied across
    usersNewSpreadsheet_InputSheet.getRange ( 
          SHEETS.INPUT.REFS.LSA_NAME.ROW_NO, 
          SHEETS.INPUT.REFS.LSA_NAME.COL_NO 
    ).setValue( usersLSANameFromMaster );
    
    //now delete all master sheet functionality from new spreadsheet as it is a child sheet
    usersNewSpreadsheet.deleteSheet( usersNewSpreadsheet.getSheetByName( SHEETS.MASTER_UPGRADE_SCRIPT.NAME  ) );
    usersNewSpreadsheet.deleteSheet( usersNewSpreadsheet.getSheetByName( SHEETS.MASTER_LEARNERS.NAME        ) );
    usersNewSpreadsheet.deleteSheet( usersNewSpreadsheet.getSheetByName( SHEETS.MASTER_LSAS.NAME            ) );
    usersNewSpreadsheet.deleteSheet( usersNewSpreadsheet.getSheetByName( SHEETS.MASTER_ANNOUNCEMENTS.NAME   ) );
    usersNewSpreadsheet.deleteSheet( usersNewSpreadsheet.getSheetByName( SHEETS.MASTER_TEMPLATE.NAME        ) );
    usersNewSpreadsheet.deleteSheet( usersNewSpreadsheet.getSheetByName( SHEETS.MASTER_ASSETS.NAME        ) );
    Logger.log( "Deleted the 6 Master Only Sheets from the new workbook being " + modeDescription + "ed" );

    //now show/hide the settings sheets in the new workbook so it's ready for use
    usersNewSpreadsheet.getSheetByName( SHEETS.SETTINGS_LEARNERS.NAME       ).showSheet();
    usersNewSpreadsheet.getSheetByName( SHEETS.SETTINGS_LESSONS.NAME        ).showSheet();
    usersNewSpreadsheet.getSheetByName( SHEETS.SETTINGS_TARGET_GRADES.NAME  ).showSheet();
    usersNewSpreadsheet.getSheetByName( SHEETS.SETTINGS_LESSON_TARGETS.NAME ).showSheet();
    if( modeCode != MODES.CREATE.CODE ) {
      //if it's an upgrade or repair then hide the sheet if it still exists under the same name in the old version
      //and if it was hidden in the old workbook
      let oldLearnersSheet = usersSpreadsheet.getSheetByName( SHEETS.SETTINGS_LEARNERS.NAME );
      if( oldLearnersSheet && oldLearnersSheet.isSheetHidden() ) { 
        usersNewSpreadsheet.getSheetByName( oldLearnersSheet ).hideSheet();
      }
      let oldLessonsSheet = usersSpreadsheet.getSheetByName( SHEETS.SETTINGS_LESSONS.NAME );
      if( oldLessonsSheet && oldLessonsSheet.isSheetHidden() ) { 
        usersNewSpreadsheet.getSheetByName( oldLessonsSheet ).hideSheet();
      }
      let oldTargetGradesSheet = usersSpreadsheet.getSheetByName( SHEETS.SETTINGS_TARGET_GRADES.NAME );
      if( oldTargetGradesSheet && oldTargetGradesSheet.isSheetHidden() ) { 
        usersNewSpreadsheet.getSheetByName( oldTargetGradesSheet ).hideSheet();
      }
      let oldLessonTargetsSheet = usersSpreadsheet.getSheetByName( SHEETS.SETTINGS_LESSON_TARGETS.NAME );
      if( oldLessonTargetsSheet && oldLessonTargetsSheet.isSheetHidden() ) { 
        usersNewSpreadsheet.getSheetByName( oldLessonTargetsSheet ).hideSheet();
      }
    }

    //unhide the sheets that The LSAs definitely need to see
    usersNewSpreadsheet.getSheetByName( SHEETS.INPUT.NAME           ).showSheet();
    usersNewSpreadsheet.getSheetByName( SHEETS.MY_FILES.NAME        ).showSheet();
    usersNewSpreadsheet.getSheetByName( SHEETS.MOBILE_MAIN.NAME     ).showSheet();
    //monile - input should be hidden by default as mobile version will have been un-activated during upgrader/repair/create
    if( !usersNewSpreadsheet.getSheetByName( SHEETS.MOBILE_THIS_RECORD.NAME ).isSheetHidden() ) {
      usersNewSpreadsheet.getSheetByName( SHEETS.MOBILE_THIS_RECORD.NAME ).showSheet();
    }

    //hide the sheets that are locked so can't and shouldn't then be unhidden by the LSAs
    usersNewSpreadsheet.getSheetByName( SHEETS.GLOBAL_SETTINGS.NAME ).hideSheet();
    usersNewSpreadsheet.getSheetByName( SHEETS.MASTER_HELP.NAME     ).hideSheet();

    //loop through record of support tabs 1 to 25 hiding 6 onwards and unhiding 1 thru 5
    for ( i = 0; i < 25; i++) {
      if( i < 5 && usersNewSpreadsheet.getSheetByName(""+(i+1)).isSheetHidden() ) {
        usersNewSpreadsheet.getSheetByName(""+(i+1)).showSheet();
      } 
      else if( i >= 5 && !usersNewSpreadsheet.getSheetByName(""+(i+1)).isSheetHidden() ) {
        usersNewSpreadsheet.getSheetByName(""+(i+1)).hideSheet();
      } 
    }
    Logger.log( "Show / Hide sheets accordingly, for the new workbook being " + modeDescription + "ed" );

    //now edit the settings after the copy - including linking this upgraded/repaired/created file to the latest master
    //this upgraded file's main user's email address
    usersNewSpreadsheet_SettingsSheet.getRange(
      SHEETS.GLOBAL_SETTINGS.REFS.ROW_NO_THIS_FILES_ID,
      SHEETS.GLOBAL_SETTINGS.REFS.COL_NO
    ).setValue( usersNewSpreadsheetId );
    usersNewSpreadsheet_SettingsSheet.getRange(
      SHEETS.GLOBAL_SETTINGS.REFS.ROW_NO_MAIN_USERS_EMAIL,
      SHEETS.GLOBAL_SETTINGS.REFS.COL_NO
    ).setValue( usersEmailAddress );
    //this upgraded file's fileId
    usersNewSpreadsheet_SettingsSheet.getRange(
      SHEETS.GLOBAL_SETTINGS.REFS.ROW_NO_MASTERS_LINK_TO_THIS_CHILD,
      SHEETS.GLOBAL_SETTINGS.REFS.COL_NO
    ).setValue( usersNewSpreadsheetId );
    //the latest master's file id
    usersNewSpreadsheet_SettingsSheet.getRange(
      SHEETS.GLOBAL_SETTINGS.REFS.ROW_NO_MASTER_FILE_ID,
      SHEETS.GLOBAL_SETTINGS.REFS.COL_NO
    ).setValue( masterFile.getId() );
    //the latest master's version
    usersNewSpreadsheet_SettingsSheet.getRange(
      SHEETS.GLOBAL_SETTINGS.REFS.ROW_NO_MASTERS_VERSION,
      SHEETS.GLOBAL_SETTINGS.REFS.COL_NO
    ).setValue( mastersVersionNumber );

    //now set the appropriate annoucement if upgrade or repair
    //for create the correct announcement is already copied in from the master spreadsheet
    if( modeCode == MODES.UPDATE.CODE || 
        ( modeCode == MODES.REPAIR.CODE && usersOrigVersionNumber != mastersVersionNumber ) ) {

      //add just upgraded release notes announcement to the queue
      var justUpgradedPrefix = usersNewSpreadsheet_SettingsSheet.getRange( 
          SHEETS.GLOBAL_SETTINGS.REFS.ROW_NO_JUST_UPDATED_TXT, 
          SHEETS.GLOBAL_SETTINGS.REFS.COL_NO
        ).getValue();
      var releaseNotes = usersNewSpreadsheet_SettingsSheet.getRange( 
          SHEETS.GLOBAL_SETTINGS.REFS.ROW_NO_RELEASE_NOTES, 
          SHEETS.GLOBAL_SETTINGS.REFS.COL_NO
        ).getValue();
      var announcementText = justUpgradedPrefix + "\n" + releaseNotes;
      Logger.log( "Adding announcement to the queue:\n" + announcementText );
      AddAnnouncementToTheQueue(  announcementText, usersNewSpreadsheet_SettingsSheet, 
                                  SHEETS.GLOBAL_SETTINGS.REFS.ROW_NO_PENDING_ANNOUNCEMENT );
    }
    else if( modeCode == MODES.REPAIR.CODE && usersOrigVersionNumber == mastersVersionNumber ) {
      //add 'repair successful' release notes to the queue
      var announcementText =  "Your workbook has been repaired. If your issue persists please use the top [LSA Menu] " +
                              "to email support.";
      Logger.log( "Adding announcement to the queue:\n" + announcementText );
      AddAnnouncementToTheQueue( announcementText, usersNewSpreadsheet_SettingsSheet, 
                                  SHEETS.GLOBAL_SETTINGS.REFS.ROW_NO_PENDING_ANNOUNCEMENT );
    }

    Logger.log( "Updated the Global Settings on the file being " + modeDescription + "ed to link to correct version numbers" +
                " and user email address. " +
                " Also updated the file id record of the child AND master files so link TO the master is maintained" );

    //now link the latest master back to the upgraded file to complete the link
    master_LSAsSheet.getRange( 
      foundLSARowNum, 
      SHEETS.MASTER_LSAS.REFS.COL_NO_USERS_FILE_ID
    ).setValue( usersNewSpreadsheetId );
    master_LSAsSheet.getRange( 
      foundLSARowNum, 
      SHEETS.MASTER_LSAS.REFS.COL_NO_USERS_VERSION
    ).setValue( mastersVersionNumber );
    Logger.log( "Updated the Master's LSA sheet to link back to this new upgraded file so reverse link is maintained" );   
    
    //update the newly updated files filename so that the user knows the file is complete
    usersNewFile.setName( usersNewFileFilename );

    //Update the old decommissioned file so it knows it's old and decomissioned
    if( modeCode != MODES.CREATE.CODE ) {
      GetMasterSettingsCellFromOldVersionWorkbook( usersSpreadsheet, users_SettingsSheet, 
          SHEETS.GLOBAL_SETTINGS.REFS.HANDLE_MASTERS_LINK_TO_THIS_CHILD, SHEETS.GLOBAL_SETTINGS.REFS.ROW_NO_MASTERS_LINK_TO_THIS_CHILD 
      ).setValue( usersNewSpreadsheetId );

      usersSpreadsheet.rename( usersSpreadsheet.getName() + "_DECOMISSIONED" );
      Logger.log( "Updated and renamed old decomissioned file so it knows it's decomissioned" ); 
    }


    //set owner of newly updated file, if necessary - NOT NECESSARY IF RUNNING UNDER WEBAPP UNDER DEV ACCOUNT
    SpreadsheetApp.flush(); //clear every thing down as best as possible before we potentially revoke our own permissions
    if( runningAsUser && usersNewFile.getOwner().getEmail() != devEmailAddress.trim() )
    {
      usersNewFile.setOwner( devEmailAddress.trim() );
      Logger.log( "Set Upgraded file's owner to: " + devEmailAddress ); 
    }      

    //FINISHED SUCCESSFULLY
    returnData.success = 1;
    returnData.errorMessage = null;
    returnData.affectedFileId = usersNewSpreadsheetId;
    Logger.log( ( modeDescription.toUpperCase() ) + " COMPLETED SUCCESSFULLY" );
    return;
  }


}
