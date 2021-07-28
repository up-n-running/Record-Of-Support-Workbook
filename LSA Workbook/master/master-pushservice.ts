function PushToChildWorksheets( selectedLSANos: Array<number>|null, 
        newMasterFileId: string|null, 
        newAnnouncementText: string|null, 
        syncLearners: boolean, 
        syncHelpCentreLinks: boolean ,
        syncNewDeploymentIds: boolean,
        syncAdminGlobalSettings: boolean,
        overrideTimeoutStartTime?: number|null ) {
  Logger.log( "PushToChildWorksheets called" );

  //stringify push service params to create a config string
  let pushServiceConfig = 
        ( ( selectedLSANos ) ? selectedLSANos.length : -1 ) + "|" +
        ( ( newMasterFileId ) ? newMasterFileId : "" ) + "|" +
        ( ( newAnnouncementText ) ? newAnnouncementText.substring( 1, 5 ) : "" ) + "|" +
        ( ( syncLearners ) ? "true" : "false" ) + "|" +
        ( ( syncHelpCentreLinks ) ? "true" : "false" ) + "|" +
        ( ( syncNewDeploymentIds ) ? "true" : "false" ) + "|" + 
        ( ( syncAdminGlobalSettings ) ? "true" : "false" ) + "|";

  //take start time snapshot
  let pushServiceStartTime = ( overrideTimeoutStartTime && overrideTimeoutStartTime > 0 ) ? 
        overrideTimeoutStartTime : new Date().getTime();

  //get sheet and ui
  let spreadsheet = SpreadsheetApp.getActive();
  let ui = SpreadsheetApp.getUi();
  let mastersGlobalSettingsSheet: GoogleAppsScript.Spreadsheet.Sheet = spreadsheet.getSheetByName( SHEETS.GLOBAL_SETTINGS.NAME );

  //read the push service settings from the global settings
  let serviceTimeoutSecs  = mastersGlobalSettingsSheet.getRange( SHEETS.GLOBAL_SETTINGS.REFS.ROW_NO_PUSH_SVC_TIMEOUT,
        SHEETS.GLOBAL_SETTINGS.REFS.COL_NO ).getValue();
  let lastPushConfig      = mastersGlobalSettingsSheet.getRange( SHEETS.GLOBAL_SETTINGS.REFS.ROW_NO_LAST_PUSH_CONFIG,
        SHEETS.GLOBAL_SETTINGS.REFS.COL_NO ).getValue();
  let lastPushTimedoutTime = mastersGlobalSettingsSheet.getRange( SHEETS.GLOBAL_SETTINGS.REFS.ROW_NO_LAST_PUSH_TIMEDOUT,
        SHEETS.GLOBAL_SETTINGS.REFS.COL_NO ).getValue();

  //switch the view to the LSAs worksheet
  let lsaWorksheet: GoogleAppsScript.Spreadsheet.Sheet = 
        spreadsheet.setActiveSheet( spreadsheet.getSheetByName( SHEETS.MASTER_LSAS.NAME ) );

  //read all lsa numbers for later
  let lsaNoColumnValues = lsaWorksheet.getRange( 
    SHEETS.MASTER_LSAS.REFS.ROW_NO_FIRST_LSA, 
    SHEETS.MASTER_LSAS.REFS.COL_NO_LSA_NOS,
    SHEETS.MASTER_LSAS.REFS.ROW_NO_LAST_LSA - SHEETS.MASTER_LSAS.REFS.ROW_NO_FIRST_LSA + 1,
    1 ).getValues();

  //read all File Ids for later
  let fileIdColumnValues = lsaWorksheet.getRange( 
    SHEETS.MASTER_LSAS.REFS.ROW_NO_FIRST_LSA, 
    SHEETS.MASTER_LSAS.REFS.COL_NO_USERS_FILE_ID,
    SHEETS.MASTER_LSAS.REFS.ROW_NO_LAST_LSA - SHEETS.MASTER_LSAS.REFS.ROW_NO_FIRST_LSA + 1,
    1 ).getValues();    

  //check if the last push service timed out and if so offer to carry on from where we left off
  let lsaNosToExclude: Array<number> = new Array();
  if( lastPushTimedoutTime > 0 ) {
    //read all push statuses
    let statusColumnValues = lsaWorksheet.getRange( 
          SHEETS.MASTER_LSAS.REFS.ROW_NO_FIRST_LSA, 
          SHEETS.MASTER_LSAS.REFS.COL_NO_PUSH_SVC_STATUS,
          SHEETS.MASTER_LSAS.REFS.ROW_NO_LAST_LSA - SHEETS.MASTER_LSAS.REFS.ROW_NO_FIRST_LSA + 1,
          1 ).getValues();
    let awaitingLSANumbers: Array<number> = new Array();
    let noOfWaitingFound:   number = 0;
    let tmpStatus: string|null = null;
    let tmpLSANo:  number|null = null;
    for( let i = 0; i < statusColumnValues.length; i++ ) {
      tmpStatus = statusColumnValues[ i ][ 0 ];
      tmpLSANo  = lsaNoColumnValues[ i ][ 0 ];
      if( tmpStatus == SHEETS.MASTER_LSAS.STATUSES.WAITING ) { 
        noOfWaitingFound++;
        awaitingLSANumbers.push( tmpLSANo );
      }
    }
    //see if it finished off in a timeout state

    if( pushServiceConfig == lastPushConfig ) { 
      let btn = ui.alert( "Continue the last timed-out Push?", 
      "The last time the push service was run it timed out with " + awaitingLSANumbers.length + " files left to go, and you " + 
      "have just repeated the action to prompt it to continue. This is good practice, have a gold star! \n\n"+
      "If you actually wanted to start a brand new push then you can click 'No', but this is NOT RECOMMENDED.\n\n" +
      "Press 'Yes' to continue the previous, timed-out push (recommended)",
      ui.ButtonSet.YES_NO );
      if( btn == ui.Button.YES ) {
        selectedLSANos = awaitingLSANumbers;
      }
    }
    else {
      let btn = ui.alert( "Warning!, last Push timed out.", 
      "Last time the push service ran, it timed-out and it is currently waiting for the user to repeat that action " +
      " so that it can carry on from where it last left off and complete successfully.\n" +
      "It appears that you are trying to perform a different push operation to the previous timed out one. " +
      "If you continue your push will be performed, however the saved state of the last timed-out push service wil be lost.\n\n" +
      "Ideally you should cancel and repeat the action (ie by clicking the button or menu item) used in the last push service to allow " +
      "it to finish off and complete successfully. If you are sure you want to perform YOUR push anyway then press 'OK' (not recommended).\n\n" +
      "Press Cancel to click the button or menu item used previously to allow the previos push to complete.", 
      ui.ButtonSet.OK_CANCEL );
      if( btn == ui.Button.CANCEL ) {
        return false;
      }
    }

  }

  //if selectedLSANos is null then we should populate the array with ALL LSA numbers!
  let noOfLsaRecords: number = SHEETS.MASTER_LSAS.REFS.ROW_NO_LAST_LSA - SHEETS.MASTER_LSAS.REFS.ROW_NO_FIRST_LSA + 1;
  if( selectedLSANos == null ) {
    selectedLSANos = new Array();
    let tmpFileId: string|null = null;
    let tmpLSANo:  number|null = null;
    for( let i = 0; i < lsaNoColumnValues.length; i++ ) {
      tmpFileId = fileIdColumnValues[i][0];
      tmpLSANo  = lsaNoColumnValues[i][0];
      if( tmpFileId != "" ) {
        selectedLSANos.push( tmpLSANo );
      }
    }
  }
  Logger.log( "selectedLSANos = " + selectedLSANos );

  //get the Master's version number and file id from this sheet as we may need it later
  let thisMastersActualFileId: string = spreadsheet.getId();
  let thisMastersVersionNo: string = mastersGlobalSettingsSheet.getRange( SHEETS.GLOBAL_SETTINGS.REFS.ROW_NO_VERSION_NO,
                                                              SHEETS.GLOBAL_SETTINGS.REFS.COL_NO ).getValue();

  //clear the Push Service Last Status column down so it can be re-populated as we loop through the LSAs below
  let lastStatusesRange = lsaWorksheet.getRange( 
        SHEETS.MASTER_LSAS.REFS.ROW_NO_FIRST_LSA, 
        SHEETS.MASTER_LSAS.REFS.COL_NO_PUSH_SVC_STATUS,
        noOfLsaRecords, 1 );
  lastStatusesRange.setValue( "" );

  //create a rangelist for the Push Service Last Status column with all the LSA Numbers we want to update in it
  //then we'll set the status to "-" on all the cells at the same time. doing it cell by cell talkes about 0.5 seconds per cell
  //which can take nearly a minute when were pushing for all LSA records!
  let pushStatusColLetter: string = columnToLetter( SHEETS.MASTER_LSAS.REFS.COL_NO_PUSH_SVC_STATUS );
  let arrayOfLSAPushStatusCellsInA1Notation: Array<string> = new Array();
  for( var i=0; i<selectedLSANos.length; i++ ) {
    arrayOfLSAPushStatusCellsInA1Notation.push( 
        pushStatusColLetter + (selectedLSANos[i] + SHEETS.MASTER_LSAS.REFS.ROW_NO_FIRST_LSA - 1 ) );
  }
  lsaWorksheet.getRangeList( arrayOfLSAPushStatusCellsInA1Notation ).setValue( SHEETS.MASTER_LSAS.STATUSES.WAITING );

  //FLAG THE SERVICE AS STARTING BY SAVING THE CONFIG ANF CHEARING THE LAST TIMEOUT DATE IN GLOBAL SETTINGS
  mastersGlobalSettingsSheet.getRange( SHEETS.GLOBAL_SETTINGS.REFS.ROW_NO_LAST_PUSH_CONFIG,
    SHEETS.GLOBAL_SETTINGS.REFS.COL_NO ).setValue( pushServiceConfig );
  mastersGlobalSettingsSheet.getRange( SHEETS.GLOBAL_SETTINGS.REFS.ROW_NO_LAST_PUSH_TIMEDOUT,
    SHEETS.GLOBAL_SETTINGS.REFS.COL_NO ).setValue( -1 );

  //if syncLearners is set then we must collect all the Learner data into a neat array before looping so it can be passed
  //down to each of the child LSA Workbooks
  let allMasterLearnerObjects: Array<Object> = syncLearners ? MasterLearnerSearch( null, null, null, null, true, true ) : null;

  //prepare to loop through each of the LSA numbers in the array and check if their LSA Workbook has already been generated
  //these are used to manage the iteration through the loop
  let noOfWorksheetsFound: number= 0; let noOfWorksheetsUpdated: number = 0; let noOfWorksheetsFailed: number = 0; 
  let recordNo: number = -1; let rowNo: number = -1;
  let childFileId: string = null; let childWorkbook: GoogleAppsScript.Spreadsheet.Spreadsheet = null;
  let childGlobalSettingsSheet = null; let skipRemaining: boolean = false; let timedout: boolean = false;

  //these are used to sanity check the shild spreadsheet is linked correctly
  var childsMasterFileIdLink = null, childsFileIdFromItsGlobalSettings = null, thisIterationErrorMessage = null;
  //these are used to store the actual child cells we'll be reading/writing from/to
  var childsMasterVersionCell = null, childsMasterFileIdCell = null;

  //actually loop through each of the LSA numbers in the array and check if their LSA Workbook has already been generated
  for (var i=0; !skipRemaining && i<selectedLSANos.length; i++) {
    recordNo = selectedLSANos[i];
    rowNo = recordNo + SHEETS.MASTER_LSAS.REFS.ROW_NO_FIRST_LSA - 1;
    childFileId = lsaWorksheet.getRange( rowNo, SHEETS.MASTER_LSAS.REFS.COL_NO_USERS_FILE_ID ).getValue();
    if( childFileId != "" )
    {
      noOfWorksheetsFound++;
      thisIterationErrorMessage = "";
      var childUpdateWasRequired = false;
      try {
        childWorkbook = SpreadsheetApp.openById( childFileId );
        Logger.log( "New Workook Found with ID '"+childFileId+"': " + childWorkbook );        
        childGlobalSettingsSheet = childWorkbook.getSheetByName( SHEETS.GLOBAL_SETTINGS.NAME );

        //*** sanity check that the child workbook is not decomissioned and is linked back to the master correctly ***
        //find the child's link back to the master and the settings cells we'll need later
        childsMasterFileIdCell  = GetMasterSettingsCellFromOldVersionWorkbook( childWorkbook, childGlobalSettingsSheet, 
              SHEETS.GLOBAL_SETTINGS.REFS.HANDLE_MASTER_FILE_ID, SHEETS.GLOBAL_SETTINGS.REFS.ROW_NO_MASTER_FILE_ID );
        childsMasterVersionCell = GetMasterSettingsCellFromOldVersionWorkbook( childWorkbook, childGlobalSettingsSheet, 
              SHEETS.GLOBAL_SETTINGS.REFS.HANDLE_MASTERS_VERSION, SHEETS.GLOBAL_SETTINGS.REFS.ROW_NO_MASTERS_VERSION );
        childsMasterFileIdLink = childsMasterFileIdCell.getValue();

        //find the spreadsheet's setting that stores it's own file id so we can check it's correct - 
        //ALWAYS ON ROW 2 regardless of what version the chils spreadsheet is on
        childsFileIdFromItsGlobalSettings = childGlobalSettingsSheet.getRange( SHEETS.GLOBAL_SETTINGS.REFS.ROW_NO_THIS_FILES_ID,
                                                                    SHEETS.GLOBAL_SETTINGS.REFS.COL_NO ).getValue();

        //check the file is linked properly
        if( childsFileIdFromItsGlobalSettings != childFileId ) {
          thisIterationErrorMessage += "The Child File For LSA number " + recordNo + " appears to have been decommissioned\n" +
              "The LSAs Workbook thinks its File Id is: " + childsFileIdFromItsGlobalSettings +"\n" +
              "The actual File ID of the LSAs Workbook is: " + childFileId +"\n\n" +
              "The file has probably been upgraded and repaired and a new current file been created\n" +
              "This new file should be the one linked to on the Master Workbook's LSAs Sheet\n\n" +
              "Please contact support and ask them to correct this link manually\n\n";
        }
        if( newMasterFileId == null && childsMasterFileIdLink != thisMastersActualFileId ) {
          thisIterationErrorMessage += "The Child File For LSA number " + recordNo + " appears to be linkend to the wrong master\n" +
              "Child's link back to Master File ID: " + childsFileIdFromItsGlobalSettings + "\n" +
              "This Master Spreadsheets File ID: " + thisMastersActualFileId + "\n" +
              "Perhaps you are not using the current active Master Workbook?\n" +
              "Please contact support and ask them to look into this as there could be something wrong\n\n";
        }
        Logger.log( "Finished Master/Child Link Validation, thisIterationErrorMessage = '" + thisIterationErrorMessage +"'" );  
        if( thisIterationErrorMessage == ""  ) {

          //we passed vaildation sanity checks so lets make the updates
          //are we updating the link to a new master spreadsheet?
          if( newMasterFileId != null ) {
            //this means we need to relink the child back to this new master
            if( childsMasterFileIdCell.getValue() != newMasterFileId ) {
              childUpdateWasRequired = true;
              childsMasterFileIdCell.setValue( newMasterFileId );
              Logger.log( "childsMasterFileIdCell ("+childsMasterFileIdCell.getA1Notation()+") set to: " + newMasterFileId );
            }
            if(  childsMasterVersionCell != thisMastersVersionNo ) {
              childUpdateWasRequired = true;
              childsMasterVersionCell.setValue( thisMastersVersionNo );
              Logger.log( "childsMasterVersionCell ("+childsMasterFileIdCell.getA1Notation()+") set to: " + thisMastersVersionNo );
            }
          }

          //are we sending a announcement to the child spreadsheet (ie was the newAnnouncementText parameter passed into this function)?
          if( newAnnouncementText != null ) {
            var rowNoPendingAncnmt = GetMasterSettingsRowFromOldVersionWorkbook( childWorkbook, childGlobalSettingsSheet, 
                    SHEETS.GLOBAL_SETTINGS.REFS.HANDLE_PENDING_ANNOUNCEMENT, SHEETS.GLOBAL_SETTINGS.REFS.ROW_NO_PENDING_ANNOUNCEMENT );
            Logger.log( "Calling  AddAnnouncementToTheQueue( text, sheet, "+rowNoPendingAncnmt+");" );
            AddAnnouncementToTheQueue( newAnnouncementText, childGlobalSettingsSheet, rowNoPendingAncnmt );
          }

          //are we syncing updated learner records down to the child spreadsheet?
          if( syncLearners != null && syncLearners == true ) {
            let errorText = AddOrEditLearnersOnChildLearnerSheet( childWorkbook, true, null, allMasterLearnerObjects, false, true, null );
            if( errorText ) {
              throw ( "Error syncing Master Learner Records: " + errorText );
            }
          }

          //are we syncing help centre link data down to the child spreadsheet?
          if( syncHelpCentreLinks != null && syncHelpCentreLinks == true ) {
            let errorText = pushAllHelpCentreSheetDate_NoVersionChecking(
                  spreadsheet.getSheetByName( SHEETS.MASTER_HELP.NAME ),
                  childWorkbook.getSheetByName( SHEETS.MASTER_HELP.NAME )
            );
            if( errorText ) {
              throw ( "Error syncing Help Centre Links: " + errorText );
            }
          }

          //are we syncing a new master webapp deploymentId down to child spreadsheet?
          if( syncNewDeploymentIds != null && syncNewDeploymentIds == true ) {
            //TO DO: THIS CAN BE REPLACED BY A CALL TO pushMasterSettingsCellValueToOldVersionWorkbook_ NOW
            let childsMasterWebappAllDeployIdCell  = GetMasterSettingsCellFromOldVersionWorkbook( childWorkbook, 
                  childGlobalSettingsSheet, SHEETS.GLOBAL_SETTINGS.REFS.HANDLE_WEBAPP_ALL_DEPLOYID, 
                  SHEETS.GLOBAL_SETTINGS.REFS.ROW_NO_WEBAPP_ALL_DEPLOYID );
            if( childsMasterWebappAllDeployIdCell == null ) {
              throw ( "Error syncing Deployment Id, coiunt not fild child settings cell" );
            }
            childsMasterWebappAllDeployIdCell.setValue(
              mastersGlobalSettingsSheet.getRange( 
                    SHEETS.GLOBAL_SETTINGS.REFS.ROW_NO_WEBAPP_ALL_DEPLOYID,
                    SHEETS.GLOBAL_SETTINGS.REFS.COL_NO 
              ).getValue()
            );
          }

          if( syncAdminGlobalSettings != null && syncAdminGlobalSettings == true ) {

            //quicklinks
            let firstQLRow:number = SHEETS.GLOBAL_SETTINGS.REFS.ROW_NO_QUICKLINKS_FIRST;
            let lastQLRow: number = SHEETS.GLOBAL_SETTINGS.REFS.ROW_NO_QUICKLINKS_LAST;
            let qlHandlePrefix: string = SHEETS.GLOBAL_SETTINGS.REFS.HANDLEPREFIX_QUICKLINKS + "_";
            let qlNo: number = -1;
            let qlHandle: string = "";
            for( let qlRow: number = firstQLRow; qlRow <= lastQLRow; qlRow ++ ) {
              qlNo = qlRow - firstQLRow + 1;
              qlHandle = "[" + qlHandlePrefix + qlNo + "]";
              Logger.log( qlHandle );
              pushMasterSettingsCellValueToOldVersionWorkbook_( mastersGlobalSettingsSheet, childWorkbook, childGlobalSettingsSheet, 
                qlHandle, qlRow, false, false );
            }

            //other non-quicklick settings
            pushMasterSettingsCellValueToOldVersionWorkbook_( mastersGlobalSettingsSheet, childWorkbook, childGlobalSettingsSheet, 
                  SHEETS.GLOBAL_SETTINGS.REFS.HANDLE_MAX_PAST_DAYS_ROS, SHEETS.GLOBAL_SETTINGS.REFS.ROW_NO_MAX_PAST_DAYS_ROS, false, false );
            pushMasterSettingsCellValueToOldVersionWorkbook_( mastersGlobalSettingsSheet, childWorkbook, childGlobalSettingsSheet, 
                  SHEETS.GLOBAL_SETTINGS.REFS.HANDLE_MAX_ROS_DEL_MINS, SHEETS.GLOBAL_SETTINGS.REFS.ROW_NO_MAX_ROS_DEL_MINS, false, false );
            pushMasterSettingsCellValueToOldVersionWorkbook_( mastersGlobalSettingsSheet, childWorkbook, childGlobalSettingsSheet, 
                  SHEETS.GLOBAL_SETTINGS.REFS.HANDLE_SUPPORT_EMAIL_TO, SHEETS.GLOBAL_SETTINGS.REFS.ROW_NO_SUPPORT_EMAIL_TO, false, false );
            pushMasterSettingsCellValueToOldVersionWorkbook_( mastersGlobalSettingsSheet, childWorkbook, childGlobalSettingsSheet, 
                  SHEETS.GLOBAL_SETTINGS.REFS.HANDLE_SUPPORT_EMAIL_CCS, SHEETS.GLOBAL_SETTINGS.REFS.ROW_NO_SUPPORT_EMAIL_CCS, false, false );

          }
        }
      }
      catch ( e ) {
          thisIterationErrorMessage += "There was an unexpected error when " + ( childUpdateWasRequired ? "updating" : "checking" ) +
            " the spreadsheet for LSA Number " + recordNo + ".\n\n" +
            "Please contact support to advise them of this, as this should not happen!\n\n" +
            "The error message returned by the system is:\n" + catchErrorToString( e ) + "\n\n";
      }

      //if there was an error then build an error alert with appropriate advice in it
      thisIterationErrorMessage = thisIterationErrorMessage == "" ? "" : 
          ( thisIterationErrorMessage + "\nWould you like to continue on to push to the remaining LSA Workbooks?" );
      
      if( thisIterationErrorMessage != ""  ) {
        noOfWorksheetsFailed ++;
        let btnPressed = ui.alert( "Issue with an update to an LSA Workbook", thisIterationErrorMessage, ui.ButtonSet.YES_NO );
        skipRemaining = ( btnPressed == ui.Button.NO );
        lsaWorksheet.getRange( rowNo, SHEETS.MASTER_LSAS.REFS.COL_NO_PUSH_SVC_STATUS ).setValue( SHEETS.MASTER_LSAS.STATUSES.FAILURE );
      }
      else {
        //if there WASNT an error then update the push status so the user knows where we are at
        noOfWorksheetsUpdated++;
        lsaWorksheet.getRange( rowNo, SHEETS.MASTER_LSAS.REFS.COL_NO_PUSH_SVC_STATUS ).setValue( SHEETS.MASTER_LSAS.STATUSES.SUCCESS );
      }

      if( !skipRemaining && ( ( new Date().getTime() - pushServiceStartTime ) > serviceTimeoutSecs*1000 ) && i<(selectedLSANos.length-1) ) {
        skipRemaining = true;
        timedout = true;
      }

      Logger.log( "skip = " + skipRemaining );
    }
    else {
      //this record was skipped because there is no spreadsheet to update
      //not updating these cells because each cell update seems to take 0.5 seocnds, x100 = lot of time doing nothing
      //lsaWorksheet.getRange( rowNo, SHEETS.MASTER_LSAS.REFS.COL_NO_PUSH_SVC_STATUS ).setValue( "Skipped" );
    }
  }

  //if it timed out then record this fact and alert the user
  if( timedout ) {
    mastersGlobalSettingsSheet.getRange( SHEETS.GLOBAL_SETTINGS.REFS.ROW_NO_LAST_PUSH_TIMEDOUT,
      SHEETS.GLOBAL_SETTINGS.REFS.COL_NO ).setValue( new Date().getTime() );
    ui.alert( "PUSH SERVICE TIMED OUT", 
          "" + noOfWorksheetsFound + " worksheets attempted and " + noOfWorksheetsUpdated + " worksheets updated successfully.\n\n" +
          "IMPORTANT: There were " + ( selectedLSANos.length - noOfWorksheetsUpdated - noOfWorksheetsFailed ) + " worksheets " +
          "that the push service did not action.\n" +
          "*** YOU NEED TO REPEAT THIS PUSH REQUEST straight away in order to reattempt these untried workbooks ***\n\n" +
          "Whichever button or menu item you clicked to trigger this push service, please press it again immediately!", 
          ui.ButtonSet.OK );
  }
  else {
    ui.alert( "Push Service Completed Successfully", "" + noOfWorksheetsFound + " worksheets found and " + noOfWorksheetsUpdated + 
          " worksheets updated successfully.", ui.ButtonSet.OK );
  }
  
  
  Logger.log( "PushToChildWorksheets finished" );
}

function pushMasterSettingsCellValueToOldVersionWorkbook_( mastersGlobalSettingsSheet, childWorkbook, childGlobalSettingsSheet, 
        settingHandle, currentVersionMastersSettingsRow, useRowNumIfHandleNotFound, throwErrorIfNotFound ) {
  let childsGlobalSettingsCell  = GetMasterSettingsCellFromOldVersionWorkbook( childWorkbook, 
        childGlobalSettingsSheet, settingHandle, ( ( useRowNumIfHandleNotFound ) ? currentVersionMastersSettingsRow : null ) );
  if( childsGlobalSettingsCell == null ) {
    if( throwErrorIfNotFound ) {
      throw new Error( "Error syncing " + settingHandle + ", could not fild child settings cell" );
    }
  }
  else {
    let valueToPush = mastersGlobalSettingsSheet.getRange(  currentVersionMastersSettingsRow, SHEETS.GLOBAL_SETTINGS.REFS.COL_NO ).getValue()
Logger.log( "About to update child Global Settings Cell '" + childsGlobalSettingsCell.getA1Notation() + "' to value: " + valueToPush );
    childsGlobalSettingsCell.setValue( valueToPush );
  }
}