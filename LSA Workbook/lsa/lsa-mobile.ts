function Mobile_ResetInputData( mobileMainSheet?: GoogleAppsScript.Spreadsheet.Sheet ) {

  let spreadsheet = SpreadsheetApp.getActive();
  mobileMainSheet = ( mobileMainSheet ) ? mobileMainSheet : spreadsheet.getSheetByName( SHEETS.MOBILE_MAIN.NAME );

  let ui = getSpoofUIObject( mobileMainSheet );

  ResetInputData( spreadsheet, ui );
}

function Mobile_SendSelectedRecords_DontEmail( mobileMainSheet?: GoogleAppsScript.Spreadsheet.Sheet ) {

  let spreadsheet = SpreadsheetApp.getActive();
  mobileMainSheet = ( mobileMainSheet ) ? mobileMainSheet : spreadsheet.getSheetByName( SHEETS.MOBILE_MAIN.NAME );
  let inputSheet = spreadsheet.getSheetByName( SHEETS.INPUT.NAME );

  let ui = getSpoofUIObject( mobileMainSheet );

  let selectedRecordNos: Array<number> = mobile_getSelectedRecordNos_( mobileMainSheet, ui );

  Logger.log( "selectedRecordNos = " + selectedRecordNos ); 

  if( selectedRecordNos.length > 0 )
  {
    ExportToSendRoSsFromList(selectedRecordNos, spreadsheet, inputSheet, null, ui );
  }
}

function getSpoofUIObject( mobileMainSheet?: GoogleAppsScript.Spreadsheet.Sheet ) {
  return getSpoofUIObject_FromAlertSheetDef( mobileMainSheet, SHEETS.MOBILE_MAIN.REFS.ALERT_BOX );
}

function mobile_getSelectedRecordNos_( mobileMainSheet: GoogleAppsScript.Spreadsheet.Sheet, ui: Object ) {
  let width = SHEETS.MOBILE_MAIN.REFS.COL_NO_RECORD_NO - SHEETS.MOBILE_MAIN.REFS.COL_NO_CHECKBOXES + 1;
  let height = SHEETS.MOBILE_MAIN.REFS.ROW_NO_LAST_RECORD - SHEETS.MOBILE_MAIN.REFS.ROW_NO_FIRST_RECORD + 1
  let recordCheckBoxesData = mobileMainSheet.getRange( 
    SHEETS.MOBILE_MAIN.REFS.ROW_NO_FIRST_RECORD,
    SHEETS.MOBILE_MAIN.REFS.COL_NO_CHECKBOXES,
    height,
    width
  ).getValues();
  let recordStatusColourIndexes = mobileMainSheet.getRange( 
    SHEETS.MOBILE_MAIN.REFS.ROW_NO_FIRST_RECORD,
    SHEETS.MOBILE_MAIN.REFS.COL_NO_HIDDEN,
    height,
    width
  ).getValues();
  
  let selectedRecordNos: Array<number> = new Array();
  for( let ri = 0; ri < height; ri++ ) {
    if( recordCheckBoxesData[ ri ][ 0 ] && recordStatusColourIndexes[ ri ][ 0 ] < 9 ) {
      selectedRecordNos.push( recordCheckBoxesData[ ri ][ width-1 ] );
    }
  }

  return selectedRecordNos;
}

function checkForInstalledMobileTriggerOnWorkbook_( ) {

  //parse params
  let spreadsheet = SpreadsheetApp.getActive();
  
  //var to store answer
  let triggerExists = false;

  //check global settings to see if trigger is flagged as running
  let globalSettingsSheet = spreadsheet.getSheetByName( SHEETS.GLOBAL_SETTINGS.NAME );
  let mobileServiceActivatedBy = globalSettingsSheet.getRange( 
        SHEETS.GLOBAL_SETTINGS.REFS.ROW_NO_MOBILE_SERVICE_ON,
        SHEETS.GLOBAL_SETTINGS.REFS.COL_NO
  ).getValue();

  if( mobileServiceActivatedBy ) {
    let onEditTriggerFunctionName = "InstalledTrigger_MainWorkbook_OnEdit";
    Logger.log("Checking existance under this user of Installed onEdit Trigger: " + onEditTriggerFunctionName );
    
    let triggers = ScriptApp.getUserTriggers(spreadsheet);
    
    triggers.forEach(function (trigger) {
      if(trigger.getEventType() === ScriptApp.EventType.ON_EDIT && trigger.getHandlerFunction() === onEditTriggerFunctionName) {
        triggerExists = true; 
      }
    });
  }

  Logger.log( "triggerExists = " + triggerExists )

  if( !triggerExists ) {
    ShowInstallTriggerAlert( spreadsheet );
  }

  return triggerExists;
}

function SetupInstalledTriggersOnWorkbook( spreadsheet?: GoogleAppsScript.Spreadsheet.Spreadsheet|null ) {
  
  //parse params
  spreadsheet = (spreadsheet) ? spreadsheet : SpreadsheetApp.getActive();
  let installWasRequired = false;

  if(!checkForInstalledMobileTriggerOnWorkbook_() ) {
    Logger.log("Creating Installed onEdit Trigger: InstalledTrigger_MainWorkbook_OnEdit");
    ScriptApp.newTrigger('InstalledTrigger_MainWorkbook_OnEdit')
      .forSpreadsheet(spreadsheet)
      .onEdit()
      .create();
    Logger.log("Trigger Created");

    //update global settings to show trigger is installed by user
    let globalSettingsSheet = spreadsheet.getSheetByName( SHEETS.GLOBAL_SETTINGS.NAME );
    let mobileServiceStatusCell = globalSettingsSheet.getRange( 
          SHEETS.GLOBAL_SETTINGS.REFS.ROW_NO_MOBILE_SERVICE_ON,
          SHEETS.GLOBAL_SETTINGS.REFS.COL_NO
    );
    let mobileServiceActivatedBy = mobileServiceStatusCell.getValue();
    mobileServiceActivatedBy = ( mobileServiceActivatedBy ) ? mobileServiceActivatedBy : "";
    let activeUserEmail = Session.getActiveUser().getEmail();
    if( !mobileServiceActivatedBy.includes( activeUserEmail ) ) {
      mobileServiceActivatedBy += "|" + activeUserEmail;
      mobileServiceStatusCell.setValue( mobileServiceActivatedBy );
    }

    //startup mobile environment
    let mobileMainSheet = spreadsheet.getSheetByName( SHEETS.MOBILE_MAIN.NAME );
    let fakeMobileUi = getSpoofUIObject_FromAlertSheetDef( 
      mobileMainSheet, 
      SHEETS.MOBILE_MAIN.REFS.ALERT_BOX 
    );
    fakeMobileUi.alert( "Installation Complete", "Congratulations, The Mobile Workbook is installed and ready to use", fakeMobileUi.ButtonSet.OK, null, 5 );
    releaseLockOnMainSheet_(mobileMainSheet, 1 );

    let mobileInputSheet = spreadsheet.getSheetByName( SHEETS.MOBILE_THIS_RECORD.NAME );
    if( mobileInputSheet.isSheetHidden() ) {
      mobileInputSheet.showSheet();
    }
    if( mobileMainSheet.isSheetHidden() ) {
      mobileMainSheet.showSheet();
    }
    installWasRequired = true;
  }
  else {
    Logger.log("Trigger Already exists, no action required." );
  }
  return installWasRequired;

}

function ShowInstallTriggerAlert( spreadsheet?: GoogleAppsScript.Spreadsheet.Spreadsheet|null ) {

  //parse params
  spreadsheet = (spreadsheet) ? spreadsheet : SpreadsheetApp.getActive();

  let mobileMainSheet = spreadsheet.getSheetByName( SHEETS.MOBILE_MAIN.NAME );

  getLockOnMainSheet_(mobileMainSheet, 1 );

  let fakeMobileUi = getSpoofUIObject_FromAlertSheetDef( 
        spreadsheet.getSheetByName( SHEETS.MOBILE_MAIN.NAME ), 
        SHEETS.MOBILE_MAIN.REFS.ALERT_BOX 
  );

  fakeMobileUi.alert( "Please Activate the Mobile Version", "To use the Mobile version of your Workbook,\n" +
        "you must first install a background service.\n\nYou must do this from your Computer and not your " + 
        "Phone/Tablet, Sorry.\n\nAccess this Workbook on your computer and go to [LSA Menu] --> " + 
        "[Mobile Version] --> [Activate].\n\nThen press OK once this is done.", 
        fakeMobileUi.ButtonSet.OK, "checkForInstalledMobileTriggerOnWorkbook_" );
}

function InstallMobileServiceFromLSAMenu( ui?: any ) {
  ui = ( ui ) ? ui : SpreadsheetApp.getUi();

  let spreadsheet = SpreadsheetApp.getActive();
  let mobileMainSheet = spreadsheet.getSheetByName( SHEETS.MOBILE_MAIN.NAME );

  let installed = SetupInstalledTriggersOnWorkbook( spreadsheet );

  let mobileInputSheet = spreadsheet.getSheetByName( SHEETS.MOBILE_THIS_RECORD.NAME );
  if( mobileInputSheet.isSheetHidden() ) {
    mobileInputSheet.showSheet();
  }
  if( mobileMainSheet.isSheetHidden() ) {
    mobileMainSheet.showSheet();
  }
  releaseLockOnMainSheet_(mobileMainSheet, 1 );

  if( installed ) {
    ui.alert( "Mobile Workbook Activated",
              "The Mobile Version of the Workbook is now setup and ready for use. The Sheets: 'Mobile - Main' and 'Mobile - Input' are now available for you to use.", 
              ui.ButtonSet.OK );
  }
  else { 
    ui.alert( "Mobile Workbook Already Activated",
              "The Mobile Version of the Workbook was already activated. You use the sheets 'Mobile - Main' and 'Mobile - Input' on your mobile or tablet using the Google Sheets App", 
              ui.ButtonSet.OK );
   }
}

function UnInstallMobileServiceFromLSAMenu( ui?: any ) {
  ui = ( ui ) ? ui : SpreadsheetApp.getUi();

  let spreadsheet = SpreadsheetApp.getActive();
  let mobileMainSheet = spreadsheet.getSheetByName( SHEETS.MOBILE_MAIN.NAME );

  let triggerExisted = false;

  //update global settings to show trigger is installed by user
  let globalSettingsSheet = spreadsheet.getSheetByName( SHEETS.GLOBAL_SETTINGS.NAME );
  let mobileServiceStatusCell = globalSettingsSheet.getRange( 
        SHEETS.GLOBAL_SETTINGS.REFS.ROW_NO_MOBILE_SERVICE_ON,
        SHEETS.GLOBAL_SETTINGS.REFS.COL_NO
  );
  let mobileServiceActivatedBy = mobileServiceStatusCell.getValue();
Logger.log( "1 mobileServiceActivatedBy = " + mobileServiceActivatedBy );
  mobileServiceActivatedBy = ( mobileServiceActivatedBy ) ? mobileServiceActivatedBy : "";
Logger.log( "2 mobileServiceActivatedBy = " + mobileServiceActivatedBy );
  let activeUserEmail = Session.getActiveUser().getEmail();
Logger.log( "replacing: '" + "|" + activeUserEmail + "'" );
mobileServiceActivatedBy = mobileServiceActivatedBy.replace( "|" + activeUserEmail, "" );
Logger.log( "3 mobileServiceActivatedBy = " + mobileServiceActivatedBy );
  mobileServiceActivatedBy = ( mobileServiceActivatedBy ) ? mobileServiceActivatedBy : false;
Logger.log( "4 mobileServiceActivatedBy = " + mobileServiceActivatedBy );
  mobileServiceStatusCell.setValue( mobileServiceActivatedBy );

  //remove all accessible triggers
  let onEditTriggerFunctionName = "InstalledTrigger_MainWorkbook_OnEdit";
  Logger.log("Checking existance under this user of Installed onEdit Trigger: " + onEditTriggerFunctionName );
  let triggers = ScriptApp.getUserTriggers(spreadsheet);
  triggers.forEach(function (trigger) {
    if(trigger.getEventType() === ScriptApp.EventType.ON_EDIT && trigger.getHandlerFunction() === onEditTriggerFunctionName) {
      ScriptApp.deleteTrigger( trigger );
      triggerExisted = true;
    }
  });

  let mobileInputSheet = spreadsheet.getSheetByName( SHEETS.MOBILE_THIS_RECORD.NAME );
  if( !mobileInputSheet.isSheetHidden() ) {
    mobileInputSheet.hideSheet();
  }
  if( mobileMainSheet.isSheetHidden() ) {
    mobileMainSheet.showSheet();
  }
  releaseLockOnMainSheet_(mobileMainSheet, 1 );

  if( triggerExisted ) {
    ui.alert( "Mobile Workbook Deactivated",
              "The Mobile Version of the Workbook was deactivated successfully .", 
              ui.ButtonSet.OK );
    ShowInstallTriggerAlert( spreadsheet );
  }
  else { 
    if( mobileServiceActivatedBy ) {
      ui.alert( "Unable to deactivate Mobile Workbook", 
                "The Mobile Version of the Workbook was activated by a different user/users: ("+mobileServiceActivatedBy.substring(1)+") .\n Only that/those user/users can deactivate, sorry.", 
                ui.ButtonSet.OK );
    } 
    else {
      ui.alert( "Mobile Workbook Already Deactivated", 
                "The Mobile Version of the Workbook was already inactive. No action was required.", 
                ui.ButtonSet.OK );
    }
   }
}