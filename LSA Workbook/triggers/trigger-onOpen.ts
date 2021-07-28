function onOpen( event ){
  //Logger.log( 'onOpen trigger commencing');

  //here were going to add an extra menu or two on the toolbal, but before we do
  //we need to know if this is the master spreadsheet and whether we're logged in as the developer or not
  var thisSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var settingsSheet = thisSpreadsheet.getSheetByName( SHEETS.GLOBAL_SETTINGS.NAME );
  var isMaster = isAMasterNotAChild( thisSpreadsheet, settingsSheet, false );
  var isDev    = currentUserIsADev( thisSpreadsheet, settingsSheet );

  // This line calls the SpreadsheetApp and gets its UI   
  // Or DocumentApp or FormApp.
  var ui = SpreadsheetApp.getUi();
 
  //These lines create the menu items and 
  // tie them to functions that all exists in this project's Apps Script .gs files
  var mainMenu = ui.createMenu('LSA Menu')
      .addItem('New Day Cleanup', 'ResetInputData')
      .addItem("Generate All Records of Support", 'ExportAllToSendRoSs')
      .addItem('Check Signatures for Today (Input Sheet)', 'RefreshAllFileModifiedDates_InputSheet')
      .addItem('Check all Signatures (My Files Sheet)', 'RefreshAllFileModifiedDates_MyFilesSheet')
      .addItem('Send Learner Reminder Emails', 'ReminderEmailForSelectedRecords')
      .addSeparator()
      .addSubMenu( ui.createMenu('Actions for Selected Records (Input Sheet)')
          .addItem('Generate only Selected Record(s) of Support', 'SendSelectedRecords')
          .addItem("Flag Selected as 'No Signature Required'", 'FlagSelectedRecordsAsAutoSign')
          .addItem('Delete Selected Record(s) of Support', 'DeletePDFFromSelectedRecords')
      )
      .addSeparator()
      .addItem('Check for new Messages', 'CheckForAccouncements')
      .addSeparator()
      .addSubMenu( ui.createMenu('Show / Hide Settings')
          .addItem('Show Settings Sheets', 'ShowSettingsWorksheets')
          .addItem('Hide Settings Sheets', 'HideSettingsWorksheets')
      )
      .addSubMenu( ui.createMenu('Help & Support')
          .addItem('View our User Guides', 'ShowGetHelpPopup')
          .addItem('Email Support', 'ShowEmailSupportPopup')
      )
      .addSeparator()
      .addSubMenu( ui.createMenu('Mobile / Tablet Workbook')
          .addItem('Activate Mobile Version', 'InstallMobileServiceFromLSAMenu')
          .addItem('Deactivate Mobile Version', 'UnInstallMobileServiceFromLSAMenu')
      )
      .addSeparator()
      .addSubMenu( ui.createMenu('Advanced Functions')
          .addItem('Export All To Single Unsigned PDF', 'ExportAllRecordsToLocalPDF')
      );

  //conditional submenus
  if( isDev ) {
    mainMenu.addSeparator()
    .addSubMenu( ui.createMenu('Development Menu')
        .addItem('Test Master RoS API Webapp', 'TestMasterRoSAPIWebapp')
        .addItem('Test Check for Announcements', 'CheckForAccouncements')
        .addItem( 'Test Overrulling Data Validation on Input Sheet', 'TestDiasablingValidationRulesBeforeCopying' )
        .addItem( 'Test Writing Null in a Range', 'TestWritingNullsInRange' )
        .addItem( 'Test setting up Installed Triggers', 'SetupInstalledTriggersOnWorkbook' )
        .addItem( 'Invalidate User Authentication', 'InvalidateUserAuthentication' )
        .addItem( 'Audit Protected Ranges', 'AuditSpreadSheetProtections' )
        .addItem( 'Refresh Settings Learner Data', 'RefreshSettingsLearnerSheetDataFromMasterDatabaseFromChild' )
    );
  }
  if( !isMaster ) {
      mainMenu.addSeparator()
      .addSubMenu( ui.createMenu('Update this Workbook' )
          .addItem( 'Upgrade to Latest Version', 'CheckForUpdates' )
          .addItem( 'Repair this Workbook', 'CheckForRepair' )
      );
  }
  mainMenu.addToUi();

  if( isMaster )
  {
    //extra functionality ONLY.if it's the master spreadsheet
    var masterMenu = ui.createMenu('MASTER Menu')
    .addItem('Send Announcement to All LSAs', 'AnnounceToAllLSAs')
    .addSeparator()
    .addSubMenu( ui.createMenu('Actions for Selected LSAs')
      .addItem('Create Workbook(s) for New Starter(s)', 'CreateNewChildWorkbookForSelectedLSAs')
      .addItem('Send Announcement to Selected LSAs', 'AnnounceToSelectedLSAs')
      .addItem('Open an LSAs Folder', 'OpenLSAsDirectory')
      .addItem('Delete an LSAs Folder', 'DeleteLSAsDirectory')
    )
    .addSeparator()
    .addSubMenu( ui.createMenu('Deployment Menu')
      .addItem("Deploy New Version!", 'DeployNewVersion')
      .addItem("Rebuild after RoS Form Template Changes", 'RegenerateFromMasterTemplate')
      .addItem("Remove Test Data Ready For Deployment", 'ClearTestDataReadyForDeployment')
      .addItem("Hide Non-Master Sheets Ready For Deployment", 'HideNonMasterSheetsReadyForDeployment')
      .addItem("Push Help Link Changes To All Children", 'PushHelpLinkChangesToAllChildren')
      .addItem("Push Admin's Global Settings to All Children", 'PushAdminGlobalSettingsChangesToAllChildren')
      .addItem("Push DeployId(s) To All Children", 'PushDeployIdChangesToAllChildren')
    );

    //conditional dev submenu
    if( isDev ) {
      masterMenu.addSeparator()
      .addSubMenu( ui.createMenu('Development Menu')
          .addItem('Test Master RoS API Webapp', 'TestMasterRoSAPIWebapp')
          .addItem('Test Master Push Service', 'TestPushToChildWorksheets')
          .addItem('Copy Protections From Another Workbook', 'CopyProtectionsFromAnotherWorkbook')
          .addItem('Send Test Learner Email (Test HTML Template)', 'TestLearnerEmail')
          .addItem('Populate Test Area LSA Workbooks From Live', 'CopyLiveWorkbooksToTest' )
          .addItem('Populate Test Area Johns Workbooks From Live', 'CopyJohnMilnersWorkbooksToTest' )
          .addItem('Test Create Learner RoS From Sheets Template', 'TestCreateLearnerRoSFromSheetsTemplate' ) 
      );
    }

    masterMenu.addToUi();
  }
  else {
    //extra functionality ONLY if it's the child spreadsheet

    //check we are not in a copied or previous versions copy of the spreadsheet
    //it has to be the one linked to from master
    //this line uopdated the 'this files id' setting to this files id - in case we're working with a copy or a decomissioned wb
    refreshMasterLinkFilesIdsInGlobalSettings( thisSpreadsheet, settingsSheet );
    //compare newly refreshed 'this files id' setting with the 'this files id according to the master' setting.
    if( !isCorrectInstance( settingsSheet ) )
    {
      //add extra menu to allow user to delete workbook
      var trashMenu = ui.createMenu('Decomissioned Workbook Menu')
      .addItem('Send to Trash Bin', 'TrashThisFileAfterUpgrade')
      .addToUi();

      thisSpreadsheet.setActiveSheet( thisSpreadsheet.getSheetByName( SHEETS.INPUT.NAME ) );
      ui.alert( 'Wrong copy of Workbook', 
                "You must use the latest version of your workbook, and not a copy of it.\n\n" +
                "It looks like you either:\n" +
                "  - Made a copy of your workbook - copies are not allowed.\n" +
                "  - Opened the old decomissioned spreadhseet after an upgrade from a pervious version.\n\n" +
                "It looks like the correct file is this one (you can copy and paste this link):\n" +
                "https://docs.google.com/spreadsheets/d/" + getCorrectInstance( settingsSheet ) + "\n\n" +
                "If you switch to the 'Input' sheet, you will also see a hyperlink to this new file in the red bar across the top\n\n" +
                "Please use the 'Decomissioned Workbook Menu' at the top to send this workbook to the Trash Bin", 
                ui.ButtonSet.OK );
    }

    //Announcement code
    CheckForAccouncements( thisSpreadsheet, settingsSheet );
  }

  checkForInstalledMobileTriggerOnWorkbook_();
  //Logger.log( 'onOpen trigger finishing');
}

function InvalidateUserAuthentication() {
  ScriptApp.invalidateAuth();
}

function isCorrectInstance( globalSettingsSheet ) {
  globalSettingsSheet = ( globalSettingsSheet ) 
                        ? globalSettingsSheet 
                        : SpreadsheetApp.getActive().getSheetByName( SHEETS.GLOBAL_SETTINGS.NAME );
  Logger.log( 'isCorrectInstance started, globalSettingsSheet = ' + globalSettingsSheet );

  var thisFilesId = globalSettingsSheet.getRange(
        SHEETS.GLOBAL_SETTINGS.REFS.ROW_NO_THIS_FILES_ID, 
        SHEETS.GLOBAL_SETTINGS.REFS.COL_NO 
      ).getValue();

  var mastersLinkToThisChild = globalSettingsSheet.getRange(
        SHEETS.GLOBAL_SETTINGS.REFS.ROW_NO_MASTERS_LINK_TO_THIS_CHILD, 
        SHEETS.GLOBAL_SETTINGS.REFS.COL_NO 
      ).getValue();
  
  Logger.log('isCorrectInstance finished, returning:' + ( thisFilesId == mastersLinkToThisChild ) );  
  return ( thisFilesId == mastersLinkToThisChild );
}

function getCorrectInstance( globalSettingsSheet ) {
  globalSettingsSheet = ( globalSettingsSheet ) 
                        ? globalSettingsSheet 
                        : SpreadsheetApp.getActive().getSheetByName( SHEETS.GLOBAL_SETTINGS.NAME );

  var mastersLinkToThisChild = globalSettingsSheet.getRange(
      SHEETS.GLOBAL_SETTINGS.REFS.ROW_NO_MASTERS_LINK_TO_THIS_CHILD, 
      SHEETS.GLOBAL_SETTINGS.REFS.COL_NO 
    ).getValue();
  
  return mastersLinkToThisChild;
}

function refreshMasterLinkFilesIdsInGlobalSettings( spreadsheet, globalSettingsSheet ) {
  spreadsheet = ( spreadsheet ) ? spreadsheet : SpreadsheetApp.getActive();
  globalSettingsSheet = ( globalSettingsSheet ) ? globalSettingsSheet : spreadsheet.getSheetByName( SHEETS.GLOBAL_SETTINGS.NAME );

  globalSettingsSheet.getRange(
    SHEETS.GLOBAL_SETTINGS.REFS.ROW_NO_THIS_FILES_ID, 
    SHEETS.GLOBAL_SETTINGS.REFS.COL_NO 
  ).setValue( spreadsheet.getId() );
}
