function CheckForUpdates() {
  CheckForUpgradesOrRepair( true );
}

function CheckForRepair() {
  CheckForUpgradesOrRepair( false );
}

function CheckForUpgradesOrRepair( upgradeNotRepair ) {

  //parse input params
  upgradeNotRepair = ( upgradeNotRepair ) ? true : false;

  let spreadsheet = SpreadsheetApp.getActive();
  let ui = SpreadsheetApp.getUi();
  let settingsSheet = spreadsheet.getSheetByName( SHEETS.GLOBAL_SETTINGS.NAME );

  if( !isAuthorised_( spreadsheet, ui, settingsSheet, true, true, false, false ) ) { return; };

  let thisVersionNumber = settingsSheet.getRange(
        SHEETS.GLOBAL_SETTINGS.REFS.ROW_NO_VERSION_NO, 
        SHEETS.GLOBAL_SETTINGS.REFS.COL_NO 
      ).getValue();

  let rootAllLSAsFolderId = settingsSheet.getRange(
        SHEETS.GLOBAL_SETTINGS.REFS.ROW_NO_DIR_ID_LSAS, 
        SHEETS.GLOBAL_SETTINGS.REFS.COL_NO 
      ).getValue();

  let latestMastersVersionNumber = settingsSheet.getRange(
        SHEETS.GLOBAL_SETTINGS.REFS.ROW_NO_MASTERS_VERSION, 
        SHEETS.GLOBAL_SETTINGS.REFS.COL_NO 
      ).getValue();

  if( upgradeNotRepair && thisVersionNumber == latestMastersVersionNumber ) {
    ui.alert( "No Update Available", 
              "Congratulations, you are already on the latest version of the RoS Workbook\n\n" +
              "You are on version " + thisVersionNumber + "\n" +
              "The latest version available is " + latestMastersVersionNumber + "\n\n" +
              "We will send you an announcement as soon as the next version is available,\n" +
              "but thanks for checking in.",
              ui.ButtonSet.OK );
  }
  else if( !upgradeNotRepair && thisVersionNumber != latestMastersVersionNumber ) {
    let alertResponse = ui.alert( "Please Upgrade Instead", 
              "You cannot repair because you are not on the latest version of the RoS Workbook\n\n" +
              "You are on version " + thisVersionNumber + "\n" +
              "The latest version available is " + latestMastersVersionNumber + "\n\n" +
              "However, upgrading your Workbook will also perform a repair\n\n" +
              "Press OK To upgrade Your Workbook",
              ui.ButtonSet.OK_CANCEL );
    if( alertResponse == ui.Button.OK ) {
      CheckForUpgradesOrRepair( true );
    }
  }
  else {
    //generate HTML success alert with LSA Directory hyperlink embedded
    let thisFile = DriveApp.getFileById(spreadsheet.getId());
    let alertHTMLTemplate = HtmlService.createTemplateFromFile('html/html-modal-upgrade-repair');
    alertHTMLTemplate.upgradeNotRepair = upgradeNotRepair;
    alertHTMLTemplate.thisVersion = thisVersionNumber;
    alertHTMLTemplate.newVersion = latestMastersVersionNumber;
    alertHTMLTemplate.thisWorkbookUrl = thisFile.getUrl();
    alertHTMLTemplate.lsaDirectoryUrl = 
          getLSAsFolderFromWorkbookFile( thisFile, null, rootAllLSAsFolderId ).getUrl();
    let alertMessageHTML = alertHTMLTemplate.evaluate().getContent();
    let alertMessage =  HtmlService.createHtmlOutput( alertMessageHTML ).setWidth(600).setHeight(510);
    ui.showModalDialog( alertMessage, ( upgradeNotRepair ? 'An Upgrade is Available' : 'Repair your Workbook?' ) );
  }
}

function UpgradeWorkbook() {
  return UpgradeOrRepairWorkbook( true );
}

function RepairWorkbook() {
  return UpgradeOrRepairWorkbook( false );
}

function UpgradeOrRepairWorkbook( upgradeNotRepair ) {

  //parse input params
  upgradeNotRepair = ( upgradeNotRepair ) ? true : false;
  let actionDescIngCaps = ( upgradeNotRepair ) ? "Upgrading" : "Repairing";
  let actionDescCaps =    ( upgradeNotRepair ) ? "Upgrade" : "Repair";
  let actionDescNoCaps =  ( upgradeNotRepair ) ? "upgrade" : "repair";

Logger.log( "UpgradeOrRepairWorkbook called, upgradeNotRepair = " + upgradeNotRepair );

  let spreadsheet = SpreadsheetApp.getActive();
  let ui = SpreadsheetApp.getUi();

  spreadsheet.toast( "Please be patient, the "+actionDescNoCaps+" may take a while.", actionDescIngCaps + "...", 25 );
    
  var thisFilesId = spreadsheet.getId();

  var webAppReturnData = CallMasterRoSWebapp( null, ( upgradeNotRepair ? "update" : "repair" ), thisFilesId, null, null, 
        null, null, null, null, null, null, null, null );

  if( webAppReturnData == null ) {
    return null;
  }
  if( webAppReturnData.success != 1.0 )
  {
    ui.alert( "Could not " + actionDescCaps, webAppReturnData.errorMessage, ui.ButtonSet.OK );
    return null;
  }

  //setup the onedit installed trigger as this user
  let newSpreadsheet = SpreadsheetApp.openById( webAppReturnData.affectedFileId );
  //setupInstalledTriggersOnWorkbook_( newSpreadsheet );

  //switch back to Input Sheet so user can see it's gone red
  spreadsheet.setActiveSheet( spreadsheet.getSheetByName( SHEETS.INPUT.NAME ) );
  return newSpreadsheet.getUrl();
}

function TrashThisFileAfterUpgrade() {
  Logger.log( "TRASHING THIS FILE" );
  let thisFilesId = SpreadsheetApp.getActive().getId();
  let webAppReturnData = CallMasterRoSWebapp( null, "clean-decommissioned-workbook", thisFilesId, null, null, null, null, 
        null, null, null, null, null, null );
  if( webAppReturnData.success != 1.0 )
  {
    let ui = SpreadsheetApp.getUi();
    ui.alert( "Could Not Trash This Decomissioned Workbook", webAppReturnData.errorMessage, ui.ButtonSet.OK );
    return;
  }
  Logger.log( "TRASHED THIS FILE" );
}