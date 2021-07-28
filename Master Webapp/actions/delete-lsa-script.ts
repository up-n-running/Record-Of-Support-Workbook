function DeleteLSARecordAndFolder( lsasRowNo: number, masterFileId: string, lsasEmailAddress: string,
        authenticatedUser, returnData ) {

  //get users child spreadsheet's linked master (should be the latest version)
  let masterFile = DriveApp.getFileById( masterFileId );
  let masterSpreadsheet = SpreadsheetApp.open( masterFile );

  let globalSettingsSheet = masterSpreadsheet.getSheetByName( SHEETS.GLOBAL_SETTINGS.NAME );
  let lsaWorksheet = masterSpreadsheet.getSheetByName( SHEETS.MASTER_LSAS.NAME );

  //authenticate
  let accessLevel = getAccessLevel_(null, null, authenticatedUser );
  if( accessLevel == ACCESS_LEVELS.NONE || accessLevel == ACCESS_LEVELS.LSA ) {
    returnData.success = 0;
    returnData.errorMessage = "You do not have access to delete the LSAs Directory; You ('" + authenticatedUser +
      "') are not the master account: '" + Session.getEffectiveUser().getEmail() + "', nor are you part of the google group:" + 
      "'lsa-administrators@wlc.ac.uk'";
    Logger.log( returnData.errorMessage );
    return;
  }
  Logger.log( "User has access level: " + accessLevel )

  //get info about the LSA who has been seleted for deletion
  let lsaEmailFromMaster = lsaWorksheet.getRange( lsasRowNo, SHEETS.MASTER_LSAS.REFS.COL_NO_LSA_EMAIL ).getValue().trim();         
  let lsaFileId = lsaWorksheet.getRange( lsasRowNo, SHEETS.MASTER_LSAS.REFS.COL_NO_USERS_FILE_ID ).getValue();

  if( lsaEmailFromMaster !== lsasEmailAddress ) {
    returnData.success = 0;
    returnData.errorMessage = "Authentication Error. You requested that LSA on row number " + lsasRowNo + " be deleted and " +
          "specified the lsa's email address as: '" + lsasEmailAddress + "'.\n" +
          "However the actual LSAs email on row number " + lsasRowNo + " is: '" + lsaEmailFromMaster + "'."
    Logger.log( returnData.errorMessage );
    return;
  }

  let lsasParentFolder = getLSAsParentFolderFromLSAWorkbookId( masterSpreadsheet, null, globalSettingsSheet, lsaFileId );
  if( !lsasParentFolder ) {
    //get the root dir id for the LSAs - purely to populate the error message!
    let lsaRootDirectoryId = globalSettingsSheet.getRange( 
      SHEETS.GLOBAL_SETTINGS.REFS.ROW_NO_DIR_ID_LSAS, 
      SHEETS.GLOBAL_SETTINGS.REFS.COL_NO 
    ).getValue();
    returnData.success = 0;
    returnData.errorMessage = "COULD NOT FIND LSAs DIRECTORY:\n\nWe could not find a parent folder to the LSAs " + 
    "Workbook with file id: '" + lsaFileId + "' which, in turn, has a parent folder of the Master Root LSAs" + 
    "Directory with Directory Id: '" + lsaRootDirectoryId + "'." +
    Logger.log( returnData.errorMessage );
    return;
  }

  lsaWorksheet.getRange( lsasRowNo, SHEETS.MASTER_LSAS.REFS.COL_NO_LSA_EMAIL ).setValue("");
  lsaWorksheet.getRange( lsasRowNo, SHEETS.MASTER_LSAS.REFS.COL_NO_LSA_NAME ).setValue("");
  lsaWorksheet.getRange( lsasRowNo, SHEETS.MASTER_LSAS.REFS.COL_NO_USERS_VERSION ).setValue("");       
  lsaWorksheet.getRange( lsasRowNo, SHEETS.MASTER_LSAS.REFS.COL_NO_USERS_FILE_ID ).setValue("");
  lsasParentFolder.setTrashed( true );

  returnData.success = 1;
  returnData.errorMessage = null;
  returnData.affectedFileId = lsasParentFolder.getId();

}