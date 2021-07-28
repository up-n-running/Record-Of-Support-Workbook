/**
 */
function DeleteRecordOfSupport( lsaWorkbookFileId, rosFileId, authenticatedUser, returnData ) { 

  Logger.log("Preparing to delete Record of Support with FileId: " + lsaWorkbookFileId );

  let lsaWorkbookSpreadsheet = SpreadsheetApp.openById(lsaWorkbookFileId);
  Logger.log("LSAs Workbook File: " + lsaWorkbookSpreadsheet.getName() );
  let lsaWorkbookGlobalSettingsSheet = lsaWorkbookSpreadsheet.getSheetByName( SHEETS.GLOBAL_SETTINGS.NAME );
  let lsaWorkbookMyFilesSheet = lsaWorkbookSpreadsheet.getSheetByName( SHEETS.MY_FILES.NAME );

  //authenticate
  let accessLevel = getAccessLevel_(lsaWorkbookSpreadsheet, lsaWorkbookGlobalSettingsSheet, authenticatedUser );
  if( accessLevel == ACCESS_LEVELS.NONE ) {
    returnData.success = 0;
    returnData.errorMessage = "You do not have access to delete the Record of Support file; You ('" + authenticatedUser +
      "') are not the master account: '" + Session.getEffectiveUser().getEmail() + "', nor are you part of the google group:" + 
      "'lsa-administrators@wlc.ac.uk'\n\nNor are you the workbook's (id: "+lsaWorkbookFileId+") main user.";
    Logger.log( returnData.errorMessage );
    return;
  }
  Logger.log( "User has access level: " + accessLevel )

  //make sure the file we're trying to delete is in the lsas worksheet - this ensures they're not deleting someone else's files
  let fileFoundRowNum = findInColumn( lsaWorkbookMyFilesSheet, rosFileId, SHEETS.MY_FILES.REFS.COL_NO_USERS_FILE_ID,
                                      SHEETS.MY_FILES.REFS.ROW_NO_FIRST_FILE, SHEETS.MY_FILES.REFS.ROW_NO_LAST_FILE );
  if( fileFoundRowNum < 0 ) {
    returnData.success = 0;
    returnData.errorMessage = "You do not have access to delete the Record of Support file; The file id ("+rosFileId+")" + 
    "could not be found in your LSAs workbook on the My Files sheet on column "+SHEETS.MY_FILES.REFS.COL_NO_USERS_FILE_ID+".";
    Logger.log( returnData.errorMessage );
    return;
  }
  
  //if we are the LSA and no higher priviledge then we need to chack how old the file is before we can delete
  //and we do this by checking the no of minutes detting on the master as we know which cell that is is for definite
  let allowedMaxFileAge = null;
  if( accessLevel == ACCESS_LEVELS.LSA ) {
    //get the master file that we know is on the current version so we can read it's settings easier
    let mastersFileId = GetMasterSettingsCellFromOldVersionWorkbook( lsaWorkbookSpreadsheet, lsaWorkbookGlobalSettingsSheet, 
            SHEETS.GLOBAL_SETTINGS.REFS.HANDLE_MASTER_FILE_ID, SHEETS.GLOBAL_SETTINGS.REFS.ROW_NO_MASTER_FILE_ID ).getValue();
    let masterWorkbookSpreadsheet = null;
    let masterWorkbookGlobalSettingsSheet = null;
    //if we're calling from the master then we dont need to get it
    if( mastersFileId == "" ) {
      mastersFileId = lsaWorkbookFileId;
      masterWorkbookSpreadsheet = lsaWorkbookSpreadsheet
    }
    else {
      masterWorkbookSpreadsheet = SpreadsheetApp.openById(mastersFileId);
    }
    masterWorkbookGlobalSettingsSheet = masterWorkbookSpreadsheet.getSheetByName( SHEETS.GLOBAL_SETTINGS.NAME );
    allowedMaxFileAge = masterWorkbookGlobalSettingsSheet.getRange (
      SHEETS.GLOBAL_SETTINGS.REFS.ROW_NO_MAX_ROS_DEL_MINS,
      SHEETS.GLOBAL_SETTINGS.REFS.COL_NO
    ).getValue();
  }

  //get the file ready to be deleted
  let rosFileToDelete = null;
  let fileAccessError = "DriveApp.getFileById( rosFileId ) returned null";
  try {
    rosFileToDelete = DriveApp.getFileById( rosFileId );
  }
  catch( e ) {
    fileAccessError = e;
  }
  if( rosFileToDelete == null ) {
    returnData.success = 0;
    returnData.errorMessage = "The RoS with ID : " + rosFileId + " cannot be read from the shared drive\n\n" +
                              "The returned error was:\n" + fileAccessError
    Logger.log( returnData.errorMessage );
    return;
  }

  //file age check
  if( allowedMaxFileAge != null ) {
    let ageInMillis = (new Date()).getTime() - rosFileToDelete.getDateCreated().getTime();
    let ageInMins = ageInMillis / (1000*60);
    if( ageInMins > allowedMaxFileAge ) {
      returnData.success = 0;
      returnData.errorMessage = "Sorry, you only have access to delete RoS Files created within the last "+allowedMaxFileAge+
            " minutes.\n\nThe file was actually created "+Math.floor( ageInMins )+" minutes ago\n\n" +
            "If you have made an error in the RoS then please generate a second RoS for the same period, and in it, explain that " +
            "The first one was made in error and should be discounted.\n\n" + 
            "If you really need to delete it then please ask your manager or the support team to run this function on your Workook " + 
            "from their account as they have greater access.";
      Logger.log( returnData.errorMessage );
      return;
    }
  }

  Logger.log( "Authenticated user: " + authenticatedUser + " has permission to delete RoS");

  Logger.log( "Trashing File Now.");
  rosFileToDelete.setTrashed(true);
  Logger.log( "Trashing Successful.");

  //FINISHED SUCCESSFULLY
  returnData.success = 1;
  returnData.errorMessage = null;
  returnData.affectedFileId = rosFileId;
}

function GenerateManualSignatureRosFile( lsaWorkbookSpreadsheetId, rosPreviewSheetName, destLearnerRosFolderId,
     learnerRoSTemplateFileId, destFileName, learnerAccessEmailAddress, learnerSignatureFileId, authenticatedUser, returnData ) {

  let lsaWorkbookSpreadsheet = SpreadsheetApp.openById( lsaWorkbookSpreadsheetId );
  let lsaWorkbookGlobalSettingsSheet = lsaWorkbookSpreadsheet.getSheetByName( SHEETS.GLOBAL_SETTINGS.NAME );

  //authenticate
  let accessLevel = getAccessLevel_(lsaWorkbookSpreadsheet, lsaWorkbookGlobalSettingsSheet, authenticatedUser );
  if( accessLevel == ACCESS_LEVELS.NONE ) {
    returnData.success = 0;
    returnData.errorMessage = "You do not have access to delete the Record of Support file; You ('" + authenticatedUser +
      "') are not the master account: '" + Session.getEffectiveUser().getEmail() + "', nor are you part of the google group:" + 
      "'lsa-administrators@wlc.ac.uk'\n\nNor are you the workbook's (id: "+lsaWorkbookSpreadsheetId+") main user.";
    Logger.log( returnData.errorMessage );
    return;
  }
  Logger.log( "User has access level: " + accessLevel );

  //setup objects ready
  let sourceRosPreviewSheet = lsaWorkbookSpreadsheet.getSheetByName( rosPreviewSheetName );
  let destLearnerRosFolder = DriveApp.getFolderById( destLearnerRosFolderId );
  let rosFile = createLearnerRoSFromSheetsTemplate_( lsaWorkbookSpreadsheet, sourceRosPreviewSheet, destLearnerRosFolder, 
        learnerRoSTemplateFileId, destFileName, learnerAccessEmailAddress, learnerSignatureFileId );

  //FINISHED SUCCESSFULLY
  returnData.success = 1;
  returnData.errorMessage = null;
  returnData.affectedFileId = ( rosFile ) ? rosFile.getId() : null;
}