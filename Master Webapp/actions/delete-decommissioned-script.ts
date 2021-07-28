/**
 * Delete a workbook file if user calling webapp has permissions
 * 
 * @workbookFileId {FileId} The id of the workbook file to be deleted
 * @authenticatedUser {String} Email Address of User calling webapp
 * @returnData {Object} Struct containing 3 porperties set to success, error, and affected file id, used as JSON return deom web app
 * @onlyDeleteDecomissioned {boolen} if this is set to true it will only delete a workbook if it can been decomissioned according to its global settings.
 * @return null
 */
function DeleteWorkbook( workbookFileId, authenticatedUser, returnData, onlyAllowDeleteDecomissioned ) { 
  
  //parse params
  onlyAllowDeleteDecomissioned = (onlyAllowDeleteDecomissioned===false) ? false : true;

  Logger.log("Preparing to delete workbook with FileId: " + workbookFileId );

  var workbookFile = SpreadsheetApp.openById(workbookFileId);
  let workbookGlobalSettingsSheet = workbookFile.getSheetByName( SHEETS.GLOBAL_SETTINGS.NAME );
  Logger.log("File: " + workbookFile );

  //check it's not a master
  var thisFileIdFromSettings = GetMasterSettingsCellFromOldVersionWorkbook( workbookFile, workbookGlobalSettingsSheet, 
          SHEETS.GLOBAL_SETTINGS.REFS.HANDLE_THIS_FILES_ID, SHEETS.GLOBAL_SETTINGS.REFS.ROW_NO_THIS_FILES_ID ).getValue();
  if( thisFileIdFromSettings == "" ) {
    returnData.success = 0;
    returnData.errorMessage = "The Workbook with ID : " + workbookFileId + " cannot be deleted as it appears to be a Master" +
                              " Worksheet (Settings Cell B2 is blank);" +
    Logger.log( returnData.errorMessage );
    return;
  }

  if( onlyAllowDeleteDecomissioned && thisFileIdFromSettings == 
                                      GetMasterSettingsCellFromOldVersionWorkbook( workbookFile,
                                          workbookGlobalSettingsSheet, SHEETS.GLOBAL_SETTINGS.REFS.HANDLE_MASTERS_LINK_TO_THIS_CHILD,
                                          SHEETS.GLOBAL_SETTINGS.REFS.ROW_NO_MASTERS_LINK_TO_THIS_CHILD 
                                      ).getValue() ) {
    returnData.success = 0;
    returnData.errorMessage = "The Workbook with ID : " + workbookFileId + " cannot be deleted as it appears not to be " +
                              "decomissioned" +
    Logger.log( returnData.errorMessage );
    return;
  }
  Logger.log( "File is a valid state for deletion (onlyAllowDeleteDecomissioned = " + onlyAllowDeleteDecomissioned + ")" );

  //TO DO: CHANGE THIS TO USE THE AUTH FUNCTION IN Code.gs
  //if calling user is not the main dev account we have to authenticate
  if( authenticatedUser != Session.getEffectiveUser().getEmail() )
  {
    //user calling webapp is not master account
    var lsaAdminsGroup = GroupsApp.getGroupByEmail( GLOBAL_CONSTANTS.LSA_ADMINS_GOOGLE_GROUP_EMAIL );
    if( !lsaAdminsGroup.hasUser(authenticatedUser) ) {
      //user calling webapp is not master account or an 'LSA Admin'
      var workbooksUserEmail = GetMasterSettingsCellFromOldVersionWorkbook( workbookFile,
            workbookGlobalSettingsSheet, SHEETS.GLOBAL_SETTINGS.REFS.HANDLE_MAIN_USERS_EMAIL,
            SHEETS.GLOBAL_SETTINGS.REFS.ROW_NO_MAIN_USERS_EMAIL 
      ).getValue().trim();
      if( authenticatedUser != workbooksUserEmail ) {
        ///user calling webapp is not master account or an 'LSA Admin' or the workbook's main user
        returnData.success = 0;
        returnData.errorMessage = "You do not have access to delete this file; You ('" + authenticatedUser +
          "') are not the master account: '" + Session.getEffectiveUser().getEmail() + "', nor are you part of the google group:" + 
          "'lsa-administrators@wlc.ac.uk'\n\nNor are you the workbooks main user: '" + workbooksUserEmail + "'.";
        Logger.log( returnData.errorMessage );
        return;
      }
    }
  }
  Logger.log( "Authenticated user: " + authenticatedUser + " has permission to delete");

  Logger.log( "Trashing File Now.");
  DriveApp.getFileById(workbookFile.getId()).setTrashed(true);
  Logger.log( "Trashing Successful.");

  //FINISHED SUCCESSFULLY
  returnData.success = 1;
  returnData.errorMessage = null;
  returnData.affectedFileId = workbookFileId;
}