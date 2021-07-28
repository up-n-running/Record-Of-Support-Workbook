function getThisAppsURL() {
  Logger.log( ScriptApp.getService().getUrl() ); 
}

function doGet(e){
  return performAction(e);
}

function doPost(e){
  return performAction(e);
}

function performAction(e){
  
  //parse url params
  var action              = ( e.parameter.action && e.parameter.action !== "null" ) ? e.parameter.action : null;
  var usersSpreadsheetId  = ( e.parameter.usersSpreadsheetId && e.parameter.usersSpreadsheetId !== "null" ) ? 
                              e.parameter.usersSpreadsheetId : null;
  var usersEmailAddress   = ( e.parameter.usersEmailAddress && e.parameter.usersEmailAddress !== "null" ) ? 
                              e.parameter.usersEmailAddress : null;
  var masterSpreadsheetId = ( e.parameter.masterSpreadsheetId && e.parameter.masterSpreadsheetId !== "null" ) ? 
                              e.parameter.masterSpreadsheetId : null;
  var rosFileId           = ( e.parameter.rosFileId && e.parameter.rosFileId !== "null" ) ? e.parameter.rosFileId : null;
  var rowNo               = ( e.parameter.rowNo && e.parameter.rowNo !== "null" ) ? e.parameter.rowNo : null;

  //for generate stored signature ROS functionality
  var rosPreviewSheetName = ( e.parameter.rosPreviewSheetName && e.parameter.rosPreviewSheetName !== "null" ) ? 
                              e.parameter.rosPreviewSheetName : null;
  var learnerRosFolderId  = ( e.parameter.learnerRosFolderId && e.parameter.learnerRosFolderId !== "null" ) ? 
                              e.parameter.learnerRosFolderId : null;
  var rosTemplateFileId   = ( e.parameter.rosTemplateFileId && e.parameter.rosTemplateFileId !== "null" ) ? 
                              e.parameter.rosTemplateFileId : null;
  var destFileName        = ( e.parameter.destFileName && e.parameter.destFileName !== "null" ) ? 
                              e.parameter.destFileName : null;
  var learnerEmailAddress = ( e.parameter.learnerEmailAddress && e.parameter.learnerEmailAddress !== "null" ) ? 
                              e.parameter.learnerEmailAddress : null;
  var signatureFileId     = ( e.parameter.signatureFileId && e.parameter.signatureFileId !== "null" ) ? 
                              e.parameter.signatureFileId : null;

  //find who called the webapp
  var authenticatedUser = Session.getActiveUser().getEmail();

Logger.log('WEBAPP CALLED, EXECUTING AS: ' + Session.getEffectiveUser().getEmail() );
Logger.log('authenticatedUser: '             + authenticatedUser );
Logger.log('Parameter: action: '             + action);
Logger.log('Parameter: usersSpreadsheetId: ' + usersSpreadsheetId);
Logger.log('Parameter: usersEmailAddress: '  + usersEmailAddress);
Logger.log('Parameter: masterSpreadsheetId: '+ masterSpreadsheetId);
Logger.log('Parameter: rosFileId: '          + rosFileId);
Logger.log('Parameter: rowNo: '              + rowNo);
Logger.log('Parameter: rosPreviewSheetName: '+ rosPreviewSheetName);
Logger.log('Parameter: learnerRosFolderId: ' + learnerRosFolderId);
Logger.log('Parameter: rosTemplateFileId: '  + rosTemplateFileId);
Logger.log('Parameter: destFileName: '       + destFileName);
Logger.log('Parameter: learnerEmailAddress: '+ learnerEmailAddress);
Logger.log('Parameter: signatureFileId: '    + signatureFileId);
  
  var returnData = {
    "success"        : 0,
    "errorMessage"   : "Invalid Webapp Parameter Combination - No action has been taken",
    "affectedFileId" : null
  };
    
  if( action == "update" && usersSpreadsheetId != null ) {
    PerformUpdates( MODES.UPDATE.CODE, usersSpreadsheetId, returnData, false, null, null );
  }
  else if( action == "repair" && usersSpreadsheetId != null ) {
    PerformUpdates( MODES.REPAIR.CODE, usersSpreadsheetId, returnData, false, null, null );
  }
  else if( action == "create" && usersEmailAddress != null && masterSpreadsheetId != null ) {
    PerformUpdates( MODES.CREATE.CODE, null, returnData, false, usersEmailAddress, masterSpreadsheetId );
  }
  else if( action == "clean-decommissioned-workbook" && usersSpreadsheetId != null && authenticatedUser != null ) {
    DeleteWorkbook( usersSpreadsheetId, authenticatedUser, returnData, true );
  }  
  else if( action == "delete-lsa" && masterSpreadsheetId != null && authenticatedUser != null &&
           usersEmailAddress != null && rowNo != null) {
    DeleteLSARecordAndFolder( rowNo, masterSpreadsheetId, usersEmailAddress, authenticatedUser, returnData );
  } 
  else if( action == "delete-ros" && usersSpreadsheetId != null && rosFileId != null && authenticatedUser != null ) {
    DeleteRecordOfSupport( usersSpreadsheetId, rosFileId, authenticatedUser, returnData );
  }  
  else if( action == "generate-ros-stored-signature" && usersSpreadsheetId != null && authenticatedUser != null && 
           rosPreviewSheetName != null && learnerRosFolderId != null && rosTemplateFileId != null && destFileName != null && 
           learnerEmailAddress != null && signatureFileId != null ) {
    GenerateManualSignatureRosFile( usersSpreadsheetId, rosPreviewSheetName, learnerRosFolderId, rosTemplateFileId, 
          destFileName, learnerEmailAddress, signatureFileId, authenticatedUser, returnData );
  }  
  else {
    Logger.log( returnData.errorMessage );  
  }

  Logger.log(returnData);  
  
  //return the JSON response back to the app that called this API
  var JSONString = JSON.stringify(returnData);
  var JSONOutput = ContentService.createTextOutput(JSONString);
  JSONOutput.setMimeType(ContentService.MimeType.JSON);
  return JSONOutput;
}


const ACCESS_LEVELS = {
  DEV   : 3,
  ADMIN : 2,
  LSA   : 1,
  NONE  : 0
}
function getAccessLevel_( lsaWorkbookSpreadsheet, workbookGlobalSettingsSheet, authenticatedUserEmail ) {
  //if calling user is not the main dev account we have to autnenticate
  if( authenticatedUserEmail == Session.getEffectiveUser().getEmail() )
  {
    return ACCESS_LEVELS.DEV;
  }
  else {
    //user calling webapp is not master account
    var lsaAdminsGroup = GroupsApp.getGroupByEmail( GLOBAL_CONSTANTS.LSA_ADMINS_GOOGLE_GROUP_EMAIL );
    if( lsaAdminsGroup.hasUser(authenticatedUserEmail) )
    {
      return ACCESS_LEVELS.ADMIN;
    }
    else if (!lsaWorkbookSpreadsheet && !workbookGlobalSettingsSheet ) {
      return ACCESS_LEVELS.NONE;
    }
    else {
      //user calling webapp is not master account or an 'LSA Admin'
      var workbooksUserEmail = GetMasterSettingsCellFromOldVersionWorkbook( lsaWorkbookSpreadsheet,
            workbookGlobalSettingsSheet, SHEETS.GLOBAL_SETTINGS.REFS.HANDLE_MAIN_USERS_EMAIL,
            SHEETS.GLOBAL_SETTINGS.REFS.ROW_NO_MAIN_USERS_EMAIL 
      ).getValue().trim();
      if( authenticatedUserEmail == workbooksUserEmail ) {
        return ACCESS_LEVELS.LSA;
      }
      else{
        return ACCESS_LEVELS.NONE;
      }
    }
  }
}
