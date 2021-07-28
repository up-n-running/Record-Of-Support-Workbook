function userHasAuthenticated_( spreadsheet?: GoogleAppsScript.Spreadsheet.Spreadsheet ): boolean {
  
  //parse params
  spreadsheet = (spreadsheet) ? spreadsheet : SpreadsheetApp.getActive();

  //check authenticated status
  Logger.log( "Checking ScriptApp.getAuthorizationInfo" );
  let authInfo = ScriptApp.getAuthorizationInfo(ScriptApp.AuthMode.LIMITED );
  Logger.log( "Authorised = " + (authInfo.getAuthorizationStatus() === ScriptApp.AuthorizationStatus.NOT_REQUIRED) );
  Logger.log( "authInfo.getAuthorizationStatus() = " + authInfo.getAuthorizationStatus() );
  Logger.log( "authInfo.getAuthorizationUrl() = " + authInfo.getAuthorizationUrl() );

  return authInfo.getAuthorizationStatus() === ScriptApp.AuthorizationStatus.NOT_REQUIRED;
}

function getAuthenticationURL( spreadsheet?: GoogleAppsScript.Spreadsheet.Spreadsheet ): string {
  
  //parse params
  spreadsheet = (spreadsheet) ? spreadsheet : SpreadsheetApp.getActive();

  //check authenticated status
  let authInfo = ScriptApp.getAuthorizationInfo(ScriptApp.AuthMode.FULL);
  return authInfo.getAuthorizationUrl();
}