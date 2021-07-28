function CallMasterRoSWebapp( 
      usersGlobalSettingsSheet,
      action, 
      usersSpreadsheetId, 
      usersEmailAddress, 
      masterSpreadsheetId, 
      rosFileId,
      rowNo: number,
      rosPreviewSheetName,
      learnerRosFolderId,
      rosTemplateFileId,
      destFileName,
      learnerEmailAddress,
      signatureFileId ) {

  Logger.log( "Calling CallMasterRoSWebapp" );
  Logger.log(ScriptApp.getService().getUrl());
  
  //parse usersGlobalSettingsSheet param in case it's missing
  usersGlobalSettingsSheet = ( usersGlobalSettingsSheet ) ? usersGlobalSettingsSheet :
                                 SpreadsheetApp.getActive().getSheetByName( SHEETS.GLOBAL_SETTINGS.NAME );

  var webappDeploymentId = usersGlobalSettingsSheet.getRange( 
        SHEETS.GLOBAL_SETTINGS.REFS.ROW_NO_WEBAPP_ALL_DEPLOYID ,
        SHEETS.GLOBAL_SETTINGS.REFS.COL_NO 
      ).getValue().trim();
  var webappExecUrlBase = "https://script.google.com/a/macros/wlc.ac.uk/s/" + 
                          webappDeploymentId +
                          "/exec"

  var url = webappExecUrlBase +
            '?action='              + encodeURIComponent( action              ) +
            '&usersSpreadsheetId='  + encodeURIComponent( usersSpreadsheetId  ) +
            '&usersEmailAddress='   + encodeURIComponent( usersEmailAddress   ) +
            '&masterSpreadsheetId=' + encodeURIComponent( masterSpreadsheetId ) +
            '&rosFileId='           + encodeURIComponent( rosFileId           ) +
            '&rowNo='               + encodeURIComponent( rowNo               ) +
            '&rosPreviewSheetName=' + encodeURIComponent( rosPreviewSheetName ) +
            '&learnerRosFolderId='  + encodeURIComponent( learnerRosFolderId  ) +
            '&rosTemplateFileId='   + encodeURIComponent( rosTemplateFileId   ) +
            '&destFileName='        + encodeURIComponent( destFileName        ) +
            '&learnerEmailAddress=' + encodeURIComponent( learnerEmailAddress ) +
            '&signatureFileId='     + encodeURIComponent( signatureFileId     );

  Logger.log( "url: " + url );

  var params:GoogleAppsScript.URL_Fetch.URLFetchRequestOptions = { 
    method: 'get',
    headers: { 
      Authorization: "Bearer " + ScriptApp.getOAuthToken() 
    },  
    muteHttpExceptions: false 
  };

  var response = UrlFetchApp.fetch( url, params );

  //var response = UrlFetchApp.fetch(url); // get api endpoint
  Logger.log( "Response code: " + response.getResponseCode() );
  Logger.log( "response.getContentText(): " + response.getContentText() );

  var responseText = response.getContentText(); // get the response content as text
  var json = null;
  try {
    json = JSON.parse(responseText); //parse text into json
    Logger.log(json); //log data to logger to check
  }
  catch( e )
  {
    debugCatchError( e );
    let alertMessage =  HtmlService.createHtmlOutput( responseText ).setWidth(600).setHeight(350);
    Logger.log( "WEBAPP RESPONSE WAS NOT JSON - SHOWING HTML ERROR MESSAGE MODAL" );
    SpreadsheetApp.getUi().showModalDialog(alertMessage, "Sorry, an error occured, please email support");
  }
  return json;
}
