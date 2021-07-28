function generatePDFBlob( spreadSheet, optSheetId, ui ) {
  
  //additional parameters for exporting the sheet as a pdf
  var url_ext = 'export?exportFormat=pdf&format=pdf'   //export as pdf

      // Print either the entire Spreadsheet or the specified sheet if optSheetId is provided
      + '&id=' + spreadSheet.getId()
      + (optSheetId ? ('&gid=' + optSheetId) : '')

      // following parameters are optional...
      + '&size=7'           // paper size
      + '&portrait=true'    // orientation, false for landscape
      + '&fitw=true'        // fit to window, false for actual size
      + '&sheetnames=false&printtitle=false&pagenumbers=false'  //hide optional headers and footers
      + '&gridlines=false'  // hide gridlines
      //+ '&gid=1650003668&ir=false&ic=false&r1=0&r2=40&c1=0&c2=11'  // only export relevant cells to save on PDF file size
      + '&fzr=false';       // do not repeat row headers (frozen rows) on each page

Logger.log( url_ext );
  
  var options = {
    headers: {
      'Authorization': 'Bearer ' +  ScriptApp.getOAuthToken()
    },
    muteHttpExceptions: true
  }
  var response = UrlFetchApp.fetch("https://docs.google.com/spreadsheets/" + url_ext, options);
  if( response.getResponseCode() == 200 ) {
    //success
    return response.getBlob();
  }
  else if( response.getResponseCode() == 429 ){
    ui.alert( "PDF Export - Temporary Failure", "One record failed because too many requests have been made to the PDF export service in quick succession.\n\nDon't worry, it's easy to recover from this. Press OK to carry on as usual and export the remaining records, then 'Generate Records of Support' again to retry this failed export.\n\nThe system should pick up the failed export the second time around and it will catch up with the other successfully exported records with no harm done.", ui.ButtonSet.OK );
    return null;
  }
  else {
    let alertMessage =  HtmlService.createHtmlOutput( response.getContentText() ).setWidth(600).setHeight(350);
    Logger.log( "ERROR EXPORTING TO PDF - SHOWING HTML ERROR MESSAGE MODAL" );
    SpreadsheetApp.getUi().showModalDialog(alertMessage, "Sorry, an error occured, please email support");
    return null;
  }
}

function savePDFsToGivenFolder( rootRoSDirectoryId, learnerObj, learnerName, startTime, duration, lsaName, recordDate, recordNo, isAutoSign,
                                optSpreadSheetId, optSheetId, ui: any ) {

  let spreadsheet = (optSpreadSheetId) ? SpreadsheetApp.openById(optSpreadSheetId) : SpreadsheetApp.getActiveSpreadsheet();

  let blob = generatePDFBlob( spreadsheet, optSheetId, ui );

  if( blob != null ) {

    uiSensitiveToast( spreadsheet, ui, 'Saving the PDF to the Shared Drive', 'Saving Record ' + recordNo );
    //var learnerFolder = createOrGetChildFolder( rootFolderId, learnerName );

    let learnerMonthFolder = getOrCreateLearnerYearAndMonthSubDir( spreadsheet, recordDate, learnerName, 
                                                                   learnerObj, rootRoSDirectoryId );

    //parse start time and duration information
    let lessonTimeObj = getlessonTimesObjectFromInputSheetValues( recordDate, startTime, duration );

    blob.setName( generateRosFileName( learnerName, lessonTimeObj, lsaName, isAutoSign, true ) );

    var pdfFile = learnerMonthFolder.createFile(blob);
    return pdfFile;
  }
  else {
    return null;
  }
}