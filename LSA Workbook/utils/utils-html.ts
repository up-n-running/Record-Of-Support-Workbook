function include(filename){
  return HtmlService.createHtmlOutputFromFile(filename)
    .getContent();
};

function serveErrorPage( errorTitle, errorMessage ) {
Logger.log( "serveErrorPage called, errorMessage = " + errorMessage );

  let htmlTemplate = HtmlService.createTemplateFromFile("html/error-message");
  htmlTemplate.errorTitle = errorTitle;
  htmlTemplate.errorMessage = errorMessage;
  
  let htmlOutput = htmlTemplate.evaluate();
  htmlOutput.setTitle(errorTitle);
  htmlOutput.setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
Logger.log( "returning htmlOutput, htmlOutput.getContent() = " + htmlOutput.getContent() );
  return htmlOutput;
}

//returns true if logged in user is wlc
function modalLoggedInUsersCheck(  ) {
  let spreadsheet = SpreadsheetApp.getActive();

  //get email validation regex
  let settingsSheet = spreadsheet.getSheetByName( SHEETS.GLOBAL_SETTINGS.NAME );
  let emailValidationRegexString = settingsSheet.getRange( 
        SHEETS.GLOBAL_SETTINGS.REFS.ROW_NO_VALID_EMAIL_REGEX, 
        SHEETS.GLOBAL_SETTINGS.REFS.COL_NO
      ).getValue();

  let emailValidationRegex = new RegExp( emailValidationRegexString );

  return emailValidationRegex.test( Session.getActiveUser().getEmail() ) && 
         emailValidationRegex.test( Session.getEffectiveUser().getEmail() );
}