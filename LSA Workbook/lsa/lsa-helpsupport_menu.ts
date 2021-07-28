function ShowEmailSupportPopup() {
  var spreadsheet = SpreadsheetApp.getActive();
  var ui = SpreadsheetApp.getUi();
  var settingsSheet = spreadsheet.getSheetByName( SHEETS.GLOBAL_SETTINGS.NAME );

  //get mailto settings from global settings
  var mailToEmailTo = settingsSheet.getRange(
        SHEETS.GLOBAL_SETTINGS.REFS.ROW_NO_SUPPORT_EMAIL_TO, 
        SHEETS.GLOBAL_SETTINGS.REFS.COL_NO 
      ).getValue();
  var mailToEmailCC = settingsSheet.getRange(
        SHEETS.GLOBAL_SETTINGS.REFS.ROW_NO_SUPPORT_EMAIL_CCS, 
        SHEETS.GLOBAL_SETTINGS.REFS.COL_NO 
      ).getValue();
  var mailToBodyEncoded = settingsSheet.getRange(
        SHEETS.GLOBAL_SETTINGS.REFS.ROW_NO_SUPPORT_EMAIL_BODY, 
        SHEETS.GLOBAL_SETTINGS.REFS.COL_NO 
      ).getValue();

    var alertHTMLTemplate = HtmlService.createTemplateFromFile('html/html-modal-email-support');
    alertHTMLTemplate.mailtoURL = "https://mail.google.com/mail?view=cm&tf=0&to=" + mailToEmailTo + "&cc=" + mailToEmailCC + 
                                  "&su=I%20have%20some%20feedback%20about%20my%20RoS%20Generator%20Spreadsheet&body="+mailToBodyEncoded;

    alertHTMLTemplate.supportEmailAddress = mailToEmailTo;
    var alertMessageHTML = alertHTMLTemplate.evaluate().getContent();
    var alertMessage =  HtmlService.createHtmlOutput( alertMessageHTML ).setWidth(600).setHeight(280);
    ui.showModalDialog(alertMessage, 'Talk to us...');
}

function ShowGetHelpPopup() {
  var spreadsheet = SpreadsheetApp.getActive();
  var ui = SpreadsheetApp.getUi();
  var helpLinksSheet = spreadsheet.getSheetByName( SHEETS.MASTER_HELP.NAME );

  //work out which sheet the user is on
  var currentSheetName = spreadsheet.getActiveSheet().getName();

  //get array of links to list in the modal popup
  let allLinksData = helpLinksSheet.getRange( SHEETS.MASTER_HELP.REFS.ROW_NO_FIRST_RECORD, SHEETS.MASTER_HELP.REFS.COL_NO_LINK_TEXT,
                                     SHEETS.MASTER_HELP.REFS.ROW_NO_LAST_RECORD - SHEETS.MASTER_HELP.REFS.ROW_NO_FIRST_RECORD + 1, 
                                     SHEETS.MASTER_HELP.REFS.COL_NO_LINK_URL ).getValues();
  
  //loop through each row of link data, HTML conatining of a series of list items (eg <li>) as a string
  let thisLinksSheet      = "";
  let thisLinksURL        = "";
  let thisLinksText       = "";
  let thisLinkHTML        = "";
  let helpCentreLinkHTML  = "";
  let sheetOnlyLinksHTML  = "";
  let otherCommonLinksHTML= "";
  let colOffset = SHEETS.MASTER_HELP.REFS.COL_NO_LINK_TEXT;
  for( let i=0; i < allLinksData.length; i++ ) {
    thisLinksSheet  = allLinksData[i][SHEETS.MASTER_HELP.REFS.COL_NO_WHICH_SHEET  - colOffset];
    thisLinksURL    = allLinksData[i][SHEETS.MASTER_HELP.REFS.COL_NO_LINK_URL     - colOffset];
    thisLinksText   = allLinksData[i][SHEETS.MASTER_HELP.REFS.COL_NO_LINK_TEXT    - colOffset];

Logger.log( "thisLinksSheet = " + thisLinksSheet );
Logger.log( "thisLinksURL = " + thisLinksURL );
Logger.log( "thisLinksText = " + thisLinksText );

    if( thisLinksURL != "" && thisLinksText != "" ) {
      thisLinkHTML = "<li><a target=\"_blank\" href=\"" + thisLinksURL + "\">" + thisLinksText + "</a></li>\n";

      Logger.log( "thisLinkHTML = " + thisLinkHTML );

      if( thisLinksSheet == "[LINK_TO_HELP_CENTRE]" ) {
        helpCentreLinkHTML += thisLinkHTML;
      }
      else if( thisLinksSheet != "" && thisLinksSheet == currentSheetName ) {
        sheetOnlyLinksHTML += thisLinkHTML;
        Logger.log( "sheetOnlyLinksHTML = " + sheetOnlyLinksHTML );
      }
      else if ( thisLinksSheet == "" ) {
        otherCommonLinksHTML += thisLinkHTML;
        Logger.log( "otherCommonLinksHTML = " + otherCommonLinksHTML );
      }
    }
  }

  //get the modal's html template from file and convert it to an HTML String
  var alertHTMLTemplate = HtmlService.createTemplateFromFile('html/html-modal-help-centre');
  alertHTMLTemplate.sheetOnlyLinksHTML    = sheetOnlyLinksHTML;
  alertHTMLTemplate.otherCommonLinksHTML  = otherCommonLinksHTML;
  alertHTMLTemplate.helpCentreLinkHTML    = helpCentreLinkHTML;
  var alertMessageHTML = alertHTMLTemplate.evaluate().getContent();
  var alertMessage =  HtmlService.createHtmlOutput( alertMessageHTML ).setWidth(600).setHeight(360);
  ui.showModalDialog(alertMessage, 'What can we help you with?');
}
