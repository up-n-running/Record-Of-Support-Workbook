function ShowTargetGradesModal() {
  //let spreadsheet = SpreadsheetApp.getActive();
  let ui = SpreadsheetApp.getUi();

  //generate HTML success alert with LSA Directory hyperlink embedded
  let modalHTMLTemplate = HtmlService.createTemplateFromFile('html/html-modal-targets-help');
  let alertMessageHTML = modalHTMLTemplate.evaluate().getContent();
  let alertMessage =  HtmlService.createHtmlOutput( alertMessageHTML ).setWidth(700).setHeight(500);
  ui.showModalDialog( alertMessage,"Target Text Builder" );
}
