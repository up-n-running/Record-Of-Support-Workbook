function ExportAllRecordsToLocalPDF() {
  let spreadsheet = SpreadsheetApp.getActive();
  let ui = SpreadsheetApp.getUi();
  var inputSheet = spreadsheet.setActiveSheet(spreadsheet.getSheetByName(SHEETS.INPUT.NAME), true);

  //get info from sheet in order to generate filename later
  var recordDate = inputSheet.getRange( 
        SHEETS.INPUT.REFS.LESSON_DATE.ROW_NO,
        SHEETS.INPUT.REFS.LESSON_DATE.COL_NO,
      ).getValue();
  var recordLSAName = inputSheet.getRange( 
        SHEETS.INPUT.REFS.LSA_NAME.ROW_NO,
        SHEETS.INPUT.REFS.LSA_NAME.COL_NO 
      ).getValue();

  
  //check validation and generate a list of records 1 thru 25 with validation errors
  //get array of values for the valudation results, 1 for each reord of support from 1 thru 25
  var toPrintRange = inputSheet.getRange( 3, 2, 1, 25 );
  var toPrintRangeValues = toPrintRange.getValues();
  var noOfRecordsFailingValidation = 0;
  var listOfRecordNumbersFilingValidation = "";
  var noOfRecordsToPrint = 0;
  for ( let i = 0; i < 25; i++) {
    if( toPrintRangeValues[0][i] === "To Send" || toPrintRangeValues[0][i] === "Sent to Learner" || toPrintRangeValues[0][i] === "Signed" ) {
      noOfRecordsToPrint++;
    } else if( toPrintRangeValues[0][i] !== "" ){
      listOfRecordNumbersFilingValidation += ((noOfRecordsFailingValidation===0)?"":", ") + (i+1);
      noOfRecordsFailingValidation++;
    } 
  }
  
  //only go ahead if there are no records failing validation
  if( noOfRecordsFailingValidation > 0 ) {
    ui.alert( ""+noOfRecordsFailingValidation+" " + ((noOfRecordsFailingValidation===1)?"Record Has":"Records Have") + " Errors", 
              "Sorry, the following record " + ((noOfRecordsFailingValidation===1)?"number has":"numbers have") + " errors:\n" 
            + listOfRecordNumbersFilingValidation 
            + "\n\nPlease correct these errors and try again", 
              ui.ButtonSet.OK );
  } //and if there is at least one record to print
  else if( noOfRecordsToPrint == 0 ) {
    ui.alert( "You Have No Records To Print", 
              "Please complete at least one record of support before attempting to print.", 
              ui.ButtonSet.OK );
  } //and if the date field has been populated
  else if ( validateInputSheetFields( inputSheet, ui, null ) ){
    //are you sure dialogue?
    var areYouSureResponse = ui.alert('Are you sure?', 'This will generate a printable pdf file in your google drive\nand will take a while to run - please be patient.\n\nAre you sure?', ui.ButtonSet.OK_CANCEL)
    
    if( areYouSureResponse == ui.Button.OK )
    {
      spreadsheet.toast('Hiding the sheets that we dont want to print, please be patient', 'Step 1 of 2');
      
      spreadsheet.getSheetByName('Settings - Learners').hideSheet();
      spreadsheet.getSheetByName('Settings - Lessons').hideSheet();
      spreadsheet.getSheetByName('Settings - Target Grades').hideSheet();
      spreadsheet.getSheetByName('Settings - Lesson Targets').hideSheet();
      spreadsheet.getSheetByName( SHEETS.GLOBAL_SETTINGS.NAME ).hideSheet();
      spreadsheet.getSheetByName('Master').hideSheet();
      
      //loop through record of support tabs 1 to 25 hiding and unhiding as appropriate
      for ( let i = 0; i < 25; i++) {
        if( toPrintRangeValues[0][i] === "To Send" || toPrintRangeValues[0][i] === "Sent to Learner" || toPrintRangeValues[0][i] === "Signed" ) {
          if( spreadsheet.getSheetByName(""+(i+1)).isSheetHidden() ) {
            spreadsheet.getSheetByName(""+(i+1)).showSheet();
          }
        } else {
          if( !spreadsheet.getSheetByName(""+(i+1)).isSheetHidden() ) {
            spreadsheet.getSheetByName(""+(i+1)).hideSheet();
          }
        } 
      }
      
      //Export to printable PDF
      //Hide the 'Input worksheet aswe dont want to print it
      SpreadsheetApp.getActiveSpreadsheet().toast('Generating Printable PDF file, Please be patient', 'Step 2 of 2');
      inputSheet.hideSheet();
      var pdfFile = savePDFs( null, null, recordLSAName + "_EXPORT_" + Utilities.formatDate(recordDate, GLOBAL_CONSTANTS.TIMEZONE, "yyyy-MM-dd" ), ui );
      
      //Unhide the 'Input' worksheet as thats the master control worksheet
      inputSheet.showSheet();
      spreadsheet.setActiveSheet( inputSheet );
      
      //Show success dialogue
      //ui.alert( 'PDF File Created', '.PDF file generated successfully ready for printing.\n\nCheck this workbook\'s folder in your google drive.', ui.ButtonSet.OK );
      const styleHTMLInsert = ' style="font-weight:400;font-size:16px; font-family: Calibri, Candara, Segoe, Optima, Arial, sans-serif;"';
      const htmlOutput = HtmlService.createHtmlOutput('<p'+styleHTMLInsert+'>PDF file generated successfully ready for printing<br />Click to open</p>'
                                                      + '<ul>'
                                                      + '<li '+styleHTMLInsert+'><a href="' + pdfFile.getUrl() + '" target="_blank">' + pdfFile.getName() + '</a></li>'
                                                      + '</ul>').setWidth(500).setHeight(130);
      ui.showModalDialog(htmlOutput, 'Export Successful');
    }
  }
}

function savePDFs( optSSId, optSheetId, fileName, ui ) {

  //FOR EXTRA HELP WITH THIS SEE
  //https://xfanatical.com/blog/print-google-sheet-as-pdf-using-apps-script/
  
  // If a sheet ID was provided, open that sheet, otherwise assume script is
  // sheet-bound, and open the active spreadsheet.
  var spreadSheet = (optSSId) ? SpreadsheetApp.openById(optSSId) : SpreadsheetApp.getActiveSpreadsheet();

  // Get folder containing spreadsheet, for later export
  var parents = DriveApp.getFileById(spreadSheet.getId()).getParents();
  if (parents.hasNext()) {
    var folder = parents.next();
  }
  else {
    folder = DriveApp.getRootFolder();
  }
  
  var blob = generatePDFBlob( spreadSheet, optSheetId, ui );
  blob.setName(fileName + '.pdf');
  
  //from here you should be able to use and manipulate the blob to send and email or create a file per usual.
  //In this example, I save the pdf to drive
  var pdfFile = folder.createFile(blob);
  
  return pdfFile;
}
