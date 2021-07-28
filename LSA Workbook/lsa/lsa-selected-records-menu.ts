function SendSelectedRecords() {
  var spreadsheet = SpreadsheetApp.getActive();
  var inputSheet = spreadsheet.getSheetByName( SHEETS.INPUT.NAME );
  var ui = SpreadsheetApp.getUi();
  var selectedRecordNos = getSelectedRecordNos( spreadsheet, ui );

  if( selectedRecordNos.length > 0 )
  {
    var alertText = "The following records are selected:\n\n";
    var noOfRecordsToBeSent  =0;
    for (var i=0; i<selectedRecordNos.length; i++)
    {
      var recordNo = selectedRecordNos[i];
      var recordColumnNo = SHEETS.INPUT.REFS.COL_NO_RECORD_1 + recordNo - 1;
      var recordStatus = inputSheet.getRange( SHEETS.INPUT.REFS.ROW_NO_STATUS_BAR, recordColumnNo ).getValue();
      if( recordStatus === SHEETS.INPUT.STATUSES.SIGNED || recordStatus === SHEETS.INPUT.STATUSES.SIGNED_AUTOSIGN ||
          recordStatus === SHEETS.INPUT.STATUSES.UNSIGNED ) {
        alertText += "  (" + recordNo + ") No Action Required (aready sent)\n";
      }
      else if ( recordStatus === SHEETS.INPUT.STATUSES.UNSENT || recordStatus === SHEETS.INPUT.STATUSES.UNSENT_AUTOSIGN ||
                recordStatus === SHEETS.INPUT.STATUSES.SAVED_EMAILWAIT ) {
        alertText += "  (" + recordNo + ") Record will be sent to: "+
                     inputSheet.getRange( SHEETS.INPUT.REFS.ROW_NO_EMAIL_ADDRESS, recordColumnNo ).getValue()+"\n";
        noOfRecordsToBeSent++;
      }
      else {
        alertText += "  (" + recordNo + ") No Action Required\n";
      }
    }
    alertText += "\nIf you continue, a total of " + noOfRecordsToBeSent + 
                 " Record"+(noOfRecordsToBeSent==1?"":"s") + " will be sent.\n";
    alertText += "\nAre you sure you want to continue?";

    var areYouSureResponse = ui.alert('Please Review Before Sending', alertText, ui.ButtonSet.OK_CANCEL);
      
    if( areYouSureResponse == ui.Button.OK )
    {  
      ExportToSendRoSsFromList(selectedRecordNos, spreadsheet, inputSheet, null, ui );
    }
  }
}


function FlagSelectedRecordsAsAutoSign() {
  let spreadsheet = SpreadsheetApp.getActive();
  let inputSheet = spreadsheet.getSheetByName( SHEETS.INPUT.NAME );
  let ui = SpreadsheetApp.getUi();
  let selectedRecordNos = getSelectedRecordNos( spreadsheet, ui );
  let recordNosToFlag = new Array();

  if( selectedRecordNos.length > 0 )
  {
    let alertText = "The following records are selected:\n\n";
    let noOfRecordsToBeSent  =0;
    let noOfTooLateRecords = 0;
    for (var i=0; i<selectedRecordNos.length; i++)
    {
      let recordNo = selectedRecordNos[i];
      let recordColumnNo = SHEETS.INPUT.REFS.COL_NO_RECORD_1 + recordNo - 1;
      let recordStatus = inputSheet.getRange( SHEETS.INPUT.REFS.ROW_NO_STATUS_BAR, recordColumnNo ).getValue();
      
      if( recordStatus === SHEETS.INPUT.STATUSES.SIGNED || recordStatus === SHEETS.INPUT.STATUSES.SIGNED_AUTOSIGN ) {
        alertText += "  (" + recordNo + ") No Action Required (aready signed)\n";
      }
      else if ( recordStatus === SHEETS.INPUT.STATUSES.SAVED_EMAILWAIT || recordStatus === SHEETS.INPUT.STATUSES.UNSIGNED ) {
        alertText += "  (" + recordNo + ") No Action Possible (already generated)\n";
        noOfTooLateRecords++;
      }
      else if ( recordStatus === SHEETS.INPUT.STATUSES.UNSENT || recordStatus === SHEETS.INPUT.STATUSES.UNSENT_AUTOSIGN ) {
        alertText += "  (" + recordNo + ") Record will be flagged to Auto-Sign\n";
        recordNosToFlag.push( recordNo );
        noOfRecordsToBeSent++;
      }
      else {
        alertText += "  (" + recordNo + ") No Action Required\n";
      }
    }
    alertText += "\nIf you continue, a total of " + noOfRecordsToBeSent + 
                 " Record"+(noOfRecordsToBeSent==1?"":"s") + " will be flagged to Auto-Sign.\n";
    if( noOfTooLateRecords > 0 ) {
      alertText += "\nFor the  " + noOfTooLateRecords + " Record"+(noOfTooLateRecords==1?"":"s") + " that ha"+
            (noOfTooLateRecords==1?"s":"ve")+" already been generated, you could try deleting " + 
            (noOfTooLateRecords==1?"this file":"these files") + " before trying again.\n";
    }
    alertText += "\nAre you sure you want to continue?";

    var areYouSureResponse = ui.alert('Please Review Before Sending', alertText, ui.ButtonSet.OK_CANCEL);
      
    if( areYouSureResponse == ui.Button.OK )
    {  
      let autoSignReason = "";
      let promptResponse = null;
      let cancelled = false;
      while( !cancelled && autoSignReason == "" ) {
        promptResponse = ui.prompt( "Why is a Signature Not Required?", 
            "Signatures are still required from the learner if they are absent, with the exception of Authorised Absences.\n" +
            "Authorised Absences are, however, only applicable to class trips and work expereince.\n\n"+
            "Any kind of sickness or medical appointment (even pre-arranged doctors appointments) are still classed as " + 
            "Unauthorised Absence, and as such, sigantures are still required.\n\n" +
            "You must enter a valid reason for not requesting a signature, which will later be checked and may be formally audited " + 
            "by the External Agency who provide funding for the learner's support.\n\n" +
            "Please enter the reason for not requesting a signature below.\n\n" + 
            "The reason you enter here will appear on the Record"+(noOfRecordsToBeSent==1?"":"s")+" of Support.\n\n", 
            ui.ButtonSet.OK_CANCEL );
        autoSignReason = promptResponse.getResponseText();
        cancelled = ( promptResponse.getSelectedButton() == ui.Button.CANCEL );
      }
      if( !cancelled ) {
Logger.log(recordNosToFlag);
        let noOfSuccesses = 0;
        for( let i=0; i<recordNosToFlag.length; i++ ) {
          let recordNo = recordNosToFlag[i];
          let recordColumnNo = SHEETS.INPUT.REFS.COL_NO_RECORD_1 + recordNo - 1;
          inputSheet.getRange( SHEETS.INPUT.REFS.ROW_NO_AUTOSIGN_MANUALENTRY, recordColumnNo ).setValue( autoSignReason );
          noOfSuccesses++;
        }
        //FEEDBACK ONCE RECORDS HAVE BEEN FLAGGED TO AUTOSIGN
        ui.alert(
          'Finished Flagging Records to Auto-Sign', 
          'A total of ' + noOfSuccesses + ' record' + 
          (noOfSuccesses==1 ? ' was' : 's were' ) + ' flagged to Auto-Sign.' + 
          "\n\nPlease be aware that if you change the 'Lesson', 'Learner' or 'Attended' fields on " + 
          (noOfSuccesses==1 ? ' this Record' : 'any of these Records' ) +
          ", then the flag will be removed.",
          ui.ButtonSet.OK);
      }
    }
  }
}


function ReminderEmailForSelectedRecords() {
  var spreadsheet = SpreadsheetApp.getActive();
  var ui = SpreadsheetApp.getUi();

  ui.alert( "Coming in a later version", "We are sorry, this feature is not available yet in this version", ui.ButtonSet.OK );
  return;

  var selectedRecordNos = getSelectedRecordNos( spreadsheet, ui );

  if( selectedRecordNos.length > 0 )
  {
    var alertText = "The following records are selected:\n\n";
    var noOfEMailsToSend = 0;
    var noOfSignedFiles = 0;
    for (var i=0; i<selectedRecordNos.length; i++)
    {
      var recordNo = selectedRecordNos[i];
      var recordColumnLetter = columnToLetter( recordNo + 1 );
      var recordStatus = spreadsheet.getRange( "" + recordColumnLetter + "3" ).getValue();
      if( recordStatus === "Signed" ) {
        alertText += "  (" + recordNo + ") ****** EMAIL WILL BE SENT FOR ALREADY SIGNED PDF!! *****\n";
        noOfEMailsToSend++; noOfSignedFiles++;
      }
      else if ( recordStatus === "Sent to Learner" ) {
        var recordEmailAddress = spreadsheet.getRange( "" + recordColumnLetter + "24" ).getValue();
        alertText += "  (" + recordNo + ") Email will be sent to: " + recordEmailAddress + "\n";
        noOfEMailsToSend++;
      }
      else {
        alertText += "  (" + recordNo + ") No Action Required\n";
      }
    }
    alertText += "\nIf you continue, a total of " + noOfEMailsToSend + " email"+(noOfEMailsToSend==1?"":"s") + " will be sent.\n";
    alertText += ( (noOfSignedFiles > 0) 
                    ? "INCLUDING " + noOfSignedFiles + " EMAIL"+ 
                      (noOfSignedFiles==1?" FOR A RECORD THAT HAS":"S FOR RECORDS THAT HAVE")  + 
                      " ALREADY BEED SIGNED.\n" 
                    : "" 
                  );
    alertText += "\nSend these reminder emails now?";

    var areYouSureResponse = ui.alert('Please Review Before Emailing', alertText, ui.ButtonSet.OK_CANCEL);
      
    if( areYouSureResponse == ui.Button.OK )
    {
      var totalEmailsSent = 0;
      var emailTemplate = null;
      for (var i=0; i<selectedRecordNos.length; i++)
      {
        var skipThisFile = true;
        var recordNo = selectedRecordNos[i];
        var recordColumnLetter = columnToLetter( recordNo + 1 );
        var recordStatus = spreadsheet.getRange( "" + recordColumnLetter + "3" ).getValue();

        //work out if we are actually going to delete this file
        if( recordStatus == "Signed" ) {
          var areYouSureResponse = ui.alert(
                          "Consider Skipping this SIGNED Record", 
                          "The PDF for record " + recordNo + " has already been signed by the learner\n\n" + 
                          "If you press 'no' this reminder email will be skipped\n\n" +
                          "Are you sure you want to send a reminder email?",
                          ui.ButtonSet.YES_NO);
          if( areYouSureResponse == ui.Button.YES ) {
            skipThisFile = false;
          }
        }
        else if ( recordStatus === "Sent to Learner" ) {
          skipThisFile = false;
        }

        //fetch file ready for emailing if we are not skipping
        if( !skipThisFile ) {
          var recordEmailAddress = spreadsheet.getRange( "" + recordColumnLetter + "24" ).getValue();
          var recordLearnerName  = spreadsheet.getRange( "" + recordColumnLetter + "7" ).getValue();
          var recordLSAName      = spreadsheet.getRange( "E1" ).getValue();
          var fileId = spreadsheet.getRange( "" + recordColumnLetter + "25" ).getValue();
          var pdfFile;
          try {
            pdfFile = DriveApp.getFileById(fileId);
            Logger.log( pdfFile );
          }
          catch( e )
          {
            debugCatchError( e );
            ui.alert(
              "Error Accessing File", 
              "The PDF for Record '" + recordNo + "' with file id '" + fileId + 
              "' could not be accessed. Perhaps it has already been manually deleted or the file id is wrong.\n\n" + 
              "Error Details:\n\n" + e,
              ui.ButtonSet.OK );
          }
          
          SendLearnerSignEmail( spreadsheet, pdfFile, recordNo, recordEmailAddress, recordLearnerName, null, recordLSAName,
                null, null, null, true, null, emailTemplate );

          totalEmailsSent++;
        }
      }

      SpreadsheetApp.flush();
      //FEEDBACK ONCE REMINDER EMAILS HAVE BEEN SENT
      ui.alert(
        'Finished Sending Reminder Emails', 
        'A total of ' + totalEmailsSent + ' reminder email' + 
        (totalEmailsSent==1 ? ' was' : 's were' ) + ' sent',
        ui.ButtonSet.OK);
    }
  }
}


function DeletePDFFromSelectedRecords() {
  let spreadsheet = SpreadsheetApp.getActive();
  let inputSheet = spreadsheet.getSheetByName( SHEETS.INPUT.NAME );
  let ui = SpreadsheetApp.getUi();
  let selectedRecordNos = getSelectedRecordNos( spreadsheet, ui );

  if( selectedRecordNos.length > 0 )
  {
    let alertText = "The following records are selected:\n\n";
    let noOfFilesToDelete = 0;
    let noOfSignedFiles = 0;
    for (var i=0; i<selectedRecordNos.length; i++)
    {
      var recordNo = selectedRecordNos[i];
      var recordColumnNo = SHEETS.INPUT.REFS.COL_NO_RECORD_1 + recordNo - 1;;
      var recordStatus = inputSheet.getRange( SHEETS.INPUT.REFS.ROW_NO_STATUS_BAR, recordColumnNo ).getValue();
      if( recordStatus === SHEETS.INPUT.STATUSES.SIGNED ) {
        alertText += "  (" + recordNo + ") ****** RoS HAS ALREADY BEEN SIGNED *****\n";
        noOfFilesToDelete++; noOfSignedFiles++;
      }
      else if ( recordStatus == SHEETS.INPUT.STATUSES.UNSIGNED || recordStatus == SHEETS.INPUT.STATUSES.SIGNED_AUTOSIGN ){
        alertText += "  (" + recordNo + ") Deletion will be attempted\n";
        noOfFilesToDelete++;
      }
      else {
        alertText += "  (" + recordNo + ") No Action Required\n";
      }
    }
    alertText += "\nIf you continue, deletion will be attempted for a total of " + noOfFilesToDelete + 
                 " file"+(noOfFilesToDelete==1?"":"s") + ".\n";
    alertText += ( (noOfSignedFiles > 0) 
                    ? "INCLUDING " + noOfSignedFiles + " FILE"+ 
                      (noOfSignedFiles==1?" THAT HAS":"S THAT HAVE")  + 
                      " ALREADY BEED SIGNED.\n" 
                    : "" 
                 );
    alertText += "\nAre you absolutely sure you want to continue?";

    let areYouSureResponse = ui.alert('Please Review Before Deleting', alertText, ui.ButtonSet.OK_CANCEL);
      
    if( areYouSureResponse == ui.Button.OK )
    {
      let totalFilesDeleted = 0;
      for (var i=0; i<selectedRecordNos.length; i++)
      {
        let skipThisFile = true;
        let recordNo = selectedRecordNos[i];
        let recordColumnNo = SHEETS.INPUT.REFS.COL_NO_RECORD_1 + recordNo - 1;
        let recordStatus = inputSheet.getRange( SHEETS.INPUT.REFS.ROW_NO_STATUS_BAR, recordColumnNo ).getValue();

        //work out if we are actually going to delete this file
        if( recordStatus === SHEETS.INPUT.STATUSES.SIGNED ) {
          areYouSureResponse = ui.alert(
                          "Consider Skipping this SIGNED PDF", 
                          "The PDF for record " + recordNo + " has already been signed by the learner\n\n" + 
                          "If you press 'no' this file will be skipped and not deleted\n\n" +
                          "Are you sure you want to delete this file?",
                          ui.ButtonSet.YES_NO);
          if( areYouSureResponse == ui.Button.YES ) {
            skipThisFile = false;
          }
        }
        else if ( recordStatus == SHEETS.INPUT.STATUSES.UNSIGNED || recordStatus == SHEETS.INPUT.STATUSES.SIGNED_AUTOSIGN ) {
          skipThisFile = false;
        }

        //delete if we are not skipping
        if( !skipThisFile ) {
          inputSheet.getRange( SHEETS.INPUT.REFS.ROW_NO_STATUS_BAR, recordColumnNo );
          let fileId = inputSheet.getRange( SHEETS.INPUT.REFS.ROW_NO_FILE_ID, recordColumnNo ).getValue();

          spreadsheet.toast('Attempting to send Record ' + recordNo + ' to trash can.', 'Deleting, Record: ' + recordNo);
          let errorMessage = "Webapp Returned null - please notify Support to ask them to check the webapp logs";
          let webAppReturnData = CallMasterRoSWebapp( null, "delete-ros", spreadsheet.getId(), null, null, fileId, null, null,
                null, null, null, null, null );
          if( webAppReturnData != null ) {
            if( webAppReturnData.success != 1.0 )
            {
              errorMessage = webAppReturnData.errorMessage;
            }
            else {
              errorMessage = null;
              UpdateFileInfoInWorkbook(spreadsheet, inputSheet, null, recordNo, DriveApp.getFileById( fileId ), true );
              totalFilesDeleted++;
            }
          }
          
          Logger.log( errorMessage );

          if( errorMessage != null ) {
            ui.alert(
              "Error Deleting File", 
              "The PDF for Record '" + recordNo + "' with file id '" + fileId + 
              "' could not be deleted\n\n" + 
              "Error Details:\n\n" + errorMessage,
              ui.ButtonSet.OK );
          }
        }
      }

      //FEEDBACK ONCE FILES HAVE BEEN DELETED
      ui.alert(
        'Finished Deleting Files', 
        'A total of ' + totalFilesDeleted + ' file' + 
        (totalFilesDeleted==1 ? ' was' : 's were' ) + ' deleted' + 
        ( (totalFilesDeleted>0) ? '\n\nIf the learner has already received an email for any of these files, then please ensure' +
          ' that they are aware that their links to these deleted files will no longer work' : '' ),
        ui.ButtonSet.OK);
    }
  }
}
