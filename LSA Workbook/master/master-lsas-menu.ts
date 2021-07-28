function CreateNewChildWorkbookForSelectedLSAs() {

  var spreadsheet = SpreadsheetApp.getActive();
  var ui = SpreadsheetApp.getUi();
  var globalSettingsSheet = spreadsheet.getSheetByName( SHEETS.GLOBAL_SETTINGS.NAME );

  var selectedLSANos = getSelectedLSANos( spreadsheet, ui );
  //the above function ensures we're on the lsa worksheet
  var lsaWorksheet = spreadsheet.getActiveSheet();

  if( selectedLSANos.length > 0 )
  {
    //get the info about the google group that the LSA's email record must belong to in order to pass validation
    //get the google group All LSAs ALL Campuses
    let allLSAsGroupEmail = globalSettingsSheet.getRange( 
          SHEETS.GLOBAL_SETTINGS.REFS.ROW_NO_LSA_GROUP_EMAIL, 
          SHEETS.GLOBAL_SETTINGS.REFS.COL_NO 
        ).getValue();
    let allLSAsGroupAdminURL = globalSettingsSheet.getRange( 
          SHEETS.GLOBAL_SETTINGS.REFS.ROW_NO_LSA_GROUP_ADMIN_URL, 
          SHEETS.GLOBAL_SETTINGS.REFS.COL_NO 
        ).getValue();
    let lsaRootDirectoryId = globalSettingsSheet.getRange( 
          SHEETS.GLOBAL_SETTINGS.REFS.ROW_NO_DIR_ID_LSAS, 
          SHEETS.GLOBAL_SETTINGS.REFS.COL_NO 
        ).getValue();
    var allLSAsGroup = GroupsApp.getGroupByEmail( allLSAsGroupEmail );
    

    var alertText = "The following LSAs are selected:\n\n";
    var noOfWorksheetsToBeCreated = 0;
    var recordNo = -1, rowNo = -1, rowStatus = null;
    for (var i=0; i<selectedLSANos.length; i++)
    {
      recordNo = selectedLSANos[i];
      rowNo = recordNo + SHEETS.MASTER_LSAS.REFS.ROW_NO_FIRST_LSA - 1;
      rowStatus = lsaWorksheet.getRange( rowNo, SHEETS.MASTER_LSAS.REFS.COL_NO_STATUS_BAR ).getValue();
      if( rowStatus == "Awaiting Creation" ) {
        noOfWorksheetsToBeCreated++;
        alertText += "  (" + recordNo + ") Workbook will be created for: " +  
                     lsaWorksheet.getRange( rowNo, SHEETS.MASTER_LSAS.REFS.COL_NO_LSA_NAME ).getValue() + "\n";

      }
      else if ( rowStatus == "Awaiting Update" || rowStatus == "On Latest Version"){
        alertText += "  (" + recordNo + ") No Action Required (aready created)\n";
      }
      else {
        alertText += "  (" + recordNo + ") No Action Required\n";
      }

    }
    alertText += "\nIf you continue, a total of " + noOfWorksheetsToBeCreated + 
                 " Worksheet"+(noOfWorksheetsToBeCreated==1?"":"s") + " will be created.\n";
    alertText += "\nAre you sure you want to create these workbooks?";

    var areYouSureResponse = ui.alert('Please Review Before Creating Workbooks', alertText, ui.ButtonSet.OK_CANCEL);
      
    if( areYouSureResponse == ui.Button.OK )
    {
      let noOfWokrsheetsCreated = 0, lsaEmailFromMaster: string = null, lsaNameFromMaster: string = null;
      let webAppReturnData = null, skip = false;

      for (var i=0; !skip && i<selectedLSANos.length; i++) {
        recordNo = selectedLSANos[i];
        rowNo = recordNo + SHEETS.MASTER_LSAS.REFS.ROW_NO_FIRST_LSA - 1;
        rowStatus = lsaWorksheet.getRange( rowNo, SHEETS.MASTER_LSAS.REFS.COL_NO_STATUS_BAR ).getValue();

        if( rowStatus === "Awaiting Creation" ) {
          lsaEmailFromMaster = lsaWorksheet.getRange(
            rowNo,
            SHEETS.MASTER_LSAS.REFS.COL_NO_LSA_EMAIL
          ).getValue().trim();
          lsaNameFromMaster = lsaWorksheet.getRange(
            rowNo,
            SHEETS.MASTER_LSAS.REFS.COL_NO_LSA_NAME
          ).getValue();      

          spreadsheet.toast( 'Creating new workbook for ' + lsaEmailFromMaster + '. Please be patient, this can take a while.',
                             'Creating new LSA Workbook', 20 );

          //double check the email address is part of the All LSAs All Campuses group otherwise we wont allow creation of sheet
          //If they are not a memeber of that group they will not be able to use the update Webapp, amongst other things.
          if( !allLSAsGroup.hasUser(lsaEmailFromMaster) ) {
            webAppReturnData = {
              success: 0,
              errorMessage: "The email address '"+lsaEmailFromMaster+
              "' is not part of the google group: '" + allLSAsGroupEmail + "'\n\nIf you're sure you have entered their email " +
              "address correctly then please copy and paste this URL into a new tab:\n" + allLSAsGroupAdminURL + "\n\n" +
              "This page will allow you to add members to the google group, then you can retry creating this LSAs Workbook.",
              affectedFileId: null
            };
          }
          else {
            webAppReturnData = CallMasterRoSWebapp( globalSettingsSheet, "create", null, lsaEmailFromMaster, spreadsheet.getId(), null, 
                    null, null, null, null, null, null, null );
          }

          if( webAppReturnData.success != 1.0 ) {
            var response = ui.alert( "Could not Create Workbook", 
                                  "Creating workbook failed for LSA with email: " + lsaEmailFromMaster + ".\n" +
                                  "Please check for a part created workbook in the LSAs folder and delete any before trying again.\n\n" +
                                  "Error Message:\n" +
                                  webAppReturnData.errorMessage + "\n\n" +
                                  "Do you want to continue on to try the remaining LSAs?", 
                                  ui.ButtonSet.OK_CANCEL );
            skip = response == ui.Button.CANCEL;
          }
          else {
            //send welcome email then increment the tally of worksheets created
            try {
              var lsasFolder = getLSAsFolderFromWorkbookFile( DriveApp.getFileById( webAppReturnData.affectedFileId ), 
                                                              lsaNameFromMaster, lsaRootDirectoryId );
              SendLSAWelcomeEMail( lsaEmailFromMaster, lsaNameFromMaster, lsasFolder.getUrl(), spreadsheet, null );
            }
            catch( e ) {
              debugCatchError( e );
              ui.alert( "Could not send welcome email", 
                        "Creating workbook SUCCEEDED for LSA with email: " + lsaEmailFromMaster + ".\n" +
                        "But their welcome email failed to send, please email them manually\n\n" +
                        "Error Message:\n" +
                        catchErrorToString( e ) + "\n\n" +
                        "The system will now continue to the next LSAs Workbook", 
                        ui.ButtonSet.OK );
            }
            noOfWokrsheetsCreated ++;
          }
        }
      }

      //FEEDBACK ONCE FILES HAVE BEEN SENT
      ui.alert(
        'Finished Creating LSA Worksheets', 
        'A total of ' + noOfWokrsheetsCreated + ' LSA Worksheet' + 
        (noOfWokrsheetsCreated==1 ? ' was' : 's were' ) + 
        ' created successfully.', 
        ui.ButtonSet.OK);
    }
  }
}


function OpenLSAsDirectory() {
  let spreadsheet = SpreadsheetApp.getActive();
  let ui = SpreadsheetApp.getUi();

  let selectedLSANos = getSelectedLSANos( spreadsheet, ui );
  //the above function ensures we're on the lsa worksheet
  let lsaWorksheet = spreadsheet.getActiveSheet();

  if( selectedLSANos.length > 1 ) {
    ui.alert( "Please select one LSA record", "Please select just one LSA record to use this function.", ui.ButtonSet.OK );
  }                        
  else if (selectedLSANos.length == 1 ) {
    var rowNo = selectedLSANos[0] + SHEETS.MASTER_LSAS.REFS.ROW_NO_FIRST_LSA - 1;
    let rowStatus: string = lsaWorksheet.getRange( rowNo, SHEETS.MASTER_LSAS.REFS.COL_NO_STATUS_BAR ).getValue();
    if( rowStatus != "Awaiting Update" && rowStatus != "On Latest Version" ) {
      ui.alert( "No LSA Workbook Found", "Please Select an LSA whose workbook has been created already", ui.ButtonSet.OK );
    }
    else {
      var lsaFileId = lsaWorksheet.getRange( rowNo, SHEETS.MASTER_LSAS.REFS.COL_NO_USERS_FILE_ID ).getValue();
      var lsasParentFolder = getLSAsParentFolderFromLSAWorkbookId( spreadsheet, ui, null, lsaFileId );
      if( lsasParentFolder ) {
        var alertHTMLTemplate = HtmlService.createTemplateFromFile('html/html-modal-open-lsadir');
        alertHTMLTemplate.openDirUrl = lsasParentFolder.getUrl();
        alertHTMLTemplate.openDirName = lsasParentFolder.getName();
        var alertMessageHTML = alertHTMLTemplate.evaluate().getContent();
        var alertMessage =  HtmlService.createHtmlOutput( alertMessageHTML ).setWidth(600).setHeight(160);
        ui.showModalDialog(alertMessage, 'Open LSAs Folder...');
      }
      else {
        ui.alert( "No Parent Folder", 
          "The LSAs workbook appears not to belong to a directory - please inform support as this should not happen", 
          ui.ButtonSet.OK );
      }
    }
  }
}


function DeleteLSAsDirectory() {

  var spreadsheet = SpreadsheetApp.getActive();
  var ui = SpreadsheetApp.getUi();
  var globalSettingsSheet = spreadsheet.getSheetByName( SHEETS.GLOBAL_SETTINGS.NAME );

  var selectedLSANos = getSelectedLSANos( spreadsheet, ui );
  //the above function ensures we're on the lsa worksheet
  var lsaWorksheet = spreadsheet.getActiveSheet();

  if( selectedLSANos.length > 1 ) {
    ui.alert( "Please select one LSA record", "Please select just one LSA record to use this function.", ui.ButtonSet.OK );
  }                        
  else if (selectedLSANos.length == 1 ) {
    var rowNo = selectedLSANos[0] + SHEETS.MASTER_LSAS.REFS.ROW_NO_FIRST_LSA - 1;
    let rowStatus: string = lsaWorksheet.getRange( rowNo, SHEETS.MASTER_LSAS.REFS.COL_NO_STATUS_BAR ).getValue();
    if( rowStatus != "Awaiting Update" && rowStatus != "On Latest Version" ) {
      ui.alert( "No LSA Workbook Found", "Please Select an LSA whose workbook has been created already", ui.ButtonSet.OK );
    }
    else {
      //get the info about the google group that the LSA's email record must belong to in order to pass validation
      //get the google group All LSAs ALL Campuses
      let allLSAsGroupEmail = globalSettingsSheet.getRange( 
            SHEETS.GLOBAL_SETTINGS.REFS.ROW_NO_LSA_GROUP_EMAIL, 
            SHEETS.GLOBAL_SETTINGS.REFS.COL_NO 
          ).getValue();
      let allLSAsGroupAdminURL = globalSettingsSheet.getRange( 
            SHEETS.GLOBAL_SETTINGS.REFS.ROW_NO_LSA_GROUP_ADMIN_URL, 
            SHEETS.GLOBAL_SETTINGS.REFS.COL_NO 
          ).getValue();
      let allLSAsGroup = GroupsApp.getGroupByEmail( allLSAsGroupEmail );

      //get info about the LSA who has been seleted for deletion
      let lsaEmailFromMaster = lsaWorksheet.getRange( rowNo, SHEETS.MASTER_LSAS.REFS.COL_NO_LSA_EMAIL ).getValue().trim();
      let lsaNameFromMaster = lsaWorksheet.getRange( rowNo, SHEETS.MASTER_LSAS.REFS.COL_NO_LSA_NAME ).getValue();      
      var lsaFileId = lsaWorksheet.getRange( rowNo, SHEETS.MASTER_LSAS.REFS.COL_NO_USERS_FILE_ID ).getValue();

      // double check that the email address has been removed from the All LSAs All Campuses group 
      // otherwise we wont allow deletion of sheet just yet - this ensures everything is kept tidy
      if( allLSAsGroup.hasUser(lsaEmailFromMaster) ) {
        ui.alert( "Please remove the user from the Google Group",
          "The email address '"+lsaEmailFromMaster+"' is still part of the google group: '" + allLSAsGroupEmail + "'\n\n" + 
          "Please remove them before deleting their workbook." +
          "To do this, please copy and paste this URL into a new tab:\n" + allLSAsGroupAdminURL + "\n\n" +
          "This page will allow you to remove the member from the google group, then you can retry deleting this LSAs Folder.",
          ui.ButtonSet.OK );
      }
      else {
        let areYouSureResponse = ui.alert('ARE YOU REALLY SURE?', 
              "Are you really sure that you want to delete the whole folder for " + lsaNameFromMaster + "?\n\n" +
              "The LSA's whole folder, and all of its contents, including the LSA's workbook will be sent to the " +
              "Trash Bin for 30 days before being PERMANENTLY DELETED.\n\n" +
              "If you are really sure, press OK to delete the folder", 
              ui.ButtonSet.OK_CANCEL );
        if( areYouSureResponse == ui.Button.OK )
        {
          spreadsheet.toast( 'Deleting LSA Directory and all files for ' + lsaNameFromMaster + '\n\n. Please wait',
                             'Deleting LSA Folder', 20 );
          let webAppReturnData = CallMasterRoSWebapp( null, "delete-lsa", null, lsaEmailFromMaster, 
                spreadsheet.getId(), null, rowNo, null, null, null, null, null, null );

          if( webAppReturnData == null || webAppReturnData.success != 1.0 ) {
            ui.alert( "Deletion Failed, please contact support", webAppReturnData.errorMessage, ui.ButtonSet.OK );
            return null;
          }
          else {
            ui.alert(
              "Finished Deleting LSA Folder", 
              lsaNameFromMaster + "'s folder has been deleted successfully",
              ui.ButtonSet.OK);
          }
        }
      }
    }
  }
}


function AnnounceToSelectedLSAs()
{
  var spreadsheet = SpreadsheetApp.getActive();
  var ui = SpreadsheetApp.getUi();  
  var masterAnnounceSheet = spreadsheet.getSheetByName( SHEETS.MASTER_ANNOUNCEMENTS.NAME );

  var selectedLSANos = getSelectedLSANos( spreadsheet, ui );
  //the above function ensures we're on the lsa worksheet
  var lsaWorksheet = spreadsheet.getActiveSheet();

  if( selectedLSANos.length > 0 )
  {
    var alertText = "The following LSAs are selected:\n\n";
    var noOfAnnouncementsToBePushed = 0;
    var recordNo = -1, rowNo = -1, rowStatus = null;
    for (var i=0; i<selectedLSANos.length; i++)
    {
      recordNo = selectedLSANos[i];
      rowNo = recordNo + SHEETS.MASTER_LSAS.REFS.ROW_NO_FIRST_LSA - 1;
      rowStatus = lsaWorksheet.getRange( rowNo, SHEETS.MASTER_LSAS.REFS.COL_NO_STATUS_BAR ).getValue();
      if( rowStatus == "Awaiting Update" || rowStatus == "On Latest Version") {
        noOfAnnouncementsToBePushed++;
        alertText += "  (" + recordNo + ") " + 
                     lsaWorksheet.getRange( rowNo, SHEETS.MASTER_LSAS.REFS.COL_NO_LSA_NAME ).getValue() + 
                     " - Will receive announcement\n";
      }
      else {
        alertText += "  (" + recordNo + ") - No Announcement\n";
      }
    }
    alertText += "\nIf you continue, a total of " + noOfAnnouncementsToBePushed + 
                 " Announcement"+(noOfAnnouncementsToBePushed==1?"":"s") + " will be sent.\n";
    alertText += "\nAre you sure you want to to send " + (noOfAnnouncementsToBePushed==1?"this":"these") + " " +
                 "Announcement"+(noOfAnnouncementsToBePushed==1?"":"s") + "?";

    var areYouSureResponse = ui.alert('Please Review Recipients', alertText, ui.ButtonSet.OK_CANCEL);
      
    if( noOfAnnouncementsToBePushed > 0 && areYouSureResponse == ui.Button.OK )
    {
      var approvedText = PreviewAnnouncement( spreadsheet, masterAnnounceSheet );
      if( approvedText != null ) {
        PushToChildWorksheets( selectedLSANos, null, approvedText, null, false, false, false );
        ClearDownProposedAnnouncement( spreadsheet, masterAnnounceSheet, true );
      } 
    }
  }
}