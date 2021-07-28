function createLearnerRoSFromSheetsTemplate_( 
      sourceRoSPreviewSheet: GoogleAppsScript.Spreadsheet.Sheet, 
      destLearnerRosFolder: GoogleAppsScript.Drive.Folder, 
      destFileName: string,
      learnerEmailAddress: string,
      learnerSignatureFileId: string ) {

  //fetch the sheets we need from the LSA Workbook
  let spreadsheet = SpreadsheetApp.getActive();
  let globalSettingsSheet = spreadsheet.getSheetByName( SHEETS.GLOBAL_SETTINGS.NAME );

  //get Source template File
  let learnerRoSTemplateFileId = globalSettingsSheet.getRange (
      SHEETS.GLOBAL_SETTINGS.REFS.ROW_NO_FILE_ID_LEARNER_ROS,
      SHEETS.GLOBAL_SETTINGS.REFS.COL_NO
  ).getValue();
  let learnerRoSTemplateFile = DriveApp.getFileById( learnerRoSTemplateFileId );
  //let learnerRoSTemplateSpreadsheet = SpreadsheetApp.open( learnerRoSTemplateFile );

  //create Dest Learner RoS File
  let learnerDestRoSFile = learnerRoSTemplateFile.makeCopy( destFileName, destLearnerRosFolder );
  let learnerDestRoSSpreadsheet = SpreadsheetApp.open( learnerDestRoSFile );
  let learnerDestRoSSheet = learnerDestRoSSpreadsheet.getSheetByName( SHEETS.LEARNER_TEMPLATE_REMOTE.NAME );

  //copy data from source to dest
  let copyScript: Array<any> = SHEETS.LEARNER_TEMPLATE_REMOTE.COPY_SCRIPT;
  let copyScriptRow: any = null;
  let sourceRoSRange: GoogleAppsScript.Spreadsheet.Range|null = null;
  let destRoSRange: GoogleAppsScript.Spreadsheet.Range|null = null;
  let tempValue: any = null;
  for( let i:number = 0; i < copyScript.length; i++ ) {
  copyScriptRow = copyScript[ i ];
  sourceRoSRange = sourceRoSPreviewSheet.getRange( copyScriptRow.SOURCE_RANGE );
  destRoSRange = learnerDestRoSSheet.getRange( 
        copyScriptRow.DEST_ROW, copyScriptRow.DEST_COL, 
        sourceRoSRange.getHeight(), sourceRoSRange.getWidth() 
  );
  if( copyScriptRow.SINGLE_VALUE ) {
    tempValue = sourceRoSRange.getValue();
    Logger.log( "tempValue = '" + tempValue + "'" ); 
    if( copyScriptRow.COPY_BLANK || tempValue ) {
      Logger.log( "Saving value" ); 
      destRoSRange.setValue( tempValue );
    }
  }
  else {
    Logger.log( "Saving range" ); 
    destRoSRange.setValues( sourceRoSRange.getValues() );
  }
  }

  //work out if the sheet is already auto-signed
  let signatureValue = sourceRoSPreviewSheet.getRange(
      SHEETS.MASTER_TEMPLATE.REFS.ROW_NO_SIGNATURE,
      SHEETS.MASTER_TEMPLATE.REFS.COL_NO_SIGNATURE
  ).getValue(); 
  Logger.log( "signatureValue = '" + signatureValue + "'" );
  let isAutoSigned = ( signatureValue != "" );

  //set dest hidden settings
  learnerDestRoSSheet.getRange( 
      SHEETS.LEARNER_TEMPLATE_REMOTE.REFS.ROW_NO_HIDDEN_SETTINGS,
      SHEETS.LEARNER_TEMPLATE_REMOTE.REFS.COL_NO_LEARNER_EMAIL
  ).setValue( learnerEmailAddress );
  learnerDestRoSSheet.getRange( 
      SHEETS.LEARNER_TEMPLATE_REMOTE.REFS.ROW_NO_HIDDEN_SETTINGS,
      SHEETS.LEARNER_TEMPLATE_REMOTE.REFS.COL_NO_LEARNER_SIGNATURE_FILE_ID
  ).setValue( learnerSignatureFileId );

  //set Permissions on Cells for Learner so learner has access to just the cells they need
  //these are the cells with a range level ptorection on them.
  if( !isAutoSigned ) {
    let rangeProtections = learnerDestRoSSheet.getProtections(SpreadsheetApp.ProtectionType.RANGE);
    let p: GoogleAppsScript.Spreadsheet.Protection|null = null;
    for ( let j:number = 0; j < rangeProtections.length; j++ ) {
      p = rangeProtections[j];
      p.addEditors( [ learnerEmailAddress ] );
    }
  }

  //FINAL thing is to give the learner permission to access the file
  learnerDestRoSFile.addEditor( learnerEmailAddress );

  return learnerDestRoSFile;
}



function saveRosSpreadsheetToGivenFolder( spreadsheet, rootRoSDirectoryId, learnerObj, learnerName, startTime, duration, lsaName, recordDate, recordNo, isAutoSign,
                                          sourceRoSPreviewSheet, learnerEmailAddress, ui: any ) {

  spreadsheet = (spreadsheet) ?spreadsheet : SpreadsheetApp.getActive();

  uiSensitiveToast( spreadsheet, ui, 'Saving the Learner RoS File to the Shared Drive', 'Generating Record ' + recordNo, 12 );

  //get learner folder
  let learnerMonthFolder = getOrCreateLearnerYearAndMonthSubDir( spreadsheet, recordDate, learnerName, 
    learnerObj, rootRoSDirectoryId );

  //parse start time and duration information and get filename
  let lessonTimeObj = getlessonTimesObjectFromInputSheetValues( recordDate, startTime, duration );
  let rosFilename = generateRosFileName( learnerName, lessonTimeObj, lsaName, isAutoSign, false ) ;

  //let rosFile = createLearnerRoSFromSheetsTemplate_( sourceRoSPreviewSheet, learnerMonthFolder, rosFilename, 
  //        learnerEmailAddress, learnerObj.SIGNATURE_ID );

  let globalSettingsSheet = spreadsheet.getSheetByName( SHEETS.GLOBAL_SETTINGS.NAME );
  let learnerRoSTemplateFileId = globalSettingsSheet.getRange (
        SHEETS.GLOBAL_SETTINGS.REFS.ROW_NO_FILE_ID_LEARNER_ROS,
        SHEETS.GLOBAL_SETTINGS.REFS.COL_NO
  ).getValue();
  let webAppReturnData = CallMasterRoSWebapp( globalSettingsSheet, "generate-ros-stored-signature", spreadsheet.getId(),
        null, null, null, null, recordNo, learnerMonthFolder.getId(), learnerRoSTemplateFileId,  rosFilename,  learnerEmailAddress,
        learnerObj.SIGNATURE_ID );

  if( webAppReturnData.success != 1.0 ) {
    ui.alert( "Could not Create Record Of Support", 
              "Creating Record Of Support failed for record number: " + recordNo + ".\n" +
              "Error Message:\n" +
              webAppReturnData.errorMessage,
              ui.ButtonSet.OK );
    return null;
  }
  else {
    return DriveApp.getFileById( webAppReturnData.affectedFileId );
  }

}