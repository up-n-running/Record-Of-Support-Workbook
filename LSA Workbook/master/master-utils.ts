//var not const because blobs get populated on the fly
var LSA_EMAIL_STATICS = {
  HELP_VIDEO : {
    FILE_ID : "1twocba1Rkui3FF-qN7tYUdvx7ihTqqYi",   //for thumbnail image
    BLOB : null,
    VIDEO_URL: "https://youtu.be/158kQ7nf70s"  //full url to view video
  },
};

function SendLSAWelcomeEMail( lsaEmailAddress, lsaName, lsaWorkbookLink, spreadsheet, emailTemplate ) {
  //parse optional parameters if they were missing
  spreadsheet = (spreadsheet) ? spreadsheet : SpreadsheetApp.getActive();

Logger.log( "SendLSAWelcomeEMail Called" );
Logger.log( "lsaEmailAddress = " + lsaName );
Logger.log( "lsaWorkbookLink = " + lsaWorkbookLink );
Logger.log( "spreadsheet = " + spreadsheet );
Logger.log( "emailTemplate = " + emailTemplate );

  //inform user what's going on
  spreadsheet.toast( 'Sending a welcome email to the LSA at: ' + lsaEmailAddress, 'Sending Welcome Email' );

  //beild HTML as string for the email
  emailTemplate = (emailTemplate) ? emailTemplate : HtmlService.createTemplateFromFile('html/html-email-lsa-welcome');
  emailTemplate.lsaName = lsaName;
  emailTemplate.lsaWorkbookLink = lsaWorkbookLink;
  emailTemplate.helpVideoURL = LSA_EMAIL_STATICS.HELP_VIDEO.VIDEO_URL;
  let emailBodyHtml = emailTemplate.evaluate().getContent();
          
  Logger.log( "emailBodyHtml = " + emailBodyHtml );

  //get images as blobs if not already loaded into memory
  if( LSA_EMAIL_STATICS.HELP_VIDEO.BLOB == null ) {
    LSA_EMAIL_STATICS.HELP_VIDEO.BLOB = DriveApp.getFileById(LSA_EMAIL_STATICS.HELP_VIDEO.FILE_ID).getBlob();
  }

  Logger.log( "LSA_EMAIL_STATICS.HELP_VIDEO.BLOB.getName() = " + LSA_EMAIL_STATICS.HELP_VIDEO.BLOB.getName() );

  //send it        
  MailApp.sendEmail({ 
    to: lsaEmailAddress,
    subject: "Welcome to your Record of Support Workbook",
    htmlBody: emailBodyHtml,
    inlineImages: {
      helpVideoThumb: LSA_EMAIL_STATICS.HELP_VIDEO.BLOB
    }
  });

}





/**
 * Locate the actual cell in global settings of the setting represented by settingHandle
 * This is intended to be used when we have an old version of the spreadsheet open, where we have no guarantees
 * that the row numbers latch with the latest definitions defined in sheet-defs.gs
 * 
 * It does this by looking for a named range called "GLOBAL_SETTINGS_"+settingHandle (but without the [ and the ] chars)
 * and if there isnt one, it will look for the 
 * handle in column c, and if its not there either, then just use the current master version of the row num from sheet-defs.gs
 * as we have no other choice!
 *
 * @oldSpreadsheet {Spreadsheet} The Spreadsheet object containing the oldVersionSettingsSheet
 * @oldVersionSettingsSheet {Sheet} The Global Settings Sheet from the Spreadsheet
 * @settingHandle {String} The Setting Handle as a capitalised string
 * @fallBackSettingRow {integer} The row number sheet-defs.gs for current versions of the Global Settings Sheet
 * @return {integer} The row number of the found setting
 */
function GetMasterSettingsCellFromOldVersionWorkbook( oldSpreadsheet, oldVersionSettingsSheet, settingHandle, fallBackSettingRow ) {
  let oldWorkbookSettingFoundRowNum = GetMasterSettingsRowFromOldVersionWorkbook( oldSpreadsheet, oldVersionSettingsSheet, settingHandle, fallBackSettingRow );
  
  return ( oldWorkbookSettingFoundRowNum > 0 ) ? oldVersionSettingsSheet.getRange( 
          oldWorkbookSettingFoundRowNum,  SHEETS.GLOBAL_SETTINGS.REFS.COL_NO ) :
          null;
}

/**
 * Locate the row number in global settings of the setting represented by settingHandle
 * This is intended to be used when we have an old version of the spreadsheet open, where we have no guarantees
 * that the row numbers latch with the latest definitions defined in sheet-defs.gs
 * 
 * It does this by looking for a named range called "GLOBAL_SETTINGS_"+settingHandle (but without the [ and the ] chars)
 * and if there isnt one, it will look for the 
 * handle in column c, and if its not there either, then just use the current master version of the row num from sheet-defs.gs
 * as we have no other choice!
 *
 * @oldSpreadsheet {Spreadsheet} The Spreadsheet object containing the oldVersionSettingsSheet
 * @oldVersionSettingsSheet {Sheet} The Global Settings Sheet from the Spreadsheet
 * @settingHandle {String} The Setting Handle as a capitalised string
 * @fallBackSettingRow {integer} The row number sheet-defs.gs for current versions of the Global Settings Sheet
 * @return {integer} The row number of the found setting
 */
function GetMasterSettingsRowFromOldVersionWorkbook( oldSpreadsheet, oldVersionSettingsSheet, settingHandle, fallBackSettingRow ) {
  var fallBackSettingRow = ( fallBackSettingRow ) ? fallBackSettingRow : -1;
  var foundCell = oldSpreadsheet.getRangeByName( "GLOBAL_SETTINGS_" + settingHandle.substr( 1, settingHandle.length - 2) );
  var rowNum = (foundCell == null) ? -1 : foundCell.getRow();
  rowNum = (rowNum==-1) ? findInColumn( oldVersionSettingsSheet, settingHandle,
                                          SHEETS.GLOBAL_SETTINGS.REFS.COL_NO + 1, 1, 
                                          SHEETS.GLOBAL_SETTINGS.REFS.ROW_NO_PENDING_ANNOUNCEMENT + 5 ) : rowNum;
  rowNum = ( rowNum==-1) ? fallBackSettingRow : rowNum;
  Logger.log( "GetMasterSettingsCellFromOldVersionWorkbook using row:" + rowNum + " for setting: " + settingHandle ); 
        //NOW WE DO THE SAME FOR SHEETS.GLOBAL_SETTINGS.REFS.ROW_NO_MASTERS_VERSION
  return rowNum;
}





/**
 * [master-utils.gs]
 * Locate the header cell of a particular column in a spreadsheet that might not be this spreadsheet. The return a struct with the
 * col no, the row number of the row below the header and the row number of the last row on the sheet.
 * 
 * This is intended to be used when we have a potentially old version of the spreadsheet, where we have no guarantees
 * that the col numbers match with the latest definitions defined in sheet-defs.gs. if you know it's the right version then
 * set dontCheckNamedRangesToSaveTime to true an it will ignore named ranges
 * 
 * It does this by looking for a named range called SHEET_NAME_COLUMN_NAME and if there isnt one, it will just use the col data from
 * sheet-defs.gs (which you have to pass in) as we have no other choice!
 *
 * @param spreadsheet {Spreadsheet} The Spreadsheet object, that might be a different version, that we're looking at
 * @param sheetHandle {String} The Handle of the Sheet as a capitalised string eg SHEET_NAME from the named gange SHEET_NAME_COLUMN_NAME
 * @param columnHandle {String} The Handle of the Col as a capitalised string eg COLUMN_NAME from the named gange SHEET_NAME_COLUMN_NAME
 * @param fallBackColumnNo {integer} The col number from sheet-defs.gs for current version
 * @param fallback1stDataRow {integer} The 1st data row from sheet-defs.gs for current version
 * @param fallbackLastDataRow {integer=} The last data row from sheet-defs.gs for current version (or blank if it's variable length)
 * @param sheetIfNoLastFallback {Sheet=} MUST BE SET IF fallbackLastDataRow IS NOT PASSED SO IT CAN COUNT THE ROWS ON THE SHEET
 * @param useRowAndSheetDataFromThisAdjacentColDefToSaveTime {Struct=} use if ur making many calls in a row
 * @param dontCheckNamedRangesToSaveTime {bool=} if you know its the right version you can pass true here as we know col defs are right
 * @return {integer} The column number of the found setting
 */
function GetColDefByHandleFromAnyWorkbook( spreadsheet, sheetHandle, columnHandle, 
          fallBackColumnNo, fallback1stDataRow, fallbackLastDataRow, sheetIfNoLastFallback, 
          useRowAndSheetDataFromThisAdjacentColDefToSaveTime, dontCheckNamedRangesToSaveTime,
          forceUseNamedRangesAndReturnNullIfNoMatch ) {

  //Logger.log( "GetColDefByHandleFromAnyWorkbook called for setting: " + sheetHandle + "_" + columnHandle + ", dontCheckNamedRangesToSaveTime = " + dontCheckNamedRangesToSaveTime );
  let foundCell = null;
  let colDef = {
    colNo           : fallBackColumnNo,
    sheet           : sheetIfNoLastFallback,
    firstDataRowNo  : fallback1stDataRow,
    lastDataRowNo   : fallbackLastDataRow
  };
  
  if( !(dontCheckNamedRangesToSaveTime) ) {
    foundCell = spreadsheet.getRangeByName( sheetHandle + "_" + columnHandle );
  }

  if( forceUseNamedRangesAndReturnNullIfNoMatch && foundCell == null ) {
    return null;
  }

  if( foundCell ) {
    let o = useRowAndSheetDataFromThisAdjacentColDefToSaveTime;
    sheetIfNoLastFallback = ( o ) ? o.sheet : foundCell.getSheet();  //in case sheet passed in is wrong
    colDef.colNo          = foundCell.getColumn();
    colDef.sheet          = sheetIfNoLastFallback;
    colDef.firstDataRowNo = ( o ) ? o.firstDataRowNo : foundCell.getRow() + foundCell.getHeight();
    colDef.lastDataRowNo  = ( o ) ? o.lastDataRowNo : sheetIfNoLastFallback.getMaxRows();
  } else if( useRowAndSheetDataFromThisAdjacentColDefToSaveTime ) {
    let o = useRowAndSheetDataFromThisAdjacentColDefToSaveTime;
    colDef.sheet = o.sheet;
    colDef.firstDataRowNo = o.firstDataRowNo;
    colDef.lastDataRowNo = o.lastDataRowNo;
  }
  else if ( !(fallbackLastDataRow) ) {
    colDef.lastDataRowNo  = sheetIfNoLastFallback.getMaxRows();
  }

  //Logger.log( "GetColDefByHandleFromAnyWorkbook, using col def" + colDef + " for setting: " + sheetHandle + "_" + columnHandle ); 
  return colDef;
}

function debugColDef( colDef ) {
  return ( !colDef ) ? 'NULL' : ( "colNo: " + colDef.colNo + "' rows [" +colDef.firstDataRowNo+ "," +colDef.lastDataRowNo+ "], sheet = " + colDef.sheet );
}

/**
 * [master-utils.gs]
 * Take a col def from GetColDefByHandleFromAnyWorkbook and get the column data range
 *
 * @param colDef {Struct} as returned from GetColDefByHandleFromAnyWorkbook
 * @return {integer} The column number of the found setting
 */
function GetRangeFromColDef( colDef ) {
  return ( colDef ) ? 
            colDef.sheet.getRange( colDef.firstDataRowNo, colDef.colNo, colDef.lastDataRowNo-colDef.firstDataRowNo+1, 1 ) :
            null;
}


function getLSAsParentFolderFromLSAWorkbookId( spreadsheet: GoogleAppsScript.Spreadsheet.Spreadsheet, ui,
        globalSettingsSheet, lsaWorkbookId ) {

  //parse params
  if( !globalSettingsSheet ) {
    spreadsheet = ( spreadsheet ) ? spreadsheet : SpreadsheetApp.getActive();
    globalSettingsSheet = spreadsheet.getSheetByName( SHEETS.GLOBAL_SETTINGS.NAME ); 
  }

  //get masterLsaRootFolderId setting
  let masterLsaRootFolderId = globalSettingsSheet.getRange(
        SHEETS.GLOBAL_SETTINGS.REFS.ROW_NO_DIR_ID_LSAS,
        SHEETS.GLOBAL_SETTINGS.REFS.COL_NO
  ).getValue();

  //initialise search vars
  let lsaWorkbookFile: GoogleAppsScript.Drive.File = DriveApp.getFileById(lsaWorkbookId); 
  let workbooksParentFolderIterator: GoogleAppsScript.Drive.FolderIterator = lsaWorkbookFile.getParents();
  let tempWorkbooksParentFolder: GoogleAppsScript.Drive.Folder|null = null;
  let tempWorkbooksParentFolderParentFolderIterator: GoogleAppsScript.Drive.FolderIterator|null = null;
  let tempWorkbooksParentFolderTempParentFolder: GoogleAppsScript.Drive.Folder|null = null;
  let foundIt: boolean = false;

  //search
  while( !foundIt && workbooksParentFolderIterator.hasNext() ) {
    tempWorkbooksParentFolder = workbooksParentFolderIterator.next();
//Logger.log( "getLSAsParentFolderFromLSAWorkbookId checking parent: " + tempWorkbooksParentFolder );
    tempWorkbooksParentFolderParentFolderIterator = tempWorkbooksParentFolder.getParents();
    while( !foundIt && tempWorkbooksParentFolderParentFolderIterator.hasNext() ) {
      tempWorkbooksParentFolderTempParentFolder = tempWorkbooksParentFolderParentFolderIterator.next();
//Logger.log( "getLSAsParentFolderFromLSAWorkbookId checking parent's parent: " + tempWorkbooksParentFolderTempParentFolder );
      if( tempWorkbooksParentFolderTempParentFolder.getId() === masterLsaRootFolderId ) {
        foundIt = true;
      }
    }
  }

  if( foundIt ) {
//Logger.log( "getLSAsParentFolderFromLSAWorkbookId FOUND parent: " + tempWorkbooksParentFolder );
    return tempWorkbooksParentFolder;
  }
  else {
    if( ui ) {
      ui.alert( "Could not find LSAs Directory", "We could not find a parent folder to the LSAs Workbook with file id: '" +
                lsaWorkbookId + "' which, in turn, has a parent folder of the Master Root LSAs Directory with Directory Id: '" +
                + masterLsaRootFolderId + "'.", ui.ButtonSet.OK );
    }
    return null;
  }

}