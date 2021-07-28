function pushAllHelpCentreSheetDate_NoVersionChecking( fromSheet, toSheet ) {
  try {
    let allData = fromSheet.getRange(
          SHEETS.MASTER_HELP.REFS.ROW_NO_FIRST_RECORD, SHEETS.MASTER_HELP.REFS.COL_NO_LINK_TEXT,
          SHEETS.MASTER_HELP.REFS.ROW_NO_LAST_RECORD - SHEETS.MASTER_HELP.REFS.ROW_NO_FIRST_RECORD + 1,
          SHEETS.MASTER_HELP.REFS.COL_NO_LINK_URL - SHEETS.MASTER_HELP.REFS.COL_NO_LINK_TEXT + 1
    ).getValues();

    toSheet.getRange(
          SHEETS.MASTER_HELP.REFS.ROW_NO_FIRST_RECORD, SHEETS.MASTER_HELP.REFS.COL_NO_LINK_TEXT,
          SHEETS.MASTER_HELP.REFS.ROW_NO_LAST_RECORD - SHEETS.MASTER_HELP.REFS.ROW_NO_FIRST_RECORD + 1,
          SHEETS.MASTER_HELP.REFS.COL_NO_LINK_URL - SHEETS.MASTER_HELP.REFS.COL_NO_LINK_TEXT + 1
    ).setValues( allData );
  }
  catch( e ) {
    debugCatchError( e );
    return e;
  }
  //test comment 
  return null;
}