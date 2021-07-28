function getSelectedRecordNos( spreadsheet, ui ) {
  return getAllSelectedRowOrColumnRecordNumbers( false, spreadsheet, ui, SHEETS.INPUT.NAME,
                                           SHEETS.INPUT.REFS.ROW_NO_STATUS_BAR, SHEETS.INPUT.REFS.ROW_NO_RECORD_NOS, 
                                           SHEETS.INPUT.REFS.COL_NO_RECORD_1, SHEETS.INPUT.REFS.COL_NO_RECORD_LAST, null );
}

function getSelectedLSANos( spreadsheet, ui ) {
  return getAllSelectedRowOrColumnRecordNumbers( true, spreadsheet, ui, SHEETS.MASTER_LSAS.NAME,
                                           SHEETS.MASTER_LSAS.REFS.ROW_NO_FIRST_LSA, SHEETS.MASTER_LSAS.REFS.ROW_NO_LAST_LSA, 
                                           SHEETS.MASTER_LSAS.REFS.COL_NO_STATUS_BAR, SHEETS.MASTER_LSAS.REFS.COL_NO_LSA_NOS, null );
}


function getAllSelectedRowOrColumnRecordNumbers( returnRowNumsNotColumnNums, spreadsheet, ui, sheetName,
                                         rMin, rMax, cMin, cMax, friendlyOutOfBoundsHelpText ) {
  var selectedRowOrColRecordNumbers = new Array();

  if( spreadsheet.getActiveSheet().getSheetName()  != sheetName ) {
      ui.alert( "You are on the wrong Sheet", 
                "You must select cells on the sheet '" + sheetName + "' in order to use this function.", 
                ui.ButtonSet.OK );
      return new Array();
  }

  var selectedRanges = SpreadsheetApp.getActive().getActiveRangeList().getRanges();

  var rTop = -1, rBottom = -1, cLeft = -1, cRight = -1;

  //if there is an arror and we have to feed back to the user with an explation and some
  //help toext to explain the correct cells to click on then we'll need these:
  var helpDesc = returnRowNumsNotColumnNums ? "column" : "row";
  var min = returnRowNumsNotColumnNums ? cMin : rMin;
  var max = returnRowNumsNotColumnNums ? cMax : rMax;
  helpDesc = ( min == max ) ? helpDesc : ( helpDesc + "s" );
  var fullHelpText = ( friendlyOutOfBoundsHelpText ) ? friendlyOutOfBoundsHelpText : ( "Please only select cells from the coloured " + 
            "status bar on " + helpDesc + " " + min + ( ( min == max ) ? "" : ( ( ( min+1 == max ) ? " and " : " to " ) + max ) ) +
            " and then try again." );

  //loop through each of the ranges in the selection
  for(var i = 0; i < selectedRanges.length; i++) {

    var iRange = selectedRanges[i];

    rTop = iRange.getRow();
    rBottom = rTop + iRange.getHeight() - 1;
    cLeft = iRange.getColumn();
    cRight = cLeft + iRange.getWidth() - 1;

    if( rTop<rMin || rBottom > rMax || cLeft < cMin || cRight > cMax ) {

      ui.alert( "Selected cells out of bounds", 
                "The selected cells includes the " + 
                    ( (iRange.getHeight()*iRange.getWidth()) > 1 ? "range " : "cell " ) + iRange.getA1Notation() + "\n" +
                "Which is outside of the allowed range: " + 
                columnToLetter( cMin ) + rMin + ":" +columnToLetter( cMax ) + rMax + ".\n" +
                "\n" +
                fullHelpText,
                ui.ButtonSet.OK );
      return new Array();
    }

    //fetch all record numbers (or LSA numbers or File Numbers etc) from currect range in the selection
    for( var c = cLeft; !returnRowNumsNotColumnNums && c<=cRight; c++ ) {
      if( selectedRowOrColRecordNumbers.indexOf( c - cMin + 1 ) < 0 ) {
        selectedRowOrColRecordNumbers.push( c - cMin + 1 );
      }
    }
    //fetch all (or LSA numbers or File Numbers etc) from the current range in the selection
    for( var r = rTop; returnRowNumsNotColumnNums && r<=rBottom; r++ ) {
      if( selectedRowOrColRecordNumbers.indexOf( r - rMin + 1 ) < 0 ) {
        selectedRowOrColRecordNumbers.push( r - rMin + 1 );
      }
    }
  }

  if( selectedRowOrColRecordNumbers.length == 0 ) {
    ui.alert( "No Cells Were Selected", 
              "No cells were selected so no action can be taken.\n" +
              "\n" +
              fullHelpText,
              ui.ButtonSet.OK );
    return selectedRowOrColRecordNumbers;
  }

  //now sort the array of record numbers ascending so it's more nearerer
  selectedRowOrColRecordNumbers.sort(function(a, b){return a-b;});

  Logger.log( 'selectedRowOrColRecordNumbers = ' + selectedRowOrColRecordNumbers ); 

  return selectedRowOrColRecordNumbers;
}


/**
 * [utils-spreadsheet.gs]
 * Search through a ROW of data looking for a particular value and returns the column number (or -1) if not found
 * @param workSheet {Sheet} The Worksheet with the column we're looking through
 * @param valueToFind {Object} The Value to find in the column
 * @param lookupRowNo {number} The number of the row to look through
 * @param leftColNo {number} the number of the left hand column to start looking from
 * @param topRowNo {number} the number of the right hand column to stop looking at
 * @return {number} The column number of the first instance where the value is found, or -1 if not found
 */
function findInRow( workSheet, valueToFind: any, lookupRowNo, leftColNo, rightColNo, caseSensitive?: boolean ) {

  //parse params
  caseSensitive = caseSensitive ? true : false;
  valueToFind = ( !caseSensitive && valueToFind && typeof( valueToFind.toUpperCase ) == 'function' ) ? valueToFind.toUpperCase() : valueToFind;

  // create an array of data from columns lookupColumnNo through returnColumnNo
  let data: Array<Array<any>> = workSheet.getRange( lookupRowNo, leftColNo, 1, rightColNo - leftColNo +1 ).getValues();
  
  //loop through array looking for valueToFind
  let foundOnColNum: number = -1;
  let lookupValue: any = null;
  for( let cIndex=0 ; cIndex<data[0].length && foundOnColNum < 0 ; cIndex++ ) {
    lookupValue = data[0][cIndex];
    lookupValue = ( !caseSensitive && lookupValue && typeof( lookupValue.toUpperCase ) == 'function' ) ? lookupValue.toUpperCase() : lookupValue;
    if (lookupValue==valueToFind) {
      foundOnColNum = cIndex + leftColNo;
    } // if a match in foundOnColNum is found, break the loop
  }

  return foundOnColNum;
}


/**
 * Search through a COLUMN of data looking for a particular value and returns the row number (or -1) if not found
 * @param workSheet {Sheet} The Worksheet with the column we're looking through
 * @param valueToFind {Object} The Value to find in the column
 * @param lookupColumnNo {number} The number of the column to look through
 * @param topRowNo {number} the top row to start looking from
 * @param bottomRowNo {number=} the bottom row to stop looking at
 * @return {number} The row number of the first instance where the value is found, or -1 if not found
 */
function findInColumn( workSheet: GoogleAppsScript.Spreadsheet.Sheet, valueToFind: any, 
                       lookupColumnNo: number, topRowNo: number, bottomRowNo?: number|null, caseSensitive?: boolean ): number {

  //parse params
  caseSensitive = caseSensitive ? true : false;
  valueToFind = ( !caseSensitive && valueToFind && typeof( valueToFind.toUpperCase ) == 'function' ) ? valueToFind.toUpperCase() : valueToFind;
  bottomRowNo = ( bottomRowNo ) ? bottomRowNo : workSheet.getMaxRows();

  // create an array of data from columns lookupColumnNo through returnColumnNo
  let data=workSheet.getRange( topRowNo, lookupColumnNo, bottomRowNo - topRowNo +1, 1 ).getValues();

  //loop through array looking for valueToFind
  let foundOnRowNum = -1;
  let lookupValue: any = null;
  for( let rIndex=0 ; rIndex<data.length && foundOnRowNum < 0 ; rIndex++ ) {
    lookupValue = data[rIndex][0];
    lookupValue = ( !caseSensitive && lookupValue && typeof( lookupValue.toUpperCase ) == 'function' ) ? lookupValue.toUpperCase() : lookupValue;
    //Logger.log( "lookupValue = " + lookupValue );
    //Logger.log( "valueToFind = " + valueToFind );
    if (lookupValue==valueToFind) {
      foundOnRowNum = rIndex + topRowNo;
    } // if a match in lookupColumnNo is found, break the loop
  }

  return foundOnRowNum;
}


function findAllInColumn( workSheet, valueToFind: any, lookupColumnNo, topRowNo, bottomRowNo?: number|null, caseSensitive?: boolean ) {

  //parse params
  caseSensitive = caseSensitive ? true : false;
  valueToFind = ( !caseSensitive && valueToFind && typeof( valueToFind.toUpperCase ) == 'function' ) ? valueToFind.toUpperCase() : valueToFind;
  bottomRowNo = ( bottomRowNo ) ? bottomRowNo : workSheet.getMaxRows();

  // create an array of data from columns lookupColumnNo through returnColumnNo
  let data=workSheet.getRange( topRowNo, lookupColumnNo, bottomRowNo - topRowNo +1, 1 ).getValues();

  //loop through array looking for valueToFind
  let foundOnRowNums = new Array();
  let lookupValue: any = null;
  for( let rIndex=0 ; rIndex<data.length; rIndex++ ) {
    lookupValue = data[rIndex][0];
    lookupValue = ( !caseSensitive && lookupValue && typeof( lookupValue.toUpperCase ) == 'function' ) ? lookupValue.toUpperCase() : lookupValue;
    //Logger.log( "lookupValue = " + lookupValue );
    //Logger.log( "valueToFind = " + valueToFind );
    if (lookupValue==valueToFind) {
      foundOnRowNums.push( rIndex + topRowNo );
    } // if a match in lookupColumnNo is found, break the loop
  }

  return foundOnRowNums;
}

/**
 * Search through a range of data looking for a first row where all cells are blank returns the row number (or -1) if not found
 * @param workSheet {Sheet} The Worksheet with the column we're looking through
 * @param firstColumnNo {number} The number of the start column for the cols that need to be blank
 * @param lastColumnNo {number} The number of the end column for the cols that need to be blank
 * @param topRowNo {number} the top row to start looking from
 * @param bottomRowNo {number=} the bottom row to stop looking at
 * @return {number} The row number of the first instance where the value is found, or -1 if not found
 */
 function findFirstBlankRow( workSheet, firstColumnNo, lastColumnNo, topRowNo, bottomRowNo ) {

  //parse params
  bottomRowNo = ( bottomRowNo ) ? bottomRowNo : workSheet.getMaxRows();

  // create an array of data from columns lookupColumnNo through returnColumnNo
  let data=workSheet.getRange( topRowNo, firstColumnNo, bottomRowNo - topRowNo +1, lastColumnNo - firstColumnNo + 1 ).getValues();

  //loop through array looking for valueToFind
  let foundOnRowNum = -1;
  let lookupValues = [];
  let foundANonBlankCell = false;
  for( let rIndex=0 ; rIndex<data.length && foundOnRowNum < 0 ; rIndex++ ) {
    lookupValues = data[rIndex];
    foundANonBlankCell = false;
    for( let cIndex=0 ; !foundANonBlankCell && cIndex<lookupValues.length; cIndex++ ) {
      if( typeof( lookupValues[ cIndex ].trim ) == "function" ) { foundANonBlankCell = lookupValues[ cIndex ].trim() != ""; }
      else { foundANonBlankCell = lookupValues[ cIndex ] != ""; }
    }
    if (!foundANonBlankCell) {
      foundOnRowNum = rIndex + topRowNo;
    } // if a match in lookupColumnNo is found, break the loop
  }

  return foundOnRowNum;
}

function deleteCellsOnRowAndMoveUp( delRowNum, colNumArray, workSheet, bottomRowNum ) {
  //set bottom row to last row on sheet if no param is passed in
  bottomRowNum = (bottomRowNum) ? bottomRowNum : workSheet.getMaxRows();

  //loop through column number array
  var colNum = null;
  for( let c=0; c < colNumArray.length; c++ ) {
    //for each column number...

    colNum = colNumArray[ c ];

    //move all the cells below the row to be deleted, up one
    if( delRowNum < bottomRowNum ) {
      let valuesBelowArray = workSheet.getRange( delRowNum+1, colNum, bottomRowNum - delRowNum, 1 ).getValues();
      workSheet.getRange( delRowNum, colNum, bottomRowNum - delRowNum, 1 ).setValues( valuesBelowArray );
    }

    //now we've moved everything up one row, the row at the vey bottom can be blanked out
    workSheet.getRange( bottomRowNum, colNum, 1, 1 ).setValue( "" );
  }
}


function deleteCellsOnRowAndMoveUp_ColumnRange( delRowNum, lColumnNo, rColumnNo, workSheet, bottomRowNum ) {
  //set bottom row to last row on sheet if no param is passed in
  bottomRowNum = (bottomRowNum) ? bottomRowNum : workSheet.getMaxRows();

  //move all below values up one
  if( delRowNum < bottomRowNum ) {
    let valuesBelowArray = workSheet.getRange( delRowNum+1, lColumnNo, bottomRowNum - delRowNum, rColumnNo-lColumnNo+1 ).getValues();
    workSheet.getRange( delRowNum, lColumnNo, bottomRowNum - delRowNum, rColumnNo-lColumnNo+1 ).setValues( valuesBelowArray );
  }

  //now we've moved everything up one row, the row at the vey bottom can be blanked out
  workSheet.getRange( bottomRowNum, lColumnNo, 1, rColumnNo-lColumnNo+1 ).setValue( "" );
}


function deleteCellsOnColumnAndMoveLeft_RowRange( delColumnNum, tRowNo, bRowNo, workSheet, lastColumnNum ) {
  //set last column to last column on sheet if no param is passed in
  lastColumnNum = (lastColumnNum) ? lastColumnNum : workSheet.getMaxColumns();

  //move all below values up one
  if( delColumnNum < lastColumnNum ) {
    let valuesAfterArray = workSheet.getRange( tRowNo, delColumnNum+1, bRowNo-tRowNo+1, lastColumnNum-delColumnNum ).getValues();
    workSheet.getRange( tRowNo, delColumnNum, bRowNo-tRowNo+1, lastColumnNum-delColumnNum ).setValues( valuesAfterArray );
  }

  //now we've moved everything up one row, the row at the vey bottom can be blanked out
  workSheet.getRange( tRowNo, lastColumnNum, bRowNo-tRowNo+1, 1 ).setValue( "" );
}


function addCellsOnRowAndMoveDown( rowNum, colNumArray, valueArray, workSheet, bottomRowNum ) {
  //set bottom row to last row on sheet if no param is passed in
  bottomRowNum = (bottomRowNum) ? bottomRowNum : workSheet.getMaxRows();

  //loop through column number array
  let colNum = null;
  for( let c=0; c < colNumArray.length; c++ ) {
    //for each column number...
    colNum = colNumArray[ c ];
    //move all the cells apart from the bottom row, down one
    let valuesBelowArray = workSheet.getRange( rowNum, colNum, bottomRowNum - rowNum, 1 ).getValues();
    workSheet.getRange( rowNum+1, colNum, bottomRowNum - rowNum, 1 ).setValues( valuesBelowArray );
    //now we've moved everything down one row, save the value in the space created on 
    workSheet.getRange( rowNum, colNum, 1, 1 ).setValue( valueArray[c] );
  }
}


//looks down a column starting from startRowNum returning the row num of the first empty cell, or null if no empty cells
function findFirstEmptyCellInColumn( startRowNum, colNum, workSheet, bottomRowNum ) {
  //set bottom row to last row on sheet if no param is passed in
  bottomRowNum = (bottomRowNum) ? bottomRowNum : workSheet.getMaxRows();

Logger.log( "findFirstEmptyCellInColumn called" );
Logger.log( "startRowNum = " + startRowNum );
Logger.log( "bottomRowNum = " + bottomRowNum );
Logger.log( "colNum = " + colNum );
Logger.log( "workSheet = " + workSheet );

  var columnValuesArray = workSheet.getRange( startRowNum, colNum, bottomRowNum - startRowNum + 1, 1 ).getValues();
Logger.log( "columnValuesArray = " + columnValuesArray );
  var ct = 0;
  while ( ct < columnValuesArray.length && columnValuesArray[ct][0] != "" ) {
    ct++;
  }
  return (ct == columnValuesArray.length ) ? null : (ct+startRowNum);
}


/**
 * [utils-spreadsheet.gs]
 * parses findAndReplaceJSONString reasy for mutiple calls to findAndReplaceExactCellValuesInRange
 * See findAndReplaceExactCellValuesInRange documentation for more info
 * THIS IS VALIDATION ONLY AND RETURNS null IFF VALID
 * @param findAndReplaceJSONString {String} A String storing valid JSON used to define 0,1 or more F & R operations.
 * @return {String} null if valid, error String if not
 */
function validateFindAndReplaceJSON( findAndReplaceJSONString ) {
  let findAndReplaceJSON = null;
  try {
    findAndReplaceJSON = JSON.parse( findAndReplaceJSONString );
  }
  catch( e ) {
    debugCatchError( e );
    return "The find And replace JSON String is not valid JSON\n\nnfindAndReplaceJSONString:\n"+findAndReplaceJSONString+
    "Internal Error Message:\n" + e;
  }

  let tempDef = null;
  let missingFieldName = null;
  let errorString = null;
  for( let i=0; errorString === null && i < findAndReplaceJSON.length; i++ ) {
    tempDef = findAndReplaceJSON[i];
    if( tempDef.TOUPPER !== true && tempDef.TOUPPER !== false ) {
      missingFieldName = "TOUPPER";
    }
    else if( tempDef.FIND === null || tempDef.FIND === undefined ) {
      missingFieldName = "FIND";
    }
    else if( tempDef.REPLACE === null || tempDef.REPLACE === undefined ) {
      missingFieldName = "REPLACE";
    }
    if( missingFieldName != null ) {
      errorString = "The field " + missingFieldName + " on object number " + (i+1) + " of the find And replace JSON String is missing" +
                  "or invalid\n\nfindAndReplaceJSONString:\n" + findAndReplaceJSONString;
    }
  }
  return errorString;
}
/**
 * [utils-spreadsheet.gs]
 * parses findAndReplaceJSONString and uses it to run the 0, 1 or more defined search and replace operations on rangeVales2D
 * See findAndReplaceExactCellValuesInRange documentation for more info.
 * @param rangeVales2D {Array<Array<Object>>} 2d Array of values like that returned from Range.getValues() - to be manipluated
 * @param findAndReplaceJSONString {String} A String storing valid JSON used to define 0,1 or more F & R operations.
 * @return null
 */
function findAndReplaceExactCellValuesInRange_FromJSON( rangeVales2D, findAndReplaceJSONString ) {
  let findAndReplaceJSON = null;
  findAndReplaceJSON = JSON.parse( findAndReplaceJSONString );

  let tempDef = null;
  for( let i=0; i < findAndReplaceJSON.length; i++ ) {
    tempDef = findAndReplaceJSON[i];
    findAndReplaceExactCellValuesInRange( rangeVales2D, tempDef.TOUPPER, tempDef.FIND, tempDef.REPLACE );
  }
}
/**
 * [utils-spreadsheet.gs]
 * Basically the same as Edit Menu Find and Replace with 'Match Whole Cell Contents Only' ticked.
 * for case insensitive search set toUpper to true AND MAKE SURE THE findVal IS UPPER CASE TOO
 * Doesnt save back to spreadsheet just edits the values inside the array. As such pass in the 2d array returned from Range.getValues()
 * and it will return null but the 2d array you pass in will itself be edited
 * @param rangeVales2D {Array<Array<Object>>} 2d Array of values like that returned from Range.getValues()
 * @param toUpper {boolean} true if you want to touppercase the values before comparing with findVal (does not toUpper findval, findVal must be a String for this to work)
 * @param findVal {Object} - must match extire cell contents - types must match too (uses === not ==)
 * @param replaceVal {Object} - object of any valid Spreadsheet type to replace entire cell contents with.
 * @return null
 */
function findAndReplaceExactCellValuesInRange( rangeVales2D, toUpper, findVal, replaceVal ) {
  for( let yi=0; yi < rangeVales2D.length; yi++ ) {
    for( let xi=0; xi < rangeVales2D[yi].length; xi++ ) {
      if( ((toUpper&&typeof(rangeVales2D[yi][xi])=="string")?rangeVales2D[yi][xi].toUpperCase():rangeVales2D[yi][xi]) === findVal ) {
        rangeVales2D[yi][xi] = replaceVal;
      }
    }
  }
}


/**
 * Duplicate Sheet level protection between 2 sheets - works across 2 seperate spreadsheets. 
 * If sourceSheet has Sheet level protection then set up idential protection on the destSheet. 
 * DEFINED IN: utils-spreadsheets.gs
 *
 * @param {Sheet} sourceSheet The Sheet to copy from
 * @param {Sheet} destSheet The Sheet to copy to
 * @return null
 */
function DuplicateSheetLevelProtection( sourceSheet, destSheet ) {
  var sheetProtections = sourceSheet.getProtections(SpreadsheetApp.ProtectionType.SHEET);
  Logger.log( "DuplicateSheetLevelProtection, sheetProtections = " + sheetProtections );
  for( var i=0; i<sheetProtections.length; i++ ) {
    var p = sheetProtections[i];
    var p2 = destSheet.protect();
    p2.setDescription(p.getDescription());
    p2.setWarningOnly(p.isWarningOnly());  
    if (!p.isWarningOnly()) {
      p2.removeEditors(p2.getEditors());
      p2.addEditors(p.getEditors());
      // p2.setDomainEdit(p.canDomainEdit()); //  only if using an Apps domain 
    }
    var ranges = p.getUnprotectedRanges();
    var newRanges = [];
    for (var i = 0; i < ranges.length; i++) {
      newRanges.push(destSheet.getRange(ranges[i].getA1Notation()));
    } 
    p2.setUnprotectedRanges(newRanges);
  }
}
/**
 * [utils-spreadsheets.gs]
 * Duplicate Range level protections between 2 sheets - works across 2 seperate spreadsheets. 
 * If sourceSheet has Range level protection then set up idential protection on the destSheet. 
 *
 * @param {Sheet} sourceSheet The Sheet to copy from
 * @param {Sheet} destSheet The Sheet to copy to
 * @param {Array<string>} emailAddressesToAdd Array of email addresses to add access to for all protected ranges
 * @return null
 */
function DuplicateRangeLevelProtection( sourceSheet, destSheet ) {

  let protections = sourceSheet.getProtections(SpreadsheetApp.ProtectionType.RANGE);
  Logger.log( "DuplicateSheetLevelProtection, protections = " + protections );

  for (var i = 0; i < protections.length; i++) {
    let p = protections[i];
    let rangeNotation = p.getRange().getA1Notation();
    let p2 = destSheet.getRange(rangeNotation).protect();
    let pTargetAudienceIds = null;
    let p2TargetAudienceIds = null;
    p2.setDescription(p.getDescription());
    p2.setWarningOnly(p.isWarningOnly());
    if (!p.isWarningOnly()) {

      p2.removeEditors(p2.getEditors());
      p2.addEditors(p.getEditors());

      p2TargetAudienceIds = p2.getTargetAudiences();
      pTargetAudienceIds = p.getTargetAudiences();
      for( let ta2=0; ta2<p2TargetAudienceIds.length; ta2++ ) { p2.removeTargetAudience( p2TargetAudienceIds[ta2] ); }
      for( let ta=0; ta<pTargetAudienceIds.length; ta++ ) { p2.addTargetAudience( pTargetAudienceIds[ta] ); }

      // p2.setDomainEdit(p.canDomainEdit()); //  only if using an Apps domain 
   }
  }
}

function AuditSpreadSheetProtections() { 
  let spreadsheet = SpreadsheetApp.getActive();
  let ui = SpreadsheetApp.getUi();

  let accessLevelMappings = {
    dev: [["lsa.admin@wlc.ac.uk"],[]],
    admin: [["lsa.admin@wlc.ac.uk", "lsa-administrators@wlc.ac.uk"],[]],
  }

  let allProtections = spreadsheet.getProtections( SpreadsheetApp.ProtectionType.SHEET );
  allProtections = allProtections.concat( spreadsheet.getProtections( SpreadsheetApp.ProtectionType.RANGE ) );

  let thisProtection: GoogleAppsScript.Spreadsheet.Protection|any = null;
  let thisProtectionDef: ProtectionDef|null = null;
  let thisProtectionEditors: Array<GoogleAppsScript.Base.User>|null = null;
  let thisProtectionTargetAudiences: Array<GoogleAppsScript.Base.User>|null = null;
  for (var i = 0; i < allProtections.length; i++) {
    thisProtection = allProtections[i];
    Logger.log( "Checking: " + thisProtection.getDescription() );
    thisProtectionDef = getProtectionDefFromProtectionName( ui, thisProtection.getDescription() );
    if( thisProtectionDef != null ) { 

      if( ( thisProtectionDef.sheetNotRange && thisProtection.getProtectionType() != SpreadsheetApp.ProtectionType.SHEET )
       || ( !thisProtectionDef.sheetNotRange && thisProtection.getProtectionType() != SpreadsheetApp.ProtectionType.RANGE ) ) {
        if( ui.alert( "Protection Type Error", "Protection with name: '"+thisProtection.getDescription()+"'\n\nHas invalid" + 
                  "type, eg Sheet|Range", ui.ButtonSet.OK_CANCEL ) == ui.Button.CANCEL ) {
          return;          
        }
      }

      //check access levels are right
      else if( !accessLevelMappings[ thisProtectionDef.accessLevel ] ) {
        if( ui.alert( "Access Level Error", "Protection with name: '"+thisProtection.getDescription()+"'\n\nHas invalid" + 
                  "access level, eg Dev|Admin", ui.ButtonSet.OK_CANCEL ) == ui.Button.CANCEL ) {
          return;          
        }
      }
      else {
        thisProtectionEditors = thisProtection.getEditors();
        thisProtectionTargetAudiences = thisProtection.getTargetAudiences();
        Logger.log( thisProtectionTargetAudiences );
        let noOfEditorsMissingFromDef: number = 0;
        let noOfAudiencesMissingFromDef: number = 0;

        for( let e=0; e<thisProtectionEditors.length; e++ ) { 
          if( !accessLevelMappings[ thisProtectionDef.accessLevel ][0].includes( thisProtectionEditors[e].getEmail() ) )
          {
            ui.alert( "Protection Has Extra Users", "Protection with name: '"+thisProtection.getDescription()+"'\n\nHas a user:\n" + 
            "'"+thisProtectionEditors[e].getEmail()+"' which is not usual for access level: " + thisProtectionDef.accessLevel, 
            ui.ButtonSet.OK );
            noOfEditorsMissingFromDef++;
          }
        }
        for( let ta=0; ta<thisProtectionTargetAudiences.length; ta++ ) { 
          if( !accessLevelMappings[ thisProtectionDef.accessLevel ][1].includes( thisProtectionTargetAudiences[ta].getEmail() ) )
          {
            ui.alert( "Protection Has Extra Groups", "Protection with name: '"+thisProtection.getDescription()+"'\n\nHas a group user:\n" + 
            "'"+thisProtectionTargetAudiences[ta].getEmail()+"' which is not usual for access level: " + thisProtectionDef.accessLevel, 
            ui.ButtonSet.OK );
            noOfAudiencesMissingFromDef++;
          }
        }

        if( accessLevelMappings[ thisProtectionDef.accessLevel ][0].length > 
            ( thisProtectionEditors.length - noOfEditorsMissingFromDef ) ) {
          let btn = ui.alert( "Users Missing from Protection", "Protection with name: '"+thisProtection.getDescription()+"'\n\nshould have " + 
                accessLevelMappings[ thisProtectionDef.accessLevel ][0].length+" users but doesnt. It actually has " + thisProtectionEditors.length +
                " users, including " + noOfEditorsMissingFromDef + " extra unnecessary users\n\n" + 
                "Should Have: " + accessLevelMappings[ thisProtectionDef.accessLevel ][0] + "\n" +
                "Actually has: " + thisProtectionEditors + "\n\n" +
                "Do you want to add these extra users?", 
                ui.ButtonSet.YES_NO_CANCEL );
          if( btn == ui.Button.YES ) {
            thisProtection.removeEditors( accessLevelMappings[ thisProtectionDef.accessLevel ][0] );
            thisProtection.addEditors( accessLevelMappings[ thisProtectionDef.accessLevel ][0] );
          }
          else if( btn == ui.Button.CANCEL ){
            return;
          }
        }
        if( accessLevelMappings[ thisProtectionDef.accessLevel ][1].length > 
          ( thisProtectionTargetAudiences.length - noOfAudiencesMissingFromDef ) ) {
          let btn = ui.alert( "Group Users Missing from Protection", "Protection with name: '"+thisProtection.getDescription()+"'\n\nshould have " + 
                accessLevelMappings[ thisProtectionDef.accessLevel ][1].length+" group users but doesnt. It actually has " + thisProtectionTargetAudiences.length +
                " group users, including " + noOfAudiencesMissingFromDef + " extra unnecessary group users\n\n" + 
                "Should Have: " + accessLevelMappings[ thisProtectionDef.accessLevel ][1] + "\n" +
                "Actually has: " + thisProtectionTargetAudiences + "\n\n" +
                "Do you want to add these extra group users?", 
                ui.ButtonSet.YES_NO_CANCEL );
          if( btn == ui.Button.YES ) {       
            for( let ta = 0 ; ta < accessLevelMappings[ thisProtectionDef.accessLevel ][1].length ; ta++ ) {
              if( !thisProtectionTargetAudiences.includes( accessLevelMappings[ thisProtectionDef.accessLevel ][1][ ta ] ) ) {
                thisProtection.addTargetAudience( accessLevelMappings[ thisProtectionDef.accessLevel ][1][ ta ] );
              }
            }
          }
          else if( btn == ui.Button.CANCEL ){
            return;
          }
        }
      }
    }
  }
  ui.alert( "Done", "Finished, no more issues found", ui.ButtonSet.OK );
}

interface ProtectionDef { 
  sheetName: string, 
  sheetNotRange: boolean,
  accessLevel: string, 
  description: string
} 

function getProtectionDefFromProtectionName( ui: any, protectionName: string ):ProtectionDef  {

  //parse params
  ui = (ui) ? ui : SpreadsheetApp.getUi();

  let nameSplit = protectionName.split( "|" );
  
  if( nameSplit.length < 3)
  {
    ui.alert( "Protection Parse Error", "Could Not Parse Protection name: '"+protectionName+"\n\nIt should" + 
              "have 3 or 4 parts seperated by the | character", ui.ButtonSet.OK );
    return null;
  }

  if( nameSplit[1].toLowerCase() != "sheet" && nameSplit[1].toLowerCase() != "range" )
  {
    ui.alert( "Protection Parse Error", "Could Not Parse Protection name: '"+protectionName+"'\n\nSecond part should be " + 
              "'Sheet' or 'Range', but not '"+nameSplit[1]+"'", ui.ButtonSet.OK );
    return null;
  }

  let protectionDef: ProtectionDef = {
    sheetName: nameSplit[0],
    sheetNotRange: nameSplit[1].toLowerCase() == "sheet",
    accessLevel: nameSplit[2].toLowerCase(),
    description: nameSplit.splice( 0, 3 ).join("|")
  }
Logger.log( "protectionDef = " + protectionDef );
  return protectionDef;
}