function ShowAddLearnerModal() {
  //let spreadsheet = SpreadsheetApp.getActive();
  let ui = SpreadsheetApp.getUi();

  let alertHTMLTemplate = HtmlService.createTemplateFromFile('html/html-modal-learner-search');
  let alertMessageHTML = alertHTMLTemplate.evaluate().getContent();
  let alertMessage =  HtmlService.createHtmlOutput( alertMessageHTML ).setWidth(600).setHeight(420);
  ui.showModalDialog( alertMessage, 'Add a Learner' );
}

function ShowCreateShortTermLearnerModal() {
  ShowCreateOrEditLearnerModal( -1, null );
}

function ShowEditShortTermLearnerModal() {
  let spreadsheet = SpreadsheetApp.getActive();
  let ui = SpreadsheetApp.getUi();
  //let ui = {};
  let learnerSheet = spreadsheet.getSheetByName( SHEETS.SETTINGS_LEARNERS.NAME );

  let compoundLearnerObj = getSelectedLearner_( spreadsheet, ui, learnerSheet, "edit" );
  
  if ( compoundLearnerObj.LEARNER != null && compoundLearnerObj.LEARNER.CATEGORY == "Long Term" ) {
    ui.alert( "Not a Short Term Learner", "Sorry, You cannot edit '"+compoundLearnerObj.LEARNER.FORENAME+" "+compoundLearnerObj.LEARNER.SURNAME+"' as they are a Long Term learner and as such their information is controlled in the Master Database.\n\n" +
          "You can still edit the orange (or orangey-pink) cells for this learner but their grey cells are not editable.\n\n" +
          "If the data in one of this learner's grey cells is wrong then we please ask that you email your manager immediately to ask them to update it as this could cause problems for anyone else who supports this learner too.", ui.ButtonSet.OK );
    return;
  }

  if( compoundLearnerObj.ROW_NO > 0 ) {
    return ShowCreateOrEditLearnerModal( compoundLearnerObj.ROW_NO, compoundLearnerObj.LEARNER );
  }
  return null;
}

function ShowCreateOrEditLearnerModal( rowNum, learnerObj ) {
  let spreadsheet = SpreadsheetApp.getActive();
  let ui = SpreadsheetApp.getUi();

  //get email validation regex
  let settingsSheet = spreadsheet.getSheetByName( SHEETS.GLOBAL_SETTINGS.NAME );
  let emailValidationRegex = settingsSheet.getRange( 
        SHEETS.GLOBAL_SETTINGS.REFS.ROW_NO_VALID_EMAIL_REGEX, 
        SHEETS.GLOBAL_SETTINGS.REFS.COL_NO
      ).getValue();

  //Logger.log( "Calling Add/Edit Short Term Learner Modal, learner = \n" + debugLearner( learnerObj ) );
  let alertHTMLTemplate = HtmlService.createTemplateFromFile('html/html-modal-learner-create');
  alertHTMLTemplate.rowNum = rowNum;
  alertHTMLTemplate.learner = learnerObj;
  alertHTMLTemplate.emailValidationRegex = emailValidationRegex;
  let alertMessageHTML = alertHTMLTemplate.evaluate().getContent();
  let alertMessage =  HtmlService.createHtmlOutput( alertMessageHTML ).setWidth(600).setHeight(490);
  ui.showModalDialog( alertMessage, ( (rowNum < 0 ) ? 'Create' : 'Edit' ) + ' a Short-Term Learner' );
}

function MasterLearnerSearch( dataNo: string|null, forename: string|null, surname: string|null, learnerIdArray: Array<number>,
                              fetchAll: boolean, calledFromMaster: boolean ) {
  //parse params
  dataNo    = ( dataNo    ) ? dataNo.trim().toUpperCase()   : "";
  forename  = ( forename  ) ? forename.trim().toUpperCase() : "";
  surname   = ( surname   ) ? surname.trim().toUpperCase()  : "";
  learnerIdArray = ( learnerIdArray ) ? learnerIdArray      : null;

  let spreadsheet = SpreadsheetApp.getActive();

  //find master file's id
  let settingsSheet = spreadsheet.getSheetByName( SHEETS.GLOBAL_SETTINGS.NAME );
  let masterFileId = settingsSheet.getRange( 
        SHEETS.GLOBAL_SETTINGS.REFS.ROW_NO_MASTER_FILE_ID, 
        SHEETS.GLOBAL_SETTINGS.REFS.COL_NO 
      ).getValue();

  //get master file if this is not already the master
  let masterSpreadsheet=(masterFileId=="" || spreadsheet.getId()==masterFileId) ? spreadsheet : SpreadsheetApp.openById( masterFileId );

  Logger.log( "Fetching Named range MASTER_LEARNERS_LEARNER_ID from master" );
  let start = new Date();
  //GET 7 RANGES - ONE PER COLUMN - INCLUSING LOGIC FOR USING NAMES RANGES IFF THE SPREADSHEET IS NOT THIS ONE
  //Get the first column definition
  let colDefLearnerId = GetColDefByHandleFromAnyWorkbook( masterSpreadsheet, SHEETS.MASTER_LEARNERS.HANDLE, 
        SHEETS.MASTER_LEARNERS.REFS.HANDLE_LEARNER_ID, SHEETS.MASTER_LEARNERS.REFS.COL_NO_LEARNER_ID, 
        SHEETS.MASTER_LEARNERS.REFS.ROW_NO_FIRST_LEARNER, null, masterSpreadsheet.getSheetByName( SHEETS.MASTER_LEARNERS.NAME ),
        null, false, false );
  let masterLearnerSheet  = colDefLearnerId.sheet;
  let firstRowNumber      = colDefLearnerId.firstDataRowNo;
  let noOfLearnerRows     = colDefLearnerId.lastDataRowNo-firstRowNumber+1;

  Logger.log( "Done in " + ( new Date().getTime() - start.getTime() ) + "ms, now Fetching remaining Learner named ranges from master" ); start = new Date();
  //get the remianing column definitions - using info from the first one to avoid duplicating lookups
  let colDefForename = GetColDefByHandleFromAnyWorkbook( masterSpreadsheet, SHEETS.MASTER_LEARNERS.HANDLE, 
        SHEETS.MASTER_LEARNERS.REFS.HANDLE_FORENAME, SHEETS.MASTER_LEARNERS.REFS.COL_NO_FORENAME, 
        null, null, null, colDefLearnerId, calledFromMaster, false );
  let colDefNickname = GetColDefByHandleFromAnyWorkbook( masterSpreadsheet, SHEETS.MASTER_LEARNERS.HANDLE, 
        SHEETS.MASTER_LEARNERS.REFS.HANDLE_NICKNAME, SHEETS.MASTER_LEARNERS.REFS.COL_NO_NICKNAME, 
        null, null, null, colDefLearnerId, calledFromMaster, false );
  let colDefSurname = GetColDefByHandleFromAnyWorkbook( masterSpreadsheet, SHEETS.MASTER_LEARNERS.HANDLE, 
        SHEETS.MASTER_LEARNERS.REFS.HANDLE_SURNAME, SHEETS.MASTER_LEARNERS.REFS.COL_NO_SURNAME, 
        null, null, null, colDefLearnerId, calledFromMaster, false );
  let colDefCategory = GetColDefByHandleFromAnyWorkbook( masterSpreadsheet, SHEETS.MASTER_LEARNERS.HANDLE, 
        SHEETS.MASTER_LEARNERS.REFS.HANDLE_CATEGORY, SHEETS.MASTER_LEARNERS.REFS.COL_NO_CATEGORY, 
        null, null, null, colDefLearnerId, calledFromMaster, false );
  let colDefEmail = GetColDefByHandleFromAnyWorkbook( masterSpreadsheet, SHEETS.MASTER_LEARNERS.HANDLE, 
        SHEETS.MASTER_LEARNERS.REFS.HANDLE_EMAIL_ADDRESS, SHEETS.MASTER_LEARNERS.REFS.COL_NO_EMAIL_ADDRESS, 
        null, null, null, colDefLearnerId, calledFromMaster, false );
  let colDefSignType = GetColDefByHandleFromAnyWorkbook( masterSpreadsheet, SHEETS.MASTER_LEARNERS.HANDLE, 
        SHEETS.MASTER_LEARNERS.REFS.HANDLE_SIGN_TYPE, SHEETS.MASTER_LEARNERS.REFS.COL_NO_SIGN_TYPE, 
        null, null, null, colDefLearnerId, calledFromMaster, false );
  let colDefExternalId1 = GetColDefByHandleFromAnyWorkbook( masterSpreadsheet, SHEETS.MASTER_LEARNERS.HANDLE, 
        SHEETS.MASTER_LEARNERS.REFS.HANDLE_EXTERNAL_ID_1, SHEETS.MASTER_LEARNERS.REFS.COL_NO_EXTERNAL_ID_1, 
        null, null, null, colDefLearnerId, calledFromMaster, false );
  let colDefExternalId2 = GetColDefByHandleFromAnyWorkbook( masterSpreadsheet, SHEETS.MASTER_LEARNERS.HANDLE, 
        SHEETS.MASTER_LEARNERS.REFS.HANDLE_EXTERNAL_ID_2, SHEETS.MASTER_LEARNERS.REFS.COL_NO_EXTERNAL_ID_2, 
        null, null, null, colDefLearnerId, calledFromMaster, false );
  let colDeflearnerDir = GetColDefByHandleFromAnyWorkbook( masterSpreadsheet, SHEETS.MASTER_LEARNERS.HANDLE, 
        SHEETS.MASTER_LEARNERS.REFS.HANDLE_LEARNER_DIR, SHEETS.MASTER_LEARNERS.REFS.COL_NO_LEARNER_DIR, 
        null, null, null, colDefLearnerId, calledFromMaster, false );
  let colDefSignatureId = GetColDefByHandleFromAnyWorkbook( masterSpreadsheet, SHEETS.MASTER_LEARNERS.HANDLE, 
        SHEETS.MASTER_LEARNERS.REFS.HANDLE_SIGNATURE_FILE_ID, SHEETS.MASTER_LEARNERS.REFS.COL_NO_SIGNATURE_FILE_ID, 
        null, null, null, colDefLearnerId, calledFromMaster, false );

  Logger.log( "Done in " + ( new Date().getTime() - start.getTime() ) + "ms, now Reading Master Learner Data from master" ); start = new Date();
  let learnerIdRange   = GetRangeFromColDef( colDefLearnerId   ).getValues();
  let forenameRange    = GetRangeFromColDef( colDefForename    ).getValues();
  let nicknameRange    = GetRangeFromColDef( colDefNickname    ).getValues();
  let surnameRange     = GetRangeFromColDef( colDefSurname     ).getValues();
  let categoryRange    = GetRangeFromColDef( colDefCategory    ).getValues();
  let emailRange       = GetRangeFromColDef( colDefEmail       ).getValues();
  let signTypeRange    = GetRangeFromColDef( colDefSignType       ).getValues();
  let externalId1Range = GetRangeFromColDef( colDefExternalId1 ).getValues();
  let externalId2Range = GetRangeFromColDef( colDefExternalId2 ).getValues();
  let learnerDirRange  = GetRangeFromColDef( colDeflearnerDir  ).getValues();
  let signatureIdRange = GetRangeFromColDef( colDefSignatureId ).getValues();
  Logger.log( "Done in " + ( new Date().getTime() - start.getTime() ) + "ms" );

  //initialise an empty 2d array to store matching learnes by score from 1 to 10
  let learnersByScore = new Array();
  for( let i=1; i<=10; i++ ) {
    learnersByScore[i] = new Array();
  }

  //now loop through and get score for each learner
  let score;
  for( let li=0; li < noOfLearnerRows; li++ ) {
    score = 0;
    if( fetchAll && learnerIdRange[li][0] != "")
    {
      score = 10;
    }
    else if( learnerIdArray && learnerIdArray.includes( learnerIdRange[li][0] ) ) {
      score = 10;
    }
    else{
      if( dataNo != "" && learnerIdRange[li][0] == dataNo ) {
        score += 4;
      }
      if( forename != "" ) {
        if( forename == forenameRange[li][0].toUpperCase() || forename == nicknameRange[li][0].toUpperCase() ) {
          score += 3;
        }
        else if( forenameRange[li][0].toUpperCase().includes( forename ) || nicknameRange[li][0].toUpperCase().includes( forename ) ) {
          score += 1;
        }
      }
      if( surname != "" ) {
        if( surname == surnameRange[li][0].toUpperCase() ) {
          score += 3;
        }
        else if( surnameRange[li][0].toUpperCase().includes( surname ) ) {
          score += 1;
        }
      }
    }

    if( score > 0 ) {
      Logger.log( li );
      learnersByScore[ score ].push( {
        FORENAME      : forenameRange[li][0],
        NICKNAME      : nicknameRange[li][0],
        SURNAME       : surnameRange[li][0],
        LEARNER_ID    : learnerIdRange[li][0],
        CATEGORY      : categoryRange[li][0],
        EMAIL_ADDRESS : emailRange[li][0],
        SIGN_TYPE     : signTypeRange[li][0],
        EXTERNAL_ID_1 : externalId1Range[li][0],
        EXTERNAL_ID_2 : externalId2Range[li][0],
        LEARNER_DIR   : learnerDirRange[li][0],
        SIGNATURE_ID  : signatureIdRange[li][0]
      } );
    }
  }

  //create an array of matched learners with those with the highest scores first
  let matchedLearnersInOrder = new Array();
  for( let i=10; i>=1; i-- ) {
    matchedLearnersInOrder = matchedLearnersInOrder.concat( learnersByScore[i] );
  }

  //matched learners with those with the highest scores first
  return matchedLearnersInOrder;
}


function AddOrEditLearnerOnChildLearnerSheet( learnerObj, rowNum ) {
  let spreadsheet = SpreadsheetApp.getActive();
  let ui = SpreadsheetApp.getUi();
  let learnersSheet = spreadsheet.getSheetByName( SHEETS.SETTINGS_LEARNERS.NAME );
  let rowNumArray = ( rowNum && rowNum >= 0 ) ? [ rowNum ] : null;
Logger.log( "rowNumArray = " + rowNumArray );
  let feedbackText = AddOrEditLearnersOnChildLearnerSheet( spreadsheet, false, learnersSheet, [learnerObj], true, false, rowNumArray );
  SpreadsheetApp.flush();
  let alertTitle = "Learner Saved";
  if(feedbackText=="") {
    alertTitle = "Learner could not be saved";
    feedbackText = "Sorry, something went wrong, the Learner Information was not stored.";
  } 
  ui.alert( alertTitle, feedbackText, ui.ButtonSet.OK );
}


function AddOrEditLearnersOnChildLearnerSheet( spreadsheet, checkNamedRanges, learnersSheetIfKnown, learnerObjArr, addMissing, 
          onlyFeedbackErrors, rowNumsToUseIfNoMatch ) {

  //parse params
  learnersSheetIfKnown = (learnersSheetIfKnown ) ? learnersSheetIfKnown : 
        spreadsheet.getSheetByName( SHEETS.SETTINGS_LEARNERS.NAME );
  //OPTIONAL: used to force non matched records to save to specific row numbers
  //even if its over the top of an existing learner (useful when saving after user edits LearnerId)
  rowNumsToUseIfNoMatch = ( rowNumsToUseIfNoMatch ) ? rowNumsToUseIfNoMatch : null;

  let feedbackText = "";

  //GET 7 RANGES - ONE PER COLUMN - INCLUSING LOGIC FOR USING NAMES RANGES IFF THE SPREADSHEET IS NOT THIS ONE
  //Get the first column definition
  let colDefLearnerId = GetColDefByHandleFromAnyWorkbook( spreadsheet, SHEETS.SETTINGS_LEARNERS.HANDLE, 
        SHEETS.SETTINGS_LEARNERS.REFS.HANDLE_LEARNER_ID, SHEETS.SETTINGS_LEARNERS.REFS.COL_NO_LEARNER_ID, 
        SHEETS.SETTINGS_LEARNERS.REFS.ROW_NO_FIRST_LEARNER, null, learnersSheetIfKnown, null, !checkNamedRanges, 
        checkNamedRanges );
  let settingsLearnerSheet  = colDefLearnerId.sheet;
  let firstRowNumber      = colDefLearnerId.firstDataRowNo;
  let noOfLearnerRows     = colDefLearnerId.lastDataRowNo-firstRowNumber+1;

  //get the remianing column definitions - using info from the first one to avoid duplicating lookups
  let colDefForename = GetColDefByHandleFromAnyWorkbook( spreadsheet, SHEETS.SETTINGS_LEARNERS.HANDLE, 
        SHEETS.SETTINGS_LEARNERS.REFS.HANDLE_FORENAME, SHEETS.SETTINGS_LEARNERS.REFS.COL_NO_FORENAME, 
        null, null, null, colDefLearnerId, !checkNamedRanges, checkNamedRanges );
  let colDefNickname = GetColDefByHandleFromAnyWorkbook( spreadsheet, SHEETS.SETTINGS_LEARNERS.HANDLE, 
        SHEETS.SETTINGS_LEARNERS.REFS.HANDLE_NICKNAME, SHEETS.SETTINGS_LEARNERS.REFS.COL_NO_NICKNAME, 
        null, null, null, colDefLearnerId, !checkNamedRanges, checkNamedRanges );
  let colDefSurname = GetColDefByHandleFromAnyWorkbook( spreadsheet, SHEETS.SETTINGS_LEARNERS.HANDLE, 
        SHEETS.SETTINGS_LEARNERS.REFS.HANDLE_SURNAME, SHEETS.SETTINGS_LEARNERS.REFS.COL_NO_SURNAME, 
        null, null, null, colDefLearnerId, !checkNamedRanges, checkNamedRanges );
  let colDefCategory = GetColDefByHandleFromAnyWorkbook( spreadsheet, SHEETS.SETTINGS_LEARNERS.HANDLE, 
        SHEETS.SETTINGS_LEARNERS.REFS.HANDLE_CATEGORY, SHEETS.SETTINGS_LEARNERS.REFS.COL_NO_CATEGORY, 
        null, null, null, colDefLearnerId, !checkNamedRanges, checkNamedRanges );
  let colDefEmail = GetColDefByHandleFromAnyWorkbook( spreadsheet, SHEETS.SETTINGS_LEARNERS.HANDLE, 
        SHEETS.SETTINGS_LEARNERS.REFS.HANDLE_EMAIL_ADDRESS, SHEETS.SETTINGS_LEARNERS.REFS.COL_NO_EMAIL_ADDRESS, 
        null, null, null, colDefLearnerId, !checkNamedRanges, checkNamedRanges );
  let colDefSignType = GetColDefByHandleFromAnyWorkbook( spreadsheet, SHEETS.SETTINGS_LEARNERS.HANDLE, 
        SHEETS.SETTINGS_LEARNERS.REFS.HANDLE_SIGN_TYPE, SHEETS.SETTINGS_LEARNERS.REFS.COL_NO_SIGN_TYPE, 
        null, null, null, colDefLearnerId, !checkNamedRanges, checkNamedRanges );
  let colDefEditableNickname = GetColDefByHandleFromAnyWorkbook( spreadsheet, SHEETS.SETTINGS_LEARNERS.HANDLE, 
        SHEETS.SETTINGS_LEARNERS.REFS.HANDLE_EDITABLE_NICKNAME, SHEETS.SETTINGS_LEARNERS.REFS.COL_NO_EDITABLE_NICKNAME, 
        null, null, null, colDefLearnerId, !checkNamedRanges, checkNamedRanges );
  let colDefExternalId1 = GetColDefByHandleFromAnyWorkbook( spreadsheet, SHEETS.SETTINGS_LEARNERS.HANDLE, 
        SHEETS.SETTINGS_LEARNERS.REFS.HANDLE_EXTERNAL_ID_1, SHEETS.SETTINGS_LEARNERS.REFS.COL_NO_EXTERNAL_ID_1, 
        null, null, null, colDefLearnerId, !checkNamedRanges, checkNamedRanges );
  let colDefExternalId2 = GetColDefByHandleFromAnyWorkbook( spreadsheet, SHEETS.SETTINGS_LEARNERS.HANDLE, 
        SHEETS.SETTINGS_LEARNERS.REFS.HANDLE_EXTERNAL_ID_2, SHEETS.SETTINGS_LEARNERS.REFS.COL_NO_EXTERNAL_ID_2, 
        null, null, null, colDefLearnerId, !checkNamedRanges, checkNamedRanges );
  let colDefLearnerDir = GetColDefByHandleFromAnyWorkbook( spreadsheet, SHEETS.SETTINGS_LEARNERS.HANDLE, 
        SHEETS.SETTINGS_LEARNERS.REFS.HANDLE_LEARNER_DIR, SHEETS.SETTINGS_LEARNERS.REFS.COL_NO_LEARNER_DIR, 
        null, null, null, colDefLearnerId, !checkNamedRanges, checkNamedRanges );
  let colDefSignatureId = GetColDefByHandleFromAnyWorkbook( spreadsheet, SHEETS.SETTINGS_LEARNERS.HANDLE, 
    SHEETS.SETTINGS_LEARNERS.REFS.HANDLE_SIGNATURE_FILE_ID, SHEETS.SETTINGS_LEARNERS.REFS.COL_NO_SIGNATURE_FILE_ID, 
    null, null, null, colDefLearnerId, !checkNamedRanges, checkNamedRanges );

  //get the actual data in seperate columns
  let learnerIdRange   = GetRangeFromColDef( colDefLearnerId );
  let forenameRange    = GetRangeFromColDef( colDefForename  );
  let nicknameRange    = GetRangeFromColDef( colDefNickname  );
  let surnameRange     = GetRangeFromColDef( colDefSurname   );
  let categoryRange    = GetRangeFromColDef( colDefCategory  );
  let emailRange       = GetRangeFromColDef( colDefEmail     );
  let signTypeRange    = GetRangeFromColDef( colDefSignType    );
  let externalId1Range = GetRangeFromColDef( colDefExternalId1 );
  let externalId2Range = GetRangeFromColDef( colDefExternalId2 );
  let learnerDirRange  = GetRangeFromColDef( colDefLearnerDir  );
  let signatureIdRange = GetRangeFromColDef( colDefSignatureId );
  let editableNicknameRange = GetRangeFromColDef( colDefEditableNickname );

  let learnerIdValues   = learnerIdRange.getValues();
  let forenameValues    = forenameRange.getValues();
  let nicknameValues    = nicknameRange.getValues();
  let surnameValues     = surnameRange.getValues();
  let categoryValues    = categoryRange.getValues();
  let emailValues       = emailRange.getValues();
  let signTypeValues    = signTypeRange    ? signTypeRange.getValues()    : null;
  let externalId1Values = externalId1Range ? externalId1Range.getValues() : null;
  let externalId2Values = externalId2Range ? externalId2Range.getValues() : null;
  let learnerDirValues  = learnerDirRange  ? learnerDirRange.getValues()  : null;
  let signatureIdValues = signatureIdRange ? signatureIdRange.getValues() : null;
  let editableNicknameValues = editableNicknameRange.getValues();  

Logger.log( "B4 UPDATE");
Logger.log( "colDefLearnerId = " + debugColDef( colDefLearnerId ) );
Logger.log( "colDefSurname = " + debugColDef( colDefSurname ) );
Logger.log( "colDefSignType = " + debugColDef( colDefSignType ) );
Logger.log( "learnerIdValues = " + learnerIdValues );
Logger.log( "surnameValues = " + surnameValues );


  //Build a lookup table for learner object by learner id as a string
  let learnersByLearnerIdStrings = new Array();
  let learnerIdsBeingSaved = new Array();
  for( let i=0; i < learnerObjArr.length; i++ ) {
    learnersByLearnerIdStrings[ ""+ learnerObjArr[i].LEARNER_ID ] = learnerObjArr[i];
    learnerIdsBeingSaved.push( ""+learnerObjArr[i].LEARNER_ID );
  }

  //now loop through EVERY row of the learners sheet and where a learnerid matches. Update its values and maintain a list
  //of matched learner ids. If there is no match then do nothing, and if the row is blank then maintain a list of blank rows
  let tempLearnerId = null;
  let blankRecordIdxs = new Array();
  let learnerIdsMatched = new Array();
  let tempLearnerObj = null;
  for( let li=0; li < noOfLearnerRows; li++ ) {
    tempLearnerId = learnerIdValues[li][0];
    if( tempLearnerId=="" ) { 
      blankRecordIdxs.push( li ); 
    }
    else {
      learnerIdsMatched.push( ""+tempLearnerId );
      tempLearnerObj = learnersByLearnerIdStrings[ ""+ tempLearnerId ];
Logger.log( "Looking for learner: "+ tempLearnerId + " in master, found tempLearnerObj: " + tempLearnerObj );
      if( tempLearnerObj != null ) {
        forenameValues[li][0]    = tempLearnerObj.FORENAME;
        nicknameValues[li][0]    = tempLearnerObj.NICKNAME;
        surnameValues[li][0]     = tempLearnerObj.SURNAME;
      //learnerIdValues[li][0]   = tempLearnerObj.LEARNER_ID;
        categoryValues[li][0]    = tempLearnerObj.CATEGORY;
        emailValues[li][0]       = tempLearnerObj.EMAIL_ADDRESS;
        if( signTypeValues    ) { signTypeValues[li][0]    = tempLearnerObj.SIGN_TYPE; }  //null if not found in earlier version
        if( externalId1Values ) { externalId1Values[li][0] = tempLearnerObj.EXTERNAL_ID_1; }
        if( externalId2Values ) { externalId2Values[li][0] = tempLearnerObj.EXTERNAL_ID_2; }
        if( learnerDirValues  ) { learnerDirValues[li][0]  = tempLearnerObj.LEARNER_DIR; }
        if( signatureIdValues ) { signatureIdValues[li][0] = tempLearnerObj.SIGNATURE_ID; }
        if( editableNicknameValues[li][0] == "" ) {
          editableNicknameValues[li][0] =  tempLearnerObj.NICKNAME;
        }
        if(!onlyFeedbackErrors) { 
          feedbackText += ""+tempLearnerObj.FORENAME+" "+tempLearnerObj.SURNAME+" updated on row "+( firstRowNumber+li )+"\n";
        }
      }
    }
  }

  Logger.log( "AFTER STEP 1");
  Logger.log( "blankRecordIdxs = " + blankRecordIdxs );
  Logger.log( "learnerIdsMatched = " + learnerIdsMatched );
  Logger.log( "addMissing = " + addMissing );

  //now any non-matches above need adding (only if we have set the addMissing param to true)
  if( addMissing ) {
    Logger.log( "learnerIdsBeingSaved = " + learnerIdsBeingSaved );
    Logger.log( "learnerIdsMatched = " + learnerIdsMatched );
    let learnerIdsBeingAdded = removeElementsFromArray( learnerIdsBeingSaved, learnerIdsMatched );
    Logger.log( "learnerIdsBeingAdded = " + learnerIdsBeingAdded );
    let recordIdxToUse = -1;
    let updatedOrAdded = "added";
    for( let i=0; i < learnerIdsBeingAdded.length; i++ ) {
      tempLearnerObj = learnersByLearnerIdStrings[ ""+ learnerIdsBeingAdded[i] ];
      
      //if rowNumsToUseIfNoMatch list passed in then we'll save these non-matching learners to these row numbers (even if it means
      // saving over the top of existing rows (this is useful when we edit a Learner record and the user changes the learnerId -
      // here we use this to force the row num even when there is no learner id match))
      //but if its not passed in, or even if it is passed in but there's more learners to save than there are row numbers, 
      //then we'll revert to the blank rows array instead
      recordIdxToUse = -1;
      updatedOrAdded = "added";
Logger.log( "rowNumsToUseIfNoMatch = " + rowNumsToUseIfNoMatch );
      if( rowNumsToUseIfNoMatch != null && i < rowNumsToUseIfNoMatch.length ) {
        recordIdxToUse = rowNumsToUseIfNoMatch[i] - colDefLearnerId.firstDataRowNo;
        //as we're forcing the row number it might be an update not an add
        if( learnerIdValues[recordIdxToUse][0] != "" ) {
          updatedOrAdded = "updated";
        }
Logger.log( "recordIdxToUse = " + recordIdxToUse );
        //update the blank rows array so we know this row is not blank
        blankRecordIdxs = removeElementsFromArray( blankRecordIdxs, [recordIdxToUse] );
      }
      else {
        recordIdxToUse = blankRecordIdxs[ i-( rowNumsToUseIfNoMatch==null ? 0 : rowNumsToUseIfNoMatch.length ) ];
      }
      
      //check we havent run out of space on the sheet - if we have then we just stop adding more learners
      if( recordIdxToUse >= 0 ) {
        forenameValues[recordIdxToUse][0]    = tempLearnerObj.FORENAME;
        nicknameValues[recordIdxToUse][0]    = tempLearnerObj.NICKNAME;
        surnameValues[recordIdxToUse][0]     = tempLearnerObj.SURNAME;
        learnerIdValues[recordIdxToUse][0]   = tempLearnerObj.LEARNER_ID;
        categoryValues[recordIdxToUse][0]    = tempLearnerObj.CATEGORY;
        emailValues[recordIdxToUse][0]       = tempLearnerObj.EMAIL_ADDRESS;
        if( signTypeValues    ) { signTypeValues[recordIdxToUse][0]    = tempLearnerObj.SIGN_TYPE; }  //if not found in earlier version
        if( externalId1Values ) { externalId1Values[recordIdxToUse][0] = tempLearnerObj.EXTERNAL_ID_1; }
        if( externalId2Values ) { externalId2Values[recordIdxToUse][0] = tempLearnerObj.EXTERNAL_ID_2; }
        if( learnerDirValues )  { learnerDirValues[recordIdxToUse][0]  = tempLearnerObj.LEARNER_DIR; }
        if( signatureIdValues ) { signatureIdValues[recordIdxToUse][0] = tempLearnerObj.SIGNATURE_ID; }
        editableNicknameValues[recordIdxToUse][0] = tempLearnerObj.NICKNAME;
        if(!onlyFeedbackErrors) { 
          feedbackText+=tempLearnerObj.FORENAME+" "+tempLearnerObj.SURNAME+" "+updatedOrAdded+" on row "+ 
          ( firstRowNumber+recordIdxToUse )+"\n";
        }
      }
      else {
        feedbackText+=tempLearnerObj.FORENAME+" "+tempLearnerObj.SURNAME+" COULD NOT BE ADDED\n"+
              "There are no more blank rows on the Sheet, sorry.";
      }
    }
  }

  Logger.log( "AFTER UPDATE");
  Logger.log( "colDefLearnerId = " + debugColDef( colDefLearnerId ) );
  Logger.log( "colDefSurname = " + debugColDef( colDefSurname ) );
  Logger.log( "colDefSignType = " + debugColDef( colDefSignType ) );
  Logger.log( "learnerIdValues = " + learnerIdValues );
  Logger.log( "surnameValues = " + surnameValues );

  //now save all data
  learnerIdRange.setValues(learnerIdValues);
  forenameRange.setValues(forenameValues);
  nicknameRange.setValues(nicknameValues);
  surnameRange.setValues(surnameValues);
  categoryRange.setValues(categoryValues);
  emailRange.setValues(emailValues);
  if( signTypeRange    ) { 
    signTypeRange.setValues(signTypeValues); 
  }
  if( externalId1Range ) { 
    externalId1Range.setValues(externalId1Values); 
  }
  if( externalId2Range ) { 
    externalId2Range.setValues(externalId2Values); 
  }
  if( learnerDirRange  ) { 
    learnerDirRange.setValues(learnerDirValues); 
  }
  if( signatureIdRange ) { 
    signatureIdRange.setValues(signatureIdValues); 
  }
  editableNicknameRange.setValues(editableNicknameValues); 

  return feedbackText;
}


function RemoveSelectedLearner() {
  var spreadsheet = SpreadsheetApp.getActive();
  var ui = SpreadsheetApp.getUi();

  var selectedRowNos = getAllSelectedRowOrColumnRecordNumbers(true, spreadsheet, ui, SHEETS.SETTINGS_LEARNERS.NAME,
      SHEETS.SETTINGS_LEARNERS.REFS.ROW_NO_FIRST_LEARNER, SHEETS.SETTINGS_LEARNERS.REFS.ROW_NO_LAST_LEARNER, 
      SHEETS.SETTINGS_LEARNERS.REFS.COL_NO_LEARNER_NAME, SHEETS.SETTINGS_LEARNERS.REFS.COL_NO_EXTRA_SUPPORT_TEXT,
      "Please select a cell on one of the Learner's rows below and try again." );

Logger.log(selectedRowNos);

  //the above function ensures we're on the SETTINGS_LEARNERS worksheet
  var learnerSheet = spreadsheet.getActiveSheet();

  if( selectedRowNos.length > 1 ) {
    ui.alert( "Please select just one", "Please select just one Learner record at a time to delete\n\nPlease click on the name of " + 
              "the one learner you would like to delete in the first column and try again.", ui.ButtonSet.OK );
  }                        
  else if (selectedRowNos.length == 1 ) {
    let rowNo = selectedRowNos[0]+SHEETS.SETTINGS_LEARNERS.REFS.ROW_NO_FIRST_LEARNER-1;
Logger.log(rowNo);
    let learnerId = learnerSheet.getRange( rowNo, SHEETS.SETTINGS_LEARNERS.REFS.COL_NO_LEARNER_NAME ).getValue();
    if( learnerId == "" ) {
      ui.alert( "No Learner Selected", "You selected a blank row, with no Learner information on it.\n\nPlease click on the name of " + 
                "the learner you would like to delete in the first column and try again.", ui.ButtonSet.OK );
    }
    else {
      let learnerName = learnerSheet.getRange( rowNo, SHEETS.SETTINGS_LEARNERS.REFS.COL_NO_LEARNER_NAME ).getValue();
      let sureResponse = ui.alert( "Are you sure?", "Are you sure you want to delete '" + learnerName + "'?\n\n" + 
                "All of their data, including their Target Grades and Long-Term Targets (from the other settings sheets)\n" + 
                "will also be deleted. Their Records of Support will remain however!\n\nPress OK to delete their data.", ui.ButtonSet.OK_CANCEL );
      if( sureResponse == ui.Button.OK ) {
        //delete learner record
        deleteCellsOnRowAndMoveUp_ColumnRange(rowNo, SHEETS.SETTINGS_LEARNERS.REFS.COL_NO_EDITABLE_NICKNAME,
              SHEETS.SETTINGS_LEARNERS.REFS.COL_NO_SIGNATURE_FILE_ID, learnerSheet, null );

        //delete target grade data row for learner (complete row except first cell)
        let targetGradesSheet = spreadsheet.getSheetByName( SHEETS.SETTINGS_TARGET_GRADES.NAME );
        deleteCellsOnRowAndMoveUp_ColumnRange( rowNo - SHEETS.SETTINGS_LEARNERS.REFS.ROW_NO_FIRST_LEARNER + 
                    SHEETS.SETTINGS_TARGET_GRADES.REFS.ROW_NO_LESSON_NAMES + 1,
              SHEETS.SETTINGS_TARGET_GRADES.REFS.COL_NO_LEARNER_NAMES + 1,
              SHEETS.SETTINGS_LESSONS.REFS.ROW_NO_LAST_LESSON - SHEETS.SETTINGS_LESSONS.REFS.ROW_NO_FIRST_LESSON +  
                    SHEETS.SETTINGS_TARGET_GRADES.REFS.COL_NO_LEARNER_NAMES + 1, 
              targetGradesSheet, null );

        //delete long term target data row for learner (complete row except first cell)
        let lessonTargetsSheet = spreadsheet.getSheetByName( SHEETS.SETTINGS_LESSON_TARGETS.NAME );
        deleteCellsOnRowAndMoveUp_ColumnRange( rowNo - SHEETS.SETTINGS_LEARNERS.REFS.ROW_NO_FIRST_LEARNER + 
                    SHEETS.SETTINGS_LESSON_TARGETS.REFS.ROW_NO_LESSON_NAMES + 1,
              SHEETS.SETTINGS_LESSON_TARGETS.REFS.COL_NO_LEARNER_NAMES + 1,
              SHEETS.SETTINGS_LESSONS.REFS.ROW_NO_LAST_LESSON - SHEETS.SETTINGS_LESSONS.REFS.ROW_NO_FIRST_LESSON +  
                    SHEETS.SETTINGS_LESSON_TARGETS.REFS.COL_NO_LEARNER_NAMES + 1, 
              lessonTargetsSheet, null );

        SpreadsheetApp.flush();
        ui.alert( "Learner Deleted", "Learner '"+learnerName+"' has been successfully deleted, including their Target Grades and " + 
                  "Long-Term Targets data.", ui.ButtonSet.OK );
      }
    }
  }
}

/**
 * [lsa-learners.gs]
 * Pass in ONE OF learnerId and rowNumber. If you pass in both, then learnerId will be ignored
 * returns null if and only if rowNumber not passed in and learner not found by id
 */
function getChildLearnerObjByLearnerIdFromSameVersionSource_( spreadsheet, learnersSheet, learnerId, rowNumber ) {

  spreadsheet = ( spreadsheet ) ? spreadsheet : SpreadsheetApp.getActive();
  learnersSheet = ( learnersSheet ) ? learnersSheet : spreadsheet.getSheetByName( SHEETS.SETTINGS_LEARNERS.NAME );
  learnerId = ( learnerId ) ? learnerId : null;
  rowNumber = ( rowNumber ) ? rowNumber : null;

  //find learner row
  let learnerRowNo = -1;
  if( rowNumber ) {
    learnerRowNo = rowNumber;
  }
  else { 
    learnerRowNo = findInColumn( learnersSheet, learnerId, SHEETS.SETTINGS_LEARNERS.REFS.COL_NO_LEARNER_ID, 
        SHEETS.SETTINGS_LEARNERS.REFS.ROW_NO_FIRST_LEARNER, SHEETS.SETTINGS_LEARNERS.REFS.ROW_NO_LAST_LEARNER );
  }

  //get learner data from the rowNum
  if( learnerRowNo < 0 ) {
    return null;
  }
  let firstColNo = SHEETS.SETTINGS_LEARNERS.REFS.COL_NO_EDITABLE_NICKNAME;
  let lastColNo = SHEETS.SETTINGS_LEARNERS.REFS.COL_NO_SIGNATURE_FILE_ID;
  let learnerDataRaw = learnersSheet.getRange( learnerRowNo, firstColNo, 1, lastColNo - firstColNo + 1 ).getValues();
  
  //get the values from the checkboxes and save for later
  let supportStrategyValuesArray = new Array();
  for( let ss= SHEETS.SETTINGS_LEARNERS.REFS.COL_NO_SUPPORT_STRAT_FIRST; 
       ss <= SHEETS.SETTINGS_LEARNERS.REFS.COL_NO_SUPPORT_STRAT_LAST ; ss++ ) {
    supportStrategyValuesArray.push( learnerDataRaw[0][ ss - firstColNo ] );
  }

  return {
    FORENAME      : learnerDataRaw[0][ SHEETS.SETTINGS_LEARNERS.REFS.COL_NO_FORENAME - firstColNo ],
    NICKNAME      : learnerDataRaw[0][ SHEETS.SETTINGS_LEARNERS.REFS.COL_NO_EDITABLE_NICKNAME - firstColNo ],
    SURNAME       : learnerDataRaw[0][ SHEETS.SETTINGS_LEARNERS.REFS.COL_NO_SURNAME - firstColNo ],
    SUPPORT_NEED  : learnerDataRaw[0][ SHEETS.SETTINGS_LEARNERS.REFS.COL_NO_SUPPORT_NEED - firstColNo ],
    SUPPORT_STRAT_DEFAULTS: supportStrategyValuesArray,
    LEARNER_ID    : learnerDataRaw[0][ SHEETS.SETTINGS_LEARNERS.REFS.COL_NO_LEARNER_ID - firstColNo ],
    CATEGORY      : learnerDataRaw[0][ SHEETS.SETTINGS_LEARNERS.REFS.COL_NO_CATEGORY - firstColNo ],
    EMAIL_ADDRESS : learnerDataRaw[0][ SHEETS.SETTINGS_LEARNERS.REFS.COL_NO_EMAIL_ADDRESS - firstColNo ],
    SIGN_TYPE     : learnerDataRaw[0][ SHEETS.SETTINGS_LEARNERS.REFS.COL_NO_SIGN_TYPE - firstColNo ],
    EXTERNAL_ID_1 : learnerDataRaw[0][ SHEETS.SETTINGS_LEARNERS.REFS.COL_NO_EXTERNAL_ID_1 - firstColNo ],
    EXTERNAL_ID_2 : learnerDataRaw[0][ SHEETS.SETTINGS_LEARNERS.REFS.COL_NO_EXTERNAL_ID_2 - firstColNo ],
    LEARNER_DIR   : learnerDataRaw[0][ SHEETS.SETTINGS_LEARNERS.REFS.COL_NO_LEARNER_DIR - firstColNo ],
    SIGNATURE_ID  : learnerDataRaw[0][ SHEETS.SETTINGS_LEARNERS.REFS.COL_NO_SIGNATURE_FILE_ID - firstColNo ]
  };
}

function getSelectedLearner_( spreadsheet, ui, learnerSheet, actionName ) {

  let helpText = "Please click somewhere on the row of the Learner you would like to "+actionName+" and try again."

  let selectedRowNos = getAllSelectedRowOrColumnRecordNumbers(true, spreadsheet, ui, SHEETS.SETTINGS_LEARNERS.NAME,
      SHEETS.SETTINGS_LEARNERS.REFS.ROW_NO_FIRST_LEARNER, SHEETS.SETTINGS_LEARNERS.REFS.ROW_NO_LAST_LEARNER, 
      SHEETS.SETTINGS_LEARNERS.REFS.COL_NO_LEARNER_NAME, SHEETS.SETTINGS_LEARNERS.REFS.COL_NO_EXTRA_SUPPORT_TEXT,
      helpText );

  if( selectedRowNos.length > 1 ) {
    ui.alert( "Please select just one", "Please select just one Learner record at a time to "+actionName+"\n\n" +
              helpText, ui.ButtonSet.OK );
  }                        
  else if (selectedRowNos.length == 1 ) {
    let rowNo = selectedRowNos[0]+SHEETS.SETTINGS_LEARNERS.REFS.ROW_NO_FIRST_LEARNER-1;
Logger.log(rowNo);
    let learner = getChildLearnerObjByLearnerIdFromSameVersionSource_( spreadsheet, learnerSheet, null, rowNo );
    if( learner == null || learner.LEARNER_ID == "" ) {
      ui.alert( "No Learner Selected", "You selected a blank row, with no Learner information on it.\n\n" + 
                helpText, ui.ButtonSet.OK );
    }
    else {
      return { ROW_NO: rowNo, LEARNER: learner }
    }
  }

  return { ROW_NO: -1, LEARNER: null };
}

function RefreshSettingsLearnerSheetDataFromMasterDatabaseFromChild( spreadsheet?, ui?, settingsLearnersSheet?, alertOnSuccess?: boolean ) {
  
  //parse params
  spreadsheet = ( spreadsheet ) ? spreadsheet : SpreadsheetApp.getActive();
  ui = ( ui ) ? ui : SpreadsheetApp.getUi();
  settingsLearnersSheet = ( settingsLearnersSheet ) ? settingsLearnersSheet : spreadsheet.getSheetByName( SHEETS.SETTINGS_LEARNERS.NAME );
  alertOnSuccess = ( alertOnSuccess !== undefined ) ? alertOnSuccess : true;

  //get array of all learner numbers to get data for
  let learnerIdColumnValues = settingsLearnersSheet.getRange(
        SHEETS.SETTINGS_LEARNERS.REFS.ROW_NO_FIRST_LEARNER, SHEETS.SETTINGS_LEARNERS.REFS.COL_NO_LEARNER_ID,
        SHEETS.SETTINGS_LEARNERS.REFS.ROW_NO_LAST_LEARNER - SHEETS.SETTINGS_LEARNERS.REFS.ROW_NO_FIRST_LEARNER + 1, 1
  ).getValues();

  let allLearnerIdArray: Array<number> = new Array();
  let tempLearnerIdValue: number|string = -1;
  for( let li = 0 ; li < learnerIdColumnValues.length; li++ ) {
    tempLearnerIdValue = learnerIdColumnValues[ li ][ 0 ];
Logger.log( tempLearnerIdValue );
    if( tempLearnerIdValue && tempLearnerIdValue != "" && tempLearnerIdValue > 0 ) {
      allLearnerIdArray.push( parseInt( ""+tempLearnerIdValue, 10 ) );
    }
  }
Logger.log( allLearnerIdArray );

  let learnerObjectArrayFromMaster = MasterLearnerSearch( null, null, null, allLearnerIdArray, false, false );
Logger.log( learnerObjectArrayFromMaster );

  let errorFeedbackText = AddOrEditLearnersOnChildLearnerSheet( spreadsheet, false, settingsLearnersSheet, 
          learnerObjectArrayFromMaster, false, true, null );

  SpreadsheetApp.flush();

  if( errorFeedbackText ) {
    ui.alert( "Refresh Failed", "There was an issue with refreshing your Learners' information from the Master Database\n\n" +
              "Error Info:\n" + errorFeedbackText, ui.ButtonSet.OK );
    return false;
  }
  else if( alertOnSuccess ) {

    ui.alert( "Refresh Complete", "Your Learners' information has been refreshed successfully.",
              ui.ButtonSet.OK );
  }
  return true;
}

function debugLearner( learner ) {
    return ( !learner ) ? 'NULL' : 
    "FORENAME      : '" + learner.FORENAME + "'\n" +
    "NICKNAME      : '" + learner.NICKNAME + "'\n" +
    "SURNAME       : '" + learner.SURNAME + "'\n" +
    "SUPPORT_NEED  : '" + learner.SUPPORT_NEED + "'\n" +
    "SUPPORT_STRAT_DEFAULTS: '" + learner.SUPPORT_STRAT_DEFAULTS + "'\n" +
    "LEARNER_ID    : '" + learner.LEARNER_ID + "'\n" +
    "CATEGORY      : '" + learner.CATEGORY + "'\n" +
    "EMAIL_ADDRESS : '" + learner.EMAIL_ADDRESS + "'\n" +
    "SIGN_TYPE     : '" + learner.SIGN_TYPE + "'\n" +
    "EXTERNAL_ID_1 : '" + learner.EXTERNAL_ID_1 + "'\n" +
    "EXTERNAL_ID_2 : '" + learner.EXTERNAL_ID_2 + "'\n" +
    "LEARNER_DIR   : '" + learner.LEARNER_DIR + "'\n" +
    "SIGNATURE_ID  : '" + learner.SIGNATURE_ID + "'\n";
}