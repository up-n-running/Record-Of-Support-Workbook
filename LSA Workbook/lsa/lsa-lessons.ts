function ShowAddLessonModal() {
  return ShowAddOrEditLessonModal( null, null );
}

function ShowEditLessonModal() {
  let spreadsheet = SpreadsheetApp.getActive();
  let ui = SpreadsheetApp.getUi();
  let lessonSheet = spreadsheet.getSheetByName( SHEETS.SETTINGS_LESSONS.NAME );

  let lessonObj = GetSelectedLesson( spreadsheet, ui, lessonSheet, "edit" );
  
  if( lessonObj.ROW_NO > 0 ) {
    return ShowAddOrEditLessonModal( lessonObj.ROW_NO, lessonObj.NAME );
  }
  return null;
}

function ShowAddOrEditLessonModal( rowNum, lessonName ) {
  
  //parse params
  rowNum = (rowNum) ? rowNum : -1;
  lessonName = (lessonName) ? lessonName : "";
  
  let ui = SpreadsheetApp.getUi();

  let alertHTMLTemplate = HtmlService.createTemplateFromFile('html/html-modal-lesson-add-edit');
  alertHTMLTemplate.rowNum = rowNum;
  alertHTMLTemplate.lessonName = lessonName;
  let alertMessageHTML = alertHTMLTemplate.evaluate().getContent();
  let alertMessage =  HtmlService.createHtmlOutput( alertMessageHTML ).setWidth(600).setHeight(280);
  ui.showModalDialog( alertMessage, ( ( rowNum < 0 ) ? "Add" : "Edit") + " a Lesson" );
}

function AddOrEditLesson( rowNum, lessonName ) {
  let spreadsheet = SpreadsheetApp.getActive();
  let ui = SpreadsheetApp.getUi();

  let lessonSheet = spreadsheet.getSheetByName( SHEETS.SETTINGS_LESSONS.NAME );

  let addingNotEditing = ( rowNum < 0 );
  if( addingNotEditing ) {
    rowNum = findInColumn( lessonSheet, "", SHEETS.SETTINGS_LESSONS.REFS.COL_NO_LESSON_NAME, 
                           SHEETS.SETTINGS_LESSONS.REFS.ROW_NO_FIRST_LESSON, SHEETS.SETTINGS_LESSONS.REFS.ROW_NO_LAST_LESSON );
  }

  lessonSheet.getRange( rowNum, SHEETS.SETTINGS_LESSONS.REFS.COL_NO_LESSON_NAME ).setValue( lessonName );
  SpreadsheetApp.flush();

  ui.alert( "Lesson Saved", "Lesson '" + lessonName + "' has been " + ( addingNotEditing ? "added" : "updated" ) + 
            " on row " + rowNum, ui.ButtonSet.OK );

  return rowNum;
}



function RemoveSelectedLesson() {
  let spreadsheet = SpreadsheetApp.getActive();
  let ui = SpreadsheetApp.getUi();
  let lessonSheet = spreadsheet.getSheetByName( SHEETS.SETTINGS_LESSONS.NAME );

  let lessonObj = GetSelectedLesson( spreadsheet, ui, lessonSheet, "delete" );
  
  if( lessonObj.ROW_NO > 0 ) {

    let sureResponse = ui.alert( "Are you sure?", "Are you sure you want to delete '" + lessonObj.NAME + "'?\n\n" + 
              "All of the Lesson's data, including its Target Grades and Long-Term Targets (from the other settings sheets)\n" + 
              "will also be deleted.\n\nPress OK to delete the Lesson.", ui.ButtonSet.OK_CANCEL );
    if( sureResponse == ui.Button.OK ) {
      //delete lesson record
      deleteCellsOnRowAndMoveUp_ColumnRange(lessonObj.ROW_NO, SHEETS.SETTINGS_LESSONS.REFS.COL_NO_EQUIPMENT_USED,
            SHEETS.SETTINGS_LESSONS.REFS.COL_NO_LESSON_NAME, lessonSheet, null );

      //delete target grade data column for lesson (complete column except first cell)
      let targetGradesSheet = spreadsheet.getSheetByName( SHEETS.SETTINGS_TARGET_GRADES.NAME );
      deleteCellsOnColumnAndMoveLeft_RowRange( lessonObj.ROW_NO - SHEETS.SETTINGS_LESSONS.REFS.ROW_NO_FIRST_LESSON + 
            SHEETS.SETTINGS_TARGET_GRADES.REFS.COL_NO_LEARNER_NAMES + 1,
            SHEETS.SETTINGS_TARGET_GRADES.REFS.ROW_NO_LESSON_NAMES + 1,
            SHEETS.SETTINGS_LEARNERS.REFS.ROW_NO_LAST_LEARNER - SHEETS.SETTINGS_LEARNERS.REFS.ROW_NO_FIRST_LEARNER +  
            SHEETS.SETTINGS_TARGET_GRADES.REFS.ROW_NO_LESSON_NAMES + 1,
            targetGradesSheet, null );

      //delete long term target data colonm for lesson (complete column except first cell)
      let lessonTargetsSheet = spreadsheet.getSheetByName( SHEETS.SETTINGS_LESSON_TARGETS.NAME );
      deleteCellsOnColumnAndMoveLeft_RowRange( lessonObj.ROW_NO - SHEETS.SETTINGS_LESSONS.REFS.ROW_NO_FIRST_LESSON + 
            SHEETS.SETTINGS_LESSON_TARGETS.REFS.COL_NO_LEARNER_NAMES + 1,
            SHEETS.SETTINGS_LESSON_TARGETS.REFS.ROW_NO_LESSON_NAMES + 1,
            SHEETS.SETTINGS_LEARNERS.REFS.ROW_NO_LAST_LEARNER - SHEETS.SETTINGS_LEARNERS.REFS.ROW_NO_FIRST_LEARNER +  
            SHEETS.SETTINGS_LESSON_TARGETS.REFS.ROW_NO_LESSON_NAMES + 1,
            lessonTargetsSheet, null );

      SpreadsheetApp.flush();
      ui.alert( "Lesson Deleted", "Lesson '"+lessonObj.NAME+"' has been successfully deleted, including its Target Grades and " + 
                "Long-Term Targets data.", ui.ButtonSet.OK );
    }



  }
  return null;
}


function GetSelectedLesson( spreadsheet, ui, lessonSheet, actionName ) {

  let helpText = "Please click somewhere on the row of the Lesson you would like to "+actionName+" and try again."

  var selectedRowNos = getAllSelectedRowOrColumnRecordNumbers(true, spreadsheet, ui, SHEETS.SETTINGS_LESSONS.NAME,
      SHEETS.SETTINGS_LESSONS.REFS.ROW_NO_FIRST_LESSON, SHEETS.SETTINGS_LESSONS.REFS.ROW_NO_LAST_LESSON, 
      SHEETS.SETTINGS_LESSONS.REFS.COL_NO_LESSON_NAME_READONLY, SHEETS.SETTINGS_LESSONS.REFS.COL_NO_SUPPORT_STRAT_LAST,
      helpText );

  if( selectedRowNos.length > 1 ) {
    ui.alert( "Please select just one", "Please select just one Lesson record at a time to "+actionName+"\n\n" +
              helpText, ui.ButtonSet.OK );
  }                        
  else if (selectedRowNos.length == 1 ) {
    let rowNo = selectedRowNos[0]+SHEETS.SETTINGS_LESSONS.REFS.ROW_NO_FIRST_LESSON-1;
Logger.log(rowNo);
    let lessonName = lessonSheet.getRange( rowNo, SHEETS.SETTINGS_LESSONS.REFS.COL_NO_LESSON_NAME ).getValue();
    if( lessonName == "" ) {
      ui.alert( "No Lesson Selected", "You selected a blank row, with no Lesson information on it.\n\n" + 
                helpText, ui.ButtonSet.OK );
    }
    else {
      return { ROW_NO: rowNo, NAME: lessonName }
    }
  }

  return { ROW_NO: -1, NAME: null };
}