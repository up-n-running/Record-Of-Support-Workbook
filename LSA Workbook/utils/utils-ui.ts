

function getSpoofUIObject_FromAlertSheetDef( sheet: GoogleAppsScript.Spreadsheet.Sheet, alertSheetDef: any ) {
  return getSpoofUIObject_( sheet, alertSheetDef.TOP_ROW, alertSheetDef.B1_COL, alertSheetDef.B2_COL,
                            alertSheetDef.HIDDEN_COL, 1 );
}

var that = this;

function getSpoofUIObject_( sheet: GoogleAppsScript.Spreadsheet.Sheet, topRowNo: number, btn1ColNo: number,
                            btn2ColNo: number, hiddenColNo: number, fakeMode: number ) {
  return {
    ButtonSet: {
      OK: 1,
      OK_CANCEL: 3,
      YES_NO: 12,
      YES_NO_CANCEL: 14
    },
    Button: {
      OK: 1,
      CANCEL: 2,
      YES: 4,
      NO: 8,
      CLOSE: 16
    },
    getButtonText: function( buttonCode ) {  
      return ( buttonCode == this.Button.OK ) ? "Okay" : ( ( buttonCode == this.Button.CANCEL ) ? "Cancel" : ( 
               ( buttonCode == this.Button.YES ) ? "Yes" : ( ( buttonCode == this.Button.NO ) ? "No" : ( 
                 ( buttonCode == this.Button.CLOSE ) ? "Close" : "" ) ) ) );
    },
    alert: function( title: string, message: string, buttonsSet: number, 
          asyncModeCallbackFunction?: string|null, timerSeconds?: number|null ) { 
      if( this.FAKE_MODE == this.FAKE_MODES.MOBILE ) {

        //parse params
        asyncModeCallbackFunction = (asyncModeCallbackFunction) ? asyncModeCallbackFunction : null;
        timerSeconds = ( timerSeconds ) ? timerSeconds : 30;

        //get button mappings
        let cbx1ButtonMap = ( buttonsSet==this.ButtonSet.OK_CANCEL ) ? this.Button.OK : 
                              ( ( buttonsSet==this.ButtonSet.YES_NO || buttonsSet==this.ButtonSet.YES_NO_CANCEL ) ?
                                 this.Button.YES : -1 );
        let cbx2ButtonMap = ( buttonsSet==this.ButtonSet.OK ) ? this.Button.OK : 
                              ( ( buttonsSet==this.ButtonSet.YES_NO || buttonsSet==this.ButtonSet.YES_NO_CANCEL ) ?
                                 this.Button.NO : ( ( buttonsSet==this.ButtonSet.OK_CANCEL ) ? this.Button.CANCEL: -1 ) );

        let alertMode = ( cbx1ButtonMap >= 0 || cbx2ButtonMap >= 0 );

        //setup box
        sheet.getRange( topRowNo, hiddenColNo, 4, 1 ).setValues(
          [[title],[message], [buttonsSet],[
                ( ( alertMode && asyncModeCallbackFunction ) ? 
                  asyncModeCallbackFunction : ( ( alertMode ) ? timerSeconds : "" ) )
          ]]
        );
        sheet.getRange( topRowNo+2, btn1ColNo ).setValue(false);
        sheet.getRange( topRowNo+2, btn2ColNo ).setValue(false);

        if( alertMode && !asyncModeCallbackFunction ) {
          //listen for checkbox button being ticked if there are any buttons
          let cbxData = [[false, "", false]];
          let noOfLoops = 0;
          let startTime = new Date().getTime(); 
          while( noOfLoops < (timerSeconds+3) && !cbxData[0][0] && !cbxData[0][btn2ColNo-btn1ColNo] ) {
            Logger.log( "Sleeping, cbxData = " + cbxData );
            Utilities.sleep( Math.max( 0, ( 1000 * noOfLoops ) - ( new Date().getTime() - startTime  ) ) );
            //cbxData = sheet.getRange( rowNo+2, colNo+1, 1, 3 ).getValues();
            cbxData = alert_oninterval( topRowNo, btn1ColNo, btn2ColNo, hiddenColNo, timerSeconds - noOfLoops );
            Logger.log( "Finished Sleeping, cbxData[0][0] = " + cbxData[0][0] + ", cbxData[0][2] = " + cbxData[0][2] + ", noOfLoops = " + noOfLoops );
            noOfLoops++;
          }
          Logger.log( "FINISHED" );

          let btn1Ticked = cbxData[0][0];
          let btn2Ticked = cbxData[0][btn2ColNo-btn1ColNo];

          //clear down box
          sheet.getRange( topRowNo, hiddenColNo, 4, 1 ).setValues(
            [[""],[""], [""], [""]]
          );
          sheet.getRange( topRowNo+2, btn1ColNo ).setValue(false);
          sheet.getRange( topRowNo+2, btn2ColNo ).setValue(false);

          //return button press
          return buttonsSet < 0 ? null : ( btn1Ticked ? cbx1ButtonMap : cbx2ButtonMap );          
        }

        return -1; //toast mode not alert mode so no button pressed
      }
    },
    alert_async_finished: function( buttonNumber, callBackFunctionName: string ) { 
      //clear down box
      sheet.getRange( topRowNo, hiddenColNo, 4, 1 ).setValues(
        [[""],[""], [""], [""]]
      );
      sheet.getRange( topRowNo+2, btn1ColNo ).setValue(false);
      sheet.getRange( topRowNo+2, btn2ColNo ).setValue(false);

      that[ callBackFunctionName ]( this );
    },
    toast: function( message: string, title: string, timeout?: number|null|undefined ) {
      this.alert( title, message, -1 );
    },
    FAKE_MODES: {
      MOBILE: 1,
      NO_UI: 2
    },
    FAKE_MODE: fakeMode,
    SHEET: sheet
  }
}

function alert_oninterval( topRowNo, btn1ColNo, btn2ColNo, hiddenColNo, countdown ) {
  //SpreadsheetApp.flush();
  Logger.log( "alert_oninterval, hiddenColNo = " + hiddenColNo + ", rowNo = " + topRowNo );
  let timeString = ( countdown > 0 ) ? ""+countdown : "";
  SpreadsheetApp.getActive().getSheetByName( SHEETS.MOBILE_MAIN.NAME ).getRange( topRowNo+3, hiddenColNo ).setValue( timeString );
  SpreadsheetApp.flush();
  return SpreadsheetApp.getActive().getSheetByName( SHEETS.MOBILE_MAIN.NAME ).getRange( 
        topRowNo+2, btn1ColNo, 1, ( btn2ColNo - btn1ColNo + 1 )
  ).getValues();
}

function uiSensitiveToast( spreadsheet: GoogleAppsScript.Spreadsheet.Spreadsheet , 
                           ui: any, 
                           message: string,
                           title: string,
                           timeout?: number ) {
  
  Logger.log( "uiSensitiveToast called, ui = " + ui + ", ui.FAKE_MODE = " + ui.FAKE_MODE );

  if( !ui || !ui.FAKE_MODE ) {
    spreadsheet = ( spreadsheet ) ? spreadsheet : SpreadsheetApp.getActive();
    if( timeout ) {
      spreadsheet.toast( message, title, timeout );
    }
    else {
      spreadsheet.toast( message, title );
    }
  }
  else {
    ui.alert( title, message, -1 );
  }

}

function onEditCheck_FromAlertSheetDef( sheet: GoogleAppsScript.Spreadsheet.Sheet, alertSheetDef: any, event: any ) {
Logger.log( "onEditCheck_FromAlertSheetDef called" );
  //called from onEdit if the sheet matches the sheet with this alert on it
  let range = event.range;
  const row   = range.getRow(); 
  const col   = range.getColumn();
  const isSingleCell = ( ( range.getHeight() + range.getWidth() ) == 2 );

  if( isSingleCell && row == ( alertSheetDef.TOP_ROW + 2 ) && 
      ( col == alertSheetDef.B1_COL || col == alertSheetDef.B2_COL ) && 
      sheet.getName() === event.range.getSheet().getName() ) {
Logger.log( "onEditCheck_FromAlertSheetDef: BUTTON CHECKBOX TICKED" );
    let functionNameOrCountdownTimerValue = sheet.getRange( alertSheetDef.TOP_ROW + 3, alertSheetDef.HIDDEN_COL ).getValue();

    if( !isValidInteger( functionNameOrCountdownTimerValue ) )
    {
      getSpoofUIObject_FromAlertSheetDef( sheet, alertSheetDef ).alert_async_finished( 
            ( ( col == alertSheetDef.B1_COL ) ? 1 : 2 ), functionNameOrCountdownTimerValue );
    }

  }
}