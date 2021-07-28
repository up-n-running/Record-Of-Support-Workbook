function renewCheckboxes( sheetEditedName, event ) {
  //check if its a checkbox cell being updated as its ridiculously easy to delete checkboxes by accident in
  //google sheets at present
  if( CHECKBOX_RANGES[ sheetEditedName ] ) {
    let range = event.range;
    const rTop: number    = range.getRow();
    const rLeft: number   = range.getColumn();
    const rBottom: number = rTop + range.getHeight() - 1;
    const rRight: number  = rLeft + range.getWidth() - 1;

Logger.log( "CHECKING FOR MISSING CHECKBOXES" );
Logger.log( "event.range = " + event.range );
Logger.log( "rTop = " + rTop );
Logger.log( "rLeft = " + rLeft );
Logger.log( "rBottom = " + rBottom );
Logger.log( "rRight = " + rRight );

    //loop through all ranges of checkboxes to see if they overlap with the range being edited
    let thisCBRange: Array<number>|null = null;
    let cTop: number = -1; let cLeft: number = -1; let cBottom: number = -1; let cRight: number = -1;
    for( let r = 0; r < CHECKBOX_RANGES[ sheetEditedName ].length; r++ ) {

      Logger.log( "CHECKBOX_RANGES[ '"+sheetEditedName+"' ]["+r+"] = " + CHECKBOX_RANGES[ sheetEditedName ][r] );
      thisCBRange = CHECKBOX_RANGES[ sheetEditedName ][r];
      Logger.log( "thisCBRange = " + thisCBRange );
      cTop = thisCBRange[0]; cLeft = thisCBRange[1]; cBottom = thisCBRange[2]; cRight = thisCBRange[3];

Logger.log( "cTop = " + cTop );
Logger.log( "cLeft = " + cLeft );
Logger.log( "cBottom = " + cBottom );
Logger.log( "cRight = " + cRight );
Logger.log( "" );
Logger.log( "Checking do they overlap" );

      //Do the two ranges overlap
      if( rBottom >= cTop && rRight >= cLeft && rTop <= cBottom && rLeft <= cRight ) {
        //find the range where the two overlap
        const oTop = Math.max( rTop, cTop );
        const oLeft = Math.max( rLeft, cLeft );
        const oBottom = Math.min( rBottom, cBottom );
        const oRight = Math.min( rRight, cRight );
  /*
  Logger.log( "oTop = " + oTop );
  Logger.log( "oLeft = " + oLeft );
  Logger.log( "oBottom = " + oBottom );
  Logger.log( "oRight = " + oRight );
  */
        //now loop through all the OVERLAPPING cells from the original range and check if there is a checkbox in it
        //if there isnt we parse whatever value is in there and convert it to a boolean value then re-add the checkbox
        //CURSE YOU GOOGLE SHEETS FOR MAKING CHECKBOXES SO EASY TO DELETE
        let cell = null; let rule = null; let value = null;
        for (let x = oLeft; x <= oRight; x++) {
          for (let y = oTop; y <= oBottom; y++) {
            cell = range.getCell( y-rTop+1, x-rLeft+1,);
            rule = cell.getDataValidation();
            value = cell.getValue();
            if( rule == null || rule.getCriteriaType() != SpreadsheetApp.DataValidationCriteria.CHECKBOX ) {
              //if it's not a checkbox then we must make sure the value is checkbox friendly ready for when re readd the checkboxes
              //but only set the value if we have to because it takes ages each time we have to do it and onEdit needs to 
              //complete quickly
              if( value !== true && value !== false && value != "" ) {
                cell.setValue( false );
                //cell.insertCheckboxes();
              }
            }
          }
        }
        event.range.getSheet().getRange( oTop, oLeft, oBottom-oTop+1, oRight-oLeft+1 ).insertCheckboxes();
      }
    }
  }

}