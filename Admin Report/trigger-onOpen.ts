function onOpen( event ){
  
  //Logger.log( 'onOpen trigger commencing');.
  var ui = SpreadsheetApp.getUi();
 
  //These lines create the menu items and 
  // tie them to functions that all exists in this project's Apps Script .gs files
  var mainMenu = ui.createMenu('Master Report Menu')
      .addItem('New Child - Daily Report', 'CreateChildDailyReport')
      .addSeparator()
      .addItem( 'Clear Down test Snapshot data ready for Nightly Batch Jobs', 'ClearReportReadyForNewSnapshot' )
      .addItem( 'Refresh Data Only - No new child', 'RefreshRoSRawDataSnapshot' )

  mainMenu.addToUi();

}