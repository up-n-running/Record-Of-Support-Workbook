function debugCatchError( error ) {
  Logger.log( "EXCEPTION" );
  Logger.log( "error = " + error );
  Logger.log( "error.name = " + error.name );
  Logger.log( "error.message = " + error.message );
  Logger.log( "error.stack = " + error.stack );
}