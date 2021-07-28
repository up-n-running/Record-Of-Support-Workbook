function SyncMasterLearnerChangesToLSAWorkbooks() {
  let spreadsheet = SpreadsheetApp.getActive();
  //let ui = {p:"w"}; 
  let ui = SpreadsheetApp.getUi();

  Logger.log( 'Calling ReconcileFoldersWithMasterLearnerRecords( spreadsheet, ui, null )' );
  let startTime: number = new Date().getTime();
  if( ReconcileFoldersWithMasterLearnerRecords( spreadsheet, ui, null ) ) {
    //if the above reconciliation didnt throw up any errors then continue
    Logger.log( 'Calling PushToChildWorksheets( null, null, null, true )' );
    spreadsheet.toast( "Syncing learner name changes to all LSA Workbooks", "Running Push Service", 10 );

    PushToChildWorksheets( null, null, null, true, false, false, false, startTime );

    //remove last edited date on global settings so it knows it's up to date
    let settingsSheet = spreadsheet.getSheetByName( SHEETS.GLOBAL_SETTINGS.NAME );
    settingsSheet.getRange( 
          SHEETS.GLOBAL_SETTINGS.REFS.ROW_NO_LAST_MTR_LRNR_EDIT, 
          SHEETS.GLOBAL_SETTINGS.REFS.COL_NO 
    ).setValue( "" );
    SpreadsheetApp.flush();
  }
}

function ReconcileFoldersWithMasterLearnerRecords( spreadsheet, ui, settingsSheet ) {

  //parse params
  spreadsheet = ( spreadsheet ) ? spreadsheet : SpreadsheetApp.getActive();
  ui = ( ui ) ? ui :  SpreadsheetApp.getUi();
  settingsSheet = ( settingsSheet ) ? settingsSheet : spreadsheet.getSheetByName( SHEETS.GLOBAL_SETTINGS.NAME );

  //set timeout high on this notification - normally runs in a second or two but can take around 20s if there's a lot of work to 
  //do to create and rename folders on the shared drive
  spreadsheet.toast( "Syncing learner name changes to and from the shared drive Learner folders.", "Reconciling Learner Folders", 10 );

  //get info about all of the immediate subdirectories of the Learner ROS Root Folder
  let learnerParentDirId = settingsSheet.getRange( 
        SHEETS.GLOBAL_SETTINGS.REFS.ROW_NO_DIR_ID_ROS, 
        SHEETS.GLOBAL_SETTINGS.REFS.COL_NO 
      ).getValue();
  let learnerParentDir = DriveApp.getFolderById( learnerParentDirId );
  let foldersById = {};
  let foldersByName = {};
  let foldersByBoth = {};
  populateSubfolderLookups_( learnerParentDir, foldersById, foldersByName, foldersByBoth );

  //now get the Master Learner Columns that we need to generate folder names reconcile with the associated folder ids
  let masterLearnerSheet = spreadsheet.getSheetByName( SHEETS.MASTER_LEARNERS.NAME );
  let topRow = SHEETS.MASTER_LEARNERS.REFS.ROW_NO_FIRST_LEARNER;
  let height = masterLearnerSheet.getMaxRows() - topRow + 1;
  let folderIdColumn  = masterLearnerSheet.getRange( topRow, SHEETS.MASTER_LEARNERS.REFS.COL_NO_LEARNER_DIR  , height, 1 ).getValues();
  let foreNameColumn  = masterLearnerSheet.getRange( topRow, SHEETS.MASTER_LEARNERS.REFS.COL_NO_FORENAME     , height, 1 ).getValues();
  let surnameColumn   = masterLearnerSheet.getRange( topRow, SHEETS.MASTER_LEARNERS.REFS.COL_NO_SURNAME      , height, 1 ).getValues();
  let learnerIdColumn = masterLearnerSheet.getRange( topRow, SHEETS.MASTER_LEARNERS.REFS.COL_NO_LEARNER_ID   , height, 1 ).getValues();
  let learnerEmailColumn = masterLearnerSheet.getRange( topRow, SHEETS.MASTER_LEARNERS.REFS.COL_NO_EMAIL_ADDRESS, height, 1 ).getValues();

  //loop through records
  let supposedDirName = null; let learnerId = null; let folderId = null; let folderNameMatchArray = null; let learnerEmail = null;
  let foundLearnerIdsSet =  new Set();
  let foundLearnerEmailsSet =  new Set();

  let learnerDirsAlreadyMatchedById = {};
  for( let i=0; i<learnerIdColumn.length; i++ ) {
    learnerId = learnerIdColumn[i][0];
    folderId = folderIdColumn[i][0];
    //if there is no learner on this row then make sure the folderId is blank
    if( learnerId == "" ) {
      folderId = "";
    }
    //if there is a learner on this row then do the reconcilliation
    else {
      //work out what the learner's dir should be called baes on the learner's name and id in the master database
      supposedDirName = foreNameColumn[i][0] + " " + surnameColumn[i][0] + " (" + learnerIdColumn[i][0] + ")";

      //keep track of learnerid and email address in order to check for duplicates
      learnerEmail = learnerEmailColumn[i][0];
      if( foundLearnerIdsSet.has( learnerId ) ) {
        ui.alert( "Duplicate Learner Id", "Multiple records found with Learner Id: '" + learnerId + "'\n" +
                  "Learner "+supposedDirName+" will not be synced as a result./n/nOnce the sync has completed, " + 
                  "please correct the data and then re-run the sync.\n", ui.ButtonSet.OK );
        return false;
      }
      else if( foundLearnerEmailsSet.has( learnerEmail ) && !isPlaceholderEmailAddress_( learnerEmail ) ) {
        ui.alert( "Duplicate Learner Email", "Multiple records found with email address '" + learnerEmail + "'\n" +
                  "Learner "+supposedDirName+" will not be synced as a result./n/nOnce the sync has completed, " + 
                  "please correct the data and then re-run the sync.\n", ui.ButtonSet.OK );
        return false;
      }
      else {
        foundLearnerIdsSet.add( learnerId );
        foundLearnerEmailsSet.add( learnerEmail );

        //if the linked folder exists and has the right name - we need do no more checks - we're done
        if( foldersByBoth[ folderId+"|"+supposedDirName ] ) {
          //do nothing - linked folder is already called the correct thing
        }
        //if there is no linked folder or there is text in there but it's not a valid folder link.
        else if( folderId == "" ||  !foldersById[ folderId ] ) {
          //folder id invalid so lets go look for a folder with the right name
          folderNameMatchArray = new Array();
          if( foldersByName[ supposedDirName ] ) {
            //found it (though there could be more than one with this correct name)
            folderNameMatchArray = foldersByName[ supposedDirName ];
          }
          if( foldersByName[ foreNameColumn[i][0] + " " + surnameColumn[i][0] ] ) {
            //found a close match so rename it to be the exact match and use that as the folder (unless there was also an exact match too)
            //there could be more than one close match so just rename the first one ready for that one to be linked.
            folderNameMatchArray = folderNameMatchArray.concat( foldersByName[ foreNameColumn[i][0] + " " + surnameColumn[i][0] ] );
            folderNameMatchArray[0].setName( supposedDirName );
          }
          
          //did we find 0, 1 or multiple matches
          if( folderNameMatchArray.length == 0 )
          {
            //couldnt find a match, or a close match, so create a brand new folder and link that
            let newLearnerDir = learnerParentDir.createFolder( supposedDirName );
            folderId = newLearnerDir.getId();
          }
          else {
            //we found at least one match so use the first one to be linked
            folderId = folderNameMatchArray[0].getId();
          }
          if( folderNameMatchArray.length > 1 ) {
            //there were multiple matches so warn the user and make them tidy up the directories by hand before trying again.
            ui.alert( "Multiple Folder Matches", "Multiple folders found for Learner: " + supposedDirName + "\n" +
                      "Please manually merge these folders into one and then delete the extra one(s).\n" +
                      "Then re-run this folder reconcilliation", ui.ButtonSet.OK );
            return false;
          }
        }
        //if the folder being linked to exists, but it is not called the right thing.
        else {
          folderNameMatchArray = new Array();
          if( foldersByName[ supposedDirName ] ) {
            //There is another folder with the correct name so we'll use that - already exists elsewhere
            folderNameMatchArray = foldersByName[ supposedDirName ];
          }
          if( foldersByName[ foreNameColumn[i][0] + " " + surnameColumn[i][0] ] ) {
            //found a close match (possibly in addition to an exact match) so use that but rename it to be the exact match
            //it could be for another learner with the same name but different learner id, however the fact that it is
            //in the shared drive with no learnerid in dir name - only the learner name, implies that
            //it has been put there by hand to be picked up by the sync for this learner and linked and renamed to the correct
            //format
            folderNameMatchArray = folderNameMatchArray.concat( foldersByName[ foreNameColumn[i][0] + " " + surnameColumn[i][0] ] );
            folderNameMatchArray[0].setName( supposedDirName );
          }

          //did we find 0, 1 or multiple matches
          if( folderNameMatchArray.length == 0 )
          {
            //couldnt find a match so, given that the learner record already had a valid link against it,
            //and given that there are no ohter directories with this learner's name in it means that by far the
            //most likely scenario is that the user renamed the learner in the master spreadsheet and we need to sync this change down
            //to the shared drive. we do this by renaming the existing dir so it has the new name and keeping the same folderId
            foldersById[ folderId ].setName( supposedDirName );
          }
          else {
            //found a match so just re-link it and ignore the old directory
            folderId = folderNameMatchArray[0].getId();
          }
          if( folderNameMatchArray.length > 1 ) {
            ui.alert( "Multiple Folder Matches", "Multiple folders found for Learner: " + supposedDirName + "\n\n" +
                      "Please manually merge these folders into one and then delete the extra one(s).\n" +
                      "Then re-run the sync.", ui.ButtonSet.OK );
            return false;
          }
        }
      }
    }
    if( folderId != folderIdColumn[i][0] ) {
      //if the folder link had changed then write that back to the spreadsheet
      masterLearnerSheet.getRange( 
        SHEETS.MASTER_LEARNERS.REFS.ROW_NO_FIRST_LEARNER + i, 
        SHEETS.MASTER_LEARNERS.REFS.COL_NO_LEARNER_DIR 
      ).setValue( folderId );
      folderIdColumn[i][0] = folderId; //just keep the column array up to date as we go even though we don't have to - we could do it at the end (but then if we waited till the end and there was an issue half way through then none of the stuff we'd done already would get saved)
    }

    if( folderId != "" ) {
      if( learnerDirsAlreadyMatchedById[ folderId ] ) {
          ui.alert( "Duplicate Learner Records", "Multiple Learners have the Folder id: " + folderId + "\n\n" +
                    "This is probably because you have duplicate records for Learner '"+supposedDirName+"'\n" +
                    "Please remove the duplicate then re-run the sync", ui.ButtonSet.OK );
          return false;
      }
      learnerDirsAlreadyMatchedById[ folderId ] = folderId;
    }
  }

  return true;
}

function populateSubfolderLookups_( parentFolder, foldersById, foldersByName, foldersByBoth ) {
  let learnerSubfolders = parentFolder.getFolders();
  let subFolder = null, subfolderId = null, subFolderName = null;
  while (learnerSubfolders.hasNext()) {
    subFolder = learnerSubfolders.next();
    subfolderId = subFolder.getId();
    subFolderName = subFolder.getName();

    //folder names can be duplicated so we have to store arrays of values just to be on the safe side in case 2 have the same name
    //initialise array if this is the first time we've seen this dir name
    if( !foldersByName[ subFolderName ] ) { 
      foldersByName[ subFolderName ] = new Array(); 
    }
    foldersByName[ subFolderName ].push( subFolder );

    //the other two arrays cannot have duplicates
    foldersById[ subfolderId ] = subFolder;
    foldersByBoth[ subfolderId+"|"+subFolderName ] = subFolder;
  }
  return;
}
