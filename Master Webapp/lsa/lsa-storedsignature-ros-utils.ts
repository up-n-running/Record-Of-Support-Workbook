// Compiled using ts2gas 3.6.4 (TypeScript 4.2.2)
function createLearnerRoSFromSheetsTemplate_(spreadsheet, sourceRoSPreviewSheet, destLearnerRosFolder, learnerRoSTemplateFileId, destFileName, learnerEmailAddress, learnerSignatureFileId) {
    var learnerRoSTemplateFile = DriveApp.getFileById(learnerRoSTemplateFileId);
    var learnerRoSTemplateSpreadsheet = SpreadsheetApp.open(learnerRoSTemplateFile);
    //create Dest Learner RoS File
    var learnerDestRoSFile = learnerRoSTemplateFile.makeCopy(destFileName, destLearnerRosFolder);
    var learnerDestRoSSpreadsheet = SpreadsheetApp.open(learnerDestRoSFile);
    var learnerDestRoSSheet = learnerDestRoSSpreadsheet.getSheetByName(SHEETS.LEARNER_TEMPLATE_REMOTE.NAME);
    //copy data from source to dest
    var copyScript = SHEETS.LEARNER_TEMPLATE_REMOTE.COPY_SCRIPT;
    var copyScriptRow = null;
    var sourceRoSRange = null;
    var destRoSRange = null;
    var tempValue = null;
    for (var i = 0; i < copyScript.length; i++) {
        copyScriptRow = copyScript[i];
        sourceRoSRange = sourceRoSPreviewSheet.getRange(copyScriptRow.SOURCE_RANGE);
        destRoSRange = learnerDestRoSSheet.getRange(copyScriptRow.DEST_ROW, copyScriptRow.DEST_COL, sourceRoSRange.getHeight(), sourceRoSRange.getWidth());
        if (copyScriptRow.SINGLE_VALUE) {
            tempValue = sourceRoSRange.getValue();
            Logger.log("tempValue = '" + tempValue + "'");
            if (copyScriptRow.COPY_BLANK || tempValue) {
                Logger.log("Saving value");
                destRoSRange.setValue(tempValue);
            }
        }
        else {
            Logger.log("Saving range");
            destRoSRange.setValues(sourceRoSRange.getValues());
        }
    }
    //work out if the sheet is already auto-signed
    var signatureValue = sourceRoSPreviewSheet.getRange(SHEETS.MASTER_TEMPLATE.REFS.ROW_NO_SIGNATURE, SHEETS.MASTER_TEMPLATE.REFS.COL_NO_SIGNATURE).getValue();
    Logger.log("signatureValue = '" + signatureValue + "'");
    var isAutoSigned = (signatureValue != "");
    //set dest hidden settings
    learnerDestRoSSheet.getRange(SHEETS.LEARNER_TEMPLATE_REMOTE.REFS.ROW_NO_HIDDEN_SETTINGS, SHEETS.LEARNER_TEMPLATE_REMOTE.REFS.COL_NO_LEARNER_EMAIL).setValue(learnerEmailAddress);
    learnerDestRoSSheet.getRange(SHEETS.LEARNER_TEMPLATE_REMOTE.REFS.ROW_NO_HIDDEN_SETTINGS, SHEETS.LEARNER_TEMPLATE_REMOTE.REFS.COL_NO_LEARNER_SIGNATURE_FILE_ID).setValue(learnerSignatureFileId);
    //set Permissions on Cells for Learner so learner has access to just the cells they need
    //these are the cells with a range level ptorection on them.
    if (!isAutoSigned) {
        var rangeProtections = learnerDestRoSSheet.getProtections(SpreadsheetApp.ProtectionType.RANGE);
        var p = null;
        for (var j = 0; j < rangeProtections.length; j++) {
            p = rangeProtections[j];
            p.addEditors([learnerEmailAddress]);
        }
    }
    //FINAL thing is to give the learner permission to access the file
    learnerDestRoSFile.addEditor(learnerEmailAddress);
    return learnerDestRoSFile;
}
