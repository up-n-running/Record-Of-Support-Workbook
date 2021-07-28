//var not const because blobs get populated on the fly
var LEARNER_EMAIL_IMAGES = {
  PDF_HELP_1 : { FILE_ID : "1VYNzECOLhWQ-Idupxkwopo0eiNwtfCL_", BLOB : null },
  PDF_HELP_2 : { FILE_ID : "15vqGtGGG1GdgeqPn3mnj-M2mFG2UnnNe", BLOB : null },
  PDF_HELP_3 : { FILE_ID : "1dPOjIF6BiK9_u5g7DiUuUJzZQFpGhj7_", BLOB : null },
  PDF_HELP_4 : { FILE_ID : "1jR7T42dqc6g8mhVnmPsZiiT5iae2lGNt", BLOB : null },
  PDF_HELP_5 : { FILE_ID : "1i4mQaB8v329dsLWDOdaaIzxGLyG5s99e", BLOB : null },
  PDF_HELP_6 : { FILE_ID : "1GP3ZSsEcECUN3MMJ-tid_n7-cd1RGKcY", BLOB : null },
  PDF_HELP_7 : { FILE_ID : "19F8IZFKlrsd1VzJMKgwz-hzzAt_LkO69", BLOB : null },
  STORED_HELP_3 : { FILE_ID : "1DX9C058wMO2FMDnxER1OpNLwfmbrMqRQ", BLOB : null },
  STORED_HELP_4 : { FILE_ID : "1iAfbFJq96WiIjHMHkqFrkYvbuwCOE0kA", BLOB : null },
  HELP_VIDEO : { FILE_ID : "1twocba1Rkui3FF-qN7tYUdvx7ihTqqYi", BLOB : null, VIDEO_URL_PDF: "https://youtu.be/x7rR0MUX0-Y" }
};

function getInlineImages( signType: string ) {
  if( signType == "PDF" ) return {
    pdfHelp1: LEARNER_EMAIL_IMAGES.PDF_HELP_1.BLOB,
    pdfHelp2: LEARNER_EMAIL_IMAGES.PDF_HELP_2.BLOB,
    pdfHelp3: LEARNER_EMAIL_IMAGES.PDF_HELP_3.BLOB,
    pdfHelp4: LEARNER_EMAIL_IMAGES.PDF_HELP_4.BLOB,
    pdfHelp5: LEARNER_EMAIL_IMAGES.PDF_HELP_5.BLOB,
    pdfHelp6: LEARNER_EMAIL_IMAGES.PDF_HELP_6.BLOB,
    pdfHelp7: LEARNER_EMAIL_IMAGES.PDF_HELP_7.BLOB,
    helpVideoThumb: LEARNER_EMAIL_IMAGES.HELP_VIDEO.BLOB
  };
  else return {
    storedHelp1: LEARNER_EMAIL_IMAGES.PDF_HELP_1.BLOB,
    storedHelp3: LEARNER_EMAIL_IMAGES.STORED_HELP_3.BLOB,
    storedHelp4: LEARNER_EMAIL_IMAGES.STORED_HELP_4.BLOB,
    //helpVideoThumb: LEARNER_EMAIL_IMAGES.HELP_VIDEO.BLOB
  };
}

function EnsureLearnerEmailBlobsLoaded_() {
  //get images as blobs if not already loaded into memory
  if( LEARNER_EMAIL_IMAGES.PDF_HELP_1.BLOB == null ) {
    LEARNER_EMAIL_IMAGES.PDF_HELP_1.BLOB = DriveApp.getFileById(LEARNER_EMAIL_IMAGES.PDF_HELP_1.FILE_ID).getBlob();
    LEARNER_EMAIL_IMAGES.PDF_HELP_2.BLOB = DriveApp.getFileById(LEARNER_EMAIL_IMAGES.PDF_HELP_2.FILE_ID).getBlob();
    LEARNER_EMAIL_IMAGES.PDF_HELP_3.BLOB = DriveApp.getFileById(LEARNER_EMAIL_IMAGES.PDF_HELP_3.FILE_ID).getBlob();
    LEARNER_EMAIL_IMAGES.PDF_HELP_4.BLOB = DriveApp.getFileById(LEARNER_EMAIL_IMAGES.PDF_HELP_4.FILE_ID).getBlob();
    LEARNER_EMAIL_IMAGES.PDF_HELP_5.BLOB = DriveApp.getFileById(LEARNER_EMAIL_IMAGES.PDF_HELP_5.FILE_ID).getBlob();
    LEARNER_EMAIL_IMAGES.PDF_HELP_6.BLOB = DriveApp.getFileById(LEARNER_EMAIL_IMAGES.PDF_HELP_6.FILE_ID).getBlob();
    LEARNER_EMAIL_IMAGES.PDF_HELP_7.BLOB = DriveApp.getFileById(LEARNER_EMAIL_IMAGES.PDF_HELP_7.FILE_ID).getBlob();
    LEARNER_EMAIL_IMAGES.STORED_HELP_3.BLOB = DriveApp.getFileById(LEARNER_EMAIL_IMAGES.STORED_HELP_3.FILE_ID).getBlob();
    LEARNER_EMAIL_IMAGES.STORED_HELP_4.BLOB = DriveApp.getFileById(LEARNER_EMAIL_IMAGES.STORED_HELP_4.FILE_ID).getBlob();
    LEARNER_EMAIL_IMAGES.HELP_VIDEO.BLOB = DriveApp.getFileById(LEARNER_EMAIL_IMAGES.HELP_VIDEO.FILE_ID).getBlob();
  }
}

function SendLearnerSignEmail( spreadsheet, newFiles, oldUnsignedFiles, recordEmailAddress, recordLearnerName, 
                               friendlyLearnerName, recordLSAName, learnerId, learnerSignSype, 
                               learnerSignatireFileId, showToast, learnerWebappDeployId, emailTemplate ) {

  if( showToast ) {
    SpreadsheetApp.getActive().toast( "Emailing " + recordLearnerName + " at: " + recordEmailAddress, 
                                      "Emailing " + friendlyLearnerName );
  }

  if(!emailTemplate) {
    emailTemplate = HtmlService.createTemplateFromFile('html/html-email-learner-tosign');
  }

  let effectiveSignType = ( learnerSignSype==="Stored" && learnerSignatireFileId )? "Stored" : "PDF";

  //generate the subject line and introductary salutation text
  newFiles = ( newFiles ) ? newFiles : new Array();
  oldUnsignedFiles = ( oldUnsignedFiles ) ? oldUnsignedFiles : new Array();
  let subjectLine = null;
  let salutationLine = null;
  let newFilesIntro = null;
  let oldFilesIntro = null;

  if( newFiles.length > 1 ) {
    subjectLine = "Please sign your "+newFiles.length+" new records of support";
    salutationLine = "I've just sent you "+newFiles.length+" new records of support.";
    newFilesIntro = "Your new files to sign are:";
  }
  else if( newFiles.length == 1 ) {
    subjectLine = "Please sign your new record of support";
    salutationLine = "I've just sent you a new record of support.";
    newFilesIntro = "Your new file to sign is:";
  }
  else if( oldUnsignedFiles.length > 1 ) {
    subjectLine = "IMPORTANT: You still have "+oldUnsignedFiles.length+" records of support to sign";
    salutationLine = "you stil have "+oldUnsignedFiles.length+" records of support to sign.";
  }
  else if( oldUnsignedFiles.length == 1 ) {
    subjectLine = "IMPORTANT: You still have a record of support to sign";
    salutationLine = "you stil have 1 record of support to sign.";
  }

  subjectLine += "  ["+Math.floor(Math.random() * 1000000)+"]";

  oldFilesIntro = ( oldUnsignedFiles.length > 1 ) ? "You still need to sign these old files:" : 
                  ( ( oldUnsignedFiles.length == 1 ) ? "You still need to sign this old file:" : null );

Logger.log("subjectLine = " + subjectLine);
Logger.log("salutationLine = " + salutationLine);
Logger.log("newFilesIntro = " + newFilesIntro);
Logger.log("oldFilesIntro = " + oldFilesIntro);

  //set stored signatire webapp URL
  let storeSignatureLinkURL = "https://script.google.com/a/macros/college.wlc.ac.uk/s/" + learnerWebappDeployId + "/exec?learnerId="+learnerId+"&lsaWorkbookId="+spreadsheet.getId();

  emailTemplate.introText     = salutationLine;
  emailTemplate.newFilesIntro = newFilesIntro;
  emailTemplate.newFiles      = newFiles;
  emailTemplate.oldFilesIntro = oldFilesIntro;
  emailTemplate.oldFiles      = oldUnsignedFiles;
  emailTemplate.friendlyLearnerName = friendlyLearnerName;
  emailTemplate.signType      = effectiveSignType;
  emailTemplate.learnersPreferredSignType = learnerSignSype;
  emailTemplate.storeSignatureLinkURL = storeSignatureLinkURL;
  emailTemplate.helpVideoURL  = ( ( effectiveSignType == "PDF" ) ? LEARNER_EMAIL_IMAGES.HELP_VIDEO.VIDEO_URL_PDF : null );
  emailTemplate.recordLSAName = recordLSAName;
  var emailBodyHtml = emailTemplate.evaluate().getContent();
  Logger.log( emailBodyHtml );
  
  //get images as blobs if not already loaded into memory
  EnsureLearnerEmailBlobsLoaded_();

  MailApp.sendEmail({
    to: substututeIfPlaceholderEmailAddress_( recordEmailAddress ),
    subject: subjectLine,
    htmlBody: emailBodyHtml,
    inlineImages: getInlineImages( effectiveSignType )
  });
}

function isPlaceholderEmailAddress_( emailAddress ) {
  return ( emailAddress == "YOUR_EMAIL_ADDRESS@wlc.ac.uk" );
}

function substututeIfPlaceholderEmailAddress_( emailAddress ) {
  return ( isPlaceholderEmailAddress_( emailAddress ) ? Session.getActiveUser().getEmail() : emailAddress );
}