// Notes:
// Tried using the ContactsApp to get the Manager and full name for associate, but that does not seem to work.
// This made it less clean for now on requiring to fill these fields in the spreadsheet.

// Identifiers are the relevant part of the URL when you open these from google drive
var formFile = DriveApp.getFileById("replace-this"); // Survey form
var quarterlyFolder = DriveApp.getFolderById("replace-this"); // Google Drive Folder to place all created survey forms in

// Spreadsheet format header line is: Project/Team, Name, Email, Manager Email, Reviewers Email (comma separated). The script will update columns to the right with link to form and for email creation
var reviewerSheet = SpreadsheetApp.openById("replace-this").getSheetByName("Reviewers"); // Google spreadsheet with list of surveyes per associate
var formTitleTemplate = "Q2FY20 Feedback Review - %s (%s)";

var viewerEmail = ""; // ATM only editors are supported on google forms, in case it's required should be xxx@redhat.com
var filterOnlyDirectsOf = Session.getActiveUser().getEmail();  // in case the value is "" it's ignored
var emailSignature = "-- \n" + "replace-this" + "\n";

function createFormPerAssociate(title) {
  var associateFormFile = formFile.makeCopy(title, quarterlyFolder);
  var associateForm = FormApp.openById(associateFormFile.getId());
  associateForm.setTitle(title);
  return associateForm;
}

function GenerateForms() {
  var range = reviewerSheet.getRange(1,1,reviewerSheet.getLastRow(),10);

  var associateForm;

  var project; // should add to form or email? (not currently used)
  var associateName;
  var associateId;
  var managerEmail;
  var reviewers;
  var row;
  var rowValues;
  var x;
  var now;
  var email; // helper while creating drafts rather than just sending the email
  var title;

  //for (var i=1; i < 2; i++){ // start with 1 - skip header row
  for (var i=1; i<range.getLastRow(); i++){ // start with 1 - skip header row
    try {
      row = range.offset(i,0,1);
      rowValues = row.getValues()[0];
      // use names for field for readability...
      project = rowValues[0];
      associateName = rowValues[1];
      associateId = rowValues[2];
      managerEmail = rowValues[3] + "@redhat.com";
      reviewers = rowValues[4].split(",").map(function(item) { return item.trim() + "@redhat.com" })

      if (managerEmail != filterOnlyDirectsOf) {
        continue; // skip
      }

      // creeate Form for associate
      title = Utilities.formatString(formTitleTemplate, associateName, associateId);

      associateForm = createFormPerAssociate(title);

      // flag as form created for tracking
      now = new Date();
      row.getCell(1,9).setFormula('=HYPERLINK("' + associateForm.getPublishedUrl() + '", "Form Created ' + now + '")');

      // add manager as editor
      associateForm.addEditor(managerEmail);
      if (viewerEmail != managerEmail && viewerEmail != "") {
        //associateForm.addViewer(viewerEmail); // viewers aren't supported
        associateForm.addEditor(viewerEmail);
      }

      // send email to reviewers and cc manager
      for (var j = 0; j < reviewers.length; j++) {
        email = GmailApp.createDraft(reviewers[j], title, "You are kindly asked to provide feedback via this form.\n\n" + associateForm.getPublishedUrl() + "\n\nResults will be visible only to the direct lead of the reviewed person.\n\n" + emailSignature, { cc: managerEmail });
        //row.getCell(1,10).setValue("Email Draft");
        //email.send();
        //row.getCell(1,10).setValue("Email Sent");
      }
    }
    catch (err) {
      Logger.log("Error in line %s - %s", i+2, err.message); // +2 as array starts from zero and our first row is row 2.
    }
  }
}
