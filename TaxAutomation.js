/**
 * Move any email that is labelled "tax" to a Google Drive folder called Tax Deductions.
 * From there, Microsoft Power Automate can detect new files and write them to Microsoft Onedrive.
 */
function processLabeledEmails() {
  
  var driveFolderName = "Tax Deductions";
  var folders = DriveApp.getFoldersByName(driveFolderName);
  var folder = folders.hasNext() ? folders.next() : DriveApp.createFolder(driveFolderName);

  var threads = GmailApp.search('label:tax newer_than:1d'); // get any emails tagged labelled tax in the last day
  Logger.log("Found %s Threads to process", threads.length);

  // process each thread
  for (var i = 0; i < threads.length; i++) {
    
    var messages = threads[i].getMessages();
    Logger.log("Found %s Messages to process", messages.length);

    // process each message
    for (var j = 0; j < messages.length; j++) {
      var message = messages[j];
      var receivedDate = message.getDate();
      Logger.log("Processing message: " + message.getSubject());

      // process each attachment
      var attachments = messages[j].getAttachments();
      for (var k = 0; k < attachments.length; k++) {
        var attachment = attachments[k];
        //TODO - add a random name to the filename
        folder.createFile(attachment); // Save each attachment to "Deductions"
        Logger.log("Saved: " + attachment.getName());
      }

    }
  }
  GmailApp.sendEmail("charles.poulsen@gmail.com", "Tax Automation Ran " + i +  " messages", "NA" );

  //test
}



