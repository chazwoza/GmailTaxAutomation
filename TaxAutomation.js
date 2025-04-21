
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
    Logger.log("Thread %s : Found %s Messages to process", i, messages.length);

    // process each message
    for (var j = 0; j < messages.length; j++) {
      var message = messages[j];
      var receivedDate = message.getDate();
      Logger.log("Processing Thread %s Message %s Subject: %s ", i , j, message.getSubject());

  
      // process each attachment
      var attachments = messages[j].getAttachments();
      
      // if attachments exist
      if (attachments.length > 0) {
        for (var k = 0; k < attachments.length; k++) {
          var attachment = attachments[k];
          attachment.setName(addDateToFileName(attachment.getName()));
          folder.createFile(attachment); // Save each attachment to "Deductions"
          Logger.log("Saved File: %s", attachment.getName());
        } 
      } else { // if attachments dont exist
        Logger.log("No attachments for message %s", message.getSubject() );
        // TODO: save the email
      }

    }
  }
  GmailApp.sendEmail("charles.poulsen@gmail.com", "Tax Automation Ran " + i +  " messages", "NA" );
}

/**.Helper function to add date to filename in the format "_dd-mm-yyyy" before the dot */
function addDateToFileName(filename) {
  // Get current date in dd-mm-yyyy format
  var now = new Date();
  var dd = String(now.getDate()).padStart(2, '0');
  var mm = String(now.getMonth() + 1).padStart(2, '0');
  var yyyy = now.getFullYear();
  var dateStr = `${dd}-${mm}-${yyyy}`;

  // Find the last dot in the filename
  var dotIndex = filename.lastIndexOf('.');

  // If there's no extension, just append the date
  if (dotIndex === -1) {
    return `${filename}_${dateStr}`;
  }

  // Otherwise, insert the date before the extension
  var name = filename.substring(0, dotIndex);
  var ext = filename.substring(dotIndex); // includes the dot
  return `${name}_${dateStr}${ext}`;
}


/** Create a trigger to run processLabeledEmails every night */
function createMidnightTrigger() {
 ScriptApp.newTrigger("processLabeledEmails") // Replace with your function name
    .timeBased()
    .everyDays(1)
    .atHour(0) // 0 = midnight
    .create();
}


