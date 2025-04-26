
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
      let message = messages[j];
      let receivedDate = message.getDate();
      Logger.log("Processing Thread %s Message %s Subject: %s ", i , j, message.getSubject());

      // save email as .eml file
      let eml = savetoEml(message);
      folder.createFile(eml);
      
      // save attachements
      let attachments = messages[j].getAttachments();
      for (var k = 0; k < attachments.length; k++) {
        let attachment = attachments[k];
        attachment.setName(addDateToFileName(attachment.getName()));
        folder.createFile(attachment); // Save each attachment to "Deductions"
        Logger.log("Saved File: %s", attachment.getName());
      } 
  

    }
  }
  GmailApp.sendEmail("charles.poulsen@gmail.com", "Tax Automation Ran " + i +  " messages", "NA" );
}

/**
 * Convert an email message to a file
 */
function savetoEml(message) {
      let subject = message.getSubject().replace(/[\\/:*?"<>|]/g, '');
      let date = Utilities.formatDate(message.getDate(), Session.getScriptTimeZone(), "yyyy-MM-dd_HH-mm-ss");
      let emlName = `${subject} - ${date}.eml`;
      let rawContent = message.getRawContent();

      // Create a blob with the correct MIME type for .eml
      let blob = Utilities.newBlob(rawContent, "message/rfc822", emlName);
      Logger.log("Created email file: %s", blob.getName());
      return blob;
}

/**.Helper function to add date to filename in the format "_dd-mm-yyyy" before the dot */
function addDateToFileName(filename) {
  // Get current date in dd-mm-yyyy format
  let now = new Date();
  let dd = String(now.getDate()).padStart(2, '0');
  let mm = String(now.getMonth() + 1).padStart(2, '0');
  let yyyy = now.getFullYear();
  let dateStr = `${dd}-${mm}-${yyyy}`;

  // Find the last dot in the filename
  let dotIndex = filename.lastIndexOf('.');

  // If there's no extension, just append the date
  if (dotIndex === -1) {
    return `${filename}_${dateStr}`;
  }

  // Otherwise, insert the date before the extension
  let name = filename.substring(0, dotIndex);
  let ext = filename.substring(dotIndex); // includes the dot
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


