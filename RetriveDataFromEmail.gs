function processIncomingEmails() {
  var threads = GmailApp.getInboxThreads();
  
  for (var i = 0; i < threads.length; i++) {
    var messages = threads[i].getMessages();
    
    for (var j = 0; j < messages.length; j++) {
      var message = messages[j];
      var subject = message.getSubject();
      
      if (subject === "The Score") {
        var attachments = message.getAttachmentsByType('application/vnd.ms-excel');
        
        if (attachments.length > 0) {
          var sheet = SpreadsheetApp.getActiveSpreadsheet();
          var attachmentBlob = attachments[0];
          var sheetCopy = DriveApp.createFile(attachmentBlob).makeCopy();
          var sheetCopyId = sheetCopy.getId();
          var sheetCopyFile = DriveApp.getFileById(sheetCopyId);
          var sheetCopyUrl = sheetCopyFile.getUrl();
          
          sheet.toast("Sheet copied. URL: " + sheetCopyUrl, "Success", 5);
          break;
        }
      }
    }
  }
}

function onMailReceive(event) {
  processIncomingEmails();
}
