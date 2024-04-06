function convertAndOpenExcelFile() {
  var subjectToSearch = "RETRIEVED DATA2";
  var senderToSearch = "saurabh.vishwakarma@a1fenceproducts.com";
  var spreadsheetId = "1XUrre-JArWj1S6zQ23f3UBpO-LKMyZW1R4B6sm7HsDs";
  var sheetName = "Sheet1";

  // Get all threads in the inbox
  var threads = GmailApp.getInboxThreads();

  var stopLoop = false;

  // Iterate over each thread
  for (var i = 0; i < threads.length && !stopLoop; i++) {
    var thread = threads[i];
    var messages = thread.getMessages();

    // Iterate over each message in the thread
    for (var j = 0; j < messages.length; j++) {
      var message = messages[j];

      if (message.getSubject() === subjectToSearch) {
        Logger.log("Mil gaya");

        var attachments = message.getAttachments();
        Logger.log(attachments[0].getContentType());

        for (var k = 0; k < attachments.length; k++) {
          var attachment = attachments[k];

          // Check if the attachment is an Excel file
          if (attachment.getContentType() === "application/vnd.ms-excel" || attachment.getContentType() === "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet") {
            if (!attachment.isGoogleType()) {
              // Convert Excel attachment to Google Sheets format
              var fileBlob = attachment.copyBlob();
              var convertedSpreadsheet = SpreadsheetApp.create('Temp Spreadsheet');
              var tempFile = DriveApp.createFile(fileBlob);
              // var tempSpreadsheet = SpreadsheetApp.open(tempFile);
              // var tempSheet = tempSpreadsheet.getSheets()[0];

              // Copy the data to the converted spreadsheet
              tempFile.copyTo(convertedSpreadsheet);
              DriveApp.getFileById(convertedSpreadsheet.getId()).setTrashed(true);

              // Open the converted Google Sheets file
              var spreadsheet = SpreadsheetApp.openById(convertedSpreadsheet.getId());

              // Access and process the data in the converted spreadsheet as needed
              // ...

              // Delete the temporary files if desired
              tempSpreadsheet.deleteSheet(tempSheet);
              tempFile.setTrashed(true);

              // Mark the message as read (optional)
              message.markRead();

              // Log the subject of the email
              Logger.log("Email Subject: " + message.getSubject());

              // Terminate the function execution
              return;
            }
          }
        }

        stopLoop = true;
        break;
      }
    }
  }
}



function printExcelFileValuesToConsole() {
  var excelFileId = "1XUrre-JArWj1S6zQ23f3UBpO-LKMyZW1R4B6sm7HsDs"; // Replace with the actual file ID of your Excel file in Google Drive
  
  // Convert the Excel file to Google Sheets format
  var sheetsFileId = convertExcelToSheets(excelFileId);
  
  // Open the converted Google Sheets file
  var spreadsheet = SpreadsheetApp.openById(sheetsFileId);
  
  // Get the active sheet
  var sheet = spreadsheet.getActiveSheet();
  
  // Get all the values in the sheet
  var values = sheet.getDataRange().getValues();
  
  // Print the values to the console
  values.forEach(function(row) {
    row.forEach(function(cell) {
      console.log(cell);
    });
  });
}

// Function to convert Excel file to Google Sheets format
function convertExcelToSheets(excelFileId) {
  var url = "https://www.googleapis.com/drive/v3/files/" + excelFileId + "/copy";
  var payload = {
    "mimeType": "application/vnd.google-apps.spreadsheet"
  };
  var headers = {
    "Authorization": "Bearer " + ScriptApp.getOAuthToken(),
    "Content-Type": "application/json"
  };
  
  var response = UrlFetchApp.fetch(url, {
    method: "post",
    headers: headers,
    payload: JSON.stringify(payload)
  });
  
  var responseData = JSON.parse(response.getContentText());
  return responseData.id;
}
