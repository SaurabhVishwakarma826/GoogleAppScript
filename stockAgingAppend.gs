function retrieveDataFromEmailFinalP2() {
  var subjectToSearch = "Fwd: Stock Aging Statement-Auto Mail- Company: A-1 FENCE PRODUCTS COMPANY PVT. LTD. ,Plant: P5";
  // var senderToSearch = "saurabh.vishwakarma@a1fenceproducts.com";
  var spreadsheetId = "1j3uLIBgHsy7_VEnrTGTJTbtc0Q96mHgvj70NEbS8tq4";
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
        for (var k = 0; k < attachments.length; k++) {
          var attachment = attachments[k];

          // Check if the attachment is an Excel sheet
          if (attachment.getContentType() === "application/vnd.ms-excel") {
            if (!attachment.isGoogleType()) {
              var blobData = attachment.copyBlob();
              // Create a temporary file from the attachment blob
              var tempFile = DriveApp.createFile(blobData);
              var fileId = tempFile.getId();
              var con = convertExcelToSheets(fileId)
              // Convert Excel file to Google Sheets format
              var spreadsheet = SpreadsheetApp.openById(con);
            }
            // Get the active sheet
            var sheet = spreadsheet.getActiveSheet();
            // Get all the values in the sheet
            var newValues = sheet.getDataRange().getValues();

            // Get the current date
            var currentDate = new Date();


            newValues[0].push("Appended Date");  // Add header
            for (var row = 1; row < newValues.length; row++) {
              newValues[row].push(currentDate);
            }
            var newHeaders = newValues[0];

            // Copy only matching columns to the current sheet in the spreadsheet
            var currentSheet = SpreadsheetApp.openById(spreadsheetId).getSheetByName(sheetName);

            // Assuming that the first row in existingValues contains the headers
            var existingHeaders = currentSheet.getDataRange().getValues()[0];
            Logger.log(newHeaders);
            Logger.log(existingHeaders);
            // Check if headers match
            if (existingHeaders.join(',') === newHeaders.join(',')) {
              // Headers match, paste all values directly
              Logger.log("Headers match");
              if (currentSheet.getLastRow() === 0) {
                currentSheet.getRange(currentSheet.getLastRow() + 1, 1, newValues.length, newValues[0].length).setValues(newValues);
              } else {
                currentSheet.getRange(currentSheet.getLastRow() + 1, 1, newValues.length - 1, newValues[0].length).setValues(newValues.slice(1));
              }

            } else {
              // Headers don't match, 
              var matchingColumnsIndexes = [];

              // Find matching column indexes
              for (var i = 0; i < existingHeaders.length; i++) {
                var columnIndexInNewData = newHeaders.indexOf(existingHeaders[i]);
                if (columnIndexInNewData !== -1) {
                  matchingColumnsIndexes.push(columnIndexInNewData);
                }
              }

              // Create a new array with only matching columns
              var matchingColumnsValues = newValues.map(function (row) {
                return matchingColumnsIndexes.map(function (index) {
                  return row[index];
                });
              });

              // Paste the values into the current sheet
              if (currentSheet.getLastRow() === 0) {
                currentSheet.getRange(currentSheet.getLastRow() + 1, 1, matchingColumnsValues.length, matchingColumnsValues[0].length).setValues(matchingColumnsValues);
              } else {
                currentSheet.getRange(currentSheet.getLastRow() + 1, 1, matchingColumnsValues.length - 1, matchingColumnsValues[0].length).setValues(matchingColumnsValues.slice(1));
              }

            }


            // Delete the temporary spreadsheet
            DriveApp.getFileById(fileId).setTrashed(true);
            // Delete the converted Google Sheets file
            DriveApp.getFileById(con).setTrashed(true);

            // Mark the message as read (optional)
            message.markRead();

            // Log the subject of the email
            Logger.log("Email Subject: " + message.getSubject());
            // Terminate the function execution

            var mailSubject = "Today's Data Updated In Sheet"

            var mailBody = `
                          URL: \n
                          https://docs.google.com/spreadsheets/d/1XUrre-JArWj1S6zQ23f3UBpO-LKMyZW1R4B6sm7HsDs/edit#gid=695570784
            `

            Logger.log(MailApp.getRemainingDailyQuota())

            MailApp.sendEmail({
              to: "saurabh.vishwakarma@a1fenceproducts.com",
              subject: mailSubject,
              body: mailBody,
              cc: ''
            });

            return;
          }
        }
        stopLoop = true;
        break;
      }
    }
  }
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

