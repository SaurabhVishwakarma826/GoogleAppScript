function currencyExchange() {
  let ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheetName = "AED - INR";
  let sheet = ss.getSheetByName(sheetName);
  let floatPoints = 3;
  let url = "https://open.er-api.com/v6/latest/AED";
  Logger.log("fetching data");
  let response = UrlFetchApp.fetch(url);
  let data = JSON.parse(response.getContentText());
  let dateUTC = new Date(data.time_last_update_utc);
  // Adjust for IST (UTC +5 hours and 30 minutes)
  let dateIST = new Date(dateUTC.getTime() + (5.5 * 60 * 60 * 1000));
  let nextDateUTC = new Date(data.time_next_update_utc);
  let nextDateIST = new Date(nextDateUTC.getTime() + (5.5 * 60 * 60 * 1000));
  nextDateIST.setMinutes(nextDateIST.getMinutes() + 5);
  let exchangeRate = data.rates.INR.toFixed(floatPoints);
  // Format date
  let formattedDate = dateIST.toLocaleString("en-US", {
    timeZone: "Asia/Kolkata",
    month: "short",
    day: "numeric",
    year: "numeric",
    weekday: "long"
  });
  Logger.log("AED - INR Append data to the sheet");
  sheet.appendRow([formattedDate, exchangeRate, "", ""]);



 // OMR - AED - INR
  let sheetName2 = "OMR - AED - INR";
  let sheet2 = ss.getSheetByName(sheetName2);
  let url2 = "https://open.er-api.com/v6/latest/OMR";
  Logger.log("fetching data");
  let response2 = UrlFetchApp.fetch(url2);
  let data2 = JSON.parse(response2.getContentText());

  let exchangeRate2 = data2.rates.AED.toFixed(floatPoints);
  let exchangeRate21 = data2.rates.INR.toFixed(floatPoints);
  
  Logger.log("OMR - AED - INR Append data to the sheet");
  sheet2.appendRow([formattedDate, exchangeRate2, exchangeRate21, "", ""]);


  // SAR - AED -INR
  let sheetName3 = "SAR - AED -INR";
  let sheet3 = ss.getSheetByName(sheetName3);
  let url3 = "https://open.er-api.com/v6/latest/SAR";
  Logger.log("fetching data");
  let response3 = UrlFetchApp.fetch(url3);
  let data3 = JSON.parse(response3.getContentText());

  let exchangeRate3 = data3.rates.AED.toFixed(floatPoints);
  let exchangeRate31 = data3.rates.INR.toFixed(floatPoints);
  
  Logger.log("SAR - AED -INR Append data to the sheet");
  sheet3.appendRow([formattedDate, exchangeRate3, exchangeRate31, "", ""]);

}
