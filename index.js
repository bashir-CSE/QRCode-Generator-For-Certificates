// Function to convert a string to proper case (first letter of each word capitalized)
//----------------------------------------------------------------------------------------
function toProperCase(str) {
  if (typeof str !== "string" || !str) {
    return str; // Return the original value if it's not a string or is empty
  }

  // Capitalize the first letter of each word, and make the rest lowercase
  return str.replace(/\w\S*/g, function (word) {
    return word.charAt(0).toUpperCase() + word.substr(1).toLowerCase();
  });
}

// Function to format date to DD-MM-YYYY format
//-------------------------------------------------------
function formatDate(date) {
  if (
    !date ||
    Object.prototype.toString.call(date) !== "[object Date]" ||
    isNaN(date)
  ) {
    return "N/A"; // Return "N/A" if the date is invalid
  }

  var day = date.getDate().toString().padStart(2, "0"); // Ensure two-digit day
  var month = (date.getMonth() + 1).toString().padStart(2, "0"); // Ensure two-digit month (Months are 0-based)
  var year = date.getFullYear(); // Get the full year

  return `${day}-${month}-${year}`; // Return formatted date as DD-MM-YYYY
}

// Main function to generate QR codes for each student's certificate details
//---------------------------------------------------------------------------
function generateQRCode() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet(); // Get active sheet in Google Sheets
  var data = sheet.getDataRange().getValues(); // Get all data from the sheet

  // Loop through each row of data (starting from the second row)
  for (var i = 1; i < data.length; i++) {
    // Get and format data from each column (name, ID, etc.)
    var studentName = data[i][0] ? data[i][0] : "N/A"; // Convert student name to proper case
    var branchName = data[i][1] ? toProperCase(data[i][1]) : "N/A"; // branch name added
    var studentID = data[i][2] ? data[i][2] : "N/A"; // Use student ID, or "N/A" if empty
    var courseName = data[i][3] ? data[i][3].toString().toUpperCase() : "N/A"; // Convert course name to uppercase
    var fatherName = data[i][4] ? data[i][4] : "N/A"; // Convert father's name to proper case
    var motherName = data[i][5] ? data[i][5] : "N/A"; // Convert mother's name to proper case
    var mobileNumber = data[i][6] ? data[i][6] : "N/A"; // Use mobile number, or "N/A" if empty
    var issueDate = data[i][7] instanceof Date ? formatDate(data[i][7]) : "N/A"; // Format issue date or use "N/A"

    // Construct the verification message to include all student details
    var verificationMessage =
      "Your Certificate is Verified!\r\n" +
      "Branch Name: " +
      branchName +
      "\r\n" +
      "Student Name: " +
      studentName +
      "\r\n" +
      "Student ID: " +
      studentID +
      "\r\n" +
      "Course: " +
      courseName +
      "\r\n" +
      "Father's Name: " +
      fatherName +
      "\r\n" +
      "Mother's Name: " +
      motherName +
      "\r\n" +
      "Mobile: +880" +
      mobileNumber +
      "\r\n" +
      "Issue Date: " +
      issueDate +
      "\r\n";

    // Generate the QR Code URL using QuickChart API with the encoded message
    var qrCodeURL = `https://quickchart.io/qr?text=${encodeURIComponent(
      verificationMessage
    )}&size=400`;

    // Insert the QR code as an image in the sheet in column 9
    var cell = sheet.getRange(i + 1, 9); // Get the cell in column 9 (QR Code column)
    var formula = `=IMAGE("${qrCodeURL}")`; // Create formula to insert the QR code image
    cell.setFormula(formula); // Set the formula in the cell to display the image
  }
}

/**
 * This function clears the QR codes from column 9 (where QR codes are located) in the active sheet.
 * It shows a confirmation popup before clearing the QR codes.
 */
function clearQRCodes() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet(); // Get active sheet
  var ui = SpreadsheetApp.getUi(); // Get the UI to show the alert

  // Show a confirmation dialog
  var response = ui.alert(
    "Clear QR Codes",
    "Are you sure you want to clear all QR codes?",
    ui.ButtonSet.OK_CANCEL
  );

  // If the user clicks "OK", proceed to clear QR codes
  if (response == ui.Button.OK) {
    var lastRow = sheet.getLastRow(); // Get the last row with data
    var qrCodeRange = sheet.getRange(2, 9, lastRow - 1, 1); // Select range from row 2 in column 9

    // Clear the contents in the selected range (QR codes)
    qrCodeRange.clearContent();

    // Provide feedback to the user
    ui.alert(
      "QR Codes Cleared!",
      "All QR codes have been cleared from the sheet.",
      ui.ButtonSet.OK
    );

    // Log the action for reference
    Logger.log("QR Codes were cleared successfully.");
  } else {
    // If the user cancels the action, log and notify
    Logger.log("QR code clearing was canceled by the user.");
    ui.alert("Action Canceled", "QR codes were not cleared.", ui.ButtonSet.OK);
  }
}

//Download All the QR code
function exportSelectedColumnsToPDF() {
  var ui = SpreadsheetApp.getUi();

  // Show confirmation popup
  var response = ui.alert(
    "Download PDF",
    "Do you want to download the QR Code PDF?",
    ui.ButtonSet.OK_CANCEL
  );

  // If the user clicks "Cancel", exit the function
  if (response == ui.Button.CANCEL) {
    ui.alert("❌ Download Canceled");
    return;
  }

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getActiveSheet();
  var sheetId = sheet.getSheetId();
  var sheetName = sheet.getName();

  var folder = DriveApp.getRootFolder(); // Save in root Drive folder

  // Construct the correct PDF export URL with selected columns (A to I)
  var exportUrl =
    "https://docs.google.com/spreadsheets/d/" +
    ss.getId() +
    "/export?format=pdf" +
    "&portrait=true" + // Portrait mode
    "&size=A4" + // Page size
    "&gridlines=true" + // Show gridlines (ensures visible borders)
    "&printtitle=true" + // Keep headers visible
    "&top_margin=0.5" +
    "&bottom_margin=0.5" +
    "&left_margin=0.5" +
    "&right_margin=0.5" +
    "&gid=" +
    sheetId +
    "&range=A1:I" +
    sheet.getLastRow(); // Only export A to I columns dynamically

  // Fetch PDF with authorization
  var token = ScriptApp.getOAuthToken();
  var response = UrlFetchApp.fetch(exportUrl, {
    headers: { Authorization: "Bearer " + token },
  });

  var pdfBlob = response.getBlob().setName(sheetName + "_QR_Codes.pdf");
  var file = folder.createFile(pdfBlob); // Save PDF to Google Drive

  // Show success message with download link
  ui.alert("✅ PDF Created! Click OK to open.");
  Logger.log("Download PDF: " + file.getUrl());
}
