# QR Code Generator and Management in Google Sheets

This repository contains a Google Apps Script for managing QR codes in a Google Sheets document. The script enables users to generate QR codes for each student's certificate details, clear the generated QR codes, and export the data (columns A to I) as a PDF.

## Features

- **Generate QR Codes**: Generate QR codes for each student's certificate based on specific details (name, ID, course, etc.) stored in the sheet.
- **Clear QR Codes**: Clear all QR codes from the sheet after generating them with a confirmation popup.
- **Download PDF**: Export the data (columns A to I) as a PDF, including all QR codes, with visible borders and correct formatting.

## Prerequisites

- A **Google Spreadsheet** to store student certificate data.
- The script requires **Google Apps Script** to be enabled and configured in the Google Sheets document.
- **Google Drive** for saving the exported PDF files.

## Setup Instructions

1. Open your **Google Sheet** and navigate to **Extensions > Apps Script**.
2. Paste the following **JavaScript code** into the Apps Script editor.
3. Save the script and close the editor.
4. **Run the functions** directly from the Apps Script editor or link them to buttons in the Google Sheet.

## Functions

### 1. `generateQRCode()`
Generates QR codes for each student based on the details in the spreadsheet. The QR code is inserted in column 9.

#### Key Details:
- Converts names and details to proper case.
- Includes student name, ID, course, mobile, and issue date in the QR code.
- Uses the QuickChart API to generate QR codes.

### 2. `clearQRCodes()`
Clears all the QR codes from column 9 in the active sheet after a **confirmation popup**.

#### Key Details:
- Asks the user for confirmation before clearing the QR codes.
- Clears the content of cells in column 9.

### 3. `exportSelectedColumnsToPDF()`
Exports the sheet (columns A to I) as a **PDF** to Google Drive with gridlines visible and correct formatting.

#### Key Details:
- PDF is exported in **portrait** mode with **A4** size.
- The export includes only columns A to I and all rows with data.
- Saves the PDF file to the **root Google Drive folder**.
- Displays a confirmation message and download link.

## How to Use

1. **Generate QR Codes**:
   - Click on **Extensions > Apps Script** and run the `generateQRCode()` function to generate QR codes for each student's certificate.

2. **Clear QR Codes**:
   - If you want to remove QR codes, click on **Extensions > Apps Script** and run the `clearQRCodes()` function.
   - A confirmation popup will appear asking for confirmation before clearing the QR codes.

3. **Download Data as PDF**:
   - To download the data (columns A to I) as a PDF, click on **Extensions > Apps Script** and run the `exportSelectedColumnsToPDF()` function.
   - A confirmation popup will appear asking if you want to download the PDF.
   - After clicking **OK**, the PDF will be generated and saved to your Google Drive.

## Code Comments

- **`toProperCase()`**: Converts a string to proper case (capitalizes the first letter of each word).
- **`formatDate()`**: Formats a date into the `DD-MM-YYYY` format.
- **`generateQRCode()`**: Generates a QR code for each student based on their certificate details.
- **`clearQRCodes()`**: Clears all QR codes in column 9 and asks for confirmation.
- **`exportSelectedColumnsToPDF()`**: Exports selected columns (A to I) to a PDF file.

## Notes

- The QR codes are generated using the QuickChart API. The QR code will contain all the certificate details for the student.
- The script only handles **column 9** (QR Code column) for QR code insertion.
- The PDF export will include gridlines, headers, and formatting as specified in the URL parameters.

## License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.

## Acknowledgments

- **QuickChart API** for generating QR codes.
- **Google Apps Script** for seamless integration with Google Sheets.

---

This **README** will help users understand the script’s functionality, installation, and usage.
