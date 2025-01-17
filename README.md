# Automated Invoice Project
## Links
#Google SpreadSheet
https://docs.google.com/spreadsheets/d/1-idj-i378g55cMq_t_8ntNyjHfXC9dakxklBkPufZn4/edit?usp=sharing

#Google Docs
https://docs.google.com/document/d/1stLR_n6SWAdPEPsi0gbe3f6r63kgolMd6HcC1ywNt30/edit?usp=sharing

#Notion
https://www.notion.so/Invoice-Generator-Payment-Tracker-ac4da496f2dd4096a01037432c0da4bf?pvs=4

## Overview

This project consists of a Google Apps Script that automatically generates Google Docs from a Google Sheets spreadsheet. The generated documents are based on a provided template and are populated with data from the spreadsheet, including calculated total amounts for each item.

## Features

- Create Google Docs from a template.
- Automatically fill placeholders in the template with data from a Google Sheets spreadsheet.
- Calculate total amounts for items based on prices and quantities provided in the spreadsheet.
- Generate URLs to the created Google Docs and store them back in the spreadsheet.

## Prerequisites

- A Google Sheets spreadsheet with the following columns: Company Name, Company Address, Invoice Date, Invoice No, Description, Price, Quantity, Sub Total, Grand Total.
- A Google Docs template with placeholders for the data.
- Google Apps Script enabled on your Google account.

## Setup

1. **Google Sheets**:
   - Ensure your Google Sheets spreadsheet has a sheet named "Sheet1".
   - Add data starting from row 2, with the headers in row 1.

2. **Google Docs Template**:
   - Create a Google Docs template with placeholders in the following format: `{{Placeholder}}`.
   - Include placeholders for descriptions, prices, quantities, and total amounts for up to 5 items (e.g., `{{Description 1}}`, `{{Price 1}}`, `{{Quantity 1}}`, `{{Total Amount 1}}`, etc.).
   
3. **Google Apps Script**:
   - Open the Google Sheets spreadsheet.
   - Go to `Extensions` > `Apps Script`.
   - Delete any existing code and paste the provided script into the editor.
   - Save the script with a name (e.g., `AutoFill Docs`).
   - Replace the `googleDocTemplateId` and `destinationFolderId` with your own Google Docs template ID and destination folder ID.

## Script

```javascript
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  const menu = ui.createMenu('AutoFill Docs');
  menu.addItem('Create New Docs', 'createNewGoogleDocs');
  menu.addToUi();
}


function createNewGoogleDocs() {
  const googleDocTemplateId = '1stLR_n6SWAdPEPsi0gbe3f6r63kgolMd6HcC1ywNt30';
  const destinationFolderId = '1AiaisD-7Qm4CuQ29sQMI_K-9oaoP6E8C';

  try {
    const googleDocTemplate = DriveApp.getFileById(googleDocTemplateId);
    const destinationFolder = DriveApp.getFolderById(destinationFolderId);
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Output Sheet');
    
    // Find the last row with data
    const lastRow = sheet.getLastRow();
    // Get data starting from row 2
    const rows = sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).getValues();

    rows.forEach((row, index) => {
      if (row[6] && row[7]) {  // Ensure that Price and Quantity columns have data
        const copy = googleDocTemplate.makeCopy(`${row[1]} Company Details`, destinationFolder);
        const doc = DocumentApp.openById(copy.getId());
        const body = doc.getBody();
        const dateValue = new Date(row[3]);
        const friendlyDate = Utilities.formatDate(dateValue, Session.getScriptTimeZone(), "dd/MM/yyyy");

        body.replaceText('{{Company Name}}', row[1] || '');
        body.replaceText('{{Company Address}}', row[2] || '');
        body.replaceText('{{Invoice Date}}', friendlyDate);
        body.replaceText('{{Invoice No}}', row[4] || '');

        // Parse the description, price, and quantity
        const descriptions = row[5].split('/').filter(Boolean);
        const prices = row[6].split('/').filter(Boolean);
        const quantities = row[7].split('/').filter(Boolean);

        for (let i = 0; i < 5; i++) {
          const description = descriptions[i] || 'NULL';
          const price = prices[i] || 'NULL';
          const quantity = quantities[i] || 'NULL';

          // Calculate total amount for each item
          const totalAmount = (price !== 'NULL' && quantity !== 'NULL') ? (parseFloat(price) * parseFloat(quantity)).toFixed(2) : 'NULL';

          body.replaceText(`{{Description ${i + 1}}}`, description);
          body.replaceText(`{{Price ${i + 1}}}`, price);
          body.replaceText(`{{Quantity ${i + 1}}}`, quantity);
          body.replaceText(`{{Total Amount ${i + 1}}}`, totalAmount);
        }

        body.replaceText('{{Sub Total}}', row[13] || '');
        body.replaceText('{{Grand Total}}', row[14] || '');

        doc.saveAndClose();
        const url = doc.getUrl();
        sheet.getRange(index + 2, 16).setValue(url);  // Adjust index + 2 to match row number
      }
    });
  } catch (e) {
    Logger.log(`Error: ${e.message}`);
  }
}

```

## Usage

1. **Open the Spreadsheet**: Open the Google Sheets spreadsheet.
2. **Run the Script**: Click on `AutoFill Docs` in the menu and select `Create New Docs`.
3. **Check the Results**: The script will generate Google Docs based on the template and populate them with the data from the spreadsheet. The URLs to the created documents will be inserted into the spreadsheet.

## Notes

- Ensure that the placeholders in the Google Docs template match exactly with the ones used in the script.
- If there are fewer than 5 items, the placeholders for the missing items will be set to `'NULL'`.
- The script assumes the data in the spreadsheet is formatted correctly and separated by `/` where applicable.

