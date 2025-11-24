📊 Google Sheets Sales Data Auto-Mailer

Automatically filter sales data by salesman, generate a temporary Google Sheet, share it with the corresponding email, and delete the temporary file — all using Google Apps Script.

🚀 Overview

This script automates the process of sharing personalized sales reports with each salesperson.

It performs the following actions:

Reads Sheet1 (sales data)

Reads Sheet2 (list of salesmen + emails)

Filters rows belonging to each salesman

Creates a temporary Google Sheet for each

Shares the sheet with view-only access

Sends an email containing the sheet link

Deletes the temporary sheet from Google Drive

📁 Project Structure
├── README.md
└── sendFilteredSheets.gs      # Main Apps Script file

🧠 How It Works
Required Google Sheet Structure
Sheet1 (Sales Data)

Contains all sales rows

Column G (index 6) = Salesman Name

Sheet2 (Email Mapping)
Salesman	Email

Each row corresponds to one salesperson.

🧩 Code (sendFilteredSheets.gs)
function sendFilteredSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const salesSheet = ss.getSheetByName("Sheet1");
  const emailSheet = ss.getSheetByName("Sheet2");

  const salesData = salesSheet.getDataRange().getValues();
  const emailData = emailSheet.getDataRange().getValues();
  const header = salesData[0];

  for (let i = 1; i < emailData.length; i++) {
    const salesman = emailData[i][0];
    const email = emailData[i][1];

    if (!salesman || !email) continue;

    const filteredRows = salesData.filter((row, index) => {
      return index > 0 && row[6] === salesman;
    });

    if (filteredRows.length === 0) continue;

    const tempFile = SpreadsheetApp.create("SalesData_" + salesman);
    const tempSheet = tempFile.getActiveSheet();

    tempSheet.getRange(1, 1, 1, header.length).setValues([header]);
    tempSheet.getRange(2, 1, filteredRows.length, header.length).setValues(filteredRows);

    tempFile.addViewer(email);

    MailApp.sendEmail({
      to: email,
      subject: "your sales data in this sheet",
      body:
        "Dear " + salesman + ",\n\n" +
        "Here is your sales data:\n" +
        tempFile.getUrl() +
        "\n\nRegards,\nVikas Dhakad"
    });

    DriveApp.getFileById(tempFile.getId()).setTrashed(true);
  }
}

📝 Requirements

A Google Workspace or Gmail account

Access to Google Apps Script

A Google Sheet with properly formatted Sheet1 and Sheet2

Official Google Apps Script references used:

Google Apps Script → Spreadsheet Service
https://developers.google.com/apps-script/reference/spreadsheet

Google Apps Script → Drive Service
https://developers.google.com/apps-script/reference/drive

Google Apps Script → MailApp
https://developers.google.com/apps-script/reference/mail/mail-app

▶️ How to Deploy

Open your Google Sheet

Go to Extensions → Apps Script

Paste the code into your project

Save

Run the function:

sendFilteredSheets()


Approve permissions when prompted

📮 Email Output

Each salesperson receives an email containing:

A personal greeting

A link to their filtered sales data

Automatically generated temporary sheet (deleted after sending)

🔐 Permissions Needed

The script requires authorization for:

Sending emails (MailApp)

Creating files (DriveApp)

Managing spreadsheets (SpreadsheetApp)

These permissions come directly from Google's official Apps Script authorization system.

🤝 Contributing

Pull requests, improvements, and feature suggestions are welcome.
