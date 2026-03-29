# Invoice Approval Tracker

A Google Workspace automation that pulls PDF invoices and timesheets from Gmail, stores them in Drive, and surfaces them in a web-based approval interface where reviewers can approve, reject, or hold each document — all without leaving their browser.

Built with Google Apps Script, Gmail API, Drive API, and Google Sheets as the data layer.

---

## The Problem

Project managers were receiving invoices and timesheets by email, forwarding them manually, waiting for replies, and tracking approvals in a shared spreadsheet by hand. The process was slow, error-prone, and hard to audit.

## The Solution

This tool automates the entire intake-to-approval workflow:

1. Emails tagged with a Gmail label are scanned for PDF attachments
2. PDFs are saved to Google Drive (duplicates skipped via MD5 hash check)
3. A Google Sheets tracker is updated with each file and a preview URL
4. Reviewers open a web app that displays the table alongside an inline PDF viewer
5. Each reviewer sets their status (Approved / Rejected / Hold / N/A) via dropdown
6. Approved records are archived weekly with a "Week Ending" date stamp

---

## Features

- **Automated Gmail → Drive pipeline** with duplicate detection
- **Inline PDF viewer** — click any row to preview the document without leaving the page
- **Per-reviewer dropdowns** — each team member has their own approval column
- **Comments field** per document
- **Resizable split-pane layout** — drag the divider or double-click to reset 50/50
- **Auto-refresh every 60 seconds** — pauses automatically if a reviewer is mid-edit
- **Tab switching** between Invoices and Timesheets
- **Weekly archiving** — approved rows move to a records sheet with a week-ending date
- **Timesheet conversion** — `.xlsx` attachments are converted to PDF on the fly via the Sheets export API

---

## Project Structure

```
├── Code.gs       # All server-side logic (Gmail, Drive, Sheets, Web App endpoints)
└── Index.html    # Frontend approval UI (HTML/CSS/JS, served via HtmlService)
```

---

## Prerequisites

- A Google account with access to Gmail, Drive, and Sheets
- A Google Sheets spreadsheet with the following tabs:
  - `Invoices` and `Invoices_Records`
  - `Timesheets` and `Timesheets_Records`
- Google Drive folders for invoices, timesheets, and the invoice archive
- Gmail labels applied to incoming invoice and timesheet emails
- The **Drive API (v2)** enabled in Apps Script (required for `.xlsx` → PDF conversion)

---

## Setup

### 1. Create your Google Sheets tracker

Create a new Google Spreadsheet and add four tabs named exactly:
- `Invoices`
- `Invoices_Records`
- `Timesheets`
- `Timesheets_Records`

Copy the spreadsheet ID from the URL:
```
https://docs.google.com/spreadsheets/d/YOUR_TRACKER_SHEET_ID/edit
```

---

### 2. Create your Drive folders

Create three folders in Google Drive:
- **Invoice folder** — incoming PDFs land here
- **Invoice archive folder** — approved invoices move here after download
- **Timesheet folder** — timesheet PDFs land here

Copy each folder ID from its URL:
```
https://drive.google.com/drive/folders/YOUR_FOLDER_ID
```

---

### 3. Set up Gmail labels

In Gmail, create labels that will be applied to incoming invoice and timesheet emails. You'll need:
- A label for **invoices to be reviewed** (applied manually or via a filter)
- A label for **processed invoices** (applied automatically by the script)
- A label for **timesheets**

Note the exact label names — they are case-sensitive.

---

### 4. Add the script to Apps Script

1. Open your Google Sheets tracker
2. Go to **Extensions > Apps Script**
3. Delete any existing code in `Code.gs` and paste in the contents of `Code.gs` from this repo
4. Click the **+** button next to Files and add an HTML file named `Index`
5. Paste in the contents of `Index.html`

---

### 5. Fill in the configuration block

At the top of `Code.gs`, fill in every placeholder with your real values:

```js
const Label_Invoice        = "your-invoice-label";
const Label_Invoice_Search = "Your Invoice Label Display Name";
const Label_Invoice_Done   = "Your Invoice Done Label";
const Label_Timesheet      = "Your Timesheet Label";
const Label_Expense        = "Your Expense Label";

const FileId_InvoiceFolder        = "your-invoice-folder-id";
const FileId_InvoiceArchiveFolder = "your-invoice-archive-folder-id";
const FileId_TimesheetFolder      = "your-timesheet-folder-id";
const FileId_ExpenseFolder        = "your-expense-folder-id";

const FileId_Tracker = "your-tracker-sheet-id";

const Reviewers = ["Alice", "Bob", "Carol", "Dan"];
```

The `Reviewers` array controls how many approval columns appear. Add or remove names freely — the sheet headers and the web UI both update automatically.

---

### 6. Enable the Drive API

The timesheet function converts `.xlsx` files to PDF using the Drive API v2.

1. In Apps Script, go to **Services** (the **+** button in the left sidebar)
2. Find **Drive API** and add it
3. Make sure the version is set to **v2**

---

### 7. Authorize the script

Run any function (e.g. `invoice_UploadToWeb`) from the Apps Script editor. Google will prompt you to authorize the script to access Gmail, Drive, and Sheets. Follow the prompts and accept.

---

### 8. Deploy as a Web App

1. In Apps Script, click **Deploy > New Deployment**
2. Set type to **Web App**
3. Set **Execute as**: Me
4. Set **Who has access**: Anyone within your organization (or "Anyone" for external reviewers)
5. Click **Deploy** and copy the Web App URL
6. Share that URL with your reviewers

> **Re-deploying after code changes:** Each time you update `Code.gs` or `Index.html`, go to **Deploy > Manage Deployments**, click the pencil icon, choose "New version", and click **Deploy**. The URL stays the same.

---

### 9. Set up triggers (optional but recommended)

To run the Gmail → Drive sync automatically:

1. In Apps Script, go to **Triggers** (clock icon in the left sidebar)
2. Click **Add Trigger**
3. Set function to `invoice_GmailToDrive`
4. Set event source to **Time-driven**
5. Choose your preferred frequency (e.g. every hour)

Repeat for `timesheet_GmailToDrive` if needed.

> **Note:** Triggers always run under the account of whoever created them. If ownership of the spreadsheet transfers, the trigger owner must be updated accordingly.

---

## Usage

### Running the pipeline manually

From the Apps Script editor, run these functions in order:

| Function | What it does |
|---|---|
| `invoice_GmailToDrive()` | Scans Gmail, saves new PDFs to Drive, then calls `invoice_UploadToWeb()` |
| `invoice_UploadToWeb()` | Syncs Drive folder contents into the Invoices sheet |
| `timesheet_GmailToDrive()` | Converts `.xlsx` attachments to PDF and saves to Drive |
| `timesheet_UploadToWeb()` | Syncs Drive folder contents into the Timesheets sheet |

### Reviewing documents

Open the Web App URL. Click any row to load the PDF in the right-hand viewer. Use the dropdowns to set your approval status. Comments auto-save when you click away.

### Archiving approved records

At the end of the week, run:

```
invoice_ClearToRecords()
timesheet_ClearToRecords()
```

Rows where at least one reviewer has set "Approved" are moved to the `_Records` sheet with a "Week Ending" date appended. Rows with any other status (Hold, Rejected, blank) remain in the active sheet.

### Downloading approved invoices

Run `invoice_DownloadPDFs()` from the script editor or attach it to a custom menu. It will:
1. Copy all approved PDFs into a new timestamped Drive folder
2. Move the originals to the archive folder
3. Log a shareable link to the download folder

---

## How duplicate detection works

`invoice_GmailToDrive()` computes an MD5 hash of every file already in the Drive folder at runtime, then compares each new attachment's hash before saving. Files with matching content are skipped even if the filename differs. This prevents the same invoice from being saved twice if an email is re-processed.

---

## Notes

- `timesheet_GmailToDrive()` does not currently check for duplicate files. If you run it more than once on the same label, attachments will be saved again. Add a hash check (see `invoice_GmailToDrive()` for the pattern) if you plan to run it on a recurring trigger.
- The web UI auto-refreshes every 60 seconds but pauses if a reviewer is actively typing in a comment or using a dropdown, preventing lost input.
- The PDF URL column is intentionally hidden from the UI — it is stored on each table row as a `data-pdf` attribute and loaded into the iframe on click.

---

## License

MIT
