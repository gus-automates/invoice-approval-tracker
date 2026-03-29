// ============================================================
// CONFIGURATION — fill these in before running
// ============================================================

// Gmail label names (must match exactly what you see in Gmail)
const Label_Invoice        = "YOUR_INVOICE_LABEL";           // used in search queries
const Label_Invoice_Search = "YOUR_INVOICE_LABEL_DISPLAY";   // used with getUserLabelByName
const Label_Invoice_Done   = "YOUR_INVOICE_DONE_LABEL";      // applied after processing
const Label_Timesheet      = "YOUR_TIMESHEET_LABEL";
const Label_Expense        = "YOUR_EXPENSE_LABEL";

// Google Drive folder IDs (find these in the folder's URL)
const FileId_InvoiceFolder        = "YOUR_INVOICE_FOLDER_ID";
const FileId_InvoiceArchiveFolder = "YOUR_INVOICE_ARCHIVE_FOLDER_ID"; // approved files move here after download
const FileId_TimesheetFolder      = "YOUR_TIMESHEET_FOLDER_ID";
const FileId_ExpenseFolder        = "YOUR_EXPENSE_FOLDER_ID";

// Google Sheets spreadsheet ID (find this in the sheet's URL)
const FileId_Tracker = "YOUR_TRACKER_SHEET_ID";

// Sheet tab names (must match the tabs inside your spreadsheet)
const Sheet_Invoices          = "Invoices";
const Sheet_InvoicesRecords   = "Invoices_Records";
const Sheet_Timesheets        = "Timesheets";
const Sheet_TimesheetsRecords = "Timesheets_Records";
const Sheet_Expenses          = "Expenses";
const Sheet_ExpensesRecords   = "Expenses_Records";

// Reviewer names — these become column headers in the approval sheet.
// Add or remove names here; everything else adapts automatically.
const Reviewers = ["Reviewer_1", "Reviewer_2", "Reviewer_3", "Reviewer_4"];

// ============================================================
// HELPERS
// ============================================================

/**
 * Always returns the tracker spreadsheet by ID.
 * Safe to call from both container-bound and Web App contexts —
 * getActiveSpreadsheet() returns null when running as a deployed Web App,
 * so we always use openById() instead.
 */
function getTracker_() {
  return SpreadsheetApp.openById(FileId_Tracker);
}

/*
--------------------------------------------------INVOICE--------------------------------------------------
*/

function invoice_GmailToDrive() {
  const folder = DriveApp.getFolderById(FileId_InvoiceFolder);

  // Build a hash set of files already in the folder to skip duplicates
  const existingFileHashes = {};
  const existingFiles = folder.getFiles();
  while (existingFiles.hasNext()) {
    const blob = existingFiles.next().getBlob();
    const hash = computeMd5_(blob.getBytes());
    existingFileHashes[hash] = true;
  }

  // Get or create the "done" label (applied to threads after processing)
  let doneLabel = GmailApp.getUserLabelByName(Label_Invoice_Done);
  if (!doneLabel) doneLabel = GmailApp.createLabel(Label_Invoice_Done);

  // Source label object needed to remove it from threads
  const sourceLabel = GmailApp.getUserLabelByName(Label_Invoice_Search);

  // Paginate through all matching threads (GmailApp.search returns max 100 at a time)
  let page = 0;
  let threads;
  do {
    threads = GmailApp.search("label:" + Label_Invoice, page * 100, 100);

    threads.forEach(thread => {
      let threadHadPDF = false;

      thread.getMessages().forEach(message => {
        message.getAttachments().forEach(attachment => {
          const name        = attachment.getName();
          const contentType = attachment.getContentType();
          const isPDF       = contentType.includes("pdf") || name.toLowerCase().endsWith(".pdf");

          if (!isPDF) return;

          const hash = computeMd5_(attachment.getBytes());
          if (!existingFileHashes[hash]) {
            folder.createFile(attachment);
            Logger.log("Saved: " + name);
            existingFileHashes[hash] = true;
          } else {
            Logger.log("Skipped (duplicate): " + name);
          }
          // Mark thread for relabelling regardless — it was already processed
          threadHadPDF = true;
        });
      });

      if (threadHadPDF) {
        thread.addLabel(doneLabel);
        thread.removeLabel(sourceLabel);
        Logger.log("Thread relabeled: " + thread.getFirstMessageSubject());
      }
    });

    page++;
  } while (threads.length === 100);

  invoice_UploadToWeb();
  Logger.log("invoice_GmailToDrive complete.");
}

function invoice_UploadToWeb() {
  const sheet = getTracker_().getSheetByName(Sheet_Invoices);
  if (!sheet) {
    Logger.log("Sheet not found: " + Sheet_Invoices);
    return;
  }

  const headers = ["File Name", "PDF URL", ...Reviewers, "Comments"];
  ensureHeaders_(sheet, headers);

  // Read all existing URLs in one batch call (column 2) to de-duplicate
  const lastRow    = sheet.getLastRow();
  const existingUrls = lastRow > 1
    ? new Set(sheet.getRange(2, 2, lastRow - 1, 1).getValues().flat())
    : new Set();

  const folder   = DriveApp.getFolderById(FileId_InvoiceFolder);
  const files    = folder.getFilesByType(MimeType.PDF);
  const newRows  = [];

  while (files.hasNext()) {
    const file   = files.next();
    const fileUrl = "https://drive.google.com/file/d/" + file.getId() + "/preview";
    if (!existingUrls.has(fileUrl)) {
      // Row: [File Name, PDF URL, ...blank reviewer cells, blank Comments]
      newRows.push([file.getName(), fileUrl, ...Reviewers.map(() => ""), ""]);
    }
  }

  // Write all new rows in a single setValues call instead of one appendRow per file
  if (newRows.length > 0) {
    const startRow = sheet.getLastRow() + 1;
    sheet.getRange(startRow, 1, newRows.length, headers.length).setValues(newRows);
  }

  // Sort all data rows A–Z by file name
  const finalLastRow = sheet.getLastRow();
  if (finalLastRow > 2) {
    sheet.getRange(2, 1, finalLastRow - 1, headers.length).sort({ column: 1, ascending: true });
  }

  Logger.log(newRows.length + " new PDF(s) added to Invoices sheet.");
}

function invoice_ClearToRecords() {
  clearToRecords_(Sheet_Invoices, Sheet_InvoicesRecords);
}

function invoice_DownloadPDFs() {
  const sheet       = getTracker_().getSheetByName(Sheet_Invoices);
  const lastRow     = sheet.getLastRow();

  if (lastRow < 2) {
    SpreadsheetApp.getUi().alert("No data found.");
    return;
  }

  // Read all data columns dynamically — no hardcoded column range
  const totalCols   = sheet.getLastColumn();
  const data        = sheet.getRange(2, 1, lastRow - 1, totalCols).getValues();

  // Column indices (0-based): 0=File Name, 1=PDF URL, 2..N-1=Reviewers, N=Comments
  const urlColIndex = 1;
  const reviewerStart = 2;
  const reviewerEnd   = totalCols - 2; // last col is Comments

  const allowedFileIds = [];

  for (const row of data) {
    const previewUrl = row[urlColIndex];
    if (!previewUrl) continue;

    // Only include rows where at least one reviewer has set "Approved"
    const statuses   = row.slice(reviewerStart, reviewerEnd + 1).map(c => (c || "").toString().toLowerCase());
    const hasApproved = statuses.includes("approved");

    if (hasApproved) {
      const fileId = extractFileIdFromPreviewUrl_(previewUrl);
      if (fileId) allowedFileIds.push(fileId);
    }
  }

  // Exit early — no point creating an empty folder or a public link
  if (allowedFileIds.length === 0) {
    SpreadsheetApp.getUi().alert("No approved invoices found. Nothing to download.");
    return;
  }

  // Create a timestamped download folder in Drive root
  const timestamp    = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd_HH-mm");
  const outputFolder = DriveApp.getRootFolder().createFolder("Download_Invoices_" + timestamp);

  const sourceFolder = DriveApp.getFolderById(FileId_InvoiceFolder);
  const destFolder   = DriveApp.getFolderById(FileId_InvoiceArchiveFolder);

  for (const fileId of allowedFileIds) {
    try {
      const file = DriveApp.getFileById(fileId);
      file.makeCopy(file.getName(), outputFolder); // copy into download folder
      destFolder.addFile(file);                    // move to archive: add first…
      sourceFolder.removeFile(file);               // …then remove from source
    } catch (e) {
      Logger.log("Error processing fileId " + fileId + ": " + e.message);
    }
  }

  // Make the download folder accessible to anyone with the link
  outputFolder.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  Logger.log("Download folder: " + outputFolder.getUrl());
  SpreadsheetApp.getUi().alert(
    allowedFileIds.length + " invoice(s) downloaded.\n\nFolder link:\n" + outputFolder.getUrl()
  );
}

/*
--------------------------------------------------TIMESHEET--------------------------------------------------
*/

function timesheet_GmailToDrive() {
  const folder = DriveApp.getFolderById(FileId_TimesheetFolder);
  const label  = GmailApp.getUserLabelByName(Label_Timesheet);

  if (!label) {
    throw new Error(`Label '${Label_Timesheet}' not found. Check Gmail for the correct label name.`);
  }

  // NOTE: Unlike invoice_GmailToDrive(), this function does not currently check for
  // duplicate files before saving. If you run it more than once on the same label,
  // attachments will be saved again. Add de-duplication logic (MD5 hash check) if
  // this function will be run on a recurring trigger.

  const threads = label.getThreads();

  for (const thread of threads) {
    for (const message of thread.getMessages()) {
      const senderName = message.getFrom().split('<')[0].trim();

      for (const attachment of message.getAttachments()) {
        if (attachment.getContentType() !== MimeType.MICROSOFT_EXCEL) continue;

        const originalFileName  = attachment.getName().replace(/\.[^/.]+$/, '');
        const fileNameWithSender = `${senderName} - ${originalFileName}`;

        // Save .xlsx temporarily to Drive so we can convert it
        const tempExcelFile = folder.createFile(attachment);
        tempExcelFile.setName(fileNameWithSender + '.xlsx');

        // Convert .xlsx → Google Sheet (needed to use the Sheets PDF export API)
        const sheetFile = Drive.Files.insert(
          {
            title: fileNameWithSender,
            mimeType: MimeType.GOOGLE_SHEETS,
            parents: [{ id: FileId_TimesheetFolder }],
          },
          attachment.getAs(MimeType.MICROSOFT_EXCEL)
        );

        const exportUrl = `https://docs.google.com/spreadsheets/d/${sheetFile.id}/export?`;
        const exportOptions = {
          format: 'pdf', portrait: false, fitw: true, scale: 4, size: 'A4',
          top_margin: 0.25, bottom_margin: 0.25, left_margin: 0.25, right_margin: 0.25,
          sheetnames: false, printtitle: false, pagenum: 'UNDEFINED',
          gridlines: false, fzr: false,
        };

        const queryString = Object.entries(exportOptions).map(([k, v]) => `${k}=${v}`).join('&');
        const response    = UrlFetchApp.fetch(exportUrl + queryString, {
          headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() },
        });

        // Save the rendered PDF
        const pdfFile = folder.createFile(response.getBlob());
        pdfFile.setName(fileNameWithSender + '.pdf');
        Logger.log("Saved: " + pdfFile.getName());

        // Clean up the temporary Google Sheet and original .xlsx
        DriveApp.getFileById(sheetFile.id).setTrashed(true);
        tempExcelFile.setTrashed(true);
      }
    }
  }

  SpreadsheetApp.getUi().alert("All new PDFs saved to Timesheet Folder.");
}

function timesheet_UploadToWeb() {
  const sheet = getTracker_().getSheetByName(Sheet_Timesheets);
  if (!sheet) {
    SpreadsheetApp.getUi().alert("Sheet not found: " + Sheet_Timesheets);
    return;
  }

  const headers = ["Name", "PDF URL", ...Reviewers, "Comments"];
  ensureHeaders_(sheet, headers);

  // Read all existing URLs in one batch call to de-duplicate
  const lastRow    = sheet.getLastRow();
  const existingUrls = lastRow > 1
    ? new Set(sheet.getRange(2, 2, lastRow - 1, 1).getValues().flat())
    : new Set();

  const folder  = DriveApp.getFolderById(FileId_TimesheetFolder);
  const files   = folder.getFilesByType(MimeType.PDF);
  const newRows = [];

  while (files.hasNext()) {
    const file    = files.next();
    const fileUrl = "https://drive.google.com/file/d/" + file.getId() + "/preview";
    if (!existingUrls.has(fileUrl)) {
      newRows.push([file.getName(), fileUrl, ...Reviewers.map(() => ""), ""]);
    }
  }

  // Write all new rows in a single setValues call
  if (newRows.length > 0) {
    const startRow = sheet.getLastRow() + 1;
    sheet.getRange(startRow, 1, newRows.length, headers.length).setValues(newRows);
  }

  const finalLastRow = sheet.getLastRow();
  if (finalLastRow > 2) {
    sheet.getRange(2, 1, finalLastRow - 1, headers.length).sort({ column: 1, ascending: true });
  }

  SpreadsheetApp.getUi().alert(newRows.length + " PDF(s) added!");
}

function timesheet_ClearToRecords() {
  clearToRecords_(Sheet_Timesheets, Sheet_TimesheetsRecords);
}

/*
--------------------------------------------------TOOLS--------------------------------------------------
*/

/**
 * Ensures a sheet has the correct headers in row 1.
 * If the sheet is empty, writes them fresh.
 * If headers are missing or mismatched, rewrites row 1 to match.
 * Trailing underscore marks this as an internal/helper function.
 */
function ensureHeaders_(sheet, expectedHeaders) {
  const lr = sheet.getLastRow();
  if (lr === 0) {
    sheet.appendRow(expectedHeaders);
    return;
  }

  const current = sheet.getRange(1, 1, 1, Math.max(sheet.getLastColumn(), expectedHeaders.length)).getValues()[0];
  const trimmed = current.map(h => (h || "").toString().trim());

  const alreadyCorrect =
    trimmed.length === expectedHeaders.length &&
    trimmed.every((h, i) => h === expectedHeaders[i]);

  if (!alreadyCorrect) {
    sheet.getRange(1, 1, 1, expectedHeaders.length).setValues([expectedHeaders]);
  }
}

/**
 * Moves "Approved" rows from the source sheet to the archive sheet,
 * appending a "Week Ending" date column. All other rows stay in the source.
 * Trailing underscore marks this as an internal/helper function.
 *
 * @param {string} sheet_source  - Name of the active sheet  (e.g. "Invoices")
 * @param {string} sheet_archive - Name of the records sheet (e.g. "Invoices_Records")
 */
function clearToRecords_(sheet_source, sheet_archive) {
  const ss           = getTracker_();
  const sourceSheet  = ss.getSheetByName(sheet_source);
  const archiveSheet = ss.getSheetByName(sheet_archive);

  if (!sourceSheet || !archiveSheet) {
    throw new Error("Sheet not found: '" + sheet_source + "' or '" + sheet_archive + "'.");
  }

  const sourceLastRow = sourceSheet.getLastRow();
  const sourceLastCol = sourceSheet.getLastColumn();

  if (sourceLastRow < 2) {
    SpreadsheetApp.getUi().alert("No data to move.");
    return;
  }

  // Read source headers
  const sourceHeaders = sourceSheet
    .getRange(1, 1, 1, sourceLastCol).getValues()[0]
    .map(h => (h || "").toString().trim());

  // Set up archive headers — adds "Week Ending" column if not already present
  let archiveLastCol  = Math.max(archiveSheet.getLastColumn(), 1);
  let archiveHeaders  = [];

  if (archiveSheet.getLastRow() === 0) {
    archiveHeaders = [...sourceHeaders, "Week Ending"];
    archiveSheet.getRange(1, 1, 1, archiveHeaders.length).setValues([archiveHeaders]);
    archiveLastCol = archiveHeaders.length;
  } else {
    archiveHeaders = archiveSheet
      .getRange(1, 1, 1, archiveLastCol).getValues()[0]
      .map(h => (h || "").toString().trim());

    // Add "Week Ending" header if it's missing
    if (archiveHeaders.length === sourceHeaders.length) {
      archiveHeaders.push("Week Ending");
      archiveSheet.getRange(1, 1, 1, archiveHeaders.length).setValues([archiveHeaders]);
      archiveLastCol = archiveHeaders.length;
    }
  }

  const data       = sourceSheet.getRange(2, 1, sourceLastRow - 1, sourceLastCol).getValues();
  const fridayDate = getLastWeeksFriday_();
  const rowsToMove = [];
  const rowsToKeep = [];

  for (const row of data) {
    // Move row only if at least one cell is exactly "approved" (case-insensitive)
    const hasApproved = row.some(cell => (cell || "").toString().toLowerCase() === "approved");

    if (hasApproved) {
      // Build archive row: source values + "Week Ending" appended at the end
      const newRow = archiveHeaders.map((header, i) => {
        if (i < sourceHeaders.length) return row[i] !== undefined ? row[i] : "";
        if (header === "Week Ending")  return fridayDate;
        return "";
      });
      rowsToMove.push(newRow);
    } else {
      rowsToKeep.push(row); // blank, Hold, Rejected, Pending — stays in source
    }
  }

  // Append approved rows to archive in one batch write
  if (rowsToMove.length > 0) {
    const startRow = archiveSheet.getLastRow() + 1;
    archiveSheet.getRange(startRow, 1, rowsToMove.length, archiveHeaders.length).setValues(rowsToMove);
  }

  // Clear source data rows and rewrite only the kept rows
  sourceSheet.getRange(2, 1, Math.max(sourceLastRow - 1, 1), sourceLastCol).clearContent();
  if (rowsToKeep.length > 0) {
    sourceSheet.getRange(2, 1, rowsToKeep.length, sourceLastCol).setValues(rowsToKeep);
  }

  // Sort remaining source rows by column A
  const newLastRow = sourceSheet.getLastRow();
  if (newLastRow > 2) {
    sourceSheet.getRange(2, 1, newLastRow - 1, sourceLastCol).sort({ column: 1, ascending: true });
  }

  SpreadsheetApp.getUi().alert(
    `${rowsToMove.length} row(s) moved to records.\n${rowsToKeep.length} row(s) kept (not yet approved).`
  );
}

/**
 * Returns last week's Friday as a formatted string (e.g. "Mar-21").
 * Used to stamp archive rows with their week-ending date.
 * Trailing underscore marks this as an internal/helper function.
 */
function getLastWeeksFriday_() {
  const today        = new Date();
  const dayOfWeek    = today.getDay(); // Sunday=0 … Saturday=6
  const daysToFriday = (5 - dayOfWeek + 7) % 7;

  const thisFriday = new Date(today);
  thisFriday.setDate(today.getDate() + daysToFriday);

  const lastFriday = new Date(thisFriday);
  lastFriday.setDate(thisFriday.getDate() - 7);

  return Utilities.formatDate(lastFriday, Session.getScriptTimeZone(), "MMM-dd");
}

/**
 * Computes an MD5 hash for a byte array and returns it as a hex string.
 * Used to detect duplicate files before saving to Drive.
 * Trailing underscore marks this as an internal/helper function.
 *
 * @param {number[]} bytes - Raw file bytes
 * @returns {string} Hex MD5 string
 */
function computeMd5_(bytes) {
  return Utilities.computeDigest(Utilities.DigestAlgorithm.MD5, bytes)
    .map(b => ('0' + (b & 0xFF).toString(16)).slice(-2))
    .join('');
}

/**
 * Extracts the Drive file ID from a /preview URL.
 * Example input:  "https://drive.google.com/file/d/FILE_ID/preview"
 * Example output: "FILE_ID"
 * Trailing underscore marks this as an internal/helper function.
 *
 * @param {string} url
 * @returns {string|null}
 */
function extractFileIdFromPreviewUrl_(url) {
  const match = url.match(/\/d\/([^/]+)\/preview/);
  return match ? match[1] : null;
}

/*
--------------------------------------------------WEB APP--------------------------------------------------
*/

/**
 * Serves the HTML approval UI when deployed as a Web App.
 * Deploy via: Extensions > Apps Script > Deploy > New Deployment > Web App
 */
function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle("Approval Tracker")
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/**
 * Returns all data rows from a named sheet (header row excluded),
 * plus the header row itself as the first element so the frontend
 * can build column headers dynamically.
 *
 * Called from the frontend via: google.script.run.getSheetData(sheetName)
 *
 * @param {string} sheetName - Tab name (e.g. "Invoices")
 * @returns {{ headers: string[], rows: any[][] }}
 */
function getSheetData(sheetName) {
  const sheet = getTracker_().getSheetByName(sheetName);
  if (!sheet) throw new Error("Sheet not found: " + sheetName);

  const data = sheet.getDataRange().getValues();
  return {
    headers: data[0] || [],  // row 1 — used by frontend to build <th> elements
    rows:    data.slice(1),  // row 2+ — the actual invoice/timesheet data
  };
}

/**
 * Writes a single value to a specific cell in a named sheet.
 * Called from the frontend when a reviewer changes a dropdown or comment.
 *
 * @param {string} sheetName - Tab name
 * @param {number} row       - 1-based row index
 * @param {number} col       - 1-based column index
 * @param {*}      value     - The value to write
 */
function updateCell(sheetName, row, col, value) {
  const sheet = getTracker_().getSheetByName(sheetName);
  if (!sheet) throw new Error("Sheet not found: " + sheetName);
  sheet.getRange(row, col).setValue(value);
}
