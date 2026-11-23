const SHEET_NAME = "Downloads";

/**
 * Google Apps Script backend for SATKARTAR STUDY LIBRARY.
 * Create a Google Sheet, open Extensions > Apps Script, paste this code,
 * save, then Deploy > New deployment > Web app.
 * Execute as: Me
 * Who has access: Anyone
 * Copy the /exec URL into DOWNLOAD_LOG_ENDPOINT in index.html.
 */
function setup() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) sheet = ss.insertSheet(SHEET_NAME);
  if (sheet.getLastRow() === 0) {
    sheet.appendRow([
      "Timestamp","Student Name","Class","Board","Mobile",
      "State","City","Country","PDF Name","PDF URL"
    ]);
    sheet.setFrozenRows(1);
  }
}

function doPost(e) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    sheet.appendRow([
      "Timestamp","Student Name","Class","Board","Mobile",
      "State","City","Country","PDF Name","PDF URL"
    ]);
    sheet.setFrozenRows(1);
  }

  let data = {};
  try {
    data = JSON.parse((e && e.postData && e.postData.contents) || "{}");
  } catch (err) {}

  sheet.appendRow([
    new Date(),
    data.studentName || "",
    data.className || "",
    data.board || "",
    data.mobile || "",
    data.state || "",
    data.city || "",
    data.country || "",
    data.pdfName || "",
    data.pdfUrl || ""
  ]);

  return ContentService
    .createTextOutput(JSON.stringify({ok:true}))
    .setMimeType(ContentService.MimeType.JSON);
}

function doGet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAME);
  const total = sheet ? Math.max(0, sheet.getLastRow() - 1) : 0;
  return ContentService
    .createTextOutput(JSON.stringify({ok:true, downloadRequests:total}))
    .setMimeType(ContentService.MimeType.JSON);
}
