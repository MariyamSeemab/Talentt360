// =============================================================
//  TALENTS360 – Google Apps Script Backend
//  Receives video entry submissions from the HTML form,
//  saves metadata to Google Sheets, and uploads the video
//  file to a Google Drive folder.
// =============================================================
//
//  ── HOW TO SET UP ──────────────────────────────────────────
//
//  STEP 1 – Open Google Apps Script
//    • Go to https://script.google.com
//    • Click "New project"
//    • Delete any existing code and paste this entire file.
//
//  STEP 2 – Create a Google Drive folder for videos
//    • Open Google Drive → New → Folder
//    • Name it e.g. " "
//    • Open the folder → copy the ID from the URL:
//        https://drive.google.com/drive/folders/  ← FOLDER_ID is here
//    • Paste it as the value of DRIVE_FOLDER_ID below.
//
//  STEP 3 – Create a Google Sheet for entries
//    • Open Google Sheets → create a new blank spreadsheet
//    • Name the first sheet tab "Entries" (or leave as "Sheet1")
//    • Copy the spreadsheet ID from its URL:
//        https://docs.google.com/spreadsheets/d/  ← SHEET_ID is here
//    • Paste it as the value of SHEET_ID below.
//
//  STEP 4 – Fill in the config values below, then save (Ctrl+S).
//
//  STEP 5 – Deploy as Web App
//    • Click Deploy → New deployment
//    • Type: Web App
//    • Description: "Talents360 v1" (anything you like)
//    • Execute as: Me
//    • Who has access: Anyone
//    • Click Deploy → Authorize → Copy the Web App URL
//
//  STEP 6 – Paste URL into the HTML file
//    • Open talents360-v3.html
//    • Find:  var SCRIPT_URL = "PASTE_YOUR_APPS_SCRIPT_URL";
//    • Replace the placeholder with the URL you just copied.
//
//  NOTE: Every time you edit this script you must create a
//  NEW deployment (or update the existing one) for changes
//  to take effect on the live URL.
// =============================================================


// ── CONFIG – fill these two values ──────────────────────────
var DRIVE_FOLDER_ID = "18qid38C1YJTORUwJjaeHJf1QI-YCYKiv";   // Google Drive folder ID
var SHEET_ID        = "1mGnYKmawt1Sg1doGsjo_4KHsN_JrsWnk4BZ5CLUceo4";   // Google Sheet ID
var SHEET_TAB       = "Entries";                       // Sheet tab name
// ────────────────────────────────────────────────────────────


/**
 * Handles GET requests – simple health-check so you can test
 * the deployment is live by opening the URL in a browser.
 */
function doGet(e) {
  return ContentService
    .createTextOutput("Talents360 backend is running ✅")
    .setMimeType(ContentService.MimeType.TEXT);
}


/**
 * Handles POST requests sent by the HTML form.
 * Expected JSON body:
 *   { name, category, school, whatsapp, filename, mime, data }
 *   where `data` is the base64-encoded video file.
 */
function doPost(e) {
  try {
    // ── 1. Parse incoming JSON ─────────────────────────────
    var payload = JSON.parse(e.postData.contents);
    var name      = payload.name      || "";
    var category  = payload.category  || "";
    var school    = payload.school    || "";
    var whatsapp  = payload.whatsapp  || "";
    var filename  = payload.filename  || ("video_" + Date.now());
    var mime      = payload.mime      || "video/mp4";
    var b64data   = payload.data      || "";

    // ── 2. Basic server-side validation ───────────────────
    if (!name || !category || !school || !whatsapp || !b64data) {
      return jsonResponse({ status: "error", message: "Missing required fields." });
    }

    // ── 3. Upload video to Google Drive ───────────────────
    var folder    = DriveApp.getFolderById(DRIVE_FOLDER_ID);
    var blob      = Utilities.newBlob(
      Utilities.base64Decode(b64data),
      mime,
      filename
    );
    var driveFile = folder.createFile(blob);
    driveFile.setSharing(
      DriveApp.Access.ANYONE_WITH_LINK,
      DriveApp.Permission.VIEW
    );
    var fileUrl   = driveFile.getUrl();
    var fileId    = driveFile.getId();

    // ── 4. Log entry to Google Sheets ─────────────────────
    var sheet = SpreadsheetApp
      .openById(SHEET_ID)
      .getSheetByName(SHEET_TAB);

    // Add header row automatically on first submission
    if (sheet.getLastRow() === 0) {
      sheet.appendRow([
        "Timestamp",
        "Full Name",
        "Category",
        "School / Institution",
        "WhatsApp Number",
        "Video Filename",
        "Drive File ID",
        "Drive Link"
      ]);
      // Bold the header row
      sheet.getRange(1, 1, 1, 8).setFontWeight("bold");
    }

    sheet.appendRow([
      new Date(),
      name,
      category,
      school,
      whatsapp,
      filename,
      fileId,
      fileUrl
    ]);

    // ── 5. Return success ─────────────────────────────────
    return jsonResponse({ status: "ok", message: "Entry saved successfully.", fileId: fileId });

  } catch (err) {
    // Log the error for debugging in Apps Script → Executions
    console.error("Talents360 doPost error:", err);
    return jsonResponse({ status: "error", message: err.toString() });
  }
}


/**
 * Helper – returns a JSON ContentService response.
 * Note: with mode:'no-cors' on the client the response body
 * is opaque, but Apps Script still needs to return something.
 */
function jsonResponse(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
