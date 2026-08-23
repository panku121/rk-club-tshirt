/**
 * RK Cricket Club - T-Shirt Order Form
 *
 * SETUP (Google Sheet "RK Cricket Club", tab name: rkCricketClub)
 *
 * 1. Open the Google Sheet: RK Cricket Club
 * 2. Go to Extensions > Apps Script
 * 3. Paste this file in full
 * 4. Save
 * 5. Deploy > New deployment
 *    - Type: Web app
 *    - Execute as: Me
 *    - Who has access: Anyone
 * 6. Copy the Web App URL and paste it into js/script.js
 *    as the value of GOOGLE_SCRIPT_URL
 */

var SHEET_NAME = "rkCricketClub";
var BACK_NUMBER_COL = 5;
var HEADERS = [
  "Timestamp",
  "Full Name",
  "Phone Number",
  "T-Shirt Size",
  "Back Number",
  "Back Name",
  "Sleeve Length"
];

function doGet(e) {
  return handleRequest(e);
}

function doPost(e) {
  return handleRequest(e);
}

function handleRequest(e) {
  var lock = LockService.getScriptLock();
  lock.waitLock(10000);

  try {
    var sheet = getOrCreateSheet();
    ensureHeaders(sheet);

    var data = getRequestData(e);
    if (!data.fullName && !data.phone) {
      return jsonResponse({ result: "ready", message: "RK Cricket Club form connected." });
    }

    var nextRow = Math.max(sheet.getLastRow() + 1, 2);
    var backNumber = String(data.backNumber || "");
    sheet.getRange(nextRow, 1, 1, HEADERS.length).setValues([[
      new Date(),
      data.fullName || "",
      data.phone || "",
      data.size || "",
      backNumber,
      data.backName || "",
      data.sleeveLength || ""
    ]]);
    sheet.getRange(nextRow, BACK_NUMBER_COL).setNumberFormat("@").setValue(backNumber);

    return jsonResponse({ result: "success" });
  } catch (error) {
    return jsonResponse({ result: "error", message: String(error) });
  } finally {
    lock.releaseLock();
  }
}

function getOrCreateSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
  }
  return sheet;
}

function ensureHeaders(sheet) {
  var firstRow = sheet.getRange(1, 1, 1, HEADERS.length).getValues()[0];
  var hasHeaders = firstRow.join("") !== "";
  if (!hasHeaders) {
    sheet.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]);
    sheet.getRange(1, 1, 1, HEADERS.length).setFontWeight("bold");
    sheet.setFrozenRows(1);
  }
  sheet.getRange(2, BACK_NUMBER_COL, sheet.getMaxRows() - 1, 1).setNumberFormat("@");
}

function getRequestData(e) {
  if (e && e.postData && e.postData.contents && e.postData.type === "application/json") {
    return JSON.parse(e.postData.contents);
  }

  var params = (e && e.parameter) ? e.parameter : {};
  return {
    fullName: params.fullName || "",
    phone: params.phone || "",
    size: params.size || "",
    backNumber: params.backNumber || "",
    backName: params.backName || "",
    sleeveLength: params.sleeveLength || ""
  };
}

function jsonResponse(payload) {
  return ContentService
    .createTextOutput(JSON.stringify(payload))
    .setMimeType(ContentService.MimeType.JSON);
}
