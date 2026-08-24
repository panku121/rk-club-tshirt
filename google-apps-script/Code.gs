/**
 * RK Cricket Club - T-Shirt Order Form
 *
 * IMPORTANT: File save karna kaafi nahi hai.
 * Purani Web App URL old code chalaati hai, isliye naya version deploy karna zaroori hai:
 *
 * 1. Open the Google Sheet: RK Cricket Club
 * 2. Extensions > Apps Script
 * 3. Is file ko FULL paste karke Save karo
 * 4. Deploy > Manage deployments
 * 5. Existing Web app par pencil (Edit) click karo
 * 6. Version: New version
 * 7. Deploy
 * 8. Web App URL same rahe to js/script.js mat badlo
 *
 * Agar New deployment banaya ho aur URL change ho gayi ho,
 * nayi URL js/script.js mein GOOGLE_SCRIPT_URL pe paste karo.
 */

var SHEET_NAME = "rkCricketClub";
var BACK_NUMBER_COL = 5;
var DEFAULT_PAYMENT_STATUS = "Not Paid";
var HEADERS = [
  "Timestamp",
  "Full Name",
  "Phone Number",
  "T-Shirt Size",
  "Back Number",
  "Back Name",
  "Sleeve Length",
  "Order Package",
  "Payment Status"
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
    backfillPaymentStatus(sheet);

    var data = getRequestData(e);
    if (!data.fullName && !data.phone) {
      return jsonResponse({ result: "ready", message: "RK Cricket Club form connected." });
    }

    writeOrderRow(sheet, data);
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
  var headerMismatch = HEADERS.some(function (header, index) {
    return String(firstRow[index] || "").trim() !== header;
  });
  if (headerMismatch) {
    sheet.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]);
    sheet.getRange(1, 1, 1, HEADERS.length).setFontWeight("bold");
    sheet.setFrozenRows(1);
  }
  sheet.getRange(2, BACK_NUMBER_COL, sheet.getMaxRows() - 1, 1).setNumberFormat("@");
}

function writeOrderRow(sheet, data) {
  var nextRow = Math.max(sheet.getLastRow() + 1, 2);
  var backNumber = String(data.backNumber || "");
  var orderPackage = data.orderPackage || "";
  var rowValues = [
    new Date(),
    data.fullName || "",
    data.phone || "",
    data.size || "",
    backNumber,
    data.backName || "",
    data.sleeveLength || "",
    orderPackage,
    DEFAULT_PAYMENT_STATUS
  ];

  sheet.getRange(nextRow, 1, 1, rowValues.length).setValues([rowValues]);
  sheet.getRange(nextRow, BACK_NUMBER_COL).setNumberFormat("@").setValue(backNumber);
  setCellByHeader(sheet, nextRow, "Order Package", orderPackage);
  setCellByHeader(sheet, nextRow, "Payment Status", DEFAULT_PAYMENT_STATUS);
}

function setCellByHeader(sheet, row, headerName, value) {
  var lastColumn = Math.max(sheet.getLastColumn(), HEADERS.length);
  var headers = sheet.getRange(1, 1, 1, lastColumn).getValues()[0];
  var col = -1;
  for (var i = 0; i < headers.length; i++) {
    if (String(headers[i] || "").trim() === headerName) {
      col = i + 1;
      break;
    }
  }
  if (col < 1) {
    col = HEADERS.indexOf(headerName) + 1;
  }
  if (col > 0) {
    sheet.getRange(row, col).setValue(value);
  }
}

function backfillPaymentStatus(sheet) {
  var lastRow = sheet.getLastRow();
  var paymentCol = columnForHeader(sheet, "Payment Status");
  if (lastRow < 2 || paymentCol < 1) return;

  var range = sheet.getRange(2, paymentCol, lastRow - 1, 1);
  var values = range.getValues();
  var changed = false;
  for (var i = 0; i < values.length; i++) {
    if (!String(values[i][0] || "").trim()) {
      values[i][0] = DEFAULT_PAYMENT_STATUS;
      changed = true;
    }
  }
  if (changed) {
    range.setValues(values);
  }
}

function columnForHeader(sheet, headerName) {
  var lastColumn = Math.max(sheet.getLastColumn(), HEADERS.length);
  var headers = sheet.getRange(1, 1, 1, lastColumn).getValues()[0];
  for (var i = 0; i < headers.length; i++) {
    if (String(headers[i] || "").trim() === headerName) {
      return i + 1;
    }
  }
  return HEADERS.indexOf(headerName) + 1;
}

function getRequestData(e) {
  var params = {};
  copyFields(params, e && e.parameter);
  copyFields(params, firstValues(e && e.parameters));

  if (e && e.postData && e.postData.contents) {
    var type = String(e.postData.type || "");
    if (type.indexOf("json") !== -1 || type.indexOf("text/plain") !== -1) {
      try {
        copyFields(params, JSON.parse(e.postData.contents));
      } catch (err) {}
    }
  }

  return {
    fullName: params.fullName || "",
    phone: params.phone || "",
    size: params.size || "",
    backNumber: params.backNumber || "",
    backName: params.backName || "",
    sleeveLength: params.sleeveLength || "",
    orderPackage: params.orderPackage || params.packageOption || ""
  };
}

function copyFields(target, source) {
  if (!source) return;
  Object.keys(source).forEach(function (key) {
    var value = source[key];
    if (value === undefined || value === null || value === "") return;
    target[key] = value;
  });
}

function firstValues(parameters) {
  if (!parameters) return {};
  var result = {};
  Object.keys(parameters).forEach(function (key) {
    var value = parameters[key];
    result[key] = Array.isArray(value) ? value[0] : value;
  });
  return result;
}

function jsonResponse(payload) {
  return ContentService
    .createTextOutput(JSON.stringify(payload))
    .setMimeType(ContentService.MimeType.JSON);
}
