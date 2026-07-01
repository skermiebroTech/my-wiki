/**
 * Driver Installer - Analytics Webhook  (v1.13.2-compatible)
 *
 * To install:
 *   1. Open the Google Sheet that's currently receiving the analytics.
 *   2. Extensions > Apps Script.
 *   3. Delete everything in the editor and paste this file.
 *   4. Update SHEET_NAME below if your tab isn't called 'data'.
 *   5. Save (disk icon, top-left).
 *   6. Deploy > Manage deployments > select the existing deployment > Edit
 *      (pencil icon) > Version: New version > Deploy.
 *      (DO NOT create a new deployment - that would change the URL and
 *       break the SHEETS_WEBHOOK constant in Install-Drivers-auto.ps1.)
 *   7. Run a test from a machine; check Apps Script's Executions panel for
 *      "doPost" runs and confirm the row appears in the Sheet.
 *
 * If you see "ReferenceError: Cannot access 'data' before initialization" or
 * any HTML error response in the Driver Installer .log, the Apps Script
 * threw an exception. The Executions panel in the Apps Script editor will
 * show the full stack trace.
 *
 * Field list (v1.13.2):
 *   result, manufacturer, model, serial, os_version, os_build,
 *   inf_count, download_mb, missing_before, missing_after,
 *   duration_sec, script_version, installed_drivers,
 *   driver_urls, missing_after_list
 *
 * v1.13.2 schema change: the "missing_before_list" field (list of missing-
 * device descriptions BEFORE the install) was REMOVED from the PowerShell
 * payload and REPLACED with "driver_urls" - a comma-separated list of the
 * actual driver-file URLs the script downloaded this run (filtered to
 * driver payloads only; vendor catalog/descriptor/matrix URLs are excluded).
 * The corresponding column slot (#15) is repurposed in-place and renamed
 * "Driver URLs". Historical rows written before v1.13.2 will still show
 * their old missing-device descriptions in that column - rename/clear those
 * manually if the mixed semantics bother you.
 *
 * The handler is forward-compatible: any new field the PowerShell side
 * starts sending will arrive in `data` but will be silently ignored
 * unless you add it to HEADERS / appendRow below.
 */

// ----- CONFIG -----
const SHEET_NAME = 'data';   // change if your tab has a different name

// Column headers, in the exact order they appear in the appendRow call below.
// If the sheet is empty, this list is written as the first row automatically.
const HEADERS = [
  'Timestamp',
  'Result',
  'Manufacturer',
  'Model',
  'Serial',
  'OS Version',
  'OS Build',
  'INF Count',
  'Download MB',
  'Missing Before',
  'Missing After',
  'Duration (sec)',
  'Script Version',
  'Installed Drivers',
  'Driver URLs',              // v1.13.2: was 'Missing Before (list)'
  'Missing After (list)'
];


// ----- ENTRY POINT -----
function doPost(e) {
  try {
    // Defensive: e.postData may be undefined if someone hits the URL with a GET
    // or with no body. JSON.parse on undefined throws TypeError, which would
    // bubble up as an HTML error response - hence the explicit guard.
    if (!e || !e.postData || !e.postData.contents) {
      return _textResponse('ERROR: no POST body');
    }

    var data;
    try {
      data = JSON.parse(e.postData.contents);
    } catch (parseErr) {
      return _textResponse('ERROR: invalid JSON - ' + parseErr.message);
    }

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(SHEET_NAME);
    if (!sheet) {
      // Auto-create the tab on first run rather than failing silently.
      sheet = ss.insertSheet(SHEET_NAME);
    }

    // Header row on first use (or if someone cleared the sheet).
    if (sheet.getLastRow() === 0) {
      sheet.appendRow(HEADERS);
      sheet.getRange(1, 1, 1, HEADERS.length).setFontWeight('bold');
      sheet.setFrozenRows(1);
    }

    // Map JSON fields to columns. Order MUST match HEADERS above.
    // _num and _str helpers coerce missing/null values to sensible defaults
    // so a partial payload still writes a clean row.
    sheet.appendRow([
      new Date(),
      _str(data.result),
      _str(data.manufacturer),
      _str(data.model),
      _str(data.serial),
      _str(data.os_version),
      _num(data.os_build),
      _num(data.inf_count),
      _num(data.download_mb),
      _numOrBlank(data.missing_before),
      _numOrBlank(data.missing_after),
      _num(data.duration_sec),
      _str(data.script_version),
      _str(data.installed_drivers),
      _str(data.driver_urls),         // v1.13.2: was data.missing_before_list
      _str(data.missing_after_list)
    ]);

    return _textResponse('OK');

  } catch (err) {
    // Catch-all so the PowerShell client gets a PLAIN-TEXT error string
    // instead of an HTML error page. The PS script logs everything that
    // isn't exactly "OK" as a warning, but plain text is much easier to
    // read than the default Apps Script HTML.
    return _textResponse('ERROR: ' + (err && err.message ? err.message : String(err)));
  }
}


// ----- HELPERS -----
function _textResponse(s) {
  return ContentService
    .createTextOutput(s)
    .setMimeType(ContentService.MimeType.TEXT);
}

function _str(v) {
  if (v === undefined || v === null) return '';
  return String(v);
}

function _num(v) {
  if (v === undefined || v === null || v === '') return 0;
  var n = Number(v);
  return isNaN(n) ? 0 : n;
}

// Same as _num but returns '' instead of 0 for missing values - keeps the
// "missing_before / missing_after" columns visually empty when the script
// couldn't determine the count (sentinel value -1 in the payload).
function _numOrBlank(v) {
  if (v === undefined || v === null || v === '') return '';
  var n = Number(v);
  if (isNaN(n)) return '';
  if (n < 0)    return '';   // -1 sentinel from Get-MissingDriverCount on failure
  return n;
}


// ----- GET handler: serves the analytics rows as JSON for the wiki dashboard -----
// The wiki "Driver Installer Analytics" page fetches this endpoint client-side
// and renders it. Returns the most recent MAX_GET_ROWS rows, NEWEST FIRST, as an
// array of objects keyed by stable snake_case names (not the human header text).
//
// After editing this file you MUST redeploy for the change to take effect:
//   Deploy > Manage deployments > (existing deployment) > Edit (pencil) >
//   Version: New version > Deploy.   Do NOT create a new deployment - that
//   changes the URL and breaks both the PowerShell webhook and the wiki page.
//
// CORS: an Apps Script web app deployed "Execute as me / Anyone" serves GET
// responses with Access-Control-Allow-Origin:* (the /exec call 302-redirects to
// googleusercontent.com), so the wiki can fetch it directly from the browser.
const MAX_GET_ROWS = 1000;
const FIELD_KEYS = [
  'timestamp','result','manufacturer','model','serial','os_version','os_build',
  'inf_count','download_mb','missing_before','missing_after','duration_sec',
  'script_version','installed_drivers','driver_urls','missing_after_list'
];

function doGet(e) {
  // ?ping=1 keeps the old plain-text liveness check available.
  if (e && e.parameter && e.parameter.ping) {
    return _textResponse('Driver Installer webhook is alive. POST JSON to record a run.');
  }
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(SHEET_NAME);
    if (!sheet || sheet.getLastRow() < 2) {
      return _jsonResponse({ rows: [], count: 0, generated: new Date().toISOString() });
    }
    var lastRow = sheet.getLastRow();
    var lastCol = Math.min(sheet.getLastColumn(), FIELD_KEYS.length);
    var startRow = Math.max(2, lastRow - MAX_GET_ROWS + 1);   // skip header row 1
    var numRows  = lastRow - startRow + 1;
    var values   = sheet.getRange(startRow, 1, numRows, lastCol).getValues();

    var rows = [];
    for (var i = values.length - 1; i >= 0; i--) {            // newest first
      var src = values[i];
      var obj = {};
      for (var c = 0; c < lastCol; c++) {
        var v = src[c];
        if (c === 0 && v instanceof Date) { v = v.toISOString(); }
        obj[FIELD_KEYS[c]] = v;
      }
      rows.push(obj);
    }
    return _jsonResponse({ rows: rows, count: rows.length, generated: new Date().toISOString() });
  } catch (err) {
    return _jsonResponse({ error: (err && err.message) ? err.message : String(err), rows: [] });
  }
}

function _jsonResponse(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}