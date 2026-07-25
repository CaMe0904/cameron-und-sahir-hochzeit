/**
 * Cameron & Sahir — RSVP → Google Sheets
 * -------------------------------------------------------------
 * Paste this whole file into the Apps Script editor attached to your
 * Google Sheet (Extensions → Apps Script), replacing any starter code.
 * Full setup steps are in README.md.
 */

var SHEET_NAME = 'RSVPs'; // change if you want a different tab name

var HEADERS = [
  'timestamp',        // when Apps Script received the submission (server time)
  'submittedAt',       // ISO timestamp sent by the browser
  'name',              // guest / family name
  'email',
  'attendance',        // 'yes' | 'maybe' | 'no'
  'adults',
  'children',
  'guestNames',        // names of additional guests
  'dietary',           // dietary restrictions / allergies
  'song',              // song request
  'message'            // personal message
];

function doPost(e) {
  try {
    var sheet = getOrCreateSheet_();
    var p = (e && e.parameter) ? e.parameter : {};

    // simple spam guard: if the honeypot "website" field is filled, a bot filled it — ignore silently.
    if (p.website) {
      return jsonOut_({ result: 'success' });
    }

    var row = HEADERS.map(function (key) {
      if (key === 'timestamp') return new Date();
      return p[key] !== undefined ? p[key] : '';
    });

    sheet.appendRow(row);
    return jsonOut_({ result: 'success' });
  } catch (err) {
    return jsonOut_({ result: 'error', message: String(err) });
  }
}

// Lets you open the deployed /exec URL directly in a browser to sanity-check it's alive.
function doGet(e) {
  return jsonOut_({ result: 'success', message: 'RSVP endpoint is live. Submit the website RSVP form to add a row.' });
}

function getOrCreateSheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
  }
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(HEADERS);
    sheet.getRange(1, 1, 1, HEADERS.length).setFontWeight('bold');
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function jsonOut_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
