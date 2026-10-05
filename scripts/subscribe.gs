// Google Apps Script for the "Get Connected" form on the homepage.
// Stores emails in the Google Sheet this script is attached to.
//
// Setup:
// 1. Create a Google Sheet. Extensions > Apps Script. Paste this file in. Save.
// 2. Deploy > New deployment > type "Web app".
//    Execute as: Me. Who has access: Anyone. Deploy and approve access.
// 3. Copy the Web App URL into SUBSCRIBE_URL in index.html.

var SHEET_NAME = 'Subscribers';

function getSheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    sheet.appendRow(['email', 'subscribed_at']);
  }
  return sheet;
}

// Receives the form POST from the website.
function doPost(e) {
  var email = String((e.parameter && e.parameter.email) || '').trim().toLowerCase();
  var trap = (e.parameter && e.parameter.website) || ''; // honeypot: bots fill it, people don't

  if (trap || !/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email) || email.length > 254) {
    return ContentService.createTextOutput('invalid');
  }

  var lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    var sheet = getSheet_();
    var existing = sheet.getRange(1, 1, sheet.getLastRow(), 1).getValues()
      .map(function (r) { return String(r[0]).toLowerCase(); });
    if (existing.indexOf(email) === -1) {
      sheet.appendRow([email, new Date()]);
    }
  } finally {
    lock.releaseLock();
  }
  return ContentService.createTextOutput('ok');
}

// Run by hand from the Apps Script editor after publishing a post.
// Edit the subject and link, then click Run.
// Free Gmail accounts can send to ~100 recipients per day.
function sendUpdate() {
  var subject = 'New post: TITLE HERE';
  var link = 'https://dolomitiinvestor.github.io/Website/articles/SLUG.html';

  var body = 'A new post is up on Dolomiti Investor:\n\n' + link +
    '\n\n--\nReply "unsubscribe" to stop getting these emails.';

  var rows = getSheet_().getDataRange().getValues().slice(1);
  rows.forEach(function (r) {
    if (r[0]) MailApp.sendEmail({ to: r[0], subject: subject, body: body });
  });
}
