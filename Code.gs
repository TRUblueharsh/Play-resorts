/**
 * Meliá Hotels — lead form receiver.
 *
 * Paste this into the Apps Script project behind the /exec URL used by index.html
 * (GOOGLE_SCRIPT_URL), then: Deploy → Manage deployments → ✏️ → Version: New version → Deploy.
 * "New deployment" would mint a different URL and the site would keep hitting the old one.
 *
 * Sheet1 column order this expects (A → H) — matches the live sheet:
 *   A TimeStamp | B Full Name | C Country Code | D phone | E gmail
 *   F wedding date | G guest count | H Budget
 * I (Phone Corrected) and J (1st Follow up) are filled by hand — the script never touches them.
 *
 * The form also sends a formType param. There is no column for it, so it is not stored;
 * add one and extend the row below if that ever needs recording.
 */

// The site submits with navigator.sendBeacon, which POSTs. The lead is in the query string
// either way, so both verbs run the same code.
function doPost(e) {
  return doGet(e);
}

function doGet(e) {
  try {
    // Never getActiveSheet(): it can land on whichever tab was opened last (this file has
    // an LP-2026 tab too, so that is a real hazard here, not a theoretical one).
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Sheet1');
    if (!sheet) return ContentService.createTextOutput('Error: Sheet1 not found');

    const p = (e && e.parameter) || {};

    // Country code goes in as digits only — no "+".
    const countryCode = String(p.countryCode || '').replace(/[^0-9]/g, '');
    // A leading "+" makes Sheets read the cell as a formula, which is where the #ERROR!
    // in D10 came from.
    const phone = String(p.phone || '').trim().replace(/^\+/, '');

    // Write to the FIRST empty row under the header, scanning TimeStamp (column A) top-down.
    // getLastRow()/appendRow() look at every column, so one stray value far down in the
    // follow-up columns pushed leads to row 999+ where nobody could see them.
    const stamps = sheet.getRange(1, 1, sheet.getMaxRows(), 1).getValues();
    let row = 0;
    for (let i = 1; i < stamps.length; i++) {
      if (String(stamps[i][0]).trim() === '') { row = i + 1; break; }
    }
    if (!row) row = stamps.length + 1; // sheet is completely full — fall back to the bottom

    // Format C and D as plain text BEFORE writing, so long numbers don't become 5.5E9
    // and leading zeros survive.
    sheet.getRange(row, 3, 1, 2).setNumberFormat('@');

    sheet.getRange(row, 1, 1, 8).setValues([[
      p.timestamp   || new Date().toISOString(),
      p.name        || '',
      countryCode,
      phone,
      p.email       || '',
      p.weddingDate || '',
      p.guestCount  || '',
      p.budget      || ''
    ]]);

    // Reporting the row back makes it obvious where a lead landed if one ever goes missing.
    return ContentService.createTextOutput(
      'Success: wrote to sheet "' + sheet.getName() + '" row ' + row
    );
  } catch (error) {
    return ContentService.createTextOutput('Error: ' + error.toString());
  }
}
