/**
 * Playa Resorts — lead form receiver.
 *
 * Sheet1 column order this expects (A → I):
 *   A Timestamp | B Form Type | C Name | D Email | E Country Code | F Mobile
 *   G Guest Count | H Wedding date | I Budget
 * J (Lead Quality) and K (Phone Corrected) are filled by hand — the script never touches them.
 */
// The site submits with navigator.sendBeacon, which POSTs. The lead is in the query string
// either way, so both verbs run the same code.
function doPost(e) {
  return doGet(e);
}

function doGet(e) {
  try {
    // Never getActiveSheet(): it can land on whichever tab was opened last (e.g. LP-2026).
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Sheet1');
    if (!sheet) return ContentService.createTextOutput('Error: Sheet1 not found');

    const p = (e && e.parameter) || {};

    // Country code goes in as digits only — no "+".
    const countryCode = String(p.countryCode || '').replace(/[^0-9]/g, '');
    // A leading "+" makes Sheets read the cell as a formula, which is where #ERROR! came from.
    const phone = String(p.phone || '').trim().replace(/^\+/, '');

    // Write to the FIRST empty row under the header, scanning Timestamp (column A) top-down.
    // getLastRow()/appendRow() look at every column, so one stray value far down in the
    // follow-up columns pushed leads to row 999+ where nobody could see them.
    const stamps = sheet.getRange(1, 1, sheet.getMaxRows(), 1).getValues();
    let row = 0;
    for (let i = 1; i < stamps.length; i++) {
      if (String(stamps[i][0]).trim() === '') { row = i + 1; break; }
    }
    if (!row) row = stamps.length + 1; // sheet is completely full — fall back to the bottom

    // Format E and F as plain text BEFORE writing, so long numbers don't become 5.5E9
    // and leading zeros survive.
    sheet.getRange(row, 5, 1, 2).setNumberFormat('@');

    sheet.getRange(row, 1, 1, 9).setValues([[
      p.timestamp   || new Date().toISOString(),
      p.formType    || 'Unknown',
      p.name        || '',
      p.email       || '',
      countryCode,
      phone,
      p.guestCount  || '',
      p.weddingDate || '',
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
