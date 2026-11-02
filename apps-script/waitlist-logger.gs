// Deployment (one-time):
// 1. Create a new Google Sheet — this becomes your waitlist/contact record.
// 2. In the Sheet: Extensions > Apps Script.
// 3. Delete any starter code and paste this whole file in.
// 4. Run `setup` once from the Apps Script editor toolbar — creates the
//    Waitlist / Updates / Messages tabs with header rows. The first run
//    will prompt you to authorize the script.
// 5. Deploy > New deployment > type "Web app".
//      - Execute as: Me
//      - Who has access: Anyone
// 6. Copy the resulting Web App URL and send it back so it can be wired
//    into the site's contact form.
// Redeploy (Deploy > Manage deployments > edit > new version) any time you
// change this file, or the live URL won't reflect your edits.

const SHEETS_BY_REASON = {
  waitlist: { name: "Waitlist", headers: ["Timestamp", "Email"] },
  updates: { name: "Updates", headers: ["Timestamp", "Email"] },
  message: { name: "Messages", headers: ["Timestamp", "Email", "Message"] },
};

function setup() {
  Object.keys(SHEETS_BY_REASON).forEach((reason) => getOrCreateSheet(reason));
}

function getOrCreateSheet(reason) {
  const config = SHEETS_BY_REASON[reason] || SHEETS_BY_REASON.message;
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = spreadsheet.getSheetByName(config.name);
  if (!sheet) {
    sheet = spreadsheet.insertSheet(config.name);
    sheet.appendRow(config.headers);
  }
  return sheet;
}

// Neutralizes leading =, +, -, @ so Sheets can't interpret a submitted
// value as a formula when the sheet is opened (spreadsheet/CSV injection).
function sanitizeCell(value) {
  const s = String(value || "");
  return /^[=+\-@]/.test(s) ? "'" + s : s;
}

function doPost(e) {
  let data;
  try {
    data = JSON.parse(e.postData.contents);
  } catch (err) {
    return reject();
  }
  // Drop junk from bots posting straight to this URL: unknown reasons,
  // malformed emails, oversized fields.
  const email = String(data.email || "").trim();
  const message = String(data.message || "");
  if (!SHEETS_BY_REASON[data.reason]) return reject();
  if (!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email) || email.length > 254) return reject();
  if (message.length > 5000) return reject();
  if (data.reason === "message" && !message.trim()) return reject();

  const sheet = getOrCreateSheet(data.reason);

  const row = [new Date(), sanitizeCell(email)];
  if (data.reason === "message") row.push(sanitizeCell(message));
  sheet.appendRow(row);

  return ContentService
    .createTextOutput(JSON.stringify({ result: "success" }))
    .setMimeType(ContentService.MimeType.JSON);
}

function reject() {
  return ContentService
    .createTextOutput(JSON.stringify({ result: "rejected" }))
    .setMimeType(ContentService.MimeType.JSON);
}
