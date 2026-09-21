/**
 * SMARTWEB 360 — saves workshop leads into the Google Sheet.
 *
 * SETUP (one time):
 * 1. Open the Sheet → Extensions → Apps Script. Delete any code and paste this file.
 * 2. Click Deploy → New deployment → type: Web app.
 *      Execute as: Me      Who has access: Anyone
 * 3. Authorize when asked, then copy the "Web app URL".
 * 4. Paste that URL into LEADS_URL in index.html and push.
 */
const SHEET_ID = "1nKr29MiEer3lIRuHnX4lSj63ZVkcG92wcXrxhY8cJdI";
const TAB_NAME = "Leads";
const HEADERS = ["Timestamp", "Name", "WhatsApp", "City", "Mode", "Amount", "Status", "Source"];

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const d = JSON.parse(e.postData.contents);
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const sh = ss.getSheetByName(TAB_NAME) || ss.insertSheet(TAB_NAME);
    if (sh.getLastRow() === 0) {
      sh.appendRow(HEADERS);
      sh.getRange(1, 1, 1, HEADERS.length).setFontWeight("bold");
      sh.setFrozenRows(1);
    }
    sh.appendRow([
      new Date(), d.name || "", "'" + (d.phone || ""), d.city || "",
      d.mode || "", d.amount || "", "Payment pending", d.source || ""
    ]);
    return ContentService.createTextOutput(JSON.stringify({ ok: true })).setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ ok: false, error: String(err) })).setMimeType(ContentService.MimeType.JSON);
  } finally {
    lock.releaseLock();
  }
}

function doGet() {
  return ContentService.createTextOutput("SMARTWEB 360 leads endpoint is live.");
}
