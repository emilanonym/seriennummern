/*** Konfiguration ***/
const SHEET_ID   = '1RwZyvB43_xdKkP-xP0EHsI75h8XLP9zzDYRBtx6XbRY';
const SHEET_NAME = 'Seriennummern';
const SECRET     = 'sn7Qe2vKptaX8'; // selbe Zeichenfolge wie in index.html

function doGet(e) {
  const p = (e && e.parameter) || {};

  if (p.action === 'add') {
    let out;
    if (p.token !== SECRET) {
      out = { status: 'error', message: 'unauthorized' };
    } else {
      out = addScan(p.code);
    }
    const json = JSON.stringify(out);
    if (p.callback) {
      return ContentService
        .createTextOutput(p.callback + '(' + json + ')')
        .setMimeType(ContentService.MimeType.JAVASCRIPT);
    }
    return ContentService.createTextOutput(json)
      .setMimeType(ContentService.MimeType.JSON);
  }

  return HtmlService.createHtmlOutput(
    '<p style="font-family:sans-serif">Scanner läuft über GitHub Pages.</p>'
  );
}

function getSheet_() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    sheet.appendRow(['Seriennummer', 'Zeitstempel']);
  }
  sheet.getRange('A:A').setNumberFormat('@');
  return sheet;
}

function addScan(code) {
  code = String(code || '').trim();
  if (!code) return { status: 'empty' };

  // Verhindert doppelte Einträge bei gleichzeitigen Scans:
  // Prüfen und Schreiben passieren nur unter exklusivem Lock.
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
  } catch (err) {
    return { status: 'error', message: 'Sheet gerade belegt, bitte nochmal scannen' };
  }

  try {
    const sheet = getSheet_();
    const lastRow = sheet.getLastRow();

    if (lastRow >= 2) {
      const vorhandene = sheet.getRange(2, 1, lastRow - 1, 1).getValues().flat().map(String);
      if (vorhandene.indexOf(code) !== -1) {
        return { status: 'duplicate', code: code, count: lastRow - 1 };
      }
    }

    const ts = Utilities.formatDate(new Date(), 'Europe/Berlin', 'yyyy-MM-dd HH:mm:ss');
    sheet.appendRow([code, ts]);
    return { status: 'ok', code: code, count: sheet.getLastRow() - 1 };
  } finally {
    lock.releaseLock();
  }
}
