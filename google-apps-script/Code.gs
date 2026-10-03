/**
 * Bait ul Izzah Quran Academy – Registration form → Google Sheet
 *
 * Saves every registration to a Google Sheet, including forms that were
 * started but never submitted (status "Incomplete"). Each visitor gets one
 * row that is updated as they type, so there are no duplicate rows.
 *
 * SETUP (one time, ~3 minutes)
 * 1. Create a new Google Sheet (e.g. "Bait ul Izzah Registrations").
 * 2. In the sheet: Extensions → Apps Script. Delete any code there and paste this file.
 * 3. Click Deploy → New deployment → gear icon → "Web app".
 *      Execute as:      Me
 *      Who has access:  Anyone
 *    Click Deploy and approve the permissions.
 * 4. Copy the "Web app URL" (ends in /exec) and paste it into index.html:
 *      const SHEET_URL = 'https://script.google.com/macros/s/XXXX/exec';
 * 5. Commit index.html. Done – rows appear in the "Registrations" tab.
 *
 * If you edit this script later: Deploy → Manage deployments → Edit → Version: New version.
 */

const SHEET_NAME = 'Registrations';

const COLUMNS = [
  ['id', 'Entry ID'],
  ['status', 'Status'],
  ['firstSeen', 'Started At'],
  ['lastUpdated', 'Last Updated'],
  ['submittedAt', 'Submitted At'],
  ['fullName', 'Full Name'],
  ['age', 'Age (Years)'],
  ['marital', 'Marital Status'],
  ['family', 'Family Members'],
  ['relation', 'Father/Husband/Guardian'],
  ['guardianName', 'Guardian Name'],
  ['guardianOccupation', 'Guardian Occupation'],
  ['country', 'Country'],
  ['city', 'City'],
  ['languages', 'Languages'],
  ['education', 'Education'],
  ['prevEducation', 'Previous Islamic Education'],
  ['institute', 'Institute'],
  ['mobile', 'Mobile Number'],
  ['whatsapp', 'WhatsApp Number'],
  ['email', 'Email Address'],
  ['page', 'Page'],
  ['device', 'Device']
];

function getSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) sh = ss.insertSheet(SHEET_NAME);
  if (sh.getLastRow() === 0) {
    sh.appendRow(COLUMNS.map(c => c[1]));
    sh.getRange(1, 1, 1, COLUMNS.length).setFontWeight('bold').setBackground('#d8f3dc');
    sh.setFrozenRows(1);
  }
  return sh;
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const data = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    if (!data.id) return json_({ ok: false, error: 'missing id' });

    const sh = getSheet_();
    const now = new Date();
    const last = sh.getLastRow();

    // Find existing row for this entry
    let rowIndex = -1;
    if (last > 1) {
      const ids = sh.getRange(2, 1, last - 1, 1).getValues();
      for (let i = ids.length - 1; i >= 0; i--) {
        if (String(ids[i][0]) === String(data.id)) { rowIndex = i + 2; break; }
      }
    }

    let row;
    if (rowIndex > 0) {
      row = sh.getRange(rowIndex, 1, 1, COLUMNS.length).getValues()[0];
    } else {
      row = COLUMNS.map(() => '');
      row[colIdx_('firstSeen')] = now;
    }

    // Never downgrade a submitted entry back to Incomplete
    const prevStatus = row[colIdx_('status')];
    const status = (prevStatus === 'Submitted' && data.status !== 'Submitted') ? 'Submitted' : (data.status || 'Incomplete');

    COLUMNS.forEach(([key], i) => {
      if (['id', 'status', 'firstSeen', 'lastUpdated', 'submittedAt'].includes(key)) return;
      if (data[key] !== undefined) row[i] = data[key];
    });
    row[colIdx_('id')] = data.id;
    row[colIdx_('status')] = status;
    row[colIdx_('lastUpdated')] = now;
    if (data.status === 'Submitted' && !row[colIdx_('submittedAt')]) row[colIdx_('submittedAt')] = now;

    // Keep phone numbers as text (stops Sheets from dropping the leading 0 / +)
    ['mobile', 'whatsapp'].forEach(k => {
      const v = row[colIdx_(k)];
      if (v && String(v).charAt(0) !== "'") row[colIdx_(k)] = "'" + v;
    });

    if (rowIndex > 0) {
      sh.getRange(rowIndex, 1, 1, COLUMNS.length).setValues([row]);
    } else {
      sh.appendRow(row);
      rowIndex = sh.getLastRow();
    }
    sh.getRange(rowIndex, colIdx_('status') + 1)
      .setBackground(status === 'Submitted' ? '#b7e4c7' : '#fdf3dc');

    return json_({ ok: true, row: rowIndex, status: status });
  } catch (err) {
    return json_({ ok: false, error: String(err) });
  } finally {
    lock.releaseLock();
  }
}

function doGet() {
  return json_({ ok: true, message: 'Bait ul Izzah registration endpoint is running.' });
}

function colIdx_(key) {
  return COLUMNS.findIndex(c => c[0] === key);
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
