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
 * If you edit this script later: Deploy → Manage deployments → Edit (pencil) → Version: New version → Deploy.
 * (Editing the existing deployment keeps the same /exec URL, so the website needs no change.)
 *
 * EMAIL ALERTS: run the function "testEmailAlert" once from the editor (select it in the
 * toolbar → Run) to grant the "send email" permission and receive a sample alert.
 */

const SHEET_NAME = 'Registrations';
const ALERT_EMAIL = 'baitulizzah2024@gmail.com';   // gets an email for every new submitted registration

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
    const isNewSubmission = data.status === 'Submitted' && !row[colIdx_('submittedAt')];
    if (isNewSubmission) row[colIdx_('submittedAt')] = now;

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

    if (isNewSubmission) {
      try { sendAlert_(row, rowIndex); } catch (mailErr) { console.error('Email alert failed: ' + mailErr); }
    }

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

// ===== Email alert for each new registration =====
function sendAlert_(row, rowIndex) {
  const get = k => {
    const v = row[colIdx_(k)];
    return (v === '' || v === null || v === undefined) ? '-' : String(v).replace(/^'/, '');
  };
  const relation = get('relation') === '-' ? 'Guardian' : get('relation');
  const fields = [
    ['Full Name', get('fullName')],
    ['Mobile Number', get('mobile')],
    ['WhatsApp Number', get('whatsapp')],
    ['Age', get('age') === '-' ? '-' : get('age') + ' Years'],
    ['Marital Status', get('marital')],
    ['Family Members', get('family')],
    ['Education', get('education')],
    [relation + "'s Name", get('guardianName')],
    [relation + "'s Occupation", get('guardianOccupation')],
    ['Country', get('country')],
    ['City', get('city')],
    ['Languages', get('languages')],
    ['Previous Islamic Education', get('prevEducation')],
    ['Institute', get('institute')],
    ['Email Address', get('email')]
  ];
  const esc = s => String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
  const waDigits = (get('whatsapp') !== '-' ? get('whatsapp') : get('mobile')).replace(/\D/g, '').replace(/^0/, '92');
  const sheetUrl = SpreadsheetApp.getActiveSpreadsheet().getUrl() + '#gid=' + getSheet_().getSheetId() + '&range=A' + rowIndex;

  const html =
    '<div style="font-family:Arial,sans-serif;max-width:560px;color:#1a2e1a">' +
    '<h2 style="color:#2d6a4f;margin:0 0 4px">New Registration</h2>' +
    '<p style="margin:0 0 16px;color:#6b6b5a">Bait ul Izzah Quran Academy website</p>' +
    '<table style="border-collapse:collapse;width:100%">' +
    fields.map(([k, v], i) =>
      '<tr style="background:' + (i % 2 ? '#ffffff' : '#f6faf6') + '">' +
      '<td style="padding:8px 10px;font-weight:bold;width:45%;border-bottom:1px solid #e5eee5">' + esc(k) + '</td>' +
      '<td style="padding:8px 10px;border-bottom:1px solid #e5eee5">' + esc(v) + '</td></tr>').join('') +
    '</table>' +
    '<p style="margin:18px 0 0">' +
    (waDigits ? '<a href="https://wa.me/' + waDigits + '" style="background:#25D366;color:#fff;padding:10px 18px;border-radius:20px;text-decoration:none;font-weight:bold;margin-right:8px">Reply on WhatsApp</a>' : '') +
    '<a href="' + sheetUrl + '" style="background:#2d6a4f;color:#fff;padding:10px 18px;border-radius:20px;text-decoration:none;font-weight:bold">Open in Sheet</a>' +
    '</p></div>';

  const plain = 'New Registration – Bait ul Izzah Quran Academy\n\n' +
    fields.map(([k, v]) => k + ': ' + v).join('\n') + '\n\nSheet: ' + sheetUrl;

  MailApp.sendEmail({
    to: ALERT_EMAIL,
    subject: 'New Registration: ' + get('fullName') + ' (' + get('mobile') + ')',
    body: plain,
    htmlBody: html,
    name: 'Bait ul Izzah Website'
  });
}

// Run this once from the Apps Script editor to authorize email sending and get a sample alert.
function testEmailAlert() {
  const sample = COLUMNS.map(() => '');
  const set = (k, v) => { sample[colIdx_(k)] = v; };
  set('fullName', 'Sample Student'); set('mobile', '03001234567'); set('whatsapp', '03001234567');
  set('age', 25); set('relation', 'Father'); set('guardianName', 'Sample Father');
  set('country', 'Pakistan'); set('city', 'Okara'); set('languages', 'Urdu, English');
  sendAlert_(sample, 2);
}
