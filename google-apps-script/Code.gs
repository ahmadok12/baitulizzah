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
 *
 * DAILY INCOMPLETE-FORMS EMAIL: run "setupDailyDigest" once from the editor. Every morning at
 * 9 AM (Pakistan time) you'll get a list of people who started the form in the last 24 hours
 * but didn't submit, with WhatsApp buttons to follow up. Run "sendIncompleteDigest" to test now.
 *
 * SUBMISSIONS TAB: every submitted registration is also copied, permanently, to the
 * "Submissions" tab (one row per registration, never overwritten). Run "rebuildSubmissions"
 * once to copy in registrations that were submitted before this tab existed.
 *
 * VISITOR TRACKING: the website logs each visit (country, city, device, source – no IP address
 * is stored) to the "Visits" tab. Run "setupVisitDashboard" once to create the
 * "Visitor Dashboard" tab with live totals. The 9 AM daily email also includes a visitor summary.
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
    if (data.type === 'visit') return logVisit_(data);
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
      try { recordSubmission_(row); } catch (subErr) { console.error('Submissions copy failed: ' + subErr); }
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
    ['WhatsApp Number', get('whatsapp')],
    ['Mobile Number', get('mobile')],
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
    subject: 'New Registration: ' + get('fullName') + ' (' + (get('whatsapp') !== '-' ? get('whatsapp') : get('mobile')) + ')',
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


// ===== Submissions tab: permanent record of every submitted registration =====
const SUBMISSIONS_SHEET = 'Submissions';
const SUBMISSION_FIELDS = ['submittedAt', 'fullName', 'whatsapp', 'mobile', 'email', 'age', 'marital', 'family', 'relation',
  'guardianName', 'guardianOccupation', 'country', 'city', 'languages', 'education', 'prevEducation', 'institute', 'id'];

function getSubmissionsSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(SUBMISSIONS_SHEET);
  if (!sh) sh = ss.insertSheet(SUBMISSIONS_SHEET);
  if (sh.getLastRow() === 0) {
    const headers = ['#'].concat(SUBMISSION_FIELDS.map(k => k === 'submittedAt' ? 'Submitted At' : COLUMNS[colIdx_(k)][1]));
    sh.appendRow(headers);
    sh.getRange(1, 1, 1, headers.length).setFontWeight('bold').setBackground('#b7e4c7');
    sh.setFrozenRows(1);
  }
  return sh;
}

function recordSubmission_(row) {
  const sh = getSubmissionsSheet_();
  const n = sh.getLastRow();   // header row counts as 1, so this is the next serial number
  const values = SUBMISSION_FIELDS.map(k => {
    const v = row[colIdx_(k)];
    return (k === 'whatsapp' || k === 'mobile') && v && String(v).charAt(0) !== "'" ? "'" + v : v;
  });
  sh.appendRow([n].concat(values));
}

// Run once to copy registrations submitted before the Submissions tab existed (skips ones already copied).
function rebuildSubmissions() {
  const reg = getSheet_();
  const sub = getSubmissionsSheet_();
  const idCol = SUBMISSION_FIELDS.length + 1;   // Entry ID is the last column
  const have = sub.getLastRow() > 1 ? sub.getRange(2, idCol, sub.getLastRow() - 1, 1).getValues().map(r => String(r[0])) : [];
  if (reg.getLastRow() < 2) return;
  const rows = reg.getRange(2, 1, reg.getLastRow() - 1, COLUMNS.length).getValues()
    .filter(r => r[colIdx_('status')] === 'Submitted' && have.indexOf(String(r[colIdx_('id')])) === -1)
    .sort((a, b) => new Date(a[colIdx_('submittedAt')]) - new Date(b[colIdx_('submittedAt')]));
  rows.forEach(recordSubmission_);
}

// ===== Visitor tracking =====
const VISITS_SHEET = 'Visits';
const VISIT_COLUMNS = ['Time', 'Date', 'Country', 'Region', 'City', 'Device', 'Source', 'Referrer', 'Browser Language', 'Time Zone', 'Visitor ID', 'New / Returning'];

function getVisitsSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(VISITS_SHEET);
  if (!sh) sh = ss.insertSheet(VISITS_SHEET);
  if (sh.getLastRow() === 0) {
    sh.appendRow(VISIT_COLUMNS);
    sh.getRange(1, 1, 1, VISIT_COLUMNS.length).setFontWeight('bold').setBackground('#d8f3dc');
    sh.setFrozenRows(1);
  }
  return sh;
}

function logVisit_(d) {
  const clip = (v, n) => String(v || '').slice(0, n || 120);
  const now = new Date();
  getVisitsSheet_().appendRow([
    now,
    Utilities.formatDate(now, 'Asia/Karachi', 'yyyy-MM-dd'),
    clip(d.country, 60) || 'Unknown', clip(d.region, 60), clip(d.city, 60) || 'Unknown',
    clip(d.device, 20), clip(d.source, 40), clip(d.referrer, 200),
    clip(d.lang, 20), clip(d.tz, 60), clip(d.visitor, 40), d.returning ? 'Returning' : 'New'
  ]);
  return json_({ ok: true });
}

// Run once: builds a "Visitor Dashboard" tab with live formulas (safe to run again).
function setupVisitDashboard() {
  getVisitsSheet_();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName('Visitor Dashboard');
  if (sh) sh.clear(); else sh = ss.insertSheet('Visitor Dashboard', 0);
  const V = "Visits!";
  sh.getRange('A1').setValue('Bait ul Izzah – Website Visitors').setFontSize(16).setFontWeight('bold').setFontColor('#2d6a4f');
  sh.getRange('A2').setValue('Updates automatically. Dates in Pakistan time.').setFontColor('#6b6b5a');

  const kpis = [
    ['Visits today', '=COUNTIFS(' + V + 'A2:A, ">="&TODAY())'],
    ['Last 7 days', '=COUNTIFS(' + V + 'A2:A, ">="&(NOW()-7))'],
    ['Last 30 days', '=COUNTIFS(' + V + 'A2:A, ">="&(NOW()-30))'],
    ['All time', '=COUNTA(' + V + 'A2:A)'],
    ['Countries (all time)', '=IFERROR(COUNTUNIQUE(FILTER(' + V + 'C2:C, ' + V + 'C2:C<>"")),0)']
  ];
  kpis.forEach(([label, f], i) => {
    sh.getRange(4, 1 + i).setValue(label).setFontWeight('bold').setFontColor('#6b6b5a').setFontSize(9);
    sh.getRange(5, 1 + i).setFormula(f).setFontSize(20).setFontWeight('bold').setFontColor('#1b4d36');
  });
  sh.getRange(4, 1, 2, 5).setBackground('#f6faf6');

  const block = (row, col, title, formula) => {
    sh.getRange(row, col).setValue(title).setFontWeight('bold').setFontColor('#ffffff').setBackground('#2d6a4f');
    sh.getRange(row, col + 1).setBackground('#2d6a4f');
    sh.getRange(row + 1, col).setFormula(formula);
  };
  // Last 30 days breakdowns
  const last30 = V + 'A2:A>=NOW()-30';
  block(8, 1, 'Top countries (30 days)',
    '=IFERROR(QUERY(FILTER({' + V + 'C2:C}, ' + last30 + '), "select Col1, count(Col1) group by Col1 order by count(Col1) desc limit 20 label Col1 \'Country\', count(Col1) \'Visits\'", 0), "No visits yet")');
  block(8, 4, 'Top cities (30 days)',
    '=IFERROR(QUERY(FILTER({' + V + 'E2:E&", "&' + V + 'C2:C}, ' + last30 + '), "select Col1, count(Col1) group by Col1 order by count(Col1) desc limit 25 label Col1 \'City\', count(Col1) \'Visits\'", 0), "No visits yet")');
  block(8, 7, 'Traffic source (30 days)',
    '=IFERROR(QUERY(FILTER({' + V + 'G2:G}, ' + last30 + '), "select Col1, count(Col1) group by Col1 order by count(Col1) desc label Col1 \'Source\', count(Col1) \'Visits\'", 0), "No visits yet")');
  block(8, 10, 'Device (30 days)',
    '=IFERROR(QUERY(FILTER({' + V + 'F2:F}, ' + last30 + '), "select Col1, count(Col1) group by Col1 order by count(Col1) desc label Col1 \'Device\', count(Col1) \'Visits\'", 0), "No visits yet")');
  block(36, 1, 'Visits per day (30 days)',
    '=IFERROR(QUERY(FILTER({' + V + 'B2:B}, ' + last30 + '), "select Col1, count(Col1) group by Col1 order by Col1 desc label Col1 \'Date\', count(Col1) \'Visits\'", 0), "No visits yet")');

  [1, 4, 7, 10].forEach(c => sh.setColumnWidth(c, 200));
  [2, 5, 8, 11].forEach(c => sh.setColumnWidth(c, 70));
  [3, 6, 9].forEach(c => sh.setColumnWidth(c, 24));
  sh.setFrozenRows(5);
  ss.setActiveSheet(sh);
}

function visitSummary_(since) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(VISITS_SHEET);
  if (!sh || sh.getLastRow() < 2) return { total: 0 };
  const rows = sh.getRange(2, 1, sh.getLastRow() - 1, VISIT_COLUMNS.length).getValues()
    .filter(r => r[0] instanceof Date && r[0] >= since);
  const tally = idxFn => {
    const m = {};
    rows.forEach(r => { const k = idxFn(r); if (k) m[k] = (m[k] || 0) + 1; });
    return Object.entries(m).sort((a, b) => b[1] - a[1]);
  };
  return {
    total: rows.length,
    countries: tally(r => r[2]),
    cities: tally(r => r[4] && r[4] !== 'Unknown' ? r[4] + ', ' + r[2] : ''),
    sources: tally(r => r[6]),
    devices: tally(r => r[5])
  };
}

// ===== Daily email: visitors + forms started but not submitted =====
function sendIncompleteDigest() {
  const sh = getSheet_();
  const last = sh.getLastRow();
  const rows = last >= 2 ? sh.getRange(2, 1, last - 1, COLUMNS.length).getValues() : [];
  const since = new Date(Date.now() - 24 * 60 * 60 * 1000);
  const visits = visitSummary_(since);
  const clean = v => (v === '' || v === null || v === undefined) ? '' : String(v).replace(/^'/, '');
  const pending = rows
    .map((r, i) => ({ r, row: i + 2 }))
    .filter(({ r }) => r[colIdx_('status')] !== 'Submitted' && r[colIdx_('lastUpdated')] instanceof Date && r[colIdx_('lastUpdated')] >= since);
  if (!pending.length && !visits.total) return;   // nothing to report – no email

  const esc = s => String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
  const ssUrl = SpreadsheetApp.getActiveSpreadsheet().getUrl() + '#gid=' + sh.getSheetId();
  const tz = 'Asia/Karachi';
  const rowsHtml = pending.map(({ r, row }, i) => {
    const phone = clean(r[colIdx_('whatsapp')]) || clean(r[colIdx_('mobile')]);
    const digits = phone.replace(/\D/g, '').replace(/^0/, '92');
    const name = clean(r[colIdx_('fullName')]) || '(no name)';
    const place = [clean(r[colIdx_('city')]), clean(r[colIdx_('country')])].filter(Boolean).join(', ');
    const when = Utilities.formatDate(r[colIdx_('lastUpdated')], tz, 'd MMM, h:mm a');
    return '<tr style="background:' + (i % 2 ? '#fff' : '#f6faf6') + '">' +
      '<td style="padding:8px 10px;border-bottom:1px solid #e5eee5"><b>' + esc(name) + '</b>' + (place ? '<br><span style="color:#6b6b5a;font-size:12px">' + esc(place) + '</span>' : '') + '</td>' +
      '<td style="padding:8px 10px;border-bottom:1px solid #e5eee5">' + (esc(phone) || '-') + '</td>' +
      '<td style="padding:8px 10px;border-bottom:1px solid #e5eee5;color:#6b6b5a;font-size:12px">' + esc(when) + '</td>' +
      '<td style="padding:8px 10px;border-bottom:1px solid #e5eee5">' +
        (digits.length >= 10 ? '<a href="https://wa.me/' + digits + '" style="background:#25D366;color:#fff;padding:6px 12px;border-radius:14px;text-decoration:none;font-weight:bold;font-size:12px">WhatsApp</a>' : '<span style="color:#999;font-size:12px">no number</span>') +
        ' <a href="' + ssUrl + '&range=A' + row + '" style="color:#2d6a4f;font-size:12px">row ' + row + '</a></td></tr>';
  }).join('');

  const list = (arr, n) => (arr || []).slice(0, n).map(([k, v]) =>
    '<tr><td style="padding:5px 10px;border-bottom:1px solid #eef3ee">' + esc(k) + '</td><td align="right" style="padding:5px 10px;border-bottom:1px solid #eef3ee;font-weight:bold">' + v + '</td></tr>').join('');
  const mini = (title, arr, n) => '<td valign="top" style="padding:0 8px 0 0;width:50%"><table style="border-collapse:collapse;width:100%;font-size:13px">' +
    '<tr><th align="left" colspan="2" style="padding:6px 10px;background:#d8f3dc;color:#1b4d36">' + title + '</th></tr>' + (list(arr, n) || '<tr><td style="padding:5px 10px;color:#999">–</td></tr>') + '</table></td>';
  const dashUrl = SpreadsheetApp.getActiveSpreadsheet().getUrl();
  const visitsHtml = visits.total ?
    '<h2 style="color:#2d6a4f;margin:0 0 4px">' + visits.total + ' website visit' + (visits.total > 1 ? 's' : '') + ' in the last 24 hours</h2>' +
    '<p style="margin:0 0 12px;color:#6b6b5a">From ' + visits.countries.length + ' countr' + (visits.countries.length === 1 ? 'y' : 'ies') + '.</p>' +
    '<table style="border-collapse:collapse;width:100%"><tr>' + mini('Countries', visits.countries, 10) + mini('Cities', visits.cities, 10) + '</tr>' +
    '<tr><td colspan="2" style="height:10px"></td></tr><tr>' + mini('Source', visits.sources, 6) + mini('Device', visits.devices, 4) + '</tr></table>' +
    '<p style="margin:12px 0 26px"><a href="' + dashUrl + '" style="color:#2d6a4f;font-weight:bold">Open the Visitor Dashboard →</a></p>' : '';

  const formsHtml = !pending.length ? '' :
    '<h2 style="color:#2d6a4f;margin:0 0 4px">' + pending.length + ' incomplete registration' + (pending.length > 1 ? 's' : '') + '</h2>' +
    '<p style="margin:0 0 14px;color:#6b6b5a">Started on the website in the last 24 hours but not submitted. A quick WhatsApp message can help them finish.</p>' +
    '<table style="border-collapse:collapse;width:100%;font-size:14px"><tr style="background:#2d6a4f;color:#fff">' +
    '<th align="left" style="padding:8px 10px">Name</th><th align="left" style="padding:8px 10px">Number</th><th align="left" style="padding:8px 10px">Last activity</th><th align="left" style="padding:8px 10px">Follow up</th></tr>' +
    rowsHtml + '</table>' +
    '<p style="margin:16px 0 0"><a href="' + ssUrl + '" style="color:#2d6a4f;font-weight:bold">Open the Registrations sheet →</a></p>';
  const html = '<div style="font-family:Arial,sans-serif;max-width:640px;color:#1a2e1a">' + visitsHtml + formsHtml + '</div>';

  const plain = pending.map(({ r }) => (clean(r[colIdx_('fullName')]) || '(no name)') + ' – ' +
    (clean(r[colIdx_('whatsapp')]) || clean(r[colIdx_('mobile')]) || 'no number')).join('\n');

  MailApp.sendEmail({
    to: ALERT_EMAIL,
    subject: 'Daily report: ' + visits.total + ' visit' + (visits.total === 1 ? '' : 's') +
      (pending.length ? ', ' + pending.length + ' incomplete registration' + (pending.length > 1 ? 's' : '') : ''),
    body: 'Website visits (last 24 hours): ' + visits.total + '\n' +
      (visits.countries || []).slice(0, 10).map(([k, v]) => '  ' + k + ': ' + v).join('\n') +
      (pending.length ? '\n\nIncomplete registrations:\n' + plain : '') + '\n\nSheet: ' + ssUrl,
    htmlBody: html,
    name: 'Bait ul Izzah Website'
  });
}

// Run once: schedules sendIncompleteDigest every day at 9 AM Pakistan time (safe to run again).
function setupDailyDigest() {
  ScriptApp.getProjectTriggers()
    .filter(t => t.getHandlerFunction() === 'sendIncompleteDigest')
    .forEach(t => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger('sendIncompleteDigest').timeBased().everyDays(1).atHour(9).inTimezone('Asia/Karachi').create();
}
