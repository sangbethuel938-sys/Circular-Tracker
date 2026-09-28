/**
 * EATTA Circular Register - Google Apps Script API
 *
 * SETUP:
 * 1. Create a Google Sheet named "EATTA Circular Register".
 * 2. Extensions > Apps Script, paste this file.
 * 3. In Project Settings > Script properties add:
 *      APP_SHARED_SECRET = the same secret used by Streamlit
 * 4. Deploy > New deployment > Web app
 *      Execute as: Me
 *      Who has access: Anyone (the shared secret still protects writes)
 * 5. Copy the Web app URL into .streamlit/secrets.toml.
 */

const REGISTER_SHEET = 'Circular Register';
const HEADERS = [
  'Serial No.',
  'Circular Number',
  'Date',
  'Subject',
  'To',
  'From',
  'Status',
  'File Name',
  'Local File Path',
  'Date Recorded',
  'Remarks'
];

function doGet() {
  return jsonResponse({
    ok: true,
    service: 'EATTA Circular Register API',
    version: '1.0',
    status: 'running'
  });
}

function doPost(e) {
  try {
    const payload = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    verifySecret(payload.secret);

    const action = String(payload.action || '').toLowerCase();
    if (action === 'reserve') return reserveSerial(payload);
    if (action === 'complete') return completeRecord(payload);
    if (action === 'cancel') return cancelRecord(payload);
    if (action === 'list') return listRecords(payload);
    if (action === 'health') return jsonResponse({ ok: true, service: 'EATTA Circular Register API', version: '1.0', status: 'running' });

    return jsonResponse({ ok: false, error: 'Unknown action.' });
  } catch (err) {
    return jsonResponse({ ok: false, error: String(err && err.message ? err.message : err) });
  }
}

function verifySecret(received) {
  const expected = PropertiesService.getScriptProperties().getProperty('APP_SHARED_SECRET');
  if (!expected) throw new Error('APP_SHARED_SECRET is not configured in Apps Script properties.');
  if (!received || received !== expected) throw new Error('Unauthorized request.');
}

function getRegisterSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(REGISTER_SHEET);
  if (!sheet) sheet = ss.insertSheet(REGISTER_SHEET);

  if (sheet.getLastRow() === 0) {
    sheet.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]);
    sheet.getRange(1, 1, 1, HEADERS.length)
      .setFontWeight('bold')
      .setBackground('#0B6B3A')
      .setFontColor('#FFFFFF');
    sheet.setFrozenRows(1);
    sheet.autoResizeColumns(1, HEADERS.length);
  }
  return sheet;
}

function reserveSerial(payload) {
  const year = Number(payload.year || new Date().getFullYear());
  if (!Number.isInteger(year) || year < 2000 || year > 2200) throw new Error('Invalid year.');

  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const sheet = getRegisterSheet();
    const lastRow = sheet.getLastRow();
    const prefix = `EATTA/CIR/${year}/`;
    let maxSeq = 0;

    if (lastRow >= 2) {
      const values = sheet.getRange(2, 2, lastRow - 1, 1).getDisplayValues().flat();
      values.forEach(v => {
        const s = String(v || '').trim();
        if (s.startsWith(prefix)) {
          const n = Number(s.slice(prefix.length));
          if (Number.isInteger(n) && n > maxSeq) maxSeq = n;
        }
      });
    }

    const seq = maxSeq + 1;
    const circularNumber = `${prefix}${String(seq).padStart(3, '0')}`;
    const now = new Date();
    const row = [seq, circularNumber, '', '', '', '', 'RESERVED', '', '', now, 'Serial reserved by app'];
    sheet.appendRow(row);

    return jsonResponse({
      ok: true,
      serial_no: seq,
      circular_number: circularNumber,
      row_number: sheet.getLastRow()
    });
  } finally {
    lock.releaseLock();
  }
}

function completeRecord(payload) {
  const number = String(payload.circular_number || '').trim();
  if (!number) throw new Error('circular_number is required.');

  const sheet = getRegisterSheet();
  const rowNumber = findCircularRow(sheet, number);
  if (!rowNumber) throw new Error(`Circular ${number} was not reserved.`);

  const existing = sheet.getRange(rowNumber, 1, 1, HEADERS.length).getValues()[0];
  const status = String(payload.status || 'Issued').trim();
  const row = [
    existing[0],
    number,
    payload.date || '',
    payload.subject || '',
    payload.to || '',
    payload.from || '',
    status,
    payload.file_name || '',
    payload.local_file_path || '',
    new Date(),
    payload.remarks || ''
  ];
  sheet.getRange(rowNumber, 1, 1, HEADERS.length).setValues([row]);
  return jsonResponse({ ok: true, row_number: rowNumber, circular_number: number });
}

function cancelRecord(payload) {
  const number = String(payload.circular_number || '').trim();
  const sheet = getRegisterSheet();
  const rowNumber = findCircularRow(sheet, number);
  if (!rowNumber) throw new Error(`Circular ${number} was not found.`);
  sheet.getRange(rowNumber, 7).setValue('Cancelled');
  sheet.getRange(rowNumber, 11).setValue(payload.remarks || 'Cancelled');
  return jsonResponse({ ok: true, row_number: rowNumber });
}

function listRecords(payload) {
  const sheet = getRegisterSheet();
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return jsonResponse({ ok: true, records: [] });

  const rows = sheet.getRange(2, 1, lastRow - 1, HEADERS.length).getDisplayValues();
  const records = rows.map(r => ({
    serial_no: r[0], circular_number: r[1], date: r[2], subject: r[3],
    to: r[4], from: r[5], status: r[6], file_name: r[7],
    local_file_path: r[8], date_recorded: r[9], remarks: r[10]
  }));
  return jsonResponse({ ok: true, records: records.reverse() });
}

function findCircularRow(sheet, number) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return null;
  const values = sheet.getRange(2, 2, lastRow - 1, 1).getDisplayValues().flat();
  const index = values.findIndex(v => String(v).trim() === number);
  return index === -1 ? null : index + 2;
}

function jsonResponse(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
