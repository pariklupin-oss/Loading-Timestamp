/**
 * Loading Log connector for the existing AppSheet source spreadsheet.
 * Bind this script to the spreadsheet with LOADING_HEADER and LOADING_ITEMS,
 * then deploy as a web app that runs as the signed-in user accessing it.
 */
const LOADING_WEB_TABLES = {
  header: {
    sheet: 'LOADING_HEADER',
    fields: ['LOADING_ID', 'DATE', 'SHIFT', 'LOADING INCHARGE', 'HELPER COUNT']
  },
  item: {
    sheet: 'LOADING_ITEMS',
    fields: ['ITEM_ID', 'LOADING_ID', 'Customer', 'LOADING START', 'LOADING END', 'TOTAL_HRS', 'VEHICLE NO', 'VEHICL FEET', 'Remarks']
  },
  stage: {
    sheet: 'STAGE_TIME',
    fields: ['Date', 'Customer Name', 'Item Name', 'Item Code', 'Invoice Number', 'OQC End', 'Loading End', 'Invoice Time', 'Gate Out Time', 'GAP-1 (Production and Loading)', 'GAP-2 (Loading and Billing)', 'GAP-3 (Billing and Gate Out)']
  }
};

function doGet(e) {
  const callback = String(e && e.parameter && e.parameter.callback || '');
  if (!/^[A-Za-z_$][\w.$]{0,100}$/.test(callback)) return loadingWebText_('Invalid callback.');
  try {
    const result = loadingWebLoad_();
    return ContentService.createTextOutput(callback + '(' + JSON.stringify(result) + ');')
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  } catch (err) {
    const safe = { ok: false, error: String(err && err.message || err) };
    return ContentService.createTextOutput(callback + '(' + JSON.stringify(safe) + ');')
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
}

function doPost(e) {
  try {
    const body = JSON.parse(e && e.postData && e.postData.contents || '{}');
    if (body.action !== 'save' || !body.header || !Array.isArray(body.items)) throw new Error('Invalid save request.');
    const lock = LockService.getDocumentLock();
    lock.waitLock(15000);
    try { loadingWebSave_(body); } finally { lock.releaseLock(); }
    return loadingWebText_(JSON.stringify({ ok: true, id: String(body.header.id) }));
  } catch (err) {
    return loadingWebText_(JSON.stringify({ ok: false, error: String(err && err.message || err) }));
  }
}

function loadingWebLoad_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) throw new Error('Bind this script to the LOADING TIMESTAMP Google Sheet first.');
  const tz = ss.getSpreadsheetTimeZone();
  const headerSheet = loadingWebRequireTable_(ss, LOADING_WEB_TABLES.header);
  const itemSheet = loadingWebRequireTable_(ss, LOADING_WEB_TABLES.item);
  const stageSheet = loadingWebRequireTable_(ss, LOADING_WEB_TABLES.stage);
  const customerSheet = ss.getSheetByName('CUSTOMER_MASTER');
  const headers = loadingWebReadRows_(headerSheet, LOADING_WEB_TABLES.header, tz).map(r => ({
    id: r.LOADING_ID,
    date: r.DATE,
    shift: r.SHIFT,
    incharge: r['LOADING INCHARGE'],
    helperCount: Number(r['HELPER COUNT']) || 0,
    createdAt: r.DATE
  })).filter(r => r.id);
  const items = loadingWebReadRows_(itemSheet, LOADING_WEB_TABLES.item, tz).map(r => ({
    id: r.ITEM_ID,
    loadingId: r.LOADING_ID,
    customer: r.Customer,
    start: r['LOADING START'],
    end: r['LOADING END'],
    totalHours: Number(r.TOTAL_HRS) || 0,
    vehicleNo: r['VEHICLE NO'],
    vehicleFeet: r['VEHICL FEET'],
    remarks: r.Remarks
  })).filter(r => r.id && r.loadingId);
  const stageRows = loadingWebReadRows_(stageSheet, LOADING_WEB_TABLES.stage, tz).map(r => ({
    date: r.Date,
    customerName: r['Customer Name'],
    itemName: r['Item Name'],
    itemCode: r['Item Code'],
    invoiceNumber: r['Invoice Number'],
    oqcEnd: r['OQC End'],
    loadingEnd: r['Loading End'],
    invoiceTime: r['Invoice Time'],
    gateOutTime: r['Gate Out Time'],
    gap1: r['GAP-1 (Production and Loading)'],
    gap2: r['GAP-2 (Loading and Billing)'],
    gap3: r['GAP-3 (Billing and Gate Out)']
  })).filter(r => r.invoiceNumber || r.itemCode || r.itemName);
  const customerNames = customerSheet && customerSheet.getLastRow() > 1
    ? [...new Set(customerSheet.getRange(2, 1, customerSheet.getLastRow() - 1, 1).getDisplayValues().flat().map(name => String(name || '').trim().replace(/\s+/g, ' ')).filter(Boolean))]
    : [];
  return { ok: true, headers: headers, items: items, stageRows: stageRows, customerNames: customerNames };
}

function loadingWebSave_(body) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) throw new Error('Bind this script to the LOADING TIMESTAMP Google Sheet first.');
  const headerSheet = loadingWebRequireTable_(ss, LOADING_WEB_TABLES.header);
  const itemSheet = loadingWebRequireTable_(ss, LOADING_WEB_TABLES.item);
  const h = body.header;
  if (!h.id || !h.date || !h.shift || !h.incharge) throw new Error('Required shift fields are missing.');
  const headerValues = {
    LOADING_ID: String(h.id),
    DATE: loadingWebParseDate_(h.date),
    SHIFT: String(h.shift),
    'LOADING INCHARGE': String(h.incharge),
    'HELPER COUNT': Number(h.helperCount) || 0
  };
  loadingWebWriteByKey_(headerSheet, LOADING_WEB_TABLES.header, headerValues, 'LOADING_ID');
  body.items.forEach(item => {
    if (!item.id || !item.vehicleNo || !item.customer || !item.start || !item.end) throw new Error('A vehicle row is missing a required value.');
    const values = {
      ITEM_ID: String(item.id),
      LOADING_ID: String(h.id),
      Customer: String(item.customer),
      'LOADING START': loadingWebTimeFraction_(item.start),
      'LOADING END': loadingWebTimeFraction_(item.end),
      TOTAL_HRS: Number(item.totalHours) || 0,
      'VEHICLE NO': String(item.vehicleNo),
      'VEHICL FEET': String(item.vehicleFeet || ''),
      Remarks: String(item.remarks || '')
    };
    loadingWebWriteByKey_(itemSheet, LOADING_WEB_TABLES.item, values, 'ITEM_ID');
  });
}

function loadingWebRequireTable_(ss, definition) {
  const sheet = ss.getSheetByName(definition.sheet);
  if (!sheet) throw new Error('Required sheet tab not found: ' + definition.sheet);
  const lastCol = Math.max(sheet.getLastColumn(), 1);
  const names = sheet.getRange(1, 1, 1, lastCol).getDisplayValues()[0].map(v => String(v).trim().toUpperCase());
  const missing = definition.fields.filter(field => names.indexOf(field.toUpperCase()) < 0);
  if (missing.length) throw new Error('Missing columns in ' + definition.sheet + ': ' + missing.join(', '));
  return sheet;
}

function loadingWebReadRows_(sheet, definition, timezone) {
  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  if (lastRow < 2 || !lastCol) return [];
  const raw = sheet.getRange(1, 1, lastRow, lastCol).getValues();
  const display = sheet.getRange(1, 1, lastRow, lastCol).getDisplayValues();
  const names = raw[0].map(v => String(v).trim().toUpperCase());
  return raw.slice(1).map((row, rowIndex) => {
    const item = {};
    definition.fields.forEach(field => {
      const col = names.indexOf(field.toUpperCase());
      const value = row[col];
      if (field.toUpperCase() === 'DATE') item[field] = loadingWebNormalizeDate_(value, display[rowIndex + 1][col], timezone);
      else if (['OQC END', 'LOADING END', 'INVOICE TIME', 'GATE OUT TIME'].indexOf(field.toUpperCase()) >= 0) item[field] = loadingWebNormalizeTimestamp_(value, display[rowIndex + 1][col], timezone);
      else if (field === 'LOADING START' || field === 'LOADING END') item[field] = String(display[rowIndex + 1][col] || '').trim().slice(0, 5);
      else item[field] = value instanceof Date ? Utilities.formatDate(value, timezone, 'HH:mm:ss') : String(display[rowIndex + 1][col] || '').trim();
    });
    return item;
  });
}

function loadingWebWriteByKey_(sheet, definition, values, keyName) {
  const lastCol = sheet.getLastColumn();
  const names = sheet.getRange(1, 1, 1, lastCol).getDisplayValues()[0].map(v => String(v).trim().toUpperCase());
  const keyCol = names.indexOf(keyName.toUpperCase()) + 1;
  if (!keyCol) throw new Error('Key column not found: ' + keyName);
  const keyValue = String(values[keyName]);
  const rowCount = Math.max(sheet.getLastRow() - 1, 0);
  const existing = rowCount ? sheet.getRange(2, keyCol, rowCount, 1).getDisplayValues().flat() : [];
  const match = existing.findIndex(value => String(value).trim() === keyValue);
  const targetRow = match < 0 ? Math.max(sheet.getLastRow() + 1, 2) : match + 2;
  Object.keys(values).forEach(field => {
    const col = names.indexOf(field.toUpperCase()) + 1;
    if (!col) throw new Error('Column not found in ' + definition.sheet + ': ' + field);
    const cell = sheet.getRange(targetRow, col);
    cell.setValue(values[field]);
    if (field === 'DATE') cell.setNumberFormat('yyyy-mm-dd');
    if (field === 'LOADING START' || field === 'LOADING END') cell.setNumberFormat('hh:mm:ss');
  });
}

function loadingWebParseDate_(value) {
  const match = String(value).match(/^(\d{4})-(\d{2})-(\d{2})$/);
  if (!match) throw new Error('Date must use YYYY-MM-DD.');
  return new Date(Number(match[1]), Number(match[2]) - 1, Number(match[3]));
}

function loadingWebNormalizeDate_(value, shown, timezone) {
  if (value instanceof Date) return Utilities.formatDate(value, timezone, 'yyyy-MM-dd');
  const text = String(shown || '').trim();
  if (/^\d{4}-\d{2}-\d{2}$/.test(text)) return text;
  const dmy = text.match(/^(\d{1,2})[-/.](\d{1,2})[-/.](\d{4})$/);
  if (dmy) return dmy[3] + '-' + dmy[2].padStart(2, '0') + '-' + dmy[1].padStart(2, '0');
  return text;
}

function loadingWebNormalizeTimestamp_(value, shown, timezone) {
  if (value instanceof Date) return Utilities.formatDate(value, timezone, 'yyyy-MM-dd HH:mm:ss');
  const text = String(shown || '').trim();
  const iso = text.match(/^(\d{4})-(\d{2})-(\d{2})[ T](\d{2}:\d{2})(?::(\d{2}))?$/);
  if (iso) return iso[1] + '-' + iso[2] + '-' + iso[3] + ' ' + iso[4] + ':' + (iso[5] || '00');
  const dmy = text.match(/^(\d{1,2})[-/.](\d{1,2})[-/.](\d{4})\s+(\d{1,2}:\d{2})(?::(\d{2}))?$/);
  if (dmy) return dmy[3] + '-' + dmy[2].padStart(2, '0') + '-' + dmy[1].padStart(2, '0') + ' ' + dmy[4].padStart(5, '0') + ':' + (dmy[5] || '00');
  return text;
}

function loadingWebTimeFraction_(value) {
  const match = String(value).match(/^(\d{1,2}):(\d{2})(?::(\d{2}))?$/);
  if (!match) throw new Error('Time must use HH:MM.');
  const hours = Number(match[1]), minutes = Number(match[2]), seconds = Number(match[3] || 0);
  if (hours > 23 || minutes > 59 || seconds > 59) throw new Error('Invalid time value.');
  return (hours * 3600 + minutes * 60 + seconds) / 86400;
}

function loadingWebText_(value) {
  return ContentService.createTextOutput(value).setMimeType(ContentService.MimeType.TEXT);
}
