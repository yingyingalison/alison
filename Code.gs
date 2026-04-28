/**
 * Sales Follow Up — web app. Uses the "Follow Up" tab: A–H as your original columns,
 * I–Q for app fields (Created, reply, type, person, contact, proposal fields).
 * Spreadsheet: https://docs.google.com/spreadsheets/d/1hVAzdNadD2WJUj3kJILDy6g3MmTC4--p7m-oY_uiyUA
 */
var SPREADSHEET_ID = '1hVAzdNadD2WJUj3kJILDy6g3MmTC4--p7m-oY_uiyUA';
var SHEET_NAME = 'Follow Up';
var COLS = 17; // A–Q

var LEGACY_HEADERS_ = [
  'Lead Name',
  'Phone',
  'Address',
  'Date Acquired',
  'Note Average Bill',
  'Note Call',
  'Note 1',
  'Note2'
];
var APP_HEADERS_ = [
  'Created',
  'Customer Reply',
  'Customer Type',
  'Person Follow Up',
  'Saved Contact in Phone',
  'Proposal Size',
  'Proposal Price',
  'Panel Pcs',
  'Expected Saving'
];

function onOpen() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) {
    return;
  }
  ss
    .getUi()
    .createMenu('Follow Up App')
    .addItem('Open web app', 'showWebAppUrl')
    .addToUi();
}

function showWebAppUrl() {
  var url = ScriptApp.getService().getUrl();
  if (!url) {
    SpreadsheetApp.getUi().alert('Deploy the web app first: Deploy → New deployment → Web app.');
    return;
  }
  SpreadsheetApp.getUi().alert('Web app URL:\n' + url);
}

function doGet() {
  ensureSheet_();
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle('Sales Follow Up')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function getSpreadsheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  if (ss) {
    return ss;
  }
  return SpreadsheetApp.openById(SPREADSHEET_ID);
}

function ensureSheet_() {
  var ss = getSpreadsheet_();
  var sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) {
    sh = ss.insertSheet(SHEET_NAME);
    sh.getRange(1, 1, 1, 8).setValues([LEGACY_HEADERS_]);
  }
  var hasExt = sh.getRange(1, 9).getValue();
  if (!hasExt || String(hasExt).trim() === '') {
    sh.getRange(1, 9, 1, COLS).setValues([APP_HEADERS_]);
    sh.getRange(1, 9, 1, COLS).setFontWeight('bold').setBackground('#fef3c7');
  }
  return sh;
}

function isSheetDate_(v) {
  return v && Object.prototype.toString.call(v) === '[object Date]' && !isNaN(v.getTime());
}

function padRow_(r) {
  var a = r.slice();
  while (a.length < COLS) {
    a.push('');
  }
  return a;
}

function isRowVisiblyEmpty_(r) {
  for (var k = 0; k < r.length; k++) {
    var v = r[k];
    if (v != null && v !== '' && String(v).trim() !== '') {
      return false;
    }
  }
  return true;
}

/**
 * Rows without Created (column I) are "legacy" for Note Call / Note 1 layout,
 * but columns J–Q are still read when you type them manually in the sheet.
 * @param {Array} r
 * @param {number} rowIndex — 1-based sheet row
 */
function mapRowToItem_(r, rowIndex) {
  r = padRow_(r);
  var hasCreated = isSheetDate_(r[8]);
  var replyStr = String(r[9] == null ? '' : r[9]).trim();
  var item = {
    id: rowIndex,
    leadName: r[0],
    customerPhone: r[1],
    address: r[2],
    dateAcquired: null,
    dateAcquiredText: null,
    noteAverageBill: r[4],
    customerReply: replyStr,
    customerType: String(r[10] == null ? '' : r[10]).trim(),
    personFollowUp: String(r[11] == null ? '' : r[11]).trim(),
    savedContact: String(r[12] == null ? '' : r[12]).trim(),
    proposalSize: String(r[13] == null ? '' : r[13]).trim(),
    proposalPrice: String(r[14] == null ? '' : r[14]).trim(),
    panelPcs: String(r[15] == null ? '' : r[15]).trim(),
    expectedSaving: String(r[16] == null ? '' : r[16]).trim(),
    note: r[7],
    created: hasCreated && isSheetDate_(r[8]) ? r[8].getTime() : null,
    isLegacy: !hasCreated
  };
  if (r[3]) {
    if (isSheetDate_(r[3])) {
      item.dateAcquired = r[3].getTime();
    } else {
      item.dateAcquiredText = String(r[3]);
    }
  }
  if (!hasCreated) {
    item.noteCall = r[5];
    item.note1 = r[6];
  }
  if (replyStr) {
    item.chipLabel = replyStr;
  } else if (!hasCreated) {
    item.chipLabel = 'Sheet';
  } else {
    item.chipLabel = '—';
  }
  return item;
}

/**
 * @param {Object} data
 */
function saveFollowUp(data) {
  var sh = ensureSheet_();
  var now = new Date();
  var dAcq = data.dateAcquired ? new Date(data.dateAcquired) : '';
  var row = [
    String(data.leadName || '').trim(),
    String(data.customerPhone || '').trim(),
    String(data.address || '').trim(),
    dAcq,
    String(data.noteAverageBill || '').trim(),
    String(data.noteCall || '').trim(),
    '', // column G (Note 1 - empty)
    String(data.note || '').trim(), // column H (Note 2)
    now,
    String(data.customerReply || '').trim(),
    String(data.customerType || '').trim(),
    String(data.personFollowUp || '').trim(),
    String(data.savedContact || '').trim(),
    String(data.proposalSize || '').trim(),
    String(data.proposalPrice || '').trim(),
    String(data.panelPcs || '').trim(),
    String(data.expectedSaving || '').trim()
  ];
  sh.appendRow(row);
  return { ok: true };
}

/**
 * Update customer reply status for an existing row in sheet.
 * @param {number} rowId - 1-based row number in "Follow Up" sheet.
 * @param {string} customerReply
 * @returns {{ok: boolean, rowId: number, customerReply: string}}
 */
function updateCustomerReply(rowId, customerReply) {
  var sh = ensureSheet_();
  var row = Number(rowId);
  if (!row || row < 2) {
    throw new Error('Invalid row to update.');
  }
  var reply = String(customerReply || '').trim();
  if (!reply) {
    throw new Error('Customer Reply is required.');
  }
  // Column J is Customer Reply in this sheet layout.
  sh.getRange(row, 10).setValue(reply);
  return { ok: true, rowId: row, customerReply: reply };
}

/**
 * Newest sheet rows first (last row in spreadsheet appears first in list).
 * @returns {Array<Object>}
 */
function getFollowUps() {
  var sh = ensureSheet_();
  var last = sh.getLastRow();
  if (last < 2) {
    return [];
  }
  var lastCol = sh.getLastColumn();
  var numCols = Math.max(COLS, lastCol);
  var range = sh.getRange(2, 1, last, numCols);
  var values = range.getValues();
  var out = [];
  for (var i = values.length - 1; i >= 0; i--) {
    var r = values[i];
    if (isRowVisiblyEmpty_(r)) {
      continue;
    }
    out.push(mapRowToItem_(r, i + 2));
  }
  return out;
}
