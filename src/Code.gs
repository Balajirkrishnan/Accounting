// ============================================================
// ACCOUNTING TOOLS - Google Apps Script
// ============================================================
// Paste this entire file into the Apps Script editor
// (Extensions > Apps Script) and save. Reload your spreadsheet
// to see the "Accounting Tools" menu.
//
// NEW WORKFLOW:
// 1. Click "Accounting Tools" > "Import CSV..."
// 2. A dialog opens — paste your CSV text directly into it
// 3. Click "Import" in the dialog
// 4. Data appears in the "main" tab with formatting applied
// 5. Fill in the "Category" column (G) manually
// ============================================================

// ============================================================
// SECTION 1: MENU SETUP
// ============================================================

function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Accounting Tools')
    .addItem('Import CSV...', 'showImportDialog')
    .addToUi();
}

function onInstall(e) {
  onOpen(e);
}

// ============================================================
// SECTION 2: SHEET MANAGEMENT
// ============================================================

var MAIN_SHEET_NAME  = 'main';
var INDEX_SHEET_NAME = '_ID_Index';

var MAIN_HEADERS = ['Date', 'Type', 'Description', 'Amount', 'Status', 'Month', 'Category'];

function getOrCreateMainSheet(ss) {
  var sheet = ss.getSheetByName(MAIN_SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(MAIN_SHEET_NAME);
    sheet.getRange(1, 1, 1, MAIN_HEADERS.length).setValues([MAIN_HEADERS]);
    sheet.setFrozenRows(1);
    var headerRange = sheet.getRange(1, 1, 1, MAIN_HEADERS.length);
    headerRange.setFontWeight('bold');
    headerRange.setBackground('#d9d9d9');
    sheet.setColumnWidth(1, 100);
    sheet.setColumnWidth(2, 80);
    sheet.setColumnWidth(3, 350);
    sheet.setColumnWidth(4, 100);
    sheet.setColumnWidth(5, 80);
    sheet.setColumnWidth(6, 80);
    sheet.setColumnWidth(7, 120);
  }
  return sheet;
}

function getOrCreateIdIndexSheet(ss) {
  var sheet = ss.getSheetByName(INDEX_SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(INDEX_SHEET_NAME);
    sheet.getRange('A1').setValue('transaction_id');
    sheet.hideSheet();
  }
  return sheet;
}

function applyMainSheetFormatting(mainSheet) {
  var lastRow = mainSheet.getLastRow();
  if (lastRow < 2) return;
  var dataRows = lastRow - 1;

  mainSheet.getRange(2, 1, dataRows, 1).setNumberFormat('MM/dd/yyyy');
  mainSheet.getRange(2, 4, dataRows, 1).setNumberFormat('#,##0.00');

  var rules = mainSheet.getConditionalFormatRules();
  var ruleExists = rules.some(function(r) {
    return r.getRanges().some(function(rng) { return rng.getColumn() === 4; });
  });
  if (!ruleExists) {
    var negativeRule = SpreadsheetApp.newConditionalFormatRule()
      .whenNumberLessThan(0)
      .setFontColor('#cc0000')
      .setRanges([mainSheet.getRange(2, 4, mainSheet.getMaxRows() - 1, 1)])
      .build();
    rules.push(negativeRule);
    mainSheet.setConditionalFormatRules(rules);
  }

  mainSheet.getRange(2, 7, dataRows, 1).setBackground('#fff9c4');
}

// ============================================================
// SECTION 3: CSV PARSING (pure string — no Sheets involvement)
// ============================================================

/**
 * Parses a full CSV string into a 2D array of clean string values.
 * Handles quoted fields, embedded commas, embedded newlines, and "" escapes.
 * This runs entirely on the raw text — Sheets never touches the values.
 */
function parseCSVText(csvText) {
  var rows = [];
  var row  = [];
  var current = '';
  var inQuotes = false;
  // Normalise line endings
  var text = csvText.replace(/\r\n/g, '\n').replace(/\r/g, '\n');

  for (var i = 0; i < text.length; i++) {
    var ch     = text[i];
    var nextCh = (i + 1 < text.length) ? text[i + 1] : '';

    if (inQuotes) {
      if (ch === '"' && nextCh === '"') {
        current += '"';   // escaped quote "" → "
        i++;
      } else if (ch === '"') {
        inQuotes = false; // closing quote
      } else {
        current += ch;    // normal char inside quotes (including newlines)
      }
    } else {
      if (ch === '"') {
        inQuotes = true;  // opening quote
      } else if (ch === ',') {
        row.push(current.trim());
        current = '';
      } else if (ch === '\n') {
        row.push(current.trim());
        current = '';
        if (row.some(function(f) { return f !== ''; })) {
          rows.push(row); // only push non-blank rows
        }
        row = [];
      } else {
        current += ch;
      }
    }
  }
  // Last field / last row
  row.push(current.trim());
  if (row.some(function(f) { return f !== ''; })) {
    rows.push(row);
  }

  return rows;
}

// ============================================================
// SECTION 4: DATA HELPERS
// ============================================================

function parseCurrency(str) {
  if (!str) return NaN;
  str = String(str).trim();
  var isNegative = str.charAt(0) === '-';
  if (isNegative) str = str.substring(1);
  str = str.replace(/\$/g, '').replace(/,/g, '');
  var value = parseFloat(str);
  return isNegative ? -value : value;
}

function parseSheetDate(str) {
  if (!str) return null;
  var parts = String(str).trim().split('/');
  if (parts.length !== 3) return null;
  var month = parseInt(parts[0], 10);
  var day   = parseInt(parts[1], 10);
  var year  = parseInt(parts[2], 10);
  if (isNaN(month) || isNaN(day) || isNaN(year)) return null;
  return new Date(year, month - 1, day);
}

function toMonthKey(dateObj) {
  if (!dateObj) return '';
  var y = dateObj.getFullYear();
  var m = dateObj.getMonth() + 1;
  return y + '-' + (m < 10 ? '0' + m : m);
}

function normalizeDescription(str) {
  if (!str) return '';
  return String(str).trim().replace(/\s+/g, ' ');
}

// ============================================================
// SECTION 5: DUPLICATE DETECTION
// ============================================================

function loadExistingIds(ss) {
  var indexSheet = getOrCreateIdIndexSheet(ss);
  var lastRow = indexSheet.getLastRow();
  var existingIds = new Set();
  if (lastRow < 2) return existingIds;
  var values = indexSheet.getRange(2, 1, lastRow - 1, 1).getValues();
  values.forEach(function(r) { if (r[0]) existingIds.add(String(r[0])); });
  return existingIds;
}

function saveNewIds(ss, newIds) {
  if (!newIds || newIds.length === 0) return;
  var indexSheet = getOrCreateIdIndexSheet(ss);
  var nextRow = indexSheet.getLastRow() + 1;
  indexSheet.getRange(nextRow, 1, newIds.length, 1).setValues(
    newIds.map(function(id) { return [id]; })
  );
}

function getRowKey(rowObj) {
  var id = rowObj['ID'] || '';
  if (id.trim() !== '') return id.trim();
  return [rowObj['DATE'] || '', rowObj['DESCRIPTION'] || '', rowObj['AMOUNT'] || ''].join('|');
}

// ============================================================
// SECTION 6: HTML DIALOG — paste CSV as raw text
// ============================================================

function showImportDialog() {
  var html = HtmlService.createHtmlOutput(
    '<!DOCTYPE html>' +
    '<html>' +
    '<head>' +
    '<style>' +
    '  body { font-family: Arial, sans-serif; margin: 12px; font-size: 13px; }' +
    '  textarea { width: 100%; height: 320px; font-size: 11px; font-family: monospace;' +
    '             box-sizing: border-box; border: 1px solid #ccc; padding: 6px; }' +
    '  button { margin-top: 10px; padding: 8px 20px; background: #1a73e8; color: white;' +
    '           border: none; border-radius: 4px; cursor: pointer; font-size: 13px; }' +
    '  button:hover { background: #1558b0; }' +
    '  #status { margin-top: 8px; font-size: 12px; color: #555; }' +
    '</style>' +
    '</head>' +
    '<body>' +
    '<p style="margin:0 0 6px 0"><strong>Paste your bank CSV below</strong> (include the header row):</p>' +
    '<textarea id="csv" placeholder="Paste CSV text here..."></textarea>' +
    '<br>' +
    '<button onclick="doImport()">Import</button>' +
    '<div id="status"></div>' +
    '<script>' +
    'function doImport() {' +
    '  var csv = document.getElementById("csv").value.trim();' +
    '  if (!csv) { document.getElementById("status").innerText = "Please paste CSV text first."; return; }' +
    '  document.getElementById("status").innerText = "Importing...";' +
    '  google.script.run' +
    '    .withSuccessHandler(function(msg) {' +
    '      document.getElementById("status").innerText = msg;' +
    '    })' +
    '    .withFailureHandler(function(err) {' +
    '      document.getElementById("status").innerText = "Error: " + err.message;' +
    '    })' +
    '    .processCSVText(csv);' +
    '}' +
    '</script>' +
    '</body>' +
    '</html>'
  )
  .setWidth(560)
  .setHeight(460)
  .setTitle('Import CSV');

  SpreadsheetApp.getUi().showModalDialog(html, 'Import CSV');
}

// ============================================================
// SECTION 7: IMPORT PROCESSOR (called from the dialog)
// ============================================================

/**
 * Called from the HTML dialog with the raw CSV text string.
 * Parses it entirely in-script — no Sheets cell mangling.
 * Returns a status string shown in the dialog.
 */
function processCSVText(csvText) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var mainSheet = getOrCreateMainSheet(ss);

  // Parse the raw CSV string into a 2D array
  var rows = parseCSVText(csvText);

  if (rows.length < 2) {
    return 'No data found. Please paste a CSV with a header row and at least one data row.';
  }

  // Build header → column index map
  var headerRow = rows[0];
  var colIndex  = {};
  headerRow.forEach(function(h, idx) {
    colIndex[h.trim().toUpperCase()] = idx;
  });

  // Validate required columns
  var required = ['DATE', 'TRANSACTION TYPE', 'DESCRIPTION', 'AMOUNT', 'ID'];
  var missing  = required.filter(function(c) { return colIndex[c] === undefined; });
  if (missing.length > 0) {
    return 'Missing required columns: ' + missing.join(', ') +
           '\nFound: ' + Object.keys(colIndex).join(', ');
  }

  // Load existing IDs
  var existingIds = loadExistingIds(ss);

  var newRows  = [];
  var newIds   = [];
  var skipped  = 0;
  var errors   = [];

  for (var r = 1; r < rows.length; r++) {
    var row = rows[r];

    // Build rowObj
    var rowObj = {};
    headerRow.forEach(function(h, idx) {
      rowObj[h.trim().toUpperCase()] = row[idx] !== undefined ? row[idx] : '';
    });

    // Duplicate check
    var key = getRowKey(rowObj);
    if (existingIds.has(key)) { skipped++; continue; }

    // Parse date
    var dateObj = parseSheetDate(rowObj['DATE']);
    if (!dateObj) {
      errors.push('Row ' + (r + 1) + ': bad date "' + rowObj['DATE'] + '"');
      continue;
    }

    // Parse amount
    var amount = parseCurrency(rowObj['AMOUNT']);
    if (isNaN(amount)) {
      errors.push('Row ' + (r + 1) + ': bad amount "' + rowObj['AMOUNT'] + '"');
      continue;
    }

    var type        = rowObj['TRANSACTION TYPE'] || '';
    var description = normalizeDescription(rowObj['DESCRIPTION']);
    var status      = rowObj['STATUS'] || '';
    var monthKey    = toMonthKey(dateObj);

    newRows.push([dateObj, type, description, amount, status, monthKey, '']);
    newIds.push(key);
    existingIds.add(key);
  }

  if (newRows.length > 0) {
    var firstNewRow = mainSheet.getLastRow() + 1;
    mainSheet.getRange(firstNewRow, 1, newRows.length, MAIN_HEADERS.length).setValues(newRows);
    saveNewIds(ss, newIds);
    applyMainSheetFormatting(mainSheet);
  }

  var summary = 'Done!  Imported: ' + newRows.length + '  |  Duplicates skipped: ' + skipped;
  if (errors.length > 0) {
    summary += '\nErrors (' + errors.length + '): ' + errors.slice(0, 3).join('; ');
    if (errors.length > 3) summary += ' ... and ' + (errors.length - 3) + ' more';
  }
  return summary;
}
