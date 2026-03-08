// ============================================================
// ACCOUNTING TOOLS - Google Apps Script
// ============================================================
// Paste this entire file into the Apps Script editor
// (Extensions > Apps Script) and save. Reload your spreadsheet
// to see the "Accounting Tools" menu.
//
// WORKFLOW:
// 1. Click "Accounting Tools" > "Import CSV from staging tab"
//    (this creates the _CSV_Import tab if it doesn't exist)
// 2. Paste your bank CSV into the _CSV_Import tab, starting at A1
// 3. Click "Accounting Tools" > "Import CSV from staging tab" again
// 4. Data appears in the "main" tab with formatting applied
// 5. Fill in the "Category" column (G) manually
// ============================================================

// ============================================================
// SECTION 1: MENU SETUP
// ============================================================

function onOpen() {
  SpreadsheetApp.getActiveSpreadsheet()
    .addMenu('Accounting Tools', [
      { name: 'Import CSV from staging tab', functionName: 'runCsvImport' },
      { name: 'Clear staging tab', functionName: 'clearStagingTab' }
    ]);
}

function onInstall(e) {
  onOpen(e);
}

// ============================================================
// SECTION 2: SHEET MANAGEMENT
// ============================================================

var MAIN_SHEET_NAME    = 'main';
var STAGING_SHEET_NAME = '_CSV_Import';
var INDEX_SHEET_NAME   = '_ID_Index';

var MAIN_HEADERS = ['Date', 'Type', 'Description', 'Amount', 'Status', 'Month', 'Category'];

function getOrCreateMainSheet(ss) {
  var sheet = ss.getSheetByName(MAIN_SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(MAIN_SHEET_NAME);
    // Write header row
    sheet.getRange(1, 1, 1, MAIN_HEADERS.length).setValues([MAIN_HEADERS]);
    // Freeze header row
    sheet.setFrozenRows(1);
    // Style header row
    var headerRange = sheet.getRange(1, 1, 1, MAIN_HEADERS.length);
    headerRange.setFontWeight('bold');
    headerRange.setBackground('#d9d9d9');
    // Set column widths
    sheet.setColumnWidth(1, 100); // Date
    sheet.setColumnWidth(2, 80);  // Type
    sheet.setColumnWidth(3, 350); // Description
    sheet.setColumnWidth(4, 100); // Amount
    sheet.setColumnWidth(5, 80);  // Status
    sheet.setColumnWidth(6, 80);  // Month
    sheet.setColumnWidth(7, 120); // Category
  }
  return sheet;
}

function getOrCreateStagingSheet(ss) {
  var sheet = ss.getSheetByName(STAGING_SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(STAGING_SHEET_NAME);
    sheet.getRange('A1').setValue(
      'Paste your bank CSV here starting at cell A1 (overwrite this text). ' +
      'Row 1 must be the header row.'
    );
    sheet.setTabColor('#b7e1cd'); // light green to indicate staging area
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
  if (lastRow < 2) return; // No data rows yet

  var dataRows = lastRow - 1; // Exclude header

  // Column A: Date format
  mainSheet.getRange(2, 1, dataRows, 1).setNumberFormat('MM/dd/yyyy');

  // Column D: Amount - number format, negative in red
  var amountRange = mainSheet.getRange(2, 4, dataRows, 1);
  amountRange.setNumberFormat('#,##0.00');

  // Highlight negative amounts in red using conditional formatting
  var rules = mainSheet.getConditionalFormatRules();
  // Only add the rule once (check if it already exists)
  var ruleExists = rules.some(function(r) {
    var ranges = r.getRanges();
    return ranges.some(function(rng) {
      return rng.getColumn() === 4;
    });
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

  // Column G: Category - light yellow fill
  mainSheet.getRange(2, 7, dataRows, 1).setBackground('#fff9c4');
}

// ============================================================
// SECTION 3: DATA PARSING HELPERS
// ============================================================

/**
 * Parses a single quoted CSV line into an array of field values.
 * Handles: quoted fields, embedded commas inside quotes, escaped "" sequences.
 * Example: '"DATE","TRANSACTION TYPE","DESCRIPTION"' → ['DATE', 'TRANSACTION TYPE', 'DESCRIPTION']
 */
function parseCSVLine(line) {
  var fields = [];
  var current = '';
  var inQuotes = false;

  for (var i = 0; i < line.length; i++) {
    var ch = line[i];
    var nextCh = (i + 1 < line.length) ? line[i + 1] : '';

    if (ch === '"') {
      if (inQuotes && nextCh === '"') {
        // Escaped double-quote inside a quoted field: "" → "
        current += '"';
        i++; // skip the second quote
      } else {
        // Toggle quoted mode
        inQuotes = !inQuotes;
      }
    } else if (ch === ',' && !inQuotes) {
      // Field separator outside quotes
      fields.push(current);
      current = '';
    } else {
      current += ch;
    }
  }
  // Push the last field
  fields.push(current);

  return fields;
}

/**
 * Reads all rows from column A of the staging sheet where each cell contains
 * one raw CSV line (as happens when Sheets doesn't auto-parse the paste).
 * Calls parseCSVLine() on each non-empty row and returns a 2D array.
 * Returns null if nothing could be parsed.
 */
function reParseRawCSV(stagingSheet) {
  var lastRow = stagingSheet.getLastRow();
  if (lastRow < 1) return null;

  var rawData = stagingSheet.getRange(1, 1, lastRow, 1).getValues();
  var parsed = [];

  for (var i = 0; i < rawData.length; i++) {
    var cellText = String(rawData[i][0] || '').trim();
    if (cellText === '') continue; // skip blank rows
    var fields = parseCSVLine(cellText);
    parsed.push(fields);
  }

  return parsed.length > 0 ? parsed : null;
}

/**
 * Parses a currency string like "-$78.00" or "$2,200.00" to a float.
 * Returns NaN if unparseable.
 */
function parseCurrency(str) {
  if (!str || typeof str !== 'string') return NaN;
  str = str.trim();
  var isNegative = str.charAt(0) === '-';
  if (isNegative) str = str.substring(1);
  str = str.replace(/\$/g, '').replace(/,/g, '');
  var value = parseFloat(str);
  return isNegative ? -value : value;
}

/**
 * Parses a date string "MM/DD/YYYY" into a JavaScript Date.
 * Uses explicit year/month/day constructor to avoid timezone drift.
 * Returns null if unparseable.
 */
function parseSheetDate(str) {
  if (!str || typeof str !== 'string') return null;
  var parts = str.trim().split('/');
  if (parts.length !== 3) return null;
  var month = parseInt(parts[0], 10);
  var day   = parseInt(parts[1], 10);
  var year  = parseInt(parts[2], 10);
  if (isNaN(month) || isNaN(day) || isNaN(year)) return null;
  return new Date(year, month - 1, day); // month is 0-indexed in JS
}

/**
 * Returns "YYYY-MM" from a Date object for the Month column.
 */
function toMonthKey(dateObj) {
  if (!dateObj) return '';
  var year  = dateObj.getFullYear();
  var month = dateObj.getMonth() + 1;
  return year + '-' + (month < 10 ? '0' + month : month);
}

/**
 * Normalizes description text: trims and collapses runs of
 * whitespace (spaces, tabs) into a single space.
 */
function normalizeDescription(str) {
  if (!str) return '';
  return String(str).trim().replace(/\s+/g, ' ');
}

// ============================================================
// SECTION 4: DUPLICATE DETECTION
// ============================================================

/**
 * Loads all known transaction IDs from _ID_Index into a Set.
 */
function loadExistingIds(ss) {
  var indexSheet = getOrCreateIdIndexSheet(ss);
  var lastRow = indexSheet.getLastRow();
  var existingIds = new Set();
  if (lastRow < 2) return existingIds; // Only header row, no data
  var values = indexSheet.getRange(2, 1, lastRow - 1, 1).getValues();
  values.forEach(function(row) {
    if (row[0]) existingIds.add(String(row[0]));
  });
  return existingIds;
}

/**
 * Appends new transaction IDs to the _ID_Index sheet in one batch write.
 */
function saveNewIds(ss, newIds) {
  if (!newIds || newIds.length === 0) return;
  var indexSheet = getOrCreateIdIndexSheet(ss);
  var lastRow = indexSheet.getLastRow();
  var nextRow = lastRow + 1;
  var data = newIds.map(function(id) { return [id]; });
  indexSheet.getRange(nextRow, 1, data.length, 1).setValues(data);
}

/**
 * Returns the effective unique key for a row.
 * Uses the ID column if present; otherwise generates a composite key.
 */
function getRowKey(rowObj) {
  var id = rowObj['ID'] || rowObj['id'] || '';
  if (id && String(id).trim() !== '') {
    return String(id).trim();
  }
  // Fallback composite key
  var date   = rowObj['DATE'] || rowObj['Date'] || '';
  var desc   = rowObj['DESCRIPTION'] || rowObj['Description'] || '';
  var amount = rowObj['AMOUNT'] || rowObj['Amount'] || '';
  return [date, desc, amount].join('|');
}

// ============================================================
// SECTION 5: IMPORT ORCHESTRATOR
// ============================================================

/**
 * Main entry point: reads staging tab, processes rows, writes to main sheet.
 */
function runCsvImport() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  // Ensure all required tabs exist
  var stagingSheet = getOrCreateStagingSheet(ss);
  var mainSheet    = getOrCreateMainSheet(ss);

  // Read staging tab data
  var lastRow = stagingSheet.getLastRow();
  var lastCol = stagingSheet.getLastColumn();

  if (lastRow < 2) {
    SpreadsheetApp.getUi().alert(
      'No data found in the ' + STAGING_SHEET_NAME + ' tab.\n\n' +
      'Please paste your CSV content starting at cell A1, then run the import again.'
    );
    return;
  }

  // Read all values from staging tab
  var allValues = stagingSheet.getRange(1, 1, lastRow, lastCol).getValues();

  // --- Raw CSV detection ---
  // When a user pastes a quoted CSV string into Sheets, the entire string
  // lands in cell A1 as plain text (Sheets does NOT auto-parse it).
  // Detect this by checking if only column A has data AND A1 looks like a
  // CSV line. If so, re-parse column A rows ourselves.
  if (lastCol === 1) {
    var firstCell = String(allValues[0][0] || '');
    if (firstCell.indexOf('"') !== -1 && firstCell.indexOf(',') !== -1) {
      var reParsed = reParseRawCSV(stagingSheet);
      if (!reParsed || reParsed.length === 0) {
        SpreadsheetApp.getUi().alert(
          'Could not parse the CSV text in the ' + STAGING_SHEET_NAME + ' tab.\n\n' +
          'Please verify the CSV format is correct and try again.'
        );
        return;
      }
      allValues = reParsed;
    }
  }
  // --- End raw CSV detection ---

  // Build header → column index map from row 1
  var headerRow = allValues[0];
  var colIndex  = {};
  headerRow.forEach(function(header, idx) {
    colIndex[String(header).trim().toUpperCase()] = idx;
  });

  // Validate required columns are present
  var requiredCols = ['DATE', 'TRANSACTION TYPE', 'DESCRIPTION', 'AMOUNT', 'ID'];
  var missingCols = requiredCols.filter(function(col) {
    return colIndex[col] === undefined;
  });
  if (missingCols.length > 0) {
    SpreadsheetApp.getUi().alert(
      'Missing required columns in the CSV:\n  ' + missingCols.join(', ') + '\n\n' +
      'Please check that your CSV has these headers and try again.'
    );
    return;
  }

  // Load existing IDs for duplicate detection
  var existingIds = loadExistingIds(ss);

  // Process each data row
  var newRows  = [];
  var newIds   = [];
  var skipped  = 0;
  var errors   = [];

  var importBatch = Utilities.formatDate(
    new Date(),
    ss.getSpreadsheetTimeZone(),
    "yyyy-MM-dd'T'HH:mm:ss"
  );

  for (var r = 1; r < allValues.length; r++) {
    var row = allValues[r];

    // Skip completely empty rows
    if (row.every(function(cell) { return cell === '' || cell === null || cell === undefined; })) {
      continue;
    }

    // Build a row object keyed by column name
    var rowObj = {};
    headerRow.forEach(function(header, idx) {
      rowObj[String(header).trim().toUpperCase()] = row[idx];
    });

    // Duplicate check
    var key = getRowKey(rowObj);
    if (existingIds.has(key)) {
      skipped++;
      continue;
    }

    // Parse fields
    var dateStr  = String(rowObj['DATE'] || '').trim();
    var dateObj  = parseSheetDate(dateStr);
    if (!dateObj) {
      errors.push('Row ' + (r + 1) + ': could not parse date "' + dateStr + '"');
      continue;
    }

    var amountRaw = rowObj['AMOUNT'];
    // Sheets may have already parsed the currency to a number during paste
    var amount;
    if (typeof amountRaw === 'number') {
      amount = amountRaw;
    } else {
      amount = parseCurrency(String(amountRaw || ''));
      if (isNaN(amount)) {
        errors.push('Row ' + (r + 1) + ': could not parse amount "' + amountRaw + '"');
        continue;
      }
    }

    var type        = String(rowObj['TRANSACTION TYPE'] || '').trim();
    var description = normalizeDescription(rowObj['DESCRIPTION']);
    var status      = String(rowObj['STATUS'] || '').trim();
    var monthKey    = toMonthKey(dateObj);

    // Build output row matching MAIN_HEADERS order:
    // Date | Type | Description | Amount | Status | Month | Category
    newRows.push([dateObj, type, description, amount, status, monthKey, '']);
    newIds.push(key);
    existingIds.add(key); // Prevent intra-batch duplicates
  }

  // Batch write new rows to main sheet
  if (newRows.length > 0) {
    var firstNewRow = mainSheet.getLastRow() + 1;
    mainSheet
      .getRange(firstNewRow, 1, newRows.length, MAIN_HEADERS.length)
      .setValues(newRows);
    saveNewIds(ss, newIds);
    applyMainSheetFormatting(mainSheet);
  }

  // Show summary
  var summary = 'Import complete!\n\n' +
    '  Rows imported: ' + newRows.length + '\n' +
    '  Duplicates skipped: ' + skipped;
  if (errors.length > 0) {
    summary += '\n  Errors (' + errors.length + '):\n    ' + errors.slice(0, 5).join('\n    ');
    if (errors.length > 5) summary += '\n    ... and ' + (errors.length - 5) + ' more';
  }
  SpreadsheetApp.getUi().alert(summary);
}

// ============================================================
// SECTION 6: UTILITY
// ============================================================

/**
 * Clears the staging tab content (keeps the tab, removes data).
 */
function clearStagingTab() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var stagingSheet = ss.getSheetByName(STAGING_SHEET_NAME);
  if (!stagingSheet) {
    SpreadsheetApp.getUi().alert(STAGING_SHEET_NAME + ' tab not found.');
    return;
  }
  stagingSheet.clearContents();
  stagingSheet.getRange('A1').setValue(
    'Paste your bank CSV here starting at cell A1. Row 1 must be the header row.'
  );
  SpreadsheetApp.getActiveSpreadsheet().toast('Staging tab cleared.', 'Done', 3);
}
