// ============================================================
// JAG Life Group Roster - Google Apps Script Backend
// Spreadsheet: https://docs.google.com/spreadsheets/d/1Cg9m7lUu536JlSXbY4HifWQpOw9nQ2DtBRDZRzIXIn4
// Version: 1.20.3 (2026-04-06)
// ============================================================

const VERSION      = '1.20.3';
const VERSION_DATE = '2026-04-06';

const SPREADSHEET_ID    = '1Cg9m7lUu536JlSXbY4HifWQpOw9nQ2DtBRDZRzIXIn4';
const ROSTER_SHEET_NAME = 'Roster';   // year-agnostic — supports 2026 and beyond
const MEMBERS_SHEET_NAME = 'Members';

// ---- Entry Point ----

function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('JAG LG Roster v' + VERSION)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1.0')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function getVersion() {
  return { version: VERSION, date: VERSION_DATE };
}

// ---- Data Fetching ----

function getAllData() {
  // B: open spreadsheet once and share it — avoids two separate openById calls
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  return {
    entries:     getRosterEntries(ss),
    members:     getMembers(ss),
    version:     VERSION,      // A: version piggybacked on the data call — eliminates getVersion() round-trip
    versionDate: VERSION_DATE
  };
}

// Maps header text → column index so reads are schema-version-agnostic.
// Adding/reordering columns in the sheet never breaks existing reads.
function _rosterColMap(headers) {
  const m = {};
  headers.forEach(function(h, i) {
    switch (String(h).toLowerCase().trim()) {
      case 'date':             m.date = i;        break;
      case 'group':            m.group = i;       break;
      case 'event type':       m.eventType = i;   break;
      case 'venue':            m.venue = i;       break;
      case 'organiser':        m.organiser = i;   break;
      case 'p&w':              m.pw = i;          break;
      case 'facilitator':      m.facilitator = i; break;
      case 'food preparation':
      case 'food':             m.food = i;        break;
      case 'reporting':        m.reporting = i;   break;
      case 'notes':            m.notes = i;       break;
      case 'ice breaker':      m.iceBreaker = i;  break;
      case 'last updated':     m.updatedAt = i;   break;
      case 'time':             m.time = i;        break;
    }
  });
  return m;
}

function getRosterEntries(ss) {
  const sheet = (ss || SpreadsheetApp.openById(SPREADSHEET_ID)).getSheetByName(ROSTER_SHEET_NAME);
  if (!sheet || sheet.getLastRow() <= 1) return [];

  const data = sheet.getDataRange().getValues();
  // Auto-detect header row: if row 1 has no recognized column names, it's the notice row
  const headerRowIdx = Object.keys(_rosterColMap(data[0])).length > 0 ? 0 : 1;
  const col  = _rosterColMap(data[headerRowIdx]);
  const tz   = Session.getScriptTimeZone();
  const g    = function(row, key) { return col[key] !== undefined ? row[col[key]] : ''; };
  const entries = [];

  for (let i = headerRowIdx + 1; i < data.length; i++) {
    const row = data[i];
    if (!g(row, 'date')) continue;
    const dateObj = new Date(g(row, 'date'));

    const rawUpdatedAt = g(row, 'updatedAt');
    const rawTime      = g(row, 'time');
    // Use 'UTC' (not script timezone) — Sheets stores time-only values as UTC fractions.
    // Applying the local timezone shifts the time by +10/+11 hours, corrupting the value.
    const timeStr      = rawTime instanceof Date
                         ? Utilities.formatDate(rawTime, 'UTC', 'HH:mm')
                         : String(rawTime || '');
    entries.push({
      rowIndex:    i + 1,
      date:        Utilities.formatDate(dateObj, tz, 'yyyy-MM-dd'),
      group:       String(g(row, 'group')       || ''),
      eventType:   String(g(row, 'eventType')   || ''),
      venue:       String(g(row, 'venue')       || ''),
      organiser:   String(g(row, 'organiser')   || ''),
      pw:          String(g(row, 'pw')          || ''),
      facilitator: String(g(row, 'facilitator') || ''),
      food:        String(g(row, 'food')        || ''),
      reporting:   String(g(row, 'reporting')   || ''),
      notes:       String(g(row, 'notes')       || ''),
      iceBreaker:  String(g(row, 'iceBreaker')  || ''),
      updatedAt:   rawUpdatedAt ? Utilities.formatDate(new Date(rawUpdatedAt), tz, "yyyy-MM-dd'T'HH:mm:ss") : '',
      time:        timeStr
    });
  }

  return entries;
}

function getMembers(ss) {
  const sheet = (ss || SpreadsheetApp.openById(SPREADSHEET_ID)).getSheetByName(MEMBERS_SHEET_NAME);
  if (!sheet || sheet.getLastRow() <= 1) return [];

  const data = sheet.getDataRange().getValues();
  // Auto-detect: if row 1 col A is 'Name', headers are in row 1 — members start at index 1.
  // Otherwise row 1 is the notice row — skip to index 2.
  const memberStartIdx = String(data[0][0]).toLowerCase().trim() === 'name' ? 1 : 2;
  const members = [];

  for (let i = memberStartIdx; i < data.length; i++) {
    const row = data[i];
    if (!row[0]) continue;
    members.push({
      rowIndex:      i + 1,
      name:          String(row[0]),
      group:         String(row[1]),
      canOrganise:   row[2] === true,
      canPW:         row[3] === true,
      canFacilitate: row[4] === true,
      canReport:     row[5] === true,
      active:        row[6] === true,
      roleType:      String(row[7] || 'Adult'),
      canDrive:      row[8] === true
    });
  }

  return members;
}

// ---- Roster CRUD ----

function saveRosterEntry(entry) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName(ROSTER_SHEET_NAME);
    if (!sheet) return { success: false, error: 'Roster sheet not found' };

    const dateParts = entry.date.split('-');
    const dateObj = new Date(
      parseInt(dateParts[0]),
      parseInt(dateParts[1]) - 1,
      parseInt(dateParts[2])
    );

    // Read sheet once: serves both header mapping and row write.
    // Building rowData by column position makes writes column-order agnostic —
    // reordering columns in the sheet never breaks saves.
    const data         = sheet.getDataRange().getValues();
    const headerRowIdx = Object.keys(_rosterColMap(data[0])).length > 0 ? 0 : 1;
    const col          = _rosterColMap(data[headerRowIdx]);
    const numCols      = data[headerRowIdx].length;

    // C: Apply @-format BEFORE writing. Range = current data rows + 50 row buffer (covers
    // appendRow for new entries without formatting the entire 1000-row sheet every save).
    if (col.time !== undefined) {
      const fmtRows = Math.min(sheet.getMaxRows() - 1, sheet.getLastRow() + 50);
      if (fmtRows > 0) sheet.getRange(2, col.time + 1, fmtRows, 1).setNumberFormat('@');
    }

    const rowData = new Array(numCols).fill('');
    const s = function(key, val) { if (col[key] !== undefined) rowData[col[key]] = val; };
    s('date',        dateObj);
    s('group',       entry.group);
    s('eventType',   entry.eventType);
    s('venue',       entry.venue       || '');
    s('organiser',   entry.organiser   || '');
    s('pw',          entry.pw          || '');
    s('facilitator', entry.facilitator || '');
    s('food',        entry.food        || '');
    s('reporting',   entry.reporting   || '');
    s('notes',       entry.notes       || '');
    s('iceBreaker',  entry.iceBreaker  || '');
    s('updatedAt',   new Date());
    s('time',        entry.time        || '');

    if (entry.rowIndex) {
      sheet.getRange(entry.rowIndex, 1, 1, numCols).setValues([rowData]);
    } else {
      sheet.appendRow(rowData);
    }

    sortRosterSheet(sheet);
    return { success: true };
  } catch (e) {
    return { success: false, error: e.toString() };
  }
}

// Batch version: saves all entries in one server round-trip (one sheet open, one read, one sort).
// Always prefer this over calling saveRosterEntry in a loop.
function saveRosterEntries(entries) {
  try {
    const ss    = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName(ROSTER_SHEET_NAME);
    if (!sheet) return { success: false, error: 'Roster sheet not found' };

    const data         = sheet.getDataRange().getValues();
    const headerRowIdx = Object.keys(_rosterColMap(data[0])).length > 0 ? 0 : 1;
    const col          = _rosterColMap(data[headerRowIdx]);
    const numCols      = data[headerRowIdx].length;

    // C: Apply @-format BEFORE any writes. Range = current data rows + 50 row buffer.
    // Must happen before writes: Sheets auto-converts 'HH:mm' strings to fractions if the
    // cell format isn't already @; applying @ after the write corrupts the stored value.
    if (col.time !== undefined) {
      const fmtRows = Math.min(sheet.getMaxRows() - 1, sheet.getLastRow() + 50);
      if (fmtRows > 0) sheet.getRange(2, col.time + 1, fmtRows, 1).setNumberFormat('@');
    }

    entries.forEach(function(entry) {
      const dp      = entry.date.split('-');
      const dateObj = new Date(parseInt(dp[0]), parseInt(dp[1]) - 1, parseInt(dp[2]));
      const rowData = new Array(numCols).fill('');
      const s = function(key, v) { if (col[key] !== undefined) rowData[col[key]] = v; };
      s('date',        dateObj);
      s('group',       entry.group);
      s('eventType',   entry.eventType);
      s('venue',       entry.venue       || '');
      s('organiser',   entry.organiser   || '');
      s('pw',          entry.pw          || '');
      s('facilitator', entry.facilitator || '');
      s('food',        entry.food        || '');
      s('reporting',   entry.reporting   || '');
      s('notes',       entry.notes       || '');
      s('iceBreaker',  entry.iceBreaker  || '');
      s('updatedAt',   new Date());
      s('time',        entry.time        || '');

      if (entry.rowIndex) {
        sheet.getRange(entry.rowIndex, 1, 1, numCols).setValues([rowData]);
      } else {
        sheet.appendRow(rowData);
      }
    });

    sortRosterSheet(sheet);
    return { success: true };
  } catch (e) {
    return { success: false, error: e.toString() };
  }
}

function deleteRosterEntry(rowIndex) {
  try {
    SpreadsheetApp.openById(SPREADSHEET_ID)
      .getSheetByName(ROSTER_SHEET_NAME)
      .deleteRow(rowIndex);
    return { success: true };
  } catch (e) {
    return { success: false, error: e.toString() };
  }
}

// ---- Members CRUD ----

function saveMember(member) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName(MEMBERS_SHEET_NAME);
    if (!sheet) return { success: false, error: 'Members sheet not found' };

    const rowData = [
      member.name,
      member.group,
      member.canOrganise   === true,
      member.canPW         === true,
      member.canFacilitate === true,
      member.canReport     === true,
      member.active        !== false,
      member.roleType      || 'Adult',
      member.canDrive      === true
    ];

    if (member.rowIndex) {
      sheet.getRange(member.rowIndex, 1, 1, 9).setValues([rowData]);
    } else {
      sheet.appendRow(rowData);
    }

    return { success: true };
  } catch (e) {
    return { success: false, error: e.toString() };
  }
}

function deleteMember(rowIndex) {
  try {
    SpreadsheetApp.openById(SPREADSHEET_ID)
      .getSheetByName(MEMBERS_SHEET_NAME)
      .deleteRow(rowIndex);
    return { success: true };
  } catch (e) {
    return { success: false, error: e.toString() };
  }
}

// ---- Helpers ----

function sortRosterSheet(sheet) {
  const lastRow = sheet.getLastRow();
  if (lastRow <= 2) return;
  // Auto-detect data start row: if row 1 col A is 'Date', headers are in row 1 — data starts at 2.
  // Otherwise a notice row exists in row 1 — data starts at 3.
  const row1ColA     = String(sheet.getRange(1, 1).getValue()).toLowerCase().trim();
  const dataStartRow = row1ColA === 'date' ? 2 : 3;
  if (lastRow < dataStartRow) return;
  sheet.getRange(dataStartRow, 1, lastRow - dataStartRow + 1, sheet.getLastColumn()).sort([
    { column: 1, ascending: true },
    { column: 2, ascending: true }
  ]);
}

// ---- Data-Safe Deployment Notes ----
// Deploying a new Apps Script version NEVER modifies sheet data — only the
// code changes.  getRosterEntries() maps columns by header name, so adding
// or reordering columns in the sheet is always safe.
//
// When a schema-breaking change is needed (new column, renamed header):
//   1. Bump VERSION (MINOR bump) and add a migrateSchemaToVXY() function below.
//      The function name MUST match the new version number (e.g. v1.12 → migrateSchemaToV112).
//   2. Run it ONCE from the Apps Script editor (never auto-run on load).
//   3. The migration inserts the new column/header without touching other data.
//   4. Only after migration succeeds should the new code be deployed.

// ---- Schema Migration: v1.19 → v1.20 ----
// Run ONCE to remove the Event ID column from the Roster tab.
// The app no longer generates or reads UUIDs — rowIndex is used for all row lookups.
// After running, call formatSheets() to refresh column widths.
function migrateSchemaToV120() {
  const ss    = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(ROSTER_SHEET_NAME);
  if (!sheet) { Logger.log('Roster sheet not found.'); return; }

  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const lower   = headers.map(function(h) { return String(h).toLowerCase().trim(); });
  const idCol   = lower.indexOf('event id');

  if (idCol < 0) {
    Logger.log('Event ID column not found — already removed or never existed.');
    return;
  }

  sheet.deleteColumn(idCol + 1);
  Logger.log('Removed Event ID column (was col ' + (idCol + 1) + '). Run formatSheets() to refresh formatting.');
}

// ---- Schema Migration: v1.20 → v1.21 ----
// Run ONCE to insert a notice row as row 1 in both Roster and Members sheets.
// Before migration: notice sits to the right of headers in row 1.
// After migration: notice occupies row 1 (frozen); column headers in row 2; data from row 3.
// Run order: migrateSchemaToV120 → migrateSchemaToV121 → fixMembersSheetGhostRows → formatSheets
// Idempotent: skips sheets where row 1 is already the notice row (col A ≠ known header).
// After running, call formatSheets() to write the notice content and finalize formatting.
function migrateSchemaToV121() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);

  function insertNoticeRow(sheet, sheetName, expectedHeader) {
    const row1ColA = String(sheet.getRange(1, 1).getValue()).toLowerCase().trim();
    if (row1ColA !== expectedHeader) {
      Logger.log(sheetName + ': already migrated (row 1 col A = "' + row1ColA + '") — skipping.');
      return;
    }
    sheet.insertRowBefore(1);
    Logger.log(sheetName + ': inserted blank notice row 1. Headers now in row 2. Run formatSheets() to write notice content.');
  }

  const rosterSheet  = ss.getSheetByName(ROSTER_SHEET_NAME);
  const membersSheet = ss.getSheetByName(MEMBERS_SHEET_NAME);
  if (rosterSheet)  insertNoticeRow(rosterSheet,  'Roster',  'date');
  if (membersSheet) insertNoticeRow(membersSheet, 'Members', 'name');

  Logger.log('migrateSchemaToV121 complete. Run formatSheets() to finalize.');
}

// ---- Utility: Fix Members Sheet Ghost Rows ----
// Run ONCE if Older Sunday School / Harvest members landed at rows 1001+.
// Root cause: formatSheets() applies checkbox validation to all rows in the sheet,
// making getLastRow() return ~1000, so appendRow() lands rows after that.
// This function deletes blank rows between real member rows (bottom-to-top to avoid
// index shifting), then logs how many were removed.
// After running, call formatSheets() to re-apply formatting.
function fixMembersSheetGhostRows() {
  const ss    = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(MEMBERS_SHEET_NAME);
  if (!sheet) { Logger.log('Members sheet not found.'); return; }

  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) { Logger.log('Members sheet is empty.'); return; }

  // Read only col A (Name) — sufficient to identify blank rows
  const names = sheet.getRange(2, 1, lastRow - 1, 1).getValues(); // rows 2..lastRow

  // Collect blank-row runs (1-indexed), traversing bottom-to-top
  const blanks = [];
  let runEnd = -1;
  for (let i = names.length - 1; i >= 0; i--) {
    const isEmpty = !names[i][0] || String(names[i][0]).trim() === '';
    if (isEmpty) {
      if (runEnd === -1) runEnd = i + 2; // convert to 1-indexed sheet row
    } else {
      if (runEnd !== -1) {
        blanks.push({ start: i + 3, end: runEnd }); // blank run starts one row below current real row
        runEnd = -1;
      }
    }
  }
  if (runEnd !== -1) blanks.push({ start: 2, end: runEnd }); // blanks immediately below header

  if (blanks.length === 0) {
    Logger.log('Members sheet is already compact — no ghost rows found.');
    return;
  }

  // Delete from highest rows first (blanks array is already in that order) to avoid index shifting
  let totalDeleted = 0;
  blanks.forEach(function(b) {
    const count = b.end - b.start + 1;
    sheet.deleteRows(b.start, count);
    totalDeleted += count;
    Logger.log('Deleted rows ' + b.start + '–' + b.end + ' (' + count + ' ghost rows)');
  });

  Logger.log('Done: removed ' + totalDeleted + ' ghost rows. Run formatSheets() to re-apply formatting.');
}

// ---- Sheet Formatting ----
// Run formatSheets() from the Apps Script editor any time to:
//   • Apply column widths, dropdowns, date formats, alternating row colours
//   • Re-run safely after adding new columns — fully idempotent
// This function ONLY changes formatting — it never reads or writes sheet data.
//
// How to add formatting for a new Roster field:
//   1. Add its JS key → pixel width to the `widths` map in _formatRosterSheet()
//   2. If it needs a dropdown, add a validation block (copy the Group pattern)
//   3. If it's system-managed, add a header note (copy the id/updatedAt pattern)
//   4. Run formatSheets() from the editor — done.

function formatSheets() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  _formatRosterSheet(ss);
  _formatMembersSheet(ss);
  Logger.log('formatSheets complete.');
}

function _formatRosterSheet(ss) {
  const sheet = ss.getSheetByName(ROSTER_SHEET_NAME);
  if (!sheet) { Logger.log('Roster sheet not found.'); return; }

  // Auto-detect layout: if row 1 has no recognized column headers, it's the notice row (post-migration)
  const row1vals     = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const hasNoticeRow = Object.keys(_rosterColMap(row1vals)).length === 0;
  const headerRow    = hasNoticeRow ? 2 : 1;   // 1-indexed sheet row containing column headers
  const dataStartRow = hasNoticeRow ? 3 : 2;   // 1-indexed first data row

  const headers  = hasNoticeRow
    ? sheet.getRange(2, 1, 1, sheet.getLastColumn()).getValues()[0]
    : row1vals;
  const col      = _rosterColMap(headers);
  const maxRows  = sheet.getMaxRows();
  const dataRows = maxRows - dataStartRow + 1;

  // Use recognized columns only for count — immune to notice cell pollution in getLastColumn()
  const dataColCount = Object.keys(col).length > 0 ? Math.max(...Object.values(col)) + 1 : 0;

  // --- Column widths (add new fields here) ---
  const widths = {
    date: 120, group: 70, eventType: 130, venue: 160,
    organiser: 130, pw: 130, facilitator: 130, food: 120,
    reporting: 130, notes: 220, iceBreaker: 160,
    time: 80, updatedAt: 145
  };
  Object.entries(widths).forEach(function([key, w]) {
    if (col[key] !== undefined) sheet.setColumnWidth(col[key] + 1, w);
  });

  // --- Freeze: notice row + header row if post-migration, otherwise just header row ---
  sheet.setFrozenRows(hasNoticeRow ? 2 : 1);

  // --- Last Updated: header note ---
  if (col.updatedAt !== undefined) {
    sheet.getRange(headerRow, col.updatedAt + 1).setNote('Auto-stamped by the app on every save. Do not edit manually.');
  }

  // --- Time: header note ---
  if (col.time !== undefined) {
    sheet.getRange(headerRow, col.time + 1).setNote('24-hour format, e.g. 18:30 for 6:30 PM. Leave blank if no fixed time.');
  }

  // --- Group: dropdown validation ---
  if (col.group !== undefined) {
    const v = SpreadsheetApp.newDataValidation()
      .requireValueInList(['JAG1', 'JAG2'], true).setAllowInvalid(false).build();
    sheet.getRange(dataStartRow, col.group + 1, dataRows, 1).setDataValidation(v);
  }

  // --- Event Type: dropdown validation ---
  if (col.eventType !== undefined) {
    const v = SpreadsheetApp.newDataValidation()
      .requireValueInList(['Youth Hour', 'Separated LG', 'Combined', 'Special', 'Cancelled', 'Replaced'], true)
      .setAllowInvalid(false).build();
    sheet.getRange(dataStartRow, col.eventType + 1, dataRows, 1).setDataValidation(v);
  }

  // --- Alternating row colours ---
  // clearFormat() removes ALL explicit cell formatting (backgrounds, number formats, fonts) so
  // banding applies uniformly across every column. Number formats are re-applied after banding.
  sheet.getBandings().forEach(function(b) { b.remove(); });
  if (dataColCount > 0) {
    sheet.getRange(dataStartRow, 1, dataRows, dataColCount).clearFormat();
    sheet.getRange(dataStartRow, 1, dataRows, dataColCount)
      .applyRowBanding()
      .setFirstRowColor('#f5f3ff')
      .setSecondRowColor('#ffffff');
    // Re-apply number formats after clearFormat() wiped them
    if (col.date !== undefined)
      sheet.getRange(dataStartRow, col.date + 1, dataRows, 1).setNumberFormat('ddd dd/mm/yyyy');
    if (col.updatedAt !== undefined)
      sheet.getRange(dataStartRow, col.updatedAt + 1, dataRows, 1).setNumberFormat('dd/mm/yyyy hh:mm');
    if (col.time !== undefined && dataRows > 0)
      sheet.getRange(dataStartRow, col.time + 1, dataRows, 1).setNumberFormat('@');
  }

  // --- Portal notice ---
  sheet.getRange(1, 1, 1, sheet.getMaxColumns()).breakApart();
  if (hasNoticeRow) {
    // Post-migration: notice spans data columns in row 1 (idempotent — safe to re-run)
    if (dataColCount > 0) {
      sheet.getRange(1, 1, 1, dataColCount).merge()
        .setValue('⚠️  Please use the JAG Roster Portal to make changes — do not edit this sheet directly.\n🔗  https://tinyurl.com/jagrosterportal')
        .setBackground('#fef08a')
        .setFontColor('#713f12')
        .setFontWeight('bold')
        .setFontSize(9)
        .setWrap(true)
        .setVerticalAlignment('middle')
        .setHorizontalAlignment('center');
      sheet.setRowHeight(1, 48);
    }
  } else {
    // Pre-migration: notice to the right of data; clear stale columns first (growing-column fix)
    const maxCols = sheet.getMaxColumns();
    if (dataColCount > 0 && maxCols > dataColCount + 1) {
      sheet.getRange(1, dataColCount + 2, 1, maxCols - dataColCount - 1).clear();
    }
    const rNoticeCol = dataColCount + 2;
    sheet.getRange(1, rNoticeCol, 1, 4).merge()
      .setValue('⚠️  Please use the JAG Roster Portal to make changes — do not edit this sheet directly.\n🔗  https://tinyurl.com/jagrosterportal')
      .setBackground('#fef08a')
      .setFontColor('#713f12')
      .setFontWeight('bold')
      .setFontSize(9)
      .setWrap(true)
      .setVerticalAlignment('middle')
      .setHorizontalAlignment('center');
    sheet.setColumnWidth(rNoticeCol, 320);
    sheet.setRowHeight(1, 48);
  }

  Logger.log('Roster sheet formatted (' + dataColCount + ' data columns, ' + (hasNoticeRow ? 'post' : 'pre') + '-migration).');
}

function _formatMembersSheet(ss) {
  const sheet = ss.getSheetByName(MEMBERS_SHEET_NAME);
  if (!sheet) { Logger.log('Members sheet not found.'); return; }

  // Auto-detect layout: if row 1 col A is 'Name', headers are in row 1 (pre-migration).
  // Otherwise row 1 is the notice row (post-migration).
  const row1ColA     = String(sheet.getRange(1, 1).getValue()).toLowerCase().trim();
  const hasNoticeRow = row1ColA !== 'name';
  const headerRow    = hasNoticeRow ? 2 : 1;
  const dataStartRow = hasNoticeRow ? 3 : 2;

  const maxRows        = sheet.getMaxRows();
  const dataRows       = maxRows - dataStartRow + 1;
  const DATA_COL_COUNT = 9; // Members schema is always 9 columns

  // --- Column widths (positional, matches Members schema order) ---
  [160, 70, 105, 80, 110, 90, 70, 90, 80].forEach(function(w, i) {
    sheet.setColumnWidth(i + 1, w);
  });

  // --- Freeze: notice row + header row if post-migration, otherwise just header row ---
  sheet.setFrozenRows(hasNoticeRow ? 2 : 1);

  // --- Read headers from the correct row for dropdown/checkbox column detection ---
  const headers = sheet.getRange(headerRow, 1, 1, DATA_COL_COUNT).getValues()[0];
  const lower   = headers.map(function(h) { return String(h).toLowerCase().trim(); });

  // --- Group: dropdown ---
  const groupIdx = lower.indexOf('group');
  if (groupIdx >= 0) {
    const v = SpreadsheetApp.newDataValidation()
      .requireValueInList(['JAG1', 'JAG2', 'Both'], true).setAllowInvalid(false).build();
    sheet.getRange(dataStartRow, groupIdx + 1, dataRows, 1).setDataValidation(v);
  }

  // --- Role Type: dropdown ---
  const roleIdx = lower.indexOf('role type');
  if (roleIdx >= 0) {
    const v = SpreadsheetApp.newDataValidation()
      .requireValueInList(['Adult', 'Student', 'Older Sunday School', 'Harvest'], true).setAllowInvalid(false).build();
    sheet.getRange(dataStartRow, roleIdx + 1, dataRows, 1).setDataValidation(v);
  }

  // --- Boolean columns: checkbox ---
  ['can organise', 'can p&w', 'can facilitate', 'can report', 'active', 'can drive'].forEach(function(name) {
    const idx = lower.indexOf(name);
    if (idx >= 0) {
      sheet.getRange(dataStartRow, idx + 1, dataRows, 1)
        .setDataValidation(SpreadsheetApp.newDataValidation().requireCheckbox().build());
    }
  });

  // --- Alternating row colours ---
  sheet.getBandings().forEach(function(b) { b.remove(); });
  sheet.getRange(dataStartRow, 1, dataRows, DATA_COL_COUNT).clearFormat();
  sheet.getRange(dataStartRow, 1, dataRows, DATA_COL_COUNT)
    .applyRowBanding()
    .setFirstRowColor('#f5f3ff')
    .setSecondRowColor('#ffffff');

  // --- Portal notice ---
  sheet.getRange(1, 1, 1, sheet.getMaxColumns()).breakApart();
  if (hasNoticeRow) {
    // Post-migration: notice spans data columns in row 1 (idempotent — safe to re-run)
    sheet.getRange(1, 1, 1, DATA_COL_COUNT).merge()
      .setValue('⚠️  Please use the JAG Roster Portal to make changes — do not edit this sheet directly.\n🔗  https://tinyurl.com/jagrosterportal')
      .setBackground('#fef08a')
      .setFontColor('#713f12')
      .setFontWeight('bold')
      .setFontSize(9)
      .setWrap(true)
      .setVerticalAlignment('middle')
      .setHorizontalAlignment('center');
    sheet.setRowHeight(1, 48);
  } else {
    // Pre-migration: notice to the right of data; clear stale columns first (growing-column fix)
    const mMaxCols = sheet.getMaxColumns();
    if (mMaxCols > DATA_COL_COUNT + 1) {
      sheet.getRange(1, DATA_COL_COUNT + 2, 1, mMaxCols - DATA_COL_COUNT - 1).clear();
    }
    const mNoticeCol = DATA_COL_COUNT + 2;
    sheet.getRange(1, mNoticeCol, 1, 4).merge()
      .setValue('⚠️  Please use the JAG Roster Portal to make changes — do not edit this sheet directly.\n🔗  https://tinyurl.com/jagrosterportal')
      .setBackground('#fef08a')
      .setFontColor('#713f12')
      .setFontWeight('bold')
      .setFontSize(9)
      .setWrap(true)
      .setVerticalAlignment('middle')
      .setHorizontalAlignment('center');
    sheet.setColumnWidth(mNoticeCol, 320);
    sheet.setRowHeight(1, 48);
  }

  Logger.log('Members sheet formatted (' + DATA_COL_COUNT + ' data columns, ' + (hasNoticeRow ? 'post' : 'pre') + '-migration).');
}


