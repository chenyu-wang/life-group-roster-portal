// ============================================================
// JAG Life Group Roster - Google Apps Script Backend
// Spreadsheet: https://docs.google.com/spreadsheets/d/1Cg9m7lUu536JlSXbY4HifWQpOw9nQ2DtBRDZRzIXIn4
// Version: 1.28.0 (2026-04-12)
// ============================================================

const VERSION      = '1.28.0';
const VERSION_DATE = '2026-04-12';

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
  const firstRowMap  = _rosterColMap(data[0]);
  const headerRowIdx = Object.keys(firstRowMap).length > 0 ? 0 : 1;
  const col          = headerRowIdx === 0 ? firstRowMap : _rosterColMap(data[headerRowIdx]);
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
    const firstRowMap  = _rosterColMap(data[0]);
    const headerRowIdx = Object.keys(firstRowMap).length > 0 ? 0 : 1;
    const col          = headerRowIdx === 0 ? firstRowMap : _rosterColMap(data[headerRowIdx]);
    const numCols      = data[headerRowIdx].length;

    // Sort only when the row order can change: new row or date edited.
    const needsSort = !entry.rowIndex || (function() {
      const oldCell = data[entry.rowIndex - 1] && data[entry.rowIndex - 1][col.date];
      if (!oldCell) return true;
      const old = new Date(oldCell);
      return old.getFullYear() !== dateObj.getFullYear() ||
             old.getMonth()    !== dateObj.getMonth()    ||
             old.getDate()     !== dateObj.getDate();
    })();

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

    SpreadsheetApp.flush(); // commit writes before sort so getLastColumn() sees col M
    if (needsSort) {
      sortRosterSheet(sheet);
      return { success: true };
    }
    return { success: true, stable: true }; // rowIndices unchanged — client can skip loadData()
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
    const firstRowMap  = _rosterColMap(data[0]);
    const headerRowIdx = Object.keys(firstRowMap).length > 0 ? 0 : 1;
    const col          = headerRowIdx === 0 ? firstRowMap : _rosterColMap(data[headerRowIdx]);
    const numCols      = data[headerRowIdx].length;

    // Sort only when the row order can change: any new row or any entry with a changed date.
    const needsSort = entries.some(function(entry) {
      if (!entry.rowIndex) return true;
      const dp      = entry.date.split('-');
      const newDate = new Date(parseInt(dp[0]), parseInt(dp[1]) - 1, parseInt(dp[2]));
      const oldCell = data[entry.rowIndex - 1] && data[entry.rowIndex - 1][col.date];
      if (!oldCell) return true;
      const old = new Date(oldCell);
      return old.getFullYear() !== newDate.getFullYear() ||
             old.getMonth()    !== newDate.getMonth()    ||
             old.getDate()     !== newDate.getDate();
    });

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

    SpreadsheetApp.flush(); // commit writes before sort so getLastColumn() sees col M
    if (needsSort) {
      sortRosterSheet(sheet);
      return { success: true };
    }
    return { success: true, stable: true }; // rowIndices unchanged — client can skip loadData()
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

// ---- Migration: Roster Group="Both" ----
// Run ONCE after deploying v1.26.0.
// Merges legacy JAG1+JAG2 row pairs for combined events into a single row with Group="Both".
// Also cleans up ghost duplicate rows. Idempotent — safe to re-run.
// Delete this function after confirmed run on the live sheet.
function migrateRosterToGroupBoth() {
  const ss    = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(ROSTER_SHEET_NAME);
  if (!sheet) { Logger.log('Roster sheet not found.'); return; }

  const data         = sheet.getDataRange().getValues();
  const firstRowMap  = _rosterColMap(data[0]);
  const headerRowIdx = Object.keys(firstRowMap).length > 0 ? 0 : 1;
  const col          = headerRowIdx === 0 ? firstRowMap : _rosterColMap(data[headerRowIdx]);
  const dataStart    = headerRowIdx + 2; // 1-indexed first data row (sheet row)

  // Update Group column validation to accept "Both" before writing — the old rule
  // only allows JAG1/JAG2 and would reject the setValue("Both") calls below.
  if (col.group !== undefined) {
    const dataRows = sheet.getMaxRows() - dataStart + 1;
    if (dataRows > 0) {
      const v = SpreadsheetApp.newDataValidation()
        .requireValueInList(['JAG1', 'JAG2', 'Both'], true).setAllowInvalid(false).build();
      sheet.getRange(dataStart, col.group + 1, dataRows, 1).setDataValidation(v);
    }
    SpreadsheetApp.flush();
  }

  // Index all data rows by date+eventType key
  const byKey = {}; // key → [{ rowIndex, group, row }]
  for (let i = dataStart - 1; i < data.length; i++) {
    const row  = data[i];
    const date = row[col.date];
    if (!date) continue;
    const key = String(date) + '|' + String(row[col.eventType] || '');
    if (!byKey[key]) byKey[key] = [];
    byKey[key].push({ rowIndex: i + 1, group: String(row[col.group] || ''), row });
  }

  const toUpdate  = []; // { rowIndex, colIndex (1-indexed), value }
  const toDelete  = []; // rowIndex values (1-indexed), deleted bottom-up

  Object.values(byKey).forEach(function(rows) {
    const et = String((rows[0].row[col.eventType] || '')).trim();
    if (et === 'Separated LG') return; // per-group rows are correct — leave untouched

    const bothRows = rows.filter(r => r.group === 'Both');
    const jag1Rows = rows.filter(r => r.group === 'JAG1');
    const jag2Rows = rows.filter(r => r.group === 'JAG2');

    if (bothRows.length > 0) {
      // Already migrated: keep first Both row, delete all duplicates and legacy JAG1/JAG2
      bothRows.slice(1).forEach(r => toDelete.push(r.rowIndex));
      jag1Rows.forEach(r => toDelete.push(r.rowIndex));
      jag2Rows.forEach(r => toDelete.push(r.rowIndex));
      return;
    }

    // Not yet migrated: pick the JAG1 row (or JAG2 if none) as the keeper
    const keepRow = jag1Rows[0] || jag2Rows[0];
    if (!keepRow) return;

    // Change the keeper's Group cell to "Both"
    toUpdate.push({ rowIndex: keepRow.rowIndex, colIndex: col.group + 1, value: 'Both' });

    // Delete all other JAG1 rows (duplicates) and all JAG2 rows
    jag1Rows.forEach(r => { if (r !== keepRow) toDelete.push(r.rowIndex); });
    jag2Rows.forEach(r => toDelete.push(r.rowIndex));
  });

  // Apply Group cell updates first
  toUpdate.forEach(u => sheet.getRange(u.rowIndex, u.colIndex).setValue(u.value));
  SpreadsheetApp.flush();

  // Delete rows bottom-up so row indices don't shift during deletion
  const sorted = [...new Set(toDelete)].sort((a, b) => b - a);
  sorted.forEach(rowIndex => sheet.deleteRow(rowIndex));
  SpreadsheetApp.flush();

  Logger.log('migrateRosterToGroupBoth: updated ' + toUpdate.length + ' rows to Group="Both", deleted ' + sorted.length + ' rows.');
}

// ---- Utility: Rebuild Last Updated Column ----
// Run ONCE to reset col M (Last Updated) to a clean datetime column.
// Clears any stale content/format from data rows, re-writes the header, and applies
// the correct datetime format. Run formatSheets() afterwards to restore banding.
function rebuildLastUpdatedColumn() {
  const ss    = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(ROSTER_SHEET_NAME);
  if (!sheet) { Logger.log('Roster sheet not found.'); return; }

  const headerRow    = 2;   // post-migration: headers in row 2
  const dataStartRow = 3;
  const targetCol    = 13;  // col M (1-indexed)
  const lastRow      = sheet.getLastRow();

  // Ensure header says 'Last Updated'
  sheet.getRange(headerRow, targetCol).setValue('Last Updated');

  // Clear all stale content + formatting from data rows, then apply datetime format
  if (lastRow >= dataStartRow) {
    const dataRange = sheet.getRange(dataStartRow, targetCol, lastRow - dataStartRow + 1, 1);
    dataRange.clearContent();
    dataRange.clearFormat();
    dataRange.setNumberFormat('dd/mm/yyyy hh:mm');
  }

  // Match column width from _formatRosterSheet
  sheet.setColumnWidth(targetCol, 145);

  SpreadsheetApp.flush();
  Logger.log('Last Updated column (col M) rebuilt. Run formatSheets() to restore banding.');
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
      .requireValueInList(['JAG1', 'JAG2', 'Both'], true).setAllowInvalid(false).build();
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
  // clearFormat() removes ALL explicit cell formatting so banding applies uniformly.
  // Number formats are re-applied after banding, then setBackground(null) ensures
  // setNumberFormat() hasn't left any implicit explicit backgrounds that override banding.
  sheet.getBandings().forEach(function(b) { b.remove(); });
  if (dataColCount > 0) {
    const dataRange = sheet.getRange(dataStartRow, 1, dataRows, dataColCount);
    dataRange.clearFormat();
    dataRange.applyRowBanding().setFirstRowColor('#f5f3ff').setSecondRowColor('#ffffff');
    // Re-apply number formats after clearFormat() wiped them
    if (col.date !== undefined)
      sheet.getRange(dataStartRow, col.date + 1, dataRows, 1).setNumberFormat('ddd dd/mm/yyyy');
    if (col.updatedAt !== undefined)
      sheet.getRange(dataStartRow, col.updatedAt + 1, dataRows, 1).setNumberFormat('dd/mm/yyyy hh:mm');
    if (col.time !== undefined && dataRows > 0)
      sheet.getRange(dataStartRow, col.time + 1, dataRows, 1).setNumberFormat('@');
    // Final pass: setBackground(null) clears any implicit backgrounds set by setNumberFormat,
    // ensuring banding colours show through uniformly on all columns including Last Updated.
    dataRange.setBackground(null);
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
      .requireValueInList(['JAG1', 'JAG2', 'Both', 'Sunday School'], true).setAllowInvalid(false).build();
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


