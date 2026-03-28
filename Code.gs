// ============================================================
// JAG Life Group Roster - Google Apps Script Backend
// Spreadsheet: https://docs.google.com/spreadsheets/d/1Cg9m7lUu536JlSXbY4HifWQpOw9nQ2DtBRDZRzIXIn4
// Version: 1.17.3 (2026-03-28)
// ============================================================

const VERSION      = '1.17.3';
const VERSION_DATE = '2026-03-28';

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
  return {
    entries: getRosterEntries(),
    members: getMembers()
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
      case 'event id':         m.id   = i;        break;
    }
  });
  return m;
}

function getRosterEntries() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(ROSTER_SHEET_NAME);
  if (!sheet || sheet.getLastRow() <= 1) return [];

  const data = sheet.getDataRange().getValues();
  const col  = _rosterColMap(data[0]);   // derive positions from header row
  const tz   = Session.getScriptTimeZone();
  const g    = function(row, key) { return col[key] !== undefined ? row[col[key]] : ''; };
  const entries = [];

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (!g(row, 'date')) continue;
    const dateObj = new Date(g(row, 'date'));

    const rawUpdatedAt = g(row, 'updatedAt');
    const rawTime      = g(row, 'time');
    const timeStr      = rawTime instanceof Date
                         ? Utilities.formatDate(rawTime, tz, 'HH:mm')
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
      time:        timeStr,
      id:          String(g(row, 'id')          || '')
    });
  }

  return entries;
}

function getMembers() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(MEMBERS_SHEET_NAME);
  if (!sheet || sheet.getLastRow() <= 1) return [];

  const data = sheet.getDataRange().getValues();
  const members = [];

  for (let i = 1; i < data.length; i++) {
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

    const entryId = entry.id || Utilities.getUuid();

    // Read sheet once: serves both header mapping and ID lookup.
    // Building rowData by column position makes writes column-order agnostic —
    // reordering columns in the sheet never breaks saves.
    const data    = sheet.getDataRange().getValues();
    const col     = _rosterColMap(data[0]);
    const numCols = data[0].length;
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
    s('id',          entryId);

    // Prefer ID-based lookup for reliable editing
    if (entry.id && col.id !== undefined) {
      for (let i = 1; i < data.length; i++) {
        if (String(data[i][col.id]) === String(entry.id)) {
          sheet.getRange(i + 1, 1, 1, numCols).setValues([rowData]);
          sortRosterSheet(sheet);
          return { success: true };
        }
      }
    }

    // Fall back to rowIndex (for rows without an ID yet) or append new
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
  sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).sort([
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

// ---- Schema Migration: v1.12 → v1.13 ----
// Run ONCE to reorder Roster columns for human readability.
// Moves Time to col D (after Event Type) and Ice Breaker to col I (after Facilitator).
// New order: Date · Group · Event Type · Time · Venue · Organiser · P&W · Facilitator ·
//            Ice Breaker · Food · Reporting · Notes · Last Updated · Event ID
// After running, call formatSheets() to refresh column widths and formatting.
function migrateSchemaToV113() {
  const ss    = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(ROSTER_SHEET_NAME);
  if (!sheet) { Logger.log('Roster sheet not found.'); return; }

  function getHeaderMap() {
    return sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0]
      .map(function(h) { return String(h).toLowerCase().trim(); });
  }

  let lower = getHeaderMap();

  // Idempotency check: Time at col D (idx 3), Ice Breaker at col I (idx 8)
  if (lower.indexOf('time') === 3 && lower.indexOf('ice breaker') === 8) {
    Logger.log('Already on v1.13 schema — nothing to do.'); return;
  }

  // Step 1: Move Time to col D (1-based position 4)
  const timeCol = lower.indexOf('time') + 1;
  if (timeCol > 0 && timeCol !== 4) {
    sheet.moveColumns(sheet.getRange(1, timeCol, 1, 1), 4);
    Logger.log('Moved Time → col D.');
    lower = getHeaderMap();
  }

  // Step 2: Move Ice Breaker to col I (1-based position 9)
  const ibCol = lower.indexOf('ice breaker') + 1;
  if (ibCol > 0 && ibCol !== 9) {
    sheet.moveColumns(sheet.getRange(1, ibCol, 1, 1), 9);
    Logger.log('Moved Ice Breaker → col I.');
  }

  Logger.log('migrateSchemaToV113 complete. Run formatSheets() to refresh column formatting.');
}

// ---- Schema Migration: v1.15 → v1.16 ----
// Run ONCE to add Can Drive column (col I) to the Members tab.
// After running, call formatSheets() to apply checkbox formatting.
function migrateSchemaToV116() {
  const ss    = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(MEMBERS_SHEET_NAME);
  if (!sheet) { Logger.log('Members sheet not found.'); return; }

  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const lower   = headers.map(function(h) { return String(h).toLowerCase().trim(); });

  if (lower.indexOf('can drive') >= 0) {
    Logger.log('Can Drive column already exists — nothing to do.');
    return;
  }

  const newCol = sheet.getLastColumn() + 1;
  sheet.getRange(1, newCol).setValue('Can Drive');
  Logger.log('Added Can Drive column at position ' + newCol + '. Run formatSheets() to apply checkbox formatting.');
}

// ---- One-time import: Older Sunday School members ----
// Run ONCE from the Apps Script editor, then DELETE this function.
function importOlderSSMembers() {
  const ss    = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(MEMBERS_SHEET_NAME);
  if (!sheet) { Logger.log('Members sheet not found.'); return; }

  const names = [
    'Anne-Grace', 'Averley', 'Chelsea', 'Chloe', 'Emma', 'Ethan', 'Faris',
    'Grace', 'James Wijaya', 'Josiah', 'Judith', 'Kaitlyn', 'Lukas',
    'Magdalene', 'Oliver', 'Tiana', 'Timothy'
  ];

  const existing = sheet.getDataRange().getValues().slice(1)
    .map(function(r) { return String(r[0]).toLowerCase().trim(); });

  let added = 0;
  names.forEach(function(name) {
    if (!existing.includes(name.toLowerCase().trim())) {
      // [Name, Group, Organise, P&W, Facilitate, Report, Active, RoleType, Drive]
      sheet.appendRow([name, 'Both', false, false, false, false, true, 'Older Sunday School', false]);
      added++;
    } else {
      Logger.log('Skipped (already exists): ' + name);
    }
  });
  Logger.log('importOlderSSMembers: added ' + added + ' member(s). Delete this function after confirming.');
}

// ---- One-time import: Harvest members (JAG1) ----
// Run ONCE from the Apps Script editor, then DELETE this function.
function importHarvestMembers() {
  const ss    = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(MEMBERS_SHEET_NAME);
  if (!sheet) { Logger.log('Members sheet not found.'); return; }

  const names = ['James Han', 'Cherry', 'Marko', 'Samyar', 'Samira', 'Ava'];

  const existing = sheet.getDataRange().getValues().slice(1)
    .map(function(r) { return String(r[0]).toLowerCase().trim(); });

  let added = 0;
  names.forEach(function(name) {
    if (!existing.includes(name.toLowerCase().trim())) {
      // [Name, Group, Organise, P&W, Facilitate, Report, Active, RoleType, Drive]
      sheet.appendRow([name, 'JAG1', false, false, false, false, true, 'Harvest', false]);
      added++;
    } else {
      Logger.log('Skipped (already exists): ' + name);
    }
  });
  Logger.log('importHarvestMembers: added ' + added + ' member(s). Delete this function after confirming.');
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

  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const col     = _rosterColMap(headers);   // header-based: safe after column reorder
  const maxRows  = sheet.getMaxRows();
  const dataRows = maxRows - 1;

  // --- Column widths (add new fields here) ---
  const widths = {
    date: 120, group: 70, eventType: 130, venue: 160,
    organiser: 130, pw: 130, facilitator: 130, food: 120,
    reporting: 130, notes: 220, iceBreaker: 160,
    time: 80, updatedAt: 145, id: 240
  };
  Object.entries(widths).forEach(function([key, w]) {
    if (col[key] !== undefined) sheet.setColumnWidth(col[key] + 1, w);
  });

  // --- Freeze header ---
  sheet.setFrozenRows(1);

  // --- Date: readable format ---
  if (col.date !== undefined) {
    sheet.getRange(2, col.date + 1, dataRows, 1).setNumberFormat('ddd dd/mm/yyyy');
  }

  // --- Last Updated: datetime format + note ---
  if (col.updatedAt !== undefined) {
    sheet.getRange(2, col.updatedAt + 1, dataRows, 1).setNumberFormat('dd/mm/yyyy hh:mm');
    sheet.getRange(1, col.updatedAt + 1).setNote('Auto-stamped by the app on every save. Do not edit manually.');
  }

  // --- Event ID: de-emphasised colour + note ---
  if (col.id !== undefined) {
    sheet.getRange(2, col.id + 1, dataRows, 1).setFontColor('#94a3b8');
    sheet.getRange(1, col.id + 1).setNote('UUID auto-generated by the app. Do not edit — used for reliable row lookup.');
  }

  // --- Time: header note ---
  if (col.time !== undefined) {
    sheet.getRange(1, col.time + 1).setNote('24-hour format, e.g. 18:30 for 6:30 PM. Leave blank if no fixed time.');
  }

  // --- Group: dropdown validation ---
  if (col.group !== undefined) {
    const v = SpreadsheetApp.newDataValidation()
      .requireValueInList(['JAG1', 'JAG2'], true).setAllowInvalid(false).build();
    sheet.getRange(2, col.group + 1, dataRows, 1).setDataValidation(v);
  }

  // --- Event Type: dropdown validation ---
  if (col.eventType !== undefined) {
    const v = SpreadsheetApp.newDataValidation()
      .requireValueInList(['Youth Hour', 'Separated LG', 'Combined', 'Special', 'Cancelled', 'Replaced'], true)
      .setAllowInvalid(false).build();
    sheet.getRange(2, col.eventType + 1, dataRows, 1).setDataValidation(v);
  }

  // --- Alternating row colours ---
  sheet.getBandings().forEach(function(b) { b.remove(); });
  sheet.getRange(2, 1, dataRows, headers.length)
    .applyRowBanding()
    .setFirstRowColor('#f5f3ff')
    .setSecondRowColor('#ffffff');

  // --- Portal notice (right of data, always visible in header row) ---
  sheet.getRange(1, 1, 1, sheet.getMaxColumns()).breakApart();
  const rNoticeCol = headers.length + 2;
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

  Logger.log('Roster sheet formatted (' + headers.length + ' columns).');
}

function _formatMembersSheet(ss) {
  const sheet = ss.getSheetByName(MEMBERS_SHEET_NAME);
  if (!sheet) { Logger.log('Members sheet not found.'); return; }

  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const lower   = headers.map(function(h) { return String(h).toLowerCase().trim(); });
  const maxRows  = sheet.getMaxRows();
  const dataRows = maxRows - 1;

  // --- Column widths (positional, matches Members schema order) ---
  [160, 70, 105, 80, 110, 90, 70, 90, 80].forEach(function(w, i) {
    if (i < headers.length) sheet.setColumnWidth(i + 1, w);
  });

  // --- Freeze header ---
  sheet.setFrozenRows(1);

  // --- Group: dropdown ---
  const groupIdx = lower.indexOf('group');
  if (groupIdx >= 0) {
    const v = SpreadsheetApp.newDataValidation()
      .requireValueInList(['JAG1', 'JAG2', 'Both'], true).setAllowInvalid(false).build();
    sheet.getRange(2, groupIdx + 1, dataRows, 1).setDataValidation(v);
  }

  // --- Role Type: dropdown ---
  const roleIdx = lower.indexOf('role type');
  if (roleIdx >= 0) {
    const v = SpreadsheetApp.newDataValidation()
      .requireValueInList(['Adult', 'Student', 'Older Sunday School', 'Harvest'], true).setAllowInvalid(false).build();
    sheet.getRange(2, roleIdx + 1, dataRows, 1).setDataValidation(v);
  }

  // --- Boolean columns: checkbox ---
  ['can organise', 'can p&w', 'can facilitate', 'can report', 'active', 'can drive'].forEach(function(name) {
    const idx = lower.indexOf(name);
    if (idx >= 0) {
      sheet.getRange(2, idx + 1, dataRows, 1)
        .setDataValidation(SpreadsheetApp.newDataValidation().requireCheckbox().build());
    }
  });

  // --- Alternating row colours ---
  sheet.getBandings().forEach(function(b) { b.remove(); });
  sheet.getRange(2, 1, dataRows, headers.length)
    .applyRowBanding()
    .setFirstRowColor('#f5f3ff')
    .setSecondRowColor('#ffffff');

  // --- Portal notice (right of data, always visible in header row) ---
  // Break apart entire header row first so any stale merge is fully covered regardless of column count changes.
  sheet.getRange(1, 1, 1, sheet.getMaxColumns()).breakApart();
  const mNoticeCol = headers.length + 2;
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

  Logger.log('Members sheet formatted (' + headers.length + ' columns).');
}


