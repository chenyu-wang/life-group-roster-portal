# JAG Life Group Roster Portal — Development Guide

## Project Overview
Google Apps Script web app for JAG Life Group scheduling at Aflame Church.
- **Code.gs** — server-side backend (Apps Script)
- **Index.html** — single-file SPA frontend (all CSS + JS inline)
- **Google Sheet** (source of truth): `1Cg9m7lUu536JlSXbY4HifWQpOw9nQ2DtBRDZRzIXIn4`

---

## MANDATORY: Check and Update CLAUDE.md Each Session

**At the start of every session**, read CLAUDE.md and verify it matches the current code:
- Schema table reflects actual sheet columns
- Migration history is complete and up to date
- UI patterns reference actual helper functions (not stale names)
- Any removed functions or features are no longer mentioned

**After any change that affects rules, schema, or patterns**, update CLAUDE.md in the same commit — not as a separate follow-up. CLAUDE.md must never lag behind the code.

---

## MANDATORY: Commit and Push After Every Change

After completing any change (version bump included), always:

1. **Stage** the changed files (`Code.gs`, `Index.html`, `CLAUDE.md`, or whichever were modified)
2. **Commit** with a message in the format:
   ```
   vX.Y.Z — <short description of what changed>
   ```
   Example: `v1.13.0 — reorder sheet columns, fix unscheduled date format, update CLAUDE.md`
3. **Push** to the remote if one is configured:
   ```bash
   git remote -v   # check if remote exists
   git push        # push if remote is available
   ```
   If no remote is configured, commit only — do not error or warn.

**Never skip the commit step**, even for small label or style fixes. The commit history is how deployed versions are tracked.

---

## MANDATORY: Version Bump on Every Change

**Always update the version in BOTH files before finishing any change, no exceptions.**

### Code.gs (top of file)
```js
const VERSION      = 'X.Y.Z';
const VERSION_DATE = 'YYYY-MM-DD';
```
Also update the comment on line 4: `// Version: X.Y.Z (YYYY-MM-DD)`

### Index.html (line 6)
```html
<!-- Version: X.Y.Z (YYYY-MM-DD) -->
```

### Version convention
- **PATCH** (Z): UI-only changes, label tweaks, display logic, bug fixes — Index.html only needs bumping
- **MINOR** (Y): New features, new fields, new event types — bump both files
- **MAJOR** (X): Breaking schema changes, full rewrites

**Both files must always share the same version number.** When Code.gs changes, bump both. When only Index.html changes (UI/display fixes), bump Index.html to catch up — Code.gs must never be behind Index.html. At the end of every session, both files must show the same version.

---

## Schema / Data Protection Rules

### Core principle
**Deploying new code to Apps Script NEVER modifies Google Sheet data** — only the script logic changes. The sheet is the source of truth and must be treated as read-only from a structural standpoint unless an explicit migration is run.

### How reads are protected
`getRosterEntries()` maps columns by **header name** via `_rosterColMap()`, not by position. This means:
- Reordering columns in the sheet → safe, reads still work
- Adding new columns manually in the sheet → safe, ignored gracefully
- Old sheet missing a new column → safe, field defaults to empty string
- **Never hardcode `row[N]` indices** — always go through `_rosterColMap()`

### When a schema migration IS needed
Only required when adding a new column or renaming an existing header. Steps:
1. Bump the version (MINOR bump: X.Y → X.Y+1, e.g. v1.11 → v1.12)
2. Write a `migrateSchemaToVXY()` function in Code.gs. **The function name MUST match the new version number** — e.g. v1.12 → `migrateSchemaToV112()`. The function must:
   - Check if migration is already done (idempotent — safe to re-run)
   - Use `insertColumnBefore()` or `appendColumn` — never delete or overwrite data
   - Write only the new header cell; leave all data rows untouched
3. Deploy the new code **first** (old data reads safely due to header mapping)
4. Run `migrateSchemaToVXY()` **once** from the Apps Script editor
5. Verify the sheet has the new header, then the new field is live
6. **After the migration has been confirmed on the live sheet, remove the migration function from Code.gs** — it is a one-time tool, not permanent code

### Migration history
| Version | Change | Migration function |
|---------|--------|--------------------|
| v1.5.0 | Added Ice Breaker (col K), shifted Last Updated to col L | `migrateSchemaToV15()` |
| v1.6.0 | Added Time (col M) and Event ID (col N) | `migrateSchemaToV16()` + `backfillEventIds()` |
| v1.13.0 | Reordered columns for human readability: Time → col D, Ice Breaker → col I | `migrateSchemaToV113()` ✓ deleted |
| v1.16.0 | Added Can Drive (col I) to Members tab | `migrateSchemaToV116()` ✓ deleted |
| v1.17.0 | Added Older Sunday School member type; added 17 Older SS members via seeder | `importOlderSSMembers()` ✓ deleted |
| v1.17.1 | Kept Harvest member type; added 6 JAG1 Harvest members via seeder; role toggles disabled for Harvest/Older SS in edit form | `importHarvestMembers()` ✓ deleted |
| v1.18.0 | Fixed time timezone bug (UTC+8 Perth shifted 18:30→02:30); batch save for performance; Time column set to plain text | `fixTimeValues()` ✓ deleted |
| v1.20.0 | Removed Event ID column (UUID); rowIndex used for all row lookups | `migrateSchemaToV120()` ✓ deleted |
| v1.20.0 | Fix Members sheet ghost rows (Older SS/Harvest at row 1001+) | `fixMembersSheetGhostRows()` ✓ deleted |
| v1.20.1 | Notice row inserted as row 1; column headers in row 2; data from row 3 | `migrateSchemaToV121()` ✓ deleted |
| v1.21.0 | Added Sunday School as a valid Members Group; no schema change (dropdown-only) | n/a |
| v1.21.1 | Fixed stale loadData() race condition wiping updatedAt display after save | n/a |
| v1.21.2 | Added SpreadsheetApp.flush() before sortRosterSheet so getLastColumn() sees col M | n/a |
| v1.22.0 | Rebuild Last Updated col M: clear stale content/format, re-apply dd/mm/yyyy hh:mm | `rebuildLastUpdatedColumn()` |
| v1.23.0 | Performance: remove per-save setNumberFormat('@') from saves (formatSheets() covers it); fix Cancelled/Replaced optimistic cache; fix member save/delete optimistic cache | n/a |
| v1.24.0 | Performance: skip sortRosterSheet() when no date changed and no new row added | n/a |
| v1.25.0 | Performance: fix _rosterColMap() double-call in all three data functions; skip loadData() after stable saves (no sort needed) | n/a |
| v1.26.0 | Data structure: combined events (Youth Hour, Combined, Special, Cancelled, Replaced) now saved as single row with Group="Both" instead of two JAG1+JAG2 rows | `migrateRosterToGroupBoth()` |

### Current schema (v1.26.0, 13 columns — Roster tab)
> Row 1: portal notice (merged, frozen). Row 2: column headers. Row 3+: data.
> **Row structure**: Separated LG → 2 rows (JAG1 + JAG2). All other event types → 1 row (Group="Both").
| Col | Sheet Header | JS field | Notes |
|-----|-------------|----------|-------|
| A | Date | date | Formatted `ddd dd/mm/yyyy` by formatSheets() |
| B | Group | group | Dropdown: JAG1, JAG2, Both. "Both" = applies to both groups (all non-Separated events) |
| C | Event Type | eventType | Dropdown: Youth Hour, Separated LG, Combined, Special, Cancelled, Replaced |
| D | Time | time | 24h text e.g. `18:30`; blank = no fixed time |
| E | Venue | venue | |
| F | Organiser | organiser | |
| G | P&W | pw | |
| H | Facilitator | facilitator | |
| I | Ice Breaker | iceBreaker | Optional; shown in form and card **only for Youth Hour**; blank for all other event types |
| J | Food | food | |
| K | Reporting | reporting | |
| L | Notes | notes | Special events: `Label: Value\n...` per line |
| M | Last Updated | updatedAt | Auto-stamped; do not edit |

### Members tab schema (fixed, 9 columns)
| Col | Sheet Header | Notes |
|-----|-------------|-------|
| A | Name | |
| B | Group | Dropdown: JAG1, JAG2, Both, Sunday School |
| C | Can Organise | Checkbox |
| D | Can P&W | Checkbox |
| E | Can Facilitate | Checkbox |
| F | Can Report | Checkbox |
| G | Active | Checkbox |
| H | Role Type | Dropdown: Adult, Student, Older Sunday School, Harvest |
| I | Can Drive | Checkbox; used to label members as "Drive" in group share messages |

- Members tab uses positional reads (row[0]–row[8]) — column order must not change
- To add a Members column: add `migrateSchemaToVXY()` (see naming rule above) and update `getMembers()` + `saveMember()`

---

---

## Sheet Formatting (`formatSheets`)

Run `formatSheets()` from the Apps Script editor any time to apply human-readable formatting. It is **fully idempotent** — safe to re-run after any schema change. It never reads or writes data.

### What it applies
| Element | Roster | Members |
|---------|--------|---------|
| Column widths | Per field (see schema table) | Fixed per column |
| Header freeze | Row 1 | Row 1 |
| Alternating row colours | ✓ light purple / white | ✓ light purple / white |
| Date format | `ddd dd/mm/yyyy` on Date col | — |
| Datetime format | `dd/mm/yyyy hh:mm` on Last Updated | — |
| Dropdown validation | Group, Event Type | Group, Role Type |
| Checkbox validation | — | Can Organise, Can P&W, Can Facilitate, Can Report, Active |
| Portal notice row | Row 1 (merged, frozen) | Row 1 (merged, frozen) |
| Header notes | Last Updated, Time | — |

### Adding a new field — formatting checklist
1. Run schema migration (`migrateSchemaToVXY()`) to add the column header
2. Add the JS field key → pixel width to `widths` map in `_formatRosterSheet()`
3. If it needs a dropdown, add a validation block (copy the Group pattern)
4. If it's system-managed (auto-filled, not for humans), add a header note + de-emphasise text
5. Run `formatSheets()` from the editor
6. Update the schema table in this file

---

## Event Type Rules

| Event Type | Combined? | Notes |
|------------|-----------|-------|
| Youth Hour | Yes (shared) | Week 1 Friday |
| Separated LG | No | Week 2 & 4 Friday |
| Combined | Yes (shared) | Week 3 Friday |
| Special | Yes (shared) | Week 5+; only show fields that have data |
| Cancelled | N/A | Notes only |
| Replaced | N/A | Notes only |

- **All non-Separated events**: saves ONE row with Group="Both"; shared fields use `shared-*` IDs
- **Separated LG**: saves TWO rows (JAG1 + JAG2); per-group fields use `{group}-*` IDs
- **Organiser**: single shared field for combined events; per-group field for Separated LG; hide in card if empty


---

## Key UI Patterns

- `displayName(fullName)` — first name only unless duplicate first name among active members
- `resolveDisplayName(val)` — reverse lookup from display name back to full name on save
- `buildSelect(id, options, selected)` — renders `<input>` + `<datalist>` for free-text override; options show display names
- `suggestEventType(date)` — always drives the default event type preload (never use stored `eventType` as default)
- After save/delete: call `finishSave(msg, skipRefresh)` — sets `currentView='home'`, calls `setActiveNav('home')`, then `loadData()` unless `skipRefresh=true`
- `submitForm` has 4 explicit paths: Cancelled/Replaced → "Both" entry; Special → "Both" entry; Combined → "Both" entry; Separated LG → per-group entries
- `buildEventCard` combined section: `entries.find(e => e.group === 'Both') || entries.find(e => e.group === 'JAG1')` — handles both new and legacy formats
- Nav updates: use `setActiveNav(view)` — do not inline the `['home','add','members'].forEach(...)` pattern
- Card footers: use `buildCardFooter(friday, entries, ts)` — do not duplicate the Share/Edit button HTML
- Encoding entries for `onclick`: use `encodeEntries(obj)` — replaces `JSON.stringify(obj).split('"').join('&quot;')`

---

## Codebase Hygiene Rules

These rules apply on every change. Violations must be cleaned up in the same PR/commit.

1. **Remove one-time functions after use** — migration functions (`migrateSchemaToVXY`), seeder functions (`importX`), setup functions (`initializeSheets`) are tools, not permanent code. Once confirmed run on the live sheet, delete them.
2. **Remove dead code** — unused helpers, unreachable branches (old event type strings like `'Home-Based LG'`), empty function bodies. If it's not called, delete it.
3. **No duplicated patterns** — repeated nav toggle, post-save navigation, card footer HTML, and JSON encoding patterns must use the shared helpers above. If you find yourself duplicating a pattern, extract a helper first.
4. **Cache repeated computations** — don't call the same function twice in the same scope (e.g. `getFridaysRange()` called twice in `getScheduleDates()`). Assign to a variable.
5. **Use `getLastColumn()` over hardcoded column counts** — `sortRosterSheet` and similar range operations must use `sheet.getLastColumn()`, not a literal number that can drift.
