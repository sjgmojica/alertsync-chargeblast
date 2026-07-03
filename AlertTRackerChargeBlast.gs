/**
 * AlertTracker_ChargeblastSync.gs
 * ---------------------------------------------------------------
 * Syncs data between:
 *   File1: POC_ALERT_TRACKER
 *   File2: POC_ESCALATIONS TO CHARGEBLAST
 *
 * Task 1: Copy the "synced" columns from File2 -> File1 (matched by alert_id)
 * Task 2: If File1's "Refunded?" column == "NO", upsert that row into File2
 *         (matched by alert_id). If "YES" (or blank), skip.
 *
 * DUPLICATE alert_id HANDLING: if an alert_id appears more than once in
 * EITHER file, it's ambiguous which row is "correct" — that alert_id is
 * SKIPPED entirely for both tasks (not synced in either direction), and
 * a summary email is sent listing every duplicate found in that run.
 *
 * Columns are resolved by READING THE HEADER ROW at runtime, not by
 * fixed letter/index — reshuffled columns won't break the logic, as
 * long as the header TEXT stays the same.
 *
 * PERFORMANCE: all reads/writes are batched (one getValues() /
 * setValues() call per sheet per run), not per-row.
 *
 * INSTALL: Bind this script to File1 (POC_ALERT_TRACKER) via
 * Extensions > Apps Script. It reaches File2 by ID, so it doesn't
 * need to be bound there too.
 * ---------------------------------------------------------------
 */

// ====================== CONFIG — CHECK THESE ======================
const AlertSync_FILE1_ID = '1TSYnYtywInAl0LVRBdu6nR7JWV1C_QEyB1qNsPdKU2A'; // POC_ALERT_TRACKER
const AlertSync_FILE2_ID = '1Wv-NB_1M4BNLyXwNkVIntqFsG9vRl8VsXm_LrZ78KHI'; // POC_ESCALATIONS TO CHARGEBLAST

// Tab names, confirmed via AlertSync_dumpHeaders() — pinned explicitly so a future
// tab reorder in either spreadsheet can't silently break the sync.
const AlertSync_FILE1_SHEET_NAME = 'POC_ALERT_TRACKER';
const AlertSync_FILE2_SHEET_NAME = 'POC_ESCALATIONS TO CHARGEBLAST';

// Row where headers live (data starts on the next row)
const AlertSync_HEADER_ROW = 1;

// Header text for the alert_id column — must match exactly (case-insensitive,
// whitespace-trimmed) in BOTH files.
const AlertSync_ALERT_ID_HEADER = 'alert_id';

// Header text of File1's YES/NO gate column (column C = "Refunded?").
const AlertSync_REFUNDED_HEADER = 'Refunded?';

// Header text of the 3 columns that get copied File2 -> File1.
// These must use the SAME header text in both files so the script
// can match them up on either side.
const AlertSync_SYNC_HEADERS = ['Chargeblast Action Date', 'Reference Information', 'Notes from Chargeblast'];

// Max seconds to wait for the script lock before giving up (prevents
// two triggers from writing to the same sheet at the same time).
const AlertSync_LOCK_WAIT_SECONDS = 30;

// TESTING: when true, no data is written anywhere — the script only
// logs what it WOULD do (View > Logs / View > Executions after running).
// Set back to false once you've verified the logged output looks right.
const AlertSync_DRY_RUN = false;

// Email address to notify when duplicate alert_ids are found.
const AlertSync_EMAIL_NOTIFY_TO = 'test@gmail.com';

// Set to false to skip sending duplicate-notification emails entirely.
// Duplicates are still detected, skipped from syncing, and logged —
// this only silences the email step.
const AlertSync_SEND_DUPLICATE_EMAILS = false;
// ====================================================================


function AlertSync_runSync() {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(AlertSync_LOCK_WAIT_SECONDS * 1000)) {
    Logger.log('AlertSync_runSync: could not acquire lock, another run is in progress. Skipping.');
    return;
  }
  try {
    const file1Sheet = AlertSync_getSheet_(AlertSync_FILE1_ID, AlertSync_FILE1_SHEET_NAME);
    const file2Sheet = AlertSync_getSheet_(AlertSync_FILE2_ID, AlertSync_FILE2_SHEET_NAME);

    const file1Headers = AlertSync_getHeaderMap_(file1Sheet);
    const file2Headers = AlertSync_getHeaderMap_(file2Sheet);
    const file1AlertCol = AlertSync_getColIndex_(file1Headers, AlertSync_ALERT_ID_HEADER, 'File1');
    const file2AlertCol = AlertSync_getColIndex_(file2Headers, AlertSync_ALERT_ID_HEADER, 'File2');

    const file1Data = file1Sheet.getDataRange().getValues();
    const file2Data = file2Sheet.getDataRange().getValues();

    // Find duplicate alert_ids in EACH file up front. Any alert_id that's
    // duplicated anywhere is skipped for BOTH tasks — safer than guessing
    // which row is the "real" one.
    const file1Dupes = AlertSync_findDuplicateAlertIds_(file1Data, file1AlertCol, 'File1');
    const file2Dupes = AlertSync_findDuplicateAlertIds_(file2Data, file2AlertCol, 'File2');
    const skipIds = AlertSync_unionSets_(file1Dupes.idToRows, file2Dupes.idToRows);

    const task1Result = AlertSync_syncChargeblastFieldsToTracker_(
      file1Data, file2Data, file1Headers, file2Headers, file1AlertCol, file2AlertCol, skipIds
    );
    const task2Result = AlertSync_upsertUnrefundedRowsToChargeblast_(
      file1Data, file2Data, file1Headers, file2Headers, file1AlertCol, file2AlertCol, skipIds
    );

    if (AlertSync_DRY_RUN) {
      Logger.log('[DRY RUN] No changes written. Task1 would change ' +
        task1Result.changed + '; Task2 would update ' + task2Result.file2Changed +
        ' / insert ' + task2Result.newRows.length + ' new row(s).');
    } else {
      if (task1Result.changed) {
        file1Sheet.getRange(1, 1, file1Data.length, file1Data[0].length).setValues(file1Data);
      }
      if (task2Result.file2Changed) {
        file2Sheet.getRange(1, 1, file2Data.length, file2Data[0].length).setValues(file2Data);
      }
      if (task2Result.newRows.length > 0) {
        const startRow = file2Sheet.getLastRow() + 1;
        file2Sheet.getRange(startRow, 1, task2Result.newRows.length, file2Data[0].length)
          .setValues(task2Result.newRows);
      }
    }

    // Consolidated summary — every alert_id actually touched this run,
    // so you can cross-check against the sheets once it's done.
    const summary = [];
    summary.push('===== SYNC SUMMARY' + (AlertSync_DRY_RUN ? ' (DRY RUN — nothing written)' : '') + ' =====');
    summary.push('Task1 updated (' + task1Result.updatedIds.length + '): ' +
      (task1Result.updatedIds.length ? task1Result.updatedIds.join(', ') : 'none'));
    summary.push('Task2 updated (' + task2Result.updatedIds.length + '): ' +
      (task2Result.updatedIds.length ? task2Result.updatedIds.join(', ') : 'none'));
    summary.push('Task2 inserted (' + task2Result.insertedIds.length + '): ' +
      (task2Result.insertedIds.length ? task2Result.insertedIds.join(', ') : 'none'));
    summary.push('Duplicates skipped (' + skipIds.size + '): ' +
      (skipIds.size ? Array.from(skipIds).join(', ') : 'none'));
    Logger.log(summary.join('\n'));

    // One consolidated email per run, covering duplicates found in either file.
    if (skipIds.size > 0) {
      AlertSync_notifyDuplicates_(file1Sheet, file2Sheet, file1Dupes, file2Dupes);
    }
  } finally {
    lock.releaseLock();
  }
}

/**
 * Scans a sheet's data (2D array, header row already excluded via
 * AlertSync_HEADER_ROW) for alert_ids that appear more than once.
 * Returns { idToRows: Map<alertId, [sheetRow, ...]> } — only entries
 * with 2+ rows are included.
 */
function AlertSync_findDuplicateAlertIds_(data, alertCol) {
  const seen = {}; // alertId -> [sheetRow, ...]
  for (let r = AlertSync_HEADER_ROW; r < data.length; r++) {
    const alertId = AlertSync_normalizeId_(data[r][alertCol - 1]);
    if (!alertId) continue;
    if (!seen[alertId]) seen[alertId] = [];
    seen[alertId].push(r + 1); // 1-indexed sheet row
  }
  const dupes = {};
  Object.keys(seen).forEach(id => {
    if (seen[id].length > 1) dupes[id] = seen[id];
  });
  return { idToRows: dupes };
}

function AlertSync_unionSets_(mapA, mapB) {
  const result = new Set();
  Object.keys(mapA).forEach(id => result.add(id));
  Object.keys(mapB).forEach(id => result.add(id));
  return result;
}

/** yyyy-MM-dd HH:mm:ss in the script's timezone. */
function AlertSync_formatTimestamp_() {
  return Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');
}

/**
 * Builds the subject/body for a duplicate-alert_id notification from a
 * flat list of entries: [{ filename, alertId, count }, ...]
 *
 * Subject: [yyyy-MM-dd HH:mm:ss] Alert Sync Duplicate alert_id
 * Body (repeated per entry):
 *   [yyyy-MM-dd HH:mm:ss] Alert Sync Duplicate
 *   Filename: <filename>
 *   Alert_id: <alert_id>
 *   No of duplicates: <count>
 */
function AlertSync_buildDuplicateEmail_(entries) {
  const timestamp = AlertSync_formatTimestamp_();
  const subject = '[' + timestamp + '] Alert Sync Duplicate alert_id';

  const lines = [];
  entries.forEach(e => {
    lines.push('[' + timestamp + '] Alert Sync Duplicate');
    lines.push('Filename: ' + e.filename);
    lines.push('Alert_id: ' + e.alertId);
    lines.push('No of duplicates: ' + e.count);
    lines.push('');
  });
  if (AlertSync_DRY_RUN) lines.push('(AlertSync_DRY_RUN is ON — this was a test run; nothing was written.)');

  return { subject, body: lines.join('\n') };
}

/**
 * Sends a summary email listing every duplicate alert_id found in this
 * run: which file, the alert_id, and how many times it's duplicated.
 */
function AlertSync_notifyDuplicates_(file1Sheet, file2Sheet, file1Dupes, file2Dupes) {
  const file1Name = file1Sheet.getParent().getName();
  const file2Name = file2Sheet.getParent().getName();

  const entries = [];
  Object.keys(file1Dupes.idToRows).forEach(id =>
    entries.push({ filename: file1Name, alertId: id, count: file1Dupes.idToRows[id].length }));
  Object.keys(file2Dupes.idToRows).forEach(id =>
    entries.push({ filename: file2Name, alertId: id, count: file2Dupes.idToRows[id].length }));

  if (entries.length === 0) return;

  if (!AlertSync_SEND_DUPLICATE_EMAILS) {
    Logger.log('Duplicate alert_id(s) found (email disabled): ' +
      entries.map(e => e.filename + ': ' + e.alertId + ' (x' + e.count + ')').join('; '));
    return;
  }

  const to = AlertSync_EMAIL_NOTIFY_TO || Session.getEffectiveUser().getEmail();
  if (!to) {
    Logger.log('AlertSync_notifyDuplicates_: no email address available, skipping email. See log for duplicate list.');
    return;
  }

  const { subject, body } = AlertSync_buildDuplicateEmail_(entries);
  MailApp.sendEmail(to, subject, body);
  Logger.log('Duplicate notification emailed to ' + to);
}

/**
 * TASK 1
 * Copy AlertSync_SYNC_HEADERS columns from File2 into File1, matched on alert_id.
 * Mutates file1Data in place; caller decides whether/how to write it.
 * Any alert_id in skipIds is left untouched.
 */
function AlertSync_syncChargeblastFieldsToTracker_(file1Data, file2Data, file1Headers, file2Headers, file1AlertCol, file2AlertCol, skipIds) {
  const file2SyncCols = AlertSync_SYNC_HEADERS.map(h => AlertSync_getColIndex_(file2Headers, h, 'File2'));
  const file1SyncCols = AlertSync_SYNC_HEADERS.map(h => AlertSync_getColIndex_(file1Headers, h, 'File1'));

  const file2Map = {}; // normalized alert_id -> [val, val, val]
  for (let r = AlertSync_HEADER_ROW; r < file2Data.length; r++) {
    const row = file2Data[r];
    const alertId = AlertSync_normalizeId_(row[file2AlertCol - 1]);
    if (!alertId || skipIds.has(alertId)) continue;
    file2Map[alertId] = file2SyncCols.map(c => row[c - 1]);
  }

  let changed = false;
  const updatedIds = [];
  for (let r = AlertSync_HEADER_ROW; r < file1Data.length; r++) {
    const alertId = AlertSync_normalizeId_(file1Data[r][file1AlertCol - 1]);
    if (!alertId || skipIds.has(alertId)) continue;
    const match = file2Map[alertId];
    if (!match) continue;

    let rowChanged = false;
    file1SyncCols.forEach((col, i) => {
      if (!AlertSync_valuesEqual_(file1Data[r][col - 1], match[i])) {
        Logger.log('[Task1] alert_id ' + alertId + ': File1 row ' + (r + 1) +
          ' col ' + col + ' "' + file1Data[r][col - 1] + '" -> "' + match[i] + '"');
        file1Data[r][col - 1] = match[i];
        rowChanged = true;
      }
    });
    if (rowChanged) {
      changed = true;
      updatedIds.push(alertId);
    }
  }

  return { changed, updatedIds };
}

/**
 * TASK 2
 * For every File1 row where "Refunded?" === "NO", upsert it into
 * File2 (matched by alert_id). "YES" (or blank) rows are skipped, and
 * any alert_id in skipIds is skipped regardless of Refunded? value.
 * Mutates file2Data in place for updates; returns newRows separately
 * for the caller to append.
 */
function AlertSync_upsertUnrefundedRowsToChargeblast_(file1Data, file2Data, file1Headers, file2Headers, file1AlertCol, file2AlertCol, skipIds) {
  const refundedCol = AlertSync_getColIndex_(file1Headers, AlertSync_REFUNDED_HEADER, 'File1');

  const file2LastCol = file2Data[0].length;
  const colMap = []; // index = file2 col# - 1, value = file1 col# (or null)
  for (const [header, file2Col] of Object.entries(file2Headers)) {
    colMap[file2Col - 1] = file1Headers[header] || null;
  }

  const file2RowIndexByAlertId = {};
  for (let r = AlertSync_HEADER_ROW; r < file2Data.length; r++) {
    const alertId = AlertSync_normalizeId_(file2Data[r][file2AlertCol - 1]);
    if (alertId && !skipIds.has(alertId)) file2RowIndexByAlertId[alertId] = r;
  }

  let file2Changed = false;
  const newRows = [];
  const updatedIds = [];
  const insertedIds = [];

  for (let r = AlertSync_HEADER_ROW; r < file1Data.length; r++) {
    const row = file1Data[r];
    const alertId = AlertSync_normalizeId_(row[file1AlertCol - 1]);
    if (!alertId || skipIds.has(alertId)) continue;

    const refundedValue = String(row[refundedCol - 1] || '').trim().toUpperCase();
    if (refundedValue !== 'NO') continue;

    const existingIndex = file2RowIndexByAlertId[alertId];
    if (existingIndex !== undefined) {
      // Row already exists in File2 — upsert only the columns that
      // actually changed, leaving every other cell in that row alone
      // (including any File2-only column with no source in File1).
      let rowChanged = false;
      for (let c = 0; c < file2LastCol; c++) {
        const file1Col = colMap[c];
        if (!file1Col) continue; // no matching File1 column — leave untouched
        const newVal = row[file1Col - 1];
        if (!AlertSync_valuesEqual_(file2Data[existingIndex][c], newVal)) {
          Logger.log('[Task2] alert_id ' + alertId + ': File2 row ' + (existingIndex + 1) +
            ' col ' + (c + 1) + ' "' + file2Data[existingIndex][c] + '" -> "' + newVal + '"');
          file2Data[existingIndex][c] = newVal;
          rowChanged = true;
        }
      }
      if (rowChanged) {
        file2Changed = true;
        updatedIds.push(alertId);
      }
    } else {
      // No existing row for this alert_id — insert a full new row,
      // copying every mapped column regardless of whether it's
      // populated or blank in File1.
      const rowToWrite = [];
      for (let c = 0; c < file2LastCol; c++) {
        const file1Col = colMap[c];
        rowToWrite.push(file1Col ? row[file1Col - 1] : '');
      }
      Logger.log('[Task2] alert_id ' + alertId + ': would INSERT new row into File2.');
      newRows.push(rowToWrite);
      insertedIds.push(alertId);
      file2RowIndexByAlertId[alertId] = file2Data.length + newRows.length - 1;
    }
  }

  return { file2Changed, newRows, updatedIds, insertedIds };
}

/**
 * Helper: open a spreadsheet by ID and return the given tab,
 * or the first tab if no name is configured / found.
 */
function AlertSync_getSheet_(spreadsheetId, sheetName) {
  const ss = SpreadsheetApp.openById(spreadsheetId);
  if (sheetName) {
    const sheet = ss.getSheetByName(sheetName);
    if (sheet) return sheet;
  }
  return ss.getSheets()[0];
}

/**
 * Reads row AlertSync_HEADER_ROW and returns a map of
 * lowercased/trimmed header text -> 1-indexed column number.
 */
function AlertSync_getHeaderMap_(sheet) {
  const lastCol = sheet.getLastColumn();
  const headerRowValues = sheet.getRange(AlertSync_HEADER_ROW, 1, 1, lastCol).getValues()[0];
  const map = {};
  headerRowValues.forEach((h, i) => {
    const key = String(h).trim().toLowerCase();
    if (key) map[key] = i + 1; // 1-indexed column
  });
  return map;
}

/**
 * Looks up a header's column index (case-insensitive, trimmed).
 * Throws a clear error if the header can't be found, instead of
 * silently writing to the wrong column.
 */
function AlertSync_getColIndex_(headerMap, headerName, fileLabel) {
  const key = String(headerName).trim().toLowerCase();
  const col = headerMap[key];
  if (!col) {
    throw new Error(
      'Header "' + headerName + '" not found in ' + fileLabel +
      '. Check CONFIG constants match the actual header text in row ' + AlertSync_HEADER_ROW + '.'
    );
  }
  return col;
}

/**
 * Normalizes an alert_id for comparison: converts to string and trims.
 * Prevents a numeric-looking alert_id (e.g. 1074390846) stored as a
 * Number in one sheet and a String in the other from failing to match.
 */
function AlertSync_normalizeId_(val) {
  if (val === null || val === undefined) return '';
  return String(val).trim();
}

/**
 * Normalizes a single cell value for equality comparison. Sheets can
 * return the same real-world value as different JS types (e.g. a
 * date-formatted cell comes back as a Date object in one sheet but a
 * plain string in another) — without this, those would look "changed"
 * on every run even though nothing actually differs.
 */
function AlertSync_normalizeForCompare_(v) {
  if (v instanceof Date) return v.getTime();
  if (v === null || v === undefined) return '';
  return String(v).trim();
}

function AlertSync_valuesEqual_(a, b) {
  return AlertSync_normalizeForCompare_(a) === AlertSync_normalizeForCompare_(b);
}

function AlertSync_rowsEqual_(rowA, rowB) {
  if (rowA.length !== rowB.length) return false;
  return rowA.every((v, i) => AlertSync_valuesEqual_(v, rowB[i]));
}

/**
 * "1 duplicate alert_id" vs "2 duplicate alert_ids" — grammatically
 * correct count + noun, used in email/log wording.
 */
function AlertSync_pluralize_(count, singular, plural) {
  return count + ' ' + (count === 1 ? singular : (plural || singular + 's'));
}


// ====================================================================
// DEBUG — run this if AlertSync_testConfig() reports a missing header. Logs
// the actual tab name and raw header row contents for both files,
// with no interpretation, so you can see exactly what's really there.
// ====================================================================
function AlertSync_dumpHeaders() {
  const out = [];

  [['File1', AlertSync_FILE1_ID, AlertSync_FILE1_SHEET_NAME], ['File2', AlertSync_FILE2_ID, AlertSync_FILE2_SHEET_NAME]].forEach(([label, id, sheetName]) => {
    try {
      const ss = SpreadsheetApp.openById(id);
      out.push(label + ' spreadsheet name: "' + ss.getName() + '"');
      out.push(label + ' all tabs: ' + ss.getSheets().map(s => '"' + s.getName() + '"').join(', '));

      const sheet = AlertSync_getSheet_(id, sheetName);
      out.push(label + ' tab being used: "' + sheet.getName() + '"');

      const lastCol = sheet.getLastColumn();
      const headerRow = sheet.getRange(AlertSync_HEADER_ROW, 1, 1, lastCol).getValues()[0];
      out.push(label + ' row ' + AlertSync_HEADER_ROW + ' raw values (' + lastCol + ' columns):');
      headerRow.forEach((h, i) => {
        out.push('  col ' + (i + 1) + ': "' + h + '"');
      });
      out.push('');
    } catch (e) {
      out.push(label + ' ERROR: ' + e);
      out.push('');
    }
  });

  Logger.log(out.join('\n'));
}


// ====================================================================
// DEBUG — inspect exactly why one alert_id's synced columns aren't
// transferring. Pass an alert_id you know has a value in File2's
// "Chargeblast Action Date" (or any AlertSync_SYNC_HEADERS column) that isn't
// showing up in File1. Logs raw value, JS type, and whether the
// equality check thinks they already match — no writes.
// ====================================================================
function AlertSync_debugColumnSync(alertId) {
  const targetId = AlertSync_normalizeId_(alertId);
  const out = [];

  const file1Sheet = AlertSync_getSheet_(AlertSync_FILE1_ID, AlertSync_FILE1_SHEET_NAME);
  const file2Sheet = AlertSync_getSheet_(AlertSync_FILE2_ID, AlertSync_FILE2_SHEET_NAME);
  const file1Headers = AlertSync_getHeaderMap_(file1Sheet);
  const file2Headers = AlertSync_getHeaderMap_(file2Sheet);
  const file1AlertCol = AlertSync_getColIndex_(file1Headers, AlertSync_ALERT_ID_HEADER, 'File1');
  const file2AlertCol = AlertSync_getColIndex_(file2Headers, AlertSync_ALERT_ID_HEADER, 'File2');

  const file1Data = file1Sheet.getDataRange().getValues();
  const file2Data = file2Sheet.getDataRange().getValues();

  let file1Row = null, file2Row = null;
  for (let r = AlertSync_HEADER_ROW; r < file1Data.length; r++) {
    if (AlertSync_normalizeId_(file1Data[r][file1AlertCol - 1]) === targetId) { file1Row = r; break; }
  }
  for (let r = AlertSync_HEADER_ROW; r < file2Data.length; r++) {
    if (AlertSync_normalizeId_(file2Data[r][file2AlertCol - 1]) === targetId) { file2Row = r; break; }
  }

  out.push('alert_id: ' + targetId);
  out.push('Found in File1 at sheet row: ' + (file1Row !== null ? file1Row + 1 : 'NOT FOUND'));
  out.push('Found in File2 at sheet row: ' + (file2Row !== null ? file2Row + 1 : 'NOT FOUND'));
  out.push('');

  if (file1Row === null || file2Row === null) {
    Logger.log(out.join('\n'));
    return;
  }

  AlertSync_SYNC_HEADERS.forEach(header => {
    const c1 = AlertSync_getColIndex_(file1Headers, header, 'File1');
    const c2 = AlertSync_getColIndex_(file2Headers, header, 'File2');
    const v1 = file1Data[file1Row][c1 - 1];
    const v2 = file2Data[file2Row][c2 - 1];
    out.push('Column: "' + header + '"');
    out.push('  File1 col ' + c1 + ': raw=' + JSON.stringify(v1) + '  type=' + typeof v1 + (v1 instanceof Date ? ' (Date)' : ''));
    out.push('  File2 col ' + c2 + ': raw=' + JSON.stringify(v2) + '  type=' + typeof v2 + (v2 instanceof Date ? ' (Date)' : ''));
    out.push('  normalized File1: ' + JSON.stringify(AlertSync_normalizeForCompare_(v1)));
    out.push('  normalized File2: ' + JSON.stringify(AlertSync_normalizeForCompare_(v2)));
    out.push('  AlertSync_valuesEqual_: ' + AlertSync_valuesEqual_(v1, v2));
    out.push('');
  });

  Logger.log(out.join('\n'));
}



function AlertSync_testConfig() {
  const results = [];
  try {
    const file1Sheet = AlertSync_getSheet_(AlertSync_FILE1_ID, AlertSync_FILE1_SHEET_NAME);
    const file2Sheet = AlertSync_getSheet_(AlertSync_FILE2_ID, AlertSync_FILE2_SHEET_NAME);
    results.push('File1 tab: "' + file1Sheet.getName() + '" (' + (file1Sheet.getLastRow() - AlertSync_HEADER_ROW) + ' data rows)');
    results.push('File2 tab: "' + file2Sheet.getName() + '" (' + (file2Sheet.getLastRow() - AlertSync_HEADER_ROW) + ' data rows)');

    const file1Headers = AlertSync_getHeaderMap_(file1Sheet);
    const file2Headers = AlertSync_getHeaderMap_(file2Sheet);

    const check = (label, headerName, headerMap, fileLabel) => {
      try {
        const col = AlertSync_getColIndex_(headerMap, headerName, fileLabel);
        results.push('OK   ' + label + ' -> ' + fileLabel + ' column ' + col + ' ("' + headerName + '")');
      } catch (e) {
        results.push('FAIL ' + label + ' -> NOT FOUND in ' + fileLabel + ' ("' + headerName + '")');
      }
    };

    check('alert_id', AlertSync_ALERT_ID_HEADER, file1Headers, 'File1');
    check('alert_id', AlertSync_ALERT_ID_HEADER, file2Headers, 'File2');
    check('Refunded?', AlertSync_REFUNDED_HEADER, file1Headers, 'File1');
    AlertSync_SYNC_HEADERS.forEach(h => {
      check('sync col', h, file1Headers, 'File1');
      check('sync col', h, file2Headers, 'File2');
    });

    const file1Data = file1Sheet.getDataRange().getValues();
    const file2Data = file2Sheet.getDataRange().getValues();
    const file1AlertCol = AlertSync_getColIndex_(file1Headers, AlertSync_ALERT_ID_HEADER, 'File1');
    const file2AlertCol = AlertSync_getColIndex_(file2Headers, AlertSync_ALERT_ID_HEADER, 'File2');
    const file1Dupes = AlertSync_findDuplicateAlertIds_(file1Data, file1AlertCol);
    const file2Dupes = AlertSync_findDuplicateAlertIds_(file2Data, file2AlertCol);
    results.push('');
    results.push('Duplicate alert_ids in File1: ' + Object.keys(file1Dupes.idToRows).length);
    results.push('Duplicate alert_ids in File2: ' + Object.keys(file2Dupes.idToRows).length);

    results.push('');
    results.push(AlertSync_DRY_RUN
      ? 'AlertSync_DRY_RUN is ON — running AlertSync_runSync() next will only log planned changes, nothing will be written.'
      : 'AlertSync_DRY_RUN is OFF — running AlertSync_runSync() next will WRITE real changes to both files.');
  } catch (e) {
    results.push('ERROR: ' + e);
  }

  Logger.log(results.join('\n'));
}


// ====================================================================
// AUTOMATION — run AlertSync_setupTriggers() ONCE (manually, from the editor)
// to wire everything up. Safe to re-run; it clears old triggers first.
// ====================================================================

function AlertSync_setupTriggers() {
  AlertSync_removeTriggers_();

  ScriptApp.newTrigger('AlertSync_runSync')
    .timeBased()
    .everyMinutes(15)
    .create();

  ScriptApp.newTrigger('AlertSync_onEditFile1_')
    .forSpreadsheet(SpreadsheetApp.openById(AlertSync_FILE1_ID))
    .onEdit()
    .create();

  ScriptApp.newTrigger('AlertSync_onEditFile2_')
    .forSpreadsheet(SpreadsheetApp.openById(AlertSync_FILE2_ID))
    .onEdit()
    .create();

  Logger.log('Triggers installed: 1 time-based + 2 on-edit.');
}

// This project is shared with other scripts (Returns, CSV-Tools, etc.)
// bound to the same spreadsheet. ONLY remove triggers whose handler
// function belongs to this sync — never touch unrelated triggers.
const AlertSync_OWNED_HANDLERS = ['AlertSync_runSync', 'AlertSync_onEditFile1_', 'AlertSync_onEditFile2_'];

function AlertSync_removeTriggers_() {
  ScriptApp.getProjectTriggers()
    .filter(t => AlertSync_OWNED_HANDLERS.includes(t.getHandlerFunction()))
    .forEach(t => ScriptApp.deleteTrigger(t));
}

// Edits in File1 trigger a full sync (needed for Task 2 — any edit
// could be the one that sets Refunded? to "NO").
function AlertSync_onEditFile1_(e) {
  AlertSync_runSync();
}

// Edits in File2: only react if the edit touched one of AlertSync_SYNC_HEADERS.
// Reads File1 ONCE (not per-row), applies all matching updates to an
// in-memory copy, and writes back in a single batched call. Skips (and
// reports) any edited alert_id that's duplicated in File1 or File2.
function AlertSync_onEditFile2_(e) {
  try {
    if (!e || !e.range) { AlertSync_runSync(); return; } // fallback if event is missing

    const file2Sheet = e.range.getSheet();
    const file2Headers = AlertSync_getHeaderMap_(file2Sheet);
    const syncCols = AlertSync_SYNC_HEADERS.map(h => AlertSync_getColIndex_(file2Headers, h, 'File2'));
    const alertCol = AlertSync_getColIndex_(file2Headers, AlertSync_ALERT_ID_HEADER, 'File2');

    const startCol = e.range.getColumn();
    const endCol = startCol + e.range.getNumColumns() - 1;
    const touchesSyncCols = syncCols.some(c => c >= startCol && c <= endCol);
    if (!touchesSyncCols) return; // edit was outside the synced columns

    const startRow = e.range.getRow();
    const numRows = e.range.getNumRows();

    const editedRows = [];
    for (let i = 0; i < numRows; i++) {
      const sheetRow = startRow + i;
      if (sheetRow <= AlertSync_HEADER_ROW) continue;

      const rowValues = file2Sheet
        .getRange(sheetRow, 1, 1, file2Sheet.getLastColumn())
        .getValues()[0];
      const alertId = AlertSync_normalizeId_(rowValues[alertCol - 1]);
      if (!alertId) continue;

      editedRows.push({ alertId, values: syncCols.map(c => rowValues[c - 1]) });
    }
    if (editedRows.length === 0) return;

    const lock = LockService.getScriptLock();
    if (!lock.tryLock(AlertSync_LOCK_WAIT_SECONDS * 1000)) {
      Logger.log('AlertSync_onEditFile2_: could not acquire lock. Skipping — the 15-min timer will catch it.');
      return;
    }

    try {
      const file1Sheet = AlertSync_getSheet_(AlertSync_FILE1_ID, AlertSync_FILE1_SHEET_NAME);
      const file1Headers = AlertSync_getHeaderMap_(file1Sheet);
      const file1AlertCol = AlertSync_getColIndex_(file1Headers, AlertSync_ALERT_ID_HEADER, 'File1');
      const file1SyncCols = AlertSync_SYNC_HEADERS.map(h => AlertSync_getColIndex_(file1Headers, h, 'File1'));

      const file1Data = file1Sheet.getDataRange().getValues();
      const file1Dupes = AlertSync_findDuplicateAlertIds_(file1Data, file1AlertCol);

      // Also check duplicates among the edited alert_ids themselves
      // (e.g. a paste that includes the same alert_id twice).
      const editedIdCounts = {};
      editedRows.forEach(({ alertId }) => {
        editedIdCounts[alertId] = (editedIdCounts[alertId] || 0) + 1;
      });

      const rowIndexByAlertId = {};
      for (let r = AlertSync_HEADER_ROW; r < file1Data.length; r++) {
        const id = AlertSync_normalizeId_(file1Data[r][file1AlertCol - 1]);
        if (id) rowIndexByAlertId[id] = r;
      }

      const skipped = []; // [{ alertId, filename, count }, ...]
      let changed = false;
      const file1Name = file1Sheet.getParent().getName();
      const file2Name = file2Sheet.getParent().getName();
      editedRows.forEach(({ alertId, values }) => {
        if (file1Dupes.idToRows[alertId]) {
          skipped.push({ alertId, filename: file1Name, count: file1Dupes.idToRows[alertId].length });
          return;
        }
        if (editedIdCounts[alertId] > 1) {
          skipped.push({ alertId, filename: file2Name, count: editedIdCounts[alertId] });
          return;
        }
        const rowIndex = rowIndexByAlertId[alertId];
        if (rowIndex === undefined) return; // no matching row in File1
        file1SyncCols.forEach((col, idx) => {
          file1Data[rowIndex][col - 1] = values[idx];
        });
        changed = true;
      });

      if (changed && !AlertSync_DRY_RUN) {
        file1Sheet.getRange(1, 1, file1Data.length, file1Data[0].length).setValues(file1Data);
      } else if (changed && AlertSync_DRY_RUN) {
        Logger.log('[DRY RUN][AlertSync_onEditFile2_] Would write updated sync columns to File1.');
      }

      if (skipped.length > 0) {
        if (AlertSync_SEND_DUPLICATE_EMAILS) {
          const to = AlertSync_EMAIL_NOTIFY_TO || Session.getEffectiveUser().getEmail();
          if (to) {
            const { subject, body } = AlertSync_buildDuplicateEmail_(skipped);
            MailApp.sendEmail(to, subject, body);
          }
        }
        Logger.log('AlertSync_onEditFile2_ skipped (email ' + (AlertSync_SEND_DUPLICATE_EMAILS ? 'sent' : 'disabled') + '): ' +
          skipped.map(s => s.alertId).join(', '));
      }
    } finally {
      lock.releaseLock();
    }
  } catch (err) {
    Logger.log('AlertSync_onEditFile2_ error: ' + err);
  }
}