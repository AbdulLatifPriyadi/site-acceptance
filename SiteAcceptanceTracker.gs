// ============================================================
// Google Apps Script — Site Acceptance Tracker
// Attach this script to your Google Spreadsheet:
//   Extensions → Apps Script → paste this code → Save → Deploy
// ============================================================

// --- CONFIGURATION ---
// TRACKER_SHEET: name of the sheet that holds your site data
var TRACKER_SHEET = 'Sheet1';   // <-- change if your sheet tab has a different name

// FORWARD_SHEET: name of the sheet where forward submissions are logged
var FORWARD_SHEET = 'Forward Log';  // <-- change if you prefer a different name

// CLOCKIN_SHEET: name of the sheet where clock-in / clock-out data is stored
var CLOCKIN_SHEET = 'Clockin';  // <-- change if your sheet tab has a different name

// Column indices in TRACKER_SHEET (0-based, A=0, B=1, etc.)
var COL = {
  ring:        1,  // B — RING
  seq:         2,  // C — Seq
  batchRfi:    3,  // D — Batch RFI
  batchCsg:    4,  // E — Batch CSG
  tpSiteId:    5,  // F — TP Site ID
  surgeId:     6,  // G — Surge ID
  duId:        7,  // H — DU ID
  siteName:    8,  // I — Site name
  subcontPlan: 9,  // J — Subcont Plan
  gap:        10,  // K — GAP Analysis
  remark:     11,  // L — Remark
  fo:         12,  // M — FO
  city:       13,  // N — City
  tp:         14,  // O — TP
  km:         15,  // P — Km
  peOwner:    16,  // Q — PE Owner
  lat:        17,  // R — Latitude
  lng:        18,  // S — Longitude
  subcontCsg: 24,  // Y — Subcont for Batch CSG (used by the PRPO tab)
  csgStatus:  26,  // AA — CSG Status
  priority:   32   // AG — Priority
};

// --- COLUMN HEADERS for Forward Log sheet ---
var FWD_COL = {
  timestamp:  0,  // A
  teamLeader: 1,  // B
  duId:       2,  // C
  siteName:   3,  // D
  message:    4,  // E
  remark:     5,  // F
  state:      6,  // G
  isLatest:   7   // H
};

// --- HELPER: get or create the Forward Log sheet ---
function getForwardSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(FORWARD_SHEET);
  if (!sheet) {
    sheet = ss.insertSheet(FORWARD_SHEET);
    // Add headers
    sheet.getRange(1, 1, 1, 8).setValues([[
      'Timestamp', 'Team Leader', 'DU ID', 'Site Name',
      'Message', 'Remark', 'State JSON', 'Is Latest'
    ]]);
    sheet.getRange(1, 1, 1, 8)
      .setFontWeight('bold')
      .setBackground('#1c2530')
      .setFontColor('#ffffff');
    sheet.setFrozenRows(1);
  }
  return sheet;
}

// --- HELPER: get tracker sheet ---
function getTrackerSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(TRACKER_SHEET);
  if (!sheet) {
    // Try the first sheet as fallback
    var sheets = ss.getSheets();
    if (sheets.length > 0) {
      return sheets[0];
    }
    throw new Error('Tracker sheet not found. Please check TRACKER_SHEET config.');
  }
  return sheet;
}

// --- HELPER: get clockin sheet ---
function getClockinSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(CLOCKIN_SHEET);
  if (!sheet) {
    throw new Error('Clockin sheet not found. Please check CLOCKIN_SHEET config.');
  }
  return sheet;
}

// --- HELPER: parse float from various formats (e.g. "-6,91511" → -6.91511) ---
function parseNum(v) {
  if (v === null || v === undefined || v === '') return null;
  var s = String(v).trim();
  if (s === '' || s === '-') return null;
  // Replace comma with dot for European decimal notation
  s = s.replace(',', '.');
  var n = parseFloat(s);
  return isNaN(n) ? null : n;
}

// --- DO GET: serve tracker data, history, summaries ---
function doGet(e) {
  var param = e.parameter;
  var type = param.type || '';

  try {
    if (type === 'tracker') {
      return serveTracker();
    } else if (type === 'summary') {
      return serveSummary();
    } else if (type === 'clockin') {
      return serveClockin();
    } else if (type === 'plan') {
      return servePlan(param.date);
    } else if (type === 'planDates') {
      return servePlanDates();
    } else if (type === 'activeDate') {
      return serveActiveDate();
    } else if (type === 'planRoster') {
      return servePlanRoster();
    } else if (type === 'planRosterDoc') {
      return servePlanRosterDoc();
    } else if (type === 'planRosterEdits') {
      return servePlanRosterEdits();
    } else {
      // Default: return full forward log
      return serveForwardLog();
    }
  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ error: err.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// Serve clockin data from CLOCKIN_SHEET
// Returns { [DU_ID]: { name, team, clockIn, clockOut } }
// clockOut being empty/null means the person is still working
function serveClockin() {
  try {
    var sheet = getClockinSheet();
    var lastRow = sheet.getLastRow();

    if (lastRow < 2) {
      return ContentService
        .createTextOutput(JSON.stringify({}))
        .setMimeType(ContentService.MimeType.JSON);
    }

    // Read headers from row 1 (auto-detect columns)
    var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    var colMap = {};
    for (var ci = 0; ci < headers.length; ci++) {
      var h = String(headers[ci] || '').trim().toLowerCase();
      colMap[h] = ci;
    }

    // Auto-detect column indices by header name (exact match first)
    var idxDuId    = colMap['du id']   !== undefined ? colMap['du id']   : -1;
    var idxName    = colMap['name']    !== undefined ? colMap['name']    :
                     colMap['nama']    !== undefined ? colMap['nama']    : -1;
    var idxClockIn  = colMap['clock in']  !== undefined ? colMap['clock in']  : -1;
    var idxClockOut = colMap['clock out'] !== undefined ? colMap['clock out'] : -1;
    var idxTeam    = colMap['team']    !== undefined ? colMap['team']    :
                     colMap['subcont'] !== undefined ? colMap['subcont'] : -1;
    var idxAccount = colMap['account']      !== undefined ? colMap['account']      :
                     colMap['account id']   !== undefined ? colMap['account id']   :
                     colMap['accountid']    !== undefined ? colMap['accountid']    : -1;

    // Partial match fallback
    if (idxDuId === -1) {
      for (var k in colMap) {
        if (k.indexOf('du') !== -1 && k.indexOf('id') !== -1) { idxDuId = colMap[k]; break; }
      }
    }
    if (idxAccount === -1) {
      for (var kAcc in colMap) {
        if (kAcc.indexOf('account') !== -1) { idxAccount = colMap[kAcc]; break; }
      }
    }
    if (idxClockIn === -1) {
      for (var k2 in colMap) {
        if (k2.indexOf('clock') !== -1 && k2.indexOf('in') !== -1 && k2.indexOf('out') === -1) { idxClockIn = colMap[k2]; break; }
      }
    }
    if (idxClockOut === -1) {
      for (var k3 in colMap) {
        if (k3.indexOf('clock') !== -1 && k3.indexOf('out') !== -1) { idxClockOut = colMap[k3]; break; }
      }
    }

    var data = sheet.getRange(2, 1, lastRow - 1, headers.length).getValues();
    var result = {};

    for (var i = 0; i < data.length; i++) {
      var row = data[i];
      var duId = idxDuId >= 0 ? String(row[idxDuId] || '').trim().toUpperCase() : '';
      if (!duId) continue;

      var clockOut = idxClockOut >= 0 ? String(row[idxClockOut] || '').trim() : '';
      result[duId] = {
        name:     idxName     >= 0 ? String(row[idxName]     || '').trim() : '',
        team:     idxTeam     >= 0 ? String(row[idxTeam]     || '').trim() : '',
        account:  idxAccount  >= 0 ? String(row[idxAccount]  || '').trim() : '',
        clockIn:  idxClockIn  >= 0 ? String(row[idxClockIn]  || '').trim() : '',
        clockOut: clockOut
      };
    }

    return ContentService
      .createTextOutput(JSON.stringify(result))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ error: err.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// Serve tracker data from TRACKER_SHEET (reads ALL rows, unaffected by filters)
function serveTracker() {
  var sheet = getTrackerSheet();
  var lastRow = sheet.getLastRow();
  var lastCol = sheet.getLastColumn();

  if (lastRow < 2) {
    return ContentService
      .createTextOutput(JSON.stringify([]))
      .setMimeType(ContentService.MimeType.JSON);
  }

  // Read all data (skip header row 1)
  var data = sheet.getRange(2, 1, lastRow - 1, lastCol).getValues();

  var tracker = [];
  for (var i = 0; i < data.length; i++) {
    var row = data[i];
    var duId = String(row[COL.duId] || '').trim();
    var siteName = String(row[COL.siteName] || '').trim();

    // Skip rows with no identifiable site
    if (!duId && !siteName) continue;

    tracker.push({
      batchRfi:    String(row[COL.batchRfi]    || '').trim(),
      batchCsg:    String(row[COL.batchCsg]    || '').trim(),
      tpId:        String(row[COL.tpSiteId]   || '').trim(),
      surgeId:     String(row[COL.surgeId]     || '').trim(),
      duId:        duId,
      siteName:    siteName,
      subcontPlan: String(row[COL.subcontPlan] || '').trim(),
      subcontCsg:  String(row[COL.subcontCsg]  || '').trim(),
      gap:         String(row[COL.gap]         || '').trim(),
      remark:      String(row[COL.remark]      || '').trim(),
      fo:         String(row[COL.fo]          || '').trim(),
      city:        String(row[COL.city]        || '').trim(),
      tp:          String(row[COL.tp]          || '').trim(),
      km:          String(row[COL.km]          || '').trim(),
      peOwner:     String(row[COL.peOwner]     || '').trim(),
      lat:         parseNum(row[COL.lat]),
      lng:         parseNum(row[COL.lng]),
      ring:        String(row[COL.ring]        || '').trim(),
      seq:         String(row[COL.seq]         || '').trim(),
      csgStatus:   String(row[COL.csgStatus]   || '').trim(),
      priority:    String(row[COL.priority]    || '').trim()
    });
  }

  return ContentService
    .createTextOutput(JSON.stringify(tracker))
    .setMimeType(ContentService.MimeType.JSON);
}

// Serve full forward log (for History view)
function serveForwardLog() {
  var sheet = getForwardSheet();
  var lastRow = sheet.getLastRow();

  if (lastRow < 2) {
    return ContentService
      .createTextOutput(JSON.stringify([]))
      .setMimeType(ContentService.MimeType.JSON);
  }

  var data = sheet.getRange(2, 1, lastRow - 1, 8).getValues();
  var log = [];

  for (var i = 0; i < data.length; i++) {
    var row = data[i];
    var ts = row[FWD_COL.timestamp];

    // Parse timestamp
    var timestamp = null;
    if (ts instanceof Date) {
      timestamp = ts.getTime();
    } else if (typeof ts === 'number') {
      timestamp = ts;  // Serial date number
    } else if (typeof ts === 'string' && ts) {
      var d = new Date(ts);
      if (!isNaN(d.getTime())) timestamp = d.getTime();
    }

    log.push({
      teamLeader: String(row[FWD_COL.teamLeader] || '').trim(),
      duId:       String(row[FWD_COL.duId]       || '').trim(),
      siteName:   String(row[FWD_COL.siteName]   || '').trim(),
      message:    String(row[FWD_COL.message]    || '').trim(),
      remark:     String(row[FWD_COL.remark]     || '').trim(),
      ts:         timestamp,
      state:      parseState(row[FWD_COL.state]),
      isLatest:   row[FWD_COL.isLatest] === '' ? true :
                  (String(row[FWD_COL.isLatest] || '').toLowerCase() === 'true')
    });
  }

  return ContentService
    .createTextOutput(JSON.stringify(log))
    .setMimeType(ContentService.MimeType.JSON);
}

// Serve lightweight summary for Leaderboard (no message/state fields)
function serveSummary() {
  var sheet = getForwardSheet();
  var lastRow = sheet.getLastRow();

  if (lastRow < 2) {
    return ContentService
      .createTextOutput(JSON.stringify([]))
      .setMimeType(ContentService.MimeType.JSON);
  }

  var data = sheet.getRange(2, 1, lastRow - 1, 8).getValues();
  var log = [];

  for (var i = 0; i < data.length; i++) {
    var row = data[i];
    var ts = row[FWD_COL.timestamp];
    var timestamp = null;
    if (ts instanceof Date) {
      timestamp = ts.getTime();
    } else if (typeof ts === 'number') {
      timestamp = ts;
    }

    log.push({
      teamLeader: String(row[FWD_COL.teamLeader] || '').trim(),
      duId:       String(row[FWD_COL.duId]       || '').trim(),
      siteName:   String(row[FWD_COL.siteName]   || '').trim(),
      remark:     String(row[FWD_COL.remark]     || '').trim(),
      ts:         timestamp,
      isLatest:   row[FWD_COL.isLatest] === '' ? true :
                  (String(row[FWD_COL.isLatest] || '').toLowerCase() === 'true')
    });
  }

  return ContentService
    .createTextOutput(JSON.stringify(log))
    .setMimeType(ContentService.MimeType.JSON);
}

// Parse state JSON from a cell value
function parseState(raw) {
  if (!raw) return null;
  try {
    if (typeof raw === 'object') return raw;
    return JSON.parse(String(raw));
  } catch (e) {
    return null;
  }
}

// --- DO POST: handle form submissions, deletes ---
function doPost(e) {
  try {
    var raw = e.postData ? e.postData.contents : e.parameter;
    var body = (typeof raw === 'string') ? JSON.parse(raw) : raw;

    // Handle delete request
    if (body.type === 'delete') {
      return handleDelete(body);
    }

    // Handle bulk edit of tracker rows
    if (body.type === 'trackerBulkEdit') {
      return handleBulkEdit(body);
    }

    // Handle daily plan save (per-date replace)
    if (body.type === 'planSave') {
      return handlePlanSave(body);
    }

    // Handle setting the surge's active date (broadcast to subcont clients)
    if (body.type === 'setActiveDate') {
      return handleSetActiveDate(body);
    }

    // Handle plan roster delta save (surge-managed additions/removals)
    // DEPRECATED: superseded by planRosterCommand.
    if (body.type === 'planRosterSave') {
      return handlePlanRosterSave(body);
    }

    // Handle a plan roster command (the single roster write path)
    if (body.type === 'planRosterCommand') {
      return handlePlanRosterCommand(body);
    }

    // Handle plan roster field edits (account/subcon/teamType/resourceRemark propagation)
    // DEPRECATED: superseded by planRosterCommand.
    if (body.type === 'planRosterEdit') {
      return handlePlanRosterEdit(body);
    }

    // Handle removal of a stored plan roster field edit
    // DEPRECATED: superseded by planRosterCommand.
    if (body.type === 'planRosterEditDelete') {
      return handlePlanRosterEditDelete(body);
    }

    // Handle forward/submission
    return handleForward(body);

  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ status: 'error', message: err.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// Record a forward/submission entry
function handleForward(body) {
  var teamLeader = String(body.teamLeader || '').trim() || 'Unknown';
  var duId       = String(body.duId       || '').trim();
  var siteName   = String(body.siteName   || '').trim();
  var message    = String(body.message    || '').trim();
  var remark     = String(body.remark     || '').trim();
  var ts         = body.ts || Date.now();
  var state      = body.state || null;

  var sheet = getForwardSheet();
  var lastRow = sheet.getLastRow() + 1;

  // Normalize DU ID for comparison (strip _NEW, _CSG suffixes)
  var normDuId = duId.toUpperCase().replace(/_NEW$/, '').replace(/_CSG$/, '');

  // Mark all previous entries with the same normalized DU ID as NOT latest
  if (lastRow > 1) {
    var existingData = sheet.getRange(2, 1, lastRow - 2, 8).getValues();
    var updates = [];
    for (var i = 0; i < existingData.length; i++) {
      var existingDuId = String(existingData[i][FWD_COL.duId] || '').trim()
        .toUpperCase().replace(/_NEW$/, '').replace(/_CSG$/, '');
      if (existingDuId === normDuId && String(existingData[i][FWD_COL.isLatest] || '').toLowerCase() !== 'false') {
        updates.push({ row: i + 2, col: FWD_COL.isLatest + 1 });
      }
    }
    // Batch update: set isLatest = FALSE for previous entries
    for (var u = 0; u < updates.length; u++) {
      sheet.getRange(updates[u].row, updates[u].col).setValue('FALSE');
    }
  }

  // Append the new entry
  var stateJson = (state && typeof state === 'object') ? JSON.stringify(state) : '';
  sheet.getRange(lastRow, 1, 1, 8).setValues([[
    new Date(ts),
    teamLeader,
    duId,
    siteName,
    message,
    remark,
    stateJson,
    'TRUE'
  ]]);

  return ContentService
    .createTextOutput(JSON.stringify({ status: 'ok', row: lastRow }))
    .setMimeType(ContentService.MimeType.JSON);
}

// Delete entries by DU ID
function handleDelete(body) {
  var duId    = String(body.duId    || '').trim();
  var pass    = String(body.password || '').trim();
  var DEL_PASS = 'latif123';  // <-- change this to your preferred delete password

  if (pass !== DEL_PASS) {
    return ContentService
      .createTextOutput(JSON.stringify({ status: 'error', message: 'Wrong password' }))
      .setMimeType(ContentService.MimeType.JSON);
  }

  var sheet = getForwardSheet();
  var lastRow = sheet.getLastRow();

  if (lastRow < 2) {
    return ContentService
      .createTextOutput(JSON.stringify({ status: 'ok', deleted: 0 }))
      .setMimeType(ContentService.MimeType.JSON);
  }

  var data = sheet.getRange(2, 1, lastRow - 1, 8).getValues();
  var rowsToDelete = [];

  for (var i = 0; i < data.length; i++) {
    var existingDuId = String(data[i][FWD_COL.duId] || '').trim();
    if (existingDuId === duId) {
      rowsToDelete.push(i + 2);  // Sheet rows are 1-based, row 1 is header
    }
  }

  // Delete rows in reverse order to preserve row indices
  for (var d = rowsToDelete.length - 1; d >= 0; d--) {
    sheet.deleteRow(rowsToDelete[d]);
  }

  return ContentService
    .createTextOutput(JSON.stringify({ status: 'ok', deleted: rowsToDelete.length }))
    .setMimeType(ContentService.MimeType.JSON);
}

// --- Bulk edit tracker rows (GAP, Remark, etc.) ---
function handleBulkEdit(body) {
  var updates = body.updates || [];
  if (!updates.length) {
    return ContentService
      .createTextOutput(JSON.stringify({ status: 'ok', updated: 0 }))
      .setMimeType(ContentService.MimeType.JSON);
  }

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(TRACKER_SHEET);
  if (!sheet) {
    return ContentService
      .createTextOutput(JSON.stringify({ status: 'error', message: 'Tracker sheet not found' }))
      .setMimeType(ContentService.MimeType.JSON);
  }

  // Build column name to index map from row 1
  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  var colMap = {};
  for (var ci = 0; ci < headers.length; ci++) {
    colMap[String(headers[ci] || '').trim()] = ci + 1; // 1-based
  }

  // Group updates by row number
  var rowUpdates = {};
  for (var ui = 0; ui < updates.length; ui++) {
    var u = updates[ui];
    if (!rowUpdates[u.row]) rowUpdates[u.row] = {};
    rowUpdates[u.row][u.field] = u.value;
  }

  var rowNums = Object.keys(rowUpdates).sort(function(a, b) { return Number(b) - Number(a); }); // desc to avoid row shift
  var updatedCount = 0;

  for (var ri = 0; ri < rowNums.length; ri++) {
    var rowNum = Number(rowNums[ri]);
    var fields = rowUpdates[rowNums[ri]];
    var fieldNames = Object.keys(fields);

    for (var fi = 0; fi < fieldNames.length; fi++) {
      var field = fieldNames[fi];
      var colNum = colMap[field];
      if (!colNum) continue;
      sheet.getRange(rowNum, colNum).setValue(fields[field]);
      updatedCount++;
    }
  }

  return ContentService
    .createTextOutput(JSON.stringify({ status: 'ok', updated: updatedCount }))
    .setMimeType(ContentService.MimeType.JSON);
}

// ============================================================
// DAILY PLAN TAB
// Sheet name: "plan"
// Columns: date, account, subcon, teamType, resourceRemark,
//          duId, siteName, activityRemark, dailyPlanActivity, result,
//          rowIdx (deprecated), rosterId
// result is the 10th column (0-based index 9) — user-entered daily result text.
// rowIdx is the 11th column (0-based index 10) — deprecated positional key.
// rosterId is the 12th column (0-based index 11) — the STABLE join key
// between a roster member and their plan rows. Matching on rosterId (instead
// of the account name or a positional index) is what makes renames and roster
// re-orderings safe: a row can never shift onto the wrong person.
// ============================================================

var PLAN_COL_ROSTER_ID = 12; // 1-based column index of rosterId
var PLAN_NUM_COLS = 12;

// --- HELPER: get or create the Plan sheet ---
function getOrCreatePlanSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('plan');
  if (!sheet) {
    sheet = ss.insertSheet('plan');
    sheet.getRange(1, 1, 1, PLAN_NUM_COLS).setValues([[
      'date', 'account', 'subcon', 'teamType', 'resourceRemark',
      'duId', 'siteName', 'activityRemark', 'dailyPlanActivity', 'result',
      'rowIdx', 'rosterId'
    ]]);
    sheet.setFrozenRows(1);
  } else {
    // Existing sheet: make sure the rosterId header is present. The column is
    // APPENDED (position 12) so no existing column shifts.
    var header = sheet.getRange(1, PLAN_COL_ROSTER_ID).getValue();
    if (header !== 'rosterId') {
      sheet.getRange(1, PLAN_COL_ROSTER_ID).setValue('rosterId');
    }
  }
  // Force column A (date) to plain text so Sheet doesn't auto-convert
  // 'YYYY-MM-DD' strings into Date objects (which break the strict-equality
  // date filter in servePlan / servePlanDates).
  var maxRows = Math.max(sheet.getLastRow(), sheet.getMaxRows());
  if (maxRows > 1) {
    sheet.getRange(2, 1, maxRows - 1, 1).setNumberFormat('@');
  } else {
    sheet.getRange('A2:A').setNumberFormat('@');
  }
  return sheet;
}

// Serve rows for a single date from the plan sheet.
// date is YYYY-MM-DD; if blank, returns all rows.
function servePlan(date) {
  var sh = getOrCreatePlanSheet();
  var last = sh.getLastRow();
  if (last < 2) {
    return ContentService
      .createTextOutput(JSON.stringify([]))
      .setMimeType(ContentService.MimeType.JSON);
  }
  var values = sh.getRange(2, 1, last - 1, PLAN_NUM_COLS).getValues();
  var want = date ? String(date) : '';
  var rows = [];
  for (var i = 0; i < values.length; i++) {
    var r = values[i];
    // Apps Script auto-detects the 'date' column as a Date type — normalize
    // back to YYYY-MM-DD so the strict equality filter matches the client input.
    var dateCell = _planDateCellToString(r[0]);
    if (want && dateCell !== want) continue;
    rows.push({
      date:             dateCell,
      account:          String(r[1] == null ? '' : r[1]).trim(),
      subcon:           String(r[2] == null ? '' : r[2]).trim(),
      teamType:         String(r[3] == null ? '' : r[3]).trim(),
      resourceRemark:   String(r[4] == null ? '' : r[4]).trim(),
      duId:             String(r[5] == null ? '' : r[5]).trim(),
      siteName:         String(r[6] == null ? '' : r[6]).trim(),
      activityRemark:   String(r[7] == null ? '' : r[7]).trim(),
      dailyPlanActivity:String(r[8] == null ? '' : r[8]).trim(),
      result:           String(r[9] == null ? '' : r[9]),
      rowIdx:           r[10] === '' || r[10] === null || isNaN(Number(r[10])) ? -1 : Number(r[10]),
      rosterId:         String(r[11] == null ? '' : r[11]).trim()
    });
  }
  return ContentService
    .createTextOutput(JSON.stringify(rows))
    .setMimeType(ContentService.MimeType.JSON);
}

// Normalize a plan-date cell value to 'YYYY-MM-DD' string. Returns '' for empty.
// Sheets auto-detects date-formatted strings into Date objects on write/read;
// getValues() then returns Date, whose String() is the long form, not 'YYYY-MM-DD'.
function _planDateCellToString(v) {
  if (v === '' || v === null || v === undefined) return '';
  if (Object.prototype.toString.call(v) === '[object Date]' && !isNaN(v.getTime())) {
    return Utilities.formatDate(v, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  }
  return String(v);
}

// Serve a sorted-desc unique list of dates that have at least one plan row.
function servePlanDates() {
  var sh = getOrCreatePlanSheet();
  var last = sh.getLastRow();
  if (last < 2) {
    return ContentService
      .createTextOutput(JSON.stringify([]))
      .setMimeType(ContentService.MimeType.JSON);
  }
  var data = sh.getRange(2, 1, last - 1, 11).getValues();
  var seen = {};
  var dates = [];
  for (var i = 0; i < data.length; i++) {
    var d = _planDateCellToString(data[i][0]);
    if (d && !seen[d]) { seen[d] = true; dates.push(d); }
  }
  dates.sort();
  dates.reverse();
  return ContentService
    .createTextOutput(JSON.stringify(dates))
    .setMimeType(ContentService.MimeType.JSON);
}

// Save all rows for a given date.
//
// Merge strategy: rows are keyed by `account` (a stable, unique identity), NOT
// by the positional `rowIdx`. Positional keys break whenever the roster
// changes (e.g. a member is deleted and every later index shifts), which
// silently reassigns one person's plan data to another person.
//
// Write strategy: existing rows for the date are updated IN PLACE and new rows
// are appended. The previous "clear the date's rows, then rewrite the whole
// block" approach could permanently destroy a date's data if the second write
// failed (quota, timeout, dropped connection) after the clear had committed.
//
// A script lock serializes the read-modify-write so two subconts saving the
// same date concurrently cannot clobber each other.
function handlePlanSave(body) {
  // --- Input validation ---
  var date = String(body.date || '').trim();
  if (!/^\d{4}-\d{2}-\d{2}$/.test(date)) {
    return ContentService
      .createTextOutput(JSON.stringify({ status: 'error', message: 'Invalid date (YYYY-MM-DD)' }))
      .setMimeType(ContentService.MimeType.JSON);
  }
  if (!body.rows || !Array.isArray(body.rows)) {
    return ContentService
      .createTextOutput(JSON.stringify({ status: 'error', message: 'rows must be an array' }))
      .setMimeType(ContentService.MimeType.JSON);
  }
  var rows = body.rows;
  // Hard cap so a malformed or hostile payload cannot exhaust the sheet or
  // trip the execution-time limit mid-write.
  var MAX_PLAN_ROWS = 2000;
  if (rows.length > MAX_PLAN_ROWS) {
    return ContentService
      .createTextOutput(JSON.stringify({ status: 'error', message: 'Too many rows (max ' + MAX_PLAN_ROWS + ')' }))
      .setMimeType(ContentService.MimeType.JSON);
  }
  // Accounts the surge has removed from the roster. Their rows are deleted
  // from the date instead of being preserved forever.
  var prune = [];
  if (body.pruneAccounts && Array.isArray(body.pruneAccounts)) {
    for (var p0 = 0; p0 < body.pruneAccounts.length; p0++) {
      var pa = String(body.pruneAccounts[p0] || '').trim();
      if (pa) prune.push(pa);
    }
  }
  var pruneSet = {};
  for (var p1 = 0; p1 < prune.length; p1++) pruneSet[prune[p1]] = true;
  // Roster ids the surge has removed. Rows belonging to these members are
  // deleted from the date as well — matching by the stable id so a rename
  // cannot leave an orphaned row behind.
  var pruneIdSet = {};
  if (body.pruneRosterIds && Array.isArray(body.pruneRosterIds)) {
    for (var pi = 0; pi < body.pruneRosterIds.length; pi++) {
      var pid = String(body.pruneRosterIds[pi] || '').trim();
      if (pid) pruneIdSet[pid] = true;
    }
  }

  // Serialise concurrent saves. Without this, two users saving the same date
  // can both read the pre-save state and the second write drops the first's rows.
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(30000);
  } catch (lockErr) {
    return ContentService
      .createTextOutput(JSON.stringify({ status: 'error', message: 'Server busy, please retry' }))
      .setMimeType(ContentService.MimeType.JSON);
  }

  try {
    var sh = getOrCreatePlanSheet();
    var last = sh.getLastRow();

    // Flush any pending writes before reading — makes read-modify-write more atomic
    // under concurrent access (one user saves while another is mid-save).
    SpreadsheetApp.flush();

    // --- 1. Read all existing data rows ---
    var existing = [];
    var existingRowNums = []; // parallel: sheet row numbers for each existing entry
    if (last > 1) {
      var vals = sh.getRange(2, 1, last - 1, PLAN_NUM_COLS).getValues();
      for (var i = 0; i < vals.length; i++) {
        existingRowNums.push(i + 2); // sheet row number (1-based)
        existing.push(vals[i]);
      }
    }

    // --- 2. Build the incoming map keyed by rosterId ---
    // rosterId is the stable join key. The legacy `account` name is kept as a
    // fallback for rows written before the migration (no rosterId yet) so a
    // first save after the upgrade still attaches to the right person — and it
    // writes the resolved rosterId back, so the next save is clean.
    var incomingById = Object.create(null);
    var incomingByAccountLc = Object.create(null); // legacy fallback
    var incomingOrder = [];
    var warnings = [];
    for (var j = 0; j < rows.length; j++) {
      var r = rows[j];
      if (!r || typeof r !== 'object') continue; // skip malformed entries
      var account = String(r.account == null ? '' : r.account).trim();
      var rosterId = String(r.rosterId == null ? '' : r.rosterId).trim();
      if (!account && !rosterId) continue; // rows without an identity can't be matched
      var idx = Number(r.rowIdx);
      var rowVals = [
        date,
        account,
        String(r.subcon            == null ? '' : r.subcon),
        String(r.teamType          == null ? '' : r.teamType),
        String(r.resourceRemark    == null ? '' : r.resourceRemark),
        String(r.duId              == null ? '' : r.duId),
        String(r.siteName          == null ? '' : r.siteName),
        String(r.activityRemark    == null ? '' : r.activityRemark),
        String(r.dailyPlanActivity == null ? '' : r.dailyPlanActivity),
        String(r.result            == null ? '' : r.result),
        isNaN(idx) ? -1 : idx,
        rosterId
      ];
      if (rosterId) {
        if (!Object.prototype.hasOwnProperty.call(incomingById, rosterId)) incomingOrder.push(rosterId);
        incomingById[rosterId] = rowVals;
      }
      if (account) {
        var lcAcct = account.toLowerCase();
        if (!Object.prototype.hasOwnProperty.call(incomingByAccountLc, lcAcct)) {
          incomingByAccountLc[lcAcct] = { rowVals: rowVals, rosterId: rosterId };
        }
      }
    }

    // --- 3. Update existing rows for this date in place ---
    var seenIds = Object.create(null);
    var seenAccountLc = Object.create(null);
    var updated = 0;
    var unchanged = 0;
    var deleteRowNums = [];
    for (var k = 0; k < existing.length; k++) {
      var dateCellK = _planDateCellToString(existing[k][0]);
      if (dateCellK !== date) continue; // different date — leave untouched
      var existingAccount = String(existing[k][1] == null ? '' : existing[k][1]).trim();
      var existingId = String(existing[k][11] == null ? '' : existing[k][11]).trim();
      // Explicitly pruned account (removed from the roster) — drop the row.
      if (existingAccount && pruneSet[existingAccount]) {
        deleteRowNums.push(existingRowNums[k]);
        continue;
      }
      // Explicitly pruned roster id — drop the row. This is the reliable path
      // (the id survives renames); the account-name path above is the legacy
      // fallback for clients that have not migrated yet.
      if (existingId && pruneIdSet[existingId]) {
        deleteRowNums.push(existingRowNums[k]);
        continue;
      }
      // Match by rosterId first — the stable key.
      var matchVals = null;
      var matchId = '';
      if (existingId && Object.prototype.hasOwnProperty.call(incomingById, existingId) && !seenIds[existingId]) {
        matchVals = incomingById[existingId];
        matchId = existingId;
      } else if (existingAccount) {
        // Legacy fallback: this sheet row predates the rosterId column, so
        // match it by account name (case-insensitively).
        var lcExisting = existingAccount.toLowerCase();
        if (Object.prototype.hasOwnProperty.call(incomingByAccountLc, lcExisting) && !seenAccountLc[lcExisting]) {
          matchVals = incomingByAccountLc[lcExisting].rowVals;
          matchId = incomingByAccountLc[lcExisting].rosterId;
          seenAccountLc[lcExisting] = true;
          if (!matchId) {
            warnings.push('Row for "' + existingAccount + '" has no roster id yet — matched by name');
          }
        }
      }
      if (matchVals) {
        // Compare against the current row so we can report "unchanged" honestly.
        var changed = false;
        for (var cc = 1; cc < PLAN_NUM_COLS; cc++) {
          if (cc === 10) continue; // rowIdx is informational only
          var cur = existing[k][cc] === null || existing[k][cc] === undefined ? '' : String(existing[k][cc]);
          var nw = matchVals[cc] === null || matchVals[cc] === undefined ? '' : String(matchVals[cc]);
          if (cur !== nw) { changed = true; break; }
        }
        // Always carry the rosterId into the sheet row (column 12).
        sh.getRange(existingRowNums[k], 1, 1, PLAN_NUM_COLS).setValues([matchVals]);
        if (matchId) seenIds[matchId] = true;
        if (changed) updated++; else unchanged++;
      }
      // No incoming row matches this row — leave the existing row as-is
      // (another subcont's row that wasn't in this submission).
    }

    // --- 4. Append genuinely new rows after the last used row ---
    var newRows = [];
    for (var m = 0; m < incomingOrder.length; m++) {
      var idM = incomingOrder[m];
      if (!seenIds[idM]) {
        newRows.push(incomingById[idM]);
        seenIds[idM] = true;
      }
    }
    // Legacy rows that arrived without a rosterId and did not match any
    // existing sheet row still need appending.
    var seenAcctKeys = Object.keys(incomingByAccountLc);
    for (var la = 0; la < seenAcctKeys.length; la++) {
      var lcK = seenAcctKeys[la];
      if (seenAccountLc[lcK]) continue;
      var entry = incomingByAccountLc[lcK];
      if (entry.rosterId && seenIds[entry.rosterId]) continue;
      newRows.push(entry.rowVals);
      if (entry.rosterId) seenIds[entry.rosterId] = true;
      seenAccountLc[lcK] = true;
    }
    var appended = 0;
    if (newRows.length > 0) {
      // Only append into rows that are actually free (the sheet may have
      // trailing blank rows). Find the first free row at or after the last
      // populated one.
      var writeAt = last + 1;
      for (var s = 0; s < existing.length; s++) {
        var isEmpty = true;
        for (var c = 0; c < PLAN_NUM_COLS; c++) {
          var cv = existing[s][c];
          if (cv !== '' && cv !== null && cv !== undefined) { isEmpty = false; break; }
        }
        if (!isEmpty) writeAt = Math.max(writeAt, existingRowNums[s] + 1);
      }
      // Ensure we don't overwrite rows belonging to other dates.
      sh.getRange(writeAt, 1, newRows.length, PLAN_NUM_COLS).setValues(newRows);
      appended = newRows.length;
    }

    // --- 5. Delete pruned rows (from the bottom up so indices stay valid) ---
    if (deleteRowNums.length > 0) {
      deleteRowNums.sort(function(a, b) { return b - a; });
      for (var d = 0; d < deleteRowNums.length; d++) {
        sh.deleteRow(deleteRowNums[d]);
      }
    }

    return ContentService
      .createTextOutput(JSON.stringify({
        status:    'ok',
        written:   updated + appended,
        updated:   updated,
        appended:  appended,
        unchanged: unchanged,
        deleted:   deleteRowNums.length,
        warnings:  warnings
      }))
      .setMimeType(ContentService.MimeType.JSON);
  } finally {
    lock.releaseLock();
  }
}

// Serve the surge's active plan date (script-wide shared property).
// Returns { date: 'YYYY-MM-DD' } or { date: '' } if unset.
function serveActiveDate() {
  var d = PropertiesService.getScriptProperties().getProperty('planActiveDate') || '';
  return ContentService
    .createTextOutput(JSON.stringify({ date: d }))
    .setMimeType(ContentService.MimeType.JSON);
}

// Set the surge's active plan date. Stored in script properties (survives redeploys).
// No auth — matches the existing trust model (surge check is client-side).
function handleSetActiveDate(body) {
  var date = String(body.date || '').trim();
  if (!/^\d{4}-\d{2}-\d{2}$/.test(date)) {
    return ContentService
      .createTextOutput(JSON.stringify({ status: 'error', message: 'Invalid date (YYYY-MM-DD)' }))
      .setMimeType(ContentService.MimeType.JSON);
  }
  PropertiesService.getScriptProperties().setProperty('planActiveDate', date);
  return ContentService
    .createTextOutput(JSON.stringify({ status: 'ok', date: date }))
    .setMimeType(ContentService.MimeType.JSON);
}

// Serve the surge-managed plan roster delta. Layered on top of the
// hardcoded base roster in the client. Stored in script properties
// (survives redeploys). Returns { added: [...], removed: [...] }.
//
// DEPRECATED: superseded by servePlanRosterDoc(). Kept for one release so an
// older client build does not break against a newer backend.
function servePlanRoster() {
  var raw = PropertiesService.getScriptProperties().getProperty('planRoster') || '';
  var delta;
  try {
    delta = raw ? JSON.parse(raw) : { added: [], removed: [] };
    if (!Array.isArray(delta.added))   delta.added   = [];
    if (!Array.isArray(delta.removed)) delta.removed = [];
  } catch (e) {
    delta = { added: [], removed: [] };
  }
  return ContentService
    .createTextOutput(JSON.stringify({ added: delta.added, removed: delta.removed }))
    .setMimeType(ContentService.MimeType.JSON);
}

// ============================================================
// PLAN ROSTER DOCUMENT (single source of truth)
//
// The roster used to live in three places at once — a hardcoded PLAN_ROSTER
// in the client, a `planRoster` delta, and a `planRosterEdits` override map —
// and the client merged them inconsistently (by name in some places, by array
// index in others). That produced the "Account missing" and "Add team doesn't
// sync" symptoms.
//
// Now the server owns ONE document, stored in the `planRosterDoc` script
// property:
//   { version: <int>, accounts: [ { id, account, subcon, teamType,
//                                   resourceRemark, removedAt } ] }
//
// `id` is a stable, immutable identifier. It is the join key between the
// roster and the plan sheet rows, so renaming a member never orphans their
// plan history and re-ordering the roster never shifts a row onto the wrong
// person. `version` increments on every mutation so a client holding a stale
// roster can detect it and re-apply its command.
//
// The legacy `planRosterEdits` map is still readable (see servePlanRosterEdits)
// so an old client build keeps working for one release, but nothing writes to
// it any more.
// ============================================================

var PLAN_ROSTER_PROP = 'planRosterDoc';

// Canonical subcont keys. Casing is not significant — the user confirmed
// `Poca` and `POCA` are the same team — so every subcont is stored uppercase.
var PLAN_SUBCONT_KEYS = [
  'ADW', 'CAA', 'CCSI', 'INDOHR', 'INNOVIS', 'INTISEL', 'MKU', 'MPJ',
  'NIM', 'PAB', 'POCA', 'TRITAMA', 'YPIT', 'ZMG'
];

// Resolve any free-text subcont value to its canonical uppercase key.
// Returns '' when the value matches no known subcont.
function _planCanonicalSubcont(value) {
  var v = String(value == null ? '' : value).trim();
  if (!v) return '';
  var up = v.toUpperCase();
  for (var i = 0; i < PLAN_SUBCONT_KEYS.length; i++) {
    if (PLAN_SUBCONT_KEYS[i] === up) return up;
  }
  // Fallback: strip trailing digits so "INTISEL12" resolves to "INTISEL".
  var stripped = up.replace(/\d+$/, '');
  for (var j = 0; j < PLAN_SUBCONT_KEYS.length; j++) {
    if (PLAN_SUBCONT_KEYS[j] === stripped) return stripped;
  }
  return '';
}

// Generate a stable, collision-resistant roster id.
// Uses the account name so ids are reproducible for a given account, but
// salted with a random suffix so two members can never share an id even if
// their names differ only by case.
function _planNewRosterId(account) {
  var base = String(account || '').toLowerCase().replace(/[^a-z0-9]+/g, '');
  if (!base) base = 'acct';
  var rand = Math.random().toString(36).substring(2, 8);
  return base.substring(0, 16) + '_' + rand;
}

// Read the roster document, normalising its shape. Returns
// { version: 0, accounts: [] } when unset or corrupt.
function _planReadRosterDoc() {
  var raw = PropertiesService.getScriptProperties().getProperty(PLAN_ROSTER_PROP) || '';
  var doc;
  try {
    doc = raw ? JSON.parse(raw) : null;
  } catch (e) {
    doc = null;
  }
  if (!doc || typeof doc !== 'object') doc = {};
  if (!Array.isArray(doc.accounts)) doc.accounts = [];
  var version = Number(doc.version);
  if (isNaN(version) || version < 0) version = 0;
  doc.version = version;
  // Drop entries without an id or account, and make sure every account has
  // one so downstream code can rely on the shape.
  var clean = [];
  for (var i = 0; i < doc.accounts.length; i++) {
    var a = doc.accounts[i];
    if (!a || typeof a !== 'object') continue;
    var acct = String(a.account == null ? '' : a.account).trim();
    if (!acct) continue;
    var id = String(a.id == null ? '' : a.id).trim();
    if (!id) id = _planNewRosterId(acct);
    clean.push({
      id:              id,
      account:         acct,
      subcon:          _planCanonicalSubcont(a.subcon) || String(a.subcon || '').trim().toUpperCase(),
      teamType:        String(a.teamType == null ? '' : a.teamType).trim(),
      resourceRemark:  String(a.resourceRemark == null ? '' : a.resourceRemark).trim(),
      removedAt:       a.removedAt ? String(a.removedAt) : ''
    });
  }
  doc.accounts = clean;
  return doc;
}

// The base roster, used ONLY to seed an empty server roster on first run.
// After that the server roster is the sole source of truth and this list is
// never merged at render time.
// Keep in sync with the site-acceptance team list.
var PLAN_ROSTER_SEED = [
  { account: 'intisel_tl_rahmat_32051919',     subcon: 'INTISEL',  teamType: 'WL',  resourceRemark: 'Dedicated/Shared' },
  { account: 'intisel_tl_deni_32053528',       subcon: 'INTISEL',  teamType: 'WL',  resourceRemark: 'Dedicated/Shared' },
  { account: 'intisel_tl_kusuma_32170422',     subcon: 'INTISEL',  teamType: 'WL',  resourceRemark: 'Dedicated/Shared' },
  { account: 'nim_tl_gustira_32053405',        subcon: 'NIM',      teamType: 'WL',  resourceRemark: 'Dedicated/Shared' },
  { account: 'tal_tl_acep_32044619',          subcon: 'TRITAMA',  teamType: 'WL',  resourceRemark: 'Dedicated/Shared' },
  { account: 'intisel_tl_mulyana_32732321',    subcon: 'INTISEL',  teamType: 'WL',  resourceRemark: 'Dedicated/Shared' },
  { account: 'andri_poca_wj',                 subcon: 'POCA',     teamType: 'WL',  resourceRemark: 'Dedicated/Shared' },
  { account: 'tal_tl_cahya_32041711',          subcon: 'TRITAMA',  teamType: 'WL',  resourceRemark: 'Dedicated/Shared' },
  { account: 'pab_tl_fajarfm_33091410',        subcon: 'PAB',      teamType: 'WL',  resourceRemark: 'Dedicated/Shared' },
  { account: 'intisel_tl_nazar_32780821',      subcon: 'INTISEL',  teamType: 'WL',  resourceRemark: 'Dedicated/Shared' },
  { account: 'nim_tl_dindin_32043209',         subcon: 'NIM',      teamType: 'WL',  resourceRemark: 'Dedicated/Shared' },
  { account: 'pab_tl_ricky_32062430',          subcon: 'PAB',      teamType: 'WL',  resourceRemark: 'Dedicated/Shared' },
  { account: 'pab_tl_fajar_31720219',          subcon: 'PAB',      teamType: 'WL',  resourceRemark: 'Dedicated/Shared' },
  { account: 'pab_mb_zhonaade_35191211',       subcon: 'PAB',      teamType: 'WL',  resourceRemark: 'Dedicated/Shared' },
  { account: 'pab_tl_agusferdi_33091450',      subcon: 'PAB',      teamType: 'WL',  resourceRemark: 'Dedicated/Shared' },
  { account: 'dedi_samsudin',                  subcon: 'ZMG',      teamType: 'WL',  resourceRemark: 'Dedicated/Shared' },
  { account: 'adw_tl_agung_32092512',          subcon: 'ADW',      teamType: 'CSG', resourceRemark: 'Dedicated/Shared' },
  { account: 'adw_tl_yudi_32092712',           subcon: 'ADW',      teamType: 'CSG', resourceRemark: 'Dedicated/Shared' },
  { account: 'pab_mb_perridarmawan_32731619',  subcon: 'PAB',      teamType: 'CSG', resourceRemark: 'Dedicated/Shared' },
  { account: 'poca_tl_yadi_32042903',          subcon: 'POCA',     teamType: 'WL',  resourceRemark: 'Dedicated/Shared' },
  { account: 'chr_jabo_yanaoktaviana',         subcon: 'INDOHR',   teamType: 'WL',  resourceRemark: 'Dedicated/Shared' },
  { account: 'inv_tl_sandi_32042901',          subcon: 'INNOVIS',  teamType: 'WL',  resourceRemark: 'Dedicated/Shared' },
  { account: 'caa_tl_reza.r_32020226',         subcon: 'CAA',      teamType: 'WL',  resourceRemark: 'Dedicated/Shared' }
];

// Seed the roster document from PLAN_ROSTER_SEED when it is empty. Returns
// true if a seed was written.
function _planSeedRosterIfEmpty() {
  var raw = PropertiesService.getScriptProperties().getProperty(PLAN_ROSTER_PROP) || '';
  if (raw) {
    // Already seeded (or already populated) — leave it alone.
    try {
      var parsed = JSON.parse(raw);
      if (parsed && Array.isArray(parsed.accounts) && parsed.accounts.length > 0) return false;
    } catch (e) { /* corrupt — reseed below */ }
  }
  var accounts = PLAN_ROSTER_SEED.map(function(r) {
    return {
      id:             _planNewRosterId(r.account),
      account:        r.account,
      subcon:         _planCanonicalSubcont(r.subcon),
      teamType:       r.teamType,
      resourceRemark: r.resourceRemark,
      removedAt:      ''
    };
  });
  _planWriteRosterDoc({ version: 1, accounts: accounts });
  return true;
}

// Write the roster document.
function _planWriteRosterDoc(doc) {
  PropertiesService.getScriptProperties().setProperty(PLAN_ROSTER_PROP, JSON.stringify(doc));
  return doc;
}

// Find an account entry by id (exact) or by account name (case-insensitive).
// Returns the entry, or null.
function _planFindAccount(doc, id, account) {
  var i;
  if (id) {
    for (i = 0; i < doc.accounts.length; i++) {
      if (doc.accounts[i].id === id) return doc.accounts[i];
    }
  }
  if (account) {
    var lc = String(account).trim().toLowerCase();
    for (i = 0; i < doc.accounts.length; i++) {
      if (String(doc.accounts[i].account).trim().toLowerCase() === lc) return doc.accounts[i];
    }
  }
  return null;
}

// Serve the roster document. Always returns the full authoritative roster so
// the client can replace its in-memory copy wholesale.
function servePlanRosterDoc() {
  // Seed from the base roster the first time only. After that the stored
  // document is the single source of truth.
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    _planSeedRosterIfEmpty();
  } catch (e) {
    // Seeding is best-effort — a busy lock must not block the read.
  } finally {
    try { lock.releaseLock(); } catch (e2) {}
  }
  var doc = _planReadRosterDoc();
  return ContentService
    .createTextOutput(JSON.stringify({
      status:   'ok',
      version:  doc.version,
      accounts: doc.accounts
    }))
    .setMimeType(ContentService.MimeType.JSON);
}

// Apply one roster command and return the full resulting roster.
//
// ops:
//   upsert      — add a member, or update an existing member's fields
//   remove      — soft-remove a member (sets removedAt)
//   rename      — change a member's display account name
//   setSubcont  — move a member to another team
//
// The client sends its known `version`. If the server's version has moved on
// the command is NOT applied and { status:'stale', roster } is returned so the
// client can re-apply against the fresh roster.
function handlePlanRosterCommand(body) {
  var op = String(body.op || '').trim();
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(30000);
  } catch (lockErr) {
    return ContentService
      .createTextOutput(JSON.stringify({ status: 'error', message: 'Server busy, please retry' }))
      .setMimeType(ContentService.MimeType.JSON);
  }
  try {
    var doc = _planReadRosterDoc();
    var clientVersion = Number(body.version);
    if (!isNaN(clientVersion) && clientVersion !== doc.version) {
      // Another user changed the roster since this client last read it.
      // Do not apply — hand back the current roster so the client can retry.
      return ContentService
        .createTextOutput(JSON.stringify({
          status: 'stale',
          version: doc.version,
          roster: doc.accounts,
          message: 'Roster changed on the server — please retry'
        }))
        .setMimeType(ContentService.MimeType.JSON);
    }

    var id = String(body.id || '').trim();
    var account = String(body.account == null ? '' : body.account).trim();
    var found = _planFindAccount(doc, id, account);
    var warnings = [];

    if (op === 'upsert') {
      var subcon = _planCanonicalSubcont(body.subcon);
      var teamType = String(body.teamType == null ? '' : body.teamType).trim();
      var resourceRemark = String(body.resourceRemark == null ? '' : body.resourceRemark).trim();
      if (found) {
        // Update fields. A blank account/subcon keeps the existing value so a
        // partial update cannot wipe the member's identity.
        if (account) found.account = account;
        if (subcon) found.subcon = subcon;
        else if (body.subcon !== undefined) warnings.push('Unknown subcont "' + body.subcon + '" for ' + found.account + ' — kept "' + found.subcon + '"');
        if (body.teamType !== undefined) found.teamType = teamType;
        if (body.resourceRemark !== undefined) found.resourceRemark = resourceRemark;
        found.removedAt = '';
      } else {
        if (!account) {
          warnings.push('Cannot add a member without an account name');
        } else {
          // A new member with an account that already exists (case-insensitively)
          // updates that member instead of creating a duplicate.
          var lc = account.toLowerCase();
          var dup = null;
          for (var d = 0; d < doc.accounts.length; d++) {
            if (String(doc.accounts[d].account).trim().toLowerCase() === lc) { dup = doc.accounts[d]; break; }
          }
          if (dup) {
            dup.account = account;
            if (subcon) dup.subcon = subcon;
            if (body.teamType !== undefined) dup.teamType = teamType;
            if (body.resourceRemark !== undefined) dup.resourceRemark = resourceRemark;
            dup.removedAt = '';
            found = dup;
          } else {
            found = {
              id: _planNewRosterId(account),
              account: account,
              subcon: subcon,
              teamType: teamType,
              resourceRemark: resourceRemark,
              removedAt: ''
            };
            doc.accounts.push(found);
          }
        }
      }
    } else if (op === 'remove') {
      if (!found) {
        warnings.push('Member not found — nothing to remove');
      } else {
        found.removedAt = new Date().toISOString();
      }
    } else if (op === 'rename') {
      var newAccount = String(body.newAccount == null ? '' : body.newAccount).trim();
      if (!found) {
        warnings.push('Member not found — nothing to rename');
      } else if (!newAccount) {
        warnings.push('New account name is required');
      } else {
        found.account = newAccount;
      }
    } else if (op === 'setSubcont') {
      var newSub = _planCanonicalSubcont(body.subcon);
      if (!found) {
        warnings.push('Member not found — team unchanged');
      } else if (!newSub) {
        warnings.push('Unknown subcont "' + body.subcon + '" — team unchanged');
      } else {
        found.subcon = newSub;
      }
    } else {
      warnings.push('Unknown roster command "' + op + '"');
    }

    doc.version = doc.version + 1;
    _planWriteRosterDoc(doc);

    return ContentService
      .createTextOutput(JSON.stringify({
        status:   'ok',
        version:  doc.version,
        roster:   doc.accounts,
        warnings: warnings
      }))
      .setMimeType(ContentService.MimeType.JSON);
  } finally {
    lock.releaseLock();
  }
}

// ============================================================
// ONE-SHOT ROSTER MIGRATION
//
// Run this ONCE in the Apps Script editor to convert the old, corrupted roster
// state into the new roster document:
//
//   1. Merge duplicate account variants (the user confirmed these are the same
//      people — `warsidi_poca_tl` / `Warsidi_poca_tl` / `warsidi.poca_tl` /
//      `warsidi_poca_wj` are all one person, canonical name `warsidi_poca_tl`).
//   2. Normalise every subcont to its canonical uppercase key
//      (`Poca` + `POCA` -> `POCA`, `Intisel` -> `INTISEL`, ...).
//   3. Drop the debug/test accounts (`TESTACC`, `FIXCHECK_1`).
//   4. Assign a stable id to every account.
//   5. Back-fill `rosterId` on every `plan` sheet row, matching by account
//      name case-insensitively.
//
// It reads the legacy `planRosterEdits` map, so run it BEFORE deploying the
// new client (which no longer writes to that map).
//
// Returns a human-readable summary. Safe to run twice: the second run finds
// nothing to do.
// ============================================================
function migratePlanRoster() {
  var report = [];
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(30000);
  } catch (lockErr) {
    return 'MIGRATION ABORTED: ' + lockErr.message;
  }
  try {
    // --- Read the legacy sources ---
    var legacyEdits = {};
    try {
      var rawEdits = PropertiesService.getScriptProperties().getProperty('planRosterEdits') || '{}';
      legacyEdits = JSON.parse(rawEdits) || {};
    } catch (e) {
      legacyEdits = {};
    }
    var legacyDelta = { added: [], removed: [] };
    try {
      var rawDelta = PropertiesService.getScriptProperties().getProperty('planRoster') || '';
      legacyDelta = rawDelta ? (JSON.parse(rawDelta) || legacyDelta) : legacyDelta;
    } catch (e) {
      legacyDelta = { added: [], removed: [] };
    }
    if (!Array.isArray(legacyDelta.added)) legacyDelta.added = [];
    if (!Array.isArray(legacyDelta.removed)) legacyDelta.removed = [];

    // --- Build the merged account set ---
    // Order: legacy edits first (richest data), then legacy added, then any
    // base-roster entries we know about. Later duplicates are skipped.
    var byLc = {};
    var order = [];
    function consider(acct, subcon, teamType, resourceRemark, source) {
      var name = String(acct == null ? '' : acct).trim();
      if (!name) return null;
      var lc = name.toLowerCase();
      // Accounts the user confirmed are the same person.
      var canonical = null;
      if (lc === 'warsidi_poca_tl' || lc === 'warsidi.poca_tl' ||
          lc === 'warsidi_poca_wj' || lc === 'warsidi.poca_wj') {
        canonical = 'warsidi_poca_tl';
        lc = canonical;
      }
      // Debug/test residue from development — drop it.
      if (canonical === null && (name === 'TESTACC' || name === 'FIXCHECK_1' ||
          name === 'BROKEN_TEST')) {
        report.push('DROPPED test account: ' + name);
        return null;
      }
      var existing = byLc[lc];
      if (existing) {
        // Same person — keep the first record but fill in any blank fields.
        var changed = false;
        if (!existing.subcon && subcon) { existing.subcon = subcon; changed = true; }
        if (!existing.teamType && teamType) { existing.teamType = teamType; changed = true; }
        if (!existing.resourceRemark && resourceRemark) { existing.resourceRemark = resourceRemark; changed = true; }
        if (changed) {
          report.push('MERGED duplicate "' + existing.account + '" (also seen as "' + name + '") — filled blank fields');
        } else {
          report.push('MERGED duplicate "' + name + '" into "' + existing.account + '"');
        }
        return existing;
      }
      var rec = {
        id:             _planNewRosterId(canonical || name),
        account:        canonical || name,
        subcon:         _planCanonicalSubcont(subcon),
        teamType:       String(teamType == null ? '' : teamType).trim(),
        resourceRemark: String(resourceRemark == null ? '' : resourceRemark).trim(),
        removedAt:      ''
      };
      byLc[lc] = rec;
      order.push(rec);
      return rec;
    }

    var editKeys = Object.keys(legacyEdits);
    for (var i = 0; i < editKeys.length; i++) {
      var e = legacyEdits[editKeys[i]] || {};
      consider(e.account || editKeys[i], e.subcon, e.teamType, e.resourceRemark, 'edit');
    }
    for (var a = 0; a < legacyDelta.added.length; a++) {
      var ar = legacyDelta.added[a] || {};
      consider(ar.account, ar.subcon, ar.teamType, ar.resourceRemark, 'added');
    }

    // --- Normalise subcons, reporting every change ---
    for (var n = 0; n < order.length; n++) {
      var rec = order[n];
      var before = String(rec.subcon || '');
      var canon = _planCanonicalSubcont(before);
      if (canon && canon !== before) {
        report.push('SUBCONT ' + rec.account + ': "' + before + '" -> "' + canon + '"');
        rec.subcon = canon;
      } else if (!canon && before) {
        report.push('WARNING unknown subcont "' + before + '" for ' + rec.account + ' — left as-is');
        rec.subcon = before.toUpperCase();
      }
    }

    // --- Removed accounts ---
    var removedLc = {};
    for (var rm = 0; rm < legacyDelta.removed.length; rm++) {
      removedLc[String(legacyDelta.removed[rm]).trim().toLowerCase()] = true;
    }
    var kept = [];
    for (var k = 0; k < order.length; k++) {
      if (removedLc[order[k].account.toLowerCase()]) {
        report.push('REMOVED (was in removed list): ' + order[k].account);
        continue;
      }
      kept.push(order[k]);
    }

    // --- Write the roster document ---
    var doc = { version: 1, accounts: kept };
    _planWriteRosterDoc(doc);
    report.push('ROSTER: ' + kept.length + ' accounts written, version 1');

    // --- Back-fill rosterId on the plan sheet ---
    var sh = getOrCreatePlanSheet();
    var last = sh.getLastRow();
    var backfilled = 0, unmatched = 0;
    if (last > 1) {
      var vals = sh.getRange(2, 1, last - 1, 11).getValues();
      var rowIdMap = {};   // sheet row number -> roster id
      var rowAcctMap = {}; // sheet row number -> account (for the fallback write)
      for (var v = 0; v < vals.length; v++) {
        var sheetRow = v + 2;
        var acct = String(vals[v][1] == null ? '' : vals[v][1]).trim();
        if (!acct) continue;
        var lcAcct = acct.toLowerCase();
        var hit = byLc[lcAcct];
        if (!hit && (lcAcct === 'warsidi.poca_tl' || lcAcct === 'warsidi_poca_wj')) {
          hit = byLc['warsidi_poca_tl'];
        }
        if (hit) {
          rowIdMap[sheetRow] = hit.id;
          rowAcctMap[sheetRow] = hit.account;
        } else {
          unmatched++;
          rowAcctMap[sheetRow] = acct;
        }
      }
      var rowNums = Object.keys(rowIdMap);
      if (rowNums.length > 0) {
        // Write rosterId into column 12 (L). Do it in one batched call.
        var batch = [];
        for (var b = 0; b < rowNums.length; b++) {
          var rn = Number(rowNums[b]);
          batch.push([rn, rowIdMap[rn]]);
        }
        // Group consecutive rows to minimise API calls.
        var startRow = batch[0][0], values = [], prev = batch[0][0];
        values.push([batch[0][1]]);
        for (var t = 1; t < batch.length; t++) {
          if (batch[t][0] === prev + 1) {
            values.push([batch[t][1]]);
          } else {
            sh.getRange(startRow, 12, values.length, 1).setValues(values);
            startRow = batch[t][0];
            values = [[batch[t][1]]];
          }
          prev = batch[t][0];
        }
        sh.getRange(startRow, 12, values.length, 1).setValues(values);
        backfilled = batch.length;
      }
      // Rename the account column for merged rows so the sheet shows the
      // canonical name rather than the old variant.
      var renameRows = [];
      for (var rr = 2; rr <= last; rr++) {
        if (rowAcctMap[rr] && byLc[String(rowAcctMap[rr]).toLowerCase()] &&
            byLc[String(rowAcctMap[rr]).toLowerCase()].account !== rowAcctMap[rr]) {
          renameRows.push([rr, byLc[String(rowAcctMap[rr]).toLowerCase()].account]);
        }
      }
      if (renameRows.length > 0) {
        var rStart = renameRows[0][0], rVals = [[renameRows[0][1]]];
        for (var q = 1; q < renameRows.length; q++) {
          if (renameRows[q][0] === rStart + rVals.length) {
            rVals.push([renameRows[q][1]]);
          } else {
            sh.getRange(rStart, 2, rVals.length, 1).setValues(rVals);
            rStart = renameRows[q][0];
            rVals = [[renameRows[q][1]]];
          }
        }
        sh.getRange(rStart, 2, rVals.length, 1).setValues(rVals);
        report.push('RENAMED ' + renameRows.length + ' sheet row(s) to the canonical account name');
      }
    }
    report.push('SHEET: ' + backfilled + ' row(s) back-filled with rosterId, ' + unmatched + ' unmatched');
    report.push('NOTE: run this again later if new legacy rows appear; it is safe to re-run');

    return report.join('\n');
  } finally {
    lock.releaseLock();
  }
}

// Save the surge-managed plan roster delta. Light validation keeps the
// schema tight so the client can rely on it. Same trust model as
// handleSetActiveDate: no auth, surge check is client-side.
// DEPRECATED: superseded by handlePlanRosterCommand().
function handlePlanRosterSave(body) {
  var added   = Array.isArray(body.added)   ? body.added   : [];
  var removed = Array.isArray(body.removed) ? body.removed : [];
  // Normalize each added entry and drop entries with empty account.
  var cleanAdded = added.map(function(r) {
    return {
      account:        String(r.account || '').trim(),
      subcon:         String(r.subcon || '').trim(),
      teamType:       String(r.teamType || '').trim(),
      resourceRemark: String(r.resourceRemark || '').trim()
    };
  }).filter(function(r) { return r.account !== ''; });
  var cleanRemoved = removed.map(String).filter(function(s) { return s !== ''; });

  // Serialise concurrent roster writes so one surge action can't drop another's.
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(30000);
  } catch (lockErr) {
    return ContentService
      .createTextOutput(JSON.stringify({ status: 'error', message: 'Server busy, please retry' }))
      .setMimeType(ContentService.MimeType.JSON);
  }
  try {
    PropertiesService.getScriptProperties().setProperty('planRoster',
      JSON.stringify({ added: cleanAdded, removed: cleanRemoved }));
    return ContentService
      .createTextOutput(JSON.stringify({ status: 'ok', added: cleanAdded.length, removed: cleanRemoved.length }))
      .setMimeType(ContentService.MimeType.JSON);
  } finally {
    lock.releaseLock();
  }
}

// Serve planRosterEdits: a map of account -> { account, subcon, teamType, resourceRemark }
// overrides. Returns the raw JSON map stored in ScriptProperties.
function servePlanRosterEdits() {
  var raw = PropertiesService.getScriptProperties().getProperty('planRosterEdits') || '{}';
  return ContentService
    .createTextOutput(raw)
    .setMimeType(ContentService.MimeType.JSON);
}

// Save plan roster field edits. The client sends an array of edit objects,
// each containing account + the 4 editable roster fields. Merges into the
// existing map: new values override old ones for the same account.
//
// A script lock serialises the read-merge-write so two concurrent saves
// cannot clobber each other's edits.
function handlePlanRosterEdit(body) {
  var edits = Array.isArray(body.edits) ? body.edits : [];
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(30000);
  } catch (lockErr) {
    return ContentService
      .createTextOutput(JSON.stringify({ status: 'error', message: 'Server busy, please retry' }))
      .setMimeType(ContentService.MimeType.JSON);
  }
  try {
    var props = {};
    try {
      var existing = PropertiesService.getScriptProperties().getProperty('planRosterEdits') || '{}';
      props = JSON.parse(existing);
    } catch (e) {
      props = {};
    }
    if (!props || typeof props !== 'object') props = {};
    for (var i = 0; i < edits.length; i++) {
      var e = edits[i];
      var acct = String(e.account || '').trim();
      if (!acct) continue;
      // Remove any existing key that differs only by case, otherwise the same
      // person can end up with two entries ("Warsidi_poca_tl" vs "warsidi_poca_tl")
      // and an edit to one of them will not apply to the other.
      var lc = acct.toLowerCase();
      Object.keys(props).forEach(function(k){
        if (k !== acct && String(k).trim().toLowerCase() === lc) delete props[k];
      });
      props[acct] = {
        account:        acct,
        subcon:         String(e.subcon || '').trim(),
        teamType:       String(e.teamType || '').trim(),
        resourceRemark: String(e.resourceRemark || '').trim()
      };
    }
    PropertiesService.getScriptProperties().setProperty('planRosterEdits', JSON.stringify(props));
    return ContentService
      .createTextOutput(JSON.stringify({ status: 'ok', count: edits.length }))
      .setMimeType(ContentService.MimeType.JSON);
  } finally {
    lock.releaseLock();
  }
}

// Remove one stored plan roster field edit (matched case-insensitively).
// Used to clear stale/duplicate entries so a rename does not leave the old
// account behind.
function handlePlanRosterEditDelete(body) {
  var acct = String(body.account || '').trim();
  if (!acct) {
    return ContentService
      .createTextOutput(JSON.stringify({ status: 'error', message: 'account is required' }))
      .setMimeType(ContentService.MimeType.JSON);
  }
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(30000);
  } catch (lockErr) {
    return ContentService
      .createTextOutput(JSON.stringify({ status: 'error', message: 'Server busy, please retry' }))
      .setMimeType(ContentService.MimeType.JSON);
  }
  try {
    var props = {};
    try {
      var existing = PropertiesService.getScriptProperties().getProperty('planRosterEdits') || '{}';
      props = JSON.parse(existing);
    } catch (e) {
      props = {};
    }
    if (!props || typeof props !== 'object') props = {};
    var lc = acct.toLowerCase();
    var removed = 0;
    Object.keys(props).forEach(function(k){
      if (String(k).trim().toLowerCase() === lc) { delete props[k]; removed++; }
    });
    PropertiesService.getScriptProperties().setProperty('planRosterEdits', JSON.stringify(props));
    return ContentService
      .createTextOutput(JSON.stringify({ status: 'ok', removed: removed }))
      .setMimeType(ContentService.MimeType.JSON);
  } finally {
    lock.releaseLock();
  }
}
