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
  team:       12,  // M — Team
  city:       13,  // N — City
  tp:         14,  // O — TP
  km:         15,  // P — Km
  peOwner:    16,  // Q — PE Owner
  lat:        17,  // R — Latitude
  lng:        18   // S — Longitude
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

    // Partial match fallback
    if (idxDuId === -1) {
      for (var k in colMap) {
        if (k.indexOf('du') !== -1 && k.indexOf('id') !== -1) { idxDuId = colMap[k]; break; }
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
      gap:         String(row[COL.gap]         || '').trim(),
      remark:      String(row[COL.remark]      || '').trim(),
      team:        String(row[COL.team]        || '').trim(),
      city:        String(row[COL.city]        || '').trim(),
      tp:          String(row[COL.tp]          || '').trim(),
      km:          String(row[COL.km]          || '').trim(),
      peOwner:     String(row[COL.peOwner]     || '').trim(),
      lat:         parseNum(row[COL.lat]),
      lng:         parseNum(row[COL.lng]),
      ring:        String(row[COL.ring]        || '').trim(),
      seq:         String(row[COL.seq]         || '').trim()
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
//          duId, siteName, activityRemark, dailyPlanActivity, rowIdx
// rowIdx is the 10th column (0-based index 9) — it preserves the
// client-side roster position so copy-from-last-date matches
// correctly even after a user reorders or duplicates rows.
// ============================================================

// --- HELPER: get or create the Plan sheet ---
function getOrCreatePlanSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('plan');
  if (!sheet) {
    sheet = ss.insertSheet('plan');
    sheet.getRange(1, 1, 1, 10).setValues([[
      'date', 'account', 'subcon', 'teamType', 'resourceRemark',
      'duId', 'siteName', 'activityRemark', 'dailyPlanActivity', 'rowIdx'
    ]]);
    sheet.setFrozenRows(1);
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
  var values = sh.getRange(2, 1, last - 1, 10).getValues();
  var want = date ? String(date) : '';
  var rows = [];
  for (var i = 0; i < values.length; i++) {
    var r = values[i];
    if (want && String(r[0]) !== want) continue;
    rows.push({
      date:             String(r[0]),
      account:          String(r[1]),
      subcon:           String(r[2]),
      teamType:         String(r[3]),
      resourceRemark:   String(r[4]),
      duId:             String(r[5]),
      siteName:         String(r[6]),
      activityRemark:   String(r[7]),
      dailyPlanActivity:String(r[8]),
      rowIdx:           r[9] === '' || r[9] === null || isNaN(Number(r[9])) ? -1 : Number(r[9])
    });
  }
  return ContentService
    .createTextOutput(JSON.stringify(rows))
    .setMimeType(ContentService.MimeType.JSON);
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
  var data = sh.getRange(2, 1, last - 1, 1).getValues();
  var seen = {};
  var dates = [];
  for (var i = 0; i < data.length; i++) {
    var d = String(data[i][0]);
    if (d && !seen[d]) { seen[d] = true; dates.push(d); }
  }
  dates.sort();
  dates.reverse();
  return ContentService
    .createTextOutput(JSON.stringify(dates))
    .setMimeType(ContentService.MimeType.JSON);
}

// Save all rows for a given date. Per-date replace strategy — server
// deletes every row matching the date, then writes the submitted set
// in the order received. Idempotent; last writer wins for concurrent edits.
function handlePlanSave(body) {
  var date = String(body.date || '').trim();
  var rows = body.rows || [];
  if (!date) {
    return ContentService
      .createTextOutput(JSON.stringify({ status: 'error', message: 'Missing date' }))
      .setMimeType(ContentService.MimeType.JSON);
  }

  // Delete all rows matching this date (reverse order to keep indices stable)
  var sh = getOrCreatePlanSheet();
  var last = sh.getLastRow();
  if (last > 1) {
    var existing = sh.getRange(2, 1, last - 1, 1).getValues();
    for (var i = existing.length - 1; i >= 0; i--) {
      if (String(existing[i][0]) === date) sh.deleteRow(i + 2);
    }
  }

  // Write the new set if any
  var written = 0;
  if (rows.length) {
    var out = rows.map(function(r) {
      return [
        date,
        String(r.account           || ''),
        String(r.subcon            || ''),
        String(r.teamType          || ''),
        String(r.resourceRemark    || ''),
        String(r.duId              || ''),
        String(r.siteName          || ''),
        String(r.activityRemark    || ''),
        String(r.dailyPlanActivity || ''),
        (r.rowIdx === '' || r.rowIdx === null || r.rowIdx === undefined) ? '' : Number(r.rowIdx)
      ];
    });
    sh.getRange(sh.getLastRow() + 1, 1, out.length, 10).setValues(out);
    written = out.length;
  }

  return ContentService
    .createTextOutput(JSON.stringify({ status: 'ok', written: written }))
    .setMimeType(ContentService.MimeType.JSON);
}
