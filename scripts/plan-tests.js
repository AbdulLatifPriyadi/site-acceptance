// Unit tests for the Plan tab roster rework.
//
// The functions are extracted from the real index.html / SiteAcceptanceTracker.gs
// so these tests exercise the shipped code, not a copy of it.
//
// Run with:  node scripts\plan-tests.js
const fs = require('fs');
const vm = require('vm');
const path = require('path');

const ROOT = path.resolve(__dirname, '..');
const HTML = path.join(ROOT, 'index.html');
const GS   = path.join(ROOT, 'SiteAcceptanceTracker.gs');
const html = fs.readFileSync(HTML, 'utf8');
const gs   = fs.readFileSync(GS, 'utf8');

let pass = 0, fail = 0;
function check(name, actual, expected) {
  const a = JSON.stringify(actual), e = JSON.stringify(expected);
  if (a === e) { console.log(`  [PASS] ${name}`); pass++; }
  else { console.log(`  [FAIL] ${name}\n         got:      ${a}\n         expected: ${e}`); fail++; }
}
function section(t){ console.log('\n== ' + t + ' =='); }

function extractClient(src, from, to) {
  const a = src.indexOf(from);
  if (a < 0) throw new Error('not found: ' + from);
  const b = src.indexOf(to, a);
  if (b < 0) throw new Error('end marker not found: ' + to);
  return src.slice(a, b);
}

// ---------------------------------------------------------------- extract ---
const canonBlock = extractClient(html,
  '  const PLAN_SUBCONTS = [',
  '\n  const PLAN_RESOURCE_REMARK_OPTS');

const client = new Function(canonBlock + '\nreturn { canonicalSubcontKey, PLAN_SUBCONTS, PLAN_SUBCONT_SET };')();
const { canonicalSubcontKey, PLAN_SUBCONTS } = client;

// ------------------------------------------------------------------ tests ---
section('T1 — subcont canonicalisation');
check('Poca -> POCA',           canonicalSubcontKey('Poca'),   'POCA');
check('POCA -> POCA',           canonicalSubcontKey('POCA'),   'POCA');
check('poca -> POCA',           canonicalSubcontKey('poca'),   'POCA');
check('  Poca  -> POCA',        canonicalSubcontKey('  Poca  '), 'POCA');
check('Intisel -> INTISEL',     canonicalSubcontKey('Intisel'),'INTISEL');
check('INTISEL -> INTISEL',     canonicalSubcontKey('INTISEL'),'INTISEL');
check('empty -> empty',         canonicalSubcontKey(''),       '');
check('null -> empty',          canonicalSubcontKey(null),     '');
check('unknown -> empty',       canonicalSubcontKey('NOSUCH'), '');
check('trailing digits fallback', canonicalSubcontKey('INTISEL12'), 'INTISEL');
check('all canonical keys are uppercase',
  PLAN_SUBCONTS.every(k => k === k.toUpperCase()), true);
check('no duplicate subcont keys',
  new Set(PLAN_SUBCONTS).size, PLAN_SUBCONTS.length);

// --------------------------------------------------- server-side subcont ---
section('T3/T9 — server canonical subcont + id generation');
const gsCtx = {
  PropertiesService: { getScriptProperties: () => ({ getProperty: () => '', setProperty: () => {} }) },
  ContentService: { createTextOutput: (t) => ({ setMimeType: () => ({ _t: t }) }), MimeType: { JSON: 'json' } },
  LockService: { getScriptLock: () => ({ waitLock: () => {}, releaseLock: () => {} }) }
};
const gsFn = new Function(
  'PropertiesService','ContentService','LockService',
  gs + '\nreturn { _planCanonicalSubcont, _planNewRosterId, _planFindAccount, _planReadRosterDoc, PLAN_SUBCONT_KEYS };'
)(gsCtx.PropertiesService, gsCtx.ContentService, gsCtx.LockService);

check('server: Poca -> POCA',      gsFn._planCanonicalSubcont('Poca'),   'POCA');
check('server: Intisel -> INTISEL',gsFn._planCanonicalSubcont('Intisel'),'INTISEL');
check('server: unknown -> empty', gsFn._planCanonicalSubcont('NOSUCH'), '');
check('server: empty -> empty',   gsFn._planCanonicalSubcont(''),       '');
check('server: id is a truncated slug + random suffix',
  /^[a-z0-9]{1,16}_[a-z0-9]{6}$/.test(gsFn._planNewRosterId('intisel_tl_rahmat_32051919')), true);
check('server: id slug strips non-alphanumerics',
  gsFn._planNewRosterId('intisel_tl_rahmat_32051919').split('_')[0], 'intiseltlrahmat3');
check('server: id suffix is 6 chars',
  gsFn._planNewRosterId('abc').split('_').pop().length, 6);
check('server: two ids for the same account differ',
  gsFn._planNewRosterId('abc') !== gsFn._planNewRosterId('abc'), true);

// -------------------------------------------------- client roster merging ---
section('T5 — roster merge by rosterId, legacy account fallback');
const mergeSrc = extractClient(html,
  '  function _mergeSavedRows(roster, saved){',
  '\n  async function fetchPlanData');
const merge = new Function('canonicalSubcontKey',
  mergeSrc + '\nreturn _mergeSavedRows;')(canonicalSubcontKey);

const roster = [
  { id: 'r1', account: 'alpha_one',   subcon: 'PAB',  teamType: 'WL', resourceRemark: 'Shared' },
  { id: 'r2', account: 'Warsidi_poca_tl', subcon: 'POCA', teamType: 'WL', resourceRemark: 'Shared' },
  { id: 'r3', account: 'gamma_three', subcon: 'NIM',  teamType: 'CSG',resourceRemark: 'Dedicated' }
];
const saved = [
  { rosterId: 'r1', account: 'alpha_one', subcon: 'PAB', teamType: 'WL', resourceRemark: 'Shared',
    duId: 'DU1', siteName: 'Site A', activityRemark: 'WORK', dailyPlanActivity: 'MOS', result: 'ok' },
  // r2's server row still uses the OLD capitalisation and has NO rosterId —
  // the legacy fallback must still attach it.
  { rosterId: '', account: 'warsidi_poca_tl', subcon: 'POCA', teamType: 'WL', resourceRemark: 'Shared',
    duId: 'DU2', siteName: 'Site B', activityRemark: 'WORK', dailyPlanActivity: 'ATP', result: 'done' },
  { rosterId: 'r3', account: 'gamma_three', subcon: 'NIM', teamType: 'CSG', resourceRemark: 'Dedicated',
    duId: '', siteName: '', activityRemark: '', dailyPlanActivity: '', result: '' }
];
const merged = merge(roster, saved);
check('merge returns a Map', merged.rows instanceof Map, true);
check('merge order matches roster order', merged.order, ['r1','r2','r3']);
check('r1 attached by rosterId', merged.rows.get('r1').duId, 'DU1');
check('r2 attached via case-insensitive account fallback', merged.rows.get('r2').duId, 'DU2');
check('r2 keeps the roster display name', merged.rows.get('r2').account, 'Warsidi_poca_tl');
check('r3 with no server data gets empty plan fields', merged.rows.get('r3').duId, '');
check('every merged row carries its rosterId',
  merged.order.every(id => merged.rows.get(id).rosterId === id), true);

// The critical regression: re-ordering the roster must NOT move plan data.
const rosterReordered = [roster[1], roster[0], roster[2]];
const mergedReordered = merge(rosterReordered, saved);
check('re-ordering the roster does not move plan data',
  mergedReordered.rows.get('r1').duId, 'DU1');
check('re-ordering the roster does not move plan data (r2)',
  mergedReordered.rows.get('r2').duId, 'DU2');
check('re-ordered order reflects the new roster order',
  mergedReordered.order, ['r2','r1','r3']);

// An unmatched server row must be DROPPED, never mis-assigned.
const mergedExtra = merge(roster, saved.concat([
  { rosterId: 'rX', account: 'stranger_xyz', duId: 'DUX', siteName: 'Site X' }
]));
check('unmatched server row is dropped (not mis-assigned)',
  mergedExtra.rows.get('r1').duId, 'DU1');
check('unmatched server row did not leak into any row',
  Array.from(mergedExtra.rows.values()).some(r => r.duId === 'DUX'), false);

// ------------------------------------------------- roster command detection ---
section('T6 — roster op detection (rename / team move / remark)');
const detectSrc = extractClient(html,
  '  function _detectRosterOps() {',
  "\n  document.getElementById('planSetActiveBtn')");
if (!detectSrc.includes('function _detectRosterOps')) {
  console.log('  [FAIL] could not extract _detectRosterOps');
  fail++;
}
const detect = new Function(
  'planRowOrder','planRows','planRosterById','canonicalSubcontKey',
  detectSrc + '\nreturn _detectRosterOps();'
);

const order = ['r1','r2'];
const rows = new Map([
  ['r1', { rosterId:'r1', account:'alpha_one', subcon:'PAB', teamType:'WL', resourceRemark:'Shared' }],
  ['r2', { rosterId:'r2', account:'RENAMED_guy', subcon:'CCSI', teamType:'CSG', resourceRemark:'Dedicated' }]
]);
const byId = new Map([
  ['r1', { id:'r1', account:'alpha_one', subcon:'PAB', teamType:'WL', resourceRemark:'Shared' }],
  ['r2', { id:'r2', account:'gamma_three', subcon:'NIM', teamType:'WL', resourceRemark:'Shared' }]
]);
const ops = detect(order, rows, byId, canonicalSubcontKey);
const renameOps = ops.filter(o => o.op === 'rename');
const subOps   = ops.filter(o => o.op === 'setSubcont');
const upsertOps= ops.filter(o => o.op === 'upsert');
check('rename detected', renameOps.map(o => o.payload.newAccount), ['RENAMED_guy']);
check('rename keeps the roster id', renameOps.map(o => o.payload.id), ['r2']);
check('team move detected', subOps.map(o => o.payload.subcont), ['CCSI']);
check('team move keeps the roster id', subOps.map(o => o.payload.id), ['r2']);
check('remark change detected as upsert', upsertOps.length >= 1, true);

// No changes -> no ops.
const noOps = detect(order, new Map([
  ['r1', { rosterId:'r1', account:'alpha_one', subcon:'PAB', teamType:'WL', resourceRemark:'Shared' }],
  ['r2', { rosterId:'r2', account:'gamma_three', subcon:'NIM', teamType:'WL', resourceRemark:'Shared' }]
]), byId, canonicalSubcontKey);
check('no changes -> no roster ops', noOps.length, 0);

// ------------------------------------------------------------ excel export ---
section('T10 — Excel export builder');
const exportSrc = extractClient(html,
  "  document.getElementById('planExportBtn').addEventListener('click'",
  '\n  // ----- Share: WhatsApp messages');
if (!exportSrc.includes('aoa_to_sheet')) {
  console.log('  [FAIL] could not extract the export handler');
  fail++;
} else {
  console.log('  [OK]   export handler extracted');
  pass++;
  check('export header includes Roster ID',
    exportSrc.includes("['Roster ID', 'Account', 'Subcon'"), true);
  check('export fetches fresh server data',
    exportSrc.includes('fetchPlanData(date)'), true);
  check('export filters by team',
    exportSrc.includes('canonicalSubcontKey(sub).toLowerCase() === mySub'), true);
  check('export filename includes the date',
    exportSrc.includes("'plan-' + date"), true);
  check('export filename includes the subcont for a team user',
    exportSrc.includes("mySub ? '-' + mySub : ''"), true);
  check('export contains no credentials',
    !/SUBCONT_TOKENS|VALID_TOKENS|password/i.test(exportSrc), true);
}

// ------------------------------------------------------------- server plan ---
section('T4 — server plan save / serve');
check('servePlan returns rosterId',
  /rosterId:\s*String\(r\[11\]/.test(gs), true);
check('handlePlanSave matches by rosterId',
  gs.includes('incomingById'), true);
check('handlePlanSave has a legacy account fallback',
  gs.includes('incomingByAccountLc'), true);
check('plan sheet writes 12 columns',
  gs.includes('PLAN_NUM_COLS = 12'), true);
check('rosterId column is appended (column 12, after rowIdx)',
  gs.includes('PLAN_COL_ROSTER_ID = 12'), true);
check('save response includes the diff',
  gs.includes('unchanged: unchanged') && gs.includes('warnings:  warnings'), true);
check('prune by roster id is supported',
  gs.includes('pruneIdSet'), true);

// ---------------------------------------------------------- command type ---
section('T6 — roster command payload contract');
// The server dispatches on body.type. If the client omits it the request
// falls through to the default handler and every roster write fails.
const cmdSrc = extractClient(html,
  '  async function postPlanRosterCommand(op, payload){',
  '\n  // Queue a roster command');
check('command body includes type: planRosterCommand',
  cmdSrc.includes("type: 'planRosterCommand'"), true);
check('command body includes the version',
  cmdSrc.includes('version: planRosterVersion'), true);
check('command body includes the op',
  cmdSrc.includes('op: op'), true);
check('command sends method POST',
  cmdSrc.includes("method: 'POST'"), true);
check('server dispatches on planRosterCommand',
  gs.includes("body.type === 'planRosterCommand'"), true);

// ---------------------------------------------------------------- summary ---
console.log('\n' + pass + ' passed, ' + fail + ' failed');
process.exit(fail ? 1 : 0);
