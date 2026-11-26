// Integration test for the Plan tab roster flow.
//
// The real client functions are extracted from index.html and run inside a
// Node vm context with a mocked Apps Script backend. All state lives INSIDE
// the context so there is no cross-context copying.
const fs = require('fs');
const vm = require('vm');

const html = fs.readFileSync('C:/Users/legio/Downloads/site-acceptance/index.html', 'utf8');
const all = [...html.matchAll(/<script(?![^>]*\ssrc=)[^>]*>([\s\S]*?)<\/script>/gi)];
const src = all[all.length - 1][1];

function extractClient(src, from, to) {
  const a = src.indexOf(from);
  if (a < 0) throw new Error('not found: ' + from);
  const b = src.indexOf(to, a);
  if (b < 0) throw new Error('end marker not found: ' + to);
  return src.slice(a, b);
}

let pass = 0, fail = 0;
const results = [];
const t = (name, actual, expected) => {
  const a = JSON.stringify(actual), e = JSON.stringify(expected);
  results.push({ name, ok: a === e, a, e });
};

const canonBlock   = extractClient(html, '  const PLAN_SUBCONTS = [', '\n  const PLAN_RESOURCE_REMARK_OPTS');
const mergeBlock   = extractClient(html, '  function _mergeSavedRows(roster, saved){', '\n  async function fetchPlanData');
const detectBlock  = extractClient(html, '  function _detectRosterOps() {', "\n  document.getElementById('planSetActiveBtn')");
const applyBlock   = extractClient(html, '  function _applyRosterDoc(accounts, version){', '\n  // Fetch the roster document');
const fetchBlock   = extractClient(html, '  async function fetchPlanRosterDoc(){', '\n  // Send one roster command');
const cmdBlock     = extractClient(html, '  async function postPlanRosterCommand(op, payload){', '\n  // Queue a roster command');
const queueBlock   = extractClient(html, '  function _queueRosterOp(op, payload){', '\n  // Show the roster sync banner');
const retryBlock   = extractClient(html, '  async function retryPendingRosterOps(){', '\n  // The effective roster');
const bannerBlock  = extractClient(html, '  function _renderRosterSyncBanner(){', '\n  // Show the roster count');
const statusBlock  = extractClient(html, '  function _renderRosterStatus(){', '\n  // Retry every queued roster command');

const script = [canonBlock, mergeBlock, detectBlock, applyBlock, fetchBlock,
                cmdBlock, queueBlock, retryBlock, bannerBlock, statusBlock].join('\n');

const sandbox = {
  console,
  JSON, String, Number, Array, Object, Math, Date, isNaN, Map, Set, Promise, Error, RegExp,
  serverConfigured: () => true,
  SCRIPT_URL: 'https://mock.test/exec',
  document: { getElementById: () => null }
};

// The roster state must live INSIDE the context so it shares the same
// bindings as the extracted functions.
const stateSrc = `
  var planRosterAccounts = [];
  var planRosterById = new Map();
  var planRosterByAccountLc = new Map();
  var planRosterVersion = 0;
  var planRosterLoaded = false;
  var planRosterLastSync = null;
  var planRosterPendingOps = [];
  var planRows = new Map();
  var planRowOrder = [];
  var planDirty = false;
  var mySubcont = '';
`;

// The mock backend is defined INSIDE the context so it shares the context's
// globals (no cross-realm copying).
const mockSrc = `
  var MOCK_ROSTER = { version: 3, accounts: [
    { id:'r1', account:'alpha_one',      subcon:'PAB',  teamType:'WL',  resourceRemark:'Shared',    removedAt:'' },
    { id:'r2', account:'Warsidi_poca_tl', subcon:'POCA', teamType:'WL',  resourceRemark:'Shared',    removedAt:'' },
    { id:'r3', account:'gamma_three',    subcon:'NIM',  teamType:'CSG', resourceRemark:'Dedicated', removedAt:'' },
    { id:'r4', account:'removed_guy',    subcon:'PAB',  teamType:'WL',  resourceRemark:'Shared',    removedAt:'2026-01-01T00:00:00.000Z' }
  ]};
  var MOCK_CALLS = [];
  globalThis.__mockFetchRef = fetchWithTimeout;
  var MOCK_FAIL_NEXT = false;

  function mockJson(o, status) {
    var st = (status === undefined || status === null) ? 200 : status;
    return { ok: st >= 200 && st < 300, status: st,
             json: function(){ return Promise.resolve(o); },
             text: function(){ return Promise.resolve(JSON.stringify(o)); } };
  }

  function fetchWithTimeout(url, options) {
    options = options || {};
    var method = options.method || 'GET';
    var body = options.body ? JSON.parse(options.body) : null;
    MOCK_CALLS.push({ url: url, method: method, body: body });
    if (MOCK_FAIL_NEXT) { MOCK_FAIL_NEXT = false; return Promise.reject(new TypeError('Failed to fetch')); }
    if (method === 'GET') {
      if (url.indexOf('type=planRosterDoc') >= 0) {
        MOCK_CALLS[MOCK_CALLS.length - 1].served = MOCK_ROSTER.accounts.length;
        return Promise.resolve(mockJson({ status:'ok', version:MOCK_ROSTER.version, roster:MOCK_ROSTER.accounts }));
      }
      return Promise.resolve(mockJson({}));
    }
    var b = body;
    // The server dispatches on the "type" field — mirror that exactly so a
    // missing type is caught here rather than in production.
    if (b.type === 'planRosterCommand') {
      if (b.version !== MOCK_ROSTER.version) {
        return Promise.resolve(mockJson({ status:'stale', version:MOCK_ROSTER.version, roster:MOCK_ROSTER.accounts, message:'stale' }));
      }
      MOCK_ROSTER.version++;
      if (b.op === 'upsert') {
        var hit = null;
        for (var i = 0; i < MOCK_ROSTER.accounts.length; i++) { if (MOCK_ROSTER.accounts[i].id === b.id) { hit = MOCK_ROSTER.accounts[i]; break; } }
        if (hit) {
          if (b.account) hit.account = b.account;
          if (b.subcon) hit.subcon = b.subcon;
          if (b.teamType !== undefined) hit.teamType = b.teamType;
          if (b.resourceRemark !== undefined) hit.resourceRemark = b.resourceRemark;
        } else {
          MOCK_ROSTER.accounts.push({ id:'rNEW', account:b.account, subcon:b.subcon, teamType:b.teamType,
                                      resourceRemark:b.resourceRemark, removedAt:'' });
        }
      }
      if (b.op === 'rename') {
        for (var j = 0; j < MOCK_ROSTER.accounts.length; j++) { if (MOCK_ROSTER.accounts[j].id === b.id) MOCK_ROSTER.accounts[j].account = b.newAccount; }
      }
      if (b.op === 'remove') {
        for (var k = 0; k < MOCK_ROSTER.accounts.length; k++) { if (MOCK_ROSTER.accounts[k].id === b.id) MOCK_ROSTER.accounts[k].removedAt = '2026-01-01T00:00:00.000Z'; }
      }
      MOCK_CALLS[MOCK_CALLS.length - 1].responseKeys = Object.keys(mockJson({ status:'ok', version:MOCK_ROSTER.version, roster:MOCK_ROSTER.accounts, warnings: [] }));
      return Promise.resolve(mockJson({ status:'ok', version:MOCK_ROSTER.version, roster:MOCK_ROSTER.accounts, warnings: [] }));
    }
    // Anything else is not a roster command — the real server would fall
    // through to its default handler.
    return Promise.resolve(mockJson({ status:'error', message:'Unknown request type: ' + (b && b.type) }));
  }
`;

const ctx = vm.createContext(sandbox);
// All parts run in ONE script so `var` declarations are shared globals.
vm.runInContext(stateSrc + '\n' + mockSrc + '\n' + script + `
(function(){
  var _origApply = _applyRosterDoc;
  var _trace = [];
  _applyRosterDoc = function(accounts, version) {
    _trace.push('applyRosterDoc: accounts=' + (accounts ? accounts.length : 'null') + ' version=' + version);
    var r = _origApply(accounts, version);
    _trace.push('  -> planRosterAccounts.length=' + planRosterAccounts.length +
                ' globalThis.planRosterAccounts.length=' + globalThis.planRosterAccounts.length);
    return r;
  };
  globalThis.__trace = _trace;
})();
` + `this.__api = {
  canonicalSubcontKey: canonicalSubcontKey,
  _mergeSavedRows: _mergeSavedRows,
  _detectRosterOps: _detectRosterOps,
  _applyRosterDoc: _applyRosterDoc,
  postPlanRosterCommand: postPlanRosterCommand,
  _queueRosterOp: _queueRosterOp,
  retryPendingRosterOps: retryPendingRosterOps,
  fetchPlanRosterDoc: fetchPlanRosterDoc,
  effectiveRoster: function(){ return planRosterAccounts.filter(function(a){ return !a.removedAt; }); },
  getState: function(){ return JSON.stringify({
    planRosterAccounts: planRosterAccounts, planRosterVersion: planRosterVersion,
    planRosterLoaded: planRosterLoaded, planRosterPendingOps: planRosterPendingOps,
    planRosterById: Array.from(planRosterById.entries()),
    calls: MOCK_CALLS, mockRoster: MOCK_ROSTER,
    globalType: typeof globalThis.planRosterAccounts,
    globalIsArray: Array.isArray(globalThis.planRosterAccounts),
    globalLen: (globalThis.planRosterAccounts && globalThis.planRosterAccounts.length) || 0,
    sameRef: globalThis.planRosterAccounts === planRosterAccounts,
    mockAccLen: MOCK_ROSTER.accounts.length, trace: globalThis.__trace,
    fetchWithTimeoutIsMock: fetchWithTimeout === globalThis.__mockFetchRef,
    fetchName: fetchWithTimeout.name
  }); },
  setFailNext: function(v){ MOCK_FAIL_NEXT = v; },
  resetCalls: function(){ MOCK_CALLS.length = 0; }
};`, ctx);

const api = sandbox.__api;

(async () => {
  // ------------------------------------------------- roster load (T2) ------
  const loaded = await api.fetchPlanRosterDoc();
  t('roster fetch succeeds', loaded, true);
  const st = JSON.parse(api.getState());
  t('roster applied', st.planRosterAccounts.length, 4);
  t('version applied', st.planRosterVersion, 3);
  t('roster marked loaded', st.planRosterLoaded, true);
  t('byId index built', st.planRosterById.length, 4);
  t('effectiveRoster excludes the soft-removed member', api.effectiveRoster().length, 3);
  t('subcont normalised to uppercase',
    st.planRosterAccounts.map(a => a.subcon), ['PAB','POCA','NIM','PAB']);

  // ------------------------------------------------- roster command (T6) --
  const r1 = await api.postPlanRosterCommand('rename', { id: 'r1', newAccount: 'alpha_renamed' });
  require('fs').writeFileSync('C:/Users/legio/AppData/Local/Temp/kilo/state.json', api.getState());
  t('rename succeeds', r1.ok, true);
  const s1 = JSON.parse(api.getState());
  const call = s1.calls.find(c => c.body && c.body.type === 'planRosterCommand');
  t('a command was sent', !!call, true);
  t('command op is rename', call && call.body.op, 'rename');
  t('command carries the roster id', call && call.body.id, 'r1');
  t('command carries the new name', call && call.body.newAccount, 'alpha_renamed');
  t('command carries the current version', call && call.body.version, 3);
  t('version advanced after the command', s1.planRosterVersion, 4);
  t('roster reflects the rename',
    (s1.planRosterAccounts.find(a => a.id === 'r1') || { account: '<<missing: ' + s1.planRosterAccounts.map(a=>a.id).join(',') + '>>' }).account,
    'alpha_renamed');

  // ---------------------------------------------- stale version retry (T6) --
  api.resetCalls();
  const r2 = await api.postPlanRosterCommand('upsert', { id: 'r3', account: 'gamma_three', teamType: 'WL' });
  t('command after a version bump still succeeds', r2.ok, true);
  t('the client resynced its version',
    JSON.parse(api.getState()).planRosterVersion, JSON.parse(api.getState()).mockRoster.version);

  // ----------------------------------------- failed command is queued (T6) --
  api.setFailNext(true);
  const r3 = await api.postPlanRosterCommand('upsert', { id: 'r3', account: 'gamma_three' });
  t('a failed command reports failure', r3.ok, false);
  api._queueRosterOp('upsert', { id: 'r3', account: 'gamma_three' });
  t('the failed command is queued', JSON.parse(api.getState()).planRosterPendingOps.length, 1);
  const drained = await api.retryPendingRosterOps();
  t('the retry reports success', drained, true);
  t('the queue drains', JSON.parse(api.getState()).planRosterPendingOps.length, 0);

  // --------------------------------------- merge by rosterId (T5) --------
  // Use a fresh roster so the earlier rename does not affect this fixture.
  const roster = [
    { id: 'r1', account: 'alpha_one',      subcon: 'PAB',  teamType: 'WL', resourceRemark: 'Shared', removedAt: '' },
    { id: 'r2', account: 'Warsidi_poca_tl', subcon: 'POCA', teamType: 'WL', resourceRemark: 'Shared', removedAt: '' },
    { id: 'r3', account: 'gamma_three',    subcon: 'NIM',  teamType: 'CSG',resourceRemark: 'Dedicated', removedAt: '' }
  ];
  const saved = [
    { rosterId: 'r1', account: 'alpha_one', duId: 'DU1', siteName: 'Site A' },
    { rosterId: '', account: 'warsidi_poca_tl', duId: 'DU2', siteName: 'Site B' },
    { rosterId: 'r3', account: 'gamma_three', duId: '', siteName: '' }
  ];
  const m = api._mergeSavedRows(roster, saved);
  t('r1 merged by rosterId', m.rows.get('r1').duId, 'DU1');
  t('r2 merged by the case-insensitive account fallback', m.rows.get('r2').duId, 'DU2');
  t('r2 keeps the roster display name', m.rows.get('r2').account, 'Warsidi_poca_tl');
  t('merge order follows the roster', m.order, ['r1','r2','r3']);
  const reordered = api._mergeSavedRows([roster[1], roster[0], roster[2]], saved);
  t('re-ordering the roster does not move plan data (r1)', reordered.rows.get('r1').duId, 'DU1');
  t('re-ordering the roster does not move plan data (r2)', reordered.rows.get('r2').duId, 'DU2');
  const extra = api._mergeSavedRows(roster, saved.concat([{ rosterId:'rX', account:'stranger', duId:'DUX' }]));
  t('an unmatched server row is dropped, not mis-assigned',
    Array.from(extra.rows.values()).some(r => r.duId === 'DUX'), false);

  // ---------------------------------------- roster op detection (T6) ------
  // The function reads the module-level state, so set it directly.
  const detected = vm.runInContext(`(function(){
    planRowOrder = ['r1','r2'];
    planRows = new Map([
      ['r1', { rosterId:'r1', account:'renamed_on_screen', subcon:'PAB', teamType:'WL', resourceRemark:'Shared' }],
      ['r2', { rosterId:'r2', account:'Warsidi_poca_tl', subcon:'CCSI', teamType:'WL', resourceRemark:'Shared' }]
    ]);
    planRosterById = new Map([
      ['r1', { id:'r1', account:'alpha_one', subcon:'PAB', teamType:'WL', resourceRemark:'Shared' }],
      ['r2', { id:'r2', account:'Warsidi_poca_tl', subcon:'POCA', teamType:'WL', resourceRemark:'Shared' }]
    ]);
    return JSON.stringify(_detectRosterOps());
  })()`, ctx);
  const ops = JSON.parse(detected);
  t('rename detected', ops.filter(o => o.op === 'rename').map(o => o.payload.newAccount), ['renamed_on_screen']);
  t('rename keeps the roster id', ops.filter(o => o.op === 'rename').map(o => o.payload.id), ['r1']);
  t('team move detected', ops.filter(o => o.op === 'setSubcont').map(o => o.payload.subcont), ['CCSI']);
  t('team move keeps the roster id', ops.filter(o => o.op === 'setSubcont').map(o => o.payload.id), ['r2']);

  // ------------------------------------------------------------ report ----
  console.log('');
  for (const r of results) {
    if (r.ok) { console.log('  [PASS] ' + r.name); pass++; }
    else { console.log('  [FAIL] ' + r.name + '\n         got:      ' + r.a + '\n         expected: ' + r.e); fail++; }
  }
  console.log('\n' + pass + ' passed, ' + fail + ' failed');
  process.exit(fail ? 1 : 0);
})();
