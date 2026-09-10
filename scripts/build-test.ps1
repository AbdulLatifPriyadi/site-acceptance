# Build index-test.html from index.html with mock data interceptors

$html = Get-Content "$PSScriptRoot\..\index.html" -Raw -Encoding UTF8

# Find </head> and insert mock script after it
$headEnd = $html.IndexOf('</head>')
if ($headEnd -eq -1) {
    Write-Host "[build-test] ERROR: could not find </head> in index.html"
    exit 1
}

$beforeHead = $html.Substring(0, $headEnd + 7)  # includes </head>
$afterHead = $html.Substring($headEnd + 7)  # starts after the > of </head>

$mocks = @"

<!-- ============================================================
  LOCAL TEST FILE — DO NOT DEPLOY TO PRODUCTION
  Generated from index.html by scripts\build-test.ps1
  To regenerate after changes: run scripts\build-test.bat
================================================================ -->

<script>
// ============================================================
// MOCK DATA — standalone test mode (no Apps Script needed)
// ============================================================

var MOCK_TRACKER = [
  { batchRfi:'BATCH-A', batchCsg:'CSG-01', tpId:'TP001', surgeId:'SURG-001', duId:'32CMH_0057_CSG', siteName:'PTI_DESA_UTAMA', subcontPlan:'Intisel', gap:'GAP-A', remark:'OK - installed', fo:'RFS', city:'Bandung', tp:'Telkom', km:'12', peOwner:'John Doe', lat:-6.9175, lng:107.6191, ring:'Andir - Ciroyom', seq:'1', csgStatus:'Active', priority:'P2' },
  { batchRfi:'BATCH-A', batchCsg:'CSG-01', tpId:'TP002', surgeId:'SURG-002', duId:'32CMH_0024_NEW', siteName:'TBG_BAROS', subcontPlan:'NIM', gap:'GAP-B', remark:'Pending ATP', fo:'Termination done', city:'Bandung', tp:'Telkom', km:'8', peOwner:'Jane Doe', lat:-6.9148, lng:107.6172, ring:'Andir - Ciroyom', seq:'2', csgStatus:'Done', priority:'P2' },
  { batchRfi:'BATCH-A', batchCsg:'CSG-02', tpId:'TP003', surgeId:'SURG-003', duId:'32BDG_0224_NEW', siteName:'TBG_GURUMINDA', subcontPlan:'PAB', gap:'', remark:'Material delay', fo:'Pulling Cable', city:'Bandung', tp:'Telkom', km:'15', peOwner:'Bob Smith', lat:-6.9210, lng:107.6100, ring:'Dago - Hegarsari', seq:'3', csgStatus:'Pending', priority:'P1' },
  { batchRfi:'BATCH-B', batchCsg:'CSG-02', tpId:'TP004', surgeId:'SURG-004', duId:'32SKB_0072_NEW', siteName:'PTI_PASAR_PELABUHAN', subcontPlan:'ADW', gap:'GAP-C', remark:'OK', fo:'Survey', city:'Sukabumi', tp:'Telkom', km:'22', peOwner:'Alice Lee', lat:-6.9250, lng:107.6050, ring:'Dago - Hegarsari', seq:'4', csgStatus:'Not Started', priority:'P1' },
  { batchRfi:'BATCH-B', batchCsg:'CSG-03', tpId:'TP005', surgeId:'SURG-005', duId:'32TNG_0011_CSG', siteName:'TNG_TANJUNG', subcontPlan:'CAA', gap:'GAP-A', remark:'Integration done', fo:'Materials', city:'Tanggerang', tp:'Telkom', km:'30', peOwner:'Charlie Ng', lat:-6.9300, lng:107.5900, ring:'Andir - Ciroyom', seq:'5', csgStatus:'Active', priority:'P2' },
  { batchRfi:'BATCH-C', batchCsg:'CSG-03', tpId:'TP006', surgeId:'SURG-006', duId:'32CMH_0099_WL', siteName:'CMH_KOMP_PLUIT', subcontPlan:'TRITAMA', gap:'', remark:'Work in progress', fo:'Permit', city:'Bandung', tp:'Telkom', km:'5', peOwner:'David Tan', lat:-6.9100, lng:107.6150, ring:'Cikawao - Turangga', seq:'6', csgStatus:'', priority:'P3' },
  { batchRfi:'BATCH-C', batchCsg:'CSG-04', tpId:'TP007', surgeId:'SURG-007', duId:'32BDG_0333_NEW', siteName:'BDG_OFFICE_TLT', subcontPlan:'INNOVIS', gap:'GAP-B', remark:'Closing only', fo:'NY Assigned', city:'Bandung', tp:'Telkom', km:'10', peOwner:'Eve Wong', lat:-6.9000, lng:107.6200, ring:'Cikawao - Turangga', seq:'7', csgStatus:'Done', priority:'P3' },
  { batchRfi:'', batchCsg:'', tpId:'TP008', surgeId:'', duId:'32CMH_0010_CSG', siteName:'CMH_DAGO', subcontPlan:'Intisel', gap:'GAP-A', remark:'', fo:'DRM', city:'Bandung', tp:'Telkom', km:'3', peOwner:'Frank Lee', lat:-6.9050, lng:107.6100, ring:'Dago - Hegarsari', seq:'8', csgStatus:'Active', priority:'P1' },
  { batchRfi:'BATCH-D', batchCsg:'CSG-04', tpId:'TP009', surgeId:'SURG-008', duId:'32SKB_0055_WL', siteName:'SKB_MALL', subcontPlan:'Poca', gap:'GAP-C', remark:'ATP passed', fo:'', city:'Sukabumi', tp:'Telkom', km:'18', peOwner:'Grace Ho', lat:-6.9280, lng:107.6080, ring:'Sukabungah - Sukajadi', seq:'9', csgStatus:'Pending', priority:'P2' },
  { batchRfi:'BATCH-D', batchCsg:'', tpId:'TP010', surgeId:'', duId:'32TNG_0077_NEW', siteName:'TNG_CITY_CENTER', subcontPlan:'INDOHR', gap:'', remark:'Waiting material', fo:'', city:'Tanggerang', tp:'Telkom', km:'25', peOwner:'Henry Wu', lat:-6.9350, lng:107.5850, ring:'Sukabungah - Sukajadi', seq:'10', csgStatus:'', priority:'P2' },
  { batchRfi:'BATCH-E', batchCsg:'CSG-05', tpId:'TP011', surgeId:'SURG-009', duId:'32CJR_0043_NEW', siteName:'CJR_SITE_ALPHA', subcontPlan:'Intisel', gap:'ATP', remark:'Ready for test', fo:'RFS', city:'Cirebon', tp:'Telkom', km:'40', peOwner:'Test User', lat:-6.7320, lng:108.5521, ring:'Cirebon - Sumber', seq:'11', csgStatus:'Active', priority:'P1' },
  { batchRfi:'BATCH-E', batchCsg:'CSG-05', tpId:'TP012', surgeId:'SURG-010', duId:'32BDG_0673_NEW', siteName:'BDG_SITE_BETA', subcontPlan:'NIM', gap:'Install', remark:'In progress', fo:'Drop', city:'Bandung', tp:'Telkom', km:'18', peOwner:'Another User', lat:-6.9001, lng:107.6200, ring:'Cirebon - Sumber', seq:'12', csgStatus:'Working', priority:'P1' },
  { batchRfi:'BATCH-E', batchCsg:'CSG-06', tpId:'TP013', surgeId:'SURG-011', duId:'32BDG_0298_NEW', siteName:'BDG_SITE_GAMMA', subcontPlan:'PAB', gap:'ON AIR', remark:'Completed', fo:'Pulling Done', city:'Bandung', tp:'Telkom', km:'12', peOwner:'Final User', lat:-6.9100, lng:107.6300, ring:'Cirebon - Sumber', seq:'13', csgStatus:'Done', priority:'P1' }
];

var MOCK_PLAN = [
  { date:'2026-09-04', account:'adw_tl_agung_32092512', subcon:'ADW', teamType:'CSG', resourceRemark:'Dedicated/Shared', duId:'32CMH_0057_CSG', siteName:'PTI_DESA_UTAMA', activityRemark:'WORK', dailyPlanActivity:'Integration', result:'', rowIdx:16 },
  { date:'2026-09-04', account:'caa_tl_rezar_32020226', subcon:'CAA', teamType:'WL', resourceRemark:'Dedicated/Shared', duId:'32CMH_0024_CSG', siteName:'TBG_BAROS', activityRemark:'WORK', dailyPlanActivity:'Install', result:'Done, waiting ATP', rowIdx:23 },
  { date:'2026-09-04', account:'pab_tl_fajarfm_33091410', subcon:'PAB', teamType:'WL', resourceRemark:'Dedicated/Shared', duId:'32BDG_0224_NEW', siteName:'TBG_GURUMINDA', activityRemark:'OTHER PROJECT', dailyPlanActivity:'', result:'', rowIdx:9 },
  { date:'2026-09-04', account:'nim_tl_gustira_32053405', subcon:'NIM', teamType:'WL', resourceRemark:'Dedicated/Shared', duId:'32CMH_0099_WL', siteName:'CMH_KOMP_PLUIT', activityRemark:'WORK', dailyPlanActivity:'MOS+Install', result:'', rowIdx:3 },
  { date:'2026-09-04', account:'intisel_tl_rahmat_32051919', subcon:'Intisel', teamType:'WL', resourceRemark:'Dedicated/Shared', duId:'32SKB_0072_NEW', siteName:'PTI_PASAR_PELABUHAN', activityRemark:'WORK', dailyPlanActivity:'ATP', result:'', rowIdx:0 },
  { date:'2026-09-04', account:'intisel_tl_deni_32053528', subcon:'Intisel', teamType:'WL', resourceRemark:'Dedicated/Shared', duId:'32TNG_0011_CSG', siteName:'TNG_TANJUNG', activityRemark:'IDLE', dailyPlanActivity:'', result:'Site locked, rescheduled', rowIdx:1 },
  { date:'2026-09-04', account:'inv_tl_sandi_32042901', subcon:'INNOVIS', teamType:'WL', resourceRemark:'Dedicated/Shared', duId:'32BDG_0333_NEW', siteName:'BDG_OFFICE_TLT', activityRemark:'Closing Only', dailyPlanActivity:'MOS', result:'', rowIdx:22 }
];

var MOCK_CLOCKIN = {
  '32CMH_0057_CSG': { name:'Agung (ADW)', team:'CSG', clockIn:'07:30', clockOut:'', account:'adw_tl_agung_32092512' },
  '32CMH_0024_CSG': { name:'Rezar (CAA)', team:'WL', clockIn:'08:00', clockOut:'', account:'caa_tl_rezar_32020226' },
  '32BDG_0224_NEW': { name:'Fajar (PAB)', team:'WL', clockIn:'07:45', clockOut:'17:00', account:'pab_tl_fajarfm_33091410' },
  '32CMH_0099_WL':  { name:'Gustira (NIM)', team:'WL', clockIn:'07:15', clockOut:'', account:'nim_tl_gustira_32053405' },
  '32SKB_0072_NEW': { name:'Rahmat (Intisel)', team:'WL', clockIn:'08:30', clockOut:'', account:'intisel_tl_rahmat_32051919' },
  '32TNG_0011_CSG': { name:'Deni (Intisel)', team:'WL', clockIn:'07:00', clockOut:'', account:'intisel_tl_deni_32053528' },
  '32BDG_0333_NEW': { name:'Sandi (INNOVIS)', team:'WL', clockIn:'08:15', clockOut:'', account:'inv_tl_sandi_32042901' },
  '32CJR_0043_NEW': { name:'Tester (Intisel)', team:'CSG', clockIn:'07:30', clockOut:'', account:'int_tl_test_00000001' },
  '32BDG_0673_NEW': { name:'Another (NIM)', team:'WL', clockIn:'08:00', clockOut:'', account:'nim_tl_another_00000002' },
  '32BDG_0298_NEW': { name:'Final (PAB)', team:'WL', clockIn:'07:45', clockOut:'', account:'pab_tl_final_00000003' }
};

var MOCK_MSGS = [
  { id:'m1', name:'sudah waktunya ganti ke ira', teamLeader:'Agung', timestamp:new Date().toISOString() },
  { id:'m2', name:'aku adalah surge', teamLeader:'Agung', timestamp:new Date().toISOString() },
  { id:'m3', name:'otw cico', teamLeader:'Rezar', timestamp:new Date().toISOString() }
];

// Intercept fetch() — return mock data based on URL type param
window.fetch = function(url, options) {
  var urlStr = (typeof url === 'string') ? url : (url && url.url ? url.url : '');
  var isPost = options && options.method === 'POST';
  var type = '';
  if (!isPost) {
    var m = urlStr.match(/[?&]type=([^&]+)/);
    if (m) type = m[1];
  } else {
    try { var body = JSON.parse(options.body || '{}'); type = body.type || ''; } catch(e) {}
  }
  var data;
  if (type === 'tracker') data = MOCK_TRACKER;
  else if (type === 'plan') data = MOCK_PLAN;
  else if (type === 'planDates') data = ['2026-09-04', '2026-09-03', '2026-09-02', '2026-09-01'];
  else if (type === 'activeDate') data = { date: '2026-09-04' };
  else if (type === 'planRoster') data = { added: [], removed: [] };
  else if (type === 'planRosterEdits') data = {};
  else if (type === 'clockin') data = MOCK_CLOCKIN;
  else if (type === 'clockinSubmit') data = { status: 'ok' };
  else data = MOCK_MSGS;

  var res = new Response(JSON.stringify(data), { status: 200, headers: { 'Content-Type': 'application/json' } });
  res.ok = true;
  return Promise.resolve(res);
};

</script>

"@

$output = $beforeHead + $mocks + $afterHead
$output | Out-File -FilePath "$PSScriptRoot\..\index-test.html" -Encoding UTF8 -NoNewline

Write-Host "[build-test] Done. Generated index-test.html ($(($output.Length / 1KB).ToString('N0')) KB)"
Write-Host "Open index-test.html in your browser to test."
