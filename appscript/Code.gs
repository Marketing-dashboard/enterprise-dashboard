// ============================================================
// Enterprise Daily Tracker Dashboard — Auto-Updater
// Install in Google Sheet: Extensions > Apps Script
// 1. Paste this code, save, run createTriggers() once
// 2. Set REPO_OWNER, REPO_NAME below
// 3. Run setGitHubToken('ghp_yourToken') once to store token
// ============================================================

var REPO_OWNER = 'Marketing-dashboard';
var REPO_NAME  = 'enterprise-dashboard';
var FILE_PATH  = 'index.html';

var JLR_TRIGGER_ALIASES = ['land rover', 'jaguar'];

function setGitHubToken(token) {
  PropertiesService.getScriptProperties().setProperty('GITHUB_TOKEN', token);
  Logger.log('Token saved.');
}

// ── ENTRY POINT ───────────────────────────────────────────────
function updateDashboard() {
  var token = PropertiesService.getScriptProperties().getProperty('GITHUB_TOKEN');
  if (!token) throw new Error('GitHub token not set. Run setGitHubToken("ghp_...") first.');

  var ss = SpreadsheetApp.getActiveSpreadsheet();

  // Find sheets
  var pannelSheet = null, trigSheet = null;
  var sheets = ss.getSheets();
  for (var i = 0; i < sheets.length; i++) {
    var name = sheets[i].getName().toLowerCase();
    if (name.indexOf('pannel') !== -1 || name.indexOf('panel') !== -1) pannelSheet = sheets[i];
    if (name.indexOf('trigger') !== -1) trigSheet = sheets[i];
  }
  if (!pannelSheet || !trigSheet) throw new Error('Could not find Raw_Pannel or Raw_triggers sheet.');

  var pannelData = pannelSheet.getDataRange().getValues();
  var trigData   = trigSheet.getDataRange().getValues();

  // ── Index pannel headers ─────────────────────────────────────
  var ph = pannelData[0];
  var pIdx = {
    month:  indexof(ph, 'Month'),
    day:    indexof(ph, 'Day'),
    cost:   indexof(ph, 'Cost'),
    leads:  indexof(ph, 'Leads'),
    brand:  indexof(ph, 'Brand'),
    model:  indexof(ph, 'Model'),
    buType: indexof(ph, 'BU Type')
  };

  // ── Index trigger headers ────────────────────────────────────
  var th = trigData[0];
  var tIdx = {
    date:     indexof(th, 'Date'),
    brand:    indexof(th, 'brand'),
    model:    indexof(th, 'model'),
    triggered:indexof(th, 'Triggered'),
    inList:   indexof(th, 'Triggered in List_ID')
  };

  // ── Build trigger map: normDate|modelLower → total ──────────
  var trigMap = {};
  for (var r = 1; r < trigData.length; r++) {
    var row    = trigData[r];
    var date   = normDate(row[tIdx.date]);
    var brand  = String(row[tIdx.brand]  || '').toLowerCase().trim();
    var model  = String(row[tIdx.model]  || '').toLowerCase().trim();
    var trig   = parseNum(row[tIdx.triggered]);
    var inList = parseNum(row[tIdx.inList]);
    var isPlus = JLR_TRIGGER_ALIASES.indexOf(brand) !== -1 || brand === 'citroen';
    var total  = isPlus ? trig + inList : trig;
    if (!date || !model) continue;
    var key = date + '|' + model;
    trigMap[key] = (trigMap[key] || 0) + total;
  }

  // ── Aggregate pannel: month|brand|model|buType|day ──────────
  var rowMap = {};
  for (var r = 1; r < pannelData.length; r++) {
    var row    = pannelData[r];
    var month  = String(row[pIdx.month]  || '').trim();
    var day    = normDate(row[pIdx.day]);
    var brand  = String(row[pIdx.brand]  || '').trim();
    var model  = String(row[pIdx.model]  || '').trim();
    var buType = String(row[pIdx.buType] || '').trim();
    var cost   = parseNum(row[pIdx.cost]);
    var leads  = parseNum(row[pIdx.leads]);
    if (!month || !brand || !model) continue;
    var key = month + '|' + brand + '|' + model + '|' + buType + '|' + day;
    if (!rowMap[key]) rowMap[key] = { month: month, day: day, brand: brand, model: model, buType: buType, spends: 0, leads: 0 };
    rowMap[key].spends += cost;
    rowMap[key].leads  += leads;
  }

  // ── Attach triggered leads & build final rows ────────────────
  var dataRows = [];
  for (var k in rowMap) {
    var row  = rowMap[k];
    var mk   = row.day + '|' + row.model.toLowerCase();
    var trig = Math.round((trigMap[mk] || 0) * 100) / 100;
    dataRows.push({
      month:     row.month,
      day:       row.day,
      brand:     row.brand,
      model:     row.model,
      buType:    row.buType,
      spends:    Math.round(row.spends * 100) / 100,
      leads:     Math.round(row.leads  * 100) / 100,
      triggered: trig
    });
  }

  // Sort: brand → model → day
  dataRows.sort(function(a, b) {
    return a.brand.localeCompare(b.brand) || a.model.localeCompare(b.model) || a.day.localeCompare(b.day);
  });

  // ── Serialize to JS ──────────────────────────────────────────
  var lines = dataRows.map(function(r) {
    return '{month:' + jstr(r.month) + ',day:' + jstr(r.day) +
           ',brand:' + jstr(r.brand) + ',model:' + jstr(r.model) +
           ',buType:' + jstr(r.buType) + ',spends:' + r.spends +
           ',leads:' + r.leads + ',triggered:' + r.triggered + '}';
  });
  var newDataRaw = 'var DATA_RAW = [\n' + lines.join(',\n') + '\n];';

  // ── Push to GitHub ───────────────────────────────────────────
  pushToGitHub(token, newDataRaw);
  Logger.log('Done. ' + dataRows.length + ' rows pushed to GitHub.');
}

// ── GITHUB HELPER ─────────────────────────────────────────────
function pushToGitHub(token, newDataRaw) {
  var apiUrl = 'https://api.github.com/repos/' + REPO_OWNER + '/' + REPO_NAME + '/contents/' + FILE_PATH;

  var getResp = UrlFetchApp.fetch(apiUrl, {
    headers: { 'Authorization': 'token ' + token, 'Accept': 'application/vnd.github.v3+json' },
    muteHttpExceptions: true
  });
  if (getResp.getResponseCode() !== 200)
    throw new Error('GitHub GET failed: ' + getResp.getContentText());

  var fileInfo = JSON.parse(getResp.getContentText());
  var sha      = fileInfo.sha;
  var b64      = fileInfo.content.replace(/\n/g, '');
  var current  = Utilities.newBlob(Utilities.base64Decode(b64)).getDataAsString();

  var updated = current.replace(/var DATA_RAW = \[[\s\S]*?\];/, newDataRaw);
  if (updated === current) { Logger.log('No data changes — skipping push.'); return; }

  var payload = JSON.stringify({
    message: 'Auto-update DATA_RAW [' + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm') + ']',
    content: Utilities.base64Encode(Utilities.newBlob(updated).getBytes()),
    sha: sha, branch: 'main'
  });

  var putResp = UrlFetchApp.fetch(apiUrl, {
    method: 'PUT',
    headers: { 'Authorization': 'token ' + token, 'Content-Type': 'application/json', 'Accept': 'application/vnd.github.v3+json' },
    payload: payload, muteHttpExceptions: true
  });
  if (putResp.getResponseCode() !== 200 && putResp.getResponseCode() !== 201)
    throw new Error('GitHub PUT failed (' + putResp.getResponseCode() + '): ' + putResp.getContentText());
  Logger.log('Pushed: ' + putResp.getResponseCode());
}

// ── TRIGGERS ──────────────────────────────────────────────────
function onSheetEdit(e) {
  // Only run if Raw_Pannel or Raw_triggers was edited
  var name = e && e.source && e.source.getActiveSheet() ? e.source.getActiveSheet().getName().toLowerCase() : '';
  if (name.indexOf('pannel') !== -1 || name.indexOf('panel') !== -1 || name.indexOf('trigger') !== -1) {
    updateDashboard();
  }
}

function createTriggers() {
  // Remove old triggers
  ScriptApp.getProjectTriggers().forEach(function(t) { ScriptApp.deleteTrigger(t); });

  // On edit trigger (fires whenever the sheet is edited)
  ScriptApp.newTrigger('onSheetEdit').forSpreadsheet(SpreadsheetApp.getActiveSpreadsheet()).onEdit().create();

  // Daily refresh at 7 AM as backup
  ScriptApp.newTrigger('updateDashboard').timeBased().everyDays(1).atHour(7).create();

  Logger.log('Triggers created: onEdit + daily 7AM.');
}

// ── UTILITIES ─────────────────────────────────────────────────
function normDate(val) {
  if (!val) return '';
  var tz = Session.getScriptTimeZone();
  if (val instanceof Date) return Utilities.formatDate(val, tz, 'yyyy-MM-dd');
  var s = String(val).trim();
  if (/^\d{4}-\d{2}-\d{2}$/.test(s)) return s;
  // "1-Jul-2026"
  var mon = {Jan:'01',Feb:'02',Mar:'03',Apr:'04',May:'05',Jun:'06',Jul:'07',Aug:'08',Sep:'09',Oct:'10',Nov:'11',Dec:'12'};
  var p = s.split('-');
  if (p.length === 3 && isNaN(p[1])) return p[2] + '-' + (mon[p[1]] || '00') + '-' + p[0].padStart(2,'0');
  // Try parsing
  var d = new Date(s);
  if (!isNaN(d)) return Utilities.formatDate(d, tz, 'yyyy-MM-dd');
  return s;
}

function parseNum(v) { var n = parseFloat(v); return isNaN(n) ? 0 : n; }
function jstr(s)     { return JSON.stringify(String(s || '')); }
function indexof(arr, name) {
  var n = name.toLowerCase();
  for (var i = 0; i < arr.length; i++) if (String(arr[i]).toLowerCase().trim() === n) return i;
  return -1;
}
