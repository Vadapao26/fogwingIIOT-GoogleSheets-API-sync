/**
 * FOGWING -> GOOGLE SHEETS  |  Single-file version with menu, sidebar, status & sync log
 * Files in the Apps Script project:  Code.gs (this)  +  Sidebar.html
 *
 * CREDENTIALS ARE NEVER STORED IN THIS CODE.
 * Use the menu: Fogwing Sync > Set Credentials. They are saved in Script Properties.
 */

const CONFIG = {
  SHEET_NAME: 'Data',
  LOG_SHEET_NAME: 'Sync Log',
  BASE_URL: 'https://api.fogwing.net/api/v1/dstorage/getData/withApiKey',
  TIMEZONE: 'Asia/Kolkata',
  FIELDS: ['date', 'material_type', 'process_type', 'shift_date', 'shift_name',
           'shift_range', 'source', 'time', 'weight', 'batch_no', 'vehicle_no'],
  HEADERS: ['Date', 'Material Type', 'Process Type', 'Shift Date', 'Shift Name',
            'Shift Range', 'Source', 'Time', 'Weight', 'Batch No', 'Vehicle No'],
  LOG_HEADERS: ['Time', 'Mode', 'From', 'To', 'Fetched', 'Added', 'Skipped', 'Status', 'Error']
};

const AUTO = {
  HOUR: 23,        // ~11 PM
  NEAR_MINUTE: 30, // ~11:30 PM IST (Google runs it within about an hour of this)
  DAYS_BACK: 0     // 0 = today only. 1 = also re-check yesterday (duplicates are skipped)
};

/* ================= MENU ================= */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Fogwing Sync')
    .addItem('Open Fetch Panel', 'openPanel')
    .addItem('Fetch Today Now', 'fetchTodayNow')
    .addSeparator()
    .addItem('Check Status', 'checkStatus')
    .addItem('Test API Connection', 'testConnectionMenu')
    .addSeparator()
    .addItem('Set Credentials', 'setupCredentials')
    .addItem('Turn ON Daily Auto-Fetch', 'createDailyTrigger')
    .addItem('Turn OFF Daily Auto-Fetch', 'removeDailyTrigger')
    .addToUi();
}

function openPanel() {
  const html = HtmlService.createHtmlOutputFromFile('Sidebar').setTitle('Fogwing Sync');
  SpreadsheetApp.getUi().showSidebar(html);
}

/* ================= CREDENTIALS (Script Properties) ================= */
const CRED_KEYS = {
  apiKey: 'FOGWING_API_KEY',
  accountId: 'FOGWING_ACCOUNT_ID',
  edgeEui: 'FOGWING_EDGE_EUI'
};

function getCred_(name) {
  return PropertiesService.getScriptProperties().getProperty(CRED_KEYS[name]) || '';
}

function credsSet_() {
  return !!(getCred_('apiKey') && getCred_('accountId') && getCred_('edgeEui'));
}

// Menu: Fogwing Sync > Set Credentials (values typed in a popup, never written in code)
function setupCredentials() {
  const ui = SpreadsheetApp.getUi();
  const ask = (title, text) => {
    const r = ui.prompt(title, text, ui.ButtonSet.OK_CANCEL);
    if (r.getSelectedButton() !== ui.Button.OK) return null;
    return r.getResponseText().trim();
  };
  const apiKey = ask('Fogwing API key', 'Paste your Fogwing API key:');
  if (apiKey === null) return;
  const accountId = ask('Account ID', 'Enter your Fogwing account ID:');
  if (accountId === null) return;
  const edgeEui = ask('Edge EUI', 'Enter the device edge EUI:');
  if (edgeEui === null) return;
  if (!apiKey || !accountId || !edgeEui) {
    ui.alert('All three values are required. Nothing was saved.');
    return;
  }
  const props = PropertiesService.getScriptProperties();
  props.setProperty(CRED_KEYS.apiKey, apiKey);
  props.setProperty(CRED_KEYS.accountId, accountId);
  props.setProperty(CRED_KEYS.edgeEui, edgeEui);
  ui.alert('Credentials saved securely in Script Properties.');
}

/* ================= MENU ACTIONS ================= */
function fetchTodayNow() {
  const ui = SpreadsheetApp.getUi();
  const today = todayStr_(0);
  try {
    const r = syncRange_(today, today, 'Manual');
    ui.alert('Done (' + today + ')',
      'Fetched: ' + r.fetched + '\nNew rows added: ' + r.added +
      '\nDuplicates skipped: ' + r.skipped, ui.ButtonSet.OK);
  } catch (err) {
    ui.alert('Fetch failed: ' + err.message);
  }
}

function checkStatus() {
  const s = getStatus_();
  const lines = [
    'Auto-fetch: ' + (s.autoOn ? 'ON (around ' + AUTO.HOUR + ':' + AUTO.NEAR_MINUTE + ' IST daily)' : 'OFF'),
    'Credentials: ' + (s.keySaved ? 'Set' : 'NOT SET (use Set Credentials)'),
    'Total rows in Data: ' + s.totalRows,
    ''
  ];
  if (s.last) {
    lines.push('Last run: ' + s.last.time + ' (' + s.last.mode + ')');
    lines.push('Range: ' + s.last.from + ' to ' + s.last.to);
    lines.push('Result: ' + s.last.fetched + ' fetched, ' + s.last.added + ' added, ' + s.last.skipped + ' skipped');
    lines.push('Status: ' + s.last.status + (s.last.error ? '\nError: ' + s.last.error : ''));
  } else {
    lines.push('No runs logged yet.');
  }
  SpreadsheetApp.getUi().alert('Fogwing Sync Status', lines.join('\n'), SpreadsheetApp.getUi().ButtonSet.OK);
}

function testConnectionMenu() {
  const r = testConnection();
  SpreadsheetApp.getUi().alert('API Test', r.message, SpreadsheetApp.getUi().ButtonSet.OK);
}

/* ================= FUNCTIONS CALLED BY THE SIDEBAR ================= */
// startIso / endIso come from <input type="date"> as yyyy-mm-dd
function fetchFromPanel(startIso, endIso) {
  if (!startIso || !endIso) throw new Error('Please pick both dates.');
  if (startIso > endIso) throw new Error('Start date is after end date.');
  const start = isoToApi_(startIso);
  const end = isoToApi_(endIso);
  const r = syncRange_(start, end, 'Manual (panel)');
  return r;
}

function getStatusForPanel() {
  return getStatus_();
}

// Fetches today's data without writing anything to the sheet
function testConnection() {
  try {
    const today = todayStr_(0);
    const json = fetchJson_(today, today);
    const n = extractRecords_(json).length;
    return { ok: true, message: 'Connection OK. API returned ' + n + ' record(s) for ' + today + '. Nothing was written.' };
  } catch (err) {
    return { ok: false, message: 'Connection FAILED: ' + err.message };
  }
}

/* ================= AUTO FETCH ================= */
function autoSyncToday() {
  const start = todayStr_(AUTO.DAYS_BACK);
  const end = todayStr_(0);
  try {
    syncRange_(start, end, 'Auto');
  } catch (err) {
    // Already written to Sync Log by syncRange_. Email so a failure is noticed.
    MailApp.sendEmail(
      Session.getEffectiveUser().getEmail(),
      'Fogwing auto-sync failed',
      'Date range ' + start + ' to ' + end + '\nError: ' + err.message
    );
  }
}

function createDailyTrigger() {
  removeDailyTrigger(true);
  ScriptApp.newTrigger('autoSyncToday')
    .timeBased()
    .everyDays(1)
    .atHour(AUTO.HOUR)
    .nearMinute(AUTO.NEAR_MINUTE)
    .inTimezone(CONFIG.TIMEZONE)
    .create();
  notify_('Daily auto-fetch is ON (around ' + AUTO.HOUR + ':' + AUTO.NEAR_MINUTE + ' IST).');
}

function removeDailyTrigger(silent) {
  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction() === 'autoSyncToday') ScriptApp.deleteTrigger(t);
  });
  if (silent !== true) notify_('Daily auto-fetch is OFF.');
}

/* ================= CORE SYNC ================= */
function syncRange_(startDate, endDate, mode) {
  try {
    const r = doSync_(startDate, endDate);
    logRun_(mode, startDate, endDate, r.fetched, r.added, r.skipped, 'OK', '');
    return r;
  } catch (err) {
    logRun_(mode, startDate, endDate, 0, 0, 0, 'FAILED', err.message);
    throw err;
  }
}

function doSync_(startDate, endDate) {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(30000)) throw new Error('Another sync is running. Try again shortly.');
  try {
    const json = fetchJson_(startDate, endDate);
    const records = extractRecords_(json);
    const sheet = getDataSheet_();

    const lastRow = sheet.getLastRow();
    const existing = new Set();
    if (lastRow > 1) {
      sheet.getRange(2, 1, lastRow - 1, CONFIG.FIELDS.length)
        .getDisplayValues()
        .forEach(row => existing.add(makeKey_(row)));
    }

    const newRows = [];
    let skipped = 0;
    records.forEach(rec => {
      const row = CONFIG.FIELDS.map(f => (rec[f] === undefined || rec[f] === null) ? '' : rec[f]);
      const key = makeKey_(row);
      if (existing.has(key)) { skipped++; return; }
      existing.add(key);
      row[8] = Number(row[8]);
      newRows.push(row);
    });

    if (newRows.length) {
      const startRow = sheet.getLastRow() + 1;
      sheet.getRange(startRow, 1, newRows.length, CONFIG.FIELDS.length).setNumberFormat('@');
      sheet.getRange(startRow, 9, newRows.length, 1).setNumberFormat('0.0##'); // weight numeric
      sheet.getRange(startRow, 1, newRows.length, CONFIG.FIELDS.length).setValues(newRows);
    }
    return { fetched: records.length, added: newRows.length, skipped: skipped };
  } finally {
    lock.releaseLock();
  }
}

function fetchJson_(startDate, endDate) {
  if (!credsSet_()) throw new Error('Credentials not set. Use Fogwing Sync > Set Credentials.');
  const key = getCred_('apiKey');

  const url = CONFIG.BASE_URL +
    '?apiKey=' + encodeURIComponent(key) +
    '&accountID=' + encodeURIComponent(getCred_('accountId')) +
    '&edgeEUI=' + encodeURIComponent(getCred_('edgeEui')) +
    '&startDate=' + startDate +
    '&endDate=' + endDate;

  const resp = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
  const code = resp.getResponseCode();
  if (code !== 200) {
    throw new Error('Fogwing API returned HTTP ' + code + ': ' + resp.getContentText().slice(0, 200));
  }
  let json;
  try {
    json = JSON.parse(resp.getContentText());
  } catch (e) {
    throw new Error('Response was not valid JSON.');
  }
  if (json && json.statusCode && Number(json.statusCode) !== 200) {
    throw new Error('Fogwing: ' + (json.message || '') + ' ' + (json.description || ''));
  }
  return json;
}

// Real response shape: { statusCode, message, data: [ { payload: {...fields...}, ... } ] }
function extractRecords_(json) {
  return findRecords_(json).map(r =>
    (r && r.payload && typeof r.payload === 'object') ? r.payload : r);
}

function findRecords_(json) {
  if (Array.isArray(json)) return json;
  if (json && typeof json === 'object') {
    if ('weight' in json && 'material_type' in json) return [json];
    const keys = Object.keys(json);
    for (const k of keys) if (Array.isArray(json[k])) return json[k];
    for (const k of keys) {
      if (json[k] && typeof json[k] === 'object') {
        const r = findRecords_(json[k]);
        if (r.length) return r;
      }
    }
  }
  return [];
}

function makeKey_(row) {
  return row.map((v, i) => i === 8 ? String(parseFloat(v)) : String(v).trim()).join('|');
}

/* ================= SHEETS: DATA + LOG ================= */
function getDataSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(CONFIG.SHEET_NAME);
  if (!sheet) sheet = ss.insertSheet(CONFIG.SHEET_NAME);
  if (sheet.getLastRow() === 0) {
    sheet.getRange(1, 1, 1, CONFIG.HEADERS.length)
      .setValues([CONFIG.HEADERS]).setFontWeight('bold');
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function getLogSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(CONFIG.LOG_SHEET_NAME);
  if (!sheet) sheet = ss.insertSheet(CONFIG.LOG_SHEET_NAME);
  if (sheet.getLastRow() === 0) {
    sheet.getRange(1, 1, 1, CONFIG.LOG_HEADERS.length)
      .setValues([CONFIG.LOG_HEADERS]).setFontWeight('bold');
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function logRun_(mode, from, to, fetched, added, skipped, status, error) {
  try {
    const sheet = getLogSheet_();
    const row = sheet.getLastRow() + 1;
    const time = Utilities.formatDate(new Date(), CONFIG.TIMEZONE, 'dd-MM-yyyy hh:mm a');
    sheet.getRange(row, 1, 1, 4).setNumberFormat('@');
    sheet.getRange(row, 1, 1, CONFIG.LOG_HEADERS.length)
      .setValues([[time, mode, from, to, fetched, added, skipped, status, error]]);
  } catch (e) {
    Logger.log('Could not write sync log: ' + e.message);
  }
}

/* ================= STATUS ================= */
function getStatus_() {
  const autoOn = ScriptApp.getProjectTriggers()
    .some(t => t.getHandlerFunction() === 'autoSyncToday');
  const keySaved = credsSet_();

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const dataSheet = ss.getSheetByName(CONFIG.SHEET_NAME);
  const totalRows = dataSheet ? Math.max(dataSheet.getLastRow() - 1, 0) : 0;

  let last = null;
  const logSheet = ss.getSheetByName(CONFIG.LOG_SHEET_NAME);
  if (logSheet && logSheet.getLastRow() > 1) {
    const v = logSheet.getRange(logSheet.getLastRow(), 1, 1, CONFIG.LOG_HEADERS.length).getDisplayValues()[0];
    last = { time: v[0], mode: v[1], from: v[2], to: v[3],
             fetched: v[4], added: v[5], skipped: v[6], status: v[7], error: v[8] };
  }
  return { autoOn: autoOn, keySaved: keySaved, totalRows: totalRows, last: last,
           autoTime: AUTO.HOUR + ':' + AUTO.NEAR_MINUTE + ' IST' };
}

/* ================= HELPERS ================= */
function todayStr_(daysBack) {
  const d = new Date(new Date().getTime() - (daysBack || 0) * 24 * 60 * 60 * 1000);
  return Utilities.formatDate(d, CONFIG.TIMEZONE, 'dd-MM-yyyy');
}

function isoToApi_(iso) { // yyyy-mm-dd -> dd-mm-yyyy
  const p = iso.split('-');
  return p[2] + '-' + p[1] + '-' + p[0];
}

function notify_(msg) {
  try { SpreadsheetApp.getUi().alert(msg); } catch (e) { Logger.log(msg); }
}
