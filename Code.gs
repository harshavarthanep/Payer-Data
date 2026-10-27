/**
 * Enrollment Export — Google Apps Script server side.
 *
 * Setup (once):
 *   1. script.google.com → New project (or Extensions → Apps Script from a Google Sheet).
 *   2. Replace Code.gs with this file. Add an HTML file named "Index" and paste Index.html into it.
 *   3. Deploy → New deployment → Web app. Execute as: "User accessing the web app".
 *      Who has access: your Google Workspace domain. Share the web app URL with your team.
 *   4. Each person opens the URL, signs in under Connection settings once, then pastes enrollment links.
 *
 * Each person's platform token is stored in their own Apps Script user properties. Treat it like a password.
 */
var DEFAULT_BASE = 'https://prod-backend.withassured.com';
var XLSX_MIME = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';

function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('Enrollment Overview Export')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

/* When the script is bound to a Google Sheet, add a menu that opens the exporter in a dialog. */
function onOpen() {
  try {
    SpreadsheetApp.getUi().createMenu('Enrollment Export').addItem('Open exporter', 'openExporter').addToUi();
  } catch (e) { /* standalone script: no spreadsheet UI */ }
}
function openExporter() {
  var html = HtmlService.createHtmlOutputFromFile('Index').setWidth(1200).setHeight(820);
  SpreadsheetApp.getUi().showModelessDialog(html, 'Enrollment Export');
}

function props_() { return PropertiesService.getUserProperties(); }
function base_() { return (props_().getProperty('ASSURED_BASE') || DEFAULT_BASE).replace(/\/+$/, ''); }

function getSettings() {
  return { base: base_(), hasToken: !!props_().getProperty('ASSURED_TOKEN') };
}

function saveSettings(s) {
  s = s || {};
  if (s.base) {
    if (!/^https:\/\/[^\s/]+/.test(s.base)) throw new Error('The API address must start with https://');
    props_().setProperty('ASSURED_BASE', String(s.base).trim());
  }
  if (s.token) props_().setProperty('ASSURED_TOKEN', String(s.token).trim().replace(/^Bearer\s+/i, ''));
  return getSettings();
}

function forgetToken() { props_().deleteProperty('ASSURED_TOKEN'); return getSettings(); }

function signIn(email, password) {
  var res = UrlFetchApp.fetch(base_() + '/api/v1/users/login/', {
    method: 'post', contentType: 'application/json', muteHttpExceptions: true,
    payload: JSON.stringify({ email: email, password: password })
  });
  var j = {};
  try { j = JSON.parse(res.getContentText()); } catch (e) {}
  var jwt = j && j.data && j.data.jwt;
  var tok = typeof jwt === 'string' ? jwt : (jwt && (jwt.access || jwt.token || jwt.access_token));
  if (res.getResponseCode() >= 300 || !tok) {
    throw new Error('Sign-in failed (' + res.getResponseCode() + '). Check your email and password, or paste a token instead.');
  }
  props_().setProperty('ASSURED_TOKEN', tok);
  return getSettings();
}

/**
 * GET several API paths in parallel. Returns [{status, json}] in the same order; empty paths return null.
 * Only paths under /api/ on the configured platform address are allowed.
 */
function apiGetAll(paths) {
  var token = props_().getProperty('ASSURED_TOKEN');
  if (!token) throw new Error('NO_TOKEN');
  var base = base_();
  var idx = [], reqs = [];
  (paths || []).forEach(function (p, i) {
    if (!p) return;
    if (!/^\/api\/[\w\-./?=&%]+$/.test(p)) throw new Error('Blocked path: ' + p);
    idx.push(i);
    reqs.push({ url: base + p, method: 'get', muteHttpExceptions: true, headers: { Authorization: 'Bearer ' + token, Accept: 'application/json' } });
  });
  var out = (paths || []).map(function () { return null; });
  if (!reqs.length) return out;
  var res = UrlFetchApp.fetchAll(reqs);
  res.forEach(function (r, k) {
    var j = null;
    try { j = JSON.parse(r.getContentText()); } catch (e) {}
    out[idx[k]] = { status: r.getResponseCode(), json: j };
  });
  return out;
}

function saveXlsxToDrive(filename, base64) {
  var blob = Utilities.newBlob(Utilities.base64Decode(base64), XLSX_MIME, filename);
  return DriveApp.createFile(blob).getUrl();
}

/* sheets: [{name, rows: [[...], ...]}] — writes every value as text so ZIP codes, NPIs and IDs keep leading zeros. */
function writeToSheet(title, sheets) {
  var ss = SpreadsheetApp.create(title);
  var first = ss.getSheets()[0];
  sheets.forEach(function (s, i) {
    var sh = i === 0 ? first.setName(s.name) : ss.insertSheet(s.name);
    var rows = s.rows || [];
    if (!rows.length) return;
    var w = rows.reduce(function (m, r) { return Math.max(m, r.length); }, 0);
    var data = rows.map(function (r) { var x = r.slice(); while (x.length < w) x.push(''); return x.map(function (v) { return v == null ? '' : String(v); }); });
    var range = sh.getRange(1, 1, data.length, w);
    range.setNumberFormat('@');
    range.setValues(data);
    sh.getRange(1, 1, 1, w).setFontWeight('bold');
    sh.setFrozenRows(1);
    try { sh.autoResizeColumns(1, w); } catch (e) {}
  });
  return ss.getUrl();
}
