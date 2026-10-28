/**
 * Enrollment Export — Google Apps Script server side.
 *
 * Setup (once):
 *   1. script.google.com → New project.
 *   2. Replace Code.gs with this file. Add an HTML file named exactly "Index" and paste Index.html into it.
 *   3. Deploy → New deployment → type "Web app".
 *        Execute as: "User accessing the web app"   (each person signs in with their own Assured account)
 *        Who has access: "Anyone within <your domain>"
 *      After any change: Deploy → Manage deployments → Edit (pencil) → Version: New version → Deploy.
 *   4. Each person opens the web app URL, approves the Google permissions once, and signs in with their
 *      Assured email and password once. "Keep me signed in" keeps them signed in on later visits.
 *
 * Sign-in details are stored only in each person's own Apps Script user properties.
 */
var BASE = 'https://prod-backend.withassured.com'; // Assured platform API
var XLSX_MIME = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';

function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('Enrollment Export')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function props_() { return PropertiesService.getUserProperties(); }

function post_(path, data, auth) {
  var opt = { method: 'post', contentType: 'application/json', muteHttpExceptions: true, payload: JSON.stringify(data), headers: { Accept: 'application/json' } };
  if (auth) opt.headers.Authorization = 'Bearer ' + auth;
  var r = UrlFetchApp.fetch(BASE + path, opt), t = r.getContentText(), j = null;
  try { j = JSON.parse(t); } catch (e) { j = t; }
  return { status: r.getResponseCode(), json: j };
}
function message_(j) {
  if (!j) return '';
  if (typeof j === 'string') return j.replace(/<[^>]+>/g, ' ').replace(/\s+/g, ' ').trim().slice(0, 200);
  var e = j.errors && j.errors[0];
  var m = (e && (e.detail || e.message)) || j.detail || j.msg || j.message || j.error || (j.non_field_errors && j.non_field_errors[0]) || '';
  if (!m) { for (var k in j) { if (Array.isArray(j[k]) && typeof j[k][0] === 'string') { m = k + ': ' + j[k][0]; break; } } }
  return String(m).replace(/\.+$/, '').slice(0, 300);
}
function tokens_(j) {
  if (!j || typeof j !== 'object') return null;
  var jwt = (j.data && j.data.jwt) || j.jwt || (j.access ? j : null);
  if (!jwt) return null;
  if (typeof jwt === 'string') return { access: jwt, refresh: '' };
  return jwt.access ? { access: jwt.access, refresh: jwt.refresh || '' } : null;
}
function wantsCode_(j) {
  var flag = j && typeof j === 'object' && (j.mfa_required || j.otp_required || (j.data && (j.data.mfa_required || j.data.otp_required || j.data.mfa_enabled)));
  return !!flag || /\b(otp|one[- ]time|verification code|mfa|two[- ]?factor|authenticator|code sent|enter the code)\b/i.test(message_(j));
}
function loginBody_(email, pass, code) {
  var b = { email: String(email).trim(), password: pass };
  if (code) { code = String(code).trim(); b.otp = code; b.code = code; b.mfa_code = code; b.otp_code = code; b.one_time_code = code; }
  return b;
}

function getSession() {
  var p = props_(), u = p.getProperty('ASSURED_USER');
  return { signedIn: !!p.getProperty('ASSURED_TOKEN'), user: u ? JSON.parse(u) : null };
}

/* Returns {user} when signed in, or {needCode, message} when the platform asks for a verification code. */
function signIn(email, password, code, keep) {
  var r = post_('/api/v1/users/login/', loginBody_(email, password, code));
  var t = tokens_(r.json);
  if (r.status >= 200 && r.status < 300 && t) {
    var d = (r.json && r.json.data) || {};
    var user = { name: [d.first_name, d.last_name].filter(String).join(' ') || email, email: d.email || email };
    var p = props_();
    p.setProperty('ASSURED_TOKEN', t.access);
    if (t.refresh) p.setProperty('ASSURED_REFRESH', t.refresh); else p.deleteProperty('ASSURED_REFRESH');
    p.setProperty('ASSURED_USER', JSON.stringify(user));
    if (keep && !code) { p.setProperty('ASSURED_EMAIL', String(email).trim()); p.setProperty('ASSURED_PASS', password); }
    else { p.deleteProperty('ASSURED_EMAIL'); p.deleteProperty('ASSURED_PASS'); }
    return { user: user };
  }
  if (wantsCode_(r.json)) return { needCode: true, message: message_(r.json) || 'Enter the verification code sent to you.' };
  throw new Error((message_(r.json) || 'Sign-in failed') + ' (' + r.status + ').');
}

/* Ask the platform to email a set-password link (for accounts that sign in without a password). */
function resetPassword(email) {
  var r = post_('/api/v1/users/password-reset/', { email: String(email).trim() });
  if (r.status >= 200 && r.status < 300) return true;
  throw new Error((message_(r.json) || 'The email could not be sent') + ' (' + r.status + ').');
}

/* Sign in with the session copied from the platform by the Assured sign-in bookmark. */
function useSession(access, refresh, keep) {
  var p = props_();
  var r = UrlFetchApp.fetch(BASE + '/api/v1/users/logged-in-user-details/', { muteHttpExceptions: true, headers: { Authorization: 'Bearer ' + access, Accept: 'application/json' } });
  if (r.getResponseCode() !== 200) throw new Error('The platform did not accept that sign-in (' + r.getResponseCode() + '). Click the bookmark on the platform again.');
  var j = {}; try { j = JSON.parse(r.getContentText()); } catch (e) {}
  var u = j.data || j;
  var user = { name: [u.first_name, u.last_name].filter(String).join(' ') || u.email || 'Signed in', email: u.email || '' };
  p.setProperty('ASSURED_TOKEN', access);
  if (refresh) p.setProperty('ASSURED_REFRESH', refresh); else p.deleteProperty('ASSURED_REFRESH');
  p.setProperty('ASSURED_USER', JSON.stringify(user));
  p.deleteProperty('ASSURED_EMAIL'); p.deleteProperty('ASSURED_PASS');
  return { user: user };
}

function signOut() {
  ['ASSURED_TOKEN', 'ASSURED_REFRESH', 'ASSURED_USER', 'ASSURED_EMAIL', 'ASSURED_PASS'].forEach(function (k) { props_().deleteProperty(k); });
  return true;
}

/* Keep the session alive: renew with the refresh token, else sign in again with the saved email and password. */
function renew_() {
  var p = props_(), refresh = p.getProperty('ASSURED_REFRESH');
  if (refresh) {
    var r = post_('/api/v1/users/token/refresh/', { refresh: refresh }, p.getProperty('ASSURED_TOKEN'));
    if (r.status === 200 && r.json && r.json.access) {
      p.setProperty('ASSURED_TOKEN', r.json.access);
      if (r.json.refresh) p.setProperty('ASSURED_REFRESH', r.json.refresh);
      return true;
    }
  }
  var email = p.getProperty('ASSURED_EMAIL'), pass = p.getProperty('ASSURED_PASS');
  if (email && pass) {
    try { var s = signIn(email, pass, '', true); return !!(s && s.user); } catch (e) { return false; }
  }
  return false;
}

function fetchPaths_(paths) {
  var token = props_().getProperty('ASSURED_TOKEN'), idx = [], reqs = [];
  paths.forEach(function (p, i) {
    if (!p) return;
    if (!/^\/api\/[\w\-./?=&%]+$/.test(p)) throw new Error('Blocked path: ' + p);
    idx.push(i);
    reqs.push({ url: BASE + p, method: 'get', muteHttpExceptions: true, headers: { Authorization: 'Bearer ' + token, Accept: 'application/json' } });
  });
  var out = paths.map(function () { return null; });
  if (!reqs.length) return out;
  UrlFetchApp.fetchAll(reqs).forEach(function (r, k) {
    var j = null;
    try { j = JSON.parse(r.getContentText()); } catch (e) {}
    out[idx[k]] = { status: r.getResponseCode(), json: j };
  });
  return out;
}

/* GET several API paths in parallel. Returns [{status, json}] in order, or {signedOut: true} when the session cannot be renewed. */
function apiGetAll(paths) {
  paths = paths || [];
  if (!props_().getProperty('ASSURED_TOKEN') && !renew_()) return { signedOut: true };
  var out = fetchPaths_(paths);
  if (out.some(function (r) { return r && r.status === 401; })) {
    if (!renew_()) { props_().deleteProperty('ASSURED_TOKEN'); return { signedOut: true }; }
    out = fetchPaths_(paths);
  }
  return out;
}

function saveXlsxToDrive(filename, base64) {
  return DriveApp.createFile(Utilities.newBlob(Utilities.base64Decode(base64), XLSX_MIME, filename)).getUrl();
}

/* sheets: [{name, rows}] — every value is written as text so ZIP codes, NPIs and IDs keep leading zeros. */
function writeToSheet(title, sheets) {
  var ss = SpreadsheetApp.create(title), first = ss.getSheets()[0], used = {};
  sheets.forEach(function (s, i) {
    var name = String(s.name || ('Sheet ' + (i + 1))).replace(/[\[\]*?:\/\\]/g, ' ').slice(0, 99), k = 2;
    while (used[name]) name = String(s.name).slice(0, 95) + ' ' + (k++);
    used[name] = 1;
    var sh = i === 0 ? first.setName(name) : ss.insertSheet(name), rows = s.rows || [];
    if (!rows.length) return;
    var w = rows.reduce(function (m, r) { return Math.max(m, r.length); }, 1);
    var data = rows.map(function (r) { var x = r.slice(); while (x.length < w) x.push(''); return x.map(function (v) { return v == null ? '' : String(v); }); });
    var range = sh.getRange(1, 1, data.length, w);
    range.setNumberFormat('@');
    range.setValues(data);
    sh.getRange(1, 1, 1, w).setFontWeight('bold').setBackground('#e8eff0');
    sh.setFrozenRows(1);
    try { sh.autoResizeColumns(1, Math.min(w, 60)); } catch (e) {}
  });
  return ss.getUrl();
}
