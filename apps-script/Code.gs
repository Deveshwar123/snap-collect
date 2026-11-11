/**
 * snap-collect backend (Google Apps Script web app).
 * Run setup() once from the editor to authorize and create the Drive folder + Sheet.
 */
var FOLDER_NAME = 'Snap Dataset (snap-collect)';
var SHEET_NAME = 'Snap Dataset — submissions';
var HEADER = ['received_at', 'submission_id', 'part', 'name', 'n_clips', 'snap_check_flags', 'ua', 'lang', 'tz', 'folder_url', 'round_points', 'session_points', 'show_on_leaderboard', 'match_scores'];
var MAX_ROUND_POINTS = 11;  // 3 attempts x 2 points + one 5-point streak bonus
var MAX_ROUNDS = 6;
var MAX_CLIP_BYTES = 1.5 * 1024 * 1024;
var MAX_CLIPS = 6;

function setup() {
  var r = getResources_();
  Logger.log('Folder: ' + r.folder.getUrl());
  Logger.log('Sheet: ' + r.ss.getUrl());
}

function doGet(e) {
  if (e && e.parameter && e.parameter.action === 'leaderboard') return json_(leaderboard_());
  return json_({ ok: true, service: 'snap-collect' });
}

// Top 10 sessions by points, opted-in players only. Points are computed in the
// browser, so the server caps them per round and counts at most MAX_ROUNDS rounds.
function leaderboard_() {
  var cache = CacheService.getScriptCache();
  var hit = cache.get('leaderboard');
  if (hit) return JSON.parse(hit);
  var sheet = getResources_().sheet;
  var rows = sheet.getDataRange().getValues();
  var head = rows.shift() || [];
  var c = {}; head.forEach(function (h, i) { c[h] = i; });
  if (c.round_points === undefined) return { ok: true, top: [], updated_at: new Date().toISOString() };
  var by = {};
  rows.forEach(function (r) {
    var id = r[c.submission_id];
    if (!id || r[c.show_on_leaderboard] !== true) return;
    var s = by[id] || (by[id] = { name: String(r[c.name] || '').replace(/^'/, '').slice(0, 24), parts: {} });
    var part = Number(r[c.part]) || 1;
    if (part > MAX_ROUNDS) return;
    // Snaps are uploaded one at a time, so a round arrives as several rows: sum them, capped per round.
    var add = Math.max(0, Number(r[c.round_points]) || 0);
    s.parts[part] = Math.min(MAX_ROUND_POINTS, (s.parts[part] || 0) + add);
  });
  var top = Object.keys(by).map(function (k) {
    var pts = 0; for (var p in by[k].parts) pts += by[k].parts[p];
    return { name: by[k].name, points: pts };
  })
    .filter(function (x) { return x.points > 0; })
    .sort(function (a, b) { return b.points - a.points; }).slice(0, 10);
  var out = { ok: true, top: top, updated_at: new Date().toISOString() };
  cache.put('leaderboard', JSON.stringify(out), 20);
  return out;
}

function doPost(e) {
  var lock = LockService.getScriptLock();
  var locked = false;
  try {
    var d = JSON.parse(e.postData.contents);
    if (d.website) return json_({ ok: false, error: 'rejected' });
    if (d.consent !== true || d.age_ok !== true) return json_({ ok: false, error: 'consent required' });

    var id = String(d.submission_id || '').replace(/[^A-Za-z0-9-]/g, '').slice(0, 64);
    if (!id) return json_({ ok: false, error: 'missing submission_id' });
    var name = String(d.name || '').trim().slice(0, 40);
    if (/^[=+\-@]/.test(name)) name = "'" + name;  // never let Sheets read a name as a formula
    if (!name) return json_({ ok: false, error: 'missing name' });
    var clips = d.clips;
    if (!Array.isArray(clips) || clips.length < 1 || clips.length > MAX_CLIPS) {
      return json_({ ok: false, error: 'bad clip count' });
    }

    // Validate everything before writing anything.
    var decoded = [];
    for (var i = 0; i < clips.length; i++) {
      var c = clips[i];
      var b64 = String(c.base64 || '');
      if (!b64) return json_({ ok: false, error: 'empty clip' });
      var pad = (b64.slice(-2) === '==') ? 2 : (b64.slice(-1) === '=' ? 1 : 0);
      var approx = Math.floor(b64.length * 3 / 4) - pad;
      if (approx > MAX_CLIP_BYTES) return json_({ ok: false, error: 'clip too large' });
      decoded.push({ index: parseInt(c.index, 10) || (i + 1), mime: String(c.mime || ''), b64: b64, check: c.snap_check });
    }

    lock.waitLock(30000);
    locked = true;

    var r = getResources_();
    var subFolders = r.folder.getFoldersByName(id);
    var sub = subFolders.hasNext() ? subFolders.next() : r.folder.createFolder(id);
    var flags = [];
    decoded.forEach(function (c) {
      var bytes = Utilities.base64Decode(c.b64);
      if (bytes.length > MAX_CLIP_BYTES) throw new Error('clip too large');
      var blob = Utilities.newBlob(bytes, c.mime.split(';')[0] || 'application/octet-stream',
        id + '_' + c.index + '.' + ext_(c.mime));
      sub.createFile(blob);
      flags.push(c.index + ':' + (c.check === false ? 'false' : 'true'));
    });

    ensureHeader_(r.sheet);
    var scores = clips.map(function (c) { return (typeof c.match_score === 'number') ? Math.round(c.match_score) : ''; }).join(',');
    r.sheet.appendRow([
      new Date(), id, d.part || 1, name, decoded.length, flags.join(','),
      String(d.ua || '').slice(0, 300), String(d.lang || '').slice(0, 20),
      String(d.tz || '').slice(0, 60), sub.getUrl(),
      Math.max(0, Math.min(MAX_ROUND_POINTS, parseInt(d.round_points, 10) || 0)),
      Math.max(0, parseInt(d.session_points, 10) || 0),
      d.show_on_leaderboard === true, scores
    ]);
    CacheService.getScriptCache().remove('leaderboard');
    return json_({ ok: true });
  } catch (err) {
    return json_({ ok: false, error: String(err && err.message || err) });
  } finally {
    if (locked) lock.releaseLock();
  }
}

// Sheets created before the game version have the old 10-column header; extend it once.
function ensureHeader_(sheet) {
  var n = sheet.getLastColumn();
  if (n >= HEADER.length) return;
  sheet.getRange(1, 1, 1, HEADER.length).setValues([HEADER]);
}

function ext_(mime) {
  var m = String(mime).toLowerCase();
  if (m.indexOf('webm') > -1) return 'webm';
  if (m.indexOf('m4a') > -1 || m.indexOf('x-m4a') > -1 || m.indexOf('aac') > -1) return 'm4a';
  if (m.indexOf('mp4') > -1 || m.indexOf('quicktime') > -1) return 'mp4';
  if (m.indexOf('ogg') > -1) return 'ogg';
  if (m.indexOf('wav') > -1) return 'wav';
  if (m.indexOf('mpeg') > -1 || m.indexOf('mp3') > -1) return 'mp3';
  return 'bin';
}

function getResources_() {
  var props = PropertiesService.getScriptProperties();
  var folderId = props.getProperty('FOLDER_ID');
  var sheetId = props.getProperty('SHEET_ID');
  var folder, ss;
  if (folderId) {
    try { folder = DriveApp.getFolderById(folderId); } catch (e) { folder = null; }
  }
  if (!folder) {
    folder = DriveApp.createFolder(FOLDER_NAME);
    props.setProperty('FOLDER_ID', folder.getId());
  }
  if (sheetId) {
    try { ss = SpreadsheetApp.openById(sheetId); } catch (e) { ss = null; }
  }
  if (!ss) {
    ss = SpreadsheetApp.create(SHEET_NAME);
    props.setProperty('SHEET_ID', ss.getId());
    var sh = ss.getSheets()[0];
    sh.appendRow(HEADER);
    sh.setFrozenRows(1);
    DriveApp.getFileById(ss.getId()).moveTo(folder);
  }
  return { folder: folder, ss: ss, sheet: ss.getSheets()[0] };
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
