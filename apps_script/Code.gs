/**
 * Receives the Rockridge photo reviews from the static site and writes them into this spreadsheet.  (version 2)
 * Container-bound script: open the Google Sheet -> Extensions -> Apps Script -> paste this file ->
 * Deploy -> Manage deployments -> pencil -> Version: New version -> Deploy.  First time: Deploy -> New deployment ->
 * Web app (Execute as: Me, Who has access: Anyone) and copy the Web app URL (ends in /exec).
 *
 * One row per (reviewer, photo_id). Sending the same photo again updates its row instead of adding one.
 * Test mode: the site sends reviewer "Name (TEST)"; those rows are kept apart and can be deleted with action "clearTest".
 * The endpoint is public, so it only accepts the listed names, photo ids like R001, and clips text. Text that starts
 * with = + - @ is stored as plain text so it can never run as a formula.
 */
const VERSION = 2;
const SHEET_NAME = 'Responses';
const REVIEWERS = ['Clodagh', 'Michael', 'Stanley', 'Charlie'];
const PHOTO_BASE = 'https://stanely18.github.io/rockridge-photo-review/photos/';
const HEADERS = ['reviewer', 'photo_id', 'segment', 'required', 'stanley_label', 'my_decision', 'my_severity',
                 'my_comment', 'photo_url', 'client_updated_at', 'received_at'];

function doGet() {
  return json_({ ok: true, version: VERSION, info: 'Rockridge review endpoint is running' });
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(25000);
  try {
    const data = JSON.parse(e.postData.contents);
    if (data.action === 'ping') return json_({ ok: true, version: VERSION });
    const base = String(data.reviewer || '').replace(/ \(TEST\)$/, '');
    if (REVIEWERS.indexOf(base) === -1) return json_({ ok: false, error: 'unknown reviewer' });
    const isTest = / \(TEST\)$/.test(String(data.reviewer));
    const sh = sheet_();
    if (data.action === 'clearTest') {
      if (!isTest) return json_({ ok: false, error: 'only (TEST) rows can be cleared' });
      return json_({ ok: true, version: VERSION, deleted: clearRows_(sh, data.reviewer) });
    }
    const rows = Array.isArray(data.rows) ? data.rows.slice(0, 400) : [];
    const last = sh.getLastRow();
    const index = {};
    if (last > 1) {
      sh.getRange(2, 1, last - 1, 2).getValues().forEach(function (r, i) { index[r[0] + '|' + r[1]] = i + 2; });
    }
    const now = new Date().toISOString();
    const fresh = [];
    const results = [];
    rows.forEach(function (r) {
      if (!/^R\d{3}$/.test(String(r.photo_id))) return;
      const decision = r.decision === 'Yes' || r.decision === 'No' ? r.decision : '';
      const sev = [0, 1, 2, 3].indexOf(r.severity) > -1 ? String(r.severity) : '';
      const vals = [data.reviewer, r.photo_id, clip_(r.segment, 60), r.required ? 'yes' : 'no', clip_(r.stanley_label, 5),
                    decision, sev, clip_(r.comment, 2000), PHOTO_BASE + r.photo_id + '.jpg', clip_(r.updated_at, 25), now];
      const key = data.reviewer + '|' + r.photo_id;
      if (index[key]) {
        sh.getRange(index[key], 1, 1, HEADERS.length).setValues([vals]);
        results.push([r.photo_id, 'updated']);
      } else {
        fresh.push(vals);
        results.push([r.photo_id, 'added']);
      }
    });
    if (fresh.length) {
      const start = sh.getLastRow() + 1;
      if (start + fresh.length - 1 > sh.getMaxRows()) sh.insertRowsAfter(sh.getMaxRows(), fresh.length + 50);
      sh.getRange(start, 1, fresh.length, HEADERS.length).setNumberFormat('@').setValues(fresh);
    }
    return json_({ ok: true, version: VERSION, saved: results, updated: results.length - fresh.length, added: fresh.length });
  } catch (err) {
    return json_({ ok: false, error: String(err) });
  } finally {
    lock.releaseLock();
  }
}

function sheet_() {
  const ss = SpreadsheetApp.getActive();
  let sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) {
    sh = ss.insertSheet(SHEET_NAME);
    sh.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]).setFontWeight('bold');
    sh.getRange(2, 1, sh.getMaxRows() - 1, HEADERS.length).setNumberFormat('@'); // plain text
    sh.setFrozenRows(1);
  }
  return sh;
}

function clearRows_(sh, reviewer) {
  const last = sh.getLastRow();
  if (last < 2) return 0;
  const col = sh.getRange(2, 1, last - 1, 1).getValues();
  let deleted = 0;
  for (let i = col.length - 1; i >= 0; i--) {   // bottom-up so row numbers stay valid
    if (col[i][0] === reviewer) { sh.deleteRow(i + 2); deleted++; }
  }
  return deleted;
}

// Text from a public endpoint must never become a spreadsheet formula (=IMPORTDATA(...) could leak cells):
// anything starting with = + - @ or a tab/CR gets a leading apostrophe so it is stored as plain text.
function clip_(v, n) {
  const t = String(v == null ? '' : v).slice(0, n);
  return /^[=+\-@\t\r]/.test(t) ? "'" + t : t;
}
function json_(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }
