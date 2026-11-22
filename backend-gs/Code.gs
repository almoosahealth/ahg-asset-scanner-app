/**
 * AHG Asset Scanner — back end v2 (bound to the inventory Google Sheet)
 * Deploy: Deploy → New deployment → Web app → Execute as: Me, Who has access: Anyone
 *
 * Tabs (Settings, Scans, Staff are created automatically on first use):
 *   ASSET_SHEET  – the asset list. Any columns. Found / Location / Scanned At / Scanned By are added if missing.
 *   Settings     – Key | Value. Controls the app (locations, columns, code pattern…). Edit here, no code change.
 *   Scans        – every scan, append-only (history + unlisted assets).
 *   Staff        – Name column. Fill it and the app shows a one-tap name picker.
 */

const ASSET_SHEET = 'Sheet1';
const SETTINGS_SHEET = 'Settings';
const SCANS_SHEET = 'Scans';
const STAFF_SHEET = 'Staff';
const TOKEN = 'CHANGE_ME';

const DEFAULTS = [
  ['Title', 'AHG Asset Scanner'],
  ['Locations', 'Room, Main station (CN), Back up store, Equipment store, With Biomed, In ASH, Other'],
  ['Room Area', 'Room'],
  ['Rooms', '4001-4026'],
  ['Match Columns', 'Asset Number, Old Asset ID'],
  ['Card Columns', 'Product Category, Status, Comment, Handwritten Note'],
  ['Group Column', 'Product Category'],
  ['Code Pattern', 'BM[\\s_-]?\\d{4,6}|ICU[\\s_-]?\\d{1,4}'],
];
const OLD_LOCATIONS = 'On unit, Store, CN Backup store, NM store, With Biomed, In ASH, Other';
const LATEST_COLS = ['Found', 'Location', 'Scanned At', 'Scanned By'];
const SCAN_HEADERS = ['Scanned At', 'Code', 'Raw QR Text', 'On List', 'Location', 'Scanned By', 'Asset Row'];

// ---------- HTTP ----------
function doGet(e) {
  const p = (e && e.parameter) || {};
  if (p.token !== TOKEN) return json_({ ok: false, error: 'Bad token' });
  try {
    ensure_();
    return json_({ ok: true, settings: readSettings_(), rows: readRows_(), staff: readStaff_() });
  } catch (err) {
    return json_({ ok: false, error: String(err.message || err) });
  }
}

function doPost(e) {
  let b = {};
  try { b = JSON.parse((e && e.postData && e.postData.contents) || '{}'); }
  catch (err) { return json_({ ok: false, error: 'Bad JSON' }); }
  if (b.token !== TOKEN) return json_({ ok: false, error: 'Bad token' });

  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    return json_(save_(b));
  } catch (err) {
    return json_({ ok: false, error: String(err.message || err) });
  } finally {
    lock.releaseLock();
  }
}

// ---------- Setup ----------
function ensure_() {
  const ss = SpreadsheetApp.getActive();

  let st = ss.getSheetByName(SETTINGS_SHEET);
  if (!st) {
    st = ss.insertSheet(SETTINGS_SHEET);
    st.getRange(1, 1, 1, 2).setValues([['Key', 'Value']]).setFontWeight('bold');
    st.getRange(2, 1, DEFAULTS.length, 2).setValues(DEFAULTS);
    st.autoResizeColumns(1, 2);
  } else {
    // Add any new setting rows; move an untouched old Locations default to the new one
    const have = st.getDataRange().getDisplayValues();
    const keys = have.map(r => String(r[0]).trim());
    const li = keys.indexOf('Locations');
    if (li > 0 && String(have[li][1]).trim() === OLD_LOCATIONS) st.getRange(li + 1, 2).setValue(DEFAULTS[1][1]);
    const add = DEFAULTS.filter(([k]) => keys.indexOf(k) < 0);
    if (add.length) st.getRange(st.getLastRow() + 1, 1, add.length, 2).setValues(add);
  }

  let sc = ss.getSheetByName(SCANS_SHEET);
  if (!sc) {
    sc = ss.insertSheet(SCANS_SHEET);
    sc.getRange(1, 1, 1, SCAN_HEADERS.length).setValues([SCAN_HEADERS]).setFontWeight('bold');
    sc.setFrozenRows(1);
  }

  let sf = ss.getSheetByName(STAFF_SHEET);
  if (!sf) {
    sf = ss.insertSheet(STAFF_SHEET);
    sf.getRange(1, 1).setValue('Name').setFontWeight('bold');
  }

  const sh = assets_();
  const lastCol = Math.max(sh.getLastColumn(), 1);
  const headers = sh.getRange(1, 1, 1, lastCol).getDisplayValues()[0];
  const missing = LATEST_COLS.filter(c => headers.indexOf(c) < 0);
  if (missing.length) sh.getRange(1, lastCol + 1, 1, missing.length).setValues([missing]).setFontWeight('bold');
}

// ---------- Reads ----------
function readSettings_() {
  const out = {};
  DEFAULTS.forEach(([k, v]) => (out[k] = v));
  const st = SpreadsheetApp.getActive().getSheetByName(SETTINGS_SHEET);
  st.getDataRange().getDisplayValues().slice(1).forEach(([k, v]) => {
    if (k && String(v).trim()) out[String(k).trim()] = String(v).trim();   // blank value → keep default
  });
  return out;
}

function readRows_() {
  const values = assets_().getDataRange().getDisplayValues();
  const headers = values.shift();
  return values
    .map((r, i) => {
      const o = { _row: i + 2 };
      headers.forEach((h, c) => { if (h) o[h] = r[c]; });
      return o;
    })
    .filter(o => headers.some(h => h && o[h]));
}

function readStaff_() {
  const sf = SpreadsheetApp.getActive().getSheetByName(STAFF_SHEET);
  return sf.getDataRange().getDisplayValues().slice(1).map(r => String(r[0]).trim()).filter(Boolean);
}

// ---------- Save a scan ----------
function save_(b) {
  const sh = assets_();
  const lastCol = sh.getLastColumn();
  const headers = sh.getRange(1, 1, 1, lastCol).getDisplayValues()[0];
  const col = name => headers.indexOf(name) + 1;
  const matchCols = list_(readSettings_()['Match Columns']).map(col).filter(c => c > 0);
  const key = norm_(b.asset);
  const hits = r => matchCols.some(c => norm_(r[c - 1]) === key);

  let row = -1;
  if (key) {
    // Fast path: the phone sends the row it matched — read just that row and confirm
    const hint = Number(b.row);
    if (hint > 1 && hint <= sh.getLastRow() && hits(sh.getRange(hint, 1, 1, lastCol).getDisplayValues()[0])) {
      row = hint;
    } else {
      const values = sh.getRange(2, 1, Math.max(sh.getLastRow() - 1, 1), lastCol).getDisplayValues();
      for (let i = 0; i < values.length && row < 0; i++) if (hits(values[i])) row = i + 2;
    }
  }

  const now = new Date();
  if (row > 0) {
    const cF = col('Found'), cL = col('Location'), cA = col('Scanned At'), cB = col('Scanned By');
    if (cL === cF + 1 && cA === cF + 2 && cB === cF + 3) {
      sh.getRange(row, cF, 1, 4).setValues([['Yes', b.location || '', now, b.by || '']]);   // one write
    } else {
      sh.getRange(row, cF).setValue('Yes');
      sh.getRange(row, cL).setValue(b.location || '');
      sh.getRange(row, cA).setValue(now);
      sh.getRange(row, cB).setValue(b.by || '');
    }
  }

  SpreadsheetApp.getActive().getSheetByName(SCANS_SHEET).appendRow([
    now, b.asset || '', b.raw || '', row > 0 ? 'Yes' : 'No', b.location || '', b.by || '', row > 0 ? row : ''
  ]);
  return { ok: true, row: row > 0 ? row : null, unlisted: row < 0 };
}

// ---------- Helpers ----------
function assets_() {
  const sh = SpreadsheetApp.getActive().getSheetByName(ASSET_SHEET);
  if (!sh) throw new Error('Asset tab not found: ' + ASSET_SHEET);
  return sh;
}
function list_(s) { return String(s || '').split(',').map(x => x.trim()).filter(Boolean); }
function norm_(s) { return String(s || '').toUpperCase().replace(/[^A-Z0-9]/g, ''); }
function json_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}

/** Run once from the editor (▶ Run → setupNow) to create the tabs and authorize. */
function setupNow() { ensure_(); }
