/**
 * Kniha jízd – backend (Google Apps Script, vázaný na tabulku).
 *
 * Zdroj pravdy je skrytý list "_kniha_data". List "Kniha jízd" vzniká kopií vzoru
 * "108P-Opicake(Honza)" a při každé změně se z dat znovu vykreslí do dvouřádkových
 * bloků vzoru (datum, odkud/kam, km, řidič). Ostatní listy se nemění.
 */

const SOURCE_SHEET = '108P-Opicake(Honza)'; // vzor
const TARGET_SHEET = 'Kniha jízd';          // sem se vykresluje
const DATA_SHEET = '_kniha_data';           // skrytá databáze
const FALLBACK_FIRST_ROW = 40;              // první řádek jízd, pokud se hlavička nenajde
const COL_DATE = 3;    // C – Datum (sloučené přes 2 řádky)
const COL_FROM = 5;    // E – odkud (horní řádek) / kam (dolní řádek)
const COL_KM = 7;      // G – Ujeté km (sloučené přes 2 řádky)
const COL_DRIVER = 16; // P – Řidič (nový sloupec za tabulkou vzoru)
const COL_LAST_TEMPLATE = 15; // O – poslední sloupec vzoru
const MAX_TEXT = 80;

// ---------- HTTP ----------

function doGet(e) {
  return handle_((e && e.parameter) || {});
}

function doPost(e) {
  let p;
  try {
    p = JSON.parse(e.postData.contents);
  } catch (err) {
    return out_({ ok: false, error: 'Neplatný požadavek.' });
  }
  return handle_(p);
}

function handle_(p) {
  try {
    const action = p.action || 'list';
    if (action === 'list') return out_({ ok: true, rows: readRows_() });

    const lock = LockService.getScriptLock();
    lock.waitLock(25000);
    try {
      if (action === 'export') {
        render_(readRows_());
        const stamp = Utilities.formatDate(new Date(), 'Europe/Prague', 'yyyy-MM-dd');
        return out_({ ok: true, filename: 'kniha-jizd-' + stamp + '.xlsx', base64: exportXlsx_() });
      }
      let rows = readRows_();
      if (action === 'add') {
        const r = validate_(p);
        r.id = Utilities.getUuid();
        rows.push(r);
      } else if (action === 'update') {
        const r = validate_(p);
        const i = rows.findIndex(x => x.id === p.id);
        if (i < 0) throw new Error('Řádek už neexistuje, obnov stránku.');
        r.id = p.id;
        rows[i] = r;
      } else if (action === 'delete') {
        const n = rows.length;
        rows = rows.filter(x => x.id !== p.id);
        if (rows.length === n) throw new Error('Řádek už neexistuje, obnov stránku.');
      } else {
        throw new Error('Neznámá akce.');
      }
      writeRows_(rows);
      render_(rows);
      return out_({ ok: true, rows: readRows_() });
    } finally {
      lock.releaseLock();
    }
  } catch (err) {
    return out_({ ok: false, error: String(err && err.message ? err.message : err) });
  }
}

function out_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

// ---------- validace ----------

function validate_(p) {
  const datum = String(p.datum || '');
  const m = /^(\d{4})-(\d{2})-(\d{2})$/.exec(datum);
  if (!m) throw new Error('Neplatné datum.');
  const d = new Date(Date.UTC(+m[1], +m[2] - 1, +m[3]));
  if (d.getUTCFullYear() !== +m[1] || d.getUTCMonth() !== +m[2] - 1 || d.getUTCDate() !== +m[3]) {
    throw new Error('Neplatné datum.');
  }
  const odkud = String(p.odkud || '').trim();
  const kam = String(p.kam || '').trim();
  const ridic = String(p.ridic || '').trim();
  if (!odkud || !kam) throw new Error('Vyplň odkud i kam.');
  if (odkud.length > MAX_TEXT || kam.length > MAX_TEXT || ridic.length > MAX_TEXT) {
    throw new Error('Text je příliš dlouhý (max ' + MAX_TEXT + ' znaků).');
  }
  const km = Number(String(p.km).replace(',', '.'));
  if (!isFinite(km) || km < 0 || km > 100000) throw new Error('Neplatný počet km.');
  return { datum: datum, odkud: odkud, kam: kam, km: Math.round(km * 10) / 10, ridic: ridic };
}

// ---------- datový list ----------

function ss_() {
  return SpreadsheetApp.getActiveSpreadsheet();
}

function dataSheet_() {
  const ss = ss_();
  let sh = ss.getSheetByName(DATA_SHEET);
  if (!sh) {
    sh = ss.insertSheet(DATA_SHEET);
    sh.getRange(1, 1, 1, 6).setValues([['id', 'datum', 'odkud', 'kam', 'km', 'ridic']]);
    sh.getRange('B:B').setNumberFormat('@');
    sh.hideSheet();
  }
  return sh;
}

function readRows_() {
  const sh = dataSheet_();
  const last = sh.getLastRow();
  if (last < 2) return [];
  const tz = ss_().getSpreadsheetTimeZone();
  return sh.getRange(2, 1, last - 1, 6).getValues()
    .filter(r => r[0])
    .map(r => ({
      id: String(r[0]),
      datum: r[1] instanceof Date ? Utilities.formatDate(r[1], tz, 'yyyy-MM-dd') : String(r[1]),
      odkud: String(r[2]),
      kam: String(r[3]),
      km: Number(r[4]),
      ridic: String(r[5])
    }));
}

function writeRows_(rows) {
  const sh = dataSheet_();
  const last = sh.getLastRow();
  if (last > 1) sh.getRange(2, 1, last - 1, 6).clearContent();
  if (!rows.length) return;
  sh.getRange(2, 2, rows.length, 1).setNumberFormat('@');
  sh.getRange(2, 1, rows.length, 6).setValues(
    rows.map(r => [r.id, r.datum, r.odkud, r.kam, r.km, r.ridic])
  );
}

// ---------- vykreslení do vzoru ----------

function targetSheet_() {
  const ss = ss_();
  let sh = ss.getSheetByName(TARGET_SHEET);
  if (sh) return sh;
  const src = ss.getSheetByName(SOURCE_SHEET);
  if (!src) throw new Error('V tabulce chybí vzorový list „' + SOURCE_SHEET + '“.');
  sh = src.copyTo(ss);
  sh.setName(TARGET_SHEET);
  return sh;
}

function firstRow_(sh) {
  const vals = sh.getRange(1, COL_DATE, Math.min(80, sh.getMaxRows()), 1).getValues();
  for (let i = 0; i < vals.length; i++) {
    if (String(vals[i][0]).trim().toLowerCase() === 'datum') return i + 1 + 2; // hlavička má 2 řádky
  }
  return FALLBACK_FIRST_ROW;
}

// Počet souvislých dvouřádkových bloků (sloučené buňky ve sloupci Datum) od prvního řádku.
function capacity_(sh, first) {
  const h = sh.getMaxRows() - first + 1;
  if (h < 2) return 0;
  const merged = sh.getRange(first, COL_DATE, h, 1).getMergedRanges()
    .filter(m => m.getNumRows() === 2)
    .sort((a, b) => a.getRow() - b.getRow());
  let cap = 0;
  let expect = first;
  for (let i = 0; i < merged.length; i++) {
    if (merged[i].getRow() !== expect) break;
    cap++;
    expect += 2;
  }
  return cap;
}

function render_(rows) {
  const sh = targetSheet_();
  const first = firstRow_(sh);
  let cap = capacity_(sh, first);
  if (cap === 0) throw new Error('Ve vzoru jsem nenašel tabulku jízd (dvouřádkové bloky pod hlavičkou „Datum“).');

  // Dost místa? Když ne, vložím další bloky před poslední (aby se rozšířily součty) a okopíruji formát.
  const need = rows.length - cap;
  if (need > 0) {
    const lastStart = first + (cap - 1) * 2;
    sh.insertRowsBefore(lastStart, need * 2);
    const srcBlock = sh.getRange(lastStart + need * 2, COL_DATE, 2, COL_LAST_TEMPLATE - COL_DATE + 1);
    for (let i = 0; i < need; i++) {
      const dst = sh.getRange(lastStart + i * 2, COL_DATE, 2, COL_LAST_TEMPLATE - COL_DATE + 1);
      srcBlock.copyTo(dst, SpreadsheetApp.CopyPasteType.PASTE_FORMAT, false);
      srcBlock.copyTo(dst, SpreadsheetApp.CopyPasteType.PASTE_DATA_VALIDATION, false);
    }
    cap = capacity_(sh, first);
  }

  const n = cap * 2;

  // Sloupec Řidič: hlavička a formát podle sloupce Ujeté km.
  sh.getRange(first - 2, COL_KM, 2, 1).copyTo(sh.getRange(first - 2, COL_DRIVER, 2, 1), SpreadsheetApp.CopyPasteType.PASTE_FORMAT, false);
  sh.getRange(first - 2, COL_DRIVER).setValue('Řidič');
  sh.getRange(first, COL_KM, n, 1).copyTo(sh.getRange(first, COL_DRIVER, n, 1), SpreadsheetApp.CopyPasteType.PASTE_FORMAT, false);
  sh.setColumnWidth(COL_DRIVER, 130);

  // Vymazat vzorová data a zapsat jízdy seřazené podle data.
  [COL_DATE, COL_FROM, COL_KM, COL_DRIVER].forEach(c => sh.getRange(first, c, n, 1).clearContent());
  const sorted = rows.slice().sort((a, b) => (a.datum < b.datum ? -1 : a.datum > b.datum ? 1 : 0));
  const dates = [], places = [], kms = [], drivers = [];
  for (let i = 0; i < cap; i++) {
    const r = sorted[i];
    if (r) {
      const p = r.datum.split('-').map(Number);
      dates.push([new Date(Date.UTC(p[0], p[1] - 1, p[2], 12))], ['']);
      places.push([r.odkud], [r.kam]);
      kms.push([r.km], ['']);
      drivers.push([r.ridic], ['']);
    } else {
      dates.push([''], ['']);
      places.push([''], ['']);
      kms.push([''], ['']);
      drivers.push([''], ['']);
    }
  }
  sh.getRange(first, COL_DATE, n, 1).setNumberFormat('d.m.yyyy').setValues(dates);
  sh.getRange(first, COL_FROM, n, 1).setValues(places);
  sh.getRange(first, COL_KM, n, 1).setValues(kms);
  sh.getRange(first, COL_DRIVER, n, 1).setValues(drivers);
  SpreadsheetApp.flush();
}

// ---------- export do .xlsx ----------

// Exportuje jen list "Kniha jízd" (kopie do dočasné tabulky, aby soubor neobsahoval ostatní listy).
function exportXlsx_() {
  const sh = targetSheet_();
  const tmp = SpreadsheetApp.create('kniha-jizd-export-tmp');
  try {
    const copy = sh.copyTo(tmp);
    copy.setName(TARGET_SHEET);
    tmp.getSheets().forEach(s => {
      if (s.getSheetId() !== copy.getSheetId()) tmp.deleteSheet(s);
    });
    SpreadsheetApp.flush();
    const url = 'https://docs.google.com/spreadsheets/d/' + tmp.getId() + '/export?format=xlsx';
    const resp = UrlFetchApp.fetch(url, {
      headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() },
      muteHttpExceptions: true
    });
    if (resp.getResponseCode() !== 200) throw new Error('Export selhal (HTTP ' + resp.getResponseCode() + ').');
    return Utilities.base64Encode(resp.getBlob().getBytes());
  } finally {
    DriveApp.getFileById(tmp.getId()).setTrashed(true);
  }
}
