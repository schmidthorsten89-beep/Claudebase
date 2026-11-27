/**
 * Stempelzeiten – Google-Sheets-Speicher
 * ---------------------------------------
 * Dieses Skript macht ein Google Sheet zum Online-Speicher der Stempelzeiten-App.
 *
 * Einrichtung (einmalig):
 *  1. Neues Google Sheet anlegen (z. B. "Stempelzeiten").
 *  2. Erweiterungen → Apps Script. Den Inhalt von Code.gs durch diese Datei ersetzen.
 *  3. Unten bei TOKEN ein eigenes, langes Passwort eintragen und speichern.
 *  4. Bereitstellen → Neue Bereitstellung → Typ "Web-App"
 *       Ausführen als: Ich      Zugriff: Jeder
 *     → Zugriff autorisieren → Web-App-URL kopieren.
 *  5. In der App unter Einstellungen → Google Sheets die URL und das Passwort eintragen.
 *
 * Nach einer Änderung an diesem Skript: Bereitstellen → Bereitstellungen verwalten →
 * Bearbeiten → Version "Neue Version" → Bereitstellen (die URL bleibt gleich).
 */

const TOKEN = 'HIER-EIGENES-PASSWORT-EINTRAGEN';

const SHEET_DATA = 'Daten';   // Rohdaten der App (nicht von Hand ändern)
const SHEET_VIEW = 'Zeiten';  // lesbare Tabelle, wird von der App befüllt
const VIEW_HEADER = ['Datum', 'Wochentag', 'Beginn', 'Ende', 'Ort', 'Pause (h)', 'Abwesenheit', 'Plan',
                     'Soll (h)', 'Ist (h)', 'HO (h)', 'Büro (h)', '+/- (h)', 'Feiertag', 'Notiz'];

function doGet(e) {
  return json_({ ok: true, app: 'stempelzeiten', hinweis: 'Web-App läuft. Bitte die URL in der App eintragen.' });
}

function doPost(e) {
  let req;
  try { req = JSON.parse(e.postData.contents); } catch (err) { return json_({ ok: false, error: 'Ungültige Anfrage' }); }
  if (!TOKEN || TOKEN === 'HIER-EIGENES-PASSWORT-EINTRAGEN') return json_({ ok: false, error: 'Im Skript ist noch kein Passwort (TOKEN) eingetragen.' });
  if (req.token !== TOKEN) return json_({ ok: false, error: 'Falsches Passwort' });

  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    return json_(sync_(req));
  } finally {
    lock.releaseLock();
  }
}

/**
 * Abgleich: Jeder Eintrag (Tag oder Einstellungen) trägt einen Zeitstempel u.
 * Der neuere Stand gewinnt. Gelöschte Tage werden als {deleted:true, u} gemerkt.
 */
function sync_(req) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = sheet_(ss, SHEET_DATA, ['Schlüssel', 'Daten (JSON)', 'Geändert']);
  const values = sh.getLastRow() > 1 ? sh.getRange(2, 1, sh.getLastRow() - 1, 2).getValues() : [];
  const store = {}, rowOf = {};
  values.forEach((r, i) => {
    const key = String(r[0]);
    if (!key) return;
    try { store[key] = JSON.parse(r[1]); rowOf[key] = i + 2; } catch (err) { /* defekte Zeile ignorieren */ }
  });

  const changes = req.changes || {};        // { '2026-11-27': {...,u}, '_settings': {...,u} }
  const rows = req.rows || {};              // lesbare Zeilen je Tag (von der App berechnet)
  const newRows = [], viewUpdates = {};
  Object.keys(changes).forEach(key => {
    const inc = changes[key];
    if (!inc || typeof inc.u !== 'number') return;
    const cur = store[key];
    if (cur && typeof cur.u === 'number' && cur.u >= inc.u) return;   // Server ist neuer
    store[key] = inc;
    const line = [key, JSON.stringify(inc), new Date(inc.u)];
    if (rowOf[key]) sh.getRange(rowOf[key], 1, 1, 3).setValues([line]);
    else newRows.push(line);
    if (key !== '_settings') viewUpdates[key] = inc.deleted ? null : (rows[key] || null);
  });
  if (newRows.length) sh.getRange(sh.getLastRow() + 1, 1, newRows.length, 3).setValues(newRows);

  if (Object.keys(viewUpdates).length) updateView_(ss, viewUpdates);

  return { ok: true, store: store, serverTime: Date.now() };
}

function updateView_(ss, updates) {
  const sh = sheet_(ss, SHEET_VIEW, VIEW_HEADER);
  const n = sh.getLastRow() - 1;
  const keys = n > 0 ? sh.getRange(2, 1, n, 1).getDisplayValues().map(r => r[0]) : [];
  const rowOf = {};
  keys.forEach((k, i) => { rowOf[toKey_(k)] = i + 2; });
  const toDelete = [];
  const append = [];
  Object.keys(updates).forEach(key => {
    const row = updates[key];
    if (row && row.length === VIEW_HEADER.length) {
      if (rowOf[key]) sh.getRange(rowOf[key], 1, 1, row.length).setValues([row]);
      else append.push(row);
    } else if (rowOf[key]) toDelete.push(rowOf[key]);
  });
  toDelete.sort((a, b) => b - a).forEach(r => sh.deleteRow(r));
  if (append.length) sh.getRange(sh.getLastRow() + 1, 1, append.length, VIEW_HEADER.length).setValues(append);
  const last = sh.getLastRow();
  if (last > 2) {
    // nach Datum sortieren (Spalte A enthält TT.MM.JJJJ als Text → Hilfssortierung)
    const data = sh.getRange(2, 1, last - 1, VIEW_HEADER.length).getValues();
    data.sort((a, b) => toKey_(a[0]).localeCompare(toKey_(b[0])));
    sh.getRange(2, 1, data.length, VIEW_HEADER.length).setValues(data);
  }
}

function toKey_(d) {
  if (d instanceof Date) return Utilities.formatDate(d, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  const m = String(d).match(/^(\d{2})\.(\d{2})\.(\d{4})$/);
  return m ? m[3] + '-' + m[2] + '-' + m[1] : String(d);
}

function sheet_(ss, name, header) {
  let sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
    sh.getRange(1, 1, 1, header.length).setValues([header]).setFontWeight('bold');
    sh.setFrozenRows(1);
    sh.getRange('A:A').setNumberFormat('@');   // Datum als Text, damit nichts umgewandelt wird
  }
  return sh;
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
