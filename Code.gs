/**
 * Magazyn QR — backend w Google Sheets (Apps Script).
 * Arkusz "Produkty": Kod | Nazwa
 * Arkusz "Wpisy":    Data | Produkt | Kod | Ilość | PP | Kto | ID
 *
 * TOKEN to PIN, który pracownicy wpisują raz w aplikacji na telefonie.
 * NIE wpisuj go do index.html. Zmiana PIN-u tutaj = wszyscy muszą wpisać nowy.
 */
const TOKEN = 'ZMIEN-NA-SWOJ-PIN';

function setup() {
  const ss = SpreadsheetApp.getActive();
  let p = ss.getSheetByName('Produkty');
  if (!p) {
    p = ss.insertSheet('Produkty');
    p.appendRow(['Kod', 'Nazwa']);
    p.appendRow(['MAG-ZEGAR', 'Zegar']);
    p.setFrozenRows(1);
  }
  let w = ss.getSheetByName('Wpisy');
  if (!w) {
    w = ss.insertSheet('Wpisy');
    w.appendRow(['Data', 'Produkt', 'Kod', 'Ilość', 'PP', 'Kto', 'ID']);
    w.setFrozenRows(1);
    w.getRange('A:A').setNumberFormat('yyyy-mm-dd hh:mm');
    w.getRange('E:E').setNumberFormat('@');
  }
  // starsza wersja arkusza bez kolumny ID
  if (w.getLastColumn() < 7) w.getRange(1, 7, 1, 1).setValues([['ID']]);
}

function json(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function rows(name) {
  const sh = SpreadsheetApp.getActive().getSheetByName(name);
  if (!sh || sh.getLastRow() < 2) return [];
  return sh.getRange(2, 1, sh.getLastRow() - 1, Math.max(sh.getLastColumn(), 7)).getValues();
}

function newId() {
  return Utilities.getUuid().slice(0, 12);
}

function doGet(e) {
  const p = e.parameter || {};
  if (p.token !== TOKEN) return json({ ok: false, error: 'Złe hasło (TOKEN)' });
  if (p.action === 'items') {
    return json({ ok: true, items: rows('Produkty').filter(r => r[0]).map(r => ({ code: String(r[0]), name: String(r[1]) })) });
  }
  if (p.action === 'entries') {
    const limit = Number(p.limit) || 5000;
    const tz = Session.getScriptTimeZone();
    const all = rows('Wpisy')
      .map((r, i) => ({ r, row: i + 2 }))
      .filter(x => x.r[2]);
    const out = all.slice(-limit).reverse().map(({ r, row }) => ({
      at: r[0] instanceof Date ? Utilities.formatDate(r[0], tz, 'yyyy-MM-dd HH:mm') : String(r[0]),
      name: String(r[1]), code: String(r[2]), qty: Number(r[3]), pp: String(r[4]), by: String(r[5]),
      id: r[6] ? String(r[6]) : 'row-' + row
    }));
    return json({ ok: true, entries: out });
  }
  return json({ ok: false, error: 'Nieznana akcja' });
}

function doPost(e) {
  let body;
  try { body = JSON.parse(e.postData.contents); } catch (err) { return json({ ok: false, error: 'Zły format danych' }); }
  if (body.token !== TOKEN) return json({ ok: false, error: 'Złe hasło (TOKEN)' });
  const ss = SpreadsheetApp.getActive();
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    if (body.action === 'addItem') {
      const code = String(body.code || '').trim(), name = String(body.name || '').trim();
      if (!code || !name) return json({ ok: false, error: 'Brak kodu lub nazwy' });
      const all = rows('Produkty');
      if (all.some(r => String(r[0]) === code)) return json({ ok: true });
      if (all.some(r => String(r[1]).trim().toLowerCase() === name.toLowerCase()))
        return json({ ok: false, error: 'Produkt „' + name + '” już jest na liście' });
      ss.getSheetByName('Produkty').appendRow([code, name]);
      return json({ ok: true });
    }
    if (body.action === 'deleteItem') {
      const code = String(body.code || '');
      const all = rows('Produkty');
      for (let i = 0; i < all.length; i++) {
        if (String(all[i][0]) === code) { ss.getSheetByName('Produkty').deleteRow(i + 2); return json({ ok: true }); }
      }
      return json({ ok: false, error: 'Nie znaleziono produktu — odśwież listę' });
    }
    if (body.action === 'addEntries') {
      const lines = (body.lines || []).filter(l => Number(l.qty) > 0 && String(l.pp || '').trim());
      if (!lines.length) return json({ ok: false, error: 'Brak poprawnych wierszy' });
      const now = new Date();
      const data = lines.map(l => [now, String(body.name), String(body.code), Number(l.qty), String(l.pp).trim(), String(body.by || ''), newId()]);
      const sh = ss.getSheetByName('Wpisy');
      sh.getRange(sh.getLastRow() + 1, 1, data.length, 7).setValues(data);
      return json({ ok: true, saved: data.length });
    }
    if (body.action === 'deleteEntry') {
      const id = String(body.id || '');
      const sh = ss.getSheetByName('Wpisy');
      let row = -1;
      if (id.indexOf('row-') === 0) {
        row = Number(id.slice(4));
      } else {
        const all = rows('Wpisy');
        for (let i = 0; i < all.length; i++) if (String(all[i][6]) === id) { row = i + 2; break; }
      }
      if (row < 2 || row > sh.getLastRow()) return json({ ok: false, error: 'Nie znaleziono wpisu — odśwież listę' });
      sh.deleteRow(row);
      return json({ ok: true });
    }
    return json({ ok: false, error: 'Nieznana akcja' });
  } finally {
    lock.releaseLock();
  }
}
