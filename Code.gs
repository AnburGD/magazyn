/**
 * Magazyn QR — backend w Google Sheets (Apps Script).
 * Arkusz "Produkty": Kod | Nazwa | Min
 * Arkusz "Wpisy":    Data | Produkt | Kod | Ilość | PP | Kto | ID | Typ | Dokument
 *
 * Typ: Wydanie (puste = Wydanie, starsze wpisy), Przyjęcie, Korekta (inwentaryzacja).
 * Stan = przyjęcia − wydania + korekty (Ilość korekty ma znak, np. -3).
 *
 * PIN (wpisywany raz w aplikacji na telefonie) NIE jest w tym pliku — trzymamy go
 * we właściwościach skryptu: Ustawienia projektu (⚙) → Właściwości skryptu →
 * Właściwość: PIN, Wartość: twój PIN. Dzięki temu przy aktualizacji wklejasz cały plik bez zmian.
 * Zmiana PIN-u tam = wszyscy muszą wpisać nowy.
 *
 * Ochrona przed zgadywaniem PIN-u: każda błędna próba czeka ~2 s, a po
 * MAX_FAILS błędnych próbach w ciągu 10 minut skrypt blokuje się na LOCK_MIN minut
 * (dla wszystkich) i zapisuje zdarzenie w arkuszu "Bezpieczeństwo".
 */
const MAX_FAILS = 30, LOCK_MIN = 5;
function pin() {
  return PropertiesService.getScriptProperties().getProperty('PIN') || '';
}

// zwraca odpowiedź z błędem albo null, gdy PIN się zgadza
function checkPin(given) {
  const p = pin();
  if (!p) return json({ ok: false, error: 'Brak PIN-u w ustawieniach skryptu — kierownik musi go ustawić (Właściwości skryptu → PIN)' });
  const cache = CacheService.getScriptCache();
  if (cache.get('lockout')) return json({ ok: false, locked: true, error: 'Zbyt wiele błędnych prób PIN-u. Spróbuj ponownie za kilka minut.' });
  if (typeof given === 'string' && given === p) return null;
  Utilities.sleep(2000);
  const fails = (Number(cache.get('fails')) || 0) + 1;
  cache.put('fails', String(fails), 600);
  if (fails >= MAX_FAILS) {
    cache.put('lockout', '1', LOCK_MIN * 60);
    cache.remove('fails');
    logSecurity('Blokada na ' + LOCK_MIN + ' min: ' + fails + ' błędnych prób PIN-u w ciągu 10 minut');
  }
  return json({ ok: false, error: 'Złe hasło (TOKEN)' });
}

function logSecurity(msg) {
  try {
    const ss = SpreadsheetApp.getActive();
    const sh = ss.getSheetByName('Bezpieczeństwo') || ss.insertSheet('Bezpieczeństwo');
    if (sh.getLastRow() === 0) sh.appendRow(['Data', 'Zdarzenie']);
    sh.appendRow([new Date(), msg]);
  } catch (err) {}
}

function readItems() {
  return json({ ok: true, items: rows('Produkty').filter(r => r[0]).map(r => ({ code: String(r[0]), name: String(r[1]), min: Number(r[2]) || 0 })) });
}

function readEntries(limitParam) {
  const limit = Number(limitParam) || 5000;
  const tz = Session.getScriptTimeZone();
  const sheet = rows('Wpisy');
  const all = sheet
    .map((r, i) => ({ r, row: i + 2 }))
    .filter(x => x.r[2]);
  const out = all.slice(-limit).reverse().map(({ r, row }) => ({
    at: r[0] instanceof Date ? Utilities.formatDate(r[0], tz, 'yyyy-MM-dd HH:mm') : String(r[0]),
    name: String(r[1]), code: String(r[2]), qty: Number(r[3]), pp: String(r[4]), by: String(r[5]),
    id: r[6] ? String(r[6]) : 'row-' + row, type: typeOf(r), doc: String(r[8] || '')
  }));
  return json({ ok: true, entries: out, stock: stockMap(sheet) });
}

function setup() {
  const ss = SpreadsheetApp.getActive();
  let p = ss.getSheetByName('Produkty');
  if (!p) {
    p = ss.insertSheet('Produkty');
    p.appendRow(['Kod', 'Nazwa', 'Min']);
    p.appendRow(['MAG-ZEGAR', 'Zegar']);
    p.setFrozenRows(1);
  }
  let w = ss.getSheetByName('Wpisy');
  if (!w) {
    w = ss.insertSheet('Wpisy');
    w.appendRow(['Data', 'Produkt', 'Kod', 'Ilość', 'PP', 'Kto', 'ID', 'Typ', 'Dokument']);
    w.setFrozenRows(1);
    w.getRange('A:A').setNumberFormat('yyyy-mm-dd hh:mm');
    w.getRange('E:E').setNumberFormat('@');
  }
  ensureColumns();
}

// starsze wersje arkusza bez kolumn ID / Typ / Dokument / Min — dopisz nagłówki
function ensureColumns() {
  const ss = SpreadsheetApp.getActive();
  const w = ss.getSheetByName('Wpisy'), p = ss.getSheetByName('Produkty');
  const head = w.getRange(1, 1, 1, 9).getValues()[0];
  if (!head[6]) w.getRange(1, 7).setValue('ID');
  if (!head[7]) w.getRange(1, 8).setValue('Typ');
  if (!head[8]) { w.getRange(1, 9).setValue('Dokument'); w.getRange('I:I').setNumberFormat('@'); }
  if (!p.getRange(1, 3).getValue()) p.getRange(1, 3).setValue('Min');
}

function json(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function rows(name) {
  const sh = SpreadsheetApp.getActive().getSheetByName(name);
  if (!sh || sh.getLastRow() < 2) return [];
  return sh.getRange(2, 1, sh.getLastRow() - 1, Math.max(sh.getLastColumn(), 9)).getValues();
}

const TYPES = { 'Przyjęcie': 'P', 'Korekta': 'K' };  // reszta (też puste) = Wydanie 'W'
function typeOf(r) { return TYPES[String(r[7])] || 'W'; }
function signed(r) { const q = Number(r[3]) || 0, t = typeOf(r); return t === 'P' || t === 'K' ? q : -q; }

function stockMap(all) {
  const m = {};
  all.forEach(r => { if (r[2]) m[String(r[2])] = (m[String(r[2])] || 0) + signed(r); });
  return m;
}

function addRows(data) {
  const sh = SpreadsheetApp.getActive().getSheetByName('Wpisy');
  sh.getRange(sh.getLastRow() + 1, 1, data.length, 9).setValues(data);
}

function newId() {
  return Utilities.getUuid().slice(0, 12);
}

// Zapytania GET (z PIN-em w adresie) nie są obsługiwane — wszystko idzie POST-em.
// Parametry adresu są celowo ignorowane, żeby PIN nigdy nie był przetwarzany z URL-a.
function doGet() {
  return json({ ok: false, error: 'Nieobsługiwane zapytanie — odśwież aplikację' });
}

function doPost(e) {
  let body;
  try { body = JSON.parse(e.postData.contents); } catch (err) { return json({ ok: false, error: 'Zły format danych' }); }
  const bad = checkPin(body.token);
  if (bad) return bad;
  if (body.action === 'items') return readItems();
  if (body.action === 'entries') return readEntries(body.limit);
  const ss = SpreadsheetApp.getActive();
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    ensureColumns();
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
      const data = lines.map(l => [now, String(body.name), String(body.code), Number(l.qty), String(l.pp).trim(), String(body.by || ''), newId(), 'Wydanie', '']);
      addRows(data);
      return json({ ok: true, saved: data.length });
    }
    if (body.action === 'addReceipt') {
      const qty = Number(body.qty);
      if (!(qty > 0)) return json({ ok: false, error: 'Wpisz ilość' });
      addRows([[new Date(), String(body.name), String(body.code), qty, '', String(body.by || ''), newId(), 'Przyjęcie', String(body.doc || '').trim()]]);
      return json({ ok: true });
    }
    if (body.action === 'inventory') {
      const counted = Number(body.counted), code = String(body.code || '');
      if (!(counted >= 0) || String(body.counted) === '') return json({ ok: false, error: 'Wpisz policzoną ilość' });
      const before = stockMap(rows('Wpisy'))[code] || 0, diff = counted - before;
      addRows([[new Date(), String(body.name), code, diff, '', String(body.by || ''), newId(), 'Korekta', 'Spis: ' + counted]]);
      return json({ ok: true, before: before, diff: diff });
    }
    if (body.action === 'setMin') {
      const code = String(body.code || ''), min = Math.max(0, Number(body.min) || 0);
      const all = rows('Produkty');
      for (let i = 0; i < all.length; i++) {
        if (String(all[i][0]) === code) { ss.getSheetByName('Produkty').getRange(i + 2, 3).setValue(min); return json({ ok: true }); }
      }
      return json({ ok: false, error: 'Nie znaleziono produktu — odśwież listę' });
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
