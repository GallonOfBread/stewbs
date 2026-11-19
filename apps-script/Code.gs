// Spicy Uno signatures: the API behind stewbs.com/signatures, plus the share-link menu.
//
// Setup (once, in the spreadsheet):
//   1. Extensions → Apps Script. Paste this file as Code.gs, and LinkDialog.html as an
//      HTML file named LinkDialog. Save.
//   2. Reload the spreadsheet, then use Spicy Uno → Set up tabs.
//   3. Deploy → New deployment → Web app. Execute as: Me. Who has access: Anyone.
//      Copy the web app URL into PROD_API in signatures.html.
//   Later code changes: Deploy → Manage deployments → pencil → Version: New version.
//   That keeps the same URL. (The spreadsheet menu always runs the latest saved code.)
//
// Anyone can call the web app URL, but every request needs a code from the Links tab
// that hasn't expired, and a code can only change wins dated within its own days.

const SITE_URL = 'https://stewbs.com/signatures#';
const LINKS_TAB = 'Links';
// Columns A–H. "To" is last so links made before ranges existed still read correctly
// (an empty To means a one-day link).
const LINK_HEADERS = ['Code', 'Deck', 'From', 'Created', 'First opened', 'Expires', 'Note', 'To'];
const WIN_HEADERS = ['ID', 'Color', 'Value', 'Signer', 'Date', 'Card ID'];
const LINK_HOURS = 24;     // a link works for this long after it's first opened
const UNOPENED_DAYS = 14;  // a link nobody opens stops working after this
const MAX_PER_DAY_LINKS = 31;
const MAX_NAME = 40;

// Must match DECKS in signatures.html
const DECKS = {
  gridfetti: { tab: 'Gridfetti 3.0', colors: ['Gold', 'Pink', 'Purple', 'Teal', 'Wild'] },
  matrix:    { tab: 'Matrix',        colors: ['Blue', 'Green', 'Red', 'Yellow', 'Wild'] },
};
// The deck the share-link dialog starts on
const DEFAULT_DECK = 'matrix';
const VALUES = ['0', '1', '2', '3', '4', '5', '6', '7', '8', '9', 'Draw 2', 'Skip', 'Reverse'];
const WILD_VALUES = ['Draw 4', 'Basic'];
// Must match SUFFIX_SEQ in signatures.html
const SUFFIX_SEQ = '123456789ABCDEFGHIJKLMNOPQRSTUVWXYZ@#$%!&()*+-/;<=>?^|~0αβγδεζηθικλμνξοπρστυφχψω';

class ApiError extends Error {
  constructor(code) { super(code); this.code = code; }
}

// ── Web app ────────────────────────────────────────────────────────────────

function doGet() {
  return ContentService.createTextOutput('Spicy Uno signatures endpoint.');
}

// Every call is a POST with a JSON body: { action, code, ... }.
// The page sends it as text/plain so the browser skips the CORS preflight.
function doPost(e) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(15000);
  } catch (err) {
    return json({ ok: false, error: 'busy' });
  }
  try {
    let req;
    try { req = JSON.parse(e.postData.contents); } catch (err) { throw new ApiError('bad-request'); }
    const handler = Object.prototype.hasOwnProperty.call(ACTIONS, req.action) ? ACTIONS[req.action] : null;
    if (!handler) throw new ApiError('bad-request');
    const link = checkLink(req.code, req.action === 'open');
    return json(Object.assign({ ok: true }, handler(link, req)));
  } catch (err) {
    if (err instanceof ApiError) return json({ ok: false, error: err.code });
    console.error(err);
    return json({ ok: false, error: 'server' });
  } finally {
    lock.releaseLock();
  }
}

const ACTIONS = {
  // Everything the page needs: the link's days and deck, and the deck's full history
  open(link) {
    return linkInfo(link, readWins(link.deck));
  },

  // Fill in or change who signed a win, or clear both to move it back to unposted
  save(link, req) {
    const wins = readWins(link.deck);
    const win = ownWin(wins, link, req.id);
    const signer = cleanText(req.signer, MAX_NAME);
    const cardId = cleanText(req.cardId, 16);
    const clearing = signer === '?' && cardId === '';
    if (!clearing) {
      if (!signer || signer === '?') throw new ApiError('bad-request');
      if (!validCardId(cardId)) throw new ApiError('bad-request');
    }
    // One write for columns D–F. The date is rewritten as MM/DD/YY, which also tidies
    // any date Sheets stored in another format.
    const date = keyToMdy(dayKey(win.date));
    deckSheet(link.deck).getRange(win.row, 4, 1, 3).setNumberFormat('@')
      .setValues([[asText(signer), date, asText(cardId)]]);
    win.signer = signer;
    win.date = date;
    win.cardId = cardId;
    return linkInfo(link, wins);
  },

  // Add a win on one of the link's days. The server picks the ID so two links can't collide.
  add(link, req) {
    const deck = DECKS[link.deck];
    const color = String(req.color || '');
    const value = String(req.value || '');
    if (deck.colors.indexOf(color) === -1) throw new ApiError('bad-request');
    if ((color === 'Wild' ? WILD_VALUES : VALUES).indexOf(value) === -1) throw new ApiError('bad-request');
    const key = req.date ? dayKey(req.date) : link.fromKey;
    if (!inRange(key, link)) throw new ApiError('not-allowed');
    const date = keyToMdy(key);
    const sheet = deckSheet(link.deck);
    const wins = readWins(link.deck);
    const id = String(wins.reduce((max, w) => Math.max(max, Number(w.id)), 0) + 1);
    const row = Math.max(sheet.getLastRow(), 1) + 1;
    sheet.getRange(row, 1, 1, 6).setNumberFormat('@').setValues([[id, color, value, '?', date, '']]);
    wins.push({ row: row, id: id, color: color, value: value, signer: '?', date: date, cardId: '' });
    return Object.assign({ id: id }, linkInfo(link, wins));
  },

  // Only the newest win in the deck, and only if it's within the link's days, so IDs stay in order
  delete(link, req) {
    const wins = readWins(link.deck);
    const last = wins[wins.length - 1];
    if (!last || last.id !== String(req.id)) throw new ApiError('not-allowed');
    if (!inRange(dayKey(last.date), link)) throw new ApiError('not-allowed');
    deckSheet(link.deck).deleteRow(last.row);
    wins.pop();
    return linkInfo(link, wins);
  },
};

function linkInfo(link, wins) {
  return {
    deck: link.deck,
    from: keyToMdy(link.fromKey),
    to: keyToMdy(link.toKey),
    expires: link.expires ? link.expires.toISOString() : null,
    wins: wins.map(w => ({
      id: w.id, color: w.color, value: w.value, signer: w.signer, date: w.date, cardId: w.cardId,
    })),
  };
}

function json(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

// ── Links ──────────────────────────────────────────────────────────────────

// A row of the Links tab as an object. Old one-day links have no To, so To = From.
function linkFromRow(r) {
  const fromKey = dayKey(r[2]);
  const toKey = dayKey(r[7]) || fromKey;
  return {
    code: String(r[0]), deck: String(r[1]), fromKey: fromKey, toKey: toKey,
    created: r[3], firstOpened: r[4], expires: isDate(r[5]) ? r[5] : null,
  };
}

// Finds the link for a code and checks it's still good. Opening a link for the first
// time starts its 24 hours.
function checkLink(code, opening) {
  if (typeof code !== 'string' || !/^[A-Za-z0-9]{12}$/.test(code)) throw new ApiError('invalid-link');
  const sheet = linksSheet();
  const rows = sheet.getDataRange().getValues();
  const now = new Date();
  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][0]) !== code) continue;
    const link = linkFromRow(rows[i]);
    if (!DECKS[link.deck] || !link.fromKey) throw new ApiError('invalid-link');
    if (!link.firstOpened) {
      if (!isDate(link.created) || now - link.created > UNOPENED_DAYS * 864e5) throw new ApiError('expired');
      if (!opening) throw new ApiError('invalid-link');
      link.expires = new Date(now.getTime() + LINK_HOURS * 36e5);
      sheet.getRange(i + 1, 5, 1, 2).setValues([[now, link.expires]]);
    }
    if (!link.expires || now > link.expires) throw new ApiError('expired');
    return link;
  }
  throw new ApiError('invalid-link');
}

function linkUsable(link, now) {
  if (link.firstOpened) return !!link.expires && now <= link.expires;
  return isDate(link.created) && now - link.created <= UNOPENED_DAYS * 864e5;
}

function inRange(key, link) {
  return !!key && key >= link.fromKey && key <= link.toKey;
}

// Codes avoid look-alike characters (0/O, 1/l/I). 12 characters ≈ 69 bits of randomness.
function newCode() {
  const abc = 'ABCDEFGHJKLMNPQRSTUVWXYZabcdefghijkmnopqrstuvwxyz23456789';
  const bytes = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256,
    Utilities.getUuid() + Utilities.getUuid() + Date.now());
  return bytes.slice(0, 12).map(b => abc[(b + 256) % abc.length]).join('');
}

// ── Spreadsheet menu (only people with edit access to the sheet see this) ───

function onOpen() {
  SpreadsheetApp.getUi().createMenu('Spicy Uno')
    .addItem('Share links…', 'showLinkDialog')
    .addItem('Set up tabs', 'setUpTabs')
    .addToUi();
}

function setUpTabs() {
  const ss = SpreadsheetApp.getActive();
  const made = [];
  Object.keys(DECKS).forEach(id => {
    const name = DECKS[id].tab;
    if (ss.getSheetByName(name)) return;
    const s = ss.insertSheet(name);
    s.getRange('A:F').setNumberFormat('@');
    s.getRange(1, 1, 1, WIN_HEADERS.length).setValues([WIN_HEADERS]).setFontWeight('bold');
    s.setFrozenRows(1);
    made.push(name);
  });
  if (!ss.getSheetByName(LINKS_TAB)) {
    const l = ss.insertSheet(LINKS_TAB);
    l.getRange('A:C').setNumberFormat('@');
    l.getRange('H:H').setNumberFormat('@');
    l.getRange(1, 1, 1, LINK_HEADERS.length).setValues([LINK_HEADERS]).setFontWeight('bold');
    l.setFrozenRows(1);
    l.hideSheet();
    made.push(LINKS_TAB + ' (hidden)');
  }
  SpreadsheetApp.getUi().alert(made.length ? 'Added tabs: ' + made.join(', ') : 'All tabs already exist.');
}

function showLinkDialog() {
  const html = HtmlService.createHtmlOutputFromFile('LinkDialog').setWidth(460).setHeight(560);
  SpreadsheetApp.getUi().showModalDialog(html, 'Share links');
}

// What the dialog needs: decks with their default days, and the links still working
function dialogData() {
  const today = todayKey();
  const to = latestWeekdayOnOrBefore(today);
  return {
    defaultDeck: DEFAULT_DECK,
    decks: Object.keys(DECKS).map(id => {
      const latest = latestDay(id);
      let from = latest ? nextWeekday(latest) : to;
      if (from > to) from = to;  // already caught up
      return { id: id, name: DECKS[id].tab, from: from, to: to };
    }),
    active: activeLinks(),
  };
}

// Called from the dialog. Makes one link for the whole range, or one per day.
// Dates come in as YYYY-MM-DD from the date inputs. Overlapping links are allowed.
function createShareLinks(deck, fromIso, toIso, perDay, skipWeekends) {
  if (!DECKS[deck]) throw new Error('Unknown deck.');
  const fromKey = isoKey(fromIso), toKey = isoKey(toIso);
  if (!fromKey || !toKey) throw new Error('Pick both days.');
  if (fromKey > toKey) throw new Error('"From" has to be on or before "To".');
  let ranges = [[fromKey, toKey]];
  if (perDay) {
    ranges = [];
    for (let k = fromKey; k <= toKey; k = addDays(k, 1)) {
      if (skipWeekends && isWeekend(k)) continue;
      ranges.push([k, k]);
    }
    if (!ranges.length) throw new Error('There are no weekdays in that range.');
    if (ranges.length > MAX_PER_DAY_LINKS) throw new Error('That would make more than ' + MAX_PER_DAY_LINKS + ' links.');
  }
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    const sheet = linksSheet();
    if (!sheet.getRange(1, 8).getValue()) sheet.getRange(1, 8).setValue('To').setFontWeight('bold');
    const now = new Date();
    let row = Math.max(sheet.getLastRow(), 1) + 1;
    const made = ranges.map(([f, t]) => {
      const code = newCode();
      // Dates as plain text so Sheets doesn't turn them into Date values
      sheet.getRange(row, 1, 1, 3).setNumberFormat('@').setValues([[code, deck, keyToMdy(f)]]);
      sheet.getRange(row, 4, 1, 4).setValues([[now, '', '', '']]);
      sheet.getRange(row, 8).setNumberFormat('@').setValue(keyToMdy(t));
      row++;
      return { url: SITE_URL + code, label: rangeLabel(f, t) };
    });
    return { deck: DECKS[deck].tab, links: made, active: activeLinks() };
  } finally {
    lock.releaseLock();
  }
}

// Called from the dialog: stop a link working right away
function cancelLink(code) {
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    const sheet = linksSheet();
    const rows = sheet.getDataRange().getValues();
    for (let i = rows.length - 1; i >= 1; i--) {
      if (String(rows[i][0]) === String(code)) sheet.deleteRow(i + 1);
    }
    return activeLinks();
  } finally {
    lock.releaseLock();
  }
}

// Links that still work (or will once opened), newest first
function activeLinks() {
  const now = new Date();
  const tz = sheetTz();
  return linksSheet().getDataRange().getValues().slice(1)
    .map(linkFromRow)
    .filter(l => DECKS[l.deck] && l.fromKey && linkUsable(l, now))
    .reverse()
    .map(l => ({
      code: l.code,
      url: SITE_URL + l.code,
      deck: DECKS[l.deck].tab,
      label: rangeLabel(l.fromKey, l.toKey),
      status: l.firstOpened ? 'Works until ' + Utilities.formatDate(l.expires, tz, 'EEE h:mm a') : 'Not opened yet',
    }));
}

// The latest day in a deck's wins or its links, as a key
function latestDay(deck) {
  let latest = '';
  readWins(deck).forEach(w => { const k = dayKey(w.date); if (k > latest) latest = k; });
  linksSheet().getDataRange().getValues().slice(1).forEach(r => {
    if (String(r[1]) !== deck) return;
    const l = linkFromRow(r);
    if (l.toKey > latest) latest = l.toKey;
  });
  return latest;
}

// ── Sheet helpers ──────────────────────────────────────────────────────────

function linksSheet() {
  const s = SpreadsheetApp.getActive().getSheetByName(LINKS_TAB);
  if (!s) throw new ApiError('not-set-up');
  return s;
}

function deckSheet(deck) {
  const s = SpreadsheetApp.getActive().getSheetByName(DECKS[deck].tab);
  if (!s) throw new ApiError('not-set-up');
  return s;
}

// Wins as shown in the sheet (display values, so dates stay as typed).
// Skips the header and any row whose ID isn't a number.
function readWins(deck) {
  const sheet = deckSheet(deck);
  const last = sheet.getLastRow();
  if (last < 1) return [];
  return sheet.getRange(1, 1, last, 6).getDisplayValues()
    .map((r, i) => ({
      row: i + 1, id: String(r[0]).trim(), color: String(r[1]).trim(), value: String(r[2]).trim(),
      signer: String(r[3]).trim(), date: String(r[4]).trim(), cardId: String(r[5]).trim(),
    }))
    .filter(w => /^\d+$/.test(w.id));
}

// A win this link may change: it exists and is within the link's days
function ownWin(wins, link, id) {
  const win = wins.find(w => w.id === String(id));
  if (!win) throw new ApiError('not-found');
  if (!inRange(dayKey(win.date), link)) throw new ApiError('not-allowed');
  return win;
}

function validCardId(cid) {
  const m = /^([^\s_]{1,10})_(.)$/u.exec(cid);
  return !!m && SUFFIX_SEQ.indexOf(m[2]) !== -1;
}

function cleanText(s, max) {
  return String(s == null ? '' : s).replace(/[\t\r\n]+/g, ' ').trim().slice(0, max);
}

// Stop names like "=IMPORTXML(...)" from becoming formulas in the sheet
function asText(s) {
  return /^[=+\-@]/.test(s) ? "'" + s : s;
}

// ── Dates ──────────────────────────────────────────────────────────────────
// Days are handled as "2026-10-30" style keys (they sort and compare as text).
// Wins store MM/DD/YY.

function sheetTz() {
  return SpreadsheetApp.getActive().getSpreadsheetTimeZone();
}

function isDate(v) {
  return Object.prototype.toString.call(v) === '[object Date]' && !isNaN(v);
}

// Key from "10/30/26", "10/30/2026", "9/4/24", a Date, or a full date string like
// "Wed Nov 18 2026 00:00:00 GMT-0600"
function dayKey(v) {
  if (isDate(v)) return Utilities.formatDate(v, sheetTz(), 'yyyy-MM-dd');
  const s = String(v == null ? '' : v).trim();
  const m = /^(\d{1,2})\/(\d{1,2})\/(\d{2}|\d{4})$/.exec(s);
  if (m) {
    const y = m[3].length === 2 ? '20' + m[3] : m[3];
    return y + '-' + m[1].padStart(2, '0') + '-' + m[2].padStart(2, '0');
  }
  if (/[a-z]{3} \d{1,2} \d{4}/i.test(s)) {
    const d = new Date(s.replace(/\s*\(.*\)$/, ''));
    if (isDate(d)) return Utilities.formatDate(d, sheetTz(), 'yyyy-MM-dd');
  }
  return '';
}

function isoKey(s) {
  return /^\d{4}-\d{2}-\d{2}$/.test(String(s)) ? String(s) : '';
}

function keyToMdy(k) {
  return k ? k.slice(5, 7) + '/' + k.slice(8, 10) + '/' + k.slice(2, 4) : '';
}

// Calendar math on keys, done in UTC so time zones and daylight saving can't shift a day
function keyToUtc(k) {
  return new Date(Date.UTC(Number(k.slice(0, 4)), Number(k.slice(5, 7)) - 1, Number(k.slice(8, 10))));
}

function addDays(k, n) {
  const d = keyToUtc(k);
  d.setUTCDate(d.getUTCDate() + n);
  return Utilities.formatDate(d, 'UTC', 'yyyy-MM-dd');
}

function isWeekend(k) {
  const day = keyToUtc(k).getUTCDay();
  return day === 0 || day === 6;
}

function nextWeekday(k) {
  let d = addDays(k, 1);
  while (isWeekend(d)) d = addDays(d, 1);
  return d;
}

function latestWeekdayOnOrBefore(k) {
  let d = k;
  while (isWeekend(d)) d = addDays(d, -1);
  return d;
}

function todayKey() {
  return Utilities.formatDate(new Date(), sheetTz(), 'yyyy-MM-dd');
}

// "Mon, Nov 16" or "Mon, Nov 16 – Fri, Nov 20"
function rangeLabel(f, t) {
  const fmt = k => Utilities.formatDate(keyToUtc(k), 'UTC', 'EEE, MMM d');
  return f === t ? fmt(f) : fmt(f) + ' – ' + fmt(t);
}
