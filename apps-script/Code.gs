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
// Audit links (stewbs.com/audit) are the exception: they can fix any win in their deck.

const SITE_URL = 'https://stewbs.com/signatures#';
const AUDIT_URL = 'https://stewbs.com/audit#';
const LINKS_TAB = 'Links';
const AUDIT_TAB = 'Audit';
// Columns A–I. "To" and "Kind" come last so older links still read correctly
// (an empty To means a one-day link; an empty Kind means a normal logging link).
const LINK_HEADERS = ['Code', 'Deck', 'From', 'Created', 'First opened', 'Expires', 'Note', 'To', 'Kind'];
const WIN_HEADERS = ['ID', 'Color', 'Value', 'Signer', 'Date', 'Card ID'];
// One row per approved or retired card (Status says which). Fingerprint is the card's
// signatures when it was marked, so the audit page can tell when a card changed afterwards. Color and value are part of
// the key because a mislogged win can share a card ID with a card of another value.
const AUDIT_HEADERS = ['Deck', 'Color', 'Value', 'Card ID', 'Status', 'Fingerprint', 'Reviewed'];
const LINK_HOURS = 24;     // a link works for this long after it's first opened
const AUDIT_HOURS = 168;   // an audit link works for a week after it's first opened
const UNOPENED_DAYS = 14;  // a link nobody opens stops working after this
const MAX_PER_DAY_LINKS = 31;
const MAX_NAME = 40;

// Must match DECKS in signatures.html
// Oldest first. The IDs are saved in links, so they never change (gridfetti is 3.0).
const DECKS = {
  original:   { tab: 'Original',      colors: ['Blue', 'Green', 'Red', 'Yellow', 'Wild'] },
  gridfetti1: { tab: 'Gridfetti 1.0', colors: ['Gold', 'Pink', 'Purple', 'Teal', 'Wild'] },
  gridfetti2: { tab: 'Gridfetti 2.0', colors: ['Gold', 'Pink', 'Purple', 'Teal', 'Wild'] },
  harvest:    { tab: 'Harvest',       colors: ['Banana', 'Berry', 'Orange', 'Plum', 'Wild'] },
  discolight: { tab: 'Discolight',    colors: ['Blue', 'Green', 'Orange', 'Purple', 'Wild'] },
  gridfetti:  { tab: 'Gridfetti 3.0', colors: ['Gold', 'Pink', 'Purple', 'Teal', 'Wild'] },
  matrix:     { tab: 'Matrix',        colors: ['Blue', 'Green', 'Red', 'Yellow', 'Wild'] },
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
    // Audit links only use the audit actions, and logging links only the others
    const has = (o, k) => Object.prototype.hasOwnProperty.call(o, k);
    const kind = has(AUDIT_ACTIONS, req.action) ? 'audit' : '';
    const handler = kind ? AUDIT_ACTIONS[req.action] : has(ACTIONS, req.action) ? ACTIONS[req.action] : null;
    if (!handler) throw new ApiError('bad-request');
    const link = checkLink(req.code, req.action === 'open' || req.action === 'auditOpen', kind);
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

  // Remove a win within the link's days. Its number is left as a gap.
  delete(link, req) {
    const wins = readWins(link.deck);
    const win = ownWin(wins, link, req.id);
    deckSheet(link.deck).deleteRow(win.row);
    wins.splice(wins.indexOf(win), 1);
    return linkInfo(link, wins);
  },

  // Undo of a delete: put the row back with its original number, in number order.
  // Only into an unused number, and only on one of the link's days.
  restore(link, req) {
    const id = String(req.id || '');
    if (!/^\d+$/.test(id)) throw new ApiError('bad-request');
    const wins = readWins(link.deck);
    if (wins.some(w => w.id === id)) throw new ApiError('not-allowed');
    const key = dayKey(req.date);
    if (!inRange(key, link)) throw new ApiError('not-allowed');
    const color = cleanText(req.color, 20);
    const value = cleanText(req.value, 20);
    if (DECKS[link.deck].colors.indexOf(color) === -1 || !value) throw new ApiError('bad-request');
    const signer = cleanText(req.signer, MAX_NAME) || '?';
    const cardId = cleanText(req.cardId, 16);
    const sheet = deckSheet(link.deck);
    const after = wins.find(w => Number(w.id) > Number(id));
    let row;
    if (after) {
      sheet.insertRowBefore(after.row);
      row = after.row;
    } else {
      row = Math.max(sheet.getLastRow(), 1) + 1;
    }
    sheet.getRange(row, 1, 1, 6).setNumberFormat('@')
      .setValues([[id, color, value, asText(signer), keyToMdy(key), asText(cardId)]]);
    return linkInfo(link, readWins(link.deck));
  },
};

// ── Audit actions ──────────────────────────────────────────────────────────

const AUDIT_ACTIONS = {
  // The deck's whole history, plus which cards have been approved or retired
  auditOpen(link) {
    return auditInfo(link, readWins(link.deck));
  },

  // Save one color and value's fixes and approvals, all at once.
  // ops: { op: 'edit', id, before, signer, date, cardId } | { op: 'add', signer, date, cardId }
  //    | { op: 'delete', id, before }. `before` is the win as the page saw it; if the sheet
  //    has changed since (someone fixed it from a logging link), nothing is saved.
  // approve / retire: every card ID of this color and value that's approved, or retired
  //    (lost or ruined, so it won't be signed again).
  auditCommit(link, req) {
    const deck = DECKS[link.deck];
    const color = String(req.color || '');
    const value = String(req.value || '');
    if (deck.colors.indexOf(color) === -1) throw new ApiError('bad-request');
    if ((color === 'Wild' ? WILD_VALUES : VALUES).indexOf(value) === -1) throw new ApiError('bad-request');
    const ops = Array.isArray(req.ops) ? req.ops : [];
    const approve = Array.isArray(req.approve) ? req.approve.map(String) : [];
    const retire = Array.isArray(req.retire) ? req.retire.map(String) : [];
    if (ops.length > 2000 || approve.length > 2000 || retire.length > 2000) throw new ApiError('bad-request');

    const sheet = deckSheet(link.deck);
    let wins = readWins(link.deck);
    const seen = {};
    const edits = [], adds = [], deletes = [];
    // Check everything before writing anything
    ops.forEach(op => {
      op = op || {};
      let win = null;
      if (op.op === 'edit' || op.op === 'delete') {
        const id = String(op.id);
        if (seen[id]) throw new ApiError('bad-request');
        seen[id] = true;
        win = wins.find(w => w.id === id);
        if (!win) throw new ApiError('conflict');
        if (win.color !== color || win.value !== value) throw new ApiError('bad-request');
        if (!sameWin(win, op.before)) throw new ApiError('conflict');
      }
      if (op.op === 'delete') { deletes.push(win); return; }
      if (op.op !== 'edit' && op.op !== 'add') throw new ApiError('bad-request');
      const signer = cleanText(op.signer, MAX_NAME);
      const cardId = cleanText(op.cardId, 16);
      const key = dayKey(op.date);
      if (!signer || signer === '?' || !validCardId(cardId) || !key) throw new ApiError('bad-request');
      const row = [asText(signer), keyToMdy(key), asText(cardId)];
      if (win) edits.push({ win: win, row: row });
      else adds.push(row);
    });

    // Edits in place, then new wins at the bottom, then deletes from the bottom up
    // (so the row numbers we read are still right)
    edits.forEach(e => sheet.getRange(e.win.row, 4, 1, 3).setNumberFormat('@').setValues([e.row]));
    if (adds.length) {
      let id = wins.reduce((max, w) => Math.max(max, Number(w.id)), 0);
      const row = Math.max(sheet.getLastRow(), 1) + 1;
      sheet.getRange(row, 1, adds.length, 6).setNumberFormat('@')
        .setValues(adds.map(r => [String(++id), color, value, r[0], r[1], r[2]]));
    }
    deletes.sort((a, b) => b.row - a.row).forEach(w => sheet.deleteRow(w.row));
    wins = readWins(link.deck);

    // This color and value's rows become exactly `approve` and `retire`
    const signed = wins.filter(isSigned);
    const ours = Object.create(null);
    signed.forEach(w => { if (w.color === color && w.value === value) ours[w.cardId] = true; });
    const audit = auditSheet();
    const width = AUDIT_HEADERS.length;
    const rows = audit.getLastRow() > 1 ? audit.getRange(2, 1, audit.getLastRow() - 1, width).getDisplayValues() : [];
    const kept = rows.filter(r => {
      if (r[0] !== link.deck) return true;
      // Drop this color and value's old rows, and rows for cards nobody has signed any more
      if (r[1] === color && r[2] === value) return false;
      return signed.some(w => w.color === r[1] && w.value === r[2] && w.cardId === r[3]);
    }).map(r => r.map(asText));
    const now = Utilities.formatDate(new Date(), sheetTz(), 'yyyy-MM-dd HH:mm');
    const status = Object.create(null);
    approve.forEach(cid => { status[cid] = 'approved'; });
    retire.forEach(cid => { status[cid] = 'retired'; });
    Object.keys(status).filter(cid => ours[cid]).forEach(cid => {
      kept.push([link.deck, color, value, asText(cid), status[cid], fingerprint(signed, color, value, cid), now]);
    });
    if (rows.length) audit.getRange(2, 1, rows.length, width).clearContent();
    if (kept.length) audit.getRange(2, 1, kept.length, width).setNumberFormat('@').setValues(kept);
    return auditInfo(link, wins);
  },
};

function auditInfo(link, wins) {
  const audit = auditSheet();
  const rows = audit.getLastRow() > 1 ? audit.getRange(2, 1, audit.getLastRow() - 1, AUDIT_HEADERS.length).getDisplayValues() : [];
  return {
    deck: link.deck,
    expires: link.expires ? link.expires.toISOString() : null,
    wins: wins.map(w => ({
      id: w.id, color: w.color, value: w.value, signer: w.signer, date: w.date, cardId: w.cardId,
    })),
    approvals: rows.filter(r => r[0] === link.deck && (r[4] === 'approved' || r[4] === 'retired'))
      .map(r => ({ color: r[1], value: r[2], cardId: r[3], status: r[4], fingerprint: r[5] })),
  };
}

const isSigned = w => !!w.cardId && !!w.signer && w.signer !== '?';

// A card's signatures as one string: "12:Bella:2026-10-30|40:Cole:2026-11-02".
// Must match fingerprint() in audit.html.
function fingerprint(wins, color, value, cid) {
  return wins.filter(w => w.color === color && w.value === value && w.cardId === cid && isSigned(w))
    .sort((a, b) => Number(a.id) - Number(b.id))
    .map(w => w.id + ':' + w.signer + ':' + dayKey(w.date))
    .join('|');
}

function sameWin(win, before) {
  if (!before) return false;
  return win.signer === String(before.signer) && win.cardId === String(before.cardId) &&
    dayKey(win.date) === dayKey(before.date);
}

function linkInfo(link, wins) {
  return {
    deck: link.deck,
    from: keyToMdy(link.fromKey),
    to: keyToMdy(link.toKey),
    expires: link.expires ? link.expires.toISOString() : null,
    wins: wins.map(w => ({
      id: w.id, color: w.color, value: w.value, signer: w.signer, date: w.date, cardId: w.cardId,
    })),
    retired: retiredCards(link.deck, wins),
  };
}

// Cards retired on an audit link, as long as nobody has signed them since
function retiredCards(deck, wins) {
  const s = SpreadsheetApp.getActive().getSheetByName(AUDIT_TAB);
  if (!s || s.getLastRow() < 2) return [];
  const signed = wins.filter(isSigned);
  return s.getRange(2, 1, s.getLastRow() - 1, AUDIT_HEADERS.length).getDisplayValues()
    .filter(r => r[0] === deck && r[4] === 'retired' && r[5] === fingerprint(signed, r[1], r[2], r[3]))
    .map(r => ({ color: r[1], value: r[2], cardId: r[3] }));
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
    kind: String(r[8] == null ? '' : r[8]).trim(),
  };
}

// Finds the link for a code and checks it's still good, and that it's the right kind
// (an audit link on the logging page doesn't work, and the other way round).
// Opening a link for the first time starts its clock: 24 hours, or a week for audits.
function checkLink(code, opening, kind) {
  if (typeof code !== 'string' || !/^[A-Za-z0-9]{12}$/.test(code)) throw new ApiError('invalid-link');
  const sheet = linksSheet();
  const rows = sheet.getDataRange().getValues();
  const now = new Date();
  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][0]) !== code) continue;
    const link = linkFromRow(rows[i]);
    if (!DECKS[link.deck] || link.kind !== (kind || '')) throw new ApiError('invalid-link');
    if (link.kind !== 'audit' && !link.fromKey) throw new ApiError('invalid-link');
    if (!link.firstOpened) {
      if (!isDate(link.created) || now - link.created > UNOPENED_DAYS * 864e5) throw new ApiError('expired');
      if (!opening) throw new ApiError('invalid-link');
      const hours = link.kind === 'audit' ? AUDIT_HOURS : LINK_HOURS;
      link.expires = new Date(now.getTime() + hours * 36e5);
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
    l.getRange('H:I').setNumberFormat('@');
    l.getRange(1, 1, 1, LINK_HEADERS.length).setValues([LINK_HEADERS]).setFontWeight('bold');
    l.setFrozenRows(1);
    l.hideSheet();
    made.push(LINKS_TAB + ' (hidden)');
  }
  if (!ss.getSheetByName(AUDIT_TAB)) {
    auditSheet();
    made.push(AUDIT_TAB + ' (hidden)');
  }
  SpreadsheetApp.getUi().alert(made.length ? 'Added tabs: ' + made.join(', ') : 'All tabs already exist.');
}

function showLinkDialog() {
  const html = HtmlService.createHtmlOutputFromFile('LinkDialog').setWidth(460).setHeight(640);
  SpreadsheetApp.getUi().showModalDialog(html, 'Share links');
}

// What the dialog needs: decks with their default days, and the links still working
function dialogData() {
  const today = todayKey();
  const to = latestWeekdayOnOrBefore(today);
  const ss = SpreadsheetApp.getActive();
  return {
    defaultDeck: DEFAULT_DECK,
    decks: Object.keys(DECKS).map(id => {
      // A deck without its tab yet is listed, but can't get links until Set up tabs makes it
      if (!ss.getSheetByName(DECKS[id].tab)) return { id: id, name: DECKS[id].tab, from: to, to: to, missing: true };
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
    ensureLinkHeaders(sheet);
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

// Called from the dialog: one link that can review and fix every card in a deck, for a week
function createAuditLink(deck) {
  if (!DECKS[deck]) throw new Error('Unknown deck.');
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    const sheet = linksSheet();
    ensureLinkHeaders(sheet);
    const row = Math.max(sheet.getLastRow(), 1) + 1;
    const code = newCode();
    sheet.getRange(row, 1, 1, 3).setNumberFormat('@').setValues([[code, deck, '']]);
    sheet.getRange(row, 4, 1, 4).setValues([[new Date(), '', '', '']]);
    sheet.getRange(row, 8, 1, 2).setNumberFormat('@').setValues([['', 'audit']]);
    return { deck: DECKS[deck].tab, links: [{ url: AUDIT_URL + code, label: 'Audit link' }], active: activeLinks() };
  } finally {
    lock.releaseLock();
  }
}

// Links tabs made before To and Kind existed get their headers
function ensureLinkHeaders(sheet) {
  if (!sheet.getRange(1, 8).getValue()) sheet.getRange(1, 8).setValue('To').setFontWeight('bold');
  if (!sheet.getRange(1, 9).getValue()) sheet.getRange(1, 9).setValue('Kind').setFontWeight('bold');
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
    .filter(l => DECKS[l.deck] && (l.fromKey || l.kind === 'audit') && linkUsable(l, now))
    .reverse()
    .map(l => {
      const audit = l.kind === 'audit';
      return {
        code: l.code,
        url: (audit ? AUDIT_URL : SITE_URL) + l.code,
        deck: DECKS[l.deck].tab,
        label: audit ? 'Audit link' : rangeLabel(l.fromKey, l.toKey),
        status: l.firstOpened
          ? 'Works until ' + Utilities.formatDate(l.expires, tz, audit ? 'EEE MMM d h:mm a' : 'EEE h:mm a')
          : audit ? 'Not opened yet · works 7 days once opened' : 'Not opened yet',
      };
    });
}

// The latest day in a deck's wins or its links, as a key
function latestDay(deck) {
  let latest = '';
  readWins(deck).forEach(w => { const k = dayKey(w.date); if (k > latest) latest = k; });
  linksSheet().getDataRange().getValues().slice(1).forEach(r => {
    if (String(r[1]) !== deck) return;
    const l = linkFromRow(r);
    if (l.kind === 'audit') return;
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

// The Audit tab, made (hidden) the first time it's needed
function auditSheet() {
  const ss = SpreadsheetApp.getActive();
  let s = ss.getSheetByName(AUDIT_TAB);
  if (s) return s;
  s = ss.insertSheet(AUDIT_TAB);
  s.getRange('A:G').setNumberFormat('@');
  s.getRange(1, 1, 1, AUDIT_HEADERS.length).setValues([AUDIT_HEADERS]).setFontWeight('bold');
  s.setFrozenRows(1);
  s.hideSheet();
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
