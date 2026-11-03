/**
 * Wagyu Stock Check — Google Apps Script backend
 *
 * 1. Paste this into Extensions > Apps Script of a new Google Sheet (file: Code.gs)
 * 2. Add an HTML file named "Index" and paste the page code into it
 * 3. Run setup() once and allow access
 * 4. Deploy > New deployment > Web app
 *      Execute as: Me    Who has access: Anyone
 * 5. Share the web app URL. Nobody needs to log in to use it.
 *
 * The page counts one product (Wagyu). Its average weight, price and the
 * manager PIN live in the Settings tab, so they can change without a redeploy.
 */

const TZ = 'Australia/Brisbane';
const PRODUCT = 'Wagyu';
const LOG = 'Log';
const SETTINGS = 'Settings';
const DATE_FMT = 'yyyy-mm-dd hh:mm:ss';
const SHOW_FMT = 'yyyy-MM-dd HH:mm:ss';

/** Settings rows: [name, default value, number format, what it does]. */
const SETTING_ROWS = [
  ['Manager PIN', '1234', '@', 'Change this. 10 wrong tries locks the manager screen for 15 minutes.'],
  ['Avg kg per unit', 0.1425, '0.0000', 'Average weight of one ' + PRODUCT + ' unit in kg. Used for the approx kg.'],
  ['Price per kg', '', '$0.00', 'Optional. When set, differences also show in dollars.'],
  ['Page title', PRODUCT + ' Stock Check', '@', 'Shown at the top of the page.']
];
const DEFAULT_AVG_KG = 0.1425;

const COL = {
  id: 1, at: 2, by: 3, product: 4, count: 5, avgKg: 6, totalKg: 7,
  prevCount: 8, prevAt: 9, change: 10, note: 11,
  stock: 12, diff: 13, diffKg: 14, price: 15, diffCost: 16, mgrAt: 17
};

const LOG_HEADERS = [
  'Entry ID', 'Counted at', 'Counted by', 'Product', 'Count (units)', 'Avg kg / unit',
  'Approx total kg', 'Previous count', 'Previous count at', 'Change (units)', 'Note',
  'Manager stock (units)', 'Difference (units)', 'Difference (kg)', 'Price ($/kg)',
  'Difference ($)', 'Manager updated at'
];

/* ---------------------------------------------------------------- setup */

function setup() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  ss.setSpreadsheetTimeZone(TZ);

  // Log
  let log = ss.getSheetByName(LOG);
  if (!log) log = ss.insertSheet(LOG, 0);
  if (log.getLastRow() === 0) {
    log.getRange(1, 1, 1, LOG_HEADERS.length).setValues([LOG_HEADERS])
      .setFontWeight('bold').setBackground('#2b2422').setFontColor('#ffffff')
      .setWrap(true).setVerticalAlignment('middle');
    log.setFrozenRows(1);
    log.setRowHeight(1, 42);
    log.setColumnWidths(1, LOG_HEADERS.length, 110);
    [COL.id, COL.at, COL.prevAt, COL.mgrAt].forEach(c => log.setColumnWidth(c, 160));
    log.setColumnWidth(COL.note, 220);
  }
  applyRules_(log);

  // Settings
  let set = ss.getSheetByName(SETTINGS);
  if (!set) set = ss.insertSheet(SETTINGS);
  if (set.getLastRow() === 0) {
    set.getRange(1, 1, 1, 3).setValues([['Setting', 'Value', 'What it does']])
      .setFontWeight('bold').setBackground('#2b2422').setFontColor('#ffffff');
    set.setFrozenRows(1);
    set.setColumnWidth(1, 160);
    set.setColumnWidth(2, 190);
    set.setColumnWidth(3, 440);
  }
  ensureSettings_(set);

  // Remove the blank starter tab if it is still empty
  ['Sheet1', '시트1'].forEach(n => {
    const s = ss.getSheetByName(n);
    if (s && s.getLastRow() === 0 && ss.getSheets().length > 1) ss.deleteSheet(s);
  });

  ss.setActiveSheet(log);
  return 'Setup done';
}

/** Add any missing setting rows. Values that are already there are left alone. */
function ensureSettings_(set) {
  const have = {};
  const last = set.getLastRow();
  if (last > 0) {
    set.getRange(1, 1, last, 1).getValues()
      .forEach(r => { have[String(r[0]).trim().toLowerCase()] = true; });
  }
  SETTING_ROWS.forEach(([name, value, format, help]) => {
    if (have[name.toLowerCase()]) return;
    const r = set.getLastRow() + 1;
    set.getRange(r, 2).setNumberFormat(format);
    set.getRange(r, 1, 1, 3).setValues([[name, value, help]]);
    set.getRange(r, 3).setFontColor('#6b615a');
  });
}

function applyRules_(log) {
  const rows = Math.max(log.getMaxRows() - 1, 1);
  const neg = SpreadsheetApp.newConditionalFormatRule()
    .whenNumberLessThan(0)
    .setBackground('#fde4e1').setFontColor('#a8231a')
    .setRanges([
      log.getRange(2, COL.diff, rows, 2),
      log.getRange(2, COL.diffCost, rows, 1)
    ])
    .build();
  const keep = log.getConditionalFormatRules().filter(r => {
    const ranges = r.getRanges();
    return !ranges.some(g => g.getColumn() === COL.diff || g.getColumn() === COL.diffCost);
  });
  keep.push(neg);
  log.setConditionalFormatRules(keep);
}

function ensureRows_(sheet, lastNeeded) {
  const max = sheet.getMaxRows();
  if (lastNeeded > max) {
    sheet.insertRowsAfter(max, Math.max(1000, lastNeeded - max));
    applyRules_(sheet);
  }
}

/* ---------------------------------------------------------------- web app */

function doGet() {
  const cfg = config_();
  const t = HtmlService.createTemplateFromFile('Index');
  t.pageTitle = cfg.title;
  t.configJson = JSON.stringify({ product: PRODUCT, avgKg: cfg.avgKg }).replace(/</g, '\\u003c');
  return t.evaluate()
    .setTitle(cfg.title)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, viewport-fit=cover');
}

/**
 * Save one Wagyu count.
 * payload = { by: 'Sam', count: 8, note: 'put out 6 from the cool room' }
 */
function submitCount(payload) {
  if (!payload || typeof payload !== 'object') throw new Error('Nothing to save.');
  const by = cleanText_(payload.by, 40);
  if (!by) throw new Error('Please enter your name.');
  const note = cleanText_(payload.note, 200);
  const raw = payload.count;
  const count = raw === '' || raw === null || raw === undefined ? NaN : Number(raw);
  if (!Number.isInteger(count) || count < 0 || count > 9999) {
    throw new Error('Check the count. Use a whole number from 0 to 9999.');
  }
  const cfg = config_();

  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const log = logSheet_();
    const now = new Date();
    const prev = lastCount_(log);
    const r = log.getLastRow() + 1;
    ensureRows_(log, r);

    const id = Utilities.formatDate(now, TZ, 'yyMMdd-HHmmss') + '-' + Utilities.getUuid().slice(0, 4);
    log.getRange(r, 1, 1, LOG_HEADERS.length).setValues([[
      id,
      now,
      by,
      PRODUCT,
      count,
      cfg.avgKg,
      '=IF(E' + r + '="","",ROUND(E' + r + '*F' + r + ',3))',
      prev ? prev.count : '',
      prev ? prev.at : '',
      '=IF(H' + r + '="","",E' + r + '-H' + r + ')',
      note,
      '',
      '=IF(L' + r + '="","",E' + r + '-L' + r + ')',
      '=IF(M' + r + '="","",ROUND(M' + r + '*F' + r + ',3))',
      cfg.price,
      '=IF(OR(M' + r + '="",O' + r + '=""),"",ROUND(M' + r + '*F' + r + '*O' + r + ',2))',
      ''
    ]]);
    formatRows_(log, r, 1);
    SpreadsheetApp.flush();

    return { ok: true, count: count, at: Utilities.formatDate(now, TZ, SHOW_FMT) };
  } finally {
    lock.releaseLock();
  }
}

/* ---------------------------------------------------------------- manager */

function getRecent(pin, limit) {
  checkPin_(pin);
  const log = logSheet_();
  const last = log.getLastRow();
  if (last < 2) return { rows: [] };

  const span = Math.min(last - 1, 3000);
  const vals = log.getRange(last - span + 1, 1, span, LOG_HEADERS.length).getValues();
  const max = Math.max(10, Math.min(Number(limit) || 150, 500));

  const rows = [];
  vals.forEach((r, i) => {
    if (r[COL.id - 1] === '' || String(r[COL.product - 1]) !== PRODUCT) return;
    rows.push({
      idx: i,
      id: String(r[COL.id - 1]),
      atMs: r[COL.at - 1] instanceof Date ? r[COL.at - 1].getTime() : 0,
      at: fmt_(r[COL.at - 1]),
      by: String(r[COL.by - 1]),
      count: num_(r[COL.count - 1]),
      avgKg: num_(r[COL.avgKg - 1]),
      prevCount: num_(r[COL.prevCount - 1]),
      prevAt: fmt_(r[COL.prevAt - 1]),
      note: String(r[COL.note - 1]),
      stock: num_(r[COL.stock - 1]),
      price: num_(r[COL.price - 1])
    });
  });
  rows.sort((a, b) => (b.atMs - a.atMs) || (b.idx - a.idx));
  return { rows: rows.slice(0, max) };
}

function setStock(pin, id, value) {
  checkPin_(pin);
  let v = '';
  if (value !== '' && value !== null && value !== undefined) {
    v = Number(value);
    if (!Number.isInteger(v) || v < 0 || v > 9999) throw new Error('Enter a whole number (0 or more).');
  }
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const log = logSheet_();
    const last = log.getLastRow();
    if (last < 2) throw new Error('Entry not found. Refresh the list.');
    const cell = log.getRange(2, COL.id, last - 1, 1)
      .createTextFinder(String(id)).matchEntireCell(true).findNext();
    if (!cell) throw new Error('Entry not found. Refresh the list.');
    const r = cell.getRow();
    const now = new Date();
    log.getRange(r, COL.stock).setValue(v);
    log.getRange(r, COL.mgrAt).setValue(v === '' ? '' : now).setNumberFormat(DATE_FMT);
    SpreadsheetApp.flush();
    return { ok: true, id: String(id), stock: v, at: v === '' ? '' : Utilities.formatDate(now, TZ, SHOW_FMT) };
  } finally {
    lock.releaseLock();
  }
}

/* ---------------------------------------------------------------- helpers */

function logSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let log = ss.getSheetByName(LOG);
  if (!log) { setup(); log = ss.getSheetByName(LOG); }
  return log;
}

/** Settings tab as { pin, avgKg, price, title }, with safe fallbacks. */
function config_() {
  const s = {};
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SETTINGS);
  if (sh && sh.getLastRow() > 0) {
    sh.getRange(1, 1, sh.getLastRow(), 2).getValues().forEach(r => {
      const k = String(r[0]).trim().toLowerCase();
      if (k && !(k in s)) s[k] = r[1];
    });
  }
  const avg = num_(s['avg kg per unit']);
  const price = num_(s['price per kg']);
  const title = String(s['page title'] == null ? '' : s['page title']).trim();
  return {
    pin: String(s['manager pin'] == null ? '' : s['manager pin']).trim(),
    avgKg: avg !== null && avg > 0 && avg < 50 ? avg : DEFAULT_AVG_KG,
    price: price !== null && price > 0 ? price : '',
    title: title || PRODUCT + ' Stock Check'
  };
}

/** Latest Wagyu count by time, so sorting the sheet doesn't break it. */
function lastCount_(log) {
  const last = log.getLastRow();
  if (last < 2) return null;
  const span = Math.min(last - 1, 5000);
  const vals = log.getRange(last - span + 1, COL.at, span, COL.count - COL.at + 1).getValues();
  let best = null;
  vals.forEach(r => {
    const at = r[0];
    const n = r[COL.count - COL.at];
    if (!(at instanceof Date) || String(r[COL.product - COL.at]) !== PRODUCT || n === '') return;
    if (!best || at.getTime() >= best.at.getTime()) best = { at: at, count: n };
  });
  return best;
}

function formatRows_(log, start, n) {
  log.getRange(start, COL.at, n, 1).setNumberFormat(DATE_FMT);
  log.getRange(start, COL.prevAt, n, 1).setNumberFormat(DATE_FMT);
  log.getRange(start, COL.mgrAt, n, 1).setNumberFormat(DATE_FMT);
  log.getRange(start, COL.avgKg, n, 2).setNumberFormat('0.000');
  log.getRange(start, COL.diffKg, n, 1).setNumberFormat('0.000');
  log.getRange(start, COL.price, n, 1).setNumberFormat('$0.00');
  log.getRange(start, COL.diffCost, n, 1).setNumberFormat('$0.00;-$0.00');
}

function checkPin_(pin) {
  const cache = CacheService.getScriptCache();
  const fails = Number(cache.get('pinFails') || 0);
  if (fails >= 10) throw new Error('Too many wrong PINs. Try again in 15 minutes.');
  const real = config_().pin;
  if (!real) throw new Error('Set a Manager PIN in the Settings tab first.');
  if (String(pin == null ? '' : pin).trim() !== real) {
    cache.put('pinFails', String(fails + 1), 900);
    throw new Error('Wrong PIN.');
  }
  if (fails) cache.remove('pinFails');
}

/** Trim, collapse whitespace, and stop text from being read as a formula. */
function cleanText_(s, max) {
  const t = String(s == null ? '' : s)
    .replace(/[\u0000-\u001f\u007f]/g, ' ')
    .replace(/\s+/g, ' ')
    .trim()
    .slice(0, max);
  return /^[=+\-@]/.test(t) ? "'" + t : t;
}

function fmt_(v) {
  return v instanceof Date ? Utilities.formatDate(v, TZ, SHOW_FMT) : (v === '' || v == null ? '' : String(v));
}

/** Number from a cell; tolerates "$89.90" or "0.15 kg" typed as text. */
function num_(v) {
  if (v === '' || v === null || v === undefined) return null;
  if (typeof v === 'number') return isFinite(v) ? v : null;
  const t = String(v).replace(/\$|kg|\s/gi, '');
  if (t === '') return null;
  const n = Number(t);
  return isFinite(n) ? n : null;
}
