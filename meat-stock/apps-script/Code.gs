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
 * Staff enter two numbers: the Wagyu pieces they counted in the store and the
 * stock the system shows in kg. The Log turns the count into kg (count × average
 * weight per piece) and compares it with the system. A gap of about one piece or
 * more turns red (Settings → "Red when off by (kg)"); smaller gaps are normal
 * differences in piece weight.
 */

const TZ = 'Australia/Brisbane';
const PRODUCT = 'Wagyu';
const LOG = 'Log';
const SETTINGS = 'Settings';
const DATE_FMT = 'yyyy-mm-dd hh:mm:ss';
const SHOW_FMT = 'yyyy-MM-dd HH:mm:ss';

const DEFAULT_AVG_KG = 0.1455;      // 19 pieces weighed on 4 Nov 2026: 2.764 kg ÷ 19
const OLD_DEFAULT_AVG_KG = 0.1425;  // the first guess; replaced once if it is still in Settings
const DEFAULT_TOL_KG = 0.1;         // about 70% of one piece
const SCHEMA = '3';                 // bump when the Log layout or Settings rows change
const RULES_VERSION = '3';          // bump when applyRules_ changes

/** Settings rows: [name, default value, number format, what it does]. */
const SETTING_ROWS = [
  ['Manager PIN', '1234', '@', 'Change this. 10 wrong tries locks the manager screen for 15 minutes.'],
  ['Avg kg per unit', DEFAULT_AVG_KG, '0.0000', 'Average weight of one ' + PRODUCT + ' piece in kg. 19 pieces weighed 2.764 kg on 4 Nov 2026, so 0.1455.'],
  ['Red when off by (kg)', DEFAULT_TOL_KG, '0.000', 'Counted kg vs system kg: a gap this size or bigger turns red. One piece is about 0.145 kg; smaller gaps are normal weight differences.'],
  ['Price per kg', '', '$0.00', 'Optional. When set, differences also show in dollars.'],
  ['Page title', PRODUCT + ' Stock Check', '@', 'Shown at the top of the page.']
];

const COL = {
  id: 1, at: 2, by: 3, product: 4, count: 5, avgKg: 6, countKg: 7, sysKg: 8,
  diffKg: 9, diffUnits: 10, price: 11, diffCost: 12, prevCount: 13, prevAt: 14, change: 15, note: 16
};

const LOG_HEADERS = [
  'Entry ID', 'Counted at', 'Counted by', 'Product', 'Count (units)', 'Avg kg / unit',
  'Counted kg (approx)', 'System stock (kg)', 'Difference (kg)', 'Difference (units)',
  'Price ($/kg)', 'Difference ($)', 'Previous count', 'Previous count at', 'Change (units)', 'Note'
];

/* ---------------------------------------------------------------- setup */

function setup() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  ss.setSpreadsheetTimeZone(TZ);

  // Settings first: the red rule on the Log reads its threshold from here
  let set = ss.getSheetByName(SETTINGS);
  if (!set) set = ss.insertSheet(SETTINGS);
  if (set.getLastRow() === 0) {
    set.getRange(1, 1, 1, 3).setValues([['Setting', 'Value', 'What it does']])
      .setFontWeight('bold').setBackground('#307029').setFontColor('#ffffff');
    set.setFrozenRows(1);
    set.setColumnWidth(1, 170);
    set.setColumnWidth(2, 190);
    set.setColumnWidth(3, 480);
  }
  ensureSettings_(set);

  // Log (a Log made by older code is kept as "Log (old)" and a new one starts)
  let log = ss.getSheetByName(LOG);
  if (log && isOldLog_(log)) log = retireLog_(ss, log);
  if (!log) log = ss.insertSheet(LOG, 0);
  if (log.getLastRow() === 0) {
    log.getRange(1, 1, 1, LOG_HEADERS.length).setValues([LOG_HEADERS])
      .setFontWeight('bold').setBackground('#307029').setFontColor('#ffffff')
      .setWrap(true).setVerticalAlignment('middle');
    log.setFrozenRows(1);
    log.setRowHeight(1, 42);
    log.setColumnWidths(1, LOG_HEADERS.length, 110);
    [COL.id, COL.at, COL.prevAt].forEach(c => log.setColumnWidth(c, 160));
    log.setColumnWidth(COL.note, 220);
  }
  applyRules_(log);

  // Remove the blank starter tab if it is still empty
  ['Sheet1', '시트1'].forEach(n => {
    const s = ss.getSheetByName(n);
    if (s && s.getLastRow() === 0 && ss.getSheets().length > 1) ss.deleteSheet(s);
  });

  ss.setActiveSheet(log);
  return 'Setup done';
}

/** Add missing setting rows and refresh the help text. Values already there are left alone. */
function ensureSettings_(set) {
  const rowOf = {};
  const last = set.getLastRow();
  if (last > 0) {
    set.getRange(1, 1, last, 1).getValues()
      .forEach((r, i) => { rowOf[String(r[0]).trim().toLowerCase()] = i + 1; });
  }
  SETTING_ROWS.forEach(([name, value, format, help]) => {
    const at = rowOf[name.toLowerCase()];
    if (at) { set.getRange(at, 3).setValue(help).setFontColor('#566352'); return; }
    const r = set.getLastRow() + 1;
    set.getRange(r, 2).setNumberFormat(format);
    set.getRange(r, 1, 1, 3).setValues([[name, value, help]]);
    set.getRange(r, 3).setFontColor('#566352');
  });
}

/** The first version guessed 0.1425 kg a piece. If that guess is still there, use the measured average. */
function updateOldAverage_(set) {
  const last = set.getLastRow();
  if (last < 1) return;
  const vals = set.getRange(1, 1, last, 2).getValues();
  for (let i = 0; i < vals.length; i++) {
    if (String(vals[i][0]).trim().toLowerCase() !== 'avg kg per unit') continue;
    const v = num_(vals[i][1]);
    if (v !== null && Math.abs(v - OLD_DEFAULT_AVG_KG) < 1e-9) set.getRange(i + 1, 2).setValue(DEFAULT_AVG_KG);
    return;
  }
}

/** A Log made by older code (manager stock in units) has "Previous count" in column H. */
function isOldLog_(log) {
  return log.getLastRow() > 0 && String(log.getRange(1, 8).getValue()).trim() === 'Previous count';
}

/** Keep the old rows in a renamed tab, or clear the tab if it has none, so the new Log starts clean. */
function retireLog_(ss, log) {
  if (log.getLastRow() <= 1) {
    log.clear();
    log.clearConditionalFormatRules();
    return log;
  }
  let name = LOG + ' (old)';
  for (let i = 2; ss.getSheetByName(name); i++) name = LOG + ' (old ' + i + ')';
  log.setName(name);
  return null;
}

/**
 * Red background on Difference (kg / units / $) when the counted kg and the
 * system kg are apart by the "Red when off by (kg)" amount or more, either way.
 * Blank cells stay plain.
 */
function applyRules_(log, tolKg) {
  const tol = tolKg == null ? config_().tolKg : tolKg;
  const rows = Math.max(log.getMaxRows() - 1, 1);
  const ref = '$' + colLetter_(COL.diffKg) + '2';
  const formula = '=AND(ISNUMBER(' + ref + '),ABS(' + ref + ')>=' + tol + ')';
  const rule = range => SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied(formula)
    .setBackground('#fde4e1').setFontColor('#a8231a')
    .setRanges([range])
    .build();
  const mine = [COL.diffKg, COL.diffUnits, COL.diffCost];
  const keep = log.getConditionalFormatRules()
    .filter(r => !r.getRanges().some(g => mine.indexOf(g.getColumn()) !== -1));
  keep.push(rule(log.getRange(2, COL.diffKg, rows, 2)), rule(log.getRange(2, COL.diffCost, rows, 1)));
  log.setConditionalFormatRules(keep);
  PropertiesService.getScriptProperties().setProperty('rules', rulesKey_(tol));
}

function rulesKey_(tol) { return RULES_VERSION + ':' + tol; }

/** Keep the red rule in line with Settings (re-applied after "Red when off by (kg)" changes). */
function ensureRules_() {
  try {
    const tol = config_().tolKg;
    if (PropertiesService.getScriptProperties().getProperty('rules') === rulesKey_(tol)) return;
    applyRules_(logSheet_(), tol);
  } catch (e) {
    console.error(e); // formatting must never stop the page from opening
  }
}

/**
 * Runs when the page opens. A sheet made with older code is brought up to date
 * once (new Log layout, new Settings rows, the measured average weight), then
 * the red rule is kept in line with Settings.
 */
function prepare_() {
  try {
    const props = PropertiesService.getScriptProperties();
    if (props.getProperty('schema') !== SCHEMA) {
      const lock = LockService.getScriptLock();
      if (lock.tryLock(15000)) {
        try {
          if (props.getProperty('schema') !== SCHEMA) {
            setup();   // new Settings rows, new Log layout, red rule
            const set = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SETTINGS);
            if (set) updateOldAverage_(set);
            props.setProperty('schema', SCHEMA);
          }
        } finally {
          lock.releaseLock();
        }
      }
    }
  } catch (e) {
    console.error(e);
  }
  ensureRules_();
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
  prepare_();
  const cfg = config_();
  const t = HtmlService.createTemplateFromFile('Index');
  t.pageTitle = cfg.title;
  t.configJson = JSON.stringify({ product: PRODUCT }).replace(/</g, '\\u003c');
  return t.evaluate()
    .setTitle(cfg.title)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, viewport-fit=cover');
}

/**
 * Save one Wagyu count.
 * payload = { by: 'Sam', count: 19, systemKg: 2.764, note: 'put out 6 from the cool room' }
 */
function submitCount(payload) {
  if (!payload || typeof payload !== 'object') throw new Error('Nothing to save.');
  const by = cleanText_(payload.by, 40);
  if (!by) throw new Error('Please enter your name.');
  const note = cleanText_(payload.note, 200);
  const count = wholeNumber_(payload.count);
  if (count === null) throw new Error('Check the count. Use a whole number from 0 to 9999.');
  const sysKg = kgNumber_(payload.systemKg);
  if (sysKg === null) throw new Error('Check the system stock. Use kg, like 2.764.');
  const cfg = config_();

  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const log = logSheet_();
    const now = new Date();
    const prev = lastCount_(log);
    const r = log.getLastRow() + 1;
    ensureRows_(log, r);

    const c = col => colLetter_(col) + r;
    const id = Utilities.formatDate(now, TZ, 'yyMMdd-HHmmss') + '-' + Utilities.getUuid().slice(0, 4);
    log.getRange(r, 1, 1, LOG_HEADERS.length).setValues([[
      id,
      now,
      by,
      PRODUCT,
      count,
      cfg.avgKg,
      '=IF(' + c(COL.count) + '="","",ROUND(' + c(COL.count) + '*' + c(COL.avgKg) + ',3))',
      sysKg,
      '=IF(OR(' + c(COL.countKg) + '="",' + c(COL.sysKg) + '=""),"",ROUND(' + c(COL.countKg) + '-' + c(COL.sysKg) + ',3))',
      '=IF(OR(' + c(COL.diffKg) + '="",N(' + c(COL.avgKg) + ')=0),"",ROUND(' + c(COL.diffKg) + '/' + c(COL.avgKg) + ',1))',
      cfg.price,
      '=IF(OR(' + c(COL.diffKg) + '="",' + c(COL.price) + '=""),"",ROUND(' + c(COL.diffKg) + '*' + c(COL.price) + ',2))',
      prev ? prev.count : '',
      prev ? prev.at : '',
      '=IF(' + c(COL.prevCount) + '="","",' + c(COL.count) + '-' + c(COL.prevCount) + ')',
      note
    ]]);
    formatRows_(log, r, 1);
    SpreadsheetApp.flush();

    return { ok: true, count: count, systemKg: sysKg, at: Utilities.formatDate(now, TZ, SHOW_FMT) };
  } finally {
    lock.releaseLock();
  }
}

/* ---------------------------------------------------------------- manager */

function getRecent(pin, limit) {
  const cfg = checkPin_(pin);
  const log = logSheet_();
  const last = log.getLastRow();
  if (last < 2) return { rows: [], total: 0, tolKg: cfg.tolKg };

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
      systemKg: num_(r[COL.sysKg - 1]),
      price: num_(r[COL.price - 1]),
      prevCount: num_(r[COL.prevCount - 1]),
      prevAt: fmt_(r[COL.prevAt - 1]),
      note: String(r[COL.note - 1])
    });
  });
  rows.sort((a, b) => (b.atMs - a.atMs) || (b.idx - a.idx));
  return { rows: rows.slice(0, max), total: rows.length, tolKg: cfg.tolKg };
}

/**
 * Manager: delete every count in the Log tab and start again.
 * The header row, the Settings tab and the PIN stay. The page asks "Are you sure?" first.
 */
function resetAll(pin) {
  checkPin_(pin);
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const log = logSheet_();
    const last = log.getLastRow();
    const n = last > 1 ? last - 1 : 0;
    if (n) log.getRange(2, 1, n, Math.max(log.getLastColumn(), LOG_HEADERS.length)).clearContent();
    applyRules_(log);
    SpreadsheetApp.flush();
    return { ok: true, cleared: n };
  } finally {
    lock.releaseLock();
  }
}

/* ---------------------------------------------------------------- helpers */

function logSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const log = ss.getSheetByName(LOG);
  if (log && log.getLastRow() > 0 && !isOldLog_(log)) return log;
  setup();
  return ss.getSheetByName(LOG);
}

/** Settings tab as { pin, avgKg, tolKg, price, title }, with safe fallbacks. */
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
  const tol = num_(s['red when off by (kg)']);
  const price = num_(s['price per kg']);
  const title = String(s['page title'] == null ? '' : s['page title']).trim();
  return {
    pin: String(s['manager pin'] == null ? '' : s['manager pin']).trim(),
    avgKg: avg !== null && avg > 0 && avg < 50 ? avg : DEFAULT_AVG_KG,
    tolKg: tol !== null && tol > 0 && tol < 100 ? tol : DEFAULT_TOL_KG,
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
  log.getRange(start, COL.avgKg, n, 1).setNumberFormat('0.0000');
  log.getRange(start, COL.countKg, n, 2).setNumberFormat('0.000');
  log.getRange(start, COL.diffKg, n, 1).setNumberFormat('+0.000;-0.000;0.000');
  log.getRange(start, COL.diffUnits, n, 1).setNumberFormat('+0.0;-0.0;0.0');
  log.getRange(start, COL.price, n, 1).setNumberFormat('$0.00');
  log.getRange(start, COL.diffCost, n, 1).setNumberFormat('+$0.00;-$0.00;$0.00');
}

/** Checks the manager PIN and returns the Settings so callers don't read them twice. */
function checkPin_(pin) {
  const cache = CacheService.getScriptCache();
  const fails = Number(cache.get('pinFails') || 0);
  if (fails >= 10) throw new Error('Too many wrong PINs. Try again in 15 minutes.');
  const cfg = config_();
  if (!cfg.pin) throw new Error('Set a Manager PIN in the Settings tab first.');
  if (String(pin == null ? '' : pin).trim() !== cfg.pin) {
    cache.put('pinFails', String(fails + 1), 900);
    throw new Error('Wrong PIN.');
  }
  if (fails) cache.remove('pinFails');
  return cfg;
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

/** Whole number 0–9999, or null. */
function wholeNumber_(v) {
  if (v === '' || v === null || v === undefined) return null;
  const n = Number(v);
  return Number.isInteger(n) && n >= 0 && n <= 9999 ? n : null;
}

/** Kilograms with up to 3 decimals ("2,764" works too), or null. */
function kgNumber_(v) {
  if (v === '' || v === null || v === undefined) return null;
  const t = typeof v === 'number' ? v : String(v).trim().replace(',', '.');
  if (t === '') return null;
  const n = Number(t);
  if (!isFinite(n) || n < -9999 || n > 9999) return null;
  return Math.round(n * 1000) / 1000;
}

function colLetter_(n) {
  let s = '';
  for (; n > 0; n = Math.floor((n - 1) / 26)) s = String.fromCharCode(65 + (n - 1) % 26) + s;
  return s;
}
