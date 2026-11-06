/**
 * Meat Stock Check — Google Apps Script backend
 *
 * 1. Paste this into Extensions > Apps Script of a new Google Sheet (file: Code.gs)
 * 2. Add an HTML file named "Index" and paste the page code into it
 * 3. Run setup() once and allow access
 * 4. Deploy > New deployment > Web app
 *      Execute as: Me    Who has access: Anyone
 * 5. Share the web app URL. Nobody needs to log in to use it.
 *
 * Staff count each product in the store (pieces) and copy the stock the system
 * shows for it (kg or pieces, set per product). The Log compares the two and a
 * gap at or over the product's "Red when off by" turns red.
 *
 * The Products tab lists what gets counted: name, Korean name, system unit
 * (kg / pcs), average kg per piece (kg products), red threshold, price and
 * status (Active / Check later / Off). Changes there show up without a redeploy.
 */

const TZ = 'Australia/Brisbane';
const LOG = 'Log';
const SETTINGS = 'Settings';
const PRODUCTS = 'Products';
const DATE_FMT = 'yyyy-mm-dd hh:mm:ss';
const SHOW_FMT = 'yyyy-MM-dd HH:mm:ss';
const SCHEMA = '4';          // bump when the sheet layout changes
const RULES_VERSION = '4';   // bump when applyRules_ changes

const WAGYU = 'Wagyu';
const WAGYU_AVG_KG = 0.1455;      // 19 pieces weighed 2.764 kg on 4 Nov 2026
const OLD_WAGYU_AVG_KG = 0.1425;  // the first guess
const TITLE = 'Meat Stock Check';
const OLD_TITLE = 'Wagyu Stock Check';

/** Products tab: one row per product. */
const PRODUCT_HEADERS = ['Product', '한글 이름', 'System unit (kg / pcs)', 'Avg kg per piece',
  'Red when off by', 'Price per system unit', 'Status (Active / Check later / Off)', 'Note for staff'];
const P = { name: 0, ko: 1, unit: 2, avg: 3, tol: 4, price: 5, status: 6, note: 7 };
const DEFAULT_PRODUCTS = [
  [WAGYU, '와규', 'kg', WAGYU_AVG_KG, 0.1, '', 'Active', ''],
  ['Pork belly', '냉장 삼겹살', 'pcs', '', 1, '', 'Active', ''],
  ['Wagyu soup cuts', '와규 국거리', 'kg', '', 0.1, '', 'Check later', '현재 매장 재고 없음 · 나중에 확인']
];

/** Settings rows: [name, default value, number format, what it does]. */
const SETTING_ROWS = [
  ['Manager PIN', '1234', '@', 'Change this. 10 wrong tries locks the manager screen for 15 minutes.'],
  ['Page title', TITLE, '@', 'Shown at the top of the page.']
];
/** Wagyu-only versions kept these in Settings; they now live on the Wagyu row of Products. */
const MOVED_SETTINGS = ['avg kg per unit', 'red when off by (kg)', 'price per kg'];

const COL = {
  id: 1, at: 2, by: 3, product: 4, count: 5, avgKg: 6, counted: 7, sys: 8, diff: 9, diffPcs: 10,
  price: 11, diffCost: 12, prevCount: 13, prevAt: 14, change: 15, note: 16, unit: 17, tol: 18
};

const LOG_HEADERS = [
  'Entry ID', 'Counted at', 'Counted by', 'Product', 'Count (pcs)', 'Avg kg / piece',
  'Counted (system unit)', 'System stock', 'Difference', 'Difference (pcs)',
  'Price (per system unit)', 'Difference ($)', 'Previous count', 'Previous count at', 'Change (pcs)', 'Note',
  'System unit', 'Red when off by'
];

/* ---------------------------------------------------------------- setup */

function setup() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  ss.setSpreadsheetTimeZone(TZ);

  // Settings
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

  // Products (seeded once; Wagyu takes the numbers that used to sit in Settings)
  if (!ss.getSheetByName(PRODUCTS)) seedProducts_(ss.insertSheet(PRODUCTS), set);
  removeMovedSettings_(set);

  // Log: a Log from the first version is kept as "Log (old)"; the Wagyu-only Log is extended in place
  let log = ss.getSheetByName(LOG);
  if (log && isFirstLog_(log)) log = retireLog_(ss, log);
  if (log && isWagyuOnlyLog_(log)) upgradeWagyuOnlyLog_(log);
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
  const rowOf = settingRows_(set);
  SETTING_ROWS.forEach(([name, value, format, help]) => {
    const at = rowOf[name.toLowerCase()];
    if (at) { set.getRange(at, 3).setValue(help).setFontColor('#566352'); return; }
    const r = set.getLastRow() + 1;
    set.getRange(r, 2).setNumberFormat(format);
    set.getRange(r, 1, 1, 3).setValues([[name, value, help]]);
    set.getRange(r, 3).setFontColor('#566352');
  });
}

/** { 'setting name in lower case': row number } */
function settingRows_(set) {
  const rowOf = {};
  const last = set.getLastRow();
  if (last > 0) {
    set.getRange(1, 1, last, 1).getValues()
      .forEach((r, i) => { const k = String(r[0]).trim().toLowerCase(); if (k && !rowOf[k]) rowOf[k] = i + 1; });
  }
  return rowOf;
}

function seedProducts_(prod, set) {
  const old = {};
  const rowOf = settingRows_(set);
  MOVED_SETTINGS.forEach(k => { if (rowOf[k]) old[k] = set.getRange(rowOf[k], 2).getValue(); });
  const rows = DEFAULT_PRODUCTS.map(r => r.slice());
  const avg = num_(old['avg kg per unit']);
  if (avg !== null && avg > 0 && avg < 50 && Math.abs(avg - OLD_WAGYU_AVG_KG) > 1e-9) rows[0][P.avg] = avg;
  const tol = num_(old['red when off by (kg)']);
  if (tol !== null && tol > 0) rows[0][P.tol] = tol;
  const price = num_(old['price per kg']);
  if (price !== null && price > 0) rows[0][P.price] = price;

  const n = PRODUCT_HEADERS.length;
  prod.getRange(1, 1, 1, n).setValues([PRODUCT_HEADERS])
    .setFontWeight('bold').setBackground('#307029').setFontColor('#ffffff').setWrap(true).setVerticalAlignment('middle');
  prod.setFrozenRows(1);
  prod.setRowHeight(1, 42);
  prod.setColumnWidths(1, n, 120);
  prod.setColumnWidth(P.name + 1, 150);
  prod.setColumnWidth(P.status + 1, 150);
  prod.setColumnWidth(P.note + 1, 280);
  prod.getRange(2, P.avg + 1, 50, 1).setNumberFormat('0.0000');
  prod.getRange(2, P.price + 1, 50, 1).setNumberFormat('$0.00');
  prod.getRange(2, P.unit + 1, 50, 1).setDataValidation(SpreadsheetApp.newDataValidation()
    .requireValueInList(['kg', 'pcs'], true).setAllowInvalid(false).build());
  prod.getRange(2, P.status + 1, 50, 1).setDataValidation(SpreadsheetApp.newDataValidation()
    .requireValueInList(['Active', 'Check later', 'Off'], true).setAllowInvalid(false).build());
  prod.getRange(2, 1, rows.length, n).setValues(rows);
}

function removeMovedSettings_(set) {
  const rowOf = settingRows_(set);
  MOVED_SETTINGS.map(k => rowOf[k]).filter(Boolean).sort((a, b) => b - a).forEach(r => set.deleteRow(r));
}

/** First version: manager stock in units, "Previous count" in column H. */
function isFirstLog_(log) {
  return log.getLastRow() > 0 && String(log.getRange(1, 8).getValue()).trim() === 'Previous count';
}

/** Wagyu-only version: same first 16 columns, everything in kg, no unit / threshold columns. */
function isWagyuOnlyLog_(log) {
  return log.getLastRow() > 0 && String(log.getRange(1, 8).getValue()).trim() === 'System stock (kg)';
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

/** Rename the headers, add System unit + Red when off by, and fill them in for the Wagyu rows already there. */
function upgradeWagyuOnlyLog_(log) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const wagyu = productList_(ss.getSheetByName(PRODUCTS)).filter(p => p.name === WAGYU)[0];
  const tol = wagyu ? wagyu.tol : 0.1;
  log.getRange(1, 1, 1, LOG_HEADERS.length).setValues([LOG_HEADERS]);
  log.getRange(1, COL.unit, 1, 2).setFontWeight('bold').setBackground('#307029').setFontColor('#ffffff').setWrap(true);
  const last = log.getLastRow();
  if (last > 1) {
    const ids = log.getRange(2, COL.id, last - 1, 1).getValues();
    log.getRange(2, COL.unit, last - 1, 2).setValues(ids.map(r => (r[0] === '' ? ['', ''] : ['kg', tol])));
    formatRows_(log, 2, ids.map(() => 'kg'));
  }
}

/**
 * Red background on Difference / Difference (pcs) / Difference ($) when the gap
 * is at or over that row's "Red when off by", either way. Blank cells stay plain.
 */
function applyRules_(log) {
  const rows = Math.max(log.getMaxRows() - 1, 1);
  const d = '$' + colLetter_(COL.diff) + '2';
  const t = '$' + colLetter_(COL.tol) + '2';
  const formula = '=AND(ISNUMBER(' + d + '),N(' + t + ')>0,ABS(' + d + ')>=' + t + ')';
  const rule = range => SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied(formula)
    .setBackground('#fde4e1').setFontColor('#a8231a')
    .setRanges([range])
    .build();
  const mine = [COL.diff, COL.diffPcs, COL.diffCost];
  const keep = log.getConditionalFormatRules()
    .filter(r => !r.getRanges().some(g => mine.indexOf(g.getColumn()) !== -1));
  keep.push(rule(log.getRange(2, COL.diff, rows, 2)), rule(log.getRange(2, COL.diffCost, rows, 1)));
  log.setConditionalFormatRules(keep);
  PropertiesService.getScriptProperties().setProperty('rules', RULES_VERSION);
}

function ensureRules_() {
  try {
    if (PropertiesService.getScriptProperties().getProperty('rules') === RULES_VERSION) return;
    applyRules_(logSheet_());
  } catch (e) {
    console.error(e); // formatting must never stop the page from opening
  }
}

/**
 * Runs when the page opens. A sheet made with older code is brought up to date once
 * (Products tab, Log columns, page title), then the red rule is kept up to date.
 */
function prepare_() {
  try {
    const props = PropertiesService.getScriptProperties();
    if (props.getProperty('schema') !== SCHEMA) {
      const lock = LockService.getScriptLock();
      if (lock.tryLock(15000)) {
        try {
          if (props.getProperty('schema') !== SCHEMA) {
            setup();
            renameOldTitle_();
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

/** "Wagyu Stock Check" was the default title; with more products it becomes "Meat Stock Check". */
function renameOldTitle_() {
  const set = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SETTINGS);
  if (!set) return;
  const at = settingRows_(set)['page title'];
  if (at && String(set.getRange(at, 2).getValue()).trim() === OLD_TITLE) set.getRange(at, 2).setValue(TITLE);
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
  const products = products_().filter(p => p.status !== 'Off')
    .map(p => ({ name: p.name, ko: p.ko, unit: p.unit, status: p.status, note: p.note }));
  t.configJson = JSON.stringify({ products: products }).replace(/</g, '\\u003c');
  return t.evaluate()
    .setTitle(cfg.title)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, viewport-fit=cover');
}

/**
 * Save one round of counts.
 * payload = { by: 'Sam', note: 'morning count',
 *             items: [{ product: 'Wagyu', count: 19, system: 2.764 }, { product: 'Pork belly', count: 12, system: 12 }] }
 */
function submitCount(payload) {
  if (!payload || typeof payload !== 'object') throw new Error('Nothing to save.');
  const by = cleanText_(payload.by, 40);
  if (!by) throw new Error('Please enter your name.');
  const note = cleanText_(payload.note, 200);
  // a page still open from the Wagyu-only version sends { count, systemKg }
  const items = Array.isArray(payload.items) ? payload.items
    : (payload.count !== undefined ? [{ product: WAGYU, count: payload.count, system: payload.systemKg }] : []);
  if (!items.length) throw new Error('Enter at least one product.');
  if (items.length > 50) throw new Error('Too many products in one save.');

  const catalog = {};
  products_().forEach(p => { if (p.status !== 'Off') catalog[p.name] = p; });
  const seen = {};
  const clean = items.map(it => {
    const name = String((it && it.product) || '');
    const p = catalog[name];
    if (!p) throw new Error('"' + name + '" is not on the product list. Refresh the page.');
    if (seen[name]) throw new Error('"' + name + '" is entered twice.');
    seen[name] = true;
    const count = wholeNumber_(it.count, false);
    if (count === null) throw new Error('Check the count for ' + name + '. Use a whole number from 0 to 9999.');
    const sys = p.unit === 'pcs' ? wholeNumber_(it.system, true) : kgNumber_(it.system);
    if (sys === null) throw new Error('Check the system stock for ' + name + (p.unit === 'pcs' ? '. Use whole pieces.' : '. Use kg, like 2.764.'));
    return { p: p, count: count, sys: sys };
  });

  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const log = logSheet_();
    const now = new Date();
    const prev = lastCounts_(log, clean.map(c => c.p.name));
    const start = log.getLastRow() + 1;
    ensureRows_(log, start + clean.length - 1);

    const batch = Utilities.formatDate(now, TZ, 'yyMMdd-HHmmss') + '-' + Utilities.getUuid().slice(0, 4);
    const values = clean.map((c, i) => rowValues_(start + i, batch + '.' + (i + 1), now, by, note, c, prev[c.p.name]));
    log.getRange(start, 1, values.length, LOG_HEADERS.length).setValues(values);
    formatRows_(log, start, clean.map(c => c.p.unit));
    SpreadsheetApp.flush();

    return {
      ok: true,
      at: Utilities.formatDate(now, TZ, SHOW_FMT),
      items: clean.map(c => ({ product: c.p.name, count: c.count, system: c.sys, unit: c.p.unit }))
    };
  } finally {
    lock.releaseLock();
  }
}

/** One Log row. The sheet formulas turn the count into the system unit and compare. */
function rowValues_(r, id, now, by, note, c, prev) {
  const a = col => colLetter_(col) + r;
  const E = a(COL.count), F = a(COL.avgKg), G = a(COL.counted), H = a(COL.sys), I = a(COL.diff),
    K = a(COL.price), M = a(COL.prevCount), Q = a(COL.unit);
  return [
    id,
    now,
    by,
    c.p.name,
    c.count,
    c.p.unit === 'kg' && c.p.avgKg > 0 ? c.p.avgKg : '',
    '=IF(OR(' + E + '="",' + Q + '=""),"",IF(' + Q + '="pcs",' + E + ',IF(' + E + '=0,0,IF(N(' + F + ')>0,ROUND(' + E + '*' + F + ',3),""))))',
    c.sys,
    '=IF(OR(' + G + '="",' + H + '=""),"",ROUND(' + G + '-' + H + ',3))',
    '=IF(' + I + '="","",IF(' + Q + '="pcs",' + I + ',IF(N(' + F + ')>0,ROUND(' + I + '/' + F + ',1),"")))',
    c.p.price,
    '=IF(OR(' + I + '="",' + K + '=""),"",ROUND(' + I + '*' + K + ',2))',
    prev ? prev.count : '',
    prev ? prev.at : '',
    '=IF(' + M + '="","",' + E + '-' + M + ')',
    note,
    c.p.unit,
    c.p.tol
  ];
}

/* ---------------------------------------------------------------- manager */

function getRecent(pin, limit) {
  checkPin_(pin);
  const products = products_();
  const list = products.map(p => ({ name: p.name, ko: p.ko, unit: p.unit, status: p.status, note: p.note }));
  const log = logSheet_();
  const last = log.getLastRow();
  if (last < 2) return { rows: [], total: 0, products: list };

  const tolOf = {};
  products.forEach(p => { tolOf[p.name] = p.tol; });
  const span = Math.min(last - 1, 3000);
  const vals = log.getRange(last - span + 1, 1, span, LOG_HEADERS.length).getValues();
  const max = Math.max(10, Math.min(Number(limit) || 200, 600));

  const rows = [];
  vals.forEach((r, i) => {
    if (r[COL.id - 1] === '' || String(r[COL.product - 1]) === '') return;
    const product = String(r[COL.product - 1]);
    const unit = String(r[COL.unit - 1]).trim() === 'pcs' ? 'pcs' : 'kg';
    const tol = num_(r[COL.tol - 1]);
    rows.push({
      idx: i,
      id: String(r[COL.id - 1]),
      atMs: r[COL.at - 1] instanceof Date ? r[COL.at - 1].getTime() : 0,
      at: fmt_(r[COL.at - 1]),
      by: String(r[COL.by - 1]),
      product: product,
      count: num_(r[COL.count - 1]),
      avgKg: num_(r[COL.avgKg - 1]),
      system: num_(r[COL.sys - 1]),
      unit: unit,
      tol: tol !== null && tol > 0 ? tol : (tolOf[product] || (unit === 'pcs' ? 1 : 0.1)),
      price: num_(r[COL.price - 1]),
      prevCount: num_(r[COL.prevCount - 1]),
      prevAt: fmt_(r[COL.prevAt - 1]),
      note: String(r[COL.note - 1])
    });
  });
  rows.sort((a, b) => (b.atMs - a.atMs) || (b.idx - a.idx));
  return { rows: rows.slice(0, max), total: rows.length, products: list };
}

/**
 * Manager: delete every count in the Log tab and start again.
 * The header row, Products, Settings and the PIN stay. The page asks "Are you sure?" first.
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
  if (log && log.getLastRow() > 0 && !isFirstLog_(log) && !isWagyuOnlyLog_(log)) return log;
  setup();
  return ss.getSheetByName(LOG);
}

/** Settings tab as { pin, title }. */
function config_() {
  const s = {};
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SETTINGS);
  if (sh && sh.getLastRow() > 0) {
    sh.getRange(1, 1, sh.getLastRow(), 2).getValues().forEach(r => {
      const k = String(r[0]).trim().toLowerCase();
      if (k && !(k in s)) s[k] = r[1];
    });
  }
  const title = String(s['page title'] == null ? '' : s['page title']).trim();
  return {
    pin: String(s['manager pin'] == null ? '' : s['manager pin']).trim(),
    title: title || TITLE
  };
}

/** Products tab as a list, in sheet order. Falls back to the built-in list if the tab is missing. */
function products_() {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(PRODUCTS);
  const list = productList_(sh);
  if (list.length) return list;
  return productList_(null);
}

function productList_(sh) {
  const raw = sh && sh.getLastRow() > 1
    ? sh.getRange(2, 1, sh.getLastRow() - 1, PRODUCT_HEADERS.length).getValues()
    : (sh ? [] : DEFAULT_PRODUCTS);
  const out = [];
  const seen = {};
  raw.forEach(r => {
    const name = String(r[P.name] == null ? '' : r[P.name]).trim();
    if (!name || seen[name]) return;
    seen[name] = true;
    const unit = /^(pcs?|pieces?|ea|each|개)$/i.test(String(r[P.unit]).trim()) ? 'pcs' : 'kg';
    const status = /^(off|no|n|숨김|끔)$/i.test(String(r[P.status]).trim()) ? 'Off'
      : /check|later|확인/i.test(String(r[P.status])) ? 'Check later' : 'Active';
    const avg = num_(r[P.avg]);
    const tol = num_(r[P.tol]);
    const price = num_(r[P.price]);
    out.push({
      name: name,
      ko: String(r[P.ko] == null ? '' : r[P.ko]).trim(),
      unit: unit,
      avgKg: avg !== null && avg > 0 && avg < 50 ? avg : null,
      tol: tol !== null && tol > 0 ? tol : (unit === 'pcs' ? 1 : 0.1),
      price: price !== null && price > 0 ? price : '',
      status: status,
      note: String(r[P.note] == null ? '' : r[P.note]).trim()
    });
  });
  return out;
}

/** Latest count per product, by time, so sorting the sheet doesn't break it. */
function lastCounts_(log, names) {
  const out = {};
  const last = log.getLastRow();
  if (last < 2) return out;
  const want = {};
  names.forEach(n => { want[n] = true; });
  const span = Math.min(last - 1, 5000);
  const vals = log.getRange(last - span + 1, COL.at, span, COL.count - COL.at + 1).getValues();
  vals.forEach(r => {
    const at = r[0];
    const name = String(r[COL.product - COL.at]);
    const n = r[COL.count - COL.at];
    if (!want[name] || !(at instanceof Date) || n === '') return;
    if (!out[name] || at.getTime() >= out[name].at.getTime()) out[name] = { at: at, count: n };
  });
  return out;
}

/** Number formats for new rows; numbers show their unit ("2.764 kg", "12 pcs"). */
function formatRows_(log, start, units) {
  if (!units.length) return;
  const n = units.length;
  log.getRange(start, COL.at, n, 1).setNumberFormat(DATE_FMT);
  log.getRange(start, COL.count, n, COL.change - COL.count + 1).setNumberFormats(units.map(u => {
    const kg = u !== 'pcs';
    const q = kg ? '0.000" kg"' : '0" pcs"';
    return ['0', '0.0000', q, q,
      kg ? '+0.000" kg";-0.000" kg";0.000" kg"' : '+0" pcs";-0" pcs";0" pcs"',
      kg ? '+0.0;-0.0;0.0' : '+0;-0;0',
      '$0.00', '+$0.00;-$0.00;$0.00', '0', DATE_FMT, '+0;-0;0'];
  }));
  log.getRange(start, COL.tol, n, 1).setNumberFormats(units.map(u => [u !== 'pcs' ? '0.000" kg"' : '0" pcs"']));
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

/** Whole number 0–9999 (or −9999–9999 when a system shows negative stock), else null. */
function wholeNumber_(v, allowNegative) {
  if (v === '' || v === null || v === undefined) return null;
  const t = typeof v === 'number' ? v : String(v).trim();
  if (t === '') return null;
  const n = Number(t);
  return Number.isInteger(n) && n <= 9999 && n >= (allowNegative ? -9999 : 0) ? n : null;
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
