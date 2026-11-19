/**
 * Digital Den Order Logger: Google Sheets connector.
 *
 * Drive layout it expects:
 *   Digital Den/                 (ROOT_FOLDER)
 *     Payment Details            (one file for all years; name contains "Payment")
 *     2026/  Rahul Yadav 2026, Sourav Mondal 2026   (names contain the dealer keyword)
 *     2027/  ...                 (created by the app's "Start new year" button)
 *
 * Set ACCESS_KEY, then Deploy > New deployment > Web app (Execute as: Me, Who has access: Anyone).
 * Every request must carry the key; without it nothing is read or written.
 */
const ACCESS_KEY = 'PUT-A-LONG-RANDOM-SECRET-HERE';
const ROOT_FOLDER = 'Digital Den';
const DEALER_KEYWORDS = { rahul: ['rahul'], sourav: ['sourav', 'sourab'] };
const CLIENTS_KEYWORD = 'payment';
const HISTORY_TAB = 'Settled history';
const HISTORY_HEADER = ['Settled on', 'Client', 'Order Date', 'Transaction Amount', 'Given', 'Given Date / Method'];
const MIN_COLS = { rahul: 30, sourav: 30, clients: 8, history: 6 };

function doPost(e) { return handle_(JSON.parse((e.postData && e.postData.contents) || '{}')); }
function doGet(e) { return handle_((e && e.parameter) || {}); }

function handle_(req) {
  try {
    if (!ACCESS_KEY || ACCESS_KEY.indexOf('PUT-') === 0 || req.key !== ACCESS_KEY) return out_({ ok: false, error: 'Wrong or missing access key.' });
    const ctx = context_(req.year);
    if (req.action === 'read') return out_({ ok: true, data: readAll_(ctx), years: ctx.years, year: ctx.year, missing: ctx.missing });
    const lock = LockService.getScriptLock();
    lock.waitLock(25000);
    try {
      if (req.action === 'write') {
        const err = write_(ctx, req.changes || []);
        if (err) return out_({ ok: false, conflict: true, error: err });
      } else if (req.action === 'newYear') {
        const made = newYear_(ctx, String(req.to || ''));
        const c2 = context_(made);
        SpreadsheetApp.flush();
        return out_({ ok: true, data: readAll_(c2), years: c2.years, year: c2.year, missing: c2.missing });
      } else return out_({ ok: false, error: 'Unknown action.' });
      SpreadsheetApp.flush();
      return out_({ ok: true, data: readAll_(ctx), years: ctx.years, year: ctx.year, missing: ctx.missing });
    } finally { lock.releaseLock(); }
  } catch (err) {
    return out_({ ok: false, error: String((err && err.message) || err) });
  }
}

function out_(obj) { return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON); }

/* ---------- find the files ---------- */
function root_() {
  const it = DriveApp.getFoldersByName(ROOT_FOLDER);
  if (!it.hasNext()) throw new Error('Folder "' + ROOT_FOLDER + '" was not found in your Google Drive.');
  return it.next();
}
function sheetsIn_(folder) {
  const out = [], it = folder.getFilesByType(MimeType.GOOGLE_SHEETS);
  while (it.hasNext()) { const f = it.next(); if (!f.isTrashed()) out.push(f); }
  return out;
}
function yearFolders_(root) {
  const out = {}, it = root.getFolders();
  while (it.hasNext()) { const f = it.next(); if (/^\d{4}$/.test(f.getName()) && !f.isTrashed()) out[f.getName()] = f; }
  return out;
}
function context_(year) {
  const root = root_(), folders = yearFolders_(root), years = Object.keys(folders).sort();
  if (!years.length) throw new Error('No year folder (like "2026") inside "' + ROOT_FOLDER + '".');
  const y = year && folders[year] ? String(year) : years[years.length - 1];
  const sheets = {}, missing = [];
  const clients = sheetsIn_(root).filter(function (f) { return f.getName().toLowerCase().indexOf(CLIENTS_KEYWORD) >= 0; })[0];
  if (clients) { sheets.clients = { id: clients.getId(), tab: '' }; sheets.history = { id: clients.getId(), tab: HISTORY_TAB, create: true, header: HISTORY_HEADER }; }
  else missing.push('Payment Details');
  const files = sheetsIn_(folders[y]);
  Object.keys(DEALER_KEYWORDS).forEach(function (k) {
    const f = files.filter(function (x) { const n = x.getName().toLowerCase(); return DEALER_KEYWORDS[k].some(function (w) { return n.indexOf(w) >= 0; }); })[0];
    if (f) sheets[k] = { id: f.getId(), tab: '' }; else missing.push(k + ' ' + y);
  });
  return { root: root, folders: folders, years: years, year: y, sheets: sheets, missing: missing, open: {} };
}
function sheet_(ctx, key, create) {
  const c = ctx.sheets[key];
  if (!c) throw new Error('No sheet found for ' + key + ' in ' + ctx.year + '.');
  const ss = ctx.open[c.id] = ctx.open[c.id] || SpreadsheetApp.openById(c.id);
  let sh = c.tab ? ss.getSheetByName(c.tab) : ss.getSheets()[0];
  if (!sh && c.create && create) {
    sh = ss.insertSheet(c.tab, ss.getSheets().length);
    if (c.header) sh.getRange(1, 1, 1, c.header.length).setValues([c.header]).setFontWeight('bold');
  }
  return { ss: ss, sh: sh };
}

/* ---------- read ---------- */
function cellOut_(v, tz) {
  if (v instanceof Date) return { d: Utilities.formatDate(v, tz, 'yyyy-MM-dd') };
  return v;
}
function readAll_(ctx) {
  const data = {};
  Object.keys(ctx.sheets).forEach(function (k) {
    const o = sheet_(ctx, k, false);
    if (!o.sh) { data[k] = { values: [], name: '' }; return; }
    const tz = o.ss.getSpreadsheetTimeZone();
    const rows = Math.max(o.sh.getLastRow(), 2), cols = Math.max(o.sh.getLastColumn(), MIN_COLS[k] || 1);
    data[k] = {
      name: o.ss.getName(),
      values: o.sh.getRange(1, 1, rows, cols).getValues().map(function (r) { return r.map(function (v) { return cellOut_(v, tz); }); })
    };
  });
  return data;
}

/* ---------- write ---------- */
function same_(a, b) {
  const norm = function (v) {
    if (v === null || v === undefined) return '';
    if (typeof v === 'object' && v.d) return 'D' + v.d;
    return String(v).trim().toUpperCase();
  };
  return norm(a) === norm(b);
}
function serial_(iso) {
  const p = iso.split('-').map(Number);
  return (Date.UTC(p[0], p[1] - 1, p[2]) - Date.UTC(1899, 11, 30)) / 86400000;
}
/* changes: cell writes {sheet,row,col,v|f|date,was,styleRow}
            {sheet,op:'insertCols',cols:[1-based columns, inserted before each]}
            {sheet,op:'copyFormat',from:[row,col,rows,cols],to:[row,col]} */
function write_(ctx, changes) {
  for (var i = 0; i < changes.length; i++) {
    const c = changes[i];
    if (c.op || c.f || c.was === undefined) continue;
    const o = sheet_(ctx, c.sheet, false);
    if (!o.sh) continue;
    const now = cellOut_(o.sh.getRange(c.row, c.col).getValue(), o.ss.getSpreadsheetTimeZone());
    if (!same_(now, c.was)) return 'The sheet was changed from another device while you were filling this in (' + c.sheet + ' row ' + c.row + '). Nothing was saved. Please try again.';
  }
  changes.filter(function (c) { return c.op === 'insertCols'; }).forEach(function (c) {
    const sh = sheet_(ctx, c.sheet, true).sh;
    c.cols.slice().sort(function (a, b) { return b - a; }).forEach(function (col) { sh.insertColumnBefore(col); });
  });
  changes.filter(function (c) { return c.op === 'copyFormat'; }).forEach(function (c) {
    const sh = sheet_(ctx, c.sheet, true).sh;
    sh.getRange(c.from[0], c.from[1], c.from[2], c.from[3]).copyTo(sh.getRange(c.to[0], c.to[1], c.from[2], c.from[3]), SpreadsheetApp.CopyPasteType.PASTE_FORMAT, false);
  });
  const groups = {};
  changes.forEach(function (c) {
    if (c.op || !c.styleRow) return;
    const k = c.sheet + ':' + c.row + ':' + c.styleRow;
    const g = groups[k] = groups[k] || { sheet: c.sheet, row: c.row, styleRow: c.styleRow, min: c.col, max: c.col };
    g.min = Math.min(g.min, c.col); g.max = Math.max(g.max, c.col);
  });
  Object.keys(groups).forEach(function (k) {
    const g = groups[k], w = g.max - g.min + 1, sh = sheet_(ctx, g.sheet, true).sh;
    sh.getRange(g.styleRow, g.min, 1, w).copyTo(sh.getRange(g.row, g.min, 1, w), SpreadsheetApp.CopyPasteType.PASTE_FORMAT, false);
  });
  changes.forEach(function (c) {
    if (c.op) return;
    const r = sheet_(ctx, c.sheet, true).sh.getRange(c.row, c.col);
    if (c.f) { r.setFormula(c.f); return; }
    if (c.date) {
      if (!/[dy]/i.test(r.getNumberFormat())) r.setNumberFormat('dd/mm/yyyy');
      r.setValue(serial_(c.date));
      return;
    }
    r.setValue(c.v === undefined ? '' : c.v);
  });
  return null;
}

/* ---------- start a new year ----------
   Copies every dealer sheet of the latest year into a new year folder, keeps the layout and the
   row-2 totals, clears the orders and payments, and puts last year's LEFT (AA2) into the GOT column
   of row 2 as a negative "Previous Year Due", the same way the yearly sheets were started by hand. */
function newYear_(ctx, to) {
  if (!/^\d{4}$/.test(to)) throw new Error('Give the new year as 4 digits.');
  if (ctx.folders[to]) throw new Error('The ' + to + ' folder already exists.');
  const from = ctx.years[ctx.years.length - 1];
  if (+to <= +from) throw new Error('The new year must be after ' + from + '.');
  const src = context_(from);
  const folder = ctx.root.createFolder(to);
  Object.keys(DEALER_KEYWORDS).forEach(function (k) {
    if (!src.sheets[k]) return;
    const file = DriveApp.getFileById(src.sheets[k].id);
    let name = file.getName().indexOf(from) >= 0 ? file.getName().split(from).join(to) : file.getName() + ' ' + to;
    const copy = file.makeCopy(name, folder);
    const oldSh = SpreadsheetApp.openById(src.sheets[k].id).getSheets()[0];
    const left = Number(oldSh.getRange('AA2').getValue()) || 0;
    const sh = SpreadsheetApp.openById(copy.getId()).getSheets()[0];
    const last = sh.getLastRow(), cols = Math.max(sh.getLastColumn(), 30);
    if (last >= 3) sh.getRange(3, 1, last - 2, cols).clearContent();
    const row2 = sh.getRange(2, 1, 1, cols).getValues()[0];
    if (!row2.some(function (v) { return String(v).toLowerCase().indexOf('previous year due') >= 0; })) sh.getRange('W2').setValue('Previous Year Due');
    sh.getRange('Y2').setValue(-left);
  });
  return to;
}
