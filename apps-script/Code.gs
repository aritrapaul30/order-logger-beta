/**
 * Digital Den Order Logger (beta): Google Sheets connector for several sheets.
 *
 * Fill in ACCESS_KEY and the spreadsheet IDs in SHEETS, then
 * Deploy > New deployment > Web app (Execute as: Me, Who has access: Anyone).
 * Every request must carry the key; without it nothing is read or written.
 */
const ACCESS_KEY = 'PUT-A-LONG-RANDOM-SECRET-HERE';
const SHEETS = {
  rahul:   { id: 'RAHUL-SPREADSHEET-ID',   tab: '' },            // '' = first tab
  sourav:  { id: 'SOURAV-SPREADSHEET-ID',  tab: '' },
  clients: { id: 'PAYMENTS-SPREADSHEET-ID', tab: '' },
  history: { id: 'PAYMENTS-SPREADSHEET-ID', tab: 'Settled history', create: true,
             header: ['Settled on', 'Client', 'Order Date', 'Transaction Amount', 'Given', 'Given Date / Method'] }
};
const MIN_COLS = { rahul: 30, sourav: 30, clients: 8, history: 6 };

function doPost(e) { return handle_(JSON.parse((e.postData && e.postData.contents) || '{}')); }
function doGet(e) { return handle_((e && e.parameter) || {}); }

function handle_(req) {
  try {
    if (!ACCESS_KEY || ACCESS_KEY.indexOf('PUT-') === 0 || req.key !== ACCESS_KEY) return out_({ ok: false, error: 'Wrong or missing access key.' });
    if (req.action === 'read') return out_({ ok: true, data: readAll_() });
    if (req.action === 'write') {
      const lock = LockService.getScriptLock();
      lock.waitLock(25000);
      try {
        const err = write_(req.changes || []);
        if (err) return out_({ ok: false, conflict: true, error: err });
        SpreadsheetApp.flush();
        return out_({ ok: true, data: readAll_() });
      } finally { lock.releaseLock(); }
    }
    return out_({ ok: false, error: 'Unknown action.' });
  } catch (err) {
    return out_({ ok: false, error: String((err && err.message) || err) });
  }
}

function out_(obj) { return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON); }

const open_ = {};
function sheet_(key, create) {
  const c = SHEETS[key];
  if (!c) throw new Error('Unknown sheet ' + key);
  const ss = open_[c.id] = open_[c.id] || SpreadsheetApp.openById(c.id);
  let sh = c.tab ? ss.getSheetByName(c.tab) : ss.getSheets()[0];
  if (!sh && c.create && create) {
    sh = ss.insertSheet(c.tab);
    if (c.header) sh.getRange(1, 1, 1, c.header.length).setValues([c.header]).setFontWeight('bold');
  }
  return { ss: ss, sh: sh };
}

function cellOut_(v, tz) {
  if (v instanceof Date) return { d: Utilities.formatDate(v, tz, 'yyyy-MM-dd') };
  return v;
}

function readAll_() {
  const data = {};
  Object.keys(SHEETS).forEach(function (k) {
    const o = sheet_(k, false);
    if (!o.sh) { data[k] = { values: [], name: '' }; return; }
    const tz = o.ss.getSpreadsheetTimeZone();
    const rows = Math.max(o.sh.getLastRow(), 2), cols = Math.max(o.sh.getLastColumn(), MIN_COLS[k] || 1);
    data[k] = {
      name: o.ss.getName() + ' / ' + o.sh.getName(),
      values: o.sh.getRange(1, 1, rows, cols).getValues().map(function (r) { return r.map(function (v) { return cellOut_(v, tz); }); })
    };
  });
  return data;
}

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
function write_(changes) {
  // 1. conflict check: every plain cell we change must still hold what the page saw
  for (var i = 0; i < changes.length; i++) {
    const c = changes[i];
    if (c.op || c.f || c.was === undefined) continue;
    const o = sheet_(c.sheet, false);
    if (!o.sh) continue;
    const now = cellOut_(o.sh.getRange(c.row, c.col).getValue(), o.ss.getSpreadsheetTimeZone());
    if (!same_(now, c.was)) return 'The sheet was changed from another device while you were filling this in (' + c.sheet + ' row ' + c.row + '). Nothing was saved. Please try again.';
  }
  // 2. structure: insert columns (right to left so positions stay valid)
  changes.filter(function (c) { return c.op === 'insertCols'; }).forEach(function (c) {
    const sh = sheet_(c.sheet, true).sh;
    c.cols.slice().sort(function (a, b) { return b - a; }).forEach(function (col) { sh.insertColumnBefore(col); });
  });
  // 3. formats: explicit copies, then new rows copy the row above
  changes.filter(function (c) { return c.op === 'copyFormat'; }).forEach(function (c) {
    const sh = sheet_(c.sheet, true).sh;
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
    const g = groups[k], w = g.max - g.min + 1, sh = sheet_(g.sheet, true).sh;
    sh.getRange(g.styleRow, g.min, 1, w).copyTo(sh.getRange(g.row, g.min, 1, w), SpreadsheetApp.CopyPasteType.PASTE_FORMAT, false);
  });
  // 4. values
  changes.forEach(function (c) {
    if (c.op) return;
    const r = sheet_(c.sheet, true).sh.getRange(c.row, c.col);
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
