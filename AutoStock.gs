/***************************************************************
 * AutoStock v2.0 - warehouse stock on Google Sheets + Apps Script
 *
 * Sheets (names must match exactly, or change them in CONFIG):
 *   - Master    : item list + stock. Item Name/Size/Type/Brand/Unit
 *                 MAY be edited here. Code & Qty are protected (script).
 *   - Stock In  : the admin records incoming items
 *   - Stock Out : the admin records outgoing items
 *   - Archive   : items whose stock is 0 (filled automatically)
 *
 * HOW IT WORKS:
 *   Master is never "added to" or "subtracted from". On every change,
 *   stock is RECALCULATED from scratch:  Stock = total In - total Out.
 *   So Ctrl+Z, editing a qty, deleting a row are all safe, any number of times.
 *
 * INSTALL (once per file):
 *   1. From the Google Sheet: Extensions -> Apps Script
 *   2. Delete everything in the editor, paste THIS WHOLE file, Save (Ctrl+S)
 *   3. Pick the "setup" function in the dropdown -> Run -> allow access
 *   4. Back to the sheet, refresh. Done.
 *   (setup also installs the triggers. Safe to run again.)
 ***************************************************************/

/* ============================================================
 * VERSION & CHANGELOG
 * ------------------------------------------------------------
 * v2.0 (26/10/2026) - REWRITTEN. Master is recalculated from the logs
 *                     on every change (not added to / subtracted from).
 *                     -> Ctrl+Z, editing qty, deleting a row = safe.
 *                     -> Edit name/spec/type/brand/unit directly in
 *                        Master; the change is copied to every log.
 *                     -> A typo in a log can be fixed in the log;
 *                        "orphan" codes (no transactions) clear themselves.
 *                     -> Not enough stock = red warning (the entry is
 *                        not rejected, so Ctrl+Z always works).
 *                     -> Changing the unit on any sheet updates it everywhere.
 *                     One menu button left: Recalculate Stock.
 * ============================================================ */
const VERSION = '2.0';

const CONFIG = {
  SHEET_MASTER: 'Master',
  SHEET_IN: 'Stock In',
  SHEET_OUT: 'Stock Out',
  SHEET_ARCHIVE: 'Archive',
  CODE_DIGITS: 4,          // 0001, 0002, ... (past 9999 it becomes 5 digits)
  COLOR_ERROR: '#f4c7c3',  // light red = something is wrong
  COLOR_OK: '#d9ead3',     // light green = counted
  // Units offered in the dropdown (editable; other units can still be typed)
  UNIT_LIST: ['pcs', 'unit', 'rod', 'roll', 'set', 'box', 'meter', 'kg', 'liter', 'sheet'],
};

const HEADER_MASTER  = ['Code', 'Item Name', 'Size/Spec', 'Type', 'Brand', 'Qty', 'Unit'];
const HEADER_IN      = ['Date', 'Item Code', 'Item Name', 'Size/Spec', 'Type', 'Brand', 'Qty In', 'Unit', 'Project', 'Notes', 'Status'];
const HEADER_OUT     = ['Date', 'Item Code', 'Item Name', 'Size/Spec', 'Type', 'Brand', 'Qty Out', 'Unit', 'Project', 'Notes', 'Status'];
const HEADER_ARCHIVE = ['Code', 'Item Name', 'Size/Spec', 'Type', 'Brand', 'Qty', 'Unit', 'Date Sold Out'];

// Column index (from 0) - Stock In & Stock Out use the SAME order
const L = { DATE: 0, CODE: 1, NAME: 2, SPEC: 3, TYPE: 4, BRAND: 5, QTY: 6, UNIT: 7, PROJECT: 8, NOTES: 9, STATUS: 10 };
const LOG_WIDTH = 11;
const MASTER_WIDTH = 7;
const ARCHIVE_WIDTH = 8;

/* ============================================================
 * SETUP - run once (Run -> setup). Safe to run again.
 * ============================================================ */
function setup() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const master  = prepareSheet_(ss, CONFIG.SHEET_MASTER, HEADER_MASTER);
  const stockIn = prepareSheet_(ss, CONFIG.SHEET_IN, HEADER_IN);
  const stockOut = prepareSheet_(ss, CONFIG.SHEET_OUT, HEADER_OUT);
  const archive = prepareSheet_(ss, CONFIG.SHEET_ARCHIVE, HEADER_ARCHIVE);

  // Code columns = TEXT, so "0001" does not turn into 1
  master.getRange('A2:A').setNumberFormat('@');
  archive.getRange('A2:A').setNumberFormat('@');
  stockIn.getRange('B2:B').setNumberFormat('@');
  stockOut.getRange('B2:B').setNumberFormat('@');
  stockIn.getRange('A2:A').setNumberFormat('dd/mm/yyyy');
  stockOut.getRange('A2:A').setNumberFormat('dd/mm/yyyy');
  archive.getRange('H2:H').setNumberFormat('dd/mm/yyyy');

  // Code dropdown in both logs (taken from Master; typing by hand also works)
  const codeRule = SpreadsheetApp.newDataValidation()
    .requireValueInRange(master.getRange('A2:A'), true)
    .setAllowInvalid(true)
    .build();
  stockIn.getRange('B2:B').setDataValidation(codeRule);
  stockOut.getRange('B2:B').setDataValidation(codeRule);

  // Unit dropdown in both logs
  const unitRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(CONFIG.UNIT_LIST, true)
    .setAllowInvalid(true)
    .build();
  stockIn.getRange('H2:H').setDataValidation(unitRule);
  stockOut.getRange('H2:H').setDataValidation(unitRule);

  // Protect the script-owned columns (a warning appears if someone types there).
  // Deliberately a "warning", not a hard lock: so the Recalculate button still works
  // when an admin clicks it. Even if it is overwritten, the script puts it back.
  // The Status column in the logs is NOT protected, so Delete row & pasting full rows work.
  [master, stockIn, stockOut, archive].forEach(function (sh) {
    sh.getProtections(SpreadsheetApp.ProtectionType.RANGE).forEach(function (p) {
      if (String(p.getDescription()).indexOf('autostock:') === 0) p.remove();
    });
  });
  protect_(master, 'A2:A', 'Code is filled by the script');
  protect_(master, 'F2:F', 'Qty is calculated from the logs');
  protect_(archive, 'A2:A', 'Code is filled by the script');
  protect_(archive, 'F2:F', 'Qty is calculated from the logs');
  protect_(archive, 'H2:H', 'Date Sold Out is filled by the script');

  // Order the tabs & remove the default empty tab ("Sheet1")
  [master, stockIn, stockOut, archive].forEach(function (sh, i) {
    ss.setActiveSheet(sh);
    ss.moveActiveSheet(i + 1);
  });
  ss.getSheets().forEach(function (sh) {
    if (MANAGED_SHEETS.indexOf(sh.getName()) >= 0) return;
    if (/^(sheet|lembar)\s?\d+$/i.test(sh.getName()) && sh.getLastRow() === 0 && sh.getLastColumn() === 0) {
      ss.deleteSheet(sh);
    }
  });
  ss.setActiveSheet(stockIn);

  // Triggers: remove every old trigger of this project, install fresh ones.
  //  - onEditHandler   : typing / paste / Ctrl+Z
  //  - onChangeHandler : deleting/inserting rows (Delete row is NOT an "edit" to Google)
  ScriptApp.getProjectTriggers().forEach(function (t) { ScriptApp.deleteTrigger(t); });
  ScriptApp.newTrigger('onEditHandler').forSpreadsheet(ss).onEdit().create();
  ScriptApp.newTrigger('onChangeHandler').forSpreadsheet(ss).onChange().create();

  recalculate_(ss, null, true);
  ss.toast('Setup complete. Stock has been calculated from the logs.', 'AutoStock v' + VERSION, 6);
}

function prepareSheet_(ss, name, header) {
  let sh = ss.getSheetByName(name);
  if (!sh) sh = ss.insertSheet(name);
  sh.getRange(1, 1, 1, header.length).setValues([header])
    .setFontWeight('bold').setBackground('#eeeeee');
  sh.setFrozenRows(1);
  return sh;
}

function protect_(sheet, a1, description) {
  sheet.getRange(a1).protect()
    .setDescription('autostock:' + description)
    .setWarningOnly(true);
}

/* ============================================================
 * TRIGGERS - called automatically (do not run by hand)
 * ============================================================ */
const MANAGED_SHEETS = [CONFIG.SHEET_MASTER, CONFIG.SHEET_IN, CONFIG.SHEET_OUT, CONFIG.SHEET_ARCHIVE];

// On every typing / paste / Ctrl+Z
function onEditHandler(e) {
  if (!e) return;
  const ss = e.source || SpreadsheetApp.getActiveSpreadsheet();
  if (!e.range) { recalculate_(ss, null); return; } // safety: event without cell info
  const name = e.range.getSheet().getName();
  if (MANAGED_SHEETS.indexOf(name) < 0) return;

  // If an identity column (Name..Brand) is edited in a log, remember the rows:
  // those rows are "re-matched" (fixing a typo / moving to another item).
  let edit = null;
  if (name === CONFIG.SHEET_IN || name === CONFIG.SHEET_OUT) {
    const c1 = e.range.getColumn(), c2 = c1 + e.range.getNumColumns() - 1;
    // Only Project/Notes edited -> stock does not change, nothing to do
    if (c1 >= L.PROJECT + 1 && c2 <= L.NOTES + 1) return;
    if (c1 <= L.BRAND + 1 && c2 >= L.NAME + 1) {
      const rows = [];
      const r0 = e.range.getRow(), r1 = r0 + e.range.getNumRows() - 1;
      for (let r = Math.max(r0, 2); r <= r1; r++) rows.push(r - 2); // row 2 = index 0
      edit = { log: name === CONFIG.SHEET_IN ? 'in' : 'out', rows: rows };
    }
    // If the Unit column is edited in a log, remember the rows: that new unit
    // is used for the item in Master, Archive, and every log row.
    if (c1 <= L.UNIT + 1 && c2 >= L.UNIT + 1) {
      const unitRows = [];
      const r0 = e.range.getRow(), r1 = r0 + e.range.getNumRows() - 1;
      for (let r = Math.max(r0, 2); r <= r1; r++) unitRows.push(r - 2);
      edit = edit || { log: name === CONFIG.SHEET_IN ? 'in' : 'out', rows: [] };
      edit.unit = unitRows;
    }
  }
  recalculate_(ss, edit);
}

// Structural changes: deleting/inserting rows, etc. Normal edits are handled by onEditHandler.
function onChangeHandler(e) {
  if (!e || e.changeType === 'EDIT' || e.changeType === 'FORMAT') return;
  recalculate_(e.source || SpreadsheetApp.getActiveSpreadsheet(), null);
}

/* ============================================================
 * MENU
 * ============================================================ */
function onOpen() {
  try {
    SpreadsheetApp.getUi().createMenu('🛠 Modification Tools')
      .addItem('🔄 Recalculate Stock', 'recalculateButton')
      .addToUi();
  } catch (err) {
    // Ignored (context without a UI)
  }
}

function recalculateButton() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  recalculate_(ss, null, true);
  ss.toast('Stock has been recalculated from the logs.', 'AutoStock', 4);
}

/* ============================================================
 * READ -> CALCULATE -> WRITE
 * ============================================================ */
// repaint = true: Status & Master colours are fully repainted (setup & Recalculate button).
// On normal typing only the changed cells are painted, so the script stays fast.
function recalculate_(ss, edit, repaint) {
  const lock = LockService.getDocumentLock();
  if (!lock.tryLock(30000)) return; // another run is busy; the next edit will recalculate
  try {
    const sh = {
      master: ss.getSheetByName(CONFIG.SHEET_MASTER),
      in: ss.getSheetByName(CONFIG.SHEET_IN),
      out: ss.getSheetByName(CONFIG.SHEET_OUT),
      archive: ss.getSheetByName(CONFIG.SHEET_ARCHIVE),
    };
    if (!sh.master || !sh.in || !sh.out || !sh.archive) {
      throw new Error('Sheets are incomplete. Run setup first.');
    }
    const data = {
      master: read_(sh.master, MASTER_WIDTH),
      archive: read_(sh.archive, ARCHIVE_WIDTH),
      in: read_(sh.in, LOG_WIDTH),
      out: read_(sh.out, LOG_WIDTH),
    };
    const result = computeStock_(data, edit, new Date());

    writeLog_(sh.in, data.in, result.in, repaint);
    writeLog_(sh.out, data.out, result.out, repaint);
    writeList_(sh.master, data.master, result.master, MASTER_WIDTH, repaint);
    writeList_(sh.archive, data.archive, result.archive, ARCHIVE_WIDTH, repaint);
  } finally {
    lock.releaseLock();
  }
}

function read_(sheet, width) {
  const last = sheet.getLastRow();
  if (last < 2) return [];
  return sheet.getRange(2, 1, last - 1, width).getValues();
}

// Write only the columns the script owns, and only the cells that changed.
// Qty, Project, Notes are NEVER written by the script (they belong to the admin).
// Before writing, the cells are read again: if the admin has just typed there,
// the script backs off (that typing triggers the next recalculation itself).
function writeLog_(sheet, before, after, repaint) {
  const n = after.rows.length;
  if (n === 0) return;
  const blocks = [
    { c0: L.DATE, c1: L.DATE },
    { c0: L.CODE, c1: L.CODE, text: true },
    { c0: L.NAME, c1: L.BRAND },
    { c0: L.UNIT, c1: L.UNIT },
    { c0: L.STATUS, c1: L.STATUS },
  ];
  blocks.forEach(function (b) {
    let i = 0;
    while (i < n) {
      if (sameRow_(before[i], after.rows[i], b.c0, b.c1)) { i++; continue; }
      let j = i;
      while (j + 1 < n && !sameRow_(before[j + 1], after.rows[j + 1], b.c0, b.c1)) j++;
      safeWrite_(sheet, i, j, b.c0, b.c1, before, after.rows, b.text);
      i = j + 1;
    }
  });

  // Status column colour: only rows whose status changed (or all of them on repaint)
  const colorOf = function (s) {
    return [s === '' ? null : (s === 'OK' ? CONFIG.COLOR_OK : CONFIG.COLOR_ERROR)];
  };
  const changed = function (i) {
    return repaint || !sameValue_(before[i][L.STATUS], after.status[i]);
  };
  let i = 0;
  while (i < n) {
    if (!changed(i)) { i++; continue; }
    let j = i;
    while (j + 1 < n && changed(j + 1)) j++;
    sheet.getRange(i + 2, L.STATUS + 1, j - i + 1, 1)
      .setBackgrounds(after.status.slice(i, j + 1).map(colorOf));
    i = j + 1;
  }
}

function sameRow_(a, b, c0, c1) {
  for (let c = c0; c <= c1; c++) if (!sameValue_(a[c], b[c])) return false;
  return true;
}

// Write rows i..j, columns c0..c1, only if the sheet still holds what was read earlier
function safeWrite_(sheet, i, j, c0, c1, before, rows, text) {
  const range = sheet.getRange(i + 2, c0 + 1, j - i + 1, c1 - c0 + 1);
  const now = range.getValues();
  for (let x = 0; x < now.length; x++) {
    for (let c = c0; c <= c1; c++) {
      if (!sameValue_(now[x][c - c0], before[i + x][c])) return; // the admin is typing here
    }
  }
  if (text) range.setNumberFormat('@');
  range.setValues(rows.slice(i, j + 1).map(function (r) { return r.slice(c0, c1 + 1); }));
}

// Master & Archive are rewritten in full (sorted by code) when something changed.
function writeList_(sheet, before, after, width, repaint) {
  const nBefore = before.length, nAfter = after.length;
  let same = nBefore === nAfter;
  for (let i = 0; same && i < nAfter; i++) {
    if (!sameRow_(before[i], after[i], 0, width - 1)) same = false;
  }
  if (same && !repaint) return; // nothing changed: no need to write or repaint
  if (!same) {
    // Check again: if the admin has just edited this sheet, the script backs off for now
    const now = read_(sheet, width);
    if (now.length !== nBefore) return;
    for (let i = 0; i < nBefore; i++) {
      if (!sameRow_(now[i], before[i], 0, width - 1)) return;
    }
    if (nAfter > 0) {
      sheet.getRange(2, 1, nAfter, 1).setNumberFormat('@');
      sheet.getRange(2, 1, nAfter, width).setValues(after);
    }
    if (nBefore > nAfter) sheet.getRange(nAfter + 2, 1, nBefore - nAfter, width).clearContent();
  }
  // Negative qty = red (not enough stock, check Stock Out)
  const nColor = Math.max(nBefore, nAfter);
  if (nColor > 0) {
    const colors = [];
    for (let i = 0; i < nColor; i++) {
      colors.push([i < nAfter && after[i][5] < 0 ? CONFIG.COLOR_ERROR : null]);
    }
    sheet.getRange(2, 6, nColor, 1).setBackgrounds(colors);
  }
}

function sameValue_(a, b) {
  if (a instanceof Date && b instanceof Date) return a.getTime() === b.getTime();
  if (a instanceof Date || b instanceof Date) return false;
  if (typeof a === 'number' && typeof b === 'number') return a === b;
  return String(a === null || a === undefined ? '' : a) === String(b === null || b === undefined ? '' : b);
}

/* ============================================================
 * THE BRAIN - pure function, no Google Sheet access.
 * (Kept separate so it can be tested outside Google.)
 *
 * data.master  : [[code, name, spec, type, brand, qty, unit]]
 * data.archive : [[code, name, spec, type, brand, qty, unit, dateSoldOut]]
 * data.in      : [[date, code, name, spec, type, brand, qty, unit, project, notes, status]]
 * data.out     : same as in
 * edit         : null, or { log: 'in'|'out', rows: [index...], unit: [index...] }
 *                = log rows whose identity / unit columns were just edited
 * ============================================================ */
function computeStock_(data, edit, now) {
  const stockIn = data.in.map(function (r) { return r.slice(); });
  const stockOut = data.out.map(function (r) { return r.slice(); });
  const log = { in: stockIn, out: stockOut };

  // 1. Item catalog (code -> identity) from Master + Archive.
  //    This is the "memory" of names/specs the admin may have edited in Master.
  const catalog = {};
  data.master.concat(data.archive).forEach(function (r) {
    const k = padCode_(r[0]);
    if (k === '' || catalog[k]) return;
    catalog[k] = { name: text_(r[1]), spec: text_(r[2]), type: text_(r[3]), brand: text_(r[4]), unit: text_(r[6]) };
  });

  // Tidy the codes in the logs ("1" -> "0001")
  [stockIn, stockOut].forEach(function (rows) {
    rows.forEach(function (r) { r[L.CODE] = padCode_(r[L.CODE]); });
  });

  const validQty = function (r) { const n = toNumber_(r[L.QTY]); return !isNaN(n) && n > 0; };
  const emptyQty = function (r) { return text_(r[L.QTY]) === ''; };
  const rowId = function (r) {
    return { name: text_(r[L.NAME]), spec: text_(r[L.SPEC]), type: text_(r[L.TYPE]), brand: text_(r[L.BRAND]), unit: text_(r[L.UNIT]) };
  };
  const findCode = function (id, except) {
    const key = identityKey_(id);
    for (const k in catalog) {
      if (k !== except && identityKey_(catalog[k]) === key) return k;
    }
    return '';
  };

  // 2. Log rows whose identity was just edited -> match them again
  if (edit && edit.rows) {
    const rows = log[edit.log];
    edit.rows.forEach(function (i) {
      const r = rows[i];
      if (!r) return;
      const k = r[L.CODE];
      const id = rowId(r);
      if (k === '' || id.name === '') return;            // no code / name deleted -> let it be filled again
      if (catalog[k] && identityKey_(catalog[k]) === identityKey_(id)) return; // unchanged
      const other = findCode(id, k);
      if (other) { r[L.CODE] = other; return; }          // it is actually another existing item
      if (edit.log === 'in' && usageCount_(log, k) === 1) {
        // Only this row uses the code -> treat it as a typo fix: rename it in the catalog
        const unit = id.unit || (catalog[k] ? catalog[k].unit : '');
        catalog[k] = { name: id.name, spec: id.spec, type: id.type, brand: id.brand, unit: unit };
        return;
      }
      r[L.CODE] = ''; // a different item -> in: gets a new code; out: a code must be picked
    });
  }

  // 3. Codes in the logs that are not in the catalog (e.g. a Master row was deleted) -> restore from the log
  [stockIn, stockOut].forEach(function (rows) {
    rows.forEach(function (r) {
      const k = r[L.CODE];
      if (k !== '' && !catalog[k] && validQty(r) && text_(r[L.NAME]) !== '') catalog[k] = rowId(r);
    });
  });

  // 4. Rows with a qty but no code -> find the same item, or (in) create a new code
  stockIn.forEach(function (r) {
    if (r[L.CODE] !== '' || !validQty(r) || text_(r[L.NAME]) === '') return;
    let k = findCode(rowId(r), '');
    if (!k) { k = newCode_(catalog, log); catalog[k] = rowId(r); }
    r[L.CODE] = k;
  });
  stockOut.forEach(function (r) {
    if (r[L.CODE] !== '' || !validQty(r) || text_(r[L.NAME]) === '') return;
    const k = findCode(rowId(r), '');
    if (k) r[L.CODE] = k;
  });

  // 5. Add up
  const totalIn = {}, totalOut = {}, lastDate = {}, used = {};
  const add = function (rows, total) {
    rows.forEach(function (r) {
      const k = r[L.CODE];
      if (k === '' || !catalog[k] || !validQty(r)) return;
      total[k] = (total[k] || 0) + toNumber_(r[L.QTY]);
      used[k] = true;
      const t = r[L.DATE] instanceof Date ? r[L.DATE] : (r[L.DATE] ? null : now);
      if (t && (!lastDate[k] || t > lastDate[k])) lastDate[k] = t;
    });
  };
  add(stockIn, totalIn);
  add(stockOut, totalOut);
  const stock = function (k) { return round_((totalIn[k] || 0) - (totalOut[k] || 0)); };

  // 6. "Orphan" codes (not a single transaction with a qty) are cleared from log rows without a qty.
  //    Example: Ctrl+Z on a new item's qty -> its code disappears too, leaves no trace,
  //    and the code number is reused for the next new item.
  //    Safe: codes of items that still have transactions (Master/Archive) are not touched.
  [stockIn, stockOut].forEach(function (rows) {
    rows.forEach(function (r) {
      if (r[L.CODE] !== '' && !used[r[L.CODE]] && emptyQty(r)) r[L.CODE] = '';
    });
  });

  // 7. Unit: empty in the catalog -> take it from the log
  [stockIn, stockOut].forEach(function (rows) {
    rows.forEach(function (r) {
      const k = r[L.CODE];
      if (k !== '' && catalog[k] && catalog[k].unit === '' && text_(r[L.UNIT]) !== '') catalog[k].unit = text_(r[L.UNIT]);
    });
  });

  // 7b. A unit just edited in a log -> becomes the item's unit in the catalog
  //     (then copied to Master, Archive, and every log row in steps 8-9)
  if (edit && edit.unit) {
    edit.unit.forEach(function (i) {
      const r = log[edit.log][i];
      if (!r) return;
      const k = r[L.CODE], unit = text_(r[L.UNIT]);
      if (k !== '' && catalog[k] && unit !== '') catalog[k].unit = unit;
    });
  }

  // 8. Match the identity in the logs to the catalog + automatic date + status
  const status = { in: [], out: [] };
  ['in', 'out'].forEach(function (kind) {
    log[kind].forEach(function (r) {
      const k = r[L.CODE];
      if (k !== '' && catalog[k]) {
        r[L.NAME] = catalog[k].name; r[L.SPEC] = catalog[k].spec;
        r[L.TYPE] = catalog[k].type; r[L.BRAND] = catalog[k].brand;
        if (catalog[k].unit !== '') r[L.UNIT] = catalog[k].unit; // same unit on every sheet
      }
      if (validQty(r) && text_(r[L.DATE]) === '') r[L.DATE] = now;
      const s = rowStatus_(kind, r, k, catalog, validQty(r), emptyQty(r), stock, totalIn, totalOut);
      r[L.STATUS] = s;
      status[kind].push(s);
    });
  });

  // 9. Build Master (stock != 0) and Archive (stock = 0), sorted by code
  const usedCodes = Object.keys(used).sort(sortCode_);
  const master = [], archive = [];
  usedCodes.forEach(function (k) {
    const b = catalog[k], q = stock(k);
    if (q === 0) archive.push([k, b.name, b.spec, b.type, b.brand, 0, b.unit, lastDate[k] || '']);
    else master.push([k, b.name, b.spec, b.type, b.brand, q, b.unit]);
  });

  return {
    in: { rows: stockIn, status: status.in },
    out: { rows: stockOut, status: status.out },
    master: master,
    archive: archive,
  };
}

function rowStatus_(kind, r, k, catalog, valid, empty, stock, totalIn, totalOut) {
  const filled = r.slice(0, L.STATUS).some(function (v) { return text_(v) !== ''; });
  if (!filled || empty) return '';
  if (!valid) return '⚠ Qty must be a number above 0';
  if (kind === 'in') {
    if (k === '') return '⚠ Fill in Item Name (or pick a Code)';
    if (!catalog[k]) return '⚠ Unknown code';
    return 'OK';
  }
  if (k === '') return text_(r[L.NAME]) === '' ? '⚠ Pick an Item Code' : '⚠ Item not found - pick a Code';
  if (!catalog[k]) return '⚠ Unknown code';
  if (stock(k) < 0) return '⚠ Not enough stock: in ' + round_(totalIn[k] || 0) + ', out ' + round_(totalOut[k] || 0);
  return 'OK';
}

// How many rows (in both logs) use this code
function usageCount_(log, k) {
  let n = 0;
  log.in.concat(log.out).forEach(function (r) { if (r[L.CODE] === k) n++; });
  return n;
}

// Next code = highest in the catalog & logs + 1
function newCode_(catalog, log) {
  let max = 0;
  const check = function (k) { const n = toNumber_(k); if (!isNaN(n) && n > max) max = n; };
  Object.keys(catalog).forEach(check);
  log.in.concat(log.out).forEach(function (r) { check(r[L.CODE]); });
  return padCode_(max + 1);
}

function sortCode_(a, b) {
  const na = toNumber_(a), nb = toNumber_(b);
  if (!isNaN(na) && !isNaN(nb)) return na - nb;
  return a < b ? -1 : a > b ? 1 : 0;
}

// Identity key: lowercase + ALL punctuation & spaces removed.
// "Elbow BS sch 40" = "elbow bs sch. 40". Different numbers stay different (sch 40 ≠ sch 45).
function identityKey_(id) {
  return [id.name, id.spec, id.type, id.brand].map(function (x) {
    return text_(x).toLowerCase().replace(/[^a-z0-9]/g, '');
  }).join('|');
}

function padCode_(x) {
  const s = text_(x);
  if (s === '') return '';
  const n = toNumber_(s);
  if (isNaN(n)) return s;
  let p = String(Math.floor(n));
  while (p.length < CONFIG.CODE_DIGITS) p = '0' + p;
  return p;
}

function text_(x) {
  if (x === null || x === undefined) return '';
  return String(x).trim();
}

function toNumber_(x) {
  if (x === null || x === undefined || x === '') return NaN;
  if (typeof x === 'number') return x;
  const n = Number(String(x).trim().replace(',', '.'));
  return isNaN(n) ? NaN : n;
}

// Avoid 0.30000000000000004
function round_(x) {
  return Math.round(x * 1e6) / 1e6;
}
