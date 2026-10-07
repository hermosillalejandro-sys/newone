/**
 * Stack Ranking – bonuses
 *
 *  - Checkbox in IZ:
 *      TRUE  -> JE:JG of that row are frozen as values.
 *      FALSE -> JE:JG get the base formula back (template JE3:JG3).
 *    Every row in the edit is processed (paste, fill down, several checkboxes
 *    toggled at once, multi-range selections). The action is idempotent: each
 *    row is set to whatever its checkbox says, so repeating it is harmless.
 *
 *  - validarBonosAvanzado(): writes OK / Mismatch in JM for every data row.
 *    It also runs by itself for the rows touched by an edit in IZ, JE:JG or JJ.
 */

const SHEET_NAME = "Stack Ranking";
const FIRST_ROW = 7;           // first data row (rows 5-6 are headers)
const TEMPLATE_ROW = 3;        // JE3:JG3 hold the base formulas
const COL_CHECK = 260;         // IZ
const COL_JE = 265;            // JE (Tech + Booster), JF, JG
const COL_COMMENT = 270;       // JJ (Comments)
const COL_RESULT = 273;        // JM

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("Bonos")
    .addItem("Validar bonos (JM)", "validarBonosAvanzado")
    .addItem("Sincronizar checkboxes IZ", "sincronizarCheckboxes")
    .addToUi();
}

// If this project already has another onEdit, delete this one and call
// onEditBonos(e) from the existing onEdit instead.
function onEdit(e) {
  onEditBonos(e);
}

function onEditBonos(e) {
  if (!e || !e.range) return;
  const sheet = e.range.getSheet();
  if (sheet.getName() !== SHEET_NAME) return;

  // e.range only covers the active range; include the whole selection so no
  // row of a multi-range edit is skipped.
  const ranges = [e.range];
  try {
    const list = sheet.getActiveRangeList();
    if (list) list.getRanges().forEach(r => ranges.push(r));
  } catch (err) {}

  const checkRows = rowsTouching(ranges, COL_CHECK, COL_CHECK);
  const validateRows = rowsTouching(ranges, COL_JE, COL_JE + 2)
    .concat(rowsTouching(ranges, COL_COMMENT, COL_COMMENT))
    .concat(checkRows);
  if (!validateRows.length) return;

  const lock = LockService.getDocumentLock();
  lock.waitLock(25000);
  try {
    if (checkRows.length) aplicarCheckboxes(sheet, uniqueSorted(checkRows));
    validarFilas(sheet, uniqueSorted(validateRows));
  } finally {
    lock.releaseLock();
  }
}

/** Safety net: applies IZ to every data row. */
function sincronizarCheckboxes() {
  const sheet = getSheet();
  const rows = allDataRows(sheet);
  if (!rows.length) return;
  aplicarCheckboxes(sheet, rows);
  validarFilas(sheet, rows);
  SpreadsheetApp.getActiveSpreadsheet().toast("IZ synced for " + rows.length + " rows.", "Bonus Audit");
}

function validarBonosAvanzado() {
  const sheet = getSheet();
  const rows = allDataRows(sheet);
  if (!rows.length) return;
  validarFilas(sheet, rows);
  SpreadsheetApp.getActiveSpreadsheet().toast("Review complete. Please check column JM.", "Bonus Audit");
}

// ---------------------------------------------------------------------------

function aplicarCheckboxes(sheet, rows) {
  const first = rows[0];
  const n = rows[rows.length - 1] - first + 1;
  const checks = sheet.getRange(first, COL_CHECK, n, 1).getValues();

  const freeze = [], restore = [];
  rows.forEach(r => {
    const v = checks[r - first][0];
    (v === true || String(v).toUpperCase() === "TRUE" ? freeze : restore).push(r);
  });

  // Freeze: current values (from the formulas) written back as plain values.
  if (freeze.length) {
    SpreadsheetApp.flush();
    const vals = sheet.getRange(first, COL_JE, n, 3).getValues();
    runs(freeze).forEach(([start, len]) => {
      sheet.getRange(start, COL_JE, len, 3)
        .setValues(vals.slice(start - first, start - first + len));
    });
  }

  // Restore: copy JE3:JG3 formulas; relative references adjust to each row.
  if (restore.length) {
    const template = sheet.getRange(TEMPLATE_ROW, COL_JE, 1, 3);
    runs(restore).forEach(([start, len]) => {
      template.copyTo(sheet.getRange(start, COL_JE, len, 3),
        SpreadsheetApp.CopyPasteType.PASTE_FORMULA, false);
    });
  }
  SpreadsheetApp.flush();
}

function validarFilas(sheet, rows) {
  const first = rows[0];
  const n = rows[rows.length - 1] - first + 1;
  const values = sheet.getRange(first, COL_JE, n, COL_COMMENT - COL_JE + 1).getValues();
  runs(rows).forEach(([start, len]) => {
    const out = [];
    for (let r = start; r < start + len; r++) {
      const v = values[r - first];
      out.push([evaluarBono(v[0], v[1], v[2], v[COL_COMMENT - COL_JE])]);
    }
    sheet.getRange(start, COL_RESULT, len, 1).setValues(out);
  });
}

/** Same rules as the original validarBonosAvanzado, for one row. */
function evaluarBono(jeVal, jfVal, jgVal, commentVal) {
  var je = parseFloat(jeVal) || 0;
  var jf = parseFloat(jfVal) || 0;
  var jg = parseFloat(jgVal) || 0;
  var total = je + jf + jg;
  var comment = commentVal ? String(commentVal) : "";

  if (comment === "" || isNaN(total)) return "";

  var text = comment.replace(/,/g, '');
  var numbers = text.match(/\d+(?:\.\d+)?/g);
  var status = "Mismatch";

  if (numbers) {
    var nums = numbers.map(Number);

    var hasExactMatch = nums.some(function(n) { return Math.abs(n - total) < 0.05; });

    var sumAll = nums.reduce(function(a, b) { return a + b; }, 0);
    var matchesSum = Math.abs(sumAll - total) < 0.05;

    var maxNum = Math.max.apply(null, nums);
    var matchesSumMinusMax = Math.abs((sumAll - maxNum) - total) < 0.05;

    var components = [];
    if (je > 0) components.push(je);
    if (jf > 0) components.push(jf);
    if (jg > 0) components.push(jg);

    var allComponentsFound = false;
    if (components.length > 0) {
      var numsCopy = nums.slice();
      allComponentsFound = true;
      for (var k = 0; k < components.length; k++) {
        var comp = components[k];
        var foundIdx = numsCopy.findIndex(function(n) { return Math.abs(n - comp) < 0.05; });
        if (foundIdx !== -1) {
          numsCopy.splice(foundIdx, 1);
        } else {
          allComponentsFound = false;
          break;
        }
      }
    }

    var zeroMatch = (total === 0 && nums.indexOf(0) !== -1);

    if (hasExactMatch || matchesSum || matchesSumMinusMax || allComponentsFound || zeroMatch) {
      status = "OK";
    }
  }

  return status === "Mismatch" ? "Mismatch: Total = " + total.toFixed(2) : "OK";
}

// ---------------------------------------------------------------------------

function getSheet() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
  if (!sheet) throw new Error('Sheet "' + SHEET_NAME + '" not found.');
  return sheet;
}

function allDataRows(sheet) {
  const rows = [];
  for (let r = FIRST_ROW; r <= sheet.getLastRow(); r++) rows.push(r);
  return rows;
}

/** Data rows of the given ranges that intersect columns [c1, c2]. */
function rowsTouching(ranges, c1, c2) {
  const rows = [];
  ranges.forEach(rg => {
    if (rg.getLastColumn() < c1 || rg.getColumn() > c2) return;
    const last = Math.min(rg.getLastRow(), rg.getSheet().getLastRow());
    for (let r = Math.max(rg.getRow(), FIRST_ROW); r <= last; r++) rows.push(r);
  });
  return rows;
}

function uniqueSorted(rows) {
  return Array.from(new Set(rows)).sort((a, b) => a - b);
}

/** Sorted rows -> [[start, length], ...] contiguous runs. */
function runs(rows) {
  const out = [];
  let i = 0;
  while (i < rows.length) {
    let j = i;
    while (j + 1 < rows.length && rows[j + 1] === rows[j] + 1) j++;
    out.push([rows[i], j - i + 1]);
    i = j + 1;
  }
  return out;
}
