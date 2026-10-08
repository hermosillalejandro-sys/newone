/**
 * Stack Ranking – bonuses
 *
 *  - JM7: =VALIDAR_BONOS(JE7:JG, JJ7:JJ)
 *    validarBonosAvanzado() writes that formula in JM7 for you.
 *    Live check: "OK" when the money in JE:JG matches what is approved in
 *    Comments (JJ), "Mismatch: Total = ..." when it must be reviewed.
 *    It recalculates by itself whenever JE:JG or JJ change.
 *
 *  - Checkbox in IZ:
 *      TRUE  -> JE:JG of that row are frozen as values.
 *      FALSE -> JE:JG get the base formula back (template JE3:JG3).
 *    Every IZ edit checks ALL rows, so pastes, fill down, several checkboxes
 *    at once, filtered ranges or a click missed earlier are all covered.
 *    Rows already in the right state are left alone.
 */

const SHEET_NAME = "Stack Ranking";
const FIRST_ROW = 7;           // first data row (rows 5-6 are headers)
const TEMPLATE_ROW = 3;        // JE3:JG3 hold the base formulas
const COL_CHECK = 260;         // IZ
const COL_JE = 265;            // JE (Tech + Booster), JF, JG
const COL_RESULT = 273;        // JM
const JM_FORMULA = "=VALIDAR_BONOS(JE7:JG, JJ7:JJ)";

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

  // Only react when the edit touches IZ (e.range or any part of the selection).
  const ranges = [e.range];
  try {
    const list = sheet.getActiveRangeList();
    if (list) list.getRanges().forEach(r => ranges.push(r));
  } catch (err) {}
  if (!rowsTouching(ranges, COL_CHECK, COL_CHECK).length) return;

  const lock = LockService.getDocumentLock();
  if (!lock.tryLock(20000)) return;   // the next IZ edit catches up on this one
  try {
    aplicarCheckboxes(sheet);
  } finally {
    lock.releaseLock();
  }
}

/**
 * Puts the live formula in JM7 (clears JM7 down first so the results have
 * room). Running it again is harmless.
 */
function validarBonosAvanzado() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
  if (!sheet) throw new Error('Sheet "' + SHEET_NAME + '" not found.');
  const n = sheet.getMaxRows() - FIRST_ROW + 1;
  if (n <= 0) return;
  sheet.getRange(FIRST_ROW, COL_RESULT, n, 1).clearContent();
  sheet.getRange(FIRST_ROW, COL_RESULT).setFormula(JM_FORMULA);
  SpreadsheetApp.getActiveSpreadsheet().toast('Review complete. Please check column JM.', 'Bonus Audit');
}

/** Same check as every IZ edit, from the menu. */
function sincronizarCheckboxes() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
  if (!sheet) throw new Error('Sheet "' + SHEET_NAME + '" not found.');
  const done = aplicarCheckboxes(sheet);
  SpreadsheetApp.getActiveSpreadsheet().toast(
    done.frozen + " rows frozen, " + done.restored + " formulas restored.", "Bonus Audit");
}

/**
 * Validates every row: "OK" or "Mismatch: Total = ...".
 * Use once in JM7: =VALIDAR_BONOS(JE7:JG, JJ7:JJ)
 *
 * @param {Array} montos JE:JG (Tech + Booster, Total Performance, Total Compliance)
 * @param {Array} comentarios JJ (Comments)
 * @customfunction
 */
function VALIDAR_BONOS(montos, comentarios) {
  if (!Array.isArray(montos)) montos = [[montos, 0, 0]];
  if (!Array.isArray(comentarios)) comentarios = [[comentarios]];
  return montos.map(function(m, i) {
    var c = comentarios[i] ? comentarios[i][0] : "";
    return [evaluarBono(m[0], m[1], m[2], c)];
  });
}

// ---------------------------------------------------------------------------

/**
 * Brings every data row in line with its IZ checkbox:
 *   IZ TRUE  and JE:JG still has formulas -> freeze as values.
 *   IZ FALSE and JE:JG has no formulas    -> restore JE3:JG3 formulas.
 * Rows already in the right state are not touched, so a row missed by an
 * earlier edit is fixed on the next one.
 */
function aplicarCheckboxes(sheet) {
  const n = sheet.getLastRow() - FIRST_ROW + 1;
  if (n <= 0) return { frozen: 0, restored: 0 };
  const checks = sheet.getRange(FIRST_ROW, COL_CHECK, n, 1).getValues();
  const block = sheet.getRange(FIRST_ROW, COL_JE, n, 3);
  const formulas = block.getFormulas();

  const freeze = [], restore = [];
  for (let i = 0; i < n; i++) {
    const v = checks[i][0];
    const checked = v === true || String(v).toUpperCase() === "TRUE";
    const hasFormula = formulas[i].some(f => f !== "");
    if (checked && hasFormula) freeze.push(FIRST_ROW + i);
    else if (!checked && !hasFormula) restore.push(FIRST_ROW + i);
  }

  // Freeze: current values (from the formulas) written back as plain values.
  if (freeze.length) {
    const vals = block.getValues();
    runs(freeze).forEach(([start, len]) => {
      const k = start - FIRST_ROW;
      sheet.getRange(start, COL_JE, len, 3).setValues(vals.slice(k, k + len));
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
  return { frozen: freeze.length, restored: restore.length };
}

/** Same rules as the original validarBonosAvanzado script, for one row. */
function evaluarBono(jeVal, jfVal, jgVal, commentVal) {
  var je = parseFloat(jeVal) || 0;
  var jf = parseFloat(jfVal) || 0;
  var jg = parseFloat(jgVal) || 0;
  var total = je + jf + jg;
  // A typed "0.00" arrives as the number 0: it is still a comment.
  var comment = (commentVal === "" || commentVal == null) ? "" : String(commentVal);

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
