/**
 * Completion Tracker – training filter
 *
 * Layout: column A holds the inputs, column B is a spacer, results in C:I.
 *   A1      link to the source spreadsheet
 *   A2      Issue Types
 *   A3:A20  training names
 *   C3      =FILTRAR_CAPACITACION(IMPORTRANGE(A1, "Compliance!A8:T"), A2, A3:A20)
 *
 * A2 holds one or more Issue Types separated by commas, in priority order:
 *   Onsite Activities – Training Iniciatives, Training / Huddles
 *   - 1st Issue Type (priority): every row is kept. The comment is only
 *     compared to A3:A20 to tell which training it was.
 *   - 2nd, 3rd... Issue Types: a row is kept only when its comment matches a
 *     training in A3:A20 with a stricter check: words that are already in the
 *     Issue Type name ("training", "huddles") do not count, and at least
 *     SECONDARY_MIN_WORDS distinctive words must be found.
 *   An Issue Type can be written in part ("Training / Huddles" matches
 *   "Onsite Activities – Training / Huddles").
 *
 * When a training name carries a date ("... Weekly Edition - 10.5.2026"),
 * the row's Date of Offense must fall between that date and DATE_WINDOW_DAYS
 * later. Otherwise the row is still listed (the comment does mention it) but
 * the training is flagged instead of assigned.
 *
 * Returns WD ID | Issue Type | Comment | Date | Approved? | Training | Match %
 * priority rows first, then the others, each in source order.
 *
 * Tolerant on purpose: case, accents, punctuation, dashes, word order,
 * small typos ("Trainig", "insder"), shortened words ("train") and missing
 * words all still match.
 */

// Source columns inside Compliance!A8:T (0 = column A)
const SRC_STATUS = 0;    // A  Status (Approved/Denied)
const SRC_DATE = 1;      // B  Date of Offense
const SRC_WDID = 2;      // C  WDID
const SRC_ISSUE = 4;     // E  Issue Type
const SRC_COMMENT = 5;   // F  Supervisor Comment

const DEFAULT_MIN_MATCH = 0.3;      // 1st Issue Type: share of training words to name the training
const SECONDARY_MIN_MATCH = 0.4;    // other Issue Types: share of distinctive words found
const SECONDARY_MIN_WORDS = 2;      // other Issue Types: distinctive words found, at least
const DATE_WINDOW_DAYS = 6;         // training date + 6 days = the whole week

const STOPWORDS = ["the", "a", "an", "of", "for", "to", "and", "in", "on", "at", "by",
  "with", "de", "la", "el", "los", "las", "y", "en", "del", "para", "por", "con"];

/**
 * Filters the Compliance source by Issue Type and training name.
 *
 * @param {Array} datos IMPORTRANGE(A1, "Compliance!A8:T")
 * @param {string} issueTypes Issue Types separated by commas, priority first (A2)
 * @param {Array} capacitaciones Training names (A3:A20)
 * @param {number} [minimo] Optional minimum match to name a training for the 1st Issue Type, 0-1 (default 0.3)
 * @customfunction
 */
function FILTRAR_CAPACITACION(datos, issueTypes, capacitaciones, minimo) {
  if (!Array.isArray(datos)) return [["No data. Check the link in A1 and allow access in IMPORTRANGE."]];
  const minMatch = (typeof minimo === "number" && minimo > 0) ? minimo : DEFAULT_MIN_MATCH;

  const issues = String(issueTypes == null ? "" : issueTypes).split(",")
    .map(t => palabrasIssue(t)).filter(w => w.length);
  if (!issues.length) return [["Write at least one Issue Type in A2."]];

  const nombres = [].concat(capacitaciones || []).flat()
    .map(String).map(s => s.trim()).filter(s => s !== "");
  const frases = nombres.map(n => ({ nombre: n, palabras: palabrasClave(n), fecha: fechaDeNombre(n) }))
    .filter(f => f.palabras.length);

  const porNivel = issues.map(() => []);
  datos.forEach(row => {
    const wdid = row[SRC_WDID];
    if (wdid === "" || wdid == null || String(wdid).trim().toUpperCase() === "WDID") return;

    const rowIssue = normalizar(row[SRC_ISSUE]).split(" ").filter(w => w);
    const nivel = issues.findIndex(t => contieneTodas(rowIssue, t));
    if (nivel === -1) return;

    const comment = row[SRC_COMMENT];
    const words = normalizar(comment).split(" ").filter(w => w);
    const joined = words.join("");
    const rowDate = fechaDeFila(row[SRC_DATE]);

    // Every training the comment matches, with its date check.
    const candidatos = [];
    frases.forEach(f => {
      let palabras = f.palabras, minWords = 1, min = minMatch, ignoradas = [];
      if (nivel > 0) {
        // Stricter: words already in the Issue Type name prove nothing.
        ignoradas = f.palabras.filter(p => rowIssue.some(w => palabraCoincide(p, w)));
        palabras = f.palabras.filter(p => ignoradas.indexOf(p) === -1);
        if (!palabras.length) return;
        minWords = Math.min(SECONDARY_MIN_WORDS, palabras.length);
        min = SECONDARY_MIN_MATCH;
      }
      const found = encontradas(palabras, words, joined);
      const score = found.length / palabras.length;
      if (found.length < minWords || score < min) return;
      const rango = f.fecha ? rangoFechas(f.fecha, rowDate) : null;
      const fechaOk = !rango || (rowDate && rowDate >= rango.inicio && rowDate <= rango.fin);
      candidatos.push({ f, found, total: palabras.length, score, ignoradas, rango, fechaOk });
    });

    // Secondary Issue Types need a match; the priority one is always kept.
    if (nivel > 0 && !candidatos.length) return;

    candidatos.sort((a, b) => (b.fechaOk - a.fechaOk) || (b.score - a.score));
    const c = candidatos[0];
    let training, detalle;
    if (!frases.length) {
      training = "";
      detalle = "Priority Issue Type · no training names in A3:A20";
    } else if (!c) {
      training = "⚠ Not identified";
      detalle = "0% · Priority Issue Type, kept anyway · comment does not mention any training in A3:A20";
    } else {
      const partes = [Math.round(c.score * 100) + "%",
        "found " + c.found.join(", ") + " (" + c.found.length + " of " + c.total + " key words)"];
      if (c.ignoradas.length) partes.push("not counted: " + c.ignoradas.join(", ") + " (part of the Issue Type)");
      partes.push(nivel === 0 ? "Priority Issue Type" : "Secondary Issue Type, strict check passed");
      if (c.fechaOk) {
        training = c.f.nombre;
        if (c.rango) partes.push("date within " + fmt(c.rango.inicio) + "–" + fmt(c.rango.fin));
      } else {
        training = "⚠ Date mismatch – check";
        partes.push("looks like \"" + c.f.nombre + "\" but date " +
          (rowDate ? fmt(rowDate) : "is missing") + " is outside " + fmt(c.rango.inicio) + "–" + fmt(c.rango.fin));
      }
      detalle = partes.join(" · ");
    }

    const status = String(row[SRC_STATUS] == null ? "" : row[SRC_STATUS]).trim() || "Pending";
    porNivel[nivel].push([wdid, row[SRC_ISSUE], comment, row[SRC_DATE], status, training, detalle]);
  });

  const out = [].concat.apply([], porNivel);
  return out.length ? out : [["No matches", "", "", "", "", "", ""]];
}

// ---------------------------------------------------------------------------

/** Lowercase, no accents, any punctuation or dash -> single space. */
function normalizar(v) {
  return String(v == null ? "" : v)
    .toLowerCase()
    .normalize("NFD").replace(/[̀-ͯ]/g, "")
    .replace(/[^a-z0-9]+/g, " ")
    .trim();
}

/** Words of a training name that matter: no stopwords, no numbers/dates. */
function palabrasClave(nombre) {
  const seen = {};
  return normalizar(nombre).split(" ")
    .filter(w => w.length > 1 && !/^\d+$/.test(w) && STOPWORDS.indexOf(w) === -1)
    .filter(w => (seen[w] ? false : (seen[w] = true)));
}

/** Words of an Issue Type from A2 (no stopwords). */
function palabrasIssue(texto) {
  return normalizar(texto).split(" ").filter(w => w && STOPWORDS.indexOf(w) === -1);
}

/** True when every word of the A2 Issue Type is in the row's Issue Type. */
function contieneTodas(rowIssue, wanted) {
  return wanted.every(p => rowIssue.some(w => palabraCoincide(p, w)));
}

/** Training words found in the comment. */
function encontradas(palabras, words, joined) {
  return palabras.filter(p =>
    words.some(w => palabraCoincide(p, w)) ||
    (p.length >= 5 && joined.indexOf(p) !== -1));                 // "agentinsider"
}

/** Same word allowing typos and shortened forms ("train" ~ "training"). */
function palabraCoincide(p, w) {
  if (p === w) return true;
  if (w.length >= 4 && p.indexOf(w) === 0) return true;           // train -> training
  if (p.length >= 4 && w.indexOf(p) === 0) return true;           // insider -> insiders
  const maxErr = p.length >= 8 ? 2 : p.length >= 4 ? 1 : 0;
  return maxErr > 0 && Math.abs(p.length - w.length) <= maxErr && distancia(p, w) <= maxErr;
}

/** Levenshtein edit distance. */
function distancia(a, b) {
  let prev = [];
  for (let j = 0; j <= b.length; j++) prev[j] = j;
  for (let i = 1; i <= a.length; i++) {
    const cur = [i];
    for (let j = 1; j <= b.length; j++) {
      cur[j] = Math.min(prev[j] + 1, cur[j - 1] + 1,
        prev[j - 1] + (a[i - 1] === b[j - 1] ? 0 : 1));
    }
    prev = cur;
  }
  return prev[b.length];
}

// ---------------------------------------------------------------------------

/** Month/day(/year) inside a training name: 10.5.2026, 10/5, 10-05-26. */
function fechaDeNombre(nombre) {
  const m = String(nombre).match(/(\d{1,2})[.\/-](\d{1,2})(?:[.\/-](\d{2,4}))?/);
  if (!m) return null;
  const mes = +m[1], dia = +m[2];
  if (mes < 1 || mes > 12 || dia < 1 || dia > 31) return null;
  let anio = m[3] ? +m[3] : null;
  if (anio !== null && anio < 100) anio += 2000;
  return { mes, dia, anio };
}

/** Date of Offense as a date at midnight (Date object or "MM/dd/yyyy" text). */
function fechaDeFila(v) {
  if (v instanceof Date && !isNaN(v)) return new Date(v.getFullYear(), v.getMonth(), v.getDate());
  const m = String(v == null ? "" : v).match(/(\d{1,2})\/(\d{1,2})\/(\d{4})/);
  return m ? new Date(+m[3], +m[1] - 1, +m[2]) : null;
}

/** Training date .. training date + DATE_WINDOW_DAYS (year from the row if missing). */
function rangoFechas(f, rowDate) {
  const anio = f.anio || (rowDate ? rowDate.getFullYear() : new Date().getFullYear());
  const inicio = new Date(anio, f.mes - 1, f.dia);
  const fin = new Date(anio, f.mes - 1, f.dia + DATE_WINDOW_DAYS);
  return { inicio, fin };
}

function fmt(d) {
  const p = n => (n < 10 ? "0" : "") + n;
  return p(d.getMonth() + 1) + "/" + p(d.getDate()) + "/" + d.getFullYear();
}
