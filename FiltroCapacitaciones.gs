/**
 * Completion Tracker – training filter
 *
 * In A3 (with the source link in A1):
 *   =FILTRAR_CAPACITACION(IMPORTRANGE(A1, "Compliance!A8:T"), I2, I3:I20)
 *
 * I2 holds one or more Issue Types separated by commas, in priority order:
 *   Onsite Activities – Training Iniciatives, Training / Huddles
 *   - 1st Issue Type (priority): lenient word check (DEFAULT_MIN_MATCH).
 *   - 2nd, 3rd... Issue Types: stricter word check; words that are already in
 *     the Issue Type name ("training", "huddles") do not count, and at least
 *     SECONDARY_MIN_WORDS distinctive words of the training must be found.
 *   An Issue Type can be written in part ("Training / Huddles" matches
 *   "Onsite Activities – Training / Huddles").
 *
 * Returns WD ID | Issue Type | Comment | Date | Match % | Training,
 * priority rows first, then the others, each in source order.
 *
 * Tolerant on purpose: case, accents, punctuation, dashes, word order,
 * small typos ("Trainig", "insder"), shortened words ("train") and missing
 * words all still match. Numbers/dates in the training name are ignored,
 * because supervisors rarely type them the same way.
 */

// Source columns inside Compliance!A8:T (0 = column A)
const SRC_DATE = 1;      // B  Date of Offense
const SRC_WDID = 2;      // C  WDID
const SRC_ISSUE = 4;     // E  Issue Type
const SRC_COMMENT = 5;   // F  Supervisor Comment

const DEFAULT_MIN_MATCH = 0.3;      // 1st Issue Type: share of training words found
const SECONDARY_MIN_MATCH = 0.4;    // other Issue Types: share of distinctive words found
const SECONDARY_MIN_WORDS = 2;      // other Issue Types: distinctive words found, at least

const STOPWORDS = ["the", "a", "an", "of", "for", "to", "and", "in", "on", "at", "by",
  "with", "de", "la", "el", "los", "las", "y", "en", "del", "para", "por", "con"];

/**
 * Filters the Compliance source by Issue Type and training name.
 *
 * @param {Array} datos IMPORTRANGE(A1, "Compliance!A8:T")
 * @param {string} issueTypes Issue Types separated by commas, priority first (I2)
 * @param {Array} capacitaciones Training names (I3:I20)
 * @param {number} [minimo] Optional minimum match for the 1st Issue Type, 0-1 (default 0.3)
 * @customfunction
 */
function FILTRAR_CAPACITACION(datos, issueTypes, capacitaciones, minimo) {
  if (!Array.isArray(datos)) return [["No data. Check the link in A1 and allow access in IMPORTRANGE."]];
  const minMatch = (typeof minimo === "number" && minimo > 0) ? minimo : DEFAULT_MIN_MATCH;

  const issues = String(issueTypes == null ? "" : issueTypes).split(",")
    .map(t => palabrasIssue(t)).filter(w => w.length);
  if (!issues.length) return [["Write at least one Issue Type in I2."]];

  const nombres = [].concat(capacitaciones || []).flat()
    .map(String).map(s => s.trim()).filter(s => s !== "");
  const frases = nombres.map(n => ({ nombre: n, palabras: palabrasClave(n) }))
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

    let best = { score: frases.length ? 0 : 1, nombre: "" };
    frases.forEach(f => {
      let palabras = f.palabras, minWords = 1, min = minMatch;
      if (nivel > 0) {
        // Stricter: words already in the Issue Type name prove nothing.
        palabras = f.palabras.filter(p => !rowIssue.some(w => palabraCoincide(p, w)));
        if (!palabras.length) return;
        minWords = Math.min(SECONDARY_MIN_WORDS, palabras.length);
        min = SECONDARY_MIN_MATCH;
      }
      const found = encontradas(palabras, words, joined);
      const s = found / palabras.length;
      if (found >= minWords && s >= min && s > best.score) best = { score: s, nombre: f.nombre };
    });

    if (best.score > 0) {
      porNivel[nivel].push([wdid, row[SRC_ISSUE], comment, row[SRC_DATE],
        Math.round(best.score * 100) / 100, best.nombre]);
    }
  });

  const out = [].concat.apply([], porNivel);
  return out.length ? out : [["No matches", "", "", "", "", ""]];
}

// ---------------------------------------------------------------------------

/** Lowercase, no accents, any punctuation or dash -> single space. */
function normalizar(v) {
  return String(v == null ? "" : v)
    .toLowerCase()
    .normalize("NFD").replace(/[\u0300-\u036f]/g, "")
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

/** Words of an Issue Type from I2 (no stopwords). */
function palabrasIssue(texto) {
  return normalizar(texto).split(" ").filter(w => w && STOPWORDS.indexOf(w) === -1);
}

/** True when every word of the I2 Issue Type is in the row's Issue Type. */
function contieneTodas(rowIssue, wanted) {
  return wanted.every(p => rowIssue.some(w => palabraCoincide(p, w)));
}

/** How many training words are found in the comment. */
function encontradas(palabras, words, joined) {
  let found = 0;
  palabras.forEach(p => {
    if (words.some(w => palabraCoincide(p, w))) found++;
    else if (p.length >= 5 && joined.indexOf(p) !== -1) found++;   // "agentinsider"
  });
  return found;
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
