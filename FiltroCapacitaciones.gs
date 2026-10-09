/**
 * Completion Tracker – training filter
 *
 * In A3 (with the source link in A1):
 *   =FILTRAR_CAPACITACION(IMPORTRANGE(A1, "Compliance!A8:T"), I2, I3:I20)
 *
 * Returns WD ID | Issue Type | Comment | Date | Match % | Training
 * for every source row whose Issue Type matches I2 and whose Supervisor
 * Comment looks like any training name in I3:I20.
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

const ISSUE_MIN_SIMILARITY = 0.85;  // Issue Type comes from a drop-down: strict
const DEFAULT_MIN_MATCH = 0.3;      // share of training words found in the comment

const STOPWORDS = ["the", "a", "an", "of", "for", "to", "and", "in", "on", "at", "by",
  "with", "de", "la", "el", "los", "las", "y", "en", "del", "para", "por", "con"];

/**
 * Filters the Compliance source by Issue Type and training name.
 *
 * @param {Array} datos IMPORTRANGE(A1, "Compliance!A8:T")
 * @param {string} issueType Issue Type to keep (I2)
 * @param {Array} capacitaciones Training names (I3:I20)
 * @param {number} [minimo] Optional minimum match, 0-1 (default 0.3)
 * @customfunction
 */
function FILTRAR_CAPACITACION(datos, issueType, capacitaciones, minimo) {
  if (!Array.isArray(datos)) return [["No data. Check the link in A1 and allow access in IMPORTRANGE."]];
  const minMatch = (typeof minimo === "number" && minimo > 0) ? minimo : DEFAULT_MIN_MATCH;
  const wantedIssue = normalizar(issueType);

  const nombres = [].concat(capacitaciones || []).flat()
    .map(String).map(s => s.trim()).filter(s => s !== "");
  const frases = nombres.map(n => ({ nombre: n, palabras: palabrasClave(n) }))
    .filter(f => f.palabras.length);

  const out = [];
  datos.forEach(row => {
    const wdid = row[SRC_WDID];
    if (wdid === "" || wdid == null || String(wdid).trim().toUpperCase() === "WDID") return;
    if (wantedIssue && similitud(normalizar(row[SRC_ISSUE]), wantedIssue) < ISSUE_MIN_SIMILARITY) return;

    const comment = row[SRC_COMMENT];
    let best = { score: frases.length ? 0 : 1, nombre: "" };
    if (frases.length) {
      const words = normalizar(comment).split(" ").filter(w => w);
      const joined = words.join("");
      frases.forEach(f => {
        const s = puntaje(f.palabras, words, joined);
        if (s > best.score) best = { score: s, nombre: f.nombre };
      });
    }
    if (best.score >= minMatch) {
      out.push([wdid, row[SRC_ISSUE], comment, row[SRC_DATE], Math.round(best.score * 100) / 100, best.nombre]);
    }
  });
  return out.length ? out : [["No matches", "", "", "", "", ""]];
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

/** Share of the training words found in the comment (0-1). */
function puntaje(palabras, words, joined) {
  let found = 0;
  palabras.forEach(p => {
    if (words.some(w => palabraCoincide(p, w))) found++;
    else if (p.length >= 5 && joined.indexOf(p) !== -1) found++;   // "agentinsider"
  });
  return found / palabras.length;
}

/** Same word allowing typos and shortened forms ("train" ~ "training"). */
function palabraCoincide(p, w) {
  if (p === w) return true;
  if (w.length >= 4 && p.indexOf(w) === 0) return true;           // train -> training
  if (p.length >= 4 && w.indexOf(p) === 0) return true;           // insider -> insiders
  const maxErr = p.length >= 8 ? 2 : p.length >= 4 ? 1 : 0;
  return maxErr > 0 && Math.abs(p.length - w.length) <= maxErr && distancia(p, w) <= maxErr;
}

/** 1 = identical, 0 = completely different. */
function similitud(a, b) {
  if (a === b) return 1;
  const len = Math.max(a.length, b.length);
  return len ? 1 - distancia(a, b) / len : 1;
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
