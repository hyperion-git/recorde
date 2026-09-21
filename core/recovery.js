import { DEFAULT_SETTINGS } from './settings.js';
// Self-healing for equations whose custom XML part is gone.
//
// When a picture survives but its XML part doesn't (deleted by a foreign editor,
// or pasted from another doc), the only source left is the picture's
// altTextDescription. That field is also where users type arbitrary captions, so
// we must guess whether it's recoverable LaTeX before re-rendering — a wrong
// guess would silently turn "Figure 3 our results" into a broken equation.
// Pure: numbers/strings only, no Office/Word/MathJax/DOM. taskpane.js owns the
// side effects; this module owns the decision and the degraded payload it builds.

// Heuristic source-vs-prose classifier. Tuned to err toward false (skip) on
// ambiguous text: a missed recovery is recoverable by the user, a false positive
// corrupts visible content.
export function looksLikeEquation(s) {
  if (typeof s !== 'string') return false;
  const t = s.trim();
  if (!t || t.length > 2000) return false;
  // LaTeX control chars / grouping / sub-superscript / math delimiters are a
  // strong signal — prose almost never contains them.
  if (/[\\{}^_$]/.test(t)) return true;
  // Otherwise require a math operator AND that it doesn't read like a sentence
  // (two-or-more runs of letters separated by non-letters → prose like
  // "Figure 3 our results"; "1+2=3" or "x=y" survive).
  if (!/[=+\-*/<>]/.test(t)) return false;
  return !/[A-Za-z]{2,}[^A-Za-z]+[A-Za-z]{2,}/.test(t);
}

// Build a payload that re-renders the recovered source and re-writes a fresh XML
// part. The `degraded` flag is for the caller's telemetry/UI only; buildXml
// ignores it (it reads uuid/latex/font/sizePt/displayMode), so it never reaches
// storage — proven by the round-trip test. Font defaults to the app default (Termes) and mode to
// 'inline' because the original choices are unknowable from alt text alone.
export function buildDegradedPayload({ uuid, latex, sizePt }) {
  const n = Number(sizePt);
  const validSize = Number.isFinite(n) && n >= 6 && n <= 96 ? n : 11;
  return {
    uuid,
    latex: String(latex).trim(),
    font: DEFAULT_SETTINGS.font,
    sizePt: validSize,
    displayMode: 'inline',
    degraded: true,
  };
}
