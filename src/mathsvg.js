// Host-independent rendering/sizing helpers.
//
// Everything here is pure (numbers/strings) or operates only on a passed-in SVG
// DOM element — no Office.js, no globals beyond the DOM the caller hands in. That
// makes this module unit-testable under `node --test` (with a DOM shim for the
// SVG-element helpers). taskpane.js imports from here; keep it free of any
// `Office`/`Word`/`MathJax` references.

// MathJax TeX-font ex-to-em ratio. Used only by the viewBox-less fallback in
// scaleSvgToPt to convert MathJax's ex-unit dimensions into absolute pt.
export const EX_PER_EM = 0.430554;

// MathJax 4 font name convention: 'mathjax-<short>'. Maps the dropdown's short
// value to the full font identifier MathJax expects.
export function mathJaxFontName(short) {
  return 'mathjax-' + short;
}

// Newlines in math mode should be equivalent to spaces per TeX semantics, but in
// practice the AMS + \newcommand packages treat some token streams differently
// when split across lines. Flatten before the MathJax boundary; storage and the
// textarea keep the user's newlines so long inputs stay readable.
export function flattenLatex(s) {
  return String(s).replace(/[\r\n]+/g, ' ');
}

// Size policy:
//   selection → body when Word can't resolve a unique size (empty/mixed/unsized
//     selection): ambiguous selections behave like a cursor position.
//   No cross-mode capping: a legitimate 96pt title deserves a 96pt equation.
//   clamp() defends against stale/corrupt inputs (e.g. a sizePt loaded from an
//     older XML part written before this validation existed).
export function computeRenderSize({ mode, selectionPt, bodyPt, fixedPt }) {
  const MIN_PT = 6, MAX_PT = 96;
  const clamp = (pt, fallback) =>
    Number.isFinite(pt) && pt >= MIN_PT && pt <= MAX_PT ? pt : fallback;

  const body = clamp(bodyPt, 11);
  if (mode === 'fixed') return clamp(fixedPt, body);
  if (mode === 'body') return body;
  return clamp(selectionPt, body);
}

// Inline descent in pt from a MathJax viewBox string ("vx vy vw vh"), with the
// math baseline at SVG y=0:
//   ascent  = -vy       (>= 0; content above baseline)
//   descent = vy + vh   (>= 0 for ink at/below baseline; small for non-descended
//                        equations like 1+2+3, large for integrals/fractions)
//   descent_pt = descent * targetPt / 1000   (1000 internal units = 1 em)
// Returns 0 when the viewBox is missing/malformed.
export function descentPtFromViewBox(viewBox, targetPt) {
  const vb = String(viewBox || '').trim().split(/\s+/).map(parseFloat);
  if (vb.length !== 4 || !vb.every(Number.isFinite)) return 0;
  const [, vy, , vh] = vb;
  const descent = vy + vh;
  return descent > 0 ? descent * targetPt / 1000 : 0;
}

// Scale the SVG to the target size by deriving width/height from the viewBox
// geometry. MathJax's viewBox uses 1000 units = 1 em — stable across v3/v4. The
// width/height *attributes* depend on MathJax's unit choice (ex in v3), so
// computing from viewBox sidesteps the unit question. Fallback to the v3
// ex-attribute path when viewBox is absent.
export function scaleSvgToPt(svgEl, targetPt) {
  const vb = (svgEl.getAttribute('viewBox') || '').split(/\s+/).map(parseFloat);
  if (vb.length === 4 && vb.every(Number.isFinite)) {
    const [, , vw, vh] = vb;
    const factor = targetPt / 1000;
    svgEl.setAttribute('width',  (vw * factor).toFixed(3) + 'pt');
    svgEl.setAttribute('height', (vh * factor).toFixed(3) + 'pt');
    return;
  }
  const ptPerEx = EX_PER_EM * targetPt;
  for (const attr of ['width', 'height']) {
    const v = svgEl.getAttribute(attr);
    if (v && v.endsWith('ex')) {
      svgEl.setAttribute(attr, (parseFloat(v) * ptPerEx).toFixed(3) + 'pt');
    }
  }
}

// MathJax SVG uses fill/stroke="currentColor" so it inherits the host page's
// text color. Word's SVG renderer has no CSS context, so currentColor resolves
// to white/transparent and the equation appears invisible. Replace it with an
// explicit color (default black) before insert.
export function inlineColors(svgEl, color = '#000000') {
  svgEl.setAttribute('color', color);
  for (const el of svgEl.querySelectorAll('*')) {
    if (el.getAttribute('fill')   === 'currentColor') el.setAttribute('fill',   color);
    if (el.getAttribute('stroke') === 'currentColor') el.setAttribute('stroke', color);
  }
}

// Centering policy (WP1.4): a paragraph is safe to center only when it holds the
// equation alone — no surrounding prose. Word's paragraph text excludes the
// inline picture, so an equation-only paragraph reads as "empty" once invisible
// characters are stripped. /\s/ catches space/tab/newline but misses NBSP in
// some engines and never covers zero-width space / BOM / vertical tab — all
// formatting artifacts Word readily emits. We strip those explicitly, written as
// \u escapes (not literal glyphs) so the set stays reviewable and survives
// editors that normalize invisible characters:
//   \u00A0 NBSP   \u200B zero-width space   \uFEFF BOM   \u000B vertical tab
export function isParagraphTextEmpty(text) {
  if (text == null) return true;
  return String(text).replace(/[\s\u00A0\u200B\uFEFF\u000B]/g, '') === '';
}

// The message of the first TeX error inside a MathJax render, or null. MathJax
// does not throw on bad input; it renders an <merror> whose SVG node carries the
// message in data-mjx-error. Pure over any DOM-like node with querySelector.
export function findRenderError(node) {
  if (!node || typeof node.querySelector !== 'function') return null;
  const el = node.querySelector('[data-mjx-error]');
  if (!el) return null;
  const msg = el.getAttribute('data-mjx-error');
  return msg && msg.trim() ? msg.trim() : 'Syntax error';
}
