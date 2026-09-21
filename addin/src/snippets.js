// LaTeX snippet / palette model — pure (strings only), so the caret math is
// unit-tested. The palette UI in taskpane.js renders PALETTE and calls
// applySnippet against the <textarea>'s value + selection. No Office/DOM refs.

// Caret placeholder inside a snippet: where the cursor lands after insertion, and
// where a selected expression is wrapped. U+25AE (▮) never appears in real LaTeX,
// so it's an unambiguous marker.
export const SNIPPET_CURSOR = '▮';

// Insert `snippet` into `value`, replacing the selection [selStart, selEnd).
// Returns { value, caret }:
//   - The snippet's first SNIPPET_CURSOR marks where the caret ends up; with no
//     marker, the caret follows the inserted text.
//   - Currently-selected text is dropped INTO the marker position, so you can
//     select an expression and wrap it (select "x" + "\sqrt{▮}" → "\sqrt{x}").
//     With no marker, the selection is replaced.
// Pure and total: clamps indices, never throws.
export function applySnippet(value, selStart, selEnd, snippet) {
  const v = String(value);
  const start = Math.max(0, Math.min(selStart | 0, v.length));
  const end = Math.max(start, Math.min(selEnd | 0, v.length));
  const selected = v.slice(start, end);

  const m = snippet.indexOf(SNIPPET_CURSOR);
  let body, caretInBody;
  if (m === -1) {
    body = snippet;
    caretInBody = body.length;            // caret after the inserted snippet
  } else {
    const before = snippet.slice(0, m);
    const after = snippet.slice(m + SNIPPET_CURSOR.length);
    body = before + selected + after;     // wrap the selection at the marker
    caretInBody = before.length + selected.length;
  }
  return {
    value: v.slice(0, start) + body + v.slice(end),
    caret: start + caretInBody,
  };
}

// Curated palette. Each item: { label (button text), snippet, title (tooltip) }.
// Plain symbols insert as-is; structures carry a ▮ caret marker (template form).
const C = SNIPPET_CURSOR;
export const PALETTE = [
  {
    name: 'Greek',
    items: [
      { label: 'α', snippet: '\\alpha' },   { label: 'β', snippet: '\\beta' },
      { label: 'γ', snippet: '\\gamma' },    { label: 'δ', snippet: '\\delta' },
      { label: 'ε', snippet: '\\epsilon' },  { label: 'ζ', snippet: '\\zeta' },
      { label: 'η', snippet: '\\eta' },      { label: 'θ', snippet: '\\theta' },
      { label: 'κ', snippet: '\\kappa' },    { label: 'λ', snippet: '\\lambda' },
      { label: 'μ', snippet: '\\mu' },       { label: 'ν', snippet: '\\nu' },
      { label: 'ξ', snippet: '\\xi' },       { label: 'π', snippet: '\\pi' },
      { label: 'ρ', snippet: '\\rho' },      { label: 'σ', snippet: '\\sigma' },
      { label: 'τ', snippet: '\\tau' },      { label: 'φ', snippet: '\\phi' },
      { label: 'χ', snippet: '\\chi' },      { label: 'ψ', snippet: '\\psi' },
      { label: 'ω', snippet: '\\omega' },
      { label: 'Γ', snippet: '\\Gamma' },    { label: 'Δ', snippet: '\\Delta' },
      { label: 'Θ', snippet: '\\Theta' },    { label: 'Λ', snippet: '\\Lambda' },
      { label: 'Ξ', snippet: '\\Xi' },       { label: 'Π', snippet: '\\Pi' },
      { label: 'Σ', snippet: '\\Sigma' },    { label: 'Φ', snippet: '\\Phi' },
      { label: 'Ψ', snippet: '\\Psi' },      { label: 'Ω', snippet: '\\Omega' },
    ],
  },
  {
    name: 'Relations',
    items: [
      { label: '≤', snippet: '\\leq' },      { label: '≥', snippet: '\\geq' },
      { label: '≠', snippet: '\\neq' },      { label: '≈', snippet: '\\approx' },
      { label: '≡', snippet: '\\equiv' },    { label: '∝', snippet: '\\propto' },
      { label: '∈', snippet: '\\in' },       { label: '⊂', snippet: '\\subset' },
      { label: '∪', snippet: '\\cup' },      { label: '∩', snippet: '\\cap' },
      { label: '→', snippet: '\\to' },       { label: '↔', snippet: '\\leftrightarrow' },
      { label: '⇒', snippet: '\\Rightarrow' }, { label: '∀', snippet: '\\forall' },
      { label: '∃', snippet: '\\exists' },   { label: '¬', snippet: '\\neg' },
    ],
  },
  {
    name: 'Operators',
    items: [
      { label: '±', snippet: '\\pm' },       { label: '×', snippet: '\\times' },
      { label: '·', snippet: '\\cdot' },     { label: '÷', snippet: '\\div' },
      { label: '∞', snippet: '\\infty' },    { label: '∂', snippet: '\\partial' },
      { label: '∇', snippet: '\\nabla' },    { label: '∑', snippet: `\\sum_{${C}}^{}` },
      { label: '∏', snippet: `\\prod_{${C}}^{}` }, { label: '∫', snippet: `\\int_{${C}}^{}` },
      { label: '∮', snippet: `\\oint_{${C}}^{}` },
    ],
  },
  {
    name: 'Structures',
    items: [
      { label: 'a/b', snippet: `\\frac{${C}}{}`,        title: 'Fraction' },
      { label: '√',   snippet: `\\sqrt{${C}}`,          title: 'Square root' },
      { label: 'ⁿ√',  snippet: `\\sqrt[${C}]{}`,        title: 'nth root' },
      { label: 'xⁿ',  snippet: `^{${C}}`,               title: 'Superscript' },
      { label: 'xₙ',  snippet: `_{${C}}`,               title: 'Subscript' },
      { label: '( )', snippet: `\\left(${C}\\right)`,   title: 'Auto-sized parens' },
      { label: '[ ]', snippet: `\\left[${C}\\right]`,   title: 'Auto-sized brackets' },
      { label: 'vec', snippet: `\\vec{${C}}`,           title: 'Vector arrow' },
      { label: 'hat', snippet: `\\hat{${C}}`,           title: 'Hat' },
      { label: 'bar', snippet: `\\bar{${C}}`,           title: 'Bar' },
      { label: 'dot', snippet: `\\dot{${C}}`,           title: 'Dot' },
      { label: 'matrix', snippet: `\\begin{matrix} ${C} \\end{matrix}`, title: 'Matrix' },
      { label: 'cases',  snippet: `\\begin{cases} ${C} \\end{cases}`,   title: 'Cases' },
    ],
  },
];
