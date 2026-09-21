// Equation numbering (WP2.2) — the pure, host-agnostic core.
//
// Three numbering *styles* share one document-level counter but differ in WHERE
// the number lives, which is the whole reason they need separate strategies:
//   'inline' — the number is rendered INTO the equation SVG (it trails the
//              equation: "…  (3)"). Self-contained, identical on Word web and
//              desktop, and renumbering just re-renders the stored source.
//   'table'  — the number is real Word text in a borderless 1×2 table, so it
//              sits flush-right (journal style). Web-compatible.
//   'field'  — a native Word SEQ field that Word auto-renumbers and that
//              cross-references can target. DESKTOP-ONLY (insertField is not
//              supported on Word web).
//
// This module decides *what* the number is and *whether* a style is usable; the
// Word glue that actually places/renumbers pictures lives in taskpane.js. Keeping
// the decisions here keeps them unit-testable without an Office host.

export const NUMBERING_STYLES = ['inline', 'table', 'field'];

// Per-document numbering state, persisted in its own XML part (storage.js).
// `nextNumber` is a fast monotonic counter for inserts; "Renumber all" recomputes
// the authoritative sequence from document order, so drift from out-of-order
// inserts is self-healing. Default style: 'table' — the journal layout with the
// number on the right text border (the inline style trails the equation).
export const DEFAULT_NUMBERING = { style: 'table', nextNumber: 1 };

// The displayed number text. One place to change if we later offer [n] / n.
// formats — every render and every renumber goes through here.
export function formatNumber(n) {
  return `(${n})`;
}

// Where the number physically lives for a style — drives the insert and renumber
// dispatch. 'svg': baked into the equation picture (inline). 'cell': real Word
// text in a table cell (table). 'field': a native SEQ field (field).
export function numberPlacement(style) {
  if (style === 'table') return 'cell';
  if (style === 'field') return 'field';
  return 'svg';
}

// Render-time LaTeX for a numbered equation. Only the 'inline' style bakes the
// number into the math (\qquad gives a clear gap before it); 'table'/'field' add
// the number as document content after insert, so the math is left untouched.
// The RAW user LaTeX is what gets stored — this decoration is applied only on the
// way to MathJax, so click-to-edit shows clean source and renumber can re-apply a
// fresh number cleanly.
export function decorateLatexForNumber(latex, number, style = 'inline') {
  if (style === 'inline' && Number.isFinite(number) && number > 0) {
    return `${latex} \\qquad ${formatNumber(number)}`;
  }
  return latex;
}

// Is a style usable on this host? Only 'field' is gated: insertField is
// desktop-only, so on Word web the field style can't place new numbers.
export function isStyleSupported(style, caps = {}) {
  if (style === 'field') return !!caps.canInsertField;
  return NUMBERING_STYLES.includes(style);
}

// Coerce a (possibly persisted/stale/unsupported-on-this-host) style to a usable
// one. An unknown style or a 'field' style opened on Word web falls back to
// 'inline' so numbering never silently breaks.
export function normalizeStyle(style, caps = {}) {
  if (NUMBERING_STYLES.includes(style) && isStyleSupported(style, caps)) return style;
  return 'inline';
}

// Given equations in DOCUMENT order, assign 1..n to the numbered ones (skipping
// unnumbered equations) and return only the assignments that CHANGED. Renumber
// then has to touch just the equations whose displayed number actually moved.
//   ordered: [{ uuid, numbered, number }]  → [{ uuid, number }]
export function assignNumbers(ordered) {
  let n = 0;
  const changed = [];
  for (const eq of ordered) {
    if (!eq.numbered) continue;
    n += 1;
    if (eq.number !== n) changed.push({ uuid: eq.uuid, number: n });
  }
  return changed;
}

// The number to give a freshly-numbered equation, and the counter to persist
// next. Editing an already-numbered equation keeps its number (pass existingNumber);
// a brand-new number consumes the counter.
export function nextNumberFor(state, existingNumber = 0) {
  if (Number.isFinite(existingNumber) && existingNumber > 0) {
    return { number: existingNumber, nextNumber: state.nextNumber };
  }
  const number = state.nextNumber;
  return { number, nextNumber: number + 1 };
}

// ---- Style migration (WP2.2 increment 4) ----
// Where an equation's number physically lives NOW, from its stored payload: an
// unnumbered equation is always a bare picture, whatever style the document uses.
export function placementOf(eq) {
  return eq && eq.numbered ? numberPlacement(eq.numberStyle) : 'svg';
}

// The structural operation that takes an equation from its stored state to a
// target { numbered, style }. This is the single dispatch point for BOTH the
// edit path ("Number this" toggled on a table equation) and a document-wide
// style switch, so the two can never disagree about what has to happen:
//   'none'     — nothing structural (same placement, same numbered state)
//   'rerender' — bare picture stays bare, but the baked-in number changes
//                (inline number added/removed)
//   'recell'   — stays in its 1×2 table; only the number cell changes
//                (static text ↔ SEQ field)
//   'wrap'     — bare picture → 1×2 table (number becomes document content)
//   'unwrap'   — 1×2 table → bare picture (number baked into the SVG or gone)
export function migrationOp(before, after) {
  const from = placementOf(before);
  const to = after && after.numbered ? numberPlacement(after.style) : 'svg';
  if (from === 'svg' && to === 'svg') {
    return !!(before && before.numbered) === !!(after && after.numbered) ? 'none' : 'rerender';
  }
  if (from !== 'svg' && to !== 'svg') return from === to ? 'none' : 'recell';
  return from === 'svg' ? 'wrap' : 'unwrap';
}

// Plan a document-wide switch to `targetStyle`: every NUMBERED equation whose
// placement differs gets an op; unnumbered equations are never touched (a style
// governs where numbers live, not whether an equation has one). Order is
// preserved so the caller can convert in document order.
//   equations: [{ uuid, numbered, numberStyle, … }]  → [{ uuid, op }]
export function planStyleMigration(equations, targetStyle) {
  const plan = [];
  for (const eq of equations) {
    if (!eq.numbered) continue;
    const op = migrationOp(eq, { numbered: true, style: targetStyle });
    if (op !== 'none') plan.push({ uuid: eq.uuid, op });
  }
  return plan;
}
