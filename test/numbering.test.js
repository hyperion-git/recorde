import { test } from 'node:test';
import assert from 'node:assert/strict';

import {
  formatNumber, decorateLatexForNumber, isStyleSupported, normalizeStyle,
  assignNumbers, nextNumberFor, numberPlacement, DEFAULT_NUMBERING, NUMBERING_STYLES,
} from '../src/numbering.js';

test('decorateLatexForNumber bakes the number into inline-style LaTeX', () => {
  // WHY: for the inline style the number is rendered INTO the SVG. If the
  // decoration doesn't reach MathJax, a "numbered" equation renders with no
  // number at all — the feature silently does nothing.
  const out = decorateLatexForNumber('E = mc^2', 7, 'inline');
  assert.ok(out.includes('(7)'), 'expected the (7) tag in the render LaTeX');
  assert.ok(out.startsWith('E = mc^2'), 'original equation must be preserved');
});

test('decorateLatexForNumber leaves the equation untouched for table/field', () => {
  // WHY: table/field styles add the number as DOCUMENT content (a cell / a SEQ
  // field). Baking it into the SVG too would show the number twice.
  assert.equal(decorateLatexForNumber('E = mc^2', 7, 'table'), 'E = mc^2');
  assert.equal(decorateLatexForNumber('E = mc^2', 7, 'field'), 'E = mc^2');
});

test('decorateLatexForNumber never emits a number for unnumbered equations', () => {
  // WHY: number 0 means "not numbered". A stray "(0)" must never appear.
  assert.equal(decorateLatexForNumber('x', 0, 'inline'), 'x');
  assert.equal(decorateLatexForNumber('x', -1, 'inline'), 'x');
  assert.equal(decorateLatexForNumber('x', NaN, 'inline'), 'x');
});

test('decoration is applied to RAW source, so numbers cannot compound', () => {
  // WHY: storage keeps the raw LaTeX; renumber re-decorates the raw source with a
  // fresh number. Decorating from raw each time guarantees one number, never the
  // "(3) (4) (5)" pile-up that decorating an already-decorated string would give.
  const raw = '\\alpha + \\beta';
  const first = decorateLatexForNumber(raw, 3, 'inline');
  const renumbered = decorateLatexForNumber(raw, 9, 'inline');
  assert.ok(first.includes('(3)') && !first.includes('(9)'));
  assert.ok(renumbered.includes('(9)') && !renumbered.includes('(3)'));
});

test('formatNumber is the single place numbers are formatted', () => {
  // WHY: every render and renumber routes display text through here, so changing
  // the format later (e.g. [n]) is a one-line change with no drift.
  assert.equal(formatNumber(12), '(12)');
});

test('isStyleSupported gates only the field style, and only on capable hosts', () => {
  // WHY: insertField is desktop-only. On Word web (no canInsertField) the field
  // style must report unsupported so we never try to insert a field that fails.
  assert.equal(isStyleSupported('inline', {}), true);
  assert.equal(isStyleSupported('table', {}), true);
  assert.equal(isStyleSupported('field', {}), false);
  assert.equal(isStyleSupported('field', { canInsertField: true }), true);
});

test('normalizeStyle degrades an unusable style to inline, never breaks', () => {
  // WHY: a 'field' document reopened on Word web must keep working (just without
  // native fields), and a corrupt/unknown style must not wedge numbering.
  assert.equal(normalizeStyle('field', {}), 'inline');           // web: no fields
  assert.equal(normalizeStyle('field', { canInsertField: true }), 'field');
  assert.equal(normalizeStyle('bogus', {}), 'inline');
  assert.equal(normalizeStyle('table', {}), 'table');
});

test('assignNumbers gives contiguous 1..n to numbered equations only', () => {
  // WHY: unnumbered equations interleaved with numbered ones must not consume a
  // number — the visible sequence has to read 1,2,3 with no gaps.
  const changed = assignNumbers([
    { uuid: 'a', numbered: true, number: 0 },
    { uuid: 'b', numbered: false, number: 0 },
    { uuid: 'c', numbered: true, number: 0 },
  ]);
  assert.deepEqual(changed, [{ uuid: 'a', number: 1 }, { uuid: 'c', number: 2 }]);
});

test('assignNumbers returns ONLY equations whose number changed', () => {
  // WHY: each returned equation triggers a costly SVG re-render + in-place
  // replace in Word. An already-correct equation must not be touched.
  const changed = assignNumbers([
    { uuid: 'a', numbered: true, number: 1 },   // already correct
    { uuid: 'b', numbered: true, number: 5 },   // should become 2
  ]);
  assert.deepEqual(changed, [{ uuid: 'b', number: 2 }]);
});

test('assignNumbers on an already-ordered set changes nothing', () => {
  const changed = assignNumbers([
    { uuid: 'a', numbered: true, number: 1 },
    { uuid: 'b', numbered: true, number: 2 },
  ]);
  assert.deepEqual(changed, []);
});

test('nextNumberFor consumes the counter for a new number', () => {
  const r = nextNumberFor({ style: 'inline', nextNumber: 4 }, 0);
  assert.equal(r.number, 4);
  assert.equal(r.nextNumber, 5);
});

test('nextNumberFor keeps an existing number and leaves the counter alone', () => {
  // WHY: editing a numbered equation must reuse its number. Consuming a fresh one
  // would skip a value and leave a permanent gap in the sequence.
  const r = nextNumberFor({ style: 'inline', nextNumber: 9 }, 3);
  assert.equal(r.number, 3);
  assert.equal(r.nextNumber, 9);
});

test('DEFAULT_NUMBERING and NUMBERING_STYLES are coherent; the default puts numbers on the right border', () => {
  // WHY: the user expects equation numbers at the right text border by default;
  // only the table style does that (inline trails the equation, field is
  // desktop-only).
  assert.ok(NUMBERING_STYLES.includes(DEFAULT_NUMBERING.style));
  assert.equal(DEFAULT_NUMBERING.style, 'table');
  assert.equal(DEFAULT_NUMBERING.nextNumber, 1);
});

test('numberPlacement maps each style to where its number lives', () => {
  // WHY: this drives the insert/renumber DISPATCH. 'svg' re-renders the picture;
  // 'cell' rewrites a table cell; 'field' is a native field. A wrong mapping would
  // send a table equation down the picture-rerender path (and never touch its
  // number cell), so renumber would silently no-op.
  assert.equal(numberPlacement('inline'), 'svg');
  assert.equal(numberPlacement('table'), 'cell');
  assert.equal(numberPlacement('field'), 'field');
  assert.equal(numberPlacement('bogus'), 'svg');   // safe default
});

test('table/field equations are NOT decorated even when decorate is asked', () => {
  // WHY: for table/field the number is document content (cell text / SEQ field).
  // If decorate baked it into the SVG too, the number would appear twice. This is
  // the invariant that lets the same render path serve all three styles.
  assert.equal(decorateLatexForNumber('x=1', 5, 'table'), 'x=1');
  assert.equal(decorateLatexForNumber('x=1', 5, 'field'), 'x=1');
  assert.ok(decorateLatexForNumber('x=1', 5, 'inline').includes('(5)'));
});

// ---- Style migration (increment 4) ----
import { placementOf, migrationOp, planStyleMigration } from '../src/numbering.js';

test('placementOf: an unnumbered equation is a bare picture whatever its style says', () => {
  // WHY: numberStyle is stored per equation even when numbered=false (it just
  // records the doc style at the time). Migration must not try to unwrap a
  // table that was never created.
  assert.equal(placementOf({ numbered: false, numberStyle: 'table' }), 'svg');
  assert.equal(placementOf({ numbered: true, numberStyle: 'table' }), 'cell');
  assert.equal(placementOf({ numbered: true, numberStyle: 'field' }), 'field');
  assert.equal(placementOf({ numbered: true, numberStyle: 'inline' }), 'svg');
  assert.equal(placementOf(null), 'svg');
});

test('migrationOp: the full before→after matrix', () => {
  // WHY: this is the one dispatch the edit path and the style switch both rely
  // on. A wrong cell here either leaves a stray 1×2 table behind, or tries to
  // delete a table that doesn't exist and takes user content with it.
  const bare = { numbered: false, numberStyle: 'inline' };
  const inl = { numbered: true, numberStyle: 'inline' };
  const tab = { numbered: true, numberStyle: 'table' };
  const fld = { numbered: true, numberStyle: 'field' };
  const to = (numbered, style) => ({ numbered, style });

  // bare picture ↔ bare picture: only the baked-in number changes
  assert.equal(migrationOp(bare, to(false, 'table')), 'none');
  assert.equal(migrationOp(bare, to(true, 'inline')), 'rerender');
  assert.equal(migrationOp(inl, to(false, 'inline')), 'rerender');
  assert.equal(migrationOp(inl, to(true, 'inline')), 'none');
  // into a table
  assert.equal(migrationOp(bare, to(true, 'table')), 'wrap');
  assert.equal(migrationOp(inl, to(true, 'field')), 'wrap');
  // out of a table
  assert.equal(migrationOp(tab, to(false, 'table')), 'unwrap');
  assert.equal(migrationOp(tab, to(true, 'inline')), 'unwrap');
  assert.equal(migrationOp(fld, to(false, 'inline')), 'unwrap');
  // table ↔ field: the picture stays put, only the number cell is rewritten
  assert.equal(migrationOp(tab, to(true, 'field')), 'recell');
  assert.equal(migrationOp(fld, to(true, 'table')), 'recell');
  assert.equal(migrationOp(tab, to(true, 'table')), 'none');
  // a missing stored payload (recovered picture) counts as a bare picture
  assert.equal(migrationOp(null, to(true, 'table')), 'wrap');
});

test('planStyleMigration converts only numbered equations that differ, in order', () => {
  // WHY: switching the document style must not number equations the user left
  // unnumbered, must skip equations already in the target style, and must keep
  // document order so the follow-up renumber sees a stable sequence.
  const eqs = [
    { uuid: 'a', numbered: true, numberStyle: 'inline' },
    { uuid: 'b', numbered: false, numberStyle: 'inline' },
    { uuid: 'c', numbered: true, numberStyle: 'table' },
    { uuid: 'd', numbered: true, numberStyle: 'field' },
  ];
  assert.deepEqual(planStyleMigration(eqs, 'table'), [
    { uuid: 'a', op: 'wrap' },
    { uuid: 'd', op: 'recell' },
  ]);
  assert.deepEqual(planStyleMigration(eqs, 'inline'), [
    { uuid: 'c', op: 'unwrap' },
    { uuid: 'd', op: 'unwrap' },
  ]);
  assert.deepEqual(planStyleMigration([], 'field'), []);
});
