import { test } from 'node:test';
import assert from 'node:assert/strict';

import { SNIPPET_CURSOR, applySnippet, PALETTE } from '../src/snippets.js';

test('applySnippet: plain symbol inserts at the caret, caret follows it', () => {
  const r = applySnippet('ab', 1, 1, '\\pi');
  assert.equal(r.value, 'a\\pib');
  assert.equal(r.caret, 4);   // after the 3-char insert at index 1
});

test('applySnippet: a template with no selection drops the caret at the marker', () => {
  const r = applySnippet('', 0, 0, '\\frac{▮}{}');
  assert.equal(r.value, '\\frac{}{}');
  assert.equal(r.caret, 6);   // between the first braces (after "\frac{")
});

test('applySnippet: a selection is WRAPPED at the marker, caret after it', () => {
  // WHY: select an expression + click √ → \sqrt{expr}, the palette's power move.
  const r = applySnippet('xy', 0, 1, '\\sqrt{▮}');
  assert.equal(r.value, '\\sqrt{x}y');
  assert.equal(r.caret, 7);   // after the wrapped "x"
});

test('applySnippet: a selection with a markerless snippet is replaced', () => {
  const r = applySnippet('xy', 0, 1, '\\pi');
  assert.equal(r.value, '\\piy');
  assert.equal(r.caret, 3);
});

test('applySnippet: the marker never survives into the output', () => {
  const r = applySnippet('', 0, 0, '\\frac{▮}{}');
  assert.ok(!r.value.includes(SNIPPET_CURSOR));
});

test('applySnippet: out-of-range / negative indices are clamped, never throw', () => {
  assert.deepEqual(applySnippet('ab', 99, 99, 'X'), { value: 'abX', caret: 3 });
  assert.deepEqual(applySnippet('ab', -5, -5, 'X'), { value: 'Xab', caret: 1 });
});

test('PALETTE is well-formed: named categories, labelled items, ≤1 marker each', () => {
  assert.ok(Array.isArray(PALETTE) && PALETTE.length > 0);
  for (const cat of PALETTE) {
    assert.ok(typeof cat.name === 'string' && cat.name.length > 0);
    assert.ok(Array.isArray(cat.items) && cat.items.length > 0);
    for (const it of cat.items) {
      assert.ok(typeof it.label === 'string' && it.label.length > 0, `label for ${it.snippet}`);
      assert.ok(typeof it.snippet === 'string' && it.snippet.length > 0, `snippet for ${it.label}`);
      // At most one caret marker per snippet (applySnippet only honors the first).
      assert.ok(it.snippet.split(SNIPPET_CURSOR).length <= 2, `multiple markers in ${it.label}`);
    }
  }
});
