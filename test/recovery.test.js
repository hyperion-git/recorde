import { test } from 'node:test';
import assert from 'node:assert/strict';
import { DOMParser } from 'linkedom';

import { looksLikeEquation, buildDegradedPayload } from '../src/recovery.js';

test('looksLikeEquation: LaTeX markers always pass', () => {
  assert.equal(looksLikeEquation('\\frac{1}{2}'), true);
  assert.equal(looksLikeEquation('x_i^2'), true);
});

test('looksLikeEquation: operator-bearing non-prose passes', () => {
  assert.equal(looksLikeEquation('1+2=3'), true);
});

test('looksLikeEquation: empty / whitespace-only fails', () => {
  assert.equal(looksLikeEquation(''), false);
  assert.equal(looksLikeEquation('   '), false);
});

test('looksLikeEquation: prose with operators is rejected', () => {
  // WHY: a caption like this must never be re-rendered as an equation — doing so
  // would replace visible text with a broken render.
  assert.equal(looksLikeEquation('Figure 3 our results'), false);
});

test('looksLikeEquation: non-strings fail', () => {
  assert.equal(looksLikeEquation(null), false);
  assert.equal(looksLikeEquation(undefined), false);
  assert.equal(looksLikeEquation(42), false);
});

test('looksLikeEquation: over-long input fails', () => {
  assert.equal(looksLikeEquation('a='.repeat(3000)), false);
});

test('buildDegradedPayload: trims latex and fills defaults', () => {
  assert.deepEqual(
    buildDegradedPayload({ uuid: 'u', latex: ' x ', sizePt: 13 }),
    { uuid: 'u', latex: 'x', font: 'termes', sizePt: 13, displayMode: 'inline', degraded: true },
  );
});

test('buildDegradedPayload: out-of-range / missing sizePt → 11', () => {
  assert.equal(buildDegradedPayload({ uuid: 'u', latex: 'x', sizePt: 500 }).sizePt, 11);
  assert.equal(buildDegradedPayload({ uuid: 'u', latex: 'x', sizePt: undefined }).sizePt, 11);
});

test('degraded payload round-trips through storage, dropping the flag', async () => {
  // WHY: the `degraded` key is harmless metadata for the caller; it must not leak
  // into the persisted XML part. buildXml only serializes the known fields.
  globalThis.DOMParser = (await import('linkedom')).DOMParser;
  const { buildXml, parseXml } = await import('../src/storage.js');

  const xml = buildXml(buildDegradedPayload({ uuid: 'u9', latex: '\\alpha=\\beta', sizePt: 9 }));
  assert.ok(!xml.includes('degraded'), 'degraded flag must not reach storage');

  const d = parseXml(xml);
  assert.equal(d.font, 'termes');
  assert.equal(d.displayMode, 'inline');
  assert.equal(d.sizePt, 9);
  assert.equal(d.latex, '\\alpha=\\beta');
});
