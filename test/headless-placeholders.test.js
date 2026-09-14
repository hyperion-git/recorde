import { test } from 'node:test';
import assert from 'node:assert/strict';
import { findPlaceholders, validatePlaceholders } from '../headless/lib/placeholders.mjs';

test('finds every kind with label and trimmed LaTeX', () => {
  // WHY: this is the agent contract. A kind or label that fails to parse means
  // the agent's equation silently stays as literal text in the delivered file.
  const t = 'a [[math: x^2 ]] b [[display: \\int f]] c [[eq#lbl: y=1]] d [[ref: lbl]] e';
  const p = findPlaceholders(t);
  assert.deepEqual(p.map((x) => [x.kind, x.label, x.latex]), [
    ['math', null, 'x^2'], ['display', null, '\\int f'], ['eq', 'lbl', 'y=1'], ['ref', null, 'lbl'],
  ]);
  assert.equal(t.slice(p[0].start, p[0].end), '[[math: x^2 ]]');
});

test('closing brackets inside LaTeX do not end the placeholder early', () => {
  // WHY: intervals and matrices end in "]" all the time. The (?!\]) lookahead
  // makes the match stop at the first "]]" that is not followed by another "]".
  const p = findPlaceholders('[[math: f(x) = [a,b]]] rest');
  assert.equal(p.length, 1);
  assert.equal(p[0].latex, 'f(x) = [a,b]');
});

test('a placeholder never spans a U+FFFC object marker', () => {
  // WHY: U+FFFC stands for a tab/break/drawing between runs; matching across
  // it would splice out document content that is not part of the placeholder.
  assert.equal(findPlaceholders('[[math: a ￼ b]]').length, 0);
});

test('dollar syntax is opt-in and skips currency-like text', () => {
  const t = 'costs $5 and $10 but $x^2$ and $$\\sum$$';
  assert.equal(findPlaceholders(t).length, 0);
  const p = findPlaceholders(t, { dollar: true });
  assert.deepEqual(p.map((x) => [x.kind, x.latex]), [['math', 'x^2'], ['display', '\\sum']]);
});

test('validatePlaceholders rejects empty math and malformed refs', () => {
  assert.match(validatePlaceholders(findPlaceholders('[[math:   ]]')), /empty math/);
  assert.match(validatePlaceholders(findPlaceholders('[[ref#x: y]]')), /ref/);
  assert.equal(validatePlaceholders(findPlaceholders('[[ref: y]] [[eq: z]]')), null);
});

test('per-equation options: parsed, validated, and refused on refs', () => {
  // WHY: a paper mixes fonts/colours/styles per equation (the showcase paper
  // does); options let the headless route honour that without a second pass.
  const [p] = findPlaceholders('[[eq#a{font=stix2,color=#1f77b4,size=12,style=field}: x]]');
  assert.deepEqual([p.kind, p.label, p.latex], ['eq', 'a', 'x']);
  assert.deepEqual(p.options, { font: 'stix2', color: '#1f77b4', size: '12', style: 'field' });
  assert.equal(validatePlaceholders([p]), null);
  assert.match(validatePlaceholders(findPlaceholders('[[math{style=field}: x]]')), /numbered/);
  assert.match(validatePlaceholders(findPlaceholders('[[math{colour=red}: x]]')), /unknown option/);
  assert.match(validatePlaceholders(findPlaceholders('[[display{color=red}: x]]')), /#rrggbb/);
  assert.match(validatePlaceholders(findPlaceholders('[[ref{font=tex}: a]]')), /no options/);
  assert.deepEqual(findPlaceholders('[[math: {a}]]')[0].options, {});
});
