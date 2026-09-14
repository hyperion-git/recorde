import { test } from 'node:test';
import assert from 'node:assert/strict';
import { DOMParser } from 'linkedom';

// mathsvg's SVG helpers call `new DOMParser()` only indirectly via the elements
// the caller builds; here we build those elements with linkedom and expose
// DOMParser as a global is unnecessary — we construct elements directly below.
const svgFrom = (s) => new DOMParser().parseFromString(s, 'image/svg+xml').documentElement;

import {
  mathJaxFontName, flattenLatex, computeRenderSize,
  descentPtFromViewBox, scaleSvgToPt, inlineColors, EX_PER_EM,
  isParagraphTextEmpty, findRenderError,
} from '../src/mathsvg.js';

test('mathJaxFontName maps short → mathjax-<short>', () => {
  assert.equal(mathJaxFontName('stix2'), 'mathjax-stix2');
  assert.equal(mathJaxFontName('tex'), 'mathjax-tex');
});

test('flattenLatex collapses any newline run to a single space', () => {
  assert.equal(flattenLatex('a\nb'), 'a b');
  assert.equal(flattenLatex('x\r\n\ny'), 'x y');
  assert.equal(flattenLatex('no newlines'), 'no newlines');
});

test('computeRenderSize: body mode returns the (clamped) body size', () => {
  assert.equal(computeRenderSize({ mode: 'body', bodyPt: 12 }), 12);
});

test('computeRenderSize: fixed mode honors a valid fixedPt', () => {
  assert.equal(computeRenderSize({ mode: 'fixed', fixedPt: 14, bodyPt: 11 }), 14);
});

test('computeRenderSize: a corrupt fixedPt (500pt) falls back to body', () => {
  // WHY: a sizePt loaded from an older/foreign XML part must not produce a
  // 500pt equation; out-of-range inputs degrade to the body size.
  assert.equal(computeRenderSize({ mode: 'fixed', fixedPt: 500, bodyPt: 12 }), 12);
});

test('computeRenderSize: selection mode uses selectionPt when valid', () => {
  assert.equal(computeRenderSize({ mode: 'selection', selectionPt: 20, bodyPt: 11 }), 20);
});

test('computeRenderSize: ambiguous (null) selection falls back to body', () => {
  assert.equal(computeRenderSize({ mode: 'selection', selectionPt: null, bodyPt: 11 }), 11);
});

test('computeRenderSize: corrupt body defaults to 11', () => {
  assert.equal(computeRenderSize({ mode: 'body', bodyPt: NaN }), 11);
});

test('descentPtFromViewBox: real descenders yield positive pt', () => {
  // vy=-800, vh=1000 → descent 200 internal units → 200*10/1000 = 2pt
  assert.equal(descentPtFromViewBox('0 -800 1200 1000', 10), 2);
});

test('descentPtFromViewBox: above-baseline-only content is ~0', () => {
  // WHY: equations like 1+2+3 sit on the baseline; descent must be 0 (no shift).
  assert.equal(descentPtFromViewBox('0 -1000 1200 1000', 10), 0);
});

test('descentPtFromViewBox: missing/malformed viewBox → 0', () => {
  assert.equal(descentPtFromViewBox('', 10), 0);
  assert.equal(descentPtFromViewBox('1 2 3', 10), 0);
  assert.equal(descentPtFromViewBox(null, 10), 0);
});

test('scaleSvgToPt derives width/height from the viewBox (1000u = 1em)', () => {
  const svg = svgFrom('<svg viewBox="0 0 2000 1000"></svg>');
  scaleSvgToPt(svg, 10);
  assert.equal(svg.getAttribute('width'), '20.000pt');
  assert.equal(svg.getAttribute('height'), '10.000pt');
});

test('scaleSvgToPt falls back to ex attributes when viewBox is absent', () => {
  const svg = svgFrom('<svg width="2ex" height="1ex"></svg>');
  scaleSvgToPt(svg, 10);
  assert.ok(svg.getAttribute('width').endsWith('pt'));
  assert.ok(Math.abs(parseFloat(svg.getAttribute('width')) - 2 * EX_PER_EM * 10) < 1e-3);
});

test('inlineColors replaces currentColor with black by default', () => {
  const svg = svgFrom('<svg><path fill="currentColor"></path><rect stroke="currentColor"></rect></svg>');
  inlineColors(svg);
  assert.equal(svg.getAttribute('color'), '#000000');
  assert.equal(svg.querySelector('path').getAttribute('fill'), '#000000');
  assert.equal(svg.querySelector('rect').getAttribute('stroke'), '#000000');
});

test('inlineColors honors a custom color', () => {
  const svg = svgFrom('<svg><path fill="currentColor"></path></svg>');
  inlineColors(svg, '#ff0000');
  assert.equal(svg.querySelector('path').getAttribute('fill'), '#ff0000');
});

test('isParagraphTextEmpty: null/undefined/empty count as empty', () => {
  // WHY: Word reports an equation-only paragraph's text as empty/whitespace once
  // the inline picture is excluded — that's the safe-to-center signal.
  assert.equal(isParagraphTextEmpty(''), true);
  assert.equal(isParagraphTextEmpty(null), true);
  assert.equal(isParagraphTextEmpty(undefined), true);
});

test('isParagraphTextEmpty: plain whitespace (space/tab/newline) is empty', () => {
  assert.equal(isParagraphTextEmpty('   '), true);
  assert.equal(isParagraphTextEmpty('\t\n'), true);
});

test('isParagraphTextEmpty: invisible Word artifacts are empty', () => {
  // WHY: NBSP/zero-width/BOM/vertical-tab survive a bare /\s/ strip in some
  // engines and would falsely block centering of an equation-only paragraph.
  assert.equal(isParagraphTextEmpty('\u00A0'), true);
  assert.equal(isParagraphTextEmpty('\u200B\uFEFF\u000B'), true);
});

test('isParagraphTextEmpty: any visible character makes it non-empty', () => {
  // WHY: real prose must never be centered; a single glyph (even amid invisibles)
  // is enough to keep the paragraph's alignment untouched.
  assert.equal(isParagraphTextEmpty('x'), false);
  assert.equal(isParagraphTextEmpty('  word  '), false);
  assert.equal(isParagraphTextEmpty('a\u00A0b'), false);
});

test('findRenderError: reads the message MathJax stores on an merror node', () => {
  // WHY: MathJax never throws on bad TeX — it draws an error box. The pane must
  // read that message to show it under the source instead of a silent red box.
  const bad = svgFrom('<svg xmlns="http://www.w3.org/2000/svg"><g data-mjx-error="Missing close brace"><rect/></g></svg>');
  assert.equal(findRenderError(bad), 'Missing close brace');
  const ok = svgFrom('<svg xmlns="http://www.w3.org/2000/svg"><g><rect/></g></svg>');
  assert.equal(findRenderError(ok), null);
  assert.equal(findRenderError(null), null);
});
