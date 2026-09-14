import { test } from 'node:test';
import assert from 'node:assert/strict';

import { TEST_PAPER, testPaperEquations, paperOutline } from '../src/testpaper.js';
import { KNOWN_FONTS } from '../src/settings.js';
import { NUMBERING_STYLES } from '../src/numbering.js';
import { DEFAULT_PREAMBLE } from '../src/preamble.js';
import { flattenLatex } from '../src/mathsvg.js';
import { initMathJaxLikePane, renderError } from '../scripts/mathjax-node.mjs';
import { findPlaceholders, validatePlaceholders } from '../headless/lib/placeholders.mjs';

test('test paper: well-formed blocks; Times body; Termes/table defaults; deviations only in showcase blocks', () => {
  // WHY: the paper mirrors RevTeX typography (Times text, newtxmath-like
  // Termes, right-margin numbers). Section IV deliberately exercises every
  // other option — but only blocks flagged `showcase` may deviate, so a stray
  // font or style elsewhere is caught as a mistake rather than a feature.
  let eqCount = 0, ooxml = 0;
  const fonts = new Set(), styles = new Set(), sizes = new Set(), colored = [];
  let multiPart = 0;
  for (const b of TEST_PAPER) {
    const kinds = ['text', 'eq', 'ooxml', 'parts'].filter((k) => k in b);
    assert.equal(kinds.length, 1, `block needs exactly one of text/eq/ooxml/parts: ${JSON.stringify(b).slice(0, 60)}`);
    if (b.ooxml) { ooxml++; assert.match(b.ooxml, new RegExp(`w:cols w:num="${b.cols}"`)); continue; }
    if (b.fmt) assert.equal(b.fmt.font, 'Times New Roman');
    const eqs = b.eq ? [b.eq] : b.parts ? b.parts.filter((x) => typeof x !== 'string') : [];
    if (b.parts) {
      assert.ok(b.parts.every((x, i) => (typeof x === 'string') === (i % 2 === 0)), 'parts alternate text/equation');
      if (eqs.length > 1) multiPart++;
    }
    for (const e of eqs) {
      eqCount++;
      assert.ok(KNOWN_FONTS.includes(e.font));
      assert.equal(e.sizeMode, 'fixed');
      fonts.add(e.font); sizes.add(e.fixedPt);
      if (e.color) colored.push(e);
      if (e.numbered) { assert.equal(e.displayMode, 'display'); assert.ok(NUMBERING_STYLES.includes(e.style)); styles.add(e.style); }
      const deviates = e.font !== 'termes' || e.fixedPt !== 10 || !!e.color || (e.numbered && e.style !== 'table') || !!e.align;
      if (deviates) assert.ok(b.showcase, `non-default equation outside a showcase block: ${e.latex.slice(0, 40)}`);
    }
  }
  assert.ok(eqCount >= 50, `enough equations to look like a paper (${eqCount})`);
  assert.ok(multiPart >= 6, 'several paragraphs carry more than one inline equation');
  assert.deepEqual([...fonts].sort(), [...KNOWN_FONTS].sort(), 'every bundled font appears');
  assert.deepEqual([...styles].sort(), [...NUMBERING_STYLES].sort(), 'every numbering style appears');
  assert.ok(sizes.size >= 4, 'several fixed sizes');
  assert.ok(colored.length >= 1, 'a coloured equation');
  const aligns = new Set(testPaperEquations().map((e) => e.align).filter(Boolean));
  assert.deepEqual([...aligns].sort(), ['center', 'left'], 'both display alignments are forced somewhere');
  assert.equal(ooxml, 2, 'one single-column and one two-column section break');
});

test('test paper: every equation renders with the default preamble (no merror)', async () => {
  const MathJax = await initMathJaxLikePane();
  await MathJax.tex2svgPromise(DEFAULT_PREAMBLE);
  for (const [i, e] of testPaperEquations().entries()) {
    const err = await renderError(MathJax, flattenLatex(e.latex), e.displayMode === 'display');
    assert.equal(err, null, `equation #${i + 1} (${e.latex.slice(0, 40)}): ${err}`);
  }
});

test('paperOutline: the headless twin carries every equation with matching options', () => {
  // WHY: both routes must produce the same paper. Every add-in equation must
  // appear once as a valid placeholder, with its deviations encoded as options.
  const outline = paperOutline();
  const eqs = testPaperEquations();
  const found = outline.filter((b) => b.kind === 'p').flatMap((b) => findPlaceholders(b.text));
  assert.equal(found.length, eqs.length);
  assert.equal(outline.filter((b) => b.kind === 'section').length, 2);
  for (const b of outline) if (b.kind === 'p') assert.equal(validatePlaceholders(findPlaceholders(b.text)), null);
  found.forEach((p, i) => {
    const e = eqs[i];
    assert.equal(p.latex, e.latex);
    assert.equal(p.kind, e.displayMode === 'inline' ? 'math' : e.numbered ? 'eq' : 'display');
    assert.equal(p.options.font ?? 'termes', e.font);
    assert.equal(p.options.color ?? null, e.color ?? null);
    assert.equal(Number(p.options.size ?? 10), e.fixedPt);
    if (e.numbered) assert.equal(p.options.style ?? 'table', e.style);
    assert.equal(p.options.align ?? null, e.align ?? null);
  });
});
