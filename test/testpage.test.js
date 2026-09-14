import { test } from 'node:test';
import assert from 'node:assert/strict';

import { TEST_PAGE, testPageEquations } from '../src/testpage.js';
import { KNOWN_FONTS } from '../src/settings.js';
import { NUMBERING_STYLES } from '../src/numbering.js';
import { DEFAULT_PREAMBLE } from '../src/preamble.js';
import { flattenLatex } from '../src/mathsvg.js';
import { initMathJaxLikePane, renderError } from '../scripts/mathjax-node.mjs';

test('test page: well-formed blocks with settings the pane actually offers', () => {
  // WHY: the runner feeds these straight into performInsert; an unknown font or
  // style would silently fall back and the "test" would test the wrong thing.
  for (const b of TEST_PAGE) {
    const kinds = ['h1', 'h2', 'p', 'eq'].filter((k) => k in b);
    assert.equal(kinds.length, 1, `block must have exactly one of h1/h2/p/eq: ${JSON.stringify(b).slice(0, 60)}`);
    if (!b.eq) continue;
    assert.ok(b.eq.latex.trim(), 'empty latex');
    if (b.eq.font) assert.ok(KNOWN_FONTS.includes(b.eq.font), `unknown font ${b.eq.font}`);
    if (b.eq.displayMode) assert.ok(['inline', 'display'].includes(b.eq.displayMode));
    if (b.eq.style) assert.ok(NUMBERING_STYLES.includes(b.eq.style), `unknown style ${b.eq.style}`);
    if (b.eq.numbered) assert.equal(b.eq.displayMode, 'display', 'only display equations are numbered');
    if (b.eq.displayMode === 'display') assert.ok(!b.before && !b.after, 'display equations get their own paragraph');
  }
  // Every numbering style and most fonts are covered.
  const eqs = testPageEquations();
  for (const st of NUMBERING_STYLES) assert.ok(eqs.some((e) => e.numbered && e.style === st), `style ${st} not exercised`);
  assert.ok(new Set(eqs.map((e) => e.font)).size >= 5, 'at least five fonts exercised');
});

test('test page: every equation renders with the default preamble (no merror)', async () => {
  const MathJax = await initMathJaxLikePane();
  await MathJax.tex2svgPromise(DEFAULT_PREAMBLE);
  for (const [i, e] of testPageEquations().entries()) {
    const err = await renderError(MathJax, flattenLatex(e.latex), e.displayMode === 'display');
    assert.equal(err, null, `equation #${i + 1}: ${err}`);
  }
});
