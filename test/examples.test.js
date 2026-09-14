import { test } from 'node:test';
import assert from 'node:assert/strict';

import { EXAMPLES, findExample } from '../src/examples.js';
import { DEFAULT_PREAMBLE } from '../src/preamble.js';
import { flattenLatex } from '../src/mathsvg.js';
import { initMathJaxLikePane, renderError } from '../scripts/mathjax-node.mjs';

test('examples: unique ids, valid modes, non-empty sources', () => {
  const ids = EXAMPLES.map((e) => e.id);
  assert.equal(new Set(ids).size, ids.length, 'duplicate example id');
  for (const e of EXAMPLES) {
    assert.ok(['inline', 'display'].includes(e.mode), `${e.id}: bad mode`);
    assert.ok(e.latex.trim().length > 0, `${e.id}: empty latex`);
    assert.ok(e.name.length > 0, `${e.id}: empty name`);
    if (e.numbered) assert.equal(e.mode, 'display', `${e.id}: only display equations can be numbered`);
  }
  assert.equal(findExample('euler').mode, 'inline');
  assert.equal(findExample('nope'), null);
});

test('examples: every example renders through MathJax with the default preamble (no merror)', async () => {
  // WHY: the picker is the user's first impression AND the manual smoke-test set;
  // an example that renders an error box would look like a broken add-in. The
  // 'macros' example additionally proves the default preamble registers.
  const MathJax = await initMathJaxLikePane();
  await MathJax.tex2svgPromise(DEFAULT_PREAMBLE);
  for (const e of EXAMPLES) {
    const err = await renderError(MathJax, flattenLatex(e.latex), e.mode === 'display');
    assert.equal(err, null, `${e.id}: ${err}`);
  }
});
