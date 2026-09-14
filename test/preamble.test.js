import { test } from 'node:test';
import assert from 'node:assert/strict';

import { DEFAULT_PREAMBLE } from '../src/preamble.js';
import { initMathJaxLikePane, renderError } from '../scripts/mathjax-node.mjs';

// Commands MathJax cannot handle — they hang, error, or silently misrender. The
// preamble is applied once at startup, so a single bad command would break ALL
// rendering. This guard keeps the ported subset honest. (Headless rendering of
// every macro is the other half — done via the Playwright battery, not here.)
const FORBIDDEN = [
  '\\NewDocumentCommand', '\\DeclareDocumentCommand', '\\DeclareFontFamily',
  '\\DeclareMathSymbol', '\\DeclareSymbolFont', '\\stackunder', '\\stackon',
  '\\bBigg@', '\\ooalign', '\\scalebox', '\\makebox', '\\sfrac', '\\vv{',
  '\\mathpalette', '\\arrowfill@', '\\DOTSI', '\\sffamily',
];

test('DEFAULT_PREAMBLE is a non-empty string', () => {
  assert.equal(typeof DEFAULT_PREAMBLE, 'string');
  assert.ok(DEFAULT_PREAMBLE.length > 500);
});

test('DEFAULT_PREAMBLE contains no MathJax-incompatible commands', () => {
  for (const cmd of FORBIDDEN) {
    assert.ok(!DEFAULT_PREAMBLE.includes(cmd), `preamble must not use ${cmd}`);
  }
});

test('DEFAULT_PREAMBLE defines the expected macro kinds', () => {
  // Sanity that it is the real ported content, not an empty stub.
  assert.ok(DEFAULT_PREAMBLE.includes('\\newcommand{\\QmOp}'));
  assert.ok(DEFAULT_PREAMBLE.includes('\\DeclarePairedDelimiter\\Avg'));
  assert.ok(DEFAULT_PREAMBLE.includes('\\DeclareMathOperator{\\Real}'));
  assert.ok(DEFAULT_PREAMBLE.includes('\\newcommand{\\Vect}'));
});

test('DEFAULT_PREAMBLE registers in MathJax and every ported macro + AFP colour renders', async () => {
  // WHY: the preamble is applied once per MathJax load; a single bad definition
  // aborts the whole registration and every macro-using equation then errors.
  const MathJax = await initMathJaxLikePane();
  const pre = await renderError(MathJax, DEFAULT_PREAMBLE, false);
  assert.equal(pre, null, `preamble registration: ${pre}`);
  const samples = [
    '\\SfOpI{T}\\SfOpBI{S}\\OperandPlaceholder\\OpPlaceholder',
    '\\TensorMatrixOverarrow{T}\\barbelow{x}\\SecondOrderTensor{M}\\DeltaTransversal\\DeltaLongitudinal',
    '\\InnerBracketRightNested{a}\\InnerBracketLeftNested{b}',
    '\\Arctan x\\Abb f\\CpVPVCal\\SumInt_{k} f_k',
    '\\Fvect{v}\\eql{def}\\eql[x]{y}\\overrightharpoon{a}\\overleftharpoon{b}',
    '\\PhaseSpaceArgs\\PhaseSpaceArgsVec\\intRlap_0^1 f\\vast(x\\vast)\\Vast[y\\Vast]',
    '\\textcolor{afp-c0}{x}+\\textcolor{afp-red-3}{y}+\\textcolor{afp-gray-medium}{z}+\\textcolor{afp-blue}{w}',
    '\\QmOpHamiltonian\\Avg*{\\HOp{H}}\\OpComm{\\QmOpX,\\QmOpP}\\VNabla\\cdot\\Vect{E}',
  ];
  for (const s of samples) {
    const err = await renderError(MathJax, s, true);
    assert.equal(err, null, `${s}: ${err}`);
  }
});

// ---- Acceptance against the author's paper template ----
import { readFileSync } from 'node:fs';
import { flattenLatex } from '../src/mathsvg.js';

test('every math snippet of the paper template renders with DEFAULT_PREAMBLE (no error, no undefined macro)', async () => {
  // WHY: the preamble exists so the author's manuscripts paste into Word
  // unchanged. The template's main document is the reference usage of every
  // macro; a snippet that errors or shows a red undefined macro is a port gap.
  // Refresh the fixture with scripts/extract-template-math.mjs when the
  // template changes.
  const { snippets } = JSON.parse(readFileSync(new URL('./fixtures/template-math.json', import.meta.url), 'utf8'));
  assert.ok(snippets.length >= 12, 'fixture present');
  const MathJax = await initMathJaxLikePane();
  await MathJax.tex2svgPromise(DEFAULT_PREAMBLE);
  const problems = [];
  for (const s of snippets) {
    const node = await MathJax.tex2svgPromise(flattenLatex(s.tex), { display: s.display });
    const html = MathJax.startup.adaptor.outerHTML(node);
    const err = /data-mjx-error="([^"]*)"/.exec(html);
    if (err) problems.push(`ERROR "${err[1]}" in: ${s.tex.slice(0, 80)}`);
    else if (/fill="red"|mathcolor="red"/.test(html)) problems.push(`undefined macro in: ${s.tex.slice(0, 80)}`);
  }
  assert.deepEqual(problems, []);
});

test('faithful ports: sans-italic operators, rule under tensors, starred delimiters, SI units', async () => {
  // WHY: these were approximations before; each must now produce the intended
  // construct (a real glyph, not a red fallback) so template documents look
  // like the PDF.
  const MathJax = await initMathJaxLikePane();
  await MathJax.tex2svgPromise(DEFAULT_PREAMBLE);
  for (const tex of ['\\SfOpI{A}\\SfOpBI{B}', '\\barbelow{T}', '\\OpComm*{\\frac{a}{b}, c}', '\\Abs*{x}',
    '\\SI{1.2(3)}{\\kilo\\hertz}', '\\SI{10}{\\micro\\second}', '\\si{\\metre\\per\\second\\squared}',
    '\\CharFunc{A}', '\\mathbbm{1}', '\\PhaseSpaceArgs[n]', '\\xsize{4}( x \\xsize{4})', '\\sym{\\int}',
    '\\II\\hbar\\EE^{\\pdg}']) {
    const html = MathJax.startup.adaptor.outerHTML(await MathJax.tex2svgPromise(tex, { display: false }));
    assert.doesNotMatch(html, /data-mjx-error|fill="red"|mathcolor="red"/, tex);
  }
});
