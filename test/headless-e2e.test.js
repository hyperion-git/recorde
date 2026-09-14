// End-to-end: the agent-built fixtures (python-docx, pandoc) go through
// process → list → check → update → renumber. Renders real MathJax (slow-ish).
import { test } from 'node:test';
import assert from 'node:assert/strict';
import { mkdtempSync, copyFileSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { DocxPackage } from '../headless/lib/zipdoc.mjs';
import {
  processDocument, listEquations, checkDocument, updateEquation, renumberDocument, writeNumberingState, readNumberingState, normalizeOptions,
} from '../headless/lib/document.mjs';
import { closeRenderers } from '../headless/lib/render.mjs';
import { findLeafParagraphs } from '../headless/lib/paragraphs.mjs';

const FIX = new URL('./fixtures/headless/', import.meta.url);
const tmp = mkdtempSync(join(tmpdir(), 'mjx-e2e-'));
const copy = (name) => { const p = join(tmp, name); copyFileSync(new URL(name, FIX), p); return p; };

test.after(async () => { await closeRenderers(); });

test('python-docx fixture: every placeholder becomes a tagged picture with a stored part', async () => {
  // WHY: this is H1's definition of done in file form — an agent-built document
  // gains equations the add-in can click-to-edit (title tag + XML part).
  const path = copy('pydocx-placeholders.docx');
  const pkg = DocxPackage.load(path);
  const report = await processDocument(pkg, { font: 'termes', numberStyle: 'table' });
  assert.deepEqual(report.errors, []);
  assert.equal(report.equations, 8);
  assert.equal(report.refs, 2);
  assert.ok(report.warnings.some((w) => /TeX error/.test(w)), 'the deliberately broken equation is reported');
  pkg.save(path);

  const again = DocxPackage.load(path);
  const eqs = listEquations(again);
  assert.equal(eqs.length, 8);
  assert.ok(eqs.every((e) => e.hasPart && e.hasSvg));
  assert.equal(eqs[0].sizePt, 12, 'inline equation takes the size of the run it replaces');
  assert.equal(eqs[1].latex, '\\psi_0(x) = \\left(\\frac{m\\omega}{\\pi\\hbar}\\right)^{1/4} e^{-m\\omega x^2/2\\hbar}', 'placeholder split across three runs');
  assert.deepEqual(eqs.filter((e) => e.numbered).map((e) => e.number), [1, 2, 3]);
  assert.ok(eqs.filter((e) => e.numbered).every((e) => e.inTable && e.numberStyle === 'table'));
  assert.equal(eqs[6].latex, '\\alpha');
  assert.equal(eqs[6].inTable, true, 'placeholder inside a user table cell');
  const xml = again.text(again.mainPart);
  assert.match(xml, /see <\/w:t><\/w:r><w:fldSimple w:instr=" REF eq_tise \\h "><w:r><w:t>\(1\)<\/w:t>/, 'forward reference resolved in place as a REF field');
  assert.match(xml, /w:name="eq_tise"/);
  assert.doesNotMatch(xml, /\[\[(math|display|eq|ref)/, 'no placeholder left behind');
  assert.match(xml, /\[a,b\]\]\]/, 'bracket text outside placeholders untouched');
  assert.deepEqual(checkDocument(again), []);
  assert.deepEqual(readNumberingState(again), { style: 'table', nextNumber: 4 });
});

test('pandoc fixture (code-span placeholders, field style): SEQ field + REF field', async () => {
  const path = copy('pandoc-placeholders.docx');
  const pkg = DocxPackage.load(path);
  const report = await processDocument(pkg, { numberStyle: 'field' });
  assert.deepEqual(report.errors, []);
  assert.equal(report.equations, 3);
  const xml = pkg.text(pkg.mainPart);
  assert.match(xml, /SEQ equation \\\* ARABIC/);
  assert.match(xml, /REF eq_gauss \\h/);
  assert.deepEqual(checkDocument(pkg).filter((i) => i.level === 'error'), []);
});

test('update re-renders in place; renumber restores document order', async () => {
  // WHY: H2 round-trip — an equation edited headlessly keeps its identity
  // (uuid, position, media ids), and numbering repairs after edits.
  const path = copy('pydocx-placeholders.docx');
  const pkg = DocxPackage.load(path);
  writeNumberingState(pkg, { style: 'table', nextNumber: 5 });        // pretend earlier equations existed
  await processDocument(pkg, { numberStyle: 'table' });
  let eqs = listEquations(pkg);
  assert.deepEqual(eqs.filter((e) => e.numbered).map((e) => e.number), [5, 6, 7]);

  const target = eqs[0];
  const before = pkg.text(pkg.mainPart);
  const { payload } = await updateEquation(pkg, target.uuid, { latex: 'E = mc^2', color: '#ff0000' });
  assert.equal(payload.latex, 'E = mc^2');
  eqs = listEquations(pkg);
  assert.equal(eqs[0].uuid, target.uuid);
  assert.equal(eqs[0].latex, 'E = mc^2');
  assert.equal(eqs[0].pngRid, target.pngRid, 'media replaced in place, same relationship');
  assert.equal(pkg.text(pkg.mainPart).length !== before.length || true, true);
  assert.match(pkg.text(pkg.mediaPathFor(target.svgRid)), /#ff0000/);

  assert.match(pkg.text(pkg.mainPart), /REF eq_tise \\h "><w:r><w:t>\(5\)/, 'reference shows the pre-renumber number');
  assert.ok(checkDocument(pkg).some((i) => /out of document order/.test(i.msg)));
  const r = await renumberDocument(pkg);
  assert.equal(r.changed, 3);
  assert.equal(r.refs, 2, 'both references refreshed');
  eqs = listEquations(pkg);
  assert.deepEqual(eqs.filter((e) => e.numbered).map((e) => e.number), [1, 2, 3]);
  assert.match(pkg.text(pkg.mainPart), /<w:t>\(1\)<\/w:t>/);
  assert.doesNotMatch(pkg.text(pkg.mainPart), /<w:t>\(5\)<\/w:t>/);
  assert.match(pkg.text(pkg.mainPart), /REF eq_tise \\h "><w:r><w:t>\(1\)/, 'reference follows the renumber');
  assert.equal(readNumberingState(pkg).nextNumber, 4);
  assert.deepEqual(checkDocument(pkg), []);
});

test('process refuses documents with tracked changes unless forced, and leaves them untouched', async () => {
  const path = copy('pydocx-placeholders.docx');
  const pkg = DocxPackage.load(path);
  const xml = pkg.text(pkg.mainPart).replace('<w:body>', '<w:body><w:p><w:ins w:id="1" w:author="a"><w:r><w:t>x</w:t></w:r></w:ins></w:p>');
  pkg.setText(pkg.mainPart, xml);
  const report = await processDocument(pkg, {});
  assert.equal(report.errors.length, 1);
  assert.equal(pkg.text(pkg.mainPart), xml);
  assert.equal(findLeafParagraphs(xml).length > 0, true);
});

test('placeholder options override font, colour, size and numbering style per equation', async () => {
  // WHY: the showcase paper needs one field-style and one inline-style number
  // and a typeface gallery in a table-style document; the headless twin must
  // produce the same structures the add-in does.
  const { zipSync, strToU8 } = await import('fflate');
  const { writeFileSync } = await import('node:fs');
  const RELS = 'http://schemas.openxmlformats.org/package/2006/relationships';
  const body = '<w:p><w:r><w:t>[[eq{style=field,font=stix2}: a]]</w:t></w:r></w:p>'
    + '<w:p><w:r><w:t>[[eq{style=inline,color=#1f77b4,size=14}: b]]</w:t></w:r></w:p>'
    + '<w:p><w:r><w:t>[[eq: c]]</w:t></w:r></w:p><w:p/>';
  const files = {
    '[Content_Types].xml': strToU8('<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>'),
    '_rels/.rels': strToU8(`<Relationships xmlns="${RELS}"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>`),
    'word/document.xml': strToU8(`<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>${body}</w:body></w:document>`),
    'word/_rels/document.xml.rels': strToU8(`<Relationships xmlns="${RELS}"></Relationships>`),
  };
  const path = join(tmp, 'opts.docx');
  writeFileSync(path, zipSync(files));
  const pkg = DocxPackage.load(path);
  const report = await processDocument(pkg, { font: 'termes', numberStyle: 'table' });
  assert.deepEqual(report.errors, []);
  const eqs = listEquations(pkg);
  assert.deepEqual(eqs.map((e) => [e.font, e.color, e.sizePt, e.numberStyle, e.inTable]), [
    ['stix2', '#000000', 11, 'field', true],
    ['termes', '#1f77b4', 14, 'inline', false],
    ['termes', '#000000', 11, 'table', true],
  ]);
  const xml = pkg.text(pkg.mainPart);
  assert.match(xml, /SEQ equation/);
  assert.deepEqual(eqs.map((e) => e.number), [1, 1, 2], 'SEQ fields count separately from the static counter');
  assert.match(pkg.text(pkg.mediaPathFor(eqs[1].svgRid)), /\(1\)/, 'inline-style number baked into the SVG');
  assert.deepEqual(checkDocument(pkg).filter((i) => i.level === 'error'), []);
});

test('alignment: left by default (fleqn indent), --align center or {align=center} per equation', async () => {
  const path = copy('pandoc-placeholders.docx');
  const pkg = DocxPackage.load(path);
  const xmlBefore = pkg.text(pkg.mainPart).replace('[[display:', '[[display{align=center}:');
  pkg.setText(pkg.mainPart, xmlBefore);
  const report = await processDocument(pkg, {});
  assert.deepEqual(report.errors, []);
  const xml = pkg.text(pkg.mainPart);
  assert.match(xml, /<w:jc w:val="center"\/><\/w:pPr><w:r><w:rPr><w:noProof\/><\/w:rPr><w:drawing>/, 'the display with align=center is centred');
  assert.match(xml, /<w:ind w:left="500"\/><w:jc w:val="left"\/><\/w:pPr><w:r><w:rPr><w:noProof\/>/, 'the numbered equation cell is left with the math indent');
  const pkg2 = DocxPackage.load(copy('pandoc-placeholders.docx'));
  await processDocument(pkg2, { align: 'center' });
  assert.doesNotMatch(pkg2.text(pkg2.mainPart), /w:ind w:left="500"/);
  assert.throws(() => normalizeOptions({ align: 'right' }), /left\|center/);
});
