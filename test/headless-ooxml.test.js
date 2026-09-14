import { test } from 'node:test';
import assert from 'node:assert/strict';
import {
  emitPictureRun, emitNumberTable, emitNumberRuns, emitRefRun, withJc, rewriteNumberInCell, enclosingElement,
  findEquationPictures, nextDocPrId, sectionTextWidthTwips, sectionTextWidthAt, defaultFontSizePt, ptToEmu, positionHalfPoints, bookmarkName,
} from '../headless/lib/ooxml.mjs';
import * as ooxml from '../headless/lib/ooxml.mjs';

const run = () => emitPictureRun({ pngRid: 'rId7', svgRid: 'rId8', widthPt: 30.5, heightPt: 12, descentPt: 3.4,
  uuid: 'u-1', latex: 'a<b & "c"', docPrId: 5, displayMode: 'inline' });

test('picture run carries the add-in tags: mjx:<uuid> title, LaTeX descr, SVG blip, baseline shift', () => {
  // WHY: these four things are what make a headless equation indistinguishable
  // from an add-in insert — click-to-edit keys on the title, recovery on the
  // descr, Word renders the svgBlip, and the shift seats inline math on the line.
  const r = run();
  assert.match(r, /title="mjx:u-1"/);
  assert.match(r, /descr="a&lt;b &amp; &quot;c&quot;"/);
  assert.match(r, /<asvg:svgBlip [^>]*r:embed="rId8"/);
  assert.match(r, /<a:blip r:embed="rId7"/);
  assert.match(r, /<w:position w:val="-7"\/>/);           // -round(2·3.4)
  assert.match(r, new RegExp(`<wp:extent cx="${ptToEmu(30.5)}" cy="${ptToEmu(12)}"/>`));
});

test('display pictures get no baseline shift and EMU/half-point conversions round', () => {
  const r = emitPictureRun({ pngRid: 'a', svgRid: 'b', widthPt: 1, heightPt: 1, descentPt: 9, uuid: 'u', latex: 'x', docPrId: 1, displayMode: 'display' });
  assert.doesNotMatch(r, /w:position/);
  assert.equal(ptToEmu(1), 12700);
  assert.equal(positionHalfPoints(2.24), -4);
});

test('findEquationPictures round-trips what emitPictureRun wrote, in document order', () => {
  const xml = `<w:body><w:p>${run()}</w:p><w:tbl><w:tr><w:tc><w:p>${run().replace('mjx:u-1', 'mjx:u-2')}</w:p></w:tc></w:tr></w:tbl></w:body>`;
  const pics = findEquationPictures(xml);
  assert.deepEqual(pics.map((p) => p.uuid), ['u-1', 'u-2']);
  assert.equal(pics[0].latex, 'a<b & "c"');
  assert.equal(pics[0].pngRid, 'rId7');
  assert.equal(pics[0].svgRid, 'rId8');
  assert.equal(xml.slice(pics[1].runStart, pics[1].runEnd).startsWith('<w:r>'), true);
  assert.equal(nextDocPrId(xml), 6);
});

test('number table: borderless 1×2, 100 % autofit, 54pt number cell (36pt in narrow columns)', () => {
  // WHY: in a two-column section a fixed full-page table overflows the column
  // and the number falls off the page. 100 %/autofit lets Word fit the table
  // to its column, as the add-in's autoFitWindow does.
  const t = emitNumberTable({ runXml: '<w:r/>', numberXml: emitNumberRuns({ number: 3 }), textWidthTwips: 9360 });
  assert.match(t, /<w:tblW w:w="5000" w:type="pct"\/>/);
  assert.match(t, /<w:tblLayout w:type="autofit"\/>/);
  assert.match(t, /<w:gridCol w:w="8280"\/><w:gridCol w:w="1080"\/>/);
  assert.match(t, /<w:top w:val="nil"\/>/);
  assert.match(t, /<w:jc w:val="right"\/><\/w:pPr><w:r><w:t>\(3\)<\/w:t>/);
  const narrow = emitNumberTable({ runXml: '<w:r/>', numberXml: '<w:r/>', textWidthTwips: 4536 });
  assert.match(narrow, /<w:gridCol w:w="3816"\/><w:gridCol w:w="720"\/>/);
});

test('sectionTextWidthAt: the governing section is the next sectPr; columns divide the width', () => {
  // WHY: a paragraph in a two-column section must get the COLUMN width, and
  // paragraph-level breaks without pgSz inherit the page geometry.
  const body = '<w:body><w:p>A</w:p><w:p><w:pPr><w:sectPr><w:type w:val="continuous"/><w:cols w:num="1"/></w:sectPr></w:pPr></w:p>'
    + '<w:p>B</w:p><w:p><w:pPr><w:sectPr><w:cols w:num="2" w:space="288"/></w:sectPr></w:pPr></w:p>'
    + '<w:p>C</w:p><w:sectPr><w:pgSz w:w="12240"/><w:pgMar w:left="1080" w:right="1080"/></w:sectPr></w:body>';
  assert.equal(sectionTextWidthAt(body, body.indexOf('A')), 10080);
  assert.equal(sectionTextWidthAt(body, body.indexOf('B')), (10080 - 288) / 2);
  assert.equal(sectionTextWidthAt(body, body.indexOf('C')), 10080);
  assert.equal(sectionTextWidthAt('<w:body/>', 0), 9360);
});

test('number runs: SEQ field for the field style, bookmark when labelled; refs mirror it', () => {
  const f = emitNumberRuns({ number: 2, style: 'field', label: 'a-b', bookmarkId: 4 });
  assert.match(f, /<w:bookmarkStart w:id="4" w:name="eq_a_b"\/>/);
  assert.match(f, /<w:fldSimple w:instr=" SEQ equation \\\* ARABIC "><w:r><w:t>2<\/w:t>/);
  assert.equal(emitRefRun({ number: 2, label: 'a-b' }), '<w:r><w:t>(2)</w:t></w:r>');
  assert.match(emitRefRun({ number: 2, label: 'a-b', field: true }), /REF eq_a_b \\h/);
  assert.equal(bookmarkName('x'.repeat(50)).length, 40);
});

test('rewriteNumberInCell updates static text and SEQ cached results', () => {
  assert.match(rewriteNumberInCell('<w:tc><w:p><w:r><w:t>(3)</w:t></w:r></w:p></w:tc>', 7), /\(7\)/);
  const seq = '<w:tc><w:p><w:r><w:t>(</w:t></w:r><w:fldSimple w:instr=" SEQ equation \\* ARABIC "><w:r><w:t>3</w:t></w:r></w:fldSimple></w:p></w:tc>';
  assert.match(rewriteNumberInCell(seq, 9), /<w:t>9<\/w:t>/);
});

test('withJc replaces or inserts justification without disturbing other pPr', () => {
  assert.equal(withJc('', 'center'), '<w:pPr><w:jc w:val="center"/></w:pPr>');
  assert.equal(withJc('<w:pPr><w:jc w:val="left"/><w:rPr/></w:pPr>', 'center'), '<w:pPr><w:jc w:val="center"/><w:rPr/></w:pPr>');
  assert.equal(withJc('<w:pPr><w:pStyle w:val="X"/><w:rPr/></w:pPr>', 'center'), '<w:pPr><w:pStyle w:val="X"/><w:jc w:val="center"/><w:rPr/></w:pPr>');
});

test('enclosingElement is nesting-aware', () => {
  const xml = '<w:tbl><w:tr><w:tc><w:tbl><w:tr><w:tc>X</w:tc></w:tr></w:tbl></w:tc></w:tr></w:tbl>';
  const pos = xml.indexOf('X');
  const inner = enclosingElement(xml, pos, 'w:tbl');
  assert.equal(xml.slice(inner.start, inner.end), '<w:tbl><w:tr><w:tc>X</w:tc></w:tr></w:tbl>');
  assert.equal(enclosingElement('<w:p>Y</w:p>', 6, 'w:tbl'), null);
});

test('section text width and default font size fall back sensibly', () => {
  assert.equal(sectionTextWidthTwips('<w:sectPr><w:pgSz w:w="11906"/><w:pgMar w:top="1" w:right="1000" w:left="1500"/></w:sectPr>'), 9406);
  assert.equal(sectionTextWidthTwips('<w:body/>'), 9360);
  assert.equal(defaultFontSizePt('<w:styles><w:docDefaults><w:rPrDefault><w:rPr><w:sz w:val="20"/></w:rPr></w:rPrDefault></w:docDefaults></w:styles>'), 10);
  assert.equal(defaultFontSizePt(''), 11);
});

test('TeX display skips: one body size above and below, continuation without indent', () => {
  // WHY: LaTeX's \abovedisplayskip/\belowdisplayskip equal the body size
  // (10/11/12 pt); without them the display sits flush against the prose and
  // the page reads dense. The continuation text after a display is the same
  // paragraph in TeX, so it must not be re-indented.
  const { displaySkipTwips, withSpacing, withoutFirstLineIndent, emitDisplayParagraph, emitNumberTable } = ooxml;
  assert.equal(displaySkipTwips(10), 200);
  assert.equal(displaySkipTwips(12), 240);
  assert.equal(displaySkipTwips(undefined), 220);
  assert.equal(withSpacing('', 200, 200), '<w:pPr><w:spacing w:before="200" w:after="200"/></w:pPr>');
  assert.equal(withSpacing('<w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/><w:jc w:val="both"/></w:pPr>', 200, 200),
    '<w:pPr><w:spacing w:before="200" w:after="200" w:line="240" w:lineRule="auto"/><w:jc w:val="both"/></w:pPr>');
  assert.equal(withSpacing('<w:pPr><w:pStyle w:val="X"/><w:ind w:firstLine="240"/><w:jc w:val="both"/></w:pPr>', 200, 200),
    '<w:pPr><w:pStyle w:val="X"/><w:spacing w:before="200" w:after="200"/><w:ind w:firstLine="240"/><w:jc w:val="both"/></w:pPr>');
  assert.equal(withoutFirstLineIndent('<w:pPr><w:ind w:left="100" w:firstLine="240"/></w:pPr>'), '<w:pPr><w:ind w:left="100" w:firstLine="0"/></w:pPr>');
  assert.equal(withoutFirstLineIndent('<w:pPr><w:ind w:hanging="240"/></w:pPr>'), '<w:pPr><w:ind w:hanging="240"/></w:pPr>');
  assert.match(emitDisplayParagraph('<w:r/>', '', 200, 'center'), /<w:spacing w:before="200" w:after="200"\/><w:jc w:val="center"\/>/);
  const t = emitNumberTable({ runXml: '<w:r/>', numberXml: '<w:r/>', textWidthTwips: 9360, skipTwips: 200 });
  assert.equal((t.match(/<w:spacing w:before="200" w:after="200"\/>/g) || []).length, 2, 'both cells carry the skip');
});

test('display alignment: left = fleqn with a 25 pt math indent (default), center = plain LaTeX', () => {
  // WHY: the user asked for left-aligned displays by default with centred as
  // an option; LaTeX's fleqn indents by \mathindent so the equation does not
  // hug the margin. Both the bare paragraph and the numbering table's equation
  // cell must follow the same rule.
  const { emitDisplayParagraph, emitNumberTable, withLeftIndent, MATH_INDENT_TWIPS } = ooxml;
  assert.equal(MATH_INDENT_TWIPS, 500);
  const left = emitDisplayParagraph('<w:r/>', '', 200);
  assert.match(left, /<w:spacing[^>]*\/><w:ind w:left="500" w:firstLine="0"\/><w:jc w:val="left"\/>/);
  const center = emitDisplayParagraph('<w:r/>', '', 200, 'center');
  assert.match(center, /<w:jc w:val="center"\/>/);
  assert.doesNotMatch(center, /w:ind/);
  assert.equal(withLeftIndent('<w:pPr><w:ind w:left="720" w:firstLine="240"/><w:jc w:val="both"/></w:pPr>', 500),
    '<w:pPr><w:ind w:left="500" w:firstLine="0"/><w:jc w:val="both"/></w:pPr>');
  const t = emitNumberTable({ runXml: '<w:r/>', numberXml: '<w:r/>', textWidthTwips: 9360 });
  assert.match(t, /<w:ind w:left="500"\/><w:jc w:val="left"\/><\/w:pPr><w:r\/>/);
  const tc = emitNumberTable({ runXml: '<w:r/>', numberXml: '<w:r/>', textWidthTwips: 9360, align: 'center' });
  assert.match(tc, /<w:jc w:val="center"\/><\/w:pPr><w:r\/>/);
});
