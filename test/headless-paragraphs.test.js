import { test } from 'node:test';
import assert from 'node:assert/strict';
import {
  findLeafParagraphs, parseParagraph, serializeParagraph, joinedText, spliceParagraph, runSizePt,
} from '../headless/lib/paragraphs.mjs';

const P = (inner, pPr = '') => `<w:p>${pPr}${inner}</w:p>`;
const R = (text, rPr = '') => `<w:r>${rPr}<w:t xml:space="preserve">${text}</w:t></w:r>`;

test('parse → serialize is the identity for ordinary paragraphs', () => {
  // WHY: Word is fussy about re-serialized XML; a paragraph we merely looked at
  // must come back byte-identical.
  const xml = P(R('Hello ', '<w:rPr><w:b/></w:rPr>') + R('world') + '<w:bookmarkStart w:id="0" w:name="x"/>', '<w:pPr><w:jc w:val="both"/></w:pPr>');
  assert.equal(serializeParagraph(parseParagraph(xml)), xml);
});

test('joinedText concatenates run text and marks non-text with U+FFFC', () => {
  // WHY: a tab is content (a placeholder must not span it); a hyperlink's
  // wrapper tags are not (its text must still be matchable).
  const p = parseParagraph(P(R('ab') + '<w:r><w:tab/><w:t>c</w:t></w:r>' + '<w:hyperlink r:id="rId9">' + R('L') + '</w:hyperlink>' + '<w:r><w:drawing/></w:r>'));
  assert.equal(joinedText(p).text, 'ab￼cL￼');
});

test('spliceParagraph replaces a range spanning three runs, keeping each rPr', () => {
  // WHY: docx-js and Word fragment text across runs unpredictably; a placeholder
  // typed as one string may arrive as three runs. The split must keep the
  // surrounding formatting on the prefix and suffix pieces.
  const p = parseParagraph(P(R('see [[ma', '<w:rPr><w:i/></w:rPr>') + R('th: x') + R(']] ok', '<w:rPr><w:b/></w:rPr>')));
  const { text } = joinedText(p);
  const start = text.indexOf('[['), end = text.indexOf(']]') + 2;
  const out = serializeParagraph(spliceParagraph(p, [{ start, end, replacement: ['<EQ/>'] }]));
  assert.equal(out, P(R('see ', '<w:rPr><w:i/></w:rPr>') + '<EQ/>' + R(' ok', '<w:rPr><w:b/></w:rPr>')));
});

test('spliceParagraph handles two edits in one run and an edit at the very start', () => {
  const p = parseParagraph(P(R('[[a]] mid [[b]] end')));
  const edits = [{ start: 0, end: 5, replacement: ['<A/>'] }, { start: 10, end: 15, replacement: ['<B/>'] }];
  assert.equal(serializeParagraph(spliceParagraph(p, edits)), P('<A/>' + R(' mid ') + '<B/>' + R(' end')));
});

test('spliceParagraph keeps bookmarks and drops covered object segments', () => {
  const p = parseParagraph(P('<w:bookmarkStart w:id="1" w:name="k"/>' + R('x') + '<w:r><w:tab/></w:r>' + R('y') + '<w:bookmarkEnd w:id="1"/>'));
  // text = 'x￼y' — replace the tab (offset 1) only
  const out = serializeParagraph(spliceParagraph(p, [{ start: 1, end: 2, replacement: ['<T/>'] }]));
  assert.equal(out, P('<w:bookmarkStart w:id="1" w:name="k"/>' + R('x') + '<T/>' + R('y') + '<w:bookmarkEnd w:id="1"/>'));
});

test('findLeafParagraphs skips text-box containers but finds their inner paragraphs', () => {
  const inner = P(R('inner'));
  const xml = '<w:body>' + P(R('a')) + '<w:p><w:r><w:pict><w:txbxContent>' + inner + '</w:txbxContent></w:pict></w:r></w:p>' + P(R('b')) + '</w:body>';
  const leaves = findLeafParagraphs(xml).map((l) => l.xml);
  assert.deepEqual(leaves, [P(R('a')), inner, P(R('b'))]);
});

test('runSizePt reads half-points from rPr', () => {
  assert.equal(runSizePt({ rPr: '<w:rPr><w:sz w:val="24"/></w:rPr>' }), 12);
  assert.equal(runSizePt({ rPr: '' }), null);
});
