import { test } from 'node:test';
import assert from 'node:assert/strict';
import { mkdtempSync, writeFileSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { zipSync, strToU8 } from 'fflate';
import { DocxPackage, REL } from '../headless/lib/zipdoc.mjs';

const RELS_NS = 'http://schemas.openxmlformats.org/package/2006/relationships';
function minimalDocx(extra = {}) {
  const files = {
    '[Content_Types].xml': strToU8('<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
      + '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
      + '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>'),
    '_rels/.rels': strToU8(`<Relationships xmlns="${RELS_NS}"><Relationship Id="rId1" Type="${REL.officeDocument}" Target="word/document.xml"/></Relationships>`),
    'word/document.xml': strToU8('<w:document><w:body><w:p/></w:body></w:document>'),
    'word/_rels/document.xml.rels': strToU8(`<Relationships xmlns="${RELS_NS}"><Relationship Id="rId3" Type="x" Target="styles.xml"/></Relationships>`),
    ...extra,
  };
  const dir = mkdtempSync(join(tmpdir(), 'mjx-'));
  const path = join(dir, 'min.docx');
  writeFileSync(path, zipSync(files));
  return path;
}

test('media, rels and content types are added consistently and survive save/load', () => {
  // WHY: a picture whose rId, media file or content type is missing makes Word
  // "repair" the document — the failure mode agents cannot see until delivery.
  const path = minimalDocx();
  const pkg = DocxPackage.load(path);
  const svg = pkg.addMedia('svg', strToU8('<svg/>'));
  const png = pkg.addMedia('png', new Uint8Array([1, 2, 3]));
  assert.equal(svg.path, 'word/media/image1.svg');
  assert.equal(png.path, 'word/media/image2.png');
  assert.equal(svg.rId, 'rId1');                 // lowest free id; rId3 was taken
  assert.equal(png.rId, 'rId2');
  const out = path.replace('min', 'out');
  pkg.save(out);
  const again = DocxPackage.load(out);
  assert.equal(again.mediaPathFor('rId2'), 'word/media/image2.png');
  assert.match(again.text('[Content_Types].xml'), /Extension="svg" ContentType="image\/svg\+xml"/);
  assert.match(again.text('[Content_Types].xml'), /Extension="png"/);
  assert.equal((again.text('[Content_Types].xml').match(/Extension="png"/g) || []).length, 1);
  assert.deepEqual(again.bytes('word/media/image2.png'), new Uint8Array([1, 2, 3]));
});

test('custom XML items take the next free index and are fully wired', () => {
  // WHY: Word templates often ship item1.xml (bibliography, cover page props);
  // colliding with it corrupts the document.
  const path = minimalDocx({ 'customXml/item1.xml': strToU8('<b:Sources xmlns:b="urn:b" xmlns="urn:b"/>') });
  const pkg = DocxPackage.load(path);
  const n = pkg.addCustomXml('<equation xmlns="urn:mathjax-office:equations" uuid="u"/>', 'urn:mathjax-office:equations');
  assert.equal(n, 2);
  assert.ok(pkg.has('customXml/item2.xml') && pkg.has('customXml/itemProps2.xml') && pkg.has('customXml/_rels/item2.xml.rels'));
  assert.match(pkg.text('customXml/itemProps2.xml'), /ds:itemID="\{[0-9A-F-]{36}\}"/);
  assert.match(pkg.text('customXml/itemProps2.xml'), /ds:uri="urn:mathjax-office:equations"/);
  assert.match(pkg.text('[Content_Types].xml'), /PartName="\/customXml\/itemProps2.xml"/);
  assert.ok(pkg.readRels('word/document.xml').some((r) => r.type === REL.customXml && r.target === '../customXml/item2.xml'));
  const items = pkg.listCustomXml();
  assert.deepEqual(items.map((i) => [i.index, i.namespace]), [[1, 'urn:b'], [2, 'urn:mathjax-office:equations']]);
  pkg.removeCustomXml(2);
  assert.equal(pkg.listCustomXml().length, 1);
  assert.ok(!pkg.readRels('word/document.xml').some((r) => r.type === REL.customXml));
  assert.doesNotMatch(pkg.text('[Content_Types].xml'), /itemProps2/);
});

test('untouched parts are carried byte-for-byte', () => {
  const bytes = new Uint8Array([0, 255, 13, 10, 7]);
  const path = minimalDocx({ 'word/opaque.bin': bytes });
  const pkg = DocxPackage.load(path);
  const out = path.replace('min', 'out2');
  pkg.save(out);
  assert.deepEqual(DocxPackage.load(out).bytes('word/opaque.bin'), bytes);
});
