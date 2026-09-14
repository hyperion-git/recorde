import { test } from 'node:test';
import assert from 'node:assert/strict';
import { DOMParser } from 'linkedom';

// storage.parseXml uses `new DOMParser()`, a browser global. Provide it.
globalThis.DOMParser = DOMParser;

import {
  buildXml, parseXml, embedSourceInSvg, readSourceFromSvg,
  buildPreambleXml, parsePreambleXml,
  buildNumberingXml, parseNumberingXml,
} from '../src/storage.js';

test('buildXml → parseXml round-trips, preserving XML metacharacters', () => {
  // WHY: the LaTeX source is the whole point of the storage layer. If escaping
  // drops or mangles < & " ' it silently corrupts equations on reload.
  const p = {
    uuid: 'u1',
    latex: 'a < b & c > d "q" \'ap\' \\frac{1}{2}',
    font: 'stix2', sizePt: 12, displayMode: 'display',
  };
  const back = parseXml(buildXml(p));
  assert.equal(back.uuid, 'u1');
  assert.equal(back.latex, p.latex);
  assert.equal(back.font, 'stix2');
  assert.equal(back.sizePt, 12);
  assert.equal(back.displayMode, 'display');
});

test('buildXml actually escapes < and &', () => {
  const xml = buildXml({ uuid: 'u', latex: 'x<y&z', font: 'tex', sizePt: 11 });
  assert.ok(xml.includes('&lt;'), 'expected &lt;');
  assert.ok(xml.includes('&amp;'), 'expected &amp;');
  assert.ok(!xml.includes('x<y'), 'raw < must not survive');
});

test('parseXml returns null on non-equation XML', () => {
  assert.equal(parseXml('<foo/>'), null);
});

test('parseXml applies defaults for missing attributes', () => {
  const xml = '<equation xmlns="urn:mathjax-office:equations" uuid="z"><latex>x</latex></equation>';
  const d = parseXml(xml);
  assert.equal(d.font, 'tex');
  assert.equal(d.sizePt, 11);
  assert.equal(d.displayMode, 'inline');
  assert.equal(d.color, '#000000');
  // WHY: parts written before WP2.2 have no numbering attributes — they must
  // load as plain unnumbered equations, not crash or invent a number.
  assert.equal(d.numbered, false);
  assert.equal(d.number, 0);
});

test('buildXml → parseXml round-trips numbered + number + numberStyle', () => {
  // WHY: the number, its on/off flag, AND the per-equation style must survive a
  // save/reload — renumber dispatches on numberStyle, so losing it would send a
  // table equation down the inline re-render path.
  const back = parseXml(buildXml({
    uuid: 'u', latex: 'x', font: 'tex', sizePt: 11, displayMode: 'display',
    numbered: true, number: 4, numberStyle: 'table',
  }));
  assert.equal(back.numbered, true);
  assert.equal(back.number, 4);
  assert.equal(back.numberStyle, 'table');
});

test('numberStyle defaults to inline on parts that predate it', () => {
  // WHY: increment-1 numbered equations have no numberStyle attribute; they ARE
  // inline, so they must default to inline and renumber via the picture path.
  const xml = '<equation xmlns="urn:mathjax-office:equations" uuid="z" '
    + 'numbered="true" number="2"><latex>x</latex></equation>';
  assert.equal(parseXml(xml).numberStyle, 'inline');
});

test('numbered flag is a real boolean, not the string "false"', () => {
  // WHY: a truthy "false" string would make every reloaded equation look numbered.
  const back = parseXml(buildXml({ uuid: 'u', latex: 'x', font: 'tex', sizePt: 11, numbered: false }));
  assert.strictEqual(back.numbered, false);
});

test('buildNumberingXml → parseNumberingXml round-trips style + counter', () => {
  // WHY: the per-document numbering style and next-number counter travel in the
  // .docx; if they don't round-trip, numbering restarts at 1 on every reopen.
  const back = parseNumberingXml(buildNumberingXml({ style: 'table', nextNumber: 12 }));
  assert.equal(back.style, 'table');
  assert.equal(back.nextNumber, 12);
});

test('parseNumberingXml returns null on non-numbering XML', () => {
  assert.equal(parseNumberingXml('<foo/>'), null);
});

test('buildNumberingXml defaults to inline / 1', () => {
  const back = parseNumberingXml(buildNumberingXml());
  assert.equal(back.style, 'inline');
  assert.equal(back.nextNumber, 1);
});

test('buildXml → parseXml round-trips the color', () => {
  const back = parseXml(buildXml({ uuid: 'u', latex: 'x', font: 'tex', sizePt: 11, color: '#ff8800' }));
  assert.equal(back.color, '#ff8800');
});

test('embedSourceInSvg → readSourceFromSvg recovers the payload', () => {
  const svg = '<svg viewBox="0 0 1 1"><path></path></svg>';
  const withDesc = embedSourceInSvg(svg, {
    uuid: 'u2', latex: '\\alpha', font: 'pagella', sizePt: 9, displayMode: 'inline',
  });
  const d = readSourceFromSvg(withDesc);
  assert.equal(d.uuid, 'u2');
  assert.equal(d.latex, '\\alpha');
  assert.equal(d.font, 'pagella');
});

test('buildPreambleXml → parsePreambleXml round-trips multi-line text + metachars', () => {
  // WHY: macro defs are multi-line and contain < & { }; escaping must survive.
  const text = '\\newcommand{\\R}{\\mathbb{R}}\n\\newcommand{\\lt}{x < y \\& z}';
  assert.equal(parsePreambleXml(buildPreambleXml(text)), text);
});

test('parsePreambleXml returns null on non-preamble XML', () => {
  assert.equal(parsePreambleXml('<foo/>'), null);
});
