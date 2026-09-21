// Equation source persistence.
//
// Primary store: Word custom XML parts in namespace `urn:mathjax-office:equations`.
// One part per equation, keyed by a UUID also written into the inline picture's
// `altTextTitle`. writeXmlPart deletes any pre-existing part with the same
// UUID before adding, so updates don't leave stale duplicates.
//
// Fallback: same payload XML-encoded into an SVG `<desc>` element so the source
// can be recovered from an SVG copy-pasted into another document.

export const NS = 'urn:mathjax-office:equations';
export const PREAMBLE_NS = 'urn:mathjax-office:preamble';
export const NUMBERING_NS = 'urn:mathjax-office:numbering';
export const ALT_PREFIX = 'mjx:';

export function newUuid() {
  return crypto.randomUUID();
}

// XML escape for attribute values and text content.
function escapeXml(s) {
  return String(s).replace(/[&<>"']/g, (c) =>
    ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&apos;' })[c]);
}

// `writer` (optional) records which tool wrote the part (the headless CLI sets
// "mjx-docx/<version>"); the add-in omits it and parseXml ignores it.
export function buildXml({
  uuid, latex, font, sizePt, displayMode = 'inline', color = '#000000',
  numbered = false, number = 0, numberStyle = 'inline', writer = null,
}) {
  const esc = escapeXml;
  return `<equation xmlns="${NS}" uuid="${esc(uuid)}" font="${esc(font)}" `
    + `sizePt="${esc(sizePt)}" displayMode="${esc(displayMode)}" color="${esc(color)}" `
    + `numbered="${numbered ? 'true' : 'false'}" number="${esc(number)}" `
    + `numberStyle="${esc(numberStyle)}"${writer ? ` writer="${esc(writer)}"` : ''}>`
    + `<latex>${esc(latex)}</latex>`
    + `</equation>`;
}

export function parseXml(xmlString) {
  const doc = new DOMParser().parseFromString(xmlString, 'application/xml');
  const eq = doc.getElementsByTagNameNS(NS, 'equation')[0];
  if (!eq) return null;
  const latex = eq.getElementsByTagNameNS(NS, 'latex')[0]?.textContent ?? '';
  return {
    uuid: eq.getAttribute('uuid') || '',
    font: eq.getAttribute('font') || 'tex',
    sizePt: parseFloat(eq.getAttribute('sizePt') || '11'),
    displayMode: eq.getAttribute('displayMode') || 'inline',
    color: eq.getAttribute('color') || '#000000',
    // numbered/number/numberStyle are absent on parts written before WP2.2 —
    // default to an unnumbered inline equation so older documents just have no
    // equation numbers. numberStyle is stored PER equation (not only per doc) so
    // renumber dispatches correctly in a mixed-style document.
    numbered: eq.getAttribute('numbered') === 'true',
    number: parseInt(eq.getAttribute('number') || '0', 10) || 0,
    numberStyle: eq.getAttribute('numberStyle') || 'inline',
    latex,
  };
}

// Primary: write to a custom XML part, replacing any existing part with the
// same UUID. Resolves with the new part's id.
export async function writeXmlPart(payload) {
  const xml = buildXml(payload);
  return Word.run(async (context) => {
    const parts = context.document.customXmlParts.getByNamespace(NS);
    parts.load('items');
    await context.sync();

    const blobs = parts.items.map((p) => p.getXml());
    await context.sync();

    for (let i = 0; i < parts.items.length; i++) {
      const data = parseXml(blobs[i].value);
      if (data && data.uuid === payload.uuid) parts.items[i].delete();
    }

    const newPart = context.document.customXmlParts.add(xml);
    newPart.load('id');
    await context.sync();
    return newPart.id;
  });
}

// Primary: look up by uuid, scanning our namespace.
export async function readXmlPartByUuid(uuid) {
  return Word.run(async (context) => {
    const parts = context.document.customXmlParts.getByNamespace(NS);
    parts.load('items');
    await context.sync();
    const blobs = parts.items.map((p) => p.getXml());
    await context.sync();
    for (const blob of blobs) {
      const data = parseXml(blob.value);
      if (data && data.uuid === uuid) return data;
    }
    return null;
  });
}

// Fallback: stuff payload inside the SVG's <desc>. Returns the modified SVG string.
export function embedSourceInSvg(svgString, payload) {
  const xml = buildXml(payload);
  return svgString.replace(/<svg([^>]*)>/, (_, attrs) =>
    `<svg${attrs}><desc>${xml.replace(/]]>/g, ']]]]><![CDATA[>')}</desc>`);
}

// Fallback: try to extract payload from an SVG's <desc>.
export function readSourceFromSvg(svgString) {
  const match = svgString.match(/<desc[^>]*>([\s\S]*?)<\/desc>/);
  if (!match) return null;
  return parseXml(match[1]);
}

// ---- Document macro preamble (WP2.4) ----
// One \newcommand preamble per document, in its own custom XML part so it travels
// with the .docx and is prepended to every render.

export function buildPreambleXml(text) {
  return `<preamble xmlns="${PREAMBLE_NS}">${escapeXml(text)}</preamble>`;
}

export function parsePreambleXml(xmlString) {
  const doc = new DOMParser().parseFromString(xmlString, 'application/xml');
  const el = doc.getElementsByTagNameNS(PREAMBLE_NS, 'preamble')[0];
  return el ? (el.textContent ?? '') : null;
}

// Replace the document's preamble part (delete-then-add). Blank text clears it.
export async function writePreamble(text) {
  return Word.run(async (context) => {
    const parts = context.document.customXmlParts.getByNamespace(PREAMBLE_NS);
    parts.load('items');
    await context.sync();
    for (const p of parts.items) p.delete();
    if (text && text.trim()) context.document.customXmlParts.add(buildPreambleXml(text));
    await context.sync();
  });
}

export async function readPreamble() {
  return Word.run(async (context) => {
    const parts = context.document.customXmlParts.getByNamespace(PREAMBLE_NS);
    parts.load('items');
    await context.sync();
    if (!parts.items.length) return '';
    const blob = parts.items[0].getXml();
    await context.sync();
    return parsePreambleXml(blob.value) || '';
  });
}

// ---- Document numbering state (WP2.2) ----
// One part per document holding the active numbering style and the next-number
// counter, so numbering choices travel with the .docx.

export function buildNumberingXml({ style = 'inline', nextNumber = 1 } = {}) {
  return `<numbering xmlns="${NUMBERING_NS}" style="${escapeXml(style)}" `
    + `nextNumber="${escapeXml(nextNumber)}"/>`;
}

export function parseNumberingXml(xmlString) {
  const doc = new DOMParser().parseFromString(xmlString, 'application/xml');
  const el = doc.getElementsByTagNameNS(NUMBERING_NS, 'numbering')[0];
  if (!el) return null;
  return {
    style: el.getAttribute('style') || 'inline',
    nextNumber: parseInt(el.getAttribute('nextNumber') || '1', 10) || 1,
  };
}

// Replace the document's numbering-state part (delete-then-add).
export async function writeNumberingState(state) {
  return Word.run(async (context) => {
    const parts = context.document.customXmlParts.getByNamespace(NUMBERING_NS);
    parts.load('items');
    await context.sync();
    for (const p of parts.items) p.delete();
    context.document.customXmlParts.add(buildNumberingXml(state));
    await context.sync();
  });
}

// Read the document's numbering state, or null if none has been written yet
// (caller applies DEFAULT_NUMBERING).
export async function readNumberingState() {
  return Word.run(async (context) => {
    const parts = context.document.customXmlParts.getByNamespace(NUMBERING_NS);
    parts.load('items');
    await context.sync();
    if (!parts.items.length) return null;
    const blob = parts.items[0].getXml();
    await context.sync();
    return parseNumberingXml(blob.value);
  });
}

// All stored equation payloads, keyed by uuid — used by "Renumber all" to look up
// each tagged picture's source without one Word round-trip per picture.
export async function readAllEquations() {
  return Word.run(async (context) => {
    const parts = context.document.customXmlParts.getByNamespace(NS);
    parts.load('items');
    await context.sync();
    const blobs = parts.items.map((p) => p.getXml());
    await context.sync();
    const byUuid = {};
    for (const blob of blobs) {
      const data = parseXml(blob.value);
      if (data && data.uuid) byUuid[data.uuid] = data;
    }
    return byUuid;
  });
}
