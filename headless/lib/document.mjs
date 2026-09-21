// Equation operations on a DocxPackage: the placeholder pass (`process`), and
// list / update / renumber / check. This is the headless twin of the add-in's
// performInsert / readAllEquations / onRenumberAll — same payload XML (storage.js),
// same alt-text tags, same numbering rules (numbering.js).
import { DOMParser } from 'linkedom';
globalThis.DOMParser ??= DOMParser;   // storage.parseXml uses the browser global

import { readFileSync } from 'node:fs';
import { randomUUID } from 'node:crypto';
import {
  NS, NUMBERING_NS, PREAMBLE_NS, buildXml, parseXml, buildNumberingXml, parseNumberingXml,
  buildPreambleXml, parsePreambleXml, embedSourceInSvg,
} from '../../core/storage.js';
import { decorateLatexForNumber, assignNumbers, numberPlacement, DEFAULT_NUMBERING, NUMBERING_STYLES } from '../../core/numbering.js';
import { buildDegradedPayload } from '../../core/recovery.js';
import { KNOWN_FONTS, DEFAULT_SETTINGS } from '../../core/settings.js';
import { renderEquation } from './render.mjs';
import { svgToPng } from './raster.mjs';
import {
  findLeafParagraphs, parseParagraph, serializeParagraph, joinedText, spliceParagraph, runSizePt, textOutside,
} from './paragraphs.mjs';
import { findPlaceholders, validatePlaceholders } from './placeholders.mjs';
import {
  emitPictureRun, emitDisplayParagraph, emitNumberTable, emitNumberRuns, emitRefRun, rewriteNumberInCell,
  enclosingElement, findEquationPictures, nextDocPrId, nextBookmarkId, sectionTextWidthAt, defaultFontSizePt,
  bookmarkName, displaySkipTwips, withoutFirstLineIndent,
} from './ooxml.mjs';

const pkgJson = JSON.parse(readFileSync(new URL('../../package.json', import.meta.url), 'utf8'));
export const WRITER = `mjx-docx/${pkgJson.version}`;

// ---- stored state (custom XML parts) ----

export function readEquationParts(pkg) {
  const byUuid = {};
  const items = [];
  for (const it of pkg.listCustomXml()) {
    if (it.namespace !== NS) continue;
    const data = parseXml(it.xml);
    if (data && data.uuid) { byUuid[data.uuid] = data; items.push({ index: it.index, uuid: data.uuid }); }
  }
  return { byUuid, items };
}

export function writeEquationPart(pkg, payload) {
  const { items } = readEquationParts(pkg);
  for (const it of items) if (it.uuid === payload.uuid) pkg.removeCustomXml(it.index);
  pkg.addCustomXml(buildXml({ ...payload, writer: WRITER }), NS);
}

export function readNumberingState(pkg) {
  const it = pkg.listCustomXml().find((x) => x.namespace === NUMBERING_NS);
  return it ? parseNumberingXml(it.xml) || { ...DEFAULT_NUMBERING } : null;
}

export function writeNumberingState(pkg, state) {
  for (const it of pkg.listCustomXml()) if (it.namespace === NUMBERING_NS) pkg.removeCustomXml(it.index);
  pkg.addCustomXml(buildNumberingXml(state), NUMBERING_NS);
}

export function readPreamble(pkg) {
  const it = pkg.listCustomXml().find((x) => x.namespace === PREAMBLE_NS);
  return it ? parsePreambleXml(it.xml) || '' : '';
}

export function writePreamble(pkg, text) {
  for (const it of pkg.listCustomXml()) if (it.namespace === PREAMBLE_NS) pkg.removeCustomXml(it.index);
  if (text && text.trim()) pkg.addCustomXml(buildPreambleXml(text), PREAMBLE_NS);
}

// ---- list ----

// Equations in document order, each joined with its stored payload (or a
// degraded payload recovered from the picture's alt text, like the add-in's
// recoverFromAltText). `inTable` tells whether the picture sits in a table.
export function listEquations(pkg) {
  const xml = pkg.text(pkg.mainPart);
  const { byUuid } = readEquationParts(pkg);
  return findEquationPictures(xml).map((pic) => {
    const stored = byUuid[pic.uuid];
    const payload = stored || buildDegradedPayload({ uuid: pic.uuid, latex: pic.latex, sizePt: 11 });
    const inTable = !!enclosingElement(xml, pic.runStart, 'w:tbl');
    return { ...payload, uuid: pic.uuid, docPrId: pic.docPrId, pngRid: pic.pngRid, svgRid: pic.svgRid,
      inTable, hasPart: !!stored, hasSvg: !!pic.svgRid };
  });
}

// ---- rendering + media ----

async function renderForDocument(payload, preamble) {
  const renderLatex = decorateLatexForNumber(payload.latex, payload.number, payload.numberStyle);
  const r = await renderEquation({ ...payload, latex: renderLatex, preamble });
  if (!r.svg) throw new Error(`render failed for "${payload.latex}": ${r.error}`);
  const svg = embedSourceInSvg(r.svg, payload);
  const { png } = await svgToPng(svg, r.widthPt, r.heightPt);
  return { svg, png, widthPt: r.widthPt, heightPt: r.heightPt, descentPt: r.descentPt, error: r.error };
}

// ---- process (placeholder pass) ----

export function normalizeOptions(opts = {}) {
  const font = opts.font || DEFAULT_SETTINGS.font;   // Termes ≈ newtxmath, like the pane
  if (!KNOWN_FONTS.includes(font)) throw new Error(`unknown font "${font}" (known: ${KNOWN_FONTS.join(', ')})`);
  const numberStyle = opts.numberStyle || 'table';
  if (!NUMBERING_STYLES.includes(numberStyle)) throw new Error(`unknown numbering style "${numberStyle}" (inline|table|field)`);
  const sizePt = opts.sizePt == null ? null : Number(opts.sizePt);
  if (sizePt != null && !(sizePt >= 4 && sizePt <= 96)) throw new Error('--size must be 4…96 pt');
  const color = opts.color || '#000000';
  if (!/^#[0-9a-fA-F]{6}$/.test(color)) throw new Error('--color must be #rrggbb');
  const align = opts.align || 'left';
  if (!['left', 'center'].includes(align)) throw new Error(`unknown alignment "${align}" (left|center)`);
  return { font, numberStyle, sizePt, color, align, preamble: opts.preamble || null, dollar: !!opts.dollar,
    staticRefs: !!opts.staticRefs, force: !!opts.force };
}

// Replace every placeholder in the main document part. Mutates `pkg`. Returns a
// report { equations, refs, warnings, errors } — `errors` non-empty means the
// document was NOT modified.
export async function processDocument(pkg, rawOpts = {}) {
  const opts = normalizeOptions(rawOpts);
  const report = { equations: 0, refs: 0, warnings: [], errors: [] };
  let xml = pkg.text(pkg.mainPart);

  if (!opts.force && /<w:(ins|del)\b/.test(xml)) {
    report.errors.push('document has tracked changes (w:ins/w:del); accept them first or pass --force');
    return report;
  }
  for (const name of pkg.names()) {
    if (/^word\/(header|footer)\d*\.xml$/.test(name) && /\[\[(math|display|eq|ref)\b/.test(pkg.text(name))) {
      report.warnings.push(`${name}: placeholders in headers/footers are not processed`);
    }
  }

  const preamble = opts.preamble != null ? opts.preamble : readPreamble(pkg);
  const stylesDefaultPt = defaultFontSizePt(pkg.text('word/styles.xml'));
  let docPrId = nextDocPrId(xml);
  let bookmarkId = nextBookmarkId(xml);
  const numbering = readNumberingState(pkg) || { ...DEFAULT_NUMBERING };
  const style = opts.numberStyle;

  // Phase A — find every placeholder (document order) to assign numbers and
  // resolve labels before any XML is touched, so [[ref:]] can precede its target.
  const paras = findLeafParagraphs(xml);
  const work = [];
  const labels = new Map();
  let nextNumber = numbering.nextNumber;
  // Field-style equations are counted by Word's SEQ sequence, not by our
  // static counter (same split as Renumber-all), so static numbers stay
  // contiguous. Existing SEQ fields in the document come first.
  let nextField = (xml.match(/ SEQ equation /g) || []).length + 1;
  for (const para of paras) {
    const p = parseParagraph(para.xml);
    const { text } = joinedText(p);
    if (!text.includes('[[') && !(opts.dollar && text.includes('$'))) continue;
    const found = findPlaceholders(text, { dollar: opts.dollar });
    if (!found.length) continue;
    const bad = validatePlaceholders(found);
    if (bad) { report.errors.push(`${bad} — in paragraph: "${excerpt(text)}"`); continue; }
    for (const ph of found) {
      if (ph.kind === 'eq') {
        ph.number = (ph.options && ph.options.style ? ph.options.style : style) === 'field' ? nextField++ : nextNumber++;
        if (ph.label) {
          if (labels.has(ph.label)) report.errors.push(`duplicate label "${ph.label}"`);
          labels.set(ph.label, ph.number);
        }
      }
    }
    work.push({ para, p, text, found });
  }
  for (const w of work) for (const ph of w.found) {
    if (ph.kind === 'ref' && !labels.has(ph.latex)) {
      report.errors.push(`unknown label "${ph.latex}" in [[ref: …]] — in paragraph: "${excerpt(w.text)}"`);
    }
  }
  if (report.errors.length) return report;
  if (!work.length) { report.warnings.push('no placeholders found'); return report; }

  // Phase B — render and splice, last paragraph first so offsets stay valid.
  const payloads = [];
  for (const w of [...work].reverse()) {
    const { para, p, text, found } = w;
    const { map } = joinedText(p);
    const edits = [];
    const blocks = [];        // block-level XML keyed by marker index
    for (const ph of found) {
      if (ph.kind === 'ref') {
        const number = labels.get(ph.latex);
        const rPr = runRPrAt(p, map, ph.start);
        edits.push({ start: ph.start, end: ph.end, replacement: [emitRefRun({ number, label: ph.latex, field: !opts.staticRefs, rPr })] });
        report.refs += 1;
        continue;
      }
      const displayMode = ph.kind === 'math' ? 'inline' : 'display';
      const numbered = ph.kind === 'eq';
      const o = ph.options || {};
      if (o.font && !KNOWN_FONTS.includes(o.font)) { report.errors.push(`unknown font "${o.font}" in placeholder "${excerpt(ph.latex)}"`); continue; }
      // The display skip follows the BODY text size around the equation (TeX:
      // the font class), not the equation's own size option.
      const bodyPt = runSizeAt(p, map, ph.start) ?? stylesDefaultPt;
      const skip = displaySkipTwips(bodyPt);
      const sizePt = o.size ? Number(o.size) : (opts.sizePt ?? bodyPt);
      const eqStyle = numbered ? (o.style || style) : 'inline';
      const eqAlign = o.align || opts.align;
      const payload = {
        uuid: randomUUID(), latex: ph.latex, font: o.font || opts.font, sizePt, displayMode, color: o.color || opts.color,
        numbered, number: numbered ? ph.number : 0, numberStyle: eqStyle,
      };
      const r = await renderForDocument(payload, preamble);
      if (r.error) report.warnings.push(`TeX error in "${excerpt(ph.latex)}": ${r.error} (inserted with the error mark)`);
      const svgRef = pkg.addMedia('svg', Buffer.from(r.svg, 'utf8'));
      const pngRef = pkg.addMedia('png', r.png);
      const runXml = emitPictureRun({
        pngRid: pngRef.rId, svgRid: svgRef.rId, widthPt: r.widthPt, heightPt: r.heightPt, descentPt: r.descentPt,
        uuid: payload.uuid, latex: payload.latex, docPrId: docPrId++, displayMode,
      });
      payloads.push(payload);
      report.equations += 1;
      if (displayMode === 'inline') {
        edits.push({ start: ph.start, end: ph.end, replacement: [runXml] });
        continue;
      }
      let block;
      if (!numbered || numberPlacement(eqStyle) === 'svg') {
        block = emitDisplayParagraph(runXml, withoutFirstLineIndent(stripSectPr(p.pPr)), skip, eqAlign);
      } else {
        const numberXml = emitNumberRuns({ number: ph.number, style: eqStyle, label: ph.label, bookmarkId: ph.label ? bookmarkId++ : null });
        block = emitNumberTable({ runXml, numberXml, textWidthTwips: sectionTextWidthAt(xml, para.start), skipTwips: skip, align: eqAlign });
      }
      if (numbered && numberPlacement(eqStyle) === 'svg' && ph.label) {
        // Inline-number style: the number is inside the picture; bookmark the
        // whole equation paragraph so a REF field still resolves to something.
        block = block.replace(/<w:p>/, `<w:p><w:bookmarkStart w:id="${bookmarkId}" w:name="${bookmarkName(ph.label)}"/>`)
          .replace(/<\/w:p>$/, `<w:bookmarkEnd w:id="${bookmarkId}"/></w:p>`);
        bookmarkId += 1;
      }
      const marker = `<mjx:block idx="${blocks.length}"/>`;
      blocks.push(block);
      edits.push({ start: ph.start, end: ph.end, replacement: [marker] });
    }
    const spliced = serializeParagraph(spliceParagraph(p, edits));
    const after = xml.slice(para.end).replace(/^\s+/, '');
    const replacement = blocks.length
      ? splitAtBlocks(spliced, blocks, p.pPr, after)
      : spliced;
    xml = xml.slice(0, para.start) + replacement + xml.slice(para.end);
  }

  pkg.setText(pkg.mainPart, xml);
  for (const payload of payloads) writeEquationPart(pkg, payload);
  writeNumberingState(pkg, { style, nextNumber });
  if (opts.preamble != null) writePreamble(pkg, opts.preamble);
  return report;
}

function excerpt(s) { const t = String(s).replace(/\s+/g, ' ').trim(); return t.length > 60 ? t.slice(0, 57) + '…' : t; }

function runSizeAt(p, map, offset) {
  const loc = map[offset];
  const c = loc && p.children[loc.childIndex];
  return c && c.type === 'run' ? runSizePt(c) : null;
}
// The formatting of the run a [[ref:]] sits in, minus character style and font:
// a reference should look like the prose around it (bold/italic/size), not like
// the code span a pandoc author had to wrap the placeholder in.
function runRPrAt(p, map, offset) {
  const loc = map[offset];
  const c = loc && p.children[loc.childIndex];
  if (!c || c.type !== 'run' || !c.rPr) return '';
  const inner = c.rPr.replace(/^<w:rPr[^>]*>|<\/w:rPr>$/g, '').replace(/<w:rStyle\b[^>]*\/>|<w:rFonts\b[^>]*\/>/g, '');
  return inner ? `<w:rPr>${inner}</w:rPr>` : '';
}
function stripSectPr(pPr) {
  return pPr.replace(/<w:sectPr\b[\s\S]*?<\/w:sectPr>|<w:sectPr\b[^>]*\/>/, '');
}

// A spliced paragraph containing <mjx:block idx="n"/> markers becomes a
// sequence: paragraph-piece, block, paragraph-piece, … Pieces that hold no
// content are dropped, except the last one when the original paragraph carried
// the section properties (they must survive in a paragraph). Text after a
// block continues the paragraph, so (as in TeX) it gets no first-line indent.
// A block that ends up directly before a table, a cell end or the section end
// is followed by an empty paragraph, as WordprocessingML requires.
function splitAtBlocks(pxml, blocks, pPr, after) {
  const openEnd = pxml.indexOf('>') + 1;
  const open = pxml.slice(0, openEnd);
  const bodyStart = pxml.startsWith(pPr, openEnd) ? openEnd + pPr.length : openEnd;
  const body = pxml.slice(bodyStart, pxml.lastIndexOf('</w:p>'));
  const pieces = body.split(/<mjx:block idx="(\d+)"\/>/);
  const hasSect = /<w:sectPr\b/.test(pPr);
  const plainPPr = stripSectPr(pPr);
  const out = [];
  for (let i = 0; i < pieces.length; i++) {
    if (i % 2 === 1) { out.push(blocks[parseInt(pieces[i], 10)]); continue; }
    const content = pieces[i];
    const isLast = i === pieces.length - 1;
    const empty = !/<w:t\b|<w:drawing\b|<w:fldSimple\b|<w:hyperlink\b|<w:tab\/>|<w:br\b/.test(content);
    if (empty && !(isLast && hasSect)) continue;
    const pp = isLast ? pPr : plainPPr;
    out.push(open + (i > 0 ? withoutFirstLineIndent(pp) : pp) + content + '</w:p>');
  }
  const last = out[out.length - 1];
  const needsTrailingP = last.startsWith('<w:tbl>') && !/^<w:p\b/.test(after);
  if (needsTrailingP) out.push(hasSect ? `<w:p>${pPr}</w:p>` : '<w:p/>');
  return out.join('');
}

// ---- update ----

// Re-render one equation in place with changed fields (latex/font/sizePt/color),
// keeping its number, style, position and picture identity.
export async function updateEquation(pkg, uuid, changes = {}) {
  let xml = pkg.text(pkg.mainPart);
  const pic = findEquationPictures(xml).find((p) => p.uuid === uuid);
  if (!pic) throw new Error(`no equation picture tagged ${uuid}`);
  const { byUuid } = readEquationParts(pkg);
  const base = byUuid[uuid] || buildDegradedPayload({ uuid, latex: pic.latex, sizePt: 11 });
  const payload = { ...base, ...changes, uuid };
  delete payload.degraded;
  if (!KNOWN_FONTS.includes(payload.font)) payload.font = DEFAULT_SETTINGS.font;
  const r = await renderForDocument(payload, readPreamble(pkg));
  const pngPath = pic.pngRid && pkg.mediaPathFor(pic.pngRid);
  const svgPath = pic.svgRid && pkg.mediaPathFor(pic.svgRid);
  let pngRid = pic.pngRid, svgRid = pic.svgRid;
  if (pngPath && pngPath.endsWith('.png')) pkg.setBytes(pngPath, r.png); else ({ rId: pngRid } = pkg.addMedia('png', r.png));
  if (svgPath && svgPath.endsWith('.svg')) pkg.setBytes(svgPath, Buffer.from(r.svg, 'utf8')); else ({ rId: svgRid } = pkg.addMedia('svg', Buffer.from(r.svg, 'utf8')));
  const runXml = emitPictureRun({
    pngRid, svgRid, widthPt: r.widthPt, heightPt: r.heightPt, descentPt: r.descentPt,
    uuid, latex: payload.latex, docPrId: pic.docPrId, displayMode: payload.displayMode,
  });
  xml = xml.slice(0, pic.runStart) + runXml + xml.slice(pic.runEnd);
  pkg.setText(pkg.mainPart, xml);
  writeEquationPart(pkg, payload);
  return { payload, error: r.error };
}

// ---- renumber ----

// Mirror of the add-in's Renumber-all: static numbers (inline + table styles)
// are reassigned 1..n in document order; SEQ fields get sequential cached
// results (Word recomputes them anyway). Returns { changed, fields }.
export async function renumberDocument(pkg) {
  const eqs = listEquations(pkg);
  const staticEqs = eqs.filter((e) => numberPlacement(e.numberStyle) !== 'field');
  const changed = assignNumbers(staticEqs.map((e) => ({ uuid: e.uuid, numbered: e.numbered, number: e.number })));
  const byUuid = Object.fromEntries(eqs.map((e) => [e.uuid, e]));
  let done = 0;
  for (const { uuid, number } of changed) {
    const eq = byUuid[uuid];
    if (numberPlacement(eq.numberStyle) === 'cell') {
      if (rewriteCellNumber(pkg, uuid, number)) { writeEquationPart(pkg, { ...eq, number }); done += 1; }
    } else {
      await updateEquation(pkg, uuid, { number });
      done += 1;
    }
  }
  // Field-style equations: sequential cached results in document order.
  let fields = 0;
  const fieldEqs = eqs.filter((e) => e.numbered && numberPlacement(e.numberStyle) === 'field');
  fieldEqs.forEach((eq, i) => { if (rewriteCellNumber(pkg, eq.uuid, i + 1)) fields += 1; });
  const staticNumbered = staticEqs.filter((e) => e.numbered).length;
  const state = readNumberingState(pkg) || { ...DEFAULT_NUMBERING };
  writeNumberingState(pkg, { ...state, nextNumber: staticNumbered + 1 });
  // REF fields carry a cached "(n)"; refresh them so the file reads right even
  // before Word updates fields.
  const xml = pkg.text(pkg.mainPart);
  const { updated, refs } = rewriteRefFields(xml, labelNumbers(xml, readEquationParts(pkg).byUuid));
  if (refs) pkg.setText(pkg.mainPart, updated);
  return { changed: done, fields, refs };
}

// Bookmark name → the number currently shown for it. The bookmark wraps the
// number cell content (table/field styles) or the whole equation paragraph
// (inline style, where the number lives in the picture → stored payload).
export function labelNumbers(xml, byUuid = {}) {
  const map = new Map();
  for (const m of xml.matchAll(/<w:bookmarkStart w:id="(\d+)" w:name="(eq_[^"]*)"\/>/g)) {
    const rest = xml.slice(m.index);
    const endAt = rest.search(new RegExp(`<w:bookmarkEnd w:id="${m[1]}"/>`));
    const span = rest.slice(0, endAt < 0 ? 8000 : endAt);
    const seq = /<w:fldSimple w:instr=" SEQ [^"]*">\s*<w:r>(?:<w:rPr>[\s\S]*?<\/w:rPr>)?<w:t(?:\s[^>]*)?>(\d+)</.exec(span);
    const stat = /<w:t(?:\s[^>]*)?>\((\d+)\)</.exec(span);
    const pic = /title="mjx:([^"]+)"/.exec(span);
    const n = seq ? +seq[1] : stat ? +stat[1] : pic && byUuid[pic[1]] ? byUuid[pic[1]].number : 0;
    if (n > 0) map.set(m[2], n);
  }
  return map;
}

// Rewrite the cached result of every REF field that targets a known bookmark.
export function rewriteRefFields(xml, map) {
  let refs = 0;
  const updated = xml.replace(
    /(<w:fldSimple w:instr=" REF (eq_[^ "]+) \\h ">\s*<w:r>(?:<w:rPr>[\s\S]*?<\/w:rPr>)?<w:t(?:\s[^>]*)?>)\((\d+)\)(<\/w:t>)/g,
    (all, pre, name, old, post) => {
      if (!map.has(name)) return all;
      if (map.get(name) !== +old) refs += 1;
      return `${pre}(${map.get(name)})${post}`;
    });
  return { updated, refs };
}

function rewriteCellNumber(pkg, uuid, number) {
  let xml = pkg.text(pkg.mainPart);
  const pic = findEquationPictures(xml).find((p) => p.uuid === uuid);
  if (!pic) return false;
  const tbl = enclosingElement(xml, pic.runStart, 'w:tbl');
  if (!tbl) return false;
  const tblXml = xml.slice(tbl.start, tbl.end);
  // The number cell is the cell AFTER the one holding the picture.
  const cellOfPic = enclosingElement(xml, pic.runStart, 'w:tc');
  if (!cellOfPic) return false;
  const nextOpen = xml.indexOf('<w:tc>', cellOfPic.end) >= 0 ? xml.indexOf('<w:tc>', cellOfPic.end) : xml.indexOf('<w:tc ', cellOfPic.end);
  if (nextOpen < 0 || nextOpen >= tbl.end) return false;
  const numCell = enclosingElement(xml, nextOpen + 1, 'w:tc');
  if (!numCell) return false;
  const cellXml = xml.slice(numCell.start, numCell.end);
  const rewritten = rewriteNumberInCell(cellXml, number);
  if (rewritten === cellXml) return false;
  xml = xml.slice(0, numCell.start) + rewritten + xml.slice(numCell.end);
  pkg.setText(pkg.mainPart, xml);
  void tblXml;
  return true;
}

// ---- check ----

export function checkDocument(pkg) {
  const issues = [];
  const xml = pkg.text(pkg.mainPart);
  const ct = pkg.text('[Content_Types].xml');
  const pics = findEquationPictures(xml);
  const { byUuid, items } = readEquationParts(pkg);
  const rels = pkg.readRels(pkg.mainPart);
  const relIds = new Set(rels.map((r) => r.id));
  const seen = new Set();
  for (const p of pics) {
    if (seen.has(p.uuid)) issues.push({ level: 'error', msg: `duplicate equation uuid ${p.uuid} (click-to-edit will hit the first)` });
    seen.add(p.uuid);
    if (!byUuid[p.uuid]) issues.push({ level: 'warn', msg: `picture ${p.uuid} has no stored source part (add-in will recover from alt text)` });
    for (const [what, rid] of [['png', p.pngRid], ['svg', p.svgRid]]) {
      if (!rid) { if (what === 'svg') issues.push({ level: 'info', msg: `picture ${p.uuid} has no SVG (raster-only, e.g. Word web)` }); continue; }
      if (!relIds.has(rid)) { issues.push({ level: 'error', msg: `picture ${p.uuid}: ${what} relationship ${rid} missing` }); continue; }
      const path = pkg.mediaPathFor(rid);
      if (!pkg.has(path)) issues.push({ level: 'error', msg: `picture ${p.uuid}: media ${path} missing from package` });
    }
  }
  for (const it of items) if (!seen.has(it.uuid)) issues.push({ level: 'warn', msg: `orphan source part for ${it.uuid} (no picture in the document)` });
  if (pics.some((p) => p.svgRid) && !/Extension="svg"/i.test(ct)) issues.push({ level: 'error', msg: 'content type for .svg missing' });
  if (pics.some((p) => p.pngRid) && !/Extension="png"/i.test(ct)) issues.push({ level: 'error', msg: 'content type for .png missing' });
  const ids = [...xml.matchAll(/<wp:docPr\b[^>]*?\sid="(\d+)"/g)].map((m) => m[1]);
  if (new Set(ids).size !== ids.length) issues.push({ level: 'error', msg: 'duplicate wp:docPr ids (Word may repair the file on open)' });
  for (const r of rels) {
    if (r.mode === 'External') continue;
    if (!pkg.has(pkg.resolveTarget(pkg.mainPart, r.target))) issues.push({ level: 'error', msg: `dangling relationship ${r.id} → ${r.target}` });
  }
  if (/\[\[(math|display|eq|ref)(#[\w:.-]+)?:/.test(xml)) issues.push({ level: 'warn', msg: 'unprocessed placeholders remain (run `process`)' });
  const numbered = pics.map((p) => byUuid[p.uuid]).filter((e) => e && e.numbered && numberPlacement(e.numberStyle) !== 'field');
  const expected = numbered.map((_, i) => i + 1);
  if (numbered.some((e, i) => e.number !== expected[i])) issues.push({ level: 'info', msg: 'static equation numbers are out of document order (run `renumber`)' });
  const { refs } = rewriteRefFields(xml, labelNumbers(xml, byUuid));
  if (refs) issues.push({ level: 'info', msg: `${refs} equation reference(s) show a stale number (run \`renumber\`)` });
  for (const m of xml.matchAll(/ REF (eq_[^ "]+) /g)) {
    if (!xml.includes(`w:name="${m[1]}"`)) issues.push({ level: 'error', msg: `reference to missing equation bookmark ${m[1]}` });
  }
  return issues;
}
