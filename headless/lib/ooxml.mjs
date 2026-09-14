// WordprocessingML emitters for equation pictures, display paragraphs and the
// 1×2 numbering table. Pure string builders — mirrors what the add-in makes Word
// produce through Office.js (see the headless roadmap's Office.js ↔ OOXML table).
//
// [unverified against a Word-written golden fixture — headless roadmap H0.1]
// Shapes follow ECMA-376 + the MS-ODRAWXML svgBlip extension; verify byte-level
// details (effectExtent, useLocalDpi) against a document the add-in produced.

import { encodeAttr, encodeXml } from './paragraphs.mjs';

export const NS = {
  w: 'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
  r: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
  wp: 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing',
  a: 'http://schemas.openxmlformats.org/drawingml/2006/main',
  pic: 'http://schemas.openxmlformats.org/drawingml/2006/picture',
  asvg: 'http://schemas.microsoft.com/office/drawing/2016/SVG/main',
};
export const SVG_BLIP_EXT_URI = '{96DAC541-7B7A-43D3-8B79-37D633B846F1}';
export const ALT_PREFIX = 'mjx:';
export const NUMBER_CELL_WIDTH_PT = 54;           // = add-in NUMBER_CELL_WIDTH_PT

export const ptToEmu = (pt) => Math.round(pt * 12700);
export const ptToTwips = (pt) => Math.round(pt * 20);
// Word's w:position is in half-points; the add-in sets Font.position = -descentPt.
export const positionHalfPoints = (descentPt) => -Math.round(2 * descentPt);

// The <w:r> holding one equation picture: PNG blip + SVG svgBlip extension,
// alt-text title "mjx:<uuid>" (click-to-edit key) and the LaTeX as description.
// Inline equations get the baseline shift the add-in applies on desktop.
export function emitPictureRun({ pngRid, svgRid, widthPt, heightPt, descentPt = 0, uuid, latex, docPrId, displayMode = 'inline' }) {
  const cx = ptToEmu(widthPt), cy = ptToEmu(heightPt);
  const shift = displayMode === 'inline' && descentPt > 0 ? positionHalfPoints(descentPt) : 0;
  const rPr = shift ? `<w:rPr><w:noProof/><w:position w:val="${shift}"/></w:rPr>` : '<w:rPr><w:noProof/></w:rPr>';
  const title = encodeAttr(ALT_PREFIX + uuid);
  const descr = encodeAttr(latex);
  return `<w:r>${rPr}<w:drawing>`
    + `<wp:inline distT="0" distB="0" distL="0" distR="0" xmlns:wp="${NS.wp}" xmlns:r="${NS.r}">`
    + `<wp:extent cx="${cx}" cy="${cy}"/>`
    + `<wp:effectExtent l="0" t="0" r="0" b="0"/>`
    + `<wp:docPr id="${docPrId}" name="Equation ${docPrId}" title="${title}" descr="${descr}"/>`
    + `<wp:cNvGraphicFramePr><a:graphicFrameLocks xmlns:a="${NS.a}" noChangeAspect="1"/></wp:cNvGraphicFramePr>`
    + `<a:graphic xmlns:a="${NS.a}"><a:graphicData uri="${NS.pic}">`
    + `<pic:pic xmlns:pic="${NS.pic}">`
    + `<pic:nvPicPr><pic:cNvPr id="0" name="Equation ${docPrId}" descr="${descr}"/><pic:cNvPicPr/></pic:nvPicPr>`
    + `<pic:blipFill><a:blip r:embed="${pngRid}">`
    + `<a:extLst><a:ext uri="${SVG_BLIP_EXT_URI}"><asvg:svgBlip xmlns:asvg="${NS.asvg}" r:embed="${svgRid}"/></a:ext></a:extLst>`
    + `</a:blip><a:stretch><a:fillRect/></a:stretch></pic:blipFill>`
    + `<pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${cx}" cy="${cy}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr>`
    + `</pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r>`;
}

// pPr with a given justification, replacing any existing <w:jc>. Keeps the rest
// of the paragraph's properties (style, spacing) so the equation paragraph
// inherits the author's layout. Element order: w:jc precedes w:rPr in pPr.
export function withJc(pPr, val) {
  const jc = `<w:jc w:val="${val}"/>`;
  if (!pPr) return `<w:pPr>${jc}</w:pPr>`;
  if (/<w:jc\b/.test(pPr)) return pPr.replace(/<w:jc\b[^>]*\/>|<w:jc\b[^>]*>[\s\S]*?<\/w:jc>/, jc);
  if (/<w:pPr\s*\/>/.test(pPr)) return `<w:pPr>${jc}</w:pPr>`;
  const rpr = pPr.indexOf('<w:rPr');
  if (rpr >= 0) return pPr.slice(0, rpr) + jc + pPr.slice(rpr);
  const sect = pPr.indexOf('<w:sectPr');
  if (sect >= 0) return pPr.slice(0, sect) + jc + pPr.slice(sect);
  return pPr.replace(/<\/w:pPr>\s*$/, jc + '</w:pPr>');
}

// TeX's display spacing: LaTeX sets \abovedisplayskip = \belowdisplayskip to
// one body font size (10 pt at 10 pt, 11 at 11, 12 at 12; the stretch/shrink
// components have no Word equivalent). Word adds it as paragraph spacing on the
// equation paragraph (or the number-table's cell paragraphs), so the display
// sits one line-ish clear of the prose above and below, as on a TeX page.
export function displaySkipTwips(bodyPt) {
  const pt = Number.isFinite(bodyPt) && bodyPt > 0 ? bodyPt : 11;
  return ptToTwips(pt);
}

// pPr with w:spacing before/after set (twips), keeping any line-spacing
// attributes. Schema order puts w:spacing before w:ind / w:jc / w:rPr.
export function withSpacing(pPr, beforeTwips, afterTwips) {
  const attrs = `w:before="${beforeTwips}" w:after="${afterTwips}"`;
  if (!pPr || /<w:pPr\s*\/>/.test(pPr)) return `<w:pPr><w:spacing ${attrs}/></w:pPr>`;
  if (/<w:spacing\b/.test(pPr)) {
    return pPr.replace(/<w:spacing\b([^>]*?)\/?>/, (all, a) => {
      let rest = a.replace(/\s*w:(before|after|beforeLines|afterLines|beforeAutospacing|afterAutospacing)="[^"]*"/g, '');
      return `<w:spacing ${attrs}${rest}/>`;
    });
  }
  const el = `<w:spacing ${attrs}/>`;
  const m = /<w:(ind|contextualSpacing|mirrorIndents|suppressOverlap|jc|textDirection|textAlignment|outlineLvl|divId|cnfStyle|rPr|sectPr)\b/.exec(pPr);
  if (m) return pPr.slice(0, m.index) + el + pPr.slice(m.index);
  return pPr.replace(/<\/w:pPr>\s*$/, el + '</w:pPr>');
}

// pPr for the text that CONTINUES a paragraph after a display: TeX sets no
// paragraph indent there, so drop a positive first-line indent.
export function withoutFirstLineIndent(pPr) {
  return (pPr || '').replace(/(<w:ind\b[^>]*?)\s+w:firstLine="\d+"/, '$1 w:firstLine="0"');
}

// Display alignment: 'left' = LaTeX's fleqn (flush left, indented by
// \mathindent = 25 pt), 'center' = plain LaTeX. Left is the default.
export const MATH_INDENT_TWIPS = ptToTwips(25);

// pPr with the left indent set (twips), replacing any existing w:ind left/
// firstLine. Schema order: w:ind follows w:spacing and precedes w:jc.
export function withLeftIndent(pPr, twips) {
  const ind = `<w:ind w:left="${twips}" w:firstLine="0"/>`;
  if (!pPr || /<w:pPr\s*\/>/.test(pPr)) return `<w:pPr>${ind}</w:pPr>`;
  if (/<w:ind\b/.test(pPr)) return pPr.replace(/<w:ind\b[^>]*\/>/, ind);
  const m = /<w:(contextualSpacing|mirrorIndents|suppressOverlap|jc|textDirection|textAlignment|outlineLvl|divId|cnfStyle|rPr|sectPr)\b/.exec(pPr);
  if (m) return pPr.slice(0, m.index) + ind + pPr.slice(m.index);
  return pPr.replace(/<\/w:pPr>\s*$/, ind + '</w:pPr>');
}

// A display-equation paragraph: the picture run alone, aligned per `align`,
// with TeX's display skips above and below (`skipTwips`, default: none).
export function emitDisplayParagraph(runXml, pPr = '', skipTwips = 0, align = 'left') {
  let p = withJc(pPr, align === 'center' ? 'center' : 'left');
  if (align !== 'center') p = withLeftIndent(p, MATH_INDENT_TWIPS);
  if (skipTwips > 0) p = withSpacing(p, skipTwips, skipTwips);
  return `<w:p>${p}${runXml}</w:p>`;
}

// Bookmark names: letter first, [A-Za-z0-9_], ≤ 40 chars.
export function bookmarkName(label) {
  return ('eq_' + String(label).replace(/[^A-Za-z0-9_]/g, '_')).slice(0, 40);
}

// Content of the number cell / an inline "(n)": static text, or a SEQ field for
// the field style; optionally wrapped in a bookmark so [[ref:]] can target it.
export function emitNumberRuns({ number, style = 'table', label = null, bookmarkId = null }) {
  const text = `(${number})`;
  let inner = style === 'field'
    ? `<w:r><w:t>(</w:t></w:r><w:fldSimple w:instr=" SEQ equation \\* ARABIC "><w:r><w:t>${number}</w:t></w:r></w:fldSimple><w:r><w:t>)</w:t></w:r>`
    : `<w:r><w:t>${encodeXml(text)}</w:t></w:r>`;
  if (label != null && bookmarkId != null) {
    const name = bookmarkName(label);
    inner = `<w:bookmarkStart w:id="${bookmarkId}" w:name="${name}"/>${inner}<w:bookmarkEnd w:id="${bookmarkId}"/>`;
  }
  return inner;
}

// A reference to a labelled equation: static "(n)" (default) or a REF field
// whose cached result is "(n)" so it reads right before any field update.
export function emitRefRun({ number, label, field = false, rPr = '' }) {
  const text = encodeXml(`(${number})`);
  if (!field) return `<w:r>${rPr}<w:t>${text}</w:t></w:r>`;
  return `<w:fldSimple w:instr=" REF ${bookmarkName(label)} \\h "><w:r>${rPr}<w:t>${text}</w:t></w:r></w:fldSimple>`;
}

const NO_BORDERS = '<w:tblBorders><w:top w:val="nil"/><w:left w:val="nil"/><w:bottom w:val="nil"/>'
  + '<w:right w:val="nil"/><w:insideH w:val="nil"/><w:insideV w:val="nil"/></w:tblBorders>';

// The borderless 1×2 numbering table (add-in "flush right" style): equation
// centred in a wide left cell, number right-aligned in a narrow right cell.
// Width = 100 % of the column it sits in with autofit (like the add-in's
// autoFitWindow), so Word and LibreOffice fit it to a two-column layout; the
// grid is a hint computed from `textWidthTwips`, the governing section's
// COLUMN width. The number cell is 54 pt (the add-in's) on wide columns and
// 36 pt on narrow ones (< 300 pt) — "(99)" at 10 pt needs ~24 pt.
export function emitNumberTable({ runXml, numberXml, textWidthTwips, skipTwips = 0, align = 'left' }) {
  const numW = ptToTwips(textWidthTwips < 6000 ? 36 : NUMBER_CELL_WIDTH_PT);
  const eqW = Math.max(numW, textWidthTwips - numW);
  const sp = skipTwips > 0 ? `<w:spacing w:before="${skipTwips}" w:after="${skipTwips}"/>` : '';
  const eqPPr = align === 'center' ? `${sp}<w:jc w:val="center"/>` : `${sp}<w:ind w:left="${MATH_INDENT_TWIPS}"/><w:jc w:val="left"/>`;
  return '<w:tbl>'
    + `<w:tblPr><w:tblW w:w="5000" w:type="pct"/><w:jc w:val="left"/>${NO_BORDERS}<w:tblLayout w:type="autofit"/>`
    + '<w:tblCellMar><w:left w:w="0" w:type="dxa"/><w:right w:w="0" w:type="dxa"/></w:tblCellMar>'
    + '<w:tblLook w:val="0000" w:firstRow="0" w:lastRow="0" w:firstColumn="0" w:lastColumn="0" w:noHBand="0" w:noVBand="0"/></w:tblPr>'
    + `<w:tblGrid><w:gridCol w:w="${eqW}"/><w:gridCol w:w="${numW}"/></w:tblGrid>`
    + '<w:tr>'
    + `<w:tc><w:tcPr><w:tcW w:w="${eqW}" w:type="dxa"/><w:vAlign w:val="center"/></w:tcPr>`
    + `<w:p><w:pPr>${eqPPr}</w:pPr>${runXml}</w:p></w:tc>`
    + `<w:tc><w:tcPr><w:tcW w:w="${numW}" w:type="dxa"/><w:vAlign w:val="center"/></w:tcPr>`
    + `<w:p><w:pPr>${sp}<w:jc w:val="right"/></w:pPr>${numberXml}</w:p></w:tc>`
    + '</w:tr></w:tbl>';
}

// Rewrite the displayed number inside a number cell's XML (renumber). Handles
// both the static "(n)" text and a SEQ field's cached result. Returns the new
// cell XML, unchanged if no number was found.
export function rewriteNumberInCell(cellXml, number) {
  const seq = /(<w:fldSimple w:instr=" SEQ [^"]*">\s*<w:r>(?:<w:rPr>[\s\S]*?<\/w:rPr>)?<w:t(?:\s[^>]*)?>)\d+(<\/w:t>)/;
  if (seq.test(cellXml)) return cellXml.replace(seq, `$1${number}$2`);
  return cellXml.replace(/(<w:t(?:\s[^>]*)?>)\((\d+)\)(<\/w:t>)/, `$1(${number})$3`);
}

// Innermost <tag>…</tag> element enclosing string offset `pos` (nesting-aware).
// Returns { start, end } (end exclusive) or null.
export function enclosingElement(xml, pos, tag) {
  const open = new RegExp(`<${tag}(?=[\\s>])`, 'g');
  const closeTag = `</${tag}>`;
  const tokens = [];
  let m;
  while ((m = open.exec(xml)) && m.index < pos) tokens.push({ i: m.index, open: true });
  let ci = xml.indexOf(closeTag);
  while (ci >= 0 && ci < pos) { tokens.push({ i: ci, open: false }); ci = xml.indexOf(closeTag, ci + 1); }
  tokens.sort((a, b) => a.i - b.i);
  const stack = [];
  for (const t of tokens) { if (t.open) stack.push(t.i); else stack.pop(); }
  if (!stack.length) return null;
  const start = stack[stack.length - 1];
  // Find the matching close from `start` forward.
  let depth = 0;
  const re = new RegExp(`<${tag}(?=[\\s>])|</${tag}>`, 'g');
  re.lastIndex = start;
  while ((m = re.exec(xml))) {
    if (m[0].startsWith('</')) { depth -= 1; if (depth === 0) return { start, end: m.index + closeTag.length }; }
    else depth += 1;
  }
  return null;
}

// Every equation picture in `xml` in document order (tables included, since
// the walk is textual): [{ uuid, latex, docPrId, pngRid, svgRid, runStart,
// runEnd, widthPt, heightPt }]. Pictures without the mjx: title are ignored.
export function findEquationPictures(xml) {
  const out = [];
  const re = /<wp:docPr\s[^>]*title="mjx:([^"]+)"[^>]*>/g;
  let m;
  while ((m = re.exec(xml))) {
    const tag = m[0];
    const uuid = m[1];
    const descr = /descr="([^"]*)"/.exec(tag);
    const id = /\sid="(\d+)"/.exec(tag);
    const run = enclosingElement(xml, m.index, 'w:r');
    if (!run) continue;
    const runXml = xml.slice(run.start, run.end);
    const png = /<a:blip\s[^>]*r:embed="([^"]+)"/.exec(runXml);
    const svg = /<asvg:svgBlip\s[^>]*r:embed="([^"]+)"/.exec(runXml);
    const ext = /<wp:extent cx="(\d+)" cy="(\d+)"/.exec(runXml);
    out.push({
      uuid,
      latex: descr ? decodeAttr(descr[1]) : '',
      docPrId: id ? parseInt(id[1], 10) : 0,
      pngRid: png ? png[1] : null,
      svgRid: svg ? svg[1] : null,
      runStart: run.start,
      runEnd: run.end,
      widthPt: ext ? parseInt(ext[1], 10) / 12700 : 0,
      heightPt: ext ? parseInt(ext[2], 10) / 12700 : 0,
    });
  }
  return out;
}

function decodeAttr(s) {
  return s.replace(/&(lt|gt|amp|quot|apos|#x[0-9a-fA-F]+|#\d+);/g, (m, e) => {
    switch (e) {
      case 'lt': return '<'; case 'gt': return '>'; case 'amp': return '&';
      case 'quot': return '"'; case 'apos': return "'";
      default: return String.fromCodePoint(e[1] === 'x' ? parseInt(e.slice(2), 16) : parseInt(e.slice(1), 10));
    }
  });
}

export function nextDocPrId(xml) {
  let max = 0;
  for (const m of xml.matchAll(/<wp:docPr\b[^>]*?\sid="(\d+)"/g)) max = Math.max(max, parseInt(m[1], 10));
  return max + 1;
}

export function nextBookmarkId(xml) {
  let max = -1;
  for (const m of xml.matchAll(/<w:bookmarkStart\s[^>]*w:id="(\d+)"/g)) max = Math.max(max, parseInt(m[1], 10));
  return max + 1;
}

// Column width (twips) available to text at string offset `pos`: the section
// governing a paragraph is the first <w:sectPr> that FOLLOWS it (a paragraph-
// level break, or the body-level one at the end). Page width − margins,
// divided among the section's columns. Falls back to US Letter with 1in
// margins (9360) when unspecified; missing pgSz/pgMar inherit from the last
// section that has them (Word's behaviour for paragraph-level breaks).
export function sectionTextWidthAt(xml, pos) {
  const sects = [];
  const re = /<w:sectPr(?=[\s>])[\s\S]*?<\/w:sectPr>|<w:sectPr[^>]*\/>/g;
  let m;
  while ((m = re.exec(xml))) sects.push({ at: m.index, xml: m[0] });
  if (!sects.length) return 9360;
  let idx = sects.findIndex((s) => s.at >= pos);
  if (idx < 0) idx = sects.length - 1;
  const pick = (attrRe, dflt) => {
    for (let i = idx; i < sects.length; i++) { const r = attrRe.exec(sects[i].xml); if (r) return parseInt(r[1], 10); }
    for (let i = idx - 1; i >= 0; i--) { const r = attrRe.exec(sects[i].xml); if (r) return parseInt(r[1], 10); }
    return dflt;
  };
  const w = pick(/<w:pgSz\s[^>]*w:w="(\d+)"/, 12240);
  const l = pick(/<w:pgMar\s[^>]*w:left="(\d+)"/, 1440);
  const r = pick(/<w:pgMar\s[^>]*w:right="(\d+)"/, 1440);
  const cols = /<w:cols\s[^>]*w:num="(\d+)"/.exec(sects[idx].xml);
  const n = cols ? Math.max(1, parseInt(cols[1], 10)) : 1;
  const spaceM = /<w:cols\s[^>]*w:space="(\d+)"/.exec(sects[idx].xml);
  const space = n > 1 ? (spaceM ? parseInt(spaceM[1], 10) : 720) : 0;
  const text = Math.max(2000, w - l - r);
  return Math.max(1000, Math.floor((text - space * (n - 1)) / n));
}

// Text width of the document's last section (single-column figure), kept for
// callers that need one number for the whole document.
export function sectionTextWidthTwips(xml) {
  return sectionTextWidthAt(xml, xml.length);
}

// Default body font size (pt) from styles.xml docDefaults (w:sz half-points).
export function defaultFontSizePt(stylesXml) {
  const dd = /<w:docDefaults>[\s\S]*?<\/w:docDefaults>/.exec(stylesXml || '');
  const m = dd && /<w:sz\s+w:val="(\d+)"/.exec(dd[0]);
  return m ? parseInt(m[1], 10) / 2 : 11;
}
