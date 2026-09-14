// String-level WordprocessingML paragraph/run model. Pure.
//
// Word and docx-js fragment a paragraph's text across <w:r> runs unpredictably
// (spell-check state, formatting, revision marks), so placeholders are matched on
// the JOINED text of a paragraph and then spliced back into runs, splitting the
// runs at the match boundaries while keeping each run's own <w:rPr>.
//
// We deliberately work on the XML *string*, not a DOM: Word is fussy about
// re-serialized XML (self-closing forms, attribute order, entity choices), so
// every byte we did not intend to change stays exactly as it was.

const P_OPEN = /<w:p(?=[\s>\/])/g;

// Leaf paragraphs: every <w:p>…</w:p> that contains no nested <w:p> (nesting
// occurs inside text boxes). Returns [{ start, end, xml }] in document order.
// Self-closing <w:p/> are skipped (nothing to match in them).
export function findLeafParagraphs(xml) {
  const out = [];
  const re = /<w:p(?=[\s>])[^>]*?(\/?)>/g;
  let m;
  while ((m = re.exec(xml))) {
    if (m[1] === '/') continue;
    const start = m.index;
    const close = xml.indexOf('</w:p>', re.lastIndex);
    if (close < 0) break;
    const inner = xml.slice(re.lastIndex, close);
    P_OPEN.lastIndex = 0;
    if (P_OPEN.test(inner)) continue;        // has nested paragraphs → not a leaf
    const end = close + '</w:p>'.length;
    out.push({ start, end, xml: xml.slice(start, end) });
    re.lastIndex = end;
  }
  return out;
}

// Split a paragraph into its open tag, optional <w:pPr>, a sequence of top-level
// children (runs and everything else, e.g. <w:hyperlink>, <w:bookmarkStart/>)
// and the close tag. Children are kept as opaque strings except runs, which are
// parsed (see parseRun). Runs directly inside hyperlinks/ins/sdt are NOT
// descended into — a placeholder in a hyperlink is left alone by design.
export function parseParagraph(pxml) {
  const openEnd = pxml.indexOf('>') + 1;
  const open = pxml.slice(0, openEnd);
  const closeStart = pxml.lastIndexOf('</w:p>');
  let body = pxml.slice(openEnd, closeStart);
  let pPr = '';
  const pprM = /^\s*<w:pPr(?:\s[^>]*)?>[\s\S]*?<\/w:pPr>|^\s*<w:pPr(?:\s[^>]*)?\/>/.exec(body);
  if (pprM) { pPr = pprM[0]; body = body.slice(pprM[0].length); }
  const children = [];
  let i = 0;
  while (i < body.length) {
    if (body.startsWith('<w:r>', i) || body.startsWith('<w:r ', i)) {
      const end = body.indexOf('</w:r>', i);
      if (end < 0) { children.push({ type: 'raw', xml: body.slice(i) }); break; }
      const rxml = body.slice(i, end + '</w:r>'.length);
      children.push({ type: 'run', ...parseRun(rxml) });
      i = end + '</w:r>'.length;
    } else {
      // Advance to the next top-level run start (if any); keep the rest opaque.
      const next = nextRunStart(body, i);
      children.push({ type: 'raw', xml: body.slice(i, next < 0 ? body.length : next) });
      i = next < 0 ? body.length : next;
    }
  }
  return { open, pPr, children, close: '</w:p>' };
}

function nextRunStart(body, from) {
  const re = /<w:r(?=[\s>])/g;
  re.lastIndex = from;
  const m = re.exec(body);
  return m ? m.index : -1;
}

// A run: open tag attrs, optional <w:rPr>, and content segments — each <w:t>
// as { type:'t', text } (XML-decoded), anything else as { type:'raw', xml }.
export function parseRun(rxml) {
  const openEnd = rxml.indexOf('>') + 1;
  const attrs = rxml.slice('<w:r'.length, openEnd - 1);
  let body = rxml.slice(openEnd, rxml.lastIndexOf('</w:r>'));
  let rPr = '';
  const m = /^\s*<w:rPr(?:\s[^>]*)?>[\s\S]*?<\/w:rPr>|^\s*<w:rPr(?:\s[^>]*)?\/>/.exec(body);
  if (m) { rPr = m[0]; body = body.slice(m[0].length); }
  const segments = [];
  const re = /<w:t(?:\s[^>]*)?>([\s\S]*?)<\/w:t>|<w:t(?:\s[^>]*)?\/>/g;
  let last = 0, t;
  while ((t = re.exec(body))) {
    if (t.index > last) segments.push({ type: 'raw', xml: body.slice(last, t.index) });
    segments.push({ type: 't', text: decodeXml(t[1] ?? '') });
    last = re.lastIndex;
  }
  if (last < body.length) segments.push({ type: 'raw', xml: body.slice(last) });
  return { attrs, rPr, segments };
}

export function serializeRun(run) {
  const parts = run.segments.map((s) => (s.type === 't'
    ? `<w:t xml:space="preserve">${encodeXml(s.text)}</w:t>`
    : s.xml));
  return `<w:r${run.attrs}>${run.rPr}${parts.join('')}</w:r>`;
}

export function serializeParagraph(p) {
  return p.open + p.pPr + p.children.map((c) => (c.type === 'run' ? serializeRun(c) : c.xml)).join('') + p.close;
}

// Children/segments that carry no visible content and must neither block a
// match nor be dropped by a splice: whitespace between elements, bookmarks,
// proofing marks, comment anchors, rendered-page-break hints.
function isInert(xml) {
  // Anything that can show up on the page counts as content; bare container
  // tags (hyperlink/sdt/ins open+close), bookmarks and proofing marks do not.
  return !/<w:(t|tab|br|cr|drawing|pict|object|sym|fldChar|instrText|footnoteReference|endnoteReference|ptab|noBreakHyphen|softHyphen|delText)\b/.test(xml);
}

// The paragraph's text as Word would show it, with U+FFFC for every non-text
// segment (tab, break, drawing, field char…) and for every content-bearing
// non-run child (hyperlink, field, sdt), plus a map text offset → location.
export function joinedText(p) {
  let text = '';
  const map = [];
  p.children.forEach((c, ci) => {
    if (c.type !== 'run') {
      if (!isInert(c.xml)) { text += '￼'; map.push({ childIndex: ci, segIndex: -1, offsetInSeg: 0 }); }
      return;
    }
    c.segments.forEach((s, si) => {
      if (s.type === 't') {
        for (let k = 0; k < s.text.length; k++) { text += s.text[k]; map.push({ childIndex: ci, segIndex: si, offsetInSeg: k }); }
      } else if (!isInert(s.xml)) {
        text += '￼'; map.push({ childIndex: ci, segIndex: si, offsetInSeg: 0 });
      }
    });
  });
  return { text, map };
}

// Replace text ranges [start,end) of the paragraph with new top-level children.
// `edits` = [{ start, end, replacement: [xmlString, …] }], non-overlapping, any
// order; every start/end refers to joinedText offsets. Runs are split at the
// boundaries (their rPr duplicated), covered content is dropped, inert children
// are always kept. Returns a NEW paragraph object.
export function spliceParagraph(p, edits) {
  const sorted = [...edits].sort((a, b) => a.start - b.start);
  const out = { ...p, children: [] };
  let pos = 0;          // text offset of the next visible character
  let ei = 0;           // next edit to emit
  const covered = (off) => sorted.some((e) => e.start <= off && off < e.end);
  const emitDue = (at) => {   // emit every edit starting at or before absolute offset `at`
    while (ei < sorted.length && sorted[ei].start <= at) {
      out.children.push(...sorted[ei].replacement.map((xml) => ({ type: 'raw', xml })));
      ei += 1;
    }
  };
  for (const c of p.children) {
    if (c.type !== 'run') {
      if (isInert(c.xml)) { out.children.push(c); continue; }
      emitDue(pos);
      if (!covered(pos)) out.children.push(c);
      pos += 1;
      continue;
    }
    let segs = [];
    const flush = () => { if (segs.length) out.children.push({ ...c, segments: segs }); segs = []; };
    for (const s of c.segments) {
      if (s.type !== 't') {
        if (isInert(s.xml)) { segs.push(s); continue; }
        if (ei < sorted.length && sorted[ei].start <= pos) { flush(); emitDue(pos); }
        if (!covered(pos)) segs.push(s);
        pos += 1;
        continue;
      }
      const len = s.text.length;
      let k = 0;
      while (k < len) {
        // Emit any edit that starts here, flushing the run so far first.
        if (ei < sorted.length && sorted[ei].start <= pos + k) { flush(); emitDue(pos + k); }
        if (covered(pos + k)) { k += 1; continue; }
        // Keep the visible stretch up to the next edit start (or the end).
        const nextStart = ei < sorted.length ? sorted[ei].start - pos : len;
        const stop = Math.min(len, Math.max(k + 1, nextStart));
        segs.push({ type: 't', text: s.text.slice(k, stop) });
        k = stop;
      }
      pos += len;
    }
    flush();
  }
  emitDue(Infinity);   // an edit starting exactly at the end of the text
  // Merge adjacent text segments produced by the loop (cosmetic).
  for (const c of out.children) {
    if (c.type !== 'run') continue;
    const merged = [];
    for (const s of c.segments) {
      const last = merged[merged.length - 1];
      if (s.type === 't' && last && last.type === 't') last.text += s.text; else merged.push({ ...s });
    }
    c.segments = merged;
  }
  return out;
}

export function decodeXml(s) {
  return s.replace(/&(lt|gt|amp|quot|apos|#x[0-9a-fA-F]+|#\d+);/g, (m, e) => {
    switch (e) {
      case 'lt': return '<'; case 'gt': return '>'; case 'amp': return '&';
      case 'quot': return '"'; case 'apos': return "'";
      default: return String.fromCodePoint(e[1] === 'x' ? parseInt(e.slice(2), 16) : parseInt(e.slice(1), 10));
    }
  });
}

export function encodeXml(s) {
  return String(s).replace(/[&<>]/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;' })[c]);
}

export function encodeAttr(s) {
  return String(s).replace(/[&<>"\n\t]/g, (c) => ({
    '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', '\n': '&#10;', '\t': '&#9;',
  })[c]);
}

// The half-point font size of a run (w:sz in its rPr), or null.
export function runSizePt(run) {
  const m = /<w:sz\s+w:val="(\d+)"/.exec(run.rPr || '');
  return m ? parseInt(m[1], 10) / 2 : null;
}

// Does the paragraph hold anything but whitespace once `ranges` are removed?
export function textOutside(text, ranges) {
  let s = '';
  let cursor = 0;
  for (const r of [...ranges].sort((a, b) => a.start - b.start)) { s += text.slice(cursor, r.start); cursor = r.end; }
  s += text.slice(cursor);
  return s;
}
