// Minimal OPC package model over a .docx zip. Every part we do not touch is
// carried byte-for-byte (Word is fussy about re-serialized XML), and only the
// handful of parts the equation pass needs are edited as text: the main
// document, its relationships, [Content_Types].xml, and our customXml items.

import { readFileSync, writeFileSync } from 'node:fs';
import { randomUUID } from 'node:crypto';
import { unzipSync, zipSync, strFromU8, strToU8 } from 'fflate';

const CT_PATH = '[Content_Types].xml';
const REL_NS = 'http://schemas.openxmlformats.org/package/2006/relationships';
const REL = {
  officeDocument: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument',
  image: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image',
  customXml: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/customXml',
  customXmlProps: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/customXmlProps',
};
export const CONTENT_TYPES = {
  svg: 'image/svg+xml',
  png: 'image/png',
  xml: 'application/xml',
  customXmlProps: 'application/vnd.openxmlformats-officedocument.customXmlProperties+xml',
};

export class DocxPackage {
  constructor(entries) {
    this.entries = entries;            // Map<name, Uint8Array>
    this.mainPart = this.resolveMainPart();
  }

  static load(path) {
    const raw = unzipSync(new Uint8Array(readFileSync(path)));
    const entries = new Map();
    for (const [name, bytes] of Object.entries(raw)) {
      if (name.endsWith('/')) continue;           // directory entries
      entries.set(name, bytes);
    }
    if (!entries.has(CT_PATH)) throw new Error(`${path}: not an OOXML package ([Content_Types].xml missing)`);
    return new DocxPackage(entries);
  }

  save(path) {
    // [Content_Types].xml first — Word tolerates other orders but this is the
    // canonical layout every producer uses.
    const ordered = {};
    ordered[CT_PATH] = this.entries.get(CT_PATH);
    for (const [name, bytes] of this.entries) if (name !== CT_PATH) ordered[name] = bytes;
    writeFileSync(path, zipSync(ordered, { level: 6 }));
  }

  has(name) { return this.entries.has(name); }
  names() { return [...this.entries.keys()]; }
  bytes(name) { return this.entries.get(name); }
  setBytes(name, bytes) { this.entries.set(name, bytes); }
  text(name) { const b = this.entries.get(name); return b == null ? null : strFromU8(b); }
  setText(name, str) { this.entries.set(name, strToU8(str)); }
  remove(name) { this.entries.delete(name); }

  // ---- main document ----
  resolveMainPart() {
    const rels = this.readRels('');                   // package-level _rels/.rels
    const main = rels.find((r) => r.type === REL.officeDocument);
    return main ? normalize('', main.target) : 'word/document.xml';
  }

  // ---- relationships ----
  static relsPathFor(partPath) {
    if (!partPath) return '_rels/.rels';
    const i = partPath.lastIndexOf('/');
    return (i < 0 ? '_rels/' : partPath.slice(0, i + 1) + '_rels/') + partPath.slice(i + 1) + '.rels';
  }

  readRels(partPath) {
    const xml = this.text(DocxPackage.relsPathFor(partPath));
    if (!xml) return [];
    const out = [];
    for (const m of xml.matchAll(/<Relationship\s([^>]*?)\/?>/g)) {
      const a = m[1];
      const get = (k) => { const r = new RegExp(`\\b${k}="([^"]*)"`).exec(a); return r ? r[1] : null; };
      out.push({ id: get('Id'), type: get('Type'), target: get('Target'), mode: get('TargetMode') });
    }
    return out;
  }

  // Add a relationship from `partPath` to `target` (relative to the part's
  // folder, as OPC expects). Returns the new rId.
  addRel(partPath, type, target) {
    const path = DocxPackage.relsPathFor(partPath);
    let xml = this.text(path);
    if (!xml) {
      xml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="${REL_NS}"></Relationships>`;
    }
    const used = new Set(this.readRels(partPath).map((r) => r.id));
    let n = 1;
    while (used.has(`rId${n}`)) n += 1;
    const id = `rId${n}`;
    const rel = `<Relationship Id="${id}" Type="${type}" Target="${target}"/>`;
    xml = xml.replace(/<\/Relationships>\s*$/, rel + '</Relationships>');
    this.setText(path, xml);
    return id;
  }

  removeRel(partPath, id) {
    const path = DocxPackage.relsPathFor(partPath);
    const xml = this.text(path);
    if (!xml) return;
    this.setText(path, xml.replace(new RegExp(`<Relationship\\s[^>]*\\bId="${id}"[^>]*?/>\\s*`), ''));
  }

  // Resolve a relationship target from a part to a package path.
  resolveTarget(partPath, target) { return normalize(partPath, target); }

  // ---- content types ----
  ensureDefault(ext, contentType) {
    let ct = this.text(CT_PATH);
    if (new RegExp(`<Default\\s[^>]*Extension="${ext}"`, 'i').test(ct)) return;
    ct = ct.replace(/<Types([^>]*)>/, `<Types$1><Default Extension="${ext}" ContentType="${contentType}"/>`);
    this.setText(CT_PATH, ct);
  }

  ensureOverride(partName, contentType) {
    let ct = this.text(CT_PATH);
    const abs = partName.startsWith('/') ? partName : '/' + partName;
    if (ct.includes(`PartName="${abs}"`)) return;
    ct = ct.replace(/<\/Types>\s*$/, `<Override PartName="${abs}" ContentType="${contentType}"/></Types>`);
    this.setText(CT_PATH, ct);
  }

  removeOverride(partName) {
    const abs = partName.startsWith('/') ? partName : '/' + partName;
    const ct = this.text(CT_PATH);
    this.setText(CT_PATH, ct.replace(new RegExp(`<Override\\s[^>]*PartName="${abs.replace(/[.*+?^${}()|[\]\\]/g, '\\$&')}"[^>]*/>`), ''));
  }

  // ---- media ----
  // Store `bytes` as word/media/imageN.<ext> and relate it from the main part.
  // Returns { path, rId }.
  addMedia(ext, bytes) {
    const dir = this.mainPart.slice(0, this.mainPart.lastIndexOf('/') + 1) + 'media/';
    let n = 1;
    for (const name of this.entries.keys()) {
      const m = new RegExp(`^${dir}image(\\d+)\\.`).exec(name);
      if (m) n = Math.max(n, parseInt(m[1], 10) + 1);
    }
    const path = `${dir}image${n}.${ext}`;
    this.setBytes(path, bytes);
    this.ensureDefault(ext, CONTENT_TYPES[ext] || 'application/octet-stream');
    const rId = this.addRel(this.mainPart, REL.image, `media/image${n}.${ext}`);
    return { path, rId };
  }

  // Path of the media file behind a main-part rId, or null.
  mediaPathFor(rId) {
    const rel = this.readRels(this.mainPart).find((r) => r.id === rId);
    return rel ? this.resolveTarget(this.mainPart, rel.target) : null;
  }

  // ---- custom XML parts ----
  // Word's layout: customXml/itemN.xml (the payload), customXml/itemPropsN.xml
  // (datastore id + schema namespace), customXml/_rels/itemN.xml.rels, and a
  // customXml relationship from the main part. Returns the item index.
  addCustomXml(xml, namespace) {
    let n = 1;
    for (const name of this.entries.keys()) {
      const m = /^customXml\/item(\d+)\.xml$/.exec(name);
      if (m) n = Math.max(n, parseInt(m[1], 10) + 1);
    }
    const item = `customXml/item${n}.xml`;
    const props = `customXml/itemProps${n}.xml`;
    this.setText(item, `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n${xml}`);
    this.setText(props, `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n`
      + `<ds:datastoreItem ds:itemID="{${randomUUID().toUpperCase()}}" xmlns:ds="http://schemas.openxmlformats.org/officeDocument/2006/customXml">`
      + `<ds:schemaRefs><ds:schemaRef ds:uri="${namespace}"/></ds:schemaRefs></ds:datastoreItem>`);
    this.setText(`customXml/_rels/item${n}.xml.rels`,
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="${REL_NS}">`
      + `<Relationship Id="rId1" Type="${REL.customXmlProps}" Target="itemProps${n}.xml"/></Relationships>`);
    this.ensureDefault('xml', CONTENT_TYPES.xml);
    this.ensureOverride(props, CONTENT_TYPES.customXmlProps);
    this.addRel(this.mainPart, REL.customXml, `../customXml/item${n}.xml`);
    return n;
  }

  // All custom XML items: [{ index, path, xml, namespace }] (namespace = the
  // root element's default namespace, which is what Office.js getByNamespace
  // keys on; itemProps is informational).
  listCustomXml() {
    const out = [];
    for (const name of [...this.entries.keys()].sort()) {
      const m = /^customXml\/item(\d+)\.xml$/.exec(name);
      if (!m) continue;
      const xml = this.text(name);
      const root = /<([\w.:-]+)([^>]*)>/.exec(xml.replace(/^<\?xml[^>]*\?>\s*/, ''));
      const ns = root && /\sxmlns="([^"]+)"/.exec(root[2]);
      out.push({ index: parseInt(m[1], 10), path: name, xml, namespace: ns ? ns[1] : null });
    }
    return out;
  }

  removeCustomXml(index) {
    const item = `customXml/item${index}.xml`;
    const props = `customXml/itemProps${index}.xml`;
    this.remove(item); this.remove(props); this.remove(`customXml/_rels/item${index}.xml.rels`);
    this.removeOverride(props);
    const rel = this.readRels(this.mainPart).find((r) => r.type === REL.customXml && this.resolveTarget(this.mainPart, r.target) === item);
    if (rel) this.removeRel(this.mainPart, rel.id);
  }
}

// Resolve `target` relative to the folder of `fromPart` ('' = package root).
function normalize(fromPart, target) {
  if (target.startsWith('/')) return target.slice(1);
  const base = fromPart.includes('/') ? fromPart.slice(0, fromPart.lastIndexOf('/')).split('/') : [];
  for (const seg of target.split('/')) {
    if (seg === '..') base.pop(); else if (seg !== '.' && seg !== '') base.push(seg);
  }
  return base.join('/');
}

export { REL };
