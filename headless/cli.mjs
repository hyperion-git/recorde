// mjx-docx — headless MathJax equations for .docx files (agent-facing CLI).
// Subcommands: process, render, list, update, renumber, check, preview.
import { parseArgs } from 'node:util';
import { readFileSync, writeFileSync, mkdirSync, existsSync } from 'node:fs';
import { spawnSync } from 'node:child_process';
import { basename, dirname, join, resolve } from 'node:path';
import { DocxPackage } from './lib/zipdoc.mjs';
import { renderEquation, closeRenderers } from './lib/render.mjs';
import { svgToPng } from './lib/raster.mjs';
import { processDocument, listEquations, updateEquation, renumberDocument, checkDocument, WRITER } from './lib/document.mjs';

const USAGE = `mjx-docx — MathJax equations for Word documents, without Word (${WRITER})

  mjx-docx process <in.docx> [-o out.docx] [--font termes|tex|newcm|stix2|pagella|asana]   (default termes ≈ newtxmath)
                   [--size <pt>] [--color #rrggbb] [--number-style table|inline|field] [--align left|center]
                   [--preamble <file.tex>] [--dollar] [--static-refs] [--force] [--json]
      Replace [[math: …]] [[display: …]] [[eq: …]] [[eq#label: …]] [[ref: label]]
      placeholders with equation pictures. [[ref:]] becomes a REF field showing
      "(n)" (--static-refs: plain text). Defaults may also come from mjx.json next
      to the document ({ "font", "size", "color", "numberStyle", "align", "preamble", "dollar" }).
      Display alignment: left (LaTeX fleqn, 25 pt math indent; default) or center.
  mjx-docx render '<latex>' [--display] [--font …] [--size <pt>] [--color …] [--svg|--png|--json] [-o file]
  mjx-docx list <file.docx> [--json]
  mjx-docx update <file.docx> --uuid <uuid> [--latex '…'] [--font …] [--size <pt>] [--color …] [-o out.docx]
  mjx-docx renumber <file.docx> [-o out.docx]
  mjx-docx check <file.docx> [--json]            exit 1 on errors
  mjx-docx preview <file.docx> [--out <dir>] [--dpi <n>]   LibreOffice → PDF → PNG pages
`;

function fail(msg, code = 2) { process.stderr.write(`mjx-docx: ${msg}\n`); process.exit(code); }

function opts(args, spec) {
  try {
    return parseArgs({ args, options: spec, allowPositionals: true, strict: true });
  } catch (e) { fail(e.message); }
}

function loadConfig(docPath) {
  const cfg = join(dirname(resolve(docPath)), 'mjx.json');
  if (!existsSync(cfg)) return {};
  try { return JSON.parse(readFileSync(cfg, 'utf8')); } catch (e) { fail(`${cfg}: ${e.message}`); }
}

async function cmdProcess(args) {
  const { values: v, positionals: [file] } = opts(args, {
    out: { type: 'string', short: 'o' }, font: { type: 'string' }, size: { type: 'string' }, color: { type: 'string' },
    'number-style': { type: 'string' }, align: { type: 'string' }, preamble: { type: 'string' }, dollar: { type: 'boolean' },
    'static-refs': { type: 'boolean' }, force: { type: 'boolean' }, json: { type: 'boolean' },
  });
  if (!file) fail('process: missing <in.docx>');
  const cfg = loadConfig(file);
  const preambleText = v.preamble ? readFileSync(v.preamble, 'utf8')
    : cfg.preamble ? readFileSync(join(dirname(resolve(file)), cfg.preamble), 'utf8') : null;
  const pkg = DocxPackage.load(file);
  const report = await processDocument(pkg, {
    font: v.font ?? cfg.font, sizePt: v.size ?? cfg.size, color: v.color ?? cfg.color,
    numberStyle: v['number-style'] ?? cfg.numberStyle, align: v.align ?? cfg.align, preamble: preambleText,
    dollar: v.dollar ?? cfg.dollar, staticRefs: v['static-refs'] ?? cfg.staticRefs, force: v.force,
  });
  const out = v.out || file;
  if (!report.errors.length) pkg.save(out);
  if (v.json) { console.log(JSON.stringify({ ...report, out: report.errors.length ? null : out }, null, 2)); }
  else {
    for (const w of report.warnings) console.error(`warning: ${w}`);
    for (const e of report.errors) console.error(`error: ${e}`);
    if (!report.errors.length) console.log(`${out}: ${report.equations} equation(s), ${report.refs} reference(s)`);
  }
  await closeRenderers();
  process.exit(report.errors.length ? 1 : 0);
}

async function cmdRender(args) {
  const { values: v, positionals: [latex] } = opts(args, {
    display: { type: 'boolean' }, font: { type: 'string' }, size: { type: 'string' }, color: { type: 'string' },
    svg: { type: 'boolean' }, png: { type: 'boolean' }, json: { type: 'boolean' }, out: { type: 'string', short: 'o' },
  });
  if (!latex) fail('render: missing <latex>');
  const r = await renderEquation({ latex, font: v.font || 'termes', sizePt: Number(v.size || 11),
    displayMode: v.display ? 'display' : 'inline', color: v.color || '#000000' });
  if (!r.svg) fail(`render failed: ${r.error}`, 1);
  if (r.error) console.error(`warning: TeX error: ${r.error}`);
  let outBytes, text;
  if (v.png) { const { png, pxW, pxH } = await svgToPng(r.svg, r.widthPt, r.heightPt); outBytes = png; r.pxW = pxW; r.pxH = pxH; }
  else if (v.json) text = JSON.stringify({ ...r, svg: undefined, svgLength: r.svg.length }, null, 2);
  else text = r.svg;
  if (v.out) writeFileSync(v.out, outBytes ?? text);
  else if (outBytes) process.stdout.write(outBytes);
  else console.log(text);
  await closeRenderers();
}

function cmdList(args) {
  const { values: v, positionals: [file] } = opts(args, { json: { type: 'boolean' } });
  if (!file) fail('list: missing <file.docx>');
  const eqs = listEquations(DocxPackage.load(file));
  if (v.json) { console.log(JSON.stringify(eqs, null, 2)); return; }
  if (!eqs.length) { console.log('no equations'); return; }
  eqs.forEach((e, i) => {
    const num = e.numbered ? ` (${e.number}) [${e.numberStyle}]` : '';
    const flags = [e.hasPart ? '' : 'no-part', e.hasSvg ? '' : 'raster-only', e.inTable ? 'table' : ''].filter(Boolean).join(',');
    console.log(`${String(i + 1).padStart(3)}  ${e.uuid}  ${e.displayMode.padEnd(7)} ${e.font.padEnd(7)} ${String(e.sizePt).padStart(4)}pt${num}  ${e.latex}${flags ? `  {${flags}}` : ''}`);
  });
}

async function cmdUpdate(args) {
  const { values: v, positionals: [file] } = opts(args, {
    uuid: { type: 'string' }, latex: { type: 'string' }, font: { type: 'string' }, size: { type: 'string' },
    color: { type: 'string' }, out: { type: 'string', short: 'o' },
  });
  if (!file || !v.uuid) fail('update: needs <file.docx> --uuid <uuid>');
  const changes = {};
  if (v.latex != null) changes.latex = v.latex;
  if (v.font) changes.font = v.font;
  if (v.size) changes.sizePt = Number(v.size);
  if (v.color) changes.color = v.color;
  const pkg = DocxPackage.load(file);
  const { payload, error } = await updateEquation(pkg, v.uuid, changes);
  pkg.save(v.out || file);
  if (error) console.error(`warning: TeX error: ${error}`);
  console.log(`${v.out || file}: updated ${payload.uuid} → ${payload.latex}`);
  await closeRenderers();
}

async function cmdRenumber(args) {
  const { values: v, positionals: [file] } = opts(args, { out: { type: 'string', short: 'o' } });
  if (!file) fail('renumber: missing <file.docx>');
  const pkg = DocxPackage.load(file);
  const r = await renumberDocument(pkg);
  pkg.save(v.out || file);
  console.log(`${v.out || file}: ${r.changed} number(s) changed, ${r.fields} SEQ field(s) and ${r.refs} reference(s) refreshed`);
  await closeRenderers();
}

function cmdCheck(args) {
  const { values: v, positionals: [file] } = opts(args, { json: { type: 'boolean' } });
  if (!file) fail('check: missing <file.docx>');
  const issues = checkDocument(DocxPackage.load(file));
  if (v.json) console.log(JSON.stringify(issues, null, 2));
  else if (!issues.length) console.log(`${file}: OK`);
  else for (const i of issues) console.log(`${i.level.padEnd(5)} ${i.msg}`);
  process.exit(issues.some((i) => i.level === 'error') ? 1 : 0);
}

function cmdPreview(args) {
  const { values: v, positionals: [file] } = opts(args, { out: { type: 'string' }, dpi: { type: 'string' } });
  if (!file) fail('preview: missing <file.docx>');
  const outDir = resolve(v.out || 'preview');
  mkdirSync(outDir, { recursive: true });
  const so = spawnSync('soffice', ['--headless', '--convert-to', 'pdf', '--outdir', outDir, resolve(file)], { encoding: 'utf8' });
  if (so.status !== 0) fail(`LibreOffice conversion failed: ${so.stderr || so.stdout}`, 1);
  const pdf = join(outDir, basename(file).replace(/\.docx$/i, '') + '.pdf');
  const prefix = join(outDir, basename(file).replace(/\.docx$/i, ''));
  const pt = spawnSync('pdftoppm', ['-r', v.dpi || '150', '-png', pdf, prefix], { encoding: 'utf8' });
  if (pt.status !== 0) fail(`pdftoppm failed: ${pt.stderr}`, 1);
  console.log(`${outDir}: PDF + PNG pages (${basename(prefix)}-N.png). LibreOffice ≠ Word for baseline/spacing; treat as a sanity check.`);
}

export async function main(argv) {
  const [cmd, ...rest] = argv;
  switch (cmd) {
    case 'process': return cmdProcess(rest);
    case 'render': return cmdRender(rest);
    case 'list': return cmdList(rest);
    case 'update': return cmdUpdate(rest);
    case 'renumber': return cmdRenumber(rest);
    case 'check': return cmdCheck(rest);
    case 'preview': return cmdPreview(rest);
    case '-h': case '--help': case 'help': case undefined: process.stdout.write(USAGE); return;
    case '--version': case '-V': console.log(WRITER); return;
    default: fail(`unknown command "${cmd}"\n\n${USAGE}`);
  }
}
