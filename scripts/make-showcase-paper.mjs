// Build the showcase paper headlessly from the SAME source as the add-in's
// "Test paper" (addin/src/testpaper.js): outline → python-docx (layout, placeholders)
// → mjx-docx process (equations) → check → examples/Recorde-Showcase.docx.
// Usage: node scripts/make-showcase-paper.mjs [--preview] [--out file.docx]
import { spawnSync } from 'node:child_process';
import { copyFileSync, existsSync, mkdirSync } from 'node:fs';
import { resolve } from 'node:path';
import { paperOutline } from '../addin/src/testpaper.js';
import { DocxPackage } from '../headless/lib/zipdoc.mjs';
import { processDocument, checkDocument, listEquations } from '../headless/lib/document.mjs';
import { closeRenderers } from '../headless/lib/render.mjs';

const args = process.argv.slice(2);
const out = resolve(args.includes('--out') ? args[args.indexOf('--out') + 1] : 'examples/Recorde-Showcase.docx');
const preview = args.includes('--preview');
// Python with python-docx: MJX_PYTHON="<cmd> [args]" overrides; else a
// micromamba env named by MJX_MAMBA_ENV (default sci-base) if micromamba is on
// PATH; else plain python3.
const PY = (() => {
  if (process.env.MJX_PYTHON) { const [c, ...a] = process.env.MJX_PYTHON.split(/\s+/); return [c, a]; }
  const env = process.env.MJX_MAMBA_ENV || 'sci-base';
  const mm = spawnSync('micromamba', ['--version'], { encoding: 'utf8' });
  return mm.status === 0 ? ['micromamba', ['run', '-n', env, 'python']] : ['python3', []];
})();

mkdirSync('build', { recursive: true });
const raw = resolve('build/showcase-placeholders.docx');
const py = spawnSync(PY[0], [...PY[1], 'scripts/build-showcase-docx.py', raw],
  { input: JSON.stringify(paperOutline()), encoding: 'utf8' });
if (py.status !== 0) { console.error(py.stderr); process.exit(1); }

const pkg = DocxPackage.load(raw);
const report = await processDocument(pkg, { font: 'termes', numberStyle: 'table', color: '#000000' });
for (const w of report.warnings) console.error('warning:', w);
if (report.errors.length) { for (const e of report.errors) console.error('error:', e); process.exit(1); }
pkg.save(out);
const issues = checkDocument(DocxPackage.load(out));
for (const i of issues) console.log(`${i.level}: ${i.msg}`);
const eqs = listEquations(DocxPackage.load(out));
console.log(`${out}: ${eqs.length} equations (${eqs.filter((e) => e.numbered).length} numbered, `
  + `${new Set(eqs.map((e) => e.font)).size} fonts, styles: ${[...new Set(eqs.filter((e) => e.numbered).map((e) => e.numberStyle))].join('/')})`);
const win = '/mnt/c/Users/Public/MathJaxAddin/Recorde-Showcase.docx';
if (existsSync('/mnt/c/Users/Public/MathJaxAddin')) { copyFileSync(out, win); console.log('copied to', win); }
if (preview) {
  const r = spawnSync('node', ['headless/bin/mjx-docx.mjs', 'preview', out, '--out', 'build/showcase-preview'], { encoding: 'utf8' });
  console.log(r.stdout.trim() || r.stderr.trim());
}
await closeRenderers();
if (issues.some((i) => i.level === 'error')) process.exit(1);
