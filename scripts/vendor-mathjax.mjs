// Vendor MathJax 4.1.2 + its SVG font packages from node_modules into
// addin/assets/vendor/mathjax/ (gitignored) so the add-in serves them same-origin — no
// CDN at runtime (offline + reproducible fonts; WP3.3). Run by `postinstall` and
// `npm run vendor`. Defensive: a missing package warns and is skipped (so a
// partial/offline install never fails `npm ci`); build.js fails loudly if the
// core bundle didn't make it into dist.
import { cp, mkdir, rm, access } from 'node:fs/promises';
import { fileURLToPath } from 'node:url';
import { dirname, join } from 'node:path';

const ROOT = join(dirname(fileURLToPath(import.meta.url)), '..');
const NM = join(ROOT, 'node_modules');
const OUT = join(ROOT, 'addin', 'assets', 'vendor', 'mathjax');
// Fonts offered in the picker; each needs its SVG data pre-vendored because the
// runtime font-switch fetches it from this same origin.
const FONTS = ['tex', 'newcm', 'termes', 'stix2', 'pagella', 'asana'];

const exists = async (p) => { try { await access(p); return true; } catch { return false; } };

await rm(OUT, { recursive: true, force: true });
await mkdir(join(OUT, 'fonts'), { recursive: true });

// Core bundle: tex-svg.js already includes loader+startup+tex (ams/newcommand)+SVG
// output + the embedded base font, so it's the only core file to serve.
const core = join(NM, 'mathjax', 'tex-svg.js');
if (await exists(core)) {
  await cp(core, join(OUT, 'tex-svg.js'));
  console.log('vendored mathjax/tex-svg.js');
} else {
  console.warn('WARN: node_modules/mathjax/tex-svg.js missing — run `npm install`.');
}

// SRE speech worker + locale data (mathmaps): MathJax 4 spawns a speech worker at
// startup that importScripts these from <root>/sre/. Vendor them or the worker
// 404s and rejects startup.promise (rendering never becomes ready).
const sreDir = join(NM, 'mathjax', 'sre');
if (await exists(sreDir)) {
  await cp(sreDir, join(OUT, 'sre'), { recursive: true });
  console.log('vendored sre/ (speech worker + mathmaps)');
} else {
  console.warn('WARN: node_modules/mathjax/sre missing — speech worker would 404.');
}

// TeX extension packages (mathtools, physics, braket, bbm, boldsymbol, color, …):
// MathJax loads [tex]/<pkg> on demand from <root>/input/tex/extensions/. Vendor the
// input/ tree so enabling any package (or autoload/\require) doesn't 404.
const inputDir = join(NM, 'mathjax', 'input');
if (await exists(inputDir)) {
  await cp(inputDir, join(OUT, 'input'), { recursive: true });
  console.log('vendored input/ (TeX extension packages)');
} else {
  console.warn('WARN: node_modules/mathjax/input missing — extensions would 404.');
}

// SVG halves of each font package (drop chtml/, cjs/, mjs/, def/, examples/…).
for (const f of FONTS) {
  const pkg = join(NM, '@mathjax', `mathjax-${f}-font`);
  const svgJs = join(pkg, 'svg.js');
  if (!(await exists(svgJs))) { console.warn(`WARN: @mathjax/mathjax-${f}-font missing — skipped.`); continue; }
  const dst = join(OUT, 'fonts', `mathjax-${f}-font`);
  await mkdir(dst, { recursive: true });
  await cp(svgJs, join(dst, 'svg.js'));
  const svgDir = join(pkg, 'svg');
  if (await exists(svgDir)) await cp(svgDir, join(dst, 'svg'), { recursive: true });
  console.log(`vendored font mathjax-${f}`);
}
console.log('MathJax vendored → addin/assets/vendor/mathjax/');
