// Build the deployable add-in site into dist/ for one HTTPS origin.
//
// Repository layout vs served layout: the pane lives in addin/src/, the pure
// modules it shares with the headless CLI in core/. Both are served from ONE
// directory, <base>/src/ (pane files import them as ./x.js), so the published
// URLs — <base>/src/taskpane.html, <base>/assets/… — never change with the
// repository layout. addin/dev-server.cjs applies the same merge for `npm start`.
import { readFile, writeFile, rm, mkdir, cp, access, readdir } from 'node:fs/promises';
import { fileURLToPath } from 'node:url';
import { dirname, join } from 'node:path';

const ADDIN = dirname(fileURLToPath(import.meta.url));
const ROOT = join(ADDIN, '..');
const CORE = join(ROOT, 'core');
const DIST = join(ROOT, 'dist');
const DEV_ORIGIN = 'https://localhost:3000';

const DEV_ID = '2a7c1e3a-9f4b-4c1d-8e1f-3c2a5b9d7e10';
const PROD_ID = '9971ece4-30c4-421f-920e-8df1563b6bd5';

const BASE_URL = (process.env.BASE_URL || process.argv[2] || '').replace(/\/+$/, '');
if (!BASE_URL) {
  console.error('Usage: BASE_URL=https://host/path node addin/build.js   (or pass it as $1)');
  process.exit(1);
}
if (!/^https:\/\//.test(BASE_URL)) {
  console.error(`BASE_URL must be https:// (Office requires HTTPS). Got: ${BASE_URL}`);
  process.exit(1);
}

await rm(DIST, { recursive: true, force: true });
await mkdir(DIST, { recursive: true });
await cp(join(ADDIN, 'src'), join(DIST, 'src'), { recursive: true });
// Merge core/ into the served src/ — a name clash would silently shadow a pane file.
const paneFiles = new Set(await readdir(join(ADDIN, 'src')));
for (const f of await readdir(CORE)) {
  if (!f.endsWith('.js')) continue;
  if (paneFiles.has(f)) { console.error(`core/${f} clashes with addin/src/${f}; rename one.`); process.exit(1); }
  await cp(join(CORE, f), join(DIST, 'src', f));
}
await cp(join(ADDIN, 'assets'), join(DIST, 'assets'), { recursive: true });
await cp(join(ADDIN, 'index.html'), join(DIST, 'index.html'));   // support page (manifest SupportUrl = site root)
for (const f of ['LICENSE', 'NOTICE', 'THIRD-PARTY-NOTICES.md']) {
  await cp(join(ROOT, f), join(DIST, f));
}

// WP3.3: the self-hosted MathJax must ship or the deployed add-in can't render.
// It's vendored from node_modules by `prepare` (gitignored), so a checkout that
// skipped `npm install` would otherwise produce a broken dist — fail loudly.
try {
  await access(join(DIST, 'assets', 'vendor', 'mathjax', 'tex-svg.js'));
} catch {
  console.error('Missing vendored MathJax (addin/assets/vendor/mathjax/tex-svg.js). '
    + 'Run `npm install` (or `npm run vendor`) before building.');
  process.exit(1);
}

// Cache-busting (learned the hard way): Word/WebView2 and the Pages CDN cache the
// pane by URL for max-age (10 min on Pages), so a fresh deploy could serve a
// stale HTML with new JS (or vice versa) — the v1.12.0 pane "didn't show". Give
// every release a distinct URL for its WHOLE module graph: the manifest's pane
// URL, the module <script>, and each relative import in dist/src/*.js carry
// ?v=<package version>. The source trees are left untouched.
const { version } = JSON.parse(await readFile(join(ROOT, 'package.json'), 'utf8'));
const V = encodeURIComponent(version);
const srcDir = join(DIST, 'src');
let rewrittenImports = 0;
for (const f of await readdir(srcDir)) {
  const fp = join(srcDir, f);
  if (f.endsWith('.js')) {
    let js = await readFile(fp, 'utf8');
    js = js.replace(/(from\s+|import\s+)(['"])(\.\/[\w-]+\.js)\2/g, (m, kw, q, spec) => {
      rewrittenImports++;
      return `${kw}${q}${spec}?v=${V}${q}`;
    });
    await writeFile(fp, js);
  } else if (f === 'taskpane.html') {
    let html = await readFile(fp, 'utf8');
    html = html.replace('src="./taskpane.js"', `src="./taskpane.js?v=${V}"`);
    await writeFile(fp, html);
  }
}

let manifest = await readFile(join(ADDIN, 'manifest.xml'), 'utf8');
manifest = manifest.split('src/taskpane.html"').join(`src/taskpane.html?v=${V}"`);

// 1) Dev origin → deploy origin.
const occurrences = manifest.split(DEV_ORIGIN).length - 1;
manifest = manifest.split(DEV_ORIGIN).join(BASE_URL);

// 2) Dev <Id> → prod <Id> (distinct identity for the deployed add-in). Fail
// loudly if the expected dev Id isn't present — a silent miss here is exactly
// the bug that produced the load blocker.
if (!manifest.includes(DEV_ID)) {
  console.error(`Expected dev <Id> ${DEV_ID} in addin/manifest.xml; not found. Aborting.`);
  process.exit(1);
}
manifest = manifest.split(DEV_ID).join(PROD_ID);

await writeFile(join(DIST, 'manifest.xml'), manifest);

console.log(
  `Built dist/ for ${BASE_URL} — rewrote ${occurrences} dev-origin URL(s) ` +
  `and dev Id → prod Id ${PROD_ID} in manifest.xml; versioned ${rewrittenImports} import(s) + pane URLs with ?v=${V}.`,
);
