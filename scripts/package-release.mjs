// Build the downloadable release packages into release/ (gitignored):
//
//   recorde-word-addin-<v>.zip   the Word add-in: released manifest.xml, INSTALL.md,
//                                 sideload installers (Windows/Mac), showcase document, licences
//   recorde-mjx-docx-<v>.zip     the headless CLI: recorde-<v>.tgz (npm pack), INSTALL.md,
//                                 the agent skill, showcase document, licences
//   manifest.xml                 the released manifest on its own (Upload My Add-in)
//   SHA256SUMS
//
// Needs dist/ built for the public origin first (BASE_URL=… npm run build); the
// manifest that goes into the zip is checked here the same way validate:dist
// checks it — a dev manifest (localhost, dev <Id>) must never ship.
//
//   npm run package [-- --out <dir>]
import { readFileSync, writeFileSync, mkdirSync, rmSync, existsSync, readdirSync, statSync } from 'node:fs';
import { spawnSync } from 'node:child_process';
import { createHash } from 'node:crypto';
import { join, resolve, dirname } from 'node:path';
import { fileURLToPath } from 'node:url';
import { zipSync } from 'fflate';

const ROOT = resolve(dirname(fileURLToPath(import.meta.url)), '..');
const args = process.argv.slice(2);
const OUT = resolve(args.includes('--out') ? args[args.indexOf('--out') + 1] : join(ROOT, 'release'));
const DEV_ID = '2a7c1e3a-9f4b-4c1d-8e1f-3c2a5b9d7e10';

const fail = (msg) => { console.error(`package: ${msg}`); process.exit(1); };
const rel = (p) => join(ROOT, p);
const read = (p) => readFileSync(rel(p));

const { version } = JSON.parse(read('package.json'));

// --- the released manifest -------------------------------------------------
const manifestPath = rel('dist/manifest.xml');
if (!existsSync(manifestPath)) fail('dist/manifest.xml missing — run `BASE_URL=https://… npm run build` first');
const manifest = readFileSync(manifestPath, 'utf8');
if (manifest.includes('localhost')) fail('dist/manifest.xml still points at localhost');
if (manifest.includes(DEV_ID)) fail('dist/manifest.xml still carries the dev <Id>');
const id = manifest.match(/<Id>([^<]+)<\/Id>/)?.[1];
if (!id) fail('dist/manifest.xml has no <Id>');
const manVersion = manifest.match(/<Version>([^<]+)<\/Version>/)?.[1];
if (manVersion !== `${version}.0`) fail(`dist/manifest.xml <Version> ${manVersion} ≠ package.json ${version}.0 — rebuild`);

// The Windows installers register the add-in under its <Id>; they must agree.
for (const f of ['install-windows.cmd', 'uninstall-windows.cmd']) {
  if (!readFileSync(rel(`addin/install/${f}`), 'utf8').includes(`set "ID=${id}"`)) fail(`addin/install/${f} does not carry the manifest <Id> ${id}`);
}

// --- the npm tarball (headless CLI) ------------------------------------------
rmSync(OUT, { recursive: true, force: true });
mkdirSync(OUT, { recursive: true });
const packDir = join(OUT, '.pack');
mkdirSync(packDir);
// --ignore-scripts: `prepare` (MathJax vendoring for the add-in) is irrelevant to the tarball and would pollute --json.
const pack = spawnSync('npm', ['pack', '--pack-destination', packDir, '--json', '--ignore-scripts'], { cwd: ROOT, encoding: 'utf8' });
if (pack.status !== 0) fail(`npm pack failed:\n${pack.stderr}`);
const packed = JSON.parse(pack.stdout);
const tgzName = (Array.isArray(packed) ? packed[0] : Object.values(packed)[0]).filename;   // npm ≤11 array, npm 12 object
const tgz = readFileSync(join(packDir, tgzName));
rmSync(packDir, { recursive: true, force: true });

// --- zips --------------------------------------------------------------------
const mtime = new Date();
const file = (buf, exec = false) => [buf instanceof Uint8Array ? buf : new Uint8Array(buf),
  exec ? { mtime, os: 3, attrs: 0o100755 << 16 } : { mtime, os: 3, attrs: 0o100644 << 16 }];
const licences = { 'LICENSE': file(read('LICENSE')), 'NOTICE': file(read('NOTICE')), 'THIRD-PARTY-NOTICES.md': file(read('THIRD-PARTY-NOTICES.md')) };

const wordDir = `recorde-word-addin-${version}`;
const wordZip = zipSync({ [wordDir]: {
  'manifest.xml': file(Buffer.from(manifest)),
  'INSTALL.md': file(read('docs/INSTALL-word.md')),
  'install-windows.cmd': file(read('addin/install/install-windows.cmd')),
  'uninstall-windows.cmd': file(read('addin/install/uninstall-windows.cmd')),
  'install-mac.command': file(read('addin/install/install-mac.command'), true),
  'uninstall-mac.command': file(read('addin/install/uninstall-mac.command'), true),
  'Recorde-Showcase.docx': file(read('examples/Recorde-Showcase.docx')),
  'PRIVACY.md': file(read('PRIVACY.md')),
  ...licences,
} }, { level: 6 });

const cliDir = `recorde-mjx-docx-${version}`;
const skillFiles = Object.fromEntries(readdirSync(rel('skills/recorde')).map((f) => [f, file(read(`skills/recorde/${f}`))]));
const cliZip = zipSync({ [cliDir]: {
  [tgzName]: [new Uint8Array(tgz), { mtime, os: 3, attrs: 0o100644 << 16, level: 0 }],   // already gzip-compressed
  'INSTALL.md': file(read('docs/INSTALL-mjx-docx.md')),
  'skill': { 'recorde': skillFiles },
  'Recorde-Showcase.docx': file(read('examples/Recorde-Showcase.docx')),
  ...licences,
} }, { level: 6 });

const outputs = {
  [`${wordDir}.zip`]: wordZip,
  [`${cliDir}.zip`]: cliZip,
  'manifest.xml': Buffer.from(manifest),
};
const sums = [];
for (const [name, data] of Object.entries(outputs)) {
  writeFileSync(join(OUT, name), data);
  sums.push(`${createHash('sha256').update(data).digest('hex')}  ${name}`);
}
writeFileSync(join(OUT, 'SHA256SUMS'), sums.join('\n') + '\n');

console.log(`Release ${version} → ${OUT}`);
for (const name of [...Object.keys(outputs), 'SHA256SUMS']) {
  console.log(`  ${name.padEnd(40)} ${(statSync(join(OUT, name)).size / 1024).toFixed(0).padStart(6)} kB`);
}
console.log(`  (tarball inside ${cliDir}.zip: ${tgzName}, ${(tgz.length / 1024).toFixed(0)} kB)`);
