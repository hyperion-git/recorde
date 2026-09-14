// Offline structural manifest check (WP0.4). Dependency-free so it runs in CI
// without npm install and without the flaky remote validator. It does NOT
// replace full XSD validation (`npm run validate`, or the documented offline
// lxml/XSD approach) — it catches the regressions a developer actually causes:
// a renamed resource, a fat-fingered resid, a missing required element, a bad
// version, or localhost URLs left in a built manifest.
import { readFile } from 'node:fs/promises';

const path = process.argv[2] || 'manifest.xml';
const xml = await readFile(path, 'utf8');
const errors = [];

// Required top-level elements.
for (const tag of ['Id', 'Version', 'DefaultSettings', 'SourceLocation', 'Permissions']) {
  if (!new RegExp(`<${tag}[ >/]`).test(xml)) errors.push(`missing <${tag}>`);
}

// Version: 1–4 dotted integers.
const ver = xml.match(/<Version>([^<]*)<\/Version>/)?.[1];
if (ver && !/^\d{1,5}(\.\d{1,5}){0,3}$/.test(ver)) errors.push(`bad <Version> "${ver}"`);

// resid resolution: every resid="X" must have a matching id="X" somewhere
// (resource ids are a subset of all ids, so this never false-negatives).
const resids = [...new Set([...xml.matchAll(/\bresid="([^"]+)"/g)].map((m) => m[1]))];
const ids = new Set([...xml.matchAll(/\bid="([^"]+)"/g)].map((m) => m[1]));
const missing = resids.filter((r) => !ids.has(r));
if (missing.length) errors.push(`unresolved resid(s): ${missing.join(', ')}`);

// A built/dist manifest must not still point at the dev origin.
if (/dist[\\/]/.test(path) && xml.includes('localhost')) {
  errors.push('built manifest still contains localhost URLs (build rewrite failed)');
}

// A built/dist manifest must carry the prod <Id>, not the dev one — a shared Id
// collides with the localhost dev registration in Office and fails to load.
if (/dist[\\/]/.test(path) && xml.includes('2a7c1e3a-9f4b-4c1d-8e1f-3c2a5b9d7e10')) {
  errors.push('built manifest still uses the DEV <Id> (prod Id rewrite failed)');
}

if (errors.length) {
  console.error(`✗ ${path} failed structural checks:`);
  for (const e of errors) console.error('  - ' + e);
  process.exit(1);
}
console.log(`✓ ${path}: ${resids.length} resid(s) resolve, required elements present, version "${ver}" OK.`);
