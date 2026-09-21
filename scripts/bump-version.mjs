// Single-source version bump (WP0.5). Keeps package.json and addin/manifest.xml in
// lockstep — Office wants a 4-part version, so manifest gets "<v>.0".
//
//   npm run bump -- 1.1.0
import { readFile, writeFile } from 'node:fs/promises';

const v = process.argv[2];
if (!v || !/^\d{1,5}(\.\d{1,5}){0,2}$/.test(v)) {
  console.error('Usage: npm run bump -- <major.minor.patch>   e.g. npm run bump -- 1.1.0');
  process.exit(1);
}

const pkg = JSON.parse(await readFile('package.json', 'utf8'));
pkg.version = v;
await writeFile('package.json', JSON.stringify(pkg, null, 2) + '\n');

let man = await readFile('addin/manifest.xml', 'utf8');
man = man.replace(/<Version>[^<]*<\/Version>/, `<Version>${v}.0</Version>`);
await writeFile('addin/manifest.xml', man);

console.log(`Bumped → package.json ${v}, addin/manifest.xml ${v}.0. Remember to redeploy so caches refresh.`);
