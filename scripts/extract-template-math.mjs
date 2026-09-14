// Extract every math snippet from the author's paper template main document
// into test/fixtures/template-math.json, the acceptance fixture for the
// built-in preamble (test/preamble.test.js): each snippet must render with
// DEFAULT_PREAMBLE without a TeX error or an undefined (red) macro.
// Usage: node scripts/extract-template-math.mjs [~/dev/paper-template/10-main.tex]
import { readFileSync, writeFileSync } from 'node:fs';
const src = process.argv[2] || `${process.env.HOME}/dev/paper-template/10-main.tex`;
const main = readFileSync(src, 'utf8').replace(/(^|[^\\])%.*$/gm, '$1');
const out = [];
const re = /\$\$([\s\S]+?)\$\$|\\\[([\s\S]+?)\\\]|\\begin\{(equation\*?|align\*?|gather\*?|multline\*?)\}([\s\S]+?)\\end\{\3\}|(?<!\$)\$([^$\n]+?)\$(?!\$)|\\\(([\s\S]+?)\\\)/g;
for (const m of main.matchAll(re)) {
  const tex = (m[1] ?? m[2] ?? m[4] ?? m[5] ?? m[6]).replace(/\\label\{[^}]*\}|\\nonumber/g, '').replace(/\\\\\s*$/, '').trim();
  if (tex) out.push({ tex, display: !(m[5] ?? m[6]) });
}
// Only the math is recorded — no document text, no author/affiliation macros.
writeFileSync('test/fixtures/template-math.json', JSON.stringify({ source: '10-main.tex', snippets: out }, null, 2) + '\n');
console.log(`${out.length} snippets`);
