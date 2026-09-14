// Install the mathjax-docx skill for Claude Code.
//   node scripts/install-skill.mjs               → ~/.claude/skills/mathjax-docx  (you, every project)
//   node scripts/install-skill.mjs --project DIR → DIR/.claude/skills/mathjax-docx (that project only)
//   --copy   copy the folder instead of symlinking (symlink = follows repo updates)
//   --remove undo
// The skill's commands call `mjx-docx`, so also run `npm link` once in this repo.
import { existsSync, lstatSync, mkdirSync, rmSync, symlinkSync, cpSync, realpathSync } from 'node:fs';
import { join, resolve } from 'node:path';
import { homedir } from 'node:os';

const args = process.argv.slice(2);
const opt = (k) => args.includes(k);
const project = args.includes('--project') ? resolve(args[args.indexOf('--project') + 1]) : null;
const src = resolve(new URL('../skills/mathjax-docx', import.meta.url).pathname);
const base = project ? join(project, '.claude', 'skills') : join(homedir(), '.claude', 'skills');
const dst = join(base, 'mathjax-docx');

if (opt('--remove')) {
  if (existsSync(dst) || isLink(dst)) { rmSync(dst, { recursive: true, force: true }); console.log(`removed ${dst}`); }
  else console.log(`nothing at ${dst}`);
  process.exit(0);
}
mkdirSync(base, { recursive: true });
if (isLink(dst) || existsSync(dst)) {
  if (isLink(dst) && realpathSync(dst) === realpathSync(src)) { console.log(`already installed: ${dst} → ${src}`); process.exit(0); }
  console.error(`${dst} exists; remove it first (or run with --remove)`); process.exit(1);
}
let how = 'symlink';
if (opt('--copy')) { cpSync(src, dst, { recursive: true }); how = 'copy'; }
else {
  try { symlinkSync(src, dst, 'dir'); }
  catch (e) { cpSync(src, dst, { recursive: true }); how = `copy (symlink failed: ${e.code})`; }
}
console.log(`installed ${dst} (${how})`);
console.log('Next: `npm link` in this repo puts `mjx-docx` on your PATH (the skill calls it); or set');
console.log(`RECORDE_DIR=${resolve(src, '../..')} and use node $RECORDE_DIR/headless/bin/mjx-docx.mjs.`);

function isLink(p) { try { return lstatSync(p).isSymbolicLink(); } catch { return false; } }
