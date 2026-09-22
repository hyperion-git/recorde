# Installing mjx-docx (equations in .docx without Word)

`mjx-docx` is a command-line tool that replaces `[[math: …]]` placeholders in a
`.docx` with MathJax equation pictures — the same pictures the Recorde Word
add-in inserts, so the equations stay click-to-edit in Word. It is a Node.js
package; Word, Python and a network connection are not needed at run time.

This file ships in `recorde-mjx-docx-<version>.zip` together with the npm
package `recorde-<version>.tgz`, the agent skill (`skill/recorde/`) and
`Recorde-Showcase.docx`.

## What you need

- **Node.js 20 or newer** with npm (<https://nodejs.org>). Linux, macOS,
  Windows.
- Disk: about 200 MB after install (MathJax and six font packages are fetched
  from the npm registry during installation).
- Optional: **LibreOffice** and **poppler-utils** (`pdftoppm`) for
  `mjx-docx preview` (page images of the finished document). Optional:
  Python with `python-docx`, pandoc or docx-js — whatever you already use to
  build documents; `mjx-docx` only post-processes the file.

## Install from the release zip

```sh
unzip recorde-mjx-docx-<version>.zip
cd recorde-mjx-docx-<version>
npm install -g ./recorde-<version>.tgz
mjx-docx --version              # → mjx-docx/<version>
```

`npm install -g` puts `mjx-docx` on your PATH (on Linux/macOS with a
system-wide Node, prefix `sudo` or use a user-level npm prefix). The PNG
fallback inside each picture uses `@resvg/resvg-js`, an optional native
dependency; if it is unavailable for your platform, `cairosvg` or ImageMagick
(`convert`) on the PATH are used instead.

**Update:** install the newer tarball the same way. **Uninstall:**
`npm uninstall -g recorde`.

## Install from the repository (development)

```sh
git clone https://github.com/hyperion-git/recorde.git && cd recorde
npm install             # MathJax + fonts + zip/raster libraries, and the add-in toolchain
npm link                # `mjx-docx` on the PATH, following the checkout
```

Without `npm link`, run `node headless/bin/mjx-docx.mjs …` from the checkout.

## Agent skill (Claude Code and others)

The package includes a skill that teaches a coding agent the placeholder
grammar and the process/check/preview procedure.

```sh
mjx-docx skill install                       # → ~/.claude/skills/recorde (all your projects)
mjx-docx skill install --project ~/my-paper  # → that project's .claude/skills/ only
mjx-docx skill remove                        # undo; `--copy` copies instead of symlinking
```

Claude Code lists it as `recorde` and applies it whenever a Word document
needs formulas; `/recorde` invokes it by hand. For other agents, paste
`skill/recorde/AGENTS-snippet.md` into the project's `AGENTS.md`. The skill
runs `mjx-docx`, so keep the command on the PATH.

## First run

```sh
mjx-docx render 'E = mc^2' --svg -o e.svg          # renders without a document
mjx-docx list Recorde-Showcase.docx                # 76 equations from the zip's demo paper
mjx-docx check Recorde-Showcase.docx               # structural lint (exit 1 = problems)
```

Build a document with placeholders — `[[math: \hbar\omega]]` inline,
`[[display: …]]`, `[[eq#label: …]]` numbered, `[[ref: label]]` — then

```sh
mjx-docx process paper.docx [--font termes] [--number-style table] [--align left]
mjx-docx check paper.docx
mjx-docx preview paper.docx                        # LibreOffice → preview/paper-1.png …
```

`mjx-docx --help` lists every command and option; the placeholder grammar,
builder recipes (python-docx, pandoc, docx-js) and the round-trip commands
(`list`, `update`, `renumber`) are in `skill/recorde/SKILL.md`.

Documents produced this way open in Word without a repair prompt; with the
Recorde add-in installed, clicking an equation loads its LaTeX into the pane.

## Troubleshooting

- `mjx-docx: command not found` after `npm install -g`: the global bin
  directory is not on the PATH — `npm prefix -g` shows it (`<prefix>/bin`, on
  Windows the prefix itself).
- `Cannot find package 'mathjax'`: the install was interrupted before the
  dependencies arrived; re-run `npm install -g ./recorde-<version>.tgz`.
- `preview` fails: it needs `soffice` (LibreOffice) and `pdftoppm` on the PATH.
- Issues: <https://github.com/hyperion-git/recorde/issues>.
