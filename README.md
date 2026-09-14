# Recorde — LaTeX equations in Word, with editable source

Recorde is a Word add-in that inserts LaTeX equations as vector pictures
(MathJax 4 SVG) and keeps the LaTeX inside the document, so every equation stays
click-to-edit. A companion command-line tool, `mjx-docx`, does the same for
`.docx` files without Word, so scripts and coding agents can produce documents
whose equations open in the add-in.

- **Six math fonts** — TeX Gyre Termes (default, the Times face closest to
  `newtxmath`), Computer Modern, New Computer Modern, STIX Two, Pagella, Asana.
- **Numbering** — flush-right journal style (borderless table), inline `(n)`,
  or native Word SEQ fields; renumber in document order; switch styles and
  convert existing equations.
- **TeX layout rules** — display equations get `\abovedisplayskip` /
  `\belowdisplayskip` (one body size) and continue the paragraph without an
  indent; left-aligned (`fleqn`, 25 pt math indent) by default or centred.
- **Macro preamble** — the author's paper-template macros (operators,
  delimiters, derivatives, vectors, a siunitx subset, 174 named colours) are
  built in; add per-document `\newcommand`s in the pane.
- **Colour, sizes, examples** — swatches from a scientific-figure palette,
  size matched to the surrounding text or fixed, an examples menu, keyboard
  shortcuts, an equation list, a one-click test page and a three-page test
  paper that exercises everything.
- **Self-hosted, offline-capable** — MathJax and all font data are served with
  the add-in; nothing is fetched from a CDN and nothing leaves the document
  (see `PRIVACY.md`).
- **English and German** UI.

## Install (end users)

**What you need**

- Word for Microsoft 365 on Windows or Mac (current channel), or Word on the
  web. The pane detects what the host can do at start-up: vector SVG insert
  needs ImageCoercion 1.2 (Microsoft 365 desktop builds since 2019), the web
  falls back to a high-resolution raster picture, and a host with neither
  shows the preview but disables Insert. Baseline-aligned inline math needs
  Word 2507 or later. Older perpetual versions (2016/2019/2021) are untested.
- Network access to the hosting site on first use (Word caches the pane
  afterwards). No admin rights for *Upload My Add-in*; no installer; nothing
  is written outside the document and Word's add-in cache.

**Steps.** The add-in is web-hosted; the only file to install is
`manifest.xml` from the release you were given (it points at the hosted build).

- **Word for Windows / Mac:** Home → Add-ins → *More Add-ins* → **My Add-ins** →
  *Upload My Add-in* → choose `manifest.xml`. Word on the web: Insert → Add-ins
  → *Upload My Add-in*.
- **Organisation-wide:** an admin deploys the same manifest through the
  Microsoft 365 admin centre (*Integrated apps*).

A **Recorde** button (in a group of the same name) appears on the Home tab; it
opens the task pane.

### Supported hosts

| Capability | Requirement set | Word Win/Mac | Word web |
|---|---|---|---|
| Render + preview | — | ✅ | ✅ |
| Insert vector SVG equation | ImageCoercion 1.2 | ✅ | ❌ (raster PNG instead) |
| Baseline-align inline math | WordApiDesktop 1.3 (Win 2507+ / Mac 16.99+, Microsoft 365 only — not LTSC) | ✅ | ❌ (sits slightly high) |
| Click-to-edit, storage, recovery, numbering (inline/table) | WordApi 1.3 | ✅ | ✅ |
| SEQ-field numbering | WordApi 1.5 + desktop | ✅ | ❌ (option disabled) |

A host that supports neither coercion still previews but cannot insert.

## Using the pane

1. Type LaTeX. The preview renders live (errors are shown inline, with the
   offending token); an *Examples* menu gives starting points.
2. Choose **Inline** or **Display**. Display equations go on their own line,
   with TeX spacing above and below, left-aligned (or centred via Settings).
3. Tick **Number this equation** for a display equation. The numbering style
   (Settings → Numbering) is per document: *Flush right (table)* (default:
   the number sits on the right text border), *Inline* (trails the
   equation), or *Word field* (desktop, also flush right). Changing the style offers to convert the existing
   numbered equations; **Renumber all** re-sequences in document order.
4. **Insert** (Ctrl+Enter). Click any inserted equation later to load it back
   into the pane and **Update** it in place; **New equation** (Esc) clears.
5. Settings (collapsible): font, size (match selection / match body / fixed
   pt), colour with swatches, display alignment, and the per-document
   `\newcommand` preamble (applied to every equation on top of the built-in
   macros).

The Examples menu's *Testing* group inserts a one-page test page and a
RevTeX-style three-page test paper (76 equations across every option); the
same paper is produced headlessly by `node scripts/make-showcase-paper.mjs`.

## Headless: equations without Word (`mjx-docx`)

CLI agents and build pipelines never have a running Word. `headless/` ships the
same renderer as an OOXML post-processor: write equations as placeholders while
building a `.docx` with python-docx, docx-js or pandoc, then run the pass.

**Install:** Node.js 20 or newer, then

```
git clone https://github.com/hyperion-git/recorde.git && cd recorde
npm install            # MathJax + fonts + zip/raster libraries (~200 MB with fonts)
npm link               # optional: puts `mjx-docx` on your PATH; else use node headless/bin/mjx-docx.mjs
```

`preview` additionally needs LibreOffice and poppler-utils (`pdftoppm`); the
PNG fallback inside each picture uses `@resvg/resvg-js` (installed with
`npm install`; cairosvg or ImageMagick are used if it is missing). No Word,
no Python, no network access at run time.

**Install the agent skill (Claude Code).** After `npm install` and `npm link`:

```
npm run skill:install                       # → ~/.claude/skills/recorde (all your projects)
npm run skill:install -- --project ~/my-paper   # → that project's .claude/skills/ only
```

It symlinks `skills/recorde` (so `git pull` updates it; `--copy` copies
instead, `--remove` undoes). Claude Code lists it as `recorde` and applies
it whenever a Word document needs formulas; `/recorde` invokes it by hand.
For other agents, paste `skills/recorde/AGENTS-snippet.md` into the
project's `AGENTS.md`.

```
node headless/bin/mjx-docx.mjs process paper.docx [--font termes] [--align left]
node headless/bin/mjx-docx.mjs check paper.docx                    # structural lint (exit 1 = fix first)
node headless/bin/mjx-docx.mjs preview paper.docx                  # LibreOffice → preview/paper-1.png
node headless/bin/mjx-docx.mjs list|update|renumber|render …       # round-trip editing
```

Placeholders: `[[math: …]]` inline, `[[display: …]]`, `[[eq#label: …]]` numbered,
`[[ref: label]]`, with per-equation `{font=,color=,size=,style=,align=}` options.
The output is what the add-in would have inserted — same SVG (PNG fallback),
`mjx:<uuid>` alt-text title, LaTeX description, source in a
`urn:mathjax-office:equations` custom XML part — so equations stay click-to-edit
in Word and `list`/`update` can edit what a human changed in Word. Agent
procedure and gotchas: `skills/recorde/SKILL.md` (also linked from
`.claude/skills/`); plan and status: `headless/ROADMAP.md`.

## How it works

- **Render** — MathJax 4.1.2 SVG output, vendored under `assets/vendor/mathjax`
  with the six font packages; `fontCache: 'none'` inlines every glyph so each
  picture is self-contained. Switching fonts re-bootstraps MathJax with a new
  `output.font`.
- **Insert** — `setSelectedDataAsync` with `XmlSvg` coercion on desktop, a
  300-DPI raster PNG via `Image` coercion on the web. The picture's
  `altTextTitle` is `mjx:<uuid>`, its `altTextDescription` the LaTeX.
- **Storage** — one custom XML part per equation (namespace
  `urn:mathjax-office:equations`) plus parts for the numbering state and the
  document preamble; the payload is also embedded in the SVG's `<desc>`.
  A pasted equation whose part is missing is recovered from its alt text.
- **Baseline** — inline pictures are lowered by the equation's descent
  (`Word.Font.position`, desktop only) so the math baseline meets the text
  baseline without clipping descenders.
- **Numbering** — a document counter; the table style is a borderless 1×2
  table with a 54 pt number cell; the field style a `SEQ equation` field; the
  inline style bakes `(n)` into the picture. `numbering.js` holds the pure
  rules (assignment, migration), `taskpane.js` the Word glue.
- **Pure core** — `src/mathsvg.js`, `numbering.js`, `settings.js`,
  `storage.js` (build/parse), `preamble.js`, `placeholders`/`paragraphs`/
  `ooxml` in `headless/lib` are host-independent and unit-tested with
  `node --test`; the headless CLI imports them directly.

## Development

Needs Node.js 20+, npm, and for the in-Word loop Word for Microsoft 365 on the
same machine (the dev server sideloads the manifest); the showcase-paper script
also needs Python with python-docx.

```sh
npm install            # toolchain + MathJax + fonts (postinstall vendors MathJax)
npm run install-certs  # once: trust the localhost dev certificate
npm start              # serve src/ over HTTPS and sideload into Word
npm test               # unit + headless end-to-end tests (node --test)
npm run validate:manifest
```

Build and deploy: `BASE_URL=https://<host>/<path> npm run build` writes
`dist/` with every URL rewritten and versioned (`?v=<version>`), then
`npm run validate:dist`. The GitHub Pages workflow does this on every push to
`main`. Bump with `npm run bump -- X.Y.Z` (syncs `package.json` and
`manifest.xml`) before deploying so caches refresh. Manual in-Word checks are
listed in `deploy/VERIFY.md`; the release procedure in `docs/RELEASE.md`;
changes in `CHANGELOG.md`.

## Known limitations

- Word on the web inserts raster pictures, cannot shift the baseline, and
  cannot create SEQ fields (existing ones display). On desktop builds older
  than 2507 (or any LTSC build) inline math also sits on its box bottom,
  slightly high; the Settings panel's host line and a start-up notice say so.
- Display alignment, spacing and numbering are Word paragraph/table
  formatting: a user can change them afterwards, and the add-in does not
  re-apply them on update.
- Cross-references inside the pane (`\ref`-style) are not yet available; the
  headless route provides labels and `REF` fields.
- Macros that MathJax 4 cannot express (starred `\Ds`, `\bBigg@` sizes, stretchy
  harpoons, esvect, stackengine) are approximated — see the header of
  `src/preamble.js`.
- Fonts are embedded per equation as SVG paths; very large equations on the
  web path may exceed the picture size limit and are reported.

## Repository

`https://github.com/hyperion-git/recorde` — issues and releases live there; the
hosted build is `https://hyperion-git.github.io/recorde/`.

## License

Apache License 2.0 — see `LICENSE` and `NOTICE`. Bundled MathJax and font data
are covered by `THIRD-PARTY-NOTICES.md` (Apache 2.0, GUST Font License, SIL OFL).
