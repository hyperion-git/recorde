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
- **Macro preamble** — a physics paper template's macros (operators,
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

## Install

Each [release](https://github.com/hyperion-git/recorde/releases) ships two
packages. Pick the one you need; each contains its own `INSTALL.md`.

| Package | For | Guide |
|---|---|---|
| `recorde-word-addin-<v>.zip` | Writing equations in Word. Contains `manifest.xml`, installer scripts for Windows and Mac, the showcase document. | [docs/INSTALL-word.md](docs/INSTALL-word.md) |
| `recorde-mjx-docx-<v>.zip` | Producing `.docx` files with equations from scripts or coding agents, without Word. Contains the npm package and the agent skill. | [docs/INSTALL-mjx-docx.md](docs/INSTALL-mjx-docx.md) |

**Word (short version):** extract the zip, run `install-windows.cmd` or
`sh install-mac.command`, restart Word — or upload `manifest.xml` via
*Home → Add-ins → More Add-ins → My Add-ins → Upload My Add-in* (also on Word
on the web). Needs Word for Microsoft 365 (desktop or web); the guide has the
host matrix.

**Headless (short version):** Node.js 20+, then
`npm install -g ./recorde-<v>.tgz` from the zip and `mjx-docx skill install`
for Claude Code. From a checkout: `npm install && npm link`.

## Using the pane

1. Type LaTeX. The preview renders live (errors are shown inline, with the
   offending token); an *Examples* menu gives starting points.
2. Choose **Inline** or **Display**. Display equations go on their own line,
   with TeX spacing above and below, left-aligned (or centred via Settings).
3. Tick **Number this equation** for a display equation. The numbering style
   (Settings → Numbering) is per document: *Flush right (table)* (default),
   *Inline*, or *Word field* (desktop). Changing the style offers to convert
   the existing numbered equations; **Renumber all** re-sequences in document
   order.
4. **Insert** (Ctrl+Enter). Click any inserted equation later to load it back
   into the pane and **Update** it in place; **New equation** (Esc) clears.
5. Settings (collapsible): font, size (match selection / match body / fixed
   pt), colour with swatches, display alignment, the per-document
   `\newcommand` preamble, and a host line showing what this Word build
   supports (insert mode, baseline shift, fields).

## Using `mjx-docx`

Write equations as placeholders while building a `.docx` with python-docx,
docx-js or pandoc, then run the pass:

```
mjx-docx process paper.docx [--font termes] [--align left]   # replace placeholders
mjx-docx check paper.docx                                     # structural lint (exit 1 = fix first)
mjx-docx preview paper.docx                                   # LibreOffice → preview/paper-1.png
mjx-docx list|update|renumber|render …                        # round-trip editing
mjx-docx skill install [--project DIR]                        # the agent skill for Claude Code
```

Placeholders: `[[math: …]]` inline, `[[display: …]]`, `[[eq#label: …]]` numbered,
`[[ref: label]]`, with per-equation `{font=,color=,size=,style=,align=}` options.
The output is what the add-in would have inserted — same SVG (PNG fallback),
`mjx:<uuid>` alt-text title, LaTeX description, source in a
`urn:mathjax-office:equations` custom XML part — so equations stay click-to-edit
in Word and `list`/`update` can edit what a human changed in Word. Grammar,
builder recipes and gotchas: `skills/recorde/SKILL.md`.

## Repository layout

```
addin/          the Word add-in: manifest.xml (dev), src/ (pane), assets/ (icons; vendor/ = MathJax,
                gitignored), index.html (support page), build.js, dev-server.cjs, install/ (sideload scripts)
core/           pure modules shared by pane and CLI: mathsvg, numbering, settings, storage, preamble,
                recovery, raster — unit-tested, host-independent
headless/       the mjx-docx CLI: bin/, cli.mjs, lib/ (placeholders, paragraphs, ooxml, zipdoc, render, raster)
skills/recorde/ the agent skill (SKILL.md, AGENTS-snippet.md); shipped in the npm package
examples/       Recorde-Showcase.docx (76 equations, produced headlessly), Recorde-TestPage.docx
docs/           INSTALL-word.md, INSTALL-mjx-docx.md, RELEASE.md, VERIFY.md (in-Word checks),
                ROADMAP.md, ROADMAP-headless.md, release notes, archive/
scripts/        package-release.mjs (the zips), vendor-mathjax.mjs, bump-version.mjs,
                check-manifest.mjs, showcase/fixture generators
test/           node --test suites + fixtures
```

One `package.json` serves both deliverables: the add-in is built into `dist/`
and hosted (GitHub Pages); the CLI is the npm package (`files` whitelists
`core/`, `headless/`, `skills/`). The pane is served from one merged directory,
`<base>/src/` = `addin/src/` + `core/`, so the published URLs never depend on
the repository layout.

## How it works

- **Render** — MathJax 4.1.2 SVG output, vendored under `addin/assets/vendor/mathjax`
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
  baseline without clipping descenders. Measured in Word 2608: Word applies
  the lowering fully, clamped to the picture's height, but seats the box
  *including* the bottom effect extent it adds on SVG import (0.75 pt) on the
  baseline — so the add-in reads that margin back and adds it to the shift.
  The headless emitter writes the run position directly with a zero margin.
- **Numbering** — a document counter; the table style is a borderless 1×2
  table with a 54 pt number cell; the field style a `SEQ equation` field; the
  inline style bakes `(n)` into the picture. `core/numbering.js` holds the pure
  rules (assignment, migration), `addin/src/taskpane.js` the Word glue.

## Development

Needs Node.js 20+, npm, and for the in-Word loop Word for Microsoft 365 on the
same machine (the dev server sideloads the manifest); the showcase-paper script
also needs Python with python-docx.

```sh
npm install            # toolchain + MathJax + fonts
npm run install-certs  # once: trust the localhost dev certificate
npm start              # vendor MathJax, serve addin/ over HTTPS, sideload into Word
npm test               # unit + headless end-to-end tests (node --test)
npm run validate:manifest
```

Build, package and release: `BASE_URL=https://<host>/<path> npm run build`
writes `dist/` with every URL rewritten and versioned (`?v=<version>`);
`npm run package` builds the two zips into `release/`. The Pages workflow
deploys `dist/` on every push to `main`; a `v*` tag makes the release workflow
attach the packages to a draft GitHub release. Bump with `npm run bump -- X.Y.Z`
(syncs `package.json` and `addin/manifest.xml`). Manual in-Word checks:
`docs/VERIFY.md`; release procedure: `docs/RELEASE.md`; changes: `CHANGELOG.md`.

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
  `core/preamble.js`.
- Fonts are embedded per equation as SVG paths; very large equations on the
  web path may exceed the picture size limit and are reported.

## Repository

`https://github.com/hyperion-git/recorde` — issues and releases live there; the
hosted build is `https://hyperion-git.github.io/recorde/`.

## License

Apache License 2.0 — see `LICENSE` and `NOTICE`. Bundled MathJax and font data
are covered by `THIRD-PARTY-NOTICES.md` (Apache 2.0, GUST Font License, SIL OFL).
