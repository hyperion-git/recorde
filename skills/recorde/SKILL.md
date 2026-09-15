---
name: recorde
description: Put LaTeX equations into .docx files you build or edit headlessly (python-docx, docx-js, pandoc), as MathJax SVG pictures that stay click-to-edit in Word with the Recorde add-in. Use whenever a Word document needs formulas and no Word is running.
---

# MathJax equations in .docx without Word

You write equations as **placeholders** in the text while building the document
with whatever tool you already use; then one command replaces them with rendered
equation pictures. Never hand-write drawing or customXml XML.

## Prerequisites

Node.js 20+ and a checkout of the Recorde repository
(`https://github.com/hyperion-git/recorde`) with `npm install` and `npm link`
run once: the first fetches MathJax, the six font packages and the zip/raster
libraries, the second puts the `mjx-docx` command on the PATH — every command
below assumes it. Without `npm link`, replace `mjx-docx` by
`node $RECORDE_DIR/headless/bin/mjx-docx.mjs`. Optional: LibreOffice +
poppler-utils for `preview`; Python with python-docx only if you build
documents that way. Word is not needed and nothing is fetched at run time.
Not sure it is installed? `mjx-docx --version` prints `mjx-docx/<version>`.

## Procedure

1. Build the document as usual (python-docx, docx-js, pandoc, or edit an
   existing file). Put each equation in a placeholder — see the grammar below.
   Keep a placeholder inside one run when you control runs; splitting across
   runs is tolerated, but never split it across paragraphs.
2. Run the equation pass:

       mjx-docx process paper.docx [--font termes] [--number-style table] [--align left]

   It rewrites the file in place (`-o out.docx` to keep the original). Exit 1 =
   nothing was written; read the `error:` lines (unknown label, empty
   placeholder, tracked changes).
3. Check before delivery:

       mjx-docx check paper.docx      # exit 1 on structural errors
       mjx-docx preview paper.docx    # LibreOffice → preview/paper-1.png …

   Look at the page images. LibreOffice ignores the inline baseline shift and
   spaces tables differently from Word — judge content, not fine layout.
4. To change equations in a document that already has them (yours or one a
   human edited in Word): `list` → `update --uuid … --latex …` → `check`.
   Run `renumber` after reordering or deleting numbered equations.

## Placeholder grammar

| Placeholder | Result |
|---|---|
| `[[math: \hbar\omega]]` | inline equation in the running text, sized like the run it replaces |
| `[[display: \int_0^\infty e^{-x^2}\,dx]]` | display equation on its own centred paragraph |
| `[[eq: i\hbar\partial_t\psi = H\psi]]` | numbered display equation (`(1)`, `(2)`, …) |
| `[[eq#schrodinger: …]]` | numbered + label |
| `[[ref: schrodinger]]` | the label's number, `(3)`, as a REF field (forward references fine) |
| `[[eq#lbl{font=stix2,color=#1f77b4,size=12,style=field,align=center}: …]]` | per-equation options (any kind but `ref`): `font`, `color`, `size` (pt), `style` (numbered only), `align` (display only) |

- The match ends at the first `]]` not followed by another `]`, so
  `[[math: f(x) = [a,b]]]` is fine.
- Display and numbered placeholders that share a paragraph with text split the
  paragraph: text before, the equation block, text after.
- `--dollar` additionally accepts `$…$` and `$$…$$` (off by default: currency).
- Display alignment (`--align`, `mjx.json` `align`, or `{align=…}`): `left` is
  the default — LaTeX's `fleqn`, flush left with the 25 pt `\mathindent`;
  `center` gives the plain-LaTeX centred look. The Word add-in has the same
  choice under Settings → Display.
- Vertical spacing follows TeX: a display gets `\abovedisplayskip` =
  `\belowdisplayskip` = one body font size (10 pt at 10 pt, 12 at 12) as
  paragraph spacing, and the text after it continues the paragraph without a
  first-line indent. Do not add empty paragraphs around placeholders to "make
  room" — that doubles the gap.
- Two-column sections: numbering tables fit the column automatically, but an
  equation wider than the column is clipped. Break long equations with
  `\begin{aligned} … \\ … \end{aligned}` (about 19 em per line at 10 pt in a
  3.4-inch column). The showcase paper (`node scripts/make-showcase-paper.mjs`
  in the Recorde repo)
  is the reference for what fits.
- Numbering styles (`--number-style`): `table` (default; number flush right in a
  borderless 1×2 table, renders everywhere), `inline` (number baked into the
  picture), `field` (Word SEQ field; desktop Word updates it, Word web shows
  the cached value).
- Fonts: `termes` (TeX Gyre Termes, the Times clone closest to newtxmath;
  default), `tex` (Computer Modern), `newcm`, `stix2`, `pagella`, `asana`. `--size <pt>` overrides the per-run size; `--color #rrggbb`.
- Document macros: `--preamble macros.tex` (`\newcommand` lines) — stored in the
  document so the Word add-in renders identically.
- Per-document defaults: an `mjx.json` next to the file
  (`{"font":"termes","numberStyle":"table","preamble":"macros.tex"}`).

## Builder recipes

**python-docx** — a placeholder is just text: `p.add_run('[[math: E=mc^2]]')`.
For an inline equation the run's font size becomes the equation size.

**pandoc (markdown)** — pandoc's markdown reader treats `\x` as escapes and
drops raw TeX. Wrap every placeholder in a code span so its bytes survive:
`` `[[math: \sqrt{2}]]` `` and a display on its own line:
`` `[[display: \sum_k k]]` ``. Then `pandoc -f markdown -o out.docx in.md` and
`process`. Do **not** use pandoc's own `$…$` math: it emits native OMML
equations, a different, non-editable-by-the-add-in type — never mix.

**docx-js** — `new TextRun('[[eq#x: …]]')`; keep one placeholder per TextRun.

## What the pass writes (so you know what not to touch)

Each equation is an inline picture (SVG with a PNG fallback in `word/media/`)
whose alt-text title is `mjx:<uuid>` and description is the LaTeX; its source
lives in a `customXml/itemN.xml` part in namespace
`urn:mathjax-office:equations`. The Recorde Word add-in keys on exactly these,
so equations round-trip: click one in Word to edit it; run `list`/`update` here
to edit it headlessly. Never edit `word/media` or `customXml` by hand — use
`update`, which keeps the ids stable.

## Gotchas

- Inline equations sit on the baseline only in Word (via a run position shift
  the pass writes: `w:position` = −2·descent half-points, zero effect extent;
  Word lowers a picture by that value, clamped to the picture's height).
  LibreOffice previews show them slightly high — expected.
- A TeX error does not stop the pass: the equation is inserted with MathJax's
  error mark and a `warning:` is printed. Fix the LaTeX and `update` it.
- Placeholders in headers/footers and text boxes are not processed (warned).
- Documents with tracked changes are refused (`--force` to override; the pass
  ignores revision marks, so accept changes first).
- `renumber` reassigns static numbers in document order and refreshes SEQ and
  REF cached results; Word's own field update (F9) agrees with it.
- `check` exit 1 means Word may "repair" the file on open — fix before sending.
