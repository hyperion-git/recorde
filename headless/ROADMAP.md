> **Status (2026-09-14, v1.18.0):** the headless toolset lives in this repo under `headless/`
> (`mjx-docx` CLI: process/render/list/update/renumber/check/preview) with the skill in
> `skills/recorde/`. Done: H0.3 headless render (worker per font), H1.1–H1.4, H2.1, H2.2,
> the skill text of H2.3. Since v1.18: per-equation placeholder options (font/color/size/style/align),
> TeX display skips + fleqn/centre alignment, column-aware autofit numbering tables, REF-field references
> refreshed by `renumber`, Termes default, and the showcase paper built from the add-in's own source
> (`scripts/make-showcase-paper.mjs`). NOT done: H0.1 golden fixture from desktop Word (the OOXML shapes are
> spec-derived and LibreOffice-checked only — verify in Word, then pin), H0.2 monorepo split
> (the CLI imports the pure modules straight from `src/`), the H2.3 agent trial, H3.
> Plan text below is the original and unchanged.

# MathJax Office — Headless toolset & agent skill

Status: **planning (H0 not started)**. Companion to the Word add-in in
the add-in in this repository (then v1.10.0). Goal: let a CLI agent (Claude Code, Codex, …)
place MathJax equations into Office documents it assembles, with output that is
**indistinguishable from what the add-in inserts** — same SVG, same alt text,
same custom XML part — so the documents stay click-to-edit in Word and equations
round-trip in both directions.

## The core observation

CLI agents never have a running Word. The document skills they use (the `docx`
skill from the scientific-skills marketplace is the reference on this machine)
all work the same way: **build new files with docx-js or python-docx, edit
existing ones by unpack → edit `word/document.xml` → repack, preview via
LibreOffice → PDF, read via pandoc.** Office.js — the add-in's entire integration
surface — does not exist in that loop. The integration point is therefore OOXML,
and the headless tool must reproduce in files what the add-in does through the
API:

| Add-in (Office.js) | Headless (OOXML) |
|---|---|
| MathJax in the task-pane webview | MathJax 4.1.2 in Node via `mathjax/node-main.mjs` (liteDOM adaptor); fonts already in `node_modules/@mathjax/*` |
| `setSelectedDataAsync(…, XmlSvg)` | `w:drawing/wp:inline` with a PNG `a:blip` plus the `asvg:svgBlip` extension; `word/media/imageN.{svg,png}` |
| Canvas PNG fallback (`src/raster.js`) | `@resvg/resvg-js` rasterization (no browser) |
| `pic.altTextTitle = 'mjx:<uuid>'`, `altTextDescription = latex` | `wp:docPr title="mjx:<uuid>" descr="<latex>"` |
| `Font.position = -descentPt` | `w:rPr/w:position w:val="-<2·descentPt>"` (half-points) |
| Own paragraph, centered (WP1.4) | Fresh `w:p` with `w:jc w:val="center"` |
| Numbered: 1×2 borderless table / SEQ field (WP2.2) | Same `w:tbl` XML; `w:fldSimple w:instr=" SEQ equation \* ARABIC "` |
| `customXmlParts.add(xml)` | `customXml/itemN.xml` + `itemPropsN.xml` + rels + content-type override |
| `customXmlParts.getByNamespace` + alt-text scan (`readAllEquations`) | Walk `document.xml` for tagged pictures, join with XML parts by uuid |

The last two rows are the payoff: an agent-built document opens in Word and the
add-in loads its source on click; a human-edited document can be listed and
updated by the agent.

## Guiding principles

- **One renderer.** The add-in and the CLI consume the same core package. Any
  render difference is a bug, not a feature.
- **Golden fixture before code.** Every OOXML shape is verified against a
  document *Word itself wrote* via the add-in, not against the spec.
- **Post-process, don't build.** The agent keeps using whatever docx tool it
  already knows; math is one extra pass over the finished file. No new document
  builder to learn, no dependency on a specific language.
- **Acceptance = round-trip in real Word.** Not "the XML validates".
- **The skill carries procedure, not code.** Decision tree + gotchas; all
  mechanics live in the CLI.

## Architecture: three layers

```
packages/core   pure render + payload logic (shared)      ← extracted from src/
packages/cli    `mjx-docx` — OOXML post-processor         ← new
skills/         SKILL.md + AGENTS.md snippet              ← new
packages/addin  the existing Word add-in                  ← moved, unchanged
```

**Repo layout decision (make in H0.2):**

> **Recommended:** convert this repository into an npm-workspace monorepo
> (`packages/core|addin|cli`, `skills/`). One renderer version, one `bump`, the
> existing CI/Pages deploy keep working (`build.js` just gains a path prefix).
> This directory then becomes redundant, or holds only the skill.
> **Considered:** keep the add-in repo as is and make this directory a separate
> CLI repo depending on it (git dependency). Two version streams for one
> renderer; every core fix needs a cross-repo bump. Only worth it if the add-in
> repo must stay "pure add-in" for distribution reasons.

### Core API (packages/core)

```js
renderEquation({ latex, font, sizePt, displayMode, color, preamble })
  → { svg, widthPt, heightPt, descentPt }        // pure; no PNG, no uuid
buildEquationPayload({ uuid, latex, font, sizePt, displayMode, color, numbered, number, numberStyle })
  → xml string                                   // = storage.buildXml
parseEquationPayload(xml) → payload              // = storage.parseXml
```

Moves, unchanged: `mathsvg.js`, `preamble.js`, `numbering.js` (pure parts),
`recovery.js`, the `build*/parse*` halves of `storage.js`, and the pure
constants of `raster.js`. Stays in the add-in: everything that touches `Word.*`,
`Office.*`, `localStorage`, the DOM of the pane. `DOMParser`/`XMLSerializer`
come from linkedom in Node (the tests already do this), `crypto.randomUUID` is
native.

### CLI surface (packages/cli, bin `mjx-docx`)

| Command | Purpose |
|---|---|
| `mjx-docx process in.docx [-o out.docx]` | Replace placeholders with equations; add media, parts, numbering. **The one command an agent needs.** |
| `mjx-docx render 'E=mc^2' [--display] [--font termes] [--size 11] [--svg|--png|--json]` | Emit SVG/PNG/metrics for agents that build their own XML |
| `mjx-docx list file.docx [--json]` | Enumerate equations in document order: uuid, latex, mode, font, size, number |
| `mjx-docx update file.docx --uuid U --latex '…'` | Re-render one equation in place (media, alt text, XML part, position) |
| `mjx-docx renumber file.docx` | Reassign static numbers in document order; refresh SEQ fields |
| `mjx-docx preview file.docx` | LibreOffice → PDF → PNG pages into `./preview/` for a visual check |
| `mjx-docx check file.docx` | Lint: orphan XML parts, pictures without parts, duplicate uuids, content-type/rels consistency |

Document-level defaults (`font`, `sizePt`, `color`, `numberStyle`, `preamble`)
come from flags or an optional `mjx.json` next to the document; the preamble is
also written to the `urn:mathjax-office:preamble` part so the add-in re-renders
identically.

### Placeholder syntax (agent contract)

```
[[math: \hbar\omega]]                inline
[[display: \int_0^\infty e^{-x^2}\,dx = \tfrac{\sqrt\pi}{2}]]   own centered paragraph
[[eq: i\hbar\partial_t\psi = H\psi]]                             numbered display
[[eq#schrodinger: …]]                numbered + label
[[ref: schrodinger]]                 → "(3)" — static text, or REF field to a bookmark (field style)
```

- Regex: `\[\[(math|display|eq)(#[\w:-]+)?:\s*(.*?)\]\](?!\])` — the
  `(?!\])` keeps `f(x)=[a,b]]]` intact. No nesting.
- Matched on **joined paragraph run text**, then spliced back into runs —
  docx-js and Word both fragment text across `w:r` unpredictably.
- `--dollar` opt-in accepts `$…$` / `$$…$$` for pandoc/markdown authors (off by
  default: currency).
- Pandoc route: `pandoc -f markdown-tex_math_dollars …` leaves the placeholders
  as literal text; `process` does the rest. (Default pandoc would emit OMML.)

### The skill (skills/recorde/SKILL.md)

Procedure the agent follows, in this order:

1. Build the document with your usual tool (docx-js / python-docx / pandoc);
   write equations as placeholders, one placeholder per run where you control
   runs.
2. `mjx-docx process`. Never hand-write drawing or customXml XML.
3. `mjx-docx preview` and look at the pages; `mjx-docx check` before delivery.
4. To edit an existing document's equations: `list` → `update`, never touch
   `word/media` by hand.

Gotchas the skill must carry: inline pictures sit on the block bottom without
the `w:position` shift (CLI handles it — don't "fix" it); display equations need
their own paragraph (CLI splits it); LibreOffice preview ≠ Word rendering for
baseline; OMML from pandoc is a *different, non-editable* equation type — never
mix.

Discovery: Claude Code via `.claude/skills/` (repo) or a plugin; Codex via its
skills directory (`.agents/skills/` / `~/.agents/skills/` — **verify against the
installed Codex version**) plus an `AGENTS.md` paragraph. One canonical folder,
symlinked.

## Sizing legend

`S` ≤ half day · `M` ~1–2 days · `L` ~3–5 days · `XL` multi-week / spike.
(Solo, part-time cadence.)

## Milestones

| Milestone | Theme | Definition of done |
|---|---|---|
| **H0** | Ground truth + shared core | Golden fixture committed; core package extracted; add-in tests still green |
| **H1** | `process` end-to-end | Agent-built .docx with inline/display/numbered equations opens in Word and the add-in edits them |
| **H2** | Round-trip + skill | `list`/`update`/`renumber` on a Word-edited document; skill used successfully by one agent session |
| **H3** | Reach | PowerPoint, Python wrapper, MCP wrapper — only as demand appears |

Critical path: H0.1 → H0.2 → H1.1 → H1.2 → H1.3 → H1.4 → H2.

---

## H0 — Ground truth + shared core

### H0.1 — Golden fixture  · `S` · low risk · **do first**
- **Why:** everything downstream is "match what Word writes". Without a real
  file the OOXML details (svgBlip ext, EMU extents, itemProps GUID, rels,
  content types, `w:position` sign) are guesswork.
- **Approach:** in desktop Word with the add-in, insert into a fresh document:
  one inline (with descender, e.g. `\int_0^1 y\,dy`), one display, one numbered
  (table style) and one numbered (field style) equation, plus a preamble and a
  non-default font. Save. `unpack.py` it (docx skill) and commit the tree as
  `test/fixtures/golden-addin/`. Also commit one Word-*web* file (PNG path, no
  SVG) as `golden-web/` — `list` must handle it.
- **Done when:** both fixtures are committed with a `NOTES.md` recording Word
  version and add-in version.

### H0.2 — Workspace + core extraction  · `M` · low risk
- **Approach:** npm workspaces; move `src/` → `packages/addin/src/`, pure
  modules → `packages/core/src/`; add-in imports from `@recorde/core`.
  `build.js`, `vendor-mathjax.mjs`, `bump-version.mjs`, CI, Pages deploy adjust
  paths only. Bump sync now covers three `package.json` + `manifest.xml`.
- **Tests:** existing suites move with their modules; `npm test` at the root runs
  all workspaces. Add a test that `packages/core` has zero references to
  `Office`, `Word`, `window`, `document`, `localStorage` (grep-based guard).
- **Done when:** `npm test` green in all packages; `BASE_URL=… npm run build`
  still yields a valid `dist/`; the pane still inserts in Word (manual).

### H0.3 — Headless render  · `S/M` · medium risk
- **Approach:** `renderEquation()` boots MathJax through `node-main.mjs` with
  `output.font` = the requested font package, `fontCache: 'none'`, then applies
  `scaleSvgToPt` / `inlineColors` / `descentPtFromViewBox` on a linkedom
  element. PNG via resvg at `TARGET_DPI` (reuse `computeRasterSize`).
- **Risk:** font switching in Node — the add-in re-bootstraps MathJax per font;
  in Node each font may need its own MathJax instance or a per-process font
  (measure; cache instances per font). SRE speech worker must be disabled
  headless.
- **Tests:** rendered SVG for `1+2+3` matches the add-in's viewBox for the same
  input (extract from the golden fixture, compare numerically); descent for
  `\int` > threshold; PNG dimensions = pt × DPI/72.
- **Done when:** `mjx-docx render` produces an SVG byte-identical (modulo
  `<desc>` and ids) to the fixture's `media/image1.svg`.

---

## H1 — `process` end-to-end

### H1.1 — OOXML writer primitives  · `M` · medium risk
- **Approach:** a small zip-aware document model (JSZip or `node:zlib` +
  own zip; keep deps minimal): read/write parts, add a media file, add a
  relationship, add a content-type default/override, append a customXml item
  (`itemN.xml`, `itemPropsN.xml` with `ds:datastoreItem ds:itemID="{GUID}"` and
  `ds:schemaRef ds:uri="urn:mathjax-office:equations"`, `_rels/itemN.xml.rels`,
  document rel of type `…/relationships/customXml`). Preserve everything else
  byte-for-byte (Word is fussy about re-serialized XML; treat unknown parts as
  opaque bytes).
- **Reference:** the golden fixture — diff each emitted part against it.
- **Tests:** add-then-read round-trip; a document that already has customXml
  items (docx-js emits none, Word templates often do) gets the next free index;
  content types not duplicated.
- **Done when:** a fixture-derived document with one part *removed* is
  reconstructed identically by the writer.

### H1.2 — Equation drawing emitter  · `M` · medium risk
- **Approach:** `emitInlineEquation({svgRelId, pngRelId, widthPt, heightPt,
  descentPt, uuid, latex, displayMode})` → `w:r` (with `w:position` for inline)
  containing `w:drawing/wp:inline` (`wp:extent` in EMU = pt × 12700, `wp:docPr`
  title/descr, `a:blip r:embed=png` + `a:extLst/a:ext uri="{96DAC541-7B7A-43D3-8B79-37D633B846F1}"/asvg:svgBlip r:embed=svg`).
  Display: wrap in its own centered `w:p`; numbered: 1×2 borderless table with
  `NUMBER_CELL_WIDTH_PT` right cell, static `(n)` or SEQ `fldSimple`.
- **Tests:** emitted XML equals the fixture's drawing modulo rel ids and uuid
  (normalize, then string-compare); EMU rounding; `w:position` = −round(2·descentPt).
- **Done when:** the emitter reproduces all four fixture equations.

### H1.3 — Placeholder pass  · `M` · medium risk
- **Approach:** for each `w:p`, join `w:t` text with run indices; regex; split
  the affected run(s) at match boundaries preserving `w:rPr`; insert the
  emitted run/paragraph/table; for `display`/`eq` in a paragraph with other
  text, split the paragraph around it (WP1.4 semantics, using
  `isParagraphTextEmpty`). Labels → bookmark around the number; `[[ref:]]` →
  static number (default) or `REF` field.
- **Edge cases (tests):** placeholder split across three runs; two placeholders
  in one paragraph; placeholder inside a table cell; inside a header/footer
  (skip with warning); `]]` inside LaTeX; unknown label → error with paragraph
  text quoted; document with tracked changes (`w:ins`/`w:del`) → refuse unless
  `--force`.
- **Done when:** a docx-js-built and a python-docx-built test document both
  process into files whose equations the add-in opens on click (manual, desktop
  Word). Record the result in `test/fixtures/NOTES.md`.

### H1.4 — `preview` + `check`  · `S` · low risk
- **Approach:** `soffice --headless --convert-to pdf`, `pdftoppm -r 150`;
  `check` walks rels/parts/pictures and reports inconsistencies.
- **Verify:** LibreOffice honours `svgBlip` (≥ 7.0 should) and how it treats
  `w:position` on a picture run. Document the difference to Word in the skill.
- **Done when:** the H1.3 documents preview with visible, correctly sized
  equations, and `check` is clean on the golden fixtures.

---

## H2 — Round-trip + skill

### H2.1 — `list`  · `S` · low risk
- Walk `document.xml` (body order, incl. tables) for `wp:docPr` with
  `title^="mjx:"`; join with parsed XML parts by uuid; fall back to
  `buildDegradedPayload` from `descr` (mirrors `recoverFromAltText`). Works on
  the web fixture (PNG only).

### H2.2 — `update` + `renumber`  · `M` · medium risk
- `update`: re-render, replace the two media files in place (same rel ids),
  rewrite `wp:extent`, `w:position`, `descr`, and the XML part (delete-then-add,
  like `writeXmlPart`). `renumber`: mirror `assignNumbers` over document order;
  static numbers rewritten, SEQ fields left to Word (`w:dirty="true"` on
  `updateFields` setting so Word refreshes on open).
- **Done when:** an equation inserted by the add-in, updated by the CLI, reopens
  in the add-in with the new source; and vice versa.

### H2.3 — Skill + agent trial  · `S/M` · low risk
- Write `SKILL.md` (procedure above), an `AGENTS.md` snippet, and an example
  session transcript. Then run one real task in a fresh Claude Code session and
  one in Codex: "write a two-page note on the harmonic oscillator as .docx with
  numbered equations". Capture what the agent got wrong and fold it into the
  skill's gotchas.
- **Done when:** both agents produce a document that passes `check` and opens
  with editable equations, without human intervention beyond the prompt.

---

## H3 — Reach (only on demand)

- **H3.1 PowerPoint** · `L` — same drawing XML in `p:pic`, per-slide
  placement; no customXml-by-namespace API parity on the add-in side yet
  (add-in WP3.1 first).
- **H3.2 Python wrapper** · `S` — `mjx_docx.process(path)` shelling out to the
  CLI, for python-docx-centric pipelines. Not a reimplementation.
- **H3.3 MCP server** · `S` — thin wrapper exposing `render`/`process`/`list`
  as tools. Zero new logic.
- **H3.4 OMML export** · `XL` — only if a consumer needs native equations;
  tracked as add-in WP3.2, shared MathML → OMML XSL step.

## Cross-cutting

- **Fixtures are the spec.** When Word and the plan disagree, Word wins; update
  the plan.
- **Tests grow with commands**; every CLI command has a fixture-based test;
  manual Word checks are logged in `test/fixtures/NOTES.md` with date + version.
- **Versioning:** core, add-in and CLI share one version via `bump`; the CLI
  writes its version into each XML part's `<equation>` as a `writer` attribute
  (optional attr — the add-in's `parseXml` already ignores unknowns).
- **Dependencies:** MathJax 4.1.2 (pinned), linkedom, resvg-js, one zip lib.
  No headless browser.

## Open questions (decide during H0)

1. Monorepo vs. sibling repo (see layout decision above).
2. Does Word require the PNG fallback next to the SVG, or does `svgBlip` alone
   render? (Fixture answers it — check whether the add-in's `XmlSvg` insert
   produced a PNG.) Keep the fallback regardless for older readers.
3. Namespace lookup: does `customXmlParts.getByNamespace` key on the root
   element namespace or on `itemProps` `schemaRef`? Emit both, but the fixture
   tells us what Word writes.
4. Default `numberStyle` for agent output: `table` (renders everywhere) vs
   `field` (native, desktop-only). Leaning `table`.

## Suggested execution order

1. H0.1 → H0.2 → H0.3  *(ground truth, shared core, headless render)*
2. H1.1 → H1.2 → H1.3 → H1.4  *(`process` works in real Word)*
3. H2.1 → H2.2 → H2.3  *(round-trip, skill, agent trial)*
4. H3.x as needed
