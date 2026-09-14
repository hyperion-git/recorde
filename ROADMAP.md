# MathJax Office — Development Program

Status (v1.22.1 / M2 complete except pane cross-references, M3 partly): **inserts on Word desktop (vector SVG) and Word on the web (raster)**; colour, palette, macro preamble (paper template), three numbering styles with migration and renumber, TeX display spacing/alignment, self-hosted MathJax, German UI, dark theme, headless `mjx-docx` CLI + agent skill (`headless/ROADMAP.md`). First public release planned as v2.0.0 (`docs/RELEASE.md`).
web (raster PNG)**. LaTeX → SVG/PNG insert, 6 MathJax-4 fonts, click-to-edit,
custom-XML-part storage with alt-text cross-document recovery, settings persistence,
visible error UX, ribbon button, manifest validated against Microsoft's XSD. (v1.0
was desktop-only — superseded.)

**Supported-host reality (verified against Microsoft docs):** the SVG insert path
uses `Office.CoercionType.XmlSvg` (**ImageCoercion 1.2**, Word Win/Mac only). As of
M1 (WP1.5), Word on the web — which only has **ImageCoercion 1.1** — inserts a
raster **PNG** via `Image` coercion instead, so the product now works on both
desktop and web. The inline baseline shift still needs **WordApiDesktop 1.3** (Word
2507+, desktop), so on the web inline math sits slightly high but unclipped.

> **Review provenance.** This plan was cross-reviewed by Codex (gpt-5.5, xhigh)
> against the actual code. That review caught the web-vs-desktop error above
> (since corrected) and two more issues now folded in. Two code fixes were
> applied as a result: the baseline-shift `try/catch` now syncs *inside* the try
> so the `WordApiDesktop 1.3` rejection is actually caught (`taskpane.js`
> `onPrimaryAction`), and MathJax is pinned to **4.1.2** for reproducible
> rendering (`loadMathJax`).

This program takes it from "works on my desktop" to a distributable,
maintainable, multi-host add-in. Four milestones, sequenced by dependency.

## Guiding principles

- **Know your host matrix.** Every insert/format API has a requirement set and a
  host support list. Detect at runtime; degrade visibly. The web ≠ desktop.
- **Deployable before featureful.** You cannot dogfood or share what you can't
  ship. M0 comes first even though M2 is more fun.
- **Guard the pure core with tests before refactoring it.** The sizing/descent
  math and the storage round-trip are where silent corruption hides; lock them
  before features start mutating them.
- **Fail loudly.** Insert/`Word.run`/storage failures must reach the user.
- **One version, bumped deliberately.** Office and static hosts both cache; a
  version bump is the update mechanism. Pin external deps for reproducibility.

## Sizing legend

`S` ≤ half day · `M` ~1–2 days · `L` ~3–5 days · `XL` multi-week / needs a spike.
(Solo, part-time cadence. Relative, not commitments.)

## Release milestones

| Release | Milestone | Theme | Definition of done |
|---|---|---|---|
| v1.0 | — | (done) | Works on desktop Word, manifest valid |
| **v1.1** | **M0** ✅ | Deployable + guarded | Host matrix, hosted, prod manifest, CI, unit tests |
| **v1.2** | **M1** ✅ | Robust + web-capable | Visible errors, settings persist, cross-doc recovery, web raster insert |
| **v2.0** | **M2** | Power features | Palette, numbering, color, macros |
| **v3.0** | **M3** | Platform reach | Multi-host, OMML, offline, localized |

Critical path: **M0 → M1 → (M2 ∥ M3 partially)**. M2 (UX) precedes M3 because
multi-host (WP3.1) inherits whatever the palette/numbering UI becomes.

---

## M0 — Foundation: deployable + guarded  (release v1.1)

### WP0.0 — Supported-host & requirement-set matrix  · `S` · low risk
- **Why:** the whole plan depends on which APIs work where; this was the single
  biggest gap in the first draft. SVG insert is ImageCoercion 1.2 (desktop);
  baseline shift is WordApiDesktop 1.3 (desktop, 2507+); raster insert is
  ImageCoercion 1.1 (incl. Word web).
- **Approach:** a documented matrix (host × requirement set × behavior) in the
  README, plus runtime checks via `Office.context.requirements.isSetSupported(…)`
  that drive Insert availability and the insert path (see WP1.5). No silent
  failures on unsupported hosts.
- **Done when:** opening the pane on any host either enables the correct insert
  path or shows a clear "not supported here" state.

### WP0.1 — Extract pure core + unit tests  · `M` · low risk
- **Why:** the bug-prone logic (ex→pt scaling, viewBox descent, XML escaping) is
  buried inside `taskpane.js` functions that can't be imported.
- **Approach:** lift host-independent helpers into `src/mathsvg.js`
  (`computeRenderSize`, `scaleSvgToPt`, a new `descentPtFromViewBox`,
  `flattenLatex`, `mathJaxFontName`); export `buildXml`/`parseXml` from
  `storage.js`. **Module/DOM setup must be decided:** `package.json` has no
  `"type":"module"` and `dev-server.js` is CommonJS — either rename it `.cjs`
  and add `"type":"module"`, or keep tests as `.mjs`. DOM-touching helpers
  (`parseXml` → `DOMParser`, `scaleSvgToPt`/`inlineColors` → SVG DOM) need a DOM
  shim (`linkedom`/`jsdom`); the rest run in plain Node.
- **Tests (`node --test`):** XML round-trip preserves `< & " '` + newlines;
  `computeRenderSize` clamps a corrupt 500 pt input to body; scaling derives
  width/height from a known viewBox; descent is **below a small threshold** for
  `1+2+3` (not zero — it's metric padding) and large for `\frac`/`\int`;
  `flattenLatex` collapses newlines; `mathJaxFontName` mapping.
- **Done when:** `npm test` is green and the pure core has no `Office`/`Word` refs.

### WP0.2 — Packaging build  · `S` · low risk
- **Approach:** `build.js` reads `BASE_URL`, copies `src/` + `assets/` → `dist/`,
  emits `dist/manifest.xml` with every `https://localhost:3000` replaced. npm
  scripts: `build`, `validate:dist`.
- **Done when:** `BASE_URL=… npm run build` yields a `dist/` that serves
  standalone and whose manifest validates.

### WP0.3 — Hosting + deploy  · `M` · low risk
- **Approach:** `git init`; publish `dist/` to a static host (GitHub Pages /
  Cloudflare Pages). **Validation target is desktop Word** (Win/Mac) for the
  insert path; Word web is validated only for pane + preview until WP1.5 lands.
- **Done when:** the prod manifest sideloads and inserts on desktop Word from
  the public URL.

### WP0.4 — CI  · `S` · low risk
- **Approach:** GitHub Actions on push/PR: install → `node --test` → manifest
  schema-validate (offline XSD — Microsoft's remote validator is flaky; reuse
  the lxml/XSD approach with the reluctant-quantifier patch) → `build`.

### WP0.5 — Versioning & cache strategy  · `S` · low risk
- **Approach:** single version source; bump script syncing `package.json` ↔
  `manifest.xml`; cache-bust assets so updates reach cached clients.
- **Done (partial):** MathJax core pinned to **4.1.2**. *Remaining:* font
  packages are still fetched at MathJax's chosen version — full reproducibility
  arrives with self-hosting (WP3.3); pull it earlier if reproducible rendering
  matters before then.

---

## M1 — Robust + web-capable  (release v1.2) — ✅ SHIPPED

> All five WPs implemented (see `M1-PLAN.md` for the build plan and the
> host-integration checklist to run in real Word). Pure logic is unit-tested
> (`node --test`, 67 cases); the host-integration items flagged there still need a
> Word run to confirm.

### WP1.1 — Visible error & status UX  · `S/M` · low risk
- **Why (corrected):** MathJax-load and font-load failures are *already*
  surfaced in the preview (`bootMathJax`/`onFontChange`), and the baseline
  try/catch is now fixed. The real unhandled gap is the **insert/update path**:
  `insertSvgAtSelection`, the `Word.run` tagging block, and `writeXmlPart` in
  `onPrimaryAction` have no `try/catch` — failures vanish.
- **Approach:** a dedicated notice region in `taskpane.html` (distinct from the
  `#status` edit banner); funnel insert/storage catch blocks through a
  `notify(level, msg)` helper with actionable text (offline, unsupported host,
  picture vanished on update, degraded storage).
- **Done when:** a forced insert failure surfaces a clear message instead of a
  silent no-op.

### WP1.2 — Settings persistence  · `S` · low risk
- **Approach:** persist last font, size-mode, fixed pt, display mode —
  per-user defaults via `localStorage`, optionally per-document last-used via
  `Office.context.document.settings`.
- **Done when:** choosing STIX Two and reopening the pane preselects it.

### WP1.3 — Alt-text source recovery + self-heal  · `S/M` · medium risk
- **Why (re-scoped down from `L`):** the LaTeX is *already* written to
  `pic.altTextDescription` in `onPrimaryAction`, and alt text travels with a
  copied picture. So cross-document recovery does **not** need OOXML spelunking
  first.
- **Approach (Codex design):** in `onSelectionChanged`, load
  `items/altTextTitle,altTextDescription`. Keep the primary `readXmlPartByUuid`
  path; if it returns `null`, recover `latex = altTextDescription.trim()`
  (reject empty/non-equation), build a degraded payload (`font:'tex'`,
  `sizePt:11`-or-body, `displayMode:'inline'`), **self-heal** via `writeXmlPart`,
  and load the UI immediately. Notify: "Recovered source from alt text; font,
  size, and display mode were reset."
- **Metadata:** `font`/`sizePt`/`displayMode` are lost (not in alt text). Do
  **not** pack machine metadata into `altTextDescription` (keep it human-readable
  for screen readers); if metadata recovery becomes important before OOXML,
  extend `altTextTitle` from `mjx:<uuid>` to `mjx:<uuid>;f=stix2;s=11;m=d`.
- **OOXML `<desc>` recovery is deferred** to a later spike — only needed for
  exact metadata, alt-text that's been stripped/edited, or SVG copies whose XML
  part is gone. Not relevant to the raster path (no SVG `<desc>`).
- **Done when:** copy an equation A→B (no shared XML part), click it in B → LaTeX
  loads, a fresh XML part is written, and a second click loads via the XML part.

### WP1.4 — Display-mode paragraph handling  · `S` · low risk
- **Approach:** insert display equations into their own paragraph (split first),
  or only center when the paragraph contains nothing but the equation.
- **Done when:** a display insert mid-paragraph leaves neighboring text aligned.

### WP1.5 — Word-web raster insert fallback  · `M` · medium risk
- **Why:** `XmlSvg` is desktop-only, but ImageCoercion 1.1 (raster `Image`)
  supports Word on the web — this is what makes the add-in usable there.
- **Approach (Codex design):** `detectInsertMode()` via
  `isSetSupported('ImageCoercion','1.2')` → SVG path, else `'1.1'` → PNG path,
  else disable Insert. Refactor `buildSvgForInsert()` to also return
  `widthPt/heightPt`. Add `svgToPngBase64(svg, wPt, hPt)`: pt→px (`×96/72`),
  render at `scale = clamp(devicePixelRatio||2, 2, 4)` into a canvas via a
  same-origin Blob URL, cap max pixels, return base64. Add
  `insertPngAtSelection(base64, wPt, hPt)` using
  `setSelectedDataAsync(base64, { coercionType: Image, imageWidth: wPt,
  imageHeight: hPt })`. Extract `tagSelectedPicture(...)` shared by both paths;
  `writeXmlPart` runs after both. The raster picture has no SVG `<desc>`, so it
  relies on the XML part + alt text (+ WP1.3 for cross-doc).
- **Trade-offs / pitfalls:** raster quality depends on canvas scale; no reliable
  baseline shift on web (prefer visible/unclipped over baseline perfection);
  canvas tainting (safe today — MathJax paths are inlined); image payload size
  limits for huge equations; verify the inserted `inlinePictures` item is still
  selectable for tagging after an `Image` insert. OOXML insertion is **not**
  worth it here — keep it as a later spike only if PNG quality is unacceptable.
- **Done when:** desktop inserts SVG; Word web (only 1.1) inserts a crisp PNG at
  the requested size, tagged and stored, with click-to-edit working.

---

## M2 — Power features  (release v2.0)  — ✅ SHIPPED except WP2.2 native cross-refs in the pane

### WP2.3 — Color  · `S` · low risk — ✅ v1.2 (+ swatches v1.15)
- `inlineColors()` currently forces black. Add a color control, thread it
  through render, store it in the payload for round-trip.

### WP2.1 — Symbol / template palette  · `L` · medium risk — ✅ v1.3 (+ examples menu v1.12)
- The biggest unlock for non-LaTeX users. Collapsible palette (common, Greek,
  operators, structures); snippet model with cursor placeholders (`\frac{▮}{}`)
  inserts at the caret. Pure, testable snippet logic.

### WP2.4 — Macro preamble  · `M` · medium risk — ✅ v1.4 (+ template port v1.7/1.16/1.22)
- Per-document `\newcommand` preamble (XML part), prepended to every render and
  stored with payloads so equations re-render identically after reopen.

### WP2.2 — Equation numbering + cross-references  · `XL` / **spike** · high risk — static ✅ v1.8–1.17 (inline/table/field styles, renumber, migration); native cross-refs: labels + REF fields in the headless route (v1.18), pane UI open
- **Re-classified up.** Split into two:
  - **Static numbering** (`M`): a document-level counter (XML part); render
    right-aligned `(n)`; a "renumber" action (renumber on demand — live stability
    on reorder is the hard part). Depends on WP1.4 paragraph handling.
  - **Word-native cross-references** (`XL`/spike): bookmarks/fields are
    desktop-leaning and largely read-only on Word web, and `(n)` baked into SVG
    gives no native ref behavior. Spike feasibility before committing.

---

## M3 — Platform reach  (release v3.0)

### WP3.3 — Self-hosted MathJax  · `M` · low risk — ✅ v1.5
- Bundle MathJax + needed font packages into `assets/`; drop the jsDelivr
  dependency and its `<AppDomain>`. Also delivers full font-version
  reproducibility (the remaining half of WP0.5) and offline/locked-down support.

### WP3.4 — Localization  · `M` · low risk — ✅ v1.6 (de-DE)
- `<Override Locale="de-DE">` manifest strings (e.g. *Gleichungen*) + a localized
  pane-UI string table.

### WP3.1 — Multi-host (PowerPoint, Excel)  · `L` · medium risk
- The engine is host-agnostic; `XmlSvg` works in PowerPoint (incl. web). Add
  `Presentation`/`Workbook` hosts; branch host-specific storage/selection. Reuse
  M2's palette/numbering UI — hence M2 first.

### WP3.5 — Accessibility  · `M` · medium risk
- Embed assistive MathML alongside the SVG; richer alt text than raw LaTeX.

### WP3.2 — OMML interop  · `XL` · high risk
- MathJax MathML → OMML via `MML2OMML.XSL` to emit *native* Word equations.
  Genuinely involved: shipping/licensing the XSL, running XSLT in the Office
  webview (`XSLTProcessor`), and inserting valid WordprocessingML via OOXML (not
  SVG). Importing native equations *back* to LaTeX is a separate XL with no
  off-the-shelf importer. Scope as opt-in "insert as native equation," not a
  replacement.

---

## Cross-cutting (every milestone)

- **Tests grow with features** — each WP adds `node --test` cases; pure logic is
  mandatory-tested, host integration is manually checked per the README
  checklist (no Office mock until it pays for itself).
- **Docs** — README carries the host matrix + sideload/checklist; this
  `ROADMAP.md` is the living plan.
- **Security/privacy** — `PRIVACY.md` (nothing leaves the document; MathJax and
  fonts are self-hosted since WP3.3, no CDN); support URL = the Pages site; a EULA
  only if ever AppSource-bound.

## Suggested execution order

1. WP0.0 → WP0.1 → WP0.2 → WP0.3 → WP0.4 → WP0.5  *(v1.1: deployable + guarded)*
2. WP1.1 → WP1.2 → WP1.4 → WP1.3 → WP1.5  *(v1.2: robust + web-capable)*
3. WP2.3 → WP2.1 → WP2.4 → WP2.2-static, then the WP2.2-native spike  *(v2.0)*
4. WP3.3 / WP3.4 (quick wins) → WP3.1 → WP3.5 → WP3.2  *(v3.0: reach)*
