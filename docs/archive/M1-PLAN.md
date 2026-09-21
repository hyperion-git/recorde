# M1 Implementation Plan (Robust + web-capable → v1.2)

> Generated from a multi-agent design pass (5 per-WP designers + adversarial
> blocker verifier + synthesis), all cross-checked against current Microsoft
> Learn / office-js docs. This is the execution spec; ROADMAP.md is the program.

## 0. Load blocker (S0) — ships first, gates dogfooding

**Root cause (confirmed):** `build.js` rewrites the dev origin but never the
`<Id>`, so the deployed/sideloaded manifest inherits the **dev `<Id>`**
`2a7c1e3a-9f4b-4c1d-8e1f-3c2a5b9d7e10` (verified live) and `<Version>1.0.0.0</Version>`.
Office keys add-in *registrations* by `<Id>` (desktop: `HKCU\…\WEF\Developer`
registry + `CustomUIValidationCache`; **web: browser local storage** — *not*
server-side). Sideloading a new manifest under an Id already mapped to the dead
`https://localhost:3000` dev entry collides → "Add-in can't be loaded". This
explains desktop for certain; web iff that browser profile previously sideloaded
the same Id. (ES-module / URL / CSP hypotheses were *refuted* by evidence:
correct `application/javascript` MIME, all assets 200, no XFO/CSP.)

**Fix:**
- `build.js`: add `DEV_ID`/`PROD_ID` constants; after origin rewrite, assert the
  dev Id is present then `manifest.split(DEV_ID).join(PROD_ID)`.
  `PROD_ID = 9971ece4-30c4-421f-920e-8df1563b6bd5` (generated once; keep STABLE so
  redeploys update the same prod add-in in place — do **not** derive from URL).
- `scripts/check-manifest.mjs`: add a dist-only guard that fails if the built
  manifest still contains the dev Id.
- Run `npm run bump -- 1.1.0` (package.json → 1.1.0, manifest `<Version>` → 1.1.0.0;
  bump-version.mjs already correct — the miss was operational).

**Confirm/refute:** after rebuild + cache-clear + re-sideload, the pane loads on
both hosts ⇒ confirmed. If it *still* fails with a brand-new unique Id + cleared
cache ⇒ refuted; attach DevTools and read the real console error (next suspect:
an unguarded `isSetSupported`/insert path throwing on web, or office.js init).

## 1. Shared refactors first (behavior-preserving; one commit)

All five WPs edit `src/taskpane.js` (esp. `onPrimaryAction`, `onSelectionChanged`).
Land these refactors first so each WP is a localized add.

- **S1 — empty pure modules:** `src/errors.js`, `src/settings.js`,
  `src/recovery.js`, `src/raster.js`. WP1.4 instead adds one export to `mathsvg.js`.
- **S2 — `notify()` funnel:** keep `notify('info'|'warn'|'error', msg)`; add
  `clearNotice()`, `notifyResult({level,msg})`, and a `noticeFromAppFlow` flag.
  CSS for all three levels already exists in `taskpane.html` — no HTML change.
- **S3 — `insertMode` supersedes `canInsertSvg`:** `let insertMode = 'none'`
  (`'svg'|'png'|'none'`) + `detectInsertMode(req)` (1.2→svg, 1.1→png, else none).
  Update `Office.onReady`, `setMathJaxBusy`, `onPrimaryAction`, and the onReady
  notice (3 arms). **Gate the `'png'` arm OFF (return 'none' for 1.1-only) until
  WP1.5 lands** so web never shows an enabled-but-broken Insert in the interim.
- **S4 — extract `tagSelectedPicture(uuid, latex, descentPt, displayMode,
  allowBaselineShift)`:** lift the alt-text-tag + baseline-shift `Word.run` out
  of `onPrimaryAction`; shared by SVG + PNG paths. Returns true if a pic was tagged.

## 2. Order

S0 → S1–S4 → **WP1.1 → WP1.2 → WP1.4 → WP1.3 → WP1.5**.

| WP | New module + test | taskpane.js edits |
|----|-------------------|-------------------|
| 1.1 | errors.js / errors.test.js | installGlobalErrorHandlers() first in onReady; wrap onPrimaryAction steps (hard-abort insert+update-locate; soft-warn center/tag/store) |
| 1.2 | settings.js / settings.test.js | applySettingsToUi(loadSavedSettings()) between wireUi & bootMathJax; saveCurrentSettings() on success path |
| 1.4 | isParagraphTextEmpty in mathsvg.js + cases | ensureOwnParagraphForDisplay() pre-split (new display inserts only); emptiness-gated center |
| 1.3 | recovery.js / recovery.test.js | onSelectionChanged: load altTextDescription; recover+self-heal when XML part null |
| 1.5 | raster.js / raster.test.js | flip png arm on; buildSvgForInsert returns w/h pt; svgToPngBase64, insertPngAtSelection, tagLastBodyPicture; branch onPrimaryAction on insertMode |

## 3. Module signatures (the integration contract)

- **errors.js:** `isLikelyOffline(error)`, `describeInsertError(step, error) →
  {level,msg}` (step ∈ insert|update-locate|tag|store; offline wins over code;
  unknown code → error + appended raw message; never throws),
  `pictureVanishedNotice() → {level:'info',msg}`.
- **settings.js:** `DEFAULT_SETTINGS = {font:'tex',sizeMode:'selection',fixedPt:11,
  displayMode:'inline'}`, `KNOWN_FONTS = [tex,newcm,termes,stix2,pagella,asana]`,
  `parseSettings(raw)` (null/''/bad-json/object → normalized or defaults),
  `serializeSettings(obj)` (strips garbage), `normalizeSettings(obj)`
  (font/sizeMode whitelist; fixedPt clamp [6,96] w/ string coercion; non-object →
  defaults).
- **recovery.js:** `looksLikeEquation(s)` (non-empty, plausible, length-capped),
  `buildDegradedPayload({uuid,latex,sizePt}) → {uuid,latex(trimmed),font:'tex',
  sizePt(clamped→11),displayMode:'inline',degraded:true}` (the `degraded` key must
  NOT leak into buildXml — it only reads uuid/latex/font/sizePt/displayMode).
- **raster.js:** `PX_PER_PT = 96/72`, `MAX_BASE64_LENGTH`, `ptToPx(pt)`,
  `clampScale(s)` (→[2,4], default 2), `computeRasterSize(wPt,hPt,dpr,maxPixels?)
  → {scale,pxW,pxH}` (≥1px; reduce scale to fit budget), `stripDataUri(s)`,
  `isPayloadTooLarge(base64)`.
- **mathsvg.js:** `isParagraphTextEmpty(text)` (true for null/undefined/whitespace
  incl. NBSP/ZWSP/BOM/VT; false for real text).

## 4. Conflict map — `onPrimaryAction` final shape (see workflow output for full body)

Step labels are the shared vocabulary: `insert`, `update-locate`, `tag`, `store`.
Hard-abort (`return`) only on insert (SVG/PNG) and update-locate; everything
post-insert is soft (warn + continue) so a placed picture is never erased. WP1.4
pre-split sits *before* the insert try (gated `!editingUuid`); WP1.5 PNG branch
sits *inside* WP1.1's insert try; WP1.2 `saveCurrentSettings()` is the last
success line. `onSelectionChanged` is touched by WP1.3 only; never call
`saveCurrentSettings()` there (equation metadata ≠ user default).

## 5. Open risks (verify in real Word — gated on S0)

1. **WP1.4 caret split semantics** — `insertParagraph('','Before')` collapsed-caret
   landing isn't doc-pinned. Highest integration risk; spike on desktop. Mitigation:
   explicit `getRange('Start').select()` re-collapse.
2. **WP1.5 PNG `imageWidth/Height` unit (pt vs px)** — office-js #2714 confirmed on
   PPT-web, unconfirmed on Word-web. Pass points (per docs); one-line ×72/96 fix if
   web shows 4/3 oversizing.
3. **WP1.5 picture reachable via `getSelection().inlinePictures` after Image insert**
   — `tagLastBodyPicture` body-scan fallback (assumes newest = last in body).
4. **WP1.5 web edit/replace** — confirm `setSelectedDataAsync(Image)` replaces a
   selected picture rather than inserting beside (else duplicates on update).
5. **`document.settings.get` null across versions (office-js #1585)** — handled by
   defaulting; per-doc settings best-effort.

## 6. Test plan (all pure-logic tests automatable now via `node --test`)

New: `test/errors.test.js`, `test/settings.test.js`, `test/recovery.test.js`,
`test/raster.test.js`; appended cases in `test/mathsvg.test.js`. Host-integration
checks are manual, gated on S0 (full checklist in the workflow output / README).
