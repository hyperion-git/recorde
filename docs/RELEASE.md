# Release procedure

The first public release is planned as **v2.0.0** (roadmap M2). Until then the
1.x versions are internal builds. A release is a manifest people install plus
the hosted build it points to, so the gate is verification in real Word.

## 1. Gate: verify in Word (desktop + web)

Work through `deploy/VERIFY.md` on the current build. Blocking sections for the
first release, none of which have been run in Word since v1.17:

- §3 numbering styles incl. *Style migration (v1.17)*, *Display spacing
  (v1.20)*, *Display alignment (v1.21)*, *Showcase paper (v1.19)*.
- Headless round-trip (H0.1 of `headless/ROADMAP.md`): open
  `deploy/Recorde-Showcase.docx` in desktop Word — no repair prompt, equations
  render, clicking one loads it in the pane. Then unzip one add-in-written
  document into `test/fixtures/golden-addin/` and diff against
  `headless/lib/ooxml.mjs`.
- §4 dark theme, §5 pane features, §6 regression smoke; the same on Word web
  (raster path, field style disabled).

Fix what fails, ship as 1.x, re-verify.

## 2. Decisions to make before tagging

- **Name/URL.** Decided: public repository `hyperion-git/recorde`, Pages at
  `https://hyperion-git.github.io/recorde/`. The XML namespaces
  (`urn:mathjax-office:*`) and storage keys keep the old name on purpose —
  they identify existing documents' data and must never change.
- **Support URL.** The manifest's `SupportUrl` is rewritten to the Pages root
  at build time; `index.html` there is the support page (install steps,
  privacy, contact).
- **Field-style numbering on the web** shows cached values only — acceptable
  for release (documented) or hide the option entirely?

## 3. Cut the release

```sh
npm run bump -- 2.0.0
BASE_URL=https://hyperion-git.github.io/recorde npm run build && npm run validate:dist
npm test
# CHANGELOG.md: move "Unreleased" to "2.0.0 — <date>"
git add -A && git commit -m "release: v2.0.0"
git push origin main                      # Pages deploy runs; wait for it
gh run watch --exit-status $(gh run list --workflow "Deploy to GitHub Pages" --limit 1 --json databaseId -q '.[0].databaseId')
git tag -a v2.0.0 -m "Recorde v2.0.0" && git push origin v2.0.0
gh release create v2.0.0 --title "Recorde v2.0.0" --notes-file docs/release-notes-2.0.0.md \
  dist/manifest.xml deploy/Recorde-Showcase.docx
```

Attach `dist/manifest.xml` (the installable file) and the showcase document.
A draft release can be created earlier with `--draft` to review the notes.

## 4. After tagging

- Copy `dist/manifest.xml` to `deploy/` and to the Windows sideload folder.
- Smoke-test the *released* manifest in a fresh Word profile (Upload My Add-in).
- Update `ROADMAP.md` status and `~/.claude/state/active/mathjax-office.md`.
