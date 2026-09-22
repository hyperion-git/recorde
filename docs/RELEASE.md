# Release procedure

The first public release is planned as **v2.0.0** (roadmap M2). Until then the
1.x versions are internal builds. A release is two downloadable packages plus
the hosted build the Word package points to, so the gate is verification in
real Word.

## 1. Gate: verify in Word (desktop + web)

Work through `docs/VERIFY.md` on the current build. Blocking sections for the
first release, none of which have been run in Word since v1.17:

- §3 numbering styles incl. *Style migration (v1.17)*, *Display spacing
  (v1.20)*, *Display alignment (v1.21)*, *Showcase paper (v1.19)*.
- Headless round-trip (H0.1 of `docs/ROADMAP-headless.md`): open
  `examples/Recorde-Showcase.docx` in desktop Word — no repair prompt, equations
  render, clicking one loads it in the pane. Then unzip one add-in-written
  document into `test/fixtures/golden-addin/` and diff against
  `headless/lib/ooxml.mjs`.
- §4 dark theme, §5 pane features, §6 regression smoke; the same on Word web
  (raster path, field style disabled).
- The Word package's installers: `install-windows.cmd` on a Windows machine
  without the developer registration, `sh install-mac.command` on a Mac
  (both from an extracted `recorde-word-addin-<v>.zip` built by `npm run package`).
- The headless package: `npm install -g ./recorde-<v>.tgz` on a machine
  without the checkout, `mjx-docx --version`, `mjx-docx list Recorde-Showcase.docx`.

Fix what fails, ship as 1.x, re-verify.

## 2. Decisions to make before tagging

- **Name/URL.** Decided: public repository `hyperion-git/recorde`, Pages at
  `https://hyperion-git.github.io/recorde/`. The XML namespaces
  (`urn:mathjax-office:*`) and storage keys keep the old name on purpose —
  they identify existing documents' data and must never change.
- **Support URL.** The manifest's `SupportUrl` is rewritten to the Pages root
  at build time; `addin/index.html` there is the support page (install steps,
  privacy, contact).
- **Field-style numbering on the web** shows cached values only — acceptable
  for release (documented) or hide the option entirely?

## 3. What a release contains

`npm run package` (after `BASE_URL=https://hyperion-git.github.io/recorde npm run build`)
writes into `release/`:

| Asset | Contents |
|---|---|
| `recorde-word-addin-<v>.zip` | `manifest.xml` (released: Pages URLs, prod `<Id>`), `INSTALL.md` (= `docs/INSTALL-word.md`), `install-/uninstall-windows.cmd`, `install-/uninstall-mac.command`, `Recorde-Showcase.docx`, `PRIVACY.md`, licences |
| `recorde-mjx-docx-<v>.zip` | `recorde-<v>.tgz` (`npm pack`: `core/`, `headless/`, `skills/`), `INSTALL.md` (= `docs/INSTALL-mjx-docx.md`), `skill/recorde/`, `Recorde-Showcase.docx`, licences |
| `manifest.xml` | the released manifest alone, for *Upload My Add-in* |
| `SHA256SUMS` | checksums of the three files above |

The packager refuses a manifest that still contains `localhost` or the dev
`<Id>`, or whose `<Version>` differs from `package.json`; the Windows installers
must carry the manifest's `<Id>` (they register under it).

## 4. Cut the release

```sh
npm run bump -- 2.0.0
# CHANGELOG.md: move "Unreleased" to "2.0.0 — <date>"; docs/release-notes-2.0.0.md current
npm test
BASE_URL=https://hyperion-git.github.io/recorde npm run build && npm run validate:dist
npm run package                           # inspect release/ once by hand
git add -A && git commit -m "release: v2.0.0"
git push origin main                      # Pages deploy runs; wait for it — the manifest points there
gh run watch --exit-status $(gh run list --workflow "Deploy to GitHub Pages" --limit 1 --json databaseId -q '.[0].databaseId')
git tag -a v2.0.0 -m "Recorde v2.0.0" && git push origin v2.0.0
```

Pushing the tag runs `.github/workflows/release.yml`: it rebuilds, tests,
packages, creates the GitHub release **as a draft** (with
`docs/release-notes-<v>.md` as the notes when that file exists) if it does not
exist yet, and attaches the four assets (`--clobber`, so re-running replaces
them). An existing draft — v2.0.0 was drafted by hand — is reused. Then:

```sh
gh run watch --exit-status $(gh run list --workflow "Release packages" --limit 1 --json databaseId -q '.[0].databaseId')
gh release view v2.0.0                    # check assets + notes
gh release edit v2.0.0 --draft=false      # publish — the only manual, outward-facing step
```

Local fallback if the workflow is unavailable:
`gh release upload v2.0.0 release/*.zip release/manifest.xml release/SHA256SUMS --clobber`.

## 5. After tagging

- Copy `release/manifest.xml` to the local sideload folder (`deploy/`,
  untracked) and to the Windows sideload folder, or run the released
  `install-windows.cmd` — that is the smoke test of the released manifest in
  a fresh Word profile.
- Update `docs/ROADMAP.md` status.
