# Changelog

All notable changes to Recorde (formerly "MathJax Office"). Versions are the
add-in manifest versions; dates are commit dates.

## Unreleased — planned v2.0.0 (first release)
- Host capability line in Settings (Word version, insert mode, baseline shift,
  fields), a start-up notice when the baseline-shift API is missing, and a
  visible warning when the shift call fails (1.23.2).
- Public repository `hyperion-git/recorde` (squashed history); hosted build at
  `https://hyperion-git.github.io/recorde/`; support via GitHub issues.
- Requirements documented (README, support page, skill); `engines.node >= 20`.

## 1.23.1 — 2026-09-14
- Equation numbers vertically centred on the equation (table/field styles).

## 1.23.0 — 2026-09-14
- Numbering defaults to the flush-right table style; cell padding zeroed so
  the number sits exactly on the right text border.

## 1.22.3 — 2026-09-14
- Ribbon button and pane title say "Recorde" instead of "Equations".

## 1.22.2 — 2026-09-14
- Release preparation: README rewrite, CHANGELOG, PRIVACY, support page,
  release checklist, roadmap status.

## 1.22.1 — 2026-09-14
- Termes is the default font on every path (headless CLI, alt-text recovery),
  verified side by side against a newtxmath render.

## 1.22.0 — 2026-09-14
- Built-in preamble ports the paper template faithfully where MathJax 4 allows:
  sans-italic operators, rule-under-tensor, `\mathbbm`, phase-space optionals,
  `\xsize`/`\sym`, the template's local `\II`/`\EE`/`\pdg`, a siunitx subset.
  Acceptance test: every math snippet of the template's main document renders.

## 1.21.0 — 2026-09-14
- Display alignment: left (LaTeX `fleqn`, 25 pt math indent; default) or
  centred — per document in the pane, per equation in document scripts and the
  headless `{align=}` option / `--align`.

## 1.20.0 — 2026-09-14
- TeX display spacing: one body size above and below display equations, no
  first-line indent on the continuation text (both routes).

## 1.19.0 — 2026-09-14
- Showcase test paper: three pages, 72 equations, every font, numbering style,
  size, colour and the macro preamble; several inline equations per paragraph.
- Headless: per-equation placeholder options (`font`, `color`, `size`,
  `style`); column-aware numbering tables; the same paper built headlessly.

## 1.18.0 — 2026-09-14
- Headless toolset `mjx-docx` (process / render / list / update / renumber /
  check / preview) and the `recorde` agent skill. Placeholder grammar with
  labels and REF-field references; python-docx and pandoc fixtures.

## 1.17.0 — 2026-09-14
- Numbering style migration: switching the document style offers to convert
  existing equations; "Number this" works for table/field equations.

## 1.16.x — 2026-09-14
- RevTeX-style test paper; remaining template macros; 174 AFP colours as
  `\definecolor`.

## 1.15.0 — 2026-09-14
- Colour swatches from the AFP scientific-figure palette.

## 1.14.0 — 2026-09-14
- One-click test page from the Examples menu; German add-in name fix.

## 1.13.0 — 2026-09-14
- Tiered in-cell SVG placement for table/field numbering; physics examples.

## 1.12.x — 2026-09-14
- Examples picker, sticky actions, shortcuts, inline errors, collapsible
  settings, in-context preview, equation list; versioned pane URLs so a release
  can never load from cache.

## 1.11.x — 2026-09-14
- Recorde identity: name, mark, icons; theme fix.

## 1.10.0 — 2026-06-15
- Dark theme and Fluent 2 pane refresh.

## 1.8.0 – 1.9.0 — 2026-06-15
- Equation numbering: inline and flush-right table styles; native SEQ-field
  style (desktop); Renumber all.

## 1.7.x — 2026-06-14
- MathJax extensions and the default macro preamble from the paper template;
  Termes as default font.

## 1.5.0 – 1.6.1 — 2026-06-14
- Self-hosted MathJax and fonts (no CDN); German localisation; vendored speech
  worker so self-hosted startup completes.

## 1.2.0 – 1.4.0 — 2026-06-14
- Equation colour; symbol/template palette; per-document macro preamble.

## 1.1.x — 2026-06-14
- Word-web raster insert at 300 DPI; visible errors; settings persistence;
  cross-document alt-text recovery; deployable build + CI + GitHub Pages.
