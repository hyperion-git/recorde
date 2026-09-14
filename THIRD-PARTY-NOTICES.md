# Third-party notices

MathJax Office is licensed under the Apache License 2.0 (see `LICENSE`). The
served bundle (`dist/`) additionally contains the following third-party
components, vendored from npm by `scripts/vendor-mathjax.mjs` into
`assets/vendor/mathjax/` so the add-in runs same-origin without a CDN.

## MathJax 4.1.2

- Package: `mathjax` — https://www.mathjax.org/
- License: Apache License 2.0 (`node_modules/mathjax/LICENSE`)
- Copyright: The MathJax Consortium
- Vendored: `tex-svg.js`, the `input/` TeX extensions, and the `sre/` speech-rule
  engine worker and math maps.

## MathJax font packages 4.1.2

The SVG font data (glyph path tables) served by the add-in comes from these
packages, each published by the MathJax project under the Apache License 2.0
(per each package's `package.json`). The glyph designs derive from the upstream
fonts listed, whose own licenses continue to apply to the glyph shapes.

| Package | Upstream font | Upstream license |
|---|---|---|
| `@mathjax/mathjax-tex-font` | MathJax TeX fonts (from Computer Modern / AMS fonts) | Apache License 2.0 |
| `@mathjax/mathjax-newcm-font` | New Computer Modern (Antonis Tsolomitis) | GUST Font License |
| `@mathjax/mathjax-termes-font` | TeX Gyre Termes (GUST e-foundry) | GUST Font License |
| `@mathjax/mathjax-pagella-font` | TeX Gyre Pagella (GUST e-foundry) | GUST Font License |
| `@mathjax/mathjax-stix2-font` | STIX Two (STI Pub Companies) | SIL Open Font License 1.1 |
| `@mathjax/mathjax-asana-font` | Asana Math (Apostolos Syropoulos) | SIL Open Font License 1.1 |

The font data is redistributed unmodified. The SIL Open Font License and the
GUST Font License both permit bundling and redistribution with software; neither
license's reserved font names are used as names of derived fonts here.

## Build and test tooling (not redistributed)

`office-addin-*` packages from Microsoft (MIT) are development dependencies only
and are not part of the served bundle.
