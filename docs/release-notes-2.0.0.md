Recorde inserts LaTeX equations into Word as vector pictures and keeps the LaTeX in the document, so every equation stays click-to-edit. This is the first public release.

**Two packages** — each with its own `INSTALL.md`:

- `recorde-word-addin-2.0.0.zip` — the Word add-in. Extract, run `install-windows.cmd` or `sh install-mac.command`, restart Word; or upload the included `manifest.xml` via *Home → Add-ins → More Add-ins → My Add-ins → Upload My Add-in* (desktop or web). Organisations deploy the same file through the Microsoft 365 admin centre. `manifest.xml` is also attached on its own.
- `recorde-mjx-docx-2.0.0.zip` — `mjx-docx`, the headless command-line tool that puts the same equations into `.docx` files built by python-docx, docx-js or pandoc (Node.js 20+: `npm install -g ./recorde-2.0.0.tgz`), with a skill for coding agents (`mjx-docx skill install`).

**Highlights**
- Live preview with inline error reporting, inline and display equations, click-to-edit, an equation list, keyboard shortcuts and an examples menu.
- Six math fonts — Termes (Times, `newtxmath`-like) by default — colour swatches, sizes matched to the surrounding text or fixed.
- Numbering: flush-right journal style, inline `(n)`, or native Word SEQ fields; renumber in document order; change style and convert existing equations.
- TeX layout rules: display spacing of one body size, no indent after a display, left-aligned (`fleqn`) or centred.
- Built-in macro preamble ported from a physics paper template (operators, delimiters, derivatives, vectors, siunitx subset, 174 named colours) plus per-document `\newcommand`s.
- Word on the web: raster insert at 300 DPI; desktop: vector SVG with baseline-aligned inline math.
- German UI, dark theme, fully self-hosted (no CDN), nothing leaves the document (see PRIVACY.md).

**Requirements:** Word for Windows/Mac (Microsoft 365) for vector insert and field numbering; Word on the web inserts raster pictures. See `INSTALL.md` in the Word package for the host matrix.

**Check it:** `Recorde-Showcase.docx` (in both zips) is a three-page test paper produced headlessly; open it in Word and click any equation.
