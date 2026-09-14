Recorde inserts LaTeX equations into Word as vector pictures and keeps the LaTeX in the document, so every equation stays click-to-edit. This is the first public release.

**Install:** download `manifest.xml` below and use *Home → Add-ins → More Add-ins → My Add-ins → Upload My Add-in* in Word (desktop or web). Organisations deploy the same file through the Microsoft 365 admin centre.

**Highlights**
- Live preview with inline error reporting, inline and display equations, click-to-edit, an equation list, keyboard shortcuts and an examples menu.
- Six math fonts — Termes (Times, `newtxmath`-like) by default — colour swatches, sizes matched to the surrounding text or fixed.
- Numbering: flush-right journal style, inline `(n)`, or native Word SEQ fields; renumber in document order; change style and convert existing equations.
- TeX layout rules: display spacing of one body size, no indent after a display, left-aligned (`fleqn`) or centred.
- Built-in macro preamble ported from a physics paper template (operators, delimiters, derivatives, vectors, siunitx subset, 174 named colours) plus per-document `\newcommand`s.
- Word on the web: raster insert at 300 DPI; desktop: vector SVG with baseline-aligned inline math.
- German UI, dark theme, fully self-hosted (no CDN), nothing leaves the document (see PRIVACY.md).
- `mjx-docx`: a headless CLI that puts the same equations into `.docx` files built by python-docx, docx-js or pandoc, with a skill for coding agents.

**Requirements:** Word for Windows/Mac (Microsoft 365) for vector insert and field numbering; Word on the web inserts raster pictures. See the README's host matrix.

**Files:** `manifest.xml` (the add-in), `Recorde-Showcase.docx` (a three-page test paper produced headlessly; open it in Word and click any equation).
