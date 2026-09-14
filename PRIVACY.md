# Privacy

Recorde runs entirely inside the Word task pane and the document.

- **No data leaves the document.** LaTeX source, rendered pictures, numbering
  state and the macro preamble are stored in the `.docx` itself (custom XML
  parts and picture alt text). Nothing is sent to any server; there is no
  telemetry, no account, no analytics.
- **What is loaded from the network:** only the add-in's own static files
  (pane HTML/JS, MathJax and font data) from the host the manifest points to
  (`https://hyperion-git.github.io/recorde/` for the published build).
  Word fetches these the same way it loads any web add-in; no request carries
  document content.
- **Local storage:** the pane remembers your last font, size, colour and
  alignment choices in the document's settings (travels with the file) and in
  the browser storage of the Word host (per machine). Both can be cleared by
  Word's add-in cache reset.
- **Headless tool:** `mjx-docx` runs on your machine, reads and writes the
  files you name, and makes no network requests.
