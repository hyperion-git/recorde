# In-Word verification package (current build: see `manifest.xml` here; sections are tagged with the version that introduced them)

`manifest.xml` here points at the live GitHub Pages build
(https://hyperion-git.github.io/recorde/), built from `main` on every push.

## 1. Load the add-in in Word for Windows

The developer machine uses a sideload registration for the add-in (registry
`HKCU\Software\Microsoft\Office\16.0\WEF\Developer`, value `9971ece4-…` →
`C:\Users\Public\MathJaxAddin\manifest.xml`); the build step copies the current
manifest into that folder.

1. Close Word completely (check the tray / Task Manager for WINWORD.EXE).
2. Start Word, open a blank document.
3. Home tab → **Recorde** button (group of the same name) opens the task pane.
4. Since v1.12.2 every release has versioned URLs (`taskpane.html?v=…`, and all
   JS imports), so a new manifest always loads the matching pane. If something
   still looks stale, close Word and delete the contents of
   `%LOCALAPPDATA%\Microsoft\Office\16.0\Wef\` (last resort).

Fresh machine or registration lost: double-click `sideload.reg` (adds the same
registry value), then steps 1–3. `unsideload.reg` removes it.

DevTools: right-click inside the pane → **Inspect** (Edge WebView2). The Console
tab shows any boot error; the pane also surfaces errors in its status line.

## 2. The reported numbering bug

Report: "All but the latex label for numbering is bugged." Please capture the
exact symptom:

- [ ] Pane open, mode = **Display**. What does the *Numbering* section show?
      (style dropdown text, checkbox label, enabled/disabled state)
- [ ] Any red text in the status line? Any Console errors?
- [ ] Screenshot of the pane (light theme) — the section between *Size* and the
      *Insert / New / Renumber* buttons.
- [ ] Which Office display language is active (File → Options → Language)?

## 3. Numbering styles

For each style: pane → *Numbering* dropdown → style; mode Display; tick
**Number this equation**; type `E = mc^2`; Insert. Repeat with `a^2+b^2=c^2`.

Inline style
- [ ] Preview shows `E = mc^2  (1)`; second equation previews `(2)`.
- [ ] Both inserted, numbers `(1)` `(2)` trail the equation on the same line.
- [ ] Click equation 1 → pane loads it as "(1)", editing keeps the number;
      **Update** re-inserts with the same number.

Flush right (table)
- [ ] Equation centred, `(n)` flush right in a borderless 1×2 table.
- [ ] Click-to-edit works inside the table cell.

Word field (desktop)
- [ ] Option is enabled in the dropdown (desktop only).
- [ ] Inserted number is a SEQ field: Alt+F9 toggles to `{ SEQ equation \* ARABIC }`.
- [ ] Cut equation 2's paragraph, paste it above equation 1 → **Renumber all**
      → numbers follow document order (also Ctrl+A, F9 should agree).

Renumber all (mixed)
- [ ] With inline + table equations in one document, delete the first one →
      **Renumber all** → remaining equations renumber from (1) in document order.

Showcase paper (v1.19)
- [ ] Examples → Testing → "Test paper (RevTeX-style, two columns)…" in a NEW document: ~72 equations,
      takes a minute or two; final notice "… 0 failed". Two-column body from Sec. I on.
- [ ] Sec. IV.A: one inline-style number "(13)" inside the picture; one SEQ-field equation (Alt+F9).
- [ ] Sec. IV.B: six fonts visibly differ; IV.C: 8/10/12/14 pt inline sizes, blue equation, AFP colours.
- [ ] Compare with deploy/Recorde-Showcase.docx (headless twin of the same source, also in
      C:\Users\Public\MathJaxAddin): opens without repair prompt; click any equation → pane loads it;
      numbers flush right inside the columns; report layout differences.

Host line + inline baseline (v1.23.2)
- [ ] Settings panel, last line: "Host: Word <version> (PC) · insert: vector SVG · baseline shift: yes/no ·
      fields: yes/no". REPORT THIS LINE. If "baseline shift: no", an info notice explains why (Word 2507+ needed).
- [ ] With "baseline shift: yes": insert `dog` (or `\int_0^1 y\,dy`) inline in a line of 9–12 pt text,
      zoom 300 %+: the bottoms of d and o sit ON the text baseline (v1.23.3 adds Word's bottom
      effectExtent — 0.75 pt on the 2026-09-15 sample — to the shift; before, the math sat that much high).
      Save as C:\Users\Public\MathJaxAddin\inline-test.docx: the run should carry w:position ≈
      -2·(descent + 0.75) half-points. Red warning "lowering … failed" → report its text.

Number position (v1.23)
- [ ] New document: Settings → Numbering shows *Flush right (table)* by default.
- [ ] A numbered equation's "(n)" ends exactly on the right text border (compare with a right-aligned
      paragraph above it); the equation starts 25 pt from the left border (fleqn) — no extra cell padding.
- [ ] Same in a two-column section (number on the column's right edge).
- [ ] For a tall (multi-line `aligned`) equation the "(n)" is vertically centred on the equation, not at
      its top (v1.23.1).

Display alignment (v1.21)
- [ ] Settings → Display shows "Left (indented, TeX fleqn)" selected by default in a new document; the
      summary line reads "… · Left".
- [ ] A new display equation sits flush left, indented 25 pt, with the number (table style) still flush
      right; switch to "Centred" → the next equation is centred. The choice persists with the document.
- [ ] Test paper Sec. IV.D: one left and one centred equation regardless of the setting.

Display spacing (v1.20, TeX skips)
- [ ] A new display equation (any style) has ≈ one body-size gap above and below it (10 pt in the test
      paper, 11 pt in a Normal document); the numbering table's row is taller by the same amount.
- [ ] The paragraph typed AFTER a display equation does not inherit the gap (spacing before = 0).
- [ ] Unwrapping a table equation ("Number this" off) keeps the gap on the bare paragraph.

Style migration (v1.17)
- [ ] With inline-numbered equations present, switch *Numbering* to *Flush right
      (table)* → notice "N numbered equation(s) still use a different numbering
      style" with a **Convert to "Flush right (table)"** button.
- [ ] Click it → each inline equation becomes a borderless 1×2 table (equation
      centred, `(n)` flush right), no empty paragraph left above the table,
      numbers 1..n in document order; unnumbered equations untouched.
- [ ] Switch back to *Inline* → convert → tables disappear, `(n)` trails each
      equation on its own centred line, no leftover table rows.
- [ ] Desktop: switch *Table* → *Word field* → convert → number cells become SEQ
      fields (Alt+F9), picture unchanged.
- [ ] Click a table equation → **Number this equation** is enabled → untick →
      Update → equation leaves its table (bare, centred, no number).
- [ ] Click an unnumbered display equation in a table-style document → tick
      **Number this equation** → Update → it is wrapped into a numbering table
      with the next number.
- [ ] Undo (Ctrl+Z) after a conversion: note how many steps it takes; report.

## 4. Dark theme

- [ ] File → Account → Office Theme → **Dark Gray** → pane switches to dark;
      the equation preview stays paper-white with a dark equation.
- [ ] **Black** → same.
- [ ] Back to **Colorful** → pane returns to light. Check text contrast of the
      status line, dropdowns and disabled buttons in both themes.

## 5. Pane features (v1.12.0)

- [ ] **Examples** dropdown (top right of *LaTeX source*): pick "Maxwell’s
      equations" → source + Display mode load, preview renders; "Schrödinger
      equation" and "Preamble macros" also tick *Number this equation*.
- [ ] **Sticky buttons**: with Settings open and the pane short, Insert / New /
      Renumber stay pinned at the bottom while the rest scrolls.
- [ ] **Ctrl+Enter** inserts (or updates when editing); **Esc** in the editor
      while editing an existing equation returns to "New equation" (Esc on a
      fresh draft does nothing).
- [ ] **Syntax errors**: type `\frac{a}{` → red "Missing close brace" under the
      source and a red editor border; fixing it clears both.
- [ ] **Settings** section collapsed by default, summary reads "Termes · Match
      selection" with a colour swatch; open state is remembered after reopening
      the pane.
- [ ] **Inline preview in context**: Inline mode shows "where … holds for all x."
      around the equation, Display mode shows the equation alone.
- [ ] **Equations in this document**: open it → numbered list in document order
      with "(n)" or "–" and the source; click an entry → Word selects that
      equation and the pane loads it (entry highlighted). Count in the header.

## 6. Regression smoke

- [ ] Inline equation (mode Inline) inserts and baseline-aligns with text.
- [ ] Clicking an existing inline equation reloads it into the pane.
- [ ] Symbols palette inserts a snippet at the caret.
- [ ] Macro preamble: `\vb{x}` (or another template macro) renders.

Report results per checkbox; for anything failing, a screenshot + the Console
output is the most useful evidence.
