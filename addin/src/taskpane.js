import {
  newUuid, writeXmlPart, readXmlPartByUuid,
  embedSourceInSvg, ALT_PREFIX, readPreamble, writePreamble,
  readNumberingState, writeNumberingState, readAllEquations,
} from './storage.js';
import {
  decorateLatexForNumber, normalizeStyle, nextNumberFor, assignNumbers,
  numberPlacement, formatNumber, DEFAULT_NUMBERING, migrationOp, planStyleMigration,
} from './numbering.js';
import {
  mathJaxFontName, flattenLatex, computeRenderSize,
  scaleSvgToPt, inlineColors, descentPtFromViewBox, isParagraphTextEmpty,
  findRenderError,
} from './mathsvg.js';
import { describeInsertError, pictureVanishedNotice } from './errors.js';
import { parseSettings, serializeSettings, DEFAULT_SETTINGS } from './settings.js';
import { looksLikeEquation, buildDegradedPayload } from './recovery.js';
import { computeRasterSize, stripDataUri, isPayloadTooLarge } from './raster.js';
import { applySnippet, PALETTE } from './snippets.js';
import { EXAMPLES, findExample } from './examples.js';
import { TEST_PAGE } from './testpage.js';
import { TEST_PAPER } from './testpaper.js';
import { findColorSet, normalizeHex } from './colors.js';
import { strings as getStrings } from './i18n.js';
import { DEFAULT_PREAMBLE } from './preamble.js';

// Host capability state, resolved once in Office.onReady.
//   insertMode — how (or whether) this host can insert an equation picture:
//     'svg'  → ImageCoercion 1.2 (Word desktop): native SVG insert.
//     'png'  → ImageCoercion 1.1 (incl. Word web): raster fallback (WP1.5).
//     'none' → no picture insert here; Insert is disabled with a notice.
//   canShiftBaseline — WordApiDesktop 1.3 (desktop, Word 2507+): lets us lower
//     an inline equation's run so its math baseline meets the text baseline.
// See README "Supported hosts".
let insertMode = 'none';        // 'svg' | 'png' | 'none'
let canShiftBaseline = false;
// canInsertField — WordApi 1.5 + a desktop signal (SVG insert ⇒ Word desktop).
// Field insertion is desktop-only: Word web satisfies the 1.5 version check yet
// throws InvalidArgument on insertField, so we also require the desktop path and
// still wrap the call in try/catch. Gates the 'field' numbering style. (WP2.2-C)
let canInsertField = false;

let mathJaxReady = false;
let loadedFont = null;          // short name of the font MathJax is loaded with
let loadingFont = null;         // short name of the font a load is in flight for
let mathJaxLoadPromise = null;  // in-flight (re)load, shared by concurrent callers
let mathJaxLoadToken = 0;       // generation guard so a superseded load can't win
let editingUuid = null;
let editingNumber = 0;          // the loaded equation's number (0 = unnumbered/new)
let editingStyle = 'inline';    // the loaded equation's numbering style
let selectionCheckPending = false;
let suppressSelection = false;  // true while Renumber-all drives the selection
// Per-document numbering state (style + next-number counter), loaded from the
// doc's XML part in Office.onReady; defaults until then. (WP2.2)
let numberingState = { ...DEFAULT_NUMBERING };
let STR = getStrings('en');     // active UI string table (WP3.4); set in Office.onReady

// Font display names, parallel to the <select id="font"> options. Used for the
// "Loading … font" hint and nowhere load-bearing.
const FONT_LABELS = {
  tex: 'Computer Modern', newcm: 'New Computer Modern', termes: 'Termes',
  stix2: 'STIX Two', pagella: 'Pagella', asana: 'Asana Math',
};

const $ = (id) => document.getElementById(id);

// Width (pt) of the narrow right-hand number cell in the flush-right table style.
// The wide equation cell takes the remaining column width. (WP2.2 table style)
const NUMBER_CELL_WIDTH_PT = 54;   // ≈ 0.75in, fits "(99)"

// TeX's display spacing: LaTeX's \abovedisplayskip = \belowdisplayskip is one
// body font size (10 pt at 10 pt, 11 at 11, 12 at 12). Applied as paragraph
// spacing before/after a display equation (or its numbering table's cell
// paragraphs), so displays sit clear of the prose the way they do on a TeX
// page instead of flush against it. Follows the BODY size, not the equation's.
function displaySkipPt(bodyPt) {
  return Number.isFinite(bodyPt) && bodyPt > 0 ? bodyPt : 11;
}
// Host capability line (Settings panel) + a one-time notice when the baseline
// shift is unavailable. Font.position is WordApiDesktop 1.3 (Word 2507 / Mac
// 16.99, Microsoft 365 only) — without it inline math sits on its box bottom,
// slightly high, and nothing else in the pane would reveal why.
function hostVersion() {
  const d = Office.context.diagnostics || {};
  return { app: d.host || 'Word', version: d.version || '?', platform: d.platform || '?' };
}
function showHostInfo() {
  const el = $('host-info');
  if (!el) return;
  const { app, version, platform } = hostVersion();
  const yn = (b) => (b ? STR.hostYes : STR.hostNo);
  const insert = insertMode === 'svg' ? STR.insertSvg : insertMode === 'png' ? STR.insertPng : STR.insertNone;
  el.textContent = fmt(STR.hostLine, { app, version, platform, insert, shift: yn(canShiftBaseline), fields: yn(canInsertField) });
  if (insertMode === 'svg' && !canShiftBaseline) notify('info', fmt(STR.noBaselineShift, { version }));
}

// Display alignment (Settings → Display): 'left' = LaTeX fleqn — flush left,
// indented by \mathindent = 25 pt; 'center' = plain LaTeX. The document
// scripts may override per equation (req.align).
const MATH_INDENT_PT = 25;
function displayAlign() {
  return $('display-align').value === 'center' ? 'center' : 'left';
}
// Apply the alignment to a display-equation paragraph proxy (queued, no sync).
function alignDisplayParagraph(p, align) {
  if (align === 'center') { p.alignment = Word.Alignment.centered; p.leftIndent = 0; }
  else { p.alignment = Word.Alignment.left; p.leftIndent = MATH_INDENT_PT; }
}
// Body size for the skip when an operation has no insert context (migrations).
async function resolveDisplaySkipPt() {
  try { const { bodyPt } = await resolveSizes(); return displaySkipPt(bodyPt); } catch { return 11; }
}
const TOO_LARGE_MSG =
  'This equation is too large to insert as an image on the web. Try a ' +
  'smaller size, or insert it on Word desktop.';

Office.onReady(({ host }) => {
  if (host !== Office.HostType.Word) {
    $('preview').textContent = 'This add-in only supports Word in v1.';
    return;
  }
  applyTheme();   // match Word's light/dark theme before first paint
  // Re-evaluate if the OS preference flips while the pane is open (the Office
  // theme itself has no change event we can rely on across hosts).
  try { window.matchMedia('(prefers-color-scheme: dark)').addEventListener('change', applyTheme); } catch { /* no matchMedia */ }
  // Resolve host capabilities once. isSetSupported is synchronous.
  const req = Office.context.requirements;
  insertMode = detectInsertMode(req);
  canShiftBaseline = req.isSetSupported('WordApiDesktop', '1.3');
  canInsertField = insertMode === 'svg' && req.isSetSupported('WordApi', '1.5');

  installGlobalErrorHandlers();   // first, so any boot-time failure is visible
  STR = getStrings(Office.context.displayLanguage || globalThis.navigator?.language);
  applyI18n();
  showHostInfo();                 // what this build can do — the first thing to read in a bug report
  wireUi();
  applySettingsToUi(loadSavedSettings());   // restore defaults BEFORE bootMathJax
  bootMathJax();                            // so MathJax boots with the saved font
  registerSelectionHandler();
  loadPreamble();                           // restore this doc's macro preamble
  loadNumberingState();                     // restore this doc's numbering counter
  syncStyleOptions();                       // enable the field style only on capable hosts
  updateNumberingUi();
  updateUiState();

  if (insertMode === 'none') {
    notify('warn',
      'This host can’t insert equation pictures. You can preview here, but ' +
      'Insert is disabled. Word on Windows or Mac (desktop) is required; a ' +
      'web-compatible insert is on the roadmap.');
  } else if (insertMode === 'png') {
    notify('info',
      'Inserting as a high-resolution image (vector SVG insert needs Word ' +
      'desktop). Equations stay editable; inline baseline alignment is ' +
      'desktop-only.');
  }
});

// User-facing notice surfacing host-capability and insert/storage failures that
// would otherwise be invisible. One funnel, three levels ('info'|'warn'|'error');
// CSS for each lives in taskpane.html. notify('info','') (or clearNotice) hides it.
let noticeFromAppFlow = false;   // true while a real app message is shown, so the
                                 // uncaught-error handler (WP1.1) won't clobber it
function notify(level, msg) {
  const el = $('notice');
  if (!el) return;
  el.textContent = msg || '';
  el.className = msg ? `notice notice--${level}` : 'notice';
  el.hidden = !msg;
  noticeFromAppFlow = !!msg;
}
// A notice carrying one action button (e.g. "Convert 3 equations"). The button
// is part of the notice, so any later notify()/clearNotice() removes it too —
// a stale offer can never outlive the situation that produced it.
function notifyAction(level, msg, label, onClick) {
  notify(level, msg);
  const el = $('notice');
  if (!el || !msg) return;
  const btn = document.createElement('button');
  btn.type = 'button';
  btn.className = 'notice__action';
  btn.textContent = label;
  btn.addEventListener('click', () => { clearNotice(); onClick(); });
  el.append(' ', btn);
}
function clearNotice() { notify('info', ''); }
function notifyResult(result) { if (result) notify(result.level, result.msg); }

// Map host capabilities to an insert strategy. ImageCoercion 1.2 (Word desktop)
// → native SVG; 1.1 (incl. Word web) → raster PNG (enabled in WP1.5); else none.
function detectInsertMode(req) {
  if (req.isSetSupported('ImageCoercion', '1.2')) return 'svg';
  if (req.isSetSupported('ImageCoercion', '1.1')) return 'png';   // Word web: raster
  return 'none';
}

// Sync the pane to Word's theme: set <html data-theme> from the live Office theme
// (Colorful / Dark Gray / Black), falling back to the OS prefers-color-scheme.
// taskpane.html defines the dark token set under :root[data-theme="dark"]; the
// preview stays "paper"-light in both themes so equations never go invisible.
function applyTheme() {
  let dark = false;
  try {
    const bg = Office.context.officeTheme && Office.context.officeTheme.bodyBackgroundColor;
    if (bg) dark = isDarkColor(bg);
    else if (window.matchMedia) dark = window.matchMedia('(prefers-color-scheme: dark)').matches;
  } catch {
    if (window.matchMedia) dark = window.matchMedia('(prefers-color-scheme: dark)').matches;
  }
  document.documentElement.dataset.theme = dark ? 'dark' : 'light';
}

// Perceived-luminance test for a #rgb / #rrggbb color string.
function isDarkColor(hex) {
  const m = /^#?([0-9a-f]{3}|[0-9a-f]{6})$/i.exec(String(hex).trim());
  if (!m) return false;
  let h = m[1];
  if (h.length === 3) h = h[0] + h[0] + h[1] + h[1] + h[2] + h[2];
  const n = parseInt(h, 16);
  const r = (n >> 16) & 255, g = (n >> 8) & 255, b = n & 255;
  return (0.299 * r + 0.587 * g + 0.114 * b) < 128;
}

// Surface otherwise-invisible uncaught failures (async rejections, event-handler
// throws, late MathJax errors) in the notice region — so problems are visible
// without DevTools. Never overwrites a specific app message already shown.
function notifyUncaught(detail) {
  if (noticeFromAppFlow) return;
  notify('error', `Unexpected error: ${detail}. If it persists, reload the task pane.`);
}
function installGlobalErrorHandlers() {
  window.addEventListener('error', (e) =>
    notifyUncaught(e.message || e.error?.message || 'script error'));
  window.addEventListener('unhandledrejection', (e) =>
    notifyUncaught(e.reason?.message || e.reason?.code || String(e.reason) || 'promise rejection'));
}

function wireUi() {
  $('latex').addEventListener('input', () => { updateLineNumbers(); renderPreview(); });
  $('latex').addEventListener('scroll', () => { $('latex-gutter').scrollTop = $('latex').scrollTop; });
  $('mode-inline').addEventListener('change', onModeChange);
  $('mode-display').addEventListener('change', onModeChange);
  $('number-this').addEventListener('change', renderPreview);
  $('display-align').addEventListener('change', () => { saveCurrentSettings(); updateSettingsSummary(); });
  $('number-style').addEventListener('change', onNumberStyleChange);
  $('renumber').addEventListener('click', () => onRenumberAll().finally(refreshEquationListIfOpen));
  $('font').addEventListener('change', onFontChange);
  $('size-mode').addEventListener('change', () => {
    $('size-pt').disabled = $('size-mode').value !== 'fixed';
  });
  $('color').addEventListener('input', renderPreview);
  $('preamble').addEventListener('input', onPreambleInput);
  $('insert').addEventListener('click', () => onPrimaryAction().finally(refreshEquationListIfOpen));
  $('new-eq').addEventListener('click', onNewEquation);
  // Examples picker (examples.js) — a native <select> so it's keyboard- and
  // theme-correct for free; loading one starts a new equation.
  buildExamples();
  $('examples').addEventListener('change', onExampleChange);
  // Keyboard: Ctrl/⌘+Enter inserts or updates from anywhere in the pane; Esc in
  // the editor drops an in-progress EDIT (never a fresh draft — that would lose text).
  document.addEventListener('keydown', onGlobalKeydown);
  $('latex').addEventListener('keydown', onEditorKeydown);
  // Collapsed settings carry a one-line summary of font · size · colour.
  for (const id of ['font', 'size-mode', 'size-pt', 'color']) {
    $(id).addEventListener('change', updateSettingsSummary);
  }
  restoreUiState();
  // Prefab colour swatches (colors.js): pick a set, click a swatch → colour input.
  $('color-set').addEventListener('change', () => { saveUiState(); renderSwatches(); });
  $('swatches').addEventListener('click', onSwatchClick);
  $('color').addEventListener('input', markActiveSwatch);
  renderSwatches();
  // Equation list: scanned only while its section is open (one document walk).
  $('eq-list-details').addEventListener('toggle', refreshEquationListIfOpen);
  $('eq-list').addEventListener('click', onEquationListClick);
  buildPalette();
  // Delegated: one listener for every palette button (they carry their snippet).
  $('palette').addEventListener('click', (e) => {
    const btn = e.target.closest('.palette__btn');
    if (btn) insertSnippet(btn.dataset.snippet);
  });
  updateLineNumbers();
}

// Render the symbol palette from PALETTE (snippets.js) into #palette.
function buildPalette() {
  const root = $('palette');
  if (!root) return;
  for (const cat of PALETTE) {
    const head = document.createElement('div');
    head.className = 'palette__cat';
    head.textContent = cat.name;
    root.appendChild(head);
    const row = document.createElement('div');
    row.className = 'palette__row';
    for (const item of cat.items) {
      const b = document.createElement('button');
      b.type = 'button';
      b.className = 'palette__btn';
      b.textContent = item.label;
      b.title = item.title || item.snippet;
      b.dataset.snippet = item.snippet;
      row.appendChild(b);
    }
    root.appendChild(row);
  }
}

// Insert a palette snippet at the caret (wrapping any selection), then keep
// focus in the editor and refresh the gutter + preview.
function insertSnippet(snippet) {
  const ta = $('latex');
  const { value, caret } = applySnippet(ta.value, ta.selectionStart, ta.selectionEnd, snippet);
  ta.value = value;
  ta.setSelectionRange(caret, caret);
  ta.focus();
  updateLineNumbers();
  renderPreview();
}

// Render the line-number gutter. Counts logical lines (\n-separated); a
// soft-wrapped long line still counts as one — matching the model the user
// sees in their LaTeX source. If they need exact visual-row numbering for
// wrapped content, the workaround is to keep lines short.
function updateLineNumbers() {
  const lineCount = Math.max(1, $('latex').value.split('\n').length);
  let nums = '';
  for (let i = 1; i <= lineCount; i++) nums += i + '\n';
  $('latex-gutter').textContent = nums;
}

// ---- Examples picker (examples.js) ----
const TEST_PAGE_OPTION = '__testpage';
const TEST_PAPER_OPTION = '__testpaper';
function buildExamples() {
  const sel = $('examples');
  for (const ex of EXAMPLES) {
    const o = document.createElement('option');
    o.value = ex.id;
    o.textContent = ex.name;
    sel.appendChild(o);
  }
  // One-click test page (testpage.js) in its own group at the end.
  const grp = document.createElement('optgroup');
  grp.label = STR.testPageGroup;
  for (const [value, label] of [[TEST_PAGE_OPTION, STR.testPage], [TEST_PAPER_OPTION, STR.testPaper]]) {
    const o = document.createElement('option');
    o.value = value;
    o.textContent = label;
    grp.appendChild(o);
  }
  sel.appendChild(grp);
}

// Load an example as a NEW equation: source + mode, and tick "Number this
// equation" when the example asks for it (display mode only — updateNumberingUi
// keeps the box disabled otherwise). The select snaps back to its placeholder.
function onExampleChange() {
  const sel = $('examples');
  const choice = sel.value;
  sel.value = '';
  if (choice === TEST_PAGE_OPTION) { runDocumentScript(TEST_PAGE); return; }
  if (choice === TEST_PAPER_OPTION) { runDocumentScript(TEST_PAPER, { font: 'termes', sizeMode: 'fixed', fixedPt: 10, style: 'table' }); return; }
  const ex = findExample(choice);
  if (!ex) return;
  onNewEquation();
  $('latex').value = ex.latex;
  updateLineNumbers();
  if (ex.mode === 'display') $('mode-display').checked = true;
  else                       $('mode-inline').checked = true;
  updateNumberingUi();
  if (ex.numbered && !$('number-this').disabled) $('number-this').checked = true;
  renderPreview();
  $('latex').focus();
}

// ---- Document scripts: test page (testpage.js) and test paper (testpaper.js) ----
// Appends a scripted document to the END of the current one: headings/prose via
// the Word API (with per-block typography from `fmt`), equations via
// performInsert() with explicit settings, optional OOXML blocks (section
// breaks). Serialised, one block at a time; failures are counted, not fatal.
// The pane's own font dropdown is restored afterwards (performInsert switches
// MathJax fonts as it goes).
let scriptRunning = false;
async function runDocumentScript(blocks, eqDefaults = {}) {
  if (scriptRunning) return;
  if (!mathJaxReady || insertMode === 'none') { notify('warn', STR.testPageUnavailable); return; }
  scriptRunning = true;
  const sel = $('examples');
  sel.disabled = true;
  clearNotice();
  const total = blocks.reduce((n, b) => n + (b.eq ? 1 : 0) + (b.parts ? b.parts.filter((x) => typeof x !== 'string').length : 0), 0);
  let done = 0, failed = 0, layoutFailed = false;
  const DEFAULTS = { font: 'termes', color: '#000000', displayMode: 'display', sizeMode: 'body',
                     fixedPt: 11, numbered: false, style: numberingState.style, align: displayAlign(),
                     editingUuid: null, editingNumber: 0, ...eqDefaults };
  try {
    for (const b of blocks) {
      if (b.ooxml) {
        try { await appendOoxml(b.ooxml); } catch (e) { layoutFailed = true; console.warn('OOXML block failed:', e.message); }
        continue;
      }
      await appendScriptParagraph(b);
      // A paragraph is `text` | `before`+`eq`+`after` | `eq` | `parts` (an
      // alternating list of strings and { latex, … } equations — several inline
      // equations in one paragraph, as real prose has).
      const parts = b.parts ? b.parts.slice(typeof b.parts[0] === 'string' ? 1 : 0)
        : b.eq ? [b.eq, ...(b.after ? [b.after] : [])] : [];
      for (const part of parts) {
        if (typeof part === 'string') { await appendTextAfterSelection(part, b.fmt); continue; }
        done++;
        $('status').textContent = STR.testPageProgress.replace('{n}', done).replace('{total}', total);
        let r = null;
        try { r = await performInsert({ ...DEFAULTS, ...part }); } catch { r = null; }
        if (!r) failed++;
        else if (b.fmt) await applyFormatToEquationTable(b.fmt);
      }
    }
  } finally {
    scriptRunning = false;
    sel.disabled = false;
    try { await ensureMathJax($('font').value); } catch { /* preview font may differ until next change */ }
    updateUiState();
    refreshEquationListIfOpen();
  }
  notify(failed || layoutFailed ? 'warn' : 'info',
    STR.testPageDone.replace('{ok}', done - failed).replace('{failed}', failed)
    + (layoutFailed ? ' ' + STR.testPaperLayoutFailed : ''));
}

// New paragraph at the end of the body with the block's text, built-in style
// and typography; leaves the selection collapsed at its end so the next insert
// lands there.
async function appendScriptParagraph(b) {
  await Word.run(async (context) => {
    const text = b.h1 || b.h2 || b.p || b.before || b.text || (b.parts && typeof b.parts[0] === 'string' ? b.parts[0] : '');
    const para = context.document.body.insertParagraph(text, Word.InsertLocation.end);
    para.styleBuiltIn = b.h1 ? Word.Style.heading1 : b.h2 ? Word.Style.heading2 : Word.Style.normal;
    para.alignment = Word.Alignment.left;   // a new paragraph may inherit "centered" from a display equation
    para.leftIndent = 0;                     // …or the fleqn math indent
    applyParagraphFormat(para, b.fmt);
    para.getRange(Word.RangeLocation.end).select();
    await context.sync();
  });
}

const ALIGN = { left: 'Left', centered: 'Centered', right: 'Right', justified: 'Justified' };
function applyParagraphFormat(para, fmt) {
  if (!fmt) return;
  if (fmt.font) para.font.name = fmt.font;
  if (fmt.size) para.font.size = fmt.size;
  para.font.bold = !!fmt.bold;
  para.font.italic = !!fmt.italic;
  if (fmt.align) para.alignment = ALIGN[fmt.align] || 'Left';
  if (fmt.firstLineIndent != null) para.firstLineIndent = fmt.firstLineIndent;
  if (fmt.leftIndent != null) para.leftIndent = fmt.leftIndent;
  if (fmt.rightIndent != null) para.rightIndent = fmt.rightIndent;
  // A paragraph appended after a display equation inherits its display skips;
  // an explicit fmt always resets them (0 unless the fmt says otherwise).
  para.spaceBefore = fmt.spaceBefore != null ? fmt.spaceBefore : 0;
  para.spaceAfter = fmt.spaceAfter != null ? fmt.spaceAfter : 0;
  if (fmt.lineSpacing != null) para.lineSpacing = fmt.lineSpacing;
}

// A table-style equation just inserted leaves the selection in its table: give
// the number cell (and the table's paragraph marks) the block's text font so
// "(n)" matches the body text, not the document default.
async function applyFormatToEquationTable(fmt) {
  if (!fmt || !fmt.font) return;
  try {
    await Word.run(async (context) => {
      const table = context.document.getSelection().parentTableOrNullObject;
      await context.sync();
      if (table.isNullObject) return;
      table.font.name = fmt.font;
      if (fmt.size) table.font.size = fmt.size;
      table.font.bold = false;
      table.font.italic = false;
      await context.sync();
    });
  } catch { /* cosmetic */ }
}

// Prose after an inline equation: insert after the selected picture (in the
// block's font) and move the selection past it.
async function appendTextAfterSelection(text, fmt) {
  await Word.run(async (context) => {
    const r = context.document.getSelection().getRange(Word.RangeLocation.end)
      .insertText(text, Word.InsertLocation.after);
    if (fmt && fmt.font) { r.font.name = fmt.font; if (fmt.size) r.font.size = fmt.size; r.font.bold = !!fmt.bold; r.font.italic = !!fmt.italic; }
    r.getRange(Word.RangeLocation.end).select();
    await context.sync();
  });
}

// Raw OOXML at the end of the body (used for continuous section breaks that
// switch the column count). Throws on hosts that reject it.
async function appendOoxml(pkg) {
  await Word.run(async (context) => {
    context.document.body.insertOoxml(pkg, Word.InsertLocation.end);
    await context.sync();
  });
}

// ---- Keyboard shortcuts ----
function onGlobalKeydown(e) {
  if ((e.ctrlKey || e.metaKey) && e.key === 'Enter') {
    e.preventDefault();
    if (!$('insert').disabled) onPrimaryAction().finally(refreshEquationListIfOpen);
  }
}
function onEditorKeydown(e) {
  if (e.key === 'Escape' && editingUuid) {
    e.preventDefault();
    onNewEquation();
  }
}

// ---- Settings summary (collapsed <details>) ----
function updateSettingsSummary() {
  const font = FONT_LABELS[$('font').value] || $('font').value;
  const sizeSel = $('size-mode');
  const size = sizeSel.value === 'fixed'
    ? `${$('size-pt').value} pt`
    : sizeSel.options[sizeSel.selectedIndex].textContent;
  const align = $('display-align').selectedOptions[0]?.textContent.split(' (')[0] || '';
  $('settings-summary').textContent = `${font} · ${size} · ${align}`;
  $('settings-swatch').style.background = $('color').value;
  markActiveSwatch();
}

// Remember the settings section's open state and the chosen swatch set (per
// user, best-effort).
const UI_STATE_KEY = 'mathjax-office:ui';
function restoreUiState() {
  try {
    const st = JSON.parse(localStorage.getItem(UI_STATE_KEY) || '{}');
    if (st.settingsOpen) $('settings-details').open = true;
    if (st.colorSet && findColorSet(st.colorSet).id === st.colorSet) $('color-set').value = st.colorSet;
  } catch { /* no localStorage */ }
  $('settings-details').addEventListener('toggle', saveUiState);
}
function saveUiState() {
  try {
    localStorage.setItem(UI_STATE_KEY, JSON.stringify({
      settingsOpen: $('settings-details').open,
      colorSet: $('color-set').value,
    }));
  } catch { /* ignore */ }
}

// ---- Colour swatches (colors.js) ----
function renderSwatches() {
  const set = findColorSet($('color-set').value);
  const root = $('swatches');
  root.innerHTML = '';
  for (const g of set.groups) {
    const row = document.createElement('div');
    row.className = 'swatch-row';
    if (g.name) {
      const lab = document.createElement('span');
      lab.className = 'swatch-row__label';
      lab.textContent = g.name;
      row.appendChild(lab);
    }
    for (const k of g.colors) {
      const b = document.createElement('button');
      b.type = 'button';
      b.className = 'swatch-btn';
      b.style.background = k.hex;
      b.dataset.hex = k.hex;
      b.title = `${k.name} ${k.hex}`;
      b.setAttribute('aria-label', `${k.name} ${k.hex}`);
      row.appendChild(b);
    }
    root.appendChild(row);
  }
  markActiveSwatch();
}

function markActiveSwatch() {
  const cur = normalizeHex($('color').value);
  for (const b of $('swatches').querySelectorAll('.swatch-btn')) {
    b.classList.toggle('swatch-btn--active', b.dataset.hex === cur);
  }
}

function onSwatchClick(e) {
  const b = e.target.closest('.swatch-btn');
  if (!b) return;
  $('color').value = b.dataset.hex;
  updateSettingsSummary();
  renderPreview();
}

// ---- Equation list (document order) ----
function refreshEquationListIfOpen() {
  if ($('eq-list-details').open) return refreshEquationList();
}

async function refreshEquationList() {
  const list = $('eq-list');
  try {
    const [order, byUuid] = await Promise.all([getTaggedPictureUuidsInOrder(), readAllEquations()]);
    list.innerHTML = '';
    $('eq-count').textContent = order.length ? String(order.length) : '';
    if (!order.length) {
      const li = document.createElement('li');
      li.className = 'eqlist__empty';
      li.textContent = STR.eqListEmpty;
      list.appendChild(li);
      return;
    }
    for (const uuid of order) {
      const d = byUuid[uuid];
      const li = document.createElement('li');
      const b = document.createElement('button');
      b.type = 'button';
      b.className = 'eqlist__item' + (uuid === editingUuid ? ' eqlist__item--active' : '');
      b.dataset.uuid = uuid;
      const num = document.createElement('span');
      num.className = 'eqlist__num';
      num.textContent = d && d.numbered && d.number ? formatNumber(d.number) : '–';
      const src = document.createElement('code');
      src.className = 'eqlist__src';
      src.textContent = summarizeSource(d ? d.latex : '');
      src.title = d ? d.latex : '';
      b.append(num, src);
      li.appendChild(b);
      list.appendChild(li);
    }
  } catch {
    list.innerHTML = '';
    notify('warn', STR.eqListError);
  }
}

function summarizeSource(latex) {
  const t = flattenLatex(String(latex || '')).trim();
  return t.length > 60 ? t.slice(0, 59) + '…' : t;
}

// Click → select the picture in the document; the selection handler then loads
// it into the editor exactly as a manual click would.
async function onEquationListClick(e) {
  const btn = e.target.closest('.eqlist__item');
  if (!btn) return;
  try {
    const found = await selectExistingPicture(btn.dataset.uuid);
    if (found) await onSelectionChanged();
    else notify('warn', STR.eqNotFound);
  } catch {
    notify('warn', STR.eqNotFound);
  }
}

// Apply the selected font by re-bootstrapping MathJax — the only reliable way
// to change its output font (see loadMathJax). Disables Insert while the font
// package loads, then re-renders the preview in the new font.
async function onFontChange() {
  const short = $('font').value;
  setMathJaxBusy(true, `Loading ${FONT_LABELS[short] || short} font…`);
  try {
    await ensureMathJax(short);
  } catch (e) {
    setMathJaxBusy(false);
    $('preview').textContent = `Font load failed: ${e.message}`;
    return;
  }
  setMathJaxBusy(false);
  await renderPreview();
}

// Initial MathJax load on task-pane open, using the default font selection.
async function bootMathJax() {
  setMathJaxBusy(true, 'Loading MathJax…');
  try {
    await ensureMathJax($('font').value);
  } catch (e) {
    setMathJaxBusy(false);
    $('preview').textContent = `Render engine failed to load: ${e.message}`;
    return;
  }
  setMathJaxBusy(false);
  await renderPreview();
}

// Reflect (re)loading in the UI: disable Insert while the engine or a font
// package is in flight; keep New disabled too so a reset can't race a load.
function setMathJaxBusy(busy, msg) {
  // Insert also stays disabled on hosts that can't insert at all (insertMode
  // 'none'), regardless of MathJax readiness.
  $('insert').disabled = busy || !mathJaxReady || insertMode === 'none';
  $('new-eq').disabled = busy;
  $('renumber').disabled = busy || !mathJaxReady || insertMode === 'none';
  if (busy && msg) {
    $('preview').innerHTML = '';
    $('preview').textContent = msg;
  }
}

// Ensure MathJax is loaded with `short` as the active font. No-op (resolved
// promise) when it already is; otherwise kicks off a (re)load and returns the
// shared in-flight promise so concurrent callers await the same work.
function ensureMathJax(short) {
  if (mathJaxReady && loadedFont === short) return Promise.resolve();
  // A load for this same font is already running — join it instead of tearing
  // down window.MathJax mid-startup, which can throw inside MathJax.
  if (loadingFont === short && mathJaxLoadPromise) return mathJaxLoadPromise;
  mathJaxLoadPromise = loadMathJax(short);
  return mathJaxLoadPromise;
}

// (Re)bootstrap MathJax 4 with `short` as the output font. MathJax binds its
// font to the output jax during startup and exposes no in-place swap, so a
// font change tears down window.MathJax and re-injects tex-svg.js with a new
// output.font. The alternate font's data (@mathjax/mathjax-<name>-font) is
// fetched on demand from the same CDN, which manifest.xml already allow-lists.
// Resolves once startup typesetting is ready. A generation token guards against
// a stale load (from rapid font switching) overwriting a newer one's state.
function loadMathJax(short) {
  const token = ++mathJaxLoadToken;
  loadingFont = short;

  // Tear down any prior instance so the fresh startup binds the new font.
  document.getElementById('mathjax-script')?.remove();
  document.querySelectorAll('style[id^="MJX"]').forEach((s) => s.remove());
  delete window.MathJax;
  mathJaxReady = false;
  loadedFont = null;

  // Same-origin location of the vendored MathJax (WP3.3): <base>/assets/vendor/mathjax/.
  // Resolved from the pane's own URL so it's correct in dev, under the Pages
  // subpath, and inside the Office task pane (where the document base can differ).
  const vendorBase = new URL('../assets/vendor/mathjax/', document.baseURI).href;

  return new Promise((resolve, reject) => {
    window.MathJax = {
      loader: {
        // Extensions backing the default preamble (preamble.js) + common paper
        // packages. autoload pulls any OTHER [tex] package when its command is
        // first used, so users don't need \require. (The require extension itself
        // hangs self-hosted startup — verified — so it's deliberately omitted.)
        // All served from our own origin (vendored under assets/vendor/mathjax).
        load: [
          '[tex]/ams', '[tex]/newcommand', '[tex]/mathtools', '[tex]/boldsymbol',
          '[tex]/color', '[tex]/cancel', '[tex]/physics',
          '[tex]/braket', '[tex]/autoload',
        ],
        // Override only the fonts path (its default is the CDN); extensions
        // resolve to [mathjax]/input/tex/extensions/ which is already local.
        paths: { fonts: vendorBase + 'fonts' },
      },
      tex: {
        packages: { '[+]': [
          'ams', 'newcommand', 'mathtools', 'boldsymbol', 'color',
          'cancel', 'physics', 'braket', 'autoload',
        ] },
        inlineMath: [['$', '$']],
      },
      // fontCache:'none' inlines every glyph as an SVG <path>, so each inserted
      // equation is self-contained in the Word document (no shared <defs>).
      svg: { fontCache: 'none' },
      output: {
        font: mathJaxFontName(short),
        // MathJax 4's automatic inline line-breaking splits inline math into
        // multiple <svg>s separated by <mjx-break>. We take the first <svg>, so
        // Word would otherwise see only the first fragment ("1+2+3" → "1"),
        // with spacing commands dropped at break points. Disable it.
        linebreaks: { inline: false },
      },
      // We render via tex2svgPromise and take the SVG. enableMenu:false hides the
      // a11y menu; enableSpeech/enableEnrichment:false skip per-render speech work.
      // MathJax 4 still spawns the SRE speech worker at STARTUP regardless, so the
      // sre/ files are self-hosted too (vendor-mathjax.mjs) — otherwise the worker's
      // importScripts 404s and rejects startup.promise, so mathJaxReady never flips.
      options: { enableMenu: false, enableSpeech: false, enableEnrichment: false },
      startup: {
        typeset: false,  // we render on demand via tex2svgPromise
        ready: () => {
          MathJax.startup.defaultReady();
          MathJax.startup.promise.then(async () => {
            if (token !== mathJaxLoadToken) { resolve(); return; }  // superseded
            // Register the built-in default macros ONCE (raw — do NOT flatten, or
            // the % comments would swallow the rest of the line). They persist for
            // every later render, so renders don't re-prepend them: re-running
            // \DeclareMathOperator / \DeclarePairedDelimiter errors "already
            // defined". See preamble.js (ported from the author's paper template).
            try {
              await MathJax.tex2svgPromise(DEFAULT_PREAMBLE);
            } catch (e) {
              console.warn('Default macro preamble failed to register:', e.message);
            }
            mathJaxReady = true;
            loadedFont = short;
            loadingFont = null;
            resolve();
          });
        },
      },
    };
    const s = document.createElement('script');
    s.id = 'mathjax-script';
    // Self-hosted (WP3.3): MathJax 4.1.2 + all six SVG font packages are vendored
    // under assets/vendor/mathjax (scripts/vendor-mathjax.mjs, pinned in
    // package.json) — no CDN at runtime, so it's offline-capable and font-version
    // reproducible. MathJax derives its [mathjax] root from this script's URL.
    s.src = vendorBase + 'tex-svg.js';
    s.async = true;
    s.onerror = () => {
      if (token === mathJaxLoadToken) loadingFont = null;
      reject(new Error('could not load the math renderer'));
    };
    document.head.appendChild(s);
  });
}

// ---- Settings persistence (WP1.2) ----
// Per-user defaults live in localStorage; per-document last-used (when the host
// supports it) in Office document settings and takes precedence so reopening a
// given doc restores what was last used there. settings.js validates everything,
// so a corrupt/stale blob can never reach the UI.
const SETTINGS_KEY = 'mathjax-office:settings';

function loadSavedSettings() {
  try {
    const perDoc = Office.context.document.settings.get(SETTINGS_KEY);
    if (perDoc != null) return parseSettings(perDoc);
  } catch { /* document.settings unavailable on this host */ }
  try {
    return parseSettings(localStorage.getItem(SETTINGS_KEY));
  } catch {
    return { ...DEFAULT_SETTINGS };
  }
}

function readSettingsFromUi() {
  return {
    font: $('font').value,
    sizeMode: $('size-mode').value,
    fixedPt: parseFloat($('size-pt').value) || 11,
    displayMode: getDisplayMode(),
    color: $('color').value,
    displayAlign: displayAlign(),
  };
}

function applySettingsToUi(s) {
  $('font').value = s.font;
  $('size-mode').value = s.sizeMode;
  $('size-pt').value = s.fixedPt;
  $('size-pt').disabled = s.sizeMode !== 'fixed';
  $('color').value = s.color;
  $('display-align').value = s.displayAlign || 'left';
  if (s.displayMode === 'display') $('mode-display').checked = true;
  else                             $('mode-inline').checked = true;
  updateSettingsSummary();
}

// Remember the just-used font/size/mode as defaults. Both stores are best-effort:
// localStorage can throw (private mode / quota), document.settings may be absent.
function saveCurrentSettings() {
  const settings = readSettingsFromUi();
  try { localStorage.setItem(SETTINGS_KEY, serializeSettings(settings)); } catch { /* ignore */ }
  try {
    Office.context.document.settings.set(SETTINGS_KEY, settings);
    Office.context.document.settings.saveAsync();   // fire-and-forget
  } catch { /* ignore */ }
}

function getDisplayMode() {
  return $('mode-display').checked ? 'display' : 'inline';
}

// ---- Equation numbering (WP2.2) ----
// Numbering applies to DISPLAY equations only, so the checkbox is enabled only in
// display mode; switching back to inline clears it.
function onModeChange() {
  updateNumberingUi();
  renderPreview();
}

function updateNumberingUi() {
  const isDisplay = getDisplayMode() === 'display';
  const cb = $('number-this');
  // Numbering applies to display equations. Toggling it on an EXISTING equation
  // of any style is fine: performInsert derives the structural change
  // (migrationOp) and wraps/unwraps the 1×2 table as needed.
  cb.disabled = !isDisplay;
  if (!isDisplay) cb.checked = false;
}

// Active numbering style: the equation's own when editing (so its preview/update
// match how it was created), else the document's selected style for new inserts.
function effectiveStyle() {
  return editingUuid ? editingStyle : numberingState.style;
}

// Enable the 'field' style option only where field insertion works (desktop +
// WordApi 1.5). On Word web it stays disabled so the user can't pick a style that
// would throw on insert.
function syncStyleOptions() {
  const fieldOpt = $('number-style').querySelector('option[value="field"]');
  if (fieldOpt) fieldOpt.disabled = !canInsertField;
}

// Switch the document's numbering style. New equations use it immediately;
// existing numbered equations of another style are converted only on request
// (the offer below) — restructuring tables is not something a dropdown change
// should do silently, and Undo won't roll back a multi-equation conversion.
function onNumberStyleChange() {
  numberingState = { ...numberingState, style: normalizeStyle($('number-style').value, { canInsertField }) };
  writeNumberingState(numberingState).catch(() => { /* best-effort persist */ });
  if (!editingUuid) editingStyle = numberingState.style;
  renderPreview();
  offerStyleMigration().catch(() => { /* the offer is a convenience, never fatal */ });
}

// If numbered equations of another style exist, show a one-click offer to
// convert them all to the document's style (WP2.2 increment 4).
async function offerStyleMigration() {
  if (insertMode === 'none') return;
  const byUuid = await readAllEquations();
  const plan = planStyleMigration(Object.values(byUuid), numberingState.style);
  if (!plan.length) { if (noticeFromAppFlow) clearNotice(); return; }
  const styleLabel = $('number-style').selectedOptions[0]?.textContent || numberingState.style;
  notifyAction('info',
    fmt(STR.migrateOffer, { n: plan.length }),
    fmt(STR.migrateAction, { style: styleLabel }),
    () => migrateAllToStyle(numberingState.style).finally(refreshEquationListIfOpen));
}

// Convert every numbered equation to `targetStyle` (wrap/unwrap/recell as the
// plan says), then renumber in document order so the sequence is consistent
// whatever mix of styles the document had before.
async function migrateAllToStyle(targetStyle) {
  await withDocumentOperation(STR.converting, async () => {
    const orderedUuids = await getTaggedPictureUuidsInOrder();
    const byUuid = await readAllEquations();
    const ordered = orderedUuids.map((uuid) => byUuid[uuid]).filter(Boolean);
    const plan = planStyleMigration(ordered, targetStyle);
    let done = 0, failed = 0;
    for (const { uuid, op } of plan) {
      const payload = byUuid[uuid];
      const target = { ...payload, numbered: true, numberStyle: targetStyle };
      const ok = await applyEquationChange(payload, target, op);
      if (ok) done += 1; else failed += 1;
    }
    // The equation being edited (if numbered) now lives in the new style.
    if (editingUuid && byUuid[editingUuid]?.numbered) editingStyle = targetStyle;
    await renumberAllCore();
    notify(failed ? 'warn' : 'info', fmt(STR.migrateDone, { n: done, failed }));
  });
}

// {name} placeholders in a UI string.
function fmt(str, vars) {
  return String(str).replace(/\{(\w+)\}/g, (_, k) => (k in vars ? String(vars[k]) : `{${k}}`));
}

// True when the current equation should carry a number.
function numberingActive() {
  return $('number-this').checked && getDisplayMode() === 'display';
}

// The number this equation would display: its own when editing an already-
// numbered one, otherwise the next counter value.
function prospectiveNumber() {
  return editingNumber > 0 ? editingNumber : numberingState.nextNumber;
}

// Restore this document's numbering counter + style. Absent part → defaults.
// normalizeStyle coerces a style this host can't honor (e.g. a 'field' doc opened
// on Word web) back to 'inline' so numbering never silently breaks.
async function loadNumberingState() {
  try {
    const s = await readNumberingState();
    if (s) numberingState = { style: normalizeStyle(s.style, { canInsertField }), nextNumber: s.nextNumber };
  } catch {
    /* non-fatal: numbering just starts from the default counter */
  }
  $('number-style').value = numberingState.style;
}

// Apply the active locale's strings to every [data-i18n] element (WP3.4). Dynamic
// strings (status/buttons) are localized in updateUiState; notices/errors are not
// yet localized (English fallback).
function applyI18n() {
  for (const el of document.querySelectorAll('[data-i18n]')) {
    const v = STR[el.dataset.i18n];
    if (v) el.textContent = v;
  }
}

function updateUiState() {
  if (editingUuid) {
    $('status').textContent = STR.editing.replace('{id}', editingUuid.slice(0, 8));
    $('status').className = 'editing';
    $('insert').textContent = STR.update;
  } else {
    $('status').textContent = STR.statusNew;
    $('status').className = 'new';
    $('insert').textContent = STR.insert;
  }
}

// Prepend the document macro preamble (WP2.4) before MathJax sees the equation,
// then flatten newlines so \newcommand defs expand in every render. Empty preamble
// → just the flattened equation. Shared by the preview and the insert path.
function withPreamble(latex) {
  const pre = $('preamble').value.trim();
  return flattenLatex(pre ? pre + '\n' + latex : latex);
}

// Macro preamble persistence: re-render live on each keystroke; debounce the write
// to the document's XML part.
let preambleSaveTimer = null;
function onPreambleInput() {
  renderPreview();
  clearTimeout(preambleSaveTimer);
  preambleSaveTimer = setTimeout(savePreamble, 600);
}
async function savePreamble() {
  try {
    await writePreamble($('preamble').value);
  } catch (e) {
    notifyResult(describeInsertError('store', e));
  }
}
async function loadPreamble() {
  try {
    const text = await readPreamble();
    if (text) {
      $('preamble').value = text;
      renderPreview();
    }
  } catch {
    /* non-fatal: the preamble just won't preload */
  }
}

async function renderPreview() {
  if (!mathJaxReady) return;
  const tex = $('latex').value.trim();
  const preview = $('preview');
  const inline = getDisplayMode() !== 'display';
  preview.style.color = $('color').value;   // MathJax SVG uses currentColor
  preview.innerHTML = '';
  preview.classList.toggle('preview--inline', inline && !!tex);
  if (!tex) { showRenderError(null); return; }
  // WYSIWYG: when numbering this equation, preview it with the number it would
  // get (the next counter value, or its own when editing) so the user sees the
  // real result. The number is decorated only for rendering — never stored.
  const render = numberingActive()
    ? decorateLatexForNumber(tex, prospectiveNumber(), effectiveStyle())
    : tex;
  try {
    const node = await MathJax.tex2svgPromise(withPreamble(render), { display: !inline });
    // MathJax draws TeX errors as an <merror> instead of throwing; surface the
    // message under the source so the user sees WHAT is wrong, not just a red box.
    showRenderError(findRenderError(node));
    if (inline) {
      // In context: a line of document-like text around the equation, so the
      // baseline and size read as they will in the paragraph.
      preview.append(contextSpan(STR.previewCtxBefore), node, contextSpan(STR.previewCtxAfter));
    } else {
      preview.appendChild(node);
    }
  } catch (e) {
    showRenderError(e.message || String(e));
  }
}

function contextSpan(text) {
  const span = document.createElement('span');
  span.className = 'ctx';
  span.textContent = text;
  return span;
}

function showRenderError(msg) {
  const box = $('latex-error');
  const editor = document.querySelector('.editor');
  if (msg) {
    box.textContent = msg;
    box.hidden = false;
    editor.classList.add('editor--error');
  } else {
    box.textContent = '';
    box.hidden = true;
    editor.classList.remove('editor--error');
  }
}

async function resolveSizes() {
  return Word.run(async (context) => {
    const sel = context.document.getSelection();
    const body = context.document.body;
    sel.font.load('size');
    body.font.load('size');
    await context.sync();
    return {
      selectionPt: typeof sel.font.size === 'number' ? sel.font.size : null,
      bodyPt: body.font.size || 11,
    };
  });
}

// Inline equations: crop the viewBox so its bottom edge coincides with the
// math baseline (y=0 in MathJax SVG coords). Word places inline-picture
// bottoms on the text baseline, so this aligns the two without relying on
// the `vertical-align` CSS that Word's XmlSvg coercion drops.
//
// To avoid clipping descenders (bottoms of integrals, fractions, big
// operators, letters like y/g/j), we also set overflow="visible" so the
// SVG content below the cropped viewport still renders — falling into the
// line's natural descender region, where text descenders like `y` already
// live. This relies on Word's XmlSvg renderer honoring SVG2's overflow
// semantics; if it doesn't, descenders go back to being clipped (no
// regression — that's just the prior alignBaselineToBottom behavior).
// True iff a real glyph extends below the math baseline (y > 0 in MathJax
// SVG coords). Reads viewBox descent (vy + vh); MathJax v4 sets this to
// ≤83 internal units for purely-above-baseline content (metric padding
// only) and ≥158 for content with actual descenders. 100 is the separator.
// alignBaselineToBottom is kept for historical reference but no longer called.
// We now preserve full SVG content and shift the run via Word.Font.position
// (WordApiDesktop 1.3) so math-baseline aligns with text-baseline without
// clipping descenders. See onPrimaryAction.
function alignBaselineToBottom(svgEl) {
  const vb = (svgEl.getAttribute('viewBox') || '').split(/\s+/).map(parseFloat);
  if (vb.length !== 4) return;
  const [vx, vy, vw, vh] = vb;

  const ascentVb = -vy;
  const descentVb = vy + vh;
  if (ascentVb <= 0 || descentVb <= 0.5) return;  // nothing meaningful to crop

  svgEl.setAttribute('viewBox', `${vx} ${vy} ${vw} ${ascentVb}`);
  svgEl.setAttribute('overflow', 'visible');

  const heightAttr = svgEl.getAttribute('height') || '';
  const hMatch = heightAttr.match(/^([\d.]+)(ex|pt|px|em)?$/);
  if (hMatch) {
    const totalH = parseFloat(hMatch[1]);
    const unit = hMatch[2] || '';
    svgEl.setAttribute('height', (totalH * (ascentVb / vh)).toFixed(3) + unit);
  }

  // One pass on the style attribute: strip MathJax's vertical-align (Word
  // drops it anyway, and stale values confuse downstream readers), and add
  // overflow:visible as a belt-and-braces companion to the attribute above.
  const styleAttr = svgEl.getAttribute('style') || '';
  let newStyle = styleAttr.replace(/vertical-align:\s*-?[\d.]+(?:ex|em|px|pt);?\s*/gi, '');
  if (!/overflow\s*:/i.test(newStyle)) {
    newStyle = (newStyle ? newStyle.replace(/;?\s*$/, ';') : '') + 'overflow:visible';
  }
  if (newStyle !== styleAttr) svgEl.setAttribute('style', newStyle);
}

async function buildSvgForInsert(latex, font, targetPt, displayMode, color) {
  const node = await MathJax.tex2svgPromise(withPreamble(latex), { display: displayMode === 'display' });
  const svgEl = node.querySelector('svg');
  // Inline: descent (in pt) from the MathJax viewBox so the caller can shift the
  // inserted picture down via Word.Font.position. The SVG is left intact (no
  // cropping) — descenders stay visible. See descentPtFromViewBox in mathsvg.js.
  const descentPt = displayMode === 'inline'
    ? descentPtFromViewBox(svgEl.getAttribute('viewBox'), targetPt)
    : 0;
  scaleSvgToPt(svgEl, targetPt);
  inlineColors(svgEl, color);
  if (!svgEl.getAttribute('xmlns')) {
    svgEl.setAttribute('xmlns', 'http://www.w3.org/2000/svg');
  }
  // scaleSvgToPt wrote width/height in pt; the raster path (WP1.5) needs them to
  // size the PNG canvas and the insert. Fall back to targetPt if somehow unset.
  const widthPt = parseFloat(svgEl.getAttribute('width')) || targetPt;
  const heightPt = parseFloat(svgEl.getAttribute('height')) || targetPt;
  return {
    svg: new XMLSerializer().serializeToString(svgEl),
    descentPt,
    widthPt,
    heightPt,
  };
}

function insertSvgAtSelection(svgString) {
  return new Promise((resolve, reject) => {
    Office.context.document.setSelectedDataAsync(
      svgString,
      { coercionType: Office.CoercionType.XmlSvg },
      (res) => res.status === Office.AsyncResultStatus.Succeeded
        ? resolve() : reject(res.error),
    );
  });
}

// Tag the just-inserted inline picture for click-to-edit, and — only when
// allowBaselineShift (the desktop SVG path) and the host supports it — lower the
// run by descentPt so the math baseline meets the text baseline. Shared by the
// SVG and (WP1.5) PNG insert paths. Returns true if a picture was found+tagged.
async function tagSelectedPicture(uuid, latex, descentPt, displayMode, allowBaselineShift) {
  return Word.run(async (context) => {
    const sel = context.document.getSelection();
    const pics = sel.inlinePictures;
    pics.load('items');
    await context.sync();
    const pic = pics.items[pics.items.length - 1];
    if (!pic) return false;
    pic.altTextTitle = ALT_PREFIX + uuid;
    pic.altTextDescription = latex;
    await context.sync();
    if (allowBaselineShift && canShiftBaseline && displayMode === 'inline' && descentPt > 0) {
      // Font.position (WordApiDesktop 1.3) lowers the run. Office.js surfaces an
      // unsupported-API rejection at context.sync(), NOT at the assignment, so the
      // sync MUST be inside the try or it fails the whole Word.run. On failure the
      // math floats above baseline by descentPt but stays visible; the tag/alt-text
      // already synced, so storage is unaffected.
      //
      // Word sizes the imported SVG to whole pixels and records the remainder as
      // a bottom "effect extent" (measured 2026-09-15: wp:effectExtent b="9525"
      // EMU = 0.75 pt on a 9 pt insert; 0 on OOXML we write ourselves). The box
      // Word seats on the baseline INCLUDES that margin, so the ink ends that much
      // above the baseline before any shift — a plain -descent left the math
      // ~0.75 pt high. InlinePicture exposes no effectExtent, so read the run's
      // OOXML back and add the margin to the shift. Word also clamps a picture's
      // lowering to its own height (calibrated), never reached here.
      try {
        let marginPt = 0;
        try {
          const ox = pic.getRange('Whole').getOoxml();
          await context.sync();
          const m = /<wp:effectExtent\b[^>]*\bb="(\d+)"/.exec(ox.value || '');
          if (m) marginPt = parseInt(m[1], 10) / 12700;
        } catch { /* no OOXML readback: fall back to the pure descent */ }
        pic.getRange('Whole').font.position = -(descentPt + marginPt);
        await context.sync();
      } catch (e) {
        console.warn('Font.position unavailable; baseline not shifted:', e.message);
        notify('warn', fmt(STR.baselineShiftFailed, { error: e.message }));
      }
    }
    return true;
  });
}

// Rasterize an SVG string to a base64 PNG sized for the web insert path. Draws
// via a same-origin Blob URL (not a data: URI) so the canvas isn't tainted —
// MathJax paths are inlined, so toDataURL is permitted. Renders at ~300 DPI
// (print quality) via computeRasterSize, regardless of the screen's pixel ratio.
function svgToPngBase64(svgString, wPt, hPt) {
  const { pxW, pxH } = computeRasterSize(wPt, hPt, window.devicePixelRatio || 2);
  return new Promise((resolve, reject) => {
    const url = URL.createObjectURL(new Blob([svgString], { type: 'image/svg+xml;charset=utf-8' }));
    const img = new Image();
    img.onload = () => {
      try {
        const canvas = document.createElement('canvas');
        canvas.width = pxW;
        canvas.height = pxH;
        canvas.getContext('2d').drawImage(img, 0, 0, pxW, pxH);
        resolve(stripDataUri(canvas.toDataURL('image/png')));
      } catch (e) {
        reject(e);
      } finally {
        URL.revokeObjectURL(url);
      }
    };
    img.onerror = () => {
      URL.revokeObjectURL(url);
      reject(new Error('could not rasterize the equation'));
    };
    img.src = url;
  });
}

// Insert a base64 PNG at the selection, sized in points. ImageCoercion 1.1 (the
// raster path) supports Word on the web; imageWidth/imageHeight are in points.
function insertPngAtSelection(base64, wPt, hPt) {
  return new Promise((resolve, reject) => {
    Office.context.document.setSelectedDataAsync(
      base64,
      { coercionType: Office.CoercionType.Image, imageWidth: wPt, imageHeight: hPt },
      (res) => res.status === Office.AsyncResultStatus.Succeeded
        ? resolve() : reject(res.error),
    );
  });
}

// Place the equation picture at the current selection: SVG via XmlSvg coercion on
// desktop, raster PNG (ImageCoercion 1.1) on the web. Returns null on success, or
// a {level,msg} notice on a pre-insert failure (rasterize / too-large) that the
// caller should surface. The actual insert call can still throw — callers wrap it
// so a HARD failure stops before tagging/storing a picture that doesn't exist.
// Shared by the insert path and Renumber-all.
// Rasterize an SVG to a base64 PNG for the web path. Returns {base64} on success
// or {err:{level,msg}} for a rasterize failure / too-large payload. Shared by the
// selection insert and the table-cell insert.
async function rasterizeForInsert(svgString, widthPt, heightPt) {
  let base64;
  try {
    base64 = await svgToPngBase64(svgString, widthPt, heightPt);
  } catch (e) {
    return { err: describeInsertError('insert', e) };
  }
  if (isPayloadTooLarge(base64)) return { err: { level: 'error', msg: TOO_LARGE_MSG } };
  return { base64 };
}

async function insertPictureAtSelection(svgString, widthPt, heightPt) {
  if (insertMode === 'svg') {
    await insertSvgAtSelection(svgString);
    return null;
  }
  const { base64, err } = await rasterizeForInsert(svgString, widthPt, heightPt);
  if (err) return err;
  await insertPngAtSelection(base64, widthPt, heightPt);
  return null;
}

// Fallback tagger for the raster path: after an Image insert the new picture may
// not be reachable via getSelection().inlinePictures, so tag the newest picture
// in the body instead. No baseline shift (web lacks WordApiDesktop 1.3). Assumes
// newest = last in body (true at/after document end; a mid-document insert may
// mis-tag — acceptable: storage still holds the source under this uuid).
async function tagLastBodyPicture(uuid, latex) {
  try {
    await Word.run(async (context) => {
      const pics = context.document.body.inlinePictures;
      pics.load('items');
      await context.sync();
      const pic = pics.items[pics.items.length - 1];
      if (pic) {
        pic.altTextTitle = ALT_PREFIX + uuid;
        pic.altTextDescription = latex;
        await context.sync();
      }
    });
  } catch (e) {
    notifyResult(describeInsertError('tag', e));
  }
}

// Put a new display equation on its own paragraph so centering it doesn't drag
// neighboring text along. Returns true if it moved the insertion point to a fresh
// paragraph. Only meaningful for NEW inserts (an update replaces in place).
//
// [unverified — needs a Word-desktop spike] The exact collapsed-caret landing
// after insertParagraph isn't pinned by the Office docs, so we re-collapse to the
// new paragraph's start explicitly. If the caret is already on an empty paragraph
// (or a list item) we leave it alone. Worst case the equation shares its line as
// before WP1.4 — no regression, and the emptiness-gated center below still guards
// against centering prose.
async function ensureOwnParagraphForDisplay() {
  return Word.run(async (context) => {
    const sel = context.document.getSelection();
    const para = sel.paragraphs.getFirst();
    para.load('text,isListItem');
    await context.sync();
    if (para.isListItem) return false;                  // don't disturb list structure
    if (isParagraphTextEmpty(para.text)) return false;  // already on its own line
    const newPara = sel.insertParagraph('', Word.InsertLocation.after);
    newPara.getRange(Word.RangeLocation.start).select();
    await context.sync();
    return true;
  });
}

async function onPrimaryAction() {
  if (!mathJaxReady || insertMode === 'none') return;
  const latex = $('latex').value.trim();
  if (!latex) return;
  clearNotice();
  const result = await performInsert({
    latex,
    font: $('font').value,
    color: $('color').value,
    displayMode: getDisplayMode(),
    sizeMode: $('size-mode').value,
    fixedPt: parseFloat($('size-pt').value) || 11,
    numbered: numberingActive(),
    style: effectiveStyle(),
    align: displayAlign(),
    editingUuid,
    editingNumber,
  });
  if (!result) return;
  // Stay in edit mode for the equation we just (re)inserted.
  editingUuid = result.uuid;
  editingNumber = result.number;
  editingStyle = result.number > 0 ? result.style : numberingState.style;
  updateUiState();
  updateNumberingUi();
  saveCurrentSettings();   // remember the just-used font/size/mode as defaults
}

// The insert pipeline proper, driven by an explicit request instead of the pane
// controls so the test-page runner can call it too. req = { latex, font, color,
// displayMode, sizeMode, fixedPt, numbered, style, editingUuid, editingNumber }.
// Returns { uuid, number, style } on success, null on a hard failure (already
// reported via notify).
async function performInsert(req) {
  if (!mathJaxReady || insertMode === 'none') return null;
  const latex = String(req.latex || '').trim();
  if (!latex) return null;
  const { font, color, displayMode, sizeMode, fixedPt } = req;
  let editing = req.editingUuid || null;

  // Resolving the host font size can fail (host quirk / transient). Degrade to
  // the body default rather than aborting the whole insert. SOFT: warn (the
  // insert continues), never the hard 'insert' error — that would falsely claim
  // the insert failed and the red notice would outlive the successful insert.
  let selectionPt = null, bodyPt = 11;
  try {
    ({ selectionPt, bodyPt } = await resolveSizes());
  } catch (e) {
    notify('warn', 'Could not read the surrounding text size; using the document default.');
  }
  const sizePt = computeRenderSize({ mode: sizeMode, selectionPt, bodyPt, fixedPt });
  const skipPt = displaySkipPt(bodyPt);
  const align = req.align === 'center' ? 'center' : req.align === 'left' ? 'left' : displayAlign();

  const uuid = editing || newUuid();

  // The stored payload of the equation being edited: its numbered state + style
  // tell us where its number lives NOW, which decides the structural change
  // below (migrationOp). Missing part (recovered picture) → bare picture.
  let prior = null;
  if (editing) {
    try { prior = await readXmlPartByUuid(editing); } catch { prior = null; }
  }

  // Numbering (WP2.2): decide this equation's number + style before rendering.
  // Editing keeps the equation's own style and (if numbered) its number; a freshly-
  // numbered equation consumes the counter. The number is baked into the render
  // only for the inline style (decorateLatexForNumber); the RAW latex is always
  // what gets stored, so renumber can re-apply a fresh number cleanly.
  // An already-numbered equation keeps its style even where this host could not
  // CREATE it (a field equation edited on Word web): nothing new is inserted,
  // so normalizing would wrongly unwrap it. Only a new number is host-gated.
  const numbered = !!req.numbered;
  const keepsStyle = numbered && prior && prior.numbered && prior.numberStyle === req.style;
  const style = numbered ? (keepsStyle ? req.style : normalizeStyle(req.style, { canInsertField })) : 'inline';
  let number = 0;
  let counterConsumed = false;
  if (numbered) {
    const r = nextNumberFor(numberingState, editing ? (req.editingNumber || 0) : 0);
    number = r.number;
    counterConsumed = r.nextNumber !== numberingState.nextNumber;
    numberingState = { ...numberingState, nextNumber: r.nextNumber };
  }
  const renderLatex = decorateLatexForNumber(latex, number, style);

  // Guarantee the active MathJax font matches the dropdown before rendering the
  // SVG that gets inserted. Normally onFontChange already settled this; this is
  // cheap insurance (resolves immediately when the font is already loaded).
  await ensureMathJax(font);
  const { svg, descentPt, widthPt, heightPt } = await buildSvgForInsert(renderLatex, font, sizePt, displayMode, color);
  const svgString = embedSourceInSvg(svg, { uuid, latex, font, sizePt, displayMode, color, numbered, number, numberStyle: style });

  if (editing) {
    // Update path: re-locate the existing picture so the insert replaces it.
    let found = false;
    try {
      found = await selectExistingPicture(editing);
    } catch (e) {
      // Can't even look for it — abort rather than insert a stray duplicate.
      notifyResult(describeInsertError('update-locate', e));
      return null;
    }
    if (!found) {
      // Picture vanished (user deleted it); fall through to insert as new.
      editing = null;
      notifyResult(pictureVanishedNotice());
    }
  }

  // A NEW numbered equation in the table OR field style is placed as a borderless
  // 1×2 table (its own block, with the number cell — static text or a SEQ field).
  // An UPDATE whose numbering moves the picture into or out of a table (or swaps
  // the number cell between static text and a SEQ field) goes through the same
  // structural ops as a style migration. Everything else — inline-style inserts
  // and in-place updates — takes the selection path below.
  const placement = numberPlacement(style);
  const newTableEquation = numbered && !editing && (placement === 'cell' || placement === 'field');
  const structOp = editing ? migrationOp(prior, { numbered, style }) : 'none';
  const stored = { uuid, latex, font, sizePt, displayMode, color, numbered, number, numberStyle: style };

  if (editing && (structOp === 'wrap' || structOp === 'unwrap')) {
    // HARD failure → the ops leave the old picture untouched on failure, so stop
    // before storing a payload that describes a structure that doesn't exist.
    let ok = false;
    try {
      ok = structOp === 'wrap'
        ? await wrapPictureIntoTable(stored, { widthPt, heightPt }, svgString)
        : await unwrapPictureFromTable(stored, { widthPt, heightPt, descentPt }, svgString);
    } catch (e) {
      notifyResult(describeInsertError('insert', e));
      return null;
    }
    if (!ok) {
      notifyResult(describeInsertError('insert', new Error('could not restructure the numbered equation')));
      return null;
    }
  } else if (newTableEquation) {
    // HARD failure → stop before tagging/storing a picture that doesn't exist.
    // insertTableEquation* handle placement, the number cell, and tagging.
    try {
      if (insertMode === 'svg') {
        await insertTableEquationSvg({ uuid, latex, number, style, widthPt, heightPt, skipPt, align }, svgString);
      } else {
        // Web/raster path; 'field' is desktop-gated so style is 'table' here.
        const { base64, err } = await rasterizeForInsert(svgString, widthPt, heightPt);
        if (err) { notifyResult(err); return null; }
        await insertTableEquationPng({ uuid, latex, number, style, skipPt, align }, base64);
      }
    } catch (e) {
      notifyResult(describeInsertError('insert', e));
      return null;
    }
  } else {
    // Display equations go on their own line so centering doesn't drag neighboring
    // text along. New inserts only — an update replaces in place. SOFT.
    if (displayMode === 'display' && !editing) {
      try {
        await ensureOwnParagraphForDisplay();
      } catch (e) {
        notify('warn',
          'Could not place the equation on its own line; ' +
          'it may share a line with nearby text.');
      }
    }

    // Insert (SVG on desktop, raster PNG on web). HARD failure: nothing was
    // placed, so surface it and stop before tagging/storing a phantom picture.
    try {
      const err = await insertPictureAtSelection(svgString, widthPt, heightPt);
      if (err) { notifyResult(err); return null; }
    } catch (e) {
      notifyResult(describeInsertError('insert', e));
      return null;
    }

    // For display mode, center the paragraph and give it TeX's display skips
    // above and below — but ONLY if it contains nothing but the equation, so we
    // never center or re-space the user's prose. SOFT: a failure here must not
    // erase a placed equation.
    if (displayMode === 'display') {
      try {
        await Word.run(async (context) => {
          const p = context.document.getSelection().paragraphs.getFirst();
          p.load('text');
          await context.sync();
          if (isParagraphTextEmpty(p.text)) {
            alignDisplayParagraph(p, align);
            p.spaceBefore = skipPt;
            p.spaceAfter = skipPt;
            await context.sync();
          }
        });
      } catch (e) {
        notify('warn', 'Equation inserted, but centering its paragraph failed.');
      }
    }

    // Tag for click-to-edit and (desktop) baseline-shift. SOFT: the equation is
    // already placed; a tag failure only costs click-to-edit, not the insert. If
    // the selection-based tag found no picture (raster path on web), fall back to
    // tagging the newest picture in the body.
    let tagged = false;
    try {
      tagged = await tagSelectedPicture(uuid, latex, descentPt, displayMode, insertMode === 'svg');
    } catch (e) {
      notifyResult(describeInsertError('tag', e));
    }
    if (!tagged) await tagLastBodyPicture(uuid, latex);

    // Table ↔ field on an update: the picture was replaced in its cell above;
    // now swap the number cell's content. SOFT: the equation is in place.
    if (structOp === 'recell') {
      try {
        if (!(await rewriteNumberCell(uuid, style, number))) {
          notify('warn', 'Updated the equation, but its number cell could not be rewritten.');
        }
      } catch (e) {
        notify('warn', 'Updated the equation, but its number cell could not be rewritten.');
      }
    }
  }

  // Persist the source payload (storage handles replace-by-uuid). SOFT: a
  // storage failure costs click-to-edit recovery, but the equation (and its
  // alt text) is already in the document.
  try {
    await writeXmlPart(stored);
  } catch (e) {
    notifyResult(describeInsertError('store', e));
  }

  // Persist the bumped counter so the NEXT equation gets the next number, even
  // across pane reloads. SOFT: only the counter; the equation is already stored.
  if (counterConsumed) {
    try {
      await writeNumberingState(numberingState);
    } catch (e) {
      notifyResult(describeInsertError('store', e));
    }
  }

  return { uuid, number, style };
}

// Collect every tagged equation picture in document order, INCLUDING pictures
// inside table cells. body.inlinePictures does NOT reliably descend into tables
// (verified against the Word API docs), so we recurse body → tables → rows →
// cells → cell.body. Returns picture proxies (valid in `context`) with their uuid
// and altTextTitle loaded. Used by edit-locate and Renumber-all so the flush-right
// table style (where equations live in cells) is fully supported.
async function collectTaggedPictures(context) {
  const out = [];
  async function scan(body) {
    const pics = body.inlinePictures;
    const tables = body.tables;
    pics.load('items/altTextTitle');
    tables.load('items');
    await context.sync();
    for (const p of pics.items) {
      const t = p.altTextTitle || '';
      if (t.startsWith(ALT_PREFIX)) out.push({ pic: p, uuid: t.slice(ALT_PREFIX.length) });
    }
    for (const table of tables.items) {
      const rows = table.rows;
      rows.load('items');
      await context.sync();
      for (const row of rows.items) {
        const cells = row.cells;
        cells.load('items');
        await context.sync();
        for (const cell of cells.items) await scan(cell.body);
      }
    }
  }
  await scan(context.document.body);
  return out;
}

// Find a picture tagged with the given UUID and select its range so a subsequent
// setSelectedDataAsync replaces it. Returns true if found (incl. in a table cell).
async function selectExistingPicture(uuid) {
  return Word.run(async (context) => {
    const hit = (await collectTaggedPictures(context)).find((t) => t.uuid === uuid);
    if (!hit) return false;
    hit.pic.getRange().select();
    await context.sync();
    return true;
  });
}

// ---- Renumber all (WP2.2) ----
// Static numbering: a number is assigned when an equation is inserted, so numbers
// follow insertion order, not document order. "Renumber all" reassigns 1..n in
// document order and re-renders the affected pictures.

// UUIDs of all tagged equation pictures, in document order — including those in
// table cells (collectTaggedPictures recurses), so flush-right table equations
// are renumbered too.
async function getTaggedPictureUuidsInOrder() {
  return Word.run(async (context) => {
    return (await collectTaggedPictures(context)).map((t) => t.uuid);
  });
}

// ---- Strategy B: flush-right number in a borderless 1×2 table (WP2.2) ----

// Format a freshly-inserted 1×2 numbering table: strip all borders, fit the text
// column, center the equation cell, and pin a narrow right-aligned number cell.
// All queued (no sync) so the caller batches it. TableCell.width is read-only —
// columnWidth is the settable one.
function formatNumberTable(table, skipPt = 0, align = 'left') {
  for (const loc of ['Top', 'Bottom', 'Left', 'Right', 'InsideHorizontal', 'InsideVertical']) {
    table.getBorder(loc).type = Word.BorderType.none;
  }
  table.autoFitWindow();
  table.alignment = Word.Alignment.left;
  const eqCell = table.getCell(0, 0);
  const numCell = table.getCell(0, 1);
  eqCell.horizontalAlignment = align === 'center' ? Word.Alignment.centered : Word.Alignment.left;
  if (align !== 'center') eqCell.body.paragraphs.getFirst().leftIndent = MATH_INDENT_PT;
  numCell.horizontalAlignment = Word.Alignment.right;
  numCell.columnWidth = NUMBER_CELL_WIDTH_PT;
  // Word's Normal Table style pads cells 5.4 pt left/right, which would leave
  // the number that far short of the right text border and shift the fleqn
  // indent. Zero the horizontal padding (WordApi 1.3), as the headless emitter
  // does with tblCellMar, so "(n)" sits exactly on the border.
  for (const cell of [eqCell, numCell]) {
    cell.setCellPadding(Word.CellPaddingLocation.left, 0);
    cell.setCellPadding(Word.CellPaddingLocation.right, 0);
    // The number sits in the vertical middle of the row — i.e. of the equation
    // picture, since both cells carry the same display skips (Word's default is
    // top-aligned; the headless emitter writes w:vAlign center).
    cell.verticalAlignment = Word.VerticalAlignment.centered;
  }
  // TeX display skips, carried by the cell paragraphs (both cells, so the row
  // height is the same either side).
  if (skipPt > 0) {
    for (const cell of [eqCell, numCell]) {
      const p = cell.body.paragraphs.getFirst();
      p.spaceBefore = skipPt;
      p.spaceAfter = skipPt;
    }
  }
}

// Fill the right-hand number cell. 'field' style drops a native SEQ field that
// Word auto-numbers (rendered as "(n)" between literal parens); every other style
// writes static "(n)" text. Queued on the table-creation context (no sync). SEQ
// insertion is WordApi 1.5 / desktop-only — only reached when canInsertField.
function fillNumberCell(table, style, number) {
  fillNumberCellOf(table.getCell(0, 1), style, number);
}
function fillNumberCellOf(numCell, style, number) {
  if (style === 'field') {
    const body = numCell.body;
    body.clear();
    const open = body.insertText('(', Word.InsertLocation.start);
    body.insertText(')', Word.InsertLocation.end);
    open.getRange(Word.RangeLocation.after).insertField(
      Word.InsertLocation.after, Word.FieldType.seq, 'equation \\* ARABIC', false);
  } else {
    numCell.value = formatNumber(number);
  }
}

// Insert a NEW numbered equation as a flush-right table (web/raster path): the
// equation PNG goes in the left cell via the Word-native
// insertInlinePictureFromBase64 (web-supported, returns the proxy so we tag in
// the same batch — no selection round-trip). The 'field' style never reaches the
// web path (it's desktop-gated), so this is always the static-number style.
async function insertTableEquationPng({ uuid, latex, number, style, skipPt = 0, align = 'left' }, base64) {
  return Word.run(async (context) => {
    const para = context.document.getSelection().paragraphs.getFirst();
    const table = para.insertTable(1, 2, Word.InsertLocation.after, [['', '']]);
    formatNumberTable(table, skipPt, align);
    fillNumberCell(table, style, number);
    const pic = table.getCell(0, 0).body.insertInlinePictureFromBase64(base64, Word.InsertLocation.replace);
    pic.altTextTitle = ALT_PREFIX + uuid;
    pic.altTextDescription = latex;
    await context.sync();
  });
}

// Insert a NEW numbered equation as a flush-right table (desktop/SVG path).
// Three tiers, because dropping an SVG into a table cell has host quirks we
// cannot probe from outside Word:
//   1. Word API: insertInlinePictureFromBase64 with the SVG straight into the
//      equation cell (no selection round-trip; Word desktop stores SVG natively).
//   2. Common API: collapse the selection to an insertion point at the cell START
//      (selecting the cell's whole content makes a CELL selection, where image
//      insertion is unsupported) and setSelectedDataAsync(XmlSvg).
//   3. Raster: PNG into the cell via the web path (warns: not vector).
// If all fail the empty table is deleted and the error surfaces as a hard insert
// failure — never a stray table without an equation.
async function insertTableEquationSvg({ uuid, latex, number, style, widthPt, heightPt, skipPt = 0, align = 'left' }, svgString) {
  // Tier 1 — build the table and try the direct in-cell insert in one batch.
  const tier1 = await Word.run(async (context) => {
    const para = context.document.getSelection().paragraphs.getFirst();
    const table = para.insertTable(1, 2, Word.InsertLocation.after, [['', '']]);
    formatNumberTable(table, skipPt, align);
    fillNumberCell(table, style, number);
    await context.sync();
    try {
      const pic = table.getCell(0, 0).body.insertInlinePictureFromBase64(utf8ToBase64(svgString), Word.InsertLocation.start);
      pic.altTextTitle = ALT_PREFIX + uuid;
      pic.altTextDescription = latex;
      await context.sync();
      return true;
    } catch (e) {
      console.warn('SVG in-cell insert unsupported, trying selection path:', e.message);
      // Leave an insertion point at the start of the equation cell for tier 2.
      table.getCell(0, 0).body.getRange(Word.RangeLocation.start).select();
      await context.sync();
      return false;
    }
  });
  if (tier1) return;

  // Tier 2 — Common API at the insertion point inside the cell.
  let tier2 = false;
  try {
    await insertSvgAtSelection(svgString);
    tier2 = await tagPictureInSelectedCell(uuid, latex);
  } catch (e) {
    console.warn('Selection-path SVG insert failed in cell:', e.message);
  }
  if (tier2) return;

  // Tier 3 — raster fallback into the cell (the selection is still in it).
  const { base64, err } = await rasterizeForInsert(svgString, widthPt, heightPt);
  if (!err) {
    const tier3 = await Word.run(async (context) => {
      const cell = context.document.getSelection().parentTableCellOrNullObject;
      await context.sync();
      if (cell.isNullObject) return false;
      const pic = cell.body.insertInlinePictureFromBase64(base64, Word.InsertLocation.start);
      pic.altTextTitle = ALT_PREFIX + uuid;
      pic.altTextDescription = latex;
      await context.sync();
      return true;
    });
    if (tier3) {
      notify('warn', 'Inserted the numbered equation as a raster image (this host would not place an SVG in the table cell).');
      return;
    }
  }

  // Nothing worked: remove the empty table so the document isn't left with debris.
  await Word.run(async (context) => {
    const table = context.document.getSelection().parentTableOrNullObject;
    await context.sync();
    if (!table.isNullObject) { table.delete(); await context.sync(); }
  });
  throw new Error(err ? err.msg : 'could not place the equation in the numbering table');
}

// After a Common-API insert into a cell, tag the picture now living in that
// cell (via the selection, else via the containing cell). Returns true if tagged.
async function tagPictureInSelectedCell(uuid, latex) {
  return Word.run(async (context) => {
    const sel = context.document.getSelection();
    let pics = sel.inlinePictures;
    pics.load('items');
    await context.sync();
    let pic = pics.items[pics.items.length - 1];
    if (!pic) {
      const cell = sel.parentTableCellOrNullObject;
      await context.sync();
      if (!cell.isNullObject) {
        pics = cell.body.inlinePictures;
        pics.load('items');
        await context.sync();
        pic = pics.items[pics.items.length - 1];
      }
    }
    if (!pic) return false;
    pic.altTextTitle = ALT_PREFIX + uuid;
    pic.altTextDescription = latex;
    await context.sync();
    return true;
  });
}

function utf8ToBase64(str) {
  return btoa(unescape(encodeURIComponent(str)));
}

// Renumber a table-style equation: rewrite the sibling number cell's text. No
// picture re-render needed (the number is Word text, not in the SVG), so this is
// far cheaper than the inline-style re-render. Returns false if the picture or its
// table has vanished.
async function renumberTableCell(payload, number) {
  const ok = await rewriteNumberCell(payload.uuid, payload.numberStyle, number);
  if (ok) {
    try {
      await writeXmlPart({ ...payload, numbered: true, number });
    } catch (e) {
      notifyResult(describeInsertError('store', e));
    }
  }
  return ok;
}

// Refresh every SEQ field so field-style numbers reflect current document order.
// Word owns field numbering, so renumbering field equations is just a field
// update — there is no bulk-update API, so iterate getByTypes([seq]) + updateResult.
async function updateAllSeqFields() {
  await Word.run(async (context) => {
    const fields = context.document.body.fields.getByTypes([Word.FieldType.seq]);
    fields.load('items');
    await context.sync();
    for (const f of fields.items) f.updateResult();
    await context.sync();
  });
}

// Re-render one inline-style equation with `number` baked in and replace its
// picture in place. Returns false if the picture has vanished.
async function rerenderEquationWithNumber(payload, number) {
  return applyEquationChange(payload, { ...payload, numbered: true, number }, 'rerender');
}

// ---- Structural ops (WP2.2 increment 4) ----
// One equation, from its stored `payload` to `target` (a full payload with the
// new numbered/number/numberStyle), performing the structural `op` from
// migrationOp: re-render in place, rewrite the number cell, wrap into a 1×2
// table, or unwrap back to a bare picture. Re-renders in the equation's OWN
// font/size/colour. Persists the new payload only after the document change
// succeeded, so storage never describes a structure that isn't there.
// [web caveat] as for Renumber-all: a replaced raster picture may not be
// re-taggable via the selection, so click-to-edit can be lost on Word web.
async function applyEquationChange(payload, target, op) {
  let ok = false;
  if (op === 'none') {
    ok = true;
  } else if (op === 'recell') {
    ok = await rewriteNumberCell(payload.uuid, target.numberStyle, target.number);
  } else {
    await ensureMathJax(payload.font);
    const renderLatex = decorateLatexForNumber(payload.latex, target.number, target.numberStyle);
    const { svg, descentPt, widthPt, heightPt } =
      await buildSvgForInsert(renderLatex, payload.font, payload.sizePt, payload.displayMode, payload.color);
    const svgString = embedSourceInSvg(svg, target);
    if (op === 'wrap') ok = await wrapPictureIntoTable(target, { widthPt, heightPt }, svgString);
    else if (op === 'unwrap') ok = await unwrapPictureFromTable(target, { widthPt, heightPt, descentPt }, svgString);
    else ok = await replacePictureInPlace(target, { widthPt, heightPt, descentPt }, svgString);
  }
  if (!ok) return false;
  try {
    await writeXmlPart(target);
  } catch (e) {
    notifyResult(describeInsertError('store', e));
  }
  return true;
}

// Replace an equation's picture where it stands (bare or in a cell).
async function replacePictureInPlace(eq, { widthPt, heightPt, descentPt }, svgString) {
  if (!(await selectExistingPicture(eq.uuid))) return false;
  const err = await insertPictureAtSelection(svgString, widthPt, heightPt);
  if (err) { notifyResult(err); return false; }
  await tagSelectedPicture(eq.uuid, eq.latex, descentPt, eq.displayMode, insertMode === 'svg');
  return true;
}

// Rewrite the number cell next to a table-style equation: static "(n)" text or
// a SEQ field, per `style`. Returns false if the picture or its table is gone.
async function rewriteNumberCell(uuid, style, number) {
  return Word.run(async (context) => {
    const hit = (await collectTaggedPictures(context)).find((t) => t.uuid === uuid);
    if (!hit) return false;
    const cell = hit.pic.parentTableCellOrNullObject;
    const table = hit.pic.parentTableOrNullObject;
    cell.load('rowIndex,cellIndex');
    await context.sync();
    if (cell.isNullObject || table.isNullObject) return false;
    fillNumberCellOf(table.getCell(cell.rowIndex, cell.cellIndex + 1), style, number);
    await context.sync();
    return true;
  });
}

// Bare picture → borderless 1×2 numbering table. The old picture is deleted and
// its paragraph becomes the anchor: the table goes right after it (that is what
// insertTableEquation* do with the selection's paragraph), and the anchor is
// removed afterwards if it held nothing but the equation. On a placement failure
// the old picture is already gone — insertTableEquation* re-throw after cleaning
// up their empty table, and the caller reports the hard error; the source stays
// recoverable from storage (the equation list can re-insert it).
async function wrapPictureIntoTable(eq, { widthPt, heightPt }, svgString) {
  const anchored = await Word.run(async (context) => {
    const hit = (await collectTaggedPictures(context)).find((t) => t.uuid === eq.uuid);
    if (!hit) return false;
    const para = hit.pic.paragraph;
    hit.pic.delete();
    para.getRange(Word.RangeLocation.start).select();
    await context.sync();
    return true;
  });
  if (!anchored) return false;

  const { uuid, latex, number, numberStyle: style } = eq;
  const skipPt = await resolveDisplaySkipPt();
  const align = displayAlign();
  if (insertMode === 'svg') {
    await insertTableEquationSvg({ uuid, latex, number, style, widthPt, heightPt, skipPt, align }, svgString);
  } else {
    const { base64, err } = await rasterizeForInsert(svgString, widthPt, heightPt);
    if (err) throw new Error(err.msg);
    await insertTableEquationPng({ uuid, latex, number, style, skipPt, align }, base64);
  }

  // Drop the anchor paragraph if the equation was all it held. SOFT: at worst an
  // empty line stays above the table.
  try {
    await Word.run(async (context) => {
      const hit = (await collectTaggedPictures(context)).find((t) => t.uuid === uuid);
      if (!hit) return;
      const table = hit.pic.parentTableOrNullObject;
      const before = table.getParagraphBeforeOrNullObject();
      before.load('text');
      await context.sync();
      if (table.isNullObject || before.isNullObject) return;
      if (isParagraphTextEmpty(before.text)) {
        before.delete();
        await context.sync();
      }
    });
  } catch (e) {
    console.warn('Could not remove the empty paragraph above the numbering table:', e.message);
  }
  return true;
}

// 1×2 numbering table → bare (centred) picture on its own paragraph after the
// table; the table is deleted only once the new picture is placed and tagged, so
// a failed insert leaves the original untouched. Refuses to touch a table that
// is not the 1×2 shape we create (the user may have moved the equation into
// their own table) — then the picture is replaced in place instead.
async function unwrapPictureFromTable(eq, { widthPt, heightPt, descentPt }, svgString) {
  const skipPt = await resolveDisplaySkipPt();
  const prepared = await Word.run(async (context) => {
    const hit = (await collectTaggedPictures(context)).find((t) => t.uuid === eq.uuid);
    if (!hit) return 'missing';
    const table = hit.pic.parentTableOrNullObject;
    table.load('rowCount');
    await context.sync();
    if (table.isNullObject) return 'notable';
    const firstRow = table.rows.getFirst();
    firstRow.load('cellCount');
    await context.sync();
    if (table.rowCount !== 1 || firstRow.cellCount !== 2) return 'foreign';
    const para = table.insertParagraph('', Word.InsertLocation.after);
    alignDisplayParagraph(para, displayAlign());
    para.spaceBefore = skipPt;
    para.spaceAfter = skipPt;
    para.getRange(Word.RangeLocation.start).select();
    await context.sync();
    return 'ok';
  });
  if (prepared === 'missing') return false;
  if (prepared !== 'ok') {
    notify('warn', 'The equation is not in a numbering table this add-in created, so it was updated in place.');
    return replacePictureInPlace(eq, { widthPt, heightPt, descentPt }, svgString);
  }

  const err = await insertPictureAtSelection(svgString, widthPt, heightPt);
  if (err) { notifyResult(err); return false; }
  const tagged = await tagSelectedPicture(eq.uuid, eq.latex, descentPt, eq.displayMode, insertMode === 'svg');
  if (!tagged) await tagLastBodyPicture(eq.uuid, eq.latex);

  // Both pictures now carry the uuid; the old one is the one still inside a table.
  await Word.run(async (context) => {
    const hits = (await collectTaggedPictures(context)).filter((t) => t.uuid === eq.uuid);
    const tables = hits.map((h) => h.pic.parentTableOrNullObject);
    await context.sync();
    for (const t of tables) if (!t.isNullObject) t.delete();
    await context.sync();
  });
  return true;
}

async function onRenumberAll() {
  await withDocumentOperation('Renumbering…', renumberAllCore);
}

// Run a document-wide operation that drives the selection and re-renders
// equations in their own fonts: silence the selection handler, keep the edit
// state, restore the UI font afterwards. Errors surface as a hard insert error.
async function withDocumentOperation(busyMsg, fn) {
  if (!mathJaxReady || insertMode === 'none' || suppressSelection) return;
  clearNotice();
  setMathJaxBusy(true, busyMsg);
  suppressSelection = true;
  const restoreEditing = editingUuid;
  try {
    await fn();
  } catch (e) {
    notifyResult(describeInsertError('insert', e));
  } finally {
    suppressSelection = false;
    editingUuid = restoreEditing;
    // Operations re-render equations in their own fonts, so restore the UI font
    // for the live preview before re-enabling the controls.
    try { await ensureMathJax($('font').value); } catch { /* leave as loaded */ }
    setMathJaxBusy(false);
    updateUiState();
    renderPreview();
  }
}

async function renumberAllCore() {
  const orderedUuids = await getTaggedPictureUuidsInOrder();
  const byUuid = await readAllEquations();
  const ordered = orderedUuids.map((uuid) => byUuid[uuid]).filter(Boolean);

  // Field equations are numbered by Word's SEQ counter — a single field update
  // refreshes them all. Everything else uses our counter + per-equation rerender.
  const hasFieldEqs = ordered.some((p) => p.numbered && numberPlacement(p.numberStyle) === 'field');
  const staticEqs = ordered
    .filter((p) => numberPlacement(p.numberStyle) !== 'field')
    .map((p) => ({ uuid: p.uuid, numbered: p.numbered, number: p.number }));

  const changed = assignNumbers(staticEqs);
  let done = 0;
  for (const { uuid, number } of changed) {
    const payload = byUuid[uuid];
    // Dispatch by the equation's OWN style: table equations just get their
    // number cell rewritten (cheap); inline equations re-render the picture.
    const ok = numberPlacement(payload.numberStyle) === 'cell'
      ? await renumberTableCell(payload, number)
      : await rerenderEquationWithNumber(payload, number);
    if (ok) done += 1;
  }
  if (hasFieldEqs) await updateAllSeqFields();

  // The counter continues after the static numbered equations (field equations
  // are counted by Word, not by us).
  const staticNumbered = staticEqs.filter((e) => e.numbered).length;
  numberingState = { ...numberingState, nextNumber: staticNumbered + 1 };
  await writeNumberingState(numberingState);
  if (!noticeFromAppFlow) {
    notify('info',
      done ? `Renumbered ${done} equation${done === 1 ? '' : 's'}.`
        : hasFieldEqs ? 'Updated equation fields.'
          : 'Equation numbers are already in order.');
  }
}

function onNewEquation() {
  editingUuid = null;
  editingNumber = 0;
  editingStyle = numberingState.style;
  $('latex').value = '';
  updateLineNumbers();
  $('mode-inline').checked = true;
  $('size-mode').value = 'selection';
  $('size-pt').value = 11;
  $('size-pt').disabled = true;
  $('number-this').checked = false;
  updateNumberingUi();
  $('preview').innerHTML = '';
  $('preview').classList.remove('preview--inline');
  showRenderError(null);
  updateSettingsSummary();
  updateUiState();
  $('latex').focus();
}

// ---- Selection change handler: click an equation to load it for editing ----

function registerSelectionHandler() {
  Office.context.document.addHandlerAsync(
    Office.EventType.DocumentSelectionChanged,
    onSelectionChanged,
    (res) => {
      if (res.status === Office.AsyncResultStatus.Failed) {
        console.warn('Could not register selection handler:', res.error.message);
      }
    },
  );
}

// Fires on every cursor move. Keep cheap: only fetch altTextTitle.
async function onSelectionChanged() {
  // Renumber-all selects each picture as it replaces it; ignore those moves so we
  // don't load every equation into the editor mid-renumber.
  if (suppressSelection) return;
  if (selectionCheckPending) return;
  selectionCheckPending = true;
  try {
    const hit = await Word.run(async (context) => {
      const sel = context.document.getSelection();
      const pics = sel.inlinePictures;
      pics.load('items/altTextTitle,items/altTextDescription');
      await context.sync();
      for (const p of pics.items) {
        const t = p.altTextTitle || '';
        if (t.startsWith(ALT_PREFIX)) {
          return { uuid: t.slice(ALT_PREFIX.length), altDesc: p.altTextDescription || '' };
        }
      }
      return null;
    });

    // No tagged equation in selection: keep current edit state. Don't clear
    // — user may be clicking around to position cursor before re-inserting.
    if (!hit || hit.uuid === editingUuid) return;
    const { uuid, altDesc } = hit;

    let data = await readXmlPartByUuid(uuid);
    if (!data) {
      // Tagged picture but no XML part (pasted from another doc, or the part was
      // stripped). Recover the LaTeX from the alt text and self-heal a fresh part.
      data = await recoverFromAltText(uuid, altDesc);
      if (!data) return;   // alt text isn't a usable equation source — don't guess
    }

    editingUuid = uuid;
    editingNumber = data.number || 0;
    // A numbered equation keeps its own style; an unnumbered one would take the
    // DOCUMENT style if the user numbers it now.
    editingStyle = data.numbered ? (data.numberStyle || 'inline') : numberingState.style;
    $('latex').value = data.latex;
    updateLineNumbers();
    if (data.displayMode === 'display') $('mode-display').checked = true;
    else                                 $('mode-inline').checked  = true;
    $('number-this').checked = !!data.numbered;
    updateNumberingUi();
    $('size-mode').value = 'fixed';
    $('size-pt').value = data.sizePt;
    $('size-pt').disabled = false;
    if (data.color) $('color').value = data.color;

    // Restore the equation's own font so the preview matches it and a
    // subsequent "Update equation" doesn't silently re-render in whatever font
    // happened to be loaded. Only fonts we still offer are restored; an unknown
    // font (older/foreign doc) leaves the dropdown as-is.
    if (data.font && FONT_LABELS[data.font] && data.font !== loadedFont) {
      $('font').value = data.font;
      setMathJaxBusy(true, `Loading ${FONT_LABELS[data.font]} font…`);
      try {
        await ensureMathJax(data.font);
      } catch (e) {
        // Font fetch failed: degrade gracefully instead of throwing out of the
        // handler (which would skip the UI updates below and leave the pane
        // mislabeled). renderPreview() no-ops while MathJax isn't ready.
        notify('warn', `Could not load the ${FONT_LABELS[data.font]} font; preview may be unavailable.`);
      } finally {
        setMathJaxBusy(false);
      }
    }

    updateSettingsSummary();
    updateUiState();
    await renderPreview();
    refreshEquationListIfOpen();
  } finally {
    selectionCheckPending = false;
  }
}

// Recover an equation's source from its picture alt text when the primary XML
// part is gone (cross-document paste, or a stripped part). The alt text is the
// human-readable LaTeX we wrote on insert; if it still looks like an equation we
// rebuild a degraded payload (font/size/display aren't stored in alt text) and
// self-heal a fresh XML part so a second click loads via the fast path. Returns
// the payload, or null if the alt text isn't a usable source.
async function recoverFromAltText(uuid, altDesc) {
  const latex = (altDesc || '').trim();
  if (!looksLikeEquation(latex)) return null;
  let bodyPt = 11;
  try { ({ bodyPt } = await resolveSizes()); } catch { /* default body size */ }
  const payload = buildDegradedPayload({ uuid, latex, sizePt: bodyPt });
  try {
    await writeXmlPart(payload);   // self-heal so the next click takes the fast path
  } catch (e) {
    notifyResult(describeInsertError('store', e));  // still usable for this edit
  }
  notify('info',
    'Recovered this equation’s source from its alt text. Font, size, and display ' +
    'mode were reset to defaults — adjust and Update if needed.');
  return payload;
}
