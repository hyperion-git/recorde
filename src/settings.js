// Persisted UI defaults: serialize / validate the user's font, sizing, and
// display preferences.
//
// Pure (strings/numbers only): the localStorage / Office `settings` glue lives
// in taskpane.js. Settings can arrive from two stores — Office's roaming
// settings (already deserialized to an object) or a JSON string in
// localStorage — so normalizeSettings is the single choke point that turns any
// untrusted/stale input into exactly the four known keys with valid values.
// That keeps a corrupt or version-skewed blob from ever reaching the renderer.

// Dropdown values, in display order. Anything outside this set degrades to 'tex'.
export const KNOWN_FONTS = ['tex', 'newcm', 'termes', 'stix2', 'pagella', 'asana'];

export const DEFAULT_SETTINGS = {
  font: 'termes',   // ≈ newtxmath (Times) — matches the author's paper template
  sizeMode: 'selection',
  fixedPt: 11,
  displayMode: 'inline',
  color: '#000000',
  // Display-equation alignment. 'left' = LaTeX's fleqn: flush left, indented by
  // \mathindent (25 pt); 'center' = the plain LaTeX default. Left is our default.
  displayAlign: 'left',
};

const SIZE_MODES = ['selection', 'body', 'fixed'];
const DISPLAY_MODES = ['inline', 'display'];
export const DISPLAY_ALIGNS = ['left', 'center'];
const HEX_COLOR = /^#[0-9a-fA-F]{6}$/;   // 6-digit hex from <input type="color">


// Match computeRenderSize's clamp range: a fixed pt outside [6,96] is treated as
// corrupt and reset to the default rather than producing an absurd equation.
const MIN_PT = 6, MAX_PT = 96;

// Always returns a NEW object with exactly the known keys, each valid.
// Extra keys are dropped so a persisted blob can't smuggle unexpected fields
// through to the rest of the app.
export function normalizeSettings(obj) {
  if (obj === null || typeof obj !== 'object') return { ...DEFAULT_SETTINGS };

  const font = KNOWN_FONTS.includes(obj.font) ? obj.font : DEFAULT_SETTINGS.font;
  const sizeMode = SIZE_MODES.includes(obj.sizeMode) ? obj.sizeMode : 'selection';
  const displayMode = DISPLAY_MODES.includes(obj.displayMode) ? obj.displayMode : 'inline';

  const pt = Number(obj.fixedPt);
  const fixedPt = Number.isFinite(pt) && pt >= MIN_PT && pt <= MAX_PT ? pt : 11;

  const color = HEX_COLOR.test(obj.color) ? obj.color.toLowerCase() : '#000000';
  const displayAlign = DISPLAY_ALIGNS.includes(obj.displayAlign) ? obj.displayAlign : DEFAULT_SETTINGS.displayAlign;

  return { font, sizeMode, fixedPt, displayMode, color, displayAlign };
}

export function serializeSettings(obj) {
  return JSON.stringify(normalizeSettings(obj));
}

// Accepts either a JSON string (localStorage) or an already-deserialized object
// (Office settings.get). Invalid JSON / null / empty falls back to defaults.
export function parseSettings(raw) {
  if (raw !== null && typeof raw === 'object') return normalizeSettings(raw);
  if (raw === null || raw === '') return { ...DEFAULT_SETTINGS };
  try {
    return normalizeSettings(JSON.parse(raw));
  } catch {
    return { ...DEFAULT_SETTINGS };
  }
}
