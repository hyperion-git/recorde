// Host-independent sizing math for rasterizing an SVG equation to a PNG.
//
// Word-on-the-web's insertInlinePictureFromBase64 (ImageCoercion 1.1) can't take
// SVG, so the desktop SVG path degrades to a raster PNG on the web. The actual
// canvas draw + Blob → base64 conversion needs the DOM and lives in taskpane.js;
// everything here is pure number/string math so it's testable under `node --test`.
// Keep this free of any Office/Word/MathJax/canvas/DOM references.

// CSS reference pixels per point: 96 CSS px per inch, 72 pt per inch.
export const PX_PER_PT = 96 / 72;

// Target print resolution for the raster (web) insert path. The image is placed
// at its point size, so effective DPI = 96 × scale; scale = TARGET_DPI/96 = 3.125
// puts it at 300 DPI regardless of the screen's devicePixelRatio. Vector SVG
// (desktop) is resolution-independent and ignores this.
export const TARGET_DPI = 300;

// Cap on an inserted image's base64 length. Word's coercion path chokes on very
// large payloads; this is a sane ceiling well below where things break, and lets
// the caller surface a friendly error instead of a host failure.
export const MAX_BASE64_LENGTH = 4_000_000;

export function ptToPx(pt) {
  return pt * PX_PER_PT;
}

// Render scale (≈ devicePixelRatio): supersample so the PNG stays crisp on HiDPI
// displays. Clamped to [2,4] — below 2 looks soft, above 4 wastes the pixel
// budget. Non-finite / too-small inputs default to the 2 floor.
export function clampScale(s) {
  const n = Number(s);
  if (!Number.isFinite(n) || n < 2) return 2;
  if (n > 4) return 4;
  return n;
}

// Pick a supersample scale and pixel dimensions for the PNG. The base scale is
// the larger of the 300-DPI print target (TARGET_DPI/96) and the screen's
// devicePixelRatio (clamped) — so the raster is print-quality on any display and
// extra-crisp on HiDPI ones. It then backs the scale off if the area would exceed
// maxPixels — protecting against multi-megapixel canvases for large display
// equations. The pixel budget wins over DPI, so a very large equation can fall
// below 300 DPI (and even below 1×); pxW/pxH never drop below 1 (a 0×0 canvas is
// invalid), even for sub-point inputs.
export function computeRasterSize(wPt, hPt, dpr, maxPixels = 4_000_000) {
  // A near-zero floor keeps the loop terminating without forbidding the
  // sub-1× scale a large equation may need to fit the budget.
  const FLOOR = 0.01;
  let scale = Math.max(TARGET_DPI / 96, clampScale(dpr));
  const dims = (s) => ({
    pxW: Math.max(1, Math.round(ptToPx(wPt) * s)),
    pxH: Math.max(1, Math.round(ptToPx(hPt) * s)),
  });

  let { pxW, pxH } = dims(scale);
  if (pxW * pxH > maxPixels) {
    // Solve for the largest scale whose (un-rounded) area fits the budget:
    //   ptToPx(wPt)*ptToPx(hPt) * scale^2 <= maxPixels.
    // Rounding to integer px can re-inflate the area by a sub-pixel sliver, so
    // step down until the rounded area actually fits, never below the floor.
    const areaPerScaleSq = ptToPx(wPt) * ptToPx(hPt);
    scale = Math.max(FLOOR, Math.sqrt(maxPixels / areaPerScaleSq));
    ({ pxW, pxH } = dims(scale));
    while (pxW * pxH > maxPixels && scale > FLOOR) {
      scale = Math.max(FLOOR, scale - 0.001);
      ({ pxW, pxH } = dims(scale));
    }
  }
  return { scale, pxW, pxH };
}

// A canvas.toDataURL() result is 'data:image/png;base64,<payload>'. The host
// insert API wants just the payload; strip the prefix if present, else assume the
// string is already raw base64. The base64 alphabet (+ / =) is left untouched.
export function stripDataUri(s) {
  const m = String(s).match(/^data:[^;,]*;base64,(.*)$/s);
  return m ? m[1] : s;
}

export function isPayloadTooLarge(base64) {
  return base64.length > MAX_BASE64_LENGTH;
}
