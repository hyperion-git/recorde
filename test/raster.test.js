import { test } from 'node:test';
import assert from 'node:assert/strict';

import {
  PX_PER_PT, MAX_BASE64_LENGTH, TARGET_DPI, ptToPx, clampScale,
  computeRasterSize, stripDataUri, isPayloadTooLarge,
} from '../core/raster.js';

test('ptToPx converts pt → CSS px (72pt = 96px)', () => {
  assert.equal(ptToPx(72), 96);
  assert.equal(ptToPx(11), 11 * PX_PER_PT);
});

test('clampScale clamps to [2,4] and defaults junk to the 2 floor', () => {
  // WHY: scale drives supersampling; out-of-range / non-finite dpr values must
  // degrade to a safe crisp-but-bounded default rather than a 0× or 100× canvas.
  assert.equal(clampScale(1), 2);
  assert.equal(clampScale(3), 3);
  assert.equal(clampScale(5), 4);
  assert.equal(clampScale(undefined), 2);
  assert.equal(clampScale(0), 2);
  assert.equal(clampScale(NaN), 2);
});

test('computeRasterSize: in-budget size renders at the 300-DPI target scale', () => {
  // dpr=2 is below the print target, so the 300-DPI floor (TARGET_DPI/96) wins.
  const r = computeRasterSize(33, 11, 2);
  const expected = TARGET_DPI / 96;
  assert.equal(r.scale, expected);
  assert.equal(r.pxW, Math.round(ptToPx(33) * expected));
  assert.equal(r.pxH, Math.round(ptToPx(11) * expected));
});

test('computeRasterSize: effective resolution is 300 DPI (72pt = 1in → 300px)', () => {
  // The image is inserted at its point size, so DPI = pxW / (wPt/72). A 72pt
  // (1-inch) equation must rasterize to 300px to print at 300 DPI.
  const r = computeRasterSize(72, 72, 1);
  assert.equal(r.pxW, 300);
  assert.equal(Math.round(r.pxW / (72 / 72)), TARGET_DPI);
});

test('computeRasterSize: sub-point inputs still yield a valid ≥1×1 canvas', () => {
  // WHY: a 0×0 canvas is invalid; tiny equations must round up to 1px.
  const r = computeRasterSize(0.1, 0.1, 2);
  assert.ok(r.pxW >= 1 && r.pxH >= 1);
});

test('computeRasterSize: oversized area backs the scale off under maxPixels', () => {
  // WHY: a large display equation at 4× could blow past the pixel budget; the
  // scale must shrink until the area fits (with a small rounding tolerance).
  const r = computeRasterSize(960, 960, 4, 1_600_000);
  assert.ok(r.scale < 4);
  assert.ok(r.pxW * r.pxH <= 1_600_000 + 2);
});

test('computeRasterSize: small equation keeps full 4× under the default budget', () => {
  const r = computeRasterSize(33, 11, 4, 4_000_000);
  assert.equal(r.scale, 4);
});

test('stripDataUri returns the base64 payload, leaving bare strings alone', () => {
  assert.equal(stripDataUri('data:image/png;base64,AAAB'), 'AAAB');
  assert.equal(stripDataUri('AAAB'), 'AAAB');
  // WHY: the base64 alphabet (+ / =) must survive untouched.
  assert.equal(stripDataUri('AAAB+/=='), 'AAAB+/==');
});

test('isPayloadTooLarge flags only payloads over the cap', () => {
  assert.equal(isPayloadTooLarge('x'.repeat(10)), false);
  assert.equal(isPayloadTooLarge('x'.repeat(MAX_BASE64_LENGTH + 1)), true);
});
