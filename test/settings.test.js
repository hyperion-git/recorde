import { test } from 'node:test';
import assert from 'node:assert/strict';

import {
  KNOWN_FONTS, DEFAULT_SETTINGS,
  normalizeSettings, serializeSettings, parseSettings,
} from '../src/settings.js';

test('KNOWN_FONTS has the exact dropdown order', () => {
  assert.deepEqual(KNOWN_FONTS, ['tex', 'newcm', 'termes', 'stix2', 'pagella', 'asana']);
});

test('DEFAULT_SETTINGS is the documented baseline', () => {
  assert.deepEqual(DEFAULT_SETTINGS, {
    font: 'termes', sizeMode: 'selection', fixedPt: 11, displayMode: 'inline', color: '#000000', displayAlign: 'left',
  });
});

test('parseSettings(null) falls back to defaults', () => {
  // WHY: a fresh install / cleared store reads null; the app must boot with a
  // valid, complete settings object rather than undefined fields.
  assert.deepEqual(parseSettings(null), DEFAULT_SETTINGS);
});

test('parseSettings on empty / malformed JSON falls back to defaults', () => {
  // WHY: a truncated or corrupt localStorage write must not throw on load.
  assert.deepEqual(parseSettings(''), DEFAULT_SETTINGS);
  assert.deepEqual(parseSettings('not json{'), DEFAULT_SETTINGS);
});

test('round-trip: serialize then parse preserves a fully-valid object', () => {
  const valid = { font: 'stix2', sizeMode: 'fixed', fixedPt: 14, displayMode: 'display', color: '#3366ff', displayAlign: 'center' };
  assert.deepEqual(parseSettings(serializeSettings(valid)), valid);
});

test('normalizeSettings drops an unknown font to the default', () => {
  assert.equal(normalizeSettings({ font: 'comic-sans' }).font, 'termes');
});

test('normalizeSettings drops an unknown sizeMode to selection', () => {
  assert.equal(normalizeSettings({ sizeMode: 'huge' }).sizeMode, 'selection');
});

test('normalizeSettings rejects out-of-range fixedPt', () => {
  // WHY: clamp range mirrors the renderer's [6,96]; a 999pt or 0pt blob is
  // corrupt and must reset rather than render an absurd size.
  assert.equal(normalizeSettings({ fixedPt: 999 }).fixedPt, 11);
  assert.equal(normalizeSettings({ fixedPt: 0 }).fixedPt, 11);
});

test('normalizeSettings accepts the inclusive fixedPt bounds', () => {
  assert.equal(normalizeSettings({ fixedPt: 96 }).fixedPt, 96);
  assert.equal(normalizeSettings({ fixedPt: 6 }).fixedPt, 6);
});

test('normalizeSettings coerces a numeric string fixedPt', () => {
  // WHY: a number <input> can hand back a string; coercion keeps it valid.
  assert.equal(normalizeSettings({ fixedPt: '14' }).fixedPt, 14);
});

test('normalizeSettings validates the color (6-digit hex lowercased, junk → black)', () => {
  // WHY: <input type=color> always yields #rrggbb; a stale/foreign blob with a
  // named color, 3-digit hex, or non-string must reset rather than reach the SVG.
  assert.equal(normalizeSettings({ color: '#AB12EF' }).color, '#ab12ef');
  assert.equal(normalizeSettings({ color: 'red' }).color, '#000000');
  assert.equal(normalizeSettings({ color: '#fff' }).color, '#000000');
  assert.equal(normalizeSettings({ color: 12345 }).color, '#000000');
});

test('normalizeSettings keeps a valid font and defaults the rest', () => {
  assert.deepEqual(normalizeSettings({ font: 'pagella' }), { ...DEFAULT_SETTINGS, font: 'pagella' });
});

test('normalizeSettings on non-objects returns defaults', () => {
  assert.deepEqual(normalizeSettings(undefined), DEFAULT_SETTINGS);
  assert.deepEqual(normalizeSettings(42), DEFAULT_SETTINGS);
  assert.deepEqual(normalizeSettings('x'), DEFAULT_SETTINGS);
});

test('parseSettings accepts an already-deserialized object (Office settings.get)', () => {
  assert.equal(parseSettings({ font: 'termes' }).font, 'termes');
});

test('serializeSettings normalizes before stringifying (bad font/pt, extra keys stripped)', () => {
  // WHY: a write must persist only the four valid keys, never round-trip
  // garbage back into the store.
  assert.deepEqual(JSON.parse(serializeSettings({ font: 'bad', fixedPt: 999, extra: 'x' })), DEFAULT_SETTINGS);
});

test('displayAlign: left by default, center accepted, anything else degrades to left', () => {
  // WHY: the alignment drives paragraph formatting on every display insert; an
  // unknown value must not leave equations unaligned or throw in Word.
  assert.equal(normalizeSettings({}).displayAlign, 'left');
  assert.equal(normalizeSettings({ displayAlign: 'center' }).displayAlign, 'center');
  assert.equal(normalizeSettings({ displayAlign: 'right' }).displayAlign, 'left');
  assert.equal(parseSettings('{"displayAlign":"center"}').displayAlign, 'center');
});
