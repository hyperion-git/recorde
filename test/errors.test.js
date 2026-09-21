import { test } from 'node:test';
import assert from 'node:assert/strict';

import { isLikelyOffline, describeInsertError, pictureVanishedNotice } from '../addin/src/errors.js';

// isLikelyOffline reads globalThis.navigator?.onLine. Set & restore it around the
// one test that depends on it so the rest of the suite sees the real (or absent)
// navigator and exercises the message-signature fallback path. Node defines
// `navigator` as a getter-only accessor, so we swap the property descriptor
// rather than assign, and restore the original in finally.
function withNavigator(nav, fn) {
  const prev = Object.getOwnPropertyDescriptor(globalThis, 'navigator');
  try {
    Object.defineProperty(globalThis, 'navigator', { value: nav, configurable: true });
    fn();
  } finally {
    if (prev) Object.defineProperty(globalThis, 'navigator', prev);
    else delete globalThis.navigator;
  }
}

test('AccessDenied → error, read-only/protected', () => {
  const r = describeInsertError('insert', { code: 'AccessDenied', message: 'x' });
  assert.equal(r.level, 'error');
  assert.match(r.msg, /read-only|protected/i);
});

test('ApiNotAvailable → error, needs Windows or Mac desktop', () => {
  const r = describeInsertError('insert', { code: 'ApiNotAvailable' });
  assert.equal(r.level, 'error');
  assert.match(r.msg, /Windows or Mac|desktop/i);
});

test('store step generic failure → warn, about the saved source', () => {
  // WHY: the equation is already inserted; only the click-to-edit source is lost,
  // which is recoverable — so this must not be surfaced as a hard error.
  const r = describeInsertError('store', { code: 'GeneralException', message: 'x' });
  assert.equal(r.level, 'warn');
  assert.match(r.msg, /source|click-to-edit|save/i);
});

test('tag step generic failure → warn, about click-to-edit', () => {
  const r = describeInsertError('tag', { code: 'GeneralException' });
  assert.equal(r.level, 'warn');
  assert.match(r.msg, /click-to-edit/i);
});

test('RequestPayloadSizeLimitExceeded → too large / try smaller', () => {
  const r = describeInsertError('insert', { code: 'RequestPayloadSizeLimitExceeded' });
  assert.match(r.msg, /too large|smaller/i);
});

test('unknown code on insert → error, actionable, includes raw message', () => {
  const r = describeInsertError('insert', { code: 'SomethingNew', message: 'boom' });
  assert.equal(r.level, 'error');
  assert.match(r.msg, /try again/i);
  assert.ok(r.msg.includes('boom'), 'expected raw message to be appended');
  assert.match(r.msg, /\(boom\)/);
});

test('unknown code appends the raw detail in parentheses', () => {
  const r = describeInsertError('insert', { code: 'GeneralException', message: 'the detail' });
  assert.match(r.msg, /\(the detail\)/);
});

test('never throws on a string / empty-object / undefined error', () => {
  for (const e of ['oops', {}, undefined]) {
    const r = describeInsertError('insert', e);
    assert.equal(typeof r.msg, 'string');
  }
});

test('isLikelyOffline: navigator.onLine === false wins', () => {
  withNavigator({ onLine: false }, () => {
    assert.equal(isLikelyOffline({}), true);
  });
});

test('isLikelyOffline: network message signatures', () => {
  assert.equal(isLikelyOffline({ message: 'Failed to fetch' }), true);
  assert.equal(isLikelyOffline({ code: 'GeneralException', message: 'NET::ERR' }), true);
  assert.equal(isLikelyOffline({ code: 'AccessDenied', message: 'denied' }), false);
});

test('offline detection wins over code mapping in describeInsertError', () => {
  const r = describeInsertError('insert', { code: 'GeneralException', message: 'failed to fetch' });
  assert.match(r.msg, /offline/i);
});

test('pictureVanishedNotice → info, explains re-insert as new', () => {
  const r = pictureVanishedNotice();
  assert.equal(r.level, 'info');
  assert.match(r.msg, /no longer in the document|new equation/i);
});
