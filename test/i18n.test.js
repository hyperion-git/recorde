import { test } from 'node:test';
import assert from 'node:assert/strict';

import { STRINGS, pickLocale, strings } from '../src/i18n.js';

test('pickLocale matches by primary subtag, falls back to en', () => {
  assert.equal(pickLocale('de-DE'), 'de');
  assert.equal(pickLocale('de'), 'de');
  assert.equal(pickLocale('en-US'), 'en');
  assert.equal(pickLocale('fr'), 'en');        // unsupported → English
  assert.equal(pickLocale(undefined), 'en');
  assert.equal(pickLocale(''), 'en');
  assert.equal(pickLocale('DE-de'), 'de');     // case-insensitive
});

test('strings() returns the matched table', () => {
  assert.equal(strings('de-DE').insert, 'Gleichung einfügen');
  assert.equal(strings('en').insert, 'Insert equation');
  assert.equal(strings('xx').statusNew, 'New equation');
});

test('every locale table has the exact same keys as en (no missing translations)', () => {
  const enKeys = Object.keys(STRINGS.en).sort();
  for (const [lang, table] of Object.entries(STRINGS)) {
    assert.deepEqual(Object.keys(table).sort(), enKeys, `locale "${lang}" key mismatch`);
  }
});

test('the editing template carries the {id} placeholder in every locale', () => {
  for (const [lang, table] of Object.entries(STRINGS)) {
    assert.ok(table.editing.includes('{id}'), `locale "${lang}" editing missing {id}`);
  }
});
