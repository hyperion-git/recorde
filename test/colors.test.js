import { test } from 'node:test';
import assert from 'node:assert/strict';

import { COLOR_SETS, findColorSet, normalizeHex } from '../addin/src/colors.js';

test('colour sets: unique ids, valid lowercase hex, no duplicate names within a row', () => {
  // WHY: every hex is assigned straight to <input type=color>, which silently
  // rejects anything but #rrggbb — a bad value would leave the old colour selected.
  const ids = COLOR_SETS.map((s) => s.id);
  assert.equal(new Set(ids).size, ids.length);
  for (const set of COLOR_SETS) {
    assert.ok(set.groups.length > 0, `${set.id}: no groups`);
    for (const g of set.groups) {
      const names = g.colors.map((k) => k.name);
      assert.equal(new Set(names).size, names.length, `${set.id}/${g.name}: duplicate names`);
      for (const k of g.colors) assert.match(k.hex, /^#[0-9a-f]{6}$/, `${set.id}/${g.name}/${k.name}: ${k.hex}`);
    }
  }
});

test('colour sets: the AFP qualitative cycle is in order and matches the skill palette', () => {
  const q = findColorSet('qualitative').groups[0].colors;
  assert.equal(q.length, 10);
  assert.deepEqual(q.slice(0, 4).map((k) => k.hex), ['#2e2cb8', '#db002b', '#1f8a70', '#fd7400']);
  assert.equal(findColorSet('shades').groups.length, 8);
  assert.equal(findColorSet('grays').groups.length, 4);
  assert.equal(findColorSet('pastel').groups.map((g) => g.colors.length).join(','), '11,11,11');
});

test('normalizeHex accepts #RRGGBB in any case and rejects the rest', () => {
  assert.equal(normalizeHex('#2E2CB8'), '#2e2cb8');
  assert.equal(normalizeHex('2e2cb8'), '#2e2cb8');
  assert.equal(normalizeHex('#fff'), null);
  assert.equal(normalizeHex(''), null);
});
