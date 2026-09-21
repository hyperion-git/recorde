// Prefab colour swatches for the equation colour picker — the AFP palette from
// the scientific-figure skill (references/color-palette.md), so equations in
// Word match the figures made with it. Pure data: sets → groups (rows) → colours.
// Hex values are lowercase 6-digit so they can be assigned to <input type=color>.

const c = (name, hex) => ({ name, hex });

export const COLOR_SETS = [
  {
    id: 'qualitative',
    groups: [
      { name: 'C0–C9', colors: [
        c('blue', '#2e2cb8'), c('red', '#db002b'), c('green', '#1f8a70'), c('orange', '#fd7400'),
        c('blueberry', '#1c809e'), c('limegreen', '#bedb43'), c('pink', '#9b2f5c'), c('gold', '#fabd1e'),
        c('violet', '#5c2d99'), c('forest', '#2c6b2f'),
      ] },
      { name: 'reserve', colors: [c('turquoise', '#004358'), c('yellow', '#ffe119')] },
    ],
  },
  {
    id: 'presentation',
    groups: [
      { name: 'C0–C9', colors: [
        c('blue', '#3e36de'), c('red', '#db002b'), c('green', '#009b4e'), c('orange', '#fd7400'),
        c('blueberry', '#1b9aaa'), c('limegreen', '#bedb43'), c('pink', '#cd5789'), c('gold', '#fabd1e'),
        c('violet', '#8a57cd'), c('forest', '#30c13f'),
      ] },
    ],
  },
  {
    id: 'semantic',
    groups: [
      { name: '', colors: [
        c('ink', '#000000'), c('cold', '#2e2cb8'), c('hot', '#db002b'), c('reference', '#919191'),
        c('growth', '#1f8a70'), c('warning', '#fd7400'), c('highlight', '#5c2d99'),
      ] },
    ],
  },
  {
    id: 'pastel',
    groups: [
      { name: 'dark', colors: [
        c('emerald', '#096637'), c('lime', '#476609'), c('yellow', '#c18711'), c('orange', '#bf5c06'),
        c('red', '#b11e13'), c('pink', '#9b2f5c'), c('violet', '#5c2d99'), c('kobalt', '#1c1896'),
        c('blue', '#005694'), c('cyan', '#005f6a'), c('gray', '#4b4b4b'),
      ] },
      { name: 'medium', colors: [
        c('emerald', '#009b4e'), c('lime', '#608e05'), c('yellow', '#ecac2d'), c('orange', '#f9943b'),
        c('red', '#e83e34'), c('pink', '#cd5789'), c('violet', '#8a57cd'), c('kobalt', '#3e36de'),
        c('blue', '#137fcc'), c('cyan', '#1b9aaa'), c('gray', '#919191'),
      ] },
      { name: 'light', colors: [
        c('emerald', '#56cd84'), c('lime', '#9cce38'), c('yellow', '#ffd265'), c('orange', '#ffb678'),
        c('red', '#fd7d7d'), c('pink', '#f891bd'), c('violet', '#c191f8'), c('kobalt', '#a9a6ff'),
        c('blue', '#b2d5fc'), c('cyan', '#89e9f6'), c('gray', '#bebebe'),
      ] },
    ],
  },
  {
    id: 'shades',
    groups: [
      { name: 'red',    colors: ['#490006', '#650008', '#96030f', '#ca0011', '#f64756', '#f9baba'].map((h, i) => c(`red ${i + 1}`, h)) },
      { name: 'orange', colors: ['#491600', '#763100', '#c65300', '#ff7715', '#f7ab6a', '#fbd0ab'].map((h, i) => c(`orange ${i + 1}`, h)) },
      { name: 'yellow', colors: ['#4f3a02', '#9b6f03', '#dfa204', '#fabd1e', '#fcd169', '#fce5ad'].map((h, i) => c(`yellow ${i + 1}`, h)) },
      { name: 'green',  colors: ['#0c2b0e', '#1d4825', '#2c6b2f', '#3a9300', '#84d700', '#c6f46f'].map((h, i) => c(`green ${i + 1}`, h)) },
      { name: 'blue',   colors: ['#16193b', '#35478c', '#4e7ac7', '#7fb2f0', '#add5f7', '#cce5fa'].map((h, i) => c(`blue ${i + 1}`, h)) },
      { name: 'purple', colors: ['#25064d', '#36175e', '#553285', '#7b52ab', '#9768d1', '#bb92ef'].map((h, i) => c(`purple ${i + 1}`, h)) },
      { name: 'pink',   colors: ['#351a23', '#532131', '#84274f', '#cc559c', '#dd99cc', '#efbee1'].map((h, i) => c(`pink ${i + 1}`, h)) },
      { name: 'kobalt', colors: ['#000c59', '#1806a0', '#0d2cd3', '#425fef', '#a8d5f5', '#dbeffa'].map((h, i) => c(`kobalt ${i + 1}`, h)) },
    ],
  },
  {
    id: 'grays',
    groups: [
      { name: 'neutral', colors: ['#323232', '#464646', '#656565', '#939393', '#c1c1c1', '#eeeeee'].map((h, i) => c(`neutral ${i + 1}`, h)) },
      { name: 'warm',    colors: ['#34312e', '#46423e', '#635d58', '#938b83', '#b4aaa0', '#ede0d3'].map((h, i) => c(`warm ${i + 1}`, h)) },
      { name: 'cold',    colors: ['#2c3032', '#3d4346', '#5a6165', '#808d93', '#9aa9b1', '#cbdee8'].map((h, i) => c(`cold ${i + 1}`, h)) },
      { name: 'green',   colors: ['#2d322f', '#3e4642', '#5a655f', '#83938b', '#acc1b6', '#d3ede0'].map((h, i) => c(`green ${i + 1}`, h)) },
    ],
  },
];

export function findColorSet(id) {
  return COLOR_SETS.find((s) => s.id === id) || COLOR_SETS[0];
}

// Normalise a colour for comparison with <input type=color> values.
export function normalizeHex(hex) {
  const m = /^#?([0-9a-f]{6})$/i.exec(String(hex || '').trim());
  return m ? '#' + m[1].toLowerCase() : null;
}
