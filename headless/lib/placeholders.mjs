// Placeholder grammar — the agent contract (headless roadmap H1.3). Pure.
//
//   [[math: \hbar\omega]]            inline equation in the running text
//   [[display: \int_0^\infty …]]     display equation on its own centred paragraph
//   [[eq: i\hbar\partial_t\psi=H\psi]]  numbered display equation
//   [[eq#label: …]]                  numbered + label (referenceable)
//   [[ref: label]]                   the label's number, e.g. "(3)"
//   [[eq#label{font=stix2,color=#1f77b4,size=12,style=field,align=center}: …]]
//                                    per-equation options (any kind but ref):
//                                    font, color (#rrggbb), size (pt), style
//                                    (inline|table|field, numbered only),
//                                    align (left|center, display/eq only)
//
// `(?!\])` after the closing brackets keeps `f(x) = [a,b]]]` intact: the match
// ends at the FIRST `]]` that is not followed by another `]`. No nesting.
// U+FFFC (object replacement) marks non-text run content (tabs, breaks,
// drawings) in the joined paragraph text, so a placeholder can never span one.

export const KINDS = ['math', 'display', 'eq', 'ref'];

const CORE = String.raw`\[\[(math|display|eq|ref)(?:#([\w:.-]+))?(?:\{([^{}]*)\})?:[ \t]*([^￼]*?)\]\](?!\])`;
export const OPTION_KEYS = ['font', 'color', 'size', 'style', 'align'];
export const PLACEHOLDER_RE = new RegExp(CORE, 'g');
// Opt-in $…$ / $$…$$ for markdown/pandoc authors. Off by default: currency.
const DOLLAR_RE = /\$\$([^￼$]+?)\$\$|(?<![\\$\w])\$(?!\s)([^￼$\n]+?)(?<!\s)\$(?![\w$])/g;

// Find every placeholder in `text`. Returns [{ start, end, kind, label, latex }]
// in text order. With `dollar`, $$…$$ → display and $…$ → math are found too.
export function findPlaceholders(text, { dollar = false } = {}) {
  const out = [];
  const s = String(text);
  for (const m of s.matchAll(PLACEHOLDER_RE)) {
    const kind = m[1];
    const label = m[2] || null;
    const options = parseOptions(m[3]);
    const latex = m[4].trim();
    out.push({ start: m.index, end: m.index + m[0].length, kind, label, options, latex });
  }
  if (dollar) {
    for (const m of s.matchAll(DOLLAR_RE)) {
      const start = m.index, end = m.index + m[0].length;
      if (out.some((p) => start < p.end && end > p.start)) continue; // inside a [[…]]
      out.push({
        start, end,
        kind: m[1] != null ? 'display' : 'math',
        label: null,
        options: {},
        latex: (m[1] ?? m[2]).trim(),
      });
    }
    out.sort((a, b) => a.start - b.start);
  }
  return out;
}

// "font=stix2,color=#ff0000" → { font, color }. Unknown keys are kept so the
// validator can name them.
function parseOptions(src) {
  const o = {};
  if (!src) return o;
  for (const part of src.split(',')) {
    const [k, ...rest] = part.split('=');
    const key = k.trim();
    if (key) o[key] = rest.join('=').trim();
  }
  return o;
}

// Validate a placeholder list for one paragraph; returns an error string or null.
export function validatePlaceholders(list) {
  for (const p of list) {
    if (p.kind === 'ref') {
      if (!p.latex || p.label) return `[[ref: …]] takes a label, e.g. [[ref: schrodinger]] (got "${p.latex}")`;
      if (Object.keys(p.options || {}).length) return '[[ref: …]] takes no options';
      continue;
    }
    if (!p.latex) return `empty ${p.kind} placeholder`;
    for (const [k, v] of Object.entries(p.options || {})) {
      if (!OPTION_KEYS.includes(k)) return `unknown option "${k}" (allowed: ${OPTION_KEYS.join(', ')})`;
      if (!v) return `option "${k}" needs a value`;
      if (k === 'style' && p.kind !== 'eq') return 'option "style" applies to numbered [[eq: …]] only';
      if (k === 'style' && !['inline', 'table', 'field'].includes(v)) return `unknown style "${v}" (inline|table|field)`;
      if (k === 'align' && p.kind === 'math') return 'option "align" applies to display equations only';
      if (k === 'align' && !['left', 'center'].includes(v)) return `unknown align "${v}" (left|center)`;
      if (k === 'color' && !/^#[0-9a-fA-F]{6}$/.test(v)) return `color must be #rrggbb (got "${v}")`;
      if (k === 'size' && !(Number(v) >= 4 && Number(v) <= 96)) return `size must be 4…96 pt (got "${v}")`;
    }
  }
  return null;
}
