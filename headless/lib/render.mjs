// Headless equation rendering — the add-in's buildSvgForInsert without a browser.
// Same MathJax packages, same font packages, same post-processing helpers from
// core/mathsvg.js (scale to pt from the viewBox, inline currentColor, descent).
import { Worker } from 'node:worker_threads';
import { DOMParser } from 'linkedom';
import { scaleSvgToPt, inlineColors, descentPtFromViewBox } from '../../core/mathsvg.js';
import { KNOWN_FONTS } from '../../core/settings.js';

const workers = new Map();     // font → { worker, pending: Map<id, {resolve}>, ready }
let nextId = 1;

function workerFor(font) {
  if (!KNOWN_FONTS.includes(font)) throw new Error(`unknown font "${font}" (known: ${KNOWN_FONTS.join(', ')})`);
  if (workers.has(font)) return workers.get(font);
  const worker = new Worker(new URL('./render-worker.mjs', import.meta.url), { workerData: { font } });
  const entry = { worker, pending: new Map(), ready: null };
  entry.ready = new Promise((resolve, reject) => {
    worker.once('message', (m) => (m.ready ? resolve() : reject(new Error('render worker failed to start'))));
    worker.once('error', reject);
  });
  worker.on('message', (m) => {
    const p = entry.pending.get(m.id);
    if (p) { entry.pending.delete(m.id); p.resolve(m); }
    if (!entry.pending.size) worker.unref();   // idle workers never keep the CLI alive
  });
  worker.on('error', (e) => { for (const p of entry.pending.values()) p.resolve({ svg: '', error: e.message }); entry.pending.clear(); });
  workers.set(font, entry);
  return entry;
}

// → { svg, widthPt, heightPt, descentPt, error }. `error` is MathJax's TeX
// error text (the SVG then contains an <merror>), null when clean.
export async function renderEquation({ latex, font = 'tex', sizePt = 11, displayMode = 'inline', color = '#000000', preamble = '' }) {
  const entry = workerFor(font);
  await entry.ready;
  const id = nextId++;
  const raw = await new Promise((resolve) => {
    entry.worker.ref();                        // a pending render keeps the loop alive
    entry.pending.set(id, { resolve });
    entry.worker.postMessage({ id, latex, display: displayMode === 'display', preamble });
  });
  if (!raw.svg) return { svg: '', widthPt: 0, heightPt: 0, descentPt: 0, error: raw.error || 'no SVG produced' };
  const doc = new DOMParser().parseFromString(raw.svg, 'image/svg+xml');
  const svgEl = doc.documentElement;
  const descentPt = displayMode === 'inline' ? descentPtFromViewBox(svgEl.getAttribute('viewBox'), sizePt) : 0;
  scaleSvgToPt(svgEl, sizePt);
  inlineColors(svgEl, color);
  if (!svgEl.getAttribute('xmlns')) svgEl.setAttribute('xmlns', 'http://www.w3.org/2000/svg');
  const widthPt = parseFloat(svgEl.getAttribute('width')) || sizePt;
  const heightPt = parseFloat(svgEl.getAttribute('height')) || sizePt;
  return { svg: svgEl.toString(), widthPt, heightPt, descentPt, error: raw.error };
}

export async function closeRenderers() {
  for (const { worker } of workers.values()) await worker.terminate();
  workers.clear();
}
