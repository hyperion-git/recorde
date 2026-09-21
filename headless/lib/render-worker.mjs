// One MathJax instance per worker = one font per worker. MathJax's component
// registry is process-global, so a second init() in the same thread cannot
// switch fonts; worker threads have their own module registry. Mirrors the
// pane's loader (taskpane.js loadMathJax) so headless renders match the add-in.
import { parentPort, workerData } from 'node:worker_threads';
import { DEFAULT_PREAMBLE } from '../../core/preamble.js';
import { flattenLatex, mathJaxFontName } from '../../core/mathsvg.js';

// MathJax's component loader logs "No version information…" for every [tex]
// extension in Node; nothing actionable, so keep the CLI output clean.
const warn = console.warn;
console.warn = (...a) => { if (!/No version information|Invalid option/.test(String(a[0]))) warn(...a); };

const PACKAGES = ['ams', 'newcommand', 'mathtools', 'boldsymbol', 'color', 'cancel', 'physics', 'braket', 'autoload'];
const MathJax = (await import('mathjax')).default;
await MathJax.init({
  loader: { load: ['input/tex', 'output/svg', ...PACKAGES.map((p) => `[tex]/${p}`)] },
  tex: { packages: { '[+]': PACKAGES } },
  svg: { fontCache: 'none' },
  output: { font: mathJaxFontName(workerData.font), linebreaks: { inline: false } },
  startup: { typeset: false },
});
// Register the built-in macro preamble once (raw, like the pane).
try { await MathJax.tex2svgPromise(DEFAULT_PREAMBLE); } catch { /* non-fatal */ }

parentPort.on('message', async ({ id, latex, display, preamble }) => {
  try {
    const tex = flattenLatex(preamble ? preamble + '\n' + latex : latex);
    const node = await MathJax.tex2svgPromise(tex, { display: !!display });
    const html = MathJax.startup.adaptor.outerHTML(node);
    const svg = (/<svg[\s\S]*<\/svg>/.exec(html) || [''])[0];
    const err = /data-mjx-error="([^"]*)"/.exec(html);
    parentPort.postMessage({ id, svg, error: err ? err[1].replace(/&quot;/g, '"').replace(/&amp;/g, '&') : null });
  } catch (e) {
    parentPort.postMessage({ id, svg: '', error: e.message });
  }
});
parentPort.postMessage({ ready: true });
