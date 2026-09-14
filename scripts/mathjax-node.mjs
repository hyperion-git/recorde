// Node-side MathJax mirroring the pane's loader/packages (taskpane.js
// loadMathJax), for tests that must prove a LaTeX string renders without an
// <merror>. Shared by test/examples.test.js and test/testpage.test.js.
export async function initMathJaxLikePane() {
  const MathJax = (await import('mathjax')).default;
  await MathJax.init({
    loader: { load: ['input/tex', 'output/svg', '[tex]/ams', '[tex]/newcommand', '[tex]/mathtools',
                     '[tex]/boldsymbol', '[tex]/color', '[tex]/cancel', '[tex]/physics', '[tex]/braket',
                     '[tex]/autoload'] },
    tex: { packages: { '[+]': ['ams', 'newcommand', 'mathtools', 'boldsymbol', 'color', 'cancel',
                                'physics', 'braket', 'autoload'] } },
    svg: { fontCache: 'none' },
    startup: { typeset: false },
  });
  return MathJax;
}

// The TeX error message MathJax embedded in the render, or null.
export async function renderError(MathJax, latex, display) {
  const node = await MathJax.tex2svgPromise(latex, { display });
  const html = MathJax.startup.adaptor.outerHTML(node);
  const m = /data-mjx-error="([^"]*)"/.exec(html);
  return m ? m[1] : null;
}
