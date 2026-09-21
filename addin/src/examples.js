// Example equations for the "Examples" picker — pure data, so the set is
// unit-tested (shape, and that every example actually renders through MathJax
// with the default macro preamble). Loading one starts a NEW equation with the
// given source and mode; `numbered` also ticks "Number this equation" so the
// numbering path is exercised. Doubles as the manual smoke-test set
// (docs/VERIFY.md). Sources keep real newlines: the editor shows them as
// typed and flattenLatex() joins lines before MathJax sees them.

export const EXAMPLES = [
  {
    id: 'gaussian', name: 'Gaussian integral', mode: 'display',
    latex: String.raw`\int_{-\infty}^{\infty} e^{-x^2}\,dx = \sqrt{\pi}`,
  },
  {
    id: 'euler', name: 'Euler’s identity', mode: 'inline',
    latex: String.raw`e^{i\pi} + 1 = 0`,
  },
  {
    id: 'quadratic', name: 'Quadratic formula', mode: 'display',
    latex: String.raw`x = \frac{-b \pm \sqrt{b^2 - 4ac}}{2a}`,
  },
  {
    id: 'maxwell', name: 'Maxwell’s equations', mode: 'display',
    latex: [
      String.raw`\begin{aligned}`,
      String.raw`  \nabla \cdot \mathbf{E} &= \frac{\rho}{\varepsilon_0} \\`,
      String.raw`  \nabla \cdot \mathbf{B} &= 0 \\`,
      String.raw`  \nabla \times \mathbf{E} &= -\frac{\partial \mathbf{B}}{\partial t} \\`,
      String.raw`  \nabla \times \mathbf{B} &= \mu_0 \mathbf{J} + \mu_0 \varepsilon_0 \frac{\partial \mathbf{E}}{\partial t}`,
      String.raw`\end{aligned}`,
    ].join('\n'),
  },
  {
    id: 'schroedinger', name: 'Schrödinger equation', mode: 'display', numbered: true,
    latex: String.raw`i\hbar \frac{\partial}{\partial t} \Psi(\mathbf{r}, t) = \left[ -\frac{\hbar^2}{2m} \nabla^2 + V(\mathbf{r}, t) \right] \Psi(\mathbf{r}, t)`,
  },
  {
    id: 'matrix', name: 'Rotation matrix', mode: 'display',
    latex: [
      String.raw`R(\theta) = \begin{pmatrix}`,
      String.raw`  \cos\theta & -\sin\theta \\`,
      String.raw`  \sin\theta & \cos\theta`,
      String.raw`\end{pmatrix}`,
    ].join('\n'),
  },
  {
    id: 'cases', name: 'Piecewise function', mode: 'display',
    latex: [
      String.raw`|x| = \begin{cases}`,
      String.raw`  x & x \ge 0 \\`,
      String.raw`  -x & x < 0`,
      String.raw`\end{cases}`,
    ].join('\n'),
  },
  {
    // Uses macros from the default preamble (preamble.js): \HOp, \Avg*, \Abs, \E, \I.
    id: 'macros', name: 'Preamble macros', mode: 'display', numbered: true,
    latex: String.raw`\Avg*{\HOp{H}} = \sum_n \Abs{c_n}^2 E_n, \qquad c_n = \Abs{c_n}\,\E^{\I\varphi_n}`,
  },
  {
    // 0+1D path integral: the propagator with the discretised measure. \I from the
    // preamble; \mathcal for the measure; \dot, \lim, \prod.
    id: 'pathintegral', name: 'Path integral (0+1D)', mode: 'display', numbered: true,
    latex: [
      String.raw`\begin{aligned}`,
      String.raw`  \langle x_f, t_f \,|\, x_i, t_i \rangle`,
      String.raw`    &= \int_{x(t_i)=x_i}^{x(t_f)=x_f} \mathcal{D}x(t)\,`,
      String.raw`       \exp\!\left[ \frac{\I}{\hbar} \int_{t_i}^{t_f} \!dt\, \left( \frac{m}{2}\dot{x}^2 - V(x) \right) \right], \\`,
      String.raw`  \mathcal{D}x &= \lim_{N\to\infty} \left( \frac{m}{2\pi \I \hbar\,\varepsilon} \right)^{N/2} \prod_{k=1}^{N-1} dx_k,`,
      String.raw`  \qquad \varepsilon = \frac{t_f - t_i}{N}`,
      String.raw`\end{aligned}`,
    ].join('\n'),
  },
  {
    // Scalar field theory: generating functional with a \phi^4 interaction. Uses
    // \mathbb, \mathcal, index placement, \partial^\mu and the functional derivative.
    id: 'fieldtheory', name: 'Path integral (field theory)', mode: 'display', numbered: true,
    latex: [
      String.raw`\begin{aligned}`,
      String.raw`  Z[J] &= \int \mathcal{D}\phi \; \exp\!\left\{ \I \int d^4x \left[`,
      String.raw`    \tfrac{1}{2}\,\partial_\mu \phi\, \partial^\mu \phi - \tfrac{1}{2} m^2 \phi^2 - \frac{\lambda}{4!}\phi^4 + J\phi \right] \right\},`,
      String.raw`    \qquad \phi : \mathbb{R}^{1,3} \to \mathbb{R}, \\`,
      String.raw`  \langle 0 | \mathcal{T}\, \phi(x_1)\cdots\phi(x_n) | 0 \rangle`,
      String.raw`    &= \left. \frac{(-\I)^n}{Z[0]} \frac{\delta^n Z[J]}{\delta J(x_1)\cdots\delta J(x_n)} \right|_{J=0}`,
      String.raw`\end{aligned}`,
    ].join('\n'),
  },
  {
    // Incompressible Navier–Stokes: physics-package \vb and \pdv plus the preamble's
    // bold nabla operators (\VNabla, \VNablaLaplace).
    id: 'navierstokes', name: 'Navier–Stokes', mode: 'display', numbered: true,
    latex: [
      String.raw`\rho \left( \pdv{\vb{u}}{t} + (\vb{u} \cdot \VNabla)\, \vb{u} \right)`,
      String.raw`  = -\VNabla p + \mu\, \VNablaLaplace \vb{u} + \rho\, \vb{g},`,
      String.raw`\qquad \VNabla \cdot \vb{u} = 0`,
    ].join('\n'),
  },
  {
    // Field quantisation in a cavity: mode sum, ladder operators (\HOp), the
    // preamble's commutator delimiter \OpComm, \E and \I, bold vectors, \text.
    id: 'cavity', name: 'Field quantisation (cavity)', mode: 'display', numbered: true,
    latex: [
      String.raw`\begin{aligned}`,
      String.raw`  \HOp{H} &= \sum_{\vb{k},\lambda} \hbar\omega_k \left( \HOp{a}^\dagger_{\vb{k}\lambda} \HOp{a}_{\vb{k}\lambda} + \tfrac{1}{2} \right), \\`,
      String.raw`  \HOp{\vb{E}}(\vb{r}, t) &= \I \sum_{\vb{k},\lambda} \sqrt{\frac{\hbar\omega_k}{2\varepsilon_0 V}}\;`,
      String.raw`    \vb{e}_{\vb{k}\lambda} \left[ \HOp{a}_{\vb{k}\lambda}\, \E^{\I(\vb{k}\cdot\vb{r} - \omega_k t)} - \text{h.c.} \right], \\`,
      String.raw`  \OpComm{\HOp{a}_{\vb{k}\lambda}, \HOp{a}^\dagger_{\vb{k}'\lambda'}} &= \delta_{\vb{k}\vb{k}'}\, \delta_{\lambda\lambda'}`,
      String.raw`\end{aligned}`,
    ].join('\n'),
  },
];

export function findExample(id) {
  return EXAMPLES.find((e) => e.id === id) || null;
}
