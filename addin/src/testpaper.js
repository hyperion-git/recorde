// "Test paper" (Examples → Testing): a manuscript laid out like a RevTeX
// two-column APS paper, built entirely through the add-in — Times body text,
// Termes equations (the font newtxmath maps to), flush-right table numbering as
// in a journal, inline math in running text, a reference list. Section IV is a
// deliberate showcase: every numbering style, every typeface, fixed sizes,
// colours and the built-in macro preamble. Pure data; the runner in taskpane.js
// (runDocumentScript) applies each block's `fmt`; paperOutline() exports the
// same paper as placeholder text for the headless CLI (scripts/make-showcase-paper.mjs),
// so both routes produce the same document.
//
// Block: { text | before+eq+after | eq | parts | ooxml, fmt?, showcase? }.
// parts: alternating strings and equations, e.g. ['For ', eq, ' and ', eq, '.'].
// fmt keys: font, size, bold, italic, align ('left'|'centered'|'right'|
// 'justified'), firstLineIndent, leftIndent, rightIndent, spaceBefore,
// spaceAfter, lineSpacing (pt). eq: as in testpage.js; the paper's defaults are
// Termes, fixed 10 pt, table style — blocks marked `showcase: true` may deviate.

const TIMES = 'Times New Roman';
export const PAPER_FONT = TIMES;

const F = {
  title:    { font: TIMES, size: 14, bold: true, align: 'centered', spaceBefore: 12, spaceAfter: 8 },
  authors:  { font: TIMES, size: 10, align: 'centered', spaceAfter: 2 },
  affil:    { font: TIMES, size: 9, italic: true, align: 'centered', spaceAfter: 2 },
  dated:    { font: TIMES, size: 9, align: 'centered', spaceAfter: 8 },
  abstract: { font: TIMES, size: 9, align: 'justified', leftIndent: 36, rightIndent: 36, spaceAfter: 12 },
  section:  { font: TIMES, size: 10, bold: true, align: 'centered', spaceBefore: 10, spaceAfter: 4 },
  subsec:   { font: TIMES, size: 10, bold: true, italic: true, align: 'centered', spaceBefore: 6, spaceAfter: 3 },
  body:     { font: TIMES, size: 10, align: 'justified', firstLineIndent: 12, spaceAfter: 0, lineSpacing: 12 },
  bodyNoIndent: { font: TIMES, size: 10, align: 'justified', spaceAfter: 0, lineSpacing: 12 },
  caption:  { font: TIMES, size: 9, italic: true, align: 'left', spaceBefore: 4, spaceAfter: 1 },
  ack:      { font: TIMES, size: 10, bold: true, align: 'centered', spaceBefore: 10, spaceAfter: 4 },
  ref:      { font: TIMES, size: 9, align: 'justified', leftIndent: 14, firstLineIndent: -14, spaceAfter: 0, lineSpacing: 11 },
};

// A paragraph carrying a section break (continuous) with `cols` columns applies
// those columns to the section that ENDS at it — the standard OOXML way to make
// part of a document multi-column. Word must accept paragraph-level sectPr via
// insertOoxml; the runner treats a failure as non-fatal (single column).
function sectionBreak(cols) {
  const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
  return `<pkg:package xmlns:pkg="http://schemas.microsoft.com/office/2006/xmlPackage">
<pkg:part pkg:name="/_rels/.rels" pkg:contentType="application/vnd.openxmlformats-package.relationships+xml"><pkg:xmlData>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>
</pkg:xmlData></pkg:part>
<pkg:part pkg:name="/word/document.xml" pkg:contentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"><pkg:xmlData>
<w:document xmlns:w="${W}"><w:body><w:p><w:pPr><w:sectPr><w:type w:val="continuous"/><w:cols w:num="${cols}" w:space="288"/></w:sectPr></w:pPr></w:p></w:body></w:document>
</pkg:xmlData></pkg:part></pkg:package>`;
}
const SECTION_COLS = { 1: sectionBreak(1), 2: sectionBreak(2) };

const R = String.raw;
const EQ  = { font: 'termes', sizeMode: 'fixed', fixedPt: 10, displayMode: 'inline' };   // inline math
const DEQ = { ...EQ, displayMode: 'display', numbered: true, style: 'table' };          // numbered display
const UEQ = { ...EQ, displayMode: 'display' };                                          // unnumbered display
const m = (latex) => ({ latex, ...EQ });
const d = (latex, extra = {}) => ({ latex, ...DEQ, ...extra });
const u = (latex, extra = {}) => ({ latex, ...UEQ, ...extra });

// The one equation the typeface gallery repeats in every font.
const GALLERY = R`\I\hbar\,\partial_t \psi = \left[ -\frac{\hbar^2}{2m}\Laplace + V(\Vect{r}) \right] \psi, \quad \int \abs{\psi}^2 \,\D^3 r = 1`;
const FONTS = [['tex', 'Computer Modern (tex)'], ['newcm', 'New Computer Modern (newcm)'], ['termes', 'Termes (termes) — the paper default'],
  ['stix2', 'STIX Two (stix2)'], ['pagella', 'Pagella (pagella)'], ['asana', 'Asana Math (asana)']];

export const TEST_PAPER = [
  { text: 'Phase shifts in light-pulse atom interferometers from a Baker–Campbell–Hausdorff expansion', fmt: F.title },
  { text: 'First Author and Second Author', fmt: F.authors },
  { text: 'Affiliation line (institute, city, country)', fmt: F.affil },
  { text: '(Dated: September 14, 2026)', fmt: F.dated },
  { text: 'We outline how the phase of a Mach–Zehnder light-pulse atom interferometer follows from an operator-algebraic treatment of the free evolution and the beam-splitter pulses. Expanding the product of evolution operators with the Baker–Campbell–Hausdorff formula isolates the gravitational phase and shows which commutators contribute at each order in the pulse separation. The treatment is extended to a linear gravity gradient and to finite pulse duration, and the shot-noise-limited sensitivity is derived. This document is a typesetting test: every equation was inserted with the Recorde add-in, and Sec. IV exercises every option the add-in offers.', fmt: F.abstract },
  { ooxml: SECTION_COLS[1], cols: 1 },

  { text: 'I. INTRODUCTION', fmt: F.section },
  { parts: ['Light-pulse atom interferometers measure inertial effects through the phase accumulated along two spatially separated paths [1, 2]. For a three-pulse sequence with pulse separation ',
      m(R`T`), ', effective wave number ', m(R`k_{\mathrm{eff}} = 2k`), ' and local acceleration ', m(R`g`),
      ', the leading phase is ', m(R`\Delta\phi = k_{\mathrm{eff}}\, g\, T^2`), '. The internal-state labelling introduced by Bordé [3] makes the two arms distinguishable at read-out.'], fmt: F.body },
  { parts: ['The scaling with ', m(R`T^2`), ' is what makes long drop times attractive: at ', m(R`T = 1\,\mathrm{s}`),
      ' a phase of order ', m(R`10^{8}\,\mathrm{rad}`), ' accrues, and a resolution of ', m(R`\delta\phi \sim 10^{-3}\,\mathrm{rad}`),
      ' translates into ', m(R`\delta g / g \sim 10^{-11}`), '. Realising this requires control of the gravity gradient ',
      m(R`\Gamma = \partial_z g`), ' and of the finite pulse duration ', m(R`\tau`), ', both treated below.'], fmt: F.body },
  { text: 'We write the evolution as a product of unitaries and expand it to expose the commutator structure. Section II states the Hamiltonian and the expansion, Sec. III gives the phase, the gradient correction and the sensitivity, and Sec. IV collects typesetting checks. The appendix carries the explicit commutator algebra.', fmt: F.body },

  { text: 'II. OPERATOR TREATMENT', fmt: F.section },
  { text: 'A. Hamiltonian', fmt: F.subsec },
  { text: 'For a two-level atom of mass m in a uniform gravitational field, driven by a travelling wave, the Hamiltonian in the rotating frame reads', fmt: F.body },
  { eq: d(R`\begin{aligned} \HOp{H} &= \frac{\HOp{p}^2}{2m} + m g\,\HOp{z} \\ &\quad + \frac{\hbar\Omega}{2}\left( \E^{\I(k\HOp{z} - \omega t)}\,\HOp{\sigma}_+ + \text{h.c.} \right), \end{aligned}`) },
  { parts: ['where ', m(R`\HOp{\sigma}_+ = \ket{e}\!\bra{g}`), ' raises the internal state and ', m(R`\Omega`),
      ' is the Rabi frequency. Between pulses the atom evolves freely under the first two terms, ',
      m(R`\HOp{H}_0 = \HOp{p}^2/2m + mg\HOp{z}`), '. A linear gravity gradient adds ', m(R`-\tfrac{1}{2} m \Gamma \HOp{z}^2`), '.'], fmt: F.bodyNoIndent },

  { text: 'B. Baker–Campbell–Hausdorff expansion', fmt: F.subsec },
  { text: 'The free-evolution operators of the two arms do not commute with the momentum kicks. Their product is rewritten with', fmt: F.body },
  { eq: d(R`\begin{aligned} \E^{\HOp{A}}\,\E^{\HOp{B}} = \exp\Big( &\HOp{A} + \HOp{B} + \tfrac{1}{2}\OpComm{\HOp{A}, \HOp{B}} \\ &+ \tfrac{1}{12}\OpComm{\HOp{A}, \OpComm{\HOp{A}, \HOp{B}}} \\ &- \tfrac{1}{12}\OpComm{\HOp{B}, \OpComm{\HOp{A}, \HOp{B}}} + \cdots \Big), \end{aligned}`) },
  { parts: ['which terminates for the quadratic Hamiltonian of Eq. (1) because ', m(R`\OpComm{\HOp{z}, \HOp{p}} = \I\hbar`),
      ' is a c-number: all nested commutators beyond the second order vanish identically. The free evolution over ',
      m(R`T`), ' is'], fmt: F.bodyNoIndent },
  { eq: d(R`\begin{aligned} \HOp{U}_0(T) &= \E^{-\I \HOp{H}_0 T/\hbar} \\ &= \E^{-\I m g \HOp{z} T/\hbar}\; \E^{-\I \HOp{p}^2 T/2m\hbar} \\ &\quad \times \E^{\I g T^2 \HOp{p}/2\hbar}\; \E^{-\I m g^2 T^3/6\hbar}, \end{aligned}`) },
  { text: 'where the last two factors are exactly the second-order commutator terms. With a gradient the algebra no longer closes at second order; keeping terms to first order in Γ,', fmt: F.bodyNoIndent },
  { eq: d(R`\begin{aligned} \HOp{U}_\Gamma(T) &= \HOp{U}_0(T)\,\E^{\I m \Gamma T \HOp{Q}/2\hbar} + \BigOh(\Gamma^2), \\ \HOp{Q} &= \HOp{z}^2 + \frac{T}{m}\{\HOp{z},\HOp{p}\} + \frac{T^2}{3m^2}\HOp{p}^2, \end{aligned}`) },
  { parts: ['with the anticommutator ', m(R`\{\HOp{z},\HOp{p}\} \equiv \HOp{z}\HOp{p} + \HOp{p}\HOp{z}`), '.'], fmt: F.bodyNoIndent },

  { text: 'C. Beam-splitter pulses', fmt: F.subsec },
  { parts: ['In the ', m(R`\{\ket{g,p},\,\ket{e,p+\hbar k}\}`), ' basis a resonant pulse of area ', m(R`\theta = \Omega\tau`),
      ' and laser phase ', m(R`\phi_L = k z_L - \omega t_L`), ' acts as'], fmt: F.body },
  { eq: d(R`\HOp{U}_\theta = \begin{pmatrix} \cos\frac{\theta}{2} & -\I\,\E^{-\I\phi_L}\sin\frac{\theta}{2} \\[4pt] -\I\,\E^{\I\phi_L}\sin\frac{\theta}{2} & \cos\frac{\theta}{2} \end{pmatrix},`) },
  { parts: ['so that ', m(R`\theta = \pi/2`), ' splits and ', m(R`\theta = \pi`), ' mirrors. The laser phase ', m(R`\phi_L`),
      ' imprinted at each pulse is what the interferometer reads out; the population in the excited state after the sequence is'], fmt: F.bodyNoIndent },
  { eq: d(R`\begin{aligned} P_e &= \tfrac{1}{2}\left[ 1 - C \cos\left(\Delta\phi + \phi_0\right) \right], \\ \phi_0 &= \begin{cases} 0, & \text{Mach–Zehnder}, \\[2pt] \pi/2, & \text{mid-fringe}, \end{cases} \end{aligned}`) },
  { parts: ['with contrast ', m(R`C \le 1`), '.'], fmt: F.bodyNoIndent },

  { text: 'III. RESULTS', fmt: F.section },
  { text: 'A. Phase and gradient correction', fmt: F.subsec },
  { text: 'Collecting the c-number terms of the expansion for the π/2–π–π/2 sequence gives the familiar phase', fmt: F.body },
  { eq: d(R`\Delta\phi = \vb{k}_{\mathrm{eff}} \cdot \vb{g}\; T^2,`) },
  { parts: ['independent of the initial momentum. The gradient contributes through the initial position ', m(R`z_0`),
      ' and velocity ', m(R`v_0`), ' of the wave packet,'], fmt: F.bodyNoIndent },
  { eq: d(R`\Delta\phi_\Gamma = k_{\mathrm{eff}} \Gamma T^2 \left( z_0 + v_0 T + \tfrac{7}{12} g T^2 \right),`) },
  { text: 'to first order in Γ — the source of the systematic that gradiometer configurations cancel. Finite pulse duration rescales the effective separation,', fmt: F.bodyNoIndent },
  { eq: d(R`\begin{aligned} T &\;\to\; T_{\mathrm{eff}} = T + \frac{2\tau}{\pi}\left( 1 - \frac{\pi}{4} \right), \\ \Delta\phi &= k_{\mathrm{eff}}\, g\, T_{\mathrm{eff}}^2 + \BigOh(\tau^2). \end{aligned}`) },

  { text: 'B. Sensitivity', fmt: F.subsec },
  { parts: ['For rubidium (', m(R`\lambda = \SI{780}{\nano\metre}`), ') and ', m(R`T = \SI{100}{\milli\second}`),
      ' the phase amounts to about ', m(R`1.6 \times 10^{6}\,\mathrm{rad}`),
      '. With N detected atoms at the shot-noise limit the acceleration sensitivity per shot is'], fmt: F.body },
  { eq: d(R`\delta g = \frac{1}{k_{\mathrm{eff}}\, T^2 \sqrt{N}}, \qquad \frac{\delta g}{g} \approx 6 \times 10^{-10}`) },
  { parts: ['for ', m(R`N = 10^6`), '. Averaging ', m(R`n`), ' shots of cycle time ', m(R`T_c`),
      ' improves this as the square root of the total time, giving the Allan deviation'], fmt: F.bodyNoIndent },
  { eq: d(R`\sigma_g(\tau_{\mathrm{avg}}) = \frac{1}{k_{\mathrm{eff}} T^2 \sqrt{N}}\, \sqrt{\frac{T_c}{\tau_{\mathrm{avg}}}}, \qquad \frac{1}{n}\sum_{i=1}^{n} \delta g_i \;\propto\; n^{-1/2}.`) },
  { parts: ['The velocity distribution of the source enters through the Doppler detuning ', m(R`\delta_D = k v_z`),
      '; integrating the two-level response over a Gaussian of width ', m(R`\sigma_v`), ' gives the contrast'], fmt: F.bodyNoIndent },
  { eq: d(R`\begin{aligned} C &= \int_{-\infty}^{\infty} \frac{\Omega^2}{\Omega^2 + k^2 v_z^2}\, \frac{\E^{-v_z^2/2\sigma_v^2}}{\sqrt{2\pi}\,\sigma_v}\, \D v_z \\ &\approx 1 - \frac{k^2\sigma_v^2}{\Omega^2} \qquad (k\sigma_v \ll \Omega). \end{aligned}`) },

  { text: 'IV. TYPESETTING CHECKS', fmt: F.section },
  { text: 'This section exists to exercise the add-in rather than the physics; every equation here is click-to-edit like the others.', fmt: F.body },
  { text: 'A. Numbering styles', fmt: F.subsec },
  { text: 'The paper numbers its equations flush right in a borderless table (the journal style). The same counter can also place the number inside the picture, trailing the equation:', fmt: F.body },
  { eq: d(R`\HOp{a}\,\ket{n} = \sqrt{n}\,\ket{n-1}`, { style: 'inline' }), showcase: true },
  { text: 'or as a native Word SEQ field, which Word itself keeps in sequence (desktop only; Alt+F9 reveals the field code):', fmt: F.bodyNoIndent },
  { eq: d(R`\HOp{H} = \hbar\omega\left( \HOp{a}^\dagger \HOp{a} + \tfrac{1}{2} \right)`, { style: 'field' }), showcase: true },
  { text: 'Renumber all reorders every static number in document order and refreshes the fields.', fmt: F.bodyNoIndent },

  { text: 'B. Typefaces', fmt: F.subsec },
  { text: 'The same equation set in each of the six bundled MathJax fonts. Termes matches the Times body text; the others are shown for comparison.', fmt: F.body },
  ...FONTS.flatMap(([font, label]) => [
    { text: label, fmt: F.caption },
    { eq: u(GALLERY, { font }), showcase: true },
  ]),

  { text: 'C. Sizes and colour', fmt: F.subsec },
  { parts: ['Fixed sizes in running text: ', m(R`\E^{\I\pi} + 1 = 0`), ' at 10 pt, ', { ...m(R`\E^{\I\pi} + 1 = 0`), fixedPt: 8 },
      ' at 8 pt, ', { ...m(R`\E^{\I\pi} + 1 = 0`), fixedPt: 12 }, ' at 12 pt and ', { ...m(R`\E^{\I\pi} + 1 = 0`), fixedPt: 14 },
      ' at 14 pt; descenders such as ', m(R`\int_0^1 y\,\D y = \tfrac{1}{2}`), ' sit on the baseline on desktop Word.'], fmt: F.body, showcase: true },
  { text: 'A whole equation in a swatch colour, and colour inside an equation through the built-in AFP palette macros:', fmt: F.bodyNoIndent },
  { eq: d(R`\nabla \times \vb{B} - \frac{1}{c^2}\,\partial_t \vb{E} = \mu_0 \vb{J}`, { color: '#1f77b4' }), showcase: true },
  { eq: d(R`\HOp{H} = \textcolor{afp-c0}{\HOp{H}_{\mathrm{kin}}} + \textcolor{afp-c1}{\HOp{H}_{\mathrm{grav}}} + {\color{afp-red-3} \HOp{H}_{\mathrm{int}}}`), showcase: true },

  { text: 'D. Alignment', fmt: F.subsec },
  { text: 'Display equations follow the Settings → Display choice: left (LaTeX fleqn, indented by the 25 pt math indent) is the default; centred is the plain-LaTeX look. Both, forced here regardless of the setting:', fmt: F.body },
  { eq: d(R`\HOp{U}(t) = \mathcal{T}\exp\!\left[ -\frac{\I}{\hbar}\int_0^t \HOp{H}(t')\,\D t' \right]`, { align: 'left' }), showcase: true },
  { eq: d(R`\HOp{U}(t) = \mathcal{T}\exp\!\left[ -\frac{\I}{\hbar}\int_0^t \HOp{H}(t')\,\D t' \right]`, { align: 'center' }), showcase: true },

  { text: 'E. Macro preamble', fmt: F.subsec },
  { parts: ['The built-in preamble ports the author’s paper template: operators, delimiters, derivatives, vectors, a siunitx subset (', m(R`\SI{1.2(3)}{\kilo\hertz}`), ', ', m(R`\si{\metre\per\second\squared}`), ') and the colour names used above are available in every equation without a \\require.'], fmt: F.body },
  { eq: d(R`\Real\!\left[ \bra{\psi} \HOp{U}^\dagger \HOp{O} \HOp{U} \ket{\psi} \right] = \sum_{n} p_n \, \expval{\HOp{O}}_n,`) },
  { eq: d(R`\Div \Vect{j} + \partial_t \rho = 0, \qquad \dv{\theta}{t} = \frac{1}{\I\hbar}\OpComm{\HOp{H}, \HOp{\theta}},`) },
  { eq: d(R`\begin{aligned} \mathcal{L}[\DensityOperator] &= -\frac{\I}{\hbar}\OpComm{\HOp{H}, \DensityOperator} + \sum_k \gamma_k\, \mathcal{D}[\HOp{L}_k]\DensityOperator, \\ \mathcal{D}[\HOp{L}]\DensityOperator &= \HOp{L} \DensityOperator \HOp{L}^\dagger - \tfrac{1}{2}\{ \HOp{L}^\dagger \HOp{L}, \DensityOperator \}. \end{aligned}`) },

  { text: 'V. CONCLUSION', fmt: F.section },
  { text: 'The operator expansion reproduces the standard result and makes the order-by-order bookkeeping explicit; the gradient and finite-duration corrections follow from the same algebra. In Word, the equations above are vector pictures with their LaTeX source stored in the document: click any of them to reopen it in the pane, or list them in the pane’s equation list.', fmt: F.body },

  { text: 'ACKNOWLEDGMENTS', fmt: F.ack },
  { text: 'This document was generated by the Recorde add-in as a typesetting test; a byte-compatible twin is produced headlessly by mjx-docx from the same source.', fmt: F.bodyNoIndent },

  { text: 'APPENDIX: COMMUTATOR ALGEBRA', fmt: F.section },
  { parts: ['With ', m(R`\HOp{A} = -\I \HOp{p}^2 T/2m\hbar`), ' and ', m(R`\HOp{B} = -\I m g \HOp{z} T/\hbar`), ' the commutators needed in Eq. (2) are'], fmt: F.body },
  { eq: d(R`\begin{aligned} \OpComm{\HOp{A}, \HOp{B}} &= -\frac{g T^2}{2\hbar^2}\,\OpComm{\HOp{p}^2, \HOp{z}} = \frac{\I g T^2}{\hbar}\,\HOp{p}, \\ \OpComm{\HOp{A}, \OpComm{\HOp{A}, \HOp{B}}} &= 0, \\ \OpComm{\HOp{B}, \OpComm{\HOp{A}, \HOp{B}}} &= -\frac{\I m g^2 T^3}{\hbar}. \end{aligned}`) },
  { parts: ['Inserting these into the expansion and using the disentangling identity, valid when ',
      m(R`\OpComm{\HOp{X}, \OpComm{\HOp{X},\HOp{Y}}} = \OpComm{\HOp{Y}, \OpComm{\HOp{X},\HOp{Y}}} = 0`), ','], fmt: F.bodyNoIndent },
  { eq: d(R`\E^{\HOp{X} + \HOp{Y}} = \E^{\HOp{X}}\,\E^{\HOp{Y}}\,\E^{-\frac{1}{2}\OpComm{\HOp{X},\HOp{Y}}},`) },
  { parts: ['yields Eq. (3). The three-pulse phase then follows from the c-number factors alone, with ',
      m(R`t_{j} = t_1 + (j-1)T`), ':'], fmt: F.bodyNoIndent },
  { eq: d(R`\Delta\phi = \phi_L(t_1) - 2\phi_L(t_2) + \phi_L(t_3) = k_{\mathrm{eff}}\, g\, T^2.`) },

  { text: '[1] M. Kasevich and S. Chu, Phys. Rev. Lett. 67, 181 (1991).', fmt: { ...F.ref, spaceBefore: 8 } },
  { text: '[2] A. Peters, K. Y. Chung, and S. Chu, Nature 400, 849 (1999).', fmt: F.ref },
  { text: '[3] Ch. J. Bordé, Phys. Lett. A 140, 10 (1989).', fmt: F.ref },
  { text: '[4] P. Storey and C. Cohen-Tannoudji, J. Phys. II France 4, 1999 (1994).', fmt: F.ref },
  { text: '[5] K. Bongs, R. Launay, and M. A. Kasevich, Appl. Phys. B 84, 599 (2006).', fmt: F.ref },
  { text: '[6] A. Roura, W. Zeller, and W. P. Schleich, New J. Phys. 16, 123012 (2014).', fmt: F.ref },
  { ooxml: SECTION_COLS[2], cols: 2 },
];

export function testPaperEquations() {
  return TEST_PAPER.flatMap((b) => (b.eq ? [b.eq] : b.parts ? b.parts.filter((x) => typeof x !== 'string') : []));
}

// ---- Headless twin ----
// The paper as placeholder text: [{ kind:'p', text, fmt } | { kind:'section', cols }].
// Equation attributes that differ from the CLI defaults (Termes, table style,
// run size, black) become placeholder options so `mjx-docx process` reproduces
// the add-in's choices exactly.
export function paperOutline() {
  const ph = (eq) => {
    const o = [];
    if (eq.font !== 'termes') o.push(`font=${eq.font}`);
    if (eq.color) o.push(`color=${eq.color}`);
    if (eq.sizeMode === 'fixed' && eq.fixedPt !== 10) o.push(`size=${eq.fixedPt}`);
    if (eq.numbered && eq.style && eq.style !== 'table') o.push(`style=${eq.style}`);
    if (eq.align) o.push(`align=${eq.align}`);
    const kind = eq.displayMode === 'inline' ? 'math' : eq.numbered ? 'eq' : 'display';
    return `[[${kind}${o.length ? `{${o.join(',')}}` : ''}: ${eq.latex}]]`;
  };
  return TEST_PAPER.map((b) => {
    if (b.ooxml) return { kind: 'section', cols: b.cols };
    const text = b.parts ? b.parts.map((x) => (typeof x === 'string' ? x : ph(x))).join('')
      : b.eq ? [b.before || '', ph(b.eq), b.after || ''].join('')
        : b.text;
    return { kind: 'p', text, fmt: b.fmt || null };
  });
}
