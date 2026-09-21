// Built-in default macro preamble, ported from the author's paper template
// (05-cmds.tex). Applied ONCE after each MathJax (re)load — NOT prepended per
// render — because MathJax persists definitions and \DeclareMathOperator /
// \DeclarePairedDelimiter error on re-definition. See loadMathJax in taskpane.js.
//
// Backing extensions (loaded in taskpane.js): ams (\DeclareMathOperator),
// mathtools (\DeclarePairedDelimiter, \mathrlap), boldsymbol, physics, braket,
// color, cancel, autoload.
//
// Acceptance: every math snippet of the template's main document
// (test/fixtures/template-math.json, refreshed by scripts/extract-template-math.mjs)
// must render with this preamble without a TeX error or an undefined macro.
//
// What MathJax 4 cannot express, and what is done instead:
//   • xparse \NewDocumentCommand — \Ds/\FuncDs have no starred form (the
//     unstarred \mathop{}\!d form is ported); \PhaseSpaceArgs(Vec) take ONE
//     optional (the superscript; the template has two); \intRlap is a plain
//     \int (its e{_^} limit capture has no MathJax equivalent).
//   • \mathpalette/\raisebox{\depth}: \sym/\raisesym raise by a fixed 0.2ex
//     via \raise instead of by the symbol's own depth.
//   • \bBigg@{n}: \xsize{n}, \vast, \Vast all give \Bigg (MathJax's largest).
//   • stackengine: \barbelow puts the template's .5em × .075ex rule under the
//     argument with \underset; \TensorMatrixOverarrow uses \leftrightarrow for
//     mathabx's \smallleftrightarrow; \eql → mathtools \xlongequal;
//     \over*harpoon → \overset with a harpoon (not stretchy); \Fvect (esvect
//     \vv) → \overrightarrow.
//   • bbm: \mathbbm is an alias of \mathbb (the double-struck range of every
//     bundled font has the 𝟙 glyph); \CharFunc uses it.
//   • siunitx: a small \SI/\si/\num subset with the SI prefixes and the units
//     the author's papers use (10-main.tex); "1.2(3)" stays literal.
//   • Text-mode only, DROPPED: \MyItem* (enumerate labels), \Comment*, \aff*.
//   • Macros the template's MAIN document defines (\II, \EE, \pdg) are included
//     so the author's manuscripts paste unchanged.
//   • AFP colours (03-colors.tex) ARE ported as \definecolor{…}{RGB}{r,g,b}
//     (\colorlet resolved at port time) — usable as \textcolor{afp-c0}{…}.

export const DEFAULT_PREAMBLE = String.raw`
% Euler / imaginary
\newcommand{\I}{\mathrm{i}}
\newcommand{\E}{\mathrm{e}}
% Function wrappers
\newcommand{\SFunc}[2]{#1\left(#2\right)}
\newcommand{\VFunc}[2]{\boldsymbol{#1}\left(#2\right)}
% Operator-notation fonts
\newcommand{\HOp}[1]{\widehat{#1}}
\newcommand{\SfOp}[1]{\mathsf{#1}}
\newcommand{\SfOpB}[1]{\boldsymbol{\mathsf{#1}}}
\newcommand{\AbstractVector}[1]{\mathsf{#1}}
\newcommand{\AbstractTensor}[1]{\mathsf{#1}}
\newcommand{\AbstractOperator}[1]{\mathsf{#1}}
% Density operator + field notation
\newcommand{\DensityOperator}{\QmOp{\rho}}
\newcommand{\RSpaceField}[1]{#1\left(\boldsymbol{x},t\right)}
\newcommand{\KSpaceField}[1]{#1\left(\boldsymbol{k},\omega\right)}
\newcommand{\GenField}[2]{{#1}\left( {#2} \right)}
\newcommand{\TMat}[1]{\underline{\underline{#1}}}
% Paired delimiters (starred form auto-sizes: \Abs*{x})
\DeclarePairedDelimiter\Avg{\langle}{\rangle}
\DeclarePairedDelimiter\Abs{\lvert}{\rvert}
\DeclarePairedDelimiter\Norm{\lVert}{\rVert}
\DeclarePairedDelimiter\OpComm{[}{]}
\DeclarePairedDelimiter\OpAComm{\{}{\}}
\DeclarePairedDelimiter\InnerBracket{[}{]}
% Mode brackets (rigged Hilbert space)
\DeclarePairedDelimiter\ModeBraket{(}{)}
\DeclarePairedDelimiter\ModeKet{\lvert}{)}
\DeclarePairedDelimiter\ModeBra{(}{\rvert}
\newcommand{\ModeOperator}[1]{\boldsymbol{\mathscr{#1}}}
\newcommand{\ModeBraketSeparator}{\vert}
% Vector calculus
\DeclareMathOperator{\VNabla}{\boldsymbol{\nabla}}
\DeclareMathOperator{\VNablaLaplace}{\boldsymbol{\nabla}^2}
\DeclareMathOperator{\Div}{div}
\DeclareMathOperator{\Rot}{rot}
\DeclareMathOperator{\Grad}{grad}
\DeclareMathOperator{\Laplace}{\Delta}
\DeclareMathOperator{\VLaplace}{\boldsymbol{\Delta}}
\DeclareMathOperator{\BigOh}{\mathcal{O}}
\DeclareMathOperator{\SmallOh}{\mathcal{o}}
% Deltas & special functions
\newcommand{\DDelta}[1]{\delta\left(#1\right)}
\newcommand{\KDelta}[1]{\delta_{#1}}
\newcommand{\GDelta}[2]{\delta\left(#1,#2\right)}
\DeclareMathOperator{\Sign}{sgn}
\newcommand{\mathbbm}[1]{\mathbb{#1}}
\newcommand{\CharFunc}[1]{\mathbbm{1}_{#1}}
% Inverse hyperbolic / trig
\DeclareMathOperator{\Artanh}{artanh}
\DeclareMathOperator{\Arcoth}{arcoth}
\DeclareMathOperator{\Arsinh}{arsinh}
\DeclareMathOperator{\Arcosh}{arcosh}
\DeclareMathOperator{\Arccot}{arccot}
% Complex analysis & linear algebra
\DeclareMathOperator{\Real}{Re}
\DeclareMathOperator{\Imag}{Im}
\DeclareMathOperator{\Kern}{ker}
\DeclareMathOperator{\Image}{img}
\DeclareMathOperator{\Dim}{dim}
\DeclareMathOperator{\Span}{span}
\DeclareMathOperator{\Dom}{dom}
\DeclareMathOperator{\Det}{det}
\newcommand{\Transpose}{\intercal}
% Distributions
\DeclareMathOperator{\HeavisideTheta}{\theta}
\DeclareMathOperator{\DiracDelta}{\delta}
% Differentials & functional D (xparse starred variants dropped)
\newcommand{\D}{\mathrm{d}}
\newcommand{\FuncD}{\mathscr{D}}
\newcommand{\Ds}[1]{\mathop{}\!\mathrm{d}#1}
\newcommand{\FuncDs}[1]{\mathop{}\!\mathscr{D}#1}
% Principal value / c.c. / h.c.
\DeclareMathOperator{\CpVCal}{\mathcal{P}}
\DeclareMathOperator{\CpVpv}{\mathrm{p.v.}}
\DeclareMathOperator{\CpVPV}{\mathrm{P.V.}}
\DeclareMathOperator{\ComplexConjugate}{\mathrm{c.c.}}
\DeclareMathOperator{\HermitianConjugate}{\mathrm{h.c.}}
% Vectors & components
\newcommand{\Vect}[1]{\boldsymbol{#1}}
\newcommand{\CovComp}[2]{{#1}^{\left(#2\right)}}
\newcommand{\ConComp}[2]{{#1}_{#2}}
% Integration domains (\mathrlap from mathtools)
\newcommand{\RIntInf}{\int_{\mathrlap{-\infty}}^{+\infty}}
\newcommand{\RIntR}{\int_{\mathrlap{\mathbb{R}}}}
\newcommand{\RIInt}{\int_{\mathrlap{\mathbb{R}^2}}}
\newcommand{\RIIInt}{\int_{\mathrlap{\mathbb{R}^3}}}
\newcommand{\VarInt}[2]{\int_{\mathrlap{#1}}^{\mathrlap{#2}}}
\newcommand{\VarIntU}[1]{\int_{\mathrlap{#1}}}
\newcommand{\VarIntL}[1]{\int^{\mathrlap{#1}}}
% Fourier-transform helpers (sfrac rewritten as plain /)
\newcommand{\FTupFac}[2]{\E^{-\I#1 \cdot #2}}
\newcommand{\FTdnFac}[2]{\E^{+\I#1 \cdot #2}}
\newcommand{\FTpreFac}[1]{\frac{1}{\left(2\pi\right)^{#1/2}}}
% Matrix notation + spin arrows
\newcommand{\BoMatrix}[1]{\boldsymbol{\mathbf{#1}}}
\newcommand{\UlMatrix}[1]{\underline{\underline{#1}}}
\newcommand{\Vpar}{\upharpoonleft\hspace{-3pt}\upharpoonright}
\newcommand{\Vapar}{\upharpoonleft\hspace{-3pt}\downharpoonright}
% Quantum operators core
\newcommand{\OpDg}[1]{#1^\dagger}
\newcommand{\QmOp}[1]{\hat{#1}}
\newcommand{\QmOpVec}[1]{\QmOp{#1}}
% Phase space
\newcommand{\PhaseSpaceX}{\chi}
\newcommand{\PhaseSpaceP}{\wp}
\newcommand{\PhaseSpaceXVec}{\Vect{\chi}}
\newcommand{\PhaseSpacePVec}{\Vect{\wp}}
\newcommand{\PhaseSpaceAction}{\mathcal{A}}
% Quantum operator instances
\newcommand{\QmOpP}{\QmOp{p}}
\newcommand{\QmOpX}{\QmOp{x}}
\newcommand{\QmOpQ}{\QmOp{q}}
\newcommand{\QmOpPVec}{\QmOp{\Vect{p}}}
\newcommand{\QmOpXVec}{\QmOp{\Vect{x}}}
\newcommand{\QmOpQVec}{\QmOp{\Vect{q}}}
\newcommand{\QmOpBraketProjector}[1]{\Ket{#1}\Bra{#1}}
\newcommand{\QmOpBraketTransition}[2]{\Ket{#2}\Bra{#1}}
\newcommand{\QmOpProjector}{\QmOp{P}}
\newcommand{\QmUnity}{\mathbb{1}}
\newcommand{\QmOpMoellerPlus}{\QmOp{\Omega}_+}
\newcommand{\QmOpMoellerMinus}{\QmOp{\Omega}_-}
\newcommand{\QmOpMoellerPM}{\QmOp{\Omega}_\pm}
\newcommand{\QmOpHamiltonian}{\QmOp{H}}
\newcommand{\QmOpHamiltonianFree}{\QmOp{H}_0}
\newcommand{\QmOpInteraction}{\QmOp{V}}
\newcommand{\QmOpTimeEvolution}{\QmOp{U}}
\newcommand{\QmOpUPath}[2]{{\QmOp{U}^{(#1)}_{#2}}}
\newcommand{\QmOpUPathAdj}[2]{{\QmOp{U}^{\dagger(#1)}_{#2}}}
\newcommand{\QmOpDisplacement}{\QmOp{D}}
% Superoperators
\newcommand{\OpInnerDerivation}{\SfOp{D}}
\newcommand{\OpOrdering}{\SfOp{T}}
\newcommand{\OpSMatrix}{\SfOp{S}}
\newcommand{\OpGreen}{\SfOp{G}}
\newcommand{\OpGreenFunction}{\QmOp{G}}
\newcommand{\OpTMatrix}{\QmOp{T}}
% Orderings
\newcommand{\OpTorderFw}{\mathsf{T}_{\mathsf{-}}}
\newcommand{\OpTorderBw}{\mathsf{T}_{\mathsf{+}}}
\newcommand{\OpTorderL}{\mathsf{T}_{\mathsf{L}}}
\newcommand{\OpTorderR}{\mathsf{T}_{\mathsf{R}}}
\newcommand{\OpOrderTime}{\mathsf{T}}
\newcommand{\OpOrderAntiTime}{\overline{\mathsf{T}}}
\newcommand{\OpTorderCt}{\mathsf{T}_{\mathsf{c}}}
\newcommand{\OpLiouville}{\mathsf{L}}
\newcommand{\OpAbstract}[1]{\mathsf{#1}}
\newcommand{\OpOrderGen}[1]{:\mspace{2mu}\mathrel{#1}\mspace{2mu}:}
\newcommand{\OpNormOrder}[1]{:\mspace{2mu}\mathrel{#1}\mspace{2mu}:_{\mathsf{+}}}
\newcommand{\OpANormOrder}[1]{{:\mspace{2mu}\mathrel{#1}\mspace{2mu}:}_{\mathsf{-}}}
\newcommand{\OpSymOrder}[1]{{:\mspace{2mu}\mathrel{#1}\mspace{2mu}:}_{\mathsf{s}}}
\newcommand{\OpWeylOrder}[1]{{:\mspace{2mu}\mathrel{#1}\mspace{2mu}:}_{\mathsf{w}}}
% === Ports of the remaining template macros (MathJax approximations) =========
% Sans-italic operators: MathJax cannot switch to sans italic inside text — plain sans.
\newcommand{\SfOpI}[1]{\mathsfit{#1}}
\newcommand{\SfOpBI}[1]{\mathbfsfit{#1}}
\newcommand{\OperandPlaceholder}{\boldsymbol{\cdot}}
\newcommand{\OpPlaceholder}{\mathord{\textcolor{afp-gray-medium}{\bullet}}}
% Tensor decorations (stackengine \barbelow → underline)
\newcommand{\TensorMatrixOverarrow}[1]{\overset{\scriptscriptstyle\leftrightarrow}{#1}}
\newcommand{\barbelow}[1]{\underset{\rule{.5em}{.075ex}}{#1}}
\newcommand{\TensorMatrixUnderlined}[1]{\underline{\underline{#1}}}
\newcommand{\SecondOrderTensor}[1]{\TensorMatrixUnderlined{#1}}
\newcommand{\DeltaTransversal}{{\SecondOrderTensor{\Vect{\delta}}^{\perp}}}
\newcommand{\DeltaLongitudinal}{{\SecondOrderTensor{\Vect{\delta}}}^{\parallel}}
% Nested inner brackets (paired delimiters can't carry a subscript on the closer)
\newcommand{\InnerBracketRightNested}[1]{\left[#1\right]_{R}}
\newcommand{\InnerBracketLeftNested}[1]{\left[#1\right]_{L}}
% Operators missed in the first port
\DeclareMathOperator{\Arctan}{arctan}
\DeclareMathOperator{\Abb}{map}
\DeclareMathOperator{\CpVPVCal}{\mathcal{P.V.}}
\DeclareMathOperator*{\SumInt}{\sum\mkern-18mu\int}
% Vectors / arrows / equals-with-label
\newcommand{\Fvect}[1]{\overrightarrow{\boldsymbol{#1}}}
\newcommand{\eql}[2][]{\xlongequal[#1]{#2}}
\newcommand{\overrightharpoon}[1]{\overset{\rightharpoonup}{#1}}
\newcommand{\overleftharpoon}[1]{\overset{\leftharpoonup}{#1}}
% xparse variants without their optional/starred forms
\newcommand{\PhaseSpaceArgs}[1][]{\PhaseSpaceX^{#1},\PhaseSpaceP^{#1},\PhaseSpaceAction^{#1}}
\newcommand{\PhaseSpaceArgsVec}[1][]{\PhaseSpaceXVec^{#1},\PhaseSpacePVec^{#1},\PhaseSpaceAction^{#1}}
\newcommand{\intRlap}{\int}
% Big delimiters (the template's extra sizes 4/5 → \Bigg)
\newcommand{\xsize}[1]{\Bigg}
\newcommand{\vast}{\Bigg}
\newcommand{\Vast}{\Bigg}
% \sym: raise a symbol so it sits on the baseline (template: by its own depth)
\newcommand{\raisesym}[2]{\raise{0.2ex}{#1}}
\newcommand{\sym}[1]{\raise{0.2ex}{#1}}
% Local macros of the template's main document
\newcommand{\II}{\mathrm{i}}
\newcommand{\EE}{\mathrm{e}}
\newcommand{\pdg}{{\vphantom{\dagger}}}
% siunitx subset: \SI{value}{unit}, \si{unit}, \num{value}; prefixes + units
\newcommand{\num}[1]{#1}
\newcommand{\si}[1]{\mathrm{#1}}
\newcommand{\SI}[2]{#1\,\mathrm{#2}}
\newcommand{\qty}[2]{#1\,\mathrm{#2}}
\newcommand{\unit}[1]{\mathrm{#1}}
\newcommand{\per}{/}
\newcommand{\squared}{^{2}}
\newcommand{\cubed}{^{3}}
\newcommand{\tothe}[1]{^{#1}}
\newcommand{\femto}{f}\newcommand{\pico}{p}\newcommand{\nano}{n}\newcommand{\micro}{\mu}
\newcommand{\milli}{m}\newcommand{\centi}{c}\newcommand{\kilo}{k}\newcommand{\mega}{M}
\newcommand{\giga}{G}\newcommand{\tera}{T}
\newcommand{\metre}{m}\newcommand{\meter}{m}\newcommand{\second}{s}\newcommand{\gram}{g}
\newcommand{\kilogram}{kg}\newcommand{\hertz}{Hz}\newcommand{\kelvin}{K}\newcommand{\mole}{mol}
\newcommand{\ampere}{A}\newcommand{\volt}{V}\newcommand{\watt}{W}\newcommand{\joule}{J}
\newcommand{\newton}{N}\newcommand{\pascal}{Pa}\newcommand{\tesla}{T}\newcommand{\ohm}{\Omega}
\newcommand{\radian}{rad}\newcommand{\steradian}{sr}\newcommand{\electronvolt}{eV}
\newcommand{\degree}{^{\circ}}\newcommand{\percent}{\%}\newcommand{\minute}{min}\newcommand{\hour}{h}
% === AFP colours from the template (03-colors.tex), \colorlet aliases resolved =
% Use as \textcolor{afp-c0}{…} or \color{afp-red-3}; names match the pane's swatches.
\definecolor{afp-c0}{RGB}{46,44,184}
\definecolor{afp-c1}{RGB}{219,0,43}
\definecolor{afp-c2}{RGB}{31,138,112}
\definecolor{afp-c3}{RGB}{253,116,0}
\definecolor{afp-c4}{RGB}{28,128,158}
\definecolor{afp-c5}{RGB}{190,219,67}
\definecolor{afp-c6}{RGB}{155,47,92}
\definecolor{afp-c7}{RGB}{250,189,30}
\definecolor{afp-c8}{RGB}{92,45,153}
\definecolor{afp-c9}{RGB}{44,107,47}
\definecolor{afp-blue}{RGB}{46,44,184}
\definecolor{afp-red}{RGB}{219,0,43}
\definecolor{afp-green}{RGB}{31,138,112}
\definecolor{afp-orange}{RGB}{253,116,0}
\definecolor{afp-blueberry}{RGB}{28,128,158}
\definecolor{afp-limegreen}{RGB}{190,219,67}
\definecolor{afp-pink}{RGB}{155,47,92}
\definecolor{afp-gold}{RGB}{250,189,30}
\definecolor{afp-violet}{RGB}{92,45,153}
\definecolor{afp-forest}{RGB}{44,107,47}
\definecolor{afp-turquoise}{RGB}{0,67,88}
\definecolor{afp-yellow}{RGB}{255,225,25}
\definecolor{afp-pres-c0}{RGB}{62,54,222}
\definecolor{afp-pres-c1}{RGB}{219,0,43}
\definecolor{afp-pres-c2}{RGB}{0,155,78}
\definecolor{afp-pres-c3}{RGB}{253,116,0}
\definecolor{afp-pres-c4}{RGB}{27,154,170}
\definecolor{afp-pres-c5}{RGB}{190,219,67}
\definecolor{afp-pres-c6}{RGB}{205,87,137}
\definecolor{afp-pres-c7}{RGB}{250,189,30}
\definecolor{afp-pres-c8}{RGB}{138,87,205}
\definecolor{afp-pres-c9}{RGB}{48,193,63}
\definecolor{afp-red-1}{RGB}{73,0,6}
\definecolor{afp-red-2}{RGB}{101,0,8}
\definecolor{afp-red-3}{RGB}{150,3,15}
\definecolor{afp-red-4}{RGB}{202,0,17}
\definecolor{afp-red-5}{RGB}{246,71,86}
\definecolor{afp-red-6}{RGB}{249,186,186}
\definecolor{afp-orange-1}{RGB}{73,22,0}
\definecolor{afp-orange-2}{RGB}{118,49,0}
\definecolor{afp-orange-3}{RGB}{198,83,0}
\definecolor{afp-orange-4}{RGB}{255,119,21}
\definecolor{afp-orange-5}{RGB}{247,171,106}
\definecolor{afp-orange-6}{RGB}{251,208,171}
\definecolor{afp-yellow-1}{RGB}{79,58,2}
\definecolor{afp-yellow-2}{RGB}{155,111,3}
\definecolor{afp-yellow-3}{RGB}{223,162,4}
\definecolor{afp-yellow-4}{RGB}{250,189,30}
\definecolor{afp-yellow-5}{RGB}{252,209,105}
\definecolor{afp-yellow-6}{RGB}{252,229,173}
\definecolor{afp-green-1}{RGB}{12,43,14}
\definecolor{afp-green-2}{RGB}{29,72,37}
\definecolor{afp-green-3}{RGB}{44,107,47}
\definecolor{afp-green-4}{RGB}{58,147,0}
\definecolor{afp-green-5}{RGB}{132,215,0}
\definecolor{afp-green-6}{RGB}{198,244,111}
\definecolor{afp-blue-1}{RGB}{22,25,59}
\definecolor{afp-blue-2}{RGB}{53,71,140}
\definecolor{afp-blue-3}{RGB}{78,122,199}
\definecolor{afp-blue-4}{RGB}{127,178,240}
\definecolor{afp-blue-5}{RGB}{173,213,247}
\definecolor{afp-blue-6}{RGB}{204,229,250}
\definecolor{afp-purple-1}{RGB}{37,6,77}
\definecolor{afp-purple-2}{RGB}{54,23,94}
\definecolor{afp-purple-3}{RGB}{85,50,133}
\definecolor{afp-purple-4}{RGB}{123,82,171}
\definecolor{afp-purple-5}{RGB}{151,104,209}
\definecolor{afp-purple-6}{RGB}{187,146,239}
\definecolor{afp-pink-1}{RGB}{53,26,35}
\definecolor{afp-pink-2}{RGB}{83,33,49}
\definecolor{afp-pink-3}{RGB}{132,39,79}
\definecolor{afp-pink-4}{RGB}{204,85,156}
\definecolor{afp-pink-5}{RGB}{221,153,204}
\definecolor{afp-pink-6}{RGB}{239,190,225}
\definecolor{afp-kobalt-1}{RGB}{0,12,89}
\definecolor{afp-kobalt-2}{RGB}{24,6,160}
\definecolor{afp-kobalt-3}{RGB}{13,44,211}
\definecolor{afp-kobalt-4}{RGB}{66,95,239}
\definecolor{afp-kobalt-5}{RGB}{168,213,245}
\definecolor{afp-kobalt-6}{RGB}{219,239,250}
\definecolor{afp-gray-neutral-1}{RGB}{50,50,50}
\definecolor{afp-gray-neutral-2}{RGB}{70,70,70}
\definecolor{afp-gray-neutral-3}{RGB}{101,101,101}
\definecolor{afp-gray-neutral-4}{RGB}{147,147,147}
\definecolor{afp-gray-neutral-5}{RGB}{193,193,193}
\definecolor{afp-gray-neutral-6}{RGB}{238,238,238}
\definecolor{afp-gray-warm-1}{RGB}{52,49,46}
\definecolor{afp-gray-warm-2}{RGB}{70,66,62}
\definecolor{afp-gray-warm-3}{RGB}{99,93,88}
\definecolor{afp-gray-warm-4}{RGB}{147,139,131}
\definecolor{afp-gray-warm-5}{RGB}{180,170,160}
\definecolor{afp-gray-warm-6}{RGB}{237,224,211}
\definecolor{afp-gray-cold-1}{RGB}{44,48,50}
\definecolor{afp-gray-cold-2}{RGB}{61,67,70}
\definecolor{afp-gray-cold-3}{RGB}{90,97,101}
\definecolor{afp-gray-cold-4}{RGB}{128,141,147}
\definecolor{afp-gray-cold-5}{RGB}{154,169,177}
\definecolor{afp-gray-cold-6}{RGB}{203,222,232}
\definecolor{afp-gray-green-1}{RGB}{45,50,47}
\definecolor{afp-gray-green-2}{RGB}{62,70,66}
\definecolor{afp-gray-green-3}{RGB}{90,101,95}
\definecolor{afp-gray-green-4}{RGB}{131,147,139}
\definecolor{afp-gray-green-5}{RGB}{172,193,182}
\definecolor{afp-gray-green-6}{RGB}{211,237,224}
\definecolor{afp-bluegreen-1}{RGB}{11,37,89}
\definecolor{afp-bluegreen-2}{RGB}{24,59,89}
\definecolor{afp-bluegreen-3}{RGB}{42,81,89}
\definecolor{afp-bluegreen-4}{RGB}{50,115,85}
\definecolor{afp-bluegreen-5}{RGB}{92,140,70}
\definecolor{afp-bluegreen-6}{RGB}{175,199,106}
\definecolor{afp-bluegreen-7}{RGB}{217,225,154}
\definecolor{afp-redorange-1}{RGB}{98,13,19}
\definecolor{afp-redorange-2}{RGB}{123,21,36}
\definecolor{afp-redorange-3}{RGB}{154,33,33}
\definecolor{afp-redorange-4}{RGB}{191,57,31}
\definecolor{afp-redorange-5}{RGB}{217,101,24}
\definecolor{afp-redorange-6}{RGB}{233,146,91}
\definecolor{afp-redorange-7}{RGB}{241,214,181}
\definecolor{afp-pastel-emerald-dark}{RGB}{9,102,55}
\definecolor{afp-pastel-emerald-medium}{RGB}{0,155,78}
\definecolor{afp-pastel-emerald-light}{RGB}{86,205,132}
\definecolor{afp-pastel-lime-dark}{RGB}{71,102,9}
\definecolor{afp-pastel-lime-medium}{RGB}{96,142,5}
\definecolor{afp-pastel-lime-light}{RGB}{156,206,56}
\definecolor{afp-pastel-yellow-dark}{RGB}{193,135,17}
\definecolor{afp-pastel-yellow-medium}{RGB}{236,172,45}
\definecolor{afp-pastel-yellow-light}{RGB}{255,210,101}
\definecolor{afp-pastel-orange-dark}{RGB}{191,92,6}
\definecolor{afp-pastel-orange-medium}{RGB}{249,148,59}
\definecolor{afp-pastel-orange-light}{RGB}{255,182,120}
\definecolor{afp-pastel-red-dark}{RGB}{177,30,19}
\definecolor{afp-pastel-red-medium}{RGB}{232,62,52}
\definecolor{afp-pastel-red-light}{RGB}{253,125,125}
\definecolor{afp-pastel-pink-dark}{RGB}{155,47,92}
\definecolor{afp-pastel-pink-medium}{RGB}{205,87,137}
\definecolor{afp-pastel-pink-light}{RGB}{248,145,189}
\definecolor{afp-pastel-violet-dark}{RGB}{92,45,153}
\definecolor{afp-pastel-violet-medium}{RGB}{138,87,205}
\definecolor{afp-pastel-violet-light}{RGB}{193,145,248}
\definecolor{afp-pastel-kobalt-dark}{RGB}{28,24,150}
\definecolor{afp-pastel-kobalt-medium}{RGB}{62,54,222}
\definecolor{afp-pastel-kobalt-light}{RGB}{169,166,255}
\definecolor{afp-pastel-blue-dark}{RGB}{0,86,148}
\definecolor{afp-pastel-blue-medium}{RGB}{19,127,204}
\definecolor{afp-pastel-blue-light}{RGB}{178,213,252}
\definecolor{afp-pastel-cyan-dark}{RGB}{0,95,106}
\definecolor{afp-pastel-cyan-medium}{RGB}{27,154,170}
\definecolor{afp-pastel-cyan-light}{RGB}{137,233,246}
\definecolor{afp-pastel-gray-dark}{RGB}{75,75,75}
\definecolor{afp-pastel-gray-medium}{RGB}{145,145,145}
\definecolor{afp-pastel-gray-light}{RGB}{190,190,190}
\definecolor{afp-cold}{RGB}{46,44,184}
\definecolor{afp-hot}{RGB}{219,0,43}
\definecolor{afp-reference}{RGB}{145,145,145}
\definecolor{afp-growth}{RGB}{31,138,112}
\definecolor{afp-warning}{RGB}{253,116,0}
\definecolor{afp-highlight}{RGB}{92,45,153}
\definecolor{afp-mark-error}{RGB}{177,30,19}
\definecolor{afp-mark-note}{RGB}{62,54,222}
\definecolor{afp-mark-answer}{RGB}{0,155,78}
\definecolor{afp-mark-muted}{RGB}{147,147,147}
\definecolor{cset-aps-blue}{RGB}{46,44,184}
\definecolor{cset-aps-red}{RGB}{219,0,43}
\definecolor{cset-aps-green}{RGB}{31,138,112}
\definecolor{cset-aps-orange}{RGB}{253,116,0}
\definecolor{cset-aps-blueberry}{RGB}{28,128,158}
\definecolor{cset-aps-limegreen}{RGB}{190,219,67}
\definecolor{cset-aps-yellow}{RGB}{255,225,25}
\definecolor{cset-aps-turquoise}{RGB}{0,67,88}
\definecolor{cset-aps-kobalt-medium}{RGB}{66,95,239}
\definecolor{cset-aps-kobalt-dark}{RGB}{24,6,160}
\definecolor{cset-aps-my-label-red}{RGB}{177,30,19}
\definecolor{cset-aps-my-label-blue}{RGB}{62,54,222}
\definecolor{cset-aps-my-label-gray}{RGB}{147,147,147}
`;
