"""Build the headless test fixtures: agent-style documents with equation
placeholders, made by python-docx and by pandoc — the two builders a CLI agent
actually uses. Run: micromamba run -n sci-base python scripts/make-headless-fixtures.py
Outputs test/fixtures/headless/pydocx-placeholders.docx and pandoc-placeholders.docx."""
import subprocess, pathlib
from docx import Document
from docx.shared import Pt

out = pathlib.Path(__file__).resolve().parent.parent / 'test' / 'fixtures' / 'headless'
out.mkdir(parents=True, exist_ok=True)

doc = Document()
doc.add_heading('Harmonic oscillator', level=1)
p = doc.add_paragraph('The energy ')
r = p.add_run('[[math: E = \\hbar\\omega\\left(n + \\tfrac{1}{2}\\right)]]')
r.font.size = Pt(12)
p.add_run(' is quantised, see [[ref: tise]]. Brackets like f(x) = [a,b]]] survive.')
# a placeholder deliberately split across three runs
p2 = doc.add_paragraph()
p2.add_run('Ground state [[math: \\psi_0(x) = ')
p2.add_run('\\left(\\frac{m\\omega}{\\pi\\hbar}\\right)^{1/4}')
p2.add_run(' e^{-m\\omega x^2/2\\hbar}]] and that is all.')
doc.add_paragraph('[[display: \\langle x \\rangle = \\int_{-\\infty}^{\\infty} \\psi^* x \\psi \\, dx]]')
doc.add_paragraph('[[eq#tise: i\\hbar\\,\\partial_t \\psi = \\hat H \\psi]]')
doc.add_paragraph('Two in one paragraph: [[eq: a^2 + b^2 = c^2]] then text [[eq#euler: e^{i\\pi} + 1 = 0]] tail.')
doc.add_paragraph('Reference to the second: [[ref: euler]].')
t = doc.add_table(rows=1, cols=2)
t.cell(0, 0).text = 'in a cell: [[math: \\alpha]]'
t.cell(0, 1).text = 'plain'
doc.add_paragraph('Broken on purpose: [[math: \\frac{1]]')
doc.save(out / 'pydocx-placeholders.docx')

# Pandoc route: pandoc's markdown reader treats backslashes as escapes and drops
# raw TeX, so placeholders go in code spans (`…`) — pandoc then keeps every byte
# and the equation pass replaces the code run. Same for a display line.
md = r'''# Pandoc route

Inline `[[math: \sqrt{2}]]` and a display:

`[[display: \sum_{k=0}^{n} k = \frac{n(n+1)}{2}]]`

Numbered `[[eq#gauss: \nabla\cdot\mathbf{E} = \rho/\varepsilon_0]]` then `[[ref: gauss]]`.
'''
(out / 'pandoc-placeholders.md').write_text(md)
subprocess.run(['pandoc', '-f', 'markdown', '-o', str(out / 'pandoc-placeholders.docx'),
                str(out / 'pandoc-placeholders.md')], check=True)
print('fixtures written to', out)
