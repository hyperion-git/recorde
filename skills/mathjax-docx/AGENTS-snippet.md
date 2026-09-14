<!-- Paste into a project's AGENTS.md / CLAUDE.md that produces Word documents. -->
## Equations in Word documents

Write formulas as placeholders (`[[math: …]]` inline, `[[display: …]]`,
`[[eq#label: …]]` numbered, `[[ref: label]]`) while building the .docx, then run
`node <recorde-repo>/headless/bin/mjx-docx.mjs process file.docx` followed by
`check` and `preview`. Full procedure and grammar: `skills/mathjax-docx/SKILL.md`
in the Recorde repo. Never emit pandoc `$…$` math (OMML) or hand-written
drawing XML for equations.
