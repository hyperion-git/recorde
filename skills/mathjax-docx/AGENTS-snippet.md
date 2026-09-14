<!-- Paste into a project's AGENTS.md / CLAUDE.md that produces Word documents. -->
## Equations in Word documents

Write formulas as placeholders (`[[math: …]]` inline, `[[display: …]]`,
`[[eq#label: …]]` numbered, `[[ref: label]]`) while building the .docx, then run
`mjx-docx process file.docx` followed by `mjx-docx check` and `mjx-docx preview`
(`mjx-docx` comes from `npm link` in the Recorde repo,
https://github.com/hyperion-git/recorde). Full procedure and grammar:
`skills/mathjax-docx/SKILL.md` there. Never emit pandoc `$…$` math (OMML) or
hand-written drawing XML for equations.
