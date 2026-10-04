# PJM ülesanne 2 – alapeatükk 1.2

`PJM_T2_alapeatukk_1.2.docx` – rühmatöö (probleemipuu, eesmärgipuu, huvipoolte analüüs, alternatiivide võrdlus, lisad).

Uuesti genereerimine:

```bash
cd build
node build_doc.js stage1.docx                 # tekst ja tabelid (docx-js)
python3 postprocess.py stage1.docx out.docx   # joonised Wordi jooniste alana + SEQ-pealdised
```

Sisu muutmiseks: tekst `build/build_doc.js`, joonised `build/figures.py`,
huvipooled `build/stakeholders.json`, alternatiivide hinded ja kaalud `build/content.js`.
