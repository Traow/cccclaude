# PJM ülesanne 2 – „Juventus uuesti ja paremini“, alapeatükk 1.2

`Juventus_uuesti_ja_paremini_T2_1.2.docx` – rühmatöö (probleemipuu, eesmärgipuu, huvipoolte analüüs, alternatiivide võrdlus, lisad).

Uuesti genereerimine:

```bash
cd build
node build_doc.js stage1.docx                 # tekst ja tabelid (docx-js)
python3 postprocess.py stage1.docx out.docx   # joonised Wordi jooniste alana + SEQ-pealdised
```

Sisu muutmiseks: tekst `build/build_doc.js`, joonised `build/figures.py`,
huvipooled `build/stakeholders.json`, alternatiivide hinded ja kaalud `build/content.js`.
