"""Asendab kohatäited jooniste ala (wpc) jooniste ja SEQ-pealdiseväljadega."""
import json, re, sys, zipfile, shutil, os
sys.path.insert(0, os.path.dirname(__file__))
import figures

src, dst = sys.argv[1], sys.argv[2]
sh = json.load(open(os.path.join(os.path.dirname(__file__), "stakeholders.json")))
figs = {
    1: figures.problem_tree(),
    2: figures.objective_tree(),
    3: figures.stakeholder_matrix(sh),
    4: figures.program_figure(),
}

zin = zipfile.ZipFile(src)
xml = zin.read("word/document.xml").decode("utf8")

for n, c in figs.items():
    pat = re.compile(r"<w:p>(?:(?!</w:p>).)*?§§FIG:%d§§.*?</w:p>" % n, re.S)
    xml, k = pat.subn(lambda m: c.xml(1000 + n), xml, count=1)
    assert k == 1, n

counters = {}
def seq(m):
    rpr, kind = m.group(1), m.group(2)
    counters[kind] = counters.get(kind, 0) + 1
    return (f'<w:fldSimple w:instr=" SEQ {kind} \\* ARABIC "><w:r>{rpr}'
            f'<w:t>{counters[kind]}</w:t></w:r></w:fldSimple>')
xml, k = re.subn(r'<w:r>(<w:rPr>(?:(?!</w:rPr>).)*</w:rPr>)<w:t xml:space="preserve">§SEQ:(\w+)§</w:t></w:r>', seq, xml)
assert "§" not in xml, "kohatäide jäi alles"
print("pealdised:", counters)

# nimeruumid (wpc, wps, a) peavad olema juurelemendis deklareeritud
root = re.search(r"<w:document[^>]*>", xml).group(0)
for pre, uri in [("wpc", "http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas"),
                 ("wps", "http://schemas.microsoft.com/office/word/2010/wordprocessingShape"),
                 ("wp", "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing")]:
    if f"xmlns:{pre}=" not in root:
        xml = xml.replace(root, root[:-1] + f' xmlns:{pre}="{uri}">', 1)
        root = re.search(r"<w:document[^>]*>", xml).group(0)

zout = zipfile.ZipFile(dst, "w", zipfile.ZIP_DEFLATED)
for item in zin.infolist():
    data = zin.read(item.filename)
    if item.filename == "word/document.xml":
        data = xml.encode("utf8")
    zout.writestr(item, data)
zout.close()
print("ok", dst)
