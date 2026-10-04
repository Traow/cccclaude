"""Joonised Wordi jooniste alana (wpc) – kastid + ühendajad (connector).

Kõik mõõdud sentimeetrites, teisendatakse EMU-deks (1 cm = 360 000 EMU).
Joonise laius = teksti laius (16 cm).
"""
from xml.sax.saxutils import escape

CM = 360000
FONT = "Times New Roman"


# Teksti laius on juhendi veeriste korral 15 cm; joonised on kavandatud 16 cm laiusele
SCALE = 15.0 / 16.0
# Juhend ei soovita joonistel taustavärve – kastid on valge taustaga, eristus raamivärviga
NO_BG = True


def emu(v):
    return int(round(v * CM * SCALE))


class Canvas:
    def __init__(self, name, width, height):
        self.name = name
        self.w = width
        self.h = height
        self.items = []
        self.next_id = 2
        self.boxes = {}

    def _id(self):
        i = self.next_id
        self.next_id += 1
        return i

    # --- kujundid -------------------------------------------------------
    def box(self, key, cx, y, w, h, text, fill="DEEAF6", line="2F5496",
            bold=False, size=10, color="000000", dash=None, geom="rect",
            line_w=9525, align="center", vert="horz", rot=0):
        """Ristkülik tekstiga. cx = keskpunkti x, y = ülaserv."""
        sid = self._id()
        if fill and NO_BG:
            fill = "FFFFFF"
        x = cx - w / 2
        self.boxes[key] = dict(id=sid, x=x, y=y, w=w, h=h)
        paras = text.split("\n")
        ptxt = ""
        for p in paras:
            ptxt += (
                f'<w:p><w:pPr><w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="auto"/>'
                f'<w:jc w:val="{align}"/></w:pPr>'
                f'<w:r><w:rPr><w:rFonts w:ascii="{FONT}" w:hAnsi="{FONT}" w:cs="{FONT}"/>'
                f'{"<w:b/><w:bCs/>" if bold else ""}<w:color w:val="{color}"/>'
                f'<w:sz w:val="{size*2}"/><w:szCs w:val="{size*2}"/></w:rPr>'
                f'<w:t xml:space="preserve">{escape(p)}</w:t></w:r></w:p>'
            )
        fill_xml = (f'<a:solidFill><a:srgbClr val="{fill}"/></a:solidFill>'
                    if fill else "<a:noFill/>")
        if line:
            dash_xml = f'<a:prstDash val="{dash}"/>' if dash else ""
            line_xml = (f'<a:ln w="{line_w}"><a:solidFill><a:srgbClr val="{line}"/></a:solidFill>'
                        f'{dash_xml}</a:ln>')
        else:
            line_xml = "<a:ln><a:noFill/></a:ln>"
        rot_attr = f' rot="{rot}"' if rot else ""
        av = '<a:avLst><a:gd name="adj" fmla="val 12000"/></a:avLst>' if geom == "roundRect" else "<a:avLst/>"
        self.items.append(
            f'<wps:wsp><wps:cNvPr id="{sid}" name="{escape(key)}"/><wps:cNvSpPr/>'
            f'<wps:spPr><a:xfrm{rot_attr}><a:off x="{emu(x)}" y="{emu(y)}"/><a:ext cx="{emu(w)}" cy="{emu(h)}"/></a:xfrm>'
            f'<a:prstGeom prst="{geom}">{av}</a:prstGeom>{fill_xml}{line_xml}</wps:spPr>'
            f'<wps:txbx><w:txbxContent>{ptxt}</w:txbxContent></wps:txbx>'
            f'<wps:bodyPr rot="0" vert="{vert}" wrap="square" lIns="18000" tIns="18000" rIns="18000" bIns="18000" '
            f'anchor="ctr" anchorCtr="0"><a:noAutofit/></wps:bodyPr></wps:wsp>'
        )
        return sid

    def label(self, key, cx, y, w, h, text, size=10, bold=False, align="center", color="000000"):
        return self.box(key, cx, y, w, h, text, fill=None, line=None, size=size,
                        bold=bold, align=align, color=color)

    def _ln(self, color, width, arrow_end, arrow_start=False, dash=None):
        tail = '<a:tailEnd type="triangle" w="med" len="med"/>' if arrow_end else ""
        head = '<a:headEnd type="triangle" w="med" len="med"/>' if arrow_start else ""
        d = f'<a:prstDash val="{dash}"/>' if dash else ""
        return (f'<a:ln w="{width}"><a:solidFill><a:srgbClr val="{color}"/></a:solidFill>'
                f'{d}{head}{tail}</a:ln>')

    def connect_up(self, child, parent, color="404040", width=12700):
        """Nurkühendaja (bentConnector3) lapse ülaservast vanema alaservani.

        Ühendaja on mõlemast otsast kastide külge kinnitatud (stCxn/endCxn),
        nii et Wordis kaste nihutades liiguvad ka ühendajad kaasa.
        Ristküliku ühenduspunktid: 0 = üles, 1 = vasak, 2 = alla, 3 = parem.
        """
        c = self.boxes[child]
        p = self.boxes[parent]
        x1, y1 = c["x"] + c["w"] / 2, c["y"]
        x2, y2 = p["x"] + p["w"] / 2, p["y"] + p["h"]
        cid = self._id()
        cxn = (f'<wps:cNvCnPr><a:stCxn id="{c["id"]}" idx="0"/>'
               f'<a:endCxn id="{p["id"]}" idx="2"/></wps:cNvCnPr>')
        H = y1 - y2
        W = abs(x2 - x1)
        if W < 0.01:
            xfrm = (f'<a:xfrm flipV="1"><a:off x="{emu(x1)}" y="{emu(y2)}"/>'
                    f'<a:ext cx="0" cy="{emu(H)}"/></a:xfrm>')
            geom = '<a:prstGeom prst="straightConnector1"><a:avLst/></a:prstGeom>'
        else:
            mx, my = (x1 + x2) / 2, (y1 + y2) / 2
            flip = ' flipV="1"' if x2 < x1 else ""
            xfrm = (f'<a:xfrm rot="16200000"{flip}><a:off x="{emu(mx - H / 2)}" y="{emu(my - W / 2)}"/>'
                    f'<a:ext cx="{emu(H)}" cy="{emu(W)}"/></a:xfrm>')
            geom = ('<a:prstGeom prst="bentConnector3"><a:avLst>'
                    '<a:gd name="adj1" fmla="val 50000"/></a:avLst></a:prstGeom>')
        self.items.append(
            f'<wps:wsp><wps:cNvPr id="{cid}" name="Ühendaja {cid}"/>{cxn}'
            f'<wps:spPr>{xfrm}{geom}{self._ln(color, width, True)}</wps:spPr>'
            f'<wps:bodyPr/></wps:wsp>'
        )

    def line(self, x1, y1, x2, y2, color="404040", width=12700, arrow=True, dash=None,
             both=False):
        """Sirge (kinnitamata) ühendaja, nt teljed."""
        lid = self._id()
        flips = ""
        if x2 < x1:
            flips += ' flipH="1"'
        if y2 < y1:
            flips += ' flipV="1"'
        xfrm = (f'<a:xfrm{flips}><a:off x="{emu(min(x1, x2))}" y="{emu(min(y1, y2))}"/>'
                f'<a:ext cx="{emu(abs(x2 - x1))}" cy="{emu(abs(y2 - y1))}"/></a:xfrm>')
        self.items.append(
            f'<wps:wsp><wps:cNvPr id="{lid}" name="Joon {lid}"/><wps:cNvCnPr/>'
            f'<wps:spPr>{xfrm}<a:prstGeom prst="straightConnector1"><a:avLst/></a:prstGeom>'
            f'{self._ln(color, width, arrow, both, dash)}</wps:spPr><wps:bodyPr/></wps:wsp>'
        )

    def xml(self, doc_pr_id):
        return (
            '<w:p><w:pPr><w:keepNext/><w:spacing w:before="120" w:after="0" w:line="240" w:lineRule="auto"/>'
            '<w:ind w:firstLine="0"/><w:jc w:val="center"/></w:pPr><w:r><w:rPr><w:noProof/></w:rPr><w:drawing>'
            f'<wp:inline distT="0" distB="0" distL="0" distR="0"><wp:extent cx="{emu(self.w)}" cy="{emu(self.h)}"/>'
            '<wp:effectExtent l="0" t="0" r="0" b="0"/>'
            f'<wp:docPr id="{doc_pr_id}" name="{escape(self.name)}"/><wp:cNvGraphicFramePr/>'
            '<a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">'
            '<a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas">'
            '<wpc:wpc><wpc:bg/><wpc:whole/>'
            + "".join(self.items)
            + '</wpc:wpc></a:graphicData></a:graphic></wp:inline></w:drawing></w:r></w:p>'
        )


# ---------------------------------------------------------------------------
# Joonis 1 – probleemipuu
# ---------------------------------------------------------------------------
EFF2 = "FBE5D6"  # tagajärjed
EFF_L = "C55A11"
CORE = "F4B183"
CORE_L = "843C0C"
CAUSE = "DEEAF6"
CAUSE_L = "2F5496"
ROOT = "F2F2F2"
ROOT_L = "595959"


def tree(c, rows, links, colors):
    """rows: list of (row_y, height, [(key, cx, w, text, style)])"""
    for y, h, boxes in rows:
        for key, cx, w, text, style in boxes:
            fill, line, bold, size, dash = colors[style]
            c.box(key, cx, y, w, h, text, fill=fill, line=line, bold=bold, size=size, dash=dash,
                  line_w=28575 if style == "core" else 12700)
    for child, parent in links:
        c.connect_up(child, parent)


def problem_tree():
    c = Canvas("Probleemipuu", 16.0, 13.7)
    colors = {
        "eff": (EFF2, EFF_L, False, 11, None),
        "core": (CORE, CORE_L, True, 11, None),
        "cause": (CAUSE, CAUSE_L, False, 11, None),
        "root": (ROOT, ROOT_L, False, 11, None),
    }
    L = [2.70, 8.00, 13.30]
    R = [1.35, 4.05, 6.65, 9.35, 11.95, 14.65]
    rows = [
        (0.0, 1.70, [
            ("T4", 4.50, 6.2, "T4. Merchandise'i ja piletimüügi potentsiaal jääb osaliselt kasutamata", "eff"),
            ("T5", 11.50, 6.2, "T5. Klubi maine ja brändi väärtus kahanevad pikas plaanis", "eff"),
        ]),
        (2.45, 1.80, [
            ("T1", L[0], 4.7, "T1. Traditsioonilised fännid kritiseerivad brändi ega samastu sellega", "eff"),
            ("T2", L[1], 4.7, "T2. Fännibaas jaguneb „vanaks“ ja „uueks“ kogukonnaks", "eff"),
            ("T3", L[2], 4.7, "T3. Emotsionaalne side ja lojaalsus klubiga nõrgenevad", "eff"),
        ]),
        (5.00, 1.40, [
            ("CORE", 8.0, 12.6, "PÕHIPROBLEEM: 2017. aasta rebrändingu järel on Juventuse fännibaas killustunud ega samastu ühtse identiteediga", "core"),
        ]),
        (7.15, 1.80, [
            ("P1", L[0], 4.7, "P1. Uus identiteet katkestas sideme klubi ajaloo ja sümboolikaga", "cause"),
            ("P2", L[1], 4.7, "P2. Fänne ei kaasatud rebrändingu protsessi", "cause"),
            ("P3", L[2], 4.7, "P3. Brändistrateegia eelistab globaalset ärilist sihtrühma", "cause"),
        ]),
        (9.75, 2.75, [
            ("P1a", R[0], 2.5, "P1a. Ovaalne vapp ja Torino härg kadusid", "root"),
            ("P1b", R[1], 2.5, "P1b. „J“-logo tundub fännidele korporatiivne", "root"),
            ("P2a", R[2], 2.5, "P2a. Muudatus esitleti valmis lahendusena", "root"),
            ("P2b", R[3], 2.5, "P2b. Puudub fännide alaline tagasiside­kanal", "root"),
            ("P3a", R[4], 2.5, "P3a. Fookus elustiili- ja moe­toodetel", "root"),
            ("P3b", R[5], 2.5, "P3b. Sõnumid ei väärtusta klubi pärandit", "root"),
        ]),
    ]
    links = [("P1a", "P1"), ("P1b", "P1"), ("P2a", "P2"), ("P2b", "P2"), ("P3a", "P3"), ("P3b", "P3"),
             ("P1", "CORE"), ("P2", "CORE"), ("P3", "CORE")]
    tree(c, rows, links, colors)
    for t in ("T1", "T2", "T3"):
        c.connect_up("CORE", t)
    c.connect_up("T1", "T4")
    c.connect_up("T2", "T4")
    c.connect_up("T2", "T5")
    c.connect_up("T3", "T5")
    c.label("lg", 8.0, 12.7, 16.0, 0.9,
            "Oranž raam – tagajärjed; paks raam – põhiprobleem; sinine raam – otsesed põhjused; hall raam – algpõhjused. Nool: põhjus → tagajärg.",
            size=10, color="404040")
    return c


# ---------------------------------------------------------------------------
# Joonis 2 – eesmärgipuu
# ---------------------------------------------------------------------------
def objective_tree():
    c = Canvas("Eesmärgipuu", 16.0, 13.7)
    IN, IN_L = "E2EFDA", "548235"
    CORE_G, CORE_GL = "A9D18E", "385723"
    OUT, OUT_L = "FFFFFF", "7F7F7F"
    colors = {
        "end": ("FFF2CC", "BF9000", False, 11, None),
        "core": (CORE_G, CORE_GL, True, 11, None),
        "in": (IN, IN_L, False, 11, None),
        "out": (OUT, OUT_L, False, 11, "dash"),
    }
    L = [2.70, 8.00, 13.30]
    R = [1.35, 4.05, 6.65, 9.35, 11.95, 14.65]
    rows = [
        (0.0, 1.70, [
            ("E4", 4.50, 6.2, "E4. Merchandise'i ja piletimüügi potentsiaal on täielikult kasutatud", "end"),
            ("E5", 11.50, 6.2, "E5. Klubi maine ja brändi väärtus kasvavad", "end"),
        ]),
        (2.45, 1.80, [
            ("E1", L[0], 4.7, "E1. Traditsioonilised fännid samastuvad klubi brändiga", "end"),
            ("E2", L[1], 4.7, "E2. Kohalikud ja globaalsed fännid on üks kogukond", "end"),
            ("E3", L[2], 4.7, "E3. Emotsionaalne side ja lojaalsus klubiga on tugevad", "end"),
        ]),
        (5.00, 1.40, [
            ("CORE", 8.0, 12.6, "PEAEESMÄRK: Juventuse fännibaas on taas ühendatud ja samastub klubi identiteediga", "core"),
        ]),
        (7.15, 1.80, [
            ("O1", L[0], 4.7, "O1. Identiteet ühendab klubi ajaloo ja kaasaegsuse", "in"),
            ("O2", L[1], 4.7, "O2. Fännid osalevad brändi kujundamises", "in"),
            ("O3", L[2], 4.7, "O3. Bränd teenib nii kohalikke kui ka globaalseid fänne", "in"),
        ]),
        (9.75, 2.75, [
            ("O1a", R[0], 2.5, "O1a. Ajalooline vapp on taas kasutusel", "in"),
            ("O1b", R[1], 2.5, "O1b. Põhilogo on fännide seas aktsepteeritud", "out"),
            ("O2a", R[2], 2.5, "O2a. Lahendus on fännidega koos testitud", "in"),
            ("O2b", R[3], 2.5, "O2b. Toimib alaline fännide nõukogu", "out"),
            ("O3a", R[4], 2.5, "O3a. Fännipoes on pärand­kollektsioon", "in"),
            ("O3b", R[5], 2.5, "O3b. Sõnumid väärtustavad pärandit", "in"),
        ]),
    ]
    links = [("O1a", "O1"), ("O1b", "O1"), ("O2a", "O2"), ("O2b", "O2"), ("O3a", "O3"), ("O3b", "O3"),
             ("O1", "CORE"), ("O2", "CORE"), ("O3", "CORE")]
    tree(c, rows, links, colors)
    for t in ("E1", "E2", "E3"):
        c.connect_up("CORE", t)
    c.connect_up("E1", "E4")
    c.connect_up("E2", "E4")
    c.connect_up("E2", "E5")
    c.connect_up("E3", "E5")
    c.label("lg", 8.0, 12.7, 16.0, 0.9,
            "Kollane raam – soovitud püsiseisund (eesmärgid); roheline raam – vahendid projekti ulatuses; "
            "kriipsjoon – vahendid väljaspool projekti (programmi teised projektid). Nool: vahend → eesmärk.",
            size=10, color="404040")
    return c


# ---------------------------------------------------------------------------
# Joonis 3 – huvipoolte mõju-huvi maatriks
# ---------------------------------------------------------------------------
def stakeholder_matrix(stakeholders):
    c = Canvas("Huvipoolte maatriks", 16.0, 10.9)
    ox, oy = 1.0, 0.75     # joonise ala vasak ülanurk
    W, H = 14.6, 9.4
    hw, hh = W / 2, H / 2
    quads = [
        (ox, oy, "HOIA RAHUL", "FFF2CC", "top"),
        (ox + hw, oy, "VÕTMEISIKUD", "F8CBAD", "top"),
        (ox, oy + hh, "JÄLGI", "E2EFDA", "bottom"),
        (ox + hw, oy + hh, "HOIA INFORMEERITUNA", "DEEAF6", "bottom"),
    ]
    for i, (x, y, t, f, where) in enumerate(quads):
        c.box(f"Q{i}", x + hw / 2, y, hw, hh, "", fill=f, line="7F7F7F", line_w=12700)
        ly = y + 0.08 if where == "top" else y + hh - 0.6
        c.label(f"QL{i}", x + hw / 2, ly, hw, 0.5, t, size=11, bold=True, color="404040")
    c.line(ox, oy + H, ox + W + 0.3, oy + H, width=15875)
    c.line(ox, oy + H, ox, oy - 0.6, width=15875)
    c.label("xl", ox + W / 2, oy + H + 0.1, 8.0, 0.5, "HUVI projekti vastu (1–5)", size=11, bold=True)
    c.label("yl", ox + 2.8, 0.0, 5.2, 0.5, "MÕJU projektile (1–5)", size=11, bold=True, align="left")
    c.label("lo", ox + 0.35, oy + H + 0.1, 0.6, 0.5, "1", size=11)
    c.label("hi", ox + W - 0.35, oy + H + 0.1, 0.6, 0.5, "5", size=11)

    # skaala: 1–3 madal pool, 4–5 kõrge pool
    frac = {1: 0.08, 2: 0.25, 3: 0.41, 4: 0.64, 5: 0.86}
    for s in stakeholders:
        text = f'{s["nr"]}. {s["short"]}'
        w = len(text) * 0.195 + 0.5
        x = ox + W * frac[s["interest"]] + s.get("dx", 0)
        x = min(max(x, ox + w / 2 + 0.15), ox + W - w / 2 - 0.15)
        y = oy + H * (1 - frac[s["influence"]]) + s.get("dy", 0)
        c.box("S" + str(s["nr"]), x, y - 0.33, w, 0.66, text,
              fill="FFFFFF", line="404040", size=11, geom="roundRect")
    return c


# ---------------------------------------------------------------------------
# Joonis 4 – strateegia, programm, projekt ja portfell
# ---------------------------------------------------------------------------
def program_figure():
    c = Canvas("Programm ja portfell", 16.0, 9.0)
    c.box("S", 8.0, 0.0, 12.0, 1.1,
          "STRATEEGIA: Juventus kui globaalne jalgpalli- ja elustiilibränd („Black and White and More“, 2017)",
          fill="D9D9D9", line="404040", bold=True, size=11)
    c.box("PR", 8.0, 1.9, 12.0, 1.1,
          "PROGRAMM „Juventus uuesti ja paremini“ – bränd ja fännisuhted (omanik: klubi juhtkond)",
          fill="DEEAF6", line="2F5496", bold=True, size=11)
    c.box("P1", 2.75, 4.0, 5.1, 1.9,
          "Projekt 1 (käesolev, 3 kuud): pärandvapp, kaasav disain ja lansseerimine",
          fill="E2EFDA", line="548235", bold=True, size=11)
    c.box("P2", 8.0, 4.0, 5.0, 1.9,
          "Projekt 2 (jätkuprojekt): alaline fännide nõukogu ja kogukonnaplatvorm",
          fill="FFFFFF", line="7F7F7F", size=11, dash="dash")
    c.box("P3", 13.25, 4.0, 5.1, 1.9,
          "Projekt 3: pärandkollektsiooni täismahus tootmine ja litsentsimine",
          fill="FFFFFF", line="7F7F7F", size=11, dash="dash")
    c.box("PF", 8.0, 6.9, 15.6, 1.9,
          "PORTFELL (ressursid ja teenused): klubi turundus- ja brändieelarve, kaubamärgiportfell, "
          "litsentsi- ja merchandise'i teenused, digikanalid; inimressurss: 8 spetsialistirolli, "
          "partnerid (tehniline partner, litsentsiaadid)",
          fill="FBE5D6", line="C55A11", size=11)
    c.connect_up("PR", "S")
    for p in ("P1", "P2", "P3"):
        c.connect_up(p, "PR")
        c.connect_up("PF", p)
    return c
