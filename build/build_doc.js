// Genereerib töö DOCX-faili (joonised lisatakse hiljem postprocess.py abil)
const fs = require("fs");
const path = require("path");
const {
  Document, Packer, Paragraph, TextRun, HeadingLevel, AlignmentType, Table, TableRow, TableCell,
  WidthType, BorderStyle, ShadingType, PageBreak, Footer, PageNumber, LevelFormat,
  TableOfContents, VerticalAlign,
} = require("docx");
const { members: M, criteria, alternatives } = require("./content.js");
const stakeholders = JSON.parse(fs.readFileSync(path.join(__dirname, "stakeholders.json"), "utf8"));

const FONT = "Times New Roman";
const TEXT_W = 9071; // 16 cm DXA

// ---------- abifunktsioonid ----------
// Lihtne märgendus: **paks**, *kald*
function runs(text, opts = {}) {
  const out = [];
  const re = /(\*\*[^*]+\*\*|\*[^*]+\*)/g;
  let last = 0;
  let m;
  while ((m = re.exec(text))) {
    if (m.index > last) out.push(new TextRun({ text: text.slice(last, m.index), ...opts }));
    const t = m[0];
    if (t.startsWith("**")) out.push(new TextRun({ ...opts, text: t.slice(2, -2), bold: true }));
    else out.push(new TextRun({ ...opts, text: t.slice(1, -1), italics: true }));
    last = m.index + t.length;
  }
  if (last < text.length) out.push(new TextRun({ text: text.slice(last), ...opts }));
  return out;
}
const P = (text, o = {}) => new Paragraph({ children: runs(text, o.run || {}), ...o.para });
const H1 = (t, pageBreak = true) => new Paragraph({ heading: HeadingLevel.HEADING_1, pageBreakBefore: pageBreak, children: [new TextRun(t)] });
const H2 = (t) => new Paragraph({ heading: HeadingLevel.HEADING_2, children: [new TextRun(t)] });
const H3 = (t) => new Paragraph({ heading: HeadingLevel.HEADING_3, children: [new TextRun(t)] });
const B = (t, level = 0) => new Paragraph({ numbering: { reference: "bullets", level }, children: runs(t) });
const N = (t, ref = "num1") => new Paragraph({ numbering: { reference: ref, level: 0 }, children: runs(t) });
const planned = (t) => new Paragraph({ style: "Planned", children: runs(t) });
const FIG = (n) => new Paragraph({ children: [new TextRun(`§§FIG:${n}§§`)] });

// Pealdis (Caption): SEQ-väli lisatakse postprocess.py abil
function caption(kind, title) {
  return new Paragraph({
    style: "Caption",
    keepNext: kind === "Tabel",
    children: [
      new TextRun({ text: `${kind} `, bold: true }),
      new TextRun({ text: `§SEQ:${kind}§`, bold: true }),
      new TextRun({ text: ". ", bold: true }),
      ...runs(title),
    ],
  });
}
const source = (t) => new Paragraph({ style: "Source", children: runs(t) });

const border = { style: BorderStyle.SINGLE, size: 4, color: "808080" };
const borders = { top: border, bottom: border, left: border, right: border };

function cell(text, width, { header = false, fill, align = AlignmentType.LEFT, bold = false } = {}) {
  const paras = String(text).split("\n").map((line) =>
    new Paragraph({ style: "TableText", alignment: align, children: runs(line, { bold: header || bold }) }));
  return new TableCell({
    borders,
    width: { size: width, type: WidthType.DXA },
    shading: header ? { fill: "D9E2F3", type: ShadingType.CLEAR, color: "auto" }
      : fill ? { fill, type: ShadingType.CLEAR, color: "auto" } : undefined,
    margins: { top: 40, bottom: 40, left: 80, right: 80 },
    verticalAlign: VerticalAlign.CENTER,
    children: paras,
  });
}

function table(widths, header, rows, opts = {}) {
  const total = widths.reduce((a, b) => a + b, 0);
  return new Table({
    width: { size: total, type: WidthType.DXA },
    columnWidths: widths,
    rows: [
      new TableRow({ tableHeader: true, cantSplit: true, children: header.map((h, i) => cell(h, widths[i], { header: true, align: AlignmentType.CENTER })) }),
      ...rows.map((r) => new TableRow({
        cantSplit: true,
        children: r.map((v, i) => {
          const o = typeof v === "object" && v !== null && !Array.isArray(v) ? v : { t: v };
          return cell(o.t, widths[i], { fill: o.fill || (opts.fills && opts.fills(r, i)), align: o.align || (opts.align && opts.align[i]) || AlignmentType.LEFT, bold: o.bold });
        }),
      })),
    ],
  });
}

const fmt = (x) => x.toFixed(2).replace(".", ",");
const gap = () => new Paragraph({ style: "TableText", children: [] });

// ---------- sisu ----------
const children = [];
const PJ = M.PJ, AI = M.AI, EX = M.EX, OM = M.OM;

// TIITELLEHT
const tc = (t, o = {}) => new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 0 }, children: [new TextRun({ text: t, ...o })] });
children.push(
  tc("TARTU ÜLIKOOL"), tc("Pärnu kolledž"), tc("[Õppekava nimi]"),
  new Paragraph({ spacing: { before: 2000 }, alignment: AlignmentType.CENTER, children: [new TextRun(`${PJ}, ${AI}, ${EX}, ${OM}`)] }),
  new Paragraph({ spacing: { before: 1800, after: 240 }, alignment: AlignmentType.CENTER, children: [new TextRun({ text: "„JUVENTUS UUESTI JA PAREMINI“: JUVENTUS FC REBRÄNDIMISE PROJEKT", bold: true, size: 28 })] }),
  tc("Rühmatöö õppeaines „Projektijuhtimise meetodid ja tehnikad“"),
  tc("Ülesanne 2: alapeatükk 1.2 – probleemide, eesmärkide ja huvipoolte analüüs"),
  new Paragraph({ spacing: { before: 2400 }, alignment: AlignmentType.RIGHT, children: [new TextRun("Juhendaja: T. Tamberg")] }),
  new Paragraph({ spacing: { before: 2600 }, alignment: AlignmentType.CENTER, children: [new TextRun("Pärnu 2026")] }),
);

// SISUKORD
children.push(
  new Paragraph({ style: "TocHeading", pageBreakBefore: true, children: [new TextRun("SISUKORD")] }),
  new TableOfContents("Sisukord", { hyperlink: true, headingStyleRange: "1-3" }),
  P("*Sisukord uueneb Wordis faili avamisel (vajadusel: paremklõps sisukorral → Update Field).*", { run: { color: "7F7F7F", size: 20 } }),
);

// SISSEJUHATUS
children.push(
  H1("SISSEJUHATUS"),
  P("Jaanuaris 2017 esitles Itaalia jalgpalliklubi Juventus FC agentuuriga Interbrand loodud uut visuaalset identiteeti, mille keskmes on minimalistlik „J“-logo. Ajalooline ovaalne vapp koos triipude ja Torino härjaga asendati, et klubi saaks kasvada globaalseks elustiilibrändiks ja jõuda ka „laste, naiste ja millenniumlasteni“ (Football Italia, 2017; Design Week, 2017). Fännide reaktsioon oli suures osas eitav: uut logo peeti liiga anonüümseks ja korporatiivseks (Dezeen, 2017; The Drum, 2017). Rühma hinnangul on rebrändingu järel fännibaas killustunud."),
  P("Käesoleva rühmatöö eesmärk on projektijuhi vaatest välja selgitada, **miks** projekti „Juventus uuesti ja paremini“ vaja on, millist püsivat seisundit see peaks looma ning milline lahendus ühendab fännid uuesti nii, et see sobiks kõigile osapooltele. Projekti teostusfaas kestab 3 kuud ja eelarve on ligikaudu 96 800 eurot."),
  P("Töö on üles ehitatud projekti määratluse loogikas: esimene peatükk määratleb projekti, teine kirjeldab planeerimist ning kolmas elluviimist ja lõpetamist (Tamberg, 2022; AXELOS, 2017). Käesolevas ülesandes on terviklikult koostatud alapeatükk 1.2: metoodika, probleemipuu, eesmärgipuu, huvipoolte analüüs ja kaasamise strateegia ning lahendusalternatiivide võrdlus. Teiste peatükkide eeldatav sisu on esitatud kursiivis või esialgsete andmetena. Iga joonise ja tabeli juures on märgitud vastutaja; tööjaotus ja vahetähtajad on lisas 1, õppetunnid lisas 2 ja vestlus tehisaruga lisas 3."),
  P("Projekt on õppejuhtum: see ei ole seotud Juventus FC-ga ega esinda klubi seisukohti. Omaniku rolli täidab rühmas Ragnar Dietrich."),
);

// 1. PROJEKTI MÄÄRATLEMINE
children.push(
  H1("1. PROJEKTI MÄÄRATLEMINE"),
  H2("1.1. Projekti idee ja taust"),
  P("Projekti idee on rebrändida Juventus FC nii, et 2017. aasta identiteedimuutuse järel killustunud fännibaas saaks uuesti ühendatud. Projekti aluseks on klubi 2017. aasta rebränding „Black and White and More“, mille eesmärk oli laiendada klubi tegevust jalgpallist kaugemale – moe, meedia ja merchandise’i valdkonda (Design Week, 2017; It’s Nice That, 2017)."),
  planned(`[Täiendatakse: ärijuhtumi esialgne kirjeldus – strateegiline sobivus, valikute hindamine, ärisuhted, tasuvus ja teostatavus (Tamberg, 2022, slaid 8–9). Vastutaja: ${PJ}.]`),
  H2("1.2. Probleemide, eesmärkide ja huvipoolte analüüs"),
  P("Alapeatüki eesmärk on enne planeerimist veenduda, et valitakse õige projekt. Omanikul võib olla idee „teeme uue logo“, kuid projektijuht peab kindlaks tegema, kas see on eesmärgi saavutamiseks mõistlik ja efektiivne. Probleemiks loetakse lahknevust soovitava ja olemasoleva seisundi vahel **enne projekti**, mitte projekti käigus tekkida võivaid riske (Tamberg, 2022, slaid 11–12). Analüüs liigub järjekorras probleem → põhjuste analüüs → eesmärgid → huvipooled → lahendusideed → lahenduse valik (Tamberg, 2022, slaid 10)."),
);

// 1.2.1 METOODIKA
children.push(
  H3("1.2.1. Metoodika ja kasutatud tehnikad"),
  P("Analüüsi aluseks on loogilise raamistiku lähenemine (*Logical Framework Approach*, LFA), mille järgi koostatakse esmalt probleemipuu, see teisendatakse eesmärgipuuks ja valitakse eesmärgipuust projekti ulatusse kuuluvad harud (European Commission, 2004). Rühm kasutas järgmisi tehnikaid:"),
  N("**Dokumendi- ja meediaanalüüs** – 2017. aasta rebrändingu eesmärkide ja fännide reaktsiooni kaardistamine erialameedia ja klubi teadaannete põhjal (Football Italia, 2017; Dezeen, 2017). Kvantitatiivsed näitajad (näitarvude tehnika) kogub projekti esimesel kuul turu-uuringu analüütik (Tamberg, 2022, slaid 13)."),
  N("**Ajurünnak ja 635-meetod** probleemide, põhjuste ja lahendusideede kogumiseks; kalasaba-diagrammi kategooriaid (sümboolika, protsess, strateegia, kommunikatsioon, tooted) kasutati kontrollnimekirjana (Ishikawa, 1990; Tamberg, 2022, slaid 18 ja 35)."),
  N("**Probleemipuu** põhjus–tagajärg seoste kaardistamiseks ja **probleemi lause** sõnastamiseks (Tamberg, 2022, slaid 16–20)."),
  N("**5 korda miks** üksikute algpõhjuste sügavuse kontrolliks (Ohno, 1988)."),
  N("**Eesmärgipuu**, milles negatiivsed seisundid sõnastati soovitud püsiseisunditeks ja märgiti projekti ulatus (Tamberg, 2022, slaid 31–34)."),
  N("**Huvipoolte analüüs**: kontrollküsimused huvipoolte leidmiseks, huvi- ja mõjuhinnangud skaalal 1–5, mõju-huvi maatriks ning kaasamise tasemed (Mendelow, 1981; Bryson, 2004; IAP2, 2018; Tamberg, 2022, slaid 22–30)."),
  N("**Benchmarking ja kaalutud mitmekriteeriumiline võrdlus** koos tundlikkusanalüüsiga lahendusalternatiivide hindamiseks (Tamberg, 2022, slaid 35–36)."),
  P("**Probleemianalüüsi tehnika valik.** Rühm võrdles tehnikaid (tabel 1) ja arutles, milline neist aitab kõige paremini vastata küsimusele, miks projekti vaja on. Kalasaba-diagramm ja „5 korda miks“ sobivad protsessi- või kvaliteediprobleemi põhjuste leidmiseks, kuid ei näita probleemi tagajärgi ega teisene otse eesmärkideks (Learn Lean Sigma, s.a.; LinkedIn, s.a.). SWOT kirjeldab organisatsiooni sise- ja väliskeskkonda, mitte põhjuslikke seoseid. Fännibaasi killustumine on mitme põhjusega probleem, mis puudutab väga erinevate huvidega osapooli (omanik, sponsorid, kohalikud ja globaalsed fännid). Seetõttu valis rühm **põhitehnikaks probleemipuu**: see näitab nii põhjuseid kui ka tagajärgi (sh ärilisi), seob eri huvipoolte probleemid ühiste algpõhjuste kaudu ja teiseneb otse eesmärgipuuks (European Commission, 2004; Tamberg, 2022, slaid 16–17)."),
  caption("Tabel", "Probleemianalüüsi tehnikate võrdlus"),
  table([1700, 2550, 2550, 2271],
    ["Tehnika", "Tugevused", "Nõrkused", "Kasutus käesolevas töös"],
    [
      ["**Probleemipuu** (LFA)", "Näitab põhjusi ja tagajärgi; teiseneb eesmärgipuuks; sobib mitme huvipoolega probleemile", "Võib lihtsustada vastastikmõjusid; sõltub osalejate teadmistest", "Põhitehnika (joonis 1)"],
      ["Kalasaba-diagramm (Ishikawa)", "Süstemaatilised põhjuste kategooriad; hea ajurünnaku struktuur", "Ainult põhjused, tagajärjed puuduvad; ei teisene eesmärkideks", "Kategooriad kontrollnimekirjana ajurünnakul"],
      ["5 korda miks", "Lihtne ja kiire; viib sümptomist algpõhjuseni", "Lineaarne, üks põhjusahel; oht peatuda liiga vara", "Algpõhjuste sügavuse kontroll (nt P1b)"],
      ["SWOT", "Ülevaade sise- ja väliskeskkonnast", "Ei näita põhjuslikke seoseid", "Keskkonna analüüs alapeatükis 1.3"],
    ]),
  source(`Allikas: autorite koostatud (European Commission, 2004; Ishikawa, 1990; Ohno, 1988; Learn Lean Sigma, s.a.; LinkedIn, s.a.). Vastutaja: ${PJ}.`),
  P(`**Tehisaru ja teaduslike lisaallikate kasutamine.** Rühm kasutas suurt keelemudelit Claude (Anthropic, 2026) töö struktuuri ja analüüsi kavandi koostamiseks, allikate leidmiseks ning valmis versioonile kriitilise tagasiside saamiseks; päringud koostas ja väljundi kvaliteeti kontrollis ${AI}. Tehisaru pakutud väited ja viited kontrolliti algallikatest; huvi- ja mõjuhinnangud ning alternatiivide hinded on rühma ekspertarvamus. Vestlus ja selle põhjal tehtud täiendused on esitatud lisas 3. Spordiklubi brändi ja fännide samastumise mõistmiseks kasutati teaduskirjandust brändiväärtusest (Keller, 1993) ja spordifännide samastumisest klubiga (Wann & Branscombe, 1993).`),
);

// 1.2.2 PROBLEEMIPUU
children.push(
  H3("1.2.2. Probleemipuu ja probleemi lause"),
  P("Põhiprobleem määratleti esmalt erinevate huvipoolte vaatest (Tamberg, 2022, slaid 17): traditsiooniliste fännide jaoks on probleemiks klubi ajaloolise sümboolika kadumine, omaniku ja turundusosakonna jaoks brändi negatiivne vastuvõtt ning kasutamata müügipotentsiaal, sponsorite jaoks ebastabiilne meediapilt. Ühine nimetaja on see, et **2017. aasta rebrändingu järel on fännibaas killustunud ega samastu ühtse identiteediga**. Probleemipuus (joonis 1) on põhiprobleemi all kolm otsest põhjust ja kuus algpõhjust ning selle kohal tagajärjed."),
  FIG(1),
  caption("Joonis", `Juventuse fännibaasi killustumise probleemipuu (autorite koostatud; Dezeen, 2017; Football Italia, 2017 põhjal). Vastutaja: ${EX}`),
  P("Algpõhjuste sügavust kontrolliti tehnikaga „5 korda miks“. Näiteks: *Miks ei samastu traditsioonilised fännid brändiga?* – Logo tundub neile anonüümne ja korporatiivne. *Miks?* – Ovaalne vapp ja Torino härg, mis olid klubi vappides aastakümneid, eemaldati (Dezeen, 2017). *Miks need eemaldati?* – Identiteet loodi eelkõige globaalse elustiilibrändi vajadustest, mitte fännide väärtustest lähtudes (Design Week, 2017). Ahel näitas, et probleem ei ole ainult logo kujundus, vaid **ajaloolise sideme ja kaasamise puudumine**. Seega ei lahendaks probleemi järjekordne „uus logo“ ilma fännide kaasamiseta."),
  P("Rühm märgib, et fännibaasi killustumise ulatust tuleb mõõta. Projekti esimesel kuul viib turu-uuringu analüütik läbi fännide küsitluse ja sotsiaalmeedia analüüsi, mis annab lähtetaseme eesmärgipuu näitajatele."),
  P("**Probleemi lause.** Ilma selle projektita ei saa Juventus oma fänne uuesti ühendada, sest 2017. aasta rebränding eemaldas klubi ajaloolise sümboolika ilma fänne kaasamata ning seab esikohale globaalse ärilise sihtrühma, mistõttu ei samastu suur osa traditsioonilisi toetajaid klubi brändiga, fännibaas jaguneb „vanaks“ ja „uueks“ kogukonnaks ning klubi kaotab lojaalsust, mainet ja müügitulu."),
);

// 1.2.3 EESMÄRGIPUU
children.push(
  H3("1.2.3. Eesmärgipuu ja projekti ulatus"),
  P("Eesmärgipuu (joonis 2) koostati probleemipuu negatiivsete seisundite ümbersõnastamisel soovitud, tulevikus püsivateks seisunditeks – mitte tegevusteks (Tamberg, 2022, slaid 31–34). Näiteks „uus logo on valmis“ on projekti tulem, mitte eesmärk; eesmärk on, et fännid samastuksid klubi identiteediga. Seejärel otsustati, milliste alameesmärkide saavutamine kuulub 3-kuulise projekti ulatusse ja millised jäetakse programmi teistele projektidele (vt alapeatükk 1.2.6)."),
  FIG(2),
  caption("Joonis", `Juventuse fännibaasi taasühendamise eesmärgipuu (autorite koostatud). Vastutaja: ${OM}`),
  P("Peaeesmärgi saavutamist mõõdetakse järgmiste näitajatega (lähtetase mõõdetakse projekti 1. kuul):"),
  B("fännide samastumise indeks (küsitlus, skaala 1–8) tõuseb 6 kuu jooksul pärast lansseerimist vähemalt 1 palli võrra nii kohalike kui ka globaalsete fännide seas (Wann & Branscombe, 1993);"),
  B("brändi kohta tehtud negatiivsete sotsiaalmeediapostituste osakaal väheneb poole võrra võrreldes lähtetasemega;"),
  B("pärandkollektsiooni müük moodustab esimesel hooajal vähemalt 10% merchandise’i käibest;"),
  B("ametlike fänniklubide liikmete arv kasvab."),
  P("Projekti ulatusse kuuluvad alameesmärgid O1a, O2a, O3a ja O3b, mis on saavutatavad 3 kuu ja 96 800-eurose eelarvega. Põhilogo laialdane aktsepteerimine (O1b) ja alaline fännide nõukogu (O2b) jäetakse programmi jätkuprojektidele, sest need vajavad pikemat aega ja klubi juhtimiskorralduse muutmist."),
);

// 1.2.4 HUVIPOOLED
const quadrantOf = (s) => (s.influence >= 4 ? (s.interest >= 4 ? "Võtmeisik" : "Hoia rahul") : (s.interest >= 4 ? "Hoia informeerituna" : "Jälgi"));
const qFill = { "Võtmeisik": "F8CBAD", "Hoia rahul": "FFF2CC", "Hoia informeerituna": "DEEAF6", "Jälgi": "E2EFDA" };
const shRows = stakeholders.map((s) => {
  const quadrant = quadrantOf(s);
  return [
    { t: `${s.nr}. ${s.name}` },
    { t: String(s.interest), align: AlignmentType.CENTER },
    { t: s.wish },
    { t: String(s.influence), align: AlignmentType.CENTER },
    { t: s.mode },
    { t: `**${quadrant}:** ${s.strategy.split(": ").slice(1).join(": ")}`, fill: qFill[quadrant] },
  ];
});
children.push(
  H3("1.2.4. Huvipoolte analüüs ja kaasamise strateegia"),
  P("Huvipooled leiti kontrollküsimuste abil: kelle vaated ja kogemused on asjakohased, kes on otsustajad, kes hakkavad otsuste järgi tegutsema, kelle toetus on edu jaoks oluline, kellel on õigus tulemustest kasu saada ja kes võib tunda end ohustatuna (Tamberg, 2022, slaid 22). Lähtepunktiks olid rühma määratletud huvipooled – omanikud, osanikud, mängijad, toetajad, fännid, sponsorid ning kohalik omavalitsus ja elanikud –, mida täiendati „väravavahtidega“ (kaubamärgiametid, liiga) ning meedia ja litsentsipartneritega (Tamberg, 2022, slaid 24). Huvi ja mõju hinnati skaalal 1–5 rühma konsensuse alusel; hinnangud valideeritakse projekti esimesel kuul intervjuude ja küsitlusega. Tabelis 2 on esitatud huvipoolte huvi tingimused ja kaasamise strateegia, joonisel 3 nende paiknemine mõju-huvi maatriksis (Mendelow, 1981; Eden & Ackermann, 1998)."),
  caption("Tabel", "Huvipoolte analüüs ja kaasamise strateegia"),
  table([1900, 680, 1850, 680, 1650, 2311],
    ["Huvipool", "Huvi (1–5)", "Huvi (osalemise, toetamise) tingimus", "Mõju (1–5)", "Osalemise või mõju viis", "Kaasamise strateegia"],
    shRows),
  source(`Allikas: autorite koostatud (Tamberg, 2022, slaid 28; Mendelow, 1981). Värv vastab maatriksi ruudule joonisel 3. Vastutaja: ${PJ}.`),
  FIG(3),
  caption("Joonis", `Huvipoolte mõju-huvi maatriks (autorite koostatud Mendelow, 1981 ja Tamberg, 2022, slaid 29 põhjal). Vastutaja: ${PJ}`),
  P("**Erinevasuunalised huvid.** Huvipoolte eesmärgid peavad olema ühitatavad (Tamberg, 2022, slaid 21), kuid rebrändingu puhul vastanduvad need tugevalt (tabel 3). Kõige teravam on vastuolu traditsiooniliste fännide, kes soovivad vana vapi tagasitoomist, ning omaniku ja turundusosakonna vahel, kes soovivad kaitsta 2017. aasta investeeringut ja globaalset brändi. Projektijuhi ülesanne on vastuolud varakult nähtavaks teha ja leida kokkulepped, mille korral ühine huvi – tugev ja ühtne fännibaas – kaalub üles erihuvid (Bryson, 2004)."),
  caption("Tabel", "Huvipoolte vastandlikud huvid ja kavandatud kokkulepped"),
  table([2300, 3100, 3671],
    ["Vastuolu", "Huvipoolte soovid", "Huvide ühitamine ja kokkulepe"],
    [
      ["Traditsioon vs globaalne bränd", "Toetajad (5) ja kohalikud fännid (7) soovivad ajaloolist vappi; omanik (1) ja turundus (2) globaalset „J“-brändi", "Kahetasandiline lahendus: „J“ jääb äriliseks kaubamärgiks, pärandvapp tuleb tagasi mängusärgile ja pärandkollektsiooni; kokkulepe kinnitatakse juhtkomitees"],
      ["Kohalikud vs globaalsed fännid", "Kohalikud (7) väärtustavad Torino ajalugu, globaalsed (6) kaasaegset disaini ja digisisu", "Disaini testitakse mõlema rühmaga (fookusgrupid Torinos, veebiküsitlus); kampaania räägib ühist lugu „ajalugu + tulevik“"],
      ["Muutus vs stabiilsus", "Sponsorid (8) ja litsentsipartnerid (10) soovivad stabiilsust ja vanade laovarude läbimüüki", "Üleminekuperiood ja tootmisgraafik lepitakse kokku 1. kuul; pärandkollektsioon on täiendus, mitte asendus"],
      ["Linna sümbol vs klubi kaubamärk", "Torino linn (9) soovib härja sümboli korrektset kasutust; klubi soovib kaubamärgiõigust", "Kaubamärgiõiguse spetsialist selgitab kasutusõiguse ja vajadusel sõlmitakse linnaga kasutuskokkulepe"],
      ["Kulu vs tulu", "Osanikud (3) soovivad kulude põhjendatust; fännid kvaliteetset lahendust", "Eelarve 96 800 € koos reserviga, mõõdetavad näitajad ja etapiviisiline otsustamine"],
    ]),
  source(`Allikas: autorite koostatud. Sulgudes on huvipoole number tabelist 2. Vastutaja: ${PJ}.`),
  P("**Kaasamise strateegia.** Kaasamise intensiivsus valiti vastavalt mõjukuse ja huvi määrale (Mendelow, 1981) ning IAP2 kaasamise spektri tasemetele – informeerimine, konsulteerimine, kaasamine, koostöö ja otsustusõiguse andmine (IAP2, 2018):"),
  B("**Võtmeisikud** (1, 2, 8) – koostöö: juhtkomitee liikmed, kes kinnitavad iga etapi lõpus lahenduse ja ärijuhtumi aktuaalsuse; sponsoritega lepitakse kokku tootmis- ja üleminekugraafik."),
  B("**Hoia rahul** (11, 12) – konsulteerimine: meedia saab eelinfot ja lugusid suhtekorraldusplaani järgi; kaubamärgiametitele esitatakse nõuetekohased taotlused juba 1. kuul, et vältida „väravavahtide“ tekitatud viivitusi (Tamberg, 2022, slaid 24)."),
  B("**Hoia informeerituna** (5, 6, 7, 10) – kaasamine: fännid osalevad fookusgruppides, disaini testimisel ja hääletusel. Eesmärk on tõsta organiseeritud toetajate mõju ja muuta nad projekti toetajateks ehk viia nad võtmeisikute rühma (Tamberg, 2022, slaid 29)."),
  B("**Jälgi** (3, 4, 9, 13) – informeerimine: investorsuhete teade, mängijad kaasatakse kampaaniasse saadikutena, linnaga konsulteeritakse härja sümboli osas ning liiga nõuded kontrollitakse enne särgi kujunduse kinnitamist."),
  P(`Kaasamise eest vastutab projektijuht ${PJ}. Huvipoolte register vaadatakse üle iga kuu lõpus, sest huvipoolte huvi ja mõju võivad projekti käigus muutuda – näiteks võib meedia tähelepanu järsult tõsta fännide mõju.`),
);

// 1.2.5 ALTERNATIIVID
const best = "B+D";
const altRows = alternatives.map((a) => [
  { t: `**${a.id}** – ${a.name}`, fill: a.id === best ? "E2EFDA" : undefined },
  ...a.scores.map((s) => ({ t: String(s), align: AlignmentType.CENTER, fill: a.id === best ? "E2EFDA" : undefined })),
  { t: `**${fmt(a.total)}**`, align: AlignmentType.CENTER, fill: a.id === best ? "E2EFDA" : undefined },
]);
const A_ = Object.fromEntries(alternatives.map((a) => [a.id, a]));
children.push(
  H3("1.2.5. Lahendusideed ja alternatiivide võrdlus"),
  P("Lahendusideed koguti ajurünnaku ja 635-meetodi abil ning benchmarking’u teel (Tamberg, 2022, slaid 35): vaadeldi, kuidas teised klubid on oma sümboolikat muutnud ning millist rolli on seejuures mänginud fännide kaasamine. Ideed rühmitati neljaks põhimõtteliselt erinevaks alternatiiviks, millele lisati kombinatsioon ja nullalternatiiv:"),
  B("**A – ajaloolise ovaalse vapi täielik taastamine**: „J“-logost loobutakse ja naastakse 2017. aasta eelse vapi juurde."),
  B("**B – kahetasandiline bränd**: „J“ jääb äriliseks ja digitaalseks kaubamärgiks, ajalooline vapp (sh triibud ja härg) tuleb tagasi pärandsümbolina mängusärgile ja pärandkollektsiooni."),
  B("**C – „J“-logo ümberkujundamine koos fännidega**: uus logo, mis lisab „J“-le ajaloolisi elemente ja mille disain valitakse fännide hääletusel."),
  B("**D – kogukonna- ja kommunikatsiooniprogramm**: visuaali ei muudeta, kuid käivitatakse fännide kaasamise ja ajaloo väärtustamise kampaania."),
  B("**B+D** – pärandvapp koos fännide kaasamise programmiga (disaini testimine, fännifoorum, kampaania)."),
  B("**0 – nullalternatiiv**: senine bränd jätkub muutmata kujul."),
  P("Alternatiive võrreldi huvipoolte eesmärkide saavutamise määra, teostatavuse, kulu ja riskantsuse lõikes (Tamberg, 2022, slaid 36). Kriteeriumide kaalud lepiti kokku enne hindamist, et vältida kaalude sobitamist eelistatud lahendusele (tabel 4)."),
  caption("Tabel", "Lahendusalternatiivide kaalutud võrdlus (hinded 1–5, 5 = parim)"),
  table([2411, 860, 860, 860, 860, 860, 860, 1500],
    ["Alternatiiv", ...criteria.map((c) => `${c.id}\n${Math.round(c.w * 100)}%`), "Kaalutud summa"],
    altRows),
  source(`Kriteeriumid: ${criteria.map((c) => `${c.id} – ${c.name}`).join("; ")}. Allikas: autorite koostatud. Vastutaja: ${OM} (kaasteostaja ${EX}).`),
  P(`Kaalutud summa järgi on parim kombinatsioon **B+D** (${fmt(A_["B+D"].total)}), sellele järgneb alternatiiv B (${fmt(A_.B.total)}). Vana vapi täielik taastamine (A, ${fmt(A_.A.total)}) rahuldaks küll traditsioonilisi fänne, kuid tühistaks 2017. aasta investeeringu, tekitaks vastuseisu omanikus ja sponsorites ning ei mahu 96 800-eurose eelarve ja 3 kuu sisse, sest kõik kaubamärgid, tooted ja kanalid tuleks ümber teha. Samal põhjusel jääb eelarvest ja ajast välja uue logo loomine (C, ${fmt(A_.C.total)}). Ainult kommunikatsioon (D, ${fmt(A_.D.total)}) on odav ja kiire, kuid ei kõrvalda algpõhjust P1 – ajaloolise sümboolika puudumist.`),
  P(`Tundlikkusanalüüs: kui eelarve kriteeriumi kaal tõsta 30%-ni, jääb B+D endiselt parimaks (${fmt(A_["B+D"].sensCost)}; B ${fmt(A_.B.sensCost)}, D ${fmt(A_.D.sensCost)}). Riski kaalu tõstmine 25%-ni järjestust ei muuda (B+D ${fmt(A_["B+D"].sensRisk)}). Seega on otsus kaalude suhtes stabiilne ning projekti omanikule soovitatakse alternatiivi **B+D**.`),
  P("Alternatiivid täidavad eesmärgipuu alameesmärke erinevalt (tabel 5): A ja B katavad sümboolika haru, C ja D kaasamise ja kommunikatsiooni haru. Ainult kombinatsioon katab kõik kolm haru."),
  caption("Tabel", "Alternatiivide panus eesmärgipuu alameesmärkidesse"),
  table([2411, 860, 860, 860, 860, 860, 860, 1500],
    ["Alternatiiv", "O1a", "O1b", "O2a", "O2b", "O3a", "O3b", "Kokku"],
    [
      ["A – vana vapi taastamine", "1", "0,5", "0", "0", "0,5", "0,5", "2,5"],
      ["B – kahetasandiline bränd", "1", "0,5", "0", "0", "1", "0,5", "3,0"],
      ["C – „J“ ümberkujundamine", "0,5", "1", "1", "0,5", "0", "0,5", "3,5"],
      ["D – kommunikatsioon", "0", "0", "0,5", "0,5", "0", "1", "2,0"],
      ["B+D", "1", "0,5", "1", "0,5", "1", "1", "5,0"],
      ["0 – nullalternatiiv", "0", "0", "0", "0", "0", "0", "0,0"],
    ].map((r) => r.map((v, i) => ({ t: v, align: i ? AlignmentType.CENTER : AlignmentType.LEFT }))),
  ),
  source(`Märkus: 1 – täielik panus, 0,5 – osaline panus, 0 – panus puudub; tähised vastavad joonisele 2. Allikas: autorite koostatud. Vastutaja: ${OM}.`),
);

// 1.2.6 ÄRIJUHTUM, PROGRAMM, PORTFELL
children.push(
  H3("1.2.6. Ärijuhtumi seire, programm ja portfell"),
  P("Projekti omanikul võib olla idee, kuid projektijuhi ülesanne on tagada, et kavandatav töö viib omaniku ja huvipoolte eesmärgile võimalikult efektiivselt lähemale (Tamberg, 2022, slaid 8; AXELOS, 2017). Analüüsi tulemusel tehakse omanikule ettepanek mitte luua täiesti uut logo, vaid taastada klubi ajalooline sümboolika pärandvapina ja kaasata fännid lahenduse loomisse."),
  P("Projekt kuulub programmi „Juventus uuesti ja paremini“, mis teenib klubi strateegilist eesmärki olla nii tugev jalgpalliklubi kui ka globaalne elustiilibränd (joonis 4). Programm võimaldab ühitada mitut omavahel seotud projekti: käesolev 3-kuuline projekt loob pärandvapi ja lansseerib selle, jätkuprojektid loovad alalise fännide nõukogu ja pärandkollektsiooni täismahus tootmise. Portfelli tasandil otsustatakse, milliseid klubi ressursse (brändi- ja turunduseelarve, kaubamärgiportfell, litsentsi- ja merchandise’i teenused, digikanalid) projektidele eraldatakse (PMI, 2021; ISO, 2012)."),
  FIG(4),
  caption("Joonis", `Projekti seos strateegia, programmi ja portfelliga (autorite koostatud). Vastutaja: ${PJ}`),
  P("Ärijuhtumi aktuaalsust ja projekti sisu adekvaatsust jälgitakse kogu projekti vältel paindlikult:"),
  B("iga kuu lõpus (PRINCE2 etapipiir) vaatab juhtkomitee üle probleemi- ja eesmärgipuu, huvipoolte registri ja näitajad ning otsustab jätkamise, muutmise või lõpetamise (AXELOS, 2017);"),
  B("1. kuu turu-uuringu tulemused on otsustuspunkt: kui fännid ei toeta pärandvapi ideed, hinnatakse alternatiivid uuesti;"),
  B("disaini testitakse fännidega iteratiivselt (kaks vooru) ja lahendust kohandatakse tagasiside põhjal;"),
  B("kui väliskeskkond muutub (nt omaniku strateegia, sponsorlepingud või kaubamärgiõiguslik takistus), muudetakse vajadusel projekti ulatust või programmi koosseisu."),
);

// 1.3–3
const team = [
  ["Brändistrateeg", "Brändi arhitektuur (põhilogo + pärandvapp), sõnumid", "1.–2. kuu"],
  ["Graafiline disainer", "Pärandvapi ja brändiraamatu kujundus", "1.–2. kuu"],
  ["Kaubamärgiõiguse spetsialist", "Õiguskontroll, registreerimine, Torino härja kasutusõigus", "1.–3. kuu"],
  ["Digitaalse sisu spetsialist", "Digikampaania, sotsiaalmeedia ja veebisisu", "2.–3. kuu"],
  ["Merchandise’i tootmiskoordinaator", "Pärandkollektsiooni prototüübid, koostöö litsentsiaatidega", "2.–3. kuu"],
  ["Sisekommunikatsiooni töötaja", "Töötajate ja mängijate informeerimine, brändisaadikud", "1.–3. kuu"],
  ["Turu-uuringu analüütik", "Fännide küsitlus, fookusgrupid, lähtetase ja järelmõõtmine", "1. ja 3. kuu"],
  ["Suhtekorraldusjuht", "Meediasuhted, lansseerimine, kriisikommunikatsioon", "1.–3. kuu"],
];
const budget = [
  ["Turu-uuring ja fännide kaasamine", 9800],
  ["Brändistrateegia ja kontseptsioon", 12000],
  ["Visuaalne identiteet (disain, brändiraamat)", 22000],
  ["Kaubamärgi õiguskaitse", 11000],
  ["Merchandise’i prototüübid ja tootmise ettevalmistus", 14000],
  ["Lansseerimiskampaania, digisisu ja suhtekorraldus", 19000],
  ["Reserv ootamatute kulude katteks", 9000],
];
const eur = (n) => n.toLocaleString("fr-FR").replace(/ | /g, " ") + " €";
const btotal = budget.reduce((s, b) => s + b[1], 0);
children.push(
  H2("1.3. Keskkonna ja esmaste riskide analüüs"),
  planned(`[Eeldatav sisu: PESTLE ja SWOT; esmaste riskide register (nt fännide uus vastureaktsioon, kaubamärgiõiguslik takistus, sponsori vastuseis, ajakava nihkumine), maandamine ja väljumisstrateegia (Tamberg, 2022, slaid 11). Vastutaja: ${EX}.]`),
  H2("1.4. Projekti määratlus ja projektiettepanek"),
  P("Projekti nimi on „Juventus uuesti ja paremini“. Projekti teostusfaas kestab 3 kuud, eelarve on ligikaudu 96 800 eurot. Projekti meeskonnas on kaheksa spetsialistirolli (tabel 6)."),
  caption("Tabel", "Projekti meeskonna rollid (esialgne)"),
  table([2600, 4671, 1800],
    ["Roll", "Vastutus", "Kaasatus"],
    team.map((r) => r.map((v) => ({ t: v })))),
  source(`Allikas: autorite koostatud. Vastutaja: ${PJ}.`),
  planned(`[Täiendatakse: projekti eesmärk ja tulemid, ulatus ja piirid, projekti organisatsioon ja juhtkomitee. Vastutaja: ${PJ}.]`),
  H1("2. PROJEKTI PLANEERIMINE"),
  H2("2.1. Projekti ulatus ja tööde struktuur"),
  planned("[Eeldatav sisu: tulemipõhine tööde jaotuse struktuur (WBS), vastutusmaatriks (RACI).]"),
  H2("2.2. Ajakava"),
  P("Teostusfaas kestab 3 kuud ja on jaotatud etappideks: **1. kuu** – turu-uuring, brändistrateegia ja kontseptsioon, õiguskontroll ja kaubamärgitaotlused; **2. kuu** – pärandvapi ja brändiraamatu kujundus, kaks testimisvooru fännidega, sponsorite ja litsentsiaatidega kokkulepped; **3. kuu** – merchandise’i prototüübid, lansseerimiskampaania, avalik esitlus ja järelmõõtmise algus."),
  planned("[Täiendatakse: verstapostid, Gantti diagramm, kriitiline tee.]"),
  H2("2.3. Ressursid ja eelarve"),
  P(`Eelarve on üles ehitatud etapiviisiliselt, kontseptsioonist ja õiguskaitsest kuni rakenduse ja avaliku lansseerimiseni. Suurim osa vahenditest on suunatud visuaalse identiteedi loomisele ning brändi turule toomisele (tabel 7).`),
  caption("Tabel", "Projekti esialgne eelarve kulukategooriate kaupa"),
  table([6171, 1500, 1400],
    ["Kulukategooria", "Summa", "Osakaal"],
    [
      ...budget.map(([k, v]) => [{ t: k }, { t: eur(v), align: AlignmentType.RIGHT }, { t: `${Math.round((v / btotal) * 100)}%`, align: AlignmentType.RIGHT }]),
      [{ t: "**Kokku**" }, { t: `**${eur(btotal)}**`, align: AlignmentType.RIGHT }, { t: "**100%**", align: AlignmentType.RIGHT }],
    ]),
  source(`Allikas: autorite koostatud. Kategooriate jaotus on esialgne ja täpsustatakse planeerimisel. Vastutaja: ${OM}.`),
  H2("2.4. Riskide, kvaliteedi ja kommunikatsiooni juhtimine"),
  planned("[Eeldatav sisu: riskijuhtimise plaan, kvaliteedikriteeriumid, kommunikatsiooniplaan huvipoolte analüüsi põhjal.]"),
  H1("3. PROJEKTI ELLUVIIMINE, SEIRE JA LÕPETAMINE"),
  planned("[Eeldatav sisu: seire ja kontroll, muudatuste juhtimine, tulemite üleandmine, tulemuste kasutuskava, õppetundide kogumine.]"),
);

// KOKKUVÕTE
children.push(
  H1("KOKKUVÕTE"),
  P("Rühmatöö alapeatükis 1.2 analüüsiti, miks on Juventus FC-l vaja projekti „Juventus uuesti ja paremini“. Probleemipuu näitas, et 2017. aasta rebrändingu järel killustunud fännibaasi peamised põhjused on sideme katkemine klubi ajaloolise sümboolikaga, fännide kaasamata jätmine ja globaalset ärilist sihtrühma eelistav brändistrateegia. Eesmärgipuu sõnastas soovitud püsiseisundi – ühendatud fännibaasi, mis samastub klubi identiteediga – ning eristas 3-kuulise projekti ulatusse kuuluvad alameesmärgid programmi teiste projektide omadest."),
  P("Huvipoolte analüüs tuvastas 13 huvipoolt, kellest võtmeisikud on omanik, klubi juhatus ja turundusosakond ning sponsorid. Peamine vastuolu – traditsioon versus globaalne bränd – lahendatakse kahetasandilise brändiga ja fännide kaasamisega. Alternatiivide võrdlus näitas, et efektiivseim on pärandvapi taastamine koos fännide kaasamise programmiga (B+D), mis mahub 96 800-eurose eelarve ja 3 kuu sisse ning on kaalude muutmise suhtes stabiilne."),
);

// KASUTATUD ALLIKAD
const refs = [
  "Anthropic. (2026). *Claude* [Suur keelemudel]. https://claude.ai",
  "AXELOS. (2017). *Managing successful projects with PRINCE2* (6th ed.). TSO.",
  "Bryson, J. M. (2004). What to do when stakeholders matter: Stakeholder identification and analysis techniques. *Public Management Review, 6*(1), 21–53. https://doi.org/10.1080/14719030410001675722",
  "Design Week. (2017, jaanuar). *Juventus seeks to go “beyond football” with new brand*. https://www.designweek.co.uk/issues/16-22-january-2017/juventus-seeks/",
  "Dezeen. (2017, 17. jaanuar). *Juventus FC faces fan uprising after launching minimal new logo*. https://www.dezeen.com/2017/01/17/juventus-football-club-faces-fan-uprising-after-minimalist-new-logo-graphics-design/",
  "Eden, C., & Ackermann, F. (1998). *Making strategy: The journey of strategic management*. Sage.",
  "European Commission. (2004). *Aid delivery methods. Volume 1: Project cycle management guidelines*. EuropeAid Cooperation Office.",
  "Football Italia. (2017, 16. jaanuar). *Juventus present new logo*. https://www.football-italia.net/96994/juventus-present-new-logo",
  "IAP2. (2018). *IAP2 spectrum of public participation*. International Association for Public Participation. https://www.iap2.org",
  "Ishikawa, K. (1990). *Introduction to quality control*. 3A Corporation.",
  "ISO. (2012). *ISO 21500:2012 Guidance on project management*. International Organization for Standardization.",
  "It’s Nice That. (2017, 17. jaanuar). *Juventus football club given a new identity by Interbrand*. https://www.itsnicethat.com/news/juventus-football-club-interbrand-rebrand-170117",
  "Keller, K. L. (1993). Conceptualizing, measuring, and managing customer-based brand equity. *Journal of Marketing, 57*(1), 1–22. https://doi.org/10.1177/002224299305700101",
  "Learn Lean Sigma. (s.a.). *Fishbone diagram vs 5 whys analysis*. https://www.learnleansigma.com/root-cause-analysis/fishbone-diagram-vs-5-whys-analysis/",
  "LinkedIn. (s.a.). *How do you compare and contrast problem tree with other problem analysis techniques?* https://www.linkedin.com/advice/0/how-do-you-compare-contrast-problem-tree",
  "Mendelow, A. L. (1981). Environmental scanning: The impact of the stakeholder concept. In *Proceedings of the Second International Conference on Information Systems* (pp. 407–418). ICIS.",
  "Ohno, T. (1988). *Toyota production system: Beyond large-scale production*. Productivity Press.",
  "PMI. (2021). *A guide to the project management body of knowledge (PMBOK guide)* (7th ed.). Project Management Institute.",
  "PMI. (s.a.). *Stakeholder management strategies: Applying risk management*. https://www.pmi.org/learning/library/stakeholder-management-strategies-applying-risk-management-7479",
  "Stakeholdermap.com. (s.a.). *Stakeholder analysis*. https://www.stakeholdermap.com/stakeholder-analysis.html",
  "Tamberg, T. (2022). *Projektijuhtimise meetodid ja tehnikad. 2. Projektide määratlemine* [Loengumaterjalid, slaidid 1–36]. Tartu Ülikooli Pärnu kolledž.",
  "Tartu Ülikooli Pärnu kolledž. (2025). *Üliõpilastööde juhend*. https://parnu.ut.ee/sites/default/files/2025-09/PC_juhend_2025.pdf",
  "The Drum. (2017). *Juventus upset fans after rebranding their famous club crest*. https://www.thedrum.com/news/juventus-upset-fans-after-rebranding-their-famous-club-crest",
  "Wann, D. L., & Branscombe, N. R. (1993). Sports fans: Measuring degree of identification with their team. *International Journal of Sport Psychology, 24*(1), 1–17.",
];
children.push(H1("KASUTATUD ALLIKAD"), ...refs.map((r) => new Paragraph({ style: "Reference", children: runs(r) })));

// LISAD
const tasks = [
  ["1", "Töö koordineerimine, sissejuhatus, 1.1, metoodika 1.2.1 (tabel 1), 1.2.6 (joonis 4), kokkuvõte", `${PJ} (projektijuht)`, AI, "09.10.2026"],
  ["2", "Probleemipuu (joonis 1), probleemi lause, 5 korda miks", `${EX} (ekspert)`, PJ, "09.10.2026"],
  ["3", "Eesmärgipuu ja projekti ulatus (joonis 2), tulemusnäitajad", `${OM} (omanik)`, EX, "11.10.2026"],
  ["4", "Huvipoolte tabel (tabel 2), vastandlikud huvid (tabel 3), maatriks (joonis 3), kaasamise strateegia", PJ, OM, "11.10.2026"],
  ["5", "Lahendusideed, alternatiivide võrdlus (tabelid 4 ja 5), tundlikkusanalüüs", OM, EX, "12.10.2026"],
  ["6", "Meeskonna rollid ja eelarve (tabelid 6 ja 7)", OM, PJ, "12.10.2026"],
  ["7", "Tehisaru päringud, tagasiside ja töö täiendamine (lisa 3)", `${AI} (AI prompter)`, "kõik", "13.10.2026"],
  ["8", "Vormistuse ja kvaliteedi tagamine: blankett, stiilid, pealdised, joonte joondus, viited", `${AI} (kvaliteedikontroll)`, PJ, "14.10.2026"],
  ["9", "Õppetundide analüüs (lisa 2)", `kõik (koond: ${AI})`, "–", "14.10.2026"],
  ["10", "Lõplik ülevaatus ja esitamine", PJ, "kõik", "16.10.2026"],
];
children.push(
  H1("LISAD"),
  H2("Lisa 1. Rühma tööjaotus ja vahetähtajad"),
  P(`Rühma juht (projektijuht) on ${PJ}, omaniku rolli täidab ${OM}, valdkonna eksperdina tegutseb ${EX} ning tehisaru päringute ja kvaliteedikontrolli eest (sh vormistus) vastutab ${AI}. Peamine töökanal on Google Workspace, kus hoitakse töö versioone, tabeleid ja tehisaruga peetud vestlusi; toetav kanal kiireks suhtluseks on Facebook Messengeri grupp.`),
  caption("Tabel", "Tööjaotus, vastutajad ja vahetähtajad"),
  table([500, 3771, 1900, 1300, 1600],
    ["Nr", "Ülesanne või objekt", "Vastutaja", "Kaasteostaja", "Vahetähtaeg"],
    tasks.map((r) => r.map((v, i) => ({ t: v, align: [0, 4].includes(i) ? AlignmentType.CENTER : AlignmentType.LEFT })))),
  source("Allikas: autorite koostatud. Kuupäevad on esialgsed ja täpsustatakse vastavalt õppeaine tähtajale."),

  H2("Lisa 2. Töö õppetunnid"),
  P("**Miks on korrektne vormistus selle ülesande juures oluline?** Rebrändingu projektis on visuaalne ja täpne esitus sisu osa: probleemi- ja eesmärgipuud ning huvipoolte maatriksit kasutatakse omaniku ja huvipooltega kokkulepete tegemisel. Joondatud kastid, ühtne kirjasuurus ja kastide külge kinnitatud ühendajad muudavad põhjus–tagajärg seosed üheselt loetavaks ning võimaldavad joonist kiiresti muuta. Pealdised ja viited tagavad, et tekstis saab objektidele viidata ja lugeja eristab rühma hinnanguid allikatest pärit faktidest."),
  P("**Milliseid oskusi arendasime?** Probleemi eristamist riskist, põhjuslike seoste modelleerimist, eesmärkide sõnastamist püsiseisunditena (mitte „uus logo“), huvipoolte hindamist ja vastandlike huvide ühitamist, kaalutud otsustusanalüüsi koos tundlikkusanalüüsiga, Wordi jooniste ala ja pealdiste kasutamist ning tehisaru väljundi kriitilist hindamist."),
  P("**Mis toetas ja mis takistas õppimist?** Toetasid slaidikogu näited (slaid 17, 29 ja 33), rühmaliikmete rollijaotus (omanik, ekspert, projektijuht, kvaliteedikontroll) ning tehisaru kiire tagasiside. Takistasid see, et fännibaasi killustumise kohta puuduvad avalikud kvantitatiivsed andmed (need tuleb koguda turu-uuringuga), ning see, et huvi- ja mõjuhinnangud on subjektiivsed, kuni neid pole huvipooltega valideeritud."),
  P("**Kas alternatiivsed ideed täidavad kaardistatud eesmärke erinevalt?** Jah. Tabelist 5 nähtub, et vana vapi taastamine (A) ja kahetasandiline bränd (B) katavad sümboolika haru, logo ümberkujundamine koos fännidega (C) ja kommunikatsiooniprogramm (D) kaasamise ja kommunikatsiooni haru. Ükski alternatiiv üksi ei kata kõiki alameesmärke, seetõttu on parim lahendus kombinatsioon B+D."),
  P("**Millist tagasisidet saime tehisarult ja kas täiendasime tööd?** Tehisaru tagasiside ja selle põhjal tehtud muudatused on esitatud lisas 3."),
  P("**Kuidas tagasime rühmas usalduse ja vastutuse?** Usalduse ja vastutuse tagame teadmisega, et meil kõigil on ühine eesmärk – omandada kõrgharidus. Lisaks on igal objektil nimeline vastutaja ja kaasteostaja (lisa 1), kes kontrollib vastutaja tööd enne vahetähtaega. Kõik versioonid ja tehisaruga peetud vestlused on Google Workspace’is kõigile nähtavad ning kiired küsimused lahendatakse Messengeri grupis, mis loob läbipaistvuse ja võimaldab iga liikme panust hinnata."),

  H2("Lisa 3. Vestlus tehisaruga ja selle põhjal tehtud täiendused"),
  P(`Tehisaruna kasutati keelemudelit Claude (Anthropic, 2026); päringud koostas ${AI}. Allpool on vestluse sisu lühendatult; täielik vestlus on rühma Google Workspace’i kaustas kõigile liikmetele nähtav.`),
  P("**Päring 1 (lühendatult):** „Vaadake üle slaidikogu 2 slaidid 1–36 ning muud seonduvad juhendid. Esitage terviklikku peatükkide struktuuri sisaldav DOCX-fail, milles on alapeatükk 1.2: metoodika, probleemide puu, eesmärkide puu, huvigruppide tabel ja kaasamise strateegia, alternatiivide võrdlus; lisad tööjaotuse ja õppetundidega.“"),
  P("**Päring 2 (lühendatult):** „Meie projekt on Juventus FC rebrändimine „Juventus uuesti ja paremini“. 2017. aasta rebränding põhjustas fännibaasi killustumise. Meeskond: Herman Ra Truvek (projektijuht), Robi Mustsaar (AI prompter/kvaliteedikontroll), Hugo-Christopher Saar (ekspert), Ragnar Dietrich (omanik). Huvipooled, kaheksa spetsialistirolli, eelarve ca 96 800 € (7 kulukategooriat, ca 9000 € reserv), teostusfaas 3 kuud.“"),
  P("**Tehisaru tagasiside ja rühma tegevus:**"),
  N("Probleemipuusse ei tohi panna projekti riske (nt „fännid võivad ka uut lahendust kritiseerida“) – need kuuluvad alapeatükki 1.3. *Arvestatud.*", "num2"),
  N("„Uus logo“ on projekti tulem, mitte eesmärk; eesmärk on fännide samastumine klubiga. *Arvestatud: eesmärgipuus on seisundid ja mõõdetavad näitajad.*", "num2"),
  N("96 800 eurot ja 3 kuud on täieliku globaalse rebrändingu jaoks vähe (2017. aasta identiteedi lõi rahvusvaheline agentuur Interbrand). Seetõttu on eelistatud alternatiivid, mis kasutavad olemasolevat „J“-kaubamärki. *Arvestatud: kriteeriumid K3 ja K4.*", "num2"),
  N(`Eelarve jaotus kategooriate kaupa (tabel 7) on tehisaru pakutud näidisjaotus, mis vastab antud kogusummale ja reservile. *Avatud: kinnitavad ${OM} ja ${PJ}.*`, "num2"),
  N(`Rühma viidatud lisamaterjali link 2017. aasta rebrändingu kohta ei jõudnud tehisaruni; kasutati Dezeeni, Football Italia, Design Weeki ja It’s Nice That’i artikleid. *Avatud: ${AI} lisab rühma allika viidetesse.*`, "num2"),
  N(`Torino härg on linna sümbol, mistõttu tuleb selle kasutusõigus pärandvapis kontrollida. *Avatud: ${EX} koos kaubamärgiõiguse spetsialistiga.*`, "num2"),
  N(`Tehisaru ei pääsenud ligi LinkedIni artiklile ega kolledži blanketile. *Avatud: ${PJ} loeb artikli läbi, ${AI} tõstab töö blanketti.*`, "num2"),
  P("**Rühma hinnang tehisaru kasutamisele:** tehisaru kiirendas struktuuri loomist ja aitas leida nõrku kohti, kuid sisulised hinnangud, projekti kontekst ja allikate kontroll jäid rühma vastutusele."),
);

// ---------- dokument ----------
const doc = new Document({
  creator: "Rühm",
  title: "Juventus uuesti ja paremini",
  features: { updateFields: true },
  styles: {
    default: {
      document: { run: { font: FONT, size: 24 }, paragraph: { spacing: { line: 360, after: 120 } } },
    },
    paragraphStyles: [
      { id: "Normal", name: "Normal", run: { font: FONT, size: 24 }, paragraph: { spacing: { line: 360, after: 120 }, alignment: AlignmentType.JUSTIFIED } },
      { id: "Heading1", name: "Heading 1", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: { size: 28, bold: true, font: FONT, allCaps: false }, paragraph: { spacing: { before: 0, after: 360 }, alignment: AlignmentType.LEFT, outlineLevel: 0, keepNext: true } },
      { id: "Heading2", name: "Heading 2", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: { size: 24, bold: true, font: FONT }, paragraph: { spacing: { before: 360, after: 240 }, alignment: AlignmentType.LEFT, outlineLevel: 1, keepNext: true } },
      { id: "Heading3", name: "Heading 3", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: { size: 24, bold: true, font: FONT }, paragraph: { spacing: { before: 240, after: 120 }, alignment: AlignmentType.LEFT, outlineLevel: 2, keepNext: true } },
      { id: "Caption", name: "caption", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: { size: 24, bold: false, italics: false, color: "000000" }, paragraph: { spacing: { before: 120, after: 120, line: 240 }, alignment: AlignmentType.LEFT } },
      { id: "Source", name: "Allikas", basedOn: "Normal", next: "Normal",
        run: { size: 20 }, paragraph: { spacing: { before: 60, after: 240, line: 240 }, alignment: AlignmentType.LEFT } },
      { id: "TableText", name: "Tabeli tekst", basedOn: "Normal",
        run: { size: 20 }, paragraph: { spacing: { before: 0, after: 0, line: 240 }, alignment: AlignmentType.LEFT } },
      { id: "Planned", name: "Eeldatav sisu", basedOn: "Normal",
        run: { italics: true, color: "7F7F7F" } },
      { id: "Reference", name: "Allikaloend", basedOn: "Normal",
        paragraph: { indent: { left: 567, hanging: 567 }, alignment: AlignmentType.LEFT } },
      { id: "TocHeading", name: "TOC Heading", basedOn: "Normal", run: { size: 28, bold: true }, paragraph: { spacing: { after: 360 } } },
    ],
  },
  numbering: {
    config: [
      { reference: "bullets", levels: [{ level: 0, format: LevelFormat.BULLET, text: "•", alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 567, hanging: 283 } } } }] },
      { reference: "num1", levels: [{ level: 0, format: LevelFormat.DECIMAL, text: "%1.", alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 567, hanging: 283 } } } }] },
      { reference: "num2", levels: [{ level: 0, format: LevelFormat.DECIMAL, text: "%1.", alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 567, hanging: 283 } } } }] },
    ],
  },
  sections: [{
    properties: {
      titlePage: true,
      page: { size: { width: 11906, height: 16838 }, margin: { top: 1418, bottom: 1418, left: 1701, right: 1134, footer: 709 } },
    },
    footers: {
      default: new Footer({ children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ children: [PageNumber.CURRENT] })] })] }),
      first: new Footer({ children: [new Paragraph("")] }),
    },
    children,
  }],
});

Packer.toBuffer(doc).then((buf) => {
  fs.writeFileSync(process.argv[2] || "stage1.docx", buf);
  console.log("ok");
});
