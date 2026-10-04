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
    if (t.startsWith("**")) out.push(new TextRun({ text: t.slice(2, -2), bold: true, ...opts }));
    else out.push(new TextRun({ text: t.slice(1, -1), italics: true, ...opts }));
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

// TIITELLEHT
const tc = (t, o = {}) => new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 0 }, children: [new TextRun({ text: t, ...o })] });
children.push(
  tc("TARTU ÜLIKOOL"), tc("Pärnu kolledž"), tc("[Õppekava nimi]"),
  new Paragraph({ spacing: { before: 2400 }, alignment: AlignmentType.CENTER, children: [new TextRun(`${M.A}, ${M.B}, ${M.C}, ${M.D}`)] }),
  new Paragraph({ spacing: { before: 1800, after: 240 }, alignment: AlignmentType.CENTER, children: [new TextRun({ text: "PÄRNU AASTARINGSE KUURORTLINNA ARENDAMINE: MADALHOOAJA KÜLASTATAVUSE SUURENDAMISE PROJEKT", bold: true, size: 28 })] }),
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
  P("Pärnu on tuntud kui Eesti suvepealinn, kuid linna turismimajanduse suurim nõrkus on tugev hooajalisus: suvel on majutusettevõtete tubade täituvus keskmiselt 60–80%, novembrist märtsini aga vaid 39–50% (Pärnu Linnavalitsus, 2018). Pärnu turismiettevõtjad on hooajalisuse vähendamiseks pakkunud välja mitmeid ideid, sh ilmast sõltumatu siseranna rajamise (ERR, 2023). Käesoleva rühmatöö eesmärk on projektijuhi vaatest välja selgitada, **miks** projekti vaja on, millist püsivat seisundit see peaks looma ning kas projekti omaniku esialgne idee on eesmärkide saavutamiseks mõistlik ja efektiivne."),
  P("Töö on üles ehitatud projekti määratluse (*project brief*) loogikas: esimene peatükk määratleb projekti, teine kirjeldab planeerimist ning kolmas elluviimist ja lõpetamist (Tamberg, 2022; AXELOS, 2017). Käesolevas ülesandes on terviklikult koostatud alapeatükk 1.2, mis sisaldab kasutatud metoodika kirjeldust, probleemipuud, eesmärgipuud, huvipoolte analüüsi ja kaasamise strateegiat ning lahendusalternatiivide võrdlust. Teiste peatükkide eeldatav sisu on esitatud kursiivis märkustena ja täidetakse järgmistes ülesannetes."),
  P(`Iga joonise ja tabeli juures on märgitud vastutaja; rühma tööjaotus ja vahetähtajad on esitatud lisas 1, töö õppetunnid lisas 2 ja vestlus tehisaruga lisas 3. Projekt ja selle omanik on õppeülesande raames eeldatud ning ei esinda Pärnu Linnavalitsuse ametlikku seisukohta.`),
);

// 1. PROJEKTI MÄÄRATLEMINE
children.push(
  H1("1. PROJEKTI MÄÄRATLEMINE"),
  H2("1.1. Projekti idee ja taust"),
  planned("[Eeldatav sisu – koostatud ülesandes 1 / täiendatakse: projekti idee, algataja ja omanik, strateegiline sobivus Pärnu linna arengukavaga, ärijuhtumi esialgne kirjeldus (Tamberg, 2022, slaid 8–9). Vastutaja: " + M.A + ".]"),
  P("Projekti idee lähtekoht on Pärnu linna arengukava 2018–2035 eesmärk vähendada turismi hooajalisust (Pärnu Linnavalitsus, 2018) ning turismiettevõtjate ettepanek rajada linna aastaringne siserand (ERR, 2023). Õppeülesandes eeldame, et projekti omanik on Pärnu Linnavalitsus ja projektijuht on " + M.A + "."),
  H2("1.2. Probleemide, eesmärkide ja huvipoolte analüüs"),
  P("Alapeatüki eesmärk on enne planeerimist veenduda, et valitakse õige projekt. Probleemiks loetakse lahknevust soovitava ja olemasoleva seisundi vahel **enne projekti**, mitte projekti käigus tekkida võivaid riske (Tamberg, 2022, slaid 11–12). Analüüs liigub järjekorras probleem → põhjuste analüüs → eesmärgid → huvipooled → lahendusideed → lahenduse valik (Tamberg, 2022, slaid 10)."),
);

// 1.2.1 METOODIKA
children.push(
  H3("1.2.1. Metoodika ja kasutatud tehnikad"),
  P("Analüüsi aluseks on loogilise raamistiku lähenemine (*Logical Framework Approach*, LFA), mille järgi koostatakse esmalt probleemipuu, see teisendatakse eesmärgipuuks ja valitakse eesmärgipuust projekti ulatusse kuuluvad harud (European Commission, 2004). Rühm kasutas järgmisi tehnikaid:"),
  N("**Dokumendianalüüs ja näitarvude tehnika** – Pärnu linna ja maakonna arengudokumentide ning majutusstatistika võrdlemine hooaegade lõikes (Tamberg, 2022, slaid 13; Pärnumaa Omavalitsuste Liit, 2019)."),
  N("**Ajurünnak ja 635-meetod** probleemide, põhjuste ja lahendusideede kogumiseks; kalasaba-diagrammi kategooriaid (taristu, sündmused, turundus, ligipääs, juhtimine) kasutati kontrollnimekirjana, et ükski põhjuste rühm ei jääks märkamata (Ishikawa, 1990; Tamberg, 2022, slaid 18 ja 35)."),
  N("**Probleemipuu** põhjus–tagajärg seoste kaardistamiseks ja **probleemi lause** sõnastamiseks (Tamberg, 2022, slaid 16–20)."),
  N("**5 korda miks** üksikute algpõhjuste sügavuse kontrolliks (Ohno, 1988)."),
  N("**Eesmärgipuu**, milles negatiivsed seisundid sõnastati soovitud püsiseisunditeks ning märgiti projekti ulatus (Tamberg, 2022, slaid 31–34; European Commission, 2004)."),
  N("**Huvipoolte analüüs**: kontrollküsimused huvipoolte leidmiseks, huvi- ja mõjuhinnangud skaalal 1–5, mõju-huvi maatriks ning kaasamise tasemed (Mendelow, 1981; Bryson, 2004; IAP2, 2018; Tamberg, 2022, slaid 22–30)."),
  N("**Benchmarking ja kaalutud mitmekriteeriumiline võrdlus** lahendusalternatiivide hindamiseks koos tundlikkusanalüüsiga (Tamberg, 2022, slaid 35–36)."),
  P("**Probleemianalüüsi tehnika valik.** Rühm võrdles probleemi analüüsiks sobivaid tehnikaid (tabel 1) ja arutles, milline neist aitab kõige paremini vastata küsimusele, miks projekti vaja on. Kalasaba-diagramm ja „5 korda miks“ sobivad hästi protsessi- või kvaliteediprobleemi põhjuste leidmiseks, kuid ei näita probleemi tagajärgi ega teisene otse eesmärkideks (Learn Lean Sigma, s.a.; LinkedIn, s.a.). SWOT kirjeldab organisatsiooni sise- ja väliskeskkonda, mitte põhjuslikke seoseid, mistõttu kasutatakse seda alles alapeatükis 1.3. Pärnu hooajalisus on mitme põhjusega süsteemne probleem, mis puudutab paljusid huvipooli, seega valis rühm **põhitehnikaks probleemipuu**: see näitab nii põhjuseid kui tagajärgi, seob eri huvipoolte probleemid ühiste algpõhjuste kaudu ja teiseneb otse eesmärgipuuks (European Commission, 2004; Tamberg, 2022, slaid 16–17)."),
  caption("Tabel", "Probleemianalüüsi tehnikate võrdlus"),
  table([1700, 2550, 2550, 2271],
    ["Tehnika", "Tugevused", "Nõrkused", "Kasutus käesolevas töös"],
    [
      ["**Probleemipuu** (LFA)", "Näitab põhjusi ja tagajärgi; teiseneb eesmärgipuuks; sobib mitme huvipoolega süsteemsele probleemile", "Võib lihtsustada vastastikmõjusid; sõltub osalejate teadmistest", "Põhitehnika (joonis 1)"],
      ["Kalasaba-diagramm (Ishikawa)", "Süstemaatilised põhjuste kategooriad; hea ajurünnaku struktuur", "Ainult põhjused, tagajärjed puuduvad; ei teisene eesmärkideks", "Kategooriad kontrollnimekirjana ajurünnakul"],
      ["5 korda miks", "Lihtne ja kiire; viib sümptomist algpõhjuseni", "Lineaarne, üks põhjusahel; oht peatuda liiga vara", "Algpõhjuste sügavuse kontroll (nt P3b)"],
      ["SWOT", "Ülevaade sise- ja väliskeskkonnast", "Ei näita põhjuslikke seoseid", "Keskkonna analüüs alapeatükis 1.3"],
    ]),
  source(`Allikas: autorite koostatud (European Commission, 2004; Ishikawa, 1990; Ohno, 1988; Learn Lean Sigma, s.a.; LinkedIn, s.a.). Vastutaja: ${M.A}.`),
  P("**Tehisaru ja teaduslike lisaallikate kasutamine.** Rühm kasutas suurt keelemudelit Claude (Anthropic, 2026) töö struktuuri esialgse kavandi koostamiseks, tehnikate võrdluse lähteallikate leidmiseks ning valmis versioonile kriitilise tagasiside saamiseks. Tehisaru pakutud väited ja viited kontrolliti algallikatest; arvulised hinnangud (huvi, mõju, alternatiivide punktid) on rühma ekspertarvamus, mitte tehisaru väljund. Vestlus ja selle põhjal tehtud täiendused on esitatud lisas 3. Hooajalisuse kui nähtuse mõistmiseks kasutati teaduskirjandust turismi hooajalisuse (Butler, 2001) ja sündmusturismi (Getz & Page, 2016) kohta."),
);

// 1.2.2 PROBLEEMIPUU
children.push(
  H3("1.2.2. Probleemipuu ja probleemi lause"),
  P("Põhiprobleem määratleti esmalt erinevate huvipoolte vaatest (Tamberg, 2022, slaid 17): ettevõtjate jaoks on probleemiks madal talvine käive, linna jaoks ebaühtlane linnaruumi kasutus ja maksutulu, elanike jaoks teenuste vähesus talvel. Ühine nimetaja on see, et **Pärnu külastatavus ja turismitulu langevad madalhooajal järsult**. Probleemipuus (joonis 1) on põhiprobleemi all kolm otsest põhjust ja kuus algpõhjust ning selle kohal tagajärjed."),
  FIG(1),
  caption("Joonis", `Pärnu madalhooaja probleemipuu (autorite koostatud; andmed: Pärnu Linnavalitsus, 2018; ERR, 2023). Vastutaja: ${M.B}`),
  P("Algpõhjuste sügavust kontrolliti tehnikaga „5 korda miks“. Näiteks: *Miks on talvel vähe teenuseid?* – Ettevõtted lühendavad lahtiolekuaegu. *Miks?* – Külastajaid on vähe ja lahtihoidmine ei tasu ära. *Miks on külastajaid vähe?* – Puuduvad põhjused tulla (atraktsioonid, sündmused). Ahel näitas, et P3b on **nõiaring**, mida ei saa lahendada ainult ettevõtjaid veenes – vaja on samaaegselt luua nõudlust (P1, P2)."),
  P("**Probleemi lause.** Ilma selle projektita ei saa Pärnu vähendada turismi hooajalisust, sest oktoobrist aprillini puuduvad ilmast sõltumatud põhjused linna külastamiseks ja sihtkohta turundatakse peamiselt suvise rannapuhkusena, mistõttu langeb majutuse täituvus igal talvel ligikaudu 40–50%-ni, paljud ettevõtted sulgevad ja linna konkurentsivõime aastaringse elukeskkonnana nõrgeneb."),
);

// 1.2.3 EESMÄRGIPUU
children.push(
  H3("1.2.3. Eesmärgipuu ja projekti ulatus"),
  P("Eesmärgipuu (joonis 2) koostati probleemipuu negatiivsete seisundite ümbersõnastamisel soovitud, tulevikus püsivateks seisunditeks – mitte tegevusteks (Tamberg, 2022, slaid 31–34). Seejärel otsustati, milliste alameesmärkide saavutamine kuulub käesoleva projekti ulatusse ja millised jäetakse programmi teistele projektidele (vt alapeatükk 1.2.6)."),
  FIG(2),
  caption("Joonis", `Pärnu aastaringse külastatavuse eesmärgipuu (autorite koostatud). Vastutaja: ${M.B}`),
  P("Peaeesmärgi saavutamist mõõdetakse järgmiste näitajatega (lähtetase 2025. aasta madalhooaja andmed):"),
  B("majutuse täituvus novembrist märtsini vähemalt 55% aastaks 2030 (lähtetase ca 39–50%);"),
  B("madalhooaja majutatute arv kasvab 15% kolme aasta jooksul;"),
  B("aastaringselt avatud toitlustus- ja vaba aja ettevõtete osakaal kasvab;"),
  B("madalhooajal toimuvate rahvusvahelise kõlapinnaga sündmuste arv kasvab vähemalt neljani aastas."),
  P("Projekti ulatusse kuuluvad alameesmärgid O1b, O2a, O2b ja O3b, mis on saavutatavad koostöö, turunduse ja sündmuste korraldamise kaudu. Aastaringse taristu investeeringud (O1a) ja transpordiühendused (O3a) jäetakse programmi eraldi projektidele, sest need vajavad pikemat ettevalmistust, teiste otsustajate (volikogu, riik) heakskiitu ja keskkonnaalaseid kooskõlastusi."),
);

// 1.2.4 HUVIPOOLED
const shRows = stakeholders.map((s) => {
  const quadrant = s.influence >= 4 ? (s.interest >= 4 ? "Võtmeisik" : "Hoia rahul") : (s.interest >= 4 ? "Hoia informeerituna" : "Jälgi");
  const q = { "Võtmeisik": "F8CBAD", "Hoia rahul": "FFF2CC", "Hoia informeerituna": "DEEAF6", "Jälgi": "E2EFDA" }[quadrant];
  return [
    { t: `${s.nr}. ${s.name}` },
    { t: String(s.interest), align: AlignmentType.CENTER },
    { t: s.wish },
    { t: String(s.influence), align: AlignmentType.CENTER },
    { t: s.mode },
    { t: `**${quadrant}:** ${s.strategy.split(": ").slice(1).join(": ")}`, fill: q },
  ];
});
children.push(
  H3("1.2.4. Huvipoolte analüüs ja kaasamise strateegia"),
  P("Huvipooled leiti kontrollküsimuste abil: kelle vaated ja kogemused on asjakohased, kes on otsustajad, kes hakkavad otsuste järgi tegutsema, kelle toetus on edu jaoks oluline, kellel on õigus tulemustest kasu saada ja kes võib tunda end ohustatuna (Tamberg, 2022, slaid 22). Huvi ja mõju hinnati skaalal 1–5 rühma konsensuse alusel; hinnangud valideeritakse projekti algatamisel intervjuudega. Tabelis 2 on esitatud huvipoolte huvi tingimused ja kaasamise strateegia, joonisel 3 nende paiknemine mõju-huvi maatriksis (Mendelow, 1981; Eden & Ackermann, 1998)."),
  caption("Tabel", "Huvipoolte analüüs ja kaasamise strateegia"),
  table([1900, 680, 1850, 680, 1650, 2311],
    ["Huvipool", "Huvi (1–5)", "Huvi (osalemise, toetamise) tingimus", "Mõju (1–5)", "Osalemise või mõju viis", "Kaasamise strateegia"],
    shRows),
  source(`Allikas: autorite koostatud (Tamberg, 2022, slaid 28; Mendelow, 1981). Värv vastab maatriksi ruudule joonisel 3. Vastutaja: ${M.C}.`),
  FIG(3),
  caption("Joonis", `Huvipoolte mõju-huvi maatriks (autorite koostatud Mendelow, 1981 ja Tamberg, 2022, slaid 29 põhjal). Vastutaja: ${M.C}`),
  P("**Erinevasuunalised huvid.** Huvipoolte eesmärgid on ühitatavad (Tamberg, 2022, slaid 21), kuid mitmes kohas vastanduvad (tabel 3). Projektijuhi ülesanne on vastuolud varakult nähtavaks teha ja leida kokkulepped, mille korral ühine huvi (madalhooaja külastatavuse kasv) kaalub üles erihuvid (Bryson, 2004)."),
  caption("Tabel", "Huvipoolte vastandlikud huvid ja kavandatud kokkulepped"),
  table([2300, 3100, 3671],
    ["Vastuolu", "Huvipoolte soovid", "Huvide ühitamine ja kokkulepe"],
    [
      ["Suurinvesteering vs eelarvedistsipliin", "Majutusettevõtted (3) soovivad siseranda ja avalikku kaasrahastust; volikogu (2) ja linnavalitsus (1) tõendatud tasuvust", "Siserand viiakse eraldi teostatavusuuringu projekti (joonis 4); otsus tehakse uuringu ja käesoleva projekti tulemusnäitajate põhjal"],
      ["Sündmuste maht vs elukvaliteet", "Korraldajad (5) ja ettevõtted (4) soovivad rohkem ja hilisemaid sündmusi; elanikud (6) vaikust ja parkimist", "„Hea naabruse“ kokkulepe: müraajad, parkimise ja ühistranspordi korraldus, elanike soodustus sündmustele, avalik arutelu enne kalendri kinnitamist"],
      ["Suured vs väikesed ettevõtted", "Suured spaad soovivad ühisturunduses suuremat nähtavust, väikeettevõtted võrdset kohtlemist", "Koostöömemorandum: ühisturunduse panus ja nähtavus proportsionaalselt voodikohtade arvuga, väikeettevõtetele ühispaketid"],
      ["Arendus vs looduskaitse", "Ettevõtjad soovivad rannaalal taristut; keskkonnaühendused (10) kaitset", "Käesolev projekt ei ehita rannaalale; jätkuprojektis keskkonnamõju eelhinnang enne investeerimisotsust"],
      ["Hind vs tulu", "Külastajad (11) soovivad soodsat hinda; ettevõtjad kõrgemat keskmist arvet", "Paketid lisaväärtusega (sündmus + majutus + toit), soodustused madalaima nõudlusega nädalatel"],
    ]),
  source(`Allikas: autorite koostatud. Sulgudes on huvipoole number tabelist 2. Vastutaja: ${M.C}.`),
  P("**Kaasamise strateegia.** Kaasamise intensiivsus valiti vastavalt mõjukuse ja huvi määrale (Mendelow, 1981) ning IAP2 kaasamise spektri tasemetele (informeerimine, konsulteerimine, kaasamine, koostöö, otsustusõiguse andmine) (IAP2, 2018):"),
  B("**Võtmeisikud** (1, 3, 7) – koostöö tasand: liikmed projekti juhtkomitees, kes kinnitab iga etapi lõpus ärijuhtumi aktuaalsuse; sõlmitakse koostöömemorandum ja ühisturunduse rahastusvalem."),
  B("**Hoia rahul** (2, 8, 10) – konsulteerimine: kvartaalsed lühiülevaated mõõdetavate näitajatega, varajane konsultatsioon rahastajate ja kooskõlastajatega, et vältida „väravavahtide“ hilist vastuseisu (Tamberg, 2022, slaid 24)."),
  B("**Hoia informeerituna** (4, 5, 12, 13) – kaasamine töörühmadesse (sündmuste kalender, lahtiolekuajad); eesmärk on huvi kasvatades viia ettevõtjad ja korraldajad võtmeisikute rühma (Tamberg, 2022, slaid 29)."),
  B("**Jälgi** (6, 9, 11) – informeerimine ja konsulteerimine: uudiskiri, linna veeb, avalik arutelu ja külastajaküsitlused; elanike puhul on mõju madal, kuid nende rahulolematus võib kanduda volikogu kaudu, seetõttu kaasatakse nad enne sündmuste kalendri kinnitamist."),
  P("Kaasamise eest vastutab projektijuht, iga huvipoolte rühma jaoks määratakse rühma liikmest kontaktisik (lisa 1). Huvipoolte register vaadatakse üle iga etapi lõpus, sest huvipoolte huvi ja mõju võivad projekti käigus muutuda."),
);

// 1.2.5 ALTERNATIIVID
const altRows = alternatives.map((a) => [
  { t: `**${a.id}** – ${a.name}`, fill: a.id === "A+C" ? "E2EFDA" : undefined },
  ...a.scores.map((s) => ({ t: String(s), align: AlignmentType.CENTER, fill: a.id === "A+C" ? "E2EFDA" : undefined })),
  { t: `**${fmt(a.total)}**`, align: AlignmentType.CENTER, fill: a.id === "A+C" ? "E2EFDA" : undefined },
]);
const A_ = Object.fromEntries(alternatives.map((a) => [a.id, a]));
children.push(
  H3("1.2.5. Lahendusideed ja alternatiivide võrdlus"),
  P("Lahendusideed koguti ajurünnaku ja 635-meetodi abil ning benchmarking’u teel (Tamberg, 2022, slaid 35). Võrdlusalustena vaadeldi talviste sündmuste programme teistes Põhjamaade kuurortides ja ERR-i (2023) artiklis eeskujuna nimetatud Saksamaa Brandenburgi liidumaa siseveekeskust. Ideed rühmitati kolmeks põhimõtteliselt erinevaks alternatiiviks, millele lisati nende kombinatsioon ja nullalternatiiv:"),
  B("**A – „Talvine Pärnu“ sündmuste programm ja ühisturundus**: ühine madalhooaja sündmuste kalender, 3–4 suursündmust, majutuse-sündmuse paketid ja koondatud turunduseelarve."),
  B("**B – Siserand**: aastaringne vee- ja vabaajakeskus konverentsi- ja kontserdivõimekusega (ettevõtjate esialgne idee; ERR, 2023)."),
  B("**C – „Pärnu Pass“**: digitaalne ühispileti ja broneerimise platvorm, mis koondab madalhooaja teenused ning kogub nõudluse andmeid."),
  B("**A+C** – etapiviisiline kombinatsioon: esmalt sündmuste programm ja ühisturundus, seejärel platvorm."),
  B("**0 – nullalternatiiv**: iga osapool jätkab senist tegevust."),
  P("Alternatiive võrreldi huvipoolte eesmärkide saavutamise määra, teostatavuse, kulu ja riskantsuse lõikes (Tamberg, 2022, slaid 36). Kriteeriumide kaalud lepiti kokku enne hindamist, et vältida kaalude sobitamist eelistatud lahendusele (tabel 4)."),
  caption("Tabel", "Lahendusalternatiivide kaalutud võrdlus (hinded 1–5, 5 = parim)"),
  table([2411, 860, 860, 860, 860, 860, 860, 1500],
    ["Alternatiiv", ...criteria.map((c) => `${c.id}\n${Math.round(c.w * 100)}%`), "Kaalutud summa"],
    altRows),
  source(`Kriteeriumid: ${criteria.map((c) => `${c.id} – ${c.name}`).join("; ")}. Allikas: autorite koostatud. Vastutaja: ${M.D}.`),
  P(`Kaalutud summa järgi on parim kombinatsioon **A+C** (${fmt(A_["A+C"].total)}), sellele järgneb alternatiiv A (${fmt(A_.A.total)}). Ettevõtjate algne idee B (${fmt(A_.B.total)}) annaks küll suurima mõju põhieesmärgile, kuid on kõige kallim, aeganõudvam ja riskantsem ning vajab keskkonnakooskõlastusi. Tundlikkusanalüüs näitas, et kui kulu kaal tõsta 30%-ni, on A (${fmt(A_.A.sensCost)}) ja A+C (${fmt(A_["A+C"].sensCost)}) praktiliselt võrdsed ning B langeb veelgi (${fmt(A_.B.sensCost)}); riski kaalu tõstmine 25%-ni järjestust ei muuda. Seega on otsus kaalude suhtes stabiilne: alustada tuleb alternatiivist A ja lisada C teise etapina siis, kui esimese etapi tulemused õigustavad platvormi kulu.`),
  P("Alternatiivid täidavad eesmärgipuu alameesmärke erinevalt (tabel 5): A katab peamiselt sündmuste ja turunduse haru, C ligipääsu ja teenuste kättesaadavuse haru ning B aastaringse taristu haru."),
  caption("Tabel", "Alternatiivide panus eesmärgipuu alameesmärkidesse"),
  table([2771, 900, 900, 900, 900, 900, 900, 900],
    ["Alternatiiv", "O1a", "O1b", "O2a", "O2b", "O3a", "O3b", "Kokku"],
    [
      ["A – sündmused ja ühisturundus", "0", "1", "1", "1", "0,5", "0,5", "4,0"],
      ["B – siserand", "1", "0,5", "0,5", "0", "0", "0,5", "2,5"],
      ["C – „Pärnu Pass“", "0", "0,5", "0,5", "0,5", "0,5", "1", "3,0"],
      ["A+C", "0", "1", "1", "1", "0,5", "1", "4,5"],
      ["0 – nullalternatiiv", "0", "0", "0", "0", "0", "0", "0,0"],
    ].map((r) => r.map((v, i) => ({ t: v, align: i ? AlignmentType.CENTER : AlignmentType.LEFT }))),
  ),
  source(`Märkus: 1 – täielik panus, 0,5 – osaline panus, 0 – panus puudub; tähised vastavad joonisele 2. Allikas: autorite koostatud. Vastutaja: ${M.D}.`),
);

// 1.2.6 ÄRIJUHTUM, PROGRAMM, PORTFELL
children.push(
  H3("1.2.6. Ärijuhtumi seire, programm ja portfell"),
  P("Projekti omanikul võib olla idee, kuid projektijuhi ülesanne on tagada, et kavandatav töö viib omaniku ja huvipoolte eesmärgile võimalikult efektiivselt lähemale (Tamberg, 2022, slaid 8; AXELOS, 2017). Analüüsi tulemusel tehakse omanikule ettepanek käsitleda siseranna ideed mitte kohe investeerimisprojektina, vaid jätkuprojektina, mille aluseks on teostatavus- ja keskkonnamõju eeluuring ning käesoleva projekti tegelikud nõudlusandmed."),
  P("Projekt kuulub programmi „Aastaringne Pärnu“, mis teenib Pärnu linna arengukava strateegilist eesmärki (joonis 4). Programm võimaldab ühitada mitut omavahel seotud projekti ja jagada ressursse; portfelli tasandil otsustatakse, milliseid linna ja partnerite ressursse ning teenuseid (eelarve, toetusmeetmed, sihtkoha turundus, kultuuriasutuste sündmused) projektidele eraldatakse (PMI, 2021; ISO, 2012)."),
  FIG(4),
  caption("Joonis", `Projekti seos strateegia, programmi ja portfelliga (autorite koostatud). Vastutaja: ${M.A}`),
  P("Ärijuhtumi aktuaalsust ja projekti sisu adekvaatsust jälgitakse kogu projekti vältel paindlikult:"),
  B("iga etapi lõpus (PRINCE2 etapipiir) vaatab juhtkomitee üle probleemi- ja eesmärgipuu, huvipoolte registri ja tulemusnäitajad ning otsustab jätkamise, muutmise või lõpetamise (AXELOS, 2017);"),
  B("pärast iga madalhooaja sündmust analüüsitakse külastatavuse ja täituvuse andmeid; kui näitajad ei parane, korrigeeritakse sündmuste ja paketti sisu järgmiseks iteratsiooniks;"),
  B("kui väliskeskkond muutub (nt Rail Balticu ühenduse avanemine, uus toetusmeede või eraarendaja investeering), hinnatakse alternatiivid uuesti ning vajadusel muudetakse projekti ulatust või programmi koosseisu."),
);

// 1.3–3 eeldatav
children.push(
  H2("1.3. Keskkonna ja esmaste riskide analüüs"),
  planned(`[Eeldatav sisu: PESTLE ja SWOT; esmaste riskide register (tõenäosus × mõju), riskide maandamise ja väljumisstrateegia (Tamberg, 2022, slaid 11). Vastutaja: ${M.D}.]`),
  H2("1.4. Projekti määratlus ja projektiettepanek"),
  planned(`[Eeldatav sisu: projekti eesmärk ja tulemid, ulatus ja piirid, ärijuhtum (strateegiline sobivus, valikute hindamine, ärisuhted, tasuvus ja rahastamine, teostatavus), projekti organisatsioon (Tamberg, 2022, slaid 8–9). Vastutaja: ${M.A}.]`),
  H1("2. PROJEKTI PLANEERIMINE"),
  H2("2.1. Projekti ulatus ja tööde struktuur"),
  planned("[Eeldatav sisu: tulemipõhine tööde jaotuse struktuur (WBS), vastutusmaatriks (RACI).]"),
  H2("2.2. Ajakava"),
  planned("[Eeldatav sisu: etapid ja verstapostid, Gantti diagramm, kriitiline tee.]"),
  H2("2.3. Ressursid ja eelarve"),
  planned("[Eeldatav sisu: ressursiplaan, eelarve ja rahastamisallikad.]"),
  H2("2.4. Riskide, kvaliteedi ja kommunikatsiooni juhtimine"),
  planned("[Eeldatav sisu: riskijuhtimise plaan, kvaliteedikriteeriumid, kommunikatsiooniplaan huvipoolte analüüsi põhjal.]"),
  H1("3. PROJEKTI ELLUVIIMINE, SEIRE JA LÕPETAMINE"),
  planned("[Eeldatav sisu: seire ja kontroll, muudatuste juhtimine, tulemite üleandmine, tulemuste kasutuskava, õppetundide kogumine.]"),
);

// KOKKUVÕTE
children.push(
  H1("KOKKUVÕTE"),
  P("Rühmatöö alapeatükis 1.2 analüüsiti, miks on Pärnul vaja madalhooaja külastatavuse suurendamise projekti. Probleemipuu näitas, et põhiprobleemi – külastatavuse ja turismitulu järsu languse madalhooajal – peamised põhjused on ilmast sõltumatute atraktsioonide ja sündmuste vähesus, suvele keskenduv turundus ning piiratud ligipääs ja teenuste kättesaadavus talvel. Eesmärgipuu sõnastas soovitud püsiseisundi ja eristas projekti ulatusse kuuluvad alameesmärgid programmi teiste projektide omadest."),
  P("Huvipoolte analüüs tuvastas 13 huvipoolt, kellest võtmeisikud on linnavalitsus, majutus- ja spaaettevõtted ning sihtkoha turundusorganisatsioon. Peamised vastuolud (suurinvesteering vs eelarvedistsipliin, sündmuste maht vs elukvaliteet, arendus vs looduskaitse) lahendatakse kokkulepete ja mõjukusele vastava kaasamisega. Alternatiivide võrdlus näitas, et efektiivseim on sündmuste programmi ja ühisturunduse ning digitaalse platvormi etapiviisiline kombinatsioon; omaniku esialgne siseranna idee jätkab programmis teostatavusuuringuna."),
);

// KASUTATUD ALLIKAD
const refs = [
  "Anthropic. (2026). *Claude* [Suur keelemudel]. https://claude.ai",
  "AXELOS. (2017). *Managing successful projects with PRINCE2* (6th ed.). TSO.",
  "Bryson, J. M. (2004). What to do when stakeholders matter: Stakeholder identification and analysis techniques. *Public Management Review, 6*(1), 21–53. https://doi.org/10.1080/14719030410001675722",
  "Butler, R. W. (2001). Seasonality in tourism: Issues and implications. In T. Baum & S. Lundtorp (Eds.), *Seasonality in tourism* (pp. 5–21). Pergamon.",
  "Eden, C., & Ackermann, F. (1998). *Making strategy: The journey of strategic management*. Sage.",
  "ERR. (2023). *Pärnu ettevõtjad tahavad rajada linna siseranna*. https://www.err.ee/1609560280/parnu-ettevotjad-tahavad-rajada-linna-siseranna",
  "European Commission. (2004). *Aid delivery methods. Volume 1: Project cycle management guidelines*. EuropeAid Cooperation Office.",
  "Getz, D., & Page, S. J. (2016). Progress and prospects for event tourism research. *Tourism Management, 52*, 593–631. https://doi.org/10.1016/j.tourman.2015.03.007",
  "Hohmann, C. (2018, 13. september). *Goal tree, not for strategy only*. https://hohmannchris.wordpress.com/2018/09/13/goal-tree-not-for-strategy-only/",
  "IAP2. (2018). *IAP2 spectrum of public participation*. International Association for Public Participation. https://www.iap2.org",
  "Ishikawa, K. (1990). *Introduction to quality control*. 3A Corporation.",
  "ISO. (2012). *ISO 21500:2012 Guidance on project management*. International Organization for Standardization.",
  "Learn Lean Sigma. (s.a.). *Fishbone diagram vs 5 whys analysis*. https://www.learnleansigma.com/root-cause-analysis/fishbone-diagram-vs-5-whys-analysis/",
  "LinkedIn. (s.a.). *How do you compare and contrast problem tree with other problem analysis techniques?* https://www.linkedin.com/advice/0/how-do-you-compare-contrast-problem-tree",
  "Mendelow, A. L. (1981). Environmental scanning: The impact of the stakeholder concept. In *Proceedings of the Second International Conference on Information Systems* (pp. 407–418). ICIS.",
  "Ohno, T. (1988). *Toyota production system: Beyond large-scale production*. Productivity Press.",
  "Pärnu Linnavalitsus. (2018). *Pärnu linna arengukava 2018–2035*. Pärnu Linnavalitsus.",
  "Pärnumaa Omavalitsuste Liit. (2019). *Pärnumaa arengustrateegia 2035*. https://parnumaa.ee/wp-content/uploads/2020/04/Arengustrateegia-2035.pdf",
  "PMI. (2021). *A guide to the project management body of knowledge (PMBOK guide)* (7th ed.). Project Management Institute.",
  "PMI. (s.a.). *Stakeholder management strategies: Applying risk management*. https://www.pmi.org/learning/library/stakeholder-management-strategies-applying-risk-management-7479",
  "SSWM. (s.a.). *Problem tree analysis*. Sustainable Sanitation and Water Management Toolbox. http://www.sswm.info/content/problem-tree-analysis",
  "Stakeholdermap.com. (s.a.). *Stakeholder analysis*. https://www.stakeholdermap.com/stakeholder-analysis.html",
  "Tamberg, T. (2022). *Projektijuhtimise meetodid ja tehnikad. 2. Projektide määratlemine* [Loengumaterjalid, slaidid 1–36]. Tartu Ülikooli Pärnu kolledž.",
  "Tartu Ülikooli Pärnu kolledž. (2025). *Üliõpilastööde juhend*. https://parnu.ut.ee/sites/default/files/2025-09/PC_juhend_2025.pdf",
];
children.push(H1("KASUTATUD ALLIKAD"), ...refs.map((r) => new Paragraph({ style: "Reference", children: runs(r) })));

// LISAD
const tasks = [
  ["1", "Töö koordineerimine, sissejuhatus, alapeatükid 1.1 ja 1.2.1 (metoodika, tabel 1), 1.2.6 (joonis 4), kokkuvõte", M.A + " (projektijuht)", M.D, "09.10.2026"],
  ["2", "Probleemipuu (joonis 1), probleemi lause, 5 korda miks", M.B, M.A, "09.10.2026"],
  ["3", "Eesmärgipuu ja projekti ulatus (joonis 2), tulemusnäitajad", M.B, M.D, "11.10.2026"],
  ["4", "Huvipoolte tabel (tabel 2), vastandlike huvide tabel (tabel 3), maatriks (joonis 3), kaasamise strateegia", M.C, M.A, "11.10.2026"],
  ["5", "Lahendusideed, alternatiivide võrdlus (tabelid 4 ja 5), tundlikkusanalüüs", M.D, M.C, "12.10.2026"],
  ["6", "Vormistuse tagamine: blankett, stiilid, pealdised, joonte joondus, viited ja allikaloend", M.D, "kõik", "14.10.2026"],
  ["7", "Tehisaru tagasiside küsimine ja töö täiendamine (lisa 3)", M.A, "kõik", "13.10.2026"],
  ["8", "Õppetundide analüüs (lisa 2)", "kõik (koond: " + M.C + ")", "–", "14.10.2026"],
  ["9", "Lõplik ülevaatus ja esitamine", M.A, "kõik", "16.10.2026"],
];
children.push(
  H1("LISAD"),
  H2("Lisa 1. Rühma tööjaotus ja vahetähtajad"),
  P(`Rühma juht (projektijuht) on ${M.A}, vormistuse eest vastutab ${M.D}. Rühm kohtub vahetähtaegade järel veebis; töö versioonid hoitakse ühises kaustas, kus iga muudatuse tegija on nähtav.`),
  caption("Tabel", "Tööjaotus, vastutajad ja vahetähtajad"),
  table([500, 3971, 1700, 1300, 1600],
    ["Nr", "Ülesanne või objekt", "Vastutaja", "Kaasteostaja", "Vahetähtaeg"],
    tasks.map((r) => r.map((v, i) => ({ t: v, align: [0, 4].includes(i) ? AlignmentType.CENTER : AlignmentType.LEFT })))),
  source("Allikas: autorite koostatud. Kuupäevad on esialgsed ja täpsustatakse vastavalt õppeaine tähtajale."),

  H2("Lisa 2. Töö õppetunnid"),
  P("**Miks on korrektne vormistus selle ülesande juures oluline?** Probleemi- ja eesmärgipuu ning huvipoolte maatriks on suhtlusvahendid: neid kasutatakse omaniku ja huvipooltega kokkulepete tegemisel. Joondatud kastid, ühtne kirjasuurus ja kastide külge kinnitatud ühendajad muudavad põhjus–tagajärg seosed üheselt loetavaks ning võimaldavad joonist projekti käigus kiiresti muuta. Pealdised ja viited tagavad, et tekstist saab objektidele viidata ja et lugeja eristab rühma hinnanguid allikatest pärit faktidest."),
  P("**Milliseid oskusi arendasime?** Probleemi eristamist riskist, põhjuslike seoste modelleerimist, eesmärkide sõnastamist püsiseisunditena, huvipoolte hindamist ja kaasamise kavandamist, kaalutud otsustusanalüüsi koos tundlikkusanalüüsiga, Wordi jooniste ala ja pealdiste kasutamist ning tehisaru väljundi kriitilist hindamist."),
  P("**Mis toetas ja mis takistas õppimist?** Toetasid slaidikogu näited (slaid 17, 29 ja 33), rühmaarutelu ja tehisaru kiire tagasiside. Takistasid avalike andmete killustatus (täituvusnäitajad pärinevad eri aastatest), mõnede veebiallikate piiratud kättesaadavus ning see, et huvi- ja mõjuhinnangud on subjektiivsed, kuni neid pole huvipooltega valideeritud."),
  P("**Kas alternatiivsed ideed täidavad kaardistatud eesmärke erinevalt?** Jah. Tabelist 5 nähtub, et sündmuste programm (A) katab sündmuste ja turunduse haru, siserand (B) aastaringse taristu haru ning platvorm (C) teenuste kättesaadavuse haru. Ükski alternatiiv üksi ei kata kõiki alameesmärke, seetõttu on mõistlik kombineeritud ja etapiviisiline lahendus ning programmi tasandil mitu projekti."),
  P("**Millist tagasisidet saime tehisarult ja kas täiendasime tööd?** Tehisaru tagasiside ja selle põhjal tehtud muudatused on esitatud lisas 3."),
  P("**Kuidas tagasime rühmas usalduse ja vastutuse?** Igal objektil on nimeline vastutaja ja kaasteostaja (lisa 1), kes kontrollib vastutaja tööd enne vahetähtaega (nelja silma põhimõte). Otsused (nt tehnika valik ja kriteeriumide kaalud) tehti koosolekul konsensusega ja kaalud fikseeriti enne hindamist. Kõik versioonid ja tehisaruga peetud vestlused on ühises kaustas kõigile nähtavad, mis loob läbipaistvuse ja võimaldab panust hinnata."),

  H2("Lisa 3. Vestlus tehisaruga ja selle põhjal tehtud täiendused"),
  P("Tehisaruna kasutati keelemudelit Claude (Anthropic, 2026). Allpool on vestluse sisu lühendatult; täielik vestlus on rühma ühises kaustas kõigile liikmetele nähtav."),
  P("**Rühma päring (lühendatult):** „Vaadake üle slaidikogu 2 slaidid 1–36 ning muud seonduvad juhendid. Esitage terviklikku peatükkide struktuuri sisaldav DOCX-fail, milles on alapeatükk 1.2: metoodika (sh tehisaru ja teaduslike allikate kasutus), probleemide puu, eesmärkide puu, huvigruppide tabel ja kaasamise strateegia, alternatiivide võrdlus; lisad tööjaotuse ja õppetundidega. Joonistel kasutada jooniste ala ja ühendajaid, joondust ning pealdiseid.“"),
  P("**Tehisaru vastus ja tagasiside töö versioonile:**"),
  N("Probleemipuusse ei tohi panna projekti riske (nt „projekt võib hilineda“) – probleem kirjeldab olukorda enne projekti (Tamberg, 2022, slaid 11). *Kontrollitud: probleemipuus on ainult olemasoleva olukorra puudujäägid.*", "num2"),
  N("Eesmärgid peavad olema sõnastatud püsiseisundina, mitte tegevusena („siserand on ehitatud“ on tulem, mitte eesmärk). *Arvestatud: eesmärgipuus on seisundid, siserand on käsitletud alternatiivina.*", "num2"),
  N("Alternatiivide võrdluses peavad kaalud olema fikseeritud enne hindamist ning tulemuse stabiilsust tuleb kontrollida. *Täiendatud: lisati tundlikkusanalüüs, mis näitas, et kulu suurema kaalu korral on A ja A+C võrdsed – soovitust täpsustati etapiviisiliseks.*", "num2"),
  N("Huvipoolte tabelist peab selguma kaasamise strateegia ja vastandlikud huvid. *Arvestatud: tabel 3 ja kaasamise strateegia IAP2 tasemete kaupa.*", "num2"),
  N("Avalikud statistikanäitajad (täituvus 39–50% vs 60–80%) tuleb kontrollida algallikast (Statistikaamet, arengukava) ja märkida aasta. *Avatud: kontrollib " + M.B + " enne esitamist.*", "num2"),
  N("Tehisaru ei pääsenud ligi LinkedIni artiklile ega kolledži blanketile – rühm peab artikli ise läbi lugema ja töö blanketti üle tõstma. *Avatud: " + M.A + " ja " + M.D + ".*", "num2"),
  P("**Rühma hinnang tehisaru kasutamisele:** tehisaru kiirendas struktuuri loomist ja aitas leida nõrku kohti, kuid sisulised hinnangud, kohalik kontekst ja allikate kontroll jäid rühma vastutusele."),
);

// ---------- dokument ----------
const doc = new Document({
  creator: "Rühm",
  title: "Pärnu aastaringse kuurortlinna arendamine",
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
