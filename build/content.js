// Töö sisu: alternatiivide hindamine ja abiandmed (kasutatakse build_doc.js-is)

const members = {
  PJ: "Herman Ra Truvek",
  AI: "Robi Mustsaar",
  EX: "Hugo-Christopher Saar",
  OM: "Ragnar Dietrich",
};

// Lahendusalternatiivide hindamiskriteeriumid ja kaalud (summa 100%)
const criteria = [
  { id: "K1", name: "Mõju fännibaasi taasühendamisele", w: 0.25 },
  { id: "K2", name: "Eesmärgipuu alameesmärkide katvus (O1–O3)", w: 0.20 },
  { id: "K3", name: "Mahtumine eelarvesse (96 800 €)", w: 0.15 },
  { id: "K4", name: "Teostatavus 3 kuu jooksul", w: 0.15 },
  { id: "K5", name: "Riskitase (5 = madal risk)", w: 0.10 },
  { id: "K6", name: "Huvipoolte toetus, vastuolude vähesus", w: 0.15 },
];

const alternatives = [
  { id: "A", name: "Ajaloolise ovaalse vapi täielik taastamine", scores: [3, 2, 2, 2, 2, 2] },
  { id: "B", name: "Kahetasandiline bränd: „J“ + pärandvapp", scores: [4, 4, 4, 4, 4, 4] },
  { id: "C", name: "„J“ logo ümberkujundamine koos fännidega", scores: [4, 3, 2, 2, 3, 3] },
  { id: "D", name: "Ainult kogukonna- ja kommunikatsiooniprogramm", scores: [2, 2, 5, 5, 4, 3] },
  { id: "B+D", name: "Pärandvapp koos fännide kaasamise programmiga", scores: [5, 5, 3, 4, 4, 5] },
  { id: "0", name: "Nullalternatiiv (senine bränd)", scores: [1, 1, 5, 5, 2, 1] },
];

function weighted(scores, weights) {
  return scores.reduce((s, v, i) => s + v * weights[i], 0);
}

const baseWeights = criteria.map((c) => c.w);
alternatives.forEach((a) => (a.total = weighted(a.scores, baseWeights)));

// Tundlikkusanalüüs: kulu kaal 30%, ülejäänud proportsionaalselt vähendatud
function reweight(idx, newW) {
  const rest = 1 - baseWeights[idx];
  return baseWeights.map((w, i) => (i === idx ? newW : (w * (1 - newW)) / rest));
}
const sensCost = reweight(2, 0.3);
const sensRisk = reweight(4, 0.25);
alternatives.forEach((a) => {
  a.sensCost = weighted(a.scores, sensCost);
  a.sensRisk = weighted(a.scores, sensRisk);
});

module.exports = { members, criteria, alternatives };
