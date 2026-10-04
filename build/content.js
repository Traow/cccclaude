// Töö sisu: alternatiivide hindamine ja abiandmed (kasutatakse build_doc.js-is)

const members = {
  A: "Liige A",
  B: "Liige B",
  C: "Liige C",
  D: "Liige D",
};

// Lahendusalternatiivide hindamiskriteeriumid ja kaalud (summa 100%)
const criteria = [
  { id: "K1", name: "Mõju põhieesmärgile (madalhooaja külastatavus ja tulu)", w: 0.25 },
  { id: "K2", name: "Eesmärgipuu alameesmärkide katvus (O1–O3)", w: 0.20 },
  { id: "K3", name: "Kulu ja rahastatavus", w: 0.15 },
  { id: "K4", name: "Teostatavus ja aeg esimese tulemuseni", w: 0.15 },
  { id: "K5", name: "Riskitase (5 = madal risk)", w: 0.10 },
  { id: "K6", name: "Huvipoolte toetus, vastuolude vähesus", w: 0.15 },
];

const alternatives = [
  { id: "A", name: "„Talvine Pärnu“ sündmuste programm ja ühisturundus", scores: [4, 3, 4, 5, 4, 4] },
  { id: "B", name: "Siseranna ehk aastaringse vee- ja vabaajakeskuse rajamine", scores: [5, 3, 1, 1, 2, 3] },
  { id: "C", name: "Digitaalne sihtkohaplatvorm „Pärnu Pass“", scores: [3, 3, 4, 3, 3, 4] },
  { id: "A+C", name: "A ja C etapiviisiline kombinatsioon", scores: [5, 4, 3, 4, 4, 4] },
  { id: "0", name: "Nullalternatiiv (senine tegevus)", scores: [1, 1, 5, 5, 3, 1] },
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
