// Test (jsdom): le cedole entrano nei flussi XIRR una per una, alla loro data reale
// (prima: un unico flusso col totale alla data di oggi). Solo dati fittizi.
// L'XIRR dei flussi viene confrontato con un calcolo indipendente (bisezione, stessa
// convenzione giorni/365,25 di xirr()). Uso: node scripts/test-xirr-cedole.mjs

import { JSDOM } from 'jsdom';
import fs from 'node:fs/promises';
import assert from 'node:assert/strict';
import path from 'node:path';
import { fileURLToPath } from 'node:url';

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const html = await fs.readFile(path.join(__dirname, '..', 'index.html'), 'utf8');

const HEADER = ['Operazione', 'Data valuta', 'Descrizione', 'Titolo', 'Isin', 'Segno', 'Quantita',
  'Divisa', 'Prezzo', 'Cambio', 'Controvalore', 'Commissioni amministrato'];
const BTP = 'IT0005442089';
const ROWS = [['Intestazione fittizia'], HEADER,
  // acquisto 10.000 nominali a 100 (commissione 0) → flusso −10.000
  ['02/01/2024', '04/01/2024', 'Compravendita titoli', 'BTP TEST', BTP, 'A', 10000, 'EUR', 100, 1, 10000, 0],
  // due cedole nette (Quantità = nominale), date diverse
  ['01/07/2024', '01/07/2024', 'Stacco Cedole', 'BTP TEST', BTP, '', 10000, 'EUR', 0, 1, 87.5, ''],
  ['02/01/2025', '02/01/2025', 'Stacco Cedole', 'BTP TEST', BTP, '', 10000, 'EUR', 0, 1, 87.5, ''],
];

// XIRR indipendente: bisezione su NPV con anni = giorni / 365,25 (come xirr() in index.html)
function xirrManuale(flows) {
  const t0 = Math.min(...flows.map(f => f.date.getTime()));
  const npv = r => flows.reduce((s, f) => s + f.amount / Math.pow(1 + r, (f.date.getTime() - t0) / (365.25 * 864e5)), 0);
  let lo = -0.99, hi = 1;
  for (let i = 0; i < 200; i++) { const mid = (lo + hi) / 2; if (npv(lo) * npv(mid) <= 0) hi = mid; else lo = mid; }
  return ((lo + hi) / 2) * 100;
}

const dom = new JSDOM(html, {
  url: 'http://localhost/', runScripts: 'dangerously', resources: undefined, pretendToBeVisual: true,
  beforeParse(win) {
    win.fetch = async () => { throw new Error('offline nei test'); };
    win.XLSX = { read: () => ({ SheetNames: ['S'], Sheets: { S: ROWS } }), utils: { sheet_to_json: ws => ws } };
  },
});
const win = dom.window;
win.handleFile(new win.File(['x'], 'fittizio.xlsx'));
await new Promise(r => setTimeout(r, 100));
const d = win.eval('importedData').find(x => x.isin === BTP);

// 1. Flussi: acquisto + una cedola per data, nessun flusso "cedole totali" alla data di oggi
const flows = [...d.cashFlows].map(f => ({ amount: Math.round(f.amount * 100) / 100, date: f.date.slice(0, 10) })); // copia nel realm di Node
assert.deepEqual(flows, [
  { amount: -10000, date: new Date(2024, 0, 2).toISOString().slice(0, 10) },
  { amount: 87.5,   date: new Date(2024, 6, 1).toISOString().slice(0, 10) },
  { amount: 87.5,   date: new Date(2025, 0, 2).toISOString().slice(0, 10) },
], 'flussi XIRR: una cedola per riga, alla sua data reale');
assert.equal(d.cedole, 175, 'totale cedole (display e realizzato) invariato');
console.log('✓ 1. Due cedole come flussi separati alle date reali; totale cedole invariato (175)');

// 2. XIRR della dashboard sui flussi + valore finale a data fissa = XIRR calcolato a mano
const conFinale = [...d.cashFlows.map(f => ({ amount: f.amount, date: new Date(f.date) })),
                   { amount: 10000, date: new Date(2026, 0, 2) }];
const xDashboard = win.xirr(conFinale);
const xManuale = xirrManuale(conFinale);
assert.ok(Math.abs(xDashboard - xManuale) < 1e-6, `XIRR dashboard ${xDashboard} = manuale ${xManuale}`);
console.log(`✓ 2. XIRR ${xDashboard.toFixed(6)}% = calcolo manuale ${xManuale.toFixed(6)}%`);

// 3. Ogni riga "Stacco Cedole" dell'export ha un flusso XIRR uguale al suo Controvalore, alla sua data
for (const r of ROWS.filter(r => r[2] === 'Stacco Cedole')) {
  const [g, m, a] = r[0].split('/').map(Number);
  const dataAttesa = new Date(a, m - 1, g).toISOString();
  const ok = d.cashFlows.some(f => Math.abs(f.amount - r[10]) < 0.005 && f.date === dataAttesa);
  assert.ok(ok, `cedola del ${r[0]}: flusso ${r[10]} alla sua data non trovato`);
}
console.log('✓ 3. Ogni riga cedola: flusso XIRR = Controvalore, alla data della riga');

console.log('\nTutti i test XIRR cedole superati.');
