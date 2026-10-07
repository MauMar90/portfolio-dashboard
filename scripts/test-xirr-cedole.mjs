// Test (jsdom): le cedole entrano nei flussi XIRR una per una, alla loro data reale e AL LORDO
// della ritenuta (netto / (1 − aliquota del titolo)), così tutti i flussi XIRR sono "prima delle
// imposte" come controvalori e valore finale. Realizzato e "Cedole incassate" restano invariati.
// Cedola su un titolo non di Stato → avviso. Solo dati fittizi.
// L'XIRR viene confrontato con un calcolo indipendente (bisezione, stessa convenzione
// giorni/365,25 di xirr()). Uso: node scripts/test-xirr-cedole.mjs

import { JSDOM } from 'jsdom';
import fs from 'node:fs/promises';
import assert from 'node:assert/strict';
import path from 'node:path';
import { fileURLToPath } from 'node:url';

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const html = await fs.readFile(path.join(__dirname, '..', 'index.html'), 'utf8');

const HEADER = ['Operazione', 'Data valuta', 'Descrizione', 'Titolo', 'Isin', 'Segno', 'Quantita',
  'Divisa', 'Prezzo', 'Cambio', 'Controvalore', 'Commissioni amministrato'];
const BTP = 'IT0005442089', ETF = 'IE00B4L5Y983';
const ALIQUOTA_BTP = 0.125;
const ROWS = [['Intestazione fittizia'], HEADER,
  // acquisto 10.000 nominali a 100 (commissione 0) → flusso −10.000
  ['02/01/2024', '04/01/2024', 'Compravendita titoli', 'BTP TEST', BTP, 'A', 10000, 'EUR', 100, 1, 10000, 0],
  // due cedole NETTE da 87,50 (lorde 100 = 87,50 / 0,875), date diverse
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

async function importa(rows) {
  const win = new JSDOM(html, {
    url: 'http://localhost/', runScripts: 'dangerously', resources: undefined, pretendToBeVisual: true,
    beforeParse(w) {
      w.fetch = async () => { throw new Error('offline nei test'); };
      w.XLSX = { read: () => ({ SheetNames: ['S'], Sheets: { S: rows } }), utils: { sheet_to_json: ws => ws } };
    },
  }).window;
  win.handleFile(new win.File(['x'], 'fittizio.xlsx'));
  await new Promise(r => setTimeout(r, 100));
  return win;
}

const win = await importa(ROWS);
const d = win.eval('importedData').find(x => x.isin === BTP);

// 1. Flussi: acquisto + una cedola LORDA per data reale; nessun flusso alla data di oggi
const flows = [...d.cashFlows].map(f => ({ amount: Math.round(f.amount * 100) / 100, date: f.date.slice(0, 10) })); // copia nel realm di Node
assert.deepEqual(flows, [
  { amount: -10000, date: new Date(2024, 0, 2).toISOString().slice(0, 10) },
  { amount: 100,    date: new Date(2024, 6, 1).toISOString().slice(0, 10) },
  { amount: 100,    date: new Date(2025, 0, 2).toISOString().slice(0, 10) },
], 'flussi XIRR: una cedola lorda (87,50 / 0,875 = 100) per riga, alla sua data reale');
console.log('✓ 1. Due cedole come flussi separati alle date reali, al lordo della ritenuta (100 = 87,50 / 0,875)');

// 2. XIRR della dashboard = XIRR calcolato a mano sui flussi lordi con le date reali
const conFinale = [...d.cashFlows.map(f => ({ amount: f.amount, date: new Date(f.date) })),
                   { amount: 10000, date: new Date(2026, 0, 2) }];
const xDashboard = win.xirr(conFinale);
const xManuale = xirrManuale([
  { amount: -10000, date: new Date(2024, 0, 2) }, { amount: 100, date: new Date(2024, 6, 1) },
  { amount: 100, date: new Date(2025, 0, 2) },    { amount: 10000, date: new Date(2026, 0, 2) }]);
assert.ok(Math.abs(xDashboard - xManuale) < 1e-6, `XIRR dashboard ${xDashboard} = manuale ${xManuale}`);
console.log(`✓ 2. XIRR ${xDashboard.toFixed(6)}% = calcolo manuale sui flussi lordi ${xManuale.toFixed(6)}%`);

// 3. Ogni riga "Stacco Cedole": flusso XIRR = Controvalore / (1 − aliquota), alla data della riga
for (const r of ROWS.filter(r => r[2] === 'Stacco Cedole')) {
  const [g, m, a] = r[0].split('/').map(Number);
  const dataAttesa = new Date(a, m - 1, g).toISOString();
  const atteso = r[10] / (1 - ALIQUOTA_BTP);
  const ok = d.cashFlows.some(f => Math.abs(f.amount - atteso) < 0.005 && f.date === dataAttesa);
  assert.ok(ok, `cedola del ${r[0]}: flusso ${atteso} alla sua data non trovato`);
}
console.log('✓ 3. Ogni riga cedola: flusso XIRR = Controvalore / (1 − 12,5%), alla data della riga');

// 4. Realizzato e cedole incassate invariati (netto 175, lordo 200, cedole 175)
assert.equal(d.cedole, 175, 'cedole incassate (nette) invariate');
assert.ok(Math.abs(d.realizedPLnet - 175) < 0.005 && Math.abs(d.realizedPLgross - 200) < 0.005, `realizzato invariato: netto ${d.realizedPLnet}, lordo ${d.realizedPLgross}`);
console.log('✓ 4. Realizzato invariato (netto 175, lordo 200) e cedole incassate nette (175)');

// 5. Cedola su un titolo non di Stato (aliquota ≠ 12,5%) → avviso; su BTP nessun avviso
{
  const avvisiBtp = [...win.eval('importWarnings')];
  assert.ok(!avvisiBtp.some(a => /aliquota di tabella/.test(a)), 'nessun avviso per cedole BTP');
  const w2 = await importa([['Intestazione fittizia'], HEADER,
    ['01/07/2024', '01/07/2024', 'Stacco Cedole', 'TITOLO TEST', ETF, '', 10, 'EUR', 0, 1, 7.4, '']]);
  const avvisi = [...w2.eval('importWarnings')];
  assert.ok(avvisi.some(a => a.includes(ETF) && /cedola riportata al lordo con l'aliquota di tabella, verificare la ritenuta effettiva/.test(a)),
    'avviso cedola su titolo non di Stato: ' + avvisi.join(' | '));
  console.log('✓ 5. Cedola su titolo non di Stato: avviso "riportata al lordo con l\'aliquota di tabella"; BTP senza avviso');
}

console.log('\nTutti i test XIRR cedole superati.');
