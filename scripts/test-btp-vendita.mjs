// Test (jsdom) della convenzione export Fineco per il BTP Futura (isNominal) e dei flussi XIRR.
// Verificato su un eseguito reale Fineco: per il BTP il Prezzo export × nominale / 100 è il
// "Totale" della vendita = corso secco + rateo netto − commissione (commissione GIÀ dedotta,
// a differenza degli ETF dove il prezzo va corretto di + comm/qty). Solo dati fittizi.
// Uso: node scripts/test-btp-vendita.mjs

import { JSDOM } from 'jsdom';
import fs from 'node:fs/promises';
import assert from 'node:assert/strict';
import path from 'node:path';
import { fileURLToPath } from 'node:url';

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const html = await fs.readFile(path.join(__dirname, '..', 'index.html'), 'utf8');

const HEADER = ['Operazione', 'Data valuta', 'Descrizione', 'Titolo', 'Isin', 'Segno', 'Quantita',
  'Divisa', 'Prezzo', 'Cambio', 'Controvalore', 'Commissioni amministrato'];
const BTP = 'IT0005442089', BTP_V = 'IT0005442097', ETF = 'IE00B4L5Y983', ORO = 'FR0013416716';
// riga export: Controvalore = quantità × prezzo (per il BTP: nominale × prezzo / 100)
const T = (d, isin, s, q, p, c) => [d, d, 'Compravendita titoli', 'TITOLO TEST', isin, s, q, 'EUR', p, 1,
  Math.round((isin.startsWith('IT00054420') ? q * p / 100 : q * p) * 100) / 100, c];

async function importa(rows) {
  const dom = new JSDOM(html, {
    url: 'http://localhost/', runScripts: 'dangerously', resources: undefined, pretendToBeVisual: true,
    beforeParse(win) {
      win.fetch = async () => { throw new Error('offline nei test'); };
      win.XLSX = { read: () => ({ SheetNames: ['S'], Sheets: { S: [['Intestazione fittizia'], HEADER, ...rows] } }),
                   utils: { sheet_to_json: ws => ws } };
    },
  });
  const win = dom.window;
  win.handleFile(new win.File(['x'], 'fittizio.xlsx'));
  await new Promise(r => setTimeout(r, 100));
  return { data: win.eval('importedData'), warnings: win.eval('importWarnings').join('\n') };
}

// ── 1. Vendita BTP in formato export ──
// Acquisto 5.000 nominali a 100 (commissione 0). Vendita 1.000 nominali (ISIN …097):
// corso secco 80, rateo netto 2 €, commissione 10 € → Totale 800 + 2 − 10 = 792 → Prezzo export 79,2.
// Atteso dalla dashboard: realizzato −208 = 792 − 1.000.
// Scarto NOTO rispetto a Fineco: Fineco calcola (80 × 10 − 10) − 1.000 = −210; la differenza (+2)
// è il rateo netto, che l'export movimenti non riporta e quindi non si può separare.
{
  const { data } = await importa([T('03/05/2021', BTP, 'A', 5000, 100, 0), T('07/03/2025', BTP_V, 'V', 1000, 79.2, 10)]);
  const d = data.find(x => x.isin === BTP);
  assert.equal(d.qty, 4, 'quote residue (4.000 nominali)');
  assert.ok(Math.abs(d.price - 100) < 1e-9, 'PMC 100: ' + d.price);
  assert.ok(Math.abs(d.realizedPLgross - (-208)) < 0.005, 'realizzato lordo −208 (Fineco −210, scarto = rateo netto): ' + d.realizedPLgross);
  assert.ok(Math.abs(d.realizedPLnet - (-208)) < 0.005, 'minusvalenza: nessuna imposta, netto = lordo: ' + d.realizedPLnet);
  const amounts = [...d.cashFlows].map(f => Math.round(f.amount * 100) / 100); // copia nel realm di Node (deepEqual stretto)
  assert.deepEqual(amounts, [-5000, 792], 'flussi XIRR: −5.000 / +792 (= Totale incassato)');
  console.log('✓ 1. Vendita BTP: realizzato −208 (Fineco −210, scarto = rateo netto), flusso +792, PMC 100');
}

// ── 2. Ogni riga di compravendita: flusso XIRR = Controvalore (− acquisti, + vendite) ──
{
  const rows = [
    T('02/01/2025', ETF, 'A', 100, 10.0295, 2.95), T('03/02/2025', ETF, 'V', 40, 11.95, 2),
    T('01/07/2025', ORO, 'A', 8, 140, 0),
    T('03/05/2021', BTP, 'A', 5000, 100, 0), T('07/03/2025', BTP_V, 'V', 1000, 79.2, 10),
  ];
  const { data } = await importa(rows);
  for (const r of rows) {
    const chiave = r[4] === BTP_V ? BTP : r[4];
    const d = data.find(x => x.isin === chiave);
    const atteso = (r[5] === 'A' ? -1 : 1) * r[10];
    const trovato = d.cashFlows.some(f => Math.abs(f.amount - atteso) < 0.005);
    assert.ok(trovato, `${r[4]} ${r[5]} del ${r[0]}: flusso ${atteso} non trovato in ${JSON.stringify(d.cashFlows.map(f => f.amount))}`);
  }
  console.log(`✓ 2. ${rows.length} righe di compravendita: ogni flusso XIRR coincide con il Controvalore`);
}

// ── 3. Avvisi visibili su convenzioni non verificate (calcoli invariati) ──
{
  const pulito = await importa([T('03/05/2021', BTP, 'A', 5000, 100, 0), T('01/07/2025', ORO, 'A', 8, 140, 0)]);
  assert.ok(!/non verificata/.test(pulito.warnings), 'nessun avviso con BTP a 100 senza commissione e oro senza commissione');

  const btpComm = await importa([T('03/05/2021', BTP, 'A', 5000, 100, 9.95)]);
  assert.ok(/BTP/.test(btpComm.warnings) && /non verificata/.test(btpComm.warnings), 'avviso BTP con commissione: ' + btpComm.warnings);

  const btpPrezzo = await importa([T('03/05/2021', BTP, 'A', 5000, 98, 0)]);
  assert.ok(/BTP/.test(btpPrezzo.warnings) && /non verificata/.test(btpPrezzo.warnings), 'avviso BTP con prezzo ≠ 100: ' + btpPrezzo.warnings);

  const oro = await importa([T('01/07/2025', ORO, 'A', 8, 140, 5)]);
  assert.ok(/oro/i.test(oro.warnings) && /non verificata/.test(oro.warnings), 'avviso oro con commissione: ' + oro.warnings);
  console.log('✓ 3. Avvisi: BTP con commissione, BTP a prezzo ≠ 100, oro con commissione; nessun avviso nei casi puliti');
}

console.log('\nTutti i test BTP superati.');
