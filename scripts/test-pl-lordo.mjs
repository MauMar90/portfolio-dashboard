// Test (jsdom) del KPI "P/L % lordo" della home in index.html:
//   P/L% lordo = Σ(valore mercato − costo di carico) / Σ costo di carico × 100, solo posizioni aperte,
//   aggregando prima gli euro (mai media delle percentuali di riga). Il costo di carico è il PMC
//   della dashboard, già senza commissioni per gli ETF (vedi handleFile). Prezzo di mercato 0 =
//   dato mancante: KPI non disponibile ("—") con avviso visibile.
// Solo dati fittizi. Uso: node scripts/test-pl-lordo.mjs

import { JSDOM } from 'jsdom';
import fs from 'node:fs/promises';
import assert from 'node:assert/strict';
import path from 'node:path';
import { fileURLToPath } from 'node:url';

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const html = await fs.readFile(path.join(__dirname, '..', 'index.html'), 'utf8');

function load(seedPositions = [], xlsxRows = null) {
  const dom = new JSDOM(html, {
    url: 'http://localhost/',
    runScripts: 'dangerously',
    resources: undefined,
    pretendToBeVisual: true,
    beforeParse(win) {
      win.localStorage.setItem('pfPositions', JSON.stringify(seedPositions));
      win.fetch = async () => { throw new Error('offline nei test'); };
      win.Element.prototype.scrollIntoView = () => {};
      if (xlsxRows) {
        win.XLSX = {
          read: () => ({ SheetNames: ['S'], Sheets: { S: xlsxRows } }),
          utils: { sheet_to_json: ws => ws },
        };
      }
    },
  });
  return dom.window;
}

const kpi = win => {
  const el = win.document.getElementById('kPLpct');
  const warn = win.document.getElementById('kPLpctWarn');
  assert.ok(el, 'deve esistere il KPI #kPLpct nella home');
  assert.ok(warn, "deve esistere l'avviso #kPLpctWarn");
  return { text: el.textContent.trim(), warnText: warn.textContent.trim(), warnVisible: warn.style.display !== 'none' };
};
const eur = s => parseFloat(s.replace('€', '').replace(/\./g, '').replace(',', '.').trim());

// ── A. ETF con commissione, convenzione export Fineco (fix 6f62137) ──
{
  const HEADER = ['Operazione', 'Data valuta', 'Descrizione', 'Titolo', 'Isin', 'Segno', 'Quantita',
    'Divisa', 'Prezzo', 'Cambio', 'Controvalore', 'Commissioni amministrato'];
  const rows = [['Intestazione fittizia'], HEADER,
    ['02/01/2025', '02/01/2025', 'Compravendita titoli', 'TITOLO TEST', 'IE00B4L5Y983', 'A', 100, 'EUR', 10.0295, 1, 1002.95, 2.95]];
  const win = load([], rows);
  win.handleFile(new win.File(['x'], 'fittizio.xlsx'));
  await new Promise(r => setTimeout(r, 100));
  win.applyImport();
  win.eval("positions.find(p => p.isin === 'IE00B4L5Y983').mkt = 11; render();");
  // confronto numerico: in it-IT i numeri a 4 cifre non hanno separatore delle migliaia ("1000,00")
  assert.ok(Math.abs(eur(win.document.getElementById('kCost').textContent) - 1000) < 0.005, 'costo di carico senza commissione = 1.000');
  assert.ok(Math.abs(eur(win.document.getElementById('kPL').textContent) - 100) < 0.005, 'P/L = +100');
  const k = kpi(win);
  assert.equal(k.text, '+10,00%', 'A: P/L% lordo');
  assert.equal(k.warnVisible, false, 'A: nessun avviso');
  console.log('✓ A. ETF con commissione → +10,00%');
}

// ── B. Coerenza con gli euro della home (ETF + BTP aperto, scala isNominal) ──
{
  const win = load([
    { name: 'ETF1', isin: 'XX0000000001', ac: 'Azioni', qty: 10, price: 50, mkt: 55, closed: false },        // costo 500, P/L +50
    { name: 'ETF2', isin: 'XX0000000002', ac: 'Azioni', qty: 4, price: 200, mkt: 190, closed: false },       // costo 800, P/L −40
    { name: 'BTP Futura', isin: 'IT0005442089', ac: 'Obbligazioni', qty: 3, price: 97.5, mkt: 98.2,
      closed: false, isNominal: true },                                                                     // costo 2.925, P/L +21
  ]);
  const rowPL = [...win.document.querySelectorAll('#holdBody tr.main-row')].map(tr => eur(tr.children[4].textContent));
  const sumPL = rowPL.reduce((s, x) => s + x, 0);
  assert.equal(rowPL.length, 3, 'tre righe aperte');
  assert.ok(Math.abs(sumPL - 31) < 0.005, 'Σ P/L righe = 31: ' + sumPL);
  assert.ok(Math.abs(eur(win.document.getElementById('kPL').textContent) - sumPL) < 0.005, 'kPL = Σ P/L righe');
  assert.ok(Math.abs(eur(win.document.getElementById('kCost').textContent) - 4225) < 0.005, 'kCost = Σ costo righe = 4.225');
  assert.equal(kpi(win).text, win.eval('fmtPct')(31 / 4225 * 100), 'B: P/L% = Σ P/L / Σ costo × 100');
  assert.equal(kpi(win).text, '+0,73%', 'B: valore atteso');
  console.log('✓ B. Coerenza con kPL/kCost (incluso BTP) → +0,73%');
}

// ── C. Più titoli: aggregazione in euro, NON media delle percentuali ──
const posC = [
  { name: 'P1', isin: 'XX0000000003', ac: 'Azioni', qty: 1, price: 100, mkt: 120, closed: false }, // +20 su 100 = +20%
  { name: 'P2', isin: 'XX0000000004', ac: 'Azioni', qty: 9, price: 100, mkt: 110, closed: false }, // +90 su 900 = +10%
];
{
  const k = kpi(load(posC));
  assert.equal(k.text, '+11,00%', 'C: (20 + 90) / (100 + 900) = 11%');
  assert.notEqual(k.text, '+15,00%', 'C: non deve essere la media delle percentuali');
  console.log('✓ C. Due titoli → +11,00% (non 15%)');
}

// ── D. Posizione chiusa (con P/L realizzato e mkt 0) esclusa da numeratore e denominatore ──
{
  const win = load([...posC,
    { name: 'Chiusa', isin: 'XX0000000005', ac: 'Azioni', qty: 0, price: 30, mkt: 0, closed: true,
      realizedPLgross: 500, realizedPLnet: 370 }]);
  const k = kpi(win);
  assert.equal(k.text, '+11,00%', 'D: la chiusa non cambia il KPI');
  assert.equal(k.warnVisible, false, 'D: mkt 0 di una posizione CHIUSA non genera avviso');
  console.log('✓ D. Posizione chiusa esclusa → +11,00%');
}

// ── E. Posizione aperta con prezzo di mercato 0 → KPI non disponibile + avviso ──
{
  const win = load([...posC,
    { name: 'SenzaPrezzo', isin: 'XX0000000006', ac: 'Azioni', qty: 5, price: 10, mkt: 0, closed: false }]);
  const k = kpi(win);
  assert.equal(k.text, '—', 'E: KPI non disponibile');
  assert.equal(k.warnVisible, true, 'E: avviso visibile');
  assert.ok(k.warnText.includes('SenzaPrezzo'), "E: l'avviso nomina la posizione senza prezzo: " + k.warnText);

  win.eval("positions.find(p => p.name === 'SenzaPrezzo').mkt = 10; render();");
  const k2 = kpi(win);
  assert.equal(k2.text, '+10,48%', 'E: con tutti i prezzi torna numerico (110 / 1.050 = 10,476%)');
  assert.equal(k2.warnVisible, false, 'E: avviso nascosto quando i prezzi ci sono');
  console.log('✓ E. Prezzo mancante → "—" con avviso; torna numerico quando il prezzo c\'è');
}

console.log('\nTutti i test P/L% lordo superati.');
