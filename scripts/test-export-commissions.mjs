// Test di regressione (jsdom) per la convenzione commissioni dell'export movimenti Fineco
// dentro handleFile() di index.html:
//   - acquisti: Prezzo export = esecuzione + comm/qty  (commissione già inclusa)
//   - vendite:  Prezzo export = esecuzione - comm/qty  (commissione già dedotta)
// PMC = prezzo di esecuzione senza commissioni (come il "P.zo medio di carico" Fineco),
// flussi XIRR = controvalore (qty × Prezzo export), nessuna commissione applicata due volte.
// Solo dati fittizi. Uso: node scripts/test-export-commissions.mjs

import { JSDOM } from 'jsdom';
import fs from 'node:fs/promises';
import assert from 'node:assert/strict';
import path from 'node:path';
import { fileURLToPath } from 'node:url';

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const html = await fs.readFile(path.join(__dirname, '..', 'index.html'), 'utf8');

const HEADER = ['Operazione', 'Data valuta', 'Descrizione', 'Titolo', 'Isin', 'Segno', 'Quantita',
  'Divisa', 'Prezzo', 'Cambio', 'Controvalore', 'Commissioni amministrato'];
const ETF_ISIN = 'IE00B4L5Y983'; // ISIN pubblico presente in FINECO_MAP; operazioni fittizie
const BTP_ISIN = 'IT0005442089';

// riga export: [data, data, desc, titolo, isin, segno, qty, divisa, prezzo, cambio, controvalore, comm]
const row = (date, isin, segno, qty, price, comm) =>
  [date, date, 'Compravendita titoli', 'TITOLO TEST', isin, segno, qty, 'EUR', price, 1,
   (segno === 'A' ? 1 : -1) * qty * price, comm];

async function importRows(rows) {
  const dom = new JSDOM(html, {
    url: 'http://localhost/',
    runScripts: 'dangerously',
    resources: undefined, // niente CDN: XLSX viene sostituito da uno stub che restituisce `rows`
    pretendToBeVisual: true,
    beforeParse(win) {
      win.fetch = async () => { throw new Error('offline nei test'); };
      win.Element.prototype.scrollIntoView = () => {};
      win.XLSX = {
        read: () => ({ SheetNames: ['S'], Sheets: { S: [['Intestazione fittizia'], HEADER, ...rows] } }),
        utils: { sheet_to_json: ws => ws },
      };
    },
  });
  const win = dom.window;
  win.handleFile(new win.File(['x'], 'fittizio.xlsx'));
  await new Promise(r => setTimeout(r, 100)); // FileReader asincrono
  const data = win.eval('importedData');
  assert.ok(data.length > 0, 'handleFile deve produrre importedData');
  return { win, data };
}

const close = (actual, expected, msg, eps = 1e-9) =>
  assert.ok(Math.abs(actual - expected) <= eps, `${msg}: atteso ${expected}, ottenuto ${actual}`);

// ── 1. PMC con commissione: 100 quote, Prezzo export 10,0295, comm 2,95 → esecuzione 10,0000 ──
{
  const { win, data } = await importRows([row('02/01/2025', ETF_ISIN, 'A', 100, 10.0295, 2.95)]);
  const d = data[0];
  close(d.price, 10, 'PMC = prezzo di esecuzione senza commissioni');
  assert.equal(d.commTot, 2.95, 'commTot (solo informativo) invariato');

  win.applyImport();
  win.eval(`positions.find(p => p.isin === '${ETF_ISIN}').mkt = 11; render();`);
  const cells = [...win.document.querySelectorAll('#holdBody tr.main-row td')].map(td => td.textContent.trim());
  assert.equal(cells[5], '+10,00%', 'P/L% di riga a prezzo 11 deve essere +10,00%');
  assert.equal(cells[4], '€ 100,00', 'P/L € di riga = 100 × (11 − 10)');
  console.log('✓ 1. PMC con commissione inclusa nel Prezzo export → 10,0000 e P/L% +10,00%');
}

// ── 2. Vendita: Prezzo export già netto di commissione, non va sottratta di nuovo ──
// Acquisto 100 @ 10,0295 (comm 2,95 → esec. 10,00); vendita 40 @ 11,95 (comm 2 → esec. 12,00).
// P/L realizzato lordo = 40 × (12,00 − 10,00) = 80,00 (come il P&L Fineco: esecuzione vs esecuzione).
{
  const { data } = await importRows([
    row('02/01/2025', ETF_ISIN, 'A', 100, 10.0295, 2.95),
    row('03/02/2025', ETF_ISIN, 'V', 40, 11.95, 2),
  ]);
  const d = data[0];
  assert.equal(d.qty, 60, 'quote residue');
  close(d.price, 10, 'PMC invariato dopo vendita parziale');
  close(d.realizedPLgross, 80, 'P/L realizzato lordo senza doppio conteggio commissioni', 0.005);
  close(d.realizedPLnet, 80 * (1 - 0.26), 'P/L realizzato netto (aliquota _default invariata)', 0.005);
  console.log('✓ 2. Vendita: realizzato lordo 80,00 (commissione non sottratta due volte)');

  // ── 3. XIRR: flussi = ∓ qty × Prezzo export, nessun aggiustamento per comm ──
  assert.equal(d.cashFlows.length, 2, 'due flussi (nessuna cedola)');
  close(d.cashFlows[0].amount, -(100 * 10.0295), 'flusso acquisto = −qty × Prezzo export');
  close(d.cashFlows[1].amount, +(40 * 11.95), 'flusso vendita = +qty × Prezzo export');
  console.log('✓ 3. Flussi XIRR = controvalori export (−1002,95 / +478,00)');
}

// ── 4. Posizione chiusa e riaperta: il costo medio riparte correttamente ──
// A 10 @ 10,10 (comm 1 → esec. 10,00); V 10 @ 10,90 (comm 1 → esec. 11,00); A 5 @ 12,20 (comm 1 → esec. 12,00)
{
  const { data } = await importRows([
    row('02/01/2025', ETF_ISIN, 'A', 10, 10.10, 1),
    row('03/02/2025', ETF_ISIN, 'V', 10, 10.90, 1),
    row('03/03/2025', ETF_ISIN, 'A', 5, 12.20, 1),
  ]);
  const d = data[0];
  assert.equal(d.closed, false, 'posizione riaperta: non chiusa');
  assert.equal(d.qty, 5, 'quote dopo la riapertura');
  close(d.price, 12, 'PMC = solo esecuzione del nuovo ciclo');
  close(d.realizedPLgross, 10, 'realizzato del ciclo chiuso = 10 × (11 − 10)', 0.005);
  assert.equal(new Date(d.firstDate).getMonth(), 2, 'firstDate = data della riapertura (marzo)');

  const { data: dataClosed } = await importRows([
    row('02/01/2025', ETF_ISIN, 'A', 10, 10.10, 1),
    row('03/02/2025', ETF_ISIN, 'V', 10, 10.90, 1),
  ]);
  assert.equal(dataClosed[0].closed, true, 'vendita totale → posizione chiusa');
  assert.equal(dataClosed[0].qty, 0, 'qty esattamente 0');
  console.log('✓ 4. Chiusura e riapertura: PMC 12,0000, realizzato 10,00');
}

// ── 5. BTP Futura (isNominal) con commissione 0: comportamento invariato ──
{
  const { data } = await importRows([row('03/05/2021', BTP_ISIN, 'A', 5000, 98, 0)]);
  const d = data[0];
  assert.equal(d.qty, 5, 'qty nominale ÷ 1000');
  close(d.price, 98, 'PMC BTP in % del nominale');
  close(d.cashFlows[0].amount, -4900, 'flusso BTP = −qty × prezzo × 10');
  console.log('✓ 5. BTP con commissione 0: invariato');
}

console.log('\nTutti i test commissioni Fineco superati.');
