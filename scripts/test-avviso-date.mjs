// Test (jsdom): avviso per date non plausibili nelle righe dell'export (compravendite e cedole),
// es. una data rimasta come numero seriale di Excel ("44313" → anno 44313). Calcoli invariati.
// Solo dati fittizi. Uso: node scripts/test-avviso-date.mjs

import { JSDOM } from 'jsdom';
import fs from 'node:fs/promises';
import assert from 'node:assert/strict';
import path from 'node:path';
import { fileURLToPath } from 'node:url';

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const html = await fs.readFile(path.join(__dirname, '..', 'index.html'), 'utf8');

const H = ['Operazione', 'Data valuta', 'Descrizione', 'Titolo', 'Isin', 'Segno', 'Quantita', 'Divisa', 'Prezzo', 'Cambio', 'Controvalore', 'Commissioni amministrato'];
const ETF = 'IE00B4L5Y983', BTP = 'IT0005442089';
const T = (d, isin = ETF) => [d, d, 'Compravendita titoli', 'TITOLO TEST', isin, 'A', 10, 'EUR', 100, 1, 1000, 0];
const C = d => [d, d, 'Stacco Cedole', 'TITOLO TEST', BTP, '', 5000, 'EUR', 0, 1, 16.41, ''];

async function importa(rows) {
  const w = new JSDOM(html, {
    url: 'http://localhost/', runScripts: 'dangerously', resources: undefined, pretendToBeVisual: true,
    beforeParse(w) {
      w.fetch = async () => { throw new Error('offline nei test'); };
      w.XLSX = { read: () => ({ SheetNames: ['S'], Sheets: { S: [['Intestazione fittizia'], H, ...rows] } }), utils: { sheet_to_json: ws => ws } };
    },
  }).window;
  w.handleFile(new w.File(['x'], 'fittizio.xlsx'));
  await new Promise(r => setTimeout(r, 100));
  return { avvisi: [...w.eval('importWarnings')], dati: [...w.eval('importedData')].map(({ years, ...x }) => x) };
}
const avvisiData = a => a.filter(x => /data non valida/i.test(x));

// 1. Date valide (anche vecchie, es. 2021) → nessun avviso
{
  const { avvisi } = await importa([T('27/04/2021'), T('02/01/2025'), C('28/10/2025')]);
  assert.equal(avvisiData(avvisi).length, 0, 'nessun avviso su date valide: ' + avvisi.join(' | '));
  console.log('✓ 1. Date valide: nessun avviso');
}

// 2. Numero seriale Excel su compravendita e su cedola → avviso con il testo originale; calcoli invariati
{
  const r = await importa([T('44313'), C('45658')]);
  const a = avvisiData(r.avvisi);
  assert.equal(a.length, 2, 'due righe con data non valida segnalate: ' + r.avvisi.join(' | '));
  assert.ok(a.some(x => x.includes(ETF) && x.includes('"44313"')) && a.some(x => x.includes(BTP) && x.includes('"45658"')), 'avviso con ISIN e testo originale della data');
  assert.equal(r.dati.length, 2, 'le righe vengono comunque importate come prima (calcoli invariati)');
  console.log('✓ 2. Data come numero seriale Excel (compravendita e cedola): avviso con ISIN e valore originale; calcoli invariati');
}

// 3. Data futura oltre l'anno prossimo → avviso
{
  const anno = new Date().getFullYear() + 5;
  const { avvisi } = await importa([T(`01/01/${anno}`)]);
  assert.equal(avvisiData(avvisi).length, 1, 'data troppo nel futuro segnalata');
  console.log('✓ 3. Data troppo nel futuro: avviso');
}

console.log('\nTutti i test avviso date superati.');
