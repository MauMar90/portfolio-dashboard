// Test (jsdom): avviso nel riquadro "Qualità import" per ogni riga dell'export con descrizione
// non riconosciuta, a calcoli invariati. Riconosciute: "Compravendita titoli", "Stacco Cedole",
// "Sottoscrizione BTP Futura (collocamento)". Solo dati fittizi.
// Uso: node scripts/test-avviso-descrizioni.mjs

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
const R = (desc, d, isin, s, q, p, ctv, c) => [d, d, desc, 'TITOLO TEST', isin, s, q, 'EUR', p, 1, ctv, c];

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
  const box = win.document.getElementById('importQuality');
  // "years" dipende dall'istante dell'import (Date corrente): escluso dal confronto dei calcoli
  return { data: JSON.stringify([...win.eval('importedData')].map(({ years, ...resto }) => resto)), warnings: [...win.eval('importWarnings')], box: box.textContent, boxVisible: box.style.display !== 'none' };
}

// 1. Solo descrizioni riconosciute → nessun avviso "non riconosciuta"
{
  const r = await importa([
    R('Sottoscrizione BTP Futura (collocamento)', '03/05/2021', BTP, 'A', 5000, 100, 5000, 0),
    R('Compravendita titoli', '02/01/2025', ETF, 'A', 10, 100, 1000, 0),
    R('Stacco Cedole', '28/10/2021', BTP, '', 5000, 0, 16.41, ''),
  ]);
  assert.ok(!r.warnings.some(w => /non riconosciuta/.test(w)), 'nessun avviso con descrizioni note: ' + r.warnings.join(' | '));
  console.log('✓ 1. Descrizioni riconosciute (compravendita, cedola, sottoscrizione BTP): nessun avviso');
}

// 2. Descrizione sconosciuta → avviso visibile con descrizione, ISIN e trattamento; calcoli invariati
{
  const base = [R('Compravendita titoli', '03/05/2021', BTP, 'A', 5000, 100, 5000, 0)];
  const ignota = R('Premio fedeltà BTP Futura', '28/04/2025', BTP, 'A', 5000, 1, 50, 0);
  const conAvviso = await importa([...base, ignota]);
  const riferimento = await importa([...base, ['28/04/2025', '28/04/2025', 'Compravendita titoli', ...ignota.slice(3)]]);
  const avviso = conAvviso.warnings.find(w => /non riconosciuta/.test(w));
  assert.ok(avviso, 'avviso presente: ' + conAvviso.warnings.join(' | '));
  assert.ok(avviso.includes('Premio fedeltà BTP Futura') && avviso.includes(BTP) && /acquisto/.test(avviso), 'avviso con descrizione, ISIN e trattamento: ' + avviso);
  assert.ok(conAvviso.boxVisible && conAvviso.box.includes('Premio fedeltà BTP Futura'), 'avviso visibile nel riquadro Qualità import');
  assert.equal(conAvviso.data, riferimento.data, 'calcoli invariati: stessi risultati di una riga di compravendita equivalente');
  console.log('✓ 2. Descrizione sconosciuta: avviso visibile (descrizione, ISIN, trattata come acquisto); calcoli invariati');
}

console.log('\nTutti i test avviso descrizioni superati.');
