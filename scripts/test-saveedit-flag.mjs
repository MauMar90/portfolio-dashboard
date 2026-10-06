// Test (jsdom): saveEdit non deve cancellare gli attributi strutturali del titolo (isNominal,
// sellOf) e le voci di customSecurities già rovinate vanno riparate all'apertura.
// Caso reale riprodotto: customSecurities['IT0005442089'] = {ticker, ac, name} senza isNominal.
// Solo dati fittizi. Uso: node scripts/test-saveedit-flag.mjs

import { JSDOM } from 'jsdom';
import fs from 'node:fs/promises';
import assert from 'node:assert/strict';
import path from 'node:path';
import { fileURLToPath } from 'node:url';

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const html = await fs.readFile(path.join(__dirname, '..', 'index.html'), 'utf8');

const HEADER = ['Operazione', 'Data valuta', 'Descrizione', 'Titolo', 'Isin', 'Segno', 'Quantita',
  'Divisa', 'Prezzo', 'Cambio', 'Controvalore', 'Commissioni amministrato'];
const BTP = 'IT0005442089', BTP_V = 'IT0005442097';
const ROWS = [['Intestazione fittizia'], HEADER,
  ['03/05/2021', '03/05/2021', 'Compravendita titoli', 'BTP TEST', BTP, 'A', 5000, 'EUR', 100, 1, 5000, 0],
  ['07/03/2025', '11/03/2025', 'Compravendita titoli', 'BTP TEST', BTP_V, 'V', 1000, 'EUR', 79.2, 1, 792, 10]];

function apri(storage = {}) {
  return new JSDOM(html, {
    url: 'http://localhost/', runScripts: 'dangerously', resources: undefined, pretendToBeVisual: true,
    beforeParse(win) {
      for (const [k, v] of Object.entries(storage)) win.localStorage.setItem(k, v);
      win.fetch = async () => { throw new Error('offline nei test'); };
      win.Element.prototype.scrollIntoView = () => {}; win.alert = () => {};
      win.XLSX = { read: () => ({ SheetNames: ['S'], Sheets: { S: ROWS } }), utils: { sheet_to_json: ws => ws } };
    },
  }).window;
}
async function importaEApplica(win) {
  win.handleFile(new win.File(['x'], 'fittizio.xlsx'));
  await new Promise(r => setTimeout(r, 100));
  win.applyImport();
  return win.eval(`positions.find(p => p.isin === '${BTP}')`);
}
const storageDi = win => Object.fromEntries([...Array(win.localStorage.length)].map((_, i) => [win.localStorage.key(i), win.localStorage.getItem(win.localStorage.key(i))]));

// 1. Caso reale: voce BTP in customSecurities senza isNominal (nome e asset class personalizzati)
{
  const win = apri({ pfCustomSecurities: JSON.stringify({
    [BTP]:   { ticker: 'BTP Futura', ac: 'Obbligazioni Italia', name: 'BTP Futura 2037' },
    [BTP_V]: { ticker: 'BTP Futura', ac: 'Obbligazioni', name: 'BTP Futura' },        // anche sellOf perso
  }) });
  const m = win.eval(`FINECO_MAP['${BTP}']`), mv = win.eval(`FINECO_MAP['${BTP_V}']`);
  assert.equal(m.isNominal, true, 'isNominal ripristinato all\'apertura');
  assert.equal(m.name, 'BTP Futura 2037', 'nome personalizzato conservato');
  assert.equal(m.ac, 'Obbligazioni Italia', 'asset class personalizzata conservata');
  assert.equal(mv.sellOf, BTP, 'sellOf della linea di vendita ripristinato');
  const p = await importaEApplica(win);
  assert.equal(p.qty, 4, 'import dopo la riparazione: 4 quote (5.000 − 1.000 nominali), non 5.000+');
  assert.equal(p.isNominal, true);
  console.log('✓ 1. Voce rovinata in customSecurities: isNominal/sellOf ripristinati, nome e asset class conservati, import corretto');
}

// 2. Modifica con ✎ (prezzo di mercato) + riapertura + reimport: il BTP resta nominale
{
  let win = apri();
  let p = await importaEApplica(win);
  assert.equal(p.qty, 4);
  win.editRow(p.name); win.document.getElementById('eMkt').value = '74.84'; win.saveEdit(p.name);
  assert.equal(win.eval(`FINECO_MAP['${BTP}'].isNominal`), true, 'saveEdit conserva isNominal');
  assert.equal(JSON.parse(win.localStorage.getItem('pfCustomSecurities'))[BTP].isNominal, true, 'anche la voce salvata conserva isNominal');
  win = apri(storageDi(win));
  p = await importaEApplica(win);
  assert.equal(p.qty, 4, 'reimport dopo ✎ e riapertura: 4 quote (prima del fix: 5.004)');
  assert.equal(p.mkt, 74.84, 'prezzo di mercato inserito a mano conservato');
  console.log('✓ 2. ✎ sul BTP + riapertura + reimport: quote 4, prezzo manuale 74,84 conservato');
}

console.log('\nTutti i test saveEdit / attributi titolo superati.');
