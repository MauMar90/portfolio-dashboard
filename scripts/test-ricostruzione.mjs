// Test di accettazione (jsdom): all'apertura le posizioni importate vengono ricostruite dallo
// storico pfImportedOps con la stessa funzione dell'import. I valori ricostruiti devono essere
// identici al centesimo a quelli di import + "Applica al portafoglio", per tutti gli strumenti,
// conservando prezzo di mercato (BTP inserito a mano) e nomi personalizzati. pfImportedOps non
// viene mai scritto; se la ricostruzione fallisce restano le posizioni salvate. Solo dati fittizi.
// Uso: node scripts/test-ricostruzione.mjs

import { JSDOM } from 'jsdom';
import fs from 'node:fs/promises';
import assert from 'node:assert/strict';
import path from 'node:path';
import { fileURLToPath } from 'node:url';

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const html = await fs.readFile(path.join(__dirname, '..', 'index.html'), 'utf8');

const H = ['Operazione', 'Data valuta', 'Descrizione', 'Titolo', 'Isin', 'Segno', 'Quantita', 'Divisa', 'Prezzo', 'Cambio', 'Controvalore', 'Commissioni amministrato'];
const T = (d, isin, s, q, p, c) => [d, d, 'Compravendita titoli', 'TITOLO TEST', isin, s, q, 'EUR', p, 1, Math.round((isin.startsWith('IT00054420') ? q * p / 100 : q * p) * 100) / 100, c];
const C = (d, isin, q, v) => [d, d, 'Stacco Cedole', 'TITOLO TEST', isin, '', q, 'EUR', 0, 1, v, ''];
const ROWS = [['Intestazione fittizia'], H,
  T('02/01/2024', 'IE00B4L5Y983', 'A', 100, 10.0295, 2.95), T('01/02/2024', 'IE00B4L5Y983', 'A', 50, 12.03, 1.5), T('03/06/2024', 'IE00B4L5Y983', 'V', 60, 12.95, 2),
  T('02/01/2024', 'IE00BTJRMP35', 'A', 10, 10.10, 1), T('01/03/2024', 'IE00BTJRMP35', 'V', 10, 10.90, 1), T('02/04/2024', 'IE00BTJRMP35', 'A', 5, 12.20, 1),
  T('02/01/2024', 'LU1287023003', 'A', 20, 150, 0), T('05/05/2025', 'LU1287023003', 'V', 20, 155, 0),
  T('01/07/2024', 'FR0013416716', 'A', 8, 140, 0), T('02/09/2024', 'FR0013416716', 'A', 2, 150, 0),
  T('03/05/2021', 'IT0005442089', 'A', 5000, 100, 0), T('07/03/2025', 'IT0005442097', 'V', 1000, 79.2, 10),
  C('28/10/2021', 'IT0005442089', 5000, 16.41), C('28/04/2022', 'IT0005442089', 5000, 16.41), C('28/10/2025', 'IT0005442089', 4000, 21)];
const BTP = 'IT0005442089';

function apri(storage = {}) {
  return new JSDOM(html, {
    url: 'http://localhost/', runScripts: 'dangerously', resources: undefined, pretendToBeVisual: true,
    beforeParse(w) {
      for (const [k, v] of Object.entries(storage)) w.localStorage.setItem(k, v);
      w.fetch = async () => { throw new Error('offline nei test'); };
      w.Element.prototype.scrollIntoView = () => {}; w.alert = () => {}; w.confirm = () => true;
      w.XLSX = { read: () => ({ SheetNames: ['S'], Sheets: { S: ROWS } }), utils: { sheet_to_json: ws => ws } };
    },
  }).window;
}
const storageDi = w => Object.fromEntries([...Array(w.localStorage.length)].map((_, i) => [w.localStorage.key(i), w.localStorage.getItem(w.localStorage.key(i))]));
const perIsin = w => Object.fromEntries([...w.eval('positions')].filter(p => p.isin).map(p => [p.isin, JSON.stringify(p)]));

// ── Stato di riferimento: import + "Applica al portafoglio", poi ✎ (prezzo BTP, ETF rinominato) ──
let w = apri({ pfPositions: JSON.stringify([{ name: 'MANUALE', ac: 'Azioni', qty: 3, price: 20, mkt: 21, closed: false }]) });
w.handleFile(new w.File(['x'], 'fittizio.xlsx'));
await new Promise(r => setTimeout(r, 120));
w.applyImport();
const btpName = w.eval(`positions.find(p => p.isin === '${BTP}').name`);
w.editRow(btpName); w.document.getElementById('eMkt').value = '74.84'; w.saveEdit(btpName);
w.editRow('SWDA'); w.document.getElementById('eN').value = 'MONDO'; w.document.getElementById('eAC').value = 'Azioni'; w.saveEdit('SWDA');
const riferimento = perIsin(w);
const manualeRif = JSON.stringify(w.eval(`positions.find(p => p.name === 'MANUALE')`));
const storage = storageDi(w);
assert.equal(Object.keys(riferimento).length, 5, '5 strumenti importati');

// ── 1. Valori salvati "vecchi" (come calcolati da una versione precedente) → ricostruiti all'apertura ──
{
  const pos = JSON.parse(storage.pfPositions);
  for (const p of pos) if (p.isin) { p.price += 1; p.realizedPLgross = 32.2; p.realizedPLnet = -2.59; p.cashFlows = [{ amount: -1, date: new Date().toISOString() }]; p.commTot = 0; }
  const opsPrima = storage.pfImportedOps;
  w = apri({ ...storage, pfPositions: JSON.stringify(pos) });
  const ricostruito = perIsin(w);
  for (const isin of Object.keys(riferimento)) assert.equal(ricostruito[isin], riferimento[isin], `${isin}: ricostruito identico a import + Applica`);
  assert.equal(w.eval(`positions.find(p => p.isin === '${BTP}').mkt`), 74.84, 'prezzo BTP inserito a mano conservato');
  assert.ok(w.eval(`positions.some(p => p.name === 'MONDO' && p.isin === 'IE00B4L5Y983')`), 'nome personalizzato conservato');
  assert.equal(JSON.stringify(w.eval(`positions.find(p => p.name === 'MANUALE')`)), manualeRif, 'posizione manuale (senza ISIN) invariata');
  assert.equal(w.localStorage.getItem('pfImportedOps'), opsPrima, 'pfImportedOps mai scritto');
  console.log('✓ 1. Valori salvati vecchi → all\'apertura ricostruiti identici a import + Applica (5 strumenti); prezzo BTP, nome e posizione manuale conservati; storico non scritto');
}

// ── 2. Riapertura senza modifiche: stabile (stessi valori a ogni apertura) ──
{
  w = apri(storage);
  const a = perIsin(w);
  w = apri(storageDi(w));
  assert.deepEqual(perIsin(w), a, 'due aperture consecutive danno gli stessi valori');
  for (const isin of Object.keys(riferimento)) assert.equal(a[isin], riferimento[isin]);
  console.log('✓ 2. Aperture successive: valori stabili e identici al riferimento');
}

// ── 3. Ricostruzione che fallisce (storico con data non valida) → restano le posizioni salvate ──
{
  const opsRotti = JSON.parse(storage.pfImportedOps);
  opsRotti[BTP].ops[0].date = 'non è una data';
  const salvate = JSON.parse(storage.pfPositions);
  salvate.find(p => p.isin === BTP).realizedPLgross = 99.99;   // valore salvato riconoscibile
  w = apri({ ...storage, pfImportedOps: JSON.stringify(opsRotti), pfPositions: JSON.stringify(salvate) });
  assert.equal(w.eval(`positions.find(p => p.isin === '${BTP}').realizedPLgross`), 99.99, 'errore di ricostruzione: resta la posizione salvata');
  assert.equal(w.eval('positions.length'), salvate.length, 'nessuna posizione persa');
  console.log('✓ 3. Ricostruzione non riuscita: restano le posizioni salvate');
}

// ── 4. Una posizione rimossa dal portafoglio non ricompare all'apertura ──
{
  const senzaGold = JSON.parse(storage.pfPositions).filter(p => p.isin !== 'FR0013416716');
  w = apri({ ...storage, pfPositions: JSON.stringify(senzaGold) });
  assert.ok(!w.eval(`positions.some(p => p.isin === 'FR0013416716')`), 'posizione rimossa non resuscitata dallo storico');
  console.log('✓ 4. Posizione rimossa: non ricompare');
}

console.log('\nTutti i test di ricostruzione superati.');
