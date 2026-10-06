// Test (jsdom): backup dei dati (storico pfImportedOps + dati inseriti a mano) ed eventuale
// ripristino con deduplica; data dell'ultimo backup sempre visibile con avviso oltre 30 giorni.
// Solo dati fittizi. Uso: node scripts/test-backup.mjs

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
const BTP = 'IT0005442089';
const ROWS = [['Intestazione fittizia'], H,
  T('02/01/2024', 'IE00B4L5Y983', 'A', 100, 10.0295, 2.95), T('02/04/2025', 'IE00B4L5Y983', 'A', 1, 12, 0), T('02/04/2025', 'IE00B4L5Y983', 'A', 1, 12, 0), // 2 righe identiche reali
  T('03/06/2024', 'IE00B4L5Y983', 'V', 60, 12.95, 2),
  T('01/07/2024', 'FR0013416716', 'A', 8, 140, 0),
  T('02/01/2024', 'LU1287023003', 'A', 20, 150, 0), T('05/05/2025', 'LU1287023003', 'V', 20, 155, 0),
  T('03/05/2021', BTP, 'A', 5000, 100, 0), T('07/03/2025', 'IT0005442097', 'V', 1000, 79.2, 10),
  C('28/10/2021', BTP, 5000, 16.41), C('28/10/2025', BTP, 4000, 21)];

function apri(storage = {}) {
  let ultimoBlob = null, ultimoNome = null;
  const w = new JSDOM(html, {
    url: 'http://localhost/', runScripts: 'dangerously', resources: undefined, pretendToBeVisual: true,
    beforeParse(w) {
      for (const [k, v] of Object.entries(storage)) w.localStorage.setItem(k, v);
      w.fetch = async () => { throw new Error('offline nei test'); };
      w.Element.prototype.scrollIntoView = () => {}; w.alert = () => {}; w.confirm = () => true;
      w.URL.createObjectURL = b => { ultimoBlob = b; return 'blob:test'; };
      w.URL.revokeObjectURL = () => {};
      w.HTMLAnchorElement.prototype.click = function () { ultimoNome = this.download; };
      w.XLSX = { read: () => ({ SheetNames: ['S'], Sheets: { S: ROWS } }), utils: { sheet_to_json: ws => ws } };
    },
  }).window;
  w.__download = async () => ({ nome: ultimoNome, testo: await new Promise(r => { const fr = new w.FileReader(); fr.onload = () => r(fr.result); fr.readAsText(ultimoBlob); }) });
  return w;
}
const storageDi = w => Object.fromEntries([...Array(w.localStorage.length)].map((_, i) => [w.localStorage.key(i), w.localStorage.getItem(w.localStorage.key(i))]));
const perIsin = w => Object.fromEntries([...w.eval('positions')].filter(p => p.isin).map(p => [p.isin, JSON.stringify(p)]));
const conteggi = w => JSON.stringify(Object.fromEntries(Object.entries(w.eval('importedOpsStore')).map(([k, v]) => [k, [v.ops.length, (v.cedole || []).length]])));
const testoStato = w => w.document.getElementById('backupStatus').textContent + ' | ' + (w.document.getElementById('backupWarn').style.display !== 'none' ? 'AVVISO: ' + w.document.getElementById('backupWarn').textContent : 'nessun avviso');

// ── Stato di partenza: import + Applica, prezzo BTP con ✎, nome personalizzato, posizione e
//    operazione manuali, target, seed BTP ──
let w = apri({
  pfPositions: JSON.stringify([{ name: 'MANUALE', ac: 'Azioni', qty: 3, price: 20, mkt: 21, closed: false }]),
  pfTargets: JSON.stringify({ Azioni: 70, Obbligazioni: 25, 'Liquidità': 5 }),
  pfBtpFuturaSeedOp: JSON.stringify({ qty: 5, price: 100, date: '2021-05-03', comm: 0 }),
});
w.handleFile(new w.File(['x'], 'fittizio.xlsx'));
await new Promise(r => setTimeout(r, 120));
w.applyImport();
const btpNome = w.eval(`positions.find(p => p.isin === '${BTP}').name`);
w.editRow(btpNome); w.document.getElementById('eMkt').value = '74.84'; w.saveEdit(btpNome);
w.editRow('SWDA'); w.document.getElementById('eN').value = 'MONDO'; w.saveEdit('SWDA');
w.eval(`operations.push({ ticker: 'MANUALE', tipo: 'A', data: '2026-05-04', qty: 1, price: 22, comm: 0, costBasis: 22, id: 1 }); save();`);
const riferimento = perIsin(w), conteggiRif = conteggi(w);
const manualeRif = JSON.stringify(w.eval(`positions.find(p => p.name === 'MANUALE')`));

// 1. Senza backup: stato visibile in intestazione + avviso
assert.match(testoStato(w), /Nessun backup.*AVVISO/s, 'senza backup: avviso visibile — ' + testoStato(w));
console.log('✓ 1. Nessun backup: data in intestazione ("Nessun backup") e avviso visibile');

// 2. Esporta: file .backup.json con storico + dati inseriti a mano; data aggiornata, avviso sparito
w.esportaBackup();
const { nome, testo } = await w.__download();
assert.match(nome, /^portfolio-backup-\d{4}-\d{2}-\d{2}\.backup\.json$/, 'nome file escluso dal .gitignore: ' + nome);
const backup = JSON.parse(testo);
assert.equal(backup.formato, 'portfolio-dashboard-backup'); assert.equal(backup.versione, 1);
for (const k of ['pfImportedOps', 'pfCustomSecurities', 'pfOperations', 'pfTargets', 'pfBtpFuturaSeedOp', 'posizioniImportate', 'posizioniManuali']) assert.ok(k in backup.dati, 'contiene ' + k);
assert.ok(backup.dati.pfOperations.every(o => !o.fromImport), 'solo operazioni manuali');
assert.ok(!/AVVISO/.test(testoStato(w)) && /Ultimo backup/.test(testoStato(w)), 'dopo il backup: data e nessun avviso — ' + testoStato(w));
console.log(`✓ 2. Esporta: ${nome} con storico e dati manuali; intestazione aggiornata, avviso sparito`);

// 3. Browser vuoto → ripristino da file: posizioni identiche al centesimo, prezzo BTP e nomi conservati
{
  const v = apri();
  await v.ripristinaBackupDaFile(new v.File([testo], nome, { type: 'application/json' }));
  const r = perIsin(v);
  for (const isin of Object.keys(riferimento)) assert.equal(r[isin], riferimento[isin], `${isin}: identico dopo il ripristino`);
  assert.equal(Object.keys(r).length, Object.keys(riferimento).length, 'stesso numero di posizioni importate');
  assert.equal(JSON.stringify(v.eval(`positions.find(p => p.name === 'MANUALE')`)), manualeRif, 'posizione manuale ripristinata');
  assert.equal(conteggi(v), conteggiRif, 'storico identico (incluse le 2 righe identiche reali)');
  assert.equal(v.eval(`FINECO_MAP['${BTP}'].isNominal`), true, 'BTP nominale dopo il ripristino');
  assert.deepEqual(JSON.parse(JSON.stringify(v.eval('targets'))), { Azioni: 70, Obbligazioni: 25, 'Liquidità': 5 }, 'target ripristinati');
  assert.ok(v.localStorage.getItem('pfBtpFuturaSeedOp'), 'seed BTP ripristinato');
  assert.match(v.document.getElementById('backupEsito').textContent, /Ripristino completato/);

  // 4. Secondo ripristino dello stesso file: nessun duplicato, nulla cambia
  const prima = perIsin(v), storicoPrima = v.localStorage.getItem('pfImportedOps');
  await v.ripristinaBackupDaFile(new v.File([testo], nome, { type: 'application/json' }));
  assert.deepEqual(perIsin(v), prima, 'posizioni invariate al secondo ripristino');
  assert.equal(v.localStorage.getItem('pfImportedOps'), storicoPrima, 'storico invariato: nessun duplicato');
  assert.equal(v.eval(`operations.filter(o => !o.fromImport).length`), 1, 'operazione manuale non duplicata');
  console.log('✓ 3. Ripristino in browser vuoto: posizioni identiche al centesimo, prezzo BTP, nomi, manuali, target, seed');
  console.log('✓ 4. Secondo ripristino: nessun duplicato');
}

// 5. File non valido → errore visibile e nessuna modifica
{
  const v = apri(storageDi(w)); const prima = JSON.stringify(storageDi(v));
  await v.ripristinaBackupDaFile(new v.File(['{ non è json'], 'rotto.json'));
  assert.match(v.document.getElementById('backupEsito').textContent, /non valido/i);
  await v.ripristinaBackupDaFile(new v.File([JSON.stringify({ formato: 'altro', versione: 1, dati: {} })], 'altro.json'));
  assert.match(v.document.getElementById('backupEsito').textContent, /non valido/i);
  const rotto = JSON.parse(testo); rotto.dati.pfImportedOps[BTP].ops[0].date = 'non è una data';
  await v.ripristinaBackupDaFile(new v.File([JSON.stringify(rotto)], 'data-rotta.json'));
  assert.match(v.document.getElementById('backupEsito').textContent, /non valido/i);
  assert.equal(JSON.stringify(storageDi(v)), prima, 'nessuna modifica ai dati salvati');
  console.log('✓ 5. File non valido (non JSON, formato errato, data non valida): errore visibile, nessuna modifica');
}

// 6. Età del backup: 31 giorni → avviso; 5 giorni → nessun avviso
{
  const g = n => new Date(Date.now() - n * 864e5).toISOString();
  let v = apri({ ...storageDi(w), pfLastBackup: g(31) });
  assert.match(testoStato(v), /AVVISO.*30 giorni/s, '31 giorni: avviso — ' + testoStato(v));
  v = apri({ ...storageDi(w), pfLastBackup: g(5) });
  assert.ok(!/AVVISO/.test(testoStato(v)), '5 giorni: nessun avviso — ' + testoStato(v));
  console.log('✓ 6. Backup di 31 giorni: avviso; di 5 giorni: nessun avviso');
}

console.log('\nTutti i test backup superati.');
