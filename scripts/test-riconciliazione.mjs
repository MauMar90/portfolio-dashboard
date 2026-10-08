// Test (jsdom): riconciliazione con l'export del portafoglio Fineco ("Portafoglio sintesi").
// L'export portafoglio è riconosciuto all'import e usato SOLO per il controllo: non scrive mai in
// pfImportedOps, non tocca posizioni né flussi XIRR. La posizione della dashboard è ricalcolata
// alla data della fotografia (giorno di calendario Europe/Rome). Solo dati fittizi.
// Uso: node scripts/test-riconciliazione.mjs

process.env.TZ = 'Europe/Rome';   // i giorni di calendario vanno verificati come in Italia

import { JSDOM } from 'jsdom';
import fs from 'node:fs/promises';
import assert from 'node:assert/strict';
import path from 'node:path';
import { fileURLToPath } from 'node:url';

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const html = await fs.readFile(path.join(__dirname, '..', 'index.html'), 'utf8');

const ETF_A = 'IE00B4L5Y983', ETF_B = 'IE00BTJRMP35', BTP = 'IT0005442089', ORO = 'FR0013416716';
const HM = ['Operazione', 'Data valuta', 'Descrizione', 'Titolo', 'Isin', 'Segno', 'Quantita', 'Divisa', 'Prezzo', 'Cambio', 'Controvalore', 'Commissioni amministrato'];
const T = (d, isin, s, q, p, c = 0) => [d, d, 'Compravendita titoli', 'TITOLO TEST', isin, s, q, 'EUR', p, 1, Math.round((isin.startsWith('IT00054420') ? q * p / 100 : q * p) * 100) / 100, c];
const MOVIMENTI = [['RISULTATO RICERCA MOVIMENTI TITOLI'], HM,
  T('03/05/2021', BTP, 'A', 5000, 100),
  T('01/06/2026', ETF_A, 'A', 10, 100.5, 5),        // PMC 100,00 (commissione inclusa nel Prezzo export)
  T('01/06/2026', ETF_B, 'A', 20, 50),
  T('25/06/2026', ETF_B, 'A', 5, 52),               // stesso giorno della fotografia → incluso
  T('26/06/2026', ETF_A, 'A', 3, 110),              // giorno successivo → escluso
];
const HP = ['Titolo', 'ISIN', 'Simbolo', 'Mercato', 'Strumento', 'Valuta', 'Quantità', 'P.zo medio di carico', 'Cambio di carico', 'Valore di carico', 'P.zo di mercato', 'Cambio di mercato', 'Valore di mercato €', 'Var%', 'Var €', 'Var in valuta', 'Rateo'];
const P = (isin, q, pmc, str = 'ETF') => ['TITOLO TEST', isin, 'SIMB', 'AFF', str, 'EUR', q, pmc, 1, q * pmc, pmc, 1, q * pmc, 0, 0, 0, 0];
const portafoglio = righe => [['Portafoglio di sintesi'], [], HP, [], ...righe, ['Totale', '', '', '', '', '', '', '', '', '', '', 'Valore di carico', 'Valore di mercato', 'Var%', 'Var']];
// fotografia "vera" al 25/06/2026: ETF_A 10 @100, ETF_B 25 @ (20×50 + 5×52)/25 = 50,4, BTP 5.000 nominali @100
const FOTO_OK = [P(ETF_A, 10, 100), P(ETF_B, 25, 50.4), P(BTP, 5000, 100, 'Obbligazione')];
const FOTOGRAFIA = '2026-06-25', MS_FOTO = new Date(2026, 5, 25, 23, 8).getTime();

let righeCorrenti = MOVIMENTI;
async function apri(storage = {}) {
  const w = new JSDOM(html, {
    url: 'http://localhost/', runScripts: 'dangerously', resources: undefined, pretendToBeVisual: true,
    beforeParse(w) {
      for (const [k, v] of Object.entries(storage)) w.localStorage.setItem(k, v);
      w.fetch = async () => { throw new Error('offline nei test'); };
      w.Element.prototype.scrollIntoView = () => {}; w.__alert = null; w.alert = m => { w.__alert = m; };
      w.XLSX = { read: () => ({ SheetNames: ['S'], Sheets: { S: righeCorrenti } }), utils: { sheet_to_json: ws => ws } };
    },
  }).window;
  return w;
}
async function carica(w, righe, nome = 'f.xlsx', lastModified = Date.now()) {
  righeCorrenti = righe;
  w.handleFile(new w.File(['x'], nome, { lastModified }));
  await new Promise(r => setTimeout(r, 120));
}
const stato = w => JSON.stringify({ ops: w.localStorage.getItem('pfImportedOps'), mem: Object.keys(w.eval('importedOpsStore')).sort(), pos: w.eval('positions') });

// Stato di partenza: movimenti importati e applicati + un'operazione manuale
const w = await apri();
await carica(w, MOVIMENTI); w.applyImport();
w.eval(`operations.push({ ticker: 'MANUALE', tipo: 'A', data: '2026-06-10', qty: 1, price: 20, comm: 0, costBasis: 20, id: 1 }); save();`);
const statoPrima = stato(w);
const flussiPrima = JSON.stringify(w.eval('positions').map(p => p.cashFlows));

// 1. Riconoscimento: l'export portafoglio non passa dall'import movimenti e non lascia traccia
await carica(w, portafoglio(FOTO_OK), 'portafoglio-export.xls', MS_FOTO);
assert.equal(w.__alert, null, 'nessun errore "Invalid time value": ' + w.__alert);
assert.equal(stato(w), statoPrima, 'pfImportedOps, storico in memoria (nessuna voce vuota) e posizioni invariati');
assert.equal(JSON.stringify(w.eval('positions').map(p => p.cashFlows)), flussiPrima, 'flussi XIRR invariati');
assert.equal(w.document.getElementById('importResult').style.display, 'none', 'nessuna anteprima di import movimenti');
assert.equal(w.document.getElementById('ricData').value, FOTOGRAFIA, 'data proposta = data del file (giorno Europe/Rome)');
console.log('✓ 1. Export portafoglio riconosciuto: nessun errore, nessuna traccia in storico/posizioni/XIRR; data proposta dal file');

// 2. Mai scambiati: movimenti ≠ portafoglio
assert.equal(w.eExportPortafoglio(MOVIMENTI), false, 'export movimenti non è un export portafoglio');
assert.equal(w.eExportPortafoglio(portafoglio(FOTO_OK)), true, 'export portafoglio riconosciuto');
console.log('✓ 2. Export movimenti e export portafoglio non vengono mai scambiati');

// 3. Caso coincidente (con operazione nello stesso giorno incluso e giorno dopo escluso)
w.eseguiRiconciliazione();
let r = w.eval('JSON.parse(localStorage.getItem("pfRiconciliazione"))');
assert.equal(r.coincidenti, 3); assert.equal(r.totale, 3); assert.equal(r.differenze.length, 0);
const sintesi = () => w.document.getElementById('ricSintesi').textContent;
assert.match(sintesi(), /Riconciliazione al 25\/06\/2026: 3 su 3 coincidono/);
assert.equal(w.document.querySelector('#ricBox details'), null, 'nessun dettaglio espandibile se tutto coincide');
assert.equal(stato(w), statoPrima, 'la riconciliazione non modifica nulla');
console.log('✓ 3. Coincidente: "3 su 3", operazione del 25/06 inclusa e del 26/06 esclusa, nessun dettaglio');

// 4. Giorno di calendario Europe/Rome: 26/06 a mezzanotte (UTC 25/06 22:00) escluso, 25/06 incluso
const conf = (righe, giorno) => JSON.parse(JSON.stringify(w.confrontaConFineco(w.leggiExportPortafoglio(portafoglio(righe)), giorno)));
assert.equal(conf(FOTO_OK, '2026-06-25').coincidenti, 3, 'al 25/06 coincide');
assert.equal(conf(FOTO_OK, '2026-06-24').differenze.find(d => d.isin === ETF_B).tipo, 'quantita', 'al 24/06 manca l\'acquisto del 25/06');
console.log('✓ 4. Giorno di calendario Europe/Rome (non UTC): 25/06 incluso, 26/06 escluso');

// 5. Quantità diversa; PMC oltre e entro la tolleranza; ISIN da una sola parte (entrambi i versi)
let c = conf([P(ETF_A, 11, 100), P(ETF_B, 25, 50.4), P(BTP, 5000, 100, 'Obbligazione')], FOTOGRAFIA);
assert.deepEqual(c.differenze.map(d => [d.isin, d.tipo]), [[ETF_A, 'quantita']]);
c = conf([P(ETF_A, 10, 100.001), P(ETF_B, 25, 50.40004), P(BTP, 5000, 100, 'Obbligazione')], FOTOGRAFIA);
assert.deepEqual(c.differenze.map(d => [d.isin, d.tipo]), [[ETF_A, 'pmc']], 'PMC oltre 0,0001 segnalato; 0,00004 entro la tolleranza');
c = conf([P(ETF_A, 10, 100), P(BTP, 5000, 100, 'Obbligazione'), P(ORO, 8, 140, 'ETC')], FOTOGRAFIA);
assert.deepEqual(c.differenze.map(d => [d.isin, d.tipo]).sort(), [[ETF_B, 'solo_dashboard'], [ORO, 'solo_fineco']].sort());
console.log('✓ 5. Quantità diversa, PMC oltre tolleranza (0,00004 entro), ISIN solo in dashboard / solo in Fineco');

// 6. Differenze → riga di sintesi espandibile + ricerca a ritroso della data (senza cambiarla)
//    Fotografia vera del 01/06 caricata indicando il 25/06: coincide al 01/06
await carica(w, portafoglio([P(ETF_A, 10, 100), P(ETF_B, 20, 50), P(BTP, 5000, 100, 'Obbligazione')]), 'p.xls', MS_FOTO);
w.eseguiRiconciliazione();
r = w.eval('JSON.parse(localStorage.getItem("pfRiconciliazione"))');
assert.equal(r.giorno, FOTOGRAFIA, 'la data indicata non viene cambiata');
assert.equal(r.giornoProbabile, '2026-06-01', 'ricerca a ritroso: coincide al 01/06');
assert.match(sintesi(), /2 su 3 coincidono/);
const det = w.document.querySelector('#ricBox details');
assert.ok(det, 'dettaglio espandibile presente se ci sono differenze');
assert.match(det.textContent, /coincide al 01\/06\/2026: probabilmente è la data della fotografia/);
assert.match(det.textContent, /1 operazione manuale.*non inclusa/i, 'operazioni manuali elencate come nota');
assert.equal(w.document.getElementById('ricData').value, FOTOGRAFIA, 'campo data non modificato');
console.log('✓ 6. Differenze: sintesi espandibile; a ritroso "coincide al 01/06/2026" senza cambiare la data; nota operazioni manuali');

// 7. Esito salvato a parte (sopravvive alla riapertura) e NON incluso nel backup
const w2 = await apri(Object.fromEntries([...Array(w.localStorage.length)].map((_, i) => [w.localStorage.key(i), w.localStorage.getItem(w.localStorage.key(i))])));
assert.match(w2.document.getElementById('ricSintesi').textContent, /Riconciliazione al 25\/06\/2026: 2 su 3/);
let blob = null; w2.URL.createObjectURL = b => { blob = b; return 'blob:x'; }; w2.URL.revokeObjectURL = () => {}; w2.HTMLAnchorElement.prototype.click = () => {};
w2.esportaBackup();
const testo = await new Promise(res => { const fr = new w2.FileReader(); fr.onload = () => res(fr.result); fr.readAsText(blob); });
assert.ok(!/pfRiconciliazione|giornoProbabile/.test(testo), 'esito della riconciliazione non incluso nel backup');
assert.equal(stato(w), statoPrima, 'nessuna modifica a storico, posizioni o flussi in tutto il test');
console.log('✓ 7. Esito visibile alla riapertura, non incluso nel backup; nessuna modifica a storico/posizioni/XIRR');

console.log('\nTutti i test riconciliazione superati.');
