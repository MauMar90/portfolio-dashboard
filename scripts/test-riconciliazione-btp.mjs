// Test (jsdom): riconciliazione BTP con l'export portafoglio Fineco. Fineco espone il nominale e il
// carico "secco" (senza commissione d'acquisto); la dashboard conta quote da 1.000 nominali e
// include la commissione d'acquisto nel costo. Solo dati fittizi.
// Uso: node scripts/test-riconciliazione-btp.mjs

process.env.TZ = 'Europe/Rome';

import { JSDOM } from 'jsdom';
import fs from 'node:fs/promises';
import assert from 'node:assert/strict';
import path from 'node:path';
import { fileURLToPath } from 'node:url';

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const html = await fs.readFile(path.join(__dirname, '..', 'index.html'), 'utf8');

const BTP = 'IT0005442089';
const HM = ['Operazione', 'Data valuta', 'Descrizione', 'Titolo', 'Isin', 'Segno', 'Quantita', 'Divisa', 'Prezzo', 'Cambio', 'Controvalore', 'Commissioni amministrato'];
// BTP: Prezzo export = Totale eseguito / nominale × 100 (commissione d'acquisto inclusa)
const T = (d, s, nom, prezzo, comm = 0) => [d, d, 'Compravendita titoli', 'BTP TEST', BTP, s, nom, 'EUR', prezzo, 1, Math.round(nom * prezzo / 100 * 100) / 100, comm];
const HP = ['Titolo', 'ISIN', 'Simbolo', 'Mercato', 'Strumento', 'Valuta', 'Quantità', 'P.zo medio di carico', 'Cambio di carico', 'Valore di carico', 'P.zo di mercato', 'Cambio di mercato', 'Valore di mercato €', 'Var%', 'Var €', 'Var in valuta', 'Rateo'];
const P = (nom, pmc) => ['BTP TEST', BTP, 'SIMB', 'MOT', 'Obbligazione', 'EUR', nom, pmc, 1, nom * pmc / 100, pmc, 1, nom * pmc / 100, 0, 0, 0, 0];
const portafoglio = righe => [['Portafoglio di sintesi'], [], HP, [], ...righe];

async function conStorico(movimenti) {
  let cur = [['RISULTATO RICERCA MOVIMENTI TITOLI'], HM, ...movimenti];
  const w = new JSDOM(html, {
    url: 'http://localhost/', runScripts: 'dangerously', resources: undefined, pretendToBeVisual: true,
    beforeParse(w) {
      w.fetch = async () => { throw new Error('offline nei test'); };
      w.Element.prototype.scrollIntoView = () => {}; w.alert = () => {};
      w.XLSX = { read: () => ({ SheetNames: ['S'], Sheets: { S: cur } }), utils: { sheet_to_json: ws => ws } };
    },
  }).window;
  w.handleFile(new w.File(['x'], 'f.xlsx')); await new Promise(r => setTimeout(r, 120));
  w.applyImport();
  return w;
}
const conf = (w, righe, giorno = '2026-06-25') => JSON.parse(JSON.stringify(w.confrontaConFineco(w.leggiExportPortafoglio(portafoglio(righe)), giorno)));

// 1. Sottoscrizione 6.000 nominali a 100, Fineco 6.000 @ 100 → coincide (6 quote, PMC 100)
{
  const w = await conStorico([T('03/05/2021', 'A', 6000, 100)]);
  const d = w.posizioniAlGiorno('2026-06-25')[BTP];
  assert.equal(d.qty, 6, 'dashboard: 6 quote da 1.000 nominali');
  assert.equal(d.price, 100, 'dashboard: PMC 100');
  const c = conf(w, [P(6000, 100)]);
  assert.equal(c.coincidenti, 1); assert.equal(c.totale, 1); assert.equal(c.differenze.length, 0);
  console.log('✓ 1. BTP 6.000 nominali @100 nello storico e in Fineco: coincide (6 quote, PMC 100)');

  // 1b. Fineco con nominale diverso → differenza di quantità
  const k = conf(w, [P(5000, 100)]);
  assert.deepEqual(k.differenze.map(x => [x.isin, x.tipo, x.qDash, x.qFin]), [[BTP, 'quantita', 6, 5]]);
  console.log('✓ 1b. Nominale Fineco 5.000 contro 6.000: differenza di quantità segnalata (6 vs 5 quote)');
}

// 2. Differenza dovuta SOLO alla commissione d'acquisto → spiegazione esplicita
{
  // 6.000 nominali a 100 secco + 6 € di commissione: Prezzo export = 6.006 / 6.000 × 100 = 100,1
  const w = await conStorico([T('03/05/2021', 'A', 6000, 100.1, 6)]);
  const c = conf(w, [P(6000, 100)]);
  assert.equal(c.differenze.length, 1);
  const d = c.differenze[0];
  assert.equal(d.tipo, 'pmc');
  assert.equal(d.soloCommissione, true, 'riconosciuta come sola commissione d\'acquisto');
  assert.ok(Math.abs(d.commissione - 6) < 0.005, 'commissione ricostruita 6 €: ' + d.commissione);
  w.localStorage.setItem('pfRiconciliazione', JSON.stringify({ ...c, controllo: new Date().toISOString() }));
  w.renderRiconciliazione();
  const testo = w.document.querySelector('#ricBox details').textContent;
  assert.match(testo, /differenza dovuta solo alla commissione d'acquisto \(6,00 €\): la dashboard la include nel costo, Fineco espone il carico secco/);
  console.log('✓ 2. BTP con commissione d\'acquisto 6 €: "differenza dovuta solo alla commissione d\'acquisto (6,00 €)"');

  // 2b. Differenza NON spiegata dalla commissione → nessuna spiegazione sulla commissione
  const k = conf(w, [P(6000, 99)]);
  assert.equal(k.differenze[0].tipo, 'pmc');
  assert.ok(!k.differenze[0].soloCommissione, 'scarto non spiegato dalla commissione');
  w.localStorage.setItem('pfRiconciliazione', JSON.stringify(k)); w.renderRiconciliazione();
  assert.doesNotMatch(w.document.querySelector('#ricBox details').textContent, /commissione/i, 'nessuna frase sulla commissione se non è la causa');
  console.log('✓ 2b. PMC Fineco 99 (non spiegato dalla commissione): differenza senza frase sulla commissione');
}

console.log('\nTutti i test riconciliazione BTP superati.');
