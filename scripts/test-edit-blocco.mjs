// Test (jsdom): in modifica (✎) le posizioni importate hanno quote e PMC bloccati (vengono
// dall'import); restano modificabili prezzo di mercato, nome e asset class. Le posizioni
// manuali (senza ISIN) restano interamente modificabili. Solo dati fittizi.
// Uso: node scripts/test-edit-blocco.mjs

import { JSDOM } from 'jsdom';
import fs from 'node:fs/promises';
import assert from 'node:assert/strict';
import path from 'node:path';
import { fileURLToPath } from 'node:url';

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const html = await fs.readFile(path.join(__dirname, '..', 'index.html'), 'utf8');

const win = new JSDOM(html, {
  url: 'http://localhost/', runScripts: 'dangerously', resources: undefined, pretendToBeVisual: true,
  beforeParse(w) {
    w.localStorage.setItem('pfPositions', JSON.stringify([
      { name: 'ETF IMPORTATO', isin: 'XX0000000001', ticker: 'ETF IMPORTATO', ac: 'Azioni', qty: 10, price: 50, mkt: 55, closed: false },
      { name: 'MANUALE', ac: 'Azioni', qty: 3, price: 20, mkt: 21, closed: false },  // senza ISIN
    ]));
    w.fetch = async () => { throw new Error('offline nei test'); }; w.alert = () => {};
  },
}).window;
const $ = id => win.document.getElementById(id);
const pos = name => win.eval(`positions.find(p => p.name === ${JSON.stringify(name)})`);

// 1. Posizione importata: campi quote/PMC disabilitati; anche forzandoli, il salvataggio li ignora
win.editRow('ETF IMPORTATO');
assert.equal($('eQty').disabled, true, 'quote disabilitate per posizione importata');
assert.equal($('ePrice').disabled, true, 'PMC disabilitato per posizione importata');
assert.equal($('eMkt').disabled, false, 'prezzo di mercato modificabile');
$('eQty').value = '999'; $('ePrice').value = '1'; $('eMkt').value = '60'; $('eN').value = 'ETF RINOMINATO'; $('eAC').value = 'Obbligazioni';
win.saveEdit('ETF IMPORTATO');
let p = pos('ETF RINOMINATO');
assert.ok(p, 'nome modificato');
assert.equal(p.qty, 10, 'quote invariate'); assert.equal(p.price, 50, 'PMC invariato');
assert.equal(p.mkt, 60, 'prezzo di mercato aggiornato'); assert.equal(p.ac, 'Obbligazioni', 'asset class aggiornata');
console.log('✓ 1. Importata: quote e PMC bloccati (anche forzando i campi); prezzo, nome e asset class modificati');

// 2. Posizione manuale (senza ISIN): tutto modificabile come prima
win.editRow('MANUALE');
assert.equal($('eQty').disabled, false); assert.equal($('ePrice').disabled, false);
$('eQty').value = '4'; $('ePrice').value = '22';
win.saveEdit('MANUALE');
p = pos('MANUALE');
assert.equal(p.qty, 4, 'quote manuali modificate'); assert.equal(p.price, 22, 'PMC manuale modificato');
console.log('✓ 2. Manuale (senza ISIN): quote e PMC restano modificabili');

console.log('\nTutti i test blocco modifica superati.');
