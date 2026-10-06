// Test (jsdom): controlli di qualità sull'export movimenti Fineco. Solo avvisi nel riquadro
// "Qualità import", calcoli invariati. Solo dati fittizi.
// - commissione negativa, o cella commissione non vuota non leggibile come numero;
// - Controvalore ≠ quote × Prezzo (BTP: nominale × Prezzo / 100) oltre 0,01 € (righe cedola escluse);
// - divisa diversa da EUR.
// Uso: node scripts/test-controlli-import.mjs

import { JSDOM } from 'jsdom';
import fs from 'node:fs/promises';
import assert from 'node:assert/strict';
import path from 'node:path';
import { fileURLToPath } from 'node:url';

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const html = await fs.readFile(path.join(__dirname, '..', 'index.html'), 'utf8');

const H = ['Operazione', 'Data valuta', 'Descrizione', 'Titolo', 'Isin', 'Segno', 'Quantita', 'Divisa', 'Prezzo', 'Cambio', 'Controvalore', 'Commissioni amministrato'];
const ETF = 'IE00B4L5Y983', BTP = 'IT0005442089';
// riga: [data, data, desc, titolo, isin, segno, quantità, divisa, prezzo, cambio, controvalore, commissione]
const R = ({ d = '02/01/2025', desc = 'Compravendita titoli', isin = ETF, s = 'A', q = 10, div = 'EUR', p = 100, ctv, c = 0 } = {}) =>
  [d, d, desc, 'TITOLO TEST', isin, s, q, div, p, 1, ctv ?? (isin.startsWith('IT00054420') ? q * p / 100 : q * p), c];

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
  return { avvisi: [...w.eval('importWarnings')], dati: [...w.eval('importedData')] };
}
const ha = (avvisi, re) => avvisi.some(a => re.test(a));

// 1. Righe corrette (ETF, BTP, cedola, commissione vuota o con virgola decimale) → nessun avviso di qualità
{
  const { avvisi } = await importa([
    R({ c: '' }), R({ d: '03/02/2025', q: 3, p: 33.33333, ctv: 100, c: '1,74' }),          // arrotondamento 0,00001 €
    R({ isin: BTP, q: 5000, p: 100 }), R({ d: '28/10/2025', desc: 'Stacco Cedole', isin: BTP, s: '', q: 5000, p: 0, ctv: 16.41, c: '' }),
  ]);
  assert.ok(!ha(avvisi, /commissione|Controvalore|divisa/i), 'nessun avviso su righe corrette: ' + avvisi.join(' | '));
  console.log('✓ 1. Righe corrette (ETF, BTP nominale, cedola, commissione vuota o "1,74"): nessun avviso');
}

// 2. Commissione negativa e commissione non numerica → avviso; calcoli invariati (non numerica = 0)
{
  const neg = await importa([R({ c: -5 })]);
  assert.ok(ha(neg.avvisi, /commissione negativa/i), 'avviso commissione negativa: ' + neg.avvisi.join(' | '));
  const testo = await importa([R({ c: 'n.d.' })]);
  assert.ok(ha(testo.avvisi, /commissione "n\.d\." non leggibile/i), 'avviso commissione non numerica: ' + testo.avvisi.join(' | '));
  const zero = await importa([R({ c: 0 })]);
  assert.equal(JSON.stringify(testo.dati.map(({ years, ...x }) => x)), JSON.stringify(zero.dati.map(({ years, ...x }) => x)), 'commissione non numerica trattata come 0 (come prima)');
  console.log('✓ 2. Commissione negativa e non numerica ("n.d."): avviso; calcoli invariati');
}

// 3. Controvalore: tolleranza 0,01 € per riga; BTP confrontato con nominale × Prezzo / 100; cedole escluse
{
  const entro = await importa([R({ q: 10, p: 100, ctv: 1000.009 })]);
  assert.ok(!ha(entro.avvisi, /Controvalore/), 'scarto 0,009 € entro tolleranza');
  const oltre = await importa([R({ q: 10, p: 100, ctv: 1000.02 })]);
  assert.ok(ha(oltre.avvisi, /Controvalore/), 'scarto 0,02 € segnalato: ' + oltre.avvisi.join(' | '));
  const btpOk = await importa([R({ isin: BTP, q: 1000, p: 79.2, ctv: 792 })]);
  assert.ok(!ha(btpOk.avvisi, /Controvalore/), 'BTP: 1.000 nominali × 79,2 / 100 = 792, nessun avviso');
  const btpKo = await importa([R({ isin: BTP, q: 1000, p: 79.2, ctv: 79200 })]);
  assert.ok(ha(btpKo.avvisi, /Controvalore.*nominale/), 'BTP con Controvalore incoerente: avviso');
  const cedola = await importa([R({ desc: 'Stacco Cedole', isin: BTP, s: '', q: 5000, p: 0, ctv: 16.41, c: '' })]);
  assert.ok(!ha(cedola.avvisi, /Controvalore/), 'riga cedola esclusa dal controllo');
  console.log('✓ 3. Controvalore: tolleranza 0,01 € (0,009 ok, 0,02 segnalato), BTP su nominale/100, cedole escluse');
}

// 4. Divisa diversa da EUR → avviso (anche sulle cedole)
{
  const usd = await importa([R({ div: 'USD' }), R({ d: '28/10/2025', desc: 'Stacco Cedole', isin: BTP, s: '', div: 'USD', q: 5000, p: 0, ctv: 16.41, c: '' })]);
  assert.equal(usd.avvisi.filter(a => /divisa USD/.test(a)).length, 2, 'due righe in USD segnalate: ' + usd.avvisi.join(' | '));
  console.log('✓ 4. Divisa diversa da EUR: avviso su compravendita e cedola');
}

console.log('\nTutti i test controlli import superati.');
