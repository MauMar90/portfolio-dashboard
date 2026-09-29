// Test funzionale con jsdom per la logica di applyAutoPrices() dentro index.html.
// Non tocca PMC/XIRR/import: verifica solo l'integrazione di prices.json.
// Uso: node scripts/test-dashboard-prices.mjs

import { JSDOM } from 'jsdom';
import fs from 'node:fs/promises';
import assert from 'node:assert/strict';
import path from 'node:path';
import { fileURLToPath } from 'node:url';

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const htmlPath = path.join(__dirname, '..', 'index.html');
const html = await fs.readFile(htmlPath, 'utf8');

async function runScenario({ seedPositions, mockPricesJson, mockFetchThrows }) {
  const dom2 = new JSDOM(html, {
    url: 'http://localhost/',
    runScripts: 'dangerously',
    resources: undefined, // non scarica lo script XLSX esterno: non serve per questo test
    pretendToBeVisual: true,
    beforeParse(win) {
      win.localStorage.setItem('pfPositions', JSON.stringify(seedPositions));
      win.fetch = async () => {
        if (mockFetchThrows) throw new Error('rete non disponibile');
        return { ok: true, json: async () => mockPricesJson };
      };
    },
  });

  // applyAutoPrices() è asincrona e lanciata a fine script: aspettiamo un tick.
  await new Promise(r => setTimeout(r, 100));

  const win2 = dom2.window;
  const positions = JSON.parse(win2.localStorage.getItem('pfPositions') || '[]');
  const priceStatusText = win2.document.getElementById('priceStatus').textContent;
  const staleWarnEl = win2.document.getElementById('priceStaleWarn');
  const staleWarnText = staleWarnEl.textContent;
  const staleWarnVisible = staleWarnEl.style.display !== 'none';
  const meta = JSON.parse(win2.localStorage.getItem('pfPriceMeta') || 'null');

  return { positions, priceStatusText, staleWarnText, staleWarnVisible, meta };
}

// ── Scenario 1: caso normale, mix di ok / stale / manuale / chiusa ───────────
{
  const seedPositions = [
    { name: 'SWDA', isin: 'IE00B4L5Y983', ac: 'Azioni', qty: 10, price: 90, mkt: 95, closed: false },
    { name: 'XMME', isin: 'IE00BTJRMP35', ac: 'Azioni', qty: 5, price: 40, mkt: 42, closed: false },
    { name: 'PosizioneManuale', ac: 'Azioni', qty: 1, price: 10, mkt: 10, closed: false }, // niente isin
    { name: 'VecchiaChiusa', isin: 'IE00B53QG562', ac: 'Azioni', qty: 0, price: 30, mkt: 0, closed: true },
    { name: 'BTP Futura', isin: 'IT0005442089', ac: 'Obbligazioni', qty: 3, price: 97.5, mkt: 98.2, closed: false, isNominal: true },
  ];
  const mockPricesJson = {
    generated_at: '2026-08-23T17:00:00Z',
    entries: {
      'IE00B4L5Y983': { close: 102.34, status: 'ok', session_date: '2026-08-23' },
      'IE00BTJRMP35': { close: 45.0, status: 'stale', session_date: '2026-08-15' },
      'IE00B53QG562': { close: 999, status: 'ok', session_date: '2026-08-23' }, // chiusa: non deve applicarsi
      // IT0005442089 (BTP Futura) volutamente NON presente nel mock: deve restare invariato e SENZA warning (fuori scope)
    },
    warnings: [],
  };

  const { positions, priceStatusText, staleWarnText, staleWarnVisible, meta } = await runScenario({ seedPositions, mockPricesJson });

  assert.equal(positions.find(p => p.name === 'SWDA').mkt, 102.34, 'SWDA: mkt deve aggiornarsi (status ok)');
  assert.equal(positions.find(p => p.name === 'XMME').mkt, 42, 'XMME: status stale -> mkt NON deve cambiare, resta il valore precedente (42)');
  assert.equal(positions.find(p => p.name === 'PosizioneManuale').mkt, 10, 'Posizione manuale (senza isin): MAI toccata');
  assert.equal(positions.find(p => p.name === 'VecchiaChiusa').mkt, 0, 'Posizione chiusa: MAI toccata anche se isin presente nel json');
  assert.equal(positions.find(p => p.name === 'BTP Futura').mkt, 98.2,'BTP Futura: mai automatizzato, resta manuale');

  assert.ok(priceStatusText.includes('23/8/2026') || priceStatusText.includes('23/08/2026'), 'priceStatus deve mostrare la data della sessione applicata (solo ok): ' + priceStatusText);
  assert.ok(!priceStatusText.includes('⚠'), 'priceStatus (header) NON deve più contenere il warning inline: ' + priceStatusText);
  assert.equal(staleWarnVisible, true, 'il banner stale deve essere visibile');
  assert.equal(staleWarnText, '⚠️ Prezzi non aggiornati: XMME', 'il banner deve elencare esplicitamente i nomi, non un tooltip (e NON deve includere BTP Futura, fuori scope): ' + staleWarnText);
  assert.equal(meta.staleNames.length, 1);
  assert.equal(meta.reachable, true);

  console.log('✓ Scenario 1 (normale: ok applicato, stale NON applicato ma segnalato in chiaro, manuale/chiusa/BTP intatte) OK');
}

// ── Scenario 2: prices.json irraggiungibile -> nessuna modifica a positions ──
{
  const seedPositions = [
    { name: 'SWDA', isin: 'IE00B4L5Y983', ac: 'Azioni', qty: 10, price: 90, mkt: 95, closed: false },
  ];
  const { positions } = await runScenario({
    seedPositions,
    mockPricesJson: null,
    mockFetchThrows: true,
  });
  assert.equal(positions.find(p => p.name === 'SWDA').mkt, 95, 'Fetch fallito: mkt deve restare quello di localStorage, invariato');
  console.log('✓ Scenario 2 (prices.json irraggiungibile, dashboard resta invariata) OK');
}

// ── Scenario 3: entries malformato (nessun campo entries) -> nessuna modifica ─
{
  const seedPositions = [
    { name: 'SWDA', isin: 'IE00B4L5Y983', ac: 'Azioni', qty: 10, price: 90, mkt: 95, closed: false },
  ];
  const { positions } = await runScenario({
    seedPositions,
    mockPricesJson: { foo: 'bar' }, // niente campo "entries"
  });
  assert.equal(positions.find(p => p.name === 'SWDA').mkt, 95, 'JSON malformato: mkt deve restare invariato');
  console.log('✓ Scenario 3 (prices.json malformato, dashboard resta invariata) OK');
}

// ── Scenario 4: tutti gli ETF risultano stale -> nessun mkt applicato, ma il
//    banner deve comunque elencarli tutti in chiaro ───────────────────────────
{
  const seedPositions = [
    { name: 'SWDA', isin: 'IE00B4L5Y983', ac: 'Azioni', qty: 10, price: 90, mkt: 95, closed: false },
    { name: 'XMME', isin: 'IE00BTJRMP35', ac: 'Azioni', qty: 5, price: 40, mkt: 42, closed: false },
  ];
  const mockPricesJson = {
    generated_at: '2026-08-23T17:00:00Z',
    entries: {
      'IE00B4L5Y983': { close: 110, status: 'stale', session_date: '2026-08-10' },
      'IE00BTJRMP35': { close: 45, status: 'stale', session_date: '2026-08-10' },
    },
    warnings: [],
  };
  const { positions, priceStatusText, staleWarnText, staleWarnVisible, meta } =
    await runScenario({ seedPositions, mockPricesJson });

  assert.equal(positions.find(p => p.name === 'SWDA').mkt, 95, 'Tutti stale: SWDA.mkt deve restare invariato');
  assert.equal(positions.find(p => p.name === 'XMME').mkt, 42, 'Tutti stale: XMME.mkt deve restare invariato');
  assert.equal(staleWarnVisible, true, 'il banner deve comunque comparire');
  assert.ok(staleWarnText.includes('SWDA') && staleWarnText.includes('XMME'), 'entrambi i titoli devono comparire nel banner: ' + staleWarnText);
  assert.equal(priceStatusText, '', 'senza alcun ok mai applicato, priceStatus deve restare vuoto: ' + JSON.stringify(priceStatusText));
  assert.equal(meta.sessionDate, null);

  console.log('✓ Scenario 4 (tutti stale: nessun mkt toccato, entrambi segnalati per nome) OK');
}

// ── Scenario 5: ISIN coperto dalla pipeline ma ASSENTE da entries (errore di
//    fetch mai riuscito) -> deve generare warning, non solo essere ignorato ──
{
  const seedPositions = [
    { name: 'SWDA', isin: 'IE00B4L5Y983', ac: 'Azioni', qty: 10, price: 90, mkt: 95, closed: false },
    { name: 'EM57', isin: 'LU1287023003', ac: 'Obbligazioni', qty: 3, price: 150, mkt: 155, closed: false },
  ];
  const mockPricesJson = {
    generated_at: '2026-08-23T17:00:00Z',
    entries: {
      'IE00B4L5Y983': { close: 102.34, status: 'ok', session_date: '2026-08-23' },
      // LU1287023003 (EM57) volutamente assente: simula un fetch mai riuscito
    },
    warnings: ['EM57.MI (LU1287023003): network down'],
  };
  const { positions, staleWarnText, staleWarnVisible } = await runScenario({ seedPositions, mockPricesJson });

  assert.equal(positions.find(p => p.name === 'SWDA').mkt, 102.34, 'SWDA aggiornato normalmente');
  assert.equal(positions.find(p => p.name === 'EM57').mkt, 155, 'EM57: entry assente -> mkt NON deve essere toccato');
  assert.equal(staleWarnVisible, true, 'CORREZIONE: un ISIN coperto ma assente da entries deve generare un warning visibile');
  assert.ok(staleWarnText.includes('EM57'), 'EM57 deve comparire esplicitamente nel banner: ' + staleWarnText);

  console.log('✓ Scenario 5 (ISIN coperto ma assente per errore reale -> warning visibile) OK');
}

// ── Scenario 6: status 'error' esplicito (difensivo: update-prices.js non lo
//    scrive mai oggi, ma se comparisse la dashboard deve gestirlo comunque) ──
{
  const seedPositions = [
    { name: 'X57E', isin: 'LU0290357176', ac: 'Obbligazioni', qty: 2, price: 200, mkt: 210, closed: false },
  ];
  const mockPricesJson = {
    generated_at: '2026-08-23T17:00:00Z',
    entries: {
      'LU0290357176': { close: 0, status: 'error', session_date: null },
    },
    warnings: [],
  };
  const { positions, staleWarnText, staleWarnVisible } = await runScenario({ seedPositions, mockPricesJson });

  assert.equal(positions.find(p => p.name === 'X57E').mkt, 210, "status 'error': mkt NON deve essere toccato");
  assert.equal(staleWarnVisible, true, "status 'error' deve generare un warning visibile");
  assert.ok(staleWarnText.includes('X57E'), 'X57E deve comparire nel banner: ' + staleWarnText);

  console.log("✓ Scenario 6 (status 'error' esplicito, gestione difensiva) OK");
}

// ── Scenario 7: primo import, nessun mkt precedente (mkt=0) -----------------
//    stale/error non devono MAI popolare mkt con un valore inventato.
{
  const seedPositions = [
    { name: 'CSEMU', isin: 'IE00B53QG562', ac: 'Azioni', qty: 8, price: 28, mkt: 0, closed: false }, // primo import: mkt ancora 0
    { name: 'LCJP', isin: 'LU1781541252', ac: 'Azioni', qty: 4, price: 22, mkt: 0, closed: false },  // primo import: mkt ancora 0
  ];
  const mockPricesJson = {
    generated_at: '2026-08-23T17:00:00Z',
    entries: {
      'IE00B53QG562': { close: 30.0, status: 'stale', session_date: '2026-08-01' }, // stale
      // LU1781541252 (LCJP) assente: errore
    },
    warnings: [],
  };
  const { positions, staleWarnText } = await runScenario({ seedPositions, mockPricesJson });

  assert.equal(positions.find(p => p.name === 'CSEMU').mkt, 0, "Primo import + stale: mkt deve restare 0, MAI il close stale applicato");
  assert.equal(positions.find(p => p.name === 'LCJP').mkt, 0, "Primo import + entry assente (errore): mkt deve restare 0, non undefined/inventato");
  assert.ok(staleWarnText.includes('CSEMU') && staleWarnText.includes('LCJP'), 'entrambi segnalati come non aggiornati: ' + staleWarnText);

  console.log('✓ Scenario 7 (primo import senza mkt precedente: stale/error non inventano un prezzo) OK');
}

// ── Scenario 8: EXUS e Gold (nuova aggiunta) riconosciuti come coperti da
//    AUTO_PRICE_ISINS lato dashboard -- non solo che il fetch funzioni lato
//    script (già verificato in test-update-prices.mjs), ma che la dashboard
//    stessa li tratti come titoli automatizzati: EXUS con status ok si
//    aggiorna, Gold assente da entries genera un warning (prova che il suo
//    ISIN è ora nell'insieme "coperto", altrimenti sarebbe stato ignorato
//    silenziosamente come un titolo fuori scope).
{
  const seedPositions = [
    { name: 'EXUS', isin: 'IE0006WW1TQ4', ac: 'Azioni', qty: 6, price: 38, mkt: 40, closed: false },
    { name: 'GOLD', isin: 'FR0013416716', ac: 'Oro', qty: 2, price: 150, mkt: 150, closed: false },
  ];
  const mockPricesJson = {
    generated_at: '2026-08-26T19:56:39.013Z',
    entries: {
      'IE0006WW1TQ4': { close: 40.53, status: 'ok', session_date: '2026-08-26' },
      // FR0013416716 (Gold) volutamente assente: simula un fetch non ancora riuscito
    },
    warnings: [],
  };
  const { positions, staleWarnText, staleWarnVisible } = await runScenario({ seedPositions, mockPricesJson });

  assert.equal(positions.find(p => p.name === 'EXUS').mkt, 40.53, 'EXUS (nuova aggiunta): status ok -> mkt deve aggiornarsi normalmente');
  assert.equal(positions.find(p => p.name === 'GOLD').mkt, 150, 'Gold (nuova aggiunta): entry assente -> mkt NON deve essere toccato');
  assert.equal(staleWarnVisible, true, 'Gold deve generare un warning visibile, prova che il suo ISIN è ora incluso in AUTO_PRICE_ISINS');
  assert.ok(staleWarnText.includes('GOLD'), 'GOLD deve comparire esplicitamente nel banner: ' + staleWarnText);
  assert.ok(!staleWarnText.includes('EXUS'), 'EXUS non deve comparire nel banner (è andato ok): ' + staleWarnText);

  console.log('✓ Scenario 8 (EXUS e Gold riconosciuti come coperti da AUTO_PRICE_ISINS) OK');
}

console.log('\n✓ Tutti gli scenari jsdom sono passati.');
