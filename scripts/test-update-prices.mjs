// Test manuale (non un framework): mocka global.fetch per validare la logica di
// update-prices.js senza dipendere dalla raggiungibilità reale di Yahoo.
// Uso: node scripts/test-update-prices.mjs
//
// NOTA: update-prices.js esegue main() automaticamente al top-level (nessun
// export, nessun guard) -- per questo la logica pura di data/soglia vive in
// price-utils.mjs, importabile liberamente senza innescare fetch reali.

import fs from 'node:fs/promises';
import fsSync from 'node:fs';
import os from 'node:os';
import crypto from 'node:crypto';
import path from 'node:path';
import { fileURLToPath, pathToFileURL } from 'node:url';
import assert from 'node:assert/strict';
import { toMilanSessionDate, daysBetween, deriveStaleStatus, STALE_DAYS } from './price-utils.mjs';

const __dirname = path.dirname(fileURLToPath(import.meta.url));

// ── Isolamento: il test NON deve mai scrivere il prices.json del repository ──
// update-prices.js calcola il percorso di prices.json dalla posizione del proprio file
// (__dirname/../prices.json) e non è modificabile (pipeline congelata). Il test ne esegue
// quindi una COPIA identica in una cartella temporanea (tmp/scripts/), insieme alle sue
// dipendenze locali: la copia scrive tmp/prices.json. GITHUB_OUTPUT punta anch'esso a tmp.
const REPO_PRICES_PATH = path.join(__dirname, '..', 'prices.json');
const sha256 = p => crypto.createHash('sha256').update(fsSync.existsSync(p) ? fsSync.readFileSync(p) : '').digest('hex');
const repoPricesHashBefore = sha256(REPO_PRICES_PATH);

const TMP_ROOT = await fs.mkdtemp(path.join(os.tmpdir(), 'test-update-prices-'));
process.on('exit', () => { try { fsSync.rmSync(TMP_ROOT, { recursive: true, force: true }); } catch {} });
const TMP_SCRIPTS = path.join(TMP_ROOT, 'scripts');
await fs.mkdir(TMP_SCRIPTS);
// Copia di update-prices.js e di tutte le sue dipendenze locali (import relativi, ricorsivi):
// se la pipeline ne aggiungesse una nuova, viene copiata automaticamente.
const daCopiare = ['update-prices.js'];
for (let i = 0; i < daCopiare.length; i++) {
  const src = path.join(__dirname, daCopiare[i]);
  assert.ok(fsSync.existsSync(src), `dipendenza della pipeline non trovata: ${daCopiare[i]}`);
  await fs.copyFile(src, path.join(TMP_SCRIPTS, daCopiare[i]));
  const codice = fsSync.readFileSync(src, 'utf8');
  for (const m of codice.matchAll(/from\s+['"]\.\/([^'"]+)['"]|import\(\s*['"]\.\/([^'"?]+)/g)) {
    const dep = m[1] || m[2];
    if (!daCopiare.includes(dep)) daCopiare.push(dep);
  }
}
const PIPELINE_URL = pathToFileURL(path.join(TMP_SCRIPTS, 'update-prices.js')).href;
const PRICES_PATH = path.join(TMP_ROOT, 'prices.json');
const GITHUB_OUTPUT_PATH = path.join(TMP_ROOT, '.test-github-output');
console.log(`Pipeline copiata in cartella temporanea con le sue dipendenze: ${daCopiare.join(', ')}`);

const nowSec = Math.floor(Date.now() / 1000);
const daysAgoSec = (n) => nowSec - n * 86400;
const todayMilan = toMilanSessionDate(nowSec);

// ══════════════════════════════════════════════════════════════════════════
// PARTE 1 — unit test puri (nessun fetch, nessun I/O): timezone e soglia
// ══════════════════════════════════════════════════════════════════════════

// CET (inverno, UTC+1): 23:30 UTC del 15 gennaio è già le 00:30 del 16 a Milano.
assert.equal(
  toMilanSessionDate(Math.floor(Date.UTC(2026, 0, 15, 23, 30, 0) / 1000)),
  '2026-01-16',
  'CET: 23:30 UTC del 15/01 deve cadere nel 16/01 a Milano (UTC+1)'
);
// CET, stesso giorno a metà giornata: nessuno scavalcamento di data.
assert.equal(
  toMilanSessionDate(Math.floor(Date.UTC(2026, 0, 15, 12, 0, 0) / 1000)),
  '2026-01-15',
  'CET: mezzogiorno UTC resta lo stesso giorno civile a Milano'
);
// CEST (estate, UTC+2): 22:30 UTC del 15 luglio è già le 00:30 del 16 a Milano.
assert.equal(
  toMilanSessionDate(Math.floor(Date.UTC(2026, 6, 15, 22, 30, 0) / 1000)),
  '2026-07-16',
  'CEST: 22:30 UTC del 15/07 deve cadere nel 16/07 a Milano (UTC+2)'
);
// CEST, stesso giorno a metà giornata.
assert.equal(
  toMilanSessionDate(Math.floor(Date.UTC(2026, 6, 15, 12, 0, 0) / 1000)),
  '2026-07-15',
  'CEST: mezzogiorno UTC resta lo stesso giorno civile a Milano'
);

// Soglia: esattamente 5 giorni è ancora 'ok' (regola: stale solo se età > soglia).
assert.equal(deriveStaleStatus('2026-08-01', '2026-08-06'), 'ok', '5 giorni esatti: ancora ok');
assert.equal(deriveStaleStatus('2026-08-01', '2026-08-07'), 'stale', '6 giorni: stale');
assert.equal(STALE_DAYS, 5, 'soglia invariata a 5 giorni (decisione esplicita, non va cambiata qui)');

console.log('✓ Parte 1 (timezone Europe/Rome + soglia) OK');

// ══════════════════════════════════════════════════════════════════════════
// PARTE 2 — integrazione con fetch mockato, su più esecuzioni consecutive
// ══════════════════════════════════════════════════════════════════════════

function chartResponse({
  close, exchangeName = 'MIL', fullExchangeName = 'Milan', currency = 'EUR', instrumentType = 'ETF',
  tsOffsetDays = 0, marketOpen = false, omitTradingPeriod = false, omitRegularMarketTime = false,
  priorBar = null, // { close, tsOffsetDays } opzionale: barra precedente reale, per il fallback a mercato aperto
}) {
  const ts = daysAgoSec(tsOffsetDays);
  const regularEnd = ts + 30600; // sessione di 8.5h, come Milano (09:00-17:30), coerente con la barra "odierna" di questo mock
  const regularMarketTime = marketOpen
    ? regularEnd - 3600  // 1h prima della chiusura prevista: sessione ancora aperta
    : regularEnd + 300;  // 5 min dopo la chiusura prevista: sessione conclusa, close definitivo

  const timestamps = priorBar ? [daysAgoSec(priorBar.tsOffsetDays), ts] : [ts];
  const closes = priorBar ? [priorBar.close, close] : [close];

  const meta = { exchangeName, fullExchangeName, currency, instrumentType };
  if (!omitRegularMarketTime) meta.regularMarketTime = regularMarketTime;
  if (!omitTradingPeriod) meta.currentTradingPeriod = { regular: { start: ts, end: regularEnd } };

  return {
    chart: {
      result: [{ meta, timestamp: timestamps, indicators: { quote: [{ close: closes }] } }],
      error: null,
    },
  };
}

// Scenario stabile, riusato identico in run 1 e run 2 per verificare che una
// ripetizione a parità di dati non generi un cambiamento sostanziale.
const MOCKS = {
  'SWDA.MI':  () => ({ ok: true, json: async () => chartResponse({ close: 102.34, tsOffsetDays: 0 }) }),              // ok, fresco, nessun preesistente
  'XMME.MI':  () => ({ ok: true, json: async () => chartResponse({ close: 45.67, tsOffsetDays: 8 }) }),               // stale (>5gg), nessun preesistente: va comunque scritto
  'CSEMU.MI': () => { throw new Error('network down'); },                                                             // errore di rete; preesistente FRESCO
  'LCJP.MI':  () => ({ ok: false, status: 503 }),                                                                     // errore HTTP; preesistente VECCHIO, status 'ok' -> deve ricalcolarsi a 'stale'
  'XEON.MI':  () => { throw new Error('network down'); },                                                             // errore di rete; NESSUN preesistente: non deve inventare nulla
  'X57E.MI':  () => ({ ok: true, json: async () => chartResponse({ close: 0, tsOffsetDays: 0 }) }),                   // prezzo 0 -> rifiutato; preesistente già 'stale', deve solo aggiornare il conteggio giorni
  'EM57.MI':  () => ({ ok: true, json: async () => chartResponse({ close: 98.5, tsOffsetDays: 1 }) }),                // ok, fresco
  'EXUS.MI':  () => ({ ok: true, json: async () => chartResponse({ close: 40.53, tsOffsetDays: 0 }) }),               // ok, fresco -- nuova aggiunta
  'GOLD.MI':  () => ({ ok: true, json: async () => chartResponse({ close: 157.74, tsOffsetDays: 0 }) }),              // ok, fresco -- nuova aggiunta
};

global.fetch = async (url) => {
  const ticker = decodeURIComponent(url.match(/chart\/([^?]+)/)[1]);
  const mock = MOCKS[ticker];
  if (!mock) throw new Error('mock non definito per ' + ticker);
  return mock();
};

const oldDate = toMilanSessionDate(daysAgoSec(10)); // ben oltre la soglia di 5gg
const freshDate = toMilanSessionDate(daysAgoSec(2)); // entro la soglia
const alreadyStaleDate = toMilanSessionDate(daysAgoSec(7));

const preExisting = {
  generated_at: '2000-01-01T00:00:00Z',
  entries: {
    'IE00B53QG562': { ticker: 'CSEMU.MI', close: 30.0, currency: 'EUR', exchange: 'MIL', session_date: freshDate, fetched_at: '2000-01-01T00:00:00Z', source: 'yahoo_chart', status: 'ok' },
    'LU1781541252': { ticker: 'LCJP.MI', close: 20.0, currency: 'EUR', exchange: 'MIL', session_date: oldDate, fetched_at: '2000-01-01T00:00:00Z', source: 'yahoo_chart', status: 'ok' },
    'LU0290357176': { ticker: 'X57E.MI', close: 95.0, currency: 'EUR', exchange: 'MIL', session_date: alreadyStaleDate, fetched_at: '2000-01-01T00:00:00Z', source: 'yahoo_chart', status: 'stale', note: `ultimo close disponibile: ${alreadyStaleDate} (7 giorni fa)` },
  },
  warnings: [],
};

async function resetGithubOutput() {
  try { await fs.rm(GITHUB_OUTPUT_PATH); } catch {}
  process.env.GITHUB_OUTPUT = GITHUB_OUTPUT_PATH;
}

async function readSubstantiveChangeFlag() {
  const raw = await fs.readFile(GITHUB_OUTPUT_PATH, 'utf8').catch(() => '');
  const line = raw.split('\n').find(l => l.startsWith('substantive_change='));
  return line ? line.split('=')[1] === 'true' : null;
}

async function runUpdatePrices() {
  await resetGithubOutput();
  await import(PIPELINE_URL + '?t=' + Date.now() + Math.random());
  await new Promise(r => setTimeout(r, 200));
  const result = JSON.parse(await fs.readFile(PRICES_PATH, 'utf8'));
  const substantiveChange = await readSubstantiveChangeFlag();
  return { result, substantiveChange };
}

// ── Run 1: stato iniziale pre-popolato, prima esecuzione dello script ────────
await fs.writeFile(PRICES_PATH, JSON.stringify(preExisting, null, 2));
const run1 = await runUpdatePrices();
const e1 = run1.result.entries;

assert.equal(e1['IE00B4L5Y983'].status, 'ok', 'SWDA deve essere ok');
assert.equal(e1['IE00B4L5Y983'].close, 102.34);

assert.equal(e1['IE00BTJRMP35'].status, 'stale', 'XMME deve essere stale (8gg), pur senza preesistente');
assert.equal(e1['IE00BTJRMP35'].close, 45.67, 'XMME stale ma il close va comunque scritto');

assert.equal(e1['IE00B53QG562'].close, 30.0, 'CSEMU: errore di rete -> valore precedente invariato');
assert.equal(e1['IE00B53QG562'].status, 'ok', 'CSEMU: preesistente ancora fresco (2gg) -> resta ok anche se il fetch fallisce');

assert.equal(e1['LU1781541252'].close, 20.0, 'LCJP: HTTP 503 -> valore precedente invariato');
assert.equal(e1['LU1781541252'].status, 'stale', 'LCJP: CORREZIONE CRITICA — preesistente ok ma vecchio di 10gg, fetch fallisce di nuovo -> deve ricalcolarsi a stale, non restare ok congelato');
assert.ok(e1['LU1781541252'].note && e1['LU1781541252'].note.includes('10 giorni fa'), 'LCJP: il note deve riportare i giorni reali trascorsi: ' + JSON.stringify(e1['LU1781541252'].note));

assert.equal(e1['LU0290358497'], undefined, 'XEON: nessun preesistente + fetch fallito -> NON deve comparire alcuna entry inventata');

assert.equal(e1['LU0290357176'].status, 'stale', 'X57E: prezzo 0 rifiutato, resta stale');
assert.equal(e1['LU0290357176'].close, 95.0, 'X57E: valore precedente invariato');
assert.ok(e1['LU0290357176'].note.includes('7 giorni fa'), 'X57E: il conteggio giorni nel note non deve restare congelato: ' + e1['LU0290357176'].note);

assert.equal(e1['LU1287023003'].status, 'ok', 'EM57 deve essere ok');
assert.equal(e1['LU1287023003'].close, 98.5);

assert.equal(e1['IE0006WW1TQ4'].status, 'ok', 'EXUS (nuova aggiunta): deve essere coperto dalla pipeline e risultare ok');
assert.equal(e1['IE0006WW1TQ4'].ticker, 'EXUS.MI');
assert.equal(e1['IE0006WW1TQ4'].close, 40.53);

assert.equal(e1['FR0013416716'].status, 'ok', 'Gold (nuova aggiunta): deve essere coperto dalla pipeline e risultare ok');
assert.equal(e1['FR0013416716'].ticker, 'GOLD.MI');
assert.equal(e1['FR0013416716'].close, 157.74);

assert.ok(!e1['IT0005442089'], 'BTP Futura non deve MAI comparire in prices.json');

assert.equal(run1.result.warnings.length, 4, 'CSEMU, LCJP, XEON, X57E devono generare un warning ciascuno');
assert.equal(run1.substantiveChange, true, 'run 1: rispetto allo stato preesistente ci sono cambiamenti sostanziali reali');

console.log('✓ Run 1 (stato iniziale, incl. ricalcolo status su entry preesistenti) OK');

// ── Run 2: stessi identici mock, eseguito subito dopo -> nessun cambiamento
//    sostanziale, il commit non dovrebbe scattare ─────────────────────────────
const run2 = await runUpdatePrices();
assert.equal(run2.substantiveChange, false, 'run 2: dati identici al run 1 -> nessun cambiamento sostanziale, niente commit');
// fetched_at cambia legittimamente ad ogni fetch riuscito anche a prezzo
// invariato (è un timestamp tecnico, correttamente escluso dal confronto
// sostanziale) -- qui confrontiamo solo i campi che DEVONO restare identici.
for (const isin of Object.keys(run1.result.entries)) {
  const a = run1.result.entries[isin], b = run2.result.entries[isin];
  assert.equal(b.close, a.close, `run 2: close invariato per ${isin}`);
  assert.equal(b.status, a.status, `run 2: status invariato per ${isin}`);
  assert.equal(b.session_date, a.session_date, `run 2: session_date invariata per ${isin}`);
  assert.equal(b.note, a.note, `run 2: note invariato per ${isin}`);
}

console.log('✓ Run 2 (ripetizione identica -> substantive_change=false) OK');

// ── Run 3: un prezzo reale cambia -> deve tornare a essere un cambiamento
//    sostanziale ────────────────────────────────────────────────────────────
MOCKS['SWDA.MI'] = () => ({ ok: true, json: async () => chartResponse({ close: 103.00, tsOffsetDays: 0 }) });
const run3 = await runUpdatePrices();
assert.equal(run3.substantiveChange, true, 'run 3: SWDA è cambiato di prezzo -> deve tornare true');
assert.equal(run3.result.entries['IE00B4L5Y983'].close, 103.00);

console.log('✓ Run 3 (variazione di prezzo reale -> substantive_change=true) OK');

await fs.rm(GITHUB_OUTPUT_PATH).catch(() => {});

// ══════════════════════════════════════════════════════════════════════════
// PARTE 3 — regola "sessione conclusa" (regularMarketTime vs
// currentTradingPeriod.regular.end), verificata su payload Yahoo reali:
// a mercato aperto l'ultima barra è un riflesso del prezzo intraday, non un
// close. fetchOne() è chiamata direttamente, in isolamento, con un mock
// dedicato per ciascuno scenario.
// ══════════════════════════════════════════════════════════════════════════

const { fetchOne } = await import(PIPELINE_URL + '?t=' + Date.now() + Math.random());
// L'import innesca comunque, come sempre, l'esecuzione automatica di main() in
// background (nessun guard, per progettazione invariata). Aspettiamo che le sue
// 7 chiamate asincrone si esauriscano prima di riassegnare global.fetch qui
// sotto, per evitare che si sovrappongano ai mock dedicati di questa sezione.
await new Promise(r => setTimeout(r, 300));

// 1. Mercato aperto -> ultimo close intraday scartato -> selezionato il
//    precedente close valido (entro soglia -> ok)
{
  global.fetch = async () => ({
    ok: true,
    json: async () => chartResponse({
      close: 999,            // valore "intraday" fittizio, deve essere scartato
      tsOffsetDays: 0,
      marketOpen: true,
      priorBar: { close: 100, tsOffsetDays: 1 }, // close reale di ieri, entro soglia
    }),
  });
  const entry = await fetchOne('TEST.MI', todayMilan);
  assert.equal(entry.close, 100, 'Mercato aperto: il close intraday (999) deve essere scartato, selezionato il precedente (100)');
  assert.equal(entry.status, 'ok', 'Precedente entro soglia (1gg) -> ok');
  assert.equal(entry.session_date, toMilanSessionDate(daysAgoSec(1)), 'session_date deve essere quella del precedente close reale, non di oggi');
}
console.log('✓ Test 3.1 (mercato aperto -> close intraday scartato, precedente selezionato) OK');

// 2. Mercato chiuso -> ultimo close odierno accettato normalmente
{
  global.fetch = async () => ({
    ok: true,
    json: async () => chartResponse({ close: 105, tsOffsetDays: 0, marketOpen: false }),
  });
  const entry = await fetchOne('TEST.MI', todayMilan);
  assert.equal(entry.close, 105, 'Mercato chiuso: il close odierno deve essere accettato');
  assert.equal(entry.status, 'ok');
  assert.equal(entry.session_date, todayMilan);
}
console.log('✓ Test 3.2 (mercato chiuso -> close odierno accettato) OK');

// 3. Mercato aperto + precedente entro soglia -> comportamento normale (variante
//    con un precedente più vecchio di 1gg ma comunque entro i 5gg, per
//    differenziare esplicitamente dal Test 3.1)
{
  global.fetch = async () => ({
    ok: true,
    json: async () => chartResponse({
      close: 999, tsOffsetDays: 0, marketOpen: true,
      priorBar: { close: 97.5, tsOffsetDays: 3 }, // entro soglia (3gg <= 5gg)
    }),
  });
  const entry = await fetchOne('TEST.MI', todayMilan);
  assert.equal(entry.close, 97.5);
  assert.equal(entry.status, 'ok', 'Precedente a 3gg, entro soglia -> ok, comportamento normale');
}
console.log('✓ Test 3.3 (mercato aperto + precedente entro soglia -> comportamento normale) OK');

// 4. Mercato aperto + precedente OLTRE soglia -> stale, mkt non deve essere
//    modificato (questo lo garantisce la dashboard leggendo status, qui
//    verifichiamo solo che update-prices.js produca lo status corretto)
{
  global.fetch = async () => ({
    ok: true,
    json: async () => chartResponse({
      close: 999, tsOffsetDays: 0, marketOpen: true,
      priorBar: { close: 95.0, tsOffsetDays: 8 }, // oltre soglia (8gg > 5gg)
    }),
  });
  const entry = await fetchOne('TEST.MI', todayMilan);
  assert.equal(entry.close, 95.0, 'Il close scritto è quello reale precedente, non il valore intraday scartato');
  assert.equal(entry.status, 'stale', 'Precedente a 8gg, oltre soglia -> stale (la dashboard non applicherà questo a mkt)');
  assert.ok(entry.note && entry.note.includes('8 giorni fa'), 'Il note deve riportare i giorni reali del close realmente usato: ' + entry.note);
}
console.log('✓ Test 3.4 (mercato aperto + precedente oltre soglia -> stale, close reale preservato) OK');

// 5. regular.end mancante/non valido -> nessun dato intraday accettato:
//    fetchOne deve lanciare un errore (comportamento difensivo, equivalente a
//    un errore di validazione: nessun close scritto da questa chiamata)
for (const [label, overrides] of [
  ['currentTradingPeriod assente', { omitTradingPeriod: true }],
  ['regularMarketTime assente', { omitRegularMarketTime: true }],
]) {
  global.fetch = async () => ({
    ok: true,
    json: async () => chartResponse({ close: 110, tsOffsetDays: 0, ...overrides }),
  });
  await assert.rejects(
    () => fetchOne('TEST.MI', todayMilan),
    /regular\.start\/end|regularMarketTime/,
    `Caso limite (${label}): fetchOne deve lanciare un errore, non assumere che il dato sia un close definitivo`
  );
}
// Variante aggiuntiva: regular.end presente ma non numerico (stringa) -- stesso trattamento
{
  global.fetch = async () => ({
    ok: true,
    json: async () => {
      const base = chartResponse({ close: 110, tsOffsetDays: 0 });
      base.chart.result[0].meta.currentTradingPeriod.regular.end = 'non-un-numero';
      return base;
    },
  });
  await assert.rejects(() => fetchOne('TEST.MI', todayMilan), /regular\.start\/end|regularMarketTime/, 'regular.end non numerico: deve essere trattato come mancante, non come valido');
}
console.log('✓ Test 3.5 (regular.end/regularMarketTime mancanti o non validi -> errore, nessun dato accettato) OK');

// 6. Caso anomalo: l'ultima barra della serie NON appartiene alla sessione
//    corrente (es. la sessione odierna non è ancora rappresentata in timestamp[]).
//    regularMarketTime < regular.end da solo NON deve bastare per scartare
//    indiscriminatamente l'ultimo close storico -- va verificato esplicitamente
//    che quell'ultima barra ricada davvero nella finestra [regular.start, regular.end].
{
  const lastBarTs = daysAgoSec(2); // ultima barra della serie: chiusura reale di 2 giorni fa
  const todaySessionStart = daysAgoSec(0); // la sessione "corrente" secondo Yahoo è oggi...
  const todaySessionEnd = todaySessionStart + 30600;
  global.fetch = async () => ({
    ok: true,
    json: async () => ({
      chart: {
        result: [{
          meta: {
            exchangeName: 'MIL', fullExchangeName: 'Milan', currency: 'EUR', instrumentType: 'ETF',
            // ...ma la sessione odierna non è ancora nemmeno iniziata al momento del
            // fetch (es. pre-apertura): l'ultimo scambio noto è precedente a
            // todaySessionStart, quindi ovviamente anche < todaySessionEnd.
            regularMarketTime: todaySessionStart - 3600,
            currentTradingPeriod: { regular: { start: todaySessionStart, end: todaySessionEnd } },
          },
          timestamp: [lastBarTs], // la serie NON contiene alcuna barra per la sessione odierna
          indicators: { quote: [{ close: [88.0] }] },
        }],
        error: null,
      },
    }),
  });
  const entry = await fetchOne('TEST.MI', todayMilan);
  assert.equal(entry.close, 88.0, "L'ultima barra della serie non è quella odierna (non ricade in [regular.start, regular.end]): è già un close storico reale, NON va scartata come intraday");
  const expectedStatus = deriveStaleStatus(toMilanSessionDate(lastBarTs), todayMilan);
  assert.equal(entry.status, expectedStatus, 'status coerente con la reale età di quel close (2 giorni fa)');
}
console.log('✓ Test 3.6 (ultima barra non appartiene alla sessione corrente -> close storico NON scartato) OK');

// Le esecuzioni di main() lanciate in background dagli import vanno lasciate concludere
// prima di verificare che il prices.json del repository sia rimasto intatto.
await new Promise(r => setTimeout(r, 500));
assert.equal(sha256(REPO_PRICES_PATH), repoPricesHashBefore, 'il prices.json del repository NON deve essere modificato dal test');
assert.ok(!fsSync.existsSync(path.join(__dirname, '.test-github-output')), 'nessun file .test-github-output nel repository');
console.log('✓ prices.json del repository intatto (hash invariato), nessun file di output nel repository');

console.log('\n✓ Tutti i test passati.');


