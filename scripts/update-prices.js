// Aggiorna prices.json interrogando l'endpoint Chart di Yahoo Finance per i 7 ETF
// quotati su Borsa Italiana. Progettato per essere eseguito da una GitHub Action
// schedulata, ma può anche essere lanciato manualmente: `node scripts/update-prices.js`.
//
// Principi non negoziabili (vedi discussione di progetto):
// - Ogni ISIN è elaborato in modo indipendente: un fallimento su un titolo non
//   deve compromettere gli altri.
// - "Stale" (close valido ma di una sessione precedente, oltre la soglia) è
//   distinto da "errore" (dato mancante/non valido) ma entrambi hanno lo stesso
//   effetto pratico: NON aggiornano positions[].mkt lato dashboard. La soglia è
//   una regola di validità, non solo un'informazione da mostrare in UI.
// - Un fetch fallito NON scrive mai un prezzo nuovo, ma l'entry preesistente per
//   quell'ISIN va comunque rivalutata rispetto a OGGI: uno status 'ok' assegnato
//   l'ultima volta che il fetch è riuscito non deve restare congelato per sempre
//   se i fetch successivi falliscono per giorni o settimane -- altrimenti un
//   prezzo vecchio di settimane risulterebbe presentato come "aggiornato".
// - BTP Futura non è in questa mappa: resta manuale, per decisione esplicita.

import fs from 'node:fs/promises';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import { toMilanSessionDate, daysBetween, deriveStaleStatus } from './price-utils.mjs';

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const PRICES_PATH = path.join(__dirname, '..', 'prices.json');

// ISIN -> ticker Yahoo (.MI). Stessa mappatura concettuale di FINECO_MAP nella
// dashboard, ma tenuta separata e minimale qui: questo script non deve dipendere
// dal codice della dashboard, solo condividerne gli ISIN come identità.
const TICKERS = {
  'IE00B4L5Y983': 'SWDA.MI',
  'IE00BTJRMP35': 'XMME.MI',
  'IE00B53QG562': 'CSEMU.MI',
  'LU1781541252': 'LCJP.MI',
  'LU0290358497': 'XEON.MI',
  'LU0290357176': 'X57E.MI',
  'LU1287023003': 'EM57.MI',
  'IE0006WW1TQ4': 'EXUS.MI',
  'FR0013416716': 'GOLD.MI',
};

async function loadExistingData() {
  try {
    const raw = await fs.readFile(PRICES_PATH, 'utf8');
    const json = JSON.parse(raw);
    return {
      entries: (json && typeof json.entries === 'object' && json.entries) || {},
      warnings: Array.isArray(json?.warnings) ? json.warnings : [],
    };
  } catch {
    return { entries: {}, warnings: [] }; // primo run, o file assente/corrotto: si riparte da zero
  }
}

export async function fetchOne(yahooTicker, todayMilan) {
  // range=5d copre weekend + un ponte di 2-3 giorni festivi; se anche questo non
  // basta (festività più lunghe), il titolo risulterà "stale" con più giorni di
  // scarto -- non un errore, solo un'informazione più vecchia.
  const url = `https://query1.finance.yahoo.com/v8/finance/chart/${encodeURIComponent(yahooTicker)}?range=5d&interval=1d`;
  const res = await fetch(url, {
    headers: {
      'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0 Safari/537.36',
      'Accept': 'application/json',
    },
  });
  if (!res.ok) throw new Error(`HTTP ${res.status}`);

  const json = await res.json();
  if (json?.chart?.error) throw new Error(`Yahoo error: ${JSON.stringify(json.chart.error)}`);

  const result = json?.chart?.result?.[0];
  if (!result) throw new Error('Nessun result nella risposta');

  const meta = result.meta || {};
  if (meta.exchangeName !== 'MIL') throw new Error(`exchangeName inatteso: ${meta.exchangeName}`);
  if (meta.fullExchangeName !== 'Milan') throw new Error(`fullExchangeName inatteso: ${meta.fullExchangeName}`);
  if (meta.currency !== 'EUR') throw new Error(`currency inattesa: ${meta.currency}`);
  if (meta.instrumentType !== 'ETF') throw new Error(`instrumentType inatteso: ${meta.instrumentType}`);

  const timestamps = result.timestamp || [];
  const closesRaw = result.indicators?.quote?.[0]?.close || [];
  if (!timestamps.length || !closesRaw.length) throw new Error('Serie storica vuota');

  // Verificato empiricamente su payload Yahoo reali (a mercato aperto): l'ultima
  // barra della serie, quando rappresenta la sessione odierna ancora in corso, non
  // è un close -- è uno specchio in tempo reale di meta.regularMarketPrice (stesso
  // valore, byte per byte). meta.marketState non è presente in questo endpoint e
  // non può essere usato come segnale. Il confronto affidabile, verificato sugli
  // stessi payload, è: currentTradingPeriod.regular.end (chiusura prevista per la
  // sessione corrente, calcolata da Yahoo sul calendario reale dell'exchange) contro
  // regularMarketTime (istante dell'ultimo scambio noto a Yahoo).
  const regularStart = result.meta?.currentTradingPeriod?.regular?.start;
  const regularEnd = result.meta?.currentTradingPeriod?.regular?.end;
  const regularMarketTime = result.meta?.regularMarketTime;
  if (!Number.isFinite(regularStart) || !Number.isFinite(regularEnd) || !Number.isFinite(regularMarketTime)) {
    // Caso limite esplicito: senza questi valori non possiamo stabilire se la
    // sessione odierna sia conclusa. Non si assume che lo sia -- comportamento
    // difensivo identico a un errore di validazione: nessun close scritto, valore
    // precedente preservato da chi chiama questa funzione.
    throw new Error('currentTradingPeriod.regular.start/end o regularMarketTime mancanti/non validi: impossibile stabilire se la sessione odierna è conclusa');
  }

  const closes = closesRaw.slice();
  const lastTs = timestamps[timestamps.length - 1];
  // L'osservazione empirica (barra odierna con timestamp coincidente all'apertura
  // della sessione corrente) NON va assunta valida a priori: va verificata caso per
  // caso. Se l'ultima barra della serie non appartiene alla finestra [regular.start,
  // regular.end] -- es. la serie non contiene ancora una barra per la sessione
  // odierna -- allora è già un close storico reale, e va lasciata intatta a
  // prescindere dal fatto che la sessione corrente sia ancora aperta.
  const lastBarIsCurrentSession = lastTs >= regularStart && lastTs <= regularEnd;
  if (lastBarIsCurrentSession && regularMarketTime < regularEnd) {
    // Sessione odierna non ancora conclusa E l'ultima barra è davvero quella di
    // oggi: non è un close definitivo. Trattata esattamente come un null -- la
    // scansione sotto, invariata, ripiega naturalmente sul close precedente reale
    // della serie.
    closes[closes.length - 1] = null;
  }

  // Ultimo indice con un close non-null e finito -- Yahoo può restituire null per
  // sessioni senza scambi in strumenti poco liquidi (oltre che per il caso appena
  // gestito sopra).
  let idx = -1;
  for (let i = closes.length - 1; i >= 0; i--) {
    if (closes[i] !== null && Number.isFinite(closes[i])) { idx = i; break; }
  }
  if (idx === -1) throw new Error('Nessun close valido negli ultimi 5 giorni');

  const close = closes[idx];
  if (!(close > 0)) throw new Error(`Prezzo non valido: ${close}`);

  // Data di sessione derivata dal timestamp Yahoo del close scelto, convertito
  // in Europe/Rome (regole IANA, CET/CEST gestiti automaticamente) -- non dal
  // gmtoffset grezzo restituito da Yahoo né dal fuso orario del runner.
  const sessionDate = toMilanSessionDate(timestamps[idx]);
  const status = deriveStaleStatus(sessionDate, todayMilan);

  const entry = {
    ticker: yahooTicker,
    close,
    currency: meta.currency,
    exchange: meta.exchangeName,
    session_date: sessionDate,
    fetched_at: new Date().toISOString(),
    source: 'yahoo_chart',
    status,
  };
  if (status === 'stale') {
    entry.note = `ultimo close disponibile: ${sessionDate} (${daysBetween(sessionDate, todayMilan)} giorni fa)`;
  }
  return entry;
}

// Rivaluta un'entry preesistente quando il fetch di oggi per quell'ISIN è
// fallito: NON introduce un prezzo nuovo (close/session_date restano quelli
// dell'ultimo dato realmente valido), ma lo status va ricalcolato in base
// all'età reale rispetto a oggi. Uno status può solo peggiorare qui (ok -> stale
// se l'età supera la soglia), mai migliorare: "tornare ok" richiede un fetch
// riuscito con un dato nuovo, non il semplice passare del tempo.
function reassessAge(prevEntry, todayMilan) {
  if (!prevEntry) return prevEntry; // nessuna entry precedente: resta assente
  const recomputedStatus = deriveStaleStatus(prevEntry.session_date, todayMilan);
  const effectiveStatus = recomputedStatus === 'stale' ? 'stale' : prevEntry.status;
  const ageDays = daysBetween(prevEntry.session_date, todayMilan);
  if (effectiveStatus === prevEntry.status && effectiveStatus !== 'stale') {
    return prevEntry; // nessun cambiamento: non toccare nulla, nemmeno il testo del note
  }
  const next = { ...prevEntry, status: effectiveStatus };
  if (effectiveStatus === 'stale') {
    // Il testo dei "giorni fa" va aggiornato ad ogni run, anche se lo status
    // era già 'stale' in precedenza: altrimenti il conteggio resterebbe fermo
    // al valore del giorno in cui è diventato stale per la prima volta.
    next.note = `ultimo close disponibile: ${prevEntry.session_date} (${ageDays} giorni fa)`;
  }
  return next;
}

// Proiezione "sostanziale" di un'entry: solo i campi che rappresentano
// informazione finanziaria/di stato reale, esclusi i timestamp tecnici
// (fetched_at) che cambiano ad ogni fetch riuscito indipendentemente dal fatto
// che il prezzo sia lo stesso. Usata per decidere se prices.json è cambiato in
// modo che valga la pena di un commit.
function substantiveProjection(entries) {
  const out = {};
  for (const isin of Object.keys(entries).sort()) {
    const e = entries[isin];
    out[isin] = {
      ticker: e.ticker ?? null,
      close: e.close ?? null,
      currency: e.currency ?? null,
      exchange: e.exchange ?? null,
      session_date: e.session_date ?? null,
      status: e.status ?? null,
      note: e.note ?? null,
    };
  }
  return out;
}

async function main() {
  const { entries, warnings: previousWarnings } = await loadExistingData();
  const previousSubstantive = JSON.stringify(substantiveProjection(entries));

  const todayMilan = toMilanSessionDate(Math.floor(Date.now() / 1000));
  const warnings = [];
  let successCount = 0;

  for (const [isin, yahooTicker] of Object.entries(TICKERS)) {
    try {
      entries[isin] = await fetchOne(yahooTicker, todayMilan);
      successCount++;
    } catch (err) {
      warnings.push(`${yahooTicker} (${isin}): ${err.message}`);
      // Nessuna scrittura di un prezzo nuovo: il valore precedente (close,
      // session_date) resta esattamente com'era. Lo status, però, va comunque
      // rivalutato rispetto a oggi -- vedi reassessAge(). Se non esisteva alcuna
      // entry precedente, la chiave NON va creata (nemmeno con valore undefined):
      // "mai avuto un dato valido" deve restare "chiave assente", non una entry
      // fantasma che romperebbe la proiezione sostanziale e l'assunzione della
      // dashboard che `!entry` significhi "nessun dato".
      if (entries[isin]) {
        entries[isin] = reassessAge(entries[isin], todayMilan);
      }
    }
  }

  if (successCount === 0) {
    // Tutti e 7 i ticker sono falliti nello stesso run: molto più probabile un
    // problema di infrastruttura (Yahoo down, formato cambiato, rete) che un
    // giorno di mercato chiuso -- in quel caso normalmente falliscono tutti
    // "stale" insieme, non con un errore. Fai fallire il job (segnale visibile
    // in GitHub Actions) senza comunque toccare prices.json.
    console.error('Tutti i ticker sono falliti in questo run:');
    warnings.forEach(w => console.error(' - ' + w));
    process.exit(1);
  }

  const output = {
    generated_at: new Date().toISOString(),
    entries,
    warnings,
  };

  await fs.writeFile(PRICES_PATH, JSON.stringify(output, null, 2) + '\n', 'utf8');

  // Cambiamento sostanziale = differisce la proiezione sostanziale delle entry
  // (ISIN, ticker, close, currency, exchange, session_date, status, note) e/o
  // l'elenco dei warning. generated_at e fetched_at NON entrano in questo
  // confronto: cambiano ad ogni run per definizione e da soli non giustificano
  // un commit.
  const newSubstantive = JSON.stringify(substantiveProjection(entries));
  const warningsChanged = JSON.stringify(previousWarnings) !== JSON.stringify(warnings);
  const hasSubstantiveChange = newSubstantive !== previousSubstantive || warningsChanged;

  if (process.env.GITHUB_OUTPUT) {
    await fs.appendFile(process.env.GITHUB_OUTPUT, `substantive_change=${hasSubstantiveChange}\n`, 'utf8');
  }

  console.log(`prices.json aggiornato: ${successCount}/${Object.keys(TICKERS).length} ticker OK.`);
  console.log(`Cambiamento sostanziale: ${hasSubstantiveChange ? 'sì' : 'no'}.`);
  if (warnings.length) {
    console.log(`${warnings.length} avviso/i:`);
    warnings.forEach(w => console.log(' - ' + w));
  }
}

main().catch(err => {
  console.error('Errore fatale:', err);
  process.exit(1);
});
