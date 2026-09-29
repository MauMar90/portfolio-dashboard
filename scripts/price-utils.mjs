// Funzioni pure, senza side-effect, condivise tra update-prices.js e i relativi
// test. Isolate qui esclusivamente perché update-prices.js esegue main() in
// automatico al top-level (nessun export, nessun guard): importare direttamente
// da lì per testare la logica di data/soglia innescherebbe un fetch reale ad ogni
// import. Nessuna logica finanziaria/fiscale qui dentro -- solo calendario.

// Converte un timestamp Unix (secondi) nella data (YYYY-MM-DD) del giorno civile
// a Milano, usando le regole IANA della timezone Europe/Rome (gestisce CET/CEST
// e il cambio ora legale automaticamente). Non dipende in alcun modo dal fuso
// orario della macchina che esegue il codice (es. runner GitHub Actions in UTC).
export function toMilanSessionDate(tsSeconds) {
  const dt = new Date(tsSeconds * 1000);
  const parts = new Intl.DateTimeFormat('en-US', {
    timeZone: 'Europe/Rome',
    year: 'numeric',
    month: '2-digit',
    day: '2-digit',
  }).formatToParts(dt);
  const y = parts.find(p => p.type === 'year').value;
  const m = parts.find(p => p.type === 'month').value;
  const d = parts.find(p => p.type === 'day').value;
  return `${y}-${m}-${d}`;
}

// Differenza in giorni di calendario tra due date "YYYY-MM-DD" (entrambe intese
// come mezzanotte UTC pura: qui contano solo i giorni di calendario, non i fusi --
// la conversione al fuso corretto è già avvenuta a monte in toMilanSessionDate).
export function daysBetween(dateStrEarlier, dateStrLater) {
  const a = new Date(dateStrEarlier + 'T00:00:00Z');
  const b = new Date(dateStrLater + 'T00:00:00Z');
  return Math.round((b - a) / 86400000);
}

// Soglia di validità (giorni di calendario): oltre questa soglia un close --
// per quanto formalmente valido -- è considerato 'stale', non applicabile
// automaticamente a positions[].mkt. Decisione esplicita e definitiva.
export const STALE_DAYS = 5;

// Stato derivato dalla sola età del dato rispetto a "oggi" (data di sessione
// Milano). Pura funzione della soglia: usata sia al momento del fetch riuscito
// sia per rivalutare un'entry preesistente quando il fetch di oggi fallisce.
export function deriveStaleStatus(sessionDate, todayMilanDate, staleDays = STALE_DAYS) {
  const ageDays = daysBetween(sessionDate, todayMilanDate);
  return ageDays > staleDays ? 'stale' : 'ok';
}
