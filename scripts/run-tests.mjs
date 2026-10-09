// Esegue tutta la suite: controllo di sintassi dello script inline di index.html (node --check)
// e ogni scripts/test-*.mjs, ciascuno in un processo separato. Riepilogo per file con il numero
// di casi superati (✓); exit code 1 se un controllo fallisce.
// Uso: npm test   (oppure: node scripts/run-tests.mjs)

import { spawnSync } from 'node:child_process';
import fs from 'node:fs/promises';
import os from 'node:os';
import path from 'node:path';
import { fileURLToPath } from 'node:url';

const ROOT = path.join(path.dirname(fileURLToPath(import.meta.url)), '..');
const SCRIPTS = path.join(ROOT, 'scripts');
const node = args => spawnSync(process.execPath, args, { cwd: ROOT, encoding: 'utf8' });
let falliti = 0;

// 1. Sintassi dello script inline di index.html (estratto in un file temporaneo)
const html = await fs.readFile(path.join(ROOT, 'index.html'), 'utf8');
const inline = [...html.matchAll(/<script>([\s\S]*?)<\/script>/g)].map(m => m[1]).join('\n');
const tmp = await fs.mkdtemp(path.join(os.tmpdir(), 'run-tests-'));
const file = path.join(tmp, 'inline.js');
await fs.writeFile(file, inline);
const check = node(['--check', file]);
await fs.rm(tmp, { recursive: true, force: true });
if (check.status === 0) console.log('node --check index.html: OK');
else { falliti++; console.log('node --check index.html: ERRORE\n' + check.stderr); }

// 2. Tutti i test
const test = (await fs.readdir(SCRIPTS)).filter(f => /^test-.*\.mjs$/.test(f)).sort();
let casi = 0;
for (const t of test) {
  const r = node([path.join(SCRIPTS, t)]);
  const out = (r.stdout || '') + (r.stderr || '');
  const n = (out.match(/✓/g) || []).length;
  casi += n;
  console.log(`${t.padEnd(40)} ${r.status === 0 ? 'ok    ' : 'FALLITO'} (${n} ✓)`);
  if (r.status !== 0) {
    falliti++;
    console.log(out.split('\n').filter(l => /Error|assert/i.test(l)).slice(0, 3).map(l => '    ' + l.trim()).join('\n'));
  }
}

console.log(`\n${test.length} file, ${casi} casi superati${falliti ? `, ${falliti} FALLITI` : ', tutto verde'}.`);
process.exit(falliti ? 1 : 0);
