import { readFileSync } from 'node:fs';
import { createRequire } from 'node:module';
globalThis.JSZip = createRequire(import.meta.url)('../vendor/jszip.min.js');
const { loadWorkbook } = await import('../js/xlsx.js');
const { computeWeek } = await import('../js/calc.js');
const m = await loadWorkbook(readFileSync(process.argv[2]));
const ctx = { rules: m.rules, thresholds: Object.fromEntries(m.lists.projects.map(p => [p.name, p.threshold])) };
let tot = 0, h = 0, pd = 0, pdx = 0;
for (const w of m.workers) for (const s of m.weeks) { const r = computeWeek(w, s, ctx); tot += r.cost.total; h += r.t.h; pd += r.t.pd; }
for (const w of m.workers) for (const v of Object.values(w.cached.pd)) pdx += v;
console.log('total', tot.toFixed(2), 'hours', h, 'pd app', pd, 'pd excel', pdx);
