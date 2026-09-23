import { readFileSync } from 'node:fs';
import { createRequire } from 'node:module';
const require = createRequire(import.meta.url);
globalThis.JSZip = require('../vendor/jszip.min.js');
const { loadWorkbook } = await import('../js/xlsx.js');
const { computeWeek } = await import('../js/calc.js');
const file = process.argv[2];
const m = await loadWorkbook(readFileSync(file));
console.log(m.sheetName, m.company, m.location, m.job, m.weeks, 'workers', m.workers.length, 'spare', m.spare.length);
console.log('lists', JSON.stringify(m.lists));
const ctx = { rules: m.rules, thresholds: Object.fromEntries(m.lists.projects.map(p => [p.name, p.threshold])) };
let checked = 0, bad = 0;
for (const w of m.workers) {
  for (const start of m.weeks) {
    const res = computeWeek(w, start, ctx);
    for (const d of res.days) {
      for (const k of ['reg','ot','dt','pd']) {
        const exp = w.cached[k][d.iso] ?? null;
        const got = d[k];
        const e0 = exp ?? 0, g0 = got ?? 0;
        checked++;
        if (Math.abs(e0 - g0) > 1e-9) { bad++; if (bad < 25) console.log('MISMATCH', w.row, w.name, w.payClass, w.perDiem, d.iso, k, 'excel', exp, 'app', got, 'h', d.h); }
      }
    }
  }
}
console.log('checked', checked, 'mismatches', bad);
console.log('notes', m.workers.filter(w=>w.notes.length).length, 'hourText', m.workers.filter(w=>Object.keys(w.hourText).length).map(w=>[w.name,w.hourText]).slice(0,5));
console.log('rules', m.rules.length, 'unknown cells', m.workers.reduce((a,w)=>a+Object.values(w.unk).reduce((b,o)=>b+Object.keys(o).length,0),0));
