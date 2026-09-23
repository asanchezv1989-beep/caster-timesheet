import { readFileSync } from 'node:fs';
import { createRequire } from 'node:module';
globalThis.JSZip = createRequire(import.meta.url)('../vendor/jszip.min.js');
const x = await import('../js/xlsx.js');
const m = await x.loadWorkbook(readFileSync(process.argv[2]));
const c = {};
for (const w of m.workers) for (const k in w.unk) for (const iso in w.unk[k]) { const key = k + ' ' + iso.slice(5); c[key] = (c[key]||0)+1; }
console.log(Object.entries(c).sort((a,b)=>b[1]-a[1]).slice(0,12));
const w = m.workers.find(w => Object.keys(w.unk).length);
console.log(w.row, w.name, JSON.stringify(w.unk).slice(0,200));
