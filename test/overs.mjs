import { readFileSync } from 'node:fs';
import { createRequire } from 'node:module';
globalThis.JSZip = createRequire(import.meta.url)('../vendor/jszip.min.js');
const { loadWorkbook } = await import('../js/xlsx.js');
const { computeWeek } = await import('../js/calc.js');
const m = await loadWorkbook(readFileSync(process.argv[2]));
const byW = [];
let tot={reg:0,ot:0,dt:0,pd:0,ml:0,tt:0,total:0}, hrs={reg:0,ot:0,dt:0,pd:0};
for (const w of m.workers){ let c=0; for (const k in w.over) for (const [iso,v] of Object.entries(w.over[k])) c++; if(c) byW.push([w.row,w.name,c, Object.fromEntries(Object.entries(w.over).map(([k,o])=>[k,Object.entries(o).filter(([i,v])=>v!==null).length]))]);
 for (const s of m.weeks){const r=computeWeek(w,s); for(const k in tot) tot[k]+=r.cost[k]; for(const k in hrs) hrs[k]+=r.t[k];}}
console.log(byW.length, byW.slice(0,15));
console.log(hrs, tot);
