import { readFileSync, writeFileSync } from 'node:fs';
import { createRequire } from 'node:module';
globalThis.JSZip = createRequire(import.meta.url)('../vendor/jszip.min.js');
const { loadWorkbook, exportWorkbook } = await import('../js/xlsx.js');
const buf = readFileSync(process.argv[2]);
const m = await loadWorkbook(buf);
const w = m.workers[1];
const sp = m.spare[0];
const changes = new Map([
  [w.row, new Map([[m.dayCol['2026-09-21'], 12], [m.dayCol['2026-09-22'], 10.5], [4, '7 Day Sub']])],
  [w.row + w.off.ML, new Map([[m.dayCol['2026-09-21'], 100]])],
  [sp.row, new Map([[1, 'V999'], [2, 'Test Person & Co'], [3, 'KEQ4 Caster'], [4, 'Days Worked'], [5, 'SW 1099'], [m.dayCol['2026-09-21'], 10]])],
  [w.row + 2, new Map([[m.dayCol['2026-09-21'], 5]])], // OT es fórmula: debe saltarse
]);
const { blob, skipped } = await exportWorkbook(buf, m.sheetName, changes);
writeFileSync('test/out.xlsx', Buffer.from(await blob.arrayBuffer()));
const m2 = await loadWorkbook(readFileSync('test/out.xlsx'));
const w2 = m2.workers.find(x => x.row === w.row);
const n2 = m2.workers.find(x => x.row === sp.row);
console.log('skipped', skipped, 'aaron', w2.hours['2026-09-21'], w2.hours['2026-09-22'], w2.perDiem, w2.ml['2026-09-21'], 'new', n2 && [n2.id, n2.name, n2.payClass, n2.hours['2026-09-21']], 'workers', m2.workers.length);
