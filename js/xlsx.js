// Lectura y escritura del .xlsx sin tocar formato ni fórmulas: se leen los XML
// con expresiones regulares y al exportar solo se reescriben las celdas cambiadas.

import { colNum, colName } from './colref.js';
import { parseRule } from './rules.js';

export { colNum, colName };
const Zip = () => globalThis.JSZip;

const unesc = (s) =>
  s.replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/&quot;/g, '"').replace(/&apos;/g, "'")
    .replace(/&#(\d+);/g, (_, d) => String.fromCharCode(+d)).replace(/&#x([0-9a-f]+);/gi, (_, h) => String.fromCharCode(parseInt(h, 16)))
    .replace(/&amp;/g, '&');
const esc = (s) => String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');

const attr = (tag, name) => {
  const m = tag.match(new RegExp('\\s' + name + '="([^"]*)"'));
  return m ? m[1] : null;
};
const textOf = (xml) => {
  let out = '';
  for (const m of xml.matchAll(/<t(?:\s[^>]*)?>([\s\S]*?)<\/t>/g)) out += m[1];
  return unesc(out);
};

export function serialToIso(serial) {
  const ms = Date.UTC(1899, 11, 30) + Math.round(serial) * 86400000;
  return new Date(ms).toISOString().slice(0, 10);
}

function resolvePath(base, target) {
  if (target.startsWith('/')) return target.slice(1);
  const parts = base.split('/');
  parts.pop();
  for (const p of target.split('/')) {
    if (p === '..') parts.pop();
    else if (p !== '.') parts.push(p);
  }
  return parts.join('/');
}

async function readRels(zip, path) {
  const relPath = path.replace(/([^/]+)$/, '_rels/$1.rels');
  const f = zip.file(relPath);
  if (!f) return [];
  const xml = await f.async('string');
  return [...xml.matchAll(/<Relationship\b[^>]*>/g)].map((m) => ({
    id: attr(m[0], 'Id'),
    type: attr(m[0], 'Type') || '',
    target: resolvePath(path, attr(m[0], 'Target')),
  }));
}

async function sheetPaths(zip) {
  const wb = await zip.file('xl/workbook.xml').async('string');
  const rels = await readRels(zip, 'xl/workbook.xml');
  return [...wb.matchAll(/<sheet\b[^>]*>/g)].map((m) => {
    const rid = attr(m[0], 'r:id');
    return { name: unesc(attr(m[0], 'name')), path: rels.find((r) => r.id === rid)?.target };
  });
}

async function sharedStrings(zip) {
  const f = zip.file('xl/sharedStrings.xml');
  if (!f) return [];
  const xml = await f.async('string');
  return [...xml.matchAll(/<si>([\s\S]*?)<\/si>/g)].map((m) => textOf(m[1]));
}

// Devuelve Map fila -> Map columna -> {v, f, ft}  (ft = texto de la fórmula)
function parseCells(xml, ss) {
  const rows = new Map();
  const shared = new Map(); // si -> texto de la fórmula maestra
  const pendingShared = [];
  for (const rm of xml.matchAll(/<row\b[^>]*?(?:\/>|>([\s\S]*?)<\/row>)/g)) {
    const r = +attr(rm[0], 'r');
    const cells = new Map();
    if (rm[1]) {
      for (const cm of rm[1].matchAll(/<c\b([^>]*?)(?:\/>|>([\s\S]*?)<\/c>)/g)) {
        const ref = attr(cm[0], 'r');
        const t = attr(cm[0], 't');
        const inner = cm[2] || '';
        const vm = inner.match(/<v>([\s\S]*?)<\/v>/);
        const f = /<f\b/.test(inner);
        let ft = null;
        if (f) {
          const fm = inner.match(/<f\b[^>]*?(?:\/>|>([\s\S]*?)<\/f>)/);
          const si = attr(fm[0], 'si');
          if (fm[1]) {
            ft = unesc(fm[1]);
            if (si != null && attr(fm[0], 't') === 'shared') shared.set(si, ft);
          } else if (si != null) pendingShared.push([r, ref, si]);
        }
        let v = null;
        if (t === 's' && vm) v = ss[+vm[1]];
        else if (t === 'inlineStr') v = textOf(inner);
        else if (t === 'str' || t === 'e') v = vm ? unesc(vm[1]) : null;
        else if (t === 'b') v = vm ? vm[1] === '1' : null;
        else if (vm) v = Number(vm[1]);
        cells.set(colNum(ref.replace(/\d+/g, '')), { v, f, ft });
      }
    }
    rows.set(r, cells);
  }
  // las fórmulas compartidas dependientes usan el texto de la maestra (sus reglas no dependen de referencias relativas)
  for (const [r, ref, si] of pendingShared) {
    const c = rows.get(r).get(colNum(ref.replace(/\d+/g, '')));
    c.ft = shared.get(si) ?? null;
  }
  return rows;
}

const clean = (v) => (typeof v === 'string' ? v.trim() : v);
const numOrNull = (v) => (typeof v === 'number' && Number.isFinite(v) ? v : null);

export async function loadWorkbook(buffer) {
  const zip = await Zip().loadAsync(buffer);
  const sheets = await sheetPaths(zip);
  const ts = sheets.find((s) => /^timesheet$/i.test(s.name)) || sheets.find((s) => /timesheet/i.test(s.name));
  if (!ts) throw new Error('No sheet named "Timesheet" in this file.');
  const ss = await sharedStrings(zip);
  const rows = parseCells(await zip.file(ts.path).async('string'), ss);
  const get = (r, c) => rows.get(r)?.get(c) || { v: null, f: false };
  const val = (r, c) => clean(get(r, c).v);

  // Semanas: fecha de lunes en fila 5, columnas I, R, AA... (cada 9)
  const weekCols = [];
  const weeks = [];
  for (let c = colNum('I'); ; c += 9) {
    const v = get(5, c).v;
    if (typeof v !== 'number') break;
    weekCols.push(c);
    weeks.push(serialToIso(v));
  }
  if (!weeks.length) throw new Error('Could not find the week dates in row 5.');

  const dayCell = new Map(); // iso -> columna
  weeks.forEach((start, k) => {
    for (let i = 0; i < 7; i++) {
      const d = new Date(start + 'T00:00:00Z');
      d.setUTCDate(d.getUTCDate() + i);
      dayCell.set(d.toISOString().slice(0, 10), weekCols[k] + i);
    }
  });

  // Tabla de reglas (fórmulas REG/OT/DT/PD leídas y deduplicadas)
  const isoOfCol = new Map([...dayCell].map(([iso, c]) => [c, iso]));
  const rules = [];
  const ruleIdx = new Map();
  const ruleId = (desc) => {
    const k = JSON.stringify(desc);
    if (!ruleIdx.has(k)) { ruleIdx.set(k, rules.length); rules.push(desc); }
    return ruleIdx.get(k);
  };
  const resolve = (desc) => {
    if (!desc) return null;
    const d = { ...desc };
    if (d.holCol != null) { d.holIso = isoOfCol.get(d.holCol) ?? null; delete d.holCol; if (!d.holIso) return null; }
    if (d.whCol != null) { const k = weekCols.findIndex((c) => c + 7 === d.whCol); delete d.whCol; if (k < 0) return null; d.whWeek = weeks[k]; }
    return d;
  };
  const RULE_LABELS = ['REG', 'OT', 'DT', 'PD'];
  // por cada día: ids de regla; las celdas sin fórmula o con fórmula desconocida van a `over`
  const blockRules = (r, off) => {
    const rule = {};
    const over = { reg: {}, ot: {}, dt: {}, pd: {} };
    const unk = {};
    for (const [iso, c] of dayCell) {
      rule[iso] = RULE_LABELS.map((lab) => {
        if (off[lab] == null) return -1;
        const cell = get(r + off[lab], c);
        const key = lab.toLowerCase();
        if (!cell.f) { over[key][iso] = numOrNull(cell.v); return -1; }
        const desc = resolve(parseRule(lab, cell.ft));
        if (!desc) {
          over[key][iso] = numOrNull(cell.v);
          (unk[key] ||= {})[iso] = true;
          return -1;
        }
        return ruleId(desc);
      });
    }
    return { rule, over, unk };
  };

  const LABELS = ['REG', 'OT', 'DT', 'PD', 'ML', 'TT'];
  const maxRow = Math.max(...rows.keys());
  const workers = [];
  const spare = [];
  for (let r = 7; r < maxRow; r++) {
    if (val(r + 1, 7) !== 'REG') continue;
    const off = {};
    for (let k = 1; k <= 8; k++) {
      const lab = val(r + k, 7);
      if (LABELS.includes(lab) && off[lab] == null) off[lab] = k;
      if (k > 1 && val(r + k + 1, 7) === 'REG') break;
    }
    const name = val(r, 2);
    if (!name || typeof name !== 'string' || /management daily signature/i.test(name)) {
      // de un bloque vacío solo guardamos las reglas de un día típico (para trabajadores nuevos)
      const br = blockRules(r, off);
      const sample = Object.values(br.rule).find((ids) => ids.every((x) => x >= 0)) || null;
      spare.push({ row: r, off, rule: sample });
      continue;
    }
    const pick = (lab) => {
      const o = {};
      if (off[lab] == null) return o;
      for (const [iso, c] of dayCell) {
        const v = numOrNull(get(r + off[lab], c).v);
        if (v !== null) o[iso] = v;
      }
      return o;
    };
    const { rule, over, unk } = blockRules(r, off);
    const hours = {};
    const hourText = {};
    for (const [iso, c] of dayCell) {
      const v = get(r, c).v;
      if (typeof v === 'number') hours[iso] = v;
      else if (typeof v === 'string' && v.trim()) hourText[iso] = v.trim();
    }
    const rateCell = (lab) => get(r + (off[lab] ?? 99), 8);
    const otC = rateCell('OT');
    const dtC = rateCell('DT');
    workers.push({
      key: 'r' + r,
      row: r,
      off,
      id: val(r, 1) == null ? '' : String(val(r, 1)),
      name,
      project: val(r, 3) || '',
      perDiem: val(r, 4) || '',
      payClass: val(r, 5) || '',
      trade: val(r, 6) == null ? '' : String(val(r, 6)),
      rate: numOrNull(rateCell('REG').v),
      otRateFixed: otC.f ? null : numOrNull(otC.v),
      dtRateFixed: dtC.f ? null : numOrNull(dtC.v),
      pdRate: numOrNull(rateCell('PD').v),
      mlRate: numOrNull(rateCell('ML').v),
      ttRate: numOrNull(rateCell('TT').v),
      hours,
      hourText,
      ml: pick('ML'),
      tt: pick('TT'),
      // valores calculados que guardó Excel, para verificar el motor
      cached: { reg: pick('REG'), ot: pick('OT'), dt: pick('DT'), pd: pick('PD') },
      // reglas por día [reg, ot, dt, pd] -> índice en `rules`
      rule,
      // celdas escritas a mano encima de la fórmula (festivos, ajustes) o con fórmula que no se reconoce
      over,
      unk,
      notes: [],
    });
  }
  // Comentarios de celda
  const rels = await readRels(zip, ts.path);
  const cm = rels.find((r) => /\/comments$/.test(r.type));
  if (cm && zip.file(cm.target)) {
    const xml = await zip.file(cm.target).async('string');
    const authors = [...xml.matchAll(/<author>([\s\S]*?)<\/author>/g)].map((m) => unesc(m[1]));
    for (const m of xml.matchAll(/<comment\b([^>]*)>([\s\S]*?)<\/comment>/g)) {
      const ref = attr(m[0], 'ref');
      const row = +ref.replace(/[A-Z]+/g, '');
      const w = [...workers].reverse().find((x) => x.row <= row);
      if (w && row - w.row < 8) {
        w.notes.push({ ref, author: authors[+attr(m[0], 'authorId')] || '', text: textOf(m[2]).replace(/^[^:\n]*:\n/, '').trim() });
      }
    }
  }

  // Listas (hoja Lists)
  const lists = { payClasses: [], perDiems: [], projects: [] };
  const lsh = sheets.find((s) => /^lists$/i.test(s.name));
  if (lsh) {
    const lr = parseCells(await zip.file(lsh.path).async('string'), ss);
    const lv = (r, c) => clean(lr.get(r)?.get(c)?.v ?? null);
    for (let r = 2; r <= 60; r++) {
      if (lv(r, 2)) lists.payClasses.push(String(lv(r, 2)));
      if (lv(r, 8)) lists.perDiems.push(String(lv(r, 8)));
      if (lv(r, 10)) lists.projects.push({ name: String(lv(r, 10)), threshold: numOrNull(lv(r, 11)) });
    }
  }

  const colOf = Object.fromEntries(dayCell);
  return {
    sheetName: ts.name,
    company: val(1, 6) || '',
    location: val(2, 6) || '',
    job: val(3, 3) || '',
    weeks,
    dayCol: colOf,
    rules,
    workers,
    spare,
    lists,
  };
}

// ---------- Exportar ----------

function cellXml(ref, style, value) {
  const s = style ? ` s="${style}"` : '';
  if (value === null || value === undefined || value === '') return `<c r="${ref}"${s}/>`;
  if (typeof value === 'number') return `<c r="${ref}"${s}><v>${value}</v></c>`;
  return `<c r="${ref}"${s} t="inlineStr"><is><t xml:space="preserve">${esc(value)}</t></is></c>`;
}

// Aplica cambios {row -> {col -> value}} sobre el XML de la hoja.
export function patchSheet(xml, changes) {
  const skipped = [];
  const done = new Set();
  const patchRow = (rowXml, r) => {
    const want = changes.get(r);
    const open = rowXml.match(/^<row\b[^>]*?(\/?)>/);
    let head = open[0];
    let body = '';
    if (open[1] === '/') head = head.replace(/\s*\/>$/, '>');
    else body = rowXml.slice(open[0].length, rowXml.length - '</row>'.length);
    const cells = [...body.matchAll(/<c\b[^>]*?(?:\/>|>[\s\S]*?<\/c>)/g)].map((m) => {
      const ref = attr(m[0], 'r');
      return { col: colNum(ref.replace(/\d+/g, '')), xml: m[0] };
    });
    for (const [col, value] of want) {
      const ref = colName(col) + r;
      const i = cells.findIndex((c) => c.col === col);
      if (i >= 0) {
        if (/<f\b/.test(cells[i].xml)) { skipped.push(ref); continue; }
        cells[i].xml = cellXml(ref, attr(cells[i].xml, 's'), value);
      } else {
        const rowStyle = attr(head, 's');
        cells.push({ col, xml: cellXml(ref, rowStyle, value) });
      }
    }
    cells.sort((a, b) => a.col - b.col);
    done.add(r);
    return head + cells.map((c) => c.xml).join('') + '</row>';
  };
  let out = xml.replace(/<row\b[^>]*?(?:\/>|>[\s\S]*?<\/row>)/g, (m) => {
    const r = +attr(m, 'r');
    return changes.has(r) ? patchRow(m, r) : m;
  });
  // filas que no existían
  const missing = [...changes.keys()].filter((r) => !done.has(r)).sort((a, b) => a - b);
  for (const r of missing) {
    const rowXml = patchRow(`<row r="${r}">` + '</row>', r);
    let inserted = false;
    out = out.replace(/<row\b[^>]*?r="(\d+)"/g, (m, rr) => {
      if (!inserted && +rr > r) { inserted = true; return rowXml + m; }
      return m;
    });
    if (!inserted) out = out.replace('</sheetData>', rowXml + '</sheetData>');
  }
  return { xml: out, skipped };
}

export async function exportWorkbook(buffer, sheetName, changes) {
  const zip = await Zip().loadAsync(buffer);
  const sheets = await sheetPaths(zip);
  const ts = sheets.find((s) => s.name === sheetName);
  const xml = await zip.file(ts.path).async('string');
  const { xml: patched, skipped } = patchSheet(xml, changes);
  zip.file(ts.path, patched);
  // que Excel recalcule todo al abrir
  let wb = await zip.file('xl/workbook.xml').async('string');
  if (/<calcPr\b/.test(wb)) {
    wb = wb.replace(/<calcPr\b([^>]*?)(\/?)>/, (m, a, sl) => `<calcPr${a.replace(/\sfullCalcOnLoad="[^"]*"/, '')} fullCalcOnLoad="1"${sl}>`);
  } else {
    wb = wb.replace('</workbook>', '<calcPr fullCalcOnLoad="1"/></workbook>');
  }
  zip.file('xl/workbook.xml', wb);
  const blob = await zip.generateAsync({
    type: typeof Blob !== 'undefined' ? 'blob' : 'nodebuffer',
    mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    compression: 'DEFLATE',
  });
  return { blob, skipped };
}
