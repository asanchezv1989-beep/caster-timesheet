// Lee el texto de las fórmulas REG/OT/DT/PD de cada celda y lo convierte en una
// regla que el motor sabe aplicar. Cada archivo (y a veces cada semana) trae su
// propia variante: pay classes distintas, festivo restado, mínimo de días fijo,
// per diem por umbral del proyecto... Si la fórmula no se reconoce devuelve null
// y la celda se trata como valor fijo (el que guardó Excel).

import { colNum } from './colref.js';

function norm(f) {
  // quita prefijos de Excel y espacios fuera de comillas
  return f
    .replace(/_xlpm\.|_xlfn\.|_xludf\./g, '')
    .split('"')
    .map((s, i) => (i % 2 ? s : s.replace(/\s+/g, '')))
    .join('"')
    .replace(/^=/, '');
}

// Separa argumentos de nivel superior (respeta paréntesis y comillas)
function splitTop(s) {
  const out = [];
  let depth = 0, q = false, cur = '';
  for (const ch of s) {
    if (ch === '"') q = !q;
    if (!q) {
      if (ch === '(') depth++;
      if (ch === ')') depth--;
      if (ch === ',' && depth === 0) { out.push(cur); cur = ''; continue; }
    }
    cur += ch;
  }
  out.push(cur);
  return out;
}

// contenido entre el paréntesis que abre en `start` y su cierre
function inside(s, start) {
  let depth = 0, q = false;
  for (let i = start; i < s.length; i++) {
    const ch = s[i];
    if (ch === '"') q = !q;
    if (q) continue;
    if (ch === '(') depth++;
    if (ch === ')' && --depth === 0) return s.slice(start + 1, i);
  }
  return null;
}

const names = (s) => [...s.matchAll(/pc="([^"]*)"/g)].map((m) => m[1].trim());

const REG_RESULTS = {
  'MAX(MIN(h,40-prior),0)': 'w40',
  'IF(dow>=6,0,MIN(h,8))': 'd8wk',
  'MIN(h,8)': 'm8',
};

function parseReg(f) {
  const i = f.indexOf('SWITCH(TRUE,');
  if (i < 0 || !/prior,SUM\(/.test(f)) return null;
  const args = splitTop(inside(f, i + 'SWITCH'.length));
  args.shift(); // TRUE
  const def = args.length % 2 ? args.pop() : null;
  if (def !== '0') return null;
  const rule = { k: 'reg', w40: [], d8wk: [], m8: [] };
  for (let j = 0; j < args.length; j += 2) {
    const type = REG_RESULTS[args[j + 1]];
    const pcs = names(args[j]);
    if (!type || !pcs.length) return null;
    rule[type].push(...pcs);
  }
  // festivo: hol,N($BB12) -> horas de ese día no cuentan para las 40
  const hol = f.match(/hol,N\(\$?([A-Z]+)\$?\d+\)/);
  if (hol) {
    if (!/-h-hol/.test(f)) return null;
    rule.holCol = colNum(hol[1]);
  } else if (!/prior,SUM\([^)]*\)-h,/.test(f)) return null;
  return rule;
}

function parseDt(f) {
  const m = f.match(/IF\(AND\((.*),dow=7\),h,0\)/);
  if (!m || !/IF\(h="","",/.test(f)) return null;
  return { k: 'dt', pcs: names(m[1]) };
}

function parseOt(f) {
  return /^IF\([A-Z]+\d+="","",[A-Z]+\d+-[A-Z]+\d+-[A-Z]+\d+\)$/.test(f) ? { k: 'ot' } : null;
}

function parsePd(f) {
  if (/pd_type/.test(f)) {
    const wh = f.match(/week_hours,\$?([A-Z]+)\$?\d+/);
    const ok = /"DaysWorked"|"Days Worked"/.test(f) && /IF\(hours>0,1,0\)/.test(f) && /week_hours>=threshold/.test(f) && /IF\(_t=0,999999,_t\)/.test(f);
    if (!wh || !ok) return null;
    return { k: 'pd', v: 'thr', whCol: colNum(wh[1]) };
  }
  if (/COUNTIF\(/.test(f) && /qd>=(mind|\d+)/.test(f)) {
    const crit = f.match(/COUNTIF\([^,]*,">=(\d+(?:\.\d+)?)"\)/);
    const m1 = f.match(/mind,IF\(ISNUMBER\(SEARCH\("([^"]+)",proj\)\),(\d+),(\d+)\),/);
    // mínimo fijo: "mind,5," o directamente "qd>=5"
    const m2 = f.match(/mind,(\d+),/) || f.match(/qd>=(\d+),/);
    const dw = f.match(/"Days Worked",IF\(N\(h\)>=(\d+),1,0\)/);
    const sub = f.match(/"7 Day Sub",IF\(qd>=(?:mind|\d+),1,IF\(N\(h\)>=(\d+),1,0\)\)/);
    if (!crit || !(m1 || m2) || !dw || !sub || !/IF\(pd="","",/.test(f)) return null;
    return {
      k: 'pd', v: 'qd', crit: +crit[1], dwMin: +dw[1], subMin: +sub[1],
      word: m1 ? m1[1] : null, minIf: m1 ? +m1[2] : +m2[1], minElse: m1 ? +m1[3] : +m2[1],
    };
  }
  return null;
}

const PARSERS = { REG: parseReg, OT: parseOt, DT: parseDt, PD: parsePd };

export function parseRule(label, formula) {
  if (!formula) return null;
  try {
    return PARSERS[label](norm(formula));
  } catch {
    return null;
  }
}

// Reglas por defecto (plantilla CASTER), para bloques sin fórmulas legibles
export const DEFAULT_RULES = {
  reg: { k: 'reg', w40: ['Southern Welding W2', 'Elite Industrial Mechanical'], d8wk: ['Elite Industrial Refractory', 'Elite Refractory BM', 'Elite Refractory Laborer', 'Gunite'], m8: ['SW 1099'] },
  ot: { k: 'ot' },
  dt: { k: 'dt', pcs: ['Elite Industrial Refractory', 'Elite Refractory BM'] },
  pd: { k: 'pd', v: 'qd', crit: 8, dwMin: 8, subMin: 8, word: 'caster', minIf: 5, minElse: 6 },
};
