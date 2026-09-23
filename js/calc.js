// Motor de cálculo: aplica, celda por celda, la regla que se leyó de la fórmula del
// Excel (ver rules.js). Si una celda no tiene regla (vacía o escrita a mano) usa
// el valor fijo de `over`, igual que Excel.
import { DEFAULT_RULES } from './rules.js';

export const PAY_CLASSES = [
  'Southern Welding W2',
  'Elite Industrial Mechanical',
  'Elite Industrial Refractory',
  'Elite Refractory BM',
  'Elite Refractory Laborer',
  'SW 1099',
  'Elite 1099',
  'Gunite',
];
export const PER_DIEMS = ['Days Worked', '7 Day Sub'];

export const isNum = (v) => typeof v === 'number' && Number.isFinite(v);
const n = (v) => (isNum(v) ? v : 0);

// ISO 'YYYY-MM-DD' -> 1 (lunes) .. 7 (domingo), igual que WEEKDAY(x,2)
export function dow(iso) {
  const d = new Date(iso + 'T00:00:00Z').getUTCDay();
  return d === 0 ? 7 : d;
}

export function addDays(iso, k) {
  const d = new Date(iso + 'T00:00:00Z');
  d.setUTCDate(d.getUTCDate() + k);
  return d.toISOString().slice(0, 10);
}

export function weekDates(startIso) {
  return Array.from({ length: 7 }, (_, i) => addDays(startIso, i));
}

const same = (a, b) => String(a || '').trim().toLowerCase() === String(b || '').trim().toLowerCase();
const inList = (list, pc) => list.some((x) => same(x, pc));

function regFor(rule, pc, h, d, prior) {
  if (inList(rule.w40, pc)) return Math.max(Math.min(h, 40 - prior), 0);
  if (inList(rule.d8wk, pc)) return d >= 6 ? 0 : Math.min(h, 8);
  if (inList(rule.m8, pc)) return Math.min(h, 8);
  return 0;
}

// ctx = { rules: [...], thresholds: {proyecto: horas} }
export function computeWeek(w, startIso, ctx = {}) {
  const dates = weekDates(startIso);
  const pc = w.payClass || '';
  const hrs = dates.map((iso) => w.hours[iso]);
  const tab = ctx.rules || [];
  const fallback = w.ruleDefault || null; // ids para trabajadores nuevos
  const ruleOf = (iso, i) => {
    const id = (w.rule?.[iso] || fallback)?.[i];
    if (id == null) return [DEFAULT_RULES.reg, DEFAULT_RULES.ot, DEFAULT_RULES.dt, DEFAULT_RULES.pd][i];
    return id >= 0 ? tab[id] : null;
  };
  const weekSum = (start) => weekDates(start).reduce((a, iso) => a + n(w.hours[iso]), 0);
  const threshold = () => {
    const key = Object.keys(ctx.thresholds || {}).find((p) => same(p, w.project));
    const t = key ? ctx.thresholds[key] : 0;
    return !t ? 999999 : t;
  };

  let prior = 0;
  const days = dates.map((iso, i) => {
    const h = hrs[i];
    const d = dow(iso);
    const ov = w.over || {};
    const has = (k) => ov[k] && Object.prototype.hasOwnProperty.call(ov[k], iso);
    const [rReg, rOt, rDt, rPd] = [0, 1, 2, 3].map((k) => ruleOf(iso, k));
    let reg = null, ot = null, dt = null, pd = null;

    if (has('reg')) reg = ov.reg[iso];
    else if (rReg && isNum(h)) {
      const hol = rReg.holIso ? n(w.hours[rReg.holIso]) : 0;
      reg = regFor(rReg, pc, h, d, prior - hol);
    }
    if (has('dt')) dt = ov.dt[iso];
    else if (rDt && isNum(h)) dt = inList(rDt.pcs, pc) && d === 7 ? h : 0;
    if (has('ot')) ot = ov.ot[iso];
    else if (rOt && isNum(h)) ot = h - n(reg) - n(dt);
    prior += n(h);

    if (has('pd')) pd = ov.pd[iso];
    else if (rPd?.v === 'qd') {
      const qd = hrs.filter((x) => isNum(x) && x >= rPd.crit).length;
      const mind = rPd.word && String(w.project || '').toLowerCase().includes(rPd.word.toLowerCase()) ? rPd.minIf : rPd.minElse;
      if (!w.perDiem) pd = null;
      else if (same(w.perDiem, 'Days Worked')) pd = n(h) >= rPd.dwMin ? 1 : 0;
      else if (same(w.perDiem, '7 Day Sub')) pd = qd >= mind ? 1 : n(h) >= rPd.subMin ? 1 : 0;
      else pd = 0;
    } else if (rPd?.v === 'thr') {
      if (same(w.perDiem, 'Days Worked')) pd = n(h) > 0 ? 1 : 0;
      else if (same(w.perDiem, '7 Day Sub')) pd = weekSum(rPd.whWeek || startIso) >= threshold() ? 1 : n(h) > 0 ? 1 : 0;
      else pd = 0;
    }
    const manual = ['reg', 'ot', 'dt', 'pd'].filter(has);
    return { iso, h: isNum(h) ? h : null, reg, ot, dt, pd, ml: w.ml?.[iso] ?? null, tt: w.tt?.[iso] ?? null, manual };
  });
  const sum = (k) => days.reduce((a, x) => a + n(x[k]), 0);
  const t = { h: sum('h'), reg: sum('reg'), ot: sum('ot'), dt: sum('dt'), pd: sum('pd'), ml: sum('ml'), tt: sum('tt') };
  const r = rates(w);
  const cost = {
    reg: t.reg * r.reg,
    ot: t.ot * r.ot,
    dt: t.dt * r.dt,
    pd: t.pd * r.pd,
    ml: t.ml * r.ml,
    tt: t.tt * r.tt,
  };
  cost.total = cost.reg + cost.ot + cost.dt + cost.pd + cost.ml + cost.tt;
  return { days, t, cost };
}

// OT/DT salen de la tarifa base (=H8*1.5, =H8*2) salvo que el archivo traiga un valor fijo.
export function rates(w) {
  const reg = n(w.rate);
  return {
    reg,
    ot: isNum(w.otRateFixed) ? w.otRateFixed : reg * 1.5,
    dt: isNum(w.dtRateFixed) ? w.dtRateFixed : reg * 2,
    pd: n(w.pdRate),
    ml: n(w.mlRate),
    tt: n(w.ttRate),
  };
}
