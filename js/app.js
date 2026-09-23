import { loadWorkbook, exportWorkbook, serialToIso } from './xlsx.js';
import { computeWeek, weekDates, addDays, dow, isNum, PAY_CLASSES, PER_DIEMS } from './calc.js';
import * as db from './db.js';
import { t, getLang, setLang } from './i18n.js';
import * as ms from './graph.js';
import { colName, colNum } from './colref.js';

// ---------- estado ----------
let S = null; // { meta, workers, spare, lists, exportedSig }
let FILE = null; // ArrayBuffer original
const ui = { tab: 'day', date: null, week: 0, q: '', dayFilter: 'all', weekOnlyHours: true, wFilter: 'all', sumWeek: -1, focusKey: null };

const $ = (s, el = document) => el.querySelector(s);
const view = $('#view');
const esc = (s) => String(s ?? '').replace(/[&<>"]/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' })[c]);
const money = (v) => (v || 0).toLocaleString('en-US', { style: 'currency', currency: 'USD', minimumFractionDigits: 2 });
const money0 = (v) => (v || 0).toLocaleString('en-US', { style: 'currency', currency: 'USD', maximumFractionDigits: 0 });
const num = (v) => (isNum(v) ? (Math.round(v * 100) / 100).toLocaleString('en-US') : '');
const blank = (v) => (isNum(v) && v !== 0 ? num(v) : '');
const WD = { en: ['Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat', 'Sun'], es: ['Lun', 'Mar', 'Mié', 'Jue', 'Vie', 'Sáb', 'Dom'] };
const wd = (iso) => WD[getLang()][dow(iso) - 1];
const mdy = (iso) => `${iso.slice(5, 7)}/${iso.slice(8, 10)}/${iso.slice(0, 4)}`;
const md = (iso) => `${+iso.slice(5, 7)}/${+iso.slice(8, 10)}`;
const todayIso = () => {
  const d = new Date();
  return `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}-${String(d.getDate()).padStart(2, '0')}`;
};
const parseH = (s) => {
  s = String(s).trim().replace(',', '.');
  if (s === '') return null;
  const v = Number(s);
  return Number.isFinite(v) && v >= 0 && v <= 24 ? v : undefined;
};
const shortPC = (pc) =>
  ({ 'Southern Welding W2': 'SW W2', 'Elite Industrial Mechanical': 'EI Mech', 'Elite Industrial Refractory': 'EI Refr', 'Elite Refractory BM': 'ER BM', 'Elite Refractory Laborer': 'ER Lab', 'SW 1099': 'SW 1099', Gunite: 'Gunite' })[pc] || pc;

function toast(msg, ms = 2600) {
  const el = $('#toast');
  el.textContent = msg;
  el.hidden = false;
  clearTimeout(toast.t);
  toast.t = setTimeout(() => (el.hidden = true), ms);
}

// ---------- persistencia ----------
// Varios archivos: 'index' = [{id, label, fileName, importedAt}], 'current' = id,
// 'state:<id>' = modelo editado, 'file:<id>' = xlsx original.
let FILES = [];
let CUR = null;
let saveT;
function save() {
  clearTimeout(saveT);
  const id = CUR, st = S;
  saveT = setTimeout(() => db.set('state:' + id, st).catch((e) => toast(e.message)), 250);
  updateHeader();
}

const snap = (w) =>
  JSON.parse(JSON.stringify({ id: w.id, name: w.name, project: w.project, perDiem: w.perDiem, payClass: w.payClass, trade: w.trade, rate: w.rate, pdRate: w.pdRate, hours: w.hours, ml: w.ml, tt: w.tt }));

// "REFRAC NIGHTS (3) (1).xlsx" -> "REFRAC NIGHTS": así una versión nueva reemplaza a la anterior
const baseName = (fn) => fn.replace(/\.xlsx$/i, '').replace(/(\s*\(\d+\))+$/, '').replace(/[_\s]+OUT$/i, (m) => m).trim();
const fileId = (fn) => baseName(fn).toUpperCase().replace(/[^A-Z0-9]+/g, '-');

const hasPending = (st) => {
  if (!st) return false;
  const prev = S; S = st;
  const n = changeList().length, pend = n && getSig() !== st.exportedSig;
  S = prev;
  return pend ? n : 0;
};

async function importFile(file) {
  return importBuffer(file.name, () => file.arrayBuffer(), null);
}

async function importBuffer(fileName, getBuf, remote, force = false) {
  const file = { name: fileName };
  const id = fileId(file.name);
  const existing = FILES.find((f) => f.id === id);
  if (existing) {
    const st = id === CUR ? S : await db.get('state:' + id);
    const n = hasPending(st);
    if (n && !force && !confirm(t('replaceQ', { n, f: existing.label }))) return;
  }
  view.innerHTML = `<div class="empty">${t('reading')}</div>`;
  try {
    const buf = await getBuf();
    const m = await loadWorkbook(buf);
    for (const w of m.workers) {
      delete w.cached;
      w.orig = snap(w);
    }
    const st = {
      meta: { fileName: file.name, remote: remote || null, importedAt: new Date().toISOString(), sheetName: m.sheetName, company: m.company, location: m.location, job: m.job, weeks: m.weeks, dayCol: m.dayCol },
      rules: m.rules,
      workers: m.workers,
      spare: m.spare,
      lists: m.lists,
      exportedSig: null,
    };
    // si reemplaza al archivo abierto, que switchFile no vuelva a guardar el estado viejo encima
    if (id === CUR) { clearTimeout(saveT); S = null; }
    await db.set('file:' + id, buf);
    await db.set('state:' + id, st);
    const entry = { id, label: baseName(file.name), job: m.job, fileName: file.name, importedAt: st.meta.importedAt, remote: !!remote };
    FILES = [...FILES.filter((f) => f.id !== id), entry].sort((x, y) => x.label.localeCompare(y.label));
    await db.set('index', FILES);
    await switchFile(id);
    ui.tab = 'day';
    render();
    if (!force) toast(`${entry.label}: ${m.workers.length} ${t('workers').toLowerCase()}`);
  } catch (e) {
    console.error(e);
    toast(t('badFile', { e: e.message }), 5000);
    render();
  }
}

async function switchFile(id) {
  clearTimeout(saveT);
  if (S && CUR) await db.set('state:' + CUR, S);
  CUR = id;
  S = (await db.get('state:' + id)) || null;
  FILE = (await db.get('file:' + id)) || null;
  if (!S || !FILE) { S = null; FILE = null; }
  await db.set('current', id);
  ui.q = '';
  if (S) initDate();
}

async function removeFile(id) {
  await Promise.all([db.del('state:' + id), db.del('file:' + id)]);
  FILES = FILES.filter((f) => f.id !== id);
  await db.set('index', FILES);
  if (id === CUR) {
    S = null; FILE = null; CUR = null;
    if (FILES[0]) await switchFile(FILES[0].id);
  }
  render();
}

function openFiles() {
  ui.modalKey = null;
  $('#modal-body').innerHTML = `
  <div class="m-head"><h2>${t('files')}</h2><button class="btn icon" data-act="mclose" aria-label="${t('close')}"><svg viewBox="0 0 24 24"><path d="M18 6 6 18M6 6l12 12"/></svg></button></div>
  <div class="m-body">
    <div class="card list">${FILES.map((f) => `
      <div class="item ${f.id === CUR ? 'has' : ''}">
        <div class="who" data-act="usefile" data-id="${f.id}">
          <div class="nm">${esc(f.label)} ${f.id === CUR ? `<span class="badge ok">${t('open')}</span>` : ''}</div>
          <div class="meta"><span>${esc(f.fileName)}</span><span>· ${new Date(f.importedAt).toLocaleDateString('en-US')}</span></div>
        </div>
        <button class="btn sm danger" data-act="rmfile" data-id="${f.id}" aria-label="${t('removeFile')}"><svg viewBox="0 0 24 24" style="width:18px;height:18px"><path d="M3 6h18M8 6V4h8v2M6 6l1 14h10l1-14"/></svg></button>
      </div>`).join('') || `<div class="empty">—</div>`}</div>
    <p class="muted small">${t('filesHint')}</p>
    <button class="btn primary" data-act="pick" style="width:100%;height:48px"><svg viewBox="0 0 24 24"><path d="M12 5v14M5 12h14"/></svg>${t('addFile')}</button>
  </div>`;
  if (!$('#modal').open) $('#modal').showModal();
}

// ---------- SharePoint / OneDrive ----------
const eqv = (a, b) => {
  const n = (v) => (v === '' || v === undefined || v === null ? null : typeof v === 'string' ? v.trim() : v);
  a = n(a); b = n(b);
  if (a === null || b === null) return a === b;
  return typeof a === 'number' || typeof b === 'number' ? Number(a) === Number(b) : String(a).toLowerCase() === String(b).toLowerCase();
};

async function spRun(fn, pending) {
  try {
    return await fn();
  } catch (e) {
    $('#toast').hidden = true;
    if (ms.isNeedLogin(e)) {
      if (confirm(t('spLoginQ'))) await ms.signIn(pending);
      return;
    }
    console.error(e);
    toast(t('spError', { e: e.message }), 7000);
  }
}

function pickRemote(mode, q = '') {
  return spRun(async () => {
    toast(t('spSearching'), 20000);
    ui.pick = { mode, q, list: await ms.findFiles(q) };
    $('#toast').hidden = true;
    renderPicker();
  }, { act: 'pick', mode });
}

function renderPicker() {
  const p = ui.pick;
  ui.modalKey = null;
  $('#modal-body').innerHTML = `
  <div class="m-head"><h2>${t(p.mode === 'link' ? 'spLinkTitle' : 'spOpenTitle')}</h2><button class="btn icon" data-act="mclose" aria-label="${t('close')}"><svg viewBox="0 0 24 24"><path d="M18 6 6 18M6 6l12 12"/></svg></button></div>
  <div class="m-body">
    <form class="row" data-form="spsearch" style="margin-bottom:10px">
      <div class="search grow"><svg viewBox="0 0 24 24"><circle cx="11" cy="11" r="7"/><path d="m20 20-3.5-3.5"/></svg>
      <input class="input" name="q" type="search" placeholder="${t('spSearchPh')}" value="${esc(p.q)}"></div>
      <button class="btn">${t('spSearch')}</button>
    </form>
    <p class="muted small" style="margin:0 0 8px">${p.q ? '' : t('spRecent')}</p>
    <div class="card list">${p.list.map((f, i) => `
      <div class="item"><div class="who" data-act="rpick" data-i="${i}">
        <div class="nm">${esc(f.name)}</div>
        <div class="meta"><span>${f.modified ? new Date(f.modified).toLocaleString('en-US') : ''}</span><span>${esc(f.folder)}</span></div>
      </div></div>`).join('') || `<div class="empty">${t('spNone')}</div>`}</div>
  </div>`;
  if (!$('#modal').open) $('#modal').showModal();
}

async function usePicked(i) {
  const f = ui.pick.list[i];
  const ref = { driveId: f.driveId, itemId: f.itemId, name: f.name, webUrl: f.webUrl };
  $('#modal').close();
  if (ui.pick.mode === 'link') {
    if (baseName(f.name).toUpperCase() !== baseName(S.meta.fileName).toUpperCase() && !confirm(t('spLinkOtherQ', { a: f.name, b: S.meta.fileName }))) return;
    S.meta.remote = ref;
    const e = FILES.find((x) => x.id === CUR);
    if (e) { e.remote = true; db.set('index', FILES); }
    save();
    render();
    toast(t('spLinked', { f: f.name }));
  } else {
    await spRun(() => importBuffer(f.name, () => ms.download(ref), ref));
  }
}

function refreshRemote() {
  const ref = S.meta.remote;
  return spRun(() => importBuffer(ref.name, () => ms.download(ref), ref), { act: 'refresh' });
}

// Escribe en el Excel en línea solo las celdas cambiadas. Antes relee la hoja en vivo:
// ubica a cada trabajador por nombre (por si movieron filas) y avisa si alguien más
// cambió una celda desde que se importó.
function sendRemote() {
  const ref = S?.meta.remote;
  if (!ref) return pickRemote('link');
  const ch = changeList();
  if (!ch.length) return toast(t('noChanges'));
  return spRun(async () => {
    toast(t('spReading'), 60000);
    const live = await ms.readSheet(ref, S.meta.sheetName);
    const c0 = colNum(live.startCol);
    const cell = (r, c) => {
      const rr = r - live.startRow, cc = c - c0;
      return { v: live.values[rr]?.[cc], f: live.formulas[rr]?.[cc] };
    };
    if (serialToIso(Number(cell(5, colNum('I')).v)) !== S.meta.weeks[0]) throw new Error(t('spWeeksDiffer'));

    const byName = new Map();
    const maxR = live.startRow + live.values.length;
    for (let r = 7; r < maxR; r++) {
      const nm = cell(r, 2).v;
      if (typeof nm === 'string' && nm.trim() && cell(r + 1, 7).v === 'REG') {
        const k = nm.trim().toLowerCase();
        byName.set(k, [...(byName.get(k) || []), r]);
      }
    }
    const rowFor = new Map();
    const lost = [];
    for (const w of S.workers) {
      if (!ch.some((c) => c.k === w.key)) continue;
      if (w.isNew) {
        const b = cell(w.row, 2).v;
        if ((b === '' || b == null) && cell(w.row + 1, 7).v === 'REG') rowFor.set(w.key, w.row);
        else lost.push(w.name || '—');
        continue;
      }
      const nm = String(w.orig.name).trim().toLowerCase();
      if (String(cell(w.row, 2).v ?? '').trim().toLowerCase() === nm) rowFor.set(w.key, w.row);
      else if ((byName.get(nm) || []).length === 1) rowFor.set(w.key, byName.get(nm)[0]);
      else lost.push(w.orig.name);
    }

    const cells = [], conflicts = [], skipped = [];
    let already = 0;
    for (const c of ch) {
      const base = rowFor.get(c.k);
      if (base == null) continue;
      const w = S.workers.find((x) => x.key === c.k);
      const r = base + (c.row - w.row);
      const addr = colName(c.col) + r;
      const lv = cell(r, c.col);
      if (typeof lv.f === 'string' && lv.f.startsWith('=')) { skipped.push(addr); continue; }
      if (eqv(lv.v, c.v)) { already++; continue; }
      if (!eqv(lv.v, c.from)) conflicts.push(`${c.who} · ${c.what}: SharePoint ${lv.v === '' || lv.v == null ? '∅' : lv.v} → ${c.v ?? '∅'}`);
      cells.push({ addr, v: c.v });
    }
    $('#toast').hidden = true;
    if (lost.length && !confirm(t('spLostQ', { n: lost.length, l: lost.slice(0, 8).join(', ') }))) return;
    if (conflicts.length && !confirm(t('spConflictQ', { n: conflicts.length, l: conflicts.slice(0, 8).join('\n') }))) return;
    if (!cells.length) {
      toast(t('spNothing', { n: already }));
      if (!lost.length && !skipped.length) await importBuffer(ref.name, () => ms.download(ref), ref, true);
      return;
    }
    if (!confirm(t('spSendQ', { n: cells.length, f: ref.name }))) return;
    const session = await ms.openSession(ref);
    let failed;
    try {
      failed = await ms.writeCells(ref, S.meta.sheetName, session, cells, (d, n) => toast(t('spWriting', { d, n }), 60000));
    } finally {
      await ms.closeSession(ref, session);
    }
    $('#toast').hidden = true;
    if (!failed.length && !lost.length && !skipped.length) {
      // todo quedó en SharePoint: el archivo en línea pasa a ser la nueva base
      const keep = { date: ui.date, tab: ui.tab };
      await importBuffer(ref.name, () => ms.download(ref), ref, true);
      Object.assign(ui, keep);
      ui.week = weekOf(ui.date);
      render();
      toast(t('spSent', { n: cells.length }), 4000);
    } else {
      alert(t('spPartial', { ok: cells.length - failed.length, bad: failed.length + lost.length + skipped.length, l: [...failed.map((f) => `${f.addr}: ${f.error}`), ...lost, ...skipped].slice(0, 10).join('\n') }));
      render();
    }
  }, { act: 'send' });
}

// ---------- cálculos ----------
const ctx = () => ({ rules: S.rules || [], thresholds: Object.fromEntries((S.lists.projects || []).map((p) => [p.name, p.threshold])) });
const cw = (w, start) => computeWeek(w, start, ctx());
const active = () => S.workers.filter((w) => !w.removed);
const allDates = () => S.meta.weeks.flatMap((s) => weekDates(s));
const weekOf = (iso) => S.meta.weeks.findIndex((s) => iso >= s && iso <= addDays(s, 6));

function initDate() {
  const dates = allDates();
  const td = todayIso();
  ui.date = td < dates[0] ? dates[0] : td > dates.at(-1) ? dates.at(-1) : td;
  ui.week = weekOf(ui.date);
}

function dayResult(w, iso) {
  const k = weekOf(iso);
  if (k < 0) return null;
  return cw(w, S.meta.weeks[k]).days.find((d) => d.iso === iso);
}

function missing(w) {
  const m = [];
  if (!w.perDiem) m.push(t('missingPD'));
  if (!w.payClass) m.push(t('missingPC'));
  return m;
}

// ---------- cambios para exportar ----------
function changeList() {
  if (!S) return [];
  const out = [];
  const colOf = S.meta.dayCol;
  for (const w of S.workers) {
    const o = w.orig;
    const now = w.removed ? { id: '', name: '', project: '', perDiem: '', payClass: '', trade: '', rate: o.rate, pdRate: o.pdRate, hours: {}, ml: {}, tt: {} } : w;
    const label = w.removed ? `${o.name} ✕` : w.name || o.name;
    const fields = [['id', 1], ['name', 2], ['project', 3], ['perDiem', 4], ['payClass', 5], ['trade', 6]];
    for (const [f, col] of fields) {
      if ((now[f] ?? '') !== (o[f] ?? '')) {
        let v = now[f] ?? '';
        if (f === 'id' && /^\d+$/.test(v)) v = Number(v);
        out.push({ k: w.key, who: label, what: t(f === 'id' ? 'id' : f), from: o[f], to: now[f], row: w.row, col, v });
      }
    }
    if ((now.rate ?? null) !== (o.rate ?? null) && w.off.REG != null)
      out.push({ k: w.key, who: label, what: t('rate'), from: o.rate, to: now.rate, row: w.row + w.off.REG, col: 8, v: now.rate });
    if ((now.pdRate ?? null) !== (o.pdRate ?? null) && w.off.PD != null)
      out.push({ k: w.key, who: label, what: t('pdRate'), from: o.pdRate, to: now.pdRate, row: w.row + w.off.PD, col: 8, v: now.pdRate });
    const dayDiff = (key, rowOff, lab) => {
      if (rowOff == null) return;
      const a = o[key] || {}, b = now[key] || {};
      for (const iso of new Set([...Object.keys(a), ...Object.keys(b)])) {
        if ((a[iso] ?? null) !== (b[iso] ?? null) && colOf[iso])
          out.push({ k: w.key, who: label, what: `${lab} ${wd(iso)} ${md(iso)}`, from: a[iso], to: b[iso], row: w.row + rowOff, col: colOf[iso], v: b[iso] ?? null });
      }
    };
    dayDiff('hours', 0, t('hours'));
    dayDiff('ml', w.off.ML, 'ML');
    dayDiff('tt', w.off.TT, 'TT');
  }
  return out;
}
const getSig = () => JSON.stringify(changeList().map((c) => [c.row, c.col, c.v]));

function updateHeader() {
  const cur = FILES.find((f) => f.id === CUR);
  $('#hdr-job').textContent = S ? `${cur?.label || S.meta.job || 'Novelis'}${FILES.length > 1 ? ' ▾' : ''}` : 'Novelis';
  $('#hdr-sub').textContent = S ? `${S.meta.job ? S.meta.job + ' · ' : ''}${S.meta.location} · ${mdy(S.meta.weeks[0])} – ${mdy(addDays(S.meta.weeks.at(-1), 6))}` : '';
  const n = S ? changeList().length : 0;
  const d = $('#hdr-dirty');
  d.hidden = !n || getSig() === S.exportedSig;
  d.textContent = t('unsaved', { n });
}

// ---------- render ----------
function render() {
  // conservar foco y scroll de tablas entre renders
  const a = document.activeElement;
  const fk = a?.dataset?.fk;
  const wrap = $('.tbl-wrap', view);
  const sc = wrap ? [wrap.scrollTop, wrap.scrollLeft] : null;

  document.querySelectorAll('[data-i18n]').forEach((el) => (el.textContent = t(el.dataset.i18n)));
  document.querySelectorAll('#tabs button').forEach((b) => {
    b.classList.toggle('on', b.dataset.tab === ui.tab);
    b.disabled = !S && b.dataset.tab !== 'file';
  });
  document.documentElement.lang = getLang();
  updateHeader();
  if (!S) {
    view.innerHTML = importView();
    bindDrop();
    return;
  }
  view.innerHTML = { day: dayView, week: weekView, workers: workersView, summary: summaryView, file: fileView }[ui.tab]();
  if (sc && $('.tbl-wrap', view)) [$('.tbl-wrap', view).scrollTop, $('.tbl-wrap', view).scrollLeft] = sc;
  if (fk) {
    const el = view.querySelector(`[data-fk="${CSS.escape(fk)}"]`);
    if (el) {
      el.focus({ preventScroll: true });
      if (el.type === 'search') el.setSelectionRange(el.value.length, el.value.length);
      else el.select?.();
    }
  }
  if (ui.tab === 'file') bindDrop();
  if (ui.tab !== 'day') $('#quickbar').hidden = true;
}

function importView() {
  return `
  <h2>${t('importTitle')}</h2>
  <p class="muted">${t('importText')}</p>
  <div class="drop" id="drop">
    <svg viewBox="0 0 24 24"><path d="M14 3H6a2 2 0 0 0-2 2v14a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V9z"/><path d="M14 3v6h6M12 11v6M9 14l3-3 3 3"/></svg>
    <p class="row" style="justify-content:center;flex-wrap:wrap"><button class="btn primary" data-act="pick">${t('choose')}</button>${ms.configured() ? `<button class="btn" data-act="spopen">${t('spOpenBtn')}</button>` : ''}</p>
    <p class="muted small">${t('dropHere')}</p>
  </div>
  ${ms.configured() ? '' : spSection()}
  <p class="muted small" style="margin-top:16px">${t('installHint')}</p>
  <p class="row"><button class="btn sm" data-act="lang">${t('lang')}</button></p>`;
}

function bindDrop() {
  const d = $('#drop');
  if (!d) return;
  d.ondragover = (e) => { e.preventDefault(); d.classList.add('over'); };
  d.ondragleave = () => d.classList.remove('over');
  d.ondrop = (e) => {
    e.preventDefault();
    d.classList.remove('over');
    const f = e.dataTransfer.files[0];
    if (f) importFile(f);
  };
}

const searchBox = (fk) => `
  <div class="search grow"><svg viewBox="0 0 24 24"><circle cx="11" cy="11" r="7"/><path d="m20 20-3.5-3.5"/></svg>
  <input class="input" type="search" placeholder="${t('search')}" value="${esc(ui.q)}" data-act="q" data-fk="${fk}"></div>`;

const matches = (w) => {
  const q = ui.q.trim().toLowerCase();
  return !q || [w.name, w.id, w.trade, w.payClass].some((x) => String(x || '').toLowerCase().includes(q));
};
const byName = (a, b) => a.name.localeCompare(b.name);

// ----- Día -----
function dayView() {
  const iso = ui.date;
  const prev = addDays(iso, -1);
  const dates = allDates();
  const list = active().sort(byName);
  let withH = 0, totH = 0, pd = 0;
  const rows = list.map((w) => ({ w, r: dayResult(w, iso) }));
  for (const { r } of rows) {
    if (isNum(r.h) && r.h > 0) { withH++; totH += r.h; }
    pd += r.pd || 0;
  }
  const f = ui.dayFilter;
  const shown = rows.filter(({ w, r }) => {
    if (!matches(w)) return false;
    if (f === 'in') return isNum(r.h);
    if (f === 'out') return !isNum(r.h);
    if (f === 'prev') return isNum(w.hours[prev]);
    return true;
  });
  const k = weekOf(iso);
  return `
  <div class="datebar">
    <button class="btn icon" data-act="day" data-d="-1" ${iso <= dates[0] ? 'disabled' : ''} aria-label="prev"><svg viewBox="0 0 24 24"><path d="m15 18-6-6 6-6"/></svg></button>
    <label class="datebtn"><b>${wd(iso)} ${mdy(iso)}</b><span>${t('week')} ${k + 1} · ${md(S.meta.weeks[k])} – ${md(addDays(S.meta.weeks[k], 6))}</span>
      <input type="date" value="${iso}" min="${dates[0]}" max="${dates.at(-1)}" data-act="date"></label>
    <button class="btn icon" data-act="day" data-d="1" ${iso >= dates.at(-1) ? 'disabled' : ''} aria-label="next"><svg viewBox="0 0 24 24"><path d="m9 18 6-6-6-6"/></svg></button>
  </div>
  <div class="stats">
    <div class="stat"><b>${withH}</b><span>${t('workersHrs')}</span></div>
    <div class="stat"><b>${num(totH) || 0}</b><span>${t('totalHrs')}</span></div>
    <div class="stat"><b>${pd}</b><span>${t('pdDays')}</span></div>
  </div>
  <div class="row" style="margin-bottom:8px">${searchBox('q-day')}
    <button class="btn" data-act="copyprev" title="${t('copyPrev')}"><svg viewBox="0 0 24 24"><rect x="8" y="8" width="13" height="13" rx="2"/><path d="M16 8V5a2 2 0 0 0-2-2H5a2 2 0 0 0-2 2v9a2 2 0 0 0 2 2h3"/></svg><span class="hide-sm">${t('copyPrev')}</span></button>
  </div>
  <div class="chips" style="margin-bottom:10px">
    ${[['all', t('all'), rows.length], ['in', t('entered'), rows.filter((x) => isNum(x.r.h)).length], ['out', t('notEntered'), rows.filter((x) => !isNum(x.r.h)).length], ['prev', t('yesterday'), list.filter((w) => isNum(w.hours[prev])).length]]
      .map(([v, l, c]) => `<button class="chip ${f === v ? 'on' : ''}" data-act="dayf" data-v="${v}">${l}<small>${c}</small></button>`).join('')}
  </div>
  <div class="card list">
    ${shown.map(({ w, r }) => dayItem(w, r, iso)).join('') || `<div class="empty">—</div>`}
  </div>
  <p class="muted small">${t('holidayHint')}</p>`;
}

function dayItem(w, r, iso) {
  const has = isNum(r.h);
  const brk = [];
  if (has || r.manual.length) {
    if (r.reg) brk.push(`REG <b>${num(r.reg)}</b>`);
    if (r.ot) brk.push(`OT <b>${num(r.ot)}</b>`);
    if (r.dt) brk.push(`DT <b>${num(r.dt)}</b>`);
  }
  if (r.pd) brk.push(`PD <b>✓</b>`);
  const miss = missing(w).map((m) => `<span class="badge warn">${m}</span>`).join('');
  const man = r.manual.length ? `<span class="badge man" title="${t('manualCell')}">M ${r.manual.map((x) => x.toUpperCase()).join('/')}</span>` : '';
  const txt = w.hourText?.[iso] ? `<span class="badge bad">"${esc(w.hourText[iso])}"</span>` : '';
  return `<div class="item ${has && r.h > 0 ? 'has' : ''}">
    <div class="who" data-act="open" data-k="${w.key}">
      <div class="nm">${esc(w.name)}</div>
      <div class="meta">${w.trade ? `<span class="badge">${esc(w.trade)}</span>` : ''}<span>${esc(shortPC(w.payClass))}</span>${miss}${man}${txt}
        ${brk.length ? `<span class="brk">· ${brk.join(' · ')}</span>` : ''}</div>
    </div>
    <input class="hr ${has ? 'filled' : ''}" inputmode="decimal" enterkeyhint="next" autocomplete="off" value="${has ? r.h : ''}"
      data-act="hr" data-k="${w.key}" data-iso="${iso}" data-fk="h-${w.key}-${iso}" aria-label="${esc(w.name)}">
  </div>`;
}

// ----- Semana -----
function weekView() {
  const k = ui.week;
  const start = S.meta.weeks[k];
  const dates = weekDates(start);
  const list = active().sort(byName).filter(matches).map((w) => ({ w, r: cw(w, start) }));
  const shown = ui.weekOnlyHours ? list.filter(({ r }) => r.t.h > 0 || r.t.pd > 0 || r.days.some((d) => d.manual.length)) : list;
  const tot = { h: 0, reg: 0, ot: 0, dt: 0, pd: 0, cost: 0, day: dates.map(() => 0) };
  for (const { r } of shown) {
    for (const f of ['h', 'reg', 'ot', 'dt', 'pd']) tot[f] += r.t[f];
    tot.cost += r.cost.total;
    r.days.forEach((d, i) => (tot.day[i] += d.h || 0));
  }
  const dcls = (iso) => (dow(iso) === 6 ? 'sat' : dow(iso) === 7 ? 'sun' : '');
  return `
  <div class="chips" style="margin-bottom:10px">
    ${S.meta.weeks.map((s, i) => `<button class="chip ${i === k ? 'on' : ''}" data-act="week" data-v="${i}">W${i + 1}<small>${md(s)}</small></button>`).join('')}
  </div>
  <div class="row wrap" style="margin-bottom:10px">${searchBox('q-week')}
    <label class="chip ${ui.weekOnlyHours ? 'on' : ''}" style="display:flex;align-items:center"><input type="checkbox" hidden data-act="onlyh" ${ui.weekOnlyHours ? 'checked' : ''}>${t('withHours')}</label>
  </div>
  <div class="tbl-wrap">
  <table>
    <thead><tr><th>${t('worker')} <small>(${shown.length})</small></th>
      ${dates.map((d) => `<th class="${dcls(d)}">${wd(d)}<br><small>${md(d)}</small></th>`).join('')}
      <th>${t('hours')}</th><th>REG</th><th>OT</th><th>DT</th><th>PD</th><th>${t('cost')}</th></tr></thead>
    <tbody>
    ${shown.map(({ w, r }) => `<tr>
      <td class="nmc" data-act="open" data-k="${w.key}" title="${esc(w.name)}">${esc(w.name)}</td>
      ${r.days.map((d) => `<td class="${dcls(d.iso)}"><input class="cell-in ${isNum(d.h) ? 'filled' : ''}" inputmode="decimal" value="${isNum(d.h) ? d.h : ''}" data-act="hr" data-k="${w.key}" data-iso="${d.iso}" data-fk="w-${w.key}-${d.iso}"></td>`).join('')}
      <td class="wk"><b>${num(r.t.h)}</b></td>
      <td class="${r.days.some((d) => d.manual.includes('reg')) ? 'man' : ''}">${blank(r.t.reg)}</td>
      <td>${blank(r.t.ot)}</td><td>${blank(r.t.dt)}</td><td>${blank(r.t.pd)}</td>
      <td>${money0(r.cost.total)}</td></tr>`).join('')}
    </tbody>
    <tfoot><tr><td>${t('total')}</td>${tot.day.map((h) => `<td>${num(h)}</td>`).join('')}
      <td>${num(tot.h)}</td><td>${num(tot.reg)}</td><td>${num(tot.ot)}</td><td>${num(tot.dt)}</td><td>${num(tot.pd)}</td><td>${money0(tot.cost)}</td></tr></tfoot>
  </table></div>`;
}

// ----- Personal -----
function workersView() {
  const list = active().sort(byName).filter(matches);
  const flt = ui.wFilter === 'miss' ? list.filter((w) => missing(w).length || !isNum(w.rate)) : list;
  const nMiss = list.filter((w) => missing(w).length || !isNum(w.rate)).length;
  return `
  <div class="row" style="margin-bottom:8px">${searchBox('q-w')}
    <button class="btn primary" data-act="add" ${S.spare.length ? '' : 'disabled'}><svg viewBox="0 0 24 24"><path d="M12 5v14M5 12h14"/></svg>${t('addWorker')}</button></div>
  <p class="muted small" style="margin:0 0 8px">${S.spare.length ? t('spareLeft', { n: S.spare.length }) : t('noSpare')}</p>
  <div class="chips" style="margin-bottom:10px">
    <button class="chip ${ui.wFilter === 'all' ? 'on' : ''}" data-act="wf" data-v="all">${t('all')}<small>${list.length}</small></button>
    <button class="chip ${ui.wFilter === 'miss' ? 'on' : ''}" data-act="wf" data-v="miss">${t('missingInfo')}<small>${nMiss}</small></button>
  </div>
  <div class="card list">
  ${flt.map((w) => `<div class="item"><div class="who" data-act="open" data-k="${w.key}">
      <div class="nm">${esc(w.name)} ${w.isNew ? `<span class="badge acc">${t('newWorker')}</span>` : ''}</div>
      <div class="meta"><span>${esc(w.id)}</span>${w.trade ? `<span class="badge">${esc(w.trade)}</span>` : ''}<span>${esc(shortPC(w.payClass))}</span><span>${esc(w.perDiem)}</span>
      ${missing(w).map((m) => `<span class="badge warn">${m}</span>`).join('')}${isNum(w.rate) ? '' : `<span class="badge warn">${t('noRate')}</span>`}
      ${w.notes?.length ? `<span class="badge">✎ ${w.notes.length}</span>` : ''}</div></div>
      <div class="num" style="text-align:right"><b>${isNum(w.rate) && w.rate ? money(w.rate) : '—'}</b><div class="muted small">/hr</div></div></div>`).join('') || `<div class="empty">—</div>`}
  </div>`;
}

// ----- Ficha del trabajador -----
function openWorker(key) {
  ui.modalKey = key;
  ui.modalWeek = ui.tab === 'day' ? weekOf(ui.date) : ui.week;
  renderModal();
  const dlg = $('#modal');
  if (!dlg.open) dlg.showModal();
}

function renderModal() {
  const w = S.workers.find((x) => x.key === ui.modalKey);
  if (!w) return $('#modal').close();
  const a = document.activeElement;
  const fk = a?.dataset?.fk;
  const k = ui.modalWeek;
  const r = cw(w, S.meta.weeks[k]);
  const opts = (arr, cur) => {
    const all = [...new Set([...arr, cur].filter(Boolean))];
    return `<option value="">—</option>` + all.map((v) => `<option ${v === cur ? 'selected' : ''}>${esc(v)}</option>`).join('');
  };
  const pcs = S.lists.payClasses.length ? S.lists.payClasses : PAY_CLASSES;
  const pds = S.lists.perDiems.length ? S.lists.perDiems : PER_DIEMS;
  const projs = S.lists.projects.map((p) => p.name);
  const cell = (d, key) => {
    const v = d[key];
    const man = d.manual.includes(key);
    return `<td class="${man ? 'man' : ''} ${v === 0 ? 'zero' : ''}" ${man ? `title="${t('manualCell')}"` : ''}>${isNum(v) ? num(v) : ''}${man ? '<sup>M</sup>' : ''}</td>`;
  };
  const inRow = (key, lab) => `<tr><td>${lab}</td>${r.days.map((d) => {
    const v = key === 'h' ? w.hours[d.iso] : w[key]?.[d.iso];
    return `<td><input class="cell-in ${isNum(v) ? 'filled' : ''}" inputmode="decimal" value="${isNum(v) ? v : ''}" data-act="mcell" data-f="${key}" data-iso="${d.iso}" data-fk="m-${key}-${d.iso}"></td>`;
  }).join('')}<td class="wk"><b>${num(r.t[key])}</b></td><td>${key === 'h' ? '' : key === 'ml' ? money(r.cost.ml) : money(r.cost.tt)}</td></tr>`;
  const calcRow = (key, lab) => `<tr><td>${lab}</td>${r.days.map((d) => cell(d, key)).join('')}<td class="wk"><b>${num(r.t[key])}</b></td><td>${money(r.cost[key])}</td></tr>`;

  $('#modal-body').innerHTML = `
  <div class="m-head"><h2>${esc(w.name || t('newWorker'))}</h2><button class="btn icon" data-act="mclose" aria-label="${t('close')}"><svg viewBox="0 0 24 24"><path d="M18 6 6 18M6 6l12 12"/></svg></button></div>
  <div class="m-body">
    <div class="form">
      <div class="field full"><label>${t('name')}</label><input class="input" data-act="mf" data-f="name" value="${esc(w.name)}" data-fk="mf-name"></div>
      <div class="field"><label>${t('id')}</label><input class="input" data-act="mf" data-f="id" value="${esc(w.id)}" data-fk="mf-id"></div>
      <div class="field"><label>${t('trade')}</label><input class="input" data-act="mf" data-f="trade" value="${esc(w.trade)}" data-fk="mf-trade"></div>
      <div class="field"><label>${t('project')}</label><select class="input" data-act="mf" data-f="project">${opts(projs, w.project)}</select></div>
      <div class="field"><label>${t('payClass')}</label><select class="input" data-act="mf" data-f="payClass">${opts(pcs, w.payClass)}</select></div>
      <div class="field"><label>${t('perDiem')}</label><select class="input" data-act="mf" data-f="perDiem">${opts(pds, w.perDiem)}</select></div>
      <div class="field"><label>${t('rate')}</label><input class="input" inputmode="decimal" data-act="mf" data-f="rate" value="${isNum(w.rate) ? w.rate : ''}" data-fk="mf-rate"></div>
      <div class="field"><label>${t('pdRate')}</label><input class="input" inputmode="decimal" data-act="mf" data-f="pdRate" value="${isNum(w.pdRate) ? w.pdRate : ''}" data-fk="mf-pd"></div>
    </div>
    ${w.notes?.length ? `<h3>${t('notes')}</h3>${w.notes.map((n) => `<div class="note"><b>${esc(n.ref)}</b> ${esc(n.text)} <span class="muted">— ${esc(n.author)}</span></div>`).join('')}` : ''}
    <h3>${t('week')}</h3>
    <div class="chips" style="margin-bottom:8px">${S.meta.weeks.map((s, i) => `<button class="chip ${i === k ? 'on' : ''}" data-act="mweek" data-v="${i}">W${i + 1}<small>${md(s)}</small></button>`).join('')}</div>
    <div class="tbl-wrap" style="max-height:none">
    <table>
      <thead><tr><th></th>${r.days.map((d) => `<th>${wd(d.iso)}<br><small>${md(d.iso)}</small></th>`).join('')}<th>${t('total')}</th><th>${t('cost')}</th></tr></thead>
      <tbody>
        ${inRow('h', t('hours'))}
        ${calcRow('reg', 'REG')}${calcRow('ot', 'OT')}${calcRow('dt', 'DT')}${calcRow('pd', 'PD')}
        ${inRow('ml', 'ML')}${inRow('tt', 'TT')}
      </tbody>
      <tfoot><tr><td>${t('total')}</td><td colspan="8"></td><td>${money(r.cost.total)}</td></tr></tfoot>
    </table></div>
    <p class="muted small">${t('holidayHint')} · OT ${money(isNum(w.otRateFixed) ? w.otRateFixed : (w.rate || 0) * 1.5)} · DT ${money(isNum(w.dtRateFixed) ? w.dtRateFixed : (w.rate || 0) * 2)} · ML ${money(w.mlRate || 0)}</p>
    <div class="row" style="margin-top:16px"><button class="btn danger" data-act="remove">${t('remove')}</button><span class="spacer"></span><button class="btn dark" data-act="mclose">${t('close')}</button></div>
  </div>`;
  if (fk) $(`[data-fk="${CSS.escape(fk)}"]`, $('#modal-body'))?.focus({ preventScroll: true });
}

// ----- Resumen -----
function summaryView() {
  const ks = ui.sumWeek < 0 ? S.meta.weeks.map((_, i) => i) : [ui.sumWeek];
  const cat = { reg: [0, 0], ot: [0, 0], dt: [0, 0], pd: [0, 0], ml: [0, 0], tt: [0, 0] };
  const byPC = new Map();
  const byWeek = S.meta.weeks.map(() => ({ h: 0, cost: 0, n: 0 }));
  const alerts = [];
  let total = 0, hours = 0;
  const people = new Set();
  for (const w of active()) {
    let wh = 0;
    S.meta.weeks.forEach((s, k) => {
      const r = cw(w, s);
      byWeek[k].h += r.t.h;
      byWeek[k].cost += r.cost.total;
      if (r.t.h > 0) byWeek[k].n++;
      if (!ks.includes(k)) return;
      for (const c in cat) { cat[c][0] += r.t[c]; cat[c][1] += r.cost[c]; }
      total += r.cost.total;
      hours += r.t.h;
      wh += r.t.h;
      if (r.t.h > 0) people.add(w.key);
      const pc = w.payClass || '—';
      const e = byPC.get(pc) || { h: 0, cost: 0 };
      e.h += r.t.h; e.cost += r.cost.total;
      byPC.set(pc, e);
      const man = r.days.filter((d) => d.manual.some((m) => isNum(d[m]))).length;
      if (man) alerts.push({ cls: 'man', k: w.key, msg: t('alertManual', { n: w.name, c: man }) + ` (W${k + 1})` });
      for (const d of r.days) if (w.hourText?.[d.iso]) alerts.push({ cls: 'bad', k: w.key, msg: t('alertText', { n: w.name, t: w.hourText[d.iso], d: mdy(d.iso) }) });
    });
    if (wh > 0 && missing(w).length) alerts.unshift({ cls: 'warn', k: w.key, msg: t('alertMissing', { n: w.name, w: missing(w).join(', ') }) });
  }
  const maxW = Math.max(...byWeek.map((x) => x.cost), 1);
  const maxPC = Math.max(...[...byPC.values()].map((x) => x.cost), 1);
  // asistencia de la semana elegida (o la de la fecha actual)
  const ak = ui.sumWeek < 0 ? weekOf(ui.date) : ui.sumWeek;
  const adates = weekDates(S.meta.weeks[ak]);
  const att = adates.map((iso) => {
    let n = 0, h = 0;
    for (const w of active()) if (isNum(w.hours[iso]) && w.hours[iso] > 0) { n++; h += w.hours[iso]; }
    return { iso, n, h };
  });
  return `
  <div class="chips" style="margin-bottom:12px">
    <button class="chip ${ui.sumWeek < 0 ? 'on' : ''}" data-act="sumw" data-v="-1">${t('allWeeks')}</button>
    ${S.meta.weeks.map((s, i) => `<button class="chip ${i === ui.sumWeek ? 'on' : ''}" data-act="sumw" data-v="${i}">W${i + 1}<small>${md(s)}</small></button>`).join('')}
  </div>
  <div class="kpis">
    <div class="kpi big"><span>${t('grandTotal')}</span><b>${money(total)}</b></div>
    <div class="kpi"><span>${t('hours')}</span><b>${num(hours)}</b></div>
    <div class="kpi"><span>${t('headcount')}</span><b>${people.size}</b></div>
  </div>
  <h3>${t('byCategory')}</h3>
  <div class="tbl-wrap plain" style="max-height:none"><table class="plain">
    <thead><tr><th>${t('category')}</th><th>${t('hours')} / ${t('days')}</th><th>${t('cost')}</th></tr></thead>
    <tbody>${Object.entries(cat).map(([c, [h, co]]) => `<tr><td>${c.toUpperCase()}</td><td>${num(h)}</td><td>${money(co)}</td></tr>`).join('')}</tbody>
    <tfoot><tr><td>${t('total')}</td><td></td><td>${money(total)}</td></tr></tfoot>
  </table></div>
  <h3>${t('byPayClass')}</h3>
  <div class="tbl-wrap plain" style="max-height:none"><table class="plain">
    <thead><tr><th>${t('payClass')}</th><th>${t('hours')}</th><th>${t('cost')}</th><th class="bar-cell"></th></tr></thead>
    <tbody>${[...byPC.entries()].sort((a, b) => b[1].cost - a[1].cost).map(([pc, e]) => `<tr><td>${esc(pc)}</td><td>${num(e.h)}</td><td>${money(e.cost)}</td><td class="bar-cell"><div class="bar" style="width:${(e.cost / maxPC) * 100}%"></div></td></tr>`).join('')}</tbody>
  </table></div>
  <h3>${t('byWeek')}</h3>
  <div class="tbl-wrap plain" style="max-height:none"><table class="plain">
    <thead><tr><th>${t('week')}</th><th>${t('headcount')}</th><th>${t('hours')}</th><th>${t('cost')}</th><th class="bar-cell"></th></tr></thead>
    <tbody>${byWeek.map((e, i) => `<tr><td>W${i + 1} · ${md(S.meta.weeks[i])}</td><td>${e.n}</td><td>${num(e.h)}</td><td>${money(e.cost)}</td><td class="bar-cell"><div class="bar" style="width:${(e.cost / maxW) * 100}%"></div></td></tr>`).join('')}</tbody>
  </table></div>
  <h3>${t('attendance')} · W${ak + 1}</h3>
  <div class="tbl-wrap plain" style="max-height:none"><table class="plain">
    <thead><tr><th></th>${att.map((a) => `<th>${wd(a.iso)} ${md(a.iso)}</th>`).join('')}</tr></thead>
    <tbody><tr><td>${t('headcount')}</td>${att.map((a) => `<td>${a.n || ''}</td>`).join('')}</tr>
    <tr><td>${t('hours')}</td>${att.map((a) => `<td>${num(a.h)}</td>`).join('')}</tr></tbody>
  </table></div>
  <h3>${t('alerts')} (${alerts.length})</h3>
  <div class="card">${alerts.map((a) => `<div class="alert ${a.cls}" data-act="open" data-k="${a.k}" style="cursor:pointer"><span class="dot"></span><span>${esc(a.msg)}</span></div>`).join('') || `<div class="empty">${t('noAlerts')}</div>`}</div>`;
}

// ----- Archivo -----
function spSection() {
  const icon = '<svg viewBox="0 0 24 24"><path d="M7 18a5 5 0 0 1-.6-9.96A6 6 0 0 1 18 9a4.5 4.5 0 0 1 0 9z"/></svg>';
  if (!ms.configured()) {
    return `<h3>SharePoint / OneDrive</h3>
    <div class="card pad">
      <p class="small" style="margin-top:0">${t('spSetup')}</p>
      <div class="field"><label>Application (client) ID</label>
        <input class="input" data-act="cid" placeholder="xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx" autocomplete="off" spellcheck="false"></div>
      <p class="muted small">${t('spRedirect')}<br><code style="user-select:all">${esc(ms.redirectUri())}</code></p>
    </div>`;
  }
  const u = ms.user();
  const r = S.meta.remote;
  const n = changeList().length;
  return `<h3>SharePoint / OneDrive</h3>
  <div class="card pad">
    <div class="row wrap small" style="margin-bottom:10px">
      <span class="grow muted">${u ? `${t('spAs')} <b>${esc(u.username)}</b>` : t('spNotSigned')}</span>
      ${u ? `<button class="btn sm" data-act="spout">${t('spSignOut')}</button>` : `<button class="btn sm" data-act="spin">${t('spSignIn')}</button>`}
    </div>
    ${r ? `<p class="small" style="margin:0 0 10px">${t('spLinkedTo')} <a href="${esc(r.webUrl)}" target="_blank" rel="noopener"><b>${esc(r.name)}</b></a></p>
      <button class="btn primary" data-act="spsend" style="width:100%;height:48px" ${n ? '' : 'disabled'}>${icon}${t('spSendBtn', { n })}</button>
      <div class="row wrap" style="margin-top:8px"><button class="btn sm" data-act="sprefresh">${t('spRefresh')}</button><button class="btn sm" data-act="splink">${t('spRelink')}</button></div>`
    : `<p class="small muted" style="margin:0 0 10px">${t('spNotLinked')}</p>
      <button class="btn primary" data-act="splink" style="width:100%;height:44px">${icon}${t('spLinkBtn')}</button>`}
    <div class="row wrap" style="margin-top:8px"><button class="btn sm" data-act="spopen">${t('spOpenBtn')}</button><span class="spacer"></span><button class="btn sm" data-act="cidreset">${t('spChangeId')}</button></div>
  </div>`;
}

function fileView() {
  const ch = changeList();
  const m = S.meta;
  const fmt = (v) => (v === undefined || v === null || v === '' ? '∅' : esc(v));
  return `
  <div class="card pad">
    <dl class="kv">
      <dt>${t('file')}</dt><dd>${esc(m.fileName)}</dd>
      <dt>${t('company')}</dt><dd>${esc(m.company)}</dd>
      <dt>${t('location')}</dt><dd>${esc(m.location)}</dd>
      <dt>${t('job')}</dt><dd>${esc(m.job)}</dd>
      <dt>${t('weeks')}</dt><dd>${m.weeks.length} · ${mdy(m.weeks[0])} – ${mdy(addDays(m.weeks.at(-1), 6))}</dd>
      <dt>${t('workers')}</dt><dd>${active().length}</dd>
      <dt>${t('imported')}</dt><dd>${new Date(m.importedAt).toLocaleString('en-US')}</dd>
    </dl>
  </div>
  ${spSection()}
  <p style="margin-top:16px"><button class="btn ${S.meta.remote && ms.configured() ? '' : 'primary'}" data-act="export" style="width:100%;height:48px"><svg viewBox="0 0 24 24"><path d="M12 3v12M7 10l5 5 5-5M5 21h14"/></svg>${t('export')}</button></p>
  <p class="muted small">${t('exportHint')}</p>
  <h3>${t('changes')} (${ch.length})</h3>
  <div class="card changes">${ch.map((c) => `<div class="chg"><span><b>${esc(c.who)}</b> · ${esc(c.what)}</span><span class="num"><s>${fmt(c.from)}</s> → <b>${fmt(c.to)}</b></span></div>`).join('') || `<div class="empty">${t('noChanges')}</div>`}</div>
  <h3>${t('importOther')} (${FILES.length})</h3>
  <div class="card list" style="margin-bottom:10px">${FILES.map((f) => `
    <div class="item ${f.id === CUR ? 'has' : ''}"><div class="who" data-act="usefile" data-id="${f.id}">
      <div class="nm">${esc(f.label)} ${f.id === CUR ? `<span class="badge ok">${t('open')}</span>` : ''}</div>
      <div class="meta"><span>${esc(f.fileName)}</span></div></div></div>`).join('')}</div>
  <div class="drop" id="drop" style="padding:20px"><button class="btn" data-act="pick">${t('addFile')}</button><p class="muted small" style="margin:8px 0 0">${t('dropHere')} · ${t('filesHint')}</p></div>
  <div class="row wrap" style="margin-top:18px"><button class="btn sm" data-act="lang">${t('lang')}</button><span class="spacer"></span><button class="btn sm danger" data-act="wipe">${t('wipe')}</button></div>
  <p class="muted small">${t('installHint')}</p>`;
}

async function doExport() {
  const ch = changeList();
  const changes = new Map();
  for (const c of ch) {
    if (!changes.has(c.row)) changes.set(c.row, new Map());
    changes.get(c.row).set(c.col, c.v);
  }
  try {
    const { blob, skipped } = await exportWorkbook(FILE, S.meta.sheetName, changes);
    const d = new Date();
    const stamp = `${String(d.getMonth() + 1).padStart(2, '0')}-${String(d.getDate()).padStart(2, '0')} ${String(d.getHours()).padStart(2, '0')}${String(d.getMinutes()).padStart(2, '0')}`;
    const name = S.meta.fileName.replace(/\.xlsx$/i, '') + ` (${stamp}).xlsx`;
    const file = new File([blob], name, { type: blob.type });
    if (navigator.canShare?.({ files: [file] }) && /iPhone|iPad|Android/i.test(navigator.userAgent)) {
      await navigator.share({ files: [file], title: name }).catch(() => {});
    } else {
      const a = document.createElement('a');
      a.href = URL.createObjectURL(blob);
      a.download = name;
      a.click();
      setTimeout(() => URL.revokeObjectURL(a.href), 5000);
    }
    S.exportedSig = getSig();
    save();
    toast(skipped.length ? t('skipped', { n: skipped.length, c: skipped.slice(0, 6).join(', ') }) : t('exported'), skipped.length ? 6000 : 2600);
  } catch (e) {
    console.error(e);
    toast(e.message, 5000);
  }
}

// ---------- edición ----------
function setHours(w, iso, raw, field = 'hours') {
  const v = parseH(raw);
  if (v === undefined) return false;
  w[field] ||= {};
  if (v === null) delete w[field][iso];
  else w[field][iso] = v;
  save();
  return true;
}

function copyPrev() {
  const iso = ui.date;
  const prev = addDays(iso, -1);
  if (!S.meta.dayCol[prev]) return toast(t('nothingCopy'));
  const list = active().filter((w) => isNum(w.hours[prev]) && w.hours[prev] > 0 && !isNum(w.hours[iso]));
  if (!list.length) return toast(t('nothingCopy'));
  if (!confirm(t('copyPrevQ', { n: list.length, d: `${wd(prev)} ${mdy(prev)}` }))) return;
  for (const w of list) w.hours[iso] = w.hours[prev];
  save();
  render();
  toast(t('copied', { n: list.length }));
}

function addWorker() {
  if (!S.spare.length) return toast(t('noSpare'));
  const sp = S.spare.shift();
  const w = {
    key: 'r' + sp.row, row: sp.row, off: sp.off, isNew: true,
    id: '', name: '', project: S.lists.projects[0]?.name || '', perDiem: '', payClass: '', trade: '',
    rate: null, pdRate: null, mlRate: 0.5, ttRate: null, hours: {}, hourText: {}, ml: {}, tt: {}, over: {}, notes: [],
  };
  w.orig = snap(w);
  w.ruleDefault = sp.rule || null;
  w.pdRate = 115;
  S.workers.push(w);
  save();
  openWorker(w.key);
  setTimeout(() => $('[data-fk="mf-name"]')?.focus(), 50);
}

function removeWorker(w) {
  if (!confirm(t('removeQ', { n: w.name || '—' }))) return;
  if (w.isNew) {
    S.workers = S.workers.filter((x) => x !== w);
    S.spare.push({ row: w.row, off: w.off });
    S.spare.sort((a, b) => a.row - b.row);
  } else w.removed = true;
  save();
  $('#modal').close();
  render();
}

// ---------- quickbar (horas rápidas en la vista Día) ----------
const QUICK = [8, 10, 10.5, 11, 12];
function showQuick(input) {
  const w = S.workers.find((x) => x.key === input.dataset.k);
  const qb = $('#quickbar');
  const own = mostCommon(w);
  const vals = [...new Set([...(own ? [own] : []), ...QUICK])];
  qb.innerHTML = `<span class="who">${esc(w.name)}</span>` + vals.map((v) => `<button data-q="${v}">${v}</button>`).join('') + `<button class="alt" data-q="">${t('clear')}</button>`;
  qb.hidden = false;
  qb.dataset.fk = input.dataset.fk;
}
function mostCommon(w) {
  const c = {};
  for (const v of Object.values(w.hours)) if (v > 0) c[v] = (c[v] || 0) + 1;
  const best = Object.entries(c).sort((a, b) => b[1] - a[1])[0];
  return best ? +best[0] : null;
}

// ---------- eventos ----------
document.addEventListener('click', (e) => {
  const db_ = e.target.closest('.datebtn');
  if (db_) {
    e.preventDefault();
    try { db_.querySelector('input').showPicker(); } catch {}
    return;
  }
  const el = e.target.closest('[data-act], [data-tab]');
  if (!el) return;
  if (el.dataset.tab) {
    if (el.disabled) return;
    ui.tab = el.dataset.tab;
    ui.q = '';
    render();
    window.scrollTo(0, 0);
    return;
  }
  const act = el.dataset.act;
  const w = el.dataset.k && S?.workers.find((x) => x.key === el.dataset.k);
  switch (act) {
    case 'pick': $('#file-input').click(); if ($('#modal').open) $('#modal').close(); break;
    case 'lang': setLang(getLang() === 'es' ? 'en' : 'es'); render(); break;
    case 'day': {
      ui.date = addDays(ui.date, +el.dataset.d);
      ui.week = weekOf(ui.date);
      render();
      break;
    }
    case 'dayf': ui.dayFilter = el.dataset.v; render(); break;
    case 'copyprev': copyPrev(); break;
    case 'week': ui.week = +el.dataset.v; render(); break;
    case 'wf': ui.wFilter = el.dataset.v; render(); break;
    case 'sumw': ui.sumWeek = +el.dataset.v; render(); break;
    case 'open': if (w) openWorker(w.key); break;
    case 'add': addWorker(); break;
    case 'mclose': $('#modal').close(); break;
    case 'mweek': ui.modalWeek = +el.dataset.v; renderModal(); break;
    case 'remove': removeWorker(S.workers.find((x) => x.key === ui.modalKey)); break;
    case 'export': doExport(); break;
    case 'spopen': pickRemote('open'); break;
    case 'splink': pickRemote('link'); break;
    case 'rpick': usePicked(+el.dataset.i); break;
    case 'spsend': sendRemote(); break;
    case 'sprefresh':
      if (!hasPending(S) || confirm(t('replaceQ', { n: hasPending(S), f: FILES.find((f) => f.id === CUR)?.label }))) refreshRemote();
      break;
    case 'spin': ms.signIn(null); break;
    case 'spout': if (confirm(t('spSignOutQ'))) ms.signOut(); break;
    case 'cidreset': if (confirm(t('spChangeIdQ'))) { ms.setClientId(''); render(); } break;
    case 'wipe':
      if (confirm(t('wipeQ', { f: FILES.find((f) => f.id === CUR)?.label || '' }))) removeFile(CUR);
      break;
    case 'files': openFiles(); break;
    case 'usefile':
      switchFile(el.dataset.id).then(() => { $('#modal').close(); ui.tab = 'day'; render(); window.scrollTo(0, 0); });
      break;
    case 'rmfile': {
      const f = FILES.find((x) => x.id === el.dataset.id);
      db.get('state:' + f.id).then((st) => {
        const n = f.id === CUR ? hasPending(S) : hasPending(st);
        if (confirm(t('removeFileQ', { f: f.label }) + (n ? ` (${t('unsaved', { n })}!)` : ''))) removeFile(f.id).then(() => (FILES.length ? openFiles() : $('#modal').close()));
      });
      break;
    }
  }
});

document.addEventListener('change', (e) => {
  const el = e.target;
  if (el.dataset?.act === 'cid') {
    const v = el.value.trim();
    if (!/^[0-9a-f-]{36}$/i.test(v)) return toast(t('spBadId'));
    ms.setClientId(v);
    ms.init().then(() => render()).catch((er) => toast(er.message, 6000));
    return;
  }
  const act = el.dataset?.act;
  if (!act) return;
  if (el.id === 'file-input') return;
  if (act === 'date') {
    if (el.value && S.meta.dayCol[el.value]) { ui.date = el.value; ui.week = weekOf(el.value); }
    else toast(t('outOfRange', { a: mdy(S.meta.weeks[0]), b: mdy(addDays(S.meta.weeks.at(-1), 6)) }));
    render();
  } else if (act === 'hr') {
    const w = S.workers.find((x) => x.key === el.dataset.k);
    if (!setHours(w, el.dataset.iso, el.value)) { toast('0 – 24'); el.value = w.hours[el.dataset.iso] ?? ''; return; }
    render();
  } else if (act === 'onlyh') {
    ui.weekOnlyHours = el.checked;
    render();
  } else if (act === 'mcell') {
    const w = S.workers.find((x) => x.key === ui.modalKey);
    const f = el.dataset.f === 'h' ? 'hours' : el.dataset.f;
    if (!setHours(w, el.dataset.iso, el.value, f)) { toast('0 – 24'); return; }
    renderModal();
    render();
  } else if (act === 'mf') {
    const w = S.workers.find((x) => x.key === ui.modalKey);
    const f = el.dataset.f;
    if (f === 'rate' || f === 'pdRate') {
      const s = el.value.trim().replace(',', '.').replace('$', '');
      const v = s === '' ? null : Number(s);
      if (s !== '' && !Number.isFinite(v)) return toast('#');
      w[f] = v;
    } else w[f] = el.value.trim();
    save();
    renderModal();
    render();
  }
});

document.addEventListener('submit', (e) => {
  if (e.target.dataset?.form === 'spsearch') {
    e.preventDefault();
    pickRemote(ui.pick?.mode || 'open', e.target.q.value.trim());
  }
});

document.addEventListener('input', (e) => {
  if (e.target.dataset?.act === 'q') {
    ui.q = e.target.value;
    render();
  }
});

// Enter pasa al siguiente campo de horas
document.addEventListener('keydown', (e) => {
  const el = e.target;
  if (e.key !== 'Enter' || !el.classList?.contains('hr') && !el.classList?.contains('cell-in')) return;
  e.preventDefault();
  const scope = el.closest('#modal-body') || view;
  const all = [...scope.querySelectorAll('input.hr, input.cell-in')];
  const i = all.indexOf(el);
  const next = el.classList.contains('hr') ? all[i + 1] : all[i + 7] || all[i + 1];
  (next || el).focus();
  if (!next) el.blur();
});

document.addEventListener('focusin', (e) => {
  if (e.target.classList?.contains('hr')) showQuick(e.target);
});
document.addEventListener('focusout', (e) => {
  if (e.target.classList?.contains('hr')) setTimeout(() => { if (!document.activeElement?.classList?.contains('hr')) $('#quickbar').hidden = true; }, 120);
});

const qb = $('#quickbar');
qb.addEventListener('pointerdown', (e) => e.preventDefault());
qb.addEventListener('click', (e) => {
  const b = e.target.closest('button[data-q]');
  if (!b) return;
  const input = view.querySelector(`[data-fk="${CSS.escape(qb.dataset.fk)}"]`);
  if (!input) return;
  const w = S.workers.find((x) => x.key === input.dataset.k);
  setHours(w, input.dataset.iso, b.dataset.q);
  // pasar al siguiente
  const all = [...view.querySelectorAll('input.hr')];
  const next = all[all.indexOf(input) + 1];
  const nextFk = b.dataset.q && next ? next.dataset.fk : input.dataset.fk;
  render();
  const el = view.querySelector(`[data-fk="${CSS.escape(nextFk)}"]`);
  if (el) { el.focus({ preventScroll: false }); el.scrollIntoView({ block: 'center', behavior: 'smooth' }); showQuick(el); }
});

$('#file-input').addEventListener('change', (e) => {
  const f = e.target.files[0];
  e.target.value = '';
  if (f) importFile(f);
});
$('#hdr-dirty').addEventListener('click', () => { ui.tab = 'file'; render(); });
$('#modal').addEventListener('close', () => { ui.modalKey = null; render(); });

// ---------- arranque ----------
(async () => {
  try {
    FILES = (await db.get('index')) || [];
    const old = await db.get('state');
    if (old && !FILES.length) {
      const id = fileId(old.meta.fileName);
      await db.set('state:' + id, old);
      await db.set('file:' + id, await db.get('file'));
      FILES = [{ id, label: baseName(old.meta.fileName), job: old.meta.job, fileName: old.meta.fileName, importedAt: old.meta.importedAt }];
      await db.set('index', FILES);
      await Promise.all([db.del('state'), db.del('file')]);
    }
    const cur = await db.get('current');
    const id = FILES.find((f) => f.id === cur)?.id || FILES[0]?.id;
    if (id) await switchFile(id);
  } catch (e) {
    console.error(e);
  }
  render();
  if (ms.configured()) {
    try {
      const pending = await ms.init();
      if (location.hash.includes('code=') || location.hash.includes('state=')) history.replaceState(null, '', location.pathname);
      render();
      if (pending?.act === 'pick') pickRemote(pending.mode);
      else if (pending?.act === 'send' && S) { ui.tab = 'file'; render(); sendRemote(); }
      else if (pending?.act === 'refresh' && S) refreshRemote();
    } catch (e) {
      console.error(e);
      toast(t('spError', { e: e.message }), 6000);
    }
  }
  if ('serviceWorker' in navigator && location.protocol !== 'file:') {
    navigator.serviceWorker.register('sw.js').catch(() => {});
    let reloaded = false;
    navigator.serviceWorker.addEventListener('controllerchange', () => {
      if (reloaded || document.querySelector('input:focus')) return;
      reloaded = true;
      location.reload();
    });
  }
})();
