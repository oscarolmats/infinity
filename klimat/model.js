// Data model for the Klimatanalys page. The table is always rendered from this,
// never read back from the DOM.
//
// A row:   { id, raw: [cell text as loaded], cells: {header: text}, type, desc,
//            area, hasArea, volume, count, length,
//            res: resource key | null, generic: Boverket key an EPD replaced | null,
//            layers: [{ name, t (mm), res, generic, mix: { name, pct, res, generic } | null }] | null }
import { RES, MANUAL, factorOf, normUnit, wasteMultiplier } from './resources.js';

export const app = {
  rows: [],
  headers: [],        // file columns, in file order
  groupColumns: [],   // the text columns that can be grouped on
  title: '',
  sourceFile: '',     // the Excel/CSV the rows came from (originalFileName in projects)
  fileHandle: null,   // where "Spara" writes, after "Spara som"
  dirty: false,
  view: { filter: 'all', search: '', groupBy: '', cols: 'overview', open: new Set(), ticked: new Set(), drawer: null, copied: null },
  bta: ''
};

// Same parsing as INFINITY's parseNumberLike: spaces out, comma as decimal point, NaN if empty.
export function parseNum(v) {
  if (typeof v === 'number') return v;
  if (v == null) return NaN;
  const n = parseFloat(String(v).trim().replace(/\s+/g, '').replace(',', '.'));
  return Number.isFinite(n) ? n : NaN;
}
const num0 = (v) => { const n = parseNum(v); return Number.isFinite(n) ? n : 0; };

let nextId = 1;
const findCol = (lower, ...names) => lower.findIndex(h => names.includes(h));

// Builds rows from a header row and data rows (arrays of cell text).
export function makeRows(headers, dataRows) {
  const clean = headers.map(h => String(h ?? '').trim());
  const lower = clean.map(h => h.toLowerCase());
  const idx = {
    type: findCol(lower, 'type', 'object type', 'typ'),
    desc: findCol(lower, 'description', 'beskrivning'),
    area: findCol(lower, 'net area', 'area'),
    volume: findCol(lower, 'volume', 'volym'),
    count: findCol(lower, 'count', 'antal'),
    length: findCol(lower, 'length', 'längd')
  };
  idx.component = findCol(lower, 'component', 'komponent');
  idx.material = findCol(lower, 'material');
  const rows = dataRows
    .filter(row => Array.isArray(row) && row.some(c => String(c ?? '').trim() !== ''))
    .map(row => rowFrom(row.map(c => String(c ?? '')), clean, idx));
  // groupable: every named column with some value that isn't a number (as before)
  const isNumber = (v) => /^-?[\d\s]+([.,]\d+)?$/.test(String(v).trim());
  const groupColumns = clean.filter((h, i) => h && rows.some(r => String(r.raw[i] ?? '').trim() !== '' && !isNumber(r.raw[i])));
  return { rows, headers: clean, groupColumns, idx };
}

function rowFrom(raw, headers, idx, id) {
  const get = (i) => (i >= 0 ? String(raw[i] ?? '').trim() : '');
  const cells = {};
  headers.forEach((h, i) => { if (h) cells[h] = String(raw[i] ?? '').trim(); });
  const areaN = parseNum(get(idx.area));
  return {
    id: id ?? nextId++, raw, cells,
    // the row's name: its type, or what it is when the type is empty
    type: get(idx.type >= 0 ? idx.type : 0) || get(idx.component ?? -1) || get(idx.material ?? -1) || '(tom)',
    desc: get(idx.desc),
    area: Number.isFinite(areaN) ? areaN : 0, hasArea: Number.isFinite(areaN),
    volume: num0(get(idx.volume)), count: num0(get(idx.count)), length: num0(get(idx.length)),
    res: null, generic: null, layers: null
  };
}

export function columnIndex() {
  const lower = app.headers.map(h => h.toLowerCase());
  return {
    component: findCol(lower, 'component', 'komponent'), material: findCol(lower, 'material'),
    type: findCol(lower, 'type', 'object type', 'typ'), desc: findCol(lower, 'description', 'beskrivning'),
    area: findCol(lower, 'net area', 'area'), volume: findCol(lower, 'volume', 'volym'),
    count: findCol(lower, 'count', 'antal'), length: findCol(lower, 'length', 'längd')
  };
}
// A new row typed in by the user ("Lägg till rad").
export function addRow(values) {
  const raw = app.headers.map(h => String(values[h] ?? '').trim());
  const r = rowFrom(raw, app.headers, columnIndex());
  app.rows.unshift(r);
  return r;
}

export function setRows({ rows, headers, groupColumns }) {
  app.rows = rows;
  app.headers = headers;
  app.groupColumns = groupColumns;
}

// INFINITY's row signature: the first six values, JSON-encoded (getRowSignature).
export const sigOf = (values) => JSON.stringify(values.slice(0, 6).map(v => String(v ?? '')));
export const rowSig = (r) => sigOf(r.raw);

/* ---------- calculation (same rules as INFINITY) ----------
   Boverket: inbyggd vikt = factor × Volume (kg/m³) or × Net Area (kg/m²);
             inköpt vikt = inbyggd × spillfaktor; A1-A3 and A4 on inbyggd, A5 on inköpt.
   EPD/own:  also kg/m (× Length) and kg/st (× Count); A1-A3, A4 and A5 on inbyggd. */
export const ZERO = { w: 0, bought: 0, a13: 0, a4: 0, a5: 0 };
export function impact(q, key) {
  if (!key || !RES[key]) return null;
  const x = RES[key];
  const { f, u } = factorOf(key);
  const unit = normUnit(u), fac = Number(f);
  let w = null;
  if (Number.isFinite(fac)) {
    if (unit === 'kg/m3') w = fac * q.volume;
    else if (unit === 'kg/m2') w = fac * q.area;
    else if (x.kind === 'custom' && unit === 'kg/m') w = fac * q.length;
    else if (x.kind === 'custom' && unit === 'kg/st') w = fac * q.count;
  }
  if (w == null || !Number.isFinite(w)) return { ...ZERO, noWeight: true };
  const waste = wasteMultiplier(x.w);
  const bought = waste ? w * waste : w;
  // boverketRules: a saved Boverket resource that isn't in this Boverket version
  if (x.kind === 'custom' && !x.boverketRules) return { w, bought, a13: w * (x.a13 || 0), a4: w * (x.a4 || 0), a5: w * (x.a5 || 0) };
  return { w, bought, a13: w * (x.a13 || 0), a4: w * (x.a4 || 0), a5: waste ? bought * (x.a5 || 0) : 0 };
}
export const addImpact = (a, b) => (b ? { w: a.w + b.w, bought: a.bought + b.bought, a13: a.a13 + b.a13, a4: a.a4 + b.a4, a5: a.a5 + b.a5 } : a);
export const tot = (i) => (i ? i.a13 + i.a4 + i.a5 : 0);

// A layer's quantities, as in INFINITY (cloneRowWithMultiplier): with a thickness and
// a Net Area the layer is Net Area × thickness (mm -> m) and keeps the whole area;
// otherwise volume and area are split by thickness share (evenly without thicknesses).
export function layerQty(row, l) {
  const t = Number(l.t) || 0;
  if (t > 0 && row.hasArea) return { volume: row.area * t / 1000, area: row.area };
  const sumT = row.layers.reduce((s, x) => s + (Number(x.t) || 0), 0);
  const share = sumT ? t / sumT : 1 / row.layers.length;
  return { volume: row.volume * share, area: row.area * share };
}
// The quantity sets a row or layer is made of: [{ q, key, generic, weight share }]
export function parts(row) {
  const base = { count: row.count, length: row.length };
  if (!row.layers) return [{ q: { ...base, volume: row.volume, area: row.area }, key: row.res, generic: row.generic }];
  const out = [];
  row.layers.forEach(l => {
    const { volume, area } = layerQty(row, l);
    const p2 = l.mix ? l.mix.pct / 100 : 0;
    out.push({ q: { ...base, volume: volume * (1 - p2), area: area * (1 - p2) }, key: l.res, generic: l.generic || null, layer: l });
    if (l.mix) out.push({ q: { ...base, volume: volume * p2, area: area * p2 }, key: l.mix.res, generic: l.mix.generic || null, layer: l, part2: true });
  });
  return out;
}

// Recomputed after every change: per-row results, status, per-layer results.
export function derive() {
  app.rows.forEach(row => {
    if (row.layers) {
      row.layerResults = row.layers.map(l => {
        const { volume, area } = layerQty(row, l);
        const p2 = l.mix ? l.mix.pct / 100 : 0;
        const base = { count: row.count, length: row.length };
        let i = impact({ ...base, volume: volume * (1 - p2), area: area * (1 - p2) }, l.res);
        if (l.mix) i = addImpact(i || ZERO, impact({ ...base, volume: volume * p2, area: area * p2 }, l.mix.res));
        return { ...l, volume, impact: i, mapped: !!l.res && (!l.mix || !!l.mix.res) };
      });
      const done = row.layerResults.filter(l => l.mapped).length;
      row.status = done === row.layers.length ? 'ok' : (done || row.layers.some(l => l.res || (l.mix && l.mix.res))) ? 'part' : 'none';
      row.impact = row.layerResults.some(l => l.impact) ? row.layerResults.reduce((s, l) => addImpact(s, l.impact), ZERO) : null;
    } else {
      row.status = row.res ? 'ok' : 'none';
      row.impact = impact({ volume: row.volume, area: row.area, count: row.count, length: row.length }, row.res);
    }
  });
}

// Project totals plus what EPDs saved against the Boverket resources they replaced.
export function totals(rows = app.rows) {
  const t = rows.reduce((s, r) => addImpact(s, r.impact), ZERO);
  let reduction = 0, withGeneric = 0, epdKg = 0;
  rows.forEach(r => parts(r).forEach(p => {
    if (!p.key || !RES[p.key] || RES[p.key].kind !== 'custom') return;
    const now = tot(impact(p.q, p.key));
    epdKg += now;
    if (p.generic && RES[p.generic]) { reduction += tot(impact(p.q, p.generic)) - now; withGeneric++; }
  }));
  const total = t.a13 + t.a4 + t.a5;
  return { ...t, total, reduction, hasReduction: withGeneric > 0, epdShare: total > 0 ? epdKg / total : 0, epdKg };
}

/* ---------- undo / redo: snapshots of the rows' mappings and layers ---------- */
const undoStack = [], redoStack = [];
const MAX_UNDO = 50;
function snapshot() {
  return JSON.stringify({
    rows: app.rows.map(r => ({ id: r.id, raw: r.raw, res: r.res, generic: r.generic, layers: r.layers })),
    manual: MANUAL
  });
}
function restore(json) {
  const s = JSON.parse(json);
  const idx = columnIndex();
  app.rows = s.rows.map(x => Object.assign(rowFrom(x.raw, app.headers, idx, x.id), { res: x.res, generic: x.generic, layers: x.layers }));
  Object.keys(MANUAL).forEach(k => delete MANUAL[k]);
  Object.assign(MANUAL, s.manual);
}
// Call before every change the user can undo.
export function checkpoint() {
  undoStack.push(snapshot());
  if (undoStack.length > MAX_UNDO) undoStack.shift();
  redoStack.length = 0;
  app.dirty = true;
}
export function undo() { if (!undoStack.length) return false; redoStack.push(snapshot()); restore(undoStack.pop()); app.dirty = true; return true; }
export function redo() { if (!redoStack.length) return false; undoStack.push(snapshot()); restore(redoStack.pop()); app.dirty = true; return true; }
export function clearHistory() { undoStack.length = 0; redoStack.length = 0; }
export const canUndo = () => undoStack.length > 0;
export const canRedo = () => redoStack.length > 0;
