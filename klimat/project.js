// Project files: the same JSON as INFINITY has always saved (version 1.2), so old
// projects open here and files saved here keep INFINITY's keys. Rows are stored as
// their original cells; layering and climate are keyed by the row's first six
// values (getRowSignature). The table is rebuilt from data on open: nothing depends
// on saved HTML, so no tableHTML is written.
import { RES, MANUAL, factorOf, normUnit, WEIGHT_UNITS, addCustom, clearCustom, clearManual,
  boverketByName, boverketByIndex, isCustom } from './resources.js';
import { app, makeRows, setRows, sigOf, rowSig, parts, impact } from './model.js';

// columns INFINITY added itself (climate results, layer names): rebuilt, never read back
const DERIVED_HEADER = /^(skiktnamn|klimatresurs|klimatpåverkan|emissionsfaktor|omräkningsfaktor|spillfaktor|inbyggd vikt|inköpt vikt)/i;
const bIndex = (key) => (key && /^b\d+$/.test(key) ? key.slice(1) : '');

function climateEntry(key, res, generic, q) {
  const x = RES[res], m = factorOf(res);
  const e = { key, resourceName: x.n, factor: m.f, unit: m.u, waste: x.w, a1a3: x.a13, a4: x.a4, a5: x.a5, isCustom: x.kind === 'custom' };
  if (x.kind === 'custom' && x.declUnit) e.declaredUnit = x.declUnit;
  if (generic && RES[generic]) {
    // what the EPD replaced, as INFINITY stored it for the reduction figure
    const g = RES[generic], gm = factorOf(generic), gi = impact(q, generic) || {};
    e.originalBoverket = { name: g.n, factor: gm.f, unit: gm.u, waste: g.w, a1a3: g.a13, a4: g.a4, a5: g.a5,
      a1a3Impact: gi.a13, a4Impact: gi.a4, a5Impact: gi.a5 };
  }
  return e;
}
const layerResource = (key) => (bIndex(key) ? `boverket:${bIndex(key)}` : isCustom(key) ? 'custom:manual' : 'none');
const layerType = (key) => (bIndex(key) ? 'boverket' : isCustom(key) ? 'custom' : 'none');

export function buildProject() {
  const layerData = [], climateData = [], seen = new Set();
  app.rows.forEach((r, idx) => {
    const sig = rowSig(r);
    if (seen.has(sig)) return; // identical rows share one entry, as in INFINITY
    seen.add(sig);
    const ps = parts(r);
    if (!r.layers) { if (r.res) climateData.push(climateEntry(sig, r.res, r.generic, ps[0].q)); return; }
    const layerKey = `row-${Date.now()}-${idx}`;
    const shared = r.layers.map((_, i) => `${layerKey}_Layer_${i + 1}`);
    const material = (name, percent, res) => ({ name, percent, climateResource: layerResource(res) });
    const mixed = r.layers.map((l, i) => (l.mix ? { layerIndex: i + 1,
      material1: material(l.name || 'Material 1', 100 - l.mix.pct, l.res),
      material2: material(l.mix.name || 'Material 2', l.mix.pct, l.mix.res) } : null)).filter(Boolean);
    layerData.push({ key: sig, count: r.layers.length, thicknesses: r.layers.map(l => Number(l.t) || 0),
      layerKey, sharedLayerKeys: shared, layerNames: r.layers.map(l => l.name || ''),
      mixedLayerConfigs: mixed.length ? mixed : undefined,
      climateResources: r.layers.map(l => bIndex(l.res) || (isCustom(l.res) ? 'manual' : '')),
      climateTypes: r.layers.map(l => layerType(l.res)),
      climateFactors: r.layers.map(() => 1) });
    ps.forEach(p => {
      if (!p.key) return;
      const i = r.layers.indexOf(p.layer);
      climateData.push(climateEntry(`${sig}::layer::${shared[i]}${p.part2 ? '_mat2' : ''}`, p.key, p.generic, p.q));
    });
  });
  const manualFactors = {};
  Object.entries(MANUAL).forEach(([k, m]) => { if (RES[k]) manualFactors[RES[k].n] = m; });
  const v = app.view;
  return {
    version: '1.2',
    timestamp: new Date().toISOString(),
    originalFileName: app.sourceFile,
    headers: [...app.headers],
    rows: app.rows.map(r => app.headers.map((h, i) => r.raw[i] ?? '')),
    layerData, climateData,
    undoStack: [], redoStack: [],
    filterValue: v.search,
    groupByValue: v.groupBy ? String(app.headers.indexOf(v.groupBy)) : '',
    // settings the new page adds; the classic page ignores unknown keys
    view: { title: app.title, groupBy: v.groupBy, cols: v.cols, filter: v.filter, bta: app.bta, manualFactors }
  };
}

// Reads a project from this page or from the classic INFINITY page.
// Returns a short summary of what was read.
export function loadProject(p, fileName) {
  if (!p || !Array.isArray(p.headers) || !Array.isArray(p.rows)) throw new Error('not a project');
  const view = p.view || {};
  const keep = p.headers.map((h, i) => [String(h ?? '').trim(), i]).filter(([h]) => h && !DERIVED_HEADER.test(h));
  const raw = p.rows.filter(row => Array.isArray(row) && keep.some(([, i]) => String(row[i] ?? '').trim() !== ''));
  const built = makeRows(keep.map(([h]) => h), raw.map(row => keep.map(([, i]) => row[i] ?? '')));
  if (!built.rows.length) throw new Error('no rows');

  // resources from the previous project go, Boverket stays
  clearCustom();
  clearManual();
  Object.entries(view.manualFactors || {}).forEach(([name, m]) => { const k = boverketByName(name); if (k) MANUAL[k] = m; });

  const climate = new Map((p.climateData || []).filter(c => c && c.key).map(c => [c.key, c]));
  const layering = new Map((p.layerData || []).filter(l => l && l.key).map(l => [l.key, l]));
  const used = new Set(), kept = new Set();
  let lost = 0;
  const fromEntry = (c) => {
    if (!c || !c.resourceName) return null;
    const k = c.isCustom ? null : boverketByName(c.resourceName);
    if (k) {
      if (!WEIGHT_UNITS.includes(normUnit(RES[k].u)) && WEIGHT_UNITS.includes(normUnit(c.unit)) && Number(c.factor) > 0)
        MANUAL[k] = { f: Number(c.factor), u: c.unit }; // the factor typed in for e.g. XPS/EPS
      return k;
    }
    // EPD, own resource or a name not in this Boverket version: keep the saved values
    kept.add(c.resourceName);
    return addCustom({ n: c.resourceName, f: c.factor ?? null, u: c.unit ?? null, w: c.waste ?? null,
      a13: c.a1a3 ?? 0, a4: c.a4 ?? 0, a5: c.a5 ?? 0, declUnit: c.declaredUnit || '', source: 'projekt',
      boverketRules: !c.isCustom });
  };
  const genericOf = (c) => (c && c.originalBoverket && c.originalBoverket.name ? boverketByName(c.originalBoverket.name) : null);
  // preferIndex: older saves can't tell the layers apart by key, so the index list decides
  const resolve = (ckey, idx, type, preferIndex = false) => {
    const c = climate.get(ckey);
    if (c) used.add(ckey);
    const byIndex = boverketByIndex(idx);
    if (preferIndex && byIndex) {
      if (c && c.resourceName === RES[byIndex].n) fromEntry(c); // picks up a typed factor
      return { res: byIndex, generic: null };
    }
    if (c && c.resourceName) return { res: fromEntry(c), generic: genericOf(c) };
    if (byIndex) return { res: byIndex, generic: null };
    if (type && type !== 'none' && type !== 'boverket') lost++;
    return { res: null, generic: null };
  };
  const idxOf = (v) => (typeof v === 'string' && v.startsWith('boverket:') ? v.slice(9) : '');

  built.rows.forEach((r, i) => {
    const sig = sigOf(raw[i]); // first six values as saved, like getRowSignature
    const L = layering.get(sig);
    if (L && L.count > 0) {
      used.add(sig);
      const mixes = L.mixedLayerConfigs || [];
      const exact = Array.isArray(L.sharedLayerKeys) && L.sharedLayerKeys.length === L.count;
      r.layers = Array.from({ length: L.count }, (_, li) => {
        const sk = exact ? L.sharedLayerKeys[li] : (L.layerKey || '');
        const mix = mixes.find(m => m.layerIndex === li + 1);
        const type = (L.climateTypes || [])[li];
        const m1 = resolve(`${sig}::layer::${sk}`, mix ? idxOf(mix.material1?.climateResource) : (type === 'boverket' ? (L.climateResources || [])[li] : ''), type, !exact);
        const layer = { name: (L.layerNames || [])[li] || '', t: (L.thicknesses || [])[li] ?? '', res: m1.res, generic: m1.generic, mix: null };
        if (mix) {
          const m2 = resolve(`${sig}::layer::${sk}_mat2`, idxOf(mix.material2?.climateResource), mix.material2?.climateResource?.split(':')[0], !exact);
          // "Material 2" is the placeholder INFINITY writes for an unnamed material
          layer.mix = { name: mix.material2?.name === 'Material 2' ? '' : (mix.material2?.name || ''), pct: Number(mix.material2?.percent) || 0, res: m2.res, generic: m2.generic };
        }
        return layer;
      });
    } else {
      const m = resolve(sig, '', null);
      r.res = m.res; r.generic = m.generic;
    }
  });
  const orphans = [...climate.keys(), ...layering.keys()].filter(k => !used.has(k) && !used.has(k.split('::layer::')[0])).length;
  setRows(built);

  const fromApp = app.headers[Number(p.groupByValue)];
  const groupBy = [view.groupBy, p.groupByValue === '' && !view.groupBy ? '' : fromApp,
    app.groupColumns.find(h => h.toLowerCase() === 'type'), app.groupColumns[0]].find(h => h === '' || (h && app.groupColumns.includes(h)));
  Object.assign(app.view, { filter: view.filter || 'all', search: p.filterValue || '', ticked: new Set(), drawer: null,
    groupBy: groupBy ?? '', cols: view.cols || app.view.cols, open: new Set() });
  if (view.bta != null && view.bta !== '') app.bta = String(view.bta);
  app.sourceFile = p.originalFileName || app.sourceFile;
  app.title = view.title || fileName.replace(/(_projekt)?\.[^.]+$/, '');

  const mapped = app.rows.filter(r => r.layers ? r.layers.every(l => l.res && (!l.mix || l.mix.res)) : r.res).length;
  const layered = app.rows.filter(r => r.layers).length;
  const notes = [`Öppnade ${app.rows.length} rader: ${mapped} mappade, ${layered} skiktade.`];
  if (kept.size) notes.push(`${kept.size} EPD/egna resurser behölls med sparade värden.`);
  if (lost) notes.push(`${lost} skikt hade en EPD utan sparade värden och står som ej mappade.`);
  if (orphans) notes.push(`${orphans} sparade mappningar hittade ingen rad.`);
  return notes.join(' ');
}
