// Climate resources: Boverket (same pinned version as before), the generic "Luft"
// resource INFINITY has always appended, EPD files from /epd and the user's own
// (EPD/manual) resources. Every resource lives in RES under a key:
//   "b<index>"  Boverket, index into the list as loaded (Luft last), the same index
//               INFINITY's project files store as climateResources / "boverket:<i>"
//   "c<n>"      an EPD or manually entered resource (kind "custom")

export const BOVERKET_API_VERSION = '02.07.000';
const API = 'https://api.boverket.se/klimatdatabas/api/Klimat/v2/GetAllResources';

const LUFT = {
  Name: 'Luft',
  Categories: [{ Text: 'Generiska resurser' }],
  Conversions: [{ Value: 1.2, Unit: 'kg/m3' }],
  WasteFactor: 1,
  Description: 'Generisk luftresurs med densitet 1,2 kg/m³',
  DataItems: [{ DataValueItems: [
    { DataModuleCode: 'A1-A3 Conservative', Value: 0 },
    { DataModuleCode: 'A4', Value: 0 },
    { DataModuleCode: 'A5.1', Value: 0 }
  ] }]
};

export const RES = {};
// Factors typed in for Boverket resources whose own factor gives no weight
// (no unit, or MJ/liter): key -> { f, u }
export const MANUAL = {};
export const status = { boverket: 'loading', version: null, published: null, latest: null, count: 0, epds: [] };

export const normUnit = (u) => String(u ?? '').replace('³', '3').replace('²', '2').trim();
export const WEIGHT_UNITS = ['kg/m3', 'kg/m2'];
export const CUSTOM_UNITS = ['kg/m3', 'kg/m2', 'kg/m', 'kg/st'];

// Same field fallbacks as INFINITY used when mapping (performClimateResourceMapping).
function fromBoverket(r, i) {
  const conv = r.Conversions && r.Conversions[0];
  const items = (r.DataItems && r.DataItems[0] && r.DataItems[0].DataValueItems) || [];
  const val = (code) => { const it = items.find(x => x.DataModuleCode === code); return it && it.Value !== undefined ? it.Value : null; };
  return {
    key: 'b' + i, kind: 'boverket', idx: i, id: r.Id ?? null,
    n: r.Name || 'Namnlös resurs',
    c: (r.Categories && r.Categories[0] && r.Categories[0].Text) || 'Övrigt',
    f: (conv && conv.Value) || r.ConservativeDataConversionFactor || r.ConversionFactor || r.Factor || null,
    u: (conv && conv.Unit) || null,
    w: r.WasteFactor || null,
    a13: val('A1-A3 Conservative'), a4: val('A4'), a5: val('A5.1'),
    desc: r.Description || ''
  };
}

export const boverketKeys = () => Object.keys(RES).filter(k => k[0] === 'b');
export const customKeys = () => Object.keys(RES).filter(k => k[0] === 'c');
export const isCustom = (key) => !!key && RES[key] && RES[key].kind === 'custom';
export const resName = (key) => (key && RES[key] ? RES[key].n : '');
export const boverketByName = (name) => boverketKeys().find(k => RES[k].n === name) || null;
export const boverketByIndex = (i) => (i !== '' && i != null && RES['b' + i] ? 'b' + i : null);
export function boverketById(id, fallbackName) {
  const k = id != null ? boverketKeys().find(x => Number(RES[x].id) === id) : null;
  return k || (fallbackName ? boverketByName(fallbackName) : null);
}

export const factorOf = (key) => MANUAL[key] || { f: RES[key].f, u: RES[key].u };
// A Boverket resource whose factor can't give a weight needs a typed factor.
export const lacksWeightUnit = (key) => RES[key].kind === 'boverket' && !WEIGHT_UNITS.includes(normUnit(RES[key].u));

// Waste factor as a multiplier. The old EPD form asked for "0.05 = 5 % spill",
// so a value below 1 is read as a share and turned into 1.05.
export const wasteMultiplier = (w) => {
  const n = Number(w);
  if (!Number.isFinite(n) || n <= 0) return null;
  return n < 1 ? 1 + n : n;
};

let customSeq = 0;
// Adds an EPD/manual resource, reusing an identical one already in the project.
export function addCustom(x) {
  const same = customKeys().find(k => {
    const r = RES[k];
    return r.n === x.n && r.f === x.f && normUnit(r.u) === normUnit(x.u) && r.w === x.w && r.a13 === x.a13 && r.a4 === x.a4 && r.a5 === x.a5;
  });
  if (same) return same;
  const key = 'c' + (++customSeq);
  RES[key] = { key, kind: 'custom', c: 'EPD och egna resurser', declUnit: x.declUnit || '', source: x.source || '', ...x };
  return key;
}
export function clearCustom() { customKeys().forEach(k => delete RES[k]); }
export function clearManual() { Object.keys(MANUAL).forEach(k => delete MANUAL[k]); }

let boverketPromise = null;
export function loadBoverket() {
  if (boverketPromise) return boverketPromise;
  const get = (url) => fetch(url).then(r => r.json()).catch(() => null);
  boverketPromise = Promise.all([get(`${API}/${BOVERKET_API_VERSION}/sv/json`), get(`${API}/latest/sv/json`)])
    .then(([data, latest]) => {
      const list = data && Array.isArray(data.Resources) ? [...data.Resources, LUFT] : [LUFT];
      list.forEach((r, i) => { RES['b' + i] = fromBoverket(r, i); });
      status.boverket = data && Array.isArray(data.Resources) ? 'ok' : 'error';
      status.version = data ? data.Version || null : null;
      status.published = data && data.PublishedDate ? new Date(data.PublishedDate).toLocaleDateString('sv-SE') : null;
      status.latest = latest ? latest.Version || null : null;
      status.count = list.length;
      // kept for anything that still reads the old globals
      window.boverketResources = list;
      window.climateResources = list;
      return status;
    });
  return boverketPromise;
}

let epdPromise = null;
// EPD files listed in epd/index.json, parsed with the same EpdParser as before.
export function loadEpds() {
  if (epdPromise) return epdPromise;
  epdPromise = (async () => {
    if (typeof window.EpdParser !== 'function') return [];
    const parser = new window.EpdParser();
    parser.setDebug(false);
    let files = [];
    try { const r = await fetch('epd/index.json'); if (r.ok) files = await r.json(); } catch { /* no list */ }
    const out = await Promise.all(files.map(async (file) => {
      try {
        const r = await fetch('epd/' + file);
        if (!r.ok) return null;
        const data = parser.parseEpdCsv(await r.text());
        return { file, name: data.name || file.replace(/\.csv$/, ''), data };
      } catch { return null; }
    }));
    status.epds = out.filter(Boolean);
    return status.epds;
  })();
  return epdPromise;
}
