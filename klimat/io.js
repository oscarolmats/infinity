// Reading quantity take-offs and writing the Excel export.
import { RES, factorOf, isCustom, wasteMultiplier } from './resources.js';
import { app, parts, impact, totals, tot, parseNum } from './model.js';

// Excel goes through the server (/api/upload) as before, so cell text - and with it
// the row signatures saved projects are keyed on - stays exactly the same.
// Without the server (a static host) the file is read in the browser instead.
export async function readSpreadsheet(file) {
  const ext = (file.name.split('.').pop() || '').toLowerCase();
  if (ext === 'csv' || ext === 'tsv' || ext === 'txt') return parseDelimited(await file.text());
  try {
    const form = new FormData();
    form.append('file', file);
    const res = await fetch('/api/upload', { method: 'POST', body: form });
    if (!res.ok) throw new Error('upload failed');
    const doc = new DOMParser().parseFromString(await res.text(), 'text/html');
    const headers = [...doc.querySelectorAll('thead th')].map(th => th.textContent);
    const rows = [...doc.querySelectorAll('tbody tr')].map(tr => [...tr.children].map(td => td.textContent));
    return [headers, ...rows];
  } catch (err) {
    if (!window.XLSX) throw err;
    const wb = window.XLSX.read(new Uint8Array(await file.arrayBuffer()), { type: 'array' });
    const table = window.XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]], { header: 1, defval: '', raw: true });
    return table.map(row => row.map(c => (c === '' || c == null ? '' : String(c))));
  }
}

// Same delimiter guess as before: comma, semicolon or tab, whichever is most common.
export function parseDelimited(text) {
  const lines = text.replace(/\r\n?/g, '\n').split('\n').filter(l => l.length > 0);
  if (!lines.length) return [];
  const sample = lines.slice(0, 10).join('\n');
  const count = (re) => (sample.match(re) || []).length;
  const comma = count(/,/g), semicolon = count(/;/g), tab = count(/\t/g);
  const d = comma >= semicolon && comma >= tab ? ',' : (semicolon >= tab ? ';' : '\t');
  return lines.map(line => line.split(d).map(c => c.trim()));
}

/* Export: built from the data, not from the screen, so the column view, grouping and
   collapsed rows don't change the file. One row per material or layer (a mixed layer
   gives one row per material); header names match INFINITY's columns. */
const CLIMATE_HEADERS = ['Klimatresurs', 'Klimatresurs typ', 'Omräkningsfaktor', 'Omräkningsfaktor enhet', 'Spillfaktor',
  'Emissionsfaktor A1-A3', 'Emissionsfaktor A4', 'Emissionsfaktor A5',
  'Inbyggd vikt', 'Inköpt vikt', 'Klimatpåverkan A1-A3', 'Klimatpåverkan A4', 'Klimatpåverkan A5'];
// quantities and results get two decimals in Excel; factors keep full precision
const TWO_DECIMALS = ['net area', 'length', 'volume', 'inbyggd vikt', 'inköpt vikt', 'klimatpåverkan a1-a3', 'klimatpåverkan a4', 'klimatpåverkan a5'];

function dataRows(source) {
  const out = [];
  const textCols = new Set(app.groupColumns);
  const col = (name) => app.headers.find(h => h.toLowerCase() === name);
  const volH = col('volume'), areaH = col('net area'), thickH = col('thickness');
  const fileValues = (r, over) => app.headers.map((h, i) => {
    if (over && h in over) return over[h];
    const v = r.raw[i] ?? '';
    if (textCols.has(h)) return v;
    const n = parseNum(v);
    return String(v).trim() === '' ? '' : (Number.isFinite(n) ? n : v);
  });
  const climate = (key, q) => {
    if (!key || !RES[key]) return CLIMATE_HEADERS.map(() => '');
    const x = RES[key], m = factorOf(key), i = impact(q, key), w = i && !i.noWeight;
    return [x.n, isCustom(key) ? 'EPD / egen resurs' : 'Generisk klimatresurs', m.f ?? '', m.u ?? '', wasteMultiplier(x.w) ?? '',
      x.a13 ?? '', x.a4 ?? '', x.a5 ?? '',
      w ? i.w : '', w ? i.bought : '', w ? i.a13 : '', w ? i.a4 : '', w ? i.a5 : ''];
  };
  source.forEach(r => {
    if (!r.layers) { out.push([...fileValues(r), '', ...climate(r.res, parts(r)[0].q)]); return; }
    parts(r).forEach(p => {
      const l = p.layer, li = r.layers.indexOf(l);
      const n1 = l.name || `Skikt ${li + 1}`;
      const name = p.part2 ? ((l.mix && l.mix.name) || `${n1}, material 2`) : n1;
      const over = {};
      if (volH) over[volH] = p.q.volume;
      if (areaH) over[areaH] = p.q.area;
      if (thickH && l.t) over[thickH] = l.t / 1000;
      out.push([...fileValues(r, over), name, ...climate(p.key, p.q)]);
    });
  });
  return out;
}

function summaryRows(source, scopeLabel) {
  const t = totals(source);
  const bta = parseFloat(app.bta);
  const by = app.view.groupBy || app.groupColumns[0] || '';
  const groups = new Map();
  source.forEach(r => {
    const k = by ? (String(r.cells[by] ?? '').trim() || '(tom)') : 'Alla';
    const g = groups.get(k) || { rows: 0, kg: 0 };
    g.rows++; g.kg += tot(r.impact); groups.set(k, g);
  });
  const kg = { 1: '0.00' };
  // [row values, number formats per column]
  return [
    [['INFINITY klimatanalys']],
    [['Projekt', app.title]],
    [['Fil', app.sourceFile]],
    [['Exporterad', new Date().toLocaleString('sv-SE')]],
    [['Urval', scopeLabel]],
    [[]],
    [['Klimatpåverkan', 'kg CO₂e']],
    [['A1-A3', t.a13], kg], [['A4', t.a4], kg], [['A5', t.a5], kg], [['Totalt A1-A5', t.total], kg],
    ...(t.hasReduction ? [[['Reduktion med EPD', t.reduction], kg]] : []),
    [[]],
    [['BTA (m²)', bta > 0 ? bta : '']],
    [['Per m² BTA (kg CO₂e/m²)', bta > 0 ? t.total / bta : ''], kg],
    [[]],
    [['Mappade rader', source.filter(r => r.status === 'ok').length, 'av', source.length]],
    [[]],
    [[`Per ${by || 'rad'}`, 'Rader', 'A1-A5 (kg CO₂e)', 'Andel']],
    ...[...groups].sort((a, b) => b[1].kg - a[1].kg).map(([k, g]) => [[k, g.rows, g.kg, t.total ? g.kg / t.total : 0], { 2: '0.00', 3: '0.0%' }])
  ];
}

// Returns { filename, blob, rows } or null when there is nothing to export.
export function buildExport(source, scopeLabel) {
  const X = window.XLSX;
  const rows = dataRows(source);
  if (!rows.length) return null;
  const headers = [...app.headers, 'Skiktnamn', ...CLIMATE_HEADERS];
  const ws = X.utils.aoa_to_sheet([headers, ...rows]);
  headers.forEach((h, c) => {
    if (!TWO_DECIMALS.includes(h.toLowerCase())) return;
    for (let r = 1; r <= rows.length; r++) {
      const cell = ws[X.utils.encode_cell({ r, c })];
      if (cell && cell.t === 'n') cell.z = '0.00';
    }
  });
  const sum = summaryRows(source, scopeLabel);
  const summary = X.utils.aoa_to_sheet(sum.map(([values]) => values));
  sum.forEach(([, formats], r) => Object.entries(formats || {}).forEach(([c, z]) => {
    const cell = summary[X.utils.encode_cell({ r, c: Number(c) })];
    if (cell && cell.t === 'n') cell.z = z;
  }));
  const wb = X.utils.book_new();
  X.utils.book_append_sheet(wb, ws, 'Data');
  X.utils.book_append_sheet(wb, summary, 'Sammanfattning');
  const base = (app.title || 'export').replace(/[\\/:*?"<>|]/g, '').trim() || 'export';
  const filename = `${base}_klimatanalys_${new Date().toISOString().slice(0, 10)}.xlsx`;
  const blob = new Blob([X.write(wb, { bookType: 'xlsx', type: 'array' })], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
  return { filename, blob, rows: rows.length };
}

export function download(blob, filename) {
  const a = document.createElement('a');
  a.href = URL.createObjectURL(blob);
  a.download = filename;
  document.body.appendChild(a);
  a.click();
  a.remove();
  setTimeout(() => URL.revokeObjectURL(a.href), 1000);
}
