// Rendering and interaction for the Klimatanalys page. Everything is drawn from the
// model (model.js); every change goes model -> derive() -> render.
import { RES, MANUAL, status as resStatus, resName, isCustom, factorOf, lacksWeightUnit, boverketKeys, customKeys,
  boverketById, addCustom, normUnit, wasteMultiplier, loadEpds } from './resources.js';
import { app, derive, totals, tot, parseNum, checkpoint, undo, redo, canUndo, canRedo, addRow, layerQty } from './model.js';
import { buildExport, download } from './io.js';

export const $ = (id) => document.getElementById(id);
export const esc = (s) => String(s ?? '').replace(/[&<>"]/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]));
// one decimal with a point, as elsewhere in INFINITY
const f = (n, d = 1) => Number(n).toLocaleString('sv-SE', { minimumFractionDigits: d, maximumFractionDigits: d }).replace(',', '.');
const fNum = (v) => { if (String(v ?? '').trim() === '') return ''; const n = parseNum(v); return Number.isFinite(n) ? n.toLocaleString('sv-SE', { maximumFractionDigits: 3 }).replace(',', '.') : esc(v); };
const clone = (o) => JSON.parse(JSON.stringify(o));
const STATUS_LABEL = { ok: 'Mappad', part: 'Delvis', none: 'Ej mappad' };
const pill = (s) => `<span class="pill ${s}">${STATUS_LABEL[s]}</span>`;
const rowLabel = (r) => (r.desc ? `${r.type} · ${r.desc}` : r.type);
const groupKey = (r) => String(r.cells[app.view.groupBy] ?? '').trim() || '(tom)';
const rowById = (rid) => app.rows.find(x => x.id === Number(rid));
const usesCustom = (r) => (r.layers ? r.layers.some(l => isCustom(l.res) || (l.mix && isCustom(l.mix.res))) : isCustom(r.res));
const epdTag = (key) => (isCustom(key) ? '<span class="epd-tag" title="EPD eller egen resurs">EPD</span>' : '');

let handlers = {};
let changed = () => {};

/* ---------- toast ---------- */
let toastTimer;
export function toast(text, ms = 3200) {
  $('toast').textContent = text; $('toast').hidden = false;
  clearTimeout(toastTimer); toastTimer = setTimeout(() => { $('toast').hidden = true; }, ms);
}

/* ---------- columns ---------- */
const FILTERS = [
  { key: 'all', label: 'Alla', test: () => true },
  { key: 'unmapped', label: 'Ej mappade', test: r => r.status !== 'ok' },
  { key: 'mapped', label: 'Mappade', test: r => r.status === 'ok' },
  { key: 'layered', label: 'Skiktade', test: r => !!r.layers },
  { key: 'epd', label: 'EPD', test: usesCustom, onlyWhenUsed: true }
];
const COLSETS = [
  { key: 'overview', label: 'Översikt' },
  { key: 'qty', label: 'Mängder' },
  { key: 'climate', label: 'Klimat' },
  { key: 'all', label: 'Alla' }
];
const resCell = (r) => {
  if (r.layers) {
    // a layered row is mapped per layer: clicking shows the layers to map them in place
    const missing = r.layerResults.filter(l => !l.mapped).length;
    return `<button class="cellbtn${missing ? ' res-missing' : ''}" type="button" data-act="toggle-layers" data-id="${r.id}" title="Visa skikten och mappa dem">${missing ? `${missing} av ${r.layers.length} skikt saknar` : 'Per skikt'}</button>`;
  }
  return `${epdTag(r.res)}<button class="cellbtn${r.res ? '' : ' res-missing'}" type="button" data-act="map" data-id="${r.id}" title="${r.res ? 'Byt klimatresurs' : 'Mappa klimatresurs'}">${r.res ? esc(resName(r.res)) : 'Saknas – välj'}</button>`;
};
// status is clickable when something is left to do
const statusCell = (r) => {
  if (r.status === 'ok') return pill('ok');
  const act = r.layers ? `data-act="toggle-layers" data-id="${r.id}" title="Visa skikten och mappa dem"` : `data-act="map" data-id="${r.id}" title="Mappa klimatresurs"`;
  return `<button type="button" class="pill ${r.status}" ${act}>${STATUS_LABEL[r.status]}</button>`;
};
const STATUS_COL = { kind: 'status', h: 'Status', cell: statusCell };
const RES_COL = { kind: 'res', h: 'Klimatresurs', climate: true, cell: resCell };
const impactCol = (h, get, extra = {}) => ({ kind: 'climate', h, r: true, climate: true, d: 0,
  cell: r => (r.impact ? f(get(r.impact), 0) : '–'), sum: r => (r.impact ? get(r.impact) : 0), ...extra });
const WEIGHT_COL = impactCol('Inbyggd vikt', i => i.w, { unit: 'kg' });
const BOUGHT_COL = impactCol('Inköpt vikt', i => i.bought, { unit: 'kg' });
const PHASE_COLS = [impactCol('A1–A3', i => i.a13), impactCol('A4', i => i.a4), impactCol('A5', i => i.a5)];
const TOTAL_COL = impactCol('A1–A5', i => tot(i), { unit: 'kg CO₂e', share: true });
// Only these quantities make sense to add up on a group row (not Thickness, Height …).
const SUMMABLE = ['net area', 'area', 'length', 'längd', 'volume', 'volym', 'count', 'antal'];
const fileCol = (h) => {
  const numeric = !app.groupColumns.includes(h), low = h.toLowerCase();
  return { kind: 'file', h, r: numeric, numeric,
    cell: r => (numeric ? fNum(r.cells[h]) : esc(r.cells[h] ?? '')),
    sum: numeric && SUMMABLE.includes(low) ? (r => { const n = parseNum(r.cells[h]); return Number.isFinite(n) ? n : 0; }) : null,
    d: low === 'count' || low === 'antal' ? 0 : 2 };
};
function columnsFor(set) {
  const files = app.headers.filter(Boolean).map(fileCol);
  const pick = (...names) => names.map(n => files.find(c => c.h.toLowerCase() === n)).filter(Boolean);
  if (set === 'overview') return [STATUS_COL, ...pick('material', 'volume', 'volym'), RES_COL, TOTAL_COL];
  if (set === 'qty') return [STATUS_COL, ...files.filter(c => c.numeric)];
  if (set === 'climate') return [STATUS_COL, RES_COL, WEIGHT_COL, BOUGHT_COL, ...PHASE_COLS, TOTAL_COL];
  return [STATUS_COL, ...files, RES_COL, WEIGHT_COL, BOUGHT_COL, ...PHASE_COLS, TOTAL_COL];
}

/* ---------- overview ---------- */
// With "(ingen)" grouping there are no groups to jump to, so rank by the first text column.
const topKey = (r) => (app.view.groupBy ? groupKey(r) : (String(r.cells[app.groupColumns[0]] ?? '').trim() || '(tom)'));
function renderOverview() {
  const t = totals();
  const total = t.total;
  $('kTotal').innerHTML = `${f(total / 1000)}<small>ton CO₂e</small>`;
  const ph = [['A1–A3', t.a13, '--a13'], ['A4', t.a4, '--a4'], ['A5', t.a5, '--a5']];
  $('kStack').innerHTML = ph.map(([, v, c]) => `<i style="width:${total ? v / total * 100 : 0}%;background:var(${c})"></i>`).join('');
  $('kLegend').innerHTML = ph.map(([l, v, c]) => `<span style="--c:var(${c})">${l} <b class="num">${f(v / 1000)} t</b></span>`).join('');
  const extra = [];
  if (t.epdKg > 0) extra.push(`<span>EPD <b class="num">${f(t.epdShare * 100)} %</b> av påverkan</span>`);
  if (t.hasReduction) {
    const base = total + t.reduction;
    extra.push(t.reduction >= 0
      ? `<span class="good">Reduktion med EPD <b class="num good">${f(t.reduction / 1000)} t</b> (${f(base ? t.reduction / base * 100 : 0)} %)</span>`
      : `<span>EPD ger <b class="num">${f(-t.reduction / 1000)} t</b> mer än generiska värden</span>`);
  }
  $('kExtra').innerHTML = extra.join('');
  const bta = parseFloat(app.bta);
  $('kPerM2').innerHTML = bta > 0 ? `${f(total / bta)}<small>kg CO₂e/m²</small>` : '–<small>ange BTA</small>';
  if (document.activeElement !== $('btaInput')) $('btaInput').value = app.bta;
  const mapped = app.rows.filter(r => r.status === 'ok').length;
  $('kMapped').innerHTML = `${mapped}<small>av ${app.rows.length} rader</small>`;
  $('kMappedBar').style.width = `${app.rows.length ? mapped / app.rows.length * 100 : 0}%`;
  const left = app.rows.length - mapped;
  $('showUnmapped').textContent = left ? `Visa ${left} som återstår` : 'Allt är mappat';
  // follows the active grouping, so a click lands on a group that is actually in the table
  const byGroup = new Map();
  app.rows.forEach(r => { const k = topKey(r); byGroup.set(k, (byGroup.get(k) || 0) + tot(r.impact)); });
  const top = [...byGroup].filter(([, v]) => v > 0).sort((a, b) => b[1] - a[1]).slice(0, 4);
  $('kTopTitle').textContent = `Störst påverkan per ${(app.view.groupBy || app.groupColumns[0] || 'rad').toLowerCase()}`;
  $('kTop').innerHTML = top.length ? top.map(([g, v]) =>
    `<button type="button" data-group="${esc(g)}"><span class="tname">${esc(g)}</span><span class="tval num">${f(v / total * 100, 0)} %</span><span class="tbar"><i style="width:${v / total * 100}%"></i></span></button>`
  ).join('') : '<span class="d" style="margin:0">Visas när något är mappat</span>';
}

function renderChips() {
  $('chips').innerHTML = FILTERS.map(x => {
    const n = app.rows.filter(x.test).length;
    if (x.onlyWhenUsed && !n && app.view.filter !== x.key) return '';
    return `<button class="chip" type="button" data-filter="${x.key}" aria-pressed="${app.view.filter === x.key}">${x.label}<b class="num">${n}</b></button>`;
  }).join('');
  $('colsets').innerHTML = COLSETS.map(c =>
    `<button type="button" data-cols="${c.key}" aria-pressed="${app.view.cols === c.key}">${c.label}</button>`).join('');
}

export function visibleRows() {
  const q = app.view.search.trim().toLowerCase();
  const test = (FILTERS.find(x => x.key === app.view.filter) || FILTERS[0]).test;
  return app.rows.filter(r => test(r) && (!q || [...Object.values(r.cells), r.res ? resName(r.res) : ''].some(v => String(v).toLowerCase().includes(q))));
}
const rowsInGroup = (g) => visibleRows().filter(r => groupKey(r) === g);

function groupStatus(rows) {
  const ok = rows.filter(r => r.status === 'ok').length;
  return ok === rows.length ? 'ok' : (ok || rows.some(r => r.status === 'part')) ? 'part' : 'none';
}
function rowActions(r) {
  const map = r.layers ? '' : `<button class="rowbtn${r.res ? ' quiet' : ''}" type="button" data-act="map" data-id="${r.id}">${r.res ? 'Byt' : 'Mappa'}</button>`;
  const copy = r.layers ? `<button class="rowbtn quiet" type="button" data-act="copy-layers" data-id="${r.id}">Kopiera skikt</button>` : '';
  return `${map}<button class="rowbtn quiet" type="button" data-act="layer" data-id="${r.id}">${r.layers ? 'Redigera skikt' : 'Skikta'}</button>${copy}`;
}

/* ---------- table ---------- */
function renderTable() {
  const v = app.view;
  const cols = columnsFor(v.cols);
  const rows = visibleRows();
  const total = totals().total;
  const allTicked = rows.length && rows.every(r => v.ticked.has(r.id));
  const head = `<thead><tr><th class="cb"><input type="checkbox" data-act="tick-all" aria-label="Markera alla synliga"${allTicked ? ' checked' : ''}></th><th class="obj">Objekt</th>${cols.map(c => `<th class="${c.r ? 'r' : ''}${c.climate ? ' climate' : ''}">${esc(c.h)}${c.unit ? ` <span style="text-transform:none">(${c.unit})</span>` : ''}</th>`).join('')}<th class="acts"><span class="sr">Åtgärder</span></th></tr></thead>`;

  // long text (e.g. Material lists from IFC) is cut with … and shown in full on hover
  const txtTd = (c, text, html) => `<td class="txt" title="${esc(text)}">${html}</td>`;
  const itemRow = (r) => {
    const cells = cols.map(c => {
      let val = c.cell(r);
      if (c.kind === 'file' && !c.numeric) return txtTd(c, r.cells[c.h] ?? '', val);
      if (c.share && r.impact && total) val += `<span class="share"><i style="width:${tot(r.impact) / total * 100}%"></i></span>`;
      return `<td class="${c.r ? 'r num' : ''}">${val}</td>`;
    }).join('');
    let html = `<tr class="item${v.ticked.has(r.id) ? ' ticked' : ''}" data-id="${r.id}">
      <td class="cb"><input type="checkbox" data-act="tick" data-id="${r.id}" aria-label="Markera ${esc(rowLabel(r))}"${v.ticked.has(r.id) ? ' checked' : ''}></td>
      <td class="obj" title="${esc(rowLabel(r))}"><button class="name" type="button" data-act="details" data-id="${r.id}"><span class="t">${esc(r.type)}</span>${r.desc ? `<span class="d">${esc(r.desc)}</span>` : ''}</button>${r.layers ? `<button class="tag" type="button" data-act="toggle-layers" data-id="${r.id}" aria-expanded="${v.open.has('L' + r.id)}">${r.layers.length} skikt</button>` : ''}</td>
      ${cells}<td class="acts">${rowActions(r)}</td></tr>`;
    if (r.layers && v.open.has('L' + r.id)) {
      html += r.layerResults.map((l, li) => {
        const cellsL = cols.map(c => {
          let val = '';
          const low = c.h.toLowerCase();
          if (c.kind === 'status') val = pill(l.mapped ? 'ok' : (l.res || (l.mix && l.mix.res)) ? 'part' : 'none');
          else if (c.kind === 'file' && low === 'material') val = esc(l.name);
          else if (c.kind === 'file' && (low === 'volume' || low === 'volym')) val = f(l.volume, 2);
          else if (c.kind === 'file' && low === 'thickness') val = l.t ? fNum(l.t / 1000) : '';
          else if (c.kind === 'res') {
            const btn = (key, part) => `${epdTag(key)}<button class="cellbtn${key ? '' : ' res-missing'}" type="button" data-act="map-layer" data-id="${r.id}" data-layer="${li}" data-part="${part}">${key ? esc(resName(key)) : 'Saknas – välj'}</button>`;
            val = btn(l.res, 1) + (l.mix ? ` + ${btn(l.mix.res, 2)}` : '');
          } else if (c.kind === 'climate') val = c.cell({ impact: l.impact });
          return `<td class="${c.r ? 'r num' : ''}">${val}</td>`;
        }).join('');
        return `<tr class="layer"><td class="cb"></td><td class="obj">${esc(l.name)}${l.mix ? ` + ${esc(l.mix.name || 'material 2')} (${l.mix.pct} %)` : ''} · ${l.t ? `${fNum(l.t)} mm` : 'jämn fördelning'}</td>${cellsL}<td class="acts"></td></tr>`;
      }).join('');
    }
    return html;
  };

  let body = '';
  if (!rows.length) {
    body = `<tr><td colspan="${cols.length + 3}" class="empty">Inga rader matchar. Prova ett annat filter eller en annan sökning.</td></tr>`;
  } else if (!v.groupBy) {
    body = rows.map(itemRow).join('');
  } else {
    const groups = new Map();
    rows.forEach(r => { const k = groupKey(r); if (!groups.has(k)) groups.set(k, []); groups.get(k).push(r); });
    groups.forEach((list, g) => {
      const open = v.open.has(g);
      const ticked = list.every(r => v.ticked.has(r.id));
      const sumCells = cols.map(c => {
        if (c.kind === 'status') return `<td>${pill(groupStatus(list))}</td>`;
        if (c.kind === 'res') {
          const missing = list.filter(r => r.status !== 'ok').length;
          const names = [...new Set(list.filter(r => r.res).map(r => resName(r.res)))];
          const text = missing ? `${missing} av ${list.length} saknar` : names.length > 2 ? `${names.length} olika resurser` : (names.join(', ') || 'Per skikt');
          const title = missing ? 'Mappa gruppens omappade rader' : 'Byt klimatresurs för hela gruppen';
          return `<td><button class="cellbtn${missing ? ' res-missing' : ''}" type="button" data-act="map-group-res" data-group="${esc(g)}" title="${title}">${esc(text)}</button></td>`;
        }
        if (c.sum) {
          const s = list.reduce((a, r) => a + c.sum(r), 0);
          let val = s ? f(s, c.d ?? 2) : '–';
          if (c.share && s && total) val += `<span class="share"><i style="width:${s / total * 100}%"></i></span>`;
          return `<td class="r num">${val}</td>`;
        }
        if (c.kind === 'file' && !c.numeric) {
          // shared value for the whole group, or how many different ones there are
          const values = [...new Set(list.map(r => r.cells[c.h] || ''))];
          return values.length === 1 ? txtTd(c, values[0], esc(values[0])) : `<td><span class="d" style="margin:0">${values.length} olika</span></td>`;
        }
        return '<td></td>';
      }).join('');
      const hasUnmapped = list.some(r => !r.layers && !r.res);
      body += `<tr class="grp${open ? ' open' : ''}" data-group="${esc(g)}">
        <td class="cb"><input type="checkbox" data-act="tick-group" data-group="${esc(g)}" aria-label="Markera gruppen ${esc(g)}"${ticked ? ' checked' : ''}></td>
        <td class="obj" title="${esc(g)}"><button class="chev" type="button" data-act="toggle-group" data-group="${esc(g)}" aria-expanded="${open}" aria-label="Fäll ut ${esc(g)}">›</button>${esc(g)} <span class="d">(${list.length})</span></td>
        ${sumCells}
        <td class="acts">${hasUnmapped ? `<button class="rowbtn" type="button" data-act="map-group" data-group="${esc(g)}">Mappa grupp</button>` : ''}<button class="rowbtn quiet" type="button" data-act="layer-group" data-group="${esc(g)}">Skikta grupp</button></td></tr>`;
      if (open) body += list.map(itemRow).join('');
    });
  }
  $('table').innerHTML = head + `<tbody>${body}</tbody>`;
  $('footInfo').textContent = `${rows.length} av ${app.rows.length} rader visas`;
  const n = v.ticked.size;
  $('bulk').hidden = n === 0;
  $('bulkPaste').hidden = !v.copied;
  $('bulkCount').textContent = n === 1 ? '1 rad markerad' : `${n} rader markerade`;
  sizeTable();
}

// Table box height = window minus the sticky top bar and the table's own toolbar,
// bulk bar and footer, so the whole table (header row included) fits on screen.
export function sizeTable() {
  const panel = document.querySelector('.panel');
  if (!panel || $('workspace').hidden) return;
  const chrome = ['.toolbar', '#bulk', '.tfoot'].reduce((h, sel) => h + (panel.querySelector(sel)?.offsetHeight || 0), 0);
  const avail = window.innerHeight - $('topbar').offsetHeight - chrome - 64;
  document.querySelector('.tablebox').style.maxHeight = `${Math.max(240, avail)}px`;
}

function renderGroupOptions() {
  $('groupBy').innerHTML = app.groupColumns.map(h => `<option value="${esc(h)}">${esc(h)}</option>`).join('') + '<option value="">(ingen)</option>';
  $('groupBy').value = app.view.groupBy;
}
function renderCopied() {
  const c = app.view.copied;
  $('copied').hidden = !c;
  if (c) $('copiedText').textContent = `Kopierat: ${c.label}, ${c.layers.length} skikt`;
}
// header: title, file, save/undo buttons
export function renderChrome() {
  const has = app.rows.length > 0;
  $('start').hidden = has;
  $('workspace').hidden = !has;
  $('docTitle').textContent = has ? (app.title || 'Namnlöst projekt') + (app.dirty ? ' •' : '') : 'Ingen fil öppnad';
  $('docSub').textContent = has ? [app.sourceFile, `${app.rows.length} rader`, app.fileHandle ? `sparas till ${app.fileHandle.name}` : ''].filter(Boolean).join(' · ')
    : 'Öppna en mängdavtagning eller ett sparat projekt';
  document.title = has ? `${app.title || 'Projekt'} – INFINITY` : 'INFINITY';
  document.querySelectorAll('[data-needs-rows]').forEach(b => { b.disabled = !has; });
  $('undoBtn').disabled = !canUndo();
  $('redoBtn').disabled = !canRedo();
}
export function renderStatus() {
  const s = resStatus, el = $('resStatus');
  if (s.boverket === 'loading') { el.className = 'res-status'; el.textContent = 'Ansluter till Boverket…'; return; }
  el.className = 'res-status ' + (s.boverket === 'ok' ? 'ok' : 'error');
  el.innerHTML = s.boverket === 'ok'
    ? `Boverket ${esc(s.version || '')}${s.latest && s.latest !== s.version ? `<span class="upd" title="Det finns en nyare version av klimatdatabasen">⚠ ${esc(s.latest)} finns</span>` : ''}`
    : 'Boverket kunde inte laddas';
  el.title = s.boverket === 'ok' ? `${s.count} resurser${s.published ? `, uppdaterad ${s.published}` : ''}${s.epds.length ? `. ${s.epds.length} EPD-filer.` : ''}` : 'Klimatdatabasen gick inte att nå. Bara Luft finns att mappa.';
}

export function renderAll() {
  derive();
  renderChrome();
  if (!app.rows.length) return;
  renderGroupOptions();
  renderOverview(); renderChips(); renderCopied(); renderTable();
  if (app.view.drawer) fillDrawer(app.view.drawer);
}
// after a change the user made
function commit() { renderAll(); changed(); }

/* ---------- mapping helpers ---------- */
// Assigns a resource; an EPD remembers the Boverket resource it replaced (for the reduction).
function setRes(obj, key) {
  const old = obj.res;
  if (isCustom(key)) obj.generic = old && !isCustom(old) ? old : (obj.generic || null);
  else obj.generic = null;
  obj.res = key;
}
function clearRes(obj) { obj.res = null; obj.generic = null; }

/* ---------- catalogue: one place for row, group, bulk, table-layer and layer-modal mapping ---------- */
// target: { rows, label, clearTicks } | { row, layer, part, label } | { draft, part, label }
let catTarget = null, catTab = 'boverket', epdSel = null;
const catLayer = (t) => (t.draft != null ? lmRows[t.draft] : t.row.layers[t.layer]);
const catObj = (t) => { const l = catLayer(t); return t.part === 2 ? l.mix : l; };
function targetCurrent(t) {
  if (t.rows) {
    const keys = [...new Set(t.rows.filter(r => !r.layers).map(r => r.res))];
    return keys.length === 1 ? keys[0] : null;
  }
  return catObj(t).res;
}
function targetHasResource(t) {
  if (t.rows) return t.rows.some(r => !r.layers && r.res);
  return !!catObj(t).res;
}
function openCatalog(target) {
  catTarget = target;
  $('catTarget').textContent = target.label;
  $('catSearch').value = ''; $('epdSearch').value = '';
  $('catMsg').textContent = '';
  const cur = targetCurrent(target);
  renderCatList('');
  renderEpdList('');
  if (cur && isCustom(cur)) { setCatTab('epd'); $('epdList').value = 'c:' + cur; pickEpd('c:' + cur); }
  else {
    setCatTab('boverket');
    if (cur) $('catList').value = cur;
    showCatDetails($('catList').value);
    clearCustomForm();
  }
  $('catClear').hidden = !targetHasResource(target);
  $('catalog').hidden = false;
  (catTab === 'epd' ? $('epdSearch') : $('catSearch')).focus();
  if (!boverketKeys().length) $('catMsg').textContent = 'Boverkets klimatdatabas laddas fortfarande…';
}
function closeCatalog() { $('catalog').hidden = true; catTarget = null; }
function setCatTab(tab) {
  catTab = tab;
  $('catTabs').querySelectorAll('button').forEach(b => b.setAttribute('aria-pressed', String(b.dataset.tab === tab)));
  $('catBoverket').hidden = tab !== 'boverket';
  $('catEpd').hidden = tab !== 'epd';
  $('catMsg').textContent = '';
  $('catUse').disabled = tab === 'boverket' ? !$('catList').value : false;
}
function renderCatList(q) {
  q = q.trim().toLowerCase();
  const byCat = new Map();
  boverketKeys().forEach(k => {
    const x = RES[k];
    if (q && !x.n.toLowerCase().includes(q) && !x.c.toLowerCase().includes(q)) return;
    if (!byCat.has(x.c)) byCat.set(x.c, []);
    byCat.get(x.c).push(k);
  });
  const groups = [...byCat].sort((a, b) => a[0].localeCompare(b[0], 'sv'));
  $('catList').innerHTML = groups.map(([c, keys]) =>
    `<optgroup label="${esc(c || 'Övrigt')}">${keys.sort((a, b) => RES[a].n.localeCompare(RES[b].n, 'sv')).map(k => `<option value="${k}">${esc(RES[k].n)}</option>`).join('')}</optgroup>`).join('');
  const hits = groups.reduce((s, [, keys]) => s + keys.length, 0);
  $('catCount').textContent = hits ? `${hits} av ${boverketKeys().length} resurser` : 'Inga resurser matchar sökningen';
  showCatDetails($('catList').value);
}
function showCatDetails(key) {
  if (catTab === 'boverket') $('catUse').disabled = !key;
  if (!key) { $('catDetails').innerHTML = '<p class="d" style="margin:0">Välj en resurs för att se detaljer.</p>'; return; }
  const x = RES[key];
  const row = (k, v) => `<dt>${k}</dt><dd class="num">${v}</dd>`;
  const ef = (v) => (v != null ? `${fNum(v)} kg CO₂e/kg` : '–');
  let html = `<h3>${esc(x.n)}</h3><dl class="kv">${row('Kategori', esc(x.c || '–'))}${row('Omräkningsfaktor', esc(x.f != null ? `${fNum(x.f)} ${x.u || '(enhet saknas)'}` : 'Ej angiven'))}${row('Spillfaktor', x.w != null ? fNum(x.w) : '–')}${row('A1–A3', ef(x.a13))}${row('A4', ef(x.a4))}${row('A5', ef(x.a5))}</dl>`;
  if (x.desc) html += `<p class="desc">${esc(x.desc)}</p>`;
  if (lacksWeightUnit(key)) {
    const m = MANUAL[key] || {};
    html += `<div class="manual"><p>Omräkningsfaktorn ${x.u ? `anges i ${esc(x.u)}` : 'saknar enhet'}, så vikten kan inte räknas fram. Ange en egen omräkningsfaktor:</p>
      <div class="manual-row"><input type="number" id="manF" min="0" step="any" value="${m.f ?? ''}" placeholder="t.ex. 32" aria-label="Omräkningsfaktor">
      <select id="manU" aria-label="Enhet"><option value="kg/m³"${m.u === 'kg/m²' ? '' : ' selected'}>kg/m³ (× Volume)</option><option value="kg/m²"${m.u === 'kg/m²' ? ' selected' : ''}>kg/m² (× Net Area)</option></select></div></div>`;
  }
  $('catDetails').innerHTML = html;
}

// EPD tab: the project's own EPD/manual resources and the EPD files in /epd
function renderEpdList(q) {
  q = q.trim().toLowerCase();
  const own = customKeys().filter(k => !q || RES[k].n.toLowerCase().includes(q));
  const files = resStatus.epds.map((e, i) => [e, i]).filter(([e]) => !q || e.name.toLowerCase().includes(q));
  $('epdList').innerHTML =
    (own.length ? `<optgroup label="I projektet">${own.map(k => `<option value="c:${k}">${esc(RES[k].n)}</option>`).join('')}</optgroup>` : '') +
    (files.length ? `<optgroup label="EPD-filer">${files.map(([e, i]) => `<option value="e:${i}">${esc(e.name)}</option>`).join('')}</optgroup>` : '') +
    (!own.length && !files.length ? '<option disabled>Inga EPD:er hittades</option>' : '');
  if (epdSel) $('epdList').value = epdSel;
}
const F = { name: 'cfName', decl: 'cfDecl', waste: 'cfWaste', factor: 'cfFactor', unit: 'cfUnit', a13: 'cfA13', a4: 'cfA4', a5: 'cfA5' };
const unitOption = (u) => ({ 'kg/m3': 'kg/m³', 'kg/m2': 'kg/m²', 'kg/m': 'kg/m', 'kg/st': 'kg/st' }[normUnit(u)] || 'kg/m³');
function clearCustomForm() {
  epdSel = null;
  Object.values(F).forEach(id => { $(id).value = ''; });
  $(F.unit).value = 'kg/m³';
  $('epdInfo').hidden = true;
}
function pickEpd(value) {
  epdSel = value;
  $('catMsg').textContent = '';
  if (!value) return;
  if (value.startsWith('c:')) {
    const x = RES[value.slice(2)];
    $(F.name).value = x.n; $(F.decl).value = x.declUnit || ''; $(F.waste).value = x.w ?? '';
    $(F.factor).value = x.f ?? ''; $(F.unit).value = unitOption(x.u);
    $(F.a13).value = x.a13 ?? ''; $(F.a4).value = x.a4 ?? ''; $(F.a5).value = x.a5 ?? '';
    $('epdInfo').hidden = false;
    $('epdInfo').innerHTML = `<b>Används i projektet</b><span>Ändrar du värdena skapas en ny resurs, de rader som redan använder den här påverkas inte.</span>`;
    return;
  }
  const e = resStatus.epds[Number(value.slice(2))];
  if (!e) return;
  const d = e.data;
  const decl = { qm: 'm2', m2: 'm2', 'm²': 'm2', m3: 'm3', 'm³': 'm3', kg: 'kg', pcs: 'st', st: 'st', piece: 'st', m: 'm' }[String(d.refUnit || '').toLowerCase()] || '';
  // with factor 1 in the declared unit the EPD's per-unit values count against that quantity
  const unit = { m2: 'kg/m²', m3: 'kg/m³', st: 'kg/st', m: 'kg/m' }[decl];
  $(F.name).value = d.name || e.name; $(F.decl).value = decl;
  $(F.waste).value = '';
  $(F.factor).value = unit ? 1 : ''; $(F.unit).value = unit || 'kg/m³';
  // EPDs declare per reference quantity (e.g. 1000 kg): the form wants per one unit
  const refQ = Number(d.refQuantity) > 0 ? Number(d.refQuantity) : 1;
  const per = (v) => (v == null || v === '' ? '' : Number((v / refQ).toPrecision(6)));
  $(F.a13).value = per(d.a1a3); $(F.a4).value = per(d.a4); $(F.a5).value = per(d.a5);
  const declText = { m2: 'm²', m3: 'm³' }[decl] || decl || 'okänd enhet';
  $('epdInfo').hidden = false;
  $('epdInfo').innerHTML = `<b>${esc(d.name || e.name)}</b>
    <span>${refQ !== 1 ? `Deklarerad per ${fNum(refQ)} ${esc(declText)}, omräknat per ${esc(declText)}` : `Värden per ${esc(declText)}`}${d.declarationOwner ? ` · ${esc(d.declarationOwner)}` : ''}${d.registrationNumber ? ` · ${esc(d.registrationNumber)}` : ''}</span>
    <span>${unit ? `Omräkningsfaktor 1 ${esc(unit)} gör att värdena räknas per ${esc(declText)}.` : 'Ange materialets densitet som omräkningsfaktor.'}</span>
    ${d.url ? `<a href="${esc(d.url)}" target="_blank" rel="noopener">Visa EPD</a>` : ''}`;
}
// Reads the form into a resource (reusing an identical one), or explains what's missing.
function customFromForm() {
  const val = (id) => $(id).value.trim();
  const numOr = (id, d) => { const v = val(id); if (v === '') return d; const n = parseFloat(v.replace(',', '.')); return Number.isFinite(n) ? n : NaN; };
  const n = val(F.name), fac = numOr(F.factor, NaN), a13 = numOr(F.a13, NaN);
  if (!n) return { error: 'Ange ett namn.', focus: F.name };
  if (!(fac > 0)) return { error: 'Ange en omräkningsfaktor.', focus: F.factor };
  if (!Number.isFinite(a13)) return { error: 'Ange klimatpåverkan A1–A3.', focus: F.a13 };
  const a4 = numOr(F.a4, 0), a5 = numOr(F.a5, 0), w = numOr(F.waste, null);
  if ([a4, a5].some(Number.isNaN) || Number.isNaN(w)) return { error: 'Kontrollera siffrorna.', focus: F.a4 };
  const src = epdSel && epdSel.startsWith('e:') ? resStatus.epds[Number(epdSel.slice(2))]?.file : '';
  return { key: addCustom({ n, declUnit: val(F.decl), f: fac, u: $(F.unit).value, w, a13, a4, a5, source: src || 'manuell' }) };
}

function applyToTarget(key) {
  const t = catTarget;
  if (t.draft != null) {
    setRes(catObj(t), key);
    const focusId = t.part === 2 ? `lm-m2res-${t.draft}` : `lm-res-${t.draft}`;
    closeCatalog(); renderLayerModal(); $(focusId) && $(focusId).focus();
    return;
  }
  checkpoint();
  if (t.rows) {
    const targets = t.rows.filter(r => !r.layers);
    targets.forEach(r => setRes(r, key));
    const skipped = t.rows.length - targets.length;
    toast(`${resName(key)} mappad på ${targets.length} ${targets.length === 1 ? 'rad' : 'rader'}${skipped ? `, ${skipped} skiktad${skipped === 1 ? '' : 'e'} hoppades över` : ''}.`);
  } else {
    setRes(catObj(t), key);
    toast(`${resName(key)} mappad på skiktet ${catLayer(t).name || ''}.`.replace(' .', '.'));
  }
  if (t.clearTicks) app.view.ticked.clear();
  closeCatalog();
  commit();
}
function useCatalog() {
  if (!catTarget) return;
  if (catTab === 'epd') {
    const r = customFromForm();
    if (r.error) { $('catMsg').textContent = r.error; $(r.focus).focus(); return; }
    applyToTarget(r.key);
    return;
  }
  const key = $('catList').value;
  if (!key) return;
  if (lacksWeightUnit(key)) {
    const fv = parseFloat($('manF') && $('manF').value);
    if (!(fv > 0)) { $('catMsg').textContent = 'Ange en omräkningsfaktor för den här resursen.'; $('manF') && $('manF').focus(); return; }
    MANUAL[key] = { f: fv, u: $('manU').value };
  }
  applyToTarget(key);
}
// Removing a mapping puts the row (or layer) back to "Ej mappad".
function unmapRows(rows) {
  const targets = rows.filter(r => !r.layers && r.res);
  if (targets.length) checkpoint();
  targets.forEach(clearRes);
  toast(targets.length ? `Mappningen borttagen på ${targets.length} ${targets.length === 1 ? 'rad' : 'rader'}.` : 'De markerade raderna var inte mappade.');
  return targets.length;
}
function clearCatalogTarget() {
  const t = catTarget;
  if (!t) return;
  if (t.draft != null) { clearRes(catObj(t)); closeCatalog(); renderLayerModal(); return; }
  if (t.rows) unmapRows(t.rows);
  else { checkpoint(); clearRes(catObj(t)); toast(`Mappningen borttagen på skiktet ${catLayer(t).name || ''}.`); }
  if (t.clearTicks) app.view.ticked.clear();
  closeCatalog();
  commit();
}

function mapRows(rows, label, clearTicks) {
  if (!rows.length) { toast('Inga rader att mappa.'); return; }
  openCatalog({ rows, label, clearTicks });
}
function layerRows(rows, clearTicks) {
  if (!rows.length) { toast('Inga rader att skikta.'); return; }
  openLayerModal(rows, rows.length === 1 ? rowLabel(rows[0]) : `${rows.length} rader`, clearTicks);
}

/* ---------- skiktningsrutan ---------- */
let lmTargetRows = [], lmRows = [], lmThickMm = null, lmClearTicks = false;
const newLayer = (t = '') => ({ name: '', t, res: null, generic: null, mix: null });
const LM_PRESETS = {
  'Innervägg utan isolering': () => [
    { name: 'Gips', t: 12.5, res: boverketById(6000000020, 'Gipsskiva, standardskiva'), mix: null },
    { name: 'Aluminium profil', t: 45, res: boverketById(6000000159, 'Aluminiumprofiler, primär'), mix: { name: 'Luft', res: boverketById(null, 'Luft'), pct: 99 } },
    { name: 'Gips', t: 12.5, res: boverketById(6000000020, 'Gipsskiva, standardskiva'), mix: null }],
  'Innervägg med isolering (stenull)': () => [
    { name: 'Gips', t: 12.5, res: boverketById(6000000020, 'Gipsskiva, standardskiva'), mix: null },
    { name: 'Aluminium profil', t: 45, res: boverketById(6000000159, 'Aluminiumprofiler, primär'), mix: { name: 'Stenull', res: boverketById(6000000123, 'Stenull, skivor och rullar'), pct: 99 } },
    { name: 'Gips', t: 12.5, res: boverketById(6000000020, 'Gipsskiva, standardskiva'), mix: null }]
};
function openLayerModal(rows, label, clearTicks) {
  lmTargetRows = rows;
  lmClearTicks = !!clearTicks;
  const layered = rows.find(r => r.layers);
  lmRows = layered ? clone(layered.layers).map(l => ({ ...l, mix: l.mix || null })) : [newLayer(), newLayer()];
  // the element's own thickness (Thickness in metres in the IFC export), single row only
  const tKey = rows.length === 1 ? Object.keys(rows[0].cells).find(k => k.toLowerCase() === 'thickness') : null;
  const tM = tKey ? parseNum(rows[0].cells[tKey]) : NaN;
  lmThickMm = tM > 0 ? Math.round(tM * 10000) / 10 : null;
  $('lmSub').textContent = label + (lmThickMm ? ` · Byggdelens tjocklek ${fNum(lmThickMm)} mm` : '');
  $('lmRemove').hidden = !rows.some(r => r.layers);
  $('lmPaste').hidden = !app.view.copied;
  $('lmMsg').textContent = '';
  renderLayerModal();
  $('layerModal').hidden = false;
  $('lm-name-0') && $('lm-name-0').focus();
}
function closeLayerModal() { $('layerModal').hidden = true; lmTargetRows = []; }
const lmTotal = () => lmRows.reduce((s, l) => s + (Number(l.t) > 0 ? Number(l.t) : 0), 0);
const lmMissing = (l) => !l.res || (l.mix && !l.mix.res);
function lmResBtn(id, part, i, value, aria) {
  return `<button type="button" class="lm-res${value ? '' : ' res-missing'}" id="${id}" data-act="lm-cat" data-i="${i}" data-part="${part}" title="${esc(value ? resName(value) : 'Välj klimatresurs')}" aria-label="${esc(aria)}">${isCustom(value) ? 'EPD: ' : ''}${esc(value ? resName(value) : 'Välj klimatresurs…')}</button>`;
}
function lmBarHtml() {
  const total = lmTotal();
  return lmRows.map(l => {
    const t = Number(l.t) > 0 ? Number(l.t) : 0;
    const name = l.mix ? `${l.name || 'Material 1'} + ${l.mix.name || 'Material 2'} (${100 - l.mix.pct} % / ${l.mix.pct} %)` : (l.name || 'Namnlöst skikt');
    return `<div class="lm-seg${l.mix ? ' lm-seg-mixed' : ''}" style="flex:${total ? t / total : 1}" title="${esc(name)}${t ? `, ${fNum(t)} mm` : ''}"><span>${t ? fNum(t) : '–'}</span></div>`;
  }).join('');
}
function renderLayerModal() {
  $('lmBar').innerHTML = lmBarHtml();
  $('lmRows').innerHTML = lmRows.map((l, i) => `
    <div class="lm-layer${l.mix ? ' lm-layer-mixed' : ''}" data-i="${i}">
      <div class="lm-grid">
        <span class="lm-n">${i + 1}</span>
        <input type="text" id="lm-name-${i}" data-f="name" value="${esc(l.name)}" placeholder="t.ex. Gips" aria-label="Namn, skikt ${i + 1}">
        <input type="number" id="lm-t-${i}" data-f="t" min="0" step="0.5" value="${esc(l.t)}" aria-label="Tjocklek i mm, skikt ${i + 1}">
        ${lmResBtn(`lm-res-${i}`, 1, i, l.res, `Klimatresurs, skikt ${i + 1}`)}
        <label class="lm-mixbox"><input type="checkbox" id="lm-mix-${i}" data-f="mix"${l.mix ? ' checked' : ''}> Blandat</label>
        <button type="button" class="lm-del" data-act="lm-del" data-i="${i}" title="Ta bort skikt" aria-label="Ta bort skikt ${i + 1}"${lmRows.length === 1 ? ' disabled' : ''}>✕</button>
      </div>
      ${l.mix ? `<div class="lm-grid">
        <span class="lm-n" title="Material 2 i samma skikt">+</span>
        <input type="text" id="lm-m2name-${i}" data-f="m2name" value="${esc(l.mix.name)}" placeholder="Material 2, t.ex. Stenull" aria-label="Material 2, skikt ${i + 1}">
        <div class="lm-unit" data-unit="%"><input type="number" id="lm-pct-${i}" data-f="pct" min="1" max="99" value="${esc(l.mix.pct)}" aria-label="Andel material 2 i procent"></div>
        ${lmResBtn(`lm-m2res-${i}`, 2, i, l.mix.res, `Klimatresurs, material 2 i skikt ${i + 1}`)}
        <span></span><span></span></div>` : ''}
    </div>`).join('');
  renderLayerTotals();
}
function renderLayerTotals() {
  const total = lmTotal(), filled = lmRows.filter(l => Number(l.t) > 0).length, missing = lmRows.filter(lmMissing).length;
  const out = [];
  if (!filled) {
    $('lmTotal').textContent = 'Ingen tjocklek angiven';
    out.push('<span class="res-missing">Skikten delas jämnt</span>');
  } else {
    $('lmTotal').innerHTML = `Totalt <b class="num">${fNum(total)} mm</b>${lmThickMm ? ` av ${fNum(lmThickMm)} mm` : ''}`;
    if (filled < lmRows.length) out.push('<span class="res-missing">Tjocklek saknas på något skikt</span>');
    else if (lmThickMm) {
      const diff = total - lmThickMm;
      out.push(Math.abs(diff) < 0.05 ? '<span class="ok-text">Tjockleken stämmer</span>' : `<span class="res-missing">${diff > 0 ? fNum(diff) + ' mm för mycket' : fNum(-diff) + ' mm kvar'}</span>`);
    }
  }
  out.push(missing ? `<span class="res-missing">${missing} skikt saknar klimatresurs</span>` : '<span class="ok-text">Alla skikt har klimatresurs</span>');
  $('lmStatus').innerHTML = out.join(' · ');
}
function applyLayers() {
  const filled = lmRows.filter(l => Number(l.t) > 0).length;
  if (filled && filled < lmRows.length) { $('lmMsg').textContent = 'Ange tjocklek för alla skikt, eller lämna alla tomma för att dela jämnt.'; return; }
  const layers = lmRows.map((l, i) => ({
    name: l.name.trim() || `Skikt ${i + 1}`, t: Number(l.t) || 0, res: l.res, generic: l.generic || null,
    mix: l.mix ? { ...l.mix, name: (l.mix.name || '').trim() || 'Material 2' } : null
  }));
  checkpoint();
  lmTargetRows.forEach(r => { r.layers = clone(layers); r.res = null; r.generic = null; });
  const n = lmTargetRows.length;
  if (lmClearTicks) app.view.ticked.clear();
  closeLayerModal();
  commit();
  toast(`${n === 1 ? 'Raden' : `${n} rader`} delad${n === 1 ? '' : 'e'} i ${layers.length} skikt. Klicka på "${layers.length} skikt" för att se dem.`);
}

// "Kopiera skikt": remembered until replaced or cleared, pasted on ticked rows or in the layer modal
function copyLayers(r) {
  app.view.copied = { layers: clone(r.layers), label: rowLabel(r) };
  renderCopied(); renderTable();
  toast(`Skiktningen kopierad (${r.layers.length} skikt). Markera rader och klicka på Klistra in skiktning.`);
}

/* ---------- details drawer ---------- */
function fillDrawer(rid) {
  const r = rowById(rid);
  if (!r) { closeDrawer(); return; }
  $('dTitle').textContent = rowLabel(r);
  $('dSub').innerHTML = `${[r.cells[app.headers.find(h => h.toLowerCase().includes('byggdel')) || ''], r.cells[app.headers.find(h => h.toLowerCase() === 'component') || '']].filter(Boolean).map(esc).join(' · ')} ${pill(r.status)}`;
  const kv = (pairs) => `<dl class="kv">${pairs.map(([k, v]) => `<dt>${k}</dt><dd class="num">${v}</dd>`).join('')}</dl>`;
  const fileRows = app.headers.filter(Boolean).map(h => [esc(h), app.groupColumns.includes(h) ? esc(r.cells[h] || '–') : (fNum(r.cells[h]) || '–')]);
  let html = `<section><h3>Från mängdavtagningen</h3>${kv(fileRows)}</section>`;
  const resInfo = (key) => {
    const x = RES[key], m = factorOf(key);
    const rows = [['Resurs', `${epdTag(key)}<button class="cellbtn" type="button" data-act="map" data-id="${r.id}" title="Byt klimatresurs">${esc(x.n)}</button>`],
      ['Omräkningsfaktor', m.f != null ? `${fNum(m.f)} ${esc(m.u || '(enhet saknas)')}` : 'Ej angiven'],
      ['Spillfaktor', wasteMultiplier(x.w) != null ? fNum(wasteMultiplier(x.w)) : '–'],
      ['Emissionsfaktor A1–A3', x.a13 != null ? fNum(x.a13) : '–']];
    if (isCustom(key) && x.declUnit) rows.splice(1, 0, ['Deklarerad enhet', esc(x.declUnit)]);
    return kv(rows);
  };
  if (r.layers) {
    const layerRes = (li, part, key) => `${epdTag(key)}<button class="cellbtn${key ? '' : ' res-missing'}" type="button" data-act="map-layer" data-id="${r.id}" data-layer="${li}" data-part="${part}">${key ? esc(resName(key)) : 'Saknas – välj'}</button>`;
    html += `<section><h3>Skikt</h3><div class="layers">${r.layerResults.map((l, li) => `<div><span>${esc(l.name)}${l.mix ? ` + ${esc(l.mix.name)}` : ''}</span><span class="num">${l.t ? `${fNum(l.t)} mm` : 'jämnt'} · ${f(layerQty(r, l).volume, 2)} m³</span><small>${layerRes(li, 1, l.res)}${l.mix ? ` + ${layerRes(li, 2, l.mix.res)} (${l.mix.pct} %)` : ''}</small></div>`).join('')}</div></section>`;
  } else if (r.res) {
    html += `<section><h3>Klimatresurs</h3>${resInfo(r.res)}</section>`;
  } else {
    html += `<section><h3>Klimatresurs</h3><button class="cellbtn res-missing" type="button" data-act="map" data-id="${r.id}">Ingen klimatresurs ännu – välj</button></section>`;
  }
  if (r.impact) {
    const all = totals().total;
    const pairs = [['Inbyggd vikt', f(r.impact.w, 0) + ' kg'], ['Inköpt vikt', f(r.impact.bought, 0) + ' kg'], ['A1–A3', f(r.impact.a13, 0) + ' kg CO₂e'], ['A4', f(r.impact.a4, 0) + ' kg CO₂e'], ['A5', f(r.impact.a5, 0) + ' kg CO₂e'], ['Andel av projektet', f(all ? tot(r.impact) / all * 100 : 0) + ' %']];
    const red = totals([r]);
    if (red.hasReduction) pairs.push(['Jämfört med generisk', `${red.reduction >= 0 ? '−' : '+'}${f(Math.abs(red.reduction), 0)} kg CO₂e`]);
    html += `<section><h3>Resultat</h3>${kv(pairs)}</section>`;
  }
  $('dBody').innerHTML = html;
  $('dMap').hidden = !!r.layers;
  $('dMap').textContent = r.res ? 'Byt klimatresurs' : 'Mappa klimatresurs';
  $('dLayer').textContent = r.layers ? 'Redigera skikt' : 'Skikta';
  $('dUnmap').hidden = !r.res;
  $('dCopy').hidden = !r.layers;
}
function openDrawer(rid) { app.view.drawer = Number(rid); fillDrawer(app.view.drawer); $('scrim').hidden = false; $('drawer').hidden = false; $('dClose').focus(); }
function closeDrawer() { $('scrim').hidden = true; $('drawer').hidden = true; app.view.drawer = null; }

// mapping clicks shared by the table and the details drawer
function handleMapClick(el) {
  const act = el.dataset.act;
  if (act === 'map') { const r = rowById(el.dataset.id); mapRows([r], rowLabel(r)); return true; }
  if (act === 'map-layer') {
    const r = rowById(el.dataset.id), li = Number(el.dataset.layer), part = Number(el.dataset.part);
    const l = r.layers[li];
    openCatalog({ row: r, layer: li, part, label: `${rowLabel(r)} · skikt ${l.name}${part === 2 ? `, ${l.mix.name || 'material 2'}` : ''}` });
    return true;
  }
  return false;
}

/* ---------- Lägg till rad ---------- */
function openRowModal() {
  const lists = app.headers.filter(Boolean).map((h, i) => {
    const values = [...new Set(app.rows.map(r => r.cells[h]).filter(Boolean))].slice(0, 300);
    return `<label>${esc(h)}<input type="text" data-h="${esc(h)}" list="rowList${i}"${app.groupColumns.includes(h) ? '' : ' inputmode="decimal"'}><datalist id="rowList${i}">${values.map(v => `<option value="${esc(v)}">`).join('')}</datalist></label>`;
  });
  $('rowFields').innerHTML = lists.join('');
  $('rowModal').hidden = false;
  $('rowFields').querySelector('input')?.focus();
}
function saveRowModal() {
  const values = {};
  $('rowFields').querySelectorAll('input[data-h]').forEach(i => { values[i.dataset.h] = i.value; });
  if (!Object.values(values).some(v => v.trim())) { toast('Fyll i minst ett fält.'); return; }
  checkpoint();
  const r = addRow(values);
  $('rowModal').hidden = true;
  // a new text value in a column makes that column groupable
  app.headers.forEach(h => { if (values[h] && !app.groupColumns.includes(h) && !/^-?[\d\s]+([.,]\d+)?$/.test(values[h].trim())) app.groupColumns.push(h); });
  app.view.search = ''; $('search').value = '';
  if (app.view.groupBy) app.view.open.add(groupKey(r));
  commit();
  toast(`Raden ${rowLabel(r)} är tillagd.`);
}

/* ---------- export ---------- */
function exportExcel(scope) {
  $('fileMenu').hidden = true;
  if (!window.XLSX) { toast('Excel-biblioteket kunde inte laddas. Kontrollera nätverket och försök igen.'); return; }
  derive();
  const source = scope === 'visible' ? visibleRows() : app.rows;
  const out = buildExport(source, scope === 'visible' ? 'Det som visas (filter och sökning)' : 'Alla rader');
  if (!out) { toast('Inga rader att exportera med nuvarande filter och sökning.'); return; }
  download(out.blob, out.filename);
  toast(`Exporterad: ${out.filename} (${out.rows} rader${scope === 'visible' ? ', det som visas' : ''}).`);
}

function doUndo() { if (undo()) { commit(); toast('Ångrat.', 1500); } }
function doRedo() { if (redo()) { commit(); toast('Gjort om.', 1500); } }

/* ---------- events ---------- */
export function initUi(h) {
  handlers = h;
  changed = h.onChange || (() => {});

  // header + menu
  $('undoBtn').addEventListener('click', doUndo);
  $('redoBtn').addEventListener('click', doRedo);
  $('fileMenuBtn').addEventListener('click', (e) => { e.stopPropagation(); const m = $('fileMenu'); m.hidden = !m.hidden; $('fileMenuBtn').setAttribute('aria-expanded', String(!m.hidden)); });
  document.addEventListener('click', (e) => { if (!e.target.closest('.menu')) { $('fileMenu').hidden = true; $('fileMenuBtn').setAttribute('aria-expanded', 'false'); } });
  $('fileMenu').addEventListener('click', (e) => { if (e.target.closest('button')) $('fileMenu').hidden = true; });
  document.querySelectorAll('[data-open-excel]').forEach(b => b.addEventListener('click', () => handlers.onOpenExcel()));
  document.querySelectorAll('[data-open-project]').forEach(b => b.addEventListener('click', () => handlers.onOpenProject()));
  $('saveBtn').addEventListener('click', () => handlers.onSave());
  $('saveAsBtn').addEventListener('click', () => handlers.onSaveAs());
  document.querySelectorAll('[data-export]').forEach(b => b.addEventListener('click', () => exportExcel(b.dataset.export)));
  $('addRowBtn').addEventListener('click', openRowModal);
  $('rowCancel').addEventListener('click', () => { $('rowModal').hidden = true; });
  $('rowSave').addEventListener('click', saveRowModal);
  $('rowFields').addEventListener('keydown', (e) => { if (e.key === 'Enter') { e.preventDefault(); saveRowModal(); } });

  // toolbar + overview
  $('search').addEventListener('input', (e) => { app.view.search = e.target.value; renderTable(); });
  $('chips').addEventListener('click', (e) => { const b = e.target.closest('[data-filter]'); if (!b) return; app.view.filter = b.dataset.filter; renderChips(); renderTable(); changed(false); });
  $('colsets').addEventListener('click', (e) => { const b = e.target.closest('[data-cols]'); if (!b) return; app.view.cols = b.dataset.cols; renderChips(); renderTable(); changed(false); });
  $('groupBy').addEventListener('change', (e) => {
    app.view.groupBy = e.target.value;
    app.view.open = new Set(); // groups and layers start collapsed
    renderOverview(); renderTable(); changed(false);
  });
  $('btaInput').addEventListener('input', (e) => { app.bta = e.target.value; app.dirty = true; renderOverview(); renderChrome(); changed(false); });
  $('showUnmapped').addEventListener('click', () => { app.view.filter = 'unmapped'; app.rows.forEach(r => app.view.open.add(groupKey(r))); renderChips(); renderTable(); });
  $('kTop').addEventListener('click', (e) => {
    const b = e.target.closest('[data-group]');
    if (!b) return;
    const g = b.dataset.group;
    if (!app.view.groupBy) { app.view.groupBy = app.groupColumns[0] || ''; $('groupBy').value = app.view.groupBy; }
    // show the whole group: clear what could hide its rows, open it and scroll to it
    app.view.filter = 'all'; app.view.search = ''; $('search').value = '';
    app.view.open.add(g);
    renderChips(); renderTable();
    const row = [...document.querySelectorAll('tr.grp')].find(t => t.dataset.group === g);
    if (row) row.scrollIntoView({ block: 'center', behavior: 'smooth' });
  });
  window.addEventListener('resize', sizeTable);

  // table
  $('table').addEventListener('click', (e) => {
    const el = e.target.closest('[data-act]');
    if (!el) return;
    const act = el.dataset.act;
    if (act === 'tick' || act === 'tick-all' || act === 'tick-group') return; // handled on change
    if (handleMapClick(el)) return;
    if (act === 'map-group') {
      const rows = rowsInGroup(el.dataset.group).filter(r => !r.layers && !r.res);
      mapRows(rows, `Grupp ${el.dataset.group}: ${rows.length} rader utan klimatresurs`);
    }
    if (act === 'map-group-res') {
      // the group's resource cell: map what's missing, or swap the resource for the whole group
      const g = el.dataset.group, plain = rowsInGroup(g).filter(r => !r.layers);
      const missing = plain.filter(r => !r.res);
      if (!plain.length) { app.view.open.add(g); renderTable(); toast('Gruppen består av skiktade rader. Mappa skikten under varje rad.'); return; }
      if (missing.length) mapRows(missing, `Grupp ${g}: ${missing.length} rader utan klimatresurs`);
      else mapRows(plain, `Grupp ${g}: byt klimatresurs på alla ${plain.length} rader`);
    }
    if (act === 'layer') layerRows([rowById(el.dataset.id)]);
    if (act === 'copy-layers') copyLayers(rowById(el.dataset.id));
    if (act === 'layer-group') layerRows(rowsInGroup(el.dataset.group));
    if (act === 'details') openDrawer(el.dataset.id);
    if (act === 'toggle-layers') { const k = 'L' + el.dataset.id; app.view.open.has(k) ? app.view.open.delete(k) : app.view.open.add(k); renderTable(); }
    if (act === 'toggle-group') { const k = el.dataset.group; app.view.open.has(k) ? app.view.open.delete(k) : app.view.open.add(k); renderTable(); }
  });
  $('table').addEventListener('change', (e) => {
    const el = e.target, act = el.dataset.act, t = app.view.ticked;
    if (act === 'tick') { const rid = Number(el.dataset.id); el.checked ? t.add(rid) : t.delete(rid); }
    if (act === 'tick-group') rowsInGroup(el.dataset.group).forEach(r => (el.checked ? t.add(r.id) : t.delete(r.id)));
    if (act === 'tick-all') visibleRows().forEach(r => (el.checked ? t.add(r.id) : t.delete(r.id)));
    renderTable();
  });
  // S over a row opens the skiktningsruta for it, as before
  let hovered = null;
  $('table').addEventListener('mouseover', (e) => { const tr = e.target.closest('tr.item'); hovered = tr ? Number(tr.dataset.id) : null; });
  $('table').addEventListener('mouseleave', () => { hovered = null; });
  document.addEventListener('keydown', (e) => {
    if ((e.key === 's' || e.key === 'S') && !e.ctrlKey && !e.metaKey && !e.altKey && hovered != null && !modalOpen() && !e.target.closest('input, select, textarea')) {
      const r = rowById(hovered);
      if (r) { e.preventDefault(); layerRows([r]); }
    }
  });

  // bulk bar
  const tickedRows = () => app.rows.filter(r => app.view.ticked.has(r.id));
  $('bulkMap').addEventListener('click', () => { const rows = tickedRows(); mapRows(rows, `${rows.length} markerade rader`, true); });
  $('bulkLayer').addEventListener('click', () => layerRows(tickedRows(), true));
  $('bulkClear').addEventListener('click', () => { app.view.ticked.clear(); renderTable(); });
  $('bulkPaste').addEventListener('click', () => {
    const c = app.view.copied;
    if (!c) return;
    const rows = tickedRows();
    checkpoint();
    rows.forEach(r => { r.layers = clone(c.layers); r.res = null; r.generic = null; });
    app.view.ticked.clear();
    commit();
    toast(`Skiktningen inklistrad på ${rows.length} ${rows.length === 1 ? 'rad' : 'rader'} (${c.layers.length} skikt).`);
  });
  $('copiedClear').addEventListener('click', () => { app.view.copied = null; renderCopied(); renderTable(); });
  $('bulkUnmap').addEventListener('click', () => { if (unmapRows(tickedRows())) app.view.ticked.clear(); commit(); });

  // catalogue
  $('catTabs').addEventListener('click', (e) => {
    const b = e.target.closest('[data-tab]'); if (!b) return;
    setCatTab(b.dataset.tab);
    if (b.dataset.tab === 'epd') { renderEpdList($('epdSearch').value); $('epdSearch').focus(); } else $('catSearch').focus();
  });
  $('catClear').addEventListener('click', clearCatalogTarget);
  $('catSearch').addEventListener('input', (e) => renderCatList(e.target.value));
  $('catList').addEventListener('change', () => { $('catMsg').textContent = ''; showCatDetails($('catList').value); });
  $('catList').addEventListener('dblclick', useCatalog);
  $('epdSearch').addEventListener('input', (e) => renderEpdList(e.target.value));
  $('epdList').addEventListener('change', () => pickEpd($('epdList').value));
  $('customForm').addEventListener('submit', (e) => { e.preventDefault(); useCatalog(); });
  $('customForm').addEventListener('input', () => { $('catMsg').textContent = ''; });
  $('catUse').addEventListener('click', useCatalog);
  $('catCancel').addEventListener('click', closeCatalog);
  $('catalog').addEventListener('click', (e) => { if (e.target === $('catalog')) closeCatalog(); });
  loadEpds().then(() => { if (!$('catalog').hidden) renderEpdList($('epdSearch').value); renderStatus(); });

  // layer modal
  $('lmRows').addEventListener('input', (e) => {
    const layer = e.target.closest('.lm-layer'); if (!layer) return;
    const l = lmRows[Number(layer.dataset.i)], v = e.target.value;
    const field = e.target.dataset.f;
    if (field === 'name') l.name = v;
    else if (field === 't') l.t = v === '' ? '' : Math.max(0, parseFloat(v) || 0);
    else if (field === 'm2name') l.mix.name = v;
    else if (field === 'pct') l.mix.pct = Math.min(99, Math.max(1, Math.round(parseFloat(v) || 1)));
    else return;
    // update the bar and totals without rebuilding the inputs (keeps focus)
    $('lmBar').innerHTML = lmBarHtml();
    renderLayerTotals();
  });
  $('lmRows').addEventListener('change', (e) => {
    if (e.target.dataset.f !== 'mix') return;
    const i = Number(e.target.closest('.lm-layer').dataset.i);
    lmRows[i].mix = e.target.checked ? { name: '', res: null, generic: null, pct: 50 } : null;
    renderLayerModal();
    $(e.target.id) && $(e.target.id).focus();
  });
  $('lmRows').addEventListener('click', (e) => {
    const b = e.target.closest('[data-act]'); if (!b) return;
    const i = Number(b.dataset.i);
    if (b.dataset.act === 'lm-del' && lmRows.length > 1) { lmRows.splice(i, 1); renderLayerModal(); }
    if (b.dataset.act === 'lm-cat') {
      const part = Number(b.dataset.part), l = lmRows[i];
      openCatalog({ draft: i, part, label: part === 2 ? `Skikt ${i + 1}, material 2: ${l.mix.name || 'namnlöst'}` : `Skikt ${i + 1}: ${l.name || 'namnlöst'}` });
    }
  });
  $('lmAdd').addEventListener('click', () => {
    const rest = lmThickMm ? Math.round((lmThickMm - lmTotal()) * 10) / 10 : 0;
    lmRows.push(newLayer(rest > 0 ? rest : ''));
    renderLayerModal();
    $(`lm-name-${lmRows.length - 1}`).focus();
  });
  Object.entries(LM_PRESETS).forEach(([name, make]) => {
    const b = document.createElement('button');
    b.type = 'button'; b.className = 'chip'; b.textContent = name;
    b.addEventListener('click', () => { lmRows = make().map(l => ({ ...l, generic: null, mix: l.mix ? { ...l.mix, generic: null } : null })); $('lmMsg').textContent = `Snabbvalet "${name}" är ifyllt.`; renderLayerModal(); });
    $('lmPresets').appendChild(b);
  });
  $('lmCancel').addEventListener('click', closeLayerModal);
  $('lmPaste').addEventListener('click', () => {
    const c = app.view.copied; if (!c) return;
    lmRows = clone(c.layers).map(l => ({ ...l, mix: l.mix || null }));
    $('lmMsg').textContent = `Skikten från ${c.label} är inklistrade.`;
    renderLayerModal();
  });
  $('lmApply').addEventListener('click', applyLayers);
  $('lmRemove').addEventListener('click', () => {
    const n = lmTargetRows.filter(r => r.layers).length;
    checkpoint();
    lmTargetRows.forEach(r => { r.layers = null; });
    closeLayerModal();
    commit();
    toast(`Skiktningen borttagen på ${n} ${n === 1 ? 'rad' : 'rader'}.`);
  });

  // drawer
  $('dClose').addEventListener('click', closeDrawer);
  $('scrim').addEventListener('click', closeDrawer);
  $('dMap').addEventListener('click', () => { const r = rowById(app.view.drawer); mapRows([r], rowLabel(r)); });
  $('dLayer').addEventListener('click', () => layerRows([rowById(app.view.drawer)]));
  $('dUnmap').addEventListener('click', () => { unmapRows([rowById(app.view.drawer)]); commit(); });
  $('dCopy').addEventListener('click', () => copyLayers(rowById(app.view.drawer)));
  $('dBody').addEventListener('click', (e) => { const el = e.target.closest('[data-act]'); if (el) handleMapClick(el); });

  // keyboard
  document.addEventListener('keydown', (e) => {
    const mod = e.ctrlKey || e.metaKey;
    if (e.key === 'Escape') {
      if (!$('catalog').hidden) { closeCatalog(); return; }
      if (!$('rowModal').hidden) { $('rowModal').hidden = true; return; }
      if (!$('layerModal').hidden) { closeLayerModal(); return; }
      if (!$('drawer').hidden) closeDrawer();
      $('fileMenu').hidden = true;
      return;
    }
    if (mod && e.key.toLowerCase() === 's') { e.preventDefault(); if (app.rows.length) (e.shiftKey ? handlers.onSaveAs : handlers.onSave)(); return; }
    const inField = e.target.closest && e.target.closest('input, select, textarea');
    if (mod && !inField && !modalOpen()) {
      if (e.key.toLowerCase() === 'z' && !e.shiftKey) { e.preventDefault(); doUndo(); }
      else if (e.key.toLowerCase() === 'y' || (e.key.toLowerCase() === 'z' && e.shiftKey)) { e.preventDefault(); doRedo(); }
    }
  });
}
const modalOpen = () => !$('catalog').hidden || !$('layerModal').hidden || !$('rowModal').hidden;

// After a new file or project: reset what belongs to the previous one.
export function resetView() {
  closeDrawer(); closeCatalog(); closeLayerModal();
  $('rowModal').hidden = true;
  $('search').value = app.view.search;
}
