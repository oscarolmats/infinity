// Klimatanalys: start-up, files and the IFC Viewer tab.
import { loadBoverket, clearCustom, clearManual } from './resources.js';
import { app, makeRows, setRows, clearHistory } from './model.js';
import { buildProject, loadProject } from './project.js';
import { readSpreadsheet, download } from './io.js';
import { initUi, renderAll, renderStatus, renderChrome, toast, resetView, $ } from './ui.js';

/* ---------- tabs ---------- */
function showTab(ifc) {
  document.body.classList.toggle('ifc-active', ifc);
  $('tabIfc').classList.toggle('active', ifc);
  $('tabTable').classList.toggle('active', !ifc);
}
$('tabIfc').addEventListener('click', () => showTab(true));
$('tabTable').addEventListener('click', () => showTab(false));

/* ---------- guards ---------- */
// Same password as before for opening files outside local development.
let authenticated = false;
function allowedToOpen() {
  const local = location.hostname === 'localhost' && location.port === '3000';
  if (local || authenticated) return true;
  const pw = prompt('Ange lösenord för att fortsätta:');
  if (pw === null) return false;
  if (pw === 'fritidsgård') { authenticated = true; return true; }
  alert('Felaktigt lösenord');
  return false;
}
const discardOk = () => !app.dirty || confirm('Du har osparade ändringar. Vill du fortsätta ändå? De osparade ändringarna försvinner.');
window.addEventListener('beforeunload', (e) => { if (app.dirty) { e.preventDefault(); e.returnValue = ''; } });

/* ---------- opening ---------- */
// A new take-off starts a new project: no mappings, no EPDs, no history.
function startFrom(table, name, source) {
  const [headers = [], ...data] = table;
  const built = makeRows(headers, data);
  if (!built.rows.length) { toast(`Hittade inga rader i ${name}.`); return false; }
  setRows(built);
  clearCustom(); clearManual(); clearHistory();
  const typeCol = built.groupColumns.find(h => h.toLowerCase() === 'type');
  Object.assign(app.view, { filter: 'all', search: '', ticked: new Set(), open: new Set(), drawer: null,
    groupBy: typeCol || built.groupColumns[0] || '' });
  app.title = name.replace(/\.[^.]+$/, '');
  app.sourceFile = source;
  app.fileHandle = null;
  app.dirty = false;
  resetView();
  renderAll();
  const noVolume = built.rows.filter(r => !r.volume).length;
  toast(`Läste in ${built.rows.length} rader från ${name}.${noVolume ? ` ${noVolume} saknar volym.` : ''}`, 4500);
  return true;
}
async function openSpreadsheet(file) {
  if (!discardOk()) return;
  toast(`Läser ${file.name}…`, 20000);
  try {
    startFrom(await readSpreadsheet(file), file.name, file.name);
  } catch (err) {
    console.error(err);
    toast(`Kunde inte läsa ${file.name}. Är det en Excel- eller CSV-fil?`, 5000);
  }
}
async function openProjectFile(file, handle = null) {
  if (!discardOk()) return;
  let p;
  try { p = JSON.parse(await file.text()); } catch { toast(`${file.name} är ingen projektfil (JSON).`, 5000); return; }
  try {
    await loadBoverket(); // saved resources are matched against Boverket by name
    const msg = loadProject(p, file.name);
    clearHistory();
    app.fileHandle = handle;
    app.dirty = false;
    resetView();
    renderAll();
    toast(msg, 7000);
  } catch (err) {
    console.error(err);
    toast(`Kunde inte öppna ${file.name} som projekt.`, 5000);
  }
}

$('excelInput').addEventListener('change', (e) => { const f = e.target.files[0]; e.target.value = ''; if (f) openSpreadsheet(f); });
$('projectInput').addEventListener('change', (e) => { const f = e.target.files[0]; e.target.value = ''; if (f) openProjectFile(f); });
const pickerTypes = [{ description: 'INFINITY-projekt', accept: { 'application/json': ['.json'] } }];
async function chooseProject() {
  // with the File System Access API, "Spara" can later write back to the same file
  if ('showOpenFilePicker' in window) {
    try {
      const [handle] = await window.showOpenFilePicker({ types: pickerTypes });
      openProjectFile(await handle.getFile(), handle);
    } catch (err) { if (err.name !== 'AbortError') $('projectInput').click(); }
    return;
  }
  $('projectInput').click();
}

/* ---------- saving ---------- */
const projectName = () => `${(app.title || 'projekt').replace(/[\\/:*?"<>|]/g, '').trim() || 'projekt'}_projekt.json`;
function saved(name) {
  app.dirty = false;
  renderChrome();
  toast(`Sparat: ${name}`, 2500);
}
async function saveAs() {
  if (!app.rows.length) return;
  const json = JSON.stringify(buildProject(), null, 2);
  if ('showSaveFilePicker' in window) {
    try {
      const handle = await window.showSaveFilePicker({ suggestedName: app.fileHandle ? app.fileHandle.name : projectName(), types: pickerTypes });
      const w = await handle.createWritable();
      await w.write(json); await w.close();
      app.fileHandle = handle;
      saved(handle.name);
    } catch (err) { if (err.name !== 'AbortError') toast('Kunde inte spara filen: ' + err.message, 5000); }
    return;
  }
  download(new Blob([json], { type: 'application/json' }), projectName());
  saved(projectName());
}
async function save() {
  if (!app.rows.length) return;
  if (!app.fileHandle) { saveAs(); return; }
  try {
    const w = await app.fileHandle.createWritable();
    await w.write(JSON.stringify(buildProject(), null, 2)); await w.close();
    saved(app.fileHandle.name);
  } catch (err) {
    // permission withdrawn or the file moved: ask where to save instead
    if (err.name !== 'AbortError') saveAs();
  }
}

initUi({
  onOpenExcel: () => { if (allowedToOpen()) $('excelInput').click(); },
  onOpenProject: chooseProject,
  onSave: save,
  onSaveAs: saveAs,
  onChange: () => renderChrome()
});

/* ---------- drag and drop: Excel/CSV or a project ---------- */
let dragDepth = 0;
const hasFiles = (e) => [...(e.dataTransfer?.types || [])].includes('Files');
document.addEventListener('dragenter', (e) => { if (!hasFiles(e) || document.body.classList.contains('ifc-active')) return; e.preventDefault(); dragDepth++; $('drop').hidden = false; });
document.addEventListener('dragover', (e) => { if (hasFiles(e) && !document.body.classList.contains('ifc-active')) e.preventDefault(); });
document.addEventListener('dragleave', () => { if (--dragDepth <= 0) { dragDepth = 0; $('drop').hidden = true; } });
document.addEventListener('drop', (e) => {
  if (!hasFiles(e) || document.body.classList.contains('ifc-active')) return;
  e.preventDefault();
  dragDepth = 0; $('drop').hidden = true;
  const file = e.dataTransfer.files[0];
  if (!file) return;
  const ext = (file.name.split('.').pop() || '').toLowerCase();
  if (ext === 'json') openProjectFile(file);
  else if (['xlsx', 'csv', 'tsv', 'txt'].includes(ext)) { if (allowedToOpen()) openSpreadsheet(file); }
  else toast('Släpp en Excel-fil (.xlsx), en CSV-fil eller ett projekt (.json).', 4000);
});

/* ---------- quantities sent from the IFC Viewer ---------- */
window.addEventListener('message', (event) => {
  if (!event.data || event.data.type !== 'ifcinfinity-export') return;
  const rows = event.data.rows;
  if (!Array.isArray(rows) || rows.length < 2) return;
  showTab(false);
  if (!discardOk()) return;
  startFrom(rows.map(r => (Array.isArray(r) ? r.map(c => String(c ?? '')) : [])), 'IFC-export', 'IFC Viewer');
});

/* ---------- start ---------- */
renderAll();
loadBoverket().then(() => { renderStatus(); if (app.rows.length) renderAll(); });
