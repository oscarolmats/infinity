import * as THREE from "three";
import * as OBC from "@thatopen/components";
import * as OBF from "@thatopen/components-front";
import type * as FRAGS from "@thatopen/fragments";
import writeExcelFile from "write-excel-file/browser";
import { generateBuildingIfc } from "./ifcGenerator";
import { FootprintEditor } from "./footprintEditor";
import { type SyncedFile, readFileMeta, saveSourceFiles, loadSourceFiles } from "./ifcFileSync";

const container = document.getElementById("container") as HTMLDivElement;
const statusSelection = document.getElementById("status-selection") as HTMLSpanElement;
const statusMode = document.getElementById("status-mode") as HTMLSpanElement;
const statusProgress = document.getElementById("status-progress") as HTMLDivElement;
const statusProgressLabel = document.getElementById("status-progress-label") as HTMLSpanElement;
const statusProgressFill = document.getElementById("status-progress-fill") as HTMLDivElement;
const filesAddButton = document.getElementById("files-add") as HTMLButtonElement;
const filesRescanButton = document.getElementById("files-rescan") as HTMLButtonElement;
const filesStatus = document.getElementById("files-status") as HTMLDivElement;
const filesAutoUpdateCheckbox = document.getElementById("files-auto-update") as HTMLInputElement;
const propertiesPanel = document.getElementById("properties-panel") as HTMLDivElement;
const propertiesContent = document.getElementById("properties-content") as HTMLDivElement;
const propertiesClose = document.getElementById("properties-close") as HTMLButtonElement;
const modelListEl = document.getElementById("model-list") as HTMLDivElement;
const treeToggle = document.getElementById("tree-toggle") as HTMLButtonElement;
const treePanel = document.getElementById("tree-panel") as HTMLDivElement;
const treeClose = document.getElementById("tree-close") as HTMLButtonElement;
const treeRefresh = document.getElementById("tree-refresh") as HTMLButtonElement;
const treeTabs = document.getElementById("tree-tabs") as HTMLDivElement;
const treeContent = document.getElementById("tree-content") as HTMLDivElement;
const visibilityHide = document.getElementById("visibility-hide") as HTMLButtonElement;
const visibilityIsolate = document.getElementById("visibility-isolate") as HTMLButtonElement;
const visibilityShowAll = document.getElementById("visibility-show-all") as HTMLButtonElement;
const visibilityXray = document.getElementById("visibility-xray") as HTMLButtonElement;
const zoomExtents = document.getElementById("zoom-extents") as HTMLButtonElement;
const zoomSelected = document.getElementById("zoom-selected") as HTMLButtonElement;
const duplicatesToggle = document.getElementById("duplicates-toggle") as HTMLButtonElement;
const duplicatesPanel = document.getElementById("duplicates-panel") as HTMLDivElement;
const duplicatesClose = document.getElementById("duplicates-close") as HTMLButtonElement;
const duplicatesRefresh = document.getElementById("duplicates-refresh") as HTMLButtonElement;
const duplicatesXray = document.getElementById("duplicates-xray") as HTMLButtonElement;
const duplicatesSortSelect = document.getElementById("duplicates-sort-select") as HTMLSelectElement;
const duplicatesToleranceInput = document.getElementById(
  "duplicates-tolerance-input",
) as HTMLInputElement;
const duplicatesToleranceValue = document.getElementById(
  "duplicates-tolerance-value",
) as HTMLSpanElement;
const duplicatesStatus = document.getElementById("duplicates-status") as HTMLSpanElement;
const duplicatesContent = document.getElementById("duplicates-content") as HTMLDivElement;
const clipToggle = document.getElementById("clip-toggle") as HTMLButtonElement;
const clipClear = document.getElementById("clip-clear") as HTMLButtonElement;
const measureToggle = document.getElementById("measure-toggle") as HTMLButtonElement;
const measureClear = document.getElementById("measure-clear") as HTMLButtonElement;
const takeoffToggle = document.getElementById("takeoff-toggle") as HTMLButtonElement;
const takeoffPanel = document.getElementById("takeoff-panel") as HTMLDivElement;
const takeoffClose = document.getElementById("takeoff-close") as HTMLButtonElement;
const takeoffRefresh = document.getElementById("takeoff-refresh") as HTMLButtonElement;
const takeoffAutoClassify = document.getElementById("takeoff-auto-classify") as HTMLButtonElement;
const takeoffExport = document.getElementById("takeoff-export") as HTMLButtonElement;
const takeoffExportKlimat = document.getElementById("takeoff-export-klimat") as HTMLButtonElement;
const takeoffExportExcludeUnclassified = document.getElementById(
  "takeoff-export-exclude-unclassified",
) as HTMLInputElement;
const takeoffStatus = document.getElementById("takeoff-status") as HTMLSpanElement;
const takeoffTableWrapper = document.getElementById("takeoff-table-wrapper") as HTMLDivElement;
const takeoffColumnsList = document.getElementById("takeoff-columns-list") as HTMLDivElement;
const takeoffAddColumnSelect = document.getElementById(
  "takeoff-add-column-select",
) as HTMLSelectElement;
const takeoffAddColumnBtn = document.getElementById("takeoff-add-column-btn") as HTMLButtonElement;
const takeoffBulkBar = document.getElementById("takeoff-bulk-bar") as HTMLDivElement;
const takeoffBulkCount = document.getElementById("takeoff-bulk-count") as HTMLSpanElement;
const takeoffBulkExclude = document.getElementById("takeoff-bulk-exclude") as HTMLButtonElement;
const takeoffBulkClear = document.getElementById("takeoff-bulk-clear") as HTMLButtonElement;
const generateToggle = document.getElementById("generate-toggle") as HTMLButtonElement;
const generatePanel = document.getElementById("generate-panel") as HTMLDivElement;
const generateClose = document.getElementById("generate-close") as HTMLButtonElement;
const footprintCanvas = document.getElementById("footprint-canvas") as HTMLCanvasElement;
const footprintInfo = document.getElementById("footprint-info") as HTMLSpanElement;
const footprintUndo = document.getElementById("footprint-undo") as HTMLButtonElement;
const footprintClear = document.getElementById("footprint-clear") as HTMLButtonElement;
const footprintEdgeOverlay = document.getElementById("footprint-edge-overlay") as HTMLDivElement;
const footprintZoomIn = document.getElementById("footprint-zoom-in") as HTMLButtonElement;
const footprintZoomOut = document.getElementById("footprint-zoom-out") as HTMLButtonElement;
const footprintZoomLevel = document.getElementById("footprint-zoom-level") as HTMLSpanElement;
const generateFloorsInput = document.getElementById("generate-floors") as HTMLInputElement;
const generateHeightInput = document.getElementById("generate-height") as HTMLInputElement;
const generateWindowRatioInput = document.getElementById(
  "generate-window-ratio",
) as HTMLInputElement;
const generateSubmit = document.getElementById("generate-submit") as HTMLButtonElement;
const generateStatus = document.getElementById("generate-status") as HTMLDivElement;

const components = new OBC.Components();

const worlds = components.get(OBC.Worlds);
const world = worlds.create<
  OBC.SimpleScene,
  OBC.OrthoPerspectiveCamera,
  OBF.RendererWith2D
>();

world.scene = new OBC.SimpleScene(components);
const postproductionRenderer = new OBF.PostproductionRenderer(components, container);
world.renderer = postproductionRenderer;
world.renderer.showLogo = false;
world.camera = new OBC.OrthoPerspectiveCamera(components);

// Klick/dubbelklick för 3D-interaktion binds till själva canvasen (inte
// den omslutande containern), annars bubblar klick på toolbar-knappar och
// paneler upp och triggar oavsiktliga raycasts/markeringar i scenen.
const canvas = world.renderer.three.domElement;

components.init();

// Demand rendering: only render when the scene actually changes.
// turnOffOnManualMode (default true) disables the expensive COLOR_PEN
// post-processing during MANUAL-mode frames and re-enables it 50 ms after the
// last scene change, giving a fast preview during interaction and a sharp
// still image when the camera comes to rest.
postproductionRenderer.mode = OBC.RendererMode.MANUAL;
postproductionRenderer.needsUpdate = true;
// Cap render resolution on high-DPI screens — halves GPU work on Retina.
postproductionRenderer.three.setPixelRatio(Math.min(window.devicePixelRatio, 1.5));

world.scene.setup();
world.scene.three.background = null;

await world.camera.controls.setLookAt(15, 15, 15, 0, 0, 0);

// Mittenknappen (mouse3) ska panorera precis som högerklick redan gör -
// vanlig CAD-konvention, och bekvämt när högerklick också används för annat.
world.camera.controls.mouseButtons.middle = world.camera.controls.mouseButtons.right;

// COLOR_PEN lägger på ritade kantlinjer ovanpå den vanliga färgade
// skuggningen, så att intilliggande ytor med liknande gråton (t.ex. platta
// tak-/väggpaneler) går att skilja åt utan att behöva rotera kameran. Sätts
// efter att kameran är klar - postproduktionens djup-/kant-pass behöver den.
// Standardfärgen (#888) är nästan osynlig mot modellens egen gråa skuggning,
// så kantlinjerna sätts uttryckligen till svart. width är en texel-baserad
// samplingsoffset i kantdetekterings-shadern (inte pixlar rakt av) - 1 är
// bibliotekets eget default men gav en ganska tjock linje i praktiken, så den
// är sänkt till 0.5 för en tunnare kontur.
postproductionRenderer.postproduction.enabled = true;
postproductionRenderer.postproduction.style = OBF.PostproductionAspect.COLOR_PEN;
postproductionRenderer.postproduction.edgesPass.color.setHex(0x000000);
postproductionRenderer.postproduction.edgesPass.width = 0.1; //Linjetjocklek - default 1, men det blev för tjock

//components.get(OBC.Grids).create(world);

// IFC -> Fragments
const ifcLoader = components.get(OBC.IfcLoader);
await ifcLoader.setup({
  autoSetWasm: false,
  wasm: {
    path: "https://unpkg.com/web-ifc@0.0.77/",
    absolute: true,
  },
});

const workerUrl = await OBC.FragmentsManager.getWorker();
const fragments = components.get(OBC.FragmentsManager);
fragments.init(workerUrl);

world.camera.controls.addEventListener("update", () => {
  fragments.core.update();
  postproductionRenderer.needsUpdate = true;
});

const modelNames = new Map<string, string>();
// Rå IFC (STEP) -text per modell, sparad vid inläsning/generering så att
// flera modeller senare kan slås ihop och exporteras till en gemensam fil.
const modelIfcText = new Map<string, string>();

interface IfcHeaderInfo {
  originatingSystem?: string;
  schema?: string;
  author?: string;
  organization?: string;
  timestamp?: string;
  description?: string;
}

/** Plockar ut innehållet mellan de balanserade parenteserna direkt efter
 *  "KEYWORD(" - en regex med icke-girig matchning duger inte här eftersom
 *  fälten (t.ex. FILE_NAME:s författar-/organisationslistor) själva
 *  innehåller nästlade parenteser. */
function extractStepCall(text: string, keyword: string): string | null {
  const idx = text.indexOf(`${keyword}(`);
  if (idx === -1) return null;
  let depth = 0;
  let start = -1;
  let inString = false;
  for (let i = idx + keyword.length; i < text.length; i++) {
    const ch = text[i];
    if (ch === "'") {
      inString = !inString;
      continue;
    }
    if (inString) continue;
    if (ch === "(") {
      if (depth === 0) start = i + 1;
      depth++;
    } else if (ch === ")") {
      depth--;
      if (depth === 0) return text.slice(start, i);
    }
  }
  return null;
}

/** Delar upp en STEP-argumentlista på toppnivå-kommatecken, utan att gå
 *  sönder på kommatecken inuti nästlade listor eller citerade strängar. */
function splitStepArgs(argsText: string): string[] {
  const args: string[] = [];
  let depth = 0;
  let current = "";
  let inString = false;
  for (const ch of argsText) {
    if (ch === "'") {
      inString = !inString;
      current += ch;
      continue;
    }
    if (!inString) {
      if (ch === "(") depth++;
      if (ch === ")") depth--;
      if (ch === "," && depth === 0) {
        args.push(current.trim());
        current = "";
        continue;
      }
    }
    current += ch;
  }
  if (current.trim()) args.push(current.trim());
  return args;
}

function unquoteStep(value: string | undefined): string | undefined {
  const match = value?.trim().match(/^'(.*)'$/);
  return match ? match[1] : undefined;
}

/** Returnerar första strängen i ett STEP-listargument, t.ex. "('Ola')" -> "Ola". */
function firstStepListItem(value: string | undefined): string | undefined {
  const trimmed = value?.trim();
  if (!trimmed) return undefined;
  const listMatch = trimmed.match(/^\(([\s\S]*)\)$/);
  if (!listMatch) return unquoteStep(trimmed);
  return unquoteStep(splitStepArgs(listMatch[1])[0]);
}

/** Läser HEADER-sektionen ur en rå IFC-fil (STEP/SPF) - bl.a. vilken
 *  programvara modellen exporterades från (FILE_NAME:s "originating system",
 *  fält 6) och vilket IFC-schema (FILE_SCHEMA). Inte en fullt spec-enlig
 *  STEP-parser, men tillräcklig för header-fält som alltid är citerade
 *  strängar eller enkla listor av dem. */
function parseIfcHeader(text: string): IfcHeaderInfo | null {
  const headerEnd = text.indexOf("ENDSEC", text.indexOf("HEADER"));
  if (headerEnd === -1) return null;
  const headerText = text.slice(0, headerEnd);

  const result: IfcHeaderInfo = {};

  const fileName = extractStepCall(headerText, "FILE_NAME");
  if (fileName) {
    const args = splitStepArgs(fileName);
    result.timestamp = unquoteStep(args[1]);
    result.author = firstStepListItem(args[2]);
    result.organization = firstStepListItem(args[3]);
    result.originatingSystem = unquoteStep(args[5]);
  }

  const fileSchema = extractStepCall(headerText, "FILE_SCHEMA");
  if (fileSchema) result.schema = firstStepListItem(fileSchema);

  const fileDescription = extractStepCall(headerText, "FILE_DESCRIPTION");
  if (fileDescription) result.description = firstStepListItem(splitStepArgs(fileDescription)[0]);

  // FILE_NAME:s "originating system" är ofta bara exportverktygets egen
  // versionssträng (t.ex. "21.1.20.44 - Exporter..."), inte själva
  // programnamnet. IFCAPPLICATION-posten i DATA-sektionen (dess
  // ApplicationFullName, 3:e fältet) är den plats IFC-spec:en avsett för det
  // läsbara programnamnet, och ger typiskt ett mycket tydligare svar (t.ex.
  // "Autodesk Revit 2021 (ENU)") - används i första hand när den finns.
  const appMatch = text.match(/IFCAPPLICATION\([^,]*,\s*'([^']*)'\s*,\s*'([^']*)'\s*,\s*'([^']*)'\s*\)/);
  if (appMatch) {
    const [, , applicationFullName] = appMatch;
    if (applicationFullName) result.originatingSystem = applicationFullName;
  }

  return result;
}

function formatIfcHeaderSummary(header: IfcHeaderInfo | null): { line: string; tooltip: string } | null {
  if (!header) return null;
  const mainParts = [header.originatingSystem, header.schema].filter(Boolean);
  if (mainParts.length === 0 && !header.author && !header.organization) return null;

  const tooltipParts = [
    header.originatingSystem ? `Program: ${header.originatingSystem}` : null,
    header.schema ? `Schema: ${header.schema}` : null,
    header.author ? `Författare: ${header.author}` : null,
    header.organization ? `Organisation: ${header.organization}` : null,
    header.timestamp ? `Skapad: ${header.timestamp}` : null,
    header.description ? `Beskrivning: ${header.description}` : null,
  ].filter((part): part is string => part !== null);

  return {
    line: mainParts.length > 0 ? mainParts.join(" · ") : "IFC-header",
    tooltip: tooltipParts.join("\n"),
  };
}

fragments.list.onItemSet.add(({ value: model }) => {
  model.useCamera(world.camera.three);
  world.scene.three.add(model.object);
  fragments.core.update(true);
  postproductionRenderer.needsUpdate = true;
  renderModelList();
  treeBuiltMode = null;
  duplicatesBuilt = false;
  duplicateCandidatesByCategory = null;
  // OBS: rendera INTE trädet här - onItemSet triggas innan modellen är
  // redo i worker-tråden (bekräftat via ett "Fragments: Model not found"-fel
  // vid test), så trädbygget måste vänta tills den faktiska
  // ifcLoader.load()-anropet är klart (se loadOrReplaceSyncedFile/
  // generateSubmit, som anropar renderModelTreeIfOpen() efteråt).
});

fragments.list.onBeforeDelete.add(({ value: model }) => {
  world.scene.three.remove(model.object);
  postproductionRenderer.needsUpdate = true;
});

fragments.list.onItemDeleted.add((modelId) => {
  modelNames.delete(modelId);
  modelIfcText.delete(modelId);
  renderModelList();
  treeBuiltMode = null;
  renderModelTreeIfOpen();
  duplicatesBuilt = false;
  duplicateCandidatesByCategory = null;
});

let modelCount = 0;

// Källfiler - håller fristående IFC-filer, valda en och en (oavsett vilken
// mapp de ligger i), synkade som modeller här - manuellt via "Skanna om"
// eller automatiskt när appen öppnas / fliken får fokus.
let syncedFiles: SyncedFile[] = [];
let filesAutoUpdate = false;
let filesSyncInProgress = false;

function syncedModelId(id: string): string {
  return `syncfile:${id}`;
}

function updateFilesControlsEnabled() {
  filesRescanButton.disabled = syncedFiles.length === 0;
}

/** Global progressindikator i statusraden - används vid modellinläsning
 *  (obestämd längd, se anropen i loadOrReplaceSyncedFile/generateSubmit) och
 *  mängdberäkning (bestämd längd/procent, se computeQuantityTakeoff). Utan
 *  angivet `percent` visas en glidande "obestämd längd"-stapel istället för
 *  en fylld andel, eftersom modellinläsningens faktiska förlopp inte
 *  exponeras av det underliggande biblioteket (bara start/klart). */
function setProgress(label: string, percent?: number): void {
  statusProgress.classList.remove("hidden");
  statusProgressLabel.textContent = label;
  if (percent === undefined) {
    statusProgress.classList.add("indeterminate");
    statusProgressFill.style.width = "";
  } else {
    statusProgress.classList.remove("indeterminate");
    statusProgressFill.style.width = `${Math.min(100, Math.max(0, percent)).toFixed(0)}%`;
  }
}

function clearProgress(): void {
  statusProgress.classList.add("hidden");
  statusProgress.classList.remove("indeterminate");
}

async function loadOrReplaceSyncedFile(entry: SyncedFile): Promise<void> {
  const modelId = syncedModelId(entry.id);
  const file = await entry.handle.getFile();
  const buffer = new Uint8Array(await file.arrayBuffer());

  if (fragments.list.get(modelId)) {
    await fragments.core.disposeModel(modelId);
  }

  modelNames.set(modelId, entry.name);
  modelIfcText.set(modelId, new TextDecoder("utf-8").decode(buffer));
  setProgress(`Laddar ${entry.name}...`);
  try {
    // coordinate: true - låter fragments-biblioteket förskjuta denna modell
    // relativt den först inlästa modellens sanna koordinater, så att flera
    // modeller som delar site hamnar rätt i förhållande till varandra istället
    // för att var och en nollställs oberoende.
    await ifcLoader.load(buffer, true, modelId, {});
  } finally {
    clearProgress();
  }
  renderModelTreeIfOpen();
}

/** Läser om varje spårad fil och laddar in de som ändrats sedan senaste synk. */
async function syncFiles(): Promise<{ changed: number }> {
  if (filesSyncInProgress || syncedFiles.length === 0) return { changed: 0 };
  filesSyncInProgress = true;

  try {
    let changed = 0;
    for (const entry of syncedFiles) {
      const meta = await readFileMeta(entry.handle);
      const isModified = meta.lastModified !== entry.lastModified || meta.size !== entry.size;
      if (!isModified) continue;

      filesStatus.textContent = `Laddar ${entry.name}...`;
      entry.lastModified = meta.lastModified;
      entry.size = meta.size;
      await loadOrReplaceSyncedFile(entry);
      changed++;
    }

    return { changed };
  } finally {
    filesSyncInProgress = false;
  }
}

filesAddButton.addEventListener("click", async () => {
  try {
    const handles = await window.showOpenFilePicker({
      multiple: true,
      types: [{ description: "IFC-filer", accept: { "application/octet-stream": [".ifc"] } }],
    });

    for (const handle of handles) {
      // Om filen redan spåras (t.ex. återvald efter att behörigheten gick
      // förlorad mellan sessioner) - uppdatera dess handtag istället för att
      // lägga till en dubblett.
      let existing: SyncedFile | undefined;
      for (const entry of syncedFiles) {
        if (await handle.isSameEntry(entry.handle)) {
          existing = entry;
          break;
        }
      }

      const meta = await readFileMeta(handle);
      const entry: SyncedFile = existing ?? {
        id: crypto.randomUUID(),
        handle,
        name: handle.name,
        lastModified: meta.lastModified,
        size: meta.size,
      };
      entry.handle = handle;
      entry.lastModified = meta.lastModified;
      entry.size = meta.size;

      filesStatus.textContent = `Laddar ${entry.name}...`;
      await loadOrReplaceSyncedFile(entry);
      if (!existing) syncedFiles.push(entry);
    }

    await saveSourceFiles(syncedFiles.map(({ id, handle }) => ({ id, handle })));
    updateFilesControlsEnabled();
    filesStatus.textContent = `${syncedFiles.length} fil(er) synkade`;
  } catch (err) {
    if ((err as DOMException)?.name !== "AbortError") {
      console.error(err);
      filesStatus.textContent = "Kunde inte lägga till filerna.";
    }
  }
});

filesRescanButton.addEventListener("click", async () => {
  if (syncedFiles.length === 0) return;
  filesStatus.textContent = "Skannar om...";
  const { changed } = await syncFiles();
  filesStatus.textContent = `Uppdaterad - ${changed} ändrad(e)`;
});

filesAutoUpdateCheckbox.addEventListener("change", () => {
  filesAutoUpdate = filesAutoUpdateCheckbox.checked;
  if (filesAutoUpdate) void syncFiles();
});

async function autoSyncFilesIfEnabled() {
  if (!filesAutoUpdate || syncedFiles.length === 0) return;
  const { changed } = await syncFiles();
  if (changed) filesStatus.textContent = `Automatiskt uppdaterad - ${changed} ändrad(e)`;
}

window.addEventListener("focus", () => void autoSyncFilesIfEnabled());
document.addEventListener("visibilitychange", () => {
  if (document.visibilityState === "visible") void autoSyncFilesIfEnabled();
});

// Vid uppstart: försök återansluta till tidigare valda filer. Behörighet kan
// bara begäras (requestPermission) från en användarinitierad handling, så
// vid sidladdning kan vi bara TYSTLÄTT kontrollera (queryPermission) om den
// redan gäller.
(async () => {
  const savedSourceFiles = await loadSourceFiles();
  if (savedSourceFiles.length === 0) return;

  filesStatus.textContent = "Återansluter till tidigare filer...";
  let needsReconnect = false;

  for (const { id, handle } of savedSourceFiles) {
    const permission = await handle.queryPermission({ mode: "read" });
    if (permission !== "granted") {
      needsReconnect = true;
      continue;
    }
    const meta = await readFileMeta(handle);
    const entry: SyncedFile = {
      id,
      handle,
      name: handle.name,
      lastModified: meta.lastModified,
      size: meta.size,
    };
    await loadOrReplaceSyncedFile(entry);
    syncedFiles.push(entry);
  }

  updateFilesControlsEnabled();
  filesStatus.textContent = needsReconnect
    ? `${syncedFiles.length} fil(er) återanslutna - klicka "Lägg till fil(er)" och välj om de saknade filerna för att återfå åtkomst.`
    : `${syncedFiles.length} fil(er) återanslutna`;
})();

// Klick på objekt -> markera + visa egenskaper
const casters = components.get(OBC.Raycasters);
const caster = casters.get(world);

const HIGHLIGHT_MATERIAL: FRAGS.MaterialDefinition = {
  color: new THREE.Color("#ff6b00"),
  renderedFaces: 1,
  opacity: 1,
  transparent: false,
};

/** Röntgenvy för dubbletter: hela modellen görs nästan genomskinlig utom de
 *  hittade dubblettobjekten, som lyses upp i en tydlig färg - så det syns
 *  VAR i byggnaden de sitter, med resten kvar som rumslig referens (till
 *  skillnad från "Isolera", som gömmer allt annat helt). */
const XRAY_GHOST_MATERIAL: FRAGS.MaterialDefinition = {
  color: new THREE.Color("#c7ccd1"),
  renderedFaces: 1,
  opacity: 0.12,
  transparent: true,
  depthWrite: false,
};
/** Cyklisk palett så olika dubblettgrupper går att skilja åt visuellt i
 *  röntgenvyn istället för att alla lysas upp i samma orange färg (som gjorde
 *  det svårt att se vilken dubblett som hörde ihop med vilken). Grupper utöver
 *  paletten återanvänder färger i tur och ordning (som en diagramlegend) -
 *  samma index används både för highlight-materialet och för färgprickarna i
 *  dubblettlistan (se DuplicateGroup.colorIndex), så de går att koppla ihop. */
const XRAY_DUPLICATE_GROUP_COLORS = [
  "#f97316",
  "#3b82f6",
  "#22c55e",
  "#a855f7",
  "#ec4899",
  "#eab308",
  "#06b6d4",
  "#ef4444",
  "#84cc16",
  "#6366f1",
];

function duplicateGroupColor(colorIndex: number): string {
  return XRAY_DUPLICATE_GROUP_COLORS[colorIndex % XRAY_DUPLICATE_GROUP_COLORS.length];
}

function duplicateGroupMaterial(colorIndex: number): FRAGS.MaterialDefinition {
  return {
    color: new THREE.Color(duplicateGroupColor(colorIndex)),
    renderedFaces: 1,
    opacity: 1,
    transparent: false,
  };
}

interface SelectableItem {
  modelId: string;
  localId: number;
  /** Redan uträknade mängder (t.ex. från mängdavtagningen), så en flerval-summering
   *  slipper räkna om geometrin på nytt för varje objekt. */
  metrics?: ItemMetrics | null;
}

function itemKey(item: SelectableItem): string {
  return `${item.modelId}::${item.localId}`;
}

/**
 * Slår ihop klickade objekt med nuvarande markering enligt samma konvention som
 * filutforskare: "replace" (vanligt klick) ersätter markeringen, "toggle"
 * (Ctrl/Cmd-klick) lägger till eller tar bort - om alla klickade objekt redan är
 * markerade tas de bort, annars läggs de som saknas till.
 */
function resolveMultiSelect(
  current: SelectableItem[],
  clicked: SelectableItem[],
  mode: "replace" | "toggle",
): SelectableItem[] {
  if (mode === "replace") return clicked;

  const currentKeys = new Set(current.map(itemKey));
  const clickedKeys = new Set(clicked.map(itemKey));
  const allSelected = clicked.every((item) => currentKeys.has(itemKey(item)));

  if (allSelected) return current.filter((item) => !clickedKeys.has(itemKey(item)));

  const merged = [...current];
  for (const item of clicked) {
    if (!currentKeys.has(itemKey(item))) merged.push(item);
  }
  return merged;
}

let highlightedItems: SelectableItem[] = [];

function escapeHtml(value: string): string {
  return value.replace(
    /[&<>"']/g,
    (char) =>
      ({
        "&": "&amp;",
        "<": "&lt;",
        ">": "&gt;",
        '"': "&quot;",
        "'": "&#39;",
      })[char] as string,
  );
}

function formatValue(value: unknown): string {
  if (value === null || value === undefined || value === "") return "-";
  if (typeof value === "object") return escapeHtml(JSON.stringify(value));
  return escapeHtml(String(value));
}

function isAttribute(value: unknown): value is FRAGS.ItemAttribute {
  return (
    !!value &&
    typeof value === "object" &&
    !Array.isArray(value) &&
    "value" in (value as Record<string, unknown>)
  );
}

function renderRow(key: string, value: unknown): string {
  return `<tr><td class="prop-key">${escapeHtml(key)}</td><td class="prop-val">${formatValue(value)}</td></tr>`;
}

/** Wrappar redan renderade renderRow()-rader i en riktig &lt;table&gt; - rader
 *  får aldrig lämnas lösa utanför en table/tbody (ogiltig HTML som webbläsare
 *  tyst flyttar ut ur sin plats). */
function renderPropTable(rows: string[]): string {
  return `<table class="prop-table"><tbody>${rows.join("")}</tbody></table>`;
}

function renderPset(pset: FRAGS.ItemData): string {
  const nameAttr = pset.Name;
  const name = isAttribute(nameAttr) ? String(nameAttr.value) : "Property set";
  const properties = pset.HasProperties;
  const rows: string[] = [];

  if (Array.isArray(properties)) {
    for (const prop of properties) {
      const propNameAttr = prop.Name;
      const propName = isAttribute(propNameAttr) ? String(propNameAttr.value) : "-";
      const valueAttr = prop.NominalValue;
      const propValue = isAttribute(valueAttr) ? valueAttr.value : undefined;
      rows.push(renderRow(propName, propValue));
    }
  }

  return `
    <div class="prop-section">
      <div class="prop-section-title">${escapeHtml(name)}</div>
      ${rows.length > 0 ? renderPropTable(rows) : '<div class="prop-empty">Inga egenskaper</div>'}
    </div>
  `;
}

function attrValue(item: FRAGS.ItemData | undefined, key: string): unknown {
  if (!item) return undefined;
  const attr = item[key];
  return isAttribute(attr) ? attr.value : undefined;
}

function itemName(item: FRAGS.ItemData | undefined): string | undefined {
  const value = attrValue(item, "Name");
  return typeof value === "string" && value ? value : undefined;
}

function itemCategory(item: FRAGS.ItemData | undefined): string {
  const value = attrValue(item, "_category");
  return typeof value === "string" ? value : "";
}

function relationItems(item: FRAGS.ItemData, key: string): FRAGS.ItemData[] {
  const value = item[key];
  return Array.isArray(value) ? value : [];
}

/** Sammanfattar en materialkoppling (ett HasAssociations-objekt av
 *  materialtyp) till en läsbar rad. Hanterar de vanligaste formerna (enkelt
 *  material, skiktuppsättning, lista, sammansättning, profiluppsättning) och
 *  faller tillbaka på namn/kategori för ovanligare typer istället för att
 *  tyst hoppa över dem. */
function summarizeMaterial(item: FRAGS.ItemData): string {
  const category = itemCategory(item);

  if (category === "IFCMATERIALLAYERSETUSAGE") {
    const layerSet = relationItems(item, "ForLayerSet")[0];
    return layerSet ? summarizeMaterial(layerSet) : "Skiktuppsättning";
  }

  if (category === "IFCMATERIALLAYERSET") {
    const setName = itemName(item);
    const layerNames = relationItems(item, "MaterialLayers").map((layer) => {
      const material = relationItems(layer, "Material")[0];
      const materialName = itemName(material);
      const thickness = attrValue(layer, "LayerThickness");
      const thicknessLabel = typeof thickness === "number" ? ` (${(thickness / 1000).toFixed(3)} m)` : "";
      return `${materialName ?? "Okänt material"}${thicknessLabel}`;
    });
    return layerNames.length > 0 ? layerNames.join(", ") : (setName ?? "Skiktuppsättning");
  }

  if (category === "IFCMATERIALCONSTITUENTSET") {
    const names = relationItems(item, "MaterialConstituents").map((constituent) => {
      const material = relationItems(constituent, "Material")[0];
      return itemName(material) ?? itemName(constituent) ?? "Okänt material";
    });
    return names.length > 0 ? names.join(", ") : (itemName(item) ?? "Materialuppsättning");
  }

  if (category === "IFCMATERIALLIST") {
    const names = relationItems(item, "Materials").map((m) => itemName(m) ?? "Okänt material");
    return names.length > 0 ? names.join(", ") : "Materiallista";
  }

  if (category === "IFCMATERIALPROFILESETUSAGE") {
    const profileSet = relationItems(item, "ForProfileSet")[0];
    return profileSet ? summarizeMaterial(profileSet) : "Profiluppsättning";
  }

  if (category === "IFCMATERIALPROFILESET") {
    const names = relationItems(item, "MaterialProfiles").map((profile) => {
      const material = relationItems(profile, "Material")[0];
      return itemName(material) ?? "Okänt material";
    });
    return names.length > 0 ? names.join(", ") : (itemName(item) ?? "Profiluppsättning");
  }

  // IFCMATERIAL eller en ovanlig/okänd associationstyp - visa namnet/
  // kategorin istället för att tyst hoppa över den.
  return itemName(item) ?? (category || "Material");
}

function summarizeClassification(item: FRAGS.ItemData): string {
  const identification = attrValue(item, "Identification") ?? attrValue(item, "ItemReference");
  const name = itemName(item);
  if (identification && name) return `${identification} – ${name}`;
  return String(identification ?? name ?? "Klassificering");
}

const MATERIAL_CATEGORIES = new Set([
  "IFCMATERIAL",
  "IFCMATERIALLIST",
  "IFCMATERIALLAYERSETUSAGE",
  "IFCMATERIALLAYERSET",
  "IFCMATERIALLAYER",
  "IFCMATERIALCONSTITUENTSET",
  "IFCMATERIALCONSTITUENT",
  "IFCMATERIALPROFILESETUSAGE",
  "IFCMATERIALPROFILESET",
  "IFCMATERIALPROFILE",
]);

const CLASSIFICATION_CATEGORIES = new Set(["IFCCLASSIFICATION", "IFCCLASSIFICATIONREFERENCE"]);

/** Renderar Typ-, material- och klassificeringssektioner utifrån
 *  IsTypedBy-/HasAssociations-relationerna på objektet - data som tidigare
 *  varken hämtades in eller visades i egenskapspanelen (bara IsDefinedBy
 *  gjorde det). Okända/ovanliga kopplingstyper hamnar i en "Övriga
 *  kopplingar"-sektion istället för att tyst försvinna. */
function renderRelationsSections(data: FRAGS.ItemData): string {
  const sections: string[] = [];

  for (const typeItem of relationItems(data, "IsTypedBy")) {
    const typeName = itemName(typeItem) ?? "Typ";
    const rows: string[] = [];
    const predefinedType = attrValue(typeItem, "PredefinedType");
    if (predefinedType) rows.push(renderRow("PredefinedType", predefinedType));
    const tag = attrValue(typeItem, "Tag");
    if (tag) rows.push(renderRow("Tag", tag));

    sections.push(`
      <div class="prop-section">
        <div class="prop-section-title">Typ: ${escapeHtml(typeName)}</div>
        ${rows.length > 0 ? renderPropTable(rows) : '<div class="prop-empty">Inga attribut</div>'}
      </div>
    `);
    for (const pset of relationItems(typeItem, "HasPropertySets")) sections.push(renderPset(pset));
  }

  const materialLines: string[] = [];
  const classificationLines: string[] = [];
  const otherLines: string[] = [];

  for (const association of relationItems(data, "HasAssociations")) {
    const category = itemCategory(association);
    if (MATERIAL_CATEGORIES.has(category)) {
      materialLines.push(summarizeMaterial(association));
    } else if (CLASSIFICATION_CATEGORIES.has(category)) {
      classificationLines.push(summarizeClassification(association));
    } else {
      otherLines.push(itemName(association) ?? (category || "Koppling"));
    }
  }

  const renderLines = (title: string, lines: string[]) =>
    lines.length === 0
      ? ""
      : `
        <div class="prop-section">
          <div class="prop-section-title">${escapeHtml(title)}</div>
          <table class="prop-table"><tbody>
            ${lines.map((line) => `<tr><td class="prop-val" colspan="2">${escapeHtml(line)}</td></tr>`).join("")}
          </tbody></table>
        </div>
      `;

  sections.push(renderLines("Material", materialLines));
  sections.push(renderLines("Klassificering", classificationLines));
  sections.push(renderLines("Övriga kopplingar", otherLines));

  return sections.join("");
}

/** Måttlinjer/ytor/box i VÄRLDSKOORDINATER (redan transformerade via
 *  model.object.matrixWorld) för att visualisera respektive Beräknat-rad i
 *  3D-vyn - se showMetricVisualization. */
interface MetricsVisualization {
  thicknessLine: [THREE.Vector3, THREE.Vector3];
  widthLine: [THREE.Vector3, THREE.Vector3];
  lengthLine: [THREE.Vector3, THREE.Vector3];
  heightLine: [THREE.Vector3, THREE.Vector3];
  /** Fyra hörn för en förenklad bounding-rektangel av objektets dominanta
   *  yta - används BARA som reservlösning om ingen ytterkant alls hittades
   *  (bör i praktiken aldrig hända, se computeMetrics). Den riktiga Net
   *  Area-visualiseringen ritar faceTriangles, inte denna rektangel. */
  faceRect: [THREE.Vector3, THREE.Vector3, THREE.Vector3, THREE.Vector3];
  /** Den FAKTISKA triangulerade ytan (inte en bounding-rektangel) för
   *  objektets dominanta sida - samma trianglar som Net Area-talet summerar
   *  arean av, så en fylld yta här är alltid geometriskt konsekvent med det
   *  visade talet (inklusive ev. hål/öppningar, som syns som riktiga
   *  luckor i fyllningen). Platt lista, tre punkter per triangel. */
  faceTriangles: THREE.Vector3[];
  /** ALLA ytterkantssegment runt den dominanta ytan, inklusive kanten runt
   *  hål/öppningar (till skillnad från perimeterSegments, som bara har den
   *  yttre konturen) - ritas som kontur ovanpå faceTriangles. */
  faceOutlineSegments: THREE.Vector3[];
  /** De faktiska ytterkantssegmenten (par av punkter) för bara den YTTRE
   *  konturen (hål/öppningar borträknade) - hittade genom att plocka ut de
   *  triangelkanter som bara delas av EN triangel och sedan behålla bara den
   *  öglan med störst utbredning (se computeMetrics). Summan av deras
   *  längder är perimeter-talet. */
  perimeterSegments: THREE.Vector3[];
  /** ALLA objektets trianglar (inte bara en sida) i världskoordinater - den
   *  FAKTISKA geometrin som model.getItemsVolume mäter volymen av, så
   *  Volume-visualiseringen blir den riktiga formen (inklusive hål/
   *  urskärningar) istället för en missvisande bounding-box. Platt lista,
   *  tre punkter per triangel. */
  volumeTriangles: THREE.Vector3[];
  /** Objektets OBB-centrum - ankare för Volume-etiketten. */
  center: THREE.Vector3;
}

type MetricVisualizationKey =
  | "sideArea"
  | "volume"
  | "thickness"
  | "height"
  | "width"
  | "length"
  | "perimeter";

interface ItemMetrics {
  sideArea: number;
  volume: number;
  thickness: number;
  height: number;
  width: number;
  length: number;
  perimeter: number;
  /** Saknas när ItemMetrics byggs syntetiskt från redan cachade tal (t.ex.
   *  quantityRowToSelectable, som återanvänder mängdavtagningens siffror
   *  utan att räkna om geometrin) - se renderMetricRow/applySelection, som
   *  hanterar avsaknaden genom att inte visa några klickbara rader/mått. */
  visualization?: MetricsVisualization;
}

async function computeMetrics(
  model: FRAGS.FragmentsModel,
  localId: number,
): Promise<ItemMetrics | null> {
  const item = model.getItem(localId);
  const geometry = await item.getGeometry();
  if (!geometry) return null;

  const triangleGroups = await geometry.getTriangles();
  if (!triangleGroups) return null;

  // Cluster triangles by (sign-agnostic) face normal, so opposite-facing
  // parallel faces (e.g. front/back of a wall, door or window) end up in
  // the same bucket. The largest bucket's normal direction (picked below,
  // see "axes") anchors thicknessAxis - the actual Net Area number is
  // computed later from frontFaceTriangles (all triangles facing that ONE
  // direction, found via a dot-product threshold rather than an exact
  // rounded-key match - see there for why that matters for sloped/tapered
  // surfaces). While we're at it, collect every vertex (for the height/
  // thickness/width/length pass below) so we only walk the triangle data
  // twice in total.
  const faceClusters = new Map<string, { area: number; normal: THREE.Vector3 }>();
  const normal = new THREE.Vector3();
  const points: THREE.Vector3[] = [];
  let minY = Infinity;
  let maxY = -Infinity;

  for (const triangles of triangleGroups) {
    for (const triangle of triangles) {
      triangle.getNormal(normal);
      const key = `${Math.abs(normal.x).toFixed(2)},${Math.abs(normal.y).toFixed(2)},${Math.abs(normal.z).toFixed(2)}`;
      const area = triangle.getArea();
      const cluster = faceClusters.get(key);
      if (cluster) cluster.area += area;
      else faceClusters.set(key, { area, normal: normal.clone() });

      for (const point of [triangle.a, triangle.b, triangle.c]) {
        points.push(point);
        if (point.y < minY) minY = point.y;
        if (point.y > maxY) maxY = point.y;
      }
    }
  }

  const sortedClusters = [...faceClusters.values()].sort((a, b) => b.area - a.area);

  // Pick up to three mutually near-perpendicular face-normal directions
  // (biggest faces first). For a box-like element (wall, slab, window,
  // beam...) these are its three principal dimensions, found without
  // assuming any particular orientation in space.
  const axes: THREE.Vector3[] = [];
  for (const cluster of sortedClusters) {
    if (axes.length === 3) break;
    if (axes.every((axis) => Math.abs(axis.dot(cluster.normal)) < 0.5)) {
      axes.push(cluster.normal);
    }
  }
  while (axes.length < 3) {
    const fallback =
      axes.length === 0
        ? new THREE.Vector3(1, 0, 0)
        : new THREE.Vector3().crossVectors(axes[0], axes[1] ?? new THREE.Vector3(0, 1, 0));
    axes.push((fallback.lengthSq() < 1e-6 ? new THREE.Vector3(0, 0, 1) : fallback).normalize());
  }

  const axisData = axes.map((axis) => {
    let min = Infinity;
    let max = -Infinity;
    for (const point of points) {
      const projection = point.dot(axis);
      if (projection < min) min = projection;
      if (projection > max) max = projection;
    }
    return { axis, min, max, extent: max - min };
  });
  axisData.sort((a, b) => a.extent - b.extent);
  const [thicknessAxis, widthAxis, lengthAxis] = axisData;
  const [thickness, width, length] = axisData.map((a) => a.extent);

  // OBB-centrum: rekonstruerat från de tre (ungefär ortogonala) axlarnas
  // mittprojektion - ankare för måttvisualiseringarna nedan, i objektets EGET
  // lokala koordinatsystem (samma som triangeldatan ovan). Transformeras till
  // världskoordinater (toWorld) precis innan det sparas, eftersom modellens
  // egen placering i scenen (model.object.matrixWorld) annars inte räknas in.
  const obbCenter = new THREE.Vector3();
  for (const a of axisData) obbCenter.addScaledVector(a.axis, (a.min + a.max) / 2);

  const toWorld = (local: THREE.Vector3) => local.clone().applyMatrix4(model.object.matrixWorld);

  // Modellens egen kant-post-process (svarta konturlinjer, se edgesPass)
  // ritas ovanpå ALLT i skärmrymden efter att 3D-scenen renderats, oavsett
  // depthTest - en kontur/omkrets-linje som ligger EXAKT på objektets egen
  // yttersilhuett hamnar därför under den svarta konturen och syns nästan
  // inte. Lyfter ytterkantslinjerna en liten bit utåt (samma riktning som
  // thicknessAxis.axis, dvs. rakt ut från den valda ytan) så de hamnar
  // tydligt FRAMFÖR den riktiga ytan istället för att sammanfalla med den.
  const outlineLift = (local: THREE.Vector3) =>
    toWorld(local.clone().addScaledVector(thicknessAxis.axis, 0.03));

  const axisLine = (a: {
    axis: THREE.Vector3;
    extent: number;
  }): [THREE.Vector3, THREE.Vector3] => {
    const half = a.extent / 2;
    return [
      toWorld(obbCenter.clone().addScaledVector(a.axis, -half)),
      toWorld(obbCenter.clone().addScaledVector(a.axis, half)),
    ];
  };

  const centroid = new THREE.Vector3();
  for (const point of points) centroid.add(point);
  if (points.length > 0) centroid.divideScalar(points.length);

  // Net Area/Omkrets representeras av den fysiska ytan i ena änden av den
  // tunnaste axeln (t.ex. väggens/plattans faktiska sida), inte objektets
  // mittplan, så visualiseringen ligger dikt an mot en riktig yta.
  const faceCenter = obbCenter
    .clone()
    .addScaledVector(thicknessAxis.axis, -thicknessAxis.extent / 2);
  const halfW = widthAxis.axis.clone().multiplyScalar(widthAxis.extent / 2);
  const halfL = lengthAxis.axis.clone().multiplyScalar(lengthAxis.extent / 2);
  const faceRect: [THREE.Vector3, THREE.Vector3, THREE.Vector3, THREE.Vector3] = [
    toWorld(faceCenter.clone().sub(halfW).sub(halfL)),
    toWorld(faceCenter.clone().add(halfW).sub(halfL)),
    toWorld(faceCenter.clone().add(halfW).add(halfL)),
    toWorld(faceCenter.clone().sub(halfW).add(halfL)),
  ];

  // Net Area/Omkrets utgår båda från EN sida av objektet - trianglarna vars
  // normal pekar (ungefär) samma väg som thicknessAxis.axis, hittade via ett
  // dot-produkt-tröskelvärde snarare än en exakt avrundad nyckel-matchning
  // (som klustringen ovan använder). Det spelar roll för t.ex. ett tak med
  // lutande/fallande isolering (tapered insulation): undersidan är helt platt
  // (en enda normal, en stor klustring), men ovansidan lutar åt flera håll
  // för avvattning och triangulerar då till MÅNGA olika normaler som inte
  // avrundas till samma nyckel som undersidan - så att bara halvera den
  // största klustringen (som om den vore ett symmetriskt fram+bak-par) skulle
  // räkna undersidans area som HALVA sin verkliga storlek. Genom att summera
  // den faktiska arean av alla trianglar på thicknessAxis-sidan direkt blir
  // Net Area korrekt oavsett hur oregelbundet motsatta sidan tesselleras.
  const frontFaceTriangles: THREE.Triangle[] = [];
  for (const triangles of triangleGroups) {
    for (const triangle of triangles) {
      triangle.getNormal(normal);
      if (normal.dot(thicknessAxis.axis) > 0.5) frontFaceTriangles.push(triangle.clone());
    }
  }
  const sideArea = frontFaceTriangles.reduce((sum, triangle) => sum + triangle.getArea(), 0);

  const pointKey = (p: THREE.Vector3) => `${p.x.toFixed(4)},${p.y.toFixed(4)},${p.z.toFixed(4)}`;
  const edgeInfo = new Map<
    string,
    { count: number; a: THREE.Vector3; b: THREE.Vector3; length: number }
  >();
  for (const triangle of frontFaceTriangles) {
    const verts = [triangle.a, triangle.b, triangle.c];
    for (let i = 0; i < 3; i++) {
      const a = verts[i];
      const b = verts[(i + 1) % 3];
      const ka = pointKey(a);
      const kb = pointKey(b);
      const key = ka < kb ? `${ka}|${kb}` : `${kb}|${ka}`;
      const existing = edgeInfo.get(key);
      if (existing) existing.count += 1;
      else edgeInfo.set(key, { count: 1, a, b, length: a.distanceTo(b) });
    }
  }

  // Gränskanterna (count === 1) hör antingen till den YTTRE konturen eller
  // till kanten runt ett hål (dörr-/fönsteröppning). Grupperar dem till
  // separata slutna öglor via union-find på hörnpunkterna, och behåller bara
  // öglan med störst utbredning (bounding box-diagonal) - hål är per
  // definition mindre än (och innanför) den yttre konturen de sitter i, så
  // den största öglan är alltid den yttre. Öppningarnas kanter räknas alltså
  // INTE med i Omkrets, bara den yttre silhuetten.
  const parent = new Map<string, string>();
  const find = (key: string): string => {
    let root = key;
    while (parent.get(root) !== root) root = parent.get(root)!;
    let current = key;
    while (current !== root) {
      const next = parent.get(current)!;
      parent.set(current, root);
      current = next;
    }
    return root;
  };

  const boundaryEdges = [...edgeInfo.values()].filter((edge) => edge.count === 1);
  for (const edge of boundaryEdges) {
    const ka = pointKey(edge.a);
    const kb = pointKey(edge.b);
    if (!parent.has(ka)) parent.set(ka, ka);
    if (!parent.has(kb)) parent.set(kb, kb);
    const ra = find(ka);
    const rb = find(kb);
    if (ra !== rb) parent.set(ra, rb);
  }

  const loops = new Map<
    string,
    { edges: { a: THREE.Vector3; b: THREE.Vector3; length: number }[]; min: THREE.Vector3; max: THREE.Vector3 }
  >();
  for (const edge of boundaryEdges) {
    const root = find(pointKey(edge.a));
    let loop = loops.get(root);
    if (!loop) {
      loop = {
        edges: [],
        min: new THREE.Vector3(Infinity, Infinity, Infinity),
        max: new THREE.Vector3(-Infinity, -Infinity, -Infinity),
      };
      loops.set(root, loop);
    }
    loop.edges.push(edge);
    loop.min.min(edge.a).min(edge.b);
    loop.max.max(edge.a).max(edge.b);
  }

  const outerLoop = [...loops.values()].sort(
    (a, b) => b.max.distanceToSquared(b.min) - a.max.distanceToSquared(a.min),
  )[0];

  let outlinePerimeter = 0;
  const outlineSegmentsLocal: THREE.Vector3[] = [];
  for (const edge of outerLoop?.edges ?? []) {
    outlinePerimeter += edge.length;
    outlineSegmentsLocal.push(edge.a, edge.b);
  }

  // Om ingen triangel råkade matcha exakt (bör i praktiken aldrig hända,
  // eftersom thicknessAxis.axis självt kommer från en av trianglarnas egen
  // normal) - fall tillbaka på den enkla rektangelns kant som approximation.
  // faceRect ovan är redan i världskoordinater, så den listan behöver ingen
  // toWorld-transform, till skillnad från outlineSegmentsLocal.
  const perimeter = outlinePerimeter > 0 ? outlinePerimeter : 2 * (width + length);
  const perimeterSegments: THREE.Vector3[] =
    outlineSegmentsLocal.length > 0
      ? outlineSegmentsLocal.map((p) => outlineLift(p))
      : [
          faceRect[0],
          faceRect[1],
          faceRect[1],
          faceRect[2],
          faceRect[2],
          faceRect[3],
          faceRect[3],
          faceRect[0],
        ];

  // Net Area/Omkrets FYLLNING visas som den faktiska triangulerade ytan (inte
  // en bounding-rektangel) - exakt samma trianglar som sideArea summerade
  // arean av, så fyllningen alltid stämmer geometriskt med talet, hål och
  // urskärningar syns som riktiga luckor. faceOutlineSegments innehåller ALLA
  // gränskanter (yttre kontur + hål), till skillnad från perimeterSegments
  // ovan som bara har den yttre.
  const faceTriangles: THREE.Vector3[] = [];
  for (const triangle of frontFaceTriangles) {
    faceTriangles.push(toWorld(triangle.a), toWorld(triangle.b), toWorld(triangle.c));
  }
  const faceOutlineSegments: THREE.Vector3[] = [];
  for (const edge of boundaryEdges) {
    faceOutlineSegments.push(outlineLift(edge.a), outlineLift(edge.b));
  }


  // Volume visas som HELA objektets faktiska mesh (inte en omskriven box) -
  // samma geometri som model.getItemsVolume mäter volymen av, så
  // visualiseringen aldrig ser "fylligare" ut än vad talet faktiskt är
  // (t.ex. en vägg med ett hål ska inte se ut som en hel, tät box). "points"
  // samlades redan in ovan i triangelordning (a,b,c per triangel), så den
  // kan återanvändas rakt av istället för att gå igenom geometrin en gång
  // till.
  const volumeTriangles = points.map((point) => toWorld(point));

  const volume = await model.getItemsVolume([localId]);

  return {
    sideArea,
    volume,
    thickness,
    height: maxY - minY,
    width,
    length,
    perimeter,
    visualization: {
      thicknessLine: axisLine(thicknessAxis),
      widthLine: axisLine(widthAxis),
      lengthLine: axisLine(lengthAxis),
      heightLine: [
        toWorld(new THREE.Vector3(centroid.x, minY, centroid.z)),
        toWorld(new THREE.Vector3(centroid.x, maxY, centroid.z)),
      ],
      faceRect,
      faceTriangles,
      faceOutlineSegments,
      perimeterSegments,
      volumeTriangles,
      center: toWorld(obbCenter),
    },
  };
}

const METRIC_ROW_DEFS: { key: MetricVisualizationKey; label: string }[] = [
  { key: "sideArea", label: "Net Area" },
  { key: "volume", label: "Volume" },
  { key: "thickness", label: "Tjocklek" },
  { key: "height", label: "Höjd" },
  { key: "width", label: "Bredd" },
  { key: "length", label: "Längd" },
  { key: "perimeter", label: "Omkrets" },
];

/** Formaterar ett Beräknat-värde som text - används både för tabellraden
 *  (renderMetrics) och etiketten som skrivs ut i 3D-vyn (showMetricVisualization),
 *  så de alltid visar exakt samma tal. */
function formatMetricValue(metrics: ItemMetrics, key: MetricVisualizationKey): string {
  switch (key) {
    case "sideArea":
      return `${metrics.sideArea.toFixed(2)} m²`;
    case "volume":
      return `${metrics.volume.toFixed(2)} m³`;
    case "thickness":
      return `${metrics.thickness.toFixed(2)} m`;
    case "height":
      return `${metrics.height.toFixed(2)} m`;
    case "width":
      return `${metrics.width.toFixed(2)} m`;
    case "length":
      return `${metrics.length.toFixed(2)} m`;
    case "perimeter":
      return `${metrics.perimeter.toFixed(2)} m`;
  }
}

/** Rad i Beräknat-fliken. Klickbar (och visar motsvarande måttlinje/yta/box
 *  i 3D-vyn, se propertiesContent-klicklyssnaren) bara när geometrin faktiskt
 *  räknades om just nu (computeMetrics) - t.ex. urval som kommer från
 *  mängdavtagningens redan cachade tal (quantityRowToSelectable) har inga
 *  3D-punkter att visa, så den raden renderas som en vanlig, oklickbar rad. */
function renderMetricRow(
  metricKey: MetricVisualizationKey,
  label: string,
  valueText: string,
  clickable: boolean,
): string {
  if (!clickable) return renderRow(label, valueText);
  const active = activeMetricVisualizationKey === metricKey;
  return `<tr class="prop-metric-row${active ? " active" : ""}" data-metric="${metricKey}"><td class="prop-key">${escapeHtml(label)}</td><td class="prop-val">${escapeHtml(valueText)}</td></tr>`;
}

function renderMetrics(metrics: ItemMetrics | null): string {
  if (!metrics) {
    return `
      <div class="prop-section">
        <div class="prop-section-title">Beräknat</div>
        <div class="prop-empty">Ingen geometri</div>
      </div>
    `;
  }

  const clickable = !!metrics.visualization;
  return `
    <div class="prop-section">
      <div class="prop-section-title">Beräknat</div>
      ${renderPropTable(
        METRIC_ROW_DEFS.map(({ key, label }) =>
          renderMetricRow(key, label, formatMetricValue(metrics, key), clickable),
        ),
      )}
      ${clickable ? `<div class="prop-metric-hint">Klicka på en rad för att visa måttet i 3D-vyn</div>` : ""}
    </div>
  `;
}

function renderItemData(data: FRAGS.ItemData): string {
  const attributeRows: string[] = [];
  const psetGroups: string[] = [];

  for (const [key, value] of Object.entries(data)) {
    if (key === "_localId") continue;

    if (Array.isArray(value)) {
      if (key === "IsDefinedBy") {
        for (const pset of value) psetGroups.push(renderPset(pset));
      }
      continue;
    }
    if (isAttribute(value)) {
      const label = key.startsWith("_") ? key.slice(1) : key;
      attributeRows.push(renderRow(label, value.value));
    }
  }

  return `
    <div class="prop-section">
      <div class="prop-section-title">Attribut</div>
      ${attributeRows.length > 0 ? renderPropTable(attributeRows) : '<div class="prop-empty">Inga attribut</div>'}
    </div>
    ${psetGroups.join("")}
  `;
}

/** Statusradens vänstra fält (se #status-bar) - hålls i synk med
 *  highlightedItems överallt den ändras (clearHighlight/applySelection). */
function updateStatusSelection(): void {
  statusSelection.textContent =
    highlightedItems.length === 0 ? "Inget markerat" : `${highlightedItems.length} markerat`;
}

async function clearHighlight() {
  if (highlightedItems.length === 0) return;

  const byModel = new Map<string, number[]>();
  for (const { modelId, localId } of highlightedItems) {
    const list = byModel.get(modelId) ?? [];
    list.push(localId);
    byModel.set(modelId, list);
  }

  for (const [modelId, localIds] of byModel) {
    const model = fragments.list.get(modelId);
    await model?.resetHighlight(localIds);
  }

  highlightedItems = [];
  updateStatusSelection();
  await fragments.core.update(true);
  postproductionRenderer.needsUpdate = true;
}

async function clearSelection() {
  await clearHighlight();
  propertiesPanel.classList.add("hidden");
  propertiesContent.innerHTML = "";

  treeSelection = [];
  treeRangeAnchor = null;
  updateTreeRowSelectionStyles();

  takeoffSelection = [];
  takeoffRangeAnchor = null;
  updateTakeoffRowSelectionStyles();

  // clearHighlight() ovan återställde bara den tidigare markeringens material
  // till standard, inte till ghost-materialet - utan detta skulle det gamla
  // markerade objektet sluta vara genomskinligt när markeringen rensas medan
  // röntgenvyn är på.
  if (visibilityXrayActive) await applyVisibilityXray();
}

/** Grupperar de markerade objekten per modell - används av göm/isolera. */
function groupHighlightedByModel(): Map<string, number[]> {
  const byModel = new Map<string, number[]>();
  for (const { modelId, localId } of highlightedItems) {
    const list = byModel.get(modelId) ?? [];
    list.push(localId);
    byModel.set(modelId, list);
  }
  return byModel;
}

visibilityHide.addEventListener("click", async () => {
  const byModel = groupHighlightedByModel();
  if (byModel.size === 0) return;
  for (const [modelId, localIds] of byModel) {
    await fragments.list.get(modelId)?.setVisible(localIds, false);
  }
  // Ett gömt objekt kan inte längre vara meningsfullt markerat/highlightat.
  await clearSelection();
  await fragments.core.update(true);
  postproductionRenderer.needsUpdate = true;
});

visibilityIsolate.addEventListener("click", async () => {
  const byModel = groupHighlightedByModel();
  if (byModel.size === 0) return;
  for (const modelId of fragments.list.keys()) {
    const model = fragments.list.get(modelId);
    if (!model) continue;
    await model.setVisible(undefined, false);
    const keep = byModel.get(modelId);
    if (keep) await model.setVisible(keep, true);
  }
  await fragments.core.update(true);
  postproductionRenderer.needsUpdate = true;
});

visibilityShowAll.addEventListener("click", async () => {
  for (const modelId of fragments.list.keys()) {
    await fragments.list.get(modelId)?.resetVisible();
  }
  await fragments.core.update(true);
  postproductionRenderer.needsUpdate = true;
});

/** Röntgenvy för den vanliga markeringen (till skillnad från dubblettpanelens
 *  röntgenvy, som alltid visar dubbletterna). Ghostar allt UTOM det/de just
 *  nu markerade objekten - som Isolera, men behåller resten synligt som
 *  halvgenomskinlig rumslig referens istället för att gömma det helt. Håller
 *  sig uppdaterad om markeringen ändras medan den är på (se hooken i
 *  applySelection/clearSelection). */
let visibilityXrayActive = false;

async function applyVisibilityXray() {
  const selectedByModel = groupHighlightedByModel();
  for (const modelId of fragments.list.keys()) {
    const model = fragments.list.get(modelId);
    if (!model) continue;
    const allIds = await model.getItemsIdsWithGeometry();
    const selectedIds = new Set(selectedByModel.get(modelId) ?? []);
    const ghostIds = allIds.filter((id) => !selectedIds.has(id));
    if (ghostIds.length > 0) await model.highlight(ghostIds, XRAY_GHOST_MATERIAL);
  }
  visibilityXrayActive = true;
  visibilityXray.classList.add("active");
  await fragments.core.update(true);
  postproductionRenderer.needsUpdate = true;
}

async function clearVisibilityXray() {
  if (!visibilityXrayActive) return;

  for (const modelId of fragments.list.keys()) {
    await fragments.list.get(modelId)?.resetHighlight();
  }
  visibilityXrayActive = false;
  visibilityXray.classList.remove("active");

  // resetHighlight() ovan nollställde också en ev. pågående vanlig markering
  // - rita tillbaka den istället för att den tyst försvinner (samma mönster
  // som clearDuplicatesXray).
  if (highlightedItems.length > 0) {
    for (const [modelId, localIds] of groupHighlightedByModel()) {
      await fragments.list.get(modelId)?.highlight(localIds, HIGHLIGHT_MATERIAL);
    }
  }
  await fragments.core.update(true);
  postproductionRenderer.needsUpdate = true;
}

visibilityXray.addEventListener("click", async () => {
  if (visibilityXrayActive) {
    await clearVisibilityXray();
    return;
  }
  // Ömsesidigt uteslutande med dubblettpanelens röntgenvy - båda ghostar hela
  // modellen på olika sätt och skulle annars motverka varandra.
  if (duplicatesXrayActive) await clearDuplicatesXray();
  await applyVisibilityXray();
});

/** Ramar in en boundingbox genom att, precis som world.camera.fit() gör
 *  internt, konvertera den till ett klot (radie = största dimensionen) och
 *  anropa fitToSphere - inte fitToBox. En klotbaserad inramning är okänslig
 *  för betraktningsvinkeln (till skillnad från fitToBox, som gav en helt fel,
 *  extremt inzoomad vy för avlånga byggnader beroende på hur boxen råkade
 *  projiceras mot den aktuella kameravinkeln). */
async function fitCameraToBox(box: THREE.Box3, offset = 1.5) {
  const size = box.getSize(new THREE.Vector3());
  const center = box.getCenter(new THREE.Vector3());
  const radius = Math.max(size.x, size.y, size.z) * offset;
  await world.camera.controls.fitToSphere(new THREE.Sphere(center, radius), true);
}

// world.camera.fit() unionerar alltid in ALLA inlästa modellers fulla
// boundingbox (oavsett vilka meshes som skickas in), så den passar perfekt
// för "extents" med en tom lista - men går inte att använda för att zooma
// till bara ett urval (se zoomSelected nedan, som bygger sin egen boundingbox
// istället).
zoomExtents.addEventListener("click", async () => {
  await world.camera.fit([]);
});

/** Zoomar till de för tillfället highlightade/markerade objekten (t.ex. satta
 *  av applySelection). Delas mellan "Zoom markerat"-knappen och andra ställen
 *  där man vill hoppa till ett specifikt objekt, t.ex. "Gå till" i
 *  dubblettlistans högerklicksmeny. */
async function zoomToHighlighted() {
  if (highlightedItems.length === 0) return;
  const overallBox = new THREE.Box3();
  let hasBox = false;
  for (const [modelId, localIds] of groupHighlightedByModel()) {
    const model = fragments.list.get(modelId);
    if (!model || localIds.length === 0) continue;
    const box = await model.getMergedBox(localIds);
    if (box.isEmpty()) continue;
    overallBox.union(box.applyMatrix4(model.object.matrixWorld));
    hasBox = true;
  }
  if (hasBox) await fitCameraToBox(overallBox);
}

zoomSelected.addEventListener("click", zoomToHighlighted);

const ICON_EYE =
  '<svg class="btn-icon" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M2.062 12.348a1 1 0 0 1 0-.696 10.75 10.75 0 0 1 19.876 0 1 1 0 0 1 0 .696 10.75 10.75 0 0 1-19.876 0" /><circle cx="12" cy="12" r="3" /></svg>';
const ICON_EYE_OFF =
  '<svg class="btn-icon" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M10.733 5.076a10.744 10.744 0 0 1 11.205 6.575 1 1 0 0 1 0 .696 10.747 10.747 0 0 1-1.444 2.49" /><path d="M14.084 14.158a3 3 0 0 1-4.242-4.242" /><path d="M17.479 17.499a10.75 10.75 0 0 1-15.417-5.151 1 1 0 0 1 0-.696 10.75 10.75 0 0 1 4.446-5.143" /><path d="m2 2 20 20" /></svg>';
const ICON_TRASH =
  '<svg class="btn-icon" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M10 11v6" /><path d="M14 11v6" /><path d="M19 6v14a2 2 0 0 1-2 2H7a2 2 0 0 1-2-2V6" /><path d="M3 6h18" /><path d="M8 6V4a2 2 0 0 1 2-2h4a2 2 0 0 1 2 2v2" /></svg>';

function renderModelList() {
  const modelIds = [...fragments.list.keys()];

  if (modelIds.length === 0) {
    modelListEl.innerHTML = '<div class="model-empty">Inga modeller inlästa</div>';
    return;
  }

  modelListEl.innerHTML = modelIds
    .map((modelId) => {
      const model = fragments.list.get(modelId);
      const name = modelNames.get(modelId) ?? modelId;
      const visible = model ? model.object.visible : true;
      return `
        <div class="model-row" data-model-id="${escapeHtml(modelId)}">
          <span class="model-name" title="${escapeHtml(name)}">${escapeHtml(name)}</span>
          <button class="model-toggle" title="Visa/dölj">${visible ? ICON_EYE : ICON_EYE_OFF}</button>
          <button class="model-remove" title="Ta bort">${ICON_TRASH}</button>
        </div>
      `;
    })
    .join("");
}

modelListEl.addEventListener("click", async (event) => {
  const target = event.target as HTMLElement;
  const row = target.closest(".model-row") as HTMLElement | null;
  const modelId = row?.dataset.modelId;
  if (!modelId) return;

  const model = fragments.list.get(modelId);
  if (!model) return;

  if (target.closest(".model-toggle")) {
    model.object.visible = !model.object.visible;
    await fragments.core.update(true);
    postproductionRenderer.needsUpdate = true;
    renderModelList();
    return;
  }

  if (target.closest(".model-remove")) {
    if (highlightedItems.some((item) => item.modelId === modelId)) await clearSelection();
    await fragments.core.disposeModel(modelId);
    // Om modellen kom från en synkad källfil: glöm den ur senaste synken
    // också, annars ser nästa synk den som "oförändrad" och laddar inte in
    // den igen.
    if (modelId.startsWith("syncfile:")) {
      const id = modelId.slice("syncfile:".length);
      syncedFiles = syncedFiles.filter((entry) => entry.id !== id);
      await saveSourceFiles(syncedFiles.map(({ id: entryId, handle }) => ({ id: entryId, handle })));
      updateFilesControlsEnabled();
    }
  }
});

async function resolveMetrics(item: SelectableItem): Promise<ItemMetrics | null> {
  if (item.metrics !== undefined) return item.metrics;
  const model = fragments.list.get(item.modelId);
  if (!model) return null;
  try {
    return await computeMetrics(model, item.localId);
  } catch {
    return null;
  }
}

/** Flikarna i egenskapspanelen vid ett enda markerat objekt (se
 *  applySelection) - visas INTE vid flerval, eftersom bara Beräknat och
 *  Klassificering då har meningsfullt innehåll (Attribut/Typ & Material är
 *  per-objekt data som inte går att summera över flera olika objekt). */
type PropertiesTab = "attributes" | "metrics" | "relations" | "classification";
const PROPERTIES_TABS: { key: PropertiesTab; label: string }[] = [
  { key: "attributes", label: "Attribut" },
  { key: "metrics", label: "Beräknat" },
  { key: "relations", label: "Typ & Material" },
  { key: "classification", label: "Klassificering" },
];
/** Håller sig kvar mellan markeringar (t.ex. om man tittar på
 *  Klassificering-fliken och klickar ett annat objekt, stannar man kvar på
 *  samma flik) - samma princip som treeMode. */
let propertiesActiveTab: PropertiesTab = "attributes";

/** Beräknat-värdena för det just markerade enskilda objektet - se
 *  renderMetrics/showMetricVisualization (som både ritar 3D-geometrin via
 *  .visualization OCH skriver ut siffervärdet via formatMetricValue).
 *  Nollställs vid varje ny markering (applySelection), oavsett om det blir
 *  ett nytt enskilt objekt, flerval eller ingen markering alls. */
let currentMetrics: ItemMetrics | null = null;
let activeMetricVisualizationKey: MetricVisualizationKey | null = null;

function renderPropertiesTabs(active: PropertiesTab): string {
  return `
    <div id="properties-tabs">
      ${PROPERTIES_TABS.map(
        (tab) =>
          `<button class="prop-tab${tab.key === active ? " active" : ""}" data-prop-tab="${tab.key}">${escapeHtml(tab.label)}</button>`,
      ).join("")}
    </div>
  `;
}

/**
 * Highlightar och visar egenskaper för ett eller flera objekt. Ett enda objekt
 * visar full attribut-/pset-vy uppdelad i flikar (samma som tidigare
 * selectItem, se PROPERTIES_TABS), medan flera objekt visar en platt
 * summering (antal + summerade mängder, ingen flikrad) - samma princip som
 * gruppradarna i mängdavtagningen redan använde.
 */
async function applySelection(items: SelectableItem[], label?: string) {
  await clearHighlight();

  // Måttvisualiseringen (se renderMetricRow/showMetricVisualization) är
  // knuten till just detta enskilda objekt - nollställs vid VARJE ny
  // markering, oavsett om det blir ett nytt objekt, flerval eller inget alls.
  currentMetrics = null;
  activeMetricVisualizationKey = null;
  hideMetricVisualization();

  if (items.length === 0) {
    propertiesPanel.classList.add("hidden");
    propertiesContent.innerHTML = "";
    if (visibilityXrayActive) await applyVisibilityXray();
    return;
  }

  const byModel = new Map<string, number[]>();
  for (const { modelId, localId } of items) {
    const list = byModel.get(modelId) ?? [];
    list.push(localId);
    byModel.set(modelId, list);
  }

  highlightedItems = items;
  updateStatusSelection();
  for (const [modelId, localIds] of byModel) {
    const model = fragments.list.get(modelId);
    if (!model) continue;
    try {
      await model.highlight(localIds, HIGHLIGHT_MATERIAL);
    } catch {
      // Rent spatiala element (IfcProject/IfcSite/IfcBuilding/IfcBuildingStorey,
      // synliga via modellträdet) saknar egen geometri att highlighta.
    }
  }
  await fragments.core.update(true);
  postproductionRenderer.needsUpdate = true;

  // Håller röntgenvyn (Synlighet-sektionen) i synk med markeringen - se
  // applyVisibilityXray, som ghostar allt UTOM highlightedItems.
  if (visibilityXrayActive) await applyVisibilityXray();

  if (items.length === 1) {
    const { modelId, localId } = items[0];
    const model = fragments.list.get(modelId);
    if (!model) return;

    const [data] = await model.getItemsData([localId], {
      attributesDefault: true,
      relations: {
        IsDefinedBy: { attributes: true, relations: true },
        IsTypedBy: { attributes: true, relations: true },
        HasAssociations: { attributes: true, relations: true },
      },
    });
    const metrics = await resolveMetrics(items[0]);
    currentMetrics = metrics;

    const panes: Record<PropertiesTab, string> = {
      attributes: renderItemData(data),
      metrics: renderMetrics(metrics),
      relations: renderRelationsSections(data),
      classification: renderBoverketSelector(items) + renderGeneralSelector(items),
    };

    propertiesContent.innerHTML =
      renderPropertiesTabs(propertiesActiveTab) +
      PROPERTIES_TABS.map(
        (tab) =>
          `<div class="prop-tab-pane${tab.key === propertiesActiveTab ? "" : " hidden"}" data-prop-pane="${tab.key}">${panes[tab.key]}</div>`,
      ).join("");
    propertiesPanel.classList.remove("hidden");
    return;
  }

  const metricsList = await Promise.all(items.map(resolveMetrics));
  const numericTotals = METRIC_COLUMN_DEFS.map((col) => ({
    label: col.label,
    total: metricsList.reduce(
      (sum, m) => sum + (m ? Number(m[col.key as keyof ItemMetrics]) || 0 : 0),
      0,
    ),
  })).filter((entry) => entry.total !== 0);

  propertiesContent.innerHTML = `
    <div class="prop-section">
      <div class="prop-section-title">${escapeHtml(label ?? `${items.length} valda objekt`)}</div>
      ${renderPropTable([
        renderRow("Antal", items.length),
        ...numericTotals.map((entry) => renderRow(entry.label, entry.total.toFixed(2))),
      ])}
    </div>
    ${renderBoverketSelector(items)}
    ${renderGeneralSelector(items)}
  `;
  propertiesPanel.classList.remove("hidden");
}

async function selectItem(modelId: string, localId: number) {
  await applySelection([{ modelId, localId }]);
}

// Webbläsaren skickar ett "click" på canvasen efter varje musnedtryck+uppsläpp
// på samma element - även om musen flyttades långt mellan dem, t.ex. när man
// roterar/panorerar kameran (vänsterklick-drag) eller drar i ett befintligt
// snittplans gizmo. Utan detta skulle varje sådan drag felaktigt tolkas som
// ett klick och starta en ny markering/mätning/snitt. `mouseCanvasDownPos`
// sätts på mousedown och jämförs med klickpositionen.
let mouseCanvasDownPos: { x: number; y: number } | null = null;
const CLICK_DRAG_THRESHOLD_PX = 5;

canvas.addEventListener("mousedown", (event) => {
  mouseCanvasDownPos = { x: event.clientX, y: event.clientY };

  // Vänsterklick styr kamerarotation (mouseButtons.left = ROTATE) - flytta
  // rotationspunkten till det som ligger under muspekaren precis innan en
  // eventuell drag börjar, så rotationen sker kring det man faktiskt
  // klickade på istället för runt en fast punkt i modellens mitt.
  // setOrbitPoint flyttar bara pivoten, inte kameran, så det syns ingen
  // hopp - och om klicket missar geometrin behålls föregående pivot.
  if (event.button === 0) {
    void (async () => {
      const result = (await caster.castRay()) as unknown as { point?: THREE.Vector3 } | null;
      if (result?.point) {
        world.camera.controls.setOrbitPoint(result.point.x, result.point.y, result.point.z);
      }
    })();
  }
});

function wasCanvasDrag(event: MouseEvent): boolean {
  if (!mouseCanvasDownPos) return false;
  const dx = event.clientX - mouseCanvasDownPos.x;
  const dy = event.clientY - mouseCanvasDownPos.y;
  return dx * dx + dy * dy > CLICK_DRAG_THRESHOLD_PX * CLICK_DRAG_THRESHOLD_PX;
}

canvas.addEventListener("click", async (event) => {
  if (wasCanvasDrag(event)) return;

  // I snitt-/mätläge ska klick styra placeringen av snittet/måttet (se
  // klick-hanteraren och klippförhandsgranskningen nedan) - inte
  // highlighta/markera objekt, vilket annars stör själva mätningen/snittet.
  if (clipper.enabled || measurer.enabled) return;

  const result = await caster.castRay();

  if (!result || !("localId" in result)) {
    await clearSelection();
    return;
  }

  const hit = result as unknown as { localId: number; fragments: { modelId: string } };
  await selectItem(hit.fragments.modelId, hit.localId);
});

propertiesClose.addEventListener("click", () => {
  void clearSelection();
});

propertiesContent.addEventListener("change", (event) => {
  const target = event.target as HTMLElement;
  if (target.classList.contains("prop-boverket-select")) {
    applyBoverketCategoryToItems(highlightedItems, (target as HTMLSelectElement).value);
    return;
  }
  if (target.classList.contains("prop-general-select")) {
    applyGeneralCategoryToItems(highlightedItems, (target as HTMLSelectElement).value);
  }
});

// Bytet av aktiv flik kräver ingen omritning av innehållet - allt är redan
// renderat (se applySelection), bara .hidden som togglas om.
propertiesContent.addEventListener("click", (event) => {
  const target = event.target as HTMLElement;

  // Klick på en rad i Beräknat-fliken (bara vid enskild markering, se
  // applySelection) - togglar motsvarande måttlinje/yta/box i 3D-vyn.
  const metricRow = target.closest(".prop-metric-row") as HTMLElement | null;
  if (metricRow) {
    const key = metricRow.dataset.metric as MetricVisualizationKey | undefined;
    if (!key || !currentMetrics?.visualization) return;
    if (activeMetricVisualizationKey === key) {
      activeMetricVisualizationKey = null;
      hideMetricVisualization();
    } else {
      activeMetricVisualizationKey = key;
      showMetricVisualization(currentMetrics, key);
    }
    for (const row of propertiesContent.querySelectorAll<HTMLElement>(".prop-metric-row")) {
      row.classList.toggle("active", row.dataset.metric === activeMetricVisualizationKey);
    }
    return;
  }

  const tabButton = target.closest(".prop-tab") as HTMLButtonElement | null;
  if (!tabButton) return;

  const tab = tabButton.dataset.propTab as PropertiesTab | undefined;
  if (!tab || tab === propertiesActiveTab) return;

  propertiesActiveTab = tab;
  for (const button of propertiesContent.querySelectorAll<HTMLElement>(".prop-tab")) {
    button.classList.toggle("active", button.dataset.propTab === tab);
  }
  for (const pane of propertiesContent.querySelectorAll<HTMLElement>(".prop-tab-pane")) {
    pane.classList.toggle("hidden", pane.dataset.propPane !== tab);
  }
});

// Modellträd
type TreeMode = "spatial" | "type";
let treeMode: TreeMode = "spatial";
let treeBuiltMode: TreeMode | null = null;
let treeSelection: SelectableItem[] = [];
let treeRangeAnchor: HTMLElement | null = null;

function updateTreeRowSelectionStyles() {
  const keys = new Set(treeSelection.map(itemKey));

  for (const row of treeContent.querySelectorAll<HTMLElement>(".tree-row[data-local-id]")) {
    const key = `${row.dataset.modelId}::${row.dataset.localId}`;
    row.classList.toggle("selected", keys.has(key));
  }

  for (const node of treeContent.querySelectorAll<HTMLElement>(".tree-node")) {
    const groupRow = node.querySelector(":scope > .tree-row:not([data-local-id])");
    if (!groupRow) continue;
    const descendantRows = [...node.querySelectorAll<HTMLElement>(".tree-row[data-local-id]")];
    const allSelected =
      descendantRows.length > 0 &&
      descendantRows.every((r) => keys.has(`${r.dataset.modelId}::${r.dataset.localId}`));
    groupRow.classList.toggle("selected", allSelected);
  }
}

function collectSpatialIds(node: FRAGS.SpatialTreeItem, ids: number[]): void {
  if (node.localId !== null) ids.push(node.localId);
  for (const child of node.children ?? []) collectSpatialIds(child, ids);
}

/**
 * getSpatialStructure() alternerar mellan två sorters noder:
 * - Kategori-noder (localId === null, category = t.ex. "IFCWALL") som grupperar
 *   syskon av samma typ.
 * - Instansnoder (localId satt, category === null) som är den faktiska
 *   entiteten - dess "typ" känns bara till via förälderns category-fält.
 * Därför skickas förälderns kategori ner till instansnoder som parentCategory.
 */
function renderTreeNode(
  node: FRAGS.SpatialTreeItem,
  modelId: string,
  namesById: Map<number, string>,
  depth: number,
  parentCategory?: string,
): string {
  const children = node.children ?? [];
  const hasChildren = children.length > 0;
  const toggle = hasChildren
    ? `<button class="tree-toggle">${depth < 4 ? "▼" : "▶"}</button>`
    : `<span class="tree-toggle-spacer"></span>`;

  if (node.localId === null) {
    const label = `${node.category ?? "-"} (${children.length})`;
    const childrenHtml = hasChildren
      ? `<ul class="tree-children${depth < 4 ? "" : " collapsed"}">${children
          .map((child) =>
            renderTreeNode(child, modelId, namesById, depth + 1, node.category ?? undefined),
          )
          .join("")}</ul>`
      : "";
    return `
      <li class="tree-node">
        <div class="tree-row">
          ${toggle}
          <span class="tree-label tree-label-group" title="${escapeHtml(label)}">${escapeHtml(label)}</span>
        </div>
        ${childrenHtml}
      </li>
    `;
  }

  const name = namesById.get(node.localId);
  const label = name ?? parentCategory ?? "-";
  const sub = name ? (parentCategory ?? "") : "";
  const rowAttrs = `data-model-id="${escapeHtml(modelId)}" data-local-id="${node.localId}"`;
  const childrenHtml = hasChildren
    ? `<ul class="tree-children${depth < 4 ? "" : " collapsed"}">${children
        .map((child) => renderTreeNode(child, modelId, namesById, depth + 1))
        .join("")}</ul>`
    : "";

  return `
    <li class="tree-node">
      <div class="tree-row" ${rowAttrs}>
        ${toggle}
        <span class="tree-label" title="${escapeHtml(label)}">${escapeHtml(label)}</span>
        ${sub ? `<span class="tree-sub">${escapeHtml(sub)}</span>` : ""}
      </div>
      ${childrenHtml}
    </li>
  `;
}

async function buildModelTree(modelId: string): Promise<string> {
  const model = fragments.list.get(modelId);
  if (!model) return "";

  const root = await model.getSpatialStructure();
  const ids: number[] = [];
  collectSpatialIds(root, ids);

  const namesById = new Map<number, string>();
  if (ids.length > 0) {
    const dataList = await model.getItemsData(ids, { attributesDefault: true });
    for (const data of dataList) {
      const localIdAttr = data._localId;
      const nameAttr = data.Name;
      if (isAttribute(localIdAttr) && isAttribute(nameAttr) && nameAttr.value) {
        namesById.set(Number(localIdAttr.value), String(nameAttr.value));
      }
    }
  }

  const modelLabel = modelNames.get(modelId) ?? modelId;
  return `
    <div class="tree-model-name">${escapeHtml(modelLabel)}</div>
    ${renderModelHeaderLine(modelId)}
    <ul class="tree-root">${renderTreeNode(root, modelId, namesById, 0)}</ul>
  `;
}

/** Liten metarad under modellnamnet i modellträdet: exportprogram + IFC-schema
 *  (ur STEP-headern, se parseIfcHeader) - t.ex. "Revit 2024 · IFC4". Fler
 *  detaljer (författare, organisation, tidsstämpel) syns som tooltip. */
function renderModelHeaderLine(modelId: string): string {
  const header = parseIfcHeader(modelIfcText.get(modelId) ?? "");
  const summary = formatIfcHeaderSummary(header);
  if (!summary) return "";
  return `<div class="tree-model-meta" title="${escapeHtml(summary.tooltip)}">${escapeHtml(summary.line)}</div>`;
}

/** Platt vy grupperad på IFC-kategori (t.ex. alla väggar i hela modellen), till
 *  skillnad från den rumsliga vyn där samma väggar ligger utspridda per våningsplan. */
async function buildTypeTree(modelId: string): Promise<string> {
  const model = fragments.list.get(modelId);
  if (!model) return "";

  const ids = await model.getItemsIdsWithGeometry();
  const modelLabel = modelNames.get(modelId) ?? modelId;

  if (ids.length === 0) {
    return `
      <div class="tree-model-name">${escapeHtml(modelLabel)}</div>
      ${renderModelHeaderLine(modelId)}
      <div class="model-empty">Inga objekt</div>
    `;
  }

  const dataList = await model.getItemsData(ids, { attributesDefault: true });
  const byCategory = new Map<string, { localId: number; name: string }[]>();

  for (const data of dataList) {
    const localIdAttr = data._localId;
    if (!isAttribute(localIdAttr)) continue;
    const categoryAttr = data._category;
    const nameAttr = data.Name;
    const category = isAttribute(categoryAttr) ? String(categoryAttr.value) : "Okänd";
    const localId = Number(localIdAttr.value);
    const name = isAttribute(nameAttr) && nameAttr.value ? String(nameAttr.value) : `#${localId}`;

    const list = byCategory.get(category) ?? [];
    list.push({ localId, name });
    byCategory.set(category, list);
  }

  const categoriesHtml = [...byCategory.keys()]
    .sort()
    .map((category) => {
      const items = byCategory.get(category)!.sort((a, b) => a.name.localeCompare(b.name));
      const itemsHtml = items
        .map(
          (item) => `
            <li class="tree-node">
              <div class="tree-row" data-model-id="${escapeHtml(modelId)}" data-local-id="${item.localId}">
                <span class="tree-toggle-spacer"></span>
                <span class="tree-label" title="${escapeHtml(item.name)}">${escapeHtml(item.name)}</span>
              </div>
            </li>
          `,
        )
        .join("");
      const label = `${category} (${items.length})`;

      return `
        <li class="tree-node">
          <div class="tree-row">
            <button class="tree-toggle">▶</button>
            <span class="tree-label tree-label-group" title="${escapeHtml(label)}">${escapeHtml(label)}</span>
          </div>
          <ul class="tree-children collapsed">${itemsHtml}</ul>
        </li>
      `;
    })
    .join("");

  return `
    <div class="tree-model-name">${escapeHtml(modelLabel)}</div>
    ${renderModelHeaderLine(modelId)}
    <ul class="tree-root">${categoriesHtml}</ul>
  `;
}

async function renderModelTree() {
  const modelIds = [...fragments.list.keys()];
  if (modelIds.length === 0) {
    treeContent.innerHTML = '<div class="model-empty">Inga modeller inlästa</div>';
    treeBuiltMode = null;
    return;
  }

  treeContent.innerHTML = '<div class="model-empty">Bygger modellträd...</div>';
  const builder = treeMode === "spatial" ? buildModelTree : buildTypeTree;
  const sections = await Promise.all(modelIds.map((modelId) => builder(modelId)));
  treeContent.innerHTML = sections.join("");
  treeBuiltMode = treeMode;
  treeRangeAnchor = null;
  updateTreeRowSelectionStyles();
}

/** Bygger om trädet om panelen redan är öppen (den är det som standard sedan
 *  Fas 1) - används efter att en fils inläsning/borttagning FAKTISKT är
 *  klar (inte i onItemSet, som triggas för tidigt - se kommentaren där). */
function renderModelTreeIfOpen(): void {
  if (!treePanel.classList.contains("hidden")) void renderModelTree();
}

treeToggle.addEventListener("click", async () => {
  const isHidden = treePanel.classList.contains("hidden");

  if (!isHidden) {
    treePanel.classList.add("hidden");
    treeToggle.classList.remove("active");
    return;
  }

  treePanel.classList.remove("hidden");
  treeToggle.classList.add("active");
  if (treeBuiltMode !== treeMode) await renderModelTree();
});

treeClose.addEventListener("click", () => {
  treePanel.classList.add("hidden");
  treeToggle.classList.remove("active");
});

// Modellträd är öppet som standard (matchar Solibris alltid synliga Model
// Tree-panel) - övriga toggle-paneler (Dubbletter/Mängdavtagning/Skapa
// modell) förblir stängda tills man själv öppnar dem, eftersom de är mer
// tillfälliga verktyg snarare än en ständigt relevant strukturvy.
treePanel.classList.remove("hidden");
treeToggle.classList.add("active");
void renderModelTree();

treeRefresh.addEventListener("click", () => {
  void renderModelTree();
});

treeTabs.addEventListener("click", async (event) => {
  const target = event.target as HTMLElement;
  const button = target.closest(".tree-tab") as HTMLButtonElement | null;
  const mode = button?.dataset.treeMode as TreeMode | undefined;
  if (!button || !mode || mode === treeMode) return;

  treeMode = mode;
  for (const tab of treeTabs.querySelectorAll(".tree-tab")) {
    tab.classList.toggle("active", tab === button);
  }
  await renderModelTree();
});

treeContent.addEventListener("click", (event) => {
  const target = event.target as HTMLElement;

  if (target.classList.contains("tree-toggle")) {
    const li = target.closest(".tree-node");
    const childUl = li?.querySelector(":scope > .tree-children");
    if (childUl) {
      const collapsed = childUl.classList.toggle("collapsed");
      target.textContent = collapsed ? "▶" : "▼";
    }
    return;
  }

  const row = target.closest(".tree-row") as HTMLElement | null;
  if (!row) return;

  let clickedItems: SelectableItem[];
  let groupLabel: string | undefined;

  if (row.dataset.modelId) {
    const localId = Number(row.dataset.localId);
    if (Number.isNaN(localId)) return;
    clickedItems = [{ modelId: row.dataset.modelId, localId }];
  } else {
    // Kategori-grupprad (t.ex. "IFCWALLSTANDARDCASE (17)") - klick markerar
    // alla objekt som ligger under den, precis som gruppraderna i mängdavtagningen.
    const node = row.closest(".tree-node");
    const descendantRows = node
      ? [...node.querySelectorAll<HTMLElement>(".tree-row[data-local-id]")]
      : [];
    if (descendantRows.length === 0) return;
    clickedItems = descendantRows.map((r) => ({
      modelId: r.dataset.modelId as string,
      localId: Number(r.dataset.localId),
    }));
    groupLabel = row.querySelector(".tree-label")?.textContent ?? undefined;
  }

  let items: SelectableItem[];
  let label: string | undefined;

  if (event.shiftKey && treeRangeAnchor && row.dataset.modelId) {
    const rows = [...treeContent.querySelectorAll<HTMLElement>(".tree-row[data-local-id]")].filter(
      (r) => r.offsetParent !== null,
    );
    const anchorIndex = rows.indexOf(treeRangeAnchor);
    const currentIndex = rows.indexOf(row);

    if (anchorIndex === -1 || currentIndex === -1) {
      items = resolveMultiSelect(
        treeSelection,
        clickedItems,
        event.ctrlKey || event.metaKey ? "toggle" : "replace",
      );
    } else {
      const [start, end] =
        anchorIndex < currentIndex ? [anchorIndex, currentIndex] : [currentIndex, anchorIndex];
      items = rows.slice(start, end + 1).map((r) => ({
        modelId: r.dataset.modelId as string,
        localId: Number(r.dataset.localId),
      }));
    }
  } else {
    const mode = event.ctrlKey || event.metaKey ? "toggle" : "replace";
    items = resolveMultiSelect(treeSelection, clickedItems, mode);
    if (row.dataset.modelId) treeRangeAnchor = row;
    if (mode === "replace") label = groupLabel;
  }

  treeSelection = items;
  updateTreeRowSelectionStyles();
  void applySelection(items, label);
});

// Dubbletter - hittar objekt som delar kategori, position och storlek (inom
// en toleransgräns), eller som bara delvis upptar samma volym och position
// (t.ex. en kopia som flyttats/ändrats något), i samma modell eller mellan
// olika modeller. Vanligaste orsaken är att samma element råkat komma med i
// flera källfiler som slagits samman.
interface DuplicateCandidate {
  modelId: string;
  localId: number;
  name: string;
  box: THREE.Box3;
}

interface DuplicateGroup {
  category: string;
  items: DuplicateCandidate[];
  /** "exact" = alla objekt delar (avrundat) exakt position+storlek.
   *  "overlap" = minst ett par i gruppen delar bara en del av volymen. */
  kind: "exact" | "overlap";
  /** Index in i XRAY_DUPLICATE_GROUP_COLORS - satt en gång när gruppen skapas
   *  (inte vid sorterad visning) så färgen är stabil oavsett vilken
   *  sorteringsordning listan råkar visas i. */
  colorIndex: number;
}

let duplicateGroups: DuplicateGroup[] = [];
let duplicatesBuilt = false;
let duplicatesSelection: SelectableItem[] = [];
let duplicatesRangeAnchor: HTMLElement | null = null;
let duplicatesXrayActive = false;
/** Cache av collectDuplicateCandidates - låter toleransreglaget gruppera om
 *  billigt utan att fråga modell-API:et igen. Nollställs (se duplicatesBuilt)
 *  när modeller läggs till/tas bort. */
let duplicateCandidatesByCategory: Map<string, DuplicateCandidate[]> | null = null;

/** Avrundar till mm-precision, så att objekt som skiljer sig med bråkdelar
 *  av en millimeter (typiskt för oberoende exporter av "samma" geometri)
 *  ändå räknas som identiska. */
function roundToMm(value: number): number {
  return Math.round(value * 1000) / 1000;
}

function boxKey(box: THREE.Box3): string {
  const size = box.getSize(new THREE.Vector3());
  const center = box.getCenter(new THREE.Vector3());
  return [
    roundToMm(center.x),
    roundToMm(center.y),
    roundToMm(center.z),
    roundToMm(size.x),
    roundToMm(size.y),
    roundToMm(size.z),
  ].join("|");
}

/** Intersection-over-Union (Jaccard-index) för två boundingboxar: överlappets
 *  volym delat på UNIONENS volym. Att bara kräva att en stor andel av det
 *  MINDRE objektets volym delas (t.ex. "60% av minsta boxen") slår fortfarande
 *  ut brett på riktiga modeller - en liten kopplingsvinkel som sitter helt
 *  inuti en stor balks boundingbox skulle räknas som "100% överlapp" trots
 *  att de inte alls är dubbletter, bara olika delar av samma knutpunkt.
 *  IoU straffar den typen av storleksskillnad (unionen domineras av den
 *  stora balken, så kvoten blir låg) och blir hög bara när båda objekten
 *  verkligen är ungefär lika stora OCH i det närmaste sammanfallande. */
function boxOverlapFraction(a: THREE.Box3, b: THREE.Box3): number {
  const ix = Math.min(a.max.x, b.max.x) - Math.max(a.min.x, b.min.x);
  if (ix <= 0) return 0;
  const iy = Math.min(a.max.y, b.max.y) - Math.max(a.min.y, b.min.y);
  if (iy <= 0) return 0;
  const iz = Math.min(a.max.z, b.max.z) - Math.max(a.min.z, b.min.z);
  if (iz <= 0) return 0;

  const overlapVolume = ix * iy * iz;
  const sizeA = a.getSize(new THREE.Vector3());
  const sizeB = b.getSize(new THREE.Vector3());
  const volumeA = sizeA.x * sizeA.y * sizeA.z;
  const volumeB = sizeB.x * sizeB.y * sizeB.z;
  const unionVolume = volumeA + volumeB - overlapVolume;
  if (unionVolume <= 0) return 0;

  return overlapVolume / unionVolume;
}

/** Andel IoU (se boxOverlapFraction) som krävs för att två objekt ska räknas
 *  som en möjlig dubblett/överlappning. Justerbar via reglaget i
 *  dubblettpanelen - satt högt (0.8) som standard eftersom boundingboxar är
 *  axelriktade: två OLIKA diagonala snedstag/stag kan råka få nästan
 *  identiska (stora) boundingboxar även om de knappt delar någon verklig
 *  volym, eftersom en axelriktad box runt ett diagonalt element blir mycket
 *  större än elementet självt. Riktiga dubbletter (samma objekt kopierat)
 *  ligger typiskt runt 95-100% IoU, så 80% ger marginal utan att öppna för
 *  den bruskällan. Tunna, platta element (t.ex. bottenplattor) kan dock behöva
 *  en lägre tolerans: en liten avvikelse i höjdled tar proportionellt mycket
 *  större andel av deras redan tunna boundingbox än för ett "normalt" objekt,
 *  vilket kan trycka ner IoU under 80% trots att plattorna uppenbart är
 *  samma platta. */
let duplicatesOverlapThreshold = 0.8;

function boxesOverlapVolume(a: THREE.Box3, b: THREE.Box3): boolean {
  return boxOverlapFraction(a, b) >= duplicatesOverlapThreshold;
}

/** Den dyra delen av dubblettsökningen - hämtar boundingboxar och
 *  grunddata för varje objekt i alla inlästa modeller. Resultatet cachas
 *  (se duplicateCandidatesByCategory) så att groupDuplicateCandidates kan
 *  köras om billigt när toleransen ändras, utan att fråga modell-API:et igen. */
async function collectDuplicateCandidates(): Promise<Map<string, DuplicateCandidate[]>> {
  const byCategory = new Map<string, DuplicateCandidate[]>();

  for (const modelId of fragments.list.keys()) {
    const model = fragments.list.get(modelId);
    if (!model) continue;

    const ids = await model.getItemsIdsWithGeometry();
    if (ids.length === 0) continue;

    const [boxes, dataList] = await Promise.all([
      model.getBoxes(ids),
      model.getItemsData(ids, { attributesDefault: true }),
    ]);

    for (let i = 0; i < ids.length; i++) {
      const box = boxes[i];
      if (!box || box.isEmpty()) continue;

      const worldBox = box.clone().applyMatrix4(model.object.matrixWorld);

      const data = dataList[i];
      const categoryAttr = data._category;
      const nameAttr = data.Name;
      const category = isAttribute(categoryAttr) ? String(categoryAttr.value) : "Okänd";
      const name =
        isAttribute(nameAttr) && nameAttr.value ? String(nameAttr.value) : `#${ids[i]}`;

      const list = byCategory.get(category) ?? [];
      list.push({ modelId, localId: ids[i], name, box: worldBox });
      byCategory.set(category, list);
    }
  }

  return byCategory;
}

/** Grupperar redan insamlade kandidater (se collectDuplicateCandidates) med
 *  den aktuella toleransen - ren/synkron gruppering, billig nog att köra om
 *  varje gång toleransreglaget flyttas. */
function groupDuplicateCandidates(byCategory: Map<string, DuplicateCandidate[]>): DuplicateGroup[] {
  const groups: DuplicateGroup[] = [];

  for (const [category, items] of byCategory) {
    // Union-Find: två objekt vars volymer överlappar hamnar i samma grupp,
    // även transitivt (A överlappar B, B överlappar C -> alla tre i en
    // grupp) - fångar både exakta dubbletter och kedjor av delvis
    // överlappande kopior utan att jämförelsen blir kvadratisk över hela
    // modellen (bara inom samma kategori).
    const parent = items.map((_, i) => i);
    function find(i: number): number {
      while (parent[i] !== i) {
        parent[i] = parent[parent[i]];
        i = parent[i];
      }
      return i;
    }
    function union(a: number, b: number) {
      const ra = find(a);
      const rb = find(b);
      if (ra !== rb) parent[ra] = rb;
    }

    for (let i = 0; i < items.length; i++) {
      for (let j = i + 1; j < items.length; j++) {
        if (boxesOverlapVolume(items[i].box, items[j].box)) union(i, j);
      }
    }

    const components = new Map<number, DuplicateCandidate[]>();
    for (let i = 0; i < items.length; i++) {
      const root = find(i);
      const list = components.get(root) ?? [];
      list.push(items[i]);
      components.set(root, list);
    }

    for (const componentItems of components.values()) {
      if (componentItems.length < 2) continue;
      const key = boxKey(componentItems[0].box);
      const exact = componentItems.every((item) => boxKey(item.box) === key);
      groups.push({
        category,
        items: componentItems,
        kind: exact ? "exact" : "overlap",
        colorIndex: groups.length % XRAY_DUPLICATE_GROUP_COLORS.length,
      });
    }
  }

  return groups;
}

type DuplicatesSortMode = "default" | "volume-desc" | "volume-asc";
let duplicatesSortMode: DuplicatesSortMode = "default";

/** Representativ volym för en dubblettgrupp - det största objektet i gruppen,
 *  eftersom "overlap"-grupper kan innehålla objekt av olika storlek och det
 *  är då den stora volymen som är intressant att prioritera. */
function groupVolume(group: DuplicateGroup): number {
  let max = 0;
  for (const item of group.items) {
    const size = item.box.getSize(new THREE.Vector3());
    const volume = size.x * size.y * size.z;
    if (volume > max) max = volume;
  }
  return max;
}

function sortedDuplicateGroups(): DuplicateGroup[] {
  const groups = [...duplicateGroups];
  if (duplicatesSortMode === "volume-desc") {
    groups.sort((a, b) => groupVolume(b) - groupVolume(a));
  } else if (duplicatesSortMode === "volume-asc") {
    groups.sort((a, b) => groupVolume(a) - groupVolume(b));
  } else {
    groups.sort((a, b) => {
      if (a.kind !== b.kind) return a.kind === "exact" ? -1 : 1;
      return b.items.length - a.items.length;
    });
  }
  return groups;
}

function renderDuplicatesPanel() {
  if (duplicateGroups.length === 0) {
    duplicatesContent.innerHTML = '<div class="model-empty">Inga dubbletter hittades</div>';
    return;
  }

  duplicatesContent.innerHTML = sortedDuplicateGroups()
    .map((group, groupIndex) => {
      const itemsHtml = group.items
        .map(
          (item) => `
            <li class="tree-row" data-model-id="${escapeHtml(item.modelId)}" data-local-id="${item.localId}">
              <span class="tree-label" title="${escapeHtml(item.name)}">${escapeHtml(item.name)}</span>
              <span class="tree-sub">${escapeHtml(modelNames.get(item.modelId) ?? item.modelId)}</span>
            </li>
          `,
        )
        .join("");
      const kindLabel =
        group.kind === "exact"
          ? '<span class="duplicate-kind duplicate-kind-exact">Exakt</span>'
          : '<span class="duplicate-kind duplicate-kind-overlap">Delvis överlapp</span>';
      const volumeLabel = `${groupVolume(group).toFixed(2)} m³`;
      const colorSwatch = `<span class="duplicate-color-swatch" style="background:${duplicateGroupColor(group.colorIndex)}" title="Färg i röntgenvyn"></span>`;
      return `
        <div class="duplicate-group">
          <div class="duplicate-group-header" data-group-index="${groupIndex}">
            ${colorSwatch}
            ${kindLabel}
            <span>${escapeHtml(group.category)} (${group.items.length} st, ${volumeLabel})</span>
          </div>
          <ul class="duplicate-items">${itemsHtml}</ul>
        </div>
      `;
    })
    .join("");

  updateDuplicatesSelectionStyles();
}

function updateDuplicatesSelectionStyles() {
  const keys = new Set(duplicatesSelection.map(itemKey));

  for (const row of duplicatesContent.querySelectorAll<HTMLElement>(".tree-row[data-local-id]")) {
    const key = `${row.dataset.modelId}::${row.dataset.localId}`;
    row.classList.toggle("selected", keys.has(key));
  }

  for (const group of duplicatesContent.querySelectorAll<HTMLElement>(".duplicate-group")) {
    const rows = [...group.querySelectorAll<HTMLElement>(".tree-row[data-local-id]")];
    const allSelected =
      rows.length > 0 &&
      rows.every((r) => keys.has(`${r.dataset.modelId}::${r.dataset.localId}`));
    group.querySelector(".duplicate-group-header")?.classList.toggle("selected", allSelected);
  }
}

function updateDuplicatesStatusText() {
  const totalItems = duplicateGroups.reduce((sum, g) => sum + g.items.length, 0);
  const exactCount = duplicateGroups.filter((g) => g.kind === "exact").length;
  const overlapCount = duplicateGroups.length - exactCount;
  duplicatesStatus.textContent =
    duplicateGroups.length === 0
      ? "Inga dubbletter"
      : `${exactCount} exakt(a), ${overlapCount} delvis överlapp - ${totalItems} objekt totalt`;
}

async function runDuplicateScan() {
  duplicatesStatus.textContent = "Skannar...";
  duplicateCandidatesByCategory = await collectDuplicateCandidates();
  duplicateGroups = groupDuplicateCandidates(duplicateCandidatesByCategory);
  duplicatesBuilt = true;
  duplicatesSelection = [];
  duplicatesRangeAnchor = null;
  renderDuplicatesPanel();
  updateDuplicatesStatusText();

  // Om röntgenvyn redan var på när man skannade om (t.ex. efter att ha gömt
  // en dubblett), rita om den med de uppdaterade grupperna istället för att
  // lämna kvar gamla, nu felaktiga uppljusningar.
  if (duplicatesXrayActive) await applyDuplicatesXray();
}

/** Grupperar om de redan insamlade kandidaterna med en ny tolerans - körs när
 *  toleransreglaget släpps, utan att fråga modell-API:et igen (se
 *  collectDuplicateCandidates/duplicateCandidatesByCategory). Om ingen
 *  skanning gjorts än finns inget att gruppera om - toleransen används då
 *  automatiskt av nästa fullständiga skanning. */
async function regroupDuplicates() {
  if (!duplicateCandidatesByCategory) return;
  duplicateGroups = groupDuplicateCandidates(duplicateCandidatesByCategory);
  duplicatesSelection = [];
  duplicatesRangeAnchor = null;
  renderDuplicatesPanel();
  updateDuplicatesStatusText();
  if (duplicatesXrayActive) await applyDuplicatesXray();
}

/** Gör hela modellen halvgenomskinlig och lyser upp alla hittade dubbletter
 *  ovanpå - varje GRUPP i en egen färg (se XRAY_DUPLICATE_GROUP_COLORS) så
 *  det går att se vilka objekt som hör ihop, inte bara att de är dubbletter.
 *  Grupper som delar färg (fler grupper än paletten) samlas i samma
 *  highlight-anrop per modell, så det blir max en handfull anrop även med
 *  hundratals grupper. */
async function applyDuplicatesXray() {
  for (const modelId of fragments.list.keys()) {
    const model = fragments.list.get(modelId);
    if (!model) continue;
    const ids = await model.getItemsIdsWithGeometry();
    if (ids.length > 0) await model.highlight(ids, XRAY_GHOST_MATERIAL);
  }

  const byModelAndColor = new Map<string, Map<number, number[]>>();
  for (const group of duplicateGroups) {
    for (const item of group.items) {
      const byColor = byModelAndColor.get(item.modelId) ?? new Map<number, number[]>();
      const list = byColor.get(group.colorIndex) ?? [];
      list.push(item.localId);
      byColor.set(group.colorIndex, list);
      byModelAndColor.set(item.modelId, byColor);
    }
  }
  for (const [modelId, byColor] of byModelAndColor) {
    const model = fragments.list.get(modelId);
    if (!model) continue;
    for (const [colorIndex, localIds] of byColor) {
      await model.highlight(localIds, duplicateGroupMaterial(colorIndex));
    }
  }

  duplicatesXrayActive = true;
  duplicatesXray.classList.add("active");
  await fragments.core.update(true);
  postproductionRenderer.needsUpdate = true;
}

async function clearDuplicatesXray() {
  if (!duplicatesXrayActive) return;

  for (const modelId of fragments.list.keys()) {
    await fragments.list.get(modelId)?.resetHighlight();
  }
  duplicatesXrayActive = false;
  duplicatesXray.classList.remove("active");

  // En pågående vanlig markering (t.ex. en dubblettgrupp man klickat på)
  // nollställdes också av resetHighlight ovan - rita tillbaka den istället
  // för att den tyst försvinner.
  if (highlightedItems.length > 0) {
    const byModel = new Map<string, number[]>();
    for (const { modelId, localId } of highlightedItems) {
      const list = byModel.get(modelId) ?? [];
      list.push(localId);
      byModel.set(modelId, list);
    }
    for (const [modelId, localIds] of byModel) {
      await fragments.list.get(modelId)?.highlight(localIds, HIGHLIGHT_MATERIAL);
    }
  }

  await fragments.core.update(true);
  postproductionRenderer.needsUpdate = true;
}

duplicatesToggle.addEventListener("click", async () => {
  const isHidden = duplicatesPanel.classList.contains("hidden");

  if (!isHidden) {
    duplicatesPanel.classList.add("hidden");
    duplicatesToggle.classList.remove("active");
    void clearDuplicatesXray();
    return;
  }

  duplicatesPanel.classList.remove("hidden");
  duplicatesToggle.classList.add("active");
  if (!duplicatesBuilt) await runDuplicateScan();
});

duplicatesClose.addEventListener("click", () => {
  duplicatesPanel.classList.add("hidden");
  duplicatesToggle.classList.remove("active");
  void clearDuplicatesXray();
});

duplicatesRefresh.addEventListener("click", () => {
  void runDuplicateScan();
});

duplicatesXray.addEventListener("click", () => {
  if (duplicatesXrayActive) {
    void clearDuplicatesXray();
    return;
  }
  // Ömsesidigt uteslutande med Synlighet-sektionens röntgenvy - se motsvarande
  // kommentar vid visibilityXray.
  void (async () => {
    if (visibilityXrayActive) await clearVisibilityXray();
    await applyDuplicatesXray();
  })();
});

duplicatesSortSelect.addEventListener("change", () => {
  duplicatesSortMode = duplicatesSortSelect.value as DuplicatesSortMode;
  renderDuplicatesPanel();
});

// "input" ger en live-uppdaterad procentetikett medan man drar i reglaget,
// men själva omgrupperingen (och ev. omritning av röntgenvyn) väntar till
// "change" (släppt reglage) så att den inte körs om för varje enskilt steg.
duplicatesToleranceInput.addEventListener("input", () => {
  duplicatesToleranceValue.textContent = `${duplicatesToleranceInput.value}%`;
});

duplicatesToleranceInput.addEventListener("change", () => {
  duplicatesOverlapThreshold = Number(duplicatesToleranceInput.value) / 100;
  void regroupDuplicates();
});

duplicatesContent.addEventListener("click", (event) => {
  const target = event.target as HTMLElement;
  const row = target.closest(".tree-row") as HTMLElement | null;
  const groupHeader = target.closest(".duplicate-group-header") as HTMLElement | null;

  let clickedItems: SelectableItem[];
  let groupLabel: string | undefined;

  if (row?.dataset.modelId) {
    const localId = Number(row.dataset.localId);
    if (Number.isNaN(localId)) return;
    clickedItems = [{ modelId: row.dataset.modelId, localId }];
  } else if (groupHeader) {
    const group = groupHeader.closest(".duplicate-group");
    const rows = group
      ? [...group.querySelectorAll<HTMLElement>(".tree-row[data-local-id]")]
      : [];
    if (rows.length === 0) return;
    clickedItems = rows.map((r) => ({
      modelId: r.dataset.modelId as string,
      localId: Number(r.dataset.localId),
    }));
    groupLabel = groupHeader.textContent?.replace(/\s+/g, " ").trim();
  } else {
    return;
  }

  let items: SelectableItem[];
  let label: string | undefined;

  if (event.shiftKey && duplicatesRangeAnchor && row?.dataset.modelId) {
    const rows = [
      ...duplicatesContent.querySelectorAll<HTMLElement>(".tree-row[data-local-id]"),
    ].filter((r) => r.offsetParent !== null);
    const anchorIndex = rows.indexOf(duplicatesRangeAnchor);
    const currentIndex = rows.indexOf(row);

    if (anchorIndex === -1 || currentIndex === -1) {
      items = resolveMultiSelect(
        duplicatesSelection,
        clickedItems,
        event.ctrlKey || event.metaKey ? "toggle" : "replace",
      );
    } else {
      const [start, end] =
        anchorIndex < currentIndex ? [anchorIndex, currentIndex] : [currentIndex, anchorIndex];
      items = rows.slice(start, end + 1).map((r) => ({
        modelId: r.dataset.modelId as string,
        localId: Number(r.dataset.localId),
      }));
    }
  } else {
    const mode = event.ctrlKey || event.metaKey ? "toggle" : "replace";
    items = resolveMultiSelect(duplicatesSelection, clickedItems, mode);
    if (row?.dataset.modelId) duplicatesRangeAnchor = row;
    if (mode === "replace") label = groupLabel;
  }

  duplicatesSelection = items;
  updateDuplicatesSelectionStyles();
  void applySelection(items, label);
});

let activeContextMenu: HTMLDivElement | null = null;

function closeContextMenu() {
  activeContextMenu?.remove();
  activeContextMenu = null;
}

/** Enkel flytande högerklicksmeny vid muspekaren - stängs vid klick utanför,
 *  Escape, scroll eller att fönstret tappar fokus. */
function showContextMenu(x: number, y: number, entries: { label: string; onClick: () => void }[]) {
  closeContextMenu();
  const menu = document.createElement("div");
  menu.className = "context-menu";
  for (const entry of entries) {
    const button = document.createElement("button");
    button.className = "context-menu-item";
    button.textContent = entry.label;
    button.addEventListener("click", () => {
      closeContextMenu();
      entry.onClick();
    });
    menu.appendChild(button);
  }
  document.body.appendChild(menu);
  activeContextMenu = menu;

  // Håll menyn inom synligt fönster istället för att låta den klippas/hamna
  // utanför skärmen när högerklicket sker nära en kant.
  const rect = menu.getBoundingClientRect();
  const clampedX = Math.min(x, window.innerWidth - rect.width - 4);
  const clampedY = Math.min(y, window.innerHeight - rect.height - 4);
  menu.style.left = `${Math.max(4, clampedX)}px`;
  menu.style.top = `${Math.max(4, clampedY)}px`;
}

document.addEventListener("click", (event) => {
  if (activeContextMenu && !activeContextMenu.contains(event.target as Node)) closeContextMenu();
});
document.addEventListener("keydown", (event) => {
  if (event.key === "Escape") closeContextMenu();
});
document.addEventListener("scroll", closeContextMenu, true);
window.addEventListener("blur", closeContextMenu);

duplicatesContent.addEventListener("contextmenu", (event) => {
  const target = event.target as HTMLElement;
  const row = target.closest(".tree-row") as HTMLElement | null;
  if (!row?.dataset.modelId) return;
  const localId = Number(row.dataset.localId);
  if (Number.isNaN(localId)) return;
  event.preventDefault();

  const item: SelectableItem = { modelId: row.dataset.modelId, localId };
  showContextMenu(event.clientX, event.clientY, [
    {
      label: "Gå till",
      onClick: () => {
        duplicatesSelection = [item];
        duplicatesRangeAnchor = row;
        updateDuplicatesSelectionStyles();
        void applySelection([item]).then(() => zoomToHighlighted());
      },
    },
  ]);
});

// Snitt (sektionering) och mätning - ömsesidigt uteslutande verktygsläge
const clipper = components.get(OBC.Clipper);
clipper.enabled = false;
// Standardstorleken (2 enheter) gör att snittplanets färgade platta täcker en
// stor del av vyn och skymmer modellen - autoScalePlanes håller kvar en
// konstant SKÄRMSTORLEK oavsett zoom, så en mindre bas-storlek räcker för att
// den bara ska synas som en kompakt markör vid själva snittytan.
clipper.size = 0.6;

const measurer = components.get(OBF.LengthMeasurement);
measurer.world = world;
measurer.color = new THREE.Color("#22c55e");
measurer.enabled = false;

type Tool = "select" | "clip" | "measure";

const TOOL_STATUS_LABELS: Record<Tool, string> = {
  select: "",
  clip: "Snittläge aktivt",
  measure: "Mätläge aktivt",
};

function setActiveTool(tool: Tool) {
  clipper.enabled = tool === "clip";
  measurer.enabled = tool === "measure";
  clipToggle.classList.toggle("active", tool === "clip");
  measureToggle.classList.toggle("active", tool === "measure");
  if (tool !== "clip") clipPreviewGroup.visible = false;
  if (tool !== "measure") measurePreviewGroup.visible = false;
  if (tool !== "measure") cancelPerpendicularMeasure();
  statusMode.textContent = TOOL_STATUS_LABELS[tool];
}

clipToggle.addEventListener("click", () => {
  setActiveTool(clipper.enabled ? "select" : "clip");
});

measureToggle.addEventListener("click", () => {
  setActiveTool(measurer.enabled ? "select" : "measure");
});

// Sätts av clipper.onAfterDrag precis innan webbläsarens syntetiska "click"
// avfyras för samma mus-upp (mousedown+mousemove+mouseup på canvas ger click
// oavsett rörelse) - annars skulle det sista musklicket i en drag av ett
// BEFINTLIGT snittplan (i gizmot) tolkas som ett nytt klick och skapa ett
// extra plan ovanpå det man just flyttade.
let clipperJustDragged = false;
let clipperDragging = false;
clipper.onBeforeDrag.add(() => { clipperDragging = true; });
clipper.onAfterDrag.add(() => {
  clipperJustDragged = true;
  clipperDragging = false;
  postproductionRenderer.needsUpdate = true;
});

// While dragging an existing clip plane, camera controls don't fire update
// events — set needsUpdate directly from pointer movement.
canvas.addEventListener("pointermove", () => {
  if (clipperDragging) postproductionRenderer.needsUpdate = true;
});

canvas.addEventListener("click", (event) => {
  if (clipper.enabled) {
    if (clipperJustDragged) {
      clipperJustDragged = false;
      return;
    }
    // Rotation/panorering av kameran (vänsterklick-drag) ger också ett click
    // på canvasen vid uppsläpp - annars skulle t.ex. varje kamerarotation i
    // snittläge felaktigt skapa ett nytt snittplan.
    if (wasCanvasDrag(event)) return;
    clipper.create(world);
    postproductionRenderer.needsUpdate = true;
  }
  if (measurer.enabled) {
    if (wasCanvasDrag(event)) return;
    // Vanligt punkt-till-punkt-mått som förut, plus ett vinkelrätt mått
    // (avstånd rakt in mot den första träffade ytan) som extra rad bredvid -
    // användaren väljer efteråt vilket av de två den bryr sig om, istället
    // för att behöva slå om läge i förväg.
    measurer.create();
    postproductionRenderer.needsUpdate = true;
    void handlePerpendicularMeasureClick();
  }
});

// Visualisering av Beräknat-flikens rader (Net Area/Volume/Tjocklek/Höjd/
// Bredd/Längd/Omkrets) i 3D-vyn - se renderMetricRow/showMetricVisualization.
// Byggs om från grunden vid varje klick istället för att hålla separata
// återanvändbara objekt per måtttyp, eftersom formen skiljer sig helt
// (linje/yta/box) beroende på vilken rad som klickats.
const METRIC_VIS_COLOR = new THREE.Color("#a855f7");
// Radien (som andel av visualScale) för alla tunna måttlinjer/konturer -
// kalibrerad genom visuell test: en tunnare linje (t.ex. 0.035) syns nästan
// inte bredvid modellens egna svarta kantlinjer, som redan är förhållandevis
// tjocka i den här vyn.
const OUTLINE_RADIUS_FACTOR = 0.15;

const metricsVisualizationGroup = new THREE.Group();
metricsVisualizationGroup.visible = false;
metricsVisualizationGroup.renderOrder = 999;
world.scene.three.add(metricsVisualizationGroup);

function clearMetricVisualizationGroup(): void {
  for (const child of [...metricsVisualizationGroup.children]) {
    metricsVisualizationGroup.remove(child);
    if (child instanceof THREE.Sprite) {
      child.material.map?.dispose();
      child.material.dispose();
      continue;
    }
    if (
      child instanceof THREE.Mesh ||
      child instanceof THREE.Line ||
      child instanceof THREE.LineSegments
    ) {
      child.geometry.dispose();
      if (Array.isArray(child.material)) {
        for (const material of child.material) material.dispose();
      } else {
        child.material.dispose();
      }
    }
  }
}

/** Referensstorlek (world units) för måttvisualiseringens linjer/markörer,
 *  baserad på kameraavståndet till objektet vid klicktillfället - samma
 *  clamp-mönster som redan används för mätverktygets hover-markör (se
 *  updateMeasurePreview). Etiketterna (addLabel) använder INTE detta - de
 *  ritas med sizeAttenuation:false, dvs. konstant skärmstorlek oavsett
 *  avstånd, eftersom just textläsbarhet var det uttryckliga problemet
 *  (behövde zooma väldigt nära för att läsa siffrorna). Linjerna/klotens
 *  world-space-storlek räknas bara ut EN gång per klick (inte varje frame),
 *  så den följer inte med om man zoomar efteråt - en medveten, enklare
 *  avvägning eftersom visualiseringen är en togglad markering, inte en
 *  kontinuerlig förhandsvisning. */
function visualScale(anchor: THREE.Vector3): number {
  const distance = world.camera.three.position.distanceTo(anchor);
  return THREE.MathUtils.clamp(distance * 0.025, 0.05, 1.5);
}

/** En "tjock linje" som en cylindrisk mesh istället för en THREE.Line -
 *  vanliga WebGL-linjer ritas i de flesta webbläsare alltid med 1px oavsett
 *  linewidth (GL-begränsning), vilket gjorde måttlinjerna svåra att se. En
 *  mesh ger pålitlig, synlig tjocklek i alla webbläsare. */
function addThickLine(start: THREE.Vector3, end: THREE.Vector3, radius: number): void {
  const direction = end.clone().sub(start);
  const length = direction.length();
  if (length < 1e-6) return;
  const geometry = new THREE.CylinderGeometry(radius, radius, length, 8, 1);
  geometry.translate(0, length / 2, 0);
  const mesh = new THREE.Mesh(
    geometry,
    new THREE.MeshBasicMaterial({ color: METRIC_VIS_COLOR, depthTest: false }),
  );
  mesh.position.copy(start);
  mesh.quaternion.setFromUnitVectors(new THREE.Vector3(0, 1, 0), direction.normalize());
  metricsVisualizationGroup.add(mesh);
}

/** Ritar text (t.ex. "3.42 m") till en canvas-textur på en THREE.Sprite -
 *  spriten är alltid kameravänd och kräver ingen extra CSS2D-renderare, så
 *  den kan läggas till scenen som vilket 3D-objekt som helst. */
function createLabelSprite(text: string): THREE.Sprite {
  const canvas = document.createElement("canvas");
  const context = canvas.getContext("2d");
  // Överdimensionerad canvas (renderas nedskalad via LABEL_SCREEN_HEIGHT) ger
  // skarpare text än att matcha den faktiska skärmstorleken 1:1.
  const fontSize = 64;
  const font = `200 ${fontSize}px system-ui, -apple-system, sans-serif`;
  let textWidth = fontSize * text.length * 0.6;
  if (context) {
    context.font = font;
    textWidth = context.measureText(text).width;
  }

  const paddingX = 34;
  const paddingY = 24;
  const shadowMargin = 10;
  canvas.width = Math.ceil(textWidth + paddingX * 2 + shadowMargin * 2);
  canvas.height = fontSize + paddingY * 2 + shadowMargin * 2;

  if (context) {
    // Canvas-storleksändringen ovan nollställer contexten, så font/stilar
    // måste sättas igen innan något ritas.
    context.font = font;
    context.textBaseline = "middle";
    context.textAlign = "center";

    const w = canvas.width - shadowMargin * 2;
    const h = canvas.height - shadowMargin * 2;
    const radius = 16;
    const drawBadge = (x0: number, y0: number) => {
      context.beginPath();
      context.moveTo(x0 + radius, y0);
      context.arcTo(x0 + w, y0, x0 + w, y0 + h, radius);
      context.arcTo(x0 + w, y0 + h, x0, y0 + h, radius);
      context.arcTo(x0, y0 + h, x0, y0, radius);
      context.arcTo(x0, y0, x0 + w, y0, radius);
      context.closePath();
    };

    // Mjuk skugga runt hela badgen, så den lyfter tydligt från VILKEN
    // bakgrund som helst i 3D-vyn istället för att bara luta sig på en
    // (ibland svårlästkontrasterande) kantfärg.
    context.save();
    context.shadowColor = "rgba(0, 0, 0, 0.35)";
    context.shadowBlur = 10;
    context.fillStyle = "#ffffff";
    drawBadge(shadowMargin, shadowMargin);
    context.fill();
    context.restore();

    context.strokeStyle = "#a855f7";
    context.lineWidth = 2.5;
    drawBadge(shadowMargin, shadowMargin);
    context.stroke();

    context.fillStyle = "#1e1024";
    context.fillText(text, shadowMargin + w / 2, shadowMargin + h / 2);
  }

  const texture = new THREE.CanvasTexture(canvas);
  // Canvasen är högupplöst men spriten visas litet på skärmen (se
  // LABEL_SCREEN_HEIGHT) - utan detta försöker WebGL bygga mipmaps för en
  // icke-kvadratisk canvas-storlek, vilket gav ett trasigt, "smetigt"
  // textutseende. Ren bilinjär filtrering utan mipmaps ger mjuk nedskalning
  // istället.
  texture.generateMipmaps = false;
  texture.minFilter = THREE.LinearFilter;
  texture.magFilter = THREE.LinearFilter;
  const material = new THREE.SpriteMaterial({
    map: texture,
    depthTest: false,
    transparent: true,
    // Konstant storlek på SKÄRMEN oavsett kameraavstånd (till skillnad från
    // vanlig sprite-skalning, som krymper med avståndet i perspektiv) - detta
    // var den uttryckliga bugg-rapporten: etiketterna gick inte att läsa utan
    // att zooma in väldigt nära.
    sizeAttenuation: false,
  });
  const sprite = new THREE.Sprite(material);
  sprite.userData.aspect = canvas.width / canvas.height;
  sprite.renderOrder = 1000;
  return sprite;
}

// Fast skärmstorlek (inte world units när sizeAttenuation är false, se
// createLabelSprite) - kalibrerad genom visuell test så texten är läsbar
// både inzoomat och utzoomat över hela byggnaden.
const LABEL_SCREEN_HEIGHT = 0.055;

function addLabel(position: THREE.Vector3, text: string): void {
  const sprite = createLabelSprite(text);
  const aspect = (sprite.userData.aspect as number | undefined) ?? 2;
  sprite.scale.set(LABEL_SCREEN_HEIGHT * aspect, LABEL_SCREEN_HEIGHT, 1);
  sprite.position.copy(position);
  metricsVisualizationGroup.add(sprite);
}

/** Ritar ett CAD-liknande måttstreck: en tjock linje mellan start/end, korta
 *  vinkelräta ändmarkeringar, en kula i varje ände och en etikett med det
 *  uppmätta värdet, förskjuten ut från linjen så den inte täcker den. */
function addDimensionLine(
  start: THREE.Vector3,
  end: THREE.Vector3,
  labelText: string,
  scale: number,
): void {
  const radius = scale * OUTLINE_RADIUS_FACTOR;
  addThickLine(start, end, radius);

  const direction = end.clone().sub(start);
  const length = direction.length();
  if (length < 1e-6) {
    addLabel(start, labelText);
    return;
  }
  direction.normalize();

  const arbitrary =
    Math.abs(direction.y) < 0.9 ? new THREE.Vector3(0, 1, 0) : new THREE.Vector3(1, 0, 0);
  const capDirection = new THREE.Vector3().crossVectors(direction, arbitrary).normalize();
  const capSize = Math.min(scale * 0.5, length * 0.3);
  for (const point of [start, end]) {
    addThickLine(
      point.clone().addScaledVector(capDirection, -capSize),
      point.clone().addScaledVector(capDirection, capSize),
      radius,
    );
  }

  const markerMaterial = new THREE.MeshBasicMaterial({ color: METRIC_VIS_COLOR, depthTest: false });
  for (const point of [start, end]) {
    const marker = new THREE.Mesh(new THREE.SphereGeometry(radius * 2.5, 12, 12), markerMaterial);
    marker.position.copy(point);
    metricsVisualizationGroup.add(marker);
  }

  const midpoint = start.clone().add(end).multiplyScalar(0.5);
  const labelPosition = midpoint.addScaledVector(capDirection, scale * 0.5);
  addLabel(labelPosition, labelText);
}

/** Net Area: fyller den FAKTISKA triangulerade ytan (inte en bounding-
 *  rektangel) - exakt samma trianglar som Net Area-talet summerade arean av,
 *  så fyllningen alltid stämmer geometriskt med talet. Hål/öppningar syns
 *  som riktiga luckor eftersom det helt enkelt inte finns några trianglar
 *  där. Konturen (faceOutlineSegments) ritas ovanpå, inklusive runt hål. */
function addFaceHighlight(
  triangles: THREE.Vector3[],
  outlineSegments: THREE.Vector3[],
  labelText: string,
  scale: number,
): void {
  if (triangles.length > 0) {
    const geometry = new THREE.BufferGeometry().setFromPoints(triangles);
    metricsVisualizationGroup.add(
      new THREE.Mesh(
        geometry,
        new THREE.MeshBasicMaterial({
          color: METRIC_VIS_COLOR,
          transparent: true,
          opacity: 0.32,
          side: THREE.DoubleSide,
          depthWrite: false,
          depthTest: false,
        }),
      ),
    );
  }

  const radius = scale * OUTLINE_RADIUS_FACTOR;
  for (let i = 0; i + 1 < outlineSegments.length; i += 2) {
    addThickLine(outlineSegments[i], outlineSegments[i + 1], radius);
  }

  const points = triangles.length > 0 ? triangles : outlineSegments;
  if (points.length === 0) return;
  const center = new THREE.Vector3();
  for (const point of points) center.add(point);
  center.divideScalar(points.length);
  addLabel(center, labelText);
}

/** Omkrets: ritar objektets FAKTISKA ytterkant (segmentparen som
 *  computeMetrics hittade via kant-delnings-analysen, med hål/öppningar
 *  redan uteslutna), inte en bounding-rektangel - annars stämmer inte linjen
 *  med det uppmätta talet för L-formade/urtagna/runda ytor. */
function addPerimeterOutline(segments: THREE.Vector3[], labelText: string, scale: number): void {
  const radius = scale * OUTLINE_RADIUS_FACTOR;
  for (let i = 0; i + 1 < segments.length; i += 2) {
    addThickLine(segments[i], segments[i + 1], radius);
  }
  if (segments.length === 0) return;

  const center = new THREE.Vector3();
  for (const point of segments) center.add(point);
  center.divideScalar(segments.length);
  addLabel(center, labelText);
}

/** Volume: fyller HELA objektets faktiska mesh (inte en omskriven box) -
 *  samma geometri som model.getItemsVolume mäter volymen av, så en vägg med
 *  ett hål t.ex. inte visas som en hel, tät box. */
function addVolumeHighlight(
  triangles: THREE.Vector3[],
  center: THREE.Vector3,
  labelText: string,
): void {
  if (triangles.length > 0) {
    const geometry = new THREE.BufferGeometry().setFromPoints(triangles);
    metricsVisualizationGroup.add(
      new THREE.Mesh(
        geometry,
        new THREE.MeshBasicMaterial({
          color: METRIC_VIS_COLOR,
          transparent: true,
          opacity: 0.35,
          side: THREE.DoubleSide,
          depthWrite: false,
          depthTest: false,
        }),
      ),
    );
  }
  addLabel(center, labelText);
}

function showMetricVisualization(metrics: ItemMetrics, key: MetricVisualizationKey): void {
  const visualization = metrics.visualization;
  if (!visualization) return;

  clearMetricVisualizationGroup();
  const labelText = formatMetricValue(metrics, key);
  const scale = visualScale(visualization.center);
  switch (key) {
    case "thickness":
      addDimensionLine(...visualization.thicknessLine, labelText, scale);
      break;
    case "width":
      addDimensionLine(...visualization.widthLine, labelText, scale);
      break;
    case "length":
      addDimensionLine(...visualization.lengthLine, labelText, scale);
      break;
    case "height":
      addDimensionLine(...visualization.heightLine, labelText, scale);
      break;
    case "sideArea":
      addFaceHighlight(visualization.faceTriangles, visualization.faceOutlineSegments, labelText, scale);
      break;
    case "perimeter":
      addPerimeterOutline(visualization.perimeterSegments, labelText, scale);
      break;
    case "volume":
      addVolumeHighlight(visualization.volumeTriangles, visualization.center, labelText);
      break;
  }
  metricsVisualizationGroup.visible = true;
  postproductionRenderer.needsUpdate = true;
}

function hideMetricVisualization(): void {
  clearMetricVisualizationGroup();
  metricsVisualizationGroup.visible = false;
  postproductionRenderer.needsUpdate = true;
}

// Litet, halvgenomskinligt plan som visar var ett snitt skulle hamna när
// man för muspekaren över en yta i snittläge - samma kompakta storlek som
// det faktiska snittplanet (clipper.size) så förhandsgranskningen matchar.
const clipPreviewFill = new THREE.Mesh(
  new THREE.PlaneGeometry(0.6, 0.6),
  new THREE.MeshBasicMaterial({
    color: new THREE.Color("#0091d5"),
    transparent: true,
    opacity: 0.3,
    side: THREE.DoubleSide,
    depthWrite: false,
  }),
);
const clipPreviewOutline = new THREE.LineSegments(
  new THREE.EdgesGeometry(clipPreviewFill.geometry),
  new THREE.LineBasicMaterial({ color: new THREE.Color("#0091d5") }),
);
const clipPreviewGroup = new THREE.Group();
clipPreviewGroup.add(clipPreviewFill, clipPreviewOutline);
clipPreviewGroup.visible = false;
clipPreviewGroup.renderOrder = 999;
world.scene.three.add(clipPreviewGroup);

let clipPreviewRafScheduled = false;

async function updateClipPreview() {
  if (!clipper.enabled) {
    clipPreviewGroup.visible = false;
    return;
  }

  const result = (await caster.castRay()) as unknown as
    | { point?: THREE.Vector3; normal?: THREE.Vector3 }
    | null;

  if (!result?.point || !result.normal) {
    clipPreviewGroup.visible = false;
    return;
  }

  clipPreviewGroup.position.copy(result.point);
  clipPreviewGroup.quaternion.setFromUnitVectors(
    new THREE.Vector3(0, 0, 1),
    result.normal.clone().normalize(),
  );
  clipPreviewGroup.visible = true;
}

canvas.addEventListener("mousemove", () => {
  if (!clipper.enabled) {
    clipPreviewGroup.visible = false;
    return;
  }
  if (clipPreviewRafScheduled) return;
  clipPreviewRafScheduled = true;
  requestAnimationFrame(() => {
    clipPreviewRafScheduled = false;
    void updateClipPreview().then(() => { postproductionRenderer.needsUpdate = true; });
  });
});

canvas.addEventListener("mouseleave", () => {
  clipPreviewGroup.visible = false;
  measurePreviewGroup.visible = false;
  postproductionRenderer.needsUpdate = true;
});

// Mätverktyget har en egen inbyggd snapp-markör, men den är bara 6px och lätt
// att missa. Den här förhandsgranskningen gör punkt-/linje-/ytträffar lika
// tydliga som snittförhandsgranskningen ovan - färgkodad efter träfftyp
// (0 = punkt, 1 = linje, 2 = yta; se FRAGS.SnappingClass) och skalad efter
// kameraavstånd så den syns lika bra oavsett zoomnivå.
const MEASURE_SNAP_COLORS: Record<number, THREE.Color> = {
  0: new THREE.Color("#f59e0b"),
  1: new THREE.Color("#22c55e"),
  2: new THREE.Color("#0091d5"),
};

const measureMarkerSphere = new THREE.Mesh(
  new THREE.SphereGeometry(1, 16, 16),
  new THREE.MeshBasicMaterial({ depthTest: false, transparent: true, opacity: 0.85 }),
);
const measureEdgeLine = new THREE.Line(
  new THREE.BufferGeometry().setFromPoints([new THREE.Vector3(), new THREE.Vector3()]),
  new THREE.LineBasicMaterial({ depthTest: false, linewidth: 2 }),
);
measureEdgeLine.visible = false;

const measurePreviewGroup = new THREE.Group();
measurePreviewGroup.add(measureMarkerSphere, measureEdgeLine);
measurePreviewGroup.visible = false;
measurePreviewGroup.renderOrder = 999;
world.scene.three.add(measurePreviewGroup);

let measurePreviewRafScheduled = false;

// "Vinkelrätt" mätning (yta till yta): första klicket sätter en ankarpunkt +
// planets normal, andra klicket projicerar sin träffpunkt på normalen genom
// ankaret - så måttet blir det vinkelräta avståndet till den första ytan
// (t.ex. väggtjocklek) istället för den råa diagonalen mellan klickpunkterna.
// Separat state/flöde från measurer.create() eftersom biblioteket inte
// exponerar ett sätt att styra var den andra punkten hamnar.
interface PerpendicularAnchor {
  point: THREE.Vector3;
  // null när första klicket träffade något utan ytnormal (t.ex. en ren
  // punkt-/linjesnap) - ankaret sätts ändå så klick-paret följer samma
  // två-klicks-rytm som det vanliga måttet, det blir bara inget vinkelrätt
  // mått för just det paret.
  normal: THREE.Vector3 | null;
}

let perpendicularAnchor: PerpendicularAnchor | null = null;

const perpendicularAnchorMarker = new THREE.Mesh(
  new THREE.SphereGeometry(1, 16, 16),
  new THREE.MeshBasicMaterial({
    color: new THREE.Color("#0091d5"),
    depthTest: false,
    transparent: true,
    opacity: 0.85,
  }),
);
perpendicularAnchorMarker.visible = false;
const perpendicularPreviewLine = new THREE.Line(
  new THREE.BufferGeometry().setFromPoints([new THREE.Vector3(), new THREE.Vector3()]),
  new THREE.LineBasicMaterial({ color: new THREE.Color("#0091d5"), depthTest: false }),
);
perpendicularPreviewLine.visible = false;
measurePreviewGroup.add(perpendicularAnchorMarker, perpendicularPreviewLine);

function cancelPerpendicularMeasure() {
  perpendicularAnchor = null;
  perpendicularAnchorMarker.visible = false;
  perpendicularPreviewLine.visible = false;
  postproductionRenderer.needsUpdate = true;
}

async function updateMeasurePreview() {
  if (!measurer.enabled) {
    measurePreviewGroup.visible = false;
    return;
  }

  const result = (await caster.castRay({ snappingClasses: [0, 1, 2] })) as unknown as
    | {
        point?: THREE.Vector3;
        normal?: THREE.Vector3;
        snappingClass?: number;
        snappedEdgeP1?: THREE.Vector3;
        snappedEdgeP2?: THREE.Vector3;
      }
    | null;

  if (!result?.point) {
    measurePreviewGroup.visible = false;
    return;
  }

  const color = MEASURE_SNAP_COLORS[result.snappingClass ?? 2] ?? MEASURE_SNAP_COLORS[2];
  const distance = world.camera.three.position.distanceTo(result.point);
  const scale = THREE.MathUtils.clamp(distance * 0.012, 0.03, 0.5);

  measureMarkerSphere.position.copy(result.point);
  measureMarkerSphere.scale.setScalar(scale);
  (measureMarkerSphere.material as THREE.MeshBasicMaterial).color.copy(color);

  // Visar var det vinkelräta måttet skulle hamna bredvid den vanliga
  // hovermarkören, så båda syns innan man klickar färdigt.
  if (perpendicularAnchor?.normal) {
    const { point: p1, normal: n1 } = perpendicularAnchor;
    const offset = result.point.clone().sub(p1).dot(n1);
    const projected = p1.clone().addScaledVector(n1, offset);
    perpendicularPreviewLine.geometry.setFromPoints([p1, projected]);
    perpendicularPreviewLine.visible = true;
  } else {
    perpendicularPreviewLine.visible = false;
  }

  if (result.snappingClass === 1 && result.snappedEdgeP1 && result.snappedEdgeP2) {
    measureEdgeLine.geometry.setFromPoints([result.snappedEdgeP1, result.snappedEdgeP2]);
    (measureEdgeLine.material as THREE.LineBasicMaterial).color.copy(color);
    measureEdgeLine.visible = true;
  } else {
    measureEdgeLine.visible = false;
  }

  if (perpendicularAnchor) {
    perpendicularAnchorMarker.position.copy(perpendicularAnchor.point);
    perpendicularAnchorMarker.scale.setScalar(scale);
    perpendicularAnchorMarker.visible = true;
  } else {
    perpendicularAnchorMarker.visible = false;
  }

  measurePreviewGroup.visible = true;
}

canvas.addEventListener("mousemove", () => {
  if (!measurer.enabled) {
    measurePreviewGroup.visible = false;
    return;
  }
  if (measurePreviewRafScheduled) return;
  measurePreviewRafScheduled = true;
  requestAnimationFrame(() => {
    measurePreviewRafScheduled = false;
    void updateMeasurePreview().then(() => { postproductionRenderer.needsUpdate = true; });
  });
});

async function handlePerpendicularMeasureClick() {
  const result = (await caster.castRay()) as unknown as
    | { point?: THREE.Vector3; normal?: THREE.Vector3 }
    | null;

  if (!perpendicularAnchor) {
    // Sätts även utan normal (t.ex. en punkt-/linjesnap) så klickparet ändå
    // följer samma två-klicks-rytm som det vanliga måttet - annars skulle
    // nästa klick felaktigt tolkas som ett nytt förstaklick.
    if (result?.point) {
      perpendicularAnchor = { point: result.point.clone(), normal: result.normal?.clone().normalize() ?? null };
    }
    return;
  }

  const { point: p1, normal: n1 } = perpendicularAnchor;
  if (n1 && result?.point) {
    const offset = result.point.clone().sub(p1).dot(n1);
    if (Math.abs(offset) > 1e-6) {
      const p2 = p1.clone().addScaledVector(n1, offset);
      measurer.list.add(new OBF.Line(p1, p2));
      postproductionRenderer.needsUpdate = true;
    }
  }
  cancelPerpendicularMeasure();
}

function isTypingTarget(target: EventTarget | null): boolean {
  return (
    target instanceof HTMLInputElement ||
    target instanceof HTMLTextAreaElement ||
    target instanceof HTMLSelectElement ||
    (target instanceof HTMLElement && target.isContentEditable)
  );
}

// Tangenterna 5/0/9 växlar mellan att markera objekt, sektionera och mäta -
// samma tre lägen som knapparna i verktygsfältet styr.
window.addEventListener("keydown", (event) => {
  if (isTypingTarget(event.target)) return;

  if (event.code === "Delete" || event.code === "Backspace") {
    if (clipper.enabled) { clipper.delete(world); postproductionRenderer.needsUpdate = true; }
    if (measurer.enabled) { measurer.delete(); postproductionRenderer.needsUpdate = true; }
    return;
  }

  if (event.code === "Escape") {
    cancelPerpendicularMeasure();
    return;
  }

  if (event.code === "Digit5" || event.code === "Numpad5") {
    setActiveTool("select");
  } else if (event.code === "Digit0" || event.code === "Numpad0") {
    setActiveTool(clipper.enabled ? "select" : "clip");
  } else if (event.code === "Digit9" || event.code === "Numpad9") {
    setActiveTool(measurer.enabled ? "select" : "measure");
  }
});

clipClear.addEventListener("click", () => {
  clipper.deleteAll();
  postproductionRenderer.needsUpdate = true;
});

measureClear.addEventListener("click", () => {
  measurer.list.clear();
  cancelPerpendicularMeasure();
  postproductionRenderer.needsUpdate = true;
});

// Mängdavtagning
interface QuantityRow {
  modelId: string;
  localId: number;
  values: Record<string, string | number>;
}

interface TakeoffColumn {
  key: string;
  label: string;
  kind: "text" | "number" | "select";
  grouped: boolean;
}

/**
 * Boverkets byggdelskategorisering för klimatdeklaration (förordning 2021:789 § 5):
 * bärande konstruktion delas i grundläggning och övrigt, plus klimatskärm och
 * innerväggar. "Ingår ej" låter användaren aktivt markera objekt som inte
 * omfattas (t.ex. ytskikt), skilt från att bara lämna klassningen tom.
 */
const BOVERKET_COLUMN_KEY = "boverketCategory";
const BOVERKET_COLUMN_LABEL = "Boverket-kategori (klimatdeklaration)";
/** Kortare rubriktext bara för tabellhuvudet i mängdavtagningen - annars
 *  blev hela kolumnen lika bred som den fulla, oböjliga rubrikraden trots
 *  att själva innehållet (den bredd-begränsade selecten) är mycket smalare.
 *  Den fulla texten finns kvar som tooltip på rubrikcellen. */
const BOVERKET_COLUMN_HEADER_SHORT = "Boverket-kategori";
const BOVERKET_CATEGORIES = [
  "Bärande konstruktionsdelar – Grundläggning",
  "Bärande konstruktionsdelar – Övriga",
  "Klimatskärm",
  "Innerväggar",
  "Ingår ej",
] as const;

/**
 * Allmän byggdelsklassificering, inspirerad av den bifogade referensfilen
 * "Building Elements - General.classification" (ett Solibri-exempel som
 * klassar objekt efter kategori + Pset_*Common.IsExternal/LoadBearing, t.ex.
 * "External Walls"/"Load Bearing Walls"/"Partition Walls"). Filen är ett
 * proprietärt Java-serialiserat binärformat (inte parsningsbart i webbläsaren),
 * så kategorierna nedan är en handöversatt, IFC-anpassad motsvarighet till
 * den taxonomin - till skillnad från Boverket-klassningen (som bara har 5
 * kategorier för klimatdeklarationens obligatoriska omfattning) är den här
 * mycket mer finmaskig, tänkt för allmän byggdelsgruppering/mängdavtagning.
 * Några av originalfilens allra smalaste underkategorier (t.ex. separata
 * "Electrical Cabinets"/"Lightning"/"Hatch") har slagits ihop till bredare,
 * mer tillförlitligt IFC-kategoriserbara grupper istället för att gissa på
 * namnmönster som inte går att verifiera utan Solibris egen motor.
 */
const GENERAL_COLUMN_KEY = "generalClassification";
const GENERAL_COLUMN_LABEL = "Byggdel (allmän klassificering)";
const GENERAL_COLUMN_HEADER_SHORT = "Byggdel";
const GENERAL_CATEGORIES = [
  "Grundläggning – Pålar",
  "Grundläggning – Fundament",
  "Grundläggning – Bottenplatta",
  "Bärande – Pelare och balkar",
  "Bärande väggar",
  "Ytterväggar",
  "Innerväggar – Icke bärande",
  "Bjälklag",
  "Takkonstruktion",
  "Trappor och ramper",
  "Räcken – Inomhus",
  "Räcken – Utomhus",
  "Fönster",
  "Ytterdörrar",
  "Innerdörrar",
  "Undertak och ytskikt",
  "Möbler och inredning",
  "VVS-installationer",
  "Ventilationssystem",
  "Elinstallationer",
  "Brand och säkerhet",
  "Hissar och rulltrappor",
  "Balkonger",
  "Övrig utrustning",
] as const;

/** Flaggar objekt som ingår i en dubblettgrupp (se findDuplicates) med
 *  Ja/Nej, så det går att sortera/gruppera mängdavtagningen på misstänkta
 *  dubbletter utan att behöva öppna Dubbletter-panelen separat. */
const DUPLICATE_COLUMN_KEY = "possibleDuplicate";
const DUPLICATE_COLUMN_LABEL = "Möjlig dubblett";

/** Samma relationsdata (IsTypedBy/HasAssociations) som egenskapspanelens
 *  Typ-/Material-/Klassificeringssektioner, sammanfattade till en textrad per
 *  objekt så de kan bli kolumner i mängdavtagningen - t.ex. för att gruppera
 *  eller filtrera på material. */
const TYPE_COLUMN_KEY = "typeName";
const TYPE_COLUMN_LABEL = "Typ";
const MATERIAL_COLUMN_KEY = "material";
const MATERIAL_COLUMN_LABEL = "Material";
const CLASSIFICATION_COLUMN_KEY = "classification";
const CLASSIFICATION_COLUMN_LABEL = "Klassificering";

interface TakeoffTableRow {
  cells: (string | number)[];
  /** De underliggande objekten som raden representerar - ett enda vid ogrupperad
   *  vy, flera när raden är en gruppsumma. */
  sourceRows: QuantityRow[];
  groupLabel?: string;
}

const METRIC_COLUMN_DEFS: { key: string; label: string }[] = [
  { key: "sideArea", label: "Net Area" },
  { key: "volume", label: "Volume" },
  { key: "thickness", label: "Tjocklek (m)" },
  { key: "height", label: "Höjd (m)" },
  { key: "width", label: "Bredd (m)" },
  { key: "length", label: "Längd (m)" },
  { key: "perimeter", label: "Omkrets (m)" },
];
const METRIC_KEYS = new Set(METRIC_COLUMN_DEFS.map((c) => c.key));
const FRIENDLY_ATTRIBUTE_LABELS: Record<string, string> = {
  _category: "Kategori",
  _guid: "GUID",
  Name: "Namn",
};

/** Kolumnnyckel-prefix + separator för property set-/quantity set-egenskaper
 *  (se psetPropertyEntries) - "pset:Pset_BeamCommon::LoadBearing" - så att de
 *  inte krockar med vanliga platta attributnycklar (t.ex. "Name") som redan
 *  används för Attribut-kolumnerna. */
const PSET_COLUMN_PREFIX = "pset:";
const PSET_COLUMN_SEPARATOR = "::";

function makePsetColumnKey(psetName: string, propName: string): string {
  return `${PSET_COLUMN_PREFIX}${psetName}${PSET_COLUMN_SEPARATOR}${propName}`;
}

function parsePsetColumnKey(key: string): { psetName: string; propName: string } | null {
  if (!key.startsWith(PSET_COLUMN_PREFIX)) return null;
  const rest = key.slice(PSET_COLUMN_PREFIX.length);
  const sepIndex = rest.indexOf(PSET_COLUMN_SEPARATOR);
  if (sepIndex === -1) return null;
  return { psetName: rest.slice(0, sepIndex), propName: rest.slice(sepIndex + PSET_COLUMN_SEPARATOR.length) };
}

/** Fylls under computeQuantityTakeoff med de pset-kolumnnycklar vars värden
 *  visat sig vara numeriska (t.ex. quantity set-egenskaper som area/volym) -
 *  läses av columnKind så att de kan summeras vid gruppering precis som de
 *  inbyggda Beräknat-kolumnerna. */
const numericPsetColumns = new Set<string>();

function columnKind(key: string): "text" | "number" | "select" {
  if (key === BOVERKET_COLUMN_KEY || key === GENERAL_COLUMN_KEY) return "select";
  if (METRIC_KEYS.has(key)) return "number";
  if (numericPsetColumns.has(key)) return "number";
  return "text";
}

function columnLabel(key: string): string {
  if (key === "model") return "Modell";
  if (key === BOVERKET_COLUMN_KEY) return BOVERKET_COLUMN_LABEL;
  if (key === GENERAL_COLUMN_KEY) return GENERAL_COLUMN_LABEL;
  if (key === DUPLICATE_COLUMN_KEY) return DUPLICATE_COLUMN_LABEL;
  if (key === TYPE_COLUMN_KEY) return TYPE_COLUMN_LABEL;
  if (key === MATERIAL_COLUMN_KEY) return MATERIAL_COLUMN_LABEL;
  if (key === CLASSIFICATION_COLUMN_KEY) return CLASSIFICATION_COLUMN_LABEL;
  const metric = METRIC_COLUMN_DEFS.find((c) => c.key === key);
  if (metric) return metric.label;
  const psetKey = parsePsetColumnKey(key);
  if (psetKey) return `${psetKey.propName} (${psetKey.psetName})`;
  return FRIENDLY_ATTRIBUTE_LABELS[key] ?? key;
}

/** Läser ut egenskaper/kvantiteter från ett objekts IsDefinedBy-relationer
 *  (property sets OCH quantity sets - samma data som egenskapspanelens
 *  Pset-sektioner, se renderPset) till platta kolumnvärden. En quantity sets
 *  "egenskaper" ligger under HasQuantities istället för HasProperties och har
 *  sitt värde under en typad nyckel (LengthValue/AreaValue/...) istället för
 *  NominalValue - båda formerna hanteras här. */
const QUANTITY_VALUE_KEYS = [
  "NominalValue",
  "LengthValue",
  "AreaValue",
  "VolumeValue",
  "WeightValue",
  "CountValue",
  "TimeValue",
];

function psetPropertyValue(prop: FRAGS.ItemData): unknown {
  for (const key of QUANTITY_VALUE_KEYS) {
    const attr = prop[key];
    if (isAttribute(attr)) return attr.value;
  }
  return undefined;
}

function psetPropertyEntries(data: FRAGS.ItemData): { key: string; value: unknown }[] {
  const entries: { key: string; value: unknown }[] = [];
  for (const pset of relationItems(data, "IsDefinedBy")) {
    const psetName = itemName(pset);
    if (!psetName) continue;
    const props = [...relationItems(pset, "HasProperties"), ...relationItems(pset, "HasQuantities")];
    for (const prop of props) {
      const propName = itemName(prop);
      if (!propName) continue;
      const value = psetPropertyValue(prop);
      if (value === undefined) continue;
      entries.push({ key: makePsetColumnKey(psetName, propName), value });
    }
  }
  return entries;
}

let quantityRows: QuantityRow[] = [];
let availableAttributeColumns: string[] = [];
let availablePsetColumns: string[] = [];
let lastTakeoffTable: { headers: string[]; rows: TakeoffTableRow[] } | null = null;
let takeoffSelection: SelectableItem[] = [];
let takeoffRangeAnchor: number | null = null;

/** Kryssrutemarkering för bulk-åtgärder (t.ex. "Sätt 'Ingår ej'") - separat
 *  från den vanliga rad-/gruppmarkeringen ovan, som styr 3D-highlight/
 *  egenskapspanelen. Nyckel: "modelId::localId" (samma format som itemKey). */
const takeoffCheckedItems = new Set<string>();

function getCheckedTakeoffItems(): SelectableItem[] {
  return quantityRows
    .filter((r) => takeoffCheckedItems.has(itemKey({ modelId: r.modelId, localId: r.localId })))
    .map((r) => ({ modelId: r.modelId, localId: r.localId }));
}

function updateTakeoffBulkBar() {
  const count = takeoffCheckedItems.size;
  takeoffBulkBar.classList.toggle("hidden", count === 0);
  takeoffBulkCount.textContent = `${count} markerad${count === 1 ? "" : "e"}`;
}

/** Boverket-klassning per objekt, hålls utanför quantityRows så den överlever
 *  att mängdavtagningen räknas om (t.ex. via Uppdatera). */
const boverketCategoryByItem = new Map<string, string>();

function isBoverketColumnVisible(): boolean {
  return takeoffColumns.some((c) => c.key === BOVERKET_COLUMN_KEY && !c.grouped);
}

/** Om alla objekt i en grupp (eller markering) delar samma Boverket-klassning
 *  returneras den, annars "mixed" - används för att visa "(Blandat)" istället
 *  för ett missvisande värde. */
function resolveGroupBoverketValue(items: { modelId: string; localId: number }[]): string | "mixed" {
  const values = new Set(
    items.map((item) => boverketCategoryByItem.get(itemKey(item)) ?? ""),
  );
  if (values.size > 1) return "mixed";
  return [...values][0] ?? "";
}

function renderBoverketOptions(current: string | "mixed"): string {
  const placeholderLabel = current === "mixed" ? "(Blandat)" : "Ej klassad";
  const placeholderSelected = current === "mixed" || current === "";
  return [
    `<option value=""${placeholderSelected ? " selected" : ""}>${placeholderLabel}</option>`,
    ...BOVERKET_CATEGORIES.map(
      (cat) =>
        `<option value="${escapeHtml(cat)}"${current === cat ? " selected" : ""}>${escapeHtml(cat)}</option>`,
    ),
  ].join("");
}

function renderBoverketCell(rowIndex: number, sourceRows: QuantityRow[]): string {
  const options = renderBoverketOptions(resolveGroupBoverketValue(sourceRows));
  return `<td class="takeoff-col-narrow"><select class="boverket-select" data-row-index="${rowIndex}">${options}</select></td>`;
}

/** Samma klassificeringskontroll som i mängdavtagningen, men för egenskapspanelen
 *  - låter en klassificera ett eller flera markerade objekt oavsett om markeringen
 *  kom från modellträdet eller mängdavtagningen. */
function renderBoverketSelector(items: SelectableItem[]): string {
  const options = renderBoverketOptions(resolveGroupBoverketValue(items));
  return `
    <div class="prop-section">
      <div class="prop-section-title">Klimatdeklaration</div>
      <select class="prop-boverket-select">${options}</select>
    </div>
  `;
}

/** Sätter (eller rensar) Boverket-klassningen för en uppsättning objekt -
 *  används både från mängdavtagningens tabellrader och från egenskapspanelens
 *  markeringsvy, så en klassificering kan göras oavsett varifrån objekten valdes. */
function applyBoverketCategoryToItems(items: { modelId: string; localId: number }[], value: string) {
  for (const item of items) {
    const key = itemKey(item);
    if (value) boverketCategoryByItem.set(key, value);
    else boverketCategoryByItem.delete(key);

    const row = quantityRows.find(
      (r) => r.modelId === item.modelId && r.localId === item.localId,
    );
    if (row) row.values[BOVERKET_COLUMN_KEY] = value;
  }
  if (!takeoffPanel.classList.contains("hidden")) renderTakeoffTable();
}

/** Samma mönster som boverketCategoryByItem/de fem funktionerna ovan, men för
 *  den allmänna byggdelsklassificeringen (se GENERAL_CATEGORIES) - hålls som
 *  en helt egen, parallell klassning istället för att byggas in i Boverket-
 *  funktionerna, så de två klassificeringarna aldrig kan störa varandra. */
const generalCategoryByItem = new Map<string, string>();

function isGeneralColumnVisible(): boolean {
  return takeoffColumns.some((c) => c.key === GENERAL_COLUMN_KEY && !c.grouped);
}

function resolveGroupGeneralValue(items: { modelId: string; localId: number }[]): string | "mixed" {
  const values = new Set(items.map((item) => generalCategoryByItem.get(itemKey(item)) ?? ""));
  if (values.size > 1) return "mixed";
  return [...values][0] ?? "";
}

function renderGeneralOptions(current: string | "mixed"): string {
  const placeholderLabel = current === "mixed" ? "(Blandat)" : "Ej klassad";
  const placeholderSelected = current === "mixed" || current === "";
  return [
    `<option value=""${placeholderSelected ? " selected" : ""}>${placeholderLabel}</option>`,
    ...GENERAL_CATEGORIES.map(
      (cat) =>
        `<option value="${escapeHtml(cat)}"${current === cat ? " selected" : ""}>${escapeHtml(cat)}</option>`,
    ),
  ].join("");
}

function renderGeneralCell(rowIndex: number, sourceRows: QuantityRow[]): string {
  const options = renderGeneralOptions(resolveGroupGeneralValue(sourceRows));
  return `<td class="takeoff-col-narrow"><select class="general-select" data-row-index="${rowIndex}">${options}</select></td>`;
}

function renderGeneralSelector(items: SelectableItem[]): string {
  const options = renderGeneralOptions(resolveGroupGeneralValue(items));
  return `
    <div class="prop-section">
      <div class="prop-section-title">Byggdelsklassificering</div>
      <select class="prop-general-select">${options}</select>
    </div>
  `;
}

function applyGeneralCategoryToItems(items: { modelId: string; localId: number }[], value: string) {
  for (const item of items) {
    const key = itemKey(item);
    if (value) generalCategoryByItem.set(key, value);
    else generalCategoryByItem.delete(key);

    const row = quantityRows.find(
      (r) => r.modelId === item.modelId && r.localId === item.localId,
    );
    if (row) row.values[GENERAL_COLUMN_KEY] = value;
  }
  if (!takeoffPanel.classList.contains("hidden")) renderTakeoffTable();
}

/** Letar upp värdet för en Pset-egenskap oavsett vilket specifikt property set
 *  den råkar ligga i (t.ex. Pset_WallCommon.IsExternal vs
 *  Pset_SlabCommon.IsExternal) - samma idé som "Pset_*Common.IsExternal" i
 *  den bifogade Solibri-referensklassningen, fast matchat mot de redan
 *  beräknade pset-kolumnerna (se psetPropertyEntries) istället för att fråga
 *  modellen på nytt. */
function findPsetBoolean(row: QuantityRow, propName: string): boolean | undefined {
  const suffix = `${PSET_COLUMN_SEPARATOR}${propName}`;
  for (const [key, value] of Object.entries(row.values)) {
    if (!key.startsWith(PSET_COLUMN_PREFIX) || !key.endsWith(suffix)) continue;
    if (value === "true") return true;
    if (value === "false") return false;
  }
  return undefined;
}

/**
 * Föreslår en Boverket-kategori (klimatdeklaration) utifrån IFC-kategori och
 * vanliga Pset-attribut (IsExternal/LoadBearing) - samma sorts regelbaserade
 * logik som exempelklassningen "Building Elements - General.classification"
 * (kategori + Pset_*Common.IsExternal/LoadBearing → t.ex. "External Walls",
 * "Load Bearing Walls", "Partition Walls").
 *
 * Klimatdeklarationens obligatoriska omfattning är bara grundläggning, övrig
 * bärande stomme, klimatskärm och innerväggar - därför räknas möbler,
 * installationer (VVS/el/vent) och liknande som "Ingår ej". Genuint osäkra
 * kategorier (trappor, räcken, ytskikt, generiska proxy-element m.m.) ges
 * INGEN gissning (returnerar null) - hellre lämna dem oklassade för manuell
 * bedömning än riskera fel i ett underlag som ska användas i en lagstadgad
 * klimatdeklaration.
 */
function suggestBoverketCategory(row: QuantityRow): string | null {
  const category = String(row.values._category ?? "").toUpperCase();
  const isExternal = findPsetBoolean(row, "IsExternal");
  const loadBearing = findPsetBoolean(row, "LoadBearing");

  if (category === "IFCFOOTING" || category === "IFCPILE") {
    return "Bärande konstruktionsdelar – Grundläggning";
  }

  if (category === "IFCSLAB") {
    const predefinedType = String(row.values.PredefinedType ?? "").toUpperCase();
    const nameAndDescription = `${row.values.Name ?? ""} ${row.values.Description ?? ""}`.toLowerCase();
    if (
      predefinedType === "BASESLAB" ||
      nameAndDescription.includes("grund") ||
      nameAndDescription.includes("foundation")
    ) {
      return "Bärande konstruktionsdelar – Grundläggning";
    }
    if (loadBearing === false) return null;
    return "Bärande konstruktionsdelar – Övriga";
  }

  if (category === "IFCBEAM" || category === "IFCCOLUMN" || category === "IFCMEMBER") {
    if (loadBearing === false) return null;
    return "Bärande konstruktionsdelar – Övriga";
  }

  if (category === "IFCWALL" || category === "IFCWALLSTANDARDCASE") {
    if (isExternal === true) return "Klimatskärm";
    if (isExternal === false) return "Innerväggar";
    return null;
  }

  if (category === "IFCCURTAINWALL" || category === "IFCROOF") return "Klimatskärm";

  if (category === "IFCWINDOW" || category === "IFCDOOR") {
    if (isExternal === true) return "Klimatskärm";
    if (isExternal === false) return "Innerväggar";
    return null;
  }

  const NOT_INCLUDED_CATEGORIES = new Set([
    "IFCFURNISHINGELEMENT",
    "IFCFURNITURE",
    "IFCSYSTEMFURNITUREELEMENT",
    "IFCFLOWTERMINAL",
    "IFCFLOWSEGMENT",
    "IFCFLOWFITTING",
    "IFCFLOWCONTROLLER",
    "IFCFLOWSTORAGEDEVICE",
    "IFCFLOWMOVINGDEVICE",
    "IFCDISTRIBUTIONELEMENT",
    "IFCDISTRIBUTIONFLOWELEMENT",
    "IFCDISTRIBUTIONCONTROLELEMENT",
    "IFCELECTRICAPPLIANCE",
    "IFCSANITARYTERMINAL",
    "IFCLIGHTFIXTURE",
    "IFCOUTLET",
    "IFCPIPESEGMENT",
    "IFCPIPEFITTING",
    "IFCDUCTSEGMENT",
    "IFCDUCTFITTING",
    "IFCCABLECARRIERSEGMENT",
    "IFCCABLESEGMENT",
    "IFCFIRESUPPRESSIONTERMINAL",
    "IFCTRANSPORTELEMENT",
    "IFCENERGYCONVERSIONDEVICE",
    "IFCSWITCHINGDEVICE",
  ]);
  if (NOT_INCLUDED_CATEGORIES.has(category)) return "Ingår ej";

  return null;
}

/** Klassar automatiskt alla objekt som ännu inte har en Boverket-klassning
 *  (varken manuell eller från en tidigare körning) - rör aldrig en befintlig
 *  klassning. Se suggestBoverketCategory för regeluppsättningen. */
function autoClassifyBoverketCategories(): {
  classified: number;
  skippedExisting: number;
  unclear: number;
} {
  let classified = 0;
  let skippedExisting = 0;
  let unclear = 0;

  for (const row of quantityRows) {
    const key = itemKey(row);
    if (boverketCategoryByItem.has(key)) {
      skippedExisting++;
      continue;
    }
    const suggestion = suggestBoverketCategory(row);
    if (!suggestion) {
      unclear++;
      continue;
    }
    boverketCategoryByItem.set(key, suggestion);
    row.values[BOVERKET_COLUMN_KEY] = suggestion;
    classified++;
  }

  if (!takeoffPanel.classList.contains("hidden")) renderTakeoffTable();
  return { classified, skippedExisting, unclear };
}

/**
 * Föreslår en allmän byggdelskategori (se GENERAL_CATEGORIES) utifrån
 * IFC-kategori och samma Pset-attribut (IsExternal/LoadBearing) som
 * suggestBoverketCategory använder, fast med en mer finmaskig indelning som
 * följer referensfilens taxonomi (skiljer t.ex. på bärande/icke-bärande
 * väggar och inomhus-/utomhusräcken, och delar upp installationer i
 * VVS/ventilation/el istället för en enda "Ingår ej"-hink).
 *
 * Mindre försiktig än suggestBoverketCategory på ett par ställen (t.ex.
 * klassar balkar/pelare oavsett LoadBearing-flaggan, och en vägg utan känd
 * IsExternal/LoadBearing hamnar i "Innerväggar – Icke bärande" som ett
 * catch-all) - det matchar hur referensfilens egen sista, ovillkorade regel
 * för väggar ("Partition Walls") fungerar, och den här klassningen är inte
 * tänkt för en lagstadgad deklaration på samma sätt som Boverket-kolumnen,
 * så ett rimligt catch-all väger tyngre än att lämna allt oklassat.
 */
function suggestGeneralCategory(row: QuantityRow): string | null {
  const category = String(row.values._category ?? "").toUpperCase();
  const isExternal = findPsetBoolean(row, "IsExternal");
  const loadBearing = findPsetBoolean(row, "LoadBearing");
  const nameAndDescription = `${row.values.Name ?? ""} ${row.values.Description ?? ""}`.toLowerCase();

  if (nameAndDescription.includes("balkong") || nameAndDescription.includes("balcony")) {
    return "Balkonger";
  }

  if (category === "IFCPILE") return "Grundläggning – Pålar";
  if (category === "IFCFOOTING") return "Grundläggning – Fundament";

  if (category === "IFCSLAB") {
    const predefinedType = String(row.values.PredefinedType ?? "").toUpperCase();
    if (
      predefinedType === "BASESLAB" ||
      nameAndDescription.includes("grund") ||
      nameAndDescription.includes("foundation")
    ) {
      return "Grundläggning – Bottenplatta";
    }
    return "Bjälklag";
  }

  if (category === "IFCROOF") return "Takkonstruktion";

  if (category === "IFCBEAM" || category === "IFCCOLUMN" || category === "IFCMEMBER") {
    return "Bärande – Pelare och balkar";
  }

  if (category === "IFCWALL" || category === "IFCWALLSTANDARDCASE") {
    if (isExternal === true) return "Ytterväggar";
    if (loadBearing === true) return "Bärande väggar";
    return "Innerväggar – Icke bärande";
  }
  if (category === "IFCCURTAINWALL") return "Ytterväggar";

  if (category === "IFCSTAIR" || category === "IFCRAMP") return "Trappor och ramper";

  if (category === "IFCRAILING") {
    return isExternal === true ? "Räcken – Utomhus" : "Räcken – Inomhus";
  }

  if (category === "IFCWINDOW") return "Fönster";

  if (category === "IFCDOOR") {
    return isExternal === true ? "Ytterdörrar" : "Innerdörrar";
  }

  if (category === "IFCCOVERING") return "Undertak och ytskikt";

  if (
    category === "IFCFURNITURE" ||
    category === "IFCFURNISHINGELEMENT" ||
    category === "IFCSYSTEMFURNITUREELEMENT"
  ) {
    return "Möbler och inredning";
  }

  if (category.includes("PIPE") || category === "IFCSANITARYTERMINAL") {
    return "VVS-installationer";
  }
  if (category.includes("DUCT") || category === "IFCAIRTERMINAL") {
    return "Ventilationssystem";
  }
  if (
    category.includes("CABLE") ||
    category === "IFCELECTRICAPPLIANCE" ||
    category === "IFCLIGHTFIXTURE" ||
    category === "IFCOUTLET" ||
    category === "IFCSWITCHINGDEVICE" ||
    category === "IFCELECTRICDISTRIBUTIONBOARD"
  ) {
    return "Elinstallationer";
  }
  if (category === "IFCFIRESUPPRESSIONTERMINAL" || category === "IFCALARM") {
    return "Brand och säkerhet";
  }
  if (category === "IFCTRANSPORTELEMENT") return "Hissar och rulltrappor";
  if (category === "IFCDISCRETEACCESSORY") return "Övrig utrustning";

  return null;
}

/** Klassar automatiskt alla objekt som ännu inte har en allmän byggdels-
 *  klassning - samma icke-destruktiva princip som autoClassifyBoverketCategories. */
function autoClassifyGeneralCategories(): {
  classified: number;
  skippedExisting: number;
  unclear: number;
} {
  let classified = 0;
  let skippedExisting = 0;
  let unclear = 0;

  for (const row of quantityRows) {
    const key = itemKey(row);
    if (generalCategoryByItem.has(key)) {
      skippedExisting++;
      continue;
    }
    const suggestion = suggestGeneralCategory(row);
    if (!suggestion) {
      unclear++;
      continue;
    }
    generalCategoryByItem.set(key, suggestion);
    row.values[GENERAL_COLUMN_KEY] = suggestion;
    classified++;
  }

  if (!takeoffPanel.classList.contains("hidden")) renderTakeoffTable();
  return { classified, skippedExisting, unclear };
}

function updateTakeoffRowSelectionStyles() {
  if (!lastTakeoffTable) return;
  const keys = new Set(takeoffSelection.map(itemKey));
  for (const rowEl of takeoffTableWrapper.querySelectorAll<HTMLElement>(".takeoff-row")) {
    const index = Number(rowEl.dataset.rowIndex);
    const tableRow = lastTakeoffTable.rows[index];
    if (!tableRow) continue;
    const allSelected = tableRow.sourceRows.every((r) => keys.has(`${r.modelId}::${r.localId}`));
    rowEl.classList.toggle("selected", allSelected);
  }
}

let takeoffColumns: TakeoffColumn[] = [
  { key: "model", label: columnLabel("model"), kind: "text", grouped: false },
  { key: "_category", label: columnLabel("_category"), kind: "text", grouped: false },
  { key: "Name", label: columnLabel("Name"), kind: "text", grouped: false },
  { key: "sideArea", label: columnLabel("sideArea"), kind: "number", grouped: false },
  { key: "volume", label: columnLabel("volume"), kind: "number", grouped: false },
  {
    key: BOVERKET_COLUMN_KEY,
    label: columnLabel(BOVERKET_COLUMN_KEY),
    kind: "select",
    grouped: false,
  },
];

function quantityRowToSelectable(row: QuantityRow): SelectableItem {
  return {
    modelId: row.modelId,
    localId: row.localId,
    metrics: {
      sideArea: Number(row.values.sideArea) || 0,
      volume: Number(row.values.volume) || 0,
      thickness: Number(row.values.thickness) || 0,
      height: Number(row.values.height) || 0,
      width: Number(row.values.width) || 0,
      length: Number(row.values.length) || 0,
      perimeter: Number(row.values.perimeter) || 0,
    },
  };
}

async function computeQuantityTakeoff() {
  quantityRows = [];

  const modelIds = [...fragments.list.keys()];
  const modelItemIds = new Map<string, number[]>();
  let total = 0;

  for (const modelId of modelIds) {
    const model = fragments.list.get(modelId);
    if (!model) continue;
    const ids = await model.getItemsIdsWithGeometry();
    modelItemIds.set(modelId, ids);
    total += ids.length;
  }

  let done = 0;
  takeoffStatus.textContent = `Beräknar mängder... 0/${total}`;
  if (total > 0) setProgress(`Beräknar mängder... 0/${total}`, 0);

  const CHUNK_SIZE = 25;
  const discoveredAttributes = new Set<string>();
  const discoveredPsetColumns = new Set<string>();
  numericPsetColumns.clear();

  for (const modelId of modelIds) {
    const model = fragments.list.get(modelId);
    const ids = modelItemIds.get(modelId);
    if (!model || !ids || ids.length === 0) continue;

    const dataList = await model.getItemsData(ids, {
      attributesDefault: true,
      relations: {
        IsTypedBy: { attributes: true, relations: true },
        HasAssociations: { attributes: true, relations: true },
        IsDefinedBy: { attributes: true, relations: true },
      },
    });
    const modelLabel = modelNames.get(modelId) ?? modelId;

    for (let i = 0; i < ids.length; i += CHUNK_SIZE) {
      const chunkIds = ids.slice(i, i + CHUNK_SIZE);
      const chunkData = dataList.slice(i, i + CHUNK_SIZE);

      const chunkMetrics = await Promise.all(
        chunkIds.map((localId) => computeMetrics(model, localId)),
      );

      for (let j = 0; j < chunkIds.length; j++) {
        const data = chunkData[j];
        const metrics = chunkMetrics[j];
        const values: Record<string, string | number> = { model: modelLabel };

        for (const [key, value] of Object.entries(data)) {
          if (key === "_localId" || Array.isArray(value) || !isAttribute(value)) continue;
          values[key] = typeof value.value === "number" ? value.value : String(value.value ?? "-");
          discoveredAttributes.add(key);
        }

        // Samma Typ-/Material-/Klassificeringsdata som egenskapspanelen (se
        // renderRelationsSections), sammanfattad till en textrad per kolumn.
        const typeItem = relationItems(data, "IsTypedBy")[0];
        values[TYPE_COLUMN_KEY] = typeItem ? (itemName(typeItem) ?? "-") : "-";

        const materialNames: string[] = [];
        const classificationNames: string[] = [];
        for (const association of relationItems(data, "HasAssociations")) {
          const category = itemCategory(association);
          if (MATERIAL_CATEGORIES.has(category)) materialNames.push(summarizeMaterial(association));
          else if (CLASSIFICATION_CATEGORIES.has(category)) classificationNames.push(summarizeClassification(association));
        }
        values[MATERIAL_COLUMN_KEY] = materialNames.length > 0 ? materialNames.join(", ") : "-";
        values[CLASSIFICATION_COLUMN_KEY] =
          classificationNames.length > 0 ? classificationNames.join(", ") : "-";

        // Alla property set- och quantity set-egenskaper (se psetPropertyEntries),
        // så vilken pset-kolumn som helst kan läggas till precis som Attribut.
        for (const entry of psetPropertyEntries(data)) {
          discoveredPsetColumns.add(entry.key);
          if (typeof entry.value === "number") {
            values[entry.key] = entry.value;
            numericPsetColumns.add(entry.key);
          } else {
            values[entry.key] = String(entry.value ?? "-");
          }
        }

        if (metrics) {
          values.sideArea = Number(metrics.sideArea.toFixed(2));
          values.volume = Number(metrics.volume.toFixed(2));
          values.thickness = Number(metrics.thickness.toFixed(2));
          values.height = Number(metrics.height.toFixed(2));
          values.width = Number(metrics.width.toFixed(2));
          values.length = Number(metrics.length.toFixed(2));
          values.perimeter = Number(metrics.perimeter.toFixed(2));
        }

        const localId = chunkIds[j];
        values[BOVERKET_COLUMN_KEY] = boverketCategoryByItem.get(itemKey({ modelId, localId })) ?? "";
        values[GENERAL_COLUMN_KEY] = generalCategoryByItem.get(itemKey({ modelId, localId })) ?? "";

        quantityRows.push({ modelId, localId, values });
      }

      done += chunkIds.length;
      takeoffStatus.textContent = `Beräknar mängder... ${done}/${total}`;
      setProgress(`Beräknar mängder... ${done}/${total}`, (done / total) * 100);
    }
  }

  availableAttributeColumns = [...discoveredAttributes].sort();
  availablePsetColumns = [...discoveredPsetColumns].sort();

  takeoffStatus.textContent = `${quantityRows.length} objekt beräknade, letar dubbletter...`;
  setProgress("Letar dubbletter...");
  await runDuplicateScan();
  const duplicateKeys = new Set(
    duplicateGroups.flatMap((group) => group.items.map((item) => itemKey(item))),
  );
  for (const row of quantityRows) {
    row.values[DUPLICATE_COLUMN_KEY] = duplicateKeys.has(itemKey(row)) ? "Ja" : "Nej";
  }

  takeoffStatus.textContent = `${quantityRows.length} objekt beräknade`;
  clearProgress();
  renderColumnEditor();
}

function formatCell(
  value: string | number | undefined,
  kind: "text" | "number" | "select",
): string | number {
  if (value === undefined || value === "") return "-";
  if (kind === "number" && typeof value === "number") return Number(value.toFixed(2));
  return String(value);
}

function buildTakeoffTable(): { headers: string[]; rows: TakeoffTableRow[] } {
  const groupedColumns = takeoffColumns.filter((c) => c.grouped);
  // Boverket- och Byggdels-kolumnerna renderas som redigerbara selects istället
  // för vanliga textceller (se renderTakeoffTable/renderBoverketCell/
  // renderGeneralCell), utom när de själva används som grupperingsnyckel - då
  // är värdet per definition enhetligt inom gruppen och kan visas som en
  // vanlig statisk cell.
  const displayColumns = takeoffColumns.filter(
    (c) => (c.key !== BOVERKET_COLUMN_KEY && c.key !== GENERAL_COLUMN_KEY) || c.grouped,
  );

  if (groupedColumns.length === 0) {
    const headers = displayColumns.map((c) => c.label);
    const rows: TakeoffTableRow[] = quantityRows.map((row) => ({
      cells: displayColumns.map((c) => formatCell(row.values[c.key], c.kind)),
      sourceRows: [row],
    }));
    return { headers, rows };
  }

  const numericColumns = displayColumns.filter((c) => !c.grouped && c.kind === "number");
  const groups = new Map<string, QuantityRow[]>();

  for (const row of quantityRows) {
    const key = groupedColumns.map((c) => String(row.values[c.key] ?? "-")).join("␟");
    const group = groups.get(key) ?? [];
    group.push(row);
    groups.set(key, group);
  }

  const headers = [
    ...groupedColumns.map((c) => c.label),
    "Antal",
    ...numericColumns.map((c) => c.label),
  ];

  const rows: TakeoffTableRow[] = [...groups.values()]
    .map((groupRows) => {
      const first = groupRows[0];
      const keyCells = groupedColumns.map((c) => formatCell(first.values[c.key], c.kind));
      const sums = numericColumns.map((c) => {
        const sum = groupRows.reduce((acc, r) => acc + (Number(r.values[c.key]) || 0), 0);
        return Number(sum.toFixed(2));
      });
      return {
        cells: [...keyCells, groupRows.length, ...sums],
        sourceRows: groupRows,
        groupLabel: keyCells.join(" / "),
      };
    })
    .sort((a, b) => String(a.cells[0]).localeCompare(String(b.cells[0])));

  return { headers, rows };
}

function isTakeoffRowChecked(sourceRows: QuantityRow[]): boolean {
  return sourceRows.every((r) =>
    takeoffCheckedItems.has(itemKey({ modelId: r.modelId, localId: r.localId })),
  );
}

function renderTakeoffTable() {
  if (quantityRows.length === 0) {
    takeoffTableWrapper.innerHTML = '<div class="model-empty">Inga mängder beräknade än</div>';
    lastTakeoffTable = null;
    return;
  }

  const table = buildTakeoffTable();
  lastTakeoffTable = table;
  const isGrouped = takeoffColumns.some((c) => c.grouped);

  // Boverket-/Byggdels-kolumnerna renderas som redigerbara selects (se
  // renderBoverketCell/renderGeneralCell) och är därför inte med i
  // buildTakeoffTable()s vanliga celler - de splicas in separat här. Ogrupperad
  // vy visar alla kolumner 1:1, så de ska hamna på sin egen plats i
  // kolumnordningen; grupperad vy har redan en fast form (nyckelkolumner,
  // Antal, summerade kolumner) där de alltid hamnar sist. Om båda är aktiva
  // samtidigt måste de splicas in i STIGANDE indexordning (den som står
  // tidigast i takeoffColumns först) - annars hamnar den andras insticksindex
  // (beräknat mot den FULLA kolumnlistan) fel efter att den första redan
  // flyttat om cellsHtml.
  const specialColumns = [
    isBoverketColumnVisible()
      ? { key: BOVERKET_COLUMN_KEY, label: BOVERKET_COLUMN_LABEL, headerShort: BOVERKET_COLUMN_HEADER_SHORT, render: renderBoverketCell }
      : null,
    isGeneralColumnVisible()
      ? { key: GENERAL_COLUMN_KEY, label: GENERAL_COLUMN_LABEL, headerShort: GENERAL_COLUMN_HEADER_SHORT, render: renderGeneralCell }
      : null,
  ]
    .filter((c) => c !== null)
    .map((c) => ({
      ...c,
      insertIndex: isGrouped ? -1 : takeoffColumns.findIndex((col) => col.key === c.key),
    }))
    .sort((a, b) => a.insertIndex - b.insertIndex);

  const numberOfTrailingNumeric = table.rows[0]
    ? table.rows[0].cells.filter((c) => typeof c === "number").length
    : 0;

  const bodyRows = table.rows
    .map((row, index) => {
      const cellsHtml = row.cells.map((cell, cellIndex) => {
        const isNumeric = cellIndex >= row.cells.length - numberOfTrailingNumeric;
        return `<td${isNumeric ? ' style="text-align:right"' : ""}>${escapeHtml(String(cell))}</td>`;
      });
      for (const col of specialColumns) {
        const cellHtml = col.render(index, row.sourceRows);
        if (col.insertIndex >= 0) cellsHtml.splice(col.insertIndex, 0, cellHtml);
        else cellsHtml.push(cellHtml);
      }
      const checked = isTakeoffRowChecked(row.sourceRows) ? " checked" : "";
      const checkboxCell = `<td><input type="checkbox" class="takeoff-row-checkbox" data-row-index="${index}"${checked} /></td>`;
      return `<tr class="takeoff-row" data-row-index="${index}">${checkboxCell}${cellsHtml.join("")}</tr>`;
    })
    .join("");

  const headers = [...table.headers];
  for (const col of specialColumns) {
    if (col.insertIndex >= 0) headers.splice(col.insertIndex, 0, col.label);
    else headers.push(col.label);
  }

  const headerCellsHtml = headers
    .map((h) => {
      const special = specialColumns.find((c) => c.label === h);
      return special
        ? `<th class="takeoff-col-narrow" title="${escapeHtml(special.label)}">${escapeHtml(special.headerShort)}</th>`
        : `<th>${escapeHtml(h)}</th>`;
    })
    .join("");

  takeoffTableWrapper.innerHTML = `
    <table class="takeoff-table">
      <thead><tr><th><input type="checkbox" id="takeoff-select-all" /></th>${headerCellsHtml}</tr></thead>
      <tbody>${bodyRows}</tbody>
    </table>
  `;
  takeoffRangeAnchor = null;
  updateTakeoffRowSelectionStyles();
  updateTakeoffBulkBar();
}

function renderColumnEditor() {
  takeoffColumnsList.innerHTML = takeoffColumns
    .map(
      (col, index) => `
        <div class="takeoff-column-chip" data-index="${index}">
          <button
            class="takeoff-column-move"
            data-dir="left"
            title="Flytta kolumnen åt vänster"
            ${index === 0 ? "disabled" : ""}
          >◀</button>
          <label class="takeoff-column-check">
            <input type="checkbox" class="takeoff-column-group-toggle" ${col.grouped ? "checked" : ""} />
            Gruppera
          </label>
          <span class="takeoff-column-label">${escapeHtml(col.label)}</span>
          <button
            class="takeoff-column-move"
            data-dir="right"
            title="Flytta kolumnen åt höger"
            ${index === takeoffColumns.length - 1 ? "disabled" : ""}
          >▶</button>
          <button class="takeoff-column-remove" title="Ta bort kolumn">✕</button>
        </div>
      `,
    )
    .join("");

  const activeKeys = new Set(takeoffColumns.map((c) => c.key));
  const modelOption = activeKeys.has("model") ? "" : `<option value="model">Modell</option>`;
  const boverketOption = activeKeys.has(BOVERKET_COLUMN_KEY)
    ? ""
    : `<option value="${BOVERKET_COLUMN_KEY}">${escapeHtml(BOVERKET_COLUMN_LABEL)}</option>`;
  const generalOption = activeKeys.has(GENERAL_COLUMN_KEY)
    ? ""
    : `<option value="${GENERAL_COLUMN_KEY}">${escapeHtml(GENERAL_COLUMN_LABEL)}</option>`;
  const duplicateOption = activeKeys.has(DUPLICATE_COLUMN_KEY)
    ? ""
    : `<option value="${DUPLICATE_COLUMN_KEY}">${escapeHtml(DUPLICATE_COLUMN_LABEL)}</option>`;
  const relationOptions = [
    [TYPE_COLUMN_KEY, TYPE_COLUMN_LABEL],
    [MATERIAL_COLUMN_KEY, MATERIAL_COLUMN_LABEL],
    [CLASSIFICATION_COLUMN_KEY, CLASSIFICATION_COLUMN_LABEL],
  ]
    .filter(([key]) => !activeKeys.has(key))
    .map(([key, label]) => `<option value="${key}">${escapeHtml(label)}</option>`)
    .join("");
  const metricOptions = METRIC_COLUMN_DEFS.filter((c) => !activeKeys.has(c.key))
    .map((c) => `<option value="${escapeHtml(c.key)}">${escapeHtml(c.label)}</option>`)
    .join("");
  const attributeOptions = availableAttributeColumns
    .filter((key) => !activeKeys.has(key))
    .map((key) => `<option value="${escapeHtml(key)}">${escapeHtml(columnLabel(key))}</option>`)
    .join("");

  // Grupperade per property set/quantity set (t.ex. "Pset_BeamCommon"), inte
  // en enda platt lista - annars blir listan snabbt oöverskådlig när många
  // psets med flera egenskaper var upptäcks över alla inlästa modeller.
  const psetGroups = new Map<string, string[]>();
  for (const key of availablePsetColumns) {
    if (activeKeys.has(key)) continue;
    const parsed = parsePsetColumnKey(key);
    if (!parsed) continue;
    const options = psetGroups.get(parsed.psetName) ?? [];
    options.push(`<option value="${escapeHtml(key)}">${escapeHtml(parsed.propName)}</option>`);
    psetGroups.set(parsed.psetName, options);
  }
  const psetOptgroups = [...psetGroups.entries()]
    .sort(([a], [b]) => a.localeCompare(b))
    .map(([psetName, options]) => `<optgroup label="${escapeHtml(psetName)}">${options.join("")}</optgroup>`)
    .join("");

  if (
    !modelOption &&
    !boverketOption &&
    !generalOption &&
    !duplicateOption &&
    !relationOptions &&
    !metricOptions &&
    !attributeOptions &&
    !psetOptgroups
  ) {
    takeoffAddColumnSelect.innerHTML = '<option value="">Inga fler kolumner</option>';
    takeoffAddColumnSelect.disabled = true;
    return;
  }

  takeoffAddColumnSelect.disabled = false;
  takeoffAddColumnSelect.innerHTML = `
    ${modelOption ? `<optgroup label="Övrigt">${modelOption}</optgroup>` : ""}
    ${metricOptions ? `<optgroup label="Beräknat">${metricOptions}</optgroup>` : ""}
    ${relationOptions ? `<optgroup label="Typ/material">${relationOptions}</optgroup>` : ""}
    ${boverketOption ? `<optgroup label="Klimatdeklaration">${boverketOption}</optgroup>` : ""}
    ${generalOption ? `<optgroup label="Byggdelsklassificering">${generalOption}</optgroup>` : ""}
    ${duplicateOption ? `<optgroup label="Dubbletter">${duplicateOption}</optgroup>` : ""}
    ${attributeOptions ? `<optgroup label="Attribut">${attributeOptions}</optgroup>` : ""}
    ${psetOptgroups}
  `;
}

async function exportQuantitiesToExcel() {
  if (quantityRows.length === 0) return;

  // Byter tillfälligt ut quantityRows mot en filtrerad kopia innan
  // buildTakeoffTable() körs, så grupperingen/summeringen (som annars är
  // ganska invecklad) återanvänds precis som den redan fungerar för den
  // vanliga tabellvyn - påverkar bara exporten, inte det som visas i appen.
  // Ingen await mellan bytet och återställningen, så inget annat hinner läsa
  // det filtrerade tillståndet under tiden.
  const excludeUnclassified = takeoffExportExcludeUnclassified.checked;
  const originalRows = quantityRows;
  if (excludeUnclassified) {
    quantityRows = quantityRows.filter((row) => {
      const value = row.values[BOVERKET_COLUMN_KEY];
      return value !== "" && value !== "Ingår ej";
    });
  }
  const table = buildTakeoffTable();
  quantityRows = originalRows;

  if (table.rows.length === 0) {
    takeoffStatus.textContent = "Inget att exportera - alla rader är ej klassade/Ingår ej";
    return;
  }

  const showBoverketColumn = isBoverketColumnVisible();
  const headers = [...table.headers, ...(showBoverketColumn ? [BOVERKET_COLUMN_LABEL] : [])];
  const rows = table.rows.map((r) => {
    if (!showBoverketColumn) return r.cells;
    const value = resolveGroupBoverketValue(r.sourceRows);
    return [...r.cells, value === "mixed" ? "(Blandat)" : value || "-"];
  });
  await writeExcelFile([headers, ...rows]).toFile("mangdavtagning.xlsx");
}

/** Mängdavtagningen dockar längs hela nedkanten - när den är öppen krymps de
 *  övre panelerna (se CSS-regeln för body.takeoff-open) så de aldrig hamnar
 *  bakom/under den istället för att bara skära av dem. */
function setTakeoffOpen(open: boolean) {
  takeoffPanel.classList.toggle("hidden", !open);
  document.body.classList.toggle("takeoff-open", open);
  takeoffToggle.classList.toggle("active", open);
}

takeoffToggle.addEventListener("click", async () => {
  const isHidden = takeoffPanel.classList.contains("hidden");

  if (!isHidden) {
    setTakeoffOpen(false);
    return;
  }

  setTakeoffOpen(true);
  renderColumnEditor();
  if (quantityRows.length === 0) {
    await computeQuantityTakeoff();
  }
  renderTakeoffTable();
});

takeoffClose.addEventListener("click", () => {
  setTakeoffOpen(false);
});

takeoffRefresh.addEventListener("click", async () => {
  await computeQuantityTakeoff();
  renderTakeoffTable();
});

takeoffAutoClassify.addEventListener("click", () => {
  if (quantityRows.length === 0) return;
  const boverket = autoClassifyBoverketCategories();
  const general = autoClassifyGeneralCategories();
  const summarize = (label: string, r: { classified: number; skippedExisting: number; unclear: number }) => {
    const parts = [`${r.classified} föreslagna`];
    if (r.skippedExisting > 0) parts.push(`${r.skippedExisting} redan klassade`);
    if (r.unclear > 0) parts.push(`${r.unclear} osäkra`);
    return `${label}: ${parts.join(", ")}`;
  };
  takeoffStatus.textContent = `Klassningsförslag - ${summarize("Boverket", boverket)}. ${summarize("Byggdel", general)}.`;
});

takeoffExport.addEventListener("click", () => {
  void exportQuantitiesToExcel();
});

function exportToKlimatanalys(): void {
  const table = buildTakeoffTable();
  if (!table || table.rows.length === 0) {
    takeoffStatus.textContent = "Inga beräknade mängder att exportera";
    return;
  }
  const headers = [...table.headers];
  const rows = table.rows.map((r) => [...r.cells] as (string | number)[]);
  window.parent.postMessage({ type: "ifcinfinity-export", rows: [headers, ...rows] }, "*");
  takeoffStatus.textContent = "Exporterat till Klimatanalys ✓";
}

takeoffExportKlimat.addEventListener("click", exportToKlimatanalys);

takeoffColumnsList.addEventListener("click", (event) => {
  const target = event.target as HTMLElement;
  const chip = target.closest(".takeoff-column-chip") as HTMLElement | null;
  if (!chip) return;
  const index = Number(chip.dataset.index);

  if (target.classList.contains("takeoff-column-remove")) {
    takeoffColumns.splice(index, 1);
    renderColumnEditor();
    renderTakeoffTable();
    return;
  }

  if (target.classList.contains("takeoff-column-move")) {
    const swapWith = target.dataset.dir === "left" ? index - 1 : index + 1;
    if (swapWith < 0 || swapWith >= takeoffColumns.length) return;
    [takeoffColumns[index], takeoffColumns[swapWith]] = [takeoffColumns[swapWith], takeoffColumns[index]];
    renderColumnEditor();
    renderTakeoffTable();
  }
});

takeoffColumnsList.addEventListener("change", (event) => {
  const target = event.target as HTMLElement;
  if (!target.classList.contains("takeoff-column-group-toggle")) return;
  const chip = target.closest(".takeoff-column-chip") as HTMLElement | null;
  if (!chip) return;
  const index = Number(chip.dataset.index);
  takeoffColumns[index].grouped = (target as HTMLInputElement).checked;
  renderTakeoffTable();
});

takeoffAddColumnBtn.addEventListener("click", () => {
  const key = takeoffAddColumnSelect.value;
  if (!key) return;
  takeoffColumns.push({ key, label: columnLabel(key), kind: columnKind(key), grouped: false });
  renderColumnEditor();
  renderTakeoffTable();
});

takeoffTableWrapper.addEventListener("change", (event) => {
  const target = event.target as HTMLElement;

  if (target.id === "takeoff-select-all" && lastTakeoffTable) {
    const checked = (target as HTMLInputElement).checked;
    for (const tableRow of lastTakeoffTable.rows) {
      for (const r of tableRow.sourceRows) {
        const key = itemKey({ modelId: r.modelId, localId: r.localId });
        if (checked) takeoffCheckedItems.add(key);
        else takeoffCheckedItems.delete(key);
      }
    }
    renderTakeoffTable();
    return;
  }

  if (target.classList.contains("takeoff-row-checkbox") && lastTakeoffTable) {
    const index = Number((target as HTMLInputElement).dataset.rowIndex);
    const tableRow = lastTakeoffTable.rows[index];
    if (!tableRow) return;
    const checked = (target as HTMLInputElement).checked;
    for (const r of tableRow.sourceRows) {
      const key = itemKey({ modelId: r.modelId, localId: r.localId });
      if (checked) takeoffCheckedItems.add(key);
      else takeoffCheckedItems.delete(key);
    }
    updateTakeoffBulkBar();
    return;
  }

  if (!lastTakeoffTable) return;
  const isBoverketSelect = target.classList.contains("boverket-select");
  const isGeneralSelect = target.classList.contains("general-select");
  if (!isBoverketSelect && !isGeneralSelect) return;

  const index = Number((target as HTMLSelectElement).dataset.rowIndex);
  const tableRow = lastTakeoffTable.rows[index];
  if (!tableRow) return;

  const value = (target as HTMLSelectElement).value;
  if (isBoverketSelect) applyBoverketCategoryToItems(tableRow.sourceRows, value);
  else applyGeneralCategoryToItems(tableRow.sourceRows, value);
});

takeoffBulkExclude.addEventListener("click", () => {
  const items = getCheckedTakeoffItems();
  if (items.length === 0) return;
  takeoffCheckedItems.clear();
  applyBoverketCategoryToItems(items, "Ingår ej");
});

takeoffBulkClear.addEventListener("click", () => {
  takeoffCheckedItems.clear();
  renderTakeoffTable();
});

takeoffTableWrapper.addEventListener("click", (event) => {
  const target = event.target as HTMLElement;
  if (target.tagName === "SELECT" || target.tagName === "OPTION" || target.tagName === "INPUT") return;
  const row = target.closest(".takeoff-row") as HTMLElement | null;
  if (!row || !lastTakeoffTable) return;

  const index = Number(row.dataset.rowIndex);
  const tableRow = lastTakeoffTable.rows[index];
  if (!tableRow) return;

  const clickedItems = tableRow.sourceRows.map(quantityRowToSelectable);
  const isToggle = event.ctrlKey || event.metaKey;
  let items: SelectableItem[];
  let label: string | undefined;

  if (event.shiftKey && takeoffRangeAnchor !== null) {
    const [start, end] =
      takeoffRangeAnchor < index ? [takeoffRangeAnchor, index] : [index, takeoffRangeAnchor];
    items = [];
    for (let i = start; i <= end; i++) {
      items.push(...lastTakeoffTable.rows[i].sourceRows.map(quantityRowToSelectable));
    }
  } else {
    const mode = isToggle ? "toggle" : "replace";
    items = resolveMultiSelect(takeoffSelection, clickedItems, mode);
    takeoffRangeAnchor = index;
    if (mode === "replace") label = tableRow.groupLabel;
  }

  takeoffSelection = items;
  updateTakeoffRowSelectionStyles();
  void applySelection(items, label);
});

// Skapa egen modell
const footprintEditor = new FootprintEditor(footprintCanvas);

function renderFootprintInfo() {
  if (footprintEditor.isClosed()) {
    footprintInfo.textContent = `Form klar: ${footprintEditor.getArea().toFixed(1)} m²`;
    return;
  }
  const count = footprintEditor.pointCount();
  footprintInfo.textContent =
    count === 0
      ? "Klicka för att rita byggnadens utbredning (minst 3 punkter)"
      : `${count} punkt${count === 1 ? "" : "er"} — klicka nära startpunkten för att stänga formen`;
}

/** Lägger ett redigerbart måttfält (m) direkt ovanpå mittpunkten av varje
 *  ritad sida i figuren, så att en handritad form kan finjusteras till exakta
 *  mått utan en separat lista. Sidan som stänger formen (sista → första
 *  punkten) visas skrivskyddad eftersom dess längd är en följd av att formen
 *  sluts - se FootprintEditor.setEdgeLength. */
function renderFootprintEdges() {
  const count = footprintEditor.getEdgeCount();
  if (count === 0) {
    footprintEdgeOverlay.innerHTML = "";
    return;
  }
  const inputs: string[] = [];
  for (let i = 0; i < count; i++) {
    const isClosing = footprintEditor.isClosingEdge(i);
    const length = footprintEditor.getEdgeLength(i);
    const { x, y } = footprintEditor.getEdgeMidpointScreen(i);
    const title = isClosing
      ? `Sida ${i + 1} (stängande) - längden bestäms automatiskt av att formen sluts`
      : `Sida ${i + 1} - ange längd i meter`;
    inputs.push(`
      <input
        type="number"
        min="0.1"
        step="0.1"
        value="${length.toFixed(2)}"
        data-edge-index="${i}"
        title="${escapeHtml(title)}"
        style="left:${x}px; top:${y}px;"
        ${isClosing ? "disabled" : ""}
      />
    `);
  }
  footprintEdgeOverlay.innerHTML = inputs.join("");
}

function renderFootprintZoom() {
  const zoom = footprintEditor.getZoom();
  footprintZoomLevel.textContent = `${Math.round(zoom * 100)}%`;
  footprintZoomOut.disabled = zoom <= 1;
  footprintZoomIn.disabled = zoom >= 4;
}

function updateFootprintUI() {
  renderFootprintInfo();
  renderFootprintEdges();
  renderFootprintZoom();
}

footprintEditor.onChange = updateFootprintUI;
renderFootprintZoom();

footprintEdgeOverlay.addEventListener("change", (event) => {
  const target = event.target as HTMLElement;
  if (target.tagName !== "INPUT") return;
  const input = target as HTMLInputElement;
  const index = Number(input.dataset.edgeIndex);
  const newLength = Number(input.value);
  if (!footprintEditor.setEdgeLength(index, newLength)) {
    renderFootprintEdges();
  }
});

footprintUndo.addEventListener("click", () => {
  footprintEditor.undo();
});

footprintClear.addEventListener("click", () => {
  footprintEditor.clear();
});

footprintZoomIn.addEventListener("click", () => {
  footprintEditor.zoomIn();
});

footprintZoomOut.addEventListener("click", () => {
  footprintEditor.zoomOut();
});

generateToggle.addEventListener("click", () => {
  generatePanel.classList.remove("hidden");
  generateToggle.classList.add("active");
});

generateClose.addEventListener("click", () => {
  generatePanel.classList.add("hidden");
  generateToggle.classList.remove("active");
});

generateSubmit.addEventListener("click", async () => {
  const footprint = footprintEditor.getFootprint();
  const floors = Math.round(Number(generateFloorsInput.value));
  const floorHeight = Number(generateHeightInput.value);
  const windowRatio = Number(generateWindowRatioInput.value);

  if (!footprint) {
    generateStatus.textContent =
      "Rita en form med minst 3 punkter och stäng den (klicka nära startpunkten) först.";
    return;
  }
  if (!Number.isFinite(floors) || floors <= 0) {
    generateStatus.textContent = "Ange ett giltigt antal våningsplan.";
    return;
  }
  if (!Number.isFinite(floorHeight) || floorHeight <= 0) {
    generateStatus.textContent = "Ange en giltig våningshöjd (m).";
    return;
  }
  if (!Number.isFinite(windowRatio) || windowRatio < 0 || windowRatio > 100) {
    generateStatus.textContent = "Ange en giltig fönsterandel (0-100%).";
    return;
  }

  generateStatus.textContent = "Genererar modell...";

  const area = footprintEditor.getArea();
  const ifcText = generateBuildingIfc({ footprint, floors, floorHeight, windowRatio });
  const buffer = new TextEncoder().encode(ifcText);
  const modelId = `model-${modelCount++}`;
  const name = `Genererad byggnad (${floors} plan, ${area.toFixed(0)} m²)`;
  modelNames.set(modelId, name);
  modelIfcText.set(modelId, ifcText);

  setProgress("Genererar modell...");
  try {
    // coordinate: true - se motsvarande kommentar i loadOrReplaceSyncedFile.
    await ifcLoader.load(buffer, true, modelId, {});
    renderModelTreeIfOpen();
    generateStatus.textContent = `${name} skapad`;
    generatePanel.classList.add("hidden");
    generateToggle.classList.remove("active");
  } catch (error) {
    modelNames.delete(modelId);
    modelIfcText.delete(modelId);
    generateStatus.textContent = "Kunde inte generera modellen.";
    console.error(error);
  } finally {
    clearProgress();
  }
});

// Re-render after viewport resize so the scene fills the new canvas size.
window.addEventListener("resize", () => { postproductionRenderer.needsUpdate = true; });
