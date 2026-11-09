// Flyttar en hel IFC-modell i höjdled (IFC:s Z-axel) genom att skriva om
// rå STEP-text. Alla produkters placeringar bildar ett träd av
// IfcLocalPlacement där bara rötterna (PlacementRelTo = $) är placerade
// direkt i världskoordinater - genom att förskjuta just rötternas origo
// följer hela modellen med, utan att någon geometri behöver röras.
//
// Rötternas befintliga IfcCartesianPoint/IfcAxis2Placement3D ändras INTE på
// plats, eftersom samma punkt (t.ex. origo #x=(0.,0.,0.)) ofta delas med
// annan geometri. Istället skapas nya entiteter som bara rotplaceringarna
// pekar om till.

export interface ElevationShiftResult {
  text: string;
  /** Antal rotplaceringar (IfcLocalPlacement utan PlacementRelTo) som flyttades. */
  movedPlacements: number;
  /** Antal IfcBuildingStorey vars Elevation-attribut justerades. */
  movedStoreys: number;
  /** Meter per längdenhet i filen (t.ex. 0.001 för millimeter). */
  unitScale: number;
}

interface EntityRef {
  /** Index för "#id" i texten. */
  start: number;
  /** Index för entitetsnamnet (t.ex. "IFCLOCALPLACEMENT"). */
  typeStart: number;
  type: string;
}

/** Plockar ut argumenten mellan den balanserade parentesen som börjar vid
 *  `openIdx` (som måste peka på "("), samt index för avslutande ";". */
function readArgs(text: string, openIdx: number): { args: string; end: number } | null {
  let depth = 0;
  let inString = false;
  for (let i = openIdx; i < text.length; i++) {
    const ch = text[i];
    if (ch === "'") {
      inString = !inString;
      continue;
    }
    if (inString) continue;
    if (ch === "(") depth++;
    else if (ch === ")") {
      depth--;
      if (depth === 0) {
        const semi = text.indexOf(";", i);
        return { args: text.slice(openIdx + 1, i), end: semi === -1 ? i + 1 : semi + 1 };
      }
    }
  }
  return null;
}

/** Delar en argumentlista på toppnivå-kommatecken (ignorerar nästlade listor/strängar). */
function splitArgs(args: string): string[] {
  const parts: string[] = [];
  let depth = 0;
  let inString = false;
  let current = "";
  for (const ch of args) {
    if (ch === "'") inString = !inString;
    if (!inString) {
      if (ch === "(") depth++;
      else if (ch === ")") depth--;
      else if (ch === "," && depth === 0) {
        parts.push(current.trim());
        current = "";
        continue;
      }
    }
    current += ch;
  }
  if (current.trim() !== "" || parts.length > 0) parts.push(current.trim());
  return parts;
}

function refId(value: string | undefined): number | null {
  const match = value?.match(/^#(\d+)$/);
  return match ? Number(match[1]) : null;
}

/** STEP-reella tal måste innehålla en decimalpunkt ("12." / "1.5E-7"). */
function stepReal(value: number): string {
  let s = Number(value.toPrecision(15)).toString().toUpperCase();
  if (!s.includes(".")) {
    const e = s.indexOf("E");
    s = e === -1 ? `${s}.` : `${s.slice(0, e)}.${s.slice(e)}`;
  }
  return s;
}

const SI_PREFIX_SCALE: Record<string, number> = {
  KILO: 1e3,
  HECTO: 1e2,
  DECA: 1e1,
  DECI: 1e-1,
  CENTI: 1e-2,
  MILLI: 1e-3,
  MICRO: 1e-6,
};

const CONVERSION_UNIT_SCALE: Record<string, number> = {
  FOOT: 0.3048,
  "US SURVEY FOOT": 1200 / 3937,
  INCH: 0.0254,
  YARD: 0.9144,
  MILE: 1609.344,
};

function indexEntities(text: string): { entities: Map<number, EntityRef>; maxId: number } {
  const entities = new Map<number, EntityRef>();
  let maxId = 0;
  const re = /#(\d+)\s*=\s*([A-Za-z0-9_]+)\s*\(/g;
  let match: RegExpExecArray | null;
  while ((match = re.exec(text))) {
    const id = Number(match[1]);
    entities.set(id, {
      start: match.index,
      typeStart: match.index + match[0].indexOf(match[2]),
      type: match[2].toUpperCase(),
    });
    if (id > maxId) maxId = id;
  }
  return { entities, maxId };
}

function entityArgs(
  text: string,
  entity: EntityRef | undefined,
): { parts: string[]; end: number } | null {
  if (!entity) return null;
  const open = text.indexOf("(", entity.typeStart);
  const read = readArgs(text, open);
  return read ? { parts: splitArgs(read.args), end: read.end } : null;
}

/** Meter per längdenhet enligt projektets IfcUnitAssignment (fallback: meter). */
function detectLengthUnitScale(text: string, entities: Map<number, EntityRef>): number {
  const candidates: number[] = [];
  for (const entity of entities.values()) {
    if (entity.type !== "IFCUNITASSIGNMENT") continue;
    const args = entityArgs(text, entity);
    const list = args?.parts[0]?.replace(/^\(|\)$/g, "") ?? "";
    for (const ref of list.split(",")) {
      const refNum = refId(ref.trim());
      if (refNum !== null) candidates.push(refNum);
    }
  }

  for (const id of candidates) {
    const entity = entities.get(id);
    const args = entityArgs(text, entity);
    if (!entity || !args) continue;
    if (entity.type === "IFCSIUNIT" && args.parts[1]?.toUpperCase() === ".LENGTHUNIT.") {
      const prefix = args.parts[2]?.replace(/\./g, "").toUpperCase();
      return prefix && prefix !== "$" ? (SI_PREFIX_SCALE[prefix] ?? 1) : 1;
    }
    if (
      entity.type === "IFCCONVERSIONBASEDUNIT" &&
      args.parts[1]?.toUpperCase() === ".LENGTHUNIT."
    ) {
      const name = args.parts[2]?.replace(/'/g, "").toUpperCase() ?? "";
      return CONVERSION_UNIT_SCALE[name] ?? 1;
    }
  }
  return 1;
}

/**
 * Returnerar en ny IFC-text där hela modellen är flyttad `offsetMeters`
 * meter uppåt (negativt värde = nedåt).
 */
export function shiftIfcElevation(text: string, offsetMeters: number): ElevationShiftResult {
  const { entities, maxId } = indexEntities(text);
  const unitScale = detectLengthUnitScale(text, entities);
  const dz = offsetMeters / unitScale;

  let nextId = maxId + 1;
  const newEntities: string[] = [];
  const edits: { start: number; end: number; replacement: string }[] = [];
  // Flera rotplaceringar kan dela samma axelplacering - återanvänd den nya.
  const replacedAxis = new Map<number, number>();
  let movedPlacements = 0;

  for (const [id, entity] of entities) {
    if (entity.type !== "IFCLOCALPLACEMENT") continue;
    const placement = entityArgs(text, entity);
    if (!placement || placement.parts[0] !== "$") continue;

    const axisId = refId(placement.parts[1]);
    if (axisId === null) continue;

    let newAxisId = replacedAxis.get(axisId);
    if (newAxisId === undefined) {
      const axisEntity = entities.get(axisId);
      if (axisEntity?.type !== "IFCAXIS2PLACEMENT3D") continue;
      const axis = entityArgs(text, axisEntity);
      const pointId = refId(axis?.parts[0]);
      const pointEntity = pointId !== null ? entities.get(pointId) : undefined;
      if (!axis || pointEntity?.type !== "IFCCARTESIANPOINT") continue;
      const point = entityArgs(text, pointEntity);
      const coords = point?.parts[0]?.replace(/^\(|\)$/g, "").split(",").map((c) => c.trim());
      if (!coords || coords.length < 2) continue;

      const z = coords.length >= 3 ? Number(coords[2]) : 0;
      if (!Number.isFinite(z)) continue;

      const newPointId = nextId++;
      newAxisId = nextId++;
      newEntities.push(
        `#${newPointId}=IFCCARTESIANPOINT((${coords[0]},${coords[1]},${stepReal(z + dz)}));`,
        `#${newAxisId}=IFCAXIS2PLACEMENT3D(#${newPointId},${axis.parts[1] ?? "$"},${axis.parts[2] ?? "$"});`,
      );
      replacedAxis.set(axisId, newAxisId);
    }

    edits.push({
      start: entity.start,
      end: placement.end,
      replacement: `#${id}=IFCLOCALPLACEMENT($,#${newAxisId});`,
    });
    movedPlacements++;
  }

  if (movedPlacements === 0) return { text, movedPlacements, movedStoreys: 0, unitScale };

  // Våningsplanens Elevation-attribut är ett fristående tal (inte härlett ur
  // placeringen) som många program - t.ex. IDA ICE - läser våningshöjderna
  // ifrån, så det måste flyttas lika mycket som geometrin. Index 9 är
  // Elevation i både IFC2x3 och IFC4; saknat värde ($) lämnas orört.
  let movedStoreys = 0;
  for (const [id, entity] of entities) {
    if (entity.type !== "IFCBUILDINGSTOREY") continue;
    const storey = entityArgs(text, entity);
    if (!storey || storey.parts.length < 10) continue;
    const elevation = Number(storey.parts[9]);
    if (storey.parts[9] === "$" || !Number.isFinite(elevation)) continue;
    const parts = [...storey.parts];
    parts[9] = stepReal(elevation + dz);
    edits.push({
      start: entity.start,
      end: storey.end,
      replacement: `#${id}=IFCBUILDINGSTOREY(${parts.join(",")});`,
    });
    movedStoreys++;
  }

  // Nya entiteter läggs sist i DATA-sektionen (sista ENDSEC före filslut).
  const fileEnd = text.lastIndexOf("END-ISO-10303-21");
  const dataEnd = text.lastIndexOf("ENDSEC", fileEnd === -1 ? text.length : fileEnd);
  if (dataEnd === -1) throw new Error("Hittade inte slutet på IFC-filens DATA-sektion.");
  edits.push({ start: dataEnd, end: dataEnd, replacement: `${newEntities.join("\n")}\n` });

  // Bygg ihop i en enda linjär passage - stora filer kan ha tusentals
  // rotplaceringar, och upprepad slice/konkatenering blir då kvadratisk.
  edits.sort((a, b) => a.start - b.start);
  const chunks: string[] = [];
  let cursor = 0;
  for (const edit of edits) {
    chunks.push(text.slice(cursor, edit.start), edit.replacement);
    cursor = edit.end;
  }
  chunks.push(text.slice(cursor));
  return { text: chunks.join(""), movedPlacements, movedStoreys, unitScale };
}
