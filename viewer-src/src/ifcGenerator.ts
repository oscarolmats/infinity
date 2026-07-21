// Genererar en enkel massmodell (platta byggnadskroppar) som en giltig IFC4 STEP-fil,
// utifrån area per våningsplan, antal våningsplan och våningshöjd. Filen laddas sedan
// genom samma IfcLoader-pipeline som uppladdade IFC-filer, så alla objekt blir klickbara,
// markerbara och mätbara på exakt samma sätt.

export interface FootprintPoint {
  x: number;
  y: number;
}

export interface BuildingParams {
  /** Byggnadens utbredning i planet, som en sluten polygon (minst 3 punkter, i meter). */
  footprint: FootprintPoint[];
  floors: number;
  floorHeight: number;
  /** Andel av varje väggs yta som ska täckas av fönster, 0-100 (%). */
  windowRatio: number;
}

const WINDOW_WIDTH = 1.0;
const WINDOW_HEIGHT = 1.5;
const WINDOW_AREA = WINDOW_WIDTH * WINDOW_HEIGHT;
const WINDOW_MIN_GAP = 0.3;

const GUID_CHARS =
  "0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz_$";

function generateGuid(): string {
  let result = "";
  for (let i = 0; i < 22; i++) {
    result += GUID_CHARS[Math.floor(Math.random() * GUID_CHARS.length)];
  }
  return result;
}

function real(n: number): string {
  return Number.isInteger(n) ? `${n}.` : String(n);
}

function ref(id: number | null): string {
  return id === null ? "$" : `#${id}`;
}

class StepModel {
  private lines: string[] = [];
  private counter = 0;

  next(content: string): number {
    this.counter += 1;
    this.lines.push(`#${this.counter}=${content};`);
    return this.counter;
  }

  toString(): string {
    return this.lines.join("\n");
  }
}

function point3(step: StepModel, x: number, y: number, z: number): number {
  return step.next(`IFCCARTESIANPOINT((${real(x)},${real(y)},${real(z)}))`);
}

function direction3(step: StepModel, x: number, y: number, z: number): number {
  return step.next(`IFCDIRECTION((${real(x)},${real(y)},${real(z)}))`);
}

function axis2Placement3D(
  step: StepModel,
  location: [number, number, number],
  refDirection?: [number, number, number],
): number {
  const locationId = point3(step, ...location);
  const refDirectionId = refDirection ? direction3(step, ...refDirection) : null;
  return step.next(`IFCAXIS2PLACEMENT3D(${ref(locationId)},$,${ref(refDirectionId)})`);
}

function localPlacement(
  step: StepModel,
  relativeTo: number | null,
  location: [number, number, number],
  refDirection?: [number, number, number],
): number {
  const axisId = axis2Placement3D(step, location, refDirection);
  return step.next(`IFCLOCALPLACEMENT(${ref(relativeTo)},${ref(axisId)})`);
}

function rectangleProfile(step: StepModel, xDim: number, yDim: number): number {
  const pointId = step.next(`IFCCARTESIANPOINT((0.,0.))`);
  const placementId = step.next(`IFCAXIS2PLACEMENT2D(${ref(pointId)},$)`);
  return step.next(
    `IFCRECTANGLEPROFILEDEF(.AREA.,$,${ref(placementId)},${real(xDim)},${real(yDim)})`,
  );
}

/** En godtycklig sluten polygonprofil, definierad av dess hörnpunkter i planet. */
function polygonProfile(step: StepModel, points: FootprintPoint[]): number {
  const pointIds = points.map((p) => step.next(`IFCCARTESIANPOINT((${real(p.x)},${real(p.y)}))`));
  const closedIds = [...pointIds, pointIds[0]];
  const polylineId = step.next(`IFCPOLYLINE((${closedIds.map((id) => ref(id)).join(",")}))`);
  return step.next(`IFCARBITRARYCLOSEDPROFILEDEF(.AREA.,$,${ref(polylineId)})`);
}

/** Extruderar byggnadens fotavtryck (polygonprofil) uppåt (lokal Z) med given höjd. */
function extrudedPolygonSolid(step: StepModel, points: FootprintPoint[], depth: number): number {
  const profileId = polygonProfile(step, points);
  const originId = point3(step, 0, 0, 0);
  const positionId = step.next(`IFCAXIS2PLACEMENT3D(${ref(originId)},$,$)`);
  const directionId = direction3(step, 0, 0, 1);
  return step.next(
    `IFCEXTRUDEDAREASOLID(${ref(profileId)},${ref(positionId)},${ref(directionId)},${real(depth)})`,
  );
}

/**
 * Bygger en rätblocks-solid: en rektangulär profil (xDim x yDim) centrerad på
 * (centerX, centerY), som sträcks (extruderas) uppåt (lokal Z) från startZ
 * till startZ+depth.
 */
function extrudedSolid(
  step: StepModel,
  xDim: number,
  yDim: number,
  depth: number,
  centerX = 0,
  centerY = 0,
  startZ = 0,
): number {
  const profileId = rectangleProfile(step, xDim, yDim);
  const originId = point3(step, centerX, centerY, startZ);
  const positionId = step.next(`IFCAXIS2PLACEMENT3D(${ref(originId)},$,$)`);
  const directionId = direction3(step, 0, 0, 1);
  return step.next(
    `IFCEXTRUDEDAREASOLID(${ref(profileId)},${ref(positionId)},${ref(directionId)},${real(depth)})`,
  );
}

function productShape(
  step: StepModel,
  contextId: number,
  items: number | number[],
  representationType: "SweptSolid" | "CSG" = "SweptSolid",
): number {
  const itemIds = Array.isArray(items) ? items : [items];
  const itemRefs = itemIds.map((id) => ref(id)).join(",");
  const shapeRepId = step.next(
    `IFCSHAPEREPRESENTATION(${ref(contextId)},'Body','${representationType}',(${itemRefs}))`,
  );
  return step.next(`IFCPRODUCTDEFINITIONSHAPE($,$,(${ref(shapeRepId)}))`);
}

/**
 * En rektangulär "stans"-solid som går rakt igenom väggens tjocklek (lokal
 * Y-axel), centrerad på (centerX, centerZ) i väggens egna lokala koordinater.
 * Används som andra operand i en IFCBOOLEANRESULT för att skapa ett riktigt
 * hål i väggen där ett fönster sitter.
 */
function wallPunchSolid(
  step: StepModel,
  width: number,
  height: number,
  depth: number,
  centerX: number,
  centerZ: number,
): number {
  const profileId = rectangleProfile(step, width, height);
  const locationId = point3(step, centerX, -depth / 2, centerZ);
  const axisId = direction3(step, 0, 1, 0);
  const refDirId = direction3(step, 1, 0, 0);
  const positionId = step.next(
    `IFCAXIS2PLACEMENT3D(${ref(locationId)},${ref(axisId)},${ref(refDirId)})`,
  );
  const extrudeDirId = direction3(step, 0, 0, 1);
  return step.next(
    `IFCEXTRUDEDAREASOLID(${ref(profileId)},${ref(positionId)},${ref(extrudeDirId)},${real(depth)})`,
  );
}

/** Skapar en (halvgenomskinlig) ytstil och returnerar dess id, för att ge en
 * grupp geometrier ett eget utseende (t.ex. blått glas för fönster). */
function surfaceStyle(
  step: StepModel,
  color: [number, number, number],
  transparency: number,
): number {
  const colorId = step.next(`IFCCOLOURRGB($,${real(color[0])},${real(color[1])},${real(color[2])})`);
  const renderingId = step.next(
    `IFCSURFACESTYLERENDERING(${ref(colorId)},${real(transparency)},$,$,$,$,$,$,.NOTDEFINED.)`,
  );
  return step.next(`IFCSURFACESTYLE($,.BOTH.,(${ref(renderingId)}))`);
}

function applyStyle(step: StepModel, itemId: number, styleId: number): void {
  step.next(`IFCSTYLEDITEM(${ref(itemId)},(${ref(styleId)}),$)`);
}

export function generateBuildingIfc({
  footprint,
  floors,
  floorHeight,
  windowRatio,
}: BuildingParams): string {
  const step = new StepModel();
  const wallThickness = 0.2;
  const slabThickness = 0.2;
  // Karmens och glasets djup, något mindre än väggens tjocklek så fönstret
  // sitter infällt i öppningen istället för att sticka ut ur väggen.
  const windowThickness = 0.16;
  const frameBorder = 0.06;
  const glassThickness = windowThickness * 0.6;
  const windowSillHeight = Math.max(0, (floorHeight - WINDOW_HEIGHT) / 2);

  // En vägg per kant i den ritade polygonen: väggens position är kantens
  // mittpunkt, riktningen är kantens enhetsvektor, och längden är kantens
  // egen längd (till skillnad från den tidigare kvadratiska modellen där
  // alla fyra väggar delade samma längd).
  const edges = footprint.map((p1, i) => {
    const p2 = footprint[(i + 1) % footprint.length];
    const dx = p2.x - p1.x;
    const dy = p2.y - p1.y;
    const length = Math.hypot(dx, dy);
    return {
      pos: [(p1.x + p2.x) / 2, (p1.y + p2.y) / 2, 0] as [number, number, number],
      dir: [dx / length, dy / length, 0] as [number, number, number],
      length,
    };
  });

  const personId = step.next(`IFCPERSON($,$,'Anvandare',$,$,$,$,$)`);
  const orgId = step.next(`IFCORGANIZATION($,'IFCinfinity',$,$,$)`);
  const personOrgId = step.next(`IFCPERSONANDORGANIZATION(${ref(personId)},${ref(orgId)},$)`);
  const applicationId = step.next(
    `IFCAPPLICATION(${ref(orgId)},'1.0','IFCinfinity Generator','IFCinfinity')`,
  );
  const ownerHistoryId = step.next(
    `IFCOWNERHISTORY(${ref(personOrgId)},${ref(applicationId)},$,.ADDED.,$,$,$,0)`,
  );

  const glassStyleId = surfaceStyle(step, [0.55, 0.75, 0.85], 0.35);
  const frameStyleId = surfaceStyle(step, [0.92, 0.92, 0.9], 0);

  const lengthUnitId = step.next(`IFCSIUNIT(*,.LENGTHUNIT.,$,.METRE.)`);
  const areaUnitId = step.next(`IFCSIUNIT(*,.AREAUNIT.,$,.SQUARE_METRE.)`);
  const volumeUnitId = step.next(`IFCSIUNIT(*,.VOLUMEUNIT.,$,.CUBIC_METRE.)`);
  const angleUnitId = step.next(`IFCSIUNIT(*,.PLANEANGLEUNIT.,$,.RADIAN.)`);
  const unitAssignmentId = step.next(
    `IFCUNITASSIGNMENT((${ref(lengthUnitId)},${ref(areaUnitId)},${ref(volumeUnitId)},${ref(angleUnitId)}))`,
  );

  const worldPlacementId = axis2Placement3D(step, [0, 0, 0], [1, 0, 0]);
  const geomContextId = step.next(
    `IFCGEOMETRICREPRESENTATIONCONTEXT($,'Model',3,1.0E-5,${ref(worldPlacementId)},$)`,
  );

  const projectId = step.next(
    `IFCPROJECT('${generateGuid()}',${ref(ownerHistoryId)},'Genererad byggnad',$,$,$,$,(${ref(geomContextId)}),${ref(unitAssignmentId)})`,
  );

  const sitePlacementId = localPlacement(step, null, [0, 0, 0]);
  const siteId = step.next(
    `IFCSITE('${generateGuid()}',${ref(ownerHistoryId)},'Tomt',$,$,${ref(sitePlacementId)},$,$,.ELEMENT.,$,$,$,$,$)`,
  );

  const buildingPlacementId = localPlacement(step, sitePlacementId, [0, 0, 0]);
  const buildingId = step.next(
    `IFCBUILDING('${generateGuid()}',${ref(ownerHistoryId)},'Byggnad',$,$,${ref(buildingPlacementId)},$,$,.ELEMENT.,$,$,$)`,
  );

  step.next(
    `IFCRELAGGREGATES('${generateGuid()}',${ref(ownerHistoryId)},$,$,${ref(projectId)},(${ref(siteId)}))`,
  );
  step.next(
    `IFCRELAGGREGATES('${generateGuid()}',${ref(ownerHistoryId)},$,$,${ref(siteId)},(${ref(buildingId)}))`,
  );

  const storeyIds: number[] = [];
  const elementsByStorey: number[][] = [];

  for (let i = 0; i < floors; i++) {
    const elevation = i * floorHeight;
    const storeyPlacementId = localPlacement(step, buildingPlacementId, [0, 0, elevation]);
    const storeyId = step.next(
      `IFCBUILDINGSTOREY('${generateGuid()}',${ref(ownerHistoryId)},'Plan ${i + 1}',$,$,${ref(storeyPlacementId)},$,$,.ELEMENT.,${real(elevation)})`,
    );
    storeyIds.push(storeyId);

    const elements: number[] = [];

    const slabSolidId = extrudedPolygonSolid(step, footprint, slabThickness);
    const slabShapeId = productShape(step, geomContextId, slabSolidId);
    const slabPlacementId = localPlacement(step, storeyPlacementId, [0, 0, -slabThickness]);
    const slabId = step.next(
      `IFCSLAB('${generateGuid()}',${ref(ownerHistoryId)},'Golv plan ${i + 1}',$,$,${ref(slabPlacementId)},${ref(slabShapeId)},$,.FLOOR.)`,
    );
    elements.push(slabId);

    // Hålet i väggen matchar karmens yttermått exakt, så det inte blir
    // något synligt mellanrum mellan hålkanten och karmen.
    const openingWidth = WINDOW_WIDTH;
    const openingHeight = WINDOW_HEIGHT;
    const openingDepth = wallThickness + 0.4;
    const windowCenterZ = windowSillHeight + WINDOW_HEIGHT / 2;

    for (const wallDef of edges) {
      // Antal fönster på just den här väggen, avrundat till närmaste heltal
      // utifrån önskad andel av väggens egen yta, begränsat så de får plats.
      const wallArea = wallDef.length * floorHeight;
      const targetWindowArea = wallArea * (windowRatio / 100);
      const rawWindowCount = Math.round(targetWindowArea / WINDOW_AREA);
      const maxWindowsPerWall = Math.max(
        0,
        Math.floor(wallDef.length / (WINDOW_WIDTH + WINDOW_MIN_GAP)),
      );
      const windowCount = Math.min(rawWindowCount, maxWindowsPerWall);

      const segmentWidth = wallDef.length / windowCount;
      const windowOffsets: number[] = [];
      for (let w = 0; w < windowCount; w++) {
        windowOffsets.push(-wallDef.length / 2 + segmentWidth * (w + 0.5));
      }

      // Stansa ut ett riktigt hål i väggkroppen för varje fönster (boolesk
      // subtraktion), så väggen får en verklig öppning både inifrån och utifrån.
      let wallBodyId = extrudedSolid(step, wallDef.length, wallThickness, floorHeight);
      for (const offset of windowOffsets) {
        const punchId = wallPunchSolid(
          step,
          openingWidth,
          openingHeight,
          openingDepth,
          offset,
          windowCenterZ,
        );
        wallBodyId = step.next(`IFCBOOLEANRESULT(.DIFFERENCE.,${ref(wallBodyId)},${ref(punchId)})`);
      }

      const wallShapeId = productShape(
        step,
        geomContextId,
        wallBodyId,
        windowOffsets.length > 0 ? "CSG" : "SweptSolid",
      );
      const wallPlacementId = localPlacement(step, storeyPlacementId, wallDef.pos, wallDef.dir);
      const wallId = step.next(
        `IFCWALLSTANDARDCASE('${generateGuid()}',${ref(ownerHistoryId)},'Vagg plan ${i + 1}',$,$,${ref(wallPlacementId)},${ref(wallShapeId)},$,$)`,
      );
      elements.push(wallId);

      for (const offset of windowOffsets) {
        // Fönstret byggs som en karm (fyra ramdelar) med infällt glas,
        // centrerat i väggens tjocklek så det sitter i hålet istället för
        // att sväva utanpå. Karm och glas delar inte samma yta (glaset är
        // infällt innanför karmens öppning), så ingen geometri överlappar.
        const frameTopId = extrudedSolid(
          step,
          WINDOW_WIDTH,
          windowThickness,
          frameBorder,
          0,
          0,
          WINDOW_HEIGHT - frameBorder,
        );
        const frameBottomId = extrudedSolid(step, WINDOW_WIDTH, windowThickness, frameBorder);
        const frameLeftId = extrudedSolid(
          step,
          frameBorder,
          windowThickness,
          WINDOW_HEIGHT - 2 * frameBorder,
          -WINDOW_WIDTH / 2 + frameBorder / 2,
          0,
          frameBorder,
        );
        const frameRightId = extrudedSolid(
          step,
          frameBorder,
          windowThickness,
          WINDOW_HEIGHT - 2 * frameBorder,
          WINDOW_WIDTH / 2 - frameBorder / 2,
          0,
          frameBorder,
        );
        for (const frameId of [frameTopId, frameBottomId, frameLeftId, frameRightId]) {
          applyStyle(step, frameId, frameStyleId);
        }

        const glassId = extrudedSolid(
          step,
          WINDOW_WIDTH - 2 * frameBorder,
          glassThickness,
          WINDOW_HEIGHT - 2 * frameBorder,
          0,
          0,
          frameBorder,
        );
        applyStyle(step, glassId, glassStyleId);

        const windowShapeId = productShape(
          step,
          geomContextId,
          [frameTopId, frameBottomId, frameLeftId, frameRightId, glassId],
        );
        const windowPos: [number, number, number] = [
          wallDef.pos[0] + wallDef.dir[0] * offset,
          wallDef.pos[1] + wallDef.dir[1] * offset,
          windowSillHeight,
        ];
        const windowPlacementId = localPlacement(step, storeyPlacementId, windowPos, wallDef.dir);
        const windowId = step.next(
          `IFCWINDOW('${generateGuid()}',${ref(ownerHistoryId)},'Fonster plan ${i + 1}',$,$,${ref(windowPlacementId)},${ref(windowShapeId)},$,${real(WINDOW_HEIGHT)},${real(WINDOW_WIDTH)},$,$,$)`,
        );
        elements.push(windowId);
      }
    }

    elementsByStorey.push(elements);
  }

  const roofElevation = floors * floorHeight;
  const roofSolidId = extrudedPolygonSolid(step, footprint, slabThickness);
  const roofShapeId = productShape(step, geomContextId, roofSolidId);
  const roofPlacementId = localPlacement(step, buildingPlacementId, [0, 0, roofElevation]);
  const roofId = step.next(
    `IFCSLAB('${generateGuid()}',${ref(ownerHistoryId)},'Tak',$,$,${ref(roofPlacementId)},${ref(roofShapeId)},$,.ROOF.)`,
  );
  elementsByStorey[floors - 1].push(roofId);

  step.next(
    `IFCRELAGGREGATES('${generateGuid()}',${ref(ownerHistoryId)},$,$,${ref(buildingId)},(${storeyIds.map((s) => ref(s)).join(",")}))`,
  );

  for (let i = 0; i < floors; i++) {
    step.next(
      `IFCRELCONTAINEDINSPATIALSTRUCTURE('${generateGuid()}',${ref(ownerHistoryId)},$,$,(${elementsByStorey[i].map((e) => ref(e)).join(",")}),${ref(storeyIds[i])})`,
    );
  }

  const header = [
    "ISO-10303-21;",
    "HEADER;",
    "FILE_DESCRIPTION((''),'2;1');",
    `FILE_NAME('generated.ifc','${new Date().toISOString()}',('IFCinfinity'),('IFCinfinity'),'IFCinfinity Generator','IFCinfinity','');`,
    "FILE_SCHEMA(('IFC4'));",
    "ENDSEC;",
    "DATA;",
  ].join("\n");

  const footer = ["ENDSEC;", "END-ISO-10303-21;"].join("\n");

  return `${header}\n${step.toString()}\n${footer}\n`;
}
