// Enkel 2D-ritare (ovanifrån/planvy) för att rita en byggnads utbredning som
// en polygon med valfria hörnpunkter, istället för att bara ange en area.
//
// Två genvägar utöver klick-för-klick-ritning:
// - Håll Shift när en punkt läggs till för att låsa den nya kanten till en
//   rät vinkel (horisontell eller vertikal relativt föregående punkt).
// - Håll ner vänster musknapp och dra för att rita en vanlig rektangel
//   direkt, med hörnen i start- och släpp-punkten.

export interface FootprintPoint {
  x: number;
  y: number;
}

const BASE_SCALE = 20; // pixlar per meter vid zoom = 1
const MIN_ZOOM = 1;
const MAX_ZOOM = 4;
const ZOOM_STEP = 1.25;
const CLOSE_RADIUS = 12; // pixlar, avstånd till startpunkten för att stänga formen
const DRAG_THRESHOLD = 6; // pixlar, rörelse innan en mousedown räknas som en dragning
const BACKGROUND_COLOR = "#fafafa";
const GRID_COLOR = "rgba(0, 0, 0, 0.08)";
const AXIS_COLOR = "rgba(0, 0, 0, 0.25)";
const EDGE_COLOR = "#0091d5";
const FILL_COLOR = "rgba(0, 145, 213, 0.12)";
const START_POINT_COLOR = "#ff6b00";
const POINT_COLOR = "#0091d5";
const DRAG_RECT_COLOR = "#0091d5";

export class FootprintEditor {
  private points: FootprintPoint[] = [];
  private closed = false;
  private hover: FootprintPoint | null = null;
  private nearStart = false;
  private dragStart: FootprintPoint | null = null;
  private dragCurrent: FootprintPoint | null = null;
  private isDragging = false;
  private zoom = 1;
  private readonly baseWidth: number;
  private readonly baseHeight: number;
  private readonly ctx: CanvasRenderingContext2D;

  onChange: (() => void) | null = null;

  constructor(private readonly canvas: HTMLCanvasElement) {
    const ctx = canvas.getContext("2d");
    if (!ctx) throw new Error("Kunde inte skapa 2D-kontext för canvas");
    this.ctx = ctx;
    this.baseWidth = canvas.width;
    this.baseHeight = canvas.height;

    canvas.addEventListener("mousedown", (event) => this.handleMouseDown(event));
    canvas.addEventListener("mousemove", (event) => this.handleMouseMove(event));
    canvas.addEventListener("mouseup", (event) => this.handleMouseUp(event));
    canvas.addEventListener("mouseleave", () => this.handleMouseLeave());

    this.render();
  }

  private get scale(): number {
    return BASE_SCALE * this.zoom;
  }

  private toWorld(event: MouseEvent): FootprintPoint {
    const rect = this.canvas.getBoundingClientRect();
    const px = event.clientX - rect.left - this.canvas.width / 2;
    const py = event.clientY - rect.top - this.canvas.height / 2;
    return { x: px / this.scale, y: py / this.scale };
  }

  private toScreen(point: FootprintPoint): FootprintPoint {
    return {
      x: point.x * this.scale + this.canvas.width / 2,
      y: point.y * this.scale + this.canvas.height / 2,
    };
  }

  getZoom(): number {
    return this.zoom;
  }

  /** Ändrar zoomnivån (1 = normal, upp till MAX_ZOOM) och skalar upp canvasens
   *  faktiska pixelstorlek i samma takt - så att synligt mätt-område i meter
   *  förblir detsamma, men allt (linjer, mätfält) blir fysiskt större på
   *  skärmen. Tänkt att användas tillsammans med en scrollbar wrapper runt
   *  canvasen så att man kan nå mätfält som annars skulle överlappa varandra
   *  när sidorna är korta. */
  setZoom(zoom: number): void {
    const clamped = Math.min(MAX_ZOOM, Math.max(MIN_ZOOM, zoom));
    if (clamped === this.zoom) return;
    this.zoom = clamped;
    this.canvas.width = Math.round(this.baseWidth * this.zoom);
    this.canvas.height = Math.round(this.baseHeight * this.zoom);
    this.render();
    this.onChange?.();
  }

  zoomIn(): void {
    this.setZoom(this.zoom * ZOOM_STEP);
  }

  zoomOut(): void {
    this.setZoom(this.zoom / ZOOM_STEP);
  }

  /** Låser en punkt till horisontell eller vertikal riktning relativt föregående punkt. */
  private applyRightAngleSnap(world: FootprintPoint): FootprintPoint {
    if (this.points.length === 0) return world;
    const last = this.points[this.points.length - 1];
    const dx = world.x - last.x;
    const dy = world.y - last.y;
    return Math.abs(dx) > Math.abs(dy) ? { x: world.x, y: last.y } : { x: last.x, y: world.y };
  }

  private handleMouseDown(event: MouseEvent): void {
    if (this.closed || event.button !== 0) return;
    this.dragStart = this.toWorld(event);
    this.dragCurrent = this.dragStart;
    this.isDragging = false;
  }

  private handleMouseMove(event: MouseEvent): void {
    if (this.closed) return;
    const world = this.toWorld(event);

    if (this.dragStart) {
      const startScreen = this.toScreen(this.dragStart);
      const currentScreen = this.toScreen(world);
      const moved = Math.hypot(currentScreen.x - startScreen.x, currentScreen.y - startScreen.y);
      if (moved > DRAG_THRESHOLD) this.isDragging = true;
      this.dragCurrent = world;
      this.render();
      return;
    }

    this.nearStart = this.points.length >= 3 && this.isNearStart(world);
    this.hover = event.shiftKey ? this.applyRightAngleSnap(world) : world;
    this.render();
  }

  private handleMouseUp(event: MouseEvent): void {
    if (this.closed || !this.dragStart || event.button !== 0) return;
    const world = this.toWorld(event);

    if (this.isDragging) {
      const x1 = Math.min(this.dragStart.x, world.x);
      const x2 = Math.max(this.dragStart.x, world.x);
      const y1 = Math.min(this.dragStart.y, world.y);
      const y2 = Math.max(this.dragStart.y, world.y);
      this.points = [
        { x: x1, y: y1 },
        { x: x2, y: y1 },
        { x: x2, y: y2 },
        { x: x1, y: y2 },
      ];
      this.closed = true;
    } else if (this.points.length >= 3 && this.isNearStart(world)) {
      this.closed = true;
    } else {
      this.points.push(event.shiftKey ? this.applyRightAngleSnap(world) : world);
    }

    this.dragStart = null;
    this.dragCurrent = null;
    this.isDragging = false;
    this.hover = null;
    this.render();
    this.onChange?.();
  }

  private handleMouseLeave(): void {
    // Avbryt en pågående dragning om muspekaren lämnar canvasen utan att släppas.
    this.dragStart = null;
    this.dragCurrent = null;
    this.isDragging = false;
    this.hover = null;
    this.nearStart = false;
    this.render();
  }

  private isNearStart(world: FootprintPoint): boolean {
    const start = this.toScreen(this.points[0]);
    const point = this.toScreen(world);
    return Math.hypot(point.x - start.x, point.y - start.y) < CLOSE_RADIUS;
  }

  undo(): void {
    if (this.closed) {
      this.closed = false;
    } else {
      this.points.pop();
    }
    this.render();
    this.onChange?.();
  }

  clear(): void {
    this.points = [];
    this.closed = false;
    this.hover = null;
    this.nearStart = false;
    this.dragStart = null;
    this.dragCurrent = null;
    this.isDragging = false;
    this.render();
    this.onChange?.();
  }

  isClosed(): boolean {
    return this.closed;
  }

  pointCount(): number {
    return this.points.length;
  }

  /** Returnerar den slutna polygonen, eller null om formen inte är klar (minst 3 punkter, stängd). */
  getFootprint(): FootprintPoint[] | null {
    return this.closed && this.points.length >= 3 ? this.points.slice() : null;
  }

  getArea(): number {
    if (this.points.length < 3) return 0;
    let sum = 0;
    for (let i = 0; i < this.points.length; i++) {
      const p1 = this.points[i];
      const p2 = this.points[(i + 1) % this.points.length];
      sum += p1.x * p2.y - p2.x * p1.y;
    }
    return Math.abs(sum) / 2;
  }

  /** Antal ritade sidor: en sida mellan varje par av angränsande punkter, plus
   *  (när formen är stängd) den stängande sidan mellan sista och första punkten. */
  getEdgeCount(): number {
    if (this.points.length < 2) return 0;
    return this.closed ? this.points.length : this.points.length - 1;
  }

  /** Sant för sidan som stänger formen (sista → första punkten) - dess längd
   *  bestäms av att formen måste slutas och kan därför inte sättas fritt. */
  isClosingEdge(index: number): boolean {
    return this.closed && index === this.points.length - 1;
  }

  getEdgeLength(index: number): number {
    const n = this.points.length;
    const a = this.points[index];
    const b = this.points[(index + 1) % n];
    return Math.hypot(b.x - a.x, b.y - a.y);
  }

  getEdgeLengths(): number[] {
    const count = this.getEdgeCount();
    const lengths: number[] = [];
    for (let i = 0; i < count; i++) lengths.push(this.getEdgeLength(i));
    return lengths;
  }

  /** Mittpunkten för en sida, i canvasens egna pixelkoordinater - för att
   *  kunna placera ett måttfält direkt ovanpå sidan i figuren. */
  getEdgeMidpointScreen(index: number): FootprintPoint {
    const n = this.points.length;
    const a = this.points[index];
    const b = this.points[(index + 1) % n];
    return this.toScreen({ x: (a.x + b.x) / 2, y: (a.y + b.y) / 2 });
  }

  /** Sätter längden (i meter) för sidan mellan punkt `index` och `index + 1`, genom
   *  att flytta den andra punkten längs sidans befintliga riktning och skjuta alla
   *  därpå följande punkter i samma led - så att deras egna sidors längder och
   *  riktningar bevaras oförändrade. Går inte att sätta på den stängande sidan
   *  (se isClosingEdge) eftersom dess längd är en följd av att formen sluts, inte
   *  ett fritt val. Returnerar false om ändringen avvisades. */
  setEdgeLength(index: number, newLength: number): boolean {
    const n = this.points.length;
    if (index < 0 || index >= this.getEdgeCount() || this.isClosingEdge(index)) return false;
    if (!Number.isFinite(newLength) || newLength <= 0) return false;

    const a = this.points[index];
    const b = this.points[index + 1];
    const dx = b.x - a.x;
    const dy = b.y - a.y;
    const currentLength = Math.hypot(dx, dy);
    if (currentLength < 1e-9) return false;

    const ux = dx / currentLength;
    const uy = dy / currentLength;
    const deltaX = a.x + ux * newLength - b.x;
    const deltaY = a.y + uy * newLength - b.y;

    for (let i = index + 1; i < n; i++) {
      this.points[i] = { x: this.points[i].x + deltaX, y: this.points[i].y + deltaY };
    }

    this.render();
    this.onChange?.();
    return true;
  }

  private render(): void {
    const ctx = this.ctx;
    const { width, height } = this.canvas;

    ctx.fillStyle = BACKGROUND_COLOR;
    ctx.fillRect(0, 0, width, height);

    ctx.strokeStyle = GRID_COLOR;
    ctx.lineWidth = 1;
    for (let x = width / 2; x < width; x += this.scale) this.gridLineV(x);
    for (let x = width / 2; x > 0; x -= this.scale) this.gridLineV(x);
    for (let y = height / 2; y < height; y += this.scale) this.gridLineH(y);
    for (let y = height / 2; y > 0; y -= this.scale) this.gridLineH(y);

    ctx.strokeStyle = AXIS_COLOR;
    ctx.beginPath();
    ctx.moveTo(width / 2, 0);
    ctx.lineTo(width / 2, height);
    ctx.moveTo(0, height / 2);
    ctx.lineTo(width, height / 2);
    ctx.stroke();

    if (this.dragStart && this.dragCurrent && this.isDragging) {
      const a = this.toScreen(this.dragStart);
      const b = this.toScreen(this.dragCurrent);
      const x = Math.min(a.x, b.x);
      const y = Math.min(a.y, b.y);
      const w = Math.abs(b.x - a.x);
      const h = Math.abs(b.y - a.y);
      ctx.fillStyle = FILL_COLOR;
      ctx.fillRect(x, y, w, h);
      ctx.strokeStyle = DRAG_RECT_COLOR;
      ctx.lineWidth = 2;
      ctx.strokeRect(x, y, w, h);
      return;
    }

    if (this.points.length === 0) return;

    const screenPoints = this.points.map((p) => this.toScreen(p));

    if (this.closed) {
      ctx.beginPath();
      ctx.moveTo(screenPoints[0].x, screenPoints[0].y);
      for (const p of screenPoints.slice(1)) ctx.lineTo(p.x, p.y);
      ctx.closePath();
      ctx.fillStyle = FILL_COLOR;
      ctx.fill();
    }

    ctx.strokeStyle = EDGE_COLOR;
    ctx.lineWidth = 2;
    ctx.beginPath();
    ctx.moveTo(screenPoints[0].x, screenPoints[0].y);
    for (const p of screenPoints.slice(1)) ctx.lineTo(p.x, p.y);
    if (this.closed) ctx.closePath();
    ctx.stroke();

    if (!this.closed && this.hover) {
      const last = screenPoints[screenPoints.length - 1];
      const hoverScreen = this.toScreen(this.hover);
      ctx.strokeStyle = "rgba(0, 145, 213, 0.5)";
      ctx.setLineDash([4, 4]);
      ctx.beginPath();
      ctx.moveTo(last.x, last.y);
      ctx.lineTo(hoverScreen.x, hoverScreen.y);
      ctx.stroke();
      ctx.setLineDash([]);
    }

    for (let i = 0; i < screenPoints.length; i++) {
      const p = screenPoints[i];
      const isStart = i === 0;
      const radius = isStart && this.nearStart ? 9 : isStart ? 6 : 4;
      ctx.beginPath();
      ctx.arc(p.x, p.y, radius, 0, Math.PI * 2);
      ctx.fillStyle = isStart ? START_POINT_COLOR : POINT_COLOR;
      ctx.fill();
    }
  }

  private gridLineV(x: number): void {
    const ctx = this.ctx;
    ctx.beginPath();
    ctx.moveTo(x, 0);
    ctx.lineTo(x, this.canvas.height);
    ctx.stroke();
  }

  private gridLineH(y: number): void {
    const ctx = this.ctx;
    ctx.beginPath();
    ctx.moveTo(0, y);
    ctx.lineTo(this.canvas.width, y);
    ctx.stroke();
  }
}
