import type { ConnectorGlue, FormulaShapeDraft, PageDisplayList, PagePrimitive, ShapeSnapshot } from '@betteroffice/vsdx';
import type { ModelPoint } from '@betteroffice/vsdx';
import { cellValue } from './components/ribbon/commands';

export type ConnectorSide = 'north' | 'east' | 'south' | 'west' | 'centre';

/** Edge sides that grow an AutoConnect chevron; the centre never does. */
export type AutoConnectSide = Exclude<ConnectorSide, 'centre'>;

export const AUTO_CONNECT_SIDES: readonly AutoConnectSide[] = ['north', 'east', 'south', 'west'];

/** A hover chevron outside one shape edge, glued to that edge's connection point. */
export interface AutoConnectArrow { side: AutoConnectSide; point: ConnectionPoint; dir?: ModelPoint; }

/** Paint and hit state for the hover chevrons of one shape. */
export interface AutoConnectOverlayState {
  arrows: readonly AutoConnectArrow[];
  hovered: AutoConnectSide | null;
  alpha: number;
}

/** Screen-pixel look of the chevrons; geometry below derives from these. */
export const AUTO_CONNECT_GAP_PX = 10;
export const AUTO_CONNECT_SIZE_PX = 16;
export const AUTO_CONNECT_HIT_PX = 26;
/** Hover halo in screen pixels; keeps arrows alive across the edge-to-chevron gap. */
export const AUTO_CONNECT_HALO_PX = 40;
/** Chevron fade on hover, in milliseconds. */
export const AUTO_CONNECT_FADE_MS = 120;
/** Edge proximity that reveals connection points with no tool armed, in screen pixels. */
export const HOVER_PROXIMITY_PX = 12;
/** Grab radius of one connection point, in screen pixels. */
export const HOVER_POINT_HIT_PX = 10;
/** Visio ring diameter for one connection point, in screen pixels. */
export const HOVER_POINT_SIZE_PX = 9;
/** Minimum drag that creates a free-ended connector, in model inches. */
export const HOVER_FREE_DRAG_INCHES = 0.05;
/** Gap between a source shape and a Quick-Shape insert, in model inches. */
export const QUICK_SHAPE_GAP_INCHES = 0.5;

/** Thumbnail ids for the Quick-Shapes flyout, in display order. */
export const QUICK_SHAPE_IDS: readonly string[] = ['rectangle', 'square', 'circle', 'ellipse', 'rightTriangle'];

/** Centre and size of a Quick-Shape insert offset from a source edge, in model inches. */
export interface QuickShapePlacement {
  x: number;
  y: number;
  width: number;
  height: number;
  from: ConnectionPoint;
  to: ConnectionPoint;
}

/** A snap target in model inches. Centre glue is dynamic; outline glue pins to a Connection row. */
export interface ConnectionPoint extends ModelPoint { side: ConnectorSide; toCell?: string; }

export interface ConnectorDragEndpoint { shapeId: string; point: ConnectionPoint; }

export type ConnectorEndpointGlue = 'point' | 'dynamic' | 'unglued';

export interface ConnectorOverlayRoute {
  route: readonly ModelPoint[];
  selected: boolean;
  beginGlue: ConnectorEndpointGlue;
  endGlue: ConnectorEndpointGlue;
}

export interface ConnectorOverlayScene {
  hoverPoints: readonly ConnectionPoint[];
  snapPoint: ConnectionPoint | null;
  previewRoute: readonly ModelPoint[] | null;
  reroutePreview: ReadonlyArray<readonly ModelPoint[]>;
  connectors: ReadonlyArray<ConnectorOverlayRoute>;
}

export const CONNECTOR_SNAP_INCHES = 0.12;
export const CONNECTOR_GLUE_MATCH_INCHES = 1e-6;
const MIN_SPAN_INCHES = 0.01;

function cellNumber(shape: ShapeSnapshot, name: string): number | undefined {
  const raw = cellValue(shape, name);
  if (raw === undefined) return undefined;
  const parsed = Number(raw);
  return Number.isFinite(parsed) ? parsed : undefined;
}

/** Mirrors the engine rule: nonzero OneD, else a full Begin/End endpoint set. */
export function isConnectorShape(shape: ShapeSnapshot): boolean {
  const oneD = cellValue(shape, 'OneD');
  if (oneD !== undefined) {
    const parsed = Number(oneD.trim().replace(/^=/, ''));
    return Number.isNaN(parsed) ? true : parsed !== 0;
  }
  return ['BeginX', 'BeginY', 'EndX', 'EndY'].every((name) => cellValue(shape, name) !== undefined);
}

function shapeBounds(shape: ShapeSnapshot): { left: number; bottom: number; right: number; top: number; centre: ModelPoint } | null {
  const pinX = cellNumber(shape, 'PinX');
  const pinY = cellNumber(shape, 'PinY');
  const width = cellNumber(shape, 'Width');
  const height = cellNumber(shape, 'Height');
  if (pinX === undefined || pinY === undefined || width === undefined || height === undefined) return null;
  if (!(width > 0) || !(height > 0)) return null;
  const locPinX = cellNumber(shape, 'LocPinX') ?? width / 2;
  const locPinY = cellNumber(shape, 'LocPinY') ?? height / 2;
  const left = pinX - locPinX;
  const bottom = pinY - locPinY;
  const centre = { x: pinX, y: pinY };
  return { left, bottom, right: left + width, top: bottom + height, centre };
}

/** Local rotation and mirrors in model inches; mirrors the engine scene rule. */
interface ShapeOrientation { angle: number; flipX: boolean; flipY: boolean; }

function shapeOrientation(shape: ShapeSnapshot): ShapeOrientation {
  return {
    angle: cellNumber(shape, 'Angle') ?? 0,
    flipX: (cellNumber(shape, 'FlipX') ?? 0) !== 0,
    flipY: (cellNumber(shape, 'FlipY') ?? 0) !== 0,
  };
}

function orientPoint(orientation: ShapeOrientation, pin: ModelPoint, point: ModelPoint): ModelPoint {
  if (orientation.angle === 0 && !orientation.flipX && !orientation.flipY) return { x: point.x, y: point.y };
  const cos = Math.cos(orientation.angle);
  const sin = Math.sin(orientation.angle);
  const sx = orientation.flipX ? -1 : 1;
  const sy = orientation.flipY ? -1 : 1;
  const dx = point.x - pin.x;
  const dy = point.y - pin.y;
  return {
    x: pin.x + cos * sx * dx - sin * sy * dy,
    y: pin.y + sin * sx * dx + cos * sy * dy,
  };
}

function orientDirection(orientation: ShapeOrientation, dir: ModelPoint): ModelPoint {
  if (orientation.angle === 0 && !orientation.flipX && !orientation.flipY) return { ...dir };
  const cos = Math.cos(orientation.angle);
  const sin = Math.sin(orientation.angle);
  const sx = orientation.flipX ? -1 : 1;
  const sy = orientation.flipY ? -1 : 1;
  return { x: cos * sx * dir.x - sin * sy * dir.y, y: sin * sx * dir.x + cos * sy * dir.y };
}

const CONNECTION_ROWS: Record<AutoConnectSide, { row: number; toCell: string }> = {
  north: { row: 0, toCell: 'Connections.X1' },
  east: { row: 1, toCell: 'Connections.X2' },
  south: { row: 2, toCell: 'Connections.X3' },
  west: { row: 3, toCell: 'Connections.X4' },
};

/** True when both X and Y cells of one Connection row survived into the snapshot. */
function connectionRowExists(shape: ShapeSnapshot, rowIndex: number): boolean {
  let hasX = false;
  let hasY = false;
  for (const cell of shape.cells) {
    if (cell.locator.section !== 'Connection') continue;
    const row = cell.locator.row;
    const index = row !== null && typeof row === 'object' && 'index' in row ? row.index : undefined;
    if (index !== rowIndex) continue;
    if (cell.name === 'X') hasX = true;
    if (cell.name === 'Y') hasY = true;
  }
  return hasX && hasY;
}

/** Five snap targets in model inches: outline midpoints plus the dynamic centre. */
export function connectionPointsForShape(shape: ShapeSnapshot): ConnectionPoint[] {
  const bounds = shapeBounds(shape);
  if (!bounds) return [];
  const pin = { x: bounds.centre.x, y: bounds.centre.y };
  const orientation = shapeOrientation(shape);
  const centreX = (bounds.left + bounds.right) / 2;
  const centreY = (bounds.bottom + bounds.top) / 2;
  const outline = (side: AutoConnectSide, x: number, y: number): ConnectionPoint => {
    const moved = orientPoint(orientation, pin, { x, y });
    const row = CONNECTION_ROWS[side];
    return connectionRowExists(shape, row.row) ? { side, ...moved, toCell: row.toCell } : { side, ...moved };
  };
  return [
    outline('north', centreX, bounds.top),
    outline('east', bounds.right, centreY),
    outline('south', centreX, bounds.bottom),
    outline('west', bounds.left, centreY),
    { side: 'centre', x: pin.x, y: pin.y },
  ];
}

/** Edge midpoints as AutoConnect arrows; connectors and boundless shapes offer none. */
export function autoConnectArrowsForShape(shape: ShapeSnapshot): AutoConnectArrow[] {
  if (isConnectorShape(shape)) return [];
  const orientation = shapeOrientation(shape);
  const arrows: AutoConnectArrow[] = [];
  for (const point of connectionPointsForShape(shape)) {
    if (point.side === 'centre') continue;
    arrows.push({ side: point.side, point, dir: orientDirection(orientation, AUTO_CONNECT_DIRS[point.side]) });
  }
  return arrows;
}

const AUTO_CONNECT_DIRS: Record<AutoConnectSide, ModelPoint> = {
  north: { x: 0, y: 1 },
  east: { x: 1, y: 0 },
  south: { x: 0, y: -1 },
  west: { x: -1, y: 0 },
};

/** Screen-pixel density of one model inch; chevron geometry derives from it. */
export function autoConnectMetrics(frame: PageDisplayList, zoom: number): { pagePerModel: number; pixelsPerInch: number; modelPerPixel: number } {
  const scale = Math.hypot(frame.paintTransform.a, frame.paintTransform.b);
  const pagePerModel = Number.isFinite(scale) && scale > 0 ? scale : 96;
  const pixelsPerInch = pagePerModel * (Number.isFinite(zoom) && zoom > 0 ? zoom : 1);
  return { pagePerModel, pixelsPerInch, modelPerPixel: 1 / pixelsPerInch };
}

/** Chevron centre in model inches, held a fixed screen-pixel gap outside the edge. */
export function autoConnectArrowCenter(arrow: AutoConnectArrow, frame: PageDisplayList, zoom: number): ModelPoint {
  const metrics = autoConnectMetrics(frame, zoom);
  const dir = arrow.dir ?? AUTO_CONNECT_DIRS[arrow.side];
  const offset = (AUTO_CONNECT_GAP_PX + AUTO_CONNECT_SIZE_PX / 2) * metrics.modelPerPixel;
  return { x: arrow.point.x + dir.x * offset, y: arrow.point.y + dir.y * offset };
}

/** Nearest chevron within the screen-pixel hit radius of a canvas point. */
export function autoConnectArrowAt(arrows: readonly AutoConnectArrow[], frame: PageDisplayList, zoom: number, canvas: ModelPoint): AutoConnectArrow | null {
  if (!arrows.length) return null;
  const metrics = autoConnectMetrics(frame, zoom);
  const radius = (AUTO_CONNECT_HIT_PX / 2) * metrics.modelPerPixel;
  let best: AutoConnectArrow | null = null;
  let bestDistance = radius;
  for (const arrow of arrows) {
    const page = modelToPage(frame, autoConnectArrowCenter(arrow, frame, zoom));
    const distance = Math.hypot(page.x - canvas.x, page.y - canvas.y) / metrics.pagePerModel;
    if (distance <= bestDistance) { best = arrow; bestDistance = distance; }
  }
  return best;
}

/** Chevron centre in canvas CSS pixels, for anchoring the Quick-Shapes flyout. */
export function autoConnectArrowCss(arrow: AutoConnectArrow, frame: PageDisplayList, zoom: number): ModelPoint {
  const page = modelToPage(frame, autoConnectArrowCenter(arrow, frame, zoom));
  return { x: page.x * zoom, y: page.y * zoom };
}

/** True while a canvas point stays near a hovered shape, in screen pixels. */
export function autoConnectHaloHit(shape: ShapeSnapshot, frame: PageDisplayList, zoom: number, canvas: ModelPoint): boolean {
  const outline = connectionPointsForShape(shape).filter((point) => point.side !== 'centre');
  if (outline.length !== 4 || isConnectorShape(shape)) return false;
  const topLeft = modelToPage(frame, { x: Math.min(...outline.map((point) => point.x)), y: Math.max(...outline.map((point) => point.y)) });
  const bottomRight = modelToPage(frame, { x: Math.max(...outline.map((point) => point.x)), y: Math.min(...outline.map((point) => point.y)) });
  const halo = AUTO_CONNECT_HALO_PX / (Number.isFinite(zoom) && zoom > 0 ? zoom : 1);
  const left = Math.min(topLeft.x, bottomRight.x) - halo;
  const right = Math.max(topLeft.x, bottomRight.x) + halo;
  const top = Math.min(topLeft.y, bottomRight.y) - halo;
  const bottom = Math.max(topLeft.y, bottomRight.y) + halo;
  return canvas.x >= left && canvas.x <= right && canvas.y >= top && canvas.y <= bottom;
}

/** Insert geometry for a Quick-Shape dropped from one source edge, in model inches. */
export function quickShapePlacement(source: ShapeSnapshot, side: AutoConnectSide, width: number, height: number, gap = QUICK_SHAPE_GAP_INCHES): QuickShapePlacement | null {
  if (!Number.isFinite(width) || !Number.isFinite(height) || width <= 0 || height <= 0) return null;
  const points = connectionPointsForShape(source);
  const from = points.find((point) => point.side === side);
  if (!from) return null;
  const dir = orientDirection(shapeOrientation(source), AUTO_CONNECT_DIRS[side]);
  const facing: AutoConnectSide = Math.abs(dir.x) >= Math.abs(dir.y) ? (dir.x > 0 ? 'west' : 'east') : (dir.y > 0 ? 'south' : 'north');
  const toCell = `Connections.X${AUTO_CONNECT_SIDES.indexOf(facing) + 1}`;
  const length = Math.hypot(dir.x, dir.y) || 1;
  const ux = dir.x / length;
  const uy = dir.y / length;
  const hx = width / 2;
  const hy = height / 2;
  const ax = Math.abs(ux);
  const ay = Math.abs(uy);
  const extent = ax < 1e-12 ? hy / Math.max(ay, 1e-12) : ay < 1e-12 ? hx / Math.max(ax, 1e-12) : Math.min(hx / ax, hy / ay);
  const x = from.x + ux * (gap + extent);
  const y = from.y + uy * (gap + extent);
  const edge = facing === 'west' ? { x: x - hx, y } : facing === 'east' ? { x: x + hx, y } : facing === 'south' ? { x, y: y - hy } : { x, y: y + hy };
  return {
    x,
    y,
    width,
    height,
    from,
    to: { side: facing, x: edge.x, y: edge.y, toCell },
  };
}

/** Matches the engine's ShapeRouteStyle 1 route: one bend, horizontal first. */
export function routeConnector(from: ModelPoint, to: ModelPoint): ModelPoint[] {
  const start = { x: from.x, y: from.y };
  const end = { x: to.x, y: to.y };
  if (start.x === end.x || start.y === end.y) return [start, end];
  return [start, { x: end.x, y: start.y }, end];
}

/** Centre endpoints glue dynamically; outline endpoints pin to a Connection row. */
export function connectorGlue(shapeId: string, point: ConnectionPoint): ConnectorGlue {
  return point.toCell === undefined ? { shapeId } : { shapeId, toCell: point.toCell };
}

/** Inch formula without exponent noise or negative zero. */
export function formatInches(value: number): string {
  if (!Number.isFinite(value)) throw new Error('Connector geometry must be finite.');
  const rounded = Number(value.toFixed(6));
  return String(Object.is(rounded, -0) ? 0 : rounded);
}

/** 1D draft whose route the engine resolves from glue at layout time. */
export function connectorDraft(from: ModelPoint, to: ModelPoint): FormulaShapeDraft {
  const midX = (from.x + to.x) / 2;
  const midY = (from.y + to.y) / 2;
  const cell = (cellName: string, formula: string) => ({ locator: { cellName }, name: cellName, formula });
  return {
    name: 'Dynamic connector',
    cells: [
      cell('OneD', '1'),
      cell('BeginX', formatInches(from.x)),
      cell('BeginY', formatInches(from.y)),
      cell('EndX', formatInches(to.x)),
      cell('EndY', formatInches(to.y)),
      cell('PinX', formatInches(midX)),
      cell('PinY', formatInches(midY)),
      cell('Width', formatInches(Math.max(Math.abs(to.x - from.x), MIN_SPAN_INCHES))),
      cell('Height', formatInches(Math.max(Math.abs(to.y - from.y), MIN_SPAN_INCHES))),
      cell('ShapeRouteStyle', '1'),
      cell('EndArrow', '4'),
      cell('LineColor', 'RGB(23,32,51)'),
      cell('LineWeight', '0.02'),
      cell('LinePattern', '1'),
    ],
  };
}

/** Nearest snap target within the model-space threshold. */
export function nearestConnectionPoint(points: readonly ConnectionPoint[], at: ModelPoint, threshold = CONNECTOR_SNAP_INCHES): ConnectionPoint | null {
  let best: ConnectionPoint | null = null;
  let bestDistance = threshold;
  for (const point of points) {
    const distance = Math.hypot(point.x - at.x, point.y - at.y);
    if (distance <= bestDistance) { best = point; bestDistance = distance; }
  }
  return best;
}

/** Closest point with no threshold; used once a containing shape is known. */
export function nearestConnectionPointAnywhere(points: readonly ConnectionPoint[], at: ModelPoint): ConnectionPoint | null {
  let best: ConnectionPoint | null = null;
  let bestDistance = Number.POSITIVE_INFINITY;
  for (const point of points) {
    const distance = Math.hypot(point.x - at.x, point.y - at.y);
    if (distance < bestDistance) { best = point; bestDistance = distance; }
  }
  return best;
}

/** Nearest connection point inside a fixed screen-pixel grab radius. */
export function hoverPointAt(points: readonly ConnectionPoint[], frame: PageDisplayList, zoom: number, at: ModelPoint, hitPx = HOVER_POINT_HIT_PX): ConnectionPoint | null {
  if (!points.length) return null;
  const metrics = autoConnectMetrics(frame, zoom);
  return nearestConnectionPoint(points, at, Math.max(0, hitPx) * metrics.modelPerPixel);
}

/** Snap within threshold, else the nearest point of a shape containing the drop. */
export function dropTargetForPoint(shapes: readonly ShapeSnapshot[], at: ModelPoint): { shapeId: string; point: ConnectionPoint } | null {
  let best: { shapeId: string; point: ConnectionPoint } | null = null;
  let bestDistance = CONNECTOR_SNAP_INCHES;
  for (const shape of shapes) {
    if (isConnectorShape(shape)) continue;
    const point = nearestConnectionPoint(connectionPointsForShape(shape), at, bestDistance);
    if (!point) continue;
    bestDistance = Math.hypot(point.x - at.x, point.y - at.y);
    best = { shapeId: shape.id, point };
  }
  if (best) return best;
  let fallback: { shapeId: string; point: ConnectionPoint } | null = null;
  let fallbackDistance = Number.POSITIVE_INFINITY;
  for (const shape of shapes) {
    if (isConnectorShape(shape)) continue;
    const targets = connectionPointsForShape(shape).filter((point) => point.side !== 'centre');
    if (!targets.length) continue;
    const left = Math.min(...targets.map((point) => point.x));
    const right = Math.max(...targets.map((point) => point.x));
    const bottom = Math.min(...targets.map((point) => point.y));
    const top = Math.max(...targets.map((point) => point.y));
    if (at.x < left - CONNECTOR_SNAP_INCHES || at.x > right + CONNECTOR_SNAP_INCHES) continue;
    if (at.y < bottom - CONNECTOR_SNAP_INCHES || at.y > top + CONNECTOR_SNAP_INCHES) continue;
    const point = nearestConnectionPointAnywhere(connectionPointsForShape(shape), at);
    if (!point) continue;
    const distance = Math.hypot(point.x - at.x, point.y - at.y);
    if (distance < fallbackDistance) { fallbackDistance = distance; fallback = { shapeId: shape.id, point }; }
  }
  return fallback;
}

function findPrimitive(primitives: readonly PagePrimitive[], id: string, depth = 0): PagePrimitive | null {
  if (depth >= 256) return null;
  for (const primitive of primitives) {
    if (primitive.id === id) return primitive;
    if (primitive.kind === 'group') {
      const nested = findPrimitive(primitive.primitives, id, depth + 1);
      if (nested) return nested;
    }
  }
  return null;
}

/** Engine-resolved route in model inches, or null when the connector did not paint as a path. */
export function connectorRouteFromFrame(frame: PageDisplayList, sourcePartPath: string, sourceId: number): ModelPoint[] | null {
  const primitive = findPrimitive(frame.primitives, `${sourcePartPath}:${sourceId}`);
  if (!primitive || primitive.kind !== 'shape') return null;
  const route: ModelPoint[] = [];
  for (const command of primitive.path) {
    if ((command.type === 'move' || command.type === 'line') && Number.isFinite(Number(command.x)) && Number.isFinite(Number(command.y))) {
      route.push({ x: Number(command.x), y: Number(command.y) });
    }
  }
  return route.length >= 2 ? route : null;
}

/** Point-glued outline match wins over a coincident centre; otherwise unglued. */
export function classifyConnectorEndpoint(at: ModelPoint, shapes: readonly ShapeSnapshot[], tolerance = CONNECTOR_GLUE_MATCH_INCHES): ConnectorEndpointGlue {
  let dynamic = false;
  for (const shape of shapes) {
    if (isConnectorShape(shape)) continue;
    for (const point of connectionPointsForShape(shape)) {
      if (Math.hypot(point.x - at.x, point.y - at.y) > tolerance) continue;
      if (point.side !== 'centre') return 'point';
      dynamic = true;
    }
  }
  return dynamic ? 'dynamic' : 'unglued';
}

/** Glue states for both route ends, read from geometry. */
export function connectorEndpointGlue(route: readonly ModelPoint[], shapes: readonly ShapeSnapshot[]): [ConnectorEndpointGlue, ConnectorEndpointGlue] {
  if (route.length < 2) return ['unglued', 'unglued'];
  return [classifyConnectorEndpoint(route[0], shapes), classifyConnectorEndpoint(route[route.length - 1], shapes)];
}

export interface MovedShapeGeometry { x: number; y: number; width: number; height: number; }

/** Connection points of a dragged shape at its preview geometry, in model inches. */
export function movedShapePoints(shape: ShapeSnapshot, geometry: MovedShapeGeometry, resize: boolean): ConnectionPoint[] {
  if (!resize) {
    const pinX = cellNumber(shape, 'PinX');
    const pinY = cellNumber(shape, 'PinY');
    if (pinX === undefined || pinY === undefined) return [];
    const dx = geometry.x - pinX;
    const dy = geometry.y - pinY;
    return connectionPointsForShape(shape).map((point) => ({ ...point, x: point.x + dx, y: point.y + dy }));
  }
  const moved: ShapeSnapshot = {
    ...shape,
    cells: shape.cells.map((cell) => {
      if (cell.name === 'PinX') return { ...cell, formula: String(geometry.x), value: String(geometry.x) };
      if (cell.name === 'PinY') return { ...cell, formula: String(geometry.y), value: String(geometry.y) };
      if (cell.name === 'Width') return { ...cell, formula: String(geometry.width), value: String(geometry.width) };
      if (cell.name === 'Height') return { ...cell, formula: String(geometry.height), value: String(geometry.height) };
      return cell;
    }),
  };
  return connectionPointsForShape(moved);
}

function gluedSide(points: readonly ConnectionPoint[], at: ModelPoint): number {
  let best = -1;
  let bestDistance = CONNECTOR_GLUE_MATCH_INCHES;
  for (let index = 0; index < points.length; index++) {
    const distance = Math.hypot(points[index].x - at.x, points[index].y - at.y);
    if (distance <= bestDistance) { best = index; bestDistance = distance; }
  }
  return best;
}

function flattenConnectorShapes(shapes: readonly ShapeSnapshot[]): ShapeSnapshot[] {
  return shapes.flatMap((shape) => [shape, ...flattenConnectorShapes(shape.children)]);
}

/** Whole-route recompute for connectors glued to a dragged shape, in model inches. */
export function reroutePreviewForMove(shapes: readonly ShapeSnapshot[], frame: PageDisplayList, sourcePartPath: string, dragged: ShapeSnapshot, geometry: MovedShapeGeometry): ModelPoint[][] {
  if (isConnectorShape(dragged) || !shapes.some((shape) => shape.id === dragged.id)) return [];
  const before = connectionPointsForShape(dragged);
  if (!before.length) return [];
  const width = cellNumber(dragged, 'Width');
  const height = cellNumber(dragged, 'Height');
  const resize = width !== undefined && height !== undefined && (geometry.width !== width || geometry.height !== height);
  const after = movedShapePoints(dragged, geometry, resize);
  if (after.length !== before.length) return [];
  const previews: ModelPoint[][] = [];
  for (const shape of flattenConnectorShapes(shapes)) {
    if (!isConnectorShape(shape)) continue;
    const route = connectorRouteFromFrame(frame, sourcePartPath, shape.sourceId);
    if (!route || route.length < 2) continue;
    const beginSide = gluedSide(before, route[0]);
    const endSide = gluedSide(before, route[route.length - 1]);
    if (beginSide < 0 && endSide < 0) continue;
    const nextBegin = beginSide >= 0 ? after[beginSide] : route[0];
    const nextEnd = endSide >= 0 ? after[endSide] : route[route.length - 1];
    previews.push(routeConnector(nextBegin, nextEnd));
  }
  return previews;
}

/** Model inches to page pixels at zoom 1; callers scale linearly by zoom. */
export function modelToPage(frame: PageDisplayList, point: ModelPoint): ModelPoint {
  const t = frame.paintTransform;
  return { x: t.a * point.x + t.c * point.y + t.e, y: t.b * point.x + t.d * point.y + t.f };
}

function applyModelTransform(ctx: CanvasRenderingContext2D, frame: PageDisplayList, dpr: number, zoom: number): void {
  const t = frame.paintTransform;
  ctx.setTransform(dpr * zoom, 0, 0, dpr * zoom, 0, 0);
  ctx.transform(t.a, t.b, t.c, t.d, t.e, t.f);
}

function dot(ctx: CanvasRenderingContext2D, at: ModelPoint, radius: number, fill: string): void {
  ctx.beginPath();
  ctx.arc(at.x, at.y, radius, 0, Math.PI * 2);
  ctx.fillStyle = fill;
  ctx.fill();
}

function ring(ctx: CanvasRenderingContext2D, at: ModelPoint, radius: number, color: string, width: number): void {
  ctx.beginPath();
  ctx.arc(at.x, at.y, radius, 0, Math.PI * 2);
  ctx.strokeStyle = color;
  ctx.lineWidth = width;
  ctx.stroke();
}

function strokeRoute(ctx: CanvasRenderingContext2D, route: readonly ModelPoint[], color: string, width: number): void {
  ctx.beginPath();
  route.forEach((point, index) => { if (index === 0) ctx.moveTo(point.x, point.y); else ctx.lineTo(point.x, point.y); });
  ctx.strokeStyle = color;
  ctx.lineWidth = width;
  ctx.lineJoin = 'round';
  ctx.lineCap = 'round';
  ctx.stroke();
}

/** Hollow Visio ring on the outline, at a fixed screen size. */
function paintHoverPoint(ctx: CanvasRenderingContext2D, at: ModelPoint, radius: number, color: string, width: number): void {
  ctx.beginPath();
  ctx.arc(at.x, at.y, radius, 0, Math.PI * 2);
  ctx.strokeStyle = color;
  ctx.lineWidth = width;
  ctx.stroke();
}

/** Filled arrowhead in model inches, pointing along the final segment. */
export function arrowheadPolygon(tip: ModelPoint, tail: ModelPoint, length = 0.15, halfWidth = 0.055): ModelPoint[] {
  const dx = tip.x - tail.x;
  const dy = tip.y - tail.y;
  const size = Math.hypot(dx, dy);
  if (!(size > 0)) return [];
  const ux = dx / size;
  const uy = dy / size;
  return [
    { ...tip },
    { x: tip.x - ux * length - uy * halfWidth, y: tip.y - uy * length + ux * halfWidth },
    { x: tip.x - ux * length + uy * halfWidth, y: tip.y - uy * length - ux * halfWidth },
  ];
}

function paintArrowhead(ctx: CanvasRenderingContext2D, route: readonly ModelPoint[], fill: string): void {
  if (route.length < 2) return;
  const polygon = arrowheadPolygon(route[route.length - 1], route[route.length - 2]);
  if (!polygon.length) return;
  ctx.beginPath();
  polygon.forEach((point, index) => { if (index === 0) ctx.moveTo(point.x, point.y); else ctx.lineTo(point.x, point.y); });
  ctx.closePath();
  ctx.fillStyle = fill;
  ctx.fill();
}

function segmentMidpoints(route: readonly ModelPoint[]): ModelPoint[] {
  const result: ModelPoint[] = [];
  for (let index = 1; index < route.length; index++) {
    result.push({ x: (route[index - 1].x + route[index].x) / 2, y: (route[index - 1].y + route[index].y) / 2 });
  }
  return result;
}

/** Paints one selected endpoint from its glue state. */
export function paintConnectorEndpoint(ctx: CanvasRenderingContext2D, at: ModelPoint, glue: ConnectorEndpointGlue): void {
  if (glue === 'point') dot(ctx, at, 0.06, '#16a34a');
  else if (glue === 'dynamic') ring(ctx, at, 0.09, '#16a34a', 0.025);
  else ring(ctx, at, 0.09, '#9aa5b4', 0.025);
}

/** Overlay chrome in model space; zoom applies once through the canvas transform. */
export function paintConnectorOverlay(ctx: CanvasRenderingContext2D, frame: PageDisplayList, dpr: number, zoom: number, scene: ConnectorOverlayScene): void {
  applyModelTransform(ctx, frame, dpr, zoom);
  for (const connector of scene.connectors) {
    paintArrowhead(ctx, connector.route, '#172033');
    if (!connector.selected) continue;
    strokeRoute(ctx, connector.route, '#16a34a', 0.015);
    if (connector.route.length >= 1) paintConnectorEndpoint(ctx, connector.route[0], connector.beginGlue);
    if (connector.route.length >= 2) paintConnectorEndpoint(ctx, connector.route[connector.route.length - 1], connector.endGlue);
    ctx.fillStyle = '#2563eb';
    for (const midpoint of segmentMidpoints(connector.route)) ctx.fillRect(midpoint.x - 0.045, midpoint.y - 0.045, 0.09, 0.09);
  }
  const hoverModelPerPixel = autoConnectMetrics(frame, zoom).modelPerPixel;
  for (const point of scene.hoverPoints) paintHoverPoint(ctx, point, (HOVER_POINT_SIZE_PX / 2) * hoverModelPerPixel, '#16a34a', 1.25 * hoverModelPerPixel);
  if (scene.snapPoint) ring(ctx, scene.snapPoint, 0.1, '#ffffff', 0.025);
  if (scene.previewRoute && scene.previewRoute.length >= 2) {
    strokeRoute(ctx, scene.previewRoute, '#172033', 0.02);
    paintArrowhead(ctx, scene.previewRoute, '#172033');
  }
  for (const route of scene.reroutePreview) {
    if (route.length < 2) continue;
    ctx.save();
    ctx.setLineDash([0.08, 0.05]);
    strokeRoute(ctx, route, '#172033', 0.02);
    ctx.restore();
    paintArrowhead(ctx, route, '#172033');
  }
}

/** Grey hover chevrons at a fixed screen size; zoom scales only their model-inch math. */
export function paintAutoConnectOverlay(ctx: CanvasRenderingContext2D, frame: PageDisplayList, dpr: number, zoom: number, state: AutoConnectOverlayState | null): void {
  if (!state || !state.arrows.length || !(state.alpha > 0)) return;
  applyModelTransform(ctx, frame, dpr, zoom);
  const metrics = autoConnectMetrics(frame, zoom);
  const half = (AUTO_CONNECT_SIZE_PX / 2) * metrics.modelPerPixel;
  ctx.save();
  ctx.globalAlpha = Math.max(0, Math.min(1, state.alpha));
  ctx.lineJoin = 'round';
  ctx.lineCap = 'round';
  for (const arrow of state.arrows) {
    const center = autoConnectArrowCenter(arrow, frame, zoom);
    const dir = arrow.dir ?? AUTO_CONNECT_DIRS[arrow.side];
    const tip = { x: center.x + dir.x * half, y: center.y + dir.y * half };
    const base = { x: center.x - dir.x * half, y: center.y - dir.y * half };
    const first = { x: base.x - dir.y * half * 0.9, y: base.y + dir.x * half * 0.9 };
    const second = { x: base.x + dir.y * half * 0.9, y: base.y - dir.x * half * 0.9 };
    ctx.beginPath();
    ctx.moveTo(tip.x, tip.y);
    ctx.lineTo(first.x, first.y);
    ctx.lineTo(second.x, second.y);
    ctx.closePath();
    ctx.fillStyle = state.hovered === arrow.side ? '#374151' : '#6b7280';
    ctx.fill();
    ctx.strokeStyle = '#ffffff';
    ctx.lineWidth = 1.5 * metrics.modelPerPixel;
    ctx.stroke();
  }
  ctx.restore();
}
