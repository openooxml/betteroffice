import type { Affine, ConnectorChrome, GeometryPathCommand, PageDisplayList, PagePrimitive, PageSnapshot } from '@betteroffice/vsdx';
import { findShapePlacement } from './components/ribbon/commands';

export interface ChromePoint { x: number; y: number; }
export interface ChromeSegment { a: ChromePoint; b: ChromePoint; mid: ChromePoint; index: number; }
export interface SelectedConnectorChrome { chrome: ConnectorChrome; start: ChromePoint; end: ChromePoint; segments: ChromeSegment[]; draggable: ChromePoint[] | null; }
export interface ChromeSelection { pageId: string; shapeId: string; }

/** Chrome for the selected shape when it is a connector with a painted route. */
export function selectedConnectorChrome(frame: PageDisplayList, page: PageSnapshot, selection: ChromeSelection): SelectedConnectorChrome | null {
  if (page.id !== selection.pageId) return null;
  const placement = findShapePlacement(page.shapes, selection.shapeId);
  if (!placement) return null;
  const id = `${page.sourcePartPath}:${placement.shape.sourceId}`;
  const chrome = frame.connectors?.find((entry) => entry.id === id);
  if (!chrome) return null;
  const path = findShapePath(frame.primitives, id);
  if (!path || path.nested) return null;
  const runs = connectorRuns(path.commands, path.transform);
  const vertices = runs.flat();
  if (vertices.length < 2) return null;
  const segments: ChromeSegment[] = [];
  let offset = 0;
  for (const run of runs) {
    for (let index = 1; index < run.length; index++) {
      const a = run[index - 1], b = run[index];
      if (a.x === b.x && a.y === b.y) continue;
      segments.push({ a, b, mid: { x: (a.x + b.x) / 2, y: (a.y + b.y) / 2 }, index: offset + index - 1 });
    }
    offset += run.length;
  }
  return { chrome, start: vertices[0], end: vertices[vertices.length - 1], segments, draggable: runs.length === 1 && chrome.routable ? runs[0] : null };
}

/** `nested` marks a connector under a group, whose ancestor transforms this chrome does not compose. */
function findShapePath(primitives: readonly PagePrimitive[], id: string, depth = 0): { commands: readonly GeometryPathCommand[]; transform?: Affine; nested: boolean } | null {
  if (depth >= 256) return null;
  for (const primitive of primitives) {
    if (primitive.kind === 'shape' && primitive.id === id) return { commands: primitive.path, transform: primitive.transform, nested: depth > 0 };
    if (primitive.kind === 'group') {
      const nested = findShapePath(primitive.primitives, id, depth + 1);
      if (nested) return nested;
    }
  }
  return null;
}

/** Straight move/line runs; curves and closes break the run. */
export function connectorRuns(path: readonly GeometryPathCommand[], transform?: Affine): ChromePoint[][] {
  const apply = (x: number, y: number): ChromePoint => transform
    ? { x: transform.a * x + transform.c * y + transform.e, y: transform.b * x + transform.d * y + transform.f }
    : { x, y };
  const runs: ChromePoint[][] = [];
  let run: ChromePoint[] = [];
  const commit = () => { if (run.length) runs.push(run); run = []; };
  for (const command of path) {
    if (command.type === 'move') { commit(); run.push(apply(Number(command.x), Number(command.y))); }
    else if (command.type === 'line') {
      if (!run.length) { run.push(apply(Number(command.x), Number(command.y))); continue; }
      run.push(apply(Number(command.x), Number(command.y)));
    } else if (typeof command.x === 'number' && typeof command.y === 'number') { commit(); run.push(apply(command.x, command.y)); }
    else commit();
  }
  commit();
  return runs.filter((entry) => entry.length > 1);
}

const GLUE_SHAPE = '#16a34a';
const GLUE_RING = '#16a34a';
const GLUE_FREE = '#64748b';
const SEGMENT_DOT = '#2563eb';

const ROUTE_EPSILON = 1e-6;

/** The segment whose dot is under the pointer, if any. */
export function hitSegmentDot(selected: SelectedConnectorChrome, point: ChromePoint, tolerance: number): ChromeSegment | null {
  let best: ChromeSegment | null = null;
  let nearest = tolerance;
  for (const segment of selected.segments) {
    const distance = Math.hypot(segment.mid.x - point.x, segment.mid.y - point.y);
    if (distance <= nearest) { nearest = distance; best = segment; }
  }
  return best;
}

/** Chrome repainted for a drag preview route. */
export function previewChrome(selected: SelectedConnectorChrome, vertices: ChromePoint[]): SelectedConnectorChrome {
  const segments: ChromeSegment[] = [];
  for (let index = 1; index < vertices.length; index++) {
    const a = vertices[index - 1], b = vertices[index];
    if (Math.hypot(b.x - a.x, b.y - a.y) < ROUTE_EPSILON) continue;
    segments.push({ a, b, mid: { x: (a.x + b.x) / 2, y: (a.y + b.y) / 2 }, index: index - 1 });
  }
  return { chrome: selected.chrome, start: vertices[0], end: vertices[vertices.length - 1], segments, draggable: vertices };
}

/** Perpendicular-only segment move with re-plumbing; parallel travel is discarded. */
export function dragSegmentRoute(vertices: readonly ChromePoint[], segmentIndex: number, anchor: ChromePoint, pointer: ChromePoint): ChromePoint[] {
  const a = vertices[segmentIndex], b = vertices[segmentIndex + 1];
  if (!a || !b) return [...vertices];
  const dx = b.x - a.x, dy = b.y - a.y;
  const length = Math.hypot(dx, dy);
  if (length < ROUTE_EPSILON) return [...vertices];
  const normal = { x: -dy / length, y: dx / length };
  const travel = (pointer.x - anchor.x) * normal.x + (pointer.y - anchor.y) * normal.y;
  if (Math.abs(travel) < ROUTE_EPSILON) return [...vertices];
  const route = [
    ...vertices.slice(0, segmentIndex + 1).map((vertex) => ({ ...vertex })),
    { x: a.x + normal.x * travel, y: a.y + normal.y * travel },
    { x: b.x + normal.x * travel, y: b.y + normal.y * travel },
    ...vertices.slice(segmentIndex + 1).map((vertex) => ({ ...vertex })),
  ];
  return simplifyRoute(route);
}

function simplifyRoute(route: ChromePoint[]): ChromePoint[] {
  const points: ChromePoint[] = [];
  for (const vertex of route) {
    const last = points[points.length - 1];
    if (!last || Math.hypot(vertex.x - last.x, vertex.y - last.y) >= ROUTE_EPSILON) points.push(vertex);
  }
  const kept: ChromePoint[] = [];
  for (const vertex of points) {
    kept.push(vertex);
    while (kept.length >= 3) {
      const tip = kept[kept.length - 1], prev = kept[kept.length - 2], first = kept[kept.length - 3];
      if (Math.hypot(tip.x - first.x, tip.y - first.y) < ROUTE_EPSILON) { kept.pop(); kept.pop(); continue; }
      const ux = prev.x - first.x, uy = prev.y - first.y;
      const vx = tip.x - prev.x, vy = tip.y - prev.y;
      const cross = ux * vy - uy * vx;
      const dot = ux * vx + uy * vy;
      const scale = Math.hypot(ux, uy) * Math.hypot(vx, vy);
      if (scale > 0 && Math.abs(cross) / scale < ROUTE_EPSILON && dot > 0) kept.splice(kept.length - 2, 1);
      else break;
    }
  }
  return kept.length >= 2 ? kept : points.slice(0, 2);
}

/** Endpoint markers plus one dot per straight segment, in constant screen sizes. */
export function paintConnectorChrome(context: CanvasRenderingContext2D, frame: PageDisplayList, selected: SelectedConnectorChrome, dpr: number, zoom: number): void {
  const paint = frame.paintTransform;
  const scale = zoom * Math.hypot(paint.a, paint.b);
  if (!Number.isFinite(scale) || scale <= 0) return;
  context.save();
  try {
    context.setTransform(dpr * zoom, 0, 0, dpr * zoom, 0, 0);
    context.transform(paint.a, paint.b, paint.c, paint.d, paint.e, paint.f);
    const dotRadius = 4 / scale;
    const endRadius = 5 / scale;
    const ringWidth = 2 / scale;
    if (selected.draggable) for (const segment of selected.segments) {
      context.beginPath();
      context.arc(segment.mid.x, segment.mid.y, dotRadius, 0, Math.PI * 2);
      context.fillStyle = SEGMENT_DOT;
      context.fill();
    }
    paintEndpoint(context, selected.start, selected.chrome.begin, endRadius, ringWidth);
    paintEndpoint(context, selected.end, selected.chrome.end, endRadius, ringWidth);
  } finally {
    context.restore();
  }
}

function paintEndpoint(context: CanvasRenderingContext2D, point: ChromePoint, glue: ConnectorChrome['begin'], radius: number, ringWidth: number): void {
  context.beginPath();
  context.arc(point.x, point.y, radius, 0, Math.PI * 2);
  if (glue === 'shape') {
    context.fillStyle = GLUE_SHAPE;
    context.fill();
  } else {
    context.fillStyle = '#ffffff';
    context.fill();
    context.lineWidth = ringWidth;
    context.strokeStyle = glue === 'point' ? GLUE_RING : GLUE_FREE;
    context.stroke();
  }
}
