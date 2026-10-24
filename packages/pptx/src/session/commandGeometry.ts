import type { HitTestResult, PositionedTextLine, SlideDisplayList, SlidePrimitive, TextBoxPrimitive } from '../types';
import type { PresentationCommandFrame, PresentationCommandMotion, PresentationCommandSelected, PresentationCommandSelection } from './commandProtocol';
import { caretColumnOffset, graphemeBoundaries, verticalCaretMove, wordBoundary } from '../../../../shared/textSelection';

export interface CommandRect { x: number; y: number; width: number; height: number }

export interface CommandGeometryLine {
  box: TextBoxPrimitive;
  line: PositionedTextLine;
  transforms: readonly SlidePrimitive[];
}

export interface CommandGeometryIndex {
  hits: { primitive: SlidePrimitive; transforms: readonly SlidePrimitive[]; clips: readonly SlidePrimitive[]; textLength?: number }[];
  lines: Map<string, CommandGeometryLine[]>;
}

type GeometryBounds = Pick<SlidePrimitive, 'x' | 'y' | 'w' | 'h' | 'transform'>;
type CaretLine = Pick<PositionedTextLine, 'columnOffset' | 'x' | 'y' | 'height' | 'start' | 'end' | 'caretStops'>;
type HitPrimitive = GeometryBounds & { kind: SlidePrimitive['kind']; shapeId?: string; storyId?: string; lines?: CaretLine[] };
interface NavigationLine { line: CaretLine }
interface NavigationIndex { lines: Map<string, NavigationLine[]> }

export interface CommandHitGeometryIndex extends NavigationIndex {
  size?: { width: number; height: number };
  hits: { primitive: HitPrimitive; transforms: readonly GeometryBounds[]; clips: readonly GeometryBounds[]; textLength?: number }[];
}

const displayLists = new WeakMap<PresentationCommandFrame, SlideDisplayList>();

export function decodeCommandDisplayList(frame: PresentationCommandFrame): SlideDisplayList {
  let list = displayLists.get(frame);
  if (!list) {
    list = JSON.parse(new TextDecoder().decode(frame.displayList)) as SlideDisplayList;
    displayLists.set(frame, list);
  }
  return list;
}

function targetKey(shapeId: string, storyId?: string): string { return JSON.stringify([shapeId, storyId]); }

function textBoxLength(box: TextBoxPrimitive): number {
  return box.paragraphs.reduce((length, paragraph) =>
    length + paragraph.runs.reduce((sum, run) => sum + run.text.length, 0), Math.max(0, box.paragraphs.length - 1));
}

export function indexGeometry(list: SlideDisplayList): CommandGeometryIndex {
  const index: CommandGeometryIndex = { hits: [], lines: new Map() };
  function visit(primitives: readonly SlidePrimitive[], parents: readonly SlidePrimitive[], clips: readonly SlidePrimitive[]): void {
    for (const primitive of primitives) {
      const transforms = [...parents, primitive];
      if (primitive.shapeId) index.hits.push({ primitive, transforms, clips,
        ...(primitive.kind === 'textBox' ? { textLength: textBoxLength(primitive) } : {}) });
      if (primitive.kind === 'textBox' && primitive.shapeId && primitive.storyId) {
        const key = targetKey(primitive.shapeId, primitive.storyId);
        const lines = index.lines.get(key) ?? [];
        for (const line of primitive.lines) lines.push({ box: primitive, line, transforms });
        index.lines.set(key, lines);
      } else if (primitive.kind === 'chart' || primitive.kind === 'table') {
        visit(primitive.primitives, transforms, [...clips, primitive]);
      }
    }
  }
  visit(list.primitives, [], []);
  return index;
}

export function indexHitGeometry(list: SlideDisplayList): CommandHitGeometryIndex {
  const index: CommandHitGeometryIndex = { hits: [], lines: new Map(), size: { width: list.width, height: list.height } };
  function visit(primitives: readonly SlidePrimitive[], parents: readonly GeometryBounds[], clips: readonly GeometryBounds[]): void {
    for (const source of primitives) {
      const primitive: HitPrimitive = { kind: source.kind, shapeId: source.shapeId,
        x: source.x, y: source.y, w: source.w, h: source.h,
        ...(source.transform ? { transform: { ...source.transform } } : {}) };
      const transforms = [...parents, primitive];
      if (source.shapeId) index.hits.push({ primitive, transforms, clips,
        ...(source.kind === 'textBox' ? { textLength: textBoxLength(source) } : {}) });
      if (source.kind === 'textBox' && source.shapeId && source.storyId) {
        primitive.storyId = source.storyId;
        primitive.lines = source.lines.map((line) => ({ columnOffset: line.columnOffset, x: line.x, y: line.y, height: line.height,
          start: line.start, end: line.end, caretStops: line.caretStops.map((stop) => ({ position: stop.position, x: stop.x })) }));
        const key = targetKey(source.shapeId, source.storyId);
        const lines = index.lines.get(key) ?? [];
        for (const line of primitive.lines) lines.push({ line });
        index.lines.set(key, lines);
      } else if (source.kind === 'chart' || source.kind === 'table') {
        visit(source.primitives, transforms, [...clips, primitive]);
      }
    }
  }
  visit(list.primitives, [], []);
  return index;
}

export function findLine<T extends NavigationLine>(index: { lines: Map<string, T[]> }, selection: PresentationCommandSelected | null): T | null {
  if (!selection || selection.kind === 'object') return null;
  const lines = index.lines.get(targetKey(selection.shapeId, selection.storyId)) ?? [];
  return lines.find(({ line }) => selection.focus >= line.start &&
    (selection.focus < line.end || selection.focus === line.end && selection.affinity === 'upstream'))
    ?? lines.find(({ line }) => selection.focus >= line.start && selection.focus <= line.end)
    ?? lines[lines.length - 1] ?? null;
}

function transformPoint(primitive: GeometryBounds, point: { x: number; y: number }, inverse = false): { x: number; y: number } {
  const cx = primitive.x + primitive.w / 2;
  const cy = primitive.y + primitive.h / 2;
  const angle = (primitive.transform?.rotationDeg ?? 0) * Math.PI / 180;
  const cos = Math.cos(angle); const sin = Math.sin(angle);
  let x = point.x - cx; let y = point.y - cy;
  if (inverse) {
    const rotatedX = x * cos + y * sin;
    y = y * cos - x * sin;
    x = rotatedX;
  }
  if (primitive.transform?.flipH) x = -x;
  if (primitive.transform?.flipV) y = -y;
  return inverse ? { x: cx + x, y: cy + y }
    : { x: cx + x * cos - y * sin, y: cy + x * sin + y * cos };
}

export function transformedRect(transforms: readonly GeometryBounds[], rect: CommandRect): CommandRect {
  const points = [{ x: rect.x, y: rect.y }, { x: rect.x + rect.width, y: rect.y },
    { x: rect.x, y: rect.y + rect.height }, { x: rect.x + rect.width, y: rect.y + rect.height }];
  for (let i = transforms.length - 1; i >= 0; i -= 1) {
    for (let j = 0; j < points.length; j += 1) points[j] = transformPoint(transforms[i], points[j]);
  }
  const x = Math.min(...points.map((point) => point.x));
  const y = Math.min(...points.map((point) => point.y));
  return { x, y, width: Math.max(...points.map((point) => point.x)) - x,
    height: Math.max(...points.map((point) => point.y)) - y };
}

function stopX(line: CaretLine, position: number): number {
  const first = line.caretStops[0];
  if (!first) return line.x;
  return line.caretStops.reduce((best, stop) =>
    Math.abs(stop.position - position) < Math.abs(best.position - position) ? stop : best, first).x;
}

export function selectionGeometry(index: CommandGeometryIndex, selection: PresentationCommandSelected | null): {
  caret: CommandRect | null; selectionRects: CommandRect[];
} {
  const entry = findLine(index, selection);
  if (!entry || !selection) return { caret: null, selectionRects: [] };
  const caret = transformedRect(entry.transforms, {
    x: stopX(entry.line, selection.focus), y: entry.line.y, width: 1, height: entry.line.height,
  });
  const selectionRects: CommandRect[] = [];
  const start = Math.min(selection.anchor, selection.focus);
  const end = Math.max(selection.anchor, selection.focus);
  for (const item of index.lines.get(targetKey(selection.shapeId, selection.storyId)) ?? []) {
    if (end <= item.line.start || start >= item.line.end) continue;
    const a = stopX(item.line, Math.max(start, item.line.start));
    const b = stopX(item.line, Math.min(end, item.line.end));
    selectionRects.push(transformedRect(item.transforms, {
      x: Math.min(a, b), y: item.line.y, width: Math.abs(b - a), height: item.line.height,
    }));
  }
  return { caret, selectionRects };
}

export function hitGeometry(index: Pick<CommandHitGeometryIndex, 'hits'>, point: { x: number; y: number },
  target?: PresentationCommandSelection | null, borderScale?: number): (HitTestResult & { affinity?: PresentationCommandSelection['affinity'] }) | null {
  for (let i = index.hits.length - 1; i >= 0; i -= 1) {
    const entry = index.hits[i];
    let local = point;
    let clipped = false;
    for (const primitive of entry.transforms) {
      local = transformPoint(primitive, local, true);
      if (entry.clips.includes(primitive) && !contains(primitive, local)) {
        clipped = true;
        if (!target) break;
      }
    }
    const primitive = entry.primitive;
    if (target && (primitive.shapeId !== target.shapeId || primitive.storyId !== target.storyId)) continue;
    if ((!target && (clipped || !contains(primitive, local))) || !primitive.shapeId) continue;
    if (!target && borderScale !== undefined) {
      const band = Math.min(6 / borderScale, primitive.w / 3, primitive.h / 3);
      if (local.x - primitive.x <= band || primitive.x + primitive.w - local.x <= band ||
        local.y - primitive.y <= band || primitive.y + primitive.h - local.y <= band) {
        return { kind: 'shape', shapeId: primitive.shapeId };
      }
    }
    if (primitive.kind === 'textBox' && primitive.storyId) {
      const nearest = target || entry.clips.some((parent) => 'kind' in parent && parent.kind === 'table');
      const columns = primitive.lines?.some((line, index, lines) => index > 0 &&
        (line.columnOffset ?? 0) !== (lines[index - 1].columnOffset ?? 0));
      let nearestDistance = Number.POSITIVE_INFINITY;
      const line = primitive.lines?.reduce<CaretLine | undefined>((best, candidate) => {
        if (!nearest && (local.y < candidate.y || local.y > candidate.y + candidate.height)) return best;
        const distance = (value: CaretLine) => {
          let left = value.x;
          let right = value.x;
          for (const stop of value.caretStops) {
            left = Math.min(left, stop.x);
            right = Math.max(right, stop.x);
          }
          const dx = Math.max(left - local.x, local.x - right, 0);
          const dy = Math.max(value.y - local.y, local.y - value.y - value.height, 0);
          return columns ? dx * dx + dy * dy : dy;
        };
        const candidateDistance = distance(candidate);
        if (!best || candidateDistance < nearestDistance) {
          nearestDistance = candidateDistance;
          return candidate;
        }
        return best;
      }, undefined);
      const first = line?.caretStops[0];
      if (line && first) {
        const stop = line.caretStops.reduce((best, candidate) =>
          Math.abs(candidate.x - local.x) < Math.abs(best.x - local.x) ? candidate : best, first);
        return { kind: 'text', shapeId: primitive.shapeId, storyId: primitive.storyId, position: stop.position,
          affinity: stop.position === line.end ? 'upstream' : 'downstream' };
      }
    }
    return { kind: 'shape', shapeId: primitive.shapeId };
  }
  return null;
}

function contains(primitive: GeometryBounds, point: { x: number; y: number }): boolean {
  return point.x >= primitive.x && point.x <= primitive.x + primitive.w &&
    point.y >= primitive.y && point.y <= primitive.y + primitive.h;
}

interface CellHit {
  primitive: { kind: string; shapeId?: string; storyId?: string };
  clips: readonly object[];
  textLength?: number;
}

/** Tab and Shift+Tab inside a table: the whole text of the next or previous cell in reading order;
 *  the selection stays put past the last or first cell, and null means it is not in a table cell. */
function cellGeometry(hits: readonly CellHit[], selection: PresentationCommandSelected | null,
  backward: boolean): PresentationCommandSelection | null {
  if (!selection || selection.kind === 'object') return null;
  const isTable = (clip: object) => 'kind' in clip && clip.kind === 'table';
  const table = hits.find(({ primitive }) => primitive.kind === 'textBox' &&
    primitive.shapeId === selection.shapeId && primitive.storyId === selection.storyId)?.clips.filter(isTable).pop();
  if (!table) return null;
  const cells: { storyId: string; end: number }[] = [];
  for (const { primitive, clips, textLength } of hits) {
    const storyId = primitive.storyId;
    if (primitive.kind !== 'textBox' || !storyId || primitive.shapeId !== selection.shapeId ||
      clips.filter(isTable).pop() !== table || cells.some((cell) => cell.storyId === storyId)) continue;
    cells.push({ storyId, end: textLength ?? 0 });
  }
  const target = cells[cells.findIndex((cell) => cell.storyId === selection.storyId) + (backward ? -1 : 1)];
  if (!target) return { ...selection };
  return { slideId: selection.slideId, shapeId: selection.shapeId, storyId: target.storyId, anchor: 0,
    focus: target.end, affinity: target.end > 0 ? 'upstream' : 'downstream' };
}

export function navigateGeometry(
  index: NavigationIndex & { hits?: readonly CellHit[] }, selection: PresentationCommandSelected | null,
  motion: PresentationCommandMotion, extend: boolean,
  text?: string,
): PresentationCommandSelection | null {
  if (motion === 'nextCell' || motion === 'previousCell') {
    return index.hits ? cellGeometry(index.hits, selection, motion === 'previousCell') : null;
  }
  const entry = findLine(index, selection);
  if (!entry || !selection || selection.kind === 'object') return null;
  const lines = index.lines.get(targetKey(selection.shapeId, selection.storyId)) ?? [];
  let position = selection.focus;
  let destination = entry;
  let goalX: number | undefined;
  if (!extend && selection.anchor !== selection.focus && (motion === 'left' || motion === 'right')) {
    position = motion === 'left' ? Math.min(selection.anchor, selection.focus) : Math.max(selection.anchor, selection.focus);
  } else if (motion === 'wordLeft' || motion === 'wordRight') {
    if (text === undefined) return null;
    position = wordBoundary(text, position, motion === 'wordLeft' ? -1 : 1);
  } else if (motion === 'storyStart') {
    destination = lines[0]!;
    position = 0;
  } else if (motion === 'storyEnd') {
    destination = lines[lines.length - 1]!;
    position = text?.length ?? destination.line.end;
  } else if (motion === 'lineStart' || motion === 'lineEnd') {
    const stops = entry.line.caretStops;
    if (!stops.length) return null;
    position = motion === 'lineStart' ? stops[0]!.position : stops[stops.length - 1]!.position;
  }
  else if (motion === 'left' || motion === 'right') {
    const stops = text === undefined ? lines.flatMap(({ line }) => line.caretStops.map((stop) => stop.position))
      : [...graphemeBoundaries(text)];
    const candidates = stops.filter((stop) => motion === 'left' ? stop < position : stop > position);
    if (candidates.length) position = motion === 'left' ? Math.max(...candidates) : Math.min(...candidates);
  } else if (motion === 'up' || motion === 'down') {
    const caretLines = lines.map(({ line }) => line);
    goalX = selection.goalX ?? stopX(entry.line, position) - caretColumnOffset(caretLines, lines.indexOf(entry));
    const moved = verticalCaretMove(caretLines,
      { position, lineIndex: lines.indexOf(entry) }, motion, goalX);
    destination = lines[moved.lineIndex]!;
    position = moved.position;
  }
  if (motion === 'left' || motion === 'right' || motion === 'wordLeft' || motion === 'wordRight') {
    destination = findLine(index, { ...selection, focus: position,
      affinity: motion === 'left' || motion === 'wordLeft' ? 'upstream' : 'downstream' }) ?? entry;
  }
  const result = { ...selection, anchor: extend ? selection.anchor : position, focus: position,
    affinity: position === destination.line.end ? 'upstream' as const : 'downstream' as const };
  if (goalX === undefined) delete result.goalX;
  else result.goalX = goalX;
  return result;
}
