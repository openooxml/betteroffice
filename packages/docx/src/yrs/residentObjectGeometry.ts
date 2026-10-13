import type { YrsCellLoc, YrsLoc, YrsStickyPosition } from './index';
import type { ResidentEngineSession } from './residentEngineSession';
import type { DocxCommandSelection } from './residentCommandProtocol';
import type { RetainedFrame } from '../layout/render/frameDelta';
import { createYrsPositionProjection } from './yrsPositionProjection';
import { createDisplayListQueries } from '../layout/render/displayListQueries';
import { deriveDisplayListSelectedCellRects, type DisplayListTableRegion } from '../layout/render/displayListTables';
import { loadRustDisplayListQueryEngine } from '../layout/render/rustDisplayList';
import type { DisplayPrimitive } from '../layout/render/displayList';

export interface ResidentObjectRect {
  pageId: string;
  pageIndex: number;
  x: number;
  y: number;
  width: number;
  height: number;
}

export interface ResidentObjectGeometry {
  handle: string;
  kind: 'text' | 'cell' | 'image';
  story: string;
  rects: ResidentObjectRect[];
}

export interface ResidentFrameGeometry {
  version: string;
  frameEpoch: number;
  layoutRevision: number;
  selection: ResidentObjectRect[];
  objects: ResidentObjectGeometry[];
  unbuiltPages: number[];
}

interface ObjectTarget {
  kind: ResidentObjectGeometry['kind'];
  position: YrsStickyPosition;
  paraId: string;
  story: string;
}

interface ObjectRegistry { targets: Map<string, ObjectTarget>; identities: Map<string, string> }
const registries = new WeakMap<ResidentEngineSession, ObjectRegistry>();

export function resetResidentObjectHandles(engine: ResidentEngineSession): void {
  registries.set(engine, { targets: new Map(), identities: new Map() });
}

function registry(engine: ResidentEngineSession): ObjectRegistry {
  let value = registries.get(engine);
  if (!value) { value = { targets: new Map(), identities: new Map() }; registries.set(engine, value); }
  return value;
}

function identity(position: YrsStickyPosition): string {
  return JSON.stringify([position.story, Array.from(position.encoded)]);
}

function issue(engine: ResidentEngineSession, kind: ObjectTarget['kind'], at: YrsLoc, identityAt = at): string {
  const position = engine.encodeStickyPosition(identityAt);
  const entries = registry(engine);
  const key = JSON.stringify([kind, at.story, identity(position), kind === 'text' ? at.paraId : null]);
  const existing = entries.identities.get(key);
  if (existing) return existing;
  const handle = crypto.randomUUID();
  entries.targets.set(handle, { kind, position, paraId: at.paraId, story: at.story });
  entries.identities.set(key, handle);
  return handle;
}

function resolveCell(engine: ResidentEngineSession, story: string): { at: YrsLoc; cell: YrsCellLoc } | null {
  const root = story === 'body' || story.startsWith('body:') ? 'body' : story.split(':').slice(0, 2).join(':');
  const projection = createYrsPositionProjection(engine.geometryReader, root);
  if (!projection) return null;
  for (const table of projection.allTables()) {
    const cell = table.cells.find(candidate => candidate.story === story);
    if (!cell) continue;
    const at = projection.blockBoundaryAt(table.start);
    return at ? { at, cell: { story: table.story, tableIndex: table.tableIndex, row: cell.row, column: cell.column } } : null;
  }
  return null;
}

export function resolveResidentObject(engine: ResidentEngineSession, handle: string): (YrsLoc & { kind: ObjectTarget['kind']; cell?: YrsCellLoc }) | null {
  const target = registry(engine).targets.get(handle);
  if (!target) return null;
  if (!engine.geometryReader.hasStory(target.story) || !engine.geometryReader.hasStory(target.position.story)) return null;
  if (target.kind !== 'image') {
    const paragraph = engine.geometryReader.paragraphSpans(target.story)
      .find(span => target.kind === 'cell' || span.paraId === target.paraId);
    if (!paragraph) return null;
    const cell = target.kind === 'cell' ? resolveCell(engine, target.story) : null;
    if (target.kind === 'cell' && !cell) return null;
    const at = cell?.at ?? { story: target.story, paraId: paragraph.paraId, offset: paragraph.length };
    if (identity(engine.encodeStickyPosition(at)) !== identity(target.position)) return null;
    return { story: target.story, paraId: paragraph.paraId, offset: 0, kind: target.kind, ...(cell ? { cell: cell.cell } : {}) };
  }
  const at = engine.resolveStickyPosition(target.position);
  if (!at || !engine.geometryReader.hasStory(at.story)) return null;
  if (identity(engine.encodeStickyPosition(at)) !== identity(target.position)) return null;
  if (target.kind === 'image') {
    let offset = 0;
    const index = engine.geometryReader.paragraphSpans(at.story).findIndex(span => span.paraId === at.paraId);
    if (index < 0) return null;
    const found = engine.paragraphSegments(at.story, [index])[0]?.find(segment => {
      const start = offset;
      offset += segment.kind === 'text' ? segment.text.length : 1;
      return start === at.offset;
    });
    if (found?.kind !== 'embed' || found.embedKind !== 'image') return null;
  }
  return { ...at, kind: target.kind };
}

function builtStoryRanges(frame: RetainedFrame): Map<string, Array<[number, number]>> {
  const roots = new Map<string, Array<[number, number]>>();
  const add = (root: string, primitives: readonly DisplayPrimitive[], span?: [number, number]) => {
    let from = span?.[0] ?? Infinity;
    let to = span?.[1] ?? -Infinity;
    for (const primitive of primitives) {
      from = Math.min(from, primitive.fragmentDocStart ?? primitive.docStart ?? Infinity);
      to = Math.max(to, primitive.fragmentDocEnd ?? primitive.docEnd ?? primitive.docStart ?? -Infinity);
    }
    if (!Number.isFinite(from) || !Number.isFinite(to)) return;
    const ranges = roots.get(root) ?? [];
    ranges.push([from, to]);
    roots.set(root, ranges);
  };
  for (const page of frame.displayList.pages) {
    if (page.unbuilt) continue;
    add('body', page.primitives, page.positionSpan);
    for (const region of [page.header, page.footer]) if (region) add(`hf:${region.rId}`, region.primitives);
    for (const area of page.noteAreas ?? []) {
      const kind = area.kind ?? 'footnote';
      const notes = new Map<string, DisplayPrimitive[]>();
      for (const primitive of area.primitives ?? []) {
        const id = primitive.groupId?.match(/^(footnote|endnote)-(\d+)$/)?.[2];
        if (id === undefined) continue;
        const primitives = notes.get(id) ?? [];
        primitives.push(primitive);
        notes.set(id, primitives);
      }
      for (const [id, primitives] of notes) add(`${kind === 'footnote' ? 'fn' : 'en'}:${id}`, primitives);
    }
  }
  return roots;
}

export async function readResidentFrameGeometry(
  engine: ResidentEngineSession, frame: RetainedFrame, selection: DocxCommandSelection | null, layoutRevision: number
): Promise<ResidentFrameGeometry> {
  const native = await loadRustDisplayListQueryEngine();
  const queries = createDisplayListQueries(frame.displayList, native);
  const objects: ResidentObjectGeometry[] = [];
  let selected: ResidentObjectRect[] = [];
  const selectionRoot = selection && selection.anchor.story === selection.focus.story
    ? selection.anchor.story === 'body' || selection.anchor.story.startsWith('body:') ? 'body'
      : selection.anchor.story.split(':').slice(0, 2).join(':') : null;
  try {
    for (const [root, ranges] of [...builtStoryRanges(frame)].sort(([a], [b]) => a < b ? -1 : a > b ? 1 : 0)) {
      if (!engine.geometryReader.hasStory(root)) continue;
      const projected = engine.residentGeometry(root, ranges, selectionRoot === root && selection
        ? [selection.anchor, selection.focus] : undefined);
      const [prefix, id] = root.split(':');
      const rects = (from: number, to: number) => (root === 'body' ? queries.rangeRects(from, to)
        : prefix === 'hf' ? [...queries.hfRangeRects('header', id!, from, to), ...queries.hfRangeRects('footer', id!, from, to)]
        : queries.noteRangeRects(prefix === 'fn' ? 'footnote' : 'endnote', Number(id), from, to))
        .map(rect => ({ ...rect, pageId: frame.pages[rect.pageIndex]!.pageId.toString() }));
      if (projected.selection) {
        const [a, b] = projected.selection;
        if (a !== b) selected = rects(Math.min(a, b), Math.max(a, b));
      }
      const tables = projected.tables;
      const cells = tables.flatMap(table => table.cells.map(cell => ({
        tableKey: String(table.from), row: cell.row, col: cell.column,
        rowSpan: cell.rowspan, colSpan: cell.colspan,
      })));
      const tableKeyOf = (position: number | undefined) => position === undefined ? null
        : tables.find(table => table.from <= position && position < table.to)?.from.toString() ?? null;
      const noteKind = prefix === 'fn' ? 'footnote' : 'endnote';
      const list = prefix === 'fn' || prefix === 'en' ? { ...frame.displayList,
        pages: frame.displayList.pages.map(page => ({ ...page,
          primitives: (page.noteAreas ?? []).filter(area => (area.kind ?? 'footnote') === noteKind)
            .flatMap(area => (area.primitives ?? []).filter(primitive => primitive.groupId === `${noteKind}-${id}`)),
        })),
      } : frame.displayList;
      const regions: DisplayListTableRegion[] = prefix === 'hf'
        ? [{ kind: 'header', rId: id }, { kind: 'footer', rId: id }] : [{ kind: 'body' }];
      const cellGeometry = new Map<string, ResidentObjectRect[]>();
      for (const region of regions) for (const box of deriveDisplayListSelectedCellRects(list, cells, tableKeyOf, region)) {
        const key = `${box.tableKey}:${box.row}:${box.col}`;
        const bounds = cellGeometry.get(key) ?? [];
        bounds.push({ pageId: frame.pages[box.pageIndex]!.pageId.toString(), pageIndex: box.pageIndex,
          x: box.x, y: box.y, width: box.w, height: box.h });
        cellGeometry.set(key, bounds);
      }
      for (const table of tables) for (const cell of table.cells) {
        const boxes = cellGeometry.get(`${table.from}:${cell.row}:${cell.column}`) ?? [];
        if (boxes.length) objects.push({ kind: 'cell', story: cell.story,
          handle: issue(engine, 'cell', { story: cell.story, paraId: cell.paraId, offset: 0 }, table.at), rects: boxes });
      }
      for (const paragraph of projected.paragraphs) {
        const { story, paraId } = paragraph;
        let offset = 0;
        for (const segment of engine.paragraphSegments(story, [paragraph.index])[0] ?? []) {
          const at = { story, paraId, offset };
          offset += segment.kind === 'text' ? segment.text.length : 1;
          if (segment.kind !== 'embed' || segment.embedKind !== 'image') continue;
          const position = paragraph.from + Math.max(0, at.offset - paragraph.leading);
          const image = root === 'body' || prefix === 'hf' ? queries.imageByPos(position, prefix === 'hf' ? 'header' : 'body', id)
            ?? (prefix === 'hf' ? queries.imageByPos(position, 'footer', id) : null) : null;
          if (image) objects.push({ kind: 'image', story, handle: issue(engine, 'image', at),
            rects: [{ ...image.rect, pageId: frame.pages[image.rect.pageIndex]!.pageId.toString() }] });
        }
        const boxes = rects(paragraph.from, paragraph.to);
        if (boxes.length) objects.push({ kind: 'text', story,
          handle: issue(engine, 'text', { story, paraId, offset: 0 }, { story, paraId, offset: paragraph.length }), rects: boxes });
      }
    }
    const state = queries.sourceState();
    if (state.status === 'error') throw new Error(`Object layout is unavailable: ${String(state.error)}`);
    return { version: engine.proposalEngine.version(), frameEpoch: frame.frameEpoch, layoutRevision,
      selection: selected, objects, unbuiltPages: frame.pages.filter(page => page.page.unbuilt).map(page => page.pageIndex) };
  } finally { queries.dispose(); }
}
