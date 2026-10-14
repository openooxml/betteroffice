import type { YrsCellLoc, YrsLoc, YrsStickyPosition, YrsStoryRange } from './index';
import type { ResidentEngineSession } from './residentEngineSession';
import type { DocxCommandSelection } from './residentCommandProtocol';
import type { RetainedFrame } from '../layout/render/frameDelta';
import { DisplayPositionIndex } from './displayPositionIndex';
import { createYrsPositionProjection } from './yrsPositionProjection';
import { createDisplayListQueries } from '../layout/render/displayListQueries';
import { deriveDisplayListSelectedCellRects, type DisplayListTableRegion } from '../layout/render/displayListTables';
import { loadRustDisplayListQueryEngine } from '../layout/render/rustDisplayList';
import { readResidentRangeHighlights } from './residentRangeHighlights';

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

const registries = new WeakMap<ResidentEngineSession, Map<string, ObjectTarget>>();

export function resetResidentObjectHandles(engine: ResidentEngineSession): void {
  registries.set(engine, new Map());
}

function registry(engine: ResidentEngineSession): Map<string, ObjectTarget> {
  let value = registries.get(engine);
  if (!value) { value = new Map(); registries.set(engine, value); }
  return value;
}

function identity(position: YrsStickyPosition): string {
  return JSON.stringify([position.story, Array.from(position.encoded)]);
}

function issue(engine: ResidentEngineSession, kind: ObjectTarget['kind'], at: YrsLoc, identityAt = at): string {
  const position = engine.encodeStickyPosition(identityAt);
  for (const [handle, target] of registry(engine)) {
    if (target.kind === kind && target.story === at.story && identity(target.position) === identity(position) &&
      (kind !== 'text' || target.paraId === at.paraId)) return handle;
  }
  const handle = crypto.randomUUID();
  registry(engine).set(handle, { kind, position, paraId: at.paraId, story: at.story });
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
  const target = registry(engine).get(handle);
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

export async function readResidentFrameGeometry(
  engine: ResidentEngineSession, frame: RetainedFrame, selection: DocxCommandSelection | null, layoutRevision: number
): Promise<ResidentFrameGeometry> {
  const native = await loadRustDisplayListQueryEngine();
  const queries = createDisplayListQueries(frame.displayList, native);
  const index = new DisplayPositionIndex({ ...engine.geometryReader, selectionText: engine.selectionText });
  const objects: ResidentObjectGeometry[] = [];
  const ranges: YrsStoryRange[] = [];
  if (selection && selection.anchor.story === selection.focus.story) ranges.push({
    story: selection.anchor.story, start: selection.anchor, end: selection.focus,
  });
  const highlights = await readResidentRangeHighlights(engine, frame, ranges, layoutRevision);
  try {
    const roots = engine.geometryReader.storyIds().filter(story => story === 'body' || /^(hf|fn|en):[^:]+$/.test(story));
    for (const root of roots) {
      const projection = index.projection(root);
      if (!projection) continue;
      const [prefix, id] = root.split(':');
      const rects = (from: number, to: number) => (root === 'body' ? queries.rangeRects(from, to)
        : prefix === 'hf' ? [...queries.hfRangeRects('header', id!, from, to), ...queries.hfRangeRects('footer', id!, from, to)]
        : queries.noteRangeRects(prefix === 'fn' ? 'footnote' : 'endnote', Number(id), from, to))
        .map(rect => ({ ...rect, pageId: frame.pages[rect.pageIndex]!.pageId.toString() }));
      const tables = projection.allTables();
      const stories = new Set([root, ...tables.flatMap(table => table.cells.map(cell => cell.story))]);
      const cells = tables.flatMap(table => table.cells.map(cell => ({
        tableKey: String(table.start), row: cell.row, col: cell.column,
        rowSpan: cell.rowspan, colSpan: cell.colspan,
      })));
      const tableKeyOf = (position: number | undefined) => position === undefined ? null
        : projection.tableAtPosition(position)?.start.toString() ?? null;
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
        const at = engine.geometryReader.paragraphSpans(cell.story)[0];
        const tableAt = projection.blockBoundaryAt(table.start);
        const boxes = cellGeometry.get(`${table.start}:${cell.row}:${cell.column}`) ?? [];
        if (at && tableAt && boxes.length) objects.push({ kind: 'cell', story: cell.story,
          handle: issue(engine, 'cell', { story: cell.story, paraId: at.paraId, offset: 0 }, tableAt), rects: boxes });
      }
      for (const story of stories) {
        const spans = engine.geometryReader.paragraphSpans(story);
        for (let p = 0; p < spans.length; p += 1) {
          const paraId = spans[p]!.paraId;
          let offset = 0;
          for (const segment of engine.paragraphSegments(story, [p])[0] ?? []) {
            const at = { story, paraId, offset };
            offset += segment.kind === 'text' ? segment.text.length : 1;
            if (segment.kind !== 'embed' || segment.embedKind !== 'image') continue;
            const position = index.positionOf(at, root);
            if (position === null) continue;
            const image = root === 'body' || prefix === 'hf' ? queries.imageByPos(position, prefix === 'hf' ? 'header' : 'body', id)
              ?? (prefix === 'hf' ? queries.imageByPos(position, 'footer', id) : null) : null;
            if (image) objects.push({ kind: 'image', story, handle: issue(engine, 'image', at),
              rects: [{ ...image.rect, pageId: frame.pages[image.rect.pageIndex]!.pageId.toString() }] });
          }
          const from = index.positionOf({ story, paraId, offset: 0 }, root);
          const to = index.positionOf({ story, paraId, offset: spans[p]!.length }, root);
          const boxes = from !== null && to !== null ? rects(from, to) : [];
          if (boxes.length) objects.push({ kind: 'text', story,
            handle: issue(engine, 'text', { story, paraId, offset: 0 }, { story, paraId, offset: spans[p]!.length }), rects: boxes });
        }
      }
    }
    const state = queries.sourceState();
    if (state.status === 'error') throw new Error(`Object layout is unavailable: ${String(state.error)}`);
    return { version: engine.proposalEngine.version(), frameEpoch: frame.frameEpoch, layoutRevision,
      selection: highlights.rects[0] ?? [], objects, unbuiltPages: highlights.unbuiltPages };
  } finally { queries.dispose(); }
}
