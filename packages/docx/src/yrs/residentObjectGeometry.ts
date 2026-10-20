import type { YrsCellLoc, YrsLoc, YrsStickyPosition, YrsStorySegment, YrsTableRange } from './index';
import type { ResidentEngineSession } from './residentEngineSession';
import type { DocxCommandSelection } from './residentCommandProtocol';
import type { RetainedFrame } from '../layout/render/frameDelta';
import { createYrsPositionProjection } from './yrsPositionProjection';
import type { DisplayListHitRegion, DisplayListRect } from '../layout/render/displayListQueries';
import { deriveDisplayListSelectedCellRects, type DisplayListSelectedCellSpec, type DisplayListTableRegion } from '../layout/render/displayListTables';
import { findImagePrimitiveByDocPos } from '../layout/render/displayListImages';
import { rootOf } from './residentCommandTableSelection';
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
  const root = rootOf(story);
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

type ResidentParagraphGeometry = ReturnType<ResidentEngineSession['residentGeometry']>['paragraphs'][number];

/** Display position of story offset `offset` in `paragraph`, past the content of its inline controls. */
function displayPosition(paragraph: ResidentParagraphGeometry, offset: number): number {
  let position = paragraph.from + Math.max(0, offset - paragraph.leading);
  for (const control of paragraph.controls ?? []) if (control.offset < offset) position += control.to - control.from - 1;
  return position;
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

/** Page-local boxes of a multi-cell range; null when no multi-cell range is active. */
export function readResidentCellSelectionRects(
  engine: ResidentEngineSession, frame: RetainedFrame, range: YrsTableRange | null
): DisplayListRect[] | null {
  if (!range || range.anchor.story !== range.head.story || range.anchor.tableIndex !== range.head.tableIndex ||
    (range.anchor.row === range.head.row && range.anchor.column === range.head.column)) return null;
  const root = rootOf(range.anchor.story);
  if (root !== 'body' || !engine.geometryReader.hasStory(root)) return [];
  const table = createYrsPositionProjection(engine.geometryReader, root)?.allTables()
    .find((entry) => entry.story === range.anchor.story && entry.tableIndex === range.anchor.tableIndex);
  const covering = (at: YrsCellLoc) => table?.cells.find((cell) => cell.row <= at.row && at.row < cell.row + cell.rowspan &&
    cell.column <= at.column && at.column < cell.column + cell.colspan);
  const anchorCell = covering(range.anchor);
  const headCell = covering(range.head);
  if (!anchorCell || !headCell) return [];
  const ranges: Array<[number, number]> = [];
  for (const page of frame.displayList.pages) {
    if (page.unbuilt) continue;
    let from = Infinity;
    let to = -Infinity;
    for (const primitive of page.primitives) {
      from = Math.min(from, primitive.fragmentDocStart ?? primitive.docStart ?? Infinity);
      to = Math.max(to, primitive.fragmentDocEnd ?? primitive.docEnd ?? primitive.docStart ?? -Infinity);
    }
    if (Number.isFinite(from) && Number.isFinite(to)) ranges.push([from, to]);
  }
  if (!ranges.length) return [];
  const projected = engine.residentGeometry(root, ranges);
  const display = projected.tables.find((entry) =>
    entry.cells.some((cell) => cell.story === anchorCell.story) &&
    entry.cells.some((cell) => cell.story === headCell.story));
  if (!display) return [];
  const top = Math.min(range.anchor.row, range.head.row);
  const bottom = Math.max(range.anchor.row, range.head.row);
  const left = Math.min(range.anchor.column, range.head.column);
  const right = Math.max(range.anchor.column, range.head.column);
  const specs: DisplayListSelectedCellSpec[] = (table?.cells ?? [])
    .filter((cell) => cell.row <= bottom && cell.row + cell.rowspan - 1 >= top &&
      cell.column <= right && cell.column + cell.colspan - 1 >= left)
    .map((cell) => ({
      tableKey: String(display.from), row: cell.row, col: cell.column, rowSpan: cell.rowspan, colSpan: cell.colspan,
    }));
  const tableKeyOf = (position: number | undefined) => position === undefined ? null
    : projected.tables.find((entry) => entry.from <= position && position < entry.to)?.from.toString() ?? null;
  return deriveDisplayListSelectedCellRects(frame.displayList, specs, tableKeyOf, { kind: 'body' })
    .filter((box) => frame.pages[box.pageIndex] && !frame.pages[box.pageIndex]!.page.unbuilt)
    .map((box) => ({ pageIndex: box.pageIndex, x: box.x, y: box.y, width: box.w, height: box.h }));
}

export async function readResidentFrameGeometry(
  engine: ResidentEngineSession, frame: RetainedFrame, selection: DocxCommandSelection | null, layoutRevision: number
): Promise<ResidentFrameGeometry> {
  const objects: ResidentObjectGeometry[] = [];
  let selected: ResidentObjectRect[] = [];
  const selectionRoot = selection && selection.anchor.story === selection.focus.story
    ? rootOf(selection.anchor.story) : null;
  const located = (rects: DisplayListRect[]): ResidentObjectRect[] => rects
    .filter(rect => frame.pages[rect.pageIndex] && !frame.pages[rect.pageIndex]!.page.unbuilt)
    .map(rect => ({ ...rect, pageId: frame.pages[rect.pageIndex]!.pageId.toString() }));
  for (const [root, ranges] of [...builtStoryRanges(frame)].sort(([a], [b]) => a < b ? -1 : a > b ? 1 : 0)) {
    if (!engine.geometryReader.hasStory(root)) continue;
    const projected = engine.residentGeometry(root, ranges, selectionRoot === root && selection
      ? [selection.anchor, selection.focus] : undefined);
    const [prefix, id] = root.split(':');
    const spans: Array<[number, number]> = projected.paragraphs.map(paragraph => [paragraph.from, paragraph.to]);
    const range = projected.selection && projected.selection[0] !== projected.selection[1]
      ? [Math.min(...projected.selection), Math.max(...projected.selection)] as [number, number] : null;
    if (range) spans.push(range);
    const regions: Array<[DisplayListHitRegion, string]> = root === 'body' ? [['body', '']]
      : prefix === 'hf' ? [['header', id!], ['footer', id!]] : [[prefix === 'fn' ? 'footnote' : 'endnote', String(Number(id))]];
    const answered = regions.map(([region, part]) => engine.residentRangeRects(region, part, spans));
    const rects = (index: number) => located(answered.flatMap(batch => batch[index] ?? []));
    if (range) selected = rects(spans.length - 1);
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
    const tableRegions: DisplayListTableRegion[] = prefix === 'hf'
      ? [{ kind: 'header', rId: id }, { kind: 'footer', rId: id }] : [{ kind: 'body' }];
    const cellGeometry = new Map<string, ResidentObjectRect[]>();
    for (const region of tableRegions) for (const box of deriveDisplayListSelectedCellRects(list, cells, tableKeyOf, region)) {
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
    const imageStarts: number[] = [];
    if (root === 'body' || prefix === 'hf') for (const page of frame.displayList.pages) {
      const bands = root === 'body' ? [page.primitives]
        : [page.header, page.footer].flatMap(band => band && band.rId === id ? [band.primitives] : []);
      for (const primitive of bands.flat()) if (primitive.kind === 'image' && primitive.docStart !== undefined) imageStarts.push(primitive.docStart);
    }
    const byStory = new Map<string, ResidentParagraphGeometry[]>();
    for (const paragraph of projected.paragraphs) {
      if (!imageStarts.some(start => start >= paragraph.from && start <= paragraph.to)) continue;
      const group = byStory.get(paragraph.story);
      if (group) group.push(paragraph);
      else byStory.set(paragraph.story, [paragraph]);
    }
    const segments = new Map<ResidentParagraphGeometry, YrsStorySegment[]>();
    for (const [story, paragraphs] of byStory) {
      const read = engine.paragraphSegments(story, paragraphs.map(paragraph => paragraph.index));
      paragraphs.forEach((paragraph, index) => segments.set(paragraph, read[index] ?? []));
    }
    for (const [index, paragraph] of projected.paragraphs.entries()) {
      const { story, paraId } = paragraph;
      let offset = 0;
      for (const segment of segments.get(paragraph) ?? []) {
        const at = { story, paraId, offset };
        offset += segment.kind === 'text' ? segment.text.length : 1;
        if (segment.kind !== 'embed' || segment.embedKind !== 'image') continue;
        const position = displayPosition(paragraph, at.offset);
        const image = root === 'body' || prefix === 'hf'
          ? findImagePrimitiveByDocPos(frame.displayList, position, prefix === 'hf' ? 'header' : 'body', id)
            ?? (prefix === 'hf' ? findImagePrimitiveByDocPos(frame.displayList, position, 'footer', id) : null) : null;
        if (image) objects.push({ kind: 'image', story, handle: issue(engine, 'image', at),
          rects: [{ pageIndex: image.pageIndex, x: image.primitive.x, y: image.primitive.y, width: image.primitive.w,
            height: image.primitive.h, pageId: frame.pages[image.pageIndex]!.pageId.toString() }] });
      }
      const boxes = rects(index);
      if (boxes.length) objects.push({ kind: 'text', story,
        handle: issue(engine, 'text', { story, paraId, offset: 0 }, { story, paraId, offset: paragraph.length }), rects: boxes });
    }
  }
  return { version: engine.proposalEngine.version(), frameEpoch: frame.frameEpoch, layoutRevision,
    selection: selected, objects, unbuiltPages: frame.pages.filter(page => page.page.unbuilt).map(page => page.pageIndex) };
}
