import { isEngineTrap } from '../../../../shared/office-session/trap';
import type { YrsCellLoc, YrsLoc, YrsTableRange } from './index';
import type { DocxCommandSelection, DocxWorkerCommandEngine } from './residentCommandProtocol';
import { tableAnchors, type TablePayload } from './storyPlainText';
import { createYrsPositionProjection } from './yrsPositionProjection';

export type TableSelectionTarget = 'table' | 'row' | 'column';

export function rootOf(story: string): string {
  return story === 'body' || story.startsWith('body:') ? 'body' : story.split(':').slice(0, 2).join(':');
}

export function readCellSelection(engine: Pick<DocxWorkerCommandEngine, 'cellSelection'>): YrsTableRange | null {
  try {
    return engine.cellSelection();
  } catch (error) {
    if (isEngineTrap(error)) throw error;
    return null;
  }
}

type ProjectedCellTable = { story: string; tableIndex: number; cells: { story: string; row: number; column: number }[] };

/** The tables of `root` as the position projection lays them out: only what cell location reads. */
function projectedCellTables(engine: DocxWorkerCommandEngine, root: string): ProjectedCellTable[] | null {
  const projection = createYrsPositionProjection(engine.geometryReader, root);
  if (!projection) return null;
  return projection.allTables().map((table) => ({
    story: table.story,
    tableIndex: table.tableIndex,
    cells: table.cells.map((cell) => ({ story: cell.story, row: cell.row, column: cell.column })),
  }));
}

function cellIn(tables: readonly ProjectedCellTable[] | null, story: string): YrsCellLoc | null {
  if (!tables) return null;
  let best: YrsCellLoc | null = null;
  for (const table of tables) {
    const cell = table.cells.find((candidate) => candidate.story === story || story.startsWith(`${candidate.story}:`));
    if (!cell) continue;
    if (!best || table.story.length > best.story.length) {
      best = { story: table.story, tableIndex: table.tableIndex, row: cell.row, column: cell.column };
    }
  }
  return best;
}

const projectedTables = new WeakMap<object, Map<string, { key: string; tables: ProjectedCellTable[] | null }>>();
let shadowCheck = false;
const cacheCounts = { built: 0, reused: 0, uncached: 0 };

/** @internal Every resolveLiveCell call also builds a fresh projection and throws when the answers differ. */
export function setCellLocationShadowCheck(enabled: boolean): void {
  shadowCheck = enabled;
}

/** @internal How often resolveLiveCell rebuilt, reused or could not cache the projected tables. */
export function cellLocationCacheCounts(): { built: number; reused: number; uncached: number } {
  return { ...cacheCounts };
}

/**
 * The innermost table cell holding `story`, as the position projection of its root finds it. The
 * projected tables are kept per engine and root until the engine's structure key changes: a key
 * that changes with every embed, table row/cell/story change, story creation or removal, and with
 * the document instance. Engines without a key project on every call.
 */
export function resolveLiveCell(engine: DocxWorkerCommandEngine, story: string): YrsCellLoc | null {
  const root = rootOf(story);
  if (story === root) return null;
  const key = typeof engine.cellStructureKey === 'function' ? engine.cellStructureKey() : null;
  let tables: ProjectedCellTable[] | null;
  if (key === null) {
    cacheCounts.uncached++;
    tables = projectedCellTables(engine, root);
  } else {
    let roots = projectedTables.get(engine);
    if (!roots) projectedTables.set(engine, roots = new Map());
    const cached = roots.get(root);
    if (cached?.key === key) {
      cacheCounts.reused++;
      tables = cached.tables;
    } else {
      cacheCounts.built++;
      tables = projectedCellTables(engine, root);
      roots.set(root, { key, tables });
    }
  }
  const located = cellIn(tables, story);
  if (shadowCheck && key !== null) {
    const fresh = cellIn(projectedCellTables(engine, root), story);
    if (JSON.stringify(fresh) !== JSON.stringify(located)) {
      throw new Error(`cell location cache differs for ${story}: ${JSON.stringify(located)} != ${JSON.stringify(fresh)}`);
    }
  }
  return located;
}

function validPayload(value: unknown): value is TablePayload {
  return !!value && typeof value === 'object' && Array.isArray((value as TablePayload).rows);
}

function anchorsOf(engine: DocxWorkerCommandEngine, story: string, tableIndex: number) {
  if (typeof engine.tablePayload !== 'function') return null;
  const payload = engine.tablePayload(story, tableIndex);
  if (!validPayload(payload) || payload.rows.length === 0) return null;
  const { anchors, columns } = tableAnchors(payload);
  if (columns === 0) return null;
  return { payload, anchors, columns };
}

export function tableRangeFor(
  engine: DocxWorkerCommandEngine, focused: YrsCellLoc, target: TableSelectionTarget
): YrsTableRange | null {
  const resolved = anchorsOf(engine, focused.story, focused.tableIndex);
  if (!resolved) return null;
  const { payload, columns } = resolved;
  const covered = resolved.anchors.some((entry) => entry.row <= focused.row && focused.row < entry.row + entry.rowspan &&
    entry.column <= focused.column && focused.column < entry.column + entry.colspan);
  if (!covered) return null;
  const endRow = payload.rows.length - 1;
  const endColumn = columns - 1;
  if (target === 'row') {
    return { anchor: { ...focused, column: 0 }, head: { ...focused, column: endColumn } };
  }
  if (target === 'column') {
    return { anchor: { ...focused, row: 0 }, head: { ...focused, row: endRow } };
  }
  return { anchor: { ...focused, row: 0, column: 0 }, head: { ...focused, row: endRow, column: endColumn } };
}

/** The first caret of a cell: its first block, entering a block control that comes before any paragraph. */
export function cellStart(engine: Pick<DocxWorkerCommandEngine, 'geometryReader'>, story: string): YrsLoc | null {
  for (const segment of engine.geometryReader.storySegments(story)) {
    if (segment.kind === 'pilcrow') break;
    if (segment.kind === 'embed' && segment.embedKind === 'blockSdt' && typeof segment.payload.story === 'string') {
      const inner = cellStart(engine, segment.payload.story);
      if (inner) return inner;
    }
  }
  const first = engine.geometryReader.paragraphSpans(story)[0];
  return first ? { story, paraId: first.paraId, offset: 0 } : null;
}

export function cellSelectionHeadPosition(engine: DocxWorkerCommandEngine, focus: YrsLoc): YrsLoc | null {
  const active = readCellSelection(engine);
  if (!active || (active.anchor.row === active.head.row && active.anchor.column === active.head.column)) return null;
  const caret = resolveLiveCell(engine, focus.story);
  if (!caret || active.anchor.story !== caret.story || active.anchor.tableIndex !== caret.tableIndex ||
    active.head.story !== caret.story || active.head.tableIndex !== caret.tableIndex) return null;
  const resolved = anchorsOf(engine, caret.story, caret.tableIndex);
  const head = resolved?.anchors.find((entry) => entry.row <= active.head.row && active.head.row < entry.row + entry.rowspan &&
    entry.column <= active.head.column && active.head.column < entry.column + entry.colspan);
  return head ? cellStart(engine, head.story) : null;
}

export function cellsForCaret(engine: DocxWorkerCommandEngine, story: string): YrsTableRange | null {
  const cell = resolveLiveCell(engine, story);
  return cell ? { anchor: cell, head: { ...cell } } : null;
}

interface CommonTable {
  story: string;
  tableIndex: number;
  anchor: { row: number; column: number; story: string };
  head: { row: number; column: number; story: string };
}

function commonTableFor(engine: DocxWorkerCommandEngine, first: string, second: string): CommonTable | null {
  if (first === second || rootOf(first) !== rootOf(second)) return null;
  const projection = createYrsPositionProjection(engine.geometryReader, rootOf(first));
  if (!projection) return null;
  let best: { story: string; tableIndex: number } | null = null;
  for (const table of projection.allTables()) {
    const holds = (story: string) => table.cells.some((cell) => cell.story === story || story.startsWith(`${cell.story}:`));
    if (!holds(first) || !holds(second)) continue;
    if (!best || table.story.length > best.story.length) best = { story: table.story, tableIndex: table.tableIndex };
  }
  if (!best) return null;
  const resolved = anchorsOf(engine, best.story, best.tableIndex);
  if (!resolved) return null;
  const map = (story: string) => resolved.anchors.find((entry) => entry.story === story) ??
    resolved.anchors.find((entry) => story.startsWith(`${entry.story}:`));
  const anchor = map(first);
  const head = map(second);
  if (!anchor || !head || anchor.story === head.story) return null;
  return {
    ...best,
    anchor: { row: anchor.row, column: anchor.column, story: anchor.story },
    head: { row: head.row, column: head.column, story: head.story },
  };
}

export interface PointerCellRange {
  cells: YrsTableRange;
  textAnchor: YrsLoc;
}

export function resolvePointerCellRange(
  engine: DocxWorkerCommandEngine, anchorAt: YrsLoc, hitAt: YrsLoc
): PointerCellRange | null {
  const common = commonTableFor(engine, anchorAt.story, hitAt.story);
  if (!common) return null;
  const active = readCellSelection(engine);
  const sameTable = !!active && active.anchor.story === common.story && active.anchor.tableIndex === common.tableIndex &&
    active.head.story === common.story && active.head.tableIndex === common.tableIndex;
  const anchor: YrsCellLoc = sameTable && active
    ? active.anchor
    : { story: common.story, tableIndex: common.tableIndex, row: common.anchor.row, column: common.anchor.column };
  const resolved = anchorsOf(engine, common.story, common.tableIndex);
  const covering = resolved?.anchors.find((entry) => entry.row <= anchor.row && anchor.row < entry.row + entry.rowspan &&
    entry.column <= anchor.column && anchor.column < entry.column + entry.colspan);
  const coveringStory = covering?.story ?? common.anchor.story;
  const textAnchor = anchorAt.story === coveringStory ? anchorAt : (cellStart(engine, coveringStory) ?? anchorAt);
  return {
    cells: {
      anchor,
      head: { story: common.story, tableIndex: common.tableIndex, row: common.head.row, column: common.head.column },
    },
    textAnchor,
  };
}

export function shiftArrowCellDelta(
  unit: string, direction: string
): { dRow: number; dCol: number } | null {
  if ((unit === 'character' || unit === 'word') && (direction === 'backward' || direction === 'forward')) {
    return { dRow: 0, dCol: direction === 'forward' ? 1 : -1 };
  }
  if ((unit === 'line' || unit === 'paragraph') && (direction === 'up' || direction === 'down')) {
    return { dRow: direction === 'down' ? 1 : -1, dCol: 0 };
  }
  return null;
}

export interface ShiftArrowCellRange {
  selection: DocxCommandSelection;
  cells: YrsTableRange;
  goalX?: number;
}

export function resolveShiftArrowCellRange(
  engine: DocxWorkerCommandEngine,
  base: DocxCommandSelection,
  moved: YrsLoc,
  unit: string,
  direction: string,
  goalX?: number
): ShiftArrowCellRange | null {
  if (unit === 'page') {
    let pointer = resolvePointerCellRange(engine, base.anchor, moved);
    if (!pointer && base.anchor.story !== moved.story) {
      const cell = resolveLiveCell(engine, moved.story);
      const resolved = cell && base.anchor.story === cell.story ? anchorsOf(engine, cell.story, cell.tableIndex) : null;
      const edge = direction === 'down' ? resolved?.anchors[0] : resolved?.anchors.at(-1);
      const at = edge ? cellStart(engine, edge.story) : null;
      if (cell && edge && at) {
        pointer = {
          textAnchor: at,
          cells: {
            anchor: { story: cell.story, tableIndex: cell.tableIndex, row: edge.row, column: edge.column },
            head: cell,
          },
        };
      }
    }
    return pointer ? {
      selection: { anchor: pointer.textAnchor, focus: pointer.textAnchor }, cells: pointer.cells, goalX,
    } : null;
  }
  const delta = shiftArrowCellDelta(unit, direction);
  const caret = delta ? resolveLiveCell(engine, base.focus.story) : null;
  if (!delta || !caret) return null;
  const active = readCellSelection(engine);
  const sameTable = !!active && active.anchor.story === caret.story && active.anchor.tableIndex === caret.tableIndex &&
    active.head.story === caret.story && active.head.tableIndex === caret.tableIndex;
  const multi = !!sameTable && !!active &&
    (active.anchor.row !== active.head.row || active.anchor.column !== active.head.column);
  const leftCell = moved.story !== base.focus.story;
  const stuck = !leftCell && moved.paraId === base.focus.paraId && moved.offset === base.focus.offset;
  if (!multi && !leftCell && !stuck) return null;
  const resolved = anchorsOf(engine, caret.story, caret.tableIndex);
  if (!resolved) return null;
  const anchor: YrsCellLoc = multi && active
    ? active.anchor
    : { story: caret.story, tableIndex: caret.tableIndex, row: caret.row, column: caret.column };
  const start: YrsCellLoc = multi && active ? active.head : anchor;
  const head: YrsCellLoc = {
    ...start,
    row: Math.max(0, Math.min(resolved.payload.rows.length - 1, start.row + delta.dRow)),
    column: Math.max(0, Math.min(resolved.columns - 1, start.column + delta.dCol)),
  };
  const covering = resolved.anchors.find((entry) => entry.row <= anchor.row && anchor.row < entry.row + entry.rowspan &&
    entry.column <= anchor.column && anchor.column < entry.column + entry.colspan);
  const coveringStory = covering?.story;
  const at = coveringStory && base.anchor.story === coveringStory
    ? base.anchor
    : (coveringStory ? cellStart(engine, coveringStory) : null);
  if (!at) return null;
  const selection = { anchor: at, focus: at };
  return { selection, cells: { anchor, head }, ...(direction === 'up' || direction === 'down' ? { goalX } : {}) };
}
