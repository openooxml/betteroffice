import type { YrsCellLoc, YrsLoc, YrsTableLoc, YrsTableRange } from './index';
import { stickyPositionKey } from './residentCommandAnchors';
import { createYrsPositionProjection } from './yrsPositionProjection';
import { tableAnchors, type TablePayload } from './storyPlainText';
import type {
  DocxCommandIntent,
  DocxCommandSelection,
  ResidentDocxCommandEngine,
  ResidentDocxCommandState,
} from './residentCommandProtocol';

export type TableCommandKind =
  | 'insertTableRow' | 'insertTableColumn'
  | 'deleteTableRow' | 'deleteTableColumn' | 'deleteTable'
  | 'mergeTableCells' | 'splitTableCell' | 'setTableCellShading';

export type TableCommandIntent = Extract<DocxCommandIntent, { kind: TableCommandKind }>;

export interface DocxTableCommandTarget {
  focused: YrsCellLoc;
  range: YrsTableRange;
  table: YrsTableLoc;
  near: YrsLoc | null;
}

const TABLE_COMMANDS: ReadonlySet<string> = new Set<string>([
  'insertTableRow', 'insertTableColumn', 'deleteTableRow', 'deleteTableColumn',
  'deleteTable', 'mergeTableCells', 'splitTableCell', 'setTableCellShading',
]);

export function isTableCommandKind(kind: string): kind is TableCommandKind {
  return TABLE_COMMANDS.has(kind);
}

function rootOf(story: string): string {
  return story === 'body' || story.startsWith('body:') ? 'body'
    : story.split(':').slice(0, 2).join(':');
}

function sameGridCell(left: YrsCellLoc, right: YrsCellLoc): boolean {
  return left.row === right.row && left.column === right.column;
}

function sameCell(left: YrsCellLoc, right: YrsCellLoc): boolean {
  return left.story === right.story && left.tableIndex === right.tableIndex && sameGridCell(left, right);
}

function liveCell(engine: ResidentDocxCommandEngine, story: string): YrsCellLoc | null {
  const projection = createYrsPositionProjection(engine.geometryReader, rootOf(story));
  if (!projection) return null;
  let best: { story: string; tableIndex: number; row: number; column: number } | null = null;
  for (const table of projection.allTables()) {
    const cell = table.cells.find((candidate) => candidate.story === story);
    if (!cell) continue;
    if (!best || table.story.length > best.story.length) {
      best = { story: table.story, tableIndex: table.tableIndex, row: cell.row, column: cell.column };
    }
  }
  return best;
}

function readCellSelection(engine: ResidentDocxCommandEngine): YrsTableRange | null {
  try {
    return engine.cellSelection();
  } catch (error) {
    if (error instanceof WebAssembly.RuntimeError) throw error;
    return null;
  }
}

export function resolveTableTarget(
  engine: ResidentDocxCommandEngine,
  selection: DocxCommandSelection
): { focused: YrsCellLoc; range: YrsTableRange; table: YrsTableLoc } | null {
  const focused = liveCell(engine, selection.focus.story);
  if (!focused) return null;
  const table = { story: focused.story, tableIndex: focused.tableIndex };
  const anchor = liveCell(engine, selection.anchor.story);
  if (anchor && anchor.story === focused.story && anchor.tableIndex === focused.tableIndex && !sameCell(anchor, focused)) {
    return { focused, range: { anchor, head: focused }, table };
  }
  let range: YrsTableRange = { anchor: focused, head: focused };
  const selected = readCellSelection(engine);
  if (selected && selected.anchor.story === focused.story && selected.head.story === focused.story &&
    selected.anchor.tableIndex === focused.tableIndex && selected.head.tableIndex === focused.tableIndex &&
    (sameGridCell(selected.anchor, focused) || sameGridCell(selected.head, focused))) {
    range = selected;
  }
  return { focused, range, table };
}

export function nearTable(engine: ResidentDocxCommandEngine, table: YrsTableLoc): YrsLoc | null {
  const segments = engine.geometryReader.storySegments(table.story);
  let ordinal = 0;
  let target = -1;
  for (let index = 0; index < segments.length; index += 1) {
    const segment = segments[index]!;
    if (segment.kind !== 'embed' || segment.embedKind !== 'table') continue;
    if (ordinal === table.tableIndex) {
      target = index;
      break;
    }
    ordinal += 1;
  }
  if (target < 0) return null;
  const following = segments.slice(target + 1).find((segment) => segment.kind === 'pilcrow');
  const preceding = segments.slice(0, target).reverse().find((segment) => segment.kind === 'pilcrow');
  const paragraph = following ?? preceding;
  return paragraph?.kind === 'pilcrow'
    ? { story: table.story, paraId: paragraph.paraId, offset: 0 }
    : null;
}

export function ancestorTableLocs(story: string): YrsTableLoc[] {  const locs: { loc: YrsTableLoc; row: number; column: number }[] = [];
  for (const match of story.matchAll(/:t(\d+):r(\d+)c(\d+)(?=:|$)/g)) {
    if (match.index === undefined) continue;
    const tableIndex = Number(match[1]);
    const row = Number(match[2]);
    const column = Number(match[3]);
    if (![tableIndex, row, column].every(Number.isSafeInteger)) continue;
    locs.push({ loc: { story: story.slice(0, match.index), tableIndex }, row, column });
  }
  return locs.map(({ loc }) => loc);
}

export type ResolvedTableSelection =
  | { ok: true; selection: DocxCommandSelection }
  | { ok: false; reason: string };

function resolveTableSelectionEnd(
  engine: ResidentDocxCommandEngine,
  state: ResidentDocxCommandState,
  at: unknown
): YrsLoc | 'stale' | null {
  if (!at || typeof at !== 'object' || Array.isArray(at)) return null;
  const loc = at as Record<string, unknown>;
  if (typeof loc.story !== 'string' || !loc.story || typeof loc.paraId !== 'string' || !loc.paraId ||
    !Number.isSafeInteger(loc.offset) || (loc.offset as number) < 0 || (loc.offset as number) > 0xffffffff) return null;
  let resolved: YrsLoc = { story: loc.story, paraId: loc.paraId, offset: loc.offset as number };
  if (loc.sticky !== undefined) {
    if (!(loc.sticky instanceof Uint8Array)) return null;
    const issued = state.stickyPositions.get(stickyPositionKey(resolved.story, loc.sticky));
    if (!issued || issued.story !== resolved.story || !engine.geometryReader.hasStory(issued.story)) return 'stale';
    const position = engine.resolveStickyPosition(issued);
    if (!position || position.story !== resolved.story) return 'stale';
    resolved = position;
  }
  if (!engine.geometryReader.hasStory(resolved.story)) return null;
  if (engine.geometryReader.paragraphIdCount(resolved.story, resolved.paraId) !== 1) return null;
  const span = engine.geometryReader.locateParagraph(resolved.story, resolved.paraId);
  if (resolved.offset > span.end - span.start) return null;
  return resolved;
}

export function resolveTableSelectionEnds(
  engine: ResidentDocxCommandEngine,
  state: ResidentDocxCommandState,
  selection: DocxCommandSelection
): ResolvedTableSelection {
  if (!selection || typeof selection !== 'object' || Array.isArray(selection)) return { ok: false, reason: 'invalid-selection' };
  const anchor = resolveTableSelectionEnd(engine, state, selection.anchor);
  if (anchor === 'stale') return { ok: false, reason: 'stale-anchor' };
  const focus = resolveTableSelectionEnd(engine, state, selection.focus);
  if (focus === 'stale') return { ok: false, reason: 'stale-anchor' };
  if (!anchor || !focus) return { ok: false, reason: 'invalid-selection' };
  return { ok: true, selection: { anchor, focus } };
}

export function preflightTableCommand(
  engine: ResidentDocxCommandEngine,
  command: TableCommandIntent,
  selection: DocxCommandSelection
): { target: DocxTableCommandTarget } | { reason: string } {
  const resolved = resolveTableTarget(engine, selection);
  if (!resolved) return { reason: 'invalid-cell-target' };
  if (rootOf(resolved.table.story) !== 'body') return { reason: 'invalid-selection' };
  if (command.payload.suggesting !== undefined &&
    command.kind !== 'insertTableRow' && command.kind !== 'deleteTableRow') {
    return { reason: 'unsupported-suggesting' };
  }
  if (command.kind === 'mergeTableCells' && sameGridCell(resolved.range.anchor, resolved.range.head)) {
    return { reason: 'invalid-selection' };
  }
  if (command.kind === 'splitTableCell') {
    const found = liveAnchor(engine, resolved.focused);
    if (!found) return { reason: 'invalid-cell-target' };
    if (command.payload.rows < found.rowspan || command.payload.columns < found.colspan ||
      (command.payload.rows === 1 && command.payload.columns === 1)) {
      return { reason: 'invalid-table-size' };
    }
  }
  return { target: { ...resolved, near: nearTable(engine, resolved.table) } };
}

function liveAnchor(engine: ResidentDocxCommandEngine, at: YrsCellLoc) {
  const payload = engine.tablePayload(at.story, at.tableIndex) as TablePayload | null;
  if (!payload || !Array.isArray(payload.rows)) return null;
  return tableAnchors(payload).anchors.find((anchor) =>
    anchor.row <= at.row && at.row < anchor.row + anchor.rowspan &&
    anchor.column <= at.column && at.column < anchor.column + anchor.colspan) ?? null;
}

function liveCellStory(engine: ResidentDocxCommandEngine, at: YrsCellLoc): string | null {
  return liveAnchor(engine, at)?.story ?? null;
}

function nearestCellStory(engine: ResidentDocxCommandEngine, at: YrsCellLoc): string | null {
  const payload = engine.tablePayload(at.story, at.tableIndex) as TablePayload | null;
  if (!payload || !Array.isArray(payload.rows)) return null;
  const anchors = tableAnchors(payload).anchors;
  const covering = anchors.find((anchor) =>
    anchor.row <= at.row && at.row < anchor.row + anchor.rowspan &&
    anchor.column <= at.column && at.column < anchor.column + anchor.colspan);
  if (covering) return covering.story;
  const column = anchors
    .filter((anchor) => anchor.column <= at.column && at.column < anchor.column + anchor.colspan)
    .sort((left, right) => Math.abs(left.row - at.row) - Math.abs(right.row - at.row));
  const ordered = anchors.slice().sort((left, right) => (left.row - at.row) || (left.column - at.column));
  return (column[0] ?? ordered[0])?.story ?? null;
}

export function syncCellSelectionWithCaret(engine: ResidentDocxCommandEngine, story: string): void {
  const cell = liveCell(engine, story);
  if (cell) engine.setCellSelection({ anchor: cell, head: cell });
  else engine.clearCellSelection();
}

export interface TableCommandSink {
  caret(at: YrsLoc): void;
  repair(): boolean;
}

function caretInCellStory(
  engine: ResidentDocxCommandEngine,
  story: string | null,
  keep: YrsLoc | null
): YrsLoc | null {
  if (!story || !engine.geometryReader.hasStory(story)) return null;
  if (keep && keep.story === story) {
    const spans = engine.geometryReader.paragraphSpans(story);
    const length = spans.find((span) => span.paraId === keep.paraId)?.length;
    if (length !== undefined) return { story, paraId: keep.paraId, offset: Math.min(keep.offset, length) };
  }
  const paragraph = engine.geometryReader.paragraphSpans(story)[0];
  return paragraph ? { story, paraId: paragraph.paraId, offset: 0 } : null;
}

export function applyTableCommand(
  engine: ResidentDocxCommandEngine,
  state: ResidentDocxCommandState,
  command: TableCommandIntent,
  target: DocxTableCommandTarget,
  keep: YrsLoc | null,
  sink: TableCommandSink
): void {
  const version = engine.proposalEngine.version();
  engine.beginUndoCapture();
  engine.addUndoBoundary();
  const place = (at: YrsLoc | null) => {
    if (!at) {
      sink.repair();
      return;
    }
    sink.caret(at);
    syncCellSelectionWithCaret(engine, at.story);
  };
  switch (command.kind) {
    case 'insertTableRow':
      engine.setCellSelection({ anchor: target.focused, head: target.focused });
      engine.insertRow(target.focused, command.payload.side, command.payload.suggesting);
      place(caretInCellStory(engine, nearestCellStory(engine, target.focused), keep));
      break;
    case 'insertTableColumn':
      engine.setCellSelection({ anchor: target.focused, head: target.focused });
      engine.insertColumn(target.focused, command.payload.side);
      place(caretInCellStory(engine, nearestCellStory(engine, target.focused), keep));
      break;
    case 'deleteTableRow':
    case 'deleteTableColumn': {
      const receipt = command.kind === 'deleteTableRow'
        ? engine.deleteRow(target.range, command.payload.suggesting)
        : engine.deleteColumn(target.range);
      if (receipt.deletedTable) {
        place(target.near);
        break;
      }
      const surviving = readCellSelection(engine)?.anchor ?? target.focused;
      place(caretInCellStory(engine, nearestCellStory(engine, surviving), keep));
      break;
    }
    case 'deleteTable':
      engine.deleteTable(target.table);
      place(target.near);
      break;
    case 'mergeTableCells': {
      engine.mergeCells(target.range);
      const surviving = readCellSelection(engine)?.anchor ?? target.range.anchor;
      place(caretInCellStory(engine, nearestCellStory(engine, surviving), keep));
      break;
    }
    case 'splitTableCell': {
      const topLeft = liveCellStory(engine, target.focused) ?? target.focused.story;
      engine.splitCell(target.focused, command.payload.rows, command.payload.columns);
      place(caretInCellStory(engine, topLeft, keep));
      break;
    }
    case 'setTableCellShading':
      engine.setCellShading(target.range, command.payload.color);
      break;
  }
  engine.addUndoBoundary();
  if (engine.proposalEngine.version() !== version) state.storedMarks = null;
}
