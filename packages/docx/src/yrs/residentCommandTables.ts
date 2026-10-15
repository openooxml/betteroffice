import type { YrsCellLoc, YrsLoc, YrsTableLoc, YrsTableRange } from './index';
import { tableAnchors, type TableCellAnchor, type TablePayload } from './storyPlainText';
import type {
  DocxCommandIntent,
  DocxCommandSelection,
  ResidentDocxCommandEngine,
  ResidentDocxCommandState,
} from './residentCommandProtocol';
import { readCellSelection, resolveLiveCell, rootOf } from './residentCommandTableSelection';

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

const TABLE_COMMANDS = {
  insertTableRow: true, insertTableColumn: true, deleteTableRow: true, deleteTableColumn: true,
  deleteTable: true, mergeTableCells: true, splitTableCell: true, setTableCellShading: true,
} satisfies Record<TableCommandKind, true>;

export function isTableCommandKind(kind: string): kind is TableCommandKind {
  return Object.hasOwn(TABLE_COMMANDS, kind);
}

function resolveTableTarget(
  engine: ResidentDocxCommandEngine,
  selection: DocxCommandSelection
): { focused: YrsCellLoc; range: YrsTableRange; table: YrsTableLoc } | null {
  const focused = resolveLiveCell(engine, selection.focus.story);
  if (!focused) return null;
  const table = { story: focused.story, tableIndex: focused.tableIndex };
  const selected = readCellSelection(engine);
  const range = selected && [selected.anchor, selected.head].every((end) =>
    end.story === focused.story && end.tableIndex === focused.tableIndex) &&
    Math.min(selected.anchor.row, selected.head.row) <= focused.row &&
    focused.row <= Math.max(selected.anchor.row, selected.head.row) &&
    Math.min(selected.anchor.column, selected.head.column) <= focused.column &&
    focused.column <= Math.max(selected.anchor.column, selected.head.column)
    ? selected : { anchor: focused, head: focused };
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

/** Live tables holding a story, innermost first; empty outside tables. */
export function liveTableAncestors(engine: ResidentDocxCommandEngine, story: string): YrsTableLoc[] {
  const tables: YrsTableLoc[] = [];
  for (let cell = resolveLiveCell(engine, story); cell; cell = resolveLiveCell(engine, cell.story)) {
    tables.push({ story: cell.story, tableIndex: cell.tableIndex });
  }
  return tables;
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
  if (command.kind === 'mergeTableCells' &&
    expandedRangeAnchorCount(engine, resolved.focused, resolved.range) < 2) {
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
  const removes = command.kind === 'deleteTable' || command.kind === 'deleteTableRow' || command.kind === 'deleteTableColumn';
  return { target: { ...resolved, near: removes ? nearTable(engine, resolved.table) : null } };
}

function liveAnchors(engine: ResidentDocxCommandEngine, at: YrsCellLoc) {
  const payload = engine.tablePayload(at.story, at.tableIndex) as TablePayload | null;
  return payload && Array.isArray(payload.rows) ? tableAnchors(payload).anchors : null;
}

function coveringAnchor(anchors: readonly TableCellAnchor[], at: YrsCellLoc): TableCellAnchor | null {
  return anchors.find((anchor) =>
    anchor.row <= at.row && at.row < anchor.row + anchor.rowspan &&
    anchor.column <= at.column && at.column < anchor.column + anchor.colspan) ?? null;
}

function expandedRangeAnchorCount(
  engine: ResidentDocxCommandEngine,
  at: YrsCellLoc,
  range: YrsTableRange
): number {
  const anchors = liveAnchors(engine, at);
  if (!anchors) return 0;
  let top = Math.min(range.anchor.row, range.head.row);
  let bottom = Math.max(range.anchor.row, range.head.row) + 1;
  let left = Math.min(range.anchor.column, range.head.column);
  let right = Math.max(range.anchor.column, range.head.column) + 1;
  for (;;) {
    const beforeTop = top;
    const beforeBottom = bottom;
    const beforeLeft = left;
    const beforeRight = right;
    for (const anchor of anchors) {
      const intersects = anchor.row < bottom && anchor.row + anchor.rowspan > top &&
        anchor.column < right && anchor.column + anchor.colspan > left;
      if (!intersects) continue;
      top = Math.min(top, anchor.row);
      bottom = Math.max(bottom, anchor.row + anchor.rowspan);
      left = Math.min(left, anchor.column);
      right = Math.max(right, anchor.column + anchor.colspan);
    }
    if (beforeTop === top && beforeBottom === bottom && beforeLeft === left && beforeRight === right) break;
  }
  return anchors.filter((anchor) =>
    anchor.row >= top && anchor.row + anchor.rowspan <= bottom &&
    anchor.column >= left && anchor.column + anchor.colspan <= right).length;
}

function liveAnchor(engine: ResidentDocxCommandEngine, at: YrsCellLoc) {
  const anchors = liveAnchors(engine, at);
  return anchors ? coveringAnchor(anchors, at) : null;
}

function nearestCellStory(engine: ResidentDocxCommandEngine, at: YrsCellLoc): string | null {
  const anchors = liveAnchors(engine, at);
  if (!anchors) return null;
  const covering = coveringAnchor(anchors, at);
  if (covering) return covering.story;
  const column = anchors
    .filter((anchor) => anchor.column <= at.column && at.column < anchor.column + anchor.colspan)
    .sort((left, right) => Math.abs(left.row - at.row) - Math.abs(right.row - at.row));
  const ordered = anchors.slice().sort((left, right) =>
    (Math.abs(left.row - at.row) - Math.abs(right.row - at.row)) ||
    (Math.abs(left.column - at.column) - Math.abs(right.column - at.column)));
  return (column[0] ?? ordered[0])?.story ?? null;
}

export function syncCellSelectionWithCaret(engine: ResidentDocxCommandEngine, story: string): void {
  const cell = resolveLiveCell(engine, story);
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
  keep: YrsLoc
): YrsLoc | null {
  if (!story || !engine.geometryReader.hasStory(story)) return null;
  if (keep.story === story) {
    const spans = engine.geometryReader.paragraphSpans(story);
    const length = spans.find((span) => span.paraId === keep.paraId)?.length;
    if (length !== undefined) return { story, paraId: keep.paraId, offset: Math.min(keep.offset, length) };
  }
  const paragraph = engine.geometryReader.paragraphSpans(story)[0];
  return paragraph ? { story, paraId: paragraph.paraId, offset: 0 } : null;
}

/** Runs inside the caller's discrete undo group. */
export function applyTableCommand(
  engine: ResidentDocxCommandEngine,
  state: ResidentDocxCommandState,
  command: TableCommandIntent,
  target: DocxTableCommandTarget,
  keep: YrsLoc,
  sink: TableCommandSink
): void {
  const version = engine.proposalEngine.version();
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
      place(caretInCellStory(engine, keep.story, keep));
      break;
    case 'insertTableColumn':
      engine.setCellSelection({ anchor: target.focused, head: target.focused });
      engine.insertColumn(target.focused, command.payload.side);
      place(caretInCellStory(engine, keep.story, keep));
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
      const topLeft = liveAnchor(engine, target.focused)?.story ?? target.focused.story;
      engine.splitCell(target.focused, command.payload.rows, command.payload.columns);
      place(caretInCellStory(engine, topLeft, keep));
      break;
    }
    case 'setTableCellShading':
      engine.setCellShading(target.range, command.payload.color);
      break;
  }
  if (engine.proposalEngine.version() !== version) state.storedMarks = null;
}
