import type { YrsCellLoc, YrsTableRange } from './index';
import type { DocxCommandPayloads, DocxCommandSelection, DocxWorkerCommandEngine } from './residentCommandProtocol';
import { tableAnchors, type TablePayload } from './storyPlainText';
import { cellStart, readCellSelection, resolveLiveCell } from './residentCommandTableSelection';
import { isUnavailableTablePayload } from './yrsPositionProjection';

const CELL_STORY = /:t\d+:r\d+c\d+(?:~\d+)?(?=:|$)/;

export interface DocxCellNavigationPlan {
  selection: DocxCommandSelection;
  cells: YrsTableRange | null;
  insertRow?: YrsCellLoc;
}

function cellSelection(engine: DocxWorkerCommandEngine, story: string): DocxCommandSelection | null {
  const at = cellStart(engine, story);
  return at && { anchor: at, focus: at };
}

export function resolveCellNavigation(
  engine: DocxWorkerCommandEngine,
  selection: DocxCommandSelection,
  intent: DocxCommandPayloads['navigateCell'],
  readOnly: boolean
): DocxCellNavigationPlan | 'read-only' | null {
  const story = selection.focus.story;
  const live = resolveLiveCell(engine, story);
  if (!live) return CELL_STORY.test(story) ? null : { selection, cells: null };
  const table = { story: live.story, tableIndex: live.tableIndex };
  const payload = engine.tablePayload(table.story, table.tableIndex) as TablePayload | null;
  if (!payload || isUnavailableTablePayload(payload) || !Array.isArray(payload.rows)) return null;
  const { anchors } = tableAnchors(payload);
  const currentIndex = anchors.findIndex((cell) => cell.story === story || story.startsWith(`${cell.story}:`));
  if (currentIndex < 0) return null;
  const current = anchors[currentIndex]!;
  const at = { ...table, row: current.row, column: current.column };
  const selected = readCellSelection(engine);
  const sameTable = selected?.anchor.story === table.story && selected.head.story === table.story &&
    selected.anchor.tableIndex === table.tableIndex && selected.head.tableIndex === table.tableIndex &&
    selected.head.row === current.row && selected.head.column === current.column;
  const anchor = intent.extend && sameTable ? selected!.anchor : at;
  const index = currentIndex + (intent.direction === 'forward' ? 1 : -1);
  const target = anchors[index];
  if (!target) {
    if (index >= anchors.length && intent.createLastRow) {
      if (readOnly) return 'read-only';
      return { selection, cells: { anchor, head: at }, insertRow: { ...at, row: payload.rows.length - 1 } };
    }
    return { selection, cells: { anchor, head: at } };
  }
  const next = cellSelection(engine, target.story);
  const head = { ...table, row: target.row, column: target.column };
  return next ? { selection: next, cells: { anchor: intent.extend ? anchor : head, head } } : null;
}

export function applyCellNavigation(
  engine: DocxWorkerCommandEngine,
  plan: DocxCellNavigationPlan,
  intent: DocxCommandPayloads['navigateCell']
): DocxCommandSelection {
  if (!plan.cells) return plan.selection;
  if (plan.insertRow) {
    const row = plan.insertRow.row + 1;
    engine.insertRow(plan.insertRow, 'below', intent.suggesting);
    const payload = engine.tablePayload(plan.insertRow.story, plan.insertRow.tableIndex) as unknown as TablePayload;
    const first = tableAnchors(payload).anchors.find((cell) => cell.row === row);
    const selection = first ? cellSelection(engine, first.story) : null;
    if (!first || !selection) throw new Error('Inserted row lacks a caret target');
    plan.selection = selection;
    plan.cells.head = { ...plan.insertRow, row, column: first.column };
    if (!intent.extend) plan.cells.anchor = plan.cells.head;
  }
  engine.setCellSelection(plan.cells);
  return plan.selection;
}
