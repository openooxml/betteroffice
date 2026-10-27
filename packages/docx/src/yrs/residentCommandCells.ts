import type { YrsCellLoc, YrsTableRange } from './index';
import type { DocxCommandPayloads, DocxCommandSelection, DocxWorkerCommandEngine } from './residentCommandProtocol';
import { tableAnchors, type TablePayload } from './storyPlainText';

export interface DocxCellNavigationPlan {
  selection: DocxCommandSelection;
  cells: YrsTableRange | null;
  insertRow?: YrsCellLoc;
}

function cellSelection(engine: DocxWorkerCommandEngine, story: string): DocxCommandSelection | null {
  const paragraph = engine.geometryReader.paragraphSpans(story)[0];
  if (!paragraph) return null;
  const at = { story, paraId: paragraph.paraId, offset: 0 };
  return { anchor: at, focus: at };
}

export function resolveCellNavigation(
  engine: DocxWorkerCommandEngine,
  selection: DocxCommandSelection,
  intent: DocxCommandPayloads['navigateCell'],
  readOnly: boolean
): DocxCellNavigationPlan | 'read-only' | null {
  const story = selection.focus.story;
  const matches = Array.from(story.matchAll(/:t(\d+):r(\d+)c(\d+)(?:~\d+)?(?=:|$)/g));
  const match = matches.at(-1);
  if (!match || match.index === undefined) return { selection, cells: null };
  const table = { story: story.slice(0, match.index), tableIndex: Number(match[1]) };
  if (!Number.isSafeInteger(table.tableIndex)) return null;
  const payload = engine.tablePayload(table.story, table.tableIndex) as TablePayload | null;
  if (!payload || !Array.isArray(payload.rows)) return null;
  const { anchors } = tableAnchors(payload);
  const currentIndex = anchors.findIndex((cell) => cell.story === story);
  if (currentIndex < 0) return null;
  const current = anchors[currentIndex]!;
  const at = { ...table, row: current.row, column: current.column };
  const selected = engine.cellSelection();
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
