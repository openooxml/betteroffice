import { normalizeRange, selectionKeyReducer } from '@betteroffice/xlsx';
import type {
  CellAddr, CellInputEdit, CellRange, Direction, GridMeta, KeyInput, Selection, SelectionData, SelectionHidden, SelectionLimits,
  WorkbookSession,
} from '@betteroffice/xlsx';

const SHEET_ROWS = 1_048_576;
const SHEET_COLS = 16_384;
export const MAX_CLEAR_CELLS = 100_000;

type UsedRange = () => CellRange | null | undefined;

export const reachesSheetEdge = (range: CellRange) => range.bottom >= SHEET_ROWS - 1 || range.right >= SHEET_COLS - 1;

export function copiedRange(range: CellRange, used: UsedRange): CellRange {
  if (!reachesSheetEdge(range)) return range;
  const bounds = used();
  const end = (start: number, last: number, usedLast: number | undefined, size: number) =>
    last >= size - 1 ? Math.max(start, Math.min(last, usedLast ?? start)) : last;
  return { ...range, bottom: end(range.top, range.bottom, bounds?.bottom, SHEET_ROWS),
    right: end(range.left, range.right, bounds?.right, SHEET_COLS) };
}

export async function copiedSelection(
  session: Pick<WorkbookSession, 'call'>, sheet: number, selection: Selection
): Promise<CellRange> {
  const range = normalizeRange(selection);
  if (!reachesSheetEdge(range)) return range;
  const { usedRange } = await session.call.selectionData(sheet, 0, 0, 'usedRange');
  return copiedRange(range, () => usedRange);
}

export function clearedCells(range: CellRange, used: UsedRange): CellInputEdit[] | null {
  const edge = reachesSheetEdge(range);
  const bounds = edge ? used() : range;
  const edits: CellInputEdit[] = [];
  if (!bounds) return edits;
  const top = Math.max(range.top, bounds.top);
  const bottom = Math.min(range.bottom, bounds.bottom);
  const left = Math.max(range.left, bounds.left);
  const right = Math.min(range.right, bounds.right);
  if (top > bottom || left > right) return edits;
  if (edge && (bottom - top + 1) * (right - left + 1) > MAX_CLEAR_CELLS) return null;
  for (let row = top; row <= bottom; row++) {
    for (let col = left; col <= right; col++) {
      edits.push({ row, col, input: '' });
    }
  }
  return edits;
}

export function selectionHidden(grid: GridMeta | undefined): SelectionHidden {
  const contains = (spans: [number, number][] | undefined, index: number) => {
    let low = 0;
    let high = spans?.length ?? 0;
    while (low < high) {
      const mid = (low + high) >>> 1;
      if (spans![mid]![1] < index) low = mid + 1;
      else high = mid;
    }
    return low < (spans?.length ?? 0) && spans![low]![0] <= index;
  };
  return {
    isRowHidden: (row) => contains(grid?.hiddenRowSpans, row),
    isColHidden: (col) => contains(grid?.hiddenColSpans, col),
  };
}

export function selectionDataRequest(key: KeyInput): Direction | 'region' | 'usedRange' | null {
  if (!key.ctrlKey && !key.metaKey) return null;
  switch (key.key) {
    case 'ArrowUp': return 'up';
    case 'ArrowDown': return 'down';
    case 'ArrowLeft': return 'left';
    case 'ArrowRight': return 'right';
    case 'End': return 'usedRange';
    case 'a':
    case 'A': return 'region';
    default: return null;
  }
}

export function sameSelection(first: Selection | null, second: Selection): boolean {
  if (!first || first.focus.row !== second.focus.row || first.focus.col !== second.focus.col ||
    first.anchor.row !== second.anchor.row || first.anchor.col !== second.anchor.col ||
    Boolean(first.range) !== Boolean(second.range)) return false;
  const a = normalizeRange(first);
  const b = normalizeRange(second);
  return a.top === b.top && a.left === b.left && a.bottom === b.bottom && a.right === b.right;
}

interface NavigationContext {
  owner: object;
  sheet: number;
  selection: Selection;
  navigation: number;
  version: string;
}

interface NavigationOptions {
  context(): NavigationContext | null;
  activeCell?(selection: Selection): CellAddr;
  limits(): SelectionLimits;
  hidden?(): SelectionHidden;
  query(context: NavigationContext, direction: Direction | 'usedRange' | undefined): SelectionData | Promise<SelectionData>;
  apply(selection: Selection): boolean | Promise<boolean>;
}

interface DeferredAction {
  run(): void;
  cancel?(): void;
}

interface NavigationBatch {
  context: NavigationContext;
  keys: (KeyInput | DeferredAction)[];
  target?: Selection;
}

const isAction = (key: KeyInput | DeferredAction): key is DeferredAction => 'run' in key;

function samePosition(first: NavigationContext, second: NavigationContext | null): boolean {
  return second !== null && first.owner === second.owner && first.sheet === second.sheet &&
    first.selection === second.selection && first.navigation === second.navigation;
}

function joins(batch: NavigationBatch, context: NavigationContext): boolean {
  return samePosition(batch.context, context) || Boolean(batch.target) &&
    batch.context.owner === context.owner && batch.context.sheet === context.sheet &&
    batch.context.navigation === context.navigation && sameSelection(batch.target!, context.selection);
}

export function createSelectionDataNavigation() {
  let pending: NavigationBatch | null = null;
  const drain = async (batch: NavigationBatch, options: NavigationOptions) => {
    let replay = false;
    try {
      while (pending === batch && batch.keys.length > 0) {
        const key = batch.keys[0]!;
        const selected = options.context();
        if (!samePosition(batch.context, selected) || !selected) return;
        if (isAction(key)) {
          batch.keys.shift();
          key.run();
          const after = options.context();
          if (!after) { replay = true; return; }
          batch.context = after;
          continue;
        }
        const request = selectionDataRequest(key);
        const selection = options.activeCell && (!key.shiftKey || request === 'region')
          ? { ...selected.selection, focus: options.activeCell(selected.selection), range: normalizeRange(selected.selection) }
          : selected.selection;
        let data: SelectionData | undefined;
        try {
          data = await options.query(selection === selected.selection ? selected : { ...selected, selection },
            request === 'region' || !request ? undefined : request);
        } catch {}
        const current = options.context();
        if (pending !== batch || !samePosition(selected, current) || !current) return;
        if (selected.version !== current.version) continue;
        batch.keys.shift();
        const action = selectionKeyReducer(selection, key, options.limits(), data, options.hidden?.());
        if (action.type !== 'move') continue;
        batch.target = action.selection;
        const applied = options.apply(action.selection);
        const applying = options.context();
        if (!applying) return;
        batch.context = applying;
        if (!await applied || pending !== batch) return;
        const accepted = options.context();
        if (!accepted || accepted.owner !== applying.owner || accepted.sheet !== applying.sheet ||
          accepted.navigation !== applying.navigation) return;
        batch.context = accepted;
        batch.target = undefined;
      }
    } finally {
      if (pending === batch) pending = null;
      const rest = batch.keys.splice(0);
      while (rest.length > 0) {
        const key = rest.shift()!;
        if (isAction(key)) {
          if (replay) key.run(); else key.cancel?.();
          continue;
        }
        const context = replay && !pending ? options.context() : null;
        if (!context) continue;
        pending = { context, keys: [key, ...rest] };
        void drain(pending, options);
        break;
      }
    }
  };
  return {
    enqueue(key: KeyInput, options: NavigationOptions): void {
      const context = options.context();
      if (!context) return;
      const input = { key: key.key, ctrlKey: key.ctrlKey, metaKey: key.metaKey, shiftKey: key.shiftKey };
      if (pending && joins(pending, context)) {
        pending.keys.push(input);
        return;
      }
      const batch = { context, keys: [input] };
      pending = batch;
      void drain(batch, options);
    },
    defer(context: NavigationContext | null, run: () => void, behindActions = false, cancel?: () => void): boolean {
      if (!pending || !context || !joins(pending, context)) return false;
      if (behindActions && !pending.keys.some(isAction)) return false;
      pending.keys.push({ run, cancel });
      return true;
    },
  };
}
