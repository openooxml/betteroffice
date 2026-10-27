import { normalizeRange, selectionKeyReducer } from '@betteroffice/xlsx';
import type {
  CellAddr, CellInputEdit, CellRange, Direction, GridMeta, KeyInput, MergedRange, Selection, SelectionData, SelectionHidden,
  SelectionLimits, SelectionMerges, WorkbookHandle, WorkbookSession,
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

export function selectionFrameContains(grid: GridMeta | undefined, cell: CellAddr): boolean {
  if (!grid) return false;
  const contains = (index: number, indices: number[] | undefined, start: number, offsets: number[]) =>
    indices ? indices.includes(index) : index >= start && index < start + offsets.length - 1;
  return contains(cell.row, grid.rowIndices, grid.startRow, grid.rowOffsets) &&
    contains(cell.col, grid.colIndices, grid.startCol, grid.colOffsets);
}

export function selectionMerges(ranges: readonly MergedRange[] | undefined): SelectionMerges {
  return {
    mergeAt: (row, col) => {
      const merged = ranges?.find((range) => row >= Math.min(range.start.row, range.end.row) &&
        row <= Math.max(range.start.row, range.end.row) && col >= Math.min(range.start.col, range.end.col) &&
        col <= Math.max(range.start.col, range.end.col));
      return merged ? { top: Math.min(merged.start.row, merged.end.row), left: Math.min(merged.start.col, merged.end.col),
        bottom: Math.max(merged.start.row, merged.end.row), right: Math.max(merged.start.col, merged.end.col) } : null;
    },
  };
}

export function mergeAnchorCell(cell: CellAddr, merges: SelectionMerges): CellAddr {
  const merged = merges.mergeAt(cell.row, cell.col);
  return merged && (merged.top !== cell.row || merged.left !== cell.col) ? { row: merged.top, col: merged.left } : cell;
}

type MergeWorkbook = Pick<WorkbookHandle, 'version' | 'mergedRanges'>;
const workbookMergeCache = new WeakMap<MergeWorkbook, { version: string; sheets: Map<number, SelectionMerges> }>();

export function workbookSelectionMerges(
  workbook: MergeWorkbook | null | undefined,
  sheet: number, fallback?: SelectionMerges
): SelectionMerges {
  if (workbook) {
    try {
      const version = workbook.version();
      let cached = workbookMergeCache.get(workbook);
      if (!cached || cached.version !== version) {
        cached = { version, sheets: new Map() };
        workbookMergeCache.set(workbook, cached);
      }
      let geometry = cached.sheets.get(sheet);
      if (!geometry) {
        geometry = selectionMerges(workbook.mergedRanges(sheet, 'A1:XFD1048576'));
        cached.sheets.set(sheet, geometry);
      }
      return geometry;
    } catch {}
  }
  return fallback ?? selectionMerges(undefined);
}

export function anchorSelection(selection: Selection, merges: SelectionMerges): Selection {
  const collapsed = selection.anchor.row === selection.focus.row && selection.anchor.col === selection.focus.col;
  const anchor = mergeAnchorCell(selection.anchor, merges);
  const focus = collapsed ? mergeAnchorCell(selection.focus, merges) : selection.focus;
  if (anchor === selection.anchor && focus === selection.focus) return selection;
  const range = normalizeRange(selection);
  return { anchor, focus, ...(!collapsed || selection.range && (range.top !== range.bottom || range.left !== range.right)
    ? { range } : {}) };
}

export function selectionDataRequest(key: KeyInput): Direction | 'region' | 'usedRange' | 'rowEnd' | null {
  if (!key.ctrlKey && !key.metaKey) return key.key === 'End' ? 'rowEnd' : null;
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

export interface NavigationContext {
  owner: object;
  sheet: number;
  selection: Selection;
  navigation: number;
  version: string;
}

export interface NavigationOptions {
  context(): NavigationContext | null;
  activeCell?(selection: Selection): CellAddr;
  limits(): SelectionLimits;
  hidden?(): SelectionHidden;
  merges?(context?: NavigationContext): SelectionMerges;
  known?(context: NavigationContext, cell: CellAddr): boolean;
  query(context: NavigationContext, direction: Direction | 'usedRange' | 'rowEnd' | 'merge' | undefined): SelectionData | Promise<SelectionData>;
  apply(selection: Selection, sheet?: number): boolean | Promise<boolean>;
}

interface DeferredAction {
  run(): void;
  cancel?(): void;
}

interface SelectionTarget {
  selection: Selection;
  sheet?: number;
  options: NavigationOptions;
  done(value: boolean): void;
}

interface NavigationBatch {
  context: NavigationContext;
  keys: (KeyInput | DeferredAction | SelectionTarget)[];
  target?: Selection;
}

const isAction = (key: KeyInput | DeferredAction | SelectionTarget): key is DeferredAction => 'run' in key;

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
  let inAction = false;
  let actionInsert = 0;
  let learned: { owner: object; sheet: number; version: string; ranges: CellRange[]; cells: Set<string> } | null = null;
  const ranges = (context: NavigationContext) => {
    if (!learned || learned.owner !== context.owner || learned.sheet !== context.sheet || learned.version !== context.version) {
      learned = { owner: context.owner, sheet: context.sheet, version: context.version, ranges: [], cells: new Set() };
    }
    return learned.ranges;
  };
  const merges = (context: NavigationContext, fallback?: SelectionMerges): SelectionMerges => ({
    mergeAt: (row, col) => ranges(context).find((range) => row >= range.top && row <= range.bottom &&
      col >= range.left && col <= range.right) ?? fallback?.mergeAt(row, col) ?? null,
  });
  const known = (context: NavigationContext, cell: CellAddr) => {
    ranges(context);
    return Boolean(learned?.cells.has(`${cell.row}:${cell.col}`) || merges(context).mergeAt(cell.row, cell.col));
  };
  const learn = (context: NavigationContext, data: SelectionData | undefined, cell: CellAddr) => {
    ranges(context);
    if (data) {
      const at = data.edge ?? cell;
      learned!.cells.add(`${at.row}:${at.col}`);
    }
    if (data?.merge && !ranges(context).some((range) => range.top === data.merge!.top && range.left === data.merge!.left)) {
      ranges(context).push(data.merge);
    }
  };
  const originSelection = (selection: Selection) => selection.anchor.row === 0 && selection.anchor.col === 0 &&
    selection.focus.row === 0 && selection.focus.col === 0;
  const knownSelection = (context: NavigationContext, selection: Selection, options: NavigationOptions) =>
    options.known?.(context, selection.anchor) && options.known(context, selection.focus);
  const resolve = async (context: NavigationContext, selection: Selection, options: NavigationOptions) => {
    let next = selection;
    const collapsed = selection.anchor.row === selection.focus.row && selection.anchor.col === selection.focus.col;
    for (const endpoint of ['anchor', 'focus'] as const) {
      const cell = next[endpoint];
      let data: SelectionData | undefined;
      if (context.sheet !== options.context()?.sheet || !options.known?.(context, cell)) {
        try { data = await options.query({ ...context, selection: { anchor: cell, focus: cell } }, 'merge'); } catch {}
      }
      learn(context, data, cell);
      const target = data && 'merge' in data ? data.merge ? { row: data.merge.top, col: data.merge.left } : cell
        : mergeAnchorCell(cell, merges(context, context.sheet === options.context()?.sheet ? options.merges?.() : undefined));
      if ((collapsed || endpoint === 'anchor') && (target.row !== cell.row || target.col !== cell.col)) {
        next = { ...next, [endpoint]: target };
      }
      if (collapsed) {
        const { range, ...collapsed } = next;
        return { ...collapsed, anchor: target, focus: target,
          ...(range && (range.top !== range.bottom || range.left !== range.right) ? { range } : {}) };
      }
    }
    return next !== selection && !selection.range ? { ...next, range: normalizeRange(selection) } : next;
  };
  const drain = async (batch: NavigationBatch, options: NavigationOptions) => {
    let replay = false;
    try {
      while (pending === batch && batch.keys.length > 0) {
        const key = batch.keys[0]!;
        const itemOptions = 'selection' in key ? key.options : options;
        const selected = itemOptions.context();
        if (!samePosition(batch.context, selected) || !selected) return;
        if (isAction(key)) {
          batch.keys.shift();
          inAction = true;
          actionInsert = 0;
          try { key.run(); } finally { inAction = false; }
          await Promise.resolve();
          const after = itemOptions.context();
          if (!after) { replay = true; return; }
          batch.context = after;
          continue;
        }
        const target = 'selection' in key ? key : null;
        const input = target ? null : key as KeyInput;
        const request = input ? selectionDataRequest(input) : null;
        const selection = input && itemOptions.activeCell && (!input.shiftKey || input.key === 'Tab' || input.key === 'Enter' || request === 'region')
          ? { ...selected.selection, focus: itemOptions.activeCell(selected.selection), range: normalizeRange(selected.selection) }
          : selected.selection;
        const destination = target?.sheet === undefined ? selected : { ...selected, sheet: target.sheet };
        const local = !target && !request && knownSelection(selected, selection, itemOptions)
          ? selectionKeyReducer(selection, input!, itemOptions.limits(), undefined, itemOptions.hidden?.(), merges(selected, itemOptions.merges?.()))
          : undefined;
        const localMove = local?.type === 'move' ? local : undefined;
        let data: SelectionData | undefined;
        if (!target && !localMove) {
          try {
            data = await itemOptions.query(selection === selected.selection ? selected : { ...selected, selection },
              request === 'region' ? undefined : request ?? 'merge');
          } catch {}
        }
        const current = itemOptions.context();
        if (pending !== batch || !samePosition(selected, current) || !current) return;
        if (selected.version !== current.version) continue;
        learn(selected, data, selection.focus);
        const action = target ? { type: 'move' as const, selection: target.selection }
          : localMove ? { ...localMove, selection: anchorSelection(localMove.selection, merges(selected, itemOptions.merges?.())) }
          : selectionKeyReducer(selection, input!, itemOptions.limits(), data, itemOptions.hidden?.(), merges(selected, itemOptions.merges?.()));
        if (action.type !== 'move') { batch.keys.shift(); continue; }
        const next = target || localMove && !knownSelection(selected, action.selection, itemOptions) || data && 'merge' in data
          ? await resolve(destination, action.selection, itemOptions) : action.selection;
        const resolved = itemOptions.context();
        if (pending !== batch || !samePosition(selected, resolved) || !resolved) return;
        if (selected.version !== resolved.version) continue;
        batch.keys.shift();
        action.selection = next;
        batch.target = action.selection;
        const applied = itemOptions.apply(action.selection, target?.sheet);
        const applying = itemOptions.context();
        if (!applying) { target?.done(false); return; }
        batch.context = applying;
        const completed = await applied;
        target?.done(completed);
        if (!completed || pending !== batch) return;
        const accepted = itemOptions.context();
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
          if (replay) {
            const context = options.context();
            if (pending && context && joins(pending, context)) pending.keys.push(key);
            else key.run();
          } else key.cancel?.();
          continue;
        }
        const context = replay && !pending ? options.context() : null;
        if (!context) { if ('selection' in key) key.done(false); continue; }
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
      const input = { key: key.key, ctrlKey: key.ctrlKey, metaKey: key.metaKey, shiftKey: key.shiftKey, altKey: key.altKey };
      if (!pending && !selectionDataRequest(input) && knownSelection(context, context.selection, options)) {
        const selection = options.activeCell && (!input.shiftKey || input.key === 'Tab' || input.key === 'Enter')
          ? { ...context.selection, focus: options.activeCell(context.selection) } : context.selection;
        const geometry = merges(context, options.merges?.());
        const action = selectionKeyReducer(selection, input, options.limits(), undefined, options.hidden?.(), geometry);
        if (action.type === 'move' && knownSelection(context, action.selection, options)) {
          void options.apply(anchorSelection(action.selection, geometry));
          return;
        }
      }
      if (pending && joins(pending, context)) {
        if (inAction) pending.keys.splice(actionInsert++, 0, input); else pending.keys.push(input);
        return;
      }
      const batch = { context, keys: [input] };
      pending = batch;
      void drain(batch, options);
    },
    merges,
    known,
    isPending: (context: NavigationContext | null) => Boolean(pending && context && joins(pending, context)),
    beforeInput(options: NavigationOptions, run: () => void, cancel?: () => void): boolean {
      if (inAction) return false;
      const context = options.context();
      if (!context) return false;
      if (pending && joins(pending, context)) { pending.keys.push({ run, cancel }); return true; }
      if (originSelection(context.selection) || knownSelection(context, context.selection, options)) return false;
      void this.select(context.selection, options);
      if (pending && joins(pending, context)) { pending.keys.push({ run, cancel }); return true; }
      return false;
    },
    select(selection: Selection, options: NavigationOptions, sheet?: number, replace = false): Promise<boolean> {
      const context = options.context();
      if (!context) return Promise.resolve(false);
      if (replace && pending && !pending.keys.some(isAction)) pending = null;
      const destination = sheet === undefined ? context : { ...context, sheet };
      if (!pending && (originSelection(selection) || knownSelection(destination, selection, options))) {
        return Promise.resolve(options.apply(anchorSelection(selection, merges(destination, options.merges?.(destination))), sheet));
      }
      return new Promise((done) => {
        const target = { selection, sheet, options, done };
        if (pending && joins(pending, context)) {
          if (inAction) pending.keys.splice(actionInsert++, 0, target); else pending.keys.push(target);
          return;
        }
        pending = { context, keys: [target] };
        void drain(pending, options);
      });
    },
    defer(context: NavigationContext | null, run: () => void, behindActions = false, cancel?: () => void): boolean {
      if (inAction || !pending || !context || !joins(pending, context)) return false;
      if (behindActions && !pending.keys.some(isAction)) return false;
      pending.keys.push({ run, cancel });
      return true;
    },
  };
}
