import { normalizeRange, selectionKeyReducer } from '@betteroffice/xlsx';
import type { Direction, GridMeta, KeyInput, Selection, SelectionData, SelectionHidden, SelectionLimits } from '@betteroffice/xlsx';

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

export function selectionDataRequest(key: KeyInput): Direction | 'region' | null {
  if (!key.ctrlKey && !key.metaKey) return null;
  switch (key.key) {
    case 'ArrowUp': return 'up';
    case 'ArrowDown': return 'down';
    case 'ArrowLeft': return 'left';
    case 'ArrowRight': return 'right';
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
  limits(): SelectionLimits;
  hidden?(): SelectionHidden;
  query(context: NavigationContext, direction: Direction | undefined): SelectionData | Promise<SelectionData>;
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
        let data: SelectionData | undefined;
        try {
          data = await options.query(selected, request === 'region' || !request ? undefined : request);
        } catch {}
        const current = options.context();
        if (pending !== batch || !samePosition(selected, current) || !current) return;
        if (selected.version !== current.version) continue;
        batch.keys.shift();
        const action = selectionKeyReducer(selected.selection, key, options.limits(), data, options.hidden?.());
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
