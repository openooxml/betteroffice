import { describe, expect, it, mock } from 'bun:test';
import { normalizeRange, selectionAt } from '@betteroffice/xlsx';
import type { Selection, SelectionData } from '@betteroffice/xlsx';
import { deriveLimits } from './viewer/sessionGeometry';
import { createSelectionDataNavigation, sameSelection, selectionHidden } from './selectionData';

function deferred<T>() {
  let resolve!: (value: T) => void;
  let reject!: (error: unknown) => void;
  const promise = new Promise<T>((done, fail) => { resolve = done; reject = fail; });
  return { promise, resolve, reject };
}

const tick = () => new Promise<void>((resolve) => setTimeout(resolve, 0));
const right = { key: 'ArrowRight', ctrlKey: true };

function navigation() {
  let current = { owner: {}, sheet: 0, navigation: 0, version: 'v1', selection: selectionAt({ row: 1, col: 1 }) };
  const query = mock(async (_context: typeof current, _direction: string | undefined): Promise<SelectionData> => ({
    edge: null, region: null,
  }));
  const apply = mock((selection: Selection) => {
    current = { ...current, navigation: current.navigation + 1, selection };
    return true;
  });
  const queue = createSelectionDataNavigation();
  return {
    get current() { return current; },
    update(next: Partial<typeof current>) { current = { ...current, ...next }; },
    query, apply,
    press: (key = right) => queue.enqueue(key, {
      context: () => current, limits: () => ({ rows: 100, cols: 20, rowsPerPage: 10 }), query, apply,
    }),
  };
}

describe('selection data navigation', () => {
  it('chains rapid repeated jumps from each accepted target', async () => {
    const nav = navigation();
    const first = deferred<SelectionData>();
    nav.query.mockReturnValueOnce(first.promise).mockResolvedValueOnce({ edge: { row: 1, col: 9 } });
    nav.press();
    nav.press();
    expect(nav.query).toHaveBeenCalledTimes(1);
    first.resolve({ edge: { row: 1, col: 4 } });
    await tick();
    expect(nav.query.mock.calls[1]![0].selection.focus).toEqual({ row: 1, col: 4 });
    expect(nav.current.selection.focus).toEqual({ row: 1, col: 9 });
    expect(nav.apply).toHaveBeenCalledTimes(2);
  });

  it('chains rapid select-all presses from the region to the full extent', async () => {
    const nav = navigation();
    const region = { top: 0, left: 0, bottom: 2, right: 2 };
    nav.query.mockResolvedValue({ region });
    nav.press({ key: 'a', ctrlKey: true });
    nav.press({ key: 'a', ctrlKey: true });
    await tick();
    expect(normalizeRange(nav.apply.mock.calls[0]![0])).toEqual(region);
    expect(normalizeRange(nav.current.selection)).toEqual({ top: 0, left: 0, bottom: 99, right: 19 });
    expect(nav.current.selection.focus).toEqual({ row: 1, col: 1 });
  });

  it('re-queries after an edit without losing the shortcut', async () => {
    const nav = navigation();
    const first = deferred<SelectionData>();
    const fresh = deferred<SelectionData>();
    nav.query.mockReturnValueOnce(first.promise).mockReturnValueOnce(fresh.promise);
    nav.press();
    nav.update({ version: 'v2' });
    first.resolve({ edge: { row: 1, col: 4 } });
    await tick();
    expect(nav.apply).not.toHaveBeenCalled();
    expect(nav.query).toHaveBeenCalledTimes(2);
    expect(nav.query.mock.calls[1]![0].version).toBe('v2');
    fresh.resolve({ edge: { row: 1, col: 7 } });
    await tick();
    expect(nav.current.selection.focus).toEqual({ row: 1, col: 7 });
  });

  it('falls back to the directional limit after a query failure', async () => {
    const nav = navigation();
    nav.query.mockRejectedValueOnce(new Error('Query failed'));
    nav.press();
    await tick();
    expect(nav.current.selection.focus).toEqual({ row: 1, col: 19 });
  });

  for (const change of ['owner', 'sheet', 'selection', 'navigation'] as const) {
    it(`drops pending shortcuts after an external ${change} change`, async () => {
      const nav = navigation();
      const first = deferred<SelectionData>();
      nav.query.mockReturnValueOnce(first.promise);
      nav.press();
      nav.press();
      if (change === 'owner') nav.update({ owner: {} });
      if (change === 'sheet') nav.update({ sheet: 1 });
      if (change === 'selection') nav.update({ selection: selectionAt({ row: 2, col: 2 }) });
      if (change === 'navigation') nav.update({ navigation: 1 });
      first.resolve({ edge: { row: 1, col: 4 } });
      await tick();
      expect(nav.apply).not.toHaveBeenCalled();
      expect(nav.query).toHaveBeenCalledTimes(1);
      nav.press();
      await tick();
      expect(nav.apply).toHaveBeenCalledTimes(1);
    });
  }

  it('retains queued presses while an asynchronous placement finishes', async () => {
    const queue = createSelectionDataNavigation();
    const placed = deferred<boolean>();
    let current = { owner: {}, sheet: 0, navigation: 0, version: 'v1', selection: selectionAt({ row: 0, col: 0 }) };
    const query = mock(async (context: typeof current): Promise<SelectionData> => ({
      edge: { row: 0, col: context.selection.focus.col + 2 },
    }));
    let target!: Selection;
    const options = {
      context: () => current,
      limits: () => ({ rows: 100, cols: 20, rowsPerPage: 10 }), query,
      apply: (selection: Selection) => {
        target = selection;
        current = { ...current, navigation: current.navigation + 1 };
        if (current.navigation === 1) return placed.promise;
        current = { ...current, selection };
        return true;
      },
    };
    queue.enqueue(right, options);
    queue.enqueue(right, options);
    await tick();
    current = { ...current, selection: structuredClone(target) };
    queue.enqueue(right, options);
    placed.resolve(true);
    await tick();
    expect(current.selection.focus).toEqual({ row: 0, col: 6 });
    expect(query).toHaveBeenCalledTimes(3);
  });
});

describe('selection equality', () => {
  it('compares range presence and anchors symmetrically', () => {
    const region = { ...selectionAt({ row: 2, col: 2 }), range: { top: 0, left: 0, bottom: 2, right: 2 } };
    const endpoints = { anchor: { row: 0, col: 0 }, focus: { row: 2, col: 2 } };
    const changedAnchor = { ...region, anchor: { row: 1, col: 1 } };
    for (const other of [endpoints, changedAnchor, { ...region, range: { ...region.range, right: 3 } }]) {
      expect(sameSelection(region, other)).toBe(false);
      expect(sameSelection(other, region)).toBe(false);
    }
    const collapsed = selectionAt({ row: 2, col: 2 });
    const explicit = { ...collapsed, range: { top: 2, left: 2, bottom: 2, right: 2 } };
    expect(sameSelection(collapsed, explicit)).toBe(false);
    expect(sameSelection(explicit, collapsed)).toBe(false);
    expect(sameSelection(region, structuredClone(region))).toBe(true);
    expect(sameSelection(endpoints, structuredClone(endpoints))).toBe(true);
  });
});

it('reads hidden spans outside the painted track indices', () => {
  const hidden = selectionHidden({ startRow: 0, startCol: 0, rowOffsets: [0, 24], colOffsets: [0, 96],
    hiddenRowSpans: [[1, 3], [100, 200]], hiddenColSpans: [[0, 0], [10, 20]] });
  for (const row of [1, 2, 3, 100, 200]) expect(hidden.isRowHidden(row)).toBe(true);
  for (const row of [0, 4, 99, 201]) expect(hidden.isRowHidden(row)).toBe(false);
  for (const col of [0, 10, 20]) expect(hidden.isColHidden(col)).toBe(true);
  for (const col of [1, 9, 21]) expect(hidden.isColHidden(col)).toBe(false);
  expect(selectionHidden(undefined).isRowHidden(1)).toBe(false);
});

it('keeps sheet indices reachable when hidden spans collapse the pixel extent', () => {
  const limits = deriveLimits({ width: 96, height: 48, commands: [], grid: {
    startRow: 0, startCol: 0, rowOffsets: [0, 24, 48], colOffsets: [0, 96],
    hiddenRowSpans: [[1, 899_999], [900_001, 1_048_575]], hiddenColSpans: [[1, 16_383]],
  } }, { contentHeight: 48, contentWidth: 96 }, 48);
  expect(limits).toEqual({ rows: 1_048_576, cols: 16_384, rowsPerPage: 2 });
});
