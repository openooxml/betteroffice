import { describe, expect, it, mock } from 'bun:test';
import { normalizeRange, selectionAt } from '@betteroffice/xlsx';
import type { KeyInput, Selection, SelectionData } from '@betteroffice/xlsx';
import { deriveLimits } from './viewer/sessionGeometry';
import { MAX_CLEAR_CELLS, clearedCells, createSelectionDataNavigation, sameSelection, selectionDataRequest, selectionHidden } from './selectionData';

function deferred<T>() {
  let resolve!: (value: T) => void;
  let reject!: (error: unknown) => void;
  const promise = new Promise<T>((done, fail) => { resolve = done; reject = fail; });
  return { promise, resolve, reject };
}

const tick = () => new Promise<void>((resolve) => setTimeout(resolve, 0));
const right = { key: 'ArrowRight', ctrlKey: true };

function navigation(anchorActive = false) {
  let current = { owner: {}, sheet: 0, navigation: 0, version: 'v1', selection: selectionAt({ row: 1, col: 1 }) };
  let open = true;
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
    setOpen(value: boolean) { open = value; },
    query, apply,
    press: (key: KeyInput = right) => queue.enqueue(key, {
      activeCell: anchorActive ? (selection) => selection.anchor : undefined,
      context: () => open ? current : null, limits: () => ({ rows: 100, cols: 20, rowsPerPage: 10 }), query, apply,
    }),
    defer: (task: () => void, behindActions?: boolean, cancel?: () => void) =>
      queue.defer(current, task, behindActions, cancel),
  };
}

describe('Ctrl+A sheet selection clears', () => {
  const sheet = { top: 0, left: 0, bottom: 1_048_575, right: 16_383 };

  it('refuses a sparse used range spanning A1:XFD1048576 before expanding it', () => {
    expect(clearedCells(sheet, () => sheet)).toBeNull();
  });

  it('applies the host cell limit to the intersection with the used range', () => {
    const column = { ...sheet, right: 0 };
    expect(clearedCells(column, () => ({ ...sheet, bottom: MAX_CLEAR_CELLS - 1 })))
      .toHaveLength(MAX_CLEAR_CELLS);
    expect(clearedCells(column, () => ({ ...sheet, bottom: MAX_CLEAR_CELLS }))).toBeNull();
    expect(clearedCells({ ...column, top: 1 }, () => ({ ...sheet, bottom: 0 }))).toEqual([]);
  });

  it('preserves clears larger than the limit when the selection does not reach a sheet edge', () => {
    const used = mock(() => sheet);
    expect(clearedCells({ top: 0, left: 0, bottom: MAX_CLEAR_CELLS, right: 0 }, used))
      .toHaveLength(MAX_CLEAR_CELLS + 1);
    expect(used).not.toHaveBeenCalled();
  });
});

describe('selection data navigation', () => {
  it('requests the used range for Ctrl/Meta End and the row end for plain End', () => {
    expect(selectionDataRequest({ key: 'End' })).toBe('rowEnd');
    expect(selectionDataRequest({ key: 'End', shiftKey: true })).toBe('rowEnd');
    expect(selectionDataRequest({ key: 'Home' })).toBeNull();
    for (const modifier of [{ ctrlKey: true }, { metaKey: true }]) {
      expect(selectionDataRequest({ key: 'End', ...modifier })).toBe('usedRange');
      expect(selectionDataRequest({ key: 'End', shiftKey: true, ...modifier })).toBe('usedRange');
    }
  });

  it('moves plain End to the queried row end and extends it with Shift', async () => {
    const nav = navigation(true);
    nav.update({ selection: { anchor: { row: 1, col: 1 }, focus: { row: 4, col: 3 } } });
    nav.query.mockResolvedValueOnce({ edge: { row: 4, col: 12 } }).mockResolvedValueOnce({ edge: { row: 1, col: 9 } });
    nav.press({ key: 'End', shiftKey: true });
    nav.press({ key: 'End' });
    await tick();
    expect(nav.query.mock.calls.map(([context, direction]) => [context.selection.focus, direction])).toEqual([
      [{ row: 4, col: 3 }, 'rowEnd'],
      [{ row: 1, col: 1 }, 'rowEnd'],
    ]);
    expect(nav.apply.mock.calls[0]![0]).toEqual({ anchor: { row: 1, col: 1 }, focus: { row: 4, col: 12 } });
    expect(nav.current.selection).toEqual(selectionAt({ row: 1, col: 9 }));
  });

  it('chains a used-range jump and a data-edge jump through one asynchronous batch', async () => {
    const nav = navigation();
    const first = deferred<SelectionData>();
    nav.query.mockReturnValueOnce(first.promise).mockResolvedValueOnce({ edge: { row: 79, col: 0 } });
    nav.press({ key: 'End', ctrlKey: true });
    nav.press({ key: 'ArrowLeft', ctrlKey: true });
    expect(nav.query.mock.calls[0]![1]).toBe('usedRange');
    first.resolve({ usedRange: { top: 0, left: 0, bottom: 79, right: 15 }, region: null });
    await tick();
    expect(nav.apply.mock.calls[0]![0]).toEqual(selectionAt({ row: 79, col: 15 }));
    expect(nav.query.mock.calls[1]![0].selection.focus).toEqual({ row: 79, col: 15 });
    expect(nav.query.mock.calls[1]![1]).toBe('left');
    expect(nav.current.selection).toEqual(selectionAt({ row: 79, col: 0 }));
  });

  it('refreshes a pending used-range answer after edits and preserves Shift extension', async () => {
    const nav = navigation(true);
    const first = deferred<SelectionData>();
    nav.query.mockReturnValueOnce(first.promise).mockResolvedValueOnce({
      usedRange: { top: 0, left: 0, bottom: 99, right: 19 },
    });
    nav.press({ key: 'End', metaKey: true, shiftKey: true });
    nav.update({ version: 'v2' });
    first.resolve({ usedRange: { top: 0, left: 0, bottom: 79, right: 15 } });
    await tick();
    expect(nav.query).toHaveBeenCalledTimes(2);
    expect(nav.apply).toHaveBeenCalledTimes(1);
    expect(nav.current.selection).toEqual({ anchor: { row: 1, col: 1 }, focus: { row: 99, col: 19 } });
  });

  it('leaves selection unchanged after a failed or unavailable used-range query', async () => {
    for (const answer of [Promise.reject(new Error('Query failed')), Promise.resolve({ region: null })]) {
      const nav = navigation();
      nav.query.mockReturnValueOnce(answer);
      nav.press({ key: 'End', ctrlKey: true });
      await tick();
      expect(nav.apply).not.toHaveBeenCalled();
      expect(nav.current.selection).toEqual(selectionAt({ row: 1, col: 1 }));
    }
  });

  it('returns to A1 only after a successful empty used-range query', async () => {
    const nav = navigation();
    nav.query.mockResolvedValueOnce({ usedRange: null });
    nav.press({ key: 'End', ctrlKey: true });
    await tick();
    expect(nav.current.selection).toEqual(selectionAt({ row: 0, col: 0 }));
  });

  it('extends from the focus then jumps from the active anchor in one queued batch', async () => {
    const nav = navigation(true);
    nav.update({ selection: { anchor: { row: 1, col: 1 }, focus: { row: 4, col: 3 } } });
    nav.query.mockImplementation(async (context) => ({
      edge: { row: context.selection.focus.row, col: context.selection.focus.col + 2 },
    }));
    nav.press({ ...right, shiftKey: true });
    nav.press(right);
    await tick();
    expect(nav.query.mock.calls[0]![0].selection.focus).toEqual({ row: 4, col: 3 });
    expect(nav.apply.mock.calls[0]![0]).toEqual({ anchor: { row: 1, col: 1 }, focus: { row: 4, col: 5 } });
    expect(nav.query.mock.calls[1]![0].selection.focus).toEqual({ row: 1, col: 1 });
    expect(nav.current.selection).toEqual(selectionAt({ row: 1, col: 3 }));
  });

  it('preserves the active anchor and range comparison across select-all presses', async () => {
    const nav = navigation(true);
    const anchor = { row: 2, col: 2 };
    nav.update({ selection: { anchor, focus: { row: 4, col: 4 } } });
    const region = { top: 0, left: 0, bottom: 4, right: 4 };
    nav.query.mockResolvedValue({ region });
    nav.press({ key: 'a', ctrlKey: true });
    nav.press({ key: 'a', ctrlKey: true });
    await tick();
    expect(nav.query.mock.calls[0]![0].selection.focus).toEqual(anchor);
    expect(normalizeRange(nav.apply.mock.calls[0]![0])).toEqual(region);
    expect(normalizeRange(nav.current.selection)).toEqual({ top: 0, left: 0, bottom: 99, right: 19 });
    expect(nav.current.selection.anchor).toEqual(anchor);
  });

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

  it('runs later actions in order on the selection each pending jump lands on', async () => {
    const nav = navigation();
    const first = deferred<SelectionData>();
    nav.query.mockReturnValueOnce(first.promise).mockResolvedValueOnce({ edge: { row: 1, col: 9 } });
    const seen: unknown[] = [];
    nav.press();
    expect(nav.defer(() => seen.push(nav.current.selection.focus))).toBe(true);
    nav.press();
    expect(nav.defer(() => seen.push(nav.current.selection.focus))).toBe(true);
    expect(seen).toEqual([]);
    first.resolve({ edge: { row: 1, col: 4 } });
    await tick();
    expect(seen).toEqual([{ row: 1, col: 4 }, { row: 1, col: 9 }]);
    expect(nav.defer(() => seen.push('idle'))).toBe(false);
  });

  it('queues later moves only behind an already deferred action', async () => {
    const nav = navigation();
    const first = deferred<SelectionData>();
    nav.query.mockReturnValueOnce(first.promise);
    nav.press();
    expect(nav.defer(() => {}, true)).toBe(false);
    expect(nav.defer(() => {})).toBe(true);
    expect(nav.defer(() => {}, true)).toBe(true);
    first.resolve({ edge: null });
    await tick();
  });

  it('runs every action queued behind one that ends the jump queue', async () => {
    const nav = navigation();
    const first = deferred<SelectionData>();
    nav.query.mockReturnValueOnce(first.promise);
    const seen: string[] = [];
    nav.press();
    nav.defer(() => { seen.push('open'); nav.setOpen(false); });
    nav.defer(() => seen.push('b'), true);
    nav.defer(() => seen.push('Enter'), true);
    first.resolve({ edge: { row: 1, col: 4 } });
    await tick();
    expect(seen).toEqual(['open', 'b', 'Enter']);
  });

  it('resumes queued jumps after the action that ended the queue hands the grid back', async () => {
    const nav = navigation();
    const first = deferred<SelectionData>();
    nav.query.mockReturnValueOnce(first.promise).mockResolvedValueOnce({ edge: { row: 9, col: 4 } });
    const seen: unknown[] = [];
    nav.press();
    nav.defer(() => nav.setOpen(false));
    nav.defer(() => nav.setOpen(true), true);
    nav.press({ key: 'ArrowDown', ctrlKey: true });
    nav.defer(() => seen.push(nav.current.selection.focus), true);
    first.resolve({ edge: { row: 1, col: 4 } });
    await tick();
    expect(seen).toEqual([{ row: 9, col: 4 }]);
  });

  it('cancels queued actions instead of running them when the selection moves away', async () => {
    const nav = navigation();
    const first = deferred<SelectionData>();
    nav.query.mockReturnValueOnce(first.promise);
    const seen: string[] = [];
    nav.press();
    nav.defer(() => seen.push('run'), false, () => seen.push('cancel'));
    nav.update({ sheet: 1 });
    first.resolve({ edge: { row: 1, col: 4 } });
    await tick();
    expect(seen).toEqual(['cancel']);
  });

  it('runs a later action at once when the selection moved away from a pending jump', async () => {
    const nav = navigation();
    const first = deferred<SelectionData>();
    nav.query.mockReturnValueOnce(first.promise);
    nav.press();
    nav.update({ navigation: 1, selection: selectionAt({ row: 7, col: 7 }) });
    expect(nav.defer(() => {})).toBe(false);
    first.resolve({ edge: { row: 1, col: 4 } });
    await tick();
    expect(nav.current.selection.focus).toEqual({ row: 7, col: 7 });
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
  } }, 48);
  expect(limits).toEqual({ rows: 1_048_576, cols: 16_384, rowsPerPage: 2 });
});
