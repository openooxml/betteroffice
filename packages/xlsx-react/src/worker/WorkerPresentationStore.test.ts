import { describe, expect, it, mock } from 'bun:test';
import type { WorkbookSheetView } from '@betteroffice/xlsx';
import type { WorkerPaintResult } from './PeerlessWorkerPaintSource';
import { WorkerPresentationStore } from './WorkerPresentationStore';

const view: WorkbookSheetView = { sheet: 0, version: 'v1', contentWidth: 4000, contentHeight: 6000,
  frozenRows: 0, frozenCols: 0, frozenWidth: 0, frozenHeight: 0, initialScrollX: 0, initialScrollY: 0 };
const paint: WorkerPaintResult = {
  source: 'worker', version: 'v1', sequence: 1, coversThroughSeq: 1, mergedRanges: [], geometry: undefined,
  displayList: { width: 800, height: 600, commands: [] },
  request: { generation: 1, sheet: 0, navigation: 0, width: 800, height: 600, zoom: 1, dpr: 1,
    viewport: { x: 0, y: 0, width: 800, height: 600 } },
};

describe('worker presentation store', () => {
  it('advances coverage without replacing the presentation snapshot', async () => {
    const store = new WorkerPresentationStore();
    store.setPainted(paint);
    await Promise.resolve();
    const snapshot = store.getSnapshot();
    const listener = mock(() => {});
    const coverage = mock(() => {});
    store.subscribe(listener);
    const unsubscribe = store.subscribeCoverage(coverage);
    store.advanceCoverage(2);
    store.advanceCoverage(1);
    await Promise.resolve();
    expect(store.coversThroughSeq).toBe(2);
    expect(store.getSnapshot()).toBe(snapshot);
    expect(store.getSnapshot().painted).toBe(paint);
    expect(listener).not.toHaveBeenCalled();
    expect(coverage).toHaveBeenCalledTimes(1);
    unsubscribe();
    store.advanceCoverage(3);
    expect(coverage).toHaveBeenCalledTimes(1);
    store.reset();
    expect(store.coversThroughSeq).toBe(0);
  });

  it('coalesces frame and geometry publications while keeping synchronous reads current', async () => {
    const store = new WorkerPresentationStore();
    const listener = mock(() => {});
    store.subscribe(listener);
    store.setPainted(paint);
    store.setView(view);
    expect(store.viewRef.current).toBe(view);
    expect(listener).not.toHaveBeenCalled();
    await Promise.resolve();
    expect(listener).toHaveBeenCalledTimes(1);
    const snapshot = store.getSnapshot();
    store.setView({ ...view, version: 'v2' });
    await Promise.resolve();
    expect(store.viewRef.current?.version).toBe('v2');
    expect(store.getSnapshot()).toBe(snapshot);
    expect(listener).toHaveBeenCalledTimes(1);
  });

  for (const field of ['sheet', 'contentWidth', 'contentHeight', 'frozenRows', 'frozenCols',
    'frozenWidth', 'frozenHeight', 'initialScrollX', 'initialScrollY'] as const) {
    it(`publishes changed ${field}`, async () => {
      const store = new WorkerPresentationStore();
      store.setView(view);
      await Promise.resolve();
      const listener = mock(() => {});
      store.subscribe(listener);
      store.setView({ ...view, [field]: view[field] + 1 });
      await Promise.resolve();
      expect(listener).toHaveBeenCalledTimes(1);
    });
  }

  it('publishes selection immediately and cancels retired publications', async () => {
    const store = new WorkerPresentationStore();
    const listener = mock(() => {});
    const unsubscribe = store.subscribe(listener);
    store.setPainted(paint);
    const selection = { anchor: { row: 0, col: 0 }, focus: { row: 1, col: 0 } };
    store.setSelection(selection);
    expect(store.selectionRef.current).toBe(selection);
    expect(listener).toHaveBeenCalledTimes(1);
    await Promise.resolve();
    expect(listener).toHaveBeenCalledTimes(1);
    store.setView(view);
    store.cancel();
    await Promise.resolve();
    expect(listener).toHaveBeenCalledTimes(1);
    store.reset();
    expect(store.getSnapshot()).toEqual({ selection: null, view: null, painted: null });
    unsubscribe();
    store.setSelection(selection);
    expect(listener).toHaveBeenCalledTimes(2);
  });
});
