import { GlobalRegistrator } from '@happy-dom/global-registrator';
import { afterAll, afterEach, beforeAll, beforeEach, expect, spyOn, test } from 'bun:test';
import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';
import type { Layout } from '@betteroffice/docx/layout/pagination';
import { preloadEditWasm } from '@betteroffice/docx/wasm/edit';
import { ResidentWorkerOutOfMemoryError } from '@betteroffice/docx/yrs';
import { revisionPreviewKey, revisionPreviewKeyOf } from '../internals/layoutProvenance';
import { useRustDisplayList } from './useDisplayList';
import { EngineWorker, lazyFixture, PREVIEW } from './__fixtures__/lazyPages';

const ownsDom = !GlobalRegistrator.isRegistered;
if (ownsDom) GlobalRegistrator.register();
const { act, cleanup, renderHook, waitFor } = await import('@testing-library/react');
const originalWorker = globalThis.Worker;

beforeAll(() =>
  preloadEditWasm(
    new Uint8Array(
      readFileSync(
        resolve(import.meta.dir, '../../../../../docx/src/wasm/generated/edit/docx_edit_bg.wasm')
      )
    )
  )
);

const originalIdle = globalThis.requestIdleCallback;
const originalCancelIdle = globalThis.cancelIdleCallback;
let idleCallbacks = new Map<number, () => void>();
let nextIdle = 1;

beforeEach(() => {
  idleCallbacks = new Map();
  globalThis.requestIdleCallback = ((callback: () => void) => {
    const id = nextIdle++;
    idleCallbacks.set(id, callback);
    return id;
  }) as typeof requestIdleCallback;
  globalThis.cancelIdleCallback = ((id: number) => {
    idleCallbacks.delete(id);
  }) as typeof cancelIdleCallback;
});

afterEach(() => {
  cleanup();
  globalThis.Worker = originalWorker;
  globalThis.requestIdleCallback = originalIdle;
  globalThis.cancelIdleCallback = originalCancelIdle;
});

function runIdleCallbacks(): void {
  const pending = [...idleCallbacks.values()];
  idleCallbacks.clear();
  for (const callback of pending) callback();
}

afterAll(async () => {
  if (ownsDom) await GlobalRegistrator.unregister();
});

test('a worker frame builds only the pages near the viewport', async () => {
  const { engine, inputs, host } = lazyFixture();
  try {
    const overrides = { getInputs: () => inputs };
    const { result, unmount } = renderHook(() =>
      useRustDisplayList(inputs.layout as Layout, overrides, undefined, undefined, host)
    );
    await waitFor(() => expect(result.current.frame).not.toBeNull());
    const pages = () => result.current.frame!.displayList.pages;
    const first = pages();
    expect(first.length).toBeGreaterThan(5);
    expect(first.slice(0, 5).every((page) => !page.unbuilt)).toBe(true);
    expect(EngineWorker.last!.posted[0]).toMatchObject({ displayWindow: [0, 5] });

    const last = first.length - 1;
    await act(async () => {
      result.current.setDisplayWindow(last, last + 1);
    });
    await waitFor(() => expect(pages()[last]?.unbuilt).toBeFalsy());
    expect(pages().slice(5, last).every((page) => page.unbuilt)).toBe(true);
    expect(revisionPreviewKeyOf(result.current.queries)).toBe(revisionPreviewKey(PREVIEW));
    const span = pages()[5]!.positionSpan!;
    expect(span[0]).toBeLessThanOrEqual(span[1]);
    expect(
      EngineWorker.last!.posted.filter((request) => request.type === 'buildPages')
    ).toEqual([expect.objectContaining({ pages: [last] })]);

    let settled: Awaited<ReturnType<typeof result.current.settledDisplayList>> | undefined;
    await act(async () => {
      settled = await result.current.settledDisplayList(() => {});
    });
    expect(settled!.pages.some((page) => page.unbuilt)).toBe(false);
    const full = JSON.parse(engine.build_display_list_json(JSON.stringify(inputs))) as {
      pages: unknown[];
    };
    expect(settled!.pages.map((page) => page.primitives.length)).toEqual(
      full.pages.map((page) => (page as { primitives: unknown[] }).primitives.length)
    );
    unmount();
  } finally {
    engine.free();
  }
});

test('pages away from the viewport build in batches while the main thread idles', async () => {
  const { engine, inputs, host } = lazyFixture();
  try {
    const overrides = { getInputs: () => inputs };
    const { result, unmount } = renderHook(() =>
      useRustDisplayList(inputs.layout as Layout, overrides, undefined, undefined, host)
    );
    await waitFor(() => expect(result.current.frame).not.toBeNull());
    const unbuilt = () => result.current.frame!.displayList.pages.filter((page) => page.unbuilt);
    expect(unbuilt().length).toBeGreaterThan(0);
    for (let round = 0; round < 50 && unbuilt().length > 0; round += 1) {
      await waitFor(() => expect(idleCallbacks.size).toBeGreaterThan(0));
      await act(async () => runIdleCallbacks());
    }
    expect(unbuilt()).toEqual([]);
    const batches = EngineWorker.last!.posted.filter((request) => request.type === 'buildPages');
    expect(batches.length).toBeGreaterThan(0);
    for (const batch of batches) {
      if (batch.type === 'buildPages') expect(batch.pages.length).toBeLessThanOrEqual(16);
    }
    unmount();
  } finally {
    engine.free();
  }
});

test('an edit schedules idle rebuilds for formerly built pages away from the viewport', async () => {
  const { engine, inputs, host } = lazyFixture();
  try {
    const overrides = { getInputs: () => inputs };
    const { result, unmount } = renderHook(() =>
      useRustDisplayList(inputs.layout as Layout, overrides, undefined, undefined, host)
    );
    await waitFor(() => expect(result.current.frame).not.toBeNull());
    await act(async () => {
      await result.current.settledDisplayList(null);
    });
    const pages = () => result.current.frame!.displayList.pages;
    const pageCount = pages().length;
    expect(pageCount).toBeGreaterThan(5);
    expect(pages().every((page) => !page.unbuilt)).toBe(true);
    const last = pageCount - 1;
    await act(async () => {
      result.current.setDisplayWindow(last, last + 1);
    });
    const requestsBeforeEdit = EngineWorker.last!.posted.length;
    const { paraId, text } = JSON.parse(engine.paragraphs('body'))[0] as {
      paraId: string;
      text: string;
    };
    const offset = text.length - 10;
    engine.set_selection('body', paraId, offset, paraId, offset);

    await act(async () => {
      expect(await result.current.applyInput('New ')).not.toBeNull();
    });
    expect(result.current.error).toBeNull();
    expect(pages()).toHaveLength(pageCount);
    expect(pages()[last]!.unbuilt).toBeFalsy();
    expect(result.current.caret?.caretRect?.pageIndex).toBe(last);
    expect(EngineWorker.last!.posted[requestsBeforeEdit]).toMatchObject({
      type: 'applyInput',
      displayWindow: [last, last + 1],
    });
    expect(
      pages().slice(0, last).every((page) => page.unbuilt && page.primitives.length === 0)
    ).toBe(true);
    const rebuilds = () =>
      EngineWorker.last!.posted
        .slice(requestsBeforeEdit)
        .filter((request) => request.type === 'buildPages');
    expect(rebuilds()).toEqual([]);
    for (let round = 0; round < 5 && rebuilds().length === 0; round += 1) {
      await waitFor(() => expect(idleCallbacks.size).toBeGreaterThan(0));
      await act(async () => runIdleCallbacks());
    }
    expect(rebuilds()).toHaveLength(1);
    const rebuilt = rebuilds()[0]!;
    expect(rebuilt.pages.length).toBeGreaterThan(0);
    expect(rebuilt.pages.length).toBeLessThanOrEqual(16);
    expect(rebuilt.pages.every((index) => index < last)).toBe(true);
    await act(async () => {
      const settled = await result.current.settledDisplayList(null);
      expect(settled.pages.some((page) => page.unbuilt)).toBe(false);
    });
    unmount();
  } finally {
    engine.free();
  }
});

test('a failed page build hands rendering back to the main thread', async () => {
  const { engine, inputs, host } = lazyFixture();
  const errors = spyOn(console, 'error').mockImplementation(() => {});
  try {
    EngineWorker.failPageBuilds = true;
    let relayouts = 0;
    const overrides = { getInputs: () => inputs };
    const { result, unmount } = renderHook(() =>
      useRustDisplayList(inputs.layout as Layout, overrides, undefined, undefined, host, () => {
        relayouts += 1;
      })
    );
    await waitFor(() => expect(result.current.frame).not.toBeNull());
    const last = result.current.frame!.displayList.pages.length - 1;
    await act(async () => {
      result.current.setDisplayWindow(last, last + 1);
    });
    await waitFor(() => expect(relayouts).toBe(1));
    expect(result.current.frame).toBeNull();
    expect(result.current.workerSurfacesActive).toBe(false);
    unmount();
  } finally {
    errors.mockRestore();
    engine.free();
  }
});

test('a page build out of memory restarts the worker once, then reports without a main-thread fallback', async () => {
  const { engine, inputs, host } = lazyFixture();
  const warnings = spyOn(console, 'warn').mockImplementation(() => {});
  const errors = spyOn(console, 'error').mockImplementation(() => {});
  try {
    EngineWorker.outOfMemoryPageBuilds = true;
    let relayouts = 0;
    const overrides = { getInputs: () => inputs };
    const { result, rerender, unmount } = renderHook(
      ({ layout }) =>
        useRustDisplayList(layout, overrides, undefined, undefined, host, () => {
          relayouts += 1;
        }),
      { initialProps: { layout: inputs.layout as Layout } }
    );
    await waitFor(() => expect(result.current.frame).not.toBeNull());
    const first = EngineWorker.last!;
    const last = result.current.frame!.displayList.pages.length - 1;
    await act(async () => {
      result.current.setDisplayWindow(last, last + 1);
    });
    await waitFor(() => expect(relayouts).toBe(1));
    expect(first.terminated).toBe(true);
    expect(result.current.frame).not.toBeNull();
    expect(result.current.error).toBeNull();

    await act(async () => rerender({ layout: { ...inputs.layout } }));
    await waitFor(() => expect(EngineWorker.spawned).toBe(2));
    const second = EngineWorker.last!;
    expect(second.posted[0]).toMatchObject({ type: 'bootstrap' });
    await waitFor(() => expect(result.current.workerSurfacesActive).toBe(true));
    await act(async () => {
      const middle = Math.floor(last / 2);
      result.current.setDisplayWindow(middle, middle + 1);
    });
    await waitFor(() => expect(result.current.error).toBeInstanceOf(ResidentWorkerOutOfMemoryError));
    expect(relayouts).toBe(1);
    expect(EngineWorker.spawned).toBe(2);
    expect(result.current.frame).not.toBeNull();
    unmount();
  } finally {
    warnings.mockRestore();
    errors.mockRestore();
    engine.free();
  }
});
