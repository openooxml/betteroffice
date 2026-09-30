import { GlobalRegistrator } from '@happy-dom/global-registrator';
import { afterAll, afterEach, beforeAll, beforeEach, expect, spyOn, test } from 'bun:test';
import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';
import type { Layout } from '@betteroffice/docx/layout/pagination';
import { preloadEditWasm } from '@betteroffice/docx/wasm/edit';
import { ResidentWorkerOutOfMemoryError } from '@betteroffice/docx/yrs';
import {
  revisionPreviewKey,
  revisionPreviewKeyOf,
  sourceVersionOf,
  stampSourceVersion,
} from '../internals/layoutProvenance';
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

test('a page build keeps the version of the layout it fills, not the session version', async () => {
  const { engine, inputs, host } = lazyFixture();
  try {
    let version = 'v1';
    Object.assign(host, { version: () => version });
    stampSourceVersion(inputs.layout, 'v1');
    const overrides = { getInputs: () => inputs };
    const { result, unmount } = renderHook(() =>
      useRustDisplayList(inputs.layout as Layout, overrides, undefined, undefined, host)
    );
    await waitFor(() => expect(result.current.frame).not.toBeNull());
    expect(sourceVersionOf(result.current.queries)).toBe('v1');
    version = 'v2';
    const pages = () => result.current.frame!.displayList.pages;
    const last = pages().length - 1;
    await act(async () => {
      result.current.setDisplayWindow(last, last + 1);
    });
    await waitFor(() => expect(pages()[last]?.unbuilt).toBeFalsy());
    expect(sourceVersionOf(result.current.queries)).toBe('v1');
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

test('a full worker rebuild retains built pages until released and rebuilds evicted pages on demand', async () => {
  const { engine, inputs, host } = lazyFixture();
  try {
    const overrides = { getInputs: () => inputs };
    const { result, rerender, unmount } = renderHook(
      ({ resolved }) =>
        useRustDisplayList(inputs.layout as Layout, overrides, undefined, resolved, host),
      { initialProps: { resolved: undefined as ReadonlySet<number> | undefined } }
    );
    await waitFor(() => expect(result.current.frame).not.toBeNull());
    const worker = EngineWorker.last!;
    const pages = () => result.current.frame!.displayList.pages;
    for (let round = 0; round < 50 && pages().some((page) => page.unbuilt); round += 1) {
      await waitFor(() => expect(idleCallbacks.size).toBeGreaterThan(0));
      await act(async () => runIdleCallbacks());
    }
    expect(pages().every((page) => !page.unbuilt)).toBe(true);
    const last = pages().length - 1;
    expect(last).toBeGreaterThanOrEqual(5);
    const before = result.current.frame!.frameEpoch;
    const requestsBeforeRebuild = worker.posted.length;
    await act(async () => {
      result.current.setRetainBuiltPages!(true);
      rerender({ resolved: new Set([1]) });
    });
    await waitFor(() => expect(result.current.frame!.frameEpoch).toBeGreaterThan(before));
    expect(worker.posted[requestsBeforeRebuild]).toMatchObject({
      type: 'buildFrame',
      displayWindow: [0, 5],
      retainBuiltPages: true,
      expectedFrameEpoch: before,
    });
    expect(pages().every((page) => !page.unbuilt)).toBe(true);
    expect(pages()[last]!.primitives.length).toBeGreaterThan(0);
    const retained = result.current.frame!.frameEpoch;
    const requestsBeforeRelease = worker.posted.length;
    await act(async () => {
      result.current.setRetainBuiltPages!(false);
      rerender({ resolved: new Set([2]) });
    });
    await waitFor(() => expect(result.current.frame!.frameEpoch).toBeGreaterThan(retained));
    expect(worker.posted[requestsBeforeRelease]).toMatchObject({
      type: 'buildFrame',
      displayWindow: [0, 5],
      expectedFrameEpoch: retained,
    });
    expect(worker.posted[requestsBeforeRelease]).not.toHaveProperty('retainBuiltPages');
    expect(pages().slice(0, 5).every((page) => !page.unbuilt)).toBe(true);
    expect(pages().slice(5).every((page) => page.unbuilt)).toBe(true);
    expect(pages()[last]!.unbuilt).toBe(true);
    const adopted = result.current.frame!.frameEpoch;
    const requestsBeforeScroll = worker.posted.length;
    await act(async () => {
      result.current.setDisplayWindow(last, last + 1);
    });
    await waitFor(() => expect(pages()[last]!.unbuilt).toBeFalsy());
    expect(
      worker.posted.slice(requestsBeforeScroll).filter((request) => request.type === 'buildPages')
    ).toEqual([expect.objectContaining({ pages: [last], expectedFrameEpoch: adopted })]);
    expect(pages()[last]!.primitives.length).toBeGreaterThan(0);
    expect(result.current.error).toBeNull();
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

test('page builds wait for the frame of an edit in flight', async () => {
  const { engine, inputs, host } = lazyFixture();
  try {
    const overrides = { getInputs: () => inputs };
    const { result, unmount } = renderHook(() =>
      useRustDisplayList(inputs.layout as Layout, overrides, undefined, undefined, host)
    );
    await waitFor(() => expect(result.current.frame).not.toBeNull());
    const worker = EngineWorker.last!;
    const pages = () => result.current.frame!.displayList.pages;
    const last = pages().length - 1;
    expect(pages()[last]!.unbuilt).toBe(true);
    const { paraId } = JSON.parse(engine.paragraphs('body'))[0] as { paraId: string };
    engine.set_selection('body', paraId, 1, paraId, 1);
    worker.holdInputReplies = true;
    let pendingEdit: ReturnType<typeof result.current.applyInput> | undefined;
    await act(async () => {
      pendingEdit = result.current.applyInput('New ');
    });
    await waitFor(() => expect(worker.heldInputReplies).toHaveLength(1));
    const pageBuilds = () => worker.posted.filter((request) => request.type === 'buildPages');
    await act(async () => {
      runIdleCallbacks();
      result.current.setDisplayWindow(last, last + 1);
      await new Promise((resolve) => setTimeout(resolve, 150));
    });
    expect(pageBuilds()).toEqual([]);

    worker.holdInputReplies = false;
    await act(async () => {
      worker.releaseInputReplies();
      expect(await pendingEdit!).not.toBeNull();
    });
    const adopted = result.current.frame!.frameEpoch;
    await waitFor(() => expect(pages()[last]!.unbuilt).toBeFalsy());
    expect(pageBuilds()[0]).toMatchObject({ pages: [last], expectedFrameEpoch: adopted });
    expect(result.current.error).toBeNull();
    unmount();
  } finally {
    engine.free();
  }
});

test.each([false, true])(
  'background pages recover after a newer worker frame stays unadopted (idle retry: %p)',
  async (idle) => {
    const { engine, inputs, host } = lazyFixture();
    let now = performance.now();
    const clock = spyOn(performance, 'now').mockImplementation(() => now);
    try {
      const overrides = { getInputs: () => inputs };
      const { result, unmount } = renderHook(() =>
        useRustDisplayList(inputs.layout as Layout, overrides, undefined, undefined, host)
      );
      await waitFor(() => expect(result.current.frame).not.toBeNull());
      const worker = EngineWorker.last!;
      const bootstrap = worker.posted.find((request) => request.type === 'bootstrap')!;
      const adopted = result.current.frame!.frameEpoch;
      const pageBuilds = () => worker.posted.filter((request) => request.type === 'buildPages');
      if (idle) await waitFor(() => expect(idleCallbacks.size).toBeGreaterThan(0));
      await act(async () => {
        worker.postMessage({
          id: 0,
          type: 'buildFrame',
          extras: bootstrap.extras,
          expectedFrameEpoch: adopted,
          paintCaret: false,
          displayWindow: [0, 5],
        });
        await Promise.resolve();
        if (idle) runIdleCallbacks();
        else {
          result.current.setDisplayWindow(0, 4);
          await new Promise((resolve) => setTimeout(resolve, 75));
        }
      });
      expect(result.current.frame!.frameEpoch).toBe(adopted);
      expect(pageBuilds()).toEqual([]);
      now += 2001;
      await waitFor(async () => {
        if (!idle) await act(async () => runIdleCallbacks());
        expect(pageBuilds().length).toBeGreaterThan(0);
      });
      expect(pageBuilds()[0]).toMatchObject({ expectedFrameEpoch: adopted });
      await waitFor(() => expect(result.current.frame!.displayList.pages[5]!.unbuilt).toBeFalsy());
      expect(result.current.frame!.displayList.pages[5]!.primitives.length).toBeGreaterThan(0);
      expect(result.current.error).toBeNull();
      unmount();
    } finally {
      clock.mockRestore();
      engine.free();
    }
  }
);

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
      result.current.setRetainBuiltPages!(true);
      result.current.setDisplayWindow(last, last + 1);
    });
    await waitFor(() => expect(relayouts).toBe(1));
    expect(first.terminated).toBe(true);
    expect(result.current.frame).not.toBeNull();
    expect(result.current.error).toBeNull();

    await act(async () => rerender({ layout: { ...inputs.layout } }));
    await waitFor(() => expect(EngineWorker.spawned).toBe(2));
    const second = EngineWorker.last!;
    expect(second.posted[0]).toMatchObject({ type: 'bootstrap', retainBuiltPages: true });
    await waitFor(() => expect(result.current.workerSurfacesActive).toBe(true));
    await act(async () => {
      const unbuilt = result.current.frame!.displayList.pages.findIndex((page) => page.unbuilt);
      expect(unbuilt).toBeGreaterThanOrEqual(0);
      result.current.setDisplayWindow(unbuilt, unbuilt + 1);
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
