import { GlobalRegistrator } from '@happy-dom/global-registrator';
import { afterAll, afterEach, beforeAll, beforeEach, expect, spyOn, test } from 'bun:test';
import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';
import type { Layout } from '@betteroffice/docx/layout/pagination';
import { decodeFrameDelta } from '@betteroffice/docx/layout/render';
import { preloadEditWasm } from '@betteroffice/docx/wasm/edit';
import { ResidentWorkerOutOfMemoryError } from '@betteroffice/docx/yrs';
import {
  revisionPreviewKey,
  revisionPreviewKeyOf,
  sourceVersionOf,
  stampRevisionPreviewKey,
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

function runIdleCallbacks(timeRemaining = () => 50): void {
  const pending = [...idleCallbacks.values()];
  idleCallbacks.clear();
  for (const callback of pending) (callback as IdleRequestCallback)({ didTimeout: false, timeRemaining });
}

async function settleWithIdle(
  display: ReturnType<typeof useRustDisplayList>,
  relayout: (() => void) | null = null
) {
  let done = false;
  const pending = display.settledDisplayList(relayout);
  void pending.then(() => { done = true; }, () => { done = true; });
  for (let round = 0; round < 300 && !done; round += 1) {
    runIdleCallbacks();
    await new Promise((resolve) => setTimeout(resolve, 10));
  }
  return pending;
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
      settled = await settleWithIdle(result.current, () => {});
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

test('default-mode idle page builds attach immediately without background work', async () => {
  const { engine, inputs, host } = lazyFixture(100);
  const overrides = { getInputs: () => inputs };
  const hook = renderHook(() => useRustDisplayList(
    inputs.layout as Layout, overrides, undefined, undefined, host
  ));
  try {
    await waitFor(() => expect(hook.result.current.frame).not.toBeNull());
    const worker = EngineWorker.last!;
    const before = hook.result.current.frame!;
    await waitFor(() => expect(idleCallbacks.size).toBeGreaterThan(0));
    await act(async () => runIdleCallbacks());
    const builds = worker.posted.filter((request) => request.type === 'buildPages');
    expect(builds).toHaveLength(1);
    expect(builds[0]!.background).toBeUndefined();
    expect(builds[0]!.pages).toHaveLength(16);
    expect(hook.result.current.frame!.frameEpoch).toBe(before.frameEpoch + 1);
    expect(builds[0]!.pages.every(
      (index) => !hook.result.current.displayList!.pages[index]!.unbuilt
    )).toBe(true);
    expect(hook.result.current.error).toBeNull();
  } finally {
    hook.unmount();
    engine.free();
  }
});

test.each([
  ['advances the default-mode frame', false],
  ['leaves the worker-open frame unchanged', true],
] as const)('a page-build reply after a keystroke %s', async (_, workerOpen) => {
  const { engine, inputs, host } = lazyFixture(100);
  const updateListeners = new Set<(update: Uint8Array) => void>();
  Object.assign(host, {
    onUpdate: (listener: (update: Uint8Array) => void) => {
      updateListeners.add(listener);
      return () => updateListeners.delete(listener);
    },
  });
  const overrides = { getInputs: () => inputs };
  const hook = renderHook(() => useRustDisplayList(
    inputs.layout as Layout, overrides, undefined, undefined, host,
    undefined, undefined, undefined, workerOpen
  ));
  try {
    await waitFor(() => expect(hook.result.current.frame).not.toBeNull());
    const worker = EngineWorker.last!;
    const before = hook.result.current.frame!;
    worker.holdPageBuilds = true;
    await waitFor(() => expect(idleCallbacks.size).toBeGreaterThan(0));
    await act(async () => runIdleCallbacks());
    expect(worker.heldPageBuilds).toHaveLength(1);
    const build = worker.posted.find((request) => request.type === 'buildPages')!;
    expect(build.background).toBe(workerOpen ? true : undefined);
    const replyEpoch = JSON.parse(engine.resident_caret_snapshot_json()).frameEpoch as number;
    expect(replyEpoch).toBeGreaterThan(before.frameEpoch);
    expect(build.pages.every((index) => before.displayList.pages[index]!.unbuilt)).toBe(true);
    expect(updateListeners.size).toBeGreaterThan(0);
    await act(async () => {
      for (const listener of updateListeners) listener(new Uint8Array());
      worker.releasePageBuilds();
    });
    const adoptedEpoch = workerOpen ? before.frameEpoch : replyEpoch;
    expect(hook.result.current.frame!.frameEpoch).toBe(adoptedEpoch);
    expect(build.pages.every(
      (index) => Boolean(hook.result.current.displayList!.pages[index]!.unbuilt) === workerOpen
    )).toBe(true);
    if (workerOpen) expect(hook.result.current.frame).toBe(before);
    else expect(hook.result.current.queries).toBeNull();

    // Answer sync through the fake's existing frame builder.
    const postMessage = worker.postMessage.bind(worker);
    worker.postMessage = (request) => {
      postMessage(request.type === 'sync' ? { ...request, type: 'buildFrame' } : request);
      worker.posted[worker.posted.length - 1] = request;
    };
    Object.assign(host, { residentWorkerProbe: () => ({ layoutRevision: 2 }) });
    const requestsBeforeRelayout = worker.posted.length;
    await act(async () => {
      inputs.layout = { ...inputs.layout };
      hook.rerender();
    });
    await waitFor(() => {
      expect(hook.result.current.frame!.frameEpoch).toBeGreaterThan(adoptedEpoch);
      expect(hook.result.current.queries).not.toBeNull();
    });
    expect(worker.posted.slice(requestsBeforeRelayout).find(
      (request) => 'expectedFrameEpoch' in request
    )).toMatchObject({ type: 'sync', expectedFrameEpoch: adoptedEpoch });

    const { paraId } = JSON.parse(engine.paragraphs('body'))[0] as { paraId: string };
    engine.set_selection('body', paraId, 1, paraId, 1);
    const requestsBeforeInput = worker.posted.length;
    const typingEpoch = hook.result.current.frame!.frameEpoch;
    await act(async () => {
      expect(await hook.result.current.applyInput('Next ')).not.toBeNull();
    });
    expect(worker.posted.slice(requestsBeforeInput).filter(
      (request) => request.type === 'applyInput'
    )).toEqual([expect.objectContaining({ expectedFrameEpoch: typingEpoch })]);
    expect(hook.result.current.error).toBeNull();
  } finally {
    hook.unmount();
    engine.free();
  }
});

test('background posting waits for idle budget and yields to a pending visible request', async () => {
  const { engine, inputs, host } = lazyFixture(100);
  const overrides = { getInputs: () => inputs };
  const hook = renderHook(() => useRustDisplayList(
    inputs.layout as Layout, overrides, undefined, undefined, host
  ));
  try {
    await waitFor(() => expect(hook.result.current.frame).not.toBeNull());
    const worker = EngineWorker.last!;
    const builds = () => worker.posted.filter((entry) => entry.type === 'buildPages');
    await waitFor(() => expect(idleCallbacks.size).toBeGreaterThan(0));
    await act(async () => runIdleCallbacks(() => 0));
    expect(builds()).toEqual([]);
    const last = hook.result.current.displayList!.pages.length - 1;
    await act(async () => {
      hook.result.current.setDisplayWindow(last, last + 1);
      runIdleCallbacks();
    });
    await waitFor(() => expect(hook.result.current.displayList!.pages[last]!.unbuilt).toBeFalsy());
    expect(builds()).toEqual([expect.objectContaining({ pages: [last] })]);
    expect(hook.result.current.displayList!.pages[5]!.unbuilt).toBe(true);
    expect(hook.result.current.error).toBeNull();
  } finally {
    hook.unmount();
    engine.free();
  }
});

test('worker-open idle builds stop after the viewport margin', async () => {
  const { engine, inputs, host } = lazyFixture();
  try {
    const overrides = { getInputs: () => inputs };
    const { result, unmount } = renderHook(() =>
      useRustDisplayList(
        inputs.layout as Layout, overrides, undefined, undefined, host,
        undefined, undefined, undefined, true
      )
    );
    await waitFor(() => expect(result.current.frame).not.toBeNull());
    const pages = () => result.current.frame!.displayList.pages;
    expect(pages().length).toBeGreaterThan(7);
    expect(pages().slice(0, 5).every((page) => !page.unbuilt)).toBe(true);
    for (let round = 0; round < 10 && pages().slice(0, 7).some((page) => page.unbuilt); round += 1) {
      await waitFor(() => expect(idleCallbacks.size).toBeGreaterThan(0));
      await act(async () => runIdleCallbacks());
    }
    expect(pages().slice(0, 7).every((page) => !page.unbuilt)).toBe(true);
    await act(async () => {
      await new Promise((resolve) => setTimeout(resolve, 75));
      runIdleCallbacks();
    });
    expect(idleCallbacks.size).toBe(0);
    expect(pages().slice(7).every((page) => page.unbuilt)).toBe(true);
    expect(
      EngineWorker.last!.posted.filter((request) => request.type === 'buildPages')
    ).toEqual([expect.objectContaining({ pages: [5, 6] })]);
    unmount();
  } finally {
    engine.free();
  }
});

test('a visible request waits for a default-mode idle reply before building', async () => {
  const { engine, inputs, host } = lazyFixture(100);
  const overrides = { getInputs: () => inputs };
  const hook = renderHook(() => useRustDisplayList(
    inputs.layout as Layout, overrides, undefined, undefined, host
  ));
  try {
    await waitFor(() => expect(hook.result.current.frame).not.toBeNull());
    const worker = EngineWorker.last!;
    const builds = () => worker.posted.filter((entry) => entry.type === 'buildPages');
    worker.holdPageBuilds = true;
    await waitFor(() => expect(idleCallbacks.size).toBeGreaterThan(0));
    const before = hook.result.current.frame!;
    await act(async () => runIdleCallbacks());
    expect(builds()).toHaveLength(1);
    expect(builds()[0]!.pages.length).toBe(16);
    expect(hook.result.current.frame).toBe(before);
    const target = before.pages.find((page) =>
      page.page.unbuilt && !builds()[0]!.pages.includes(page.pageIndex)
    )!.pageIndex;
    await act(async () => {
      hook.result.current.setDisplayWindow(target, target + 1);
      await new Promise((resolve) => setTimeout(resolve, 20));
    });
    expect(builds()).toHaveLength(1);
    worker.holdPageBuilds = false;
    await act(async () => worker.releasePageBuilds());
    await waitFor(() => expect(hook.result.current.displayList!.pages[target]!.unbuilt).toBeFalsy());
    expect(builds()).toHaveLength(2);
    expect(builds()[1]!.pages).toEqual([target]);
    expect(builds()[1]!.expectedFrameEpoch).toBeGreaterThan(before.frameEpoch);
    const pages = hook.result.current.displayList!.pages;
    expect(builds()[0]!.pages.every((index) => !pages[index]!.unbuilt)).toBe(true);
    expect(hook.result.current.error).toBeNull();
  } finally {
    hook.unmount();
    engine.free();
  }
});

test('a settle wait promotes a worker-open background reply awaiting idle attachment', async () => {
  const { engine, inputs, host } = lazyFixture(100);
  const overrides = { getInputs: () => inputs };
  const hook = renderHook(() => useRustDisplayList(
    inputs.layout as Layout, overrides, undefined, undefined, host,
    undefined, undefined, undefined, true
  ));
  try {
    await waitFor(() => expect(hook.result.current.frame).not.toBeNull());
    const worker = EngineWorker.last!;
    const builds = () => worker.posted.filter((entry) => entry.type === 'buildPages');
    await waitFor(() => expect(idleCallbacks.size).toBeGreaterThan(0));
    const before = hook.result.current.frame!;
    await act(async () => runIdleCallbacks());
    expect(builds()).toHaveLength(1);
    await waitFor(() => expect(idleCallbacks.size).toBeGreaterThan(0));
    expect(hook.result.current.frame).toBe(before);
    let built = false;
    await act(async () => {
      const settled = await hook.result.current.settledDisplayList(null, null);
      built = settled.pages.every((page) => !page.unbuilt);
    });
    expect(built).toBe(true);
    expect(hook.result.current.error).toBeNull();
  } finally {
    hook.unmount();
    engine.free();
  }
});

test('worker-open window settling leaves far pages unbuilt while document settling builds them', async () => {
  const { engine, inputs, host } = lazyFixture();
  try {
    const overrides = { getInputs: () => inputs };
    const { result, unmount } = renderHook(() =>
      useRustDisplayList(
        inputs.layout as Layout, overrides, undefined, undefined, host,
        undefined, undefined, undefined, true
      )
    );
    await waitFor(() => expect(result.current.frame).not.toBeNull());
    const pageCount = result.current.frame!.displayList.pages.length;
    expect(pageCount).toBeGreaterThan(7);
    await act(async () => {
      const settled = await result.current.settledDisplayList(null, null, 'window');
      expect(settled.pages).toHaveLength(pageCount);
      expect(settled.pages.slice(0, 7).every((page) => !page.unbuilt)).toBe(true);
      expect(settled.pages.slice(7).every((page) => page.unbuilt)).toBe(true);
    });
    expect(
      EngineWorker.last!.posted.filter((request) => request.type === 'buildPages')
    ).toEqual([expect.objectContaining({ pages: [5, 6] })]);
    await act(async () => {
      const settled = await result.current.settledDisplayList(null, null);
      expect(settled.pages).toHaveLength(pageCount);
      expect(settled.pages.every((page) => !page.unbuilt)).toBe(true);
    });
    unmount();
  } finally {
    engine.free();
  }
});

test('worker-open background frames publish atomically from idle without replacing visible pages', async () => {
  const { engine, inputs, host } = lazyFixture(100);
  const overrides = { getInputs: () => inputs };
  const hook = renderHook(() => useRustDisplayList(
    inputs.layout as Layout, overrides, undefined, undefined, host,
    undefined, undefined, undefined, true
  ));
  try {
    await waitFor(() => expect(hook.result.current.frame).not.toBeNull());
    const worker = EngineWorker.last!;
    worker.slicePageBuilds = true;
    const before = hook.result.current.frame!;
    await waitFor(() => expect(idleCallbacks.size).toBeGreaterThan(0));
    await act(async () => runIdleCallbacks());
    expect(hook.result.current.frame).toBe(before);
    let checks = 0;
    await act(async () => runIdleCallbacks(() => checks++ === 0 ? 50 : 0));
    expect(hook.result.current.frame).toBe(before);
    expect(before.displayList.pages.slice(5).every((page) => page.unbuilt)).toBe(true);
    for (let round = 0; round < 50 && hook.result.current.frame === before; round += 1) {
      await act(async () => runIdleCallbacks());
    }
    const after = hook.result.current.frame!;
    expect(after.frameEpoch).toBe(before.frameEpoch + 1);
    expect(after.damagedPageIds.size).toBe(2);
    for (let index = 0; index < 5; index += 1) {
      expect(after.displayList.pages[index]).toBe(before.displayList.pages[index]);
    }
    expect(worker.posted.filter((entry) => entry.type === 'buildPages')).toHaveLength(1);
    expect(hook.result.current.error).toBeNull();
  } finally {
    hook.unmount();
    engine.free();
  }
});

test('a worker-open window wait settles when the viewport returns to built pages before a build starts', async () => {
  const { engine, inputs, host } = lazyFixture();
  try {
    const overrides = { getInputs: () => inputs };
    const { result, unmount } = renderHook(() =>
      useRustDisplayList(
        inputs.layout as Layout, overrides, undefined, undefined, host,
        undefined, undefined, undefined, true
      )
    );
    await waitFor(() => expect(result.current.frame).not.toBeNull());
    const pageCount = result.current.frame!.displayList.pages.length;
    expect(pageCount).toBeGreaterThan(10);
    await act(async () => {
      await result.current.settledDisplayList(null, null, 'window');
    });
    const builds = () =>
      EngineWorker.last!.posted.filter((request) => request.type === 'buildPages').length;
    const before = builds();
    let settled = false;
    await act(async () => {
      result.current.setDisplayWindow(pageCount - 2, pageCount);
      void result.current.settledDisplayList(null, null, 'window').then(() => {
        settled = true;
      });
      result.current.setDisplayWindow(0, 5);
      await new Promise((resolve) => setTimeout(resolve, 50));
    });
    expect(settled).toBe(true);
    expect(builds()).toBe(before);
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
      await settleWithIdle(result.current);
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
      const settled = await settleWithIdle(result.current);
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

test('a viewport build queued behind a page build keeps its timer when that build comes back stale', async () => {
  const { engine, inputs, host } = lazyFixture();
  const queued = new Map<number, () => void>();
  let nextTimer = -1;
  const schedule = globalThis.setTimeout;
  const unschedule = globalThis.clearTimeout;
  let timers: { mockRestore(): void } | undefined;
  let clears: { mockRestore(): void } | undefined;
  try {
    const overrides = { getInputs: () => inputs };
    const { result, unmount } = renderHook(() =>
      useRustDisplayList(inputs.layout as Layout, overrides, undefined, undefined, host)
    );
    await waitFor(() => expect(result.current.frame).not.toBeNull());
    const worker = EngineWorker.last!;
    const pages = () => result.current.frame!.displayList.pages;
    const pageBuilds = () => worker.posted.filter((request) => request.type === 'buildPages');
    const last = pages().length - 1;
    const middle = last - 1;
    expect(pages()[middle]!.unbuilt).toBe(true);

    worker.holdPageBuilds = true;
    await act(async () => {
      result.current.setDisplayWindow(last, last + 1);
    });
    await waitFor(() => expect(worker.heldPageBuilds).toHaveLength(1));
    const { paraId } = JSON.parse(engine.paragraphs('body'))[0] as { paraId: string };
    engine.set_selection('body', paraId, 1, paraId, 1);
    worker.holdInputReplies = true;
    let pendingEdit: ReturnType<typeof result.current.applyInput> | undefined;
    await act(async () => {
      pendingEdit = result.current.applyInput('New ');
    });
    await waitFor(() => expect(worker.heldInputReplies).toHaveLength(1));
    await act(async () => {
      result.current.setDisplayWindow(middle, middle + 1);
      await new Promise((resolve) => setTimeout(resolve, 20));
    });

    timers = spyOn(globalThis, 'setTimeout').mockImplementation(
      ((...input: Parameters<typeof setTimeout>) => {
        const [callback, delay, ...args] = input;
        if (delay === 50 && typeof callback === 'function') {
          const id = nextTimer--;
          queued.set(id, () => callback(...args));
          return id as unknown as ReturnType<typeof setTimeout>;
        }
        return schedule(callback, delay, ...args);
      }) as typeof setTimeout
    );
    clears = spyOn(globalThis, 'clearTimeout').mockImplementation(
      ((id?: ReturnType<typeof setTimeout>) => {
        if (!queued.delete(id as unknown as number)) unschedule(id);
      }) as typeof clearTimeout
    );
    worker.holdInputReplies = false;
    await act(async () => {
      worker.releaseInputReplies();
      expect(await pendingEdit!).not.toBeNull();
    });
    expect(pages()[middle]!.unbuilt).toBe(true);
    worker.holdPageBuilds = false;
    await act(async () => {
      worker.releasePageBuilds();
      await new Promise((resolve) => schedule(resolve, 0));
    });
    const before = pageBuilds().length;
    await act(async () => {
      const due = [...queued.values()];
      queued.clear();
      for (const run of due) run();
    });
    expect(pageBuilds().slice(before)).toEqual([expect.objectContaining({ pages: [middle] })]);
    timers.mockRestore();
    clears.mockRestore();
    await waitFor(() => expect(pages()[middle]!.unbuilt).toBeFalsy());
    expect(result.current.error).toBeNull();
    unmount();
  } finally {
    timers?.mockRestore();
    clears?.mockRestore();
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
        await act(async () => runIdleCallbacks());
        expect(pageBuilds().length).toBeGreaterThan(0);
      });
      expect(pageBuilds()[0]).toMatchObject({ expectedFrameEpoch: adopted });
      await waitFor(async () => {
        await act(async () => runIdleCallbacks());
        expect(result.current.frame!.displayList.pages[5]!.unbuilt).toBeFalsy();
      });
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

async function renderMainFallbackFixture(workerOpen = true, paragraphLength = 120) {
  const { engine, inputs, host } = lazyFixture(paragraphLength);
  const workerEngine = lazyFixture(paragraphLength).engine;
  Object.assign(inputs, { layoutRevision: 1 });
  stampSourceVersion(inputs.layout, 'main-v1');
  stampRevisionPreviewKey(inputs.layout, revisionPreviewKey(PREVIEW));
  const frameBuiltPages: number[] = [];
  const updateListeners = new Set<(update: Uint8Array) => void>();
  Object.assign(host, {
    onUpdate: (listener: (update: Uint8Array) => void) => {
      updateListeners.add(listener);
      return () => updateListeners.delete(listener);
    },
    residentCaretSnapshot: () => JSON.parse(engine.resident_caret_snapshot_json()),
    buildDisplayListJson: (input: string) => engine.build_display_list_json(input),
    buildDisplayListFrame: (input: string, epoch: number) => {
      const bytes = engine.build_display_list_frame(input, epoch);
      frameBuiltPages.push(decodeFrameDelta(bytes).operations.filter(
        (operation) => operation.kind === 'upsert' && !operation.page.unbuilt
      ).length);
      return bytes;
    },
    setDisplayWindow: (start: number, end: number) => engine.set_display_window(start, end),
    setDisplayRetainBuiltPages: (retain: boolean) => engine.set_display_retain_built_pages(retain),
    setWindowedIncrementalBuilds: (enabled: boolean) => engine.set_windowed_incremental_builds(enabled),
    buildDisplayPagesFrame: (pages: readonly number[], epoch: number) =>
      engine.build_display_pages_frame(Uint32Array.from(pages), epoch),
    releaseDisplayPagesFrame: (pages: number[], epoch: number) => {
      const bytes = engine.release_display_pages_frame(Uint32Array.from(pages), epoch);
      return bytes.length === 0 ? null : bytes;
    },
    resetFrameBase: () => engine.reset_frame_base(),
    displayHitTestRegionsJson: (page: number, x: number, y: number) =>
      engine.display_hit_test_regions_json(page, x, y),
    displayVerticalMoveJson: (position: number, direction: 'up' | 'down', goalX: number) =>
      engine.display_vertical_move_json(position, direction, goalX),
    displayRangeRectsJson: (from: number, to: number) => engine.display_range_rects_json(from, to),
    displayRangeRectsRegionJson: (region: string, partId: string, from: number, to: number) =>
      engine.display_range_rects_region_json(region, partId, from, to),
  });
  const frameBuilds = spyOn(host, 'buildDisplayListFrame');
  const pageBuilds = spyOn(host, 'buildDisplayPagesFrame');
  const releases = spyOn(host, 'releaseDisplayPagesFrame');
  const windows = spyOn(host, 'setDisplayWindow');
  const retention = spyOn(host, 'setDisplayRetainBuiltPages');
  const incremental = spyOn(host, 'setWindowedIncrementalBuilds');
  const hits = spyOn(host, 'displayHitTestRegionsJson');
  const errors = spyOn(console, 'error').mockImplementation(() => {});
  let relayouts = 0;
  const overrides = { getInputs: () => inputs };
  const hook = renderHook(
    ({ layout }) => useRustDisplayList(
      layout, overrides, undefined, undefined, host, () => { relayouts += 1; },
      undefined, undefined, workerOpen
    ),
    { initialProps: { layout: inputs.layout as Layout } }
  );
  const dispose = () => {
    hook.unmount();
    for (const spy of [frameBuilds, pageBuilds, releases, windows, retention, incremental, hits, errors]) {
      spy.mockRestore();
    }
    engine.free();
    workerEngine.free();
  };
  try {
    await waitFor(() => expect(hook.result.current.frame).not.toBeNull());
    const worker = EngineWorker.last!;
    EngineWorker.failPageBuilds = true;
    await act(async () => hook.result.current.setDisplayWindow(5, 6));
    await waitFor(() => expect(relayouts).toBe(1));
    expect(worker.terminated).toBe(true);
    expect(hook.result.current.frame).toBeNull();
    const mainLayout = { ...inputs.layout } as Layout;
    stampSourceVersion(mainLayout, 'main-v1');
    stampRevisionPreviewKey(mainLayout, revisionPreviewKey(PREVIEW));
    await act(async () => {
      hook.result.current.setDisplayWindow(0, 5);
      hook.rerender({ layout: mainLayout });
    });
    await waitFor(() => expect(hook.result.current.frame).not.toBeNull());
    expect(hook.result.current.workerSurfacesActive).toBe(false);
    expect(hook.result.current.shownFrameEngine()).toBe(host);
    expect(frameBuilds).toHaveBeenCalledTimes(1);
    return {
      ...hook, host, mainLayout, updateListeners, frameBuilds, frameBuiltPages,
      pageBuilds, releases, windows, retention, incremental, hits, errors, dispose,
    };
  } catch (error) {
    dispose();
    throw error;
  }
}

test('worker-open main fallback builds only the display window and idle margin', async () => {
  const fixture = await renderMainFallbackFixture();
  const { result, frameBuilds, frameBuiltPages, pageBuilds, windows, retention, incremental } = fixture;
  try {
    const pages = () => result.current.frame!.displayList.pages;
    expect(pages()).toHaveLength(40);
    expect(frameBuiltPages).toEqual([5]);
    expect(pages().slice(0, 5).every((page) => !page.unbuilt)).toBe(true);
    expect(pages().slice(7).every((page) => page.unbuilt)).toBe(true);
    expect(windows).toHaveBeenCalledWith(0, 5);
    expect(retention).toHaveBeenCalledWith(false);
    expect(incremental).toHaveBeenCalledWith(true);
    await idleUntil(() => pages().slice(0, 7).every((page) => !page.unbuilt));
    expect(pageBuilds.mock.calls.map(([pages]) => pages)).toEqual([[5, 6]]);
    await act(async () => {
      await new Promise((resolve) => setTimeout(resolve, 75));
      runIdleCallbacks();
    });
    expect(pages().slice(7).every((page) => page.unbuilt)).toBe(true);
    expect(frameBuilds).toHaveBeenCalledTimes(1);
    expect(frameBuiltPages.every((built) => built < pages().length)).toBe(true);
    expect(pageBuilds.mock.calls.every(([built]) => built.length < pages().length)).toBe(true);
    expect(result.current.caret).toBeNull();
    expect(result.current.error).toBeNull();
    expect(sourceVersionOf(result.current.queries)).toBe('main-v1');
    expect(revisionPreviewKeyOf(result.current.queries)).toBe(revisionPreviewKey(PREVIEW));
  } finally {
    fixture.dispose();
  }
});

test('worker-open main fallback builds distant pages, releases them at idle and rebuilds on return', async () => {
  const fixture = await renderMainFallbackFixture();
  const { result, frameBuilds, pageBuilds, releases, windows, host, hits } = fixture;
  try {
    const pages = () => result.current.frame!.displayList.pages;
    await idleUntil(() => pages().slice(0, 7).every((page) => !page.unbuilt));
    const identities = result.current.frame!.pages.map((page) => page.pageId);
    const firstPrimitives = pages()[0]!.primitives;
    await act(async () => result.current.setDisplayWindow(30, 34));
    expect(windows).toHaveBeenLastCalledWith(30, 34);
    await waitFor(() => expect(pages().slice(30, 34).every((page) => !page.unbuilt)).toBe(true));
    expect(pageBuilds.mock.calls.map(([built]) => built)).toContainEqual([30, 31, 32, 33]);
    expect(releases).not.toHaveBeenCalled();
    await idleUntil(() => pages().slice(0, 7).every((page) => page.unbuilt));
    expect(pages().slice(28, 36).every((page) => !page.unbuilt)).toBe(true);
    expect(pages().every((page, index) => (index >= 22 && index < 40) || page.unbuilt)).toBe(true);
    expect(releases.mock.calls[0]![0]).toEqual([0, 1, 2, 3, 4, 5, 6]);
    expect(result.current.frame!.pages.map((page) => page.pageId)).toEqual(identities);
    expect(result.current.caret).toBeNull();
    result.current.queries!.hitTestRegions(30, 10, 10);
    expect(hits).toHaveBeenLastCalledWith(30, 10, 10);
    expect(sourceVersionOf(result.current.queries)).toBe('main-v1');
    expect(revisionPreviewKeyOf(result.current.queries)).toBe(revisionPreviewKey(PREVIEW));

    await act(async () => result.current.setDisplayWindow(0, 5));
    await waitFor(() => expect(pages().slice(0, 5).every((page) => !page.unbuilt)).toBe(true));
    expect(pageBuilds.mock.calls.map(([built]) => built)).toContainEqual([0, 1, 2, 3, 4]);
    expect(pages()[0]!.primitives).toEqual(firstPrimitives);
    await idleUntil(() => pages().slice(28, 36).every((page) => page.unbuilt));
    expect(pages().slice(0, 7).every((page) => !page.unbuilt)).toBe(true);
    expect(result.current.shownFrameEngine()).toBe(host);
    expect(frameBuilds).toHaveBeenCalledTimes(1);
    expect(EngineWorker.spawned).toBe(1);
    expect(result.current.error).toBeNull();
  } finally {
    fixture.dispose();
  }
});

test('worker-open main page builds wait for the relayout after an edit', async () => {
  const fixture = await renderMainFallbackFixture();
  const { result, rerender, mainLayout, updateListeners, frameBuilds, pageBuilds, releases } = fixture;
  try {
    const pages = () => result.current.frame!.displayList.pages;
    await idleUntil(() => pages().slice(0, 7).every((page) => !page.unbuilt));
    const frame = result.current.frame;
    const queries = result.current.queries;
    pageBuilds.mockClear();
    await act(async () => {
      for (const listener of updateListeners) listener(new Uint8Array());
      result.current.setDisplayWindow(30, 34);
      await new Promise((resolve) => setTimeout(resolve, 75));
      runIdleCallbacks();
    });
    expect(pageBuilds).not.toHaveBeenCalled();
    expect(releases).not.toHaveBeenCalled();
    expect(result.current.frame).toBe(frame);
    expect(result.current.queries).toBe(queries);
    expect(sourceVersionOf(result.current.queries)).toBe('main-v1');
    expect(await result.current.resolveQueries()).toBeNull();
    expect(pages().slice(30, 34).every((page) => page.unbuilt)).toBe(true);

    const layout = { ...mainLayout };
    stampSourceVersion(layout, 'main-v2');
    stampRevisionPreviewKey(layout, revisionPreviewKey(PREVIEW));
    await act(async () => rerender({ layout }));
    await waitFor(() => expect(pages().slice(30, 34).every((page) => !page.unbuilt)).toBe(true));
    await idleUntil(() => pages().slice(28, 36).every((page) => !page.unbuilt));
    expect(frameBuilds).toHaveBeenCalledTimes(2);
    expect(pageBuilds.mock.calls.map(([built]) => built)).toContainEqual([28, 29, 34, 35]);
    expect(sourceVersionOf(result.current.queries)).toBe('main-v2');
    expect(result.current.error).toBeNull();
  } finally {
    fixture.dispose();
  }
});

test('worker-open main release keeps the caret page', async () => {
  const fixture = await renderMainFallbackFixture();
  const { result, host, releases } = fixture;
  const caret = spyOn(host, 'residentCaretSnapshot').mockImplementation(() => ({
    frameEpoch: result.current.frame!.frameEpoch,
    caretRect: {
      pageIndex: 2,
      pageId: result.current.frame!.pages[2]!.pageId.toString(),
      x: 10,
      y: 10,
      height: 12,
    },
  }));
  try {
    const pages = () => result.current.frame!.displayList.pages;
    const released = [0, 1, 3, 4, 5, 6];
    await idleUntil(() => pages().slice(0, 7).every((page) => !page.unbuilt));
    await act(async () => result.current.setDisplayWindow(30, 34));
    await waitFor(() => expect(pages().slice(30, 34).every((page) => !page.unbuilt)).toBe(true));
    await idleUntil(() => released.every((index) => pages()[index]!.unbuilt));
    expect(pages()[2]!.unbuilt).toBeFalsy();
    expect(releases.mock.calls[0]![0]).toEqual(released);
    expect(releases.mock.calls.every(([indices]) => !indices.includes(2))).toBe(true);
    expect(result.current.caret).toBeNull();
    expect(result.current.error).toBeNull();
  } finally {
    caret.mockRestore();
    fixture.dispose();
  }
});

test('main fallback without worker-open still builds every page', async () => {
  const fixture = await renderMainFallbackFixture(false);
  const { result, frameBuiltPages, pageBuilds, releases, windows, retention, incremental } = fixture;
  try {
    expect(result.current.displayList!.pages).toHaveLength(40);
    expect(frameBuiltPages).toEqual([40]);
    expect(result.current.displayList!.pages.every((page) => !page.unbuilt)).toBe(true);
    await act(async () => {
      result.current.setDisplayWindow(30, 34);
      runIdleCallbacks();
    });
    expect(pageBuilds).not.toHaveBeenCalled();
    expect(releases).not.toHaveBeenCalled();
    expect(windows).not.toHaveBeenCalled();
    expect(retention).not.toHaveBeenCalled();
    expect(incremental).not.toHaveBeenCalled();
    expect(result.current.error).toBeNull();
  } finally {
    fixture.dispose();
  }
});

test('worker-open main fallback settles the document through the shared page batches', async () => {
  const fixture = await renderMainFallbackFixture(true, 24);
  const { result, frameBuilds, pageBuilds, releases } = fixture;
  try {
    const pages = () => result.current.frame!.displayList.pages;
    await idleUntil(() => pages().slice(0, 7).every((page) => !page.unbuilt));
    const window = await result.current.settledDisplayList(null, null, 'window');
    expect(window.pages.slice(7).every((page) => page.unbuilt)).toBe(true);
    pageBuilds.mockClear();
    let settled: Awaited<ReturnType<typeof result.current.settledDisplayList>> | undefined;
    await act(async () => {
      settled = await result.current.settledDisplayList(null, null, 'document');
      expect(releases).not.toHaveBeenCalled();
    });
    expect(settled!.pages).toHaveLength(200);
    expect(settled!.pages.every((page) => !page.unbuilt)).toBe(true);
    expect(pageBuilds.mock.calls[0]![0]).toHaveLength(128);
    expect(pageBuilds.mock.calls.every(([built]) => built.length <= 128)).toBe(true);
    expect(frameBuilds).toHaveBeenCalledTimes(1);
    await idleUntil(() => pages().slice(13).every((page) => page.unbuilt));
    expect(releases).toHaveBeenCalledTimes(1);
    expect(result.current.error).toBeNull();
  } finally {
    fixture.dispose();
  }
});

test.each(['build', 'release'] as const)('a failed main page %s reports the main engine error', async (kind) => {
  const fixture = await renderMainFallbackFixture();
  const { result, pageBuilds, releases, host, errors } = fixture;
  try {
    const failure = new Error(`main page ${kind} failed`);
    await idleUntil(() => result.current.displayList!.pages.slice(0, 7).every((page) => !page.unbuilt));
    const target = kind === 'build' ? pageBuilds : releases;
    target.mockImplementation(() => { throw failure; });
    errors.mockClear();
    await act(async () => result.current.setDisplayWindow(30, 34));
    await waitFor(async () => {
      await act(async () => runIdleCallbacks());
      expect(result.current.error).toBe(failure);
    });
    expect(errors).toHaveBeenCalledWith('[CanvasRenderer] Building display pages failed', failure);
    expect(result.current.errorEngine).toBe(host);
    expect(result.current.loading).toBe(false);
    expect(result.current.frame).not.toBeNull();
    expect(result.current.shownFrameEngine()).toBe(host);
    expect(EngineWorker.spawned).toBe(1);
    await expect(result.current.settledDisplayList(null)).rejects.toBe(failure);
  } finally {
    fixture.dispose();
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

function renderReleaseFixture(workerOpen = true, caretAtStart = false) {
  const fixture = lazyFixture(160);
  if (caretAtStart) {
    const { paraId } = JSON.parse(fixture.engine.paragraphs('body'))[0] as { paraId: string };
    fixture.engine.set_selection('body', paraId, 1, paraId, 1);
  }
  const overrides = { getInputs: () => fixture.inputs };
  const hook = renderHook(() =>
    useRustDisplayList(
      fixture.inputs.layout as Layout,
      overrides,
      undefined,
      undefined,
      fixture.host,
      undefined,
      undefined,
      undefined,
      workerOpen
    )
  );
  return { ...fixture, ...hook };
}

async function idleUntil(done: () => boolean): Promise<void> {
  for (let round = 0; round < 50 && !done(); round += 1) {
    await waitFor(() => expect(idleCallbacks.size).toBeGreaterThan(0));
    await act(async () => runIdleCallbacks());
  }
  expect(done()).toBe(true);
}

test('worker-open releases distant display pages and rebuilds them when scrolling back', async () => {
  const { engine, result, unmount } = renderReleaseFixture();
  try {
    await waitFor(() => expect(result.current.frame).not.toBeNull());
    const worker = EngineWorker.last!;
    const pages = () => result.current.frame!.displayList.pages;
    expect(pages().length).toBeGreaterThanOrEqual(30);
    await act(async () => result.current.setDisplayWindow(0, 2));
    await idleUntil(() => pages().slice(0, 4).every((page) => !page.unbuilt));
    const identities = result.current.frame!.pages.map((page) => page.pageId);
    const firstPrimitives = pages()[0]!.primitives;
    await act(async () => result.current.setDisplayWindow(20, 22));
    await waitFor(() => expect(pages().slice(20, 22).every((page) => !page.unbuilt)).toBe(true));
    await idleUntil(() => worker.releasedIndices.length > 0);
    expect(pages().slice(18, 24).every((page) => !page.unbuilt)).toBe(true);
    const caretPage = result.current.caret?.caretRect?.pageIndex;
    expect(pages().every((page, index) =>
      (index >= 12 && index < 30) || index === caretPage || page.unbuilt
    )).toBe(true);
    expect(pages().filter((page) => !page.unbuilt).length).toBeLessThanOrEqual(2 + 2 * 2 + 2 * 8 + 1);
    expect(result.current.frame!.pages.map((page) => page.pageId)).toEqual(identities);
    const release = worker.posted.find((request) => request.type === 'releasePages')!;
    expect(release).toMatchObject({ paintCaret: false });
    expect(revisionPreviewKeyOf(result.current.queries)).toBe(revisionPreviewKey(PREVIEW));
    await act(async () => result.current.setDisplayWindow(0, 2));
    await waitFor(() => expect(pages().slice(0, 2).every((page) => !page.unbuilt)).toBe(true));
    expect(pages()[0]!.primitives).toEqual(firstPrimitives);
    await idleUntil(() => pages().slice(0, 4).every((page) => !page.unbuilt));
    await idleUntil(() => pages().slice(18, 24).every((page, index) => index + 18 === caretPage || page.unbuilt));
    await act(async () => {
      await new Promise((resolve) => setTimeout(resolve, 75));
      runIdleCallbacks();
    });
    expect(idleCallbacks.size).toBe(0);
    unmount();
  } finally {
    engine.free();
  }
});

test('worker-open retains the caret page outside the retention band', async () => {
  const { engine, result, unmount } = renderReleaseFixture(true, true);
  try {
    await waitFor(() => expect(result.current.caret?.caretRect?.pageIndex).toBe(0));
    const worker = EngineWorker.last!;
    await act(async () => result.current.setDisplayWindow(20, 22));
    await idleUntil(() => worker.releasedIndices.length > 0);
    expect(result.current.frame!.displayList.pages[0]!.unbuilt).toBeFalsy();
    expect(result.current.frame!.displayList.pages[1]!.unbuilt).toBe(true);
    expect(worker.releasedIndices).not.toContain(0);
    expect(result.current.caret?.caretRect?.pageIndex).toBe(0);
    unmount();
  } finally {
    engine.free();
  }
});

test('worker-open retains built pages until retention is disabled', async () => {
  const { engine, result, unmount } = renderReleaseFixture();
  try {
    await waitFor(() => expect(result.current.frame).not.toBeNull());
    const worker = EngineWorker.last!;
    await act(async () => {
      result.current.setRetainBuiltPages!(true);
      await result.current.settledDisplayList(null, null, 'document');
      result.current.setDisplayWindow(20, 22);
    });
    await act(async () => {
      await new Promise((resolve) => setTimeout(resolve, 75));
      runIdleCallbacks();
    });
    expect(worker.releaseRequests).toBe(0);
    expect(result.current.frame!.displayList.pages.every((page) => !page.unbuilt)).toBe(true);
    await act(async () => result.current.setRetainBuiltPages!(false));
    expect(worker.releaseRequests).toBe(0);
    await idleUntil(() => worker.releasedIndices.length > 0);
    expect(result.current.frame!.displayList.pages[1]!.unbuilt).toBe(true);
    expect(result.current.frame!.displayList.pages.slice(12, 30).every((page) => !page.unbuilt)).toBe(true);
    expect(worker.releaseRequests).toBe(1);
    unmount();
  } finally {
    engine.free();
  }
});

test('worker-open document settling suppresses release until every page is built', async () => {
  const { engine, result, unmount } = renderReleaseFixture();
  try {
    await waitFor(() => expect(result.current.frame).not.toBeNull());
    const worker = EngineWorker.last!;
    worker.holdPageBuilds = true;
    let pending: ReturnType<typeof result.current.settledDisplayList> | undefined;
    await act(async () => {
      pending = result.current.settledDisplayList(null, null, 'document');
    });
    await waitFor(() => expect(worker.heldPageBuilds).toHaveLength(1));
    await act(async () => runIdleCallbacks());
    expect(worker.releaseRequests).toBe(0);
    worker.holdPageBuilds = false;
    await act(async () => {
      worker.releasePageBuilds();
      const settled = await pending!;
      expect(settled.pages.every((page) => !page.unbuilt)).toBe(true);
      expect(worker.releaseRequests).toBe(0);
    });
    await idleUntil(() => worker.releasedIndices.length > 0);
    expect(result.current.frame!.displayList.pages[20]!.unbuilt).toBe(true);
    unmount();
  } finally {
    engine.free();
  }
});

test('worker-open retries superseded releases without dropping the worker', async () => {
  const { engine, result, unmount } = renderReleaseFixture();
  try {
    await waitFor(() => expect(result.current.frame).not.toBeNull());
    const worker = EngineWorker.last!;
    worker.supersedeNextRelease = true;
    await act(async () => result.current.setDisplayWindow(20, 22));
    await idleUntil(() => worker.releaseRequests === 1);
    expect(worker.releasedIndices).toEqual([]);
    expect(result.current.frame!.displayList.pages[1]!.unbuilt).toBeFalsy();
    expect(worker.terminated).toBe(false);
    expect(result.current.workerSurfacesActive).toBe(true);
    expect(result.current.error).toBeNull();
    await idleUntil(() => worker.releasedIndices.length > 0);
    expect(worker.releaseRequests).toBe(2);
    expect(worker.terminated).toBe(false);
    expect(EngineWorker.spawned).toBe(1);
    expect(result.current.frame!.displayList.pages[1]!.unbuilt).toBe(true);
    expect(result.current.workerSurfacesActive).toBe(true);
    expect(result.current.error).toBeNull();
    unmount();
  } finally {
    engine.free();
  }
});

test('default worker mode builds every page and never requests releases', async () => {
  const { engine, result, unmount } = renderReleaseFixture(false);
  try {
    await waitFor(() => expect(result.current.frame).not.toBeNull());
    const worker = EngineWorker.last!;
    await act(async () => {
      result.current.setDisplayWindow(20, 22);
      result.current.setRetainBuiltPages!(false);
    });
    await idleUntil(() => result.current.frame!.displayList.pages.every((page) => !page.unbuilt));
    expect(worker.releaseRequests).toBe(0);
    expect(worker.posted.some((request) => request.type === 'releasePages')).toBe(false);
    unmount();
  } finally {
    engine.free();
  }
});

test('worker-open delays release while the caret is in the typing window', async () => {
  const { engine, result, unmount } = renderReleaseFixture();
  let now = performance.now();
  const clock = spyOn(performance, 'now').mockImplementation(() => now);
  try {
    await waitFor(() => expect(result.current.frame).not.toBeNull());
    const worker = EngineWorker.last!;
    await act(async () => result.current.setDisplayWindow(20, 22));
    await idleUntil(() => result.current.frame!.displayList.pages.slice(18, 24).every((page) => !page.unbuilt));
    await act(async () => result.current.notifyCaretInput());
    await waitFor(() => expect(idleCallbacks.size).toBeGreaterThan(0));
    await act(async () => {
      await new Promise((resolve) => setTimeout(resolve, 75));
      runIdleCallbacks();
    });
    expect(worker.releaseRequests).toBe(0);
    now += 501;
    await waitFor(async () => {
      await act(async () => runIdleCallbacks());
      expect(worker.releasedIndices.length).toBeGreaterThan(0);
    });
    expect(result.current.frame!.displayList.pages[1]!.unbuilt).toBe(true);
    unmount();
  } finally {
    clock.mockRestore();
    engine.free();
  }
});
