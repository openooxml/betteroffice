import { GlobalRegistrator } from '@happy-dom/global-registrator';
import { afterAll, afterEach, beforeAll, beforeEach, expect, spyOn, test } from 'bun:test';
import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';
import type { Layout } from '@betteroffice/docx/layout/pagination';
import { createEditSession, preloadEditWasm } from '@betteroffice/docx/wasm/edit';
import { ResidentWorkerOutOfMemoryError, type YrsSession } from '@betteroffice/docx/yrs';
import type {
  ResidentEngineWorkerRequest,
  ResidentEngineWorkerResponse,
} from '@betteroffice/docx/yrs/residentEngineWorkerProtocol';
import { revisionPreviewKey, revisionPreviewKeyOf } from '../internals/layoutProvenance';
import { useRustDisplayList } from './useDisplayList';

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

/** A worker that answers from a real engine session. */
class EngineWorker {
  static engine: ReturnType<typeof createEditSession> | null = null;
  static last: EngineWorker | null = null;
  onmessage: ((event: MessageEvent<ResidentEngineWorkerResponse>) => void) | null = null;
  onerror: ((event: ErrorEvent) => void) | null = null;
  onmessageerror = null;
  posted: ResidentEngineWorkerRequest[] = [];
  constructor() {
    EngineWorker.last = this;
    EngineWorker.spawned += 1;
  }
  static failPageBuilds = false;
  static outOfMemoryPageBuilds = false;
  static spawned = 0;
  terminated = false;
  postMessage(request: ResidentEngineWorkerRequest): void {
    this.posted.push(request);
    const engine = EngineWorker.engine!;
    if (request.type === 'buildPages' && EngineWorker.outOfMemoryPageBuilds) {
      queueMicrotask(() =>
        this.onmessage?.({
          data: {
            id: request.id,
            ok: false,
            error: 'Resident engine worker ran out of memory allocating 64 bytes: unreachable',
            terminal: true,
            outOfMemory: true,
          },
        } as MessageEvent<ResidentEngineWorkerResponse>)
      );
      return;
    }
    if (request.type === 'buildPages' && EngineWorker.failPageBuilds) {
      queueMicrotask(() =>
        this.onmessage?.({
          data: { id: request.id, ok: false, error: 'page build failed' },
        } as MessageEvent<ResidentEngineWorkerResponse>)
      );
      return;
    }
    let frame: Uint8Array;
    if (request.type === 'bootstrap') {
      if (request.displayWindow) engine.set_display_window(...request.displayWindow);
      frame = engine.build_display_list_frame(request.extras, 0);
    } else if (request.type === 'buildPages') {
      frame = engine.build_display_pages_frame(
        Uint32Array.from(request.pages),
        request.expectedFrameEpoch
      );
    } else {
      return;
    }
    const caret = JSON.parse(engine.resident_caret_snapshot_json());
    queueMicrotask(() =>
      this.onmessage?.({
        data: {
          id: request.id,
          ok: true,
          frame: frame.slice().buffer,
          caret,
          selection: null,
          layoutRevision: 1,
        },
      } as MessageEvent<ResidentEngineWorkerResponse>)
    );
  }
  terminate(): void {
    this.terminated = true;
  }
}

const PREVIEW = { r1: 'accepted' } as const;

function lazyFixture() {
  const engine = createEditSession(9401);
  engine.create_story('body', 'Lazy pages. '.repeat(400), 'Normal', 'left');
  const fontId = engine.register_measure_font(
    new Uint8Array(
      readFileSync(
        resolve(
          import.meta.dir,
          '../../../../../../crates/ooxml-text/tests/fonts/LiberationSans-Regular.ttf'
        )
      )
    )
  );
  const inputs = JSON.parse(
    engine.layout_document_with_regions_json(
      JSON.stringify({
        bodyStory: 'body',
        regions: {
          sections: [
            {
              sectionId: 'main',
              properties: {
                pageWidth: 4320,
                pageHeight: 2880,
                marginTop: 300,
                marginRight: 300,
                marginBottom: 300,
                marginLeft: 300,
              },
            },
          ],
        },
        measurement: {
          fontChains: { 'calibri|0|0': [fontId] },
          defaults: { fontSize: 11, fontFamily: 'Calibri' },
          authoritativeShaping: true,
        },
        renderEnv: {},
      })
    )
  );
  EngineWorker.engine = engine;
  EngineWorker.failPageBuilds = false;
  EngineWorker.outOfMemoryPageBuilds = false;
  EngineWorker.spawned = 0;
  globalThis.Worker = EngineWorker as unknown as typeof Worker;
  const host = {
    residentWorkerProbe: () => ({ layoutRevision: 1 }),
    residentWorkerSnapshot: () => ({
      state: new Uint8Array(),
      fonts: [],
      fontsRevision: 0,
      layoutRevision: 1,
      layoutInput: JSON.stringify({ renderEnv: { revisionPreview: PREVIEW } }),
    }),
    resetFrameBase: () => {},
    encodeStateVector: () => new Uint8Array(),
    onUpdate: () => () => {},
    selection: () => null,
    applyUpdate: () => null,
  } as unknown as YrsSession;
  return { engine, inputs, host };
}

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
