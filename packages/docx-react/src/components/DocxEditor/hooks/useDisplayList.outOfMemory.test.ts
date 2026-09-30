import { GlobalRegistrator } from '@happy-dom/global-registrator';
import { afterAll, afterEach, beforeAll, expect, spyOn, test } from 'bun:test';
import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';
import type { Layout } from '@betteroffice/docx/layout/pagination';
import { createEditSession, preloadEditWasm } from '@betteroffice/docx/wasm/edit';
import { ResidentWorkerOutOfMemoryError, type YrsSession } from '@betteroffice/docx/yrs';
import type {
  ResidentEngineWorkerRequest,
  ResidentEngineWorkerResponse,
} from '@betteroffice/docx/yrs/residentEngineWorkerProtocol';
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

afterEach(() => {
  cleanup();
  globalThis.Worker = originalWorker;
});

afterAll(async () => {
  if (ownsDom) await GlobalRegistrator.unregister();
});

const REQUEST = JSON.stringify({
  bodyStory: 'body',
  regions: { sections: [{ sectionId: 'main', properties: {} }] },
  measurement: { defaults: { fontSize: 11, fontFamily: 'Calibri' } },
  renderEnv: {},
});

class FakeWorker {
  static spawned: FakeWorker[] = [];
  onmessage: ((event: MessageEvent<ResidentEngineWorkerResponse>) => void) | null = null;
  onerror: ((event: ErrorEvent) => void) | null = null;
  onmessageerror = null;
  posted: ResidentEngineWorkerRequest[] = [];
  terminated = false;
  constructor() {
    FakeWorker.spawned.push(this);
  }
  postMessage(request: ResidentEngineWorkerRequest): void {
    this.posted.push(request);
  }
  last(): ResidentEngineWorkerRequest {
    return this.posted[this.posted.length - 1]!;
  }
  replyFrame(frame: Uint8Array, frameEpoch: number, extra: object = {}): void {
    this.onmessage?.({
      data: {
        id: this.last().id,
        ok: true,
        frame: frame.slice().buffer,
        caret: { frameEpoch, caretRect: null },
        selection: null,
        layoutRevision: 1,
        ...extra,
      },
    } as MessageEvent<ResidentEngineWorkerResponse>);
  }
  outOfMemory(): void {
    this.onmessage?.({
      data: {
        id: this.last().id,
        ok: false,
        error: 'Resident engine worker ran out of memory allocating 65536 bytes: unreachable',
        terminal: true,
        outOfMemory: true,
        memory: [{ label: 'docx-edit', bufferBytes: 65536, liveBytes: 4000, peakBytes: 4000, failedAllocationBytes: 65536 }],
      },
    } as MessageEvent<ResidentEngineWorkerResponse>);
  }
  trapped(): void {
    this.onmessage?.({
      data: {
        id: this.last().id,
        ok: false,
        error: 'Resident engine worker trapped: unreachable',
        terminal: true,
      },
    } as MessageEvent<ResidentEngineWorkerResponse>);
  }
  terminate(): void {
    this.terminated = true;
  }
}

function setup() {
  FakeWorker.spawned = [];
  globalThis.Worker = FakeWorker as unknown as typeof Worker;
  const native = createEditSession(9401);
  native.create_story('body', 'Out of memory', 'Normal', 'left');
  const inputs = JSON.parse(native.layout_document_with_regions_json(REQUEST));
  const layoutJson = native.layout_document_with_regions_retained_json(REQUEST);
  const frame = (epoch: number) => {
    const bytes = native.build_display_list_frame(JSON.stringify(inputs), 0);
    new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength).setBigUint64(32, BigInt(epoch), true);
    return bytes;
  };
  const mainThreadBuilds: number[] = [];
  const engine = {
    buildDisplayListJson: (input: string) => native.build_display_list_json(input),
    resetFrameBase: () => native.reset_frame_base(),
    buildDisplayListFrame: (input: string, epoch: number) => {
      mainThreadBuilds.push(epoch);
      return native.build_display_list_frame(input, epoch);
    },
    adoptResidentWorkerLayout: () => 1,
    residentWorkerProbe: () => ({ layoutRevision: 1 }),
    residentWorkerSnapshot: () => ({ state: new Uint8Array(), fonts: [], fontsRevision: 0 }),
    encodeStateVector: () => new Uint8Array(),
    onUpdate: () => () => {},
    selection: () => null,
    applyUpdate: () => null,
  } as unknown as YrsSession;
  return { native, inputs, frame, engine, mainThreadBuilds, layoutJson };
}

test('a worker that runs out of memory is replaced once, never by the main thread', async () => {
  const { native, inputs, frame, engine, mainThreadBuilds } = setup();
  const warnings = spyOn(console, 'warn').mockImplementation(() => {});
  const errors = spyOn(console, 'error').mockImplementation(() => {});
  try {
    const overrides = { getInputs: () => inputs };
    const { result, rerender, unmount } = renderHook(
      ({ layout }) =>
        useRustDisplayList(layout, overrides, undefined, undefined, engine, undefined, 1 << 30),
      { initialProps: { layout: inputs.layout as Layout } }
    );
    const [first] = FakeWorker.spawned;
    expect(first!.last()).toMatchObject({ type: 'bootstrap', heapLimitBytes: 1 << 30 });
    await act(async () => first!.replyFrame(frame(100), 100));
    await waitFor(() => expect(result.current.frame?.frameEpoch).toBe(100));

    await act(async () => rerender({ layout: { ...inputs.layout } }));
    expect(first!.last()).toMatchObject({ type: 'buildFrame' });
    await act(async () => first!.outOfMemory());
    expect(first!.terminated).toBe(true);
    const second = FakeWorker.spawned[1];
    expect(FakeWorker.spawned).toHaveLength(2);
    expect(second!.last()).toMatchObject({
      type: 'bootstrap',
      heapLimitBytes: 1 << 30,
      expectedFrameEpoch: 100,
    });
    await act(async () => second!.replyFrame(frame(101), 101));
    await waitFor(() => expect(result.current.frame?.frameEpoch).toBe(101));
    expect(result.current.error).toBeNull();
    expect(result.current.workerSurfacesActive).toBe(true);

    await act(async () => rerender({ layout: { ...inputs.layout } }));
    await act(async () => second!.outOfMemory());
    await waitFor(() => expect(result.current.error).toBeInstanceOf(ResidentWorkerOutOfMemoryError));
    expect(result.current.workerSurfacesActive).toBe(false);
    expect(result.current.frame?.frameEpoch).toBe(101);

    await act(async () => rerender({ layout: { ...inputs.layout } }));
    await waitFor(() => expect(result.current.error).toBeInstanceOf(ResidentWorkerOutOfMemoryError));
    expect(FakeWorker.spawned).toHaveLength(2);
    expect(mainThreadBuilds).toEqual([]);
    unmount();
  } finally {
    warnings.mockRestore();
    errors.mockRestore();
    native.free();
  }
});

test('a worker layout that runs out of memory runs again in a fresh worker, then rejects', async () => {
  const { native, frame, engine, mainThreadBuilds, layoutJson } = setup();
  const warnings = spyOn(console, 'warn').mockImplementation(() => {});
  const errors = spyOn(console, 'error').mockImplementation(() => {});
  try {
    const { result, unmount } = renderHook(() =>
      useRustDisplayList(null, undefined, undefined, undefined, null)
    );
    let retried: Promise<unknown> | null = null;
    await act(async () => {
      retried = result.current.layoutInWorker(engine, REQUEST);
      FakeWorker.spawned[0]!.outOfMemory();
    });
    expect(FakeWorker.spawned).toHaveLength(2);
    expect(FakeWorker.spawned[1]!.last()).toMatchObject({ type: 'bootstrap' });
    await act(async () => FakeWorker.spawned[1]!.replyFrame(frame(1), 1, { layoutJson }));
    const computation = (await retried) as { layout: Layout } | null;
    expect(computation?.layout.pages.length).toBeGreaterThan(0);

    let refused: unknown;
    await act(async () => {
      const pending = result.current.layoutInWorker(engine, REQUEST)!;
      FakeWorker.spawned[1]!.outOfMemory();
      refused = await pending.catch((error: unknown) => error);
    });
    expect(refused).toBeInstanceOf(ResidentWorkerOutOfMemoryError);
    expect(result.current.error).toBe(refused as Error);
    expect(result.current.workerSurfacesActive).toBe(false);
    await expect(result.current.layoutInWorker(engine, REQUEST)!).rejects.toBe(refused);
    expect(FakeWorker.spawned).toHaveLength(2);
    expect(mainThreadBuilds).toEqual([]);
    expect(warnings).toHaveBeenCalledTimes(1);
    unmount();
  } finally {
    warnings.mockRestore();
    errors.mockRestore();
    native.free();
  }
});

test('a provisional layout whose completion runs out of memory completes in a fresh worker', async () => {
  const { native, frame, engine, layoutJson } = setup();
  const warnings = spyOn(console, 'warn').mockImplementation(() => {});
  try {
    const { result, unmount } = renderHook(() =>
      useRustDisplayList(null, undefined, undefined, undefined, null)
    );
    const pending = result.current.layoutInWorker(engine, REQUEST)!;
    const [first] = FakeWorker.spawned;
    await act(async () =>
      first!.replyFrame(frame(1), 1, { layoutJson, layoutProvisional: true })
    );
    const provisional = (await pending) as { complete?: Promise<{ layout: Layout } | null> };
    await waitFor(() => expect(first!.last()).toMatchObject({ type: 'completeLayout' }));
    await act(async () => first!.outOfMemory());
    expect(FakeWorker.spawned).toHaveLength(2);
    const second = FakeWorker.spawned[1]!;
    expect(second.last()).toMatchObject({ type: 'bootstrap', provisionalPages: 3 });
    await act(async () => second.replyFrame(frame(2), 2, { layoutJson }));
    const complete = await provisional.complete!;
    expect(complete?.layout.pages.length).toBeGreaterThan(0);
    expect(result.current.error).toBeNull();
    unmount();
  } finally {
    warnings.mockRestore();
    native.free();
  }
});

/** A host engine that counts its layouts, as a session's layout revision does. */
function revisedHost(engine: YrsSession) {
  const adopted: string[] = [];
  let revision = 0;
  const host = {
    ...engine,
    adoptResidentWorkerLayout: (request: string) => {
      adopted.push(request);
      return ++revision;
    },
    residentWorkerProbe: () => ({ layoutRevision: revision }),
  } as YrsSession;
  return { host, adopted, layOutHere: () => void ++revision };
}

test('a superseded provisional layout does not run again in the replacement worker', async () => {
  const { native, frame, engine, layoutJson } = setup();
  const { host, adopted } = revisedHost(engine);
  const newer = JSON.stringify({ ...JSON.parse(REQUEST), renderEnv: { preview: 'b' } });
  const warnings = spyOn(console, 'warn').mockImplementation(() => {});
  try {
    const { result, unmount } = renderHook(() =>
      useRustDisplayList(null, undefined, undefined, undefined, null)
    );
    const pending = result.current.layoutInWorker(host, REQUEST)!;
    const [first] = FakeWorker.spawned;
    await act(async () =>
      first!.replyFrame(frame(1), 1, { layoutJson, layoutProvisional: true })
    );
    const provisional = (await pending) as { complete?: Promise<unknown> };
    await act(async () => {
      const recovered = result.current.layoutInWorker(host, newer)!;
      first!.outOfMemory();
      await waitFor(() => expect(FakeWorker.spawned).toHaveLength(2));
      FakeWorker.spawned[1]!.replyFrame(frame(2), 2, { layoutJson });
      await recovered;
    });
    const second = FakeWorker.spawned[1]!;
    const sent = second.posted.length;
    await act(async () => expect(await provisional.complete).toBeNull());
    expect(second.posted).toHaveLength(sent);
    expect(second.terminated).toBe(false);
    expect(adopted).toEqual([REQUEST, newer, newer]);
    expect(FakeWorker.spawned).toHaveLength(2);
    expect(result.current.error).toBeNull();
    unmount();
  } finally {
    warnings.mockRestore();
    native.free();
  }
});

test('a provisional layout that a host layout replaced does not run again once its worker runs out of memory', async () => {
  const { native, frame, engine, layoutJson } = setup();
  const { host, adopted, layOutHere } = revisedHost(engine);
  const warnings = spyOn(console, 'warn').mockImplementation(() => {});
  try {
    const { result, unmount } = renderHook(() =>
      useRustDisplayList(null, undefined, undefined, undefined, null)
    );
    const pending = result.current.layoutInWorker(host, REQUEST)!;
    const [first] = FakeWorker.spawned;
    await act(async () =>
      first!.replyFrame(frame(1), 1, { layoutJson, layoutProvisional: true })
    );
    const provisional = (await pending) as { complete?: Promise<unknown> };
    layOutHere();
    await waitFor(() => expect(first!.last()).toMatchObject({ type: 'completeLayout' }));
    await act(async () => first!.outOfMemory());
    await act(async () => expect(await provisional.complete).toBeNull());
    expect(FakeWorker.spawned).toHaveLength(1);
    expect(adopted).toEqual([REQUEST]);
    unmount();
  } finally {
    warnings.mockRestore();
    native.free();
  }
});

test('a display build that fails beside a worker pass leaves the pass its replacement worker', async () => {
  const { native, frame, engine, layoutJson } = setup();
  const { host, adopted } = revisedHost(engine);
  const warnings = spyOn(console, 'warn').mockImplementation(() => {});
  try {
    const { result, rerender, unmount } = renderHook(
      ({ layout, resolved }) =>
        useRustDisplayList(layout, undefined, undefined, resolved, host),
      {
        initialProps: {
          layout: null as Layout | null,
          resolved: undefined as ReadonlySet<number> | undefined,
        },
      }
    );
    const pending = result.current.layoutInWorker(host, REQUEST)!;
    const [first] = FakeWorker.spawned;
    await act(async () => first!.replyFrame(frame(1), 1, { layoutJson }));
    const computation = (await pending) as { layout: Layout };
    await act(async () => rerender({ layout: computation.layout, resolved: undefined }));
    await waitFor(() => expect(result.current.frame?.frameEpoch).toBe(1));

    // A relayout and a display build are both waiting when the worker runs out of memory.
    await act(async () => {
      void result.current.layoutInWorker(host, REQUEST);
      rerender({ layout: computation.layout, resolved: new Set([7]) });
    });
    expect(first!.posted.map((request) => request.type)).toEqual(['bootstrap', 'sync', 'sync']);
    await act(async () => first!.outOfMemory());
    expect(FakeWorker.spawned).toHaveLength(2);
    const second = FakeWorker.spawned[1]!;
    expect(second.posted.map((request) => request.type)).toEqual(['bootstrap']);
    expect(second.last()).toMatchObject({ provisionalPages: 3 });
    expect(adopted).toHaveLength(3);
    expect(result.current.error).toBeNull();
    await act(async () => second.replyFrame(frame(2), 2, { layoutJson }));
    unmount();
  } finally {
    warnings.mockRestore();
    native.free();
  }
});

test('a worker that runs out of memory attaching canvases is replaced and lays out again', async () => {
  const { native, inputs, frame, engine } = setup();
  let relayouts = 0;
  const warnings = spyOn(console, 'warn').mockImplementation(() => {});
  try {
    const overrides = { getInputs: () => inputs };
    const { result, rerender, unmount } = renderHook(
      ({ layout }) =>
        useRustDisplayList(layout, overrides, undefined, undefined, engine, () => {
          relayouts += 1;
        }),
      { initialProps: { layout: inputs.layout as Layout } }
    );
    const [first] = FakeWorker.spawned;
    await act(async () => first!.replyFrame(frame(100), 100));
    await waitFor(() => expect(result.current.workerSurfacesActive).toBe(true));

    let attached: unknown;
    await act(async () => {
      const pending = result.current.attachOffscreenCanvases([], [], 1, 1, {
        color: '#000',
        width: 2,
      });
      expect(first!.last()).toMatchObject({ type: 'attachCanvases' });
      first!.outOfMemory();
      attached = await pending;
    });
    expect(attached).toBe(false);
    expect(first!.terminated).toBe(true);
    expect(relayouts).toBe(1);
    expect(result.current.workerSurfacesActive).toBe(false);
    expect(result.current.error).toBeNull();

    await act(async () => rerender({ layout: { ...inputs.layout } }));
    expect(FakeWorker.spawned).toHaveLength(2);
    const second = FakeWorker.spawned[1]!;
    expect(second.last()).toMatchObject({ type: 'bootstrap' });
    await act(async () => second.replyFrame(frame(101), 101));
    await waitFor(() => expect(result.current.workerSurfacesActive).toBe(true));
    unmount();
  } finally {
    warnings.mockRestore();
    native.free();
  }
});

test("a loading document's worker that runs out of memory twice fails the wait for its layout", async () => {
  const { native, engine } = setup();
  const warnings = spyOn(console, 'warn').mockImplementation(() => {});
  const errors = spyOn(console, 'error').mockImplementation(() => {});
  try {
    // The renderer learns the session from its first layout.
    const { result, unmount } = renderHook(() =>
      useRustDisplayList(null, undefined, undefined, undefined, null)
    );
    act(() => result.current.resetSettled());
    let waited: unknown = 'pending';
    void result.current.settledDisplayList(null, null).then(
      () => (waited = 'settled'),
      (error: unknown) => (waited = error)
    );
    await act(async () => {
      const pending = result.current.layoutInWorker(engine, REQUEST)!;
      FakeWorker.spawned[0]!.outOfMemory();
      await waitFor(() => expect(FakeWorker.spawned).toHaveLength(2));
      FakeWorker.spawned[1]!.outOfMemory();
      await pending.catch(() => {});
    });
    expect(waited).toBeInstanceOf(ResidentWorkerOutOfMemoryError);
    unmount();
  } finally {
    warnings.mockRestore();
    errors.mockRestore();
    native.free();
  }
});

test("a loading document's worker recovers while the previous document is still shown", async () => {
  const { native, engine, frame, layoutJson } = setup();
  const shown = { ...engine } as YrsSession;
  const warnings = spyOn(console, 'warn').mockImplementation(() => {});
  try {
    const { result, unmount } = renderHook(() =>
      useRustDisplayList(null, undefined, undefined, undefined, shown)
    );
    act(() => result.current.resetSettled());
    await act(async () => {
      void result.current.layoutInWorker(engine, REQUEST);
      FakeWorker.spawned[0]!.outOfMemory();
    });
    expect(FakeWorker.spawned).toHaveLength(2);
    expect(FakeWorker.spawned[1]!.last()).toMatchObject({ type: 'bootstrap' });
    await act(async () => FakeWorker.spawned[1]!.replyFrame(frame(1), 1, { layoutJson }));
    unmount();
  } finally {
    warnings.mockRestore();
    native.free();
  }
});

test('a replaced document whose worker runs out of memory again fails nothing of the next one', async () => {
  const { native, engine } = setup();
  const warnings = spyOn(console, 'warn').mockImplementation(() => {});
  try {
    const { result, unmount } = renderHook(() =>
      useRustDisplayList(null, undefined, undefined, undefined, null)
    );
    await act(async () => {
      void result.current.layoutInWorker(engine, REQUEST);
      FakeWorker.spawned[0]!.outOfMemory();
    });
    expect(FakeWorker.spawned).toHaveLength(2);
    act(() => result.current.resetSettled());
    let waited: unknown = 'pending';
    void result.current.settledDisplayList(null, null).then(
      () => (waited = 'settled'),
      (error: unknown) => (waited = error)
    );
    await act(async () => FakeWorker.spawned[1]!.outOfMemory());
    expect(waited).toBe('pending');
    expect(result.current.error).toBeNull();
    expect(FakeWorker.spawned).toHaveLength(2);
    unmount();
  } finally {
    warnings.mockRestore();
    native.free();
  }
});

test("a request sent to a replaced document's worker after another load began fails nothing of it", async () => {
  const { native, engine } = setup();
  const warnings = spyOn(console, 'warn').mockImplementation(() => {});
  try {
    const { result, unmount } = renderHook(() =>
      useRustDisplayList(null, undefined, undefined, undefined, null)
    );
    await act(async () => {
      void result.current.layoutInWorker(engine, REQUEST);
      FakeWorker.spawned[0]!.outOfMemory();
    });
    const second = FakeWorker.spawned[1]!;
    act(() => result.current.resetSettled());
    let waited: unknown = 'pending';
    void result.current.settledDisplayList(null, null).then(
      () => (waited = 'settled'),
      (error: unknown) => (waited = error)
    );
    let attached: unknown;
    await act(async () => {
      const pending = result.current.attachOffscreenCanvases([], [], 1, 1, {
        color: '#000',
        width: 2,
      });
      second.outOfMemory();
      attached = await pending;
    });
    expect(attached).toBe(false);
    expect(waited).toBe('pending');
    expect(result.current.error).toBeNull();
    unmount();
  } finally {
    warnings.mockRestore();
    native.free();
  }
});

test("a replaced document's worker started again for its display fails nothing of the next load", async () => {
  const { native, inputs, frame, engine } = setup();
  const warnings = spyOn(console, 'warn').mockImplementation(() => {});
  try {
    const overrides = { getInputs: () => inputs };
    const { result, rerender, unmount } = renderHook(
      ({ resolved }) =>
        useRustDisplayList(inputs.layout as Layout, overrides, undefined, resolved, engine),
      { initialProps: { resolved: undefined as ReadonlySet<number> | undefined } }
    );
    const [first] = FakeWorker.spawned;
    await act(async () => first!.replyFrame(frame(100), 100));
    await waitFor(() => expect(result.current.workerSurfacesActive).toBe(true));
    await act(async () => {
      const attached = result.current.attachOffscreenCanvases([], [], 1, 1, {
        color: '#000',
        width: 2,
      });
      first!.outOfMemory();
      await attached;
    });
    // Another document starts loading while this one is still shown and rebuilt.
    act(() => result.current.resetSettled());
    let waited: unknown = 'pending';
    void result.current.settledDisplayList(null, null).then(
      () => (waited = 'settled'),
      (error: unknown) => (waited = error)
    );
    await act(async () => rerender({ resolved: new Set([7]) }));
    expect(FakeWorker.spawned).toHaveLength(2);
    await act(async () => FakeWorker.spawned[1]!.outOfMemory());
    expect(waited).toBe('pending');
    expect(result.current.error).toBeNull();
    unmount();
  } finally {
    warnings.mockRestore();
    native.free();
  }
});

test("a replaced session's first worker, started after the next load began, fails nothing of it", async () => {
  const { native, engine } = setup();
  const warnings = spyOn(console, 'warn').mockImplementation(() => {});
  try {
    const { result, unmount } = renderHook(() =>
      useRustDisplayList(null, undefined, undefined, undefined, null)
    );
    // The editor records each session as it starts, before its first layout.
    act(() => result.current.recordSession(engine));
    act(() => result.current.resetSettled());
    let waited: unknown = 'pending';
    void result.current.settledDisplayList(null, null).then(
      () => (waited = 'settled'),
      (error: unknown) => (waited = error)
    );
    await act(async () => {
      void result.current.layoutInWorker(engine, REQUEST);
      FakeWorker.spawned[0]!.outOfMemory();
    });
    expect(FakeWorker.spawned).toHaveLength(1);
    expect(waited).toBe('pending');
    expect(result.current.error).toBeNull();
    unmount();
  } finally {
    warnings.mockRestore();
    native.free();
  }
});

test("a replaced document's standing out-of-memory failure fails nothing of the next load", async () => {
  const { native, inputs, frame, engine } = setup();
  const warnings = spyOn(console, 'warn').mockImplementation(() => {});
  const errors = spyOn(console, 'error').mockImplementation(() => {});
  try {
    const overrides = { getInputs: () => inputs };
    const { result, rerender, unmount } = renderHook(
      ({ layout }) => useRustDisplayList(layout, overrides, undefined, undefined, engine),
      { initialProps: { layout: inputs.layout as Layout } }
    );
    await act(async () => FakeWorker.spawned[0]!.replyFrame(frame(100), 100));
    await act(async () => rerender({ layout: { ...inputs.layout } }));
    await act(async () => FakeWorker.spawned[0]!.outOfMemory());
    await act(async () => FakeWorker.spawned[1]!.outOfMemory());
    await waitFor(() => expect(result.current.error).toBeInstanceOf(ResidentWorkerOutOfMemoryError));

    act(() => result.current.resetSettled());
    let waited: unknown = 'pending';
    void result.current.settledDisplayList(null, null).then(
      () => (waited = 'settled'),
      (error: unknown) => (waited = error)
    );
    // The replaced document lays out once more before the next one arrives.
    await act(async () => rerender({ layout: { ...inputs.layout } }));
    await act(async () => {});
    expect(waited).toBe('pending');
    expect(FakeWorker.spawned).toHaveLength(2);
    unmount();
  } finally {
    warnings.mockRestore();
    errors.mockRestore();
    native.free();
  }
});

test("an out-of-memory failure from a replaced document's worker leaves the new worker alone", async () => {
  const { native, inputs, frame, engine } = setup();
  const other = { ...engine } as YrsSession;
  const warnings = spyOn(console, 'warn').mockImplementation(() => {});
  try {
    const overrides = { getInputs: () => inputs };
    const { result, rerender, unmount } = renderHook(
      ({ layout, source }) => useRustDisplayList(layout, overrides, undefined, undefined, source),
      { initialProps: { layout: inputs.layout as Layout, source: engine } }
    );
    const [first] = FakeWorker.spawned;
    await act(async () => first!.replyFrame(frame(100), 100));
    await act(async () => rerender({ layout: { ...inputs.layout }, source: engine }));
    expect(first!.last()).toMatchObject({ type: 'buildFrame' });

    // The failure is delivered before the switch, and handled after it.
    act(() => {
      first!.outOfMemory();
      rerender({ layout: { ...inputs.layout }, source: other });
    });
    await act(async () => {});
    expect(FakeWorker.spawned).toHaveLength(2);
    const second = FakeWorker.spawned[1]!;
    expect(second.terminated).toBe(false);
    expect(second.last()).toMatchObject({ type: 'bootstrap' });
    expect(FakeWorker.spawned).toHaveLength(2);
    await act(async () => second.replyFrame(frame(1), 1));
    await waitFor(() => expect(result.current.frame?.frameEpoch).toBe(1));
    expect(result.current.error).toBeNull();
    unmount();
  } finally {
    warnings.mockRestore();
    native.free();
  }
});

test('any other worker failure still hands the display to the main thread', async () => {
  const { native, inputs, frame, engine, mainThreadBuilds } = setup();
  const errors = spyOn(console, 'error').mockImplementation(() => {});
  try {
    const overrides = { getInputs: () => inputs };
    const { result, rerender, unmount } = renderHook(
      ({ layout }) => useRustDisplayList(layout, overrides, undefined, undefined, engine),
      { initialProps: { layout: inputs.layout as Layout } }
    );
    const [worker] = FakeWorker.spawned;
    await act(async () => worker!.replyFrame(frame(100), 100));
    await waitFor(() => expect(result.current.frame?.frameEpoch).toBe(100));

    await act(async () => rerender({ layout: { ...inputs.layout } }));
    await act(async () => worker!.trapped());
    await waitFor(() => expect(mainThreadBuilds).toEqual([100]));
    await waitFor(() => expect(result.current.frame?.frameEpoch).toBe(101));
    expect(result.current.error).toBeNull();
    expect(result.current.workerSurfacesActive).toBe(false);
    expect(FakeWorker.spawned).toHaveLength(1);
    unmount();
  } finally {
    errors.mockRestore();
    native.free();
  }
});

test('typing into a worker that runs out of memory keeps the keystroke for the host and replaces the worker', async () => {
  const { native, inputs, frame, engine, mainThreadBuilds } = setup();
  const at = { story: 'body', paraId: 'p', offset: 0 };
  const typing = { ...engine, selection: () => ({ anchor: at, head: at }) } as YrsSession;
  const warnings = spyOn(console, 'warn').mockImplementation(() => {});
  try {
    const overrides = { getInputs: () => inputs };
    const { result, rerender, unmount } = renderHook(
      ({ layout }) => useRustDisplayList(layout, overrides, undefined, undefined, typing),
      { initialProps: { layout: inputs.layout as Layout } }
    );
    const [first] = FakeWorker.spawned;
    await act(async () => first!.replyFrame(frame(100), 100));
    await waitFor(() => expect(result.current.frame?.frameEpoch).toBe(100));

    let typed: unknown;
    await act(async () => {
      const pending = result.current.applyInput('x');
      await waitFor(() => expect(first!.last()).toMatchObject({ type: 'applyInput' }));
      first!.outOfMemory();
      typed = await pending;
    });
    expect(typed).toBeNull();
    expect(first!.terminated).toBe(true);
    expect(result.current.error).toBeNull();

    await act(async () => rerender({ layout: { ...inputs.layout } }));
    expect(FakeWorker.spawned).toHaveLength(2);
    expect(FakeWorker.spawned[1]!.last()).toMatchObject({ type: 'bootstrap' });
    expect(mainThreadBuilds).toEqual([]);
    unmount();
  } finally {
    warnings.mockRestore();
    native.free();
  }
});

test('a failure that arrives after unmount starts no worker', async () => {
  const { native, frame, engine, layoutJson } = setup();
  const warnings = spyOn(console, 'warn').mockImplementation(() => {});
  try {
    const { result, unmount } = renderHook(() =>
      useRustDisplayList(null, undefined, undefined, undefined, null)
    );
    const provisional = result.current.layoutInWorker(engine, REQUEST)!;
    const [first] = FakeWorker.spawned;
    await act(async () =>
      first!.replyFrame(frame(1), 1, { layoutJson, layoutProvisional: true })
    );
    await provisional;
    // An overlapping pass runs the first worker out of memory and is retried.
    await act(async () => {
      const retried = result.current.layoutInWorker(engine, REQUEST)!;
      first!.outOfMemory();
      await waitFor(() => expect(FakeWorker.spawned).toHaveLength(2));
      FakeWorker.spawned[1]!.replyFrame(frame(2), 2, { layoutJson });
      await retried;
    });
    unmount();
    // The first pass's completion is still due, and fails with the old error.
    await new Promise((resolve) => setTimeout(resolve, 400));
    expect(FakeWorker.spawned).toHaveLength(2);
  } finally {
    warnings.mockRestore();
    native.free();
  }
});
