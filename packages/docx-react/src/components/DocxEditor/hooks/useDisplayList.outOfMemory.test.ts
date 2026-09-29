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
  return { native, inputs, frame, engine, mainThreadBuilds };
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

test('a worker layout that runs out of memory is retried in a fresh worker, then refused', async () => {
  const { native, engine, mainThreadBuilds } = setup();
  const warnings = spyOn(console, 'warn').mockImplementation(() => {});
  const errors = spyOn(console, 'error').mockImplementation(() => {});
  try {
    const { result, unmount } = renderHook(() =>
      useRustDisplayList(null, undefined, undefined, undefined, null)
    );
    let outcome: unknown;
    await act(async () => {
      const pending = result.current.layoutInWorker(engine, REQUEST);
      FakeWorker.spawned[0]!.outOfMemory();
      outcome = await pending;
    });
    expect(outcome).toBeNull();
    await act(async () => {
      const pending = result.current.layoutInWorker(engine, REQUEST);
      expect(pending).not.toBeNull();
      expect(FakeWorker.spawned).toHaveLength(2);
      FakeWorker.spawned[1]!.outOfMemory();
      outcome = await pending;
    });
    expect(outcome).toBeNull();
    expect(result.current.error).toBeInstanceOf(ResidentWorkerOutOfMemoryError);
    expect(result.current.layoutInWorker(engine, REQUEST)).toBeNull();
    expect(FakeWorker.spawned).toHaveLength(2);
    expect(mainThreadBuilds).toEqual([]);
    unmount();
  } finally {
    warnings.mockRestore();
    errors.mockRestore();
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
