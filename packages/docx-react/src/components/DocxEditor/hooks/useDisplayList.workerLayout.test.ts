import { GlobalRegistrator } from '@happy-dom/global-registrator';
import { afterAll, afterEach, beforeAll, expect, spyOn, test } from 'bun:test';
import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';
import type { Layout } from '@betteroffice/docx/layout/pagination';
import { createEditSession, preloadEditWasm } from '@betteroffice/docx/wasm/edit';
import type { YrsSession } from '@betteroffice/docx/yrs';
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
  static last: FakeWorker | null = null;
  onmessage: ((event: MessageEvent<ResidentEngineWorkerResponse>) => void) | null = null;
  onerror: ((event: ErrorEvent) => void) | null = null;
  onmessageerror = null;
  posted: ResidentEngineWorkerRequest[] = [];
  constructor() {
    FakeWorker.last = this;
  }
  postMessage(request: ResidentEngineWorkerRequest): void {
    this.posted.push(request);
  }
  reply(response: ResidentEngineWorkerResponse): void {
    this.onmessage?.({ data: response } as MessageEvent<ResidentEngineWorkerResponse>);
  }
  terminate(): void {}
}

function setup() {
  const native = createEditSession(9301);
  native.create_story('body', 'Owned layout', 'Normal', 'left');
  const layoutJson = native.layout_document_with_regions_retained_json(REQUEST);
  const frame = native.build_display_list_frame(JSON.stringify({}), 0);
  const adopted: string[] = [];
  const engine = {
    adoptResidentWorkerLayout: (input: string) => {
      adopted.push(input);
      return adopted.length;
    },
    residentLayoutInWorker: () => true,
    resetFrameBase: () => {},
    residentWorkerProbe: () => ({ layoutRevision: adopted.length }),
    residentWorkerSnapshot: () => ({
      state: new Uint8Array(),
      fonts: [],
      fontsRevision: 0,
      layoutRevision: adopted.length,
    }),
    encodeStateVector: () => new Uint8Array([1]),
    onUpdate: () => () => {},
    selection: () => null,
    applyUpdate: () => null,
  } as unknown as YrsSession;
  globalThis.Worker = FakeWorker as unknown as typeof Worker;
  return { native, layoutJson, frame, engine, adopted };
}

test('a worker-run layout arrives with its frame and needs no second worker pass', async () => {
  const { native, layoutJson, frame, engine, adopted } = setup();
  try {
    const { result, rerender, unmount } = renderHook(
      ({ layout, source }) => useRustDisplayList(layout, undefined, undefined, undefined, source),
      {
        initialProps: {
          layout: null as Layout | null,
          source: null as YrsSession | null,
        },
      }
    );
    const pending = result.current.layoutInWorker(engine, REQUEST);
    expect(pending).not.toBeNull();
    expect(adopted).toEqual([REQUEST]);
    const worker = FakeWorker.last!;
    expect(worker.posted).toHaveLength(1);
    expect(worker.posted[0]).toMatchObject({ type: 'bootstrap', extras: '' });
    expect(worker.posted[0]).toHaveProperty('layoutExtras', '{}');
    worker.reply({
      id: worker.posted[0].id,
      ok: true,
      frame: frame.slice().buffer,
      caret: { frameEpoch: 1, caretRect: null },
      selection: null,
      layoutRevision: 1,
      layoutJson,
    });
    const computation = await pending!;
    expect(computation?.layout.pages.length).toBeGreaterThan(0);
    await act(async () => {
      rerender({ layout: computation!.layout, source: engine });
    });
    await waitFor(() => {
      if (result.current.error) throw result.current.error;
      expect(result.current.frame?.frameEpoch).toBe(1);
    });
    expect(result.current.workerSurfacesActive).toBe(true);
    expect(worker.posted).toHaveLength(1);
    unmount();
  } finally {
    native.free();
  }
});

test('a failed worker layout hands the pass back to the main thread', async () => {
  const { native, engine } = setup();
  const errors = spyOn(console, 'error').mockImplementation(() => {});
  try {
    const { result, unmount } = renderHook(() =>
      useRustDisplayList(null, undefined, undefined, undefined, null)
    );
    let outcome: unknown;
    await act(async () => {
      const pending = result.current.layoutInWorker(engine, REQUEST);
      FakeWorker.last!.onerror?.({ message: 'worker crashed' } as ErrorEvent);
      outcome = await pending!;
    });
    expect(outcome).toBeNull();
    expect(result.current.layoutInWorker(engine, REQUEST)).toBeNull();
    unmount();
  } finally {
    errors.mockRestore();
    native.free();
  }
});

test('a frame built for other display extras is not adopted', async () => {
  const { native, layoutJson, frame, engine } = setup();
  try {
    const { result, rerender, unmount } = renderHook(
      ({ layout, source, resolved }) =>
        useRustDisplayList(layout, undefined, undefined, resolved, source),
      {
        initialProps: {
          layout: null as Layout | null,
          source: null as YrsSession | null,
          resolved: undefined as ReadonlySet<number> | undefined,
        },
      }
    );
    const pending = result.current.layoutInWorker(engine, REQUEST);
    const worker = FakeWorker.last!;
    worker.reply({
      id: worker.posted[0].id,
      ok: true,
      frame: frame.slice().buffer,
      caret: { frameEpoch: 1, caretRect: null },
      selection: null,
      layoutRevision: 1,
      layoutJson,
    });
    const computation = await pending!;
    await act(async () => {
      rerender({ layout: computation!.layout, source: engine, resolved: new Set([7]) });
    });
    expect(worker.posted.at(-1)).toMatchObject({ type: 'buildFrame' });
    unmount();
  } finally {
    native.free();
  }
});

test('a reply without a layout hands the pass back to the main thread', async () => {
  const { native, frame, engine } = setup();
  const errors = spyOn(console, 'error').mockImplementation(() => {});
  try {
    const { result, unmount } = renderHook(() =>
      useRustDisplayList(null, undefined, undefined, undefined, null)
    );
    let outcome: unknown;
    await act(async () => {
      const pending = result.current.layoutInWorker(engine, REQUEST);
      const worker = FakeWorker.last!;
      worker.reply({
        id: worker.posted[0].id,
        ok: true,
        frame: frame.slice().buffer,
        caret: { frameEpoch: 1, caretRect: null },
        selection: null,
        layoutRevision: 1,
      });
      outcome = await pending!;
    });
    expect(outcome).toBeNull();
    expect(result.current.layoutInWorker(engine, REQUEST)).toBeNull();
    unmount();
  } finally {
    errors.mockRestore();
    native.free();
  }
});

test('a provisional layout paints first and settles only once the full layout follows', async () => {
  const { native, layoutJson, frame, engine } = setup();
  try {
    const fullLayoutJson = native.layout_document_with_regions_retained_json(REQUEST);
    const fullFrame = native.build_display_list_frame(JSON.stringify({}), 1);
    const { result, rerender, unmount } = renderHook(
      ({ layout, source }) => useRustDisplayList(layout, undefined, undefined, undefined, source),
      { initialProps: { layout: null as Layout | null, source: null as YrsSession | null } }
    );
    const pending = result.current.layoutInWorker(engine, REQUEST);
    const worker = FakeWorker.last!;
    expect(worker.posted[0]).toMatchObject({ type: 'bootstrap', provisionalPages: 3 });
    worker.reply({
      id: worker.posted[0].id,
      ok: true,
      frame: frame.slice().buffer,
      caret: { frameEpoch: 1, caretRect: null },
      selection: null,
      layoutRevision: 1,
      layoutJson,
      layoutProvisional: true,
    });
    const provisional = await pending!;
    await act(async () => {
      rerender({ layout: provisional!.layout, source: engine });
    });
    await waitFor(() => expect(result.current.frame?.frameEpoch).toBe(1));
    expect(result.current.loading).toBe(false);
    // The rest waits until the first surfaces are attached, so the worker
    // paints them first.
    expect(worker.posted).toHaveLength(1);
    await act(async () => {
      void result.current.attachOffscreenCanvases([], [], 1, 1, { color: '#000', width: 2 });
    });
    worker.reply({ id: worker.posted[1].id, ok: true });
    await waitFor(() => expect(worker.posted).toHaveLength(3));
    expect(worker.posted.map((request) => request.type)).toEqual([
      'bootstrap',
      'attachCanvases',
      'completeLayout',
    ]);
    expect(worker.posted[2]).toMatchObject({ expectedFrameEpoch: 1 });
    let settled = false;
    void result.current.settledDisplayList(() => {}).then(() => {
      settled = true;
    });
    await act(async () => {});
    expect(settled).toBe(false);

    worker.reply({
      id: worker.posted[2].id,
      ok: true,
      frame: fullFrame.slice().buffer,
      caret: { frameEpoch: 2, caretRect: null },
      selection: null,
      layoutRevision: 1,
      layoutJson: fullLayoutJson,
    });
    const complete = await provisional!.complete!;
    await act(async () => {
      rerender({ layout: complete!.layout, source: engine });
    });
    await waitFor(() => expect(result.current.frame?.frameEpoch).toBe(2));
    await waitFor(() => expect(settled).toBe(true));
    expect(worker.posted).toHaveLength(3);
    unmount();
  } finally {
    native.free();
  }
});
