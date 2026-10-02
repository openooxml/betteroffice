import { GlobalRegistrator } from '@happy-dom/global-registrator';
import { afterAll, afterEach, beforeAll, beforeEach, expect, mock, spyOn, test } from 'bun:test';
import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';
import type { Layout } from '@betteroffice/docx/layout/pagination';
import { LayoutSelectionGate, type ResidentMeasurementConfig } from '@betteroffice/docx/layout';
import { decodeFrameDelta, loadRustDisplayListQueryEngine } from '@betteroffice/docx/layout/render';
import { createEditSession, preloadEditWasm } from '@betteroffice/docx/wasm/edit';
import * as wasm from '@betteroffice/docx/yrs/wasm/index';
import {
  preloadDocxEngine,
  proposalSetIdentity,
  ResidentEngineWorkerClient,
  takePreloadedResidentEngineWorker,
  type ResidentEngineWorkerFrame,
  type ResidentProposalReply,
  type YrsRenderEnv,
  type YrsSelection,
  type YrsSession,
} from '@betteroffice/docx/yrs';
import type {
  ResidentEngineWorkerHostModule,
  ResidentEngineWorkerRequest,
  ResidentEngineWorkerResponse,
} from '@betteroffice/docx/yrs/residentEngineWorkerProtocol';
import { markSupersededLayout } from '../internals/layoutProvenance';
import { registerWorkerProposalAuthority } from '../internals/workerProposalAuthority';
import { useRustDisplayList, type ResidentFrameApplyResult } from './useDisplayList';
import { useLayoutPipeline, type UseLayoutPipelineOptions } from './useLayoutPipeline';

const ownsDom = !GlobalRegistrator.isRegistered;
if (ownsDom) GlobalRegistrator.register();
const { act, cleanup, renderHook, waitFor } = await import('@testing-library/react');
const originalWorker = globalThis.Worker;
const editModule = new WebAssembly.Module(
  new Uint8Array([0x00, 0x61, 0x73, 0x6d, 0x01, 0x00, 0x00, 0x00])
);
let compileModule: ReturnType<typeof spyOn<typeof wasm, 'editWasmModule'>>;

beforeEach(() => {
  compileModule = spyOn(wasm, 'editWasmModule').mockResolvedValue(editModule);
});

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
  takePreloadedResidentEngineWorker()?.destroy();
  compileModule.mockRestore();
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
  static instances: FakeWorker[] = [];
  onmessage: ((event: MessageEvent<ResidentEngineWorkerResponse>) => void) | null = null;
  onerror: ((event: ErrorEvent) => void) | null = null;
  onmessageerror = null;
  posted: (ResidentEngineWorkerRequest | ResidentEngineWorkerHostModule)[] = [];
  terminated = false;
  constructor() {
    FakeWorker.last = this;
    FakeWorker.instances.push(this);
  }
  postMessage(request: ResidentEngineWorkerRequest | ResidentEngineWorkerHostModule): void {
    this.posted.push(request);
  }
  requestAt(index: number): ResidentEngineWorkerRequest {
    const message = this.posted.at(index)!;
    if (!('id' in message)) throw new Error('Expected a worker request');
    return message;
  }
  reply(response: ResidentEngineWorkerResponse): void {
    this.onmessage?.({ data: response } as MessageEvent<ResidentEngineWorkerResponse>);
  }
  terminate(): void {
    this.terminated = true;
  }
}

function setup(clientId = 9301, text = 'Owned layout', request = REQUEST) {
  const native = createEditSession(clientId);
  const { paraId } = JSON.parse(native.create_story('body', text, 'Normal', 'left')) as {
    paraId: string;
  };
  const layoutJson = native.layout_document_with_regions_retained_json(request);
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
  return { native, paraId, layoutJson, frame, engine, adopted };
}

function setupLayoutPipeline() {
  const source = setup();
  Object.assign(source.engine, {
    version: () => '1',
    layoutFontRequirementsJson: () => '[]',
  });
  return source;
}

function useWorkerLayoutPipeline(
  session: YrsSession | null,
  residentMeasurementConfig: UseLayoutPipelineOptions['residentMeasurementConfig'],
  overrides?: Parameters<typeof useRustDisplayList>[1]
) {
  const display = useRustDisplayList(null, overrides);
  return useLayoutPipeline({
    document: null,
    session,
    renderEnv: {} as YrsRenderEnv,
    pageGap: 24,
    zoom: 1,
    residentMeasurementConfig,
    deferLayoutPass: () => false,
    pagesContainerRef: { current: null },
    viewportLayoutRef: { current: null },
    syncCoordinator: new LayoutSelectionGate(),
    getScrollContainer: () => null,
    layoutInWorker: display.layoutInWorker,
  });
}

test('unresolved fonts warm once and the first layout adopts the still-warming worker', async () => {
  const source = setupLayoutPipeline();
  const initialWorkers = FakeWorker.instances.length;
  let settleFonts: () => void = () => {};
  let measurement: ResidentMeasurementConfig | null = null;
  const fonts = new Promise<void>((resolve) => (settleFonts = resolve)).then(() => {
    measurement = {} as ResidentMeasurementConfig;
  });
  const preflight = mock(() => measurement);
  const hook = renderHook(() => useWorkerLayoutPipeline(source.engine, preflight));
  try {
    act(() => {
      hook.result.current.runLayoutPipeline();
      hook.result.current.runLayoutPipeline();
    });
    const spare = FakeWorker.last!;
    expect(preflight).toHaveBeenCalledTimes(2);
    expect(measurement).toBeNull();
    expect(source.adopted).toEqual([]);
    expect(spare.posted.map((request) => request.type)).toEqual(['warm']);
    expect(FakeWorker.instances.length - initialWorkers).toBe(1);
    await act(async () => {
      settleFonts();
      await fonts;
      hook.result.current.runLayoutPipeline();
    });
    expect(FakeWorker.last).toBe(spare);
    expect(FakeWorker.instances.length - initialWorkers).toBe(1);
    expect(spare.posted.map((request) => request.type)).toEqual(['warm', 'editModule', 'bootstrap']);
    expect(spare.posted[0]).toEqual({ id: 1, type: 'warm', hostModule: true });
    expect(spare.posted[1]).toEqual({ type: 'editModule', module: editModule });
    await act(async () => {
      spare.reply({ id: spare.requestAt(0).id, ok: true });
      spare.reply({
        id: spare.requestAt(2).id,
        ok: true,
        frame: source.frame.slice().buffer,
        caret: { frameEpoch: 1, caretRect: null },
        selection: null,
        layoutRevision: 1,
        layoutJson: source.layoutJson,
      });
    });
    expect(hook.result.current.layout?.pages.length).toBeGreaterThan(0);
  } finally {
    hook.unmount();
    source.native.free();
  }
});

test('session change and unmount release a font-deferred spare before another editor opens', async () => {
  for (const ending of ['session change', 'unmount']) {
    const first = setupLayoutPipeline();
    const second = setup();
    const hook = renderHook(
      ({ session }: { session: YrsSession | null }) => useWorkerLayoutPipeline(session, () => null),
      { initialProps: { session: first.engine as YrsSession | null } }
    );
    try {
      act(() => hook.result.current.runLayoutPipeline());
      const spare = FakeWorker.last!;
      expect(spare.posted.map((request) => request.type)).toEqual(['warm']);
      if (ending === 'session change') hook.rerender({ session: null });
      else hook.unmount();
      await waitFor(() => expect(spare.terminated).toBe(true));
      expect(takePreloadedResidentEngineWorker()).toBeNull();
      const next = renderHook(() => useRustDisplayList(null));
      try {
        const pending = next.result.current.layoutInWorker(second.engine, REQUEST);
        const fresh = FakeWorker.last!;
        expect(fresh).not.toBe(spare);
        expect(fresh.posted.map((request) => request.type)).toEqual(['bootstrap']);
        fresh.reply({
          id: fresh.requestAt(0).id,
          ok: true,
          frame: second.frame.slice().buffer,
          caret: { frameEpoch: 1, caretRect: null },
          selection: null,
          layoutRevision: 1,
          layoutJson: second.layoutJson,
        });
        expect(await pending).not.toBeNull();
      } finally {
        next.unmount();
      }
    } finally {
      hook.unmount();
      first.native.free();
      second.native.free();
    }
  }
});

test('a font-deferred spare the first layout does not adopt is released', async () => {
  for (const ending of ['build', 'onHost', 'preflight']) {
    const source = setupLayoutPipeline();
    Object.assign(source.engine, {
      layoutDocumentWithRegionsRetainedJson: () => source.layoutJson,
    });
    let settleFonts: () => void = () => {};
    let measurement: ResidentMeasurementConfig | null = null;
    const fonts = new Promise<void>((resolve) => (settleFonts = resolve)).then(() => {
      measurement = {} as ResidentMeasurementConfig;
    });
    const build = mock(async () => ({ pages: [] }));
    const hook = renderHook(
      ({ overrides }) => useWorkerLayoutPipeline(source.engine, () => measurement, overrides),
      { initialProps: { overrides: undefined as Parameters<typeof useRustDisplayList>[1] } }
    );
    try {
      act(() => hook.result.current.runLayoutPipeline());
      const spare = FakeWorker.last!;
      expect(spare.posted.map((request) => request.type)).toEqual(['warm']);
      if (ending === 'build') hook.rerender({ overrides: { build } });
      if (ending === 'preflight') {
        Object.assign(source.engine, {
          layoutFontRequirementsJson: () => {
            throw new Error('preflight failed');
          },
        });
      }
      await act(async () => {
        settleFonts();
        await fonts;
      });
      act(() => hook.result.current.runLayoutPipeline({ onHost: ending === 'onHost' }));
      if (ending !== 'preflight') expect(hook.result.current.layout?.pages.length).toBeGreaterThan(0);
      expect(spare.posted.map((request) => request.type)).toEqual(['warm', 'editModule']);
      expect(source.adopted).toEqual([]);
      await waitFor(() => expect(spare.terminated).toBe(true));
      expect(takePreloadedResidentEngineWorker()).toBeNull();
    } finally {
      hook.unmount();
      source.native.free();
    }
  }
});

test('font-deferred passes warm only when the worker path is eligible', () => {
  const source = setupLayoutPipeline();
  const initialWorkers = FakeWorker.instances.length;
  const build = mock(() => { throw new Error('unexpected display build'); });
  try {
    for (const disabled of ['onHost', 'Worker', 'snapshot', 'adopt', 'version', 'build']) {
      const engine = { ...source.engine };
      if (disabled === 'snapshot') Reflect.deleteProperty(engine, 'residentWorkerSnapshot');
      if (disabled === 'adopt') Reflect.deleteProperty(engine, 'adoptResidentWorkerLayout');
      if (disabled === 'version') Reflect.deleteProperty(engine, 'version');
      if (disabled === 'Worker') globalThis.Worker = undefined as unknown as typeof Worker;
      const hook = renderHook(() =>
        useWorkerLayoutPipeline(engine, () => null, disabled === 'build' ? { build } : undefined)
      );
      try {
        act(() => hook.result.current.runLayoutPipeline({ onHost: disabled === 'onHost' }));
        expect(FakeWorker.instances.length).toBe(initialWorkers);
        expect(takePreloadedResidentEngineWorker()).toBeNull();
      } finally {
        hook.unmount();
        globalThis.Worker = FakeWorker as unknown as typeof Worker;
      }
    }
    expect(build).not.toHaveBeenCalled();
  } finally {
    source.native.free();
  }
});

function setupResidentInput() {
  const document = setup();
  const paragraphs = JSON.parse(document.native.paragraphs('body')) as Array<{
    paraId: string;
    text: string;
  }>;
  const para = paragraphs[0]!;
  document.native.set_selection('body', para.paraId, para.text.length, para.paraId, para.text.length);
  const applyInput = mock((text: string, frameEpoch: number) =>
    document.native.apply_input(text, frameEpoch)
  );
  const applyDelete = mock((direction: 'backward' | 'forward', frameEpoch: number, count = 1) =>
    document.native.apply_delete(direction, frameEpoch, count)
  );
  Object.assign(document.engine, {
    resetFrameBase: () => document.native.reset_frame_base(),
    selection: () => JSON.parse(document.native.selection()) as YrsSelection,
    applyInput,
    applyDelete,
    residentDeletedUnits: () => document.native.resident_deleted_units(),
    residentCaretSnapshot: () => JSON.parse(document.native.resident_caret_snapshot_json()),
  });
  return { ...document, applyInput, applyDelete };
}

test('workerFor adopts the spare once and the next session spawns fresh', async () => {
  const first = setup();
  const second = setup();
  const initialWorkers = FakeWorker.instances.length;
  const warming = preloadDocxEngine();
  const spare = FakeWorker.last!;
  expect(spare.posted.map((request) => request.type)).toEqual(['warm']);
  spare.reply({ id: spare.requestAt(0).id, ok: true });
  await warming;
  const hook = renderHook(() =>
    useRustDisplayList(null, undefined, undefined, undefined, null)
  );
  const reply = (worker: FakeWorker, source: typeof first): void => {
    worker.reply({
      id: worker.requestAt(-1).id,
      ok: true,
      frame: source.frame.slice().buffer,
      caret: { frameEpoch: 1, caretRect: null },
      selection: null,
      layoutRevision: 1,
      layoutJson: source.layoutJson,
    });
  };
  try {
    const firstLayout = hook.result.current.layoutInWorker(first.engine, REQUEST);
    expect(FakeWorker.last).toBe(spare);
    expect(spare.posted.map((request) => request.type)).toEqual(['warm', 'editModule', 'bootstrap']);
    reply(spare, first);
    expect(await firstLayout).not.toBeNull();

    const secondLayout = hook.result.current.layoutInWorker(second.engine, REQUEST);
    const fresh = FakeWorker.last!;
    expect(fresh).not.toBe(spare);
    expect(spare.terminated).toBe(true);
    expect(fresh.posted.map((request) => request.type)).toEqual(['bootstrap']);
    expect(FakeWorker.instances.length - initialWorkers).toBe(2);
    reply(fresh, second);
    expect(await secondLayout).not.toBeNull();

    const warmingReplacement = preloadDocxEngine();
    const unusedSpare = FakeWorker.last!;
    unusedSpare.reply({ id: unusedSpare.requestAt(0).id, ok: true });
    await warmingReplacement;
    const replacementLayout = hook.result.current.layoutInWorker(first.engine, REQUEST);
    const replacement = FakeWorker.last!;
    expect(replacement).not.toBe(unusedSpare);
    expect(replacement.posted.map((request) => request.type)).toEqual(['bootstrap']);
    reply(replacement, first);
    expect(await replacementLayout).not.toBeNull();
    const remaining = takePreloadedResidentEngineWorker();
    expect(remaining).not.toBeNull();
    remaining?.destroy();
  } finally {
    hook.unmount();
    first.native.free();
    second.native.free();
  }
});

test('two display hooks cannot adopt the same spare', async () => {
  const first = setup();
  const second = setup();
  const warming = preloadDocxEngine();
  const spare = FakeWorker.last!;
  spare.reply({ id: spare.requestAt(0).id, ok: true });
  await warming;
  const firstHook = renderHook(() => useRustDisplayList(null));
  const secondHook = renderHook(() => useRustDisplayList(null));
  try {
    const firstLayout = firstHook.result.current.layoutInWorker(first.engine, REQUEST);
    const secondLayout = secondHook.result.current.layoutInWorker(second.engine, REQUEST);
    const fresh = FakeWorker.last!;
    expect(fresh).not.toBe(spare);
    for (const [worker, source] of [[spare, first], [fresh, second]] as const) {
      worker.reply({
        id: worker.requestAt(-1).id,
        ok: true,
        frame: source.frame.slice().buffer,
        caret: { frameEpoch: 1, caretRect: null },
        selection: null,
        layoutRevision: 1,
        layoutJson: source.layoutJson,
      });
    }
    expect(await firstLayout).not.toBeNull();
    expect(await secondLayout).not.toBeNull();
    firstHook.unmount();
    expect(spare.terminated).toBe(true);
    expect(fresh.terminated).toBe(false);
  } finally {
    firstHook.unmount();
    secondHook.unmount();
    first.native.free();
    second.native.free();
  }
});

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
      id: worker.requestAt(0).id,
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

test('a worker-opened document reuses its worker for the first layout', async () => {
  const { native, layoutJson, frame, engine } = setup();
  try {
    const { result, rerender, unmount } = renderHook(
      ({ layout, source }) => useRustDisplayList(
        layout, undefined, undefined, undefined, source, undefined, undefined, undefined, true
      ),
      { initialProps: { layout: null as Layout | null, source: null as YrsSession | null } }
    );
    const opening = result.current.openInWorker(engine, Uint8Array.of(1, 2, 3), 'digest', 7);
    const worker = FakeWorker.last!;
    expect(worker.posted[0]).toMatchObject({ type: 'open', digest: 'digest', generation: '7' });
    worker.reply({ id: worker.requestAt(0).id, ok: true, hostJson: '{}', stateVector: Uint8Array.of(9).buffer });
    const opened = await opening;
    expect(opened?.hostJson).toBe('{}');
    const pending = result.current.layoutInWorker(engine, REQUEST);
    expect(FakeWorker.last).toBe(worker);
    expect(worker.posted[1]).toMatchObject({ type: 'bootstrap', opened: true });
    worker.reply({
      id: worker.requestAt(1).id,
      ok: true,
      frame: frame.slice().buffer,
      caret: { frameEpoch: 1, caretRect: null },
      selection: null,
      layoutRevision: 1,
      layoutJson,
    });
    const computation = await pending!;
    await act(async () => { rerender({ layout: computation!.layout, source: engine }); });
    await waitFor(() => expect(result.current.frame?.frameEpoch).toBe(1));
    expect(worker.posted.map((request) => request.type)).toEqual(['open', 'bootstrap']);
    const encoded = opened!.encodeState();
    await waitFor(() => expect(worker.posted[2]?.type).toBe('encodeState'));
    worker.reply({ id: worker.requestAt(2).id, ok: true, state: Uint8Array.of(4, 5).buffer });
    expect(await encoded).toEqual(Uint8Array.of(4, 5));
    const count = opened!.revisionCount();
    await waitFor(() => expect(worker.posted[3]?.type).toBe('revisionCount'));
    worker.reply({ id: worker.requestAt(3).id, ok: true, revisionCount: 1 });
    expect(await count).toBe(1);
    unmount();
  } finally {
    native.free();
  }
});

test('sync recovery rejects a cached layout frame from the old owner of the same engine', async () => {
  const { native, layoutJson, frame, engine } = setup();
  const { result, rerender, unmount } = renderHook(
    ({ layout, source }) => useRustDisplayList(
      layout, undefined, undefined, undefined, source, undefined, undefined, undefined, true
    ),
    { initialProps: { layout: null as Layout | null, source: null as YrsSession | null } }
  );
  try {
    const opening = result.current.openInWorker(engine, Uint8Array.of(1));
    const old = FakeWorker.last!;
    old.reply({ id: old.requestAt(0).id, ok: true, hostJson: '{}', stateVector: new ArrayBuffer(0) });
    const opened = await opening;
    const first = result.current.layoutInWorker(engine, REQUEST);
    old.reply({
      id: old.requestAt(1).id, ok: true, frame: frame.slice().buffer,
      caret: { frameEpoch: 1, caretRect: null }, selection: null, layoutRevision: 1, layoutJson,
    });
    const cached = await first!;
    act(() => {
      const recover = opened!.fallback({ syncAccess: 'getEditorRef' });
      expect(recover && recover()).toBe(true);
    });
    expect(old.terminated).toBe(true);
    const next = result.current.layoutInWorker(engine, REQUEST);
    const current = FakeWorker.last!;
    expect(current).not.toBe(old);
    expect(current.posted[0]).toMatchObject({ type: 'bootstrap' });
    expect(current.posted[0]).not.toHaveProperty('opened');
    current.reply({
      id: current.requestAt(0).id, ok: true, frame: frame.slice().buffer,
      caret: { frameEpoch: 1, caretRect: null }, selection: null, layoutRevision: 2, layoutJson,
    });
    await next;
    await act(async () => { rerender({ layout: cached!.layout, source: engine }); });
    await waitFor(() => expect(current.posted.at(-1)).toMatchObject({ type: 'buildFrame' }));
    expect(result.current.frame).toBeNull();
  } finally {
    unmount();
    native.free();
  }
});

test('a retained worker query facade forwards within a document load but never to the next document', async () => {
  const requestB = JSON.stringify({
    ...JSON.parse(REQUEST),
    regions: { sections: [{ sectionId: 'main', properties: { pageWidth: 14400 } }] },
  });
  const documentA = setup();
  const documentB = setup(9302, 'Document B', requestB);
  const engineQueriesB = {
    displayHitTestRegionsJson: mock(() => 'null'),
    displayVerticalMoveJson: mock(() => 'null'),
    displayRangeRectsJson: mock(() => '[]'),
    displayRangeRectsRegionJson: mock(() => '[]'),
  };
  Object.assign(documentB.engine, engineQueriesB);
  try {
    const { preloadLayoutWasm } = await import('@betteroffice/docx/wasm/layout');
    await preloadLayoutWasm(
      new Uint8Array(
        readFileSync(
          resolve(import.meta.dir, '../../../../../docx/src/wasm/generated/layout/docx_layout_bg.wasm')
        )
      )
    );
    let inputs = {
      ...JSON.parse(documentA.native.retained_kernel_inputs_json()),
      ...JSON.parse(documentA.layoutJson),
    };
    const overrides = { getInputs: () => inputs };
    const { result, rerender, unmount } = renderHook(
      ({ layout, source }) => useRustDisplayList(layout, overrides, undefined, undefined, source),
      { initialProps: { layout: inputs.layout as Layout, source: documentA.engine } }
    );
    const publish = async (worker: FakeWorker, frame: Uint8Array, frameEpoch: number) => {
      await act(async () => {
        worker.reply({
          id: worker.requestAt(-1).id,
          ok: true,
          frame: frame.slice().buffer,
          caret: { frameEpoch, caretRect: null },
          selection: null,
          layoutRevision: 0,
        });
      });
      await waitFor(() => {
        if (result.current.error) throw result.current.error;
        expect(result.current.frame?.frameEpoch).toBe(frameEpoch);
        expect(result.current.queries).not.toBeNull();
        expect(result.current.workerSurfacesActive).toBe(true);
      });
      return result.current.queries!;
    };
    const workerA = FakeWorker.last!;
    expect(workerA.posted.at(-1)).toMatchObject({ type: 'bootstrap' });
    const queriesA0 = await publish(workerA, documentA.frame, 1);
    const displayListA0 = queriesA0.displayList;
    const pageSizeA0 = queriesA0.pageSize(0);
    expect(pageSizeA0).not.toBeNull();
    await queriesA0.whenReady();
    queriesA0.prime();

    const layoutA1 = documentA.native.layout_document_with_regions_retained_json(REQUEST);
    inputs = {
      ...JSON.parse(documentA.native.retained_kernel_inputs_json()),
      ...JSON.parse(layoutA1),
    };
    const frameA1 = documentA.native.build_display_list_frame('{}', 1);
    await act(async () => {
      rerender({ layout: inputs.layout, source: documentA.engine });
    });
    expect(workerA.posted.at(-1)).toMatchObject({ type: 'buildFrame', expectedFrameEpoch: 1 });
    const queriesA1 = await publish(workerA, frameA1, 2);
    await queriesA1.whenReady();
    queriesA1.prime();
    expect(queriesA1).not.toBe(queriesA0);
    expect(queriesA1.displayList).not.toBe(displayListA0);
    expect(queriesA1.displayList.pages[0]).toBe(displayListA0.pages[0]);
    expect(queriesA0.displayList).toBe(queriesA1.displayList);
    expect(queriesA0.pageSize(0)).toEqual(queriesA1.pageSize(0));

    act(() => result.current.resetSettled());
    expect(result.current.awaitingDocument()).toBe(true);
    inputs = {
      ...JSON.parse(documentB.native.retained_kernel_inputs_json()),
      ...JSON.parse(documentB.layoutJson),
    };
    await act(async () => {
      rerender({ layout: inputs.layout, source: documentB.engine });
    });
    const workerB = FakeWorker.last!;
    expect(workerB).not.toBe(workerA);
    expect(workerB.posted.at(-1)).toMatchObject({ type: 'bootstrap' });
    const queriesB = await publish(workerB, documentB.frame, 1);
    expect(result.current.awaitingDocument()).toBe(false);
    expect(queriesB.pageSize(0)).not.toEqual(pageSizeA0);
    expect(queriesB.pageSize(0)).not.toEqual(queriesA1.pageSize(0));
    const requestsToB = workerB.posted.length;

    expect(queriesA0.displayList).toBe(displayListA0);
    expect(queriesA0.displayList).not.toBe(queriesB.displayList);
    expect(queriesA0.pageSize(0)).toEqual(pageSizeA0);
    expect(queriesA0.rangeRects(0, 10)).toEqual([]);
    expect(queriesA0.hitTestRegions(0, 100, 100)).toBeNull();
    expect(queriesA0.caretRect(1)).toBeNull();
    for (const query of Object.values(engineQueriesB)) expect(query).not.toHaveBeenCalled();
    expect(workerB.posted).toHaveLength(requestsToB);
    queriesA0.dispose();
    queriesA1.dispose();
    queriesB.dispose();
    unmount();
  } finally {
    documentA.native.free();
    documentB.native.free();
  }
});

test('releasing ends forwarding from a superseded worker query facade', async () => {
  const document = setup();
  const warnings = spyOn(console, 'warn').mockImplementation(() => {});
  try {
    const { preloadLayoutWasm } = await import('@betteroffice/docx/wasm/layout');
    await preloadLayoutWasm(
      new Uint8Array(
        readFileSync(
          resolve(import.meta.dir, '../../../../../docx/src/wasm/generated/layout/docx_layout_bg.wasm')
        )
      )
    );
    const inputs = {
      ...JSON.parse(document.native.retained_kernel_inputs_json()),
      ...JSON.parse(document.layoutJson),
    };
    const overrides = { getInputs: () => inputs };
    const { result, rerender, unmount } = renderHook(
      ({ layout, resolved }) =>
        useRustDisplayList(layout, overrides, undefined, resolved, document.engine),
      {
        initialProps: {
          layout: inputs.layout as Layout,
          resolved: undefined as ReadonlySet<number> | undefined,
        },
      }
    );
    const worker = FakeWorker.last!;
    const publish = async (frame: Uint8Array, frameEpoch: number) => {
      await act(async () => {
        worker.reply({
          id: worker.requestAt(-1).id,
          ok: true,
          frame: frame.slice().buffer,
          caret: { frameEpoch, caretRect: null },
          selection: null,
          layoutRevision: 0,
        });
      });
      await waitFor(() => {
        if (result.current.error) throw result.current.error;
        expect(result.current.frame?.frameEpoch).toBe(frameEpoch);
        expect(result.current.queries).not.toBeNull();
        expect(result.current.workerSurfacesActive).toBe(true);
      });
      return result.current.queries!;
    };
    const stale = await publish(document.frame, 1);
    await stale.whenReady();
    stale.prime();
    const nextFrame = document.native.build_display_list_frame('{"resolvedCommentIds":[7]}', 1);
    await act(async () => {
      rerender({ layout: inputs.layout, resolved: new Set([7]) });
    });
    expect(worker.posted.at(-1)).toMatchObject({ type: 'buildFrame', expectedFrameEpoch: 1 });
    const live = await publish(nextFrame, 2);
    await live.whenReady();
    live.prime();
    expect(live).not.toBe(stale);
    expect(stale.displayList).toBe(live.displayList);
    expect(stale.rangeRects(1, 2).length).toBeGreaterThan(0);

    const queryEngine = await loadRustDisplayListQueryEngine();
    const reads = [
      spyOn(queryEngine, 'rangeRectsByHandle'),
      spyOn(queryEngine, 'hitTestRegionsByHandle'),
      spyOn(queryEngine, 'rangeRectsJson'),
      spyOn(queryEngine, 'hitTestRegionsJson'),
    ];
    warnings.mockClear();
    try {
      await act(async () => result.current.release());
      expect(worker.terminated).toBe(true);
      expect(result.current.presentedEngine).toBeNull();
      const requestsAfterRelease = worker.posted.length;

      expect(stale.rangeRects(1, 2)).toEqual([]);
      expect(stale.hitTestRegions(0, 100, 100)).toBeNull();
      for (const read of reads) expect(read).not.toHaveBeenCalled();
      expect(worker.posted).toHaveLength(requestsAfterRelease);
      expect(warnings).not.toHaveBeenCalled();
    } finally {
      for (const read of reads) read.mockRestore();
    }
    stale.dispose();
    live.dispose();
    unmount();
  } finally {
    warnings.mockRestore();
    document.native.free();
  }
});

test("a snapshot rebuilt after the next document starts loading keeps its own document's line", async () => {
  const requestB = JSON.stringify({
    ...JSON.parse(REQUEST),
    regions: { sections: [{ sectionId: 'main', properties: { pageWidth: 14400 } }] },
  });
  const documentA = setup();
  const documentB = setup(9302, 'Document B', requestB);
  try {
    const { preloadLayoutWasm } = await import('@betteroffice/docx/wasm/layout');
    await preloadLayoutWasm(
      new Uint8Array(
        readFileSync(
          resolve(import.meta.dir, '../../../../../docx/src/wasm/generated/layout/docx_layout_bg.wasm')
        )
      )
    );
    let inputs = {
      ...JSON.parse(documentA.native.retained_kernel_inputs_json()),
      ...JSON.parse(documentA.layoutJson),
    };
    const layoutA = inputs.layout as Layout;
    const overrides = { getInputs: () => inputs };
    const { result, rerender, unmount } = renderHook(
      ({ layout, source, resolved }) =>
        useRustDisplayList(layout, overrides, undefined, resolved, source),
      {
        initialProps: {
          layout: layoutA,
          source: documentA.engine,
          resolved: undefined as ReadonlySet<number> | undefined,
        },
      }
    );
    const publish = async (worker: FakeWorker, frame: Uint8Array, frameEpoch: number) => {
      await act(async () => {
        worker.reply({
          id: worker.requestAt(-1).id,
          ok: true,
          frame: frame.slice().buffer,
          caret: { frameEpoch, caretRect: null },
          selection: null,
          layoutRevision: 0,
        });
      });
      await waitFor(() => {
        if (result.current.error) throw result.current.error;
        expect(result.current.frame?.frameEpoch).toBe(frameEpoch);
        expect(result.current.queries).not.toBeNull();
        expect(result.current.workerSurfacesActive).toBe(true);
      });
      return result.current.queries!;
    };
    const workerA = FakeWorker.last!;
    expect(workerA.posted.at(-1)).toMatchObject({ type: 'bootstrap' });
    const queriesA0 = await publish(workerA, documentA.frame, 1);
    const displayListA0 = queriesA0.displayList;
    await queriesA0.whenReady();
    queriesA0.prime();

    act(() => result.current.resetSettled());
    expect(result.current.awaitingDocument()).toBe(true);
    const frameA1 = documentA.native.build_display_list_frame('{"resolvedCommentIds":[7]}', 1);
    await act(async () => {
      rerender({ layout: layoutA, source: documentA.engine, resolved: new Set([7]) });
    });
    expect(FakeWorker.last).toBe(workerA);
    expect(workerA.posted.at(-1)).toMatchObject({ type: 'buildFrame', expectedFrameEpoch: 1 });
    const queriesA1 = await publish(workerA, frameA1, 2);
    await queriesA1.whenReady();
    queriesA1.prime();
    const displayListA1 = queriesA1.displayList;
    const pageSizeA1 = queriesA1.pageSize(0);
    const forwardedDisplayListA0 = queriesA0.displayList;
    expect(queriesA1).not.toBe(queriesA0);
    expect(displayListA1).not.toBe(displayListA0);
    expect(pageSizeA1).not.toBeNull();
    expect(result.current.awaitingDocument()).toBe(true);

    inputs = {
      ...JSON.parse(documentB.native.retained_kernel_inputs_json()),
      ...JSON.parse(documentB.layoutJson),
    };
    await act(async () => {
      rerender({ layout: inputs.layout, source: documentB.engine, resolved: undefined });
    });
    const workerB = FakeWorker.last!;
    expect(workerB).not.toBe(workerA);
    expect(workerB.posted.at(-1)).toMatchObject({ type: 'bootstrap' });
    const queriesB = await publish(workerB, documentB.frame, 1);
    expect(result.current.awaitingDocument()).toBe(false);
    expect(queriesB.pageSize(0)).not.toEqual(pageSizeA1);

    expect(queriesA1.displayList).toBe(displayListA1);
    expect(queriesA1.displayList).not.toBe(queriesB.displayList);
    expect(queriesA1.pageSize(0)).toEqual(pageSizeA1);
    expect(forwardedDisplayListA0).toBe(displayListA1);
    queriesA0.dispose();
    queriesA1.dispose();
    queriesB.dispose();
    unmount();
  } finally {
    documentA.native.free();
    documentB.native.free();
  }
});

test('a new layout of the shown document after the next one starts loading keeps its own line', async () => {
  const requestB = JSON.stringify({
    ...JSON.parse(REQUEST),
    regions: { sections: [{ sectionId: 'main', properties: { pageWidth: 14400 } }] },
  });
  const documentA = setup();
  const documentB = setup(9302, 'Document B', requestB);
  try {
    const { preloadLayoutWasm } = await import('@betteroffice/docx/wasm/layout');
    await preloadLayoutWasm(
      new Uint8Array(
        readFileSync(
          resolve(import.meta.dir, '../../../../../docx/src/wasm/generated/layout/docx_layout_bg.wasm')
        )
      )
    );
    let inputs = {
      ...JSON.parse(documentA.native.retained_kernel_inputs_json()),
      ...JSON.parse(documentA.layoutJson),
    };
    const layoutA = inputs.layout as Layout;
    const overrides = { getInputs: () => inputs };
    const { result, rerender, unmount } = renderHook(
      ({ layout, source }) => useRustDisplayList(layout, overrides, undefined, undefined, source),
      { initialProps: { layout: layoutA, source: documentA.engine } }
    );
    const publish = async (worker: FakeWorker, frame: Uint8Array, frameEpoch: number) => {
      await act(async () => {
        worker.reply({
          id: worker.requestAt(-1).id,
          ok: true,
          frame: frame.slice().buffer,
          caret: { frameEpoch, caretRect: null },
          selection: null,
          layoutRevision: 0,
        });
      });
      await waitFor(() => {
        if (result.current.error) throw result.current.error;
        expect(result.current.frame?.frameEpoch).toBe(frameEpoch);
        expect(result.current.queries).not.toBeNull();
        expect(result.current.workerSurfacesActive).toBe(true);
      });
      return result.current.queries!;
    };
    const workerA = FakeWorker.last!;
    const queriesA0 = await publish(workerA, documentA.frame, 1);
    await queriesA0.whenReady();
    queriesA0.prime();

    act(() => result.current.resetSettled());
    expect(result.current.awaitingDocument()).toBe(true);
    const layoutA1 = documentA.native.layout_document_with_regions_retained_json(REQUEST);
    inputs = {
      ...JSON.parse(documentA.native.retained_kernel_inputs_json()),
      ...JSON.parse(layoutA1),
    };
    expect(inputs.layout).not.toBe(layoutA);
    const frameA1 = documentA.native.build_display_list_frame('{}', 1);
    await act(async () => {
      rerender({ layout: inputs.layout, source: documentA.engine });
    });
    expect(FakeWorker.last).toBe(workerA);
    expect(workerA.posted.at(-1)).toMatchObject({ type: 'buildFrame', expectedFrameEpoch: 1 });
    const queriesA1 = await publish(workerA, frameA1, 2);
    await queriesA1.whenReady();
    queriesA1.prime();
    const displayListA1 = queriesA1.displayList;
    const pageSizeA1 = queriesA1.pageSize(0);
    expect(queriesA1).not.toBe(queriesA0);
    expect(pageSizeA1).not.toBeNull();

    const layoutA2 = documentA.native.layout_document_with_regions_retained_json(REQUEST);
    inputs = {
      ...JSON.parse(documentA.native.retained_kernel_inputs_json()),
      ...JSON.parse(layoutA2),
    };
    const frameA2 = documentA.native.build_display_list_frame('{}', 2);
    await act(async () => {
      rerender({ layout: inputs.layout, source: documentA.engine });
    });
    expect(workerA.posted.at(-1)).toMatchObject({ type: 'buildFrame', expectedFrameEpoch: 2 });
    const queriesA2 = await publish(workerA, frameA2, 3);
    await queriesA2.whenReady();
    queriesA2.prime();
    expect(queriesA2.displayList).not.toBe(displayListA1);
    expect(queriesA1.displayList).toBe(queriesA2.displayList);

    inputs = {
      ...JSON.parse(documentB.native.retained_kernel_inputs_json()),
      ...JSON.parse(documentB.layoutJson),
    };
    await act(async () => {
      rerender({ layout: inputs.layout, source: documentB.engine });
    });
    const workerB = FakeWorker.last!;
    expect(workerB).not.toBe(workerA);
    expect(workerB.posted.at(-1)).toMatchObject({ type: 'bootstrap' });
    const queriesB = await publish(workerB, documentB.frame, 1);
    await queriesB.whenReady();
    queriesB.prime();
    expect(queriesB.pageSize(0)).not.toEqual(pageSizeA1);
    expect(queriesA1.displayList).toBe(displayListA1);
    expect(queriesA1.displayList).not.toBe(queriesB.displayList);
    expect(queriesA1.pageSize(0)).toEqual(pageSizeA1);
    queriesA0.dispose();
    queriesA1.dispose();
    queriesA2.dispose();
    queriesB.dispose();
    unmount();
  } finally {
    documentA.native.free();
    documentB.native.free();
  }
});

test('input that answers after the next document replaced its worker publishes nothing', async () => {
  const requestB = JSON.stringify({
    ...JSON.parse(REQUEST),
    regions: { sections: [{ sectionId: 'main', properties: { pageWidth: 14400 } }] },
  });
  const documentA = setup();
  const documentB = setup(9302, 'Document B', requestB);
  const applyLocalUpdate = mock(() => {});
  for (const document of [documentA, documentB]) {
    const paragraphs = JSON.parse(document.native.paragraphs('body')) as Array<{
      paraId: string;
      text: string;
    }>;
    const para = paragraphs[0]!;
    document.native.set_selection('body', para.paraId, para.text.length, para.paraId, para.text.length);
    Object.assign(document.engine, {
      selection: () => JSON.parse(document.native.selection()) as YrsSelection,
      applyLocalUpdate,
    });
  }
  try {
    const { preloadLayoutWasm } = await import('@betteroffice/docx/wasm/layout');
    await preloadLayoutWasm(
      new Uint8Array(
        readFileSync(
          resolve(import.meta.dir, '../../../../../docx/src/wasm/generated/layout/docx_layout_bg.wasm')
        )
      )
    );
    let inputs = {
      ...JSON.parse(documentA.native.retained_kernel_inputs_json()),
      ...JSON.parse(documentA.layoutJson),
    };
    const overrides = { getInputs: () => inputs };
    const { result, rerender, unmount } = renderHook(
      ({ layout, source }) => useRustDisplayList(layout, overrides, undefined, undefined, source),
      { initialProps: { layout: inputs.layout as Layout, source: documentA.engine } }
    );
    const workerA = FakeWorker.last!;
    await act(async () => {
      workerA.reply({
        id: workerA.requestAt(-1).id,
        ok: true,
        frame: documentA.frame.slice().buffer,
        caret: { frameEpoch: 1, caretRect: null },
        selection: documentA.engine.selection(),
        layoutRevision: 0,
      });
    });
    await waitFor(() => {
      if (result.current.error) throw result.current.error;
      expect(result.current.frame?.frameEpoch).toBe(1);
      expect(result.current.queries).not.toBeNull();
    });
    const queriesA = result.current.queries!;
    const displayListA = result.current.displayList;

    act(() => result.current.resetSettled());
    const pendingInput = result.current.applyInput('!');
    await act(async () => {
      for (let i = 0; i < 25 && workerA.posted.at(-1)?.type !== 'applyInput'; i += 1) {
        await Promise.resolve();
      }
      expect(workerA.posted.at(-1)).toMatchObject({ type: 'applyInput', expectedFrameEpoch: 1 });
    });
    const inputRequest = workerA.requestAt(-1);
    const inputFrameA = documentA.native.apply_input('!', 1);
    inputs = {
      ...JSON.parse(documentB.native.retained_kernel_inputs_json()),
      ...JSON.parse(documentB.layoutJson),
    };
    act(() => {
      workerA.reply({
        id: inputRequest.id,
        ok: true,
        frame: inputFrameA.slice().buffer,
        updates: [new Uint8Array([7]).buffer],
        deletedUnits: 0,
        caret: { frameEpoch: 2, caretRect: null },
        selection: documentA.engine.selection(),
        layoutRevision: 0,
      });
      // Replace the worker before the input reply's continuation runs.
      rerender({ layout: inputs.layout, source: documentB.engine });
    });
    const workerB = FakeWorker.last!;
    expect(workerB).not.toBe(workerA);
    expect(workerA.posted.at(-1)).toMatchObject({ type: 'destroy' });
    expect(workerB.posted.at(-1)).toMatchObject({ type: 'bootstrap' });
    const pendingQueries = result.current.resolveQueries();
    const pendingDisplayList = result.current.settledDisplayList(null, null);
    let outcome = null as ResidentFrameApplyResult | null;
    await act(async () => {
      outcome = await pendingInput;
    });
    expect(applyLocalUpdate).toHaveBeenCalledWith(new Uint8Array([7]));
    expect(outcome).toMatchObject({ frameEpoch: null, caretSynchronized: false, deletedUnits: 0 });
    expect(result.current.frame?.frameEpoch).toBe(1);
    expect(result.current.displayList).toBe(displayListA);
    expect(result.current.queries).toBe(queriesA);

    await act(async () => {
      workerB.reply({
        id: workerB.requestAt(-1).id,
        ok: true,
        frame: documentB.frame.slice().buffer,
        caret: { frameEpoch: 1, caretRect: null },
        selection: documentB.engine.selection(),
        layoutRevision: 0,
      });
    });
    await waitFor(() => {
      if (result.current.error) throw result.current.error;
      expect(result.current.frame?.frameEpoch).toBe(1);
      expect(result.current.queries).not.toBe(queriesA);
      expect(result.current.queries).not.toBeNull();
      expect(result.current.workerSurfacesActive).toBe(true);
    });
    const queriesB = result.current.queries!;
    expect(JSON.stringify(queriesB.displayList)).toContain('Document B');
    expect(queriesB.pageSize(0)).not.toEqual(queriesA.pageSize(0));
    expect((await pendingQueries)?.queries).toBe(queriesB);
    expect(await pendingDisplayList).toBe(queriesB.displayList);

    await act(async () => {
      const pending = result.current.applyInput('!');
      for (let i = 0; i < 25 && workerB.posted.at(-1)?.type !== 'applyInput'; i += 1) {
        await Promise.resolve();
      }
      expect(workerB.posted.at(-1)).toMatchObject({ type: 'applyInput', expectedFrameEpoch: 1 });
      const inputFrameB = documentB.native.apply_input('!', 1);
      workerB.reply({
        id: workerB.requestAt(-1).id,
        ok: true,
        frame: inputFrameB.slice().buffer,
        caret: { frameEpoch: 2, caretRect: null },
        selection: documentB.engine.selection(),
        layoutRevision: 0,
      });
      expect(await pending).toMatchObject({ frameEpoch: 2, caretSynchronized: false });
    });
    expect(result.current.frame?.frameEpoch).toBe(2);
    expect(JSON.stringify(result.current.queries!.displayList)).toContain('Document B');
    expect(JSON.stringify(result.current.queries!.displayList)).toContain('!');
    queriesA.dispose();
    queriesB.dispose();
    result.current.queries!.dispose();
    unmount();
  } finally {
    documentA.native.free();
    documentB.native.free();
  }
});

test('input a replaced worker rejects publishes nothing of its document', async () => {
  const requestB = JSON.stringify({
    ...JSON.parse(REQUEST),
    regions: { sections: [{ sectionId: 'main', properties: { pageWidth: 14400 } }] },
  });
  const documentA = setupResidentInput();
  const documentB = setup(9302, 'Document B', requestB);
  try {
    const { preloadLayoutWasm } = await import('@betteroffice/docx/wasm/layout');
    await preloadLayoutWasm(
      new Uint8Array(
        readFileSync(
          resolve(import.meta.dir, '../../../../../docx/src/wasm/generated/layout/docx_layout_bg.wasm')
        )
      )
    );
    let inputs = {
      ...JSON.parse(documentA.native.retained_kernel_inputs_json()),
      ...JSON.parse(documentA.layoutJson),
    };
    const overrides = { getInputs: () => inputs };
    const { result, rerender, unmount } = renderHook(
      ({ layout, source }) => useRustDisplayList(layout, overrides, undefined, undefined, source),
      { initialProps: { layout: inputs.layout as Layout, source: documentA.engine } }
    );
    const workerA = FakeWorker.last!;
    await act(async () => {
      workerA.reply({
        id: workerA.requestAt(-1).id,
        ok: true,
        frame: documentA.frame.slice().buffer,
        caret: { frameEpoch: 1, caretRect: null },
        selection: documentA.engine.selection(),
        layoutRevision: 0,
      });
    });
    await waitFor(() => {
      if (result.current.error) throw result.current.error;
      expect(result.current.frame?.frameEpoch).toBe(1);
      expect(result.current.queries).not.toBeNull();
    });
    const queriesA = result.current.queries!;
    const displayListA = result.current.displayList;

    const pendingInput = result.current.applyInput('!');
    await act(async () => {
      for (let i = 0; i < 25 && workerA.posted.at(-1)?.type !== 'applyInput'; i += 1) {
        await Promise.resolve();
      }
      expect(workerA.posted.at(-1)).toMatchObject({ type: 'applyInput', expectedFrameEpoch: 1 });
    });
    act(() => result.current.resetSettled());
    inputs = {
      ...JSON.parse(documentB.native.retained_kernel_inputs_json()),
      ...JSON.parse(documentB.layoutJson),
    };
    act(() => {
      rerender({ layout: inputs.layout, source: documentB.engine });
    });
    const workerB = FakeWorker.last!;
    expect(workerB).not.toBe(workerA);
    expect(workerA.posted.at(-1)).toMatchObject({ type: 'destroy' });
    expect(workerB.posted.at(-1)).toMatchObject({ type: 'bootstrap' });
    let outcome = null as ResidentFrameApplyResult | null;
    await act(async () => {
      outcome = await pendingInput;
    });
    expect(outcome).toEqual({ frameEpoch: null, caretSynchronized: false });
    expect(documentA.applyInput).not.toHaveBeenCalled();
    expect(documentA.applyDelete).not.toHaveBeenCalled();
    expect(result.current.displayList).toBe(displayListA);
    expect(result.current.queries).toBe(queriesA);
    expect(result.current.error).toBeNull();

    await act(async () => {
      workerB.reply({
        id: workerB.requestAt(-1).id,
        ok: true,
        frame: documentB.frame.slice().buffer,
        caret: { frameEpoch: 1, caretRect: null },
        selection: documentB.engine.selection(),
        layoutRevision: 0,
      });
    });
    await waitFor(() => {
      if (result.current.error) throw result.current.error;
      expect(result.current.frame?.frameEpoch).toBe(1);
      expect(result.current.queries).not.toBe(queriesA);
      expect(result.current.queries).not.toBeNull();
      expect(result.current.workerSurfacesActive).toBe(true);
    });
    const queriesB = result.current.queries!;
    expect(result.current.displayList).toBe(queriesB.displayList);
    expect(result.current.displayList).not.toBe(displayListA);
    expect(JSON.stringify(result.current.displayList)).toContain('Document B');
    expect(JSON.stringify(queriesB.displayList)).toContain('Document B');
    expect(queriesB.pageSize(0)).not.toEqual(queriesA.pageSize(0));
    expect(documentA.applyInput).not.toHaveBeenCalled();
    expect(documentA.applyDelete).not.toHaveBeenCalled();
    expect(result.current.error).toBeNull();
    queriesA.dispose();
    queriesB.dispose();
    unmount();
  } finally {
    documentA.native.free();
    documentB.native.free();
  }
});

test('a failed worker input replays on the host and rejoins its query line', async () => {
  const document = setupResidentInput();
  const errors = spyOn(console, 'error').mockImplementation(() => {});
  const warnings = spyOn(console, 'warn').mockImplementation(() => {});
  try {
    const { preloadLayoutWasm } = await import('@betteroffice/docx/wasm/layout');
    await preloadLayoutWasm(
      new Uint8Array(
        readFileSync(
          resolve(import.meta.dir, '../../../../../docx/src/wasm/generated/layout/docx_layout_bg.wasm')
        )
      )
    );
    const inputs = {
      ...JSON.parse(document.native.retained_kernel_inputs_json()),
      ...JSON.parse(document.layoutJson),
    };
    const overrides = { getInputs: () => inputs };
    const { result, rerender, unmount } = renderHook(
      ({ layout, resolved }) =>
        useRustDisplayList(layout, overrides, undefined, resolved, document.engine),
      {
        initialProps: {
          layout: inputs.layout as Layout,
          resolved: undefined as ReadonlySet<number> | undefined,
        },
      }
    );
    const worker = FakeWorker.last!;
    await act(async () => {
      worker.reply({
        id: worker.requestAt(-1).id,
        ok: true,
        frame: document.frame.slice().buffer,
        caret: { frameEpoch: 1, caretRect: null },
        selection: document.engine.selection(),
        layoutRevision: 0,
      });
    });
    await waitFor(() => {
      if (result.current.error) throw result.current.error;
      expect(result.current.frame?.frameEpoch).toBe(1);
      expect(result.current.queries).not.toBeNull();
    });
    const stale = result.current.queries!;
    await stale.whenReady();
    stale.prime();
    const nextFrame = document.native.build_display_list_frame('{"resolvedCommentIds":[7]}', 1);
    await act(async () => {
      rerender({ layout: inputs.layout, resolved: new Set([7]) });
    });
    expect(worker.posted.at(-1)).toMatchObject({ type: 'buildFrame', expectedFrameEpoch: 1 });
    await act(async () => {
      worker.reply({
        id: worker.requestAt(-1).id,
        ok: true,
        frame: nextFrame.slice().buffer,
        caret: { frameEpoch: 2, caretRect: null },
        selection: document.engine.selection(),
        layoutRevision: 0,
      });
    });
    await waitFor(() => {
      if (result.current.error) throw result.current.error;
      expect(result.current.frame?.frameEpoch).toBe(2);
      expect(result.current.queries).not.toBe(stale);
      expect(result.current.queries).not.toBeNull();
    });
    const workerQueries = result.current.queries!;
    await workerQueries.whenReady();
    workerQueries.prime();
    expect(stale.displayList).toBe(workerQueries.displayList);
    expect(stale.rangeRects(1, 2).length).toBeGreaterThan(0);

    const queryEngine = await loadRustDisplayListQueryEngine();
    const rangeByHandle = spyOn(queryEngine, 'rangeRectsByHandle');
    const hitByHandle = spyOn(queryEngine, 'hitTestRegionsByHandle');
    const rangeJson = spyOn(queryEngine, 'rangeRectsJson');
    const hitJson = spyOn(queryEngine, 'hitTestRegionsJson');
    const reads = [rangeByHandle, hitByHandle, rangeJson, hitJson];
    warnings.mockClear();
    let outcome = null as ResidentFrameApplyResult | null;
    try {
      document.applyInput.mockImplementation((text, frameEpoch) => {
        expect(stale.rangeRects(1, 2)).toEqual([]);
        expect(stale.hitTestRegions(0, 100, 100)).toBeNull();
        for (const read of reads) expect(read).not.toHaveBeenCalled();
        expect(warnings).not.toHaveBeenCalled();
        return document.native.apply_input(text, frameEpoch);
      });
      await act(async () => {
        const pendingInput = result.current.applyInput('!');
        for (let i = 0; i < 25 && worker.posted.at(-1)?.type !== 'applyInput'; i += 1) {
          await Promise.resolve();
        }
        expect(worker.posted.at(-1)).toMatchObject({ type: 'applyInput', expectedFrameEpoch: 2 });
        worker.onerror?.({ message: 'worker crashed' } as ErrorEvent);
        outcome = await pendingInput;
      });
      const live = result.current.queries!;
      await live.whenReady();
      live.prime();
      expect(live).not.toBe(workerQueries);
      expect(live.displayList).not.toBe(workerQueries.displayList);
      expect(stale.displayList).toBe(live.displayList);
      const rects = live.rangeRects(1, 2);
      const hit = live.hitTestRegions(0, 100, 100);
      expect(rects.length).toBeGreaterThan(0);
      for (const read of reads) read.mockClear();

      expect(stale.rangeRects(1, 2)).toEqual(rects);
      expect(stale.hitTestRegions(0, 100, 100)).toEqual(hit);
      expect(rangeByHandle).toHaveBeenCalledTimes(1);
      expect(hitByHandle).toHaveBeenCalledTimes(1);
      expect(rangeJson).not.toHaveBeenCalled();
      expect(hitJson).not.toHaveBeenCalled();
      expect(warnings).not.toHaveBeenCalled();
    } finally {
      for (const read of reads) read.mockRestore();
    }
    expect(document.applyInput).toHaveBeenCalledTimes(1);
    expect(document.applyInput).toHaveBeenCalledWith('!', 2);
    expect(document.applyDelete).not.toHaveBeenCalled();
    expect(outcome).toMatchObject({ frameEpoch: 3, caretSynchronized: false });
    expect(result.current.frame?.frameEpoch).toBe(3);
    expect(JSON.stringify(result.current.displayList)).toContain('!');
    expect(JSON.stringify(result.current.queries!.displayList)).toContain('!');
    expect(result.current.workerSurfacesActive).toBe(false);
    expect(result.current.error).toBeNull();
    stale.dispose();
    workerQueries.dispose();
    result.current.queries!.dispose();
    unmount();
  } finally {
    warnings.mockRestore();
    errors.mockRestore();
    document.native.free();
  }
});

/** Lays out the session's current state, then builds its frame on `base`. */
function laidOut(native: ReturnType<typeof createEditSession>, base: number) {
  const layoutJson = native.layout_document_with_regions_retained_json(REQUEST);
  return { frame: native.build_display_list_frame(JSON.stringify({}), base), layoutJson };
}

test('after a worker layout the host dropped, the next one paints the current text', async () => {
  const { native, paraId, layoutJson, frame, engine } = setup();
  try {
    const { result, rerender, unmount } = renderHook(
      ({ layout, source }) => useRustDisplayList(layout, undefined, undefined, undefined, source),
      { initialProps: { layout: null as Layout | null, source: null as YrsSession | null } }
    );
    const worker = () => FakeWorker.last!;
    const pass = async (built: { frame: Uint8Array; layoutJson: string }) => {
      const pending = result.current.layoutInWorker(engine, REQUEST)!;
      const request = worker().requestAt(-1);
      const epoch = decodeFrameDelta(built.frame.slice().buffer).frameEpoch;
      worker().reply({
        id: request.id,
        ok: true,
        frame: built.frame.slice().buffer,
        caret: { frameEpoch: epoch, caretRect: null },
        selection: null,
        layoutRevision: 1,
        layoutJson: built.layoutJson,
      });
      return { request, computation: (await pending)! };
    };
    const opened = await pass({ frame, layoutJson });
    await act(async () => {
      rerender({ layout: opened.computation.layout, source: engine });
    });
    await waitFor(() => expect(result.current.frame?.frameEpoch).toBe(1));

    native.insert_text('body', paraId, 0, 'Dropped ');
    const dropped = await pass(laidOut(native, 1));
    expect(dropped.request).toMatchObject({ type: 'sync', expectedFrameEpoch: 1 });

    native.insert_text('body', paraId, 0, 'Shown ');
    const shown = await pass(laidOut(native, 1));
    expect(shown.request).toMatchObject({ type: 'sync', expectedFrameEpoch: 1 });
    await act(async () => {
      rerender({ layout: shown.computation.layout, source: engine });
    });
    await waitFor(() => {
      if (result.current.error) throw result.current.error;
      expect(result.current.frame?.frameEpoch).toBeGreaterThan(1);
    });
    const text = result.current
      .frame!.displayList.pages.flatMap((page) => page.primitives)
      .map((primitive) =>
        primitive.kind === 'glyphRun' || primitive.kind === 'text' ? primitive.text : ''
      )
      .join('');
    expect(text).toContain('Shown Dropped Owned layout');
    expect(worker().posted).toHaveLength(3);
    unmount();
  } finally {
    native.free();
  }
});

test.each([[false, false], [true, false], [true, true]])(
  'a stale sync reply keeps the host path with worker-open=%s unless proposals are held=%s',
  async (experimentalWorkerOpen, holding) => {
    const { native, engine, layoutJson, frame } = setup();
    const hook = renderHook(() => useRustDisplayList(
      null, undefined, undefined, undefined, undefined, undefined, undefined, undefined,
      experimentalWorkerOpen
    ));
    try {
      const first = hook.result.current.layoutInWorker(engine, REQUEST)!;
      const worker = FakeWorker.last!;
      worker.reply({
        id: worker.requestAt(0).id, ok: true, frame: frame.slice().buffer,
        caret: { frameEpoch: 1, caretRect: null }, selection: null,
        layoutRevision: 1, layoutJson,
      });
      expect(await first).not.toBeNull();
      if (holding) {
        const snapshot = { version: '1', previewVersion: 0, proposals: [] };
        Object.assign(engine, { getProposals: () => snapshot, mirrorWorkerDocument: () => {} });
        const reply: ResidentProposalReply = {
          mirror: { version: snapshot.version, proposals: { previewVersion: 0, entries: [] } },
          result: { ok: true, snapshot }, changedStories: [],
          geometry: {
            version: snapshot.version, previewVersion: 0,
            proposals: proposalSetIdentity(snapshot), targets: {}, hidden: [],
          },
          updates: [], stateVector: new Uint8Array(),
        };
        const authority = registerWorkerProposalAuthority(engine, {
          proposal: async () => reply,
          documentRead: async () => { throw new Error('unexpected document read'); },
          handOver: async () => { throw new Error('unexpected handover'); },
        }, {
          relayout: () => {}, current: () => true, laidOut: async () => {},
          adopted: () => {}, handedOver: () => {}, contentChanged: () => {},
        });
        await authority.initialize();
        await authority.setStates({
          expectVersion: snapshot.version, expectPreviewVersion: 0, changes: [],
        }, async () => { throw new Error('unexpected main-thread toggle'); });
        expect(authority.holdsWorkerState()).toBe(true);
      }
      const older = hook.result.current.layoutInWorker(engine, REQUEST)!;
      const olderRequest = worker.requestAt(-1);
      const olderFrame = native.build_display_list_frame('{}', 0);
      const newer = hook.result.current.layoutInWorker(engine, REQUEST)!;
      const newerRequest = worker.requestAt(-1);
      const newerFrame = native.build_display_list_frame('{}', 0);
      expect(olderRequest.type).toBe('sync');
      expect(newerRequest.type).toBe('sync');
      worker.reply({
        id: newerRequest.id, ok: true, frame: newerFrame.slice().buffer,
        caret: { frameEpoch: 3, caretRect: null }, selection: null,
        layoutRevision: 3, layoutJson,
      });
      const current = await newer;
      expect(current?.layout.pages.length).toBeGreaterThan(0);
      let stale!: Awaited<typeof older>;
      await act(async () => {
        worker.reply({
          id: olderRequest.id, ok: true, frame: olderFrame.slice().buffer,
          caret: { frameEpoch: 2, caretRect: null }, selection: null,
          layoutRevision: 2, layoutJson, layoutProvisional: true,
        });
        stale = await older;
      });
      if (holding) expect(stale).toBeNull();
      else {
        expect(stale?.layout.pages.length).toBeGreaterThan(0);
        expect(stale?.complete).toBeDefined();
      }
      expect(worker.posted.some((request) => request.type === 'completeLayout')).toBe(false);
      expect(hook.result.current.error).toBeNull();
      expect(worker.terminated).toBe(false);
    } finally {
      hook.unmount();
      native.free();
    }
  }
);

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
      id: worker.requestAt(0).id,
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
        id: worker.requestAt(0).id,
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
      id: worker.requestAt(0).id,
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
    worker.reply({ id: worker.requestAt(1).id, ok: true });
    await waitFor(() => expect(worker.posted).toHaveLength(3));
    expect(worker.posted.map((request) => request.type)).toEqual([
      'bootstrap',
      'attachCanvases',
      'completeLayout',
    ]);
    expect(worker.posted[2]).toMatchObject({ expectedFrameEpoch: 1, sliceBlocks: 64 });
    let settled = false;
    void result.current.settledDisplayList(() => {}).then(() => {
      settled = true;
    });
    await act(async () => {});
    expect(settled).toBe(false);

    worker.reply({
      id: worker.requestAt(2).id,
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

test('a proposal layout is whole and reuses unchanged pages after an intervening visible build', async () => {
  let request = JSON.stringify({
    ...JSON.parse(REQUEST),
    regions: { sections: [{ sectionId: 'main', properties: {
      pageWidth: 4320, pageHeight: 2880,
      marginTop: 300, marginRight: 300, marginBottom: 300, marginLeft: 300,
    } }] },
  });
  const { native, engine } = setup(9398, 'Proposal pages. '.repeat(600), request);
  const fontId = native.register_measure_font(new Uint8Array(readFileSync(resolve(
    import.meta.dir, '../../../../../../crates/ooxml-text/tests/fonts/LiberationSans-Regular.ttf'
  ))));
  request = JSON.stringify({ ...JSON.parse(request), measurement: {
    ...JSON.parse(request).measurement,
    fontChains: { 'calibri|0|0': [fontId] }, authoritativeShaping: true,
  } });
  const layoutJson = native.layout_document_with_regions_retained_json(request);
  native.reset_frame_base();
  const frame = native.build_display_list_frame('{}', 0);
  const initialEpoch = decodeFrameDelta(frame).frameEpoch;
  const snapshot = engine.residentWorkerSnapshot.bind(engine);
  engine.residentWorkerSnapshot = (options) => ({ ...snapshot(options)!, workerAuthoritative: true });
  const originalIdle = globalThis.requestIdleCallback;
  const originalCancelIdle = globalThis.cancelIdleCallback;
  const callbacks = new Map<number, IdleRequestCallback>();
  let nextIdle = 1;
  globalThis.requestIdleCallback = (callback) => {
    const id = nextIdle++;
    callbacks.set(id, callback);
    return id;
  };
  globalThis.cancelIdleCallback = (id) => { callbacks.delete(id); };
  const runIdle = (): void => {
    const pending = [...callbacks.values()];
    callbacks.clear();
    for (const callback of pending) callback({ didTimeout: false, timeRemaining: () => 50 });
  };
  const hook = renderHook(
    ({ layout, source }) => useRustDisplayList(layout, undefined, undefined, undefined, source),
    { initialProps: { layout: null as Layout | null, source: null as YrsSession | null } }
  );
  try {
    act(() => hook.result.current.setDisplayWindow(0, 1));
    const first = hook.result.current.layoutInWorker(engine, request)!;
    const worker = FakeWorker.last!;
    worker.reply({
      id: worker.requestAt(0).id, ok: true, frame: frame.slice().buffer,
      caret: { frameEpoch: initialEpoch, caretRect: null }, selection: null, layoutRevision: 1, layoutJson,
    });
    const opened = (await act(() => first))!;
    await act(async () => hook.rerender({ layout: opened.layout, source: engine }));
    expect(hook.result.current.frame!.pages.length).toBeGreaterThan(5);
    const unchanged = hook.result.current.displayList!.pages[0];
    const proposalRequest = JSON.stringify({ ...JSON.parse(request), renderEnv: {
      revisionPreview: { proposed: 'accepted' },
    } });
    const proposal = hook.result.current.layoutInWorker(engine, proposalRequest)!;
    expect(worker.posted[1]).toMatchObject({ type: 'sync' });
    expect(worker.posted[1]).not.toHaveProperty('provisionalPages');
    const proposalJson = native.layout_document_with_regions_retained_json(proposalRequest);
    native.set_display_window(0, 1);
    native.set_windowed_incremental_builds(true);
    const proposalFrame = native.build_display_list_frame('{}', initialEpoch);
    const proposalEpoch = decodeFrameDelta(proposalFrame).frameEpoch;
    worker.reply({
      id: worker.requestAt(1).id, ok: true, frame: proposalFrame.slice().buffer,
      caret: { frameEpoch: proposalEpoch, caretRect: null }, selection: null,
      layoutRevision: 2, layoutJson: proposalJson,
    });
    const laidOut = (await act(() => proposal))!;
    expect(laidOut.complete).toBeUndefined();
    expect(laidOut.layout.partial).toBeUndefined();
    await act(async () => hook.rerender({ layout: laidOut.layout, source: engine }));
    expect(hook.result.current.displayList!.pages[0]).toBe(unchanged);
    let settled = false;
    void hook.result.current.settledDisplayList(null, null).then(() => { settled = true; });

    act(() => hook.result.current.setDisplayWindow(3, 4));
    await waitFor(() => expect(worker.posted[2]).toMatchObject({ type: 'buildPages', pages: [3] }));
    const visibleFrame = native.build_display_pages_frame(Uint32Array.of(3), proposalEpoch);
    const visibleEpoch = decodeFrameDelta(visibleFrame).frameEpoch;
    await act(async () => worker.reply({
      id: worker.requestAt(2).id, ok: true, frame: visibleFrame.slice().buffer,
      caret: { frameEpoch: visibleEpoch, caretRect: null }, selection: null, layoutRevision: 2,
    }));
    const visible = hook.result.current.displayList!.pages[3];
    let answered = 3;
    for (let round = 0; round < 100 && !settled; round += 1) {
      await act(async () => {
        await new Promise((resolve) => setTimeout(resolve, 20));
        runIdle();
        for (; answered < worker.posted.length; answered += 1) {
          const entry = worker.posted[answered]!;
          if (entry.type !== 'buildPages') continue;
          const built = native.build_display_pages_frame(Uint32Array.from(entry.pages), entry.expectedFrameEpoch);
          worker.reply({
            id: entry.id, ok: true, frame: built.slice().buffer,
            caret: { frameEpoch: decodeFrameDelta(built).frameEpoch, caretRect: null },
            selection: null, layoutRevision: 2,
          });
        }
      });
    }
    expect(settled).toBe(true);
    expect(hook.result.current.displayList!.pages[0]).toBe(unchanged);
    expect(hook.result.current.displayList!.pages[3]).toBe(visible);
    const types = worker.posted.map((entry) => ('type' in entry ? entry.type : null));
    expect(types).not.toContain('completeLayout');
    expect(types).not.toContain('buildFrame');
    const builtPages = worker.posted.flatMap((entry) => (entry.type === 'buildPages' ? entry.pages : []));
    expect(new Set(builtPages).size).toBe(builtPages.length);
    expect(builtPages).not.toContain(0);
    expect(hook.result.current.error).toBeNull();
  } finally {
    hook.unmount();
    globalThis.requestIdleCallback = originalIdle;
    globalThis.cancelIdleCallback = originalCancelIdle;
    native.free();
  }
});

test('with worker open, decisions after a font re-send lay out the whole document', async () => {
  const font = new Uint8Array(readFileSync(resolve(
    import.meta.dir, '../../../../../../crates/ooxml-text/tests/fonts/LiberationSans-Regular.ttf'
  )));
  let request = JSON.stringify({
    ...JSON.parse(REQUEST),
    regions: { sections: [{ sectionId: 'main', properties: {
      pageWidth: 4320, pageHeight: 2880,
      marginTop: 300, marginRight: 300, marginBottom: 300, marginLeft: 300,
    } }] },
  });
  const sentence = 'Decision pages. '.repeat(20);
  const { native, engine, paraId } = setup(9399, sentence.repeat(40), request);
  for (let index = 39; index > 0; index -= 1) {
    native.split_paragraph('body', paraId, index * sentence.length);
  }
  const fontId = native.register_measure_font(font);
  request = JSON.stringify({ ...JSON.parse(request), measurement: {
    ...JSON.parse(request).measurement,
    fontChains: { 'calibri|0|0': [fontId] }, authoritativeShaping: true,
  } });
  const layoutJson = native.layout_document_with_regions_retained_json(request);
  const pages = (JSON.parse(layoutJson) as { layout: Layout }).layout.pages.length;
  expect(pages).toBeGreaterThan(6);
  native.reset_frame_base();
  const frame = native.build_display_list_frame('{}', 0);
  let epoch = decodeFrameDelta(frame).frameEpoch;
  const snapshot = engine.residentWorkerSnapshot.bind(engine);
  engine.residentWorkerSnapshot = (options) => ({ ...snapshot(options)!, workerAuthoritative: true });
  const hook = renderHook(
    ({ layout, source }) => useRustDisplayList(layout, undefined, undefined, undefined, source),
    { initialProps: { layout: null as Layout | null, source: null as YrsSession | null } }
  );
  try {
    act(() => hook.result.current.setDisplayWindow(0, 1));
    const first = hook.result.current.layoutInWorker(engine, request)!;
    const worker = FakeWorker.last!;
    worker.reply({
      id: worker.requestAt(0).id, ok: true, frame: frame.slice().buffer,
      caret: { frameEpoch: epoch, caretRect: null }, selection: null, layoutRevision: 1, layoutJson,
    });
    const opened = (await act(() => first))!;
    await act(async () => hook.rerender({ layout: opened.layout, source: engine }));
    // A snapshot with a new fonts revision makes the worker register its fonts again.
    native.clear_measure_fonts();
    expect(native.register_measure_font(font)).toBe(fontId);
    for (const [index, proposed] of ['accepted', 'proposed', 'accepted'].entries()) {
      const decision = JSON.stringify({ ...JSON.parse(request), renderEnv: { revisionPreview: { proposed } } });
      const pass = hook.result.current.layoutInWorker(engine, decision)!;
      const sent = worker.posted.at(-1) as Extract<ResidentEngineWorkerRequest, { type: 'sync' }>;
      expect(sent.type).toBe('sync');
      const json = sent.provisionalPages === undefined
        ? native.layout_document_with_regions_retained_json(decision)
        : native.layout_document_with_regions_prefix_retained_json(decision, sent.provisionalPages);
      const built = native.build_display_list_frame('{}', epoch);
      epoch = decodeFrameDelta(built).frameEpoch;
      worker.reply({
        id: sent.id, ok: true, frame: built.slice().buffer,
        caret: { frameEpoch: epoch, caretRect: null }, selection: null,
        layoutRevision: index + 2, layoutJson: json,
        ...((JSON.parse(json) as { provisional?: boolean }).provisional ? { layoutProvisional: true } : {}),
      });
      const laidOut = (await act(() => pass))!;
      expect(laidOut.layout.partial).toBeUndefined();
      expect(laidOut.layout.pages).toHaveLength(pages);
      await act(async () => hook.rerender({ layout: laidOut.layout, source: engine }));
    }
  } finally {
    hook.unmount();
    native.free();
  }
});

test('with worker open, a provisional layout names its engine until the rest is asked of the worker', async () => {
  const { native, layoutJson, frame, engine } = setup();
  try {
    const { result, rerender, unmount } = renderHook(
      ({ layout, source }) =>
        useRustDisplayList(
          layout,
          undefined,
          undefined,
          undefined,
          source,
          undefined,
          undefined,
          undefined,
          true
        ),
      { initialProps: { layout: null as Layout | null, source: null as YrsSession | null } }
    );
    expect(result.current.pendingCompletion).toBeNull();
    const pending = result.current.layoutInWorker(engine, REQUEST);
    const worker = FakeWorker.last!;
    worker.reply({
      id: worker.requestAt(0).id,
      ok: true,
      frame: frame.slice().buffer,
      caret: { frameEpoch: 1, caretRect: null },
      selection: null,
      layoutRevision: 1,
      layoutJson,
      layoutProvisional: true,
    });
    const provisional = await act(() => pending!);
    expect(provisional?.layout.pages.length).toBeGreaterThan(0);
    await act(async () => {
      rerender({ layout: provisional!.layout, source: engine });
    });
    await waitFor(() => expect(result.current.frame?.frameEpoch).toBe(1));
    expect(result.current.pendingCompletion).toBe(engine);
    expect(worker.posted).toHaveLength(1);
    await act(async () => {
      void result.current.attachOffscreenCanvases([], [], 1, 1, { color: '#000', width: 2 });
    });
    worker.reply({ id: worker.requestAt(1).id, ok: true });
    await waitFor(() => expect(worker.posted).toHaveLength(3));
    expect(worker.posted[2]).toMatchObject({ type: 'completeLayout' });
    expect(result.current.pendingCompletion).toBeNull();
    unmount();
  } finally {
    native.free();
  }
});

test('an older surface timeout cannot supersede the newer provisional completion', async () => {
  const { native, layoutJson, frame, engine, adopted } = setup();
  const snapshot = engine.residentWorkerSnapshot.bind(engine);
  engine.residentWorkerSnapshot = (options) => ({
    ...snapshot(options)!, workerAuthoritative: true,
  });
  const surfaceTimers: Array<() => void> = [];
  const schedule = globalThis.setTimeout;
  const timers = spyOn(globalThis, 'setTimeout').mockImplementation(
    ((...input: Parameters<typeof setTimeout>) => {
      const [callback, delay, ...args] = input;
      if (delay === 250 && typeof callback === 'function') {
        surfaceTimers.push(() => callback(...args));
        return 0 as unknown as ReturnType<typeof setTimeout>;
      }
      return schedule(callback, delay, ...args);
    }) as typeof setTimeout
  );
  const hook = renderHook(() => useRustDisplayList(
    null, undefined, undefined, undefined, null, undefined, undefined, undefined, true
  ));
  try {
    const first = hook.result.current.layoutInWorker(engine, REQUEST)!;
    const worker = FakeWorker.last!;
    worker.reply({
      id: worker.requestAt(0).id, ok: true, frame: frame.slice().buffer,
      caret: { frameEpoch: 1, caretRect: null }, selection: null,
      layoutRevision: 1, layoutJson, layoutProvisional: true,
    });
    const older = (await act(() => first))!;
    expect(surfaceTimers).toHaveLength(1);

    const second = hook.result.current.layoutInWorker(engine, REQUEST)!;
    const nextFrame = native.build_display_list_frame('{}', 1);
    expect(worker.posted[1]).toMatchObject({ type: 'sync', snapshot: { layoutRevision: 2 } });
    worker.reply({
      id: worker.requestAt(1).id, ok: true, frame: nextFrame.slice().buffer,
      caret: { frameEpoch: 2, caretRect: null }, selection: null,
      layoutRevision: 2, layoutJson, layoutProvisional: true,
    });
    const current = (await act(() => second))!;
    expect(surfaceTimers).toHaveLength(2);
    expect(hook.result.current.pendingCompletion).toBe(engine);
    await act(async () => {
      const attaching = hook.result.current.attachOffscreenCanvases(
        [], [], 1, 1, { color: '#000', width: 2 }
      );
      worker.reply({ id: worker.requestAt(2).id, ok: true });
      expect(await attaching).toBe(true);
    });
    expect(worker.posted[3]).toMatchObject({
      type: 'completeLayout', expectedFrameEpoch: 2, sliceBlocks: 64,
    });

    await act(async () => surfaceTimers[0]!());
    expect(worker.posted.filter((request): boolean => request.type === 'completeLayout')).toEqual([
      worker.posted[3]!,
    ]);
    expect(await older.complete).toBeNull();
    const fullFrame = native.build_display_list_frame('{}', 2);
    worker.reply({
      id: worker.requestAt(3).id, ok: true, frame: fullFrame.slice().buffer,
      caret: { frameEpoch: 3, caretRect: null }, selection: null,
      layoutRevision: 2, layoutJson,
    });
    expect((await current.complete)?.layout.pages.length).toBeGreaterThan(0);
    await act(async () => surfaceTimers[1]!());
    expect(worker.posted).toHaveLength(4);
    expect(adopted).toEqual([REQUEST, REQUEST]);
    expect(hook.result.current.pendingCompletion).toBeNull();
    expect(hook.result.current.error).toBeNull();
    expect(worker.terminated).toBe(false);
  } finally {
    hook.unmount();
    timers.mockRestore();
    native.free();
  }
});

test('a worker-authoritative relayout lays out the whole document in one request', async () => {
  const { native, layoutJson, frame, engine } = setup();
  const snapshot = engine.residentWorkerSnapshot.bind(engine);
  engine.residentWorkerSnapshot = (options) => ({
    ...snapshot(options)!, workerAuthoritative: true,
  });
  try {
    const { result, unmount } = renderHook(() => useRustDisplayList(null));
    const first = result.current.layoutInWorker(engine, REQUEST);
    const worker = FakeWorker.last!;
    worker.reply({
      id: worker.requestAt(0).id, ok: true, frame: frame.slice().buffer,
      caret: { frameEpoch: 1, caretRect: null }, selection: null,
      layoutRevision: 1, layoutJson,
    });
    await first!;
    act(() => result.current.setDisplayWindow(4, 6));
    const next = result.current.layoutInWorker(engine, REQUEST);
    expect(worker.posted[1]).toMatchObject({
      type: 'sync', displayWindow: [4, 6], snapshot: { workerAuthoritative: true },
    });
    expect(worker.posted[1]).not.toHaveProperty('provisionalPages');
    const nextFrame = native.build_display_list_frame('{}', 1);
    worker.reply({
      id: worker.requestAt(1).id, ok: true, frame: nextFrame.slice().buffer,
      caret: { frameEpoch: 2, caretRect: null }, selection: null,
      layoutRevision: 2, layoutJson,
    });
    const relayout = await act(() => next!);
    expect(relayout!.complete).toBeUndefined();
    await act(() => new Promise((resolve) => setTimeout(resolve, 300)));
    expect(worker.posted.some((entry) => 'type' in entry && entry.type === 'completeLayout')).toBe(false);
    unmount();
  } finally {
    native.free();
  }
});

function settleHarness() {
  const displayList = { pages: [] };
  const overrides = {
    build: async () => displayList,
    getInputs: (): never | undefined => ({ measured: [], options: {} }) as never,
  };
  const layout = (partial: boolean) =>
    ({ pageSize: { w: 816, h: 1056 }, pages: [], ...(partial ? { partial } : {}) }) as Layout;
  const initial = layout(false);
  const hook = renderHook(
    ({ layout, resolved }: { layout: Layout | null; resolved?: ReadonlySet<number> }) =>
      useRustDisplayList(layout, overrides, undefined, resolved),
    {
      initialProps: { layout: initial } as {
        layout: Layout | null;
        resolved?: ReadonlySet<number>;
      },
    }
  );
  const settle = () => {
    const state = { settled: false, failure: null as Error | null };
    void hook.result.current.settledDisplayList(null, null).then(
      () => {
        state.settled = true;
      },
      (error: Error) => {
        state.failure = error;
      }
    );
    return state;
  };
  return { ...hook, initial, layout, settle, overrides };
}

test('a layout of part of the document never settles, even after a full one did', async () => {
  const { rerender, layout, settle } = settleHarness();
  const first = settle();
  await waitFor(() => expect(first.settled).toBe(true));
  await act(async () => {
    rerender({ layout: layout(true) });
  });
  const partial = settle();
  await act(async () => {});
  expect(partial.settled).toBe(false);
  await act(async () => {
    rerender({ layout: layout(false) });
  });
  await waitFor(() => expect(partial.settled).toBe(true));
});

test('a layout the document moved past paints but settles no wait', async () => {
  const { result, rerender, layout, settle, overrides } = settleHarness();
  const first = settle();
  await waitFor(() => expect(first.settled).toBe(true));
  const behind = layout(false);
  const behindList = { pages: [] };
  markSupersededLayout(behind);
  overrides.build = async () => behindList;
  await act(async () => {
    rerender({ layout: behind });
  });
  await waitFor(() => expect(result.current.displayList).toBe(behindList));
  const waiting = settle();
  await act(async () => {});
  expect(waiting.settled).toBe(false);
  await act(async () => {
    rerender({ layout: layout(false) });
  });
  await waitFor(() => expect(waiting.settled).toBe(true));
});

test('a reset waits for the next layout and a failure rejects', async () => {
  const { result, rerender, initial, layout, settle, overrides } = settleHarness();
  const first = settle();
  await waitFor(() => expect(first.settled).toBe(true));
  act(() => result.current.resetSettled());
  expect(result.current.awaitingDocument()).toBe(true);
  const next = settle();
  // Rebuilding the replaced document's layout, as clearing its comments does,
  // neither settles nor fails the wait.
  await act(async () => {
    rerender({ layout: initial, resolved: new Set([1]) });
  });
  const getInputs = overrides.getInputs;
  overrides.getInputs = () => undefined;
  await act(async () => {
    rerender({ layout: initial, resolved: new Set([2]) });
  });
  overrides.getInputs = getInputs;
  expect(next.settled).toBe(false);
  expect(next.failure).toBeNull();
  // The editor shows no layout while the new document's bytes load.
  await act(async () => {
    rerender({ layout: null });
  });
  expect(result.current.awaitingDocument()).toBe(true);
  await act(async () => {
    rerender({ layout: layout(false) });
  });
  await waitFor(() => expect(next.settled).toBe(true));
  act(() => result.current.resetSettled(new Error('parse failed')));
  const failed = settle();
  await waitFor(() => expect(failed.failure?.message).toBe('parse failed'));
});

test('a rejected completion after reload preserves the new session frame, queries and surfaces', async () => {
  const { native, layoutJson, frame, engine } = setup();
  const next = { ...engine } as YrsSession;
  const deferred = { reject: null as ((error: Error) => void) | null };
  const completion = spyOn(ResidentEngineWorkerClient.prototype, 'completeLayout').mockImplementation(
    () => new Promise<ResidentEngineWorkerFrame | null>((_, reject) => { deferred.reject = reject; })
  );
  const errors = spyOn(console, 'error').mockImplementation(() => {});
  try {
    const { result, rerender, unmount } = renderHook(
      ({ layout, source }) => useRustDisplayList(
        layout, undefined, undefined, undefined, source, undefined, undefined, undefined, true
      ),
      { initialProps: { layout: null as Layout | null, source: null as YrsSession | null } }
    );
    const opening = result.current.openInWorker(engine, Uint8Array.of(1));
    const oldWorker = FakeWorker.last!;
    oldWorker.reply({ id: oldWorker.requestAt(0).id, ok: true, hostJson: '{}', stateVector: new ArrayBuffer(0) });
    const opened = (await opening)!;
    const layout = result.current.layoutInWorker(engine, REQUEST)!;
    oldWorker.reply({
      id: oldWorker.requestAt(1).id,
      ok: true,
      frame: frame.slice().buffer,
      caret: { frameEpoch: 1, caretRect: null },
      selection: null,
      layoutRevision: 1,
      layoutJson,
      layoutProvisional: true,
    });
    const provisional = (await act(() => layout))!;
    await act(async () => { rerender({ layout: provisional.layout, source: engine }); });
    await waitFor(() => expect(result.current.frame).not.toBeNull());
    await act(async () => {
      const attaching = result.current.attachOffscreenCanvases([], [], 1, 1, { color: '#000', width: 2 });
      oldWorker.reply({ id: oldWorker.requestAt(-1).id, ok: true });
      await attaching;
    });
    await waitFor(() => expect(deferred.reject).not.toBeNull());
    const previousFrame = result.current.frame;
    const previousQueries = result.current.queries;
    act(() => {
      result.current.resetSettled();
      opened.destroy();
      result.current.recordSession(next);
    });
    const replacement = result.current.layoutInWorker(next, REQUEST)!;
    const newWorker = FakeWorker.last!;
    newWorker.reply({
      id: newWorker.requestAt(0).id,
      ok: true,
      frame: frame.slice().buffer,
      caret: { frameEpoch: 1, caretRect: null },
      selection: null,
      layoutRevision: 1,
      layoutJson,
    });
    const computation = (await replacement)!;
    await act(async () => { rerender({ layout: computation.layout, source: next }); });
    await waitFor(() => {
      expect(result.current.presentedEngine).toBe(next);
      expect(result.current.frame).not.toBe(previousFrame);
      expect(result.current.queries).not.toBe(previousQueries);
      expect(result.current.queries?.isReady()).toBe(true);
    });
    const currentFrame = result.current.frame;
    const currentQueries = result.current.queries;
    expect(result.current.workerSurfacesActive).toBe(true);
    act(() => result.current.setWorkerPresentationActive(true));
    await act(async () => {
      deferred.reject!(new Error('old completion failed'));
      expect(await provisional.complete!).toBeNull();
    });
    expect(result.current.frame).toBe(currentFrame);
    expect(result.current.queries).toBe(currentQueries);
    expect(currentQueries!.pageBounds(0)).not.toBeNull();
    expect(await result.current.resolveQueries()).toMatchObject({ queries: currentQueries });
    expect(result.current.workerSurfacesActive).toBe(true);
    expect(result.current.workerPresentationActive).toBe(true);
    expect(result.current.error).toBeNull();
    expect(errors).not.toHaveBeenCalled();
    expect(newWorker.posted).toHaveLength(1);
    unmount();
  } finally {
    completion.mockRestore();
    errors.mockRestore();
    native.free();
  }
});
