import { GlobalRegistrator } from '@happy-dom/global-registrator';
import { afterAll, afterEach, beforeAll, expect, mock, spyOn, test } from 'bun:test';
import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';
import type { Layout } from '@betteroffice/docx/layout/pagination';
import { createEditSession, preloadEditWasm } from '@betteroffice/docx/wasm/edit';
import { preloadDocxEngine } from '@betteroffice/docx/yrs';
import {
  takePreloadedResidentEngineWorker,
  type YrsSelection,
  type YrsSession,
} from '@betteroffice/docx/yrs';
import type {
  ResidentEngineWorkerRequest,
  ResidentEngineWorkerResponse,
} from '@betteroffice/docx/yrs/residentEngineWorkerProtocol';
import { useRustDisplayList, type ResidentFrameApplyResult } from './useDisplayList';

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
  takePreloadedResidentEngineWorker()?.destroy();
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
  posted: ResidentEngineWorkerRequest[] = [];
  terminated = false;
  constructor() {
    FakeWorker.last = this;
    FakeWorker.instances.push(this);
  }
  postMessage(request: ResidentEngineWorkerRequest): void {
    this.posted.push(request);
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
  native.create_story('body', text, 'Normal', 'left');
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
  return { native, layoutJson, frame, engine, adopted };
}

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
  spare.reply({ id: spare.posted[0].id, ok: true });
  await warming;
  const hook = renderHook(() =>
    useRustDisplayList(null, undefined, undefined, undefined, null)
  );
  const reply = (worker: FakeWorker, source: typeof first): void => {
    worker.reply({
      id: worker.posted.at(-1)!.id,
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
    expect(spare.posted.map((request) => request.type)).toEqual(['warm', 'bootstrap']);
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
    unusedSpare.reply({ id: unusedSpare.posted[0].id, ok: true });
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
  spare.reply({ id: spare.posted[0].id, ok: true });
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
        id: worker.posted.at(-1)!.id,
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
          id: worker.posted.at(-1)!.id,
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
          id: worker.posted.at(-1)!.id,
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
          id: worker.posted.at(-1)!.id,
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
        id: workerA.posted.at(-1)!.id,
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
    const inputRequest = workerA.posted.at(-1)!;
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
        id: workerB.posted.at(-1)!.id,
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
        id: workerB.posted.at(-1)!.id,
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
        id: workerA.posted.at(-1)!.id,
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
        id: workerB.posted.at(-1)!.id,
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

test('input a failed worker rejects replays on the main thread for its own document', async () => {
  const document = setupResidentInput();
  const errors = spyOn(console, 'error').mockImplementation(() => {});
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
    const { result, unmount } = renderHook(() =>
      useRustDisplayList(inputs.layout as Layout, overrides, undefined, undefined, document.engine)
    );
    const worker = FakeWorker.last!;
    await act(async () => {
      worker.reply({
        id: worker.posted.at(-1)!.id,
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
    let outcome = null as ResidentFrameApplyResult | null;
    await act(async () => {
      const pendingInput = result.current.applyInput('!');
      for (let i = 0; i < 25 && worker.posted.at(-1)?.type !== 'applyInput'; i += 1) {
        await Promise.resolve();
      }
      expect(worker.posted.at(-1)).toMatchObject({ type: 'applyInput', expectedFrameEpoch: 1 });
      worker.onerror?.({ message: 'worker crashed' } as ErrorEvent);
      outcome = await pendingInput;
    });
    expect(document.applyInput).toHaveBeenCalledTimes(1);
    expect(document.applyInput).toHaveBeenCalledWith('!', 1);
    expect(document.applyDelete).not.toHaveBeenCalled();
    expect(outcome).toMatchObject({ frameEpoch: 2, caretSynchronized: false });
    expect(result.current.frame?.frameEpoch).toBe(2);
    expect(JSON.stringify(result.current.displayList)).toContain('!');
    expect(JSON.stringify(result.current.queries!.displayList)).toContain('!');
    expect(result.current.workerSurfacesActive).toBe(false);
    expect(result.current.error).toBeNull();
    result.current.queries!.dispose();
    unmount();
  } finally {
    errors.mockRestore();
    document.native.free();
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
    expect(worker.posted[2]).toMatchObject({ expectedFrameEpoch: 1, sliceBlocks: 64 });
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
