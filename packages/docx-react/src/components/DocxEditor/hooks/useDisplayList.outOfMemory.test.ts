import { GlobalRegistrator } from '@happy-dom/global-registrator';
import { afterAll, afterEach, beforeAll, expect, mock, spyOn, test } from 'bun:test';
import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';
import { StrictMode } from 'react';
import type { Layout } from '@betteroffice/docx/layout/pagination';
import { createEditSession, preloadEditWasm } from '@betteroffice/docx/wasm/edit';
import {
  ResidentEngineWorkerClient,
  ResidentWorkerOutOfMemoryError,
  proposalSetIdentity,
  type DocxProposalSnapshot,
  type ResidentProposalReply,
  type YrsSession,
} from '@betteroffice/docx/yrs';
import type {
  ResidentEngineWorkerRequest,
  ResidentEngineWorkerResponse,
} from '@betteroffice/docx/yrs/residentEngineWorkerProtocol';
import { deferWorkerOpenReplica, requestWorkerOpenReplica } from '../internals/workerOpenReplica';
import * as workerProposals from '../internals/workerProposalAuthority';
import { EngineWorker, lazyFixture } from './__fixtures__/lazyPages';
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

function proposalAuthority(engine: YrsSession) {
  let mirror: ResidentProposalReply['mirror'] | null = null;
  const snapshot: DocxProposalSnapshot = { version: 'worker-1', previewVersion: 0, proposals: [] };
  const reply: ResidentProposalReply = {
    mirror: { version: 'worker-1', proposals: { previewVersion: 0, entries: [] } },
    result: { ok: true, snapshot },
    geometry: { version: 'worker-1', previewVersion: 0, proposals: proposalSetIdentity(snapshot), targets: {}, hidden: [] },
    changedStories: [], updates: [], stateVector: new Uint8Array(),
  };
  Object.assign(engine, {
    version: () => mirror?.version ?? 'main-1',
    mirrorWorkerDocument: (next: typeof mirror) => { mirror = next; },
    getProposals: () => ({ ...snapshot, version: engine.version() }),
  });
  const worker = {
    proposal: mock(async () => reply),
    documentRead: mock(async () => { throw new Error('unexpected worker read'); }),
    handOver: mock(async () => ({
      state: Uint8Array.of(1), version: 'worker-1', proposals: reply.mirror.proposals,
    })),
  };
  const authority = workerProposals.registerWorkerProposalAuthority(engine, worker, {
    relayout: () => {}, current: () => true, laidOut: () => Promise.resolve(),
    adopted: () => {}, handedOver: () => {}, contentChanged: () => {},
  });
  const hold = async () => {
    await authority.setStates({ expectVersion: 'worker-1', expectPreviewVersion: 0, changes: [] }, async () => {
      throw new Error('unexpected main call');
    });
    expect(authority.holdsWorkerState()).toBe(true);
  };
  return { authority, worker, hold };
}

for (const stage of ['fontRequirements', 'bootstrap'] as const) {
  test(`worker-held proposals reject ${stage} OOM without reopening or hydrating`, async () => {
    const { native, engine, mainThreadBuilds } = setup();
    const { authority, hold } = proposalAuthority(engine);
    const fallback = spyOn(console, 'error').mockImplementation(() => {});
    let mainOpens = 0;
    const replica = deferWorkerOpenReplica(engine, () => new Promise(() => {}), () => {
      mainOpens += 1;
    }, () => {});
    const { result, unmount } = renderHook(() =>
      useRustDisplayList(null, undefined, undefined, undefined, engine, undefined, undefined, undefined, true)
    );
    try {
      const opening = result.current.openInWorker(engine, Uint8Array.of(1));
      const worker = FakeWorker.spawned[0]!;
      worker.onmessage?.({ data: {
        id: worker.last().id, ok: true, hostJson: '{}', stateVector: new ArrayBuffer(0),
      } } as MessageEvent<ResidentEngineWorkerResponse>);
      const opened = await opening;
      expect(opened).not.toBeNull();
      await hold();
      const ready = replica.ready.catch((error: unknown) => error);
      const settled = result.current.settledDisplayList(null, null).catch((error: unknown) => error);
      let pending!: Promise<unknown>;
      await act(async () => {
        pending = stage === 'fontRequirements'
          ? result.current.fontRequirementsInWorker(engine, REQUEST)!
          : result.current.layoutInWorker(engine, REQUEST)!;
      });
      expect(worker.last().type).toBe(stage);
      const rejected = pending.catch((error: unknown) => error);
      await act(async () => {
        worker.outOfMemory();
        expect(await rejected).toBeInstanceOf(ResidentWorkerOutOfMemoryError);
      });
      expect(FakeWorker.spawned).toHaveLength(1);
      expect(worker.posted.filter((request) => request.type === 'open')).toHaveLength(1);
      expect(worker.posted.some((request) => request.type === 'encodeState')).toBe(false);
      expect(mainThreadBuilds).toEqual([]);
      expect(mainOpens).toBe(0);
      const failure = await ready;
      expect(failure).toBeInstanceOf(ResidentWorkerOutOfMemoryError);
      expect((failure as Error).message).toBe('Resident engine worker ran out of memory allocating 65536 bytes: unreachable');
      expect(replica.pending).toBe(false);
      expect(replica.started).toBe(true);
      expect(await settled).toBe(failure);
      expect(result.current.error).toBe(failure as Error);
      await expect(pending).rejects.toBe(failure);
      await expect(authority.getProposals(async () => engine.getProposals())).rejects.toBe(failure);
      await expect(workerProposals.beginWorkerProposalHandover(engine)!).rejects.toBe(failure);
      await expect(opened!.revisionCount()).rejects.toBe(failure);
      expect(FakeWorker.spawned).toHaveLength(1);
    } finally {
      unmount();
      fallback.mockRestore();
      native.free();
    }
  });
}

test('a page-build failure with worker-held proposals fails the document and settles waits', async () => {
  const { engine, inputs, host } = lazyFixture();
  const { authority, hold } = proposalAuthority(host);
  const replica = deferWorkerOpenReplica(host, () => new Promise(() => {}), () => {
    throw new Error('unexpected hydration');
  }, () => {});
  const errors = spyOn(console, 'error').mockImplementation(() => {});
  const unhandled = mock(() => {});
  process.on('unhandledRejection', unhandled);
  const relayout = mock(() => {});
  try {
    const overrides = { getInputs: () => inputs };
    const { result, rerender, unmount } = renderHook(({ layout }) =>
      useRustDisplayList(layout, overrides, undefined, undefined, host, relayout),
      { initialProps: { layout: inputs.layout as Layout } }
    );
    await waitFor(() => expect(result.current.frame).not.toBeNull());
    expect(result.current.displayList!.pages.some((page) => page.unbuilt)).toBe(true);
    await hold();
    EngineWorker.failPageBuilds = true;
    const ready = replica.ready.catch((error: unknown) => error);
    let failure: unknown;
    await act(async () => {
      failure = await result.current.settledDisplayList(null, null).catch((error: unknown) => error);
    });
    await waitFor(() => expect(result.current.error).toBe(failure as Error));
    expect(failure).toBeInstanceOf(Error);
    expect((failure as Error).message).toBe('page build failed');
    expect(await ready).toBe(failure);
    expect(replica.pending).toBe(false);
    await expect(authority.getProposals(async () => host.getProposals())).rejects.toBe(failure);
    await expect(workerProposals.beginWorkerProposalHandover(host)!).rejects.toBe(failure);
    expect(EngineWorker.last!.terminated).toBe(true);
    expect(EngineWorker.spawned).toBe(1);
    expect(result.current.loading).toBe(false);
    await act(async () => rerender({ layout: { ...inputs.layout } as Layout }));
    expect(result.current.error).toBe(failure as Error);
    await expect(result.current.settledDisplayList(null, null)).rejects.toBe(failure);
    expect(EngineWorker.spawned).toBe(1);
    expect(errors).toHaveBeenCalledTimes(1);
    expect(relayout).not.toHaveBeenCalled();
    expect(unhandled).not.toHaveBeenCalled();
    unmount();
  } finally {
    cleanup();
    process.off('unhandledRejection', unhandled);
    errors.mockRestore();
    engine.free();
  }
});

test('a terminal failure with worker-held proposals rejects concurrent calls after layout settles', async () => {
  const { native, inputs, frame, engine, mainThreadBuilds } = setup();
  const { authority, hold, worker: transport } = proposalAuthority(engine);
  let mainOpens = 0;
  const replica = deferWorkerOpenReplica(engine, () => new Promise(() => {}), () => {
    mainOpens += 1;
  }, () => {});
  const errors = spyOn(console, 'error').mockImplementation(() => {});
  try {
    const overrides = { getInputs: () => inputs };
    const { result, rerender, unmount } = renderHook(
      ({ layout }) =>
        useRustDisplayList(layout, overrides, undefined, undefined, engine, undefined, undefined, undefined, true),
      { initialProps: { layout: null as Layout | null } }
    );
    const opening = result.current.openInWorker(engine, Uint8Array.of(1));
    const worker = FakeWorker.spawned[0]!;
    worker.onmessage?.({ data: {
      id: worker.last().id, ok: true, hostJson: '{}', stateVector: new ArrayBuffer(0),
    } } as MessageEvent<ResidentEngineWorkerResponse>);
    const opened = await opening;
    expect(opened).not.toBeNull();
    await act(async () => rerender({ layout: inputs.layout as Layout }));
    expect(worker.last()).toMatchObject({ type: 'bootstrap', opened: true });
    await act(async () => worker.replyFrame(frame(1), 1));
    await hold();
    await act(async () => { await result.current.settledDisplayList(null, null); });
    expect(result.current.error).toBeNull();
    const ready = replica.ready.catch((error: unknown) => error);
    let pending!: Promise<ResidentProposalReply>;
    await act(async () => {
      pending = opened!.proposal({
        kind: 'setStates',
        request: { expectVersion: 'worker-1', expectPreviewVersion: 0, changes: [] },
      });
    });
    expect(worker.last()).toMatchObject({ type: 'proposal', operation: { kind: 'setStates' } });
    const rejected = pending.catch((error: unknown) => error);
    let revisionCount!: Promise<unknown>;
    let requirements!: Promise<unknown>;
    let handover!: Promise<unknown>;
    let proposals!: Promise<unknown>;
    transport.handOver.mockImplementation(() => new Promise(() => {}));
    await act(async () => {
      revisionCount = opened!.revisionCount().catch((error: unknown) => error);
      requirements = result.current.fontRequirementsInWorker(engine, REQUEST)!.catch((error: unknown) => error);
      handover = workerProposals.beginWorkerProposalHandover(engine)!.catch((error: unknown) => error);
      proposals = authority.getProposals(async () => engine.getProposals()).catch((error: unknown) => error);
    });
    expect(worker.posted.some((request) => request.type === 'fontRequirements')).toBe(true);
    expect(transport.handOver).toHaveBeenCalledTimes(1);
    let failure: unknown;
    await act(async () => {
      worker.trapped();
      failure = await rejected;
    });
    expect(failure).toBeInstanceOf(Error);
    expect(failure).not.toBeInstanceOf(ResidentWorkerOutOfMemoryError);
    expect((failure as Error).message).toBe('Resident engine worker trapped: unreachable');
    await expect(pending).rejects.toBe(failure);
    expect(await revisionCount).toBe(failure);
    expect(await requirements).toBe(failure);
    expect(await handover).toBe(failure);
    expect(await proposals).toBe(failure);
    await expect(workerProposals.beginWorkerProposalHandover(engine)!).rejects.toBe(failure);
    await expect(authority.getProposals(async () => engine.getProposals())).rejects.toBe(failure);
    await expect(authority.readParagraphs({ view: 'accepted' }, async () => {
      throw new Error('unexpected main read');
    })).rejects.toBe(failure);
    expect(await ready).toBe(failure);
    expect(replica.pending).toBe(false);
    await expect(result.current.settledDisplayList(null, null)).rejects.toBe(failure);
    expect(result.current.error).toBe(failure as Error);
    expect(result.current.loading).toBe(false);
    expect(errors).toHaveBeenCalledTimes(1);
    expect(worker.terminated).toBe(true);
    expect(FakeWorker.spawned).toHaveLength(1);
    expect(worker.posted.filter((request) => request.type === 'open')).toHaveLength(1);
    expect(worker.posted.some((request) => request.type === 'encodeState')).toBe(false);
    expect(mainOpens).toBe(0);
    expect(mainThreadBuilds).toEqual([]);
    unmount();
  } finally {
    cleanup();
    errors.mockRestore();
    native.free();
  }
});

test('a provisional completion failure with worker-held proposals fails the document and settles waits', async () => {
  const { native, frame, engine, layoutJson, mainThreadBuilds } = setup();
  const { authority, hold } = proposalAuthority(engine);
  const replica = deferWorkerOpenReplica(engine, () => new Promise(() => {}), () => {
    throw new Error('unexpected hydration');
  }, () => {});
  const errors = spyOn(console, 'error').mockImplementation(() => {});
  try {
    const { result, unmount } = renderHook(() =>
      useRustDisplayList(null, undefined, undefined, undefined, engine)
    );
    const pending = result.current.layoutInWorker(engine, REQUEST)!;
    const worker = FakeWorker.spawned[0]!;
    await act(async () => worker.replyFrame(frame(1), 1, { layoutJson, layoutProvisional: true }));
    const provisional = (await pending) as { complete?: Promise<unknown> };
    expect(provisional.complete).toBeDefined();
    await hold();
    const completed = provisional.complete!.catch((error: unknown) => error);
    const ready = replica.ready.catch((error: unknown) => error);
    const settled = result.current.settledDisplayList(null, null).catch((error: unknown) => error);
    await act(async () => {
      const attached = result.current.attachOffscreenCanvases([], [], 1, 1, { color: '#000', width: 2 });
      expect(worker.last()).toMatchObject({ type: 'attachCanvases' });
      worker.onmessage?.({ data: { id: worker.last().id, ok: true } } as MessageEvent<ResidentEngineWorkerResponse>);
      expect(await attached).toBe(true);
    });
    expect(worker.last()).toMatchObject({ type: 'completeLayout' });
    await act(async () => {
      worker.trapped();
      await completed;
    });
    const failure = await completed;
    expect(failure).toBeInstanceOf(Error);
    expect((failure as Error).message).toBe('Resident engine worker trapped: unreachable');
    expect(await ready).toBe(failure);
    expect(await settled).toBe(failure);
    expect(result.current.error).toBe(failure as Error);
    expect(result.current.loading).toBe(false);
    expect(replica.pending).toBe(false);
    await expect(authority.getProposals(async () => engine.getProposals())).rejects.toBe(failure);
    expect(worker.terminated).toBe(true);
    expect(FakeWorker.spawned).toHaveLength(1);
    expect(mainThreadBuilds).toEqual([]);
    expect(errors).toHaveBeenCalledTimes(1);
    unmount();
  } finally {
    cleanup();
    errors.mockRestore();
    native.free();
  }
});

test('a current superseded completion retries while the worker holds proposals', async () => {
  const { native, frame, engine, layoutJson, mainThreadBuilds } = setup();
  const { authority, hold } = proposalAuthority(engine);
  const hydrate = mock(() => new Promise<() => void>(() => {}));
  const fallback = mock(() => { throw new Error('unexpected hydration'); });
  const replica = deferWorkerOpenReplica(engine, hydrate, fallback, () => {});
  const hook = renderHook(() => useRustDisplayList(
    null, undefined, undefined, undefined, engine, undefined, undefined, undefined, true
  ));
  try {
    const pending = hook.result.current.layoutInWorker(engine, REQUEST)!;
    const worker = FakeWorker.spawned[0]!;
    await act(async () => worker.replyFrame(frame(1), 1, { layoutJson, layoutProvisional: true }));
    const provisional = (await pending)!;
    await hold();
    await act(async () => {
      const attaching = hook.result.current.attachOffscreenCanvases(
        [], [], 1, 1, { color: '#000', width: 2 }
      );
      worker.onmessage?.({ data: { id: worker.last().id, ok: true } } as MessageEvent<ResidentEngineWorkerResponse>);
      expect(await attaching).toBe(true);
    });
    for (let attempt = 0; attempt < 2; attempt += 1) {
      const completion = worker.last();
      expect(completion).toMatchObject({
        type: 'completeLayout', expectedFrameEpoch: 1, sliceBlocks: 64,
      });
      await act(async () => worker.onmessage?.({
        data: { id: completion.id, ok: true },
      } as MessageEvent<ResidentEngineWorkerResponse>));
      expect(worker.last()).toMatchObject({
        type: 'completeLayout', expectedFrameEpoch: 1, sliceBlocks: 64,
      });
      expect(worker.last().id).not.toBe(completion.id);
    }
    await act(async () => worker.replyFrame(frame(2), 2, { layoutJson }));
    expect((await provisional.complete)?.layout.pages.length).toBeGreaterThan(0);
    expect(worker.posted.map((request) => request.type)).toEqual([
      'bootstrap', 'attachCanvases', 'completeLayout', 'completeLayout', 'completeLayout',
    ]);
    expect(authority.holdsWorkerState()).toBe(true);
    expect(replica.pending).toBe(true);
    expect(hydrate).not.toHaveBeenCalled();
    expect(fallback).not.toHaveBeenCalled();
    expect(mainThreadBuilds).toEqual([]);
    expect(hook.result.current.error).toBeNull();
    expect(worker.terminated).toBe(false);
    expect(FakeWorker.spawned).toHaveLength(1);
  } finally {
    hook.unmount();
    native.free();
  }
});

test('a stale superseded completion never retries over the newer proposal layout', async () => {
  const { native, frame, engine, layoutJson, mainThreadBuilds } = setup();
  const { host, adopted } = revisedHost(engine);
  const { authority, hold } = proposalAuthority(host);
  const hook = renderHook(() => useRustDisplayList(null));
  try {
    const pending = hook.result.current.layoutInWorker(host, REQUEST)!;
    const worker = FakeWorker.spawned[0]!;
    await act(async () => worker.replyFrame(frame(1), 1, { layoutJson, layoutProvisional: true }));
    const provisional = (await pending)!;
    await hold();
    await act(async () => {
      const attaching = hook.result.current.attachOffscreenCanvases(
        [], [], 1, 1, { color: '#000', width: 2 }
      );
      worker.onmessage?.({ data: { id: worker.last().id, ok: true } } as MessageEvent<ResidentEngineWorkerResponse>);
      expect(await attaching).toBe(true);
    });
    const completion = worker.last();
    expect(completion.type).toBe('completeLayout');
    const next = hook.result.current.layoutInWorker(host, REQUEST)!;
    const sync = worker.last();
    expect(sync.type).toBe('sync');
    await act(async () => worker.onmessage?.({
      data: { id: completion.id, ok: true },
    } as MessageEvent<ResidentEngineWorkerResponse>));
    expect(await provisional.complete).toBeNull();
    expect(worker.last()).toBe(sync);
    await act(async () => worker.replyFrame(frame(2), 2, { layoutJson, layoutRevision: 2 }));
    expect((await next)?.layout.pages.length).toBeGreaterThan(0);
    expect(worker.posted.map((request) => request.type)).toEqual([
      'bootstrap', 'attachCanvases', 'completeLayout', 'sync',
    ]);
    expect(adopted).toEqual([REQUEST, REQUEST]);
    expect(authority.holdsWorkerState()).toBe(true);
    expect(mainThreadBuilds).toEqual([]);
    expect(hook.result.current.error).toBeNull();
    expect(worker.terminated).toBe(false);
  } finally {
    hook.unmount();
    native.free();
  }
});

test('a completed proposal hand-over allows worker OOM replacement and main-thread fallback', async () => {
  const { native, inputs, frame, engine, mainThreadBuilds } = setup();
  const { authority, hold } = proposalAuthority(engine);
  const replica = deferWorkerOpenReplica(engine, async () => {
    const handover = await workerProposals.beginWorkerProposalHandover(engine)!;
    return () => handover.complete();
  }, () => { throw new Error('unexpected hydration fallback'); }, () => {});
  const warnings = spyOn(console, 'warn').mockImplementation(() => {});
  const errors = spyOn(console, 'error').mockImplementation(() => {});
  try {
    const overrides = { getInputs: () => inputs };
    const { result, rerender, unmount } = renderHook(
      ({ layout }) => useRustDisplayList(layout, overrides, undefined, undefined, engine),
      { initialProps: { layout: inputs.layout as Layout } }
    );
    const first = FakeWorker.spawned[0]!;
    await act(async () => first.replyFrame(frame(1), 1));
    await hold();
    await act(async () => { await requestWorkerOpenReplica(engine); });
    expect(replica.pending).toBe(false);
    expect(authority.holdsWorkerState()).toBe(false);
    expect(workerProposals.registeredWorkerProposalAuthority(engine)).toBe(authority);

    await act(async () => rerender({ layout: { ...inputs.layout } }));
    await act(async () => first.outOfMemory());
    expect(first.terminated).toBe(true);
    expect(FakeWorker.spawned).toHaveLength(2);
    const second = FakeWorker.spawned[1]!;
    expect(second.last()).toMatchObject({ type: 'bootstrap' });
    await act(async () => second.replyFrame(frame(2), 2));
    expect(result.current.error).toBeNull();
    expect(mainThreadBuilds).toEqual([]);

    await act(async () => rerender({ layout: { ...inputs.layout } }));
    expect(second.last()).toMatchObject({ type: 'buildFrame' });
    await act(async () => second.trapped());
    await waitFor(() => expect(result.current.loading).toBe(false));
    expect(second.terminated).toBe(true);
    expect(result.current.error).toBeNull();
    expect(result.current.workerSurfacesActive).toBe(false);
    expect(mainThreadBuilds).toHaveLength(1);
    expect(FakeWorker.spawned).toHaveLength(2);
    unmount();
  } finally {
    cleanup();
    warnings.mockRestore();
    errors.mockRestore();
    native.free();
  }
});

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

test('a renderer remounted by StrictMode still replaces a worker that runs out of memory', async () => {
  const { native, inputs, frame, engine, mainThreadBuilds } = setup();
  const warnings = spyOn(console, 'warn').mockImplementation(() => {});
  const errors = spyOn(console, 'error').mockImplementation(() => {});
  try {
    const overrides = { getInputs: () => inputs };
    const { result, unmount } = renderHook(
      () => useRustDisplayList(inputs.layout as Layout, overrides, undefined, undefined, engine),
      { wrapper: StrictMode }
    );
    expect(FakeWorker.spawned).toHaveLength(2);
    expect(FakeWorker.spawned[0]!.terminated).toBe(true);
    const live = FakeWorker.spawned[1]!;
    expect(live.last()).toMatchObject({ type: 'bootstrap' });
    expect(result.current.loading).toBe(true);
    await act(async () => live.outOfMemory());
    expect(live.terminated).toBe(true);
    expect(FakeWorker.spawned).toHaveLength(3);
    const replacement = FakeWorker.spawned[2]!;
    expect(replacement.last()).toMatchObject({ type: 'bootstrap' });
    await act(async () => replacement.replyFrame(frame(1), 1));
    await waitFor(() => expect(result.current.frame?.frameEpoch).toBe(1));
    expect(result.current.displayList).not.toBeNull();
    expect(result.current.loading).toBe(false);
    expect(result.current.error).toBeNull();
    expect(result.current.workerSurfacesActive).toBe(true);
    expect(replacement.terminated).toBe(false);
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

test('a provisional layout whose worker ran out of memory twice rejects its completion instead of laying out here', async () => {
  const { native, inputs, frame, engine, layoutJson, mainThreadBuilds } = setup();
  let rejectCompletion!: (cause: unknown) => void;
  const delayedCompletion = new Promise<never>((_, reject) => {
    rejectCompletion = reject;
  });
  const completeLayout = spyOn(ResidentEngineWorkerClient.prototype, 'completeLayout')
    .mockImplementationOnce(() => delayedCompletion);
  const warnings = spyOn(console, 'warn').mockImplementation(() => {});
  const errors = spyOn(console, 'error').mockImplementation(() => {});
  try {
    const overrides = { getInputs: () => inputs };
    const { result, rerender, unmount } = renderHook(
      ({ layout }) => useRustDisplayList(layout, overrides, undefined, undefined, engine),
      { initialProps: { layout: null as Layout | null } }
    );
    const pending = result.current.layoutInWorker(engine, REQUEST)!;
    const first = FakeWorker.spawned[0]!;
    expect(first.last()).toMatchObject({ type: 'bootstrap', provisionalPages: 3 });
    await act(async () => first.outOfMemory());
    expect(first.terminated).toBe(true);
    expect(FakeWorker.spawned).toHaveLength(2);
    const second = FakeWorker.spawned[1]!;
    expect(second.last()).toMatchObject({ type: 'bootstrap', provisionalPages: 3 });
    await act(async () =>
      second.replyFrame(frame(1), 1, { layoutJson, layoutProvisional: true })
    );
    const provisional = (await pending) as {
      layout: Layout;
      complete?: Promise<unknown>;
    };
    expect(provisional.complete).toBeDefined();
    await act(async () => rerender({ layout: provisional.layout }));
    await waitFor(() => expect(result.current.frame?.frameEpoch).toBe(1));
    expect(result.current.error).toBeNull();

    await act(async () => {
      const attached = result.current.attachOffscreenCanvases([], [], 1, 1, {
        color: '#000',
        width: 2,
      });
      expect(second.last()).toMatchObject({ type: 'attachCanvases' });
      second.outOfMemory();
      expect(await attached).toBe(false);
    });
    await waitFor(() => expect(completeLayout).toHaveBeenCalledTimes(1));
    expect(second.terminated).toBe(true);
    expect(FakeWorker.spawned).toHaveLength(2);
    const failure = result.current.error;
    expect(failure).toBeInstanceOf(ResidentWorkerOutOfMemoryError);
    expect(failure?.message).toBe(
      'Resident engine worker ran out of memory allocating 65536 bytes: unreachable'
    );
    expect(mainThreadBuilds).toEqual([]);

    await act(async () => {
      rejectCompletion(failure);
      await expect(provisional.complete!).rejects.toBeInstanceOf(ResidentWorkerOutOfMemoryError);
    });
    await expect(provisional.complete!).rejects.toBe(failure);
    expect(FakeWorker.spawned).toHaveLength(2);
    expect(mainThreadBuilds).toEqual([]);
    expect(result.current.error).toBe(failure);
    unmount();
  } finally {
    cleanup();
    completeLayout.mockRestore();
    warnings.mockRestore();
    errors.mockRestore();
    native.free();
  }
});

test('a pass whose session no worker serves any more starts no worker when its completion runs out of memory', async () => {
  const { native, inputs, frame, engine, layoutJson, mainThreadBuilds } = setup();
  let released = false;
  const afterRelease: string[] = [];
  const engineA = new Proxy(engine, {
    get(target, key, receiver) {
      const value = Reflect.get(target, key, receiver);
      if (typeof value !== 'function') return value;
      return (...args: unknown[]) => {
        if (released) afterRelease.push(String(key));
        return value.apply(target, args);
      };
    },
  });
  const engineB = { ...engine } as YrsSession;
  let rejectCompletion!: (cause: unknown) => void;
  const delayedCompletion = new Promise<never>((_, reject) => {
    rejectCompletion = reject;
  });
  const completeLayout = spyOn(ResidentEngineWorkerClient.prototype, 'completeLayout')
    .mockImplementationOnce(() => delayedCompletion);
  const warnings = spyOn(console, 'warn').mockImplementation(() => {});
  const errors = spyOn(console, 'error').mockImplementation(() => {});
  try {
    const overrides = { getInputs: () => inputs };
    const { result, rerender, unmount } = renderHook(
      ({ layout, source }) => useRustDisplayList(layout, overrides, undefined, undefined, source),
      { initialProps: { layout: null as Layout | null, source: engineA } }
    );
    const pending = result.current.layoutInWorker(engineA, REQUEST)!;
    const workerA = FakeWorker.spawned[0]!;
    await act(async () =>
      workerA.replyFrame(frame(1), 1, { layoutJson, layoutProvisional: true })
    );
    const provisional = (await pending) as {
      layout: Layout;
      complete?: Promise<unknown>;
    };
    expect(provisional.complete).toBeDefined();
    const completed = provisional.complete!.catch((cause: unknown) => cause);
    await act(async () => rerender({ layout: provisional.layout, source: engineA }));
    await waitFor(() => expect(result.current.frame?.frameEpoch).toBe(1));
    await act(async () => {
      const attached = result.current.attachOffscreenCanvases([], [], 1, 1, {
        color: '#000',
        width: 2,
      });
      expect(workerA.last()).toMatchObject({ type: 'attachCanvases' });
      workerA.outOfMemory();
      expect(await attached).toBe(false);
    });
    await waitFor(() => expect(completeLayout).toHaveBeenCalledTimes(1));
    expect(workerA.terminated).toBe(true);
    expect(FakeWorker.spawned).toHaveLength(1);
    expect(warnings).toHaveBeenCalledTimes(1);

    await act(async () => {
      const pendingB = result.current.layoutInWorker(engineB, REQUEST)!;
      const workerB = FakeWorker.spawned[1]!;
      expect(workerB.last()).toMatchObject({ type: 'bootstrap' });
      workerB.trapped();
      expect(await pendingB).toBeNull();
    });
    expect(FakeWorker.spawned[1]!.terminated).toBe(true);
    await act(async () => rerender({ layout: inputs.layout as Layout, source: engineB }));
    await waitFor(() => {
      expect(result.current.frame).not.toBeNull();
      expect(result.current.loading).toBe(false);
    });
    expect(result.current.error).toBeNull();
    expect(result.current.displayList).not.toBeNull();
    expect(result.current.workerSurfacesActive).toBe(false);
    expect(mainThreadBuilds).toEqual([1]);
    const shownFrame = result.current.frame;
    const reportedErrors = errors.mock.calls.length;
    released = true;

    await act(async () =>
      rejectCompletion(new ResidentWorkerOutOfMemoryError('delayed completion ran out of memory', []))
    );
    expect(afterRelease).toEqual([]);
    expect(FakeWorker.spawned).toHaveLength(2);
    expect(await completed).toBeNull();
    expect(result.current.frame).toBe(shownFrame);
    expect(result.current.loading).toBe(false);
    expect(result.current.error).toBeNull();
    expect(mainThreadBuilds).toEqual([1]);
    expect(errors).toHaveBeenCalledTimes(reportedErrors);
    unmount();
  } finally {
    cleanup();
    completeLayout.mockRestore();
    warnings.mockRestore();
    errors.mockRestore();
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
    await act(async () => {
      const attaching = result.current.attachOffscreenCanvases(
        [], [], 1, 1, { color: '#000', width: 2 }
      );
      first!.onmessage?.({ data: { id: first!.last().id, ok: true } } as MessageEvent<ResidentEngineWorkerResponse>);
      expect(await attaching).toBe(true);
    });
    expect(first!.last()).toMatchObject({ type: 'completeLayout' });
    layOutHere();
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
