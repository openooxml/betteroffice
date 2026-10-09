import { GlobalRegistrator } from '@happy-dom/global-registrator';
import { afterAll, afterEach, beforeAll, expect, spyOn, test } from 'bun:test';
import type { LayoutComputation } from '@betteroffice/docx/editor';
import { LayoutSelectionGate, type ResidentMeasurementConfig } from '@betteroffice/docx/layout';
import type { Layout } from '@betteroffice/docx/layout/pagination';
import { proposalSetIdentity, type ResidentProposalReply, type YrsRenderEnv, type YrsSession } from '@betteroffice/docx/yrs';
import { isLayoutQueued, isSupersededLayout, sourceVersionOf } from '../internals/layoutProvenance';
import {
  deferWorkerOpenReplica,
  holdWorkerOpenDocument,
} from '../internals/workerOpenReplica';
import { DocxWorkerError } from '../internals/docxWorkerError';
import { registerWorkerProposalAuthority } from '../internals/workerProposalAuthority';
import type { FontRequirementsInWorker, WorkerLayoutComputation } from './useDisplayList';
import { SupersededPreviewError } from '../internals/supersededPreview';
import { useLayoutTriggers } from './useLayoutTriggers';

const ownsDom = !GlobalRegistrator.isRegistered;
if (ownsDom) GlobalRegistrator.register();
const { act, cleanup, renderHook } = await import('@testing-library/react');
const { useLayoutPipeline } = await import('./useLayoutPipeline');
const restoreFrames: Array<() => void> = [];

beforeAll(() => {
  if (!window.document.fonts) {
    Object.defineProperty(window.document, 'fonts', {
      value: { addEventListener: () => {}, removeEventListener: () => {} },
      configurable: true,
    });
  }
});

afterEach(() => {
  cleanup();
  for (const restore of restoreFrames.splice(0)) restore();
});

afterAll(async () => {
  if (ownsDom) await GlobalRegistrator.unregister();
});

interface WorkerPass {
  /** The document version the pass was asked for. */
  at: number;
  request: string;
  answer(computation?: WorkerLayoutComputation): void;
  fail(): void;
}

interface HookProps {
  session: YrsSession;
  renderEnv?: YrsRenderEnv;
  zoom?: number;
  viewport?: HTMLDivElement;
}

const MEASUREMENT: ResidentMeasurementConfig = {
  fontChains: { a: [0] },
  defaults: { fontSize: 22, fontFamily: 'Calibri' },
  compat: { noLeading: false, doNotExpandShiftReturn: false },
  authoritativeShaping: true,
};

function fakeDocument() {
  const doc = {
    version: 1,
    laidOutHere: [] as number[],
    workerAvailable: true,
    workerOwnsDocument: false,
    fontsReady: true,
    measurement: MEASUREMENT,
  };
  const session = {
    version: () => String(doc.version),
    layoutFontRequirementsJson: () => '[]',
    layoutDocumentWithRegionsRetainedJson: () => {
      doc.laidOutHere.push(doc.version);
      return JSON.stringify({ layout: { pages: [] }, notesConverged: true });
    },
    residentWorkerProbe: () => ({ layoutRevision: 1 }),
  } as unknown as YrsSession;
  return { doc, session };
}

/** A document whose version each test moves on; worker passes answer when the test says. */
async function opened({
  experimentalWorkerOpen = false,
  pendingReplica = false,
  ownsDocument = false,
  viewerSession = false,
  fontRequirementsInWorker = undefined as FontRequirementsInWorker | undefined,
} = {}) {
  let nextFrame = 0;
  const frames = new Map<number, FrameRequestCallback>();
  const requestFrame = spyOn(globalThis, 'requestAnimationFrame').mockImplementation((callback) => {
    const id = ++nextFrame;
    frames.set(id, callback);
    return id;
  });
  const cancelFrame = spyOn(globalThis, 'cancelAnimationFrame').mockImplementation((id) => {
    frames.delete(id);
  });
  restoreFrames.push(() => { requestFrame.mockRestore(); cancelFrame.mockRestore(); });
  const { doc, session } = fakeDocument();
  doc.workerOwnsDocument = ownsDocument;
  const replica = pendingReplica
    ? deferWorkerOpenReplica(session, () => new Promise(() => {}), () => {}, () => {})
    : null;
  const ensureReplica = replica ? spyOn(replica, 'ensure') : null;
  const release = viewerSession
    ? spyOn({ release: () => deferWorkerOpenReplica(session, async () => () => {}, () => {}, () => {}) }, 'release')
    : null;
  if (release) holdWorkerOpenDocument(session, release);
  const mainPreflight = spyOn(session, 'layoutFontRequirementsJson');
  restoreFrames.push(() => mainPreflight.mockRestore());
  const worker: WorkerPass[] = [];
  const errors: Error[] = [];
  const syncCoordinator = new LayoutSelectionGate();
  const hook = renderHook(({ session, renderEnv, zoom = 1, viewport }: HookProps) =>
    useLayoutPipeline({
      document: null,
      session,
      experimentalWorkerOpen,
      renderEnv: renderEnv ?? ({} as YrsRenderEnv),
      pageGap: 24,
      zoom,
      residentMeasurementConfig: () => (doc.fontsReady ? doc.measurement : null),
      deferLayoutPass: () => false,
      pagesContainerRef: { current: null },
      viewportLayoutRef: { current: viewport ?? null },
      syncCoordinator,
      getScrollContainer: () => null,
      onError: (error) => errors.push(error),
      fontRequirementsInWorker: fontRequirementsInWorker ??
        (viewerSession ? () => Promise.resolve('[]') : undefined),
      layoutInWorker: Object.assign((asked: YrsSession, request: string) =>
        doc.workerAvailable
          ? new Promise<WorkerLayoutComputation | null>((resolve) => {
              worker.push({
                at: Number(asked.version()),
                request,
                answer: (computation) =>
                  resolve(computation ?? {
                    layout: { pages: [] } as unknown as Layout, notesConverged: true,
                  }),
                fail: () => resolve(null),
              });
            })
          : null, {
        ownsDocument: (asked: YrsSession) => asked === session && doc.workerOwnsDocument,
        isViewerSession: (asked: YrsSession) => asked === session && viewerSession,
      }),
    }),
    { initialProps: { session } as HookProps }
  );
  const frame = () =>
    act(async () => {
      const pending = [...frames];
      for (const [id, callback] of pending) {
        if (!frames.delete(id)) continue;
        callback(performance.now());
      }
    });
  const answer = (index: number, computation?: WorkerLayoutComputation) =>
    act(async () => {
      worker[index]!.answer(computation);
    });
  const shown = () => sourceVersionOf(hook.result.current.layout);
  act(() => hook.result.current.runLayoutPipeline());
  if (viewerSession) await act(async () => {});
  await answer(0);
  expect(shown()).toBe('1');
  return {
    doc, session, worker, errors, hook, frame, answer, shown, replica, ensureReplica,
    mainPreflight, release,
  };
}

async function holdProposals(session: YrsSession) {
  const snapshot = { version: session.version(), previewVersion: 0, proposals: [] };
  Object.assign(session, {
    getProposals: () => snapshot,
    mirrorWorkerDocument: () => {},
  });
  const reply: ResidentProposalReply = {
    mirror: { version: snapshot.version, proposals: { previewVersion: 0, entries: [] } },
    result: { ok: true, snapshot },
    changedStories: [],
    geometry: {
      version: snapshot.version, previewVersion: 0,
      proposals: proposalSetIdentity(snapshot), targets: {}, hidden: [],
    },
    updates: [],
    stateVector: new Uint8Array(),
  };
  const authority = registerWorkerProposalAuthority(session, {
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
  }, async () => {
    throw new Error('unexpected main-thread toggle');
  });
  expect(authority.holdsWorkerState()).toBe(true);
}

test.each([false, true])('cached page totals are requested only with worker-open=%s', async (experimentalWorkerOpen) => {
  const h = await opened({ experimentalWorkerOpen });
  const request = JSON.parse(h.worker[0]!.request);
  const retainedRequest = JSON.parse(h.hook.result.current.getLayoutRequest()!);
  if (experimentalWorkerOpen) {
    expect(request.cachedPageTotals).toBe(true);
    expect(retainedRequest.cachedPageTotals).toBe(true);
  } else {
    expect(request).not.toHaveProperty('cachedPageTotals');
    expect(retainedRequest).not.toHaveProperty('cachedPageTotals');
  }
});

test('an identical pass replaces provisional pages while the replica is hydrating', async () => {
  const h = await opened({ experimentalWorkerOpen: true, pendingReplica: true, ownsDocument: true });
  h.replica!.start();
  expect(h.replica?.pending).toBe(true);
  expect(h.replica?.started).toBe(true);
  h.doc.version = 2;
  act(() => h.hook.result.current.scheduleLayout('remote'));
  await h.frame();
  let finish!: (computation: LayoutComputation | null) => void;
  const complete = new Promise<LayoutComputation | null>((resolve) => { finish = resolve; });
  const firstPages = { pages: [] } as unknown as Layout;
  await h.answer(1, { layout: firstPages, notesConverged: true, complete });
  expect(h.hook.result.current.layout).toBe(firstPages);

  act(() => h.hook.result.current.scheduleLayout('remote'));
  await h.frame();
  expect(h.worker.map((pass) => pass.at)).toEqual([1, 2, 2]);
  let finishRetry!: (computation: LayoutComputation | null) => void;
  const retryComplete = new Promise<LayoutComputation | null>((resolve) => { finishRetry = resolve; });
  const retryPages = { pages: [] } as unknown as Layout;
  await h.answer(2, { layout: retryPages, notesConverged: true, complete: retryComplete });
  await act(async () => finish({ layout: { pages: [] } as unknown as Layout, notesConverged: true }));
  expect(h.hook.result.current.layout).toBe(retryPages);
  const full = { pages: [] } as unknown as Layout;
  await act(async () => finishRetry({ layout: full, notesConverged: true }));
  expect(h.hook.result.current.layout).toBe(full);
  expect(isLayoutQueued(h.session)).toBe(false);
  act(() => h.hook.result.current.runLayoutPipeline());
  await h.frame();
  expect(h.worker).toHaveLength(3);
  expect(h.hook.result.current.layout).toBe(full);
  expect(h.doc.laidOutHere).toEqual([]);
  expect(h.ensureReplica).not.toHaveBeenCalled();
  expect(h.errors).toEqual([]);
});

test('an identical pass retries a null worker reply while the replica is hydrating', async () => {
  const h = await opened({ experimentalWorkerOpen: true, pendingReplica: true, ownsDocument: true });
  act(() => h.hook.result.current.runLayoutPipeline());
  expect(h.worker).toHaveLength(2);
  h.replica!.start();
  expect(h.replica?.pending).toBe(true);
  expect(h.replica?.started).toBe(true);
  act(() => h.hook.result.current.runLayoutPipeline());
  expect(isLayoutQueued(h.session)).toBe(true);
  await act(async () => h.worker[1]!.fail());
  expect(h.doc.laidOutHere).toEqual([]);
  await h.frame();
  expect(h.worker.map((pass) => pass.at)).toEqual([1, 1, 1]);
  await h.answer(2);
  expect(isLayoutQueued(h.session)).toBe(false);
  expect(h.doc.laidOutHere).toEqual([]);
  expect(h.ensureReplica).not.toHaveBeenCalled();
  expect(h.errors).toEqual([]);
});

test('worker ownership lost before preflight requires another layout after it recovers', async () => {
  let recovering = false;
  let recover = () => {};
  const h = await opened({
    experimentalWorkerOpen: true,
    pendingReplica: true,
    ownsDocument: true,
    fontRequirementsInWorker: () => {
      if (!recovering) return null;
      recover();
      return Promise.resolve('[]');
    },
  });
  h.replica!.start();
  h.doc.workerOwnsDocument = false;
  recovering = true;
  recover = () => { h.doc.workerOwnsDocument = true; };
  act(() => h.hook.result.current.runLayoutPipeline());
  await h.frame();
  expect(h.doc.workerOwnsDocument).toBe(true);
  expect(h.worker.map((pass) => pass.at)).toEqual([1, 1]);
  await h.answer(1);
  expect(isLayoutQueued(h.session)).toBe(false);
  expect(h.doc.laidOutHere).toEqual([]);
  expect(h.ensureReplica).not.toHaveBeenCalled();
  expect(h.errors).toEqual([]);
});

test('a zoom change updates the viewport while the replica is hydrating', async () => {
  const h = await opened({ experimentalWorkerOpen: true, pendingReplica: true, ownsDocument: true });
  h.replica!.start();
  const viewport = document.createElement('div');
  h.hook.rerender({ session: h.session, zoom: 2, viewport });
  act(() => h.hook.result.current.runLayoutPipeline());
  await h.frame();
  expect(h.worker).toHaveLength(2);
  await h.answer(1);
  expect(viewport.style.minHeight).not.toBe('');
  expect(viewport.style.marginBottom).toBe(viewport.style.minHeight);
  expect(isLayoutQueued(h.session)).toBe(false);
  expect(h.doc.laidOutHere).toEqual([]);
  expect(h.errors).toEqual([]);
});

test('a deferred host pass does not force a new session to hydrate its replica', async () => {
  const h = await opened({ experimentalWorkerOpen: true, pendingReplica: true });
  h.doc.fontsReady = false;
  act(() => h.hook.result.current.runLayoutPipeline({ onHost: true }));
  expect(h.ensureReplica).toHaveBeenCalledTimes(1);
  const next = fakeDocument();
  const replica = deferWorkerOpenReplica(next.session, () => new Promise(() => {}), () => {}, () => {});
  const ensureReplica = spyOn(replica, 'ensure');
  h.hook.rerender({ session: next.session });
  h.doc.fontsReady = true;
  act(() => h.hook.result.current.runLayoutPipeline());
  await h.frame();
  expect(ensureReplica).not.toHaveBeenCalled();
  expect(replica.pending).toBe(true);
  expect(replica.started).toBe(false);
  expect(h.worker).toHaveLength(2);
  await h.answer(1);
  expect(h.hook.result.current.layout).not.toBeNull();
  expect(isLayoutQueued(next.session)).toBe(false);
  expect(next.doc.laidOutHere).toEqual([]);
  expect(h.errors).toEqual([]);
});

test('a superseded null completion requests the worker while it holds proposals', async () => {
  const h = await opened({ experimentalWorkerOpen: true, pendingReplica: true });
  await holdProposals(h.session);
  act(() => h.hook.result.current.scheduleLayout('remote'));
  await h.frame();
  let finish!: (computation: LayoutComputation | null) => void;
  const complete = new Promise<LayoutComputation | null>((resolve) => { finish = resolve; });
  const firstPages = { pages: [] } as unknown as Layout;
  await h.answer(1, { layout: firstPages, notesConverged: true, complete });
  await act(async () => finish(null));
  expect(h.hook.result.current.layout).toBe(firstPages);
  expect(h.doc.laidOutHere).toEqual([]);
  expect(h.ensureReplica).not.toHaveBeenCalled();
  expect(isLayoutQueued(h.session)).toBe(true);
  await h.frame();
  expect(h.worker).toHaveLength(3);
  await h.answer(2);
  expect(isLayoutQueued(h.session)).toBe(false);
  expect(h.doc.laidOutHere).toEqual([]);
  expect(h.ensureReplica).not.toHaveBeenCalled();
  expect(h.errors).toEqual([]);
});

test.each(['computation', 'completion'])('a viewer whose worker holds proposals queues a worker pass for every null %s', async (result) => {
  const h = await opened({ experimentalWorkerOpen: true, ownsDocument: true });
  await holdProposals(h.session);
  act(() => h.hook.result.current.scheduleLayout('remote'));
  await h.frame();
  for (const index of [1, 2]) {
    if (result === 'computation') {
      await act(async () => h.worker[index]!.fail());
    } else {
      let finish!: (computation: LayoutComputation | null) => void;
      const complete = new Promise<LayoutComputation | null>((resolve) => { finish = resolve; });
      await h.answer(index, { layout: { pages: [] } as unknown as Layout, notesConverged: true, complete });
      await act(async () => finish(null));
    }
    expect(isLayoutQueued(h.session)).toBe(true);
    await h.frame();
    expect(h.worker).toHaveLength(index + 2);
  }
  await h.answer(3);
  expect(isLayoutQueued(h.session)).toBe(false);
  expect(h.worker.map((pass) => pass.at)).toEqual([1, 1, 1, 1]);
  expect(h.doc.laidOutHere).toEqual([]);
  expect(h.errors).toEqual([]);
});

test.each([false, true])('a null completion without held proposals keeps the host path with worker-open=%s', async (experimentalWorkerOpen) => {
  const h = await opened({ experimentalWorkerOpen, pendingReplica: true });
  act(() => h.hook.result.current.scheduleLayout('remote'));
  await h.frame();
  let finish!: (computation: LayoutComputation | null) => void;
  const complete = new Promise<LayoutComputation | null>((resolve) => { finish = resolve; });
  const firstPages = { pages: [] } as unknown as Layout;
  await h.answer(1, { layout: firstPages, notesConverged: true, complete });
  await act(async () => finish(null));
  expect(h.hook.result.current.layout).not.toBe(firstPages);
  expect(h.doc.laidOutHere).toEqual([1]);
  expect(h.ensureReplica).toHaveBeenCalledTimes(experimentalWorkerOpen ? 1 : 0);
  expect(isLayoutQueued(h.session)).toBe(false);
  await h.frame();
  expect(h.worker).toHaveLength(2);
  expect(h.errors).toEqual([]);
});

test('a live worker retries a null completion once per source version without a pending replica', async () => {
  const h = await opened({ experimentalWorkerOpen: true, ownsDocument: true });
  h.doc.version = 2;
  act(() => h.hook.result.current.scheduleLayout('remote'));
  await h.frame();
  let finish!: (computation: LayoutComputation | null) => void;
  const complete = new Promise<LayoutComputation | null>((resolve) => { finish = resolve; });
  const firstPages = { pages: [] } as unknown as Layout;
  await h.answer(1, { layout: firstPages, notesConverged: true, complete });
  expect(h.hook.result.current.layout).toBe(firstPages);
  expect(h.shown()).toBe('2');
  expect(h.replica).toBeNull();

  await act(async () => finish(null));
  expect(h.hook.result.current.layout).toBe(firstPages);
  expect(h.doc.laidOutHere).toEqual([]);
  expect(isLayoutQueued(h.session)).toBe(true);
  await h.frame();
  expect(h.worker.map((pass) => pass.at)).toEqual([1, 2, 2]);
  let finishRetry!: (computation: LayoutComputation | null) => void;
  const retryComplete = new Promise<LayoutComputation | null>((resolve) => { finishRetry = resolve; });
  const retryPages = { pages: [] } as unknown as Layout;
  await h.answer(2, { layout: retryPages, notesConverged: true, complete: retryComplete });
  expect(h.hook.result.current.layout).toBe(retryPages);
  await act(async () => finishRetry(null));
  expect(h.doc.laidOutHere).toEqual([2]);
  expect(h.hook.result.current.layout).not.toBe(retryPages);
  expect(h.shown()).toBe('2');
  expect(isLayoutQueued(h.session)).toBe(false);
  await h.frame();
  expect(h.worker).toHaveLength(3);
  expect(h.doc.laidOutHere).toEqual([2]);

  h.doc.version = 3;
  act(() => h.hook.result.current.scheduleLayout('remote'));
  await h.frame();
  let finishNext!: (computation: LayoutComputation | null) => void;
  const nextComplete = new Promise<LayoutComputation | null>((resolve) => { finishNext = resolve; });
  await h.answer(3, { layout: { pages: [] } as unknown as Layout, notesConverged: true, complete: nextComplete });
  await act(async () => finishNext(null));
  expect(h.doc.laidOutHere).toEqual([2]);
  expect(isLayoutQueued(h.session)).toBe(true);
  await h.frame();
  expect(h.worker.map((pass) => pass.at)).toEqual([1, 2, 2, 3, 3]);
  await h.answer(4);
  expect(h.shown()).toBe('3');
  const next = fakeDocument();
  next.doc.version = 3;
  h.hook.rerender({ session: next.session });
  act(() => h.hook.result.current.runLayoutPipeline());
  let finishSession!: (computation: LayoutComputation | null) => void;
  const sessionComplete = new Promise<LayoutComputation | null>((resolve) => { finishSession = resolve; });
  await h.answer(5, { layout: { pages: [] } as unknown as Layout, notesConverged: true, complete: sessionComplete });
  await act(async () => finishSession(null));
  expect(next.doc.laidOutHere).toEqual([]);
  expect(isLayoutQueued(next.session)).toBe(true);
  await h.frame();
  expect(h.worker.map((pass) => pass.at)).toEqual([1, 2, 2, 3, 3, 3, 3]);
  await h.answer(6);
  expect(isLayoutQueued(next.session)).toBe(false);
  expect(h.doc.laidOutHere).toEqual([2]);
  expect(next.doc.laidOutHere).toEqual([]);
  expect(h.errors).toEqual([]);
});

test.each(['null', 'stale'])('a worker that loses ownership keeps the host path for a %s completion', async (outcome) => {
  const h = await opened({ experimentalWorkerOpen: true, ownsDocument: true });
  h.doc.version = 2;
  act(() => h.hook.result.current.scheduleLayout('remote'));
  await h.frame();
  let finish!: (computation: LayoutComputation | null) => void;
  const complete = new Promise<LayoutComputation | null>((resolve) => { finish = resolve; });
  const firstPages = { pages: [] } as unknown as Layout;
  await h.answer(1, { layout: firstPages, notesConverged: true, complete });
  expect(h.hook.result.current.layout).toBe(firstPages);
  h.doc.workerOwnsDocument = false;
  if (outcome === 'stale') h.doc.version = 3;
  await act(async () => finish(outcome === 'null'
    ? null
    : { layout: { pages: [] } as unknown as Layout, notesConverged: true }));
  expect(h.doc.laidOutHere).toEqual([h.doc.version]);
  expect(h.shown()).toBe(String(h.doc.version));
  expect(h.hook.result.current.layout).not.toBe(firstPages);
  expect(isLayoutQueued(h.session)).toBe(false);
  await h.frame();
  expect(h.worker.map((pass) => pass.at)).toEqual([1, 2]);
  expect(h.errors).toEqual([]);
});

test('worker ownership keeps the host path for a null first result', async () => {
  const h = await opened({ experimentalWorkerOpen: true, ownsDocument: true });
  h.doc.version = 2;
  act(() => h.hook.result.current.scheduleLayout('remote'));
  await h.frame();
  await act(async () => h.worker[1]!.fail());
  expect(h.doc.laidOutHere).toEqual([2]);
  expect(h.shown()).toBe('2');
  expect(isLayoutQueued(h.session)).toBe(false);
  await h.frame();
  expect(h.worker).toHaveLength(2);
  expect(h.errors).toEqual([]);
});

test('a viewer fails a null worker layout without ensuring or laying out on the host', async () => {
  const h = await opened({ experimentalWorkerOpen: true, viewerSession: true });
  act(() => h.hook.result.current.runLayoutPipeline({ onHost: true }));
  await act(async () => {});
  await act(async () => h.worker[1]!.fail());
  expect(h.doc.laidOutHere).toEqual([]);
  expect(h.mainPreflight).not.toHaveBeenCalled();
  expect(h.release).not.toHaveBeenCalled();
  expect(h.errors).toHaveLength(1);
  expect(h.errors[0]).toBeInstanceOf(DocxWorkerError);
  expect((h.errors[0] as DocxWorkerError).stage).toBe('layout');
  h.hook.unmount();
});

test('a viewer reports an unavailable worker instead of honoring a host-layout request', async () => {
  const h = await opened({ experimentalWorkerOpen: true, viewerSession: true });
  h.doc.workerAvailable = false;
  act(() => h.hook.result.current.runLayoutPipeline({ onHost: true }));
  await act(async () => {});
  expect(h.worker).toHaveLength(1);
  expect(h.doc.laidOutHere).toEqual([]);
  expect(h.errors[0]).toBeInstanceOf(DocxWorkerError);
  expect(h.mainPreflight).not.toHaveBeenCalled();
  expect(h.release).not.toHaveBeenCalled();
  h.hook.unmount();
});

test.each(['first', 'full'])('a viewer requeues a stale %s worker layout without a host pass', async (stage) => {
  const h = await opened({ experimentalWorkerOpen: true, viewerSession: true });
  h.doc.version = 2;
  act(() => h.hook.result.current.scheduleLayout('local'));
  await h.frame();
  let finish!: (computation: LayoutComputation | null) => void;
  const complete = new Promise<LayoutComputation | null>((resolve) => { finish = resolve; });
  if (stage === 'full') await h.answer(1, {
    layout: { pages: [] } as unknown as Layout, notesConverged: true, complete,
  });
  h.doc.version = 3;
  if (stage === 'first') await h.answer(1);
  else await act(async () => finish({ layout: { pages: [] } as unknown as Layout, notesConverged: true }));
  await h.frame();
  await h.answer(2);
  expect(h.shown()).toBe('3');
  expect(h.doc.laidOutHere).toEqual([]);
  expect(h.mainPreflight).not.toHaveBeenCalled();
  expect(h.release).not.toHaveBeenCalled();
  expect(h.errors).toEqual([]);
  h.hook.unmount();
});

test.each(['missing', 'rejected'])('viewer font requirements that are %s never invoke the main preflight', async (kind) => {
  let fail = false;
  const cause = new Error('font requirements failed');
  const h = await opened({
    experimentalWorkerOpen: true, viewerSession: true,
    fontRequirementsInWorker: () => fail
      ? kind === 'missing' ? null : Promise.reject(cause)
      : Promise.resolve('[]'),
  });
  fail = true;
  act(() => h.hook.result.current.runLayoutPipeline());
  await act(async () => {});
  expect(h.errors).toHaveLength(1);
  expect(h.errors[0]).toBeInstanceOf(DocxWorkerError);
  if (kind === 'rejected') expect(h.errors[0]!.cause).toBe(cause);
  expect(h.doc.laidOutHere).toEqual([]);
  expect(h.mainPreflight).not.toHaveBeenCalled();
  expect(h.release).not.toHaveBeenCalled();
  h.hook.unmount();
});

test('a viewer exhausts null completion retries without a main layout', async () => {
  const h = await opened({ experimentalWorkerOpen: true, viewerSession: true, ownsDocument: true });
  for (let index = 1; index <= 2; index += 1) {
    if (index === 1) act(() => h.hook.result.current.runLayoutPipeline());
    else await h.frame();
    await act(async () => {});
    await h.answer(index, {
      layout: { pages: [] } as unknown as Layout,
      notesConverged: true,
      complete: Promise.resolve(null),
    });
  }
  expect(h.errors).toHaveLength(1);
  expect(h.errors[0]).toBeInstanceOf(DocxWorkerError);
  expect(h.doc.laidOutHere).toEqual([]);
  expect(h.mainPreflight).not.toHaveBeenCalled();
  expect(h.release).not.toHaveBeenCalled();
  h.hook.unmount();
});

test('worker ownership is ignored for null completions with worker-open off', async () => {
  const h = await opened({ ownsDocument: true });
  act(() => h.hook.result.current.scheduleLayout('remote'));
  await h.frame();
  let finish!: (computation: LayoutComputation | null) => void;
  const complete = new Promise<LayoutComputation | null>((resolve) => { finish = resolve; });
  const firstPages = { pages: [] } as unknown as Layout;
  await h.answer(1, { layout: firstPages, notesConverged: true, complete });
  expect(h.hook.result.current.layout).toBe(firstPages);
  await act(async () => finish(null));
  expect(h.doc.laidOutHere).toEqual([1]);
  expect(h.hook.result.current.layout).not.toBe(firstPages);
  expect(isLayoutQueued(h.session)).toBe(false);
  await h.frame();
  expect(h.worker).toHaveLength(2);
  expect(h.errors).toEqual([]);
});

test.each(['first', 'full'])('a live worker retries a stale %s result without a pending replica', async (stage) => {
  const h = await opened({ experimentalWorkerOpen: true, ownsDocument: true });
  h.doc.version = 2;
  act(() => h.hook.result.current.scheduleLayout('remote'));
  await h.frame();
  expect(isLayoutQueued(h.session)).toBe(false);
  let finish!: (computation: LayoutComputation | null) => void;
  const complete = new Promise<LayoutComputation | null>((resolve) => { finish = resolve; });
  const firstPages = { pages: [] } as unknown as Layout;
  if (stage === 'full') {
    await h.answer(1, { layout: firstPages, notesConverged: true, complete });
    expect(h.hook.result.current.layout).toBe(firstPages);
    expect(h.shown()).toBe('2');
  }
  h.doc.version = 3;
  if (stage === 'first') await h.answer(1, { layout: firstPages, notesConverged: true });
  else await act(async () => finish({ layout: { pages: [] } as unknown as Layout, notesConverged: true }));
  expect(h.hook.result.current.layout).toBe(firstPages);
  expect(h.shown()).toBe('2');
  expect(h.doc.laidOutHere).toEqual([]);
  expect(h.replica).toBeNull();
  expect(isLayoutQueued(h.session)).toBe(true);
  await h.frame();
  expect(h.worker.map((pass) => pass.at)).toEqual([1, 2, 3]);
  await h.answer(2);
  expect(h.shown()).toBe('3');
  expect(isSupersededLayout(h.hook.result.current.layout)).toBe(false);
  expect(isLayoutQueued(h.session)).toBe(false);
  await h.frame();
  expect(h.worker).toHaveLength(3);
  expect(h.doc.laidOutHere).toEqual([]);
  expect(h.errors).toEqual([]);
});

test.each(['first', 'full'])('worker ownership is ignored for stale %s results with worker-open off', async (stage) => {
  const h = await opened({ ownsDocument: true });
  h.doc.version = 2;
  act(() => h.hook.result.current.scheduleLayout('remote'));
  await h.frame();
  let finish!: (computation: LayoutComputation | null) => void;
  const complete = new Promise<LayoutComputation | null>((resolve) => { finish = resolve; });
  if (stage === 'full') {
    await h.answer(1, { layout: { pages: [] } as unknown as Layout, notesConverged: true, complete });
  }
  h.doc.version = 3;
  if (stage === 'first') await h.answer(1);
  else await act(async () => finish({ layout: { pages: [] } as unknown as Layout, notesConverged: true }));
  expect(h.doc.laidOutHere).toEqual([3]);
  expect(h.shown()).toBe('3');
  expect(isSupersededLayout(h.hook.result.current.layout)).toBe(false);
  expect(isLayoutQueued(h.session)).toBe(false);
  await h.frame();
  expect(h.worker.map((pass) => pass.at)).toEqual([1, 2]);
  expect(h.errors).toEqual([]);
});

test('worker proposals override deferred host passes, local changes and unavailable passes', async () => {
  const h = await opened({ experimentalWorkerOpen: true, pendingReplica: true });
  await holdProposals(h.session);
  h.doc.fontsReady = false;
  act(() => h.hook.result.current.runLayoutPipeline({ onHost: true }));
  h.doc.version = 2;
  h.doc.fontsReady = true;
  h.doc.workerAvailable = false;
  act(() => {
    h.hook.result.current.scheduleLayout('remote');
    h.hook.result.current.scheduleLayout('local');
  });
  await h.frame();
  expect(h.doc.laidOutHere).toEqual([]);
  expect(h.ensureReplica).not.toHaveBeenCalled();
  expect(isLayoutQueued(h.session)).toBe(true);
  h.doc.workerAvailable = true;
  await h.frame();
  expect(h.worker.map((pass) => pass.at)).toEqual([1, 2]);
  act(() => h.hook.result.current.runLayoutPipeline({ onHost: true }));
  expect(h.worker).toHaveLength(2);
  await h.answer(1);
  await h.frame();
  expect(h.worker.map((pass) => pass.at)).toEqual([1, 2, 2]);
  await h.answer(2);
  expect(h.shown()).toBe('2');
  expect(h.doc.laidOutHere).toEqual([]);
  expect(h.ensureReplica).not.toHaveBeenCalled();
  expect(h.replica?.pending).toBe(true);
  expect(h.errors).toEqual([]);
});

test('host batches and remote updates lay out in the worker, local edits here', async () => {
  const { doc, worker, errors, hook, frame, answer, shown } = await opened();

  doc.version = 2;
  act(() => hook.result.current.scheduleLayout('local', true));
  await frame();
  expect(worker.map((pass) => pass.at)).toEqual([1, 2]);
  await answer(1);
  expect(shown()).toBe('2');

  doc.version = 3;
  act(() => hook.result.current.scheduleLayout('local'));
  await frame();
  expect(doc.laidOutHere).toEqual([3]);
  expect(shown()).toBe('3');

  doc.version = 4;
  act(() => hook.result.current.scheduleLayout('remote'));
  await frame();
  expect(worker.map((pass) => pass.at)).toEqual([1, 2, 4]);
  await answer(2);
  expect(shown()).toBe('4');
  expect(doc.laidOutHere).toEqual([3]);
  expect(errors).toEqual([]);
});

test('a passive render-env trigger publishes remote and a coalesced local edit publishes local', async () => {
  const h = await opened();
  const triggers = renderHook(({ renderEnv }) => useLayoutTriggers({
    runLayoutPipeline: h.hook.result.current.runLayoutPipeline,
    updateSelectionOverlay: () => {},
    renderEnv,
  }), { initialProps: { renderEnv: {} as YrsRenderEnv } });

  const preview = { revisionPreview: { a: 'accepted' } } as YrsRenderEnv;
  act(() => {
    h.hook.rerender({ session: h.session, renderEnv: preview });
    triggers.rerender({ renderEnv: preview });
  });
  expect(h.worker).toHaveLength(2);
  await h.answer(1);
  expect(h.hook.result.current.layoutUpdateOrigin).toBe('remote');

  h.doc.version = 2;
  const next = {} as YrsRenderEnv;
  act(() => {
    h.hook.result.current.scheduleLayout('local');
    h.hook.rerender({ session: h.session, renderEnv: next });
    triggers.rerender({ renderEnv: next });
  });
  await h.frame();
  expect(h.doc.laidOutHere).toEqual([2]);
  expect(h.hook.result.current.layoutUpdateOrigin).toBe('local');
  expect(h.errors).toEqual([]);
});

test('a preview change that only adds font chains lays out in the worker', async () => {
  const { doc, session, worker, errors, hook, frame, answer } = await opened();
  doc.version = 2;
  act(() => hook.result.current.scheduleLayout('local'));
  await frame();
  expect(doc.laidOutHere).toEqual([2]);

  doc.measurement = { ...MEASUREMENT, fontChains: { a: [0], b: [1] } };
  hook.rerender({ session, renderEnv: { revisionPreview: { a: 'accepted' } } as YrsRenderEnv });
  act(() => hook.result.current.runLayoutPipeline());
  await frame();
  expect(worker.map((pass) => pass.at)).toEqual([1, 2]);
  await answer(1);

  // A chain that measures differently keeps the pass here.
  doc.measurement = { ...MEASUREMENT, fontChains: { a: [2], b: [1] } };
  hook.rerender({ session, renderEnv: {} as YrsRenderEnv });
  act(() => hook.result.current.runLayoutPipeline());
  await frame();
  expect(doc.laidOutHere).toEqual([2, 2]);

  // So does a new font chain without a preview change.
  doc.measurement = { ...MEASUREMENT, fontChains: { a: [2], b: [1], c: [3] } };
  act(() => hook.result.current.runLayoutPipeline());
  await frame();
  expect(doc.laidOutHere).toEqual([2, 2, 2]);
  expect(worker).toHaveLength(2);
  expect(errors).toEqual([]);
});

test.each([false, true])('a local edit in the same frame as a host batch keeps the pass here with worker-open=%s', async (experimentalWorkerOpen) => {
  const { doc, worker, hook, frame, shown } = await opened({ experimentalWorkerOpen, ownsDocument: true });
  doc.version = 2;
  act(() => {
    hook.result.current.scheduleLayout('local', true);
    hook.result.current.scheduleLayout('local');
  });
  await frame();
  expect(worker).toHaveLength(1);
  expect(doc.laidOutHere).toEqual([2]);
  expect(shown()).toBe('2');
});

test('a worker layout of a host batch never replaces a newer local edit', async () => {
  const { doc, worker, errors, hook, frame, answer, shown } = await opened();

  doc.version = 2;
  act(() => hook.result.current.scheduleLayout('local', true));
  await frame();
  doc.version = 3;
  act(() => hook.result.current.scheduleLayout('local'));
  await frame();
  expect(shown()).toBe('3');
  doc.version = 4;
  act(() => hook.result.current.scheduleLayout('local', true));
  await frame();
  expect(worker.map((pass) => pass.at)).toEqual([1, 2]);

  await answer(1);
  expect(shown()).toBe('3');
  await frame();
  expect(worker.map((pass) => pass.at)).toEqual([1, 2, 4]);
  await answer(2);
  expect(shown()).toBe('4');
  expect(doc.laidOutHere).toEqual([3]);
  expect(errors).toEqual([]);
});

test('updates that land while the worker lays out queue one pass for the latest state', async () => {
  const { doc, worker, errors, hook, frame, answer, shown } = await opened();

  doc.version = 2;
  act(() => hook.result.current.scheduleLayout('local', true));
  await frame();
  for (const version of [3, 4]) {
    doc.version = version;
    act(() => hook.result.current.scheduleLayout(version === 3 ? 'remote' : 'local', true));
    await frame();
  }
  expect(worker.map((pass) => pass.at)).toEqual([1, 2]);

  await answer(1);
  expect(shown()).toBe('2');
  expect(isSupersededLayout(hook.result.current.layout)).toBe(true);
  await frame();
  expect(worker.map((pass) => pass.at)).toEqual([1, 2, 4]);
  await answer(2);
  expect(shown()).toBe('4');
  expect(isSupersededLayout(hook.result.current.layout)).toBe(false);
  expect(doc.laidOutHere).toEqual([]);
  expect(errors).toEqual([]);
});

test('a stale worker reply queues one more worker pass while the replica is pending', async () => {
  const h = await opened({ experimentalWorkerOpen: true, pendingReplica: true });
  h.doc.version = 2;
  act(() => h.hook.result.current.scheduleLayout('remote'));
  await h.frame();
  expect(isLayoutQueued(h.session)).toBe(false);
  h.doc.version = 3;

  await h.answer(1);
  expect(h.shown()).toBe('2');
  expect(isSupersededLayout(h.hook.result.current.layout)).toBe(true);
  expect(isLayoutQueued(h.session)).toBe(true);
  expect(h.doc.laidOutHere).toEqual([]);
  expect(h.ensureReplica).not.toHaveBeenCalled();
  expect(h.replica?.pending).toBe(true);
  expect(h.replica?.started).toBe(false);

  await h.frame();
  expect(h.worker.map((pass) => pass.at)).toEqual([1, 2, 3]);
  expect(isLayoutQueued(h.session)).toBe(false);
  await h.answer(2);
  expect(h.shown()).toBe('3');
  expect(isSupersededLayout(h.hook.result.current.layout)).toBe(false);
  await h.frame();
  expect(h.worker).toHaveLength(3);
  expect(h.doc.laidOutHere).toEqual([]);
  expect(h.ensureReplica).not.toHaveBeenCalled();
  expect(h.replica?.pending).toBe(true);
  expect(h.replica?.started).toBe(false);
  expect(h.errors).toEqual([]);
  h.hook.unmount();
});

test('a stale worker reply lays out here without a pending replica', async () => {
  const { doc, session, worker, errors, hook, frame, answer, shown } = await opened({
    experimentalWorkerOpen: true,
  });
  doc.version = 2;
  act(() => hook.result.current.scheduleLayout('remote'));
  await frame();
  doc.version = 3;

  await answer(1);
  expect(shown()).toBe('3');
  expect(isSupersededLayout(hook.result.current.layout)).toBe(false);
  expect(isLayoutQueued(session)).toBe(false);
  await frame();
  expect(worker.map((pass) => pass.at)).toEqual([1, 2]);
  expect(doc.laidOutHere).toEqual([3]);
  expect(errors).toEqual([]);
  hook.unmount();
});

for (const pendingReplica of [true, false]) {
  const name = pendingReplica
    ? 'a stale full worker layout queues another pass while the replica is pending'
    : 'a stale full worker layout lays out here without a pending replica';
  test(name, async () => {
    const h = await opened({ experimentalWorkerOpen: true, pendingReplica });
    h.doc.version = 2;
    act(() => h.hook.result.current.scheduleLayout('remote'));
    await h.frame();
    let finish!: (computation: LayoutComputation | null) => void;
    const complete = new Promise<LayoutComputation | null>((resolve) => { finish = resolve; });
    const firstPages = { pages: [] } as unknown as Layout;
    await h.answer(1, { layout: firstPages, notesConverged: true, complete });
    expect(h.shown()).toBe('2');
    expect(h.hook.result.current.layout).toBe(firstPages);
    expect(isLayoutQueued(h.session)).toBe(false);
    h.doc.version = 3;
    const full = { pages: [] } as unknown as Layout;
    await act(async () => {
      finish({ layout: full, notesConverged: true });
    });

    if (pendingReplica) {
      expect(h.hook.result.current.layout).toBe(firstPages);
      expect(h.shown()).toBe('2');
      expect(isLayoutQueued(h.session)).toBe(true);
      expect(h.doc.laidOutHere).toEqual([]);
      expect(h.ensureReplica).not.toHaveBeenCalled();
      expect(h.replica?.pending).toBe(true);
      expect(h.replica?.started).toBe(false);
      await h.frame();
      expect(h.worker.map((pass) => pass.at)).toEqual([1, 2, 3]);
      expect(isLayoutQueued(h.session)).toBe(false);
      await h.answer(2);
      expect(h.shown()).toBe('3');
      expect(isSupersededLayout(h.hook.result.current.layout)).toBe(false);
      await h.frame();
      expect(h.worker).toHaveLength(3);
      expect(h.doc.laidOutHere).toEqual([]);
      expect(h.ensureReplica).not.toHaveBeenCalled();
      expect(h.replica?.pending).toBe(true);
      expect(h.replica?.started).toBe(false);
    } else {
      expect(h.hook.result.current.layout).not.toBe(full);
      expect(h.shown()).toBe('3');
      expect(isSupersededLayout(h.hook.result.current.layout)).toBe(false);
      expect(isLayoutQueued(h.session)).toBe(false);
      await h.frame();
      expect(h.worker.map((pass) => pass.at)).toEqual([1, 2]);
      expect(h.doc.laidOutHere).toEqual([3]);
    }
    expect(h.errors).toEqual([]);
    h.hook.unmount();
  });
}

test('a pass no change asked to run here waits for the worker pass in flight', async () => {
  const { doc, session, worker, errors, hook, frame, answer, shown } = await opened();

  doc.version = 2;
  act(() => hook.result.current.scheduleLayout('local', true));
  await frame();
  expect(isLayoutQueued(session)).toBe(false);
  act(() => hook.result.current.runLayoutPipeline());
  await frame();
  expect(worker.map((pass) => pass.at)).toEqual([1, 2]);
  // Nothing the display shows settles a wait until the queued pass runs.
  expect(isLayoutQueued(session)).toBe(true);

  await answer(1);
  expect(shown()).toBe('2');
  // The queued pass may change only the revision preview, so this one settles no wait.
  expect(hook.result.current.layoutUpdateOrigin).toBe('local');
  expect(isSupersededLayout(hook.result.current.layout)).toBe(true);
  await frame();
  expect(worker.map((pass) => pass.at)).toEqual([1, 2, 2]);
  expect(isLayoutQueued(session)).toBe(false);
  await answer(2);
  expect(hook.result.current.layoutUpdateOrigin).toBe('remote');
  expect(isSupersededLayout(hook.result.current.layout)).toBe(false);
  expect(doc.laidOutHere).toEqual([]);
  expect(errors).toEqual([]);
});

test('a queued pass that waits for fonts holds settles until it lays out', async () => {
  const { doc, session, worker, errors, hook, frame, answer } = await opened();

  doc.version = 2;
  act(() => hook.result.current.scheduleLayout('local', true));
  await frame();
  act(() => hook.result.current.runLayoutPipeline());
  await frame();
  doc.fontsReady = false;
  await answer(1);
  await frame();
  expect(worker.map((pass) => pass.at)).toEqual([1, 2]);
  expect(isLayoutQueued(session)).toBe(true);

  doc.fontsReady = true;
  act(() => hook.result.current.runLayoutPipeline());
  await frame();
  expect(worker.map((pass) => pass.at)).toEqual([1, 2, 2]);
  expect(isLayoutQueued(session)).toBe(false);
  expect(errors).toEqual([]);
});

test('a worker pass that fails with a pass queued behind it leaves the layout to that pass', async () => {
  const { doc, worker, errors, hook, frame, shown } = await opened();

  doc.version = 2;
  act(() => hook.result.current.scheduleLayout('local', true));
  await frame();
  act(() => hook.result.current.runLayoutPipeline());
  await frame();
  await act(async () => {
    worker[1]!.fail();
  });
  await frame();
  // Only the queued pass lays out, here: the failed one left it the layout.
  expect(worker.map((pass) => pass.at)).toEqual([1, 2]);
  expect(doc.laidOutHere).toEqual([2]);
  expect(shown()).toBe('2');
  expect(isSupersededLayout(hook.result.current.layout)).toBe(false);
  expect(errors).toEqual([]);
});

test('a superseded font preflight drops its pass without an error', async () => {
  let superseded = false;
  const h = await opened({
    experimentalWorkerOpen: true,
    pendingReplica: true,
    fontRequirementsInWorker: () => (superseded ? Promise.reject(new SupersededPreviewError()) : null),
  });
  superseded = true;
  h.doc.version = 2;
  act(() => h.hook.result.current.scheduleLayout('remote'));
  await h.frame();
  expect(h.errors).toEqual([]);
  expect(h.worker).toHaveLength(1);
  expect(h.doc.laidOutHere).toEqual([]);
  expect(h.shown()).toBe('1');
  expect(isLayoutQueued(h.session)).toBe(false);

  superseded = false;
  h.doc.version = 3;
  act(() => h.hook.result.current.scheduleLayout('remote'));
  await h.frame();
  await h.answer(1);
  expect(h.worker.map((pass) => pass.at)).toEqual([1, 3]);
  expect(h.shown()).toBe('3');
  expect(h.errors).toEqual([]);
  h.hook.unmount();
});

test('a pass queued behind the worker never runs after unmount', async () => {
  const { doc, worker, errors, hook, frame, answer } = await opened();
  doc.version = 2;
  act(() => hook.result.current.scheduleLayout('local', true));
  await frame();
  doc.version = 3;
  act(() => hook.result.current.scheduleLayout('local', true));
  await frame();
  hook.unmount();
  await answer(1);
  await frame();
  expect(worker).toHaveLength(2);
  expect(doc.laidOutHere).toEqual([]);
  expect(errors).toEqual([]);
});

test('a host batch lays out here when no worker takes it', async () => {
  const { doc, worker, hook, frame, shown } = await opened();
  doc.workerAvailable = false;
  doc.version = 2;
  act(() => hook.result.current.scheduleLayout('local', true));
  await frame();
  expect(worker).toHaveLength(1);
  expect(doc.laidOutHere).toEqual([2]);
  expect(shown()).toBe('2');
});

test('a new session lays out while the replaced one still has a worker pass queued', async () => {
  const { doc, worker, errors, hook, frame, answer, shown } = await opened();
  doc.version = 2;
  act(() => hook.result.current.scheduleLayout('local', true));
  await frame();
  doc.version = 3;
  act(() => hook.result.current.scheduleLayout('local', true));
  await frame();
  expect(worker).toHaveLength(2);

  const next = fakeDocument();
  next.doc.version = 7;
  hook.rerender({ session: next.session });
  act(() => hook.result.current.runLayoutPipeline());
  expect(worker.map((pass) => pass.at)).toEqual([1, 2, 7]);
  await answer(2);
  expect(shown()).toBe('7');
  await answer(1);
  await frame();
  expect(shown()).toBe('7');
  expect(worker).toHaveLength(3);
  expect(doc.laidOutHere).toEqual([]);
  expect(errors).toEqual([]);
});

test('ready editor layout settles with both registries even when proposal ids match', async () => {
  const h = await opened({ experimentalWorkerOpen: true, ownsDocument: true });
  const paragraph = { kind: 'session', sessionId: 'session', story: 'body', paraId: 'p1' } as const;
  const local = {
    version: '1', previewVersion: 1,
    proposals: [{ id: 'same', paragraph, revisionIds: ['100:1'], state: 'rejected' as const, changed: true }],
  };
  const remote = {
    version: 'worker-1', previewVersion: 1,
    proposals: [{ id: 'same', paragraph, revisionIds: ['200:1'], state: 'accepted' as const, changed: true }],
  };
  Object.assign(h.session, {
    getProposals: () => local, mirrorWorkerDocument: () => {},
    encodeStateVector: () => new Uint8Array(), storyIds: () => ['body'],
  });
  const reply: ResidentProposalReply = {
    mirror: {
      version: remote.version,
      proposals: { previewVersion: remote.previewVersion, entries: remote.proposals.map((record) => ({
        record, key: record.id, suggest: { author: 'Host', date: '2026-10-05T00:00:00Z' },
      })) },
    },
    result: { ok: true, snapshot: remote },
    changedStories: [], updates: [], stateVector: new Uint8Array(),
    geometry: { version: remote.version, previewVersion: 1, proposals: proposalSetIdentity(remote), targets: {}, hidden: [] },
  };
  const authority = registerWorkerProposalAuthority(h.session, {
    proposal: async () => reply,
    documentRead: async () => { throw new Error('unexpected document read'); },
    handOver: async () => { throw new Error('unexpected snapshot'); },
    syncUpdate: async () => ({ version: remote.version, stateVector: new Uint8Array(), repair: null }),
    integrateProposalUpdate: () => {},
  }, {
    editorPeer: true, current: () => true, laidOut: async () => {}, relayout: () => {},
    adopted: () => {}, contentChanged: () => {},
  });
  await authority.initialize();
  expect(await authority.setStates({ expectVersion: h.session.version(), expectPreviewVersion: 1, changes: [] },
    async () => { throw new Error('unexpected peer decision'); })).toMatchObject({ ok: true });
  act(() => h.hook.result.current.runLayoutPipeline());
  const preview = { '100:1': 'rejected', '200:1': 'accepted' };
  expect(JSON.parse(h.worker[1]!.request).renderEnv.revisionPreview).toEqual(preview);
  expect(Object.keys(authority.revisionPreview()!)).toHaveLength(2);
  expect(h.session.getProposals().proposals).toEqual(local.proposals);
  expect(authority.snapshot()!.proposals).toEqual(remote.proposals);
  await h.answer(1);
  expect(isLayoutQueued(h.session)).toBe(false);
  expect(h.shown()).toBe(h.session.version());
  expect(h.errors).toEqual([]);
});

test('replacement layout retains both registry decisions while initialization is pending', async () => {
  const h = await opened({ experimentalWorkerOpen: true, ownsDocument: true });
  const paragraph = { kind: 'session', sessionId: 'session', story: 'body', paraId: 'p1' } as const;
  const local = {
    version: '1', previewVersion: 1,
    proposals: [{ id: 'same', paragraph, revisionIds: ['100:1'], state: 'rejected' as const, changed: true }],
  };
  const remote = {
    version: 'worker-1', previewVersion: 1,
    proposals: [{ id: 'same', paragraph, revisionIds: ['200:1'], state: 'accepted' as const, changed: true }],
  };
  Object.assign(h.session, {
    getProposals: () => local, mirrorWorkerDocument: () => {},
    encodeStateVector: () => new Uint8Array(), storyIds: () => ['body'],
  });
  const reply: ResidentProposalReply = {
    mirror: {
      version: remote.version,
      proposals: { previewVersion: remote.previewVersion, entries: remote.proposals.map((record) => ({
        record, key: record.id, suggest: { author: 'Host', date: '2026-10-05T00:00:00Z' },
      })) },
    },
    result: { ok: true, snapshot: remote },
    changedStories: [], updates: [], stateVector: new Uint8Array(),
    geometry: { version: remote.version, previewVersion: 1, proposals: proposalSetIdentity(remote), targets: {}, hidden: [] },
  };
  const authority = registerWorkerProposalAuthority(h.session, {
    proposal: async () => reply,
    documentRead: async () => { throw new Error('unexpected document read'); },
    handOver: async () => { throw new Error('unexpected snapshot'); },
    syncUpdate: async () => ({ version: remote.version, stateVector: new Uint8Array(), repair: null }),
    integrateProposalUpdate: () => {},
  }, {
    editorPeer: true, current: () => true, laidOut: async () => {}, relayout: () => {},
    adopted: () => {}, contentChanged: () => {},
  });
  await authority.initialize();
  expect(await authority.setStates({ expectVersion: h.session.version(), expectPreviewVersion: 1, changes: [] },
    async () => { throw new Error('unexpected peer decision'); })).toMatchObject({ ok: true });
  authority.restart();
  expect(authority.initialized).toBe(false);
  act(() => h.hook.result.current.runLayoutPipeline());
  const preview = { '100:1': 'rejected', '200:1': 'accepted' };
  expect(JSON.parse(h.worker[1]!.request).renderEnv.revisionPreview).toEqual(preview);
  expect(Object.keys(authority.revisionPreview()!)).toHaveLength(2);
  expect(h.session.getProposals().proposals).toEqual(local.proposals);
  expect(authority.snapshot()!.proposals).toEqual(remote.proposals);
  await h.answer(1);
  expect(authority.initialized).toBe(false);
  expect(isLayoutQueued(h.session)).toBe(false);
  expect(h.shown()).toBe(h.session.version());
  expect(h.errors).toEqual([]);
});
