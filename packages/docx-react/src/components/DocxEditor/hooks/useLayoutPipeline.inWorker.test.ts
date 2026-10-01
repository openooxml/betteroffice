import { GlobalRegistrator } from '@happy-dom/global-registrator';
import { afterAll, afterEach, expect, spyOn, test } from 'bun:test';
import type { LayoutComputation } from '@betteroffice/docx/editor';
import { LayoutSelectionGate, type ResidentMeasurementConfig } from '@betteroffice/docx/layout';
import type { Layout } from '@betteroffice/docx/layout/pagination';
import { proposalSetIdentity, type ResidentProposalReply, type YrsRenderEnv, type YrsSession } from '@betteroffice/docx/yrs';
import { isLayoutQueued, isSupersededLayout, sourceVersionOf } from '../internals/layoutProvenance';
import { deferWorkerOpenReplica } from '../internals/workerOpenReplica';
import { registerWorkerProposalAuthority } from '../internals/workerProposalAuthority';
import type { FontRequirementsInWorker, WorkerLayoutComputation } from './useDisplayList';
import { SupersededPreviewError } from '../internals/supersededPreview';

const ownsDom = !GlobalRegistrator.isRegistered;
if (ownsDom) GlobalRegistrator.register();
const { act, cleanup, renderHook } = await import('@testing-library/react');
const { useLayoutPipeline } = await import('./useLayoutPipeline');
const restoreFrames: Array<() => void> = [];

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
  answer(computation?: WorkerLayoutComputation): void;
  fail(): void;
}

interface HookProps {
  session: YrsSession;
  renderEnv?: YrsRenderEnv;
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
  const replica = pendingReplica
    ? deferWorkerOpenReplica(session, () => new Promise(() => {}), () => {}, () => {})
    : null;
  const ensureReplica = replica ? spyOn(replica, 'ensure') : null;
  const worker: WorkerPass[] = [];
  const errors: Error[] = [];
  const syncCoordinator = new LayoutSelectionGate();
  const hook = renderHook(({ session, renderEnv }: HookProps) =>
    useLayoutPipeline({
      document: null,
      session,
      experimentalWorkerOpen,
      renderEnv: renderEnv ?? ({} as YrsRenderEnv),
      pageGap: 24,
      zoom: 1,
      residentMeasurementConfig: () => (doc.fontsReady ? doc.measurement : null),
      deferLayoutPass: () => false,
      pagesContainerRef: { current: null },
      viewportLayoutRef: { current: null },
      syncCoordinator,
      getScrollContainer: () => null,
      onError: (error) => errors.push(error),
      fontRequirementsInWorker,
      layoutInWorker: (asked) =>
        doc.workerAvailable
          ? new Promise<WorkerLayoutComputation | null>((resolve) => {
              worker.push({
                at: Number(asked.version()),
                answer: (computation) =>
                  resolve(computation ?? {
                    layout: { pages: [] } as unknown as Layout, notesConverged: true,
                  }),
                fail: () => resolve(null),
              });
            })
          : null,
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
  await answer(0);
  expect(shown()).toBe('1');
  return {
    doc, session, worker, errors, hook, frame, answer, shown, replica, ensureReplica,
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

test('a local edit in the same frame as a host batch keeps the pass here', async () => {
  const { doc, worker, hook, frame, shown } = await opened();
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
  expect(isSupersededLayout(hook.result.current.layout)).toBe(true);
  await frame();
  expect(worker.map((pass) => pass.at)).toEqual([1, 2, 2]);
  expect(isLayoutQueued(session)).toBe(false);
  await answer(2);
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
