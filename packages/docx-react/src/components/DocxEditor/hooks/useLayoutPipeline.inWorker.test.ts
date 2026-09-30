import { GlobalRegistrator } from '@happy-dom/global-registrator';
import { afterAll, expect, test } from 'bun:test';
import type { LayoutComputation } from '@betteroffice/docx/editor';
import { LayoutSelectionGate, type ResidentMeasurementConfig } from '@betteroffice/docx/layout';
import type { Layout } from '@betteroffice/docx/layout/pagination';
import type { YrsRenderEnv, YrsSession } from '@betteroffice/docx/yrs';
import { isLayoutQueued, isSupersededLayout, sourceVersionOf } from '../internals/layoutProvenance';

const ownsDom = !GlobalRegistrator.isRegistered;
if (ownsDom) GlobalRegistrator.register();
const { act, renderHook } = await import('@testing-library/react');
const { useLayoutPipeline } = await import('./useLayoutPipeline');

afterAll(async () => {
  if (ownsDom) await GlobalRegistrator.unregister();
});

interface WorkerPass {
  /** The document version the pass was asked for. */
  at: number;
  answer(): void;
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
async function opened() {
  const { doc, session } = fakeDocument();
  const worker: WorkerPass[] = [];
  const errors: Error[] = [];
  const syncCoordinator = new LayoutSelectionGate();
  const hook = renderHook(({ session, renderEnv }: HookProps) =>
    useLayoutPipeline({
      document: null,
      session,
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
      layoutInWorker: (asked) =>
        doc.workerAvailable
          ? new Promise<LayoutComputation | null>((resolve) => {
              worker.push({
                at: Number(asked.version()),
                answer: () =>
                  resolve({ layout: { pages: [] } as unknown as Layout, notesConverged: true }),
                fail: () => resolve(null),
              });
            })
          : null,
    }),
    { initialProps: { session } as HookProps }
  );
  const frame = () =>
    act(async () => {
      await new Promise((done) => setTimeout(done, 40));
    });
  const answer = (index: number) =>
    act(async () => {
      worker[index]!.answer();
      await new Promise((done) => setTimeout(done, 0));
    });
  const shown = () => sourceVersionOf(hook.result.current.layout);
  act(() => hook.result.current.runLayoutPipeline());
  await answer(0);
  expect(shown()).toBe('1');
  return { doc, session, worker, errors, hook, frame, answer, shown };
}

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
    await new Promise((done) => setTimeout(done, 0));
  });
  await frame();
  // Only the queued pass lays out, here: the failed one left it the layout.
  expect(worker.map((pass) => pass.at)).toEqual([1, 2]);
  expect(doc.laidOutHere).toEqual([2]);
  expect(shown()).toBe('2');
  expect(isSupersededLayout(hook.result.current.layout)).toBe(false);
  expect(errors).toEqual([]);
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
