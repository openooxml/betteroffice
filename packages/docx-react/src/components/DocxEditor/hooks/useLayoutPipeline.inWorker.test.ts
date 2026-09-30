import { GlobalRegistrator } from '@happy-dom/global-registrator';
import { afterAll, expect, test } from 'bun:test';
import type { LayoutComputation } from '@betteroffice/docx/editor';
import { LayoutSelectionGate, type ResidentMeasurementConfig } from '@betteroffice/docx/layout';
import type { Layout } from '@betteroffice/docx/layout/pagination';
import type { YrsRenderEnv, YrsSession } from '@betteroffice/docx/yrs';
import { sourceVersionOf } from '../internals/layoutProvenance';

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
}

/** A document whose version each test moves on; worker passes answer when the test says. */
async function opened() {
  const doc = { version: 1, laidOutHere: [] as number[], workerAvailable: true };
  const session = {
    version: () => String(doc.version),
    layoutFontRequirementsJson: () => '[]',
    layoutDocumentWithRegionsRetainedJson: () => {
      doc.laidOutHere.push(doc.version);
      return JSON.stringify({ layout: { pages: [] }, notesConverged: true });
    },
    residentWorkerProbe: () => ({ layoutRevision: 1 }),
  } as unknown as YrsSession;
  const worker: WorkerPass[] = [];
  const errors: Error[] = [];
  const hook = renderHook(() =>
    useLayoutPipeline({
      document: null,
      session,
      renderEnv: {} as YrsRenderEnv,
      pageGap: 24,
      zoom: 1,
      residentMeasurementConfig: () => ({}) as ResidentMeasurementConfig,
      deferLayoutPass: () => false,
      pagesContainerRef: { current: null },
      viewportLayoutRef: { current: null },
      syncCoordinator: new LayoutSelectionGate(),
      getScrollContainer: () => null,
      onError: (error) => errors.push(error),
      layoutInWorker: () =>
        doc.workerAvailable
          ? new Promise<LayoutComputation | null>((resolve) => {
              worker.push({
                at: doc.version,
                answer: () => resolve({ layout: { pages: [] } as unknown as Layout }),
              });
            })
          : null,
    })
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
  return { doc, worker, errors, hook, frame, answer, shown };
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
  await frame();
  expect(worker.map((pass) => pass.at)).toEqual([1, 2, 4]);
  await answer(2);
  expect(shown()).toBe('4');
  expect(doc.laidOutHere).toEqual([]);
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
