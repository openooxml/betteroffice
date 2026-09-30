import { GlobalRegistrator } from '@happy-dom/global-registrator';
import { afterAll, expect, spyOn, test } from 'bun:test';
import type { LayoutComputation } from '@betteroffice/docx/editor';
import {
  LayoutSelectionGate,
  type ResidentFontRequirement,
  type ResidentMeasurementConfig,
} from '@betteroffice/docx/layout';
import type { Layout } from '@betteroffice/docx/layout/pagination';
import type { YrsRenderEnv, YrsSession } from '@betteroffice/docx/yrs';
import { isLayoutQueued, isSupersededLayout, sourceVersionOf } from '../internals/layoutProvenance';
import { deferWorkerOpenReplica } from '../internals/workerOpenReplica';

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
  reject(error: Error): void;
}

function fakeDocument() {
  const doc = {
    version: 1,
    laidOutHere: [] as number[],
    workerAvailable: true,
    fontsReady: true,
    preflights: 0,
  };
  const session = {
    version: () => String(doc.version),
    layoutFontRequirementsJson: () => {
      doc.preflights += 1;
      return '[]';
    },
    layoutDocumentWithRegionsRetainedJson: () => {
      doc.laidOutHere.push(doc.version);
      return JSON.stringify({ layout: { pages: [] }, notesConverged: true });
    },
    residentWorkerProbe: () => ({ layoutRevision: 1 }),
  } as unknown as YrsSession;
  return { doc, session };
}

for (const proposals of ['present', 'empty', 'unavailable', 'throwing']) {
  for (const inWorker of [false, true]) {
    test(`decision fonts (${proposals}, worker=${inWorker}) warm after the layout request`, async () => {
      const { session } = fakeDocument();
      const order: string[] = [];
      const requirements: ResidentFontRequirement[] = [
        { key: 'regular', family: 'Calibri', bold: false, italic: false },
      ];
      const warm: ResidentFontRequirement[] = [
        ...requirements,
        { key: 'decision', family: 'Symbol', bold: false, italic: false },
      ];
      const preflight = (where: string, input: string) => {
        const superset = JSON.parse(input).revisionFontSuperset === true;
        order.push(`${where}:${superset ? 'superset' : 'exact'}`);
        return JSON.stringify(superset ? warm : requirements);
      };
      const layoutInputs: string[] = [];
      const measured: ResidentFontRequirement[][] = [];
      const warmed: ResidentFontRequirement[][] = [];
      if (proposals !== 'unavailable') {
        Object.assign(session, {
          getProposals: () => {
            if (proposals === 'throwing') throw new Error('Proposal registry unavailable');
            return {
              proposals: proposals === 'present' ? [{ id: 'proposal', revisionIds: ['r1'] }] : [],
            };
          },
        });
      }
      if (inWorker) {
        deferWorkerOpenReplica(session, () => new Promise(() => {}), () => {}, () => {});
      }
      Object.assign(session, {
        layoutFontRequirementsJson: (input: string) => preflight('host', input),
        layoutDocumentWithRegionsRetainedJson: (input: string) => {
          order.push('layout');
          layoutInputs.push(input);
          return JSON.stringify({ layout: { pages: [] }, notesConverged: true });
        },
      });
      const hook = renderHook(() =>
        useLayoutPipeline({
          document: null,
          session,
          renderEnv: {} as YrsRenderEnv,
          pageGap: 24,
          zoom: 1,
          residentMeasurementConfig: (required) => {
            measured.push(required);
            return {} as ResidentMeasurementConfig;
          },
          warmFontRequirements: (fonts) => warmed.push(fonts),
          deferLayoutPass: () => false,
          pagesContainerRef: { current: null },
          viewportLayoutRef: { current: null },
          syncCoordinator: new LayoutSelectionGate(),
          getScrollContainer: () => null,
          experimentalWorkerOpen: inWorker,
          fontRequirementsInWorker: (_session, input) =>
            Promise.resolve(preflight('worker', input)),
          layoutInWorker: (_session, input) => {
            order.push('layout');
            layoutInputs.push(input);
            return Promise.resolve({
              layout: { pages: [] } as unknown as Layout,
              notesConverged: true,
            });
          },
        })
      );
      try {
        await act(async () => {
          hook.result.current.runLayoutPipeline({ onHost: !inWorker });
          await new Promise((done) => setTimeout(done, 5));
        });
        const where = inWorker ? 'worker' : 'host';
        expect(order).toEqual([
          `${where}:exact`,
          'layout',
          ...(proposals === 'present' ? [`${where}:superset`] : []),
        ]);
        expect(measured).toEqual([requirements]);
        expect(warmed).toEqual(proposals === 'present' ? [warm] : []);
        const request = hook.result.current.getLayoutRequest();
        expect(request).not.toBeNull();
        for (const input of [...layoutInputs, request!]) {
          expect(JSON.parse(input)).not.toHaveProperty('revisionFontSuperset');
          expect(JSON.parse(input)).not.toHaveProperty('measurement.revisionFontSuperset');
          expect(JSON.parse(input).renderEnv).not.toHaveProperty('revisionFontSuperset');
        }
        if (proposals === 'present') {
          const pass = () =>
            act(async () => {
              hook.result.current.runLayoutPipeline({ onHost: !inWorker });
              await new Promise((done) => setTimeout(done, 5));
            });
          const supersets = () => order.filter((step) => step.endsWith(':superset')).length;
          await pass();
          expect(supersets()).toBe(1);
          Object.assign(session, {
            getProposals: () => ({
              proposals: [
                { id: 'proposal', revisionIds: ['r1'] },
                { id: 'another', revisionIds: ['r2'] },
              ],
            }),
          });
          await pass();
          expect(supersets()).toBe(2);
          Object.assign(session, {
            getProposals: () => ({ proposals: [{ id: 'proposal', revisionIds: ['r3'] }] }),
          });
          await pass();
          expect(supersets()).toBe(3);
        }
      } finally {
        hook.unmount();
      }
    });
  }
}

for (const answer of ['rejects', 'null'] as const) {
  test(`a worker superset preflight that ${answer} warms nothing and fails nothing`, async () => {
    const { session } = fakeDocument();
    const hostInputs: string[] = [];
    const errors: Error[] = [];
    const warmed: ResidentFontRequirement[][] = [];
    deferWorkerOpenReplica(session, () => new Promise(() => {}), () => {}, () => {});
    Object.assign(session, {
      getProposals: () => ({ proposals: [{ id: 'proposal', revisionIds: ['r1'] }] }),
      layoutFontRequirementsJson: (input: string) => {
        hostInputs.push(input);
        return '[]';
      },
    });
    const hook = renderHook(() =>
      useLayoutPipeline({
        document: null,
        session,
        renderEnv: {} as YrsRenderEnv,
        pageGap: 24,
        zoom: 1,
        residentMeasurementConfig: () => ({}) as ResidentMeasurementConfig,
        warmFontRequirements: (fonts) => warmed.push(fonts),
        deferLayoutPass: () => false,
        pagesContainerRef: { current: null },
        viewportLayoutRef: { current: null },
        syncCoordinator: new LayoutSelectionGate(),
        getScrollContainer: () => null,
        onError: (error) => errors.push(error),
        experimentalWorkerOpen: true,
        fontRequirementsInWorker: (_session, input) =>
          JSON.parse(input).revisionFontSuperset
            ? answer === 'rejects'
              ? Promise.reject(new Error('superset preflight failed'))
              : Promise.resolve(null)
            : Promise.resolve('[]'),
        layoutInWorker: () =>
          Promise.resolve({ layout: { pages: [] } as unknown as Layout, notesConverged: true }),
      })
    );
    try {
      await act(async () => {
        hook.result.current.runLayoutPipeline();
        await new Promise((done) => setTimeout(done, 5));
      });
      expect(sourceVersionOf(hook.result.current.layout)).toBe('1');
      expect(errors).toEqual([]);
      expect(warmed).toEqual([]);
      expect(hostInputs).toEqual([]);
    } finally {
      hook.unmount();
    }
  });
}

/** A document whose version each test moves on; worker passes answer when the test says. */
async function opened() {
  const { doc, session } = fakeDocument();
  const worker: WorkerPass[] = [];
  const errors: Error[] = [];
  const syncCoordinator = new LayoutSelectionGate();
  const hook = renderHook(({ session }) =>
    useLayoutPipeline({
      document: null,
      session,
      renderEnv: {} as YrsRenderEnv,
      pageGap: 24,
      zoom: 1,
      residentMeasurementConfig: () => (doc.fontsReady ? ({} as ResidentMeasurementConfig) : null),
      deferLayoutPass: () => false,
      pagesContainerRef: { current: null },
      viewportLayoutRef: { current: null },
      syncCoordinator,
      getScrollContainer: () => null,
      onError: (error) => errors.push(error),
      layoutInWorker: Object.assign(
        (asked: YrsSession) =>
          doc.workerAvailable
            ? new Promise<LayoutComputation | null>((resolve, reject) => {
                worker.push({
                  at: Number(asked.version()),
                  answer: () =>
                    resolve({ layout: { pages: [] } as unknown as Layout, notesConverged: true }),
                  fail: () => resolve(null),
                  reject,
                });
              })
            : null,
        { available: () => doc.workerAvailable }
      ),
    }),
    { initialProps: { session } }
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

test('a font warm-up pass never lays out here, even without a worker', async () => {
  const { doc, worker, errors, hook, frame, shown } = await opened();

  act(() => hook.result.current.scheduleWarmLayout());
  await frame();
  expect(worker.map((pass) => pass.at)).toEqual([1, 1]);
  act(() => worker[1]!.fail());
  await frame();
  expect(doc.laidOutHere).toEqual([]);

  doc.workerAvailable = false;
  const preflights = doc.preflights;
  act(() => hook.result.current.scheduleWarmLayout());
  await frame();
  expect(doc.laidOutHere).toEqual([]);
  expect(doc.preflights).toBe(preflights);

  doc.version = 2;
  act(() => {
    hook.result.current.scheduleWarmLayout();
    hook.result.current.scheduleLayout('local', true);
  });
  await frame();
  expect(doc.laidOutHere).toEqual([2]);

  doc.workerAvailable = true;
  act(() => hook.result.current.scheduleWarmLayout());
  act(() => hook.result.current.runLayoutPipeline());
  expect(worker.map((pass) => pass.at)).toEqual([1, 1, 2]);
  act(() => worker[2]!.fail());
  await frame();
  expect(doc.laidOutHere).toEqual([2, 2]);
  expect(shown()).toBe('2');
  expect(errors).toEqual([]);
});

test('a direct run queued behind a worker pass is never a warm-up', async () => {
  const { doc, worker, errors, hook, frame, shown } = await opened();
  doc.version = 2;
  act(() => hook.result.current.scheduleLayout('local', true));
  await frame();
  expect(worker.map((pass) => pass.at)).toEqual([1, 2]);

  doc.version = 3;
  act(() => {
    hook.result.current.scheduleWarmLayout();
    hook.result.current.runLayoutPipeline();
  });
  doc.workerAvailable = false;
  act(() => worker[1]!.fail());
  await frame();
  expect(doc.laidOutHere).toEqual([3]);
  expect(shown()).toBe('3');
  expect(errors).toEqual([]);
});

for (const ending of ['fails', 'throws'] as const) {
  test(`a warm-up queued behind a worker pass that ${ending} changes nothing it does`, async () => {
    const outcomes: unknown[] = [];
    for (const warmUp of [false, true]) {
      const { doc, worker, errors, hook, frame, shown } = await opened();
      doc.version = 2;
      act(() => hook.result.current.scheduleLayout('local', true));
      await frame();
      expect(worker.map((pass) => pass.at)).toEqual([1, 2]);

      if (warmUp) act(() => hook.result.current.scheduleWarmLayout());
      await frame();
      expect(worker).toHaveLength(2);
      doc.workerAvailable = false;
      const log = spyOn(console, 'error').mockImplementation(() => {});
      try {
        await act(async () => {
          if (ending === 'fails') worker[1]!.fail();
          else worker[1]!.reject(new Error('worker lost'));
          await new Promise((done) => setTimeout(done, 0));
        });
        await frame();
      } finally {
        log.mockRestore();
      }
      outcomes.push({
        laidOutHere: doc.laidOutHere,
        shown: shown(),
        errors: errors.map((error) => error.message),
      });
      hook.unmount();
    }
    expect(outcomes[1]).toEqual(outcomes[0]);
    expect(outcomes[0]).toEqual(
      ending === 'fails'
        ? { laidOutHere: [2], shown: '2', errors: [] }
        : { laidOutHere: [], shown: '1', errors: ['worker lost'] }
    );
  });
}

test('a warm-up behind a worker pass holds no settle and supersedes nothing', async () => {
  const { doc, session, worker, errors, hook, frame, answer, shown } = await opened();
  doc.version = 2;
  act(() => hook.result.current.scheduleLayout('local', true));
  await frame();
  act(() => hook.result.current.scheduleWarmLayout());
  await frame();
  expect(worker.map((pass) => pass.at)).toEqual([1, 2]);
  expect(isLayoutQueued(session)).toBe(false);

  await answer(1);
  expect(shown()).toBe('2');
  expect(isSupersededLayout(hook.result.current.layout)).toBe(false);
  await frame();
  expect(worker.map((pass) => pass.at)).toEqual([1, 2, 2]);
  await answer(2);
  expect(shown()).toBe('2');
  expect(doc.laidOutHere).toEqual([]);
  expect(errors).toEqual([]);
});

test("a warm-up waits for the full layout that replaces a pass's first pages", async () => {
  const { doc, session } = fakeDocument();
  const asked: number[] = [];
  const pages = (count: number) =>
    ({ pages: Array.from({ length: count }, () => ({})) }) as unknown as Layout;
  let finish!: (complete: LayoutComputation | null) => void;
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
      layoutInWorker: (owner) => {
        asked.push(Number(owner.version()));
        return Promise.resolve(
          asked.length === 1
            ? {
                layout: pages(1),
                notesConverged: true,
                complete: new Promise<LayoutComputation | null>((resolve) => {
                  finish = resolve;
                }),
              }
            : { layout: pages(3), notesConverged: true }
        );
      },
    })
  );
  const settle = (ms: number) =>
    act(async () => {
      await new Promise((done) => setTimeout(done, ms));
    });
  try {
    act(() => hook.result.current.runLayoutPipeline());
    await settle(0);
    expect(hook.result.current.layout?.pages).toHaveLength(1);

    act(() => hook.result.current.scheduleWarmLayout());
    await settle(40);
    expect(asked).toEqual([1]);

    finish({ layout: pages(3), notesConverged: true });
    await settle(0);
    expect(hook.result.current.layout?.pages).toHaveLength(3);
    expect(isSupersededLayout(hook.result.current.layout)).toBe(false);
    await settle(40);
    expect(asked).toEqual([1, 1]);
    expect(doc.laidOutHere).toEqual([]);
  } finally {
    hook.unmount();
  }
});

test('a real pass that fails in the worker with only a warm-up queued behind it still lays out here', async () => {
  const { doc, worker, errors, hook, frame, shown } = await opened();
  doc.version = 2;
  act(() => hook.result.current.scheduleLayout('local', true));
  await frame();
  expect(worker.map((pass) => pass.at)).toEqual([1, 2]);
  act(() => hook.result.current.scheduleWarmLayout());
  await frame();
  doc.workerAvailable = false;
  act(() => worker[1]!.fail());
  await frame();
  await frame();
  expect(errors).toEqual([]);
  expect(doc.laidOutHere).toEqual([2]);
  expect(shown()).toBe('2');
});

test('without the warm-up the failed real pass lays out here', async () => {
  const { doc, worker, errors, hook, frame, shown } = await opened();
  doc.version = 2;
  act(() => hook.result.current.scheduleLayout('local', true));
  await frame();
  doc.workerAvailable = false;
  act(() => worker[1]!.fail());
  await frame();
  expect(errors).toEqual([]);
  expect(doc.laidOutHere).toEqual([2]);
  expect(shown()).toBe('2');
});

test('a hydrated worker-open session warms decision fonts here, after the pass', async () => {
  const { session } = fakeDocument();
  const order: string[] = [];
  const warmed: ResidentFontRequirement[][] = [];
  const decision: ResidentFontRequirement = {
    key: 'decision',
    family: 'Symbol',
    bold: false,
    italic: false,
  };
  Object.assign(session, {
    getProposals: () => ({ proposals: [{ id: 'proposal', revisionIds: ['r1'] }] }),
    layoutFontRequirementsJson: (input: string) => {
      const superset = JSON.parse(input).revisionFontSuperset === true;
      order.push(superset ? 'host:superset' : 'host:exact');
      return JSON.stringify(superset ? [decision] : []);
    },
  });
  const hook = renderHook(() =>
    useLayoutPipeline({
      document: null,
      session,
      renderEnv: {} as YrsRenderEnv,
      pageGap: 24,
      zoom: 1,
      residentMeasurementConfig: () => ({}) as ResidentMeasurementConfig,
      warmFontRequirements: (fonts) => warmed.push(fonts),
      deferLayoutPass: () => false,
      pagesContainerRef: { current: null },
      viewportLayoutRef: { current: null },
      syncCoordinator: new LayoutSelectionGate(),
      getScrollContainer: () => null,
      experimentalWorkerOpen: true,
      fontRequirementsInWorker: () => null,
      layoutInWorker: () => {
        order.push('layout');
        return Promise.resolve({ layout: { pages: [] } as unknown as Layout, notesConverged: true });
      },
    })
  );
  try {
    await act(async () => {
      hook.result.current.runLayoutPipeline();
      await new Promise((done) => setTimeout(done, 5));
    });
    expect(order).toEqual(['host:exact', 'layout', 'host:superset']);
    expect(warmed).toEqual([[decision]]);
  } finally {
    hook.unmount();
  }
});
