import { GlobalRegistrator } from '@happy-dom/global-registrator';
import { afterAll, expect, mock, test } from 'bun:test';
import type { LayoutComputation } from '@betteroffice/docx/editor';
import { LayoutSelectionGate, type ResidentMeasurementConfig } from '@betteroffice/docx/layout';
import type { YrsRenderEnv, YrsSession } from '@betteroffice/docx/yrs';
import { registerWorkerProposalAuthority, workerProposalFailure } from '../internals/workerProposalAuthority';
import type { WorkerOpenedDocument } from './useDisplayList';

const ownsDom = !GlobalRegistrator.isRegistered;
if (ownsDom) GlobalRegistrator.register();
const { act, renderHook } = await import('@testing-library/react');
const { useLayoutPipeline } = await import('./useLayoutPipeline');

afterAll(async () => {
  if (ownsDom) await GlobalRegistrator.unregister();
});

test('a worker pass answering after unmount touches no session', async () => {
  let freed = false;
  const touched: string[] = [];
  const touch = (method: string) => {
    if (freed) touched.push(method);
  };
  const session = {
    version: () => {
      touch('version');
      return '1';
    },
    layoutFontRequirementsJson: () => '[]',
    layoutDocumentWithRegionsRetainedJson: () => {
      touch('layoutDocumentWithRegionsRetainedJson');
      throw new Error('null pointer passed to rust');
    },
  } as unknown as YrsSession;
  let answer: (computation: LayoutComputation | null) => void = () => {};
  const errors: Error[] = [];
  const { result, unmount } = renderHook(() =>
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
      layoutInWorker: () => new Promise((resolve) => (answer = resolve)),
    })
  );
  act(() => result.current.runLayoutPipeline());
  unmount();
  freed = true;
  await act(async () => {
    answer(null);
    await new Promise((done) => setTimeout(done, 0));
  });
  expect(touched).toEqual([]);
  expect(errors).toEqual([]);
});

test.each(['font preflight', 'initial layout'])('an editable %s failure releases the first round and its queued save', async (stage) => {
  const failure = new Error(`${stage} failed`);
  const session = {
    version: () => '1',
    getProposals: () => ({ version: '1', previewVersion: 0, proposals: [] }),
    layoutFontRequirementsJson: () => {
      if (stage === 'font preflight') throw failure;
      return '[]';
    },
  } as unknown as YrsSession;
  const worker = {
    proposal: mock(async () => { throw new Error('unexpected worker proposal'); }),
    documentRead: async () => { throw new Error('unexpected worker read'); },
    handOver: async () => { throw new Error('unexpected handover'); },
  } satisfies Pick<WorkerOpenedDocument, 'proposal' | 'documentRead' | 'handOver'>;
  let waiting!: () => void;
  const waitingForLayout = new Promise<void>((resolve) => { waiting = resolve; });
  const authority = registerWorkerProposalAuthority(session, worker, {
    editorPeer: true, current: () => true,
    laidOut: () => { waiting(); return new Promise<void>(() => {}); },
    relayout: () => {}, adopted: () => {}, contentChanged: () => {},
  });
  const main = mock(async () => { throw new Error('unexpected peer proposal'); });
  const round = authority.propose({ expectVersion: '1', proposals: [] }, main).catch((error: unknown) => error);
  const save = mock(async () => 'saved');
  const saving = authority.save(save).catch((error: unknown) => error);
  await waitingForLayout;
  const errors: Error[] = [];
  const view = renderHook(() => useLayoutPipeline({
    document: null, session, renderEnv: {}, pageGap: 24, zoom: 1,
    residentMeasurementConfig: () => ({} as ResidentMeasurementConfig),
    deferLayoutPass: () => false,
    pagesContainerRef: { current: null }, viewportLayoutRef: { current: null },
    syncCoordinator: new LayoutSelectionGate(), getScrollContainer: () => null,
    layoutInWorker: () => Promise.reject(failure),
    onError: (error) => errors.push(error),
  }));
  try {
    expect(view.result.current.layout).toBeNull();
    expect(save).not.toHaveBeenCalled();
    await act(async () => {
      view.result.current.runLayoutPipeline();
      expect(await round).toBe(failure);
      expect(await saving).toBe(failure);
    });
    expect(errors).toEqual([failure]);
    expect(workerProposalFailure(session)).toBe(failure);
    expect(worker.proposal).not.toHaveBeenCalled();
    expect(main).not.toHaveBeenCalled();
    expect(save).toHaveBeenCalledTimes(1);
  } finally {
    view.unmount();
  }
});
