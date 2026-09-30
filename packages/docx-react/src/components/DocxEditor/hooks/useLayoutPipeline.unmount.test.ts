import { GlobalRegistrator } from '@happy-dom/global-registrator';
import { afterAll, expect, test } from 'bun:test';
import type { LayoutComputation } from '@betteroffice/docx/editor';
import { LayoutSelectionGate, type ResidentMeasurementConfig } from '@betteroffice/docx/layout';
import type { YrsRenderEnv, YrsSession } from '@betteroffice/docx/yrs';

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
