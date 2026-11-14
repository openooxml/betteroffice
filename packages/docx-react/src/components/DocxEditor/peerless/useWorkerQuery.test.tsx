import { GlobalRegistrator } from '@happy-dom/global-registrator';
import { afterAll, afterEach, expect, mock, spyOn, test } from 'bun:test';
import type { DocxDocumentRead } from '@betteroffice/docx/yrs';
import { useWorkerQuery } from './useWorkerQuery';
import type { WorkerShellController } from './workerProperties';
import { proposalHost } from './workerProposalHost.testing';

const ownsDom = !GlobalRegistrator.isRegistered;
if (ownsDom) GlobalRegistrator.register();
const { act, cleanup, renderHook, waitFor } = await import('@testing-library/react');
afterEach(cleanup);
afterAll(async () => { await new Promise(resolve => setTimeout(resolve, 10)); if (ownsDom) GlobalRegistrator.unregister(); });

test('worker queries preserve sticky bytes and deduplicate equal requests', async () => {
  const state = { status: 'ready', view: { identity: { session: 's', incarnation: 1, epoch: 1 }, version: 1, throughSeq: 1, geometry: null } };
  const query = mock(async (_request: DocxDocumentRead) => ({ version: 'native', value: null }));
  const controller = { subscribe: () => () => {}, getSnapshot: () => state,
    subscribeShell: () => () => {}, getShellSnapshot: () => state, query } as unknown as WorkerShellController;
  const request = (byte: number): DocxDocumentRead & { kind: 'properties' } => ({ kind: 'properties', target: {
    kind: 'objectSelection', selection: {
      anchor: { story: 'body', paraId: 'p', offset: 0, sticky: Uint8Array.of(byte, 2) },
      focus: { story: 'body', paraId: 'p', offset: 1, sticky: Uint8Array.of(byte, 3) },
    },
  } });
  const original = request(1);
  const hook = renderHook(({ value }) => useWorkerQuery<'properties'>(controller, value), { initialProps: { value: original } });
  await waitFor(() => expect(query).toHaveBeenCalledTimes(1));
  expect(query.mock.calls[0]![0]).toEqual(original);
  expect(query.mock.calls[0]![0]).toBe(original);
  await act(async () => { hook.rerender({ value: request(1) }); });
  expect(query).toHaveBeenCalledTimes(1);
  const changed = request(4);
  await act(async () => { hook.rerender({ value: changed }); });
  await waitFor(() => expect(query).toHaveBeenCalledTimes(2));
  expect(query.mock.calls[1]![0]).toBe(changed);
});

test('review queries retain every frame without rendering for unchanged review data', async () => {
  const h = proposalHost({ automatic: false });
  const base = await h.readAt({ kind: 'reviewSnapshot', expectPreviewVersion: 0 }, h.snapshot().version);
  if (base.status !== 'ok' || !base.value) throw new Error('Review unavailable');
  let commentGeometry = base.value.commentGeometry;
  const query = spyOn(h.controller, 'query').mockImplementation((async () => ({ version: base.version, selection: null,
    value: { ...base.value, frameEpoch: h.binding.viewStore.getSnapshot().geometry!.frameEpoch, commentGeometry } })) as never);
  let renders = 0;
  const hook = renderHook(() => { renders++; return useWorkerQuery<'reviewSnapshot'>(h.controller, { kind: 'reviewSnapshot' }); });
  try {
    await waitFor(() => expect(hook.result.current?.value).toBeTruthy());
    const initialRenders = renders;
    const initialReads = query.mock.calls.length;
    for (let index = 0; index < 40; index++) {
      const geometry = structuredClone(h.binding.viewStore.getSnapshot().geometry!);
      await act(async () => { h.binding.viewStore.installFrameGeometry({ ...geometry, frameEpoch: geometry.frameEpoch + 1 }); });
    }
    expect(renders - initialRenders).toBeLessThanOrEqual(2);
    expect(query.mock.calls.length - initialReads).toBe(40);
    hook.rerender();
    expect(hook.result.current?.value?.frameEpoch).toBe(h.binding.viewStore.getSnapshot().geometry!.frameEpoch);
    commentGeometry = { note: [{ pageIndex: 0, x: 1, y: 2, width: 10, height: 12 }] };
    const geometry = h.binding.viewStore.getSnapshot().geometry!;
    await act(async () => { h.binding.viewStore.installFrameGeometry({ ...geometry, frameEpoch: geometry.frameEpoch + 1 }); });
    expect(hook.result.current?.value?.commentGeometry).toBe(commentGeometry);
  } finally { hook.unmount(); query.mockRestore(); h.controller.dispose(); }
});

test('toolbar metadata queries do not repeat for paint-only frames', async () => {
  const h = proposalHost({ automatic: false });
  const query = spyOn(h.controller, 'query').mockImplementation((async () => ({ version: 'native', selection: null,
    value: { styles: [], fontFamilies: [] } })) as never);
  const hook = renderHook(() => useWorkerQuery<'toolbarMetadata'>(h.controller, { kind: 'toolbarMetadata' }));
  try {
    await waitFor(() => expect(hook.result.current?.value).toBeTruthy());
    for (let index = 0; index < 40; index++) {
      const geometry = structuredClone(h.binding.viewStore.getSnapshot().geometry!);
      await act(async () => { h.binding.viewStore.installFrameGeometry({ ...geometry, frameEpoch: geometry.frameEpoch + 1 }); });
    }
    expect(query).toHaveBeenCalledTimes(1);
  } finally { hook.unmount(); query.mockRestore(); h.controller.dispose(); }
});
