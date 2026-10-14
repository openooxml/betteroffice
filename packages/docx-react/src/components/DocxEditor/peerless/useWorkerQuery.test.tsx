import { GlobalRegistrator } from '@happy-dom/global-registrator';
import { afterAll, afterEach, expect, mock, test } from 'bun:test';
import type { ResidentDocumentRead } from '@betteroffice/docx/yrs';
import { useWorkerQuery } from './useWorkerQuery';
import type { WorkerShellController } from './workerProperties';

const ownsDom = !GlobalRegistrator.isRegistered;
if (ownsDom) GlobalRegistrator.register();
const { act, cleanup, renderHook, waitFor } = await import('@testing-library/react');
afterEach(cleanup);
afterAll(async () => { await new Promise(resolve => setTimeout(resolve, 10)); if (ownsDom) GlobalRegistrator.unregister(); });

test('worker queries preserve sticky bytes and deduplicate equal requests', async () => {
  const state = { status: 'ready', view: { identity: { session: 's', incarnation: 1, epoch: 1 }, version: 1, throughSeq: 1, geometry: null } };
  const query = mock(async (_request: ResidentDocumentRead) => ({ version: 'native', value: null }));
  const controller = { subscribe: () => () => {}, getSnapshot: () => state, query } as unknown as WorkerShellController;
  const request = (byte: number): ResidentDocumentRead & { kind: 'properties' } => ({ kind: 'properties', target: {
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
