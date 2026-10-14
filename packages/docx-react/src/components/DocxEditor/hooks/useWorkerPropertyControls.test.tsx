import { GlobalRegistrator } from '@happy-dom/global-registrator';
import { afterAll, afterEach, expect, mock, test } from 'bun:test';
import { useWorkerPageSetupControls } from '../peerless/workerPropertyControls';
import { useWorkerWatermarkControls } from '../peerless/workerPropertyControls';
import type { createPeerlessShellController } from '../peerless/shellController';

const ownsDom = !GlobalRegistrator.isRegistered;
if (ownsDom) GlobalRegistrator.register();
const { act, cleanup, renderHook } = await import('@testing-library/react');
afterEach(cleanup);
afterAll(async () => { await new Promise(resolve => setTimeout(resolve, 10)); if (ownsDom) GlobalRegistrator.unregister(); });

function fixture() {
  const query = mock(async (request: { kind: string; target?: { kind: string } }) => ({ version: 'opened-native', value: request.kind === 'partSnapshot' ?
    { metadata: { body: { sections: [] } } } : request.target?.kind === 'watermark' ?
      { kind: 'watermark', watermark: { kind: 'text', text: 'Draft' } } : { kind: 'section', section: 'final', properties: { marginLeft: 1440 } } }));
  const applyQuery = mock(async (_query: unknown, _intent: unknown) => ({ ok: true, status: 'executed' }));
  const onError = mock(() => {});
  const focus = mock(() => {});
  const controller = { query, applyQuery, ref: { focus, getCurrentPage: () => 1 }, getSnapshot: () => ({ view: { geometry: { pages: [] } } }),
    environment: () => ({ readOnly: false, mode: 'editing' }) } as unknown as ReturnType<typeof createPeerlessShellController>;
  const hook = renderHook(() => ({ page: useWorkerPageSetupControls(controller, onError), watermark: useWorkerWatermarkControls(controller, onError) }));
  return { hook, query, applyQuery, onError, focus };
}

test('worker page setup and watermark dialogs query properties and write the opened native version', async () => {
  const { hook, applyQuery } = fixture();
  await act(async () => { await hook.result.current.page.handleOpenPageSetup(); await hook.result.current.watermark.handleOpenWatermark(); });
  expect(hook.result.current.page.currentProps).toEqual({ marginLeft: 1440 });
  expect(hook.result.current.watermark.currentWatermark).toMatchObject({ text: 'Draft' });
  await act(async () => {
    await hook.result.current.page.handlePageSetupApply({ marginLeft: 2880, orientation: 'landscape' });
    await hook.result.current.watermark.handleWatermarkApply(null);
  });
  expect(applyQuery.mock.calls.map(call => call[1])).toEqual([
    { kind: 'setSectionProperties', payload: { section: 'final', patch: { marginLeft: 2880, orientation: 'landscape' }, expectVersion: 'opened-native' } },
    { kind: 'setWatermark', payload: { watermark: null, expectVersion: 'opened-native' } },
  ]);
});

test('a refused dialog command reports failure without reopening or replacing the document', async () => {
  const { hook, applyQuery, onError } = fixture();
  await act(async () => { await hook.result.current.page.handleOpenPageSetup(); });
  applyQuery.mockImplementation(async () => ({ ok: false, failure: { message: 'stale-version' } }) as never);
  await act(async () => { await hook.result.current.page.handlePageSetupApply({ marginTop: 720 }); });
  expect(onError).toHaveBeenCalledWith(expect.objectContaining({ message: 'stale-version' }));
  expect(applyQuery).toHaveBeenCalledTimes(1);
});

test('automatic watermark size omits undefined fields from worker validation', async () => {
  const { hook, applyQuery } = fixture();
  await act(async () => { await hook.result.current.watermark.handleOpenWatermark(); });
  await act(async () => { await hook.result.current.watermark.handleWatermarkApply({ kind: 'text', text: 'Draft', font: 'Liberation Sans',
    color: '#cccccc', semitransparent: true, layout: 'diagonal', fontSize: undefined }); });
  expect((applyQuery.mock.calls[0]![1] as { payload: { watermark: unknown } }).payload.watermark).not.toHaveProperty('fontSize');
});
