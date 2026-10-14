import { GlobalRegistrator } from '@happy-dom/global-registrator';
import { afterAll, afterEach, expect, mock, test } from 'bun:test';
import { useWorkerRulerControls } from '../peerless/workerPropertyControls';
import type { WorkerShellController } from '../peerless/workerProperties';

const ownsDom = !GlobalRegistrator.isRegistered;
if (ownsDom) GlobalRegistrator.register();
const { cleanup, renderHook, waitFor } = await import('@testing-library/react');
afterEach(cleanup);
afterAll(async () => { await new Promise(resolve => setTimeout(resolve, 10)); if (ownsDom) GlobalRegistrator.unregister(); });

test('section margins remain editable while paragraph controls wait for a native selection', async () => {
  const state = { status: 'ready', view: { identity: { session: 's', incarnation: 1, epoch: 1 }, version: 0, throughSeq: 0,
    workerSelection: null, geometry: null, editState: { toolbar: null } } };
  const applyQuery = mock(async () => ({ ok: true, status: 'executed' }));
  const controller = { subscribe: () => () => {}, getSnapshot: () => state, applyQuery,
    environment: () => ({ mode: 'editing', readOnly: false }), ref: { getCurrentPage: () => 1, focus() {} },
    query: async () => ({ version: 'native', selection: null, value: { version: 'native',
      sections: [{ section: 'final', properties: { marginLeft: 1440 } }], paragraphs: [] } }),
  } as unknown as WorkerShellController;
  const hook = renderHook(() => useWorkerRulerControls(controller));
  await waitFor(() => expect(hook.result.current.sectionProps?.marginLeft).toBe(1440));
  expect(hook.result.current.editable).toBe(true);
  expect(Reflect.get(hook.result.current, 'paragraphEditable')).toBe(false);
  await hook.result.current.handleIndentLeftChange(720);
  await hook.result.current.handleTabStopsChange([{ position: 1440, alignment: 'left' }]);
  expect(applyQuery).not.toHaveBeenCalled();
  await hook.result.current.handleLeftMarginChange(2160);
  expect(applyQuery).toHaveBeenCalledTimes(1);
});
