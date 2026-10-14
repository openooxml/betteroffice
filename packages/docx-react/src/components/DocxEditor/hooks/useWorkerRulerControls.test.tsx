import { GlobalRegistrator } from '@happy-dom/global-registrator';
import { afterAll, afterEach, expect, mock, test } from 'bun:test';
import { useWorkerRulerControls } from '../peerless/workerPropertyControls';
import type { WorkerShellController } from '../peerless/workerProperties';

const ownsDom = !GlobalRegistrator.isRegistered;
if (ownsDom) GlobalRegistrator.register();
const { act, cleanup, renderHook, waitFor } = await import('@testing-library/react');
afterEach(cleanup);
afterAll(async () => { await new Promise(resolve => setTimeout(resolve, 10)); if (ownsDom) GlobalRegistrator.unregister(); });

function fixture() {
  const selection = { anchor: { story: 'body:cell', paraId: 'p', offset: 1 }, focus: { story: 'body:cell', paraId: 'p', offset: 1 } };
  const state = { status: 'ready', view: { identity: { session: 's', incarnation: 1, epoch: 1 }, version: 1, throughSeq: 1,
    workerSelection: selection, editState: { toolbar: { paragraphProperties: { indentLeft: 720, hangingIndent: true, indentFirstLine: 180 } } },
    geometry: { pages: [{ page: { pageIndex: 0 }, displayPage: { sectionIndex: 0 } }], frameEpoch: 1 } } };
  const applyQuery = mock(async (_query: unknown, _intent: unknown) => ({ ok: true, status: 'executed' }));
  const controller = { subscribe: () => () => {}, getSnapshot: () => state, applyQuery,
    environment: () => ({ readOnly: false, mode: 'editing' }), ref: { focus() {}, getCurrentPage: () => 1 },
    query: async () => ({ version: 'native', selection, value: { version: 'native', sections: [
      { section: 'ending-first', properties: { marginLeft: 1440 } }, { section: 'final', properties: { marginLeft: 2880 } },
    ], paragraphs: [{ story: 'body:cell', paraId: 'p', tabs: [{ position: 720, alignment: 'right', leader: 'dot' }] }] } }),
  } as unknown as WorkerShellController;
  const hook = renderHook(() => useWorkerRulerControls(controller));
  return { hook, applyQuery, selection };
}

test('ruler reads the current section and story paragraph and writes captured native targets', async () => {
  const { hook, applyQuery, selection } = fixture();
  await waitFor(() => expect(hook.result.current.sectionProps?.marginLeft).toBe(1440));
  expect(hook.result.current.indentLeft).toBe(720);
  expect(hook.result.current.tabStops).toEqual([{ position: 720, alignment: 'right', leader: 'dot' }]);
  await act(async () => {
    await hook.result.current.handleLeftMarginChange(2160);
    await hook.result.current.handleFirstLineIndentChange(-360);
    await hook.result.current.handleTabStopRemove(720);
  });
  expect(applyQuery.mock.calls.map(call => call[1])).toEqual([
    { kind: 'setSectionMargins', payload: { section: 'ending-first', margins: { marginLeft: 2160 }, expectVersion: 'native' } },
    { kind: 'setParagraphAttrs', payload: { selection, attrs: { indentFirstLine: 360, hangingIndent: true } } },
    { kind: 'removeTabStop', payload: { selection, positionTwips: 720, expectVersion: 'native' } },
  ]);
});

test('a ruler drag submits its last authored value once on release', async () => {
  const { hook, applyQuery } = fixture();
  await waitFor(() => expect(hook.result.current.editable).toBe(true));
  await act(async () => {
    hook.result.current.startDrag();
    await hook.result.current.handleLeftMarginChange(1800);
    await hook.result.current.handleLeftMarginChange(2160);
  });
  expect(applyQuery).not.toHaveBeenCalled();
  await act(async () => { await hook.result.current.endDrag(); });
  expect(applyQuery).toHaveBeenCalledTimes(1);
  expect(applyQuery.mock.calls[0]![1]).toMatchObject({ payload: { margins: { marginLeft: 2160 } } });
});
