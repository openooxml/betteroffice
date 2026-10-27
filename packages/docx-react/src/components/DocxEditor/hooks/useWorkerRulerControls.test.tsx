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

function liveFixture() {
  const selection = { anchor: { story: 'body', paraId: 'p', offset: 1 }, focus: { story: 'body', paraId: 'p', offset: 1 } };
  const view = (version: number, session = 's') => ({ identity: { session, incarnation: 1, epoch: 1 }, version, throughSeq: version,
    workerSelection: selection, editState: { toolbar: { paragraphProperties: { indentLeft: 720, indentFirstLine: 360 } } },
    geometry: { pages: [{ page: { pageIndex: 0 }, displayPage: { sectionIndex: 0 } }], frameEpoch: version } });
  let state = { status: 'ready', view: view(1) };
  const listeners = new Set<() => void>();
  const reads: Array<(value: unknown) => void> = [];
  const read = (version: number, paraId = 'p') => ({ version: `v${version}`,
    selection: { anchor: { ...selection.anchor, paraId }, focus: { ...selection.focus, paraId } }, value: { version: `v${version}`,
    sections: [{ section: 'final', properties: { marginLeft: 2160 } }], paragraphs: [{ story: 'body', paraId: 'p', tabs: [{ position: 1440, alignment: 'left' }] },
      { story: 'body', paraId: 'q', tabs: [] }] } });
  const publish = (next: typeof state) => { state = next; for (const listener of listeners) listener(); };
  const applyQuery = mock(async (_query: unknown, _intent: unknown) => ({ ok: true, status: 'executed' }));
  const controller = { subscribe: (listener: () => void) => { listeners.add(listener); return () => { listeners.delete(listener); }; },
    getSnapshot: () => state, applyQuery,
    environment: () => ({ readOnly: false, mode: 'editing' }), ref: { focus() {}, getCurrentPage: () => 1 },
    query: () => new Promise(resolve => { reads.push(resolve); }),
  } as unknown as WorkerShellController;
  const hook = renderHook(() => useWorkerRulerControls(controller));
  const settle = (version: number) => act(async () => { reads.shift()!(read(version)); });
  return { hook, view, read, reads, settle, applyQuery, publish: (next: typeof state) => act(() => publish(next)) };
}

test('ruler markers keep the last read while an edit re-reads, and drop it for a new document', async () => {
  const { hook, view, reads, settle, publish } = liveFixture();
  await settle(1);
  await waitFor(() => expect(hook.result.current.paragraphEditable).toBe(true));

  publish({ status: 'ready', view: view(2) });
  expect(reads).toHaveLength(1);
  expect(hook.result.current.paragraphEditable).toBe(true);
  expect(hook.result.current.editable).toBe(true);
  expect(hook.result.current.sectionProps?.marginLeft).toBe(2160);
  expect(hook.result.current.tabStops).toEqual([{ position: 1440, alignment: 'left' }]);

  publish({ status: 'ready', view: view(1, 'replaced') });
  expect(hook.result.current.paragraphEditable).toBe(false);
  expect(hook.result.current.sectionProps).toBeUndefined();
});

test('a ruler write during a re-read waits for a fresh read instead of the held one', async () => {
  const { hook, view, read, reads, settle, applyQuery, publish } = liveFixture();
  await settle(1);
  await waitFor(() => expect(hook.result.current.paragraphEditable).toBe(true));
  publish({ status: 'ready', view: view(2) });
  let written!: Promise<void>;
  act(() => { written = hook.result.current.handleIndentLeftChange(1440); });
  expect(reads).toHaveLength(2);
  expect(applyQuery).not.toHaveBeenCalled();
  reads.splice(1, 1)[0]!(read(2, 'q'));
  await act(async () => { await written; });
  expect(applyQuery.mock.calls.map(call => [(call[0] as { version: string }).version, call[1]])).toEqual([
    ['v2', { kind: 'setParagraphAttrs', payload: { attrs: { indentLeft: 1440 }, selection: read(2, 'q').selection } }],
  ]);
});

test('a tab list authored on the held read is not written over tabs a fresh read changed', async () => {
  const { hook, view, read, reads, settle, applyQuery, publish } = liveFixture();
  await settle(1);
  await waitFor(() => expect(hook.result.current.paragraphEditable).toBe(true));
  publish({ status: 'ready', view: view(2) });
  const tabs = [...hook.result.current.tabStops!, { position: 4320, alignment: 'left' as const }];
  let written!: Promise<void>;
  act(() => { written = hook.result.current.handleTabStopsChange(tabs); });
  const fresh = read(2);
  const moved = { ...fresh, value: { ...fresh.value, paragraphs: [{ story: 'body', paraId: 'p',
    tabs: [{ position: 1440, alignment: 'left' }, { position: 2880, alignment: 'center' }] }] } };
  reads.splice(1, 1)[0]!(moved);
  await act(async () => { await written; });
  expect(applyQuery).not.toHaveBeenCalled();

  act(() => { written = hook.result.current.handleTabStopsChange(tabs); });
  reads.splice(1, 1)[0]!(read(2));
  await act(async () => { await written; });
  expect(applyQuery.mock.calls.map(call => call[1])).toEqual([
    { kind: 'setTabStops', payload: { selection: read(2).selection, tabs, expectVersion: 'v2' } },
  ]);
});

test('a tab drag that started on the held read is not written over tabs the re-read changed before release', async () => {
  const { hook, view, read, reads, settle, applyQuery, publish } = liveFixture();
  await settle(1);
  await waitFor(() => expect(hook.result.current.paragraphEditable).toBe(true));
  publish({ status: 'ready', view: view(2) });
  const tabs = [...hook.result.current.tabStops!, { position: 4320, alignment: 'left' as const }];
  act(() => { hook.result.current.startDrag(); });
  const fresh = read(2);
  const movedTabs = [{ position: 1440, alignment: 'left' as const }, { position: 2880, alignment: 'center' as const }];
  const moved = { ...fresh, value: { ...fresh.value, paragraphs: [{ story: 'body', paraId: 'p', tabs: movedTabs }] } };
  await act(async () => { reads.shift()!(moved); });
  await waitFor(() => expect(hook.result.current.tabStops).toEqual(movedTabs));
  let released!: Promise<void>;
  await act(async () => { await hook.result.current.handleTabStopsChange(tabs); released = hook.result.current.endDrag(); });
  reads.shift()!(moved);
  await act(async () => { await released; });
  expect(applyQuery).not.toHaveBeenCalled();

  const next = [...hook.result.current.tabStops!, { position: 4320, alignment: 'left' as const }];
  await act(async () => {
    hook.result.current.startDrag();
    await hook.result.current.handleTabStopsChange(next);
    await hook.result.current.endDrag();
  });
  expect(applyQuery.mock.calls.map(call => call[1])).toEqual([
    { kind: 'setTabStops', payload: { selection: moved.selection, tabs: next, expectVersion: 'v2' } },
  ]);
});
