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

test('review queries keep their placement across paint-only frames and re-read for a new layout', async () => {
  const h = proposalHost({ automatic: false });
  const base = await h.readAt({ kind: 'reviewSnapshot', expectPreviewVersion: 0 }, h.snapshot().version);
  if (base.status !== 'ok' || !base.value) throw new Error('Review unavailable');
  let commentGeometry = base.value.commentGeometry;
  const query = spyOn(h.controller, 'query').mockImplementation((async () => {
    const geometry = h.binding.viewStore.getSnapshot().geometry!;
    return { version: base.version, selection: null,
      value: { ...base.value, frameEpoch: geometry.frameEpoch, layoutRevision: geometry.layoutRevision, commentGeometry } };
  }) as never);
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
    expect(query.mock.calls.length - initialReads).toBe(0);
    expect(hook.result.current?.value).toBeTruthy();
    commentGeometry = { note: [{ pageIndex: 0, x: 1, y: 2, width: 10, height: 12 }] };
    const geometry = h.binding.viewStore.getSnapshot().geometry!;
    await act(async () => { h.binding.viewStore.installFrameGeometry({ ...geometry, frameEpoch: geometry.frameEpoch + 1, layoutRevision: geometry.layoutRevision + 1 }); });
    await waitFor(() => expect(hook.result.current?.value?.commentGeometry).toBe(commentGeometry));
    expect(query.mock.calls.length - initialReads).toBe(1);
  } finally { hook.unmount(); query.mockRestore(); h.controller.dispose(); }
});

type Host = ReturnType<typeof proposalHost>;

function reads(h: Host, kind: string) {
  return h.readAt.mock.calls.filter(([request]) => request.kind === kind).length;
}

async function settle() {
  await act(async () => { await new Promise(resolve => setTimeout(resolve, 0)); });
}

async function extendSelection(h: Host, offset: number) {
  await act(async () => {
    const store = h.binding.viewStore;
    const view = store.getSnapshot();
    const throughSeq = view.throughSeq + 1;
    store.acceptWorkerState(view.identity, { epoch: view.identity.epoch, version: view.version, throughSeq,
      presentedThroughSeq: view.presentedThroughSeq, editState: view.editState,
      selection: { anchor: { story: 'body', paraId: 'p0', offset: 0 }, focus: { story: 'body', paraId: 'p0', offset } } });
    const geometry = store.getSnapshot().geometry!;
    store.installFrameGeometry({ ...geometry, coversThroughSeq: throughSeq, frameEpoch: geometry.frameEpoch + 1 });
  });
  await settle();
}

async function changeVersion(h: Host) {
  await act(async () => {
    const view = h.binding.viewStore.getSnapshot();
    h.binding.viewStore.acceptWorkerState(view.identity, { epoch: view.identity.epoch, version: view.version + 1,
      throughSeq: view.throughSeq + 1, presentedThroughSeq: view.presentedThroughSeq, selection: view.workerSelection, editState: view.editState });
  });
  await settle();
}

async function changeLayout(h: Host) {
  await act(async () => {
    const store = h.binding.viewStore;
    const geometry = store.getSnapshot().geometry!;
    store.installFrameGeometry({ ...geometry, coversThroughSeq: store.getSnapshot().throughSeq,
      frameEpoch: geometry.frameEpoch + 1, layoutRevision: geometry.layoutRevision + 1 });
  });
  await settle();
}

const documentReads: (DocxDocumentRead & { kind: 'reviewSnapshot' | 'toolbarMetadata' | 'sidebar' | 'contentControlHandles' })[] = [
  { kind: 'reviewSnapshot' }, { kind: 'toolbarMetadata' }, { kind: 'sidebar', commentIds: [], expectVersion: 'native-0' }, { kind: 'contentControlHandles' },
];

for (const request of documentReads) test(`${request.kind} reads once per document version and never per selection change`, async () => {
  const h = proposalHost({ automatic: false });
  const hook = renderHook(() => useWorkerQuery(h.controller, request));
  try {
    await waitFor(() => expect(hook.result.current).not.toBeNull());
    await settle();
    const initial = reads(h, request.kind);
    expect(initial).toBe(1);
    for (let offset = 1; offset <= 8; offset++) await extendSelection(h, offset);
    expect(reads(h, request.kind) - initial).toBe(0);
    expect(hook.result.current).not.toBeNull();
    await changeVersion(h);
    expect(reads(h, request.kind) - initial).toBe(1);
    await waitFor(() => expect(hook.result.current).not.toBeNull());
  } finally { hook.unmount(); h.controller.dispose(); }
});

async function acknowledgeSelection(h: Host, offset: number) {
  await act(async () => {
    const store = h.binding.viewStore;
    const view = store.getSnapshot();
    store.acceptWorkerState(view.identity, { epoch: view.identity.epoch, version: view.version, throughSeq: view.throughSeq + 1,
      presentedThroughSeq: view.presentedThroughSeq, editState: view.editState,
      selection: { anchor: { story: 'body', paraId: 'p0', offset: 0 }, focus: { story: 'body', paraId: 'p0', offset } } });
  });
}

async function acknowledgeDuringReads(h: Host, kind: string, release: () => void, count: number) {
  for (let offset = 1; offset <= count; offset++) {
    await acknowledgeSelection(h, offset);
    const next = h.hold(kind);
    await act(async () => { release(); });
    await settle();
    release = next;
  }
  return release;
}

for (const request of documentReads) test(`${request.kind} keeps a read that a selection-only acknowledgement lands during`, async () => {
  const h = proposalHost({ automatic: false });
  const release = h.hold(request.kind);
  const hook = renderHook(() => useWorkerQuery(h.controller, request));
  try {
    await waitFor(() => expect(reads(h, request.kind)).toBe(1));
    const { version, geometry } = h.binding.viewStore.getSnapshot();
    await acknowledgeSelection(h, 1);
    await act(async () => { release(); });
    await settle();
    await waitFor(() => expect(hook.result.current).not.toBeNull());
    expect(h.binding.viewStore.getSnapshot().version).toBe(version);
    expect(h.binding.viewStore.getSnapshot().geometry).toBe(geometry);
    expect(reads(h, request.kind)).toBe(1);
  } finally { hook.unmount(); h.controller.dispose(); }
});

for (const request of documentReads) test(`${request.kind} has its result once four selection acknowledgements in a row, each during the read, stop`, async () => {
  const h = proposalHost({ automatic: false });
  const hold = h.hold(request.kind);
  const hook = renderHook(() => useWorkerQuery(h.controller, request));
  try {
    await waitFor(() => expect(reads(h, request.kind)).toBe(1));
    const { version } = h.binding.viewStore.getSnapshot();
    const release = await acknowledgeDuringReads(h, request.kind, hold, 4);
    await act(async () => { release(); });
    await settle();
    await waitFor(() => expect(hook.result.current).not.toBeNull());
    expect(h.binding.viewStore.getSnapshot().version).toBe(version);
    expect(reads(h, request.kind)).toBeLessThanOrEqual(5);
    expect(h.peak(request.kind)).toBe(1);
  } finally { hook.unmount(); h.controller.dispose(); }
});

for (const request of documentReads) test(`${request.kind} keeps one read in flight through a drag and issues none for its moves`, async () => {
  const h = proposalHost({ automatic: false });
  const release = h.hold(request.kind);
  const hook = renderHook(({ dragging }: { dragging: boolean }) => useWorkerQuery(h.controller, request, { held: dragging }),
    { initialProps: { dragging: false } });
  try {
    await waitFor(() => expect(reads(h, request.kind)).toBe(1));
    hook.rerender({ dragging: true });
    for (let offset = 1; offset <= 6; offset++) await extendSelection(h, offset);
    const next = h.hold(request.kind);
    await act(async () => { release(); });
    await settle();
    expect(reads(h, request.kind)).toBe(1);
    for (let offset = 7; offset <= 12; offset++) await extendSelection(h, offset);
    expect(reads(h, request.kind)).toBe(1);
    hook.rerender({ dragging: false });
    await act(async () => { next(); });
    await settle();
    await waitFor(() => expect(hook.result.current).not.toBeNull());
    expect(h.peak(request.kind)).toBe(1);
    expect(reads(h, request.kind)).toBeLessThanOrEqual(7);
  } finally { hook.unmount(); h.controller.dispose(); }
});

test('a selection-scoped read held through a drag has its result once four acknowledgements during it stop', async () => {
  const h = proposalHost({ automatic: false });
  const request: DocxDocumentRead & { kind: 'sectionProperties' } = { kind: 'sectionProperties', paragraphs: [] };
  const hold = h.hold(request.kind);
  const hook = renderHook(() => useWorkerQuery(h.controller, request, { held: true }));
  try {
    await waitFor(() => expect(reads(h, request.kind)).toBe(1));
    const release = await acknowledgeDuringReads(h, request.kind, hold, 4);
    await act(async () => { release(); });
    await settle();
    await waitFor(() => expect(hook.result.current).not.toBeNull());
    expect(reads(h, request.kind)).toBeLessThanOrEqual(5);
    expect(h.peak(request.kind)).toBe(1);
  } finally { hook.unmount(); h.controller.dispose(); }
});

test('content control handles re-read when the layout revision changes', async () => {
  const h = proposalHost({ automatic: false });
  const hook = renderHook(() => useWorkerQuery<'contentControlHandles'>(h.controller, { kind: 'contentControlHandles' }));
  try {
    await waitFor(() => expect(hook.result.current).not.toBeNull());
    await settle();
    const initial = reads(h, 'contentControlHandles');
    await changeLayout(h);
    expect(reads(h, 'contentControlHandles') - initial).toBe(1);
    await extendSelection(h, 3);
    expect(reads(h, 'contentControlHandles') - initial).toBe(1);
  } finally { hook.unmount(); h.controller.dispose(); }
});

type Geometry = NonNullable<ReturnType<Host['binding']['viewStore']['getSnapshot']>['geometry']>;

const sheets = (unbuilt: readonly number[]) => [0, 1, 2].map(index => ({ page: { pageId: `page-${index}`, pageIndex: index },
  bounds: { x: 0, y: index * 1100, width: 816, height: 1056 }, transform: [1, 0, 0, 1, 0, index * 1100],
  displayPage: { unbuilt: unbuilt.includes(index) } })) as unknown as Geometry['pages'];

async function installFrame(h: Host, change: Partial<Geometry> = {}) {
  await act(async () => {
    const store = h.binding.viewStore;
    const geometry = store.getSnapshot().geometry!;
    store.installFrameGeometry({ ...geometry, coversThroughSeq: store.getSnapshot().throughSeq, frameEpoch: geometry.frameEpoch + 1, ...change });
  });
  await settle();
}

async function pagedHost(request: DocxDocumentRead, change: Partial<Geometry> = {}) {
  const h = proposalHost({ automatic: false });
  await installFrame(h, { layoutEpoch: 1, pages: sheets([2]), ...change });
  const hook = renderHook(() => useWorkerQuery(h.controller, request));
  await waitFor(() => expect(hook.result.current).not.toBeNull());
  await settle();
  return { h, hook, initial: reads(h, request.kind) };
}

const layoutReads: [DocxDocumentRead, number][] = [
  [{ kind: 'reviewSnapshot' }, 1], [{ kind: 'toolbarMetadata' }, 0], [{ kind: 'sidebar', commentIds: [], expectVersion: 'native-0' }, 0], [{ kind: 'contentControlHandles' }, 1],
];

for (const [request, rereads] of layoutReads) test(`${request.kind} ${rereads ? 're-reads once' : 'does not re-read'} when scrolling builds a page`, async () => {
  const { h, hook, initial } = await pagedHost(request);
  try {
    await installFrame(h, { pages: sheets([]) });
    expect(reads(h, request.kind) - initial).toBe(rereads);
    await installFrame(h, { pages: sheets([]) });
    expect(reads(h, request.kind) - initial).toBe(rereads);
    await waitFor(() => expect(hook.result.current).not.toBeNull());
  } finally { hook.unmount(); h.controller.dispose(); }
});

for (const [request, rereads] of layoutReads) test(`${request.kind} ${rereads ? 're-reads once' : 'does not re-read'} for a relayout at the same document version`, async () => {
  const { h, hook, initial } = await pagedHost(request);
  try {
    await installFrame(h, { layoutEpoch: 2 });
    expect(reads(h, request.kind) - initial).toBe(rereads);
    await installFrame(h);
    expect(reads(h, request.kind) - initial).toBe(rereads);
    await waitFor(() => expect(hook.result.current).not.toBeNull());
  } finally { hook.unmount(); h.controller.dispose(); }
});

for (const [request] of layoutReads) test(`${request.kind} issues no read for selection-only frames over built pages`, async () => {
  const { h, hook, initial } = await pagedHost(request);
  try {
    for (let offset = 1; offset <= 6; offset++) {
      await act(async () => {
        const store = h.binding.viewStore;
        const view = store.getSnapshot();
        const throughSeq = view.throughSeq + 1;
        store.acceptWorkerState(view.identity, { epoch: view.identity.epoch, version: view.version, throughSeq,
          presentedThroughSeq: view.presentedThroughSeq, editState: view.editState,
          selection: { anchor: { story: 'body', paraId: 'p0', offset: 0 }, focus: { story: 'body', paraId: 'p0', offset } } });
      });
      await installFrame(h, { pages: sheets([2]), caret: { x: offset, y: 0, width: 1, height: 12 }, selectionRects: [{ x: 0, y: 0, width: offset, height: 12 }] });
    }
    expect(reads(h, request.kind) - initial).toBe(0);
    expect(hook.result.current).not.toBeNull();
  } finally { hook.unmount(); h.controller.dispose(); }
});

test('content control handles re-read once when the provisional layout completes', async () => {
  const { h, hook, initial } = await pagedHost({ kind: 'contentControlHandles' }, { layoutProvisional: true });
  try {
    await installFrame(h, { layoutProvisional: false });
    expect(reads(h, 'contentControlHandles') - initial).toBe(1);
    await installFrame(h, { layoutProvisional: false });
    expect(reads(h, 'contentControlHandles') - initial).toBe(1);
  } finally { hook.unmount(); h.controller.dispose(); }
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
