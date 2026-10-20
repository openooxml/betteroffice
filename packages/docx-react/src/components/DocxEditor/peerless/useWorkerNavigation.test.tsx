import { GlobalRegistrator } from '@happy-dom/global-registrator';
import { afterAll, afterEach, expect, mock, spyOn, test } from 'bun:test';
import { useRef } from 'react';
import { createPeerlessShellController, WorkerReadSupersededError } from './shellController';
import { createPeerlessViewStore } from './viewStore';
import { useWorkerNavigation } from './useWorkerNavigation';
import { createInputAdmission } from './inputAdmission';
import { ResidentEngineWorkerClient } from '@betteroffice/docx/yrs';
import type { PeerlessEditSessionBinding } from '../hooks/usePeerlessEditSession';

const ownsDom = !GlobalRegistrator.isRegistered;
if (ownsDom) GlobalRegistrator.register();
const { act, cleanup, fireEvent, render, waitFor } = await import('@testing-library/react');
afterEach(cleanup);
afterAll(async () => {
  await act(async () => { await new Promise<void>((resolve) => setImmediate(resolve)); });
  if (ownsDom) await GlobalRegistrator.unregister();
});

test('worker bookmark target selects and scrolls', async () => {
  const identity = { session: 'bookmark', incarnation: 1, epoch: 1 };
  const loc = { story: 'body:t0:r0c0', paraId: 'target', offset: 1 };
  const viewStore = createPeerlessViewStore({ identity, version: 0, throughSeq: 0, presentedThroughSeq: 0,
    workerSelection: null, anchors: null, editState: undefined, geometry: { identity, frameEpoch: 1, layoutRevision: 1,
      coversThroughSeq: 0, caret: { x: 10, y: 1200, width: 1, height: 14 }, selectionRects: [], pages: [] } });
  const submit = mock(() => ({ status: 'submitted', ticket: { ack: Promise.resolve({ status: 'applied' }), presented: Promise.resolve('presented') } }));
  const binding = { viewStore, admission: {}, session: { flush: async () => {} }, commands: { submit } } as unknown as PeerlessEditSessionBinding;
  const read = mock(async () => ({ status: 'ok', version: 'native', value: { loc, position: 50 } }));
  const controller = createPeerlessShellController({ zoom: 1, focus() {}, read: read as never });
  controller.attach(binding);
  let navigation!: ReturnType<typeof useWorkerNavigation>;
  function Harness() {
    const scroller = useRef<HTMLDivElement>(null);
    const host = useRef<HTMLDivElement>(null);
    navigation = useWorkerNavigation(controller, () => binding, scroller, host, 1, {});
    return <div ref={scroller} data-testid="scroll"><div ref={host} /></div>;
  }
  try {
    const view = render(<Harness />);
    await act(async () => { expect(await navigation.scrollToBookmark('target')).toBe(true); });
    expect(read).toHaveBeenCalledWith({ kind: 'bookmarkTarget', story: 'body', name: 'target' });
    expect(submit).toHaveBeenCalledWith({ kind: 'select', payload: { selection: { anchor: loc, focus: loc } } }, { source: 'selection', fence: true });
    expect(view.getByTestId('scroll').scrollTop).toBe(1176);
    read.mockImplementation(async () => ({ status: 'ok', version: 'native', value: null } as never));
    await act(async () => { expect(await navigation.scrollToBookmark('missing')).toBe(false); });
    expect(submit).toHaveBeenCalledTimes(1);
  } finally { cleanup(); controller.dispose(); viewStore.dispose(); }
});

test('zoom updates the page indicator without requiring a scroll event', async () => {
  const identity = { session: 'pages', incarnation: 1, epoch: 1 };
  const viewStore = createPeerlessViewStore({ identity, version: 0, throughSeq: 0, presentedThroughSeq: 0,
    workerSelection: null, anchors: null, editState: undefined, geometry: { identity, frameEpoch: 1, layoutRevision: 1,
      coversThroughSeq: 0, caret: null, selectionRects: [], pages: [0, 1, 2].map((pageIndex) => ({
        page: { pageId: String(pageIndex + 1), pageIndex }, bounds: { x: 0, y: pageIndex * 124, width: 200, height: 100 },
        transform: [1, 0, 0, 1, 0, 0],
      })) } });
  const controller = createPeerlessShellController({ zoom: 1, focus() {}, read: async () => ({} as never) });
  controller.attach({ viewStore, admission: {} } as PeerlessEditSessionBinding);
  function Harness({ zoom }: { zoom: number }) {
    const scroller = useRef<HTMLDivElement>(null);
    const host = useRef<HTMLDivElement>(null);
    const navigation = useWorkerNavigation(controller, () => null, scroller, host, zoom, {});
    return <div ref={scroller} data-testid="scroll"><div ref={host} data-testid="host" />
      <output aria-label="Current page">{navigation.pageInfo.currentPage}</output></div>;
  }
  const box = spyOn(Element.prototype, 'getBoundingClientRect').mockImplementation(function (this: Element) {
    const top = this.getAttribute('data-testid') === 'host' ? -(this.parentElement?.scrollTop ?? 0) : 0;
    return { top, bottom: top + 60, left: 0, right: 200, width: 200, height: 60, x: 0, y: top, toJSON() {} };
  });
  try {
    const view = render(<Harness zoom={1} />);
    const scroller = view.getByTestId('scroll');
    Object.defineProperty(scroller, 'clientHeight', { value: 60 });
    scroller.scrollTop = 230;
    fireEvent.scroll(scroller);
    await waitFor(() => expect(view.getByLabelText('Current page').textContent).toBe('3'));
    view.rerender(<Harness zoom={2} />);
    expect(view.getByLabelText('Current page').textContent).toBe('2');
  } finally { cleanup(); box.mockRestore(); controller.dispose(); viewStore.dispose(); }
});

type Shown = { caret: { x: number; y: number; width: number; height: number } | null; selectionRects: { x: number; y: number; width: number; height: number }[] };

function reviewNavigation(targets: Record<string, unknown>, shown: Shown = { caret: { x: 10, y: 400, width: 1, height: 12 }, selectionRects: [] }) {
  const identity = { session: 'review', incarnation: 1, epoch: 1 };
  const viewStore = createPeerlessViewStore({ identity, version: 1, throughSeq: 0, presentedThroughSeq: 0,
    workerSelection: null, anchors: null, editState: undefined, geometry: { identity, frameEpoch: 1, layoutRevision: 1,
      coversThroughSeq: 0, ...shown, pages: [] } });
  const read = mock(async (request: { kind: string; target?: { kind: string; id?: number; revisionId?: string } }) => ({ status: 'ok', version: 'native',
    value: request.kind === 'reviewTarget' ? targets[JSON.stringify(request.target)] ?? null : null }));
  const submit = mock((_intent: unknown, _options?: unknown) => {
    const geometry = viewStore.getSnapshot().geometry!;
    if (!geometry.caret && geometry.selectionRects[0]) geometry.caret = geometry.selectionRects[0];
    return { status: 'submitted', ticket: { seq: 1, ack: Promise.resolve({ status: 'applied', seq: 1 }), presented: Promise.resolve('presented') } };
  });
  const binding = { session: { flush: async () => {} }, admission: {}, commands: { submit }, viewStore } as unknown as PeerlessEditSessionBinding;
  const controller = createPeerlessShellController({ zoom: 2, focus() {}, read: read as never, readAt: read as never });
  controller.attach(binding);
  function Harness() {
    const scroller = useRef<HTMLDivElement>(null);
    const host = useRef<HTMLDivElement>(null);
    const navigation = useWorkerNavigation(controller, () => binding, scroller, host, 2, {});
    controller.update({ zoom: 2, focus() {}, read: read as never, readAt: read as never, scrollToReviewTarget: navigation.scrollToReviewTarget });
    return <div ref={scroller} data-testid="scroll"><div ref={host} /></div>;
  }
  const view = render(<Harness />);
  return { controller, binding, viewStore, read, submit, scroller: view.getByTestId('scroll'), dispose() { controller.dispose(); viewStore.dispose(); } };
}

test('review navigation selects the first anchored range of a comment or revision and scrolls its caret into view', async () => {
  const range = { story: 'body', start: { paraId: 'p', offset: 3 }, end: { paraId: 'q', offset: 2 } };
  const { controller, read, submit, scroller, dispose } = reviewNavigation({
    [JSON.stringify({ kind: 'comment', id: 7 })]: { id: 'native', ranges: [range, { ...range, story: 'hf:rId1' }] },
    [JSON.stringify({ kind: 'revision', revisionId: '12' })]: { id: '12', ranges: [range] },
  });
  try {
    expect(await controller.ref.scrollToComment(7)).toBe(true);
    expect(read).toHaveBeenCalledWith({ kind: 'reviewTarget', target: { kind: 'comment', id: 7 } });
    expect(submit.mock.calls[0]).toEqual([{ kind: 'select', payload: { selection: {
      anchor: { story: 'body', paraId: 'p', offset: 3 }, focus: { story: 'body', paraId: 'p', offset: 3 } } } }, { source: 'selection', fence: true }]);
    expect(submit.mock.calls[1]).toEqual([{ kind: 'select', payload: { selection: {
      anchor: { story: 'body', paraId: 'q', offset: 2 }, focus: { story: 'body', paraId: 'p', offset: 3 } } } }, { source: 'selection', fence: true }]);
    expect(scroller.scrollTop).toBe(400 * 2 - 24);
    expect(await controller.revealReviewTarget({ kind: 'revision', revisionId: '12' })).toBe(true);
    expect(await controller.ref.scrollToChange(12)).toBe(true);
    expect(submit).toHaveBeenCalledTimes(6);
  } finally { dispose(); }
});

test('review navigation builds a start caret when a selected range initially has no caret', async () => {
  const range = { story: 'body', start: { paraId: 'p', offset: 3 }, end: { paraId: 'p', offset: 9 } };
  const { controller, scroller, dispose } = reviewNavigation({ [JSON.stringify({ kind: 'comment', id: 7 })]: { id: 'native', ranges: [range] } },
    { caret: null, selectionRects: [{ x: 10, y: 900, width: 40, height: 12 }, { x: 10, y: 914, width: 20, height: 12 }] });
  try {
    expect(await controller.ref.scrollToComment(7)).toBe(true);
    expect(scroller.scrollTop).toBe(900 * 2 - 24);
  } finally { dispose(); }
});

test('review navigation resolves false for an unknown comment or revision without moving the selection', async () => {
  const { controller, submit, scroller, dispose } = reviewNavigation({});
  try {
    expect(await controller.ref.scrollToComment(99)).toBe(false);
    expect(await controller.ref.scrollToComment(1.5)).toBe(false);
    expect(await controller.ref.scrollToChange(99)).toBe(false);
    expect(submit).not.toHaveBeenCalled();
    expect(scroller.scrollTop).toBe(0);
  } finally { dispose(); }
});


const reviewRange = { story: 'body', start: { paraId: 'p', offset: 3 }, end: { paraId: 'q', offset: 9 } };
const reviewTargets = { [JSON.stringify({ kind: 'comment', id: 7 })]: { id: 'native', ranges: [reviewRange] },
  [JSON.stringify({ kind: 'revision', revisionId: '12' })]: { id: '12', ranges: [reviewRange] } };

for (const member of ['scrollToComment', 'scrollToChange'] as const) {
  test(`${member} resolves false while opening and when typing or replacement supersedes its read`, async () => {
    const { controller, read, viewStore, dispose } = reviewNavigation(reviewTargets);
    const id = member === 'scrollToComment' ? 7 : 12;
    try {
      read.mockImplementationOnce(async () => { const view = viewStore.getSnapshot(); viewStore.acceptWorkerState(view.identity, { epoch: view.identity.epoch, version: view.version + 1, throughSeq: 0, presentedThroughSeq: 0, selection: null, editState: view.editState }); return { status: 'ok', version: 'native', value: reviewTargets[JSON.stringify({ kind: 'comment', id: 7 })] }; });
      expect(await controller.ref[member](id)).toBe(false);
      read.mockImplementationOnce(async () => { controller.opening({}); return { status: 'ok', version: 'native', value: null }; });
      expect(await controller.ref[member](id)).toBe(false);
      expect(await controller.ref[member](id)).toBe(false);
    } finally { dispose(); }
  });
}

for (const phase of ['ack', 'presented'] as const) {
  test(`review navigation resolves false when ${phase} is superseded`, async () => {
    const { controller, submit, dispose } = reviewNavigation(reviewTargets);
    try {
      submit.mockImplementationOnce(() => ({ status: 'submitted', ticket: { seq: 1,
        ack: phase === 'ack' ? Promise.reject(new WorkerReadSupersededError('Worker read was superseded')) : Promise.resolve({ status: 'applied', seq: 1 }),
        presented: phase === 'presented' ? Promise.reject(new WorkerReadSupersededError('Document was replaced')) : Promise.resolve('presented') } }));
      expect(await controller.ref.scrollToComment(7)).toBe(false);
    } finally { dispose(); }
  });
}

test('review navigation preserves unrelated worker failures', async () => {
  const { controller, read, dispose } = reviewNavigation(reviewTargets);
  try {
    read.mockRejectedValueOnce(new Error('Worker failed'));
    await expect(controller.ref.scrollToComment(7)).rejects.toThrow('Worker failed');
  } finally { dispose(); }
});

test('review navigation refuses a target without start geometry', async () => {
  const { controller, dispose } = reviewNavigation(reviewTargets, { caret: null, selectionRects: [] });
  try { expect(await controller.ref.scrollToComment(7)).toBe(false); } finally { dispose(); }
});

test('review navigation builds the start before scrolling a range whose end is the only available rectangle', async () => {
  const { controller, submit, viewStore, scroller, dispose } = reviewNavigation(reviewTargets,
    { caret: null, selectionRects: [{ x: 0, y: 900, width: 20, height: 12 }] });
  const original = submit.getMockImplementation()!;
  submit.mockImplementation((intent, options) => {
    const selection = (intent as { payload: { selection: { anchor: { paraId: string }; focus: { paraId: string } } } }).payload.selection;
    if (selection.anchor.paraId === 'p' && selection.focus.paraId === 'p') {
      viewStore.getSnapshot().geometry!.caret = { x: 0, y: 200, width: 1, height: 12 };
    }
    return original(intent, options);
  });
  try {
    expect(await controller.ref.scrollToComment(7)).toBe(true);
    expect(scroller.scrollTop).toBe(200 * 2 - 24);
  } finally { dispose(); }
});


for (const phase of ['ack', 'presented'] as const) {
  test(`review navigation resolves false when replacement disposes ${phase}`, async () => {
    const { controller, submit, dispose } = reviewNavigation(reviewTargets);
    const replaced = () => { controller.opening({}); throw Object.assign(new Error('Session was disposed'), { code: 'disposed' }); };
    try {
      submit.mockImplementationOnce(() => ({ status: 'submitted', ticket: { seq: 1,
        ack: phase === 'ack' ? Promise.resolve().then(replaced) : Promise.resolve({ status: 'applied', seq: 1 }),
        presented: phase === 'presented' ? Promise.resolve().then(replaced) : Promise.resolve('presented') } }));
      expect(await controller.ref.scrollToComment(7)).toBe(false);
    } finally { dispose(); }
  });

  test(`review navigation preserves unrelated ${phase} failures`, async () => {
    const { controller, submit, dispose } = reviewNavigation(reviewTargets);
    try {
      submit.mockImplementationOnce(() => ({ status: 'submitted', ticket: { seq: 1,
        ack: phase === 'ack' ? Promise.reject(new Error('Worker failed')) : Promise.resolve({ status: 'applied', seq: 1 }),
        presented: phase === 'presented' ? Promise.reject(new Error('Worker failed')) : Promise.resolve('presented') } }));
      await expect(controller.ref.scrollToComment(7)).rejects.toThrow('Worker failed');
    } finally { dispose(); }
  });
}


test('review navigation resolves false when replacement disposes the composition read barrier', async () => {
  const { controller, binding, viewStore, read, dispose } = reviewNavigation(reviewTargets);
  const admission = createInputAdmission({ session: binding.session, viewStore, echoPainter: { disable() {} } as never });
  binding.admission = admission;
  admission.reserveComposition();
  try {
    const pending = controller.ref.scrollToComment(7);
    controller.opening({});
    admission.dispose();
    expect(await pending).toBe(false);
    expect(read).not.toHaveBeenCalled();
  } finally { admission.dispose(); dispose(); }
});

for (const destroyed of [true, false]) {
  test(`review navigation ${destroyed ? 'resolves false for teardown' : 'preserves a worker failure'} during a replacement read`, async () => {
    const { controller, read, dispose } = reviewNavigation(reviewTargets);
    const port: NonNullable<ConstructorParameters<typeof ResidentEngineWorkerClient>[0]> = {
      onmessage: null, onmessageerror: null, onerror: null, postMessage() {}, terminate() {},
    };
    const client = new ResidentEngineWorkerClient(port);
    read.mockImplementationOnce(async () => {
      const pending = client.documentRead({ kind: 'reviewTarget', target: { kind: 'comment', id: 7 } });
      controller.opening({});
      if (destroyed) client.destroy();
      else port.onerror?.({ message: 'Worker failed' } as ErrorEvent);
      return pending as never;
    });
    try {
      if (destroyed) expect(await controller.ref.scrollToComment(7)).toBe(false);
      else await expect(controller.ref.scrollToComment(7)).rejects.toThrow('Resident engine worker failed: Worker failed');
    } finally { client.destroy(); dispose(); }
  });
}
