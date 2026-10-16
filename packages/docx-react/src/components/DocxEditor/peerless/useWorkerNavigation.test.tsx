import { GlobalRegistrator } from '@happy-dom/global-registrator';
import { afterAll, afterEach, expect, mock, spyOn, test } from 'bun:test';
import { useRef } from 'react';
import { createPeerlessShellController } from './shellController';
import { createPeerlessViewStore } from './viewStore';
import { useWorkerNavigation } from './useWorkerNavigation';
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

function reviewNavigation(targets: Record<string, unknown>) {
  const identity = { session: 'review', incarnation: 1, epoch: 1 };
  const viewStore = createPeerlessViewStore({ identity, version: 1, throughSeq: 0, presentedThroughSeq: 0,
    workerSelection: null, anchors: null, editState: undefined, geometry: { identity, frameEpoch: 1, layoutRevision: 1,
      coversThroughSeq: 0, caret: { x: 10, y: 400, width: 1, height: 12 }, selectionRects: [], pages: [] } });
  const read = mock(async (request: { kind: string; target?: { kind: string; id?: number; revisionId?: string } }) => ({ status: 'ok', version: 'native',
    value: request.kind === 'reviewTarget' ? targets[JSON.stringify(request.target)] ?? null : null }));
  const submit = mock((_intent: unknown, _options?: unknown) => ({ status: 'submitted', ticket: { seq: 1, ack: Promise.resolve({ status: 'applied', seq: 1 }), presented: Promise.resolve('presented') } }));
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
  return { controller, read, submit, scroller: view.getByTestId('scroll'), dispose() { controller.dispose(); viewStore.dispose(); } };
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
      anchor: { story: 'body', paraId: 'p', offset: 3 }, focus: { story: 'body', paraId: 'q', offset: 2 } } } }, { source: 'selection', fence: true }]);
    expect(scroller.scrollTop).toBe(400 * 2 - 24);
    expect(await controller.revealReviewTarget({ kind: 'revision', revisionId: '12' })).toBe(true);
    expect(await controller.ref.scrollToChange(12)).toBe(true);
    expect(submit).toHaveBeenCalledTimes(3);
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
