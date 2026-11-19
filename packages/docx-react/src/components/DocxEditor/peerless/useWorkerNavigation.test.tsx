import { GlobalRegistrator } from '@happy-dom/global-registrator';
import { afterAll, afterEach, expect, jest, mock, spyOn, test } from 'bun:test';
import { useRef } from 'react';
import { createPeerlessShellController, WorkerReadSupersededError } from './shellController';
import { createPeerlessViewStore, type PeerlessFrameGeometry, type PeerlessViewSnapshot } from './viewStore';
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

for (const event of ['window resize', 'width change', 'height change', 'plain scroll'] as const) {
  test(`${event} refreshes the page indicator and only plain scrolling reveals it`, () => {
    const identity = { session: 'pages', incarnation: 1, epoch: 1 };
    const viewStore = createPeerlessViewStore({ identity, version: 0, throughSeq: 0, presentedThroughSeq: 0,
      workerSelection: null, anchors: null, editState: undefined, geometry: { identity, frameEpoch: 1, layoutRevision: 1,
        coversThroughSeq: 0, caret: null, selectionRects: [], pages: [0, 1, 2].map((pageIndex) => ({
          page: { pageId: String(pageIndex + 1), pageIndex }, bounds: { x: 0, y: pageIndex * 124, width: 200, height: 100 },
          transform: [1, 0, 0, 1, 0, 0],
        })) } });
    const controller = createPeerlessShellController({ zoom: 1, focus() {}, read: async () => ({} as never) });
    controller.attach({ viewStore, admission: {} } as PeerlessEditSessionBinding);
    let width = 200, height = 60;
    let navigation!: ReturnType<typeof useWorkerNavigation>;
    function Harness() {
      const scroller = useRef<HTMLDivElement>(null);
      const host = useRef<HTMLDivElement>(null);
      navigation = useWorkerNavigation(controller, () => null, scroller, host, 1, {});
      return <div ref={(element) => {
        scroller.current = element;
        if (element) Object.defineProperties(element, {
          clientWidth: { configurable: true, get: () => width }, clientHeight: { configurable: true, get: () => height },
        });
      }} data-testid="scroll"><div ref={host} data-testid="host" /></div>;
    }
    const box = spyOn(Element.prototype, 'getBoundingClientRect').mockImplementation(function (this: Element) {
      const top = this.getAttribute('data-testid') === 'host' ? -(this.parentElement?.scrollTop ?? 0) : 0;
      return { top, bottom: top + height, left: 0, right: width, width, height, x: 0, y: top, toJSON() {} };
    });
    jest.useFakeTimers();
    const frames = spyOn(globalThis, 'requestAnimationFrame').mockImplementation(() => 1);
    const flushFrame = () => act(() => { frames.mock.calls.at(-1)![0](0); });
    try {
      const view = render(<Harness />);
      const scroller = view.getByTestId('scroll');
      expect(navigation.pageInfo).toEqual({ currentPage: 1, totalPages: 3, visible: false, scrollLeft: 0 });
      scroller.scrollTop = 230; scroller.scrollLeft = 12;
      if (event === 'width change') width = 180;
      if (event === 'height change') height = 80;
      if (event === 'window resize') { fireEvent(window, new Event('resize')); fireEvent(window, new Event('resize')); }
      else fireEvent.scroll(scroller);
      expect(frames).toHaveBeenCalledTimes(1);
      flushFrame();
      expect(navigation.pageInfo).toEqual({ currentPage: 3, totalPages: 3, visible: event === 'plain scroll', scrollLeft: 12 });
      fireEvent.scroll(scroller);
      flushFrame();
      expect(navigation.pageInfo.visible).toBe(true);
      const shown = navigation.pageInfo;
      const geometry = viewStore.getSnapshot().geometry!;
      act(() => { viewStore.installFrameGeometry({ ...geometry, frameEpoch: geometry.frameEpoch + 1 }); });
      expect(navigation.pageInfo).toBe(shown);
      act(() => { jest.advanceTimersByTime(300); });
      if (event === 'width change') { width = 160; fireEvent.scroll(scroller); }
      else if (event === 'height change') { height = 100; fireEvent.scroll(scroller); }
      else fireEvent(window, new Event('resize'));
      flushFrame();
      expect(navigation.pageInfo.visible).toBe(true);
      act(() => { jest.advanceTimersByTime(299); });
      expect(navigation.pageInfo.visible).toBe(true);
      act(() => { jest.advanceTimersByTime(1); });
      expect(navigation.pageInfo.visible).toBe(false);
    } finally { cleanup(); frames.mockRestore(); jest.useRealTimers(); box.mockRestore(); controller.dispose(); viewStore.dispose(); }
  });
}

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
      else await expect(controller.ref.scrollToComment(7)).rejects.toThrow('Document worker failed: Worker failed');
    } finally { client.destroy(); dispose(); }
  });
}

const collapsedTargets = { [JSON.stringify({ kind: 'comment', id: 7 })]: { id: 'native',
  ranges: [{ story: 'body', start: { paraId: 'p', offset: 3 }, end: { paraId: 'p', offset: 3 } }] } };

function presentWithForeignCommand(navigation: ReturnType<typeof reviewNavigation>, at: { paraId: string; offset: number }, version = 1,
  editState?: PeerlessViewSnapshot['editState']) {
  const { submit, viewStore } = navigation;
  submit.mockImplementationOnce(() => ({ status: 'submitted', ticket: { seq: 1, ack: Promise.resolve({ status: 'applied', seq: 1 }),
    presented: Promise.resolve().then(() => {
      const view = viewStore.getSnapshot();
      viewStore.acceptWorkerState(view.identity, { epoch: view.identity.epoch, version, throughSeq: 2, presentedThroughSeq: 2,
        selection: { anchor: { story: 'body', ...at }, focus: { story: 'body', ...at } }, editState: editState ?? view.editState });
      viewStore.installFrameGeometry({ ...view.geometry!, frameEpoch: view.geometry!.frameEpoch + 1, coversThroughSeq: 2,
        caret: { x: 10, y: 600, width: 1, height: 12 } });
      return 'presented';
    }) } }));
}

test('review navigation reveals its selection when a command that keeps it is presented in the same frame', async () => {
  const navigation = reviewNavigation(collapsedTargets);
  presentWithForeignCommand(navigation, { paraId: 'p', offset: 3 });
  try {
    expect(await navigation.controller.ref.scrollToComment(7)).toBe(true);
    expect(navigation.submit).toHaveBeenCalledTimes(1);
    expect(navigation.scroller.scrollTop).toBe(600 * 2 - 24);
  } finally { navigation.dispose(); }
});

for (const [name, at, version] of [
  ['another select replaced the selection', { paraId: 'z', offset: 0 }, 1],
  ['a later command moved the caret', { paraId: 'p', offset: 5 }, 1],
  ['a later edit changed the paragraph under the same offsets', { paraId: 'p', offset: 3 }, 2],
] as const) {
  test(`review navigation resolves false without scrolling when ${name}`, async () => {
    const navigation = reviewNavigation(collapsedTargets);
    presentWithForeignCommand(navigation, at, version);
    try {
      expect(await navigation.controller.ref.scrollToComment(7)).toBe(false);
      expect(navigation.scroller.scrollTop).toBe(0);
    } finally { navigation.dispose(); }
  });
}

test('review navigation resolves false without scrolling when a table cell selection replaced its selection', async () => {
  const navigation = reviewNavigation(collapsedTargets);
  presentWithForeignCommand(navigation, { paraId: 'p', offset: 3 }, 1, { readOnly: false, toolbar: null, effectiveFormatting: null,
    storedMarks: null, canUndo: false, canRedo: false, hyperlink: null, tableContext: { story: 'body', tableIndex: 0, rowIndex: 0,
      columnIndex: 0, rowCount: 2, columnCount: 3, hasMultiCellSelection: true, canSplitCell: true } });
  try {
    expect(await navigation.controller.ref.scrollToComment(7)).toBe(false);
    expect(navigation.scroller.scrollTop).toBe(0);
  } finally { navigation.dispose(); }
});

test('review navigation scrolls a range start to where the final frame shows it after a relayout', async () => {
  const range = { story: 'body', start: { paraId: 'p', offset: 3 }, end: { paraId: 'q', offset: 2 } };
  const navigation = reviewNavigation({ [JSON.stringify({ kind: 'comment', id: 7 })]: { id: 'native', ranges: [range] } });
  const { submit, viewStore } = navigation;
  submit.mockImplementationOnce(() => ({ status: 'submitted', ticket: { seq: 1, ack: Promise.resolve({ status: 'applied', seq: 1 }),
    presented: Promise.resolve('presented') } }));
  submit.mockImplementationOnce(() => ({ status: 'submitted', ticket: { seq: 2, ack: Promise.resolve({ status: 'noop', seq: 2 }),
    presented: Promise.resolve().then(() => {
      const view = viewStore.getSnapshot();
      viewStore.acceptWorkerState(view.identity, { epoch: view.identity.epoch, version: 1, throughSeq: 3, presentedThroughSeq: 3,
        selection: { anchor: { story: 'body', ...range.end }, focus: { story: 'body', ...range.start } }, editState: view.editState });
      viewStore.installFrameGeometry({ ...view.geometry!, frameEpoch: view.geometry!.frameEpoch + 1,
        layoutRevision: view.geometry!.layoutRevision + 1, coversThroughSeq: 3, caret: { x: 10, y: 700, width: 1, height: 12 } });
      return 'presented';
    }) } }));
  try {
    expect(await navigation.controller.ref.scrollToComment(7)).toBe(true);
    expect(submit).toHaveBeenCalledTimes(2);
    expect(navigation.scroller.scrollTop).toBe(700 * 2 - 24);
  } finally { navigation.dispose(); }
});

function revealRange(throughSeq: number, final: (first: PeerlessFrameGeometry) => Partial<PeerlessFrameGeometry>, rasterScale?: number) {
  const navigation = reviewNavigation(reviewTargets);
  const { submit, viewStore } = navigation;
  const initial = viewStore.getSnapshot().geometry!;
  viewStore.installFrameGeometry({ ...initial, frameEpoch: initial.frameEpoch + 1, layoutEpoch: 4, rasterScale: 2, documentWidth: 816 });
  submit.mockImplementationOnce(() => ({ status: 'submitted', ticket: { seq: 1, ack: Promise.resolve({ status: 'applied', seq: 1 }),
    presented: Promise.resolve().then(() => {
      const view = viewStore.getSnapshot();
      viewStore.acceptWorkerState(view.identity, { epoch: view.identity.epoch, version: 1, throughSeq: 1, presentedThroughSeq: 1,
        selection: { anchor: { story: 'body', ...reviewRange.start }, focus: { story: 'body', ...reviewRange.start } }, editState: view.editState });
      viewStore.installFrameGeometry({ ...view.geometry!, frameEpoch: view.geometry!.frameEpoch + 1, coversThroughSeq: 1 });
      return 'presented';
    }) } }));
  submit.mockImplementationOnce(() => ({ status: 'submitted', ticket: { seq: 2, ack: Promise.resolve({ status: 'applied', seq: 2 }),
    presented: Promise.resolve().then(() => {
      const view = viewStore.getSnapshot();
      const first = view.geometry!;
      expect(first.coversThroughSeq).toBe(1);
      viewStore.acceptWorkerState(view.identity, { epoch: view.identity.epoch, version: 1, throughSeq, presentedThroughSeq: throughSeq,
        selection: { anchor: { story: 'body', ...reviewRange.end }, focus: { story: 'body', ...reviewRange.start } }, editState: view.editState });
      viewStore.installFrameGeometry({ ...first, frameEpoch: first.frameEpoch + 1, coversThroughSeq: throughSeq,
        caret: null, selectionFocus: null, ...final(first) });
      if (rasterScale !== undefined) viewStore.rescaleFrameGeometry!({ ...viewStore.getSnapshot().geometry!, rasterScale });
      return 'presented';
    }) } }));
  return navigation;
}

test('review navigation scrolls a range start to the selection focus the final frame shows after a relayout', async () => {
  const navigation = revealRange(2, (first) => ({ layoutRevision: first.layoutRevision + 1, layoutEpoch: 5,
    selectionFocus: { x: 10, y: 700, width: 1, height: 12 } }));
  try {
    expect(await navigation.controller.ref.scrollToComment(7)).toBe(true);
    expect(navigation.submit).toHaveBeenCalledTimes(2);
    expect(navigation.scroller.scrollTop).toBe(700 * 2 - 24);
  } finally { navigation.dispose(); }
});

for (const [name, throughSeq, final, rasterScale] of [
  ['a later command repaginated without a new layout revision', 3, { layoutEpoch: 5, layoutProvisional: true }, undefined],
  ['the raster scale changed', 2, { layoutProvisional: true }, 3],
] as const) {
  test(`review navigation resolves false without scrolling when the final frame has no range geometry and ${name}`, async () => {
    const navigation = revealRange(throughSeq, () => final, rasterScale);
    try {
      expect(await navigation.controller.ref.scrollToComment(7)).toBe(false);
      expect(navigation.submit).toHaveBeenCalledTimes(2);
      expect(navigation.viewStore.getSnapshot().geometry!.coversThroughSeq).toBe(throughSeq);
      expect(navigation.scroller.scrollTop).toBe(0);
    } finally { navigation.dispose(); }
  });
}

for (const [name, throughSeq, final, y] of [
  ['a later command kept the layout and the final frame has no range geometry', 3, { layoutProvisional: true }, 400],
  ['the final frame covers only its own selection and has no range geometry', 2, { layoutProvisional: true }, 400],
  ['the final frame covers only its own selection and shows its focus', 2, { selectionFocus: { x: 10, y: 650, width: 1, height: 12 } }, 650],
] as const) {
  test(`review navigation scrolls a range start when ${name}`, async () => {
    const navigation = revealRange(throughSeq, () => final);
    try {
      expect(await navigation.controller.ref.scrollToComment(7)).toBe(true);
      expect(navigation.submit).toHaveBeenCalledTimes(2);
      expect(navigation.scroller.scrollTop).toBe(y * 2 - 24);
    } finally { navigation.dispose(); }
  });
}
