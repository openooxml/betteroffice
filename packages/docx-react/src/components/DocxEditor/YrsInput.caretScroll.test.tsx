import { GlobalRegistrator } from '@happy-dom/global-registrator';
import { afterAll, afterEach, beforeAll, expect, mock, spyOn, test } from 'bun:test';
import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';
import { createRef, useEffect, useMemo } from 'react';
import { resolveDisplayPageClientRect, type DisplayListQueries } from '@betteroffice/docx/layout/render';
import { LayoutSelectionGate, type ResidentMeasurementConfig } from '@betteroffice/docx/layout';
import type { Layout } from '@betteroffice/docx/layout/pagination';
import { findVerticalScrollParentOrRoot } from '@betteroffice/docx/utils/findVerticalScrollParent';
import { preloadEditWasm } from '@betteroffice/docx/wasm/edit';
import {
  createYrsInputPositionMap,
  createYrsSession,
  displayPositionToYrsLoc,
  yrsLocToDisplayPosition,
  type YrsRenderEnv,
  type YrsSession,
} from '@betteroffice/docx/yrs';
import { YrsInput, type YrsInputProps, type YrsInputRef } from './YrsInput';
import { deferWorkerOpenReplica } from './internals/workerOpenReplica';
import { useLayoutPipeline, type UseLayoutPipelineReturn } from './hooks/useLayoutPipeline';
import { useLayoutTriggers } from './hooks/useLayoutTriggers';
import { useRevisionPreview } from './hooks/useRevisionPreview';
import { scrollIntoViewDelta, scrollViewport } from './internals/viewportBand';

const ownsDom = !GlobalRegistrator.isRegistered;
if (ownsDom) GlobalRegistrator.register();
const { act, cleanup, fireEvent, render } = await import('@testing-library/react');
const ROOT = resolve(import.meta.dir, '../../../../..');
const bytes = new Uint8Array(
  readFileSync(resolve(ROOT, 'crates/docx-edit/tests/fixtures/page-fragments/pages.docx'))
);
const sessions: YrsSession[] = [];
const scrollers: HTMLDivElement[] = [];
const restoreMocks: Array<() => void> = [];

beforeAll(() => {
  if (!window.document.fonts) {
    Object.defineProperty(window.document, 'fonts', {
      value: { addEventListener: () => {}, removeEventListener: () => {} },
      configurable: true,
    });
  }
  return preloadEditWasm(
    new Uint8Array(
      readFileSync(resolve(ROOT, 'packages/docx/src/wasm/generated/edit/docx_edit_bg.wasm'))
    )
  );
});
afterEach(() => {
  cleanup();
  for (const restore of restoreMocks.splice(0)) restore();
  for (const scroller of scrollers.splice(0)) scroller.remove();
  for (const session of sessions.splice(0)) session.destroy();
});
afterAll(async () => {
  if (ownsDom) await GlobalRegistrator.unregister();
});

async function mount(
  readOnly: boolean,
  seedSelection: boolean,
  replicaReady = true,
  props: Partial<YrsInputProps> = {}
) {
  const session = await createYrsSession();
  sessions.push(session);
  if (replicaReady) session.openDocx(bytes, true);
  expect(session.selection()).toBeNull();

  const scroller = document.createElement('div');
  scroller.style.overflowY = 'auto';
  scroller.style.height = '400px';
  Object.defineProperties(scroller, {
    clientHeight: { value: 400 },
    scrollHeight: { value: 3000 },
  });
  scroller.getBoundingClientRect = () => new DOMRect(0, 0, 800, 400);
  scroller.scrollTop = 100;
  const scrollTop = scroller.scrollTop;
  const host = document.createElement('div');
  host.className = 'canvas-pages';
  host.style.height = '3000px';
  const page = document.createElement('div');
  page.className = 'canvas-page';
  page.dataset.pageIndex = '0';
  page.getBoundingClientRect = () => new DOMRect(0, -scroller.scrollTop, 800, 3000);
  host.append(page);
  scroller.append(host);
  document.body.append(scroller);
  scrollers.push(scroller);

  const input = createRef<YrsInputRef>();
  const replicaReadyRef = { current: replicaReady };
  const canvasHostRef = { current: host };
  const onStateChange = mock<YrsInputProps['onStateChange']>(() => {});
  const caretRect = mock((position: number) => ({
    pageIndex: 0,
    x: 80,
    y: 1200 + position * 80,
    width: 1,
    height: 20,
  }));
  const queries = {
    isReady: () => replicaReadyRef.current,
    caretRect,
    anchorRect: caretRect,
    pageCount: () => 1,
    pageSize: () => ({ width: 800, height: 3000 }),
    visualLines: () => [],
    visualLinesOnPage: () => [],
    visualLineExtent: () => null,
    verticalMove: () => null,
  } satisfies Pick<DisplayListQueries,
    'isReady' | 'caretRect' | 'anchorRect' | 'pageCount' | 'pageSize' |
    'visualLines' | 'visualLinesOnPage' | 'visualLineExtent' | 'verticalMove'>;
  const map = () => createYrsInputPositionMap('body', session.paragraphSpans('body'));
  const displayPositionToLoc: YrsInputProps['displayPositionToLoc'] = (position) =>
    displayPositionToYrsLoc(map(), position);
  const locToDisplayPosition: YrsInputProps['locToDisplayPosition'] = (loc) =>
    yrsLocToDisplayPosition(map(), loc);
  const inputFor = (layoutProps: Partial<YrsInputProps> = {}) => (
    <YrsInput
      ref={input}
      enabled
      readOnly={readOnly}
      seedSelection={seedSelection}
      replicaReadyRef={replicaReadyRef}
      session={session}
      inputPositionMap={map}
      displayPositionToLoc={displayPositionToLoc}
      locToDisplayPosition={locToDisplayPosition}
      displayListQueries={queries as unknown as DisplayListQueries}
      displayListFrameEpoch={replicaReadyRef.current ? 1 : 0}
      layoutUpdateOrigin="remote"
      canvasHostRef={canvasHostRef}
      onStateChange={onStateChange}
      onDirectInput={() => {}}
      {...props}
      {...layoutProps}
    />
  );
  const view = render(inputFor());
  const expectCaret = (offset: number) => {
    const loc = { story: 'body', paraId: session.paragraphs('body')[0]!.paraId, offset };
    expect(session.selection()).toEqual({ anchor: loc, head: loc });
    expect(input.current!.displaySelection()).toEqual({ anchor: offset + 1, head: offset + 1 });
    expect(caretRect).toHaveBeenCalledWith(offset + 1);
  };
  return {
    session, input, scroller, scrollTop, replicaReadyRef, onStateChange, view, inputFor, expectCaret,
    queries, canvasHostRef,
  };
}

async function mountWithPipeline(workerLayout: boolean) {
  let nextFrame = 0;
  const frames = new Map<number, FrameRequestCallback>();
  const requestFrame = spyOn(globalThis, 'requestAnimationFrame').mockImplementation((callback) => {
    const id = ++nextFrame;
    frames.set(id, callback);
    return id;
  });
  const cancelFrame = spyOn(globalThis, 'cancelAnimationFrame').mockImplementation((id) => {
    frames.delete(id);
  });
  restoreMocks.push(() => { requestFrame.mockRestore(); cancelFrame.mockRestore(); });
  const frame = async () => {
    await act(async () => {
      for (const [id, callback] of [...frames]) {
        if (frames.delete(id)) callback(performance.now());
      }
    });
    expect(frames.size).toBe(0);
  };
  const mounted = await mount(false, true);
  const { session, input, scroller, view, inputFor, queries, canvasHostRef, onStateChange } = mounted;
  let currentScrollTop = scroller.scrollTop;
  const scrollChanges: number[] = [];
  Object.defineProperty(scroller, 'scrollTop', {
    configurable: true,
    get: () => currentScrollTop,
    set: (value: number) => {
      if (value !== currentScrollTop) scrollChanges.push(value);
      currentScrollTop = value;
    },
  });
  const displayQueries = {
    ...queries,
    caretRect: mock((position: number) => queries.caretRect(position)),
  };
  const computation = () => ({ layout: { pages: [] } as unknown as Layout, notesConverged: true });
  const retainedLayout = spyOn(session, 'layoutDocumentWithRegionsRetainedJson')
    .mockImplementation(() => JSON.stringify(computation()));
  const probe = spyOn(session, 'residentWorkerProbe').mockReturnValue(
    { layoutRevision: 1 } as ReturnType<YrsSession['residentWorkerProbe']>
  );
  restoreMocks.push(() => { retainedLayout.mockRestore(); probe.mockRestore(); });
  const layoutInWorker = mock(() => workerLayout ? Promise.resolve(computation()) : null);
  const syncCoordinator = new LayoutSelectionGate();
  const measurement: ResidentMeasurementConfig = {
    fontChains: {},
    defaults: { fontSize: 22, fontFamily: 'Calibri' },
    compat: { noLeading: false, doNotExpandShiftReturn: false },
    authoritativeShaping: true,
  };
  let pipeline!: UseLayoutPipelineReturn;
  let frameEpoch = 1;
  const published: Array<{
    layout: Layout;
    origin: UseLayoutPipelineReturn['layoutUpdateOrigin'];
    frameEpoch: number;
  }> = [];
  const onLayoutComputed = (layout: Layout | null) => {
    if (layout) published.push({ layout, origin: pipeline.layoutUpdateOrigin, frameEpoch });
  };
  const getScrollContainer = () => scroller;
  const getSelectionHead = () => input.current?.displaySelection()?.head ?? 0;
  const residentMeasurementConfig = () => measurement;
  const deferLayoutPass = () => false;
  const viewportLayoutRef = { current: null };
  const updateSelectionOverlay = () => {};
  const onError = mock(() => {});
  const handleStateChange: YrsInputProps['onStateChange'] = (...args) => {
    onStateChange(...args);
    if (args[1] && !args[2]) {
      syncCoordinator.incrementStateSeq();
      syncCoordinator.requestRender();
      pipeline.scheduleLayout('local');
    }
  };
  function PipelineInput() {
    const preview = useRevisionPreview(session);
    const renderEnv = useMemo<YrsRenderEnv>(() => ({
      ...(preview.revisionPreview ? { revisionPreview: preview.revisionPreview } : {}),
    }), [preview]);
    pipeline = useLayoutPipeline({
      document: null,
      session,
      renderEnv,
      pageGap: 24,
      zoom: 1,
      residentMeasurementConfig,
      deferLayoutPass,
      displayListQueries: displayQueries as unknown as DisplayListQueries,
      pagesContainerRef: canvasHostRef,
      viewportLayoutRef,
      getSelectionHead,
      syncCoordinator,
      getScrollContainer,
      layoutInWorker,
      onLayoutComputed,
      onError,
    });
    useLayoutTriggers({
      runLayoutPipeline: pipeline.runLayoutPipeline,
      updateSelectionOverlay,
      renderEnv,
    });
    useEffect(() => pipeline.runLayoutPipeline(), [session]);
    const displayListFrameEpoch = useMemo(() => ++frameEpoch, [pipeline.layout]);
    return inputFor({
      displayListQueries: displayQueries as unknown as DisplayListQueries,
      layoutUpdateOrigin: pipeline.layoutUpdateOrigin,
      displayListFrameEpoch,
      onStateChange: handleStateChange,
    });
  }
  await act(async () => view.rerender(<PipelineInput />));
  await frame();
  expect(pipeline.layout).not.toBeNull();
  expect(onError).not.toHaveBeenCalled();
  return {
    ...mounted, queries: displayQueries, pipeline: () => pipeline, layoutInWorker, retainedLayout,
    frame, frameEpoch: () => frameEpoch, published, scrollChanges, onError,
  };
}

function offscreenCaretDelta(t: Awaited<ReturnType<typeof mount>>) {
  const queries = t.queries as unknown as DisplayListQueries;
  expect(queries.isReady()).toBe(true);
  const selection = t.input.current!.displaySelection();
  expect(selection).not.toBeNull();
  expect(selection!.anchor).toBe(selection!.head);
  const caret = queries.caretRect(selection!.head);
  expect(caret).not.toBeNull();
  const pageRect = resolveDisplayPageClientRect(
    t.canvasHostRef.current, queries, caret!.pageIndex
  );
  const pageSize = queries.pageSize(caret!.pageIndex);
  expect(pageRect).not.toBeNull();
  expect(pageSize).not.toBeNull();
  expect(pageSize!.height).toBeGreaterThan(0);
  expect(findVerticalScrollParentOrRoot(t.canvasHostRef.current)).toBe(t.scroller);
  const scaleY = pageRect!.height / pageSize!.height;
  const top = pageRect!.top + caret!.y * scaleY;
  const bottom = top + Math.max(1, caret!.height * scaleY);
  const viewport = scrollViewport(t.scroller);
  expect(top < viewport.top || bottom > viewport.bottom).toBe(true);
  const delta = scrollIntoViewDelta(viewport, top, bottom, 24);
  expect(delta).not.toBe(0);
  return delta;
}

test.each([false, true])(
  'proposal accept, reject and undo preserve scrollTop with an unchanged distant caret (worker layout=%s)',
  async (workerLayout) => {
    const t = await mountWithPipeline(workerLayout);
    const firstText = t.session.paragraphs('body')[0]!.text;
    await act(async () => {
      fireEvent.input(t.view.getByTestId('yrs-input'), { target: { value: '12345678' } });
      await t.input.current!.flushPendingInput();
    });
    await t.frame();
    expect(t.session.paragraphs('body')[0]!.text).toBe(`12345678${firstText}`);
    t.expectCaret(8);
    expect(t.pipeline().layoutUpdateOrigin).toBe('local');
    const last = t.session.paragraphs('body').at(-1)!;
    await act(async () => {
      const proposal = t.session.proposeChanges({
        expectVersion: t.session.version(),
        proposals: [{
          id: 'distant',
          paragraph: {
            kind: 'session',
            sessionId: t.session.paragraphIdentities().sessionId,
            story: 'body',
            paraId: last.paraId,
          },
          suggest: { author: 'Assistant', date: '2026-09-29T00:00:00Z' },
          op: 'insertText',
          at: 'end',
          text: ' proposal',
        }],
      });
      expect(proposal.ok).toBe(true);
    });
    await t.frame();
    const selection = t.session.selection();
    expect(selection).not.toBeNull();
    expect(selection!.anchor).toEqual(selection!.head);
    for (const [index, state] of (['accepted', 'proposed', 'rejected', 'proposed'] as const).entries()) {
      const scrollTop = 2300;
      t.scroller.scrollTop = scrollTop;
      fireEvent.scroll(t.scroller);
      expect(offscreenCaretDelta(t)).toBeLessThan(0);
      t.scrollChanges.length = 0;
      const layout = t.pipeline().layout;
      const frameEpoch = t.frameEpoch();
      const publications = t.published.length;
      const caretQueries = t.queries.caretRect.mock.calls.length;
      const version = t.session.version();
      const previewVersion = t.session.getProposals().previewVersion;
      const workerPasses = t.layoutInWorker.mock.calls.length;
      const hostPasses = t.retainedLayout.mock.calls.length;
      await act(async () => {
        const result = t.session.setProposalStates({
          expectVersion: t.session.version(),
          expectPreviewVersion: t.session.getProposals().previewVersion,
          changes: [{ id: 'distant', state }],
        });
        expect(result.ok).toBe(true);
      });
      await t.frame();
      expect(t.session.version()).toBe(version);
      expect(t.session.getProposals().previewVersion).toBe(previewVersion + 1);
      expect(t.pipeline().layout).not.toBe(layout);
      expect(t.published).toHaveLength(publications + 1);
      expect(t.published.at(-1)!.layout).toBe(t.pipeline().layout!);
      expect(t.published.at(-1)!.frameEpoch).toBeGreaterThan(frameEpoch);
      expect(t.queries.caretRect.mock.calls.length).toBeGreaterThan(caretQueries);
      expect(t.layoutInWorker).toHaveBeenCalledTimes(workerPasses + (index === 0 ? 0 : 1));
      expect(t.retainedLayout).toHaveBeenCalledTimes(hostPasses + (workerLayout && index > 0 ? 0 : 1));
      expect(t.session.selection()).toEqual(selection);
      expect(t.scrollChanges).toEqual([]);
      expect(t.scroller.scrollTop).toBe(scrollTop);
      expect(t.published.at(-1)!.origin).toBe('remote');
      expect(t.pipeline().layoutUpdateOrigin).toBe('remote');
      expect(t.onError).not.toHaveBeenCalled();
    }
  }
);

test('a delayed unchanged selection reveal preserves newer scrolling on a passive frame', async () => {
  const t = await mount(false, true);
  t.expectCaret(0);
  const selection = t.session.selection();
  const position = t.input.current!.displaySelection()!.head;
  t.scroller.scrollTop = 2300;
  expect(offscreenCaretDelta(t)).toBeLessThan(0);
  const ready = spyOn(t.queries, 'isReady').mockReturnValue(false);
  restoreMocks.push(() => ready.mockRestore());
  const caretQueries = t.queries.caretRect.mock.calls.length;
  act(() => t.input.current!.setSelectionFromDisplay(position));
  expect(ready).toHaveBeenCalled();
  expect(t.session.selection()).toEqual(selection);
  expect(t.queries.caretRect.mock.calls).toHaveLength(caretQueries);
  expect(t.scroller.scrollTop).toBe(2300);
  t.scroller.scrollTop = 2000;
  fireEvent.scroll(t.scroller);
  ready.mockReturnValue(true);
  expect(offscreenCaretDelta(t)).toBeLessThan(0);
  const passiveCaretQueries = t.queries.caretRect.mock.calls.length;
  act(() => t.view.rerender(t.inputFor({ displayListFrameEpoch: 2 })));
  expect(t.session.selection()).toEqual(selection);
  expect(t.queries.caretRect.mock.calls.length).toBeGreaterThan(passiveCaretQueries);
  expect(t.scroller.scrollTop).toBe(2000);
});

test.each(['Ctrl+Home', 'Home', 'ArrowLeft', 'ArrowUp', 'setSelectionFromDisplay'])(
  '%s reveals an unchanged caret after a passive relayout',
  async (action) => {
    const t = await mountWithPipeline(true);
    t.expectCaret(0);
    const selection = t.session.selection();
    const verticalMove = spyOn(t.queries as unknown as DisplayListQueries, 'verticalMove')
      .mockReturnValue({ position: t.input.current!.displaySelection()!.head, goalX: 80 });
    restoreMocks.push(() => verticalMove.mockRestore());
    expect(t.pipeline().layoutUpdateOrigin).toBe('remote');
    t.scroller.scrollTop = 2300;
    fireEvent.scroll(t.scroller);
    const delta = offscreenCaretDelta(t);
    expect(delta).toBeLessThan(0);
    t.scrollChanges.length = 0;
    const layout = t.pipeline().layout;
    const frameEpoch = t.frameEpoch();
    await act(async () => t.pipeline().runLayoutPipeline());
    await t.frame();
    expect(t.pipeline().layout).not.toBe(layout);
    expect(t.frameEpoch()).toBeGreaterThan(frameEpoch);
    expect(t.pipeline().layoutUpdateOrigin).toBe('remote');
    expect(t.session.selection()).toEqual(selection);
    expect(t.scrollChanges).toEqual([]);
    expect(t.scroller.scrollTop).toBe(2300);
    await act(async () => {
      if (action === 'setSelectionFromDisplay') {
        t.input.current!.setSelectionFromDisplay(t.input.current!.displaySelection()!.head);
      } else {
        fireEvent.keyDown(t.view.getByTestId('yrs-input'), {
          key: action === 'Ctrl+Home' ? 'Home' : action,
          ctrlKey: action === 'Ctrl+Home',
        });
      }
      await t.input.current!.flushPendingInput();
    });
    await t.frame();
    expect(t.session.selection()).toEqual(selection);
    expect(t.scroller.scrollTop).toBe(2300 + delta);
    expect(t.scrollChanges).toEqual([2300 + delta]);
    const caret = t.queries.caretRect(t.input.current!.displaySelection()!.head);
    const page = resolveDisplayPageClientRect(
      t.canvasHostRef.current, t.queries as unknown as DisplayListQueries, caret.pageIndex
    )!;
    expect(page.top + caret.y).toBeGreaterThanOrEqual(24);
    expect(page.top + caret.y + caret.height).toBeLessThanOrEqual(t.scroller.clientHeight - 24);
    t.scroller.scrollTop = 2300;
    fireEvent.scroll(t.scroller);
    t.scrollChanges.length = 0;
    await act(async () => t.pipeline().runLayoutPipeline());
    await t.frame();
    expect(t.scrollChanges).toEqual([]);
    expect(t.scroller.scrollTop).toBe(2300);
    expect(t.onError).not.toHaveBeenCalled();
  }
);

test('a local Delete reveals a distant caret even when its collapsed selection is unchanged', async () => {
  const t = await mountWithPipeline(true);
  const selection = t.session.selection();
  t.scroller.scrollTop = t.scrollTop;
  fireEvent.scroll(t.scroller);
  const text = t.session.paragraphs('body')[0]!.text;
  const layout = t.pipeline().layout;
  const frameEpoch = t.frameEpoch();
  const publications = t.published.length;
  const delta = offscreenCaretDelta(t);
  expect(delta).toBeGreaterThan(0);
  t.scrollChanges.length = 0;
  await act(async () => {
    fireEvent.keyDown(t.view.getByTestId('yrs-input'), { key: 'Delete' });
    await t.input.current!.flushPendingInput();
  });
  const caretQueries = t.queries.caretRect.mock.calls.length;
  await t.frame();
  expect(t.session.paragraphs('body')[0]!.text).toBe(text.slice(1));
  expect(t.session.selection()).toEqual(selection);
  expect(t.pipeline().layout).not.toBe(layout);
  expect(t.published).toHaveLength(publications + 1);
  expect(t.published.at(-1)!.layout).toBe(t.pipeline().layout!);
  expect(t.published.at(-1)!.frameEpoch).toBeGreaterThan(frameEpoch);
  expect(t.queries.caretRect.mock.calls.length).toBeGreaterThan(caretQueries);
  expect(t.pipeline().layoutUpdateOrigin).toBe('local');
  expect(t.published.at(-1)!.origin).toBe('local');
  expect(t.scroller.scrollTop).toBeGreaterThan(t.scrollTop);
  expect(t.scroller.scrollTop).toBe(t.scrollTop + delta);
  expect(t.scrollChanges).toEqual([t.scrollTop + delta]);
  expect(t.onError).not.toHaveBeenCalled();
});

test('a resident Delete frame reveals an unchanged distant caret after a remote layout', async () => {
  let session!: YrsSession;
  const applyResidentDelete = mock(async () => {
    const { head } = session.selection()!;
    session.deleteRange({
      story: head.story,
      start: { paraId: head.paraId, offset: head.offset },
      end: { paraId: head.paraId, offset: head.offset + 1 },
    });
    return { frameEpoch: 2, caretSynchronized: false, deletedUnits: 1 };
  });
  const t = await mount(false, true, true, { applyResidentDelete });
  session = t.session;
  const selection = session.selection();
  const text = session.paragraphs('body')[0]!.text;
  t.scroller.scrollTop = t.scrollTop;
  await act(async () => {
    fireEvent.keyDown(t.view.getByTestId('yrs-input'), { key: 'Delete' });
    await t.input.current!.flushPendingInput();
  });
  expect(applyResidentDelete).toHaveBeenCalledWith('forward', 1);
  expect(session.paragraphs('body')[0]!.text).toBe(text.slice(1));
  expect(session.selection()).toEqual(selection);
  expect(t.scroller.scrollTop).toBe(t.scrollTop);
  act(() => t.view.rerender(t.inputFor({ displayListFrameEpoch: 2 })));
  expect(t.scroller.scrollTop).toBeGreaterThan(t.scrollTop);
  t.scroller.scrollTop = t.scrollTop;
  act(() => t.view.rerender(t.inputFor({ displayListFrameEpoch: 3 })));
  expect(t.scroller.scrollTop).toBe(t.scrollTop);
});

test('a read-only replica lands without a selection or scrolling when seeding is disabled', async () => {
  const { session, input, scroller, scrollTop, replicaReadyRef, onStateChange, view, inputFor } =
    await mount(true, false, false);
  expect(session.selection()).toBeNull();
  expect(scroller.scrollTop).toBe(scrollTop);
  expect(onStateChange).not.toHaveBeenCalled();
  act(() => {
    session.openDocx(bytes, true);
    replicaReadyRef.current = true;
    view.rerender(inputFor());
  });
  expect(input.current!.displaySelection()).toBeNull();
  expect(session.selection()).toBeNull();
  expect(scroller.scrollTop).toBe(scrollTop);
  expect(onStateChange).not.toHaveBeenCalled();
});

test.each([
  ['read-only', true],
  ['editable', false],
] as const)('a %s input seeds a caret at the story start and scrolls to it', async (_, readOnly) => {
  const { scroller, scrollTop, expectCaret } = await mount(readOnly, true);
  expectCaret(0);
  expect(scroller.scrollTop).toBeGreaterThan(scrollTop);
});

test('keepSelectionInPlace suppresses scrolling for one selection', async () => {
  const { input, scroller, expectCaret } = await mount(false, true);
  expectCaret(0);
  const scrollTop = scroller.scrollTop;
  act(() => {
    input.current!.setSelectionFromDisplay(2);
    input.current!.keepSelectionInPlace();
  });
  expectCaret(1);
  expect(scroller.scrollTop).toBe(scrollTop);
  act(() => input.current!.setSelectionFromDisplay(3));
  expectCaret(2);
  expect(scroller.scrollTop).toBeGreaterThan(scrollTop);
});

test('read-only, a first selection set on a replica without a caret scrolls into view', async () => {
  const { session, input, scroller, scrollTop, expectCaret } = await mount(true, false);
  expect(session.selection()).toBeNull();
  act(() => input.current!.setSelectionFromDisplay(2));
  expectCaret(1);
  expect(scroller.scrollTop).toBeGreaterThan(scrollTop);
});

test('read-only, the first keyboard move on a replica without a caret scrolls into view', async () => {
  const { session, input, scroller, scrollTop, view } = await mount(true, false);
  expect(session.selection()).toBeNull();
  const textarea = view.getByTestId('yrs-input');
  await act(async () => {
    fireEvent.keyDown(textarea, { key: 'End', ctrlKey: true });
    await input.current!.flushPendingInput();
  });
  expect(session.selection()).not.toBeNull();
  expect(scroller.scrollTop).toBeGreaterThan(scrollTop);
});

// Recorded gestures replay before the input that waited for the edit peer.
async function mountPendingReplica() {
  let epoch = 0;
  let pending: [number, number] | null = null;
  const replayed: Array<[number, number]> = [];
  const mounted = await mount(true, false, false, {
    inputEpoch: () => epoch,
    applyPendingSelection: () => {
      if (!pending) return;
      const [anchor, head] = pending;
      pending = null;
      replayed.push([anchor, head]);
      mounted.input.current!.setSelectionFromDisplay(anchor, head);
      mounted.input.current!.keepSelectionInPlace();
    },
  });
  const { session, input, view, inputFor, replicaReadyRef } = mounted;
  let settle!: (loaded: boolean) => void;
  const hydrate = mock(() =>
    new Promise<() => void>((resolve, reject) => {
      settle = (loaded) =>
        loaded
          ? resolve(() => session.openDocx(bytes, true))
          : reject(new Error('The handoff failed'));
    })
  );
  const replica = deferWorkerOpenReplica(
    session,
    hydrate,
    () => {
      throw new Error('The fallback failed');
    },
    () => {
      replicaReadyRef.current = true;
    }
  );
  const finish = async (loaded: boolean) => {
    await act(async () => {
      replica.start();
      settle(loaded);
      await input.current!.flushPendingInput();
    });
    act(() => view.rerender(inputFor()));
  };
  return {
    ...mounted,
    hydrate,
    replayed,
    textarea: view.getByTestId('yrs-input'),
    record: (anchor: number, head = anchor) => {
      pending = [anchor, head];
    },
    supersede: () => {
      epoch += 1;
      pending = null;
    },
    loc: (offset: number) => ({ story: 'body', paraId: session.paragraphs('body')[0]!.paraId, offset }),
    load: () => finish(true),
    fail: () => finish(false),
  };
}

test('keys pressed while the replica loads follow the recorded gesture once it has, in order', async () => {
  const t = await mountPendingReplica();
  t.record(3);
  act(() => {
    fireEvent.keyDown(t.textarea, { key: 'ArrowRight' });
    fireEvent.keyDown(t.textarea, { key: 'ArrowRight', shiftKey: true });
  });
  expect(t.hydrate).not.toHaveBeenCalled();
  expect(t.session.selection()).toBeNull();
  expect(t.input.current!.hasPendingInput()).toBe(true);

  await t.load();
  expect(t.replayed).toEqual([[3, 3]]);
  expect(t.session.selection()).toEqual({ anchor: t.loc(3), head: t.loc(4) });
});

test('a line move pressed while the replica loads moves from the recorded caret', async () => {
  const t = await mountPendingReplica();
  t.record(3);
  act(() => {
    fireEvent.keyDown(t.textarea, { key: 'ArrowDown' });
  });

  await t.load();
  expect(t.replayed).toEqual([[3, 3]]);
  expect(t.session.selection()?.head.paraId).toBe(t.session.paragraphs('body')[1]!.paraId);
});

test('select all pressed after a click while the replica loads selects the whole story', async () => {
  const t = await mountPendingReplica();
  t.record(3);
  act(() => {
    fireEvent.keyDown(t.textarea, { key: 'a', ctrlKey: true });
  });

  await t.load();
  const map = createYrsInputPositionMap('body', t.session.paragraphSpans('body'));
  const first = map.paragraphs[0]!;
  const last = map.paragraphs[map.paragraphs.length - 1]!;
  expect(t.replayed).toEqual([[3, 3]]);
  expect(t.session.selection()).toEqual({
    anchor: { story: 'body', paraId: first.paraId, offset: 0 },
    head: { story: 'body', paraId: last.paraId, offset: last.length },
  });
});

test('a copy while the replica loads writes the text the replayed drag selects', async () => {
  const written: Array<Record<string, Promise<Blob>>> = [];
  const descriptor = Object.getOwnPropertyDescriptor(navigator, 'clipboard');
  const scope = globalThis as { ClipboardItem?: unknown };
  const clipboardItem = scope.ClipboardItem;
  Object.defineProperty(navigator, 'clipboard', {
    configurable: true,
    value: {
      write: async (items: Array<{ data: Record<string, Promise<Blob>> }>) => {
        written.push(items[0]!.data);
      },
    },
  });
  scope.ClipboardItem = class {
    constructor(readonly data: Record<string, Promise<Blob>>) {}
  };
  try {
    const t = await mountPendingReplica();
    t.record(2, 6);
    let notPrevented = true;
    act(() => {
      notPrevented = fireEvent.keyDown(t.textarea, { key: 'c', ctrlKey: true });
    });
    expect(notPrevented).toBe(false);
    expect(written).toHaveLength(1);
    expect(t.hydrate).not.toHaveBeenCalled();

    await t.load();
    const text = t.session.paragraphs('body')[0]!.text.slice(1, 5);
    expect(text).toHaveLength(4);
    expect(await (await written[0]!['text/plain']!).text()).toBe(text);
  } finally {
    if (descriptor) Object.defineProperty(navigator, 'clipboard', descriptor);
    else delete (navigator as { clipboard?: unknown }).clipboard;
    scope.ClipboardItem = clipboardItem;
  }
});

test('newer input drops keys still waiting for the replica', async () => {
  const t = await mountPendingReplica();
  t.record(3);
  act(() => {
    fireEvent.keyDown(t.textarea, { key: 'ArrowRight' });
  });
  t.supersede();
  t.record(5);
  act(() => {
    fireEvent.keyDown(t.textarea, { key: 'ArrowLeft' });
  });

  await t.load();
  expect(t.replayed).toEqual([[5, 5]]);
  expect(t.session.selection()).toEqual({ anchor: t.loc(3), head: t.loc(3) });
});

test.each([false, true])(
  'a key that waited for the replica scrolls to its caret unless the reader scrolled since (%p)',
  async (scrolled) => {
    const t = await mountPendingReplica();
    act(() => {
      fireEvent.keyDown(t.textarea, { key: 'End', ctrlKey: true });
    });
    if (scrolled) {
      act(() => {
        t.scroller.querySelector('.canvas-page')!.dispatchEvent(new Event('wheel', { bubbles: true }));
      });
    }

    await t.load();
    expect(t.session.selection()).not.toBeNull();
    if (scrolled) expect(t.scroller.scrollTop).toBe(t.scrollTop);
    else expect(t.scroller.scrollTop).toBeGreaterThan(t.scrollTop);
  }
);

test('a replica that fails to load drops the keys waiting for it', async () => {
  const t = await mountPendingReplica();
  t.record(3);
  act(() => {
    fireEvent.keyDown(t.textarea, { key: 'a', ctrlKey: true });
  });

  await t.fail();
  expect(t.replayed).toEqual([]);
  expect(t.session.selection()).toBeNull();
  expect(t.input.current!.hasPendingInput()).toBe(false);
});
