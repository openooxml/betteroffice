import { GlobalRegistrator } from '@happy-dom/global-registrator';
import { afterAll, afterEach, expect, mock, test } from 'bun:test';
import { createRef, useRef, type RefObject } from 'react';
import type { DisplayListQueries } from '@betteroffice/docx/layout/render';
import type { Layout } from '@betteroffice/docx/layout/pagination';
import type { DocxDisplaySelectionText, ResidentDocumentRead, ResidentEngineWorkerClient, YrsSession } from '@betteroffice/docx/yrs';
import type { PagedEditorRef } from './PagedEditor';
import { ViewerInput, type ViewerInputProps } from './ViewerInput';
import type { YrsInputRef } from './YrsInput';
import { usePagedEditorRefApi } from './hooks/usePagedEditorRefApi';
import { stampWorkerFrameVersion } from './internals/layoutProvenance';
import type { ViewerSelectionChange } from './internals/viewerSelectionController';
import { layoutIdOf } from '../../plugins/geometry';

const ownsDom = !GlobalRegistrator.isRegistered;
if (ownsDom) GlobalRegistrator.register();
const { act, cleanup, fireEvent, render, renderHook } = await import('@testing-library/react');
const originalClipboard = Object.getOwnPropertyDescriptor(navigator, 'clipboard');
const originalClipboardItem = globalThis.ClipboardItem;

afterEach(() => {
  cleanup();
  if (originalClipboard) Object.defineProperty(navigator, 'clipboard', originalClipboard);
  else Reflect.deleteProperty(navigator, 'clipboard');
  globalThis.ClipboardItem = originalClipboardItem;
});
afterAll(async () => {
  if (ownsDom) await GlobalRegistrator.unregister();
});

for (const viewerSelection of [true, false]) {
  test(`${viewerSelection ? 'viewer' : 'editor'} readiness ${viewerSelection ? 'uses layout' : 'waits for the replica'}`, () => {
    const onReady = mock((_ref: PagedEditorRef) => {});
    const session = {} as YrsSession;
    const layout = {} as Layout;
    const runLayoutPipeline = () => {};
    const scrollToParaIdImpl = () => false;
    const scrollToPageImpl = () => {};
    const hook = renderHook(({ layout, replicaReady }: { layout: Layout | null; replicaReady: boolean }) => {
      const ref = useRef<PagedEditorRef>(null);
      usePagedEditorRefApi({
        ref, viewerSelection, yrsInputRef: { current: null }, layout, yrsSession: session, replicaReady,
        runLayoutPipeline, getLayoutRequest: () => null, readLayoutRequest: async () => null,
        scrollToPositionImpl: () => {}, revealPositionImpl: () => 'layout-unavailable',
        scrollToParaIdImpl, scrollToPageImpl, setIsFocused: () => {},
        onReadyRef: { current: onReady }, documentFromYrs: () => null,
        yrsLocToDisplayPosition: () => null, syncYrsInputState: () => true,
        applyYrsFormatting: () => false, applyYrsCommand: () => false,
        getYrsPositionProjection: () => null, displayPositionToYrsLoc: () => null,
        getPositionAtPoint: () => null,
      });
      return ref;
    }, { initialProps: { layout: null as Layout | null, replicaReady: false } });
    expect(onReady).not.toHaveBeenCalled();
    hook.rerender({ layout, replicaReady: false });
    if (viewerSelection) {
      expect(onReady).toHaveBeenCalledTimes(1);
      expect(onReady.mock.calls[0]![0].isWorkerViewer()).toBe(true);
    } else {
      expect(onReady).not.toHaveBeenCalled();
      hook.rerender({ layout, replicaReady: true });
      expect(onReady).toHaveBeenCalledTimes(1);
      expect(onReady.mock.calls[0]![0].isWorkerViewer()).toBe(false);
    }
  });
}

function frame(version: string, preview = false, asOpened = false): DisplayListQueries {
  const lines = [{ pageIndex: 0, from: 1, to: 30 }];
  const queries = {
    displayList: { pages: [{ pageIndex: 0, primitives: [{ docStart: 1, docEnd: 30 }] }] },
    isReady: () => true,
    paragraphRects: () => [],
    visualLinesOnPage: (pageIndex: number) => lines.filter((line) => line.pageIndex === pageIndex),
    visualLineAtPosition: (position: number) => lines.find((line) => position >= line.from && position <= line.to) ?? null,
  } as unknown as DisplayListQueries;
  stampWorkerFrameVersion(queries, version, preview, asOpened);
  return queries;
}

function text(value: string): DocxDisplaySelectionText {
  return { text: value, range: null };
}

/** A viewer input behind the editor ref, whose gestures begin the way the pointer hook begins them. */
function NavigatingViewer({ props, inputRef, pagedRef }: {
  props: ViewerInputProps;
  inputRef: RefObject<YrsInputRef | null>;
  pagedRef: RefObject<PagedEditorRef | null>;
}) {
  const onReadyRef = useRef<((ref: PagedEditorRef) => void) | undefined>(undefined);
  usePagedEditorRefApi({
    viewerSelection: true,
    bumpInputEpoch: () => { inputRef.current?.beginGesture?.(); },
    ref: pagedRef,
    yrsInputRef: inputRef,
    layout: null,
    runLayoutPipeline: () => {},
    getLayoutRequest: () => null, readLayoutRequest: async () => null,
    scrollToPositionImpl: () => {},
    revealPositionImpl: () => 'layout-unavailable',
    scrollToParaIdImpl: () => true,
    scrollToPageImpl: () => {},
    setIsFocused: () => {},
    onReadyRef,
    documentFromYrs: () => null,
    yrsSession: null,
    yrsLocToDisplayPosition: () => null,
    syncYrsInputState: () => true,
    applyYrsFormatting: () => false,
    applyYrsCommand: () => false,
    getYrsPositionProjection: () => null,
    displayPositionToYrsLoc: () => null,
    getPositionAtPoint: () => null,
    navigateViewer: async () => false,
  });
  return <ViewerInput {...props} ref={inputRef} />;
}

function mount(
  queries: DisplayListQueries,
  documentKey: ViewerInputProps['document'] = { isDisplayOnly: () => false },
  pagedRef?: RefObject<PagedEditorRef | null>,
  onSelectionChange: ViewerInputProps['onSelectionChange'] = () => {}
) {
  const pending: Array<{
    request: ResidentDocumentRead;
    resolve(reply: { version: string; value: unknown }): void;
  }> = [];
  const read = ((request: ResidentDocumentRead) => new Promise<{ version: string; value: unknown }>((resolve) => {
    pending.push({ request, resolve });
  })) as unknown as ResidentEngineWorkerClient['documentRead'];
  const ref = createRef<YrsInputRef>();
  let props: ViewerInputProps = { read, story: 'body', queries, document: documentKey, onSelectionChange };
  const element = () => pagedRef
    ? <NavigatingViewer props={props} inputRef={ref} pagedRef={pagedRef} />
    : <ViewerInput {...props} ref={ref} />;
  const view = render(element());
  const show = (queries: DisplayListQueries, document = props.document) => {
    props = { ...props, queries, document };
    view.rerender(element());
  };
  const answer = async (kind: ResidentDocumentRead['kind'], expectVersion: string, version: string, value: unknown) => {
    const at = pending.findIndex((entry) => entry.request.kind === kind &&
      'expectVersion' in entry.request && entry.request.expectVersion === expectVersion);
    expect(at).toBeGreaterThanOrEqual(0);
    const [entry] = pending.splice(at, 1);
    await act(async () => { entry!.resolve({ version, value }); });
  };
  const has = (kind: ResidentDocumentRead['kind'], version: string) => pending.some((entry) =>
    entry.request.kind === kind && 'expectVersion' in entry.request && entry.request.expectVersion === version);
  return { ref, view, show, answer, has, textarea: view.getByTestId('yrs-input') as HTMLTextAreaElement };
}

test('a missing viewer navigation leaves the selection gesture and pending copy untouched', async () => {
  const pagedRef = createRef<PagedEditorRef>();
  const { ref, answer } = mount(frame('A'), undefined, pagedRef);
  act(() => ref.current!.setSelectionFromDisplay(1, 6));
  const gesture = ref.current!.currentGesture!();
  const copy = ref.current!.readSelectedText!()!;
  expect(await pagedRef.current!.navigateViewer({ kind: 'paragraphTarget', paraId: 'missing' })).toBe(false);
  expect(ref.current!.currentGesture!()).toBe(gesture);
  expect(ref.current!.displaySelection()).toEqual({ anchor: 1, head: 6 });
  await answer('selectionText', 'A', 'A', text('Alpha'));
  expect(await copy).toBe('Alpha');
});

test('viewer publication uses ordered display positions, worker paragraph identity and the presented layout', async () => {
  const queries = frame('A');
  const changes: ViewerSelectionChange[] = [];
  const { ref, show, answer, has } = mount(queries, undefined, undefined, (selection) => changes.push(selection));
  act(() => ref.current!.setSelectionFromDisplay(12, 1));
  expect(changes.at(-1)).toEqual({
    displayRange: { story: 'body', from: 1, to: 12, layoutId: layoutIdOf(queries) },
    isMultiParagraph: false,
  });
  const count = changes.length;
  await answer('selectionText', 'A', 'A', {
    text: 'AlphaBeta', range: { story: 'body', view: 'accepted',
      start: { paraId: '00000001', offset: 0 }, end: { paraId: '00000002', offset: 4 } },
  });
  expect(changes).toHaveLength(count + 1);
  expect(changes.at(-1)?.isMultiParagraph).toBe(true);
  const next = frame('A');
  show(next);
  expect(changes.at(-1)).toEqual({
    displayRange: { story: 'body', from: 1, to: 12, layoutId: layoutIdOf(next) },
    isMultiParagraph: true,
  });
  expect(has('selectionText', 'A')).toBe(false);
  show(frame('B'));
  expect(changes.at(-1)).toEqual({ displayRange: null, isMultiParagraph: false });
});

test.each([
  ['Alpha\nBeta', null, true],
  ['Alpha\tBeta', null, true],
  ['Alpha', null, false],
  ['Alpha\nBeta', {
    story: 'body', view: 'accepted',
    start: { paraId: '00000001', offset: 0 }, end: { paraId: '00000001', offset: 10 },
  }, false],
] satisfies Array<[string, DocxDisplaySelectionText['range'], boolean]>)(
  'viewer capture classifies paragraph coverage for %p', async (value, range, isMultiParagraph) => {
    const changes: ViewerSelectionChange[] = [];
    const { ref, answer } = mount(frame('A'), undefined, undefined, (selection) => changes.push(selection));
    act(() => ref.current!.setSelectionFromDisplay(1, 12));
    expect(changes.at(-1)?.isMultiParagraph).toBe(false);
    await answer('selectionText', 'A', 'A', { text: value, range });
    expect(changes.at(-1)?.isMultiParagraph).toBe(isMultiParagraph);
  }
);

test('a version change clears a viewer selection and rejects its pending copy', async () => {
  const { ref, show, answer, has } = mount(frame('A'));
  act(() => ref.current!.setSelectionFromDisplay(1, 6));
  const copy = ref.current!.readSelectedText!()!.catch((error: Error) => error);
  show(frame('B'));
  expect(ref.current!.displaySelection()).toBeNull();
  expect((await copy as Error).message).toBe('Selection cleared');
  show(frame('C'));
  await answer('selectionText', 'A', 'A', text('Alpha'));
  expect(ref.current!.displaySelection()).toBeNull();
  expect(ref.current!.readSelectedText!()).toBeNull();
  expect(has('selectionText', 'B')).toBe(false);
  expect(has('selectionText', 'C')).toBe(false);
});

test('navigation leaves a viewer selection and its copy in place', async () => {
  const pagedRef = createRef<PagedEditorRef>();
  const { ref, answer } = mount(frame('A'), undefined, pagedRef);
  act(() => ref.current!.setSelectionFromDisplay(1, 6));
  await answer('selectionText', 'A', 'A', text('Alpha'));
  act(() => {
    pagedRef.current!.scrollToPage(2);
    pagedRef.current!.scrollToPosition(10);
    pagedRef.current!.revealDisplayPosition(10);
    pagedRef.current!.scrollToParaId('00000001');
  });
  expect(ref.current!.displaySelection()).toEqual({ anchor: 1, head: 6 });
  expect(await ref.current!.readSelectedText!()).toBe('Alpha');
});

test('P1-2: a selection on presented preview pages waits for the full frame after the session handover', async () => {
  const preview = frame('preview', true);
  const { ref, show, answer } = mount(preview, { isDisplayOnly: () => true });
  show(preview, { isDisplayOnly: () => false });
  act(() => ref.current!.setSelectionFromDisplay(1, 6));
  let finished = false;
  const copied = ref.current!.readSelectedText!()!.then((value) => { finished = true; return value; });
  await answer('selectionText', 'preview', 'full', null);
  expect(finished).toBe(false);
  show(frame('full', false, true));
  expect(ref.current!.displaySelection()).toEqual({ anchor: 1, head: 6 });
  await answer('selectionText', 'full', 'full', text('Alpha'));
  expect(await copied).toBe('Alpha');
});

test('P1-3: select-all on a preview re-reads the full story before its copy settles', async () => {
  const { ref, show, answer, has, textarea } = mount(frame('preview', true), { isDisplayOnly: () => true });
  fireEvent.keyDown(textarea, { key: 'a', ctrlKey: true });
  await answer('selectionUnit', 'preview', 'preview', { anchor: 0, head: 6 });
  while (has('selectionText', 'preview')) await answer('selectionText', 'preview', 'preview', text('Alpha'));
  let finished = false;
  const copied = ref.current!.readSelectedText!()!.then((value) => { finished = true; return value; });
  await act(async () => {});
  expect(finished).toBe(false);
  show(frame('full', false, true), { isDisplayOnly: () => false });
  await answer('selectionUnit', 'full', 'full', { anchor: 0, head: 30 });
  expect(ref.current!.displaySelection()).toEqual({ anchor: 0, head: 30 });
  await answer('selectionText', 'full', 'full', text('Alpha'));
  expect(finished).toBe(false);
  await answer('selectionText', 'full', 'full', text('Alpha\nBeta\nGamma'));
  expect(await copied).toBe('Alpha\nBeta\nGamma');
});

test('P1-4: a pending copy rejects on a newer selection and leaves the clipboard untouched', async () => {
  const written: string[] = [];
  class Item {
    constructor(readonly data: Record<string, Promise<Blob>>) {}
  }
  globalThis.ClipboardItem = Item as unknown as typeof ClipboardItem;
  Object.defineProperty(navigator, 'clipboard', { configurable: true, value: {
    write: (items: Item[]) => items[0]!.data['text/plain']!.then(async (blob) => { written.push(await blob.text()); }),
    writeText: async (value: string) => { written.push(value); },
  } });
  const { ref, answer, textarea } = mount(frame('preview', true), { isDisplayOnly: () => true });
  act(() => ref.current!.setSelectionFromDisplay(1, 6));
  const pending = ref.current!.readSelectedText!()!.then(
    (value) => ({ status: 'ok', value }),
    () => ({ status: 'rejected', value: null })
  );
  fireEvent.keyDown(textarea, { key: 'c', ctrlKey: true });
  act(() => ref.current!.setSelectionFromDisplay(10, 14));
  await answer('selectionText', 'preview', 'full', null);
  await answer('selectionText', 'preview', 'preview', text('Beta'));
  expect((await pending).status).toBe('rejected');
  await act(async () => { await new Promise((resolve) => setTimeout(resolve, 80)); });
  expect(written).toEqual([]);
});

test('a copy cancelled by a version change does not write a later selection', async () => {
  const written: string[] = [];
  class Item {
    constructor(readonly data: Record<string, Promise<Blob>>) {}
  }
  globalThis.ClipboardItem = Item as unknown as typeof ClipboardItem;
  Object.defineProperty(navigator, 'clipboard', { configurable: true, value: {
    write: (items: Item[]) => items[0]!.data['text/plain']!.then(async (blob) => { written.push(await blob.text()); }),
    writeText: async (value: string) => { written.push(value); },
  } });
  const { ref, show, answer, textarea } = mount(frame('A'));
  act(() => ref.current!.setSelectionFromDisplay(1, 6));
  const copy = ref.current!.readSelectedText!()!.catch((error: Error) => error);
  fireEvent.keyDown(textarea, { key: 'c', ctrlKey: true });
  show(frame('B'));
  expect(ref.current!.displaySelection()).toBeNull();
  expect((await copy as Error).message).toBe('Selection cleared');
  act(() => ref.current!.setSelectionFromDisplay(10, 14));
  await answer('selectionText', 'A', 'A', text('Alpha'));
  expect(ref.current!.displaySelection()).toEqual({ anchor: 10, head: 14 });
  await answer('selectionText', 'B', 'B', text('Beta'));
  expect(await ref.current!.readSelectedText!()).toBe('Beta');
  await act(async () => { await new Promise((resolve) => setTimeout(resolve, 80)); });
  expect(written).toEqual([]);
});

test('horizontal keyboard movement crosses all structural positions before the next visual line', () => {
  const queries = frame('A');
  const lines = [1, 9].map((position) => ({ pageIndex: 0, from: position, to: position }));
  queries.visualLinesOnPage = (pageIndex) => lines.filter((line) => line.pageIndex === pageIndex) as unknown as
    ReturnType<DisplayListQueries['visualLinesOnPage']>;
  queries.visualLineAtPosition = (position) => (lines.find((line) => line.from === position) ?? null) as
    ReturnType<DisplayListQueries['visualLineAtPosition']>;
  const { ref, textarea } = mount(queries);
  act(() => ref.current!.setSelectionFromDisplay(1));
  fireEvent.keyDown(textarea, { key: 'ArrowRight', shiftKey: true });
  expect(ref.current!.displaySelection()).toEqual({ anchor: 1, head: 9 });
});

test('a settled viewer copy uses the synchronous textarea clipboard path', async () => {
  const { ref, answer, textarea } = mount(frame('A'));
  act(() => ref.current!.setSelectionFromDisplay(1, 6));
  await answer('selectionText', 'A', 'A', text('Alpha'));
  const data = new Map<string, string>();
  fireEvent.keyDown(textarea, { key: 'c', ctrlKey: true });
  fireEvent.copy(textarea, { clipboardData: {
    setData: (kind: string, value: string) => data.set(kind, value),
  } });
  expect(data.get('text/plain')).toBe('Alpha');
});

test('replacing a full document cancels pending viewer copies and ignores its late capture', async () => {
  const { ref, show, answer } = mount(frame('A'));
  act(() => ref.current!.setSelectionFromDisplay(1, 6));
  const copy = ref.current!.readSelectedText!()!.then(() => false, () => true);
  show(frame('B'), { isDisplayOnly: () => false });
  expect(ref.current!.displaySelection()).toBeNull();
  await answer('selectionText', 'A', 'A', text('Alpha'));
  expect(await copy).toBe(true);
  expect(ref.current!.displaySelection()).toBeNull();
});

test('unmounting the viewer cancels a pending selected-text read', async () => {
  const { ref, view, answer } = mount(frame('A'));
  act(() => ref.current!.setSelectionFromDisplay(1, 6));
  const copy = ref.current!.readSelectedText!()!.then(() => false, () => true);
  view.unmount();
  await answer('selectionText', 'A', 'A', text('Alpha'));
  expect(await copy).toBe(true);
});
