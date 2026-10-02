import { GlobalRegistrator } from '@happy-dom/global-registrator';
import { afterAll, afterEach, expect, test } from 'bun:test';
import { createRef } from 'react';
import type { DisplayListQueries } from '@betteroffice/docx/layout/render';
import type { DocxDisplaySelectionText, ResidentDocumentRead, ResidentEngineWorkerClient } from '@betteroffice/docx/yrs';
import { ViewerInput, type ViewerInputProps } from './ViewerInput';
import type { YrsInputRef } from './YrsInput';
import { stampWorkerFrameVersion } from './internals/layoutProvenance';

const ownsDom = !GlobalRegistrator.isRegistered;
if (ownsDom) GlobalRegistrator.register();
const { act, cleanup, fireEvent, render } = await import('@testing-library/react');
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

function frame(version: string, preview = false): DisplayListQueries {
  const queries = {
    displayList: { pages: [{ pageIndex: 0, primitives: [{ docStart: 1, docEnd: 30 }] }] },
    isReady: () => true,
    paragraphRects: () => [],
    visualLineAtPosition: (position: number) => position >= 1 && position <= 30 ? { from: 1, to: 30 } : null,
  } as unknown as DisplayListQueries;
  stampWorkerFrameVersion(queries, version, preview);
  return queries;
}

function text(value: string): DocxDisplaySelectionText {
  return { text: value, range: null, sticky: {
    anchor: { story: 'body', encoded: new Uint8Array([1]) },
    head: { story: 'body', encoded: new Uint8Array([2]) },
  } };
}

function mount(queries: DisplayListQueries, documentKey: ViewerInputProps['document'] = { isDisplayOnly: () => false }) {
  const pending: Array<{
    request: ResidentDocumentRead;
    resolve(reply: { version: string; value: unknown }): void;
  }> = [];
  const read = ((request: ResidentDocumentRead) => new Promise((resolve) => {
    pending.push({ request, resolve });
  })) as ResidentEngineWorkerClient['documentRead'];
  const ref = createRef<YrsInputRef>();
  let props: ViewerInputProps = { read, story: 'body', queries, document: documentKey, onSelectionChange: () => {} };
  const view = render(<ViewerInput {...props} ref={ref} />);
  const show = (queries: DisplayListQueries, document = props.document) => {
    props = { ...props, queries, document };
    view.rerender(<ViewerInput {...props} ref={ref} />);
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

test('P1-1: a superseded sticky reply cannot delete a selection mapping to the next frame', async () => {
  const { ref, show, answer } = mount(frame('A'));
  act(() => ref.current!.setSelectionFromDisplay(1, 6));
  await answer('selectionText', 'A', 'A', text('Alpha'));
  show(frame('B'));
  show(frame('C'));
  await answer('stickyPosition', 'B', 'C', null);
  await answer('stickyPosition', 'C', 'C', { anchor: 3, head: 8 });
  expect(ref.current!.displaySelection()).toEqual({ anchor: 3, head: 8 });
  await answer('selectionText', 'C', 'C', text('Alpha'));
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
  show(frame('full'));
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
  show(frame('full'), { isDisplayOnly: () => false });
  await answer('selectionUnit', 'full', 'full', { anchor: 0, head: 30 });
  expect(ref.current!.displaySelection()).toEqual({ anchor: 0, head: 30 });
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

test('horizontal keyboard movement crosses all structural positions before the next visual line', () => {
  const queries = frame('A');
  queries.visualLineAtPosition = (position) => position === 1 || position === 9
    ? { from: position, to: position } as ReturnType<DisplayListQueries['visualLineAtPosition']>
    : null;
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
