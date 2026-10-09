import { GlobalRegistrator } from '@happy-dom/global-registrator';
import { afterAll, afterEach, beforeAll, expect, mock, spyOn, test } from 'bun:test';
import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';
import { cloneElement, createRef } from 'react';
import { preloadEditWasm } from '@betteroffice/docx/wasm/edit';
import {
  createYrsInputPositionMap,
  createYrsSession,
  displayPositionToYrsLoc,
  yrsLocToDisplayPosition,
  type YrsSession,
} from '@betteroffice/docx/yrs';
import { YrsInput, type YrsInputProps, type YrsInputRef } from './YrsInput';
import type { ResidentFrameApplyResult } from './hooks/useDisplayList';
import type { DisplayListQuerySnapshot } from './hooks/displayListQueryEpochGate';
import { performYrsHistoryAction } from './yrsCommands';
import { DocxCommandAdmissionError } from '../../commands/createDocxCommandStore';
import { deferWorkerOpenReplica } from './internals/workerOpenReplica';

const ownsDom = !GlobalRegistrator.isRegistered;
if (ownsDom) GlobalRegistrator.register();
const { act, cleanup, fireEvent, render } = await import('@testing-library/react');
const sessions: YrsSession[] = [];
const globalRestores = new Set<() => void>();

function registerRestore(restore: () => void): () => void {
  const run = () => {
    if (!globalRestores.delete(run)) return;
    restore();
  };
  globalRestores.add(run);
  return run;
}

async function bounded<T>(promise: Promise<T>): Promise<T> {
  let timer!: ReturnType<typeof setTimeout>;
  try {
    return await Promise.race([
      promise,
      new Promise<never>((_, reject) => {
        timer = setTimeout(() => reject(new Error('Input replay did not settle')), 2_000);
      }),
    ]);
  } finally {
    clearTimeout(timer);
  }
}

beforeAll(() =>
  preloadEditWasm(
    new Uint8Array(
      readFileSync(
        resolve(import.meta.dir, '../../../../docx/src/wasm/generated/edit/docx_edit_bg.wasm')
      )
    )
  )
);
afterEach(() => {
  for (const restore of [...globalRestores].reverse()) restore();
  cleanup();
  for (const session of sessions.splice(0)) session.destroy();
});
afterAll(async () => {
  if (ownsDom) await GlobalRegistrator.unregister();
});

async function seededSession(): Promise<YrsSession> {
  const session = await createYrsSession();
  sessions.push(session);
  const { paraId } = session.createStory('body', 'Seed');
  session.setSelection({ story: 'body', paraId, offset: 4 });
  session.addUndoBoundary();
  return session;
}

function inputFor(
  session: YrsSession,
  input: React.Ref<YrsInputRef>,
  applyResidentInput?: YrsInputProps['applyResidentInput'],
  applyResidentDelete?: YrsInputProps['applyResidentDelete'],
  props: Partial<Pick<
    YrsInputProps,
    'isSuggesting' | 'author' | 'onPendingInputChange' | 'resolveDisplayListQueries' |
    'replicaReadyRef' | 'inputEpoch' | 'onStateChange' | 'onDirectInput' |
    'holdInput' | 'inputScope' | 'seedSelection' | 'resolveDisplayTarget'
  >> = {}
) {
  const map = (story = 'body') =>
    createYrsInputPositionMap(
      story,
      session.paragraphs(story).map((p) => ({
        paraId: p.paraId,
        length: p.text.length,
      }))
    );
  return (
    <YrsInput
      ref={input}
      enabled
      readOnly={false}
      session={session}
      inputPositionMap={map}
      displayPositionToLoc={(position, story) => displayPositionToYrsLoc(map(story), position)}
      locToDisplayPosition={(loc) => yrsLocToDisplayPosition(map(loc.story), loc)}
      onStateChange={() => {}}
      onDirectInput={() => {}}
      applyResidentInput={applyResidentInput}
      applyResidentDelete={applyResidentDelete}
      {...props}
    />
  );
}

async function mount(
  applyResidentInput?: YrsInputProps['applyResidentInput'],
  applyResidentDelete?: YrsInputProps['applyResidentDelete']
) {
  const session = await seededSession();
  const input = createRef<YrsInputRef>();
  const view = render(inputFor(session, input, applyResidentInput, applyResidentDelete));
  return { session, input, view };
}

function text(session: YrsSession): string {
  return session.paragraphs('body')[0].text;
}

function admissionCode(error: unknown): string | null {
  return error instanceof DocxCommandAdmissionError ? error.code : null;
}

async function pendingHydration() {
  const original = await seededSession();
  let retired = false;
  const oldCalls: string[] = [];
  const session = new Proxy(original, { get(target, key) {
    const value = Reflect.get(target, key);
    if (typeof value !== 'function') return value;
    return (...args: unknown[]) => {
      if (retired) oldCalls.push(String(key));
      return value.apply(target, args);
    };
  } });
  const input = createRef<YrsInputRef>();
  const replicaReadyRef = { current: false };
  let release!: () => void;
  const replica = deferWorkerOpenReplica(session,
    () => new Promise<() => void>((resolve) => { release = () => resolve(() => {}); }),
    () => {}, () => { replicaReadyRef.current = true; });
  replica.start();
  const resident = mock(async (_text: string) => null);
  const changed = mock(() => {});
  const props = { replicaReadyRef, onStateChange: changed, onDirectInput: changed,
    onPendingInputChange: changed };
  const view = render(inputFor(session, input, resident, undefined, props));
  const textarea = view.getByTestId('yrs-input') as HTMLTextAreaElement;
  return { session, input, replica, release, resident, changed, props, view, textarea,
    oldCalls, retire: () => { retired = true; } };
}

test('opening input holds both text paths, select-all, deletes and splits until replay', async () => {
  const preview = await seededSession();
  const full = await seededSession();
  const input = createRef<YrsInputRef>();
  const resident = mock(async (_text: string) => null);
  const view = render(inputFor(preview, input, resident, undefined, {
    holdInput: true, inputScope: 1, seedSelection: false,
  }));
  const textarea = view.getByTestId('yrs-input') as HTMLTextAreaElement;
  const beforeInput = new InputEvent('textInput', {
    bubbles: true, cancelable: true, inputType: 'insertText', data: 'A',
  });
  fireEvent(textarea, beforeInput);
  expect(beforeInput.defaultPrevented).toBe(true);
  fireEvent.input(textarea, { target: { value: 'B' } });
  fireEvent.keyDown(textarea, { key: 'a', ctrlKey: true });
  fireEvent.keyDown(textarea, { key: 'Delete' });
  fireEvent.input(textarea, { target: { value: 'XY' } });
  fireEvent.keyDown(textarea, { key: 'Home' });
  fireEvent.keyDown(textarea, { key: 'Delete' });
  fireEvent.keyDown(textarea, { key: 'Backspace' });
  fireEvent.keyDown(textarea, { key: 'Enter' });
  fireEvent.paste(textarea, { clipboardData: { getData: () => 'P\nQ' } });
  await act(async () => {});
  expect(text(preview)).toBe('Seed');
  expect(resident).not.toHaveBeenCalled();
  expect(input.current!.hasPendingInput()).toBe(true);
  view.rerender(inputFor(full, input, resident, undefined, { holdInput: false, inputScope: 1 }));
  await act(async () => { await input.current!.flushPendingInput(); });
  expect(full.paragraphs('body').map((paragraph) => paragraph.text)).toEqual(['', 'P', 'QY']);
  expect(full.selection()?.head.offset).toBe(1);
  expect(text(preview)).toBe('Seed');
  expect(input.current!.hasPendingInput()).toBe(false);
  expect(resident.mock.calls).toEqual([['AB'], ['XY']]);
});

test('held navigation replays between text on the full session', async () => {
  const preview = await seededSession();
  const full = await createYrsSession();
  sessions.push(full);
  const { paraId } = full.createStory('body', 'abc');
  full.setSelection({ story: 'body', paraId, offset: 3 });
  const input = createRef<YrsInputRef>();
  const resident = mock(async (_text: string) => null);
  const view = render(inputFor(preview, input, resident, undefined, { holdInput: true, inputScope: 1 }));
  const textarea = view.getByTestId('yrs-input');
  fireEvent.input(textarea, { target: { value: 'A' } });
  expect(fireEvent.keyDown(textarea, { key: 'ArrowLeft' })).toBe(false);
  fireEvent.input(textarea, { target: { value: 'B' } });
  expect(text(preview)).toBe('Seed');
  view.rerender(inputFor(full, input, resident, undefined, { inputScope: 1 }));
  await act(async () => { await input.current!.flushPendingInput(); });
  expect(text(full)).toBe('abcBA');
  expect(full.selection()?.head.offset).toBe(4);
  expect(input.current!.hasPendingInput()).toBe(false);
  expect(resident.mock.calls).toEqual([['A'], ['B']]);
});

test.each([false, true])('held table Tab between text matches ready input with shift=%s', async (shift) => {
  const preview = await seededSession();
  const full = await seededSession();
  const paraId = full.paragraphs('body')[0]!.paraId;
  full.insertTable({ story: 'body', paraId, offset: 0 }, 1, 2);
  const cells = [0, 1].map((column) => `body:t0:r0c${column}`);
  const story = cells[shift ? 1 : 0];
  const loc = { story, paraId: full.paragraphs(story)[0]!.paraId, offset: 0 };
  const ready = await createYrsSession();
  sessions.push(ready);
  ready.loadState(full.encodeState());
  ready.setSelection(loc);
  const input = createRef<YrsInputRef>();
  const replicaReadyRef = { current: true };
  const props = { inputScope: 1, replicaReadyRef };
  const view = render(inputFor(preview, input, undefined, undefined, { ...props, holdInput: true }));
  const textarea = view.getByTestId('yrs-input');
  act(() => {
    expect(input.current!.queueSelection!(async () => () => full.setSelection(loc), false, () => true)).toBe(true);
  });
  fireEvent.input(textarea, { target: { value: 'A' } });
  expect(fireEvent.keyDown(textarea, { key: 'Tab', shiftKey: shift })).toBe(false);
  fireEvent.input(textarea, { target: { value: 'B' } });
  expect(text(preview)).toBe('Seed');
  expect(cells.map((cell) => full.paragraphs(cell)[0]!.text)).toEqual(['', '']);
  view.rerender(inputFor(full, input, undefined, undefined, props));
  await act(async () => { await bounded(input.current!.flushPendingInput()); });
  const replayed = cells.map((cell) => full.paragraphs(cell)[0]!.text);
  expect(replayed).toEqual(shift ? ['B', 'A'] : ['A', 'B']);
  expect(full.selection()?.head.story).toBe(cells[shift ? 0 : 1]);
  expect(full.selection()?.head.offset).toBe(1);
  expect(input.current!.hasPendingInput()).toBe(false);
  view.unmount();
  const loaded = render(inputFor(ready, input, undefined, undefined, props));
  const loadedTextarea = loaded.getByTestId('yrs-input');
  fireEvent.input(loadedTextarea, { target: { value: 'A' } });
  expect(fireEvent.keyDown(loadedTextarea, { key: 'Tab', shiftKey: shift })).toBe(false);
  fireEvent.input(loadedTextarea, { target: { value: 'B' } });
  await act(async () => { await bounded(input.current!.flushPendingInput()); });
  expect(cells.map((cell) => ready.paragraphs(cell)[0]!.text)).toEqual(replayed);
  expect(ready.paragraphs('body').map((paragraph) => paragraph.text))
    .toEqual(full.paragraphs('body').map((paragraph) => paragraph.text));
  expect(ready.selection()).toEqual(full.selection());
  expect(ready.cellSelection()).toEqual(full.cellSelection());
});

test.each([false, true])('opening body Tab keeps native focus movement with shift=%s', async (shift) => {
  const session = await seededSession();
  const input = createRef<YrsInputRef>();
  const replicaReadyRef = { current: true };
  const props = { inputScope: 1, replicaReadyRef };
  const view = render(inputFor(session, input, undefined, undefined, { ...props, holdInput: true }));
  const textarea = view.getByTestId('yrs-input');
  expect(fireEvent.keyDown(textarea, { key: 'Tab', shiftKey: shift })).toBe(true);
  fireEvent.input(textarea, { target: { value: 'A' } });
  expect(fireEvent.keyDown(textarea, { key: 'Tab', shiftKey: shift })).toBe(true);
  fireEvent.input(textarea, { target: { value: 'B' } });
  expect(text(session)).toBe('Seed');
  view.rerender(inputFor(session, input, undefined, undefined, props));
  await act(async () => { await bounded(input.current!.flushPendingInput()); });
  expect(text(session)).toBe('SeedAB');
  expect(session.selection()?.head.offset).toBe(6);
  expect(fireEvent.keyDown(textarea, { key: 'Tab', shiftKey: shift })).toBe(true);
});

test.each([false, true])('held click in a table prevents Tab and replays cell navigation with shift=%s', async (shift) => {
  const session = await seededSession();
  session.insertTable({ ...session.selection()!.head, offset: 0 }, 1, 2);
  const body = session.selection()!.head;
  const stories = ['body:t0:r0c0', 'body:t0:r0c1'];
  const story = stories[shift ? 1 : 0];
  const cell = { story, paraId: session.paragraphs(story)[0]!.paraId, offset: 0 };
  const input = createRef<YrsInputRef>();
  const replicaReadyRef = { current: true };
  const props = { inputScope: 1, replicaReadyRef };
  const view = render(inputFor(session, input, undefined, undefined, { ...props, holdInput: true }));
  const textarea = view.getByTestId('yrs-input');
  act(() => {
    input.current!.focus();
    expect(input.current!.queueSelection!(async () => () => session.setSelection(cell), false, () => true)).toBe(true);
  });
  expect(fireEvent.keyDown(textarea, { key: 'ArrowLeft' })).toBe(false);
  expect(fireEvent.keyDown(textarea, { key: 'Tab', shiftKey: shift })).toBe(false);
  expect(session.selection()?.head).toEqual(body);
  view.rerender(inputFor(session, input, undefined, undefined, props));
  await act(async () => { await bounded(input.current!.flushPendingInput()); });
  expect(session.selection()?.head.story).toBe(stories[shift ? 0 : 1]);
  expect(session.selection()?.head.offset).toBe(0);
  expect(session.cellSelection()?.head.column).toBe(shift ? 0 : 1);
  expect(stories.map((story) => session.paragraphs(story)[0]!.text)).toEqual(['', '']);
  expect(document.activeElement).toBe(textarea);
});

test.each([false, true])('held click outside a table keeps native Tab after navigation with shift=%s', async (shift) => {
  const session = await seededSession();
  const body = session.selection()!.head;
  session.insertTable({ ...body, offset: 0 }, 1, 2);
  const stories = ['body:t0:r0c0', 'body:t0:r0c1'];
  session.setSelection({ story: stories[0], paraId: session.paragraphs(stories[0])[0]!.paraId, offset: 0 });
  const input = createRef<YrsInputRef>();
  const replicaReadyRef = { current: true };
  const verticalMove = mock(() => ({ position: 1, goalX: 0 }));
  const props = { inputScope: 1, replicaReadyRef,
    resolveDisplayListQueries: async () => ({ queries: { verticalMove } as unknown as DisplayListQuerySnapshot['queries'], frameEpoch: 1 }),
    resolveDisplayTarget: () => ({ story: stories[0], displayPosition: 1 }) };
  const view = render(inputFor(session, input, undefined, undefined, { ...props, holdInput: true }));
  const textarea = view.getByTestId('yrs-input');
  act(() => {
    input.current!.focus();
    expect(input.current!.queueSelection!(async () => () => session.setSelection(body), false, () => false)).toBe(true);
  });
  expect(fireEvent.keyDown(textarea, { key: 'Tab', shiftKey: shift })).toBe(true);
  expect(fireEvent.keyDown(textarea, { key: 'ArrowUp' })).toBe(false);
  expect(fireEvent.keyDown(textarea, { key: 'Tab', shiftKey: shift })).toBe(true);
  fireEvent.input(textarea, { target: { value: 'B' } });
  view.rerender(inputFor(session, input, undefined, undefined, props));
  await act(async () => { await bounded(input.current!.flushPendingInput()); });
  expect(verticalMove).toHaveBeenCalledTimes(1);
  expect(session.selection()?.head.story).toBe(stories[0]);
  expect(session.selection()?.head.offset).toBe(1);
  expect(stories.map((story) => session.paragraphs(story)[0]!.text)).toEqual(['B', '']);
  expect(document.activeElement).toBe(textarea);
});

test.each([false, true])('held Tab replayed outside a table leaves document, selection and focus unchanged with shift=%s', async (shift) => {
  const session = await seededSession();
  const body = session.selection()!.head;
  session.insertTable({ ...body, offset: 0 }, 1, 2);
  const story = 'body:t0:r0c0';
  const cell = { story, paraId: session.paragraphs(story)[0]!.paraId, offset: 0 };
  const state = session.encodeState();
  const version = session.version();
  const input = createRef<YrsInputRef>();
  const replicaReadyRef = { current: true };
  const verticalMove = mock(() => ({ position: 1, goalX: 0 }));
  const props = { inputScope: 1, replicaReadyRef,
    resolveDisplayListQueries: async () => ({ queries: { verticalMove } as unknown as DisplayListQuerySnapshot['queries'], frameEpoch: 1 }),
    resolveDisplayTarget: () => ({ story: 'body', displayPosition: 1 }) };
  const rectangles = spyOn(HTMLElement.prototype, 'getClientRects').mockImplementation(() =>
    [new DOMRect(0, 0, 10, 10)] as unknown as DOMRectList
  );
  const restore = registerRestore(() => rectangles.mockRestore());
  try {
    const previous = render(<button data-testid="previous" />);
    const view = render(inputFor(session, input, undefined, undefined, { ...props, holdInput: true }));
    const next = render(<button data-testid="next" />);
    const textarea = view.getByTestId('yrs-input');
    act(() => {
      input.current!.focus();
      expect(input.current!.queueSelection!(async () => () => session.setSelection(cell), false, () => true)).toBe(true);
    });
    await act(async () => {
      await new Promise<void>((resolve) => requestAnimationFrame(() => resolve()));
    });
    expect(fireEvent.keyDown(textarea, { key: 'ArrowDown' })).toBe(false);
    expect(fireEvent.keyDown(textarea, { key: 'Tab', shiftKey: shift })).toBe(false);
    expect(fireEvent.keyDown(textarea, { key: 'Tab', shiftKey: shift })).toBe(false);
    expect(document.activeElement).toBe(textarea);
    const focused = shift ? next.getByTestId('next') : previous.getByTestId('previous');
    act(() => { focused.focus(); });
    view.rerender(inputFor(session, input, undefined, undefined, props));
    await act(async () => { await bounded(input.current!.flushPendingInput()); });
    expect(verticalMove).toHaveBeenCalledTimes(1);
    const caret = { ...body, offset: 0 };
    expect(session.selection()).toEqual({ anchor: caret, head: caret });
    expect(session.cellSelection()).toBeNull();
    expect(session.encodeState()).toEqual(state);
    expect(session.version()).toBe(version);
    expect(document.activeElement).toBe(focused);
  } finally {
    restore();
  }
});

test.each([false, true])('held Tab outside a table at replay is a no-op and the next typed key lands in the editor with shift=%s', async (shift) => {
  const session = await seededSession();
  const body = session.selection()!.head;
  session.insertTable({ ...body, offset: 0 }, 1, 2);
  const stories = ['body:t0:r0c0', 'body:t0:r0c1'];
  const story = stories[0];
  const cell = { story, paraId: session.paragraphs(story)[0]!.paraId, offset: 0 };
  const input = createRef<YrsInputRef>();
  const replicaReadyRef = { current: true };
  const verticalMove = mock(() => ({ position: 1, goalX: 0 }));
  const props = { inputScope: 1, replicaReadyRef,
    resolveDisplayListQueries: async () => ({ queries: { verticalMove } as unknown as DisplayListQuerySnapshot['queries'], frameEpoch: 1 }),
    resolveDisplayTarget: () => ({ story: 'body', displayPosition: 1 }) };
  const rectangles = spyOn(HTMLElement.prototype, 'getClientRects').mockImplementation(() =>
    [new DOMRect(0, 0, 10, 10)] as unknown as DOMRectList
  );
  const restore = registerRestore(() => rectangles.mockRestore());
  try {
    const view = render(inputFor(session, input, undefined, undefined, { ...props, holdInput: true }));
    const textarea = view.getByTestId('yrs-input');
    act(() => {
      input.current!.focus();
      expect(input.current!.queueSelection!(async () => () => session.setSelection(cell), false, () => true)).toBe(true);
    });
    expect(document.activeElement).toBe(textarea);
    expect(fireEvent.keyDown(textarea, { key: 'ArrowDown' })).toBe(false);
    expect(fireEvent.keyDown(textarea, { key: 'Tab', shiftKey: shift })).toBe(false);
    fireEvent.input(textarea, { target: { value: 'X' } });
    expect(document.activeElement).toBe(textarea);
    view.rerender(inputFor(session, input, undefined, undefined, props));
    await act(async () => { await bounded(input.current!.flushPendingInput()); });
    expect(verticalMove).toHaveBeenCalledTimes(1);
    expect(stories.map((story) => session.paragraphs(story)[0]!.text)).toEqual(['', '']);
    expect(session.paragraphs('body').find((paragraph) => paragraph.paraId === body.paraId)?.text).toBe('XSeed');
    const caret = { ...body, offset: 1 };
    expect(session.selection()).toEqual({ anchor: caret, head: caret });
    expect(session.cellSelection()).toBeNull();
    expect(document.activeElement).toBe(textarea);
  } finally {
    restore();
  }
});

test('held no-op first-cell Shift+Tab keeps two undo steps across an 800 ms edit gap', async () => {
  const source = await seededSession();
  const paraId = source.paragraphs('body')[0]!.paraId;
  source.insertTable({ story: 'body', paraId, offset: 0 }, 1, 2);
  const story = 'body:t0:r0c0';
  const state = source.encodeState();
  let now = 1_000;
  const clock = spyOn(performance, 'now').mockImplementation(() => now);
  const undoClock = spyOn(Date, 'now').mockImplementation(() => now);
  const restore = registerRestore(() => {
    undoClock.mockRestore();
    clock.mockRestore();
  });
  try {
    for (const held of [false, true]) {
      now = 1_000;
      const session = await createYrsSession();
      sessions.push(session);
      session.loadState(state);
      session.setSelection({ story, paraId: session.paragraphs(story)[0]!.paraId, offset: 0 });
      const input = createRef<YrsInputRef>();
      const props = { inputScope: 1, replicaReadyRef: { current: true } };
      const view = render(inputFor(session, input, undefined, undefined, { ...props, holdInput: held }));
      const textarea = view.getByTestId('yrs-input');
      fireEvent.input(textarea, { target: { value: 'A' } });
      if (!held) await act(async () => { await bounded(input.current!.flushPendingInput()); });
      now += 400;
      expect(fireEvent.keyDown(textarea, { key: 'Tab', shiftKey: true })).toBe(false);
      if (!held) await act(async () => { await bounded(input.current!.flushPendingInput()); });
      now += 400;
      fireEvent.input(textarea, { target: { value: 'B' } });
      if (held) view.rerender(inputFor(session, input, undefined, undefined, props));
      await act(async () => { await bounded(input.current!.flushPendingInput()); });
      expect(session.paragraphs(story)[0]!.text).toBe('AB');
      expect(session.selection()?.head.story).toBe(story);
      expect(session.undo()).toBe(true);
      expect(session.paragraphs(story)[0]!.text).toBe('A');
      expect(session.undo()).toBe(true);
      expect(session.paragraphs(story)[0]!.text).toBe('');
      expect(session.canUndo()).toBe(false);
      view.unmount();
    }
  } finally {
    restore();
  }
});

test('held selection deletion follows select-all on the full session', async () => {
  const preview = await seededSession();
  const full = await seededSession();
  const input = createRef<YrsInputRef>();
  const view = render(inputFor(preview, input, undefined, undefined, { holdInput: true, inputScope: 1 }));
  act(() => {
    input.current!.selectAll();
    input.current!.deleteSelection();
  });
  expect(text(preview)).toBe('Seed');
  view.rerender(inputFor(full, input, undefined, undefined, { inputScope: 1 }));
  await act(async () => { await input.current!.flushPendingInput(); });
  expect(text(full)).toBe('');
  expect(full.selection()?.head.offset).toBe(0);
  expect(text(preview)).toBe('Seed');
});

test.each([100, 499, 500, 501, 650])(
  'worker opening replay groups undo by a %i ms held-input gap', async (gap) => {
    const session = await seededSession();
    const input = createRef<YrsInputRef>();
    const resident = mock(async (_text: string) => null);
    const replicaReadyRef = { current: true };
    let now = 1_000;
    const clock = spyOn(performance, 'now').mockImplementation(() => now);
    const restore = registerRestore(() => clock.mockRestore());
    try {
      const view = render(inputFor(session, input, resident, undefined, {
        holdInput: true, inputScope: 1, replicaReadyRef,
      }));
      const textarea = view.getByTestId('yrs-input');
      fireEvent.input(textarea, { target: { value: 'A' } });
      now += gap;
      fireEvent.input(textarea, { target: { value: 'B' } });
      now += 10_000;
      view.rerender(inputFor(session, input, resident, undefined, { inputScope: 1, replicaReadyRef }));
      await act(async () => { await bounded(input.current!.flushPendingInput()); });
      expect(text(session)).toBe('SeedAB');
      expect(resident).not.toHaveBeenCalled();
      expect(session.undoCaptureMode()).toBe('auto');
      expect(session.undo()).toBe(true);
      expect(text(session)).toBe(gap < 500 ? 'Seed' : 'SeedA');
      if (gap >= 500) {
        expect(session.undo()).toBe(true);
        expect(text(session)).toBe('Seed');
      }
      expect(session.canUndo()).toBe(false);
    } finally {
      restore();
    }
  }
);

test.each([100, 600])(
  'live input groups with the last worker replay edit after %i ms', async (gap) => {
    const session = await seededSession();
    const input = createRef<YrsInputRef>();
    const resident = mock(async (_text: string) => null);
    const replicaReadyRef = { current: true };
    let now = 1_000;
    let undoNow = 10_000;
    const clock = spyOn(performance, 'now').mockImplementation(() => now);
    const undoClock = spyOn(Date, 'now').mockImplementation(() => undoNow);
    const restore = registerRestore(() => {
      undoClock.mockRestore();
      clock.mockRestore();
    });
    const props = { inputScope: 1, replicaReadyRef };
    try {
      const view = render(inputFor(session, input, resident, undefined, { ...props, holdInput: true }));
      const textarea = view.getByTestId('yrs-input');
      fireEvent.input(textarea, { target: { value: 'A' } });
      now += 650;
      fireEvent.input(textarea, { target: { value: 'B' } });
      now += 10_000;
      view.rerender(inputFor(session, input, resident, undefined, props));
      await act(async () => { await bounded(input.current!.flushPendingInput()); });
      expect(text(session)).toBe('SeedAB');
      expect(resident).not.toHaveBeenCalled();
      expect(session.undoCaptureMode()).toBe('auto');
      undoNow += gap;
      fireEvent.input(textarea, { target: { value: 'C' } });
      await act(async () => { await bounded(input.current!.flushPendingInput()); });
      expect(text(session)).toBe('SeedABC');
      expect(resident.mock.calls).toEqual([['C']]);
      expect(session.undo()).toBe(true);
      expect(text(session)).toBe(gap < 500 ? 'SeedA' : 'SeedAB');
      if (gap >= 500) {
        expect(session.undo()).toBe(true);
        expect(text(session)).toBe('SeedA');
      }
      expect(session.undo()).toBe(true);
      expect(text(session)).toBe('Seed');
      expect(session.canUndo()).toBe(false);
    } finally {
      restore();
    }
  }
);

test('a rejected held click preserves every worker replay edit and reports its failure', async () => {
  const session = await seededSession();
  const input = createRef<YrsInputRef>();
  const replicaReadyRef = { current: true };
  const failure = new Error('selection preparation failed');
  const reportError = spyOn(console, 'error').mockImplementation(() => {});
  const restore = registerRestore(() => reportError.mockRestore());
  const props = { inputScope: 1, replicaReadyRef };
  try {
    const view = render(inputFor(session, input, undefined, undefined, { ...props, holdInput: true }));
    const textarea = view.getByTestId('yrs-input');
    fireEvent.input(textarea, { target: { value: 'A' } });
    act(() => {
      expect(input.current!.queueSelection!(async () => { throw failure; })).toBe(true);
    });
    fireEvent.input(textarea, { target: { value: 'B' } });
    view.rerender(inputFor(session, input, undefined, undefined, props));
    await act(async () => {
      await expect(bounded(input.current!.flushPendingInput())).rejects.toBe(failure);
    });
    expect(text(session)).toBe('SeedAB');
    expect(reportError).toHaveBeenCalledTimes(1);
    expect(reportError).toHaveBeenCalledWith('[YrsInput] queued input operation failed', failure);
    expect(input.current!.hasPendingInput()).toBe(false);
  } finally {
    restore();
  }
});

test.each(['navigation', 'click'] as const)(
  'a held %s separates undo without splitting the worker replay refresh', async (kind) => {
    const session = await seededSession();
    const input = createRef<YrsInputRef>();
    const replicaReadyRef = { current: true };
    const changed = mock((_selection: unknown, _docChanged: boolean) => {});
    const clock = spyOn(performance, 'now').mockReturnValue(1_000);
    const restore = registerRestore(() => clock.mockRestore());
    const props = { inputScope: 1, replicaReadyRef, onStateChange: changed };
    try {
      const view = render(inputFor(session, input, undefined, undefined, { ...props, holdInput: true }));
      const textarea = view.getByTestId('yrs-input');
      fireEvent.input(textarea, { target: { value: 'A' } });
      if (kind === 'navigation') fireEvent.keyDown(textarea, { key: 'ArrowLeft' });
      else act(() => {
        expect(input.current!.queueSelection!(async () => () => {
          const head = session.selection()!.head;
          session.setSelection({ ...head, offset: 4 });
        })).toBe(true);
      });
      fireEvent.input(textarea, { target: { value: 'B' } });
      view.rerender(inputFor(session, input, undefined, undefined, props));
      await act(async () => { await bounded(input.current!.flushPendingInput()); });
      expect(text(session)).toBe('SeedBA');
      expect(changed.mock.calls.filter(([, docChanged]) => docChanged)).toHaveLength(1);
      expect(session.undo()).toBe(true);
      expect(text(session)).toBe('SeedA');
      expect(session.undo()).toBe(true);
      expect(text(session)).toBe('Seed');
    } finally {
      restore();
    }
  }
);

test('worker replay preserves story-switch and explicit undo boundaries in manual capture', async () => {
  const session = await seededSession();
  const otherStory = 'body:other';
  const other = session.createStory(otherStory, 'Other');
  session.setUndoCaptureMode('manual');
  const input = createRef<YrsInputRef>();
  const replicaReadyRef = { current: true };
  const props = { inputScope: 1, replicaReadyRef };
  const view = render(inputFor(session, input, undefined, undefined, { ...props, holdInput: true }));
  const textarea = view.getByTestId('yrs-input');
  fireEvent.input(textarea, { target: { value: 'A' } });
  act(() => {
    input.current!.queueSelection!(async () => () => session.addUndoBoundary());
  });
  fireEvent.input(textarea, { target: { value: 'B' } });
  act(() => {
    input.current!.queueSelection!(async () => () => {
      session.setSelection({ story: otherStory, paraId: other.paraId, offset: 5 });
    });
  });
  fireEvent.input(textarea, { target: { value: 'C' } });
  const map = (story = 'body') => createYrsInputPositionMap(story,
    session.paragraphs(story).map((paragraph) => ({ paraId: paragraph.paraId, length: paragraph.text.length })));
  view.rerender(cloneElement(inputFor(session, input, undefined, undefined, props), {
    inputPositionMap: map,
    locToDisplayPosition: (loc: Parameters<typeof yrsLocToDisplayPosition>[1]) =>
      yrsLocToDisplayPosition(map(loc.story), loc),
  }));
  await act(async () => { await bounded(input.current!.flushPendingInput()); });
  expect(text(session)).toBe('SeedAB');
  expect(session.paragraphs(otherStory)[0]!.text).toBe('OtherC');
  expect(session.undoCaptureMode()).toBe('manual');
  expect(session.undo()).toBe(true);
  expect(session.paragraphs(otherStory)[0]!.text).toBe('Other');
  expect(text(session)).toBe('SeedAB');
  expect(session.undo()).toBe(true);
  expect(text(session)).toBe('SeedA');
  expect(session.undo()).toBe(true);
  expect(text(session)).toBe('Seed');
});

test('held text, horizontal navigation and composition match ready input in order', async () => {
  const held = await seededSession();
  const ready = await seededSession();
  const input = createRef<YrsInputRef>();
  const reference = createRef<YrsInputRef>();
  const replicaReadyRef = { current: true };
  const resident = mock(async (_text: string) => null);
  const view = render(inputFor(held, input, resident, undefined, {
    holdInput: true, inputScope: 1, replicaReadyRef,
  }));
  const heldTextarea = view.getByTestId('yrs-input') as HTMLTextAreaElement;
  const enter = (textarea: HTMLTextAreaElement) => fireEvent.input(textarea, { target: { value: 'A' } });
  enter(heldTextarea);
  fireEvent.keyDown(heldTextarea, { key: 'ArrowLeft' });
  fireEvent.input(heldTextarea, { target: { value: 'B' } });
  fireEvent.compositionStart(heldTextarea);
  fireEvent.compositionUpdate(heldTextarea, { data: '日' });
  heldTextarea.value = '日本';
  fireEvent.compositionEnd(heldTextarea, { data: '日本' });
  await act(async () => { await Promise.resolve(); });
  fireEvent.input(heldTextarea, { target: { value: 'C' } });
  view.rerender(inputFor(held, input, resident, undefined, { inputScope: 1, replicaReadyRef }));
  await act(async () => { await bounded(input.current!.flushPendingInput()); });
  view.unmount();
  const liveView = render(inputFor(ready, reference));
  const textarea = liveView.getByTestId('yrs-input') as HTMLTextAreaElement;
  enter(textarea);
  await act(async () => { await bounded(reference.current!.flushPendingInput()); });
  fireEvent.keyDown(textarea, { key: 'ArrowLeft' });
  fireEvent.input(textarea, { target: { value: 'B' } });
  await act(async () => { await bounded(reference.current!.flushPendingInput()); });
  fireEvent.compositionStart(textarea);
  fireEvent.compositionUpdate(textarea, { data: '日' });
  textarea.value = '日本';
  fireEvent.compositionEnd(textarea, { data: '日本' });
  await act(async () => { await bounded(reference.current!.flushPendingInput()); });
  fireEvent.input(textarea, { target: { value: 'C' } });
  await act(async () => { await bounded(reference.current!.flushPendingInput()); });
  expect(text(held)).toBe('SeedB日本CA');
  expect(text(held)).toBe(text(ready));
  expect(held.selection()?.head.offset).toBe(ready.selection()?.head.offset);
  expect(resident).not.toHaveBeenCalled();
  expect(held.undo()).toBe(true);
  expect(text(held)).toBe('SeedB日本A');
  expect(held.undo()).toBe(true);
  expect(text(held)).toBe('SeedBA');
});

test('an asynchronous navigation interrupt preserves every held edit and later live input', async () => {
  const session = await seededSession();
  const input = createRef<YrsInputRef>();
  const replicaReadyRef = { current: true };
  const resident = mock(async (_text: string) => null);
  const remove = mock(async (_direction: 'backward' | 'forward') => null);
  let release!: () => void;
  const queryReady = new Promise<null>((resolve) => { release = () => resolve(null); });
  const resolveQueries = mock(() => queryReady);
  const props = { inputScope: 1, replicaReadyRef, resolveDisplayListQueries: resolveQueries };
  const view = render(inputFor(session, input, resident, remove, { ...props, holdInput: true }));
  try {
    const textarea = view.getByTestId('yrs-input');
    fireEvent.input(textarea, { target: { value: 'AB' } });
    fireEvent.keyDown(textarea, { key: 'ArrowDown' });
    fireEvent.input(textarea, { target: { value: 'CD' } });
    fireEvent.keyDown(textarea, { key: 'Backspace' });
    fireEvent.keyDown(textarea, { key: 'Enter' });
    fireEvent.input(textarea, { target: { value: 'EF' } });
    fireEvent.keyDown(textarea, { key: 'Home' });
    fireEvent.keyDown(textarea, { key: 'Delete' });
    fireEvent.keyDown(textarea, { key: 'End' });
    fireEvent.input(textarea, { target: { value: 'G' } });
    view.rerender(inputFor(session, input, resident, remove, props));
    await act(async () => { await Promise.resolve(); });
    expect(text(session)).toBe('SeedAB');
    expect(input.current!.hasPendingInput()).toBe(true);
    fireEvent.input(textarea, { target: { value: 'H' } });
    release();
    await act(async () => { await bounded(input.current!.flushPendingInput()); });
    expect(session.paragraphs('body').map((paragraph) => paragraph.text)).toEqual(['SeedABC', 'FGH']);
    expect(input.current!.hasPendingInput()).toBe(false);
    expect(resolveQueries).toHaveBeenCalledTimes(1);
    expect(resident.mock.calls).toEqual([['H']]);
    expect(remove).not.toHaveBeenCalled();
    expect(session.undoCaptureMode()).toBe('auto');
  } finally {
    release();
  }
});

test('a held-input flush waits for replay and the full-session operation queue', async () => {
  const preview = await seededSession();
  const full = await seededSession();
  const input = createRef<YrsInputRef>();
  const replicaReadyRef = { current: false };
  let release!: () => void;
  const replica = deferWorkerOpenReplica(full,
    () => new Promise<() => void>((resolve) => { release = () => resolve(() => {}); }),
    () => {}, () => { replicaReadyRef.current = true; });
  replica.start();
  const view = render(inputFor(preview, input, undefined, undefined, { holdInput: true, inputScope: 1 }));
  const textarea = view.getByTestId('yrs-input');
  fireEvent.input(textarea, { target: { value: 'A' } });
  expect(input.current!.hasHeldInput!()).toBe(true);
  let done = false;
  const flush = input.current!.flushPendingInput().then(() => { done = true; });
  await act(async () => { await Promise.resolve(); });
  expect(done).toBe(false);
  view.rerender(inputFor(full, input, undefined, undefined, {
    holdInput: true, inputScope: 1, replicaReadyRef,
  }));
  fireEvent.input(textarea, { target: { value: 'B' } });
  await act(async () => { await Promise.resolve(); });
  expect(done).toBe(false);
  view.rerender(inputFor(full, input, undefined, undefined, { inputScope: 1, replicaReadyRef }));
  await act(async () => { await Promise.resolve(); });
  expect(input.current!.hasHeldInput!()).toBe(false);
  expect(done).toBe(false);
  expect(text(preview)).toBe('Seed');
  expect(text(full)).toBe('Seed');
  await act(async () => { release(); await flush; });
  expect(done).toBe(true);
  expect(text(full)).toBe('SeedAB');
  expect(full.selection()?.head.offset).toBe(6);
  expect(text(preview)).toBe('Seed');
});

test.each([
  ['held keys', 'scope change'],
  ['held keys', 'unmount'],
  ['active composition', 'scope change'],
  ['active composition', 'unmount'],
  ['pending composition', 'scope change'],
  ['pending composition', 'unmount'],
] as const)('a flush of %s rejects on %s before opening completes', async (kind, lifecycle) => {
  const preview = await seededSession();
  const replacement = await seededSession();
  const input = createRef<YrsInputRef>();
  const view = render(inputFor(preview, input, undefined, undefined, { holdInput: true, inputScope: 1 }));
  const textarea = view.getByTestId('yrs-input') as HTMLTextAreaElement;
  if (kind === 'held keys') fireEvent.input(textarea, { target: { value: 'discard' } });
  else {
    fireEvent.compositionStart(textarea);
    textarea.value = '日本';
    if (kind === 'pending composition') fireEvent.compositionEnd(textarea, { data: '日本' });
  }
  const flush = input.current!.flushPendingInput().catch((error) => error);
  if (lifecycle === 'unmount') view.unmount();
  else view.rerender(inputFor(replacement, input, undefined, undefined, { holdInput: true, inputScope: 2 }));
  await act(async () => {
    expect(await flush).toEqual(new Error('The document changed while flushing input'));
  });
  expect(text(preview)).toBe('Seed');
  expect(text(replacement)).toBe('Seed');
});

test.each(['scope change', 'unmount'] as const)(
  'a held-input flush rejects on %s while the replayed queue waits for the peer', async (lifecycle) => {
    const preview = await seededSession();
    const full = await seededSession();
    const replacement = await seededSession();
    const input = createRef<YrsInputRef>();
    const replicaReadyRef = { current: false };
    let release!: () => void;
    const replica = deferWorkerOpenReplica(full,
      () => new Promise<() => void>((resolve) => { release = () => resolve(() => {}); }),
      () => {}, () => { replicaReadyRef.current = true; });
    replica.start();
    const view = render(inputFor(preview, input, undefined, undefined, { holdInput: true, inputScope: 1 }));
    fireEvent.input(view.getByTestId('yrs-input'), { target: { value: 'discard' } });
    const flush = input.current!.flushPendingInput().catch((error) => error);
    view.rerender(inputFor(full, input, undefined, undefined, { inputScope: 1, replicaReadyRef }));
    await act(async () => { await Promise.resolve(); });
    expect(input.current!.hasHeldInput!()).toBe(false);
    expect(input.current!.hasPendingInput()).toBe(true);
    if (lifecycle === 'unmount') view.unmount();
    else view.rerender(inputFor(replacement, input, undefined, undefined, { inputScope: 2 }));
    expect(await flush).toEqual(new Error('The document changed while flushing input'));
    replica.cancel();
    await act(async () => { release(); await replica.ready.catch(() => {}); });
    expect(text(preview)).toBe('Seed');
    expect(text(full)).toBe('Seed');
    expect(text(replacement)).toBe('Seed');
  }
);

test.each(['held commit', 'pending commit', 'active composition'] as const)(
  'opening composition survives the same-scope session switch with %s', async (stage) => {
    const preview = await seededSession();
    const full = await seededSession();
    const input = createRef<YrsInputRef>();
    const replicaReadyRef = { current: false };
    let release!: () => void;
    const replica = deferWorkerOpenReplica(full,
      () => new Promise<() => void>((resolve) => { release = () => resolve(() => {}); }),
      () => {}, () => { replicaReadyRef.current = true; });
    replica.start();
    const view = render(inputFor(preview, input, undefined, undefined, {
      holdInput: true, inputScope: 1, seedSelection: false,
    }));
    const textarea = view.getByTestId('yrs-input') as HTMLTextAreaElement;
    fireEvent.compositionStart(textarea);
    textarea.value = '日本';
    expect(input.current!.hasHeldInput!()).toBe(true);
    let done = false;
    const flush = input.current!.flushPendingInput().then(() => { done = true; });
    if (stage !== 'active composition') fireEvent.compositionEnd(textarea, { data: '日本' });
    if (stage === 'held commit') await act(async () => { await Promise.resolve(); });
    view.rerender(inputFor(full, input, undefined, undefined, {
      holdInput: true, inputScope: 1, seedSelection: false, replicaReadyRef,
    }));
    expect(view.getByTestId('yrs-input')).toBe(textarea);
    view.rerender(inputFor(full, input, undefined, undefined, {
      holdInput: false, inputScope: 1, replicaReadyRef,
    }));
    if (stage === 'active composition') fireEvent.compositionEnd(textarea, { data: '日本' });
    fireEvent(textarea, new InputEvent('textInput', { bubbles: true, data: '日本' }));
    fireEvent.input(textarea);
    await act(async () => { await Promise.resolve(); });
    fireEvent.input(textarea);
    expect(done).toBe(false);
    expect(text(full)).toBe('Seed');
    expect(text(preview)).toBe('Seed');
    await act(async () => { release(); await flush; });
    expect(done).toBe(true);
    expect(text(full)).toBe('Seed日本');
    expect(full.selection()?.head.offset).toBe(6);
    expect(textarea.value).toBe('');
    expect(input.current!.hasPendingInput()).toBe(false);
  }
);

test.each(['scope change', 'unmount'] as const)('opening input is discarded on %s', async (lifecycle) => {
  const session = await seededSession();
  const input = createRef<YrsInputRef>();
  const view = render(inputFor(session, input, undefined, undefined, { holdInput: true, inputScope: 1 }));
  const textarea = view.getByTestId('yrs-input') as HTMLTextAreaElement;
  fireEvent.input(textarea, { target: { value: 'discard' } });
  fireEvent.compositionStart(textarea);
  textarea.value = '日本';
  fireEvent.compositionEnd(textarea, { data: '日本' });
  if (lifecycle === 'unmount') {
    view.unmount();
    render(inputFor(session, input, undefined, undefined, { inputScope: 1 }));
  } else {
    view.rerender(inputFor(session, input, undefined, undefined, { inputScope: 2 }));
  }
  await act(async () => { await input.current!.flushPendingInput(); });
  expect(text(session)).toBe('Seed');
  expect(input.current!.hasPendingInput()).toBe(false);
});

test.each([
  ['active composition', 'key'],
  ['active composition', 'composition'],
  ['pending commit', 'key'],
  ['pending commit', 'composition'],
] as const)('discarded %s ignores delayed trailing events until a fresh %s', async (stage, boundary) => {
  const preview = await seededSession();
  const full = await seededSession();
  const input = createRef<YrsInputRef>();
  const view = render(inputFor(preview, input, undefined, undefined, { holdInput: true, inputScope: 1 }));
  const textarea = view.getByTestId('yrs-input') as HTMLTextAreaElement;
  fireEvent.compositionStart(textarea);
  textarea.value = '日本';
  if (stage === 'pending commit') fireEvent.compositionEnd(textarea, { data: '日本' });
  view.rerender(inputFor(full, input, undefined, undefined, { inputScope: 2 }));
  fireEvent.compositionUpdate(textarea, { data: '日本' });
  fireEvent.input(textarea, { target: { value: '日本' }, isComposing: true });
  if (stage === 'active composition') fireEvent.compositionEnd(textarea, { data: '日本' });
  fireEvent.input(textarea, { target: { value: '日本' } });
  for (const delay of ['microtask', 'macrotask'] as const) {
    await act(async () => {
      if (delay === 'microtask') await Promise.resolve();
      else await new Promise<void>((resolve) => setTimeout(resolve, 0));
    });
    fireEvent.keyDown(textarea, { key: 'a', isComposing: true });
    fireEvent.keyDown(textarea, { key: 'Process' });
    fireEvent.keyDown(textarea, { key: 'Unidentified', keyCode: 229 });
    textarea.value = '日本';
    const beforeInput = new InputEvent('textInput', {
      bubbles: true, cancelable: true, inputType: 'insertText', data: '日本',
    });
    fireEvent(textarea, beforeInput);
    expect(beforeInput.defaultPrevented).toBe(true);
    expect(textarea.value).toBe('');
    fireEvent.input(textarea, { data: '日本', target: { value: '日本' } });
    expect(textarea.value).toBe('');
    await act(async () => { await input.current!.flushPendingInput(); });
    expect(text(preview)).toBe('Seed');
    expect(text(full)).toBe('Seed');
  }
  if (boundary === 'key') {
    fireEvent.keyDown(textarea, { key: 'C' });
    fireEvent.input(textarea, { target: { value: 'C' } });
  } else {
    fireEvent.compositionStart(textarea);
    textarea.value = 'C';
    fireEvent.compositionEnd(textarea, { data: 'C' });
  }
  await act(async () => { await input.current!.flushPendingInput(); });
  expect(text(full)).toBe('SeedC');
});

test('pending hydration preserves mixed input and event-time paste in sealed FIFO batches', async () => {
  const { session, input, release, resident, view, textarea, props } = await pendingHydration();
  let epoch = 0;
  view.rerender(inputFor(session, input, resident, undefined, { ...props, inputEpoch: () => epoch }));
  fireEvent.input(textarea, { target: { value: 'A' } });
  fireEvent.keyDown(textarea, { key: 'Enter' });
  let clipboard = 'P\r\nQ';
  fireEvent.paste(textarea, { clipboardData: { getData: () => clipboard } });
  clipboard = 'invalidated';
  fireEvent.input(textarea, { target: { value: 'B' } });
  fireEvent.compositionStart(textarea);
  textarea.value = '日本';
  fireEvent.compositionEnd(textarea, { data: '日本' });
  await act(async () => { await Promise.resolve(); });
  fireEvent.input(textarea);
  fireEvent.keyDown(textarea, { key: 'Backspace' });
  epoch += 1;
  await act(async () => { await Promise.resolve(); });
  expect(input.current!.hasPendingInput()).toBe(true);
  expect(text(session)).toBe('Seed');
  expect(resident).not.toHaveBeenCalled();
  await act(async () => { release(); await input.current!.flushPendingInput(); });
  fireEvent.input(textarea, { target: { value: 'C' } });
  await act(async () => { await input.current!.flushPendingInput(); });
  expect(session.paragraphs('body').map((p) => p.text)).toEqual(['SeedA', 'P', 'QB日C']);
  expect(session.selection()?.head.offset).toBe(4);
  expect(resident.mock.calls).toEqual([['A'], ['B'], ['C']]);
});

test('flush includes a pending-hydration IME commit exactly once', async () => {
  const { session, input, release, textarea } = await pendingHydration();
  fireEvent.compositionStart(textarea);
  let done = false;
  const flush = input.current!.flushPendingInput().then(() => { done = true; });
  textarea.value = '日本';
  fireEvent.compositionEnd(textarea, { data: '日本' });
  fireEvent(textarea, new InputEvent('textInput', { bubbles: true, data: '日本' }));
  fireEvent.input(textarea);
  await act(async () => { await Promise.resolve(); });
  fireEvent.input(textarea);
  expect(done).toBe(false);
  expect(text(session)).toBe('Seed');
  expect(input.current!.hasPendingInput()).toBe(true);
  await act(async () => { release(); await flush; });
  expect(text(session)).toBe('Seed日本');
  expect(session.selection()?.head.offset).toBe(6);
  expect(textarea.value).toBe('');
});

test.each(['replacement', 'unmount'])('retained hydration input cancels safely on %s', async (lifecycle) => {
  const { session, input, replica, release, resident, changed, view, textarea, oldCalls, retire } = await pendingHydration();
  fireEvent.input(textarea, { target: { value: 'A' } });
  fireEvent.keyDown(textarea, { key: 'Enter' });
  fireEvent.keyDown(textarea, { key: 'Backspace' });
  const flush = input.current!.flushPendingInput().catch((error) => error);
  await act(async () => { await Promise.resolve(); });
  expect(input.current!.hasPendingInput()).toBe(true);
  expect(text(session)).toBe('Seed');
  const replacement = await seededSession();
  if (lifecycle === 'unmount') view.unmount();
  else view.rerender(inputFor(replacement, input, resident));
  replica.cancel();
  retire();
  changed.mockClear();
  await act(async () => { release(); expect(await flush).toBeInstanceOf(Error); });
  expect(oldCalls).toEqual([]);
  expect(resident).not.toHaveBeenCalled();
  expect(changed).not.toHaveBeenCalled();
  expect(text(replacement)).toBe('Seed');
  if (lifecycle === 'unmount') render(inputFor(replacement, input, resident));
  act(() => input.current!.insertText('C'));
  await act(async () => { await input.current!.flushPendingInput(); });
  expect(text(replacement)).toBe('SeedC');
});

test('live hydration failure rejects a waiting flush and command', async () => {
  const { session, input, replica } = await pendingHydration();
  const errors = spyOn(console, 'error').mockImplementation(() => {});
  try {
    act(() => input.current!.insertText('A'));
    const flush = input.current!.flushPendingInput().catch((error) => error);
    const command = input.current!.runAfterPendingInput(() => 'ran').catch((error) => error);
    await act(async () => { await Promise.resolve(); });
    expect(input.current!.hasPendingInput()).toBe(true);
    const failure = new Error('hydration failed');
    replica.fail(failure);
    expect(await flush).toBe(failure);
    expect(admissionCode(await command)).toBe('input-failed');
    expect(text(session)).toBe('Seed');
  } finally {
    errors.mockRestore();
  }
});

test('flush waits for resident input and publishes the latest selection', async () => {
  let release!: () => void;
  const blocked = new Promise<void>((resolve) => {
    release = resolve;
  });
  const { session, input } = await mount(async () => {
    await blocked;
    return null;
  });
  act(() => input.current!.insertText(' accepted'));
  let done = false;
  const flush = input.current!.flushPendingInput().then(() => {
    done = true;
  });
  await Promise.resolve();
  expect(done).toBe(false);
  expect(session.paragraphs('body')[0].text).toBe('Seed');
  await act(async () => {
    release();
    await flush;
  });
  expect(session.paragraphs('body')[0].text).toBe('Seed accepted');
  expect(session.selection()?.head.offset).toBe(13);
});

test('flush seals a queued text batch before subsequent input', async () => {
  const calls: string[] = [];
  const { session, input } = await mount(async (text) => {
    calls.push(text);
    return null;
  });
  act(() => input.current!.insertText('A'));
  const first = input.current!.flushPendingInput();
  act(() => input.current!.insertText('B'));
  await act(async () => {
    await first;
    await input.current!.flushPendingInput();
  });
  expect(calls).toEqual(['A', 'B']);
  expect(session.paragraphs('body')[0].text).toBe('SeedAB');
});

test('suggesting type-over places the caret and stored formatting on the inserted text', async () => {
  const session = await seededSession();
  const paraId = session.paragraphs('body')[0]!.paraId;
  session.setSelection(
    { story: 'body', paraId, offset: 1 },
    { story: 'body', paraId, offset: 3 }
  );
  const input = createRef<YrsInputRef>();
  render(inputFor(session, input, undefined, undefined, { isSuggesting: true, author: 'Ada' }));
  act(() => {
    input.current!.applyStoredFormatting({ type: 'set', delta: { bold: true } });
    input.current!.insertText('X');
  });
  await act(async () => {
    await input.current!.flushPendingInput();
  });
  expect(text(session)).toBe('SeeXd');
  expect(session.selection()).toEqual({
    anchor: { story: 'body', paraId, offset: 4 },
    head: { story: 'body', paraId, offset: 4 },
  });
  const segments = session.storySegments('body');
  const inserted = segments.find((segment) => segment.kind === 'text' && segment.text === 'X');
  expect(inserted?.attributes.ins).toMatchObject({ author: 'Ada' });
  expect(inserted?.attributes.bold).toBe(true);
  const deleted = segments.find((segment) => segment.kind === 'text' && segment.text === 'ee');
  expect(deleted?.attributes.del).toMatchObject({ author: 'Ada' });
  expect(deleted?.attributes.bold).not.toBe(true);
  act(() => input.current!.insertText('Y'));
  await act(async () => {
    await input.current!.flushPendingInput();
  });
  expect(text(session)).toBe('SeeXYd');
  expect(session.selection()?.head).toEqual({ story: 'body', paraId, offset: 5 });
});

test('suggesting type-over across paragraphs keeps the stored formatting at the caret', async () => {
  const session = await seededSession();
  const first = session.paragraphs('body')[0]!.paraId;
  const { secondParaId: second } = session.splitParagraph({
    story: 'body',
    paraId: first,
    offset: 2,
  });
  session.setSelection(
    { story: 'body', paraId: second, offset: 1 },
    { story: 'body', paraId: first, offset: 1 }
  );
  const input = createRef<YrsInputRef>();
  render(inputFor(session, input, undefined, undefined, { isSuggesting: true, author: 'Ada' }));
  act(() => {
    input.current!.applyStoredFormatting({ type: 'set', delta: { bold: true } });
    input.current!.insertText('X');
  });
  await act(async () => {
    await input.current!.flushPendingInput();
  });
  act(() => input.current!.insertText('Y'));
  await act(async () => {
    await input.current!.flushPendingInput();
  });
  const typed = session
    .storySegments('body')
    .filter((segment) => segment.kind === 'text' && /[XY]/.test(segment.text));
  expect(typed.map((segment) => (segment.kind === 'text' ? segment.text : '')).join('')).toBe(
    'XY'
  );
  expect(typed.every((segment) => segment.attributes.bold === true)).toBe(true);
});

test('multiline suggesting type-over keeps the stored formatting on every piece', async () => {
  for (const typed of ['\nX', 'X\nY']) {
    const session = await seededSession();
    const first = session.paragraphs('body')[0]!.paraId;
    const { secondParaId: second } = session.splitParagraph({
      story: 'body',
      paraId: first,
      offset: 2,
    });
    session.setSelection(
      { story: 'body', paraId: first, offset: 1 },
      { story: 'body', paraId: second, offset: 1 }
    );
    const input = createRef<YrsInputRef>();
    render(inputFor(session, input, undefined, undefined, { isSuggesting: true, author: 'Ada' }));
    act(() => {
      input.current!.applyStoredFormatting({ type: 'set', delta: { bold: true } });
      input.current!.insertText(typed);
    });
    await act(async () => {
      await input.current!.flushPendingInput();
    });
    const inserted = session
      .storySegments('body')
      .filter((segment) => segment.kind === 'text' && /[XY]/.test(segment.text));
    expect(inserted.map((segment) => (segment.kind === 'text' ? segment.text : '')).join('')).toBe(
      typed.replace('\n', '')
    );
    expect(inserted.every((segment) => segment.attributes.bold === true)).toBe(true);
    cleanup();
  }
});

test('type-over keeps paragraph stored formatting apart from the carried override', async () => {
  const setup = async (isSuggesting: boolean) => {
    const session = await seededSession();
    const first = session.paragraphs('body')[0]!.paraId;
    const { secondParaId: second } = session.splitParagraph({
      story: 'body',
      paraId: first,
      offset: 2,
    });
    const input = createRef<YrsInputRef>();
    render(inputFor(session, input, undefined, undefined, { isSuggesting, author: 'Ada' }));
    const store = (paraId: string, delta: Record<string, boolean>) => {
      session.setSelection({ story: 'body', paraId, offset: 1 });
      act(() => input.current!.applyStoredFormatting({ type: 'set', delta }));
    };
    store(first, { italic: true });
    store(second, { bold: true });
    const type = async (text: string) => {
      act(() => input.current!.insertText(text));
      await act(async () => {
        await input.current!.flushPendingInput();
      });
    };
    const attributesOf = (text: string) => {
      const segment = session
        .storySegments('body')
        .find((candidate) => candidate.kind === 'text' && candidate.text.includes(text));
      return segment?.attributes ?? {};
    };
    return { session, first, second, type, attributesOf };
  };

  // Plain type-over formats the text with its own paragraph's stored formatting.
  const plain = await setup(false);
  plain.session.setSelection(
    { story: 'body', paraId: plain.first, offset: 1 },
    { story: 'body', paraId: plain.second, offset: 1 }
  );
  await plain.type('X');
  expect(plain.attributesOf('X').italic).toBe(true);
  expect(plain.attributesOf('X').bold).not.toBe(true);
  cleanup();

  // A suggested one carries the head's, without displacing another paragraph's.
  const suggested = await setup(true);
  suggested.session.setSelection(
    { story: 'body', paraId: suggested.second, offset: 1 },
    { story: 'body', paraId: suggested.first, offset: 1 }
  );
  await suggested.type('X\nY');
  expect(suggested.attributesOf('X').italic).toBe(true);
  expect(suggested.attributesOf('Y').italic).toBe(true);
  suggested.session.setSelection({ story: 'body', paraId: suggested.second, offset: 0 });
  await suggested.type('Z');
  expect(suggested.attributesOf('Z').bold).toBe(true);
  cleanup();
});

test('flush includes a completed IME composition exactly once', async () => {
  const { session, input, view } = await mount();
  const textarea = view.getByTestId('yrs-input') as HTMLTextAreaElement;
  fireEvent.compositionStart(textarea);
  textarea.value = '日本';
  fireEvent.compositionEnd(textarea, { data: '日本' });
  await act(async () => {
    await input.current!.flushPendingInput();
  });
  fireEvent.input(textarea);
  await act(async () => {
    await input.current!.flushPendingInput();
  });
  expect(session.paragraphs('body')[0].text).toBe('Seed日本');
  expect(session.selection()?.head.offset).toBe(6);
});

test('flush waits for an active composition and rejects if the input is removed', async () => {
  const { input, view } = await mount();
  fireEvent.compositionStart(view.getByTestId('yrs-input'));
  let done = false;
  const flush = input.current!.flushPendingInput().finally(() => {
    done = true;
  });
  await Promise.resolve();
  expect(done).toBe(false);
  view.unmount();
  await expect(flush).rejects.toThrow('unavailable');
});

test('flush rejects failed resident input instead of claiming it was committed', async () => {
  const failure = new Error('resident failure');
  const { session, input } = await mount(async () => {
    throw failure;
  });
  act(() => input.current!.insertText('lost'));
  await expect(input.current!.flushPendingInput()).rejects.toBe(failure);
  expect(session.paragraphs('body')[0].text).toBe('Seed');
});

test.each([
  ['text input', 'unmount', false],
  ['Backspace', 'session replacement', false],
  ['text input', 'unmount', true],
  ['Backspace', 'session replacement', true],
])('resident %s never calls the old session after %s (Enter queued behind it: %p)', async (operation, lifecycle, queued) => {
  const original = await seededSession();
  let retired = false;
  const afterRetirement: string[] = [];
  const session = new Proxy(original, {
    get(target, key, receiver) {
      const value = Reflect.get(target, key, receiver);
      if (typeof value !== 'function') return value;
      return (...args: unknown[]) => {
        if (retired) {
          afterRetirement.push(String(key));
          throw new Error(`Retired session method: ${String(key)}`);
        }
        return value.apply(target, args);
      };
    },
  });
  let release!: (result: ResidentFrameApplyResult) => void;
  const blocked = new Promise<ResidentFrameApplyResult>((resolve) => {
    release = resolve;
  });
  let started!: () => void;
  const applying = new Promise<void>((resolve) => {
    started = resolve;
  });
  const resident = mock((..._args: unknown[]) => {
    started();
    return blocked;
  });
  let finished!: () => void;
  const settled = new Promise<void>((resolve) => {
    finished = resolve;
  });
  const onPendingInputChange = (pending: boolean) => {
    if (!pending) finished();
  };
  const input = createRef<YrsInputRef>();
  const errors = spyOn(console, 'error').mockImplementation(() => {});
  try {
    const view = render(
      inputFor(
        session,
        input,
        operation === 'text input' ? resident : undefined,
        operation === 'Backspace' ? resident : undefined,
        { onPendingInputChange }
      )
    );
    const textarea = view.getByTestId('yrs-input');
    await act(async () => {
      if (operation === 'text input') fireEvent.input(textarea, { target: { value: 'x' } });
      else fireEvent.keyDown(textarea, { key: 'Backspace' });
      await applying;
      if (queued) fireEvent.keyDown(textarea, { key: 'Enter' });
    });
    expect(resident.mock.calls).toEqual(operation === 'text input' ? [['x']] : [['backward', 1]]);
    if (lifecycle === 'unmount') view.unmount();
    else {
      const replacement = await seededSession();
      view.rerender(inputFor(replacement, input, undefined, undefined, { onPendingInputChange }));
    }
    retired = true;
    await act(async () => {
      release({ frameEpoch: 1, caretSynchronized: true });
      await settled;
    });
    expect(afterRetirement).toEqual([]);
    expect(errors).not.toHaveBeenCalled();
  } finally {
    errors.mockRestore();
  }
});

test('undo waits behind pending typing and later typing waits behind undo', async () => {
  let release!: () => void;
  const blocked = new Promise<void>((resolve) => {
    release = resolve;
  });
  const { session, input } = await mount(async () => {
    await blocked;
    return null;
  });
  act(() => input.current!.insertText(' A'));
  const undo = input.current!.runAfterPendingInput(
    () => performYrsHistoryAction(session, false).changed
  );
  act(() => input.current!.insertText(' B'));
  expect(text(session)).toBe('Seed');
  await act(async () => {
    release();
    expect(await undo).toBe(true);
    await input.current!.flushPendingInput();
  });
  expect(text(session)).toBe('Seed B');
});

test('a command seals the text batch so input on either side stays separate', async () => {
  const calls: string[] = [];
  const { input } = await mount(async (typed) => {
    calls.push(typed);
    return null;
  });
  act(() => input.current!.insertText('x'));
  const format = input.current!.runAfterPendingInput(() => {
    calls.push('format');
  });
  act(() => input.current!.insertText('y'));
  await act(async () => {
    await format;
    await input.current!.flushPendingInput();
  });
  expect(calls).toEqual(['x', 'format', 'y']);
});

test('a key queued behind resident typing seals the text batch', async () => {
  let release!: () => void;
  const blocked = new Promise<void>((resolve) => {
    release = resolve;
  });
  const { session, input, view } = await mount(async () => {
    await blocked;
    return null;
  });
  const textarea = view.getByTestId('yrs-input');
  act(() => input.current!.insertText('A'));
  act(() => input.current!.insertText('B'));
  fireEvent.keyDown(textarea, { key: 'Backspace' });
  act(() => input.current!.insertText('C'));
  fireEvent.keyDown(textarea, { key: 'ArrowLeft' });
  act(() => input.current!.insertText('D'));
  await act(async () => {
    release();
    await input.current!.flushPendingInput();
  });
  expect(text(session)).toBe('SeedADC');
});

test('deletes queued behind busy input reach the resident engine as one batch', async () => {
  let release!: () => void;
  const blocked = new Promise<void>((resolve) => {
    release = resolve;
  });
  const deletes: string[] = [];
  const { session, input, view } = await mount(
    async () => {
      await blocked;
      return null;
    },
    async (direction, count) => {
      deletes.push(`${direction}:${count}`);
      return null;
    }
  );
  const textarea = view.getByTestId('yrs-input');
  act(() => input.current!.insertText('XYZ'));
  for (let i = 0; i < 3; i += 1) fireEvent.keyDown(textarea, { key: 'Backspace' });
  await act(async () => {
    release();
    await input.current!.flushPendingInput();
  });
  expect(deletes[0]).toBe('backward:3');
  expect(text(session)).toBe('Seed');
});

test('a resident delete is dispatched without exporting the story paragraphs', async () => {
  const session = await seededSession();
  const input = createRef<YrsInputRef>();
  const live = () =>
    createYrsInputPositionMap(
      'body',
      session.paragraphs('body').map((p) => ({ paraId: p.paraId, length: p.text.length }))
    );
  let frozen: ReturnType<typeof live> | null = null;
  const map = () => frozen ?? live();
  let exports!: ReturnType<typeof spyOn>;
  const exportsBeforeDispatch: number[] = [];
  const element = inputFor(session, input, undefined, async (_direction, count) => {
    exportsBeforeDispatch.push(exports.mock.calls.length);
    return { frameEpoch: null, caretSynchronized: false, deletedUnits: count };
  });
  const view = render(
    cloneElement(element, {
      inputPositionMap: map,
      displayPositionToLoc: (position: number) => displayPositionToYrsLoc(map(), position),
      locToDisplayPosition: (loc: Parameters<typeof yrsLocToDisplayPosition>[1]) =>
        yrsLocToDisplayPosition(map(), loc),
    })
  );
  act(() => input.current!.insertText('XY'));
  await act(async () => {
    await input.current!.flushPendingInput();
  });
  frozen = live();
  exports = spyOn(session, 'paragraphs');
  fireEvent.keyDown(view.getByTestId('yrs-input'), { key: 'Backspace' });
  await act(async () => {
    await input.current!.flushPendingInput();
  });
  exports.mockRestore();
  expect(exportsBeforeDispatch).toEqual([0]);
});

test('repeated undo requests each run', async () => {
  const { session, input } = await mount();
  act(() => input.current!.insertText(' one'));
  await act(async () => {
    await input.current!.flushPendingInput();
  });
  session.addUndoBoundary();
  act(() => input.current!.insertText(' two'));
  const undo = () =>
    input.current!.runAfterPendingInput(() => performYrsHistoryAction(session, false).changed);
  let results: boolean[] = [];
  await act(async () => {
    results = await Promise.all([undo(), undo()]);
  });
  expect(results).toEqual([true, true]);
  expect(text(session)).toBe('Seed');
});

test('a command issued during composition runs after the committed text', async () => {
  const { session, input, view } = await mount();
  const textarea = view.getByTestId('yrs-input') as HTMLTextAreaElement;
  fireEvent.compositionStart(textarea);
  let seen = null as string | null;
  const command = input.current!.runAfterPendingInput(() => {
    seen = text(session);
  });
  await Promise.resolve();
  expect(seen).toBeNull();
  textarea.value = '日本';
  await act(async () => {
    fireEvent.compositionEnd(textarea, { data: '日本' });
    await command;
  });
  expect(seen).toBe('Seed日本');
});

test('a command after lost input is refused as input-failed', async () => {
  const originalError = console.error;
  console.error = () => {};
  try {
    const { input } = await mount(async () => {
      throw new Error('resident failure');
    });
    act(() => input.current!.insertText('lost'));
    let error: unknown;
    await act(async () => {
      error = await input.current!.runAfterPendingInput(() => 'ran').catch((cause) => cause);
    });
    expect(admissionCode(error)).toBe('input-failed');
  } finally {
    console.error = originalError;
  }
});

test('input and commands after lost input still apply', async () => {
  const originalError = console.error;
  console.error = () => {};
  try {
    let fail = true;
    const { session, input } = await mount(async () => {
      if (!fail) return null;
      fail = false;
      throw new Error('resident failure');
    });
    act(() => input.current!.insertText(' lost'));
    await expect(input.current!.flushPendingInput()).rejects.toThrow('resident failure');
    act(() => input.current!.insertText(' kept'));
    let ran: unknown;
    await act(async () => {
      await input.current!.flushPendingInput();
      ran = await input.current!.runAfterPendingInput(() => 'ran');
    });
    expect(ran).toBe('ran');
    expect(text(session)).toBe('Seed kept');
  } finally {
    console.error = originalError;
  }
});

test('a flush or command waiting on composition fails when earlier input fails meanwhile', async () => {
  const originalError = console.error;
  let reported!: () => void;
  const failureReported = new Promise<void>((resolve) => {
    reported = resolve;
  });
  console.error = () => reported();
  try {
    let release!: () => void;
    const blocked = new Promise<void>((resolve) => {
      release = resolve;
    });
    let fail = true;
    const { session, input, view } = await mount(async () => {
      if (!fail) return null;
      fail = false;
      await blocked;
      throw new Error('resident failure');
    });
    const textarea = view.getByTestId('yrs-input') as HTMLTextAreaElement;
    act(() => input.current!.insertText(' lost'));
    fireEvent.compositionStart(textarea);
    const flush = input.current!.flushPendingInput().catch((cause) => cause);
    const command = input.current!.runAfterPendingInput(() => 'ran').catch((cause) => cause);
    release();
    await failureReported;
    textarea.value = '日本';
    let outcomes: unknown[] = [];
    await act(async () => {
      fireEvent.compositionEnd(textarea, { data: '日本' });
      outcomes = [await flush, admissionCode(await command)];
    });
    expect(outcomes).toEqual([new Error('resident failure'), 'input-failed']);
    let ran: unknown;
    await act(async () => {
      await input.current!.flushPendingInput();
      ran = await input.current!.runAfterPendingInput(() => 'ran');
    });
    expect(ran).toBe('ran');
    expect(text(session)).toBe('Seed日本');
  } finally {
    console.error = originalError;
  }
});

test('a command admitted before the document was replaced is refused', async () => {
  let release!: () => void;
  const blocked = new Promise<void>((resolve) => {
    release = resolve;
  });
  const first = await seededSession();
  const second = await seededSession();
  const input = createRef<YrsInputRef>();
  const resident = async () => {
    await blocked;
    return null;
  };
  const view = render(inputFor(first, input, resident));
  act(() => input.current!.insertText(' pending'));
  const command = input.current!.runAfterPendingInput(() => 'ran').catch((cause) => cause);
  act(() => {
    view.rerender(inputFor(second, input, resident));
  });
  let error: unknown;
  await act(async () => {
    release();
    error = await command;
  });
  expect(admissionCode(error)).toBe('document-replaced');
});

test.each([false, true])(
  'ArrowDown awaiting display queries never calls a replaced session (Enter queued behind it: %p)',
  async (queued) => {
    const original = await seededSession();
    const replacement = await seededSession();
    let released = false;
    const afterRelease: string[] = [];
    const session = new Proxy(original, {
      get(target, key, receiver) {
        const value = Reflect.get(target, key, receiver);
        if (typeof value !== 'function') return value;
        return (...args: unknown[]) => {
          if (released) {
            afterRelease.push(String(key));
            throw new Error(`Released session method: ${String(key)}`);
          }
          return value.apply(target, args);
        };
      },
    });
    let releaseQueries!: (value: null) => void;
    const blocked = new Promise<null>((resolve) => {
      releaseQueries = resolve;
    });
    let started!: () => void;
    const resolving = new Promise<void>((resolve) => {
      started = resolve;
    });
    const resolveDisplayListQueries = mock((_minimumFrameEpoch?: number | null) => {
      started();
      return blocked;
    });
    let finished!: () => void;
    const settled = new Promise<void>((resolve) => {
      finished = resolve;
    });
    const onPendingInputChange = (pending: boolean) => {
      if (!pending) finished();
    };
    const resident = mock(async (_text: string) => null);
    const input = createRef<YrsInputRef>();
    const setSelection = spyOn(original, 'setSelection');
    const errors = spyOn(console, 'error').mockImplementation(() => {});
    try {
      const view = render(inputFor(session, input, resident, undefined, {
        onPendingInputChange,
        resolveDisplayListQueries,
      }));
      await act(async () => {
        fireEvent.keyDown(view.getByTestId('yrs-input'), { key: 'ArrowDown' });
        await resolving;
        if (queued) fireEvent.keyDown(view.getByTestId('yrs-input'), { key: 'Enter' });
      });
      expect(resolveDisplayListQueries.mock.calls).toEqual([[null]]);
      view.rerender(inputFor(replacement, input, resident, undefined, { onPendingInputChange }));
      const replacementSelection = replacement.selection();
      setSelection.mockClear();
      released = true;
      await act(async () => {
        releaseQueries(null);
        await settled;
      });
      expect(afterRelease).toEqual([]);
      expect(setSelection).not.toHaveBeenCalled();
      expect(errors).not.toHaveBeenCalled();
      expect(replacement.selection()).toEqual(replacementSelection);
      expect(text(replacement)).toBe('Seed');
      act(() => input.current!.insertText(' B'));
      await act(async () => {
        await input.current!.flushPendingInput();
      });
      expect(resident.mock.calls).toEqual([[' B']]);
      expect(text(replacement)).toBe('Seed B');
      expect(replacement.selection()?.head.offset).toBe(6);
      expect(afterRelease).toEqual([]);
      expect(errors).not.toHaveBeenCalled();
    } finally {
      setSelection.mockRestore();
      errors.mockRestore();
    }
  }
);

test('a composition that ends as the input unmounts commits nothing to the released session', async () => {
  const original = await seededSession();
  let released = false;
  const afterRelease: string[] = [];
  const session = new Proxy(original, {
    get(target, key, receiver) {
      const value = Reflect.get(target, key, receiver);
      if (typeof value !== 'function') return value;
      return (...args: unknown[]) => {
        if (released) {
          afterRelease.push(String(key));
          throw new Error(`Released session method: ${String(key)}`);
        }
        return value.apply(target, args);
      };
    },
  });
  const view = render(inputFor(session, createRef<YrsInputRef>()));
  const textarea = view.getByTestId('yrs-input') as HTMLTextAreaElement;
  fireEvent.compositionStart(textarea);
  textarea.value = '日本';
  fireEvent.compositionEnd(textarea, { data: '日本' });
  view.unmount();
  released = true;
  await new Promise((resolve) => setTimeout(resolve, 0));
  expect(afterRelease).toEqual([]);
  expect(text(original)).toBe('Seed');
});

test('a command waiting on composition is refused when the input unmounts', async () => {
  const { input, view } = await mount();
  fireEvent.compositionStart(view.getByTestId('yrs-input'));
  const command = input.current!.runAfterPendingInput(() => 'ran').catch((cause) => cause);
  view.unmount();
  expect(admissionCode(await command)).toBe('editor-unavailable');
});

test('a stored superscript applies to the next typed text and clears subscript', async () => {
  const { session, input } = await mount();
  act(() =>
    input.current!.applyStoredFormatting({ type: 'toggle', mark: 'superscript', active: false })
  );
  expect(input.current!.storedFormatting()?.delta.other).toEqual({
    superscript: true,
    subscript: null,
  });
  act(() => input.current!.insertText('x'));
  await act(async () => {
    await input.current!.flushPendingInput();
  });
  const { paraId } = session.paragraphs('body')[0];
  const typed = session.selectionContext({
    story: 'body',
    start: { paraId, offset: 4 },
    end: { paraId, offset: 5 },
  });
  expect(typed.superscript).toBe(true);
  expect(typed.subscript).toBe(false);
});

test('typing in a table cell is offered to the resident engine', async () => {
  const session = await seededSession();
  const paraId = session.paragraphs('body')[0]!.paraId;
  session.insertTable({ story: 'body', paraId, offset: 0 }, 1, 1);
  const cell = session.storyIds().find((story) => story.startsWith('body:'))!;
  const cellParagraph = session.paragraphs(cell)[0]!.paraId;
  session.setSelection({ story: cell, paraId: cellParagraph, offset: 0 });
  const map = (story = 'body') =>
    createYrsInputPositionMap(
      story,
      session.paragraphs(story).map((p) => ({ paraId: p.paraId, length: p.text.length }))
    );
  const offered: string[] = [];
  const input = createRef<YrsInputRef>();
  render(
    <YrsInput
      ref={input}
      enabled
      readOnly={false}
      session={session}
      inputPositionMap={map}
      displayPositionToLoc={(position, story) => displayPositionToYrsLoc(map(story), position)}
      locToDisplayPosition={(loc) => yrsLocToDisplayPosition(map(loc.story), loc)}
      onStateChange={() => {}}
      onDirectInput={() => {}}
      applyResidentInput={async (text) => {
        offered.push(text);
        return null;
      }}
    />
  );
  act(() => input.current!.insertText('x'));
  await act(async () => {
    await input.current!.flushPendingInput();
  });
  expect(offered).toEqual(['x']);
  expect(session.paragraphs(cell)[0]!.text).toBe('x');
});
