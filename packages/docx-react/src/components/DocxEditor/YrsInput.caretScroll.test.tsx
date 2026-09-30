import { GlobalRegistrator } from '@happy-dom/global-registrator';
import { afterAll, afterEach, beforeAll, expect, mock, test } from 'bun:test';
import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';
import { createRef } from 'react';
import type { DisplayListQueries } from '@betteroffice/docx/layout/render';
import { preloadEditWasm } from '@betteroffice/docx/wasm/edit';
import {
  createYrsInputPositionMap,
  createYrsSession,
  displayPositionToYrsLoc,
  yrsLocToDisplayPosition,
  type YrsSession,
} from '@betteroffice/docx/yrs';
import { YrsInput, type YrsInputProps, type YrsInputRef } from './YrsInput';
import { deferWorkerOpenReplica } from './internals/workerOpenReplica';

const ownsDom = !GlobalRegistrator.isRegistered;
if (ownsDom) GlobalRegistrator.register();
const { act, cleanup, fireEvent, render } = await import('@testing-library/react');
const ROOT = resolve(import.meta.dir, '../../../../..');
const bytes = new Uint8Array(
  readFileSync(resolve(ROOT, 'crates/docx-edit/tests/fixtures/page-fragments/pages.docx'))
);
const sessions: YrsSession[] = [];
const scrollers: HTMLDivElement[] = [];

beforeAll(() =>
  preloadEditWasm(
    new Uint8Array(
      readFileSync(resolve(ROOT, 'packages/docx/src/wasm/generated/edit/docx_edit_bg.wasm'))
    )
  )
);
afterEach(() => {
  cleanup();
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
    pageSize: () => ({ width: 800, height: 3000 }),
    verticalMove: () => null,
  } satisfies Pick<DisplayListQueries, 'isReady' | 'caretRect' | 'pageSize' | 'verticalMove'>;
  const map = () => createYrsInputPositionMap('body', session.paragraphSpans('body'));
  const displayPositionToLoc: YrsInputProps['displayPositionToLoc'] = (position) =>
    displayPositionToYrsLoc(map(), position);
  const locToDisplayPosition: YrsInputProps['locToDisplayPosition'] = (loc) =>
    yrsLocToDisplayPosition(map(), loc);
  const inputFor = () => (
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
  };
}

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

// A read-only input on a replica that loads on demand; `record` stands in for a gesture the pages
// recorded meanwhile, which the host applies before input that waited.
async function mountOnDemand() {
  let epoch = 0;
  let pending: [number, number] | null = null;
  const replayed: Array<[number, number]> = [];
  const requestReplica = mock(() => {});
  const mounted = await mount(true, false, false, {
    requestReplica,
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
  const replica = deferWorkerOpenReplica(
    session,
    () =>
      new Promise<() => void>((resolve, reject) => {
        settle = (loaded) =>
          loaded
            ? resolve(() => session.openDocx(bytes, true))
            : reject(new Error('The handoff failed'));
      }),
    () => {
      throw new Error('The fallback failed');
    },
    () => {
      replicaReadyRef.current = true;
    },
    { active: () => true, request: () => replica.start() }
  );
  const finish = async (loaded: boolean) => {
    await act(async () => {
      settle(loaded);
      await input.current!.flushPendingInput();
    });
    act(() => view.rerender(inputFor()));
  };
  return {
    ...mounted,
    requestReplica,
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
  const t = await mountOnDemand();
  t.record(3);
  act(() => {
    fireEvent.keyDown(t.textarea, { key: 'ArrowRight' });
    fireEvent.keyDown(t.textarea, { key: 'ArrowRight', shiftKey: true });
  });
  expect(t.requestReplica).toHaveBeenCalled();
  expect(t.session.selection()).toBeNull();
  expect(t.input.current!.hasPendingInput()).toBe(true);

  await t.load();
  expect(t.replayed).toEqual([[3, 3]]);
  expect(t.session.selection()).toEqual({ anchor: t.loc(3), head: t.loc(4) });
});

test('a line move pressed while the replica loads moves from the recorded caret', async () => {
  const t = await mountOnDemand();
  t.record(3);
  act(() => {
    fireEvent.keyDown(t.textarea, { key: 'ArrowDown' });
  });

  await t.load();
  expect(t.replayed).toEqual([[3, 3]]);
  expect(t.session.selection()?.head.paraId).toBe(t.session.paragraphs('body')[1]!.paraId);
});

test('select all pressed after a click while the replica loads selects the whole story', async () => {
  const t = await mountOnDemand();
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
    const t = await mountOnDemand();
    t.record(2, 6);
    let notPrevented = true;
    act(() => {
      notPrevented = fireEvent.keyDown(t.textarea, { key: 'c', ctrlKey: true });
    });
    expect(notPrevented).toBe(false);
    expect(written).toHaveLength(1);

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
  const t = await mountOnDemand();
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
    const t = await mountOnDemand();
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
  const t = await mountOnDemand();
  t.record(3);
  act(() => {
    fireEvent.keyDown(t.textarea, { key: 'a', ctrlKey: true });
  });

  await t.fail();
  expect(t.replayed).toEqual([]);
  expect(t.session.selection()).toBeNull();
  expect(t.input.current!.hasPendingInput()).toBe(false);
});
