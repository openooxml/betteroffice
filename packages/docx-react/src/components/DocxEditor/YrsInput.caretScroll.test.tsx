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

async function mount(readOnly: boolean, seedSelection: boolean, replicaReady = true) {
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
  } satisfies Pick<DisplayListQueries, 'isReady' | 'caretRect' | 'pageSize'>;
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
] as const)('a %s input seeds a caret at the story start without scrolling', async (_, readOnly) => {
  const { scroller, scrollTop, expectCaret } = await mount(readOnly, true);
  expectCaret(0);
  expect(scroller.scrollTop).toBe(scrollTop);
});

test('keepSelectionInPlace suppresses scrolling for one selection', async () => {
  const { input, scroller, scrollTop, expectCaret } = await mount(false, true);
  expectCaret(0);
  expect(scroller.scrollTop).toBe(scrollTop);
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
