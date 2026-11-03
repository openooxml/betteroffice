import { GlobalRegistrator } from '@happy-dom/global-registrator';
import { afterAll, afterEach, beforeAll, expect, test } from 'bun:test';
import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';
import { preloadEditWasm } from '@betteroffice/docx/wasm/edit';
import {
  applyClipboardFragment,
  createYrsInputPositionMap,
  createYrsSession,
  displayPositionToYrsLoc,
  yrsLocToDisplayPosition,
  type DocxClipboardFragment,
  type YrsSession,
} from '@betteroffice/docx/yrs';
import type { DisplayListQueries } from '@betteroffice/docx/layout/render';
import { YrsInput } from './YrsInput';
import { clipboardFragmentHtml } from './internals/clipboardHtml';
import { yrsCellStory } from './yrsCommands';

const ownsDom = !GlobalRegistrator.isRegistered;
if (ownsDom) GlobalRegistrator.register();
const { act, cleanup, fireEvent, render } = await import('@testing-library/react');
const sessions: YrsSession[] = [];

beforeAll(() =>
  preloadEditWasm(
    new Uint8Array(
      readFileSync(
        resolve(import.meta.dir, '../../../../docx/src/wasm/generated/edit/docx_edit_bg.wasm')
      )
    )
  ),
  30_000
);
afterEach(() => {
  cleanup();
  for (const session of sessions.splice(0)) session.destroy();
});
afterAll(async () => {
  if (ownsDom) await GlobalRegistrator.unregister();
});

async function mount(readOnly: boolean, suggesting = false) {
  const session = await createYrsSession();
  sessions.push(session);
  const { paraId } = session.createStory('body', 'Seed text');
  session.setSelection({ story: 'body', paraId, offset: 0 }, { story: 'body', paraId, offset: 4 });
  const map = (story = 'body') =>
    createYrsInputPositionMap(
      story,
      session.paragraphs(story).map((p) => ({ paraId: p.paraId, length: p.text.length }))
    );
  const view = render(
    <YrsInput
      enabled
      readOnly={readOnly}
      isSuggesting={suggesting}
      session={session}
      inputPositionMap={map}
      displayPositionToLoc={(position) => displayPositionToYrsLoc(map(), position)}
      locToDisplayPosition={(loc) => yrsLocToDisplayPosition(map(loc.story), loc)}
      onStateChange={() => {}}
      onDirectInput={() => {}}
    />
  );
  const textarea = view.getByTestId('yrs-input') as HTMLTextAreaElement;
  return { session, paraId, textarea };
}

function copy(textarea: HTMLTextAreaElement) {
  const data = new Map<string, string>();
  const event = fireEvent.copy(textarea, {
    clipboardData: { setData: (type: string, value: string) => data.set(type, value) },
  });
  return { data, prevented: !event };
}

for (const readOnly of [true, false]) {
  test(`Mod+C copies the selection as plain text${readOnly ? ' while read-only' : ''}`, async () => {
    const { textarea } = await mount(readOnly);
    fireEvent.keyDown(textarea, { key: 'c', ctrlKey: true });
    // the shortcut needs a native selection to run over
    expect(textarea.value).toBe('Seed');
    expect([textarea.selectionStart, textarea.selectionEnd]).toEqual([0, 4]);
    const { data, prevented } = copy(textarea);
    expect(prevented).toBe(true);
    expect(data.get('text/plain')).toBe('Seed');
    await act(() => new Promise((resolve) => setTimeout(resolve, 0)));
    expect(textarea.value).toBe('');
  });
}

test('a copy without a selection leaves the clipboard alone', async () => {
  const { session, paraId, textarea } = await mount(true);
  session.setSelection({ story: 'body', paraId, offset: 2 });
  fireEvent.keyDown(textarea, { key: 'c', metaKey: true });
  expect(textarea.value).toBe('');
  const { data, prevented } = copy(textarea);
  expect(prevented).toBe(false);
  expect(data.size).toBe(0);
});

test('read-only keys never write, and Tab stays in its table cell', async () => {
  const { session, paraId, textarea } = await mount(true);
  const { table } = session.insertTable({ story: 'body', paraId, offset: 9 }, 1, 1);
  const cell = yrsCellStory(session, { ...table, row: 0, column: 0 })!;
  session.setSelection({ story: cell, paraId: session.paragraphs(cell)[0].paraId, offset: 0 });
  const before = JSON.stringify(session.storySegments('body'));
  const selection = JSON.stringify(session.selection());
  fireEvent.keyDown(textarea, { key: 'Tab' });
  expect(JSON.stringify(session.selection())).toBe(selection);
  for (const key of ['Enter', 'Backspace', 'Delete']) fireEvent.keyDown(textarea, { key });
  await act(() => new Promise((resolve) => setTimeout(resolve, 0)));
  expect(JSON.stringify(session.storySegments('body'))).toBe(before);
});

test('read-only select all takes the document from inside a table cell', async () => {
  const { session, paraId, textarea } = await mount(true);
  const { table } = session.insertTable({ story: 'body', paraId, offset: 9 }, 1, 1);
  const cell = yrsCellStory(session, { ...table, row: 0, column: 0 })!;
  session.setSelection({ story: cell, paraId: session.paragraphs(cell)[0].paraId, offset: 0 });
  fireEvent.keyDown(textarea, { key: 'a', ctrlKey: true });
  await act(() => new Promise((resolve) => setTimeout(resolve, 0)));
  const selection = session.selection()!;
  expect([selection.anchor.story, selection.head.story]).toEqual(['body', 'body']);
  expect(selection.anchor.offset).toBe(0);
});

function tableFragment(): DocxClipboardFragment {
  return {
    paragraphs: [{
      tables: [{
        tblPr: { width: 9360, widthType: 'dxa' },
        grid: [],
        rows: [{ trPr: {}, cells: [{ tcPr: {}, content: [{ kind: 'pilcrow' }] }] }],
      }],
      text: '',
    }],
  };
}

function pasteWith(textarea: HTMLTextAreaElement, html: string, text: string) {
  fireEvent.paste(textarea, {
    clipboardData: { getData: (type: string) => type === 'text/html' ? html : text },
  });
}

test('copy of a multi-cell selection sets text html containing a table', async () => {
  const { session, paraId, textarea } = await mount(false);
  const { table } = session.insertTable({ story: 'body', paraId, offset: 9 }, 1, 2);
  const firstCell = yrsCellStory(session, { ...table, row: 0, column: 0 })!;
  session.setSelection({ story: firstCell, paraId: session.paragraphs(firstCell)[0].paraId, offset: 0 });
  session.setCellSelection({
    anchor: { ...table, row: 0, column: 0 },
    head: { ...table, row: 0, column: 1 },
  });
  fireEvent.keyDown(textarea, { key: 'c', ctrlKey: true });
  const { data, prevented } = copy(textarea);
  expect(prevented).toBe(true);
  expect(data.get('text/plain')).toContain('\t');
  expect(data.get('text/html')).toContain('<table');
});

test('paste of html with a table inserts inside one undo group', async () => {
  const { session, textarea } = await mount(false);
  const html = clipboardFragmentHtml(tableFragment());
  let boundaries = 0;
  const boundary = session.addUndoBoundary.bind(session);
  session.addUndoBoundary = () => { boundaries += 1; return boundary(); };
  pasteWith(textarea, html, 'fallback');
  await act(() => new Promise((resolve) => setTimeout(resolve, 0)));
  expect(boundaries).toBe(2);
  expect(session.storySegments('body').some((segment) => segment.kind === 'embed')).toBe(true);
  expect(session.undo()).toBe(true);
  expect(session.storySegments('body').some((segment) => segment.kind === 'embed')).toBe(false);
});

test('paste of html with a table lands after text typed in the same tick', async () => {
  const { session, textarea } = await mount(false);
  fireEvent.input(textarea, { target: { value: 'Z' } });
  pasteWith(textarea, clipboardFragmentHtml(tableFragment()), 'fallback');
  await act(() => new Promise((resolve) => setTimeout(resolve, 0)));
  const segments = session.storySegments('body');
  expect(segments.some((segment) => segment.kind === 'embed')).toBe(true);
  expect(session.paragraphs('body').map((p) => p.text).join('\n')).toContain('Z');
  const embedIndex = segments.findIndex((segment) => segment.kind === 'embed');
  const textIndex = segments.findIndex((segment) => segment.kind === 'text' && segment.text.includes('Z'));
  expect(textIndex).toBeGreaterThanOrEqual(0);
  expect(textIndex).toBeLessThan(embedIndex);
});

test('paste options text-only over a selection replaces it and one undo restores it', async () => {
  const session = await createYrsSession();
  sessions.push(session);
  const { paraId } = session.createStory('body', 'a foo b');
  session.setSelection({ story: 'body', paraId, offset: 2 }, { story: 'body', paraId, offset: 5 });
  const map = (story = 'body') =>
    createYrsInputPositionMap(
      story,
      session.paragraphs(story).map((p) => ({ paraId: p.paraId, length: p.text.length }))
    );
  const host = document.createElement('div');
  const page = document.createElement('div');
  page.className = 'canvas-page';
  page.dataset.pageIndex = '0';
  page.getBoundingClientRect = () => new DOMRect(0, 0, 800, 1000);
  host.append(page);
  document.body.append(host);
  const queries = {
    isReady: () => true,
    caretRect: (position: number) => ({ pageIndex: 0, x: position, y: 10, width: 1, height: 20 }),
    anchorRect: (position: number) => ({ pageIndex: 0, x: position, y: 10, width: 1, height: 20 }),
    pageCount: () => 1,
    pageSize: () => ({ width: 800, height: 1000 }),
    visualLines: () => [],
    visualLinesOnPage: () => [],
    visualLineExtent: () => null,
    verticalMove: () => null,
  };
  try {
    const view = render(
      <YrsInput
        enabled
        readOnly={false}
        session={session}
        canvasHostRef={{ current: host }}
        displayListQueries={queries as unknown as DisplayListQueries}
        inputPositionMap={map}
        displayPositionToLoc={(position) => displayPositionToYrsLoc(map(), position)}
        locToDisplayPosition={(loc) => yrsLocToDisplayPosition(map(loc.story), loc)}
        onStateChange={() => {}}
        onDirectInput={() => {}}
      />
    );
    const textarea = view.getByTestId('yrs-input') as HTMLTextAreaElement;
    const fragment: DocxClipboardFragment = {
      paragraphs: [{ text: 'BAR', runs: [{ text: 'BAR', format: { bold: true } }] }],
    };
    pasteWith(textarea, clipboardFragmentHtml(fragment), 'BAR');
    await act(() => new Promise((resolve) => setTimeout(resolve, 0)));
    await act(() => new Promise((resolve) => setTimeout(resolve, 0)));
    expect(session.paragraphs('body').map((p) => p.text)).toEqual(['a BAR b']);
    const control = await view.findByRole('button', { name: 'Paste Options (Ctrl)' });
    fireEvent.click(control);
    fireEvent.keyDown(view.getByRole('menu'), { key: 'Escape' });
    await act(() => new Promise((resolve) => setTimeout(resolve, 0)));
    expect(view.queryByRole('menu')).toBeNull();
    expect(view.getByRole('button', { name: 'Paste Options (Ctrl)' })).toBe(control);
    fireEvent.click(control);
    fireEvent.click(view.getByRole('menuitemradio', { name: 'Keep Text Only' }));
    await act(() => new Promise((resolve) => setTimeout(resolve, 0)));
    expect(session.paragraphs('body').map((p) => p.text)).toEqual(['a BAR b']);
    expect(session.undo()).toBe(true);
    expect(session.paragraphs('body').map((p) => p.text)).toEqual(['a foo b']);
  } finally {
    host.remove();
  }
});

test('a paste option after the paste was undone elsewhere changes nothing', async () => {
  const session = await createYrsSession();
  sessions.push(session);
  const { paraId } = session.createStory('body', 'a foo b');
  session.setSelection({ story: 'body', paraId, offset: 2 }, { story: 'body', paraId, offset: 5 });
  const map = (story = 'body') =>
    createYrsInputPositionMap(
      story,
      session.paragraphs(story).map((p) => ({ paraId: p.paraId, length: p.text.length }))
    );
  const host = document.createElement('div');
  const page = document.createElement('div');
  page.className = 'canvas-page';
  page.dataset.pageIndex = '0';
  page.getBoundingClientRect = () => new DOMRect(0, 0, 800, 1000);
  host.append(page);
  document.body.append(host);
  const queries = {
    isReady: () => true,
    caretRect: (position: number) => ({ pageIndex: 0, x: position, y: 10, width: 1, height: 20 }),
    anchorRect: (position: number) => ({ pageIndex: 0, x: position, y: 10, width: 1, height: 20 }),
    pageCount: () => 1,
    pageSize: () => ({ width: 800, height: 1000 }),
    visualLines: () => [],
    visualLinesOnPage: () => [],
    visualLineExtent: () => null,
    verticalMove: () => null,
  };
  try {
    const view = render(
      <YrsInput
        enabled
        readOnly={false}
        session={session}
        canvasHostRef={{ current: host }}
        displayListQueries={queries as unknown as DisplayListQueries}
        inputPositionMap={map}
        displayPositionToLoc={(position) => displayPositionToYrsLoc(map(), position)}
        locToDisplayPosition={(loc) => yrsLocToDisplayPosition(map(loc.story), loc)}
        onStateChange={() => {}}
        onDirectInput={() => {}}
      />
    );
    const textarea = view.getByTestId('yrs-input') as HTMLTextAreaElement;
    const fragment: DocxClipboardFragment = {
      paragraphs: [{ text: 'BAR', runs: [{ text: 'BAR', format: { bold: true } }] }],
    };
    pasteWith(textarea, clipboardFragmentHtml(fragment), 'BAR');
    await act(() => new Promise((resolve) => setTimeout(resolve, 0)));
    await act(() => new Promise((resolve) => setTimeout(resolve, 0)));
    expect(session.paragraphs('body').map((p) => p.text)).toEqual(['a BAR b']);
    const control = await view.findByRole('button', { name: 'Paste Options (Ctrl)' });
    expect(session.undo()).toBe(true);
    expect(session.paragraphs('body').map((p) => p.text)).toEqual(['a foo b']);
    fireEvent.click(control);
    fireEvent.click(view.getByRole('menuitemradio', { name: 'Keep Text Only' }));
    await act(() => new Promise((resolve) => setTimeout(resolve, 0)));
    expect(session.paragraphs('body').map((p) => p.text)).toEqual(['a foo b']);
    expect(view.queryByRole('button', { name: 'Paste Options (Ctrl)' })).toBeNull();
  } finally {
    host.remove();
  }
});

test('pasted runs keep their own formatting instead of the previous run\'s', async () => {
  const session = await createYrsSession();
  sessions.push(session);
  const { paraId } = session.createStory('body', '');
  applyClipboardFragment(session, { story: 'body', start: { paraId, offset: 0 }, end: { paraId, offset: 0 } }, {
    paragraphs: [{ text: 'A b C d', runs: [
      { text: 'A', format: { bold: true, fontSize: 20 } },
      { text: ' b ' },
      { text: 'C', format: { italic: true } },
      { text: ' d' },
    ] }],
  });
  const runs = session.storySegments('body')
    .filter((segment) => segment.kind === 'text')
    .map((segment) => segment.kind === 'text' ? [segment.text, segment.attributes.bold === true, segment.attributes.italic === true] : []);
  expect(runs).toEqual([['A', true, false], [' b ', false, false], ['C', false, true], [' d', false, false]]);
});

test('suggesting mode pastes plain text', async () => {
  const { session, textarea } = await mount(false, true);
  pasteWith(textarea, clipboardFragmentHtml(tableFragment()), 'hello');
  await act(() => new Promise((resolve) => setTimeout(resolve, 0)));
  expect(session.storySegments('body').some((segment) => segment.kind === 'embed')).toBe(false);
  expect(session.paragraphs('body').map((p) => p.text).join('\n')).toContain('hello');
});
