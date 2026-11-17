import { GlobalRegistrator } from '@happy-dom/global-registrator';
import { afterAll, afterEach, beforeAll, expect, test } from 'bun:test';
import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';
import { createRef } from 'react';
import { preloadEditWasm } from '@betteroffice/docx/wasm/edit';
import {
  applyClipboardFragment,
  createYrsInputPositionMap,
  createYrsSession,
  displayPositionToYrsLoc,
  mergeFormattingFragment,
  yrsLocToDisplayPosition,
  type DocxClipboardFragment,
  type YrsSession,
} from '@betteroffice/docx/yrs';
import type { DisplayListQueries } from '@betteroffice/docx/layout/render';
import { YrsInput, type YrsInputRef } from './YrsInput';
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
      paragraphs: [{ text: 'BAR', runs: [{ text: 'BAR', format: { other: { bold: true } } }] }],
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
      paragraphs: [{ text: 'BAR', runs: [{ text: 'BAR', format: { other: { bold: true } } }] }],
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
      { text: 'A', format: { other: { bold: true, fontSize: { size: 40, sizeCs: 40 } } } },
      { text: ' b ', format: { other: {} } },
      { text: 'C', format: { other: { italic: true } } },
      { text: ' d', format: { other: {} } },
    ] }],
  });
  const runs = session.storySegments('body')
    .filter((segment) => segment.kind === 'text')
    .map((segment) => segment.kind === 'text' ? [segment.text, segment.attributes.bold === true, segment.attributes.italic === true] : []);
  expect(runs).toEqual([['A', true, false], [' b ', false, false], ['C', false, true], [' d', false, false]]);
});

test('merge formatting keeps the source marks and the destination run formatting', async () => {
  const session = await createYrsSession();
  sessions.push(session);
  const { paraId } = session.createStory('body', '');
  session.insertText({ story: 'body', paraId, offset: 0 }, 'ab');
  session.formatRange({ story: 'body', start: { paraId, offset: 0 }, end: { paraId, offset: 2 } }, {
    other: { fontFamily: { ascii: 'Arial', hAnsi: 'Arial' }, textColor: { rgb: '0000FF' }, fontSize: { size: 24, sizeCs: 24 } },
  });
  const source: DocxClipboardFragment = { paragraphs: [{ text: 'Bold tail', runs: [
    { text: 'Bold', format: { other: { bold: true, textColor: { rgb: 'FF0000' }, fontSize: { size: 40, sizeCs: 40 }, fontFamily: { ascii: 'Georgia', hAnsi: 'Georgia' } } } },
    { text: ' tail', format: { other: {} } },
  ] }] };
  const at = { paraId, offset: 1 };
  applyClipboardFragment(session, { story: 'body', start: at, end: at }, mergeFormattingFragment(source), undefined, true);
  const segments = session.storySegments('body').filter((segment) => segment.kind === 'text');
  expect(segments.map((segment) => segment.kind === 'text' ? segment.text : '')).toEqual(['a', 'Bold', ' tailb']);
  const attributesOf = (text: string) => {
    const found = segments.find((segment) => segment.kind === 'text' && segment.text.includes(text));
    return found?.kind === 'text' ? found.attributes : undefined;
  };
  const destination = { fontFamily: { ascii: 'Arial' }, textColor: { rgb: '0000FF' }, fontSize: { size: 24 } };
  expect(attributesOf('Bold')).toMatchObject({ bold: true, ...destination });
  expect(attributesOf(' tail')).toMatchObject(destination);
  expect(attributesOf(' tail')?.bold).not.toBe(true);
  for (const text of ['Bold', ' tail']) {
    expect(JSON.stringify(attributesOf(text))).not.toContain('FF0000');
    expect(JSON.stringify(attributesOf(text))).not.toContain('Georgia');
  }
});

test('merge formatting chosen from the paste options keeps the destination run formatting', async () => {
  const session = await createYrsSession();
  sessions.push(session);
  const { paraId } = session.createStory('body', 'ab');
  session.formatRange({ story: 'body', start: { paraId, offset: 0 }, end: { paraId, offset: 2 } }, {
    other: { fontFamily: { ascii: 'Arial', hAnsi: 'Arial' }, textColor: { rgb: '0000FF' }, fontSize: { size: 24, sizeCs: 24 } },
  });
  session.setSelection({ story: 'body', paraId, offset: 1 });
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
    const fragment: DocxClipboardFragment = { paragraphs: [{ text: 'Bold', runs: [{
      text: 'Bold',
      format: { other: { bold: true, textColor: { rgb: 'FF0000' }, fontSize: { size: 40, sizeCs: 40 }, fontFamily: { ascii: 'Georgia', hAnsi: 'Georgia' } } },
    }] }] };
    pasteWith(textarea, clipboardFragmentHtml(fragment), 'Bold');
    await act(() => new Promise((resolve) => setTimeout(resolve, 0)));
    await act(() => new Promise((resolve) => setTimeout(resolve, 0)));
    expect(session.paragraphs('body').map((p) => p.text)).toEqual(['aBoldb']);
    fireEvent.click(await view.findByRole('button', { name: 'Paste Options (Ctrl)' }));
    fireEvent.click(view.getByRole('menuitemradio', { name: 'Merge Formatting' }));
    await act(() => new Promise((resolve) => setTimeout(resolve, 0)));
    expect(session.paragraphs('body').map((p) => p.text)).toEqual(['aBoldb']);
    const bold = session.storySegments('body').filter((segment) => segment.kind === 'text' && segment.text === 'Bold');
    expect(bold).toHaveLength(1);
    const attributes = bold[0]!.kind === 'text' ? bold[0]!.attributes : {};
    expect(attributes).toMatchObject({
      bold: true, fontFamily: { ascii: 'Arial' }, textColor: { rgb: '0000FF' }, fontSize: { size: 24 },
    });
    expect(JSON.stringify(attributes)).not.toContain('FF0000');
    expect(JSON.stringify(attributes)).not.toContain('Georgia');
    expect(JSON.stringify((attributes as { fontSize?: unknown }).fontSize)).not.toContain('40');
  } finally {
    host.remove();
  }
});

test('suggesting mode pastes plain text', async () => {
  const { session, textarea } = await mount(false, true);
  pasteWith(textarea, clipboardFragmentHtml(tableFragment()), 'hello');
  await act(() => new Promise((resolve) => setTimeout(resolve, 0)));
  expect(session.storySegments('body').some((segment) => segment.kind === 'embed')).toBe(false);
  expect(session.paragraphs('body').map((p) => p.text).join('\n')).toContain('hello');
});

async function mountWithPasteOptions(text: string, select: [number, number], suggesting: boolean) {
  const session = await createYrsSession();
  sessions.push(session);
  const { paraId } = session.createStory('body', text);
  session.setSelection({ story: 'body', paraId, offset: select[0] }, { story: 'body', paraId, offset: select[1] });
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
  const input = createRef<YrsInputRef>();
  const view = render(
    <YrsInput
      ref={input}
      enabled
      readOnly={false}
      isSuggesting={suggesting}
      author="Ada"
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
  return { session, paraId, host, input, view, textarea };
}

for (const selected of [true, false]) {
  test(`suggesting keep text only ${selected ? 'over a selection leaves the struck original unformatted' : 'at a caret formats exactly the inserted text'}`, async () => {
    const { session, paraId, host, input, view, textarea } = await mountWithPasteOptions('a orig b', selected ? [2, 6] : [2, 2], true);
    try {
      act(() => input.current!.applyStoredFormatting({ type: 'set', delta: { bold: true } }));
      const fragment: DocxClipboardFragment = {
        paragraphs: [{ text: 'NEW', runs: [{ text: 'NEW', format: { other: { italic: true } } }] }],
      };
      pasteWith(textarea, clipboardFragmentHtml(fragment), 'NEW');
      await act(() => new Promise((resolve) => setTimeout(resolve, 0)));
      await act(() => new Promise((resolve) => setTimeout(resolve, 0)));
      fireEvent.click(await view.findByRole('button', { name: 'Paste Options (Ctrl)' }));
      fireEvent.click(view.getByRole('menuitemradio', { name: 'Keep Text Only' }));
      await act(() => new Promise((resolve) => setTimeout(resolve, 0)));
      const texts = () => session.storySegments('body').filter((segment) => segment.kind === 'text');
      const inserted = texts().filter((segment) => segment.kind === 'text' && segment.text === 'NEW');
      expect(inserted).toHaveLength(1);
      expect(inserted[0]!.kind === 'text' ? inserted[0]!.attributes.ins : undefined).toMatchObject({ author: 'Ada' });
      expect(inserted[0]!.kind === 'text' ? inserted[0]!.attributes.italic : undefined).not.toBe(true);
      expect(texts().filter((segment) => segment.kind === 'text' && segment.attributes.bold === true)
        .map((segment) => segment.kind === 'text' ? segment.text : '')).toEqual(['NEW']);
      if (selected) {
        const struck = texts().filter((segment) => segment.kind === 'text' && segment.attributes.del);
        expect(struck.map((segment) => segment.kind === 'text' ? segment.text : '')).toEqual(['orig']);
        expect(struck[0]!.kind === 'text' ? struck[0]!.attributes.bold : undefined).not.toBe(true);
        const length = session.paragraphs('body')[0]!.text.length;
        session.rejectChange({ story: 'body', start: { paraId, offset: 0 }, end: { paraId, offset: length } });
        expect(session.paragraphs('body').map((p) => p.text)).toEqual(['a orig b']);
        expect(texts().some((segment) => segment.kind === 'text' && segment.attributes.bold === true)).toBe(false);
      }
    } finally {
      host.remove();
    }
  });
}

test('a rejected clipboard payload pastes its plain text without formatting', async () => {
  const { session, textarea } = await mount(false);
  const payload = JSON.stringify({ v: 1, fragment: { paragraphs: [{ text: 'text', runs: [
    { text: 'text', format: { other: { runStyle: { styleId: 'Evil' } } } },
  ] }] } }).replaceAll('&', '&amp;').replaceAll('"', '&quot;');
  pasteWith(textarea, `<div data-betteroffice-docx="${payload}"><b>text</b></div>`, 'text');
  await act(() => new Promise((resolve) => setTimeout(resolve, 0)));
  await act(() => new Promise((resolve) => setTimeout(resolve, 0)));
  expect(session.paragraphs('body').map((p) => p.text)).toEqual(['text text']);
  expect(session.storySegments('body').some((segment) => segment.kind === 'text' && segment.attributes.bold === true)).toBe(false);
});
