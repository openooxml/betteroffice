import { GlobalRegistrator } from '@happy-dom/global-registrator';
import { afterAll, afterEach, beforeAll, expect, test } from 'bun:test';
import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';
import { preloadEditWasm } from '@betteroffice/docx/wasm/edit';
import {
  createYrsInputPositionMap,
  createYrsSession,
  displayPositionToYrsLoc,
  yrsLocToDisplayPosition,
  type YrsSession,
} from '@betteroffice/docx/yrs';
import { YrsInput } from './YrsInput';
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
  )
);
afterEach(() => {
  cleanup();
  for (const session of sessions.splice(0)) session.destroy();
});
afterAll(async () => {
  if (ownsDom) await GlobalRegistrator.unregister();
});

async function mount(readOnly: boolean) {
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
