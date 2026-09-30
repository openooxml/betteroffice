import { GlobalRegistrator } from '@happy-dom/global-registrator';
import { afterAll, afterEach, beforeAll, expect, test } from 'bun:test';
import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';
import type { DisplayListQueries } from '@betteroffice/docx/layout/render';
import { preloadEditWasm } from '@betteroffice/docx/wasm/edit';
import { createYrsSession, type YrsSession } from '@betteroffice/docx/yrs';
import type { PagedEditorRef } from '../PagedEditor';
import { createYrsPositionProjection } from '../internals/yrsPositionProjection';
import { yrsCellStory } from '../yrsCommands';
import { useHostSearch, type DocxSearchState } from './useHostSearch';

const ownsDom = !GlobalRegistrator.isRegistered;
if (ownsDom) GlobalRegistrator.register();
const { act, cleanup, renderHook } = await import('@testing-library/react');
const sessions: YrsSession[] = [];

beforeAll(() =>
  preloadEditWasm(
    new Uint8Array(
      readFileSync(
        resolve(import.meta.dir, '../../../../../docx/src/wasm/generated/edit/docx_edit_bg.wasm')
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

/** Body "The cat and the dog", a 1x2 table "the | THE", then "the end"; a header with "the". */
async function mount(page = 1) {
  const session = await createYrsSession();
  sessions.push(session);
  const { paraId: first } = session.createStory('body', 'The cat and the dog');
  const { secondParaId: last } = session.splitParagraph({ story: 'body', paraId: first, offset: 19 });
  const { table } = session.insertTable({ story: 'body', paraId: last, offset: 0 }, 1, 2);
  for (const [column, text] of [
    [0, 'the'],
    [1, 'THE'],
  ] as const) {
    const story = yrsCellStory(session, { ...table, row: 0, column })!;
    session.insertText({ story, paraId: session.paragraphs(story)[0].paraId, offset: 0 }, text);
  }
  session.insertText({ story: 'body', paraId: last, offset: 1 }, 'the end');
  session.createStory('hdr1', 'the header');

  const reveals: number[] = [];
  const editor = {
    getYrsSession: () => session,
    hasPendingInput: () => false,
    flushPendingInput: async () => {},
    yrsLocToDisplayPosition: (loc: Parameters<PagedEditorRef['yrsLocToDisplayPosition']>[0]) =>
      createYrsPositionProjection(session, 'body')!.positionForLoc(loc),
    revealDisplayPosition: (position: number) => {
      reveals.push(position);
      return 'scrolled' as const;
    },
  } as unknown as PagedEditorRef;
  // two pages: the first paragraph, then the table and the rest
  const firstCell = yrsCellStory(session, { ...table, row: 0, column: 0 })!;
  const splitAt = editor.yrsLocToDisplayPosition({
    story: firstCell,
    paraId: session.paragraphs(firstCell)[0].paraId,
    offset: 0,
  })!;
  const queries = (version: number) =>
    ({
      version,
      anchorRect: (position: number) => ({
        pageIndex: position < splitAt ? 0 : 1,
        x: 0,
        y: 0,
        width: 1,
        height: 1,
      }),
    }) as unknown as DisplayListQueries;
  const hook = renderHook(
    ({ version }) =>
      useHostSearch({
        pagedEditorRef: { current: editor },
        displayListQueries: queries(version),
        currentPage: () => page,
      }),
    { initialProps: { version: 0 } }
  );
  const events: Array<DocxSearchState | null> = [];
  hook.result.current.api.onSearchChange((state) => events.push(state));
  return { session, first, last, hook, reveals, events };
}

test('finds every body and table match in reading order and walks them', async () => {
  const { hook, reveals, events } = await mount();
  const api = () => hook.result.current.api;
  let state = null as DocxSearchState | null;
  await act(async () => {
    state = await api().search('the');
  });
  expect(state).toEqual({ query: 'the', options: { caseSensitive: false }, total: 5, current: 0 });
  const matches = hook.result.current.highlight!.matches;
  expect(matches).toHaveLength(5);
  expect(matches.every((match, index) => index === 0 || match.displayFrom > matches[index - 1].displayFrom)).toBe(true);
  expect(reveals).toEqual([matches[0].displayFrom]);

  act(() => {
    api().searchPrevious();
  });
  expect(api().getSearchState()?.current).toBe(4);
  act(() => {
    api().searchNext();
  });
  expect(api().getSearchState()?.current).toBe(0);
  act(() => {
    api().searchGoTo(2);
  });
  expect(hook.result.current.highlight?.current).toBe(2);
  expect(reveals.at(-1)).toBe(matches[2].displayFrom);
  expect(api().searchGoTo(9)?.current).toBe(2);

  await act(async () => {
    state = await api().search('the', { caseSensitive: true });
  });
  expect(state!.total).toBe(3);

  act(() => api().clearSearch());
  expect(hook.result.current.highlight).toBeNull();
  expect(api().getSearchState()).toBeNull();
  expect(api().searchNext()).toBeNull();
  expect(events.map((event) => event?.current ?? null)).toEqual([0, 4, 0, 2, 0, null]);
});

test('starts at the first match on the page in view', async () => {
  const { hook } = await mount(2);
  await act(async () => {
    await hook.result.current.api.search('the');
  });
  expect(hook.result.current.api.getSearchState()?.current).toBe(2);
});

test('a document change re-runs the search and keeps the current match', async () => {
  const { session, first, hook, events } = await mount();
  const api = () => hook.result.current.api;
  await act(async () => {
    await api().search('the');
  });
  act(() => {
    api().searchGoTo(3);
  });
  session.insertText({ story: 'body', paraId: first, offset: 0 }, 'the ');
  hook.rerender({ version: 1 });
  expect(api().getSearchState()).toMatchObject({ total: 6, current: 4 });
  expect(events.at(-1)).toMatchObject({ total: 6, current: 4 });

  act(() => {
    api().searchGoTo(0);
  });
  hook.rerender({ version: 2 });
  expect(events).toHaveLength(4);
});

test('an empty query clears', async () => {
  const { hook } = await mount();
  await act(async () => {
    await hook.result.current.api.search('the');
  });
  let state = null as DocxSearchState | null;
  await act(async () => {
    state = await hook.result.current.api.search('');
  });
  expect(state).toEqual({ query: '', options: { caseSensitive: false }, total: 0, current: -1 });
  expect(hook.result.current.api.getSearchState()).toBeNull();
});
