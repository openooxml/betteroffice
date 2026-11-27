import { afterEach, beforeAll, describe, expect, test } from 'bun:test';
import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';
import { preloadEditWasm } from '@betteroffice/docx/wasm/edit';
import {
  createYrsInputPositionMap,
  createYrsSession,
  type YrsLoc,
  type YrsSelectionContext,
  type YrsSession,
  type YrsStoryRange,
} from '@betteroffice/docx/yrs';
import { yrsCellStory } from './yrsCommands';
import {
  applyYrsToolbarFormatting,
  currentYrsToolbarSelection,
  storedYrsToolbarFormatting,
  withStoredYrsFormatting,
  type FormattingAction,
} from './yrsToolbar';

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
  for (const session of sessions.splice(0)) session.destroy();
});

async function selected(text: string, start: number, end: number) {
  const session = await createYrsSession();
  sessions.push(session);
  const { paraId } = session.createStory('body', text);
  session.setSelection(
    { story: 'body', paraId, offset: start },
    { story: 'body', paraId, offset: end }
  );
  const map = () =>
    createYrsInputPositionMap(
      'body',
      session.paragraphs('body').map((paragraph) => ({
        paraId: paragraph.paraId,
        length: paragraph.text.length,
      }))
    );
  return { session, map, selection: () => currentYrsToolbarSelection(session, map())! };
}

describe('toolbar formatting adapter', () => {
  test('sets paragraph spacing and clears competing spacing on only the supplied side', async () => {
    const { session, map, selection } = await selected('hello world', 0, 0);
    session.setParagraphAttrs(selection().range, { spaceAfter: 160, other: {
      spaceBeforeLines: 200, beforeAutospacing: true, spaceAfterLines: 100, afterAutospacing: true,
    } });
    expect(applyYrsToolbarFormatting(session, map(), { type: 'paragraphSpacing', before: 240 })).toBe(true);
    expect(selection().context.paragraphProperties).toMatchObject({
      spaceBefore: 240, beforeAutospacing: false, spaceAfter: 160, spaceAfterLines: 100, afterAutospacing: true,
    });
    expect(selection().context.paragraphProperties.spaceBeforeLines).toBe(0);
    expect(applyYrsToolbarFormatting(session, map(), { type: 'paragraphSpacing', after: 0 })).toBe(true);
    expect(selection().context.paragraphProperties).toMatchObject({ spaceBefore: 240, spaceAfter: 0, afterAutospacing: false });
    expect(selection().context.paragraphProperties.spaceAfterLines).toBe(0);
  });

  test('applies dialog line and paragraph spacing together in suggesting mode', async () => {
    const { session, map, selection } = await selected('hello world', 0, 0);
    expect(applyYrsToolbarFormatting(session, map(), { type: 'paragraphSpacing', before: 240, after: 120, lineSpacing: 360 },
      { name: 'Reviewer', date: '2026-10-31T10:00:00Z' })).toBe(true);
    expect(selection().context.paragraphProperties).toMatchObject({ spaceBefore: 240, spaceAfter: 120, lineSpacing: 360, lineSpacingRule: 'auto' });
  });

  test('toggles superscript and subscript exclusively over a range', async () => {
    const { session, map, selection } = await selected('hello world', 0, 5);
    expect(applyYrsToolbarFormatting(session, map(), 'superscript')).toBe(true);
    expect(selection().context.superscript).toBe(true);
    expect(selection().context.subscript).toBe(false);
    expect(applyYrsToolbarFormatting(session, map(), 'subscript')).toBe(true);
    expect(selection().context.superscript).toBe(false);
    expect(selection().context.subscript).toBe(true);
    session.setSelection(
      { story: 'body', paraId: selection().context.paraId, offset: 0 },
      { story: 'body', paraId: selection().context.paraId, offset: 11 }
    );
    expect(selection().context.subscript).toBe('mixed');
  });

  test('reads highlight from the authoritative selection', async () => {
    const { session, map, selection } = await selected('hello world', 0, 5);
    applyYrsToolbarFormatting(session, map(), { type: 'highlightColor', value: 'FFFF00' });
    expect(selection().context.highlight).toBe('yellow');
    applyYrsToolbarFormatting(session, map(), { type: 'highlightColor', value: 'none' });
    expect(selection().context.highlight).toBeNull();
  });

  test('a collapsed caret leaves formatting to the stored caret marks', async () => {
    const { session, map } = await selected('hello', 2, 2);
    expect(applyYrsToolbarFormatting(session, map(), 'bold')).toBe(false);
    expect(applyYrsToolbarFormatting(session, map(), { type: 'fontSize', value: 20 })).toBe(false);
    expect(session.storySegments('body').filter((segment) => segment.kind === 'text')
      .map((segment) => segment.attributes)).toEqual([{}]);
  });

  test('stores script and highlight for a collapsed caret and overlays them', async () => {
    const { selection } = await selected('hello', 5, 5);
    const context = selection().context;
    expect(storedYrsToolbarFormatting(context, 'superscript')).toEqual({
      type: 'toggle',
      mark: 'superscript',
      active: false,
    });
    const overlaid = withStoredYrsFormatting(selection(), {
      clear: false,
      delta: { highlight: 'green', other: { superscript: true, subscript: null } },
    });
    expect(overlaid.context.superscript).toBe(true);
    expect(overlaid.context.subscript).toBe(false);
    expect(overlaid.context.highlight).toBe('green');
    const cleared = withStoredYrsFormatting(selection(), { clear: true, delta: {} });
    expect(cleared.context.superscript).toBe(false);
    expect(cleared.context.highlight).toBeNull();
  });
});

function wholeStory(session: YrsSession, story: string): YrsStoryRange {
  const spans = session.paragraphSpans(story);
  const last = spans[spans.length - 1]!;
  return { story, start: { paraId: spans[0]!.paraId, offset: 0 }, end: { paraId: last.paraId, offset: last.length } };
}

function positionMap(session: YrsSession, story: string) {
  return createYrsInputPositionMap(story, session.paragraphSpans(story));
}

/** A body paragraph above a table of `texts`, the caret in the `anchor` cell and `anchor`..`head` selected. */
async function tableSelected(texts: string[][], anchor: [number, number], head: [number, number]) {
  const session = await createYrsSession();
  sessions.push(session);
  const { paraId } = session.createStory('body', 'Body text');
  const { table } = session.insertTable({ story: 'body', paraId, offset: 9 }, texts.length, texts[0]!.length);
  const stories = texts.map((row, rowIndex) => row.map((text, column) => {
    const story = yrsCellStory(session, { ...table, row: rowIndex, column })!;
    if (text) session.insertText({ story, paraId: session.paragraphs(story)[0]!.paraId, offset: 0 }, text);
    return story;
  }));
  const cell = ([row, column]: [number, number]) => ({ ...table, row, column });
  const caret: YrsLoc = { story: stories[anchor[0]]![anchor[1]]!, paraId: '', offset: 0 };
  caret.paraId = session.paragraphs(caret.story)[0]!.paraId;
  session.setSelection(caret);
  session.setCellSelection({ anchor: cell(anchor), head: cell(head) });
  const body = session.paragraphs('body').find((paragraph) => paragraph.text === 'Body text')!.paraId;
  return {
    session,
    stories,
    body: { story: 'body', start: { paraId: body, offset: 0 }, end: { paraId: body, offset: 9 } } as YrsStoryRange,
    map: () => positionMap(session, caret.story),
    whole: (story: string) => wholeStory(session, story),
    context: (story: string) => session.selectionContext(wholeStory(session, story)),
  };
}

const TEXTS = [['Alpha', 'Beta', 'Gamma'], ['Delta', 'Epsilon', 'Zeta']];
const SELECTED: Array<[number, number]> = [[0, 0], [0, 1], [1, 0], [1, 1]];
const OUTSIDE: Array<[number, number]> = [[0, 2], [1, 2]];
const COVERED: Array<[string, FormattingAction, (context: YrsSelectionContext) => unknown, unknown]> = [
  ['fontSize', { type: 'fontSize', value: 15 }, (context) => context.fontSize, 30],
  ['fontFamily', { type: 'fontFamily', value: 'Arial' }, (context) => context.fontFamily, 'Arial'],
  ['bold', 'bold', (context) => context.bold, true],
  ['italic', 'italic', (context) => context.italic, true],
  ['underline', 'underline', (context) => context.underline, true],
  ['strikethrough', 'strikethrough', (context) => context.strike, true],
  ['superscript', 'superscript', (context) => context.superscript, true],
  ['subscript', 'subscript', (context) => context.subscript, true],
  ['textColor', { type: 'textColor', value: 'FF0000' }, (context) => context.color, 'FF0000'],
  ['highlightColor', { type: 'highlightColor', value: 'FFFF00' }, (context) => context.highlight, 'yellow'],
];

describe('toolbar formatting over a multi-cell selection', () => {
  for (const [name, action, read, expected] of COVERED) {
    test(`${name} formats the whole text of every selected cell and nothing else`, async () => {
      const { session, stories, body, map, context } = await tableSelected(TEXTS, [0, 0], [1, 1]);
      const outside = () => [...OUTSIDE.map(([row, column]) => context(stories[row]![column]!)), session.selectionContext(body)];
      const before = outside();
      expect(applyYrsToolbarFormatting(session, map(), action)).toBe(true);
      expect(SELECTED.map(([row, column]) => read(context(stories[row]![column]!))))
        .toEqual(SELECTED.map(() => expected));
      expect(outside()).toEqual(before);
    });
  }

  test('clearFormatting clears every selected cell and nothing else', async () => {
    const { session, stories, body, map, whole, context } = await tableSelected(TEXTS, [1, 1], [0, 0]);
    for (const range of [...stories.flat().map(whole), body]) session.formatRange(range, { bold: true, fontSize: 14 });
    expect(applyYrsToolbarFormatting(session, map(), 'clearFormatting')).toBe(true);
    expect(SELECTED.map(([row, column]) => context(stories[row]![column]!)))
      .toMatchObject(SELECTED.map(() => ({ bold: false, fontSize: null })));
    expect([...OUTSIDE.map(([row, column]) => context(stories[row]![column]!)), session.selectionContext(body)])
      .toMatchObject([0, 1, 2].map(() => ({ bold: true, fontSize: 28 })));
  });

  test('cells of different and mixed sizes all take the chosen size', async () => {
    const { session, stories, map, whole, context } = await tableSelected([['Alpha', 'Beta'], ['Delta', 'Epsilon']], [0, 0], [1, 1]);
    session.formatRange(whole(stories[0]![0]!), { fontSize: 12 });
    session.formatRange(whole(stories[0]![1]!), { fontSize: 14 });
    session.formatRange({ ...whole(stories[1]![0]!), end: { ...whole(stories[1]![0]!).end, offset: 3 } }, { fontSize: 10 });
    expect(applyYrsToolbarFormatting(session, map(), { type: 'fontSize', value: 12 })).toBe(true);
    expect(stories.flat().map((story) => context(story).fontSize)).toEqual([24, 24, 24, 24]);
  });

  test('a toggle turns on across the cells unless all their text has it, and then off', async () => {
    const { session, stories, map, whole, context } = await tableSelected([['Alpha', 'Beta'], ['', 'Epsilon']], [0, 0], [1, 1]);
    const blank = stories[1]![0]!;
    session.splitParagraph({ story: blank, paraId: session.paragraphs(blank)[0]!.paraId, offset: 0 });
    session.formatRange(whole(stories[0]![0]!), { bold: true });
    const withText = [stories[0]![0]!, stories[0]![1]!, stories[1]![1]!];
    expect(applyYrsToolbarFormatting(session, map(), 'bold')).toBe(true);
    expect(withText.map((story) => context(story).bold)).toEqual([true, true, true]);
    expect(applyYrsToolbarFormatting(session, map(), 'bold')).toBe(true);
    expect(withText.map((story) => context(story).bold)).toEqual([false, false, false]);
  });

  test('one toolbar action undoes in one step', async () => {
    const { session, stories, map, whole, context } = await tableSelected([['Alpha', 'Beta'], ['Delta', 'Epsilon']], [0, 0], [1, 1]);
    session.formatRange(whole(stories[0]![0]!), { fontSize: 10 });
    session.formatRange(whole(stories[0]![1]!), { fontSize: 14 });
    session.addUndoBoundary();
    expect(applyYrsToolbarFormatting(session, map(), { type: 'fontSize', value: 12 })).toBe(true);
    expect(stories.flat().map((story) => context(story).fontSize)).toEqual([24, 24, 24, 24]);
    expect(session.undo()).toBe(true);
    expect(stories.flat().map((story) => context(story).fontSize)).toEqual([20, 28, null, null]);
  });

  test('a text selection inside one cell formats only that text', async () => {
    const { session, stories, context } = await tableSelected([['Alpha', 'Beta']], [0, 0], [0, 0]);
    const story = stories[0]![0]!;
    const paraId = session.paragraphs(story)[0]!.paraId;
    session.setSelection({ story, paraId, offset: 0 }, { story, paraId, offset: 3 });
    expect(applyYrsToolbarFormatting(session, positionMap(session, story), 'bold')).toBe(true);
    const part = (start: number, end: number) =>
      session.selectionContext({ story, start: { paraId, offset: start }, end: { paraId, offset: end } }).bold;
    expect([part(0, 3), part(3, 5), context(stories[0]![1]!).bold]).toEqual([true, false, false]);
  });

  test('a body text range beside a cell selection formats only that range', async () => {
    const { session, stories, body, context } = await tableSelected(TEXTS, [0, 0], [1, 1]);
    session.setSelection({ story: 'body', ...body.start }, { story: 'body', paraId: body.start.paraId, offset: 4 });
    expect(applyYrsToolbarFormatting(session, positionMap(session, 'body'), 'bold')).toBe(true);
    expect([
      session.selectionContext({ ...body, end: { ...body.end, offset: 4 } }).bold,
      session.selectionContext({ ...body, start: { ...body.start, offset: 4 } }).bold,
      ...stories.flat().map((story) => context(story).bold),
    ]).toEqual([true, false, false, false, false, false, false, false]);
  });
});
