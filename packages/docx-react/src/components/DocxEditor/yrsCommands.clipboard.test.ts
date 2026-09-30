import { afterEach, beforeAll, expect, test } from 'bun:test';
import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';
import { preloadEditWasm } from '@betteroffice/docx/wasm/edit';
import { createYrsSession, type YrsSession } from '@betteroffice/docx/yrs';
import {
  yrsCellStory,
  yrsSelectedText,
  yrsSelectionPlainText,
  yrsTableSelectionRange,
} from './yrsCommands';

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
  for (const session of sessions.splice(0)) session.destroy();
});

/** "one<tab>two<break>three", then a 2x2 table leading the paragraph "end". */
async function document() {
  const session = await createYrsSession();
  sessions.push(session);
  const { paraId: first } = session.createStory('body', 'onetwothree');
  session.applyRawOps('body', [
    { op: 'insertEmbed', index: 3, kind: 'tab' },
    { op: 'insertEmbed', index: 7, kind: 'break' },
  ]);
  const { secondParaId: last } = session.splitParagraph({ story: 'body', paraId: first, offset: 13 });
  const { table } = session.insertTable({ story: 'body', paraId: last, offset: 0 }, 2, 2);
  const cells = [
    [0, 0, 'a'],
    [0, 1, 'b'],
    [1, 0, 'c'],
    [1, 1, 'd'],
  ] as const;
  for (const [row, column, text] of cells) {
    const story = yrsCellStory(session, { ...table, row, column })!;
    session.insertText({ story, paraId: session.paragraphs(story)[0].paraId, offset: 0 }, text);
  }
  session.insertText({ story: 'body', paraId: last, offset: 1 }, 'end');
  return { session, first, last, table };
}

function syntheticTable(
  rows: { text: string; tcPr?: Record<string, unknown> }[][],
  grid?: unknown[]
) {
  const texts = new Map<string, string>();
  const payload = {
    grid,
    rows: rows.map((cells, row) => ({
      cells: cells.map((cell, column) => {
        const story = `body:t0:r${row}c${column}`;
        texts.set(story, cell.text);
        return { story, tcPr: cell.tcPr };
      }),
    })),
  };
  const session = {
    selection: () => ({
      anchor: { story: 'body', paraId: 'p', offset: 0 },
      head: { story: 'body', paraId: 'p', offset: 1 },
    }),
    cellSelection: () => null,
    locateParagraph: () => ({ start: 0, end: 1 }),
    tablePayload: () => payload,
    storySegments: (story) =>
      story === 'body'
        ? [{ kind: 'embed', embedKind: 'table', payload, attributes: {} }]
        : [{ kind: 'text', text: texts.get(story) ?? '', attributes: {} }],
  } satisfies Pick<
    YrsSession,
    'selection' | 'cellSelection' | 'locateParagraph' | 'tablePayload' | 'storySegments'
  >;
  return session as unknown as YrsSession;
}

function selectSyntheticRange(
  session: YrsSession,
  top: number,
  bottom: number,
  left: number,
  right: number
) {
  const loc = { story: 'body:t0:r0c0', paraId: 'p', offset: 0 };
  session.selection = () => ({ anchor: loc, head: loc });
  session.cellSelection = () => ({
    anchor: { story: 'body', tableIndex: 0, row: top, column: left },
    head: { story: 'body', tableIndex: 0, row: bottom, column: right },
  });
}

test('copies a tall table with an extreme grid span without grid padding', () => {
  const rows = Array.from({ length: 20_000 }, (_, row) => [
    { text: `row-${row}`, tcPr: row === 0 ? { colspan: 1_000_000_000 } : undefined },
  ]);
  const session = syntheticTable(rows);
  const text = yrsSelectionPlainText(session);
  expect(text.length).toBeLessThan(1 << 20);
  expect(text.split('\n')).toEqual([...rows.map(([cell]) => cell.text), '']);

  selectSyntheticRange(session, 1, rows.length - 1, 0, 16_383);
  const rangeText = yrsSelectionPlainText(session);
  expect(rangeText.length).toBeLessThan(1 << 20);
  expect(rangeText.split('\n')).toEqual(rows.slice(1).map(([cell]) => cell.text));

  selectSyntheticRange(session, 0, 1, 0, 1);
  expect(yrsSelectionPlainText(session)).toBe('row-0\t\nrow-1\t');
});

test('copies vertical and horizontal merges with empty merged-over fields', () => {
  const session = syntheticTable([
    [
      { text: 'a', tcPr: { rowspan: 2 } },
      { text: 'b', tcPr: { colspan: 2 } },
    ],
    [{ text: 'c', tcPr: { colspan: 2 } }],
    [{ text: 'd' }, { text: 'e' }, { text: 'f' }],
  ]);
  expect(yrsSelectionPlainText(session)).toBe('a\tb\t\n\tc\t\nd\te\tf\n');
  const table = { story: 'body', tableIndex: 0 };
  expect(yrsCellStory(session, { ...table, row: 1, column: 0 })).toBe('body:t0:r0c0');
  expect(yrsCellStory(session, { ...table, row: 0, column: 2 })).toBe('body:t0:r0c1');
  expect(yrsCellStory(session, { ...table, row: 1, column: 2 })).toBe('body:t0:r1c0');

  selectSyntheticRange(session, 0, 1, 0, 2);
  expect(yrsSelectionPlainText(session)).toBe('a\tb\t\n\tc\t');
});

test('carries multiple merge intervals in column order until they expire', () => {
  const session = syntheticTable([
    [
      { text: 'a' },
      { text: 'b', tcPr: { colspan: 2, rowspan: 3 } },
      { text: 'c' },
      { text: 'd', tcPr: { rowspan: 2 } },
    ],
    [{ text: 'e', tcPr: { rowspan: 2 } }, { text: 'f' }, { text: 'g' }],
    [{ text: 'h' }, { text: 'i' }],
  ]);
  expect(yrsSelectionPlainText(session)).toBe('a\tb\t\tc\td\t\ne\t\t\tf\t\tg\n\t\t\th\ti\t\n');
});

test('clamps spans to remaining rows and columns even for cells beyond the bound', () => {
  const session = syntheticTable([
    [
      { text: 'a', tcPr: { rowspan: 1_000_000_000, colspan: 1_000_000_000 } },
      { text: 'b', tcPr: { colspan: 1_000_000_000 } },
    ],
    [{ text: 'c', tcPr: { colspan: 1_000_000_000 } }],
  ]);
  const table = { story: 'body', tableIndex: 0 };
  expect(yrsCellStory(session, { ...table, row: 1, column: 16_383 })).toBe('body:t0:r0c0');
  expect(yrsCellStory(session, { ...table, row: 0, column: 16_384 })).toBe('body:t0:r0c1');
  expect(yrsCellStory(session, { ...table, row: 1, column: 16_384 })).toBe('body:t0:r1c0');
  expect(yrsCellStory(session, { ...table, row: 1, column: 16_385 })).toBeNull();
  expect(yrsCellStory(session, { ...table, row: 2, column: 0 })).toBeNull();
});

test('clamps an oversized declared table grid', () => {
  const session = syntheticTable([[{ text: 'a' }]], Array(20_000));
  const focused = { story: 'body', tableIndex: 0, row: 0, column: 0 };
  expect(yrsTableSelectionRange(session, focused, 'table')?.head.column).toBe(16_383);
});

test('copies tabs, line breaks and tables as text', async () => {
  const { session, first, last } = await document();
  session.setSelection(
    { story: 'body', paraId: first, offset: 0 },
    { story: 'body', paraId: last, offset: 4 }
  );
  expect(yrsSelectionPlainText(session)).toBe('one\ttwo\nthree\na\tb\nc\td\nend');
  expect(yrsSelectedText(session)).toBe('onetwothree\nend');

  session.setSelection(
    { story: 'body', paraId: first, offset: 5 },
    { story: 'body', paraId: first, offset: 1 }
  );
  expect(yrsSelectionPlainText(session)).toBe('ne\tt');

  session.setSelection({ story: 'body', paraId: first, offset: 2 });
  expect(yrsSelectionPlainText(session)).toBe('');
});

test('copies a cell range as tab-separated rows', async () => {
  const { session, table } = await document();
  const story = yrsCellStory(session, { ...table, row: 1, column: 1 })!;
  session.setSelection({ story, paraId: session.paragraphs(story)[0].paraId, offset: 0 });
  session.setCellSelection({
    anchor: { ...table, row: 0, column: 0 },
    head: { ...table, row: 1, column: 1 },
  });
  expect(yrsSelectionPlainText(session)).toBe('a\tb\nc\td');

  session.setCellSelection({
    anchor: { ...table, row: 0, column: 1 },
    head: { ...table, row: 1, column: 1 },
  });
  expect(yrsSelectionPlainText(session)).toBe('b\nd');
});

test('copies inline content controls and equations as their text', async () => {
  const session = await createYrsSession();
  sessions.push(session);
  const { paraId } = session.createStory('body', 'Name: , x');
  session.applyRawOps('body', [
    {
      op: 'insertEmbed',
      index: 6,
      kind: 'sdt',
      payload: { content: [{ kind: 'text', text: 'Alice' }, { kind: 'tab' }, { kind: 'text', text: 'B' }] },
    },
    { op: 'insertEmbed', index: 10, kind: 'math', payload: { plainText: 'a+b' } },
  ]);
  session.setSelection(
    { story: 'body', paraId, offset: 0 },
    { story: 'body', paraId, offset: 11 }
  );
  expect(yrsSelectionPlainText(session)).toBe('Name: Alice\tB, xa+b');
});

test('a cell holding a break or a quote copies as one quoted field', async () => {
  const { session, table } = await document();
  const first = yrsCellStory(session, { ...table, row: 0, column: 0 })!;
  const { secondParaId } = session.splitParagraph({
    story: first,
    paraId: session.paragraphs(first)[0].paraId,
    offset: 1,
  });
  session.insertText({ story: first, paraId: secondParaId, offset: 0 }, 'z');
  const second = yrsCellStory(session, { ...table, row: 0, column: 1 })!;
  const secondPara = session.paragraphs(second)[0].paraId;
  session.insertText({ story: second, paraId: secondPara, offset: 1 }, ' "x"');
  session.setSelection({ story: second, paraId: secondPara, offset: 0 });
  session.setCellSelection({
    anchor: { ...table, row: 0, column: 0 },
    head: { ...table, row: 1, column: 1 },
  });
  expect(yrsSelectionPlainText(session)).toBe('"a\nz"\t"b ""x"""\nc\td');
});
