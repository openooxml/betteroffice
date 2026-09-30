import { afterEach, beforeAll, expect, test } from 'bun:test';
import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';
import { preloadEditWasm } from '@betteroffice/docx/wasm/edit';
import { createYrsSession, type YrsSession } from '@betteroffice/docx/yrs';
import { yrsCellStory, yrsSelectedText, yrsSelectionPlainText } from './yrsCommands';

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
