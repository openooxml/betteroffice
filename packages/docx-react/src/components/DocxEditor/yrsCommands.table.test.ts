import { afterEach, beforeAll, expect, test } from 'bun:test';
import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';
import { preloadEditWasm } from '@betteroffice/docx/wasm/edit';
import { createYrsSession, type YrsSession } from '@betteroffice/docx/yrs';
import {
  currentYrsSplitCellConfig,
  currentYrsTableContext,
  currentYrsTableProperties,
  yrsCellStory,
  yrsTableSelectionStories,
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

test('table state reads its own table, not the whole story', async () => {
  const session = await createYrsSession();
  sessions.push(session);
  const { paraId } = session.createStory('body', 'before');
  session.insertTable({ story: 'body', paraId, offset: 0 }, 1, 1);
  const { table } = session.insertTable({ story: 'body', paraId, offset: 6 }, 2, 3);
  expect(table.tableIndex).toBe(1);

  const payloads = session
    .storySegments('body')
    .flatMap((segment) =>
      segment.kind === 'embed' && segment.embedKind === 'table' ? [segment.payload] : []
    );
  expect(payloads).toHaveLength(2);
  payloads.forEach((payload, index) => expect(session.tablePayload('body', index)).toEqual(payload));
  expect(session.tablePayload('body', 2)).toBeNull();
  expect(() => session.tablePayload('missing', 0)).toThrow();

  const cell = { ...table, row: 1, column: 2 };
  const story = yrsCellStory(session, cell)!;
  session.setSelection({ story, paraId: session.paragraphs(story)[0]!.paraId, offset: 0 });
  session.setCellSelection({ anchor: { ...cell, column: 1 }, head: cell });

  const storySegments = session.storySegments;
  let wholeStoryReads = 0;
  session.storySegments = (read) => {
    wholeStoryReads += 1;
    return storySegments(read);
  };
  expect(currentYrsTableContext(session)).toMatchObject({
    isInTable: true,
    rowIndex: 1,
    columnIndex: 2,
    rowCount: 2,
    columnCount: 3,
    hasMultiCellSelection: true,
  });
  expect(currentYrsTableProperties(session)).toEqual(
    payloads[1]!.tblPr as Record<string, unknown>
  );
  expect(currentYrsSplitCellConfig(session)).not.toBeNull();
  expect(yrsTableSelectionStories(session)).toEqual([
    yrsCellStory(session, { ...cell, column: 1 })!,
    story,
  ]);
  expect(wholeStoryReads).toBe(0);
});
