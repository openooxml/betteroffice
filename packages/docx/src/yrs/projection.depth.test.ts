import { beforeAll, describe, expect, it } from 'bun:test';

import { toBytes } from '../docx/rezip/parts';
import type { BlockContent } from '../types/document';
import { createEditSession } from '../wasm/edit';
import {
  buildDocxFromParts,
  canonicalStories,
  check,
  checkComposition,
  checkStory,
  nativeProjection,
  nativeStoryProjection,
  preloadSeedParityWasm,
  type Deferral,
} from './projectionParity.support';
import { wrapSession } from './yrsSessionFacade';
import { yrsToDocument } from './yrsToDocument';

const CLIENT = 48001;
const TABLE_LEVELS = 5;
const DEEP_MARKER = 'innermost-projection-depth-marker';

function syntheticDocx(body: string): Uint8Array {
  return buildDocxFromParts(new Map([['word/document.xml', toBytes(
    `<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>${body}</w:body></w:document>`
  )]]));
}

function nestedTables(levels: number): string {
  const open = '<w:tbl><w:tblGrid><w:gridCol w:w="1440"/></w:tblGrid><w:tr><w:tc>';
  const close = '</w:tc></w:tr></w:tbl>';
  const field = '<w:p><w:fldSimple w:instr="DATE"><w:r><w:t>field result</w:t></w:r></w:fldSimple></w:p>';
  return `${field}${open.repeat(levels)}<w:p><w:r><w:t>${DEEP_MARKER}</w:t></w:r></w:p>${close.repeat(levels)}<w:p/>`;
}

function expectDeepMarker(blocks: readonly BlockContent[], levels: number): void {
  for (let depth = 0; depth < levels; depth++) {
    const table = blocks.find((block) => block.type === 'table');
    if (!table) throw new Error(`missing table level ${depth + 1}`);
    blocks = table.rows[0]!.cells[0]!.content;
  }
  expect(blocks.some((block) => block.type === 'table')).toBe(false);
  expect(JSON.stringify(blocks)).toContain(DEEP_MARKER);
}

describe('Native DOCX projection depth', () => {
  beforeAll(preloadSeedParityWasm, 30_000);

  it('carries five nested parsed tables through both native JSON exports', () => {
    const raw = createEditSession(CLIENT);
    const session = wrapSession(raw, CLIENT);
    try {
      session.seedFromDocx(syntheticDocx(nestedTables(TABLE_LEVELS)));
      const base = session.materializeDocx()!;
      const source = canonicalStories(session, base);
      const ts = yrsToDocument(session, base);
      const native = nativeProjection(raw, base.originalBuffer!);
      const story = nativeStoryProjection(raw, 'body');
      expect(check(session, ts, native, source)).toEqual([]);
      expect(checkStory(canonicalStories(session, ts), story, source)).toEqual([]);
      expect(checkComposition(session, [story], native)).toEqual([]);
      expectDeepMarker(native.document.package.document.content, TABLE_LEVELS);
      expectDeepMarker(story.blocks, TABLE_LEVELS);
    } finally {
      session.destroy();
    }
  }, 30_000);

  for (const levels of [1, TABLE_LEVELS]) {
    it(`reports all ${levels} session-created tables as deferred`, () => {
      const raw = createEditSession(CLIENT);
      const session = wrapSession(raw, CLIENT);
      try {
        session.seedFromDocx(syntheticDocx('<w:p/>'));
        const base = session.materializeDocx()!;
        const source = canonicalStories(session, base);
        let storyId = 'body';
        const expected: Deferral[] = [];
        for (let depth = 0; depth < levels; depth++) {
          const { paraId } = session.paragraphs(storyId)[0]!;
          const receipt = session.insertTable({ story: storyId, paraId, offset: 0 }, 1, 1);
          expected.push({ family: 'tables', kind: 'table', storyId, paragraph: 0 });
          storyId = receipt.createdStoryIds[0]!;
        }
        const { paraId } = session.paragraphs(storyId)[0]!;
        session.insertText({ story: storyId, paraId, offset: 0 }, DEEP_MARKER);
        const ts = yrsToDocument(session, base);
        expectDeepMarker(ts.package.document.content, levels);

        const native = nativeProjection(raw, base.originalBuffer!);
        const story = nativeStoryProjection(raw, 'body');
        expect(native.deferred.filter(({ family }) => family === 'tables')).toEqual(expected);
        expect(story.deferred.filter(({ family }) => family === 'tables')).toEqual(expected);
        expect(check(session, ts, native, source)).toEqual([]);
        expect(checkStory(canonicalStories(session, ts), story, source)).toEqual([]);
        expect(checkComposition(session, [story], native)).toEqual([]);
      } finally {
        session.destroy();
      }
    }, 60_000);
  }
});
