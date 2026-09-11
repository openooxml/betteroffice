/**
 * A caret inside a note, resolved against a display list the real layout wasm
 * built. The note's own story answers, so the geometry has to land in the note
 * area and nowhere else.
 */

import { afterAll, beforeAll, describe, expect, test } from 'bun:test';
import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';

import { buildDisplayListJson, preloadLayoutWasm } from '../../wasm/layout';
import { createDisplayListQueries, type DisplayListQueries } from './displayListQueries';
import type { DisplayList } from './displayList';

// two footnotes in one area at y 900..960, each story starting at position 1
const FIXTURE = resolve(
  import.meta.dir,
  '../../../../../crates/docx-layout/tests/fixtures/notes/two-footnotes.input.json'
);

let list: DisplayList;
let queries: DisplayListQueries;

beforeAll(async () => {
  await preloadLayoutWasm();
  list = JSON.parse(buildDisplayListJson(readFileSync(FIXTURE, 'utf8'))) as DisplayList;
  queries = createDisplayListQueries(list);
  await queries.whenReady();
});

afterAll(() => queries.dispose());

/** First position of the note story that covers nothing — its end. */
function endOfNote(noteId: number): number {
  for (let position = 1; position < 200; position += 1) {
    if (queries.noteRangeRects('footnote', noteId, position, position + 1).length === 0) {
      return position;
    }
  }
  throw new Error(`footnote ${noteId} never ends`);
}

describe('noteCaretRects', () => {
  test('a caret inside a note sits on the leading edge of its position', () => {
    const covered = queries.noteRangeRects('footnote', 1, 1, 2);
    expect(covered).toHaveLength(1);

    const caret = queries.noteCaretRects('footnote', 1, 1);
    expect(caret).toHaveLength(1);
    expect(caret[0]!.width).toBe(0);
    expect(caret[0]!.x).toBe(covered[0]!.x);
    expect(caret[0]!.pageIndex).toBe(covered[0]!.pageIndex);
  });

  // A note paints on exactly one page, so there is never a second candidate to
  // choose between — unlike a header, which repeats.
  test('a note answers on the one page it paints on', () => {
    const area = list.pages[0].noteAreas?.[0];
    if (!area) throw new Error('the fixture page has no note area');
    const caret = queries.noteCaretRects('footnote', 1, 1);

    expect(caret).toHaveLength(1);
    expect(caret[0]!.pageIndex).toBe(0);
    expect(caret[0]!.y).toBeGreaterThanOrEqual(area.y ?? 0);
    expect(caret[0]!.y).toBeLessThanOrEqual((area.y ?? 0) + (area.height ?? 0));
  });

  test('a caret at the end of a note falls back to the trailing edge', () => {
    const end = endOfNote(1);
    const last = queries.noteRangeRects('footnote', 1, end - 1, end);
    expect(last).toHaveLength(1);

    const caret = queries.noteCaretRects('footnote', 1, end);
    expect(caret).toHaveLength(1);
    expect(caret[0]!.width).toBe(0);
    expect(caret[0]!.x).toBe(last[0]!.x + last[0]!.width);
  });

  test('a note the page does not carry has no caret', () => {
    expect(queries.noteCaretRects('endnote', 1, 1)).toEqual([]);
  });
});
