import { beforeAll, expect, test } from 'bun:test';
import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';
import { rezipPartsToArrayBuffer, toBytes } from '@betteroffice/docx/docx/rezip/parts';
import { preloadEditWasm } from '@betteroffice/docx/wasm/edit';
import { createYrsSession, type YrsSession, type YrsStorySegment } from '@betteroffice/docx/yrs';
import { createYrsPositionProjection } from './yrsPositionProjection';
import { YrsStorySegmentCache } from './yrsStorySegmentCache';

const WASM = resolve(
  import.meta.dir,
  '../../../../../docx/src/wasm/generated/edit/docx_edit_bg.wasm'
);
const W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"';
const OFFICE = 'application/vnd.openxmlformats-officedocument.wordprocessingml';

function docx(): Uint8Array {
  const parts = new Map<string, Uint8Array>();
  parts.set(
    '[Content_Types].xml',
    toBytes(
      `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="${OFFICE}.document.main+xml"/></Types>`
    )
  );
  parts.set(
    '_rels/.rels',
    toBytes(
      '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="r1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>'
    )
  );
  const p = (text: string) => `<w:p><w:r><w:t>${text}</w:t></w:r></w:p>`;
  const cell = (text: string) => `<w:tc>${p(text)}</w:tc>`;
  parts.set(
    'word/document.xml',
    toBytes(
      `<w:document ${W}><w:body>${p('First paragraph')}<w:tbl><w:tblGrid><w:gridCol w:w="2000"/><w:gridCol w:w="2000"/></w:tblGrid><w:tr>${cell('a')}${cell('b')}</w:tr><w:tr>${cell('c')}${cell('d')}</w:tr></w:tbl>${p('Second paragraph')}${p('Third')}</w:body></w:document>`
    )
  );
  return new Uint8Array(rezipPartsToArrayBuffer(parts));
}

beforeAll(() => preloadEditWasm(new Uint8Array(readFileSync(WASM))));

/** Records the segment reads `session` answers. */
function counted(session: YrsSession) {
  const reads = { whole: [] as string[], digests: [] as string[], units: [] as number[][] };
  const whole = session.storySegments.bind(session);
  const digests = session.storySegmentUnitDigests.bind(session);
  const units = session.storySegmentUnits.bind(session);
  session.storySegments = (story) => {
    reads.whole.push(story);
    return whole(story);
  };
  session.storySegmentUnitDigests = (story) => {
    reads.digests.push(story);
    return digests(story);
  };
  session.storySegmentUnits = (story, requested) => {
    reads.units.push([...requested]);
    return units(story, requested);
  };
  const clear = () => {
    reads.whole.length = 0;
    reads.digests.length = 0;
    reads.units.length = 0;
  };
  return { reads, clear, storySegments: whole };
}

function expectSameProjection(
  session: YrsSession,
  cache: YrsStorySegmentCache,
  storySegments: (story: string) => YrsStorySegment[]
) {
  cache.refresh();
  const cached = createYrsPositionProjection(session, 'body', cache)!;
  const fresh = createYrsPositionProjection(session, 'body', { segments: storySegments })!;
  expect(cached.size).toBe(fresh.size);
  for (let position = 0; position <= fresh.size; position++) {
    expect(cached.targetAt(position)).toEqual(fresh.targetAt(position));
    expect(cached.nodeAt(position)).toEqual(fresh.nodeAt(position));
    expect(cached.tableAtPosition(position)).toEqual(fresh.tableAtPosition(position));
  }
}

test('a cached projection re-reads only the paragraphs an edit changed', async () => {
  const session = await createYrsSession({ clientId: 77001 });
  try {
    session.seedFromDocx(docx());
    const cache = new YrsStorySegmentCache(session);
    const { reads, clear, storySegments } = counted(session);
    expectSameProjection(session, cache, storySegments);
    const cells = reads.whole.filter((story) => story.startsWith('body:t0:'));
    expect(reads.whole).toEqual(['body', ...cells]);
    expect(cells).toHaveLength(4);
    clear();
    cache.completeDigests();
    expect(reads.digests).toEqual(['body']);
    const [first, again] = session.storySegmentUnits('body', [0, 0]);
    expect(again).toEqual(first!);
    expect(first!.length).toBeGreaterThan(0);

    const [, second] = session.paragraphs('body');
    session.insertText({ story: 'body', paraId: second!.paraId, offset: 3 }, 'xyz');
    clear();
    expectSameProjection(session, cache, storySegments);
    expect(reads.digests).toEqual(['body']);
    expect(reads.units).toEqual([[1]]);
    expect(reads.whole).toEqual([]);

    const cellStory = cells[0]!;
    const [cellParagraph] = session.paragraphs(cellStory);
    session.splitParagraph({ story: cellStory, paraId: cellParagraph!.paraId, offset: 0 });
    clear();
    expectSameProjection(session, cache, storySegments);
    expect(reads.whole).toEqual([cellStory]);
    expect(reads.digests).toEqual([]);

    clear();
    expectSameProjection(session, cache, storySegments);
    expect(reads.digests).toEqual([]);
    expect(reads.whole).toEqual([]);
  } finally {
    session.destroy();
  }
});
