import { beforeAll, expect, test } from 'bun:test';
import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';
import { rezipPartsToArrayBuffer, toBytes } from '@betteroffice/docx/docx/rezip/parts';
import { createYrsSidebarProjection } from '@betteroffice/docx/layout/render';
import { preloadEditWasm } from '@betteroffice/docx/wasm/edit';
import { createYrsSession, type YrsSession, type YrsStorySegment } from '@betteroffice/docx/yrs';
import { checkIdleContinuation, fakeIdleScheduler } from '../__fixtures__/fakeIdleScheduler';
import { createYrsPositionProjection } from './yrsPositionProjection';
import { storySegmentSource, YrsStorySegmentCache } from './yrsStorySegmentCache';

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

test('refreshing an unchanged revision keeps stale paragraphs available for digest reads', async () => {
  const session = await createYrsSession({ clientId: 77002 });
  try {
    session.seedFromDocx(docx());
    const cache = new YrsStorySegmentCache(session);
    const { reads, clear, storySegments } = counted(session);
    cache.refresh();
    cache.segments('body');
    cache.completeDigests();

    const [, second] = session.paragraphs('body');
    session.insertText({ story: 'body', paraId: second!.paraId, offset: 3 }, 'xyz');
    clear();
    cache.refresh();
    cache.refresh();
    expect(cache.segments('body')).toEqual(storySegments('body'));
    expect(reads.digests).toEqual(['body']);
    expect(reads.units).toEqual([[1]]);
    expect(reads.whole).toEqual([]);
  } finally {
    session.destroy();
  }
});

test('a sidebar projection read through the cache re-reads only the edited paragraph', async () => {
  const session = await createYrsSession({ clientId: 77003 });
  try {
    session.seedFromDocx(docx());
    const direct = Object.create(session) as YrsSession;
    direct.storySegments = session.storySegments.bind(session);
    const cache = new YrsStorySegmentCache(session);
    const { reads, clear } = counted(session);
    const paragraphs = session.paragraphs('body');
    const locations = paragraphs.map(({ paraId }) => ({ story: 'body', paraId, offset: 1 }));
    const first = createYrsSidebarProjection(session, storySegmentSource(session, cache));
    const firstDirect = createYrsSidebarProjection(direct);
    for (const loc of locations) {
      expect(first.locToDisplayPoint(loc)).toEqual(firstDirect.locToDisplayPoint(loc));
    }
    cache.completeDigests();

    session.insertText({ story: 'body', paraId: paragraphs[1]!.paraId, offset: 3 }, 'xyz');
    clear();
    const edited = createYrsSidebarProjection(session);
    const editedDirect = createYrsSidebarProjection(direct);
    for (const loc of locations) {
      expect(edited.locToDisplayPoint(loc)).toEqual(editedDirect.locToDisplayPoint(loc));
    }
    expect(edited).not.toBe(first);
    expect(reads.whole).toEqual([]);
    expect(reads.units).toEqual([[1]]);
  } finally {
    session.destroy();
  }
});

for (const editBetweenSlices of ['none', 'reread', 'without rereading'] as const) {
  const name = editBetweenSlices === 'none'
    ? 'idle digest slices preserve story order and match synchronous completion'
    : `idle digest slices preserve warming after an edit ${editBetweenSlices}`;
  test(name, async () => {
    const session = await createYrsSession({ clientId: 77005 });
    const scheduler = fakeIdleScheduler({ timers: true });
    const cache = new YrsStorySegmentCache(session);
    const synchronous = new YrsStorySegmentCache(session);
    try {
      session.seedFromDocx(docx());
      const { reads, clear, storySegments } = counted(session);
      const digests = session.storySegmentUnitDigests.bind(session);
      session.storySegmentUnitDigests = (story) => {
        scheduler.now += 4;
        return digests(story);
      };
      expectSameProjection(session, cache, storySegments);
      const stories = [...reads.whole];
      for (const story of stories.slice(1)) {
        const [paragraph] = session.paragraphs(story);
        session.splitParagraph({ story, paraId: paragraph!.paraId, offset: 0 });
      }
      for (const warmed of [cache, synchronous]) {
        warmed.refresh();
        for (const story of stories) warmed.segments(story);
      }
      clear();
      cache.scheduleDigests();
      const firstCallbackTime = scheduler.now;
      synchronous.completeDigests();
      const expected = [...reads.digests];
      expect(expected).toEqual(stories);
      scheduler.now = firstCallbackTime;
      clear();
      scheduler.flushOneIdle();
      expect(reads.digests).toEqual(stories.slice(0, 2));
      const completed = [...reads.digests];
      if (editBetweenSlices !== 'none') {
        const story = stories[2]!;
        const [paragraph] = session.paragraphs(story);
        session.insertText({ story, paraId: paragraph!.paraId, offset: 0 }, 'xyz');
        if (editBetweenSlices === 'reread') {
          for (const warmed of [cache, synchronous]) {
            warmed.refresh();
            warmed.segments(story);
          }
        }
      }
      clear();
      scheduler.flushOneIdle();
      scheduler.flushOneIdle();
      expect([...completed, ...reads.digests]).toEqual(expected);
      expect(scheduler.idleWork.size).toBe(0);
      if (editBetweenSlices === 'without rereading') {
        const story = stories[2]!;
        expect(reads.whole).toEqual([story]);
        clear();
        const resumed = cache.segments(story);
        expect(resumed).toEqual(storySegments(story));
        expect({ whole: reads.whole, digests: reads.digests }).toEqual({ whole: [], digests: [] });
        synchronous.refresh();
        expect(synchronous.segments(story)).toEqual(resumed);
      }
      expect(cache).toEqual(synchronous);
    } finally {
      cache.dispose();
      synchronous.dispose();
      scheduler.restore();
      session.destroy();
    }
  });
}

for (const scenario of [
  'busy', 'no idle', 'dispose', 'dispose no idle', 'expiry during story', 'idle deadline', 'deleted story',
]) {
  const name = scenario === 'deleted story'
    ? 'idle digest slices skip a deleted story and warm its surviving neighbour'
    : `digest continuation: ${scenario}`;
  test(name, () => {
    const noIdle = scenario.includes('no idle');
    const scheduler = fakeIdleScheduler({ timers: true, noIdle });
    const completed: string[] = [];
    const live = new Set(['A', 'B', 'C']);
    const session = {
      hasStory: (story: string) => live.has(story),
      storiesChangedSince: () => ({ revision: 0, stories: live.has('B') ? [] : ['B'] }),
      storySegments: (story: string) => {
        if (!live.has(story)) throw new Error(`Missing story: ${story}`);
        return [0, 1].map((index) => ({
          kind: 'pilcrow', paraId: `${story}:${index}`, properties: {}, attributes: {},
        }));
      },
      storySegmentUnitDigests: (story: string) => {
        completed.push(story);
        if (scenario === 'expiry during story' || story === 'A') {
          scheduler.now += scenario === 'idle deadline' ? 2 : 8;
        }
        return [`${story}:0`, `${story}:1`];
      },
    } as unknown as YrsSession;
    const cache = new YrsStorySegmentCache(session);
    try {
      for (const story of ['A', 'B', 'C']) cache.segments(story);
      cache.scheduleDigests();
      expect(() => checkIdleContinuation(scheduler, scenario, {
        setNow: scheduler.setNow, dispose: () => cache.dispose(),
        checkFirst: () => {
          expect(completed).toEqual(['A']);
          if (scenario === 'deleted story') live.delete('B');
        },
        timerSlices: scenario === 'expiry during story' ? 2 : 1, resumeIdle: scenario === 'deleted story',
        deadline: {
          didTimeout: false,
          timeRemaining: () => scenario === 'idle deadline' && completed.length > 0 ? 0 : 40,
        },
        checkTimer: (index) => {
          const expected = scenario === 'deleted story' ? ['A', 'C']
            : scenario === 'expiry during story' && index === 0 ? ['A', 'B'] : ['A', 'B', 'C'];
          expect(completed).toEqual(expected);
          if (scenario === 'deleted story') {
            const fresh = new YrsStorySegmentCache(session);
            for (const story of ['A', 'C']) fresh.segments(story);
            fresh.completeDigests();
            expect(cache).toEqual(fresh);
            fresh.dispose();
          } else if (scenario !== 'expiry during story') {
            expect(scheduler.now).toBe(noIdle ? 8 : 5000);
          }
        },
      })).not.toThrow();
    } finally {
      cache.dispose();
      scheduler.restore();
    }
  });
}

test('a source whose cache was disposed reads the session directly and holds no segments', async () => {
  const session = await createYrsSession({ clientId: 77004 });
  try {
    session.seedFromDocx(docx());
    const cache = new YrsStorySegmentCache(session);
    const source = storySegmentSource(session, cache);
    source.segments('body');
    cache.dispose();
    const { reads } = counted(session);
    expect(source.segments('body')).toEqual(session.storySegments('body'));
    expect(reads.whole).toEqual(['body', 'body']);
    expect(reads.digests).toEqual([]);
  } finally {
    session.destroy();
  }
});
