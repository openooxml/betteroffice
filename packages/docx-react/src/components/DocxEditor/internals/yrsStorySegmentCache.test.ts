import { beforeAll, expect, spyOn, test } from 'bun:test';
import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';
import { rezipPartsToArrayBuffer, toBytes } from '@betteroffice/docx/docx/rezip/parts';
import { createYrsSidebarProjection } from '@betteroffice/docx/layout/render';
import { preloadEditWasm } from '@betteroffice/docx/wasm/edit';
import { createYrsSession, type YrsSession, type YrsStorySegment } from '@betteroffice/docx/yrs';
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

for (const editBetweenSlices of [false, true]) {
  const name = editBetweenSlices
    ? 'idle digest slices replay an edit after the synchronous cache was warmed'
    : 'idle digest slices preserve story order and match synchronous completion';
  test(name, async () => {
    const session = await createYrsSession({ clientId: 77005 });
    const originalIdle = globalThis.requestIdleCallback;
    const originalCancelIdle = globalThis.cancelIdleCallback;
    const callbacks = new Map<number, IdleRequestCallback>();
    let nextIdle = 1;
    globalThis.requestIdleCallback = (run) => {
      const id = nextIdle++;
      callbacks.set(id, run);
      return id;
    };
    globalThis.cancelIdleCallback = (id) => { callbacks.delete(id); };
    let now = 0;
    const clock = spyOn(performance, 'now').mockImplementation(() => now);
    const cache = new YrsStorySegmentCache(session);
    const synchronous = new YrsStorySegmentCache(session);
    try {
      session.seedFromDocx(docx());
      const { reads, clear, storySegments } = counted(session);
      const digests = session.storySegmentUnitDigests.bind(session);
      session.storySegmentUnitDigests = (story) => {
        now += 4;
        return digests(story);
      };
      expectSameProjection(session, cache, storySegments);
      const stories = [...reads.whole];
      for (const story of stories.slice(1)) {
        const [paragraph] = session.paragraphs(story);
        session.splitParagraph({ story, paraId: paragraph!.paraId, offset: 0 });
      }
      cache.refresh();
      synchronous.refresh();
      for (const story of stories) {
        cache.segments(story);
        synchronous.segments(story);
      }
      clear();
      const flush = () => {
        expect(callbacks.size).toBe(1);
        const [id, run] = callbacks.entries().next().value!;
        callbacks.delete(id);
        run({ didTimeout: false, timeRemaining: () => 40 });
      };
      cache.scheduleDigests();
      const firstCallbackTime = now;
      synchronous.completeDigests();
      const expected = [...reads.digests];
      expect(expected).toEqual(stories);
      now = firstCallbackTime;
      clear();
      flush();
      expect(reads.digests).toEqual(stories.slice(0, 2));
      const completed = [...reads.digests];
      if (editBetweenSlices) {
        const story = stories[2]!;
        const [paragraph] = session.paragraphs(story);
        session.insertText({ story, paraId: paragraph!.paraId, offset: 0 }, 'xyz');
        for (const warmed of [cache, synchronous]) {
          warmed.refresh();
          warmed.segments(story);
        }
      }
      clear();
      flush();
      flush();
      expect([...completed, ...reads.digests]).toEqual(expected);
      expect(callbacks.size).toBe(0);
      expect(cache).toEqual(synchronous);
    } finally {
      cache.dispose();
      synchronous.dispose();
      clock.mockRestore();
      globalThis.requestIdleCallback = originalIdle;
      globalThis.cancelIdleCallback = originalCancelIdle;
      session.destroy();
    }
  });
}

for (const { name, noIdle, duringStory, idleDeadline, dispose } of [
  { name: 'busy continuations finish at the original deadline', noIdle: false },
  { name: 'fallback expiry during a story finishes the overdue queue', duringStory: true },
  { name: 'the idle deadline can expire during a story', idleDeadline: true },
  { name: 'timers continue digests without requestIdleCallback', noIdle: true },
  { name: 'disposal cancels a rescheduled digest', dispose: true },
  { name: 'disposal cancels a rescheduled digest without idle callbacks', noIdle: true, dispose: true },
]) {
  test(name, () => {
    const originalIdle = globalThis.requestIdleCallback;
    const originalCancelIdle = globalThis.cancelIdleCallback;
    const callbacks = new Map<number, { run: IdleRequestCallback; timeout?: number }>();
    const timers = new Map<number, { at: number; run: () => void }>();
    let next = 0;
    let now = 0;
    const clock = spyOn(performance, 'now').mockImplementation(() => now);
    const timeout = spyOn(globalThis, 'setTimeout').mockImplementation(
      ((run: () => void, delay = 0) => {
        const id = ++next;
        timers.set(id, { at: now + delay, run });
        return id;
      }) as unknown as typeof setTimeout
    );
    const clear = spyOn(globalThis, 'clearTimeout').mockImplementation((id) => {
      timers.delete(id as unknown as number);
    });
    globalThis.requestIdleCallback = noIdle
      ? undefined as unknown as typeof requestIdleCallback
      : (run, options) => {
        const id = ++next;
        callbacks.set(id, { run, timeout: options?.timeout });
        return id;
      };
    globalThis.cancelIdleCallback = (id) => { callbacks.delete(id); };
    const completed: string[] = [];
    const session = {
      storiesChangedSince: () => ({ revision: 0, stories: [] }),
      storySegments: (story: string) => [0, 1].map((index) => ({
        kind: 'pilcrow', paraId: `${story}:${index}`, properties: {}, attributes: {},
      })),
      storySegmentUnitDigests: (story: string) => {
        completed.push(story);
        now += idleDeadline ? 2 : 8;
        return [`${story}:0`, `${story}:1`];
      },
    } as unknown as YrsSession;
    const cache = new YrsStorySegmentCache(session);
    const flushTimers = () => {
      for (const [id, timer] of [...timers].sort((a, b) => a[1].at - b[1].at)) {
        if (timer.at > now || !timers.delete(id)) continue;
        timer.run();
      }
    };
    try {
      cache.segments('A');
      cache.segments('B');
      cache.scheduleDigests();
      now = duringStory ? 4999 : 4990;
      if (duringStory) cache.segments('C');
      if (noIdle) flushTimers();
      else {
        const [id, { run }] = callbacks.entries().next().value!;
        callbacks.delete(id);
        run({
          didTimeout: false,
          timeRemaining: () => idleDeadline && completed.length > 0 ? 0 : 40,
        });
      }
      if (duringStory) expect(completed).toEqual(['A', 'B', 'C']);
      else {
        expect(completed).toEqual(['A']);
        expect([...timers.values()].some(({ at }) => at === 5000)).toBe(true);
        if (!noIdle) {
          expect([...callbacks.values()].map(({ timeout }) => timeout)).toEqual([idleDeadline ? 8 : 2]);
        }
        cache.segments('C');
        cache.scheduleDigests();
        if (dispose) {
          cache.dispose();
          expect(callbacks.size).toBe(0);
          expect(timers.size).toBe(0);
        } else if (!noIdle) {
          now = 4999;
          flushTimers();
          expect(completed).toEqual(['A']);
        }
        now = 5000;
        flushTimers();
        expect(completed).toEqual(dispose ? ['A'] : ['A', 'B', 'C']);
      }
      expect(callbacks.size).toBe(0);
      expect(timers.size).toBe(0);
    } finally {
      cache.dispose();
      timeout.mockRestore();
      clear.mockRestore();
      clock.mockRestore();
      globalThis.requestIdleCallback = originalIdle;
      globalThis.cancelIdleCallback = originalCancelIdle;
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
