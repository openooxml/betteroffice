import { expect, test } from 'bun:test';
import { createUnbuiltPageMirror, scheduleUnbuiltMirrorIdle, UNBUILT_MIRROR_AX_BUDGET, type MirrorPageText, type MirrorStamp, type PageTexts, type UnbuiltMirrorPage, type UnbuiltMirrorSnapshot } from './unbuiltPageMirror';

function fixture(count = 40, budget = UNBUILT_MIRROR_AX_BUDGET, cost = 5) {
  let blocked = false;
  let focused: number | null = null;
  let frameEpoch = 1;
  let layoutRevision = 1;
  let center = 0;
  let enabled = true;
  let snapshot: UnbuiltMirrorSnapshot<string> = { sections: [], ranges: [], nodeCount: 0 };
  const idle: ((deadline: Pick<IdleDeadline, 'timeRemaining'>) => void)[] = [];
  const reads: { pages: number[]; stamp: MirrorStamp; resolve(answer: PageTexts | null): void }[] = [];
  const built: number[] = [];
  const published: UnbuiltMirrorSnapshot<string>[] = [];
  const pages: UnbuiltMirrorPage[] = Array.from({ length: count }, (_, pageIndex) => ({ id: String(pageIndex), pageIndex, source: {}, revision: 0, unbuilt: true }));
  const mirror = createUnbuiltPageMirror({
    budget,
    idle(run) { idle.push(run); return () => { const at = idle.indexOf(run); if (at >= 0) idle.splice(at, 1); }; },
    blocked: () => blocked,
    focusedPage: () => focused,
    query: (pages, stamp) => new Promise<PageTexts | null>((resolve) => { reads.push({ pages, stamp, resolve }); }),
    build(page: MirrorPageText) { built.push(page.pageIndex); return { value: page.paragraphs[0]!.text, nodes: cost }; },
    publish(next) { snapshot = next; published.push(next); },
  });
  const update = () => mirror.update({ identity: 'session:1:1', frameEpoch, layoutRevision, center, pages, enabled });
  update();
  return {
    mirror, idle, reads, built, pages, published,
    get snapshot() { return snapshot; },
    setBlocked(value: boolean) { blocked = value; },
    focus(page: number | null) { focused = page; },
    move(page: number) { center = page; update(); },
    off() { enabled = false; update(); },
    frame(change?: (pages: UnbuiltMirrorPage[]) => void) { frameEpoch += 1; layoutRevision += 1; change?.(pages); update(); },
    slice(remaining = () => 40) { const run = idle.shift(); expect(run).toBeDefined(); run!({ timeRemaining: remaining }); },
    async answer(at = reads.length - 1, stamp = reads[at]!.stamp, suffix = '') {
      const read = reads[at]!;
      read.resolve({ ...stamp, pages: read.pages.map((pageIndex) => ({ pageIndex, paragraphs: [{ text: `Page ${pageIndex}${suffix}` }] })) });
      await Promise.resolve(); await Promise.resolve(); await Promise.resolve();
    },
    async fill() {
      for (let step = 0; idle.length && step < count * 4 + 20; step += 1) {
        const before = reads.length;
        this.slice();
        if (reads.length > before) await this.answer();
      }
      expect(idle).toHaveLength(0);
    },
  };
}

test('an empty frame has no mirror nodes, ranges or reads', async () => {
  const f = fixture(0);
  await f.fill();
  expect(f.snapshot).toEqual({ sections: [], ranges: [], nodeCount: 0 });
  expect(f.reads).toEqual([]);
});

test('fills in page order with at most sixteen pages in one read per idle slice', async () => {
  const f = fixture();
  f.slice();
  expect(f.reads.map(({ pages }) => pages)).toEqual([Array.from({ length: 16 }, (_, index) => index)]);
  expect(f.built).toEqual([]);
  await f.answer();
  await f.fill();
  expect(f.built).toEqual(Array.from({ length: 40 }, (_, index) => index));
  expect(f.reads.every(({ pages }) => pages.length <= 16)).toBe(true);
  expect(f.snapshot.nodeCount).toBe(200);
  expect(f.snapshot.ranges).toEqual([]);
});

test('publishes a completed paragraph batch once per idle slice', async () => {
  const f = fixture(4);
  f.slice();
  await f.answer();
  const before = f.published.length;
  f.slice();
  expect(f.built).toEqual([0, 1, 2, 3]);
  expect(f.published.slice(before)).toEqual([{
    sections: [0, 1, 2, 3].map((pageIndex) => ({ pageIndex, value: `Page ${pageIndex}` })),
    ranges: [], nodeCount: 20,
  }]);
});

test('stops reading and building when the deadline has three milliseconds left', async () => {
  const f = fixture(4);
  f.slice(() => 3);
  expect(f.reads).toEqual([]);
  f.slice();
  await f.answer();
  let checks = 0;
  f.slice(() => ++checks <= 2 ? 4 : 3);
  expect(f.built).toEqual([0, 1]);
  expect(f.reads).toHaveLength(1);
  await f.fill();
  expect(f.built).toEqual([0, 1, 2, 3]);
});

test('queues behind pending edits and frames without issuing a read', async () => {
  const f = fixture(3);
  f.setBlocked(true);
  f.slice();
  f.slice();
  expect(f.reads).toEqual([]);
  f.setBlocked(false);
  f.slice();
  expect(f.reads[0]!.pages).toEqual([0, 1, 2]);
  await f.answer();
  f.setBlocked(true);
  f.slice();
  expect(f.built).toEqual([]);
  f.setBlocked(false);
  await f.fill();
});

test('drops an older answer even when the page itself did not change and reads it again', async () => {
  const f = fixture(1);
  f.slice();
  f.frame();
  await f.answer(0, { frameEpoch: 1, layoutRevision: 1 }, ' obsolete');
  f.slice();
  expect(f.built).toEqual([]);
  expect(f.snapshot.sections).toEqual([]);
  expect(f.reads).toHaveLength(2);
  expect(f.reads[1]!.stamp).toEqual({ frameEpoch: 2, layoutRevision: 2 });
  await f.answer(1);
  await f.fill();
  expect(f.snapshot.sections).toEqual([{ pageIndex: 0, value: 'Page 0' }]);
});

test('drops stale frame epochs and layout revisions independently', async () => {
  for (const stamp of [{ frameEpoch: 0, layoutRevision: 1 }, { frameEpoch: 1, layoutRevision: 0 }]) {
    const f = fixture(1);
    f.slice();
    await f.answer(0, stamp, ' obsolete');
    f.slice();
    expect(f.built).toEqual([]);
    expect(f.reads).toHaveLength(2);
    await f.answer(1);
    await f.fill();
    expect(f.snapshot.sections[0]!.value).toBe('Page 0');
  }
});

test('refreshes only changed pages on a delta and keeps unchanged sections', async () => {
  const f = fixture(4);
  await f.fill();
  const reads = f.reads.length;
  f.frame((pages) => { pages[2] = { ...pages[2]!, source: {}, revision: 1 }; });
  expect(f.snapshot.sections).toHaveLength(4);
  f.slice();
  expect(f.reads.slice(reads).map(({ pages }) => pages)).toEqual([[2]]);
  await f.answer(reads, f.reads[reads]!.stamp, ' changed');
  await f.fill();
  expect(f.snapshot.sections.map(({ value }) => value)).toEqual(['Page 0', 'Page 1', 'Page 2 changed', 'Page 3']);
});

test('adds and removes pages and refreshes renumbered identities', async () => {
  const f = fixture(3);
  await f.fill();
  f.frame((pages) => {
    pages.splice(1, 1);
    pages[1] = { ...pages[1]!, pageIndex: 1 };
    pages.push({ id: 'new', pageIndex: 2, source: {}, revision: 0, unbuilt: true });
  });
  f.slice();
  expect(f.reads.at(-1)!.pages).toEqual([1, 2]);
  await f.answer();
  await f.fill();
  expect(f.snapshot.sections.map(({ pageIndex }) => pageIndex)).toEqual([0, 1, 2]);
});

test('counts actual built node costs and moves a bounded window with explicit edges', async () => {
  const f = fixture(30, 50, 9);
  await f.fill();
  expect(f.snapshot.nodeCount).toBeLessThanOrEqual(50);
  expect(f.snapshot.sections.map(({ pageIndex }) => pageIndex)).toEqual([0, 1, 2, 3, 4]);
  expect(f.snapshot.ranges).toEqual([{ from: 6, to: 30 }]);
  f.move(15);
  await f.fill();
  expect(f.snapshot.nodeCount).toBeLessThanOrEqual(50);
  expect(f.snapshot.sections.map(({ pageIndex }) => pageIndex)).toEqual([13, 14, 15, 16]);
  expect(f.snapshot.ranges).toEqual([{ from: 1, to: 13 }, { from: 18, to: 30 }]);
  expect(f.snapshot.sections.some(({ value }) => value === 'Page 0')).toBe(false);
  f.move(f.snapshot.ranges[1]!.from - 1);
  await f.fill();
  expect(f.snapshot.sections.some(({ pageIndex }) => pageIndex === 17)).toBe(true);
});

test('moving the window never drops its focused section', async () => {
  const f = fixture(30, 50, 9);
  await f.fill();
  const pinned = f.snapshot.sections[0];
  f.focus(0);
  f.move(20);
  await f.fill();
  expect(f.snapshot.sections).toContainEqual(pinned);
  expect(f.snapshot.sections.some(({ pageIndex }) => pageIndex === 20)).toBe(true);
  expect(f.snapshot.nodeCount).toBeLessThanOrEqual(50);
  f.focus(null);
  f.move(21);
  await f.fill();
  expect(f.snapshot.sections.some(({ pageIndex }) => pageIndex === 0)).toBe(false);
});

test('off cancels slices, renders no sections or edges, and ignores an in-flight reply', async () => {
  const f = fixture(20);
  f.slice();
  f.off();
  await f.answer();
  expect(f.snapshot).toEqual({ sections: [], ranges: [], nodeCount: 0 });
  expect(f.idle).toEqual([]);
  expect(f.built).toEqual([]);
});

test('a frame that changes no section, range or node count publishes nothing', async () => {
  const f = fixture(4);
  await f.fill();
  const published = f.published.length;
  f.frame();
  f.frame();
  f.move(0);
  expect(f.published).toHaveLength(published);
  f.frame((pages) => { pages[1] = { ...pages[1]!, source: {}, revision: 1 }; });
  f.slice();
  await f.answer(f.reads.length - 1, f.reads.at(-1)!.stamp, ' changed');
  await f.fill();
  expect(f.published).toHaveLength(published + 1);
  expect(f.snapshot.sections.map(({ value }) => value)).toEqual(['Page 0', 'Page 1 changed', 'Page 2', 'Page 3']);
});

test('off publishes once when sections go away and never on later frames', async () => {
  const empty = fixture(0);
  empty.off();
  empty.frame();
  empty.off();
  expect(empty.published).toEqual([]);
  const unread = fixture(20);
  const before = unread.published.length;
  unread.off();
  unread.frame();
  unread.off();
  expect(unread.published.slice(before)).toEqual([{ sections: [], ranges: [], nodeCount: 0 }]);
  const f = fixture(4);
  await f.fill();
  const published = f.published.length;
  f.off();
  f.frame();
  f.off();
  expect(f.published.slice(published)).toEqual([{ sections: [], ranges: [], nodeCount: 0 }]);
});

test('the timer fallback uses the injected clock and can be cancelled without timer spies', () => {
  const originalIdle = globalThis.requestIdleCallback;
  const originalTimeout = globalThis.setTimeout;
  const originalClear = globalThis.clearTimeout;
  let run: (() => void) | undefined;
  let cancelled = false;
  let now = 10;
  let remaining = 0;
  try {
    globalThis.requestIdleCallback = undefined as unknown as typeof requestIdleCallback;
    globalThis.setTimeout = ((callback: () => void) => { run = callback; return 1; }) as unknown as typeof setTimeout;
    globalThis.clearTimeout = (() => { cancelled = true; }) as typeof clearTimeout;
    const cancel = scheduleUnbuiltMirrorIdle((deadline) => {
      remaining = deadline.timeRemaining();
      now += 5;
      expect(deadline.timeRemaining()).toBe(3);
    }, () => now);
    run!();
    expect(remaining).toBe(8);
    cancel();
    expect(cancelled).toBe(true);
  } finally {
    globalThis.requestIdleCallback = originalIdle;
    globalThis.setTimeout = originalTimeout;
    globalThis.clearTimeout = originalClear;
  }
});
