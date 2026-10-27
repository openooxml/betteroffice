import { expect, test, type Page } from 'playwright/test';

type State = 'proposed' | 'accepted' | 'rejected';
interface Probe {
  proposed: { ok: boolean; snapshot?: { proposals: { id: string; state: State; changed: boolean }[] } } | null;
  timings: { firstPagePainted: number | null; firstInk: number | null; preview: number | null; load: number | null; layoutComplete: number | null };
  errors: string[];
  decide(ids: string[], state: State): Promise<{ ok: boolean }>;
  undoDecision(): Promise<{ ok: boolean }>;
  next(index: number): Promise<{ ok: boolean } | null>;
  jumpFirst(): Promise<{ id: string; pageIndex: number; result: { ok: boolean } | null } | null>;
  pageInked(pageIndex: number): boolean;
  firstPage(): string | null;
  firstPagePaintedCalls: number;
  started: number;
}
type ProbeWindow = Window & { __viewerHostProposalProbe: Probe; __commandInitHeld: boolean; __releaseCommandInit(): void; __posted: Record<string, number>;
  __wholeSourceHeld: boolean; __releaseWholeSource(): void; __holdReviews: boolean; __releaseReviews(): void };
const ids = Array.from({ length: 10 }, (_, index) => `host-${index + 1}`);

const read = (page: Page) => page.evaluate(() => {
  const probe = (window as unknown as ProbeWindow).__viewerHostProposalProbe;
  return { proposed: probe.proposed, timings: probe.timings, errors: probe.errors };
});
const settled = (page: Page) => read(page).catch(() => null);
const firstPage = (page: Page) => page.evaluate(() => (window as unknown as ProbeWindow).__viewerHostProposalProbe.firstPage());

async function holdCommandInit(page: Page) {
  await page.addInitScript(() => {
    const scope = window as unknown as ProbeWindow;
    const post = Worker.prototype.postMessage;
    const held: (() => void)[] = [];
    let released = false;
    scope.__commandInitHeld = false;
    scope.__releaseCommandInit = () => { released = true; for (const send of held.splice(0)) send(); };
    Worker.prototype.postMessage = function (this: Worker, message: unknown, ...rest: unknown[]) {
      if (!released && message && typeof message === 'object' && (message as { type?: unknown }).type === 'commandInit') {
        held.push(() => Reflect.apply(post, this, [message, ...rest]));
        scope.__commandInitHeld = true;
        return;
      }
      return Reflect.apply(post, this, [message, ...rest]);
    } as typeof Worker.prototype.postMessage;
  });
}

async function holdWholeSource(page: Page) {
  await page.addInitScript(() => {
    const scope = window as unknown as ProbeWindow;
    const post = Worker.prototype.postMessage;
    const held: (() => void)[] = [];
    let released = false;
    scope.__wholeSourceHeld = false;
    scope.__releaseWholeSource = () => { released = true; for (const send of held.splice(0)) send(); };
    Worker.prototype.postMessage = function (this: Worker, message: unknown, ...rest: unknown[]) {
      const request = message as { type?: unknown; source?: { cut?: unknown } } | null;
      if (!released && request && typeof request === 'object' && request.type === 'openSource' && !request.source?.cut) {
        held.push(() => Reflect.apply(post, this, [message, ...rest]));
        scope.__wholeSourceHeld = true;
        return;
      }
      return Reflect.apply(post, this, [message, ...rest]);
    } as typeof Worker.prototype.postMessage;
  });
}

async function holdReviewReads(page: Page) {
  await page.addInitScript(() => {
    const scope = window as unknown as ProbeWindow;
    const post = Worker.prototype.postMessage;
    const held: (() => void)[] = [];
    scope.__holdReviews = false;
    scope.__releaseReviews = () => { scope.__holdReviews = false; for (const send of held.splice(0)) send(); };
    Worker.prototype.postMessage = function (this: Worker, message: unknown, ...rest: unknown[]) {
      const request = message as { type?: unknown; read?: { kind?: unknown } } | null;
      if (scope.__holdReviews && request && typeof request === 'object' && request.type === 'documentRead' && request.read?.kind === 'reviewSnapshot') {
        held.push(() => Reflect.apply(post, this, [message, ...rest]));
        return;
      }
      return Reflect.apply(post, this, [message, ...rest]);
    } as typeof Worker.prototype.postMessage;
  });
}

async function recordPosts(page: Page) {
  await page.addInitScript(() => {
    const scope = window as unknown as ProbeWindow;
    const post = Worker.prototype.postMessage;
    scope.__posted = {};
    Worker.prototype.postMessage = function (this: Worker, message: unknown, ...rest: unknown[]) {
      const type = message && typeof message === 'object' ? (message as { type?: unknown }).type : undefined;
      if (typeof type === 'string') scope.__posted[type] ??= performance.now();
      return Reflect.apply(post, this, [message, ...rest]);
    } as typeof Worker.prototype.postMessage;
  });
}

async function pixels(page: Page, shots: Buffer[]) {
  return page.evaluate(async (data) => Promise.all(data.map(async (shot) => {
    const image = new Image();
    image.src = `data:image/png;base64,${shot}`;
    await image.decode();
    const canvas = document.createElement('canvas');
    canvas.width = image.width;
    canvas.height = image.height;
    const context = canvas.getContext('2d')!;
    context.fillStyle = 'white';
    context.fillRect(0, 0, image.width, image.height);
    context.drawImage(image, 0, 0);
    return { width: image.width, height: image.height, data: Array.from(context.getImageData(0, 0, image.width, image.height).data) };
  })), shots.map(shot => shot.toString('base64')));
}

async function previewShot(page: Page) {
  const preview = page.locator('[aria-label="Document preview"] canvas').first();
  await expect(preview).toBeVisible();
  const shot = await preview.screenshot();
  const [image] = await pixels(page, [shot]);
  let dark = 0;
  for (let index = 0; index < image!.data.length; index += 4) if (image!.data[index]! < 128) dark += 1;
  return { shot, dark };
}

async function differingPixels(page: Page, first: Buffer, second: Buffer) {
  const [a, b] = await pixels(page, [first, second]);
  if (a!.width !== b!.width || a!.height !== b!.height) return { size: [a!.width, a!.height, b!.width, b!.height], count: -1 };
  let count = 0;
  for (let index = 0; index < a!.data.length; index += 4) {
    const delta = Math.abs(a!.data[index]! - b!.data[index]!) + Math.abs(a!.data[index + 1]! - b!.data[index + 1]!) + Math.abs(a!.data[index + 2]! - b!.data[index + 2]!);
    if (delta > 96) count += 1;
  }
  return { size: [a!.width, a!.height], count };
}

async function layoutCompletes(page: Page) {
  await expect.poll(async () => {
    const current = await settled(page);
    return current?.timings.layoutComplete ?? (current?.errors.length ? current.errors : null);
  }, { timeout: 170_000 }).not.toBeNull();
  expect((await read(page)).errors).toEqual([]);
}

async function firstPageBeforeLayout(page: Page, parity: boolean) {
  await expect.poll(() => page.evaluate(() => (window as unknown as ProbeWindow).__commandInitHeld).catch(() => false), { timeout: 120_000 }).toBe(true);
  await expect.poll(async () => (await settled(page))?.timings.firstPagePainted ?? null, { timeout: 60_000 }).not.toBeNull();
  const preview = await previewShot(page);
  expect(preview.dark).toBeGreaterThan(100);
  const held = await read(page);
  expect(held.timings.layoutComplete).toBeNull();
  expect(held.errors).toEqual([]);
  await page.evaluate(() => (window as unknown as ProbeWindow).__releaseCommandInit());
  await layoutCompletes(page);
  await expect.poll(async () => (await settled(page))?.timings.firstInk ?? null).not.toBeNull();
  await expect(page.locator('[aria-label="Document preview"]')).toHaveCount(0);
  if (parity) {
    const final = await page.evaluate(() => document.querySelector<HTMLCanvasElement>('[data-docx-worker-editor] [data-page-index="0"] canvas')!.toDataURL());
    expect(await differingPixels(page, preview.shot, Buffer.from(final.split(',')[1]!, 'base64'))).toEqual({ size: [expect.any(Number), expect.any(Number)], count: 0 });
  }
  return held;
}

async function marks(page: Page) {
  const marks = page.locator('[data-proposal-id]');
  await expect(marks).toHaveCount(10);
  await expect.poll(() => marks.evaluateAll(elements => elements.every(element => element.getAttribute('data-ok') === 'true'))).toBe(true);
  await expect.poll(() => page.getByTestId('host-proposals').evaluate(element =>
    element.getAttribute('data-preview') === element.getAttribute('data-layout-preview'))).toBe(true);
  return marks;
}

test('a read-only viewer previews its first page, then applies ten host proposals made at open, repaints decisions and navigates to them', async ({ page }) => {
  test.setTimeout(300_000);
  await holdCommandInit(page);
  await page.goto('/docx-viewer-host-proposals.html?pages=120&pageTotal=1');
  const held = await firstPageBeforeLayout(page, false);
  expect(held.proposed).toBeNull();
  await expect.poll(async () => {
    const current = await settled(page);
    return current?.proposed ?? (current?.errors.length ? current.errors : null);
  }, { timeout: 170_000 }).not.toBeNull();
  const opened = await read(page);
  expect(opened.errors).toEqual([]);
  expect(opened.proposed!.ok).toBe(true);
  expect(opened.proposed!.snapshot!.proposals.map(({ id, state, changed }) => ({ id, state, changed })))
    .toEqual(ids.map(id => ({ id, state: 'proposed', changed: true })));
  const placed = await marks(page);
  await expect.poll(() => placed.evaluateAll(elements => elements.every(element => element.getAttribute('data-text-aligned') === 'true'))).toBe(true);
  await expect.poll(() => firstPage(page)).not.toBeNull();
  const proposed = (await firstPage(page))!;
  const decided: string[] = [];
  for (const state of ['accepted', 'rejected'] as const) {
    expect((await page.evaluate(state => (window as unknown as ProbeWindow).__viewerHostProposalProbe.decide(['host-1'], state), state)).ok).toBe(true);
    await expect(placed.first()).toHaveAttribute('data-state', state);
    await expect.poll(() => firstPage(page)).not.toBe(proposed);
    const pixels = (await firstPage(page))!;
    expect(decided).not.toContain(pixels);
    decided.push(pixels);
    expect((await page.evaluate(() => (window as unknown as ProbeWindow).__viewerHostProposalProbe.undoDecision())).ok).toBe(true);
    await expect(placed.first()).toHaveAttribute('data-state', 'proposed');
    await expect.poll(() => firstPage(page)).toBe(proposed);
  }
  for (const index of [9, 0]) {
    expect(await page.evaluate(index => (window as unknown as ProbeWindow).__viewerHostProposalProbe.next(index), index)).toEqual({ ok: true });
    await expect(placed.nth(index)).toBeInViewport();
  }
  await marks(page);
  expect((await read(page)).errors).toEqual([]);
});

test('a read-only viewer keeps host proposals past its rendered window placed across viewport frames and steps to the first in document order', async ({ page }) => {
  test.setTimeout(240_000);
  await holdReviewReads(page);
  await page.goto('/docx-viewer-host-proposals.html?pages=40&at=6');
  await expect.poll(async () => {
    const current = await settled(page);
    return current?.proposed ?? (current?.errors.length ? current.errors : null);
  }, { timeout: 170_000 }).not.toBeNull();
  expect((await read(page)).proposed!.ok).toBe(true);
  const placed = page.locator('[data-proposal-id]');
  await expect(placed).toHaveCount(10);
  const expected = ids.map((id, index) => ({ id, failure: '', page: 6 + Math.floor(index / 4), inPage: 'true' }));
  const anchors = () => placed.evaluateAll(elements => elements.map(element => ({ id: element.getAttribute('data-proposal-id'),
    failure: element.getAttribute('data-failure'), page: Number(element.getAttribute('data-page')), inPage: element.getAttribute('data-in-page') })));
  await expect.poll(anchors).toEqual(expected);
  await expect(placed.first()).not.toBeInViewport();
  await page.evaluate(() => {
    (window as unknown as ProbeWindow).__holdReviews = true;
    const scroller = document.querySelector<HTMLElement>('.docx-editor__scroll-container')!;
    scroller.scrollTop = scroller.scrollHeight / 4;
  });
  await expect.poll(() => page.evaluate(() => (window as unknown as ProbeWindow).__viewerHostProposalProbe.pageInked(10))).toBe(true);
  await expect.poll(anchors).toEqual(expected);
  await page.evaluate(() => { document.querySelector<HTMLElement>('.docx-editor__scroll-container')!.scrollTop = 0; });
  await expect.poll(anchors).toEqual(expected);
  expect(await page.evaluate(() => (window as unknown as ProbeWindow).__viewerHostProposalProbe.jumpFirst()))
    .toEqual({ id: 'host-1', pageIndex: 6, result: { ok: true } });
  await expect(placed.first()).toBeInViewport();
  await expect.poll(() => page.evaluate(() => (window as unknown as ProbeWindow).__viewerHostProposalProbe.pageInked(6))).toBe(true);
  await page.evaluate(() => (window as unknown as ProbeWindow).__releaseReviews());
  await expect.poll(() => placed.evaluateAll(elements => elements.slice(0, 4).map(element => element.getAttribute('data-text-aligned'))))
    .toEqual(['true', 'true', 'true', 'true']);
  expect((await read(page)).errors).toEqual([]);
});

for (const pageTotal of [0, 1]) {
  test(`a read-only viewer resolves paragraph, search, range, header and footer anchors after its first paint${pageTotal ? ' with page totals' : ''}`, async ({ page }) => {
    test.setTimeout(240_000);
    await page.goto(`/docx-viewer-host-proposals.html?pages=120&propose=0&headerFooter=1&pageTotal=${pageTotal}`);
    await expect.poll(async () => (await settled(page))?.timings.load ?? null, { timeout: 120_000 }).not.toBeNull();
    const targets = page.locator('[data-anchor-target]');
    await expect(targets).toHaveCount(5);
    await expect.poll(() => targets.evaluateAll(elements => elements.map(element => ({
      target: element.getAttribute('data-anchor-target'), failure: element.getAttribute('data-failure'),
      ok: element.getAttribute('data-ok') === 'true', rects: Number(element.getAttribute('data-rects')) > 0,
    })))).toEqual(['paragraph', 'search', 'range', 'header', 'footer'].map(target => ({ target, failure: '', ok: true, rects: true })));
    expect((await read(page)).errors).toEqual([]);
  });
}

test('previewFirstPage without a fonts prop paints the first page as page one, once, before the whole source opens', async ({ page }) => {
  test.setTimeout(240_000);
  await holdWholeSource(page);
  await page.goto('/docx-viewer-host-proposals.html?pages=120&propose=0&fonts=0');
  await expect.poll(() => page.evaluate(() => (window as unknown as ProbeWindow).__wholeSourceHeld).catch(() => false), { timeout: 120_000 }).toBe(true);
  await expect.poll(() => page.evaluate(() => (window as unknown as ProbeWindow).__viewerHostProposalProbe.pageInked(0)), { timeout: 60_000 }).toBe(true);
  expect((await read(page)).timings.firstPagePainted).not.toBeNull();
  await page.evaluate(() => (window as unknown as ProbeWindow).__releaseWholeSource());
  await layoutCompletes(page);
  await expect.poll(async () => (await settled(page))?.timings.firstInk ?? null).not.toBeNull();
  expect(await page.evaluate(() => (window as unknown as ProbeWindow).__viewerHostProposalProbe.firstPagePaintedCalls)).toBe(1);
  expect((await read(page)).errors).toEqual([]);
});

test('a read-only viewer without a fonts prop submits host proposals made at open before its background layout completes', async ({ page }) => {
  test.setTimeout(240_000);
  await recordPosts(page);
  await page.goto('/docx-viewer-host-proposals.html?pages=400&fonts=0');
  await expect.poll(async () => {
    const current = await settled(page);
    return current?.proposed && current.timings.layoutComplete !== null ? current.proposed : current?.errors.length ? current.errors : null;
  }, { timeout: 170_000 }).not.toBeNull();
  const opened = await read(page);
  expect(opened.errors).toEqual([]);
  expect(opened.proposed!.ok).toBe(true);
  const { posted, started } = await page.evaluate(() => ({ posted: (window as unknown as ProbeWindow).__posted,
    started: (window as unknown as ProbeWindow).__viewerHostProposalProbe.started }));
  expect(posted.edit).toBeGreaterThan(posted.commandInit!);
  expect(posted.edit! - started).toBeLessThan(opened.timings.layoutComplete!);
  await marks(page);
});

for (const viewing of [true, false]) {
  test(`previewFirstPage paints a long document's first page before its layout starts in ${viewing ? 'a read-only viewer' : 'an editor'}`, async ({ page }) => {
    test.setTimeout(240_000);
    await holdCommandInit(page);
    for (const pageTotal of [0, 1]) {
      await page.goto(`/docx-viewer-host-proposals.html?pages=120&propose=0&viewing=${Number(viewing)}&pageTotal=${pageTotal}`);
      await firstPageBeforeLayout(page, true);
      expect((await read(page)).errors).toEqual([]);
    }
  });
}
