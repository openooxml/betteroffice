import { expect, test, type Page } from 'playwright/test';

interface PointRead {
  target: {
    start: { paraId: string; offset: number };
    end: { paraId: string; offset: number };
  };
}

interface PersistedParagraph {
  kind: 'persisted';
  story: string;
  paraId: string;
}

interface ViewerEditor {
  readPositionAtPoint(clientX: number, clientY: number): Promise<PointRead | null>;
  getPositionAtPoint(clientX: number, clientY: number): PointRead | null;
  getParagraphIdentities(): Promise<{
    paragraphs: Array<{ session?: { story: string }; persisted?: PersistedParagraph }>;
  }>;
  resolveParagraphAnchors(anchors: readonly PersistedParagraph[]): Promise<{ version: string }>;
  proposeChanges(request: {
    expectVersion: string;
    proposals: readonly {
      id: string;
      paragraph: PersistedParagraph;
      suggest: { author: string; date: string };
      op: 'replaceText';
      search: string;
      replaceWith: string;
    }[];
  }): Promise<{ ok: boolean }>;
}

interface ViewerWindow {
  __viewerSelectionProbe: {
    editor: ViewerEditor | null;
    copies: string[];
    errors: string[];
    paragraphText(index: number): string;
    paragraphs: number;
    replica(): { started: boolean; loaded: boolean };
  };
  __wasmInstantiations: number;
}

const SENTINEL = '⁣unchanged';

async function instrument(page: Page) {
  await page.context().grantPermissions(['clipboard-read', 'clipboard-write']);
  await page.addInitScript(() => {
    const target = window as unknown as ViewerWindow;
    target.__wasmInstantiations = 0;
    for (const name of ['instantiate', 'instantiateStreaming', 'compile', 'compileStreaming'] as const) {
      const original = WebAssembly[name] as (...args: unknown[]) => unknown;
      (WebAssembly as unknown as Record<string, unknown>)[name] = (...args: unknown[]) => {
        target.__wasmInstantiations += 1;
        return original.apply(WebAssembly, args);
      };
    }
  });
}

async function open(page: Page, preview = false) {
  await instrument(page);
  await page.goto(`/docx-viewer-selection.html?preview=${preview ? 1 : 0}`);
  await expect(page.locator('canvas[data-page-index="0"]')).toBeVisible({ timeout: 120_000 });
}

/** A client point at fractions of the first page, inside its text column. */
async function pagePoint(page: Page, fx: number, fy: number) {
  // The preview's canvas is replaced when the whole document takes over.
  let box: Awaited<ReturnType<ReturnType<Page['locator']>['boundingBox']>> = null;
  await expect.poll(async () => (box = await page.locator('canvas[data-page-index="0"]').boundingBox())).not.toBeNull();
  return { x: box!.x + fx * box!.width, y: box!.y + fy * box!.height };
}

/** The first line of the first paragraph: one inch down and in, plus half a line. */
const FIRST_LINE = { fx: 0.2, fy: 96 / 1056 + 0.008 };

async function wasmInstantiations(page: Page) {
  return page.evaluate(() => (window as unknown as ViewerWindow).__wasmInstantiations);
}

async function replica(page: Page) {
  return page.evaluate(() => (window as unknown as ViewerWindow).__viewerSelectionProbe.replica());
}

async function copied(page: Page): Promise<string> {
  await page.evaluate(async (sentinel) => {
    (window as unknown as ViewerWindow).__viewerSelectionProbe.copies.length = 0;
    await navigator.clipboard.writeText(sentinel);
  }, SENTINEL);
  await page.keyboard.press('ControlOrMeta+C');
  let text = SENTINEL;
  await expect.poll(async () => {
    text = await page.evaluate(async (sentinel) => {
      const copies = (window as unknown as ViewerWindow).__viewerSelectionProbe.copies;
      const clipboard = await navigator.clipboard.readText();
      return clipboard !== sentinel ? clipboard : (copies.at(-1) ?? sentinel);
    }, SENTINEL);
    return text !== SENTINEL;
  }).toBe(true);
  return text;
}

async function copiedNow(page: Page): Promise<string> {
  await page.evaluate(async (sentinel) => {
    (window as unknown as ViewerWindow).__viewerSelectionProbe.copies.length = 0;
    await navigator.clipboard.writeText(sentinel);
  }, SENTINEL);
  await page.keyboard.press('ControlOrMeta+C');
  await page.waitForTimeout(200);
  return page.evaluate(async (sentinel) => {
    const copies = (window as unknown as ViewerWindow).__viewerSelectionProbe.copies;
    const clipboard = await navigator.clipboard.readText();
    return clipboard !== sentinel ? clipboard : (copies.filter((text) => text !== '').at(-1) ?? sentinel);
  }, SENTINEL);
}

async function paragraph(page: Page, index: number) {
  return page.evaluate(
    (n) => (window as unknown as ViewerWindow).__viewerSelectionProbe.paragraphText(n),
    index
  );
}

async function pointAt(page: Page, x: number, y: number) {
  return page.evaluate(
    ([px, py]) => (window as unknown as ViewerWindow).__viewerSelectionProbe.editor!.readPositionAtPoint(px!, py!),
    [x, y]
  );
}

function wordAround(text: string, offset: number): string {
  const word = (char: string | undefined) => char !== undefined && /\w/.test(char);
  const at = Math.min(offset, text.length - 1);
  const inWord = word(text[at]);
  let start = at;
  let end = at;
  const same = (char: string | undefined) => (inWord ? word(char) : char === ' ');
  while (start > 0 && same(text[start - 1])) start -= 1;
  while (end < text.length && same(text[end])) end += 1;
  return text.slice(start, end);
}

async function expectNoMainThreadDocument(page: Page, instantiations: number) {
  expect(await replica(page)).toEqual({ started: false, loaded: false });
  expect(await wasmInstantiations(page)).toBe(instantiations);
}

for (const preview of [false, true]) {
  test(`a triple-click right after the first paint selects its paragraph without a main-thread document${preview ? ' (first-page preview)' : ''}`, async ({ page }) => {
    await open(page, preview);
    const instantiations = await wasmInstantiations(page);
    const { x, y } = await pagePoint(page, FIRST_LINE.fx, FIRST_LINE.fy);
    await page.mouse.click(x, y, { clickCount: 3 });
    await expect(page.locator('[data-testid^="selection-rect-"]').first()).toBeVisible();
    expect(await copied(page)).toBe(await paragraph(page, 1));
    await expectNoMainThreadDocument(page, instantiations);
  });
}

test('readPositionAtPoint answers from the worker right after the first paint, and getPositionAtPoint on retry', async ({ page }) => {
  await open(page);
  const instantiations = await wasmInstantiations(page);
  const { x, y } = await pagePoint(page, FIRST_LINE.fx, FIRST_LINE.fy);
  let read = await pointAt(page, x, y);
  await expect.poll(async () => (read = await pointAt(page, x, y)) !== null).toBe(true);
  expect(read!.target.start.paraId).toBe('00000001');
  expect(read!.target.start).toEqual(read!.target.end);
  await expect.poll(() => page.evaluate(
    ([px, py]) => (window as unknown as ViewerWindow).__viewerSelectionProbe.editor!.getPositionAtPoint(px!, py!),
    [x, y]
  )).toEqual(read);
  await expectNoMainThreadDocument(page, instantiations);
});

test('a double-click selects the word, also when the pointer moves a little between the clicks', async ({ page }) => {
  await open(page);
  const instantiations = await wasmInstantiations(page);
  const { x, y } = await pagePoint(page, 0.35, FIRST_LINE.fy);
  let read = await pointAt(page, x, y);
  await expect.poll(async () => (read = await pointAt(page, x, y)) !== null).toBe(true);
  const word = wordAround(await paragraph(page, 1), read!.target.start.offset);
  await page.mouse.click(x, y, { clickCount: 2 });
  expect(await copied(page)).toBe(word);
  await page.mouse.move(x, y);
  await page.mouse.down({ clickCount: 1 });
  await page.mouse.up({ clickCount: 1 });
  await page.mouse.down({ clickCount: 2 });
  await page.mouse.move(x + 2, y + 1);
  await page.mouse.up({ clickCount: 2 });
  expect(await copied(page)).toBe(word);
  await expectNoMainThreadDocument(page, instantiations);
});

test('a drag across lines released at once keeps the whole range', async ({ page }) => {
  await open(page);
  const instantiations = await wasmInstantiations(page);
  const from = await pagePoint(page, 0.25, FIRST_LINE.fy);
  const to = await pagePoint(page, 0.6, FIRST_LINE.fy + 0.05);
  let start = await pointAt(page, from.x, from.y);
  await expect.poll(async () => (start = await pointAt(page, from.x, from.y)) !== null).toBe(true);
  const end = await pointAt(page, to.x, to.y);
  expect(end!.target.start.paraId).toBe('00000001');
  expect(end!.target.start.offset).toBeGreaterThan(start!.target.start.offset + 100);
  await page.mouse.move(from.x, from.y);
  await page.mouse.down();
  await page.mouse.move(to.x, to.y, { steps: 2 });
  await page.mouse.up();
  const text = await paragraph(page, 1);
  const expected = text.slice(start!.target.start.offset, end!.target.start.offset);
  const received = await copied(page);
  // Point reads and the pointer round a hit to the nearest character edge on their own.
  expect(Math.abs(received.length - expected.length)).toBeLessThanOrEqual(1);
  expect(expected.startsWith(received) || received.startsWith(expected)).toBe(true);
  expect(received.length).toBeGreaterThan(100);
  await expectNoMainThreadDocument(page, instantiations);
});

test('select-all copies the whole body', async ({ page }) => {
  await open(page);
  const instantiations = await wasmInstantiations(page);
  const { x, y } = await pagePoint(page, FIRST_LINE.fx, FIRST_LINE.fy);
  await page.mouse.click(x, y);
  await page.keyboard.press('ControlOrMeta+A');
  const count = await page.evaluate(() => (window as unknown as ViewerWindow).__viewerSelectionProbe.paragraphs);
  const texts = await Promise.all(Array.from({ length: count }, (_, index) => paragraph(page, index + 1)));
  await expect.poll(() => copied(page)).toBe(texts.join('\n'));
  await expectNoMainThreadDocument(page, instantiations);
});

test('a proposal clears the selection, and a triple-click right after it selects', async ({ page }) => {
  await open(page);
  const instantiations = await wasmInstantiations(page);
  const { x, y } = await pagePoint(page, 0.35, FIRST_LINE.fy);
  let read = await pointAt(page, x, y);
  await expect.poll(async () => (read = await pointAt(page, x, y)) !== null).toBe(true);
  await page.mouse.click(x, y, { clickCount: 2 });
  const word = await copied(page);
  expect(word).toBe(wordAround(await paragraph(page, 1), read!.target.start.offset));
  await page.evaluate(async () => {
    const editor = (window as unknown as ViewerWindow).__viewerSelectionProbe.editor!;
    const identities = await editor.getParagraphIdentities();
    const first = identities.paragraphs.find((entry) => entry.session?.story === 'body')!;
    const resolved = await editor.resolveParagraphAnchors([first.persisted!]);
    const result = await editor.proposeChanges({
      expectVersion: resolved.version,
      proposals: [{
        id: 'p1',
        paragraph: first.persisted!,
        suggest: { author: 'Reviewer', date: '2026-10-02T00:00:00Z' },
        op: 'replaceText',
        search: 'P1 ',
        replaceWith: 'P1 inserted words ahead ',
      }],
    });
    if (!result.ok) throw new Error('the proposal failed');
  });
  await expect.poll(() => copiedNow(page)).toBe(SENTINEL);
  await page.mouse.click(x, y, { clickCount: 3 });
  await expect.poll(async () => (await copied(page)).startsWith('P1 ')).toBe(true);
  await expectNoMainThreadDocument(page, instantiations);
});
