import { expect, test, type Page } from 'playwright/test';

interface ViewerSidebarsProbe {
  editor: { commands: { execute(command: string, payload: null): Promise<unknown> } } | null;
  sessions: unknown[];
  errors: string[];
  sidebarOpen: boolean;
  replica(): { started: boolean; loaded: boolean };
  sessionReads(): Record<string, number>;
}

interface ViewerWindow {
  __viewerSidebarsProbe: ViewerSidebarsProbe;
  __wasmInstantiations: number;
}

async function instrument(page: Page) {
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

async function status(page: Page) {
  return page.evaluate(() => {
    const target = window as unknown as ViewerWindow;
    return {
      wasm: target.__wasmInstantiations,
      reads: target.__viewerSidebarsProbe.sessionReads(),
      replica: target.__viewerSidebarsProbe.replica(),
      errors: target.__viewerSidebarsProbe.errors,
    };
  });
}

async function open(page: Page) {
  await instrument(page);
  await page.goto('/docx-viewer-sidebars.html');
  await expect(page.locator('canvas[data-page-index="0"]')).toBeVisible({ timeout: 120_000 });
  const firstCanvasAt = Date.now();
  await expect.poll(() => page.evaluate(() => (window as unknown as ViewerWindow).__viewerSidebarsProbe.sessions.length)).toBeGreaterThan(0);
  return { before: await status(page), firstCanvasAt };
}

async function openSidebar(page: Page, firstCanvasAt: number) {
  const autoOpen = await page.evaluate(() => (window as unknown as ViewerWindow).__viewerSidebarsProbe.sidebarOpen);
  const started = autoOpen ? firstCanvasAt : Date.now();
  if (!autoOpen) {
    expect(await page.evaluate(() => (window as unknown as ViewerWindow).__viewerSidebarsProbe.editor!.commands.execute('commentsSidebar', null))).toMatchObject({ ok: true });
  }
  await expect(page.locator('.docx-comment-card')).toBeVisible();
  console.log(JSON.stringify({ sidebarFirstCardMs: Date.now() - started }));
}

async function expectUnchanged(page: Page, before: Awaited<ReturnType<typeof status>>) {
  const reads = Object.fromEntries(Object.keys(before.reads).map((method) => [method, 0]));
  // The open path checks the first frame's caret against the main selection once.
  reads.selection = before.reads.selection ?? 0;
  expect(before.reads.selection ?? 0).toBeLessThanOrEqual(1);
  expect(await status(page)).toEqual({ ...before, reads, replica: { started: false, loaded: false }, errors: [] });
}

test('viewer comment and tracked-change cards are placed without a document replica', async ({ page }) => {
  const { before, firstCanvasAt } = await open(page);
  await openSidebar(page, firstCanvasAt);
  const comment = page.locator('.docx-comment-card');
  const tracked = page.locator('.docx-tracked-change-card');
  await expect(comment).toContainText('Check this page one comment.');
  await expect(tracked).toContainText('Document writer');
  await expect(tracked).toBeVisible();
  const scroll = await page.locator('.docx-editor__scroll-container').boundingBox();
  expect(scroll).not.toBeNull();
  for (const card of [comment, tracked]) {
    await expect.poll(async () => {
      const box = await card.boundingBox();
      return !!box && box.y > scroll!.y && box.y < scroll!.y + scroll!.height;
    }).toBe(true);
  }
  const tops = await Promise.all([comment.boundingBox(), tracked.boundingBox()]);
  expect(tops[0]!.y).not.toBe(tops[1]!.y);
  await expectUnchanged(page, before);
});

test('viewer outline headings navigate to page three through the worker', async ({ page }) => {
  const { before } = await open(page);
  await page.locator('.docx-outline-toggle').click();
  const outline = page.locator('.docx-outline-nav');
  for (const text of ['First heading', 'Second heading', 'Page three heading']) {
    await expect(outline.getByText(text, { exact: true })).toBeVisible();
  }
  await outline.getByText('Page three heading', { exact: true }).click();
  await expect.poll(async () => {
    const canvas = await page.locator('canvas[data-page-index="2"]').boundingBox();
    const scroll = await page.locator('.docx-editor__scroll-container').boundingBox();
    return !!canvas && !!scroll && canvas.y >= scroll.y && canvas.y < scroll.y + scroll.height;
  }).toBe(true);
  await expectUnchanged(page, before);
});

test('expanding a viewer comment highlights it without session reads', async ({ page }) => {
  const { before, firstCanvasAt } = await open(page);
  await openSidebar(page, firstCanvasAt);
  await page.locator('.docx-comment-card').click();
  await expect(page.locator('.docx-canvas-brighten-comment').first()).toBeVisible();
  await expectUnchanged(page, before);
});
