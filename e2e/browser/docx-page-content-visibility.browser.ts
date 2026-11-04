import { test, expect, type Browser, type Page } from 'playwright/test';
import { readFile } from 'node:fs/promises';
import { resolve } from 'node:path';
import JSZip from 'jszip';

const root = resolve(import.meta.dirname, '../..');
const PAGES = 18;
const FORCE_VISIBLE = '.paged-editor__page, .canvas-page { content-visibility: visible !important; }';
const SETTLE = '[data-docx-caret] { visibility: hidden !important; } * { caret-color: transparent !important; }';
type Rect = { x: number; y: number; width: number; height: number };

async function longDocument(): Promise<Buffer> {
  const zip = await JSZip.loadAsync(await readFile(resolve(root, 'apps/demo/public/betteroffice-demo.docx')));
  const body = Array.from({ length: PAGES }, (_, index) => {
    const number = String(index + 1).padStart(2, '0');
    const heading = `<w:p><w:pPr>${index ? '<w:pageBreakBefore/>' : ''}</w:pPr><w:r><w:t>Visibility page ${number}</w:t></w:r></w:p>`;
    return heading + Array.from({ length: 20 }, (_, line) =>
      `<w:p><w:r><w:t xml:space="preserve">Marker${number} line ${line + 1} keeps the quick brown fox jumping over the lazy dog.</w:t></w:r></w:p>`).join('');
  }).join('');
  zip.file('word/document.xml', `<?xml version="1.0" encoding="UTF-8"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>${body}<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/></w:sectPr></w:body></w:document>`);
  return zip.generateAsync({ type: 'nodebuffer' });
}

async function open(browser: Browser, worker: boolean) {
  const context = await browser.newContext({ viewport: { width: 1440, height: 1000 } });
  const page = await context.newPage();
  await page.goto(`/?format=docx&peerless=${worker ? 1 : 0}`);
  await page.locator('input[type=file]').last().setInputFiles({
    name: 'visibility.docx',
    mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
    buffer: await longDocument(),
  });
  await expect(page.locator('.canvas-page[data-page-index="1"] canvas')).toBeAttached({ timeout: 60_000 });
  await expect(page.getByRole('status').filter({ hasText: /^\d+ of \d+$/ })).toHaveText(`1 of ${PAGES}`, { timeout: 60_000 });
  await expect(page.locator('.notice, [data-testid$="-error"]')).toHaveCount(0);
  if (worker) await expect(page.locator('[data-docx-worker-editor]')).toHaveAttribute('aria-busy', 'false');
  await page.addStyleTag({ content: SETTLE });
  return { context, page };
}

const frames = (page: Page) => page.evaluate(() => new Promise<void>((done) =>
  requestAnimationFrame(() => requestAnimationFrame(() => setTimeout(done, 300)))));

async function forceVisible(page: Page, on: boolean) {
  await page.evaluate(([css, on]) => {
    document.getElementById('force-visible')?.remove();
    if (!on) return;
    const style = document.createElement('style');
    style.id = 'force-visible';
    style.textContent = css;
    document.head.append(style);
  }, [FORCE_VISIBLE, on] as const);
  await frames(page);
}

/** Scrolls until the given page is mounted, puts its top `offset` px below the scroller's top and waits for its raster. */
async function scrollToPage(page: Page, index: number, offset = 40) {
  await page.evaluate(async ([index, offset]) => {
    const find = () => document.querySelector(`.canvas-page[data-page-index="${index}"]`);
    let scroller = document.querySelector('.canvas-page[data-page-index]')!.parentElement;
    while (scroller && !/(auto|scroll)/.test(getComputedStyle(scroller).overflowY)) scroller = scroller.parentElement;
    const element = scroller ?? document.scrollingElement!;
    const frame = () => new Promise((done) => requestAnimationFrame(() => requestAnimationFrame(() => setTimeout(done, 100))));
    for (let attempt = 0; attempt < 200 && !find(); attempt++) {
      const mounted = Array.from(document.querySelectorAll<HTMLElement>('.canvas-page[data-page-index]')).map((page) => Number(page.dataset.pageIndex));
      element.scrollTop += (Math.max(...mounted) < index ? 1 : -1) * element.clientHeight * 0.8;
      await frame();
    }
    const target = find();
    if (!target) throw new Error(`Page ${index + 1} never mounted`);
    element.scrollTop += target.getBoundingClientRect().top - element.getBoundingClientRect().top - offset;
  }, [index, offset] as const);
  await expect.poll(() => page.evaluate((index) => {
    const canvas = document.querySelector<HTMLCanvasElement>(`.canvas-page[data-page-index="${index}"] canvas`);
    return !!canvas && canvas.width > 0;
  }, index), { timeout: 30_000 }).toBe(true);
  await frames(page);
}

async function state(page: Page, worker: boolean) {
  const layout = await page.evaluate(() => {
    const pages = Array.from(document.querySelectorAll<HTMLElement>('.canvas-page[data-page-index]'));
    let scroller = pages[0]?.parentElement ?? null;
    while (scroller && !/(auto|scroll)/.test(getComputedStyle(scroller).overflowY)) scroller = scroller.parentElement;
    const round = (value: number) => Math.round(value * 100) / 100;
    return {
      scrollHeight: scroller?.scrollHeight ?? 0,
      scrollTop: scroller?.scrollTop ?? 0,
      pages: pages.map((element) => {
        const rect = element.getBoundingClientRect();
        return [element.dataset.pageIndex, round(rect.left), round(rect.top), round(rect.width), round(rect.height)].join(':');
      }),
    };
  });
  const root = page.locator(worker ? '[data-docx-worker-editor]' : '.editor-stage').first();
  const clip = await page.evaluate(() => {
    let scroller = document.querySelector('.canvas-page[data-page-index]')!.parentElement;
    while (scroller && !/(auto|scroll)/.test(getComputedStyle(scroller).overflowY)) scroller = scroller.parentElement;
    const rect = (scroller ?? document.documentElement).getBoundingClientRect();
    return { x: rect.x, y: rect.y, width: Math.min(rect.width, innerWidth - rect.x), height: Math.min(rect.height, innerHeight - rect.y) };
  });
  const text = await page.evaluate(() => Array.from(document.querySelectorAll<HTMLElement>('.canvas-page[data-page-index]'))
    .map((element) => `${element.dataset.pageIndex}:${element.textContent ?? ''}`));
  return { ...layout, text, aria: await root.ariaSnapshot(), pixels: await settledShot(page, clip) };
}

/** A screenshot of the document area once two consecutive captures agree. */
async function settledShot(page: Page, clip: Rect): Promise<string> {
  let previous = '';
  for (let attempt = 0; attempt < 20; attempt++) {
    const shot = (await page.screenshot({ clip })).toString('base64');
    if (shot === previous) return shot;
    previous = shot;
    await page.waitForTimeout(250);
  }
  throw new Error('Document area never settled');
}

const selectionRects = (page: Page) => page.evaluate(() => {
  const host = document.querySelector('[data-docx-pages-host]')!.getBoundingClientRect();
  return Array.from(document.querySelectorAll<HTMLElement>('[data-docx-pages-host] div[aria-hidden="true"]'))
    .filter((element) => /59, 130, 246|255, 150, 0|255, 220, 0/.test(element.style.background))
    .map((element) => {
      const rect = element.getBoundingClientRect();
      return { x: Math.round(rect.x - host.x), y: Math.round(rect.y - host.y), width: Math.round(rect.width), height: Math.round(rect.height) };
    });
});

async function settledSelection(page: Page): Promise<Rect[]> {
  let previous = '';
  let rects: Rect[] = [];
  await expect.poll(async () => {
    rects = await selectionRects(page);
    const key = JSON.stringify(rects);
    const stable = rects.length > 0 && key === previous;
    previous = key;
    return stable;
  }, { timeout: 15_000, intervals: [300] }).toBe(true);
  return rects;
}

async function canvasBox(page: Page, index: number) {
  return (await page.locator(`[data-docx-worker-editor] .canvas-page[data-page-index="${index}"] canvas`).boundingBox())!;
}

/** Double click, triple click, a drag across a page boundary and a find on pages that started outside the viewport. */
async function interactions(page: Page) {
  const input = page.locator('[data-docx-worker-editor] .paged-editor__input');
  await scrollToPage(page, 11);
  let box = await canvasBox(page, 11);
  await page.mouse.dblclick(box.x + 200, box.y + 300);
  const word = await settledSelection(page);
  await page.mouse.click(box.x + 200, box.y + 400, { clickCount: 3 });
  const paragraph = await settledSelection(page);
  await scrollToPage(page, 12, 500);
  const upper = await canvasBox(page, 11), lower = await canvasBox(page, 12);
  await page.mouse.move(upper.x + 150, upper.y + upper.height - 260);
  await page.mouse.down();
  await page.mouse.move(lower.x + 300, lower.y + 220, { steps: 12 });
  await page.mouse.up();
  const drag = await settledSelection(page);
  box = await canvasBox(page, 12);
  await page.mouse.click(box.x + 150, box.y + 150);
  await input.press('ControlOrMeta+f');
  const dialog = page.getByTestId('find-replace-dialog');
  await dialog.getByRole('textbox', { name: 'Find text', exact: true }).fill('Marker16 line 3 ');
  await expect(dialog.locator('.docx-find-replace-dialog-status')).toHaveText('1 of 1 matches');
  const found = await settledSelection(page);
  await dialog.getByRole('button', { name: 'Close dialog', exact: true }).click();
  await scrollToPage(page, 0);
  await page.mouse.click((await canvasBox(page, 0)).x + 150, (await canvasBox(page, 0)).y + 150);
  return { word, paragraph, drag, found };
}

for (const worker of [true, false]) {
  test(`docx ${worker ? 'worker' : 'in-thread'} editor: skipped pages keep layout, accessibility tree and pixels`, async ({ browser }) => {
    test.setTimeout(180_000);
    const { context, page } = await open(browser, worker);
    try {
      if (worker) {
        expect(await page.locator('.paged-editor__page').evaluateAll((pages) =>
          pages.every((element) => getComputedStyle(element).contentVisibility === 'auto'))).toBe(true);
      }
      for (const index of [0, 8, PAGES - 1]) {
        await scrollToPage(page, index);
        const skipped = await state(page, worker);
        await forceVisible(page, true);
        const shown = await state(page, worker);
        await forceVisible(page, false);
        expect(skipped.scrollHeight).toBe(shown.scrollHeight);
        expect(skipped.scrollTop).toBe(shown.scrollTop);
        expect(skipped.pages).toEqual(shown.pages);
        expect(skipped.aria).toBe(shown.aria);
        expect(skipped.text).toEqual(shown.text);
        expect(skipped.text.join(' '), 'page text in the document tree').toContain(`Marker${String(index + 1).padStart(2, '0')} line 1 `);
        expect(skipped.pixels === shown.pixels, `pixels at page ${index + 1}`).toBe(true);
      }
    } finally {
      await context.close();
    }
  });

  test(`docx ${worker ? 'worker' : 'in-thread'} editor: scrolling never jumps as pages enter the viewport`, async ({ browser }) => {
    test.setTimeout(180_000);
    const { context, page } = await open(browser, worker);
    try {
      const steps = await page.evaluate(async () => {
        const first = document.querySelector<HTMLElement>('.canvas-page[data-page-index="0"]')!;
        let scroller = first.parentElement;
        while (scroller && !/(auto|scroll)/.test(getComputedStyle(scroller).overflowY)) scroller = scroller.parentElement;
        const element = scroller!;
        const anchor = () => first.getBoundingClientRect().top - element.getBoundingClientRect().top + element.scrollTop;
        const frame = () => new Promise((done) => requestAnimationFrame(() => requestAnimationFrame(done)));
        const trace: { scrollTop: number; anchor: number; scrollHeight: number }[] = [];
        for (let step = 0; step < 60; step++) {
          element.scrollTop += 300;
          await frame();
          trace.push({ scrollTop: element.scrollTop, anchor: anchor(), scrollHeight: element.scrollHeight });
        }
        return trace;
      });
      const clamp = steps[0]!.scrollHeight;
      for (const [index, step] of steps.entries()) {
        expect(Math.abs(step.anchor - steps[0]!.anchor), `step ${index}`).toBeLessThan(0.01);
        expect(step.scrollHeight, `step ${index}`).toBe(clamp);
        if (index > 0 && steps[index - 1]!.scrollTop < step.scrollHeight - 1100) {
          expect(Math.abs(step.scrollTop - steps[index - 1]!.scrollTop - 300), `step ${index}`).toBeLessThan(1);
        }
      }
    } finally {
      await context.close();
    }
  });
}

test('docx worker editor: selection and find on skipped pages match fully painted pages', async ({ browser }) => {
  test.setTimeout(240_000);
  const { context, page } = await open(browser, true);
  try {
    const skipped = await interactions(page);
    await forceVisible(page, true);
    const shown = await interactions(page);
    for (const key of ['word', 'paragraph', 'drag', 'found'] as const) {
      expect(skipped[key].length, key).toBeGreaterThan(0);
      expect(skipped[key], key).toEqual(shown[key]);
    }
    expect(new Set(skipped.drag.map((rect) => Math.floor(rect.y / 1000))).size, 'drag spans two pages').toBeGreaterThan(1);
  } finally {
    await context.close();
  }
});
