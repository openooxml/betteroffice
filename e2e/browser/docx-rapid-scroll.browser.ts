import { test, expect, type Page } from 'playwright/test';
import { readFile } from 'node:fs/promises';
import { resolve } from 'node:path';
import JSZip from 'jszip';

const root = resolve(import.meta.dirname, '../..');
const pageCounts = (process.env.RAPID_SCROLL_PAGES ?? '300').split(',').map(Number);
const sweepMs = Number(process.env.RAPID_SCROLL_MS ?? 10_000);

async function longDocument(pages: number) {
  const zip = await JSZip.loadAsync(await readFile(resolve(root, 'apps/demo/public/betteroffice-demo.docx')));
  const body = Array.from({ length: pages }, (_, index) =>
    `<w:p><w:pPr>${index ? '<w:pageBreakBefore/>' : ''}</w:pPr><w:r><w:t>Section ${index + 1}</w:t></w:r></w:p>` +
    Array.from({ length: 6 }, (_, line) =>
      `<w:p><w:r><w:t>Section ${index + 1} line ${line + 1}. The quick brown fox jumps over the lazy dog.</w:t></w:r></w:p>`).join('')
  ).join('');
  zip.file('word/document.xml', '<?xml version="1.0" encoding="UTF-8"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>' +
    body + '<w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/></w:sectPr></w:body></w:document>');
  return zip.generateAsync({ type: 'nodebuffer' });
}

const scrollRange = (page: Page) => page.evaluate(() => {
  const scroller = document.querySelector<HTMLElement>('[data-docx-worker-editor]')?.closest<HTMLElement>('.docx-editor__scroll-container');
  return scroller ? scroller.scrollHeight - scroller.clientHeight : 0;
});

async function openLaidOut(page: Page, pages: number) {
  await page.goto('/?format=docx');
  await page.locator('input[type=file]').last().setInputFiles({
    name: 'rapid-scroll.docx',
    mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
    buffer: await longDocument(pages),
  });
  await expect(page.locator('[data-docx-worker-editor] [data-page-id]').first()).toBeVisible({ timeout: 60_000 });
  let previous = -1;
  await expect.poll(async () => {
    const range = await scrollRange(page);
    const settled = range === previous && range > pages * 900;
    previous = range;
    return settled;
  }, { timeout: 180_000, intervals: [1000] }).toBe(true);
}

for (const pages of pageCounts) {
  test(`a rapid scroll through ${pages} laid out pages keeps the document mounted`, async ({ page }) => {
    test.setTimeout(300_000);
    const errors: string[] = [];
    page.on('pageerror', (error) => errors.push(error.message));
    page.on('console', (message) => { if (message.type() === 'error') errors.push(message.text()); });
    await openLaidOut(page, pages);
    await page.evaluate(async (durationMs) => {
      const scroller = document.querySelector<HTMLElement>('[data-docx-worker-editor]')!.closest<HTMLElement>('.docx-editor__scroll-container')!;
      const start = performance.now();
      await new Promise<void>((done) => {
        const frame = () => {
          const fraction = Math.min(1, (performance.now() - start) / durationMs);
          scroller.scrollTop = fraction * (scroller.scrollHeight - scroller.clientHeight);
          if (fraction < 1) requestAnimationFrame(frame); else done();
        };
        requestAnimationFrame(frame);
      });
    }, sweepMs);
    await page.waitForTimeout(1000);
    expect(errors.filter((error) => /Maximum update depth|error #185/.test(error))).toEqual([]);
    await expect(page.locator('.docx-error-fallback')).toHaveCount(0);
    expect(await scrollRange(page)).toBeGreaterThan(pages * 900);
    await expect(page.locator(`[data-docx-worker-editor] [data-page-id][data-page-index="${pages - 1}"]`)).toBeVisible();
  });
}
