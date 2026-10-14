import { test, expect, type Page } from 'playwright/test';
import { readFile } from 'node:fs/promises';
import { resolve } from 'node:path';
import JSZip from 'jszip';

const root = resolve(import.meta.dirname, '../..');
const PARAGRAPHS = 800;

/** The demo package with a body of dense text that fills every page to the bottom margin. */
async function longDocument(): Promise<Buffer> {
  const zip = await JSZip.loadAsync(
    await readFile(resolve(root, 'apps/demo/public/betteroffice-demo.docx'))
  );
  const xml = await zip.file('word/document.xml')!.async('string');
  const line = 'The quick brown fox jumps over the lazy dog, again and again. '.repeat(6);
  const paragraphs = Array.from(
    { length: PARAGRAPHS },
    () => `<w:p><w:r><w:t xml:space="preserve">${line}</w:t></w:r></w:p>`
  ).join('');
  const body = xml.replace(
    /<w:body>[\s\S]*?(<w:sectPr[\s\S]*<\/w:sectPr>)?\s*<\/w:body>/,
    (_, section) => `<w:body>${paragraphs}${section ?? ''}</w:body>`
  );
  zip.file('word/document.xml', body);
  return zip.generateAsync({ type: 'nodebuffer' });
}

async function open(page: Page) {
  await page.goto('/?format=docx');
  await page
    .locator('input[type=file]')
    .last()
    .setInputFiles({
      name: 'long.docx',
      mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
      buffer: await longDocument(),
    });
  // More pages than the canvas keeps bitmaps for at once.
  await expect
    .poll(() => page.locator('.canvas-page canvas').count(), { timeout: 60_000 })
    .toBeGreaterThan(12);
}

/**
 * Dark pixels in a screenshot of the middle of the page scroller: a band taller
 * than the margins and gap between one page's text and the next.
 */
async function inkAtViewportCentre(page: Page): Promise<number> {
  const box = await page.evaluate(() => {
    let scroller = document.querySelector('.canvas-pages')?.parentElement ?? null;
    while (scroller && !/(auto|scroll)/.test(getComputedStyle(scroller).overflowY)) {
      scroller = scroller.parentElement;
    }
    const rect = (scroller ?? document.documentElement).getBoundingClientRect();
    return { x: rect.left + rect.width / 2 - 150, y: rect.top + rect.height / 2 - 200 };
  });
  const png = await page.screenshot({ clip: { ...box, width: 300, height: 400 } });
  return page.evaluate(async (base64) => {
    const bytes = Uint8Array.from(atob(base64), (char) => char.charCodeAt(0));
    const image = await createImageBitmap(new Blob([bytes], { type: 'image/png' }));
    const canvas = new OffscreenCanvas(image.width, image.height);
    const context = canvas.getContext('2d')!;
    context.drawImage(image, 0, 0);
    const pixels = context.getImageData(0, 0, image.width, image.height).data;
    let dark = 0;
    for (let index = 0; index < pixels.length; index += 4) {
      if (pixels[index]! + pixels[index + 1]! + pixels[index + 2]! < 300) dark += 1;
    }
    return dark;
  }, png.toString('base64'));
}

for (const zoom of [0.8, 1.25]) {
  test(`docx: pages stay painted while scrolling under an ancestor CSS zoom of ${zoom}`, async ({
    page,
  }) => {
    test.setTimeout(240_000);
    await open(page);
    await page.addStyleTag({ content: `.editor-stage { zoom: ${zoom}; }` });
    for (const fraction of [0.3, 0.7]) {
      await page.evaluate((fraction) => {
        let scroller = document.querySelector('.canvas-pages')?.parentElement ?? null;
        while (scroller && !/(auto|scroll)/.test(getComputedStyle(scroller).overflowY)) {
          scroller = scroller.parentElement;
        }
        const target = scroller ?? document.scrollingElement!;
        target.scrollTop = (target.scrollHeight - target.clientHeight) * fraction;
      }, fraction);
      await expect.poll(() => inkAtViewportCentre(page), { timeout: 45_000 }).toBeGreaterThan(50);
    }
  });
}
