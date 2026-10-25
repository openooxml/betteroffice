import { expect, test, type Page } from 'playwright/test';
import JSZip from 'jszip';

interface DragSample { source: 'raf' | 'mutation'; count: number; top: number; left: number; bottom: number; right: number }

declare global {
  interface Window { __docxDragSamples?: { active: boolean; samples: DragSample[]; stop(): void } }
}

test.use({ launchOptions: { executablePath: process.env.CHROME || undefined } });

const OVERLAY = '[data-docx-worker-editor] div[aria-hidden="true"][style*="rgba(59, 130, 246"]';
const PAGE_WIDTH = 816;
const MARGIN = 96;

async function documentBytes() {
  const sentence = 'The quick brown fox jumps over the lazy dog while the selection follows the pointer across lines. ';
  const zip = new JSZip();
  zip.file('[Content_Types].xml',
    '<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
    '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
    '<Default Extension="xml" ContentType="application/xml"/>' +
    '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>' +
    '</Types>');
  zip.file('_rels/.rels',
    '<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
    '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>' +
    '</Relationships>');
  zip.file('word/document.xml',
    '<?xml version="1.0"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>' +
    Array.from({ length: 6 }, () => `<w:p><w:r><w:t xml:space="preserve">${sentence.repeat(4)}</w:t></w:r></w:p>`).join('') +
    '<w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/></w:sectPr>' +
    '</w:body></w:document>');
  return zip.generateAsync({ type: 'nodebuffer' });
}

async function frame(page: Page) {
  await page.evaluate(() => new Promise<void>((resolve) => requestAnimationFrame(() => resolve())));
}

async function startSampling(page: Page) {
  await page.evaluate((selector) => {
    const samples: DragSample[] = [];
    const read = (source: DragSample['source']) => {
      if (!window.__docxDragSamples?.active) return;
      const rects = Array.from(document.querySelectorAll<HTMLElement>(selector), (element) => element.getBoundingClientRect());
      const last = rects.reduce<DOMRect | undefined>((best, rect) =>
        !best || rect.bottom > best.bottom + 0.5 || Math.abs(rect.bottom - best.bottom) <= 0.5 && rect.right > best.right ? rect : best, undefined);
      const first = rects.reduce<DOMRect | undefined>((best, rect) =>
        !best || rect.top < best.top - 0.5 || Math.abs(rect.top - best.top) <= 0.5 && rect.left < best.left ? rect : best, undefined);
      samples.push({ source, count: rects.length, top: first?.top ?? 0, left: first?.left ?? 0, bottom: last?.bottom ?? 0, right: last?.right ?? 0 });
    };
    const observer = new MutationObserver(() => read('mutation'));
    observer.observe(document.querySelector('[data-docx-worker-editor]')!, { childList: true, subtree: true, attributes: true, attributeFilter: ['style'] });
    let raf = 0;
    const tick = () => { read('raf'); raf = requestAnimationFrame(tick); };
    raf = requestAnimationFrame(tick);
    window.__docxDragSamples = { active: false, samples, stop() { observer.disconnect(); cancelAnimationFrame(raf); } };
  }, OVERLAY);
}

test('drag selection follows the pointer without flicker or snapping back', async ({ page }) => {
  await page.goto('/?format=docx&peerless=1');
  await page.locator('input[type=file]').last().setInputFiles({
    name: 'drag-selection.docx',
    mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
    buffer: await documentBytes(),
  });
  const canvas = page.locator('[data-docx-worker-editor] [data-page-id] canvas').first();
  await expect(canvas).toBeVisible({ timeout: 60_000 });
  const bounds = await canvas.boundingBox();
  if (!bounds) throw new Error('Document canvas is unavailable');
  const scale = bounds.width / PAGE_WIDTH;
  const start = { x: bounds.x + (MARGIN + 1) * scale, y: bounds.y + (MARGIN + 8) * scale };
  const end = { x: bounds.x + (PAGE_WIDTH - MARGIN - 40) * scale, y: bounds.y + (MARGIN + 260) * scale };
  await startSampling(page);
  await page.mouse.move(start.x, start.y);
  await page.mouse.down();
  await page.evaluate(() => { window.__docxDragSamples!.active = true; });
  const steps = 90;
  for (let step = 1; step <= steps; step += 1) {
    await page.mouse.move(start.x + (end.x - start.x) * step / steps, start.y + (end.y - start.y) * step / steps);
    await frame(page);
  }
  await page.evaluate(() => { window.__docxDragSamples!.active = false; });
  await page.mouse.up();
  await expect.poll(async () => page.locator(OVERLAY).count()).toBeGreaterThanOrEqual(5);
  const samples = await page.evaluate(() => { window.__docxDragSamples!.stop(); return window.__docxDragSamples!.samples; });
  const first = samples.findIndex((sample) => sample.count > 0);
  expect(first, 'the drag never showed a selection').toBeGreaterThanOrEqual(0);
  const during = samples.slice(first);
  const empty = during.filter((sample) => sample.count === 0);
  const backwards: { index: number; previous: DragSample; sample: DragSample }[] = [];
  let previous: DragSample | undefined;
  for (const [index, sample] of during.entries()) {
    if (sample.count === 0) continue;
    if (previous && (sample.bottom < previous.bottom - 0.5 ||
      Math.abs(sample.bottom - previous.bottom) <= 0.5 && sample.right < previous.right - 0.5)) backwards.push({ index, previous, sample });
    previous = sample;
  }
  console.log(`[drag-selection] samples=${during.length} raf=${during.filter((sample) => sample.source === 'raf').length} ` +
    `empty=${empty.length} emptyRaf=${empty.filter((sample) => sample.source === 'raf').length} backwards=${backwards.length} ` +
    `finalBottom=${during.at(-1)?.bottom.toFixed(1)}`);
  console.log(`[drag-selection] frames ${JSON.stringify(during.filter((sample) => sample.source === 'raf').map((sample) => [sample.count, Math.round(sample.bottom), Math.round(sample.right)]))}`);
  expect(empty, 'the highlight emptied while the pointer moved away from the anchor').toEqual([]);
  expect(backwards, 'the highlight moved back toward the anchor').toEqual([]);
  const shown = during.filter((sample) => sample.count > 0);
  expect(shown.filter((sample) => Math.abs(sample.top - shown[0]!.top) > 0.5 || Math.abs(sample.left - shown[0]!.left) > 0.5),
    'the anchor end of the highlight moved').toEqual([]);
  const extents = new Set(during.filter((sample) => sample.source === 'raf' && sample.count > 0)
    .map((sample) => `${Math.round(sample.bottom)}:${Math.round(sample.right)}`));
  expect(extents.size, 'the highlight did not follow the pointer during the drag').toBeGreaterThanOrEqual(10);
  expect(shown.at(-1)!.bottom - shown[0]!.bottom).toBeGreaterThan(100 * scale);
});
