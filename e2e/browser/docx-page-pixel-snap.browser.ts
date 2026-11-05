import { test, expect, type Browser, type Page } from 'playwright/test';
import { readFile } from 'node:fs/promises';
import { resolve } from 'node:path';
import JSZip from 'jszip';
import { pinnedEvaluate } from './pinned-evaluate';

const root = resolve(import.meta.dirname, '../..');
const A4 = { width: 11906 / 15, height: 16838 / 15 };
const ZOOMS = ['50%', '75%', '90%', '100%', '125%', '150%', '200%'];
const PROBE_PAGE = 1;
const BLOCK = 192;

async function a4Document(): Promise<Buffer> {
  const zip = await JSZip.loadAsync(await readFile(resolve(root, 'apps/demo/public/betteroffice-demo.docx')));
  const line = 'Pixel snapping probe: the quick brown fox jumps over the lazy dog, 0123456789. '.repeat(3);
  const paragraphs = Array.from({ length: 160 }, () =>
    `<w:p><w:r><w:t xml:space="preserve">${line}</w:t></w:r></w:p>`).join('');
  zip.file('word/document.xml', `<?xml version="1.0" encoding="UTF-8"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>${paragraphs}<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/></w:sectPr></w:body></w:document>`);
  return zip.generateAsync({ type: 'nodebuffer' });
}

async function open(browser: Browser, devicePixelRatio: number, worker: boolean) {
  const context = await browser.newContext({ deviceScaleFactor: devicePixelRatio, viewport: { width: 1440, height: 1000 } });
  const page = await context.newPage();
  await page.goto(`/?format=docx&peerless=${worker ? 1 : 0}`);
  await page.locator('input[type=file]').last().setInputFiles({
    name: 'a4.docx',
    mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
    buffer: await a4Document(),
  });
  await expect(page.locator(`.canvas-page[data-page-index="${PROBE_PAGE}"] canvas`)).toBeAttached({ timeout: 60_000 });
  await expect(page.locator('.notice, [data-testid$="-error"]')).toHaveCount(0);
  if (worker) await expect(page.locator('[data-docx-worker-editor]')).toHaveAttribute('aria-busy', 'false');
  await page.addStyleTag({ content: '[data-docx-caret], .paged-editor__input { visibility: hidden !important; }' });
  return { context, page };
}

async function setZoom(page: Page, value: string) {
  const control = page.getByRole('combobox', { name: /^Zoom:/ });
  await control.click();
  await page.getByRole('option', { name: value, exact: true }).click();
  await expect(control).toContainText(value);
}

/** Scrolls the probe page into view and waits until its canvas holds a settled raster at the zoom. */
async function settle(page: Page, zoom: number) {
  await page.evaluate((index) => {
    const canvas = document.querySelector(`.canvas-page[data-page-index="${index}"] canvas`)!;
    let scroller = canvas.parentElement;
    while (scroller && !/(auto|scroll)/.test(getComputedStyle(scroller).overflowY)) scroller = scroller.parentElement;
    const target = scroller ?? document.scrollingElement!;
    target.scrollTop += canvas.getBoundingClientRect().top - target.getBoundingClientRect().top - 40;
  }, PROBE_PAGE);
  let previous = '';
  await expect.poll(async () => {
    const state = await page.evaluate(({ index, width }) => {
      const canvas = document.querySelector<HTMLCanvasElement>(`.canvas-page[data-page-index="${index}"] canvas`);
      if (!canvas || Math.abs(canvas.width - width * devicePixelRatio) > 1.01) return 'raster';
      const rect = canvas.getBoundingClientRect();
      return `${canvas.width}x${canvas.height}@${rect.left},${rect.top},${rect.width},${rect.height}`;
    }, { index: PROBE_PAGE, width: A4.width * zoom });
    const stable = state !== 'raster' && state === previous;
    previous = state;
    return stable;
  }, { timeout: 30_000, intervals: [400] }).toBe(true);
  await page.waitForTimeout(300);
}

interface Probe {
  devicePixelRatio: number;
  backing: { width: number; height: number };
  box: { left: number; top: number; width: number; height: number };
  pixels: { mismatched: number; total: number; ink: number; offset: [number, number] } | { unreadable: string };
}

/** The probe page's box in device pixels, and how its on-screen pixels compare with its canvas backing. */
async function probe(page: Page): Promise<Probe> {
  const geometry = await page.evaluate((index) => {
    const canvas = document.querySelector<HTMLCanvasElement>(`.canvas-page[data-page-index="${index}"] canvas`)!;
    const rect = canvas.getBoundingClientRect();
    const ratio = devicePixelRatio;
    return { devicePixelRatio: ratio, backing: { width: canvas.width, height: canvas.height },
      box: { left: rect.left * ratio, top: rect.top * ratio, width: rect.width * ratio, height: rect.height * ratio },
      viewport: { width: innerWidth * ratio, height: innerHeight * ratio } };
  }, PROBE_PAGE);
  const shot = await page.screenshot({ scale: 'device' });
  const pixels = await pinnedEvaluate(page, async ({ base64, index, box, viewport, block }) => {
    const canvas = document.querySelector<HTMLCanvasElement>(`.canvas-page[data-page-index="${index}"] canvas`)!;
    let source: ImageData;
    try {
      source = canvas.getContext('2d')!.getImageData(0, 0, canvas.width, canvas.height);
    } catch (error) {
      return { unreadable: String(error) };
    }
    const bytes = Uint8Array.from(atob(base64), (char) => char.charCodeAt(0));
    const image = await createImageBitmap(new Blob([bytes], { type: 'image/png' }));
    const screen = new OffscreenCanvas(image.width, image.height);
    const context = screen.getContext('2d')!;
    context.drawImage(image, 0, 0);
    const shotPixels = context.getImageData(0, 0, image.width, image.height);
    const left = Math.max(0, Math.round(box.left)), top = Math.max(0, Math.round(box.top));
    const right = Math.min(viewport.width, Math.round(box.left + box.width));
    const bottom = Math.min(viewport.height, Math.round(box.top + box.height));
    const x0 = Math.round((left + right) / 2 - block / 2), y0 = Math.round(top + (bottom - top) * 0.35);
    let best: { mismatched: number; total: number; ink: number; offset: [number, number] } | null = null;
    for (const dy of [-1, 0, 1]) for (const dx of [-1, 0, 1]) {
      let mismatched = 0, total = 0, ink = 0;
      for (let y = y0; y < y0 + block; y++) for (let x = x0; x < x0 + block; x++) {
        const sx = x - Math.round(box.left) + dx, sy = y - Math.round(box.top) + dy;
        if (sx < 0 || sy < 0 || sx >= source.width || sy >= source.height || x >= shotPixels.width || y >= shotPixels.height) continue;
        const a = (y * shotPixels.width + x) * 4, b = (sy * source.width + sx) * 4;
        const alpha = source.data[b + 3]! / 255;
        let differs = false;
        for (let channel = 0; channel < 3; channel++) {
          const expected = source.data[b + channel]! * alpha + 255 * (1 - alpha);
          if (Math.abs(shotPixels.data[a + channel]! - expected) > 2) differs = true;
        }
        if (shotPixels.data[a]! + shotPixels.data[a + 1]! + shotPixels.data[a + 2]! < 384) ink += 1;
        total += 1;
        if (differs) mismatched += 1;
      }
      if (!best || mismatched < best.mismatched) best = { mismatched, total, ink, offset: [dx, dy] };
    }
    return best!;
  }, { base64: shot.toString('base64'), index: PROBE_PAGE, box: geometry.box, viewport: geometry.viewport, block: BLOCK });
  return { devicePixelRatio: geometry.devicePixelRatio, backing: geometry.backing, box: geometry.box, pixels };
}

const fraction = (value: number) => Math.abs(value - Math.round(value));

type ControlVariant = 'css-integral' | 'device-integral' | 'layout-unit';

/**
 * Mismatched pixels of a bare canvas placed and sized on whole device pixels: what the browser itself does at this
 * ratio. `css-integral` also lands on whole CSS pixels, `device-integral` does not (like page boxes), and
 * `layout-unit` rounds those CSS lengths up to the 1/64 px layout unit.
 */
async function control(page: Page, variant: ControlVariant): Promise<{ mismatched: number; total: number }> {
  return pinnedEvaluate(page, async (variant) => {
    const ratio = devicePixelRatio;
    const offset = variant === 'css-integral' ? 96 : 97, size = variant === 'css-integral' ? 240 : 241;
    const css = (device: number) => variant === 'layout-unit' ? Math.ceil((device / ratio) * 64) / 64 : device / ratio;
    const canvas = document.createElement('canvas');
    canvas.width = canvas.height = size;
    Object.assign(canvas.style, { position: 'fixed', left: `${css(offset)}px`, top: `${css(offset)}px`,
      width: `${css(size)}px`, height: `${css(size)}px`, zIndex: '2147483647' });
    const context = canvas.getContext('2d')!;
    let seed = 7;
    for (let y = 0; y < size; y++) for (let x = 0; x < size; x++) {
      seed = (seed * 1103515245 + 12345) & 0x7fffffff;
      context.fillStyle = seed % 2 ? '#000' : '#fff';
      context.fillRect(x, y, 1, 1);
    }
    document.body.append(canvas);
    (window as unknown as { __snapControl: HTMLCanvasElement }).__snapControl = canvas;
  }, variant).then(async () => {
    await page.waitForTimeout(200);
    const shot = await page.screenshot({ scale: 'device' });
    return pinnedEvaluate(page, async ({ base64, variant }) => {
      const canvas = (window as unknown as { __snapControl: HTMLCanvasElement }).__snapControl;
      const offset = variant === 'css-integral' ? 96 : 97, size = canvas.width;
      const source = canvas.getContext('2d')!.getImageData(0, 0, size, size).data;
      const bytes = Uint8Array.from(atob(base64), (char) => char.charCodeAt(0));
      const image = await createImageBitmap(new Blob([bytes], { type: 'image/png' }));
      const screen = new OffscreenCanvas(image.width, image.height).getContext('2d')!;
      screen.drawImage(image, 0, 0);
      const pixels = screen.getImageData(offset, offset, size, size).data;
      let mismatched = 0;
      for (let index = 0; index < pixels.length; index += 4) if (Math.abs(pixels[index]! - source[index]!) > 2) mismatched += 1;
      canvas.remove();
      return { mismatched, total: size * size };
    }, { base64: shot.toString('base64'), variant });
  });
}

function faults(result: Probe, zoom: number, pixelExact: boolean): string[] {
  const found: string[] = [];
  const tolerance = Math.max(0.02, (result.devicePixelRatio * zoom) / 64);
  if (Math.abs(result.box.width - result.backing.width) > tolerance || Math.abs(result.box.height - result.backing.height) > tolerance) {
    found.push(`box ${result.box.width.toFixed(3)}x${result.box.height.toFixed(3)} != backing ${result.backing.width}x${result.backing.height}`);
  }
  if (fraction(result.box.left) > tolerance || fraction(result.box.top) > tolerance) {
    found.push(`box origin ${result.box.left.toFixed(3)},${result.box.top.toFixed(3)} off the device grid`);
  }
  if ('unreadable' in result.pixels) return [...found, `canvas unreadable: ${result.pixels.unreadable}`];
  if (!pixelExact) return found;
  if (result.pixels.ink < result.pixels.total * 0.01) found.push('probe block has no text');
  if (result.pixels.mismatched > result.pixels.total * 0.001) {
    found.push(`${result.pixels.mismatched}/${result.pixels.total} screen pixels differ from the canvas backing`);
  }
  return found;
}

const PROBE_IN_THREAD = process.env.PIXELSNAP_IN_THREAD === '1';

for (const worker of PROBE_IN_THREAD ? [true, false] : [true]) {
  for (const devicePixelRatio of [1, 1.5, 2]) {
    test(`docx ${worker ? 'worker' : 'in-thread'} editor: page boxes match their canvas backing on whole device pixels at DPR ${devicePixelRatio}`, async ({ browser }, info) => {
      test.setTimeout(300_000);
      const { context, page } = await open(browser, devicePixelRatio, worker);
      const failures: string[] = [];
      const results: Record<string, Probe> = {};
      try {
        const controls: Partial<Record<ControlVariant, { mismatched: number; total: number }>> = {};
        for (const variant of ['css-integral', 'device-integral', 'layout-unit'] as const) controls[variant] = await control(page, variant);
        console.log(`PIXELSNAP-CONTROL ${JSON.stringify({ worker, devicePixelRatio, controls })}`);
        for (const value of ZOOMS) {
          await setZoom(page, value);
          await settle(page, parseInt(value, 10) / 100);
          const result = await probe(page);
          results[value] = result;
          console.log(`PIXELSNAP ${JSON.stringify({ worker, zoom: value, ...result })}`);
          const zoom = parseInt(value, 10) / 100;
          const cssIntegral = [result.box.width, result.box.height].every((length) => fraction(length / result.devicePixelRatio) < 0.02);
          if (worker) for (const fault of faults(result, zoom, value === '100%' && cssIntegral)) failures.push(`${value}: ${fault}`);
        }
      } finally {
        await info.attach('pixel-snap.json', { body: JSON.stringify(results, null, 2), contentType: 'application/json' });
        await context.close();
      }
      expect(failures).toEqual([]);
    });
  }
}
