import { expect, test, type Page } from 'playwright/test';
import JSZip from 'jszip';
import type { DocxCommandFrame } from '../../packages/docx/src/yrs/residentCommandProtocol';
import type { DocxEngineWorkerResponse } from '../../packages/docx/src/yrs/residentEngineWorkerProtocol';

interface RasterProbe {
  mode: 'off' | 'synthetic' | 'fail';
  hold: boolean;
  failed: number;
  expected: Map<string, HTMLCanvasElement>;
  release(): void;
}

declare global {
  interface Window {
    __docxRasterProbe?: RasterProbe;
  }
}

async function installRasterProbe(page: Page) {
  await page.addInitScript(() => {
    const replayed = new WeakSet<Event>();
    const held: { worker: Worker; data: DocxEngineWorkerResponse }[] = [];
    let seed = 0;
    const probe: RasterProbe = {
      mode: 'off',
      hold: false,
      failed: 0,
      expected: new Map(),
      release() {
        probe.hold = false;
        for (const { worker, data } of held.splice(0)) {
          const event = new MessageEvent('message', { data });
          replayed.add(event);
          worker.dispatchEvent(event);
        }
      },
    };
    window.__docxRasterProbe = probe;
    const alphas = [0, 96, 255];
    const synthetic = (width: number, height: number) => {
      seed += 1;
      const surface = new OffscreenCanvas(width, height);
      const context = surface.getContext('2d')!;
      const image = context.createImageData(width, height);
      for (let y = 0; y < height; y++) {
        for (let x = 0; x < width; x++) {
          const offset = (y * width + x) * 4;
          const band = (x >> 4) + (y >> 4) + seed;
          image.data[offset] = (x * 7 + seed * 31) & 255;
          image.data[offset + 1] = (y * 5 + seed * 17) & 255;
          image.data[offset + 2] = (band * 23) & 255;
          image.data[offset + 3] = alphas[band % 3]!;
        }
      }
      context.putImageData(image, 0, 0);
      return surface.transferToImageBitmap();
    };
    const reference = (bitmap: ImageBitmap) => {
      const canvas = document.createElement('canvas');
      canvas.width = bitmap.width;
      canvas.height = bitmap.height;
      canvas.getContext('2d')!.drawImage(bitmap, 0, 0);
      return canvas;
    };
    const NativeWorker = window.Worker;
    window.Worker = class extends NativeWorker {
      constructor(url: string | URL, options?: WorkerOptions) {
        super(url, options);
        this.addEventListener('message', (event: MessageEvent<DocxEngineWorkerResponse>) => {
          const message = event.data;
          if (replayed.has(event) || !message.ok) return;
          const frames = [message.commandInitial?.frame, message.commandAck?.frame, message.commandFrame?.frame]
            .filter((frame): frame is DocxCommandFrame => frame !== undefined);
          if (!frames.length) return;
          if (probe.hold) {
            event.stopImmediatePropagation();
            held.push({ worker: this, data: message });
            return;
          }
          for (const frame of frames) {
            delete frame.payload.echoContext;
            const pages = frame.payload.stagedPresentation?.pages ?? [];
            if (probe.mode === 'off' || !pages.length) continue;
            pages.forEach((item, index) => {
              const { width, height } = item.bitmap;
              item.bitmap.close();
              if (probe.mode === 'fail' && index === pages.length - 1) {
                item.bitmap = { width, height, close() {} } as unknown as ImageBitmap;
                return;
              }
              item.bitmap = synthetic(width, height);
              if (probe.mode === 'synthetic') probe.expected.set(item.page.pageId, reference(item.bitmap));
            });
            if (probe.mode === 'fail') {
              probe.mode = 'off';
              probe.hold = true;
              probe.failed += 1;
            }
          }
        }, true);
      }
    };
  });
}

async function installedPixels(page: Page) {
  return page.evaluate(() => {
    let compared = 0;
    for (const [pageId, reference] of window.__docxRasterProbe!.expected) {
      const canvas = document.querySelector<HTMLCanvasElement>(`[data-docx-worker-editor] [data-page-id="${pageId}"] canvas`);
      if (!canvas) continue;
      if (canvas.width !== reference.width || canvas.height !== reference.height) {
        return `page ${pageId}: ${canvas.width}x${canvas.height}, staged ${reference.width}x${reference.height}`;
      }
      const actual = canvas.getContext('2d')!.getImageData(0, 0, canvas.width, canvas.height).data;
      const staged = reference.getContext('2d')!.getImageData(0, 0, reference.width, reference.height).data;
      let differences = 0;
      for (let index = 0; index < actual.length; index++) if (actual[index] !== staged[index]) differences += 1;
      if (differences) return `page ${pageId}: ${differences} bytes differ`;
      compared += 1;
    }
    return compared ? 'identical' : 'no installed staged page';
  });
}

async function documentBytes() {
  const zip = new JSZip();
  zip.file('[Content_Types].xml', '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
    '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
    '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>');
  zip.file('_rels/.rels', '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
    '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>');
  zip.file('word/document.xml', '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>' +
    '<w:p><w:r><w:t>Device pixel text</w:t></w:r></w:p>' +
    '<w:p><w:pPr><w:pageBreakBefore/></w:pPr><w:r><w:t>Second page</w:t></w:r></w:p>' +
    '<w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/></w:sectPr>' +
    '</w:body></w:document>');
  return zip.generateAsync({ type: 'nodebuffer' });
}

async function backingSizes(page: Page, dpr: number) {
  return page.locator('[data-docx-worker-editor] [data-page-id] canvas').evaluateAll((canvases, ratio) => {
    const visible = canvases.filter((canvas) => {
      const rect = canvas.getBoundingClientRect();
      return rect.bottom > 0 && rect.top < innerHeight && rect.right > 0 && rect.left < innerWidth;
    });
    return visible.length > 0 && visible.every((element) => {
      const canvas = element as HTMLCanvasElement;
      const rect = canvas.getBoundingClientRect();
      return Math.abs(canvas.width - Math.round(rect.width * ratio)) <= 1 &&
        Math.abs(canvas.height - Math.round(rect.height * ratio)) <= 1;
    });
  }, dpr);
}

for (const dpr of [1, 2]) {
  test(`worker pages rasterize at device pixels at DPR ${dpr}`, async ({ browser }) => {
    test.skip(process.env.DOCX_PEERLESS_HARNESS !== '1', 'Harness capabilities require DOCX_PEERLESS_HARNESS=1');
    const context = await browser.newContext({ deviceScaleFactor: dpr, viewport: { width: 1440, height: 1000 } });
    const page = await context.newPage();
    try {
      await page.goto('/?format=docx&peerless=1');
      await page.locator('input[type=file]').last().setInputFiles({ name: 'raster.docx',
        mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document', buffer: await documentBytes() });
      await expect.poll(() => page.evaluate(() => window.__docxPeerlessHarness?.capabilities), {
        message: 'Build with DOCX_PEERLESS_HARNESS=1',
      }).toMatchObject({ presentation: true, mode: true });
      await expect(page.locator('[data-docx-worker-editor] [data-page-id] canvas').first()).toBeVisible();
      await expect.poll(() => backingSizes(page, dpr)).toBe(true);
      await page.getByRole('combobox', { name: /Zoom/ }).click();
      await page.getByRole('option', { name: '150%', exact: true }).click();
      await expect(page.getByRole('combobox', { name: /Zoom/ })).toHaveText('150%');
      await expect.poll(() => backingSizes(page, dpr)).toBe(true);
    } finally { await context.close(); }
  });
}

test('installed canvases hold exactly the staged pixels across transforms, transparency, resizes and rollback', async ({ browser }) => {
  const context = await browser.newContext({ deviceScaleFactor: 2, viewport: { width: 1440, height: 1000 } });
  const page = await context.newPage();
  try {
    await installRasterProbe(page);
    await page.goto('/?format=docx&peerless=1');
    await page.locator('input[type=file]').last().setInputFiles({ name: 'raster.docx',
      mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document', buffer: await documentBytes() });
    const canvas = page.locator('[data-docx-worker-editor] [data-page-id] canvas').first();
    await expect(canvas).toBeVisible();
    await expect.poll(() => backingSizes(page, 2)).toBe(true);
    const zoom = async (level: string) => {
      await page.getByRole('combobox', { name: /Zoom/ }).click();
      await page.getByRole('option', { name: level, exact: true }).click();
      await expect(page.getByRole('combobox', { name: /Zoom/ })).toHaveText(level);
    };
    await page.evaluate(() => { window.__docxRasterProbe!.mode = 'synthetic'; });
    await zoom('150%');
    await expect.poll(() => installedPixels(page)).toBe('identical');
    await canvas.evaluate((element) => {
      const target = element as HTMLCanvasElement;
      const context = target.getContext('2d')!;
      context.fillStyle = '#ff00ff';
      context.fillRect(0, 0, target.width, target.height);
      context.setTransform(0.5, 0.1, -0.1, 0.5, 37, 19);
    });
    await canvas.click({ position: { x: 150, y: 150 } });
    const input = page.locator('[data-docx-worker-editor] textarea');
    await expect(input).toBeEditable();
    await input.pressSequentially('x');
    await expect.poll(() => installedPixels(page)).toBe('identical');
    await page.evaluate(() => { window.__docxRasterProbe!.mode = 'fail'; });
    await zoom('100%');
    await expect.poll(() => page.evaluate(() => window.__docxRasterProbe!.failed)).toBe(1);
    await expect.poll(() => installedPixels(page)).toBe('identical');
    await page.evaluate(() => window.__docxRasterProbe!.release());
    await expect.poll(() => backingSizes(page, 2)).toBe(true);
  } finally { await context.close(); }
});
