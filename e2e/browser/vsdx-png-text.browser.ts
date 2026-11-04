import { test, expect, type Page } from 'playwright/test';
import { readFile } from 'node:fs/promises';
import { extname, resolve } from 'node:path';
import JSZip from 'jszip';

const root = resolve(import.meta.dirname, '../..');
const types: Record<string, string> = { '.js': 'text/javascript', '.wasm': 'application/wasm' };
const TEXT = 'free text';
const SIZE_IN = 0.25;
const SCALE = 2;
const SPACING_IN = 0.1;

async function diagram(cells: string): Promise<Buffer> {
  const zip = await JSZip.loadAsync(await readFile(resolve(root, 'crates/vsdx-parse/tests/fixtures/text-lock.vsdx')));
  const shape = `<Shape ID='2' Type='Shape'><Cell N='Width' V='4'/><Cell N='Height' V='1'/><Cell N='PinX' V='3'/><Cell N='PinY' V='2'/><Cell N='LocPinX' V='2'/><Cell N='LocPinY' V='0.5'/><Section N='Character'><Row IX='0'><Cell N='Size' V='${SIZE_IN}'/>${cells}</Row></Section><Text>${TEXT}</Text></Shape>`;
  zip.file('visio/pages/page1.xml', `<PageContents xmlns='http://schemas.microsoft.com/office/visio/2012/main'><Shapes>${shape}</Shapes></PageContents>`);
  return zip.generateAsync({ type: 'nodebuffer' });
}

interface Ink { count: number; left: number; right: number; top: number; bottom: number; }

async function exportInk(page: Page, variants: Record<string, string>): Promise<Record<string, Ink & { differsFrom?: Ink }>> {
  const docs = Object.fromEntries(await Promise.all(Object.entries(variants).map(async ([name, cells]) => [name, await diagram(cells)] as const)));
  await page.route('**/__vsdx/**', async (route) => {
    const path = new URL(route.request().url()).pathname.replace('/__vsdx/', '');
    await route.fulfill({ body: await readFile(resolve(root, 'packages/vsdx/dist', path)), contentType: types[extname(path)] ?? 'application/octet-stream' });
  });
  await page.route('**/__fixture/doc/*', async (route) => route.fulfill({ body: docs[new URL(route.request().url()).pathname.split('/').pop()!] }));
  await page.route('**/__fixture/font', async (route) =>
    route.fulfill({ body: await readFile(resolve(root, 'packages/fonts/assets/ComicRelief-Regular.ttf')) }));
  await page.goto('/');
  return page.evaluate(async ({ names, scale }) => {
    const vsdx = await import(`${location.origin}/__vsdx/index.js`);
    await vsdx.initWasm();
    const font = new Uint8Array(await (await fetch('/__fixture/font')).arrayBuffer());
    const pixels: Record<string, ImageData> = {};
    for (const name of names) {
      const bytes = new Uint8Array(await (await fetch(`/__fixture/doc/${name}`)).arrayBuffer());
      const handle = vsdx.openDiagram(bytes, { clientId: 1, fonts: [{ family: 'Arial', bytes: font }] });
      try {
        const png = await handle.renderPng({ page: 0 }, { scale });
        const bitmap = await createImageBitmap(new Blob([png], { type: 'image/png' }));
        const canvas = new OffscreenCanvas(bitmap.width, bitmap.height);
        const ctx = canvas.getContext('2d')!;
        ctx.drawImage(bitmap, 0, 0);
        pixels[name] = ctx.getImageData(0, 0, canvas.width, canvas.height);
      } finally { handle.dispose(); }
    }
    const box = (hit: (index: number) => boolean, image: ImageData) => {
      const ink = { count: 0, left: Infinity, right: -Infinity, top: Infinity, bottom: -Infinity };
      for (let y = 0; y < image.height; y++) for (let x = 0; x < image.width; x++) {
        if (!hit((y * image.width + x) * 4)) continue;
        ink.count++; ink.left = Math.min(ink.left, x); ink.right = Math.max(ink.right, x); ink.top = Math.min(ink.top, y); ink.bottom = Math.max(ink.bottom, y);
      }
      return ink;
    };
    const base = pixels.base;
    return Object.fromEntries(names.map((name) => {
      const image = pixels[name];
      const ink = box((index) => image.data[index] < 128, image);
      const differs = name === 'base' || image.width !== base.width || image.height !== base.height ? undefined
        : box((index) => Math.abs(image.data[index] - base.data[index]) > 64, image);
      return [name, { ...ink, differsFrom: differs }];
    }));
  }, { names: Object.keys(variants), scale: SCALE });
}

test('an exported run paints with the registered face layout measured it with, spaces its glyphs and draws its underline', async ({ page }) => {
  const ink = await exportInk(page, {
    base: `<Cell N='Font' V='Arial'/>`,
    substituted: `<Cell N='Font' V='Calibri'/>`,
    spaced: `<Cell N='Font' V='Arial'/><Cell N='Letterspace' V='${SPACING_IN * 1440}'/>`,
    underlined: `<Cell N='Font' V='Arial'/><Cell N='Style' V='4'/>`,
  });
  const size = SIZE_IN * 96 * SCALE;
  expect(ink.base.count, 'the base run paints').toBeGreaterThan(100);

  expect(ink.substituted.differsFrom!.count, 'a Calibri run laid out with the registered Arial face paints with it').toBe(0);

  const spread = (ink.spaced.right - ink.spaced.left) - (ink.base.right - ink.base.left);
  expect(Math.abs(spread - (TEXT.length - 1) * SPACING_IN * 96 * SCALE), `spacing widened the run by ${spread}px`).toBeLessThan(4);

  const rule = ink.underlined.differsFrom!;
  expect(rule.count, 'the underline paints').toBeGreaterThan(0);
  expect(rule.top, 'the underline sits at or under the baseline').toBeGreaterThanOrEqual(ink.base.bottom - 2);
  expect(rule.bottom - rule.top, 'the underline is a rule, not a second line of text').toBeLessThan(0.15 * size);
  expect(Math.abs(rule.left - ink.base.left), 'the underline starts with the run').toBeLessThan(0.15 * size);
  expect(Math.abs(rule.right - ink.base.right), 'the underline ends with the run').toBeLessThan(0.25 * size);
});
