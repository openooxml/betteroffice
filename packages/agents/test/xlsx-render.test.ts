import { expect, spyOn, test } from 'bun:test';
import { createCanvas, loadImage } from '@napi-rs/canvas';
import { paintDisplayList } from '@betteroffice/xlsx';
import { openXlsx } from '../src/xlsx';
import { xlsxFixture } from './format-fixtures';

test('XLSX previews paint a non-active sheet range at the requested size', async () => {
  const book = await openXlsx(await xlsxFixture());
  const render = spyOn(book.handle, 'renderPng');
  try {
    book.handle.setActiveSheet(0);
    book.formatRanges({ version: book.handle.version(), edits: [{ sheet: 'sheet:1', range: 'B2:C4', fill: '#ff0000' }] });
    const first = book.handle.cellRect(1, 1, 1);
    const last = book.handle.cellRect(1, 3, 2);
    const version = book.handle.version();
    const saved = await book.export();
    const previewing = book.previewRange({ sheet: 'sheet:1', range: 'B2:C4', scale: 1.5 });
    expect(book.handle.sheetInfo().activeSheet).toBe(0);
    const preview = await previewing;
    const image = await loadImage(Buffer.from(preview.png));
    expect([...preview.png.subarray(0, 8)]).toEqual([137, 80, 78, 71, 13, 10, 26, 10]);
    expect(preview).toMatchObject({
      version, sheet: 'sheet:1', range: 'B2:C4', warnings: [],
      width: Math.ceil((last.x + last.w - first.x) * 1.5),
      height: Math.ceil((last.y + last.h - first.y) * 1.5),
    });
    expect(image.width).toBe(preview.width);
    expect(image.height).toBe(preview.height);
    const canvas = createCanvas(image.width, image.height);
    const ctx = canvas.getContext('2d');
    ctx.drawImage(image, 0, 0);
    expect([...ctx.getImageData(5, 5, 1, 1).data]).toEqual([255, 0, 0, 255]);
    expect(book.overview().capabilities.render).toBe(true);
    expect(book.handle.sheetInfo().activeSheet).toBe(0);
    expect(book.handle.version()).toBe(version);
    expect(await book.export()).toEqual(saved);
  } finally { render.mockRestore(); book.close(); }
});

test('XLSX pattern fills use the supplied canvas without browser globals', () => {
  const canvas = createCanvas(16, 16);
  const ctx = canvas.getContext('2d');
  paintDisplayList(ctx as unknown as CanvasRenderingContext2D, {
    width: 16, height: 16,
    commands: [{ op: 'patternFill', x: 0, y: 0, w: 16, h: 16, pattern: 'darkDown', fg: '#ff0000', bg: '#ffffff', originX: 0, originY: 0 }],
  }, 1);
  const pixels = ctx.getImageData(0, 0, 16, 16).data;
  const green = new Set<number>();
  for (let at = 0; at < pixels.length; at += 4) {
    expect(pixels[at]).toBe(255);
    green.add(pixels[at + 1]);
  }
  expect(Math.min(...green)).toBeLessThan(128);
  expect(Math.max(...green)).toBeGreaterThan(128);
});

test('XLSX previews use the engine PNG where it can draw', async () => {
  const book = await openXlsx(await xlsxFixture());
  const canvas = createCanvas(12, 7);
  const png = new Uint8Array(await canvas.encode('png'));
  const render = spyOn(book.handle, 'renderPng').mockResolvedValue(png);
  try {
    const preview = await book.previewRange({ sheet: 'sheet:1', range: 'B2:C4' });
    expect(render).toHaveBeenCalledTimes(1);
    expect(preview.png).toBe(png);
    expect([preview.width, preview.height]).toEqual([12, 7]);
  } finally { render.mockRestore(); book.close(); }
});
