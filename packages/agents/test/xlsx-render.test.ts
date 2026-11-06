import { expect, spyOn, test } from 'bun:test';
import { createCanvas, GlobalFonts, loadImage } from '@napi-rs/canvas';
import { createFontProvider, resolveLastResortFace } from '@betteroffice/fonts';
import { decodeFontFace, paintDisplayList } from '@betteroffice/xlsx';
import { openXlsx } from '../src/xlsx';
import { rangeExtent, renderXlsxRange } from '../src/xlsx-render';
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

test('XLSX canvas previews draw unbundled families with the Word last-resort face', async () => {
  const family = 'Humnst777 BT';
  expect(resolveLastResortFace(family, false, false, 'word').family).toBe('Liberation Sans');
  const provider = createFontProvider(undefined, { decoder: decodeFontFace });
  const word = Buffer.from(await provider.resolveLastResort(family, false, false, 'word')());
  expect(word.equals(Buffer.from(await provider.resolveLastResort(family, false, false)()))).toBe(false);
  const handle = {
    sheetInfo: () => ({ activeSheet: 0 }),
    setActiveSheet: () => {},
    displayList: () => ({ width: 40, height: 20, commands: [{ op: 'text', x: 2, y: 14, text: 'Hi', fontFamily: family, fontSize: 11, color: '#000000' }] }),
    renderPngFonts: () => [],
  };
  const register = spyOn(GlobalFonts, 'register');
  try {
    await renderXlsxRange(handle as never, 0, 'A1', { x: 0, y: 0, width: 40, height: 20 }, 1, provider);
    expect(register).toHaveBeenCalledTimes(1);
    expect(Buffer.from(register.mock.calls[0]![0]).equals(word)).toBe(true);
  } finally { register.mockRestore(); }
});

test('XLSX canvas previews in one process keep each workbook\'s faces apart', async () => {
  const base = createFontProvider(undefined, { decoder: decodeFontFace });
  const hebrew = await base.resolve('Noto Sans Hebrew', false, false)!();
  const sans = await base.resolve('Arial', false, false)!();
  const provider = (bytes: ArrayBuffer) => ({ ...base, resolve: () => async () => bytes });
  const book = (family: string) => ({
    sheetInfo: () => ({ activeSheet: 0 }),
    setActiveSheet: () => {},
    displayList: () => ({ width: 120, height: 24, commands: [{ op: 'text', x: 4, y: 18, text: 'Quarterly', fontFamily: family, fontSize: 12, color: '#000000' }] }),
    renderPngFonts: () => [],
  });
  const viewport = { x: 0, y: 0, width: 120, height: 24 };
  const alone = await renderXlsxRange(book('Unshared Face') as never, 0, 'A1', viewport, 1, provider(sans));
  const other = await renderXlsxRange(book('Shared Face') as never, 0, 'A1', viewport, 1, provider(hebrew));
  const shared = await renderXlsxRange(book('Shared Face') as never, 0, 'A1', viewport, 1, provider(sans));
  expect(Buffer.from(other).equals(Buffer.from(alone))).toBe(false);
  expect(Buffer.from(shared).equals(Buffer.from(alone))).toBe(true);
  const [overlappedOther, overlappedShared] = await Promise.all([
    renderXlsxRange(book('Shared Face') as never, 0, 'A1', viewport, 1, provider(hebrew)),
    renderXlsxRange(book('Shared Face') as never, 0, 'A1', viewport, 1, provider(sans)),
  ]);
  expect(Buffer.from(overlappedOther).equals(Buffer.from(other))).toBe(true);
  expect(Buffer.from(overlappedShared).equals(Buffer.from(alone))).toBe(true);
  expect(GlobalFonts.families.some(({ family }) => family.startsWith('BetterOffice Preview'))).toBe(false);
});

test('XLSX canvas preview sizes match the engine\'s 32-bit range extent', () => {
  const first = { x: 0, y: Math.fround(1.1) };
  const last = { x: Math.fround(10.3), y: Math.fround(799.9), w: Math.fround(20.7), h: Math.fround(1.2) };
  const extent = rangeExtent(first, last, 1);
  const f = Math.fround;
  expect(extent.pixelHeight).toBe(Math.ceil(f(f(f(last.y + last.h) - f(first.y)) * f(1))));
  expect(extent.pixelWidth).toBe(Math.ceil(f(f(last.x + last.w) - first.x)));
  expect(rangeExtent({ x: 0, y: 0 }, { x: 0, y: 0, w: 100, h: 50 }, 1.5)).toEqual({ width: 100, height: 50, pixelWidth: 150, pixelHeight: 75 });
});
