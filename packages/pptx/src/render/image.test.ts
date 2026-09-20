import { describe, expect, test } from 'bun:test';
import { readFile } from 'node:fs/promises';
import { resolve } from 'node:path';
import { initWasm, openPresentation } from '../wasm/loader';
import { presentationImageBlob } from './image';

function record(command: number, payload: Uint8Array): Uint8Array<ArrayBuffer> {
  const bytes = new Uint8Array(6 + payload.length);
  const view = new DataView(bytes.buffer);
  view.setUint32(0, bytes.length / 2, true);
  view.setUint16(4, command, true);
  bytes.set(payload, 6);
  return bytes;
}

function words(...values: number[]): Uint8Array<ArrayBuffer> {
  const bytes = new Uint8Array(values.length * 2);
  const view = new DataView(bytes.buffer);
  values.forEach((value, index) => view.setInt16(index * 2, value, true));
  return bytes;
}

function metafile(records: Uint8Array[], placeable = false): Uint8Array<ArrayBuffer> {
  const start = placeable ? 22 : 0;
  const bytes = new Uint8Array(start + 18 + records.reduce((sum, item) => sum + item.length, 0));
  const view = new DataView(bytes.buffer);
  if (placeable) view.setUint32(0, 0x9ac6cdd7, true);
  view.setUint16(start, 1, true);
  view.setUint16(start + 2, 9, true);
  view.setUint16(start + 4, 0x0300, true);
  view.setUint32(start + 6, (bytes.length - start) / 2, true);
  let offset = start + 18;
  for (const item of records) {
    bytes.set(item, offset);
    offset += item.length;
  }
  return bytes;
}

function bitmapRecord(): Uint8Array<ArrayBuffer> {
  const payload = new Uint8Array(20 + 40 + 16);
  const view = new DataView(payload.buffer);
  view.setUint32(0, 0x00cc0020, true);
  view.setInt16(4, 2, true);
  view.setInt16(6, 2, true);
  view.setInt16(12, -2, true);
  view.setInt16(14, 2, true);
  view.setUint32(20, 40, true);
  view.setInt32(24, 2, true);
  view.setInt32(28, 2, true);
  view.setUint16(32, 1, true);
  view.setUint16(34, 24, true);
  view.setUint32(40, 16, true);
  payload.set([255, 0, 0, 255, 255, 255, 0, 0, 0, 0, 255, 0, 255, 0, 0, 0], 60);
  return record(0x0b41, payload);
}

function bitmapMetafile(bitmap = bitmapRecord(), extras: Uint8Array[] = [], placeable = false): Uint8Array<ArrayBuffer> {
  return metafile([
    record(0x0103, words(8)),
    record(0x020c, words(-2, 2)),
    record(0x020b, words(0, 0)),
    record(0x0107, words(4, 0)),
    bitmap,
    ...extras,
    record(0, new Uint8Array()),
  ], placeable);
}

function dib(width: number, height: number): Uint8Array<ArrayBuffer> {
  const stride = Math.ceil((width * 24) / 32) * 4;
  const bytes = new Uint8Array(40 + stride * height);
  const view = new DataView(bytes.buffer);
  view.setUint32(0, 40, true);
  view.setInt32(4, width, true);
  view.setInt32(8, height, true);
  view.setUint16(12, 1, true);
  view.setUint16(14, 24, true);
  bytes.fill(0x7f, 40);
  return bytes;
}

function stretchDibits(pixels: Uint8Array, width: number, height: number): Uint8Array<ArrayBuffer> {
  const bytes = new Uint8Array(80 + pixels.length);
  const view = new DataView(bytes.buffer);
  view.setUint32(0, 81, true);
  view.setUint32(4, bytes.length, true);
  view.setInt32(40, width, true);
  view.setInt32(44, height, true);
  view.setUint32(48, 80, true);
  view.setUint32(52, 40, true);
  view.setUint32(56, 120, true);
  view.setUint32(60, pixels.length - 40, true);
  view.setUint32(68, 0x00cc0020, true);
  view.setInt32(72, width, true);
  view.setInt32(76, height, true);
  bytes.set(pixels, 80);
  return bytes;
}

function enhancedMetafile(records: Uint8Array[], width = 4, height = 3): Uint8Array<ArrayBuffer> {
  const eof = new Uint8Array(20);
  const end = new DataView(eof.buffer);
  end.setUint32(0, 14, true);
  end.setUint32(4, 20, true);
  const body = [...records, eof];
  const bytes = new Uint8Array(88 + body.reduce((sum, item) => sum + item.length, 0));
  const view = new DataView(bytes.buffer);
  view.setUint32(0, 1, true);
  view.setUint32(4, 88, true);
  view.setInt32(16, width - 1, true);
  view.setInt32(20, height - 1, true);
  view.setUint32(40, 0x464d4520, true);
  let offset = 88;
  for (const item of body) {
    bytes.set(item, offset);
    offset += item.length;
  }
  return bytes;
}

describe('presentation image blobs', () => {
  test('rejects oversized TIFF media before transferring it to Wasm', () => {
    const bytes = new Uint8Array(32 * 1024 * 1024 + 1);
    bytes.set([0x49, 0x49, 0x2a, 0]);
    expect(() => presentationImageBlob(bytes)).toThrow('TIFF image exceeds the browser transfer budget');
  });

  test('transcodes TIFF media from a presentation to PNG', async () => {
    const [wasm, pptx] = await Promise.all([
      readFile(resolve(import.meta.dir, '../wasm/generated/pptx_wasm_bg.wasm')),
      readFile(resolve(import.meta.dir, 'fixtures/tiff-image.pptx')),
    ]);
    await initWasm(wasm);
    const presentation = openPresentation(pptx);
    try {
      const blob = presentationImageBlob(presentation.mediaBytes('ppt/media/image1.tiff'));
      expect(blob.type).toBe('image/png');
      expect(new Uint8Array(await blob.arrayBuffer()).subarray(0, 8)).toEqual(
        new Uint8Array([137, 80, 78, 71, 13, 10, 26, 10])
      );
    } finally {
      presentation.dispose();
    }
  });

  test('preserves ordinary media bytes', async () => {
    const bytes = new Uint8Array([137, 80, 78, 71, 13, 10, 26, 10]);
    expect(new Uint8Array(await presentationImageBlob(bytes).arrayBuffer())).toEqual(bytes);
  });

  test.each([false, true])('unwraps a complete raster WMF (placeable=%s)', async (placeable) => {
    const bytes = bitmapMetafile(bitmapRecord(), [], placeable);
    const before = bytes.slice();
    const blob = presentationImageBlob(bytes);
    const bitmap = new Uint8Array(await blob.arrayBuffer());
    const header = new DataView(bitmap.buffer);
    expect(blob.type).toBe('image/bmp');
    expect(header.getUint16(0, true)).toBe(0x4d42);
    expect(header.getUint32(2, true)).toBe(70);
    expect(header.getUint32(10, true)).toBe(54);
    expect(bitmap.subarray(14)).toEqual(bitmapRecord().subarray(26));
    expect(bytes).toEqual(before);
  });

  test('reads a view without confusing its backing buffer offsets', async () => {
    const bytes = bitmapMetafile();
    const padded = new Uint8Array(bytes.length + 20);
    padded.set(bytes, 10);
    expect(presentationImageBlob(padded.subarray(10, 10 + bytes.length)).type).toBe('image/bmp');
  });

  test.each([false, true])('preserves wrappers whose extents precede anisotropic mapping (placeable=%s)', async (placeable) => {
    for (const lateMode of [[], [record(0x0103, words(8))]]) {
      const bytes = metafile([
        record(0x020c, words(-2, 2)),
        ...lateMode,
        record(0x020b, words(0, 0)),
        bitmapRecord(),
        record(0, new Uint8Array()),
      ], placeable);
      const blob = presentationImageBlob(bytes);
      expect(blob.type).toBe('');
      expect(new Uint8Array(await blob.arrayBuffer())).toEqual(bytes);
    }
  });

  test('does not discard vector drawing or additional bitmap records', async () => {
    for (const extra of [record(0x0213, words(1, 1)), bitmapRecord()]) {
      const bytes = bitmapMetafile(bitmapRecord(), [extra]);
      expect(new Uint8Array(await presentationImageBlob(bytes).arrayBuffer())).toEqual(bytes);
    }
  });

  test('does not replace cropped, mirrored or composited source pixels with an unmodified bitmap', async () => {
    for (const [offset, value] of [[6, 0x00ee0086], [10, -2], [14, 1], [18, 2], [22, 1]]) {
      const bitmap = bitmapRecord();
      const view = new DataView(bitmap.buffer);
      if (offset === 6) view.setUint32(offset, value, true);
      else view.setInt16(offset, value, true);
      const bytes = bitmapMetafile(bitmap);
      expect(presentationImageBlob(bytes).type).toBe('');
    }
  });

  test('unwraps a complete raster EMF', async () => {
    const pixels = dib(4, 3);
    const bytes = enhancedMetafile([stretchDibits(pixels, 4, 3)]);
    const before = bytes.slice();
    const blob = presentationImageBlob(bytes);
    const bitmap = new Uint8Array(await blob.arrayBuffer());
    const header = new DataView(bitmap.buffer);
    expect(blob.type).toBe('image/bmp');
    expect(header.getUint16(0, true)).toBe(0x4d42);
    expect(header.getUint32(2, true)).toBe(14 + pixels.length);
    expect(header.getUint32(10, true)).toBe(54);
    expect(bitmap.subarray(14)).toEqual(pixels);
    expect(bytes).toEqual(before);
  });

  test('unwraps a raster EMF read as a view into a larger buffer', async () => {
    const bytes = enhancedMetafile([stretchDibits(dib(4, 3), 4, 3)]);
    const padded = new Uint8Array(bytes.length + 20);
    padded.set(bytes, 10);
    expect(presentationImageBlob(padded.subarray(10, 10 + bytes.length)).type).toBe('image/bmp');
  });

  test('preserves EMF metafiles that carry vector ink or a second blit', async () => {
    const polyline = new Uint8Array(16);
    new DataView(polyline.buffer).setUint32(0, 87, true);
    new DataView(polyline.buffer).setUint32(4, 16, true);
    for (const extra of [[polyline], [stretchDibits(dib(4, 3), 4, 3)]]) {
      const bytes = enhancedMetafile([stretchDibits(dib(4, 3), 4, 3), ...extra]);
      expect(new Uint8Array(await presentationImageBlob(bytes).arrayBuffer())).toEqual(bytes);
    }
  });

  test('preserves EMF blits that scale, offset or composite their source', async () => {
    for (const [offset, value] of [[32, 1], [40, 3], [68, 0x00ee0086], [64, 1], [72, 3], [24, 1]]) {
      const blit = stretchDibits(dib(4, 3), 4, 3);
      new DataView(blit.buffer).setInt32(offset, value, true);
      expect(presentationImageBlob(enhancedMetafile([blit])).type).toBe('');
    }
  });

  test('bounds an EMF that declares an empty bitmap payload', async () => {
    const record = stretchDibits(dib(4, 3), 4, 3).subarray(0, 80);
    const empty = new Uint8Array(80);
    empty.set(record);
    const view = new DataView(empty.buffer);
    view.setUint32(4, 80, true);
    view.setUint32(52, 0, true);
    view.setUint32(56, 80, true);
    view.setUint32(60, 0, true);
    const bytes = enhancedMetafile([empty]);
    expect(presentationImageBlob(bytes).type).toBe('');
  });

  test('bounds malformed EMF records and rejects truncated metafiles', async () => {
    const short = enhancedMetafile([stretchDibits(dib(4, 3), 4, 3)]);
    const truncated = short.subarray(0, short.length - 2);
    const oversized = enhancedMetafile([stretchDibits(dib(4, 3), 4, 3)]);
    new DataView(oversized.buffer).setUint32(92, 0, true);
    const understated = stretchDibits(dib(4, 3), 4, 3);
    new DataView(understated.buffer).setUint32(60, dib(4, 3).length - 44, true);
    for (const bytes of [truncated, oversized, enhancedMetafile([understated])]) {
      expect(new Uint8Array(await presentationImageBlob(bytes).arrayBuffer())).toEqual(bytes);
    }
  });

  test('bounds malformed records and rejects incomplete or oversized bitmaps', async () => {
    const broken = bitmapMetafile();
    new DataView(broken.buffer).setUint32(18, 0, true);
    const incomplete = bitmapMetafile().subarray(0, bitmapMetafile().length - 2);
    const oversizedBitmap = bitmapRecord();
    const view = new DataView(oversizedBitmap.buffer);
    view.setInt16(10, 30000, true);
    view.setInt16(12, 30000, true);
    view.setInt32(30, 30000, true);
    view.setInt32(34, 30000, true);
    for (const bytes of [new Uint8Array(), broken, incomplete, bitmapMetafile(oversizedBitmap)]) {
      expect(new Uint8Array(await presentationImageBlob(bytes).arrayBuffer())).toEqual(bytes);
    }
  });
});
