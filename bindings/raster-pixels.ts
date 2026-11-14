import { inflateSync } from 'node:zlib';
import { expect } from 'bun:test';

export function pngPixels(bytes: Uint8Array): Uint8Array {
  const png = Buffer.from(bytes);
  expect([...png.subarray(0, 8)]).toEqual([137, 80, 78, 71, 13, 10, 26, 10]);
  expect([png[24], png[25], png[28]]).toEqual([8, 6, 0]);
  const stride = png.readUInt32BE(16) * 4;
  const height = png.readUInt32BE(20);
  const chunks: Buffer[] = [];
  for (let offset = 8; offset < png.length;) {
    const size = png.readUInt32BE(offset);
    if (png.toString('ascii', offset + 4, offset + 8) === 'IDAT') {
      chunks.push(png.subarray(offset + 8, offset + 8 + size));
    }
    offset += size + 12;
  }
  const filtered = inflateSync(Buffer.concat(chunks));
  expect(filtered.length).toBe(height * (stride + 1));
  const pixels = new Uint8Array(height * stride);
  for (let row = 0; row < height; row++) {
    const filter = filtered[row * (stride + 1)];
    expect(filter).toBeLessThanOrEqual(4);
    for (let col = 0; col < stride; col++) {
      const index = row * stride + col;
      const left = col >= 4 ? pixels[index - 4] : 0;
      const up = row > 0 ? pixels[index - stride] : 0;
      const corner = row > 0 && col >= 4 ? pixels[index - stride - 4] : 0;
      const estimate = left + up - corner;
      const distances = [Math.abs(estimate - left), Math.abs(estimate - up), Math.abs(estimate - corner)];
      const paeth = distances[0] <= distances[1] && distances[0] <= distances[2]
        ? left : distances[1] <= distances[2] ? up : corner;
      const prediction = [0, left, up, Math.floor((left + up) / 2), paeth][filter];
      pixels[index] = (filtered[row * (stride + 1) + col + 1] + prediction) & 255;
    }
  }
  return pixels;
}

type Glyph = { x: number; advance: number };
type Rule = { x: number; y: number; w: number; h: number };

function image(bytes: Uint8Array) {
  const png = Buffer.from(bytes);
  return { width: png.readUInt32BE(16), height: png.readUInt32BE(20), pixels: pngPixels(bytes) };
}

function inkAt(png: ReturnType<typeof image>, x: number, y: number): boolean {
  const index = (y * png.width + x) * 4;
  return png.pixels[index + 3] > 64 && Math.min(png.pixels[index], png.pixels[index + 1], png.pixels[index + 2]) < 200;
}

function inkBounds(png: ReturnType<typeof image>) {
  let left = png.width;
  let right = -1;
  for (let y = 0; y < png.height; y++) {
    for (let x = 0; x < png.width; x++) {
      if (inkAt(png, x, y)) {
        left = Math.min(left, x);
        right = Math.max(right, x);
      }
    }
  }
  expect(right).toBeGreaterThanOrEqual(left);
  return { left, right: right + 1, width: right + 1 - left };
}

export function expectPngTracking(
  spaced: Uint8Array,
  plain: Uint8Array,
  glyphs: Glyph[],
  plainGlyphs: Glyph[],
  spacing: number,
  scale: number,
) {
  expect(glyphs.length).toBe(4);
  expect(plainGlyphs.length).toBe(glyphs.length);
  for (let index = 0; index < glyphs.length; index++) {
    expect(Math.abs(glyphs[index].x - plainGlyphs[index].x - index * spacing)).toBeLessThan(0.02);
  }
  const first = glyphs[0];
  const last = glyphs[glyphs.length - 1];
  const expectedLeft = first.x * scale;
  const expectedRight = (last.x + last.advance) * scale;
  const actual = inkBounds(image(spaced));
  const baseline = inkBounds(image(plain));
  expect(Math.abs(actual.left - expectedLeft)).toBeLessThanOrEqual(2);
  expect(Math.abs(actual.right - expectedRight)).toBeLessThanOrEqual(2);
  expect(Math.abs(actual.width - (expectedRight - expectedLeft))).toBeLessThanOrEqual(2);
  const expectedGrowth = spacing * (glyphs.length - 1) * scale;
  expect(actual.width).toBeGreaterThan(baseline.width);
  expect(Math.abs(actual.width - baseline.width - expectedGrowth)).toBeLessThanOrEqual(2);
}

export function expectPngRule(decorated: Uint8Array, plain: Uint8Array, rule: Rule, scale: number) {
  const actual = image(decorated);
  const baseline = image(plain);
  expect([actual.width, actual.height]).toEqual([baseline.width, baseline.height]);
  const left = Math.ceil(rule.x * scale);
  const right = Math.floor((rule.x + rule.w) * scale);
  const top = Math.floor(rule.y * scale);
  const bottom = Math.ceil((rule.y + rule.h) * scale);
  expect(left).toBeGreaterThanOrEqual(0);
  expect(right).toBeLessThanOrEqual(actual.width);
  expect(top).toBeGreaterThanOrEqual(0);
  expect(bottom).toBeLessThanOrEqual(actual.height);
  expect(right - left).toBeGreaterThan(10 * scale);
  const longest = (png: ReturnType<typeof image>) => {
    let best = 0;
    for (let y = top; y < bottom; y++) {
      let length = 0;
      for (let x = left; x < right; x++) {
        length = inkAt(png, x, y) ? length + 1 : 0;
        best = Math.max(best, length);
      }
    }
    return best;
  };
  expect(longest(actual)).toBeGreaterThanOrEqual((right - left) * 0.9);
  expect(longest(baseline)).toBeLessThan((right - left) * 0.5);
}
