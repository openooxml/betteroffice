import { expect, test } from "bun:test";
import { imageSize } from "./image-size.ts";

const ascii = (text) => [...text].map((char) => char.charCodeAt(0));
const be16 = (n) => [n >> 8, n & 0xff];
const le16 = (n) => [n & 0xff, n >> 8];
const be32 = (n) => [n >>> 24, (n >> 16) & 0xff, (n >> 8) & 0xff, n & 0xff];
const le24 = (n) => [n & 0xff, (n >> 8) & 0xff, (n >> 16) & 0xff];
const le32 = (n) => [...le24(n), n >>> 24];
const bytes = (...parts) => Uint8Array.from(parts.flat());
const riff = (chunk, payload) => bytes(ascii("RIFF"), le32(4 + 8 + payload.length), ascii("WEBP"), ascii(chunk), le32(payload.length), payload);

const PNG = bytes([0x89], ascii("PNG\r\n\x1a\n"), be32(13), ascii("IHDR"), be32(640), be32(480), [8, 6, 0, 0, 0]);
const GIF = bytes(ascii("GIF89a"), le16(320), le16(200), [0, 0, 0]);
const JPEG = bytes(
  [0xff, 0xd8],
  [0xff, 0xe0], be16(16), ascii("JFIF\0"), [1, 1, 0], be16(1), be16(1), [0, 0],
  [0xff, 0xdb], be16(4), [0, 0],
  [0xff, 0xc2], be16(17), [8], be16(1080), be16(1920), [3, 1, 0x22, 0, 2, 0x11, 1, 3, 0x11, 1],
);
const VP8 = riff("VP8 ", [0, 0, 0, 0x9d, 0x01, 0x2a, ...le16(800 | (1 << 14)), ...le16(600 | (2 << 14))]);
const VP8L = riff("VP8L", [0x2f, ...le32((1023 - 1) | ((767 - 1) << 14)), 0, 0, 0, 0, 0]);
const VP8X = riff("VP8X", [0x10, 0, 0, 0, ...le24(4096 - 1), ...le24(2160 - 1)]);

test("reads the size of each raster format", () => {
  expect(imageSize(PNG)).toEqual([640, 480]);
  expect(imageSize(GIF)).toEqual([320, 200]);
  expect(imageSize(JPEG)).toEqual([1920, 1080]);
  expect(imageSize(VP8)).toEqual([800, 600]);
  expect(imageSize(VP8L)).toEqual([1023, 767]);
  expect(imageSize(VP8X)).toEqual([4096, 2160]);
  for (const webp of [VP8, VP8L, VP8X]) expect(webp.length).toBe(30);
});

test("reads a view of a larger buffer at its own offset", () => {
  const padded = new Uint8Array(PNG.length + 7);
  padded.set(PNG, 7);
  expect(imageSize(padded.subarray(7))).toEqual([640, 480]);
});

test("returns null for truncated or unknown bytes instead of throwing", () => {
  for (const image of [PNG, GIF, JPEG, VP8, VP8L, VP8X]) {
    for (let length = 0; length < image.length; length++) {
      expect(() => imageSize(image.subarray(0, length))).not.toThrow();
    }
  }
  expect(imageSize(PNG.subarray(0, 23))).toBeNull();
  expect(imageSize(GIF.subarray(0, 9))).toBeNull();
  expect(imageSize(JPEG.subarray(0, JPEG.length - 12))).toBeNull();
  expect(imageSize(VP8X.subarray(0, 29))).toBeNull();
  expect(imageSize(riff("VP8Q", [0, 0, 0, 0, 0, 0, 0, 0, 0, 0]))).toBeNull();
  expect(imageSize(bytes([0xff, 0xd8, 0x00, 0x00, 0, 0, 0, 0, 0, 0, 0, 0]))).toBeNull();
  expect(imageSize(bytes(ascii("plain text, not an image")))).toBeNull();
});

test("reads an SVG's width and height, falling back to its viewBox", () => {
  const svg = (attrs) => new TextEncoder().encode(`<?xml version="1.0"?>\n<svg xmlns="http://www.w3.org/2000/svg" ${attrs}><rect/></svg>`);
  expect(imageSize(svg('width="276" height="206" viewBox="0 0 10 10"'))).toEqual([276, 206]);
  expect(imageSize(svg('width="120px" height="80.5px"'))).toEqual([120, 80.5]);
  expect(imageSize(svg('viewBox="0 0 120 80"'))).toEqual([120, 80]);
  expect(imageSize(svg('width="100%" height="100%" viewBox="-5,-5,48,24"'))).toEqual([48, 24]);
  expect(imageSize(svg('width="100%"'))).toBeNull();
});
