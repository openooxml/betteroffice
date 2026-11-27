import { expect, test } from "bun:test";
import { readFileSync } from "node:fs";
import { join } from "node:path";

const ICO = readFileSync(join(import.meta.dir, "favicon.ico"));
const PNG_SIGNATURE = "89504e470d0a1a0a";

test("favicon.ico holds 16, 32 and 48 px entries as 8-bit RGBA PNGs", () => {
  expect(ICO.readUInt16LE(0)).toBe(0);
  expect(ICO.readUInt16LE(2)).toBe(1);
  const entries = Array.from({ length: ICO.readUInt16LE(4) }, (_, i) => {
    const dir = 6 + 16 * i;
    const offset = ICO.readUInt32LE(dir + 12);
    const png = ICO.subarray(offset, offset + ICO.readUInt32LE(dir + 8));
    expect(png.subarray(0, 8).toString("hex")).toBe(PNG_SIGNATURE);
    expect(png.subarray(12, 16).toString("ascii")).toBe("IHDR");
    return { width: ICO[dir] || 256, height: ICO[dir + 1] || 256, pngWidth: png.readUInt32BE(16), pngHeight: png.readUInt32BE(20), depth: png[24], colorType: png[25] };
  });
  expect(entries).toEqual([16, 32, 48].map((size) => ({ width: size, height: size, pngWidth: size, pngHeight: size, depth: 8, colorType: 6 })));
});
