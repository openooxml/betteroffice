import { expect, test } from 'bun:test';
import { EchoGlyphSources } from './residentEchoGlyphSource';

test('retained glyph providers read shipped sets and cached ink without building', () => {
  let shapes = 0;
  let outlines = 0;
  const sources = new EchoGlyphSources({
    echoGlyphSetJson() {
      shapes += 1;
      return JSON.stringify({ horizontal: 0.01, glyphIds: [1], advances: [1000], kerning: [], unsafePairs: [] });
    },
    outlineGlyphJson() {
      outlines += 1;
      return JSON.stringify({ upem: 1000, cmds: [{ t: 'M', x: 0, y: 0 }, { t: 'L', x: 500, y: 700 }, { t: 'Z' }] });
    },
  }, 1);
  const retained = sources.provider(1, true);
  expect(retained.glyphSet(1, 12, false)).toBeUndefined();
  expect(retained.glyphInk(1, 1)).toBeUndefined();
  expect([shapes, outlines]).toEqual([0, 0]);
  const built = sources.provider(1).glyphSet(1, 12, false)!;
  expect(built).toBeDefined();
  expect(retained.glyphSet(1, 12, false)).toBeUndefined();
  sources.sent([built.glyphs.set]);
  expect(retained.glyphSet(1, 12, false)).toBe(built);
  const counts = [shapes, outlines];
  expect(retained.glyphInk(1, 1)).toBeDefined();
  expect(retained.glyphSet(1, 13, false)).toBeUndefined();
  expect(retained.glyphInk(1, 2)).toBeUndefined();
  expect([shapes, outlines]).toEqual(counts);
});
