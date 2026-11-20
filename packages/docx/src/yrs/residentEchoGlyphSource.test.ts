import { expect, test } from 'bun:test';
import type { DocxCommandEchoContext } from './residentCommandProtocol';
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

test('synthetic sets ship the widened outlines and their own ink', () => {
  const calls: unknown[][] = [];
  const sources = new EchoGlyphSources({
    echoGlyphSetJson: () => JSON.stringify({ horizontal: 0.01, glyphIds: [1], advances: [100],
      kerning: [], unsafePairs: [] }),
    outlineGlyphJson(fontId, glyphId, syntheticBold) {
      calls.push([fontId, glyphId, syntheticBold === true]);
      const xMin = syntheticBold ? -40 : 0;
      const xMax = syntheticBold ? 540 : 500;
      return JSON.stringify({ upem: 1000, cmds: [{ t: 'M', x: xMin, y: 0 }, { t: 'L', x: xMax, y: 700 }, { t: 'Z' }] });
    },
  }, 1);
  const provider = sources.provider(1);
  const regular = provider.glyphSet(1, 12, false)!;
  const synthetic = provider.glyphSet(1, 12, false, true)!;
  expect(calls).toEqual([[1, 1, false], [1, 1, true]]);
  expect(synthetic.key).not.toBe(regular.key);
  expect(synthetic.glyphs.set.syntheticBold).toBe(true);
  expect(regular.glyphs.set.syntheticBold).toBeUndefined();
  // the widened contours widen the ink both left and past the advance
  expect(regular.ink).toEqual({ upem: 1000, xMin: 0, yMin: 0, yMax: 700, overhang: 400 });
  expect(synthetic.ink).toEqual({ upem: 1000, xMin: -40, yMin: 0, yMax: 700, overhang: 440 });
  expect(provider.glyphInk(1, 1, true)).toMatchObject({ xMin: -40 });
  expect(provider.glyphInk(1, 1)).toMatchObject({ xMin: 0 });
  const echo = (glyphSet: string) => ({ glyphs: { glyphSet } } as unknown as DocxCommandEchoContext);
  expect(sources.unsent(echo(synthetic.key))).toEqual([synthetic.glyphs.set]);
  expect(sources.unsent(echo(regular.key))).toEqual([regular.glyphs.set]);
  sources.sent([synthetic.glyphs.set]);
  expect(sources.unsent(echo(synthetic.key))).toEqual([]);
  expect(sources.unsent(echo(regular.key))).toEqual([regular.glyphs.set]);
  const retained = sources.provider(1, true);
  expect(retained.glyphSet(1, 12, false, true)).toBe(synthetic);
  expect(retained.glyphSet(1, 12, false)).toBeUndefined();
  expect(retained.glyphInk(1, 1, true)).toMatchObject({ xMin: -40 });
  expect(retained.glyphInk(1, 1)).toMatchObject({ xMin: 0 });
});
