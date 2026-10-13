import { expect, test } from 'bun:test';
import type { DocxEchoGlyphSet } from './residentCommandProtocol';
import { echoRound3, indexEchoGlyphSet, leadingWordWidth, outlineBounds, placeEchoGlyphs } from './residentEchoGlyphs';

function glyphSet(overrides: Partial<DocxEchoGlyphSet> = {}): DocxEchoGlyphSet {
  const alphabet = 'AVabx ';
  return { key: 'k', fontId: 1, alphabet, horizontal: Math.fround(Math.fround(960 / 72) / 2048),
    glyphIds: [...alphabet].map((char) => char.charCodeAt(0)), advances: [1366, 1366, 1139, 1139, 1024, 569],
    kerning: [[0, 1, -152]], unsafePairs: [[2, 4]], outlines: {}, ...overrides };
}

test('echo placement sums f32 advances in f64 and rounds pens as the engine does', () => {
  const glyphs = indexEchoGlyphSet(glyphSet());
  const origin = 1417 / 15;
  const h = glyphs.set.horizontal;
  const advances = [Math.fround((1366 - 152) * h), Math.fround(1366 * h), Math.fround(1139 * h)];
  const placed = placeEchoGlyphs(glyphs, 'AVa', '', origin)!;
  expect(placed.advances).toEqual(advances);
  expect(placed.glyphIds).toEqual([65, 86, 97]);
  expect(placed.xs).toEqual([echoRound3(origin), echoRound3(origin + advances[0]!),
    echoRound3(origin + advances[0]! + advances[1]!)]);
  expect(placed.end).toBe(origin + advances[0]! + advances[1]! + advances[2]!);
});

test('echo placement scales by the exact f32 the shortest JSON decimal stands for', () => {
  const glyphs = indexEchoGlyphSet(glyphSet({ horizontal: 0.0065104165 }));
  expect(glyphs.horizontal).toBe(Math.fround(0.0065104165));
  expect(glyphs.horizontal).not.toBe(0.0065104165);
  expect(placeEchoGlyphs(glyphs, 'a', '', 0)!.advances).toEqual([Math.fround(1139 * Math.fround(0.0065104165))]);
});

test('echo placement kerns the line end with the character after the line', () => {
  const glyphs = indexEchoGlyphSet(glyphSet());
  expect(placeEchoGlyphs(glyphs, 'A', 'V', 0)!.advances).toEqual([Math.fround(1214 * glyphs.set.horizontal)]);
  expect(placeEchoGlyphs(glyphs, 'A', '', 0)!.advances).toEqual([Math.fround(1366 * glyphs.set.horizontal)]);
});

test('echo placement declines characters and pairs the glyph set cannot place', () => {
  const glyphs = indexEchoGlyphSet(glyphSet());
  expect(placeEchoGlyphs(glyphs, 'ab', '', 0)).toBeDefined();
  expect(placeEchoGlyphs(glyphs, 'aZ', '', 0)).toBeUndefined();
  expect(placeEchoGlyphs(glyphs, 'ax', '', 0)).toBeUndefined();
  expect(placeEchoGlyphs(glyphs, 'a', 'x', 0)).toBeUndefined();
  expect(placeEchoGlyphs(glyphs, 'a', 'é', 0)).toBeUndefined();
  expect(placeEchoGlyphs(indexEchoGlyphSet(glyphSet({ glyphIds: [65, 86, 0, 98, 120, 32] })), 'a', '', 0)).toBeUndefined();
});

test('echo rounding rounds halves away from zero and never yields negative zero', () => {
  expect(echoRound3(0.0625)).toBe(0.063);
  expect(echoRound3(-0.0625)).toBe(-0.063);
  expect(Object.is(echoRound3(-0.0001), 0)).toBe(true);
});

test('leading word width stops at the first character a break can follow', () => {
  expect(leadingWordWidth('ab1 cd', [1, 2, 4, 8, 16, 32])).toBe(7);
  expect(leadingWordWidth('-ab', [1, 2, 4])).toBe(0);
});

test('outline bounds cover every point and control point', () => {
  expect(outlineBounds(JSON.stringify({ upem: 1000, cmds: [{ t: 'M', x: 10, y: -20 },
    { t: 'Q', cx: 50, cy: 800, x: 90, y: 0 }, { t: 'Z' }] })))
    .toEqual({ upem: 1000, xMin: 10, xMax: 90, yMin: -20, yMax: 800 });
  expect(outlineBounds(JSON.stringify({ upem: 1000, cmds: [] }))).toBeUndefined();
});
