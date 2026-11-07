import { expect, test } from 'bun:test';
import { readFileSync } from 'node:fs';
import type { PositionedTextRun } from '../types';
import { fittedPieceDecisions } from './textChunks';

const fixture = JSON.parse(readFileSync(
  new URL('../../../../crates/pptx-raster/tests/fixtures/piece-decisions.json', import.meta.url), 'utf8',
)) as { cases: Array<{ run: PositionedTextRun; decisions: Array<[number, number]> | null }> };

test('fits the pieces of random runs exactly as the native engine does', () => {
  expect(fixture.cases.length).toBeGreaterThan(300);
  for (const { run, decisions } of fixture.cases) {
    const canvas = fittedPieceDecisions(run, (text) => text.length * 8);
    expect([run.text, canvas?.map(({ start, scale }) => [start, scale]) ?? null]).toEqual([run.text, decisions]);
  }
});

test('a run whose only departure is its JSON decimal rounding is not fitted', () => {
  const glyph = (cluster: number, x: number, advance: number) => ({ glyphId: cluster + 1, cluster, x, advance, xOffset: 0, yOffset: 0 });
  const run = {
    text: 'aTo', start: 0, end: 3, x: 2040, width: 30.3798828125, fontId: 1, fontFamily: 'Corbel', fontSizePx: 20,
    bold: false, italic: false, underline: false, color: '#000000',
    glyphs: [glyph(0, 2040, 10), glyph(1, 2050, 10.3798828125), glyph(2, 2060.38, 10)],
  } as PositionedTextRun;
  expect(fittedPieceDecisions(run, (text) => text.length * 8)).toBeNull();
  run.glyphs[1].x = 2051;
  run.glyphs[2].x = 2061.3798828125;
  expect(fittedPieceDecisions(run, (text) => text.length * 8)?.map(({ start }) => start)).toEqual([0, 1]);
});

test('a glyph inside a cluster never departs, and rounding at large x is not a departure', () => {
  const glyph = (cluster: number, x: number, advance: number) => ({ glyphId: cluster + 1, cluster, x, advance, xOffset: 0, yOffset: 0 });
  const run = {
    text: 'ab', start: 0, end: 2, x: Math.fround(1039.4), width: Math.fround(24.78), fontId: 1, fontFamily: 'Corbel', fontSizePx: 20,
    bold: false, italic: false, underline: false, color: '#000000',
    glyphs: [glyph(0, Math.fround(1039.4), Math.fround(8.39)), glyph(0, Math.fround(1047.7902), Math.fround(8.39)), glyph(1, Math.fround(1056.1801), Math.fround(8.39))],
  } as PositionedTextRun;
  expect(fittedPieceDecisions(run, (text) => text.length * 8)).toBeNull();
});

test('a run whose clusters run backwards is not fitted', () => {
  const glyph = (cluster: number, x: number) => ({ glyphId: cluster + 1, cluster, x, advance: 8, xOffset: 0, yOffset: 0 });
  const run = {
    text: '\u05d0\u05d1\u05d2', start: 0, end: 3, x: 40, width: 27, fontId: 1, fontFamily: 'Corbel', fontSizePx: 20,
    bold: false, italic: false, underline: false, color: '#000000', glyphs: [glyph(2, 40), glyph(1, 49), glyph(0, 58)],
  } as PositionedTextRun;
  expect(fittedPieceDecisions(run, (text) => text.length * 8)).toBeNull();
});
