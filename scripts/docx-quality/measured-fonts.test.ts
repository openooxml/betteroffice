import { expect, test } from 'bun:test';
import { fontProvider } from '../../packages/fonts/src/provider';
import { measuredFontProvider, type FontLoad } from './measured-fonts';

test('the measured provider forwards and records glyph coverage loads', async () => {
  const provider = fontProvider(async (face) => {
    if (face.file === 'Carlito-Bold.ttf') throw new Error('offline');
    return new ArrayBuffer(face.byteLength);
  });
  const fontLoads: FontLoad[] = [];
  const measured = measuredFontProvider(provider, fontLoads);
  expect(measured.resolve('Segoe UI', false, false)!.measureAs).toBe('Segoe UI');
  expect(measured.resolveGlyphFallback).toBeDefined();
  expect(measured.resolveGlyphFallback!('Google Sans', false, false)).toBeUndefined();
  const coverage = measured.resolveGlyphFallback!('Segoe UI Black', true, false, 'word')!;
  expect((await coverage()).byteLength).toBe(
    (await provider.resolveGlyphFallback!('Segoe UI Black', true, false, 'word')!()).byteLength,
  );
  await expect(measured.resolveGlyphFallback!('Segoe UI Black', true, false, 'powerpoint')!()).rejects.toThrow('offline');
  expect(fontLoads).toEqual([
    { kind: 'glyph', family: 'Segoe UI Black', ok: true },
    { kind: 'glyph', family: 'Segoe UI Black', ok: false, error: 'Error: offline' },
  ]);
});

test('the measured provider leaves glyph coverage out when the source has none', () => {
  const { resolveGlyphFallback: _omitted, ...source } = fontProvider(async () => new ArrayBuffer(0));
  expect('resolveGlyphFallback' in measuredFontProvider(source, [])).toBe(false);
});
