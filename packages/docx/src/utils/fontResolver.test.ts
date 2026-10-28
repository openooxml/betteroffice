import { expect, test } from 'bun:test';

import { resolveFontFamily } from './fontResolver';

for (const family of [
  'Microsoft Sans Serif',
  'MS Sans Serif',
  'MS Reference Sans Serif',
  'Unknown Sans-Serif',
  'Unknown SansSerif',
  'UNKNOWN SANS SERIF',
]) {
  test(`${family} uses the sans category fallback`, () => {
    expect(resolveFontFamily(family).cssFallback).toBe(
      `"${family}", Arial, Helvetica, sans-serif`
    );
    for (const serif of ['MS Serif', 'Times New Roman']) {
      expect(resolveFontFamily(serif).cssFallback).toContain('Times New Roman');
      expect(resolveFontFamily(serif).cssFallback).toEndWith(', serif');
    }
  });
}
