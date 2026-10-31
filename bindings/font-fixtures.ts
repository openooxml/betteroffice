import { expect } from 'bun:test';
import { readFileSync } from 'node:fs';

export interface DecodedFontFace {
  font: Uint8Array;
  payload?: Uint8Array;
}

export interface FontCase {
  bytes: Uint8Array;
  codec: string;
  dict?: Uint8Array;
  maxOutput?: number;
  font?: Uint8Array;
  payload?: Uint8Array;
}

function fixture(name: string): Uint8Array {
  return new Uint8Array(readFileSync(new URL(
    `../crates/betteroffice-fonts/tests/fixtures/${name}`, import.meta.url,
  )));
}

export function fontCases(): FontCase[] {
  const regular = fixture('Synthetic-Regular.ttf');
  const regularPayload = fixture('Synthetic-Regular.uw2');
  const boldPayload = fixture('Synthetic-Bold.uw2');
  const cases: FontCase[] = [
    { bytes: fixture('Synthetic-Regular.sfnt.zst'), codec: 'zstd', font: regular },
    {
      bytes: fixture('Synthetic-Bold.sfnt.zst'), codec: 'zstd', dict: regular,
      font: fixture('Synthetic-Bold.ttf'),
    },
    { bytes: fixture('Synthetic-Regular.ttf.zst'), codec: 'woff2t+zstd', payload: regularPayload },
    {
      bytes: fixture('Synthetic-Bold.ttf.zst'), codec: 'woff2t+zstd', dict: regularPayload,
      payload: boldPayload,
    },
    { bytes: fixture('Synthetic-Bold.nodict.zst'), codec: 'woff2t+zstd', payload: boldPayload },
  ];
  for (const codec of ['zstd', 'woff2t+zstd']) {
    for (const bytes of [regular, fixture('Synthetic-Regular.woff2'), new Uint8Array([0x28, 0xb5])]) {
      cases.push({ bytes, codec, maxOutput: 1.5, font: bytes });
    }
  }
  return cases;
}

export function expectFontFace(decoded: DecodedFontFace, expected: FontCase): void {
  expect(Object.getPrototypeOf(decoded)).toBe(Object.prototype);
  if (expected.font) {
    expect(decoded.font).toEqual(expected.font);
  } else {
    expect(decoded.font.subarray(0, 4)).toEqual(new Uint8Array([0, 1, 0, 0]));
  }
  expect(decoded.payload).toEqual(expected.payload);
}

export function fontErrorCases(): (FontCase & { message: string })[] {
  const bytes = fixture('Synthetic-Regular.sfnt.zst');
  const corrupt = bytes.slice();
  corrupt[corrupt.length - 1] ^= 1;
  const cases: (FontCase & { message: string })[] = [
    { bytes: new Uint8Array([0, 1]), codec: 'brotli', maxOutput: 1.5, message: 'unsupported font face codec: brotli' },
    { bytes, codec: 'brotli', maxOutput: 1.5, message: 'unsupported font face codec: brotli' },
    { bytes: corrupt, codec: 'zstd', message: 'font face checksum failed: corrupt face or wrong dictionary' },
    { bytes: new Uint8Array([0x28, 0xb5, 0x2f, 0xfd]), codec: 'zstd', message: 'font face is truncated' },
  ];
  for (const codec of ['zstd', 'woff2t+zstd']) {
    for (const maxOutput of [NaN, Infinity, -1, 1.5, 0x80000000]) {
      cases.push({
        bytes, codec, maxOutput,
        message: 'font face maxOutput must be an integer from 0 to 2147483647',
      });
    }
  }
  return cases;
}
