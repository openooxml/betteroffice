import { expect, test } from 'bun:test';
import { readFileSync } from 'node:fs';
import { BASE_LATIN_FONTS, BUNDLED_FONTS } from '../packages/fonts/src/manifest';
import { coverageFaceFiles, knownCoverage } from '../packages/fonts/src/coverage';
import { loadBundledFontBytes } from '../packages/fonts/src/index';
import { fontAssetResponse } from '../packages/fonts/src/test-assets';
import { fontCoverage } from '../packages/docx/src/layout/measure/cmapCoverage';
import { charactersFromBits } from '../packages/docx/src/layout/measure/fontCoverage';
import { generate } from './base-latin/coverage';

test('bundled face coverage tiers are regenerated from the fonts', async () => {
  const committed = readFileSync(new URL('../packages/fonts/src/coverage.json', import.meta.url), 'utf8');
  expect(await generate()).toBe(committed);
}, 300_000);

test('every character a bundled face is known to cover is in its cmap', async () => {
  const fetch = globalThis.fetch;
  globalThis.fetch = (async (input: RequestInfo | URL) => fontAssetResponse(input)) as typeof fetch;
  try {
    const faces = new Map([...BUNDLED_FONTS, ...BASE_LATIN_FONTS].map((face) => [face.file, face]));
    const wrong: string[] = [];
    let known = 0;
    for (const file of coverageFaceFiles()) {
      const bits = knownCoverage(file);
      if (bits === undefined) continue;
      known++;
      const bytes = await loadBundledFontBytes(faces.get(file)!, { baseUrl: 'https://assets.invalid/fonts/' });
      const covers = fontCoverage(new Uint8Array(bytes));
      expect(covers).toBeDefined();
      for (const character of charactersFromBits(bits)) {
        if (!covers!(character)) wrong.push(`${file} U+${character.codePointAt(0)!.toString(16)}`);
      }
    }
    expect(known).toBeGreaterThan(50);
    expect(wrong).toEqual([]);
  } finally {
    globalThis.fetch = fetch;
  }
}, 300_000);
