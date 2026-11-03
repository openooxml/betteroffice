// usage: bun scripts/base-latin/coverage.ts [--check]
// Writes packages/fonts/src/coverage.json: nested tiers of base Latin characters
// shared by the bundled faces, and per face the largest tier its cmap covers.
import { readFileSync, writeFileSync } from 'node:fs';
import { fileURLToPath } from 'node:url';
import { BASE_LATIN_FONTS, BUNDLED_FONTS } from '../../packages/fonts/src/manifest';
import { coverageFaceFiles } from '../../packages/fonts/src/coverage';
import { loadBundledFontBytes } from '../../packages/fonts/src/index';
import { fontAssetResponse } from '../../packages/fonts/src/test-assets';
import { coverageBits, fontCoverage } from '../../packages/docx/src/layout/measure/cmapCoverage';
import { charactersFromBits } from '../../packages/docx/src/layout/measure/fontCoverage';

const OUTPUT = fileURLToPath(new URL('../../packages/fonts/src/coverage.json', import.meta.url));
const THRESHOLDS = [1000, 800, 600, 400, 250, 150];
const ALL = charactersFromBits('f'.repeat(512));

export async function generate(): Promise<string> {
  const fetch = globalThis.fetch;
  globalThis.fetch = (async (input: RequestInfo | URL) => fontAssetResponse(input)) as typeof fetch;
  try {
    const faces = new Map([...BUNDLED_FONTS, ...BASE_LATIN_FONTS].map((face) => [face.file, face]));
    const sets: Set<string>[] = [];
    for (const file of coverageFaceFiles()) {
      const bytes = await loadBundledFontBytes(faces.get(file)!, { baseUrl: 'https://assets.invalid/fonts/' });
      const covers = fontCoverage(new Uint8Array(bytes));
      if (!covers) throw new Error(`${file}: no readable cmap`);
      sets.push(new Set(ALL.filter((character) => covers(character))));
    }
    const tiers = THRESHOLDS.map((threshold) => {
      const large = sets.filter((set) => set.size >= threshold);
      return new Set(ALL.filter((character) => large.length > 0 && large.every((set) => set.has(character))));
    });
    const levels = sets.map((set) => {
      const index = tiers.findIndex((tier) => tier.size > 0 && [...tier].every((character) => set.has(character)));
      return index < 0 ? 0 : index + 1;
    });
    return JSON.stringify({
      thresholds: THRESHOLDS,
      tiers: tiers.map((tier) => coverageBits((character) => tier.has(character))),
      levels: levels.join(''),
    }) + '\n';
  } finally {
    globalThis.fetch = fetch;
  }
}

if (import.meta.main) {
  const next = await generate();
  if (process.argv.includes('--check')) {
    if (readFileSync(OUTPUT, 'utf8') !== next) {
      console.error('packages/fonts/src/coverage.json is stale: run bun scripts/base-latin/coverage.ts');
      process.exit(1);
    }
  } else writeFileSync(OUTPUT, next);
}
