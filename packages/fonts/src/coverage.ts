import coverage from './coverage.json';
import { BASE_LATIN_FONTS, BUNDLED_FONTS } from './manifest';

/** @internal Every bundled face file once, in the order `coverage.json` lists their levels. */
export function coverageFaceFiles(): string[] {
  return [...new Set([...BUNDLED_FONTS, ...BASE_LATIN_FONTS].map((face) => face.file))];
}

let levels: Map<string, number> | undefined;

/** @internal Base Latin characters `file` is known to cover, as coverage bits. */
export function knownCoverage(file: string): string | undefined {
  levels ??= new Map(coverageFaceFiles().map((name, index) => [name, Number(coverage.levels[index] ?? 0)]));
  const level = levels.get(file) ?? 0;
  return level > 0 ? coverage.tiers[level - 1] : undefined;
}
