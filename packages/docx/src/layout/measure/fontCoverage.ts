export type FontCoverage = (character: string) => boolean;

const BASE_LATIN_RANGES = [[0x20, 0x24f], [0x370, 0x4ff], [0x2000, 0x206f],
  [0x20a0, 0x20cf], [0x2190, 0x22ff], [0x25a0, 0x27bf], [0xfb00, 0xfb06]] as const;
const BASE_LATIN_ALIASES = '\uf06f\uf076\uf0a7\uf0a8\uf0b7\uf0d8\uf0fc';

function baseLatinCharacter(index: number): string | undefined {
  for (const [start, end] of BASE_LATIN_RANGES) {
    if (index <= end - start) return String.fromCodePoint(start + index);
    index -= end - start + 1;
  }
  return BASE_LATIN_ALIASES[index];
}

export function charactersFromBits(bits = ''): string[] {
  const characters: string[] = [];
  for (let word = 0; word * 8 < bits.length; word++) {
    const value = Number.parseInt(bits.slice(word * 8, word * 8 + 8), 16);
    for (let bit = 0; bit < 32; bit++) {
      const character = (value & (1 << bit)) !== 0 ? baseLatinCharacter(word * 32 + bit) : undefined;
      if (character) characters.push(character);
    }
  }
  return characters;
}

function baseLatinIndex(point: number): number | undefined {
  let offset = 0;
  for (const [start, end] of BASE_LATIN_RANGES) {
    if (point >= start && point <= end) return offset + point - start;
    offset += end - start + 1;
  }
  const alias = point >= 0 && point <= 0xffff ? BASE_LATIN_ALIASES.indexOf(String.fromCharCode(point)) : -1;
  return alias < 0 ? undefined : offset + alias;
}

/** Coverage answering from base Latin bit set words. */
export function wordsCoverage(words: Uint32Array): FontCoverage {
  return (character) => {
    const index = baseLatinIndex(character.codePointAt(0) ?? -1);
    return index !== undefined && (words[index >>> 5]! & (1 << (index & 31))) !== 0;
  };
}

/** Base Latin coverage from coverage bits (hex words). */
export function coverageFromBits(bits: string): FontCoverage {
  const words = new Uint32Array(64);
  for (let word = 0; word < 64 && word * 8 < bits.length; word++) {
    words[word] = Number.parseInt(bits.slice(word * 8, word * 8 + 8), 16);
  }
  return wordsCoverage(words);
}

export function baseLatinAlias(character: string): boolean {
  return /^[\uf0b7\uf0a7\uf06f\uf0d8\uf0fc\uf076\uf0a8]$/u.test(character);
}

export function symbolFamily(family: string): boolean {
  return /^(symbol|wingdings(?: [23])?)$/i.test(family.trim());
}

export function baseLatinCandidate(character: string, styled: boolean): boolean {
  const point = character.codePointAt(0) ?? 0;
  return (point >= 0x20 && point <= (styled ? 0x17f : 0x24f)) ||
    (point >= 0x2000 && point <= 0x206f) || (point >= 0x20a0 && point <= 0x20cf) ||
    (!styled && ((point >= 0x370 && point <= 0x4ff) ||
      (point >= 0x2190 && point <= 0x22ff) || (point >= 0x25a0 && point <= 0x27bf) ||
      (point >= 0xfb00 && point <= 0xfb06) || baseLatinAlias(character)));
}
