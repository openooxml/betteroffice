export type FontCoverage = (character: string) => boolean;

const BASE_LATIN_RANGES = [[0x20, 0x24f], [0x370, 0x4ff], [0x2000, 0x206f],
  [0x20a0, 0x20cf], [0x2190, 0x22ff], [0x25a0, 0x27bf], [0xfb00, 0xfb06]] as const;
const BASE_LATIN_CHARACTERS = [
  ...BASE_LATIN_RANGES.flatMap(([start, end]) => Array.from({ length: end - start + 1 }, (_, index) => String.fromCodePoint(start + index))),
  ...'\uf06f\uf076\uf0a7\uf0a8\uf0b7\uf0d8\uf0fc',
];

export function charactersFromBits(bits = ''): string[] {
  const characters: string[] = [];
  for (let word = 0; word * 8 < bits.length; word++) {
    const value = Number.parseInt(bits.slice(word * 8, word * 8 + 8), 16);
    for (let bit = 0; bit < 32; bit++) {
      if ((value & (1 << bit)) !== 0 && BASE_LATIN_CHARACTERS[word * 32 + bit]) {
        characters.push(BASE_LATIN_CHARACTERS[word * 32 + bit]!);
      }
    }
  }
  return characters;
}

export function coverageBits(covers: FontCoverage): string {
  const words = new Uint32Array(64);
  for (const [index, character] of BASE_LATIN_CHARACTERS.entries()) {
    if (covers(character)) words[index >>> 5]! |= 1 << (index & 31);
  }
  let length = words.length;
  while (length > 0 && words[length - 1] === 0) length--;
  return Array.from(words.subarray(0, length), word => word.toString(16).padStart(8, '0')).join('');
}

export function fontCoverage(bytes: Uint8Array): FontCoverage | undefined {
  let view: DataView;
  const tables: Array<(point: number) => boolean> = [];
  try {
    view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
    const directory = view.getUint32(0) === 0x74746366 && view.getUint32(8) > 0 ? view.getUint32(12) : 0;
    if (![0x00010000, 0x4f54544f, 0x74727565, 0x74797031].includes(view.getUint32(directory))) return undefined;
    for (let index = 0; index < view.getUint16(directory + 4); index++) {
      const record = directory + 12 + index * 16;
      if (view.getUint32(record) !== 0x636d6170) continue;
      const cmap = view.getUint32(record + 8);
      for (let entry = 0; entry < view.getUint16(cmap + 2); entry++) {
        const encoding = cmap + 4 + entry * 8;
        const platform = view.getUint16(encoding);
        const id = view.getUint16(encoding + 2);
        if (platform !== 0 && !(platform === 3 && (id === 0 || id === 1 || id === 10))) continue;
        const offset = cmap + view.getUint32(encoding + 4);
        const format = view.getUint16(offset);
        if (format === 12 || format === 13) {
          const groups = view.getUint32(offset + 12);
          if (offset + 16 + groups * 12 > view.byteLength) continue;
          tables.push((point) => {
            let left = 0;
            let right = groups;
            while (left < right) {
              const middle = (left + right) >>> 1;
              const group = offset + 16 + middle * 12;
              if (point < view.getUint32(group)) right = middle;
              else if (point > view.getUint32(group + 4)) left = middle + 1;
              else return view.getUint32(group + 8) + (format === 12 ? point - view.getUint32(group) : 0) !== 0;
            }
            return false;
          });
        } else if (format === 4) {
          const count = view.getUint16(offset + 6) / 2;
          const end = offset + 14;
          const start = end + count * 2 + 2;
          const delta = start + count * 2;
          const ranges = delta + count * 2;
          const limit = offset + view.getUint16(offset + 2);
          if (ranges + count * 2 > limit || limit > view.byteLength) continue;
          tables.push((point) => {
            if (point > 0xffff) return false;
            let left = 0;
            let right = count;
            while (left < right) {
              const middle = (left + right) >>> 1;
              if (view.getUint16(end + middle * 2) < point) left = middle + 1;
              else right = middle;
            }
            if (left === count || point < view.getUint16(start + left * 2)) return false;
            const adjustment = view.getInt16(delta + left * 2);
            const range = view.getUint16(ranges + left * 2);
            if (range === 0) return ((point + adjustment) & 0xffff) !== 0;
            const address = ranges + left * 2 + range + (point - view.getUint16(start + left * 2)) * 2;
            if (address + 2 > limit) return false;
            const glyph = view.getUint16(address);
            return glyph !== 0 && ((glyph + adjustment) & 0xffff) !== 0;
          });
        }
      }
    }
  } catch {
    return undefined;
  }
  if (tables.length === 0) return undefined;
  const covers = (point: number) => tables.some((table) => table(point));
  return (character) => {
    try {
      const point = character.codePointAt(0);
      return point !== undefined && (covers(point) || (point === 0x2011 && (covers(0x2010) || covers(0x2d))));
    } catch {
      return false;
    }
  };
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
