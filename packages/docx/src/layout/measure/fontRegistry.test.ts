import { expect, test } from 'bun:test';
import { TextMeasureFontRegistry } from './fontRegistry';

test("an unknown family asks the provider for Word's last-resort face", async () => {
  const offices: unknown[] = [];
  const registry = new TextMeasureFontRegistry(
    { registerFont: () => 3 },
    {
      bundled: {
        resolve: () => undefined,
        resolveLastResort: (_family, _bold, _italic, office) => {
          offices.push(office);
          return async () => new ArrayBuffer(16);
        },
      },
    }
  );
  expect(await registry.getFontIdChain('Lato', false, false)).toEqual([3]);
  expect(offices).toEqual(['word']);
});

test('Selawik Word chains retain Liberation Sans coverage and CJK without missing glyphs', async () => {
  const { createFontProvider } = await import('@betteroffice/fonts');
  const provider = createFontProvider();
  const registered: Uint8Array[] = [];
  const offices: unknown[] = [];
  const registry = new TextMeasureFontRegistry(
    { registerFont: bytes => { registered.push(bytes); return registered.length; } },
    {
      bundled: {
        ...provider,
        resolveGlyphFallback: (family, bold, italic, office) => {
          offices.push(office);
          return provider.resolveGlyphFallback!(family, bold, italic, office);
        },
      },
    }
  );
  const familyIds = await registry.getFontIdChain('Segoe UI', false, false);
  expect(familyIds).toEqual([1, 2]);
  expect(offices).toEqual(['word']);
  expect(registered[0]).toEqual(new Uint8Array(await provider.resolve('Segoe UI', false, false)!()));
  expect(registered[1]).toEqual(new Uint8Array(await provider.resolve('Arial', false, false)!()));
  const chain = [...familyIds, ...await registry.getScriptFallbackIds(['cjk-sc'])];
  expect(chain).toEqual([1, 2, 3]);
  for (const character of 'Latin ἀ Ӂ 中') {
    expect(chain.some(id => covers(registered[id - 1], character))).toBe(true);
  }
  expect(covers(registered[0], 'L')).toBe(true);
  for (const character of ['ἀ', 'Ӂ']) {
    expect(covers(registered[0], character)).toBe(false);
    expect(covers(registered[1], character)).toBe(true);
  }
  expect(covers(registered[2], '中')).toBe(true);
});

function covers(bytes: Uint8Array, character: string): boolean {
  const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
  const code = character.codePointAt(0)!;
  for (let table = 12; table < 12 + view.getUint16(4) * 16; table += 16) {
    if (view.getUint32(table) !== 0x636d6170) continue;
    const cmap = view.getUint32(table + 8);
    for (let record = cmap + 4; record < cmap + 4 + view.getUint16(cmap + 2) * 8; record += 8) {
      const platform = view.getUint16(record);
      const encoding = view.getUint16(record + 2);
      if (platform !== 0 && !(platform === 3 && (encoding === 1 || encoding === 10))) continue;
      const subtable = cmap + view.getUint32(record + 4);
      const format = view.getUint16(subtable);
      if (format === 12) {
        const groups = view.getUint32(subtable + 12);
        for (let group = subtable + 16; group < subtable + 16 + groups * 12; group += 12) {
          const start = view.getUint32(group);
          if (code >= start && code <= view.getUint32(group + 4) &&
              view.getUint32(group + 8) + code - start !== 0) return true;
        }
      } else if (format === 4 && code <= 0xffff) {
        const segments = view.getUint16(subtable + 6) / 2;
        const ends = subtable + 14;
        const starts = ends + segments * 2 + 2;
        const deltas = starts + segments * 2;
        const offsets = deltas + segments * 2;
        for (let segment = 0; segment < segments; segment++) {
          const start = view.getUint16(starts + segment * 2);
          if (code < start || code > view.getUint16(ends + segment * 2)) continue;
          const delta = view.getInt16(deltas + segment * 2);
          const address = offsets + segment * 2;
          const offset = view.getUint16(address);
          const glyph = offset ? view.getUint16(address + offset + (code - start) * 2) : code;
          if ((!offset || glyph !== 0) && ((glyph + delta) & 0xffff) !== 0) return true;
        }
      }
    }
  }
  return false;
}
