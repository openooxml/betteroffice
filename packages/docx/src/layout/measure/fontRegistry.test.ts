import { expect, test } from 'bun:test';
import { createFontProvider, resolveLastResortFace } from '@betteroffice/fonts';
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

for (const family of ['Segoe UI Light', 'Segoe UI Semilight', 'Segoe UI', 'Segoe UI Semibold']) {
  for (const bold of [false, true]) {
    test(`${family} ${bold} italic measures and draws the true italic last resort, not Selawik`, async () => {
      const provider = createFontProvider();
      expect(provider.resolve(family, bold, true)).toBeUndefined();
      const registered: Uint8Array[] = [];
      const registry = new TextMeasureFontRegistry(
        { registerFont: bytes => { registered.push(bytes); return registered.length; } },
        { bundled: provider }
      );
      expect(await registry.getFontIdChain(family, bold, true)).toEqual([1]);
      expect(registered).toEqual([new Uint8Array(await provider.resolveLastResort(family, bold, true, 'word')())]);
      expect(resolveLastResortFace(family, bold, true, 'word').style).toBe('italic');
    });
    for (const failPrimary of [false, true]) {
      test(`${family} ${bold} upright uses ${failPrimary ? 'Liberation measurement views' : 'Selawik with Segoe UI measurement views'}`, async () => {
        const provider = createFontProvider();
        const primary = provider.resolve(family, bold, false)!;
        const registered: Uint8Array[] = [];
        const substitutes: Array<[number, string]> = [];
        const registry = new TextMeasureFontRegistry(
          {
            registerFont: bytes => { registered.push(bytes); return registered.length; },
            registerSubstituteFont: (id, requested) => {
              substitutes.push([id, requested]);
              return id + 100;
            },
          },
          {
            bundled: {
              ...provider,
              resolve: () => failPrimary ? async () => { throw new Error('primary unavailable'); } : primary,
            },
          }
        );
        const chain = await registry.getFontIdChain(family, bold, false);
        const lastResort = new Uint8Array(await provider.resolveLastResort(family, bold, false, 'word')());
        if (failPrimary) {
          expect(chain).toEqual([101]);
          expect(registered[0]).toEqual(lastResort);
          expect(substitutes).toContainEqual([1, family]);
        } else {
          expect(chain).toEqual([101, 102]);
          expect(registered[0]).toEqual(new Uint8Array(await primary()));
          expect(registered[1]).toEqual(lastResort);
          expect(substitutes).toContainEqual([1, 'Segoe UI']);
          if (family !== 'Segoe UI') expect(substitutes).not.toContainEqual([1, family]);
          expect(substitutes).toContainEqual([2, family]);
        }
      });
    }
  }
}

for (const failPrimary of [false, true]) {
  test(`Arial ${failPrimary ? 'falls back to another base face' : 'loads only its own base face'}`, async () => {
    const provider = createFontProvider();
    const registered: Uint8Array[] = [];
    const registry = new TextMeasureFontRegistry(
      { registerFont: bytes => { registered.push(bytes); return registered.length; } },
      {
        bundled: {
          ...provider,
          resolve: (family, bold, italic) => failPrimary
            ? async () => { throw new Error('primary unavailable'); }
            : provider.resolve(family, bold, italic),
        },
      }
    );
    const chain = await registry.getFontIdChain('Arial', false, false);
    const primary = new Uint8Array(await provider.resolve('Arial', false, false)!());
    const alternate = new Uint8Array(await provider.resolveLastResort('Arial', false, false, 'word')());
    expect(alternate).not.toEqual(primary);
    expect(chain).toEqual([1]);
    expect(registered).toEqual([failPrimary ? alternate : primary]);
  });
}

test('Lucida Bright keeps its substitute measurement view', async () => {
  const provider = createFontProvider();
  const substitutes: Array<[number, string]> = [];
  const registry = new TextMeasureFontRegistry(
    {
      registerFont: () => 1,
      registerSubstituteFont: (id, family) => { substitutes.push([id, family]); return 2; },
    },
    { bundled: provider }
  );
  expect(await registry.getFontIdChain('Lucida Bright', false, false)).toEqual([2]);
  expect(substitutes).toEqual([[1, 'Lucida Bright']]);
});

test('a rejected primary still measures with real last-resort bytes', async () => {
  const wasm = await import('../wasm');
  await wasm.preloadLayoutWasm();
  const provider = createFontProvider();
  const registered: Uint8Array[] = [];
  const registry = new TextMeasureFontRegistry(
    {
      registerFont: bytes => { registered.push(bytes); return wasm.registerMeasureFont(bytes); },
      registerSubstituteFont: wasm.registerSubstituteMeasureFont,
    },
    {
      bundled: {
        ...provider,
        resolve: () => async () => { throw new Error('primary unavailable'); },
      },
    }
  );
  try {
    const chain = await registry.getFontIdChain('Segoe UI Light', false, false);
    expect(chain.length).toBeGreaterThan(0);
    expect(registered).toEqual([new Uint8Array(await provider.resolveLastResort('Segoe UI Light', false, false, 'word')())]);
    const measured = JSON.parse(wasm.measureParagraphJson(JSON.stringify({
      block: { kind: 'paragraph', runs: [{ kind: 'text', text: 'Real font metrics' }] },
      maxWidth: 200,
      fontChains: { 'segoe ui light|0|0': chain },
      defaults: { fontSize: 12, fontFamily: 'Segoe UI Light' },
      authoritativeShaping: true,
    })));
    expect(measured.kind).toBe('paragraph');
    expect(measured.lines).toHaveLength(1);
    expect(measured.totalHeight).toBeGreaterThan(0);
  } finally {
    wasm.clearMeasureFonts();
  }
});

for (const family of ['Segoe UI Light', 'Segoe UI Semilight', 'Segoe UI']) {
  test(`${family} uses Word line metrics and unscaled Selawik advances with the real engine`, async () => {
    const wasm = await import('../wasm');
    await wasm.preloadLayoutWasm();
    const provider = createFontProvider();
    const registry = new TextMeasureFontRegistry(
      {
        registerFont: wasm.registerMeasureFont,
        registerSubstituteFont: wasm.registerSubstituteMeasureFont,
      },
      { bundled: provider }
    );
    try {
      const chain = await registry.getFontIdChain(family, false, false);
      const raw = wasm.registerMeasureFont(new Uint8Array(await provider.resolve(family, false, false)!()));
      const measure = (ids: number[]) => JSON.parse(wasm.measureParagraphJson(JSON.stringify({
        block: { kind: 'paragraph', runs: [{ kind: 'text', text: 'Real font metrics' }] },
        maxWidth: 200,
        fontChains: { [`${family.toLowerCase()}|0|0`]: ids },
        defaults: { fontSize: 12, fontFamily: family },
        authoritativeShaping: true,
      })));
      const measured = measure(chain);
      const unscaled = measure([raw]);
      expect(measured.kind).toBe('paragraph');
      expect(measured.lines).toHaveLength(1);
      expect(unscaled.lines).toHaveLength(1);
      expect(measured.lines[0].ascent).toBeCloseTo(2210 / 2048 * 16, 5);
      expect(measured.lines[0].descent).toBeCloseTo(514 / 2048 * 16, 5);
      expect(measured.lines[0].lineHeight).toBeCloseTo((2210 + 514) / 2048 * 16, 5);
      expect(unscaled.lines[0].width).toBeGreaterThan(0);
      expect(measured.lines[0].width / unscaled.lines[0].width).toBeCloseTo(1.0, 5);
    } finally {
      wasm.clearMeasureFonts();
    }
  });
}

test('Selawik Word chains retain Liberation Sans coverage and CJK without missing glyphs', async () => {
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
