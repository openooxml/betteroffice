import { afterEach, expect, mock, spyOn, test } from 'bun:test';
import { readFileSync } from 'node:fs';
import * as fontManager from '@betteroffice/fonts';
import { configureDefaultFonts } from '@betteroffice/docx/layout';
import { customWorkerFontFaces, fontEnvironmentForRegistrations, prepareWorkerFontEnvironment, uncoveredFontRequirements } from './workerFonts';

afterEach(() => { mockFetch?.mockRestore(); configureDefaultFonts({}); });
let mockFetch: ReturnType<typeof spyOn> | undefined;

test('registered base-face descriptors survive worker font environments', () => {
  const face = { family: 'BetterOffice Base Latin', bold: false, italic: false };
  const environment = fontEnvironmentForRegistrations([Uint8Array.of(1), Uint8Array.of(2)],
    { 'requested|1|1': [0, 1] }, [{ key: 'requested|1|1', family: 'Requested', bold: true, italic: true }],
    ['Requested'], new Map([[1, face]]));
  expect(environment.faces[0]).toMatchObject({ family: 'Requested', weight: 'bold', style: 'italic' });
  expect(environment.faces[1]).toMatchObject({ family: face.family, weight: 'normal', style: 'normal' });
  expect(environment.chains).toEqual({ 'requested|1|1': [0, 1] });
});

test('new characters refresh an existing family and retain earlier coverage requirements', () => {
  const requirement = { key: 'requested|0|0', family: 'Requested', bold: false, italic: false };
  const prepared = [{ ...requirement, characters: ['A'] }];
  expect(uncoveredFontRequirements(prepared, [{ ...requirement, characters: ['A'] }])).toBeNull();
  expect(uncoveredFontRequirements(prepared, [{ ...requirement, characters: ['Ω'] }]))
    .toEqual([{ ...requirement, characters: ['A', 'Ω'] }]);
});

test('new characters the registered fonts already cover do not rebuild the font environment', () => {
  const requirement = { key: 'requested|0|0', family: 'Requested', bold: false, italic: false };
  const prepared = [{ ...requirement, characters: ['A'] }];
  const typed = [{ ...requirement, characters: ['B', 'Ω'] }];
  const asked: unknown[] = [];
  expect(uncoveredFontRequirements(prepared, typed, (requirements) => { asked.push(requirements); return true; })).toBeNull();
  expect(asked).toEqual([[{ ...requirement, characters: ['A', 'B', 'Ω'] }]]);
  expect(uncoveredFontRequirements(prepared, typed, () => false)).toEqual([{ ...requirement, characters: ['A', 'B', 'Ω'] }]);
  const scripted = [{ ...requirement, characters: ['A'], scripts: ['arabic' as const] }];
  expect(uncoveredFontRequirements(prepared, scripted, () => true)).toEqual(scripted);
  const added = [...typed, { key: 'other|0|0', family: 'Other', bold: false, italic: false }];
  expect(uncoveredFontRequirements(prepared, added, () => true)?.map(item => item.key)).toEqual(['requested|0|0', 'other|0|0']);
});

test('the coverage check judges only the requirements that gained characters', () => {
  const typed = { key: 'typed|0|0', family: 'Typed', bold: false, italic: false };
  const other = { key: 'other|0|0', family: 'Other', bold: false, italic: false, characters: ['A'] };
  const asked: unknown[] = [];
  const covers = (requirements: unknown[]) => { asked.push(requirements); return requirements.length === 1; };
  expect(uncoveredFontRequirements([{ ...typed, characters: ['A'] }, other], [{ ...typed, characters: ['j'] }, other], covers)).toBeNull();
  expect(asked).toEqual([[{ ...typed, characters: ['A', 'j'] }]]);
});

test('compact requirement refreshes compare decoded candidates and settle', () => {
  const requirement = { key: 'requested|0|0', family: 'Requested', bold: false, italic: false };
  const prepared = [{ ...requirement, characterBits: '0000000000000002' }];
  expect(uncoveredFontRequirements(prepared, [{ ...requirement, characters: ['A'] }])).toBeNull();
  const next = [{ ...requirement, characters: ['B'] }];
  const refreshed = uncoveredFontRequirements(prepared, next)!;
  expect(refreshed[0]?.characters).toEqual(['A', 'B']);
  expect(uncoveredFontRequirements(refreshed, next)).toBeNull();
});

test('custom WOFF and WOFF2 bytes pass unchanged to worker preparation', async () => {
  for (const signature of [0x774f4646, 0x774f4632]) {
    const data = new ArrayBuffer(16);
    new DataView(data).setUint32(0, signature);
    mockFetch = spyOn(globalThis, 'fetch').mockResolvedValue(new Response(data));
    const faces = await customWorkerFontFaces([{ family: 'Custom', src: '/face', weight: 700 }], new AbortController().signal);
    expect(faces).toMatchObject([{ family: 'Custom', weight: 'bold', data }]);
    mockFetch.mockRestore();
  }
});

test('captured registrations become face slots with style chains and substitute identities', () => {
  const environment = fontEnvironmentForRegistrations([Uint8Array.of(1), { family: 'Custom', substituteOf: 0 }],
    { 'custom|1|1': [0, 1] }, [{ key: 'custom|1|1', family: 'Custom', bold: true, italic: true }], ['Custom']);
  expect(environment).toEqual({ faces: [{ family: 'Custom', data: Uint8Array.of(1), weight: 'bold', style: 'italic' },
    { family: 'Custom', substituteOf: 0 }], chains: { 'custom|1|1': [0, 1] }, families: ['Custom'] });
});

for (const embedded of [false, true]) {
  test(`worker measurement keeps URL faces ${embedded ? 'after embeds' : 'before explicit manager defaults'}`, async () => {
    const data = new Uint8Array([1, 2]).buffer;
    const custom = new Uint8Array([3, 4]).buffer;
    const bundled = new Uint8Array([5, 6]).buffer;
    const load = async () => bundled;
    const create = spyOn(fontManager, 'createFontProvider').mockReturnValue({
      resolve: () => load, resolveFamily: () => load,
      resolveScriptFallback: () => undefined, resolveLastResort: () => load,
    });
    mockFetch = spyOn(globalThis, 'fetch').mockResolvedValue(new Response(custom));
    try {
      const environment = await prepareWorkerFontEnvironment({
        fonts: [{ family: 'Calibri', src: '/fonts/calibri.ttf' }],
        embedded: embedded ? [{ family: 'Calibri', data, weight: 'normal', style: 'normal', subsetted: false }] : [],
        requirements: [{ key: 'calibri|0|0', family: 'Calibri', bold: false, italic: false }],
        families: ['Calibri'], provider: fontManager.createFontProvider(), compat: undefined, signal: new AbortController().signal,
      });
      expect(environment.faces[0]).toMatchObject({ family: 'Calibri', data: new Uint8Array(embedded ? data : custom) });
      expect(environment.faces.filter(face => 'data' in face).map(face => face.data)).toEqual([
        ...(embedded ? [new Uint8Array(data)] : []), new Uint8Array(custom), new Uint8Array(bundled),
      ]);
      expect(environment.chains['calibri|0|0']).toEqual(embedded ? [0, 1, 3, 4] : [0, 2, 3]);
      expect(create).toHaveBeenCalledTimes(1);
      expect(mockFetch).toHaveBeenCalledTimes(1);
    } finally { create.mockRestore(); }
  });
}

test('a hosted subset reaches base Latin after embedded and app faces with the implicit provider', async () => {
  const subset = new Uint8Array(readFileSync(new URL('../../../../../fonts/assets/NotoSansHebrew-Regular.ttf', import.meta.url)));
  const base = new Uint8Array(readFileSync(new URL('../../../../../../crates/base-latin/fonts/BetterOfficeBaseLatin-Regular.ttf', import.meta.url)));
  const load = spyOn(fontManager, 'createFontProvider').mockReturnValue({
    resolve: () => { throw new Error('hosted substitute requested'); },
    resolveFamily: () => { throw new Error('hosted substitute requested'); },
    resolveLastResort: () => { throw new Error('hosted substitute requested'); },
    resolveScriptFallback: () => undefined,
    resolveBaseLatin: () => async () => base.buffer,
  });
  mockFetch = spyOn(globalThis, 'fetch').mockResolvedValue(new Response(subset));
  try {
    const environment = await prepareWorkerFontEnvironment({
      fonts: [{ family: 'Requested', src: '/subset.ttf' }], embedded: [{ family: 'Requested', data: subset.slice().buffer,
        weight: 'normal', style: 'normal', subsetted: true }],
      requirements: [{ key: 'requested|0|0', family: 'Requested', bold: false, italic: false, characters: ['Ω'] }],
      families: ['Requested'], provider: undefined, compat: undefined, signal: new AbortController().signal,
    });
    expect(environment.chains['requested|0|0']).toEqual([0, 1, 2]);
    expect(environment.faces.filter(face => 'data' in face).map(face => face.data)).toEqual([subset, subset, base]);
  } finally { load.mockRestore(); }
});

test('hosted WOFF and WOFF2 keep their compressed faces and chains for Greek text', async () => {
  const load = mock(async () => { throw new Error('unexpected base Latin load'); });
  const create = spyOn(fontManager, 'createFontProvider').mockReturnValue({
    resolve: () => { throw new Error('hosted substitute requested'); },
    resolveFamily: () => { throw new Error('hosted substitute requested'); },
    resolveLastResort: () => { throw new Error('hosted substitute requested'); },
    resolveScriptFallback: () => undefined,
    resolveBaseLatin: () => load,
  });
  try {
    for (const signature of [0x774f4646, 0x774f4632]) {
      const data = new Uint8Array(16);
      new DataView(data.buffer).setUint32(0, signature);
      mockFetch = spyOn(globalThis, 'fetch').mockResolvedValue(new Response(data));
      const environment = await prepareWorkerFontEnvironment({
        fonts: [{ family: 'Requested', src: '/face' }], embedded: [],
        requirements: [{ key: 'requested|0|0', family: 'Requested', bold: false, italic: false, characters: ['Ω'] }],
        families: ['Requested'], provider: undefined, compat: undefined, signal: new AbortController().signal,
      });
      expect(environment.chains['requested|0|0']).toEqual([0]);
      expect(environment.faces).toEqual([{ family: 'Requested', data, weight: 'normal', style: 'normal' }]);
      expect(mockFetch).toHaveBeenCalledTimes(1);
      mockFetch.mockRestore();
    }
    expect(load).not.toHaveBeenCalled();
  } finally { create.mockRestore(); }
});

test('hosted Calibri excludes implicit worker substitutes while unhosted Cambria keeps them', async () => {
  const custom = Uint8Array.of(3, 4);
  const bundled = Uint8Array.of(5, 6);
  const load = async () => bundled.buffer;
  const create = spyOn(fontManager, 'createFontProvider').mockReturnValue({
    resolve: () => load, resolveFamily: () => load,
    resolveScriptFallback: () => undefined, resolveLastResort: () => load,
  });
  mockFetch = spyOn(globalThis, 'fetch').mockResolvedValue(new Response(custom));
  try {
    const environment = await prepareWorkerFontEnvironment({
      fonts: [{ family: 'Calibri', src: '/calibri.woff2' }], embedded: [],
      requirements: ['Calibri', 'Cambria'].map(family => ({ key: `${family.toLowerCase()}|0|0`, family, bold: false, italic: false })),
      families: ['Calibri', 'Cambria'], provider: undefined, compat: undefined, signal: new AbortController().signal,
    });
    expect(environment.chains['calibri|0|0']).toEqual([0]);
    expect(environment.chains['cambria|0|0']).toHaveLength(2);
    expect(environment.faces.filter(face => 'data' in face).map(face => face.data)).toEqual([custom, bundled]);
  } finally { create.mockRestore(); }
});

test('an unknown worker family gets no implicit substitute', async () => {
  const custom = Uint8Array.of(3, 4);
  const load = async () => Uint8Array.of(5, 6).buffer;
  const create = spyOn(fontManager, 'createFontProvider').mockReturnValue({
    resolve: () => load, resolveFamily: () => load,
    resolveScriptFallback: () => undefined, resolveLastResort: () => load,
  });
  mockFetch = spyOn(globalThis, 'fetch').mockResolvedValue(new Response(custom));
  try {
    for (const hosted of [false, true]) {
      const environment = await prepareWorkerFontEnvironment({
        fonts: hosted ? [{ family: 'Custom', src: '/fonts/custom.ttf' }] : undefined,
        embedded: [], requirements: [{ key: 'custom|0|0', family: 'Custom', bold: false, italic: false }],
        families: ['Custom'], provider: undefined, compat: undefined, signal: new AbortController().signal,
      });
      expect(environment.faces).toEqual(hosted ? [{ family: 'Custom', data: custom, weight: 'normal', style: 'normal' }] : []);
      if (hosted) expect(environment.chains['custom|0|0']).toEqual([0]);
      else expect(environment.chains['custom|0|0']).toBeUndefined();
    }
  } finally { create.mockRestore(); }
});
