import { expect, test } from 'bun:test';
import { ResidentSessionFonts } from './residentFontEnvironment';

function setup() {
  const faces = new Set<FontFace>();
  let failing = false;
  const createFace = () => ({ load: async function () { if (failing) throw new Error('font load failed'); return this; } }) as FontFace;
  const engine = { decodeFont: (bytes: Uint8Array) => bytes.slice() };
  const fontSet = { add: (face: FontFace) => { faces.add(face); return fontSet as unknown as FontFaceSet; }, delete: (face: FontFace) => faces.delete(face) };
  const first = new ResidentSessionFonts(engine, fontSet, createFace);
  const second = new ResidentSessionFonts(engine, fontSet, createFace);
  const environment = { faces: [{ family: 'Shared', data: new Uint8Array(12) }], chains: { 'shared|0|0': [0] }, families: ['Shared'] };
  return { faces, first, second, environment, fail: () => { failing = true; } };
}

test('sessions remove only their own font faces on replacement and disposal', async () => {
  const s = setup();
  s.first.commit(await s.first.prepare(s.environment), () => {});
  s.second.commit(await s.second.prepare(s.environment), () => {});
  expect(s.faces.size).toBe(2);
  s.first.commit(await s.first.prepare(s.environment), () => {});
  expect(s.faces.size).toBe(2);
  s.first.dispose();
  expect(s.faces.size).toBe(1);
  s.second.dispose();
  expect(s.faces.size).toBe(0);
});

test('failed or disposed preparations preserve the active font environment', async () => {
  const s = setup();
  s.first.commit(await s.first.prepare(s.environment), () => {});
  s.fail();
  await expect(s.first.prepare(s.environment)).rejects.toThrow('font load failed');
  expect(s.first.snapshot?.revision).toBe(1);
  expect(s.faces.size).toBe(1);
  s.first.dispose();
  await expect(s.first.prepare(s.environment)).rejects.toThrow('disposed');
  expect(s.faces.size).toBe(0);
});

test('font publication failures discard prepared faces and retain the old faces', async () => {
  const s = setup();
  s.first.commit(await s.first.prepare(s.environment), () => {});
  const prepared = await s.first.prepare(s.environment);
  expect(() => s.first.commit(prepared, () => { throw new Error('registration failed'); })).toThrow('registration failed');
  expect(s.first.snapshot?.revision).toBe(1);
  expect(s.faces.size).toBe(1);
  s.first.dispose();
});

test('sliced faces reach registration with their files and get no CSS face; malformed slices are refused', async () => {
  const s = setup();
  const source = {
    ref: { format: 1, index: 'face/slices.json', indexSha256: 'a'.repeat(64), byteLength: 100, sha256: 'b'.repeat(64) },
    urls: { 'face/slices.json': 'https://fonts.example/face/slices.json' },
    timeoutMs: 30_000,
  };
  const environment = {
    faces: [{ family: 'Sliced', data: new Uint8Array(12) }, { family: 'Shared', data: new Uint8Array(12) }],
    chains: { 'sliced|0|0': [0], 'shared|0|0': [1] }, families: ['Sliced', 'Shared'],
    progressive: [{ slot: 0, source }],
  };
  let registered: unknown;
  s.first.commit(await s.first.prepare(environment), (_fonts, _chains, progressive) => { registered = progressive; });
  expect(registered).toEqual([{ slot: 0, source }]);
  const inlined = { ...source, urls: { ...source.urls, 'face/c001.zst': `data:application/octet-stream;base64,${'A'.repeat(20_000)}` } };
  await expect(s.first.prepare({ ...environment, progressive: [{ slot: 0, source: inlined }] })).resolves.toBeDefined();
  expect(s.faces.size).toBe(1);
  for (const progressive of [
    [{ slot: 2, source }],
    [{ slot: 0, source }, { slot: 0, source }],
    [{ slot: 0, source: { ...source, urls: { 'face/slices.json': 7 } } }],
    [{ slot: 0, source: { ...source, ref: { ...source.ref, sha256: 'not hex' } } }],
    [{ slot: 0, source: { ...source, urls: { 'face/slices.json': `https://fonts.example/${'a'.repeat(9000)}` } } }],
    [{ slot: 0, source: { ...source, urls: { 'face/c001.zst': `data:application/octet-stream;base64,${'A'.repeat(16 * 1024 * 1024)}` } } }],
  ]) {
    await expect(s.first.prepare({ ...environment, progressive } as typeof environment)).rejects.toThrow('Invalid worker font environment');
  }
  s.first.dispose();
});
