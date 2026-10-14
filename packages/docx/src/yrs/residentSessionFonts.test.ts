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
