import { expect, test } from 'bun:test';
import { upgradeExportFaces } from './exportFaces';

const skeleton = Uint8Array.of(0);
const whole = Uint8Array.of(1);

test('every case variant of a family gets the whole bytes', () => {
  const faces = [{ id: 1, family: 'MS Gothic', bytes: skeleton }, { id: 2, family: ' ms gothic ', bytes: skeleton }, { id: 3, family: 'Arial', bytes: skeleton }];
  upgradeExportFaces(faces, 'MS Gothic', whole);
  expect(faces.map((face) => face.bytes)).toEqual([whole, whole, skeleton]);
});

test('a host face keeps its bytes, and an older skeleton of the same style still gets the whole face', () => {
  const host = Uint8Array.of(2);
  const faces = [{ id: 1, family: 'Noto Sans KR', bytes: skeleton }, { id: 2, family: 'Noto Sans KR', bytes: host }];
  upgradeExportFaces(faces, 'Noto Sans KR', whole, (face) => face.id === 2);
  expect(faces.map((face) => face.bytes)).toEqual([whole, host]);
});

test('replacing an export face drops what an engine cached from the old bytes', async () => {
  const { readFile } = await import('node:fs/promises');
  const loader = await readFile(new URL('./loader.ts', import.meta.url), 'utf8');
  expect(/upgradeExportFaces\(exportFaces\.values\(\), family, bytes, keep\);\s+releasePng\(doc\);/.test(loader)).toBe(true);
});

