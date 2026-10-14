import { expect } from 'bun:test';
import JSZip from 'jszip';

export async function expectSavedPartsEqual(native: Uint8Array, wasm: Uint8Array) {
  if (Buffer.from(native).equals(Buffer.from(wasm))) {
    console.log('Saved bytes match.');
    return;
  }
  const left = await JSZip.loadAsync(native, { checkCRC32: true });
  const right = await JSZip.loadAsync(wasm, { checkCRC32: true });
  const parts = (zip: JSZip) => Object.keys(zip.files).filter((name) => !zip.files[name].dir).sort();
  expect(parts(left)).toEqual(parts(right));
  for (const name of parts(left)) {
    const nativePart = await left.files[name].async('uint8array');
    const wasmPart = await right.files[name].async('uint8array');
    if (name.endsWith('.xml') || name.endsWith('.rels')) {
      expect(Buffer.from(nativePart).toString('utf8'), name).toBe(Buffer.from(wasmPart).toString('utf8'));
    }
    expect(Buffer.from(nativePart), name).toEqual(Buffer.from(wasmPart));
  }
  console.log('Saved ZIP containers differ; every part byte matches.');
}
