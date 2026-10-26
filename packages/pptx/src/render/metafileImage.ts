import { browserPresentationImageBlob } from './browserImage';

export function metafileImageBytes(bytes: Uint8Array, resolve?: () => Uint8Array | undefined): Uint8Array {
  if (!resolve || !isMetafile(bytes) || browserPresentationImageBlob(bytes).type === 'image/bmp') return bytes;
  return resolve() ?? new Uint8Array();
}

function isMetafile(bytes: Uint8Array): boolean {
  const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
  if (bytes.length >= 44 && view.getUint32(0, true) === 1 && view.getUint32(40, true) === 0x464d4520) return true;
  if (bytes.length < 18) return false;
  const start = view.getUint32(0, true) === 0x9ac6cdd7 ? 22 : 0;
  return bytes.length >= start + 18 && [1, 2].includes(view.getUint16(start, true)) &&
    view.getUint16(start + 2, true) === 9 && view.getUint16(start + 4, true) === 0x0300;
}
