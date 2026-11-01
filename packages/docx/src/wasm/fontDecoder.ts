import { copyDecodedFace, fontBytes, type FontDecoder } from '../../../../shared/office-session/fontDecoder';

export const decodeFontFace: FontDecoder = async (bytes, { codec, dict, maxOutput }) => {
  const engine = await import('./edit');
  await engine.preloadEditWasm();
  return copyDecodedFace(engine.decodePackedFontFace(fontBytes(bytes), codec,
    dict === undefined ? undefined : fontBytes(dict), maxOutput));
};
