import type { FontDecoder } from '../../../../shared/office-session/fontDecoder';
import { engineKind } from '../wasm/engine';

const OUTSIDE_ENGINE_BYTES = 512 * 1024;

/**
 * Decodes large pack files with `outside` rather than the engine: wasm memory never shrinks, so a large decode in the
 * engine costs it the decode's peak for the whole session. Falls back to the engine when `outside` is unavailable.
 */
export function decodeLargeFacesOutsideEngine(engine: FontDecoder, outside: () => Promise<FontDecoder>): FontDecoder {
  return async (bytes, options) => {
    if (bytes.byteLength >= OUTSIDE_ENGINE_BYTES) {
      try { return await (await outside())(bytes, options); } catch {}
    }
    return engine(bytes, options);
  };
}

export function sessionFontDecoder(engine: FontDecoder): FontDecoder {
  return engineKind === 'wasm' ? decodeLargeFacesOutsideEngine(engine,
    () => import('@betteroffice/fonts').then((fonts) => fonts.decodeFontFaceOutsideEngine)) : engine;
}
