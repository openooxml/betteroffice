import { decodeLargeFacesOutsideEngine, type FontDecoder } from '../../../../shared/office-session/fontDecoder';
import { engineKind } from '../wasm/engine';

export function sessionFontDecoder(engine: FontDecoder): FontDecoder {
  return engineKind === 'wasm' ? decodeLargeFacesOutsideEngine(engine,
    () => import('@betteroffice/fonts').then((fonts) => fonts.decodeFontFaceOutsideEngine)) : engine;
}
