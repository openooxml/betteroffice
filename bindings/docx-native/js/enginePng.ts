import { fontUnavailable, pagePngRequest, type DocxPngOptions } from '../../../packages/docx/src/layout/render/png';
import { PngExportError, type PngSkippedEffectKind } from '../../../shared/office-session/pngExportError';
import { DEFAULT_STRICT_PICTURES, strictPicturesError } from '../../../shared/office-session/strictPictures';
import type { DocxPngEngine } from '../../../packages/docx/src/wasm/enginePng';
import type { DocxNativeExtras } from '../../native-extras-types';

export type { DocxPngEngine } from '../../../packages/docx/src/wasm/enginePng';

export const engineKind: 'wasm' | 'native' = 'native';

type NativeSession = DocxNativeExtras['members']['EditSession'];

/** The kind each name the native walk reports stands for. */
const NATIVE_EFFECT_KINDS: Record<PngSkippedEffectKind, PngSkippedEffectKind> = {
  'text.outline': 'text.outline',
  'text.emphasisMark': 'text.emphasisMark',
  'text.effect': 'text.effect',
  'text.modernEffects': 'text.modernEffects',
  'text.shadow': 'text.shadow',
  'text.leaderGlyphs': 'text.leaderGlyphs',
  'shape.effects': 'shape.effects',
  'image.filter': 'image.filter',
  'image.effects': 'image.effects',
};

/** The kinds the native walk names, throwing on a name that stands for none. */
function nativeEffectKinds(names: readonly string[]): PngSkippedEffectKind[] {
  return names.map((name) => {
    if (!Object.hasOwn(NATIVE_EFFECT_KINDS, name)) throw new Error(`The native PNG engine reported an unknown skipped effect: ${name}`);
    return NATIVE_EFFECT_KINDS[name as PngSkippedEffectKind];
  });
}

/**
 * Rasterizes the page's display list in the native engine on the calling
 * thread, with the faces and media of the session that laid it out. Scale and
 * size are checked as on the canvas engine, so both reject alike.
 */
export async function renderPng(engine: DocxPngEngine, options: Pick<DocxPngOptions, 'scale' | 'signal' | 'onDiagnostic' | 'strictPictures'>): Promise<Uint8Array> {
  const { scale } = pagePngRequest(engine.page, options);
  options.signal?.throwIfAborted();
  let report: ReturnType<NativeSession['renderPngTargetReport']>;
  try {
    report = (engine.session as NativeSession).renderPngTargetReport(
      JSON.stringify({ ...engine.displayList, fontChains: engine.fontChains, pages: [engine.page] }), '{"page":0}', JSON.stringify({ scale }));
  } catch (error) {
    const message = error instanceof Error ? error.message : '';
    if (message.startsWith('png-font-unavailable: ')) throw fontUnavailable(message.slice(22));
    if (message.startsWith('png-too-large: ')) throw new PngExportError('canvas-too-large', message.slice(15));
    throw error;
  }
  options.signal?.throwIfAborted();
  if (report.pictures.length && (options.strictPictures ?? DEFAULT_STRICT_PICTURES)) throw strictPicturesError(report.pictures);
  const effects = nativeEffectKinds(report.skippedEffects);
  if (report.pictures.length) {
    options.onDiagnostic?.({ code: 'picture-unavailable', pictures: report.pictures.map((picture) => ({ ...picture })) });
  }
  if (effects.length) options.onDiagnostic?.({ code: 'effects-skipped', page: engine.page.pageIndex, effects });
  return report.png;
}

export default renderPng;
