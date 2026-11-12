import type { BundledFontSource } from '../../packages/fonts/src/provider';

export type FontLoad = {
  kind: string;
  family: string;
  ok: boolean;
  error?: string;
  /** The advertised face file the provider loaded for this request. */
  face?: string;
};

/** Names the face file a provider call resolves to, from the provider's own manifest. */
export type FaceOf = (kind: string, args: unknown[]) => string | undefined;

const track =
  <A extends unknown[], L extends () => Promise<ArrayBuffer>>(
    fontLoads: FontLoad[],
    kind: string,
    resolve: (...args: A) => L | undefined,
    faceOf?: FaceOf,
  ) =>
  (...args: A) => {
    const load = resolve(...args);
    if (!load) return undefined;
    const face = faceOf?.(kind, args);
    const record = (entry: Omit<FontLoad, 'kind' | 'family'>) =>
      fontLoads.push({ kind, family: String(args[0]), ...entry, ...(face === undefined ? {} : { face }) });
    return Object.assign(async () => {
      try {
        const bytes = await load();
        record({ ok: true });
        return bytes;
      } catch (error) {
        record({ ok: false, error: String(error) });
        throw error;
      }
    }, load);
  };

/** The provider the harness measures with, recording every font load. */
export function measuredFontProvider(provider: BundledFontSource, fontLoads: FontLoad[], faceOf?: FaceOf) {
  return {
    resolve: track(fontLoads, 'family', provider.resolve, faceOf),
    resolveScriptFallback: track(fontLoads, 'script', provider.resolveScriptFallback, faceOf),
    ...(provider.resolveGlyphFallback
      ? { resolveGlyphFallback: track(fontLoads, 'glyph', provider.resolveGlyphFallback, faceOf) }
      : {}),
    resolveLastResort: track(fontLoads, 'fallback', provider.resolveLastResort, faceOf),
  };
}
