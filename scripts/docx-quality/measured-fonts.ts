import type { BundledFontLoader, BundledFontSource } from '../../packages/fonts/src/provider';

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
  <A extends unknown[], L extends BundledFontLoader>(
    fontLoads: FontLoad[],
    kind: string,
    resolve: (...args: A) => L | undefined,
    faceOf: FaceOf | undefined,
  ) =>
  (...args: A) => {
    const load = resolve(...args);
    if (!load) return undefined;
    const face = faceOf?.(kind, args);
    const family = kind === 'base-latin' ? 'BetterOffice Base Latin' : String(args[0]);
    const record = (entry: Omit<FontLoad, 'kind' | 'family'>) =>
      fontLoads.push({ kind, family, ...entry, ...(face === undefined ? {} : { face }) });
    const recorded = <T>(run: () => Promise<T>) => async () => {
      try {
        const result = await run();
        record({ ok: true });
        return result;
      } catch (error) {
        record({ ok: false, error: String(error) });
        throw error;
      }
    };
    const { readOnlyBytes } = load;
    return Object.assign(
      recorded(() => load()),
      load,
      readOnlyBytes ? { readOnlyBytes: recorded(() => readOnlyBytes.call(load)) } : {},
    );
  };

/** The provider the harness measures with, recording every font load. */
export function measuredFontProvider(provider: BundledFontSource, fontLoads: FontLoad[], faceOf?: FaceOf) {
  const tracked = <A extends unknown[], L extends BundledFontLoader>(
    kind: string,
    resolve: (...args: A) => L | undefined,
    attributed = true,
  ) => track(fontLoads, kind, resolve, attributed ? faceOf : undefined);
  return {
    resolve: tracked('family', provider.resolve),
    ...(provider.resolveFamily ? { resolveFamily: tracked('bundled-family', provider.resolveFamily, false) } : {}),
    resolveScriptFallback: tracked('script', provider.resolveScriptFallback),
    ...(provider.resolveGlyphFallback ? { resolveGlyphFallback: tracked('glyph', provider.resolveGlyphFallback) } : {}),
    resolveLastResort: tracked('fallback', provider.resolveLastResort),
    ...(provider.resolveBaseLatin ? { resolveBaseLatin: tracked('base-latin', provider.resolveBaseLatin) } : {}),
    ...(provider.resolveLightVariant
      ? { resolveLightVariant: tracked('light-variant', provider.resolveLightVariant, false) }
      : {}),
  };
}
