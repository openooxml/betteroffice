import type { BundledFontSource } from '../../packages/fonts/src/provider';

export type FontLoad = {
  kind: string;
  family: string;
  ok: boolean;
  error?: string;
};

const track =
  <A extends unknown[], L extends () => Promise<ArrayBuffer>>(
    fontLoads: FontLoad[],
    kind: string,
    resolve: (...args: A) => L | undefined,
  ) =>
  (...args: A) => {
    const load = resolve(...args);
    if (!load) return undefined;
    return Object.assign(async () => {
      try {
        const bytes = await load();
        fontLoads.push({ kind, family: String(args[0]), ok: true });
        return bytes;
      } catch (error) {
        fontLoads.push({
          kind,
          family: String(args[0]),
          ok: false,
          error: String(error),
        });
        throw error;
      }
    }, load);
  };

/** The provider the harness measures with, recording every font load. */
export function measuredFontProvider(provider: BundledFontSource, fontLoads: FontLoad[]) {
  return {
    resolve: track(fontLoads, 'family', provider.resolve),
    resolveScriptFallback: track(fontLoads, 'script', provider.resolveScriptFallback),
    ...(provider.resolveGlyphFallback
      ? { resolveGlyphFallback: track(fontLoads, 'glyph', provider.resolveGlyphFallback) }
      : {}),
    resolveLastResort: track(fontLoads, 'fallback', provider.resolveLastResort),
  };
}
