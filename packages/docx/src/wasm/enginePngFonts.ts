import type { BundledFontFace } from '@betteroffice/fonts';
import type { BundledFontLoader, BundledFontProvider } from '../layout/measure/fontRegistry';
import { fontUnavailable } from '../layout/render/png';

type Fonts = typeof import('@betteroffice/fonts');

/** A face the layout could not load: the face, the family it was resolved for, and why. */
export interface FailedFaceLoad {
  face: BundledFontFace | undefined;
  family: string;
  reason: string;
}

/** A provider that records the faces it fails to load, and those failures. */
export interface RecordedFontProvider {
  provider: BundledFontProvider;
  failures: FailedFaceLoad[];
}

type Methods = { [K in keyof BundledFontProvider]-?: NonNullable<BundledFontProvider[K]> };
type Resolved = { face: BundledFontFace | undefined; family: string };

const RESOLVED: { [K in keyof Methods]?: (fonts: Fonts, ...args: Parameters<Methods[K]>) => Resolved } = {
  resolveBaseLatin: (fonts, bold, italic) => ({
    family: fonts.BASE_LATIN_FAMILY,
    face: fonts.BASE_LATIN_FONTS.find((face) => face.weight === (bold ? 700 : 400) && face.style === (italic ? 'italic' : 'normal')),
  }),
  resolve: (fonts, family, bold, italic) => ({ family, face: fonts.resolveMetricCompatFace(family, bold, italic) }),
  resolveFamily: (fonts, family, bold, italic) => ({ family, face: fonts.resolveBundledFamilyFace(family, bold, italic) }),
  resolveScriptFallback: (fonts, script, bold, italic, family) => ({
    family: family ?? script, face: fonts.resolveScriptFallbackFace(script, bold, italic, family),
  }),
  resolveGlyphFallback: (fonts, family, bold, italic, office) => ({ family, face: fonts.resolveGlyphFallbackFace(family, bold, italic, office) }),
  resolveLastResort: (fonts, family, bold, italic, office) => ({ family, face: fonts.resolveLastResortChoice(family, bold, italic, office).face }),
};

function resolvedFace(fonts: Fonts, method: PropertyKey, args: unknown[]): Resolved {
  const resolve = RESOLVED[method as keyof Methods] as ((fonts: Fonts, ...args: unknown[]) => Resolved) | undefined;
  return resolve?.(fonts, ...args) ?? { face: undefined, family: String(args[0]) };
}

/**
 * Wraps `provider` so each face load that fails is recorded and still rejects, leaving the layout to skip the face
 * as it does without the wrapper. A face whose optional pack is not installed is not recorded. Every property of
 * the provider and of its loaders is kept, and each loader is wrapped once.
 */
export function recordFaceLoads(provider: BundledFontProvider, fonts: Fonts): RecordedFontProvider {
  const failures: FailedFaceLoad[] = [];
  const loaders = new WeakMap<BundledFontLoader, BundledFontLoader>();
  const methods = new Map<PropertyKey, unknown>();

  const watch = (loader: BundledFontLoader, resolved: () => Resolved): BundledFontLoader => {
    let watched = loaders.get(loader);
    if (!watched) {
      const recorded = (load: () => Promise<ArrayBuffer>) => async () => {
        try {
          return await load();
        } catch (error) {
          if (!fonts.isMissingFontAddon(error)) {
            failures.push({ ...resolved(), reason: error instanceof Error ? error.message : String(error) });
          }
          throw error;
        }
      };
      const readOnlyBytes = loader.readOnlyBytes;
      watched = Object.assign(
        recorded(() => loader()),
        loader,
        readOnlyBytes ? { readOnlyBytes: recorded(() => readOnlyBytes.call(loader)) } : {},
      );
      loaders.set(loader, watched);
    }
    return watched;
  };

  const watchedProvider = new Proxy(provider, {
    get(target, key, receiver) {
      const value: unknown = Reflect.get(target, key, receiver);
      if (typeof value !== 'function') return value;
      let method = methods.get(key);
      if (!method) {
        method = (...args: unknown[]) => {
          const loader: unknown = value.apply(target, args);
          return typeof loader === 'function'
            ? watch(loader as BundledFontLoader, () => resolvedFace(fonts, key, args))
            : loader;
        };
        methods.set(key, method);
      }
      return method;
    },
  });
  return { provider: watchedProvider, failures };
}

/** The `font-unavailable` rejection for the first recorded failure, or undefined when none was recorded. */
export function faceLoadFailure(failures: readonly FailedFaceLoad[]): ReturnType<typeof fontUnavailable> | undefined {
  const [first] = failures;
  if (!first) return undefined;
  const face = first.face ? `${first.face.family} for ${first.family}` : first.family;
  return fontUnavailable(`${face} could not load: ${first.reason}`);
}

/** `provider` wrapped by {@link recordFaceLoads}, with the font package it loads from. */
export async function recordProviderFaceLoads(provider: BundledFontProvider): Promise<RecordedFontProvider> {
  return recordFaceLoads(provider, await import('@betteroffice/fonts'));
}

/** `work`'s result, or a rejection with the signal's reason as soon as it aborts; `work` itself runs on. */
export async function untilAborted<T>(work: Promise<T>, signal?: AbortSignal): Promise<T> {
  if (!signal) return work;
  signal.throwIfAborted();
  let stop: (() => void) | undefined;
  const aborted = new Promise<never>((_, reject) => {
    stop = () => reject(signal.reason);
    signal.addEventListener('abort', stop, { once: true });
  });
  work.catch(() => {});
  try {
    return await Promise.race([work, aborted]);
  } finally {
    signal.removeEventListener('abort', stop!);
  }
}
