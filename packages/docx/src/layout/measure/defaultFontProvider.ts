/** Lazy default-font provider resolution. @packageDocumentation */

import type { BundledFontLoader, BundledFontProvider } from './fontRegistry';

export const IMPLICIT_FONT_COVERAGE = '@betteroffice/fonts/coverage';

export class ImplicitFontAbsentError extends Error {}

const implicitProviders = new WeakSet<BundledFontProvider>();

export function isImplicitFontProvider(provider: BundledFontProvider | undefined): boolean {
  return provider !== undefined && implicitProviders.has(provider);
}

function implicitFontProvider(manager: typeof import('@betteroffice/fonts')): BundledFontProvider {
  const provider = manager.createFontProvider();
  const checked = (
    loader: BundledFontLoader | undefined,
    face: import('@betteroffice/fonts').BundledFontFace | undefined
  ): BundledFontLoader | undefined => {
    if (!loader || !face?.script) return loader;
    return Object.assign(async () => {
      let bytes: ArrayBuffer;
      try { bytes = await loader(); }
      catch { throw new ImplicitFontAbsentError(face.family); }
      if (bytes.byteLength !== face.byteLength) throw new ImplicitFontAbsentError(face.family);
      return bytes;
    }, loader);
  };
  const implicit: BundledFontProvider = {
    ...provider,
    resolve: (family, bold, italic) => checked(
      provider.resolve(family, bold, italic), manager.resolveMetricCompatFace(family, bold, italic)
    ),
    resolveFamily: (family, bold, italic) => checked(
      provider.resolveFamily(family, bold, italic), manager.resolveBundledFamilyFace(family, bold, italic)
    ),
    resolveScriptFallback: (script, bold, italic) => checked(
      provider.resolveScriptFallback(script, bold, italic), manager.resolveScriptFallbackFace(script, bold, italic)
    ),
    resolveGlyphFallback: (family, bold, italic, office) => checked(
      provider.resolveGlyphFallback?.(family, bold, italic, office),
      manager.resolveGlyphFallbackFace(family, bold, italic, office)
    ),
    resolveLastResort: (family, bold, italic, office) => checked(
      provider.resolveLastResort(family, bold, italic, office),
      manager.resolveLastResortFace(family, bold, italic, office)
    ),
  };
  implicitProviders.add(implicit);
  return implicit;
}

/** The consumed font-manager contract. @public */
export interface BundledFontModule {
  createFontProvider(options?: { baseUrl?: string | URL }): BundledFontProvider;
}

/** Override the lazy `@betteroffice/fonts` default. @public */
export interface DefaultFontOptions {
  /** The imported `@betteroffice/fonts` module. */
  fonts?: BundledFontModule;
  /** Lazy alternative to {@link DefaultFontOptions.fonts}, resolving to the same module. */
  load?: () => Promise<unknown>;
  /** Fetch the faces from this base URL instead of the module's own assets. */
  baseUrl?: string | URL;
}

let options: DefaultFontOptions = {};
let resolved: Promise<BundledFontProvider | undefined> | undefined;

function configuredBaseUrl(baseUrl: string | URL | undefined): URL | undefined {
  if (baseUrl === undefined) return undefined;
  const href = typeof baseUrl === 'string' ? baseUrl : baseUrl.href;
  const locationHref = typeof location === 'undefined' ? undefined : location.href;
  try {
    return new URL(href, locationHref);
  } catch {
    throw new TypeError(
      `configureDefaultFonts baseUrl must be absolute when no browser location exists: ${href}`
    );
  }
}

/** Whether a default font source is available. */
export function hasDefaultFontSource(): boolean {
  return true;
}

/** Whether {@link configureDefaultFonts} was given a font source. */
export function hasConfiguredFontSource(): boolean {
  return options.fonts !== undefined || options.load !== undefined;
}

/** Configure the default provider and reset its memoized resolution. @public */
export function configureDefaultFonts(next: DefaultFontOptions): void {
  options = { ...next, baseUrl: configuredBaseUrl(next.baseUrl) };
  resolved = undefined;
}

async function load(): Promise<BundledFontProvider | undefined> {
  const { baseUrl, fonts, load: loader } = options;
  try {
    if (fonts === undefined && loader === undefined && baseUrl === undefined) {
      return implicitFontProvider(await import('@betteroffice/fonts'));
    }
    const module = fonts ?? (
      (loader ? await loader() : await import('@betteroffice/fonts')) as BundledFontModule
    );
    return module.createFontProvider(baseUrl === undefined ? undefined : { baseUrl });
  } catch (error) {
    console.warn('[configureDefaultFonts] the configured font source failed to load', error);
    return undefined;
  }
}

/** Resolve the default provider, or `undefined` after a retryable failure. @public */
export function resolveDefaultFontProvider(): Promise<BundledFontProvider | undefined> {
  if (resolved === undefined) {
    const promise = load();
    promise.then((provider) => {
      if (provider === undefined && resolved === promise) resolved = undefined;
    });
    resolved = promise;
  }
  return resolved;
}
