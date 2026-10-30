/** Lazy default-font provider resolution. @packageDocumentation */

import type { BundledFontLoader, BundledFontProvider } from './fontRegistry';
import type { FontDecoder } from '../../../../../shared/office-session/fontDecoder';
import { decodeFontFace } from '../../wasm/fontDecoder';

export const IMPLICIT_FONT_COVERAGE = '@betteroffice/fonts/coverage';

export class ImplicitFontAbsentError extends Error {}

const implicitProviders = new WeakSet<BundledFontProvider>();

export function isImplicitFontProvider(provider: BundledFontProvider | undefined): boolean {
  return provider !== undefined && implicitProviders.has(provider);
}

const hostedFamiliesKey = Symbol.for('@betteroffice/docx/hostedFontFamilies');

export function hostedFontFamilies(source: object | undefined): ReadonlySet<string> | undefined {
  return (source as { [hostedFamiliesKey]?: ReadonlySet<string> } | undefined)?.[hostedFamiliesKey];
}

export function excludeHostedFontFamilies(
  provider: BundledFontProvider | undefined,
  families: ReadonlySet<string> | undefined
): BundledFontProvider | undefined {
  if (!provider || !isImplicitFontProvider(provider) || !families?.size) return provider;
  const hosted = (family: string) => families.has(family.trim().toLowerCase());
  const filtered: BundledFontProvider = {
    ...provider,
    resolve: (family, bold, italic) => hosted(family) ? undefined : provider.resolve(family, bold, italic),
    resolveFamily: (family, bold, italic) => hosted(family) ? undefined : provider.resolveFamily?.(family, bold, italic),
    resolveGlyphFallback: (family, bold, italic, office) => hosted(family) ? undefined : provider.resolveGlyphFallback?.(family, bold, italic, office),
    resolveLastResort: (family, bold, italic, office) => hosted(family) ? undefined : provider.resolveLastResort?.(family, bold, italic, office),
  };
  implicitProviders.add(filtered);
  return filtered;
}

function implicitFontProvider(
  manager: typeof import('@betteroffice/fonts'),
  runtime: { decoder?: FontDecoder }
): BundledFontProvider {
  const provider = manager.createFontProvider(undefined, runtime);
  const providesFamily = (family: string) =>
    manager.resolveMetricCompatFamily(family) !== undefined ||
    manager.resolveBundledFamilyFace(family, false, false) !== undefined;
  const checked = (
    loader: BundledFontLoader | undefined,
    face: import('@betteroffice/fonts').BundledFontFace | undefined
  ): BundledFontLoader | undefined => {
    if (!face) return undefined;
    if (!loader || !face.script) return loader;
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
      provider.resolve(family, bold, italic) ?? provider.resolveFamily(family, bold, italic),
      manager.resolveMetricCompatFace(family, bold, italic) ?? manager.resolveBundledFamilyFace(family, bold, italic)
    ),
    resolveFamily: (family, bold, italic) => checked(
      provider.resolveFamily(family, bold, italic), manager.resolveBundledFamilyFace(family, bold, italic)
    ),
    resolveScriptFallback: (script, bold, italic) => checked(
      provider.resolveScriptFallback(script, bold, italic), manager.resolveScriptFallbackFace(script, bold, italic)
    ),
    resolveGlyphFallback: (family, bold, italic, office) => providesFamily(family) ? checked(
      provider.resolveGlyphFallback?.(family, bold, italic, office),
      manager.resolveGlyphFallbackFace(family, bold, italic, office)
    ) : undefined,
    resolveLastResort: (family, bold, italic, office) => providesFamily(family) ? checked(
      provider.resolveLastResort(family, bold, italic, office),
      manager.resolveLastResortFace(family, bold, italic, office)
    ) : undefined,
  };
  implicitProviders.add(implicit);
  return implicit;
}

/** The consumed font-manager contract. @public */
export interface BundledFontModule {
  createFontProvider(options?: { baseUrl?: string | URL }, runtime?: { decoder?: FontDecoder }): BundledFontProvider;
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
const engineRuntime = { decoder: decodeFontFace };
let resolved = new WeakMap<object, Promise<BundledFontProvider | undefined>>();

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
  resolved = new WeakMap();
}

async function load(runtime: { decoder?: FontDecoder }): Promise<BundledFontProvider | undefined> {
  const { baseUrl, fonts, load: loader } = options;
  try {
    if (fonts === undefined && loader === undefined && baseUrl === undefined) {
      return implicitFontProvider(await import('@betteroffice/fonts'), runtime);
    }
    const module = fonts ?? (
      (loader ? await loader() : await import('@betteroffice/fonts')) as BundledFontModule
    );
    return module.createFontProvider(baseUrl === undefined ? undefined : { baseUrl }, runtime);
  } catch (error) {
    console.warn('[configureDefaultFonts] the configured font source failed to load', error);
    return undefined;
  }
}

/** Resolve the default provider, or `undefined` after a retryable failure. @public */
export function resolveDefaultFontProvider(
  runtime: { decoder?: FontDecoder } = engineRuntime
): Promise<BundledFontProvider | undefined> {
  let pending = resolved.get(runtime);
  if (pending === undefined) {
    const promise = load(runtime);
    promise.then((provider) => {
      if (provider === undefined && resolved.get(runtime) === promise) resolved.delete(runtime);
    });
    resolved.set(runtime, promise);
    pending = promise;
  }
  return pending;
}
